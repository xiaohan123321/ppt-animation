# -*- coding: utf-8 -*-
"""整片转场预览（真实渲染版）。

为什么不用 render_gif 合成：它把椭圆裁切的图片画成带边框的矩形，
还受 256 色量化影响出现色带 —— 预览失真会误导人。

这里改用「LibreOffice 真渲染 → PIL 转场合成」：
  1. soffice 把 pptx 转 PDF，pymupdf 每页导 PNG（字体/图片/透明度全部准确）
  2. 按 plans.json 的 transition / transition_ms 合成转场帧
     · morph  → 交叉淡入（圆窗同名形状位置恒定，真实 morph 观感即如此）
     · wheelN → N 个扇区同时从中心展开（复刻参考片的光圈扫过）
     · circle → 中心圆扩大
  3. 每页停留 hold 秒后出 GIF

用法：
    python render_deck.py <pptx> <plans.json> <out.gif>
        [--width 640] [--hold 1.2] [--fps 15]
"""
import json
import math
import os
import subprocess
import sys

from PIL import Image, ImageChops, ImageDraw, ImageFilter

SOFFICE = r"D:\LibreOffice\program\soffice.exe"


def _arg(name, default, cast=str):
    if name in sys.argv:
        return cast(sys.argv[sys.argv.index(name) + 1])
    return default


def to_pages(pptx, tmp, dpi=120):
    os.makedirs(tmp, exist_ok=True)
    pdf = os.path.join(tmp, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
    if not os.path.exists(pdf):
        subprocess.run([SOFFICE, "--headless", "--convert-to", "pdf",
                        "--outdir", tmp, pptx],
                       check=True, stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL)
    import pymupdf
    d = pymupdf.open(pdf)
    out = []
    for i, p in enumerate(d):
        fp = os.path.join(tmp, f"pg{i+1:02d}.png")
        if not os.path.exists(fp):
            p.get_pixmap(dpi=dpi).save(fp)
        out.append(fp)
    return out


def mask_circle(size, t):
    W, H = size
    m = Image.new("L", size, 0)
    r = math.hypot(W, H) * t
    ImageDraw.Draw(m).ellipse([W/2 - r, H/2 - r, W/2 + r, H/2 + r], fill=255)
    return m


def mask_wheel(size, t, spokes=8):
    """N 个扇区同时从 -90° 展开 —— PowerPoint wheel 的真实形态。"""
    W, H = size
    m = Image.new("L", size, 0)
    d = ImageDraw.Draw(m)
    r = math.hypot(W, H)
    step = 360.0 / spokes
    span = step * t
    for k in range(spokes):
        a0 = -90 + k * step
        d.pieslice([W/2 - r, H/2 - r, W/2 + r, H/2 + r], a0, a0 + span, fill=255)
    return m


def mask_sweep(size, t):
    """单一扇形顺时针扫过（丝滑擦除）。"""
    W, H = size
    m = Image.new("L", size, 0)
    r = math.hypot(W, H)
    ImageDraw.Draw(m).pieslice([W/2 - r, H/2 - r, W/2 + r, H/2 + r],
                               -90, -90 + 360 * t, fill=255)
    return m


def ease_out(t):
    return 1 - (1 - t) ** 3


def anchor_boxes(pptx):
    """读每页核心锚点（coreImg）的位置尺寸，用于 morph 位移合成。"""
    from pptx import Presentation
    prs = Presentation(pptx)
    out = []
    for s in prs.slides:
        box = None
        for sp in s.shapes:
            if sp.name == "coreImg":
                box = (sp.left / 12700, sp.top / 12700,
                       sp.width / 12700, sp.height / 12700)
                break
        out.append(box)
    return out


def bloom_boxes(pptx):
    """读每页"叠层图形"（spot / moon）的位置尺寸。

    这些形状不在 coreImg 的外接框里，合成时会被当成背景的一部分、
    在硬切那一刻就出现在终点；而圆窗是滑过来的，于是屏幕上同时出现
    两个伴圆（背景一个、移动图层里一个），看起来就是画重了。
    所以要把它们单独擦掉，再跟着接力淡入。
    """
    from pptx import Presentation
    prs = Presentation(pptx)
    out = []
    for s in prs.slides:
        boxes = []
        for sp in s.shapes:
            if sp.name in ("spot", "moon"):
                boxes.append((sp.name, (sp.left / 12700, sp.top / 12700,
                                        sp.width / 12700, sp.height / 12700)))
        out.append(boxes)
    return out


def _paper(img):
    """取这一页的"纸底"颜色。

    ⚠️ 不能写死 (10,7,8)：那是深色片的底色，浅色片（米白 / 白）会把
    出血补边和挖洞补成一块近黑，圆窗一滑开就露出一块黑斑。
    做法：沿四条边各取若干点，量化到 8 的格子后取众数 —— 出血页也能取准。
    """
    w, h = img.size
    pts = []
    for i in range(11):
        t = i / 10
        pts += [(int(3 + t * (w - 7)), 3), (int(3 + t * (w - 7)), h - 4),
                (3, int(3 + t * (h - 7))), (w - 4, int(3 + t * (h - 7)))]
    hist = {}
    for p in pts:
        c = img.getpixel(p)[:3]
        k = tuple(v // 8 for v in c)
        hist.setdefault(k, []).append(c)
    best = max(hist.values(), key=len)
    n = len(best)
    return tuple(sum(c[i] for c in best) // n for i in range(3))


def _ring_paper(img, box):
    """取"洞口边缘"一圈的底色中位数。

    整页取众数还是会有 1~2 个色阶的偏差，浅底片上填出来的洞会显出一个淡圆盘。
    沿椭圆边界（r×1.04）采样 48 点取中位数，填色就和周围几乎完全一致。
    """
    x, y, w, h = box
    cx, cy = x + w / 2.0, y + h / 2.0
    rx, ry = max(1.0, w / 2.0 * 1.04), max(1.0, h / 2.0 * 1.04)
    pts = []
    for i in range(48):
        a = i / 48.0 * 2 * math.pi
        px_, py_ = int(cx + rx * math.cos(a)), int(cy + ry * math.sin(a))
        if 0 <= px_ < img.width and 0 <= py_ < img.height:
            pts.append(img.getpixel((px_, py_))[:3])
    if not pts:
        return _paper(img)
    pts.sort(key=lambda c: c[0] + c[1] + c[2])
    return pts[len(pts) // 2]


def _erase_circle(img, box, bg=None, feather=0):
    """把圆窗区域涂成底色，得到"没有圆"的背景层。

    feather>0 时用羽化蒙版贴补丁 —— 硬边平涂在平滑渐变上会显出一块"淡斑"，
    尤其补的是圆内的区域（伴圆抹除）。羽化 15~20px 就看不出来了。
    """
    c = img.copy()
    if not box:
        return c
    if bg is None:
        bg = _ring_paper(img, box)
    x, y, w, h = box
    # 多涂 5px：金环描边正好压在椭圆边界上，不多涂会留下半圆残影
    pad = max(3.0, w * 0.018)
    mask = Image.new("L", img.size, 0)
    ImageDraw.Draw(mask).ellipse([x - pad, y - pad, x + w + pad, y + h + pad], fill=255)
    if feather:
        mask = mask.filter(ImageFilter.GaussianBlur(feather))
    c.paste(Image.new("RGB", img.size, bg), (0, 0), mask)
    return c


def _ramp(n, side):
    """生成一条渐变 alpha：接缝处 255，向外淡到 0。"""
    img = Image.new("L", (n, 1))
    for i in range(n):
        t = i / max(1, n - 1)
        v = int(255 * (1 - t)) if side in ("r", "b") else int(255 * t)
        img.putpixel((i, 0), v)
    return img


def _extend_edge(tile, x, y, w, h, iw, ih):
    """把"落在画布外、裁不到"的区域补满。

    为什么必须填：圆窗常常出血到画布外（P1 圆心 792、直径 488 → 右边界 1036
    > 960）。这条内容在源页里根本不存在。等圆窗滑到画面中间，这块"洞"就露出来了
    —— 洞后面是下一页的内容，叠在一起就是一道生硬的竖边。

    补法：接缝处取最靠边的 1 列/行拉伸，再叠一条"由实到透明"的渐变，淡到背景。
      · 直接拉一整条 → 竖直糊痕 + 一道硬线；
      · 镜像 → 把"01"和金环复制一份，出重影；
      · 淡出 → 接缝连续、远端融进黑底，正好符合这套暗角美术。
    """
    sx0, sy0 = max(0, x), max(0, y)
    sx1, sy1 = min(iw, x + w), min(ih, y + h)
    if sx1 <= sx0 or sy1 <= sy0:
        return tile
    L, T = sx0 - x, sy0 - y
    R, B = w - (sx1 - x), h - (sy1 - y)

    if L > 0:
        col = tile.crop((L, 0, L + 1, h)).resize((L, h))
        tile.paste(col, (0, 0), _ramp(L, "l").resize((L, h)))
    if R > 0:
        col = tile.crop((w - R - 1, 0, w - R, h)).resize((R, h))
        tile.paste(col, (w - R, 0), _ramp(R, "r").resize((R, h)))
    if T > 0:
        row = tile.crop((0, T, w, T + 1)).resize((w, T))
        tile.paste(row, (0, 0),
                   _ramp(T, "t").transpose(Image.TRANSPOSE).resize((w, T)))
    if B > 0:
        row = tile.crop((0, h - B - 1, w, h - B)).resize((w, B))
        tile.paste(row, (0, h - B),
                   _ramp(B, "b").transpose(Image.TRANSPOSE).resize((w, B)))
    return tile


def _circle_layer(img, box):
    """从整页里抠出圆窗（含金环），裁成圆形，带 alpha。"""
    x, y, w, h = [int(round(v)) for v in box]
    w = max(2, w); h = max(2, h)
    tile = Image.new("RGB", (w, h), _paper(img))
    sx0, sy0 = max(0, x), max(0, y)
    sx1, sy1 = min(img.width, x + w), min(img.height, y + h)
    if sx1 > sx0 and sy1 > sy0:
        piece = img.crop((sx0, sy0, sx1, sy1))
        tile.paste(piece, (sx0 - x, sy0 - y))
    tile = _extend_edge(tile, x, y, w, h, img.width, img.height)
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).ellipse([0, 0, w, h], fill=255)
    out = tile.convert("RGBA")
    out.putalpha(mask)
    return out


def _fade(img, k):
    """只改 alpha，不动颜色。"""
    if k >= 0.999:
        return img
    out = img.copy()
    out.putalpha(img.split()[3].point(lambda v: int(v * max(0.0, k))))
    return out


def morph_slide_frames(a, b, a_box, b_box, ms, fps):
    """真实 morph 观感：锚点（圆窗）在页间滑动 + 缩放，其余内容交叉淡入。

    时间轴要解耦：圆窗一直在滑（全时长），文字层在前 35% 就淡完。
    如果两层同速淡，全程都是双曝光，看着像重影而不是转场。
    """
    n = max(2, int(ms / 1000 * fps))
    a_bg = _erase_circle(a, a_box)
    b_bg = _erase_circle(b, b_box)
    a_c = _circle_layer(a, a_box) if a_box else None
    b_c = _circle_layer(b, b_box) if b_box else None
    out = []
    for i in range(n):
        t = ease_out((i + 1) / n)
        # 文字/背景层：2.8 倍速提前淡完，避免长时间双曝光
        tb = ease_out(min(1.0, (i + 1) / n * 2.8))
        canvas = Image.blend(a_bg, b_bg, tb).convert("RGBA")
        if a_box and b_box:
            x = a_box[0] + (b_box[0] - a_box[0]) * t
            y = a_box[1] + (b_box[1] - a_box[1]) * t
            w = a_box[2] + (b_box[2] - a_box[2]) * t
            h = a_box[3] + (b_box[3] - a_box[3]) * t
            # 两页圆窗直径可能不同（封面 488 / 内容 428），先统一尺寸再混合
            b_same = b_c if b_c.size == a_c.size else b_c.resize(a_c.size, Image.LANCZOS)
            # 交叉溶解：b 按 tb 压在 a 上（a 保持不透明，作底层）。
            # 注意 alpha_composite(底, 前) —— 参数反了会变成"永远是上一页"。
            # 这样结果 alpha 恒为 1 —— 用 Image.blend 混 RGBA 会让透明区的黑色
            # 渗进颜色，圆边缘出现一圈脏边；只调 alpha 不会。
            layer = Image.alpha_composite(a_c, _fade(b_same, tb))
            # 擦除时多涂了一圈 pad，贴回时也要多贴一圈，否则留一道淡色细缝
            pad = max(3.0, w * 0.018)
            layer = layer.resize((max(2, int(w + 2 * pad)),
                                  max(2, int(h + 2 * pad))), Image.LANCZOS)
            canvas.alpha_composite(layer, (int(round(x - pad)), int(round(y - pad))))
        elif b_c:
            canvas.alpha_composite(b_c, (int(round(b_box[0])), int(round(b_box[1]))))
        out.append(canvas.convert("RGB"))
    return out


def carry_frames(a, b, a_box, b_box, opts, ms, fps, b_extra=None):
    """复刻 carry 入场：圆窗从"上一页所在的位置"滑到本页位置。

    opts.dx/dy 是画布比例位移，opts.scale 是起始缩放百分比（114 = 从 114% 收到 100%）。
    和 morph 的区别：morph 是"两页同时插值"，carry 是"新页的圆先站在老位置，
    再滑到新位置"——旧页在切页瞬间就没了，所以底色要快速硬切（不跟着慢慢淡）。

    b_extra：叠层图形（spot / moon）的外接框。它们在 coreImg 框外，不清掉就会
    在硬切时先出现在终点，和移动图层里的那一份重叠 —— 屏幕上两个伴圆。
    处理：先从背景擦掉，再在接力后 40% 里淡入（垫在圆窗下面，露出的是月牙部分）。
    """
    n = max(2, int(ms / 1000 * fps))
    W, H = a.size
    dx = float(opts.get("dx", 0.0)) * W
    dy = float(opts.get("dy", 0.0)) * H
    sc = float(opts.get("scale") or 100.0) / 100.0
    paper = _paper(b)
    a_bg = _erase_circle(a, a_box)
    b_bg = _erase_circle(b, b_box)
    # 背景层也要把叠层图形抹掉：它们在 coreImg 框外，不清掉会在硬切那一刻
    # 就出现在终点，而圆窗还在滑 —— 屏幕上同时有两个伴圆。
    for _n, _ex in (b_extra or []):
        b_bg = _erase_circle(b_bg, _ex, paper)
    # 裁圆窗用的源图：把伴圆抹掉，否则伴圆的碎片会被一起裁进图层、跟着缩放，
    # 屏幕上就出现"两个伴圆"（背景一个、移动图层一个）。
    # 抹掉时不能填纸色 —— 那块在圆内，填纸色会在滑动的圆上留一块白斑；
    # 沿同一方向取圆内 0.8r 处的像素（圆内是平滑渐变），补上去几乎看不出来。
    b_src = b.copy()
    bx0, by0, bw0, bh0 = b_box
    ccx, ccy = bx0 + bw0 / 2.0, by0 + bh0 / 2.0
    for _n, ex in (b_extra or []):
        ecx, ecy = ex[0] + ex[2] / 2.0, ex[1] + ex[3] / 2.0
        ddx, ddy = ecx - ccx, ecy - ccy
        norm = math.hypot(ddx, ddy) or 1.0
        # 取"伴圆内侧紧邻"的像素：往外 0.3 个框宽会正好落在伴圆盘上（取到深绿，
        # 等于用深绿把伴圆又填了一遍）。贴着伴圆边缘往里 8px 取，才是圆本身的颜色。
        back = ex[2] / 2.0 + 8.0
        sx = int(ecx - ddx / norm * back)
        sy = int(ecy - ddy / norm * back)
        fill = b.getpixel((max(0, min(b.width - 1, sx)),
                            max(0, min(b.height - 1, sy))))[:3]
        b_src = _erase_circle(b_src, ex, fill, feather=max(10, ex[2] * 0.10))
    b_c = _circle_layer(b_src, b_box)
    a_c = _circle_layer(a, a_box)
    b_same = b_c if b_c.size == a_c.size else b_c.resize(a_c.size, Image.LANCZOS)
    out = []
    hard = max(2, int(n * 0.16))          # 硬切只给前 16% 的帧，别拖成淡入淡出
    for i in range(n):
        p = ease_out((i + 1) / n)
        tb = min(1.0, (i + 1) / hard)
        canvas = Image.blend(a_bg, b_bg, tb).convert("RGBA")
        # 终点＝本页编排位置；起点＝圆心加 (dx,dy) 且尺寸乘 sc
        bx, by, bw, bh = b_box
        fx, fy = bx + bw / 2, by + bh / 2
        sx, sy = fx + dx, fy + dy
        cx = sx + (fx - sx) * p
        cy = sy + (fy - sy) * p
        w = bw * sc + (bw - bw * sc) * p
        h = bh * sc + (bh - bh * sc) * p
        x, y = cx - w / 2, cy - h / 2
        pad = max(3.0, w * 0.018)
        layer = Image.alpha_composite(a_c, _fade(b_same, tb))
        layer = layer.resize((max(2, int(w + 2 * pad)),
                              max(2, int(h + 2 * pad))), Image.LANCZOS)
        canvas.alpha_composite(layer, (int(round(x - pad)), int(round(y - pad))))
        # 叠层图形（伴圆 / 底圆）最后压上去：它们是本页固定在终点的元素，
        # 圆窗是"滑到它下面"的。这样既不重复、也不会在最后一帧突然跳出来。
        out.append(canvas.convert("RGB"))
    return out


def frames_for(a, b, trans, ms, fps, a_box=None, b_box=None):
    n = max(2, int(ms / 1000 * fps))
    out = []
    for i in range(n):
        t = ease_out((i + 1) / n)
        if trans.startswith("morph"):
            if a_box and b_box:
                # 交给锚点位移合成（圆窗真的滑过去）
                continue
            out.append(Image.blend(a, b, t))
        elif trans in ("fade", "dissolve"):
            out.append(Image.blend(a, b, t))
        elif trans.startswith("wheel"):
            sp = 8
            for d in ("1", "2", "3", "4", "6", "8", "12"):
                if trans == "wheel" + d:
                    sp = int(d)
            out.append(Image.composite(b, a, mask_wheel(a.size, t, sp)))
        elif trans in ("circle", "iris"):
            out.append(Image.composite(b, a, mask_circle(a.size, t)))
        elif trans.startswith("sweep"):
            out.append(Image.composite(b, a, mask_sweep(a.size, t)))
        elif trans.startswith("push"):
            W, H = a.size
            k = int(W * t)
            c = Image.new("RGB", a.size, (8, 6, 7))
            c.paste(a, (-k, 0)); c.paste(b, (W - k, 0))
            out.append(c)
        elif trans in ("zoom", "zoom_in"):
            # 新页从中心放大进入，旧页同时放大淡出（PowerPoint「缩放」观感）
            W, H = a.size
            s = 0.58 + 0.42 * t
            w2, h2 = max(1, int(W * s)), max(1, int(H * s))
            fg = b.resize((w2, h2), Image.LANCZOS)
            canvas = a.copy()
            canvas.paste(fg, ((W - w2) // 2, (H - h2) // 2))
            out.append(Image.blend(a, canvas, min(1.0, t * 1.6)))
        else:
            out.append(b if t > 0.5 else a)
    if trans.startswith("morph") and a_box and b_box:
        return morph_slide_frames(a, b, a_box, b_box, ms, fps)
    return out


def write_mp4(frames, out, fps):
    """帧序列 → H.264。

    为什么要 MP4：本片背景是近黑的大面积径向渐变，GIF 只有 256 色，
    不管怎么抖动都会出色带或噪点（看上去像画面脏了一块）。
    H.264 用 YUV420 渐变平滑，文件还更小。

    编码器要按序回退：本机 ffmpeg 是录屏版，h264_mf 会报
    MF_E_INVALIDMEDIATYPE，h264_amf 在没有对应硬件时也会挂。
    """
    FFMPEG = r"D:\ACLOS\Cross\recorder-release\ffmpeg.exe"
    if not os.path.exists(FFMPEG):
        print("!! 找不到 ffmpeg，跳过 mp4")
        return False
    W, H = frames[0].size
    for enc in ("libo264rt", "h264_amf", "h264_mf"):
        cmd = [FFMPEG, "-y", "-hide_banner", "-loglevel", "error",
               "-f", "rawvideo", "-pix_fmt", "rgb24", "-s", f"{W}x{H}",
               "-r", str(fps), "-i", "-",
               "-c:v", enc, "-b:v", "8M", "-pix_fmt", "yuv420p",
               "-movflags", "+faststart", out]
        p = subprocess.Popen(cmd, stdin=subprocess.PIPE,
                             stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
        try:
            for f in frames:
                p.stdin.write(f.tobytes())
            p.stdin.close()
        except BrokenPipeError:
            pass
        err = p.stderr.read().decode("utf-8", "ignore")
        if p.wait() == 0 and os.path.exists(out) and os.path.getsize(out) > 0:
            return True
        lines = [x for x in err.strip().splitlines() if x.strip()]
        print(f"!! {enc} 不可用，换下一个：{lines[-1] if lines else '未知错误'}")
    return False


def main():
    pptx, plans_json, out = sys.argv[1:4]
    width = _arg("--width", 640, int)
    hold = _arg("--hold", 1.2, float)
    fps = _arg("--fps", 15, int)
    mp4 = _arg("--mp4", None, str)

    plans = json.load(open(plans_json, encoding="utf-8"))
    tmp = os.path.join(os.path.dirname(out) or ".", "_render")
    pages = to_pages(pptx, tmp)

    imgs = []
    for p in pages:
        im = Image.open(p).convert("RGB")
        h = int(im.height * width / im.width)
        imgs.append(im.resize((width, h), Image.LANCZOS))
    size = imgs[0].size

    boxes = anchor_boxes(pptx)
    extras = bloom_boxes(pptx)
    k = width / 960.0
    def px(box):
        return None if box is None else (box[0]*k, box[1]*k, box[2]*k, box[3]*k)

    frames = []
    hold_n = max(1, int(hold * fps))
    for i, im in enumerate(imgs):
        if i == 0:
            frames += [im] * hold_n
            continue
        cfg = plans.get(str(i + 1), {})
        plan = cfg.get("plan", [])
        carry = next((it for it in plan if len(it) > 1 and it[1] == "carry"), None)
        if carry and boxes[i - 1] and boxes[i]:
            opts = carry[4] if len(carry) > 4 else {}
            frames += carry_frames(imgs[i - 1], im, px(boxes[i - 1]), px(boxes[i]),
                                   opts, carry[2] or 1000, fps,
                                   [(n, px(b)) for n, b in extras[i]])
        else:
            trans = cfg.get("transition") or "cut"
            ms = cfg.get("transition_ms") or 1000
            frames += frames_for(imgs[i - 1], im, trans, ms, fps,
                                 px(boxes[i - 1]), px(boxes[i]))
        frames += [im] * hold_n

    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    ms = int(1000 / fps)

    if mp4:
        if write_mp4(frames, mp4, fps):
            print(f"{len(imgs)} 页 {len(frames)} 帧 {len(frames)/fps:.1f}s -> {mp4}")

    # GIF 只有 256 色，这种近黑大幅渐变一定会出现"整块色带"（肉眼像一块竖着的脏斑）。
    # 做法：取一张中间帧算共享调色板，再对所有帧做误差扩散抖动。
    # 每帧各自调色板会让暗部逐帧跳色（闪烁），所以必须共享。
    ref = frames[len(frames) // 2]
    pal = ref.quantize(colors=255, method=Image.MEDIANCUT)
    qframes = [f.quantize(palette=pal, dither=Image.FLOYDSTEINBERG) for f in frames]

    qframes[0].save(out, save_all=True, append_images=qframes[1:],
                    duration=[ms] * len(qframes), loop=0,
                    optimize=False, disposal=1)
    print(f"{len(imgs)} 页 {len(qframes)} 帧 {len(qframes)/fps:.1f}s -> {out}")


if __name__ == "__main__":
    main()
