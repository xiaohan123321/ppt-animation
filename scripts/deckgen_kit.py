# -*- coding: utf-8 -*-
"""随机成稿引擎 —— 给一份内容 spec，长出一套版式 + 图形 + 动效。

和"模板"的区别
  模板：版式写死，换内容 = 填空。
  这里：每次运行按 seed 从词表里抽"锚点 / 图形母题 / 动效"，再按约束组装。
        同 seed 可复现，换 seed 就是另一套片子 —— 但都遵守同一套排版纪律
        （文字栏不被图形侵占、圆窗跨页必位移、深色处必压雾白）。

产出**成对**的 pptx 和 plan.json
  plan 里的 spid 是生成时记下来的真实 id，所以永远不会出现"找不到形状"。
  这是把 v7/v2 那套"手写 spid"彻底干掉的关键。

用法
  python scripts/gen_deck.py assets/spec_demo.json --seed 7 --out build/deck7
  → build/deck7/deck.pptx  build/deck7/plan.json  build/deck7/bg/*.jpg
"""
import argparse
import json
import math
import os
import random

from PIL import Image, ImageDraw, ImageFilter
import numpy as np
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt
from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
W, H = 960.0, 540.0

FONT = "STZhongsong"      # 华文中宋
TITLE = "STXingkai"       # 华文行楷
LATIN = "Times New Roman"


# ══════════════════════════════════════════════════════════════════
# 配色：三套皮肤，按 seed 选一套（不是每页换，是整套换）
# ══════════════════════════════════════════════════════════════════
PALETTES = {
    "qinglv": dict(                     # 青绿山水
        paper="F6F2E7", paper2="FAF6EC", ink="1F2E2C", ink2="4C5C58",
        ink3="7B8985", green="2E5A52", green_m="4A7C72", green_l="9DBAB0",
        line="D5CDBB", accent="A8804A",
        sky_t=(247, 243, 233), sky_b=(214, 228, 223),
        far1=(146, 180, 177), far2=(96, 141, 136), mid=(58, 102, 97),
        near=(30, 60, 58),
    ),
    "daiqing": dict(                    # 黛青（偏蓝灰，更冷）
        paper="F4F3EE", paper2="FAF9F4", ink="1E2830", ink2="4A5762",
        ink3="79858E", green="2C4A5C", green_m="446B7E", green_l="9AB2BF",
        line="D2D4CD", accent="9C7B4E",
        sky_t=(246, 244, 238), sky_b=(212, 222, 228),
        far1=(150, 176, 190), far2=(104, 138, 154), mid=(64, 96, 112),
        near=(32, 54, 66),
    ),
    "zhejin": dict(                     # 赭金（暖，偏宋画）
        paper="F7F1E4", paper2="FBF6EA", ink="2A241C", ink2="574E42",
        ink3="867C6E", green="4A3B28", green_m="6E5A40", green_l="BCA989",
        line="DCCFB6", accent="A87238",
        sky_t=(250, 244, 230), sky_b=(228, 218, 200),
        far1=(196, 176, 146), far2=(158, 132, 100), mid=(116, 92, 66),
        near=(66, 50, 36),
    ),
    "yanhui": dict(                     # 烟灰（冷灰，最克制）
        paper="F3F3F0", paper2="FAFAF7", ink="26282A", ink2="55595C",
        ink3="858A8D", green="3C4448", green_m="5E686D", green_l="A8B0B4",
        line="D6D8D4", accent="8C7B64",
        sky_t=(246, 246, 242), sky_b=(218, 221, 219),
        far1=(176, 184, 186), far2=(132, 143, 146), mid=(92, 102, 106),
        near=(52, 60, 64),
    ),
    "qingci": dict(                     # 青瓷（浅青，清透）
        paper="F2F6F4", paper2="F8FBFA", ink="1D2E2C", ink2="48605C",
        ink3="7A8E89", green="2A5E58", green_m="4C8880", green_l="A6C6C0",
        line="CFDCD7", accent="A9834E",
        sky_t=(242, 248, 246), sky_b=(206, 224, 221),
        far1=(154, 190, 187), far2=(106, 152, 148), mid=(64, 110, 106),
        near=(34, 68, 65),
    ),
    "zitan": dict(                      # 紫檀（深褐紫，厚重）
        paper="F5EFE9", paper2="FBF6F0", ink="2A2026", ink2="574044",
        ink3="86706F", green="4A2F38", green_m="6E4A54", green_l="BCA0A4",
        line="DCCEC8", accent="9C6B3E",
        sky_t=(248, 240, 234), sky_b=(226, 212, 208),
        far1=(190, 166, 164), far2=(152, 120, 120), mid=(112, 82, 84),
        near=(64, 44, 48),
    ),
    "songlv": dict(                     # 松绿（深绿，庄重）
        paper="F4F2E6", paper2="FAF8EE", ink="1C2A22", ink2="45564A",
        ink3="77877A", green="1F4A34", green_m="376B4E", green_l="96B49C",
        line="D4D0BC", accent="A67B3E",
        sky_t=(246, 244, 232), sky_b=(214, 224, 212),
        far1=(148, 178, 158), far2=(104, 140, 118), mid=(64, 98, 78),
        near=(32, 58, 44),
    ),
    "xuantie": dict(                    # 玄铁（冷钢灰 + 铁锈橙，工业感新风格）
        paper="F2F3F4", paper2="F8F9FA", ink="1B2226", ink2="454F55",
        ink3="767F85", green="2E3E48", green_m="4A5E6A", green_l="9FAEB6",
        line="D3D7DA", accent="B4622D",
        sky_t=(243, 244, 245), sky_b=(210, 217, 221),
        far1=(160, 172, 180), far2=(112, 126, 136), mid=(70, 84, 94),
        near=(34, 44, 52),
    ),
    "jiangzhu": dict(                   # 绛朱（深红棕 + 鎏金，宫墙新风格）
        paper="F7F0EA", paper2="FBF5EF", ink="2B1E1A", ink2="58423A",
        ink3="86706A", green="4E2E28", green_m="74453C", green_l="BFA094",
        line="DFC9BE", accent="A87E3C",
        sky_t=(249, 242, 236), sky_b=(229, 214, 208),
        far1=(198, 168, 158), far2=(158, 120, 110), mid=(116, 80, 72),
        near=(60, 38, 34),
    ),
}


def _hex_rgb(c):
    """hex 字符串 / rgb 元组 → (r,g,b) 整数元组。"""
    if isinstance(c, str):
        return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))
    return tuple(int(v) for v in c[:3])


def photo_grade(im, pal, dark=1.0, size=None, keep_color=0.22):
    """实拍照片 → 配色系统里的双色调风景。

    灰度后按 阴影色(near)→中间色(green_m)→高光色(paper2) 上色，再掺回少量
    原色 —— 风格和程序山水一致，但质感是真实照片。dark=1.0 是窗景浓度
    （阴影更深），整页底图用 0.6~0.7（亮一点才压得住正文）。
    """
    from PIL import ImageOps
    if size is None:
        size = (BW, BH)
    tw, th = size
    sw, sh = im.size
    sc = max(tw / sw, th / sh)
    im = im.resize((int(sw * sc) + 1, int(sh * sc) + 1), Image.LANCZOS)
    w, h = im.size
    im = im.crop(((w - tw) // 2, (h - th) // 2,
                  (w - tw) // 2 + tw, (h - th) // 2 + th))
    g = ImageOps.autocontrast(ImageOps.grayscale(im), cutoff=2)
    lo = tuple(int(v * (0.62 - 0.25 * dark)) for v in _hex_rgb(pal["near"]))
    mid = _hex_rgb(pal["green_m"])
    hi = _hex_rgb(pal["paper2"])
    out = ImageOps.colorize(g, black=lo, white=hi, mid=mid).convert("RGB")
    from PIL import ImageEnhance
    out = ImageEnhance.Contrast(out).enhance(1.16)
    out = ImageEnhance.Color(out).enhance(1.05)
    if keep_color > 0:
        out = Image.blend(out, im, keep_color)
    return out


def C(h):
    return RGBColor.from_string(h)


def emu(v):
    return Emu(int(v * 12700))


# ══════════════════════════════════════════════════════════════════
# 图元
# ══════════════════════════════════════════════════════════════════
def _strip_style(sp):
    """删掉 p:style —— 不删的话主题的 effectRef 会画出投影。"""
    sp.shadow.inherit = False
    st = sp._element.find(qn("p:style"))
    if st is not None:
        sp._element.remove(st)
    return sp


def _fill_alpha(sp, alpha):
    if not alpha:
        return
    sf = sp._element.spPr.find(qn("a:solidFill"))
    if sf is None:
        return
    clr = sf.find(qn("a:srgbClr"))
    if clr is None:
        return
    al = etree.SubElement(clr, qn("a:alpha"))
    al.set("val", str(int((100 - alpha) * 1000)))


def _line_alpha(sp, alpha):
    ln = sp._element.spPr.find(qn("a:ln"))
    sf = ln.find(qn("a:solidFill"))
    clr = sf.find(qn("a:srgbClr"))
    al = etree.SubElement(clr, qn("a:alpha"))
    al.set("val", str(int((100 - alpha) * 1000)))


def rect(s, x, y, w, h, color, alpha=0, name=None, rot=None):
    sp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, emu(x), emu(y), emu(w), emu(h))
    _strip_style(sp)
    sp.fill.solid()
    sp.fill.fore_color.rgb = C(color)
    sp.line.fill.background()
    _fill_alpha(sp, alpha)
    if rot:
        sp.rotation = rot
    if name:
        sp.name = name
    return sp


def rrect(s, x, y, w, h, color, alpha=0, name=None, border=None, radius=0.05):
    sp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, emu(x), emu(y), emu(w), emu(h))
    _strip_style(sp)
    sp.fill.solid()
    sp.fill.fore_color.rgb = C(color)
    sp.line.fill.background()
    geom = sp._element.spPr.find(qn("a:prstGeom"))
    av = geom.find(qn("a:avLst"))
    if av is None:
        av = etree.SubElement(geom, qn("a:avLst"))
    gd = etree.SubElement(av, qn("a:gd"))
    gd.set("name", "adj")
    gd.set("fmla", f"val {int(radius * 100000)}")
    _fill_alpha(sp, alpha)
    if border:
        sp.line.color.rgb = C(border)
        sp.line.width = Pt(0.75)
    if name:
        sp.name = name
    return sp


def disc(s, cx, cy, d, color, alpha=0, name=None):
    sp = s.shapes.add_shape(MSO_SHAPE.OVAL, emu(cx - d / 2), emu(cy - d / 2), emu(d), emu(d))
    _strip_style(sp)
    sp.fill.solid()
    sp.fill.fore_color.rgb = C(color)
    sp.line.fill.background()
    _fill_alpha(sp, alpha)
    if name:
        sp.name = name
    return sp


def ring(s, cx, cy, d, color, w=1.0, alpha=None, name=None):
    sp = s.shapes.add_shape(MSO_SHAPE.OVAL, emu(cx - d / 2), emu(cy - d / 2), emu(d), emu(d))
    _strip_style(sp)
    sp.fill.background()
    sp.line.color.rgb = C(color)
    sp.line.width = Pt(w)
    if alpha:
        _line_alpha(sp, alpha)
    if name:
        sp.name = name
    return sp


def arc(s, cx, cy, d, a0, a1, color, w=1.0, alpha=None, name=None):
    """a0/a1 单位是度，从 3 点方向顺时针。粗描边就是"厚环段"。"""
    sp = s.shapes.add_shape(MSO_SHAPE.ARC, emu(cx - d / 2), emu(cy - d / 2), emu(d), emu(d))
    _strip_style(sp)
    sp.fill.background()
    sp.line.color.rgb = C(color)
    sp.line.width = Pt(w)
    geom = sp._element.spPr.find(qn("a:prstGeom"))
    av = geom.find(qn("a:avLst"))
    if av is None:
        av = etree.SubElement(geom, qn("a:avLst"))
    for nm, val in (("adj1", a0), ("adj2", a1)):
        gd = etree.SubElement(av, qn("a:gd"))
        gd.set("name", nm)
        gd.set("fmla", f"val {int(val * 60000)}")
    if alpha:
        _line_alpha(sp, alpha)
    if name:
        sp.name = name
    return sp


def band(s, cx, cy, d, a0, a1, th, color, alpha=None, name=None, rot=None):
    """实心厚环段（blockArc）—— 比"粗描边弧"重得多，读作一整块环形砖。

    th 是厚度占半径的比例（0~0.5）；a0/a1 度，从 3 点方向顺时针。
    厚带 + 细缘线交替排布，就是"层层镶嵌"的基本单元。
    """
    sp = s.shapes.add_shape(MSO_SHAPE.BLOCK_ARC,
                            emu(cx - d / 2), emu(cy - d / 2), emu(d), emu(d))
    _strip_style(sp)
    sp.fill.solid()
    sp.fill.fore_color.rgb = C(color)
    sp.line.fill.background()
    geom = sp._element.spPr.find(qn("a:prstGeom"))
    av = geom.find(qn("a:avLst"))
    if av is None:
        av = etree.SubElement(geom, qn("a:avLst"))
    for nm, val in (("adj1", a0), ("adj2", a1), ("adj3", th)):
        gd = etree.SubElement(av, qn("a:gd"))
        gd.set("name", nm)
        gd.set("fmla", f"val {int(val * (60000 if nm != 'adj3' else 100000))}")
    _fill_alpha(sp, alpha)
    if rot:
        sp.rotation = rot
    if name:
        sp.name = name
    return sp


def frame(s, x, y, w, h, color, w_pt=6.0, alpha=None, radius=0.16, rot=None, name=None):
    """空心厚框（圆角矩形描边）—— 异形镶嵌用：圆带之间穿插一道方环。"""
    sp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, emu(x), emu(y), emu(w), emu(h))
    _strip_style(sp)
    sp.fill.background()
    sp.line.color.rgb = C(color)
    sp.line.width = Pt(w_pt)
    geom = sp._element.spPr.find(qn("a:prstGeom"))
    av = geom.find(qn("a:avLst"))
    if av is None:
        av = etree.SubElement(geom, qn("a:avLst"))
    gd = etree.SubElement(av, qn("a:gd"))
    gd.set("name", "adj")
    gd.set("fmla", f"val {int(radius * 100000)}")
    if alpha:
        _line_alpha(sp, alpha)
    if rot:
        sp.rotation = rot
    if name:
        sp.name = name
    return sp


def pie(s, cx, cy, d, a0, a1, color, alpha=None, name=None, rot=None):
    """扇形（pie）：四象限棋盘圆、吃豆人楔口的单元。a0/a1 度，3 点方向顺时针。"""
    sp = s.shapes.add_shape(MSO_SHAPE.PIE, emu(cx - d / 2), emu(cy - d / 2), emu(d), emu(d))
    _strip_style(sp)
    sp.fill.solid()
    sp.fill.fore_color.rgb = C(color)
    sp.line.fill.background()
    geom = sp._element.spPr.find(qn("a:prstGeom"))
    av = geom.find(qn("a:avLst"))
    if av is None:
        av = etree.SubElement(geom, qn("a:avLst"))
    for nm, val in (("adj1", a0), ("adj2", a1)):
        gd = etree.SubElement(av, qn("a:gd"))
        gd.set("name", nm)
        gd.set("fmla", f"val {int(val * 60000)}")
    _fill_alpha(sp, alpha)
    if rot:
        sp.rotation = rot
    if name:
        sp.name = name
    return sp


def txt(s, x, y, w, h, t, size, color, font=None, align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP, bold=False, alpha=0, name=None, lsp=1.15, wrap=True):
    tb = s.shapes.add_textbox(emu(x), emu(y), emu(w), emu(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.auto_size = None
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, ln in enumerate(str(t).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = ln
        p.alignment = align
        p.line_spacing = lsp
        for run in p.runs:
            run.font.size = Pt(size)
            run.font.name = font or FONT
            run.font.bold = bold
            run.font.color.rgb = C(color)
            # 中文必须写 a:ea，否则落回主题默认字体（书法体失效的根因）
            rPr = run.font._rPr
            ea = rPr.find(qn("a:ea"))
            if ea is None:
                ea = etree.SubElement(rPr, qn("a:ea"))
                latin = rPr.find(qn("a:latin"))
                if latin is not None:
                    latin.addnext(ea)
            ea.set("typeface", font or FONT)
            if alpha:
                sf = run.font._rPr.find(qn("a:solidFill"))
                clr = sf.find(qn("a:srgbClr")) if sf is not None else None
                if clr is not None:
                    al = etree.SubElement(clr, qn("a:alpha"))
                    al.set("val", str(int((100 - alpha) * 1000)))
    if name:
        tb.name = name
    return tb


# 窗口不一定得是圆的 —— 这些都是预设几何，改个 prst 就换形状
PRESETS = {
    "circle": "ellipse",
    "squircle": "roundRect",
    "hex": "hexagon",
    "pent": "pentagon",
    "arch": "arch",
    "para": "parallelogram",
    "teardrop": "teardrop",
}


def pic_shape(s, path, cx, cy, d, prst="ellipse", adj=None, name="coreImg"):
    """图片裁成任意预设形状。roundRect 的圆角要压小，否则像药片。"""
    p = s.shapes.add_picture(path, emu(cx - d / 2), emu(cy - d / 2), emu(d), emu(d))
    p.name = name
    geom = p._element.spPr.find(qn("a:prstGeom"))
    if geom is None:
        return p
    geom.set("prst", prst)
    av = geom.find(qn("a:avLst"))
    if av is None:
        av = etree.SubElement(geom, qn("a:avLst"))
    if adj is not None:
        gd = etree.SubElement(av, qn("a:gd"))
        gd.set("name", "adj")
        gd.set("fmla", f"val {int(adj * 100000)}")
    return p


def pic_round(s, path, cx, cy, d, name="coreImg"):
    return pic_shape(s, path, cx, cy, d, "ellipse", name=name)


# ══════════════════════════════════════════════════════════════════
# 山水底图（远层 / 近层 / 圆窗浓墨版）
# ══════════════════════════════════════════════════════════════════
BW, BH = 1440, 810
_yy, _xx = np.mgrid[0:BH, 0:BW].astype(np.float32)


def _peaks(rng, n, span=(-0.15, 1.15), h=(0.45, 1.0), w=(0.07, 0.17)):
    out = []
    for _ in range(n):
        out.append((rng.uniform(*span) * BW, rng.uniform(*h),
                    rng.uniform(*w) * BW, rng.random() < 0.4))
    return out


def _ridge(ps, base, amp, rng, jag=0.0022):
    """山脊线。先减中位数让山谷沉下去，否则宽高斯会叠成一条色带。"""
    d = _xx[0]
    prof = np.zeros_like(d)
    for px, hh, ww, sharp in ps:
        t = np.abs(d - px) / ww
        prof += (1.0 / (1.0 + t ** 1.7) if sharp else np.exp(-t ** 2)) * hh
    prof -= np.median(prof)
    prof = prof / max(np.percentile(prof, 92), 1e-6)
    x = d / BW
    prof += 0.28 * np.sin(x * np.pi * 1.25 + rng.uniform(0, 6.28))
    prof += 0.14 * np.sin(x * np.pi * 2.6 + rng.uniform(0, 6.28))
    n1 = np.convolve(rng.normal(0, 1, BW).astype(np.float32), np.ones(17) / 17, mode="same")
    return base * BH - (prof * amp * BH + n1 * jag * BH)


def _fill(img, ry, c_top, c_bot):
    m = np.clip(_yy - ry[None, :] + 0.5, 0.0, 1.0)      # 1px 抗锯齿，否则山脊一圈白边
    t = (((_yy - ry[None, :]) / (BH * 0.55)).clip(0, 1) ** 0.85)[..., None]
    layer = c_top[None, None, :] * (1 - t) + c_bot[None, None, :] * t
    img[:] = img * (1 - m[..., None]) + layer * m[..., None]


def _ridge_line(img, ry, color, alpha=0.22, width=2):
    im = Image.fromarray(np.clip(img, 0, 255).astype(np.uint8))
    d = ImageDraw.Draw(im, "RGBA")
    d.line([(i, float(ry[i]) + width * 0.5) for i in range(0, BW, 2)],
           fill=tuple(int(v) for v in color) + (int(alpha * 255),), width=width)
    return np.asarray(im, np.float32)


def _mist(img, y0, strength, thick):
    band = np.exp(-(((_yy - y0 * BH) / (thick * BH)) ** 2))
    a = (band * strength)[..., None]
    return img * (1 - a) + np.array([250, 249, 244], np.float32)[None, None, :] * a


def _haze(c, k):
    return np.asarray(c, np.float32) * (1 - k) + np.array([250, 249, 244], np.float32) * k


def _paint_mount(pal, seed, quiet=None, mtn=1.0, low=0.0):
    """返回 (far RGB, near RGBA)。两层分别做漂移就是"远近变化"。"""
    rng = np.random.default_rng(seed)
    pal = {k: (np.array(v, np.float32) if isinstance(v, (tuple, list)) else v)
           for k, v in pal.items()}
    sky_t = pal["sky_t"]
    sky_b = np.array(pal["sky_b"], np.float32)
    far1 = np.array(pal["far1"], np.float32)
    far2 = np.array(pal["far2"], np.float32)
    mid = np.array(pal["mid"], np.float32)
    near = np.array(pal["near"], np.float32)

    t = (_yy / BH) ** 0.9
    img = sky_t[None, None, :] * (1 - t[..., None]) + sky_b[None, None, :] * t[..., None]

    damp = None
    if quiet == "L":
        k = np.clip((_xx[0] / BW - 0.05) / 0.45, 0, 1)
        damp = 0.35 + 0.65 * k
    elif quiet == "R":
        k = np.clip((0.95 - _xx[0] / BW) / 0.45, 0, 1)
        damp = 0.35 + 0.65 * k

    r1 = _ridge(_peaks(rng, 7, h=(0.45, 0.95), w=(0.11, 0.22)), 0.66 + low, 0.20 * mtn, rng)
    if damp is not None:
        r1 = r1 + (1 - damp) * BH * 0.09
    _fill(img, r1, far1, _haze(far1, 0.72))
    img = _ridge_line(img, r1, np.array((22, 34, 33), np.float32), 0.12)
    img = _mist(img, 0.66, 0.46, 0.070)

    r2 = _ridge(_peaks(rng, 6, h=(0.50, 1.00), w=(0.09, 0.17)), 0.78 + low * 0.9, 0.26 * mtn, rng)
    if damp is not None:
        r2 = r2 + (1 - damp) * BH * 0.10
    _fill(img, r2, far2, _haze(far2, 0.80))
    img = _ridge_line(img, r2, np.array((22, 34, 33), np.float32), 0.18)
    img = _mist(img, 0.79, 0.50, 0.075)

    far = np.clip(img + rng.normal(0, 2.0, (BH, BW, 1)), 0, 255).astype(np.uint8)

    nimg = np.zeros((BH, BW, 3), np.float32) + pal["sky_t"]
    r3 = _ridge(_peaks(rng, 4, h=(0.55, 1.00), w=(0.10, 0.19)), 0.90 + low * 0.8, 0.20 * mtn, rng)
    if damp is not None:
        r3 = r3 + (1 - damp) * BH * 0.12
    _fill(nimg, r3, mid, _haze(mid, 0.55))
    nimg = _ridge_line(nimg, r3, np.array((22, 34, 33), np.float32), 0.26, 3)

    r4 = _ridge(_peaks(rng, 3, span=(-0.2, 1.2), h=(0.35, 0.80), w=(0.16, 0.32)),
                1.00, 0.16 * mtn, rng)
    _fill(nimg, r4, near, _haze(near, 0.42))
    nimg = _ridge_line(nimg, r4, np.array((22, 34, 33), np.float32), 0.40, 3)

    # alpha 直接用填充遮罩：之前用"山脊上方 6px 起算"的台阶，
    # 那 6px 的 RGB 还是雾白色，被半透明地叠上去 → 山脊一圈白边。
    alpha = np.clip(_yy - r3[None, :] + 0.5, 0.0, 1.0)
    near_rgba = np.dstack([np.clip(nimg, 0, 255), alpha * 255]).astype(np.uint8)

    # 圆窗里的浓墨版：压暗 + 提饱和，圆内外的明暗差就是"镜头感"
    a = far.astype(np.float32) / 255.0
    a = a ** 1.45
    lum = a @ np.array([0.299, 0.587, 0.114], np.float32)
    a = np.clip(lum[..., None] + (a - lum[..., None]) * 1.55, 0, 1)
    a = a * np.array([0.50, 0.60, 0.60], np.float32)
    win = (np.clip((a - 0.5) * 1.18 + 0.5, 0, 1) * 255).astype(np.uint8)
    return Image.fromarray(far), Image.fromarray(near_rgba, "RGBA"), Image.fromarray(win)


# ══════════════════════════════════════════════════════════════════
# 文本自适应：PowerPoint 没有文本测量 API，只能估。
# 估算准了，就能"字号自己退到放得下"，而不是溢出压住下一块。
# ══════════════════════════════════════════════════════════════════
def _adv(ch, size):
    """单字符前进宽度（pt）。中文 1em，西文/数字 0.52~0.56em。"""
    if len(ch.encode("gbk", errors="ignore")) > 1:
        return size * 1.0
    if ch.isdigit():
        return size * 0.56
    return size * 0.52


def line_count(t, size, box_w):
    if not t:
        return 1
    n, cur = 1, 0.0
    for ch in t:
        a = _adv(ch, size)
        if cur + a > box_w and cur > 0:
            n += 1
            cur = a
        else:
            cur += a
    return n


def text_h(t, size, box_w, lsp):
    return sum(line_count(p, size, box_w) for p in str(t).split("\n")) * size * lsp


def fit_size(t, box_w, box_h, smax, smin=5.5, lsp=1.45, safety=0.96):
    """二分出"最大且放得下"的字号。返回 (size, ok)。"""
    box_h *= safety
    if text_h(t, smin, box_w, lsp) > box_h:
        return smin, False
    lo, hi = smin, smax
    for _ in range(16):
        mid = (lo + hi) / 2
        if text_h(t, mid, box_w, lsp) <= box_h:
            lo = mid
        else:
            hi = mid
    return round(lo, 1), True


def truncate_to(t, size, box_w, max_h, lsp):
    """按行截断，末尾补省略号 —— 字号退到下限还放不下时的兜底。"""
    keep, cur_h = [], 0.0
    for p in str(t).split("\n"):
        n = line_count(p, size, box_w)
        if cur_h + n * size * lsp <= max_h:
            keep.append(p)
            cur_h += n * size * lsp
            continue
        # 这一行只留能塞下的部分
        room = int((max_h - cur_h) / (size * lsp))
        if room > 0:
            per = max(1, int(len(p) * room / max(1, n)))
            keep.append(p[:max(1, per - 1)] + "…")
        break
    return "\n".join(keep) if keep else "…"


OVERFLOWS = []          # [(name, 是否截断)]，生成器用它报告"哪页被压了"


def atxt(s, x, y, w, h, t, color, smax=12, smin=5.5, lsp=1.45, font=None,
         bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, name=None,
         alpha=0, wrap=True):
    """自动字号 + 溢出截断。所有正文都该走这里，不要直接写死字号。"""
    if not str(t).strip():
        return None
    size, ok = fit_size(t, w, h, smax, smin, lsp)
    body = t
    if not ok:
        body = truncate_to(t, size, w, h * 0.96, lsp)
    tb = txt(s, x, y, w, h, body, size, color, font=font, bold=bold, align=align,
             anchor=anchor, lsp=lsp, name=name, alpha=alpha, wrap=wrap)
    tb._fitted_size = size
    tb._overflow = not ok
    if not ok or size < smax * 0.72:
        OVERFLOWS.append((name or "?", not ok))
    return tb


# ══════════════════════════════════════════════════════════════════
# 更多底图：雾 / 剪影 / 泼墨 / 网格
# ══════════════════════════════════════════════════════════════════
def _arr(v):
    return np.array(v, np.float32) if isinstance(v, (tuple, list)) else v


def _sky_of(pal):
    t = (_yy / BH) ** 0.9
    return (_arr(pal["sky_t"])[None, None, :] * (1 - t[..., None])
            + _arr(pal["sky_b"])[None, None, :] * t[..., None])


def _grade(far):
    """浓墨版：压暗 + 提饱和，圆窗里外才像"同一处景、不同曝光"。"""
    a = far.astype(np.float32) / 255.0
    a = a ** 1.45
    lum = a @ np.array([0.299, 0.587, 0.114], np.float32)
    a = np.clip(lum[..., None] + (a - lum[..., None]) * 1.55, 0, 1)
    a = a * np.array([0.50, 0.60, 0.60], np.float32)
    return (np.clip((a - 0.5) * 1.18 + 0.5, 0, 1) * 255).astype(np.uint8)


def _empty_near():
    return Image.fromarray(np.zeros((BH, BW, 4), np.uint8), "RGBA")


def _paint_mist(pal, seed, quiet=None, mtn=1.0, low=0.0):
    """纯雾：只有天空和几道雾带，最干净的一种。"""
    rng = np.random.default_rng(seed)
    img = _sky_of(pal)
    r = _ridge(_peaks(rng, 5, h=(0.5, 0.95), w=(0.16, 0.30)), 0.80, 0.13, rng)
    _fill(img, r, _arr(pal["far1"]), _haze(pal["far1"], 0.55))   # 极淡远山，只为给窗口一点纹理
    for i in range(4):
        y0 = 0.22 + 0.19 * i + rng.uniform(-0.03, 0.03)
        img = _mist(img, y0, rng.uniform(0.25, 0.5), rng.uniform(0.05, 0.10))
    far = np.clip(img + rng.normal(0, 1.8, (BH, BW, 1)), 0, 255).astype(np.uint8)
    return Image.fromarray(far), _empty_near(), Image.fromarray(_grade(far))


def _paint_ridges(pal, seed, quiet=None, mtn=1.0, low=0.0):
    """剪影：两层硬边山脊，不铺雾，海报感更强。"""
    rng = np.random.default_rng(seed)
    img = _sky_of(pal)
    r1 = _ridge(_peaks(rng, 6, h=(0.5, 1.0), w=(0.10, 0.20)), 0.72 + low, 0.30 * mtn, rng)
    _fill(img, r1, _arr(pal["far2"]), _haze(pal["far2"], 0.25))
    r2 = _ridge(_peaks(rng, 4, h=(0.55, 1.0), w=(0.12, 0.24)), 0.88 + low * 0.8, 0.24 * mtn, rng)
    _fill(img, r2, _arr(pal["near"]), _haze(pal["near"], 0.12))
    img = _ridge_line(img, r2, np.array((22, 34, 33), np.float32), 0.35, 3)
    far = np.clip(img + rng.normal(0, 1.5, (BH, BW, 1)), 0, 255).astype(np.uint8)
    return Image.fromarray(far), _empty_near(), Image.fromarray(_grade(far))


def _paint_ink(pal, seed, quiet=None, mtn=1.0, low=0.0):
    """泼墨：几团柔边墨块 + 两笔飞白，抽象，配极简内容最好。"""
    rng = np.random.default_rng(seed)
    img = _sky_of(pal)
    dark = _arr(pal["near"])
    for _ in range(5):
        cx, cy = rng.uniform(0, 1) * BW, rng.uniform(0.15, 1.05) * BH
        sx, sy = rng.uniform(0.14, 0.34) * BW, rng.uniform(0.06, 0.18) * BH
        g = np.exp(-(((_xx - cx) / sx) ** 2 + ((_yy - cy) / sy) ** 2))
        img = img + (dark - img) * (g * rng.uniform(0.10, 0.28))[..., None]
    far = np.clip(img + rng.normal(0, 1.8, (BH, BW, 1)), 0, 255).astype(np.uint8)
    im = Image.fromarray(far)
    d = ImageDraw.Draw(im, "RGBA")
    for _ in range(2):
        y0 = rng.uniform(0.5, 0.95) * BH
        d.line([(0, y0), (BW, y0 + rng.uniform(-0.06, 0.06) * BH)],
               fill=(30, 36, 34, 26), width=int(rng.uniform(2, 5)))
    return im, _empty_near(), Image.fromarray(_grade(np.asarray(im)))


def _paint_grid(pal, seed, quiet=None, mtn=1.0, low=0.0):
    """纸纹网格：最素的一种，适合纯文字页。"""
    rng = np.random.default_rng(seed)
    img = _sky_of(pal)
    far = np.clip(img + rng.normal(0, 1.2, (BH, BW, 1)), 0, 255).astype(np.uint8)
    im = Image.fromarray(far)
    d = ImageDraw.Draw(im, "RGBA")
    step = int(rng.choice([36, 48, 60]))
    for x in range(0, BW, step):
        d.line([(x, 0), (x, BH)], fill=(30, 46, 42, 12), width=1)
    for y in range(0, BH, step):
        d.line([(0, y), (BW, y)], fill=(30, 46, 42, 12), width=1)
    return im, _empty_near(), Image.fromarray(_grade(np.asarray(im)))


BG_KINDS = ("mount", "mist", "ridges", "ink", "grid", "yuan")


def _paint_yuan(pal, seed, quiet=None, mtn=1.0, low=0.0):
    """远景全幅 —— 参考图4/5：山铺满整页、雾感重、推到天际，
    远处一座塔影点景。配 push 慢推近就是"一张幻灯片里的远景演变"。
    雾都压得重，正文压上去仍然读得清。"""
    rng = np.random.default_rng(seed)
    pal = {k: (np.array(v, np.float32) if isinstance(v, (tuple, list)) else v)
           for k, v in pal.items()}
    sky_t = pal["sky_t"]
    sky_b = np.array(pal["sky_b"], np.float32)
    dark = np.array((22, 34, 33), np.float32)

    t = (_yy / BH) ** 0.72
    img = sky_t[None, None, :] * (1 - t[..., None]) + sky_b[None, None, :] * t[..., None]

    # 五层山：基线逐层下移 → 山越叠越高、铺满画幅；雾逐层加重保对比
    layers = [
        (np.array(pal["far1"], np.float32), 0.22 + low * 0.3, 0.22 * mtn, 6, 0.80),
        (np.array(pal["far2"], np.float32), 0.38 + low * 0.4, 0.26 * mtn, 6, 0.68),
        (np.array(pal["mid"], np.float32), 0.54 + low * 0.5, 0.28 * mtn, 5, 0.50),
        (np.array(pal["near"], np.float32), 0.72 + low * 0.5, 0.26 * mtn, 4, 0.30),
    ]
    ridges = []
    for c, base, amp, n, haze in layers:
        r = _ridge(_peaks(rng, n, h=(0.55, 1.0), w=(0.09, 0.20)), base, amp, rng)
        _fill(img, r, c, _haze(c, haze))
        img = _ridge_line(img, r, dark, 0.15)
        img = _mist(img, base + 0.12, 0.55, 0.085)
        ridges.append(r)

    # 远塔剪影：立在第二层山最高峰旁，三叠檐
    im = Image.fromarray(np.clip(img, 0, 255).astype(np.uint8))
    d = ImageDraw.Draw(im, "RGBA")
    px = float(np.clip(np.argmin(ridges[1]) + rng.uniform(-160, 160), 0.12 * BW, 0.88 * BW))
    py = float(ridges[1][int(px)] - BH * 0.02)
    col = tuple(int(v) for v in np.array(pal["near"])) + (150,)
    for w, h in ((0.030, 0.016), (0.021, 0.014), (0.013, 0.012)):
        d.polygon([(px - BW * w / 2, py), (px + BW * w / 2, py),
                   (px + BW * w * 0.30, py - BH * h), (px - BW * w * 0.30, py - BH * h)],
                  fill=col)
        py -= BH * h * 0.92
    d.polygon([(px - BW * 0.004, py - BH * 0.012), (px + BW * 0.004, py - BH * 0.012),
               (px, py - BH * 0.026)], fill=col)

    # 底部水光：一条亮带 + 倒影雾
    d.rectangle([0, BH * 0.90, BW, BH], fill=(255, 255, 255, 26))
    a = np.asarray(im).astype(np.float32)
    a = _mist(a, 0.94, 0.60, 0.10)
    far = np.clip(a + rng.normal(0, 2.0, (BH, BW, 1)), 0, 255).astype(np.uint8)

    # 近层：给圆窗保留一道更近的深色山脊（视差层）
    nimg = np.zeros((BH, BW, 3), np.float32) + pal["sky_t"]
    r5 = _ridge(_peaks(rng, 4, h=(0.55, 1.0), w=(0.10, 0.19)), 0.92, 0.18 * mtn, rng)
    _fill(nimg, r5, np.array(pal["near"], np.float32), _haze(np.array(pal["near"], np.float32), 0.42))
    nimg = _ridge_line(nimg, r5, dark, 0.38, 3)
    alpha = np.clip(_yy - r5[None, :] + 0.5, 0.0, 1.0)
    near_rgba = np.dstack([np.clip(nimg, 0, 255), alpha * 255]).astype(np.uint8)

    # 圆窗浓墨版（同 mount 的处理）
    a = far.astype(np.float32) / 255.0
    a = a ** 1.45
    lum = a @ np.array([0.299, 0.587, 0.114], np.float32)
    a = np.clip(lum[..., None] + (a - lum[..., None]) * 1.55, 0, 1)
    a = a * np.array([0.50, 0.60, 0.60], np.float32)
    win = (np.clip((a - 0.5) * 1.18 + 0.5, 0, 1) * 255).astype(np.uint8)
    return Image.fromarray(far), Image.fromarray(near_rgba, "RGBA"), Image.fromarray(win)


def paint_bg(pal, seed, kind="mount", quiet=None, mtn=1.0, low=0.0):
    """底图调度：kind 决定画法，pal 决定色系。"""
    fn = {"mount": _paint_mount, "mist": _paint_mist, "ridges": _paint_ridges,
          "ink": _paint_ink, "grid": _paint_grid, "yuan": _paint_yuan}.get(kind, _paint_mount)
    return fn(pal, seed, quiet=quiet, mtn=mtn, low=low)
