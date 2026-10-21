# -*- coding: utf-8 -*-
"""整份演示的页面转场预览。

用法：
    python render_transition.py <pptx> <plans.json> <out.gif>
        [--width 960] [--hold 1.0] [--trans 0.6] [--fps 20]

支持 iris/circle、wheel、morph、push、split、wipe、cover、doors、zoom、
blinds、checker、strips、cut、fade 等常用转场。
"""
import json
import math
import sys
import zipfile

from PIL import Image, ImageDraw

import render_gif as RG


def page_final(pptx, idx):
    with zipfile.ZipFile(pptx) as z:
        shapes = RG.parse_slide(z, idx)
        bg = RG.background_image(z, idx)
    final = bg.copy()
    for s in shapes:
        if s.get("name", "").startswith("transition_"):
            continue
        state = {"in": ("fade", 1.0, "out"), "emph": []}
        RG.draw_shape(final, s, state)
    return final.convert("RGB"), bg, shapes


def _mask(size, kind, t):
    W, H = size
    m = Image.new("L", size, 0)
    d = ImageDraw.Draw(m)
    if kind in ("circle", "iris"):
        r = math.hypot(W, H) * t
        d.ellipse([W / 2 - r, H / 2 - r, W / 2 + r, H / 2 + r], fill=255)
    elif kind == "wheel":
        r = math.hypot(W, H)
        d.pieslice([W / 2 - r, H / 2 - r, W / 2 + r, H / 2 + r],
                   -90, -90 + 360 * t, fill=255)
    elif kind in ("split_h", "doors_h"):
        half = int(H / 2 * t)
        d.rectangle([0, H // 2 - half, W, H // 2 + half], fill=255)
    elif kind in ("split_v", "doors_v"):
        half = int(W / 2 * t)
        d.rectangle([W // 2 - half, 0, W // 2 + half, H], fill=255)
    elif kind == "blinds":
        n = 10
        strip = H / n
        for i in range(n):
            d.rectangle([0, i * strip, W, i * strip + strip * t], fill=255)
    elif kind == "strips":
        n = 10
        strip = W / n
        for i in range(n):
            d.rectangle([i * strip, 0, i * strip + strip * t, H], fill=255)
    elif kind == "checker":
        n = 8
        cw, ch = W / n, H / n
        for y in range(n):
            for x in range(n):
                if (x + y) % 2 == 0:
                    ww, hh = cw * t, ch * t
                    d.rectangle([x * cw + (cw - ww) / 2,
                                 y * ch + (ch - hh) / 2,
                                 x * cw + (cw + ww) / 2,
                                 y * ch + (ch + hh) / 2], fill=255)
    return m


def morph_frames(a_bg, b_bg, a_shapes, b_shapes, n):
    a_core = [s for s in a_shapes if s.get("name") == "camera_ring"]
    b_core = [s for s in b_shapes if s.get("name") == "camera_ring"]
    frames = []
    for i in range(n):
        t = (i + 1) / n
        eased = RG._ease(t, "smooth")
        f = Image.blend(a_bg, b_bg, eased)
        for s in a_shapes:
            if s.get("name") != "camera_ring":
                RG.draw_shape(f, s, {"in": ("fade", 1 - eased, "out"), "emph": []})
        for s in b_shapes:
            if s.get("name") != "camera_ring":
                RG.draw_shape(f, s, {"in": ("fade", eased, "out"), "emph": []})
        if a_core and b_core:
            ca, cb = a_core[0], b_core[0]
            interp = dict(cb)
            for key in ("x", "y", "w", "h", "rot", "sz"):
                interp[key] = ca.get(key, 0) + (cb.get(key, 0) - ca.get(key, 0)) * t
            layer = Image.new("RGBA", f.size, (0, 0, 0, 0))
            RG.draw_shape(layer, interp, {"in": ("fade", 1.0, "out"), "emph": []})
            f = Image.alpha_composite(f.convert("RGBA"), layer).convert("RGB")
        frames.append(f)
    return frames


def trans_frames(a, b, a_shapes, b_shapes, trans, n, a_bg=None, b_bg=None):
    W, H = a.size
    if trans == "morph":
        return morph_frames(a_bg or a.convert("RGBA"),
                            b_bg or b.convert("RGBA"),
                            a_shapes, b_shapes, n)
    frames = []
    for i in range(n):
        t = RG._ease((i + 1) / n, "smooth")
        f = Image.new("RGB", (W, H), (255, 255, 255))
        if trans == "push_u":
            f.paste(a, (0, int(-H * t)))
            f.paste(b, (0, int(H * (1 - t))))
        elif trans == "push_d":
            f.paste(a, (0, int(H * t)))
            f.paste(b, (0, int(-H * (1 - t))))
        elif trans == "push_l":
            f.paste(a, (int(-W * t), 0))
            f.paste(b, (int(W * (1 - t)), 0))
        elif trans == "push_r":
            f.paste(a, (int(W * t), 0))
            f.paste(b, (int(-W * (1 - t)), 0))
        elif trans in ("wipe_r", "wipe_l", "wipe_d", "wipe_u"):
            f.paste(a, (0, 0))
            m = RG._mask_wipe((W, H), 0, 0, W, H, t, trans[-1])
            f = Image.composite(b, a, m)
        elif trans in ("cover_u", "cover_d", "cover_l", "cover_r"):
            f.paste(a, (0, 0))
            if trans == "cover_u":
                f.paste(b, (0, int(-H * (1 - t))))
            elif trans == "cover_d":
                f.paste(b, (0, int(H * (1 - t))))
            elif trans == "cover_l":
                f.paste(b, (int(-W * (1 - t)), 0))
            else:
                f.paste(b, (int(W * (1 - t)), 0))
        elif trans in ("uncover_u", "uncover_d", "uncover_l", "uncover_r"):
            f.paste(b, (0, 0))
            if trans == "uncover_u":
                f.paste(a, (0, int(-H * t)))
            elif trans == "uncover_d":
                f.paste(a, (0, int(H * t)))
            elif trans == "uncover_l":
                f.paste(a, (int(-W * t), 0))
            else:
                f.paste(a, (int(W * t), 0))
        elif trans == "zoom":
            scaled = b.resize((max(1, int(W * (0.82 + 0.18 * t))),
                               max(1, int(H * (0.82 + 0.18 * t)))),
                              Image.Resampling.LANCZOS)
            f = Image.blend(a, b, t)
            f.paste(scaled, ((W - scaled.width) // 2,
                             (H - scaled.height) // 2))
        elif trans == "cut":
            f = b if t >= 0.5 else a
        elif trans in ("circle", "iris", "wheel", "split_h", "split_v",
                       "doors_h", "doors_v", "blinds", "strips", "checker"):
            f = Image.composite(b, a, _mask((W, H), trans, t))
        else:
            f = Image.blend(a, b, t)
        frames.append(f)
    return frames


def _speed_ms(spd):
    return {"fast": 350, "med": 600, "slow": 1000}.get(spd, 600)


def main():
    if len(sys.argv) < 4:
        print("用法: python render_transition.py <pptx> <plans.json> <out.gif> "
              "[--width 960] [--hold 1.0] [--trans 0.6] [--fps 20]")
        sys.exit(1)
    pptx, plans_json, out = sys.argv[1:4]
    width, hold, trans_s, fps = 960, 1.0, 0.6, 20
    rest = sys.argv[4:]
    i = 0
    while i < len(rest):
        key = rest[i]
        if key == "--width" and i + 1 < len(rest):
            width = int(rest[i + 1]); i += 2
        elif key == "--hold" and i + 1 < len(rest):
            hold = float(rest[i + 1]); i += 2
        elif key == "--trans" and i + 1 < len(rest):
            trans_s = float(rest[i + 1]); i += 2
        elif key == "--fps" and i + 1 < len(rest):
            fps = int(rest[i + 1]); i += 2
        else:
            i += 1
    with open(plans_json, encoding="utf-8") as f:
        plans = json.load(f)
    RG.configure_canvas(pptx, width, fps)
    idxs = sorted(int(k) for k in plans)
    pages, page_bgs, page_shapes = [], [], []
    for idx in idxs:
        img, bg, shapes = page_final(pptx, idx)
        pages.append(img)
        page_bgs.append(bg)
        page_shapes.append(shapes)

    n_hold = max(1, int(hold * fps))
    frames = []
    for k, img in enumerate(pages):
        frames.extend([img] * n_hold)
        if k + 1 < len(pages):
            nxt = plans[str(idxs[k + 1])]
            trans = nxt.get("transition", "fade")
            ms = _speed_ms(nxt.get("transition_spd", "med"))
            if nxt.get("transition_ms"):
                ms = int(nxt["transition_ms"])
            n_trans = max(1, int(ms * fps / 1000))
            frames.extend(trans_frames(
                img, pages[k + 1], page_shapes[k], page_shapes[k + 1],
                trans, n_trans, page_bgs[k], page_bgs[k + 1],
            ))
    frame_ms = 1000 // fps
    frames[0].save(out, save_all=True, append_images=frames[1:],
                   duration=[frame_ms] * len(frames), loop=0, optimize=True)
    print(f"{len(idxs)} 页 {len(frames)} 帧 {len(frames)*frame_ms/1000:.1f}s -> {out}")


if __name__ == "__main__":
    main()
