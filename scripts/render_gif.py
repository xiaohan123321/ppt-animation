# -*- coding: utf-8 -*-
"""PPT 动画 GIF 预览器。

与旧版相比，这一版按演示文稿真实画幅渲染，并支持：
  - 幻灯片背景、图片、组合形状、图片填充与裁剪
  - 浅色/深色/任意比例画幅（16:9、1:1、竖版）
  - 外层阴影、柔化边缘、径向/线性渐变
  - cinematic / zoom_turn / iris / rise / kenburns / drift_x/y 等新效果

用法：
    python render_gif.py <pptx> <plans.json> <out_dir> [page] [--width 960] [--fps 20]
"""
import io
import json
import math
import os
import posixpath
import sys
import zipfile

from lxml import etree
from PIL import Image, ImageChops, ImageDraw, ImageFilter, ImageFont, ImageOps

import anim_engine as AE

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

EMU_PT = 12700
SLIDE_W_PT = 960.0
SLIDE_H_PT = 540.0
W, H = 960, 540
FPS = 20
FRAME_MS = 1000 // FPS
PX = 1.0

FONT_PATHS = {
    "msyh": r"C:\Windows\Fonts\msyh.ttc",
    "simsun": r"C:\Windows\Fonts\simsun.ttc",
    "simhei": r"C:\Windows\Fonts\simhei.ttf",
    "simkai": r"C:\Windows\Fonts\simkai.ttf",
    "simfang": r"C:\Windows\Fonts\simfang.ttf",
    "times": r"C:\Windows\Fonts\times.ttf",
}

_EMPH_AMP = {
    "pulse": 0.11,
    "breathe": 0.035,
    "kenburns": 0.065,
}
_FADE_EFFECTS = {
    "fade", "zoom", "zoom_soft", "cinematic", "zoom_turn", "spin",
    "float", "rise", "fly_l", "fly_r", "fly_t", "fly_b",
}


def _q(tag):
    return etree.QName(tag).localname


def _xml(z, part):
    return etree.fromstring(z.read(part))


def _rel_part(part, target):
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def _rels(z, part):
    """返回 (id -> part, type suffix -> part)。"""
    base = posixpath.basename(part)
    rels_part = posixpath.join(posixpath.dirname(part), "_rels", base + ".rels")
    out, by_type = {}, {}
    if rels_part not in z.namelist():
        return out, by_type
    root = _xml(z, rels_part)
    for rel in root:
        rid = rel.get("Id")
        target = rel.get("Target")
        if not rid or not target:
            continue
        path = _rel_part(part, target)
        out[rid] = path
        typ = (rel.get("Type") or "").rsplit("/", 1)[-1]
        by_type[typ] = path
    return out, by_type


def _theme_colors(z):
    path = "ppt/theme/theme1.xml"
    if path not in z.namelist():
        return {}
    root = _xml(z, path)
    scheme = root.find(f".//{{{A}}}clrScheme")
    out = {}
    if scheme is None:
        return out
    for slot in scheme:
        srgb = slot.find(f"{{{A}}}srgbClr")
        sysc = slot.find(f"{{{A}}}sysClr")
        if srgb is not None:
            out[_q(slot.tag)] = srgb.get("val")
        elif sysc is not None:
            out[_q(slot.tag)] = sysc.get("lastClr") or "000000"
    for alias, key in (("tx1", "dk1"), ("tx2", "dk2"),
                       ("bg1", "lt1"), ("bg2", "lt2")):
        if key in out:
            out[alias] = out[key]
    return out


def _color(parent, theme):
    if parent is None:
        return None, 100
    srgb = parent.find(f"{{{A}}}srgbClr")
    if srgb is not None:
        color = srgb.get("val")
    else:
        scheme = parent.find(f"{{{A}}}schemeClr")
        if scheme is not None:
            color = theme.get(scheme.get("val"))
        else:
            sysc = parent.find(f"{{{A}}}sysClr")
            color = (sysc.get("lastClr") if sysc is not None else None)
    if not color:
        return None, 100
    clr = parent.find(f"{{{A}}}srgbClr")
    if clr is None:
        clr = parent.find(f"{{{A}}}schemeClr")
    if clr is None:
        clr = parent.find(f"{{{A}}}sysClr")
    alpha = 100
    if clr is not None:
        ae = clr.find(f"{{{A}}}alpha")
        if ae is not None:
            alpha = int(ae.get("val", "100000")) / 1000
    return color.lstrip("#").upper(), alpha


def _image_from_blip(parent, z, rels):
    if parent is None:
        return None
    blip = parent.find(f".//{{{A}}}blip")
    if blip is None:
        return None
    rid = blip.get(f"{{{R}}}embed")
    path = rels.get(rid)
    if not path or path not in z.namelist():
        return None
    try:
        img = Image.open(io.BytesIO(z.read(path))).convert("RGBA")
    except Exception:
        return None
    src = parent.find(f".//{{{A}}}srcRect")
    if src is not None:
        l = int(src.get("l", "0")) / 100000
        t = int(src.get("t", "0")) / 100000
        r = int(src.get("r", "0")) / 100000
        b = int(src.get("b", "0")) / 100000
        iw, ih = img.size
        box = (int(iw * l), int(ih * t),
               int(iw * (1 - r)), int(ih * (1 - b)))
        if box[2] > box[0] and box[3] > box[1]:
            img = img.crop(box)
    return img


def _parse_fill(parent, z, rels, theme):
    if parent is None:
        return None
    if parent.find(f"{{{A}}}noFill") is not None:
        return {"type": "none"}
    solid = parent.find(f"{{{A}}}solidFill")
    if solid is not None:
        color, alpha = _color(solid, theme)
        if color:
            return {"type": "solid", "color": color, "alpha": alpha}
    grad = parent.find(f"{{{A}}}gradFill")
    if grad is not None:
        stops = []
        for gs in grad.findall(f".//{{{A}}}gs"):
            color, alpha = _color(gs, theme)
            if color:
                stops.append({
                    "pos": int(gs.get("pos", "0")) / 100000,
                    "color": color, "alpha": alpha,
                })
        stops.sort(key=lambda s: s["pos"])
        if stops:
            path = grad.find(f".//{{{A}}}path")
            return {
                "type": "gradient", "stops": stops,
                "radial": path is not None,
            }
    blip = parent.find(f"{{{A}}}blipFill")
    if blip is not None:
        img = _image_from_blip(blip, z, rels)
        if img is not None:
            return {"type": "image", "image": img, "alpha": 100}
    return None


def _parse_line(spPr, theme):
    if spPr is None:
        return None
    ln = spPr.find(f"{{{A}}}ln")
    if ln is None or ln.find(f"{{{A}}}noFill") is not None:
        return None
    fill = ln.find(f"{{{A}}}solidFill")
    if fill is None:
        return None
    color, alpha = _color(fill, theme)
    if not color:
        return None
    return {
        "color": color,
        "alpha": alpha,
        "width_pt": int(ln.get("w", "12700")) / EMU_PT,
    }


def _parse_effects(spPr, theme, scale):
    out = {}
    if spPr is None:
        return out
    lst = spPr.find(f"{{{A}}}effectLst")
    if lst is None:
        return out
    sh = lst.find(f"{{{A}}}outerShdw")
    if sh is not None:
        color, alpha = _color(sh, theme)
        if color:
            out["shadow"] = {
                "color": color,
                "alpha": alpha,
                "blur_pt": int(sh.get("blurRad", "0")) / EMU_PT,
                "dist_pt": int(sh.get("dist", "0")) / EMU_PT,
                "dir": int(sh.get("dir", "5400000")) / 60000,
                "scale": scale,
            }
    se = lst.find(f"{{{A}}}softEdge")
    if se is not None:
        out["soft_pt"] = int(se.get("rad", "0")) / EMU_PT
    return out


def _parse_text(sp, theme):
    body = sp.find(f".//{{{P}}}txBody")
    if body is None:
        return "", [], {}
    anchor = "ctr"
    body_pr = body.find(f"{{{A}}}bodyPr")
    if body_pr is not None:
        anchor = body_pr.get("anchor", "ctr")
    paragraphs, text_lines = [], []
    align, sz, color, font, bold, italic = (
        "ctr", 14, theme.get("tx1", "000000"), "", False, False)
    rgb = None
    for p in body.findall(f"{{{A}}}p"):
        ppr = p.find(f"{{{A}}}pPr")
        if ppr is not None and ppr.get("algn"):
            align = ppr.get("algn")
        parts = []
        for child in p:
            if child.tag == f"{{{A}}}r":
                rpr = child.find(f"{{{A}}}rPr")
                if rpr is not None:
                    if rpr.get("sz"):
                        sz = int(rpr.get("sz")) / 100
                    if rpr.get("b") in ("1", "true"):
                        bold = True
                    if rpr.get("i") in ("1", "true"):
                        italic = True
                    for key in ("latin", "ea"):
                        face = rpr.find(f"{{{A}}}{key}")
                        if face is not None and face.get("typeface"):
                            font = face.get("typeface")
                    sf = rpr.find(f"{{{A}}}solidFill")
                    if sf is not None:
                        c, _ = _color(sf, theme)
                        if c:
                            color = c
                            rgb = c
                t = child.find(f"{{{A}}}t")
                if t is not None and t.text:
                    parts.append(t.text)
            elif child.tag == f"{{{A}}}br":
                parts.append("\n")
        line = "".join(parts)
        if ppr is not None and line:
            if ppr.find(f"{{{A}}}buChar") is not None:
                line = "- " + line
            elif ppr.find(f"{{{A}}}buAutoNum") is not None:
                line = "- " + line
        text_lines.append(line)
        paragraphs.append(line)
    return "\n".join(text_lines), paragraphs, {
        "anchor": anchor, "align": align, "sz": sz,
        "color": color, "font": font, "bold": bold, "italic": italic,
    }


def _xfrm(el):
    xf = el.find(f".//{{{A}}}xfrm")
    if xf is None:
        return None
    off, ext = xf.find(f"{{{A}}}off"), xf.find(f"{{{A}}}ext")
    if off is None or ext is None:
        return None
    return {
        "x": int(off.get("x", "0")) / EMU_PT,
        "y": int(off.get("y", "0")) / EMU_PT,
        "w": int(ext.get("cx", "0")) / EMU_PT,
        "h": int(ext.get("cy", "0")) / EMU_PT,
        "rot": int(xf.get("rot", "0")) / 60000,
        "flipH": xf.get("flipH") in ("1", "true"),
        "flipV": xf.get("flipV") in ("1", "true"),
    }


def _transform_box(box, tr):
    box["x"] = tr["ox"] + box["x"] * tr["sx"]
    box["y"] = tr["oy"] + box["y"] * tr["sy"]
    box["w"] *= tr["sx"]
    box["h"] *= tr["sy"]
    return box


def _parse_shape(el, kind, tr, group_ids, z, rels, theme):
    nv_path = AE._NV_PATHS[kind]
    nv = el.find(nv_path)
    if nv is None:
        return None
    box = _xfrm(el)
    if box is None:
        return None
    box = _transform_box(box, tr)
    sp_pr = el.find(f"{{{P}}}spPr")
    if sp_pr is None:
        sp_pr = el.find(f".//{{{A}}}spPr")
    fill = _parse_fill(sp_pr, z, rels, theme)
    if kind == "pic":
        img = _image_from_blip(el, z, rels)
        if img is not None:
            fill = {"type": "image", "image": img, "alpha": 100}
    geom_el = sp_pr.find(f".//{{{A}}}prstGeom") if sp_pr is not None else None
    geom = geom_el.get("prst") if geom_el is not None else "rect"
    adj = []
    if geom_el is not None:
        for gd in geom_el.findall(f".//{{{A}}}gd"):
            fmla = gd.get("fmla", "")
            if fmla.startswith("val "):
                try:
                    v = int(fmla.split()[1])
                    adj.append(v / 60000 if abs(v) > 3600 else v)
                except ValueError:
                    pass
    text, paragraphs, typography = _parse_text(el, theme)
    effects = _parse_effects(sp_pr, theme, PX)
    return {
        "id": nv.get("id"),
        "name": nv.get("name") or "",
        "kind": kind,
        "group_ids": list(group_ids),
        "x": box["x"], "y": box["y"], "w": box["w"], "h": box["h"],
        "rot": box["rot"], "flipH": box["flipH"], "flipV": box["flipV"],
        "fill": fill, "line": _parse_line(sp_pr, theme),
        "geom": geom, "adj": adj, "text": text, "paragraphs": paragraphs,
        **typography, **effects,
    }


def _walk_shapes(container, tr, group_ids, z, rels, theme, out):
    for child in container:
        kind = _q(child.tag)
        if kind == "grpSp":
            nv = child.find(f"{{{P}}}nvGrpSpPr/{{{P}}}cNvPr")
            gid = nv.get("id") if nv is not None else None
            gx = child.find(f"{{{P}}}grpSpPr/{{{A}}}xfrm")
            if gx is not None and gid:
                off, ext = gx.find(f"{{{A}}}off"), gx.find(f"{{{A}}}ext")
                choff, chext = gx.find(f"{{{A}}}chOff"), gx.find(f"{{{A}}}chExt")
                if off is not None and ext is not None:
                    ox = int(off.get("x", "0")) / EMU_PT
                    oy = int(off.get("y", "0")) / EMU_PT
                    ew = int(ext.get("cx", "1")) / EMU_PT
                    eh = int(ext.get("cy", "1")) / EMU_PT
                    cx = int(choff.get("x", "0")) / EMU_PT if choff is not None else 0
                    cy = int(choff.get("y", "0")) / EMU_PT if choff is not None else 0
                    cw = int(chext.get("cx", "1")) / EMU_PT if chext is not None else ew
                    ch = int(chext.get("cy", "1")) / EMU_PT if chext is not None else eh
                    sx = ew / cw if cw else 1
                    sy = eh / ch if ch else 1
                    ntr = {
                        "sx": tr["sx"] * sx, "sy": tr["sy"] * sy,
                        "ox": tr["ox"] + (ox - cx * sx) * tr["sx"],
                        "oy": tr["oy"] + (oy - cy * sy) * tr["sy"],
                    }
                    gbox = _xfrm(child)
                    if gbox is not None:
                        gbox = _transform_box(gbox, tr)
                        out.append({
                            "id": gid, "name": nv.get("name") or "",
                            "kind": "group", "group_ids": list(group_ids),
                            **{k: gbox[k] for k in ("x", "y", "w", "h", "rot")},
                            "fill": None, "line": None, "geom": "rect", "adj": [],
                            "text": "", "paragraphs": [], "anchor": "ctr",
                            "align": "ctr", "sz": 14, "color": "FFFFFF",
                            "font": "", "bold": False, "italic": False,
                        })
                        _walk_shapes(child, ntr, group_ids + [gid], z, rels,
                                     theme, out)
        elif kind in AE._NV_PATHS:
            shape = _parse_shape(child, kind, tr, group_ids, z, rels, theme)
            if shape is not None:
                out.append(shape)


def configure_canvas(pptx, max_width=960, fps=20):
    global W, H, PX, FPS, FRAME_MS, SLIDE_W_PT, SLIDE_H_PT
    with zipfile.ZipFile(pptx) as z:
        root = _xml(z, "ppt/presentation.xml")
        sz = root.find(f"{{{P}}}sldSz")
        if sz is not None:
            SLIDE_W_PT = int(sz.get("cx")) / EMU_PT
            SLIDE_H_PT = int(sz.get("cy")) / EMU_PT
    max_width = max(240, int(max_width))
    W = max_width
    H = max(1, int(round(max_width * SLIDE_H_PT / SLIDE_W_PT)))
    PX = W / SLIDE_W_PT
    FPS = max(1, int(fps))
    FRAME_MS = 1000 // FPS
    return W, H


def parse_slide(z, idx):
    part = f"ppt/slides/slide{idx}.xml"
    root = _xml(z, part)
    rels, _ = _rels(z, part)
    theme = _theme_colors(z)
    sp_tree = root.find(f".//{{{P}}}spTree")
    shapes = []
    if sp_tree is not None:
        _walk_shapes(sp_tree, {"sx": 1, "sy": 1, "ox": 0, "oy": 0},
                     [], z, rels, theme, shapes)
    return shapes


def _background_for_part(z, part):
    if part not in z.namelist():
        return None
    root = _xml(z, part)
    bg = root.find(f".//{{{P}}}bg")
    if bg is None:
        return None
    bg_pr = bg.find(f"{{{P}}}bgPr")
    if bg_pr is None:
        return None
    rels, _ = _rels(z, part)
    return _parse_fill(bg_pr, z, rels, _theme_colors(z))


def slide_background(z, idx):
    slide = f"ppt/slides/slide{idx}.xml"
    fill = _background_for_part(z, slide)
    if fill is not None:
        return fill
    _, types = _rels(z, slide)
    layout = types.get("slideLayout")
    if layout:
        fill = _background_for_part(z, layout)
        if fill is not None:
            return fill
        _, ltypes = _rels(z, layout)
        master = ltypes.get("slideMaster")
        if master:
            fill = _background_for_part(z, master)
            if fill is not None:
                return fill
    return {"type": "solid", "color": "FFFFFF", "alpha": 100}


def _hex_rgb(color):
    color = (color or "FFFFFF").lstrip("#")
    return tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))


def _fill_surface(fill, size):
    w, h = size
    if not fill or fill["type"] == "none":
        return Image.new("RGBA", size, (0, 0, 0, 0))
    if fill["type"] == "solid":
        c = _hex_rgb(fill["color"]) + (int(255 * fill["alpha"] / 100),)
        return Image.new("RGBA", size, c)
    if fill["type"] == "image":
        img = ImageOps.fit(fill["image"], size, method=Image.Resampling.LANCZOS)
        return img.convert("RGBA")
    if fill["type"] == "gradient":
        stops = fill["stops"]
        a, b = stops[0], stops[-1]
        c1, c2 = _hex_rgb(a["color"]), _hex_rgb(b["color"])
        surf = Image.new("RGBA", size)
        px = surf.load()
        if fill.get("radial"):
            cx, cy = w / 2, h / 2
            maxd = math.hypot(cx, cy) or 1
            for y in range(h):
                for x in range(w):
                    t = min(1, math.hypot(x - cx, y - cy) / maxd)
                    col = tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3))
                    alpha = int(255 * (a["alpha"] + (b["alpha"] - a["alpha"]) * t) / 100)
                    px[x, y] = col + (alpha,)
        else:
            for y in range(h):
                t = y / max(1, h - 1)
                col = tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3))
                alpha = int(255 * (a["alpha"] + (b["alpha"] - a["alpha"]) * t) / 100)
                for x in range(w):
                    px[x, y] = col + (alpha,)
        return surf
    return Image.new("RGBA", size, (255, 255, 255, 255))


def background_image(z, idx):
    return _fill_surface(slide_background(z, idx), (W, H)).convert("RGBA")


def _ease(p, kind="out"):
    p = max(0.0, min(1.0, p))
    if kind in ("in", "ease-in", "easeIn"):
        return p ** 3
    if kind in ("smooth", "ease-in-out", "easeInOut"):
        return p * p * (3 - 2 * p)
    if kind in ("none", "linear"):
        return p
    return 1 - (1 - p) ** 3


def shape_state(plan, t):
    out = {}
    items = AE.normalize_plan(plan) if plan else []
    for item in items:
        d = out.setdefault(
            item["spid"], {"in": None, "out": None, "emph": [], "visible": True}
        )
        delay, dur = item["delay"], item["dur"]
        if item["effect"] in AE.EXIT_EFFECTS:
            if t >= delay + dur:
                d["visible"] = False
            elif t >= delay:
                prog = (t - delay) / dur if dur else 1.0
                d["out"] = (item["effect"], _ease(prog, "in"), "in")
        elif item["effect"] in AE.EMPH_EFFECTS:
            if t >= delay:
                prog = ((t - delay) % dur) / dur if dur else 1.0
                d["emph"].append((item["effect"], prog))
        elif t >= delay:
            prog = 1.0 if t >= delay + dur else (t - delay) / dur if dur else 1.0
            d["in"] = (item["effect"], _ease(prog, item.get("ease", "out")),
                       item.get("ease", "out"))
    return out


def _state_for(shape, st):
    if shape["id"] in st:
        return st[shape["id"]]
    for gid in reversed(shape.get("group_ids", [])):
        if gid in st:
            return st[gid]
    # 未参与动画的元素（背景、圆窗图片、金环等）必须始终显示，
    # 否则整层被跳过 —— 会出现白底、图片消失只剩边框。
    return {"in": ("fade", 1.0, "out"), "emph": [], "visible": True}


def _mask_wheel(size, x, y, w, h, prog):
    m = Image.new("L", size, 0)
    d = ImageDraw.Draw(m)
    cx, cy = x + w / 2, y + h / 2
    r = math.hypot(w, h)
    d.pieslice([cx - r, cy - r, cx + r, cy + r],
               -90, -90 + 360 * prog, fill=255)
    return m


def _mask_wipe(size, x, y, w, h, prog, direction):
    m = Image.new("L", size, 0)
    d = ImageDraw.Draw(m)
    if direction == "r":
        d.rectangle([x, y, x + w * prog, y + h], fill=255)
    elif direction == "l":
        d.rectangle([x + w * (1 - prog), y, x + w, y + h], fill=255)
    elif direction == "d":
        d.rectangle([x, y, x + w, y + h * prog], fill=255)
    else:
        d.rectangle([x, y + h * (1 - prog), x + w, y + h], fill=255)
    return m


def _mask_center(size, x, y, w, h, prog, kind):
    m = Image.new("L", size, 0)
    d = ImageDraw.Draw(m)
    cx, cy = x + w / 2, y + h / 2
    if kind in ("circle", "iris"):
        r = math.hypot(w, h) * prog
        d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=255)
    elif kind == "box":
        d.rectangle([cx - w * prog / 2, cy - h * prog / 2,
                     cx + w * prog / 2, cy + h * prog / 2], fill=255)
    elif kind == "diamond":
        dx, dy = w * prog / 2, h * prog / 2
        d.polygon([(cx, cy - dy), (cx + dx, cy),
                   (cx, cy + dy), (cx - dx, cy)], fill=255)
    return m


def _mask_split(size, x, y, w, h, prog):
    m = Image.new("L", size, 0)
    d = ImageDraw.Draw(m)
    cx, cy = x + w / 2, y + h / 2
    half = h / 2 * prog
    d.rectangle([x, cy - half, x + w, cy], fill=255)
    d.rectangle([x, cy, x + w, cy + half], fill=255)
    return m


def _mask_blinds(size, x, y, w, h, prog):
    m = Image.new("L", size, 0)
    d = ImageDraw.Draw(m)
    n = 8
    strip = h / n
    for i in range(n):
        sy = y + i * strip
        d.rectangle([x, sy, x + w, sy + strip * prog], fill=255)
    return m


def _font_for(name, text):
    low = (name or "").lower()
    if "宋" in name or "song" in low or "serif" in low:
        path = FONT_PATHS["simsun"]
    elif "楷" in name or "kai" in low:
        path = FONT_PATHS["simkai"]
    elif "黑" in name or "hei" in low:
        path = FONT_PATHS["simhei"]
    elif "times" in low:
        path = FONT_PATHS["times"]
    else:
        path = FONT_PATHS["msyh"]
    return path if os.path.exists(path) else FONT_PATHS["msyh"]


def _wrap_line(draw, line, font, max_width):
    if not line:
        return [""]
    if draw.textlength(line, font=font) <= max_width:
        return [line]
    out, cur = [], ""
    for ch in line:
        trial = cur + ch
        if cur and draw.textlength(trial, font=font) > max_width:
            out.append(cur)
            cur = ch
        else:
            cur = trial
    if cur:
        out.append(cur)
    return out


def _draw_text(layer, s, box, alpha):
    text = (s.get("text") or "").strip()
    if not text:
        return
    d = ImageDraw.Draw(layer)
    size = max(6, int(s.get("sz", 14) * PX))
    path = _font_for(s.get("font", ""), text)
    font = ImageFont.truetype(path, size)
    x, y, w, h = box
    wrapped = []
    for line in text.split("\n"):
        wrapped.extend(_wrap_line(d, line, font, max(8, w - 8)))
    line_h = size * 1.22
    while len(wrapped) * line_h > max(8, h - 6) and size > 6:
        size -= 1
        font = ImageFont.truetype(path, size)
        line_h = size * 1.22
        wrapped = []
        for line in text.split("\n"):
            wrapped.extend(_wrap_line(d, line, font, max(8, w - 8)))
    total = line_h * len(wrapped)
    if s.get("anchor") == "t":
        ty = y + 4
    elif s.get("anchor") == "b":
        ty = y + h - 4 - total
    else:
        ty = y + (h - total) / 2
    color = _hex_rgb(s.get("color", "FFFFFF")) + (int(255 * alpha),)
    for line in wrapped:
        tw = d.textlength(line, font=font)
        if s.get("align") == "ctr":
            tx = x + (w - tw) / 2
        elif s.get("align") == "r":
            tx = x + w - 6 - tw
        else:
            tx = x + 6
        d.text((tx, ty), line, font=font, fill=color)
        ty += line_h


def _draw_geometry(draw, s, box, fill, line):
    x, y, w, h = box
    geom = s.get("geom", "rect")
    common = {"fill": fill, "outline": line}
    if geom == "ellipse":
        draw.ellipse([x, y, x + w, y + h], **common)
    elif geom == "roundRect":
        draw.rounded_rectangle([x, y, x + w, y + h],
                               radius=min(w, h) * 0.12, **common)
    elif geom == "pie":
        adj = s.get("adj") or []
        a0 = adj[0] if len(adj) > 0 else 0
        a1 = adj[1] if len(adj) > 1 else 90
        draw.pieslice([x, y, x + w, y + h], -a1, -a0, **common)
    elif geom == "arc":
        # OOXML 的 arc 用 adj1/adj2（度，顺时针，0° 在 3 点钟方向），
        # PIL 的 arc 角度约定相同，可直接映射。缺这个分支会退化成矩形。
        adj = s.get("adj") or []
        a0 = adj[0] if len(adj) > 0 else 0
        a1 = adj[1] if len(adj) > 1 else 360
        if a1 <= a0:
            a1 += 360
        if fill:
            draw.pieslice([x, y, x + w, y + h], start=a0, end=a1, fill=fill)
        if line:
            draw.arc([x, y, x + w, y + h], start=a0, end=a1,
                     fill=line, width=max(1, int(s.get("line_w", 2))))
    elif geom == "diamond":
        draw.polygon([(x + w / 2, y), (x + w, y + h / 2),
                      (x + w / 2, y + h), (x, y + h / 2)], **common)
    elif geom == "triangle":
        draw.polygon([(x + w / 2, y), (x + w, y + h),
                      (x, y + h)], **common)
    elif geom in ("hexagon", "homePlate"):
        draw.polygon([(x + w * .25, y), (x + w * .75, y),
                      (x + w, y + h / 2), (x + w * .75, y + h),
                      (x + w * .25, y + h), (x, y + h / 2)], **common)
    else:
        draw.rectangle([x, y, x + w, y + h], **common)


def _fill_layer(layer, s, box):
    fill = s.get("fill")
    if not fill:
        return
    x, y, w, h = [int(round(v)) for v in box]
    w, h = max(1, w), max(1, h)
    d = ImageDraw.Draw(layer)
    if fill["type"] == "solid":
        col = _hex_rgb(fill["color"]) + (int(255 * fill["alpha"] / 100),)
        _draw_geometry(d, s, box, col, None)
    elif fill["type"] == "image":
        img = ImageOps.fit(fill["image"], (w, h),
                           method=Image.Resampling.LANCZOS)
        mask = Image.new("L", (w, h), 0)
        md = ImageDraw.Draw(mask)
        local = {"geom": s.get("geom", "rect"), "adj": s.get("adj", [])}
        _draw_geometry(md, local, (0, 0, w, h), 255, None)
        layer.paste(img, (x, y), mask)
    elif fill["type"] == "gradient":
        stops = fill["stops"]
        c1, c2 = _hex_rgb(stops[0]["color"]), _hex_rgb(stops[-1]["color"])
        n = max(2, min(40, int(max(w, h) / 4)))
        for i in range(n, 0, -1):
            t = i / n
            col = tuple(int(c2[k] + (c1[k] - c2[k]) * (1 - t)) for k in range(3))
            a1, a2 = stops[0]["alpha"], stops[-1]["alpha"]
            alpha = int(255 * (a2 + (a1 - a2) * (1 - t)) / 100)
            rw, rh = w * t, h * t
            rx, ry = x + (w - rw) / 2, y + (h - rh) / 2
            if s.get("geom") == "ellipse":
                d.ellipse([rx, ry, rx + rw, ry + rh], fill=col + (alpha,))
            else:
                d.rectangle([rx, ry, rx + rw, ry + rh], fill=col + (alpha,))


def _apply_shadow(base, layer, s):
    sh = s.get("shadow")
    if not sh:
        return
    alpha = layer.split()[3]
    alpha = alpha.point(lambda v: int(v * sh["alpha"] / 100))
    shadow = Image.new("RGBA", layer.size, _hex_rgb(sh["color"]) + (0,))
    shadow.putalpha(alpha)
    blur = sh["blur_pt"] * PX
    if blur > 0:
        shadow = shadow.filter(ImageFilter.GaussianBlur(blur))
    rad = math.radians(sh["dir"])
    dx = math.cos(rad) * sh["dist_pt"] * PX
    dy = math.sin(rad) * sh["dist_pt"] * PX
    base.alpha_composite(shadow, (int(dx), int(dy)))


def draw_shape(base, s, state):
    if s.get("kind") == "group":
        return
    if state is None or state.get("visible") is False:
        return
    in_state = state.get("in")
    out_state = state.get("out")
    if in_state is None and out_state is None:
        effect, prog, ease = "static", 1.0, "out"
    elif in_state is not None:
        effect, prog, ease = in_state
    else:
        effect, prog, ease = "static", 1.0, "out"
    x, y, w, h = s["x"] * PX, s["y"] * PX, s["w"] * PX, s["h"] * PX
    alpha = prog if effect in _FADE_EFFECTS else 1.0
    scale = 1.0
    rot = s.get("rot", 0)
    dx = dy = 0.0

    if effect == "zoom":
        scale = prog
    elif effect in ("zoom_soft", "cinematic"):
        start = 0.86 if effect == "cinematic" else 0.4
        scale = start + (1 - start) * prog
    elif effect == "zoom_turn":
        scale = 0.72 + 0.28 * prog
        rot += -8 * (1 - prog)
    elif effect == "spin":
        scale = prog
        rot += 360 * (1 - prog)
    elif effect in ("float", "rise"):
        dy += (1 - prog) * H * 0.06
    elif effect in ("fly_l", "fly_r", "fly_t", "fly_b"):
        if effect == "fly_l":
            dx -= (1 - prog) * W
        elif effect == "fly_r":
            dx += (1 - prog) * W
        elif effect == "fly_t":
            dy -= (1 - prog) * H
        else:
            dy += (1 - prog) * H

    for eeff, eprog in state.get("emph", []):
        wave = math.sin(eprog * 2 * math.pi)
        if eeff in _EMPH_AMP:
            scale *= 1 + _EMPH_AMP[eeff] * wave
        elif eeff in ("sway", "sway6", "sway10", "sway15", "sway20"):
            amp = {"sway": 18, "sway6": 6, "sway10": 10,
                   "sway15": 15, "sway20": 20}[eeff]
            rot += amp * wave
        elif eeff == "rotate":
            rot += eprog * 360
        elif eeff == "drift_x":
            dx += wave * W * 0.012
        elif eeff == "drift_y":
            dy += wave * H * 0.012

    if out_state is not None:
        oeffect, oprog, _ = out_state
        alpha *= max(0.0, 1.0 - oprog)
        if oeffect == "exit_fly_l":
            dx -= oprog * W
        elif oeffect == "exit_fly_r":
            dx += oprog * W
        elif oeffect == "exit_fly_t":
            dy -= oprog * H
        elif oeffect == "exit_fly_b":
            dy += oprog * H
        elif oeffect == "exit_zoom":
            scale *= max(0.0, 1.0 - oprog)
        elif oeffect == "exit_turn":
            scale *= max(0.0, 1.0 - oprog)
            rot += 180 * oprog

    cx, cy = x + w / 2, y + h / 2
    if scale != 1.0:
        x, y = cx - w * scale / 2, cy - h * scale / 2
        w, h = w * scale, h * scale
    x += dx
    y += dy

    layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    _fill_layer(layer, s, (x, y, w, h))
    if s.get("line"):
        ln = s["line"]
        col = _hex_rgb(ln["color"]) + (int(255 * ln["alpha"] / 100),)
        d = ImageDraw.Draw(layer)
        width = max(1, int(round(ln["width_pt"] * PX)))
        _draw_geometry(d, s, (x, y, w, h), None, col)
        if width > 1:
            for k in range(1, width):
                _draw_geometry(d, s, (x - k, y - k, w + 2 * k, h + 2 * k),
                               None, col)
    _draw_text(layer, s, (x, y, w, h), alpha)

    if rot:
        layer = layer.rotate(-rot, resample=Image.Resampling.BICUBIC,
                             center=(x + w / 2, y + h / 2))
    if alpha < 1:
        a = layer.split()[3].point(lambda v: int(v * alpha))
        layer.putalpha(a)
    if s.get("soft_pt"):
        a = layer.split()[3].filter(
            ImageFilter.GaussianBlur(s["soft_pt"] * PX))
        layer.putalpha(a)

    mask = None
    if effect == "wheel":
        mask = _mask_wheel((W, H), x, y, w, h, prog)
    elif effect in ("wipe_r", "wipe_l", "wipe_d", "wipe_u"):
        mask = _mask_wipe((W, H), x, y, w, h, prog, effect[-1])
    elif effect in ("circle", "iris", "box", "diamond"):
        mask = _mask_center((W, H), x, y, w, h, prog, effect)
    elif effect == "split":
        mask = _mask_split((W, H), x, y, w, h, prog)
    elif effect == "blinds":
        mask = _mask_blinds((W, H), x, y, w, h, prog)
    if mask is not None:
        a = layer.split()[3]
        a = Image.composite(a, Image.new("L", (W, H), 0), mask)
        layer.putalpha(a)

    _apply_shadow(base, layer, s)
    base.alpha_composite(layer)


def total_duration(plan):
    items = AE.normalize_plan(plan) if plan else []
    ends = [i["delay"] + i["dur"] for i in items
            if i["effect"] not in AE.EMPH_EFFECTS]
    has_emph = any(i["effect"] in AE.EMPH_EFFECTS for i in items)
    tail = 2200 if has_emph else 600
    return max(ends) + tail if ends else 1000


def render_page_gif(pptx, idx, plan, path, width=None, fps=None):
    if width or fps:
        configure_canvas(pptx, width or W, fps or FPS)
    with zipfile.ZipFile(pptx) as z:
        shapes = parse_slide(z, idx)
        bg = background_image(z, idx)
        total = total_duration(plan)
        nframes = max(1, int(math.ceil(total / FRAME_MS)))
        frames = []
        for fi in range(nframes):
            t = fi * FRAME_MS
            st = shape_state(plan, t)
            base = bg.copy()
            for s in shapes:
                state = _state_for(s, st)
                if state and state.get("in") is not None:
                    draw_shape(base, s, state)
            frames.append(base.convert("RGB"))
    frames[0].save(path, save_all=True, append_images=frames[1:],
                   duration=[FRAME_MS] * len(frames), loop=0, optimize=False,
                   disposal=1)
    return nframes, total


def main():
    if len(sys.argv) < 4:
        print("用法: python render_gif.py <pptx> <plans.json> <out_dir> "
              "[page] [--width 960] [--fps 20]")
        sys.exit(1)
    pptx, plans_json, out = sys.argv[1:4]
    only = None
    width, fps = 960, 20
    rest = sys.argv[4:]
    i = 0
    while i < len(rest):
        if rest[i] == "--width" and i + 1 < len(rest):
            width = int(rest[i + 1]); i += 2
        elif rest[i] == "--fps" and i + 1 < len(rest):
            fps = int(rest[i + 1]); i += 2
        else:
            only = int(rest[i]); i += 1
    with open(plans_json, encoding="utf-8") as f:
        plans = json.load(f)
    configure_canvas(pptx, width, fps)
    os.makedirs(out, exist_ok=True)
    for k, cfg in plans.items():
        idx = int(k)
        if only and idx != only:
            continue
        p = os.path.join(out, f"anim_p{idx:02d}.gif")
        nf, tot = render_page_gif(pptx, idx, cfg.get("plan", []), p)
        print(f"p{idx:02d} {nf}帧 {tot/1000:.1f}s -> {p}")


if __name__ == "__main__":
    main()
