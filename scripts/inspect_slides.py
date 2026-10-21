# -*- coding: utf-8 -*-
"""导出每页可动画对象的 spid、名称、位置、尺寸、文本与视觉角色。

用法：
    python inspect_slides.py <pptx> [out.json] [--compact]

建议先跑本脚本，再写 plans.json。背景/圆环/标题/正文/图片会附带 role 提示，
但最终 spid 仍以 cNvPr@id 为准。
"""
import json
import re
import sys
import zipfile

import render_gif as RG


def _role(shape, slide_w, slide_h):
    w, h = shape.get("w", 0), shape.get("h", 0)
    if w >= slide_w * 0.9 and h >= slide_h * 0.9:
        return "background"
    if shape.get("kind") == "pic":
        return "image"
    if shape.get("geom") in ("ellipse", "donut", "ring"):
        if w >= slide_w * 0.18:
            return "ring_or_disc"
    if shape.get("text"):
        if shape.get("sz", 0) >= 30:
            return "title"
        return "text"
    if shape.get("fill") or shape.get("line"):
        return "decoration"
    return "shape"


def inspect_deck(path):
    with zipfile.ZipFile(path) as z:
        RG.configure_canvas(path, 960, 20)
        slide_parts = sorted(
            (int(m.group(1)), name)
            for name in z.namelist()
            for m in [re.match(r"ppt/slides/slide(\d+)\.xml$", name)]
            if m
        )
        slides = []
        for idx, part in slide_parts:
            shapes = RG.parse_slide(z, idx)
            root = RG._xml(z, part)
            trans = root.find(f".//{{{RG.P}}}transition")
            trans_name = None
            if trans is not None and len(trans):
                trans_name = RG._q(trans[0].tag)
            timing = root.find(f".//{{{RG.P}}}timing")
            for s in shapes:
                s["role"] = _role(s, RG.SLIDE_W_PT, RG.SLIDE_H_PT)
                fill = s.get("fill")
                if isinstance(fill, dict) and fill.get("type") == "image":
                    img = fill.get("image")
                    s["fill"]["image_size"] = list(img.size) if img else None
                for key in ("fill", "line", "shadow"):
                    if isinstance(s.get(key), dict):
                        s[key] = {k: v for k, v in s[key].items()
                                  if k != "image"}
            slides.append({
                "slide": idx,
                "transition": trans_name,
                "has_timing": timing is not None,
                "shapes": shapes,
            })
    return {
        "file": path,
        "slide_size_pt": [round(RG.SLIDE_W_PT, 2), round(RG.SLIDE_H_PT, 2)],
        "aspect": round(RG.SLIDE_W_PT / RG.SLIDE_H_PT, 4),
        "slides": slides,
    }


def main():
    if len(sys.argv) < 2:
        print("用法: python inspect_slides.py <pptx> [out.json] [--compact]")
        sys.exit(1)
    path = sys.argv[1]
    compact = "--compact" in sys.argv
    out = next((a for a in sys.argv[2:] if not a.startswith("--")), None)
    data = inspect_deck(path)
    text = json.dumps(data, ensure_ascii=False, indent=None if compact else 2)
    if out:
        with open(out, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"{len(data['slides'])} 页 -> {out}")
    else:
        print(text)


if __name__ == "__main__":
    main()
