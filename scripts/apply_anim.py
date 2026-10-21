# -*- coding: utf-8 -*-
"""通用动画应用器 —— 按 plans.json 编排表，把高级动画写入 PPTX。

用法：
    python apply_anim.py <src.pptx> <dst.pptx> <plans.json>

plans.json 结构（页码为 1-based 字符串键）：
    {
      "1": {
        "delete": ["16", "17"],            # 要删除的形状 spid（可选）
        "glow":   ["3", "4"],              # 改为辉光渐变的形状 spid（可选）
        "shadow": {"13": 30, "14": 25},    # spid -> 模糊磅数或完整参数（可选）
        "soft":   {"8": 4},                # 加柔化边缘 spid -> 磅数（可选，慎用易糊）
        "solid":  {"9": ["D4AF37", 80]},   # 改纯色填充（可选）
        "outline":{"20": ["F4C76A", 1.5]}, # 改描边（可选）
        "plan": [                          # 动画编排（必填）
            ["2", "fade", 600, 0],         # [spid, 效果, 时长ms, 延迟ms]
            ["8", "iris", 1100, 100, {"ease": "out"}]
        ]
      }
    }

plan 也支持对象写法：
    {"spid": "8", "effect": "iris", "dur": 1100, "delay": 100, "ease": "out"}

效果枚举见 references/anim_guide.md。
"""
import json
import sys
import anim_engine as AE


def _glow_cfg(value):
    if isinstance(value, dict):
        return value.get("color"), value.get("a1"), value.get("a2", 0)
    return None, None, 0


def _shadow_cfg(value):
    if isinstance(value, dict):
        return {
            "blur_pt": value.get("blur", value.get("blur_pt", 30)),
            "dist_pt": value.get("dist", value.get("dist_pt", 3)),
            "alpha": value.get("alpha", 45),
            "color": value.get("color", "000000"),
            "direction": value.get("direction", 90),
        }
    return {"blur_pt": value, "dist_pt": 3, "alpha": 45,
            "color": "000000", "direction": 90}


def apply(root, idx, cfg):
    if not cfg:
        return
    sm = AE.shape_map(root)
    for spid in cfg.get("delete", []):
        sp = sm.pop(spid, None)
        if sp is not None:
            sp.getparent().remove(sp)
    glow = cfg.get("glow", [])
    glow_items = glow.items() if isinstance(glow, dict) else [(s, None) for s in glow]
    for spid, value in glow_items:
        if spid in sm:
            AE.make_glow(sm[spid], *_glow_cfg(value))
    for spid, value in cfg.get("solid", {}).items():
        if spid in sm:
            color = value[0] if isinstance(value, (list, tuple)) else value
            alpha = value[1] if isinstance(value, (list, tuple)) and len(value) > 1 else 100
            AE.set_solid_fill(sm[spid], color, alpha)
    for spid, value in cfg.get("outline", {}).items():
        if spid not in sm:
            continue
        if isinstance(value, dict):
            AE.set_outline(sm[spid], **value)
        else:
            color = value[0] if isinstance(value, (list, tuple)) else value
            width = value[1] if isinstance(value, (list, tuple)) and len(value) > 1 else None
            AE.set_outline(sm[spid], color=color, width_pt=width)
    for spid, pt in cfg.get("soft", {}).items():
        if spid in sm:
            AE.add_soft_edge(sm[spid], pt)
    for spid, value in cfg.get("shadow", {}).items():
        if spid in sm:
            AE.add_shadow(sm[spid], **_shadow_cfg(value))
    for spid, g in cfg.get("gradient", {}).items():
        # g = [c1, c2, a1, a2, fx, fy] —— 径向渐变（球体/金属质感）
        if spid in sm:
            AE.set_radial_gradient(sm[spid], g[0], g[1],
                                   g[2] if len(g) > 2 else 100,
                                   g[3] if len(g) > 3 else 0,
                                   g[4] if len(g) > 4 else 0.5,
                                   g[5] if len(g) > 5 else 0.5)
    if "plan" in cfg:
        AE.replace_timing(
            root,
            AE.normalize_plan(cfg["plan"]),
            auto_start=bool(cfg.get("auto_start", False)),
        )
    trans = cfg.get("transition")
    if trans:
        if isinstance(trans, dict):
            AE.set_transition(
                root, trans.get("name"), trans.get("spd", "med"),
                trans.get("ms"),
            )
        else:
            AE.set_transition(
                root, trans, cfg.get("transition_spd", "med"),
                cfg.get("transition_ms"),
            )
    for m in cfg.get("morph", []):
        # m = {"spid": "11", "name": "core"} —— 给延续元素命名，供 Morph 跨页变形
        AE.rename_shape(root, m["spid"], m["name"])


def main(src, dst, plans_json):
    plans = json.load(open(plans_json, encoding="utf-8"))
    AE.save_slides(src, dst, lambda root, idx: apply(root, idx, plans.get(str(idx))))
    print("done ->", dst)


if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("用法: python apply_anim.py <src.pptx> <dst.pptx> <plans.json>")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2], sys.argv[3])
