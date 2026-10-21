# -*- coding: utf-8 -*-
"""发布前检查：spid、效果名、转场、shape name 与 Morph 同名核心。

用法：
    python preflight.py <pptx> <plans.json> [--fix-names] [--json]

退出码 1 表示存在会阻止正确播放的问题。
"""
import json
import re
import sys
import zipfile
from collections import Counter

import anim_engine as AE


def _check_slide_structure(root, idx, errors):
    """捕获最容易让 PowerPoint 修复或崩溃的 OOXML 结构错误。"""
    p = f"{{{AE.P}}}"
    direct = [AE.etree.QName(el).localname for el in root]
    if "timing" in direct and "clrMapOvr" in direct:
        if direct.index("timing") < direct.index("clrMapOvr"):
            errors.append(f"slide {idx}: p:timing must follow p:clrMapOvr")
    if "timing" in direct and "transition" in direct:
        if direct.index("timing") < direct.index("transition"):
            errors.append(f"slide {idx}: p:timing must follow p:transition")
    for ctn in root.iter(p + "cTn"):
        if "autoReverse" in ctn.attrib:
            errors.append(f"slide {idx}: cTn uses invalid autoReverse; use autoRev")
    for motion in root.iter(p + "animMotion"):
        if not (motion.get("path") or "").strip():
            errors.append(f"slide {idx}: animMotion needs a path")
        for point in motion:
            if AE.etree.QName(point).localname not in ("from", "to"):
                continue
            for axis in ("x", "y"):
                if (point.get(axis) or "").startswith("#"):
                    errors.append(
                        f"slide {idx}: animMotion {point.tag} {axis} must be numeric")
    trans = root.find(f".//{p}transition")
    if trans is not None:
        for child in trans:
            if AE.etree.QName(child).localname == "morph" and child.tag != f"{{{AE.P14}}}morph":
                errors.append(f"slide {idx}: morph transition must use the p14 namespace")


def check(pptx, plans_json):
    errors, warnings = [], []
    with open(plans_json, encoding="utf-8") as f:
        plans = json.load(f)
    with zipfile.ZipFile(pptx) as z:
        slides = {}
        for name in z.namelist():
            m = re.match(r"ppt/slides/slide(\d+)\.xml$", name)
            if not m:
                continue
            idx = int(m.group(1))
            root = AE.etree.fromstring(z.read(name))
            _check_slide_structure(root, idx, errors)
            shapes = AE.shape_map(root)
            ids = list(AE.iter_shapes(root))
            dup = [sid for sid, n in Counter(sid for sid, _ in ids).items() if n > 1]
            if dup:
                errors.append(f"slide {idx}: duplicate spid {dup}")
            missing_names = [
                sid for sid, el in shapes.items()
                if not (el.find(AE._NV_PATHS[AE.etree.QName(el).localname])
                        .get("name") or "").strip()
            ]
            if missing_names:
                warnings.append(f"slide {idx}: missing cNvPr@name {missing_names}")
            slides[idx] = shapes

        for key, cfg in plans.items():
            try:
                idx = int(key)
            except ValueError:
                errors.append(f"plan key must be a slide number: {key!r}")
                continue
            if idx not in slides:
                errors.append(f"slide {idx}: referenced by plans but not found")
                continue
            sm = slides[idx]
            for i, raw in enumerate(cfg.get("plan", [])):
                try:
                    item = AE.normalize_plan_item(raw)
                except Exception as exc:
                    errors.append(f"slide {idx} item {i}: {exc}")
                    continue
                if item["effect"] not in AE._EFFECT:
                    errors.append(f"slide {idx}: unknown effect {item['effect']!r}")
                if item["spid"] not in sm:
                    errors.append(f"slide {idx}: spid {item['spid']} not found "
                                  f"(effect {item['effect']})")
                if item["dur"] <= 0:
                    errors.append(f"slide {idx}: duration must be > 0")
                if item["delay"] < 0:
                    errors.append(f"slide {idx}: delay must be >= 0")
            trans = cfg.get("transition")
            if isinstance(trans, dict):
                trans = trans.get("name")
            if trans and trans not in AE.TRANSITIONS:
                errors.append(f"slide {idx}: unknown transition {trans!r}")
            for m in cfg.get("morph", []):
                if str(m.get("spid")) not in sm:
                    errors.append(f"slide {idx}: morph spid {m.get('spid')} not found")
                if not (m.get("name") or "").strip():
                    errors.append(f"slide {idx}: morph target needs a name")
    return {"ok": not errors, "errors": errors, "warnings": warnings}


def main():
    if len(sys.argv) < 3:
        print("用法: python preflight.py <pptx> <plans.json> "
              "[--fix-names] [--json]")
        sys.exit(1)
    pptx, plans = sys.argv[1:3]
    if "--fix-names" in sys.argv:
        import fix_names
        n = fix_names.fix(pptx)
        print(f"fixed {n} shape names")
    result = check(pptx, plans)
    if "--json" in sys.argv:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        for msg in result["errors"]:
            print("ERROR", msg)
        for msg in result["warnings"]:
            print("WARN ", msg)
        print("OK" if result["ok"] else "FAILED")
    sys.exit(0 if result["ok"] else 1)


if __name__ == "__main__":
    main()
