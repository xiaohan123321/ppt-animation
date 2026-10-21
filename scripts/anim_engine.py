# -*- coding: utf-8 -*-
"""高级动画引擎 —— 直接重写 PPTX 的 OOXML 动画时间轴与形状特效。

通用 PPT 工具只支持基础动画、无时长/延迟/柔化/阴影，且 zoom 动画缺真缩放节点。
本模块用 lxml 直接操控底层 XML：

  入场效果：fade / wheel / iris / wipe_* / zoom / zoom_soft / zoom_turn /
           cinematic / float / rise / fly_* / box / diamond / circle /
           blinds / split / spin
  强调效果：pulse / breathe / kenburns / sway / rotate / drift_x / drift_y
  特效：softEdge 柔化边缘 / outerShdw 阴影 / gradFill 径向渐变辉光

命名空间与 presetID 依据 ECMA-376 / PowerPoint 标准。
"""
import re
import zipfile
from lxml import etree

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def load_slide(zipf, idx):
    data = zipf.read(f"ppt/slides/slide{idx}.xml")
    return etree.fromstring(data), data


_NV_PATHS = {
    "sp": f"{{{P}}}nvSpPr/{{{P}}}cNvPr",
    "pic": f"{{{P}}}nvPicPr/{{{P}}}cNvPr",
    "graphicFrame": f"{{{P}}}nvGraphicFramePr/{{{P}}}cNvPr",
    "cxnSp": f"{{{P}}}nvCxnSpPr/{{{P}}}cNvPr",
    "grpSp": f"{{{P}}}nvGrpSpPr/{{{P}}}cNvPr",
}


def iter_shapes(root):
    """按文档顺序遍历可动画对象（含图片、连接线、表格/图表框与组合）。"""
    for child in root.iter():
        local = etree.QName(child).localname
        path = _NV_PATHS.get(local)
        if path is None:
            continue
        nv = child.find(path)
        if nv is not None and nv.get("id"):
            yield nv.get("id"), child


def shape_map(root):
    """spid -> shape element；图片不再被遗漏。"""
    return dict(iter_shapes(root))


# ---------- 形状特效 ----------

def _effect_lst(sp):
    spPr = sp.find(f".//{{{P}}}spPr")
    lst = spPr.find(f"{{{A}}}effectLst")
    if lst is None:
        lst = etree.SubElement(spPr, f"{{{A}}}effectLst")
    return lst


def _set_fill_el(spPr, fill):
    for tag in ("solidFill", "gradFill", "noFill", "blipFill", "pattFill", "grpFill"):
        for old in spPr.findall(f"{{{A}}}{tag}"):
            spPr.remove(old)
    anchor = None
    for tag in ("ln", "effectLst", "effectDag", "scene3d", "sp3d"):
        anchor = spPr.find(f"{{{A}}}{tag}")
        if anchor is not None:
            break
    if anchor is not None:
        anchor.addprevious(fill)
    else:
        spPr.append(fill)


def add_soft_edge(sp, rad_pt):
    rad_emu = int(rad_pt * 12700)
    lst = _effect_lst(sp)
    for old in lst.findall(f"{{{A}}}softEdge"):
        lst.remove(old)
    se = etree.SubElement(lst, f"{{{A}}}softEdge")
    se.set("rad", str(rad_emu))


def add_shadow(sp, blur_pt, dist_pt=3, alpha=40, color="000000", direction=90):
    lst = _effect_lst(sp)
    for old in lst.findall(f"{{{A}}}outerShdw"):
        lst.remove(old)
    sh = etree.SubElement(lst, f"{{{A}}}outerShdw")
    sh.set("blurRad", str(int(blur_pt * 12700)))
    sh.set("dist", str(int(dist_pt * 12700)))
    sh.set("dir", str(int(direction * 60000)))
    sh.set("rotWithShape", "0")
    clr = etree.SubElement(sh, f"{{{A}}}srgbClr")
    clr.set("val", color)
    al = etree.SubElement(clr, f"{{{A}}}alpha")
    al.set("val", str(alpha * 1000))


def set_radial_gradient(sp, c1, c2, a1=100, a2=0, fx=0.5, fy=0.5):
    """径向渐变填充：焦点色 c1(不透明度 a1%) -> 边缘 c2(a2%)。

    fx/fy 为焦点位置(0~1)，默认中心；球体用 0.35/0.3 模拟左上高光。
    """
    spPr = sp.find(f".//{{{P}}}spPr")
    gf = etree.Element(f"{{{A}}}gradFill")
    gf.set("rotWithShape", "1")
    gs = etree.SubElement(gf, f"{{{A}}}gsLst")
    g0 = etree.SubElement(gs, f"{{{A}}}gs"); g0.set("pos", "0")
    c = etree.SubElement(g0, f"{{{A}}}srgbClr"); c.set("val", c1)
    al0 = etree.SubElement(c, f"{{{A}}}alpha"); al0.set("val", str(int(a1 * 1000)))
    g1 = etree.SubElement(gs, f"{{{A}}}gs"); g1.set("pos", "100000")
    c = etree.SubElement(g1, f"{{{A}}}srgbClr"); c.set("val", c2)
    al1 = etree.SubElement(c, f"{{{A}}}alpha"); al1.set("val", str(int(a2 * 1000)))
    path = etree.SubElement(gf, f"{{{A}}}path")
    path.set("path", "circle")
    fr = etree.SubElement(path, f"{{{A}}}fillToRect")
    half = 0.05
    fr.set("l", str(int((fx - half) * 100000)))
    fr.set("t", str(int((fy - half) * 100000)))
    fr.set("r", str(int((1 - fx - half) * 100000)))
    fr.set("b", str(int((1 - fy - half) * 100000)))
    _set_fill_el(spPr, gf)


def make_glow(sp, color=None, a1=None, a2=0):
    """半透明纯色椭圆 → 辉光：中心原色@原alpha → 边缘同色@alpha=0。"""
    spPr = sp.find(f".//{{{P}}}spPr")
    fill = spPr.find(f"{{{A}}}solidFill/{{{A}}}srgbClr")
    if fill is None and color is None:
        return False
    color = color or fill.get("val")
    al = fill.find(f"{{{A}}}alpha") if fill is not None else None
    a = a1 if a1 is not None else (
        int(al.get("val")) / 1000 if al is not None else 100)
    set_radial_gradient(sp, color, color, a1=a, a2=a2)
    return True


def set_solid_fill(sp, color, alpha=100):
    spPr = sp.find(f".//{{{P}}}spPr")
    fill = etree.Element(f"{{{A}}}solidFill")
    clr = etree.SubElement(fill, f"{{{A}}}srgbClr")
    clr.set("val", color.lstrip("#").upper())
    if alpha < 100:
        etree.SubElement(clr, f"{{{A}}}alpha").set("val", str(int(alpha * 1000)))
    _set_fill_el(spPr, fill)


def set_outline(sp, color=None, width_pt=None, alpha=100, dash=None):
    spPr = sp.find(f".//{{{P}}}spPr")
    ln = spPr.find(f"{{{A}}}ln")
    if ln is None:
        ln = etree.Element(f"{{{A}}}ln")
        anchor = None
        for tag in ("effectLst", "effectDag", "scene3d", "sp3d"):
            anchor = spPr.find(f"{{{A}}}{tag}")
            if anchor is not None:
                break
        if anchor is not None:
            anchor.addprevious(ln)
        else:
            spPr.append(ln)
    if width_pt is not None:
        ln.set("w", str(int(width_pt * 12700)))
    if color is None:
        return
    for tag in ("solidFill", "gradFill", "noFill", "pattFill"):
        for old in ln.findall(f"{{{A}}}{tag}"):
            ln.remove(old)
    fill = etree.Element(f"{{{A}}}solidFill")
    clr = etree.SubElement(fill, f"{{{A}}}srgbClr")
    clr.set("val", color.lstrip("#").upper())
    if alpha < 100:
        etree.SubElement(clr, f"{{{A}}}alpha").set("val", str(int(alpha * 1000)))
    ln.insert(0, fill)
    if dash:
        for old in ln.findall(f"{{{A}}}prstDash"):
            ln.remove(old)
        etree.SubElement(ln, f"{{{A}}}prstDash").set("val", dash)


# ---------- 动画时间轴 ----------

# (presetID, subtype, class)
_EFFECT = {
    # 入场
    "fade": (10, 0, "entr"), "wheel": (21, 1, "entr"),
    "wipe_r": (22, 2, "entr"), "wipe_l": (22, 4, "entr"),
    "wipe_d": (22, 1, "entr"), "wipe_u": (22, 8, "entr"),
    "zoom": (23, 0, "entr"), "zoom_soft": (23, 0, "entr"),
    "float": (42, 0, "entr"), "spin": (23, 0, "entr"),
    "fly_l": (2, 8, "entr"), "fly_r": (2, 1, "entr"),
    "fly_t": (2, 2, "entr"), "fly_b": (2, 4, "entr"),
    "box": (4, 0, "entr"), "diamond": (8, 0, "entr"),
    "circle": (6, 0, "entr"), "blinds": (3, 0, "entr"),
    "split": (16, 0, "entr"),
    "iris": (6, 0, "entr"), "cinematic": (23, 0, "entr"),
    "zoom_turn": (23, 0, "entr"), "rise": (42, 0, "entr"),
    # 接力：从"上一页所在的位置"滑到本页位置（不依赖 morph，所有播放器都认）
    "carry": (2, 0, "entr"),
    # 只旋入不缩放：配合 carry，让弧线边滑边拧
    "turn": (14, 0, "entr"),
    # 退出（用于页面进入时自动移走转场遮挡层）
    "exit_fade": (10, 0, "exit"), "exit_zoom": (23, 0, "exit"),
    "exit_fly_l": (2, 8, "exit"), "exit_fly_r": (2, 1, "exit"),
    "exit_fly_t": (2, 2, "exit"), "exit_fly_b": (2, 4, "exit"),
    "exit_turn": (14, 0, "exit"),
    # 强调（无限循环）
    "pulse": (3, 0, "emph"), "breathe": (3, 0, "emph"),
    "sway": (14, 0, "emph"), "rotate": (14, 0, "emph"),
    "sway6": (14, 0, "emph"), "sway10": (14, 0, "emph"),
    "sway15": (14, 0, "emph"), "sway20": (14, 0, "emph"),
    "kenburns": (3, 0, "emph"), "drift_x": (42, 0, "emph"),
    "drift_y": (42, 0, "emph"), "push": (3, 0, "emph"),
}

# animEffect filter（过渡类）
_FILTER = {
    "wheel": "wheel(1)", "wipe_r": "wipe(right)", "wipe_l": "wipe(left)",
    "wipe_d": "wipe(down)", "wipe_u": "wipe(up)", "box": "box(in)",
    "diamond": "diamond(in)", "circle": "circle(in)", "iris": "circle(in)",
    "blinds": "blinds(horizontal)", "split": "split(in)",
}


EMPH_EFFECTS = {name for name, meta in _EFFECT.items() if meta[2] == "emph"}
EXIT_EFFECTS = {name for name, meta in _EFFECT.items() if meta[2] == "exit"}


def normalize_plan_item(item):
    """兼容旧数组、可扩展数组与对象写法。"""
    if isinstance(item, dict):
        opts = dict(item)
        spid = opts.pop("spid", opts.pop("id", None))
        effect = opts.pop("effect", None)
        dur = opts.pop("dur", opts.pop("duration", None))
        delay = opts.pop("delay", 0)
        return {
            "spid": str(spid), "effect": effect, "dur": int(dur),
            "delay": int(delay), "ease": opts.pop("ease", "out"), **opts,
        }

    if not isinstance(item, (list, tuple)) or len(item) < 4:
        raise ValueError(f"plan item must have at least 4 fields: {item!r}")
    spid, effect, dur, delay = item[:4]
    opts = item[4] if len(item) > 4 else {}
    if isinstance(opts, str):
        opts = {"ease": opts}
    if not isinstance(opts, dict):
        opts = {}
    return {
        "spid": str(spid), "effect": effect, "dur": int(dur),
        "delay": int(delay), "ease": opts.pop("ease", "out"), **opts,
    }


def normalize_plan(plan):
    return [normalize_plan_item(item) for item in plan]


def _ease_attrs(ease):
    if ease in (None, "none", "linear"):
        return {}
    if ease in ("out", "ease-out", "easeOut"):
        return {"accel": "0", "decel": "50000"}
    if ease in ("in", "ease-in", "easeIn"):
        return {"accel": "50000", "decel": "0"}
    if ease in ("smooth", "ease-in-out", "easeInOut"):
        return {"accel": "30000", "decel": "30000"}
    return {}


def _behavior_par(node_id, spid, effect, dur_ms, delay_ms, ease="out", extra=None):
    """构造单个动画 <p:par>。

    入场效果 = set可见性 + 变换；强调效果 = 无 set、presetClass=emph、
    repeatCount=indefinite（持续循环）。
    extra 是编排项里的其余字段（dx/dy/scale 等效果私有参数）。
    """
    p = "{%s}" % P
    a = "{%s}" % A
    if effect not in _EFFECT:
        raise ValueError(f"unknown animation effect: {effect}")
    pid, sub, cls = _EFFECT[effect]
    is_emph = (cls == "emph")
    is_exit = (cls == "exit")
    par = etree.Element(p + "par")
    cTn = etree.SubElement(par, p + "cTn")
    cTn.set("presetID", str(pid)); cTn.set("presetClass", cls)
    cTn.set("presetSubtype", str(sub)); cTn.set("fill", "hold")
    cTn.set("nodeType", "withEffect")
    for k, v in _ease_attrs(ease).items():
        cTn.set(k, v)
    st = etree.SubElement(cTn, p + "stCondLst")
    cond = etree.SubElement(st, p + "cond"); cond.set("delay", str(delay_ms))
    child = etree.SubElement(cTn, p + "childTnLst")

    def _tgt(parent):
        tg = etree.SubElement(parent, p + "tgtEl")
        stg = etree.SubElement(tg, p + "spTgt"); stg.set("spid", str(spid))
        return tg

    def _dur_ct(parent, extra=None):
        ct = etree.SubElement(parent, p + "cTn")
        ct.set("dur", str(dur_ms)); ct.set("fill", "hold")
        for k, v in _ease_attrs(ease).items():
            ct.set(k, v)
        if extra:
            for k, v in extra.items():
                ct.set(k, v)
        s = etree.SubElement(ct, p + "stCondLst")
        c = etree.SubElement(s, p + "cond"); c.set("delay", str(delay_ms))
        return ct

    def _set_visible():
        sett = etree.SubElement(child, p + "set")
        cbh = etree.SubElement(sett, p + "cBhvr")
        ct = etree.SubElement(cbh, p + "cTn")
        ct.set("dur", "1"); ct.set("fill", "hold")
        s = etree.SubElement(ct, p + "stCondLst")
        c = etree.SubElement(s, p + "cond"); c.set("delay", str(delay_ms))
        _tgt(cbh)
        anl = etree.SubElement(cbh, p + "attrNameLst")
        an = etree.SubElement(anl, p + "attrName"); an.text = "style.visibility"
        to = etree.SubElement(sett, p + "to")
        sv = etree.SubElement(to, p + "strVal"); sv.set("val", "visible")

    def _anim_effect(filter_, transition="in"):
        ae = etree.SubElement(child, p + "animEffect")
        ae.set("transition", transition); ae.set("filter", filter_)
        cbh = etree.SubElement(ae, p + "cBhvr")
        _dur_ct(cbh)
        _tgt(cbh)

    def _set_hidden(at_ms):
        sett = etree.SubElement(child, p + "set")
        cbh = etree.SubElement(sett, p + "cBhvr")
        ct = etree.SubElement(cbh, p + "cTn")
        ct.set("dur", "1"); ct.set("fill", "hold")
        s = etree.SubElement(ct, p + "stCondLst")
        c = etree.SubElement(s, p + "cond"); c.set("delay", str(at_ms))
        _tgt(cbh)
        anl = etree.SubElement(cbh, p + "attrNameLst")
        an = etree.SubElement(anl, p + "attrName"); an.text = "style.visibility"
        to = etree.SubElement(sett, p + "to")
        sv = etree.SubElement(to, p + "strVal"); sv.set("val", "hidden")

    def _scale(from_pct, to_pct, by=None, repeat=False, auto_reverse=False):
        sc = etree.SubElement(child, p + "animScale")
        cbh = etree.SubElement(sc, p + "cBhvr")
        extra = {}
        if repeat:
            extra["repeatCount"] = "indefinite"
        if auto_reverse:
            extra["autoRev"] = "1"
        _dur_ct(cbh, extra)
        _tgt(cbh)
        if by is not None:
            by_el = etree.SubElement(sc, p + "by")
            by_el.set("x", str(by)); by_el.set("y", str(by))
        else:
            frm = etree.SubElement(sc, p + "from")
            frm.set("x", str(from_pct)); frm.set("y", str(from_pct))
            to = etree.SubElement(sc, p + "to")
            to.set("x", str(to_pct)); to.set("y", str(to_pct))

    def _motion(path, repeat=False, auto_reverse=False):
        am = etree.SubElement(child, p + "animMotion")
        am.set("origin", "layout"); am.set("path", path)
        am.set("pathEditMode", "relative"); am.set("ptsTypes", "")
        cbh = etree.SubElement(am, p + "cBhvr")
        extra = {}
        if repeat:
            extra["repeatCount"] = "indefinite"
        if auto_reverse:
            extra["autoRev"] = "1"
        _dur_ct(cbh, extra)
        _tgt(cbh)
        anl = etree.SubElement(cbh, p + "attrNameLst")
        for nm in ("ppt_x", "ppt_y"):
            an = etree.SubElement(anl, p + "attrName"); an.text = nm

    def _rot(by=None, repeat=False, auto_reverse=False, from_val=None, to_val=None):
        ar = etree.SubElement(child, p + "animRot")
        if by is not None:
            ar.set("by", str(by))
        if from_val is not None:
            ar.set("from", str(from_val))
        if to_val is not None:
            ar.set("to", str(to_val))
        cbh = etree.SubElement(ar, p + "cBhvr")
        extra = {}
        if repeat:
            extra["repeatCount"] = "indefinite"
        if auto_reverse:
            extra["autoRev"] = "1"
        _dur_ct(cbh, extra)
        _tgt(cbh)
        anl = etree.SubElement(cbh, p + "attrNameLst")
        an = etree.SubElement(anl, p + "attrName"); an.text = "r"

    # ---- 退出效果（页面进入时自动把遮挡层带走） ----
    if is_exit:
        if effect == "exit_fade":
            _anim_effect("fade", transition="out")
        elif effect == "exit_zoom":
            _scale(from_pct=100000, to_pct=0)
            _anim_effect("fade", transition="out")
        elif effect == "exit_turn":
            _scale(from_pct=100000, to_pct=0)
            _rot(by=-1080000)
            _anim_effect("fade", transition="out")
        elif effect in ("exit_fly_l", "exit_fly_r", "exit_fly_t", "exit_fly_b"):
            path = {
                "exit_fly_l": "M 0 0 L -1 0",
                "exit_fly_r": "M 0 0 L 1 0",
                "exit_fly_t": "M 0 0 L 0 -1",
                "exit_fly_b": "M 0 0 L 0 1",
            }[effect]
            _motion(path)
            _anim_effect("fade", transition="out")
        _set_hidden(delay_ms + dur_ms)
        return par

    # ---- 强调效果（无 set visible，循环） ----
    if effect == "pulse":
        _scale(by=12000, from_pct=0, to_pct=0, repeat=True, auto_reverse=True)
        return par
    if effect == "breathe":
        _scale(by=4000, from_pct=0, to_pct=0, repeat=True, auto_reverse=True)
        return par
    if effect == "sway":
        _rot(by=1500000, repeat=True, auto_reverse=True)
        return par
    if effect in ("sway6", "sway10", "sway15", "sway20"):
        # 可调幅度的往复摆动（"转过去又摆回来"），幅度单位：度 -> 1/60000 度
        deg = {"sway6": 6, "sway10": 10, "sway15": 15, "sway20": 20}[effect]
        _rot(by=deg * 60000, repeat=True, auto_reverse=True)
        return par
    if effect == "rotate":
        _rot(by=21600000, repeat=True, auto_reverse=False)
        return par
    if effect == "kenburns":
        _scale(by=6500, from_pct=0, to_pct=0, repeat=True, auto_reverse=True)
        return par
    if effect == "push":
        # 远景推近（Ken Burns 单程）：100% → 114%，只进不退 —— "一张幻灯片里的演变"
        _scale(from_pct=100000, to_pct=114000)
        return par
    if effect == "drift_x":
        _motion("M -0.012 0 L 0.012 0", repeat=True, auto_reverse=True)
        return par
    if effect == "drift_y":
        _motion("M 0 -0.012 L 0 0.012", repeat=True, auto_reverse=True)
        return par

    # ---- 入场效果 ----
    _set_visible()
    if effect == "carry":
        # 接力入场：把形状先"摆"到上一页所在的位置，再滑到本页的编排位置。
        # 相对路径（pathEditMode=relative）从 (-dx,-dy) 走到 (0,0)：
        #   第 1 点 = 当前位置-(dx,dy)，结束时正好回到编排位置。
        # dx/dy 单位是画布宽/高的比例。这样翻页时圆窗是"滑"过去的，
        # 不依赖 p14:morph（WPS / 旧版 PowerPoint 不认 morph，会直接硬切）。
        o = extra or {}
        dx = float(o.get("dx", 0.0))
        dy = float(o.get("dy", 0.0))
        _motion(f"M {dx:.5f} {dy:.5f} L 0 0")
        sc = o.get("scale")
        if sc:
            _scale(from_pct=int(round(float(sc) * 1000)), to_pct=100000)
        return par
    if effect == "fade":
        _anim_effect("fade")
    elif effect in _FILTER:
        _anim_effect(_FILTER[effect])
    elif effect in ("zoom", "zoom_soft"):
        frm = "40000" if effect == "zoom_soft" else "0"
        _scale(from_pct=frm, to_pct=100000)
        _anim_effect("fade")
    elif effect in ("float", "rise"):
        _motion("M 0 0.06 L 0 0")
        _anim_effect("fade")
    elif effect in ("fly_l", "fly_r", "fly_t", "fly_b"):
        d = {"fly_l": "M -1 0 L 0 0",
             "fly_r": "M 1 0 L 0 0",
             "fly_t": "M 0 -1 L 0 0",
             "fly_b": "M 0 1 L 0 0"}[effect]
        _motion(d)
        _anim_effect("fade")
    elif effect == "turn":
        # 只转不缩：给"跟着圆一起滑进来"的弧线用（配 carry），光圈边滑边拧
        _rot(by=21600000)
    elif effect == "spin":
        _scale(from_pct=0, to_pct=100000)
        _rot(by=21600000)
        _anim_effect("fade")
    elif effect == "zoom_turn":
        _scale(from_pct=72000, to_pct=100000)
        _rot(from_val=-480000, to_val=0)
        _anim_effect("fade")
    elif effect == "cinematic":
        _motion("M 0 0.035 L 0 0")
        _scale(from_pct=86000, to_pct=100000)
        _rot(from_val=-180000, to_val=0)
        _anim_effect("fade")
    return par


def build_timing(plan, auto_start=False):
    """根据编排表生成完整 <p:timing>（单交互组，错峰自动播放）。

    plan: [(spid, effect, dur_ms, delay_ms[, options]), ...] 或对象数组。
    同一 spid 可多次出现：先入场（如 wheel）、后强调（如 breathe，delay=入场后）。
    """
    p = "{%s}" % P
    timing = etree.Element(p + "timing")
    tnLst = etree.SubElement(timing, p + "tnLst")
    par = etree.SubElement(tnLst, p + "par")
    root = etree.SubElement(par, p + "cTn")
    root.set("id", "1"); root.set("dur", "indefinite")
    root.set("restart", "never"); root.set("nodeType", "tmRoot")
    child = etree.SubElement(root, p + "childTnLst")
    seq = etree.SubElement(child, p + "seq")
    seq.set("concurrent", "1"); seq.set("nextAc", "seek")
    main = etree.SubElement(seq, p + "cTn")
    main.set("id", "2"); main.set("dur", "indefinite")
    main.set("nodeType", "mainSeq")
    child2 = etree.SubElement(main, p + "childTnLst")
    grp = etree.SubElement(child2, p + "par")
    gc = etree.SubElement(grp, p + "cTn")
    gc.set("id", "3"); gc.set("fill", "hold")
    gs = etree.SubElement(gc, p + "stCondLst")
    gcond = etree.SubElement(gs, p + "cond")
    gcond.set("delay", "0" if auto_start else "indefinite")
    gchild = etree.SubElement(gc, p + "childTnLst")
    inner = etree.SubElement(gchild, p + "par")
    ic = etree.SubElement(inner, p + "cTn")
    ic.set("id", "4"); ic.set("fill", "hold")
    ist = etree.SubElement(ic, p + "stCondLst")
    icond = etree.SubElement(ist, p + "cond"); icond.set("delay", "0")
    ichild = etree.SubElement(ic, p + "childTnLst")
    items = normalize_plan(plan)
    for i, item in enumerate(items):
        node = _behavior_par(
            5 + i, item["spid"], item["effect"], item["dur"],
            item["delay"], item.get("ease", "out"), item,
        )
        if i == 0:
            cn = node.find(f"{{{P}}}cTn")
            cn.set("nodeType", "afterEffect" if auto_start else "clickEffect")
        ichild.append(node)
    prevC = etree.SubElement(seq, p + "prevCondLst")
    pc = etree.SubElement(prevC, p + "cond")
    pc.set("evt", "onPrev"); pc.set("delay", "0")
    ob = etree.SubElement(pc, p + "tgtEl"); so = etree.SubElement(ob, p + "sldTgt")
    nextC = etree.SubElement(seq, p + "nextCondLst")
    nc = etree.SubElement(nextC, p + "cond")
    nc.set("evt", "onNext"); nc.set("delay", "0")
    ob2 = etree.SubElement(nc, p + "tgtEl"); so2 = etree.SubElement(ob2, p + "sldTgt")
    return timing


def replace_timing(root, plan, auto_start=False):
    old = root.find(f".//{{{P}}}timing")
    if old is not None:
        root.remove(old)
    timing = build_timing(plan, auto_start=auto_start)
    anchor = root.find(f".//{{{P}}}transition")
    if anchor is None:
        anchor = root.find(f".//{{{P}}}clrMapOvr")
    if anchor is None:
        anchor = root.find(f".//{{{P}}}cSld")
    anchor.addnext(timing)
    return timing


# ---------- 页面转场 ----------

# 转场名 -> (transition 子元素名, 属性)
_TRANSITIONS = {
    "fade": ("fade", {}),
    "push_u": ("push", {"dir": "u"}),
    "push_d": ("push", {"dir": "d"}),
    "push_l": ("push", {"dir": "l"}),
    "push_r": ("push", {"dir": "r"}),
    "wipe_u": ("wipe", {"dir": "u"}),
    "wipe_d": ("wipe", {"dir": "d"}),
    "wipe_l": ("wipe", {"dir": "l"}),
    "wipe_r": ("wipe", {"dir": "r"}),
    "cover_u": ("cover", {"dir": "u"}),
    "cover_d": ("cover", {"dir": "d"}),
    "cover_l": ("cover", {"dir": "l"}),
    "cover_r": ("cover", {"dir": "r"}),
    "uncover_u": ("uncover", {"dir": "u"}),
    "uncover_d": ("uncover", {"dir": "d"}),
    "uncover_l": ("uncover", {"dir": "l"}),
    "uncover_r": ("uncover", {"dir": "r"}),
    "split_h": ("split", {"orient": "horz", "dir": "in"}),
    "split_h_out": ("split", {"orient": "horz", "dir": "out"}),
    "split_v": ("split", {"orient": "vert", "dir": "in"}),
    "split_v_out": ("split", {"orient": "vert", "dir": "out"}),
    "circle": ("circle", {}),
    "iris": ("circle", {}),
    "wheel": ("wheel", {"spokes": "4"}),
    "wheel3": ("wheel", {"spokes": "3"}),
    "wheel6": ("wheel", {"spokes": "6"}),
    "wheel8": ("wheel", {"spokes": "8"}),
    "wheel12": ("wheel", {"spokes": "12"}),
    "wheel1": ("wheel", {"spokes": "1"}),
    "doors_h": ("doors", {"orient": "horz"}),
    "doors_v": ("doors", {"orient": "vert"}),
    "comb_h": ("comb", {"orient": "horz"}),
    "zoom": ("zoom", {}),
    "cut": ("cut", {}),
    "dissolve": ("dissolve", {}),
    "newsflash": ("newsflash", {}),
    "plus": ("plus", {}),
    "wedge": ("wedge", {}),
    "strips": ("strips", {}),
    "checker": ("checker", {}),
    "blinds": ("blinds", {}),
    "random": ("random", {}),
    "morph": ("morph", {"option": "byObject"}),
    "none": (None, {}),
}


def set_transition(root, trans, spd="med", duration_ms=None):
    """给一页写入 <p:transition> 转场（页面切换动画）。

    trans 见 _TRANSITIONS；spd: slow/med/fast。
    插入位置：clrMapOvr 之后、timing 之前。
    """
    name, attrs = _TRANSITIONS.get(trans, (None, {}))
    old = root.find(f".//{{{P}}}transition")
    if old is not None:
        root.remove(old)
    if name is None:
        return None
    advanced = trans == "morph" or duration_ms is not None
    t = etree.Element(
        f"{{{P}}}transition",
        nsmap={"mc": MC, "p14": P14} if advanced else None,
    )
    if spd:
        t.set("spd", spd)
    if duration_ms is not None:
        t.set(f"{{{P14}}}dur", str(int(duration_ms)))
        t.set(f"{{{MC}}}Ignorable", "p14")
    if trans == "morph":
        t.set("advClick", "1")
        t.set(f"{{{MC}}}Ignorable", "p14")
        morph = etree.SubElement(t, f"{{{P14}}}morph")
        morph.set("option", "byObject")
    else:
        etree.SubElement(t, f"{{{P}}}{name}", **attrs)
    clr = root.find(f".//{{{P}}}clrMapOvr")
    if clr is not None:
        clr.addnext(t)
    else:
        cSld = root.find(f".//{{{P}}}cSld")
        cSld.addnext(t)
    return t


def rename_shape(root, spid, name):
    """给指定 spid 的形状设置 cNvPr@name（Morph 转场靠同名形状跨页变形）。"""
    for sid, shape in iter_shapes(root):
        if sid == str(spid):
            nv = shape.find(_NV_PATHS[etree.QName(shape).localname])
            if nv is not None:
                nv.set("name", name)
                return True
    return False


TRANSITIONS = set(_TRANSITIONS)


# ---------- 打包 ----------

def save_slides(src, dst, apply_fn):
    zin = zipfile.ZipFile(src)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        m = re.match(r"ppt/slides/slide(\d+)\.xml$", item.filename)
        if m:
            idx = int(m.group(1))
            root = etree.fromstring(data)
            apply_fn(root, idx)
            data = etree.tostring(root, xml_declaration=True,
                                  encoding="UTF-8", standalone=True)
        zout.writestr(item, data)
    zin.close()
    zout.close()
