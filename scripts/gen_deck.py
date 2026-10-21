# -*- coding: utf-8 -*-
"""随机成稿引擎 v2 —— 内容 spec → 版式 + 图形 + 动效，一次生成 pptx 与 plan。

这一版解决四件事
  1. **文字永不溢出**：所有正文走 `K.atxt`（按字数二分缩字号，到底线再截断），
     不再写死字号。长文只会变小，不会压住下一块。
  2. **窗口不只圆**：circle / squircle / hex / pent / arch / para 六种裁切。
  3. **位置不只左右**：L / R / T / B 四种站位，内容块按可用矩形自适应。
  4. **母题、配色、底图都多了**：5 种母题 × 7 套配色 × 5 种底图。

成对产出
  画形状时把真实 spid 记下来，plan 由这些 id 生成 —— 不存在"编排出错找不到形状"。

用法
  python scripts/gen_deck.py assets/spec_demo.json --seed 7 --out build/s7
"""
import argparse
import json
import math
import os
import random
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import deckgen_kit as K          # noqa: E402
from deckgen_kit import (W, H, FONT, TITLE, LATIN, emu, rect, rrect, disc,  # noqa: E402
                         ring, arc, txt, atxt, pic_shape, paint_bg)
from PIL import Image            # noqa: E402
from pptx import Presentation    # noqa: E402
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402


# ══════════════════════════════════════════════════════════════════
# 锚点：位置 + 直径，按站位分池
# ══════════════════════════════════════════════════════════════════
ANCHORS = {
    "L": [(148, 262, 540), (168, 246, 460), (188, 300, 430), (160, 210, 400), (196, 330, 360)],
    "R": [(812, 262, 540), (792, 246, 460), (772, 300, 430), (800, 210, 400), (764, 330, 360)],
    # 上下站位的直径必须小：大圆横在中间，文字只剩一条缝
    # 上下站位必须"小圆 + 宽文字带"：圆一大，文字带就只剩一条缝
    "T": [(276, 106, 212), (684, 106, 212), (480, 100, 196)],
    "B": [(276, 434, 212), (684, 434, 212), (480, 440, 196)],
}
STYLES_FOR = {
    "toc": ("L", "R", "T", "B"),
    "chapter": ("L", "R"),
    "bullets": ("L", "R"),
    "bars": ("L", "R"),
    "cards": ("L", "R"),
    "timeline": ("L", "R"),
}
SHAPES = ("circle", "circle", "circle", "circle", "squircle", "arch")
FULL = (480, 270, 1080, "F")
# rings = 只有环筒；其余 = 环筒 + 一层点缀。保证每页都有多环结构。
MOTIFS = ("rings", "rings", "rings", "dots", "blades", "arcset", "eclipse")


class Rec:
    def __init__(self):
        self.items = []

    def add(self, sp, role, kind="text"):
        if sp is not None:
            self.items.append({"id": str(sp.shape_id), "name": sp.name,
                               "role": role, "kind": kind})
        return sp

    def of(self, role):
        return [i for i in self.items if i["role"] == role]


def text_box(a):
    """由锚点反推文字矩形：位置决定从哪一侧让开。"""
    cx, cy, d, style = a
    if style == "L":
        x0 = cx + d / 2 + 42
        return (x0, 74.0, W - 58 - x0, 404.0)
    if style == "R":
        return (58.0, 74.0, cx - d / 2 - 42 - 58, 404.0)
    if style == "T":
        y0 = cy + d / 2 + 30
        return (58.0, y0, W - 116, 486 - y0)
    if style == "B":
        return (58.0, 74.0, W - 116, cy - d / 2 - 30 - 74)
    return (58.0, 74.0, W - 116, 404.0)


# ══════════════════════════════════════════════════════════════════
# 图形母题
# ══════════════════════════════════════════════════════════════════


def ring_lattice(s, pal, rec, cx=480, cy=270, alpha=92):
    """封面虹膜环阵 —— 参考图1：厚环同心等距排开，十字缝切过，压在标题底下。

    每道环拆成 4 段 82° 象限弧，象限之间留 8° 缝，缝对齐横竖轴 ——
    "十"字缝就是这么出来的，不用叠色条。入场逐段 turn 整转 = 光圈打开。
    """
    radii = (0.72, 1.02, 1.32, 1.62, 1.92)          # ×H
    for i, kd in enumerate(radii):
        dd = H * kd
        th = 0.075 - i * 0.008                       # 内环厚、外环渐薄
        a = alpha - i * 7
        for q in range(4):
            a0 = q * 90 + 4
            rec.add(K.band(s, cx, cy, dd, a0, a0 + 82, th, pal["ink"], a,
                           name=f"kb{i}q{q}"), "lattice", "arc")
    # 一根细光柱斜穿环阵，给"刀光"
    rec.add(arc(s, cx, cy, H * 1.30, 84, 100, pal["ink"], max(8.0, W * 0.02),
                alpha=alpha + 2, name="kb6"), "lattice", "arc")


def motif_rings(s, pal, cx, cy, d, gap, rec, limit=None, rng=None, variant="classic"):
    """环筒 v2 —— 厚带 × 细缘线交替，层层镶嵌（替代细线堆叠）。

    层数恒定 6 层（带/线各半）+ 一道玻璃宽带，跨页 carry 一层对一层；
    带子是实心 blockArc（环形砖），缺口对齐同一族方向 —— 远看是铸铁环，
    近看每层有阶梯。variant 决定缺口怎么拧：
      classic  同心：缺口全朝文字侧，整齐镶嵌
      eccentric 偏心：每层圆心递进偏移 → 歪成三维筒
      spiral   旋进：每层缺口再转 40°，像光圈拧进来
      wedge    楔口：第二道换成 62° 厚楔
      bisect   双心：两组带各三层，错开 0.30d
      segments 断带：五段厚砖排开，中间留缝
      orb      全环：缺口收最小，靠透明度分层（替代原 reticle）
      lens     楔+盘：最内加一道半透明厚带压住照片缘（替代原 ticks）
    """
    rng = rng or random.Random(int(cx * 3 + cy * 7 + d))
    if variant == "ticks":
        variant = "lens"
    elif variant == "reticle":
        variant = "orb"

    def gfix(dd, want):
        r = dd / 2
        if limit is not None and r > limit:
            want = max(want, 2 * math.degrees(math.acos(
                max(-1.0, min(1.0, limit / r)))) + 10)
        return min(want, 300.0)

    def bnd(kd, th, alpha, want, name, cx2=None, cy2=None, color=None,
            span=None, a0=None):
        dd = d * kd
        x2 = cx if cx2 is None else cx2
        y2 = cy if cy2 is None else cy2
        g = gfix(dd, want)
        st = a0 if a0 is not None else gap + g / 2
        return rec.add(K.band(s, x2, y2, dd, st,
                              st + (span if span is not None else 360 - g),
                              th, color or pal["ink"], alpha, name=name),
                       "kernel", "arc")

    def edg(kd, alpha, name, cx2=None, cy2=None, color=None, w=1.1):
        x2 = cx if cx2 is None else cx2
        y2 = cy if cy2 is None else cy2
        return rec.add(ring(s, x2, y2, d * kd, color or pal["ink"], w,
                            alpha=alpha, name=name), "kernel", "ring")

    # 六层主阶梯：厚带 / 缘线交替，镶嵌的"砖与缝"
    stack = [(1.56, 0.045, 86), (1.47, None, 58), (1.32, 0.030, 80),
             (1.24, None, 54), (1.10, 0.075, 30), (1.02, None, 66)]
    if variant == "bisect":
        ox = math.cos(math.radians(gap + 90)) * d * 0.30
        oy = math.sin(math.radians(gap + 90)) * d * 0.30
        for i, (kd, th, al) in enumerate([(1.55, 0.040, 82), (1.42, None, 56),
                                          (1.18, 0.055, 76)]):
            bnd(kd, th or 0.012, al, 44 + i * 14, f"m{i}")
        for i, (kd, th, al) in enumerate([(1.30, 0.048, 80), (1.02, 0.030, 74),
                                          (0.86, None, 60)]):
            bnd(kd, th or 0.012, al, 36 + i * 12, f"m{i + 3}", cx + ox, cy + oy)
    elif variant == "segments":
        for i, kd in enumerate((1.58, 1.40, 1.22, 1.06, 0.90)):
            bnd(kd, 0.055 - i * 0.004, 78 + i, 0, f"m{i}",
                a0=gap + i * 68 + 8, span=46)
        edg(0.74, 60, "m5")
    elif variant == "pacman":
        # 吃豆人 —— 参考图3：一道巨厚环咬掉 100° 大楔口，楔口朝文字侧
        bnd(1.30, 0.085, 88, 100, "m0")
        edg(1.13, 58, "m1")
        bnd(1.02, 0.030, 74, 60, "m2")
    elif variant == "orb":
        for i, (kd, th, al) in enumerate(stack):
            if th is None:
                edg(kd, al, f"m{i}")
            else:
                bnd(kd, th, al, 6, f"m{i}", a0=gap + i * 120)
    else:
        for i, (kd, th, al) in enumerate(stack):
            if variant == "eccentric":
                k = (i / (len(stack) - 1) - 0.5) * d * 0.10
                dx = math.cos(math.radians(gap + 90)) * k
                dy = math.sin(math.radians(gap + 90)) * k
            else:
                dx = dy = 0.0
            if variant == "wedge" and i == 2:
                bnd(kd, 0.11, al, 0, f"m{i}", cx + dx, cy + dy,
                    a0=gap + 150, span=62)
                continue
            if th is None:
                edg(kd, al, f"m{i}", cx + dx, cy + dy)
            else:
                bnd(kd, th, al, 46 + i * 10, f"m{i}", cx + dx, cy + dy)
        if variant == "lens":
            # 半透明厚带压在照片缘上 —— 照片被"镶"进环座
            bnd(0.98, 0.06, 34, 90, "m9d")

    # 一道更淡的宽玻璃带压在图上，制造"玻璃"感
    bnd(1.20, 0.11, 78, 150, "m9", color=pal["paper"])


def motif_bands(s, pal, cx, cy, d, gap, rec, limit=None):
    k = 1.0 if d >= 500 else 0.9
    spec = [(1.44 * k, 0.046, 86), (1.26 * k, 0.066, 80), (1.10 * k, 0.094, 74)]
    for i, (kd, kw, span) in enumerate(spec):
        if limit is not None and d * kd / 2 + d * kw / 2 > limit:
            continue
        a0 = gap + 40 + i * 12
        rec.add(arc(s, cx, cy, d * kd, a0, a0 + span + 186, pal["ink"],
                    max(10.0, d * kw * k), alpha=84 - i * 3, name=f"m{i}"), "kernel", "arc")
    rec.add(ring(s, cx, cy, d * 1.44 * k, pal["ink"], 0.9, alpha=62, name="m9"),
            "kernel", "ring")


def motif_eclipse(s, pal, cx, cy, d, gap, rec, limit=None):
    rng = random.Random(int(cx + cy * 7))
    ang = math.radians(gap + rng.uniform(-40, 40))
    ox, oy = math.cos(ang) * d * 0.42, math.sin(ang) * d * 0.42
    rec.add(disc(s, cx + ox, cy + oy, d * 0.92, pal["ink"], 88, name="m0"), "kernel", "disc")
    rec.add(ring(s, cx + ox, cy + oy, d * 0.92, pal["ink"], 0.9, alpha=56, name="m1"),
            "kernel", "ring")
    rec.add(ring(s, cx, cy, d * 1.24, pal["ink"], 0.9, alpha=64, name="m2"), "kernel", "ring")
    rec.add(ring(s, cx, cy, d * 1.48, pal["ink"], 0.8, alpha=76, name="m3"), "kernel", "ring")


def motif_blades(s, pal, cx, cy, d, gap, rec, limit=None):
    """两道长斜刃。刃是"以圆心为中心"的长条，半长必须小于到文字的距离。"""
    for i in range(2):
        ln = d * (1.9 - i * 0.45)
        if limit is not None:
            ln = min(ln, limit * 1.85)
        if ln < d * 0.8:
            continue                      # 太短就不像刃了，宁可不画
        rec.add(rect(s, cx - ln / 2, cy - 3.0, ln, 6.0, pal["ink"], 92,
                     name=f"m{i}", rot=gap + 18 + i * 46), "kernel", "blade")
    rec.add(ring(s, cx, cy, d * 1.22, pal["ink"], 0.8, alpha=72, name="m2"), "kernel", "ring")


def motif_arcset(s, pal, cx, cy, d, gap, rec, limit=None):
    for i in range(3):
        if limit is not None and d * (1.26 - i * 0.16) / 2 > limit:
            continue
        a0 = gap + 30 + i * 54
        rec.add(arc(s, cx, cy, d * (1.26 - i * 0.16), a0, a0 + 96, pal["ink"],
                    1.0 + i * 0.25, alpha=56 + i * 4, name=f"m{i}"), "kernel", "arc")
    ang = math.radians(gap + 200)
    rec.add(disc(s, cx + math.cos(ang) * d * 0.62, cy + math.sin(ang) * d * 0.62,
                 max(8.0, d * 0.035), pal["accent"], 0, name="m3"), "kernel", "dot")


def motif_dots(s, pal, cx, cy, d, gap, rec, limit=None):
    if limit is None or d * 1.38 / 2 < limit:
        rec.add(arc(s, cx, cy, d * 1.38, gap + 24, gap + 274, pal["ink"],
                    max(9.0, d * 0.05), alpha=82, name="m0"), "kernel", "arc")
    rec.add(ring(s, cx, cy, d * 1.16, pal["ink"], 0.9, alpha=64, name="m8"), "kernel", "ring")
    for i in range(3):
        ang = math.radians(gap + 90 + i * 26)
        rec.add(disc(s, cx + math.cos(ang) * d * 0.86, cy + math.sin(ang) * d * 0.86,
                     max(5.0, d * 0.022), pal["ink"], 72, name=f"m{i + 1}"), "kernel", "dot")


MOTIF_FN = {"rings": motif_rings, "bands": motif_bands, "eclipse": motif_eclipse,
            "blades": motif_blades, "arcset": motif_arcset, "dots": motif_dots}


def motif_stack(s, pal, cx, cy, d, rec, rng, limit=None):
    """叠式 —— 图形上下重叠、直接压在照片上，打破"环永远在外围"的单层感。

    四种画法（每页抽一种）：
      cross_band  一条实心横带横贯照片，带上骑一根纸线 —— 像胶片压条
      disc_shift  半透明错位盘：圆心垂直偏 0.33d 压住照片上/下缘，盘从照片"底下"长出来
      shingle     三道横带纵向叠瓦，一道比一道窄、相互搭接，最下道压过照片下缘
      moon        半月粗环扣：只露上半/下半的厚环带，扣在照片缘上
    """
    st = rng.choice(("cross_band", "disc_shift", "shingle", "moon"))
    if st == "cross_band":
        h = d * 0.16
        y = cy + rng.choice((-0.24, -0.06, 0.14)) * d
        rec.add(rect(s, cx - d * 0.62, y - h / 2, d * 1.24, h, pal["ink"], 80,
                     name="stBand"), "kernel", "blade")
        rec.add(rect(s, cx - d * 0.62, y - 0.7, d * 1.24, 1.4, pal["paper"], 62,
                     name="stBandLn"), "kernel", "blade")
    elif st == "disc_shift":
        dy = rng.choice((-0.33, 0.33)) * d
        rec.add(K.disc(s, cx, cy + dy, d * 0.92, pal["ink"], 26, name="stDisc"),
                "kernel", "arc")
        rec.add(ring(s, cx, cy + dy, d * 0.92, pal["ink"], 1.2, alpha=64,
                     name="stDiscE"), "kernel", "ring")
    elif st == "shingle":
        for i, wd in enumerate((1.14, 0.90, 0.64)):
            h = d * 0.11
            y = cy + d * (0.16 + i * 0.15)
            rec.add(rect(s, cx - d * wd / 2, y - h / 2, d * wd, h, pal["ink"],
                         80 - i * 10, name=f"stSh{i}"), "kernel", "blade")
    else:  # moon
        top = rng.random() < 0.5
        a0, a1 = (188, 352) if top else (8, 172)
        rec.add(K.band(s, cx, cy, d * 1.06, a0, a1, 0.10, pal["ink"], 36,
                       name="stMoon"), "kernel", "arc")
    globals()["_last_stack"] = st


# ══════════════════════════════════════════════════════════════════
# 页眉 / 证据带
# ══════════════════════════════════════════════════════════════════
def header(s, rec, pal, box, eyebrow, title, lead=None):
    """标题也自适应：长度决定字号和占高，绝对不压到正文。"""
    x, y, w, h = box
    rec.add(atxt(s, x, y, w, 12, eyebrow, pal["green_m"], smax=7.6, smin=6.2,
                 lsp=1.2, name="eyebrow"), "content", "label")
    ty = y + 15
    th = min(h * 0.42, 74)
    rec.add(atxt(s, x, ty, w, th, title, pal["ink"], smax=24, smin=13, lsp=1.16,
                 bold=True, name="hd"), "content", "hd")
    y2 = ty + th + 6
    rect(s, x, y2, 40, 2.4, pal["accent"], name="rule")
    rec.add(s.shapes[-1], "content", "rule")
    y3 = y2 + 12
    if lead:
        lh = min(30, max(16, h * 0.12))
        rec.add(atxt(s, x, y3, w, lh, lead, pal["ink2"], smax=9.2, smin=7.2,
                     lsp=1.45, name="lead"), "content", "lead")
        y3 += lh + 8
    return y3


def evidence(s, rec, pal, x, w, items, y):
    n = len(items)
    colw = w / n
    rect(s, x, y - 12, w, 0.8, pal["line"], name="evline")
    rec.add(s.shapes[-1], "content", "rule")
    for i, (k, v) in enumerate(items):
        xx = x + i * colw
        if i:
            rect(s, xx - 9, y - 3, 0.8, 36, pal["line"], name=f"evsep{i}")
            rec.add(s.shapes[-1], "content", "rule")
        rec.add(atxt(s, xx, y, colw - 16, 11, k, pal["green_m"], smax=6.6, smin=5.4,
                     name=f"ev{i}k"), "content", "label")
        rec.add(atxt(s, xx, y + 12, colw - 16, 20, v, pal["ink"], smax=9.4, smin=7.2,
                     bold=True, lsp=1.2, name=f"ev{i}v"), "content", "emph")


def bignum(s, rec, pal, cx, cy, d, label, alpha=82):
    k = d / 420.0
    w, h = 300 * k, 200 * k
    rec.add(txt(s, cx - w / 2, cy - h / 2 + 2 * k, w, h, label, 128 * k, pal["paper"],
                font=LATIN, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                bold=True, alpha=alpha, name="coreNum", wrap=False), "kernel", "num")


# ══════════════════════════════════════════════════════════════════
# 内容块（全部按 box 排版 + 自动字号）
# ══════════════════════════════════════════════════════════════════
def blk_cover(s, rec, pal, box, spec, meta):
    txt(s, 0, 104, W, 16, meta.get("kicker", ""), 8.5, pal["ink3"],
        align=PP_ALIGN.CENTER, name="eyebrow")
    rec.add(s.shapes[-1], "content", "label")
    rec.add(atxt(s, 0, 128, W - 160, 104, meta["title"], pal["ink"], smax=52, smin=30,
                 lsp=1.12, font=TITLE, align=PP_ALIGN.CENTER, name="hd"),
            "content", "hd")
    rect(s, W / 2 - 32, 248, 64, 1.2, pal["accent"], name="rule")
    rec.add(s.shapes[-1], "content", "rule")
    rec.add(atxt(s, 0, 262, W - 200, 30, meta.get("subtitle", ""), pal["ink"],
                 smax=19, smin=13, lsp=1.2, font=TITLE, align=PP_ALIGN.CENTER,
                 name="lead"), "content", "lead")
    rec.add(atxt(s, 0, 300, W - 160, 18, meta.get("en", ""), pal["ink3"], smax=9,
                 smin=7, align=PP_ALIGN.CENTER, name="lead2"), "content", "label")
    rec.add(atxt(s, 0, 432, W - 120, 18, meta.get("author", ""), pal["ink3"],
                 smax=8.8, smin=7, align=PP_ALIGN.CENTER, name="foot"), "content", "foot")


def blk_toc(s, rec, pal, box, spec):
    x, y, w, h = box
    rec.add(atxt(s, x, y, w, 12, "CONTENTS", pal["green_m"], smax=8, smin=6.6,
                 name="eyebrow"), "content", "label")
    rec.add(atxt(s, x, y + 16, w, 34, spec.get("title", "四步，从知道到做到"), pal["ink"],
                 smax=21, smin=14, bold=True, name="hd"), "content", "hd")
    rect(s, x, y + 56, 40, 2.4, pal["accent"], name="rule")
    rec.add(s.shapes[-1], "content", "rule")
    items = spec["items"]
    top = y + 78
    avail = h - (top - y) - 16
    step = min(66, int(avail / max(1, len(items))))
    for i, it in enumerate(items):
        no, cn, en = it[0], it[1], (it[2] if len(it) > 2 else "")
        yy = top + i * step
        rect(s, x, yy, w, 0.8, pal["line"], name=f"row{i}line")
        rec.add(s.shapes[-1], "content", "rule")
        rec.add(atxt(s, x, yy + 6, 36, 22, no, pal["green"], smax=19, smin=12,
                     font=LATIN, name=f"num{i}"), "content", "num")
        rec.add(atxt(s, x + 42, yy + 7, 110, 17, cn, pal["ink"], smax=12.5, smin=9,
                     bold=True, name=f"row{i}"), "content", "row")
        if en:
            rec.add(atxt(s, x + 42, yy + 25, 140, 11, en, pal["green_l"], smax=6.5,
                         smin=5.5, name=f"row{i}en"), "content", "label")
        if len(it) > 3 and it[3]:
            rec.add(atxt(s, x + 160, yy + 8, w - 190, step - 14, it[3], pal["ink2"],
                         smax=8.6, smin=6.6, lsp=1.4, name=f"row{i}t"), "content", "row")


def blk_chapter(s, rec, pal, box, spec):
    x, y, w, h = box
    rec.add(atxt(s, x, y, w, 12, f"CHAPTER {spec['no']}", pal["green_m"], smax=8,
                 smin=6.6, name="eyebrow"), "content", "label")
    rec.add(atxt(s, x, y + 14, w, min(64, h * 0.3), spec["title"], pal["ink"],
                 smax=42, smin=22, font=TITLE, lsp=1.1, name="hd"), "content", "hd")
    cy = y + 14 + min(64, h * 0.3) + 10
    rect(s, x, cy, 44, 2.6, pal["accent"], name="rule")
    rec.add(s.shapes[-1], "content", "rule")
    cy += 16
    rec.add(atxt(s, x, cy, w - 16, 24, spec["lead"], pal["green"], smax=11.5, smin=8.5,
                 bold=True, lsp=1.3, name="lead"), "content", "lead")
    cy += 30
    room = (y + h) - cy
    if room < 56:                       # 剩余高度不够：只留标题和导语，宁可不讲
        return
    body_h = min(64, max(20, h * 0.22), room - 40)
    if spec.get("body"):
        rec.add(atxt(s, x, cy, w - 16, body_h, spec["body"], pal["ink2"], smax=8.8,
                     smin=6.4, lsp=1.55, name="row"), "content", "row")
        cy += body_h + 12
    items = spec.get("items", [])[:4]
    if items:
        avail = (y + h - 40) - cy
        step = 30
        while step > 22 and step * len(items) > avail:   # 放不下就压缩行距
            step -= 2
        fit_n = max(1, min(len(items), int(avail / step)))
        items = items[:fit_n]                            # 还放不下就少讲两条
        for j, (k, v) in enumerate(items):
            yy = cy + j * step
            rect(s, x, yy + 6, 9, 1.2, pal["accent"], name=f"card{j}bar")
            rec.add(s.shapes[-1], "content", "rule")
            rec.add(atxt(s, x + 16, yy, 78, 13, k, pal["green_m"], smax=7.6, smin=6,
                         name=f"row{j}t"), "content", "label")
            rec.add(atxt(s, x + 94, yy - 1, w - 94, step - 2, v, pal["ink2"],
                         smax=8.2, smin=6.2, lsp=1.35, name=f"row{j}v"), "content", "row")
    if spec.get("quote"):
        rec.add(atxt(s, x, y + h - 22, w - 12, 22, spec["quote"], pal["ink3"],
                     smax=8, smin=6.4, lsp=1.4, name="foot"), "content", "foot")


def blk_bullets(s, rec, pal, box, spec):
    x, y, w, h = box
    top = header(s, rec, pal, box, spec.get("eyebrow", ""), spec["title"], spec.get("lead"))
    rows = spec["rows"]
    bottom = y + h - (58 if spec.get("evidence") else 26 if spec.get("foot") else 4)
    step = max(20, min(56, int((bottom - top) / max(1, len(rows)))))
    for i, r in enumerate(rows):
        tag, score, name, desc = r[0], r[1], r[2], r[3]
        yy = top + i * step
        rect(s, x, yy, w, 0.8, pal["line"], name=f"row{i}line")
        rec.add(s.shapes[-1], "content", "rule")
        rec.add(atxt(s, x, yy + 6, 24, 20, tag, pal["accent"], smax=15, smin=10,
                     font=LATIN, name=f"num{i}"), "content", "num")
        rec.add(atxt(s, x + 28, yy + 6, 70, 14, name, pal["ink"], smax=9.8, smin=7.4,
                     bold=True, name=f"row{i}"), "content", "row")
        if score:
            rec.add(atxt(s, x + 28, yy + 21, 54, 11, str(score), pal["green_m"],
                         smax=7.2, smin=6, name=f"row{i}s"), "content", "label")
        rec.add(atxt(s, x + 100, yy + 7, w - 100, max(14, step - 12), desc, pal["ink2"],
                     smax=8.3, smin=6.2, lsp=1.4, name=f"row{i}t"), "content", "row")
    if spec.get("evidence") and h >= 300:
        evidence(s, rec, pal, x, w, spec["evidence"], y + h - 46)
    if spec.get("foot") and h >= 260:
        rec.add(atxt(s, x, y + h - 18, w, 18, spec["foot"], pal["ink3"], smax=8,
                     smin=6.2, lsp=1.3, name="foot"), "content", "foot")


def blk_bars(s, rec, pal, box, spec):
    x, y, w, h = box
    top = header(s, rec, pal, box, spec.get("eyebrow", ""), spec["title"], spec.get("lead"))
    bars = spec["bars"]
    bottom = y + h - (58 if spec.get("evidence") else 26 if spec.get("foot") else 4)
    step = max(18, min(32, int((bottom - top) / max(1, len(bars)))))
    lab_w = min(158, w * 0.34)
    bar_x = lab_w + 10
    bar_w = max(80, w - bar_x - 44)
    for i, (nm, v) in enumerate(bars):
        yy = top + i * step
        rec.add(atxt(s, x, yy, lab_w, step - 4, nm, pal["ink"], smax=8.2, smin=6.4,
                     name=f"row{i}"), "content", "row")
        rec.add(atxt(s, x + bar_x + bar_w + 4, yy - 1, 40, 14, str(v),
                     pal["green"] if v >= 60 else pal["accent"], smax=9, smin=7,
                     font=LATIN, align=PP_ALIGN.RIGHT, name=f"num{i}"), "content", "num")
        rect(s, x + bar_x, yy + 3, bar_w, max(2.6, step * 0.12), pal["line"], 55,
             name=f"track{i}")
        rec.add(s.shapes[-1], "content", "track")
        rect(s, x + bar_x, yy + 3, bar_w * v / 100.0, max(2.6, step * 0.12),
             pal["green"] if v >= 60 else pal["accent"], name=f"bar{i}")
        rec.add(s.shapes[-1], "content", "bar")
        rect(s, x + bar_x + bar_w * 0.6, yy - 1, 0.8, 9, pal["ink3"], 40, name=f"ln{i}")
        rec.add(s.shapes[-1], "content", "track")
    if spec.get("evidence") and h >= 300:
        evidence(s, rec, pal, x, w, spec["evidence"], y + h - 46)
    if spec.get("foot") and h >= 260:
        rec.add(atxt(s, x, y + h - 18, w, 18, spec["foot"], pal["ink3"], smax=8,
                     smin=6.2, lsp=1.3, name="foot"), "content", "foot")


def blk_cards(s, rec, pal, box, spec):
    x, y, w, h = box
    top = header(s, rec, pal, box, spec.get("eyebrow", ""), spec["title"], spec.get("lead"))
    cards = spec["cards"]
    cols = 2 if len(cards) > 2 and w > 380 else 1
    rows = math.ceil(len(cards) / cols)
    gap = 6
    cw = (w - gap * (cols - 1)) / cols
    bottom = y + h - (56 if spec.get("evidence") else 24 if spec.get("foot") else 4)
    ch = max(46, min(96, int((bottom - top - gap * (rows - 1)) / max(1, rows))))
    for i, (h1, body) in enumerate(cards):
        cx = x + (i % cols) * (cw + gap)
        cy = top + (i // cols) * (ch + gap)
        rrect(s, cx, cy, cw, ch, pal["paper2"], 0, name=f"card{i}", border=pal["line"])
        rec.add(s.shapes[-1], "content", "card")
        rect(s, cx, cy, cw, 1.6, pal["green"], name=f"card{i}top")
        rec.add(s.shapes[-1], "content", "rule")
        rec.add(atxt(s, cx + 10, cy + 8, cw - 18, 14, h1, pal["green"], smax=8.8,
                     smin=7, bold=True, name=f"row{i}"), "content", "row")
        rec.add(atxt(s, cx + 10, cy + 25, cw - 20, ch - 30, body, pal["ink2"],
                     smax=7.6, smin=5.8, lsp=1.38, name=f"row{i}t"), "content", "row")
    if spec.get("evidence") and h >= 300:
        evidence(s, rec, pal, x, w, spec["evidence"], y + h - 44)
    if spec.get("foot") and h >= 260:
        rec.add(atxt(s, x, y + h - 16, w, 16, spec["foot"], pal["ink3"], smax=8,
                     smin=6.2, lsp=1.3, name="foot"), "content", "foot")


def blk_timeline(s, rec, pal, box, spec):
    x, y, w, h = box
    top = header(s, rec, pal, box, spec.get("eyebrow", ""), spec["title"], spec.get("lead"))
    steps = spec["steps"]
    bottom = y + h - (26 if spec.get("foot") else 4)
    step = max(28, min(54, int((bottom - top) / max(1, len(steps)))))
    rect(s, x + 28, top + 2, 0.8, step * len(steps) - 12, pal["line"], name="footline")
    rec.add(s.shapes[-1], "content", "rule")
    for i, (yr, nm, body, out) in enumerate(steps):
        yy = top + i * step
        ring(s, x + 28, yy + 6, 8, pal["accent"] if i == 0 else pal["green"], 1.4,
             name=f"card{i}dot")
        rec.add(s.shapes[-1], "content", "dot")
        rec.add(atxt(s, x, yy - 2, 28, 15, yr, pal["green_m"], smax=8.8, smin=6.8,
                     name=f"num{i}"), "content", "num")
        rec.add(atxt(s, x + 42, yy - 3, 104, 16, nm, pal["ink"], smax=10.2, smin=8,
                     bold=True, name=f"row{i}"), "content", "row")
        rec.add(atxt(s, x + 42, yy + 14, w - 42, step - 30, body, pal["ink2"],
                     smax=7.6, smin=6.2, lsp=1.35, name=f"row{i}t"), "content", "row")
        if out:
            rec.add(atxt(s, x + 42, yy + step - 18, w - 42, 15, "→ " + out,
                         pal["accent"], smax=7.6, smin=6, name=f"row{i}o"),
                    "content", "row")
    if spec.get("foot") and h >= 260:
        rec.add(atxt(s, x, y + h - 16, w, 16, spec["foot"], pal["ink3"], smax=8,
                     smin=6.2, lsp=1.3, name="foot"), "content", "foot")



def blk_circles(s, rec, pal, spec, meta, win_paths):
    """一排小圆窗 —— 参考图第二张。

    每个圆：圆形裁图 + 细环 + 十字准星 + 圆内编号 + 圆下标题。
    这是"一个大环系散成几个小孔"的落点：入场用 iris（光圈开），错峰 130ms。
    """
    items = spec["items"]
    # 页面标题：小圆窗页也有一根眉线，压住顶部空间
    ttl = spec.get("title") or meta.get("subtitle") or ""
    if ttl:
        rec.add(atxt(s, 0, 58, W, 30, ttl, pal["ink"], smax=17, smin=12,
                     align=PP_ALIGN.CENTER, bold=True, name="hd"), "content", "hd")
        rec.add(atxt(s, 0, 92, W, 14, "CONTENTS / 目录", pal["ink3"], smax=7.6,
                     smin=6.4, align=PP_ALIGN.CENTER, name="kick"), "content", "kick")
        rect(s, W / 2 - 26, 114, 52, 1.0, pal["accent"], name="rule")
        rec.add(s.shapes[-1], "content", "rule")
    n = len(items)
    d = min(214.0, (W - 140) / n - 34)
    gap = (W - 2 * 70 - d * n) / max(1, n - 1)
    cy = 236.0 if ttl else 270.0
    for i, it in enumerate(items):
        cx = 70 + d / 2 + i * (d + gap)
        no = it[0]
        title = it[1]
        cn = it[2] if len(it) > 2 else ""
        if i < len(win_paths) and win_paths[i]:
            pic_shape(s, win_paths[i], cx, cy, d, "ellipse", name=f"c{i}img")
            rec.add(s.shapes[-1], f"c{i}", "image")
            rec.add(ring(s, cx, cy, d * 0.995, pal["paper"], 1.6, alpha=62,
                         name=f"c{i}ring"), f"c{i}", "ring")
            # 四象限棋盘：对角同色（纸白/墨），象限间留十字缝 —— 参考图2
            for q in range(4):
                a0 = q * 90 + 4
                if (q + i) % 2 == 0:
                    col, al = pal["paper"], 46
                else:
                    col, al = pal["ink"], 42
                rec.add(K.pie(s, cx, cy, d * 0.97, a0, a0 + 82, col, al,
                              name=f"c{i}q{q}"), f"c{i}", "ring")
            # 中心墨盘托编号
            rec.add(K.disc(s, cx, cy, d * 0.46, pal["ink"], 78,
                           name=f"c{i}hub"), f"c{i}", "ring")
        rec.add(txt(s, cx - d / 2, cy - 16, d, 34, no, 22, pal["paper"], font=LATIN,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, bold=True,
                    name=f"c{i}no"), f"c{i}", "num")
        rec.add(atxt(s, cx - d / 2 - 10, cy + d / 2 + 14, d + 20, 20, title,
                     pal["ink"], smax=11, smin=8, align=PP_ALIGN.CENTER,
                     bold=True, name=f"c{i}t"), f"c{i}", "row")
        if cn:
            rec.add(atxt(s, cx - d / 2 - 10, cy + d / 2 + 34, d + 20, 16, cn,
                         pal["ink3"], smax=6.6, smin=5.6, align=PP_ALIGN.CENTER,
                         name=f"c{i}cn"), f"c{i}", "label")


def blk_closing(s, rec, pal, box, spec, meta):
    rec.add(atxt(s, 0, 150, W - 120, 96, spec.get("title", "谢谢"), pal["paper"],
                 smax=36, smin=22, font=TITLE, lsp=1.15, align=PP_ALIGN.CENTER,
                 name="hd"), "content", "hd")
    rect(s, W / 2 - 32, 268, 64, 1.2, pal["accent"], name="rule")
    rec.add(s.shapes[-1], "content", "rule")
    rec.add(atxt(s, 0, 286, W - 200, 26, spec.get("lead", ""), pal["paper"], smax=11.5,
                 smin=8.5, lsp=1.4, align=PP_ALIGN.CENTER, name="lead"), "content", "lead")
    rec.add(atxt(s, 0, 404, W - 120, 18, meta.get("author", ""), pal["paper"],
                 smax=9.2, smin=7, align=PP_ALIGN.CENTER, name="foot"), "content", "foot")


BLOCKS = {"bullets": blk_bullets, "bars": blk_bars, "cards": blk_cards,
          "timeline": blk_timeline, "toc": blk_toc, "chapter": blk_chapter}


# ══════════════════════════════════════════════════════════════════
# 组页
# ══════════════════════════════════════════════════════════════════
def build(spec, seed, outdir, photo_bank=None):
    rng = random.Random(seed)
    pal = K.PALETTES[spec.get("palette") or rng.choice(list(K.PALETTES))]
    meta = spec["meta"]
    pages = spec["pages"]
    os.makedirs(os.path.join(outdir, "bg"), exist_ok=True)

    deck_bg = rng.choice(K.BG_KINDS)          # 整套的主底图类型
    photo_mode = bool(photo_bank)
    # 整套只用一张背景：切页时背景像素完全一致，"硬切"就看不出来了。
    # 代价是背景不再逐页变化 —— 变化交给窗口里的景（每页不同）和环。
    if photo_mode:
        # 实拍模式：D:\ppt_bg_cache 里的真照片做整页底图（双色调归入配色）
        im = Image.open(photo_bank[seed % len(photo_bank)]).convert("RGB")
        far0 = K.photo_grade(im, pal, dark=0.78, keep_color=0.30)
        near0 = None
    else:
        far0, near0, _ = paint_bg(pal, seed * 7 + 1, kind=deck_bg, quiet=None, mtn=0.92, low=0.14)
    main_far = os.path.join(outdir, "bg", "main_far.jpg")
    main_near = os.path.join(outdir, "bg", "main_near.png")
    far0.save(main_far, quality=86)
    if near0 is not None:
        near0.save(main_near, optimize=True)
    prs = Presentation()
    prs.slide_width = emu(W)
    prs.slide_height = emu(H)
    blank = prs.slide_layouts[6]

    plans, geo, pressed, motifs = {}, [], {}, []
    side = rng.choice(["L", "R"])
    prev = None

    for idx, page in enumerate(pages, 1):
        t = page["type"]
        s = prs.slides.add_slide(blank)
        rec = Rec()

        # ── 锚点：站位 + 尺寸 + 形状 ──
        if t in ("cover", "circles"):
            anchor = None          # 封面/小圆窗页不画单个大窗口
        elif t == "closing":
            anchor = FULL
        else:
            pool_keys = STYLES_FOR.get(t, ("L", "R"))
            style = rng.choice(pool_keys)
            if style in ("L", "R"):
                # 左右必须交替，且直径与上一页不同 —— 保证翻页一定有运镜
                style = "R" if side == "L" else "L"
                side = style
            cand = ANCHORS[style]
            if prev and prev[3] in ("L", "R") and style in ("L", "R"):
                cand = [a for a in cand if a[2] != prev[2]] or cand
            cx, cy, d = rng.choice(cand)
            if prev and prev[3] != "F" and abs(cx - prev[0]) < W * 0.40:
                far_c = [a for a in cand if abs(a[0] - prev[0]) >= W * 0.40]
                if far_c:
                    cx, cy, d = rng.choice(far_c)
            anchor = (cx, cy, d, style)
            bx, by, bw, bh = text_box(anchor)
            n_items = len(page.get("rows") or page.get("bars")
                          or page.get("cards") or page.get("items") or [])
            # 上下站位是"宽而矮"的盒子：条目多了必然挤，直接退回左右
            if bw < 260 or bh < 220 or (style in ("T", "B") and n_items > 4):
                style = "R" if side == "L" else "L"
                side = style
                cx, cy, d = rng.choice(ANCHORS[style])
                anchor = (cx, cy, d, style)
        geo.append(anchor)
        # 母题先定：环/弧/点这些是"沿着圆"画的，窗口就得是圆的；
        # 只有直线母题（blades）才允许用圆角方/拱形，否则环会切在方角上。
        motif = rng.choice(MOTIFS) if (anchor and t != "closing") else None
        if t == "closing":
            shape_key = "circle"
        elif motif == "blades":
            shape_key = rng.choice(SHAPES)
        else:
            shape_key = "circle"

        # ── 底图：背景恒定，只有窗口里的景在换 ──
        kind = deck_bg
        quiet = None
        if anchor:
            if anchor[3] == "L":
                quiet = "L"
            elif anchor[3] == "R":
                quiet = "R"
        # 每页只画"窗口里的景"：换 seed 就是另一处山，镜头感来自它
        if photo_mode:
            win = K.photo_grade(Image.open(
                photo_bank[(seed * 100 + idx * 7 + 3) % len(photo_bank)]
            ).convert("RGB"), pal, dark=0.88)
        else:
            _f, _n, win = paint_bg(pal, seed * 100 + idx, kind=kind, quiet=quiet,
                                   mtn=0.85 if t == "cover" else 1.0,
                                   low=0.26 if t == "cover" else 0.0)
        bp = os.path.join(outdir, "bg", f"p{idx:02d}")
        win.save(bp + "_win.jpg", quality=86)

        rect(s, 0, 0, W, H, pal["paper"], name="bg")
        p1 = s.shapes.add_picture(main_far, emu(0), emu(0), emu(W), emu(H))
        p1.name = "bgFar"
        rec.add(p1, "bg", "image")
        if near0 is not None:
            p2 = s.shapes.add_picture(main_near, emu(0), emu(0), emu(W), emu(H))
            p2.name = "bgNear"
        if t != "closing":          # 封面也压同样的雾白，切页才看不出接缝
            # 远景全幅底图/实拍照片面积大，纱加重才不抢正文
            veil = 30 if deck_bg == "yuan" else (18 if photo_mode else 16)
            rect(s, 0, 0, W, H, pal["paper"], veil, name="bgVeil")

        # 封面：整页环阵（贯穿画面、多个圆心相交）
        if t == "cover":
            ring_lattice(s, pal, rec)

        # ── 窗口 + 母题 ──
        if anchor and t != "closing":
            cx, cy, d, style = anchor
            prst = K.PRESETS[shape_key]
            pic_shape(s, bp + "_win.jpg", cx, cy, d, prst,
                      adj=0.07 if shape_key == "squircle" else None, name="coreImg")
            ring(s, cx, cy, d, pal["green"], 1.1, name="coreRing")
            rec.add(s.shapes[-1], "kernel", "ring")
            rec.add(s.shapes[-2], "kernel", "image")
            if shape_key == "circle":
                ring(s, cx, cy, d * 0.968, pal["paper"], 0.5, alpha=58, name="coreRing2")
                rec.add(s.shapes[-1], "kernel", "ring")
            gap = {"L": 0.0, "R": 180.0, "T": 90.0, "B": 270.0}[style]
            # 圆心沿"缺口方向"到文字栏的距离：环超出这个半径就让缺口
            _bx, _by, _bw, _bh = text_box(anchor)
            if style == "L":
                limit = _bx - 8 - cx
            elif style == "R":
                limit = cx - (_bx + _bw + 8)
            elif style == "T":
                limit = _by - 8 - cy
            else:
                limit = cy - (_by + _bh + 8)
            if t == "chapter":
                v = "pacman"      # 章节页固定吃豆人大环 —— 参考图3
            else:
                v = rng.choice(("classic", "eccentric", "spiral", "wedge", "lens",
                                "bisect", "segments", "orb"))
            # 空间窄（窗口贴页边、文字又宽）时，整圈的环会被缺口吃掉只剩两段小弧，
            # 这种页换成"断环"—— 短弧成组反而立得住。
            if limit < 210 and v in ("classic", "eccentric", "spiral", "bisect"):
                v = "segments"
            # 环筒永远是底：不再有"这页几乎没有环"的页面
            motif_rings(s, pal, cx, cy, d, gap, rec, limit=max(30.0, limit), variant=v)
            if motif != "rings":
                MOTIF_FN[motif](s, pal, cx, cy, d, gap, rec, limit=max(30.0, limit))
            globals()["_last_stack"] = ""
            if t != "chapter":
                # 叠式：图形上下重叠压住照片 —— 每个内容页随机一种
                motif_stack(s, pal, cx, cy, d, rec, rng, limit=max(30.0, limit))
            globals()["_last_motif"] = (f"rings/{v}" + ("" if motif == "rings" else f" +{motif}")
                                        + f" +{globals().get('_last_stack', '')}")
            if page.get("num", True):
                bignum(s, rec, pal, cx, cy, d,
                       page.get("no") if t == "chapter" else f"{idx:02d}",
                       80 if t == "chapter" else 84)
        elif t == "closing":
            pic_shape(s, bp + "_win.jpg", 480, 270, 1180, "ellipse", name="coreImg")
            ring(s, 480, 270, 1180, pal["green"], 1.0, name="coreRing")
            rec.add(s.shapes[-1], "kernel", "ring")
            rec.add(s.shapes[-2], "kernel", "image")
            rect(s, 0, 0, W, H, pal["green"], 44, name="bgVeil2")

        # ── 内容 ──
        if t == "cover":
            blk_cover(s, rec, pal, (0, 0, W, H), page, meta)
        elif t == "circles":
            wins = []
            for i in range(len(page["items"])):
                if photo_mode:
                    w2 = K.photo_grade(Image.open(
                        photo_bank[(seed * 211 + idx * 17 + i * 5) % len(photo_bank)]
                    ).convert("RGB"), pal, dark=0.88)
                else:
                    _f2, _n2, w2 = paint_bg(pal, seed * 211 + idx * 17 + i, kind=kind,
                                            mtn=1.0, low=0.10)
                wp = os.path.join(outdir, "bg", f"p{idx:02d}c{i}.jpg")
                w2.save(wp, quality=86)
                wins.append(wp)
            blk_circles(s, rec, pal, page, meta, wins)
        elif t == "closing":
            blk_closing(s, rec, pal, (0, 0, W, H), page, meta)
        else:
            BOX = text_box(anchor)
            if t in BLOCKS:
                BLOCKS[t](s, rec, pal, BOX, page)
            else:
                raise ValueError(f"未知页类型: {t}")

        # ── 章节页探头：下一章的编号从页边探进来 —— 参考图3 右缘的 02 ──
        if t == "chapter":
            chs = [pg for pg in pages if pg.get("type") == "chapter"]
            ci = chs.index(page)
            if ci + 1 < len(chs):
                nxt = chs[ci + 1]
                no2 = str(nxt.get("no") or f"{ci + 2:02d}")
                rec.add(atxt(s, W - 92, 206, 170, 76, no2, pal["ink"], smax=52,
                             smin=40, name="peekno", alpha=26), "content", "peek")
                ttl2 = str(nxt.get("title") or "")[:4]
                if ttl2:
                    rec.add(atxt(s, W - 92, 292, 170, 24, ttl2, pal["ink"],
                                 smax=15, smin=11, name="peekt", alpha=26),
                            "content", "peek")

        # ── 页边角标 ──
        for i, (dx, dy, sx, sy) in enumerate([(30, 30, 1, 1), (W - 30, 30, -1, 1),
                                              (30, H - 30, 1, -1), (W - 30, H - 30, -1, -1)]):
            rect(s, dx if sx > 0 else dx - 15, dy if sy > 0 else dy - 0.8, 15, 0.8,
                 pal["green"], 62, name=f"cnr{i}h")
            rec.add(s.shapes[-1], "deco", "rule")
            rect(s, dx if sx > 0 else dx - 0.8, dy if sy > 0 else dy - 15, 0.8, 15,
                 pal["green"], 62, name=f"cnr{i}v")
            rec.add(s.shapes[-1], "deco", "rule")

        motifs.append(globals().get("_last_motif", ""))
        plans[str(idx)] = make_plan(rec, prev, anchor, idx)
        if K.OVERFLOWS:
            pressed[idx] = list(K.OVERFLOWS)
        K.OVERFLOWS.clear()
        prev = anchor

    os.makedirs(outdir, exist_ok=True)
    pptx = os.path.join(outdir, "deck.pptx")
    prs.save(pptx)
    pj = os.path.join(outdir, "plan.json")
    json.dump(plans, open(pj, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return pptx, pj, geo, pal, deck_bg, pressed, motifs


def make_plan(rec, prev, cur, idx):
    plan = []
    # 背景 Ken Burns：单程慢推近 9s —— "一张幻灯片里的远景演变"（参考图4/5）
    for it in rec.of("bg"):
        plan.append([it["id"], "push", 9000, 0, {"ease": "out"}])
    if idx > 1 and prev and cur and prev[3] != "F":
        dx = round((prev[0] - cur[0]) / W, 5)
        dy = round((prev[1] - cur[1]) / H, 5)
        sc = round(prev[2] / cur[2] * 100, 2)
        opts = {"dx": dx, "dy": dy, "ease": "out"}
        if abs(sc - 100) > 1:
            opts["scale"] = sc
        # 接力不再"六个环一起动"：照片（镜头）先到，环按远近逐层追上来，
        # 每层再给不同的旋转时长 —— 内环快、外环慢，运动里就有了纵深。
        ker = rec.of("kernel")
        lead = [k for k in ker if k["kind"] in ("image", "ring")]
        rest = [k for k in ker if k not in lead]
        for it in lead:
            plan.append([it["id"], "carry", 950, 0, dict(opts)])
        for k, it in enumerate(rest):
            dly = 80 + k * 42
            plan.append([it["id"], "carry", 950, dly, dict(opts)])
            if it["kind"] in ("arc", "blade"):
                # 每层转速不同：900 + k*110 → 六层转出"错位光圈"
                plan.append([it["id"], "turn", 900 + k * 110, dly + 40,
                             {"ease": "out"}])
        # 落位后最外层慢慢摆，画面一直在动
        if rest:
            plan.append([rest[0]["id"], "sway6", 7000, 1500, {"ease": "out"}])
        base = 950
    elif idx > 1 and cur:
        base = 120
        for it in rec.of("kernel"):
            plan.append([it["id"], "zoom", 1000, 120, {"ease": "out"}])
    else:
        base = 0

    for k, it in enumerate(rec.of("lattice")):
        # 厚带/方环 turn 整转入场（动转感），细缘线 fade 追上来
        ef = "turn" if it["kind"] == "arc" else "fade"
        plan.append([it["id"], ef, 860 if ef == "turn" else 700,
                     100 + k * 60, {"ease": "out"}])
    groups = sorted({i["role"] for i in rec.items
                     if i["role"].startswith("c") and i["role"][1:].isdigit()})
    for gi, role in enumerate(groups):
        for it in rec.of(role):
            if it["kind"] in ("image", "ring", "cross"):
                ef = "iris"
            elif it["kind"] == "num":
                ef = "zoom_turn"
            else:
                ef = "fade"
            plan.append([it["id"], ef, 720, 140 + gi * 130, {"ease": "out"}])

    for i, it in enumerate(rec.of("content")):
        t = max(0, base - 60) + i * 52
        k = it["kind"]
        if k == "bar":
            plan.append([it["id"], "wipe_r", 760, t, {"ease": "out"}])
        elif k in ("track", "rule"):
            plan.append([it["id"], "fade", 340, t, {"ease": "out"}])
        elif k == "num":
            plan.append([it["id"], "zoom_turn", 620, t, {"ease": "out"}])
        elif k == "hd":
            plan.append([it["id"], "cinematic", 900, t, {"ease": "out"}])
        elif k in ("card", "row"):
            plan.append([it["id"], "cinematic", 660, t, {"ease": "out"}])
        else:
            plan.append([it["id"], "fade", 600, t, {"ease": "out"}])
    for i, it in enumerate(rec.of("deco")):
        plan.append([it["id"], "fade", 420, base + 820 + i * 40, {"ease": "out"}])
    return {"transition": "cut", "transition_spd": "med", "transition_ms": 0,
            "auto_start": True, "plan": plan}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("--seed", type=int, default=1)
    ap.add_argument("--out", default=None)
    ap.add_argument("--palette", default=None, choices=sorted(K.PALETTES),
                    help="覆盖 spec 里的配色")
    ap.add_argument("--bg-dir", default=r"D:\ppt_bg_cache",
                    help="实拍背景图库目录")
    ap.add_argument("--no-photos", action="store_true",
                    help="禁用实拍照片，强制程序山水")
    a = ap.parse_args()
    spec = json.load(open(a.spec, encoding="utf-8"))
    if a.palette:
        spec["palette"] = a.palette
    # 实拍背景库：D:\ppt_bg_cache 有图就自动用真照片（--no-photos 退回程序山水）
    import glob as _glob
    bank = []
    if not a.no_photos:
        bank = sorted(_glob.glob(os.path.join(a.bg_dir, "*.jpg")))
        print(f"   实拍图库：{len(bank)} 张（{a.bg_dir}）" if bank
              else f"   图库为空（{a.bg_dir}），用程序山水；可跑 scripts/fetch_bg.py 下载")
    out = a.out or os.path.join("build", f"seed{a.seed}")
    pptx, pj, geo, pal, bg, pressed, motifs = build(spec, a.seed, out, photo_bank=bank)
    # 烘焙动效版：plan.json 里的动画时间轴直接写进 pptx —— 不再需要手动跑 apply_anim
    import apply_anim
    anim_pptx = os.path.join(out, "deck_动效版.pptx")
    apply_anim.main(pptx, anim_pptx, pj)
    print(f"-> {pptx}\n-> {pj}\n-> {anim_pptx}")
    print(f"   底图类型 {bg}")
    if pressed:
        print("   ⚠ 以下文本框被压缩（字号退让或截断），扫一眼这几页：")
        for i, items in sorted(pressed.items()):
            cut = [n for n, tr in items if tr]
            print(f"     P{i:02d}  {len(items)} 处；" + ("截断: " + ",".join(cut[:4]) if cut else "仅缩字号"))
    for i, g in enumerate(geo, 1):
        pt = spec["pages"][i - 1].get("type", "?")
        if g is None:
            print(f"  P{i:02d} {'封面' if pt == 'cover' else pt}")
        elif g[3] == "F":
            print(f"  P{i:02d} 铺满 d=1080")
        else:
            print(f"  P{i:02d} {g[3]}  {g[0]:.0f},{g[1]:.0f}  d={g[2]:.0f}   {motifs[i - 1]}")


if __name__ == "__main__":
    main()
