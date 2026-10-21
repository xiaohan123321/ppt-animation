# -*- coding: utf-8 -*-
"""圆窗影像 v3 —— 只用平滑渐变 + 极细光丝 + 颗粒。

放弃"假装照片"（v2 的剪影/褶皱只会变成栅栏和迷彩）。
这里只做质感层：
  1. 三段径向渐变：焦点暖金 → 中段暗红 → 边缘近黑（gamma 衰减，边缘自然）
  2. 2~3 条极细光丝（exp 衰减、不弯曲）→ 方向感，不成迷彩
  3. 8bit 抖动防色带 + 颗粒 + 暗角
结构元素（扇环/数字/细环）交给 PPT 形状，图片只负责氛围。
"""
import os
import sys
import numpy as np
from PIL import Image

N = 1024
# 输出目录默认落在当前工作目录下，可用第 1 个参数覆盖
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.getcwd(), "window_img")
os.makedirs(OUT, exist_ok=True)

P0 = np.array([214, 142, 88], np.float32)    # 焦点暖金
P1 = np.array([124, 34, 22], np.float32)     # 暗红
P2 = np.array([52, 10, 12], np.float32)      # 深红
P3 = np.array([8, 3, 4], np.float32)         # 近黑


def ramp(fx, fy, sx=1.25):
    yy, xx = np.mgrid[0:N, 0:N].astype(np.float32) / N
    d = np.sqrt(((xx - fx) * sx) ** 2 + (yy - fy) ** 2) / 1.30
    return np.clip(d, 0, 1), xx, yy


def grad_img(fx, fy, sx=1.25, gamma=2.05, glow=0.62):
    d, xx, yy = ramp(fx, fy, sx)
    t = d ** gamma
    out = np.zeros((N, N, 3), np.float32)
    knots = [(0.00, P0), (0.34, P1), (0.66, P2), (1.00, P3)]
    for i in range(len(knots) - 1):
        a0, c0 = knots[i]
        a1, c1 = knots[i + 1]
        m = (t >= a0) & (t <= a1)
        u = np.clip((t - a0) / (a1 - a0), 0, 1)[..., None]
        # smoothstep，避免硬色阶
        u = u * u * (3 - 2 * u)
        out = np.where(m[..., None], c0 * (1 - u) + c1 * u, out)
    return out * glow, xx, yy


def add_threads(img, xx, yy, threads):
    """极细光丝：直斜线，exp 衰减，只用低幅度，不会形成花纹。"""
    for ang, c0, wd, amp, col in threads:
        u = xx * np.cos(ang) + yy * np.sin(ang)
        b = np.exp(-((u - c0) / wd) ** 2).astype(np.float32) * amp
        img = img + b[..., None] * col
    return img


def finish(img, xx, yy, vig=0.62, grain=2.4):
    # 关键：图片被裁成内切圆，只有中心区域可见。所以暗角必须从圆心算起，
    # 否则压暗全落在圆外，圆内永远是亮的。
    dc = np.sqrt((xx - 0.5) ** 2 + (yy - 0.5) ** 2) / 0.5   # 0=圆心 1=圆边
    v = 1 - 0.62 * np.clip((dc - 0.18) / 0.82, 0, 1) ** 1.25
    v = v * (1 - vig * np.clip((np.sqrt(((xx-0.5)*1.05)**2 + (yy-0.5)**2) - 0.44) / 0.5, 0, 1))
    img = img * v[..., None] * 0.74
    # 8bit 抖动，压掉渐变色带
    img = img + np.random.uniform(-0.75, 0.75, (N, N, 1)).astype(np.float32)
    img = img + np.random.normal(0, grain, (N, N, 1)).astype(np.float32)
    return np.clip(img, 0, 255).astype(np.uint8)


AMBER = np.array([92, 46, 14], np.float32)
ROSE = np.array([70, 18, 12], np.float32)
S = np.pi / 180


def make(name, fx, fy, sx=1.25, gamma=2.05, glow=0.62, vig=0.62, grain=2.4,
         threads=None):
    img, xx, yy = grad_img(fx, fy, sx, gamma, glow)
    if threads:
        img = add_threads(img, xx, yy, threads)
    out = finish(img, xx, yy, vig, grain)
    Image.fromarray(out).save(os.path.join(OUT, f"{name}.jpg"), quality=94)
    print("  ", name, f"焦点({fx:.2f},{fy:.2f})")


make("w01", 0.84, 0.08, glow=1.05, threads=[
    (-24 * S, 0.30, 0.048, 0.22, AMBER), (-20 * S, 0.58, 0.062, 0.26, ROSE)])
make("w02", 0.44, 0.66, glow=0.92, vig=0.62, threads=[
    (66 * S, 0.46, 0.050, 0.28, AMBER)])
make("w03", 0.02, 0.06, sx=1.35, glow=1.08, threads=[
    (28 * S, 0.34, 0.052, 0.20, AMBER), (32 * S, 0.62, 0.068, 0.24, ROSE)])
make("w04", 0.86, 0.20, glow=0.86, vig=0.50, grain=2.0, threads=[
    (-10 * S, 0.44, 0.055, 0.22, AMBER)])
make("w05", 0.10, 0.10, glow=1.02, threads=[
    (50 * S, 0.38, 0.050, 0.36, AMBER), (-28 * S, 0.66, 0.064, 0.24, ROSE)])
make("w06", 0.94, 0.86, glow=0.90, vig=0.60, grain=2.0, threads=[
    (0, 0.46, 0.055, 0.26, AMBER)])
make("w07", 0.08, -0.06, sx=1.30, glow=1.08, vig=0.58, threads=[
    (16 * S, 0.36, 0.055, 0.38, AMBER)])
make("w08", 0.50, 0.98, glow=1.00, vig=0.64, grain=2.0, threads=None)
make("w09", 0.92, 0.10, glow=0.95, vig=0.56, threads=[
    (68 * S, 0.42, 0.050, 0.30, AMBER)])
make("w10", 0.50, -0.04, sx=1.55, gamma=1.85, glow=1.10, vig=0.46, grain=2.0,
     threads=[(0, 0.40, 0.095, 0.20, AMBER)])

print("done")
