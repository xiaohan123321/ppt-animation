# 红金电影感 · 圆形遮罩动效配方

这套配方来自“手搓 PPT”类短视频的关键视觉语法：大画幅暖色背景、巨号英文数字、
大圆环/圆盘、深色半透明信息板，以及圆形遮罩推进或 Morph 圆环延续。

## 视觉层

从上到下建议 5 层：

1. **底层**：深红黑纯色，避免动画首帧闪白。
2. **背景图**：全页 `p:pic`，红金舞台、丝绸、旗帜或建筑纹理。
3. **氛围层**：大圆环、圆盘、细描边或径向暖光。
4. **巨号字**：`ONE / TWO / THREE` 或年份，低饱和暖白。
5. **内容层**：章节标签、主标题、正文、卡片。

圆环不要贴边过满。建议直径占画幅 0.55~0.8，圆心略偏离画面中心，形成运动空间。

## 推荐动画

```json
["2", "fade", 700, 0],
["2", "kenburns", 7000, 700],
["4", "zoom_soft", 1000, 120],
["4", "rotate", 16000, 1120],
["3", "cinematic", 900, 180],
["5", "zoom_turn", 750, 320],
["6", "float", 650, 450],
["7", "fade", 420, 650],
["8", "zoom", 520, 760]
```

层间不要严格排队：背景开始 0.7s 后，圆环与巨号字已经进入，形成“背景仍在推进、
前景已经落位”的空间感。

## 圆环转场：两种实现

### A. Morph 连续圆环（首选）

相邻页都保留一个名为 `camera_ring` 的大圆环，每页改变它的位置与尺寸：

```json
"transition": "morph",
"transition_spd": "slow",
"morph": [{"spid": "4", "name": "camera_ring"}]
```

效果：圆环从画面右侧流入下一页左侧，尺寸连续变化，像镜头穿过同一只圆环。

注意：

- 每页的 `camera_ring` 必须是独立形状，不能一部分在组合里、一部分在组合外。
- 相邻页的圆环可以尺寸不同，但不要一页有填充、一页完全无填充。
- 如果同名形状附近还有大量文字，Morph 会让文字也参与交叉淡化；用 `fade`/`float`
  给正文单独入场，避免“全页双曝光”。

### B. 内置圆形/轮盘转场（兼容性更高）

```json
"transition": "iris",
"transition_spd": "slow"
```

或：

```json
"transition": "wheel",
"transition_spd": "slow"
```

`iris` 从中心圆展开；`wheel` 像轮盘扫过。它们不依赖同名形状，跨软件更稳。

## 版式规则

- 主标题与巨号英文不要争夺同一视觉区域。巨号英文通常放在顶部或右上，主标题在左中。
- 正文使用 1.15~1.35 倍行距，单行不超过 18~22 个汉字。
- 深色半透明卡片配 0.8~1.2 磅金色描边与 16~24 磅柔和阴影。
- 卡片自身用 `zoom`，文字用 `fade`，卡片与文字间隔 70~110ms。
- 所有元素都围绕一个主视觉轴：圆环、巨号数字或人物照片，不要同时制造多个焦点。

## 颜色与字体

- 主色：深红 `#3A0E12`、朱红 `#B73523`、橙色高光 `#E46B2E`
- 强调：香槟金 `#E5AD51`、浅金 `#F6DA9F`
- 文字：暖白 `#FFF4E4`、浅米 `#E8D9C4`
- 标题：宋体/思源宋体；英文巨号：Arial/Helvetica Black/Impact
- 正文：微软雅黑/思源黑体

## 不要这样做

- 不要给所有元素同时加 `pulse`；只让圆环或一个核心数字呼吸。
- 不要用大磅数 `softEdge` 制造光晕，会发雾；用 `glow`。
- 不要让圆环、背景和巨号字都以接近速度旋转；背景慢、圆环中速、文字只做短入场。
- 不要在同一个入口同时叠 5 种行为；视觉上会像特效堆叠而不是镜头运动。
- 不要跨页对所有元素做 Morph；只让一个核心形状延续。

## 可直接运行的示例

示例靶子是一份 4 页小样片 `assets/demo.pptx`（960×540，封面 / 内容 / 内容 / 满版收尾）。
它由脚本生成，形状 id 是确定的，所以两份 plan 拿来就能跑：

- `assets/demo.pptx` —— 靶子（`python scripts/make_demo_deck.py` 可重新生成）
- `assets/demo_cinematic_circle_plan.json` —— 本案例的完整编排（carry + turn）
- `assets/example_plan.json` —— 最小可用编排（只做入场 + 转场，不带 carry）

```powershell
# 先查形状 id（换自己的 deck 时第一步永远是这个）
python scripts/inspect_slides.py assets/demo.pptx --compact

# 预检：spid / 效果名 / 转场是否都对得上
python scripts/preflight.py assets/demo.pptx assets/demo_cinematic_circle_plan.json

# 写动画，再补 name（不补的话部分播放器打不开）
python scripts/apply_anim.py assets/demo.pptx assets/demo_animated.pptx assets/demo_cinematic_circle_plan.json
python scripts/fix_names.py assets/demo_animated.pptx

# 看效果：carry 的运动**必须用 render_deck**（它按 carry 合成圆窗的位移和缩放）。
# render_transition 只处理页面转场，这套案例全是 cut 转场，用它只会得到 4 张静帧。
python scripts/render_deck.py assets/demo.pptx assets/demo_cinematic_circle_plan.json assets/demo_carry.gif --width 640 --hold 0.8
```

> 换成自己的 PPT 时，只需要把 plan 里的 spid 换成 `inspect_slides` 输出里的 id —— 
> 示例里是 3=图片、4=圆环、5=弧线、6=标题、7=眉标、8=分隔线。其余参数直接沿用。
> 现成产物：`assets/demo_animated.pptx`（已写好动画）、`assets/demo_carry.gif`（6.4 秒预览）。
