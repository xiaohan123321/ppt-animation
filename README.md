# ppt-animation 👾

> 电影感 PPT 动画生成器 —— 圆形遮罩转场、轮盘光圈、Ken Burns 远景推近、多环镶嵌版式，一条命令从内容直接生成整套动效 PPT。

一套 WorkBuddy / Claude 风格的 **Agent Skill**：既可以给已有 PPT/PPTX 叠加高级动画（直接重写 OOXML 动画时间轴），也可以从一份内容 spec **现场生成整套带动画的演示文稿**。

## ✨ 特性

**生成器路线（`gen_deck.py`）—— 从内容到成片**

- 🎨 **9 套配色**：青绿 / 黛青 / 赭金 / 烟灰 / 青瓷 / 紫檀 / 松绿 / 玄铁 / 绛朱
- 💍 **厚带环筒**：实心 blockArc 环形砖 × 细缘线交替，8 种结构 DNA（同心 / 偏心 / 旋进 / 楔口 / 镶镜 / 双心 / 断带 / 全环），每页随机
- 🌀 **虹膜封面**：5 道厚环 × 4 象限弧，十字缝对齐横竖轴，逐段整转入场 = 光圈打开
- 🀄 **四象限棋盘目录** + **吃豆人章节页**（下一章编号从页边探头）
- 📚 **叠式层**：横贯带 / 错位盘 / 叠瓦 / 半月，图形上下重叠压住照片，跨页 carry 一起搬移
- 🏔️ **真实照片背景**：自动搜索实拍风光（Bing/Openverse/Wikimedia），双色调归入配色系统；也可退回程序化山水（numpy 手绘，6 种画法）
- 🔭 **Ken Burns 远景推近**：每页背景 9s 单程 100%→114%，"一张幻灯片里的远景演变"
- ✂️ **文字永不溢出**：字符宽度估算 + 二分字号 + 截断保护，长内容自动退让
- 🎬 **动画全烘焙**：plan.json 的编排（carry 接力 / iris / turn / sway / push…）直接写进 pptx，放映即有

**动画路线（`apply_anim.py` / `anim_engine.py`）—— 给现有 PPT 加动效**

- 30+ 效果：morph 式延续、wheel 轮盘、iris 光圈、cinematic、zoom_turn、sway 摆动、drift 漂移、kenburns…
- 只写 OOXML `<p:timing>`，不依赖第三方放映器；WPS / 旧版 PowerPoint 自动降级
- plan.json 是纯数据编排表：`(spid, effect, duration, delay, options)`，可程序化生成

## 📦 安装

```bash
pip install -r requirements.txt
```

需要 Python 3.10+，无其他系统依赖。

## 🚀 快速开始

```bash
# 1.（可选）建实拍背景图库 —— 缓存在 D:\ppt_bg_cache，目录可改
python scripts/fetch_bg.py --count 12

# 2. 一条命令生成整套动效 PPT
python scripts/gen_deck.py assets/spec_demo.json --seed 7 --palette xuantie --out build/s7
# → build/s7/deck.pptx + plan.json + deck_动效版.pptx + bg/*

# 3. 校验编排与 spid 对应关系
python scripts/preflight.py build/s7/deck.pptx build/s7/plan.json

# 4.（可选）渲染转场 GIF 预览
python scripts/render_deck.py build/s7/deck.pptx build/s7/plan.json build/s7.gif
```

换 seed 就是另一套片子：配色、锚点序列、图形母题、山形全变，排版纪律不变。

给**已有** PPT 加动画：

```bash
python scripts/apply_anim.py 已有.pptx 已有_动效版.pptx plans.json
```

## 🗂️ 目录结构

```
├── SKILL.md                 # Agent Skill 入口（工作流 + 纪律）
├── scripts/
│   ├── gen_deck.py          # 生成器：spec → pptx + plan.json + 动效版
│   ├── deckgen_kit.py       # 绘图原语 / 配色 / 底图画法 / 文字自适应
│   ├── anim_engine.py       # OOXML 动画引擎（30+ 效果）
│   ├── apply_anim.py        # 把 plan.json 烘焙进 pptx
│   ├── fetch_bg.py          # 实拍背景图库下载（Bing/Openverse/Wikimedia）
│   ├── preflight.py         # plan ↔ spid 校验
│   ├── render_deck.py       # LibreOffice + pymupdf 转场预览
│   └── ...
├── references/              # 版式纪律 / 动画指南 / 生成器规格
├── assets/                  # 示例 spec / 示例 deck / 动画 demo
└── requirements.txt
```

## 🧠 设计原则

1. **一套背景贯穿全片**——切页帧完全一致，硬切也看不出接缝
2. **环的层数恒定**——跨页 carry 时"上一页的第 3 层"对上"本页的第 3 层"
3. **plan.json 是唯一真相**——pptx 里的每个 spid 都是生成时记录的真实 id
4. **内容比图形早 60ms 起**——窗口还在滑，第一行字已经淡入

## ⚠️ 背景图版权须知

`fetch_bg.py` 默认从 **Bing 每日壁纸**抓取实拍照片——这些图片版权归微软及原摄影师所有，
授权范围通常仅限桌面壁纸。生成内容若要**公开发布或商用**，请自行替换为可商用的图源
（如 Openverse/Wikimedia CC 图库，脚本已内置这两个兜底源），或使用你自有版权的照片。
仓库本身**不包含、也不分发**任何壁纸图片。

## 📄 License

[MIT](LICENSE) © xiaohan123321
