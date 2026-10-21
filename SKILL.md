---
name: ppt-animation
description: 给 PPT/PPTX 叠加电影感高级动画——圆形遮罩/轮盘转场、Morph 圆环延续、真缩放/浮起/旋转、背景 Ken Burns、辉光渐变、柔化阴影与逐元素错峰。直接重写 OOXML 动画时间轴、形状特效和转场，支持方图/竖版/横版真实画幅、图片与组合形状预览。触发词：高级动画、电影感 PPT、圆形遮罩、轮盘转场、炫酷动效、手搓PPT、动画 PPT、PPT 动效、presentation animation、动效升级。当用户要求给已有 PPT 加动画、把动画做得更高级/更炫、或参考「手搓PPT」类视频风格时使用。
agent_created: true
---

# PPT 高级动画引擎

为 PPT/PPTX 叠加真正的高级动画：圆形遮罩与轮盘推进、Morph 圆环延续、真缩放/浮起/
旋转、背景慢速漂移、辉光渐变、柔化阴影、逐元素错峰与精确时长。核心是直接重写
OOXML（`p:timing` 动画时间轴 + `p:spPr` 形状特效 + `p:transition` 页面转场）。

## 工作流

### 路线 A：从内容直接生成（推荐，不需要现成的 PPT）

`scripts/gen_deck.py` 是**生成器**，不是模板：给它一份内容 spec，它按 seed 现场组装
版式、图形母题、底图和动效，**同时产出 pptx 和 plan.json** —— spid 是生成时记下来的，
不存在"编排出错找不到形状"。

**背景用真实照片**：先跑 `python scripts/fetch_bg.py --count 12`
（Bing 每日壁纸 → Openverse → Wikimedia → Picsum 依次兜底，图库存
`D:\ppt_bg_cache`）。图库非空时 gen_deck 自动用实拍照片做整页底图和圆窗，
并按配色做双色调归一（`photo_grade`：near→green_m→paper2 三点映射 + 对比增强），
风格不散；`--no-photos` 退回程序山水，`--bg-dir` 换图库目录。

```powershell
python scripts/gen_deck.py assets/spec_demo.json --seed 7 --out build/s7
# → build/s7/deck.pptx  build/s7/plan.json  build/s7/deck_动效版.pptx  build/s7/bg/*.jpg
# （动效版自动烘焙，无需再跑 apply_anim；apply_anim 仅用于给旧 deck 补动画）
python scripts/preflight.py build/s7/deck.pptx build/s7/plan.json
python scripts/render_deck.py build/s7/deck.pptx build/s7/plan.json build/s7.gif --width 720
```

同一份 spec 换个 seed 就是另一套片子：配色（9 套，含玄铁/绛朱新风格，`--palette` 可指定）、
锚点序列、图形母题、
山形全变，但排版纪律不变（文字栏不被图形侵占、圆窗跨页必位移、深色处必压雾白）。
spec 格式与词表见 `references/generator.md`，示例内容 `assets/spec_demo.json`。

### 路线 B：给现成的 PPT 加动画

1. **摸画幅和形状**：`python scripts/inspect_slides.py <src.pptx> shapes.json`
   导出每页真实画幅、图片/组合/形状的 `cNvPr@id`（spid）、名称、位置、尺寸、
   文本与视觉角色。先确认背景图、圆环、标题、正文分别对应哪个 spid。
2. **写编排** `plans.json`：按页（1-based 字符串键）声明
   `delete`/`glow`/`shadow`/`solid`/`outline`/`plan`。
   格式与效果枚举见 `references/anim_guide.md`。
   **改之前先跑一遍现成示例**：`assets/demo.pptx` 是 4 页小样片，
   `assets/example_plan.json`（最小入场）和 `assets/demo_cinematic_circle_plan.json`
   （carry 圆窗案例）都是对着它写的，能直接跑通：
   `python scripts/preflight.py assets/demo.pptx assets/example_plan.json`。
   案例讲解见 `references/cinematic_circle.md`。
3. **应用动画**：`python scripts/apply_anim.py <src.pptx> <dst.pptx> <plans.json>`
4. **补兼容**：`python scripts/fix_names.py <dst.pptx>`（补 cNvPr 缺的 name，
   否则 LibreOffice/WPS 报 "source file could not be loaded"）
5. **预检**：`python scripts/preflight.py <dst.pptx> <plans.json>`
   检查 spid、效果名、转场、缺 name 与 Morph 目标；有问题时以非零退出码结束。
6. **GIF 预览**：`python scripts/render_gif.py <dst.pptx> <plans.json> <out_dir> --width 960`
   逐页生成动画 GIF；`python scripts/render_transition.py <dst.pptx> <plans.json> <out.gif>`
   生成带页面转场的完整流程 GIF。预览会读取真实画幅、背景、图片和组合形状。
7. **真实渲染预览（推荐给客户看）**：
   `python scripts/render_deck.py <dst.pptx> <plans.json> <out.gif> --width 960 --hold 1.2 --fps 30 --mp4 <out.mp4>`
   先用 LibreOffice 把每页导出成精确图像，再按转场合成帧。**图片、透明度、字体全部保真**；
   `render_gif.py` 走的是自绘矢量路径，遇到 JPEG 填充/椭圆裁切/复杂透明会失真
   （图片画成带框矩形、底色发白），只适合快速自检，不要拿去给用户看成品。
   **深色/渐变大背景一定要出 mp4 而不是 gif**：256 色在任何抖动策略下都会把近黑渐变压成
   整块色带或满屏噪点，看上去像画面脏了；H.264 的 YUV420 渐变平滑，体积还更小。

## 依赖

- Python 3 + lxml + Pillow（`pip install -r requirements.txt`；应用动画仅需 lxml）
- 渲染预览另需 `pymupdf`；出 mp4 需 ffmpeg（本机 `D:\ACLOS\Cross\recorder-release\ffmpeg.exe`）。
  编码器要按序回退 `libo264rt → h264_amf → h264_mf`：录屏版 ffmpeg 里 `h264_mf` 常报
  `MF_E_INVALIDMEDIATYPE`，没有对应硬件时 `h264_amf` 也会挂。
- 字体：微软雅黑 `msyh.ttc`、宋体 `simsun.ttc`、黑体 `simhei.ttf`、楷体 `simkai.ttf`
  （GIF 按 PPT 中的字体名称自动选择，缺失时回退到微软雅黑）
- 示例资产 `assets/`：`demo.pptx`（4 页小样片，由 `scripts/make_demo_deck.py` 生成，
  形状 id 固定）、`example_plan.json`、`demo_cinematic_circle_plan.json`、
  `demo_animated.pptx`、`demo_carry.gif`。改了 `make_demo_deck.py` 要重新生成，
  否则 plan 里的 spid 会对不上。

## 关键约束

- 光晕效果用 `glow`（径向渐变透明衰减），**不要用大磅数柔化边缘**（>5 磅叠加会发糊）。
- 圆形切割涟漪：小圆 delay 小、大圆 delay 大，从中心向外扩散。
- **画幅自适应**：预览和剧本不能假设 16:9；先读 `p:sldSz`，方图、竖版和横版都按原比例输出。
- **背景图要能被动画**：优先把背景图作为铺满页面的 `p:pic`，不要让预览器只读静态 `p:bg`；
  同时给幻灯片底层设置深色纯色填充，避免入场首帧出现白闪。
- **层次感**：用「主从重叠」而非逐个排队——背景慢(1.2~1.5s)、前景快(0.5~0.8s)、
  前景在背景进行到 60% 时重叠开始；同组元素间隔 80~120ms 形成波浪。
  光晕圆入场后追加 `breathe` 强调（同一 spid 出现两次），整页持续微呼吸。
- **电影感圆形**：大圆环 `zoom_soft` 入场后接 `rotate` 慢转；背景图 `fade` +
  `kenburns`；巨号英文数字用 `cinematic`；标题用 `zoom_turn`，正文 `float`/`fade` 错峰。
- **Morph 延续**：相邻页把同一圆环命名为 `camera_ring`，各页用
  `"morph": [{"spid": "...", "name": "camera_ring"}]`，转场会随位置/尺寸变化连续变形。
  没有同名核心时，用 `iris`（中心圆）或 `wheel`（轮盘）更稳。
- **别让同一个圆在同一个位置来回**：左右交替只解决"永远在右边"，久了会变成
  "同一个圆、同一个大小、来回横移"。真正有效的是每页给一个不同的 **(圆心, 直径)**：
  全高压边 428 / 大裁切 540 / 小孔 330 / 偏心大圆 480 / 铺满 1080 交替排。
  检查两个数：**位移 ≥ 0.6 画布宽**、**直径比在 0.8~1.6 之间** —— 两个都变才像运镜，
  只有一个变就是平移。详见 `references/anim_guide.md` 的「走位」。
- **只有一个圆 = 还是单调，要再加一层"装饰语言"**。圆窗是第一层（主体），页边细节是第二层。
  五件套：① 虚焦光斑 `blob`（角落柔光，压住大片死黑）② 对焦刻度 `tick`（沿圆周一小段弧
  排 21 根径向短刻度，每 5 根一根长的，像镜头对焦环）③ 四角取景标 `cnr`（16pt 直角细线，
  把版面"框住"）④ 脊线竖排字 `vlabel`（贴左右边缘、转 270°，杂志刊头味）⑤ 章节进度轨
  `rail`（4 段短杠 + `0N / 04`，给读者一个"我在第几章"的锚点）。
  三条实现纪律：（a）**装饰必须追加在每页内容之后**，否则已有形状 id 全部后移，
  `plans.json` 里手写的 spid 会整体错位；（b）装饰**不进** `carry` 内核 —— 它们在圆外，
  跟着 `carry` 的 `scale` 走会飞散，正确做法是接力结束后再 fade/wipe 出来；
  （c）刻度环放不下就**不放**（出血页 d=540 时圆周外没有空位），每页装饰密度不同反而更好。
- **`spc`（字距）会吃字**：给 7.5pt 的小标签设 `spc` 后，LibreOffice 会在组边界断行、
  然后把第二行裁掉 —— 表现是"2026"变"202"、"参赛手册"变"参赛手"，像字体缺字，
  其实是布局。**细小标签不要设 `spc`**，要大字距就重写字符串（插全角空格）。
- **别用 `a:gradFill path="circle"` 画柔光斑**：LibreOffice 会在形状边界留一圈硬边
  （看着像块"饼"），PowerPoint 里却正常，两端不一致最难查。改用**预生成的径向渐变 PNG**
  （圆心不透明、按 `(1-t)^2.4` 衰减到全透明）+ `a:alphaModFix` 压强度，
  任何渲染器都一致；python-pptx 按哈希去重，十页共用同一个 image part。
- **后置装饰循环里不要用全局版式变量**：`set_side()` 留下的 `CX/CY/D` 在装饰循环执行时
  已经被**最后一页**覆盖（铺满页 480/270/1080），刻度环会全部算到画布外、静默跳过。
  必须逐页从 `coreImg` 现读圆心与直径。
- **主体只有一层，视觉重量就是恒定的**：单个圆再怎么位移、缩放，停下来看还是"一个图形"。
  要冲击力就得叠层，四层结构（自下而上）：
  ① `spot` 同色实心圆，9% 透明，偏移 60~120pt，"后面还有东西"；
  ② 主圆（照片 + 双环 + 内弧），跨页 carry 的就是它；
  ③ `moon` 伴圆，**实心深色盘**，0.36~0.52×，圆心落在主圆边缘（距离 ≈ 0.95r）；
  ④ `orbit` 0.8pt 偏心细环，1.10~1.20×，打破主圆轮廓。
  **伴圆一定要实心**：两个淡色照片圆叠在一起只会糊成"肥皂泡"，没有明暗差就没有重叠感。
  动效上分两拍：主圆 carry 到位，伴圆再 `iris` 开光圈、细轨 `zoom`、数字 `zoom_turn`，
  一次全出来就又回到"没有层次"。
- **浅底（米白 / 白）不是把深底反过来就行**，有四处必须改：
  ① **卡片必须描边**：纸色卡片和纸底只差几个色阶，只靠填充等于"看不见的卡片"，
  加 0.75pt 的 `LINE` 细描边才立得住；② **柔光斑要换颜色**（深底用暖金，浅底用低饱和绿），
  强度压到 12~14%，否则会像一块脏印；③ **圆内弧线放到 0.90d 以内**，放外侧会在 0° 方向
  被渲染成一条贴边的直线，正好戳进文字栏；④ **预览器底色不能写死**（见 `anim_guide.md`
  预览器第 5 条），浅底片跑出来会有一块黑斑。
- **大圆里的巨号数字要按直径缩放**（`字号 = 118 × d/404`）：写死字号的话，
  大裁切页显得小、小孔页又溢出。
- **固定锚点＝无缝切换**（"镜头不动、景物在变"）：全片在**同一坐标同一尺寸**放一个圆形
  影像窗，相邻页都用同名 `coreImg`/`coreRing`，转场统一 `morph` + `p14:dur` 1600~2000ms。
  同名同位的形状在 Morph 中恒等不动，圆内画面交叉淡入 → 观感是镜头固定、景象流转。
  **锚点尺寸一旦被改动就会变成推拉**，所以只在封面（略大）和结尾（放大铺满）刻意改。
  ⚠️ 这条依赖 morph：播放器不认 morph 时退化成硬切。要"一定看得见"，见下面 `carry`。
- **嫌"一直在右边"就左右交替**（这才是"多样式转换"）：锚点永远待在同一个位置，久了必然单调。
  按奇偶页把圆心在 `792`（右，出血右 46）与 `168`（左，出血左 46）之间交替，
  文字栏严格镜像（`TX=62` ↔ `TX=440`）。相邻页 `coreImg` 同处一个锚点、位置相反，
  翻页时**整圆横穿画面滑过去**（实测位移 594~636 pt）。
  要点：① 圆窗内核（图片 + 金环 + 弧线 + 环/点）**全部**用 `carry` 接力入场，
  起点＝上一页圆心、终点＝本页圆心，位移以画布比例给 `dx/dy`（1.0 = 一个画布宽）；
  ② 内核里原有的 `zoom_turn`/`wipe_r`/`spin` 一律**删掉**，否则会"一边滑一边从 0 放大"，
  只剩 `sway*` 这类强调保留（delay 放到接力之后）；③ 弧线加一个 `turn`（只转不缩），边滑边拧；
  ④ 圆内的文字（大号数字、"CONTENTS"）延迟到接力结束后再浮现；⑤ 转场用 `cut`，不要用 fade/zoom。
- **`p14:morph` 不能作为唯一手段**：它只有 PowerPoint 2019/365 认。
  WPS、PowerPoint 2016 及更早版本会**直接退化成硬切** —— 用户看到的就是
  "圆瞬间跳到左边，没有过程"。所以任何"跨页延续"的效果都必须有一份**动画层实现**
  （`carry` / `fly_*` / `exit_fly_*`），morph 只能当锦上添花。
- **接力位移的坐标单位**：`p:animMotion` 相对路径里 1.0 = 整个画布宽（`M -1 0 L 0 0`
  就是从画布左外侧飞入），所以 `dx = (上一页圆心x - 本页圆心x) / 960`。
  y 用画布高做分母。
- **圆窗图片禁止缩放/位移**：图片按 `prstGeom prst="ellipse"` 裁成圆形，任何
  `kenburns`/`pulse` 都会让"图片圆"与金环错位。要动就让**圆内结构**（弧线、同心环、数字）动。
- **填充透明度必须手写 XML**：python-pptx 的 `fill.transparency` 是**静默无效**的（不报错、
  也不写入），必须往 `<a:srgbClr>` 里插 `<a:alpha val="透明度×1000"/>`。这是"半透明全变实心"
  的根因，排查时先看 XML 有没有 alpha 节点。
- **图片裁圆的暗角陷阱**：圆形只取正方形图片的**内切圆区域**，把暗角画在正方形边缘等于白画
  （圆内仍是最亮的中心）。暗角必须**从圆心算起**，焦点也要推到圆外，圆内才有一段由亮到暗的过渡。
- **扇形/弧线用 `a:prstGeom prst="arc"` 不要用 `pie`**：pie 会带出两条穿过圆心的半径线，
  看起来像被划了两刀；arc 只有弧本身。角度单位是 1/60000 度。
- **章节页仪式感**：`wheel6`/`wheel8` + `transition_ms` 1200~1400，复刻参考片的光圈扫过；
  章节页的**弧线段数＝章节号**（ONE 一弧、TWO 两弧、THREE 三弧），避免"随手放个形状"的观感。
- **嫌动画"太平"就换旋转/缩放语言**：纯 morph 只换画面、不产生位移运动。要三层叠加——
  转场 `zoom`（整页缩放翻入）/ `wheel8`（旋转扫过）、入场 `zoom_turn`（-8°转正）/ `cinematic`
  （微旋+推近）/ `spin`（转满360°）、持续 `sway6/10/15/20`（往复摆，即"转过去又摆回来"）。
  详见 `references/anim_guide.md` 的「旋转/缩放语言」。
- **`sway` 的 delay 必须晚于入场结束**（入场 delay+dur+200ms 起），否则被入场覆盖看不到摆动；
  幅度别超 15°，会歪。`arc` 形状的包围盒以圆心为中心，所以 sway 就是绕圆心转（光圈旋转）。
- **做"锚点横滑"预览时的三个坑**（`render_deck.py` 已修，改代码别改回去）：
  ① 锚点出血到画布外时，`img.crop()` 越界会被填黑、或裁成透明，圆滑进画面中间就露出一块
  "洞"，洞后面是下一页内容，叠成一道生硬竖边 → 必须用最近像素**拉伸补边**；
  ② 两页圆窗做交叉溶解时**只能改 alpha**（`alpha_composite` + 只压 alpha），
  直接 `Image.blend` 两个 RGBA 会把透明区的黑混进颜色，圆边缘一圈脏；
  ③ 文字层和圆窗的淡入**必须解耦**（文字 2.8× 速提前淡完，圆窗全时长滑行），
  同速淡会全程双曝光，像重影而不是转场。
- 每次应用动画后都要跑 `fix_names.py`，否则播放器可能打不开。
- 需要严格兼容性时可以再用微软 Open XML SDK 的 `OpenXmlValidator` 校验
  `Office2013` 与 `Microsoft365` 两个目标；示例产物已做到 0 errors。
- 支持 32 种页内动画（25 入场 + 7 强调循环）和 38 种页面转场，
  详见 `references/anim_guide.md`（务必先读）。
