# 科研流程图规范与模板（实验流程图 / 技术路线图 / 研究方案图）

> 面向论文与开题/结题报告的 Visio 制图规范。基础 API（连接、模具、导出）见 `reference/visio-pywin32.md`，本文件只讲科研场景的版式、样式与完整模板。

## 1. 科研流程图类型与结构选择

| 类型 | 典型用途 | 推荐结构 | 页面方向 |
| --- | --- | --- | --- |
| 实验流程图 | 论文 Methods 配图：实验操作顺序 | 自上而下单列 + 判断分支（质控点） | A4 纵向 |
| 技术路线图 | 开题报告/基金申请：研究总体方案 | 三栏式（研究内容→研究方法→预期成果）自左向右 | A4/A3 横向 |
| 研究方案图 | 学位论文绪论：多模块关系 | 分层分区（阶段色块分组） | 按内容 |

**通用原则：**
- 一张图讲清一件事；阶段/模块用颜色区分，同阶段同色。
- 判断分支必须有"是/否"标签；回退箭头（如质控不通过返回重做）走侧边，不与主流交叉。
- 中文论文图内文字用宋体/黑体/微软雅黑；字号 10.5 pt（五号）起，标题 12-14 pt。
- 拼写、单位、专业术语与正文一致（投稿前逐字核对）。

## 2. 版式与样式规范

### 2.1 页面设置（先固定，防"自动适应"乱尺寸）

```python
def setup_page(page, w_mm=297, h_mm=210):      # 默认 A4 横向
    ps = page.PageSheet
    ps.CellsU("PageWidth").FormulaU = "%d mm" % w_mm
    ps.CellsU("PageHeight").FormulaU = "%d mm" % h_mm
```

- 实验流程图：`setup_page(page, 210, 297)`（纵向）。
- 内容较多（多组、多分支）：A3 横向 `setup_page(page, 420, 297)`，投稿时按栏宽缩放仍清晰。

### 2.2 科研配色方案（填充/边框成对，柔和不抢字）

```python
STYLE = {                     # (FillForegnd, LineColor)
    "blue":   ("RGB(218,232,252)", "RGB(108,142,191)"),   # 数据/分析阶段
    "green":  ("RGB(213,232,212)", "RGB(130,179,102)"),   # 样本/材料阶段
    "orange": ("RGB(255,230,204)", "RGB(215,155,0)"),     # 实验/处理阶段
    "red":    ("RGB(248,206,204)", "RGB(184,84,80)"),     # 质控/风险点
    "purple": ("RGB(225,213,231)", "RGB(150,115,166)"),   # 成果/输出
    "yellow": ("RGB(255,242,204)", "RGB(214,182,86)"),    # 强调/里程碑
    "gray":   ("RGB(245,245,245)", "RGB(102,102,102)"),   # 起止/注释
}
```

### 2.3 论文用导出规范

- `page.Export("x.png")` 位图约 **96 dpi**，直接投稿分辨率不足。论文配图用：
  - **SVG（首选）**：`page.Export(r"D:\out\fig.svg")`，矢量无限清晰，可插入 Word/LaTeX 后再转 PDF。
  - **PDF**：`doc.ExportAsFixedFormat(1, r"D:\out\fig.pdf", 1, 0)`（参数见 visio-pywin32.md §7）。
- 导出前把形状铺满页面边距内（四周留 ≥10 mm），避免"自动适应内容"裁出怪比例。
- 验证：额外导出 PNG 用 Read 回看，检查文字重叠与连线走向。

## 3. 可复用 Python 辅助函数库（复制到脚本开头即可）

```python
# -*- coding: utf-8 -*-
import sys, time, pythoncom, win32com.client
for s in (sys.stdout, sys.stderr):
    try: s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass

def connect():
    try: return win32com.client.GetActiveObject("Visio.Application"), False
    except pythoncom.com_error: pass
    for i in range(4):
        try: return win32com.client.Dispatch("Visio.Application"), True
        except pythoncom.com_error as e:
            if e.hresult == -2146959355 and i < 3: time.sleep(5); continue
            raise

STYLE = {  # §2.2 的配色字典
    "blue":   ("RGB(218,232,252)", "RGB(108,142,191)"),
    "green":  ("RGB(213,232,212)", "RGB(130,179,102)"),
    "orange": ("RGB(255,230,204)", "RGB(215,155,0)"),
    "red":    ("RGB(248,206,204)", "RGB(184,84,80)"),
    "purple": ("RGB(225,213,231)", "RGB(150,115,166)"),
    "yellow": ("RGB(255,242,204)", "RGB(214,182,86)"),
    "gray":   ("RGB(245,245,245)", "RGB(102,102,102)"),
}

def style(shp, color="blue", font='"Microsoft YaHei"', size="10.5 pt",
          bold=False, font_color="RGB(51,51,51)", lw="1 pt", rounded=None):
    fill, line = STYLE[color]
    shp.CellsU("FillForegnd").FormulaU = fill
    shp.CellsU("LineColor").FormulaU = line
    shp.CellsU("LineWeight").FormulaU = lw
    shp.CellsU("Char.Font").FormulaU = font        # 注意带内层双引号
    shp.CellsU("Char.Size").FormulaU = size
    shp.CellsU("Char.Color").FormulaU = font_color
    if bold:   shp.CellsU("Char.Style").FormulaU = "1"      # 1=加粗 2=斜体 4=下划线
    if rounded: shp.CellsU("Rounding").FormulaU = rounded   # 如 "0.08 in" 圆角
    return shp

def box(page, cx, cy, w, h, text, color="blue", **kw):
    """以 (cx,cy) 为中心画矩形。单位: 英寸。"""
    shp = page.DrawRectangle(cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2)
    shp.Text = text
    return style(shp, color, **kw)

def flow_shape(page, stencil, name, cx, cy, text, color="blue", **kw):
    """从模具落流程图符号(Process/Decision/Start/End...)。"""
    shp = page.Drop(stencil.Masters.ItemU(name), cx, cy)
    shp.Text = text
    return style(shp, color, **kw)

def arrow(app, page, a, b, label="", dashed=False, color="RGB(102,102,102)"):
    """a→b 动态连接线, 实心箭头, 可带标签(是/否/n=...)。"""
    c = page.Drop(app.ConnectorToolDataObject, 0, 0)
    c.CellsU("BeginX").GlueTo(a.CellsU("PinX"))
    c.CellsU("EndX").GlueTo(b.CellsU("PinX"))
    c.CellsU("EndArrow").FormulaU = "4"
    c.CellsU("LineColor").FormulaU = color
    c.CellsU("LineWeight").FormulaU = "1 pt"
    if dashed: c.CellsU("LinePattern").FormulaU = "2"
    if label:
        c.Text = label
        c.CellsU("Char.Font").FormulaU = '"Microsoft YaHei"'
        c.CellsU("Char.Size").FormulaU = "9 pt"
    return c

def setup_page(page, w_mm=297, h_mm=210):
    ps = page.PageSheet
    ps.CellsU("PageWidth").FormulaU = "%d mm" % w_mm
    ps.CellsU("PageHeight").FormulaU = "%d mm" % h_mm
```

- 坐标单位英寸（1 in = 25.4 mm）；`box(page, cx, cy, ...)` 传**形状中心**，与 `page.Drop` 一致。
- 中文宋体论文风：`font='"SimSun"'`；开题答辩投影风：`'"Microsoft YaHei"'`。

## 4. 完整示例一：实验流程图（纵向 + 质控分支）

结构：开始 → 样本采集 → 样本处理 → 质控判断（不合格→返回处理）→ 分组实验 → 数据采集 → 数据分析 → 结论。阶段配色：样本=绿、实验=橙、质控=红、数据=蓝、成果=紫。

```python
# -*- coding: utf-8 -*-
# 粘贴 §3 辅助函数库, 然后:
app, started = connect()
app.Visible = True
doc = app.Documents.Add("")
page = app.ActivePage
setup_page(page, 210, 297)                       # A4 纵向 (8.27 x 11.69 in)
st = app.Documents.OpenEx(
    r"C:\Program Files\Microsoft Office\Root\Office16\Visio Content\2052\BASFLO_M.VSSX",
    2 + 64)

CX = 4.16                                        # 页面水平中心(英寸)
def row(y): return y

start  = flow_shape(page, st, "Start/End", CX, row(10.8), "开始", "gray", size="12 pt", bold=True)
s1     = flow_shape(page, st, "Process",   CX, row(9.5),  "样本采集与登记\n(n = 60)", "green")
s2     = flow_shape(page, st, "Process",   CX, row(8.2),  "样本预处理\n(匀浆 / 离心 / 分装)", "green")
qc     = flow_shape(page, st, "Decision",  CX, row(6.7),  "质控合格?\n(浓度/纯度检测)", "red")
e1     = flow_shape(page, st, "Process",   CX - 1.9, row(5.0), "对照组\n(生理盐水)", "orange")
e2     = flow_shape(page, st, "Process",   CX,      row(5.0), "低剂量组\n(10 mg/kg)", "orange")
e3     = flow_shape(page, st, "Process",   CX + 1.9, row(5.0), "高剂量组\n(50 mg/kg)", "orange")
d1     = flow_shape(page, st, "Data",      CX, row(3.5),  "指标检测\n(ELISA / qPCR / WB)", "blue")
d2     = flow_shape(page, st, "Process",   CX, row(2.2),  "统计分析\n(ANOVA, n = 3)", "blue")
concl  = flow_shape(page, st, "Process",   CX, row(0.9),  "结果与结论", "purple", bold=True)

arrow(app, page, start, s1)
arrow(app, page, s1, s2)
arrow(app, page, s2, qc)
arrow(app, page, qc, e2, "合格")
arrow(app, page, qc, e1, "")
arrow(app, page, e2, e3, "")
arrow(app, page, e1, d1); arrow(app, page, e2, d1); arrow(app, page, e3, d1)
arrow(app, page, d1, d2)
arrow(app, page, d2, concl)
arrow(app, page, qc, s2, "不合格 → 重新处理", dashed=True)   # 回退线(动态连接会自动绕行)

out = r"<临时目录>\exp_flow.png"
page.Export(out); print("已导出:", out, flush=True)
page.Export(r"<输出目录>\exp_flow.svg")          # 论文用矢量版

doc.SaveAs(r"<输出目录>\exp_flow.vsdx")
app.AlertResponse = 7
doc.Close()
if started: app.Quit()
```

**版式要点**：主流自上而下居中；分支左右对称展开；回退虚线让 Visio 自动绕行（`Layout()` 慎用，坐标手工排最可控）；每阶段一种颜色并在图注中说明。

## 5. 完整示例二：三栏式技术路线图（横向，开题/基金）

结构：顶部研究题目 → 三栏（研究内容 | 研究方法 | 预期成果）→ 底部总体目标。自左向右箭头。

```python
# -*- coding: utf-8 -*-
# 粘贴 §3 辅助函数库, 然后:
app, started = connect()
app.Visible = True
doc = app.Documents.Add("")
page = app.ActivePage
setup_page(page, 297, 210)                       # A4 横向 (11.69 x 8.27 in)

# 顶部标题条
title = box(page, 5.85, 7.7, 9.5, 0.7, "XXX 机制研究技术路线", "gray",
            size="14 pt", bold=True)

# 三栏表头
cols = [(2.0, "研究内容", "blue"), (5.85, "研究方法", "green"), (9.7, "预期成果", "purple")]
for cx, txt, color in cols:
    box(page, cx, 6.6, 2.9, 0.55, txt, color, size="12 pt", bold=True, rounded="0.08 in")

# 三行内容(按课题实际替换): (研究内容, 研究方法, 预期成果)
rows_y = [5.4, 4.1, 2.8]
contents = [
    ("现状调研与\n问题界定", "文献计量+实地调研", "研究框架\n与指标体系"),
    ("机制解析与\n模型构建", "实验+数值模拟", "因果机制\n解释模型"),
    ("方案优化与\n实证检验", "对照实验+灵敏度分析", "优化方案\n与政策建议"),
]
prev = None
for (t1, t2, t3), y in zip(contents, rows_y):
    c1 = box(page, 2.0,  y, 2.9, 1.0, t1, "blue")
    c2 = box(page, 5.85, y, 2.9, 1.0, t2, "green")
    c3 = box(page, 9.7,  y, 2.9, 1.0, t3, "purple")
    arrow(app, page, c1, c2)
    arrow(app, page, c2, c3)
    if prev:                                      # 纵向递进: 上一行成果 → 下一行内容(虚线)
        arrow(app, page, prev[2], c1, dashed=True)
    prev = (c1, c2, c3)

# 底部总体目标: 最后一行成果汇入
goal = box(page, 5.85, 1.4, 9.5, 0.7, "总体目标: 阐明 XX 规律并提出 XX 方案", "yellow",
           size="12 pt", bold=True)
arrow(app, page, prev[2], goal)

out = r"<临时目录>\tech_route.png"
page.Export(out); print("已导出:", out, flush=True)
page.Export(r"<输出目录>\tech_route.svg")         # 论文/申请书用矢量版
doc.SaveAs(r"<输出目录>\tech_route.vsdx")

app.AlertResponse = 7
doc.Close()
if started: app.Quit()
```

**版式要点**：
- 栏间距 ≥ 0.8 in 保证箭头标签不挤；同行三框**y 坐标严格一致**（对齐是技术路线图的生命线）。
- 文字两行以内，多了改"关键词+图注解释"。
- A3 横向（`setup_page(page, 420, 297)`）适合内容 ≥4 行或每格文字较多的场景。

## 6. 科研流程图检查清单

- [ ] 页面尺寸已固定（A4/A3），导出比例稳定
- [ ] 阶段配色统一（同阶段同色），颜色数量 ≤ 5
- [ ] 判断分支有"是/否"标签；回退线为虚线且不与主流交叉
- [ ] 同排/同列元素坐标严格对齐
- [ ] 图内字号 ≥ 10.5 pt（投影汇报 ≥ 14 pt），字体统一
- [ ] 术语、样本量、单位与正文完全一致
- [ ] 论文配图导出 SVG/PDF 矢量版；PNG 仅作验证回看
- [ ] 导出 PNG 已 Read 回看验证（布局/重叠/连线）

## 7. 常见科研流程图问题速查

| 问题 | 对策 |
| --- | --- |
| 连线穿框/绕远 | 手工坐标布局，不用 `Layout()`；回退线单独走侧边通道 |
| 导出 PNG 太小/模糊 | 96dpi 限制：论文用 SVG/PDF；或 A3 页面+矢量导出后按栏宽缩放 |
| 框内文字溢出 | 缩短文字或加大框（宽 +0.3 in）；"关键词化"而非整句 |
| 中英文混排字体乱 | 统一 `Char.Font`；中文 '"Microsoft YaHei"' / '"SimSun"'，勿混用 |
| 多页报告只需某页 | `doc.Pages.Item(i).Export(...)` 按页导出 |
