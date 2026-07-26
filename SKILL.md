---
name: illustrator-visio
description: 用本机已安装的 Adobe Illustrator 和 Microsoft Visio 自动绘图。当用户要求制作 Visio 流程图、架构图、网络拓扑图、组织结构图、.vsdx 文件，或要求用 Illustrator 制作矢量图、Logo、图标、信息图、矢量插画、.ai/.svg 文件时触发。通过 COM/JSX 脚本驱动真实软件出图，而非手写 SVG。
---

# Illustrator / Visio 自动绘图

在 Windows 上用脚本驱动**真实的** Adobe Illustrator 与 Microsoft Visio 完成制图任务。

- **Visio 路线**：Python + pywin32 调用 Visio COM（详见 `reference/visio-pywin32.md`）
- **Illustrator 路线**：生成 ExtendScript(.jsx) 脚本，经 COM `DoJavaScript` 执行（详见 `reference/illustrator-jsx.md`）

本 skill 目录记为 `<SKILL_DIR>`（即本 SKILL.md 所在目录）。

## 前置条件

Windows + 已安装 Illustrator / Visio + Python + pywin32。首次使用或出错时先运行自检：

```powershell
[Console]::OutputEncoding=[Text.Encoding]::UTF8; python -X utf8 <SKILL_DIR>/scripts/check_env.py
```

它会报告：两个 COM 是否可用、版本号、Illustrator 中文字体、Visio 模具（stencil）真实路径与完整清单。**模具清单以自检输出为准，不要凭记忆猜模具文件名。**

## 工具选择

| 任务 | 工具 |
| --- | --- |
| 流程图、业务流程、判断分支 | **Visio**（BASFLO_M 模具 + 自动连接线） |
| 系统架构、网络拓扑、机柜图 | **Visio**（NETSYM/AZURE/AWS 等专用模具） |
| 组织结构图、泳道图、BPMN | **Visio**（ORGHCH/XFUNC/BPMN 模具） |
| Logo、图标、矢量插画 | **Illustrator**（路径/渐变/排版控制最强） |
| 精致信息图、矢量数据图表 | **Illustrator**（像素级控制） |
| 最终交付 .vsdx | Visio；交付 .ai/.svg/.pdf 且重设计感 | Illustrator |

## 统一工作流（必须遵守）

1. **规划**：先在脑中/草稿确定画布尺寸、每个元素的坐标与层级，再写脚本。不要边画边想。
2. **写脚本**：临时脚本写入系统临时目录（如 `$env:TEMP` 下新建子目录），**不要污染用户项目目录**。
3. **执行**：Visio 直接跑 Python；Illustrator 用 `run_jsx.py` 执行 .jsx。
4. **导出验证图**：总是额外导出一份 PNG。
5. **验证**：若当前模型支持读图，用 Read 工具回看 PNG 检查布局/文字/重叠；不支持则用像素采样（System.Drawing）或结构化断言（对象坐标、文档尺寸）验证。发现问题→改脚本→重跑。
6. **清理**：任务结束后删除临时脚本与临时导出文件；不要留下无人看管的对话框。

## PowerShell 中文输出规则（Windows 5.1 管道乱码）

在 PowerShell 中运行 Python 必须前缀 UTF-8 设置，否则中文输出乱码：

```powershell
[Console]::OutputEncoding=[Text.Encoding]::UTF8; python -X utf8 script.py
```

## Visio 快速开始

最小可用模板（已实测，Visio 16）：

```python
# -*- coding: utf-8 -*-
import sys, time, pythoncom, win32com.client

for s in (sys.stdout, sys.stderr):
    try: s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass

def connect(progid="Visio.Application"):
    try:
        return win32com.client.GetActiveObject(progid), False   # 附加到已运行实例
    except pythoncom.com_error:
        pass
    for i in range(4):                                          # 冷启动重试
        try:
            return win32com.client.Dispatch(progid), True
        except pythoncom.com_error as e:
            if e.hresult == -2146959355 and i < 3:              # 0x80080005 启动中
                time.sleep(5); continue
            raise

app, started = connect()
app.Visible = True
doc = app.Documents.Add("")                                     # 空白绘图
page = app.ActivePage
stencil = app.Documents.OpenEx(
    r"C:\Program Files\Microsoft Office\Root\Office16\Visio Content\2052\BASFLO_M.VSSX",
    2 + 64)                                                     # 只读 + 隐藏

s1 = page.Drop(stencil.Masters.ItemU("Process"), 4.25, 8.0)     # 单位: 英寸, 原点左下, y 向上
s1.Text = "步骤一"
s2 = page.Drop(stencil.Masters.ItemU("Process"), 4.25, 5.5)
s2.Text = "步骤二"
s2.CellsU("FillForegnd").FormulaU = "RGB(198,224,180)"          # 填充色

conn = page.Drop(app.ConnectorToolDataObject, 0, 0)             # 动态连接线(无需模具)
conn.CellsU("BeginX").GlueTo(s1.CellsU("PinX"))
conn.CellsU("EndX").GlueTo(s2.CellsU("PinX"))

page.Layout()                                                   # 自动布局(可选)
page.Export(r"<临时目录>\out.png")                              # 导出验证图
doc.SaveAs(r"<目标路径>\out.vsdx")                              # 需要交付时保存

app.AlertResponse = 7                                           # 关键! 自动回答"否", 防弹窗挂死
doc.Close()
if started: app.Quit()
```

**关键规则：**

- `AlertResponse = 7` 必须在 `doc.Close()` / `app.Quit()` 之前设置，否则"是否保存更改"弹窗会把脚本永久挂死（无界面运行时无法点击）。
- 只有自己启动的实例（`started=True`）才 `Quit()`；附加到用户已开的实例时只关闭自己创建的文档。
- 模具 master 名用 `ItemU`（通用名，跨语言稳定）：`"Process"`、`"Decision"`、`"Start/End"`、`"Document"`、`"Data"`。
- 更多配方（样式、箭头、其他模具、PDF 导出、页面尺寸、遍历 master 清单）→ `reference/visio-pywin32.md`。

## Illustrator 快速开始

1. 写 JSX（UTF-8 编码）：

```javascript
// 单位: pt; 原点: 画板左下, y 轴向上
var doc = app.documents.add(DocumentColorSpace.RGB, 800, 600);  // 画板 800x600

function rgb(r, g, b) { var c = new RGBColor(); c.red = r; c.green = g; c.blue = b; return c; }

var rect = doc.pathItems.rectangle(600, 0, 300, 300);           // rectangle(top, left, 宽, 高)
rect.fillColor = rgb(230, 40, 40); rect.stroked = false;        // → 画板左上角的红色方块

var tf = doc.textFrames.add();
tf.contents = "中文标题";
try { tf.textRange.characterAttributes.textFont = app.textFonts.getByName("MicrosoftYaHei"); } catch (e) {}
tf.textRange.characterAttributes.size = 48;
tf.position = [80, 120];

var opts = new ExportOptionsPNG24();                            // 导出验证图
opts.horizontalScale = 150; opts.verticalScale = 150;           // 百分比, 100%≈72dpi
opts.artBoardClipping = true; opts.transparency = true;
doc.exportFile(new File("<临时目录>/out.png"), ExportType.PNG24, opts);

doc.saveAs(new File("<目标路径>/out.ai"));                      // 需要交付时保存
doc.close(SaveOptions.DONOTSAVECHANGES);                        // 显式不保存关闭, 无弹窗
"done";                                                         // 最后的表达式会回传给执行器打印
```

2. 执行：

```powershell
[Console]::OutputEncoding=[Text.Encoding]::UTF8; python -X utf8 <SKILL_DIR>/scripts/run_jsx.py <脚本>.jsx [--quit]
```

**关键规则：**

- 坐标系：原点画板**左下**，y **向上**，单位 pt（1pt = 1/72 英寸）。`rectangle(top, left, w, h)` 的第一个参数是顶边 y 坐标。习惯"从上往下"布局时定义辅助函数 `function Y(t){ return H - t; }`。
- `run_jsx.py` 会附加到正在运行的 Illustrator（可连续多次执行迭代）；`--quit` 仅在实例由脚本启动时执行后退出。
- JSX 抛错时执行器返回非零退出码并输出 `Line: N->` 行号定位。
- 渐变填充必须用 `GradientColor` 包装（直接 `fillColor = gradient` 会报 Internal error）；见参考文档。
- 更多配方（贝塞尔路径、渐变、图层/编组、SVG/PDF 导出、多画板、完整 Logo 与图表示例）→ `reference/illustrator-jsx.md`。

## 排错速查

| 现象 | 原因与对策 |
| --- | --- |
| `(-2146959355, '服务器运行失败')` | 应用冷启动中 COM 超时；按模板重试 3-4 次（间隔 5 秒） |
| 脚本卡住无任何输出 | 99% 是弹窗。Visio：检查是否漏了 `AlertResponse = 7`；任务管理器结束对应进程后修复脚本重跑 |
| 控制台中文乱码 | 漏了 `[Console]::OutputEncoding=...UTF8` 前缀 |
| Illustrator 字体回退/方块字 | `getByName` 失败被吞掉；先跑 check_env.py 看真实字体名（本机有 MicrosoftYaHei、FangSong、KaiTi、NotoSansSC 等） |
| Visio 模具打不开 | 路径不存在；以 check_env.py 输出的模具目录为准（注意 2052/1033 语言目录） |
| JSX 第一行莫名报错 | .jsx 文件带 BOM；run_jsx.py 已用 utf-8-sig 处理，若自己写执行器需注意 |

## 文件清单

- `scripts/check_env.py` — 环境自检（COM/字体/模具清单）
- `scripts/run_jsx.py` — Illustrator JSX 执行器（DoJavaScript，支持中文与行号报错）
- `reference/visio-pywin32.md` — Visio 完整配方
- `reference/illustrator-jsx.md` — Illustrator JSX 完整配方
