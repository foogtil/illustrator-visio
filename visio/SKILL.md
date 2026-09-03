---
name: visio
description: 用本机已安装的 Microsoft Visio 自动绘图。当用户要求制作 Visio 流程图、实验流程图、技术路线图、研究方案图、架构图、网络拓扑图、组织结构图、泳道图、BPMN、.vsdx 文件时触发。通过 Python COM(pywin32) 脚本驱动真实软件出图。
---

# Visio 自动绘图（含科研流程图规范）

在 Windows 上用 Python + pywin32 调用 Visio COM（详见 `reference/visio-pywin32.md`）驱动**真实的** Microsoft Visio 完成制图任务。

本 skill 目录记为 `<SKILL_DIR>`（即本 SKILL.md 所在目录）。

## 前置条件

Windows + 已安装 Visio + Python + pywin32。首次使用或出错时先运行自检：

```powershell
[Console]::OutputEncoding=[Text.Encoding]::UTF8; python -X utf8 <SKILL_DIR>/scripts/check_env.py
```

它会报告：COM 是否可用、版本号、Visio 模具（stencil）真实路径与完整清单。**模具清单以自检输出为准，不要凭记忆猜模具文件名。**

## 任务路由

| 任务 | 先读 |
| --- | --- |
| 实验流程图、技术路线图、研究方案图（科研向） | `reference/research-flowcharts.md`，再查 `reference/visio-pywin32.md` |
| 业务流程、判断分支 | `reference/visio-pywin32.md`（BASFLO_M 模具） |
| 系统架构、网络拓扑、机柜图 | `reference/visio-pywin32.md`（NETSYM/AZURE/AWS 模具） |
| 组织结构图、泳道图、BPMN | `reference/visio-pywin32.md`（ORGHCH/XFUNC/BPMN 模具） |

## 统一工作流（必须遵守）

1. **规划**：先在脑中/草稿确定页面尺寸、每个元素的坐标与层级，再写脚本。不要边画边想。
   - 科研流程图：先确定**页面尺寸（A4 横向等）与阶段分区**，再定坐标。见 `reference/research-flowcharts.md` §1-§2。
2. **写脚本**：临时脚本写入系统临时目录（如 `$env:TEMP` 下新建子目录），**不要污染用户项目目录**。
3. **执行**：直接跑 Python。
4. **导出验证图**：总是额外导出一份 PNG。
5. **验证**：若当前模型支持读图，用 Read 工具回看 PNG 检查布局/文字/重叠；不支持则用像素采样（System.Drawing）或结构化断言（对象坐标、文档尺寸）验证。发现问题→改脚本→重跑。
6. **清理**：任务结束后删除临时脚本与临时导出文件；不要留下无人看管的对话框。

## PowerShell 中文输出规则（Windows 5.1 管道乱码）

在 PowerShell 中运行 Python 必须前缀 UTF-8 设置，否则中文输出乱码：

```powershell
[Console]::OutputEncoding=[Text.Encoding]::UTF8; python -X utf8 script.py
```

## 快速开始

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
- 科研流程图：**固定页面尺寸 + 阶段配色统一 + 论文用导出 SVG/PDF 矢量**。详见 `reference/research-flowcharts.md`。
- 更多配方（样式、箭头、其他模具、PDF 导出、页面尺寸、遍历 master 清单）→ `reference/visio-pywin32.md`。

## 排错速查

| 现象 | 原因与对策 |
| --- | --- |
| `(-2146959355, '服务器运行失败')` | 应用冷启动中 COM 超时；按模板重试 3-4 次（间隔 5 秒） |
| 脚本卡住无任何输出 | 99% 是弹窗。检查是否漏了 `AlertResponse = 7`；任务管理器结束 VISIO.EXE 后修复脚本重跑 |
| 控制台中文乱码 | 漏了 `[Console]::OutputEncoding=...UTF8` 前缀 |
| Visio 模具打不开 | 路径不存在；以 check_env.py 输出的模具目录为准（注意 2052/1033 语言目录） |
| 导出的 PNG 很小/模糊 | 页面"自动适应内容"+位图约 96dpi 所致：固定页面尺寸，论文用导 SVG/PDF（见 `reference/research-flowcharts.md` §2.3） |

## 文件清单

- `scripts/check_env.py` — 环境自检（COM/版本/模具清单）
- `reference/visio-pywin32.md` — Visio 完整配方
- `reference/research-flowcharts.md` — 科研流程图规范与模板（实验流程图/技术路线图/研究方案图）
