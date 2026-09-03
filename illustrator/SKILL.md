---
name: illustrator
description: 用本机已安装的 Adobe Illustrator 自动绘制矢量图。当用户要求用 Illustrator 制作 Logo、图标、信息图、矢量插画、.ai/.svg 文件，或制作科研论文插图/期刊投稿图（多面板组合图、数据图表、误差棒、显著性标注、机理示意图、300dpi/CMYK/期刊栏宽导出）时触发。通过 COM/JSX(ExtendScript) 脚本驱动真实软件出图，而非手写 SVG。
---

# Illustrator 自动绘图（含科研论文插图规范）

在 Windows 上生成 ExtendScript(.jsx) 脚本，经 COM `DoJavaScript` 驱动**真实的** Adobe Illustrator 完成制图任务。

本 skill 目录记为 `<SKILL_DIR>`（即本 SKILL.md 所在目录）。

## 前置条件

Windows + 已安装 Illustrator + Python + pywin32。首次使用或出错时先运行自检：

```powershell
[Console]::OutputEncoding=[Text.Encoding]::UTF8; python -X utf8 <SKILL_DIR>/scripts/check_env.py
```

它会报告：COM 是否可用、版本号、中文字体清单。**字体名以自检输出为准，不要凭记忆猜字体名。**

## 任务路由

| 任务 | 先读 |
| --- | --- |
| 期刊论文插图：多面板组合图、数据图表（柱状/折线/散点）、误差棒、显著性标注、坐标轴规范、300dpi/CMYK 导出 | `reference/scientific-figures.md`，再查 `reference/illustrator-jsx.md` |
| 机理示意图、实验装置图、图形摘要（Graphical Abstract） | `reference/scientific-figures.md` §5-§6（配色/辅助函数）+ `reference/illustrator-jsx.md` |
| Logo、图标、矢量插画、信息图 | `reference/illustrator-jsx.md` |

## 统一工作流（必须遵守）

1. **规划**：先在脑中/草稿确定画布尺寸、每个元素的坐标与层级，再写脚本。不要边画边想。
   - 科研插图：先确定**目标期刊栏宽**（画板宽度）与面板排布，再定坐标。见 `reference/scientific-figures.md` §1。
2. **写脚本**：临时脚本写入系统临时目录（如 `$env:TEMP` 下新建子目录），**不要污染用户项目目录**。
3. **执行**：用 `run_jsx.py` 执行 .jsx。
4. **导出验证图**：总是额外导出一份 PNG。
5. **验证**：若当前模型支持读图，用 Read 工具回看 PNG 检查布局/文字/重叠；不支持则用像素采样（System.Drawing）或结构化断言（对象坐标、文档尺寸）验证。发现问题→改脚本→重跑。
6. **清理**：任务结束后删除临时脚本与临时导出文件；不要留下无人看管的对话框。

## PowerShell 中文输出规则（Windows 5.1 管道乱码）

在 PowerShell 中运行 Python 必须前缀 UTF-8 设置，否则中文输出乱码：

```powershell
[Console]::OutputEncoding=[Text.Encoding]::UTF8; python -X utf8 script.py
```

## 快速开始

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
- 科研插图三原则：**先定栏宽（画板）→ 字号/线宽按最终印刷尺寸设定 → 位图导出按期刊 dpi 换算缩放百分比**。详见 `reference/scientific-figures.md`。
- 更多配方（贝塞尔路径、渐变、图层/编组、SVG/PDF 导出、多画板、完整 Logo 与图表示例）→ `reference/illustrator-jsx.md`。

## 排错速查

| 现象 | 原因与对策 |
| --- | --- |
| `(-2146959355, '服务器运行失败')` | 应用冷启动中 COM 超时；run_jsx.py 已自动重试，仍失败则手动先开一次 Illustrator |
| 脚本卡住无任何输出 | 99% 是弹窗；任务管理器结束 Illustrator 进程后修复脚本重跑 |
| 控制台中文乱码 | 漏了 `[Console]::OutputEncoding=...UTF8` 前缀 |
| 字体回退/方块字 | `getByName` 失败被吞掉；先跑 check_env.py 看真实字体名（本机常见 MicrosoftYaHei、FangSong、KaiTi、NotoSansSC 等） |
| JSX 第一行莫名报错 | .jsx 文件带 BOM；run_jsx.py 已用 utf-8-sig 处理，若自己写执行器需注意 |
| 投稿被退"分辨率不足" | 导出 dpi 不够：见 `reference/scientific-figures.md` §3 的 dpi→缩放百分比换算表 |

## 文件清单

- `scripts/check_env.py` — 环境自检（COM/版本/中文字体）
- `scripts/run_jsx.py` — JSX 执行器（DoJavaScript，支持中文与行号报错）
- `reference/illustrator-jsx.md` — Illustrator JSX 完整配方
- `reference/scientific-figures.md` — 科研论文插图规范与模板（栏宽/字号/dpi/配色/辅助函数/多面板示例）
