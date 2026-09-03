# Illustrator 完整配方（ExtendScript / JSX）

> 基于 Illustrator 30.0 实测。坐标、导出、渐变等"坑点"均已实际验证。

## 1. 执行方式

```powershell
[Console]::OutputEncoding=[Text.Encoding]::UTF8; python -X utf8 <SKILL_DIR>/scripts/run_jsx.py script.jsx [--quit]
```

- 通过 COM `DoJavaScript` 执行；**附加到正在运行的实例**，没有才启动新实例。
- 脚本最后一个表达式的值会打印到 stdout（可回传诊断信息）。
- 出错时退出码为 1，并输出 ExtendScript 行号，如：
  `Error 24: xxx 不是函数. Line: 12->  ...`
- 同一实例内多次执行会**共享应用状态**（上一次未关闭的文档还在），便于迭代；脚本开头建议先建好新文档，失败脚本残留的文档用 `doc.close(SaveOptions.DONOTSAVECHANGES)` 清理。
- `--quit`：仅当实例由本次执行启动时，结束后退出 Illustrator。

## 2. 文档、画板与坐标系（实测）

```javascript
var doc = app.documents.add(DocumentColorSpace.RGB, 800, 600);  // 宽800 高600 pt
// → artboardRect 自动为 [0, 600, 800, 0]  (left, top, right, bottom)
doc.artboards[0].artboardRect = [0, 1080, 1920, 0];             // 改画板尺寸
doc.artboards.add([0, 600, 800, 0]);                            // 加第二画板
```

- **原点 = 画板左下角，y 轴向上，单位 pt**（1 pt = 1/72 英寸；1 mm ≈ 2.8346 pt；1 英寸 = 72 pt）。
- 习惯"从上往下"排版时定义换算函数：

```javascript
var H = 600;                       // 画板高
function Y(top) { return H - top; } // 传入"距顶距离", 返回 JSX y 坐标
```

## 3. 形状（实测）

```javascript
// 三种基本形状, 参数均为 (top, left, width, height) —— 注意 top 在前!
var rect  = doc.pathItems.rectangle(600, 0, 300, 300);        // 顶边y=600(即画板顶) → 左上区域
var ell   = doc.pathItems.ellipse(400, 500, 200, 200);        // 正圆: 宽高相等
var rr    = doc.pathItems.roundedRectangle(400, 10, 120, 80, 12, 12); // 后两个参数: 水平/垂直圆角半径

rect.fillColor   = rgb(230, 40, 40);
rect.stroked     = false;
ell.filled       = false;
ell.stroked      = true;
ell.strokeColor  = rgb(40, 40, 40);
ell.strokeWidth  = 2;
```

自由多边形（实测）：

```javascript
var tri = doc.pathItems.add();
tri.setEntirePath([[200, 300], [260, 380], [320, 300]]);  // [[x,y],...]
tri.closed = true;
tri.fillColor = rgb(80, 80, 220);
tri.stroked = false;
```

贝塞尔曲线（标准 OMV API）：

```javascript
var p = doc.pathItems.add();
var pt = p.pathPoints.add();
pt.anchor = [100, 100];
pt.leftDirection  = [100, 160];   // 控制柄 = 绝对坐标
pt.rightDirection = [160, 100];
var pt2 = p.pathPoints.add();
pt2.anchor = [300, 100];
pt2.leftDirection = [240, 100];
pt2.rightDirection = [300, 40];
p.stroked = true; p.filled = false;
```

## 4. 颜色与渐变

```javascript
function rgb(r, g, b) { var c = new RGBColor(); c.red = r; c.green = g; c.blue = b; return c; }
function cmyk(c, m, y, k) { var x = new CMYKColor(); x.cyan=c; x.magenta=m; x.yellow=y; x.black=k; return x; }
```

**渐变（实测：必须经 GradientColor 包装）**：

```javascript
var g = doc.gradients.add();
g.type = GradientType.LINEAR;              // 或 GradientType.RADIAL
g.gradientStops[0].rampPoint = 0;          // 0~100
g.gradientStops[0].color = rgb(255, 0, 0);
g.gradientStops[1].rampPoint = 100;
g.gradientStops[1].color = rgb(0, 0, 255);
// 更多色标: var s = g.gradientStops.add(); s.rampPoint = 50; s.color = ...

var gc = new GradientColor();
gc.gradient = g;
rect.fillColor = gc;   // ✔ 正确
// rect.fillColor = g;   // ✘ 实测报 "Internal error"
// g.name = "x";         // ✘ 实测报 PARM 错误, 不要给 Gradient 命名
```

## 5. 文字（中文实测可用）

```javascript
var tf = doc.textFrames.add();          // 点文本
tf.contents = "冒烟测试中文";
tf.position = [80, 120];                 // 文本锚点(左下)
var ca = tf.textRange.characterAttributes;
ca.size = 48;
ca.fillColor = rgb(30, 30, 30);
try { ca.textFont = app.textFonts.getByName("MicrosoftYaHei"); }
catch (e) { /* 字体不存在时保留默认, 不要让整个脚本崩 */ }
tf.textRange.paragraphAttributes.justification = Justification.CENTER;
```

- 本机实测可用中文字体名：`MicrosoftYaHei`、`MicrosoftYaHei-Bold`、`FangSong`、`KaiTi`、`NotoSansSC-Regular/Bold/...`（以 check_env.py 输出为准）。
- 区域文本：`doc.textFrames.areaText(pathItem)`，先把矩形当容器。
- 科研图西文字体规范（Arial/Helvetica、字号/线宽）→ `reference/scientific-figures.md` §2-§3。

## 6. 图层、编组、变换

```javascript
var lay = doc.layers.add(); lay.name = "背景";      // 新图层置于顶层
var grp = doc.groupItems.add();                      // 编组
var c = doc.pathItems.ellipse(100, 100, 40, 40);
c.move(grp, ElementPlacement.PLACEATEND);            // 移入组

item.translate(50, -20);      // 平移 dx, dy
item.rotate(45);              // 旋转(度, 默认绕中心)
item.resize(150, 150);        // 缩放(百分比)
```

## 7. 导出与保存（均已实测）

```javascript
// PNG (位图): 缩放百分比控制分辨率, 100% ≈ 72dpi, 200% ≈ 144dpi, 417% ≈ 300dpi
var png = new ExportOptionsPNG24();
png.horizontalScale = 150; png.verticalScale = 150;
png.artBoardClipping = true;      // 裁剪到画板
png.transparency = true;          // 透明背景
doc.exportFile(new File("D:/out/x.png"), ExportType.PNG24, png);

// SVG (矢量)
doc.exportFile(new File("D:/out/x.svg"), ExportType.SVG, new ExportOptionsSVG());

// PDF
doc.saveAs(new File("D:/out/x.pdf"), new PDFSaveOptions());

// AI 源文件
doc.saveAs(new File("D:/out/x.ai"));

// 关闭(不弹窗)
doc.close(SaveOptions.DONOTSAVECHANGES);
```

文件路径在 JSX 中用 `/` 分隔（ExtendScript File 也接受 `\\`，推荐 `/`）。
**期刊投稿导出规范（300/600/1200dpi、EPS/TIFF、CMYK）→ `reference/scientific-figures.md` §3-§4。**

## 8. 完整示例一：简易 Logo

```javascript
var W = 600, H = 600;
var doc = app.documents.add(DocumentColorSpace.RGB, W, H);
function rgb(r, g, b) { var c = new RGBColor(); c.red = r; c.green = g; c.blue = b; return c; }

// 渐变圆底
var g = doc.gradients.add();
g.type = GradientType.LINEAR;
g.gradientStops[0].rampPoint = 0;  g.gradientStops[0].color = rgb(64, 120, 255);
g.gradientStops[1].rampPoint = 100; g.gradientStops[1].color = rgb(140, 60, 220);
var gc = new GradientColor(); gc.gradient = g;
var circle = doc.pathItems.ellipse(H - 100, 100, 400, 400);
circle.fillColor = gc; circle.stroked = false;

// 中心字母: position = 基线左端。圆心(300,300), 220pt 字母高约 158pt、宽约 158pt,
// 故基线 ≈ 300-158/2 ≈ 221, 左端 ≈ 300-158/2 ≈ 221。导出 PNG 后回看微调即可。
var tf = doc.textFrames.add();
tf.contents = "V";
try { tf.textRange.characterAttributes.textFont = app.textFonts.getByName("MicrosoftYaHei-Bold"); } catch (e) {}
tf.textRange.characterAttributes.size = 220;
tf.textRange.characterAttributes.fillColor = rgb(255, 255, 255);
tf.position = [220, 222];

var png = new ExportOptionsPNG24();
png.horizontalScale = 200; png.verticalScale = 200;
png.artBoardClipping = true; png.transparency = true;
doc.exportFile(new File("<临时目录>/logo.png"), ExportType.PNG24, png);
doc.saveAs(new File("<目标目录>/logo.ai"));
doc.close(SaveOptions.DONOTSAVECHANGES);
"logo done";
```

## 9. 完整示例二：柱状信息图（含从上往下布局辅助）

```javascript
var W = 1000, H = 700;
var doc = app.documents.add(DocumentColorSpace.RGB, W, H);
function rgb(r, g, b) { var c = new RGBColor(); c.red = r; c.green = g; c.blue = b; return c; }
function Y(t) { return H - t; }   // 距顶 → JSX y

var data = [["一月", 42], ["二月", 58], ["三月", 35], ["四月", 73], ["五月", 66]];
var maxV = 80, chartTop = 140, chartH = 420, barW = 90, gap = 70, left0 = 120;

// 标题
var title = doc.textFrames.add();
title.contents = "月度活跃度";
try { title.textRange.characterAttributes.textFont = app.textFonts.getByName("MicrosoftYaHei-Bold"); } catch (e) {}
title.textRange.characterAttributes.size = 40;
title.textRange.characterAttributes.fillColor = rgb(40, 40, 40);
title.position = [left0, Y(50)];

// 基线
var axis = doc.pathItems.add();
axis.setEntirePath([[left0 - 20, Y(chartTop)], [left0 - 20, Y(chartTop + chartH)], [W - 80, Y(chartTop + chartH)]]);
axis.stroked = true; axis.strokeColor = rgb(120, 120, 120); axis.strokeWidth = 2; axis.filled = false;

for (var i = 0; i < data.length; i++) {
    var h = data[i][1] / maxV * chartH;
    var x = left0 + i * (barW + gap);
    var bar = doc.pathItems.rectangle(Y(chartTop + chartH - h) + 0, x, barW, h); // 顶边 = 基线向上 h
    bar.fillColor = rgb(70 + i * 30, 130, 220); bar.stroked = false;

    var lb = doc.textFrames.add();
    lb.contents = data[i][0] + " " + data[i][1];
    try { lb.textRange.characterAttributes.textFont = app.textFonts.getByName("MicrosoftYaHei"); } catch (e) {}
    lb.textRange.characterAttributes.size = 20;
    lb.position = [x, Y(chartTop + chartH + 30)];
}

var png = new ExportOptionsPNG24();
png.horizontalScale = 150; png.verticalScale = 150;
png.artBoardClipping = true; png.transparency = false;
doc.exportFile(new File("<临时目录>/chart.png"), ExportType.PNG24, png);
doc.close(SaveOptions.DONOTSAVECHANGES);
"chart done, bars=" + data.length;
```

## 10. 排错

| 现象 | 对策 |
| --- | --- |
| `Internal error` 给 fillColor 赋渐变 | 必须 `new GradientColor()` 包装，见 §4 |
| `Error 1200 ... ('PARM')` 给渐变命名 | Gradient 不支持 `name` 赋值，删掉该行 |
| `0x80080005 服务器运行失败` | 冷启动中超时；run_jsx.py 已自动重试，仍失败则手动先开一次 Illustrator |
| 中文变方块 | `getByName` 失败被 catch 吞了；跑 check_env.py 确认字体名 |
| 文字位置偏了 | `tf.position` 是锚点（基线左端），不是视觉中心；用 `tf.width/height` 读实际包围盒再微调 |
| 导出空白 | `artBoardClipping` 为 false 且元素超出画板；或画板 rect 设错，用返回值打印 `doc.artboards[0].artboardRect` 诊断 |
| 脚本第一行报错/乱码 | .jsx 保存为 UTF-8（run_jsx.py 用 utf-8-sig 读取，BOM 有无均可）；内容含中文时不要存成 ANSI |
