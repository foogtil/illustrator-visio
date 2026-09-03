# 科研论文插图规范与模板（期刊投稿向）

> 面向 SCI/中文核心期刊投稿的 Illustrator 制图规范。所有换算基于 1 mm = 2.8346 pt（1 pt = 1/72 in）。
> 各期刊要求以目标期刊最新 Guide for Authors 为准，本文件提供通用参考值。

## 1. 画板尺寸 = 期刊栏宽（投稿第一件事）

**先定画板宽度再画图**：`app.documents.add(ColorSpace, 栏宽pt, 高度pt)` 的宽直接等于最终印刷宽度，这样在画板内设定的字号/线宽就是最终印刷尺寸，无需后期缩放换算。

| 出版社/期刊 | 单栏 mm (pt) | 1.5 栏 mm (pt) | 双栏 mm (pt) | 备注 |
| --- | --- | --- | --- | --- |
| Cell Press | 85 (241) | 114 (323) | 174 (493) | 常见上限 |
| Nature 系列 | 89 (252) | 120 (340) | 183 (519) | 高度 ≤ 247 mm |
| Science 系列 | 55 (156) | 120 (340) | 176 (499) | 另有方形/宽版规格 |
| IEEE | 89 (252) | — | 182 (514) | 3.5 / 7.16 英寸 |
| Elsevier（通用） | 90 (255) | 140 (397) | 190 (539) | 部分期刊不同 |
| Springer（通用） | 84 (238) | 129 (366) | 174 (493) | 医学类常用 |

```javascript
// 例: Nature 双栏宽 + 半页高
var doc = app.documents.add(DocumentColorSpace.RGB, 519, 280);  // 183mm x 99mm
var W = 519, H = 280;
function Y(t) { return H - t; }
function MM(mm) { return mm * 2.8346; }
```

- 常规建议：**先做小尺寸（单栏）再放大**——在 1:1 尺寸下确保文字清晰，放大不会糊，缩小才会。
- 多面板图先画"排版草图"：面板数量、每面板宽高、面板间距（建议 3-5 mm = 9-14 pt）。

## 2. 字体与字号规范

- **字体**：图内西文用 Arial / Helvetica 无衬线字体；中文期刊图内中文用宋体/微软雅黑（以 check_env.py 输出为准）。全图字体统一。
- **字号（最终印刷尺寸）**：
  - 坐标轴刻度标签、图例：6-8 pt
  - 坐标轴标题：7-9 pt
  - 面板标签（A/B/C）：8-10 pt，加粗
  - 最小字号 ≥ 5 pt（低于此值印刷易糊）
  - 多数期刊**不需要图内大标题**（题注在正文/图注中）
- 检查方法：导出 PNG 后 Read 回看，文字应清晰且无重叠。

## 3. 线宽规范

| 元素 | 线宽 (pt) |
| --- | --- |
| 坐标轴 | 0.75-1 |
| 数据线/折线 | 1-1.5 |
| 误差棒、显著性标注线 | 0.5-1 |
| 辅助线/网格线（慎用） | 0.25-0.5 |
| 刻度线 | 0.5-0.75 |

- 印刷线宽下限约 0.25 pt。避免使用 3D 效果、阴影、渐变背景等"chartjunk"。

## 4. 分辨率与导出格式

### 4.1 位图导出（PNG）：dpi → 缩放百分比换算

`ExportOptionsPNG24` 的 scale 百分比与 dpi 关系：**scale% = dpi ÷ 72 × 100**。

| 目标 dpi | horizontalScale/verticalScale | 适用 |
| --- | --- | --- |
| 300 | 417 | 半色调图（照片类） |
| 600 | 833 | 组合图（线+半色调） |
| 1200 | 1667 | 纯线图（部分期刊要求 1000-1200） |

```javascript
function exportPNG(doc, path, dpi) {           // 通用高清导出
    var s = Math.round(dpi / 72 * 100);
    var o = new ExportOptionsPNG24();
    o.horizontalScale = s; o.verticalScale = s;
    o.artBoardClipping = true;
    o.transparency = false;                    // 期刊图通常要白底
    doc.exportFile(new File(path), ExportType.PNG24, o);
}
exportPNG(doc, "<输出>/fig1_300dpi.png", 300);
```

### 4.2 矢量导出（推荐用于线图/组合图）

```javascript
// PDF (现代期刊普遍接受)
doc.saveAs(new File("D:/out/fig1.pdf"), new PDFSaveOptions());

// EPS (部分老期刊要求; 使用 CMYK 文档时颜色更可控)
doc.saveAs(new File("D:/out/fig1.eps"), new EPSSaveOptions());
```

- 线图/图表优先投矢量 PDF/EPS：无限分辨率，文件小。
- AI 源文件存档：`doc.saveAs(new File("D:/out/fig1.ai"))`。

### 4.3 色彩模式（印刷 vs 在线）

```javascript
var doc = app.documents.add(DocumentColorSpace.CMYK, W, H);  // 印刷期刊
var doc = app.documents.add(DocumentColorSpace.RGB, W, H);   // 在线/无彩印要求
```

- 纯在线期刊（多数 OA）RGB 即可；要求彩印的印刷期刊用 CMYK。**不确定时先做 RGB，AI 源文件可无损转 CMYK**。

## 5. 科研配色方案

### 5.1 色盲友好配色（Okabe-Ito，推荐）

对红绿色盲读者可区分的 8 色方案，数据系列按此顺序取色：

```javascript
var PALETTE = [
    rgb(230, 159,   0),   // 橙
    rgb( 86, 180, 233),   // 天蓝
    rgb(  0, 158, 115),   // 蓝绿
    rgb(240, 228,  66),   // 黄
    rgb(  0, 114, 178),   // 蓝
    rgb(213,  94,   0),   // 朱红
    rgb(204, 121, 167),   // 紫红
    rgb(  0,   0,   0)    // 黑
];
```

### 5.2 机理图/示意图通用规则

- 同类别元素同色系（如所有"酶"绿色、所有"信号"红色），图例一一对应。
- 一图配色不超过 5-6 种，背景留白。
- 中英文标注字号一致，箭头统一线宽。

## 6. 可复用 JSX 辅助函数库（复制到脚本开头即可）

```javascript
var W = doc.width || 519, H = doc.height || 280;   // 按实际画板
function rgb(r, g, b) { var c = new RGBColor(); c.red = r; c.green = g; c.blue = b; return c; }
function MM(mm) { return mm * 2.8346; }             // 毫米 → pt
function Y(top) { return H - top; }                 // 距顶 → JSX y
function hasCJK(s) { return /[\u4e00-\u9fff]/.test(s); }

// 文字(自动选字体: 中文→雅黑系, 西文→Arial系; 返回 textFrame)
function addText(doc, str, x, y, size, color, bold) {
    var tf = doc.textFrames.add();
    tf.contents = str;
    var ca = tf.textRange.characterAttributes;
    ca.size = size; ca.fillColor = color;
    var cand = hasCJK(str)
        ? (bold ? ["MicrosoftYaHei-Bold", "NotoSansSC-Bold", "MicrosoftYaHei"] 
                : ["MicrosoftYaHei", "NotoSansSC-Regular", "KaiTi", "FangSong"])
        : (bold ? ["Arial-BoldMT", "Helvetica-Bold", "MicrosoftYaHei-Bold"] 
                : ["ArialMT", "Arial", "Helvetica"]);
    for (var i = 0; i < cand.length; i++) {
        try { ca.textFont = app.textFonts.getByName(cand[i]); break; } catch (e) {}
    }
    tf.position = [x, y];
    return tf;
}

// 面板标签 A/B/C (左上角, 加粗)
function panelLabel(doc, txt, left, top, size) {
    addText(doc, txt, left, Y(top) - size * 1.2, size || 10, rgb(0, 0, 0), true);
}

// 坐标轴 (x0,y0 = 左下原点; w/h = 轴长; 刻度外置可选)
function drawAxis(doc, x0, y0, w, h, color, lw) {
    var p = doc.pathItems.add();
    p.setEntirePath([[x0, y0], [x0 + w, y0]]);
    p.stroked = true; p.filled = false; p.strokeColor = color; p.strokeWidth = lw;
    var q = doc.pathItems.add();
    q.setEntirePath([[x0, y0], [x0, y0 + h]]);
    q.stroked = true; q.filled = false; q.strokeColor = color; q.strokeWidth = lw;
}

// 刻度线: xs = x方向刻度位置数组(相对轴长比例0~1), ys 同理
function drawTicks(doc, x0, y0, w, h, xs, ys, tickLen, color, lw) {
    var i, p;
    for (i = 0; i < xs.length; i++) {
        p = doc.pathItems.add();
        p.setEntirePath([[x0 + w * xs[i], y0 - tickLen], [x0 + w * xs[i], y0 + tickLen]]);
        p.stroked = true; p.filled = false; p.strokeColor = color; p.strokeWidth = lw;
    }
    for (i = 0; i < ys.length; i++) {
        p = doc.pathItems.add();
        p.setEntirePath([[x0 - tickLen, y0 + h * ys[i]], [x0 + tickLen, y0 + h * ys[i]]]);
        p.stroked = true; p.filled = false; p.strokeColor = color; p.strokeWidth = lw;
    }
}

// 误差棒: 竖线 + 上下短横 (cx=柱中心x, yTop/yBottom=误差上下限)
function errorBar(doc, cx, yTop, yBottom, halfW, color, lw) {
    var seg, style = function (p) {
        p.stroked = true; p.filled = false; p.strokeColor = color; p.strokeWidth = lw;
    };
    seg = doc.pathItems.add(); seg.setEntirePath([[cx, yTop], [cx, yBottom]]); style(seg);
    seg = doc.pathItems.add(); seg.setEntirePath([[cx - halfW, yTop], [cx + halfW, yTop]]); style(seg);
    seg = doc.pathItems.add(); seg.setEntirePath([[cx - halfW, yBottom], [cx + halfW, yBottom]]); style(seg);
}

// 显著性标注: 连接 x1→x2 的折线 + 星号("*"/"**"/"n.s.")
function sigBar(doc, x1, x2, y, drop, label, color, lw) {
    var p = doc.pathItems.add();
    p.setEntirePath([[x1, y - drop], [x1, y], [x2, y], [x2, y - drop]]);
    p.stroked = true; p.filled = false; p.strokeColor = color; p.strokeWidth = lw;
    var tf = addText(doc, label, 0, y + 3, 7, color, false);
    tf.position = [(x1 + x2) / 2 - tf.width / 2, y + 3];   // 居中(实测 tf.width 创建后可用)
}

// 比例尺(显微照片用): lenPt = 实际长度的 pt 数
function scaleBar(doc, x, y, lenPt, label, color, lw) {
    var p, style = function (q) {
        q.stroked = true; q.filled = false; q.strokeColor = color; q.strokeWidth = lw;
    };
    p = doc.pathItems.add(); p.setEntirePath([[x, y], [x + lenPt, y]]); style(p);
    p = doc.pathItems.add(); p.setEntirePath([[x, y - 4], [x, y + 4]]); style(p);
    p = doc.pathItems.add(); p.setEntirePath([[x + lenPt, y - 4], [x + lenPt, y + 4]]); style(p);
    var tf = addText(doc, label, 0, y + 6, 7, color, false);
    tf.position = [x + lenPt / 2 - tf.width / 2, y + 6];
}

// 通用高清 PNG 导出
function exportPNG(doc, path, dpi) {
    var s = Math.round(dpi / 72 * 100);
    var o = new ExportOptionsPNG24();
    o.horizontalScale = s; o.verticalScale = s;
    o.artBoardClipping = true; o.transparency = false;
    doc.exportFile(new File(path), ExportType.PNG24, o);
}
```

## 7. 完整示例：双面板期刊组合图

规格：Nature 双栏宽 183 mm (519 pt)，高 280 pt；面板 A = 柱状图（3 组 ± 误差棒 + 显著性标注），面板 B = 简单机制示意图；A/B 标签 10 pt 加粗；导出 300 dpi PNG + PDF + AI。

```javascript
var doc = app.documents.add(DocumentColorSpace.RGB, 519, 280);
var W = 519, H = 280;
var gray = rgb(40, 40, 40), blue = rgb(0, 114, 178);

function rgb(r, g, b) { var c = new RGBColor(); c.red = r; c.green = g; c.blue = b; return c; }
function Y(t) { return H - t; }
/* ==== 粘贴 §6 的辅助函数库 (addText/panelLabel/drawAxis/errorBar/sigBar/exportPNG) ==== */

// ---- 面板 A: 柱状图 (左半: x 0~255) ----
var x0 = 30, baseY = 70, axW = 180, axH = 150;
drawAxis(doc, x0, baseY, axW, axH, gray, 0.75);
drawTicks(doc, x0, baseY, axW, axH, [0.5], [0.5, 1], 3, gray, 0.5);

var groups = [
    { label: "Control",  mean: 1.00, sd: 0.12 },
    { label: "1 mM",     mean: 1.45, sd: 0.18 },
    { label: "10 mM",    mean: 2.10, sd: 0.22 }
];
var maxV = 2.5, barW = 28;
for (var i = 0; i < groups.length; i++) {
    var cx = x0 + 45 + i * 55;
    var mh = groups[i].mean / maxV * axH;
    var eh = groups[i].sd   / maxV * axH;
    var bar = doc.pathItems.rectangle(baseY + mh, cx - barW / 2, barW, mh);
    bar.fillColor = blue; bar.stroked = false;
    errorBar(doc, cx, baseY + mh + eh, baseY + mh - eh, 3, gray, 0.75);
    var lb = addText(doc, groups[i].label, 0, baseY - 22, 7, gray, false);
    lb.position = [cx - lb.width / 2, baseY - 22];
}
// 竖直 y 轴标题: 先 create() 再 rotate(90), 最后把 position 设为左侧中部
var yt = addText(doc, "Relative expression", 0, 0, 8, gray, false);
yt.rotate(90);
yt.position = [x0 - 16, baseY + axH / 2 - yt.height / 2];   // 轴左、垂直居中
sigBar(doc, x0 + 45, x0 + 100, baseY + axH + 18, 6, "*", gray, 0.75);      // 组1 vs 组2
sigBar(doc, x0 + 100, x0 + 155, baseY + axH + 32, 6, "**", gray, 0.75);    // 组2 vs 组3
panelLabel(doc, "A", 8, 8, 10);

// ---- 面板 B: 机制示意图 (右半: x 275~511) ----
var bx = 320, by = 180;
var cell = doc.pathItems.ellipse(by + 30, bx - 30, 60, 60);   // 细胞
cell.fillColor = rgb(230, 230, 230); cell.stroked = true;
cell.strokeColor = gray; cell.strokeWidth = 1;
var sig = doc.pathItems.add();                                  // 信号箭头
sig.setEntirePath([[bx + 55, by], [bx + 105, by]]);
sig.stroked = true; sig.filled = false; sig.strokeColor = rgb(213, 94, 0); sig.strokeWidth = 1.5;
addText(doc, "Nucleus", bx - 5, by + 15, 6, gray, false);
panelLabel(doc, "B", 265, 8, 10);

// ---- 导出 ----
exportPNG(doc, "<临时目录>/fig_preview.png", 150);     // 150dpi 快速验证回看
exportPNG(doc, "<输出目录>/Fig1_300dpi.png", 300);      // 投稿位图版
doc.saveAs(new File("<输出目录>/Fig1.pdf"), new PDFSaveOptions());
doc.saveAs(new File("<输出目录>/Fig1.ai"));
doc.close(SaveOptions.DONOTSAVECHANGES);
"fig done";
```

## 8. 投稿前检查清单

- [ ] 画板宽度 = 目标期刊栏宽（§1 表，投稿前核对 Guide for Authors）
- [ ] 图内文字 ≥ 5 pt（建议 6-8 pt），字体统一 Arial/Helvetica（或期刊指定）
- [ ] 线宽 ≥ 0.25 pt，坐标轴 0.75-1 pt
- [ ] 每个面板有 A/B/C 标签（与图注对应）
- [ ] 误差棒/显著性标注有图注说明（"Data are mean ± SD"）
- [ ] 位图 ≥ 300 dpi（线图/组合图投矢量 PDF/EPS）
- [ ] 印刷期刊已确认 RGB/CMYK 要求
- [ ] 色盲友好配色或图案+颜色双编码（不只靠颜色区分数据系列）
- [ ] 导出 PNG 已 Read 回看验证（布局/重叠/文字清晰）

## 9. 常见科研图类型速查

| 图类型 | 实现要点 |
| --- | --- |
| 柱状图 | `rectangle` + `errorBar`；组间对比用 `sigBar` |
| 折线图 | `setEntirePath` 多点连线 + 数据点用 `ellipse` |
| 散点图 | 循环 `ellipse`，用 `PALETTE` 按组取色 |
| 箱线图 | 多个 `rectangle` + 中位线 + `errorBar` |
| 热力图 | 网格 `rectangle` + 自定义颜色映射（RGB 插值） |
| 机制/通路图 | 圆角矩形 + 箭头 + `addText` 标注，配色见 §5.2 |
| 显微照片+比例尺 | 置入图片(`doc.placedItems.add()`) + `scaleBar` |
