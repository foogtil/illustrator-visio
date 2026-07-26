# Visio 完整配方（Python + pywin32 COM）

> 基于 Visio 16（Microsoft 365 桌面版）实测。所有代码片段均已验证，除非特别标注。

## 1. 连接与生命周期

```python
import time, pythoncom, win32com.client

def connect(progid="Visio.Application"):
    """返回 (app, started_by_us)。附加优先, 冷启动自动重试。"""
    try:
        return win32com.client.GetActiveObject(progid), False
    except pythoncom.com_error:
        pass
    for i in range(4):
        try:
            return win32com.client.Dispatch(progid), True
        except pythoncom.com_error as e:
            if e.hresult == -2146959355 and i < 3:  # 0x80080005 应用启动中
                time.sleep(5)
                continue
            raise

app, started = connect()
app.Visible = True   # 建议 True: 部分版本在完全无窗口时行为异常
```

**关闭顺序（防挂死，务必照做）：**

```python
app.AlertResponse = 7   # visResponseNo: 自动对"是否保存"回答"否"
doc.Close()             # Document.Close() 无参数, 不设 AlertResponse 会弹窗挂死
if started:
    app.AlertResponse = 7
    app.Quit()
```

要保留保存过的文件：先 `doc.SaveAs(path)` 再 Close；`AlertResponse=7` 只是兜底。

## 2. 文档与页面

```python
doc = app.Documents.Add("")          # 新建空白绘图(无模板)
doc = app.Documents.Open(r"D:\x.vsdx")  # 打开现有文件
page = app.ActivePage                # 当前页
pages = doc.Pages                    # 页集合, doc.Pages.Add() 加页
```

- **单位：英寸；原点左下角；y 轴向上。** 默认空白页为 Letter(8.5x11) 纵向，且默认开启"自动适应内容"——导出图片尺寸会随内容变化（实测：内容 1.24x6.9 英寸 → 导出 119x659 px，约 96dpi）。
- 手动设页面大小（关闭自适应后）：

```python
ps = page.PageSheet
ps.CellsU("PageWidth").FormulaU = "11 in"
ps.CellsU("PageHeight").FormulaU = "8.5 in"   # 横向 A 纸可写 "297 mm"/"210 mm"
```

## 3. 模具（Stencil）与 Master

### 3.1 定位模具目录（不要硬编码，先跑 check_env.py）

```python
import glob, os
content = os.path.join(app.Path, "Visio Content")     # app.Path = Office16 根目录
dirs = sorted(glob.glob(os.path.join(content, "*")))  # 语言目录: 2052=中文, 1033=英文
pref = next((d for d in dirs if d.endswith("2052")), dirs[0])
stencils = glob.glob(os.path.join(pref, "*.vss*"))
```

本机实测：`...\Office16\Visio Content\2052\` 下有 **362 个模具**，包括：
`BASFLO_M.VSSX`(基本流程图)、`BPMN_M.VSSX`、`XFUNC_M.VSSX`(跨职能/泳道)、
`ORGCH_M.VSSX`(组织结构图)、`BASIC_M.VSSX`(基本形状)、`ARROWS_M.VSSX`(箭头)、
`CONNEC_M.VSSX`(连接线)、`NETSYM_M.VSSX`/`SERVER_M.VSSX`(网络)、
`AZURE*_M.VSSX`、`AWS*_M.VSSX`、`GANTT_M.VSSX`、`TIMELN_M.VSSX` 等。

### 3.2 打开模具并取 master

```python
st = app.Documents.OpenEx(os.path.join(pref, "BASFLO_M.VSSX"), 2 + 64)
# OpenEx flags (VisOpenSaveArgs): 2=只读(RO) 4=Docked 8=最小化 16=不列入最近文件
#                                32=禁用宏 64=隐藏(Hidden) 128=不加载工作区
m = st.Masters.ItemU("Process")   # ItemU = 通用名(跨语言); Item = 本地化名
shp = page.Drop(m, 4.25, 8.0)     # 落到页面, 坐标为形状中心(PinX/PinY)
```

BASFLO_M 常用通用名：`Process`、`Decision`、`Start/End`、`Document`、`Data`、
`Predefined process`、`Connector`(离页连接符)、`Annotation`。

**不知道 master 叫什么时，遍历打印（强烈推荐）：**

```python
for i in range(1, st.Masters.Count + 1):
    mm = st.Masters.Item(i)
    print(mm.NameU, "|", mm.Name)   # 通用名 | 本地名
```

## 4. 形状操作

```python
shp.Text = "处理数据"                                  # 文本(直接支持中文)
shp.CellsU("PinX").FormulaU = "4 in"                   # 移动(中心点)
shp.CellsU("PinY").FormulaU = "6 in"
shp.CellsU("Width").FormulaU = "2 in"                  # 尺寸
shp.CellsU("Height").FormulaU = "1 in"
shp.CellsU("Angle").FormulaU = "45 deg"                # 旋转
shp.CellsU("FillForegnd").FormulaU = "RGB(198,224,180)"  # 填充色
shp.CellsU("LineColor").FormulaU = "RGB(60,60,60)"       # 边框色
shp.CellsU("LineWeight").FormulaU = "1.5 pt"             # 边框粗细
shp.CellsU("Char.Size").FormulaU = "14 pt"               # 字号
shp.CellsU("Char.Color").FormulaU = "RGB(255,0,0)"       # 字色
x = shp.CellsU("PinX").Result("in")                      # 读取坐标(返回浮点)
```

删除：`shp.Delete()`。复制：`shp.Duplicate()`（原位重叠，需再移动）。

## 5. 连接线

### 5.1 方式一：ConnectorToolDataObject（推荐，无需模具，已验证）

```python
conn = page.Drop(app.ConnectorToolDataObject, 0, 0)
conn.CellsU("BeginX").GlueTo(s1.CellsU("PinX"))   # 动态胶水: 自动选连接点
conn.CellsU("EndX").GlueTo(s2.CellsU("PinX"))
conn.Text = "是"                                    # 连线标签
```

### 5.2 方式二：从模具拖连接线 master

```python
conn_st = app.Documents.OpenEx(os.path.join(pref, "CONNEC_M.VSSX"), 2 + 64)
conn = page.Drop(conn_st.Masters.ItemU("Dynamic Connector"), 0, 0)
# 然后同样 GlueTo
```

### 5.3 连线样式

```python
conn.CellsU("EndArrow").FormulaU = "4"        # 箭头样式(0=无, 4=实心三角, 1-45)
conn.CellsU("LineColor").FormulaU = "RGB(100,100,100)"
conn.CellsU("LinePattern").FormulaU = "2"     # 2=虚线
```

`GlueTo`（粘到整个形状=动态连接）与 `GlueToPos`（粘到指定相对坐标=静态连接）：

```python
conn.CellsU("BeginX").GlueToPos(s1, 0.5, 0)   # s1 底边中点
```

## 6. 自动布局

```python
page.Layout()   # 按页面布局设置重排所有形状与连线(已验证可用)
```

布局方向/样式由 PageSheet 单元格控制（可选调优，按需尝试）：
`PlaceStyle`(1=从上到下 2=从左到右…)、`RouteStyle`、`PlaceDirection`、`LineRouteExt`。
经验：先把形状按大致位置放好再 `Layout()`，效果最可控；复杂图建议手工定位不 Layout。

## 7. 导出与保存（均已验证）

```python
page.Export(r"D:\out\diagram.png")     # 按扩展名: .png/.jpg/.svg/.emf 等
                                       # 位图约 96dpi; 要高清就导 .svg(矢量)
doc.SaveAs(r"D:\out\diagram.vsdx")     # 保存源文件
doc.ExportAsFixedFormat(1, r"D:\out\diagram.pdf", 1, 0)
# ExportAsFixedFormat(格式, 路径, 质量, 范围): 格式 1=PDF 2=XPS; 质量 1=屏幕 2=打印; 范围 0=全部页
```

## 8. 完整示例：带分支的流程图

```python
# -*- coding: utf-8 -*-
import os, sys, time, tempfile, pythoncom, win32com.client
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

app, started = connect()
app.Visible = True
doc = app.Documents.Add("")
page = app.ActivePage
st = app.Documents.OpenEx(
    r"C:\Program Files\Microsoft Office\Root\Office16\Visio Content\2052\BASFLO_M.VSSX", 2 + 64)

def drop(name, x, y, text):
    shp = page.Drop(st.Masters.ItemU(name), x, y)
    shp.Text = text
    return shp

def link(a, b, label=""):
    c = page.Drop(app.ConnectorToolDataObject, 0, 0)
    c.CellsU("BeginX").GlueTo(a.CellsU("PinX"))
    c.CellsU("EndX").GlueTo(b.CellsU("PinX"))
    if label: c.Text = label

start = drop("Start/End", 4.25, 9.8, "开始")
p1    = drop("Process",   4.25, 8.0, "采集数据")
dec   = drop("Decision",  4.25, 5.8, "数据完整?")
p2    = drop("Process",   4.25, 3.4, "生成报告")
p3    = drop("Process",   7.20, 5.8, "补齐缺失")
end   = drop("Start/End", 4.25, 1.4, "结束")

link(start, p1); link(p1, dec)
link(dec, p2, "是"); link(dec, p3, "否"); link(p3, p1)
link(p2, end)

dec.CellsU("FillForegnd").FormulaU = "RGB(255,230,153)"

out = os.path.join(tempfile.gettempdir(), "visio_demo.png")
page.Export(out)
print("已导出:", out, flush=True)

app.AlertResponse = 7
doc.Close()
if started: app.Quit()
```

## 9. 排错

| 现象 | 对策 |
| --- | --- |
| `0x80080005 服务器运行失败` | 冷启动中超时，用 §1 的 connect() 重试模式 |
| 脚本无输出卡死 | 弹窗阻塞：确认 Close/Quit 前有 `AlertResponse=7`；任务管理器结束 VISIO.EXE 重来 |
| `OpenEx` 报文件找不到 | 用 §3.1 代码动态定位；确认语言目录（2052/1033） |
| `Masters.ItemU("xxx")` 报错 | master 名不对：用 §3.2 遍历代码打印全部 NameU |
| 导出的 PNG 很小/尺寸怪 | 页面"自动适应内容"所致：先放好全部形状再导出，或用 §2 固定页面尺寸；要高清导 SVG |
| 中文变方块 | 指定字体：`shp.CellsU("Char.Font").FormulaU = '"Microsoft YaHei"'` |
