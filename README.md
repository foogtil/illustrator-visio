# illustrator-visio

AI 编程智能体（opencode / Claude Code / Codex CLI）的自动化绘图 skill —— 在 Windows 上通过 COM/JSX 驱动**真实**的 Adobe Illustrator 和 Microsoft Visio 制图，而非手写 SVG。

本仓库包含**两个独立 skill**，已针对科研绘图做针对性优化：

| skill | 目录 | 驱动工具 | 适用场景 | 针对科研的优化 |
| --- | --- | --- | --- | --- |
| `illustrator` | `illustrator/` | Illustrator（COM + ExtendScript/JSX） | Logo、图标、矢量插画、信息图、科研论文插图 | 期刊栏宽/字号/线宽/dpi 规范、色盲友好配色、可复用作图辅助函数、多面板组合图模板、投稿检查清单 |
| `visio` | `visio/` | Visio（pywin32 COM） | 流程图、实验流程图、技术路线图、研究方案图、架构图、网络拓扑、组织结构图、泳道图、BPMN | 实验/技术路线流程图模板、阶段配色、矢量导出（SVG/PDF 供论文），检查清单 |

每个子目录都自带 `SKILL.md`（触发条件 + 工作流 + 快速开始 + 排错）、`reference/`（完整配方与科研绘图模板）、`scripts/`（环境自检与执行器）以及 LICENSE。

## 前置条件

- Windows（必须，COM 自动化）
- Python 3 + pywin32
- 已安装 Adobe Illustrator / Microsoft Visio（至少其一）

## 安装

将你想用的 skill 目录放到对应 agent 的 skills 路径下（可两个都装，也可只装其一）：

```
# opencode
~/.config/opencode/skills/illustrator/
~/.config/opencode/skills/visio/

# Claude Code
~/.claude/skills/illustrator/
~/.claude/skills/visio/

# Codex CLI
~/.codex/skills/illustrator/
~/.codex/skills/visio/

# 其他 (如 math-rigor 等所在)
~/.agents/skills/illustrator/
~/.agents/skills/visio/
```

重启对应 agent 后生效。

## 快速验证

```powershell
[Console]::OutputEncoding=[Text.Encoding]::UTF8
python -X utf8 illustrator/scripts/check_env.py    # Illustrator 自检(COM/字体)
python -X utf8 visio/scripts/check_env.py          # Visio 自检(COM/模具清单)
```

## 工作流

1. **规划** —— 确定画布/页面尺寸、元素坐标与层级；科研图先定期刊栏宽或页面版式
2. **写脚本** —— 生成 Python（Visio）或 JSX（Illustrator）临时脚本
3. **执行** —— 脚本驱动真实软件出图
4. **导出 PNG / SVG 验证** —— 回看位图检查布局与文字（视觉模型用 Read，否则像素采样/结构化断言）
5. **清理** —— 删除临时文件

## 文件结构

```
illustrator-visio/
├── README.md                    # 本文件
├── LICENSE
├── illustrator/                 # skill: illustrator (Illustrator + 科研论文插图规范)
│   ├── SKILL.md
│   ├── reference/
│   │   ├── illustrator-jsx.md       # JSX 完整配方(坐标/渐变/文字/导出)
│   │   └── scientific-figures.md    # 期刊论文插图规范与模板(栏宽/字号/dpi/配色/多面板示例)
│   └── scripts/
│       ├── check_env.py             # 环境自检(COM/版本/中西文字体)
│       └── run_jsx.py               # JSX 执行器(COM DoJavaScript)
└── visio/                      # skill: visio (Visio + 科研流程图模板)
    ├── SKILL.md
    ├── reference/
    │   ├── visio-pywin32.md         # Visio 完整配方(模具/连接线/样式/导出)
    │   └── research-flowcharts.md   # 实验流程图/技术路线图规范与模板
    └── scripts/
        └── check_env.py             # 环境自检(COM/版本/模具清单)
```

## 限制

- 仅 Windows
- 需要桌面端软件许可证
- AI 模型需支持图片读取才能通过"导出→回看→迭代"闭环验证（大部分支持；不支持时用像素采样/结构化断言）

## License

MIT

## 联系方式
如有任何问题，请通过以下邮箱地址与我联系：lhy_ynmzdx@qq.com


*免责声明：本工具仅供学习交流使用。使用者需自行承担使用本工具的一切风险与责任。*
