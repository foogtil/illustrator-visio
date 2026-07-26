# illustrator-visio

AI 编程智能体（opencode / Claude Code / Codex CLI）的自动化绘图 skill —— 在 Windows 上通过 COM/JSX 驱动**真实**的 Adobe Illustrator 和 Microsoft Visio 制图，而非手写 SVG。

## 能做什么

| 场景 | 驱动工具 | 示例 |
| --- | --- | --- |
| 流程图、业务流程、架构图、网络拓扑、BPMN | Visio（pywin32 COM） | 分支流程、Azure/AWS 架构 |
| Logo、图标、矢量插画、精致信息图 | Illustrator（JSX） | 品牌 Logo、柱状图、信息图 |

## 前置条件

- Windows（必须，COM 自动化）
- Python 3 + pywin32
- 已安装 Adobe Illustrator / Microsoft Visio（至少其一）

## 安装

将 `illustrator-visio` 目录放到对应 agent 的 skills 路径下：

```
# opencode
~/.config/opencode/skills/illustrator-visio/

# Claude Code
~/.claude/skills/illustrator-visio/

# Codex CLI
~/.codex/skills/illustrator-visio/
```

重启对应 agent 后生效。

## 快速验证

```powershell
[Console]::OutputEncoding=[Text.Encoding]::UTF8
python -X utf8 scripts/check_env.py
```

输出本机 Illustrator / Visio COM 状态、版本号、可用中文字体、Visio 模具清单。

## 工作流

1. **规划** —— 确定画布尺寸、元素坐标与层级
2. **写脚本** —— 生成 Python（Visio）或 JSX（Illustrator）临时脚本
3. **执行** —— 脚本驱动真实软件出图
4. **导出 PNG 验证** —— 回看位图检查布局与文字
5. **清理** —— 删除临时文件

## 文件结构

```
illustrator-visio/
├── SKILL.md                    # skill 主入口（触发条件、决策表、快速开始、排错）
├── reference/
│   ├── visio-pywin32.md        # Visio 完整配方（模具、连接线、样式、导出、完整示例）
│   └── illustrator-jsx.md      # Illustrator JSX 完整配方（坐标、渐变、文字、导出、完整示例）
└── scripts/
    ├── check_env.py            # 环境自检
    └── run_jsx.py              # JSX 执行器（COM DoJavaScript）
```

## 限制

- 仅 Windows
- 需要桌面端软件许可证
- AI 模型需支持图片读取才能通过"导出→回看→迭代"闭环验证（大部分支持）

## License

MIT

## 联系方式
如有任何问题，请通过以下邮箱地址与我联系：lhy_ynmzdx@qq.com


*免责声明：本工具仅供学习交流使用。使用者需自行承担使用本工具的一切风险与责任。*
