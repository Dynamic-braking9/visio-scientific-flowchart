# 科研流程图生成器 (visio-scientific-flowchart)

Hermes Agent Skill：生成出版级科研流程图，支持 SVG（通用）、VSDX（Visio）、PPTX（PowerPoint）三种输出格式。

> **[English Version → README.md](README.md)**

## ✨ 功能特性

| 功能 | 说明 |
|------|------|
| 📊 ISO 5807 符号 | 处理、判断、输入输出、起止、子流程、数据库、注释 |
| 🎨 研究阶段配色 | 6 种语义化颜色（数据/预处理/分析/判断/验证/结果） |
| 📐 自动布局 | 垂直（自上而下）和水平（自左而右） |
| 🔀 分支支持 | 判断节点带"是/否"标签，虚线回环连线 |
| 📝 多格式输出 | SVG（矢量）、VSDX（Visio）、PPTX（PowerPoint） |
| 🌏 中日韩支持 | 所有格式均支持中文文字 |

## 🚀 快速开始

### SVG 输出（无依赖）

```bash
# 纯 Python 标准库，无需安装任何第三方库！
python templates/svg_flowchart.py
```

### PPTX 输出

```bash
pip install python-pptx
python templates/pptx_flowchart.py
```

### VSDX 输出（需要 Visio）

```bash
pip install pywin32
python templates/visio_com.py
```

## 使用方法

```python
from templates.svg_flowchart import FlowchartBuilder

fc = FlowchartBuilder("my_flow.svg", title="机器学习研究流程")

fc.start("开始")
fc.io_input("加载数据集", phase="data")
fc.process("特征工程", phase="preprocess")
fc.decision("质量检查？", phase="check")
fc.process("模型训练", phase="analysis")
fc.subprocess("交叉验证", phase="verify")
fc.io_output("输出报告", phase="result")
fc.end("结束")

fc.connect_all()   # 按添加顺序自动连线
fc.save()
```

### 高级用法：自定义连线与标签

```python
n1 = fc.start("开始")
n2 = fc.decision("准确率 ≥ 90%？", phase="check")
n3 = fc.io_output("输出报告", phase="result")
n4 = fc.process("调参重训", phase="analysis")

fc.connect(n1, n2)
fc.connect(n2, n3, label="是")      # 是 → 输出
fc.connect(n2, n4, label="否")      # 否 → 重训
fc.connect(n4, n2, style="dashed")  # 虚线回环
```

## 颜色方案（按研究阶段）

| 阶段 | 颜色 | 适用场景 |
|------|------|----------|
| `data` | 🔵 蓝 | 数据加载、清洗、预处理 |
| `preprocess` | 🟣 紫 | 特征工程、降维 |
| `analysis` | 🟢 绿 | 模型训练、参数优化 |
| `check` | 🟡 黄 | 质量检查、假设检验 |
| `verify` | 🟠 橙 | 交叉验证、鲁棒性检验 |
| `result` | 🩷 粉 | 结论、报告、发表 |

## ISO 5807 流程图符号

| 符号 | 形状 | 方法 | 用途 |
|------|------|------|------|
| 处理 | 矩形 | `process()` | 一般处理步骤 |
| 判断 | 菱形 | `decision()` | 条件判断 |
| 输入/输出 | 平行四边形 | `io_input()` / `io_output()` | 数据输入输出 |
| 起止 | 圆角矩形 | `start()` / `end()` | 流程开始和结束 |
| 子流程 | 双边框矩形 | `subprocess()` | 子程序/子流程 |
| 数据库 | 圆柱体 | `database()` | 数据存储 |
| 注释 | 折角矩形 | `annotation()` | 补充说明 |

## 输出格式对比

| 格式 | 可编辑 | 依赖 | 最佳用途 |
|------|--------|------|----------|
| **SVG** | Inkscape/浏览器 | 无（纯标准库） | 论文、Word、LaTeX |
| **VSDX** | Visio | Visio + pywin32 | Visio 用户 |
| **PPTX** | PowerPoint | python-pptx | 演示文稿 |

## 📁 项目结构

```
visio-scientific-flowchart/
├── SKILL.md                          # Skill 定义
├── README.md                         # 英文文档
├── README.zh.md                      # 中文文档
├── LICENSE                           # MIT + AI 声明
├── requirements.txt                  # Python 依赖
├── .gitignore
├── templates/
│   ├── __init__.py
│   ├── svg_flowchart.py              # SVG 生成器（主模块，无依赖）
│   ├── visio_com.py                  # Visio COM 自动化
│   └── pptx_flowchart.py             # PowerPoint 生成器
└── 示例_深度学习流程.svg              # 示例 SVG 输出
```

## 安装为 Hermes Agent Skill

```bash
git clone https://github.com/Dynamic-braking9/visio-scientific-flowchart.git     ~/.hermes/skills/creative/visio-scientific-flowchart
```

## 许可证

MIT — 见 [LICENSE](LICENSE)
