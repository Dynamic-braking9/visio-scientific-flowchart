---
name: visio-scientific-flowchart
description: "Use when generating scientific research flowcharts for Visio. Produces SVG (universal), VSDX (via COM automation on Windows with Visio), or PPTX (PowerPoint fallback). Supports ISO 5807 symbols, auto-layout, research-phase color schemes, and LaTeX-style math labels."
version: 1.0.0
author: Dynamic-braking9
license: MIT
metadata:
  hermes:
    tags: [visio, flowchart, scientific, research, svg, pptx, diagram, ISO-5807]
    related_skills: [excalidraw, architecture-diagram, document-generation, powerpoint]
---

# 科研流程图生成器 (Visio Scientific Flowchart)

生成符合科研规范的流程图，支持三种输出格式：SVG（通用）、VSDX（Visio COM 自动化）、PPTX（PowerPoint 备选）。

## Overview

科研论文、项目申报书、结题报告中常需流程图来展示研究方法和实验步骤。本 skill 提供完整的 Python 工具链：

1. **SVG 输出**（默认/推荐）— 所有平台可用，可直接插入 Word/Visio/LaTeX
2. **VSDX 输出**（高级）— 通过 COM 接口直接操控 Visio，需要 Windows + Visio 已安装
3. **PPTX 输出**（备选）— 通过 python-pptx 生成 PowerPoint 流程图

## When to Use

- 科研论文中的研究方法流程图
- 实验步骤示意图
- 项目申报书中的技术路线图
- 论文答辩 PPT 中的流程说明
- 系统架构或数据流示意图

Don't use for: 纯软件架构图（用 `architecture-diagram`）、手绘风格图（用 `excalidraw`）、平面建筑图。

## 输出格式选择

| 场景 | 推荐格式 | 原因 |
|------|----------|------|
| 论文投稿（Word） | SVG | 矢量无损，Word 直接插入 |
| 论文投稿（LaTeX） | SVG | 直接用 `\includesvg` 或转 PDF |
| 需要在 Visio 编辑 | VSDX | 原生格式，可继续修改 |
| 答辩 PPT | PPTX | 直接复制到演示文稿 |
| 通用分享 | SVG | 浏览器直接打开，无依赖 |

## 快速开始

### 安装依赖

```bash
# SVG 输出（无额外依赖，纯 Python 标准库即可）

# PPTX 输出
pip install python-pptx

# VSDX 输出（需要 Windows + Visio 已安装）
pip install pywin32
```

### 基本用法（SVG）

```python
from templates.svg_flowchart import FlowchartBuilder

fc = FlowchartBuilder("research_flow.svg", title="深度学习图像分类研究流程")

# 添加节点
fc.start("开始")
fc.process("数据采集", phase="data")
fc.process("数据预处理", phase="data")
fc.decision("数据质量检查", phase="check")
fc.process("模型训练", phase="analysis")
fc.process("模型评估", phase="analysis")
fc.io_output("结果输出", phase="result")
fc.end("结束")

# 连接节点（自动按添加顺序连接）
fc.connect_all()

# 生成文件
fc.save()
print(f"✅ 流程图已生成: research_flow.svg")
```

### 高级用法（带分支和自定义连接）

```python
from templates.svg_flowchart import FlowchartBuilder

fc = FlowchartBuilder("complex_flow.svg", title="机器学习实验流程")

n1 = fc.start("开始")
n2 = fc.io_input("数据集加载", phase="data")
n3 = fc.process("训练集/验证集划分", phase="data")
n4 = fc.process("特征工程", phase="preprocess")
n5 = fc.decision("特征是否充分？", phase="check")
n6 = fc.process("模型选择\n(CNN/Transformer)", phase="analysis")
n7 = fc.process("超参数调优", phase="analysis")
n8 = fc.decision("性能达标？", phase="check")
n9 = fc.subprocess("交叉验证", phase="verify")
n10 = fc.io_output("实验报告", phase="result")
n11 = fc.end("结束")

# 自定义连接
fc.connect(n1, n2)
fc.connect(n2, n3)
fc.connect(n3, n4)
fc.connect(n4, n5)
fc.connect(n5, n6, label="是")
fc.connect(n5, n4, label="否", style="dashed")  # 回环
fc.connect(n6, n7)
fc.connect(n7, n8)
fc.connect(n8, n9, label="否")
fc.connect(n8, n10, label="是")
fc.connect(n9, n7, label="调整", style="dashed")
fc.connect(n10, n11)

fc.save()
```

### VSDX 输出（Visio COM）

```python
from templates.visio_com import VisioFlowchart

vf = VisioFlowchart("research_flow.vsdx", title="实验流程")

vf.add_start("开始")
vf.add_process("数据采集", phase="data")
vf.add_process("模型训练", phase="analysis")
vf.add_decision("结果验证", phase="check")
vf.add_end("结束")

vf.auto_connect()
vf.save()
```

## 科研流程图 Phase 颜色体系

内置 6 种研究阶段配色，语义化映射：

| Phase | 中文含义 | 填充色 | 边框色 | 适用节点 |
|-------|----------|--------|--------|----------|
| `data` | 数据采集/处理 | 浅蓝 `#DBEAFE` | 蓝 `#3B82F6` | 数据加载、清洗、预处理 |
| `preprocess` | 预处理/特征 | 浅紫 `#EDE9FE` | 紫 `#8B5CF6` | 特征工程、降维、编码 |
| `analysis` | 分析/建模 | 浅绿 `#D1FAE5` | 绿 `#10B981` | 模型训练、参数优化 |
| `check` | 检验/判断 | 浅黄 `#FEF3C7` | 黄 `#F59E0B` | 质量检查、假设检验 |
| `verify` | 验证/确认 | 浅橙 `#FFEDD5` | 橙 `#F97316` | 交叉验证、鲁棒性检验 |
| `result` | 结果/输出 | 浅粉 `#FCE7F3` | 粉 `#EC4899` | 结论、报告、发表 |

特殊节点颜色（固定）：
- **开始/结束**（椭圆）：浅灰 `#F3F4F6`，灰边框 `#6B7280`
- **I/O**（平行四边形）：浅青 `#CFFAFE`，青边框 `#06B6D4`

## ISO 5807 标准流程图符号

| 符号 | 形状 | Python 方法 | 用途 |
|------|------|------------|------|
| 处理 (Process) | 矩形 | `process()` | 一般处理步骤 |
| 判断 (Decision) | 菱形 | `decision()` | 条件判断、分支 |
| 输入/输出 (I/O) | 平行四边形 | `io_input()` / `io_output()` | 数据输入输出 |
| 起止 (Terminal) | 圆角矩形/椭圆 | `start()` / `end()` | 流程开始和结束 |
| 子流程 (Subprocess) | 双边框矩形 | `subprocess()` | 调用子程序/子流程 |
|| 注释 (Annotation) | 折角矩形 | `annotation()` | 补充说明 |
|| 数据库 (Database) | 圆柱体 | `database()` | 数据存储 |

## 布局方向

```python
# 自上而下（默认，适合大多数论文）
fc = FlowchartBuilder("flow.svg", direction="vertical")

# 自左而右（适合宽屏展示或横向排版）
fc = FlowchartBuilder("flow.svg", direction="horizontal")

# 泳道图（规划中，暂未实现）
# fc = FlowchartBuilder("flow.svg", layout="swimlane",
#                       lanes=["数据层", "模型层", "评估层"])
```

## 节点尺寸与间距

| 参数 | 默认值 | 说明 |
|------|--------|------|
| 节点宽度 | 200px | 可通过 `node_width` 参数调整 |
| 节点高度 | 60px | 可通过 `node_height` 参数调整 |
| 水平间距 | 60px | 节点之间的水平间距 |
| 垂直间距 | 80px | 节点之间的垂直间距 |
|| 判断节点 | 160×100px | 菱形比一般节点大 |
| 字号 | 14px | 节点内文字大小 |
|| 标题字号 | 20px | 流程图标题 |

## 数学公式标签

SVG 输出支持简单的数学标记（非完整 LaTeX，但覆盖常用符号）：

```python
# 上下标用 HTML 标签
fc.process("计算 R² 值", phase="analysis")
fc.process("求解 ∂L/∂w", phase="analysis")
fc.process("σ(x) = 1/(1+e⁻ˣ)", phase="analysis")
```

## Common Pitfalls

1. **SVG 中文显示问题**：确保目标系统有中文字体（微软雅黑/宋体），SVG 使用 `font-family: 'Microsoft YaHei', SimSun, sans-serif`。

2. **VSDX 需要 Visio 已安装**：COM 自动化需要 Visio 桌面版。Office 365 网页版不支持。如未安装 Visio，使用 SVG 或 PPTX 输出。

3. **菱形判断节点文字溢出**：判断节点文字建议不超过 8 个汉字。过长会导致文字溢出菱形边界。用换行符 `\n` 分行。

4. **回环连线布局**：从下方节点连回上方节点时，SVG 需要手动调整路径。使用 `style="dashed"` 区分回环连线。

5. **论文投稿分辨率**：SVG 是矢量格式，无分辨率问题。但如果期刊要求位图，用浏览器截图或 Inkscape 导出 300dpi PNG。

6. **Word 中插入 SVG**：Word 2016+ 支持直接插入 SVG。如果期刊要求 EMF/WMF，用 Inkscape 转换。

7. **PPTX 流程图编辑性**：python-pptx 生成的是形状（shapes），不是 SmartArt。在 PowerPoint 中可以手动移动和修改。

## Verification Checklist

- [ ] 所有节点都有明确的文字标签
- [ ] 判断节点有"是/否"分支标签
- [ ] 连线方向一致（自上而下或自左而右）
- [ ] 回环连线使用虚线区分
- [ ] 颜色按研究阶段语义化使用
- [ ] 流程图有标题
- [ ] SVG 在浏览器中打开正常显示
- [ ] 中文字符正常渲染（无方框）
- [ ] 节点文字不溢出边界
