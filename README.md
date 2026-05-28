# Scientific Research Flowchart Generator (visio-scientific-flowchart)

A Hermes Agent Skill for generating publication-quality scientific research flowcharts. Supports **three output formats**: SVG (universal), VSDX (Visio COM), and PPTX (PowerPoint).

> **[中文版 → README.zh.md](README.zh.md)**

## ✨ Features

| Feature | Description |
|---------|-------------|
| 📊 ISO 5807 Symbols | Process, Decision, I/O, Terminal, Subprocess, Database, Annotation |
| 🎨 Research Phase Colors | 6 semantic color schemes (data/preprocess/analysis/check/verify/result) |
| 📐 Auto Layout | Vertical (top-to-bottom) and horizontal (left-to-right) |
| 🔀 Branching | Decision nodes with labeled Yes/No branches, dashed loop-back arrows |
| 📝 Multi-format Output | SVG (vector), VSDX (Visio), PPTX (PowerPoint) |
| 🌏 CJK Support | Chinese/Japanese/Korean text in all formats |

## 🚀 Quick Start

### SVG Output (No dependencies)

```bash
# Pure Python standard library — nothing to install!
python templates/svg_flowchart.py
```

### PPTX Output

```bash
pip install python-pptx
python templates/pptx_flowchart.py
```

### VSDX Output (Requires Visio)

```bash
pip install pywin32
python templates/visio_com.py
```

## Usage

```python
from templates.svg_flowchart import FlowchartBuilder

fc = FlowchartBuilder("my_flow.svg", title="ML Research Pipeline")

fc.start("Start")
fc.io_input("Load Dataset", phase="data")
fc.process("Feature Engineering", phase="preprocess")
fc.decision("Quality Check?", phase="check")
fc.process("Model Training", phase="analysis")
fc.subprocess("Cross Validation", phase="verify")
fc.io_output("Report Results", phase="result")
fc.end("End")

fc.connect_all()   # Auto-connect by order
fc.save()
```

### Advanced: Custom Connections with Labels

```python
n1 = fc.start("开始")
n2 = fc.decision("准确率 ≥ 90%？", phase="check")
n3 = fc.io_output("输出报告", phase="result")
n4 = fc.process("调参重训", phase="analysis")

fc.connect(n1, n2)
fc.connect(n2, n3, label="是")      # Yes branch
fc.connect(n2, n4, label="否")      # No branch
fc.connect(n4, n2, style="dashed")  # Loop back
```

## Color Schemes by Research Phase

| Phase | Color | Use For |
|-------|-------|---------|
| `data` | 🔵 Blue | Data loading, cleaning, preprocessing |
| `preprocess` | 🟣 Purple | Feature engineering, dimensionality reduction |
| `analysis` | 🟢 Green | Model training, parameter optimization |
| `check` | 🟡 Yellow | Quality checks, hypothesis testing |
| `verify` | 🟠 Orange | Cross-validation, robustness checks |
| `result` | 🩷 Pink | Results, reports, publications |

## ISO 5807 Flowchart Symbols

| Symbol | Shape | Method | Purpose |
|--------|-------|--------|---------|
| Process | Rectangle | `process()` | General processing step |
| Decision | Diamond | `decision()` | Conditional branch |
| Input/Output | Parallelogram | `io_input()` / `io_output()` | Data I/O |
| Terminal | Rounded rect | `start()` / `end()` | Flow start/end |
| Subprocess | Double-border rect | `subprocess()` | Sub-procedure |
| Database | Cylinder | `database()` | Data storage |
| Annotation | Folded corner | `annotation()` | Notes |

## Output Format Comparison

| Format | Editable | Dependencies | Best For |
|--------|----------|-------------|----------|
| **SVG** | ✏️ In Inkscape/Browser | None (stdlib only) | Papers, Word, LaTeX |
| **VSDX** | ✏️ In Visio | Visio + pywin32 | Visio users |
| **PPTX** | ✏️ In PowerPoint | python-pptx | Presentations |

## 📁 Structure

```
visio-scientific-flowchart/
├── SKILL.md                          # Skill definition
├── README.md                         # English docs
├── README.zh.md                      # Chinese docs
├── LICENSE                           # MIT + AI declaration
├── requirements.txt                  # Python dependencies
├── .gitignore
├── templates/
│   ├── __init__.py
│   ├── svg_flowchart.py              # SVG generator (main, no deps)
│   ├── visio_com.py                  # Visio COM automation
│   └── pptx_flowchart.py             # PowerPoint generator
└── 示例_深度学习流程.svg              # Demo SVG output
```

## Install as Hermes Agent Skill

```bash
git clone https://github.com/Dynamic-braking9/visio-scientific-flowchart.git     ~/.hermes/skills/creative/visio-scientific-flowchart
```

## License

MIT — see [LICENSE](LICENSE)
