#!/usr/bin/env python3
"""
Visio COM 自动化流程图生成器
需要: Windows + Microsoft Visio 已安装 + pip install pywin32

用法:
    from visio_com import VisioFlowchart
    vf = VisioFlowchart("output.vsdx", title="实验流程")
    vf.add_process("数据采集", phase="data")
    vf.save()
"""

try:
    import win32com.client
    import pythoncom
    HAS_COM = True
except ImportError:
    HAS_COM = False

import os
from typing import Optional


# ── Visio 主题颜色（RGB 值）──────────────────────

PHASE_STENCILS = {
    "data":       {"fill": (219, 234, 254), "text": (31, 41, 55)},
    "preprocess": {"fill": (237, 233, 254), "text": (31, 41, 55)},
    "analysis":   {"fill": (209, 250, 229), "text": (31, 41, 55)},
    "check":      {"fill": (254, 243, 199), "text": (31, 41, 55)},
    "verify":     {"fill": (255, 237, 213), "text": (31, 41, 55)},
    "result":     {"fill": (252, 231, 243), "text": (31, 41, 55)},
}


class VisioFlowchart:
    """
    通过 COM 接口直接操控 Microsoft Visio 生成流程图。

    需要:
    - Windows 系统
    - Microsoft Visio 已安装（桌面版）
    - pip install pywin32
    """

    # Visio 形状主控形状名（来自 Basic Flowchart Shapes stencil）
    SHAPES = {
        "terminal":   "Terminator",
        "process":    "Process",
        "decision":   "Decision",
        "io":         "Data",
        "subprocess": "Subprocess",
        "database":   "Stored Data",
        "document":   "Document",
    }

    def __init__(self, filepath: str, title: str = "",
                 direction: str = "vertical",
                 page_width: float = 11.0, page_height: float = 8.5):
        """
        Args:
            filepath: 输出 .vsdx 文件路径
            title: 流程图标题
            direction: vertical (自上而下) 或 horizontal (自左而右)
            page_width: 页面宽度（英寸）
            page_height: 页面高度（英寸）
        """
        if not HAS_COM:
            raise ImportError("需要安装 pywin32: pip install pywin32")

        self.filepath = os.path.abspath(filepath)
        self.title = title
        self.direction = direction
        self.page_w = page_width
        self.page_h = page_height

        self.visio = None
        self.doc = None
        self.page = None
        self.stencil = None
        self.shapes = []
        self._shape_ids = []
        self._initialized = False

    def _init_visio(self):
        """初始化 Visio COM 连接"""
        if self._initialized:
            return

        pythoncom.CoInitialize()
        self.visio = win32com.client.Dispatch("Visio.Application")
        self.visio.Visible = False  # 后台运行

        # 新建文档（使用基本流程图模板）
        template = os.path.join(
            self.visio.GetSetting(7),  # visio 文档路径
            "Flowchart", "Basic Flowchart.vstx"
        )
        if os.path.exists(template):
            self.doc = self.visio.Documents.Add(template)
        else:
            self.doc = self.visio.Documents.Add("")

        self.page = self.visio.ActivePage
        self.page.PageSheet.CellsU("PageWidth").FormulaU = f"{self.page_w} in"
        self.page.PageSheet.CellsU("PageHeight").FormulaU = f"{self.page_h} in"

        # 尝试打开流程图模具
        try:
            stencil_path = os.path.join(
                self.visio.GetSetting(7),
                "Flowchart", "BASFLO_M.vssx"
            )
            if os.path.exists(stencil_path):
                self.stencil = self.visio.Documents.OpenEx(
                    stencil_path, 4  # visOpenDocked
                )
        except Exception:
            pass

        # 添加标题
        if self.title:
            title_shape = self.page.DrawText(
                self.title, 1, self.page_h - 0.5, 10, self.page_h - 0.2
            )
            title_shape.CellsU("Char.Size").FormulaU = "14pt"
            title_shape.CellsU("Char.Style").FormulaU = "1"  # bold

        self._initialized = True
        self._next_x = 1.0
        self._next_y = self.page_h - 1.5

    def _get_stencil_shape(self, shape_type: str):
        """从模具获取形状"""
        master_name = self.SHAPES.get(shape_type, "Process")

        # 从已加载的模具中查找
        if self.stencil:
            try:
                return self.stencil.Masters.ItemU(master_name)
            except Exception:
                pass

        # 从所有已加载的模具中查找
        for doc in self.visio.Documents:
            try:
                return doc.Masters.ItemU(master_name)
            except Exception:
                continue

        return None

    def _add_shape(self, text: str, shape_type: str = "process",
                   phase: str = "", x: float = None, y: float = None,
                   w: float = 2.0, h: float = 0.5) -> object:
        """添加形状到页面"""
        self._init_visio()

        if x is None:
            x = self._next_x
        if y is None:
            y = self._next_y

        # 尝试从模具获取形状
        master = self._get_stencil_shape(shape_type)

        if master:
            shape = self.page.Drop(master, x + w/2, y - h/2)
            shape.CellsU("Width").FormulaU = f"{w} in"
            shape.CellsU("Height").FormulaU = f"{h} in"
        else:
            # 降级：用 DrawRectangle 绘制基本矩形
            shape = self.page.DrawRectangle(x, y - h, x + w, y)

        # 设置文字
        shape.Text = text
        shape.CellsU("Para.HorzAlign").FormulaU = "1"  # 居中

        # 设置颜色
        if phase and phase in PHASE_STENCILS:
            c = PHASE_STENCILS[phase]
            fill_rgb = c["fill"]
            shape.CellsU("FillForegnd").FormulaU = (
                f'RGB({fill_rgb[0]},{fill_rgb[1]},{fill_rgb[2]})'
            )

        self.shapes.append(shape)
        self._shape_ids.append(shape.ID)

        # 更新下一个位置
        if self.direction == "vertical":
            self._next_y = y - h - 0.5
        else:
            self._next_x = x + w + 0.3

        return shape

    # ── 公开 API ──────────────────────────────

    def add_start(self, text: str = "开始") -> object:
        return self._add_shape(text, "terminal", w=1.5, h=0.4)

    def add_end(self, text: str = "结束") -> object:
        return self._add_shape(text, "terminal", w=1.5, h=0.4)

    def add_process(self, text: str, phase: str = "analysis") -> object:
        return self._add_shape(text, "process", phase=phase)

    def add_decision(self, text: str, phase: str = "check") -> object:
        return self._add_shape(text, "decision", phase=phase, w=2.0, h=0.8)

    def add_io(self, text: str, phase: str = "data") -> object:
        return self._add_shape(text, "io", phase=phase)

    def add_subprocess(self, text: str, phase: str = "analysis") -> object:
        return self._add_shape(text, "subprocess", phase=phase)

    def add_database(self, text: str, phase: str = "data") -> object:
        return self._add_shape(text, "database", phase=phase, h=0.6)

    def connect_shapes(self, src_idx: int, dst_idx: int,
                       label: str = ""):
        """连接两个形状（按添加顺序的索引）"""
        self._init_visio()
        if src_idx >= len(self.shapes) or dst_idx >= len(self.shapes):
            raise IndexError("形状索引越界")

        src = self.shapes[src_idx]
        dst = self.shapes[dst_idx]

        # 使用 Visio 的自动连接
        connector = self.page.DrawLine(0, 0, 1, 1)
        connector.CellsU("BeginX").GlueTo(src.CellsU("PinX"))
        connector.CellsU("EndX").GlueTo(dst.CellsU("PinX"))

        # 添加箭头
        connector.CellsU("EndArrow").FormulaU = "5"  # 实心箭头

        if label:
            connector.Text = label

    def auto_connect(self):
        """按添加顺序自动连接所有形状"""
        for i in range(len(self.shapes) - 1):
            self.connect_shapes(i, i + 1)

    def add_annotation(self, text: str) -> object:
        """添加注释节点"""
        return self._add_shape(text, "annotation", w=2.0, h=0.35)

    def save(self):
        """保存 Visio 文件"""
        self._init_visio()

        # 保存为 .vsdx 格式
        save_path = self.filepath
        if not save_path.endswith('.vsdx'):
            save_path += '.vsdx'

        try:
            self.doc.SaveAs(save_path)
            print(f"✅ Visio 文件已生成: {save_path}")
        finally:
            try:
                self.visio.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()
            self._initialized = False


# ── 演示 ─────────────────────────────────────────

def demo():
    """演示：生成深度学习研究流程的 Visio 文件"""
    vf = VisioFlowchart("示例_深度学习流程.vsdx",
                        title="基于深度学习的图像分类研究流程")

    vf.add_start("开始")
    vf.add_io("CIFAR-10 数据集加载", phase="data")
    vf.add_process("数据增强与预处理", phase="data")
    vf.add_decision("数据质量检查？", phase="check")
    vf.add_process("ResNet-50 模型构建", phase="analysis")
    vf.add_process("模型训练 (Adam)", phase="analysis")
    vf.add_decision("准确率 ≥ 90%？", phase="check")
    vf.add_subprocess("交叉验证", phase="verify")
    vf.add_io("实验报告", phase="result")
    vf.add_end("结束")

    vf.auto_connect()
    vf.save()


if __name__ == "__main__":
    demo()
