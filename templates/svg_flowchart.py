#!/usr/bin/env python3
"""
科研流程图 SVG 生成器
纯 Python 标准库实现，无第三方依赖。
支持 ISO 5807 标准符号、研究阶段配色、自动布局。

用法:
    from svg_flowchart import FlowchartBuilder
    fc = FlowchartBuilder("output.svg", title="研究流程")
    fc.start("开始")
    fc.process("数据采集", phase="data")
    fc.save()
"""

import math
from dataclasses import dataclass, field
from typing import Optional
from xml.sax.saxutils import escape


# ── 配色方案 ────────────────────────────────────

PHASE_COLORS = {
    "data":       {"fill": "#DBEAFE", "stroke": "#3B82F6", "label": "数据"},
    "preprocess": {"fill": "#EDE9FE", "stroke": "#8B5CF6", "label": "预处理"},
    "analysis":   {"fill": "#D1FAE5", "stroke": "#10B981", "label": "分析"},
    "check":      {"fill": "#FEF3C7", "stroke": "#F59E0B", "label": "判断"},
    "verify":     {"fill": "#FFEDD5", "stroke": "#F97316", "label": "验证"},
    "result":     {"fill": "#FCE7F3", "stroke": "#EC4899", "label": "结果"},
}

TERMINAL_COLORS = {"fill": "#F3F4F6", "stroke": "#6B7280"}
IO_COLORS        = {"fill": "#CFFAFE", "stroke": "#06B6D4"}
SUBPROC_COLORS   = {"fill": "#FEF9C3", "stroke": "#CA8A04"}

FONT_FAMILY = "'Microsoft YaHei', 'SimSun', 'PingFang SC', 'Noto Sans CJK SC', sans-serif"


# ── 节点数据结构 ────────────────────────────────

@dataclass
class Node:
    id: str
    text: str
    shape: str          # process | decision | io | terminal | subprocess | database | annotation
    phase: str = ""
    x: float = 0
    y: float = 0
    w: float = 200
    h: float = 60
    direction: str = "input"  # for io shape


@dataclass
class Edge:
    src: str
    dst: str
    label: str = ""
    style: str = "solid"   # solid | dashed | dotted


# ── SVG 生成器 ───────────────────────────────────

class FlowchartBuilder:
    """科研流程图构建器，输出 SVG 矢量图。"""

    def __init__(self, filepath: str, title: str = "",
                 direction: str = "vertical",
                 node_width: float = 200, node_height: float = 60,
                 h_gap: float = 60, v_gap: float = 80,
                 font_size: float = 14, title_font_size: float = 20,
                 layout: str = "auto", lanes: list = None):
        self.filepath = filepath
        self.title = title
        self.direction = direction    # vertical | horizontal
        self.node_w = node_width
        self.node_h = node_height
        self.h_gap = h_gap
        self.v_gap = v_gap
        self.font_size = font_size
        self.title_fs = title_font_size
        self.layout_type = layout     # auto | swimlane
        self.lanes = lanes or []

        self.nodes: list[Node] = []
        self.edges: list[Edge] = []
        self._id_counter = 0
        self._node_map: dict[str, Node] = {}

    # ── ID 生成 ──────────────────────────────

    def _next_id(self, prefix: str = "n") -> str:
        self._id_counter += 1
        return f"{prefix}{self._id_counter}"

    # ── 节点添加 ──────────────────────────────

    def _add_node(self, text: str, shape: str, phase: str = "",
                  w: float = None, h: float = None, **kwargs) -> Node:
        nid = self._next_id()
        node = Node(
            id=nid, text=text, shape=shape, phase=phase,
            w=w or self.node_w, h=h or self.node_h, **kwargs
        )
        self.nodes.append(node)
        self._node_map[nid] = node
        return node

    def start(self, text: str = "开始") -> Node:
        return self._add_node(text, "terminal", w=140, h=50)

    def end(self, text: str = "结束") -> Node:
        return self._add_node(text, "terminal", w=140, h=50)

    def process(self, text: str, phase: str = "analysis") -> Node:
        return self._add_node(text, "process", phase=phase)

    def decision(self, text: str, phase: str = "check") -> Node:
        return self._add_node(text, "decision", phase=phase, w=160, h=100)

    def io_input(self, text: str, phase: str = "data") -> Node:
        return self._add_node(text, "io", phase=phase, direction="input")

    def io_output(self, text: str, phase: str = "result") -> Node:
        return self._add_node(text, "io", phase=phase, direction="output")

    def subprocess(self, text: str, phase: str = "analysis") -> Node:
        return self._add_node(text, "subprocess", phase=phase)

    def database(self, text: str, phase: str = "data") -> Node:
        return self._add_node(text, "database", phase=phase, h=70)

    def annotation(self, text: str) -> Node:
        return self._add_node(text, "annotation", w=180, h=45)

    # ── 连线 ──────────────────────────────────

    def connect(self, src: Node, dst: Node, label: str = "",
                style: str = "solid") -> Edge:
        edge = Edge(src=src.id, dst=dst.id, label=label, style=style)
        self.edges.append(edge)
        return edge

    def connect_all(self):
        """按节点添加顺序自动连线（跳过已手动连接的节点对）"""
        existing = {(e.src, e.dst) for e in self.edges}
        for i in range(len(self.nodes) - 1):
            pair = (self.nodes[i].id, self.nodes[i + 1].id)
            if pair not in existing:
                self.edges.append(Edge(src=pair[0], dst=pair[1]))

    # ── 布局计算 ──────────────────────────────

    def _auto_layout(self):
        """自动计算节点坐标"""
        if not self.nodes:
            return

        # 标题偏移
        title_offset = 60 if self.title else 20

        if self.direction == "vertical":
            cx = 400  # 水平居中
            y = title_offset
            for node in self.nodes:
                node.x = cx - node.w / 2
                node.y = y
                y += node.h + self.v_gap
        else:  # horizontal
            cy = 200  # 垂直居中
            x = 80
            for node in self.nodes:
                node.x = x
                node.y = cy - node.h / 2
                x += node.w + self.h_gap

    # ── SVG 图形绘制 ──────────────────────────

    def _get_colors(self, node: Node) -> dict:
        if node.shape == "terminal":
            return TERMINAL_COLORS
        if node.shape == "io":
            return IO_COLORS
        if node.shape == "subprocess":
            return SUBPROC_COLORS
        if node.phase and node.phase in PHASE_COLORS:
            return PHASE_COLORS[node.phase]
        return {"fill": "#F3F4F6", "stroke": "#9CA3AF"}

    def _draw_node_svg(self, node: Node) -> str:
        """生成单个节点的 SVG 元素"""
        colors = self._get_colors(node)
        fill = colors["fill"]
        stroke = colors["stroke"]
        x, y, w, h = node.x, node.y, node.w, node.h
        text_lines = node.text.split("\n")
        line_count = len(text_lines)
        line_h = self.font_size * 1.3
        text_block_h = line_count * line_h

        parts = []

        if node.shape == "terminal":
            # 圆角矩形（开始/结束）
            rx = h / 2
            parts.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" '
                         f'rx="{rx}" ry="{rx}" fill="{fill}" stroke="{stroke}" stroke-width="2"/>')

        elif node.shape == "decision":
            # 菱形
            cx, cy = x + w / 2, y + h / 2
            points = f"{cx},{y} {x + w},{cy} {cx},{y + h} {x},{cy}"
            parts.append(f'<polygon points="{points}" fill="{fill}" stroke="{stroke}" stroke-width="2"/>')

        elif node.shape == "io":
            # 平行四边形
            skew = 20
            if node.direction == "input":
                points = f"{x + skew},{y} {x + w},{y} {x + w - skew},{y + h} {x},{y + h}"
            else:
                points = f"{x},{y} {x + w - skew},{y} {x + w},{y + h} {x + skew},{y + h}"
            parts.append(f'<polygon points="{points}" fill="{fill}" stroke="{stroke}" stroke-width="2"/>')

        elif node.shape == "subprocess":
            # 双边框矩形
            parts.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" '
                         f'rx="4" ry="4" fill="{fill}" stroke="{stroke}" stroke-width="2"/>')
            # 内边框
            m = 6
            parts.append(f'<rect x="{x + m}" y="{y + m}" width="{w - 2*m}" height="{h - 2*m}" '
                         f'rx="2" ry="2" fill="none" stroke="{stroke}" stroke-width="1"/>')

        elif node.shape == "database":
            # 圆柱体（简化为矩形 + 椭圆顶）
            ew = w
            eh = 16
            parts.append(f'<rect x="{x}" y="{y + eh/2}" width="{w}" height="{h - eh}" '
                         f'fill="{fill}" stroke="{stroke}" stroke-width="2"/>')
            parts.append(f'<ellipse cx="{x + w/2}" cy="{y + eh/2}" rx="{w/2}" ry="{eh/2}" '
                         f'fill="{fill}" stroke="{stroke}" stroke-width="2"/>')
            parts.append(f'<ellipse cx="{x + w/2}" cy="{y + h - eh/2}" rx="{w/2}" ry="{eh/2}" '
                         f'fill="{fill}" stroke="{stroke}" stroke-width="2" '
                         f'stroke-dasharray="4,2" opacity="0.5"/>')

        elif node.shape == "annotation":
            # 折角矩形
            fold = 12
            path = (f'M{x},{y} L{x + w - fold},{y} L{x + w},{y + fold} '
                    f'L{x + w},{y + h} L{x},{y + h} Z')
            parts.append(f'<path d="{path}" fill="{fill}" stroke="{stroke}" stroke-width="1.5" '
                         f'stroke-dasharray="4,3"/>')
            # 折角
            parts.append(f'<polyline points="{x + w - fold},{y} {x + w - fold},{y + fold} {x + w},{y + fold}" '
                         f'fill="none" stroke="{stroke}" stroke-width="1.5"/>')

        else:  # process (default)
            parts.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" '
                         f'rx="6" ry="6" fill="{fill}" stroke="{stroke}" stroke-width="2"/>')

        # 绘制文字
        text_cx = x + w / 2
        text_start_y = y + (h - text_block_h) / 2 + self.font_size
        for i, line in enumerate(text_lines):
            ty = text_start_y + i * line_h
            parts.append(f'<text x="{text_cx}" y="{ty}" '
                         f'text-anchor="middle" dominant-baseline="auto" '
                         f'font-family="{FONT_FAMILY}" font-size="{self.font_size}" '
                         f'fill="#1F2937">{escape(line)}</text>')

        return "\n  ".join(parts)

    def _draw_edge_svg(self, edge: Edge) -> str:
        """生成连线的 SVG 元素"""
        src = self._node_map[edge.src]
        dst = self._node_map[edge.dst]

        # 计算连线起止点（节点中心连线与边界的交点）
        sx, sy = self._port(src, dst)
        ex, ey = self._port(dst, src)

        dash = ""
        if edge.style == "dashed":
            dash = ' stroke-dasharray="6,4"'
        elif edge.style == "dotted":
            dash = ' stroke-dasharray="3,3"'

        # 生成路径（带拐弯）
        path_d = self._edge_path(sx, sy, ex, ey, src, dst)

        parts = []
        parts.append(f'<path d="{path_d}" fill="none" stroke="#6B7280" '
                     f'stroke-width="1.8"{dash} marker-end="url(#arrowhead)"/>')

        # 连线标签
        if edge.label:
            mx = (sx + ex) / 2
            my = (sy + ey) / 2
            # 偏移标签避免与线重叠
            if abs(sx - ex) < 5:  # 垂直线
                mx += 12
            else:
                my -= 8
            parts.append(f'<rect x="{mx - 18}" y="{my - 12}" width="36" height="18" '
                         f'rx="3" fill="white" stroke="#D1D5DB" stroke-width="0.5"/>')
            parts.append(f'<text x="{mx}" y="{my + 2}" text-anchor="middle" '
                         f'font-family="{FONT_FAMILY}" font-size="11" '
                         f'fill="#6B7280">{escape(edge.label)}</text>')

        return "\n  ".join(parts)

    def _port(self, node_from: Node, node_to: Node) -> tuple:
        """计算从 node_from 到 node_to 方向的边界交点"""
        cx = node_from.x + node_from.w / 2
        cy = node_from.y + node_from.h / 2
        tx = node_to.x + node_to.w / 2
        ty = node_to.y + node_to.h / 2

        dx = tx - cx
        dy = ty - cy

        if abs(dx) < 1 and abs(dy) < 1:
            return cx, cy

        # 判断主要方向
        if self.direction == "vertical":
            if abs(dy) > abs(dx) * 0.3:
                if dy > 0:
                    return cx, node_from.y + node_from.h  # bottom
                else:
                    return cx, node_from.y  # top
            else:
                if dx > 0:
                    return node_from.x + node_from.w, cy  # right
                else:
                    return node_from.x, cy  # left
        else:
            if abs(dx) > abs(dy) * 0.3:
                if dx > 0:
                    return node_from.x + node_from.w, cy  # right
                else:
                    return node_from.x, cy  # left
            else:
                if dy > 0:
                    return cx, node_from.y + node_from.h  # bottom
                else:
                    return cx, node_from.y  # top

    def _edge_path(self, sx, sy, ex, ey, src_node, dst_node) -> str:
        """生成连线路径，带中间拐点"""
        if self.direction == "vertical":
            # 垂直布局：先垂直后水平再垂直
            mid_y = (sy + ey) / 2
            if abs(sx - ex) < 5:
                # 同列直连
                return f"M{sx},{sy} L{ex},{ey}"
            else:
                return f"M{sx},{sy} L{sx},{mid_y} L{ex},{mid_y} L{ex},{ey}"
        else:
            mid_x = (sx + ex) / 2
            if abs(sy - ey) < 5:
                return f"M{sx},{sy} L{ex},{ey}"
            else:
                return f"M{sx},{sy} L{mid_x},{sy} L{mid_x},{ey} L{ex},{ey}"

    # ── 生成完整 SVG ──────────────────────────

    def _calc_canvas_size(self) -> tuple:
        """计算画布尺寸"""
        if not self.nodes:
            return 800, 600
        max_x = max(n.x + n.w for n in self.nodes) + 80
        max_y = max(n.y + n.h for n in self.nodes) + 80
        return max(max_x, 400), max(max_y, 300)

    def _draw_legend(self, canvas_w: float) -> str:
        """绘制图例（颜色-阶段对照）"""
        used_phases = {n.phase for n in self.nodes if n.phase and n.phase in PHASE_COLORS}
        if not used_phases:
            return ""

        legend_y = max(n.y + n.h for n in self.nodes) + 50
        parts = []
        x = 30

        for phase in sorted(used_phases):
            c = PHASE_COLORS[phase]
            parts.append(f'<rect x="{x}" y="{legend_y}" width="16" height="16" '
                         f'rx="3" fill="{c["fill"]}" stroke="{c["stroke"]}" stroke-width="1.5"/>')
            parts.append(f'<text x="{x + 22}" y="{legend_y + 13}" '
                         f'font-family="{FONT_FAMILY}" font-size="12" fill="#4B5563">'
                         f'{c["label"]}</text>')
            x += 80

        return "\n  ".join(parts)

    def save(self):
        """生成并保存 SVG 文件"""
        self._auto_layout()
        cw, ch = self._calc_canvas_size()

        # 确保图例空间
        legend_y = max(n.y + n.h for n in self.nodes) + 70 if self.nodes else 0
        ch = max(ch, legend_y + 40)

        parts = []
        parts.append(f'<svg xmlns="http://www.w3.org/2000/svg" width="{cw}" height="{ch}" '
                     f'viewBox="0 0 {cw} {ch}">')

        # 样式定义
        parts.append('''  <defs>
    <marker id="arrowhead" markerWidth="10" markerHeight="7"
            refX="9" refY="3.5" orient="auto" fill="#6B7280">
      <polygon points="0 0, 10 3.5, 0 7"/>
    </marker>
  </defs>''')

        # 背景
        parts.append(f'  <rect width="{cw}" height="{ch}" fill="white"/>')

        # 标题
        if self.title:
            parts.append(f'  <text x="{cw/2}" y="35" text-anchor="middle" '
                         f'font-family="{FONT_FAMILY}" font-size="{self.title_fs}" '
                         f'font-weight="bold" fill="#111827">{escape(self.title)}</text>')

        # 绘制连线（在节点下方）
        for edge in self.edges:
            parts.append(f"  {self._draw_edge_svg(edge)}")

        # 绘制节点
        for node in self.nodes:
            parts.append(f"  {self._draw_node_svg(node)}")

        # 图例
        legend = self._draw_legend(cw)
        if legend:
            parts.append(f"  {legend}")

        parts.append("</svg>")

        svg_content = "\n".join(parts)
        with open(self.filepath, "w", encoding="utf-8") as f:
            f.write(svg_content)
        print(f"✅ 流程图已生成: {self.filepath}")


# ── 演示 ─────────────────────────────────────────

def demo():
    """演示：生成一份深度学习研究流程图"""
    fc = FlowchartBuilder("示例_深度学习流程.svg",
                          title="基于深度学习的图像分类研究流程",
                          direction="vertical")

    n1 = fc.start("开始")
    n2 = fc.io_input("CIFAR-10 数据集加载", phase="data")
    n3 = fc.process("数据增强\n(旋转/翻转/裁剪)", phase="data")
    n4 = fc.process("归一化与标准化", phase="preprocess")
    n5 = fc.decision("数据质量\n检查通过？", phase="check")
    n6 = fc.process("CNN 模型构建\n(ResNet-50)", phase="analysis")
    n7 = fc.process("模型训练\n(Adam, lr=1e-3)", phase="analysis")
    n8 = fc.decision("准确率\n≥ 90%？", phase="check")
    n9 = fc.subprocess("5-折交叉验证", phase="verify")
    n10 = fc.process("混淆矩阵分析", phase="verify")
    n11 = fc.io_output("实验报告\n(准确率 93.2%)", phase="result")
    n12 = fc.end("结束")

    # 自定义连线
    fc.connect(n1, n2)
    fc.connect(n2, n3)
    fc.connect(n3, n4)
    fc.connect(n4, n5)
    fc.connect(n5, n6, label="是")
    fc.connect(n5, n3, label="否", style="dashed")
    fc.connect(n6, n7)
    fc.connect(n7, n8)
    fc.connect(n8, n9, label="否")
    fc.connect(n8, n11, label="是")
    fc.connect(n9, n10)
    fc.connect(n10, n7, label="调整", style="dashed")
    fc.connect(n11, n12)

    fc.save()


if __name__ == "__main__":
    demo()
