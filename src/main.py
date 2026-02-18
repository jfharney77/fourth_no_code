"""
Generate a PowerPoint slide from hierarchy Excel data.

Drawing strategy: all connectors are drawn first, then all circles, so that
connectors always appear behind circles regardless of layout mode.
"""

import math
import yaml
import openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parent.parent


def load_config():
    config_path = ROOT_DIR / "config" / "config.yaml"
    with open(config_path, "r") as f:
        return yaml.safe_load(f)


def hex_to_rgb(hex_str):
    return RGBColor(
        int(hex_str[0:2], 16),
        int(hex_str[2:4], 16),
        int(hex_str[4:6], 16),
    )


def read_sheet_mapping(wb, sheet_name, header_row, parent_col, child_col):
    """Read a sheet and return an ordered dict of parent -> [children]."""
    ws = wb[sheet_name]
    mapping = {}
    for row in range(header_row + 1, ws.max_row + 1):
        parent = ws[f"{parent_col}{row}"].value
        child = ws[f"{child_col}{row}"].value
        if parent and child:
            mapping.setdefault(parent, []).append(child)
    return mapping


def read_all_data(config):
    """Read L12A and A2T sheets, return l1_to_agents and agent_to_tools."""
    wb = openpyxl.load_workbook(ROOT_DIR / config["input_file"], data_only=True)

    l12a_cfg = config["sheets"]["L12A"]
    l1_to_agents = read_sheet_mapping(
        wb, "L12A", l12a_cfg["header_row"],
        l12a_cfg["l1_column"], l12a_cfg["agents_column"],
    )

    a2t_cfg = config["sheets"]["A2T"]
    agent_to_tools = read_sheet_mapping(
        wb, "A2T", a2t_cfg["header_row"],
        a2t_cfg["agents_column"], a2t_cfg["tools_column"],
    )

    wb.close()
    return l1_to_agents, agent_to_tools


# ---------------------------------------------------------------------------
# Low-level drawing primitives
# ---------------------------------------------------------------------------

def draw_line(slide, x1, y1, x2, y2, color, weight):
    """Add a straight connector line between two points."""
    connector = slide.shapes.add_connector(1, x1, y1, x2, y2)
    connector.line.color.rgb = color
    connector.line.width = weight


def draw_circle(slide, center_x, center_y, size_emu, fill_color, font_size,
                font_color, text, abbrev_chars=0, label_font_size=None,
                label_font_color=None):
    """Add a circle centred at (center_x, center_y) with optional label below."""
    half = size_emu // 2
    shape = slide.shapes.add_shape(
        9, center_x - half, center_y - half, size_emu, size_emu,
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color

    needs_label = abbrev_chars > 0 and len(text) > abbrev_chars
    inside_text = text[:abbrev_chars] if needs_label else text

    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = inside_text
    p.alignment = PP_ALIGN.CENTER
    p.font.size = font_size
    p.font.color.rgb = font_color

    if needs_label:
        label_width = Inches(1.5)
        label_height = Inches(0.3)
        label_left = center_x - label_width // 2
        label_top = center_y + half + Inches(0.02)
        tb = slide.shapes.add_textbox(label_left, label_top, label_width, label_height)
        tf_label = tb.text_frame
        tf_label.word_wrap = True
        p_label = tf_label.paragraphs[0]
        p_label.text = text
        p_label.alignment = PP_ALIGN.CENTER
        p_label.font.size = label_font_size or font_size
        p_label.font.color.rgb = label_font_color or font_color


def _make_circle_style(cfg):
    """Build a style dict from a circle config section."""
    style = {
        "size": Inches(cfg["size_inches"]),
        "fill_color": hex_to_rgb(cfg["fill_color"]),
        "font_size": Pt(cfg["font_size"]),
        "font_color": hex_to_rgb(cfg.get("font_color", "000000")),
        "abbrev_chars": cfg.get("abbrev_chars", 0),
    }
    if cfg.get("label_font_size"):
        style["label_font_size"] = Pt(cfg["label_font_size"])
    if cfg.get("label_font_color"):
        style["label_font_color"] = hex_to_rgb(cfg["label_font_color"])
    return style


# ---------------------------------------------------------------------------
# Batch drawing helpers — work on collected positions
# ---------------------------------------------------------------------------

def _draw_all_connectors(slide, connectors):
    """Draw all connector lines. Each entry: (x1, y1, x2, y2, color, weight)."""
    for x1, y1, x2, y2, color, weight in connectors:
        draw_line(slide, x1, y1, x2, y2, color, weight)


def _draw_all_circles(slide, circles):
    """Draw all circles. Each entry: (cx, cy, style, text)."""
    for cx, cy, style, text in circles:
        draw_circle(slide, cx, cy, style["size"], style["fill_color"],
                    style["font_size"], style["font_color"], text,
                    abbrev_chars=style.get("abbrev_chars", 0),
                    label_font_size=style.get("label_font_size"),
                    label_font_color=style.get("label_font_color"))


# ---------------------------------------------------------------------------
# Position calculation — no drawing, just returns positions + draw lists
# ---------------------------------------------------------------------------

def _calc_surrounding_positions(parent_to_children, parent_centers, center_dist):
    """Calculate child positions orbiting around parent centers.
    Returns dict of child_name -> (cx, cy).
    """
    child_centers = {}
    for parent_name, children in parent_to_children.items():
        if parent_name not in parent_centers:
            continue
        px, py = parent_centers[parent_name]
        n = len(children)
        for j, child_name in enumerate(children):
            angle = 2 * math.pi * j / n - math.pi / 2
            child_centers[child_name] = (
                px + int(center_dist * math.cos(angle)),
                py + int(center_dist * math.sin(angle)),
            )
    return child_centers


def _calc_row_positions(parent_to_children, row_y, child_size, slide_width):
    """Calculate child positions in a centered horizontal row.
    Returns dict of child_name -> (cx, cy) and list of (parent_name, child_name).
    """
    all_children = []
    for parent_name, children in parent_to_children.items():
        for child_name in children:
            all_children.append((parent_name, child_name))

    if not all_children:
        return {}, []

    padding = Inches(0.3)
    total_width = len(all_children) * child_size + (len(all_children) - 1) * padding
    start_x = (slide_width - total_width) // 2 + child_size // 2

    child_centers = {}
    for i, (parent_name, child_name) in enumerate(all_children):
        cx = start_x + i * (child_size + padding)
        child_centers[child_name] = (cx, row_y)

    return child_centers, all_children


def _calc_l1_agents_surrounding(l1_to_agents, sheet_cfg, agent_size):
    """Calculate L1 and agent positions for surrounding layout.
    Returns l1_centers, agent_centers.
    """
    center_dist = Inches(sheet_cfg["center_distance_inches"])
    cluster_width = center_dist * 2 + agent_size
    cluster_padding = Inches(0.3)
    start_x = Inches(1.0)
    center_y = Inches(3.0)

    l1_centers = {}
    agent_centers = {}

    for i, (l1_name, agents) in enumerate(l1_to_agents.items()):
        cx = start_x + cluster_width // 2 + i * (cluster_width + cluster_padding)
        cy = center_y
        l1_centers[l1_name] = (cx, cy)

        n = len(agents)
        for j, agent_name in enumerate(agents):
            angle = 2 * math.pi * j / n - math.pi / 2
            agent_centers[agent_name] = (
                cx + int(center_dist * math.cos(angle)),
                cy + int(center_dist * math.sin(angle)),
            )

    return l1_centers, agent_centers


def _calc_l1_agents_bottom(l1_to_agents, l1_size, agent_size):
    """Calculate L1 and agent positions for bottom layout.
    Returns l1_centers, agent_centers, agent_parent_list.
    """
    slide_width = Inches(10.0)

    # L1 row
    l1_names = list(l1_to_agents.keys())
    l1_padding = Inches(0.5)
    total_l1_width = len(l1_names) * l1_size + (len(l1_names) - 1) * l1_padding
    l1_start_x = (slide_width - total_l1_width) // 2 + l1_size // 2
    l1_y = Inches(1.5)

    l1_centers = {}
    for i, name in enumerate(l1_names):
        l1_centers[name] = (l1_start_x + i * (l1_size + l1_padding), l1_y)

    # Agent row
    all_agents = []
    for l1_name, agents in l1_to_agents.items():
        for agent_name in agents:
            all_agents.append((l1_name, agent_name))

    agent_padding = Inches(0.3)
    total_agent_width = len(all_agents) * agent_size + (len(all_agents) - 1) * agent_padding
    agent_start_x = (slide_width - total_agent_width) // 2 + agent_size // 2
    agent_y = Inches(4.5)

    agent_centers = {}
    for i, (l1_name, agent_name) in enumerate(all_agents):
        agent_centers[agent_name] = (
            agent_start_x + i * (agent_size + agent_padding), agent_y,
        )

    return l1_centers, agent_centers, all_agents


# ---------------------------------------------------------------------------
# Connector collection helpers
# ---------------------------------------------------------------------------

def _collect_connectors(parent_centers, child_centers, parent_to_children,
                        conn_color, conn_weight):
    """Build a list of connector tuples from parent->child mappings."""
    connectors = []
    for parent_name, children in parent_to_children.items():
        if parent_name not in parent_centers:
            continue
        px, py = parent_centers[parent_name]
        for child_name in children:
            if child_name in child_centers:
                cx, cy = child_centers[child_name]
                connectors.append((px, py, cx, cy, conn_color, conn_weight))
    return connectors


# ---------------------------------------------------------------------------
# Main build
# ---------------------------------------------------------------------------

def build_slide(config, l1_to_agents, agent_to_tools):
    """Create a PowerPoint with circles for L1, agents, and tools.

    All connectors are drawn first, then all circles, ensuring connectors
    are always behind circles.
    """
    l12a_cfg = config["sheets"]["L12A"]
    l12a_conn = l12a_cfg["connector"]
    l1_style = _make_circle_style(l12a_cfg["l1_circle"])
    agent_style = _make_circle_style(l12a_cfg["agent_circle"])
    l12a_draw_conn = l12a_conn.get("enabled", True)
    l12a_conn_color = hex_to_rgb(l12a_conn["color"])
    l12a_conn_weight = Pt(l12a_conn["weight_pt"])

    a2t_cfg = config["sheets"]["A2T"]
    a2t_conn = a2t_cfg["connector"]
    tool_style = _make_circle_style(a2t_cfg["tool_circle"])
    a2t_draw_conn = a2t_conn.get("enabled", True)
    a2t_conn_color = hex_to_rgb(a2t_conn["color"])
    a2t_conn_weight = Pt(a2t_conn["weight_pt"])

    slide_width = Inches(10.0)

    # ---- Phase 1: Calculate all positions ----

    agent_layout = l12a_cfg.get("agent_layout", "bottom")
    if agent_layout == "surrounding":
        l1_centers, agent_centers = _calc_l1_agents_surrounding(
            l1_to_agents, l12a_cfg, agent_style["size"])
    else:
        l1_centers, agent_centers, _ = _calc_l1_agents_bottom(
            l1_to_agents, l1_style["size"], agent_style["size"])

    tool_layout = a2t_cfg.get("tool_layout", "top")
    if tool_layout == "surrounding":
        center_dist = Inches(a2t_cfg["center_distance_inches"])
        tool_centers = _calc_surrounding_positions(
            agent_to_tools, agent_centers, center_dist)
    else:
        tool_centers, _ = _calc_row_positions(
            agent_to_tools, row_y=Inches(0.5),
            child_size=tool_style["size"], slide_width=slide_width)

    # ---- Phase 2: Collect all connectors ----

    all_connectors = []
    if l12a_draw_conn:
        all_connectors += _collect_connectors(
            l1_centers, agent_centers, l1_to_agents,
            l12a_conn_color, l12a_conn_weight)
    if a2t_draw_conn:
        all_connectors += _collect_connectors(
            agent_centers, tool_centers, agent_to_tools,
            a2t_conn_color, a2t_conn_weight)

    # ---- Phase 3: Collect all circles ----

    all_circles = []
    for name, (cx, cy) in l1_centers.items():
        all_circles.append((cx, cy, l1_style, name))
    for name, (cx, cy) in agent_centers.items():
        all_circles.append((cx, cy, agent_style, name))
    for name, (cx, cy) in tool_centers.items():
        all_circles.append((cx, cy, tool_style, name))

    # ---- Phase 4: Draw — connectors first, then circles ----

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    _draw_all_connectors(slide, all_connectors)
    _draw_all_circles(slide, all_circles)

    output_path = ROOT_DIR / config["output_file"]
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


def main():
    config = load_config()
    l1_to_agents, agent_to_tools = read_all_data(config)

    print("L1 -> Agents:")
    for l1, agents in l1_to_agents.items():
        print(f"  {l1}: {agents}")

    print("Agent -> Tools:")
    for agent, tools in agent_to_tools.items():
        print(f"  {agent}: {tools}")

    output = build_slide(config, l1_to_agents, agent_to_tools)
    print(f"Saved to {output}")


if __name__ == "__main__":
    main()
