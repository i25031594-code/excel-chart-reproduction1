# -*- coding: utf-8 -*-
"""图18 跑道图：严格读取 Excel 数据，完整显示中文标题、标签、图例与脚注。"""
from pathlib import Path

import numpy as np
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.font_manager import FontProperties
from matplotlib.patches import Circle, Wedge, FancyBboxPatch, Rectangle
from openpyxl import load_workbook

EXCEL_FILE = "第二章 图表(后15).xlsx"
SHEET_NAME = "18 跑道图"
OUTPUT_FILE = "图18_跑道图.png"

BG = "#191E45"
WHITE = "#F7F8FC"
TEXT2 = "#D7DBEA"
MUTED = "#AEB5CE"
GRID = "#4B5277"
BLUE = "#087CC1"
BRIGHT_BLUE = "#1597E5"
CYAN = "#20A8B5"
LIGHT_BLUE = "#45A5D6"
PINK = "#EB4D6D"
YELLOW = "#FBC44F"
PURPLE = "#6936D9"
ORANGE = "#FF7A1A"


def load_chinese_fonts():
    """优先直接加载 Windows 中文字体文件，避免 Jupyter 中文丢失。"""
    regular_paths = [
        Path("C:/Windows/Fonts/msyh.ttc"),
        Path("C:/Windows/Fonts/msyhl.ttc"),
        Path("C:/Windows/Fonts/simhei.ttf"),
        Path("C:/Windows/Fonts/simsun.ttc"),
        Path("/System/Library/Fonts/PingFang.ttc"),
        Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
    ]
    bold_paths = [
        Path("C:/Windows/Fonts/msyhbd.ttc"),
        Path("C:/Windows/Fonts/simhei.ttf"),
        Path("/System/Library/Fonts/PingFang.ttc"),
        Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"),
    ]

    regular = next((p for p in regular_paths if p.exists()), None)
    bold = next((p for p in bold_paths if p.exists()), regular)

    if regular is not None:
        font_manager.fontManager.addfont(str(regular))
        normal_prop = FontProperties(fname=str(regular))
        bold_prop = FontProperties(fname=str(bold or regular), weight="bold")
        plt.rcParams["font.family"] = normal_prop.get_name()
    else:
        preferred = ["Microsoft YaHei", "SimHei", "Noto Sans CJK SC",
                     "Source Han Sans SC", "PingFang SC", "Arial Unicode MS"]
        available = {f.name for f in font_manager.fontManager.ttflist}
        selected = next((name for name in preferred if name in available), "DejaVu Sans")
        normal_prop = FontProperties(family=selected)
        bold_prop = FontProperties(family=selected, weight="bold")
        plt.rcParams["font.family"] = selected

    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["figure.dpi"] = 150
    plt.rcParams["savefig.dpi"] = 300
    return normal_prop, bold_prop


FONT, FONT_BOLD = load_chinese_fonts()


def tx(owner, x, y, text, *, bold=False, **kwargs):
    """统一文字函数：所有中文都显式指定 FontProperties。"""
    kwargs["fontproperties"] = FONT_BOLD if bold else FONT
    return owner.text(x, y, text, **kwargs)


def find_excel_file():
    """兼容 Jupyter，不使用 __file__。"""
    candidates = [
        Path.cwd() / EXCEL_FILE,
        Path.cwd() / "data" / EXCEL_FILE,
        Path.cwd() / "upload" / EXCEL_FILE,
        Path.cwd().parent / EXCEL_FILE,
    ]
    for path in candidates:
        if path.exists():
            return path
    raise FileNotFoundError(
        f"找不到 {EXCEL_FILE}。请把 Excel 与 Jupyter/.py 文件放在同一文件夹。"
    )


def read_rows(columns, start_row, end_row):
    workbook = load_workbook(find_excel_file(), data_only=True, read_only=True)
    if SHEET_NAME not in workbook.sheetnames:
        workbook.close()
        raise KeyError(f"找不到工作表：{SHEET_NAME}")
    ws = workbook[SHEET_NAME]
    rows = [[ws[f"{col}{row}"].value for col in columns]
            for row in range(start_row, end_row + 1)]
    workbook.close()
    return rows


def canvas(figsize=(10, 6.2)):
    fig, ax = plt.subplots(figsize=figsize, facecolor=BG)
    ax.set_facecolor(BG)
    return fig, ax


def header(fig, year, title, subtitle, metrics=()):
    tx(fig, 0.055, 0.925, year, color=WHITE, fontsize=21, bold=True)
    tx(fig, 0.055, 0.850, title, color=WHITE, fontsize=17, bold=True)
    if subtitle:
        tx(fig, 0.055, 0.805, subtitle, color=TEXT2, fontsize=10.5)
    if metrics:
        right = 0.94
        step = min(0.145, 0.50 / len(metrics))
        start = right - step * (len(metrics) - 1)
        for i, (value, label) in enumerate(metrics):
            x = start + i * step
            tx(fig, x, 0.925, value, color=WHITE, fontsize=13.5, ha="center")
            tx(fig, x, 0.894, label, color=MUTED, fontsize=7.3, ha="center")


def footer(fig, date="2022.06.30", source="注：数据来源于公司人力资源系统"):
    tx(fig, 0.055, 0.050, "*", color=WHITE, fontsize=9)
    tx(fig, 0.075, 0.050, source, color=MUTED, fontsize=7.5)
    tx(fig, 0.945, 0.050, date, color=TEXT2, fontsize=8, ha="right")


def style_category_axis(ax):
    for tick in ax.get_xticklabels() + ax.get_yticklabels():
        tick.set_fontproperties(FONT)
        tick.set_color(WHITE)


def finish(fig):
    output = Path.cwd() / OUTPUT_FILE
    fig.savefig(output, facecolor=BG, dpi=300)
    print(f"图片已保存：{output}")
    plt.show()

def create_chart():
    rows = read_rows(["B", "C", "D"], 3, 8)
    labels = [str(r[0]).strip() for r in rows]
    values = np.array([float(r[1]) for r in rows])
    totals = values + np.array([float(r[2]) for r in rows])
    colors = [PURPLE, BLUE, LIGHT_BLUE, CYAN, YELLOW, PINK]
    fig, ax = canvas()
    header(fig, "2022", "2022年上半年各部门人数",
           "公司总人数1664，销售部人数最多451，占比27%",
           [("1664", "公司总人数"), ("451", "销售部人数")])
    footer(fig)
    ax.set_position([.20,.12,.60,.66]); ax.set_aspect("equal"); ax.axis("off")
    for i,(lab,val,total,col) in enumerate(zip(labels,values,totals,colors)):
        r=.38+i*.13
        end=210+360*val/total
        ax.add_patch(Wedge((0,0),r,210,end,width=.09,facecolor=col,edgecolor="none"))
        ax.add_patch(Wedge((0,0),r,end,570,width=.09,facecolor="#30365E",edgecolor="none",alpha=.65))
        tx(ax, 1.34, .55-i*.19, f"{lab}  {int(val)}", color=col, fontsize=8.8, ha="right")
    tx(ax,0,.03,"27%",color=WHITE,fontsize=18,bold=True,ha="center")
    tx(ax,0,-.11,"销售部占比",color=MUTED,fontsize=7.5,ha="center")
    ax.set_xlim(-1.38,1.45); ax.set_ylim(-1.15,1.22)
    finish(fig)


if __name__ == "__main__": create_chart()

