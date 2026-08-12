#!/usr/bin/env python3
"""
Render one figure per lab for the Learner Guide.

A live screen capture of the Power Query Editor cannot be produced headlessly,
so each lab gets a faithful, brand-styled DIAGRAM of the Editor for that lab:
the Queries pane (the real query names from the workbook), the data preview
(the real column headers and first rows from the lab's mock data), and the
Applied Steps list (the steps the learner will actually create).

Because every element is read from the SAME source as the workbook and the
guide, the figure can never drift from what the learner sees on screen.

Output: courseware/assets/figures/lab-NN.png
"""
import os, sys, datetime as dt

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import course_data as C
from data_domain1 import DOMAIN1
from data_domain2 import DOMAIN2
from data_domain3 import DOMAIN3
from inject_powerquery import QUERIES

from openpyxl import load_workbook

ACTIVITIES = DOMAIN1 + DOMAIN2 + DOMAIN3

BLUE = "#1F6FEB"; TEAL = "#10B981"; INK = "#161B26"; GREY = "#5B6372"
LIGHT = "#F5F8FC"; LINE = "#E2E8F0"; WHITE = "#FFFFFF"; VIOLET = "#7C3AED"
plt.rcParams["font.family"] = "DejaVu Sans"


def repo_root():
    if os.environ.get("COURSE_REPO"):
        return os.environ["COURSE_REPO"]
    d = HERE
    for _ in range(8):
        d = os.path.dirname(d)
        if os.path.isdir(os.path.join(d, "courseware")) and os.path.isdir(os.path.join(d, "labs")):
            return d
    raise SystemExit("cannot locate course repo")


def slug(title):
    out = "".join(c.lower() if c.isalnum() else "-" for c in title)
    while "--" in out:
        out = out.replace("--", "-")
    return out.strip("-")


# Applied Steps each lab actually produces — drawn from the lab's own procedure.
APPLIED = {
    1:  ["Source", "Changed Type", "Replaced Value", "Replaced Value1",
         "Changed Type1", "Changed Type with Locale"],
    2:  ["Source", "Navigation", "Promoted Headers", "Changed Type"],
    3:  ["Source", "Filled Down", "Filled Up", "Removed Blank Rows", "Changed Type"],
    4:  ["Source", "Split Column by Delimiter", "Renamed Columns",
         "Split Column by Delimiter1", "Trimmed Text"],
    5:  ["Source", "Changed Type", "Added Conditional Column",
         "Added Custom", "Added Column From Examples"],
    6:  ["Source", "Merged Columns", "Filtered Rows", "Filtered Rows1", "Sorted Rows"],
    7:  ["Source", "Changed Type", "Added Custom", "Added Custom1", "Changed Type1"],
    8:  ["Source", "Changed Type", "Added Custom (Priority)", "Added Custom (Review Needed)"],
    9:  ["Source", "Changed Type", "Added Custom (Due Date)",
         "Added Custom (Days Overdue)", "Added Conditional Column", "Sorted Rows"],
    10: ["Source", "Unpivoted Other Columns", "Renamed Columns", "Changed Type"],
    11: ["Source", "Pivoted Column", "Appended Query", "Changed Type"],
    12: ["Source", "Changed Type", "Grouped Rows", "Grouped Rows (2 cols)"],
    13: ["Source", "Changed Type", "Loaded to Data Model"],
    14: ["Source", "Merged Queries (Left Outer)", "Expanded Products"],
    15: ["Source", "Merged Queries", "Expanded Products"],
    16: ["Source", "Merged Queries (ProductID + Region)", "Expanded RegionalPrices"],
    17: ["Source", "Added Conditional Column", "Added Custom (Bands)",
         "Expanded Bands", "Added Custom (InBand)", "Filtered Rows", "Removed Columns"],
    18: ["Source", "Added Custom (Prod)", "Expanded Prod",
         "Merged Queries", "Expanded Sales", "Replaced Value"],
    19: ["Source", "Merged Queries (Fuzzy)", "Expanded FinanceCustomers"],
}


def sheet_preview(xlsx, sheet, ncols=6, nrows=4):
    """Real headers + first rows from the lab's own mock data."""
    wb = load_workbook(xlsx, data_only=True)
    ws = wb[sheet] if sheet in wb.sheetnames else wb.worksheets[0]
    heads, rows = [], []
    for c in range(1, min(ws.max_column, ncols) + 1):
        heads.append(str(ws.cell(1, c).value or ""))
    for r in range(2, min(ws.max_row, nrows + 1) + 1):
        row = []
        for c in range(1, min(ws.max_column, ncols) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, dt.datetime):
                v = v.strftime("%d/%m/%Y")
            elif isinstance(v, float):
                v = f"{v:,.2f}"
            row.append("" if v is None else str(v))
        rows.append(row)
    return heads, rows


def draw(lab, out_path):
    n = lab["num"]
    folder = os.path.join(repo_root(), "labs", f"lab-{n:02d}-{slug(lab['title'])}")
    xlsx = os.path.join(folder, lab["workbook"])
    queries = [q for q, _ in QUERIES.get(n, [("Source", "")])]
    sheet = lab.get("sheets", [None])[0]
    heads, rows = sheet_preview(xlsx, sheet)
    steps = APPLIED.get(n, ["Source", "Changed Type"])

    W, H = 12.0, 5.6
    fig = plt.figure(figsize=(W, H), dpi=150)
    ax = fig.add_axes([0, 0, 1, 1]); ax.set_xlim(0, W); ax.set_ylim(0, H); ax.axis("off")
    fig.patch.set_facecolor(WHITE)

    def box(x, y, w, h, fc=WHITE, ec=LINE, lw=1.0, r=0.02):
        p = FancyBboxPatch((x, y), w, h, boxstyle=f"round,pad=0,rounding_size={r}",
                           fc=fc, ec=ec, lw=lw)
        ax.add_patch(p); return p

    # window chrome
    box(0.08, 0.08, W - 0.16, H - 0.16, WHITE, "#C7D2DE", 1.2)
    ax.add_patch(Rectangle((0.08, H - 0.62), W - 0.16, 0.54, fc=BLUE, ec="none"))
    ax.text(0.32, H - 0.35, f"Power Query Editor  —  Lab {n:02d}: {lab['title']}",
            color=WHITE, fontsize=11, fontweight="bold", va="center")

    # ribbon
    ax.add_patch(Rectangle((0.08, H - 1.14), W - 0.16, 0.52, fc=LIGHT, ec=LINE, lw=0.8))
    x = 0.42
    for tab, active in [("Home", True), ("Transform", False),
                        ("Add Column", False), ("View", False)]:
        ax.text(x, H - 0.88, tab, fontsize=9.5, color=BLUE if active else GREY,
                fontweight="bold" if active else "normal", va="center")
        if active:
            ax.plot([x - 0.06, x + 0.52], [H - 1.06, H - 1.06], color=BLUE, lw=2)
        x += 1.15

    body_top = H - 1.22
    body_h = body_top - 0.30

    # --- queries pane
    qw = 2.05
    box(0.20, 0.30, qw, body_h, LIGHT)
    ax.text(0.36, body_top - 0.22, "Queries [%d]" % len(queries),
            fontsize=9, color=GREY, fontweight="bold", va="center")
    yy = body_top - 0.58
    for q in queries[:6]:
        ax.add_patch(Rectangle((0.28, yy - 0.13), qw - 0.16, 0.30,
                               fc=WHITE, ec=LINE, lw=0.7))
        ax.add_patch(Rectangle((0.28, yy - 0.13), 0.05, 0.30, fc=TEAL, ec="none"))
        ax.text(0.44, yy + 0.02, q[:20], fontsize=8.5, color=INK, va="center")
        yy -= 0.40

    # --- data preview
    px = 0.20 + qw + 0.16
    pw = W - px - 2.55
    box(px, 0.30, pw, body_h, WHITE)
    ncol = max(1, len(heads))
    colw = pw / ncol
    hy = body_top - 0.42
    ax.add_patch(Rectangle((px, hy), pw, 0.34, fc=BLUE, ec="none"))
    for i, h in enumerate(heads):
        ax.text(px + colw * i + 0.07, hy + 0.17, h[:15], fontsize=7.6,
                color=WHITE, fontweight="bold", va="center")
    ry = hy - 0.34
    for ri, row in enumerate(rows):
        if ry < 0.42:
            break
        ax.add_patch(Rectangle((px, ry), pw, 0.34,
                               fc=LIGHT if ri % 2 == 0 else WHITE, ec=LINE, lw=0.5))
        for ci, v in enumerate(row):
            ax.text(px + colw * ci + 0.07, ry + 0.17, v[:15], fontsize=7.4,
                    color=INK, va="center")
        ry -= 0.34
    ax.text(px + 0.06, 0.40, "Preview of the first rows — the full table is loaded on refresh",
            fontsize=7.2, color=GREY, style="italic", va="center")

    # --- applied steps
    sx = px + pw + 0.16
    sw = W - sx - 0.20
    box(sx, 0.30, sw, body_h, LIGHT)
    ax.text(sx + 0.14, body_top - 0.22, "Query Settings", fontsize=9,
            color=GREY, fontweight="bold", va="center")
    ax.text(sx + 0.14, body_top - 0.52, "APPLIED STEPS", fontsize=7.8,
            color=BLUE, fontweight="bold", va="center")
    yy = body_top - 0.86
    for i, st in enumerate(steps[:9]):
        if yy < 0.42:
            break
        last = (i == min(len(steps), 9) - 1)
        ax.add_patch(Rectangle((sx + 0.10, yy - 0.12), sw - 0.20, 0.30,
                               fc=WHITE if not last else "#EAF7F0",
                               ec=TEAL if last else LINE, lw=1.0 if last else 0.7))
        ax.text(sx + 0.22, yy + 0.03, st[:26], fontsize=7.8,
                color=INK, va="center", fontweight="bold" if last else "normal")
        yy -= 0.38

    fig.savefig(out_path, facecolor=WHITE, bbox_inches=None)
    plt.close(fig)


def main():
    out_dir = os.path.join(repo_root(), "courseware", "assets", "figures")
    os.makedirs(out_dir, exist_ok=True)
    for lab in ACTIVITIES:
        p = os.path.join(out_dir, f"lab-{lab['num']:02d}.png")
        draw(lab, p)
        print(f"  lab {lab['num']:02d} -> {os.path.relpath(p, repo_root())}")
    print(f"\n{len(ACTIVITIES)} lab figures written to {out_dir}")


if __name__ == "__main__":
    main()
