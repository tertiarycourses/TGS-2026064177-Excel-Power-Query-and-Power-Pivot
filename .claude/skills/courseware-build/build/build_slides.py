#!/usr/bin/env python3
"""
Slide deck — Excel Power Query and Power Pivot (TGS-2026064177).

All-white Tertiary house style, built on the v2 component grammar
(build_slides_v2_reference.py): head/kicker + hairline + left tab, tile grids,
process maps with real connectors, compare tables, decision maps, worked
examples, activity cards and section dividers.

HOUSE RULE OBSERVED HERE: the deck is VISUAL and carries NO step-by-step
procedure. Every lab gets an overview card naming what is built and where the
data lives; the numbered click-by-click procedure lives ONLY in the Learner
Guide. Slides teach the concept, the decision and the pitfall.

Content is driven entirely by course_data.py + data_domainN.py.
"""
import os, sys, math, json, re
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import course_data as C
from data_domain1 import DOMAIN1
from data_domain2 import DOMAIN2
from data_domain3 import DOMAIN3
ACTIVITIES = DOMAIN1 + DOMAIN2 + DOMAIN3


def _find_repo(start):
    env = os.environ.get("COURSE_REPO")
    if env and os.path.isdir(env):
        return env
    d = start
    for _ in range(8):
        d = os.path.dirname(d)
        if os.path.isdir(os.path.join(d, "courseware")) and os.path.isdir(os.path.join(d, "labs")):
            return d
    return os.path.dirname(os.path.dirname(HERE))


REPO = _find_repo(HERE)
ASSETS = os.path.join(os.path.dirname(HERE), "assets")

# ---------------- palette ----------------
BLUE = RGBColor(0x1F, 0x6F, 0xEB); TEAL = RGBColor(0x10, 0xB9, 0x81)
AMBER = RGBColor(0xF5, 0x9E, 0x0B); INK = RGBColor(0x16, 0x1B, 0x26)
GREY = RGBColor(0x5B, 0x63, 0x72); LIGHT = RGBColor(0xF5, 0xF8, 0xFC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF); LINE = RGBColor(0xE2, 0xE8, 0xF0)
VIOLET = RGBColor(0x7C, 0x3A, 0xED); RED = RGBColor(0xDC, 0x26, 0x26)
PALETTE = [BLUE, TEAL, VIOLET, AMBER]

prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]
SLIDE_MAP = {}


def slide(): return prs.slides.add_slide(BLANK)


def rect(s, x, y, w, h, color, line=None):
    sp = s.shapes.add_shape(1, x, y, w, h); sp.fill.solid(); sp.fill.fore_color.rgb = color
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb = line; sp.line.width = Pt(1)
    sp.shadow.inherit = False; return sp


def oval(s, x, y, w, h, color):
    sp = s.shapes.add_shape(9, x, y, w, h); sp.fill.solid(); sp.fill.fore_color.rgb = color
    sp.line.fill.background(); sp.shadow.inherit = False; return sp


def roundrect(s, x, y, w, h, color, line=None, adj=0.10):
    sp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    sp.fill.solid(); sp.fill.fore_color.rgb = color
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb = line; sp.line.width = Pt(1.25)
    sp.shadow.inherit = False
    try: sp.adjustments[0] = adj
    except Exception: pass
    return sp


def diamond(s, x, y, w, h, color, line=None):
    sp = s.shapes.add_shape(MSO_SHAPE.DIAMOND, x, y, w, h)
    sp.fill.solid(); sp.fill.fore_color.rgb = color
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb = line; sp.line.width = Pt(1.25)
    sp.shadow.inherit = False; return sp


def txt(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, space=4):
    tb = s.shapes.add_textbox(x, y, w, h); tf = tb.text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor
    for i, ln in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.space_after = Pt(space)
        for t, sz, col, bold in ln:
            r = p.add_run(); r.text = t; r.font.size = Pt(sz); r.font.bold = bold
            r.font.color.rgb = col; r.font.name = "Arial"
    return tb


def label_in(sp, text, size, color, bold=True):
    tf = sp.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text; r.font.size = Pt(size); r.font.bold = bold
    r.font.color.rgb = color; r.font.name = "Arial"; return sp


def connector(s, x1, y1, x2, y2, color, width=Pt(2.0), arrow=True):
    cn = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    cn.line.color.rgb = color; cn.line.width = width
    ln = cn.line._get_or_add_ln()
    if arrow:
        t = etree.SubElement(ln, qn("a:tailEnd"))
        t.set("type", "triangle"); t.set("w", "med"); t.set("len", "med")
    return cn


def _ellipsis(text, limit):
    t = " ".join(str(text).split())
    if len(t) <= limit: return t
    cut = t[:limit]; sp = cut.rfind(" ")
    if sp > limit * 0.55: cut = cut[:sp]
    return cut.rstrip(" ,.;:-—(") + "…"


def _sentences(text, limit):
    """Keep whole sentences up to `limit` characters — a slide must never end
    mid-clause with an ellipsis. Falls back to a word-boundary trim only if the
    very first sentence is already longer than the limit."""
    t = " ".join(str(text).split())
    if len(t) <= limit:
        return t
    parts = re.split(r'(?<=[.!?]) +', t)
    out = ""
    for p in parts:
        cand = (out + " " + p).strip() if out else p
        if len(cand) > limit:
            break
        out = cand
    return out if out else _ellipsis(t, limit)


def _fit_title(title, size=29):
    n = len(title)
    if n <= 52: return size
    if n <= 66: return 25
    if n <= 82: return 22
    return 20


PAGE = {"n": 1}


def footer(s):
    PAGE["n"] += 1
    txt(s, Inches(0.4), Inches(7.05), Inches(7.5), Inches(0.35),
        [[(f"{C.SHORT_TITLE}  ·  {C.COURSE_CODE}", 9, GREY, False)]])
    txt(s, Inches(5.0), Inches(7.05), Inches(3.3), Inches(0.35),
        [[("© 2026 Tertiary Infotech Academy Pte Ltd", 9, GREY, False)]], align=PP_ALIGN.CENTER)
    txt(s, Inches(12.4), Inches(7.05), Inches(0.6), Inches(0.35),
        [[(str(PAGE["n"]), 9, GREY, False)]], align=PP_ALIGN.RIGHT)


def mark(key):
    SLIDE_MAP[key] = PAGE["n"] + 1


def head(s, title, kicker=None, kcolor=BLUE):
    rect(s, 0, 0, SW, SH, WHITE); rect(s, 0, 0, Inches(0.28), Inches(1.55), kcolor)
    if kicker:
        txt(s, Inches(0.85), Inches(0.5), Inches(11.6), Inches(0.4), [[(kicker, 14, kcolor, True)]])
    txt(s, Inches(0.85), Inches(0.88), Inches(11.9), Inches(0.78),
        [[(title, _fit_title(title), INK, True)]], anchor=MSO_ANCHOR.MIDDLE)
    rect(s, Inches(0.85), Inches(1.7), Inches(11.63), Inches(0.02), LINE)
    return s


def _logo(name):
    p = os.path.join(ASSETS, name)
    return p if os.path.exists(p) else None


P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"


def _transition(s, kind="fade", speed="med"):
    sld = s._element
    for old in sld.findall(qn("p:transition")): sld.remove(old)
    tr = etree.SubElement(sld, qn("p:transition")); tr.set("spd", speed); tr.set("advClick", "1")
    if kind == "fade": etree.SubElement(tr, qn("p:fade"))
    elif kind == "push":
        el = etree.SubElement(tr, qn("p:push")); el.set("dir", "l")
    sld.append(tr); return tr


# ============================================================ templates
def cover():
    s = slide(); rect(s, 0, 0, SW, SH, WHITE)
    rect(s, 0, 0, SW, Inches(0.22), BLUE); rect(s, 0, Inches(7.28), SW, Inches(0.22), TEAL)
    org = _logo("tertiary-infotech-logo.png")
    if org: s.shapes.add_picture(org, Inches(0.85), Inches(0.7), height=Inches(1.05))
    # Excel / Power Query badge
    rect(s, Inches(10.55), Inches(0.66), Inches(2.05), Inches(1.12), RGBColor(0x21, 0x73, 0x46))
    txt(s, Inches(10.55), Inches(0.76), Inches(2.05), Inches(0.5),
        [[("EXCEL", 21, WHITE, True)]], align=PP_ALIGN.CENTER)
    txt(s, Inches(10.55), Inches(1.26), Inches(2.05), Inches(0.42),
        [[("POWER QUERY · PIVOT", 8.5, WHITE, True)]], align=PP_ALIGN.CENTER)
    txt(s, Inches(0.9), Inches(2.3), Inches(12), Inches(0.6),
        [[("COURSE SLIDES  ·  WSQ", 16, BLUE, True)]])
    txt(s, Inches(0.9), Inches(2.85), Inches(12.0), Inches(1.9), [[(C.TITLE, 40, INK, True)]])
    rect(s, Inches(0.92), Inches(4.75), Inches(2.4), Inches(0.06), TEAL)
    txt(s, Inches(0.9), Inches(5.05), Inches(12), Inches(1.4),
        [[(f"WSQ Course Code: {C.COURSE_CODE}", 16, GREY, False)],
         [(f"Conducted by {C.ORG}  ·  {C.UEN}", 14, GREY, False)]], space=6)
    txt(s, Inches(0.9), Inches(6.5), Inches(12), Inches(0.4),
        [[(f"Version {C.VERSION}  ·  {C.VERSION_DATE}", 12, GREY, False)]])
    txt(s, Inches(0.9), Inches(6.85), Inches(12), Inches(0.34),
        [[(f"© 2026 {C.ORG}. All rights reserved.  ·  {C.WEBSITE}", 10, GREY, False)]])


def section(kicker, title, n, sub=""):
    s = slide(); rect(s, 0, 0, SW, SH, WHITE); rect(s, 0, 0, Inches(0.28), SH, BLUE)
    rect(s, Inches(0.85), Inches(2.5), Inches(0.14), Inches(2.0), TEAL)
    txt(s, Inches(1.25), Inches(2.55), Inches(11), Inches(0.6), [[(kicker, 18, BLUE, True)]])
    txt(s, Inches(1.25), Inches(3.0), Inches(11.4), Inches(1.6), [[(title, 40, INK, True)]])
    if sub: txt(s, Inches(1.27), Inches(4.55), Inches(11), Inches(0.8), [[(sub, 16, GREY, False)]])
    txt(s, Inches(10.0), Inches(0.7), Inches(2.8), Inches(1.6), [[(n, 72, LINE, True)]],
        align=PP_ALIGN.RIGHT)
    footer(s); return s


def bullets(s, x, y, w, h, items, size=18, color=INK, gap=10, mcolor=BLUE):
    tb = s.shapes.add_textbox(x, y, w, h); tf = tb.text_frame; tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); p.space_after = Pt(gap)
        lvl = it[1] if isinstance(it, tuple) else 0
        text = it[0] if isinstance(it, tuple) else it
        r = p.add_run(); r.text = ("•  " if lvl == 0 else "–  ") + text
        r.font.size = Pt(size if lvl == 0 else size - 2)
        r.font.color.rgb = color if lvl == 0 else GREY; r.font.name = "Arial"
    return tb


def content(title, items, kicker=None, size=20):
    s = head(slide(), title, kicker)
    bullets(s, Inches(0.85), Inches(1.95), Inches(11.6), Inches(4.9), items, size=size)
    footer(s); return s


def big_statement(line1, line2, kicker, color=BLUE):
    s = slide(); rect(s, 0, 0, SW, SH, WHITE); rect(s, 0, 0, Inches(0.28), SH, color)
    txt(s, Inches(1.1), Inches(2.2), Inches(11), Inches(0.5), [[(kicker, 16, color, True)]])
    txt(s, Inches(1.1), Inches(2.8), Inches(11.3), Inches(2.4), [[(line1, 38, INK, True)]])
    if line2: txt(s, Inches(1.12), Inches(4.9), Inches(11), Inches(1.2), [[(line2, 20, GREY, False)]])
    footer(s); return s


def tile_grid(title, items, kicker=None, cols=2, size=15, icons=None, accent=BLUE):
    s = head(slide(), title, kicker, kcolor=accent)
    n = len(items); rows = math.ceil(n / cols)
    X0 = Inches(0.85); Y0 = Inches(1.95); TOTW = Inches(11.63)
    # the grid must finish above the footer at 7.05 — budget to 6.80
    AREAH = Inches(6.80) - Y0
    gx = Inches(0.3); gy = Inches(0.26) if rows < 3 else Inches(0.18)
    cw = int((TOTW - gx * (cols - 1)) / cols); ch = int((AREAH - gy * (rows - 1)) / rows)
    bd = Inches(0.6) if ch > Inches(1.2) else Inches(0.46)
    for i, it in enumerate(items):
        r = i // cols; c = i % cols
        x = int(X0 + (cw + gx) * c); y = int(Y0 + (ch + gy) * r); col = PALETTE[i % len(PALETTE)]
        rect(s, x, y, cw, ch, LIGHT); rect(s, x, y, Inches(0.1), ch, col)
        oval(s, x + Inches(0.28), int(y + ch / 2 - bd / 2), bd, bd, col)
        ic = icons[i] if icons else str(i + 1)
        txt(s, x + Inches(0.28), int(y + ch / 2 - bd / 2), bd, bd,
            [[(ic, 19 if bd > Inches(0.5) else 15, WHITE, True)]],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        tx = x + Inches(1.08); tw = cw - Inches(1.32)
        if isinstance(it, tuple):
            txt(s, tx, int(y + Inches(0.14)), tw, int(ch - Inches(0.2)),
                [[(it[0], size + 2, INK, True)], [(it[1], size - 2, GREY, False)]],
                anchor=MSO_ANCHOR.MIDDLE, space=3)
        else:
            txt(s, tx, int(y + Inches(0.1)), tw, int(ch - Inches(0.16)),
                [[(it, size, INK, False)]], anchor=MSO_ANCHOR.MIDDLE)
    footer(s); return s


def ncards(title, cards, kicker=None, accent=BLUE, cols=4, synthesis=None):
    """Outlined numbered concept cards — the reference deck's signature 4-across."""
    s = head(slide(), title, kicker, kcolor=accent)
    n = len(cards); rows = math.ceil(n / cols)
    X0 = Inches(0.72); TOTW = Inches(11.85)
    gx = Inches(0.28)
    cw = int((TOTW - gx * (cols - 1)) / cols)
    y0 = Inches(1.88)
    # the synthesis band is PINNED just above the footer; cards take what is left
    SYN_H = Inches(0.96)
    BOTTOM = Inches(6.86)
    syn_y = BOTTOM - SYN_H
    card_bottom = (syn_y - Inches(0.16)) if synthesis else BOTTOM
    gy = Inches(0.24)
    ch = int((card_bottom - y0 - gy * (rows - 1)) / rows)
    for i, cd in enumerate(cards):
        ttl, body = cd[0], cd[1]
        r = i // cols; c = i % cols
        x = int(X0 + (cw + gx) * c); y = int(y0 + (ch + gy) * r)
        col = PALETTE[i % len(PALETTE)]
        rect(s, x, y, cw, ch, LIGHT)
        rect(s, x, y, Inches(0.09), ch, col)
        bd = Inches(0.46)
        oval(s, x + Inches(0.24), y + Inches(0.2), bd, bd, col)
        txt(s, x + Inches(0.24), y + Inches(0.2), bd, bd, [[(str(i + 1), 15, WHITE, True)]],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        txt(s, x + Inches(0.82), y + Inches(0.14), cw - Inches(1.0), Inches(0.6),
            [[(_ellipsis(ttl, 34), 14, col, True)]], anchor=MSO_ANCHOR.MIDDLE)
        txt(s, x + Inches(0.28), y + Inches(0.78), cw - Inches(0.54), int(ch - Inches(0.9)),
            [[(body, 11, INK, False)]])
    if synthesis:
        by = int(syn_y)
        rect(s, X0, by, TOTW, SYN_H, LIGHT)
        rect(s, X0, by, Inches(0.11), SYN_H, accent)
        txt(s, X0 + Inches(0.3), by + Inches(0.1), TOTW - Inches(0.6), Inches(0.3),
            [[(synthesis[0].upper(), 11, accent, True)]])
        txt(s, X0 + Inches(0.3), by + Inches(0.4), TOTW - Inches(0.6), Inches(0.5),
            [[(synthesis[1], 12, INK, False)]])
    footer(s); return s


def process_map(title, stages, kicker=None, color=BLUE, synthesis=None):
    s = head(slide(), title, kicker, kcolor=color)
    n = len(stages); X0 = Inches(0.85); TOTW = Inches(11.63)
    gap = Inches(0.42); cw = int((TOTW - gap * (n - 1)) / n)
    y = Inches(2.35); ch = Inches(2.35) if synthesis else Inches(3.0)
    for i, st in enumerate(stages):
        lbl, detail = (st if isinstance(st, tuple) else (st, ""))
        x = int(X0 + (cw + gap) * i)
        roundrect(s, x, y, cw, ch, LIGHT, line=LINE)
        rect(s, x, y, cw, Inches(0.11), color)
        bd = Inches(0.62)
        oval(s, int(x + cw / 2 - bd / 2), int(y + Inches(0.34)), bd, bd, color)
        txt(s, int(x + cw / 2 - bd / 2), int(y + Inches(0.34)), bd, bd,
            [[(str(i + 1), 24, WHITE, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        lbl = _ellipsis(lbl, 44)
        lsz = 13 if len(lbl) <= 28 else (12 if len(lbl) <= 38 else 11)
        txt(s, x + Inches(0.14), int(y + Inches(1.06)), cw - Inches(0.28), Inches(0.94),
            [[(lbl, lsz, INK, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space=0)
        if detail:
            # the caption band must fit two wrapped lines at 9pt, or the second
            # line is clipped by the card edge
            txt(s, x + Inches(0.08), int(y + ch - Inches(0.62)), cw - Inches(0.16), Inches(0.56),
                [[(_ellipsis(detail, 30), 9, GREY, False)]],
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space=0)
        if i < n - 1:
            cy = int(y + ch / 2)
            connector(s, int(x + cw + Inches(0.06)), cy, int(x + cw + gap - Inches(0.06)), cy, color)
    if synthesis:
        by = int(y + ch + Inches(0.34))
        rect(s, Inches(0.85), by, Inches(11.63), Inches(1.15), LIGHT)
        rect(s, Inches(0.85), by, Inches(0.11), Inches(1.15), color)
        txt(s, Inches(1.15), int(by + Inches(0.12)), Inches(11.1), Inches(0.32),
            [[(synthesis[0].upper(), 11, color, True)]])
        txt(s, Inches(1.15), int(by + Inches(0.46)), Inches(11.1), Inches(0.62),
            [[(synthesis[1], 13, INK, False)]])
    footer(s); return s


def decision_map(title, question, yes, no, kicker=None, color=VIOLET, note=None):
    s = head(slide(), title, kicker, kcolor=color)
    dx, dy, dw, dh = Inches(0.85), Inches(2.75), Inches(4.6), Inches(2.7)
    d = diamond(s, dx, dy, dw, dh, color); label_in(d, question, 12, WHITE)
    bx = Inches(6.1); bw = Inches(6.35); bh = Inches(1.5)
    ys = [Inches(2.15), Inches(4.35)]
    for (hdr, items), by, col in zip([yes, no], ys, [TEAL, AMBER]):
        roundrect(s, bx, by, bw, bh, LIGHT, line=LINE)
        rect(s, bx, by, Inches(0.11), bh, col)
        txt(s, bx + Inches(0.3), int(by + Inches(0.16)), bw - Inches(0.55), Inches(0.4),
            [[(hdr, 15, col, True)]])
        txt(s, bx + Inches(0.3), int(by + Inches(0.6)), bw - Inches(0.55), Inches(0.82),
            [[(items, 12, INK, False)]])
        connector(s, int(dx + dw), int(dy + dh / 2), bx, int(by + bh / 2), col)
    if note:
        rect(s, Inches(0.85), Inches(6.15), Inches(11.63), Inches(0.72), LIGHT)
        txt(s, Inches(1.15), Inches(6.28), Inches(11.1), Inches(0.5), [[(note, 12, GREY, False)]])
    footer(s); return s


def compare_table(title, headers, rows, kicker=None, accent=BLUE, note=None, first_frac=0.26):
    s = head(slide(), title, kicker, kcolor=accent)
    X0 = Inches(0.85); TOTW = Inches(11.63); ncol = len(headers)
    first = int(TOTW * first_frac); rest = int((TOTW - first) / (ncol - 1))
    widths = [first] + [rest] * (ncol - 1)
    y = Inches(1.95); hh = Inches(0.52); x = X0
    for i, h in enumerate(headers):
        col = accent if i == 0 else PALETTE[(i - 1) % len(PALETTE)]
        rect(s, x, y, widths[i], hh, col)
        txt(s, x + Inches(0.14), y, widths[i] - Inches(0.28), hh, [[(h, 13, WHITE, True)]],
            anchor=MSO_ANCHOR.MIDDLE)
        x += widths[i]
    ry = int(y + hh)
    BOTTOM = Inches(6.88) - (Inches(1.24) if note else Emu(0))
    avail = BOTTOM - ry
    rh = int(min(Inches(0.78), avail / max(len(rows), 1)))
    for r, row in enumerate(rows):
        x = X0
        for i, cell in enumerate(row):
            fill = LIGHT if r % 2 == 0 else WHITE
            rect(s, x, ry, widths[i], rh, fill, line=LINE)
            bold = (i == 0)
            txt(s, x + Inches(0.12), ry, widths[i] - Inches(0.24), rh,
                [[(cell, 11 if not bold else 11.5, INK if bold else GREY, bold)]],
                anchor=MSO_ANCHOR.MIDDLE)
            x += widths[i]
        ry += rh
    if note:
        rect(s, X0, int(ry + Inches(0.22)), TOTW, Inches(0.92), LIGHT)
        rect(s, X0, int(ry + Inches(0.22)), Inches(0.11), Inches(0.92), accent)
        txt(s, X0 + Inches(0.3), int(ry + Inches(0.34)), TOTW - Inches(0.6), Inches(0.66),
            [[("WHEN IT MATTERS  ", 11, accent, True), (note, 12, INK, False)]])
    footer(s); return s


def code_example(title, intro, code, explain, kicker=None, accent=TEAL, lang="M"):
    """M / DAX code on the left, the line-by-line reading on the right."""
    s = head(slide(), title, kicker, kcolor=accent)
    txt(s, Inches(0.85), Inches(1.9), Inches(11.63), Inches(0.46), [[(intro, 15, GREY, False)]])
    cx, cw = Inches(0.85), Inches(6.5)
    rect(s, cx, Inches(2.5), cw, Inches(4.15), RGBColor(0x0B, 0x12, 0x20))
    txt(s, cx + Inches(0.2), Inches(2.56), cw - Inches(0.34), Inches(0.3),
        [[(lang, 9, RGBColor(0x6A, 0x99, 0x55), True)]])
    tb = s.shapes.add_textbox(cx + Inches(0.22), Inches(2.86), cw - Inches(0.44), Inches(3.6))
    tf = tb.text_frame; tf.word_wrap = True
    KEY = RGBColor(0xC5, 0x86, 0xC0); STR = RGBColor(0xCE, 0x91, 0x78)
    PLAIN = RGBColor(0xD4, 0xD4, 0xD4); CMT = RGBColor(0x6A, 0x99, 0x55)
    for i, lnq in enumerate(code):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); p.space_after = Pt(2)
        r = p.add_run(); r.text = lnq; r.font.size = Pt(11); r.font.name = "Consolas"
        st = lnq.strip()
        col = PLAIN
        if st.startswith("//"): col = CMT
        elif re.match(r'^(let|in|shared|section)\b', st): col = KEY
        elif '"' in lnq: col = STR
        elif re.search(r'\b(if|then|else|each|and|or|not)\b', lnq): col = KEY
        r.font.color.rgb = col
    ex, ew = Inches(7.65), Inches(4.83)
    for i, (lbl, body) in enumerate(explain[:4]):
        y = int(Inches(2.5) + (Inches(1.0) + Inches(0.05)) * i)
        col = PALETTE[i % len(PALETTE)]
        rect(s, ex, y, ew, Inches(1.0), LIGHT); rect(s, ex, y, Inches(0.09), Inches(1.0), col)
        txt(s, ex + Inches(0.26), int(y + Inches(0.1)), ew - Inches(0.45), Inches(0.32),
            [[(lbl, 12, col, True)]])
        txt(s, ex + Inches(0.26), int(y + Inches(0.42)), ew - Inches(0.45), Inches(0.52),
            [[(body, 11, INK, False)]])
    footer(s); return s


def good_bad(title, good, bad, kicker=None, note=None):
    """Green ✓ vs red ✗ outlined panels — the pitfall slide."""
    s = head(slide(), title, kicker, kcolor=TEAL)
    # size the panels to the content so long lists fill the slide and short ones
    # do not leave a dead band above the note
    rows = max(len(good[1]), len(bad[1]))
    panel_h = min(Inches(4.4), Inches(0.92) + Inches(0.44) * rows)
    for (hdr, items, col, mark_), x in zip(
            [(good[0], good[1], TEAL, "✓"), (bad[0], bad[1], RED, "✗")],
            [Inches(0.85), Inches(6.95)]):
        w = Inches(5.55)
        rect(s, x, Inches(1.95), w, panel_h, WHITE, line=col)
        rect(s, x, Inches(1.95), w, Inches(0.56), col)
        txt(s, x + Inches(0.2), Inches(1.95), w - Inches(0.4), Inches(0.56),
            [[(f"{mark_}  {hdr}", 14, WHITE, True)]], anchor=MSO_ANCHOR.MIDDLE)
        tb = s.shapes.add_textbox(x + Inches(0.28), Inches(2.62), w - Inches(0.56),
                                  panel_h - Inches(0.78))
        tf = tb.text_frame; tf.word_wrap = True
        for i, it in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); p.space_after = Pt(9)
            r = p.add_run(); r.text = "•  " + it; r.font.size = Pt(12.5)
            r.font.color.rgb = INK; r.font.name = "Arial"
    if note:
        ny = int(Inches(1.95) + panel_h + Inches(0.22))
        rect(s, Inches(0.85), ny, Inches(11.65), Inches(0.7), LIGHT)
        rect(s, Inches(0.85), ny, Inches(0.11), Inches(0.7), TEAL)
        txt(s, Inches(1.2), ny, Inches(11.1), Inches(0.7), [[(note, 12, GREY, False)]],
            anchor=MSO_ANCHOR.MIDDLE)
    footer(s); return s


def takeaway(title, line, kicker=None, accent=BLUE, points=None):
    """A statement slide with one boxed rule + optional supporting points."""
    s = head(slide(), title, kicker, kcolor=accent)
    rect(s, Inches(0.85), Inches(2.0), Inches(11.63), Inches(1.5), LIGHT)
    rect(s, Inches(0.85), Inches(2.0), Inches(0.14), Inches(1.5), accent)
    txt(s, Inches(1.25), Inches(2.0), Inches(11.0), Inches(1.5), [[(line, 20, INK, True)]],
        anchor=MSO_ANCHOR.MIDDLE)
    if points:
        for i, (h, b) in enumerate(points[:3]):
            x = Inches(0.85) + i * Inches(3.99)
            col = PALETTE[i % len(PALETTE)]
            rect(s, x, Inches(3.85), Inches(3.71), Inches(2.5), LIGHT)
            rect(s, x, Inches(3.85), Inches(3.71), Inches(0.1), col)
            txt(s, x + Inches(0.26), Inches(4.1), Inches(3.2), Inches(0.5), [[(h, 14, col, True)]])
            txt(s, x + Inches(0.26), Inches(4.62), Inches(3.2), Inches(1.6), [[(b, 11.5, INK, False)]])
    footer(s); return s


def activity_overview(lab, kicker):
    """Lab briefing card — NO step-by-step. Points to the Learner Guide."""
    s = head(slide(), f"Lab {lab['num']:02d}: {lab['title']}", kicker, kcolor=TEAL)
    rect(s, Inches(10.55), Inches(0.42), Inches(1.95), Inches(0.5), TEAL)
    txt(s, Inches(10.55), Inches(0.42), Inches(1.95), Inches(0.5),
        [[(f"LAB {lab['num']:02d}  ·  30 MIN", 10.5, WHITE, True)]],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # scenario — trim on a SENTENCE boundary so the slide never ends mid-clause
    # (the Learner Guide carries the full paragraph)
    txt(s, Inches(0.85), Inches(1.92), Inches(11.63), Inches(0.95),
        [[(_sentences(lab["desc"], 300), 13.5, GREY, False)]])
    # what you'll produce
    rect(s, Inches(0.85), Inches(3.0), Inches(11.63), Inches(0.92), LIGHT)
    rect(s, Inches(0.85), Inches(3.0), Inches(0.11), Inches(0.92), TEAL)
    txt(s, Inches(1.15), Inches(3.08), Inches(11.1), Inches(0.3),
        [[("YOU'LL PRODUCE", 11, TEAL, True)]])
    txt(s, Inches(1.15), Inches(3.38), Inches(11.1), Inches(0.48),
        [[(lab["build"], 13, INK, False)]])
    # verification
    rect(s, Inches(0.85), Inches(4.02), Inches(11.63), Inches(0.92), LIGHT)
    rect(s, Inches(0.85), Inches(4.02), Inches(0.11), Inches(0.92), BLUE)
    txt(s, Inches(1.15), Inches(4.1), Inches(11.1), Inches(0.3),
        [[("HOW YOU'LL KNOW IT WORKED", 11, BLUE, True)]])
    txt(s, Inches(1.15), Inches(4.4), Inches(11.1), Inches(0.48),
        [[(_ellipsis(lab["test"], 160), 12.5, INK, False)]])
    # maps-to
    rect(s, Inches(0.85), Inches(5.04), Inches(11.63), Inches(0.72), LIGHT)
    rect(s, Inches(0.85), Inches(5.04), Inches(0.11), Inches(0.72), VIOLET)
    txt(s, Inches(1.15), Inches(5.14), Inches(11.1), Inches(0.5),
        [[("MAPS TO  ", 11, VIOLET, True), (lab["objective"], 11.5, INK, False)]])
    # lab band — file + pointer to the LG
    rect(s, Inches(0.85), Inches(5.94), Inches(11.63), Inches(0.86), WHITE, line=LINE)
    rect(s, Inches(1.0), Inches(6.08), Inches(1.15), Inches(0.5), TEAL)
    txt(s, Inches(1.0), Inches(6.08), Inches(1.15), Inches(0.5),
        [[(f"LAB {lab['num']:02d}", 12, WHITE, True)]], align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(2.35), Inches(6.02), Inches(10.0), Inches(0.34),
        [[(lab["workbook"], 13, INK, True)]])
    txt(s, Inches(2.35), Inches(6.36), Inches(10.0), Inches(0.34),
        [[(f"{lab['services']}  ·  Full step-by-step: Learner Guide, Lab {lab['num']:02d}",
           10.5, GREY, False)]])
    footer(s); return s


def trainer_slide(kicker, name, role, rows, initials, accent=BLUE):
    s = head(slide(), "About the Trainer", kicker, kcolor=accent)
    rect(s, Inches(0.85), Inches(1.95), Inches(3.6), Inches(4.7), LIGHT)
    rect(s, Inches(0.85), Inches(1.95), Inches(3.6), Inches(0.12), accent)
    oval(s, Inches(1.85), Inches(2.45), Inches(1.6), Inches(1.6), accent)
    txt(s, Inches(1.85), Inches(2.45), Inches(1.6), Inches(1.6), [[(initials, 44, WHITE, True)]],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(1.05), Inches(4.25), Inches(3.2), Inches(0.5), [[(name, 18, INK, True)]],
        align=PP_ALIGN.CENTER)
    txt(s, Inches(1.05), Inches(4.75), Inches(3.2), Inches(1.0), [[(role, 12, GREY, False)]],
        align=PP_ALIGN.CENTER)
    for i, (lbl, val) in enumerate(rows[:4]):
        y = int(Inches(1.95) + (Inches(1.15) + Inches(0.05)) * i)
        col = PALETTE[i % len(PALETTE)]
        rect(s, Inches(4.85), y, Inches(7.63), Inches(1.15), LIGHT)
        rect(s, Inches(4.85), y, Inches(0.09), Inches(1.15), col)
        txt(s, Inches(5.15), int(y + Inches(0.12)), Inches(7.1), Inches(0.32),
            [[(lbl, 12, col, True)]])
        txt(s, Inches(5.15), int(y + Inches(0.44)), Inches(7.1), Inches(0.62),
            [[(val, 12, INK, False)]])
    footer(s); return s


def browser_slide(title, url, kicker, tiles, caption=None, accent=BLUE):
    """A browser mock (LMS / practice exam) + numbered how-to tiles."""
    s = head(slide(), title, kicker, kcolor=accent)
    bx, by, bw, bh = Inches(0.85), Inches(1.95), Inches(7.0), Inches(4.3)
    rect(s, bx, by, bw, bh, WHITE, line=LINE)
    rect(s, bx, by, bw, Inches(0.52), LIGHT)
    for i, col in enumerate([RGBColor(0xFF, 0x5F, 0x57), RGBColor(0xFE, 0xBC, 0x2E),
                             RGBColor(0x28, 0xC8, 0x40)]):
        oval(s, bx + Inches(0.2) + i * Inches(0.28), by + Inches(0.17), Inches(0.16),
             Inches(0.16), col)
    rect(s, bx + Inches(1.2), by + Inches(0.11), bw - Inches(1.6), Inches(0.32), WHITE, line=LINE)
    txt(s, bx + Inches(1.35), by + Inches(0.11), bw - Inches(1.8), Inches(0.32),
        [[(url, 10.5, GREY, False)]], anchor=MSO_ANCHOR.MIDDLE)
    # page mock
    rect(s, bx + Inches(0.3), by + Inches(0.75), bw - Inches(0.6), Inches(0.5), accent)
    txt(s, bx + Inches(0.5), by + Inches(0.75), bw - Inches(1.0), Inches(0.5),
        [[(C.SHORT_TITLE, 12, WHITE, True)]], anchor=MSO_ANCHOR.MIDDLE)
    for i in range(4):
        yy = by + Inches(1.42) + i * Inches(0.62)
        rect(s, bx + Inches(0.3), yy, bw - Inches(0.6), Inches(0.5), LIGHT)
        rect(s, bx + Inches(0.3), yy, Inches(0.07), Inches(0.5), PALETTE[i % 4])
        lbls = ["Trainer Slides (PPT)", "Learner Guide (PDF)",
                "Lesson Plan (PDF)", "Activities / Lab Files"]
        txt(s, bx + Inches(0.55), yy, bw - Inches(1.0), Inches(0.5),
            [[(lbls[i], 11, INK, False)]], anchor=MSO_ANCHOR.MIDDLE)
    ex, ew = Inches(8.15), Inches(4.33)
    for i, (h, b) in enumerate(tiles[:4]):
        y = int(Inches(1.95) + (Inches(1.02) + Inches(0.06)) * i)
        col = PALETTE[i % len(PALETTE)]
        rect(s, ex, y, ew, Inches(1.02), LIGHT); rect(s, ex, y, Inches(0.09), Inches(1.02), col)
        bd = Inches(0.4)
        oval(s, ex + Inches(0.24), int(y + Inches(0.14)), bd, bd, col)
        txt(s, ex + Inches(0.24), int(y + Inches(0.14)), bd, bd, [[(str(i + 1), 13, WHITE, True)]],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        txt(s, ex + Inches(0.76), int(y + Inches(0.1)), ew - Inches(0.95), Inches(0.34),
            [[(h, 12, col, True)]])
        txt(s, ex + Inches(0.3), int(y + Inches(0.46)), ew - Inches(0.55), Inches(0.5),
            [[(b, 10.5, INK, False)]])
    if caption:
        rect(s, Inches(0.85), Inches(6.4), Inches(11.63), Inches(0.6), LIGHT)
        txt(s, Inches(1.15), Inches(6.46), Inches(11.1), Inches(0.48), [[(caption, 12, GREY, False)]])
    footer(s); return s


def brk(kind, dur, color=AMBER):
    s = slide(); rect(s, 0, 0, SW, SH, WHITE); rect(s, 0, 0, Inches(0.28), SH, color)
    rect(s, Inches(0.85), Inches(2.6), Inches(11.63), Inches(2.3), LIGHT)
    rect(s, Inches(0.85), Inches(2.6), Inches(0.14), Inches(2.3), color)
    txt(s, Inches(1.3), Inches(3.0), Inches(11), Inches(1.0), [[(kind, 40, INK, True)]])
    txt(s, Inches(1.32), Inches(4.0), Inches(11), Inches(0.6), [[(dur, 18, GREY, False)]])
    footer(s); return s


def ribbon_slide(title, kicker, tabs, groups, caption, accent=BLUE):
    """A mock of the Excel / Power Query Editor ribbon — teaches WHERE things live."""
    s = head(slide(), title, kicker, kcolor=accent)
    rx, ry, rw = Inches(0.85), Inches(2.05), Inches(11.63)
    rect(s, rx, ry, rw, Inches(0.46), LIGHT)
    x = rx + Inches(0.2)
    for i, t in enumerate(tabs):
        w = Inches(0.28) + Inches(0.105) * len(t)
        if i == 0:
            rect(s, x, ry, w, Inches(0.46), WHITE, line=LINE)
        txt(s, x, ry, w, Inches(0.46), [[(t, 11, accent if i == 0 else GREY, i == 0)]],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        x += w + Inches(0.12)
    # ribbon body with grouped buttons
    by = ry + Inches(0.46)
    rect(s, rx, by, rw, Inches(1.5), WHITE, line=LINE)
    gx = rx + Inches(0.15)
    for i, (gname, items) in enumerate(groups):
        gw = Inches(0.55) + Inches(0.72) * len(items)
        for j, it in enumerate(items):
            bxx = gx + Inches(0.12) + j * Inches(0.72)
            col = PALETTE[(i + j) % len(PALETTE)]
            rect(s, bxx, by + Inches(0.16), Inches(0.6), Inches(0.6), LIGHT)
            rect(s, bxx, by + Inches(0.16), Inches(0.6), Inches(0.08), col)
            txt(s, bxx - Inches(0.08), by + Inches(0.8), Inches(0.76), Inches(0.42),
                [[(it, 8, INK, False)]], align=PP_ALIGN.CENTER)
        txt(s, gx, by + Inches(1.22), gw, Inches(0.24), [[(gname, 9, GREY, False)]],
            align=PP_ALIGN.CENTER)
        gx += gw
        if i < len(groups) - 1:
            rect(s, gx, by + Inches(0.12), Inches(0.012), Inches(1.15), LINE)
            gx += Inches(0.14)
    for i, (h, b) in enumerate(caption[:3]):
        x = Inches(0.85) + i * Inches(3.99)
        col = PALETTE[i % len(PALETTE)]
        rect(s, x, Inches(4.35), Inches(3.71), Inches(2.0), LIGHT)
        rect(s, x, Inches(4.35), Inches(3.71), Inches(0.1), col)
        txt(s, x + Inches(0.26), Inches(4.58), Inches(3.2), Inches(0.4), [[(h, 13, col, True)]])
        txt(s, x + Inches(0.26), Inches(5.0), Inches(3.2), Inches(1.2), [[(b, 11, INK, False)]])
    footer(s); return s


# ============================================================ DECK CONTENT
TOPIC_ACTS = {t["num"]: [a for a in ACTIVITIES if a["topic"] == t["num"]] for t in C.TOPICS}
ADMIN = "COURSE ADMIN"


def admin_block(closing=False):
    """Digital attendance / TRAQOM. Repeated at the front and the end."""
    s = head(slide(), "Digital Attendance (Mandatory)", ADMIN, kcolor=AMBER)
    tiles = [("AM Attendance", "Scan at the start of each training day."),
             ("PM Attendance", "Scan after the lunch break."),
             ("Assessment Attendance", "Scan before the final assessment begins."),
             ("75% Minimum", "Attendance below 75% forfeits funding and certification.")]
    for i, (h, b) in enumerate(tiles):
        x = Inches(0.85) + (i % 2) * Inches(5.95)
        y = Inches(2.05) + (i // 2) * Inches(2.35)
        col = PALETTE[i % 4]
        rect(s, x, y, Inches(5.68), Inches(2.1), LIGHT)
        rect(s, x, y, Inches(0.1), Inches(2.1), col)
        bd = Inches(0.5)
        oval(s, x + Inches(0.28), y + Inches(0.28), bd, bd, col)
        txt(s, x + Inches(0.28), y + Inches(0.28), bd, bd, [[(str(i + 1), 16, WHITE, True)]],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        txt(s, x + Inches(0.95), y + Inches(0.3), Inches(4.5), Inches(0.45),
            [[(h, 15, col, True)]])
        txt(s, x + Inches(0.32), y + Inches(0.92), Inches(5.1), Inches(1.0),
            [[(b, 12.5, INK, False)]])
    txt(s, Inches(0.85), Inches(6.5), Inches(11.63), Inches(0.5),
        [[("The trainer will display the SSG digital attendance QR code — scan it "
           "with your phone camera and submit.", 12, GREY, False)]])
    footer(s)


def build():
    # ---------------------------------------------------------------- 1. cover
    cover()

    # ---------------------------------------------------------------- admin
    admin_block()

    trainer_slide(ADMIN, "Your Trainer", "WSQ Certified Trainer & Assessor (ACTA / ACLP)",
                  [("Qualifications", "Degree/Masters in a data or engineering discipline; ACLP-certified."),
                   ("Industry Experience", "Hands-on practitioner in data analysis and business reporting."),
                   ("Teaching Focus", "Practical, lab-driven delivery — you build every artefact yourself."),
                   ("Contact", f"{C.EMAIL}  ·  {C.TEL}")],
                  "T", accent=GREY)

    trainer_slide(ADMIN, C.TRAINER, "Principal Trainer, Tertiary Infotech Academy",
                  [("Qualifications", "PhD; ACLP-certified WSQ trainer and assessor."),
                   ("Industry Experience", "20+ years across data analytics, automation and software engineering."),
                   ("Specialisation", "Excel Power Query / Power Pivot, Power BI, Python analytics."),
                   ("Contact", f"{C.EMAIL}  ·  {C.TEL}")],
                  "AA", accent=BLUE)

    tile_grid("Let's Know Each Other", [
        ("Your name", "and the organisation you are from"),
        ("Your role", "and how much Excel you use day to day"),
        ("Your data", "the reports or files you clean up most often"),
        ("Your goal", "what you want to stop doing manually after this course"),
    ], kicker=ADMIN, cols=2, size=15, accent=TEAL)

    tile_grid("Ground Rules", [
        ("Phones on silent", "Step out quietly for calls"),
        ("Participate actively", "No question is a stupid question"),
        ("Mutual respect", "Agree to disagree"),
        ("One conversation", "One person speaks at a time"),
        ("Be punctual", "Return from breaks on time"),
        ("75% attendance", "Required for funding and certification"),
    ], kicker=ADMIN, cols=2, size=14, accent=BLUE)

    # ---- housekeeping / software
    tile_grid("What You Need for This Course", [
        ("Microsoft Excel for Windows", "Microsoft 365, 2021, 2019 or 2016"),
        ("Power Query", "Built in under Data > Get & Transform Data"),
        ("Power Pivot", "File > Options > Add-ins > COM Add-ins > enable"),
        ("Lab files", "One folder per lab, downloaded from the LMS"),
    ], kicker=ADMIN, cols=2, size=15, accent=VIOLET)

    browser_slide("Download Your Course Material", "https://lms-tms.tertiaryinfotech.com",
                  ADMIN,
                  [("Log in", "Use the account e-mailed to you before the course."),
                   ("Open this course", "Excel Power Query and Power Pivot."),
                   ("Download the files", "Slides, Learner Guide, Lesson Plan and lab workbooks."),
                   ("Save locally", "Unzip the labs folder to your Desktop before Lab 01.")],
                  caption="All courseware is on the LMS/TMS portal at "
                          "lms-tms.tertiaryinfotech.com — download it before the first lab.")

    # ---- skills framework
    tile_grid("Skills Framework Alignment", [
        ("TSC Title", C.TSC_TITLE),
        ("TSC Code", C.TSC_CODE),
        ("Proficiency Level", "Level 4 — apply and adapt analysis methods"),
        ("Assessment", "Written Assessment + Practical Performance"),
    ], kicker=ADMIN, cols=2, size=15, accent=BLUE)

    compare_table("TSC Knowledge and Abilities Assessed",
                  ["Code", "What it covers", "Where it is taught"],
                  [["K1", "Operations of data transformation techniques", "Topic 1, 2 · Labs 1–13"],
                   ["K2", "Conditions required for each transformation and join", "Topic 1–3 · Labs 3–19"],
                   ["K3", "Interpretation of query and data-model results", "Topic 2, 3 · Labs 12–19"],
                   ["K4", "Types of data analysis software and capabilities", "Topic 1, 2 · Labs 1, 2, 13"],
                   ["A1", "Review data sets to uncover trends or patterns", "Topic 1 · Labs 1–6"],
                   ["A2", "Develop methods for complex data sets", "Topic 2 · Labs 7–13"],
                   ["A3", "Apply analytics to examine business issues", "Topic 3 · Labs 14–19"]],
                  kicker=ADMIN, accent=VIOLET, first_frac=0.12,
                  note="Every K and A code is taught in a topic AND practised in a lab — "
                       "nothing is assessed that you have not built.")

    ncards("Learning Outcomes", [
        ("LO1 · Review", "Use Power Query to load, clean and reshape data sets so trends "
                         "and patterns become visible."),
        ("LO2 · Develop", "Build methods in Power Query and Power Pivot — logic, dates, "
                          "grouping and DAX measures — for complex data."),
        ("LO3 · Apply", "Apply advanced joins, approximate and fuzzy matching to examine "
                        "real business issues."),
    ], kicker=ADMIN, accent=BLUE, cols=3,
        synthesis=("How you will get there",
                   "Each outcome is taught as a concept, demonstrated by the trainer, then "
                   "built by you in the labs — 19 hands-on labs across the two days."))

    # ---- course outline (one slide per topic)
    for t in C.TOPICS:
        acts = TOPIC_ACTS[t["num"]]
        ncards(f"Course Outline — Topic {t['num']}",
               [(f"Lab {a['num']:02d}", a["title"]) for a in acts],
               kicker=f"TOPIC {t['num']} · {t['weighting']} OF THE COURSE",
               accent=PALETTE[(t["num"] - 1) % 4], cols=4 if len(acts) > 4 else 3,
               synthesis=(t["title"], t["subtitle"]))

    # ---- lesson plan
    compare_table("Lesson Plan — Day 1",
                  ["Time", "Session", "Detail"],
                  [["9:30 – 9:45", "Digital Attendance (AM)", "Scan the SSG QR code"],
                   ["9:45 – 10:15", "Introductions & Course Overview", "Outcomes, outline, assessment"],
                   ["10:15 – 12:30", "Topic 1 (Part 1)", "Power Query concepts · Labs 01–03"],
                   ["12:30 – 13:30", "Lunch Break", "—"],
                   ["13:30 – 13:40", "Digital Attendance (PM)", "Scan the SSG QR code"],
                   ["13:40 – 15:45", "Topic 1 (Part 2)", "Column shaping · Labs 04–06"],
                   ["15:45 – 18:00", "Topic 2 (Part 1)", "Logic and dates · Labs 07–09"],
                   ["18:00 – 18:30", "Day 1 Recap", "Q&A and consolidation"]],
                  kicker=ADMIN, accent=BLUE, first_frac=0.20)

    compare_table("Lesson Plan — Day 2",
                  ["Time", "Session", "Detail"],
                  [["9:30 – 9:40", "Digital Attendance (AM)", "Scan the SSG QR code"],
                   ["9:40 – 12:30", "Topic 2 (Part 2)", "Pivot, Group By, Power Pivot · Labs 10–13"],
                   ["12:30 – 13:30", "Lunch Break", "—"],
                   ["13:30 – 13:40", "Digital Attendance (PM)", "Scan the SSG QR code"],
                   ["13:40 – 16:00", "Topic 3", "Joins and matching · Labs 14–19"],
                   ["16:00 – 16:30", "Revision", "Consolidation and Q&A"],
                   ["16:30 – 16:45", "Course Feedback & TRAQOM", "Mandatory survey"],
                   ["16:45 – 18:30", "Digital Attendance + Final Assessment", "WA 1 hr + PP 1 hr"]],
                  kicker=ADMIN, accent=TEAL, first_frac=0.20)

    # ---- assessment (briefing BEFORE assessment, per house rule)
    tile_grid("Briefing for Assessment", [
        ("Clear your desk", "Phones and materials under the table or on the floor"),
        ("No photos or recording", "Assessment scripts must not be captured"),
        ("No discussion", "Work individually throughout"),
        ("Black or blue pen", "For hard-copy scripts; no correction tape or liquid paper"),
        ("Scripts collected", "All scripts are collected when time is up"),
        ("Open book", "Slides, Learner Guide and approved materials only"),
    ], kicker=ADMIN, cols=2, size=14, accent=AMBER)

    ncards("Final Assessment", [
        ("Written Assessment (WA)", "Short-Answer Questions · 1 hour · open book. "
                                    "Tests your knowledge of Power Query and Power Pivot concepts."),
        ("Practical Performance (PP)", "Hands-on tasks in Excel · 1 hour · open book. "
                                       "You perform the transformations you practised in the labs."),
        ("Result", "Competent / Not Yet Competent. Both instruments must be passed."),
        ("Re-assessment", "One re-assessment opportunity is offered if assessed Not Yet Competent."),
    ], kicker=ADMIN, accent=AMBER, cols=4,
        synthesis=("Open book means",
                   "You may refer to these slides, the Learner Guide and any approved course "
                   "material. You may not use the internet, AI assistants or discuss with others."))

    process_map("Assessment Flow",
                [("Briefing", "Rules explained"),
                 ("Digital Attendance", "Assessment scan"),
                 ("Written Assessment", "SAQ · 1 hour"),
                 ("Practical Performance", "Excel tasks · 1 hour"),
                 ("Marking & Feedback", "Assessor reviews"),
                 ("Sign-off", "Competent / NYC")],
                kicker=ADMIN, color=AMBER,
                synthesis=("Sign-off", "The assessor records the outcome, you sign the "
                                       "assessment record, and results are released through the LMS/TMS portal."))

    # NOTE: no Practice Exam slide — that is required for CERTIFICATION courses
    # (Microsoft/AWS/Google/Cisco/CompTIA). This is a generic WSQ skills course,
    # so its absence is correct per the house standard.

    tile_grid("Criteria for Funding", [
        ("75% attendance", "Based on the SSG digital attendance record"),
        ("Assessed Competent", "Both the WA and the PP must be passed"),
        ("TRAQOM survey", "Mandatory course feedback survey"),
        ("Certificate", "Certificate of Achievement issued on completion"),
    ], kicker=ADMIN, cols=2, size=15, accent=TEAL)

    # ================================================== TOPIC 1
    t = C.TOPICS[0]
    mark("topic1")
    section(f"TOPIC {t['num']}", t["title"], "01", t["subtitle"])

    big_statement("Stop repeating the same clean-up every month.",
                  "Power Query records your transformation once and replays it on every refresh.",
                  "TOPIC 1 · WHY POWER QUERY", color=BLUE)

    process_map("What Power Query Actually Does",
                [("Extract", "Connect to the source"),
                 ("Transform", "Clean and reshape"),
                 ("Load", "Worksheet or Data Model"),
                 ("Refresh", "Replay on new data")],
                kicker="TOPIC 1 · THE ETL LOOP", color=BLUE,
                synthesis=("The point",
                           "You build the recipe ONCE. Next month you replace the source file, "
                           "click Refresh, and every step runs again in the same order."))

    good_bad("Worksheet Formulas vs Power Query",
             ("Power Query", [
                 "Steps are recorded and re-run on Refresh",
                 "Handles millions of rows outside the grid",
                 "The whole clean-up is auditable in Applied Steps",
                 "New rows are picked up automatically",
                 "One query feeds many reports"]),
             ("Manual formulas / copy-paste", [
                 "Must be re-done or dragged down every month",
                 "Slows the workbook down as rows grow",
                 "No record of what was changed, or why",
                 "Breaks silently when the layout shifts",
                 "Each analyst rebuilds their own version"]),
             kicker="TOPIC 1 · WHY IT MATTERS",
             note="Rule of thumb: if you have done it twice by hand, it belongs in Power Query.")

    ribbon_slide("Where Power Query Lives in Excel", "TOPIC 1 · THE RIBBON",
                 ["Data", "Home", "Insert", "Formulas", "Review", "View"],
                 [("Get & Transform Data",
                   ["From\nTable", "From\nFile", "From\nWeb", "Get\nData"]),
                  ("Queries & Connections",
                   ["Refresh\nAll", "Queries", "Properties"])],
                 [("Data > From Table/Range",
                   "Loads the selected worksheet table into the Power Query Editor."),
                  ("Data > Get Data",
                   "Connects to an external workbook, CSV, database, folder or web source."),
                  ("Queries & Connections",
                   "The side pane listing every query, its load state and its row count.")],
                 accent=BLUE)

    tile_grid("The Power Query Editor — Five Areas You Must Know", [
        ("1. The Ribbon", "Home, Transform, Add Column and View tabs"),
        ("2. Queries Pane", "Every query in the workbook, on the left"),
        ("3. Data Preview", "The first 1000 rows — a preview, not the full data"),
        ("4. Query Settings", "Properties and the Applied Steps list"),
        ("5. Status Bar", "Column count, row count and profiling info"),
        ("Applied Steps", "The heart of it — your recorded, editable recipe"),
    ], kicker="TOPIC 1 · THE EDITOR", cols=2, size=14, accent=BLUE)

    takeaway("Applied Steps — Your Audit Trail",
             "Every click you make becomes a step you can rename, reorder, edit or delete.",
             kicker="TOPIC 1 · APPLIED STEPS", accent=TEAL,
             points=[("Order matters",
                      "Steps run top to bottom. Filtering before a merge is fast; after it is slow."),
                     ("Everything is editable",
                      "Click the gear icon on a step to reopen the dialog you used."),
                     ("Delete with care",
                      "Removing an early step can break the ones that follow it.")])

    # data types
    compare_table("Data Types — Set Them in the Query, Not the Grid",
                  ["Type", "What it holds", "Common failure"],
                  [["Text", "Names, codes, IDs, postal codes", "Leading zeros lost if typed as number"],
                   ["Whole Number", "Counts and quantities", "Will not sum while stored as text"],
                   ["Decimal Number", "Prices, amounts, rates", "Currency prefix blocks the type change"],
                   ["Date", "Calendar dates", "dd/mm/yyyy read as US format — use Using Locale"],
                   ["Date/Time", "Timestamps", "Time zone shifts the date by a day"],
                   ["True/False", "Logical flags", "The words TRUE/FALSE stored as text"],
                   ["Any", "Type not yet set", "Excel guesses — and eventually guesses wrong"]],
                  kicker="TOPIC 1 · DATA TYPES", accent=VIOLET, first_frac=0.20,
                  note="A wrong data type is the single most common cause of a broken refresh. "
                       "Set every column's type explicitly as the FIRST thing you do.")

    decision_map("Which Load Option Should I Choose?",
                 "Do you need to SEE this table on a worksheet?",
                 ("Load To > Table",
                  "Loads to a worksheet. Use for the final, report-ready output only."),
                 ("Only Create Connection",
                  "Keeps it out of the grid. Use for staging queries that feed a merge or append."),
                 kicker="TOPIC 1 · LOADING", color=VIOLET,
                 note="Tick 'Add this data to the Data Model' as well when the table will be "
                      "used in Power Pivot. Loading everything to worksheets is the fastest way "
                      "to bloat a workbook.")

    ncards("Getting Data — The Sources You Will Use", [
        ("From Table/Range", "The data already on a worksheet. The starting point for Lab 01."),
        ("From Workbook", "Another Excel file — the pattern in Lab 02 for the monthly price list."),
        ("From Text/CSV", "Exports from any system that can write a delimited file."),
        ("From Folder", "Every file in a folder combined into one table — the monthly-file pattern."),
    ], kicker="TOPIC 1 · GET DATA", accent=BLUE, cols=4,
        synthesis=("The connection is live",
                   "Power Query stores the path, not a copy. Replace the source file with a newer "
                   "one of the same shape and Refresh brings the new data straight through."))

    mark("lab01")
    activity_overview(ACTIVITIES[0], "TOPIC 1 · HANDS-ON LAB")
    mark("lab02")
    activity_overview(ACTIVITIES[1], "TOPIC 1 · HANDS-ON LAB")

    # fill
    takeaway("Fill Down and Fill Up",
             "Reports written for humans leave gaps. Fill turns them back into real data.",
             kicker="TOPIC 1 · FILL", accent=TEAL,
             points=[("Fill Down",
                      "Carries a value INTO the blank rows beneath it — the classic merged-cell repair."),
                     ("Fill Up",
                      "Carries a value into the blanks ABOVE it, for labels that sit under their group."),
                     ("Watch for null vs blank",
                      "Fill only replaces null. Cells holding an empty string need Replace Values first.")])

    mark("lab03")
    activity_overview(ACTIVITIES[2], "TOPIC 1 · HANDS-ON LAB")

    # split
    compare_table("Splitting a Column — Four Ways",
                  ["Method", "What it does", "Use it when"],
                  [["By Delimiter", "Cuts at a chosen character", "Names, codes, comma-separated lists"],
                   ["By Number of Characters", "Cuts at a fixed width", "Fixed-width legacy exports"],
                   ["By Positions", "Cuts at named positions", "Mainframe or COBOL-style records"],
                   ["By Lowercase to Uppercase", "Cuts at a case change", "camelCase or PascalCase fields"]],
                  kicker="TOPIC 1 · SPLIT COLUMN", accent=TEAL, first_frac=0.28,
                  note="Split into COLUMNS makes the row wider. Split into ROWS makes the table "
                       "taller — that is how you normalise a comma-separated list into one record "
                       "per value.")

    good_bad("Split Into Columns vs Split Into Rows",
             ("Split into Rows", [
                 "One product per row — a tidy, analysable table",
                 "Row count grows; column count stays fixed",
                 "PivotTables and merges work directly on it",
                 "The right choice for a repeating list"]),
             ("Split into Columns", [
                 "Product 1, Product 2, Product 3 … columns",
                 "Breaks the moment someone buys a fourth item",
                 "Every downstream formula must be rewritten",
                 "Only correct for a FIXED number of parts (e.g. first/last name)"]),
             kicker="TOPIC 1 · THE CHOICE THAT MATTERS",
             note="Ask: can the number of parts change? If yes, split into rows.")

    mark("lab04")
    activity_overview(ACTIVITIES[3], "TOPIC 1 · HANDS-ON LAB")

    # add columns
    ncards("Three Ways to Add a Column", [
        ("Conditional Column", "A guided dialog that writes if/then/else for you. "
                               "Best for banding and classification rules."),
        ("Custom Column", "You write the M expression yourself. Needed for anything "
                          "the dialog cannot express."),
        ("Column From Examples", "You type the answer you want on the first row and "
                                 "Power Query infers the pattern."),
        ("Index Column", "A sequential row number — useful for ranking and for "
                         "restoring an original order after sorting."),
    ], kicker="TOPIC 1 · ADD COLUMN", accent=VIOLET, cols=4,
        synthesis=("Which one first?",
                   "Try the Conditional Column dialog first, then Column From Examples, and drop "
                   "to a Custom Column only when neither can express the rule. All three produce "
                   "the same kind of Applied Step."))

    code_example("Conditional Column — What the Dialog Writes",
                 "The Value Band rule from Lab 05, as the dialog builds it in M.",
                 ['if [Amount] >= 5000 then "High"',
                  'else if [Amount] >= 1000 then "Medium"',
                  'else "Low"'],
                 [("Order is everything",
                   "The FIRST matching rule wins. Put the highest threshold at the top."),
                  ("Test the boundary",
                   "Amount exactly 5000 is High because the test is >=, not >."),
                  ("else is compulsory",
                   "Every if needs a final else, or the query returns an error."),
                  ("Set the type after",
                   "A new column arrives as Any — set it to Text or Decimal explicitly.")],
                 kicker="TOPIC 1 · CONDITIONAL COLUMN", accent=VIOLET, lang="M")

    mark("lab05")
    activity_overview(ACTIVITIES[4], "TOPIC 1 · HANDS-ON LAB")

    takeaway("Sort and Filter INSIDE the Query",
             "Filtering at source means Excel never loads the rows you did not want.",
             kicker="TOPIC 1 · SORT & FILTER", accent=BLUE,
             points=[("Filter early",
                      "A filter placed before a merge or group is dramatically faster than one after."),
                     ("Filters are permanent",
                      "Unlike a worksheet AutoFilter, a query filter removes rows from the result."),
                     ("Sort is a step too",
                      "Sorting by two columns records two steps — check the order in Applied Steps.")])

    mark("lab06")
    activity_overview(ACTIVITIES[5], "TOPIC 1 · HANDS-ON LAB")

    tile_grid("Topic 1 Recap", [
        ("ETL, recorded", "Extract, Transform, Load — replayed on every Refresh"),
        ("Types first", "Set every column's data type before anything else"),
        ("Connection Only", "Keep staging queries out of the worksheet grid"),
        ("Fill repairs gaps", "Fill Down and Fill Up rebuild machine-readable tables"),
        ("Split by rows", "When the number of parts can change"),
        ("Three add-column tools", "Conditional, Custom, and From Examples"),
    ], kicker="TOPIC 1 · RECAP", cols=2, size=14, accent=BLUE)

    brk("Lunch Break", "1 hour  ·  Digital Attendance (PM) on return")

    # ================================================== TOPIC 2
    t = C.TOPICS[1]
    mark("topic2")
    section(f"TOPIC {t['num']}", t["title"], "02", t["subtitle"])

    big_statement("Cleaning data is only half the job.",
                  "Now you add the logic, the dates and the aggregation that answer the question.",
                  "TOPIC 2 · FROM CLEAN TO USEFUL", color=TEAL)

    # M language
    compare_table("Excel Formulas vs Power Query M",
                  ["", "Excel worksheet", "Power Query M"],
                  [["Syntax", "IF(A2>=500, 0, 15)", "if [Amount] >= 500 then 0 else 15"],
                   ["Case", "IF, AND, OR — not case sensitive", "if, and, or — MUST be lowercase"],
                   ["Refers to", "Cell addresses (A2)", "Column names in [square brackets]"],
                   ["else", "Optional third argument", "Compulsory — every if needs an else"],
                   ["Runs", "Once per cell, on every calculation", "Once per row, on refresh only"],
                   ["Text join", "&  or CONCAT()", "&  (text only — numbers need Text.From)"]],
                  kicker="TOPIC 2 · THE M LANGUAGE", accent=TEAL, first_frac=0.16,
                  note="M is case sensitive throughout. 'If' or 'IF' will not work — only 'if'. "
                       "This is the single most common error when writing your first custom column.")

    code_example("Writing IF Logic in M",
                 "The Shipping Cost rule from Lab 07 — written by hand in a Custom Column.",
                 ['if [Amount] >= 500 then 0',
                  'else 15',
                  '',
                  '// a comparison alone returns true/false',
                  '[Amount] >= 500'],
                 [("Lowercase keywords",
                   "if, then, else — capitalising any of them is an immediate error."),
                  ("Columns in brackets",
                   "[Amount] refers to the column. Spelling and case must match exactly."),
                  ("No else = error",
                   "Delete the else and Power Query refuses to build the column."),
                  ("Comparisons are values",
                   "[Amount] >= 500 on its own returns TRUE/FALSE — set the type to True/False.")],
                 kicker="TOPIC 2 · IF IN M", accent=TEAL, lang="M")

    compare_table("Nesting Conditions — the AND / OR Truth Table",
                  ["Amount >= 5000", "Days Outstanding > 30", "and →", "or →"],
                  [["TRUE", "TRUE", "TRUE (Escalate)", "TRUE"],
                   ["TRUE", "FALSE", "FALSE", "TRUE (Watch)"],
                   ["FALSE", "TRUE", "FALSE", "TRUE (Watch)"],
                   ["FALSE", "FALSE", "FALSE", "FALSE (Normal)"]],
                  kicker="TOPIC 2 · AND vs OR", accent=VIOLET, first_frac=0.25,
                  note="'and' needs BOTH sides true; 'or' needs EITHER. Lab 08 builds exactly this "
                       "rule — rows SO-9201 to SO-9204 cover all four lines of this table.")

    mark("lab07")
    activity_overview(ACTIVITIES[6], "TOPIC 2 · HANDS-ON LAB")
    mark("lab08")
    activity_overview(ACTIVITIES[7], "TOPIC 2 · HANDS-ON LAB")

    # dates
    ncards("Date Functions You Will Actually Use", [
        ("Date.AddDays", "Date.AddDays([Invoice Date], 30) — the due date from payment terms."),
        ("Duration.Days", "Turns the gap between two dates into a whole number of days."),
        ("Date.Year / Month", "Pulls the year or month out for grouping and reporting."),
        ("Date.From", "Converts text or a date/time into a clean Date value."),
    ], kicker="TOPIC 2 · DATES", accent=AMBER, cols=4,
        synthesis=("The worked example in Lab 09",
                   "Invoice INV-9001 is dated 1 May 2026 with 30-day terms. Date.AddDays gives a "
                   "Due Date of 31 May 2026; against a 30 June 2026 reporting date that is exactly "
                   "30 days overdue."))

    process_map("From Invoice Date to Ageing Bucket",
                [("Invoice Date", "The raw source column"),
                 ("+ Terms Days", "Date.AddDays"),
                 ("Due Date", "When payment is expected"),
                 ("Days Overdue", "Duration.Days vs today"),
                 ("Ageing Bucket", "Conditional column")],
                kicker="TOPIC 2 · THE DEADLINE PIPELINE", color=AMBER,
                synthesis=("Why build it in the query",
                           "Every one of these columns recalculates on Refresh. Next month's "
                           "invoice file drops in and the whole ageing report rebuilds itself."))

    mark("lab09")
    activity_overview(ACTIVITIES[8], "TOPIC 2 · HANDS-ON LAB")

    # unpivot
    big_statement("Wide is for humans. Tall is for machines.",
                  "Unpivot is the single most valuable transformation in Power Query.",
                  "TOPIC 2 · UNPIVOT", color=VIOLET)

    good_bad("Why Unpivot Matters",
             ("Tall / tidy — 3 columns", [
                 "Region · Month · Sales",
                 "One row per observation",
                 "PivotTables slice it any way you like",
                 "A 13th month adds ROWS, not columns",
                 "Merges and the Data Model work directly"]),
             ("Wide / cross-tab — 13 columns", [
                 "Region · Jan · Feb · Mar … Dec",
                 "One row per region, values spread across",
                 "A PivotTable cannot group by month",
                 "A 13th month breaks every formula",
                 "Readable for people, useless for analysis"]),
             kicker="TOPIC 2 · SHAPE",
             note="Lab 10 turns 5 regions x 12 month columns into 60 tidy rows.")

    compare_table("Three Unpivot Commands — Choose Carefully",
                  ["Command", "What it unpivots", "Survives a new column?"],
                  [["Unpivot Columns", "Only the columns you selected", "No — the new month is ignored"],
                   ["Unpivot Other Columns", "Everything EXCEPT your selection", "Yes — this is the safe one"],
                   ["Unpivot Only Selected", "Only that selection, fixed forever", "No — silently misses it"]],
                  kicker="TOPIC 2 · UNPIVOT WARNINGS", accent=VIOLET, first_frac=0.30,
                  note="Select the columns you want to KEEP, then Unpivot Other Columns. Your query "
                       "then survives next year's extra month without any edit.")

    mark("lab10")
    activity_overview(ACTIVITIES[9], "TOPIC 2 · HANDS-ON LAB")

    takeaway("Pivot — the Reverse Trip",
             "Pivot Column turns rows back into columns. The Don't Aggregate option is the one that matters.",
             kicker="TOPIC 2 · PIVOT", accent=BLUE,
             points=[("Default: aggregate",
                      "Pivot sums or counts the values — right for numbers, wrong for text."),
                     ("Don't Aggregate",
                      "Keeps each individual value. Use it for attributes like Warranty or Supplier."),
                     ("Error if duplicated",
                      "Don't Aggregate fails when a key repeats — there would be two values for one cell.")])

    takeaway("Append — Stacking Tables",
             "Append adds ROWS. Merge adds COLUMNS. Confusing the two is the classic beginner error.",
             kicker="TOPIC 2 · APPEND", accent=TEAL,
             points=[("Column names must match",
                      "Exactly — including case and spaces. A mismatch creates a new column of nulls."),
                     ("Two or more tables",
                      "'Append Queries as New' keeps the sources intact and creates a third query."),
                     ("The monthly-file pattern",
                      "Append Q1 and Q2, or point From Folder at a folder of identical files.")])

    mark("lab11")
    activity_overview(ACTIVITIES[10], "TOPIC 2 · HANDS-ON LAB")

    # group by
    ncards("Group By — Aggregating Inside the Query", [
        ("Count Rows", "How many orders per region — the most common aggregation."),
        ("Sum", "Total sales value per group."),
        ("Average / Min / Max", "Typical, smallest and largest order in each group."),
        ("All Rows", "Keeps the underlying detail rows nested for later drill-through."),
    ], kicker="TOPIC 2 · GROUP BY", accent=BLUE, cols=4,
        synthesis=("Basic vs Advanced",
                   "Basic groups by one column with one aggregation. Switch to ADVANCED to group "
                   "by several columns and add several aggregations at once — that is what Lab 12 "
                   "does with 400 transactions."))

    mark("lab12")
    activity_overview(ACTIVITIES[11], "TOPIC 2 · HANDS-ON LAB")

    # power pivot
    big_statement("Power Pivot is where the tables meet.",
                  "Power Query shapes each table; the Data Model relates them and DAX measures them.",
                  "TOPIC 2 · POWER PIVOT", color=VIOLET)

    process_map("Power Query and Power Pivot — the Division of Labour",
                [("Power Query", "Get and shape each table"),
                 ("Data Model", "Load tables in, no grid"),
                 ("Relationships", "Join fact to lookup"),
                 ("DAX Measures", "Calculate on demand"),
                 ("PivotTable", "Slice and report")],
                kicker="TOPIC 2 · THE FULL STACK", color=VIOLET,
                synthesis=("The rule",
                           "Shape in Power Query. Relate and calculate in Power Pivot. Doing "
                           "row-level clean-up in DAX, or trying to relate tables in M, fights the tools."))

    compare_table("Star Schema — Fact and Lookup Tables",
                  ["Table", "Kind", "What it holds"],
                  [["Sales", "Fact", "One row per transaction — the numbers you measure"],
                   ["Products", "Lookup / Dimension", "One row per product — how you slice"],
                   ["Calendar", "Lookup / Dimension", "One row per date — time intelligence"]],
                  kicker="TOPIC 2 · THE DATA MODEL", accent=BLUE, first_frac=0.24,
                  note="Relationships run ONE-TO-MANY from the lookup table to the fact table. "
                       "Lab 13 builds exactly this three-table model.")

    good_bad("Implicit vs Explicit Measures",
             ("Explicit measure (do this)", [
                 "Written once in the Calculation Area",
                 "Total Sales := SUM(Sales[Amount])",
                 "Reusable in every PivotTable and other measures",
                 "Renaming it updates everywhere",
                 "The professional practice"]),
             ("Implicit measure (avoid)", [
                 "Created by dragging a field into Values",
                 "Excel silently picks SUM for you",
                 "Cannot be reused or referenced by name",
                 "Disappears when the field is removed",
                 "Fine for a quick look, wrong for a model"]),
             kicker="TOPIC 2 · DAX MEASURES",
             note="Always write explicit measures. Lab 13 creates four of them.")

    code_example("The Four DAX Measures You Build in Lab 13",
                 "Written in the Power Pivot Calculation Area, one measure per cell.",
                 ['Total Sales := SUM(Sales[Amount])',
                  '',
                  'Order Count := COUNTROWS(Sales)',
                  '',
                  'Average Order :=',
                  '    DIVIDE([Total Sales], [Order Count])',
                  '',
                  'Sales YTD :=',
                  '    TOTALYTD([Total Sales], Calendar[Date])'],
                 [("SUM aggregates a column",
                   "It respects whatever filters the PivotTable applies."),
                  ("COUNTROWS counts the table",
                   "Not a column — so blank values still count as rows."),
                  ("DIVIDE, never /",
                   "DIVIDE returns blank on a divide-by-zero instead of an error."),
                  ("TOTALYTD needs a date table",
                   "Mark Calendar as a Date Table first, or time intelligence returns blank.")],
                 kicker="TOPIC 2 · DAX", accent=VIOLET, lang="DAX")

    mark("lab13")
    activity_overview(ACTIVITIES[12], "TOPIC 2 · HANDS-ON LAB")

    tile_grid("Topic 2 Recap", [
        ("M is case sensitive", "if / then / else / and / or are always lowercase"),
        ("Every if needs an else", "There is no optional third argument in M"),
        ("Date.AddDays", "Derives deadlines that recalculate on refresh"),
        ("Unpivot Other Columns", "The version that survives a new column"),
        ("Append = rows, Merge = columns", "The most common mix-up"),
        ("Explicit DAX measures", "Written once, reusable everywhere"),
    ], kicker="TOPIC 2 · RECAP", cols=2, size=14, accent=TEAL)

    brk("Tea Break", "15 minutes")

    # ================================================== TOPIC 3
    t = C.TOPICS[2]
    mark("topic3")
    section(f"TOPIC {t['num']}", t["title"], "03", t["subtitle"])

    big_statement("Choosing the wrong join changes your row count silently.",
                  "No error, no warning — just a different answer. This topic makes the choice deliberate.",
                  "TOPIC 3 · JOINS", color=AMBER)

    process_map("How a Merge Works",
                [("Two queries", "A left and a right table"),
                 ("Matching columns", "The key you join on"),
                 ("Join kind", "Which rows survive"),
                 ("Nested table", "One column of tables"),
                 ("Expand", "Pick the columns you need")],
                kicker="TOPIC 3 · THE MERGE PIPELINE", color=BLUE,
                synthesis=("The step everyone forgets",
                           "A merge produces a COLUMN OF TABLES. Nothing useful appears until you "
                           "click the expand arrow and choose which columns to bring through."))

    compare_table("The Six Join Kinds",
                  ["Join kind", "Which rows you keep", "Typical use"],
                  [["Left Outer", "ALL left rows + matches from the right", "The safe default — an XLOOKUP replacement"],
                   ["Right Outer", "ALL right rows + matches from the left", "Rare — usually swap the tables instead"],
                   ["Full Outer", "Every row from BOTH tables", "Reconciling two systems"],
                   ["Inner", "Only rows that match on BOTH sides", "Confirmed, complete records only"],
                   ["Left Anti", "Left rows with NO match on the right", "Find orphans — orders with a bad product ID"],
                   ["Right Anti", "Right rows with NO match on the left", "Find unused lookups — products never sold"]],
                  kicker="TOPIC 3 · ALL SIX", accent=AMBER, first_frac=0.18,
                  note="Lab 14 runs all six on the SAME pair of tables and records the row count "
                       "each one returns — that comparison is your reference sheet for every merge.")

    decision_map("Which Join Do I Need?",
                 "Must every row of my main table survive?",
                 ("Yes → Left Outer",
                  "Keeps all your records. Unmatched rows show null. The XLOOKUP equivalent."),
                 ("No → Inner or Anti",
                  "Inner keeps only matched rows. Left Anti keeps only the UNMATCHED ones."),
                 kicker="TOPIC 3 · CHOOSING", color=AMBER,
                 note="A useful check: Inner rows + Left Anti rows = Left Outer rows. If your merge "
                      "returns MORE rows than the left table, your key is not unique.")

    good_bad("Merge vs XLOOKUP",
             ("Merge (Power Query)", [
                 "Runs on Refresh, no dragging down",
                 "Handles hundreds of thousands of rows",
                 "Unmatched rows return null",
                 "Can match on MULTIPLE columns",
                 "Fuzzy matching available"]),
             ("XLOOKUP (worksheet)", [
                 "Must be copied down as rows are added",
                 "Slows the workbook as the range grows",
                 "Unmatched rows return #N/A",
                 "Needs a helper column for composite keys",
                 "Exact or approximate only — no fuzzy"]),
             kicker="TOPIC 3 · THE COMPARISON",
             note="XLOOKUP is still the right tool for a one-off answer in a small sheet. "
                  "Lab 15 replaces a column of XLOOKUPs with a single merge.")

    mark("lab14")
    activity_overview(ACTIVITIES[13], "TOPIC 3 · HANDS-ON LAB")
    mark("lab15")
    activity_overview(ACTIVITIES[14], "TOPIC 3 · HANDS-ON LAB")

    takeaway("Merging on More Than One Column",
             "When the key is a combination — product AND region — you must select both, in the same order, in both tables.",
             kicker="TOPIC 3 · COMPOSITE KEYS", accent=VIOLET,
             points=[("Ctrl-click to add",
                      "Click the first column, then Ctrl-click the second. The order is numbered."),
                     ("Order must match",
                      "Product-then-Region on the left must meet Product-then-Region on the right."),
                     ("Watch the row count",
                      "Matching on too few columns multiplies rows — Lab 16 shows a 5x inflation.")])

    mark("lab16")
    activity_overview(ACTIVITIES[15], "TOPIC 3 · HANDS-ON LAB")

    big_statement("Power Query merges are EXACT match only.",
                  "VLOOKUP's approximate match must be rebuilt — with binning or a banding table.",
                  "TOPIC 3 · APPROXIMATE MATCH", color=TEAL)

    compare_table("Two Ways to Rebuild an Approximate Match",
                  ["", "Conditional column (binning)", "Banding table + merge"],
                  [["How", "if/then/else on the thresholds", "A table of Lower/Upper/Rate rows"],
                   ["Bands live", "Inside the query", "In a worksheet table anyone can edit"],
                   ["To change a band", "Edit the query step", "Edit a cell and Refresh"],
                   ["Best for", "Two or three stable bands", "Bands that change, or many of them"],
                   ["Lab 17 builds", "Rate A — the quick version", "The data-driven version"]],
                  kicker="TOPIC 3 · BINNING", accent=TEAL, first_frac=0.18,
                  note="Both must agree at the boundaries. Lab 17 includes reps at exactly 10,000 "
                       "and exactly 50,000 to prove it.")

    mark("lab17")
    activity_overview(ACTIVITIES[16], "TOPIC 3 · HANDS-ON LAB")

    ncards("Cross Join and Drill Down", [
        ("Cross join = every pair", "Add a custom column whose formula is the other query, "
                                    "then expand. 5 regions x 13 products = 65 rows."),
        ("Why you want it", "A report with no gaps — combinations with zero sales still "
                            "appear, showing 0 instead of vanishing."),
        ("The danger", "1,000 x 1,000 rows is a million-row table. Only cross join "
                       "small lookup tables."),
        ("Drill Down", "Right-click a single cell > Drill Down. The query now returns "
                       "ONE value you can reuse as a parameter."),
    ], kicker="TOPIC 3 · CROSS JOIN", accent=BLUE, cols=4,
        synthesis=("Lab 18 combines both",
                   "Build the complete region x product grid with a cross join, fill the gaps with "
                   "zero, then drill down to a threshold value that drives a filter step."))

    mark("lab18")
    activity_overview(ACTIVITIES[17], "TOPIC 3 · HANDS-ON LAB")

    big_statement("Real customer names never match exactly.",
                  "'Pte Ltd' vs 'Pte. Ltd.' vs a typo — fuzzy matching scores similarity instead of demanding equality.",
                  "TOPIC 3 · FUZZY MATCHING", color=VIOLET)

    compare_table("Tuning the Similarity Threshold",
                  ["Threshold", "Behaviour", "Risk"],
                  [["1.00", "Exact match only — fuzzy is effectively off", "Misses every variant"],
                   ["0.90", "Very close spellings only", "Still misses common abbreviations"],
                   ["0.80", "The recommended starting point", "Balanced — few false matches"],
                   ["0.60", "Loosely similar text matches", "False matches appear"],
                   ["0.40", "Almost anything matches", "Unusable — wrong data, silently"]],
                  kicker="TOPIC 3 · THRESHOLD", accent=VIOLET, first_frac=0.18,
                  note="Lower the threshold and you catch more real matches but also more wrong ones. "
                       "Start at 0.80, inspect the result, and adjust deliberately.")

    ncards("Fuzzy Matching Options", [
        ("Similarity threshold", "0 to 1. How alike two values must be before they pair. "
                                 "0.80 is the sensible default."),
        ("Ignore case", "Treats EVERGREEN FOODS and Evergreen Foods as the same."),
        ("Ignore space", "Handles double spaces and stray padding."),
        ("Transformation table", "A two-column From/To table teaching pairs no score can "
                                 "catch — 'ABC Group' means 'ABC Holdings'."),
    ], kicker="TOPIC 3 · OPTIONS", accent=VIOLET, cols=4,
        synthesis=("Why the transformation table is essential",
                   "No similarity score will ever link 'ABC Holdings' to 'ABC Group' — the words "
                   "are genuinely different. Only an explicit mapping can. Lab 19 supplies one."))

    mark("lab19")
    activity_overview(ACTIVITIES[18], "TOPIC 3 · HANDS-ON LAB")

    tile_grid("Topic 3 Recap", [
        ("Six join kinds", "Left/Right/Full Outer, Inner, Left/Right Anti"),
        ("Left Outer is the default", "It keeps every row of your main table"),
        ("Anti joins find problems", "Orphan records and unused lookups"),
        ("Composite keys", "Ctrl-click both columns, same order, both tables"),
        ("Approximate match", "Rebuild with binning or a banding table"),
        ("Fuzzy + transformation table", "Score similarity, then teach the exceptions"),
    ], kicker="TOPIC 3 · RECAP", cols=2, size=14, accent=AMBER)

    # ================================================== CLOSING
    mark("revision")
    section("REVISION", "Putting It All Together", "04",
            "The full pipeline, from raw file to answered business question")

    process_map("The Complete Workflow You Have Learned",
                [("Get", "Table, workbook, folder"),
                 ("Clean", "Types, fill, split, trim"),
                 ("Enrich", "Conditional and custom columns"),
                 ("Combine", "Merge and append"),
                 ("Model", "Relationships and DAX"),
                 ("Report", "PivotTable and refresh")],
                kicker="REVISION · END TO END", color=BLUE,
                synthesis=("What you can now do",
                           "Take a messy monthly extract and turn it into a refreshable, "
                           "auditable report that rebuilds itself with one click next month."))

    tile_grid("The Pitfalls That Cost Most Marks", [
        ("Wrong data type", "Set every type explicitly, first"),
        ("Capitalised If/Then", "M keywords are always lowercase"),
        ("Missing else", "Every if in M needs one"),
        ("Unpivot Only Selected", "Use Unpivot Other Columns instead"),
        ("Merging on too few columns", "Row count inflates silently"),
        ("Forgetting to expand", "A merge yields a column of tables"),
    ], kicker="REVISION · COMMON ERRORS", cols=2, size=14, accent=RED)

    tile_grid("Where to Go Next", [(r, "") for r in C.RECOMMENDED[:4]],
              kicker="REVISION · RECOMMENDED COURSES", cols=2, size=14, accent=TEAL)

    # closing admin block: Assessment → Assessment Flow → Digital Attendance → Thank You
    ncards("Final Assessment", [
        ("Written Assessment (WA)", "Short-Answer Questions · 1 hour · open book."),
        ("Practical Performance (PP)", "Hands-on Excel tasks · 1 hour · open book."),
        ("Open book means", "Slides, Learner Guide and approved materials — no internet, "
                            "no AI tools, no discussion."),
        ("Outcome", "Competent / Not Yet Competent, with one re-assessment opportunity."),
    ], kicker="ASSESSMENT", accent=AMBER, cols=4,
        synthesis=("Before you start",
                   "Clear your desk, silence your phone, and use a black or blue pen for the "
                   "hard-copy script."))

    process_map("Assessment Flow",
                [("Briefing", "Rules explained"),
                 ("Digital Attendance", "Assessment scan"),
                 ("Written Assessment", "SAQ · 1 hour"),
                 ("Practical Performance", "Excel tasks · 1 hour"),
                 ("Marking & Feedback", "Assessor reviews"),
                 ("Sign-off", "Competent / NYC")],
                kicker="ASSESSMENT", color=AMBER,
                synthesis=("Sign-off", "The assessor records the outcome, you sign the assessment "
                                       "record, and the result is released through the LMS/TMS portal."))

    admin_block(closing=True)

    browser_slide("Certificate & TRAQOM Survey (Mandatory)",
                  "https://lms-tms.tertiaryinfotech.com", "COURSE CLOSING",
                  [("Log in to the portal", "Use your course account."),
                   ("Complete the TRAQOM survey", "Mandatory SSG course-quality feedback."),
                   ("Download your certificate", "Issued once assessed Competent."),
                   ("Keep the courseware", "Slides, guide and labs stay available to you.")],
                  caption="The TRAQOM survey is mandatory for WSQ funded courses — please "
                          "complete it before you leave.",
                  accent=TEAL)

    tile_grid("Support", [
        ("Email", C.EMAIL),
        ("Telephone", C.TEL),
        ("Website", C.WEBSITE),
        ("LMS / TMS", "lms-tms.tertiaryinfotech.com"),
    ], kicker="COURSE CLOSING", cols=2, size=15, accent=BLUE)

    s = slide(); rect(s, 0, 0, SW, SH, WHITE)
    rect(s, 0, 0, SW, Inches(0.22), BLUE); rect(s, 0, Inches(7.28), SW, Inches(0.22), TEAL)
    txt(s, Inches(1.1), Inches(2.6), Inches(11.2), Inches(1.6), [[("Thank You!", 54, INK, True)]])
    txt(s, Inches(1.12), Inches(4.1), Inches(11.2), Inches(0.7),
        [[("We hope Power Query and Power Pivot save you many hours a month.", 20, GREY, False)]])
    rect(s, Inches(1.12), Inches(4.95), Inches(2.4), Inches(0.06), TEAL)
    txt(s, Inches(1.1), Inches(5.3), Inches(11.2), Inches(1.0),
        [[(f"{C.ORG}  ·  {C.UEN}", 14, GREY, False)],
         [(f"{C.WEBSITE}  ·  {C.EMAIL}  ·  {C.TEL}", 13, GREY, False)]], space=6)
    footer(s)


# ============================================================ run
build()

# restrained motion: content fades, dividers push — applied in one pass
for i, s in enumerate(prs.slides):
    _transition(s, "fade", "fast")

OUT = os.path.join(REPO, "courseware", f"{C.SHORT_TITLE}-{C.VERSION}.pptx")
os.makedirs(os.path.dirname(OUT), exist_ok=True)
prs.save(OUT)

SLIDE_MAP["_last"] = PAGE["n"]
with open(os.path.join(HERE, "slide_map.json"), "w") as f:
    json.dump(SLIDE_MAP, f, indent=2)

print(f"slides: {len(prs.slides.__iter__.__self__._sldIdLst)}")
print(f"saved:  {OUT}")
