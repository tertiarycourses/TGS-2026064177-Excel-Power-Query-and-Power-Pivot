#!/usr/bin/env python3
"""
Lesson Plan (LP) — Excel Power Query and Power Pivot (TGS-2026064177).

House format: WSQ cover page, Document Version Control Record, real Word TOC
field, Arial 11pt body, colour-coded daily schedule tables, and a copyright +
Page X of Y footer on every page.

Each training day totals exactly 8 instructional hours (9:30am-6:30pm with a
1-hour lunch; tea breaks are counted within the instructional time).

Slide references come from slide_map.json, written by build_slides.py, so the
LP's slide numbers can never drift from the deck.
"""
import os, sys, json
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import course_data as C
from data_domain1 import DOMAIN1
from data_domain2 import DOMAIN2
from data_domain3 import DOMAIN3
from prodoc import (add_cover_page, add_version_control, add_toc, add_page_numbers,
                    style_headings, enable_update_fields, _shade_cell,
                    BRAND, DARK, GREY)

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

try:
    SLIDES = json.load(open(os.path.join(HERE, "slide_map.json")))
except Exception:
    SLIDES = {}


DECK_LAST = SLIDES.get("_last", 0)


def sref(key, fallback=""):
    n = SLIDES.get(key)
    return f"Slide {n}" if n else fallback


def srange(a, b, end_inclusive=False):
    """Slide range from marker a up to marker b. With end_inclusive, b's own
    slide is the last one covered rather than the one before it."""
    x = SLIDES.get(a)
    y = DECK_LAST if b == "end" else SLIDES.get(b)
    if not x or not y:
        return sref(a)
    last = y if (end_inclusive or b == "end") else y - 1
    return f"Slides {x}–{last}" if last > x else f"Slide {x}"


# ------------------------------------------------------------------ schedule
# Each day: 9:30-18:30 with a 1-hour lunch = 8 instructional hours.
DAY1 = [
    ("9:30 – 9:45", "0:15", "Digital Attendance (AM) + Welcome",
     "Trainer welcomes learners and administers the mandatory SSG digital "
     "attendance. Housekeeping and safety briefing.",
     "Trainer-led · Demonstration", "Slides 2–8"),
    ("9:45 – 10:15", "0:30", "Introductions, Outcomes and Course Outline",
     "Trainer and learner introductions. Learning outcomes, TSC alignment, "
     "course outline, lesson plan and assessment briefing are presented.",
     "Trainer-led · Class discussion", "Slides 9–20"),
    ("10:15 – 11:15", "1:00", "Topic 1: Power Query Fundamentals",
     "What Power Query is and the ETL loop it automates. Power Query versus "
     "worksheet formulas. The ribbon, the Editor's five areas, and Applied "
     "Steps as an audit trail. Data types and why they must be set first.",
     "Trainer-led · Demonstration · Q&A", srange("topic1","lab01")),
    ("11:15 – 12:30", "1:15", "Labs 01–03: Loading, Types, Connections and Fill",
     "Lab 01 — load a table and fix every data type. Lab 02 — connect to an "
     "external workbook and load it as Connection Only. Lab 03 — repair a "
     "ragged report with Fill Down and Fill Up.",
     "Hands-on practical · Individual work", srange("lab01","lab03",end_inclusive=True)),
    ("12:30 – 13:30", "—", "LUNCH BREAK", "Not counted as instructional time.", "—", "—"),
    ("13:30 – 13:40", "0:10", "Digital Attendance (PM)",
     "Mandatory SSG digital attendance for the afternoon session.",
     "Administration", "—"),
    ("13:40 – 15:10", "1:30", "Topic 1: Shaping Columns",
     "Splitting columns by delimiter — into columns versus into rows, and why "
     "the choice matters. Conditional, custom and example-driven columns. "
     "Merging columns, sorting and filtering inside the query.",
     "Trainer-led · Demonstration · Q&A", srange("lab03","lab04")),
    ("15:10 – 16:10", "1:00", "Labs 04–06: Splitting, Adding and Filtering",
     "Lab 04 — split a name and a product list. Lab 05 — add conditional, "
     "custom and example-driven columns. Lab 06 — merge columns, then sort "
     "and filter at source.",
     "Hands-on practical · Individual work", srange("lab04","lab06",end_inclusive=True)),
    ("16:10 – 17:30", "1:20", "Topic 2: Logic and Dates in M",
     "The M language versus Excel formulas — case sensitivity and the "
     "compulsory else. Writing IF logic by hand. Nesting conditions with "
     "and / or. Date.AddDays and the deadline pipeline.",
     "Trainer-led · Demonstration · Worked example", srange("topic2","lab07")),
    ("17:30 – 18:30", "1:00", "Labs 07–09: IF, Nested Logic and Deadlines",
     "Lab 07 — write IF logic in M. Lab 08 — nest IF with AND for a priority "
     "flag. Lab 09 — calculate due dates and ageing buckets. Day 1 recap.",
     "Hands-on practical · Individual work", srange("lab07","lab09",end_inclusive=True)),
]

DAY2 = [
    ("9:30 – 9:40", "0:10", "Digital Attendance (AM) + Day 1 Recap",
     "Mandatory SSG digital attendance. Recap of Day 1 and Q&A.",
     "Administration · Class discussion", "—"),
    ("9:40 – 11:00", "1:20", "Topic 2: Reshaping and Aggregating",
     "Why wide is for humans and tall is for machines. The three unpivot "
     "commands and the trap in Unpivot Only Selected Columns. Pivot and the "
     "Don't Aggregate option. Append versus merge. The Group By feature.",
     "Trainer-led · Demonstration · Q&A", srange("lab09","lab10")),
    ("11:00 – 12:00", "1:00", "Labs 10–12: Unpivot, Pivot, Append and Group By",
     "Lab 10 — unpivot a wide monthly report into tidy rows. Lab 11 — pivot "
     "without aggregation and append two quarters. Lab 12 — summarise 400 "
     "transactions with Group By.",
     "Hands-on practical · Individual work", srange("lab10","lab12",end_inclusive=True)),
    ("12:00 – 12:30", "0:30", "Topic 2: The Power Pivot Data Model",
     "The division of labour between Power Query and Power Pivot. Star "
     "schema — fact and lookup tables. Relationships. Implicit versus "
     "explicit measures, and the four DAX measures.",
     "Trainer-led · Demonstration", srange("lab12","lab13")),
    ("12:30 – 13:30", "—", "LUNCH BREAK", "Not counted as instructional time.", "—", "—"),
    ("13:30 – 13:40", "0:10", "Digital Attendance (PM)",
     "Mandatory SSG digital attendance for the afternoon session.",
     "Administration", "—"),
    ("13:40 – 14:20", "0:40", "Lab 13: Build a Power Pivot Model with DAX",
     "Load three tables to the Data Model, draw the relationships in Diagram "
     "View, and write four explicit DAX measures, verifying each in a PivotTable.",
     "Hands-on practical · Individual work", sref("lab13")),
    ("14:20 – 15:40", "1:20", "Topic 3: Joins and Merging",
     "How a merge works and the step everyone forgets. All six join kinds and "
     "the rows each keeps. Choosing a join. Merge versus XLOOKUP. Merging on "
     "multiple fields and the row-count inflation trap.",
     "Trainer-led · Demonstration · Q&A", srange("topic3","lab14")),
    ("15:40 – 16:40", "1:00", "Labs 14–16: The Six Joins and Composite Keys",
     "Lab 14 — run all six joins and record the row counts. Lab 15 — replace "
     "XLOOKUP with a Left Outer merge. Lab 16 — merge on a composite key.",
     "Hands-on practical · Individual work", srange("lab14","lab16",end_inclusive=True)),
    ("16:40 – 17:40", "1:00", "Topic 3 + Labs 17–19: Advanced Matching",
     "Approximate match by binning and by banding table. Cross join and drill "
     "down. Fuzzy matching, the similarity threshold and transformation "
     "tables. Labs 17, 18 and 19.",
     "Trainer-led · Hands-on practical", srange("lab17","lab19",end_inclusive=True)),
    ("17:40 – 18:00", "0:20", "Revision, Feedback and TRAQOM Survey",
     "Consolidation of the complete workflow, the common pitfalls, and the "
     "mandatory TRAQOM course-quality survey.",
     "Class discussion · Administration", srange("revision","end")),
    ("18:00 – 18:30", "0:30", "Digital Attendance (Assessment) + Assessment Briefing",
     "Mandatory assessment digital attendance and the assessment briefing. "
     "The Written Assessment (1 hr) and Practical Performance (1 hr) follow.",
     "Administration · Assessment", srange("revision","end")),
]


def hdr_row(t, labels, widths=None):
    cells = t.add_row().cells
    for i, h in enumerate(labels):
        cells[i].text = ""
        r = cells[i].paragraphs[0].add_run(h)
        r.bold = True; r.font.size = Pt(9); r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        r.font.name = "Arial"
        _shade_cell(cells[i], "1F6FEB")
    return cells


def schedule_table(doc, rows):
    t = doc.add_table(rows=0, cols=6)
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_row(t, ["Time", "Dur.", "Session / Topic", "Content and Activities",
                "Instructional Method", "Slide Ref"])
    for time, dur, sess, cont, meth, sl in rows:
        cells = t.add_row().cells
        is_break = "LUNCH" in sess.upper()
        for i, val in enumerate([time, dur, sess, cont, meth, sl]):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            r = p.add_run(str(val))
            r.font.size = Pt(8.5); r.font.name = "Arial"
            r.bold = (i == 2)
            if is_break:
                r.font.color.rgb = GREY; r.italic = True
        if is_break:
            for c in cells:
                _shade_cell(c, "FFF7E6")
    widths = [Inches(0.95), Inches(0.45), Inches(1.55), Inches(2.75),
              Inches(1.30), Inches(0.75)]
    for row in t.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = w
    return t


def body(doc, text, size=11, bold=False, after=6, color=DARK):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    r = p.add_run(text); r.font.size = Pt(size); r.font.name = "Arial"
    r.bold = bold; r.font.color.rgb = color
    return p


def bullet(doc, text, size=11):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text); r.font.size = Pt(size); r.font.name = "Arial"
    return p


def main():
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = "Arial"; st.font.size = Pt(11)
    style_headings(doc)

    for s in doc.sections:
        s.left_margin = Inches(0.8); s.right_margin = Inches(0.8)
        s.top_margin = Inches(0.7); s.bottom_margin = Inches(0.7)

    logo = os.path.join(ASSETS, "tertiary-infotech-logo.png")
    add_cover_page(doc, "LESSON PLAN", C.TITLE, C.VERSION,
                   org_logo=logo, course_code=C.COURSE_CODE)

    add_version_control(doc, [
        ("v4.0", "16 August 2024", "Previous release under course code TGS-2024048312.",
         "Tertiary Infotech Academy"),
        (C.VERSION, C.VERSION_DATE,
         "Re-issued for course code TGS-2026064177. Content realigned to the "
         "published three-topic outline; out-of-scope statistics material "
         "(hypothesis testing, t-tests, regression) removed. 19 hands-on labs "
         "added with starter workbooks. Slide references refreshed.",
         "Tertiary Infotech Academy"),
    ])

    add_toc(doc)

    # ---------------------------------------------------------- 1. overview
    doc.add_heading("1. Course Overview", level=1)
    t = doc.add_table(rows=0, cols=2); t.style = "Table Grid"
    info = [
        ("Course Title", C.TITLE),
        ("Course Reference Number", C.COURSE_CODE),
        ("Training Provider", f"{C.ORG} ({C.UEN})"),
        ("Course Duration", f"{C.DAYS} days · 16 instructional hours "
                            f"(8 hours per day, excluding lunch)"),
        ("Assessment Duration", "2 hours (Written Assessment 1 hour + "
                                "Practical Performance 1 hour)"),
        ("Mode of Delivery", "Classroom-based, instructor-led with hands-on practical labs"),
        ("Class Size", "Maximum 20 learners"),
        ("TSC Title", C.TSC_TITLE),
        ("TSC Code", C.TSC_CODE),
        ("Software Required", C.SOFTWARE),
    ]
    for k, v in info:
        cells = t.add_row().cells
        cells[0].text = ""; r = cells[0].paragraphs[0].add_run(k)
        r.bold = True; r.font.size = Pt(10); r.font.name = "Arial"
        cells[1].text = ""; r2 = cells[1].paragraphs[0].add_run(v)
        r2.font.size = Pt(10); r2.font.name = "Arial"
        cells[0].width = Inches(1.9); cells[1].width = Inches(5.8)
    doc.add_paragraph()

    # ---------------------------------------------------------- 2. outcomes
    doc.add_heading("2. Learning Outcomes", level=1)
    body(doc, "By the end of this course, learners will be able to:")
    for lo in C.LEARNING_OUTCOMES:
        bullet(doc, lo)
    doc.add_paragraph()

    doc.add_heading("3. Skills Framework Alignment", level=1)
    body(doc, f"TSC Title: {C.TSC_TITLE}    ·    TSC Code: {C.TSC_CODE}", bold=True)
    body(doc, "Knowledge assessed:", bold=True, after=2)
    for k in C.TSC_KNOWLEDGE:
        bullet(doc, k, size=10)
    body(doc, "Abilities assessed:", bold=True, after=2)
    for a in C.TSC_ABILITIES:
        bullet(doc, a, size=10)
    doc.add_paragraph()

    # ---------------------------------------------------------- 4. audience
    doc.add_heading("4. Target Audience and Prerequisites", level=1)
    body(doc, "Target audience", bold=True, after=2)
    body(doc, ", ".join(C.TARGET_AUDIENCE) + ".", size=10)
    body(doc, "Prerequisites", bold=True, after=2)
    for p in C.PREREQUISITES:
        bullet(doc, p, size=10)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ---------------------------------------------------------- 5. schedule
    doc.add_heading("5. Detailed Lesson Plan", level=1)
    body(doc, "Each training day runs 9:30am to 6:30pm and delivers exactly 8 "
              "instructional hours. The 1-hour lunch break is excluded from "
              "instructional time; tea breaks are taken within the sessions shown.",
         size=10, color=GREY)

    for day, rows, theme in ((1, DAY1, C.DAY_THEMES[1]), (2, DAY2, C.DAY_THEMES[2])):
        doc.add_heading(f"Day {day} — {theme}", level=2)
        schedule_table(doc, rows)
        total = sum(
            int(d.split(":")[0]) * 60 + int(d.split(":")[1])
            for _, d, _, _, _, _ in rows if ":" in d)
        body(doc, f"Total instructional time, Day {day}: "
                  f"{total // 60} hours {total % 60:02d} minutes.",
             size=9.5, bold=True, color=BRAND)
        doc.add_paragraph()

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ---------------------------------------------------------- 6. labs
    doc.add_heading("6. Practical Activities (Labs)", level=1)
    body(doc, "The course is built around 19 hands-on labs. Each lab has its own "
              "folder under labs/ containing a starter Excel workbook with mock "
              "data and pre-built Power Query queries. Full step-by-step "
              "instructions for every lab are in the Learner Guide.",
         size=10, color=GREY)
    t = doc.add_table(rows=0, cols=4); t.style = "Table Grid"
    hdr_row(t, ["Lab", "Title", "Maps To", "Starter Workbook"])
    for a in ACTIVITIES:
        cells = t.add_row().cells
        for i, v in enumerate([f"{a['num']:02d}", a["title"],
                               a["objective"].split("—")[0].strip(), a["workbook"]]):
            cells[i].text = ""
            r = cells[i].paragraphs[0].add_run(v)
            r.font.size = Pt(8.5); r.font.name = "Arial"; r.bold = (i == 0)
        cells[0].width = Inches(0.4); cells[1].width = Inches(2.6)
        cells[2].width = Inches(2.0); cells[3].width = Inches(2.7)
    doc.add_paragraph()

    # ---------------------------------------------------------- 7. assessment
    doc.add_heading("7. Assessment", level=1)
    body(doc, C.ASSESSMENT["written"])
    body(doc, C.ASSESSMENT["practical"])
    body(doc, C.ASSESSMENT["note"], bold=True)
    body(doc, "Assessment is open book: learners may refer to the slide deck, the "
              "Learner Guide and any approved course material, but may not use the "
              "internet, AI assistants, or discuss with other learners. Learners "
              "assessed Not Yet Competent are offered one re-assessment opportunity.",
         size=10)
    body(doc, "Criteria for funding", bold=True, after=2)
    bullet(doc, "A minimum attendance rate of 75%, based on the SSG digital "
                "attendance record.", size=10)
    bullet(doc, "Completion of the assessment with an outcome of Competent.", size=10)
    doc.add_paragraph()

    # ---------------------------------------------------------- 8. resources
    doc.add_heading("8. Training Resources", level=1)
    for r in [
        "Trainer slide deck (PPTX) and learner slide deck (PDF)",
        "Learner Guide (PDF) with step-by-step instructions for all 19 labs",
        "This Lesson Plan (PDF)",
        "19 lab folders, each with a starter Excel workbook containing mock data "
        "and pre-built Power Query queries",
        "A laptop per learner with Microsoft Excel for Windows (Microsoft 365, "
        "2021, 2019 or 2016), projector and whiteboard",
    ]:
        bullet(doc, r, size=10)

    add_page_numbers(doc)
    enable_update_fields(doc)

    out = os.path.join(REPO, "courseware", f"LP-{C.SHORT_TITLE}.docx")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    doc.save(out)
    print("saved:", out)


if __name__ == "__main__":
    main()
