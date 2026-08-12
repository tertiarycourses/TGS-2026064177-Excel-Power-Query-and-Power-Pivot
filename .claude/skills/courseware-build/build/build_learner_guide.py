#!/usr/bin/env python3
"""
Learner Guide (LG) — Excel Power Query and Power Pivot (TGS-2026064177).

Produces BOTH from one source so they can never diverge:
  courseware/LG-<TITLE>.docx    house-format Word document
  LG-<TITLE>.md                 Markdown mirror at the repo root

The LG carries the DETAILED, click-by-click Power Query and Power Pivot
procedure for every lab — the ribbon path, the dialog, the exact M or DAX to
type, what you should see, and how to fix it when it goes wrong. The slide deck
deliberately carries none of this.
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
from lab_steps import LAB_STEPS
from prodoc import (add_cover_page, add_version_control, add_toc, add_page_numbers,
                    style_headings, enable_update_fields, _shade_cell,
                    BRAND, DARK, GREY)

ACTIVITIES = DOMAIN1 + DOMAIN2 + DOMAIN3
MONO = "Consolas"


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


def slug(title):
    out = "".join(c.lower() if c.isalnum() else "-" for c in title)
    while "--" in out:
        out = out.replace("--", "-")
    return out.strip("-")


# ------------------------------------------------------------------ DOCX bits
def body(doc, text, size=11, bold=False, after=6, color=DARK, italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    r = p.add_run(text); r.font.size = Pt(size); r.font.name = "Arial"
    r.bold = bold; r.italic = italic; r.font.color.rgb = color
    return p


def bullet(doc, text, size=10.5):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text); r.font.size = Pt(size); r.font.name = "Arial"
    return p


def numbered(doc, text, size=11, n=None):
    """A step. Word's List Number style numbers continuously across the whole
    document, so each lab's steps would read 41, 42, 43… Pass n to write the
    step number literally and restart at 1 for every lab."""
    if n is None:
        p = doc.add_paragraph(style="List Number")
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(text); r.font.size = Pt(size); r.font.name = "Arial"
        return p
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.left_indent = Inches(0.32)
    p.paragraph_format.first_line_indent = Inches(-0.32)
    rn = p.add_run(f"{n}.  ")
    rn.bold = True; rn.font.size = Pt(size); rn.font.name = "Arial"
    rn.font.color.rgb = BRAND
    r = p.add_run(text); r.font.size = Pt(size); r.font.name = "Arial"
    return p


def codeblock(doc, lines, lang=""):
    """A shaded monospace block for M / DAX code."""
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    cell = t.rows[0].cells[0]
    _shade_cell(cell, "F5F8FC")
    cell.text = ""
    if lang:
        p0 = cell.paragraphs[0]
        p0.paragraph_format.space_after = Pt(2)
        r0 = p0.add_run(lang)
        r0.font.size = Pt(7.5); r0.font.name = "Arial"; r0.bold = True
        r0.font.color.rgb = BRAND
        first = False
    else:
        first = True
    for i, ln in enumerate(lines):
        p = cell.paragraphs[0] if (first and i == 0) else cell.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(ln if ln else " ")
        r.font.size = Pt(9.5); r.font.name = MONO
        r.font.color.rgb = RGBColor(0x11, 0x18, 0x27)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return t


def callout(doc, label, text, fill="FFF7E6"):
    t = doc.add_table(rows=1, cols=1); t.style = "Table Grid"
    cell = t.rows[0].cells[0]; _shade_cell(cell, fill); cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(label + "  ")
    r.bold = True; r.font.size = Pt(9.5); r.font.name = "Arial"; r.font.color.rgb = BRAND
    r2 = p.add_run(text); r2.font.size = Pt(9.5); r2.font.name = "Arial"
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return t


def main():
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = "Arial"; st.font.size = Pt(11)
    style_headings(doc)
    for s in doc.sections:
        s.left_margin = Inches(0.85); s.right_margin = Inches(0.85)
        s.top_margin = Inches(0.7); s.bottom_margin = Inches(0.7)

    logo = os.path.join(ASSETS, "tertiary-infotech-logo.png")
    add_cover_page(doc, "Learner Guide", C.TITLE, C.VERSION,
                   org_logo=logo, course_code=C.COURSE_CODE)

    add_version_control(doc, [
        ("v4.0", "16 August 2024", "Previous release under course code TGS-2024048312.",
         "Tertiary Infotech Academy"),
        (C.VERSION, C.VERSION_DATE,
         "Re-issued for course code TGS-2026064177. Content realigned to the "
         "published three-topic outline; out-of-scope statistics material removed. "
         "Full step-by-step Power Query and Power Pivot procedures added for all "
         "19 labs, each with its own starter workbook.",
         "Tertiary Infotech Academy"),
    ])

    add_toc(doc)

    # ---------------------------------------------------------- introduction
    doc.add_heading("1. About This Course", level=1)
    body(doc, f"{C.TITLE} ({C.COURSE_CODE}) is a {C.DAYS}-day, 16-hour WSQ course "
              "that teaches you to turn messy Excel data into clean, refreshable "
              "reports using Power Query, and to analyse it with the Power Pivot "
              "Data Model and DAX.")
    body(doc, "Learning outcomes", bold=True, after=2)
    for lo in C.LEARNING_OUTCOMES:
        bullet(doc, lo)
    body(doc, "How to use this guide", bold=True, after=2)
    body(doc, "Section 3 explains the concepts you need. Sections 4 to 6 contain "
              "the 19 labs. Every lab gives you the exact ribbon path, the dialog "
              "settings, the M or DAX code to type, what you should see on screen, "
              "and what to do when it goes wrong. Work through them in order — "
              "later labs assume the skills built in earlier ones.", size=10.5)

    doc.add_heading("2. Before You Start", level=1)
    body(doc, "Software", bold=True, after=2)
    body(doc, C.SOFTWARE, size=10.5)
    body(doc, "Enabling Power Pivot (needed from Lab 13 onward)", bold=True, after=2)
    for i, s in enumerate([
        "Click File > Options.",
        "Select Add-ins in the left pane.",
        "At the bottom, set Manage to COM Add-ins and click Go.",
        "Tick Microsoft Power Pivot for Excel and click OK.",
        "The Power Pivot tab now appears on the ribbon.",
    ], 1):
        numbered(doc, s, size=10.5, n=i)
    callout(doc, "Note:",
            "Power Query is built into Excel 2016 and later under Data > Get & "
            "Transform Data — no add-in is required. Power Pivot is available in "
            "Microsoft 365, Office Professional Plus, and the standalone Excel "
            "2016/2019/2021 editions.")

    body(doc, "Your lab files", bold=True, after=2)
    body(doc, "Download the labs folder from the LMS/TMS portal at "
              "lms-tms.tertiaryinfotech.com and unzip it to your Desktop. Each lab "
              "has its own folder containing a starter workbook with mock data and, "
              "where appropriate, Power Query queries already created for you.",
         size=10.5)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ---------------------------------------------------------- concepts
    doc.add_heading("3. Key Concepts", level=1)

    doc.add_heading("3.1 What Power Query Does", level=2)
    body(doc, "Power Query is Excel's ETL engine. It Extracts data from a source, "
              "Transforms it through a recorded list of steps, and Loads the result "
              "to a worksheet or the Data Model. The crucial word is recorded: every "
              "click becomes an Applied Step, so next month you replace the source "
              "file, press Refresh, and the entire clean-up runs again in the same "
              "order.")
    callout(doc, "Rule of thumb:",
            "If you have done a clean-up by hand twice, it belongs in Power Query.")

    doc.add_heading("3.2 Data Types", level=2)
    body(doc, "A wrong data type is the most common cause of a broken refresh. Set "
              "the type of every column explicitly, as the first thing you do after "
              "loading.")
    t = doc.add_table(rows=0, cols=3); t.style = "Table Grid"
    hdr = t.add_row().cells
    for i, h in enumerate(["Type", "What it holds", "Common failure"]):
        hdr[i].text = ""
        r = hdr[i].paragraphs[0].add_run(h); r.bold = True; r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r.font.name = "Arial"
        _shade_cell(hdr[i], "1F6FEB")
    for row in [
        ("Text", "Names, codes, IDs, postal codes", "Leading zeros lost if typed as a number"),
        ("Whole Number", "Counts and quantities", "Will not sum while stored as text"),
        ("Decimal Number", "Prices, amounts and rates", "A currency prefix blocks the type change"),
        ("Date", "Calendar dates", "dd/mm/yyyy read as US format — use Using Locale"),
        ("True/False", "Logical flags", "The words TRUE/FALSE stored as text"),
    ]:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            r = cells[i].paragraphs[0].add_run(v)
            r.font.size = Pt(9.5); r.font.name = "Arial"; r.bold = (i == 0)
    doc.add_paragraph()

    doc.add_heading("3.3 The M Language", level=2)
    body(doc, "Power Query's formula language is called M. It is not Excel's formula "
              "language, and the differences catch everyone out at first:")
    for b in [
        "M is CASE SENSITIVE. The keywords if, then, else, and, or, each are always "
        "lowercase. Writing If or IF produces an error.",
        "Columns are referenced by name in square brackets: [Amount], not A2.",
        "Every if MUST have an else. There is no optional third argument.",
        "Text is joined with & — but both sides must be text. Use Text.From([Number]) "
        "to convert a number first.",
    ]:
        bullet(doc, b)
    codeblock(doc, ['if [Amount] >= 500 then 0 else 15'], lang="M")

    doc.add_heading("3.4 Power Query versus Power Pivot", level=2)
    body(doc, "Power Query shapes each table. Power Pivot relates those tables to "
              "one another and calculates over them with DAX measures. Do row-level "
              "clean-up in Power Query; do aggregation and time intelligence in DAX. "
              "Working against that division fights both tools.")

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ---------------------------------------------------------- labs
    sec_no = {1: "4", 2: "5", 3: "6"}
    for topic in C.TOPICS:
        n = topic["num"]
        doc.add_heading(f"{sec_no[n]}. Topic {n} — {topic['title']}", level=1)
        body(doc, topic["subtitle"], italic=True, color=GREY, size=10.5)
        for con in topic["concepts"]:
            bullet(doc, con)
        doc.add_paragraph()

        for lab in [a for a in ACTIVITIES if a["topic"] == n]:
            idx = lab["num"]
            doc.add_heading(f"{sec_no[n]}.{idx} Lab {idx:02d}: {lab['title']}", level=2)

            # meta table
            mt = doc.add_table(rows=0, cols=2); mt.style = "Table Grid"
            for k, v in [("Maps to", lab["objective"]),
                         ("Starter workbook", f"labs/lab-{idx:02d}-{slug(lab['title'])}/{lab['workbook']}"),
                         ("Tools", lab["services"]),
                         ("You will produce", lab["build"])]:
                cells = mt.add_row().cells
                cells[0].text = ""
                r = cells[0].paragraphs[0].add_run(k)
                r.bold = True; r.font.size = Pt(9.5); r.font.name = "Arial"
                cells[1].text = ""
                r2 = cells[1].paragraphs[0].add_run(v)
                r2.font.size = Pt(9.5); r2.font.name = "Arial"
                cells[0].width = Inches(1.35); cells[1].width = Inches(5.6)
            doc.add_paragraph().paragraph_format.space_after = Pt(4)

            body(doc, "Scenario", bold=True, after=2)
            body(doc, lab["desc"], size=10.5)

            # the lab's Power Query Editor figure — queries, preview and the
            # Applied Steps this lab produces
            fig = os.path.join(REPO, "courseware", "assets", "figures",
                               f"lab-{idx:02d}.png")
            if os.path.exists(fig):
                p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(2)
                p.add_run().add_picture(fig, width=Inches(6.4))
                cap = doc.add_paragraph(); cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cap.paragraph_format.space_after = Pt(8)
                rc = cap.add_run(f"Figure {idx}: the Power Query Editor for Lab "
                                 f"{idx:02d} — queries, data preview and Applied Steps.")
                rc.font.size = Pt(8.5); rc.font.name = "Arial"
                rc.italic = True; rc.font.color.rgb = GREY

            body(doc, "Step-by-step", bold=True, after=2)
            steps = LAB_STEPS.get(idx, [])
            if not steps:
                # fall back to the summary steps if a detailed script is missing
                for i, (text, cmd) in enumerate(lab["steps"], 1):
                    numbered(doc, text, n=i)
                    if cmd:
                        codeblock(doc, cmd.split("  |  "), lang="M")
            else:
                for i, step in enumerate(steps, 1):
                    numbered(doc, step["do"], n=i)
                    if step.get("code"):
                        codeblock(doc, step["code"], lang=step.get("lang", "M"))
                    if step.get("see"):
                        callout(doc, "You should see:", step["see"], fill="EAF7F0")

            body(doc, "Test it", bold=True, after=2)
            callout(doc, "Success criterion:", lab["test"], fill="EAF7F0")

            tr = LAB_STEPS.get(f"{idx}_trouble")
            if tr:
                body(doc, "Troubleshooting", bold=True, after=2)
                tt = doc.add_table(rows=0, cols=2); tt.style = "Table Grid"
                h = tt.add_row().cells
                for i, x in enumerate(["Symptom", "Fix"]):
                    h[i].text = ""
                    r = h[i].paragraphs[0].add_run(x); r.bold = True
                    r.font.size = Pt(9.5); r.font.name = "Arial"
                    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    _shade_cell(h[i], "1F6FEB")
                for sym, fix in tr:
                    cells = tt.add_row().cells
                    for i, v in enumerate([sym, fix]):
                        cells[i].text = ""
                        r = cells[i].paragraphs[0].add_run(v)
                        r.font.size = Pt(9.5); r.font.name = "Arial"
                    cells[0].width = Inches(2.5); cells[1].width = Inches(4.45)
                doc.add_paragraph()
            doc.add_paragraph()

    # ---------------------------------------------------------- reference
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    doc.add_heading("7. Quick Reference", level=1)

    doc.add_heading("7.1 The Six Join Kinds", level=2)
    t = doc.add_table(rows=0, cols=3); t.style = "Table Grid"
    h = t.add_row().cells
    for i, x in enumerate(["Join kind", "Which rows you keep", "Typical use"]):
        h[i].text = ""
        r = h[i].paragraphs[0].add_run(x); r.bold = True; r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r.font.name = "Arial"
        _shade_cell(h[i], "1F6FEB")
    for row in [
        ("Left Outer", "All left rows + matches from the right", "The safe default; XLOOKUP replacement"),
        ("Right Outer", "All right rows + matches from the left", "Rare — usually swap the tables"),
        ("Full Outer", "Every row from both tables", "Reconciling two systems"),
        ("Inner", "Only rows matching on both sides", "Confirmed, complete records only"),
        ("Left Anti", "Left rows with NO match on the right", "Find orphans"),
        ("Right Anti", "Right rows with NO match on the left", "Find unused lookups"),
    ]:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            r = cells[i].paragraphs[0].add_run(v)
            r.font.size = Pt(9.5); r.font.name = "Arial"; r.bold = (i == 0)
    doc.add_paragraph()

    doc.add_heading("7.2 Common M Expressions", level=2)
    codeblock(doc, [
        '// conditional',
        'if [Amount] >= 500 then 0 else 15',
        '',
        '// nested with and / or',
        'if [Amount] >= 5000 and [Days] > 30 then "Escalate"',
        'else if [Amount] >= 5000 or [Days] > 30 then "Watch"',
        'else "Normal"',
        '',
        '// dates',
        'Date.AddDays([Invoice Date], [Terms Days])',
        'Duration.Days(#date(2026,6,30) - [Due Date])',
        '',
        '// text',
        'Text.From([ProductID])',
        'Text.Trim([Product])',
    ], lang="M")

    doc.add_heading("7.3 The DAX Measures from Lab 13", level=2)
    codeblock(doc, [
        'Total Sales := SUM(Sales[Amount])',
        'Order Count := COUNTROWS(Sales)',
        'Average Order := DIVIDE([Total Sales], [Order Count])',
        'Sales YTD := TOTALYTD([Total Sales], Calendar[Date])',
    ], lang="DAX")

    doc.add_heading("7.4 Common Errors and What They Mean", level=2)
    t = doc.add_table(rows=0, cols=2); t.style = "Table Grid"
    h = t.add_row().cells
    for i, x in enumerate(["Message / symptom", "What it means and how to fix it"]):
        h[i].text = ""
        r = h[i].paragraphs[0].add_run(x); r.bold = True; r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r.font.name = "Arial"
        _shade_cell(h[i], "1F6FEB")
    for sym, fix in [
        ("Token Eof expected", "A syntax error in M — most often a capitalised If/Then/Else, "
                               "or a missing else clause."),
        ("Expression.Error: The column X of the table wasn't found",
         "A column was renamed or removed by an earlier step. Check the step before it."),
        ("DataFormat.Error: We couldn't convert to Number",
         "The column still holds text such as a currency prefix. Replace the prefix, then set the type."),
        ("The name 'X' wasn't recognised",
         "A misspelled function or column. M is case sensitive — Date.AddDays, not date.adddays."),
        ("Row count jumped after a merge",
         "The matching key is not unique. Add the missing column to the match, or de-duplicate first."),
        ("A column of nulls appeared after Append",
         "Column names differ between the two tables. Rename them so they match exactly."),
        ("#N/A or blank after Close & Load",
         "The query loaded but returned no rows — check the filters in Applied Steps."),
    ]:
        cells = t.add_row().cells
        for i, v in enumerate([sym, fix]):
            cells[i].text = ""
            r = cells[i].paragraphs[0].add_run(v)
            r.font.size = Pt(9.5); r.font.name = "Arial"; r.bold = (i == 0)
        cells[0].width = Inches(2.4); cells[1].width = Inches(4.55)
    doc.add_paragraph()

    doc.add_heading("8. Assessment", level=1)
    body(doc, C.ASSESSMENT["written"])
    body(doc, C.ASSESSMENT["practical"])
    body(doc, C.ASSESSMENT["note"], bold=True)
    body(doc, "The assessment is open book. You may refer to the slide deck, this "
              "Learner Guide and any approved course material, but not the internet, "
              "AI assistants, or other learners.", size=10.5)

    add_page_numbers(doc)
    enable_update_fields(doc)

    out = os.path.join(REPO, "courseware", f"LG-{C.SHORT_TITLE}.docx")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    doc.save(out)
    print("saved:", out)

    write_markdown()


# ------------------------------------------------------------------ Markdown mirror
def write_markdown():
    L = []
    L.append(f"# Learner Guide — {C.TITLE}")
    L.append("")
    L.append(f"**Course Code:** {C.COURSE_CODE}  ")
    L.append(f"**Version:** {C.VERSION} · {C.VERSION_DATE}  ")
    L.append(f"**{C.ORG}** · {C.UEN}")
    L.append("")
    L.append("## 1. About This Course")
    L.append("")
    L.append(f"{C.TITLE} ({C.COURSE_CODE}) is a {C.DAYS}-day, 16-hour WSQ course that "
             "teaches you to turn messy Excel data into clean, refreshable reports "
             "using Power Query, and to analyse it with the Power Pivot Data Model "
             "and DAX.")
    L.append("")
    L.append("**Learning outcomes**")
    L.append("")
    for lo in C.LEARNING_OUTCOMES:
        L.append(f"- {lo}")
    L.append("")
    L.append("## 2. Before You Start")
    L.append("")
    L.append(C.SOFTWARE)
    L.append("")
    L.append("**Enabling Power Pivot** (needed from Lab 13): File > Options > "
             "Add-ins > set Manage to COM Add-ins > Go > tick "
             "*Microsoft Power Pivot for Excel* > OK.")
    L.append("")
    L.append("## 3. Key Concepts")
    L.append("")
    for topic in C.TOPICS:
        L.append(f"### Topic {topic['num']} — {topic['title']}")
        L.append("")
        for con in topic["concepts"]:
            L.append(f"- {con}")
        L.append("")

    sec_no = {1: "4", 2: "5", 3: "6"}
    for topic in C.TOPICS:
        n = topic["num"]
        L.append(f"## {sec_no[n]}. Topic {n} — {topic['title']}")
        L.append("")
        for lab in [a for a in ACTIVITIES if a["topic"] == n]:
            idx = lab["num"]
            folder = f"labs/lab-{idx:02d}-{slug(lab['title'])}"
            L.append(f"### {sec_no[n]}.{idx} Lab {idx:02d}: {lab['title']}")
            L.append("")
            L.append(f"| | |")
            L.append(f"|---|---|")
            L.append(f"| **Maps to** | {lab['objective']} |")
            L.append(f"| **Starter workbook** | [`{lab['workbook']}`]({folder}/{lab['workbook']}) |")
            L.append(f"| **Tools** | {lab['services']} |")
            L.append(f"| **You will produce** | {lab['build']} |")
            L.append("")
            L.append("**Scenario**")
            L.append("")
            L.append(lab["desc"])
            L.append("")
            L.append(f"![Power Query Editor for Lab {idx:02d}]"
                     f"(courseware/assets/figures/lab-{idx:02d}.png)")
            L.append("")
            L.append(f"*Figure {idx}: the Power Query Editor for Lab {idx:02d} — "
                     f"queries, data preview and Applied Steps.*")
            L.append("")
            L.append("**Step-by-step**")
            L.append("")
            steps = LAB_STEPS.get(idx, [])
            if not steps:
                for i, (text, cmd) in enumerate(lab["steps"], 1):
                    L.append(f"{i}. {text}")
                    if cmd:
                        L.append("")
                        L.append("   ```m")
                        for part in cmd.split("  |  "):
                            L.append(f"   {part.strip()}")
                        L.append("   ```")
                    L.append("")
            else:
                for i, step in enumerate(steps, 1):
                    L.append(f"{i}. {step['do']}")
                    if step.get("code"):
                        L.append("")
                        lang = "dax" if step.get("lang") == "DAX" else "m"
                        L.append(f"   ```{lang}")
                        for c in step["code"]:
                            L.append(f"   {c}")
                        L.append("   ```")
                    if step.get("see"):
                        L.append("")
                        L.append(f"   > **You should see:** {step['see']}")
                    L.append("")
            L.append(f"**Test it:** {lab['test']}")
            L.append("")
            tr = LAB_STEPS.get(f"{idx}_trouble")
            if tr:
                L.append("**Troubleshooting**")
                L.append("")
                L.append("| Symptom | Fix |")
                L.append("|---|---|")
                for sym, fix in tr:
                    L.append(f"| {sym} | {fix} |")
                L.append("")

    L.append("## 7. Quick Reference")
    L.append("")
    L.append("### 7.1 The Six Join Kinds")
    L.append("")
    L.append("| Join kind | Which rows you keep | Typical use |")
    L.append("|---|---|---|")
    for row in [
        ("Left Outer", "All left rows + matches from the right", "The safe default; XLOOKUP replacement"),
        ("Right Outer", "All right rows + matches from the left", "Rare — usually swap the tables"),
        ("Full Outer", "Every row from both tables", "Reconciling two systems"),
        ("Inner", "Only rows matching on both sides", "Confirmed, complete records only"),
        ("Left Anti", "Left rows with NO match on the right", "Find orphans"),
        ("Right Anti", "Right rows with NO match on the left", "Find unused lookups"),
    ]:
        L.append(f"| **{row[0]}** | {row[1]} | {row[2]} |")
    L.append("")
    L.append("### 7.2 The DAX Measures from Lab 13")
    L.append("")
    L.append("```dax")
    L.append("Total Sales := SUM(Sales[Amount])")
    L.append("Order Count := COUNTROWS(Sales)")
    L.append("Average Order := DIVIDE([Total Sales], [Order Count])")
    L.append("Sales YTD := TOTALYTD([Total Sales], Calendar[Date])")
    L.append("```")
    L.append("")
    L.append("## 8. Assessment")
    L.append("")
    L.append(f"- {C.ASSESSMENT['written']}")
    L.append(f"- {C.ASSESSMENT['practical']}")
    L.append(f"- {C.ASSESSMENT['note']}")
    L.append("")
    L.append("---")
    L.append("")
    L.append(f"*© 2026 {C.ORG}. All rights reserved. · {C.WEBSITE}*")

    out = os.path.join(REPO, f"LG-{C.SHORT_TITLE}.md")
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    print("saved:", out)


if __name__ == "__main__":
    main()
