# Excel Power Query and Power Pivot — WSQ Courseware

[![Course Code](https://img.shields.io/badge/TGS-2026064177-1F6FEB?style=flat-square)](https://www.tertiarycourses.com.sg/casl-excel-power-query-and-power-pivot.html)
[![Duration](https://img.shields.io/badge/Duration-2%20days%20·%2016%20hours-10B981?style=flat-square)](#lesson-plan)
[![Labs](https://img.shields.io/badge/Labs-19-7C3AED?style=flat-square)](labs/)
[![Version](https://img.shields.io/badge/Courseware-v5.0-F59E0B?style=flat-square)](courseware/)
[![Excel](https://img.shields.io/badge/Microsoft-Excel-217346?style=flat-square&logo=microsoftexcel&logoColor=white)](#software-requirements)

Complete WSQ courseware for **Excel Power Query and Power Pivot**, delivered by
[Tertiary Infotech Academy Pte Ltd](https://www.tertiarycourses.com.sg) (UEN 201200696W).

| Course detail | Information |
|---|---|
| Course code | `TGS-2026064177` |
| Programme | WSQ |
| Duration | 2 days / 16 hours |
| Registration | [View course details and register](https://www.tertiarycourses.com.sg/casl-excel-power-query-and-power-pivot.html) |
| Funding | Up to 70% course-fee funding may be available. Eligibility, validity dates and other terms apply; check the registration page for current details. |

---

## What this course teaches

Turn a messy monthly Excel extract into a clean, refreshable, auditable report —
then model and analyse it. Everything is built by the learner across 19 hands-on labs.

| # | Learning outcome |
|---|---|
| **LO1** | Review data sets using Excel Power Query to uncover trends and patterns |
| **LO2** | Develop methods in Power Pivot to conduct complex analyses tailored to specific datasets |
| **LO3** | Apply advanced Power Query analytics to examine business issues |

**Skills Framework:** Data and Statistical Analytics (`EGS-BIN-4055-1.1`)

---

## Repository contents

```
courseware/
  Excel Power Query and Power Pivot-v5.0.pptx   trainer slide deck (102 slides)
  Excel Power Query and Power Pivot-v5.0.pdf    learner slide deck
  LP-Excel Power Query and Power Pivot.docx     Lesson Plan (+ PDF)
  LG-Excel Power Query and Power Pivot.docx     Learner Guide (+ PDF)
  assets/figures/                               per-lab Power Query Editor figures
labs/
  lab-01-…  …  lab-19-…                         one folder per lab
LG-Excel Power Query and Power Pivot.md         Learner Guide, Markdown mirror
```

Every artifact is generated from a **single source of truth**
(`.claude/skills/courseware-build/build/course_data.py` + `data_domain1-3.py` +
`lab_steps.py`), so the deck, the Lesson Plan, the Learner Guide and the labs can
never drift apart.

---

## Course outline

### Topic 1 — Review Data with Power Query
Data types · querying a table or another workbook · connection-only loads · fill up
and down · splitting by delimiter and into rows · conditional, custom and
example-driven columns · merging columns · sorting and filtering.

### Topic 2 — Data Analysis with Power Query and Power Pivot
`if` logic in M · nesting with `and` / `or` · `Date.AddDays` deadlines · pivot and
unpivot (and the unpivot trap) · Group By · appending queries · the Power Pivot Data
Model, relationships and DAX measures.

### Topic 3 — Advanced Data Analysis using Power Query
All six join kinds · left outer vs XLOOKUP · merging on multiple fields ·
approximate match by binning and banding table · cross join · drill down to a
variable · fuzzy matching and transformation tables.

---

## The labs

Each lab has **its own folder** containing a starter `.xlsx` with realistic mock data
and **Power Query queries already embedded**, plus a `README.md` brief. The full
click-by-click procedure lives in the Learner Guide.

| Lab | Title | Topic |
|---|---|---|
| 01 | [Load Sales Data and Fix Data Types](labs/lab-01-load-sales-data-and-fix-data-types/) | 1 |
| 02 | [Combine an External Workbook as a Connection Only](labs/lab-02-combine-an-external-workbook-as-a-connection-only/) | 1 |
| 03 | [Fill Down and Fill Up a Ragged Report](labs/lab-03-fill-down-and-fill-up-a-ragged-report/) | 1 |
| 04 | [Split Columns by Delimiter and into Rows](labs/lab-04-split-columns-by-delimiter-and-into-rows/) | 1 |
| 05 | [Add Conditional, Custom and Example-Driven Columns](labs/lab-05-add-conditional-custom-and-example-driven-columns/) | 1 |
| 06 | [Merge Columns, Sort and Filter](labs/lab-06-merge-columns-sort-and-filter/) | 1 |
| 07 | [Conditional Logic with IF in M](labs/lab-07-conditional-logic-with-if-in-m/) | 2 |
| 08 | [Nest IF with AND for a Priority Flag](labs/lab-08-nest-if-with-and-for-a-priority-flag/) | 2 |
| 09 | [Calculate Deadlines with Date.AddDays](labs/lab-09-calculate-deadlines-with-date-adddays/) | 2 |
| 10 | [Unpivot a Wide Report into Tidy Rows](labs/lab-10-unpivot-a-wide-report-into-tidy-rows/) | 2 |
| 11 | [Pivot, Pivot Without Aggregation, and Append](labs/lab-11-pivot-pivot-without-aggregation-and-append/) | 2 |
| 12 | [Summarise with Group By](labs/lab-12-summarise-with-group-by/) | 2 |
| 13 | [Build a Power Pivot Data Model with DAX Measures](labs/lab-13-build-a-power-pivot-data-model-with-dax-measures/) | 2 |
| 14 | [Walk Through All Six Join Kinds](labs/lab-14-walk-through-all-six-join-kinds/) | 3 |
| 15 | [Left Outer Join versus XLOOKUP](labs/lab-15-left-outer-join-versus-xlookup/) | 3 |
| 16 | [Merge on Multiple Fields](labs/lab-16-merge-on-multiple-fields/) | 3 |
| 17 | [Approximate Match: Binning and Conditional Columns](labs/lab-17-approximate-match-binning-and-conditional-columns/) | 3 |
| 18 | [Cross Join and Drill Down to a Variable](labs/lab-18-cross-join-and-drill-down-to-a-variable/) | 3 |
| 19 | [Fuzzy Matching and Transformation Tables](labs/lab-19-fuzzy-matching-and-transformation-tables/) | 3 |

The mock data carries **deliberate defects and boundary cases** — text-typed numbers,
`dd/mm/yyyy` dates, orphan product codes, exact band boundaries at 1,000 / 5,000 /
10,000 / 50,000, and inconsistently spelled company names — so each lab's lesson
actually bites.

---

## Software requirements

- **Microsoft Excel for Windows** — Microsoft 365, 2021, 2019 or 2016
- **Power Query** — built in under `Data > Get & Transform Data`
- **Power Pivot** — enable via `File > Options > Add-ins > COM Add-ins`

---

## Rebuilding the courseware

```bash
bash .claude/skills/courseware-build/build/build_courseware.sh
```

This regenerates the lab workbooks (with their Power Query queries), the slide deck,
the per-lab figures, the Lesson Plan and the Learner Guide, renders every PDF and
injects page-numbered tables of contents.

### Quality audit

```bash
# renders every page to an image and checks it against the house standards
/courseware-qa
```

The courseware is audited against the published Tertiary Infotech WSQ standards at
<https://tertiarycourses.github.io/wsqcourseware/> — the PPT 7-point and assessment
checklists, Lesson Plan slide-reference accuracy, the Learner Guide and its Markdown
mirror, and 100% deck/LG/lab alignment. Run it after every regeneration; the deck,
LP and LG must pass before anything is published.

---

## Assessment

Assessed by a **Written Assessment (SAQ, 1 hour)** covering K1–K4 and a **Practical
Performance (1 hour)** covering A1–A3. Both are open book.

> **Note:** assessment papers and answer keys are **confidential** and are
> deliberately **not** published in this repository. They are distributed through
> Google Drive and the LMS to authorised trainers and assessors only.

---

## Funding and certification

A minimum **75% attendance** (SSG digital attendance record) and an outcome of
**Competent** in both instruments are required for funding and for the Certificate
of Achievement.

---

## Support

**Tertiary Infotech Academy Pte Ltd** · UEN 201200696W
📧 <enquiry@tertiaryinfotech.com> · ☎ +65 6100 0613 · 🌐 [tertiarycourses.com.sg](https://www.tertiarycourses.com.sg)

---

© 2026 Tertiary Infotech Academy Pte Ltd. All rights reserved.
