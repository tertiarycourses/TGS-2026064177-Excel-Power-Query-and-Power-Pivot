"""
SINGLE SOURCE OF TRUTH — Excel Power Query and Power Pivot (TGS-2026064177).

Every artifact (PPT, LP, LG, LG.md, labs index, assessment) is generated from
this file plus data_domain1.py … data_domain3.py so they stay 100% aligned.

Scope note: the legacy v4 deck (course code TGS-2024048312) carried ~20 slides of
hypothesis testing / t-test / regression content plus a large block of Power BI
service + Power Platform dataflow theory. Neither appears in the published
TGS-2026064177 topic outline, so both are OUT OF SCOPE here. Nothing is taught
or assessed that is not in the outline and covered by a lab.
"""

# ------------------------------------------------------------------ metadata
TITLE        = "Excel Power Query and Power Pivot"
SHORT_TITLE  = "Excel Power Query and Power Pivot"   # used in output filenames
COURSE_CODE  = "TGS-2026064177"
VERSION      = "v5.0"
VERSION_DATE = "12 August 2026"
ORG          = "Tertiary Infotech Academy Pte Ltd"
UEN          = "UEN: 201200696W"
TRAINER      = "Dr Alfred Ang"
DAYS         = 2

WEBSITE = "www.tertiarycourses.com.sg"
EMAIL   = "enquiry@tertiaryinfotech.com"
TEL     = "+65 6100 0613"

# ------------------------------------------------------------------ TSC / Skills Framework
TSC_TITLE = "Data and Statistical Analytics"
TSC_CODE  = "EGS-BIN-4055-1.1"

TSC_KNOWLEDGE = [
    "K1  Operations of data transformation techniques, e.g. shaping, cleansing, merging and aggregation",
    "K2  Conditions required for each transformation and join technique",
    "K3  Interpretation of the results from a query and a data model",
    "K4  Types of data analysis software and their capabilities",
]

TSC_ABILITIES = [
    "A1  Review data sets to uncover trends or patterns",
    "A2  Develop new methods to conduct analysis of large complex data sets specific to each issue",
    "A3  Facilitate the discussion on areas for application of data analytics to examine issues",
]

# ------------------------------------------------------------------ outcomes
LEARNING_OUTCOMES = [
    "LO1: Review data sets using Excel Power Query to uncover trends and patterns.",
    "LO2: Develop methods in Power Pivot to conduct complex analyses tailored to specific datasets.",
    "LO3: Apply advanced Power Query analytics to examine business issues.",
]

# ------------------------------------------------------------------ topics
TOPICS = [
    dict(num=1, code="01",
         title="Review Data with Power Query",
         subtitle="Data types · Getting data · Shaping columns · Sorting and filtering",
         weighting="35%",
         concepts=[
             "Power Query is Excel's ETL engine: it Extracts data from a source, "
             "Transforms it through a recorded list of steps, and Loads the result "
             "to a worksheet or the Data Model — repeatably, on every refresh.",
             "Every transformation is recorded as an Applied Step, so the whole "
             "clean-up is auditable and re-runs on new data with one Refresh click.",
             "Data types are set once in the query, not with worksheet formatting. "
             "A wrong type is the single most common cause of a broken refresh.",
             "Loading as Connection Only keeps a staging query out of the workbook "
             "grid, so only the final shaped table consumes rows.",
             "Column-shaping tools — Fill, Split, Merge, Conditional Column, "
             "Custom Column and Column From Examples — replace fragile worksheet formulas.",
         ]),
    dict(num=2, code="02",
         title="Data Analysis with Power Query and Power Pivot",
         subtitle="Logic · Dates · Pivot and unpivot · Group By · Append · The Data Model",
         weighting="35%",
         concepts=[
             "Conditional logic in Power Query is built with the Conditional Column "
             "dialog, or written directly in M with if / then / else — which is "
             "case-sensitive and lowercase, unlike Excel's IF().",
             "Nesting conditions with and / or lets one column resolve a multi-factor "
             "business rule, such as a priority flag driven by both value and status.",
             "Date.AddDays and the Date functions turn a raw date into a deadline, an "
             "age or an ageing bucket that recalculates on every refresh.",
             "Unpivot turns a wide, human-readable report into the tall, tidy shape "
             "that every PivotTable and data model needs. Pivot reverses it.",
             "Group By aggregates rows to a summary grain — count, sum, average, min, "
             "max — without writing a single worksheet formula.",
             "Append stacks tables with the same shape; the Data Model then relates "
             "the stacked facts to their lookup tables for analysis with DAX measures.",
         ]),
    dict(num=3, code="03",
         title="Advanced Data Analysis using Power Query",
         subtitle="The six joins · Merge · Approximate match · Fuzzy matching · Transformation tables",
         weighting="30%",
         concepts=[
             "Merge joins two queries on one or more matching columns. Power Query "
             "offers six join kinds, and choosing the wrong one silently changes the row count.",
             "Left Outer keeps every row of the first table — the safe default and the "
             "direct equivalent of an XLOOKUP that must not lose records.",
             "Inner keeps only matches; Anti joins keep only non-matches, which makes "
             "them the fastest way to find orphaned or unmatched records.",
             "An approximate match (VLOOKUP's TRUE) is reproduced in Power Query by "
             "binning with a conditional column or by a banding table plus a merge.",
             "Cross Join pairs every row with every row — useful for building a "
             "complete calendar or scenario grid, dangerous on large tables.",
             "Fuzzy matching scores similarity between untidy text values; a "
             "transformation table teaches the engine the pairs it cannot guess.",
         ]),
]

# ------------------------------------------------------------------ day themes (8 instructional hours/day)
DAY_THEMES = {
    1: "Getting, shaping and cleaning data with Power Query",
    2: "Combining, modelling and advanced analysis with Power Query and Power Pivot",
}

# ------------------------------------------------------------------ assessment
ASSESSMENT = dict(
    written="Written Assessment (WA) — Short-Answer Questions (SAQ), 1 hour, open book.",
    practical="Practical Performance (PP) — hands-on Power Query and Power Pivot tasks, 1 hour, open book.",
    note="A minimum of 75% attendance is required to be eligible for assessment and funding.",
)

# ------------------------------------------------------------------ recommended courses
RECOMMENDED = [
    "WSQ - Data Analytics with Excel",
    "WSQ - Statistical Data Analysis with Excel for Beginners",
    "WSQ - Software Automation with Excel VBA Programming",
    "WSQ - Create Interactive Dashboard Using Excel",
    "WSQ - Microsoft Power BI Data Analyst Associate (PL-300)",
]

# ------------------------------------------------------------------ prerequisites / audience
PREREQUISITES = [
    "Able to operate using computer functions",
    "Minimum 3 GCE 'O' Levels Passes including English, or WPL Level 5",
    "Minimum 1 year working experience",
    "Positive learning attitude and an enthusiastic learner",
]

TARGET_AUDIENCE = [
    "Data Analyst", "Business Analyst", "Financial Analyst",
    "Marketing Analyst", "Operations Analyst",
    "Business Intelligence Developer", "Report Developer",
    "Data Consultant", "Database Administrator", "Systems Analyst",
]

SOFTWARE = (
    "Microsoft Excel for Windows (Microsoft 365, 2021, 2019 or 2016). "
    "Power Query is built in under Data > Get & Transform Data. "
    "Power Pivot is enabled from File > Options > Add-ins > COM Add-ins."
)
