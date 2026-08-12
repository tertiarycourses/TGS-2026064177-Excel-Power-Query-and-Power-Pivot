"""
Topic 2 — Data Analysis with Power Query and Power Pivot.

Covers: IF formulas in M, nested IF with AND, Date.AddDays for deadlines,
pivot, pivot-and-append, pivot without aggregation, unpivot, unpivot warnings,
Group By, appending queries, and loading to the Power Pivot Data Model with
relationships and DAX measures.
"""

DOMAIN2 = [
    dict(
        num=7, topic=2,
        title="Conditional Logic with IF in M",
        objective="LO2 · A2 · K1, K2 — Write IF logic in the Power Query M language.",
        desc=(
            "The Conditional Column dialog cannot express every rule, so the learner moves to the "
            "Custom Column window and writes M directly. The lab contrasts Excel's IF() with M's "
            "lowercase, case-sensitive if / then / else and shows why a missing else is an error."
        ),
        build="A Shipping Cost column driven by a hand-written M if expression.",
        services="Excel, Power Query Editor, Custom Column, M language",
        workbook="lab-07-if-formulas.xlsx",
        sheets=["Orders"],
        steps=[
            ("Load the Orders range. Add Column > Custom Column and name it Shipping Cost.", ""),
            ("Enter the M expression, noting that if, then and else must be lowercase.", "if [Amount] >= 500 then 0 else 15"),
            ("Delete the else clause and read the error Power Query raises, then restore it. Every if needs an else.", ""),
            ("Add a second Custom Column named Free Shipping using a comparison that returns a logical value.", "[Amount] >= 500"),
            ("Set Shipping Cost to Decimal Number and Free Shipping to True/False.", ""),
        ],
        test="Orders of 500 or more show Shipping Cost 0 and Free Shipping TRUE; the boundary row at exactly 500 qualifies.",
    ),
    dict(
        num=8, topic=2,
        title="Nest IF with AND for a Priority Flag",
        objective="LO2 · A2 · K2 — Nest conditions with and / or to resolve a multi-factor rule.",
        desc=(
            "An order is escalated only when it is both high value and already overdue. The learner "
            "nests conditions using M's and / or operators to produce a single Priority column, then "
            "verifies each branch against a worked truth table."
        ),
        build="A Priority column classifying every order as Escalate, Watch or Normal.",
        services="Excel, Power Query Editor, Custom Column, M logical operators",
        workbook="lab-08-nested-if-and.xlsx",
        sheets=["Orders"],
        steps=[
            ("Load the Orders range, which carries Amount, Status and Days Outstanding.", ""),
            ("Add Column > Custom Column named Priority with a nested M expression.",
             'if [Amount] >= 5000 and [Days Outstanding] > 30 then "Escalate" else if [Amount] >= 5000 or [Days Outstanding] > 30 then "Watch" else "Normal"'),
            ("Confirm that and requires both conditions while or requires either, using the four-row truth table in the guide.", ""),
            ("Add a Custom Column named Review Needed combining a text test with a numeric test.",
             'if [Status] <> "Closed" and [Days Outstanding] > 60 then "Yes" else "No"'),
            ("Filter Priority to Escalate and record how many orders qualify.", ""),
        ],
        test="Only orders meeting BOTH the value and ageing tests read Escalate; orders meeting one read Watch.",
    ),
    dict(
        num=9, topic=2,
        title="Calculate Deadlines with Date.AddDays",
        objective="LO2 · A2 · K1, K3 — Use Date functions to derive deadlines and ageing buckets.",
        desc=(
            "Each invoice carries payment terms in days. The learner derives the due date with "
            "Date.AddDays, compares it against a fixed reporting date to produce days overdue, and "
            "bins the result into ageing buckets the finance team recognises."
        ),
        build="An invoice ageing query with Due Date, Days Overdue and Ageing Bucket.",
        services="Excel, Power Query Editor, Date functions, Custom Column",
        workbook="lab-09-adddays-deadline.xlsx",
        sheets=["Invoices"],
        steps=[
            ("Load the Invoices range and confirm Invoice Date is typed as Date.", ""),
            ("Add Column > Custom Column named Due Date using the invoice date plus its terms.", "Date.AddDays([Invoice Date], [Terms Days])"),
            ("Set Due Date to Date, then add a Custom Column named Days Overdue against the 30 June 2026 reporting date.",
             "Duration.Days(#date(2026,6,30) - [Due Date])"),
            ("Add a Conditional Column named Ageing Bucket: <= 0 Current, <= 30 1-30 Days, <= 60 31-60 Days, else 60+ Days.", ""),
            ("Sort by Days Overdue descending to bring the worst offenders to the top.", ""),
        ],
        test="An invoice dated 1 May 2026 with 30-day terms shows Due Date 31 May 2026 and Days Overdue 30.",
    ),
    dict(
        num=10, topic=2,
        title="Unpivot a Wide Report into Tidy Rows",
        objective="LO2 · A2 · K2, K3 — Unpivot columns, and choose between Unpivot Other Columns and Unpivot Only Selected.",
        desc=(
            "The monthly sales report is laid out with one column per month — readable for a human, "
            "useless for a PivotTable. The learner unpivots the month columns into Attribute/Value "
            "pairs and learns why Unpivot Other Columns survives a new month being added while "
            "Unpivot Only Selected Columns silently misses it."
        ),
        build="A tidy long-format sales table with Month and Sales columns.",
        services="Excel, Power Query Editor, Unpivot",
        workbook="lab-10-unpivot.xlsx",
        sheets=["MonthlySales"],
        steps=[
            ("Load the MonthlySales range, which has Region plus twelve month columns.", ""),
            ("Select the Region column, then Transform > Unpivot Other Columns.", ""),
            ("Rename Attribute to Month and Value to Sales, then set Sales to Decimal Number.", ""),
            ("Compare with Unpivot Only Selected Columns: the guide explains that adding a 13th month breaks the selected-columns version but not Unpivot Other Columns.", ""),
            ("Close & Load and build a quick PivotTable to prove the tidy shape now pivots correctly.", ""),
        ],
        test="The result has exactly three columns (Region, Month, Sales) and a row count of regions x months.",
    ),
    dict(
        num=11, topic=2,
        title="Pivot, Pivot Without Aggregation, and Append",
        objective="LO2 · A2 · K2, K3 — Pivot a column, use Don't Aggregate, and append queries.",
        desc=(
            "The learner reverses an unpivot to rebuild a cross-tab, meets the Don't Aggregate option "
            "that preserves individual values instead of summing them, and then appends the Q1 and Q2 "
            "extracts into a single continuous table."
        ),
        build="A pivoted cross-tab plus an appended half-year table.",
        services="Excel, Power Query Editor, Pivot Column, Append Queries",
        workbook="lab-11-pivot-append.xlsx",
        sheets=["Q1Sales", "Q2Sales", "Attributes"],
        steps=[
            ("Load Q1Sales and Q2Sales as separate queries, each as Connection Only.", ""),
            ("On the Attributes query select the Attribute column, then Transform > Pivot Column with Value as the values column.", ""),
            ("In Advanced options choose Don't Aggregate so text values are preserved rather than counted.", ""),
            ("Home > Append Queries as New, add Q1Sales and Q2Sales, and name the result HalfYear.", ""),
            ("Confirm both quarters' column names match exactly — a mismatched name creates a new column full of nulls.", ""),
        ],
        test="HalfYear's row count equals Q1 rows plus Q2 rows, with no column of all nulls.",
    ),
    dict(
        num=12, topic=2,
        title="Summarise with Group By",
        objective="LO2 · A2 · K1, K3 — Aggregate rows to a summary grain with Group By.",
        desc=(
            "Rather than loading 20,000 transaction rows into the grid, the learner aggregates them "
            "inside the query. The lab covers basic and advanced Group By, multiple aggregations, "
            "grouping on more than one column, and keeping the underlying rows with All Rows."
        ),
        build="A regional sales summary with count, sum and average in one query.",
        services="Excel, Power Query Editor, Group By",
        workbook="lab-12-group-by.xlsx",
        sheets=["Transactions"],
        steps=[
            ("Load the Transactions range into the Power Query Editor.", ""),
            ("Transform > Group By, switch to Advanced, and group by Region.", ""),
            ("Add three aggregations: Order Count (Count Rows), Total Sales (Sum of Amount), Average Order (Average of Amount).", ""),
            ("Add Product Category as a second Group By column and observe the finer grain.", ""),
            ("Add an All Rows aggregation to keep the detail rows nested in a table column for drill-through.", ""),
        ],
        test="Total Sales across all groups equals the grand total of the source Amount column.",
    ),
    dict(
        num=13, topic=2,
        title="Build a Power Pivot Data Model with DAX Measures",
        objective="LO2 · A2 · K3, K4 — Load to the Data Model, create relationships and write DAX measures.",
        desc=(
            "The learner loads the fact and lookup tables to the Power Pivot Data Model, draws "
            "relationships in Diagram View, and writes explicit DAX measures. The lab contrasts an "
            "implicit measure (dragging a field into Values) with an explicit one, and explains why "
            "explicit measures are the professional practice."
        ),
        build="A Power Pivot data model with three related tables and four DAX measures.",
        services="Excel, Power Query Editor, Power Pivot, DAX, Diagram View",
        workbook="lab-13-power-pivot-model.xlsx",
        sheets=["Sales", "Products", "Calendar"],
        steps=[
            ("Enable Power Pivot via File > Options > Add-ins > COM Add-ins if the tab is not visible.", ""),
            ("Load Sales, Products and Calendar with Close & Load To… > Only Create Connection, ticking 'Add this data to the Data Model'.", ""),
            ("Open Power Pivot > Manage > Diagram View and drag Sales[ProductID] onto Products[ProductID], then Sales[Order Date] onto Calendar[Date].", ""),
            ("In the Calculation Area write the explicit measures, one per cell.",
             "Total Sales := SUM(Sales[Amount])"),
            ("Add the remaining measures, then build a PivotTable from the model to verify them.",
             "Order Count := COUNTROWS(Sales)  |  Average Order := DIVIDE([Total Sales], [Order Count])  |  Sales YTD := TOTALYTD([Total Sales], Calendar[Date])"),
        ],
        test="A PivotTable with Products[Category] on rows and [Total Sales] in values returns per-category totals that sum to the grand total.",
    ),
]
