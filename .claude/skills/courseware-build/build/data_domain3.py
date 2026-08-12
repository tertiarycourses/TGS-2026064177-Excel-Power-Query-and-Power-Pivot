"""
Topic 3 — Advanced Data Analysis using Power Query.

Covers: overview of joins, all six join kinds, left vs right, outer join versus
XLOOKUP, merge on multiple fields, approximate match via binning, approximate
match via conditional column, cross join, drill down to create a variable,
fuzzy matching by percentage, and merging inconsistent data with a
transformation table.
"""

DOMAIN3 = [
    dict(
        num=14, topic=3,
        title="Walk Through All Six Join Kinds",
        objective="LO3 · A3 · K2, K3 — Compare the six Power Query join kinds on the same pair of tables.",
        desc=(
            "Using one deliberately imperfect pair of tables — some orders reference a missing "
            "product, some products were never ordered — the learner runs all six join kinds in turn "
            "and records the row count each produces. The resulting comparison table is the "
            "reference sheet for every merge decision afterwards."
        ),
        build="Six merge queries and a completed row-count comparison table.",
        services="Excel, Power Query Editor, Merge Queries",
        workbook="lab-14-six-joins.xlsx",
        sheets=["Orders", "Products"],
        steps=[
            ("Load Orders and Products as Connection Only queries.", ""),
            ("Home > Merge Queries as New, match Orders[ProductID] to Products[ProductID], and choose Left Outer. Record the row count.", ""),
            ("Repeat for Right Outer, Full Outer, Inner, Left Anti and Right Anti, naming each query after its join kind.", ""),
            ("Note that Left Anti returns orders with no matching product — the orphan records.", ""),
            ("Complete the comparison table in the guide, then state which join answers 'which products have never sold?'.", ""),
        ],
        test="Left Outer returns exactly the Orders row count; Inner returns fewer; Left Anti plus Inner equals Left Outer.",
    ),
    dict(
        num=15, topic=3,
        title="Left Outer Join versus XLOOKUP",
        objective="LO3 · A3 · K2, K3 — Reproduce an XLOOKUP with a merge and expand only the needed column.",
        desc=(
            "The learner replaces a worksheet full of XLOOKUP formulas with a single Left Outer "
            "merge, expands only the required column, and compares the two approaches on refresh "
            "behaviour, row-count safety and performance on a large table."
        ),
        build="A Sales query enriched with product details by merge instead of formulas.",
        services="Excel, Power Query Editor, Merge Queries, Expand",
        workbook="lab-15-merge-vs-xlookup.xlsx",
        sheets=["Sales", "Products"],
        steps=[
            ("Inspect the existing XLOOKUP column on the worksheet and note it must be dragged down when rows are added.", ""),
            ("Load Sales into Power Query, then Home > Merge Queries with Products on ProductID using Left Outer.", ""),
            ("Click the expand arrow on the new table column and untick 'Use original column name as prefix'.", ""),
            ("Select only Product Name and Category — expanding every column bloats the query for no benefit.", ""),
            ("Confirm the row count is unchanged after the merge, proving no duplication occurred.", ""),
        ],
        test="The merged query has the same row count as Sales, and unmatched rows show null rather than #N/A.",
    ),
    dict(
        num=16, topic=3,
        title="Merge on Multiple Fields",
        objective="LO3 · A3 · K2 — Match on a composite key using more than one column.",
        desc=(
            "Prices vary by both product and region, so a single-column match returns the wrong "
            "price. The learner Ctrl-clicks two columns in each table to build a composite key, and "
            "learns that column selection order determines the pairing."
        ),
        build="A correctly priced sales query matched on Product plus Region.",
        services="Excel, Power Query Editor, Merge Queries",
        workbook="lab-16-multi-field-merge.xlsx",
        sheets=["Sales", "RegionalPrices"],
        steps=[
            ("Merge Sales with RegionalPrices on ProductID alone and observe the row count inflate — one order matches several regional prices.", ""),
            ("Delete that step. Re-merge, Ctrl-clicking ProductID then Region in Sales, and ProductID then Region in RegionalPrices.", ""),
            ("Check the match indicator at the foot of the Merge dialog before clicking OK.", ""),
            ("Expand only the Unit Price column from the matched table.", ""),
            ("Confirm the row count now equals the original Sales row count.", ""),
        ],
        test="Row count matches Sales exactly, and a North-region order carries the North price, not the South one.",
    ),
    dict(
        num=17, topic=3,
        title="Approximate Match: Binning and Conditional Columns",
        objective="LO3 · A3 · K2, K3 — Reproduce VLOOKUP's approximate match two ways.",
        desc=(
            "Power Query merges are exact-match only, so VLOOKUP's TRUE behaviour must be rebuilt. "
            "The learner solves a commission-band lookup twice: once with a conditional column for "
            "fixed bands, and once with a banding table plus a merge so the bands stay data-driven."
        ),
        build="A commission-rate column produced by both the conditional-column and banding-table methods.",
        services="Excel, Power Query Editor, Conditional Column, Merge Queries",
        workbook="lab-17-approximate-match.xlsx",
        sheets=["SalesReps", "CommissionBands"],
        steps=[
            ("Method 1: Add Column > Conditional Column named Rate A with the bands 0-9999 -> 0.02, 10000-49999 -> 0.05, else 0.08.", ""),
            ("Note the maintenance cost: changing a band means editing the query, not a worksheet cell.", ""),
            ("Method 2: load the CommissionBands table, which carries Lower, Upper and Rate columns.", ""),
            ("Cross join the two tables, then filter to rows where Sales is between Lower and Upper using a custom column.",
             "[Sales] >= [Lower] and [Sales] <= [Upper]"),
            ("Compare Rate A and the merged rate — every row must agree.", ""),
        ],
        test="Both methods return identical rates, including at the exact band boundaries of 10000 and 50000.",
    ),
    dict(
        num=18, topic=3,
        title="Cross Join and Drill Down to a Variable",
        objective="LO3 · A3 · K1, K2 — Build a complete grid with a cross join and extract a scalar by drill down.",
        desc=(
            "A sales matrix must show every region against every product, including combinations "
            "with no sales, so the report has no gaps. The learner builds that complete grid with a "
            "cross join, then uses Drill Down to turn a single cell into a reusable variable such as "
            "a target threshold."
        ),
        build="A complete region x product grid plus a drilled-down threshold variable.",
        services="Excel, Power Query Editor, Custom Column, Drill Down",
        workbook="lab-18-cross-join-drilldown.xlsx",
        sheets=["Regions", "Products", "Parameters", "Sales"],
        steps=[
            ("Load Regions and Products as Connection Only queries.", ""),
            ("On Regions add a Custom Column whose formula is simply the Products query, creating a nested table on every row.", "Products"),
            ("Expand that column to produce every region-product pair — the cross join.", ""),
            ("Merge the grid with Sales using Left Outer so unsold combinations appear as null, then replace null with 0.", ""),
            ("On the Parameters query right-click the threshold cell and choose Drill Down; the query now returns a single value usable in a filter step.", ""),
        ],
        test="The grid row count equals regions x products, and unsold combinations show 0 rather than being absent.",
    ),
    dict(
        num=19, topic=3,
        title="Fuzzy Matching and Transformation Tables",
        objective="LO3 · A3 · K2, K3 — Match inconsistent text with fuzzy matching and a transformation table.",
        desc=(
            "Customer names arrive spelled inconsistently across two systems — 'Pte Ltd' versus "
            "'Pte. Ltd.', typos, and abbreviations. The learner enables fuzzy matching, tunes the "
            "similarity threshold, and then supplies a transformation table to teach the engine the "
            "pairs that no similarity score will ever catch."
        ),
        build="A reconciled customer query matched by fuzzy logic plus a transformation table.",
        services="Excel, Power Query Editor, Fuzzy Merge, Transformation Table",
        workbook="lab-19-fuzzy-matching.xlsx",
        sheets=["CRMCustomers", "FinanceCustomers", "TransformTable"],
        steps=[
            ("Merge CRMCustomers with FinanceCustomers on Customer Name using Left Outer and note how many rows fail to match.", ""),
            ("Edit the merge, tick 'Use fuzzy matching to perform the merge', and open Fuzzy matching options.", ""),
            ("Set the Similarity threshold to 0.8 and tick 'Ignore case' and 'Ignore space'. Record the improvement.", ""),
            ("Lower the threshold to 0.6 and observe false matches appearing — the trade-off between recall and precision.", ""),
            ("Return to 0.8 and set the TransformTable query as the Transformation table so 'ABC Holdings' maps to 'ABC Group' explicitly.", ""),
        ],
        test="All customers reconcile at threshold 0.8 with the transformation table applied, and no false match is introduced.",
    ),
]
