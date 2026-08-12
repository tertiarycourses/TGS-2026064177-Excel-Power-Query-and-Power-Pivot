"""
Topic 1 — Review Data with Power Query.

Covers: data types, query from table/range, query from another workbook,
connection-only load, fill up/down, split by delimiter, split into rows,
conditional and custom columns, column from examples, merge columns,
sort and filter.

Each lab ships as labs/lab-NN-<slug>/ containing a starter .xlsx with mock data
(and, where the lab builds on a prior query, the Power Query queries already
embedded) plus the lab instructions.
"""

DOMAIN1 = [
    dict(
        num=1, topic=1,
        title="Load Sales Data and Fix Data Types",
        objective="LO1 · A1 · K1, K4 — Query data from a table or range and set correct data types.",
        desc=(
            "The Southbridge Trading sales extract arrives with everything stored as text: "
            "order dates read as words, amounts carry a stray currency prefix, and quantities "
            "will not sum. The learner loads the range into Power Query, inspects each column "
            "with the data-profiling tools, and applies the correct data type to every column "
            "so the table can be aggregated."
        ),
        build="A clean Sales query loaded to a worksheet with every column correctly typed.",
        services="Excel, Power Query Editor, Data Profiling",
        workbook="lab-01-load-and-data-types.xlsx",
        sheets=["Sales"],
        steps=[
            ("Click any cell inside the Sales range, then Data > From Table/Range. Tick 'My table has headers'.", ""),
            ("In the Power Query Editor turn on View > Column quality, Column distribution and Column profile.", ""),
            ("Set Order Date to Date, Quantity to Whole Number, Unit Price and Amount to Decimal Number.", ""),
            ("Where a type change fails, use Transform > Data Type > Using Locale to read the date as English (Singapore).", ""),
            ("Rename the query to Sales, then Home > Close & Load To > Table on a new worksheet.", ""),
        ],
        test="Quantity and Amount sum in the status bar, Order Date filters offer date pickers, and Column quality reports 0% error on every column.",
    ),
    dict(
        num=2, topic=1,
        title="Combine an External Workbook as a Connection Only",
        objective="LO1 · A1 · K4 — Query data from another Excel file and load as connection only.",
        desc=(
            "The regional price list lives in a separate workbook that the finance team refreshes "
            "monthly. The learner connects to that external file, loads it as Connection Only so it "
            "does not clutter the workbook grid, and confirms the connection refreshes independently."
        ),
        build="A PriceList staging query held as Connection Only, ready to be merged later.",
        services="Excel, Power Query Editor, External workbook source",
        workbook="lab-02-external-connection.xlsx",
        extra_files=["price-list-source.xlsx"],
        sheets=["Products"],
        steps=[
            ("Data > Get Data > From File > From Excel Workbook and select price-list-source.xlsx.", ""),
            ("In the Navigator select the PriceList table and click Transform Data rather than Load.", ""),
            ("Promote the first row to headers if needed, then set Unit Price to Decimal Number.", ""),
            ("Home > Close & Load To… > Only Create Connection. Leave 'Add this data to the Data Model' unticked.", ""),
            ("Open Queries & Connections and confirm PriceList reports 'Connection only'.", ""),
        ],
        test="The Queries & Connections pane lists PriceList as 'Connection only', and no new worksheet was created.",
    ),
    dict(
        num=3, topic=1,
        title="Fill Down and Fill Up a Ragged Report",
        objective="LO1 · A1 · K2 — Repair merged-cell style gaps with Fill Down and Fill Up.",
        desc=(
            "A departmental headcount report was typed for human eyes: the region name appears only "
            "on the first row of each block, and the reporting quarter sits below its group. The "
            "learner uses Fill Down and Fill Up to push those values across every row so the table "
            "becomes machine-readable."
        ),
        build="A rectangular headcount table with no blank grouping cells.",
        services="Excel, Power Query Editor, Fill Down / Fill Up",
        workbook="lab-03-fill-down-up.xlsx",
        sheets=["Headcount"],
        steps=[
            ("Load the Headcount range with Data > From Table/Range.", ""),
            ("Select the Region column, then Transform > Fill > Down to carry each region into its blank rows.", ""),
            ("Select the Quarter column, then Transform > Fill > Up to lift the quarter into the rows above it.", ""),
            ("Use Home > Remove Rows > Remove Blank Rows to drop the spacer rows left by the original layout.", ""),
            ("Close & Load to a new worksheet.", ""),
        ],
        test="Every row carries a Region and a Quarter value, and the row count matches the number of real headcount records.",
    ),
    dict(
        num=4, topic=1,
        title="Split Columns by Delimiter and into Rows",
        objective="LO1 · A1 · K2 — Split a column by delimiter, into columns and into rows.",
        desc=(
            "The CRM export packs the customer's full name into one field and lists every product "
            "a customer bought as a comma-separated string. The learner splits the name into "
            "First and Last Name columns, then splits the product list into rows so each purchase "
            "becomes its own record."
        ),
        build="A normalised customer table with one product per row.",
        services="Excel, Power Query Editor, Split Column",
        workbook="lab-04-split-columns.xlsx",
        sheets=["Customers"],
        steps=[
            ("Load the Customers range into the Power Query Editor.", ""),
            ("Select Full Name, then Transform > Split Column > By Delimiter, choose Space, and split at the left-most delimiter.", ""),
            ("Rename the resulting columns to First Name and Last Name.", ""),
            ("Select Products Bought, then Split Column > By Delimiter, choose Comma, and open Advanced options > Split into Rows.", ""),
            ("Use Transform > Format > Trim on the new product column to remove the leading spaces.", ""),
        ],
        test="A customer who bought three products now occupies three rows, and no product value has a leading space.",
    ),
    dict(
        num=5, topic=1,
        title="Add Conditional, Custom and Example-Driven Columns",
        objective="LO1 · A1 · K1, K2 — Add conditional columns, custom columns and Column From Examples.",
        desc=(
            "Management wants each order classified into a value band, a line total calculated, and "
            "an order reference built from existing fields. The learner uses the Conditional Column "
            "dialog for the band, a Custom Column with an M expression for the total, and Column "
            "From Examples to let Power Query infer the reference pattern."
        ),
        build="A Sales query carrying Value Band, Line Total and Order Ref columns.",
        services="Excel, Power Query Editor, Conditional Column, Custom Column, Column From Examples",
        workbook="lab-05-add-columns.xlsx",
        sheets=["Orders"],
        steps=[
            ("Load the Orders range. Add Column > Conditional Column, name it Value Band.", ""),
            ("Set the rules: Amount >= 5000 -> High, Amount >= 1000 -> Medium, else Low. Order matters; the first true rule wins.", ""),
            ("Add Column > Custom Column, name it Line Total, and enter the M expression.", "[Quantity] * [Unit Price]"),
            ("Add Column > Column From Examples > From All Columns, then type SBT-NORTH-1001 on the first row so Power Query infers the pattern.", ""),
            ("Rename the inferred column to Order Ref and set Line Total to Decimal Number.", ""),
        ],
        test="Value Band shows High/Medium/Low correctly at the 1000 and 5000 boundaries, and Line Total equals Quantity x Unit Price on every row.",
    ),
    dict(
        num=6, topic=1,
        title="Merge Columns, Sort and Filter",
        objective="LO1 · A1 · K2, K3 — Merge columns, then sort and filter data in Power Query.",
        desc=(
            "The despatch team needs a single address label field and a shortlist of the orders that "
            "actually need attention. The learner merges the address parts into one column with a "
            "chosen separator, then applies sorting and filtering inside the query so the refresh "
            "always returns the shortlist."
        ),
        build="A despatch shortlist query, sorted and filtered at source.",
        services="Excel, Power Query Editor, Merge Columns, Sort, Filter",
        workbook="lab-06-merge-sort-filter.xlsx",
        sheets=["Despatch"],
        steps=[
            ("Load the Despatch range. Ctrl-click Block, Street and Postal Code in that order.", ""),
            ("Transform > Merge Columns, choose Comma as the separator, and name the result Address Label.", ""),
            ("Filter Status to exclude Cancelled using the column drop-down.", ""),
            ("Filter Amount with Number Filters > Greater Than 1000.", ""),
            ("Sort Order Date descending, then Amount descending, and note the step order in Applied Steps.", ""),
        ],
        test="The query returns only non-cancelled orders above 1000, newest first, and Address Label reads as one comma-separated string.",
    ),
]
