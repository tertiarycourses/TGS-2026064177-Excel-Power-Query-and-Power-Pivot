# Lab 12 — Summarise with Group By

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 2:** Data Analysis with Power Query and Power Pivot  
**Maps to:** LO2 · A2 · K1, K3 — Aggregate rows to a summary grain with Group By.

**Workbook:** `lab-12-group-by.xlsx`

## Scenario

Rather than loading 20,000 transaction rows into the grid, the learner aggregates them inside the query. The lab covers basic and advanced Group By, multiple aggregations, grouping on more than one column, and keeping the underlying rows with All Rows.

## What you will build

A regional sales summary with count, sum and average in one query.

**Tools used:** Excel, Power Query Editor, Group By

## Steps

1. Load the Transactions range into the Power Query Editor.

2. Transform > Group By, switch to Advanced, and group by Region.

3. Add three aggregations: Order Count (Count Rows), Total Sales (Sum of Amount), Average Order (Average of Amount).

4. Add Product Category as a second Group By column and observe the finer grain.

5. Add an All Rows aggregation to keep the detail rows nested in a table column for drill-through.

## Test it

Total Sales across all groups equals the grand total of the source Amount column.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*