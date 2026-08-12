# Lab 11 — Pivot, Pivot Without Aggregation, and Append

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 2:** Data Analysis with Power Query and Power Pivot  
**Maps to:** LO2 · A2 · K2, K3 — Pivot a column, use Don't Aggregate, and append queries.

**Workbook:** `lab-11-pivot-append.xlsx`

## Scenario

The learner reverses an unpivot to rebuild a cross-tab, meets the Don't Aggregate option that preserves individual values instead of summing them, and then appends the Q1 and Q2 extracts into a single continuous table.

## What you will build

A pivoted cross-tab plus an appended half-year table.

**Tools used:** Excel, Power Query Editor, Pivot Column, Append Queries

## Steps

1. Load Q1Sales and Q2Sales as separate queries, each as Connection Only.

2. On the Attributes query select the Attribute column, then Transform > Pivot Column with Value as the values column.

3. In Advanced options choose Don't Aggregate so text values are preserved rather than counted.

4. Home > Append Queries as New, add Q1Sales and Q2Sales, and name the result HalfYear.

5. Confirm both quarters' column names match exactly — a mismatched name creates a new column full of nulls.

## Test it

HalfYear's row count equals Q1 rows plus Q2 rows, with no column of all nulls.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*