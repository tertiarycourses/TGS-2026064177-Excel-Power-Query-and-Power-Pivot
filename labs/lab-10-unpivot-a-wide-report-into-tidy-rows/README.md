# Lab 10 — Unpivot a Wide Report into Tidy Rows

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 2:** Data Analysis with Power Query and Power Pivot  
**Maps to:** LO2 · A2 · K2, K3 — Unpivot columns, and choose between Unpivot Other Columns and Unpivot Only Selected.

**Workbook:** `lab-10-unpivot.xlsx`

## Scenario

The monthly sales report is laid out with one column per month — readable for a human, useless for a PivotTable. The learner unpivots the month columns into Attribute/Value pairs and learns why Unpivot Other Columns survives a new month being added while Unpivot Only Selected Columns silently misses it.

## What you will build

A tidy long-format sales table with Month and Sales columns.

**Tools used:** Excel, Power Query Editor, Unpivot

## Steps

1. Load the MonthlySales range, which has Region plus twelve month columns.

2. Select the Region column, then Transform > Unpivot Other Columns.

3. Rename Attribute to Month and Value to Sales, then set Sales to Decimal Number.

4. Compare with Unpivot Only Selected Columns: the guide explains that adding a 13th month breaks the selected-columns version but not Unpivot Other Columns.

5. Close & Load and build a quick PivotTable to prove the tidy shape now pivots correctly.

## Test it

The result has exactly three columns (Region, Month, Sales) and a row count of regions x months.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*