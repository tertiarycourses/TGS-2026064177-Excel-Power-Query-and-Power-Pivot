# Lab 03 — Fill Down and Fill Up a Ragged Report

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 1:** Review Data with Power Query  
**Maps to:** LO1 · A1 · K2 — Repair merged-cell style gaps with Fill Down and Fill Up.

**Workbook:** `lab-03-fill-down-up.xlsx`

## Scenario

A departmental headcount report was typed for human eyes: the region name appears only on the first row of each block, and the reporting quarter sits below its group. The learner uses Fill Down and Fill Up to push those values across every row so the table becomes machine-readable.

## What you will build

A rectangular headcount table with no blank grouping cells.

**Tools used:** Excel, Power Query Editor, Fill Down / Fill Up

## Steps

1. Load the Headcount range with Data > From Table/Range.

2. Select the Region column, then Transform > Fill > Down to carry each region into its blank rows.

3. Select the Quarter column, then Transform > Fill > Up to lift the quarter into the rows above it.

4. Use Home > Remove Rows > Remove Blank Rows to drop the spacer rows left by the original layout.

5. Close & Load to a new worksheet.

## Test it

Every row carries a Region and a Quarter value, and the row count matches the number of real headcount records.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*