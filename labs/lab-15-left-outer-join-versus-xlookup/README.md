# Lab 15 — Left Outer Join versus XLOOKUP

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 3:** Advanced Data Analysis using Power Query  
**Maps to:** LO3 · A3 · K2, K3 — Reproduce an XLOOKUP with a merge and expand only the needed column.

**Workbook:** `lab-15-merge-vs-xlookup.xlsx`

## Scenario

The learner replaces a worksheet full of XLOOKUP formulas with a single Left Outer merge, expands only the required column, and compares the two approaches on refresh behaviour, row-count safety and performance on a large table.

## What you will build

A Sales query enriched with product details by merge instead of formulas.

**Tools used:** Excel, Power Query Editor, Merge Queries, Expand

## Steps

1. Inspect the existing XLOOKUP column on the worksheet and note it must be dragged down when rows are added.

2. Load Sales into Power Query, then Home > Merge Queries with Products on ProductID using Left Outer.

3. Click the expand arrow on the new table column and untick 'Use original column name as prefix'.

4. Select only Product Name and Category — expanding every column bloats the query for no benefit.

5. Confirm the row count is unchanged after the merge, proving no duplication occurred.

## Test it

The merged query has the same row count as Sales, and unmatched rows show null rather than #N/A.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*