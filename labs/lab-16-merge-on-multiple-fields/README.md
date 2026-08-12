# Lab 16 — Merge on Multiple Fields

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 3:** Advanced Data Analysis using Power Query  
**Maps to:** LO3 · A3 · K2 — Match on a composite key using more than one column.

**Workbook:** `lab-16-multi-field-merge.xlsx`

## Scenario

Prices vary by both product and region, so a single-column match returns the wrong price. The learner Ctrl-clicks two columns in each table to build a composite key, and learns that column selection order determines the pairing.

## What you will build

A correctly priced sales query matched on Product plus Region.

**Tools used:** Excel, Power Query Editor, Merge Queries

## Steps

1. Merge Sales with RegionalPrices on ProductID alone and observe the row count inflate — one order matches several regional prices.

2. Delete that step. Re-merge, Ctrl-clicking ProductID then Region in Sales, and ProductID then Region in RegionalPrices.

3. Check the match indicator at the foot of the Merge dialog before clicking OK.

4. Expand only the Unit Price column from the matched table.

5. Confirm the row count now equals the original Sales row count.

## Test it

Row count matches Sales exactly, and a North-region order carries the North price, not the South one.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*