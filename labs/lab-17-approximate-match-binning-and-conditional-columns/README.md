# Lab 17 — Approximate Match: Binning and Conditional Columns

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 3:** Advanced Data Analysis using Power Query  
**Maps to:** LO3 · A3 · K2, K3 — Reproduce VLOOKUP's approximate match two ways.

**Workbook:** `lab-17-approximate-match.xlsx`

## Scenario

Power Query merges are exact-match only, so VLOOKUP's TRUE behaviour must be rebuilt. The learner solves a commission-band lookup twice: once with a conditional column for fixed bands, and once with a banding table plus a merge so the bands stay data-driven.

## What you will build

A commission-rate column produced by both the conditional-column and banding-table methods.

**Tools used:** Excel, Power Query Editor, Conditional Column, Merge Queries

## Steps

1. Method 1: Add Column > Conditional Column named Rate A with the bands 0-9999 -> 0.02, 10000-49999 -> 0.05, else 0.08.

2. Note the maintenance cost: changing a band means editing the query, not a worksheet cell.

3. Method 2: load the CommissionBands table, which carries Lower, Upper and Rate columns.

4. Cross join the two tables, then filter to rows where Sales is between Lower and Upper using a custom column.

   ```m
   [Sales] >= [Lower] and [Sales] <= [Upper]
   ```

5. Compare Rate A and the merged rate — every row must agree.

## Test it

Both methods return identical rates, including at the exact band boundaries of 10000 and 50000.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*