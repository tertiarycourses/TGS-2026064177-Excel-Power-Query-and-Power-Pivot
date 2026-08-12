# Lab 18 — Cross Join and Drill Down to a Variable

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 3:** Advanced Data Analysis using Power Query  
**Maps to:** LO3 · A3 · K1, K2 — Build a complete grid with a cross join and extract a scalar by drill down.

**Workbook:** `lab-18-cross-join-drilldown.xlsx`

## Scenario

A sales matrix must show every region against every product, including combinations with no sales, so the report has no gaps. The learner builds that complete grid with a cross join, then uses Drill Down to turn a single cell into a reusable variable such as a target threshold.

## What you will build

A complete region x product grid plus a drilled-down threshold variable.

**Tools used:** Excel, Power Query Editor, Custom Column, Drill Down

## Steps

1. Load Regions and Products as Connection Only queries.

2. On Regions add a Custom Column whose formula is simply the Products query, creating a nested table on every row.

   ```m
   Products
   ```

3. Expand that column to produce every region-product pair — the cross join.

4. Merge the grid with Sales using Left Outer so unsold combinations appear as null, then replace null with 0.

5. On the Parameters query right-click the threshold cell and choose Drill Down; the query now returns a single value usable in a filter step.

## Test it

The grid row count equals regions x products, and unsold combinations show 0 rather than being absent.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*