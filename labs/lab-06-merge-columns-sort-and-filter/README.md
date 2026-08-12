# Lab 06 — Merge Columns, Sort and Filter

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 1:** Review Data with Power Query  
**Maps to:** LO1 · A1 · K2, K3 — Merge columns, then sort and filter data in Power Query.

**Workbook:** `lab-06-merge-sort-filter.xlsx`

## Scenario

The despatch team needs a single address label field and a shortlist of the orders that actually need attention. The learner merges the address parts into one column with a chosen separator, then applies sorting and filtering inside the query so the refresh always returns the shortlist.

## What you will build

A despatch shortlist query, sorted and filtered at source.

**Tools used:** Excel, Power Query Editor, Merge Columns, Sort, Filter

## Steps

1. Load the Despatch range. Ctrl-click Block, Street and Postal Code in that order.

2. Transform > Merge Columns, choose Comma as the separator, and name the result Address Label.

3. Filter Status to exclude Cancelled using the column drop-down.

4. Filter Amount with Number Filters > Greater Than 1000.

5. Sort Order Date descending, then Amount descending, and note the step order in Applied Steps.

## Test it

The query returns only non-cancelled orders above 1000, newest first, and Address Label reads as one comma-separated string.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*