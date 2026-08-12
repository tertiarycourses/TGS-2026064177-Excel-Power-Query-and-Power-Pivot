# Lab 14 — Walk Through All Six Join Kinds

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 3:** Advanced Data Analysis using Power Query  
**Maps to:** LO3 · A3 · K2, K3 — Compare the six Power Query join kinds on the same pair of tables.

**Workbook:** `lab-14-six-joins.xlsx`

## Scenario

Using one deliberately imperfect pair of tables — some orders reference a missing product, some products were never ordered — the learner runs all six join kinds in turn and records the row count each produces. The resulting comparison table is the reference sheet for every merge decision afterwards.

## What you will build

Six merge queries and a completed row-count comparison table.

**Tools used:** Excel, Power Query Editor, Merge Queries

## Steps

1. Load Orders and Products as Connection Only queries.

2. Home > Merge Queries as New, match Orders[ProductID] to Products[ProductID], and choose Left Outer. Record the row count.

3. Repeat for Right Outer, Full Outer, Inner, Left Anti and Right Anti, naming each query after its join kind.

4. Note that Left Anti returns orders with no matching product — the orphan records.

5. Complete the comparison table in the guide, then state which join answers 'which products have never sold?'.

## Test it

Left Outer returns exactly the Orders row count; Inner returns fewer; Left Anti plus Inner equals Left Outer.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*