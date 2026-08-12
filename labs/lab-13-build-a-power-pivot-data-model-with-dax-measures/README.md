# Lab 13 — Build a Power Pivot Data Model with DAX Measures

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 2:** Data Analysis with Power Query and Power Pivot  
**Maps to:** LO2 · A2 · K3, K4 — Load to the Data Model, create relationships and write DAX measures.

**Workbook:** `lab-13-power-pivot-model.xlsx`

## Scenario

The learner loads the fact and lookup tables to the Power Pivot Data Model, draws relationships in Diagram View, and writes explicit DAX measures. The lab contrasts an implicit measure (dragging a field into Values) with an explicit one, and explains why explicit measures are the professional practice.

## What you will build

A Power Pivot data model with three related tables and four DAX measures.

**Tools used:** Excel, Power Query Editor, Power Pivot, DAX, Diagram View

## Steps

1. Enable Power Pivot via File > Options > Add-ins > COM Add-ins if the tab is not visible.

2. Load Sales, Products and Calendar with Close & Load To… > Only Create Connection, ticking 'Add this data to the Data Model'.

3. Open Power Pivot > Manage > Diagram View and drag Sales[ProductID] onto Products[ProductID], then Sales[Order Date] onto Calendar[Date].

4. In the Calculation Area write the explicit measures, one per cell.

   ```m
   Total Sales := SUM(Sales[Amount])
   ```

5. Add the remaining measures, then build a PivotTable from the model to verify them.

   ```m
   Order Count := COUNTROWS(Sales)
   Average Order := DIVIDE([Total Sales], [Order Count])
   Sales YTD := TOTALYTD([Total Sales], Calendar[Date])
   ```

## Test it

A PivotTable with Products[Category] on rows and [Total Sales] in values returns per-category totals that sum to the grand total.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*