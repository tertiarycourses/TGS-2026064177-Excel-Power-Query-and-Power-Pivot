# Lab 05 — Add Conditional, Custom and Example-Driven Columns

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 1:** Review Data with Power Query  
**Maps to:** LO1 · A1 · K1, K2 — Add conditional columns, custom columns and Column From Examples.

**Workbook:** `lab-05-add-columns.xlsx`

## Scenario

Management wants each order classified into a value band, a line total calculated, and an order reference built from existing fields. The learner uses the Conditional Column dialog for the band, a Custom Column with an M expression for the total, and Column From Examples to let Power Query infer the reference pattern.

## What you will build

A Sales query carrying Value Band, Line Total and Order Ref columns.

**Tools used:** Excel, Power Query Editor, Conditional Column, Custom Column, Column From Examples

## Steps

1. Load the Orders range. Add Column > Conditional Column, name it Value Band.

2. Set the rules: Amount >= 5000 -> High, Amount >= 1000 -> Medium, else Low. Order matters; the first true rule wins.

3. Add Column > Custom Column, name it Line Total, and enter the M expression.

   ```m
   [Quantity] * [Unit Price]
   ```

4. Add Column > Column From Examples > From All Columns, then type SBT-NORTH-1001 on the first row so Power Query infers the pattern.

5. Rename the inferred column to Order Ref and set Line Total to Decimal Number.

## Test it

Value Band shows High/Medium/Low correctly at the 1000 and 5000 boundaries, and Line Total equals Quantity x Unit Price on every row.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*