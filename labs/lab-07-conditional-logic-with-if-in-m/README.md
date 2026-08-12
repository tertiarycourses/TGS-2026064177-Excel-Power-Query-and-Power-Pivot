# Lab 07 — Conditional Logic with IF in M

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 2:** Data Analysis with Power Query and Power Pivot  
**Maps to:** LO2 · A2 · K1, K2 — Write IF logic in the Power Query M language.

**Workbook:** `lab-07-if-formulas.xlsx`

## Scenario

The Conditional Column dialog cannot express every rule, so the learner moves to the Custom Column window and writes M directly. The lab contrasts Excel's IF() with M's lowercase, case-sensitive if / then / else and shows why a missing else is an error.

## What you will build

A Shipping Cost column driven by a hand-written M if expression.

**Tools used:** Excel, Power Query Editor, Custom Column, M language

## Steps

1. Load the Orders range. Add Column > Custom Column and name it Shipping Cost.

2. Enter the M expression, noting that if, then and else must be lowercase.

   ```m
   if [Amount] >= 500 then 0 else 15
   ```

3. Delete the else clause and read the error Power Query raises, then restore it. Every if needs an else.

4. Add a second Custom Column named Free Shipping using a comparison that returns a logical value.

   ```m
   [Amount] >= 500
   ```

5. Set Shipping Cost to Decimal Number and Free Shipping to True/False.

## Test it

Orders of 500 or more show Shipping Cost 0 and Free Shipping TRUE; the boundary row at exactly 500 qualifies.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*