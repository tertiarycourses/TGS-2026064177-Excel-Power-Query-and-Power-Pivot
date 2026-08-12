# Lab 08 — Nest IF with AND for a Priority Flag

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 2:** Data Analysis with Power Query and Power Pivot  
**Maps to:** LO2 · A2 · K2 — Nest conditions with and / or to resolve a multi-factor rule.

**Workbook:** `lab-08-nested-if-and.xlsx`

## Scenario

An order is escalated only when it is both high value and already overdue. The learner nests conditions using M's and / or operators to produce a single Priority column, then verifies each branch against a worked truth table.

## What you will build

A Priority column classifying every order as Escalate, Watch or Normal.

**Tools used:** Excel, Power Query Editor, Custom Column, M logical operators

## Steps

1. Load the Orders range, which carries Amount, Status and Days Outstanding.

2. Add Column > Custom Column named Priority with a nested M expression.

   ```m
   if [Amount] >= 5000 and [Days Outstanding] > 30 then "Escalate" else if [Amount] >= 5000 or [Days Outstanding] > 30 then "Watch" else "Normal"
   ```

3. Confirm that and requires both conditions while or requires either, using the four-row truth table in the guide.

4. Add a Custom Column named Review Needed combining a text test with a numeric test.

   ```m
   if [Status] <> "Closed" and [Days Outstanding] > 60 then "Yes" else "No"
   ```

5. Filter Priority to Escalate and record how many orders qualify.

## Test it

Only orders meeting BOTH the value and ageing tests read Escalate; orders meeting one read Watch.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*