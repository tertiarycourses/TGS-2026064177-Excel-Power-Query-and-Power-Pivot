# Lab 09 — Calculate Deadlines with Date.AddDays

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 2:** Data Analysis with Power Query and Power Pivot  
**Maps to:** LO2 · A2 · K1, K3 — Use Date functions to derive deadlines and ageing buckets.

**Workbook:** `lab-09-adddays-deadline.xlsx`

## Scenario

Each invoice carries payment terms in days. The learner derives the due date with Date.AddDays, compares it against a fixed reporting date to produce days overdue, and bins the result into ageing buckets the finance team recognises.

## What you will build

An invoice ageing query with Due Date, Days Overdue and Ageing Bucket.

**Tools used:** Excel, Power Query Editor, Date functions, Custom Column

## Steps

1. Load the Invoices range and confirm Invoice Date is typed as Date.

2. Add Column > Custom Column named Due Date using the invoice date plus its terms.

   ```m
   Date.AddDays([Invoice Date], [Terms Days])
   ```

3. Set Due Date to Date, then add a Custom Column named Days Overdue against the 30 June 2026 reporting date.

   ```m
   Duration.Days(#date(2026,6,30) - [Due Date])
   ```

4. Add a Conditional Column named Ageing Bucket: <= 0 Current, <= 30 1-30 Days, <= 60 31-60 Days, else 60+ Days.

5. Sort by Days Overdue descending to bring the worst offenders to the top.

## Test it

An invoice dated 1 May 2026 with 30-day terms shows Due Date 31 May 2026 and Days Overdue 30.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*