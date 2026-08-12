# Lab 01 — Load Sales Data and Fix Data Types

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 1:** Review Data with Power Query  
**Maps to:** LO1 · A1 · K1, K4 — Query data from a table or range and set correct data types.

**Workbook:** `lab-01-load-and-data-types.xlsx`

## Scenario

The Southbridge Trading sales extract arrives with everything stored as text: order dates read as words, amounts carry a stray currency prefix, and quantities will not sum. The learner loads the range into Power Query, inspects each column with the data-profiling tools, and applies the correct data type to every column so the table can be aggregated.

## What you will build

A clean Sales query loaded to a worksheet with every column correctly typed.

**Tools used:** Excel, Power Query Editor, Data Profiling

## Steps

1. Click any cell inside the Sales range, then Data > From Table/Range. Tick 'My table has headers'.

2. In the Power Query Editor turn on View > Column quality, Column distribution and Column profile.

3. Set Order Date to Date, Quantity to Whole Number, Unit Price and Amount to Decimal Number.

4. Where a type change fails, use Transform > Data Type > Using Locale to read the date as English (Singapore).

5. Rename the query to Sales, then Home > Close & Load To > Table on a new worksheet.

## Test it

Quantity and Amount sum in the status bar, Order Date filters offer date pickers, and Column quality reports 0% error on every column.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*