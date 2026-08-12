# Lab 04 — Split Columns by Delimiter and into Rows

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 1:** Review Data with Power Query  
**Maps to:** LO1 · A1 · K2 — Split a column by delimiter, into columns and into rows.

**Workbook:** `lab-04-split-columns.xlsx`

## Scenario

The CRM export packs the customer's full name into one field and lists every product a customer bought as a comma-separated string. The learner splits the name into First and Last Name columns, then splits the product list into rows so each purchase becomes its own record.

## What you will build

A normalised customer table with one product per row.

**Tools used:** Excel, Power Query Editor, Split Column

## Steps

1. Load the Customers range into the Power Query Editor.

2. Select Full Name, then Transform > Split Column > By Delimiter, choose Space, and split at the left-most delimiter.

3. Rename the resulting columns to First Name and Last Name.

4. Select Products Bought, then Split Column > By Delimiter, choose Comma, and open Advanced options > Split into Rows.

5. Use Transform > Format > Trim on the new product column to remove the leading spaces.

## Test it

A customer who bought three products now occupies three rows, and no product value has a leading space.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*