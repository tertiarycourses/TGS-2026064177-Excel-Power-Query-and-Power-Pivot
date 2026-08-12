# Lab 02 — Combine an External Workbook as a Connection Only

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 1:** Review Data with Power Query  
**Maps to:** LO1 · A1 · K4 — Query data from another Excel file and load as connection only.

**Workbook:** `lab-02-external-connection.xlsx`
**Also in this folder:** `price-list-source.xlsx`

## Scenario

The regional price list lives in a separate workbook that the finance team refreshes monthly. The learner connects to that external file, loads it as Connection Only so it does not clutter the workbook grid, and confirms the connection refreshes independently.

## What you will build

A PriceList staging query held as Connection Only, ready to be merged later.

**Tools used:** Excel, Power Query Editor, External workbook source

## Steps

1. Data > Get Data > From File > From Excel Workbook and select price-list-source.xlsx.

2. In the Navigator select the PriceList table and click Transform Data rather than Load.

3. Promote the first row to headers if needed, then set Unit Price to Decimal Number.

4. Home > Close & Load To… > Only Create Connection. Leave 'Add this data to the Data Model' unticked.

5. Open Queries & Connections and confirm PriceList reports 'Connection only'.

## Test it

The Queries & Connections pane lists PriceList as 'Connection only', and no new worksheet was created.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*