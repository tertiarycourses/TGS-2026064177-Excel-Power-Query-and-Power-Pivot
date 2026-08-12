# Lab 19 — Fuzzy Matching and Transformation Tables

**Course:** Excel Power Query and Power Pivot (TGS-2026064177)  
**Topic 3:** Advanced Data Analysis using Power Query  
**Maps to:** LO3 · A3 · K2, K3 — Match inconsistent text with fuzzy matching and a transformation table.

**Workbook:** `lab-19-fuzzy-matching.xlsx`

## Scenario

Customer names arrive spelled inconsistently across two systems — 'Pte Ltd' versus 'Pte. Ltd.', typos, and abbreviations. The learner enables fuzzy matching, tunes the similarity threshold, and then supplies a transformation table to teach the engine the pairs that no similarity score will ever catch.

## What you will build

A reconciled customer query matched by fuzzy logic plus a transformation table.

**Tools used:** Excel, Power Query Editor, Fuzzy Merge, Transformation Table

## Steps

1. Merge CRMCustomers with FinanceCustomers on Customer Name using Left Outer and note how many rows fail to match.

2. Edit the merge, tick 'Use fuzzy matching to perform the merge', and open Fuzzy matching options.

3. Set the Similarity threshold to 0.8 and tick 'Ignore case' and 'Ignore space'. Record the improvement.

4. Lower the threshold to 0.6 and observe false matches appearing — the trade-off between recall and precision.

5. Return to 0.8 and set the TransformTable query as the Transformation table so 'ABC Holdings' maps to 'ABC Group' explicitly.

## Test it

All customers reconcile at threshold 0.8 with the transformation table applied, and no false match is introduced.

---

*Tertiary Infotech Academy Pte Ltd · UEN: 201200696W · v5.0*