# Excel AI Analyzer — Demo / Sample Workbooks

Fake data by construction (safe to ship publicly). Each file is chosen to show
ONE capability that impresses a first-time / trial user — and that a naive or
cloud-based tool typically gets wrong. Regenerate anytime with `make_demos.py`.

## The files and their "wow" moment

### 1_Portfolio_Dashboard.xlsx  — MULTI-TABLE ON ONE SHEET
Three tables on a single sheet: a KPI strip, a holdings table, and a
transactions log side-by-side. Most tools flatten this into a scrambled grid.
This app splits them into three clean, separately-queryable tables.
  Try: "What's my total portfolio value and return?"
       "Which holding has the largest weight?"
       "List all my SELL transactions."

### 2_Financial_Statements.xlsx  — SMART ABOUT STATEMENTS
Income Statement / Balance Sheet / Cash Flow with real hierarchy (subtotals +
totals). The app KNOWS not to sum a column that already contains subtotals — it
reads the specific labelled row instead. Naive tools return nonsense totals.
  Try: "What was Net Income in 2025?"
       "What were Total Assets each year?"
       "Compute the gross margin for 2024."

### 3_Sales_by_Region.xlsx  — EXACT STATS OVER ALL ROWS
120 orders. Aggregate questions compute over EVERY row (exact sum/avg/median),
not a truncated sample — accurate even on large sheets.
  Try: "What's the total revenue across all 120 orders?"
       "Which region has the highest average order value?"
       "Who is the top sales rep by revenue?"

### 4_Multi_Fund_Comparison.xlsx  — ASKS THE RIGHT SHEET
Four fund holding sheets. Ask about one fund and the app routes to that sheet —
no need to hunt through tabs.
  Try: "What's in the Healthcare Fund?"
       "Which fund has the best 1-year return on its top holding?"
       "Compare the top holding of Tech Growth Fund and Dividend Fund."

## Usage
- Landing page: offer these as one-click "try a sample" downloads.
- First-run: bundle 1 + 2 as built-in samples with the suggested questions
  pre-populated, so the trial user gets an instant "it works" moment.
- All queries run 100% locally — the sample data never leaves the machine.

## Regenerate
    python make_demos.py      # writes files/ + prints suggested questions
