# Customer Data Analysis Using SQL & Excel

An end-to-end **SQL + Excel analytics project**: a messy CRM export is cleaned, loaded
into a relational model, interrogated with **15 analysis queries** (joins, GROUP BY /
HAVING, subqueries, CASE pivots, window functions) and packaged into an
**Excel workbook** with VLOOKUP/XLOOKUP lookups, pivot-style summaries and conditional
formatting.

**Live SQL report: https://gstomer001.github.io/customer-data-analysis-sql-excel/**

**Author:** Gopal Singh Tomer · Data Analyst (SQL · MySQL · Excel · Power BI · Python)

---

## 1. What this project demonstrates

| Skill | Where to see it |
| --- | --- |
| Relational data modelling | `sql/01_schema_mysql.sql` (4 tables, PK/FK, indexes) |
| Advanced SQL | `sql/02_analysis_queries.sql` — JOINs, GROUP BY + HAVING, scalar & correlated subqueries, CASE WHEN pivoting, `ROW_NUMBER() OVER (PARTITION BY ...)`, running totals |
| Data cleaning | `scripts/analysis.py` → `clean_tables()` (duplicates, blank imputation, label standardisation, referential-integrity check) |
| Excel transformation | `excel/customer_analysis.xlsx` — VLOOKUP + XLOOKUP enrichment sheet, SUMIF/COUNTIF/AVERAGEIF summary sheets, colour-scale and rule-based conditional formatting |
| Reporting | `index.html` — an interactive report that renders every query with its live result |

## 2. Repository structure

```
customer-data-analysis-sql-excel/
├── index.html                    # live SQL report (GitHub Pages entry point)
├── assets/
│   ├── app.js                    # charts + SQL explorer
│   ├── style.css
│   ├── chart.umd.min.js          # vendored Chart.js 4
│   └── report_data.js            # auto-generated query results (window.REPORT_DATA)
├── sql/
│   ├── 01_schema_mysql.sql       # MySQL DDL, LOAD DATA statements, sanity checks
│   └── 02_analysis_queries.sql   # 15 portable analysis queries
├── scripts/
│   ├── generate_data.py          # reproducible CRM extract (with injected dirt)
│   └── analysis.py               # clean -> SQLite -> run queries -> Excel -> report
├── data/
│   ├── *_raw.csv                 # messy source extracts
│   ├── *.csv                     # cleaned tables
│   ├── customer_analysis.sqlite  # queryable database
│   └── quality_report.json
└── excel/
    └── customer_analysis.xlsx    # Excel deliverable
```

## 3. How to run it locally

```bash
python -m pip install pandas numpy openpyxl

python scripts/generate_data.py   # 1) create the raw CRM extracts
python scripts/analysis.py        # 2) clean, load SQLite, run 15 queries, build Excel + report

# open index.html, or serve it:  python -m http.server 8000
```

### Running the same queries on MySQL

```bash
mysql -u root -p < sql/01_schema_mysql.sql   # schema + loaders + sample inserts
mysql -u root -p customer_analysis < sql/02_analysis_queries.sql
```

`sql/02_analysis_queries.sql` is written in a **portable dialect**: it uses no
`DATE_FORMAT`, `CONCAT` or `||`, and the year/month grain is stored as real
`order_year` / `order_month` columns, so the exact same file runs on MySQL 8 and SQLite.

## 4. Data cleaning performed

| Problem in the raw CRM export | Treatment | Rows affected |
| --- | --- | --- |
| Duplicate customer records | keep first by `customer_id` | **9** removed |
| Inconsistent labels (`CORPORATE`, `" smb "`, `" web "`) | trim + canonical mapping (`SMB` ≠ `Smb`) | **17** corrected |
| Missing `city` | imputed as `Unknown` and reported separately | 12 filled |
| Missing `email` | imputed as `not-provided@example.com` | 9 filled |
| Dates stored as text | `pd.to_datetime` with coercion + derived `order_year` / `order_month` | all rows |
| `line_total` inconsistent with quantity × price × discount | recalculated from first principles | all rows |
| Order lines with no matching order (FK violation) | dropped after the integrity check | 0 |

Result: **300 clean customers, 20 products, 1,600 orders, 4,000 order lines.**

## 5. Query coverage (15 queries)

| Query key | Technique demonstrated |
| --- | --- |
| `kpi_overview` | scalar subqueries + multi-table aggregate |
| `revenue_by_segment` | INNER JOIN + GROUP BY |
| `revenue_by_state` | share-of-total via a scalar subquery |
| `revenue_by_category` | JOIN across fact → dimension, `AVG` |
| `monthly_revenue_trend` | `SUM(...) OVER (ORDER BY ...)` running total |
| `channel_year_crosstab` | CASE WHEN pivot |
| `top_10_customers` | 3-table JOIN + `ORDER BY` + `LIMIT` |
| `customer_value_tiers` | CASE WHEN bucketing over a derived table |
| `high_value_customers` | `GROUP BY ... HAVING SUM(...) > threshold` |
| `above_average_customers` | nested subqueries (`IN` + `AVG`) |
| `customers_without_orders` | `LEFT JOIN ... WHERE ... IS NULL` |
| `repeat_vs_onetime` | CASE bucketing on a derived aggregate |
| `status_by_channel` | conditional aggregation (`SUM(CASE WHEN ...)`) |
| `top_customer_per_state` | `ROW_NUMBER() OVER (PARTITION BY state ...)` |
| `yoy_by_channel` | year-over-year growth with `NULLIF` divide-by-zero guard |

## 6. Key findings

| KPI | Value |
| --- | --- |
| Customers | **300** (299 have purchased at least once) |
| Orders analysed | 1,469 orders that contain line items |
| Total revenue | **Rs 11.79 Cr** |
| Units sold | 16,085 |
| Average order value | **Rs 80,281** |
| Top customer | CUST1220 (Chennai · Corporate) — Rs 15.31 Lakh lifetime |
| Peak month | October 2025 — Rs 55.71 Lakh |

1. **Consumers drive the business** — the Consumer segment is Rs 6.87 Cr (58% of revenue)
   from 886 orders, ahead of Corporate (Rs 3.33 Cr) and SMB (Rs 1.59 Cr).
2. **Revenue is geographically spread** — Uttar Pradesh (Rs 2.43 Cr), Tamil Nadu
   (Rs 2.39 Cr) and Delhi (Rs 2.17 Cr) are the top three states, and no single state
   exceeds ~21% of revenue, which limits concentration risk.
3. **Electronics and Furniture dominate** — Rs 5.85 Cr and Rs 4.95 Cr respectively, while
   Office Supplies is a rounding error at Rs 24.8 Lakh.
4. **Digital channels lead** — Web (Rs 5.01 Cr) plus Mobile App (Rs 3.12 Cr) together
   contribute 69% of revenue; Retail Store is only Rs 1.23 Cr.
5. **A small, high-value core matters** — 219 customers fall in the Platinum tier
   (≥ Rs 2 Lakh lifetime) and they account for Rs 10.98 Cr, i.e. ~93% of revenue.
   Retention of that cohort is the single highest-leverage action.
6. **Discounting is disciplined** — the average discount sits at roughly 7.1–7.6% across
   every category, so no category is being bought purely with price.

## 7. Excel workbook details

`excel/customer_analysis.xlsx` is regenerated on every pipeline run:

- **Orders_Enriched** — adds four lookup columns to the Orders sheet:
  `VLOOKUP` for customer name, segment and city/state, plus an `XLOOKUP` for loyalty tier.
- **Pivot_Segment** — customers, orders, revenue, average order value and revenue share,
  driven entirely by `COUNTIF`, `COUNTIFS`, `SUMIF` and `IFERROR` formulas.
- **Pivot_State** — revenue per state with share-of-total formulas and a rule-based flag
  column for markets above Rs 10 Lakh.
- **Query_Results** — the output of the eight headline SQL queries, dropped into one
  sheet and ready to pivot.
- **Conditional formatting** — colour scales on the revenue columns and a rule-based fill
  on the "Top market" flag.

> The XLOOKUP columns need Excel 365 / 2021; the VLOOKUP columns work in every version.

