# -*- coding: utf-8 -*-
"""Clean the CRM extracts, load them into SQLite, run every query in
sql/02_analysis_queries.sql and export the results for the HTML report.

The same sql/02_analysis_queries.sql file is written in a portable dialect, so
it runs unchanged on MySQL 8 (see sql/01_schema_mysql.sql) and on SQLite here.

Run:
    python scripts/generate_data.py    # creates data/*_raw.csv
    python scripts/analysis.py         # clean -> SQLite -> queries -> Excel + report

Outputs:
    data/customers.csv, products.csv, orders.csv, order_items.csv   (cleaned)
    data/customer_analysis.sqlite                                   SQLite database
    excel/customer_analysis.xlsx                                    Excel deliverable
    assets/report_data.js                                           window.REPORT_DATA
"""
from __future__ import annotations

import json
import re
import sqlite3
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
EXCEL_DIR = ROOT / "excel"
ASSETS = ROOT / "assets"
QUERY_FILE = ROOT / "sql" / "02_analysis_queries.sql"
SQLITE_DB = DATA / "customer_analysis.sqlite"

TEXT_COLS = {
    "customers": ["first_name", "last_name", "city", "state", "country", "segment", "loyalty_tier"],
    "products": ["product_name", "category"],
    "orders": ["channel", "payment_method", "status"],
}

# Canonical spellings used to standardise inconsistent labels found in the raw files.
# (Simple Title Casing cannot be used everywhere: it would turn "SMB" into "Smb".)
CANONICAL: dict[str, list[str]] = {
    "segment": ["Consumer", "Corporate", "SMB"],
    "loyalty_tier": ["Bronze", "Silver", "Gold", "Platinum"],
    "channel": ["Web", "Mobile App", "Marketplace", "Retail Store"],
    "payment_method": ["Credit Card", "UPI", "Net Banking", "Debit Card", "Cash on Delivery"],
    "status": ["Delivered", "Shipped", "Pending", "Cancelled", "Returned"],
    "category": ["Electronics", "Furniture", "Office Supplies", "Apparel"],
}

SCHEMA_SQLITE = """
CREATE TABLE customers (
    customer_id   TEXT PRIMARY KEY,
    first_name    TEXT NOT NULL,
    last_name     TEXT NOT NULL,
    email         TEXT,
    phone         TEXT,
    city          TEXT,
    state         TEXT,
    country       TEXT,
    signup_date   TEXT NOT NULL,
    segment       TEXT NOT NULL,
    loyalty_tier  TEXT NOT NULL
);
CREATE TABLE products (
    product_id    TEXT PRIMARY KEY,
    product_name  TEXT NOT NULL,
    category      TEXT NOT NULL,
    unit_price    REAL NOT NULL
);
CREATE TABLE orders (
    order_id        TEXT PRIMARY KEY,
    customer_id     TEXT NOT NULL,
    order_date      TEXT NOT NULL,
    order_year      INTEGER NOT NULL,
    order_month     TEXT NOT NULL,
    channel         TEXT NOT NULL,
    payment_method  TEXT NOT NULL,
    status          TEXT NOT NULL
);
CREATE TABLE order_items (
    order_item_id  TEXT PRIMARY KEY,
    order_id       TEXT NOT NULL,
    product_id     TEXT NOT NULL,
    quantity       INTEGER NOT NULL,
    unit_price     REAL NOT NULL,
    discount_pct   REAL NOT NULL DEFAULT 0,
    line_total     REAL NOT NULL
);
CREATE INDEX idx_orders_customer ON orders (customer_id);
CREATE INDEX idx_orders_month    ON orders (order_month);
CREATE INDEX idx_items_order     ON order_items (order_id);
"""


def clean_tables() -> tuple[dict[str, pd.DataFrame], dict]:
    """Read the raw CSVs, fix the quality problems and return clean tables."""
    report: dict = {}
    tables = {name: pd.read_csv(DATA / f"{name}_raw.csv")
              for name in ["customers", "products", "orders", "order_items"]}

    report["raw_rows"] = {name: int(len(df)) for name, df in tables.items()}

    # ---- 1) standardise text: trim whitespace, then map to canonical labels ----
    changed = 0
    for name, frame in tables.items():
        for col in TEXT_COLS.get(name, []):
            stripped = frame[col].astype("string").str.strip()
            if col in CANONICAL:
                canonical = {
                    value.upper(): value for value in CANONICAL[col]
                }
                keys = stripped.str.upper()
                frame[col] = keys.map(canonical).fillna(stripped).astype("string")
            elif col in {"city", "state", "country"}:
                frame[col] = stripped.str.title()
            else:
                frame[col] = stripped
            changed += int((stripped != frame[col]).sum())
    report["labels_standardised"] = changed

    # ---- 2) customers: de-duplicate, impute blanks, parse dates ----
    customers = tables["customers"]
    report["duplicate_customers_removed"] = int(
        customers.duplicated(subset=["customer_id"]).sum())
    customers = customers.drop_duplicates(subset=["customer_id"], keep="first").copy()

    report["missing_values"] = {
        "city": int(customers["city"].isna().sum()),
        "email": int(customers["email"].isna().sum()),
    }
    for col in ["city", "state", "country"]:
        customers[col] = customers[col].fillna("Unknown")
    customers["email"] = customers["email"].fillna("not-provided@example.com")
    customers["signup_date"] = pd.to_datetime(customers["signup_date"], errors="coerce")
    customers = customers.dropna(subset=["signup_date"]).copy()
    customers["signup_date"] = customers["signup_date"].dt.strftime("%Y-%m-%d")

    # ---- 3) orders: parse dates + derive the portable date dimensions ----
    orders = tables["orders"].copy()
    orders["order_date"] = pd.to_datetime(orders["order_date"], errors="coerce")
    orders = orders.dropna(subset=["order_date"]).copy()
    orders["order_year"] = orders["order_date"].dt.year.astype(int)
    orders["order_month"] = orders["order_date"].dt.to_period("M").astype(str)
    orders["order_date"] = orders["order_date"].dt.strftime("%Y-%m-%d")

    # ---- 4) order_items: recalc line_total + drop orphan lines ----
    items = tables["order_items"].copy()
    items["line_total"] = (items["quantity"] * items["unit_price"]
                           * (1 - items["discount_pct"])).round(2)
    # order_date is already on the orders table - keep order_items normalised
    items = items.drop(columns=["order_date"], errors="ignore")
    valid_orders = set(orders["order_id"])
    before = len(items)
    items = items[items["order_id"].isin(valid_orders)].copy()
    report["orphan_order_items_removed"] = int(before - len(items))

    clean = {
        "customers": customers.reset_index(drop=True),
        "products": tables["products"].reset_index(drop=True),
        "orders": orders.reset_index(drop=True),
        "order_items": items.reset_index(drop=True),
    }
    report["clean_rows"] = {name: int(len(df)) for name, df in clean.items()}
    return clean, report


def build_sqlite(tables: dict[str, pd.DataFrame]) -> None:
    """(Re)create the SQLite database from the cleaned tables."""
    DATA.mkdir(parents=True, exist_ok=True)
    if SQLITE_DB.exists():
        SQLITE_DB.unlink()

    for name, frame in tables.items():
        frame.to_csv(DATA / f"{name}.csv", index=False)

    conn = sqlite3.connect(SQLITE_DB)
    conn.executescript(SCHEMA_SQLITE)
    for name, frame in tables.items():
        frame.to_sql(name, conn, if_exists="append", index=False)
    conn.commit()
    conn.close()


def load_queries() -> list[dict]:
    """Split sql/02_analysis_queries.sql into individual, runnable queries."""
    text = QUERY_FILE.read_text(encoding="utf-8")
    marker = re.compile(r"^--\s*@query:\s*(?P<key>[\w_]+)\s*\|\s*(?P<title>.+)$", re.MULTILINE)
    matches = list(marker.finditer(text))

    queries = []
    for i, match in enumerate(matches):
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        body = text[match.end():end].strip()
        body = "\n".join(line for line in body.splitlines() if not line.strip().startswith("--"))
        sql = body.strip().rstrip(";").strip()
        if sql:
            queries.append({"key": match.group("key"), "title": match.group("title").strip(), "sql": sql})
    return queries


def run_queries(queries: list[dict]) -> dict[str, list[dict]]:
    """Execute every query against SQLite and return {key: [rows]}."""
    conn = sqlite3.connect(SQLITE_DB)
    conn.row_factory = sqlite3.Row
    results: dict[str, list[dict]] = {}
    for query in queries:
        try:
            cursor = conn.execute(query["sql"])
            rows = [{k: (round(float(v), 2) if isinstance(v, float) else v) for k, v in dict(r).items()}
                    for r in cursor.fetchall()]
            results[query["key"]] = rows
            print(f"  ok   {query['key']:<28} {len(rows):>4} rows")
        except sqlite3.Error as exc:  # noqa: PERF203
            print(f"  FAIL {query['key']:<28} {exc}")
            results[query["key"]] = []
    conn.close()
    return results


HEADER_FILL = PatternFill("solid", fgColor="1F6FEB")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def _write_sheet(ws, frame: pd.DataFrame, freeze: str = "A2") -> None:
    """Write a dataframe as a formatted table with a bold header row."""
    ws.append(list(frame.columns))
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for record in frame.itertuples(index=False):
        ws.append(["None" if pd.isna(v) else v for v in record])
    ws.freeze_panes = freeze
    for idx, column in enumerate(frame.columns, start=1):
        width = max(len(str(column)) + 2,
                    *(len(str(v)) + 2 for v in frame[column].head(300).astype(str)) if len(frame) else 3)
        ws.column_dimensions[get_column_letter(idx)].width = min(width, 32)


def build_excel(tables: dict[str, pd.DataFrame], results: dict[str, list[dict]]) -> Path:
    """Create the Excel deliverable: data sheets + VLOOKUP/XLOOKUP + pivot summaries."""
    EXCEL_DIR.mkdir(parents=True, exist_ok=True)
    out = EXCEL_DIR / "customer_analysis.xlsx"

    customers, products = tables["customers"], tables["products"]
    orders, items = tables["orders"], tables["order_items"]
    queries = load_queries()

    wb = Workbook()
    ws = wb.active
    ws.title = "Read_me"
    notes = [
        ["Customer Data Analysis - Excel workbook"],
        ["Generated by scripts/analysis.py (project: customer-data-analysis-sql-excel)"],
        [],
        ["Sheet", "What it contains"],
        ["Customers", "Cleaned customer master (duplicates removed, blanks imputed)"],
        ["Products", "Product catalogue with categories and unit prices"],
        ["Orders", "Order headers incl. derived order_year / order_month columns"],
        ["Order_Items", "Order line detail with recalculated line_total"],
        ["Orders_Enriched", "Orders + VLOOKUP/XLOOKUP lookups pulling customer attributes"],
        ["Pivot_Segment", "Pivot-style summary by segment (SUMIFS / COUNTIFS / AVERAGEIFS)"],
        ["Pivot_State", "Pivot-style summary by state with share of revenue"],
        ["Query_Results", "Output of the 15 SQL queries in sql/02_analysis_queries.sql"],
        [],
        ["Note", "The XLOOKUP columns need Excel 365 / 2021; VLOOKUP columns work in every version."],
    ]
    for row in notes:
        ws.append(row)
    ws["A1"].font = Font(bold=True, size=14)
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 78

    _write_sheet(wb.create_sheet("Customers"), customers)
    _write_sheet(wb.create_sheet("Products"), products)
    _write_sheet(wb.create_sheet("Orders"), orders)
    _write_sheet(wb.create_sheet("Order_Items"), items)

    # ---------------- Orders_Enriched: VLOOKUP / XLOOKUP lookups ----------------
    revenue_by_order = items.groupby("order_id", as_index=False)["line_total"].sum()
    revenue_map = dict(zip(revenue_by_order["order_id"], revenue_by_order["line_total"]))

    ws = wb.create_sheet("Orders_Enriched")
    headers = ["order_id", "customer_id", "order_date", "order_year", "order_month",
               "channel", "payment_method", "Customer_Name", "Segment", "City",
               "State", "Loyalty_Tier", "Order_Revenue"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row_idx, order in enumerate(orders.itertuples(index=False), start=2):
        ws.append([
            order.order_id, order.customer_id, order.order_date, int(order.order_year),
            order.order_month, order.channel, order.payment_method,
            f'=VLOOKUP($B{row_idx},Customers!$A:$K,2,FALSE)&" "&VLOOKUP($B{row_idx},Customers!$A:$K,3,FALSE)',
            f'=VLOOKUP($B{row_idx},Customers!$A:$K,10,FALSE)',
            f'=VLOOKUP($B{row_idx},Customers!$A:$K,6,FALSE)',
            f'=VLOOKUP($B{row_idx},Customers!$A:$K,7,FALSE)',
            f'=XLOOKUP($B{row_idx},Customers!$A:$A,Customers!$K:$K,"Not found")',
            round(float(revenue_map.get(order.order_id, 0.0)), 2),
        ])
    ws.freeze_panes = "A2"
    for idx in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(idx)].width = 16
    ws.column_dimensions["H"].width = 24

    # ---------------- Pivot_Segment (SUMIFS / COUNTIFS) ----------------
    ws = wb.create_sheet("Pivot_Segment")
    ws.append(["Segment", "Customers", "Orders", "Revenue", "Avg Order Value", "Revenue Share %"])
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    for row_idx, segment in enumerate(["Consumer", "Corporate", "SMB"], start=2):
        ws.append([
            segment,
            f'=COUNTIF(Customers!$J:$J,$A{row_idx})',
            f'=COUNTIF(Orders_Enriched!$I:$I,$A{row_idx})',
            f'=SUMIF(Orders_Enriched!$I:$I,$A{row_idx},Orders_Enriched!$M:$M)',
            f'=IFERROR(D{row_idx}/C{row_idx},0)',
            f'=IFERROR(D{row_idx}/SUM($D$2:$D$4),0)',
        ])
    ws.append(["TOTAL", "=SUM(B2:B4)", "=SUM(C2:C4)", "=SUM(D2:D4)", "=IFERROR(D5/C5,0)", "=SUM(F2:F4)"])
    ws["A5"].font = Font(bold=True)
    ws["E2"].number_format = "#,##0"
    ws["F2"].number_format = "0.0%"
    ws["F3"].number_format = "0.0%"
    ws["F4"].number_format = "0.0%"
    ws["F5"].number_format = "0.0%"
    ws.conditional_formatting.add("D2:D4", ColorScaleRule(start_type="min", start_color="FFF5F5",
                                                          end_type="max", end_color="2EA043"))

    # ---------------- Pivot_State (SUMIF / COUNTIF) ----------------
    states = sorted(customers["state"].dropna().unique().tolist())
    ws = wb.create_sheet("Pivot_State")
    ws.append(["State", "Orders", "Revenue", "Revenue Share %", "Flag"])
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    last = len(states) + 1
    for row_idx, state in enumerate(states, start=2):
        ws.append([
            state,
            f'=COUNTIF(Orders_Enriched!$K:$K,$A{row_idx})',
            f'=SUMIF(Orders_Enriched!$K:$K,$A{row_idx},Orders_Enriched!$M:$M)',
            f'=IFERROR(C{row_idx}/SUM($C$2:$C${last}),0)',
            f'=IF(C{row_idx}>1000000,"Top market","")',
        ])
    ws.append(["TOTAL", f"=SUM(B2:B{last})", f"=SUM(C2:C{last})", f"=SUM(D2:D{last})", ""])
    ws[f"A{last + 1}"].font = Font(bold=True)
    for i in range(2, last + 1):
        ws[f"D{i}"].number_format = "0.0%"
    ws.conditional_formatting.add(f"C2:C{last}", ColorScaleRule(start_type="min", start_color="FFF5F5",
                                                               end_type="max", end_color="1F6FEB"))
    ws.conditional_formatting.add(f"E2:E{last}",
                                  CellIsRule(operator="equal", formula=['"Top market"'],
                                             fill=PatternFill("solid", fgColor="D2F5D0")))
    for idx in range(1, 6):
        ws.column_dimensions[get_column_letter(idx)].width = 18
    # ---------------- Query_Results: outputs of the SQL analysis ----------------
    ws = wb.create_sheet("Query_Results")
    ws.append(["SQL Analysis Results - sql/02_analysis_queries.sql"])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append([])

    highlights = ["kpi_overview", "revenue_by_segment", "revenue_by_state",
                  "revenue_by_category", "top_10_customers", "customer_value_tiers",
                  "channel_year_crosstab", "top_customer_per_state"]
    row = 3
    for key in highlights:
        rows = results.get(key) or []
        title = next((q["title"] for q in queries if q["key"] == key), key)
        ws.cell(row=row, column=1, value=f"{title}  [{key}]").font = Font(bold=True, color="1F6FEB")
        row += 1
        if rows:
            headers = list(rows[0].keys())
            for col, head in enumerate(headers, start=1):
                cell = ws.cell(row=row, column=col, value=head)
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
            row += 1
            for record in rows:
                for col, head in enumerate(headers, start=1):
                    ws.cell(row=row, column=col, value=record.get(head))
                row += 1
        else:
            ws.cell(row=row, column=1, value="(no rows)")
            row += 1
        row += 1

    for idx in range(1, 12):
        ws.column_dimensions[get_column_letter(idx)].width = min(24, max(14, idx * 3))
    ws.column_dimensions["A"].width = 24

    return out
    return out


def main() -> None:
    print("1/5  cleaning raw CRM extracts ...")
    tables, report = clean_tables()

    print("2/5  loading SQLite database ...")
    build_sqlite(tables)

    print("3/5  running SQL queries ...")
    queries = load_queries()
    results = run_queries(queries)
    print(f"     {len(queries)} queries executed from {QUERY_FILE.name}")

    print("4/5  building the Excel workbook ...")
    xlsx = build_excel(tables, results)
    print(f"     -> {xlsx}")

    print("5/5  exporting report data ...")
    ASSETS.mkdir(parents=True, exist_ok=True)
    payload = {
        "queries": [{"key": q["key"], "title": q["title"], "sql": q["sql"]} for q in queries],
        "results": results,
        "quality": report,
        "counts": {name: int(len(df)) for name, df in tables.items()},
    }
    (ASSETS / "report_data.js").write_text(
        "/* Auto-generated by scripts/analysis.py - do not edit by hand. */\n"
        f"window.REPORT_DATA = {json.dumps(payload, indent=2)};\n",
        encoding="utf-8",
    )
    (DATA / "quality_report.json").write_text(json.dumps(report, indent=2), encoding="utf-8")

    print(f"\nRows   : {report['raw_rows']} (raw) -> {report['clean_rows']} (clean)")
    print(f"Quality: {report['duplicate_customers_removed']} duplicate customers removed, "
          f"{report['labels_standardised']} labels standardised, "
          f"{report['orphan_order_items_removed']} orphan order lines dropped")


if __name__ == "__main__":
    main()
