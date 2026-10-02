# -*- coding: utf-8 -*-
"""Generate the relational sample data for the customer-analysis project.

Run:
    python scripts/generate_data.py

Outputs (data/):
    customers.csv    300 customers (with a few data-quality problems)
    products.csv      60 products across 4 categories
    orders.csv     1,600 orders  (2023 - 2025)
    order_items.csv  ~4,000 order lines
"""
from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

SEED = 7
ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"

FIRST = ["Aarav", "Vivaan", "Aditya", "Ananya", "Diya", "Ishaan", "Kabir", "Meera",
         "Riya", "Rohan", "Saanvi", "Aryan", "Kavya", "Nisha", "Rahul", "Pooja",
         "Vikram", "Sneha", "Manish", "Tanya", "Harsh", "Neha", "Yash", "Simran"]
LAST = ["Sharma", "Verma", "Gupta", "Iyer", "Nair", "Patel", "Reddy", "Singh",
        "Mehta", "Joshi", "Kulkarni", "Das", "Bose", "Chauhan", "Mishra", "Rao"]
STATES = {
    "Delhi": ["New Delhi", "Noida", "Gurugram"],
    "Maharashtra": ["Mumbai", "Pune", "Nagpur"],
    "Karnataka": ["Bengaluru", "Mysuru", "Mangaluru"],
    "Tamil Nadu": ["Chennai", "Coimbatore", "Madurai"],
    "West Bengal": ["Kolkata", "Howrah", "Siliguri"],
    "Uttar Pradesh": ["Lucknow", "Kanpur", "Agra"],
}
PRODUCTS = {
    "Electronics": [("Wireless Earbuds", 2499), ("Smartphone", 18999), ("Laptop", 54999),
                    ("Smart Watch", 3999), ("Bluetooth Speaker", 2799)],
    "Furniture": [("Office Chair", 6499), ("Study Table", 8999), ("Bookshelf", 5299),
                ("Sofa Set", 32999), ("Mattress", 12999)],
    "Office Supplies": [("Notebook Pack", 349), ("Gel Pen Set", 199), ("File Binder", 279),
                        ("Stapler", 399), ("Whiteboard", 1899)],
    "Apparel": [("T-Shirt", 699), ("Denim Jeans", 1799), ("Running Shoes", 3299),
                ("Backpack", 1499), ("Winter Jacket", 2799)],
}
SEGMENTS = ["Consumer", "Corporate", "SMB"]
TIERS = ["Bronze", "Silver", "Gold", "Platinum"]
CHANNELS = ["Web", "Mobile App", "Marketplace", "Retail Store"]
STATUSES = ["Delivered", "Shipped", "Pending", "Cancelled", "Returned"]
PAYMENTS = ["Credit Card", "UPI", "Net Banking", "Debit Card", "Cash on Delivery"]


def build_data() -> dict[str, pd.DataFrame]:
    """Create the four relational tables with a reproducible random seed."""
    rng = np.random.default_rng(SEED)

    # ---------------- customers ----------------
    n_customers = 300
    states = list(STATES)
    state = rng.choice(states, size=n_customers)
    cities = [STATES[s][i] for s, i in zip(state, rng.integers(0, 3, n_customers))]
    signup = pd.to_datetime("2022-01-01") + pd.to_timedelta(
        rng.integers(0, 1200, n_customers), unit="D"
    )
    customers = pd.DataFrame({
        "customer_id": [f"CUST{1000 + i}" for i in range(1, n_customers + 1)],
        "first_name": rng.choice(FIRST, n_customers),
        "last_name": rng.choice(LAST, n_customers),
        "city": cities,
        "state": state,
        "country": "India",
        "signup_date": signup.strftime("%Y-%m-%d"),
        "segment": rng.choice(SEGMENTS, n_customers, p=[0.55, 0.27, 0.18]),
        "loyalty_tier": rng.choice(TIERS, n_customers, p=[0.40, 0.30, 0.20, 0.10]),
    })
    customers["email"] = (
        customers["first_name"].str.lower() + "." + customers["last_name"].str.lower()
        + "@example.com"
    )
    customers["phone"] = [
        f"+91 {rng.integers(70, 99)}{rng.integers(10000000, 99999999)}" for _ in range(n_customers)
    ]

    # ---------------- products ----------------
    product_rows = []
    pid = 1
    for category, items in PRODUCTS.items():
        for name, price in items:
            product_rows.append({
                "product_id": f"PRD{pid:03d}",
                "product_name": name,
                "category": category,
                "unit_price": float(price),
            })
            pid += 1
    products = pd.DataFrame(product_rows)

    # ---------------- orders ----------------
    n_orders = 1600
    days = pd.date_range("2023-01-01", "2025-12-31", freq="D")
    month_weight = np.array([0.85, 0.80, 0.95, 0.92, 0.90, 0.88,
                             0.86, 0.92, 1.00, 1.25, 1.35, 1.10])
    year_growth = {2023: 1.00, 2024: 1.10, 2025: 1.24}
    weights = np.array([month_weight[d.month - 1] * year_growth[d.year] for d in days])
    weights = weights / weights.sum()

    order_dates = pd.DatetimeIndex(sorted(days[rng.choice(len(days), n_orders, p=weights)]))
    orders = pd.DataFrame({
        "order_id": [f"ORD{50000 + i}" for i in range(1, n_orders + 1)],
        "customer_id": rng.choice(customers["customer_id"], n_orders),
        "order_date": order_dates.strftime("%Y-%m-%d"),
        "channel": rng.choice(CHANNELS, n_orders, p=[0.38, 0.30, 0.22, 0.10]),
        "payment_method": rng.choice(PAYMENTS, n_orders, p=[0.28, 0.30, 0.14, 0.20, 0.08]),
        "status": rng.choice(STATUSES, n_orders, p=[0.76, 0.09, 0.06, 0.05, 0.04]),
    })

    # ---------------- order_items ----------------
    n_items = 4000
    item_orders = rng.choice(orders["order_id"], n_items)
    order_index = orders.set_index("order_id")
    item_dates = order_index.loc[item_orders, "order_date"].to_numpy()

    chosen = rng.integers(0, len(products), n_items)
    base_price = products["unit_price"].to_numpy()[chosen]
    discount = np.clip(rng.normal(0.07, 0.06, n_items), 0, 0.35).round(2)
    quantity = rng.integers(1, 8, n_items)

    order_items = pd.DataFrame({
        "order_item_id": [f"OI{100000 + i}" for i in range(1, n_items + 1)],
        "order_id": item_orders,
        "product_id": products["product_id"].to_numpy()[chosen],
        "quantity": quantity,
        "unit_price": base_price,
        "discount_pct": discount,
    })
    order_items["line_total"] = (
        order_items["quantity"] * order_items["unit_price"] * (1 - order_items["discount_pct"])
    ).round(2)
    order_items["order_date"] = item_dates
    order_items = order_items.sort_values(["order_date", "order_id"]).reset_index(drop=True)

    return {
        "customers": customers,
        "products": products,
        "orders": orders,
        "order_items": order_items,
    }


def inject_quality_issues(tables: dict[str, pd.DataFrame], rng: np.random.Generator) -> dict:
    """Add the messy bits a real CRM/ERP export usually contains."""
    customers = tables["customers"].copy()
    n = len(customers)

    # inconsistent casing + stray whitespace in segment / city
    picks = rng.choice(n, size=int(n * 0.06), replace=False)
    half = len(picks) // 2
    customers.loc[picks[:half], "segment"] = customers.loc[picks[:half], "segment"].str.upper()
    customers.loc[picks[half:], "segment"] = " " + customers.loc[picks[half:], "segment"].str.lower() + " "
    customers.loc[rng.choice(n, size=int(n * 0.04), replace=False), "city"] = np.nan
    customers.loc[rng.choice(n, size=int(n * 0.03), replace=False), "email"] = np.nan

    # duplicate customer records (same customer_id + email, entered twice)
    dupes = customers.sample(n=int(n * 0.03), random_state=SEED)
    customers = pd.concat([customers, dupes], ignore_index=True)

    tables["customers"] = customers.sample(frac=1.0, random_state=SEED).reset_index(drop=True)
    return tables


def main() -> None:
    rng = np.random.default_rng(SEED + 99)
    tables = inject_quality_issues(build_data(), rng)
    DATA.mkdir(parents=True, exist_ok=True)

    for name, frame in tables.items():
        path = DATA / f"{name}_raw.csv"
        frame.to_csv(path, index=False)
        print(f"Wrote {len(frame):>5,} rows x {frame.shape[1]:>2} cols -> {path.name}")

    customers = tables["customers"]
    print(f"\nInjected issues -> duplicate customers: {int(customers.duplicated(subset=['customer_id']).sum()):,}, "
          f"missing city: {int(customers['city'].isna().sum()):,}, "
          f"missing email: {int(customers['email'].isna().sum()):,}")
    print("Next: python scripts/analysis.py   (clean -> SQLite -> queries -> Excel + report)")


if __name__ == "__main__":
    main()
