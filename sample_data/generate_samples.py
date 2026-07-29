"""Generate three realistic sample files for end-to-end testing.

File A: sales_north.xlsx   — Sales data, 60 rows
File B: sales_south.xlsx   — Sales data with slightly different columns, 40 rows
File C: employees.xlsx     — Completely different schema, 25 rows
"""

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import pandas as pd
import numpy as np
from config.settings import SAMPLE_DATA_DIR

SAMPLE_DATA_DIR.mkdir(parents=True, exist_ok=True)

np.random.seed(42)

# ── File A: North region sales ───────────────────────────
# ~60 rows, compatible schema with File B, non-English city names

products = ["Widget A", "Widget B", "Gadget X", "Gadget Y", "Premium Kit"]
cities_eu = ["München", "Köln", "Paris", "Lyon", "Milano", "Wien"]
regions = ["North", "South", "East", "West"]

north_data = {
    "order_id": range(1001, 1061),
    "product": np.random.choice(products, 60),
    "quantity": np.random.randint(1, 25, 60).astype(float),
    "unit_price": np.round(np.random.uniform(5.0, 200.0, 60), 2),
    "order_date": pd.date_range("2024-01-01", periods=60, freq="D"),
    "region": np.random.choice(regions, 60),
    "customer_city": np.random.choice(cities_eu, 60),
}
df_a = pd.DataFrame(north_data)
df_a["total"] = (df_a["quantity"] * df_a["unit_price"]).round(2)

# Duplicates (2 exact duplicate rows)
df_a.loc[55] = df_a.loc[3].copy()
df_a.loc[56] = df_a.loc[3].copy()

# Missing values
df_a.loc[10, "quantity"] = None
df_a.loc[20, "region"] = None
df_a.loc[30, "customer_city"] = None
df_a.loc[40, "quantity"] = None

# Realistic outlier (extremely high unit price)
df_a.loc[50, "unit_price"] = 9999.99
df_a.loc[50, "total"] = round(df_a.loc[50, "quantity"] * 9999.99, 2)

df_a.to_excel(SAMPLE_DATA_DIR / "sales_north.xlsx", index=False, engine="openpyxl")
print("Created sales_north.xlsx — 60 rows, 8 columns, 2 duplicates, 4 missing, 1 outlier")

# ── File B: South region sales — slightly different schema ─
# ~40 rows, shares core columns with A, has 'discount_code' instead of 'customer_city'

cities_sa = ["São Paulo", "Côte d'Azur", "Buenos Aires", "Lisboa", "Porto"]
discount_codes = ["DISC10", "DISC20", None]

south_data = {
    "order_id": range(2001, 2041),
    "product": np.random.choice(products, 40),
    "quantity": np.random.randint(1, 30, 40).astype(float),
    "unit_price": np.round(np.random.uniform(8.0, 180.0, 40), 2),
    "order_date": pd.date_range("2024-02-01", periods=40, freq="D"),
    "region": np.random.choice(regions, 40),
    "discount_code": np.random.choice(discount_codes, 40),
}
df_b = pd.DataFrame(south_data)
df_b["total"] = (df_b["quantity"] * df_b["unit_price"]).round(2)

# Missing values
df_b.loc[5, "quantity"] = None
df_b.loc[15, "discount_code"] = None
df_b.loc[25, "region"] = None

# A few rows where discount_code is actually valid (not None)
df_b.loc[0, "discount_code"] = "SUMMER24"
df_b.loc[10, "discount_code"] = "SUMMER24"
df_b.loc[20, "discount_code"] = "WELCOME5"

df_b.to_excel(SAMPLE_DATA_DIR / "sales_south.xlsx", index=False, engine="openpyxl")
print("Created sales_south.xlsx — 40 rows, 8 columns, 3 missing, different schema")

# ── File C: Employee data — completely different schema ───
# ~25 rows, no overlap with sales files

depts = ["Engineering", "Sales", "HR", "Marketing", "Finance"]
names = [
    "Anna Schmidt", "Jean Dupont", "Maria Rossi", "John Smith", "Sophie Müller",
    "Carlos García", "Elena Petrova", "Wei Zhang", "Fatima Al-Rashid", "Olga Ivanova",
    "Liam O'Brien", "Yuki Tanaka", "Ahmed Hassan", "Sofia Costa", "Lars Johansson",
    "Mei Lin", "Raj Patel", "Abdullah Al-Saud", "Ingrid Svensson", "Pedro Santos",
    "Nina Kowalski", "Klaus Weber", "Aiko Yamamoto", "Ibrahim Diallo", "Eva Nowak",
]
cities = ["München", "Zürich", "Paris", "London", "Berlin", "Wien", "Madrid"]

employee_data = {
    "employee_id": range(301, 326),
    "name": names,
    "department": np.random.choice(depts, 25),
    "salary": np.random.randint(35000, 130000, 25).astype(float),
    "hire_date": pd.date_range("2019-06-01", periods=25, freq="ME"),
    "office_city": np.random.choice(cities, 25),
}
df_c = pd.DataFrame(employee_data)

# Missing values
df_c.loc[3, "salary"] = None
df_c.loc[8, "salary"] = None
df_c.loc[12, "department"] = None
df_c.loc[20, "office_city"] = None

# Outliers (extremely high salary)
df_c.loc[18, "salary"] = 250000.0

df_c.to_excel(SAMPLE_DATA_DIR / "employees.xlsx", index=False, engine="openpyxl")
print("Created employees.xlsx — 25 rows, 6 columns, 4 missing, 1 outlier (salary)")
print("\nAll sample files generated in:", SAMPLE_DATA_DIR)
