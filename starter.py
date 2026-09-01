import json
from pathlib import Path
import matplotlib.pyplot as plt
import pandas as pd

# กำหนดโฟลเดอร์ Input / Output
# รองรับทั้งการวางไฟล์ไว้ในโฟลเดอร์เดียวกับสคริปต์ หรือในโฟลเดอร์ data/
BASE_DIR = Path(__file__).parent
DATA = BASE_DIR / "data" if (BASE_DIR / "data").exists() else BASE_DIR
OUTPUT = BASE_DIR / "output"
OUTPUT.mkdir(exist_ok=True)

dq_records = []


def log_dq(step, issue, dropped_count, details=""):
    """บันทึกเหตุการณ์การตรวจสอบและคัดกรองข้อมูลลง Data Quality Report"""
    dq_records.append(
        {
            "step": step,
            "issue_type": issue,
            "impacted_rows": dropped_count,
            "details": details,
        }
    )


# ==============================================================================
# TODO 1: Extract ข้อมูลจาก CSV, Excel และ JSON
# ==============================================================================
print("--- 1. Extracting Data ---")
df_orders_jan = pd.read_csv(DATA / "orders_2026_01.csv")
df_orders_feb = pd.read_csv(DATA / "orders_2026_02.csv")
df_customers = pd.read_csv(DATA / "customers_crm.csv")
df_products = pd.read_excel(DATA / "product_master.xlsx")

with open(DATA / "payments.json", "r", encoding="utf-8") as f:
    raw_payments = json.load(f)

# Flatten nested JSON
df_payments = pd.json_normalize(raw_payments)
df_payments.rename(
    columns={
        "payment.method": "payment_method",
        "payment.status": "payment_status",
    },
    inplace=True,
)

print(
    f"Loaded: Jan={df_orders_jan.shape}, Feb={df_orders_feb.shape}, "
    f"Cust={df_customers.shape}, Prod={df_products.shape}, Pay={df_payments.shape}"
)

# ==============================================================================
# TODO 2: Schema Alignment & Concat Orders
# ==============================================================================
print("\n--- 2. Schema Alignment & Concat Orders ---")
# 1. ปรับชื่อคอลัมน์ของเดือน ก.พ. ให้ตรงกับเดือน ม.ค.
feb_cols_map = {
    "ordered_at": "order_date",
    "qty": "quantity",
    "discount_pct": "discount",
}
df_orders_feb_aligned = df_orders_feb.rename(columns=feb_cols_map).copy()

# 2. ทำความสะอาดและแปลง discount จาก string ('5%', '10%') เป็น numeric float (0.05, 0.10)
df_orders_feb_aligned["discount"] = (
    df_orders_feb_aligned["discount"].astype(str).str.rstrip("%").str.strip()
)
df_orders_feb_aligned["discount"] = pd.to_numeric(
    df_orders_feb_aligned["discount"], errors="coerce"
)
if df_orders_feb_aligned["discount"].max() > 1.0:
    df_orders_feb_aligned["discount"] = df_orders_feb_aligned["discount"] / 100.0

df_orders_jan["discount"] = pd.to_numeric(
    df_orders_jan["discount"], errors="coerce"
)
df_orders_jan["quantity"] = pd.to_numeric(
    df_orders_jan["quantity"], errors="coerce"
)
df_orders_jan["unit_price"] = pd.to_numeric(
    df_orders_jan["unit_price"], errors="coerce"
)

df_orders_feb_aligned["quantity"] = pd.to_numeric(
    df_orders_feb_aligned["quantity"], errors="coerce"
)
df_orders_feb_aligned["unit_price"] = pd.to_numeric(
    df_orders_feb_aligned["unit_price"], errors="coerce"
)

# 3. จัดการรูปแบบ Date ให้เป็นมาตรฐานเดียวกัน
df_orders_jan["order_date"] = pd.to_datetime(df_orders_jan["order_date"])
df_orders_feb_aligned["order_date"] = pd.to_datetime(
    df_orders_feb_aligned["order_date"], format="%d/%m/%Y %H:%M"
)

# 4. รวมข้อมูลคำสั่งซื้อทั้งสองเดือน
df_orders_raw = pd.concat(
    [df_orders_jan, df_orders_feb_aligned], ignore_index=True
)
total_raw_orders = len(df_orders_raw)
log_dq(
    "Combine Orders",
    "Initial Concatenation",
    0,
    f"Total raw orders: {total_raw_orders}",
)

# ==============================================================================
# TODO 3: Clean, Standardize, Deduplicate & Data Quality Report
# ==============================================================================
print("\n--- 3. Cleaning & Deduplication ---")

# 3.1 Deduplicate Orders: เก็บแถวล่าสุดตามลำดับที่ปรากฏ (keep='last')
df_orders_dedup = df_orders_raw.drop_duplicates(
    subset=["order_id"], keep="last"
).copy()
dup_orders_dropped = total_raw_orders - len(df_orders_dedup)
log_dq(
    "Orders Deduplication",
    "Duplicate order_id",
    dup_orders_dropped,
    "Retained latest order occurrence",
)

# 3.2 กรองช่วงตัวเลข: quantity > 0, unit_price > 0, discount 0 ถึง 1
valid_num_mask = (
    (df_orders_dedup["quantity"] > 0)
    & (df_orders_dedup["unit_price"] > 0)
    & (df_orders_dedup["discount"] >= 0.0)
    & (df_orders_dedup["discount"] <= 1.0)
)
df_orders_valid = df_orders_dedup[valid_num_mask].copy()
invalid_num_count = len(df_orders_dedup) - len(df_orders_valid)
log_dq(
    "Orders Validation",
    "Invalid quantity/unit_price/discount",
    invalid_num_count,
    "Filtered out non-positive quantity or invalid discount",
)

# 3.3 Clean & Deduplicate Customers
df_customers_clean = df_customers.copy()
df_customers_clean["email"] = (
    df_customers_clean["email"].astype(str).str.strip().str.lower()
)
df_customers_clean["full_name"] = (
    df_customers_clean["full_name"].astype(str).str.strip()
)

# Map ชื่อจังหวัดให้เป็นภาษาไทยมาตรฐาน
province_map = {
    "ชลบุรี": "ชลบุรี",
    "Chonburi": "ชลบุรี",
    "ชลบุรี ": "ชลบุรี",
    "ขอนแก่น": "ขอนแก่น",
    "ขอนเเก่น": "ขอนแก่น",
    "กรุงเทพมหานคร": "กรุงเทพมหานคร",
    "Bangkok": "กรุงเทพมหานคร",
    "กทม.": "กรุงเทพมหานคร",
    "ระยอง": "ระยอง",
    "Rayong": "ระยอง",
    "Phuket": "ภูเก็ต",
    "ภูเก็ต": "ภูเก็ต",
    "Chiang Mai": "เชียงใหม่",
    "เชียงใหม่": "เชียงใหม่",
}
df_customers_clean["province"] = (
    df_customers_clean["province"]
    .astype(str)
    .str.strip()
    .map(lambda x: province_map.get(x, x))
)
raw_cust_len = len(df_customers_clean)
df_dim_customer = df_customers_clean.drop_duplicates(
    subset=["customer_id"], keep="last"
).reset_index(drop=True)
log_dq(
    "Customers Cleaning",
    "Duplicate customer_id",
    raw_cust_len - len(df_dim_customer),
    "Standardized province/email and retained latest customer profile",
)

# 3.4 Clean Products Master
df_dim_product = df_products.drop_duplicates(
    subset=["product_id"], keep="last"
).reset_index(drop=True)
log_dq(
    "Products Cleaning",
    "Duplicate product_id",
    len(df_products) - len(df_dim_product),
    "Retained unique products in Master Catalog",
)

# 3.5 Clean Payments
df_payments_dedup = df_payments.drop_duplicates(
    subset=["order_id"], keep="last"
).reset_index(drop=True)
log_dq(
    "Payments Cleaning",
    "Duplicate order_id payment events",
    len(df_payments) - len(df_payments_dedup),
    "Retained latest payment attempt per order",
)

# ==============================================================================
# TODO 4 & 5: Integrate, Validate Referential Integrity & Business Rules
# ==============================================================================
print("\n--- 4. Integration & Referential Integrity ---")

# Merge Customer Master
m_cust = df_orders_valid.merge(
    df_dim_customer[["customer_id", "full_name", "email", "province"]],
    on="customer_id",
    how="left",
    indicator=True,
)
unmatched_cust = (m_cust["_merge"] == "left_only").sum()
log_dq(
    "Integrity Check",
    "Unmatched customer_id in CRM Master",
    unmatched_cust,
    "Orders with customer_id not found in Master CRM",
)
df_orders_cust = m_cust[m_cust["_merge"] == "both"].drop(columns=["_merge"])

# Merge Product Master
m_prod = df_orders_cust.merge(
    df_dim_product[
        [
            "product_id",
            "product_name",
            "category",
            "standard_price",
            "active_flag",
        ]
    ],
    on="product_id",
    how="left",
    indicator=True,
)
unmatched_prod = (m_prod["_merge"] == "left_only").sum()
log_dq(
    "Integrity Check",
    "Unmatched product_id in Product Master",
    unmatched_prod,
    "Orders with product_id not found in Master Catalog",
)
df_orders_prod = m_prod[m_prod["_merge"] == "both"].drop(columns=["_merge"])

# Merge Payments
m_pay = df_orders_prod.merge(
    df_payments_dedup[
        [
            "order_id",
            "payment_id",
            "payment_method",
            "payment_status",
            "paid_at",
        ]
    ],
    on="order_id",
    how="left",
    indicator=True,
)
unmatched_pay = (m_pay["_merge"] == "left_only").sum()
log_dq(
    "Integrity Check",
    "Unmatched order_id in Payments",
    unmatched_pay,
    "Orders without payment records",
)
df_integrated = m_pay[m_pay["_merge"] == "both"].drop(columns=["_merge"])

# ตรวจสอบสถานะการจ่ายเงิน: นับเฉพาะ PAID เท่านั้น
non_paid_count = (df_integrated["payment_status"] != "PAID").sum()
log_dq(
    "Business Rule Validation",
    "Payment status NOT PAID (FAILED/REFUNDED)",
    non_paid_count,
    "Excluded non-successful transactions from net sales",
)

df_fact_sales = df_integrated[df_integrated["payment_status"] == "PAID"].copy()

# คำนวณ net_sales = quantity × unit_price × (1 - discount)
df_fact_sales["net_sales"] = (
    df_fact_sales["quantity"]
    * df_fact_sales["unit_price"]
    * (1 - df_fact_sales["discount"])
).round(2)

# ==============================================================================
# TODO 6: Load Dimensions, Fact และ Data Quality Report
# ==============================================================================
print("\n--- 5. Loading Star Schema & Reports ---")
df_dim_customer.to_csv(OUTPUT / "dim_customer.csv", index=False)
df_dim_product.to_csv(OUTPUT / "dim_product.csv", index=False)
df_fact_sales.to_csv(OUTPUT / "fact_sales.csv", index=False)

df_dq_report = pd.DataFrame(dq_records)
df_dq_report.to_csv(OUTPUT / "data_quality_report.csv", index=False)

# ==============================================================================
# TODO 7: สร้าง Summary Aggregation Files
# ==============================================================================
print("\n--- 6. Creating Aggregated Summaries ---")
summary_by_province = (
    df_fact_sales.groupby("province")
    .agg(
        total_orders=("order_id", "count"),
        total_quantity=("quantity", "sum"),
        total_net_sales=("net_sales", "sum"),
    )
    .reset_index()
    .sort_values(by="total_net_sales", ascending=False)
)
summary_by_province.to_csv(OUTPUT / "summary_by_province.csv", index=False)

summary_by_category = (
    df_fact_sales.groupby("category")
    .agg(
        total_orders=("order_id", "count"),
        total_quantity=("quantity", "sum"),
        total_net_sales=("net_sales", "sum"),
    )
    .reset_index()
    .sort_values(by="total_net_sales", ascending=False)
)
summary_by_category.to_csv(OUTPUT / "summary_by_category.csv", index=False)

# Optional: กราฟ Data Quality Funnel (Challenge Bonus)
stages = [
    "1. Raw Orders",
    "2. Deduplicated",
    "3. Valid Numbers",
    "4. Matched Master Data",
    "5. Paid Sales (Fact)",
]
counts = [
    total_raw_orders,
    len(df_orders_dedup),
    len(df_orders_valid),
    len(df_integrated),
    len(df_fact_sales),
]

plt.figure(figsize=(9, 4.5))
bars = plt.barh(stages[::-1], counts[::-1], color="#2b5c8f")
plt.title("TechTrove ETL Data Quality Funnel", fontsize=13, pad=12)
plt.xlabel("Record Count")
for bar in bars:
    w = bar.get_width()
    plt.text(
        w + 5,
        bar.get_y() + bar.get_height() / 2,
        f"{int(w):,}",
        ha="left",
        va="center",
        fontweight="bold",
    )
plt.xlim(0, max(counts) * 1.15)
plt.tight_layout()
plt.savefig(OUTPUT / "data_quality_funnel.png", dpi=300)
plt.close()

print(
    "\nData Integration Pipeline Completed! Files saved in 'output/' folder."
)