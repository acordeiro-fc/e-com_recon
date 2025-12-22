import pandas as pd
from config.itsperfect import BASE_URL
from utils.auth import get_itsperfect_token
from utils.pagination import fetch_paginated
from utils.helpers import safe_get

# -----------------------------------
# Mappings
# -----------------------------------
TYPE_MAP = {
    1: "Pre-order",
    2: "Direct order",
    3: "Receipt",
    4: "Sample order",
}

STATUS_MAP = {
    0: "Quantity to ship",
    1: "Sent",
    2: "Canceled",
    3: "Draft",
    4: "Quote",
}

B2B_B2C_MAP = {
    0: "Unknown",
    1: "B2B order",
    2: "B2C order",
}


# -----------------------------------
# Public API
# -----------------------------------
def fetch_sales_orders(date_from: str, date_to: str) -> pd.DataFrame:
    """
    Fetch Itsperfect B2C sales orders (Fab BV),
    including payments and lines.
    """

    headers = {"Authorization": f"Bearer {get_itsperfect_token()}"}

    url = (
        f"{BASE_URL}/sales_orders?"
        "fields=id,date,warehouse,customer,reference,country,"
        "shipping_costs_lcy,shipping_costs_fcy,"
        "discount_lcy,discount_fcy,"
        "subsidiary,type,status,webshop,marketplace_channel,"
        "currency,amount_lcy,amount_fcy,"
        "vat_amount_lcy,vat_amount_fcy,"
        "creation_date,quantity,b2b_b2c_order"
        f"&date>={date_from}&date<{date_to}"
        "&includes=payments,lines"
    )

    raw = fetch_paginated(url, headers)
    df = pd.DataFrame(raw)

    if df.empty:
        return df

    # -----------------------------------
    # Map enums
    # -----------------------------------
    df["type"] = df["type"].map(TYPE_MAP)
    df["status"] = df["status"].map(STATUS_MAP)
    df["b2b_b2c_order"] = df["b2b_b2c_order"].map(B2B_B2C_MAP)

    # -----------------------------------
    # Rename base fields
    # -----------------------------------
    df = df.rename(columns={
        "id": "Order no.",
        "date": "Date",
        "reference": "Reference",
        "shipping_costs_lcy": "Shipping costs",
        "discount_lcy": "Discount",
        "type": "Type",
        "status": "Status",
        "b2b_b2c_order": "Channel",
        "amount_lcy": "Amount",
        "vat_amount_lcy": "VAT value",
        "creation_date": "Creation date",
    })

    # -----------------------------------
    # Extract nested objects
    # -----------------------------------
    df["Warehouse"] = df["warehouse"].apply(lambda x: safe_get(x, "warehouse"))
    df["Customer ID"] = df["customer"].apply(lambda x: safe_get(x, "id"))
    df["Customer"] = df["customer"].apply(lambda x: safe_get(x, "customer_name"))
    df["Country"] = df["country"].apply(lambda x: safe_get(x, "iso2"))
    df["Subsidiary"] = df["subsidiary"].apply(lambda x: safe_get(x, "subsidiary"))
    df["Webshop"] = df["webshop"].apply(lambda x: safe_get(x, "webshop"))
    df["Currency"] = df["currency"].apply(lambda x: safe_get(x, "iso"))

    # -----------------------------------
    # Filters (B2C / Fab BV / no marketplace)
    # -----------------------------------
    df = df[df["Subsidiary"] == "Fab BV"]
    df = df[df["Channel"] == "B2C order"]
    df = df[
        df["marketplace_channel"].apply(
            lambda x: safe_get(x, "channel") is None
        )
    ]

    # -----------------------------------
    # Numeric coercion
    # -----------------------------------
    numeric_cols = [
        "Amount",
        "Shipping costs",
        "VAT value",
        "Discount",
        "amount_fcy",
        "discount_fcy",
        "shipping_costs_fcy",
        "vat_amount_fcy",
    ]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # -----------------------------------
    # Calculated amounts (FCY)
    # -----------------------------------
    df["Subtotaal excl VAT"] = df["amount_fcy"] + df["discount_fcy"]
    df["Total incl. VAT"] = (
        df["amount_fcy"]
        + df["shipping_costs_fcy"]
        + df["vat_amount_fcy"]
    )

    # -----------------------------------
    # Quantities from lines
    # -----------------------------------
    df["Total Qty"] = df["lines"].apply(
        lambda lines: sum(l.get("quantity", 0) for l in lines)
        if isinstance(lines, list)
        else 0
    )

    # -----------------------------------
    # Payments
    # -----------------------------------
    df["Payment date"] = df["payments"].apply(
        lambda payments: min(
            p.get("date") for p in payments if p.get("date")
        )
        if isinstance(payments, list) and payments
        else None
    )

    df["Payment method"] = df["payments"].apply(
        lambda payments: payments[0]
        .get("payment_method", {})
        .get("payment_method")
        if isinstance(payments, list) and payments
        else None
    )

    df["Payment amount (LCY)"] = df.apply(
        lambda row: sum(
            float(p.get("amount_rcy", 0))
            for p in row["payments"]
        ) * row["Total Qty"]
        if isinstance(row["payments"], list)
        else 0,
        axis=1,
    )

    # -----------------------------------
    # Final column order
    # -----------------------------------
    columns = [
        "Order no.",
        "Date",
        "Warehouse",
        "Customer ID",
        "Customer",
        "Reference",
        "Country",
        "Shipping costs",
        "Discount",
        "Subsidiary",
        "Type",
        "Status",
        "Webshop",
        "Channel",
        "Currency",
        "Amount",
        "VAT value",
        "Creation date",
        "Payment date",
        "Payment amount (LCY)",
        "Payment method",
        "Total Qty",
        "Subtotaal excl VAT",
        "Total incl. VAT",
    ]

    df = df[columns]

    return df