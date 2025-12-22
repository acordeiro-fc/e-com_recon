import pandas as pd
from utils.auth import get_itsperfect_token
from utils.pagination import fetch_paginated
from utils.helpers import safe_get
import streamlit as st

BASE_URL = st.secrets["ITSP_BASE_URL"]
def fetch_returns(date_from, date_to):
    headers = {"Authorization": f"Bearer {get_itsperfect_token()}"}

    url = (
        f"{BASE_URL}/sales_return_orders?"
        f"fields=id,date,warehouse,customer,return_costs_lcy,discount_lcy,"
        f"remarks,country,subsidiary,quantity,amount_lcy,postage_costs_lcy,"
        f"marketplace_channel,b2b_b2c_order"
        f"&date>={date_from}&date<{date_to}"
    )

    df = pd.DataFrame(fetch_paginated(url, headers))

    if df.empty:
        return df

    df = df[df["b2b_b2c_order"] == 2]
    df = df[df["subsidiary"].apply(lambda x: safe_get(x, "subsidiary") == "Fab BV")]
    df = df[df["marketplace_channel"].apply(lambda x: safe_get(x, "channel") is None)]

    df = df.rename(columns={
        "id": "Order no.",
        "date": "Date",
        "return_costs_lcy": "Return costs",
        "discount_lcy": "Discount",
        "amount_lcy": "Amount",
        "postage_costs_lcy": "Postage costs",
        "quantity": "Quantity",
        "remarks": "Comments"
    })

    df["Warehouse"] = df["warehouse"].apply(lambda x: safe_get(x, "warehouse"))
    df["Customer ID"] = df["customer"].apply(lambda x: safe_get(x, "id"))
    df["Customer"] = df["customer"].apply(lambda x: safe_get(x, "customer_name"))
    df["Country"] = df["country"].apply(lambda x: safe_get(x, "iso2"))
    df["Subsidiary"] = df["subsidiary"].apply(lambda x: safe_get(x, "subsidiary"))

    return df[[
        "Order no.", "Date", "Warehouse", "Customer ID", "Customer",
        "Return costs", "Discount", "Comments",
        "Country", "Subsidiary", "Quantity",
        "Amount", "Postage costs"

    ]]

