import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import time

from services.shopify_service import fetch_shopify_reports
from services.itsperfect_returns import fetch_returns
from services.itsperfect_sales import fetch_sales_orders
from utils.excel import export_to_excel

def load_reference_sheet(path, sheet_name):
    return pd.read_excel(path, sheet_name=sheet_name, dtype=str)
st.write("App started")
st.title("E-commerce Reconciliation Export")

today = datetime.today()
first = today.replace(day=1)
last = (first + relativedelta(months=1)) - relativedelta(days=1)

date_range = st.date_input(
    "Date range",
    value=(first, last)
)

reference_excel = st.file_uploader(
    label="Upload reference Excel",
    type="xlsx"
)

if isinstance(date_range, tuple) and len(date_range) == 2:
    start_date, end_date = date_range
else:
    start_date, end_date = None, None

if reference_excel is None:
    st.info("Please upload the reference Excel to enable the Generate button.")
else:
    if st.button("Generate Excel"):
        with st.spinner("In progress..."):
            status_text = st.empty()
    
            report_index = 0
            t0 = time.perf_counter()
            shopify_dfs = fetch_shopify_reports(start_date=start_date.strftime("%Y-%m-%d"), 
                                            end_date=end_date.strftime("%Y-%m-%d"))
            t1 = time.perf_counter()
            returns_df = fetch_returns(
                f"{start_date} 00:00:00",
                f"{end_date} 23:59:59"
            )
            t2 = time.perf_counter()
            sales_df = fetch_sales_orders(
                f"{start_date} 00:00:00",
                f"{end_date} 23:59:59"
            )
            t3 = time.perf_counter()
    
            sales_df_copy = sales_df.copy()
            sales_df_copy["VAT %"] = (
                sales_df_copy["VAT value"]
                .astype(str).str.replace(",", ".", regex=False)
                .pipe(pd.to_numeric, errors="coerce")
                .div(
                    sales_df_copy["Shipping costs"]
                    .astype(str).str.replace(",", ".", regex=False)
                    .pipe(pd.to_numeric, errors="coerce")
                    +
                    sales_df_copy["Amount"]
                    .astype(str).str.replace(",", ".", regex=False)
                    .pipe(pd.to_numeric, errors="coerce")
                )
                .replace([float("inf"), -float("inf")], 0)
                .fillna(0)
            )
    
            t4 = time.perf_counter()
            backend_df = load_reference_sheet(reference_excel, "Backend")
            t5 = time.perf_counter()
            old_itsp_df = load_reference_sheet(reference_excel, "Old ITSP")
            t6 = time.perf_counter()
    
            KEY_COL = "Order no."
            existing_orders = set(old_itsp_df[KEY_COL].dropna().astype(str))
            sales_df_copy[KEY_COL] = sales_df_copy[KEY_COL].astype(str)
    
            new_sales_rows = sales_df_copy[
                ~sales_df_copy[KEY_COL].isin(existing_orders)
            ]
    
            new_sales_rows["Marketplace > Channel"] = None
            new_sales_rows = new_sales_rows[old_itsp_df.columns]
    
            old_itsp_combined = pd.concat(
                [old_itsp_df, new_sales_rows],
                ignore_index=True
            )
    
            sheets = {
                **shopify_dfs,
                "ITSP Sales": sales_df,
                "ITSP Returns": returns_df,
                "Old ITSP": old_itsp_combined,
                "Backend": backend_df,
            }
            t7 = time.perf_counter()
            output = export_to_excel(sheets)
            if "excel_file" not in st.session_state:
                st.session_state.excel_file = output
            t8 = time.perf_counter()
    
            # st.info(
            #     f"""
            #     Shopify fetch: {t1 - t0:.2f}s  
            #     Returns fetch: {t2 - t1:.2f}s  
            #     Sales fetch: {t3 - t2:.2f}s  
            #     Excel export: {t4 - t3:.2f}s  
            #     Backend fetch: {t5 - t4:.2f}s
            #     Old ITSP fetch: {t6 - t5:.2f}s
            #     Combine sheets: {t7 - t6:.2f}s
            #     Excel export: {t8 - t7:.2f}s
            #     **Total:** {t8 - t0:.2f}s
            #     """
            # )
        st.download_button(
            "Download Excel",
            st.session_state.excel_file,
            file_name="ecom_recon.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    
        )

