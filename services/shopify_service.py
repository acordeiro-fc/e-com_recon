import time
import requests
import pandas as pd

from config.shopify import (
    ACCESS_TOKEN,
    ACCESS_TOKEN_ARCHIVE,
    GRAPHQL_URL,
    GRAPHQL_URL_ARCHIVE,
)

SHOPIFY_RENAME_MAPS = {
    "payments": {
        "transaction_id": "Transaction ID",
        "day": "Date",
        "order_name": "Order",
        "payment_gateway": "Payment method",
        "credit_card_type": "Accelerated checkout",
        "credit_card_tier": "Credit card",
        "shipping_country": "Sales channel",
        "billing_country": "Billing country",
        "gift_card_id": "Gift card ID",
        "gross_payments": "Gross payments",
        "refunded_payments": "Refunds",
        "net_payments": "Net payments",
    },
    "incl_returns": {
        "order_id": "Order ID",
        "sale_id": "Sale ID",
        "order_name": "Order",
        "day": "Date",
        "order_or_return": "Sale type",
        "sales_channel": "Sales channel",
        "pos_location_name": "POS location",
        "billing_country": "Billing country",
        "shipping_country": "Shipping country",
        "product_type": "Product type",
        "product_vendor": "Product vendor",
        "product_title": "Product",
        "product_variant_title": "Variant",
        "product_variant_sku": "Variant SKU",
        "quantity_ordered": "Net quantity",
        "gross_sales": "Gross sales",
        "discounts": "Discounts",
        "returns": "Returns",
        "net_sales": "Net sales",
        "shipping_charges": "Shipping",
        "taxes": "Taxes",
        "total_sales": "Total sales",
    },
    "tax": {
        "line_item_id": "Sale tax ID",
        "order_id": "Order ID",
        "day": "Date",
        "order_name": "Order",
        "product_title": "Product",
        "product_variant_title": "Variant",
        "product_variant_sku": "Variant SKU",
        "product_type": "Product type",
        "tax_country": "Country",
        "tax_region": "Region",
        "tax_name": "Name",
        "tax_rate": "Rate",
        "sales_taxes": "Amount",
        "sales_channel": "Sales channel",
    },
}

# --------------------------------------------------
# Low-level Shopify POST with retries & throttling
# --------------------------------------------------
def shopify_post(query, access_token, graphql_url, max_retries=5, initial_delay=5):
    headers = {
        "X-Shopify-Access-Token": access_token,
        "Content-Type": "application/json",
    }

    delay = initial_delay
    for attempt in range(max_retries):
        r = requests.post(graphql_url, json={"query": query}, headers=headers)

        try:
            data = r.json()
        except ValueError:
            time.sleep(2)
            continue

        errors = (
            data.get("errors")
            or data.get("data", {}).get("shopifyqlQuery", {}).get("parseErrors", [])
        )

        if errors:
            print(errors)
            if any(e.get("extensions", {}).get("code") == "THROTTLED" for e in errors):
                time.sleep(delay)
                delay *= 2
                continue
            raise Exception(f"Shopify GraphQL error: {errors}")

        return data

    raise Exception("Shopify API failed after retries")

# --------------------------------------------------
# Generic ShopifyQL fetcher (pagination)
# --------------------------------------------------
def fetch_shopifyql(
    query_template,
    access_token,
    graphql_url,
    start_date,
    end_date,
    batch_size=3000,
):
    all_rows = []
    offset = 0

    while True:
        query = query_template.format(
            start_date=start_date,
            end_date=end_date,
            limit=batch_size,
            offset=offset,
        )

        data = shopify_post(query, access_token, graphql_url)

        table = data["data"]["shopifyqlQuery"]["tableData"]
        rows = table.get("rows", [])
        cols = [c["name"] for c in table.get("columns", [])]

        if not rows:
            break

        all_rows.extend(rows)
        offset += batch_size

        if len(rows) < batch_size:
            break

    if not all_rows:
        return pd.DataFrame()

    return pd.DataFrame(all_rows, columns=cols)

# --------------------------------------------------
# Individual report functions
# --------------------------------------------------
def fetch_shopify_payments(start_date, end_date, access_token, graphql_url):
    query = """
    query {{
        shopifyqlQuery(
            query: "
            FROM payments
            SHOW gross_payments, refunded_payments, net_payments
            WHERE transaction_kind IN ('sale','change','capture','refund')
                AND order_name!=''
            GROUP BY transaction_id, day, order_name, payment_gateway,
                     credit_card_type, credit_card_tier, shipping_country,
                     billing_country, gift_card_id
            TIMESERIES day
            SINCE {start_date} UNTIL {end_date}
            ORDER BY day ASC
            LIMIT {limit}
            OFFSET {offset}
            VISUALIZE net_payments TYPE table
            "
        ) {{
            tableData {{
                columns {{ name }}
                rows
            }}
            parseErrors
        }}
    }}
    """
    df= fetch_shopifyql(query, access_token, graphql_url, start_date, end_date)
    return df.rename(columns=SHOPIFY_RENAME_MAPS["payments"])

def fetch_shopify_incl_returns(start_date, end_date, access_token, graphql_url):
    query = """
    query {{
        shopifyqlQuery(
            query: "
            FROM sales
            SHOW quantity_ordered, gross_sales, discounts, returns,
                 net_sales, shipping_charges, taxes, total_sales
            GROUP BY order_id, sale_id, order_name, day, order_or_return,
                     sales_channel, pos_location_name, billing_country,
                     shipping_country, product_type, product_vendor,
                     product_title, product_variant_title, product_variant_sku
            SINCE {start_date} UNTIL {end_date}
            ORDER BY day ASC, order_id ASC, sale_id ASC
            LIMIT {limit}
            OFFSET {offset}
            "
        ) {{
            tableData {{
                columns {{ name }}
                rows
            }}
            parseErrors
        }}
    }}
    """
    df= fetch_shopifyql(query, access_token, graphql_url, start_date, end_date)
    return df.rename(columns=SHOPIFY_RENAME_MAPS["incl_returns"])

def fetch_shopify_tax(start_date, end_date, access_token, graphql_url):
    query = """
    query {{
        shopifyqlQuery(
            query: "
            FROM sales_taxes
            SHOW sales_taxes
            GROUP BY line_item_id, order_id, day, order_fulfillment_status,
                     order_payment_status, order_name, product_title,
                     product_variant_title, product_variant_sku,
                     product_type, tax_country, tax_region, tax_name,
                     tax_rate, sales_channel, filed_by_channel, is_canceled_order
            SINCE {start_date} UNTIL {end_date}
            ORDER BY order_id ASC
            LIMIT {limit}
            OFFSET {offset}
            VISUALIZE sales_taxes TYPE table
            "
        ) {{
            tableData {{
                columns {{ name }}
                rows
            }}
            parseErrors
        }}
    }}
    """
    df= fetch_shopifyql(query, access_token, graphql_url, start_date, end_date)
    return df.rename(columns=SHOPIFY_RENAME_MAPS["tax"])

# --------------------------------------------------
# Public API: fetch all reports (live + archive)
# --------------------------------------------------
def fetch_shopify_reports(start_date, end_date):
    results = {}

    results["Shopify payments"] = pd.concat([
        fetch_shopify_payments(start_date, end_date, ACCESS_TOKEN, GRAPHQL_URL),
        fetch_shopify_payments(start_date, end_date, ACCESS_TOKEN_ARCHIVE, GRAPHQL_URL_ARCHIVE)
    ], ignore_index=True)

    results["Shopify incl. returns"] = pd.concat([
        fetch_shopify_incl_returns(start_date, end_date, ACCESS_TOKEN, GRAPHQL_URL),
        fetch_shopify_incl_returns(start_date, end_date, ACCESS_TOKEN_ARCHIVE, GRAPHQL_URL_ARCHIVE)
    ], ignore_index=True)

    results["Shopify Tax"] = pd.concat([
        fetch_shopify_tax(start_date, end_date, ACCESS_TOKEN, GRAPHQL_URL),
        fetch_shopify_tax(start_date, end_date, ACCESS_TOKEN_ARCHIVE, GRAPHQL_URL_ARCHIVE)
    ], ignore_index=True)

    return results