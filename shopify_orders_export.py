#!/usr/bin/env python3
"""
Shopify orders -> CSV (aur optional Excel), bina webhook aur bina dashboard ke.

Setup:
    pip install requests            # Excel ke liye: pip install openpyxl
    export SHOPIFY_SHOP="your-store.myshopify.com"
    export SHOPIFY_TOKEN="shpat_xxxxx"

Usage:
    python shopify_orders_export.py                       # sare orders (paged mode)
    python shopify_orders_export.py --since 2026-10-01    # sirf is date ke baad ke
    python shopify_orders_export.py --mode bulk           # bahut zyada orders (10k+)
    python shopify_orders_export.py --xlsx                # CSV ke saath Excel bhi
Output: ek row = ek line item (order ki details har row me repeat hoti hain).
"""
import argparse, csv, json, os, sys, time
import requests

SHOP = os.environ.get("SHOPIFY_SHOP")
TOKEN = os.environ.get("SHOPIFY_TOKEN")
API_VERSION = os.environ.get("SHOPIFY_API_VERSION", "2025-10")

if not SHOP or not TOKEN:
    sys.exit("SHOPIFY_SHOP aur SHOPIFY_TOKEN environment variables set karo.")

URL = f"https://{SHOP}/admin/api/{API_VERSION}/graphql.json"
HEADERS = {"X-Shopify-Access-Token": TOKEN, "Content-Type": "application/json"}

COLUMNS = [
    "order_id", "order_name", "created_at", "financial_status", "fulfillment_status",
    "currency", "total_price", "customer_name", "customer_email",
    "line_title", "line_sku", "line_quantity",
]

ORDER_FIELDS = """
  id
  name
  createdAt
  displayFinancialStatus
  displayFulfillmentStatus
  totalPriceSet { shopMoney { amount currencyCode } }
  customer { firstName lastName email }
"""
LINE_FIELDS = "title sku quantity"


def gql(query, variables=None, retries=5):
    """GraphQL call, throttle/429 par automatic retry."""
    for attempt in range(retries):
        r = requests.post(URL, headers=HEADERS,
                          json={"query": query, "variables": variables or {}}, timeout=60)
        if r.status_code == 429:
            time.sleep(2 ** attempt)
            continue
        r.raise_for_status()
        body = r.json()
        errs = body.get("errors")
        if errs:
            if any(e.get("extensions", {}).get("code") == "THROTTLED" for e in errs):
                time.sleep(2 ** attempt)
                continue
            sys.exit(f"GraphQL error: {json.dumps(errs, indent=2)}")
        return body["data"]
    sys.exit("Bahut baar retry kiya, abhi bhi throttle ho raha hai.")


def flatten(order, lines):
    """Ek order + uski line items -> CSV rows."""
    money = (order.get("totalPriceSet") or {}).get("shopMoney") or {}
    cust = order.get("customer") or {}
    base = {
        "order_id": order["id"].split("/")[-1],
        "order_name": order.get("name"),
        "created_at": order.get("createdAt"),
        "financial_status": order.get("displayFinancialStatus"),
        "fulfillment_status": order.get("displayFulfillmentStatus"),
        "currency": money.get("currencyCode"),
        "total_price": money.get("amount"),
        "customer_name": f"{cust.get('firstName') or ''} {cust.get('lastName') or ''}".strip(),
        "customer_email": cust.get("email"),
    }
    if not lines:
        return [{**base, "line_title": "", "line_sku": "", "line_quantity": ""}]
    return [{**base, "line_title": l.get("title"), "line_sku": l.get("sku"),
             "line_quantity": l.get("quantity")} for l in lines]


# ---------- Mode 1: paged (normal stores) ----------
def fetch_paged(since):
    q = f"""
    query($cursor: String, $search: String) {{
      orders(first: 100, after: $cursor, query: $search, sortKey: CREATED_AT) {{
        pageInfo {{ hasNextPage endCursor }}
        nodes {{ {ORDER_FIELDS} lineItems(first: 100) {{ nodes {{ {LINE_FIELDS} }} }} }}
      }}
    }}"""
    search = f"created_at:>={since}" if since else None
    cursor, rows, n = None, [], 0
    while True:
        data = gql(q, {"cursor": cursor, "search": search})["orders"]
        for o in data["nodes"]:
            rows.extend(flatten(o, o["lineItems"]["nodes"]))
            n += 1
        print(f"  {n} orders fetched...", end="\r")
        if not data["pageInfo"]["hasNextPage"]:
            break
        cursor = data["pageInfo"]["endCursor"]
    print()
    return rows


# ---------- Mode 2: bulk operation (10k+ orders) ----------
def fetch_bulk(since):
    search = f'(query: "created_at:>={since}")' if since else ""
    inner = f"""
    {{
      orders{search} {{
        edges {{ node {{ {ORDER_FIELDS}
          lineItems {{ edges {{ node {{ {LINE_FIELDS} }} }} }}
        }} }}
      }}
    }}"""
    start = gql('''
    mutation($q: String!) {
      bulkOperationRunQuery(query: $q) {
        bulkOperation { id status }
        userErrors { field message }
      }
    }''', {"q": inner})["bulkOperationRunQuery"]
    if start["userErrors"]:
        sys.exit(f"Bulk start error: {start['userErrors']}")
    print("Bulk operation start ho gaya, Shopify file bana raha hai...")

    while True:
        op = gql("{ currentBulkOperation { id status errorCode objectCount url } }")["currentBulkOperation"]
        print(f"  status: {op['status']}  objects: {op['objectCount']}", end="\r")
        if op["status"] == "COMPLETED":
            break
        if op["status"] in ("FAILED", "CANCELED", "EXPIRED"):
            sys.exit(f"\nBulk operation {op['status']}: {op.get('errorCode')}")
        time.sleep(5)
    print()

    if not op["url"]:  # koi order match nahi hua
        return []

    orders, lines = {}, {}
    with requests.get(op["url"], stream=True, timeout=300) as r:
        r.raise_for_status()
        for raw in r.iter_lines():
            if not raw:
                continue
            obj = json.loads(raw)
            parent = obj.get("__parentId")
            if parent:
                lines.setdefault(parent, []).append(obj)
            else:
                orders[obj["id"]] = obj
    rows = []
    for oid, o in orders.items():
        rows.extend(flatten(o, lines.get(oid, [])))
    return rows


# ---------- Output ----------
def write_csv(rows, path):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:  # utf-8-sig = Excel me Hindi/special chars sahi dikhte hain
        w = csv.DictWriter(f, fieldnames=COLUMNS)
        w.writeheader()
        w.writerows(rows)


def write_xlsx(rows, path):
    try:
        from openpyxl import Workbook
    except ImportError:
        sys.exit("Excel ke liye pehle: pip install openpyxl")
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    ws.append(COLUMNS)
    for r in rows:
        ws.append([r[c] for c in COLUMNS])
    wb.save(path)


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--mode", choices=["paged", "bulk"], default="paged")
    p.add_argument("--since", help="YYYY-MM-DD, is date ke baad ke orders")
    p.add_argument("--out", default="orders.csv")
    p.add_argument("--xlsx", action="store_true", help="Excel file bhi banao")
    a = p.parse_args()

    rows = fetch_bulk(a.since) if a.mode == "bulk" else fetch_paged(a.since)
    write_csv(rows, a.out)
    print(f"{len(rows)} rows -> {a.out}")
    if a.xlsx:
        x = os.path.splitext(a.out)[0] + ".xlsx"
        write_xlsx(rows, x)
        print(f"Excel -> {x}")


if __name__ == "__main__":
    main()
