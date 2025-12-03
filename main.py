from kintone_api import KintoneAPI
from invoice_template import create_invoice_pdf

DOMAIN = "https://kyoto-kobayashi-ss.cybozu.com"
APP_ID = 262
TOKEN = "9svLLQNwjN6Zqpm7n8a24zKDHW8w4YU90Fx177Sb"

k = KintoneAPI(DOMAIN, APP_ID, TOKEN)

records = k.fetch_all_records()

target_no = "B111"

for rec in records:
    if rec["invoice_no"]["value"] == target_no:
        header = rec
        table = header["subdata"]
        create_invoice_pdf("invoice.pdf", header, table)
        break

table   = header["subdata"]            # เอา subtable

create_invoice_pdf("invoice.pdf", header, table)
