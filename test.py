from kintone_api import KintoneAPI
from pdf_gui import ColumnSelector  # นี่คือไฟล์ GUI ของพี่

DOMAIN = "https://kyoto-kobayashi-ss.cybozu.com"
APP_ID = 262
TOKEN = "9svLLQNwjN6Zqpm7n8a24zKDHW8w4YU90Fx177Sb"

k = KintoneAPI(DOMAIN, APP_ID, TOKEN)
records = k.fetch_all_records()
df = k.to_dataframe(records)

app = QApplication(sys.argv)
selector = ColumnSelector(df)  # <-- ใส่ df แทน Excel
selector.show()
sys.exit(app.exec_())