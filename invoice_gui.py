import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QListWidget, QMessageBox
)
from kintone_api import KintoneAPI
from invoice_template import create_invoice_pdf

DOMAIN = "https://kyoto-kobayashi-ss.cybozu.com"
APP_ID = 262
TOKEN = "9svLLQNwjN6Zqpm7n8a24zKDHW8w4YU90Fx177Sb"


class InvoiceGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Kintone Invoice Generator")
        self.resize(400, 500)

        self.layout = QVBoxLayout()

        self.btn_load = QPushButton("Kintone からデータ取得")
        self.btn_load.clicked.connect(self.load_records)
        self.layout.addWidget(self.btn_load)

        self.list = QListWidget()
        self.layout.addWidget(self.list)

        self.btn_generate = QPushButton("選択したレコードでPDF生成")
        self.btn_generate.clicked.connect(self.generate_pdf)
        self.layout.addWidget(self.btn_generate)

        self.records = []

        self.setLayout(self.layout)

    def load_records(self):
        try:
            k = KintoneAPI(DOMAIN, APP_ID, TOKEN)
            self.records = k.fetch_all_records()

            self.list.clear()

            for rec in self.records:
                invo = rec["invoice_no"]["value"]
                cust = rec["customer"]["value"]
                date = rec["invoice_date"]["value"]

                self.list.addItem(f"{invo} | {cust} | {date}")

        except Exception as e:
            QMessageBox.critical(self, "エラー", str(e))

    def generate_pdf(self):
        idx = self.list.currentRow()
        if idx == -1:
            QMessageBox.warning(self, "警告", "レコードを選択してください")
            return

        rec = self.records[idx]

        header = rec
        table = header["subdata"]

        filename = f"invoice_{header['invoice_no']['value']}.pdf"

        create_invoice_pdf(filename, header, table)

        QMessageBox.information(self, "完了", f"PDF を作成しました：\n{filename}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    gui = InvoiceGUI()
    gui.show()
    sys.exit(app.exec_())
