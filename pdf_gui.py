import os
import json
import unicodedata
import pandas as pd
import numpy as np
from datetime import datetime
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog,
    QGridLayout, QScrollArea, QCheckBox, QLineEdit, QHBoxLayout,
    QMessageBox, QComboBox
)
from PyQt5.QtCore import Qt
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image as RLImage
from reportlab.lib.pagesizes import A4, A3, LETTER, landscape, portrait
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import qrcode
from io import BytesIO


# ========== CONFIG FILE ================
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "font_config.json")

PAGE_SIZES = {
    "A4": A4,
    "A3": A3,
    "LETTER": LETTER
}

ORIENTATIONS = {
    "縦": portrait,
    "横": landscape
}


# ========== HELPER FUNCTIONS ================
def normalize_to_ascii(text):
    if not isinstance(text, str):
        return str(text)
    text = unicodedata.normalize('NFKC', text)
    return ''.join(c for c in text)


def format_value(value):
    """รูปแบบข้อมูลสำหรับ PDF"""
    if pd.isna(value) or value is None:
        return ""

    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, (int, np.integer)):
        return str(int(value))

    if isinstance(value, float):
        if np.isnan(value):
            return ""
        if value == int(value):
            return str(int(value))
        return str(value)

    return str(value)


# ========== GUI CLASS ================
class ColumnSelector(QWidget):
    def __init__(self, df):
        super().__init__()
        self.df = df

        self.setWindowTitle("カラム選択 + PDF出力")
        self.resize(850, 600)

        self.checkboxes = []  # (col, display, qr, hide_zero)

        layout = QVBoxLayout()

        # ===== Scroll area =====
        scroll = QScrollArea()
        scroll_widget = QWidget()
        grid = QGridLayout()

        for i, col in enumerate(df.columns):
            label = QLabel(col)
            cb_disp = QCheckBox("表示")
            cb_qr = QCheckBox("QR")
            cb_hide_zero = QCheckBox("0非表示")

            grid.addWidget(label, i, 0)
            grid.addWidget(cb_disp, i, 1)
            grid.addWidget(cb_qr, i, 2)
            grid.addWidget(cb_hide_zero, i, 3)

            self.checkboxes.append((col, cb_disp, cb_qr, cb_hide_zero))

        scroll_widget.setLayout(grid)
        scroll.setWidget(scroll_widget)
        scroll.setWidgetResizable(True)
        layout.addWidget(scroll)

        # ===== Font + Page Settings =====
        font_layout = QHBoxLayout()

        self.font_input = QLineEdit()
        self.font_input.setPlaceholderText("内部フォント名 (例: NotoSansJP)")

        self.font_path_input = QLineEdit()
        self.font_path_input.setPlaceholderText("フォントファイルを選択")

        btn_browse = QPushButton("選択")
        btn_browse.clicked.connect(self.select_font_file)

        self.font_size_input = QLineEdit()
        self.font_size_input.setPlaceholderText("サイズ 例: 10")

        self.page_size_combo = QComboBox()
        self.page_size_combo.addItems(PAGE_SIZES.keys())

        self.orientation_combo = QComboBox()
        self.orientation_combo.addItems(ORIENTATIONS.keys())

        font_layout.addWidget(QLabel("フォント名:"))
        font_layout.addWidget(self.font_input)
        font_layout.addWidget(QLabel("ファイル:"))
        font_layout.addWidget(self.font_path_input)
        font_layout.addWidget(btn_browse)
        font_layout.addWidget(QLabel("サイズ:"))
        font_layout.addWidget(self.font_size_input)
        font_layout.addWidget(QLabel("用紙:"))
        font_layout.addWidget(self.page_size_combo)
        font_layout.addWidget(QLabel("方向:"))
        font_layout.addWidget(self.orientation_combo)

        layout.addLayout(font_layout)

        # ===== Generate PDF Button =====
        btn_pdf = QPushButton("PDF作成")
        btn_pdf.clicked.connect(self.generate_pdf)
        layout.addWidget(btn_pdf)

        self.setLayout(layout)
        self.load_font_settings()

    # -----------------------------
    def select_font_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "フォントファイル選択", "", "Font Files (*.ttf *.ttc)")
        if path:
            self.font_path_input.setText(path)

    # -----------------------------
    def load_font_settings(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                self.font_input.setText(cfg.get("font_name", "Helvetica"))
                self.font_path_input.setText(cfg.get("font_path", ""))
                self.font_size_input.setText(str(cfg.get("font_size", 10)))
                self.page_size_combo.setCurrentText(cfg.get("page_size", "A4"))
                self.orientation_combo.setCurrentText(cfg.get("orientation", "縦"))

    # -----------------------------
    def save_font_settings(self, name, path, size, page, ori):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({
                "font_name": name,
                "font_path": path,
                "font_size": size,
                "page_size": page,
                "orientation": ori
            }, f, ensure_ascii=False, indent=2)

    # -----------------------------
    def generate_pdf(self):
        font_name = self.font_input.text().strip()
        font_path = self.font_path_input.text().strip()

        try:
            font_size = int(self.font_size_input.text().strip())
        except:
            QMessageBox.critical(self, "エラー", "フォントサイズは数字で入力してください")
            return

        if not os.path.exists(font_path):
            QMessageBox.critical(self, "エラー", "フォントファイルが存在しません")
            return

        # register font
        try:
            pdfmetrics.registerFont(TTFont(font_name, font_path))
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"フォント登録に失敗: {e}")
            return

        page_key = self.page_size_combo.currentText()
        ori_key = self.orientation_combo.currentText()
        page_size = ORIENTATIONS[ori_key](PAGE_SIZES[page_key])

        self.save_font_settings(font_name, font_path, font_size, page_key, ori_key)

        selected_columns = [
            col for col, cb_disp, _, _ in self.checkboxes if cb_disp.isChecked()
        ]

        if not selected_columns:
            QMessageBox.warning(self, "警告", "1つ以上のカラムを選択してください")
            return

        # ====== PDF OUTPUT FOLDER ======
        pdf_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output_kintone.pdf")

        doc = SimpleDocTemplate(pdf_path, pagesize=page_size)
        elements = []

        # ===== TABLE HEADER =====
        table_data = [selected_columns + ["QR"]]

        # ===== ADD ROWS =====
        for _, row in self.df.iterrows():
            skip = False

            # hide zero logic
            for col, _, _, cb_hide_zero in self.checkboxes:
                if cb_hide_zero.isChecked():
                    val = row[col]
                    try:
                        if float(str(val).strip()) == 0:
                            skip = True
                            break
                    except:
                        pass

            if skip:
                continue

            row_cells = []
            qr_text_parts = []

            for col, cb_disp, cb_qr, _ in self.checkboxes:
                if cb_disp.isChecked():
                    fv = format_value(row[col])
                    row_cells.append(normalize_to_ascii(fv))

                if cb_qr.isChecked():
                    fv = format_value(row[col])
                    qr_text_parts.append(normalize_to_ascii(fv))

            # build QR
            qr_text = ",".join(qr_text_parts).strip()

            if qr_text:
                qr = qrcode.QRCode(
                    version=None,
                    error_correction=qrcode.constants.ERROR_CORRECT_M,
                    box_size=4,
                    border=4
                )
                qr.add_data(qr_text)
                qr.make(fit=True)
                img = qr.make_image(fill_color="black", back_color="white")

                buf = BytesIO()
                img.save(buf, format='PNG')
                qr_img = RLImage(BytesIO(buf.getvalue()), width=60, height=60)
            else:
                qr_img = ""

            row_cells.append(qr_img)
            table_data.append(row_cells)

        # ===== CREATE TABLE =====
        table = Table(table_data, repeatRows=1)
        table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('BACKGROUND', (0,0), (-2,0), colors.lightgrey),
            ('FONTNAME', (0,0), (-2,-1), font_name),
            ('FONTSIZE', (0,0), (-2,-1), font_size),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))

        elements.append(table)
        doc.build(elements)

        QMessageBox.information(self, "完了", f"PDF を作成しました:\n{pdf_path}")
