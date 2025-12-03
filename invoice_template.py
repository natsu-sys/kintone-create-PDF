from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.units import mm
from datetime import datetime


# ---------------------------
# Convert to Japanese era
# ---------------------------
def to_wareki(date_str):
    if not date_str:
        return ""
    y, m, d = map(int, date_str.split("-"))

    if y >= 2019:
        return f"令和{y - 2018}年{m}月{d}日"
    if y >= 1989:
        return f"平成{y - 1988}年{m}月{d}日"

    return f"{y}年{m}月{d}日"


def create_invoice_pdf(output_path, header, table_records):

    # ---------------------------
    # Register Japanese Font
    # ---------------------------
    pdfmetrics.registerFont(TTFont("JP", r"C:\Windows\Fonts\meiryo.ttc"))

    styles = {
        "title": ParagraphStyle(
            name="Title", fontName="JP", fontSize=22, alignment=TA_CENTER
        ),
        "normal": ParagraphStyle(
            name="Normal",
            fontName="JP",
            fontSize=12,
        ),
        "right": ParagraphStyle(
            name="Right",
            fontName="JP",
            fontSize=12,
            alignment=TA_RIGHT,
        ),
        "small": ParagraphStyle(
            name="Small",
            fontName="JP",
            fontSize=10,
        ),
    }

    doc = SimpleDocTemplate(output_path, pagesize=A4, leftMargin=25, rightMargin=25)
    elements = []

    # ---------------------------
    # DATE (from invoice_date)
    # ---------------------------
    wareki_date = to_wareki(header["invoice_date"]["value"])
    elements.append(Paragraph(wareki_date, styles["right"]))
    elements.append(Spacer(1, 12))

    # ---------------------------
    # TITLE
    # ---------------------------
    elements.append(Paragraph("御請求書", styles["title"]))
    elements.append(Spacer(1, 20))

    # ---------------------------
    # Customer
    # ---------------------------
    customer_block = (
        f"{header['customer']['value']}<br/>" f"担当者：{header['staff']['value']} 様"
    )
    elements.append(Paragraph(customer_block, styles["normal"]))
    elements.append(Spacer(1, 12))

    # ---------------------------
    # Price Summary
    # ---------------------------
    price_data = [
        ["本体価格", "消費税", "合計金額"],
        [
            f"¥{int(header['amount_sum']['value']):,}",
            f"¥{int(header['vat']['value']):,}",
            f"¥{int(header['total']['value']):,}",
        ],
    ]

    price_table = Table(price_data, colWidths=[65 * mm, 65 * mm, 65 * mm])
    price_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("FONTNAME", (0, 0), (-1, -1), "JP"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ]
        )
    )
    elements.append(price_table)
    elements.append(Spacer(1, 20))

    # ---------------------------
    # Table (subdata)
    # ---------------------------
    sublist = table_records["value"]

    item_data = [["商品名", "単価", "数量", "単位", "小計"]]

    for row in sublist:
        v = row["value"]

        item_data.append(
            [
                v["name"]["value"],
                f"¥{int(v['price']['value']):,}",
                int(v["qty"]["value"]),
                "個",
                f"¥{int(v['amount']['value']):,}",
            ]
        )

    items_table = Table(
        item_data,
        colWidths=[65 * mm, 20 * mm, 20 * mm, 20 * mm, 40 * mm],
    )
    items_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightblue),
                ("FONTNAME", (0, 0), (-1, -1), "JP"),
                ("ALIGN", (1, 1), (-1, -1), "CENTER"),
            ]
        )
    )
    elements.append(items_table)
    elements.append(Spacer(1, 20))

    # ---------------------------
    # Company info
    # ---------------------------
    foot = (
        "Kobayashi Manufacturing Co., Ltd.<br/>"
        "11-15, The Jar of Records, Ryuji Sho<br/>"
        "Tel: 075-9547200"
    )
    elements.append(Paragraph(foot, styles["small"]))
    elements.append(Spacer(1, 20))

    # ---------------------------
    # Notes box
    # ---------------------------
    notes = Table(
        [["備考欄"]],
        colWidths=[180 * mm],
        rowHeights=[40 * mm],
    )
    notes.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("FONTNAME", (0, 0), (-1, -1), "JP"),
            ]
        )
    )
    elements.append(notes)

    doc.build(elements)
