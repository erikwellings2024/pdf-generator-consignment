import streamlit as st
import pandas as pd
import os
import re
import calendar
import zipfile
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib import colors

st.title("📄 Consignment PDF Generator")

# =========================
# UPLOAD
# =========================
excel = st.file_uploader("Upload Excel", type=["xlsx"])
logo = st.file_uploader("Upload Logo", type=["jpg","png"])
sign = st.file_uploader("Upload Sign ITA", type=["jpg","png"])

# =========================
# STYLE
# =========================
styles = getSampleStyleSheet()

wrap_style = ParagraphStyle("wrap", parent=styles["Normal"], fontSize=6)
meta_style = ParagraphStyle("meta", parent=styles["Normal"], fontSize=8)
title_style = ParagraphStyle("title", parent=styles["Heading1"], fontSize=12, alignment=TA_CENTER)

# =========================
# FUNCTION
# =========================
def safe_int(x):
    try:
        return int(float(str(x).replace(",","")))
    except:
        return 0

def safe_filename(text):
    return re.sub(r"[\\/*?:\"<>|]","",str(text))

def get_period(df):
    dates = pd.to_datetime(df["Date"],errors="coerce").dropna()
    if dates.empty:
        return "-"
    first = dates.min()
    last_day = calendar.monthrange(first.year,first.month)[1]
    return f"1 - {last_day} {first.strftime('%b %Y')}"

# =========================
# GENERATE
# =========================
if st.button("🚀 Generate PDF"):

    if not excel:
        st.error("Upload Excel dulu")
    else:

        df = pd.read_excel(excel)
        df.columns = df.columns.str.strip()

        mapping = {
            "Vendor Name":"Vendor name",
            "store_location":"Cabang",
            "new_item_code":"Item Code",
            "item_name":"Item Name",
            "transaction_date":"Date",
            "qty":"Qty Sold",
            "Price/Unit Exclude Tax (Confirmed CM)":"Cost Price",
            "Total Purchase Exc PPN":"Total Purchase Price Exc Tax",
            "Total Purchase Inc PPN":"Total Purchase Price Inc Tax",
            "Nama PT":"KolomU"
        }

        df = df[list(mapping.keys())].rename(columns=mapping)

        headers = [
            "Vendor name","Cabang","Item Code","Item Name",
            "Date","Qty Sold","Cost Price",
            "Total Purchase Price Exc Tax","Total Purchase Price Inc Tax"
        ]

        periode = get_period(df)

        os.makedirs("output", exist_ok=True)

        pdf_files = []

        groups = df.groupby(["KolomU","Vendor name"])

        progress = st.progress(0)
        total = len(groups)

        for i,((pt,vendor),data) in enumerate(groups):

            st.write(f"Processing: {vendor}")

            file_name = f"{safe_filename(vendor)}.pdf"
            path = os.path.join("output", file_name)

            doc = SimpleDocTemplate(path, pagesize=landscape(A4))
            elements = []

            if logo:
                with open("logo.png","wb") as f:
                    f.write(logo.getbuffer())
                elements.append(Image("logo.png",width=120,height=45))

            elements.append(Paragraph("Sales Consignment Wellings",title_style))
            elements.append(Paragraph(vendor,meta_style))
            elements.append(Paragraph(f"Periode: {periode}",meta_style))
            elements.append(Paragraph(pt,meta_style))
            elements.append(Spacer(1,10))

            table_data = [headers] + data[headers].fillna("").values.tolist()

            table = Table(table_data)
            table.setStyle(TableStyle([
                ("GRID",(0,0),(-1,-1),0.25,colors.black),
                ("FONTSIZE",(0,0),(-1,-1),6)
            ]))

            elements.append(table)

            doc.build(elements)

            pdf_files.append(path)

            progress.progress((i+1)/total)

        # ZIP
        zip_path = "result.zip"
        with zipfile.ZipFile(zip_path, 'w') as z:
            for f in pdf_files:
                z.write(f)

        with open(zip_path, "rb") as f:
            st.download_button("📥 Download ZIP", f, file_name="result.zip")

        st.success("Selesai!")