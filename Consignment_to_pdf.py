import os
import re
import calendar
import pandas as pd
from tkinter import Tk, filedialog
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib import colors
from reportlab.pdfgen import canvas

LOGO_FILENAME = "LOGO WELLINGS RESIZE.jpg"
SIGN_FILENAME = "ITA SIGN.png"

# =========================
# CONSOLE HOLD
# =========================

def hold_console():
    print("\nTekan ENTER untuk keluar...")
    input()

# =========================
# PAGE NUMBER
# =========================

class NumberedCanvas(canvas.Canvas):

    def __init__(self,*args,**kwargs):
        super().__init__(*args,**kwargs)
        self.pages = []

    def showPage(self):
        self.pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total_pages = len(self.pages)

        for page in self.pages:
            self.__dict__.update(page)
            self.draw_page_number(total_pages)
            super().showPage()

        super().save()

    def draw_page_number(self,total):

        width,_ = landscape(A4)

        self.setFont("Helvetica",8)

        self.drawRightString(width-20,10,f"{self._pageNumber}/{total}")

# =========================
# STYLE
# =========================

styles = getSampleStyleSheet()

wrap_style = ParagraphStyle(
    "wrap",
    parent=styles["Normal"],
    fontSize=6,
    leading=8
)

meta_style = ParagraphStyle(
    "meta",
    parent=styles["Normal"],
    fontSize=8,
    alignment=TA_LEFT
)

title_style = ParagraphStyle(
    "title",
    parent=styles["Heading1"],
    fontSize=12,
    alignment=TA_CENTER
)

sign_style = ParagraphStyle(
    "sign",
    parent=styles["Normal"],
    fontSize=8,
    alignment=TA_CENTER
)

# =========================
# CLEAN TEXT
# =========================

def clean_text(x):

    if pd.isna(x):
        return ""

    return str(x).strip()

# =========================
# SAFE NUMBER
# =========================

def safe_int(x):

    try:

        if pd.isna(x):
            return 0

        return int(float(str(x).replace(",","")))

    except:

        return 0

# =========================
# SAFE FILENAME
# =========================

def safe_filename(text):

    text = re.sub(r"[\\/*?:\"<>|]","",str(text))

    return text.strip()

# =========================
# FORMAT PERIOD
# =========================

def get_period(df):

    dates = pd.to_datetime(df["Date"],errors="coerce").dropna()

    if dates.empty:

        return "Periode Tidak Ditemukan"

    first = dates.min()

    last_day = calendar.monthrange(first.year,first.month)[1]

    return f"1 - {last_day} {first.strftime('%b %Y')}"

# =========================
# TABLE DATA
# =========================

def build_table(df,headers):

    table_data = [[Paragraph(h,wrap_style) for h in headers]]

    for _,row in df.iterrows():

        row_data = []

        for col in headers:

            val = row[col]

            if col == "Date":

                try:
                    val = pd.to_datetime(val).strftime("%d/%m/%Y")
                except:
                    val = ""

            elif col in ["Qty Sold","Cost Price","Total Purchase Price Exc Tax","Total Purchase Price Inc Tax"]:

                val = f"{safe_int(val):,}"

            row_data.append(Paragraph(str(val),wrap_style))

        table_data.append(row_data)

    return table_data

# =========================
# GENERATE PDF
# =========================

def generate_pdf(path,vendor,periode,pt,df,headers,logo,sign):

    print(f"Generate PDF : {vendor}")

    doc = SimpleDocTemplate(
        path,
        pagesize=landscape(A4),
        leftMargin=18,
        rightMargin=18,
        topMargin=25,
        bottomMargin=25
    )

    elements = []

    if os.path.exists(logo):

        elements.append(Image(logo,width=120,height=45))

    elements.append(Paragraph("Sales Consignment Wellings",title_style))

    elements.append(Spacer(1,6))

    elements.append(Paragraph(vendor,meta_style))
    elements.append(Paragraph(f"Periode : {periode}",meta_style))
    elements.append(Paragraph(pt,meta_style))

    elements.append(Spacer(1,10))

    table_data = build_table(df,headers)

    col_width = (landscape(A4)[0]-36)/len(headers)

    table = Table(table_data,repeatRows=1,colWidths=[col_width]*len(headers))

    table.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
        ("GRID",(0,0),(-1,-1),0.25,colors.black),
        ("FONTSIZE",(0,0),(-1,-1),6)
    ]))

    elements.append(table)

    # TOTAL

    total_row = []

    for col in headers:

        if col in ["Qty Sold","Total Purchase Price Exc Tax","Total Purchase Price Inc Tax"]:

            total = df[col].apply(safe_int).sum()

            total_row.append(Paragraph(f"{total:,}",meta_style))

        else:

            total_row.append(Paragraph("TOTAL" if len(total_row)==0 else "",meta_style))

    total_table = Table([total_row],colWidths=[col_width]*len(headers))

    total_table.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.25,colors.black),
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#eeeeee")),
        ("FONTSIZE",(0,0),(-1,-1),6)
    ]))

    elements.append(Spacer(1,6))
    elements.append(total_table)

    elements.append(Spacer(1,20))

    if os.path.exists(sign):

        sign_img = Image(sign,width=90,height=45)

    else:

        sign_img = Paragraph("(TTD)",sign_style)

    sig = Table([
        ["Dibuat Oleh","Disetujui Oleh"],
        ["",""],
        [sign_img,""],
        [Paragraph("ITA",sign_style),Paragraph("Supplier / Distributor",sign_style)]
    ],colWidths=[300,300])

    sig.setStyle(TableStyle([
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE")
    ]))

    elements.append(sig)

    doc.build(elements,canvasmaker=NumberedCanvas)

    print("Selesai")

# =========================
# MAIN
# =========================

def main():

    print("=== CONSIGNMENT PDF GENERATOR ===")

    bulan = input("Report bulan (contoh Jan 26): ").strip()

    bulan_file = bulan.replace(" ","")

    filename = f"CONSIGNMENT_{bulan_file}_to_pdf.xlsx"

    Tk().withdraw()

    folder = filedialog.askdirectory(title="Pilih folder")

    excel_path = os.path.join(folder,filename)

    if not os.path.exists(excel_path):

        print("File Excel tidak ditemukan")
        return

    print("Membaca Excel...")

    df = pd.read_excel(excel_path)

    df.columns = df.columns.str.strip()

    print("Total rows :",len(df))

    # =========================
    # COLUMN MAPPING
    # =========================

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

    missing = [c for c in mapping if c not in df.columns]

    if missing:

        print("Kolom tidak ditemukan :")

        for m in missing:

            print(m)

        return

    df = df[list(mapping.keys())].rename(columns=mapping)

    df = df.sort_values(["Vendor name","Cabang","Item Code"])

    headers = [
        "Vendor name",
        "Cabang",
        "Item Code",
        "Item Name",
        "Date",
        "Qty Sold",
        "Cost Price",
        "Total Purchase Price Exc Tax",
        "Total Purchase Price Inc Tax"
    ]

    periode = get_period(df)

    output_dir = os.path.join(folder,f"output_{bulan_file}")

    os.makedirs(output_dir,exist_ok=True)

    groups = df.groupby(["KolomU","Vendor name"])

    print("Total vendor report :",len(groups))

    for (pt,vendor),data in groups:

        file_name = f"{safe_filename(vendor)}_{bulan_file}_{safe_filename(pt)}.pdf"

        output_path = os.path.join(output_dir,file_name)

        generate_pdf(
            output_path,
            vendor,
            periode,
            pt,
            data[headers].fillna(""),
            headers,
            os.path.join(folder,LOGO_FILENAME),
            os.path.join(folder,SIGN_FILENAME)
        )

    print("\nSEMUA REPORT SELESAI")

# =========================
# RUN
# =========================

if __name__ == "__main__":

    try:

        main()

    except Exception as e:

        print("ERROR :",e)

    finally:

        hold_console()