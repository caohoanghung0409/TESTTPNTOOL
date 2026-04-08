import streamlit as st
import pandas as pd
import re
import tempfile
import os
import zipfile
import xlsxwriter
import base64

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

st.set_page_config(page_title="THL TO SM", layout="centered")

# =========================
# UI (GIỮ NGUYÊN)
# =========================
st.markdown("""
<style>
header{display:none!important;}
#MainMenu{visibility:hidden;}
footer{visibility:hidden;}
.block-container{padding-top:0rem!important;}

.header{text-align:center;padding:8px 0;}
.header h1{color:#0284c7;margin:0;}
.header p{color:#64748b;margin:0;}

.card{background:white;padding:20px;border-radius:12px;}

.stButton>button{
    width:100%;
    height:42px;
    border-radius:10px;
    background:linear-gradient(90deg,#0ea5e9,#22c55e);
    color:white;
}
</style>
""", unsafe_allow_html=True)

# =========================
# COLOR
# =========================
PALETTE = ["FFF2CC","FCE4D6","E2EFDA","DDEBF7","E4DFEC","F8CBAD","C6E0B4"]

# =========================
# SAFE READ (KHÔNG OPENPYXL INPUT)
# =========================
def read_excel(path):
    return pd.read_excel(path, header=None, dtype=str)

# =========================
# BUILD CLEAN FILE (QUAN TRỌNG)
# =========================
def rebuild_clean_excel(df, path):
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet()

    for r, row in df.iterrows():
        for c, val in enumerate(row):
            ws.write(r, c, "" if pd.isna(val) else str(val))

    wb.close()

# =========================
# UI
# =========================
st.title("⚡ THL TO SM")

files = st.file_uploader("Upload 2 file Excel", type=["xlsx"], accept_multiple_files=True)

if files and len(files) == 2:

    if st.button("RUN"):

        tmp = tempfile.gettempdir()

        path_tpn = None
        path_book = None

        # save
        paths = []
        for f in files:
            p = os.path.join(tmp, f.name)
            with open(p, "wb") as out:
                out.write(f.read())
            paths.append(p)

        # classify
        for p in paths:
            dfh = read_excel(p)
            if any("Shipment Nbr" in str(x) for x in dfh.iloc[0]):
                path_tpn = p
            else:
                path_book = p

        # =========================
        # BOOK
        # =========================
        df_book = read_excel(path_book)

        row_numbers = []
        all_numbers = set()

        for _, r in df_book.iterrows():
            txt = "" if pd.isna(r.iloc[0]) else str(r.iloc[0])

            nums = set()
            for n in re.findall(r"\d+", txt):
                if len(n) == 3:
                    n = "0" + n
                if len(n) == 4:
                    nums.add(n)

            row_numbers.append(nums)
            all_numbers.update(nums)

        row_color = {i: PALETTE[i % len(PALETTE)] for i in range(len(row_numbers))}

        # =========================
        # REBUILD TPN CLEAN FILE (FIX TRIỆT ĐỂ)
        # =========================
        df_tpn = read_excel(path_tpn)
        clean_tpn = os.path.join(tmp, "clean_tpn.xlsx")
        rebuild_clean_excel(df_tpn, clean_tpn)

        # =========================
        # OPEN CLEAN FILE
        # =========================
        wb = load_workbook(clean_tpn)
        ws = wb.active

        # find shipment col
        col = None
        for c in range(1, ws.max_column + 1):
            if ws.cell(1, c).value and "Shipment Nbr" in str(ws.cell(1, c).value):
                col = c
                break

        if not col:
            st.error("Không tìm thấy Shipment Nbr")
            st.stop()

        # header style
        for c in range(1, ws.max_column + 1):
            ws.cell(1, c).fill = PatternFill("solid", fgColor="1F4E79")
            ws.cell(1, c).font = Font(color="FFFFFF", bold=True)

        # =========================
        # KET_QUA COLOR
        # =========================
        for i in range(2, ws.max_row + 1):

            val = ws.cell(i, col).value
            if not val:
                continue

            nums = set()
            for n in re.findall(r"\d+", str(val)):
                if len(n) == 3:
                    n = "0" + n
                if len(n) == 4:
                    nums.add(n)

            match = nums & all_numbers
            if not match:
                continue

            chosen = None
            for idx, rset in enumerate(row_numbers):
                if match & rset:
                    chosen = idx
                    break

            color = row_color[chosen] if chosen is not None else "DDDDDD"

            ws.cell(i, col).fill = PatternFill("solid", fgColor=color)

        out1 = os.path.join(tmp, "TPN_KET_QUA.xlsx")
        wb.save(out1)
        wb.close()

        # =========================
        # KEHOACH FILE
        # =========================
        df2 = read_excel(path_book)

        out2 = os.path.join(tmp, "TPN_KE_HOACH_XE.xlsx")
        wb2 = xlsxwriter.Workbook(out2)
        ws2 = wb2.add_worksheet()

        normal = wb2.add_format()
        red = wb2.add_format({"font_color": "#E74C3C"})

        for r, row in df2.iterrows():
            text = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])

            parts = []
            last = 0

            for m in re.finditer(r"\d+", text):
                s, e = m.span()

                if s > last:
                    parts += [normal, text[last:s]]

                num = m.group()
                check = num if len(num) != 3 else "0" + num

                if len(check) == 4 and check in all_numbers:
                    parts += [red, num]
                else:
                    parts += [normal, num]

                last = e

            if last < len(text):
                parts += [normal, text[last:]]

            try:
                ws2.write_rich_string(r, 0, *parts)
            except:
                ws2.write(r, 0, text)

        wb2.close()

        # =========================
        # ZIP
        # =========================
        zip_path = os.path.join(tmp, "RESULT.zip")

        with zipfile.ZipFile(zip_path, "w") as z:
            z.write(out1, "TPN_KET_QUA.xlsx")
            z.write(out2, "TPN_KE_HOACH_XE.xlsx")

        with open(zip_path, "rb") as f:
            data = f.read()

        st.success("DONE")

        b64 = base64.b64encode(data).decode()

        st.components.v1.html(f"""
            <a id="dl" href="data:application/zip;base64,{b64}" download="THL_TO_SM.zip"></a>
            <script>document.getElementById('dl').click();</script>
        """, height=0)
