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
# STYLE
# =========================
st.markdown("""
<style>
header {display:none;}
#MainMenu {visibility:hidden;}
footer {visibility:hidden;}

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
COLOR_PALETTE = [
    "FFFF3B30","FF34C759","FF007AFF","FFFF9500",
    "FFAF52DE","FFFF2D55","FF5856D6","FF00C7BE",
    "FFFFCC00","FFFF6B00","FF4CD964","FF1E90FF"
]

def load_safe(path):
    return load_workbook(path, data_only=True, keep_links=False)

def find_col(ws):
    for c in ws[1]:
        if c.value and "Shipment Nbr" in str(c.value):
            return c.column
    return None

# =========================
# UI
# =========================
st.title("⚡ THL TO SM")

files = st.file_uploader(
    "Upload 2 file Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

# =========================
# RUN
# =========================
if files and len(files) == 2 and st.button("RUN"):

    try:
        tmp = tempfile.gettempdir()

        file_tpn = None
        file_book = None

        # ================= CLASSIFY FILE =================
        for f in files:
            path = os.path.join(tmp, f.name)
            with open(path, "wb") as out:
                out.write(f.read())

            wb = load_workbook(path, data_only=True)
            header = [str(x.value) for x in wb.active[1]]
            wb.close()

            if any("Shipment Nbr" in h for h in header):
                file_tpn = path
            else:
                file_book = path

        # ================= BOOK DATA =================
        df = pd.read_excel(file_book, usecols=[0], dtype=str)

        all_numbers = set()
        for v in df.iloc[:, 0].dropna():
            for n in re.findall(r"\d+", str(v)):
                if len(n) == 3:
                    n = "0" + n
                if len(n) == 4:
                    all_numbers.add(n)

        # ================= LOAD TPN =================
        wb = load_safe(file_tpn)
        ws = wb.active

        col = find_col(ws)
        if col is None:
            st.error("Không tìm thấy Shipment Nbr")
            st.stop()

        header_fill = PatternFill("solid", fgColor="FF000080")

        for c in ws[1]:
            c.fill = header_fill
            c.font = Font(color="FFFFFFFF", bold=True)

        number_color = {}
        color_idx = 0

        matched_set = set()

        # ================= PROCESS ROWS =================
        for r in range(2, ws.max_row + 1):

            val = ws.cell(r, col).value
            if not val:
                continue

            nums = set()
            for n in re.findall(r"\d+", str(val)):
                if len(n) == 3:
                    n = "0" + n
                if len(n) == 4:
                    nums.add(n)

            match = nums & all_numbers

            if match:
                matched_set.update(match)

                # assign color per shipment
                for m in match:
                    if m not in number_color:
                        number_color[m] = COLOR_PALETTE[color_idx % len(COLOR_PALETTE)]
                        color_idx += 1

                first = list(match)[0]
                ws.cell(r, col).fill = PatternFill(
                    "solid",
                    fgColor=number_color[first]
                )

        # ================= SAVE RESULT =================
        out_path = os.path.join(tmp, "TPN_KET_QUA.xlsx")
        wb.save(out_path)
        wb.close()

        # ================= MATCH COUNT (FIXED) =================
        st.success(f"🎯 MATCH SHIPMENT = {len(matched_set)}")

        # ================= KEHOACH FILE (FIXED LOGIC) =================
        df2 = pd.read_excel(file_book, header=None, dtype=str)

        kehoach_path = os.path.join(tmp, "TPN_KE_HOACH_XE.xlsx")

        wb2 = xlsxwriter.Workbook(kehoach_path)
        ws2 = wb2.add_worksheet()

        normal = wb2.add_format({})
        red = wb2.add_format({'font_color': 'red'})

        max_len = 0

        for i, row in df2.iterrows():

            text = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])
            max_len = max(max_len, len(text))

            parts = []
            last = 0

            # ONLY highlight if number EXISTS in matched_set
            for m in re.finditer(r"\d+", text):
                start, end = m.span()
                num = m.group()

                if len(num) == 3:
                    num = "0" + num

                if start > last:
                    parts += [normal, text[last:start]]

                if num in matched_set:
                    parts += [red, m.group()]
                else:
                    parts += [normal, m.group()]

                last = end

            if last < len(text):
                parts += [normal, text[last:]]

            try:
                ws2.write_rich_string(i, 0, *parts)
            except:
                ws2.write(i, 0, text)

        ws2.set_column(0, 0, max_len + 3)
        wb2.close()

        # ================= ZIP =================
        zip_path = os.path.join(tmp, "THL_TO_SM.zip")

        with zipfile.ZipFile(zip_path, "w") as z:
            z.write(out_path, "TPN_KET_QUA.xlsx")
            z.write(kehoach_path, "TPN_KE_HOACH_XE.xlsx")

        with open(zip_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()

        st.markdown(f"""
        <a download="THL_TO_SM.zip" href="data:application/zip;base64,{b64}">
            👉 DOWNLOAD FILE
        </a>
        """, unsafe_allow_html=True)

    except Exception as e:
        st.error("❌ Lỗi xảy ra")
        st.exception(e)
