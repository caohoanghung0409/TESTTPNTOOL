import streamlit as st
import pandas as pd
import re
import tempfile
import os
import zipfile
import shutil
import uuid
import xlsxwriter
import base64

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.worksheet.views import Selection

st.set_page_config(page_title="THL TO SM", layout="centered")

# =========================
# UI (GIỮ NGUYÊN)
# =========================
st.markdown("""
<style>
header {display: none !important;}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
.block-container {padding-top: 0rem !important;}

.header {text-align: center; padding: 8px 0;}
.header h1 {color: #0284c7; margin: 0;}
.header p {color: #64748b; margin: 0;}

.card {background: white; padding: 20px; border-radius: 12px;}

.stButton>button {
    width: 100%;
    height: 42px;
    border-radius: 10px;
    background: linear-gradient(90deg, #0ea5e9, #22c55e);
    color: white;
}
</style>
""", unsafe_allow_html=True)


# =========================
# SAFE LOAD
# =========================
def safe_load(path, read_only=False):
    try:
        return load_workbook(path, read_only=read_only, data_only=True, keep_links=False)
    except:
        return load_workbook(path, read_only=read_only, data_only=True, keep_links=False)


def find_shipment_col(ws):
    for cell in ws[1]:
        if cell.value and "Shipment Nbr" in str(cell.value):
            return cell.column
    return None


# =========================
# UI
# =========================
st.title("⚡ THL TO SM")

files = st.file_uploader(
    "Chọn 2 file Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

if files and len(files) == 2:

    if st.button("🚀 Xử lý"):

        tmp_dir = tempfile.gettempdir()
        path_tpn, path_book1 = None, None

        for file in files:
            path = os.path.join(tmp_dir, file.name)
            with open(path, "wb") as f:
                f.write(file.read())

            wb = safe_load(path, True)
            header = [str(c.value) for c in wb.active[1]]
            wb.close()

            if any("Shipment Nbr" in str(h) for h in header):
                path_tpn = path
            else:
                path_book1 = path

        save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
        kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

        # =========================
        # LẤY SỐ TỪ BOOK1
        # =========================
        df = pd.read_excel(path_book1, header=None, dtype=str)

        def norm(x):
            x = re.sub(r"\D", "", str(x))
            if len(x) == 3:
                x = "0" + x
            return x if len(x) == 4 else None

        # =========================
        # LOAD TPN
        # =========================
        wb = safe_load(path_tpn)
        ws = wb.active
        col_index = find_shipment_col(ws)

        header_fill = PatternFill("solid", fgColor="000080")
        header_font = Font(color="FFFFFF", bold=True)
        bold_font = Font(bold=True)

        for c in ws[1]:
            c.fill = header_fill
            c.font = header_font

        for r in ws.iter_rows(min_row=2):
            for c in r:
                if c.value:
                    c.font = bold_font

        # =========================
        # ⭐ IMPORTANT FIX:
        # KHÔNG BUILD GROUP MỚI
        # → lấy trực tiếp mapping từ KEHOACH file logic gốc
        # =========================
        group_sets = []
        palette = [
            "FFFF00", "FF9999", "99FF99", "9999FF",
            "FFD966", "FF66FF", "66FFFF", "C0C0C0"
        ]

        # build group đúng từ file (KHÔNG sửa logic gốc)
        for _, row in df.iterrows():
            text = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])
            nums = set()

            for n in re.findall(r"\d+", text):
                n = norm(n)
                if n:
                    nums.add(n)

            if nums:
                group_sets.append(nums)

        # =========================
        # FIX KET_QUA ONLY (AN TOÀN)
        # =========================
        for i in range(2, ws.max_row + 1):

            cell = ws.cell(i, col_index)
            val = cell.value

            if not val:
                continue

            nums = set()
            for n in re.findall(r"\d+", str(val)):
                n = norm(n)
                if n:
                    nums.add(n)

            matched = None

            for idx, g in enumerate(group_sets):
                if nums & g:
                    matched = palette[idx % len(palette)]
                    break

            if matched:
                cell.fill = PatternFill("solid", fgColor=matched)

        ws.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]

        wb.save(save_path)
        wb.close()

        # =========================
        # FILE KEHOACH (GIỮ NGUYÊN 100%)
        # =========================
        workbook = xlsxwriter.Workbook(kehoach_path)
        worksheet = workbook.add_worksheet()

        red = workbook.add_format({'font_color': 'red'})
        normal = workbook.add_format({})

        col_width = 0

        for r, row in df.iterrows():
            text = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])
            col_width = max(col_width, len(text))

            parts = []
            last = 0

            for m in re.finditer(r"\d+", text):
                num = m.group()
                start, end = m.span()
                check = norm(num)

                if start > last:
                    parts += [normal, text[last:start]]

                parts += [red if check and any(check in g for g in group_sets) else normal, num]
                last = end

            if last < len(text):
                parts += [normal, text[last:]]

            try:
                worksheet.write_rich_string(r, 0, *parts)
            except:
                worksheet.write(r, 0, text)

        worksheet.set_column(0, 0, col_width + 3)
        workbook.close()

        # =========================
        # ZIP + AUTO DOWNLOAD
        # =========================
        zip_path = os.path.join(tmp_dir, "TPN_COMPLETE.zip")

        with zipfile.ZipFile(zip_path, "w") as z:
            z.write(save_path, "TPN_KET_QUA.xlsx")
            z.write(kehoach_path, "TPN_KE_HOACH_XE.xlsx")

        with open(zip_path, "rb") as f:
            data = f.read()

        b64 = base64.b64encode(data).decode()

        st.success("DONE")

        st.components.v1.html(f"""
            <a id="dl" href="data:application/zip;base64,{b64}" download="THL_TO_SM.zip"></a>
            <script>
                document.getElementById('dl').click();
            </script>
        """, height=0)
