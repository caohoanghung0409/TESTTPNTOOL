import streamlit as st
import pandas as pd
import re
import tempfile
import os
import zipfile
import uuid
import xlsxwriter

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.worksheet.views import Selection

st.set_page_config(page_title="TPN TOOL ⚡", layout="centered")

# =========================
# CSS UI
# =========================
st.markdown("""
<style>
header {display: none !important;}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.block-container {
    padding-top: 0rem !important;
}

html, body {
    background-color: #f1f5f9;
}

.header {
    text-align: center;
    padding: 10px 0;
}

.header h1 {
    color: #0284c7;
    margin: 0;
}

.header p {
    color: #64748b;
    margin: 0;
}

.card {
    background: white;
    padding: 20px;
    border-radius: 12px;
}

.stButton>button {
    width: 100%;
    height: 42px;
    border-radius: 10px;
    background: linear-gradient(90deg, #0ea5e9, #22c55e);
    color: white;
}

.stDownloadButton>button {
    width: 100%;
    height: 42px;
    border-radius: 10px;
    background: #16a34a;
    color: white;
}

section[data-testid="stFileUploader"] {
    border: 2px dashed #cbd5f5;
    padding: 12px;
    border-radius: 10px;
    background: #f8fafc;
}
</style>
""", unsafe_allow_html=True)

# =========================
# STATE
# =========================
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

# =========================
# HEADER
# =========================
st.markdown("""
<div class="header">
    <h1>⚡ TPN TOOL</h1>
    <p>Xử lý & đối soát Shipment nhanh chóng</p>
</div>
""", unsafe_allow_html=True)

# =========================
# SAFE LOAD (FIX STYLE ERROR)
# =========================
def safe_load(path, read_only=False):
    try:
        return load_workbook(path, read_only=read_only, data_only=True)
    except Exception:
        # fallback: đọc bằng pandas rồi save lại
        df = pd.read_excel(path, engine="openpyxl")
        fixed_path = path.replace(".xlsx", "_clean.xlsx")
        df.to_excel(fixed_path, index=False)
        return load_workbook(fixed_path, read_only=read_only, data_only=True)

# =========================
# FIND COLUMN
# =========================
def find_shipment_col(ws):
    for cell in ws[1]:
        if cell.value and "Shipment Nbr" in str(cell.value):
            return cell.column
    return None

# =========================
# UI MAIN
# =========================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "📂 Chọn 2 file Excel",
        type=["xlsx"],
        accept_multiple_files=True,
        key=f"uploader_{st.session_state['uploader_key']}"
    )

    if st.button("🚀 Bắt đầu xử lý"):

        if not uploaded_files or len(uploaded_files) != 2:
            st.error("⚠️ Chọn đúng 2 file")
            st.stop()

        with st.spinner("⏳ Đang xử lý..."):

            tmp_dir = tempfile.gettempdir()
            path_tpn, path_book1 = None, None

            for file in uploaded_files:
                path = os.path.join(tmp_dir, file.name)
                with open(path, "wb") as f:
                    f.write(file.read())

                wb_check = safe_load(path, read_only=True)
                header = [str(c.value) for c in wb_check.active[1]]
                wb_check.close()

                if any("Shipment Nbr" in h for h in header):
                    path_tpn = path
                else:
                    path_book1 = path

            save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
            kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

            # ===== FILE 2 READ =====
            df = pd.read_excel(path_book1, usecols=[0])

            all_numbers = set()
            for v in df.iloc[:, 0].dropna().astype(str):
                for num in re.findall(r"\d+", v):
                    if len(num) == 3:
                        num = "0" + num
                    if len(num) == 4:
                        all_numbers.add(num)

            # ===== FILE 1 =====
            wb = safe_load(path_tpn)
            ws = wb.active

            ws.sheet_view.topLeftCell = "A1"
            ws.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]

            col_index = find_shipment_col(ws)

            yellow = PatternFill("solid", fgColor="FFFF00")
            ketqua_numbers = set()
            count = 0

            for i in range(2, ws.max_row + 1):
                val = ws.cell(i, col_index).value
                if val:
                    nums = set()
                    for num in re.findall(r"\d+", str(val)):
                        if len(num) == 3:
                            num = "0" + num
                        if len(num) == 4:
                            nums.add(num)

                    ketqua_numbers.update(nums)

                    if nums & all_numbers:
                        ws.cell(i, col_index).fill = yellow
                        count += 1

            wb.save(save_path)
            wb.close()

            # ===== FILE 2 (HIGHLIGHT + AUTO WIDTH) =====
            df2 = pd.read_excel(path_book1, header=None)

            workbook = xlsxwriter.Workbook(kehoach_path)
            worksheet = workbook.add_worksheet()

            red_format = workbook.add_format({'font_color': 'red'})
            normal_format = workbook.add_format({})

            max_len = 0

            for row_idx, row in df2.iterrows():
                cell_value = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
                max_len = max(max_len, len(cell_value))

                parts = []
                last_idx = 0

                for match in re.finditer(r"\d+", cell_value):
                    num = match.group()
                    num_check = "0" + num if len(num) == 3 else num

                    start, end = match.span()

                    if start > last_idx:
                        parts += [normal_format, cell_value[last_idx:start]]

                    if len(num_check) == 4 and num_check in ketqua_numbers:
                        parts += [red_format, num]
                    else:
                        parts += [normal_format, num]

                    last_idx = end

                if last_idx < len(cell_value):
                    parts += [normal_format, cell_value[last_idx:]]

                if parts:
                    worksheet.write_rich_string(row_idx, 0, *parts)
                else:
                    worksheet.write(row_idx, 0, cell_value)

            worksheet.set_column(0, 0, max_len + 5)

            workbook.close()

            # ===== ZIP =====
            zip_path = os.path.join(tmp_dir, "TPN_COMPLETE.zip")
            with zipfile.ZipFile(zip_path, "w") as z:
                z.write(save_path, "TPN_KET_QUA.xlsx")
                z.write(kehoach_path, "TPN_KE_HOACH_XE.xlsx")

            with open(zip_path, "rb") as f:
                zip_data = f.read()

        st.success(f"✅ COMPLETE !!! Matched: {count}")

        st.download_button(
            "📥 Download ALL (ZIP)",
            data=zip_data,
            file_name="TPN_COMPLETE.zip"
        )

        st.session_state["uploader_key"] += 1

    st.markdown('</div>', unsafe_allow_html=True)
