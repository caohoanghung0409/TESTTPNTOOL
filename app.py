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
from openpyxl.styles import PatternFill
from openpyxl.worksheet.views import Selection

st.set_page_config(page_title="THL TO SM", layout="centered")

# =========================
# CSS
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
# STATE
# =========================
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

if "processing" not in st.session_state:
    st.session_state["processing"] = False

if "done" not in st.session_state:
    st.session_state["done"] = False

if "last_file_hash" not in st.session_state:
    st.session_state["last_file_hash"] = None


# =========================
# AUTO DOWNLOAD
# =========================
def auto_download(data, filename):
    b64 = base64.b64encode(data).decode()
    href = f"""
    <html>
    <body>
    <a id="download_link" href="data:application/zip;base64,{b64}" download="{filename}"></a>
    <script>
    document.getElementById('download_link').click();
    </script>
    </body>
    </html>
    """
    st.components.v1.html(href, height=0)


# =========================
# FIND COLUMN
# =========================
def find_shipment_col(ws):
    for cell in ws[1]:
        if cell.value:
            v = str(cell.value).replace("\xa0", " ").strip()
            if "Shipment Nbr" in v:
                return cell.column
    return None


# =========================
# UI HEADER
# =========================
st.markdown("""
<div class="header">
    <h1>⚡ THL TO SM</h1>
    <p>Xử lý & đối soát Shipment nhanh chóng</p>
</div>
""", unsafe_allow_html=True)

with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "📂 Chọn 2 file Excel",
        type=["xlsx"],
        accept_multiple_files=True,
        key=f"uploader_{st.session_state['uploader_key']}"
    )

    current_hash = None
    if uploaded_files:
        current_hash = "|".join(sorted([f.name for f in uploaded_files]))

    if current_hash != st.session_state["last_file_hash"]:
        st.session_state["done"] = False
        st.session_state["processing"] = False
        st.session_state["last_file_hash"] = current_hash

    ready = uploaded_files and len(uploaded_files) == 2
    can_run = ready and (not st.session_state["processing"]) and (not st.session_state["done"])

    # =========================
    # BUTTON RUN (ẨN KHI DONE)
    # =========================
    if ready and not st.session_state["done"]:
        if st.button("🚀 Bắt đầu xử lý", disabled=not can_run):

            st.session_state["processing"] = True
            st.session_state["done"] = False

            try:
                with st.spinner("⏳ Đang xử lý..."):

                    tmp_dir = tempfile.gettempdir()
                    path_tpn = None
                    path_book1 = None

                    for file in uploaded_files:
                        path = os.path.join(tmp_dir, file.name)

                        with open(path, "wb") as f:
                            f.write(file.read())

                        wb_check = load_workbook(path, read_only=True)
                        ws_check = wb_check.active
                        header = [str(c.value).strip() if c.value else "" for c in ws_check[1]]
                        wb_check.close()

                        if any("Shipment Nbr" in h for h in header):
                            path_tpn = path
                        else:
                            path_book1 = path

                    if not path_tpn or not path_book1:
                        st.error("❌ Không đúng định dạng 2 file!")
                        st.stop()

                    save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
                    kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

                    df = pd.read_excel(path_book1, usecols=[0], engine="openpyxl", dtype=str)

                    all_numbers = set()
                    for v in df.iloc[:, 0].dropna():
                        for num in re.findall(r"\d+", str(v)):
                            if len(num) == 3:
                                num = "0" + num
                            if len(num) == 4:
                                all_numbers.add(num)

                    wb = load_workbook(path_tpn)
                    ws = wb.active

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

                    ws.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]
                    ws.sheet_view.topLeftCell = "A1"

                    wb.save(save_path)
                    wb.close()

                    # FILE 2
                    df2 = pd.read_excel(path_book1, header=None, engine="openpyxl", dtype=str)

                    workbook = xlsxwriter.Workbook(kehoach_path)
                    worksheet = workbook.add_worksheet()

                    red_format = workbook.add_format({'font_color': 'red'})
                    normal_format = workbook.add_format({})

                    max_len = 0

                    for row_idx, row in df2.iterrows():
                        cell_value = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])
                        max_len = max(max_len, len(cell_value))

                        parts = []
                        last_idx = 0

                        for match in re.finditer(r"\d+", cell_value):
                            num = match.group()
                            start, end = match.span()

                            num_check = "0" + num if len(num) == 3 else num

                            if start > last_idx:
                                parts += [normal_format, cell_value[last_idx:start]]

                            if len(num_check) == 4 and num_check in ketqua_numbers:
                                parts += [red_format, num]
                            else:
                                parts += [normal_format, num]

                            last_idx = end

                        if last_idx < len(cell_value):
                            parts += [normal_format, cell_value[last_idx:]]

                        try:
                            worksheet.write_rich_string(row_idx, 0, *parts)
                        except:
                            worksheet.write(row_idx, 0, cell_value)

                    worksheet.set_column(0, 0, max_len + 3)
                    workbook.close()

                    # ZIP
                    zip_path = os.path.join(tmp_dir, "TPN_COMPLETE.zip")

                    with zipfile.ZipFile(zip_path, "w") as z:
                        z.write(save_path, "TPN_KET_QUA.xlsx")
                        z.write(kehoach_path, "TPN_KE_HOACH_XE.xlsx")

                    with open(zip_path, "rb") as f:
                        zip_data = f.read()

                st.success(f"✅ COMPLETE !!! Matched: {count}")

                auto_download(zip_data, "THL_TO_SM.zip")

                st.session_state["done"] = True
                st.session_state["processing"] = False

            except Exception:
                st.session_state["processing"] = False
                st.error("❌ Có lỗi xảy ra!")

    # =========================
    # BUTTON RESET
    # =========================
    if st.session_state["done"]:
        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("🔄 Xử lý file mới", use_container_width=True):
            st.session_state["uploader_key"] += 1
            st.session_state["done"] = False
            st.session_state["processing"] = False
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)
