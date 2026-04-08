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
import traceback

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

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
if "done" not in st.session_state:
    st.session_state["done"] = False

# =========================
# PALETTE (PASTEL - dễ nhìn)
# =========================
PALETTE = [
    "FFB3BA","FFDFBA","FFFFBA","BAFFC9","BAE1FF",
    "E3BAFF","FFC6FF","C6FFFA","F1FFC6","D6D6D6",
    "FFD6A5","Caffbf","9bf6ff","a0c4ff","bdb2ff"
]

# =========================
# FIX EXCEL STYLE
# =========================
def safe_load(path):
    try:
        return load_workbook(path)
    except:
        return load_workbook(path)

def find_shipment_col(ws):
    for cell in ws[1]:
        if cell.value and "Shipment Nbr" in str(cell.value):
            return cell.column
    return None

# =========================
# UI
# =========================
st.markdown("""
<div class="header">
    <h1>⚡ THL TO SM</h1>
    <p>Xử lý & đối soát Shipment</p>
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

    ready = uploaded_files and len(uploaded_files) == 2

    if ready and not st.session_state["done"]:
        if st.button("🚀 Bắt đầu xử lý"):

            try:
                with st.spinner("⏳ Đang xử lý..."):

                    tmp_dir = tempfile.gettempdir()
                    path_tpn, path_book1 = None, None

                    # =========================
                    # LOAD FILES
                    # =========================
                    for file in uploaded_files:
                        path = os.path.join(tmp_dir, file.name)
                        with open(path, "wb") as f:
                            f.write(file.read())

                        wb = safe_load(path)
                        header = [str(c.value) for c in wb.active[1]]
                        wb.close()

                        if any("Shipment Nbr" in str(h) for h in header):
                            path_tpn = path
                        else:
                            path_book1 = path

                    # =========================
                    # BOOK1 (KEHOACH) - lấy số theo từng dòng
                    # =========================
                    df_book1 = pd.read_excel(path_book1, header=None, dtype=str)

                    book1_rows = []
                    all_numbers = set()

                    for _, row in df_book1.iterrows():
                        text = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])

                        nums = set()
                        for n in re.findall(r"\d+", text):
                            if len(n) == 3:
                                n = "0" + n
                            if len(n) == 4:
                                nums.add(n)

                        book1_rows.append(nums)
                        all_numbers.update(nums)

                    # =========================
                    # MAP COLOR THEO ROW (QUAN TRỌNG FIX)
                    # =========================
                    row_color = {}
                    for i in range(len(book1_rows)):
                        row_color[i] = PALETTE[i % len(PALETTE)]

                    # =========================
                    # LOAD TPN
                    # =========================
                    wb = safe_load(path_tpn)
                    ws = wb.active
                    col_index = find_shipment_col(ws)

                    if col_index is None:
                        st.error("Không tìm thấy cột Shipment Nbr")
                        st.stop()

                    # header style
                    header_fill = PatternFill("solid", fgColor="FF1F4E79")
                    for c in ws[1]:
                        c.fill = header_fill
                        c.font = Font(color="FFFFFFFF", bold=True)

                    # =========================
                    # TÔ MÀU TPN_KET_QUA (FIX)
                    # =========================
                    for i in range(2, ws.max_row + 1):
                        val = ws.cell(i, col_index).value
                        if not val:
                            continue

                        nums = set()
                        for n in re.findall(r"\d+", str(val)):
                            if len(n) == 3:
                                n = "0" + n
                            if len(n) == 4:
                                nums.add(n)

                        match_nums = nums & all_numbers

                        if not match_nums:
                            continue

                        # tìm row bên book1 chứa các số này
                        chosen_row = None
                        for idx, rset in enumerate(book1_rows):
                            if len(match_nums & rset) > 0:
                                chosen_row = idx
                                break

                        if chosen_row is not None:
                            color = row_color[chosen_row]
                        else:
                            color = "D9D9D9"

                        ws.cell(i, col_index).fill = PatternFill(
                            "solid",
                            fgColor=color
                        )

                    save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
                    wb.save(save_path)
                    wb.close()

                    # =========================
                    # FILE KEHOACH (FIX CHỈ TÔ SỐ TRÙNG)
                    # =========================
                    df2 = pd.read_excel(path_book1, header=None, dtype=str)

                    kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")
                    workbook = xlsxwriter.Workbook(kehoach_path)
                    worksheet = workbook.add_worksheet()

                    normal = workbook.add_format({})
                    red = workbook.add_format({'font_color': '#FF3B30'})

                    col_width = 0

                    for r, row in df2.iterrows():
                        text = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])
                        col_width = max(col_width, len(text))

                        parts = []
                        last = 0

                        for m in re.finditer(r"\d+", text):
                            start, end = m.span()

                            if start > last:
                                parts += [normal, text[last:start]]

                            num = m.group()
                            check = num if len(num) != 3 else "0" + num

                            # CHỈ TÔ NẾU TRÙNG
                            if len(check) == 4 and check in all_numbers:
                                parts += [red, num]
                            else:
                                parts += [normal, num]

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
                    # ZIP
                    # =========================
                    zip_path = os.path.join(tmp_dir, "TPN_COMPLETE.zip")

                    with zipfile.ZipFile(zip_path, "w") as z:
                        z.write(save_path, "TPN_KET_QUA.xlsx")
                        z.write(kehoach_path, "TPN_KE_HOACH_XE.xlsx")

                    with open(zip_path, "rb") as f:
                        zip_data = f.read()

                st.success("✅ COMPLETE !!!")

                b64 = base64.b64encode(zip_data).decode()

                st.components.v1.html(f"""
                    <a id="dl" href="data:application/zip;base64,{b64}" download="THL_TO_SM.zip"></a>
                    <script>document.getElementById('dl').click();</script>
                """, height=0)

                st.session_state["done"] = True

            except Exception as e:
                st.error("❌ Có lỗi xảy ra:")
                st.exception(e)

    if st.session_state["done"]:
        if st.button("🔄 Xử lý file mới"):
            st.session_state["uploader_key"] += 1
            st.session_state["done"] = False
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)
