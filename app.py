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
from openpyxl import Workbook

# =========================
# UI CONFIG
# =========================
st.set_page_config(page_title="THL TO SM", layout="centered")

# =========================
# CSS (GIỮ NGUYÊN)
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
# 🔥 FIX TRIỆT ĐỂ OPENPYXL ERROR
# =========================
def safe_load(path, read_only=False):
    """
    FIX:
    - expected Fill
    - styles corruption
    - index out of range
    """

    try:
        return load_workbook(
            path,
            read_only=read_only,
            data_only=True,
            keep_links=False
        )

    except Exception:
        # =========================
        # FALLBACK CLEAN REBUILD
        # =========================
        df = pd.read_excel(path, dtype=str)

        tmp = os.path.join(
            tempfile.gettempdir(),
            f"safe_{uuid.uuid4().hex}.xlsx"
        )

        wb = Workbook()
        ws = wb.active

        for row in df.itertuples(index=False):
            ws.append(list(row))

        wb.save(tmp)
        wb.close()

        return load_workbook(
            tmp,
            read_only=read_only,
            data_only=True
        )


# =========================
# FIND COLUMN SHIPMENT
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

    ready = uploaded_files and len(uploaded_files) == 2

    if ready and not st.session_state["done"]:

        if st.button("🚀 Bắt đầu xử lý"):

            try:
                with st.spinner("⏳ Đang xử lý..."):

                    tmp_dir = tempfile.gettempdir()
                    path_tpn, path_book1 = None, None

                    # =========================
                    # phân loại file
                    # =========================
                    for file in uploaded_files:
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
                    # extract numbers
                    # =========================
                    df = pd.read_excel(path_book1, usecols=[0], dtype=str)
                    all_numbers = set()

                    for v in df.iloc[:, 0].dropna():
                        for num in re.findall(r"\d+", str(v)):
                            if len(num) == 3:
                                num = "0" + num
                            if len(num) == 4:
                                all_numbers.add(num)

                    # =========================
                    # process TPN
                    # =========================
                    wb = safe_load(path_tpn)
                    ws = wb.active
                    col_index = find_shipment_col(ws)

                    yellow = PatternFill("solid", fgColor="FFFF00")
                    header_fill = PatternFill("solid", fgColor="000080")
                    header_font = Font(color="FFFFFF", bold=True)

                    for cell in ws[1]:
                        cell.fill = header_fill
                        cell.font = header_font

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

                    # =========================
                    # kế hoạch file
                    # =========================
                    df2 = pd.read_excel(path_book1, header=None, dtype=str)

                    workbook = xlsxwriter.Workbook(kehoach_path)
                    worksheet = workbook.add_worksheet()

                    red = workbook.add_format({'font_color': 'red'})
                    normal = workbook.add_format({})

                    col_width = 0

                    for r, row in df2.iterrows():
                        text = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])
                        col_width = max(col_width, len(text))

                        parts = []
                        last = 0

                        for m in re.finditer(r"\d+", text):
                            num = m.group()
                            start, end = m.span()
                            check = "0"+num if len(num) == 3 else num

                            if start > last:
                                parts += [normal, text[last:start]]

                            parts += [red if check in ketqua_numbers else normal, num]
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
                    # zip output
                    # =========================
                    zip_path = os.path.join(tmp_dir, "TPN_COMPLETE.zip")

                    with zipfile.ZipFile(zip_path, "w") as z:
                        z.write(save_path, "TPN_KET_QUA.xlsx")
                        z.write(kehoach_path, "TPN_KE_HOACH_XE.xlsx")

                    with open(zip_path, "rb") as f:
                        zip_data = f.read()

                st.success(f"✅ COMPLETE !!! Matched: {count}")

                b64 = base64.b64encode(zip_data).decode()
                st.components.v1.html(f"""
                    <a id="dl" href="data:application/zip;base64,{b64}" download="THL_TO_SM.zip"></a>
                    <script>document.getElementById('dl').click();</script>
                """, height=0)

                st.session_state["done"] = True

            except Exception as e:
                st.error(f"❌ Có lỗi xảy ra: {e}")

    if st.session_state["done"]:
        if st.button("🔄 Xử lý file mới"):
            st.session_state["uploader_key"] += 1
            st.session_state["done"] = False
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)
