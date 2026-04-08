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
# UI STYLE
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
# NORMALIZE
# =========================
def norm(v):
    if v is None:
        return None
    s = str(v)

    if s.endswith(".0"):
        s = s[:-2]

    s = re.sub(r"\D", "", s)

    if len(s) == 3:
        s = "0" + s

    if len(s) >= 4:
        return s[:4]

    return None


# =========================
# SAFE LOAD
# =========================
def safe_load(path):
    try:
        return load_workbook(path, data_only=True)
    except:
        return load_workbook(path, data_only=True)


# =========================
# FIND COLUMN
# =========================
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


uploaded_files = st.file_uploader(
    "📂 Chọn 2 file Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files and len(uploaded_files) == 2:

    if st.button("🚀 Bắt đầu xử lý"):

        try:
            with st.spinner("Đang xử lý..."):

                tmp_dir = tempfile.gettempdir()
                path_tpn, path_kehoach = None, None

                # =========================
                # CLASSIFY FILE
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
                        path_kehoach = path

                save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
                kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

                # =========================
                # BUILD SET KE HOACH
                # =========================
                df = pd.read_excel(path_kehoach, dtype=str)

                all_numbers = set()
                for col in df.columns:
                    for v in df[col].dropna():
                        n = norm(v)
                        if n:
                            all_numbers.add(n)

                # =========================
                # PROCESS TPN
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
                    n = norm(val)

                    if n:
                        ketqua_numbers.add(n)

                        if n in all_numbers:
                            ws.cell(i, col_index).fill = yellow
                            count += 1

                ws.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]

                wb.save(save_path)
                wb.close()

                # =========================
                # CREATE KE HOACH FILE (FIX RICH STRING SAFE)
                # =========================
                df2 = pd.read_excel(path_kehoach, header=None, dtype=str)

                workbook = xlsxwriter.Workbook(kehoach_path)
                worksheet = workbook.add_worksheet()

                red = workbook.add_format({'font_color': 'red'})
                normal = workbook.add_format({})

                col_width = 0

                for r, row in df2.iterrows():
                    text = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])
                    col_width = max(col_width, len(text))

                    matches = list(re.finditer(r"\d+", text))

                    # =========================
                    # SAFE BUILD PARTS
                    # =========================
                    parts = []
                    last = 0

                    for m in matches:
                        raw = m.group()
                        n = norm(raw)

                        start, end = m.span()

                        if start > last:
                            parts.append(normal)
                            parts.append(text[last:start])

                        parts.append(red if (n and n in ketqua_numbers) else normal)
                        parts.append(raw)

                        last = end

                    if last < len(text):
                        parts.append(normal)
                        parts.append(text[last:])

                    # =========================
                    # FIX: write_rich_string SAFETY CHECK
                    # =========================
                    try:
                        if len(parts) >= 2:
                            worksheet.write_rich_string(r, 0, *parts)
                        else:
                            worksheet.write(r, 0, text)
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

            st.success(f"✅ DONE - Matched: {count}")

            b64 = base64.b64encode(zip_data).decode()
            st.components.v1.html(f"""
                <a id="dl" href="data:application/zip;base64,{b64}" download="THL_TO_SM.zip"></a>
                <script>document.getElementById('dl').click();</script>
            """, height=0)

        except Exception as e:
            st.error(f"❌ Lỗi: {e}")
