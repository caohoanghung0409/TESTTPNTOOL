import streamlit as st
import pandas as pd
import re
import tempfile
import os
import zipfile
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
        if cell.value and "Shipment Nbr" in str(cell.value):
            return cell.column
    return None


# =========================
# UI
# =========================
st.markdown("""
<div class="header">
    <h1>⚡ THL TO SM</h1>
    <p>Xử lý & đối soát Shipment nhanh chóng</p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader("📂 Chọn 2 file Excel", type=["xlsx"], accept_multiple_files=True)

if uploaded_files and len(uploaded_files) == 2:

    if st.button("🚀 Bắt đầu xử lý"):

        try:
            with st.spinner("⏳ Đang xử lý..."):

                tmp_dir = tempfile.gettempdir()
                path_tpn, path_book1 = None, None

                for file in uploaded_files:
                    path = os.path.join(tmp_dir, file.name)
                    with open(path, "wb") as f:
                        f.write(file.read())

                    wb = load_workbook(path, read_only=True)
                    header = [str(c.value) if c.value else "" for c in wb.active[1]]
                    wb.close()

                    if any("Shipment Nbr" in h for h in header):
                        path_tpn = path
                    else:
                        path_book1 = path

                save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
                kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

                df = pd.read_excel(path_book1, usecols=[0], dtype=str)

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
                        nums = set(re.findall(r"\d+", str(val)))
                        nums = {"0"+n if len(n)==3 else n for n in nums if len(n) in [3,4]}

                        ketqua_numbers.update(nums)

                        if nums & all_numbers:
                            ws.cell(i, col_index).fill = yellow
                            count += 1

                # FIX A1
                ws.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]
                ws.sheet_view.topLeftCell = "A1"

                wb.save(save_path)
                wb.close()

                # =========================
                # FILE 2 FIX AUTO WIDTH
                # =========================
                df2 = pd.read_excel(path_book1, header=None, dtype=str)

                workbook = xlsxwriter.Workbook(kehoach_path)
                worksheet = workbook.add_worksheet()

                red = workbook.add_format({'font_color': 'red'})
                normal = workbook.add_format({})

                max_len = 0

                for r, row in df2.iterrows():
                    val = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])
                    max_len = max(max_len, len(val))

                    parts = []
                    last = 0

                    for m in re.finditer(r"\d+", val):
                        num = m.group()
                        s, e = m.span()

                        num_check = "0"+num if len(num)==3 else num

                        if s > last:
                            parts += [normal, val[last:s]]

                        if len(num_check)==4 and num_check in ketqua_numbers:
                            parts += [red, num]
                        else:
                            parts += [normal, num]

                        last = e

                    if last < len(val):
                        parts += [normal, val[last:]]

                    try:
                        worksheet.write_rich_string(r, 0, *parts)
                    except:
                        worksheet.write(r, 0, val)

                # ✅ AUTO WIDTH
                worksheet.set_column(0, 0, max_len + 3)

                # ✅ FIX VIEW A1
                worksheet.write(0, 0, df2.iloc[0,0] if not pd.isna(df2.iloc[0,0]) else "")
                worksheet.freeze_panes(0, 0)

                workbook.close()

                # ZIP
                zip_path = os.path.join(tmp_dir, "TPN_COMPLETE.zip")
                with zipfile.ZipFile(zip_path, "w") as z:
                    z.write(save_path, "TPN_KET_QUA.xlsx")
                    z.write(kehoach_path, "TPN_KE_HOACH_XE.xlsx")

                with open(zip_path, "rb") as f:
                    data = f.read()

            st.success(f"✅ COMPLETE !!! Matched: {count}")

            auto_download(data, "THL_TO_SM.zip")

        except Exception:
            st.error("❌ Có lỗi xảy ra!")
