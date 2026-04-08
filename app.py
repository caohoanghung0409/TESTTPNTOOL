import streamlit as st
import pandas as pd
import re
import tempfile
import os
import zipfile
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
for key, val in {
    "uploader_key": 0,
    "processing": False,
    "done": False,
    "last_file_hash": None,
    "zip_data": None,
    "match_count": 0
}.items():
    if key not in st.session_state:
        st.session_state[key] = val

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
# FIX STYLE ERROR
# =========================
def load_workbook_safe(path):
    try:
        return load_workbook(path, data_only=True)
    except:
        tmp = tempfile.mktemp(suffix=".xlsx")
        with zipfile.ZipFile(path, 'r') as zin:
            with zipfile.ZipFile(tmp, 'w') as zout:
                for item in zin.infolist():
                    if "styles.xml" not in item.filename:
                        zout.writestr(item, zin.read(item.filename))
        return load_workbook(tmp, data_only=True)

# =========================
# NORMALIZE TEXT
# =========================
def normalize_text(s):
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\xa0", " ").replace("\ufeff", "").replace("\u200b", "")
    s = s.replace("\r", " ").replace("\n", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip().lower()

# =========================
# FIND COLUMN
# =========================
def find_shipment_col(ws):
    for r in range(1, 6):
        for cell in ws[r]:
            val = normalize_text(cell.value)
            if "shipment" in val and "nbr" in val:
                return cell.column
    return None

# =========================
# HEADER
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
        st.session_state.update({
            "done": False,
            "processing": False,
            "zip_data": None,
            "match_count": 0,
            "last_file_hash": current_hash
        })

    ready = uploaded_files and len(uploaded_files) == 2
    can_run = ready and (not st.session_state["done"])

    # =========================
    # PROCESS
    # =========================
    if st.session_state["processing"] and not st.session_state["done"]:

        try:
            with st.spinner("⏳ Đang xử lý..."):

                tmp_dir = tempfile.gettempdir()
                path_tpn = None
                path_book1 = None

                for file in uploaded_files:
                    path = os.path.join(tmp_dir, file.name)

                    with open(path, "wb") as f:
                        f.write(file.read())

                    wb_check = load_workbook_safe(path)
                    ws_check = wb_check.active
                    header = [normalize_text(c.value) for c in ws_check[1]]
                    wb_check.close()

                    if any("shipment" in h for h in header):
                        path_tpn = path
                    else:
                        path_book1 = path

                if not path_tpn or not path_book1:
                    st.error("❌ Không đúng định dạng 2 file!")
                    st.stop()

                save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
                kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

                # ===== FIX 1 =====
                df = pd.read_excel(path_book1, engine="openpyxl", dtype=str)

                col_idx = None
                for i in range(df.shape[1]):
                    if df.iloc[:, i].notna().sum() > 0:
                        col_idx = i
                        break

                if col_idx is None:
                    st.error("❌ File Book1 không có dữ liệu")
                    st.stop()

                df = df.iloc[:, [col_idx]]

                all_numbers = set()
                for v in df.iloc[:, 0].dropna().astype(str):
                    for num in re.findall(r"\d+", v):
                        if len(num) == 3:
                            num = "0" + num
                        if len(num) == 4:
                            all_numbers.add(num)

                wb = load_workbook_safe(path_tpn)
                ws = wb.active

                col_index = find_shipment_col(ws)

                if not col_index:
                    st.error("❌ Không tìm thấy cột Shipment Nbr")
                    st.stop()

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

                # ===== FIX 2 =====
                df2 = pd.read_excel(path_book1, header=None, engine="openpyxl", dtype=str)

                if df2.shape[1] == 0:
                    df2[0] = ""

                workbook = xlsxwriter.Workbook(kehoach_path)
                worksheet = workbook.add_worksheet()

                red_format = workbook.add_format({'font_color': 'red'})
                normal_format = workbook.add_format({})

                max_len = 0

                for row_idx, row in df2.iterrows():

                    # ===== FIX 3 =====
                    try:
                        val = row.iloc[0] if len(row) > 0 else ""
                        cell_value = "" if pd.isna(val) else str(val)
                    except:
                        cell_value = ""

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

                    # ===== FIX 4 =====
                    try:
                        if len(parts) >= 2:
                            worksheet.write_rich_string(row_idx, 0, *parts)
                        else:
                            worksheet.write(row_idx, 0, cell_value)
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

            st.session_state["zip_data"] = zip_data
            st.session_state["match_count"] = count
            st.session_state["done"] = True
            st.session_state["processing"] = False

        except Exception as e:
            st.session_state["processing"] = False
            st.error(f"❌ Có lỗi xảy ra: {e}")

    # =========================
    # RESULT
    # =========================
    if st.session_state["done"]:
        st.success(f"✅ COMPLETE !!! Matched: {st.session_state['match_count']}")
        auto_download(st.session_state["zip_data"], "THL_TO_SM.zip")

        if st.button("🔄 Xử lý file mới", use_container_width=True):
            st.session_state["uploader_key"] += 1
            st.session_state["done"] = False
            st.session_state["zip_data"] = None
            st.rerun()

    else:
        if ready:
            if st.button("🚀 Bắt đầu xử lý", disabled=not can_run):
                st.session_state["processing"] = True
                st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)
