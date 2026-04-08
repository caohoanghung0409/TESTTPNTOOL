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

.stButton>button:disabled {
    background: #94a3b8 !important;
    opacity: 0.6;
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

# =========================
# COLOR PALETTE (TƯƠNG PHẢN MẠNH)
# =========================
COLOR_PALETTE = [
    "FF3B30",  # đỏ
    "34C759",  # xanh lá
    "007AFF",  # xanh dương
    "FF9500",  # cam
    "AF52DE",  # tím
    "FF2D55",  # hồng đậm
    "5856D6",  # xanh tím
    "00C7BE",  # xanh ngọc
    "FFCC00",  # vàng đậm
    "FF6B00",  # cam đậm
    "4CD964",  # xanh lá nhạt đậm
    "1E90FF",  # xanh biển
]

# =========================
# FIX EXCEL
# =========================
def fix_excel_styles(path):
    tmp_dir = os.path.join(tempfile.gettempdir(), f"fix_{uuid.uuid4().hex}")
    os.makedirs(tmp_dir, exist_ok=True)

    with zipfile.ZipFile(path, 'r') as zin:
        zin.extractall(tmp_dir)

    style_path = os.path.join(tmp_dir, "xl", "styles.xml")
    if os.path.exists(style_path):
        os.remove(style_path)

    sheet_dir = os.path.join(tmp_dir, "xl", "worksheets")

    if os.path.exists(sheet_dir):
        for file in os.listdir(sheet_dir):
            if file.endswith(".xml"):
                fpath = os.path.join(sheet_dir, file)
                with open(fpath, "r", encoding="utf-8") as f:
                    content = f.read()

                content = re.sub(r'\s*s="\d+"', '', content)

                with open(fpath, "w", encoding="utf-8") as f:
                    f.write(content)

    fixed_path = path.replace(".xlsx", "_fixed.xlsx")
    shutil.make_archive(fixed_path.replace(".xlsx", ""), 'zip', tmp_dir)
    os.rename(fixed_path.replace(".xlsx", ".zip"), fixed_path)

    return fixed_path


def safe_load(path, read_only=False):
    try:
        return load_workbook(path, read_only=read_only, data_only=True, keep_links=False)
    except:
        fixed = fix_excel_styles(path)
        return load_workbook(fixed, read_only=read_only, data_only=True, keep_links=False)


def find_shipment_col(ws):
    for cell in ws[1]:
        if cell.value:
            v = str(cell.value).replace("\xa0", " ").strip()
            if "Shipment Nbr" in v:
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
                    # LẤY SỐ TỪ BOOK1
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
                    # LOAD TPN
                    # =========================
                    wb = safe_load(path_tpn)
                    ws = wb.active
                    col_index = find_shipment_col(ws)

                    yellow = PatternFill("solid", fgColor="FFFF00")
                    header_fill = PatternFill("solid", fgColor="000080")
                    header_font = Font(color="FFFFFF", bold=True)
                    bold_font = Font(bold=True)

                    for cell in ws[1]:
                        cell.fill = header_fill
                        cell.font = header_font

                    for row in ws.iter_rows(min_row=2):
                        for cell in row:
                            if cell.value:
                                cell.font = bold_font

                    ketqua_numbers = set()

                    # =========================
                    # MAP SỐ -> MÀU (QUAN TRỌNG)
                    # =========================
                    number_color_map = {}
                    color_index = 0

                    # =========================
                    # TÔ MÀU TỪ TPN
                    # =========================
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

                            # gán màu cho từng số (KHÔNG trùng, tránh giống nhau)
                            row_color = None
                            for n in nums:
                                if n in all_numbers:
                                    if n not in number_color_map:
                                        number_color_map[n] = COLOR_PALETTE[color_index % len(COLOR_PALETTE)]
                                        color_index += 1

                                    row_color = number_color_map[n]

                            # tô dòng TPN_KET_QUA
                            if nums & all_numbers:
                                ws.cell(i, col_index).fill = yellow

                    wb.save(save_path)
                    wb.close()

                    # =========================
                    # FILE KẾ HOẠCH (RICH TEXT)
                    # =========================
                    df2 = pd.read_excel(path_book1, header=None, dtype=str)

                    workbook = xlsxwriter.Workbook(kehoach_path)
                    worksheet = workbook.add_worksheet()

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
                            check = "0"+num if len(num)==3 else num

                            color_hex = number_color_map.get(check, None)
                            fmt = workbook.add_format({'font_color': color_hex}) if color_hex else normal

                            if start > last:
                                parts += [normal, text[last:start]]

                            parts += [fmt, num]
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
                    # ZIP OUTPUT
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
                    <a id="dl" href="data:application/zip;base64,{b64}" download="THL TO SM.zip"></a>
                    <script>document.getElementById('dl').click();</script>
                """, height=0)

                st.session_state["done"] = True

            except:
                st.error("❌ Có lỗi xảy ra!")

    if st.session_state["done"]:
        if st.button("🔄 Xử lý file mới"):
            st.session_state["uploader_key"] += 1
            st.session_state["done"] = False
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)
