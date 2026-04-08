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
st.title("⚡ THL TO SM")

uploaded_files = st.file_uploader(
    "Chọn 2 file Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files and len(uploaded_files) == 2:

    if st.button("🚀 Xử lý"):

        tmp_dir = tempfile.gettempdir()
        path_tpn, path_book1 = None, None

        # phân loại file
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
        # lấy số từ KEHOACH
        # =========================
        df = pd.read_excel(path_book1, header=None, dtype=str)

        group_sets = []
        palette = ["FFFF00", "FF9999", "99FF99", "9999FF", "FFD966", "FF66FF"]

        for _, row in df.iterrows():
            text = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])
            nums = set(re.findall(r"\d+", text))

            if nums:
                group_sets.append(nums)

        # =========================
        # LOAD TPN
        # =========================
        wb = safe_load(path_tpn)
        ws = wb.active
        col_index = find_shipment_col(ws)

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

        # =========================
        # CHỈ TÔ CỘT Shipment Nbr
        # =========================
        color_map = {}

        for i in range(2, ws.max_row + 1):
            cell = ws.cell(i, col_index)
            val = cell.value

            if not val:
                continue

            nums = set(re.findall(r"\d+", str(val)))

            matched_color = None

            for idx, gset in enumerate(group_sets):
                if nums & gset:
                    if idx not in color_map:
                        color_map[idx] = palette[idx % len(palette)]
                    matched_color = color_map[idx]
                    break

            if matched_color:
                cell.fill = PatternFill("solid", fgColor=matched_color)

        ws.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]

        wb.save(save_path)
        wb.close()

        # =========================
        # FILE KEHOACH GIỮ NGUYÊN
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
                check = num

                if start > last:
                    parts += [normal, text[last:start]]

                parts += [red if any(check in g for g in group_sets) else normal, num]
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
        zip_path = os.path.join(tmp_dir, "RESULT.zip")
        with zipfile.ZipFile(zip_path, "w") as z:
            z.write(save_path, "TPN_KET_QUA.xlsx")
            z.write(kehoach_path, "TPN_KE_HOACH_XE.xlsx")

        with open(zip_path, "rb") as f:
            data = f.read()

        st.success("DONE")

        b64 = base64.b64encode(data).decode()
        st.markdown(f"""
        <a href="data:application/zip;base64,{b64}" download="THL_TO_SM.zip">DOWNLOAD</a>
        """, unsafe_allow_html=True)
