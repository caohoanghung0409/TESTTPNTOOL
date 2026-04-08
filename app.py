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
# STATE
# =========================
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

if "done" not in st.session_state:
    st.session_state["done"] = False


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
st.title("THL TO SM")

uploaded_files = st.file_uploader(
    "Chọn 2 file Excel",
    type=["xlsx"],
    accept_multiple_files=True,
    key=f"uploader_{st.session_state['uploader_key']}"
)

ready = uploaded_files and len(uploaded_files) == 2


if ready and not st.session_state["done"]:
    if st.button("Xử lý"):
        try:
            with st.spinner("Processing..."):

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
                # STEP 1: lấy số từ KET_QUA (BOOK1)
                # =========================
                df = pd.read_excel(path_book1, usecols=[0], dtype=str)

                valid_numbers = set()
                for v in df.iloc[:, 0].dropna():
                    for num in re.findall(r"\d+", str(v)):
                        num = "0" + num if len(num) == 3 else num
                        if len(num) == 4:
                            valid_numbers.add(num)

                # =========================
                # STEP 2: đọc KẾ HOẠCH + CHỈ GIỮ số VALID
                # =========================
                df2 = pd.read_excel(path_book1, header=None, dtype=str)

                palette = [
                    "FFF9C4","FFE0B2","FFCDD2","D1C4E9","C8E6C9",
                    "B3E5FC","F8BBD0","DCEDC8","D7CCC8","FFECB3",
                    "CFD8DC","F0F4C3","B2DFDB","E1BEE7","FFCCBC"
                ]

                number_to_color = {}
                group_index = 0

                # CHỈ group theo dòng có VALID number
                for _, row in df2.iterrows():
                    text = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])

                    nums = set()
                    for num in re.findall(r"\d+", text):
                        num = "0" + num if len(num) == 3 else num
                        if len(num) == 4 and num in valid_numbers:
                            nums.add(num)

                    # ❗ QUAN TRỌNG: bỏ dòng không có match
                    if not nums:
                        continue

                    color = palette[group_index % len(palette)]
                    group_index += 1

                    for n in nums:
                        if n not in number_to_color:
                            number_to_color[n] = color

                # =========================
                # STEP 3: xử lý KET_QUA
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

                count = 0

                for i in range(2, ws.max_row + 1):
                    val = ws.cell(i, col_index).value
                    if not val:
                        continue

                    nums = set()
                    for num in re.findall(r"\d+", str(val)):
                        num = "0" + num if len(num) == 3 else num
                        if len(num) == 4:
                            nums.add(num)

                    matched = nums & valid_numbers
                    if not matched:
                        continue

                    # lấy màu theo mapping từ kế hoạch
                    chosen_color = None
                    for n in matched:
                        if n in number_to_color:
                            chosen_color = number_to_color[n]
                            break

                    if chosen_color:
                        ws.cell(i, col_index).fill = PatternFill(
                            "solid",
                            fgColor=chosen_color
                        )
                        count += 1

                ws.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]
                ws.sheet_view.topLeftCell = "A1"

                wb.save(save_path)
                wb.close()

                # =========================
                # STEP 4: FILE KẾ HOẠCH (GIỮ NGUYÊN, chỉ tô đỏ valid)
                # =========================
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

                        parts += [red if check in valid_numbers else normal, num]
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

            st.success(f"DONE ✔ Matched rows: {count}")

            b64 = base64.b64encode(zip_data).decode()
            st.components.v1.html(f"""
                <a id="dl" href="data:application/zip;base64,{b64}" download="THL_TO_SM.zip"></a>
                <script>document.getElementById('dl').click();</script>
            """, height=0)

            st.session_state["done"] = True

        except Exception as e:
            st.error(f"Lỗi: {e}")

if st.session_state["done"]:
    if st.button("Reset"):
        st.session_state["uploader_key"] += 1
        st.session_state["done"] = False
        st.rerun()
