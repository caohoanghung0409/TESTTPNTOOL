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

st.set_page_config(page_title="THL TO SM", layout="centered")

# =========================
# UI CSS (GIỮ NGUYÊN)
# =========================
st.markdown("""
<style>
header {display:none !important;}
#MainMenu {visibility:hidden;}
footer {visibility:hidden;}
.block-container {padding-top:0rem !important;}

.header {text-align:center;padding:8px 0;}
.header h1 {color:#0284c7;margin:0;}
.header p {color:#64748b;margin:0;}

.card {background:white;padding:20px;border-radius:12px;}

.stButton>button{
    width:100%;
    height:42px;
    border-radius:10px;
    background:linear-gradient(90deg,#0ea5e9,#22c55e);
    color:white;
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
# PALETTE
# =========================
PALETTE = [
    "FFF2CC","FCE4D6","E2EFDA","DDEBF7","E4DFEC",
    "F8CBAD","C6E0B4","BDD7EE","D9D2E9","FFD966"
]

# =========================
# 🔥 FIX EXCEL TRIỆT ĐỂ (QUAN TRỌNG NHẤT)
# =========================
def sanitize_excel(path):
    """
    Không sửa bằng openpyxl.
    Chỉ rebuild lại zip excel -> xoá styles lỗi.
    """
    tmp = os.path.join(tempfile.gettempdir(), f"fix_{uuid.uuid4().hex}")
    os.makedirs(tmp, exist_ok=True)

    with zipfile.ZipFile(path, 'r') as z:
        z.extractall(tmp)

    # 🔥 QUAN TRỌNG: xoá toàn bộ styles + theme
    for f in [
        "xl/styles.xml",
        "xl/theme/theme1.xml"
    ]:
        fp = os.path.join(tmp, f)
        if os.path.exists(fp):
            try:
                os.remove(fp)
            except:
                pass

    fixed = path.replace(".xlsx", "_fixed.xlsx")
    shutil.make_archive(fixed.replace(".xlsx", ""), 'zip', tmp)
    os.rename(fixed.replace(".xlsx", ".zip"), fixed)

    return fixed


def safe_load(path):
    try:
        return load_workbook(path)
    except:
        fixed = sanitize_excel(path)
        return load_workbook(fixed)


def find_shipment_col(ws):
    for cell in ws[1]:
        if cell.value and "Shipment Nbr" in str(cell.value):
            return cell.column
    return None


# =========================
# UI HEADER (GIỮ NGUYÊN)
# =========================
st.markdown("""
<div class="header">
    <h1>⚡ THL TO SM</h1>
    <p>Xử lý & đối soát Shipment</p>
</div>
""", unsafe_allow_html=True)

with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "📂 Chọn 2 file Excel",
        type=["xlsx"],
        accept_multiple_files=True,
        key=f"uploader_{st.session_state['uploader_key']}"
    )

    ready = uploaded and len(uploaded) == 2

    if ready and not st.session_state["done"]:
        if st.button("🚀 Bắt đầu xử lý"):

            try:
                with st.spinner("⏳ Đang xử lý..."):

                    tmp = tempfile.gettempdir()
                    path_tpn = None
                    path_book = None

                    # =========================
                    # phân loại file
                    # =========================
                    for f in uploaded:
                        p = os.path.join(tmp, f.name)
                        with open(p, "wb") as out:
                            out.write(f.read())

                        wb = safe_load(p)
                        header = [str(c.value) for c in wb.active[1]]
                        wb.close()

                        if any("Shipment Nbr" in str(h) for h in header):
                            path_tpn = p
                        else:
                            path_book = p

                    # =========================
                    # BOOK DATA
                    # =========================
                    df = pd.read_excel(path_book, header=None, dtype=str)

                    row_numbers = []
                    all_numbers = set()

                    for _, r in df.iterrows():
                        txt = "" if pd.isna(r.iloc[0]) else str(r.iloc[0])

                        nums = set()
                        for n in re.findall(r"\d+", txt):
                            if len(n) == 3:
                                n = "0" + n
                            if len(n) == 4:
                                nums.add(n)

                        row_numbers.append(nums)
                        all_numbers.update(nums)

                    row_color = {i: PALETTE[i % len(PALETTE)] for i in range(len(row_numbers))}

                    # =========================
                    # LOAD TPN (SAFE)
                    # =========================
                    wb = safe_load(path_tpn)
                    ws = wb.active

                    col = find_shipment_col(ws)
                    if not col:
                        st.error("Không tìm thấy Shipment Nbr")
                        st.stop()

                    # header style
                    for c in ws[1]:
                        c.fill = PatternFill("solid", fgColor="FF1F4E79")
                        c.font = Font(color="FFFFFFFF", bold=True)

                    # =========================
                    # TÔ MÀU KET_QUA
                    # =========================
                    for i in range(2, ws.max_row + 1):

                        val = ws.cell(i, col).value
                        if not val:
                            continue

                        nums = set()
                        for n in re.findall(r"\d+", str(val)):
                            if len(n) == 3:
                                n = "0" + n
                            if len(n) == 4:
                                nums.add(n)

                        match = nums & all_numbers
                        if not match:
                            continue

                        chosen = None
                        for idx, rset in enumerate(row_numbers):
                            if match & rset:
                                chosen = idx
                                break

                        color = row_color[chosen] if chosen is not None else "DDDDDD"

                        ws.cell(i, col).fill = PatternFill("solid", fgColor=color)

                    out1 = os.path.join(tmp, "TPN_KET_QUA.xlsx")
                    wb.save(out1)
                    wb.close()

                    # =========================
                    # FILE KEHOACH
                    # =========================
                    df2 = pd.read_excel(path_book, header=None, dtype=str)

                    out2 = os.path.join(tmp, "TPN_KE_HOACH_XE.xlsx")
                    book = xlsxwriter.Workbook(out2)
                    ws2 = book.add_worksheet()

                    normal = book.add_format()
                    red = book.add_format({"font_color": "#E74C3C"})

                    for r, row in df2.iterrows():
                        text = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])

                        parts = []
                        last = 0

                        for m in re.finditer(r"\d+", text):
                            s, e = m.span()

                            if s > last:
                                parts += [normal, text[last:s]]

                            num = m.group()
                            check = num if len(num) != 3 else "0" + num

                            if len(check) == 4 and check in all_numbers:
                                parts += [red, num]
                            else:
                                parts += [normal, num]

                            last = e

                        if last < len(text):
                            parts += [normal, text[last:]]

                        try:
                            ws2.write_rich_string(r, 0, *parts)
                        except:
                            ws2.write(r, 0, text)

                    book.close()

                    # =========================
                    # ZIP OUTPUT
                    # =========================
                    zip_path = os.path.join(tmp, "RESULT.zip")

                    with zipfile.ZipFile(zip_path, "w") as z:
                        z.write(out1, "TPN_KET_QUA.xlsx")
                        z.write(out2, "TPN_KE_HOACH_XE.xlsx")

                    with open(zip_path, "rb") as f:
                        data = f.read()

                st.success("✅ DONE")

                b64 = base64.b64encode(data).decode()

                st.components.v1.html(f"""
                    <a id="dl" href="data:application/zip;base64,{b64}" download="THL_TO_SM.zip"></a>
                    <script>document.getElementById('dl').click();</script>
                """, height=0)

                st.session_state["done"] = True

            except Exception as e:
                st.error("❌ Lỗi:")
                st.exception(e)

    if st.session_state["done"]:
        if st.button("🔄 Làm file mới"):
            st.session_state["uploader_key"] += 1
            st.session_state["done"] = False
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)
