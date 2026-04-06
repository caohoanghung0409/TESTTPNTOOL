import streamlit as st
import pandas as pd
import re
import tempfile
import os
import zipfile
import shutil
import uuid
import xlsxwriter

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.worksheet.views import Selection

st.set_page_config(page_title="THL TO SM PRO", layout="wide")

# =========================
# CSS PRO UI
# =========================
st.markdown("""
<style>
header, #MainMenu, footer {visibility: hidden;}

.block-container {
    padding-top: 1rem;
}

/* Background */
body {
    background: linear-gradient(135deg, #e0f2fe, #f0fdf4);
}

/* Card */
.card {
    background: white;
    padding: 20px;
    border-radius: 16px;
    box-shadow: 0 8px 24px rgba(0,0,0,0.08);
}

/* Title */
.title {
    font-size: 32px;
    font-weight: 700;
    color: #0f172a;
}

/* Sidebar */
.sidebar-box {
    background: white;
    padding: 15px;
    border-radius: 14px;
    box-shadow: 0 5px 20px rgba(0,0,0,0.06);
    margin-bottom: 15px;
}

/* Upload */
[data-testid="stFileUploader"] {
    border: 2px dashed #38bdf8;
    padding: 25px;
    border-radius: 12px;
    background: #f8fafc;
}

/* Buttons */
.stButton>button {
    width: 100%;
    height: 48px;
    border-radius: 12px;
    font-weight: 600;
    font-size: 16px;
    background: linear-gradient(90deg, #0284c7, #22c55e);
    color: white;
}

.stDownloadButton>button {
    width: 100%;
    height: 48px;
    border-radius: 12px;
    background: linear-gradient(90deg, #16a34a, #4ade80);
    color: white;
}

/* Metric */
.metric {
    font-size: 22px;
    font-weight: 700;
    color: #0284c7;
}
</style>
""", unsafe_allow_html=True)

# =========================
# STATE
# =========================
for k in ["uploader_key","processing","done","last_file_hash","matched"]:
    if k not in st.session_state:
        st.session_state[k] = 0 if k=="uploader_key" else False
st.session_state.setdefault("matched", 0)

# =========================
# FUNCTIONS (GIỮ NGUYÊN)
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
                content = re.sub(r'\s*s="\\d+"', '', content)
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
        if cell.value and "Shipment Nbr" in str(cell.value):
            return cell.column
    return None

# =========================
# LAYOUT
# =========================
col1, col2 = st.columns([3,1])

# =========================
# MAIN
# =========================
with col1:
    st.markdown('<div class="card">', unsafe_allow_html=True)

    st.markdown('<div class="title">⚡ THL TO SM PRO</div>', unsafe_allow_html=True)
    st.caption("Xử lý shipment nhanh – chính xác – tự động")

    uploaded_files = st.file_uploader(
        "📂 Upload 2 file Excel",
        type=["xlsx"],
        accept_multiple_files=True,
        key=f"uploader_{st.session_state['uploader_key']}"
    )

    ready = uploaded_files and len(uploaded_files) == 2
    can_run = ready and not st.session_state.processing and not st.session_state.done

    progress = st.progress(0)

    if ready:
        st.success("✔️ File hợp lệ")

        if st.button("🚀 BẮT ĐẦU"):
            st.session_state.processing = True

            with st.spinner("Đang xử lý..."):
                progress.progress(20)

                tmp_dir = tempfile.gettempdir()
                paths = []

                for f in uploaded_files:
                    p = os.path.join(tmp_dir, f.name)
                    with open(p,"wb") as file:
                        file.write(f.read())
                    paths.append(p)

                progress.progress(40)

                path_tpn = paths[0]
                path_book1 = paths[1]

                df = pd.read_excel(path_book1, usecols=[0], dtype=str)

                all_numbers = set()
                for v in df.iloc[:,0].dropna():
                    for num in re.findall(r"\d+", str(v)):
                        if len(num)==4:
                            all_numbers.add(num)

                wb = safe_load(path_tpn)
                ws = wb.active
                col_index = find_shipment_col(ws)

                yellow = PatternFill("solid", fgColor="FFFF00")

                count = 0
                for i in range(2, ws.max_row+1):
                    val = ws.cell(i, col_index).value
                    if val:
                        for num in re.findall(r"\d+", str(val)):
                            if num in all_numbers:
                                ws.cell(i,col_index).fill = yellow
                                count+=1

                save_path = os.path.join(tmp_dir,"TPN_KET_QUA.xlsx")
                wb.save(save_path)

                progress.progress(80)

                zip_path = os.path.join(tmp_dir,"TPN.zip")
                with zipfile.ZipFile(zip_path,"w") as z:
                    z.write(save_path,"TPN_KET_QUA.xlsx")

                with open(zip_path,"rb") as f:
                    zip_data = f.read()

                progress.progress(100)

            st.session_state.done = True
            st.session_state.processing = False
            st.session_state.matched = count

            st.success(f"✅ DONE – Matched: {count}")

            st.download_button("📥 Download", data=zip_data, file_name="RESULT.zip")

    else:
        st.warning("⚠️ Upload đủ 2 file")

    st.markdown('</div>', unsafe_allow_html=True)

# =========================
# SIDEBAR DASHBOARD
# =========================
with col2:
    st.markdown('<div class="sidebar-box">', unsafe_allow_html=True)
    st.markdown("### 📊 Dashboard")

    st.markdown(f"**Files:** {len(uploaded_files) if uploaded_files else 0}")
    st.markdown(f"**Status:** {'✅ Done' if st.session_state.done else '⏳ Waiting'}")
    st.markdown(f"**Matched:** {st.session_state.matched}")

    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="sidebar-box">', unsafe_allow_html=True)
    st.markdown("### ℹ️ Hướng dẫn")
    st.markdown("""
    1. Upload 2 file  
    2. Nhấn xử lý  
    3. Tải kết quả  
    """)
    st.markdown('</div>', unsafe_allow_html=True)
