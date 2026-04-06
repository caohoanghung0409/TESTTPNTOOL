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

st.set_page_config(
    page_title="THL TO SM",
    layout="wide",
    page_icon="⚙️"
)

# =========================
# ENTERPRISE STYLE (GREEN + RED)
# =========================
st.markdown("""
<style>

/* nền tối xanh rêu */
body {
    background-color: #052e1b;
}

/* ẩn UI streamlit */
header {visibility: hidden;}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

/* TOP BAR */
.topbar {
    background: #064e3b;
    padding: 14px 20px;
    border-radius: 12px;
    margin-bottom: 18px;
    border-left: 6px solid #22c55e;
}

.topbar h1 {
    color: #22c55e;
    margin: 0;
    font-size: 26px;
}

.topbar p {
    color: #d1fae5;
    margin: 2px 0 0 0;
    font-size: 13px;
}

/* PANEL */
.panel {
    background: #0b3b26;
    padding: 18px;
    border-radius: 14px;
    border: 1px solid #14532d;
}

/* INFO BOX */
.info {
    background: #052e1b;
    border-left: 5px solid #22c55e;
    padding: 10px;
    border-radius: 10px;
    color: #d1fae5;
    font-size: 13px;
}

/* ERROR BOX */
.errorbox {
    background: #450a0a;
    border-left: 5px solid #ef4444;
    padding: 10px;
    border-radius: 10px;
    color: #fecaca;
}

/* BUTTON GREEN */
.stButton>button {
    width: 100%;
    height: 45px;
    border-radius: 10px;
    background: #16a34a;
    color: white;
    font-weight: 700;
    border: none;
}

.stButton>button:hover {
    background: #22c55e;
}

/* DISABLED */
.stButton>button:disabled {
    background: #374151 !important;
    color: #9ca3af !important;
}

/* DOWNLOAD RED */
.stDownloadButton>button {
    width: 100%;
    height: 45px;
    border-radius: 10px;
    background: #dc2626;
    color: white;
    font-weight: 700;
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

if "last_file_hash" not in st.session_state:
    st.session_state["last_file_hash"] = None


# =========================
# EXCEL FIX
# =========================
def fix_excel_styles(path):
    tmp_dir = os.path.join(tempfile.gettempdir(), f"fix_{uuid.uuid4().hex}")
    os.makedirs(tmp_dir, exist_ok=True)

    with zipfile.ZipFile(path, 'r') as zin:
        zin.extractall(tmp_dir)

    style_path = os.path.join(tmp_dir, "xl", "styles.xml")
    if os.path.exists(style_path):
        os.remove(style_path)

    shutil.make_archive(path.replace(".xlsx","")+"_fixed","zip",tmp_dir)
    fixed_path = path.replace(".xlsx","_fixed.xlsx")
    os.rename(path.replace(".xlsx","")+"_fixed.zip", fixed_path)

    return fixed_path


def safe_load(path):
    try:
        return load_workbook(path, data_only=True, keep_links=False)
    except:
        fixed = fix_excel_styles(path)
        return load_workbook(fixed, data_only=True, keep_links=False)


def find_shipment_col(ws):
    for cell in ws[1]:
        if cell.value and "Shipment Nbr" in str(cell.value):
            return cell.column
    return None


# =========================
# HEADER
# =========================
st.markdown("""
<div class="topbar">
    <h1>⚙️ THL TO SM SYSTEM</h1>
    <p>Enterprise Excel Processing Dashboard</p>
</div>
""", unsafe_allow_html=True)


# =========================
# LAYOUT 2 CỘT
# =========================
col1, col2 = st.columns([1, 1])

with col1:
    st.markdown('<div class="panel">', unsafe_allow_html=True)

    st.markdown("### 📂 INPUT DATA")

    uploaded_files = st.file_uploader(
        "Upload 2 Excel files",
        type=["xlsx"],
        accept_multiple_files=True,
        key=f"up_{st.session_state['uploader_key']}"
    )

    ready = uploaded_files and len(uploaded_files) == 2

    if ready:
        st.markdown('<div class="info">✔ 2 files detected. Ready to process.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="errorbox">⚠ Please upload exactly 2 Excel files.</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)


with col2:
    st.markdown('<div class="panel">', unsafe_allow_html=True)

    st.markdown("### ⚡ ACTION PANEL")

    can_run = ready and (not st.session_state["processing"]) and (not st.session_state["done"])

    run = st.button("🚀 RUN PROCESS", disabled=not can_run)

    st.markdown("### 📦 OUTPUT")
    placeholder = st.empty()

    st.markdown('</div>', unsafe_allow_html=True)


# =========================
# PROCESS
# =========================
if run:

    st.session_state["processing"] = True

    with st.spinner("Processing data..."):

        tmp_dir = tempfile.gettempdir()
        path_tpn = None
        path_book1 = None

        for file in uploaded_files:
            path = os.path.join(tmp_dir, file.name)
            with open(path, "wb") as f:
                f.write(file.read())

            wb = safe_load(path)
            ws = wb.active
            header = [str(c.value) if c.value else "" for c in ws[1]]
            wb.close()

            if any("Shipment Nbr" in h for h in header):
                path_tpn = path
            else:
                path_book1 = path

        save_path = os.path.join(tmp_dir, "TPN_KET_QUA.xlsx")
        kehoach_path = os.path.join(tmp_dir, "TPN_KE_HOACH_XE.xlsx")

        df = pd.read_excel(path_book1, usecols=[0], dtype=str)

        all_numbers = set()
        for v in df.iloc[:,0].dropna():
            for num in re.findall(r"\d+", str(v)):
                if len(num) == 3:
                    num = "0"+num
                if len(num) == 4:
                    all_numbers.add(num)

        wb = safe_load(path_tpn)
        ws = wb.active

        col = find_shipment_col(ws)

        yellow = PatternFill("solid", fgColor="00FF00")

        ketqua = set()
        count = 0

        for i in range(2, ws.max_row+1):
            val = ws.cell(i, col).value

            if val:
                nums = set()
                for num in re.findall(r"\d+", str(val)):
                    if len(num) == 3:
                        num = "0"+num
                    if len(num) == 4:
                        nums.add(num)

                ketqua |= nums

                if nums & all_numbers:
                    ws.cell(i, col).fill = yellow
                    count += 1

        wb.save(save_path)
        wb.close()

        # fake output zip
        zip_path = os.path.join(tmp_dir, "result.zip")
        with zipfile.ZipFile(zip_path, "w") as z:
            z.write(save_path, "TPN_KET_QUA.xlsx")
            z.write(save_path, "TPN_KE_HOACH_XE.xlsx")

    st.session_state["processing"] = False
    st.session_state["done"] = True

    with placeholder:
        st.success(f"✅ DONE | MATCHED: {count}")

        with open(zip_path, "rb") as f:
            st.download_button(
                "⬇ DOWNLOAD RESULT",
                f,
                file_name="THL_TO_SM.zip"
            )

    st.session_state["uploader_key"] += 1
