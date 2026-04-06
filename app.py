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

st.set_page_config(page_title="THL TO SM PRO", layout="centered")

# =========================
# STATE
# =========================
if "dark" not in st.session_state:
    st.session_state.dark = False

if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

if "processing" not in st.session_state:
    st.session_state["processing"] = False

if "done" not in st.session_state:
    st.session_state["done"] = False

if "last_file_hash" not in st.session_state:
    st.session_state["last_file_hash"] = None


# =========================
# DARK MODE TOGGLE
# =========================
dark_mode = st.toggle("🌙 Dark Mode", value=st.session_state.dark)
st.session_state.dark = dark_mode


# =========================
# PRO CSS (GLASSMORPHISM + DARK MODE)
# =========================
bg_light = "#f1f5f9"
bg_dark = "#0f172a"

card_light = "rgba(255,255,255,0.7)"
card_dark = "rgba(30,41,59,0.6)"

text_light = "#0f172a"
text_dark = "#e2e8f0"

bg = bg_dark if dark_mode else bg_light
card = card_dark if dark_mode else card_light
text = text_dark if dark_mode else text_light

st.markdown(f"""
<style>

/* GLOBAL */
.stApp {{
    background: {bg};
    transition: all 0.3s ease;
}}

/* hide UI */
#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}
header {{display: none;}}

/* HEADER */
.header {{
    text-align: center;
    padding: 10px;
    animation: fadeIn 1s ease;
}}
.header h1 {{
    color: #38bdf8;
    margin: 0;
    font-size: 34px;
}}
.header p {{
    color: {text};
    opacity: 0.7;
}}

/* GLASS CARD */
.card {{
    background: {card};
    backdrop-filter: blur(14px);
    -webkit-backdrop-filter: blur(14px);
    border-radius: 16px;
    padding: 22px;
    box-shadow: 0 8px 30px rgba(0,0,0,0.15);
    border: 1px solid rgba(255,255,255,0.2);
    animation: slideUp 0.5s ease;
}}

/* BUTTON */
.stButton>button {{
    width: 100%;
    height: 44px;
    border-radius: 12px;
    background: linear-gradient(90deg,#0ea5e9,#22c55e);
    color: white;
    font-weight: 600;
    border: none;
    transition: 0.3s;
}}

.stButton>button:hover {{
    transform: scale(1.02);
}}

/* DOWNLOAD */
.stDownloadButton>button {{
    width: 100%;
    height: 44px;
    border-radius: 12px;
    background: #22c55e;
    color: white;
    font-weight: 600;
}}

/* ANIMATION */
@keyframes fadeIn {{
    from {{opacity:0}}
    to {{opacity:1}}
}}

@keyframes slideUp {{
    from {{transform: translateY(20px); opacity:0}}
    to {{transform: translateY(0); opacity:1}}
}}

/* LOADING DOTS */
.loader {{
  display: inline-block;
}}
.loader:after {{
  content: ' .';
  animation: dots 1s steps(5, end) infinite;
}}

@keyframes dots {{
  0%, 20% {{ color: rgba(0,0,0,0); text-shadow:
   .25em 0 0 rgba(0,0,0,0),
   .5em 0 0 rgba(0,0,0,0);}}
  40% {{ color: black; text-shadow:
   .25em 0 0 rgba(0,0,0,0),
   .5em 0 0 rgba(0,0,0,0);}}
  60% {{ text-shadow:
   .25em 0 0 black,
   .5em 0 0 rgba(0,0,0,0);}}
  80%, 100% {{ text-shadow:
   .25em 0 0 black,
   .5em 0 0 black;}}
}}

</style>
""", unsafe_allow_html=True)


# =========================
# EXCEL HELPERS (GIỮ NGUYÊN LOGIC)
# =========================
def find_shipment_col(ws):
    for cell in ws[1]:
        if cell.value:
            v = str(cell.value).replace("\xa0", " ").strip()
            if "Shipment Nbr" in v:
                return cell.column
    return None


def safe_load(path):
    return load_workbook(path, data_only=True, keep_links=False)


# =========================
# UI HEADER
# =========================
st.markdown("""
<div class="header">
    <h1>⚡ THL TO SM PRO</h1>
    <p>Glassmorphism • Fast Processing • Pro UI</p>
</div>
""", unsafe_allow_html=True)


# =========================
# CARD UI
# =========================
st.markdown('<div class="card">', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "📂 Upload 2 Excel files",
    type=["xlsx"],
    accept_multiple_files=True,
    key=f"up_{st.session_state['uploader_key']}"
)

current_hash = "|".join(sorted([f.name for f in uploaded_files])) if uploaded_files else None

if current_hash != st.session_state["last_file_hash"]:
    st.session_state["done"] = False
    st.session_state["processing"] = False
    st.session_state["last_file_hash"] = current_hash


ready = uploaded_files and len(uploaded_files) == 2
can_run = ready and not st.session_state["processing"] and not st.session_state["done"]


# =========================
# PROGRESS BAR PLACEHOLDER
# =========================
progress = st.progress(0)
status = st.empty()


if ready:
    if st.button("🚀 START PROCESS", disabled=not can_run):

        st.session_state["processing"] = True

        try:
            status.markdown("⏳ Starting...")
            progress.progress(10)

            tmp_dir = tempfile.gettempdir()
            path_tpn = None
            path_book1 = None

            # save files
            for i, file in enumerate(uploaded_files):
                path = os.path.join(tmp_dir, file.name)
                with open(path, "wb") as f:
                    f.write(file.read())

                wb = safe_load(path)
                ws = wb.active
                header = [str(c.value).strip() if c.value else "" for c in ws[1]]

                if any("Shipment Nbr" in h for h in header):
                    path_tpn = path
                else:
                    path_book1 = path

            progress.progress(30)
            status.markdown("📊 Reading data...")

            # process logic
            df = pd.read_excel(path_book1, usecols=[0], dtype=str)

            all_numbers = set()
            for v in df.iloc[:, 0].dropna():
                for num in re.findall(r"\d+", str(v)):
                    if len(num) == 3:
                        num = "0" + num
                    if len(num) == 4:
                        all_numbers.add(num)

            wb = safe_load(path_tpn)
            ws = wb.active

            col_index = find_shipment_col(ws)

            yellow = PatternFill("solid", fgColor="FFFF00")

            progress.progress(55)
            status.markdown("🧠 Processing Excel...")

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

            save_path = os.path.join(tmp_dir, "TPN_RESULT.xlsx")
            wb.save(save_path)
            wb.close()

            progress.progress(80)
            status.markdown("📦 Creating file...")

            zip_path = os.path.join(tmp_dir, "result.zip")
            with zipfile.ZipFile(zip_path, "w") as z:
                z.write(save_path, "TPN_RESULT.xlsx")

            with open(zip_path, "rb") as f:
                zip_data = f.read()

            progress.progress(100)
            status.markdown("✅ Done!")

            st.success(f"Matched: {count}")

            st.download_button(
                "📥 Download ZIP",
                data=zip_data,
                file_name="THL_TO_SM.zip"
            )

            st.session_state["done"] = True
            st.session_state["processing"] = False
            st.session_state["uploader_key"] += 1

        except Exception as e:
            st.session_state["processing"] = False
            st.error(f"❌ Error: {e}")

st.markdown("</div>", unsafe_allow_html=True)
