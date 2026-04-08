import streamlit as st
import pandas as pd
import re
import tempfile
import os
import zipfile
import xlsxwriter
import base64

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

st.set_page_config(page_title="THL TO SM", layout="centered")

# =========================
# UI
# =========================
st.markdown("""
<style>
header {display:none;}
#MainMenu {visibility:hidden;}
footer {visibility:hidden;}

.card {padding:20px;background:#fff;border-radius:12px;}
</style>
""", unsafe_allow_html=True)

# =========================
# PALETTE
# =========================
PALETTE = [
    "FFF2CC","FCE4D6","E2EFDA","DDEBF7","E4DFEC",
    "F8CBAD","C6E0B4","BDD7EE","D9D2E9","FFD966"
]

# =========================
# SAFE READ EXCEL (FIX TRIỆT ĐỂ)
# =========================
def read_excel_safe(path, header=None):
    try:
        # engine calamine ổn định hơn openpyxl khi file lỗi style
        return pd.read_excel(path, header=header, dtype=str, engine="calamine")
    except:
        # fallback
        return pd.read_excel(path, header=header, dtype=str)


def find_shipment_col(df):
    for i, c in enumerate(df.columns):
        if "Shipment Nbr" in str(c):
            return i
    return None


# =========================
# UI
# =========================
st.title("⚡ THL TO SM")

uploaded = st.file_uploader("Upload 2 file Excel", type=["xlsx"], accept_multiple_files=True)

if uploaded and len(uploaded) == 2:

    if st.button("🚀 RUN"):

        tmp = tempfile.gettempdir()
        path_tpn = None
        path_book = None

        # save files
        for f in uploaded:
            p = os.path.join(tmp, f.name)
            with open(p, "wb") as out:
                out.write(f.read())

            df_head = read_excel_safe(p, header=0)

            if any("Shipment Nbr" in str(x) for x in df_head.columns):
                path_tpn = p
            else:
                path_book = p

        # =========================
        # BOOK (KEHOACH)
        # =========================
        df_book = read_excel_safe(path_book, header=None)

        row_numbers = []
        all_numbers = set()

        for _, r in df_book.iterrows():
            txt = "" if pd.isna(r.iloc[0]) else str(r.iloc[0])

            nums = set()
            for n in re.findall(r"\d+", txt):
                if len(n) == 3:
                    n = "0" + n
                if len(n) == 4:
                    nums.add(n)

            row_numbers.append(nums)
            all_numbers.update(nums)

        # map màu theo dòng
        row_color = {i: PALETTE[i % len(PALETTE)] for i in range(len(row_numbers))}

        # =========================
        # READ TPN (SAFE)
        # =========================
        df_tpn = read_excel_safe(path_tpn, header=0)

        col_idx = find_shipment_col(df_tpn)

        if col_idx is None:
            st.error("Không tìm thấy Shipment Nbr")
            st.stop()

        # =========================
        # PROCESS KET_QUA
        # =========================
        from openpyxl import load_workbook

        wb = load_workbook(path_tpn)
        ws = wb.active

        for i in range(2, ws.max_row + 1):

            val = ws.cell(i, col_idx + 1).value
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

            ws.cell(i, col_idx + 1).fill = PatternFill(
                "solid",
                fgColor=color
            )

        out1 = os.path.join(tmp, "TPN_KET_QUA.xlsx")
        wb.save(out1)
        wb.close()

        # =========================
        # KEHOACH FILE
        # =========================
        df2 = read_excel_safe(path_book, header=None)

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
        # ZIP
        # =========================
        zip_path = os.path.join(tmp, "RESULT.zip")

        with zipfile.ZipFile(zip_path, "w") as z:
            z.write(out1, "TPN_KET_QUA.xlsx")
            z.write(out2, "TPN_KE_HOACH_XE.xlsx")

        with open(zip_path, "rb") as f:
            data = f.read()

        st.success("DONE")

        b64 = base64.b64encode(data).decode()

        st.components.v1.html(f"""
            <a id="dl" href="data:application/zip;base64,{b64}" download="THL_TO_SM.zip"></a>
            <script>document.getElementById('dl').click();</script>
        """, height=0)
