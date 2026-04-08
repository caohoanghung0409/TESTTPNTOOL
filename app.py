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
    st.components.v1.html(f"""
    <a id="dl" href="data:application/zip;base64,{b64}" download="{filename}"></a>
    <script>document.getElementById('dl').click();</script>
    """, height=0)

# =========================
# SAFE LOAD
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
# NORMALIZE
# =========================
def normalize_text(s):
    if not s:
        return ""
    s = str(s)
    s = s.replace("\xa0"," ").replace("\u200b","")
    s = s.replace("\n"," ").replace("\r"," ")
    return re.sub(r"\s+"," ",s).strip().lower()

def find_shipment_col(ws):
    for r in range(1,6):
        for cell in ws[r]:
            if "shipment" in normalize_text(cell.value):
                return cell.column
    return None

# =========================
# UI
# =========================
st.title("⚡ THL TO SM")

files = st.file_uploader("Upload 2 file", type=["xlsx"], accept_multiple_files=True)

if files and len(files)==2 and st.button("🚀 Run"):

    tmp = tempfile.gettempdir()
    path1 = os.path.join(tmp, files[0].name)
    path2 = os.path.join(tmp, files[1].name)

    open(path1,"wb").write(files[0].read())
    open(path2,"wb").write(files[1].read())

    # detect file
    def is_tpn(path):
        wb = load_workbook_safe(path)
        ws = wb.active
        header = [normalize_text(c.value) for c in ws[1]]
        wb.close()
        return any("shipment" in h for h in header)

    if is_tpn(path1):
        path_tpn, path_book = path1, path2
    else:
        path_tpn, path_book = path2, path1

    # ===== lấy list số =====
    df = pd.read_excel(path_book, header=None, dtype=str)
    all_numbers = set()

    for i in range(len(df)):
        try:
            val = str(df.iat[i,0])
        except:
            continue

        for num in re.findall(r"\d+", val):
            if len(num)==3:
                num="0"+num
            if len(num)==4:
                all_numbers.add(num)

    # ===== xử lý TPN =====
    wb = load_workbook_safe(path_tpn)
    ws = wb.active

    col = find_shipment_col(ws)
    if not col:
        st.error("Không tìm thấy cột Shipment")
        st.stop()

    yellow = PatternFill("solid", fgColor="FFFF00")
    found = set()
    count = 0

    for i in range(2, ws.max_row+1):
        val = str(ws.cell(i,col).value or "")

        nums=set()
        for n in re.findall(r"\d+", val):
            if len(n)==3:
                n="0"+n
            if len(n)==4:
                nums.add(n)

        found.update(nums)

        if nums & all_numbers:
            ws.cell(i,col).fill = yellow
            count+=1

    save1 = os.path.join(tmp,"KET_QUA.xlsx")
    wb.save(save1)

    # ===== FILE 2 SAFE =====
    save2 = os.path.join(tmp,"KE_HOACH.xlsx")
    wb2 = xlsxwriter.Workbook(save2)
    ws2 = wb2.add_worksheet()

    red = wb2.add_format({'font_color':'red'})
    normal = wb2.add_format()

    for i in range(len(df)):

        try:
            text = str(df.iat[i,0])
        except:
            text = ""

        parts=[]
        last=0

        for m in re.finditer(r"\d+", text):
            num=m.group()
            s,e=m.span()

            check = "0"+num if len(num)==3 else num

            if s>last:
                parts += [normal, text[last:s]]

            if len(check)==4 and check in found:
                parts += [red, num]
            else:
                parts += [normal, num]

            last=e

        if last<len(text):
            parts += [normal, text[last:]]

        try:
            if len(parts)>2:
                ws2.write_rich_string(i,0,*parts)
            else:
                ws2.write(i,0,text)
        except:
            ws2.write(i,0,text)

    wb2.close()

    # ZIP
    zpath = os.path.join(tmp,"OUT.zip")
    with zipfile.ZipFile(zpath,"w") as z:
        z.write(save1,"KET_QUA.xlsx")
        z.write(save2,"KE_HOACH.xlsx")

    data=open(zpath,"rb").read()

    st.success(f"Done: {count}")
    auto_download(data,"RESULT.zip")
