import streamlit as st
import pandas as pd
import re
import tempfile
import os
import zipfile
import xlsxwriter
import base64
import colorsys

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

st.set_page_config(page_title="THL TO SM", layout="centered")

# =========================
# UI (GIỮ NGUYÊN)
# =========================
st.markdown("""
<style>
header {display:none !important;}
#MainMenu {visibility:hidden;}
footer {visibility:hidden;}
.block-container {padding-top:0rem !important;}

.header{text-align:center;padding:8px 0;}
.header h1{color:#0284c7;margin:0;}
.header p{color:#64748b;margin:0;}

.card{background:white;padding:20px;border-radius:12px;}

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
# COLOR
# =========================
def gen_colors(n):
    base = ["FFD6D6","FFE0EB","EBD6FF","D6E4FF","D6F5FF","D6FFF5"]
    colors = base[:]
    for i in range(max(0, n-len(base))):
        h = (i*0.15)%1
        r,g,b = colorsys.hsv_to_rgb(h,0.3,0.95)
        colors.append('%02X%02X%02X'%(int(r*255),int(g*255),int(b*255)))
    return colors

# =========================
# SAFE LAST 4 DIGITS
# =========================
def last4(x):
    found = re.findall(r"\d+", str(x))
    if not found:
        return None
    v = found[-1]
    if len(v) == 3:
        v = "0"+v
    return v if len(v)==4 else None

# =========================
# HEADER
# =========================
st.markdown("""
<div class="header">
<h1>⚡ THL TO SM</h1>
<p>Xử lý Shipment ổn định</p>
</div>
""", unsafe_allow_html=True)

# =========================
# UI
# =========================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)

    files = st.file_uploader(
        "📂 Chọn 2 file Excel",
        type=["xlsx"],
        accept_multiple_files=True
    )

    if files and len(files)==2:

        if st.button("🚀 RUN"):

            try:
                tmp = tempfile.gettempdir()

                path1, path2 = None, None

                # =========================
                # SAVE FILES
                # =========================
                for f in files:
                    p = os.path.join(tmp, f.name)
                    with open(p,"wb") as w:
                        w.write(f.read())

                    df_check = pd.read_excel(p, dtype=str)

                    if df_check.astype(str).apply(lambda x: x.str.contains("Shipment", na=False)).any().any():
                        path1 = p
                    else:
                        path2 = p

                # =========================
                # ERP SAFE READ (PANDAS ONLY)
                # =========================
                df_tpn = pd.read_excel(path1, dtype=str).fillna("")
                df_book = pd.read_excel(path2, header=None, dtype=str).fillna("")

                # detect shipment column
                shipment_col = None
                for c in df_tpn.columns:
                    if df_tpn[c].astype(str).str.contains("Shipment", na=False).any():
                        shipment_col = c
                        break

                if shipment_col is None:
                    shipment_col = df_tpn.columns[0]

                # =========================
                # BUILD KEQUA SET
                # =========================
                ketqua = set()
                for v in df_tpn[shipment_col]:
                    x = last4(v)
                    if x:
                        ketqua.add(x)

                # =========================
                # GROUP BOOK1
                # =========================
                groups = []
                for _, r in df_book.iterrows():
                    t = str(r.iloc[0]) if len(r)>0 else ""
                    s = set()
                    for n in re.findall(r"\d+", t):
                        if len(n)==3: n="0"+n
                        if len(n)==4: s.add(n)
                    if s:
                        groups.append(s)

                colors = gen_colors(len(groups))

                # =========================
                # OUTPUT FILES
                # =========================
                out1 = os.path.join(tmp,"TPN_KET_QUA.xlsx")
                out2 = os.path.join(tmp,"TPN_KE_HOACH_XE.xlsx")

                # =========================
                # WRITE FILE 1 (OPENPYXL SAFE)
                # =========================
                wb = load_workbook(path1)
                ws = wb.active

                header_fill = PatternFill("solid", fgColor="000080")
                header_font = Font(color="FFFFFF", bold=True)

                for c in ws[1]:
                    c.fill = header_fill
                    c.font = header_font

                count = 0

                for i in range(2, ws.max_row+1):
                    val = ws.cell(i, ws.max_column).value
                    x = last4(val)

                    if x:
                        for idx,g in enumerate(groups):
                            if x in g:
                                ws.cell(i, ws.max_column).fill = PatternFill("solid", fgColor=colors[idx])
                                count+=1
                                break

                wb.save(out1)
                wb.close()

                # =========================
                # WRITE FILE 2 (KEEP FORMAT)
                # =========================
                workbook = xlsxwriter.Workbook(out2)
                sheet = workbook.add_worksheet()

                red = workbook.add_format({'font_color':'red'})
                normal = workbook.add_format({})

                for r,row in df_book.iterrows():

                    text = str(row.iloc[0]) if len(row)>0 else ""

                    parts=[]
                    last=0

                    for m in re.finditer(r"\d+",text):
                        n=m.group()
                        s,e=m.span()

                        chk = "0"+n if len(n)==3 else n

                        if s>last:
                            parts += [normal,text[last:s]]

                        parts += [red if chk in ketqua else normal,n]
                        last=e

                    if last<len(text):
                        parts += [normal,text[last:]]

                    try:
                        sheet.write_rich_string(r,0,*parts)
                    except:
                        sheet.write(r,0,text)

                workbook.close()

                # =========================
                # ZIP
                # =========================
                zip_path = os.path.join(tmp,"out.zip")
                with zipfile.ZipFile(zip_path,"w") as z:
                    z.write(out1,"TPN_KET_QUA.xlsx")
                    z.write(out2,"TPN_KE_HOACH_XE.xlsx")

                with open(zip_path,"rb") as f:
                    data = f.read()

                st.success(f"✅ DONE - MATCHED: {count}")

                b64 = base64.b64encode(data).decode()
                st.markdown(f"""
                <a href="data:application/zip;base64,{b64}" download="THL_TO_SM.zip">
                📥 DOWNLOAD FILE
                </a>
                """, unsafe_allow_html=True)

            except Exception as e:
                st.error(f"❌ Lỗi: {e}")

    st.markdown('</div>', unsafe_allow_html=True)
