import io, os, re, zipfile, tempfile
from pathlib import Path
import pandas as pd
import streamlit as st
from PIL import Image, ImageOps, ImageEnhance
import pytesseract, gdown
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import matplotlib.pyplot as plt

st.set_page_config(page_title="Chiến Binh | Báo cáo", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
.block-container{padding-top:1.4rem;padding-bottom:2rem;max-width:1450px}
[data-testid="stSidebar"]{background:#f4f7fb}
.hero{background:linear-gradient(120deg,#063b70,#0876bd);padding:26px 30px;border-radius:18px;color:white;margin-bottom:18px}
.hero h1{color:white;font-size:29px;margin:0 0 6px 0}
.hero p{color:#d9ecff;margin:0;font-size:14px}
.step-card{background:#f6f9fc;border:1px solid #e4ebf3;border-radius:14px;padding:16px 18px;margin-bottom:12px}
.metric{background:white;border:1px solid #e5eaf0;border-radius:12px;padding:14px}
.small-muted{color:#65758b;font-size:13px}
div.stButton>button{border-radius:9px;font-weight:600;min-height:42px}
div[data-testid="stDownloadButton"] button{border-radius:9px;font-weight:700}
</style>
""",unsafe_allow_html=True)

st.markdown('<div class="hero"><h1>📊 BÁO CÁO SỨC KHỎE CHIẾN BINH</h1><p>Nhập dữ liệu một lần • Kiểm tra minh bạch • Xuất Excel và ảnh báo cáo</p></div>',unsafe_allow_html=True)

DEFAULT_EMP=pd.DataFrame([
["82654","Nguyễn Văn Quyền"],["137502","Nguyễn Ngọc Tân"],["181846","Lưu Văn Nhật"],
["137501","Võ Văn Lượng"],["137506","Lưu Thị Tuyết Nhi"],["170206","Lê Thị Thu Hiền"],
["281296","Phạm Xuân Toàn"],["21512","Hà Thị Minh"],["269200","Võ Quốc Cường"]
],columns=["Mã NV","Nhân viên"])
COLS=["Mã NV","Nhân viên","Nhóm hàng","TARGET siêu thị","Tỷ lệ hoàn thành (%)"]

def extract_zip_images(raw):
    out=[]
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        bad=z.testzip()
        if bad: raise ValueError(f"File ZIP lỗi: {bad}")
        for info in z.infolist():
            if info.is_dir(): continue
            if Path(info.filename).suffix.lower() in [".png",".jpg",".jpeg",".webp",".bmp",".tif",".tiff"]:
                out.append((Path(info.filename).name,z.read(info)))
    return out

def drive_images(url):
    tmp=tempfile.mkdtemp(prefix="drive_")
    if "/folders/" in url:
        files=gdown.download_folder(url=url,output=tmp,quiet=True,use_cookies=False)
        if not files: raise ValueError("Không tải được thư mục. Kiểm tra quyền chia sẻ công khai.")
        result=[]
        for p in Path(tmp).rglob("*"):
            if p.is_file() and p.suffix.lower() in [".zip",".png",".jpg",".jpeg",".webp",".bmp"]:
                if p.suffix.lower()==".zip": result.extend(extract_zip_images(p.read_bytes()))
                else: result.append((p.name,p.read_bytes()))
        return result
    dest=os.path.join(tmp,"download")
    got=gdown.download(url=url,output=dest,quiet=True,fuzzy=True)
    if not got: raise ValueError("Không tải được link. Hãy kiểm tra quyền 'Bất kỳ ai có đường liên kết'.")
    p=Path(got)
    if p.suffix.lower()==".zip": return extract_zip_images(p.read_bytes())
    if p.suffix.lower() in [".png",".jpg",".jpeg",".webp",".bmp"]: return [(p.name,p.read_bytes())]
    raise ValueError("Link Drive cần trỏ tới ZIP hoặc ảnh.")

def run_ocr(name,raw):
    im=Image.open(io.BytesIO(raw))
    im=ImageOps.exif_transpose(im).convert("L")
    im=ImageEnhance.Contrast(im).enhance(1.8)
    if im.width<1400:
        scale=1400/im.width; im=im.resize((int(im.width*scale),int(im.height*scale)))
    return pytesseract.image_to_string(im,lang="vie+eng",config="--psm 6")

def build_excel(df, date, factor, people):
    wb=Workbook(); ws=wb.active; ws.title="KetQua"
    ws.merge_cells("A1:G1"); ws["A1"]="BÁO CÁO SỨC KHỎE CHIẾN BINH"
    ws["A1"].font=Font(size=16,bold=True,color="FFFFFF"); ws["A1"].fill=PatternFill("solid",fgColor="075A9C")
    ws.merge_cells("A2:G2"); ws["A2"]=f"Ngày báo cáo: {date} | Hệ số: {factor:.0%} | Số nhân viên: {people}"
    headers=["Mã NV","Nhân viên","Nhóm hàng","TARGET siêu thị","Tỷ lệ hoàn thành (%)","Mục tiêu cá nhân","Doanh số quy đổi"]
    for c,h in enumerate(headers,1):
        x=ws.cell(4,c,h); x.font=Font(bold=True,color="FFFFFF"); x.fill=PatternFill("solid",fgColor="1677C8"); x.alignment=Alignment(wrap_text=True)
    for r,row in enumerate(df.itertuples(index=False),5):
        vals=[row[0],row[1],row[2],row[3],row[4],row[5],row[6]]
        for c,v in enumerate(vals,1): ws.cell(r,c,v)
        for c in [4,6,7]: ws.cell(r,c).number_format='#,##0.00'
        ws.cell(r,5).number_format='0.00"%"'
    for col,w in {"A":13,"B":25,"C":28,"D":21,"E":22,"F":21,"G":21}.items(): ws.column_dimensions[col].width=w
    ws.freeze_panes="C5"; ws.auto_filter.ref=f"A4:G{max(4,4+len(df))}"
    audit=wb.create_sheet("KiemTra")
    audit.append(["Mã NV","Nhân viên","Nhóm hàng","TARGET","Tỷ lệ","Trạng thái"])
    for r,row in enumerate(df.itertuples(index=False),2):
        audit.append([row[0],row[1],row[2],row[3],row[4],"Đã nhập — cần đối chiếu ảnh gốc"])
    for c in audit[1]:
        c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="475569")
    b=io.BytesIO(); wb.save(b); b.seek(0); return b

def build_png(df,date):
    pv=df.pivot_table(index="Nhân viên",columns="Nhóm hàng",values="Tỷ lệ hoàn thành (%)",aggfunc="first",fill_value=0)
    fig,ax=plt.subplots(figsize=(max(13,4+len(pv.columns)*.55),max(5,2+len(pv)*.5)))
    vmax=max(100,float(pv.values.max())) if pv.size else 100
    im=ax.imshow(pv.values,aspect="auto",cmap="RdYlGn",vmin=0,vmax=vmax)
    ax.set_xticks(range(len(pv.columns))); ax.set_xticklabels(pv.columns,rotation=60,ha="right",fontsize=7)
    ax.set_yticks(range(len(pv.index))); ax.set_yticklabels(pv.index,fontsize=8)
    ax.set_title(f"TỶ LỆ HOÀN THÀNH — {date}",fontweight="bold",pad=15)
    for i in range(len(pv)):
        for j in range(len(pv.columns)): ax.text(j,i,f"{pv.iloc[i,j]:g}%",ha="center",va="center",fontsize=6)
    fig.colorbar(im,ax=ax,label="Tỷ lệ hoàn thành (%)"); fig.tight_layout()
    b=io.BytesIO(); fig.savefig(b,format="png",dpi=180,bbox_inches="tight"); plt.close(fig); b.seek(0); return b

if "image_items" not in st.session_state: st.session_state.image_items=[]
if "ocr_rows" not in st.session_state: st.session_state.ocr_rows=[]
if "input_df" not in st.session_state:
    st.session_state.input_df=pd.DataFrame(columns=COLS)

with st.sidebar:
    st.markdown("### ⚙️ CẤU HÌNH BÁO CÁO")
    report_date=st.text_input("Ngày báo cáo","09/11/2026")
    factor=st.number_input("Hệ số mục tiêu",min_value=0.0,max_value=5.0,value=1.10,step=0.05)
    headcount=st.number_input("Số nhân viên",min_value=1,max_value=100,value=9,step=1)
    st.divider()
    st.caption("Mục tiêu cá nhân = TARGET siêu thị × hệ số / số nhân viên.")
    st.caption("Doanh số quy đổi = mục tiêu cá nhân × tỷ lệ hoàn thành.")

tabs=st.tabs(["① NẠP DỮ LIỆU","② ĐỌC ẢNH OCR","③ KIỂM TRA & TÍNH","④ XUẤT BÁO CÁO"])
with tabs[0]:
    st.markdown("### Chọn nguồn dữ liệu")
    left,right=st.columns([1,1],gap="large")
    with left:
        z=st.file_uploader("Tải file ZIP dashboard",type=["zip"],key="zip_up")
        multi=st.file_uploader("Hoặc tải nhiều ảnh",type=["png","jpg","jpeg","webp","bmp"],accept_multiple_files=True,key="multi_up")
        if st.button("Nạp file đã chọn",type="primary",use_container_width=True):
            try:
                items=[]
                if z: items.extend(extract_zip_images(z.getvalue()))
                if multi: items.extend([(f.name,f.getvalue()) for f in multi])
                if not items: st.warning("Anh chưa chọn ZIP hoặc ảnh.")
                else:
                    st.session_state.image_items=items
                    st.session_state.ocr_rows=[]
                    st.success(f"Đã nhận {len(items)} ảnh.")
            except Exception as e: st.error(str(e))
    with right:
        url=st.text_input("Link Google Drive công khai",placeholder="Dán link ZIP hoặc thư mục Drive")
        st.caption("Quyền chia sẻ phải là: Bất kỳ ai có đường liên kết.")
        if st.button("Tải từ Google Drive",use_container_width=True):
            try:
                items=drive_images(url)
                st.session_state.image_items=items; st.session_state.ocr_rows=[]
                st.success(f"Đã tải {len(items)} ảnh từ Drive.")
            except Exception as e: st.error(str(e))
    st.divider()
    a,b,c=st.columns(3)
    a.metric("Ảnh đã nạp",len(st.session_state.image_items))
    b.metric("Số nhân viên cấu hình",int(headcount))
    c.metric("Hệ số mục tiêu",f"{factor:.0%}")
    if st.session_state.image_items:
        with st.expander("Danh sách ảnh đã nạp",expanded=True):
            for i,(name,raw) in enumerate(st.session_state.image_items,1):
                st.write(f"{i}. {name} · {len(raw):,} bytes")

with tabs[1]:
    st.markdown("### Nhận diện thông tin từ ảnh")
    st.caption("OCR trích chữ để hỗ trợ kiểm tra. Chưa tự suy đoán tên nhóm hàng hoặc tỷ lệ khi không chắc chắn.")
    st.dataframe(DEFAULT_EMP,use_container_width=True,hide_index=True)
    if st.button("▶ Chạy OCR toàn bộ ảnh",type="primary",disabled=not st.session_state.image_items):
        rows=[]
        bar=st.progress(0)
        for i,(name,raw) in enumerate(st.session_state.image_items):
            try:
                txt=run_ocr(name,raw)
                codes=[str(x) for x in DEFAULT_EMP["Mã NV"] if str(x) in (name+" "+txt)]
                rows.append({"Tên ảnh":name,"Mã NV nhận diện":codes[0] if len(codes)==1 else ("Không nhận diện" if not codes else "Nhiều mã"),"Nội dung OCR":txt,"Trạng thái":"Đã đọc — cần xác minh"})
            except Exception as e:
                rows.append({"Tên ảnh":name,"Mã NV nhận diện":"Lỗi OCR","Nội dung OCR":str(e),"Trạng thái":"Lỗi"})
            bar.progress((i+1)/len(st.session_state.image_items))
        st.session_state.ocr_rows=rows
        st.success(f"Hoàn tất OCR {len(rows)} ảnh.")
    if st.session_state.ocr_rows:
        st.dataframe(pd.DataFrame([{k:v for k,v in r.items() if k!="Nội dung OCR"} for r in st.session_state.ocr_rows]),use_container_width=True,hide_index=True)
        with st.expander("Xem nội dung OCR chi tiết"):
            for r in st.session_state.ocr_rows:
                st.markdown(f"**{r['Tên ảnh']} — {r['Mã NV nhận diện']}**")
                st.text(r["Nội dung OCR"][:5000])

with tabs[2]:
    st.markdown("### Bảng dữ liệu để tính báo cáo")
    st.info("Nhập hoặc dán dữ liệu đã xác minh. Không dùng dòng mẫu giả. Các cột bắt buộc: Mã NV, Nhân viên, Nhóm hàng, TARGET siêu thị, Tỷ lệ hoàn thành (%).")
    st.session_state.input_df=st.data_editor(st.session_state.input_df,num_rows="dynamic",use_container_width=True,key="grid_data")
    csv=st.text_area("Dán dữ liệu CSV (nếu có)",placeholder="Mã NV,Nhân viên,Nhóm hàng,TARGET siêu thị,Tỷ lệ hoàn thành (%)",height=90)
    if st.button("Nạp CSV vào bảng"):
        try:
            parsed=pd.read_csv(io.StringIO(csv))
            missing=[c for c in COLS if c not in parsed.columns]
            if missing: st.error("CSV thiếu cột: "+", ".join(missing))
            else: st.session_state.input_df=parsed[COLS]; st.success("Đã nạp dữ liệu CSV.")
        except Exception as e: st.error(f"CSV chưa hợp lệ: {e}")
    d=st.session_state.input_df.copy()
    if len(d):
        for c in ["TARGET siêu thị","Tỷ lệ hoàn thành (%)"]: d[c]=pd.to_numeric(d[c],errors="coerce")
        d["Mục tiêu cá nhân"]=d["TARGET siêu thị"]*factor/headcount
        d["Doanh số quy đổi"]=d["Mục tiêu cá nhân"]*d["Tỷ lệ hoàn thành (%)"]/100
        bad=d[COLS].isna().any(axis=1)|(d["Tỷ lệ hoàn thành (%)"]<0)|(d["Tỷ lệ hoàn thành (%)"]>10000)
        st.markdown("#### Kết quả tính tạm")
        st.dataframe(d,use_container_width=True,hide_index=True)
        st.metric("Dòng cần kiểm tra",int(bad.sum()))
        st.session_state.calculated=d
        st.session_state.has_errors=bool(bad.any())
    else:
        st.warning("Chưa có dữ liệu. Hãy nhập bảng đã xác minh hoặc dán CSV.")

with tabs[3]:
    st.markdown("### Tải báo cáo")
    d=st.session_state.get("calculated",pd.DataFrame())
    errors=st.session_state.get("has_errors",True)
    if len(d) and not errors:
        st.success("Đã qua kiểm tra định dạng cơ bản. Anh vẫn cần đối chiếu số liệu với ảnh gốc trước khi chốt.")
        x=build_excel(d[["Mã NV","Nhân viên","Nhóm hàng","TARGET siêu thị","Tỷ lệ hoàn thành (%)","Mục tiêu cá nhân","Doanh số quy đổi"]],report_date,factor,headcount)
        p=build_png(d,report_date)
        l,r=st.columns(2)
        with l: st.download_button("⬇ TẢI FILE EXCEL",x,f"Bao_cao_{report_date.replace('/','-')}.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
        with r: st.download_button("⬇ TẢI ẢNH PNG",p,f"Bao_cao_{report_date.replace('/','-')}.png","image/png",use_container_width=True)
    else:
        st.warning("Chưa thể xuất báo cáo: chưa có dữ liệu hoặc còn dòng lỗi. Hãy hoàn tất Bước 3.")
