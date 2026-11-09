import io, os, re, zipfile, tempfile
from pathlib import Path
import pandas as pd
import streamlit as st
from PIL import Image, ImageOps, ImageEnhance
import pytesseract, gdown
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
import matplotlib.pyplot as plt

st.set_page_config(page_title="Báo cáo chiến binh",layout="wide")
st.title("HỆ THỐNG BÁO CÁO SỨC KHỎE CHIẾN BINH")
st.caption("MVP: ZIP / ảnh / Google Drive công khai → OCR → rà soát → Excel + PNG")
DEFAULT=pd.DataFrame([
["82654","Nguyễn Văn Quyền"],["137502","Nguyễn Ngọc Tân"],["181846","Lưu Văn Nhật"],
["137501","Võ Văn Lượng"],["137506","Lưu Thị Tuyết Nhi"],["170206","Lê Thị Thu Hiền"],
["281296","Phạm Xuân Toàn"],["21512","Hà Thị Minh"],["269200","Võ Quốc Cường"]
],columns=["MaNV","NhanVien"])
def ocr_image(path):
    im=ImageOps.grayscale(ImageOps.exif_transpose(Image.open(path)).convert("RGB"))
    im=ImageEnhance.Contrast(im).enhance(1.8)
    if im.width<1400:
        k=1400/im.width; im=im.resize((int(im.width*k),int(im.height*k)))
    return pytesseract.image_to_string(im,lang="vie+eng",config="--psm 6")
def excel_export(df,date,factor,n):
    wb=Workbook(); ws=wb.active; ws.title="KetQua"
    ws.append(["BÁO CÁO SỨC KHỎE CHIẾN BINH"])
    ws.append([f"Ngày: {date} | Hệ số: {factor:.0%} | Số NV: {n}"])
    ws.append([])
    headers=["Mã NV","Nhân viên","Nhóm hàng","Tỷ lệ hoàn thành (%)","Mục tiêu cá nhân","Doanh số quy đổi"]
    ws.append(headers)
    for c in ws[4]:
        c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="1F4E78")
    for r in df.itertuples(index=False):
        ws.append([r.MaNV,r.NhanVien,r.NhomHang,r.TyLeHoanThanh,r.MucTieuCaNhan,r.DoanhSoQuyDoi])
    for col,w in {"A":14,"B":26,"C":28,"D":22,"E":22,"F":22}.items(): ws.column_dimensions[col].width=w
    ws.freeze_panes="C5"
    for row in ws.iter_rows(min_row=5,min_col=5,max_col=6):
        for c in row: c.number_format="#,##0.00"
    b=io.BytesIO(); wb.save(b); b.seek(0); return b
def png_export(df,date):
    p=df.pivot_table(index="NhanVien",columns="NhomHang",values="TyLeHoanThanh",aggfunc="first",fill_value=0)
    fig,ax=plt.subplots(figsize=(max(12,4+len(p.columns)*.55),max(5,2+len(p)*.5)))
    im=ax.imshow(p.values,aspect="auto",cmap="RdYlGn",vmin=0,vmax=max(100,float(p.values.max())))
    ax.set_xticks(range(len(p.columns))); ax.set_xticklabels(p.columns,rotation=65,ha="right",fontsize=7)
    ax.set_yticks(range(len(p))); ax.set_yticklabels(p.index,fontsize=8)
    ax.set_title(f"TỶ LỆ HOÀN THÀNH — {date}")
    for i in range(len(p)):
        for j in range(len(p.columns)): ax.text(j,i,f"{p.iloc[i,j]:g}%",ha="center",va="center",fontsize=6)
    fig.colorbar(im,ax=ax,label="%"); fig.tight_layout()
    b=io.BytesIO(); fig.savefig(b,format="png",dpi=180,bbox_inches="tight"); plt.close(fig); b.seek(0); return b

with st.sidebar:
    date=st.text_input("Ngày báo cáo","09/11/2026")
    factor=st.number_input("Hệ số mục tiêu",0.0,5.0,1.10,.05)
    n=st.number_input("Số nhân viên",1,100,9,1)
st.subheader("1. Nạp dữ liệu")
a,b=st.columns(2)
with a:
    zfile=st.file_uploader("Tải ZIP ảnh dashboard",type=["zip"])
    imgs=st.file_uploader("Hoặc tải nhiều ảnh",type=["png","jpg","jpeg","webp"],accept_multiple_files=True)
with b:
    drive=st.text_input("Link Google Drive công khai (ZIP hoặc thư mục)")
    getdrive=st.button("Tải từ Google Drive")
tmp=tempfile.mkdtemp(prefix="bc_"); paths=[]
if zfile:
    try:
        with zipfile.ZipFile(io.BytesIO(zfile.getvalue())) as z:
            for x in z.infolist():
                if Path(x.filename).suffix.lower() in [".png",".jpg",".jpeg",".webp",".bmp"]:
                    out=os.path.join(tmp,Path(x.filename).name)
                    with z.open(x) as src,open(out,"wb") as dst: dst.write(src.read())
    except Exception as e: st.error(f"ZIP lỗi: {e}")
if imgs:
    for f in imgs:
        p=os.path.join(tmp,f.name); open(p,"wb").write(f.getvalue())
if getdrive:
    try:
        if "/folders/" in drive:
            gdown.download_folder(drive,output=tmp,quiet=True,use_cookies=False)
        else: gdown.download(drive,output=os.path.join(tmp,"drive.zip"),quiet=True,fuzzy=True)
        for p in Path(tmp).rglob("*.zip"):
            with zipfile.ZipFile(p) as z:
                for x in z.infolist():
                    if Path(x.filename).suffix.lower() in [".png",".jpg",".jpeg",".webp",".bmp"]:
                        out=os.path.join(tmp,Path(x.filename).name)
                        with z.open(x) as src,open(out,"wb") as dst: dst.write(src.read())
    except Exception as e: st.error(f"Không tải được Drive: {e}. Hãy bật chia sẻ công khai.")
paths=list(dict.fromkeys(str(p) for p in Path(tmp).rglob("*") if p.is_file() and p.suffix.lower() in [".png",".jpg",".jpeg",".webp",".bmp"]))
st.write(f"Ảnh nhận được: **{len(paths)}**")
st.subheader("2. Danh sách nhân viên")
if "employees" not in st.session_state: st.session_state.employees=DEFAULT
st.session_state.employees=st.data_editor(st.session_state.employees,num_rows="dynamic",use_container_width=True)
if st.button("Chạy OCR",disabled=not paths,type="primary"):
    audit=[]
    for p in paths:
        try: txt=ocr_image(p)
        except Exception as e: txt=f"LỖI OCR: {e}"
        code=next((str(x) for x in st.session_state.employees.MaNV if str(x) in (Path(p).name+" "+txt)),"")
        audit.append({"Ảnh":Path(p).name,"Mã NV nhận diện":code,"Nội dung OCR":txt})
    st.session_state.audit=audit
    st.success(f"Đã đọc {len(audit)} ảnh. Cần xác nhận dữ liệu trích xuất.")
if "audit" in st.session_state:
    with st.expander("Nhật ký OCR — kiểm tra ảnh và chữ đã đọc",expanded=True):
        for x in st.session_state.audit:
            st.markdown(f"**{x['Ảnh']} | Mã NV: {x['Mã NV nhận diện'] or 'Chưa nhận diện'}**")
            st.text(x["Nội dung OCR"][:5000])
st.subheader("3. Bảng dữ liệu chuẩn để tính")
st.info("OCR ở MVP này đọc và nhận diện ảnh, nhưng chưa tự gán tỷ lệ vào nhóm hàng vì không được phép đoán khi bố cục/chữ không chắc chắn. Anh rà soát rồi nhập hoặc dán CSV để tính chính xác.")
template=pd.DataFrame([["82654","Nguyễn Văn Quyền","Tên nhóm hàng",100000000,0]],columns=["MaNV","NhanVien","NhomHang","TargetSieuThi","TyLeHoanThanh"])
if "data" not in st.session_state: st.session_state.data=template
st.session_state.data=st.data_editor(st.session_state.data,num_rows="dynamic",use_container_width=True)
csv=st.text_area("Hoặc dán CSV: MaNV,NhanVien,NhomHang,TargetSieuThi,TyLeHoanThanh")
if st.button("Nạp CSV"):
    try:
        d=pd.read_csv(io.StringIO(csv))
        st.session_state.data=d
        st.success("Đã nạp CSV.")
    except Exception as e: st.error(f"CSV lỗi: {e}")
d=st.session_state.data.copy()
need=["MaNV","NhanVien","NhomHang","TargetSieuThi","TyLeHoanThanh"]
if all(c in d.columns for c in need):
    d["TargetSieuThi"]=pd.to_numeric(d.TargetSieuThi,errors="coerce")
    d["TyLeHoanThanh"]=pd.to_numeric(d.TyLeHoanThanh,errors="coerce")
    d["MucTieuCaNhan"]=d.TargetSieuThi*factor/n
    d["DoanhSoQuyDoi"]=d.MucTieuCaNhan*d.TyLeHoanThanh/100
    bad=d[need].isna().any(axis=1)|(d.TyLeHoanThanh<0)|(d.TyLeHoanThanh>10000)
    st.subheader("4. Kết quả")
    st.dataframe(d,use_container_width=True)
    st.warning(f"Có {int(bad.sum())} dòng cần sửa." if bad.any() else "Kiểm tra cơ bản đạt. Vẫn cần đối chiếu với ảnh nguồn.")
    if len(d) and not bad.any():
        x=excel_export(d,date,factor,n); p=png_export(d,date)
        st.download_button("Tải Excel",x,f"Bao_cao_{date.replace('/','-')}.xlsx")
        st.download_button("Tải PNG",p,f"Bao_cao_{date.replace('/','-')}.png")
else: st.error("Thiếu cột dữ liệu bắt buộc.")
