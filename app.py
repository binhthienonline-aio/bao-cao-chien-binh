import io, os, re, zipfile, tempfile, unicodedata
from pathlib import Path
import pandas as pd
import streamlit as st
from PIL import Image, ImageOps, ImageEnhance
import pytesseract
import gdown
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import matplotlib.pyplot as plt

st.set_page_config(page_title="Chiến Binh | Báo cáo",page_icon="📊",layout="wide")
st.markdown("""
<style>
.block-container{padding-top:1.2rem;max-width:1500px}
[data-testid="stSidebar"]{background:#f3f6fa}
.hero{background:linear-gradient(115deg,#063b70,#0879c5);padding:24px 28px;border-radius:16px;color:white;margin-bottom:15px}
.hero h1{color:white;margin:0;font-size:27px}.hero p{color:#dceeff;margin:6px 0 0}
div.stButton>button,div.stDownloadButton>button{border-radius:9px;font-weight:650;min-height:42px}
</style>""",unsafe_allow_html=True)
st.markdown('<div class="hero"><h1>📊 BÁO CÁO SỨC KHỎE CHIẾN BINH</h1><p>Tự đọc dashboard • Rà soát số liệu • Xuất Excel và PNG</p></div>',unsafe_allow_html=True)

# TARGET lấy từ ảnh Lũy kế 09/11, cột TARGET (không phải Mục tiêu/NV)
GROUPS=[
("Máy sấy / Máy rửa chén",23.58),("Bảo hiểm thọ ICT",168.70),("Tivi TCL",111.00),
("Tủ lạnh / Đông / Mát",436.86),("Tai nghe",11.98),("Sim",57.00),("Quạt gió",17.90),
("Máy lọc KK / Hút bụi / Hút ẩm",118.87),("CE LG",706.00),("Sim DMX",43.00),
("Bảo hiểm thọ DMX",382.69),("Điện thoại Vivo",151.73),("Nạp rút",225.00),
("Trả chậm FE Credit / Samsung F+ / Shinhan / MCredit",1190.46),("Máy giặt",420.95),
("Loa",14.20),("Điện tử",776.00),("Laptop",164.84),("VAS",202.00),
("Nồi cơm & Nồi chiên",183.00),("Camera",28.85),("Phụ kiện khác IT & Nhóm khác",150.26),
("Trả chậm điện máy & Gia dụng",1668.53),("Máy lạnh",39.00),("Trả chậm Home Credit",1009.70),
("Điện thoại & Tablet Android",1036.32),("Điện tử Samsung",108.00),("Vay tiền mặt",109.56),
("Cáp sạc",19.11),("Máy lọc nước",298.52),("Ví trả sau",163.24),("Máy nước nóng",52.00),
("Sạc dự phòng",43.40),("Điện thoại Realme",253.31),("Bảo hiểm tổng",246.89),
("Đèn năng lượng mặt trời",9.00),("Mở thẻ tín dụng VPBank MWG & TPBank",1.00),("Phụ kiện công nghệ",4.01)
]
EMP=[
("82654","Nguyễn Văn Quyền"),("137502","Nguyễn Ngọc Tân"),("181846","Lưu Văn Nhật"),
("137501","Võ Văn Lượng"),("137506","Lưu Thị Tuyết Nhi"),("170206","Lê Thị Thu Hiền"),
("281296","Phạm Xuân Toàn"),("21512","Hà Thị Minh"),("269200","Võ Quốc Cường")
]
ALIASES={
"Máy sấy / Máy rửa chén":["may say","may rua chen"],
"Bảo hiểm thọ ICT":["bao hiem tho ict","bao hiem thọ ict"],
"Tivi TCL":["tivi tcl"],
"Tủ lạnh / Đông / Mát":["tu lanh","dong","mat"],
"Tai nghe":["tai nghe"],"Sim":["sim"],"Quạt gió":["quat gio"],
"Máy lọc KK / Hút bụi / Hút ẩm":["may loc kk","hut bui","hut am"],
"CE LG":["ce lg","celg"],"Sim DMX":["sim dmx"],
"Bảo hiểm thọ DMX":["bao hiem tho dmx","bao hiem thọ dmx"],
"Điện thoại Vivo":["dien thoai vivo"],"Nạp rút":["nap rut"],
"Trả chậm FE Credit / Samsung F+ / Shinhan / MCredit":["tra cham fe credit","samsung f","shinhan","mcredit"],
"Máy giặt":["may giat"],"Loa":["loa"],"Điện tử":["dien tu"],
"Laptop":["laptop"],"VAS":["vas"],"Nồi cơm & Nồi chiên":["noi com","noi chien"],
"Camera":["camera"],"Phụ kiện khác IT & Nhóm khác":["phu kien khac it","nhom khac"],
"Trả chậm điện máy & Gia dụng":["tra cham dien may","gia dung"],
"Máy lạnh":["may lanh"],"Trả chậm Home Credit":["tra cham homecredit","homecredit"],
"Điện thoại & Tablet Android":["dien thoai","tablet android"],
"Điện tử Samsung":["dien tu samsung"],"Vay tiền mặt":["vay tien mat"],
"Cáp sạc":["cap sac"],"Máy lọc nước":["may loc nuoc"],
"Ví trả sau":["vi tra sau"],"Máy nước nóng":["may nuoc nong"],
"Sạc dự phòng":["sac du phong"],"Điện thoại Realme":["dien thoai realme"],
"Bảo hiểm tổng":["bao hiem tong"],"Đèn năng lượng mặt trời":["den nang luong mat troi"],
"Mở thẻ tín dụng VPBank MWG & TPBank":["mo the tin dung vpbank"],
"Phụ kiện công nghệ":["phu kien cong nghe"]
}
def norm(s):
    s=unicodedata.normalize("NFD",str(s).lower())
    s="".join(c for c in s if unicodedata.category(c)!="Mn")
    s=s.replace("đ","d")
    return re.sub(r"[^a-z0-9%]+"," ",s).strip()
def get_images_zip(raw):
    out=[]
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        if z.testzip(): raise ValueError("ZIP bị lỗi.")
        for i in z.infolist():
            if Path(i.filename).suffix.lower() in [".png",".jpg",".jpeg",".webp"]:
                out.append((Path(i.filename).name,z.read(i)))
    return out
def ocr(raw):
    im=Image.open(io.BytesIO(raw))
    im=ImageOps.exif_transpose(im).convert("L")
    im=ImageEnhance.Contrast(im).enhance(1.7)
    if im.width<1500:
        k=1500/im.width; im=im.resize((int(im.width*k),int(im.height*k)))
    return pytesseract.image_to_string(im,lang="eng",config="--psm 6")
def match_employee(name,text):
    s=name+" "+text
    return next((e for e in EMP if e[0] in s),None)
def parse_percentages(text):
    # Heuristic card parser. Keep uncertain/missing values as blank for human review.
    lines=[norm(x) for x in text.splitlines() if x.strip()]
    out={}
    for group,aliases in ALIASES.items():
        for i,line in enumerate(lines):
            if any(norm(a) in line for a in aliases):
                context=" ".join(lines[max(0,i-1):min(len(lines),i+3)])
                nums=re.findall(r"(?<!\d)(\d{1,3}(?:\.\d{1,2})?)\s*%?",context)
                vals=[]
                for n in nums:
                    try:
                        v=float(n)
                        if 0<=v<=1000: vals.append(v)
                    except: pass
                if vals:
                    out[group]=vals[-1]
                    break
    # No-sales list from dashboard: categories explicitly named in this section are zero.
    ntext=norm(text)
    if "chua co so lieu" in ntext:
        part=ntext.split("chua co so lieu",1)[1].split("sieu thi khong co target",1)[0]
        for g,aliases in ALIASES.items():
            if any(norm(a) in part for a in aliases): out[g]=0.0
    return out
def make_xlsx(df,date,factor,staff):
    wb=Workbook(); ws=wb.active; ws.title="KetQua"
    ws.merge_cells("A1:G1"); ws["A1"]="BÁO CÁO SỨC KHỎE CHIẾN BINH"; ws["A1"].font=Font(size=16,bold=True,color="FFFFFF"); ws["A1"].fill=PatternFill("solid",fgColor="075A9C")
    ws.merge_cells("A2:G2"); ws["A2"]=f"Ngày {date} | Hệ số {factor:.0%} | Nhân viên {staff}"
    heads=["Mã NV","Nhân viên","Nhóm hàng","TARGET siêu thị","Tỷ lệ hoàn thành (%)","Mục tiêu cá nhân","Doanh số quy đổi"]
    for j,h in enumerate(heads,1):
        c=ws.cell(4,j,h); c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="1677C8")
    for i,row in enumerate(df.itertuples(index=False),5):
        for j,v in enumerate(row,1): ws.cell(i,j,v)
        for j in [4,6,7]: ws.cell(i,j).number_format="#,##0.00"
        ws.cell(i,5).number_format='0.00"%"'
    for c,w in {"A":13,"B":25,"C":42,"D":20,"E":22,"F":20,"G":20}.items(): ws.column_dimensions[c].width=w
    ws.freeze_panes="C5"; ws.auto_filter.ref=f"A4:G{4+len(df)}"
    audit=wb.create_sheet("KiemTra")
    audit.append(["Mã NV","Nhân viên","Nhóm hàng","Tỷ lệ OCR","Trạng thái"])
    for r in df.itertuples(index=False): audit.append([r[0],r[1],r[2],r[4], "Cần xác minh ảnh gốc" if pd.isna(r[4]) else "OCR trích xuất"])
    for c in audit[1]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="475569")
    b=io.BytesIO(); wb.save(b); b.seek(0); return b
def make_png(df,date):
    pv=df.pivot_table(index="Nhân viên",columns="Nhóm hàng",values="Tỷ lệ hoàn thành (%)",aggfunc="first",fill_value=0)
    fig,ax=plt.subplots(figsize=(max(15,len(pv.columns)*.5),max(5,len(pv)*.48)))
    mx=max(100,float(pv.fillna(0).values.max()))
    im=ax.imshow(pv.fillna(0).values,aspect="auto",cmap="RdYlGn",vmin=0,vmax=mx)
    ax.set_xticks(range(len(pv.columns))); ax.set_xticklabels(pv.columns,rotation=65,ha="right",fontsize=6)
    ax.set_yticks(range(len(pv.index))); ax.set_yticklabels(pv.index,fontsize=8)
    ax.set_title(f"TỶ LỆ HOÀN THÀNH — {date}",fontweight="bold")
    fig.colorbar(im,ax=ax,label="%"); fig.tight_layout()
    b=io.BytesIO(); fig.savefig(b,format="png",dpi=180,bbox_inches="tight"); plt.close(fig); b.seek(0); return b

with st.sidebar:
    st.markdown("### CẤU HÌNH")
    date=st.text_input("Ngày báo cáo","09/11/2026")
    factor=st.number_input("Hệ số mục tiêu",0.0,5.0,1.10,.05)
    staff=st.number_input("Số nhân viên",1,100,9,1)
    st.caption("Mục tiêu cá nhân = TARGET × hệ số / số nhân viên")
tabs=st.tabs(["① NẠP ẢNH","② TỰ ĐỌC & GHÉP","③ RÀ SOÁT DỮ LIỆU","④ XUẤT BÁO CÁO"])
if "images" not in st.session_state: st.session_state.images=[]
if "parsed" not in st.session_state: st.session_state.parsed=None
with tabs[0]:
    c1,c2=st.columns(2)
    with c1:
        z=st.file_uploader("Tải ZIP gồm 9 dashboard + ảnh lũy kế",type=["zip"])
        imgs=st.file_uploader("Hoặc tải ảnh riêng",type=["png","jpg","jpeg"],accept_multiple_files=True)
        if st.button("NẠP ẢNH",type="primary"):
            try:
                got=get_images_zip(z.getvalue()) if z else []
                got += [(f.name,f.getvalue()) for f in (imgs or [])]
                st.session_state.images=got; st.session_state.parsed=None
                st.success(f"Đã nạp {len(got)} ảnh.")
            except Exception as e: st.error(str(e))
    with c2:
        url=st.text_input("Link Google Drive công khai")
        if st.button("TẢI TỪ DRIVE"):
            try:
                tmp=tempfile.mkdtemp()
                if "/folders/" in url:
                    gdown.download_folder(url=url,output=tmp,quiet=True,use_cookies=False)
                    got=[]
                    for p in Path(tmp).rglob("*"):
                        if p.is_file() and p.suffix.lower()==".zip": got+=get_images_zip(p.read_bytes())
                        elif p.is_file() and p.suffix.lower() in [".png",".jpg",".jpeg"]: got.append((p.name,p.read_bytes()))
                else:
                    p=gdown.download(url=url,output=os.path.join(tmp,"source"),quiet=True,fuzzy=True)
                    got=get_images_zip(Path(p).read_bytes()) if p and p.endswith(".zip") else ([(Path(p).name,Path(p).read_bytes())] if p else [])
                st.session_state.images=got; st.session_state.parsed=None
                st.success(f"Đã tải {len(got)} ảnh.")
            except Exception as e: st.error(f"Không tải được Drive: {e}")
    st.metric("Tổng ảnh đã nạp",len(st.session_state.images))
    if st.session_state.images:
        for name,_ in st.session_state.images: st.write("•",name)
with tabs[1]:
    st.markdown("### Hệ thống tự đọc TARGET và dashboard")
    st.warning("Bản V3 dùng OCR + đối chiếu tên nhóm. Mọi giá trị OCR đều cần rà soát; nhóm chưa đọc rõ để trống, không tự đoán.")
    if st.button("CHẠY TRÍCH XUẤT TỰ ĐỘNG",type="primary",disabled=not st.session_state.images):
        imgs={name:raw for name,raw in st.session_state.images}
        target_names=[n for n in imgs if "luy" in norm(n) or "ke" in norm(n)]
        if not target_names: st.error("Không tìm thấy ảnh Lũy kế. Hãy nạp ZIP có ảnh này.")
        else:
            target_text=ocr(imgs[target_names[0]])
            empfiles={}
            for name,raw in imgs.items():
                emp=match_employee(name,"")
                if emp: empfiles[emp[0]]=(emp,raw,name)
            rows=[]
            progress=st.progress(0)
            for idx,(code,name) in enumerate(EMP):
                if code not in empfiles:
                    rows.extend([[code,name,g,t,None] for g,t in GROUPS])
                else:
                    emp,raw,fn=empfiles[code]
                    txt=ocr(raw); vals=parse_percentages(txt)
                    rows.extend([[code,name,g,t,vals.get(g,None)] for g,t in GROUPS])
                progress.progress((idx+1)/len(EMP))
            df=pd.DataFrame(rows,columns=["Mã NV","Nhân viên","Nhóm hàng","TARGET siêu thị","Tỷ lệ hoàn thành (%)"])
            df["Mục tiêu cá nhân"]=df["TARGET siêu thị"]*factor/staff
            df["Doanh số quy đổi"]=df["Mục tiêu cá nhân"]*pd.to_numeric(df["Tỷ lệ hoàn thành (%)"],errors="coerce")/100
            st.session_state.parsed=df
            st.success(f"Đã tạo ma trận {len(df)} dòng (9 nhân viên × 38 nhóm hàng).")
            st.info("Lưu ý: tên nhóm và tỷ lệ được ghép theo quy tắc OCR. Những ô trống cần rà soát, không được mặc định là 0.")
    if st.session_state.parsed is not None:
        st.dataframe(st.session_state.parsed,use_container_width=True,height=500)
with tabs[2]:
    st.markdown("### Rà soát trước khi chốt")
    if st.session_state.parsed is None: st.info("Hãy chạy trích xuất ở tab ② trước.")
    else:
        df=st.session_state.parsed.copy()
        st.caption("Sửa trực tiếp tỷ lệ hoàn thành. Ô trống nghĩa là OCR chưa đọc chắc chắn; nếu dashboard xác nhận không có doanh số, nhập 0.")
        edited=st.data_editor(df,use_container_width=True,num_rows="fixed",height=600,key="review_grid")
        edited["Mục tiêu cá nhân"]=edited["TARGET siêu thị"]*factor/staff
        edited["Doanh số quy đổi"]=pd.to_numeric(edited["Tỷ lệ hoàn thành (%)"],errors="coerce")*edited["Mục tiêu cá nhân"]/100
        st.session_state.final_df=edited
        missing=edited["Tỷ lệ hoàn thành (%)"].isna().sum()
        st.metric("Ô tỷ lệ còn thiếu",int(missing))
        if missing: st.warning("Còn ô chưa xác minh. Không nên xuất báo cáo chốt.")
        else: st.success("Đã điền đủ tỷ lệ. Hãy đối chiếu ngẫu nhiên với ảnh gốc trước khi xuất.")
with tabs[3]:
    st.markdown("### Tải báo cáo Excel và PNG")
    df=st.session_state.get("final_df")
    if df is None: st.info("Hãy hoàn tất bước rà soát trước.")
    else:
        miss=df["Tỷ lệ hoàn thành (%)"].isna().sum()
        if miss: st.error(f"Còn {miss} ô chưa xác minh. Vui lòng quay lại tab ③.")
        else:
            st.success("Dữ liệu đã đủ ô. Vẫn cần xác nhận số liệu theo ảnh nguồn.")
            x=make_xlsx(df[["Mã NV","Nhân viên","Nhóm hàng","TARGET siêu thị","Tỷ lệ hoàn thành (%)","Mục tiêu cá nhân","Doanh số quy đổi"]],date,factor,staff)
            p=make_png(df,date)
            a,b=st.columns(2)
            a.download_button("TẢI EXCEL",x,f"Bao_cao_{date.replace('/','-')}.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            b.download_button("TẢI PNG",p,f"Bao_cao_{date.replace('/','-')}.png","image/png",use_container_width=True)
