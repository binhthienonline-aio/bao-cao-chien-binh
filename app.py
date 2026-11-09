import io, os, re, tempfile, zipfile
from pathlib import Path
import pandas as pd
import streamlit as st
import gdown
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from PIL import Image, ImageOps
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

st.set_page_config(page_title="Báo cáo Chiến Binh", page_icon="📊", layout="wide")
st.markdown("""
<style>
.block-container{max-width:1600px;padding-top:1.2rem}
[data-testid="stSidebar"]{background:#f3f6fa}
.hero{background:linear-gradient(110deg,#063b70,#0879c5);padding:22px 26px;border-radius:14px;color:#fff;margin-bottom:16px}
.hero h1{color:#fff;margin:0;font-size:26px}.hero p{color:#e5f2ff;margin:6px 0 0}
div.stButton>button,div.stDownloadButton>button{border-radius:8px;font-weight:650;min-height:40px}
</style>""", unsafe_allow_html=True)
st.markdown('<div class="hero"><h1>📊 BÁO CÁO SỨC KHỎE CHIẾN BINH</h1><p>Tự nhận diện nhóm hàng • Ghép doanh số • Xuất bảng Excel và PNG theo mẫu</p></div>', unsafe_allow_html=True)

EMP_ORDER = [
    ("137501","Võ Văn Lượng"),("137502","Nguyễn Ngọc Tân"),
    ("137506","Lưu Thị Tuyết Nhi"),("170206","Lê Thị Thu Hiền"),
    ("181846","Lưu Văn Nhật"),("21512","Hà Thị Minh"),
    ("269200","Võ Quốc Cường"),("281296","Phạm Xuân Toàn"),
    ("82654","Nguyễn Văn Quyền")
]

def norm_id(v):
    s=str(v or "")
    m=re.search(r"(?<!\d)(\d{4,8})(?!\d)",s)
    return m.group(1) if m else None

def safe_num(v):
    if v is None or v == "": return None
    if isinstance(v,(int,float)): return float(v)
    s=str(v).strip().replace("\u00a0","").replace(" ","")
    if not s: return None
    # Excel cell values should be numeric; support localized numeric text conservatively.
    if "," in s and "." in s: s=s.replace(".","").replace(",",".")
    elif "," in s: s=s.replace(",",".")
    try: return float(s)
    except: return None

def find_source_sheet(wb):
    preferred=["ThiDua","THIDUA","Thi đua","Thi Đua"]
    for name in preferred:
        if name in wb.sheetnames: return wb[name]
    # fallback: sheet with a header row containing employee IDs/names and many group columns
    for ws in wb.worksheets:
        for r in range(1,min(ws.max_row,10)+1):
            vals=[ws.cell(r,c).value for c in range(1,min(ws.max_column,200)+1)]
            text=" ".join(str(x) for x in vals if x is not None).lower()
            if "hình thức thi đua" in text:
                return ws
    return None

def parse_workbook(raw):
    # data_only=True reads cached results from Excel formulas, avoiding re-calculation errors.
    wb=load_workbook(io.BytesIO(raw),data_only=True,read_only=True)
    ws=find_source_sheet(wb)
    if ws is None:
        raise ValueError("Không tìm thấy sheet ThiDua hoặc vùng dữ liệu có cấu trúc tương đương.")
    # Layout of supplied workbook: row 1 contains dynamic group names; employee rows are identified by ID.
    headers=[]
    for c in range(2,ws.max_column+1):
        h=ws.cell(1,c).value
        if h is None or str(h).strip()=="": continue
        headers.append((c,str(h).strip()))
    if not headers: raise ValueError("Không đọc được tên nhóm hàng ở dòng tiêu đề.")
    wanted={code:name for code,name in EMP_ORDER}
    rows_by_id={}
    for r in range(2,ws.max_row+1):
        label=ws.cell(r,1).value
        code=norm_id(label)
        if code in wanted:
            vals={}
            for c,h in headers:
                vals[h]=safe_num(ws.cell(r,c).value)
            rows_by_id[code]={"label":f"{wanted[code]} - {code}","values":vals,"source_row":r}
    missing=[code for code,_ in EMP_ORDER if code not in rows_by_id]
    if missing: raise ValueError("Thiếu nhân viên trong file Excel: "+", ".join(missing))
    # Drop columns that are entirely empty for the selected nine employees; preserve original order.
    keep=[]
    for c,h in headers:
        if any(rows_by_id[code]["values"].get(h) is not None for code,_ in EMP_ORDER):
            keep.append(h)
    if not keep: raise ValueError("Các cột nhóm hàng không có dữ liệu số trong 9 nhân viên.")
    records=[]
    for code,name in EMP_ORDER:
        emp=rows_by_id[code]
        records.append({"Mã NV":code,"Nhân viên":emp["label"],**{h:emp["values"].get(h) for h in keep}})
    df=pd.DataFrame(records)
    # Any blank/non-numeric source cell is missing, not zero.
    return keep,df,rows_by_id,ws.title

def validate_dashboard_zip(raw):
    names=[]
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        bad=z.testzip()
        if bad: raise ValueError(f"ZIP lỗi tại {bad}")
        for n in z.namelist():
            if Path(n).suffix.lower() in [".png",".jpg",".jpeg",".webp"]:
                names.append(Path(n).name)
    ids={norm_id(n) for n in names}
    expected={c for c,_ in EMP_ORDER}
    return names,sorted(expected-ids)

def build_xlsx(groups,df,report_date):
    wb=Workbook()
    ws=wb.active; ws.title="Báo cáo"
    n=len(groups); last_col=n+1
    # Header row and exact three-band visual hierarchy from the reference image.
    ws.cell(1,1,40 if n==40 else n)
    for j,g in enumerate(groups,2): ws.cell(1,j,g)
    ws.cell(2,1,"TỔNG")
    for j,g in enumerate(groups,2):
        vals=pd.to_numeric(df[g],errors="coerce")
        ws.cell(2,j,float(vals.sum()) if vals.notna().all() else None)
    for i,row in enumerate(df.itertuples(index=False),3):
        ws.cell(i,1,row[1])
        for j,g in enumerate(groups,2):
            v=getattr(row,g) if hasattr(row,g) else df.iloc[i-3][g]
            ws.cell(i,j,float(v) if pd.notna(v) else None)
    yellow="FFF200"; orange="FFC000"; blue="DDEBF7"; black="000000"
    thin=Side(style="thin",color="202020")
    for cell in ws[1][:last_col]:
        cell.fill=PatternFill("solid",fgColor=yellow); cell.font=Font(name="Arial",size=8,bold=True,color=black)
        cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        cell.border=Border(left=thin,right=thin,top=thin,bottom=thin)
    for cell in ws[2][:last_col]:
        cell.fill=PatternFill("solid",fgColor=orange); cell.font=Font(name="Arial",size=9,bold=True,color=black)
        cell.alignment=Alignment(horizontal="center",vertical="center")
        cell.border=Border(left=thin,right=thin,top=thin,bottom=thin)
        if cell.column>1: cell.number_format="#,##0.00"
    for r in range(3,12):
        for c in range(1,last_col+1):
            cell=ws.cell(r,c); cell.fill=PatternFill("solid",fgColor=blue)
            cell.font=Font(name="Arial",size=9,bold=(c==1),color=black)
            cell.alignment=Alignment(horizontal="left" if c==1 else "center",vertical="center")
            cell.border=Border(left=thin,right=thin,top=thin,bottom=thin)
            if c>1: cell.number_format="#,##0.00"
    ws.column_dimensions["A"].width=27
    for c in range(2,last_col+1): ws.column_dimensions[get_column_letter(c)].width=12
    ws.row_dimensions[1].height=62; ws.row_dimensions[2].height=22
    for r in range(3,12): ws.row_dimensions[r].height=19
    ws.freeze_panes="B3"
    ws.sheet_view.showGridLines=False
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_setup.orientation="landscape"; ws.page_setup.paperSize=ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=1
    ws.print_options.horizontalCentered=True
    ws.print_area=f"A1:{get_column_letter(last_col)}11"
    # Audit sheet records missing values instead of silently converting them to zero.
    audit=wb.create_sheet("Kiểm tra")
    audit.append(["Ngày báo cáo",report_date])
    audit.append(["Số nhóm hàng",n])
    audit.append(["Số nhân viên",9])
    audit.append(["Ô dữ liệu trống",int(df[groups].isna().sum().sum())])
    audit.append(["Ghi chú","Ô trống trong nguồn Excel được giữ trống, không tự đổi thành 0."])
    b=io.BytesIO(); wb.save(b); b.seek(0); return b

def build_png(groups,df,report_date):
    n=len(groups); rows=11
    fig_w=max(18,1.45*(n+1))
    fig_h=7.2
    fig,ax=plt.subplots(figsize=(fig_w,fig_h))
    ax.set_xlim(0,n+1.05); ax.set_ylim(0,11); ax.invert_yaxis(); ax.axis("off")
    colw=[2.1]+[1.0]*n
    x=[0]
    for w in colw: x.append(x[-1]+w)
    # Header row
    for c in range(n+1):
        ax.add_patch(Rectangle((x[c],0),colw[c],1,facecolor="#FFF200",edgecolor="#202020",linewidth=.5))
    ax.text(x[0]+colw[0]/2,.5,str(40 if n==40 else n),ha="center",va="center",fontsize=9,weight="bold")
    for j,g in enumerate(groups,1):
        ax.text(x[j]+colw[j]/2,.5,g,ha="center",va="center",fontsize=4.4,weight="bold",wrap=True)
    # Total row
    for c in range(n+1):
        ax.add_patch(Rectangle((x[c],1),colw[c],1,facecolor="#FFC000",edgecolor="#202020",linewidth=.5))
    ax.text(x[0]+colw[0]/2,1.5,"TỔNG",ha="center",va="center",fontsize=8,weight="bold")
    for j,g in enumerate(groups,1):
        vals=pd.to_numeric(df[g],errors="coerce")
        v=vals.sum() if vals.notna().all() else float("nan")
        ax.text(x[j]+colw[j]/2,1.5,f"{v:.2f}" if pd.notna(v) else "",ha="center",va="center",fontsize=5.2,weight="bold")
    # Employee rows
    for i,row in df.iterrows():
        y=i+2
        for c in range(n+1):
            ax.add_patch(Rectangle((x[c],y),colw[c],1,facecolor="#DDEBF7",edgecolor="#202020",linewidth=.45))
        ax.text(x[0]+.04,y+.5,row["Nhân viên"],ha="left",va="center",fontsize=6.1)
        for j,g in enumerate(groups,1):
            v=row[g]
            ax.text(x[j]+colw[j]/2,y+.5,f"{v:.2f}" if pd.notna(v) else "",ha="center",va="center",fontsize=5.1)
    ax.set_xlim(0,x[-1]); ax.set_title(f"BÁO CÁO SỨC KHỎE CHIẾN BINH — {report_date}",fontsize=11,weight="bold",pad=10)
    fig.tight_layout()
    b=io.BytesIO(); fig.savefig(b,format="png",dpi=220,bbox_inches="tight"); plt.close(fig); b.seek(0); return b

with st.sidebar:
    st.markdown("### CẤU HÌNH")
    report_date=st.text_input("Ngày báo cáo","09/11/2026")
    st.caption("Cố định 9 nhân viên theo thứ tự ảnh mẫu.")
    st.caption("Số nhóm hàng tự nhận diện từ file Excel.")

tabs=st.tabs(["① NẠP DỮ LIỆU","② KIỂM TRA NGUỒN","③ XEM TRƯỚC BẢNG","④ XUẤT EXCEL / PNG"])
if "source_data" not in st.session_state: st.session_state.source_data=None
with tabs[0]:
    left,right=st.columns(2)
    with left:
        xl=st.file_uploader("File Excel nguồn (.xlsx)",type=["xlsx"])
        if st.button("ĐỌC FILE EXCEL",type="primary",disabled=xl is None):
            try:
                groups,df,rows_by_id,sheet=parse_workbook(xl.getvalue())
                st.session_state.source_data={"groups":groups,"df":df,"sheet":sheet,"rows":rows_by_id}
                st.success(f"Đã đọc {len(groups)} nhóm hàng và 9 nhân viên từ sheet {sheet}.")
            except Exception as e: st.error(str(e))
    with right:
        z=st.file_uploader("ZIP ảnh dashboard (để đối chiếu)",type=["zip"])
        if st.button("KIỂM TRA ZIP",disabled=z is None):
            try:
                names,missing=validate_dashboard_zip(z.getvalue())
                st.session_state.zip_check={"names":names,"missing":missing}
                st.success(f"Nhận diện {len(names)} ảnh. Thiếu mã: {', '.join(missing) if missing else 'Không thiếu mã trong tên file'}.")
            except Exception as e: st.error(str(e))
    st.markdown("**Hoặc tải Excel từ Google Drive công khai**")
    url=st.text_input("Link Google Drive file Excel")
    if st.button("TẢI EXCEL TỪ DRIVE",disabled=not url):
        try:
            tmp=tempfile.mkdtemp()
            if "/folders/" in url:
                # Public Drive folder: download its contents, then select Excel workbooks.
                gdown.download_folder(url=url, output=tmp, quiet=True, use_cookies=False)
                candidates=list(Path(tmp).rglob("*.xlsx"))
                candidates += list(Path(tmp).rglob("*.xlsm"))
                if not candidates:
                    raise ValueError("Thư mục Drive không có file Excel .xlsx/.xlsm. Hãy kiểm tra quyền xem công khai.")
                if len(candidates)>1:
                    st.session_state.drive_excel_candidates=[str(x) for x in candidates]
                    st.info("Thư mục có nhiều file Excel. Chọn đúng file ở danh sách bên dưới.")
                else:
                    chosen=candidates[0]
                    groups,df,rows_by_id,sheet=parse_workbook(chosen.read_bytes())
                    st.session_state.source_data={"groups":groups,"df":df,"sheet":sheet,"rows":rows_by_id}
                    st.success(f"Đã đọc {len(groups)} nhóm hàng và 9 nhân viên từ {chosen.name}.")
            else:
                # File URL: omit fuzzy for compatibility with older gdown versions.
                p=gdown.download(url=url, output=os.path.join(tmp,"source.xlsx"), quiet=True)
                if not p: raise ValueError("Không tải được file. Kiểm tra quyền chia sẻ công khai.")
                groups,df,rows_by_id,sheet=parse_workbook(Path(p).read_bytes())
                st.session_state.source_data={"groups":groups,"df":df,"sheet":sheet,"rows":rows_by_id}
                st.success(f"Đã đọc {len(groups)} nhóm hàng và 9 nhân viên.")
        except Exception as e: st.error(f"Lỗi Google Drive: {e}")
    if st.session_state.source_data:
        st.metric("Nhóm hàng tự nhận diện",len(st.session_state.source_data["groups"]))
        st.metric("Nhân viên",9)
    if st.session_state.get("drive_excel_candidates"):
        candidate_paths=st.session_state.drive_excel_candidates
        labels=[Path(x).name for x in candidate_paths]
        chosen_label=st.selectbox("Chọn file Excel trong thư mục Drive",labels)
        if st.button("ĐỌC FILE ĐÃ CHỌN"):
            chosen=Path(candidate_paths[labels.index(chosen_label)])
            try:
                groups,df,rows_by_id,sheet=parse_workbook(chosen.read_bytes())
                st.session_state.source_data={"groups":groups,"df":df,"sheet":sheet,"rows":rows_by_id}
                st.success(f"Đã đọc {len(groups)} nhóm hàng và 9 nhân viên từ {chosen.name}.")
            except Exception as e: st.error(str(e))

with tabs[1]:
    src=st.session_state.source_data
    if not src: st.info("Nạp file Excel ở bước ① trước.")
    else:
        df=src["df"]; groups=src["groups"]
        st.write(f"**Nguồn:** sheet `{src['sheet']}`")
        st.write(f"**Nhóm hàng:** {len(groups)}")
        st.write("**Thứ tự nhân viên:**")
        st.dataframe(df[["Nhân viên"]],use_container_width=True,hide_index=True)
        blanks=int(df[groups].isna().sum().sum())
        st.metric("Ô doanh số trống trong Excel",blanks)
        if blanks: st.warning("Ô trống được giữ trống, không tự chuyển thành 0. Cần xác minh nguồn.")
        st.dataframe(df,use_container_width=True,height=360)
        if st.session_state.get("zip_check"):
            st.write("**Đối chiếu ZIP ảnh:**")
            st.write("Ảnh nhận diện:",len(st.session_state.zip_check["names"]))
            st.write("Mã nhân viên thiếu:",st.session_state.zip_check["missing"] or "Không thiếu")
            st.caption("Ở bản này ảnh dashboard dùng để xác nhận đủ nhân viên; dữ liệu doanh số lấy từ Excel nguồn để tránh OCR sai số.")
with tabs[2]:
    src=st.session_state.source_data
    if not src: st.info("Nạp Excel ở bước ① trước.")
    else:
        groups=src["groups"]; df=src["df"]
        st.markdown("**Bản xem trước — có dòng TỔNG, đúng thứ tự nhóm hàng và 9 nhân viên đã chốt.**")
        preview=df[["Nhân viên"]+groups].copy()
        total_row={"Nhân viên":"TỔNG"}
        for g in groups:
            vals=pd.to_numeric(df[g],errors="coerce")
            total_row[g]=float(vals.sum()) if vals.notna().all() else None
        preview=pd.concat([pd.DataFrame([total_row]),preview],ignore_index=True)
        def color_preview_row(row):
            bg = "#FFC000" if row.name == 0 else "#DDEBF7"
            weight = "bold" if row.name == 0 else "normal"
            return [f"background-color: {bg}; color: #000000; font-weight: {weight}; border: 1px solid #202020;"]*len(row)
        styled=(preview.style
                .apply(color_preview_row,axis=1)
                .set_table_styles([
                    {"selector":"th","props":[("background-color","#FFF200"),("color","#000000"),("font-weight","bold"),("border","1px solid #202020"),("white-space","normal"),("min-width","110px")]},
                    {"selector":"td","props":[("border","1px solid #202020"),("text-align","center"),("padding","4px")]},
                    {"selector":"th:first-child, td:first-child","props":[("min-width","190px"),("text-align","left")]}
                ])
                .format({g:"{:,.2f}" for g in groups},na_rep=""))
        st.dataframe(styled,use_container_width=True,height=520,hide_index=True)
        st.caption("Màu vàng: tiêu đề nhóm hàng • Màu cam: dòng TỔNG • Xanh nhạt: dữ liệu 9 nhân viên.")
with tabs[3]:
    src=st.session_state.source_data
    if not src: st.info("Nạp Excel ở bước ① trước.")
    else:
        groups=src["groups"]; df=src["df"]
        blanks=int(df[groups].isna().sum().sum())
        if blanks:
            st.warning(f"Có {blanks} ô trống trong nguồn Excel. Chúng sẽ được giữ trống; hãy xác minh trước khi chốt.")
        else:
            st.success(f"Dữ liệu đã đủ {9*len(groups)} ô.")
        x=build_xlsx(groups,df,report_date)
        p=build_png(groups,df,report_date)
        a,b=st.columns(2)
        a.download_button("⬇ TẢI EXCEL",x,f"Bao_cao_Chien_Binh_{report_date.replace('/','-')}.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
        b.download_button("⬇ TẢI PNG",p,f"Bao_cao_Chien_Binh_{report_date.replace('/','-')}.png","image/png",use_container_width=True)
