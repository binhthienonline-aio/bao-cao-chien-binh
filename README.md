# BÁO CÁO CHIẾN BINH V5 — LŨY KẾ × TỶ LỆ DASHBOARD

## Công thức chính thức
- **Doanh số nhân viên theo nhóm = LŨY KẾ cửa hàng của nhóm × tỷ lệ hoàn thành của nhân viên / 100.**
- Ví dụ: nhóm Sim có LŨY KẾ 20.00; nhân viên có tỷ lệ 15% → doanh số quy đổi = **3.00**.
- Cột TARGET chỉ được đọc và hiển thị để tham khảo; **không tham gia công thức doanh số**.
- Không chia TARGET cho 9 nhân viên, không áp dụng hệ số 110%, không ép tổng tỷ lệ về 100%.
- Nếu tổng tỷ lệ 9 nhân viên của một nhóm khác 100%, ứng dụng hiển thị cảnh báo và chênh lệch; không tự điều chỉnh dữ liệu.

## Nguồn dữ liệu
- Ảnh Lũy kế tháng 11: đọc động tên nhóm, **cột LŨY KẾ** và TARGET theo từng dòng. Hỗ trợ ảnh chụp dọc bảng có 38 dòng, kể cả các dòng tên nhóm xuống dòng.
- 9 ảnh Dashboard: phát hiện từng thẻ trắng trong vùng THI ĐUA, OCR nhãn và tỷ lệ.
- Danh sách "Chưa có số liệu" được đọc riêng; chỉ nhóm ghép được từ danh sách này mới tự gán 0.
- Nhóm không ghép chắc chắn để trống; không tự suy đoán.

## Đối soát
- Tab xem trước có bảng phân bổ doanh số theo nhóm.
- Ứng dụng đối chiếu tổng tỷ lệ 9 nhân viên và tổng doanh số đã phân bổ với LŨY KẾ gốc.
- Ô OCR chưa xác minh để trống, không tự đổi thành 0.

## Cập nhật Streamlit
Upload/ghi đè `app.py`, `requirements.txt`, `packages.txt` vào repository GitHub đang kết nối Streamlit Cloud.
Main file: `app.py`.

## File đầu vào
Một ZIP gồm 9 PNG Dashboard (tên file có mã nhân viên) và 1 ảnh bảng Lũy kế. Tên ảnh Lũy kế nên có chữ `Lũy kế`/`Luy ke`; nếu tên file không có, ứng dụng sẽ nhận ảnh thừa duy nhất ngoài 9 ảnh nhân viên làm ảnh Lũy kế.
