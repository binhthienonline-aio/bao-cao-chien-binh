# BÁO CÁO CHIẾN BINH V6 — IMAGE ONLY

## Quy tắc dữ liệu tháng 11
- Chỉ sử dụng ảnh Lũy kế tháng 11 và 9 ảnh Dashboard tháng 11. Không đọc hoặc kế thừa tên nhóm hàng/doanh số từ file Excel tháng 10.
- Ảnh Lũy kế mới: OCR tên chương trình, cột **THỰC HIỆN** và **TARGET**. TARGET được giữ để đối chiếu; chương trình không có TARGET (ví dụ Đồng hồ & Smartwatch) được để trống, không tự gán số.
- 9 Dashboard: OCR tỷ lệ hoàn thành từng nhóm của từng nhân viên. Các ô chưa đọc chắc chắn phải được rà soát; chỉ gán 0 khi Dashboard có căn cứ thể hiện chưa có số liệu.

## Công thức phân bổ doanh số
Với từng nhóm hàng:

`Doanh số NV_i = THỰC HIỆN cửa hàng × Tỷ lệ Dashboard NV_i / Tổng tỷ lệ Dashboard của 9 NV`

- Không dùng hệ số 110%.
- Không lấy TARGET chia 9 để suy ra doanh số.
- Khi đủ tỷ lệ của cả 9 nhân viên, tổng doanh số phân bổ của 9 nhân viên sẽ khớp THỰC HIỆN cửa hàng (sai số làm tròn chỉ ở bước hiển thị).
- Nếu thiếu tỷ lệ, tổng tỷ lệ bằng 0 trong khi THỰC HIỆN khác 0, hoặc dữ liệu nguồn chưa xác minh, hệ thống để trống kết quả và cảnh báo; không tự chia đều.
- Có bảng đối soát theo từng nhóm giữa THỰC HIỆN cửa hàng và tổng doanh số 9 nhân viên.

## Cập nhật Streamlit
Upload/ghi đè `app.py`, `requirements.txt`, `packages.txt` vào repository GitHub đang kết nối Streamlit Cloud. Main file: `app.py`.

## Đầu vào
Một ZIP gồm 9 ảnh Dashboard (tên file có mã nhân viên) và 1 ảnh Lũy kế mới. Có thể tải ZIP/ảnh trực tiếp hoặc dùng link Google Drive công khai.

## Lưu ý OCR
OCR ảnh điện thoại có thể nhầm dấu thập phân hoặc tên nhóm. Bắt buộc rà soát các cột tên nhóm, THỰC HIỆN, TARGET và ma trận tỷ lệ trước khi xuất báo cáo chính thức.
