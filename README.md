# BÁO CÁO SỨC KHỎE CHIẾN BINH — V7

## V7 sửa gì
- Đọc ảnh Lũy kế dạng màn hình điện thoại theo từng cột riêng: Tên chương trình, THỰC HIỆN, TARGET.
- Neo từng dòng theo tọa độ OCR của cột THỰC HIỆN để tránh lấy nhầm số THỰC HIỆN sang TARGET hoặc ghép lệch dòng.
- Dòng Đồng hồ & Smartwatch được giữ TARGET trống vì ảnh không có TARGET.
- Giữ nguyên nguyên tắc V6: phân bổ THỰC HIỆN cửa hàng theo tỷ trọng tỷ lệ Dashboard của 9 nhân viên; tổng phân bổ theo từng nhóm phải khớp THỰC HIỆN cửa hàng.
- Không dùng hệ số 110%; không lấy dữ liệu nhóm hàng/doanh số từ file Excel tháng 10.
- Các trường OCR thiếu hoặc độ tin cậy thấp phải rà soát trước khi tạo ma trận/xuất báo cáo.

## Cập nhật Streamlit Cloud
Thay các file `app.py`, `requirements.txt`, `packages.txt` trong repository GitHub đang kết nối Streamlit Cloud. Main file: `app.py`.

## Đầu vào
ZIP gồm 9 ảnh Dashboard và ảnh Lũy kế mới. Tên ảnh Lũy kế nên có cụm `LuyKe` hoặc `Lũy kế`.
