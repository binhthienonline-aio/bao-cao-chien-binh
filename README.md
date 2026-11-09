# Báo cáo Chiến Binh V4 — bản theo mẫu ma trận
Cập nhật lên GitHub repository đang kết nối Streamlit Cloud, ghi đè app.py và requirements.txt.
Main file: app.py

## Đầu vào
- Excel nguồn `.xlsx` có sheet `ThiDua` theo mẫu: nhóm hàng ở dòng 1 từ cột B trở đi; mã/tên nhân viên ở cột A.
- ZIP dashboard tùy chọn để xác nhận 9 mã nhân viên có trong ảnh.

## Xử lý
- Tự đọc số lượng nhóm hàng và thứ tự cột từ Excel.
- Lọc đúng 9 mã nhân viên đã chốt, theo thứ tự trong ảnh mẫu.
- Không tự đổi ô trống thành 0.
- Dòng TỔNG tính lại từ 9 nhân viên.
- Xuất Excel và PNG theo bố cục vàng/cam/xanh nhạt.

Lưu ý: doanh số phải có trong Excel nguồn; ảnh dashboard được dùng đối chiếu, không dùng OCR làm nguồn số chính để tránh sai lệch.
