# Báo cáo sức khỏe chiến binh — MVP v2
- Giao diện chia 4 tab, cấu hình bên trái.
- ZIP / nhiều ảnh / Google Drive công khai.
- Dữ liệu ảnh được giữ trong session state, không mất do thư mục tạm tạo lại ở mỗi lần Streamlit rerun.
- OCR đọc chữ và nhận diện mã NV; mapping nhóm hàng/tỷ lệ vẫn cần xác minh.
- Nhập bảng chuẩn hoặc CSV để tính và xuất Excel + PNG.
## Deploy
Upload toàn bộ file vào GitHub repository hiện tại, thay thế các file cũ. Streamlit Cloud sẽ tự khởi chạy lại.
Main file: app.py
