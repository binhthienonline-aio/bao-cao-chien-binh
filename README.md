# BÁO CÁO SỨC KHỎE CHIẾN BINH — V8

## Sửa lỗi chính
- Sửa lỗi regex làm tên nhóm OCR không được chuẩn hóa đúng.
- Ưu tiên ảnh Lũy kế điện thoại mới nếu Drive còn ảnh cũ trùng tên.
- Chỉ chấp nhận ảnh Lũy kế có đủ cột CHƯƠNG TRÌNH, THỰC HIỆN, TARGET, % HT.
- Kiểm tra đủ 39 dòng THỰC HIỆN trước khi xử lý tiếp; thiếu/thừa dòng thì dừng, không tự suy đoán.
- Không dùng ảnh Excel tháng 10 để lấy nhóm hàng hoặc số liệu tháng 11.
- Không áp dụng hệ số 110%.
- Doanh số 9 nhân viên phân bổ theo tỷ trọng Dashboard, tổng theo từng nhóm phải khớp THỰC HIỆN cửa hàng.

## Cập nhật Streamlit Cloud
Thay app.py, requirements.txt, packages.txt trong GitHub repo đang kết nối Streamlit Cloud.
