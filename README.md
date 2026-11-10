# BÁO CÁO CHIẾN BINH V5.3 — LŨY KẾ HEADER FIX

## Nền tảng và nguyên tắc dữ liệu
- Giữ nguyên quy trình V5 đã chốt: ảnh Lũy kế là nguồn doanh số gốc; 9 ảnh Dashboard là nguồn tỷ lệ phân bổ.
- Không dùng file Excel tháng 10; không áp dụng hệ số 110%.
- Doanh số nhân viên theo nhóm = LŨY KẾ cửa hàng của nhóm × tỷ lệ hoàn thành Dashboard của nhân viên / 100.
- Chỉ lấy số trong cột LŨY KẾ; không đọc TARGET hoặc các cột kế bên để thay thế.

## Sửa lỗi nhận diện tiêu đề LŨY KẾ
- Không còn quét toàn ảnh và vô tình lẫn chữ “LŨY KẾ” trong tiêu đề xanh lá phía trên.
- Tách riêng hàng tiêu đề cột màu xanh, phóng lớn trước khi OCR.
- Chấp nhận các biến thể OCR thường gặp như `LUYKE`, `LUYKEE`, `LUY + KE`, `LOY + KE`, và `TUYKE`.
- Xác định cột LŨY KẾ theo vị trí tương đối giữa MỤC TIÊU và TARGET; không dựa vào chỉ số cột cố định.
- Nếu không xác định được ô tiêu đề hoặc ranh giới cột an toàn, ứng dụng dừng và báo lỗi, không lấy nhầm TARGET.
- Bổ sung biên dòng đầu/cuối cho ảnh bảng trên điện thoại để giữ đủ 38 dòng của ảnh mẫu 09/11.

## OCR tỷ lệ 0%
- Chuẩn hóa các tỷ lệ OCR đọc được là `0`, `0%`, `0.0%` thành số 0.
- Nếu OCR vùng số màu bỏ sót nhưng OCR toàn thẻ đọc rõ token `0%`, tự động ghi nhận 0%.
- Không tự chuyển `None` hoặc ô không đọc được thành 0 nếu chưa có bằng chứng rõ ràng.

## Cài đặt GitHub / Streamlit Cloud
Đưa `app.py`, `requirements.txt`, `packages.txt` và `README.md` vào thư mục gốc repository. Main file: `app.py`.
