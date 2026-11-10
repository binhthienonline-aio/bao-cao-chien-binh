# BÁO CÁO CHIẾN BINH V5.1 — AUTO ZERO

## Nền tảng
Giữ nguyên V5 đã chốt: ảnh Lũy kế là nguồn doanh số gốc; 9 ảnh Dashboard là nguồn tỷ lệ phân bổ; không dùng file Excel tháng 10; không áp dụng hệ số 110%.

## Công thức
Doanh số nhân viên theo nhóm = LŨY KẾ cửa hàng của nhóm × tỷ lệ hoàn thành Dashboard của nhân viên / 100.

## Cập nhật OCR 0%
- Chuẩn hóa các tỷ lệ OCR đọc được là `0`, `0%`, `0.0%` thành số 0.
- Nếu OCR vùng số màu bỏ sót nhưng OCR toàn thẻ đọc rõ token `0%`, tự động ghi nhận 0%.
- Không tự chuyển `None` hoặc ô không đọc được thành 0. Nếu không có bằng chứng rõ ràng, vẫn giữ trạng thái cần rà soát.
- Danh sách “Chưa có số liệu” tiếp tục được dùng làm nguồn xác nhận 0 cho các nhóm tương ứng.

## Cài đặt GitHub / Streamlit Cloud
Đưa `app.py`, `requirements.txt`, `packages.txt` và `README.md` vào thư mục gốc repository. Main file: `app.py`.
