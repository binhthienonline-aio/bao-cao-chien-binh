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


## V5.4 — LŨY KẾ ONLY

- Chỉ nhận diện cột **LŨY KẾ** từ tiêu đề và đường biên cột; không yêu cầu OCR phải đọc được tiêu đề TARGET hoặc MỤC TIÊU/NV.
- Không đọc, không dùng TARGET trong công thức hoặc điều kiện tạo ma trận.
- Sửa lỗi lệch hàng: bắt đầu tại dòng dữ liệu đầu tiên, không đưa hàng tiêu đề vào danh sách 38 nhóm.
- OCR doanh số LŨY KẾ chạy 3 lượt độc lập; chỉ nhận số khi cả 3 lượt thống nhất. Nếu bất đồng, để trống và đánh dấu cần kiểm tra, không tự đoán.
- Công thức phân bổ: `LŨY KẾ cửa hàng × tỷ lệ Dashboard / 100`.


## V5.5 — LŨY KẾ OCR DECIMAL FIX

- Giữ nguyên chế độ LŨY KẾ ONLY; không đọc hoặc dùng TARGET.
- Chuẩn hóa số OCR theo định dạng hiển thị 2 chữ số thập phân của ảnh nguồn (ví dụ 7575 → 75.75; 659 → 6.59; 20 → 20.00).
- Tăng kích thước vùng số trước OCR; dùng 4 cách tiền xử lý ảnh khác nhau: grayscale, sharpen, Otsu và adaptive threshold.
- Chỉ tự xác nhận khi ít nhất 3 cách tiền xử lý cùng đọc ra một giá trị và không có giá trị khác có số phiếu bằng hoặc cao hơn. Nếu không đủ đồng thuận, để trống để kiểm tra, không đoán.
