# BÁO CÁO CHIẾN BINH V5 — IMAGE ONLY

## Nguyên tắc dữ liệu
- Không đọc file Excel tháng 10.
- Ảnh Lũy kế: đọc động tên nhóm hàng và cột TARGET theo từng hàng của bảng.
- 9 ảnh Dashboard: phát hiện từng thẻ trắng trong vùng THI ĐUA, OCR nhãn và tỷ lệ.
- Danh sách "Chưa có số liệu" được đọc riêng; chỉ các nhóm ghép được từ danh sách này mới tự gán 0.
- Nhóm không ghép chắc chắn để trống; không tự suy đoán.
- Tạo ma trận nhân viên × nhóm; doanh số = TARGET siêu thị × hệ số / 9 × tỷ lệ / 100.
- Xuất Excel theo mẫu vàng/cam/xanh và PNG tương ứng.

## Cập nhật Streamlit
Upload/ghi đè `app.py`, `requirements.txt`, `packages.txt` vào repository GitHub đang kết nối Streamlit Cloud.
Main file: `app.py`.

## File đầu vào
Một ZIP gồm 9 PNG dashboard, tên file có mã nhân viên (ví dụ `suc-khoe-nv-137501-202611.png`) và 1 ảnh Lũy kế (tên file có `Lũy kế` hoặc `Luy ke`). Có thể tải ZIP trực tiếp hoặc dán link Drive công khai.

## Lưu ý
OCR không phải nguồn dữ liệu tuyệt đối. Các ô chưa chắc chắn để trống, có thể sửa trực tiếp ở bước rà soát trước khi dùng báo cáo chính thức.


V5 NO110: Đã bỏ hoàn toàn hệ số 110%. Mục tiêu cá nhân = TARGET siêu thị / số nhân viên. Doanh số quy đổi = mục tiêu cá nhân × tỷ lệ Dashboard / 100.
