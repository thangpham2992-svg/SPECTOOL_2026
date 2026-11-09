# Kiểm thử giao diện – v2.4

**Ngày kiểm thử:** 09/11/2026

- **Hai model:** NEXIEZ-MR ALPHA và NEXIEZ-MRL Version02 ALPHA.
- **Hiển thị:** desktop 1220 px và mobile 390 px; không có cuộn ngang ở 390 px.
- **Check Spec mặc định:** hai model đều hiển thị duy nhất **Đạt**; không hiển thị khối kết quả dài.
- **Trường hợp cố tình nhập OH thấp:** báo **Không đạt — 1 lỗi**, hiển thị tên thông số; nhấn chuyển đến đúng ô.
- **Dữ liệu dropdown:** kiểm tra MR trần D45 (14 mục), khung cửa D74 (3 mục), vật liệu E74 (2 mục); MRL chiều cao D30 (2 mục), SUS D40 (2 mục), HIP E75 (57 mục), COP dành cho người khuyết tật D62 (4 mục).
- **Điều kiện theo vật liệu:** SUS-HL khóa màu sơn; Painted mở lại màu sơn; chuyển DSJ1/SUS-HL thì danh sách SUS chuyển sang SUS304 và giữ giá trị cũ ở trạng thái không hợp lệ để Check Spec báo lỗi thay vì tự sửa ngầm.
- **Nạp lại Excel Database:** nạp trực tiếp hai workbook qua nút chọn file; Check Spec đều đạt với cấu hình gốc.
- **Export:** thử xuất file Excel hợp lệ, một sheet `Specification`, và mở preview PDF A4 Portrait của cả hai model.
- **JavaScript:** không có lỗi khi chạy các ca thử trên Chromium.

Vẫn cần rà soát sản phẩm theo quy trình phê duyệt kỹ thuật trước khi dùng làm tài liệu chính thức; kiểm thử giao diện không phải chứng nhận mọi tổ hợp kỹ thuật.
