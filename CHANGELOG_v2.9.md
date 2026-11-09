# Thay đổi logic cabin quan sát NY-40

- Chọn NY-40 dựa trên Capacity, Door Type và Entrance Type theo `SpecForm!R13/R14` và Data Validation tại D14.
- KHÔNG khóa NY-40 vì trần FL1 hoặc tay vịn. Khi đã chọn NY-40, Check Spec đánh dấu lỗi đúng tại dòng **Trần cabin** và **Tay vịn**.
- Excel Database MRL giữ nguyên công thức `DataTable!D525` và Data Validation D14; chỉ cập nhật công thức diễn giải lỗi tại `DataTable!C525` để hỗ trợ tra cứu.
- Popup dùng hình NY-40A/B/C/D/E chính thức từ ảnh người dùng cung cấp, các mẫu không hợp lệ theo điều kiện đầu vào sẽ bị làm mờ.
- Mã lỗi chi tiết phản ánh đúng trường cần sửa, không thay đổi tiêu chuẩn áp dụng kỹ thuật.
