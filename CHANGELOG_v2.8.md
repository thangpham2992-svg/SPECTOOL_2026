# v2.8 — Đồng bộ mô tả sàn MRL với MR

- DataTable của NEXIEZ-MRL Version02 ALPHA: cập nhật `TblK28_Flooring[FlooringVN]` (B159) và `FlooringEN` (C159) theo đúng cách mô tả của NEXIEZ-MR ALPHA.
- Bản `engine-database.js` nhúng được cập nhật cùng nội dung; khi chạy offline hoặc chưa nạp lại file vẫn dùng bảng mô tả mới.
- Giữ nguyên hai loại sàn, danh sách dropdown, công thức khối lượng sàn tại D56 và công thức `SpecFormExport!C56`.
- Khi xuất Spec với sàn đá địa phương, câu mô tả theo mẫu MR: “Sàn đá granite cung cấp trong nước, chiều dày hoàn thiện sau thi công 20mm, khối lượng ... Kg (Local natural granite, finishing thickness 20 mm, weight ... Kg).”
- Không thay đổi các logic Check Spec, COP và CWT.

Đã kiểm tra cấu trúc ZIP/XLSX, 40 bảng của MRL vẫn còn đầy đủ và cú pháp JavaScript hợp lệ. Chưa thể chạy thử trình duyệt trong môi trường này do Chromium bị chặn bởi chính sách hệ thống.
