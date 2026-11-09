# Kiểm thử SpecWebTool v2.6 — Database mới

Ngày: 09/11/2026

## Database sử dụng

- `database/SpecToolGXLM2P.xlsx`: nguồn MR trong `SpecToolZip(4).zip` (không chỉnh sửa).
- `database/SpecToolGQXL2P.xlsx`: nguồn MRL trong `SpecToolZip(4).zip`, đã sửa D11 dropdown từ 1,2,3 thành 1,2,3,4.
- Cả hai giữ đủ các sheet: SpecForm, SpecFormExport, DataTable.

## Sửa lỗi tương thích MR mới

- 93 công thức dùng chung của MR và 48 công thức dùng chung của MRL được chuyển thành tham chiếu có thể tính lại trong engine nhúng.
- `app.js` hỗ trợ giải mã công thức dùng chung khi người dùng nạp Excel mới bằng giao diện, không chỉ khi dùng dữ liệu nhúng.
- Check Spec MR hiện đọc các nhánh thông báo AP16:AP27; kiểm tra OH và PIT từ bảng Standard/Short OH/PD của workbook hiện hành.

## Kiểm tra tự động đã thực hiện

- `node tools/regression.js`: **20/20 PASS** (MR/MRL, OH, PD, SDE, SUS304, màu sơn, Feature...).
- `node tools/audit_all.js`: DataTable MR **678/678** công thức tính được; MRL **395/395**, không có lỗi/mismatch với các giá trị cache có nghĩa.
- Kiểm tra bổ sung: 8/8 trường hợp PASS/FAIL theo DataTable mới, gồm MR Short OH = 4300 mm, OH 4200 mm phải FAIL.
- Export formula audit: MR **747/747**; MRL **65/65**, không lỗi trong cấu hình mặc định.
- D11 MRL: 1 / 2 / 3 / 4 được nhận diện từ Excel Database mới.
- JavaScript: tất cả file kiểm tra cú pháp thành công.

## Chưa kiểm chứng trong phiên này

- Test thao tác nạp file, tải Excel và in A4 PDF trực tiếp trên Chromium bị chặn quyền truy cập URL trong môi trường kiểm thử. Mã giao diện/Export giữ từ bản v2.4 đã được thử ở các phiên trước; **không đồng nghĩa kết quả xuất bằng workbook mới đã được thử end-to-end**.
- Chưa thể xác nhận tương đương Tool.exe 100% cho mọi cấu hình đặc biệt.
