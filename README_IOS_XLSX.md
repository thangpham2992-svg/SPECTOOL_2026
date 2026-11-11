# v5.8.3 – iOS XML BOM fix (based on v5.8.1)

## Nguyên nhân được xác định trong file nguồn
`xl/_rels/workbook.xml.rels` của cả hai Excel source có UTF-8 BOM ngay trước `<?xml ...?>`. Khi đọc chuỗi bằng JSZip, BOM trở thành ký tự U+FEFF. Một số XML parser trên iOS có thể báo `parsererror` nếu XML declaration không đứng ở đầu chuỗi.

## Sửa đổi
- Bỏ BOM U+FEFF khi đưa XML vào `DOMParser` trong export-engine.js và print-layout.js (cả app.js).
- Giữ nguyên layout XLSX/PDF, merge, styles, hidden rows/columns, logic Feature từ v5.8.1.
- PDF báo nguyên nhân thực tế trong tab thay vì trang “Preparing A4 layout…” trắng khi có lỗi.
- Giữ cách chia sẻ/tải file iOS v5.8.2.

## Kiểm thử cần làm trên iPhone
Check Spec = Đạt → Xuất XLSX → Share/Save to Files → mở Excel; sau đó xuất A4 PDF.

Chưa chạy thử trên thiết bị iOS thật.
