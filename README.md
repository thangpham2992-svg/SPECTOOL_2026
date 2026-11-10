# SpecWebTool ALPHA v2.6 — Excel Database mới (09/11/2026)

Webapp kỹ thuật thang máy Mitsubishi Electric, dùng hoàn toàn offline bằng `index.html` và có thể triển khai nguyên thư mục lên GitHub Pages.

## Mở ứng dụng

1. Giải nén ZIP đầy đủ. **Không** chỉ lấy riêng file `index.html`, vì app cần các file `.js`, `responsive-ui.css`, `database/` và `vendor/jszip.min.js`.
2. Mở `index.html` bằng Chrome / Microsoft Edge / Firefox. Internet và Python không cần thiết.
3. Chọn **NEXIEZ-MR ALPHA** hoặc **NEXIEZ-MRL Version02 ALPHA**.
4. Nhập Công trình, Hạng mục và các đặc tính. Ô nền **xanh nhạt** là có thể sửa/chọn, ô **xám** là khóa theo logic Excel.
5. Nhấn **Check Spec**. Chỉ sau khi kết quả PASS, mới dùng **Xuất Excel** hoặc **Xuất A4 PDF**.

## Giao diện responsive

- **Desktop/tablet**: nội dung / lựa chọn trên cùng một dòng; hai cụm thông số cạnh nhau khi đủ chiều rộng; chuyển nhóm I–V nhanh qua thanh điều hướng.
- **Điện thoại**: một cột các dòng, mỗi dòng có **nhãn bên trái ~43%** và **đầu vào bên phải ~57%**. Thanh thao tác Check Spec / Excel / PDF cố định ở đáy màn hình.
- Công trình và Hạng mục ở trên cùng. Không hiển thị mã ô Excel hoặc các thống kê nội bộ trên màn hình Sales.
- Giữ toàn bộ trường của SpecForm hiện tại; mã ô vẫn tồn tại trong dữ liệu DOM phục vụ Check Spec và Export.

## Export Excel

Nút **Xuất Excel (.xlsx)** dùng engine từ v2.1 để xuất `SpecFormExport` thành workbook một sheet `Specification`; loại các hàng Feature không áp dụng và xử lý Hall Indicators không cung cấp. Workbook xuất được thiết lập trang in A4 dọc.

## Export A4 PDF — đã sửa layout

Nút **Xuất A4 PDF** mở tab **Print-ready Specification**. Nhấn **In / Save as PDF** và trong hộp thoại của trình duyệt:

- Destination / Máy in: **Save as PDF**
- Paper size / Khổ giấy: **A4**
- Orientation / Chiều giấy: **Portrait / Dọc**
- Scale: **Default / 100%**; không áp dụng “Fit to page” lần thứ hai.
- Headers & footers / Đầu-chân trang trình duyệt: **Tắt**
- Background graphics / In đồ họa nền: **Bật** nếu muốn giữ màu các mục/Feature.

Bản in mới **không tạo bảng PDF từ dữ liệu text ba cột đơn giản**. Nó đọc lại chính workbook `.xlsx` `Specification` vừa được engine Export tạo ra, gồm:

- Tỷ lệ chiều rộng cột A/B/C trong file Excel;
- merged cells, gồm gộp theo cả hàng và cột;
- giá trị/công thức đã được tính lại và đóng băng;
- định dạng font, in đậm, nghiêng, nền, màu, đường viền của Excel;
- hàng bị ẩn bởi logic Feature/Hall Indicators;
- chiều cao hàng Excel làm cơ sở, có thể tự nới thêm nếu văn bản dài để tránh cắt chữ;
- ngắt trang A4 dọc, kiểm soát tiêu đề nhóm trước danh sách Feature.

**Giới hạn:** Đây là HTML/CSS print theo mô hình Excel, không phải bộ render Excel nguyên bản của Microsoft. Có thể có khác biệt nhỏ về font, ngắt trang, đường viền so với thao tác Print trực tiếp bằng Excel. In PDF vẫn cần người dùng chọn Save as PDF trong trình duyệt vì đây là Webapp HTML offline không có dịch vụ xử lý PDF phía server.

## Cập nhật Database Excel

Hai file gốc đặt trong `database/`:

- `SpecToolGXLM2P.xlsx` — NEXIEZ-MR ALPHA
- `SpecToolGQXL2P.xlsx` — NEXIEZ-MRL Version02 ALPHA

Offline: Mở **Cập nhật dữ liệu Excel**, nạp file `.xlsx` phù hợp với model. Trên GitHub Pages: thay hai workbook trong `database/` và tải lại trang. Khi Excel thêm kiểu công thức chưa được engine hỗ trợ, Check Spec phải báo không xác minh thay vì báo PASS.

**Bảo mật:** Nếu GitHub Pages public, các workbook và mã nguồn cũng truy cập được công khai. Sử dụng repository riêng, cấu hình quyền truy cập phù hợp nếu có dữ liệu nội bộ.

## Kiểm thử

- Cả MR & MRL: Check Spec PASS cấu hình mặc định, phát hiện FAIL ở các dữ liệu sai; DataTable MR mới có 31 Condition và MRL có 38 Condition.
- Bộ regression trên Database mới: **20/20** trường hợp đạt.
- Tạo và mở file `.xlsx` Excel cho hai model; kiểm tra chỉ một sheet Specification.
- Dựng/đọc bảng Excel trên tab in; MR có 134 hàng hiển thị, MRL 132 hàng ở dữ liệu mặc định.
- Bản A4 test được tạo bằng Chromium: **5 trang, A4 Portrait (594.96×841.92 pt)** cho mỗi model; không phát hiện text nằm ngoài trang.
- Mobile viewport 390px: không tràn ngang, thao tác Check Spec hoạt động.

## Lưu ý phiên bản

v2.3 thay đổi giao diện và engine trình bày PDF. Không thay đổi các công thức Check Spec, mapping, điều kiện DataTable hoặc engine Export Excel so với phiên bản v2.2. Kết quả PASS kiểm thử mặc định không thay thế phê duyệt kỹ thuật của Sales Engineering đối với các cấu hình đặc biệt.

## v2.6 — Những điểm cập nhật
- Database thay bằng **hai file mới** trong `database/` lấy từ `SpecToolZip(4).zip`.
- MRL D11 (Quantity) có các lựa chọn **1, 2, 3, 4** (file đã sửa theo yêu cầu).
- Khôi phục tham chiếu công thức dùng chung (Shared Formulas) trong workbook MR khi đọc dữ liệu nhúng và khi Sales nạp lại tệp Excel. Không dùng giá trị cache để phán đoán PASS.
- Check Spec MR theo workbook mới: `DataTable!AP16:AP27`, đặc biệt OH/PD dựa vào **B529/E521/E524/E523** cho Standard/Short OH/PD. Không dùng ngưỡng OH cũ với trần L400.
- Đối chiếu công thức DataTable: MR **678 công thức**, MRL **395 công thức**, không còn lỗi parser trong bộ đối chiếu mặc định.
- Kiểm thử 20/20 tình huống PASS/FAIL và 8 phép kiểm tra bổ sung theo Database mới. 747 công thức Export MR và 65 công thức Export MRL đều tính được ở cấu hình mặc định.
- Giữ giao diện responsive, dropdown, Check Spec nhảy đến lỗi, Excel Export và A4 PDF của v2.4.

**Phạm vi kiểm thử:** Kiểm tra công thức và xử lý bằng Node.js; chưa kiểm thử lại trực tiếp giao diện Chromium trên môi trường phiên này do trình duyệt tự động bị hạn chế truy cập. Những cấu hình đặc biệt vẫn phải được Sales Engineering kiểm tra.

## v2.3.1: Compact Check Result
- PASS: chỉ một dòng `PASS — Đạt / Passed`.
- FAIL: chỉ một dòng số lỗi và danh sách tên thông số bị lỗi; chạm/nhấn vào tên để đến đúng trường.
- INCOMPLETE: vẫn phân biệt với PASS để chặn xuất không an toàn.
- Các quy tắc DataTable, Check Spec và Export không thay đổi.

## v2.4 – Giao diện và dropdown
- Kết quả kiểm tra chỉ hiện **Đạt**, **Không đạt — số lỗi**, hoặc **Chưa xác minh**. Nếu lỗi, nhấn vào tên thông số để chuyển tới ô đó.
- Phục hồi dropdown từ Data Validation dạng công thức, Named Range, INDIRECT và OFFSET trong workbook đang nạp.
- Bổ sung dropdown cho khung cửa, vật liệu cửa tầng và vị trí đối trọng khi nguồn lựa chọn tồn tại trong DataTable.
- Tiêu đề tô xanh dương nhạt, viền dòng đồng đều trên desktop và mobile.
- Giữ nguyên công cụ Check Spec, xuất Excel và PDF của bản v2.3.


### v2.7 – Các ô tự động từ Excel
- NEXIEZ-MR: vị trí đối trọng theo `TblK5_Capacity[CWT-Position]`; khóa chỉnh sửa.
- NEXIEZ-MRL Version02: mẫu COP tự động theo công thức tại `SpecForm!F59`, phụ thuộc nút COP và mã COP; khóa chỉnh sửa.
- NEXIEZ-MR: mẫu COP tự động theo công thức tại `SpecForm!F63`, giữ khóa theo workbook.
- MRL Side drop BWH-620D/E và BWH-720F/G vẫn chọn được trong nhóm tải trọng hợp lệ.


## Cập nhật v3.1 – DSJ1 / Large Window và SUS304
- Có thể giữ SUS Standard cho vách và chi tiết cabin khi chọn cửa SUS-HL với DSJ1/Large Window.
- Cửa cabin và cửa tầng được chọn là SUS-HL + DSJ1/Large Window tự mô tả vật liệu SUS304.
- Chỉ sửa phạm vi SUS304 trong Check Spec; các điều kiện cửa, SDE, kích thước và tính năng khác giữ nguyên.
- Các file Excel Database đi kèm đã đồng bộ với công thức Webtool offline.

## Cập nhật v3.3 – Capacity và JJ
- NEXIEZ-MRL Version02: JJ là dropdown, không còn khóa. Đối với P08-CO 630 kg, dùng 800 hoặc 900 mm theo danh sách trong `DataTable!CA1080:CB1082`; những model khác căn cứ dòng Capacity trong DataTable và các lựa chọn bổ sung được khai báo sau này.
- NEXIEZ-MR P10 700 kg: Check Spec chỉ rõ lỗi khối lượng trang trí thay vì hiển thị nhầm Capacity. Giới hạn khối lượng vẫn lấy theo công thức Excel, không bỏ qua lỗi kỹ thuật.
- Danh sách cửa chống cháy EN81-58(II) 120 phút của hai model vẫn giữ nguyên.
- Xem `QA_v3.3.md` để biết phạm vi kiểm tra và giới hạn.

## v3.6: OH / PD mapping

GQXL2P: DataTable DA3:DL164 is the authoritative, normalized view of the unchanged BG8:BS43 OH/PD source. SpecForm F36/F37 and the Webtool calculate against DA:DL. Service floor is an optional free-text field. Run `node tools/qa_v36.js` to verify mappings. See `CHANGELOG_v3.6.md` and `QA_v3.6.md`.

## v3.9 — Blue headings and meaningful dimension validation

- Both Excel and PDF exports keep blue I/II/III/IV section headers; detailed Feature rows do not receive brown backgrounds.
- OH/PD/AH/BH errors show current value, actual source minimum/maximum and whether too small or too large; the link opens the correct field.
- The package bundles unchanged MR and MRL Excel Databases from v3.7. See `QA_v3.9.md` and `CHANGELOG_v3.9.md` for verification.

## v4.1 – ITP / Interphone type
When ITP is selected in Features → Signal, choose the corresponding system type under the ITP checkbox. The options are read from `DataTable!DN:DP` in the MR/MRL Excel Database and depend on Group Control Operation. For 1C-2BC with FE (including FE Hong Kong), `6V Single – 3 Master Stations` is mandatory. Unselected ITP type triggers `Vui lòng chọn loại ITP` in Check Spec and blocks Export. The selected type is exported in the Vietnamese and English ITP descriptions.

## Bản v4.2 — ITP dropdown & Ventilation

- ITP: trình chọn loại được ghim ở đầu mục V. Các tính năng (không cần mở tab Tín hiệu); chỉ xuất hiện khi bật tính năng ITP. Khi chưa chọn loại, Check Spec báo "Vui lòng chọn loại ITP".
- Excel Database: có Data Validation dạng danh sách tại SpecForm!H175 (MR) và SpecForm!H177 (MRL). Lựa chọn thay đổi theo hệ điều khiển và FE; danh sách nguồn là DataTable!DN:DP. Nên mở bằng Microsoft Excel Desktop để dùng Data Validation đầy đủ.
- Export: mô tả ITP theo giá trị đã chọn, bỏ câu ghi chú cũ "Trừ: Handrail & mirror nhập khẩu; sàn (By local)".
- MRL quạt thông gió: giữ Standard / 02 Sets. Theo G11L1P Sales Manual, Plasma Quad™ là một quạt tuần hoàn/tùy chọn riêng, chỉ áp dụng từ 825 kg và một số kích thước cabin không hỗ trợ; không tự gộp với quạt thông gió tiêu chuẩn.
- Để cập nhật GitHub Pages, thay toàn bộ file trong ZIP, đặc biệt index.html, app.js, các JS database và database/*.xlsx. Refresh trình duyệt sau khi deploy.
