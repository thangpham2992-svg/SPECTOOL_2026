# v3.7 — AECC cho SAI / AI-2200C (GQXL2P, G11L1P)

- Nguồn: Sales & Sales Engineering Manual for NEXIEZ-MRL Alpha Package Version2 [G11L1P], S-1-U4010-E, bản 2026-05-08, mục 6.4.2, bảng Signal features 1/3 (trang in 35).
- Manual: AECC Car Arrival Chime (Car), basic application, có giá trị `A` ở cột SAI-2200C. Không kết hợp AECC với AECH và AIL theo ghi chú của manual.
- Thay đổi chính xác `DataTable!J366` của `SpecToolGQXL2P.xlsx` từ `X` thành `A`.
- `DataTable!K366` vẫn dùng công thức tra cứu Group Control hiện có; 3C-AI2200C và 4C-AI2200C tra về cột J.
- `SpecForm!E165` vẫn dùng công thức VLOOKUP từ `Tbl39_Feature`; không thay đổi vị trí, tên bảng hay giá trị mặc định.
- Đồng bộ `engine-database.js` và `workbooks.js` cho chế độ offline.
- Không thay đổi MR, nhóm tính năng khác, OH/PD và Service floor của v3.6.
