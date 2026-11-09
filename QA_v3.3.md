# Kiểm thử kỹ thuật – SpecWebTool v3.3

## Những việc đã xác minh
- Đối chiếu 64 dòng Capacity trong hai bảng TblK5_Capacity; mỗi dòng được kiểm tra theo giá trị đầu vào dựng từ DataTable. Có 29 trường hợp đạt ngay trên cấu hình mẫu và 35 trường hợp báo lỗi khi các lựa chọn liên quan chưa tương thích. **35 FAIL này không đồng nghĩa 35 Capacity không hợp lệ**; mỗi cấu hình còn phụ thuộc OH/PD, nhóm điều khiển, COP, Feature và vật liệu.
- Bộ kiểm thử hồi quy Check Spec: 21/21 đạt mong đợi. Thêm 6 trường hợp nhắm tới MR P10 (khối lượng trang trí) và MRL JJ=800/900/bất hợp lệ đạt mong đợi.
- Chạy trình duyệt Chromium: MRL P08-CO 630 kg có thể chọn JJ 800 hoặc 900; chọn 800, Check Spec = Đạt, Export Excel chứa 800 mm.
- MR P10-CO 700 kg với các kích thước từ DataTable và khối lượng trang trí nhập 25 kg báo lỗi ở mục Khối lượng trang trí. Khi giảm về 10 kg, Check Spec = Đạt trong cấu hình thử nghiệm.
- Nạp Excel lại ngay trên trình duyệt thành công đối với cả MR và MRL; Check Spec = Đạt với cấu hình mặc định, Export Excel tạo được file hợp lệ.
- Source workbook MRL chứa bảng bổ sung DataTable!CA1080:CB1082: lựa chọn JJ 800/900 mm của P08-CO (630 kg). Khi muốn thêm phương án JJ cho Capacity khác, hãy thêm vào các dòng CA1083:CB1100, đúng mã Capacity và giá trị JJ.
- MRL D11 còn 1/2/3/4. Cửa chống cháy cả hai model: EN81-58(II), 120 phút như dữ liệu v3.2.

## Giới hạn cần lưu ý
- Đối với Capacity khác P08-CO 630 kg, dropdown JJ hiện lấy giá trị mặc định từ cột JJ của TblK5_Capacity, cộng các lựa chọn bổ sung nếu được khai báo tại DataTable!CA1081:CB1100. Không suy đoán các kích thước khác khi chưa có dữ liệu.
- Kiểm thử các trường hợp kỹ thuật đã nêu chưa phải kiểm định đầy đủ mọi tổ hợp cấu hình. Các yêu cầu đặc biệt vẫn cần Sales Engineering kiểm tra.
- Điều kiện Check Spec được tính lại theo DataTable; các điều kiện chưa tính được phải chặn Export.
