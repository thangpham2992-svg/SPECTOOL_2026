# v3.3 — Kiểm tra Capacity và JJ

- MRL: JJ là dropdown, không khóa; P08-CO (630 kg) cho phép 800/900 mm theo yêu cầu đã xác nhận. Với các Capacity khác, danh sách được lấy từ cột JJ trong DataTable hiện hành; không tự đoán thêm lựa chọn.
- Check Spec có ràng buộc đối chiếu JJ với Capacity; lựa chọn ngoài danh sách bị báo lỗi.
- MR: lỗi khối lượng trang trí do AP5 được gắn về trường khối lượng trang trí thay vì Capacity; P10 với Cabin kích thước chuẩn và khối lượng trang trí 25 kg vẫn có thể báo FAIL hợp lệ vì vượt ngưỡng, 10 kg thì PASS trong ca thử.
- Khi thay Capacity, giá trị AH/BH gợi ý theo trạng thái cửa chống cháy từ các cột DataTable tương ứng.
- Cửa chống cháy của cả hai model giữ EN81-58(II) – 120 phút từ Database v3.2.
- Kiểm thử hồi quy và kiểm tra dữ liệu được mô tả trong QA_v3.3.md. Không coi việc kiểm thử hữu hạn là bảo đảm mọi tổ hợp kỹ thuật.
