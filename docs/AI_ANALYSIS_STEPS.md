# Định hướng AI của T-eraClip

## Quyết định hiện tại

Đã dừng và gỡ bản thử AI thủ công theo quyết định của người dùng.

Luồng thử gồm chuẩn bị prompt, tải gói JSON, chuyển dữ liệu sang ChatGPT, dán JSON kết quả trở lại, kiểm tra nguồn và xuất thêm sheet `Phân tích AI`. Luồng này tạo nhiều thao tác hơn việc đưa Excel gốc vào ChatGPT, không đáp ứng mục tiêu tiện và nhanh.

Đã gỡ khối chuẩn bị yêu cầu, giao diện chuyển dữ liệu/nhận kết quả, mã kiểm tra kết quả và nhánh xuất sheet AI thử nghiệm. Không giữ các nút AI chưa có kết nối thật.

## Chức năng giữ nguyên

- Thu thập thảo luận trang hiện tại, toàn dự án và danh sách công việc đang lọc.
- Xuất Excel dữ liệu gốc với các sheet `Dự án`, `Công việc`, `Thảo luận`, `Tệp đính kèm` đối với kết quả dự án/danh sách.
- Các cột thông tin công việc, No., cấu trúc phản hồi và dữ liệu nguồn đã triển khai trước bản thử AI.
- Chức năng xuất JSON thảo luận vốn có, không phải gói JSON dành cho bản thử AI.
- Capture và các thiết lập hiện có.

Nếu cần phân tích ngoài T-eraClip, người dùng tự đưa Excel đã xuất vào ChatGPT và nêu yêu cầu, chỉ khi được công ty cho phép chia sẻ dữ liệu. Extension không tự upload, không nhận kết quả AI trở lại và không tự tạo sheet AI.

## Điều kiện để xem xét lại

Trải nghiệm mục tiêu: `Nhập yêu cầu -> Xem kết quả -> Tải Excel`, ngay trong T-eraClip, không chuyển file hoặc sao chép kết quả qua lại bằng tay.

Đây là hướng tương lai, chưa được triển khai. Trước khi làm tiếp cần xác minh:

1. Cách kết nối AI thật phù hợp với ứng dụng nội bộ và điều kiện dịch vụ. Đăng nhập tài khoản không tự đồng nghĩa với quyền dùng gói ChatGPT để phân tích từ extension.
2. Quyền chia sẻ dữ liệu công ty, phạm vi dữ liệu gửi đi và sự đồng ý của người dùng.
3. Chi phí, giới hạn sử dụng, nơi lưu thông tin xác thực và khả năng chỉ cài extension.
4. Cách xử lý dự án lớn, dữ liệu chưa đầy đủ, kết quả lỗi và tham chiếu nguồn.

Không tự công khai mã nguồn, thêm backend/chương trình phụ hoặc đưa API key vào extension. Không mở rộng sang Claude, nhập file tùy ý hoặc sửa dữ liệu 1Office bằng AI khi chưa có yêu cầu riêng.

## Nguyên tắc kết quả tương lai

- Giữ nguyên dữ liệu 1Office gốc, phân biệt rõ với kết quả AI.
- Nội dung thảo luận là dữ liệu, không phải chỉ dẫn điều khiển hệ thống.
- Không bịa căn cứ, người phụ trách hoặc thời hạn; ghi rõ khi thiếu thông tin.
- Tham chiếu nguồn và kiểm tra định dạng không thay thế việc xác minh kết luận.
- Không tuyên bố đã xử lý toàn bộ khi chỉ phân tích một phần dữ liệu.
