# Cài T-eraClip Thảo Luận Trên Windows / macOS

## Dành Cho Người Dùng

Không cần Node.js, npm, Git, GitHub hay công cụ lập trình. Các lệnh trong phần Developer Setup của README chỉ dành cho người phát triển.

Gói này chỉ thu thập và xuất thảo luận 1Office của công việc đơn, dự án hoặc danh sách trong báo cáo đã lọc. Không có Capture hoặc chức năng sửa mã/No. Cần Chrome và tài khoản 1Office có quyền xem dữ liệu tương ứng. Bản pilot chưa được Google xét duyệt; chưa kiểm thử trực tiếp trên máy Windows của key user.

## Cài Lần Đầu

1. Nhận `T-eraClip-Thao-luan-Noi-bo.zip` từ người phụ trách qua kênh nội bộ được phép.
2. Windows: nhấp chuột phải ZIP, chọn **Extract All / Giải nén tất cả**. macOS: mở ZIP để giải nén.
3. Đặt thư mục đã giải nén ở nơi cố định, ví dụ Documents. Bên trong có thư mục `T-eraClip-Thao-luan` và hướng dẫn này.
4. Mở Chrome, nhập `chrome://extensions` vào thanh địa chỉ, nhấn Enter.
5. Bật **Developer mode / Chế độ dành cho nhà phát triển**, giữ chế độ này bật khi dùng bản unpacked.
6. Bấm **Load unpacked / Tải tiện ích đã giải nén**.
7. Chọn thư mục `T-eraClip-Thao-luan`. File `manifest.json` phải nằm ngay trong thư mục đó. Không chọn ZIP hoặc thư mục ngoài cùng.
8. Kiểm tra T-eraClip xuất hiện và đang bật. Ghim tiện ích từ biểu tượng Tiện ích trên thanh công cụ.
9. Đăng nhập 1Office bằng tài khoản của mình. Mở công việc, dự án hoặc báo cáo đã lọc rồi bấm tiện ích để thu thập.

Đăng nhập trên trang 1Office, không gửi mật khẩu cho người hỗ trợ. Bắt đầu bằng một công việc nhỏ, đối chiếu bình luận/phản hồi trước khi thu thập hàng loạt. Giữ tab nguồn và trang tiến độ mở; khi lấy theo báo cáo, không đổi bộ lọc giữa chừng. Kiểm tra cảnh báo trước khi xuất Excel.

## Khi Có Bản Mới

1. Xuất các kết quả đang cần giữ, chờ lượt thu thập kết thúc và đóng trang tiến độ.
2. Nhận bản mới, thay toàn bộ nội dung trong thư mục cài hiện tại, không đổi vị trí thư mục.
3. Mở `chrome://extensions`, bấm **Reload / Tải lại** trên thẻ T-eraClip.
4. Mở lại tiện ích từ tab 1Office. Không cần gỡ rồi cài lại.

Bản unpacked không tự cập nhật qua GitHub. Không xóa hoặc di chuyển thư mục sau khi nạp. Kết quả Excel đã tải là file riêng, có thể chứa dữ liệu công ty và cần quản lý theo quy định nội bộ.

## Nếu Không Cài Được

- Không thấy `manifest.json`: có thể đã chọn nhầm thư mục hoặc tải mã nguồn thay vì bản đã build. **Code > Download ZIP trên GitHub không phải bản cài.**
- Developer mode hoặc Load unpacked bị chặn: liên hệ CNTT, không vượt chính sách quản lý máy.
- Có lỗi sau khi thay bản mới: tải lại tiện ích, mở lại từ tab 1Office và gửi nguyên thông báo lỗi cho người phụ trách.

Gói unpacked có thể bị sao chép, không giới hạn người cài theo danh sách key user. Chỉ gửi cho người được chỉ định là quy trình phân phối, không phải cơ chế kiểm soát truy cập. Tool không cấp thêm quyền dữ liệu 1Office. Phân phối lâu dài có thể cần Chrome Web Store hoặc chính sách doanh nghiệp được phê duyệt.

Tham khảo quy trình chính thức: https://developer.chrome.com/docs/extensions/get-started/tutorial/hello-world#load-unpacked
