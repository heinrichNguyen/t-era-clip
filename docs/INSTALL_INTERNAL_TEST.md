# T-eraClip: cài thử nội bộ

## Bản thử này làm gì?

Thu thập và xuất thảo luận 1Office trên `https://office.meygroup.vn/`: trang hiện tại, toàn bộ công việc trong dự án hoặc danh sách công việc đang lọc. Có Capture để lưu nội dung được chọn trên trang web vào máy.

Đây là bản thử nội bộ, chưa được Google xét duyệt. Không có AI, kết nối ChatGPT, cloud sync hoặc tải file đính kèm gốc. Không cần Node.js, tài khoản GitHub hoặc cài công cụ lập trình. Cần Google Chrome trên Windows hoặc macOS. Chưa kiểm thử trực tiếp trên máy Windows của đồng nghiệp.

## Cài đặt từng bước

1. Tải file `t-eraclip-test-noi-bo.zip` do người phụ trách gửi.
2. Giải nén ZIP. Bên trong có thư mục `T-eraClip` và file hướng dẫn này.
3. Đặt thư mục `T-eraClip` ở nơi cố định trên máy, ví dụ Documents. Không xóa hoặc di chuyển thư mục sau khi cài. Không mở các file HTML để cài.
4. Mở Chrome, gõ `chrome://extensions` vào thanh địa chỉ rồi nhấn Enter.
5. Bật `Chế độ dành cho nhà phát triển` (Developer mode), thường ở góc trên bên phải.
6. Bấm `Tải tiện ích đã giải nén` (Load unpacked).
7. Chọn đúng thư mục `T-eraClip` chứa `manifest.json`, không chọn ZIP hoặc thư mục ngoài cùng.
8. Kiểm tra thẻ T-eraClip xuất hiện và đang bật. Bấm biểu tượng Tiện ích ở thanh công cụ, ghim T-eraClip để dễ sử dụng.
9. Mở `https://office.meygroup.vn/`, đăng nhập bằng tài khoản 1Office của chính mình, rồi mở dự án/công việc được phép xem.

Nếu Developer mode hoặc Load unpacked bị chặn, không tìm cách vượt chính sách. Gửi ảnh lỗi cho người hỗ trợ. Không kéo thả file CRX để thay cho quy trình này.

## Thử thu thập và xuất Excel

1. Chọn một công việc nhỏ đã biết rõ số bình luận/phản hồi.
2. Mở T-eraClip, bấm `Thu thập thảo luận trang hiện tại`, kiểm tra số lượng và mở trang kết quả bằng nút tương ứng.
3. Xuất Excel, mở bằng Excel và đối chiếu No., mã/tên công việc, tác giả, thời gian, nội dung và quan hệ phản hồi.
4. Mở dự án nhỏ, vào tab `Công việc`, bấm `Kiểm tra danh sách công việc dự án`. Đối chiếu số công việc trước khi thu thập toàn bộ.
5. Bấm `Thu thập toàn bộ thảo luận dự án`, giữ tab nguồn và trang tiến độ mở. Không thao tác chuyển sang dự án khác trong lúc thu thập.
6. Thử một widget của dự án để mở danh sách đang lọc, kiểm tra extension nhận đủ tất cả các trang rồi thu thập/xuất danh sách đó. Không tiếp tục coi là đủ nếu có cảnh báo thiếu công việc.

Extension chỉ sử dụng quyền của tài khoản đang đăng nhập; không cấp thêm quyền xem dự án. Số lượng và dữ liệu có thể khác tài khoản admin. Nếu gặp lỗi quyền, dừng và gửi thông tin lỗi, không dùng chung mật khẩu admin.

Nên có một tài khoản quản lý đầy đủ và một tài khoản có phạm vi khác trong nhóm thử. Kiểm tra cả công việc trống, bình luận có phản hồi, dự án nhiều cấp và danh sách nhiều trang. Chưa coi việc chạy thành công trên một tài khoản là xác nhận chạy đúng cho toàn phòng ban.

## Thử Capture

1. Chọn một đoạn văn bản không nhạy cảm trên trang web thông thường.
2. Nhấp chuột phải, chọn mục lưu vào T-eraClip.
3. Mở T-eraClip, bấm `Capture`, kiểm tra nội dung đã lưu; thử sửa tiêu đề/ghi chú, lọc và quay lại menu chính.
4. Không thử `Reset danh sách` với dữ liệu quan trọng chưa sao lưu. Reset xóa toàn bộ mục Capture, kể cả mục đang bị bộ lọc che, không phải dữ liệu 1Office gốc.

## Dữ liệu và cập nhật

- Capture và thiết lập nằm trong Chrome profile hiện tại. Kết quả thu thập 1Office dùng bộ nhớ trang và/hoặc Chrome session storage, không đồng bộ giữa các máy.
- File Excel/JSON đã tải là file riêng trên máy, có thể chứa thông tin cá nhân và dữ liệu công ty. Chỉ gửi qua kênh nội bộ được phép.
- Extension không yêu cầu mật khẩu 1Office/Google. Đăng nhập 1Office trực tiếp trên trang của hệ thống.
- Khi cập nhật bản unpacked: đóng phiên thu thập, sao lưu Capture từ Options, giữ nguyên thư mục cài đặt, thay bằng các file bản mới rồi vào `chrome://extensions` bấm Reload trên thẻ T-eraClip. Không gỡ extension để cập nhật nếu muốn giữ dữ liệu.
- Khi chuyển sang bản Chrome Web Store: coi như một bản cài khác. Sao lưu Capture từ Options trước; sau khi cài Store có thể nhập backup qua Options. Không bảo đảm dữ liệu hoặc ID extension chuyển tự động. Không chạy hai bản cùng lúc; không gỡ bản thử trước khi xác nhận backup và bản Store.
- Backup Capture không phục hồi phiên thu thập 1Office. Xuất Excel/JSON cần giữ trước khi đóng/reset trang.

## Báo lỗi

Liên hệ `tuan.nguyen03@tanadaithanh.vn`.

Gửi: phiên bản Chrome/OS, bản T-eraClip và BUILD-INFO do người phụ trách cung cấp, thao tác đã làm, số công việc/bình luận thực tế và số thu được, thông báo lỗi, thời gian xử lý. Ảnh chụp cần che tên dự án/nội dung nhạy cảm khi không cần cho chẩn đoán. Không gửi mật khẩu, cookie, token hoặc toàn bộ dữ liệu công ty qua kênh không được phép.

## Checklist nghiệm thu trước khi mở rộng

- [ ] Cả 1–2 người cài được và dùng tài khoản riêng.
- [ ] Thu thập trang hiện tại đúng số bình luận/phản hồi, nhận đúng công việc.
- [ ] Dự án nhiều cấp giữ đúng No. và cột công việc cha/cấp 3.
- [ ] Danh sách lọc nhiều trang có số công việc ổn định giữa hai lần chạy.
- [ ] Excel mở được, đối chiếu nguồn và không thiếu trường đã biết.
- [ ] Công việc không có thảo luận được phân biệt với lỗi thu thập.
- [ ] Capture, quay lại menu, backup và reset đúng phạm vi.
- [ ] Ghi nhận thời gian chạy và lỗi, không dùng riêng kết quả admin để kết luận.

Không có yêu cầu cấp quyền quản trị macOS/Windows. Chrome vẫn có thể thay đổi chính sách hoặc hạn chế bản unpacked; bản thử này không phải phương án phát hành lâu dài.
