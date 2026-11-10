# Chính sách quyền riêng tư T-eraClip (bản dự thảo)

Trạng thái: chuẩn bị cho bản phát hành Private, chưa công bố. Nội dung được đối chiếu với mã hiện tại; người phụ trách phải duyệt trước khi đăng lên URL HTTPS công khai có thể truy cập bởi Google và người dùng.

## Phạm vi

T-eraClip giúp người dùng ghi nhận và xuất thông tin công việc, gồm thảo luận 1Office trên `office.meygroup.vn` và nội dung được người dùng chủ động chọn để lưu bằng Capture. Extension không phải dịch vụ 1Office, không cấp thêm quyền truy cập và không yêu cầu mật khẩu hệ thống.

## Dữ liệu được xử lý

- Thông tin dự án/công việc: tên, ID, mã, No., quan hệ cấp, trạng thái, thông tin người giao/thực hiện/phòng ban và các mốc thời gian nếu hệ thống cung cấp.
- Thảo luận: nội dung, tên tác giả, thời gian, quan hệ phản hồi, người được nhắc đến và metadata tệp đính kèm, gồm tên, dung lượng hoặc URL ảnh xem trước nếu có.
- Capture: văn bản đã chọn, URL/tiêu đề trang nguồn, thời gian lưu, thông tin tổ chức nội dung do người dùng nhập.
- Báo cáo lịch sử truy cập (khi người dùng chủ động yêu cầu): ID bản ghi, tài khoản, họ tên và thời gian truy cập từ nhật ký 1Office. Không thu IP hoặc thông tin trình duyệt vào báo cáo này. Đọc toàn bộ nhật ký được cấp quyền xem hoặc chỉ nhật ký các tài khoản khớp bộ lọc nhân sự người dùng chọn trong extension, không giới hạn ở các bản ghi đang hiển thị trên trang nguồn. Dữ liệu chỉ giữ trong bộ nhớ trang báo cáo đến khi đóng hoặc tải lại; file Excel đã xuất do người dùng quản lý.
- Báo cáo này bổ sung vị trí, phòng ban và công ty hiện tại từ danh sách người dùng đang hoạt động/đã khóa, nguồn chọn người dùng và cây cơ cấu 1Office. Khi lọc trước khi quét, đọc danh sách nhân sự trước để chọn tài khoản, rồi đọc nhật ký của các tài khoản khớp. Chỉ dùng ID người dùng, tài khoản và ID phòng ban để ghép, không ghép theo họ tên, không tái dựng lịch sử nhân sự. Không lưu avatar, giới tính hoặc dữ liệu khác của nguồn chọn người dùng. Thiếu quyền hoặc không ghép được thì để trống và cảnh báo; không tự suy đoán tài khoản thuộc công ty để đưa vào bộ lọc.
- Báo cáo truy cập chỉ xuất tài khoản được xác minh hiện còn active. Danh sách khóa dùng để đối chiếu trạng thái và phát hiện trùng/mâu thuẫn, không xuất nhật ký tài khoản khóa. Không xác minh được danh sách trạng thái thì chặn xuất; trạng thái hiện tại không được coi là trạng thái tại ngày chốt quá khứ.
- Thiết lập và dữ liệu tạm cần cho quá trình thu thập/xem trước.

Dữ liệu nguồn có thể chứa thông tin cá nhân hoặc nội dung nhạy cảm tùy những gì người dùng chọn. Extension không duyệt toàn bộ lịch sử trình duyệt; chỉ xử lý trang/phạm vi người dùng chủ động thao tác.

## Mục đích sử dụng

Chỉ để thực hiện chức năng ghi nhận nội dung, quản lý Capture cục bộ, thu thập/xem trước và xuất Excel/JSON theo yêu cầu người dùng. Không có quảng cáo, analytics, mua bán dữ liệu hoặc phân tích AI trong bản này.

## Lưu trữ và kết nối

Capture và thiết lập được lưu bằng `chrome.storage.local` trên Chrome profile hiện tại. Dữ liệu xem trước/phiên thu thập dùng bộ nhớ trang và/hoặc `chrome.storage.session`. Chrome session storage không phải cloud sync. Không có cơ sở dữ liệu hoặc backend của T-eraClip để nhận nội dung này.

Để thu thập, extension đọc DOM và gọi các endpoint 1Office trong origin được hỗ trợ, sử dụng phiên đăng nhập hiện có của người dùng (`credentials: include`). Trình duyệt gửi thông tin phiên cho 1Office theo cơ chế của hệ thống; T-eraClip không yêu cầu, đọc hoặc lưu mật khẩu/cookie đăng nhập riêng. Extension có thể mở công việc, chuyển trang danh sách và bấm tải thêm bình luận để đọc đủ dữ liệu. Không chủ động sửa công việc hoặc đăng thảo luận.

Không truyền dữ liệu thu thập đến OpenAI, nhà phát triển T-eraClip hoặc bên quảng cáo. Nếu người dùng tự gửi file/ảnh lỗi cho bộ phận hỗ trợ, việc chia sẻ đó là thao tác riêng của người dùng và cần tuân thủ quy định công ty.

## Xuất và xóa dữ liệu

Excel/JSON tải xuống do người dùng quản lý, có thể chứa thông tin công ty và thông tin cá nhân. Các chức năng xóa trong extension không xóa file đã tải hoặc dữ liệu gốc trên 1Office.

Người dùng có thể reset danh sách Capture hoặc xóa dữ liệu Capture/thiết lập tại Options. Các thao tác này không cam kết xóa dữ liệu phiên 1Office đang mở; đóng trang kết quả và kết thúc phiên trình duyệt để kết thúc dữ liệu tạm tương ứng. Sao lưu trước khi gỡ extension hoặc thay Chrome profile. Không đồng bộ dữ liệu giữa thiết bị.

## Quyền Chrome

- `storage`: lưu Capture, thiết lập và phiên thu thập/xem trước.
- `contextMenus`: cung cấp thao tác lưu văn bản được chọn khi nhấp chuột phải.
- `activeTab`: truy cập tạm trang người dùng chủ động mở extension/thao tác để lấy nội dung, không có quyền truy cập thường trực toàn bộ website.
- `scripting`: chạy bộ thu thập đóng gói sẵn trên tab 1Office được hỗ trợ, đọc dữ liệu và tải thêm các mục trong phạm vi đã chọn.

Không có quyền `cookies`, persistent host permissions hoặc mã JavaScript được tải từ máy chủ bên ngoài để thực thi trong extension. Toàn bộ mã chức năng extension được đóng gói cùng bản cài.

## Giới hạn bảo mật

Dữ liệu nằm trong profile và file trên máy không được T-eraClip bổ sung lớp mã hóa độc lập. Người dùng phải bảo vệ máy/profile và file xuất theo quy định công ty. Kiểm thử chức năng không phải chứng nhận an toàn bảo mật hoặc bảo đảm không có lỗi. Quyền xem dữ liệu thực tế do 1Office và tài khoản đang đăng nhập quyết định.

## Liên hệ và thay đổi

Email hỗ trợ: `tuan.nguyen03@tanadaithanh.vn`.

Đơn vị/người chịu trách nhiệm phát hành và hiệu lực của chính sách: [Cần xác nhận trước khi công bố]. Không coi tên/logo trong giao diện là bằng chứng T-eraClip đã được công bố là sản phẩm chính thức của tập đoàn.

Nếu sau này thêm AI, cloud hoặc bên xử lý dữ liệu mới, cần cập nhật chính sách và xin sự đồng ý phù hợp trước khi gửi dữ liệu; bản này chưa có các chức năng đó.
