# T-eraClip: phát hành thử và Chrome Web Store Private

## Trạng thái

Đã chuẩn bị bản ZIP cài thử và ZIP theo cấu trúc Chrome Web Store. Chưa upload, gửi xét duyệt, thanh toán hoặc được Google chấp thuận. Mã ứng dụng giữ chức năng hiện tại; chỉ bổ sung metadata/icon cho phát hành. Bản build dựa trên working tree hiện tại, kể cả thay đổi chưa commit.

Tài khoản nhà phát hành do người dùng cung cấp: `nguyenanht1000@gmail.com`, chưa đăng ký developer. Email hỗ trợ: `tuan.nguyen03@tanadaithanh.vn`. Không thu thập mật khẩu/tài khoản admin trong bộ hồ sơ.

Không cần máy công ty được quản lý để dùng bản Store Private. Giới hạn cài đặt theo tài khoản Google hoặc Google Group, không chỉ theo đuôi email. Không chọn Unlisted vì người có link đều có thể cài.

## Hai gói khác nhau

- `t-eraclip-test-noi-bo.zip`: gửi riêng cho 1–2 người thử. Giải nén, Load unpacked thư mục `T-eraClip`; hướng dẫn nằm cùng ZIP.
- `chrome-web-store.zip`: chỉ có bản chạy, `manifest.json` ở root, dùng để upload Developer Dashboard. Không upload ZIP có thư mục ngoài cùng/hướng dẫn cài thử.
- `SHA256SUMS.txt`, `BUILD-INFO.json`: thông tin nguồn build, phạm vi quyền và checksum để nhận biết đúng gói. Bản `0.1.0` chưa phải tuyên bố go-live hoặc mức ổn định 1.0.
- Không gửi cả repository, node_modules, fixture, file HTML/API đính kèm, backup, Excel dữ liệu thật, cookie hoặc profile Chrome.

Tái tạo: chạy typecheck, tests, build rồi `node scripts/package-release.mjs`. Gói mới đặt trong `releases/` theo thời điểm đóng gói; không ghi đè gói cũ. Mỗi bản cập nhật Store phải tăng version manifest, giữ version package thống nhất khi lên bản mới.

## Cài thử trước

Theo `INSTALL_INTERNAL_TEST.md`. Thu thập kết quả thực tế từ cả 1–2 user, gồm quyền khác admin, danh sách nhiều trang, No. nhiều cấp, công việc trống/phản hồi và đối chiếu Excel. Không yêu cầu người dùng sắp xếp cột giống tài khoản phát triển như điều kiện cài đặt; nếu phát hiện phụ thuộc view, ghi nhận là lỗi/giới hạn cần xử lý.

Chỉ đánh dấu nghiệm thu sau khi có bằng chứng kiểm thử thực tế. Test tự động/build thành công không thay thế cài Chrome thật, Windows hoặc tài khoản đồng nghiệp.

## Đăng ký và gửi xét duyệt: từng bước

1. Truy cập [Chrome Web Store Developer Dashboard](https://chrome.google.com/webstore/devconsole), đăng nhập đúng tài khoản nhà phát hành nêu trên.
2. Nếu chưa đăng ký, đọc điều khoản và phí hiển thị rồi tự xác nhận/thanh toán. Người phụ trách tài khoản thực hiện, không giao mật khẩu hoặc thẻ thanh toán cho AI. Bật bảo vệ tài khoản theo yêu cầu Google.
3. Chuẩn bị danh sách tài khoản Google của 1–2 tester; về sau dùng Google Group do bạn quản lý để thêm người trong phòng ban. Email công ty không tự có quyền cài nếu chưa gắn với Google Account/nhóm được chỉ định.
4. Chốt người chịu trách nhiệm phát hành/chính sách và quyền sử dụng thương hiệu. Người dùng đã yêu cầu giữ logo, nhưng không tự tuyên bố sản phẩm được tập đoàn bảo trợ/chính thức.
5. Duyệt `PRIVACY_POLICY_DRAFT.md`, thay mục chưa xác nhận và đăng thành một trang HTTPS truy cập được bởi Google/người dùng, chỉ có chính sách, không chứa dữ liệu dự án. URL hiện chưa có. Không tự công khai repo nội bộ để tạo URL.
6. Chuẩn bị ảnh chụp thật của giao diện chạy với dữ liệu giả, không dùng ảnh có tên/nội dung dự án thật. Kích thước và trường ảnh theo Dashboard hiện tại. Icon 128px đã có trong bản build; ảnh quảng bá/screenshot được bổ sung trong bộ tài liệu nếu đã tạo.
7. Bấm Add new item, upload đúng `chrome-web-store.zip`.
8. Điền Store Listing và Privacy practices bằng nội dung dự thảo bên dưới. Điền URL chính sách và email hỗ trợ. Kiểm tra lại metadata/ảnh, không giới thiệu AI hoặc tải file gốc là chức năng đã có.
9. Distribution: chọn `Private`, thêm Trusted testers hoặc Google Group. Không chọn Public/Unlisted. Thiết bị cá nhân không cần enterprise policy để cài bản Store được phép.
10. Test instructions: chức năng 1Office nằm sau đăng nhập. Cần phương án để reviewer kiểm thử, ví dụ tài khoản demo giới hạn chỉ xem dữ liệu giả trên origin được hỗ trợ, được công ty phê duyệt. Không gửi tài khoản admin/cookie/token thật; không tự tạo hoặc mở dữ liệu công ty cho reviewer. Chưa có tài khoản demo, đây là điều kiện còn thiếu, có thể cản trở xét duyệt.
11. Submit for Review chỉ sau khi người phụ trách duyệt thông tin và các điều kiện đã đủ. Bỏ chọn publish tự động để kiểm soát thời điểm phát hành sau xét duyệt. Thời gian xét duyệt/chấp thuận không được bảo đảm.
12. Khi Google duyệt và pilot đạt yêu cầu, publish Private rồi gửi link và hướng dẫn tài khoản Google cho phòng ban. Kiểm tra thời hạn staged publish hiển thị trong Dashboard; nếu để quá hạn có thể phải gửi xét duyệt lại.

## Nội dung Store Listing dự thảo

Tên: T-eraClip.

Mô tả ngắn: Thu thập, xuất thảo luận 1Office ra Excel và lưu nội dung công việc trên máy bằng Capture.

Mục đích duy nhất: Ghi nhận và xuất thông tin công việc do người dùng chủ động chọn để đối chiếu, lưu trữ và chuẩn bị báo cáo.

Mô tả chi tiết:

T-eraClip hỗ trợ người dùng nội bộ thu thập thảo luận trong hệ thống 1Office tại office.meygroup.vn và xuất dữ liệu ra Excel/JSON. Người dùng có thể thu thập trang hiện tại, toàn bộ công việc trong một dự án hoặc danh sách công việc đang lọc, xem tiến độ và đối chiếu dữ liệu trước khi xuất.

Excel dự án phân biệt công việc theo No., mã và tên, giữ thông tin công việc cha/cấp 3, bình luận/phản hồi, tác giả, thời gian và metadata tệp đính kèm. Quyền dữ liệu phụ thuộc tài khoản 1Office đang đăng nhập.

Capture là công cụ phụ để lưu văn bản được chọn, URL nguồn và ghi chú trên máy; có bộ lọc, mẫu nội dung và sao lưu JSON. Không cần tài khoản T-eraClip.

Không có AI, cloud sync, analytics hoặc tải file đính kèm gốc. Đây là bản thử nội bộ; chức năng 1Office yêu cầu quyền truy cập hệ thống của tổ chức.

## Permission Justifications

- storage: Store user-selected Capture items/settings locally and temporary collection/review sessions in Chrome session storage. No Chrome sync or developer backend.
- contextMenus: Add a user-invoked menu action to save explicitly selected text and its page context for work reporting.
- activeTab: Obtain temporary access to the page the user explicitly invokes the extension on; read selected content or the supported 1Office page. No persistent all-sites access.
- scripting: Execute packaged collectors on the supported 1Office tab to read task/discussion data and load additional items selected by the user. All collector code is bundled in the extension; no remotely hosted executable code.
- Remote code declaration: No remotely hosted extension code. 1Office JSON/HTML responses are source data, not downloaded JavaScript executed by T-eraClip.

## Privacy Practices

Không khai báo “không xử lý dữ liệu” chỉ vì dữ liệu ở trên máy. Đối chiếu các nhóm trong Dashboard hiện tại với chính sách: website content (văn bản/thảo luận), thông tin nhận diện cá nhân (tên tác giả/người thực hiện/được nhắc đến), nội dung liên lạc nếu nhóm tương ứng áp dụng. URL trang được chọn là ngữ cảnh nguồn, không phải quét toàn bộ browsing history. Có thể phát sinh dữ liệu nhạy cảm bên trong văn bản người dùng chọn; phải mô tả rõ phạm vi xử lý cục bộ và không tự đánh dấu các nhóm chưa được xem xét.

Không truy cập credentials riêng, dữ liệu thẻ thanh toán, vị trí thiết bị hoặc analytics. Không bán dữ liệu, không dùng ngoài mục đích chức năng, không dùng quyết định tín dụng. Người phụ trách xác nhận các chứng thực sau khi đọc, không để AI chấp nhận điều khoản thay.

## Điều kiện còn thiếu trước khi gửi

- [ ] Tài khoản developer đăng ký thành công và yêu cầu bảo mật tài khoản hoàn tất.
- [ ] Tester/Google Group được xác nhận.
- [ ] Người chịu trách nhiệm và nội dung chính sách được duyệt, có URL HTTPS.
- [ ] Ảnh giao diện dữ liệu giả được duyệt cho listing.
- [ ] Phương án reviewer truy cập 1Office chỉ dữ liệu giả được tổ chức cho phép.
- [ ] Pilot 1–2 user được ghi nhận đạt yêu cầu hoặc lỗi đã xử lý.
- [ ] Người phụ trách duyệt gửi xét duyệt, không tự publish rộng.

## Nguồn chính thức

- [Chuẩn bị ZIP/manifest](https://developer.chrome.com/docs/webstore/prepare)
- [Đăng ký developer](https://developer.chrome.com/docs/webstore/register)
- [Private/Trusted testers/Google Groups](https://developer.chrome.com/docs/webstore/cws-dashboard-distribution)
- [Privacy practices](https://developer.chrome.com/docs/webstore/cws-dashboard-privacy)
- [Upload, test instructions, xét duyệt và deferred publishing](https://developer.chrome.com/docs/webstore/publish)
