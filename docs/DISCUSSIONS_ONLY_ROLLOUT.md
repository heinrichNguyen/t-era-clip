# Triển khai T eraClip Thảo luận cho key user

Ngày rà soát: 03/11/2026. Phạm vi: một công việc hoặc toàn bộ công việc trong một dự án 1Office. Không bao gồm Capture, lịch sử truy cập, ERP, import, đổi No. hoặc đổi mã. Không có bước ghi ngược lên 1Office trong bản này.

## Quyết định đề xuất

Tách **gói phát hành riêng**, không chỉ ẩn nút trong bản đầy đủ. Giữ chung bộ thu thập, kiểm tra nguồn, xem trước và xuất file để tránh hai bản logic bị lệch. Chưa cần tạo nhánh Git dài hạn: working tree hiện có nhiều thay đổi chưa commit; không chuyển nhánh, reset hoặc bỏ thay đổi của người dùng. Sau khi chốt pilot, CNTT có thể tạo nhánh phát hành ngắn hạn từ một commit được duyệt. Việc build profile đã tách được phạm vi mà không nhân đôi dự án.

Build: `npm run build:discussions`; thư mục kết quả: `dist-discussions`. Đóng gói: `node scripts/package-discussions.mjs`. Bản đầy đủ vẫn build bằng `npm run build` ra `dist`.

Gói riêng chỉ có popup thảo luận, trang xem trước một đối tượng, trang thu thập toàn dự án, tài nguyên tĩnh và thư viện dùng chung. Manifest chỉ xin `storage`, `activeTab`, `scripting`; không có background, options, context menu hoặc quyền truy cập thường trực mọi website. Build dừng nếu các module Capture/import/đổi dữ liệu/lịch sử truy cập đi vào bundle. Dữ liệu minh họa và trang chụp hướng dẫn không nằm trong gói phát hành.

## Phạm vi giới thiệu cho phòng ban

1. Công việc đơn: mở chi tiết công việc, thu thập bình luận/phản hồi, xem trước, xuất Excel hoặc JSON.
2. Toàn dự án: mở dự án, kiểm tra danh sách, thu thập thảo luận của từng công việc, xem tiến độ và trạng thái lỗi, xuất Excel. Thảo luận cấp dự án là phạm vi khác; dùng nút trang hiện tại nếu cần phần này riêng.
3. Báo cáo quanh milestone: lọc khoảng ngày, người bình luận, bình luận gốc/phản hồi, hoặc bản ghi mới nhất của mỗi công việc trong phần khớp bộ lọc. Không tự đánh giá tiến độ, xác nhận hoàn thành milestone hay dùng nội dung bình luận làm quyết định phê duyệt.

Lọc ngày nằm ở luồng toàn dự án. Luồng công việc đơn hiện xuất bản thu thập, không có cùng bộ lọc. Không giới thiệu lọc theo người được giao việc, nhận diện tác giả bằng tài khoản duy nhất khi trùng tên, tải nguyên bản tệp đính kèm, AI phân tích hoặc đồng bộ hệ thống như tính năng đã có.

## Các mốc và bằng chứng nghiệm thu

| Mốc | Đầu ra | Bằng chứng và điều kiện |
| --- | --- | --- |
| M0 Chốt phạm vi | Chỉ thu thập và xuất thảo luận | Gói build riêng, không chứa trang hoặc module ghi dữ liệu. Đã có kiểm tra build, cần kiểm tra lại từng bản ZIP. |
| M1 Kiểm tra kỹ thuật | Thu thập, phân biệt phản hồi, cảnh báo thiếu, xuất file và dọn dữ liệu | Typecheck, kiểm thử tự động, build hai profile và kiểm tra gói. Ảnh hướng dẫn là giao diện thật với dữ liệu minh họa, không thay UAT. |
| M2 Pilot công việc đơn | Key user tự xuất và đối chiếu với 1Office | Có bình luận cũ cần tải thêm, phản hồi, xuống dòng, nhắc tên, tiếng Việt, metadata tệp và công việc không có bình luận. Chưa nghiệm thu tài khoản key user thật trong lượt chuẩn bị này. |
| M3 Pilot toàn dự án | Danh sách, bộ lọc và Excel đối chiếu được | Thử dự án nhỏ rồi lớn, dừng giữa lượt, lỗi một số công việc, hết phiên và tài khoản quyền hạn chế. Ghi số công việc dự kiến/đã xử lý/thành công/thiếu/lỗi và thời gian. Không hứa tốc độ chỉ từ số lượng test fixture. |
| M4 Phát hành Private | Người trong nhóm được cài, người ngoài không được cài | Chốt tài khoản publisher, Google Group, URL chính sách, quyền thương hiệu và phương án reviewer được công ty duyệt. Chưa upload hoặc gửi Google xét duyệt. |
| M5 Bàn giao phòng ban | Người dùng tự làm được và biết xử lý cảnh báo | Hướng dẫn kèm ảnh, buổi thực hành, phiếu xác nhận từ key user và đầu mối hỗ trợ. Mở rộng chỉ sau khi các lỗi mức cao đã xử lý. |

Đề xuất pilot: 2–3 key user ở các phòng ban có quyền xem khác nhau. Đây là đề xuất, không phải danh sách người đã tham gia. CNTT phối hợp người quản lý dữ liệu chọn mẫu không nhạy cảm hoặc được phép dùng. Mỗi mẫu phải có người dùng kiểm tra ngẫu nhiên nội dung đầu/cuối, tác giả, thời gian và quan hệ phản hồi trên nguồn; cảnh báo thiếu không được ký nhận là đầy đủ.

## Có nên Public trên Chrome Web Store

**Chưa nên Public.** Công cụ hiện chỉ hỗ trợ các origin Meygroup đã cấu hình và xử lý thông tin thảo luận nội bộ, chưa phải sản phẩm 1Office đa đơn vị. Kênh phù hợp là Chrome Web Store **Private**, ưu tiên Google Group do CNTT quản lý. Theo Google, Public mở cho mọi người; Unlisted cho bất kỳ ai có URL; Private giới hạn người dùng/nhóm được chỉ định. Các chế độ đều chịu yêu cầu chính sách và quy trình review như nhau. Tài khoản trong trusted testers phải gắn Google Account; trusted testers theo publisher, còn Google Group có thể chỉ định theo từng item. Domain publishing chỉ có khi Google Workspace được quản trị viên bật [1].

Unlisted không phải cơ chế bảo mật. Private chỉ giới hạn cài đặt, không thay quyền dữ liệu 1Office, không ngăn người đã tải Excel sao chép file. Không gửi dữ liệu công ty, tài khoản quản trị, cookie hoặc token để phục vụ reviewer. Chỉ dùng tài khoản demo được phê duyệt, dữ liệu giả, quyền tối thiểu; nếu chưa có thì ghi nhận là blocker gửi review.

Google yêu cầu công khai cách xử lý dữ liệu ngay cả khi chỉ xử lý/lưu trên máy [2]. Cần duyệt chính sách riêng cho bản thảo luận: nội dung bình luận, tác giả, thời gian, nhắc tên, metadata tệp, dữ liệu nhận diện công việc/dự án; sử dụng phiên đăng nhập hiện có nhưng không chủ động đọc/lưu mật khẩu hoặc cookie; xử lý trên trình duyệt, dữ liệu xem trước dùng session storage, file tải xuống do người dùng quản lý. Không tuyên bố đã được Google hoặc tập đoàn chứng nhận. Đơn mục đích cũng phù hợp hơn với hướng dẫn chất lượng của Store [3].

Trước khi cân nhắc Public: có quyết định chủ sở hữu sản phẩm/thương hiệu, nhu cầu ngoài nhóm nội bộ, hỗ trợ origin/tenant rõ ràng, UAT quyền hạn chế và dữ liệu lớn, quy trình hỗ trợ/cập nhật, privacy disclosures chính xác và phương án reviewer. Hiện chưa có bằng chứng đủ các điều kiện này.

## Cài thử có kiểm soát

Tài liệu chính để hướng dẫn thao tác hằng ngày là `Huong-dan-truc-quan-T-eraClip-Thao-luan.pdf`, dùng ảnh giao diện, khoanh vùng và đánh số. Bản Word giữ làm tài liệu tra cứu chi tiết cho CNTT và người dùng cần xem thêm. Ảnh dùng dữ liệu minh họa; không thay bằng chứng nghiệm thu trên tài khoản phòng ban.

Chưa cài hoặc publish tự động. Gói pilot có thư mục T-eraClip-Thao-luan, hướng dẫn và ZIP riêng để CNTT chuẩn bị gửi Store. CNTT kiểm tra SHA256 của file được chỉ tên tại trường archive trong BUILD-INFO.json, không nhầm với checksum của gói pilot bên ngoài. Chọn thư mục T-eraClip-Thao-luan có manifest.json bằng Load unpacked trong Chrome Extensions. Cài thử thủ công cần Developer mode và có thể bị chính sách máy công ty chặn; không yêu cầu tắt bảo vệ hoặc vượt chính sách. Bản Private khi được duyệt sẽ có link Store do CNTT cấp, chưa có link trong bộ tài liệu này.

Không dùng đồng thời bản đầy đủ và bản thảo luận khi hướng dẫn key user, tránh bấm nhầm. CNTT hướng dẫn tắt bản không dùng, không tự xóa dữ liệu đang lưu. Sau cập nhật/reload, dữ liệu session cũ có thể mất; tải file cần giữ trước khi thao tác. Bản BETA không đồng nghĩa bản triển khai chính thức.

## Rủi ro và giới hạn còn lại

- Thu thập dùng phiên đăng nhập hiện tại, không thay hay chứng minh cơ chế phân quyền backend 1Office. Phải test quyền tài khoản phòng ban và phiên hết hạn thực tế.
- Toàn dự án dùng API đọc trước, có phương án DOM dự phòng có thể chuyển trang nguồn. Giữ trang kết quả và tab nguồn mở, không đổi dự án trong lúc chạy. Dừng có thể cần chờ nhóm hiện tại hoàn tất.
- Đầy đủ là trạng thái của lượt thu thập, không phải bản sao lưu toàn hệ thống, ảnh chụp nguyên tử hoặc kết luận nghiệp vụ. Dữ liệu sau thời điểm chạy không tự cập nhật.
- Tệp đính kèm gốc không được tải cùng Excel; chỉ metadata có sẵn. File Excel không được mã hóa hay quản lý quyền chia sẻ bởi tiện ích.
- Chưa có kiểm thử trực tiếp Windows/key user phòng ban hoặc chấp thuận Store cho gói mới. Kiểm thử bằng dữ liệu giả và test tự động không thay các mốc này.

## Nguồn chính thức

[1] https://developer.chrome.com/docs/webstore/cws-dashboard-distribution (kiểm tra ngày 03/11/2026).

[2] https://developer.chrome.com/docs/webstore/program-policies/user-data-faq (kiểm tra ngày 03/11/2026).

[3] https://developer.chrome.com/docs/webstore/program-policies/quality-guidelines (kiểm tra ngày 03/11/2026).
