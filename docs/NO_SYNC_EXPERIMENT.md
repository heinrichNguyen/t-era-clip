# Đồng bộ No. theo mã công việc

## Đã xác minh trên dự án test (ID 870)

- Giao diện sửa công việc ID 100119 cho phép đổi cf19 (No.) từ 1.3 sang 1.4.
- Tải lại trang xác nhận giá trị đã lưu. Đã hoàn nguyên về 1.3.
- Các trường biểu mẫu được đối chiếu không đổi ngoài cf19.
- Đây là lưu biểu mẫu qua giao diện, KHÔNG phải kiểm thử API cập nhật riêng một trường.

## Bản đối chiếu trong extension

Sau khi kiểm tra danh sách dự án, mở “Đối chiếu No. hàng loạt (chỉ xem trước)”.
Dán hai cột từ Excel, không có tiêu đề: mã công việc, No. mới.

- Khớp chính xác mã công việc, không khớp theo tên hoặc No.
- Giữ ID nội bộ. Không suy ra hoặc sửa quan hệ cha/con từ No.
- Kiểm tra mã trùng, ID không hợp lệ, No. sai định dạng và No. xung đột với công việc ngoài phạm vi dán.
- Không tự tạo công việc mới, không xóa việc vắng trong file.
- Chỉ đọc ảnh chụp danh sách trong chrome.storage.session; dữ liệu dán nằm trong bộ nhớ trang.
- Nút xóa dữ liệu tạm xóa cả danh sách đối chiếu.
- Bản đối chiếu không có nút áp dụng lên dự án thật. Có mục kiểm thử riêng giới hạn cứng ở dự án test ID 870.

## Runner thử nghiệm ngày 15/10/2026

- Đã quan sát yêu cầu lưu thực tế: POST `/apps/work-task-task/edit?_json=1`, phản hồi JSON có `close`, `reload`, `notice`. Payload chứa toàn bộ biểu mẫu, không chỉ cf19.
- Không thử POST thiếu trường. Runner dùng luồng biểu mẫu giao diện đã render: mở công việc, sửa cf19, bấm nút Cập nhật có sẵn của 1Office.
- Chỉ chấp nhận origin office.meygroup.vn, project ID 870, task ID 100119/mã TEST01.3 và task ID 100122/mã TEST01.INSERT, parent ID 100104.
- Đọc cả hai biểu mẫu trước khi ghi; dùng No. tạm 99001.1 và 99001.2. Chặn khi các giá trị tạm đã có trong danh sách.
- Mỗi lần ghi đều kiểm tra lại nhận diện và các giá trị biểu mẫu. Mở lại biểu mẫu để đối chiếu No. và các trường còn lại. Hoàn nguyên chỉ khi biểu mẫu hiện tại còn khớp dữ liệu vừa ghi.
- Nếu yêu cầu đã gửi nhưng chưa xác minh kết quả, không tự gửi lại hoặc tự hoàn nguyên. Báo cần kiểm tra thủ công.
- Chưa chứng minh chống thay đổi đồng thời ở mức máy chủ, không có giao dịch toàn bộ lô. Cần giữ tab nguồn và trang kiểm thử mở, không sửa công việc trong lúc chạy.
- Kiểm thử mô phỏng đã bao phủ lô hai việc/hoàn nguyên, sai dự án, khóa chạy đồng thời, No. tạm trùng và phản hồi không rõ. Kết quả live phải được xác minh riêng, không suy ra từ kiểm thử mô phỏng.

### Lần chạy live đầu tiên

- Ảnh người dùng gửi cho thấy hết thời hạn 30 giây, không có dòng kết quả công việc. Theo thứ tự runner, chưa đến vòng kích hoạt lưu (vòng này chỉ bắt đầu sau khi đã đọc đủ hai biểu mẫu).
- Đọc nguồn sau lỗi: TEST01.3 vẫn 1.3, TEST01.INSERT vẫn 1.4; không có biểu mẫu sửa mở, còn hộp chi tiết TEST01.3, chưa thấy menu Sửa.
- Chưa xác định chắc chắn lỗi nằm ở đóng chi tiết hay mở menu do bản đầu chưa có nhãn từng bước. Không coi đây là lỗi máy chủ hoặc tăng timeout để che lỗi.
- Đã đổi thao tác điều khiển nút/link sang chuỗi pointerdown/mousedown/pointerup/mouseup/click. Kiểm thử mô phỏng dùng nút đóng xử lý mouseup, menu xử lý mousedown để tránh chỉ thử onclick.
- Bổ sung bước cuối và số lần kích hoạt lưu trong kết quả. Số này đếm thao tác bấm lưu, không khẳng định máy chủ đã nhận từng yêu cầu.
- Cần chạy lại bản mới để xác minh live; chưa công bố đã ghi hàng loạt thành công.

### Lỗi tab hiển thị JSON quickview

- Xác nhận tab nguồn bị điều hướng đến `/apps/work-task-task/quickview?ID=100119` và hiển thị cấu hình dialog JSON. Đã đưa tab về project ID 870, đọc lại TEST01.3 = 1.3 và TEST01.INSERT = 1.4.
- Chuỗi mouseup + click có thể vừa kích hoạt xử lý AJAX của ứng dụng vừa chạy điều hướng mặc định của anchor. Đã chặn default của click trên anchor nhưng giữ truyền sự kiện cho handler của ứng dụng.
- Cũng chặn submit dạng document, và chờ biểu mẫu có `ajax-form-done` trước khi cho phép lưu. Không tự gửi POST riêng hoặc đoán đường ghi khác.
- Thêm “Kiểm tra mở biểu mẫu (không ghi)”: chỉ mở, đọc và đóng hai biểu mẫu; không kích hoạt lưu. Dùng bước này xác minh luồng giao diện trước khi thử ghi lại.
- Kiểm thử mô phỏng có handler mở link ở mouseup và kiểm tra mọi click/submit mặc định đã bị chặn. Chưa xác minh bản sửa này trên giao diện live.

### Menu Sửa vẫn không mở, 0 lần kích hoạt lưu

- Ảnh lần chạy tiếp theo xác nhận dừng ở mở menu Sửa của TEST01.3, số lần kích hoạt lưu bằng 0.
- Đọc mã JavaScript giao diện đang tải trên tab nguồn: plugin popup mở menu click bằng sự kiện riêng `realclick`; core tạo sự kiện này từ pointerdown/pointerup có tọa độ gần nhau. Wrapper xử lý sự kiện có bước bất đồng bộ trước khi gọi handler.
- Suy luận cần kiểm chứng live: phát toàn bộ chuỗi đồng bộ có thể khiến handler nhấn chưa xử lý xong khi thả chuột. Runner nay chờ 25 ms giữa nhấn/thả và trước click, kiểm tra lại guard sau mỗi lần chờ; chờ nút menu có `pop-done` trước khi thao tác.
- Fixture mô phỏng nay có handler pointerdown bất đồng bộ và menu chỉ mở bằng `realclick`. Kiểm thử mô phỏng không chứng minh extension đã mở được menu trên hệ thống thật.
- Chưa chạy ghi live sau thay đổi này. Bước tiếp theo chỉ chạy “Kiểm tra mở biểu mẫu (không ghi)”.

### Chờ nhấn/thả chưa giải quyết được lỗi live

- Ảnh người dùng gửi sau bản sửa vẫn báo timeout mở menu Sửa, 0 lần kích hoạt lưu. Không coi giả thuyết thời gian xử lý sự kiện là nguyên nhân đã xác nhận.
- Đọc DOM tab nguồn xác nhận nút menu đang hiển thị và có `pop-done`; hộp chi tiết cũ bị ẩn, hộp hiện tại có No. 1.3.
- Đổi riêng thao tác menu: phát một `MouseEvent('realclick')` có tọa độ nút và `which = 1`, không phát thêm chuỗi pointer/click lên menu. Không truy cập biến nội bộ của ứng dụng, không gửi POST trực tiếp.
- Fixture kiểm tra menu nhận đúng một realclick mỗi lần mở, không nhận thêm pointer/click; vẫn kiểm tra bước chỉ đọc không kích hoạt lưu. Cần xác minh live trước khi chạy ghi.

### Chẩn đoán menu, không đổi thời hạn chờ

- Bổ sung mục “Chẩn đoán mở menu (bản 1)” trong kết quả: số menu `.list-menu` hiển thị trước thao tác và nhiều nhất khi chờ; số liên kết edit đúng task, số đang hiển thị và số có title chính xác Sửa; trạng thái nút menu còn trên DOM/hiển thị.
- Chỉ lưu số đếm và nhận định giới hạn theo DOM quan sát được. Không thu HTML, URL đầy đủ, token hoặc thông tin nhân sự. Không kết luận máy chủ chậm hay handler đã nhận sự kiện từ các số đếm này.
- Vẫn chờ tối đa 30 giây mỗi bước, không thêm lần bấm hoặc lần ghi. Kiểm thử riêng các trường hợp không có menu, liên kết ẩn, title khác, và hiển thị kết quả chẩn đoán trên trang đối chiếu.

### Điều chỉnh sau lần mở menu thủ công

- Ảnh live xác nhận đã đọc TEST01.3, No. 1.3, 0 lần kích hoạt lưu; sau đó dừng ở đóng hộp chi tiết cũ. Việc đọc biểu mẫu không chứng minh runner tự mở menu được.
- Đọc DOM nguồn thấy nút đóng chi tiết là `.dialog-close[pop-close]`. Runner nay chỉ chọn nút đóng đang hiển thị thuộc đúng dialog, không chọn nút trong dialog con bị ẩn.
- Menu và nút đóng dùng một helper phát realclick, không thêm chuỗi pointer/click lên cùng nút. Liên kết mở chi tiết/Sửa vẫn giữ chặn điều hướng document như trước.
- Runner thử nghiệm chuyển từ ngữ cảnh mặc định ISOLATED sang MAIN, cùng ngữ cảnh với handler giao diện. Đây là thay đổi để kiểm chứng giả thuyết lệch ngữ cảnh sự kiện, chưa xác nhận là nguyên nhân live.
- Không thêm quyền manifest, không dùng debugger, không truy cập biến nội bộ 1Office, không thêm message listener/cầu nối lệnh. Hàm vẫn giới hạn cứng origin/dự án/hai task, chỉ chạy theo nút người dùng bấm. Dữ liệu biểu mẫu tạm không được ghi storage hay log.
- Giới hạn bảo mật: MAIN chia sẻ môi trường JavaScript với trang, nên trang có thể tác động API DOM/kết quả. Không coi kết quả runner là bằng chứng chống giả mạo hoặc cơ chế cấp quyền; chưa dùng phương án này cho ghi dự án thật.
- Kiểm thử mô phỏng bao phủ đóng biểu mẫu/chi tiết bằng realclick, bỏ qua nút đóng ẩn trong dialog con, chỉ đọc hai việc không lưu và lựa chọn MAIN ở adapter. Chẩn đoán UI mang nhãn bản 2. Cần chạy lại bước không ghi trên hệ thống thật.

### Tab chuyển sang JSON sau khi chuyển runner sang MAIN

- Kiểm tra trực tiếp Chrome xác nhận tab nguồn ở quickview ID 100119 và hiện JSON cấu hình dialog. Đã đưa về project 870 qua lịch sử tab.
- Thao tác Chrome thông thường mở chi tiết, menu Sửa, biểu mẫu No. 1.3 và đóng cả biểu mẫu/chi tiết thành công. Không bấm Cập nhật, không đổi trường dữ liệu. Điều này xác minh luồng giao diện, không xác minh runner extension.
- Runner còn gọi `anchor.click()` sau chuỗi nhấn/thả. Đã bỏ click bổ sung trên cả link quickview và link Sửa; giữ cơ chế chặn default trong lúc phát nhấn/thả. Đây là loại bỏ một đường kích hoạt thừa, chưa kết luận chính xác handler nào đã điều hướng ở lần live.
- Thêm kiểm thử mô phỏng handler click chủ động điều hướng kể cả khi default bị chặn. Bước chỉ đọc phải hoàn tất mà không phát click trên anchor, không đổi URL và không lưu. Nhãn chẩn đoán bản 3.

### Inspect Sources sau khi bản 3 không mở được chi tiết

- Ảnh live bản 3 dừng tại mở chi tiết TEST01.3, 0 lần kích hoạt lưu. Fixture trước đó mở link bằng mouseup không phản ánh handler click dùng chung thực tế, nên không được dùng để khẳng định bản 3 hoạt động.
- Đọc trực tiếp Sources trên tab HTTP 1Office, không thực thi Console: `packages.plugin-pop.desktop.js?t=1791197954633`, module `init_events_on_document.js` dùng WeakMap theo lượt thao tác, reset khi pointerdown, tra `s.has(...)` trước mở/đóng. Phát riêng realclick không reset trạng thái này.
- `packages.core-app.js?t=1791197954633`, `registerPopState` đăng ký click cho anchor nội bộ rồi gọi app.load. `packages.core-js.js?t=1791197954633`, wrapper g/E chuyển thành SnapEvent và await EVENT_HANDLING trước gọi listener. Suy luận: listener preventDefault đăng ký qua bộ bọc này có thể chạy quá muộn để chặn điều hướng document trong MAIN. Chưa khẳng định toàn bộ chuỗi lỗi live chỉ có nguyên nhân này.
- Bản 4: popup phát pointerdown, chờ một lượt xử lý rồi phát một realclick (không pointerup gây realclick thứ hai); bỏ qua realclick nếu nút đã ẩn/được đóng từ pointerdown. Link dùng đúng một click, không chuỗi mouseup, với chặn default đồng bộ qua thuộc tính onclick. Tương tự chặn submit đồng bộ qua onsubmit. Giữ/gọi handler cũ và khôi phục khi kết thúc, không ghi đè handler mới nếu trang đã thay trong lúc chạy.
- Fixture đổi sang mở link bằng click, yêu cầu pointerdown trước realclick, và mô phỏng addEventListener trì hoãn handler. Kiểm tra default của click/submit bị chặn ngay, nhưng luồng ứng dụng vẫn mở biểu mẫu/lưu mô phỏng/hoàn nguyên. Đây vẫn là mô phỏng, không phải kiểm thử ghi live.

### Bản 5, bỏ điều khiển dropdown

- Bản 4 live vẫn không mở menu: có đúng một liên kết edit trong trang, nhưng liên kết nằm trong `.list-menu.hidden`. 0 lần kích hoạt lưu.
- Inspect tiếp `testEventOpen`, `getControl`, `testEventBtnClose` trong bundle popup xác nhận có điều kiện theo control, trạng thái cha/con và lượt thao tác. Chưa xác định chắc chắn điều kiện nào làm menu không mở trong runner.
- Đổi luồng: sau khi mở đúng chi tiết task, tìm liên kết Sửa có sẵn trong `.list-menu`, khớp origin/path/ID/title. Dùng handler click nội bộ đã inspect từ core-app, không mở dropdown, không unhide DOM, không tạo link/đoán endpoint hoặc POST trực tiếp.
- Chỉ chấp nhận việc link bị ẩn do `.list-menu.hidden`. Bỏ qua action có hidden/class hidden/aria-hidden/disabled/aria-disabled, sai origin/ID/title hoặc bị ẩn trực tiếp qua CSS. Không coi liên kết có trong DOM là chứng minh quyền sửa; máy chủ vẫn phải kiểm tra quyền và runner vẫn xác minh đầy đủ nhận diện biểu mẫu trước khi ghi.
- Nếu có nhiều link hợp lệ cho cùng task, chọn link cuối trong DOM để tránh menu cũ được giữ lại. Đây không phải bảo đảm chống thay đổi đồng thời; nhận diện và fingerprint của form vẫn là bước bắt buộc.
- Fixture dựng link edit sẵn trong menu thu gọn, không tạo link khi bấm menu. Bước chỉ đọc phải mở hai form với 0 tương tác dropdown, 0 lần lưu. Có ca từ chối action bị ẩn riêng, bị vô hiệu hóa, sai title/origin và link không tồn tại.
- Bản 5 chưa được xác minh live. Tiếp tục chỉ thử mở/đọc/đóng form, không chạy ghi.

### Bản 6, sửa nhận diện chi tiết theo trạng thái công việc

- Ảnh live bản 5 đã đọc TEST01.3 No. 1.3, sau đó timeout mở TEST01.INSERT, 0 lần lưu. Inspect trực tiếp tab nguồn thấy chi tiết TEST01.INSERT đã hiện, mã TEST01.INSERT và No. 1.4.
- Header công việc thứ hai có econtract/progress ID 100122 nhưng không có member. Điều kiện cũ bắt buộc anchor member nên không thể nhận diện hộp chi tiết này, dù đã mở. Đây là lỗi nhận diện cụ thể đã đối chiếu với DOM, không phải bằng chứng máy chủ cần hơn 30 giây.
- Thay điều kiện bằng mã công việc chính xác ở `.detail-field` có nhãn Mã công việc và ID chính xác qua URL action header thuộc allowlist member/econtract/progress, cùng origin. Không dùng substring ID, tên hoặc No. làm khóa. Nếu không có đủ bằng chứng thì vẫn dừng, không nới nhận diện bằng tên.
- Fixture công việc thứ nhất có member, công việc thứ hai chỉ có progress, không có member. Ca chỉ đọc phải đọc cả hai, giữ No. 1.3/1.4 và 0 lần lưu. Ca âm tính chặn ID 1001220 hoặc mã sai ở chi tiết thứ hai trước khi ghi.
- Giữ luồng dùng link Sửa sẵn có của bản 5. Bản 6 chưa được chạy live, chưa cho phép suy ra ghi hàng loạt đã thành công.

### Bản 7, không phụ thuộc menu đã từng mở

- Bản 6 nhận diện được chi tiết TEST01.INSERT nhưng không thấy link edit. Đọc DOM trước mở menu chỉ có link edit ID 100119; bấm menu TEST01.INSERT qua Chrome mới xuất hiện link edit ID 100122. Giả định mọi link edit đã được dựng trước là sai; kết quả trước có thể phụ thuộc cache menu.
- Đã xác minh trực tiếp cả hai href: `/apps/work-task-task/edit?ref=detail&_quickview=1&ID=100119` và cùng query với ID 100122. Bổ sung đúng hai URL này vào allowlist thử nghiệm, không suy rộng template sang task/dự án khác.
- Nếu không có link edit cho task, runner gắn tạm một anchor vào đúng chi tiết đã xác minh ID/mã, dùng href allowlist và handler click 1Office (đã chặn default đồng bộ). Giữ anchor đến khi form xuất hiện hoặc timeout, rồi xóa đúng node tạm đó. Không gọi POST riêng, không cần mở dropdown.
- Nếu có action edit nhưng sai origin/title, bị ẩn riêng hoặc vô hiệu hóa, dừng thay vì dùng URL dự phòng. Việc không có action không được coi là xác nhận quyền sửa; máy chủ vẫn quyết định có trả form hay không, nhận diện/fingerprint vẫn được kiểm tra trước bất kỳ lần lưu nào. Không có form thì dừng với 0 lần lưu ở bước chuẩn bị.
- Fixture task thứ hai không có menu/link edit ngay từ lần mở đầu. Có ca cả hai task đều chưa dựng menu và ca server không trả form: không mở dropdown, link tạm được dọn sau thành công hoặc lỗi, không tự gửi lại.
- Phải tải lại tab nguồn khi xác minh live để tránh menu đã được mở trong lúc inspect làm sai lệch kết quả cold-start. Bản 7 chưa được xác minh live, tiếp tục chỉ chạy nút không ghi.

### Bản 8, kiểm soát tab nguồn sau khi đối chiếu Network

- Lần cold-start người dùng gửi trước đó dừng tại tải form TEST01.3 bằng link tạm, 0 lần lưu. Khi inspect sau đó có thấy biểu mẫu Sửa trên tab nguồn, nhưng không dùng giá trị input bị công cụ che để kết luận ID sai.
- Ghi nhận lại từ trang mới trong Network: mở qua menu thủ công dùng GET edit với query `_json=1`, `reloadCsrfToken=1`, `inlineLogin=1` và tham số chống cache; 200 OK, khoảng 85 ms. Preview có `dialog.header = Sửa công việc` và `dialog.content` chứa elem form. Không xuất cookie, header/token hoặc HAR.
- Người dùng chạy lại bản 7 khi Network đang mở: hai request edit cùng cơ chế đều 200 OK, khoảng 55 ms và 43 ms. Người dùng xác nhận đã mở/đọc đủ hai form; tab nguồn không còn form sửa hiển thị. Không có thay đổi code giữa lần thất bại và thành công này. Vì vậy không kết luận URL/link tạm luôn thiếu ngữ cảnh khởi tạo.
- Đọc `packages.core-dom.js?t=1791197954633`: `VDOM.nextFrame`/`window.nextFrame` chờ hai requestAnimationFrame. Tài liệu Chrome mô tả việc dừng rAF ở tab nền: https://developer.chrome.com/blog/background_tabs/ . Đây là cơ sở cho giả thuyết trạng thái tab nền/render, chưa chứng minh là nguyên nhân duy nhất.
- Bản 8: kiểm tra URL tab nguồn vẫn đúng project 870 bằng tabs.get, kích hoạt đúng tab bằng tabs.update({active:true}) rồi mới inject. Không thêm quyền manifest, không đổi cài đặt throttling/debugger và không giành focus liên tục. tabs API cơ bản không cần thêm quyền tabs: https://developer.chrome.com/docs/extensions/reference/api/tabs .
- Runner chỉ bắt đầu khi document visible; nếu chuyển tab/thu nhỏ trong lúc chờ hoặc trước thao tác thì dừng với thông báo riêng, không chỉ báo timeout. Nếu có ghi chưa rõ kết quả vẫn không tự gửi lại/hoàn nguyên. Không tự chuyển về trang review khi kết thúc; người dùng quay lại sau khi luồng mở/đóng đã xong.
- Kiểm thử mô phỏng bao phủ thứ tự get/update/inject, không focus tab sai dự án, dừng tab hidden trước chạy, và chuyển hidden sau ghi chưa rõ. Cần xác minh cold-start bản 8 khi DevTools đã đóng trước khi thử ghi.

## Chưa được xác minh, bắt buộc trước khi bật ghi

1. Xác minh yêu cầu cập nhật thực tế, hành động, nội dung gửi, cơ chế phiên/CSRF và phản hồi. Không đoán endpoint, không lưu token vào tài liệu.
2. Chứng minh máy chủ hỗ trợ ghi riêng No., hoặc chứng minh cách giữ nguyên đầy đủ các trường khi lưu biểu mẫu.
3. Thử extension cập nhật ít nhất hai công việc trên test, đọc lại từ máy chủ và hoàn nguyên.
4. Kiểm tra lỗi giữa chừng, mất phản hồi sau khi đã ghi, thay đổi đồng thời và hết phiên. Không tự gửi lại yêu cầu ghi nếu chưa kiểm tra trạng thái hiện tại.
5. Trước mỗi lần áp dụng, đọc lại ID/mã/No. và xác nhận dự án. Dừng khi dữ liệu nguồn thay đổi; ghi nhật ký kết quả từng việc. Hoàn nguyên cũng phải kiểm tra trạng thái, không ghi đè sửa đổi mới của người khác.

Đối chiếu hợp lệ không đồng nghĩa một lô cập nhật đã an toàn. Chưa có cam kết giao dịch toàn bộ hoặc hoàn tác nguyên tử từ 1Office.

## Cập nhật 2026-10-16: kết quả bản 8 và ghi theo danh sách bản 9

- Người dùng gửi ảnh bản 8: đã ghi, đọc lại và hoàn nguyên cả hai công việc, 4 lần kích hoạt lưu. TEST01.3: 1.3 -> 99001.1 -> 1.3; TEST01.INSERT: 1.4 -> 99001.2 -> 1.4. Đây là bằng chứng kết quả runner trên hai task test, không phải xác nhận ghi hàng loạt trên dự án thật. Các mục chưa xác minh ở trên được thay thế tương ứng cho phạm vi hai task này, không suy rộng quyền hoặc hợp đồng API.
- Bản 9 nối dữ liệu dán với nút ghi theo bảng xem trước, có xác nhận riêng. Khóa origin/project 870, chỉ cho thay đổi ID 100119/TEST01.3 và ID 100122/TEST01.INSERT, cha 100104. Không tạo mới, không đổi code/cha/con. Các công việc test khác và mọi dự án thật chỉ xem trước.
- Dựng lại plan khi bấm ghi, yêu cầu inventory đầy đủ, không lỗi/trùng No. Kiểm tra toàn bộ danh sách No. trong DOM nguồn khớp snapshot (số lượng và giá trị), rồi đọc cả hai biểu mẫu được chọn trước lần lưu đầu, kiểm tra No. khớp preview. Nếu giao diện không hiển thị đủ danh sách thì dừng, chưa có API inventory tổng quát cho luồng ghi này.
- Hai giai đoạn: đưa các task thay đổi sang No. tạm đã dành riêng, đọc lại; sau đó ghi No. cuối và đọc lại. Cho phép hoán đổi 1.3/1.4 mà không tạo trùng No. ở giữa luồng. Kết quả cuối được giữ lại, KHÔNG tự hoàn nguyên như nút thử nghiệm bản 8.
- Không gửi lại hoặc tự rollback plan lỗi. Hiển thị riêng No. cuối đã xác nhận và No. tạm đã xác nhận. Có thể còn cập nhật dở, cần kiểm tra thủ công. Sau bất kỳ plan có lần kích hoạt lưu, hoặc lỗi không nhận được kết quả, khóa ghi tiếp trong trang review hiện tại; phải kiểm tra 1Office và lấy inventory mới. Khóa UI không phải cơ chế phục hồi bền vững qua đóng/reload trình duyệt.
- Vẫn chờ tối đa 30 giây mỗi bước, giữ tab nguồn visible, fingerprint các trường khác trước/sau lưu. Đây không phải kiểm soát đồng thời nguyên tử ở máy chủ; vẫn có cửa sổ người khác sửa sau khi đọc form. Không được dùng trên dự án thật cho đến khi giải quyết phạm vi/đối chiếu và phục hồi phù hợp.
- Sửa câu chẩn đoán mở form để không còn gây hiểu nhầm là chưa lưu khi toàn bộ lượt ghi đã thành công. Dừng các lần hoàn nguyên kế tiếp của bài thử cũ nếu một lần hoàn nguyên đã gửi mà chưa rõ kết quả.
- Xác minh local: 176 tests/24 files qua, typecheck và build qua. Ca mới: hoán đổi qua số tạm, giữ No. cuối, cập nhật một task, stale snapshot/form, ID ngoài allowlist, ID trùng, trùng No., No. tạm bị dùng, lỗi giữa chừng, lỗi lần ghi cuối và xác nhận UI/khóa ghi tiếp.
- Chưa chạy live bản 9. Kiểm thử tiếp: reload extension, kiểm tra lại danh sách test, mở đối chiếu mới, dán hai cột TEST01.3 [TAB] 1.4 và TEST01.INSERT [TAB] 1.3, xem trước rồi xác nhận ghi. Chỉ chạy khi snapshot hiện tại vẫn là 1.3/1.4. Sau chạy kiểm tra trên 1Office, không dùng lại preview cũ. Muốn trả lại thứ tự cũ cần lấy inventory mới và lập plan ngược riêng.

### Bản 10: phân biệt No. gốc với outline tự sinh

- Người dùng gửi lỗi bản 9 ở bước chuẩn bị, 0 lần lưu. Inspect DOM nguồn test qua Chrome xác nhận 8 ô cf19: 1, 1.1, 1.2, 1.3, 1.4, ô trống, 2, 3. Ô trống thuộc TEST01.4/ID 100143. Không gửi thao tác ghi trong lượt điều tra.
- Code inventory API dùng `readCustomNo(id) || outlineByTaskId.get(id)` làm `no`. Do đó bản 9 so sánh No. tự sinh của task trống với chuỗi rỗng thật trên UI. Đây là lỗi dùng trường hiển thị để xác minh giá trị nguồn.
- Thêm metadata `sourceNo` riêng: chuỗi rỗng là giá trị cf19 trống đã xác minh, undefined là chưa đọc được. Giữ nguyên `no` và hành vi xuất báo cáo hiện có. Collector DOM chỉ ghi sourceNo khi thật sự có ô cf19, không lấy cột thứ tự thay thế.
- Payload ghi dùng sourceTasks gồm ID và sourceNo. Runner đọc `.gantt-body-row[data-side="left"][item-type="item"]`, ID qua `item-id="item_<ID>"`, cf19 từng dòng; kiểm tra đủ số lượng, không trùng ID và khớp No. theo ID. Không còn so multiset số, không bỏ qua ô trống, không tự điền No. cho task ngoài phạm vi.
- Inventory cũ chưa có sourceNo bị từ chối trước khi activate/inject. Bắt buộc reload extension, kiểm tra danh sách dự án lại rồi mở review mới. Bản 10 chưa chạy live; việc inspect chỉ xác nhận cấu trúc DOM và No. đang hiển thị, không xác nhận đã ghi thành công.

## Bản 11, 2026-10-16: nhật ký phục hồi và giảm thao tác ghi

- Người dùng đã gửi ảnh bản 10 xác nhận hai task hoán đổi xong, 4 lần lưu. Đây là kết quả live của bản 10, không phải kiểm chứng bản 11.
- Trước khi activate/inject lượt ghi theo bảng, extension lưu và đọc lại ý định ghi trong `chrome.storage.local`, key `oneOfficeNoWriteJournalV1`. Chỉ giữ lượt gần nhất: origin/project cố định, UUID, thời gian, ID/mã task, No. trước/cuối/tạm và trạng thái đã xác nhận. Không lưu tên công việc, fingerprint biểu mẫu, cookie, CSRF/token, nội dung thảo luận hoặc danh sách nhân sự. Không thêm manifest permission hoặc server/bridge.
- Nhật ký tồn tại qua đóng review/khởi động Chrome, nhưng không tồn tại nếu gỡ extension/xóa dữ liệu profile. Đây là dữ liệu cục bộ, không phải két mã hóa hay bản backup máy chủ. Không xóa cùng nút dữ liệu báo cáo tạm. Nhật ký pending không có nút xóa; chỉ xóa nhật ký đã đối soát khi user xác nhận. Không tự hết hạn pending.
- Lưu ý giới hạn checkpoint: ghi ý định của tất cả task trước toàn lượt, cập nhật các kết quả đã đọc lại sau khi runner trả về. KHÔNG lưu tiến độ mỗi lần click lưu. Nếu review bị đóng/mất phản hồi, các task chưa có kết quả được giữ unknown/pending, dù có thể đã ghi trên 1Office. Không được suy unknown là chưa từng gửi lưu.
- Web Locks trong extension serialize ghi/đối soát/xóa nhật ký giữa các trang review. Không có locks API hoặc lưu local thất bại thì từ chối ghi. Pending chặn lượt ghi mới ngay ở wrapper, không chỉ disable UI. Không khóa được user khác trên 1Office, không hứa giao dịch nguyên tử.
- Nút `Kiểm tra trạng thái trên 1Office (không ghi)` mở hai form allowlist, đọc ID/mã/dự án/cha như trước, cho phép đọc No. tạm/trống/khác mẫu để phân loại conflict, tuyệt đối 0 lần lưu. Chỉ kết quả đọc đủ mới cập nhật journal. Tất cả ở after -> verified; tất cả ở before -> unchanged; No. tạm/một phần/khác dự kiến -> pending, yêu cầu xử lý thủ công. Không có tự tiếp tục/rollback. Sau kiểm tra cần snapshot dự án mới để ghi tiếp.
- UI bỏ nút ghi-hoàn nguyên bài thử cũ, giữ nút kiểm tra form không ghi; mọi thao tác ghi từ UI đi qua bảng có nhật ký. Runner thử nghiệm cũ vẫn giữ cho tests, không dùng làm đường ghi mới.
- Tối ưu thứ tự bằng occupancy theo No. gốc: ghi các đích đang trống trước, chỉ phá vòng bằng một No. tạm. Hai task hoán đổi: 3 lần lưu và 8 lần tải form thay vì 4/10. Một task sang No. trống: 1 lần lưu và 3 lần tải form thay vì 2/5. Chuỗi dịch không vòng: 1 lần lưu/task. Đây là số thao tác kiểm thử mô phỏng, không phải benchmark thời gian live.
- Vẫn pre-read toàn bộ task đổi trước lần ghi đầu, kiểm tra fingerprint trước/sau mỗi lần lưu, không cache/replay POST, không bỏ đọc lại, không song song ghi UI. Hiển thị formLoads/durationMs của lượt thực tế để lấy số đo sau. Chưa có endpoint bulk update được xác minh.
- Giới hạn hiện tại vẫn project 870, hai ID cho phép, DOM phải chứa đủ snapshot. Chưa phù hợp chạy 100-1000 task: mở form tuần tự và DOM virtualization vẫn là nút thắt; cần xác minh API cập nhật riêng No./bulk trước khi thay cơ chế vận chuyển. Không cam kết thời gian 30 giây cho quy mô đó.
- Ca mới bao phủ journal pending qua mở lại UI, storage lỗi chặn inject, mất kết quả giữ pending/chặn replay, đối soát before/after/temp/conflict, không xóa pending, dữ liệu journal tối thiểu, lock bị giữ, bản ghi hỏng; kiểm tra số form/save và không trùng No. qua từng bước của hoán đổi/chuỗi dịch. Bản 11 chưa chạy live.
- Xác minh cuối lượt: 196 tests/25 files qua; typecheck, build và diff whitespace check qua. Không gửi thao tác ghi trên 1Office trong lượt triển khai bản 11.

## Bản 12, 2026-10-16: kiểm tra ngầm, chỉ mở form khi lưu

- Luồng ghi theo bảng bật silentReads. Dùng hai API đọc đã có trong inventory: GET gantt project_id=870/structure=1/group_type=phase và POST gantt mode=enrich, project_id=870, ids theo nhóm tối đa 50. POST enrich là yêu cầu đọc, không phải cập nhật công việc. Không triển khai POST ghi riêng No./bulk.
- Preflight đọc snapshot mới từ server: tập ID dự án phải khớp đầy đủ bản xem trước, cf19 phải có giá trị nguồn rõ ràng (chuỗi/số, kể cả chuỗi trống); không dùng outline dự phòng, không dựa toàn bộ DOM để xác nhận danh sách. Hai task test phải đúng code, title kiểu string, parent 100104. Thiếu/sai schema hoặc ID trùng đều dừng trước lưu.
- Không mở các form preflight nữa. Khi thực sự lưu, vẫn dùng đúng open/save UI đã xác minh, kiểm tra ID/code/project/parent/No. hiện tại, giữ fingerprint FormData trong bộ nhớ để xác nhận thao tác chỉ thay cf19. Lần mở form đầu của mỗi task lấy baseline FormData, không giả vờ fingerprint API tương đương fingerprint biểu mẫu.
- Sau mỗi save, gọi lại GET structure + enrich cho đúng ID để xác nhận No. và fingerprint của các trường API ngoài cf19. Fingerprint gồm identity, các raw task.fields và enrich.fields, không dùng HTML display snippets của record hay metadata cập nhật. Không lưu raw fields/fingerprint này vào journal hoặc logs. API không bảo đảm trả mọi trường của form; chưa có chứng minh bảo vệ tất cả trường phía server.
- Request dùng credentials same-origin, cache no-store, redirect error, header X-Requested-With, AbortController 30 giây bao gồm đọc body. Không retry tự động. Sau khi gửi lưu, phản hồi đọc thiếu/sai/No. cũ/trường khác thay đổi đều giữ kết quả chưa rõ và nhật ký pending, không gửi bước tiếp theo. Chưa chứng minh no-store ngăn mọi cache ứng dụng phía server; kết quả chưa mới sẽ fail closed.
- Nút đối soát journal cũng đọc ngầm, không mở form hoặc gửi lưu. Nút kiểm tra mở form cũ vẫn chỉ đọc UI, phục vụ chẩn đoán, không phải đường cập nhật.
- Hoán đổi mô phỏng: từ 8 lượt mở form bản 11 xuống 3 (chỉ 3 lần lưu), có 8 request đọc ngầm. Đích trống: từ 3 lượt mở form xuống 1. Không suy giảm tương ứng về thời gian hoặc tổng tải backend khi chưa đo live. Hiện postcheck vẫn đọc cấu trúc toàn dự án mỗi bước; đây là điểm cần tối ưu tiếp trước quy mô 1000, không tăng concurrency ghi UI.
- Lượt điều tra chỉ thử GET edit trực tiếp trong tab mới, trả dialog content trống, không coi đó là bằng chứng XHR edit đọc đủ form. Đã đóng tab thử, không gửi ghi. Vì không có phản hồi form đủ để chứng minh parser, không chọn cách dựng FormData ngầm từ edit JSON.
- Kiểm thử mới: API preflight/postcheck hoàn thành swap với 3 form; một đích trống với 1 form; thiếu field/sai key ID/chuyển hướng/timeout chặn 0 save; stale No. hoặc protected field đổi sau save dừng sau đúng 1 save; recovery temp/trống qua API với 0 form/0 save. Bản 12 chưa chạy live.

## Bản 13: bỏ bước mở chi tiết trước biểu mẫu lưu

- Người dùng phản hồi bản 12 nhanh hơn; chưa cung cấp số đo hoặc ảnh kết quả lượt đó, không suy thành benchmark hay xác nhận mọi trường được giữ nguyên.
- Bản 13 bật directEdit cho plan: sau API preflight đã xác minh ID/mã/cha, dùng đúng editHref allowlist để mở form trực tiếp. Không click quickview trước mỗi save. Anchor tạm nằm cùng app root với link task hiện có; vẫn yêu cầu link task trong dự án, không sinh URL cho ID ngoài allowlist.
- Giữ nguyên các điều kiện form POST action/ID/code/project/parent, No. trước ghi và fingerprint; action edit sẵn có nhưng vô hiệu hóa/sai origin/title vẫn chặn, không dùng URL dự phòng để vượt qua. Form không xuất hiện thì timeout 30 giây và dọn anchor, không fallback mở chi tiết rồi gửi lại.
- Nút kiểm tra mở biểu mẫu chỉ đọc cũng thử đường trực tiếp mới, sau API identity preflight, mở/đọc/đóng cả hai form và 0 lần lưu. Cho phép xác minh ngữ cảnh mở form khi chưa có detail trước khi thử ghi bản 13.
- Vẫn 3 lần lưu/form cho swap và đọc ngầm sau mỗi save; phần giảm thêm là 3 lần mở chi tiết công việc. Không giảm thời gian chờ ổn định form hoặc tăng concurrency ghi. Không đổi journal/manifest hoặc mở phạm vi dự án thật.
- Tests mô phỏng xác nhận chỉ activate edit links, không tồn tại task-quick-view, form+anchor được dọn sau probe, chặn edit disabled và dừng timeout ở form thứ hai với đúng một save trước đó. 208 tests/25 files qua, typecheck/build/diff check qua. Bản 13 chưa xác minh live.

### Xác nhận live bản 13 và điều tra request lưu

- Người dùng gửi ảnh probe đọc đủ cả hai form trực tiếp, 0 lần lưu; No. lúc đó TEST01.3=1.6, TEST01.INSERT=1.7.
- Ảnh kết quả lượt ghi tiếp theo xác nhận TEST01.3 1.6 -> 1.7 và TEST01.INSERT 1.7 -> 1.6; 3 lần lưu, 3 lượt tải form, 8 request đọc ngầm, 6.3 giây. Đây là một lượt hoán đổi hai task, không suy ra benchmark 100-1000 task.
- Ngày 2026-10-16 bắt đầu inspect request lưu thực tế. Tab nguồn mới ID 2097984679, project 870; đọc UI nguồn xác nhận No. 1.7/1.6. Chỉ mở chi tiết TEST01.3, chưa bấm lưu và chưa gửi custom POST.
- Chưa thu được payload/response POST vì người dùng chuyển Chrome sang tab khác trong lúc bật DevTools. Đã hỏi người dùng giữ tab test trước khi tiếp tục; không coi đây là bằng chứng endpoint không hỗ trợ cập nhật riêng No. Không đoán cấu trúc POST, không xuất HAR/cookie/token và chưa bật writer ngầm.

### Inspect POST thực tế sau khi người dùng cho tiếp tục

- Trên tab HTTP project 870, mở Sửa TEST01.3/100119 bằng UI. Network ghi GET edit `_json=1&reloadCsrfToken=1&inlineLogin=1` trả 200. Không sửa giá trị form; bấm CẬP NHẬT đúng một lần với No. 1.7 giữ nguyên. Không đụng TEST01.INSERT hoặc dự án thật. Việc lưu này có thể thêm lịch sử hoạt động.
- POST thực tế: `https://office.meygroup.vn/apps/work-task-task/edit?_json=1`, HTTP 200 OK; request Content-Type multipart/form-data, Content-Length 3997 cho lần thử này. Đây không phải kích thước cố định của hợp đồng API.
- Query `_json=1`; body có 38 trường ghi nhận theo tên (không giữ giá trị): app_ctx, ID, code, title, phase_type, start_plan, end_plan, owner_ids[], type_id, priority, progress_type, progress_plan_type, progress_total, progress_unit, project_id, parent_id, desc, files_baseline, objectlinks[0][object], objectlinks[0][object_id], cf1, cf11, cf14, cf19, cf34, cf41, cf42, cf53, cf146, cf171, cf174, cf177, cf180, assign_other_off, no_allow_file_on_success, progress_required_desc, progress_required_file, inlineLogin. Không coi danh sách này là schema đầy đủ cho mọi task/user/trạng thái; checkbox/array có thể thay đổi trường được gửi.
- Request có Cookie (chỉ xác nhận sự hiện diện), Csrf-Token-Cache, Client-Instance-Id và X-Requested-With: XMLHttpRequest. Không sao chép giá trị cookie/token/client ID, không xuất HAR hoặc replay request cũ. Định dạng multipart boundary do trình duyệt sinh; không hardcode Content-Type/boundary khi gửi FormData.
- Tìm tên header trong Sources đang tải: packages.core-app.js?t=1791197954633 gắn `csrf-token-cache` từ `window.CSRF_TOKEN_CACHE`, `client-instance-id` từ `window.INSTANCE_ID` nếu có, `client-mode` từ `window.CLIENT_MODE` nếu có. Các đoạn core-app/plugin-pop/visuallayout đọc response header csrf-token-cache để cập nhật cache. Đây là mã static đã đọc, không trích giá trị global runtime. Chưa kết luận thiếu CSRF trong body là thiếu bảo vệ CSRF.
- Preview phản hồi POST: `{close: true, reload: true, notice: "Cập nhật thành công"}`. Thông báo thành công không thay thế đọc lại dữ liệu. Sau lưu, UI chi tiết TEST01.3 xác nhận No. 1.7. Chưa đối chiếu toàn bộ trường trước/sau lần lưu thủ công này, không tuyên bố mọi trường máy chủ không thay đổi.
- Đã đóng DevTools, chỉ giữ ảnh nguồn sạch `/tmp/tera-save-request-verification.png` có mã TEST01.3/No. 1.7, không có token/header. Không lưu screenshot Network chứa thông tin phiên vào file dự án.
- Kết luận: có bằng chứng lưu full-form một task bằng POST multipart, chưa có bằng chứng cập nhật riêng cf19 hoặc bulk. Chưa bật writer ngầm. Bước kỹ thuật cần xác minh tiếp là GET XHR lấy form mới có trả đủ schema/default values/arrays/checkbox và app_ctx để tái tạo chính xác payload hay không; không dùng inventory enrich thay thế toàn bộ FormData và không tái sử dụng token/cookie đã nhìn thấy.

## Bản 14: mở rộng đúng năm công việc test

- Người dùng yêu cầu thử thay đổi cả năm mã, không chốt No. đích cố định: TEST01.01/100113, TEST01.2/100116, TEST01.3/100119, TEST01.INSERT/100122, TEST01.4/100143. Giữ giới hạn origin/project 870, không mở toàn dự án hoặc production.
- Mở allowlist UI, wrapper và MAIN runner lên tối đa 5 thay đổi; đối chiếu cả ID và mã, không cho mã đúng nhưng ID khác. URL edit vẫn literal cho đúng 5 ID. Scope/journal có helper chung ngoài MAIN; MAIN giữ literal tự chứa để inject không phụ thuộc module imports.
- No. trước trong planner ưu tiên sourceNo gốc, kể cả chuỗi rỗng, thay vì outline hiển thị tự sinh. No. đích vẫn phải theo dạng số phân cấp hợp lệ, không trùng kết quả cuối. Dòng không đổi bị bỏ qua, không ép gửi lưu chỉ vì người dùng dán đủ 5 dòng.
- Đọc parent_id từ API mới của từng task, kiểm tra định dạng và giữ baseline; form project/parent phải khớp, post-read đổi parent thì dừng. Không giả định cả năm task cùng parent 100104, không thay đổi quan hệ dựa vào No. Chưa có server-side optimistic locking hoặc giao dịch nguyên tử.
- Mỗi ID có No. tạm riêng 99001.1..99001.5, tránh trùng các vòng đổi số. Scheduler vẫn ghi đích trống trước, dùng số tạm khi có vòng, không tự retry/rollback khi chưa rõ kết quả. Chưa bật POST ghi ngầm: chỉ check/recheck ngầm, lưu vẫn qua form đã kiểm tra.
- Nhật ký V1 đọc được dữ liệu cũ 1-2 task và mới tối đa 5 task, cho before trống và kiểm tra temp đúng allowlist. Recovery chỉ đọc đúng các ID thuộc journal, không mặc định hai task. Probe mở trực tiếp cả 5 form và cho phép No. trống, 0 save.
- Tests mới: UI cho ghi preview đủ năm mã; planner dùng blank thật; 5 thay đổi/5 save với blank và task có cha khác; vòng 4 task cộng task trống với 6 save, không trùng No. trung gian; recovery chọn 3 ID; journal 5 task/blank/temp riêng. 214 tests/25 files qua, typecheck/build/diff check qua. Chưa chạy live bản 14 và không gửi ghi trong lượt triển khai.
