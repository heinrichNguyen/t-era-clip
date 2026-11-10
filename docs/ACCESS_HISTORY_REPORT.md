# Báo cáo lịch sử truy cập, bản thử nghiệm

## Cách sử dụng

1. Tải lại extension sau khi cập nhật bản build.
2. Trên 1Office, mở **Lịch sử truy cập** bằng tài khoản có quyền xem nhật ký.
3. Mở T-eraClip, chọn **Báo cáo lịch sử truy cập**.
4. Chọn ngày chốt và thẻ công ty. Với TDK/TDU, chọn toàn nhóm hoặc một thành viên trong danh sách xổ xuống. Vị trí, phòng ban (tên/mã) là điều kiện bổ sung kết hợp đồng thời. Chọn Toàn bộ và để trống hai ô để quét toàn bộ. Bấm **Thu thập nhật ký**, giữ cả hai tab mở.
5. Chỉ sau khi đối chiếu đủ dữ liệu mới xuất Excel. Có thể hủy; không xuất báo cáo thu thập dở.
6. Sau khi thu thập thành công, đổi ngày chốt để tính lại ngay trên dữ liệu trong bộ nhớ, không tải lại nhật ký. Chọn **Thu thập lại từ 1Office** khi muốn làm mới dữ liệu. Đóng hoặc tải lại trang sẽ mất dữ liệu tạm.

## Quy tắc

- Chỉ xuất tài khoản có trong danh sách **active hiện tại** của 1Office, đối chiếu chính xác theo tài khoản/ID. Tài khoản trong danh sách khóa, không xác minh được hoặc tên đăng nhập trùng bị loại, dù có nhật ký lịch sử. Trạng thái này lấy khi thu thập/cập nhật nhân sự, không phải trạng thái tại ngày chốt quá khứ.
- Danh sách người dùng phải xác nhận đúng `tab=active`/`tab=block` trong phản hồi. Không đọc được danh sách hoặc không xác minh được trạng thái thì không xuất báo cáo, không mặc định mọi tài khoản trong nhật ký là active.
- Khi có bộ lọc, chỉ tìm nhật ký các tài khoản active khớp điều kiện. Khi không lọc nhân sự, vẫn đối chiếu nguồn nhật ký đầy đủ rồi loại bản ghi của tài khoản không còn active khỏi snapshot/xem trước/Excel, tính lại số lượng và thời gian cũ nhất. Đổi ngày chốt giữ phạm vi active đã xác minh. Muốn cập nhật trạng thái phải cập nhật nhân sự hoặc thu thập lại; muốn bổ sung tài khoản mới kích hoạt lại phải thu thập lại.
- Lấy thời điểm truy cập lớn nhất trước 00:00 ngày kế tiếp theo UTC+7, nhóm theo tài khoản nguồn, không nhóm theo họ tên.
- Số bản ghi không phải số phiên hay số lần đăng nhập. Hai ID khác nhau cùng thời điểm vẫn là hai bản ghi nguồn.
- Họ tên lấy từ bản ghi nguồn, chưa xác minh là tên tại thời điểm lịch sử. Không tái dựng phòng ban hoặc chức danh quá khứ.
- Người không có bản ghi trước ngày chốt không nằm trong kết quả. Không suy ra họ chưa từng truy cập, vì nhật ký có thể đã bị xóa, giới hạn thời gian lưu hoặc quyền xem.
- Chưa xác minh thời gian nguồn được 1Office hiển thị theo múi giờ nào; bản thử nghiệm dùng UTC+7, cần đối chiếu trước khi dùng báo cáo chính thức.
- Nguồn là toàn bộ nhật ký mà tài khoản đang đăng nhập được phép xem, hoặc phạm vi tài khoản khớp bộ lọc nhân sự chọn trong extension. Bộ lọc trang nguồn không áp dụng cho báo cáo.

## Cách thu thập và giới hạn

Đọc phản hồi qua đường dẫn phân trang `/apps/admin-reconciliation-histories/list?v=list&p=N&_json=1&inlineLogin=1` trong phiên đăng nhập hiện tại. Không suy đoán API chưa được xác minh, không đọc cookie/token, không thêm quyền Chrome. Tối đa sáu trang đồng thời trong cửa sổ trượt, thay trang đã đối chiếu bằng yêu cầu tiếp theo thay vì chờ cả nhóm. Phản hồi vẫn được đối chiếu theo thứ tự trang dù trả về lệch thứ tự. Giới hạn 20 giây cho một yêu cầu, 2.000 trang (100.000 bản ghi với 50 bản ghi/trang). Hủy ngăn yêu cầu tiếp theo, chờ các yêu cầu đang chạy kết thúc để không chồng nhiều lần thu thập.

Đã quan sát trực tiếp yêu cầu chuyển trang ngày 05/10/2026: phản hồi chứa `elem[name="table"]`, thuộc tính `data-props` chứa JSON với `object`, `headers`, `rows[].ID`, `rows[].vals` và `page` (`total`, `current`, `range: 50`, `rq: "p"`). Bộ đọc phân tích thuộc tính này bằng DOMParser và JSON.parse, kiểm tra schema, đủ dòng và ID duy nhất. Chỉ giữ tài khoản, tên, thời gian và ID, không giữ IP/trình duyệt. Không thực thi mã sự kiện trong phản hồi, không mở iframe hoặc chạy cả giao diện trang. Vẫn đọc được bảng HTML thật khi nguồn trả định dạng đó. Phản hồi không nhận diện được thì báo lỗi ngay, không coi là danh sách rỗng.

Nút **Xuất tất cả** đã được kiểm tra: yêu cầu `/apps/admin-reconciliation-histories/export?...&all-fields=1` trả thông báo tác vụ đang thực hiện, chờ thông báo hoàn thành. Trang tiến trình của tác vụ thử nghiệm đã hiển thị 81.156 bản ghi và 100%. Chưa xác minh thời gian hoàn thành, đường tải kết quả hoặc hợp đồng dữ liệu của tác vụ. Extension chưa tích hợp đường này. Việc bỏ iframe là tối ưu đã triển khai, không phải bằng chứng rằng 81.000 bản ghi đã được đọc trong 30 giây; vẫn cần đo thực tế và xác minh phương án lấy hàng loạt.

Đối chiếu phạm vi từng trang, số ID duy nhất và trang đầu trước/sau. Với thứ tự ID giảm dần, giữ mốc ID lớn nhất tại lúc bắt đầu, loại bản ghi mới lớn hơn mốc và chấp nhận phần trùng có nội dung nguyên vẹn khi trang bị đẩy xuống. Quét thêm trang nếu cần để đủ số ID ban đầu. Khi đối chiếu cuối, phần tăng tổng phải tương ứng số ID mới và nhóm đầu ban đầu phải còn nguyên. Giới hạn tăng 950 bản ghi trong một lần thu thập. Nếu tổng giảm, thứ tự thay đổi, bản ghi trùng bị sửa hoặc thiếu dữ liệu, dừng và không xuất. Khi nguồn không theo ID giảm dần, dùng kiểm tra nghiêm ngặt cũ, không chấp nhận tổng thay đổi.

Đây không phải snapshot giao dịch của máy chủ: thay đổi ở trang giữa mà không đổi tổng/trang đầu vẫn có thể không bị phát hiện. Không cam kết tốc độ hay thành công với hơn 80.000 bản ghi trước khi thử thực tế. Không được tự bỏ các bản ghi cũ hoặc dừng sau một số ngày, vì có thể bỏ sót người ít truy cập. Khi không lọc, vẫn phải đọc toàn bộ nhật ký. Khi lọc theo nhân sự, dùng truy vấn phía máy chủ đã xác minh như mô tả dưới đây.

## Bộ lọc trước khi quét (06/10/2026)

- Đã thử yêu cầu `range=1000`: nguồn vẫn trả 50 dòng/trang. Chưa có bằng chứng lấy được 1.000, 5.000 hoặc 10.000 dòng/lần. Không tăng song song vô hạn.
- Đã xác minh tìm kiếm `s=<tài khoản>` ở nguồn nhật ký: phản hồi xác nhận giá trị trong `page.params.s`/`vars.s`, tổng phân trang giảm về tập kết quả tìm kiếm. Extension chỉ đọc nhật ký của các tài khoản khớp bộ lọc nhân sự. Tìm kiếm là từ khóa, không giả định bằng tuyệt đối: đọc đủ kết quả rồi giữ chính xác tên tài khoản, loại các tài khoản khớp một phần hoặc trường khác. Không xác nhận được bộ lọc thì dừng, không âm thầm quét toàn tập đoàn.
- Tối đa sáu yêu cầu nhật ký đang chạy cho toàn bộ tác vụ, kể cả nhiều tài khoản và đối chiếu lại. Giữ đủ lịch sử còn được cấp quyền xem cho từng tài khoản; không dùng trường truy cập mới nhất thay cho lịch sử.
- Cây `/apps/admin-department/list?v=list&tab=department&_json=1&inlineLogin=1` đã đối chiếu đủ 460 nút. Ghép `vals.department_id` của người dùng với ID cơ cấu, kế thừa công ty gần nhất loại COMPANY/BRANCH. Mã lấy từ tiêu đề đơn vị nguồn, không suy từ chữ `R -`.
- Đơn vị mã **TDR**, ID 397, có tên nguồn **Công ty Cổ phần Bất động sản Tập đoàn Tân Á Đại Thành - Meyland**. Người dùng đã xác nhận đây là đơn vị họ gọi là Tổng công ty Meygroup. Giữ nguyên tên nguồn trong dữ liệu.
- Khi lọc, đọc nhân sự trước. Tài khoản thiếu dữ liệu của điều kiện lọc bị loại, có cảnh báo dữ liệu thiếu; không suy đoán. Không có tài khoản khớp thì không gọi nhật ký. Công ty chưa xác minh thì không quét thay thế toàn tập đoàn.
- Vị trí/phòng ban được chuyển từ HTML sang văn bản bằng parser không thực thi mã, gồm icon span và HTML mã hóa. Excel giữ tám cột cũ và bổ sung **Công ty hiện tại**; trường nhân sự được làm sạch lại trước khi xuất.
- Phạm vi lọc và giới hạn ghi trên trang, trong sheet thông tin. Mốc bắt đầu riêng từng tài khoản, không phải snapshot đồng thời toàn công ty. Đổi bộ lọc xóa kết quả tạm để không xuất nhầm phạm vi; đổi ngày chốt dùng lại phạm vi đã thu thập. Nút cập nhật nhân sự không mở rộng lại tập tài khoản nhật ký.
- Tốc độ phụ thuộc số tài khoản khớp và nhật ký của họ. Chưa đo thời gian toàn bộ TDR với bản build này, không cam kết 30 giây. Nếu gần như toàn tập đoàn khớp, tìm từng tài khoản có thể không nhanh hơn quét toàn bộ.

### Nhóm công ty (09/10/2026)

- Phạm vi theo xác nhận nghiệp vụ: TDK gồm nhân sự cấp TDK và TDH, TDP, TDC, TDW; TDU gồm nhân sự cấp TDU và TDU.CKKG, TDU.GNKG, TDU.VLXD, TDU.XMHT. TDG, TDT, TDR độc lập. Chọn riêng thành viên chỉ khớp mã công ty chính xác, không khớp một phần của mã hoặc tên.
- Thẻ chọn bố trí dạng lưới, danh sách thành viên chỉ hiện khi chọn TDK/TDU. Chỉ một phạm vi công ty mỗi lần; các điều kiện vị trí/phòng ban vẫn áp dụng đồng thời. Đổi lựa chọn xóa kết quả cũ để không xuất nhầm phạm vi.
- Cây 1Office có nhánh cha không hiển thị mã trên TDK/TDU. Nhận diện nhánh này bằng mã công ty con trực tiếp đã xác minh, giữ nhóm riêng (`companyGroup`), không tự gán mã nguồn hoặc ghép theo tên trùng. Excel vẫn giữ tên công ty nguồn; sheet thông tin báo cáo ghi rõ phạm vi toàn nhóm hoặc thành viên.
- Mọi phạm vi tiếp tục chỉ xuất tài khoản active đã xác minh. Chưa đối chiếu số nhân sự/nhật ký thực tế của từng nhóm sau thay đổi này.

## Thu thập thông tin yêu cầu mạng để tối ưu lần đầu

1. Mở **Lịch sử truy cập** trong Chrome, nhấn `Command + Option + I`.
2. Chọn tab **Network**, chọn **Fetch/XHR**. Nếu không thấy tab Network, tìm trong nút `>>`.
3. Bấm mũi tên sang trang 2 trên danh sách 1Office, không phải trong extension.
4. Chọn dòng yêu cầu vừa xuất hiện, liên quan `admin-reconciliation-histories` hoặc tải bảng nhật ký.
5. Ghi **Request URL**, **Request Method** và các mục trong **Query String Parameters / Payload**. Không sao chép Request Headers, Cookie, Authorization hoặc token.
6. Ở tab **Response**, sao chép phần phản hồi vào file `.txt`. Nếu có cookie, token hoặc mã phiên trong nội dung, loại bỏ trước khi gửi. Chỉ gửi trong kênh được công ty cho phép; dữ liệu có thể chứa thông tin truy cập cá nhân.

Không dùng Copy as cURL hoặc gửi toàn bộ HAR, vì có thể chứa thông tin xác thực. Thông tin này cần để xác minh cách 1Office lấy nhiều bản ghi/lọc ngày mà không mở hàng nghìn giao diện trang.

## Vị trí và phòng ban hiện tại

- Ngày 06/10/2026 đã đối chiếu giao diện hồ sơ người dùng và phản hồi `/app/autofill/users/items?ids[]=ID`: `items[].ID` cùng `placeholder` hiển thị phòng ban, rồi `||`, rồi vị trí. Đã thấy trường hợp chỉ có phòng ban, không có vị trí. Không dùng nhóm quyền hoặc chức vụ thay cho vị trí.
- Đọc `/apps/admin-user/list?v=list&tab=active|block&p=N&_json=1&inlineLogin=1` để lấy `rows[].ID` và `rows[].vals.username`. Ngày 06/10/2026 xác minh nguồn active có 1.046 tài khoản, block có 129 tài khoản (số liệu tại lúc kiểm tra, không cố định trong mã). Đọc cả hai danh sách để phát hiện trùng/mâu thuẫn, nhưng chỉ tra nhân sự và xuất nhật ký tài khoản active. Kiểm tra đúng tab, tổng, đủ dòng, ID trùng và trang đầu trước/sau. Không yêu cầu người dùng bật cột vị trí/phòng ban trong giao diện của họ; không sửa thiết lập hiển thị.
- Khi không lọc, sau khi quét xong nhật ký, chỉ tra thông tin cho tài khoản trong snapshot. Khi lọc, đọc nhân sự trước để chọn tài khoản. Ghép tên tài khoản chính xác sang ID người dùng, rồi tra theo lô tối đa 100 ID, tối đa bốn yêu cầu đồng thời. Không đọc hồ sơ từng người, không ghép bằng họ tên, không trả avatar/giới tính hay thông tin khác từ nguồn chọn người dùng về trang báo cáo.
- Thêm `Vị trí hiện tại` và `Phòng ban hiện tại` vào xem trước và Excel. Giữ nguyên toàn bộ cột cũ. Thông tin này không phải lịch sử điều chuyển tại ngày chốt. Ghi thời gian lấy nhân sự và cảnh báo vào sheet thông tin báo cáo.
- Tài khoản trùng, đã khóa hoặc không có trong danh sách active bị loại. Lỗi đọc danh sách trạng thái chặn xuất; lỗi nguồn thông tin vị trí/phòng ban chỉ để trống trường tương ứng và cảnh báo nếu đã xác minh active. Thiếu vị trí không được lấy phòng ban hoặc nhóm quyền điền thay.
- Kiểm tra nguồn theo lô ngày 06/10/2026 cho thấy một số tài khoản hoàn toàn không có `placeholder`. Đây là thiếu dữ liệu tùy chọn, không phải ID không hợp lệ: để trống riêng tài khoản active đó, không hủy toàn bộ kết quả. Giữ phòng ban hiển thị trong `rows[].fields.department_id` của danh sách khi tra vị trí thiếu dữ liệu. Không xuất `rows[].vals.department_id` vì giá trị này là ID số, không phải tên phòng ban.
- Nếu một lô tra cứu lỗi, giữ các lô thành công và phòng ban đã đọc từ danh sách, ghi cảnh báo và để trống vị trí chưa xác nhận. Nút **Cập nhật vị trí/phòng ban** kiểm tra lại trạng thái active và nhân sự trên snapshot có sẵn, không quét lại nhật ký. Loại người vừa khóa. Lỗi xác minh danh sách khi cập nhật sẽ ẩn báo cáo/nút xuất cũ để không tiếp tục xuất kết quả với trạng thái chưa xác nhận.
- Đổi ngày chốt dùng lại snapshot và thông tin nhân sự trong bộ nhớ, không gọi lại 1Office. Ngay cả khi nhật ký rỗng, vẫn xác minh danh sách trạng thái trước khi xuất.
- Cửa sổ trượt sáu trang là cải tiến về lịch gọi, không giảm hơn 1.600 yêu cầu nhật ký thành một yêu cầu hàng loạt. Chưa đo thời gian của bản này với toàn bộ 81.000 bản ghi, chưa cam kết đạt 30 giây. Không tăng song song vô hạn để ép tốc độ máy chủ.

## Kiểm thử thủ công cần làm

- Đối chiếu tài khoản mẫu với nhật ký tại ngày quá khứ, gồm bản ghi đúng 23:59:59 và sau ngày chốt.
- Kiểm tra nguồn trả HTML qua yêu cầu phân trang, có đủ ba tiêu đề và phạm vi tổng bản ghi. Nếu giao diện/API thay đổi, báo lỗi thay vì giả định rỗng.
- Kiểm tra tài khoản ít quyền, hết phiên đăng nhập, tab nguồn đóng/chuyển trang, hủy và nhật ký phát sinh bản ghi mới.
- Kiểm tra tài khoản có cùng họ tên và nhiều sự kiện cùng thời điểm.
- Kiểm tra Excel mở được, tiếng Việt, thời gian đúng; chỉ chia sẻ cho người được phép nhận thông tin nội bộ.

Không có backend, AI hoặc truyền dữ liệu đến nhà phát triển. Không ghi nhật ký vào storage của extension. Không phát hành gói thử nghiệm mới trong bước này.
