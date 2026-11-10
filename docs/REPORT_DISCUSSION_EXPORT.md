# Thu thập thảo luận từ báo cáo đã lọc

## Phạm vi

- Nguồn: trang `/apps/report-report/view?ID=...` trên các môi trường 1Office đã cho phép.
- Báo cáo phải có bảng với hai cột `Mã công việc` và `Tên công việc`.
- Đọc danh sách đã dựng trên báo cáo, không mở rộng danh sách sang toàn bộ dự án, không ghi dữ liệu lên 1Office.
- Giữ nguyên luồng công việc đơn/toàn dự án. Bổ sung nút báo cáo ở cả bản đầy đủ và bản chỉ có thảo luận.

## Dùng thử

1. Reload tiện ích đang nạp từ `dist` hoặc `dist-discussions` sau khi build.
2. Mở báo cáo, chọn ngày và phòng ban, chờ bảng hiển thị xong.
3. Lượt pilot đầu nên chọn phạm vi nhỏ, kiểm tra trước số công việc trong bảng.
4. Mở tiện ích, bấm **Thu thập thảo luận danh sách trong báo cáo**.
5. Popup chỉ đọc các dòng đang có trong báo cáo theo phòng ban và thời gian đã lọc, sau đó mở trang tiến độ riêng. Không tải lại toàn dự án hoặc tự đổi ý nghĩa ngày lọc của báo cáo.
6. Giữ tab báo cáo mở và không đổi bộ lọc. Trang tiến độ gửi request ngầm qua tab nguồn, không chuyển trang qua từng công việc. Theo dõi tiến độ, xem cảnh báo, lọc thảo luận và xuất Excel như trước.
7. Đối chiếu một mẫu có nhiều bình luận/phản hồi/tệp với 1Office trước khi mở rộng phạm vi.

Ngày trong báo cáo chọn **danh sách công việc**. Mặc định lấy toàn bộ thảo luận của danh sách đó. Muốn giới hạn ngày bình luận, dùng bộ lọc thảo luận ở bước xem trước như trước đây.

## Đối chiếu ID

- Ưu tiên liên kết công việc có sẵn trong cột tên. Có thể đọc cột `ID công việc` rõ nghĩa, không dùng cột `ID` chung của báo cáo.
- Khi chưa có liên kết, gửi request tra mã đầy đủ ở danh sách công việc `all_all`. Truy vấn này chỉ nối ID của dòng đã thuộc báo cáo, không bổ sung công việc vào phạm vi xuất.
- Chỉ chọn khi đã đọc đủ các dòng kết quả tìm kiếm và có đúng một liên kết khớp tên sau chuẩn hóa khoảng trắng, sau đó đọc phản hồi chi tiết bằng GET để xác nhận mã đầy đủ. Có thể tìm kiếm trả cả cha/con cùng tiền tố mã; không tự chọn dòng đầu tiên. Không dựng trang, chạy callback của 1Office hoặc điều hướng tab.
- Không lấy số cuối của mã làm ID, không tự chọn một trong nhiều ứng viên.
- Thiếu mã/ID, mã trùng, không có quyền xem hoặc mã/tên không khớp: giữ nguyên dòng để xuất, đánh dấu chưa lấy được thảo luận và tiếp tục các dòng khác. Không tự ghép tên hoặc tạo ID giả.
- Nguồn đổi trang/bộ lọc giữa chừng hoặc cấu trúc bảng không hợp lệ: dừng đối chiếu để không vượt phạm vi đã chọn.
- Mỗi dòng báo cáo có khóa nội bộ riêng để các dòng cùng thiếu ID không bị gộp. Khóa này không xuất vào cột ID và không gửi tới API 1Office.
- Tối đa 3 công việc đồng thời. Không có nhánh dự phòng mở trang công việc cho nguồn báo cáo: phản hồi thiếu hoặc không đọc được sẽ giữ dòng và báo lỗi. Luồng công việc đơn/dự án không thay đổi.

## File Excel

- Giữ 4 sheet: Dự án, Công việc, Thảo luận, Tệp đính kèm.
- Giữ các cột công việc, thảo luận, tệp hiện có và cách định dạng cũ.
- Dòng chưa xác minh ID vẫn có trong sheet Công việc, ID để trống, trạng thái `Chưa lấy được thảo luận` và lý do ở cột Lỗi. Không có bản ghi thảo luận giả, không kết luận công việc không có bình luận. Cho phép xuất kể cả khi toàn bộ dòng đều chưa xác minh ID, sau khi người dùng xác nhận giới hạn dữ liệu trong hộp xuất hiện có.
- Sheet Dự án bổ sung loại nguồn, ID báo cáo, tên và URL chứa bộ lọc. Không dùng ID báo cáo thay ID dự án.
- Với báo cáo nhiều dự án, không suy cây cha/con từ No. Tên cha lấy từ báo cáo nếu có; thông tin không xác minh được để trống.
- Thảo luận công việc thường dùng vùng `work-normal-normal`, công việc còn lại dùng `work-task-task`. Không lấy nội dung tóm tắt ở cột Thảo luận của báo cáo thay cho dữ liệu bộ thu thập.
- Chỉ xuất thông tin tệp như luồng cũ, không tải hàng loạt tệp gốc.
- Theo xác nhận mở rộng của người dùng ngày 06/11, chỉ với mã thuần dạng `.chữ_số` hoặc `0.chữ_số`, được đối chiếu cách thể hiện có/không số 0 đầu và số 0 cuối phần thập phân (ví dụ `0.2175` với `.21750`). Chuẩn hóa bằng chuỗi, không chuyển mã thành số. Không áp dụng cho mã có tiền tố, mã phân cấp hay số 0 ở đầu phần thập phân (`.02175` không tương đương `.2175`). Chỉ thử truy vấn dạng khác khi truy vấn trước trả 0 kết quả, tối đa 3 dạng. Vẫn cần duy nhất một ứng viên khớp tên trong danh sách tìm kiếm đã đọc đủ và xác minh ID/mã chi tiết. Excel giữ mã báo cáo; cảnh báo ghi rõ hai mã ở phần Cảnh báo của sheet Dự án.

## Kiểm chứng và giới hạn

- Đã kiểm tra trực tiếp bảng báo cáo, phản hồi tải báo cáo và một mẫu tìm mã -> liên kết ID -> mã chi tiết -> vùng thảo luận.
- Đã thêm test cho nguồn báo cáo, tra ID bằng GET, không thay URL nguồn, không điều hướng dự phòng, phân biệt loại công việc, giao diện bản thảo luận và cấu trúc Excel.
- Bộ đọc request hỗ trợ HTML và JSON chứa nội dung markup khai báo. Phản hồi ngoài cấu trúc nhận diện được sẽ bị từ chối, không suy diễn ID. Chưa xác minh trực tiếp toàn bộ dạng phản hồi tra ID trên môi trường thật.
- Kiểm tra Network ngày 05/11 xác nhận bảng danh sách nằm trong `template children` lồng nhau. Bộ đọc đã bổ sung duyệt `template.content` trên tài liệu tách rời, không thực thi script. Test tái hiện lỗi này đã chạy đạt; chưa xác nhận lượt thu thập thật sau sửa.
- File kết quả 99 dòng ngày 05/11 xác nhận tất cả cùng lỗi khớp ứng viên ID. Phản hồi danh sách thực tế có `vals.title` là tên dạng chữ, `fields.title` chứa liên kết khai báo với nhãn trong `data-children`. Bộ đọc đã sửa theo cấu trúc này, đồng thời kiểm tra ID và các liên kết khớp nhau.
- Với trang chi tiết tải chậm theo khối, đọc khung `board` đúng ID, lấy danh sách khối của tab đang chọn rồi xác minh `Work-Task-TaskNormalInformation`, ID công việc, ID khối và mã. Không cố định ID khối theo một máy chủ. Vị trí khối đã xác minh được giữ trong bộ nhớ trang tiến độ để giảm request cho các dòng tiếp theo, nhưng mã và ID vẫn phải kiểm tra lại từng dòng.
- Đã kiểm thử lại bằng phản hồi danh sách và khung chi tiết lưu từ Chrome cho mã `KHTD.01.06.18526`, nối được ID `89897`. Phần phản hồi khối trong lượt kiểm thử lại là mẫu tối giản theo cấu trúc vừa quan sát trực tiếp, không phải toàn bộ phản hồi gốc. Chưa nghiệm thu lượt xuất 99 dòng bằng bản sửa mới.
- Giao diện phân biệt xử lý xong danh sách với lấy đủ thảo luận. Các dòng chưa xác minh ID được tính vào số chưa lấy được, không báo hoàn tất thành công khi toàn bộ dòng đều chưa lấy được.
- Chưa nghiệm thu một lượt xuất hoàn chỉnh từ báo cáo thật bằng extension mới; chưa kiểm thử quyền hạn chế/hết phiên và hiệu năng danh sách lớn.
- Danh sách đọc tại một thời điểm và từng công việc đọc sau đó không phải ảnh chụp nguyên tử.
- Không bổ sung quyền extension, không gửi dữ liệu sang dịch vụ thứ ba, không thay đổi thiết lập hoặc báo cáo trên 1Office.

### Rà soát file kết quả ngày 06/11

- File người dùng gửi có 45 dòng: 30 hoàn thành, 15 chưa xác minh ID. Không đồng nhất trạng thái hoàn thành với việc công việc phải có bình luận.
- Đã kiểm tra trực tiếp cả 15 trường hợp trên nguồn 1Office. Có 7 dòng liên kết dùng nhãn trong `template children`, 3 dòng tìm kiếm trả 3/7/4 kết quả cùng tiền tố nhưng chỉ một ứng viên đúng tên và mã, 4 dòng lệch `0.` với `.` như quy tắc đã duyệt.
- Dòng `Test QT gia hạn - Hạnh`: báo cáo `0.2175`, ID `99981`, mã chi tiết `.21750`. Ban đầu bị chặn ở chênh lệch chữ số cuối; sau xác nhận mở rộng nêu trên, được đối chiếu theo quy tắc dạng mã thập phân và ghi cảnh báo, không dùng ngoại lệ ID cố định.
- Đã mở trực tiếp công việc ID `99981`: đường dẫn thuộc `work-normal-normal`, nhóm công việc hiển thị `Ngoài dự án`, trang chi tiết không hiển thị liên kết dự án. Không suy ra mã hay ID từ tổng số công việc.
- Các sửa đổi đã bổ sung test cho nhãn template, nhiều kết quả, tên trùng, phân trang thiếu, cảnh báo dạng mã và từ chối tương đương số học. Chưa nghiệm thu một lượt xuất đủ 45 dòng sau các sửa đổi này.
