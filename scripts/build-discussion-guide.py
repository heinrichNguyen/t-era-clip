from pathlib import Path
from PIL import Image
from docx import Document
from docx.shared import Cm, Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'deliverables' / 'discussions-key-users'
IMAGES = OUT / 'images'
for source, target, box in [
    ('02-task-review.jpg', '02-task-detail.jpg', (380, 70, 1360, 800)),
    ('04-project-review.jpg', '04-project-header.jpg', (260, 120, 1510, 550)),
    ('04-project-review.jpg', '07-task-status.jpg', (260, 875, 1510, 1160)),
]:
    with Image.open(IMAGES / source) as im:
        im.crop(box).save(IMAGES / target, quality=95)

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.top_margin, sec.bottom_margin = Cm(1.7), Cm(1.7)
sec.left_margin, sec.right_margin = Cm(1.9), Cm(1.9)
sec.footer_distance = Cm(0.8)
for name in ['Normal', 'Title', 'Subtitle', 'Heading 1', 'Heading 2', 'Caption']:
    style = doc.styles[name]
    style.font.name = 'Arial'
    style.font.color.rgb = RGBColor.from_string('182536')
    style._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), 'Arial')
doc.styles['Normal'].font.size = Pt(10.5)
doc.styles['Normal'].paragraph_format.space_after = Pt(6)
doc.styles['Normal'].paragraph_format.line_spacing = 1.1
doc.styles['Title'].font.size = Pt(25)
doc.styles['Title'].font.color.rgb = RGBColor(0, 0, 0)
doc.styles['Heading 1'].font.size = Pt(18)
doc.styles['Heading 1'].paragraph_format.space_after = Pt(12)
doc.styles['Heading 2'].font.size = Pt(12)
doc.styles['Caption'].font.size = Pt(9)
doc.styles['Caption'].font.color.rgb = RGBColor.from_string('526174')
for style in doc.styles:
    for border in style.element.xpath('.//w:pBdr'):
        border.getparent().remove(border)

footer = sec.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
footer.add_run('T-eraClip Thảo luận | Key user | Trang ').font.size = Pt(8)
f = OxmlElement('w:fldSimple'); f.set(qn('w:instr'), 'PAGE'); footer._p.append(f)

def p(text, bold=False):
    line = doc.add_paragraph()
    line.add_run(text).bold = bold
    return line

def step(number, text):
    line = doc.add_paragraph()
    line.add_run(f'{number}. ').bold = True
    line.add_run(text)

def title(text):
    doc.add_page_break()
    doc.add_heading(text, level=1)

def image(name, width, caption):
    line = doc.add_paragraph()
    line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    line.add_run().add_picture(str(IMAGES / name), width=Inches(width))
    line.paragraph_format.space_after = Pt(3)
    note = doc.add_paragraph(caption, style='Caption')
    note.paragraph_format.space_after = Pt(10)

def table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Light Shading Accent 1'
    for c, text in zip(t.rows[0].cells, headers):
        c.text = text
        for run in c.paragraphs[0].runs: run.bold = True
    repeat = OxmlElement('w:tblHeader'); t.rows[0]._tr.get_or_add_trPr().append(repeat)
    for row in rows:
        cells = t.add_row().cells
        for c, text in zip(cells, row): c.text = text
    for row in t.rows:
        for i, c in enumerate(row.cells):
            if widths: c.width = Cm(widths[i])
            for para in c.paragraphs:
                para.paragraph_format.space_after = Pt(5)
                for run in para.runs: run.font.size = Pt(10)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)

doc.add_paragraph('Hướng dẫn thu thập thảo luận 1Office', style='Title')
doc.add_paragraph('T-eraClip Thảo luận dành cho key user', style='Subtitle')
p('Bản BETA 0.1.0 | Hướng dẫn ngày 03/11/2026', True)
p('Dùng tài liệu này để xuất bình luận và phản hồi của một công việc hoặc các công việc trong một dự án, phục vụ tổng hợp thông tin và đối chiếu trước các mốc báo cáo.')
p('Tiện ích chỉ thu thập và xuất file, không đổi No., mã công việc hay nội dung trên 1Office. Bản dành cho key user không có Capture, import hoặc công cụ điều chỉnh dữ liệu.')
image('01-menu.jpg', 2.8, 'Hình 1  Menu bản chỉ có thảo luận')
p('Ảnh trong tài liệu được chụp từ giao diện ứng dụng chạy với dữ liệu minh họa. Đây không phải ảnh dữ liệu công ty hay bằng chứng đã nghiệm thu tài khoản phòng ban.', False)

title('Chuẩn bị trước khi sử dụng')
p('Dùng tài khoản 1Office của chính bạn, chỉ thu thập thông tin thuộc phạm vi được phép. Việc cài được tiện ích không cấp thêm quyền xem dự án.')
doc.add_heading('Nhận và cài đúng bản', level=2)
step(1, 'Nhận gói hoặc link cài T-eraClip Thảo luận BETA từ đầu mối CNTT. Không dùng bản đầy đủ trong buổi hướng dẫn này.')
step(2, 'Nếu có link Chrome Web Store Private đã được CNTT phát hành: đăng nhập đúng Google Account được cấp quyền, mở link và xác nhận cài theo hướng dẫn Chrome. Bộ tài liệu hiện chưa có link Store đã được duyệt.')
step(3, 'Nếu CNTT tổ chức pilot bằng gói giải nén: mở chrome://extensions/, bật Developer mode nếu chính sách cho phép, bấm Load unpacked và chọn thư mục T-eraClip-Thao-luan chứa manifest.json. Nếu máy bị chặn, liên hệ CNTT; không tắt bảo vệ để vượt chặn.')
step(4, 'Mở menu Tiện ích của Chrome, ghim T-eraClip Thảo luận để dễ truy cập. Kiểm tra đúng tên có chữ Thảo luận.')
doc.add_heading('Chọn đúng nguồn dữ liệu', level=2)
table(['Bạn cần lấy gì', 'Mở trang nào và chọn nút nào'], [
    ['Một công việc', 'Mở chi tiết công việc, chọn Thu thập thảo luận trang hiện tại.'],
    ['Từng công việc trong dự án', 'Mở đúng dự án, chọn Thu thập thảo luận toàn dự án.'],
    ['Thảo luận chung ở cấp dự án', 'Mở khung thảo luận cấp dự án, chọn Thu thập thảo luận trang hiện tại. Không thay thế thu thập từng công việc.'],
])
step(5, 'Đăng nhập 1Office trước, mở công việc hoặc dự án đúng phạm vi, rồi mới bấm biểu tượng tiện ích. Kiểm tra tên nguồn trước khi thu thập.')
step(6, 'Các địa chỉ được cấu hình hiện nay: office.meygroup.vn, testoffice.meygroup.vn và testoffice.meygroup.vn:9443, đều qua HTTPS. Tenant khác chưa được hỗ trợ.')
p('Trong lúc chạy, giữ trang kết quả và tab nguồn mở. Không chuyển sang dự án khác hoặc thao tác chỉnh sửa trên tab nguồn. Bản dữ liệu là kết quả tại thời điểm thu thập, không tự cập nhật về sau.', True)

title('Thu thập một công việc')
step(1, 'Trên 1Office, mở chi tiết đúng công việc và bảo đảm khung thảo luận đã hiển thị. Không cần bấm Sửa hoặc Lưu.')
step(2, 'Bấm biểu tượng T-eraClip Thảo luận, kiểm tra dòng Nguồn, rồi bấm Thu thập thảo luận trang hiện tại.')
step(3, 'Giữ cửa sổ tiện ích mở cho đến khi có kết quả. Hệ thống có thể mở rộng các bình luận cũ; nếu đóng popup sớm, hãy thu thập lại khi cần.')
step(4, 'Đọc số bình luận gốc, số phản hồi và cảnh báo. Chưa đầy đủ không có nghĩa công việc không có thảo luận.')
step(5, 'Bấm Xem trước và xuất file. Kiểm tra tên công việc, thời điểm thu thập, thứ tự nội dung, tác giả và phản hồi trước khi xuất.')
image('02-task-detail.jpg', 5.9, 'Hình 2  Bản xem trước công việc và cấu trúc phản hồi với dữ liệu minh họa')

title('Xuất file của một công việc')
step(1, 'Trong bản xem trước, bấm Xuất Excel. Với nhu cầu trao đổi hoặc đọc báo cáo thông thường, ưu tiên Excel; JSON dành cho đối chiếu kỹ thuật khi CNTT yêu cầu.')
step(2, 'Đọc hộp xác nhận, kiểm tra tên file và phạm vi. Bấm Tải file Excel, sau đó kiểm tra thư mục tải xuống của Chrome.')
image('03-task-export.jpg', 5.8, 'Hình 3  Xác nhận xuất Excel cho một công việc')
table(['Sheet trong Excel', 'Nội dung cần kiểm tra'], [
    ['Đối tượng', 'Thông tin công việc hoặc đối tượng nguồn và thông tin lượt thu thập.'],
    ['Thảo luận', 'Bình luận, phản hồi, ID cha, tác giả, thời gian, nội dung và người được nhắc đến.'],
    ['Tệp đính kèm', 'Tên và metadata tệp thu thập được. Không phải các tệp gốc tải kèm.'],
])
step(3, 'Đối chiếu một bình luận đầu, một bình luận cuối và một phản hồi với nguồn. Nếu số lượng hoặc nội dung không khớp, chưa sử dụng file như báo cáo đầy đủ.')
step(4, 'Nếu bản thu thập thiếu, hệ thống yêu cầu xác nhận riêng trước khi tải phần dữ liệu hiện có. Chỉ xuất một phần khi bạn hiểu giới hạn và ghi rõ giới hạn đó khi gửi báo cáo.')
p('Luồng một công việc hiện không có bộ lọc ngày/người bình luận như luồng toàn dự án. Không suy ra rằng chọn Xuất Excel sẽ tự lọc theo mốc báo cáo.', True)

title('Thu thập toàn bộ công việc của dự án')
step(1, 'Mở đúng dự án trên 1Office. Bấm tiện ích và chọn Thu thập thảo luận toàn dự án.')
step(2, 'Đợi hệ thống kiểm tra danh sách. Chỉ danh sách đầy đủ và có công việc mới tự chuyển sang trang thu thập; nếu bị chặn, đọc cảnh báo, không coi dữ liệu một phần là toàn dự án.')
step(3, 'Giữ cả tab nguồn và tab kết quả mở. Hệ thống đọc theo nhóm nhỏ; khi cần dự phòng, tab nguồn có thể tự mở các công việc. Không tự đổi trang trong lúc chạy.')
image('04-project-header.jpg', 6.6, 'Hình 4  Tiến độ và các chỉ số của lượt thu thập dự án minh họa')
step(4, 'Theo dõi số đã xử lý/tổng số, thành công, thất bại, bình luận gốc và phản hồi. Công việc có 0 bình luận vẫn có thể được thu thập thành công.')
step(5, 'Muốn dừng: bấm Dừng thu thập khi nút đang hiện. Hệ thống dừng sau nhóm hiện tại; các công việc chưa xử lý không được xem là đã kiểm tra.')
image('07-task-status.jpg', 6.6, 'Hình 5  Trạng thái từng công việc, không phải trạng thái tiến độ nghiệp vụ')
p('Thu thập thành công hoặc Hoàn thành trong bảng này nói về lượt đọc thảo luận, không xác nhận công việc hay milestone đã hoàn thành.', True)

title('Lọc thảo luận theo mốc báo cáo')
p('Bộ lọc nằm ở trang toàn dự án sau khi thu thập. Công cụ không tự nhận diện hay đánh giá milestone; key user xác định khoảng ngày cần đối chiếu theo mốc được giao.')
image('05-filters.jpg', 6.6, 'Hình 6  Bộ lọc trước khi xuất Excel toàn dự án')
table(['Trường lọc', 'Cách sử dụng'], [
    ['Từ ngày và Đến ngày', 'Chọn khoảng ngày bình luận. Tính theo giờ Việt Nam UTC+7, gồm cả ngày kết thúc. Ngày bắt đầu không được sau ngày kết thúc.'],
    ['Người bình luận', 'Lọc theo tên tác giả, không phải người thực hiện công việc. Nếu trùng tên, kết quả gồm các bản ghi cùng tên.'],
    ['Loại thảo luận', 'Chọn Tất cả, Bình luận gốc hoặc Phản hồi.'],
    ['Chỉ lấy mới nhất', 'Lấy bản ghi mới nhất trong phần đã khớp các bộ lọc của từng công việc. Giữ các bản ghi có cùng thời điểm mới nhất. Không phải bình luận mới nhất duy nhất của toàn dự án.'],
    ['Xóa bộ lọc', 'Trở về toàn bộ nội dung đã thu thập, không thu thập lại từ 1Office.'],
])
doc.add_heading('Ví dụ áp dụng', level=2)
p('Cần tổng hợp trước cuộc họp ngày 03/11: chọn từ 01/11 đến 03/11 theo phạm vi báo cáo được giao; giữ Tất cả nếu cần cả yêu cầu và phản hồi. Kiểm tra số bình luận và số công việc khớp trước khi tải.')
p('Bộ lọc chỉ áp dụng sheet Thảo luận và Tệp đính kèm. Sheet Công việc vẫn giữ toàn bộ danh sách. Bản ghi không có thời gian xác định bị loại khi lọc ngày hoặc chọn mới nhất; hãy đọc cảnh báo.', True)

title('Xuất và đọc kết quả toàn dự án')
step(1, 'Khi lượt thu thập kết thúc, bấm Xuất Excel. Kiểm tra số bình luận/phản hồi và số công việc khớp bộ lọc trong hộp xác nhận.')
image('06-project-export.jpg', 5.8, 'Hình 7  Xác nhận phạm vi xuất toàn dự án với dữ liệu minh họa')
step(2, 'Nếu có công việc thiếu, thất bại hoặc bị hủy, cần xác nhận xuất phần hiện có. Không tích xác nhận nếu bạn chưa hiểu phần dữ liệu còn thiếu; dùng Thử lại thu thập hoặc báo CNTT.')
step(3, 'Bấm Tải file Excel. Kiểm tra file trong thư mục tải xuống, mở file và rà soát các sheet dưới đây.')
table(['Sheet', 'Ý nghĩa'], [
    ['Dự án', 'Thông tin dự án, lượt thu thập, phạm vi và bộ lọc xuất.'],
    ['Công việc', 'Toàn bộ danh sách công việc đã lấy được, trạng thái thu thập và lỗi/thiếu nếu có.'],
    ['Thảo luận', 'Bình luận và phản hồi khớp bộ lọc, gắn với từng công việc.'],
    ['Tệp đính kèm', 'Metadata tệp của phần thảo luận khớp bộ lọc; không có tệp gốc tải kèm.'],
])
p('Mã và No. xuất hiện trong Excel chỉ là thông tin nhận diện. Tiện ích không cập nhật lại hai trường này. Không dùng file thảo luận làm file import điều chỉnh dự án.', True)

title('Xử lý tình huống và bảo vệ dữ liệu')
table(['Tình huống', 'Cách xử lý'], [
    ['Nút toàn dự án bị mờ', 'Mở trang chi tiết dự án thay vì trang công việc, trang chủ hoặc danh sách chung; sau đó mở lại tiện ích.'],
    ['Không có dữ liệu hoặc không thấy khung thảo luận', 'Kiểm tra đăng nhập, quyền xem, đúng tab và khung thảo luận đã tải. Không kết luận dự án không có thảo luận từ một kết quả rỗng.'],
    ['Đọc lỗi hoặc phiên hết hạn', 'Đọc thông báo, đăng nhập lại nếu cần và bắt đầu lại. Không cố vượt hạn chế quyền.'],
    ['Có công việc thất bại hoặc chưa đầy đủ', 'Xem từng dòng lỗi. Thử lại khi nguồn ổn định; chỉ xuất một phần nếu báo cáo ghi rõ giới hạn. Thử lại có thể chạy lại lượt thu thập.'],
    ['Mất dữ liệu sau reload hoặc cập nhật tiện ích', 'Dữ liệu xem trước dùng phiên tạm. Thu thập lại từ đúng nguồn. File đã tải không tự bị xóa.'],
])
doc.add_heading('Sau khi sử dụng', level=2)
step(1, 'Lưu file vào nơi nội bộ được cho phép. Kiểm tra nội dung trước khi gửi; Excel có thể chứa tên người, trao đổi và thông tin dự án nhạy cảm.')
step(2, 'Bấm Xóa dữ liệu tạm và xác nhận khi không cần giữ bản xem trước. Thao tác này không xóa file đã tải, không xóa thảo luận trên 1Office.')
step(3, 'Không đưa file hoặc ảnh có nội dung công ty lên công cụ AI/dịch vụ ngoài nếu chưa có phê duyệt. Không chia sẻ cookie, token hoặc mật khẩu để nhờ hỗ trợ.')
doc.add_heading('Gửi lỗi cho đầu mối CNTT', level=2)
p('Cung cấp: tên phiên bản tiện ích; Chrome và hệ điều hành; thời điểm lỗi; phạm vi công việc/dự án; bước đã làm; thông báo lỗi; số dự kiến và số đã đọc. Chỉ gửi link/ảnh/nội dung trong phạm vi được phép và che thông tin nhạy cảm khi cần.')
p('Trước khi dùng chính thức, key user cùng CNTT nghiệm thu trên tài khoản phòng ban thực tế. Tài liệu này không thay bằng chứng kiểm thử quyền truy cập hoặc sự chấp thuận của Google.', True)

doc.core_properties.title = 'Hướng dẫn thu thập thảo luận 1Office'
doc.core_properties.subject = 'T-eraClip Thảo luận cho key user'
doc.core_properties.author = 'CNTT'
doc.save(OUT / 'Huong-dan-key-user-T-eraClip-Thao-luan.docx')
print(OUT / 'Huong-dan-key-user-T-eraClip-Thao-luan.docx')
