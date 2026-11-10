from pathlib import Path
from math import atan2, cos, sin, pi
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor, white
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'deliverables' / 'discussions-key-users'
IMG = OUT / 'images'
PDF = OUT / 'Huong-dan-truc-quan-T-eraClip-Thao-luan.pdf'
pdfmetrics.registerFont(TTFont('UI', '/System/Library/Fonts/Supplemental/Arial.ttf'))
pdfmetrics.registerFont(TTFont('UIB', '/System/Library/Fonts/Supplemental/Arial Bold.ttf'))
W, H = 960, 540
C = canvas.Canvas(str(PDF), pagesize=(W, H))
C.setTitle('T-eraClip Thảo luận | Hướng dẫn bằng hình ảnh')
C.setAuthor('CNTT')
NAVY, MUTED, BLUE, RED = map(HexColor, ['#14263E', '#596B80', '#1745BE', '#D83638'])

def text(s, x, y, size=15, bold=False, color=NAVY):
    C.setFont('UIB' if bold else 'UI', size)
    C.setFillColor(color)
    C.drawString(x, H-y-size, s)

def para(s, x, y, w, size=16, bold=False, color=NAVY):
    style = ParagraphStyle('p', fontName='UIB' if bold else 'UI', fontSize=size,
                           leading=size*1.35, textColor=color)
    p = Paragraph(s, style)
    _, h = p.wrap(w, 1000)
    p.drawOn(C, x, H-y-h)
    return h

def rect(x, y, w, h, fill=None, stroke=None, radius=0, line=1):
    C.setLineWidth(line)
    C.setFillColor(fill or white)
    C.setStrokeColor(stroke or white)
    C.roundRect(x, H-y-h, w, h, radius, fill=bool(fill), stroke=bool(stroke))

def page(n, section, heading, hint):
    rect(0, 0, W, H, HexColor('#F8FAFD'))
    rect(0, 0, 9, H, BLUE)
    text(f'{section}  /  {n:02d}', 40, 16, 10, True, BLUE)
    text(heading, 40, 37, 28, True)
    text(hint, 40, 75, 13, color=MUTED)
    C.setStrokeColor(HexColor('#DCE3ED')); C.setLineWidth(.7)
    C.line(40, 30, 920, 30)
    text('T-eraClip Thảo luận BETA 0.1.0 | 03/11/2026', 40, 517, 9, color=MUTED)
    text('Ảnh giao diện với dữ liệu minh họa', 590, 517, 9, color=MUTED)
    text(f'{n}/8', 895, 517, 9, True, MUTED)

def photo(name, x, y, w, h=None, crop=None):
    im = Image.open(IMG/name)
    if crop: im = im.crop(crop)
    iw, ih = im.size
    if h is None: h = w*ih/iw
    assert abs(w/h-iw/ih) < .01, 'Do not distort screenshots'
    rect(x-3, y-3, w+6, h+6, white, HexColor('#D6E0EF'), 7)
    C.drawImage(ImageReader(im), x, H-y-h, w, h)
    return lambda bx, by, bw, bh: (x+bx*w/iw, y+by*h/ih, bw*w/iw, bh*h/ih)

def arrow(x1, y1, x2, y2):
    C.setStrokeColor(RED); C.setFillColor(RED); C.setLineWidth(1.7)
    C.line(x1, H-y1, x2, H-y2)
    a = atan2(y2-y1, x2-x1)
    path = C.beginPath(); path.moveTo(x2, H-y2)
    for delta in [-.45, .45]:
        path.lineTo(x2-9*cos(a+delta), H-(y2-9*sin(a+delta)))
    path.close(); C.drawPath(path, fill=1, stroke=0)

def badge(n, x, y):
    C.setFillColor(RED); C.circle(x, H-y, 12, fill=1, stroke=0)
    C.setFillColor(white); C.setFont('UIB', 12)
    C.drawCentredString(x, H-y-4, str(n))

def target(coords, n=None, badge_at=None):
    x, y, w, h = coords
    rect(x-2, y-2, w+4, h+4, stroke=RED, radius=4, line=2)
    if n is not None: badge(n, *(badge_at or (x-10, y-10)))
    return x, y, w, h

def note(s, y=470, warn=False):
    rect(40, y, 880, 33, HexColor('#FFF1DE' if warn else '#EAF1FF'), radius=6)
    text(s, 54, y+8, 12, color=HexColor('#8D4212') if warn else BLUE)

def label(n, heading, detail, x, y, w=320):
    badge(n, x+12, y+13)
    heading_height = para(heading, x+37, y, w-37, 19, True)
    if detail: para(detail, x+37, y+heading_height+10, w-37, 14, color=MUTED)

page(1, 'MỘT CÔNG VIỆC', 'Bắt đầu từ công việc cần lấy thảo luận',
     'Đã được CNTT cài tiện ích: đăng nhập 1Office, mở đúng công việc, rồi bấm biểu tượng T-eraClip.')
f = photo('01-menu.jpg', 560, 100, 265)
target(f(24, 287, 360, 49), 1, (535, 293))
label(1, 'Bấm “Thu thập thảo luận<br/>trang hiện tại”',
      'Giữ popup mở và đợi kết quả.', 55, 180, 420)
arrow(435, 285, 570, 293)
label(2, 'Chọn “Xem trước và xuất file”',
      'Nút xuất hiện sau khi đã thu thập.<br/>Cuộn xuống popup nếu cần.', 55, 345, 420)
text('Menu minh họa; thao tác từ trang công việc thực tế của bạn.', 48, 487, 11, color=MUTED)
C.showPage()

page(2, 'MỘT CÔNG VIỆC', 'Kiểm tra kết quả trước khi xuất',
     'Đúng công việc, có bình luận/phản hồi cần lấy, không bỏ qua cảnh báo thiếu.')
f = photo('02-task-detail.jpg', 45, 107, 510)
target(f(26, 180, 560, 64), 1, (565, 230))
target(f(25, 267, 931, 91), 2, (565, 325))
target(f(720, 84, 103, 44), 3, (565, 162))
label(3, 'Bấm “Xuất Excel”', '', 620, 145, 300)
arrow(615, 180, 565, 162)
label(1, 'Tên công việc', 'Đối chiếu với trang nguồn.', 620, 225, 300)
label(2, 'Số bình luận và trạng thái', 'Chưa đầy đủ: kiểm tra cảnh báo.', 620, 320, 300)
text('Nội dung và quan hệ phản hồi ở phần dưới của bản xem trước.', 45, 496, 11, color=MUTED)
C.showPage()

page(3, 'MỘT CÔNG VIỆC', 'Xác nhận và tải file Excel',
     'Đọc hộp xác nhận, bấm tải rồi mở thư mục tải xuống của Chrome.')
f = photo('03-task-export.jpg', 145, 112, 670)
target(f(398, 200, 118, 42), 1, (838, 379))
arrow(837, 391, 767, 400)
text('1  Tải file Excel', 664, 455, 18, True, BLUE)
text('3 sheet: Đối tượng  /  Thảo luận  /  Tệp đính kèm', 48, 468, 14, True)
text('Tệp đính kèm chỉ là thông tin tệp, không tải kèm tệp gốc.', 48, 491, 12, color=MUTED)
C.showPage()

page(4, 'TOÀN DỰ ÁN', 'Mở dự án và chọn thu thập toàn dự án',
     'Mở đúng trang dự án trên 1Office trước khi mở tiện ích.')
f = photo('01-menu.jpg', 560, 100, 265)
target(f(24, 389, 360, 49), 1, (535, 355))
label(1, 'Bấm “Thu thập thảo luận<br/>toàn dự án”',
      'Đợi kiểm tra danh sách.<br/>Trang kết quả sẽ tự mở.', 55, 210, 425)
arrow(440, 320, 572, 356)
para('Giữ tab nguồn và tab kết quả mở.<br/>Không đổi sang dự án khác khi đang chạy.', 92, 379, 350, 15)
text('Nút “trang hiện tại” ở trang dự án chỉ lấy thảo luận cấp dự án, không lấy từng công việc.', 48, 489, 11, color=MUTED)
C.showPage()

page(5, 'TOÀN DỰ ÁN', 'Theo dõi tiến độ và trạng thái từng công việc',
     'Chờ lượt thu thập kết thúc. Cần dừng: bấm “Dừng thu thập” khi nút đang hiện.')
f = photo('04-project-header.jpg', 75, 104, 810, crop=(0,140,1250,430))
target(f(36, 108, 1195, 99), 1, (902, 215))
g = photo('07-task-status.jpg', 75, 314, 810, crop=(0,0,1250,270))
target(g(765, 137, 191, 130), 2, (902, 430))
text('1  Tiến độ đọc thảo luận', 75, 297, 12, True, BLUE)
text('2  “Hoàn thành” là thu thập xong, không phải công việc hoàn thành.', 75, 495, 11, True, BLUE)
C.showPage()

page(6, 'TOÀN DỰ ÁN', 'Lọc phần thảo luận cần báo cáo',
     'Mốc báo cáo do bạn chọn. Công cụ không tự đánh giá milestone.')
f = photo('05-filters.jpg', 50, 183, 860)
target(f(24, 97, 578, 42), 1, (59, 226))
target(f(624, 97, 283, 42), 2, (584, 226))
target(f(922, 97, 285, 42), 3, (894, 226))
target(f(27, 154, 20, 22), 4, (59, 307))
text('1  Từ ngày, đến ngày', 50, 124, 16, True)
text('2  Người bình luận', 370, 124, 16, True)
text('3  Bình luận gốc / phản hồi', 660, 124, 16, True)
arrow(160, 153, 190, 246)
arrow(490, 153, 580, 246)
arrow(790, 153, 802, 246)
text('4  Tùy chọn: chỉ lấy mới nhất trong phần khớp bộ lọc của từng công việc.', 50, 429, 15, True)
text('Ngày tính UTC+7, gồm cả ngày kết thúc. Trùng tên tác giả: có thể gồm nhiều người.', 50, 460, 12, color=MUTED)
text('Sheet Công việc vẫn giữ toàn bộ danh sách. Đọc cảnh báo nếu thiếu thời gian bình luận.', 50, 485, 12, color=MUTED)
C.showPage()

page(7, 'TOÀN DỰ ÁN', 'Xuất đúng phạm vi đã chọn',
     'Bấm “Xuất Excel”, kiểm tra số bản ghi và công việc trong hộp xác nhận rồi tải file.')
f = photo('06-project-export.jpg', 190, 103, 580)
target(f(36, 177, 466, 39), 1, (209, 287))
target(f(388, 240, 119, 43), 2, (793, 397))
para('1  Kiểm tra<br/>phạm vi', 46, 295, 125, 14, True, BLUE)
arrow(176, 336, 230, 335)
text('2  Tải file Excel', 774, 423, 14, True, BLUE)
arrow(805, 417, 738, 406)
note('Có lỗi, thiếu hoặc đã dừng: chỉ xác nhận xuất một phần khi hiểu dữ liệu còn thiếu.', 465, True)
text('4 sheet: Dự án  /  Công việc  /  Thảo luận  /  Tệp đính kèm (không kèm tệp gốc).', 48, 500, 10, color=MUTED)
C.showPage()

page(8, 'SAU KHI SỬ DỤNG', 'Lưu file nội bộ và dọn dữ liệu tạm',
     'Kiểm tra file đã tải, sau đó bấm “Xóa dữ liệu tạm” và xác nhận.')
f = photo('01-menu.jpg', 145, 142, 670, crop=(0,500,430,640))
target(f(18, 74, 150, 42), 1, (119, 282))
arrow(120, 298, 173, 302)
rect(45, 394, 425, 76, HexColor('#EAF1FF'), radius=8)
text('Không xóa thảo luận trên 1Office.', 60, 407, 17, True, BLUE)
text('Không xóa file đã tải xuống.', 60, 438, 15, color=BLUE)
rect(490, 394, 425, 76, HexColor('#FFF1DE'), radius=8)
text('Chỉ chia sẻ file đúng phạm vi.', 505, 407, 17, True, HexColor('#8D4212'))
text('Có lỗi: gửi ảnh và bước thao tác cho CNTT.', 505, 438, 14, color=HexColor('#8D4212'))
text('Không gửi mật khẩu, cookie hoặc token trong yêu cầu hỗ trợ.', 48, 489, 12, color=MUTED)
C.showPage()

C.save()
print(PDF)
