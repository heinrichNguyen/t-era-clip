import writeExcelFile, { type CellObject, type Sheet } from 'write-excel-file/browser';
import type { AccessReport } from '../accessHistory';
import { personnelText } from '../personnelText';

export function buildAccessWorkbook(report: AccessReport): Sheet<Blob>[] {
  const text = (value: string): CellObject => ({ value, type: String, format: '@', wrap: true });
  const header = (value: string): CellObject => ({ ...text(value), fontWeight: 'bold', backgroundColor: '#243EAD', textColor: '#FFFFFF' });
  const personnel = new Map(report.personnel?.users.map((user) => [user.account, { ...user,
    position: personnelText(user.position), department: personnelText(user.department), company: personnelText(user.company ?? '') }]) ?? []);
  return [{
    sheet: 'Tổng hợp truy cập', stickyRowsCount: 1,
    columns: [8, 28, 32, 26, 18, 24, 36, 44, 54].map((width) => ({ width })),
    data: [ ['STT', 'Tài khoản', 'Họ và tên tại bản ghi', 'Truy cập cuối trước ngày chốt', 'Ngày chốt', 'Số bản ghi đến ngày chốt', 'Vị trí hiện tại', 'Phòng ban hiện tại', 'Công ty hiện tại'].map(header),
      ...report.users.map((user, index) => [ { value: index + 1 }, text(user.account), text(user.name), text(user.timeRaw), text(report.cutoff), { value: user.eventCount }, text(personnel.get(user.account)?.position ?? ''), text(personnel.get(user.account)?.department ?? ''), text(personnel.get(user.account)?.company ?? '') ]) ],
  }, {
    sheet: 'Thông tin báo cáo', columns: [{ width: 30 }, { width: 100 }],
    data: [['Thông tin', 'Giá trị'].map(header), ...[
      ['Nguồn', report.scope || 'Lịch sử truy cập 1Office, toàn bộ nhật ký tài khoản hiện tại được phép xem; không dùng bộ lọc của trang nguồn.'],
      ['Ngày chốt', `${report.cutoff}, hết ngày theo UTC+7`],
      ['Thời gian thu thập (UTC)', report.collectedAt],
      ['Mốc dữ liệu', 'Giữ các ID có tại lúc bắt đầu quét; bản ghi mới có ID lớn hơn mốc đầu không được bổ sung vào lần thu thập này. Đổi ngày chốt chỉ tính lại trên dữ liệu đã thu thập.'],
      ['Phạm vi lọc', report.scope ? 'Lọc theo trạng thái/cơ cấu hiện tại. Nhật ký và nhân sự không phải snapshot giao dịch đồng thời. Tài khoản thiếu thông tin lọc không được tự suy đoán.' : 'Không lọc theo nhân sự.'],
      ['Trạng thái tài khoản', report.personnel?.activeAccounts ? 'Chỉ tài khoản hiện còn active tại lần xác minh danh sách người dùng. Loại tài khoản đã khóa, không xác minh được hoặc trùng tên đăng nhập. Không phản ánh trạng thái tại ngày chốt lịch sử.' : 'Chưa xác minh trạng thái active'],
      ['Bản ghi đã đối chiếu', String(report.recordCount)],
      ['Bản ghi cũ nhất trong nguồn', report.oldestTime || 'Nguồn không có bản ghi'],
      ['Giới hạn', 'Chỉ phản ánh nhật ký còn được 1Office lưu và cấp quyền xem. Không có tài khoản trong báo cáo không có nghĩa chưa từng truy cập.'],
      ['Ngữ nghĩa', 'Truy cập không đồng nghĩa đăng nhập thành công. Số bản ghi không phải số lần đăng nhập.'],
      ['Thông tin nhân sự', 'Họ tên lấy từ bản ghi nguồn, chưa xác minh là tên tại thời điểm lịch sử. Không suy ra phòng ban/chức danh trong quá khứ.'],
      ['Vị trí và phòng ban', 'Thông tin hiện tại từ danh sách người dùng và nguồn chọn người dùng 1Office, ghép chính xác theo tài khoản và ID. Ô trống là chưa có dữ liệu. Không dùng họ tên để ghép.'],
      ['Thời gian lấy thông tin nhân sự (UTC)', report.personnel?.collectedAt || 'Chưa lấy được'],
      ['Cảnh báo thông tin nhân sự', report.personnel?.warnings.join('; ') || (report.personnel ? 'Không có cảnh báo ghép tài khoản' : 'Chưa lấy thông tin nhân sự')],
    ].map((row) => row.map(text))],
  }];
}

export async function downloadAccessWorkbook(report: AccessReport): Promise<void> {
  await writeExcelFile(buildAccessWorkbook(report), { fontFamily: 'Arial', fontSize: 10 })
    .toFile(`tera-clip-truy-cap-den-${report.cutoff}.xlsx`);
}
