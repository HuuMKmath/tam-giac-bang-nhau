// Nhận kết quả bài Bội chung nhỏ nhất (mã BC6-) và ghi vào bảng tính này.
// Dán vào Apps Script của Google Sheet RIÊNG cho bài này.
const PREFIX = 'BC6-';
const TOPICS = ['bc', 'ts', 'sy', 'tt'];
const HEADERS = ['Thời điểm nộp', 'Họ tên', 'Lớp', 'Ngày', 'Loại bài', 'Mã đề', 'Điểm', 'Trắc nghiệm', 'Đúng sai', 'Trả lời ngắn',
  'Thời gian (phút)', 'Giới hạn (phút)', 'Bội chung', 'Tìm BCNN bằng phân tích', 'Từ BCNN suy ra bội chung', 'Bài toán thực tế', 'Ghi chú', 'Mã kết quả'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    const code = String(body.code || '');
    if (code.indexOf(PREFIX) !== 0) return ok('bad');
    const json = Utilities.newBlob(Utilities.base64Decode(code.slice(PREFIX.length))).getDataAsString('UTF-8');
    const o = JSON.parse(json);

    const sh = SpreadsheetApp.getActiveSheet();
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }
    // Nếu mã này đã có thì không ghi lại (học sinh bấm gửi lại)
    const last = sh.getLastRow();
    if (last > 1) {
      const codes = sh.getRange(2, HEADERS.length, last - 1, 1).getValues();
      for (let i = 0; i < codes.length; i++) if (codes[i][0] === code) return ok('dup');
    }

    const topic = TOPICS.map(function (k) { return o.t && o.t[k] && o.t[k][1] ? o.t[k][0] + '/' + o.t[k][1] : ''; });
    const limit = o.L || (o.q === 'quick' ? 300 : 1500);
    const fast = (typeof o.m === 'number' && o.m < 0.25 * limit && o.s < 5) ? 'Làm nhanh, điểm thấp' : '';
    const row = [new Date(), o.n || '', o.c || '', o.d || '', o.q === 'quick' ? 'Kiểm tra nhanh' : 'Kiểm tra', o.k || '',
      o.s, o.p ? o.p[0] : '', o.p ? o.p[1] : '', o.p ? o.p[2] : '',
      typeof o.m === 'number' ? Math.round(o.m / 6) / 10 : '', o.L ? o.L / 60 : ''].concat(topic, [fast, code]);
    sh.appendRow(row);
    return ok('saved');
  } catch (err) {
    return ok('error');
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return ok('ready'); }

function ok(msg) {
  return ContentService.createTextOutput(msg);
}
