// Nhận kết quả từ các trang luyện toán và ghi vào bảng tính.
// Mỗi bài có một tiền tố mã và một trang tính riêng.
// Bài Tam giác bằng nhau (TG7-) vẫn ghi vào trang tính đầu tiên như cũ, các cột không đổi.
// Bài Bội chung nhỏ nhất (BC6-) ghi vào trang tính "Bội chung nhỏ nhất", tự tạo nếu chưa có.
const LESSONS = {
  'TG7-': {
    sheet: null, // null = trang tính đầu tiên (nơi đang lưu dữ liệu cũ)
    topics: ['tu', 'ccc', 'cgc', 'gcg'],
    names: ['Tổng quan tam giác', 'Cạnh-cạnh-cạnh', 'Cạnh-góc-cạnh', 'Góc-cạnh-góc']
  },
  'BC6-': {
    sheet: 'Bội chung nhỏ nhất',
    topics: ['bc', 'ts', 'sy', 'tt'],
    names: ['Bội chung', 'Tìm BCNN bằng phân tích', 'Từ BCNN suy ra bội chung', 'Bài toán thực tế']
  }
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    const code = String(body.code || '');
    const prefix = Object.keys(LESSONS).filter(function (p) { return code.indexOf(p) === 0; })[0];
    if (!prefix) return ok('bad');
    const cfg = LESSONS[prefix];
    const json = Utilities.newBlob(Utilities.base64Decode(code.slice(prefix.length))).getDataAsString('UTF-8');
    const o = JSON.parse(json);

    const HEADERS = ['Thời điểm nộp', 'Họ tên', 'Lớp', 'Ngày', 'Loại bài', 'Mã đề', 'Điểm', 'Trắc nghiệm', 'Đúng sai', 'Trả lời ngắn',
      'Thời gian (phút)', 'Giới hạn (phút)'].concat(cfg.names, ['Ghi chú', 'Mã kết quả']);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = cfg.sheet ? ss.getSheetByName(cfg.sheet) : ss.getSheets()[0];
    if (!sh) sh = ss.insertSheet(cfg.sheet);
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

    const topic = cfg.topics.map(function (k) { return o.t && o.t[k] && o.t[k][1] ? o.t[k][0] + '/' + o.t[k][1] : ''; });
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
