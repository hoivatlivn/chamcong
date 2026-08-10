/*************************************************************
 * HỆ THỐNG CHẤM CÔNG - HỘI VẬT LÝ VIỆT NAM
 * Google Apps Script Web App
 *
 * CÁCH CÀI ĐẶT (đọc kỹ file README.md đi kèm):
 * 1. Tạo 1 Google Sheet mới, đặt tên tuỳ ý (VD: "ChamCong_HoiVatLy_DB")
 * 2. Mở Extensions > Apps Script
 * 3. Xoá code mẫu, dán toàn bộ nội dung file Code.gs này vào
 * 4. Tạo thêm 1 file HTML tên "Index", dán nội dung file Index.html vào
 * 5. Chạy hàm khoiTaoHeThong() 1 lần (Run > khoiTaoHeThong) để tạo sheet
 * 6. Nếu bạn nâng cấp từ bản cũ, chạy thêm hàm capNhatHeThongLenBanMoi()
 * 7. Vào sheet "Config" điền toạ độ công ty, bán kính, email admin/giám đốc
 * 8. Vào sheet "Users" gán VaiTro (admin/giamdoc/truongphong/nhanvien) và
 *    PhongBan cho từng nhân viên
 * 9. Chạy hàm taoTriggerHangNgay() 1 lần để bật báo cáo tự động
 * 10. Deploy > New deployment > Web app > Execute as: Me,
 *     Who has access: Anyone with the link
 *************************************************************/

const TEN_CONG_TY = 'HỘI VẬT LÝ VIỆT NAM';

const SHEET_USERS = 'Users';
const SHEET_CHAMCONG = 'ChamCong';
const SHEET_CONFIG = 'Config';
const SHEET_SESSIONS = 'Sessions';
const SHEET_DONNGHIPHEP = 'DonNghiPhep';
const SHEET_LICHSUTHIETBI = 'LichSuThietBi';

const TRANG_THAI = {
  GIO_VAO: 'Giờ vào',
  GIO_RA: 'Giờ ra',
  NGHI_KHONG_LUONG: 'Nghỉ không lương',
  NGHI_KHONG_LY_DO: 'Nghỉ không lý do'
};

const VAI_TRO = {
  ADMIN: 'admin',
  GIAM_DOC: 'giamdoc',
  TRUONG_PHONG: 'truongphong',
  NHAN_VIEN: 'nhanvien'
};

const THOI_HAN_SESSION_NGAY = 60; // đăng nhập nhớ 60 ngày - mở lại app trong thời gian này sẽ vào thẳng, không cần đăng nhập lại

/* ============================================================
 *  KHỞI TẠO
 * ============================================================ */

function khoiTaoHeThong() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sh = ss.getSheetByName(SHEET_USERS);
  if (!sh) sh = ss.insertSheet(SHEET_USERS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Email', 'HoTen', 'MatKhauHash', 'Salt', 'VaiTro', 'PhongBan', 'ChucDanh', 'AnhDaiDien', 'NgayTao']);
  }

  sh = ss.getSheetByName(SHEET_CHAMCONG);
  if (!sh) sh = ss.insertSheet(SHEET_CHAMCONG);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Ngay', 'Email', 'HoTen', 'ThoiGian', 'TrangThai', 'Lat', 'Lng', 'KhoangCach_m', 'ThietBi', 'GhiChu']);
  }

  sh = ss.getSheetByName(SHEET_SESSIONS);
  if (!sh) sh = ss.insertSheet(SHEET_SESSIONS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Token', 'Email', 'HetHan']);
  }

  sh = ss.getSheetByName(SHEET_DONNGHIPHEP);
  if (!sh) sh = ss.insertSheet(SHEET_DONNGHIPHEP);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['MaDon', 'Email', 'HoTen', 'PhongBan', 'SoNgay', 'NgayBatDau', 'NgayKetThuc', 'NgayGui', 'TrangThai', 'NguoiDuyet', 'NgayDuyet', 'LyDo']);
  }

  sh = ss.getSheetByName(SHEET_LICHSUTHIETBI);
  if (!sh) sh = ss.insertSheet(SHEET_LICHSUTHIETBI);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Ngay', 'ThoiGian', 'Email', 'HoTen', 'HanhDong', 'ThietBiThoRaw', 'ThietBiRutGon', 'Lat', 'Lng']);
  }

  sh = ss.getSheetByName(SHEET_CONFIG);
  if (!sh) sh = ss.insertSheet(SHEET_CONFIG);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Key', 'Value', 'GhiChu']);
    sh.appendRow(['LAT', '10.7756587', 'Vĩ độ công ty (lấy từ Google Maps)']);
    sh.appendRow(['LNG', '106.7004238', 'Kinh độ công ty']);
    sh.appendRow(['RADIUS_M', '200', 'Bán kính cho phép chấm công (mét)']);
    sh.appendRow(['ADMIN_EMAILS', 'admin@gmail.com', 'Email nhận báo cáo tổng hợp hàng ngày, cách nhau dấu phẩy']);
    sh.appendRow(['GIO_CHOT', '18:00', 'Giờ chốt công mỗi ngày để gửi báo cáo & đánh dấu nghỉ không lý do']);
    sh.appendRow(['GIO_VAO_CHUAN', '08:00', 'Giờ vào làm chuẩn']);
    sh.appendRow(['GIO_RA_CHUAN', '17:00', 'Giờ tan làm chuẩn']);
    sh.appendRow(['BIEN_DO_PHUT', '15', 'Số phút cho phép trước/sau giờ chuẩn để chấm giờ vào - giờ ra']);
    sh.appendRow(['NGAY_BAO_CAO_THIET_BI', 'MONDAY', 'Thứ trong tuần gửi báo cáo đa thiết bị (viết hoa, tiếng Anh): MONDAY/TUESDAY/WEDNESDAY/THURSDAY/FRIDAY/SATURDAY/SUNDAY']);
    sh.appendRow(['THONG_BAO_TUC_THI_XA_CONG_TY', 'BAT', 'BAT = gửi cảnh báo ngay cho Giám đốc/admin khi nhân viên chấm công cách xa công ty; TAT = tắt']);
    sh.appendRow(['GOOGLE_CHAT_WEBHOOK_URL', '', 'Dán URL webhook Google Chat vào đây để nhận cảnh báo qua Google Chat thay vì email. Để trống = dùng email.']);
    sh.appendRow(['GIO_BAO_CAO_THIET_BI', '08:00', 'Giờ gửi báo cáo đa thiết bị hàng tuần']);
  }

  capNhatHeThongLenBanMoi();

  SpreadsheetApp.getUi().alert('Đã khởi tạo xong các sheet. Vào sheet "Config" để cập nhật toạ độ công ty, email; vào sheet "Users" để gán vai trò/phòng ban.');
}

function capNhatHeThongLenBanMoi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(SHEET_CHAMCONG);
  if (sheet) {
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (header.indexOf('ThietBi') === -1) {
      const viTriChen = header.length;
      sheet.insertColumnBefore(viTriChen);
      sheet.getRange(1, viTriChen).setValue('ThietBi');
    }
  }

  sheet = ss.getSheetByName(SHEET_USERS);
  if (sheet) {
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (header.indexOf('PhongBan') === -1) {
      const viTriChen = header.length;
      sheet.insertColumnBefore(viTriChen);
      sheet.getRange(1, viTriChen).setValue('PhongBan');
    }
    const headerSauPhongBan = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headerSauPhongBan.indexOf('ChucDanh') === -1) {
      const idxPhongBan = headerSauPhongBan.indexOf('PhongBan'); // 0-based
      const viTriChenChucDanh = idxPhongBan + 2; // 1-based, ngay sau cột PhongBan
      sheet.insertColumnBefore(viTriChenChucDanh);
      sheet.getRange(1, viTriChenChucDanh).setValue('ChucDanh');
    }
    const header2 = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (header2.indexOf('AnhDaiDien') === -1) {
      const viTriChen2 = header2.length; // chèn ngay trước cột cuối (NgayTao)
      sheet.insertColumnBefore(viTriChen2);
      sheet.getRange(1, viTriChen2).setValue('AnhDaiDien');
    }
    // Đã bỏ tính năng chấm công bằng khuôn mặt - nếu sheet cũ còn cột KhuonMat thì xoá đi
    const header3 = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idxKhuonMat = header3.indexOf('KhuonMat'); // 0-based
    if (idxKhuonMat !== -1) {
      sheet.deleteColumn(idxKhuonMat + 1);
    }
  }

  if (!ss.getSheetByName(SHEET_DONNGHIPHEP)) {
    sheet = ss.insertSheet(SHEET_DONNGHIPHEP);
    sheet.appendRow(['MaDon', 'Email', 'HoTen', 'PhongBan', 'SoNgay', 'NgayBatDau', 'NgayKetThuc', 'NgayGui', 'TrangThai', 'NguoiDuyet', 'NgayDuyet', 'LyDo']);
  }

  if (!ss.getSheetByName(SHEET_LICHSUTHIETBI)) {
    sheet = ss.insertSheet(SHEET_LICHSUTHIETBI);
    sheet.appendRow(['Ngay', 'ThoiGian', 'Email', 'HoTen', 'HanhDong', 'ThietBiThoRaw', 'ThietBiRutGon', 'Lat', 'Lng']);
  }

  const configSheet = ss.getSheetByName(SHEET_CONFIG);
  if (configSheet) {
    const config = layConfig();
    const macDinh = [
      ['GIO_VAO_CHUAN', '08:00', 'Giờ vào làm chuẩn'],
      ['GIO_RA_CHUAN', '17:00', 'Giờ tan làm chuẩn'],
      ['BIEN_DO_PHUT', '15', 'Số phút cho phép trước/sau giờ chuẩn để chấm giờ vào - giờ ra'],
      ['NGAY_BAO_CAO_THIET_BI', 'MONDAY', 'Thứ trong tuần gửi báo cáo đa thiết bị (viết hoa, tiếng Anh): MONDAY/TUESDAY/.../SUNDAY'],
      ['THONG_BAO_TUC_THI_XA_CONG_TY', 'BAT', 'BAT = gửi cảnh báo ngay cho Giám đốc/admin khi nhân viên chấm công cách xa công ty; TAT = tắt'],
      ['GOOGLE_CHAT_WEBHOOK_URL', '', 'Dán URL webhook Google Chat vào đây để nhận cảnh báo qua Google Chat thay vì email. Để trống = dùng email.'],
      ['GIO_BAO_CAO_THIET_BI', '08:00', 'Giờ gửi báo cáo đa thiết bị hàng tuần']
    ];
    macDinh.forEach(function (row) {
      if (config[row[0]] === undefined) {
        configSheet.appendRow(row);
      }
    });

    // Đã bỏ tính năng chấm công bằng khuôn mặt - xoá các dòng config cũ liên quan nếu còn
    const cacKeyKhuonMatCu = ['BAT_BUOC_KHUON_MAT', 'NGUONG_KHOP_KHUON_MAT'];
    const configData = configSheet.getDataRange().getValues();
    for (let i = configData.length - 1; i >= 1; i--) {
      if (cacKeyKhuonMatCu.indexOf(configData[i][0]) !== -1) {
        configSheet.deleteRow(i + 1);
      }
    }
  }

  try { SpreadsheetApp.getUi().alert('Đã cập nhật cấu trúc sheet lên phiên bản mới nhất.'); } catch (e) {}
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Chấm công Hội Vật Lý')
    .addItem('Khởi tạo hệ thống', 'khoiTaoHeThong')
    .addItem('Cập nhật lên phiên bản mới (chạy 1 lần khi nâng cấp)', 'capNhatHeThongLenBanMoi')
    .addItem('Bật báo cáo tự động hàng ngày', 'taoTriggerHangNgay')
    .addItem('Gửi báo cáo ngay bây giờ (test)', 'guiBaoCaoHangNgay')
    .addSeparator()
    .addItem('Bật báo cáo đa thiết bị hàng tuần', 'taoTriggerBaoCaoThietBiHangTuan')
    .addItem('Gửi báo cáo đa thiết bị ngay (test)', 'guiBaoCaoThietBiHangTuan')
    .addSeparator()
    .addItem('Kiểm tra webhook Google Chat (test)', 'kiemTraWebhookGoogleChat')
    .addToUi();
}

function taoTriggerHangNgay() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'guiBaoCaoHangNgay') {
      ScriptApp.deleteTrigger(t);
    }
  });

  const config = layConfig();
  const gioChotStr = config.GIO_CHOT || '18:00';
  const phan = gioChotStr.split(':');
  const gio = parseInt(phan[0], 10);
  const phut = parseInt(phan[1], 10) || 0;

  ScriptApp.newTrigger('guiBaoCaoHangNgay')
    .timeBased()
    .atHour(gio)
    .nearMinute(phut)
    .everyDays(1)
    .create();

  SpreadsheetApp.getUi().alert('Đã bật trigger gửi báo cáo tự động lúc ' + gioChotStr + ' hàng ngày.');
}

function taoTriggerBaoCaoThietBiHangTuan() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'guiBaoCaoThietBiHangTuan') {
      ScriptApp.deleteTrigger(t);
    }
  });

  const config = layConfig();
  const ngayStr = String(config.NGAY_BAO_CAO_THIET_BI || 'MONDAY').toUpperCase().trim();
  const bangThu = {
    MONDAY: ScriptApp.WeekDay.MONDAY, TUESDAY: ScriptApp.WeekDay.TUESDAY, WEDNESDAY: ScriptApp.WeekDay.WEDNESDAY,
    THURSDAY: ScriptApp.WeekDay.THURSDAY, FRIDAY: ScriptApp.WeekDay.FRIDAY, SATURDAY: ScriptApp.WeekDay.SATURDAY, SUNDAY: ScriptApp.WeekDay.SUNDAY
  };
  const thu = bangThu[ngayStr] || ScriptApp.WeekDay.MONDAY;

  const gioStr = config.GIO_BAO_CAO_THIET_BI || '08:00';
  const phan = gioStr.split(':');
  const gio = parseInt(phan[0], 10);
  const phut = parseInt(phan[1], 10) || 0;

  ScriptApp.newTrigger('guiBaoCaoThietBiHangTuan')
    .timeBased()
    .onWeekDay(thu)
    .atHour(gio)
    .nearMinute(phut)
    .create();

  SpreadsheetApp.getUi().alert('Đã bật trigger gửi báo cáo đa thiết bị vào ' + ngayStr + ' lúc ' + gioStr + ' hàng tuần.');
}

/* ============================================================
 *  WEB APP ENTRY POINT
 * ============================================================ */

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(TEN_CONG_TY + ' - Chấm công')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ============================================================
 *  HÀM TIỆN ÍCH
 * ============================================================ */

function chuanHoaGio(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
  }
  return String(value).trim();
}

function layConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    let value = data[i][1];
    if (value instanceof Date) {
      value = chuanHoaGio(value);
    }
    config[data[i][0]] = value;
  }
  return config;
}

function taoSalt() {
  return Utilities.getUuid();
}

function bamMatKhau(matKhau, salt) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    matKhau + '::' + salt
  );
  return digest.map(function (b) {
    return ('0' + (b & 0xff).toString(16)).slice(-2);
  }).join('');
}

function taoMatKhauNgauNhien(doDai) {
  const kyTu = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789';
  let kq = '';
  for (let i = 0; i < doDai; i++) {
    kq += kyTu.charAt(Math.floor(Math.random() * kyTu.length));
  }
  return kq;
}

function taoToken() {
  return Utilities.getUuid() + Utilities.getUuid();
}

function taoMaDon() {
  return Utilities.getUuid();
}

function tinhKhoangCachMet(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLng = (lng2 - lng1) * Math.PI / 180;
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
    Math.sin(dLng / 2) * Math.sin(dLng / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

function dinhDangNgay(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function chuanHoaNgay(value) {
  if (value instanceof Date) {
    return dinhDangNgay(value);
  }
  return String(value).trim();
}

function dinhDangGio(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'HH:mm:ss');
}

function soPhutTrongNgay(date) {
  return date.getHours() * 60 + date.getMinutes();
}

function chuoiGioThanhPhut(gioStr) {
  const phan = String(gioStr).split(':');
  return parseInt(phan[0], 10) * 60 + (parseInt(phan[1], 10) || 0);
}

function trongKhungGioChamCong(gioChuanStr, bienDoPhut, thoiDiem) {
  const phutHienTai = soPhutTrongNgay(thoiDiem || new Date());
  const phutChuan = chuoiGioThanhPhut(gioChuanStr);
  return Math.abs(phutHienTai - phutChuan) <= bienDoPhut;
}

// Rút gọn chuỗi User-Agent thô thành nhãn dễ đọc, dùng để so sánh "cùng 1 thiết bị"
// mà không bị sai lệch bởi khác biệt phiên bản nhỏ (ví dụ Safari 17.1 vs 17.2).
// Đây chỉ là thông tin tham khảo, không phải ID thiết bị thật.
function rutGonThietBi(userAgent) {
  if (!userAgent) return 'Không xác định';
  const ua = String(userAgent);

  let loaiMay = 'Khác';
  if (/iPhone/i.test(ua)) loaiMay = 'iPhone';
  else if (/iPad/i.test(ua)) loaiMay = 'iPad';
  else if (/Android/i.test(ua)) loaiMay = 'Android';
  else if (/Macintosh/i.test(ua)) loaiMay = 'Mac';
  else if (/Windows/i.test(ua)) loaiMay = 'Windows';
  else if (/Linux/i.test(ua)) loaiMay = 'Linux';

  let trinhDuyet = '';
  if (/Edg\//i.test(ua)) trinhDuyet = 'Edge';
  else if (/CriOS/i.test(ua)) trinhDuyet = 'Chrome (iOS)';
  else if (/Chrome\//i.test(ua)) trinhDuyet = 'Chrome';
  else if (/FxiOS/i.test(ua)) trinhDuyet = 'Firefox (iOS)';
  else if (/Firefox\//i.test(ua)) trinhDuyet = 'Firefox';
  else if (/Safari\//i.test(ua)) trinhDuyet = 'Safari';

  return loaiMay + (trinhDuyet ? ' - ' + trinhDuyet : '');
}

function ghiLichSuThietBi(email, hoTen, hanhDong, thietBiRaw, lat, lng) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LICHSUTHIETBI);
  if (!sheet) return;
  const now = new Date();
  sheet.appendRow([
    dinhDangNgay(now),
    dinhDangGio(now),
    email,
    hoTen,
    hanhDong,
    thietBiRaw || '',
    rutGonThietBi(thietBiRaw),
    lat || '',
    lng || ''
  ]);
}

/* ============================================================
 *  QUẢN LÝ USER
 * ============================================================ */

function timUserTheoEmail(email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(email).toLowerCase()) {
      return {
        rowIndex: i + 1,
        email: data[i][0],
        hoTen: data[i][1],
        matKhauHash: data[i][2],
        salt: data[i][3],
        vaiTro: data[i][4],
        phongBan: data[i][5] || '',
        chucDanh: data[i][6] || '',
        anhDaiDien: data[i][7] || '',
        ngayTao: data[i][8]
      };
    }
  }
  return null;
}

function layDanhSachUserHopLe() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  const ds = [];
  for (let i = 1; i < data.length; i++) {
    ds.push({ email: data[i][0], hoTen: data[i][1], vaiTro: data[i][4], phongBan: data[i][5] || '', chucDanh: data[i][6] || '' });
  }
  return ds;
}

function layEmailGiamDoc() {
  return layDanhSachUserHopLe()
    .filter(function (u) { return u.vaiTro === VAI_TRO.GIAM_DOC; })
    .map(function (u) { return u.email; });
}

function layEmailTruongPhong(phongBan) {
  if (!phongBan) return [];
  return layDanhSachUserHopLe()
    .filter(function (u) { return u.vaiTro === VAI_TRO.TRUONG_PHONG && u.phongBan === phongBan; })
    .map(function (u) { return u.email; });
}

/* ============================================================
 *  ĐĂNG KÝ
 * ============================================================ */

function dangKy(email, hoTen) {
  email = String(email).trim().toLowerCase();
  hoTen = String(hoTen).trim();

  if (!email || !hoTen) {
    return { success: false, message: 'Vui lòng nhập đầy đủ Họ tên và Email.' };
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { success: false, message: 'Email không hợp lệ.' };
  }
  if (timUserTheoEmail(email)) {
    return { success: false, message: 'Email này đã đăng ký tài khoản rồi. Vui lòng đăng nhập.' };
  }

  const matKhau = taoMatKhauNgauNhien(10);
  const salt = taoSalt();
  const hash = bamMatKhau(matKhau, salt);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const soUserHienCo = sheet.getLastRow() - 1;
  const vaiTro = soUserHienCo === 0 ? VAI_TRO.ADMIN : VAI_TRO.NHAN_VIEN;

  sheet.appendRow([email, hoTen, hash, salt, vaiTro, '', '', '', new Date()]);

  try {
    MailApp.sendEmail({
      to: email,
      subject: 'Mật khẩu đăng nhập hệ thống chấm công - ' + TEN_CONG_TY,
      htmlBody:
        '<p>Chào <b>' + hoTen + '</b>,</p>' +
        '<p>Bạn đã đăng ký thành công tài khoản chấm công tại <b>' + TEN_CONG_TY + '</b>.</p>' +
        '<p>Thông tin đăng nhập:</p>' +
        '<ul><li>Email: <b>' + email + '</b></li>' +
        '<li>Mật khẩu: <b style="font-size:16px;color:#c0392b">' + matKhau + '</b></li></ul>' +
        '<p>Vui lòng đăng nhập tại trang chấm công và đổi mật khẩu nếu cần.</p>' +
        '<p><i>Email tự động, vui lòng không trả lời.</i></p>'
    });
  } catch (err) {
    return { success: false, message: 'Tạo tài khoản thành công nhưng gửi email thất bại: ' + err.message };
  }

  return { success: true, message: 'Đăng ký thành công! Mật khẩu đã được gửi tới email ' + email + '. Vui lòng liên hệ admin để được gán phòng ban.' };
}

/* ============================================================
 *  ĐĂNG NHẬP / SESSION
 * ============================================================ */

function dangNhap(email, matKhau) {
  email = String(email).trim().toLowerCase();
  const user = timUserTheoEmail(email);
  if (!user) {
    return { success: false, message: 'Email chưa được đăng ký.' };
  }
  const hash = bamMatKhau(matKhau, user.salt);
  if (hash !== user.matKhauHash) {
    return { success: false, message: 'Sai mật khẩu.' };
  }

  const token = taoToken();
  const hetHan = new Date();
  hetHan.setDate(hetHan.getDate() + THOI_HAN_SESSION_NGAY);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESSIONS);
  sheet.appendRow([token, email, hetHan]);

  return {
    success: true,
    token: token,
    hoTen: user.hoTen,
    vaiTro: user.vaiTro,
    phongBan: user.phongBan,
    chucDanh: user.chucDanh,
    anhDaiDien: user.anhDaiDien,
    email: user.email
  };
}

function kiemTraSession(token) {
  if (!token) return { success: false };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESSIONS);
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      const hetHan = new Date(data[i][2]);
      if (hetHan < now) return { success: false };
      const user = timUserTheoEmail(data[i][1]);
      if (!user) return { success: false };
      return { success: true, hoTen: user.hoTen, vaiTro: user.vaiTro, phongBan: user.phongBan, chucDanh: user.chucDanh, anhDaiDien: user.anhDaiDien, email: user.email };
    }
  }
  return { success: false };
}

function dangXuat(token) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESSIONS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return { success: true };
}

/* ============================================================
 *  ĐỔI MẬT KHẨU
 * ============================================================ */

/**
 * Lưu ảnh đại diện (đã được nén/thu nhỏ ở phía client thành base64 JPEG).
 * Giới hạn kích thước để tránh vượt quá giới hạn 50.000 ký tự/ô của Google Sheets.
 */
function capNhatAnhDaiDien(token, base64Anh) {
  const session = kiemTraSession(token);
  if (!session.success) return { success: false, message: 'Phiên đăng nhập hết hạn, vui lòng đăng nhập lại.' };

  if (!base64Anh) return { success: false, message: 'Không có dữ liệu ảnh.' };
  if (base64Anh.length > 45000) {
    return { success: false, message: 'Ảnh quá lớn để lưu, vui lòng chọn ảnh khác hoặc thử lại.' };
  }

  const user = timUserTheoEmail(session.email);
  if (!user) return { success: false, message: 'Không tìm thấy tài khoản.' };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  sheet.getRange(user.rowIndex, 8).setValue(base64Anh); // cột AnhDaiDien

  return { success: true, message: 'Đã cập nhật ảnh đại diện.' };
}

function doiMatKhau(token, matKhauCu, matKhauMoi) {
  const session = kiemTraSession(token);
  if (!session.success) {
    return { success: false, message: 'Phiên đăng nhập hết hạn, vui lòng đăng nhập lại.' };
  }

  const user = timUserTheoEmail(session.email);
  if (!user) {
    return { success: false, message: 'Không tìm thấy tài khoản.' };
  }

  const hashCu = bamMatKhau(matKhauCu, user.salt);
  if (hashCu !== user.matKhauHash) {
    return { success: false, message: 'Mật khẩu hiện tại không đúng.' };
  }

  if (!matKhauMoi || matKhauMoi.length < 6) {
    return { success: false, message: 'Mật khẩu mới phải có ít nhất 6 ký tự.' };
  }
  if (matKhauMoi === matKhauCu) {
    return { success: false, message: 'Mật khẩu mới phải khác mật khẩu hiện tại.' };
  }

  const saltMoi = taoSalt();
  const hashMoi = bamMatKhau(matKhauMoi, saltMoi);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  sheet.getRange(user.rowIndex, 3).setValue(hashMoi);
  sheet.getRange(user.rowIndex, 4).setValue(saltMoi);

  try {
    const now = new Date();
    const thoiGian = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm 'ngày' dd/MM/yyyy");
    MailApp.sendEmail({
      to: user.email,
      subject: 'Mật khẩu đã được thay đổi - ' + TEN_CONG_TY,
      htmlBody:
        '<p>Chào <b>' + user.hoTen + '</b>,</p>' +
        '<p>Mật khẩu tài khoản chấm công của bạn tại <b>' + TEN_CONG_TY + '</b> vừa được đổi thành công lúc <b>' + thoiGian + '</b>.</p>' +
        '<p>Nếu không phải bạn thực hiện thay đổi này, vui lòng liên hệ ngay với quản trị viên hệ thống.</p>' +
        '<p><i>Email tự động, vui lòng không trả lời.</i></p>'
    });
  } catch (err) {
    return { success: true, message: 'Đổi mật khẩu thành công, nhưng gửi email xác nhận thất bại: ' + err.message };
  }

  return { success: true, message: 'Đổi mật khẩu thành công! Email xác nhận đã được gửi tới ' + user.email + '.' };
}

/* ============================================================
 *  QUẢN LÝ NHÂN VIÊN (chỉ admin) - gán vai trò & phòng ban
 * ============================================================ */

function layDanhSachTatCaUserChoAdmin(token) {
  const session = kiemTraSession(token);
  if (!session.success) return { success: false, message: 'Phiên đăng nhập hết hạn.' };
  if (session.vaiTro !== VAI_TRO.ADMIN) return { success: false, message: 'Chỉ admin mới có quyền này.' };

  return { success: true, danhSach: layDanhSachUserHopLe() };
}

function capNhatVaiTroPhongBan(token, email, vaiTroMoi, phongBanMoi, chucDanhMoi) {
  const session = kiemTraSession(token);
  if (!session.success) return { success: false, message: 'Phiên đăng nhập hết hạn.' };
  if (session.vaiTro !== VAI_TRO.ADMIN) return { success: false, message: 'Chỉ admin mới có quyền này.' };

  const cacVaiTroHopLe = [VAI_TRO.ADMIN, VAI_TRO.GIAM_DOC, VAI_TRO.TRUONG_PHONG, VAI_TRO.NHAN_VIEN];
  if (cacVaiTroHopLe.indexOf(vaiTroMoi) === -1) {
    return { success: false, message: 'Vai trò không hợp lệ.' };
  }

  const user = timUserTheoEmail(email);
  if (!user) return { success: false, message: 'Không tìm thấy tài khoản.' };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  sheet.getRange(user.rowIndex, 5).setValue(vaiTroMoi);
  sheet.getRange(user.rowIndex, 6).setValue(phongBanMoi || '');
  sheet.getRange(user.rowIndex, 7).setValue(chucDanhMoi || '');

  return { success: true, message: 'Đã cập nhật vai trò/phòng ban/chức danh cho ' + user.hoTen + '.' };
}

/* ============================================================
 *  ĐƠN NGHỈ PHÉP (có quy trình duyệt bởi Giám đốc)
 * ============================================================ */

function timDonTheoMa(maDon) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DONNGHIPHEP);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === maDon) {
      return {
        rowIndex: i + 1,
        maDon: data[i][0], email: data[i][1], hoTen: data[i][2], phongBan: data[i][3],
        soNgay: data[i][4], ngayBatDau: chuanHoaNgay(data[i][5]), ngayKetThuc: chuanHoaNgay(data[i][6]),
        ngayGui: data[i][7], trangThai: data[i][8], nguoiDuyet: data[i][9], ngayDuyet: data[i][10], lyDo: data[i][11]
      };
    }
  }
  return null;
}

function coDonNghiPhepDuocDuyetHomNay(email) {
  const homNay = dinhDangNgay(new Date());
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DONNGHIPHEP);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).toLowerCase() === String(email).toLowerCase() && data[i][8] === 'Đã duyệt') {
      const batDau = chuanHoaNgay(data[i][5]);
      const ketThuc = chuanHoaNgay(data[i][6]);
      if (homNay >= batDau && homNay <= ketThuc) {
        return { batDau: batDau, ketThuc: ketThuc, soNgay: data[i][4] };
      }
    }
  }
  return null;
}

function coDonChoDuyet(email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DONNGHIPHEP);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).toLowerCase() === String(email).toLowerCase() && data[i][8] === 'Chờ duyệt') {
      return true;
    }
  }
  return false;
}

function guiDonNghiPhep(token, soNgay, lyDo) {
  const session = kiemTraSession(token);
  if (!session.success) return { success: false, message: 'Phiên đăng nhập hết hạn, vui lòng đăng nhập lại.' };

  soNgay = parseInt(soNgay, 10);
  if (!soNgay || soNgay < 1 || soNgay > 30) {
    return { success: false, message: 'Số ngày nghỉ không hợp lệ (1-30 ngày).' };
  }

  if (coDonChoDuyet(session.email)) {
    return { success: false, message: 'Bạn đang có 1 đơn nghỉ phép chờ duyệt, vui lòng đợi kết quả trước khi gửi đơn mới.' };
  }
  if (coDonNghiPhepDuocDuyetHomNay(session.email)) {
    return { success: false, message: 'Bạn đang trong thời gian nghỉ phép đã được duyệt.' };
  }

  const user = timUserTheoEmail(session.email);
  const homNay = new Date();
  const ngayBatDau = dinhDangNgay(homNay);
  const ngayKetThucDate = new Date(homNay);
  ngayKetThucDate.setDate(ngayKetThucDate.getDate() + (soNgay - 1));
  const ngayKetThuc = dinhDangNgay(ngayKetThucDate);

  const maDon = taoMaDon();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DONNGHIPHEP);
  sheet.appendRow([maDon, session.email, session.hoTen, user.phongBan || '', soNgay, ngayBatDau, ngayKetThuc, new Date(), 'Chờ duyệt', '', '', lyDo || '']);

  const nguoiNhan = layEmailTruongPhong(user.phongBan).concat(layEmailGiamDoc());
  const nguoiNhanDuyNhat = nguoiNhan.filter(function (email, idx) { return nguoiNhan.indexOf(email) === idx && email; });

  if (nguoiNhanDuyNhat.length > 0) {
    try {
      MailApp.sendEmail({
        to: nguoiNhanDuyNhat.join(','),
        subject: 'Đơn xin nghỉ phép mới - ' + session.hoTen + ' - ' + TEN_CONG_TY,
        htmlBody:
          '<p>Nhân viên <b>' + session.hoTen + '</b> (' + session.email + ', phòng ban: ' + (user.phongBan || 'chưa gán') + ') vừa gửi đơn xin nghỉ phép:</p>' +
          '<ul>' +
          '<li>Số ngày nghỉ: <b>' + soNgay + ' ngày</b></li>' +
          '<li>Từ ngày: <b>' + ngayBatDau + '</b> đến ngày: <b>' + ngayKetThuc + '</b></li>' +
          '<li>Lý do: ' + (lyDo || '(không ghi)') + '</li>' +
          '</ul>' +
          '<p>Giám đốc vui lòng đăng nhập hệ thống chấm công để duyệt đơn này. Trưởng phòng nhận email này chỉ để nắm thông tin, không có quyền duyệt.</p>' +
          '<p><i>Email tự động, vui lòng không trả lời.</i></p>'
      });
    } catch (err) {
      // Không chặn việc tạo đơn nếu gửi mail lỗi
    }
  }

  return { success: true, message: 'Đã gửi đơn xin nghỉ phép (' + soNgay + ' ngày, từ ' + ngayBatDau + ' đến ' + ngayKetThuc + '). Vui lòng chờ Giám đốc duyệt.' };
}

function layDonChoDuyetChoGiamDoc(token) {
  const session = kiemTraSession(token);
  if (!session.success) return { success: false, message: 'Phiên đăng nhập hết hạn.' };
  if (session.vaiTro !== VAI_TRO.GIAM_DOC && session.vaiTro !== VAI_TRO.ADMIN) {
    return { success: false, message: 'Chỉ Giám đốc mới có quyền duyệt đơn nghỉ phép.' };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DONNGHIPHEP);
  const data = sheet.getDataRange().getValues();
  const ds = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === 'Chờ duyệt') {
      ds.push({
        maDon: data[i][0], email: data[i][1], hoTen: data[i][2], phongBan: data[i][3],
        soNgay: data[i][4], ngayBatDau: chuanHoaNgay(data[i][5]), ngayKetThuc: chuanHoaNgay(data[i][6]), lyDo: data[i][11]
      });
    }
  }
  return { success: true, danhSach: ds };
}

function duyetDonNghiPhep(token, maDon, ketQua) {
  const session = kiemTraSession(token);
  if (!session.success) return { success: false, message: 'Phiên đăng nhập hết hạn.' };
  if (session.vaiTro !== VAI_TRO.GIAM_DOC && session.vaiTro !== VAI_TRO.ADMIN) {
    return { success: false, message: 'Chỉ Giám đốc mới có quyền duyệt đơn nghỉ phép.' };
  }
  if (['Đã duyệt', 'Từ chối'].indexOf(ketQua) === -1) {
    return { success: false, message: 'Kết quả duyệt không hợp lệ.' };
  }

  const don = timDonTheoMa(maDon);
  if (!don) return { success: false, message: 'Không tìm thấy đơn nghỉ phép.' };
  if (don.trangThai !== 'Chờ duyệt') return { success: false, message: 'Đơn này đã được xử lý trước đó (' + don.trangThai + ').' };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DONNGHIPHEP);
  sheet.getRange(don.rowIndex, 9).setValue(ketQua);
  sheet.getRange(don.rowIndex, 10).setValue(session.hoTen);
  sheet.getRange(don.rowIndex, 11).setValue(new Date());

  try {
    MailApp.sendEmail({
      to: don.email,
      subject: 'Kết quả đơn xin nghỉ phép - ' + TEN_CONG_TY,
      htmlBody:
        '<p>Chào <b>' + don.hoTen + '</b>,</p>' +
        '<p>Đơn xin nghỉ phép ' + don.soNgay + ' ngày (từ ' + don.ngayBatDau + ' đến ' + don.ngayKetThuc + ') của bạn đã được ' +
        '<b style="color:' + (ketQua === 'Đã duyệt' ? '#1e8449' : '#c0392b') + '">' + ketQua + '</b> bởi ' + session.hoTen + '.</p>' +
        '<p><i>Email tự động, vui lòng không trả lời.</i></p>'
    });
  } catch (err) {
    // Không chặn việc duyệt nếu gửi mail lỗi
  }

  return { success: true, message: 'Đã ' + ketQua.toLowerCase() + ' đơn của ' + don.hoTen + '.' };
}

/* ============================================================
 *  CHẤM CÔNG
 * ============================================================ */

function layTrangThaiHomNay(token) {
  const session = kiemTraSession(token);
  if (!session.success) return { success: false, message: 'Phiên đăng nhập hết hạn.' };

  const homNay = dinhDangNgay(new Date());
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CHAMCONG);
  const data = sheet.getDataRange().getValues();

  const danhSachGioVao = [], danhSachGioRa = [];
  let nghiKhongLuong = false, nghiKhongLyDo = false;

  for (let i = 1; i < data.length; i++) {
    if (chuanHoaNgay(data[i][0]) === homNay && String(data[i][1]).toLowerCase() === session.email.toLowerCase()) {
      const tt = data[i][4];
      if (tt === TRANG_THAI.GIO_VAO) danhSachGioVao.push({ thoiGian: data[i][3], canhBao: data[i][9] || '' });
      else if (tt === TRANG_THAI.GIO_RA) danhSachGioRa.push({ thoiGian: data[i][3], canhBao: data[i][9] || '' });
      else if (tt === TRANG_THAI.NGHI_KHONG_LUONG) nghiKhongLuong = true;
      else if (tt === TRANG_THAI.NGHI_KHONG_LY_DO) nghiKhongLyDo = true;
    }
  }

  const donDuyet = coDonNghiPhepDuocDuyetHomNay(session.email);
  const dangChoDuyet = coDonChoDuyet(session.email);
  const config = layConfig();

  return {
    success: true,
    danhSachGioVao: danhSachGioVao,
    danhSachGioRa: danhSachGioRa,
    nghiKhongLuong: nghiKhongLuong,
    nghiKhongLyDo: nghiKhongLyDo,
    nghiPhepDuocDuyet: donDuyet,
    dangChoDuyetNghiPhep: dangChoDuyet,
    gioVaoChuan: config.GIO_VAO_CHUAN,
    gioRaChuan: config.GIO_RA_CHUAN,
    bienDoPhut: config.BIEN_DO_PHUT
  };
}

function chamCong(token, action, lat, lng, thietBi) {
  const session = kiemTraSession(token);
  if (!session.success) return { success: false, message: 'Phiên đăng nhập hết hạn, vui lòng đăng nhập lại.' };

  const cacActionHopLe = ['GIO_VAO', 'GIO_RA', 'NGHI_KHONG_LUONG'];
  if (cacActionHopLe.indexOf(action) === -1) {
    return { success: false, message: 'Trạng thái không hợp lệ.' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { success: false, message: 'Hệ thống đang xử lý yêu cầu khác, vui lòng thử lại sau vài giây.' };
  }

  try {
    const donDuyet = coDonNghiPhepDuocDuyetHomNay(session.email);
    if (donDuyet) {
      return { success: false, message: 'Bạn đang trong thời gian nghỉ phép đã được duyệt (từ ' + donDuyet.batDau + ' đến ' + donDuyet.ketThuc + ').' };
    }

    const homNay = dinhDangNgay(new Date());
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CHAMCONG);
    const data = sheet.getDataRange().getValues();

    let coGioVao = false, coGioRa = false, coNghiKhongLuong = false, coNghiKhongLyDo = false;
    for (let i = 1; i < data.length; i++) {
      if (chuanHoaNgay(data[i][0]) === homNay && String(data[i][1]).toLowerCase() === session.email.toLowerCase()) {
        const tt = data[i][4];
        if (tt === TRANG_THAI.GIO_VAO) coGioVao = true;
        else if (tt === TRANG_THAI.GIO_RA) coGioRa = true;
        else if (tt === TRANG_THAI.NGHI_KHONG_LUONG) coNghiKhongLuong = true;
        else if (tt === TRANG_THAI.NGHI_KHONG_LY_DO) coNghiKhongLyDo = true;
      }
    }

    if (coNghiKhongLuong || coNghiKhongLyDo) {
      const ttDaCo = coNghiKhongLuong ? TRANG_THAI.NGHI_KHONG_LUONG : TRANG_THAI.NGHI_KHONG_LY_DO;
      return { success: false, message: 'Bạn đã ghi nhận trạng thái "' + ttDaCo + '" cho hôm nay rồi.' };
    }

    const config = layConfig();
    const bienDo = parseInt(config.BIEN_DO_PHUT, 10) || 15;

    let trangThaiGhi, khoangCach = '', ghiChu = '';

    if (action === 'NGHI_KHONG_LUONG') {
      if (coGioVao || coGioRa) {
        return { success: false, message: 'Bạn đã chấm công đi làm hôm nay rồi, không thể chọn nghỉ.' };
      }
      trangThaiGhi = TRANG_THAI.NGHI_KHONG_LUONG;

    } else if (action === 'GIO_VAO') {
      trangThaiGhi = TRANG_THAI.GIO_VAO;
      var canhBaoVao = [];

      // Không chặn theo vị trí - vẫn cho ghi nhận tự do, chỉ cảnh báo nếu ở xa công ty
      if (lat == null || lng == null) {
        canhBaoVao.push('Không xác định được vị trí GPS');
      } else {
        const d1 = tinhKhoangCachMet(parseFloat(lat), parseFloat(lng), parseFloat(config.LAT), parseFloat(config.LNG));
        khoangCach = Math.round(d1);
        const banKinh = parseFloat(config.RADIUS_M) || 200;
        if (d1 > banKinh) {
          canhBaoVao.push('Cách công ty ' + Math.round(d1) + 'm (ngoài bán kính cho phép ' + banKinh + 'm)');
          guiCanhBaoTucThiXaCongTy(session.hoTen, session.email, session.phongBan, session.chucDanh, TRANG_THAI.GIO_VAO, Math.round(d1), banKinh, dinhDangGio(new Date()));
        }
      }

      // Không chặn nếu chấm trễ - vẫn cho ghi nhận, chỉ cảnh báo số phút trễ
      const phutTre = soPhutTrongNgay(new Date()) - (chuoiGioThanhPhut(config.GIO_VAO_CHUAN) + bienDo);
      if (phutTre > 0) {
        canhBaoVao.push('Vào trễ ' + phutTre + ' phút so với giờ chuẩn (' + config.GIO_VAO_CHUAN + ' ± ' + bienDo + ' phút)');
      }

      ghiChu = canhBaoVao.join('; ');

    } else if (action === 'GIO_RA') {
      trangThaiGhi = TRANG_THAI.GIO_RA;
      var canhBaoRa = [];

      // Không chặn theo vị trí - vẫn cho ghi nhận tự do, chỉ cảnh báo nếu ở xa công ty
      if (lat == null || lng == null) {
        canhBaoRa.push('Không xác định được vị trí GPS');
      } else {
        const d2 = tinhKhoangCachMet(parseFloat(lat), parseFloat(lng), parseFloat(config.LAT), parseFloat(config.LNG));
        khoangCach = Math.round(d2);
        const banKinh2 = parseFloat(config.RADIUS_M) || 200;
        if (d2 > banKinh2) {
          canhBaoRa.push('Cách công ty ' + Math.round(d2) + 'm (ngoài bán kính cho phép ' + banKinh2 + 'm)');
          guiCanhBaoTucThiXaCongTy(session.hoTen, session.email, session.phongBan, session.chucDanh, TRANG_THAI.GIO_RA, Math.round(d2), banKinh2, dinhDangGio(new Date()));
        }
      }

      // Không chặn nếu chấm sớm - vẫn cho ghi nhận, chỉ cảnh báo số phút sớm
      const phutSom = (chuoiGioThanhPhut(config.GIO_RA_CHUAN) - bienDo) - soPhutTrongNgay(new Date());
      if (phutSom > 0) {
        canhBaoRa.push('Ra sớm ' + phutSom + ' phút so với giờ chuẩn (' + config.GIO_RA_CHUAN + ' ± ' + bienDo + ' phút)');
      }

      ghiChu = canhBaoRa.join('; ');
    }

    const now = new Date();
    sheet.appendRow([
      homNay,
      session.email,
      session.hoTen,
      dinhDangGio(now),
      trangThaiGhi,
      lat || '',
      lng || '',
      khoangCach,
      thietBi || '',
      ghiChu
    ]);

    // Ghi lại thiết bị dùng để chấm công vào tab riêng, phục vụ báo cáo đa thiết bị hàng tuần
    ghiLichSuThietBi(session.email, session.hoTen, trangThaiGhi, thietBi, lat, lng);

    const thongBaoDonGian = {
      'Giờ vào': 'Đã chấm công giờ vào',
      'Giờ ra': 'Đã chấm công giờ ra',
      'Nghỉ không lương': 'Đã chấm nghỉ không lương'
    };

    return {
      success: true,
      message: thongBaoDonGian[trangThaiGhi] || ('Đã ghi nhận: ' + trangThaiGhi),
      trangThai: trangThaiGhi
    };
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 *  BÁO CÁO CHO ADMIN / GIÁM ĐỐC / TRƯỞNG PHÒNG
 * ============================================================ */

function layBaoCaoHomNay(token) {
  const session = kiemTraSession(token);
  if (!session.success) return { success: false, message: 'Phiên đăng nhập hết hạn.' };
  if ([VAI_TRO.ADMIN, VAI_TRO.GIAM_DOC, VAI_TRO.TRUONG_PHONG].indexOf(session.vaiTro) === -1) {
    return { success: false, message: 'Bạn không có quyền xem báo cáo.' };
  }

  const bc = xayDungBaoCaoNgay(new Date());

  if (session.vaiTro === VAI_TRO.TRUONG_PHONG) {
    bc.danhSach = bc.danhSach.filter(function (row) { return row.phongBan === session.phongBan; });
  }

  return { success: true, baoCao: bc };
}

function xayDungBaoCaoNgay(ngay) {
  const homNay = dinhDangNgay(ngay);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CHAMCONG);
  const data = sheet.getDataRange().getValues();

  const tongHop = {};
  for (let i = 1; i < data.length; i++) {
    if (chuanHoaNgay(data[i][0]) === homNay) {
      const key = String(data[i][1]).toLowerCase();
      if (!tongHop[key]) tongHop[key] = { dsGioVao: [], dsGioRa: [] };
      const tt = data[i][4];
      if (tt === TRANG_THAI.GIO_VAO) tongHop[key].dsGioVao.push({ thoiGian: data[i][3], canhBao: data[i][9] || '' });
      else if (tt === TRANG_THAI.GIO_RA) tongHop[key].dsGioRa.push({ thoiGian: data[i][3], canhBao: data[i][9] || '' });
      else if (tt === TRANG_THAI.NGHI_KHONG_LUONG) tongHop[key].nghiKhongLuong = true;
      else if (tt === TRANG_THAI.NGHI_KHONG_LY_DO) tongHop[key].nghiKhongLyDo = true;
    }
  }

  const tatCaUser = layDanhSachUserHopLe();
  const ketQua = [];
  tatCaUser.forEach(function (u) {
    const key = String(u.email).toLowerCase();
    const th = tongHop[key];
    const donDuyet = coDonNghiPhepDuocDuyetHomNay(u.email);
    let trangThai, thoiGian;

    const coGioVao = th && th.dsGioVao.length > 0;
    const coGioRa = th && th.dsGioRa.length > 0;

    function dinhDangDanhSach(ds) {
      // Luôn hiện số lần chấm; nếu >1 lần thì thêm lần đầu → lần cuối, kèm cảnh báo nếu có
      if (ds.length === 1) {
        return ds[0].thoiGian + (ds[0].canhBao ? ' [' + ds[0].canhBao + ']' : '') + ' (1 lần)';
      }
      const dau = ds[0];
      const cuoi = ds[ds.length - 1];
      return dau.thoiGian + (dau.canhBao ? ' [' + dau.canhBao + ']' : '') +
        ' → ' + cuoi.thoiGian + (cuoi.canhBao ? ' [' + cuoi.canhBao + ']' : '') + ' (' + ds.length + ' lần)';
    }

    if (donDuyet) {
      trangThai = 'Nghỉ phép (đã duyệt)';
      thoiGian = donDuyet.batDau + ' - ' + donDuyet.ketThuc;
    } else if (!th) {
      trangThai = TRANG_THAI.NGHI_KHONG_LY_DO;
      thoiGian = '';
    } else if (th.nghiKhongLuong) {
      trangThai = TRANG_THAI.NGHI_KHONG_LUONG;
      thoiGian = '';
    } else if (coGioVao || coGioRa) {
      trangThai = 'Đã chấm công';
      const phanVao = 'Vào: ' + (coGioVao ? dinhDangDanhSach(th.dsGioVao) : 'chưa chấm');
      const phanRa = 'Ra: ' + (coGioRa ? dinhDangDanhSach(th.dsGioRa) : 'chưa chấm');
      thoiGian = phanVao + ' - ' + phanRa;
    } else {
      trangThai = TRANG_THAI.NGHI_KHONG_LY_DO;
      thoiGian = '';
    }

    ketQua.push({ email: u.email, hoTen: u.hoTen, phongBan: u.phongBan, chucDanh: u.chucDanh, trangThai: trangThai, thoiGian: thoiGian });
  });

  return { ngay: homNay, danhSach: ketQua };
}

/**
 * Gửi email cảnh báo NGAY LẬP TỨC cho Giám đốc + admin khi 1 nhân viên chấm
 * giờ vào/giờ ra ở vị trí cách xa công ty ngoài bán kính cho phép.
 * Có thể tắt bằng Config!THONG_BAO_TUC_THI_XA_CONG_TY = 'TAT'.
 */
/**
 * Gửi cảnh báo NGAY LẬP TỨC cho Giám đốc + admin khi 1 nhân viên chấm giờ
 * vào/giờ ra ở vị trí cách xa công ty ngoài bán kính cho phép.
 * - Nếu Config!GOOGLE_CHAT_WEBHOOK_URL đã được cấu hình -> gửi qua Google Chat.
 * - Nếu chưa cấu hình (hoặc gửi Google Chat lỗi) -> tự động gửi email dự phòng.
 * Có thể tắt hoàn toàn bằng Config!THONG_BAO_TUC_THI_XA_CONG_TY = 'TAT'.
 */
function guiCanhBaoTucThiXaCongTy(hoTen, email, phongBan, chucDanh, trangThai, khoangCachM, banKinhM, thoiGianStr) {
  const config = layConfig();
  if (String(config.THONG_BAO_TUC_THI_XA_CONG_TY || 'BAT').toUpperCase() === 'TAT') return;

  const webhookUrl = String(config.GOOGLE_CHAT_WEBHOOK_URL || '').trim();
  if (webhookUrl) {
    const daGuiChat = guiCanhBaoQuaGoogleChat(webhookUrl, hoTen, email, phongBan, chucDanh, trangThai, khoangCachM, banKinhM, thoiGianStr);
    if (daGuiChat) return; // Gửi Google Chat thành công -> không cần gửi thêm email
  }

  // Chưa cấu hình webhook, hoặc gửi Google Chat thất bại -> gửi email dự phòng để không mất cảnh báo
  guiCanhBaoQuaEmail(hoTen, email, phongBan, chucDanh, trangThai, khoangCachM, banKinhM, thoiGianStr);
}

function guiCanhBaoQuaGoogleChat(webhookUrl, hoTen, email, phongBan, chucDanh, trangThai, khoangCachM, banKinhM, thoiGianStr) {
  const dongThongTin = [
    '⚠️ *Chấm công cách xa công ty* - ' + TEN_CONG_TY,
    'Nhân viên: *' + hoTen + '* (' + email + ')'
  ];
  if (phongBan) dongThongTin.push('Phòng ban: ' + phongBan);
  if (chucDanh) dongThongTin.push('Chức danh: ' + chucDanh);
  dongThongTin.push('Hành động: *' + trangThai + '* lúc ' + thoiGianStr);
  dongThongTin.push('Khoảng cách: *' + khoangCachM + 'm* (cho phép tối đa ' + banKinhM + 'm)');

  try {
    const res = UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: dongThongTin.join('\n') }),
      muteHttpExceptions: true
    });
    // Mã 200 nghĩa là Google Chat đã nhận tin nhắn thành công
    return res.getResponseCode() === 200;
  } catch (err) {
    return false;
  }
}

/**
 * Chạy hàm này (menu ⚙️ Chấm công Hội Vật Lý > Kiểm tra webhook Google Chat) sau khi
 * đã dán URL vào Config!GOOGLE_CHAT_WEBHOOK_URL, để kiểm tra webhook có hoạt động
 * không mà không cần chờ có ai chấm công cách xa công ty.
 */
function kiemTraWebhookGoogleChat() {
  const config = layConfig();
  const webhookUrl = String(config.GOOGLE_CHAT_WEBHOOK_URL || '').trim();

  if (!webhookUrl) {
    SpreadsheetApp.getUi().alert('Chưa cấu hình GOOGLE_CHAT_WEBHOOK_URL trong sheet Config. Vui lòng dán URL webhook vào đó trước rồi thử lại.');
    return;
  }

  const thanhCong = guiCanhBaoQuaGoogleChat(
    webhookUrl,
    'Nhân viên Test',
    'test@gmail.com',
    'Phòng Test',
    'Chức danh Test',
    'Giờ vào',
    999,
    200,
    dinhDangGio(new Date())
  );

  if (thanhCong) {
    SpreadsheetApp.getUi().alert('Đã gửi tin nhắn test thành công! Vào Google Chat Space kiểm tra xem đã nhận được tin nhắn "⚠️ Chấm công cách xa công ty..." chưa.');
  } else {
    SpreadsheetApp.getUi().alert('Gửi thất bại. Kiểm tra lại URL webhook đã copy đúng và đầy đủ chưa (thường rất dài, cần copy hết cả đoạn key=...&token=...).');
  }
}

function guiCanhBaoQuaEmail(hoTen, email, phongBan, chucDanh, trangThai, khoangCachM, banKinhM, thoiGianStr) {
  const config = layConfig();
  const giamDocEmails = layEmailGiamDoc();
  const adminEmails = String(config.ADMIN_EMAILS || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  const nguoiNhan = giamDocEmails.concat(adminEmails).filter(function (e, idx, arr) { return arr.indexOf(e) === idx && e; });
  if (nguoiNhan.length === 0) return;

  try {
    MailApp.sendEmail({
      to: nguoiNhan.join(','),
      subject: '⚠️ Chấm công cách xa công ty - ' + hoTen + ' - ' + TEN_CONG_TY,
      htmlBody:
        '<p>Nhân viên <b>' + hoTen + '</b> (' + email + ')' +
        (phongBan ? ', phòng ban <b>' + phongBan + '</b>' : '') +
        (chucDanh ? ', chức danh <b>' + chucDanh + '</b>' : '') +
        ' vừa chấm <b>' + trangThai + '</b> lúc <b>' + thoiGianStr + '</b>,</p>' +
        '<p>nhưng đang ở vị trí cách công ty <b style="color:#c0392b">' + khoangCachM + 'm</b> ' +
        '(bán kính cho phép: ' + banKinhM + 'm).</p>' +
        '<p><i>Email tự động gửi ngay khi phát hiện, vui lòng không trả lời.</i></p>'
    });
  } catch (err) {
    // Không chặn việc chấm công nếu gửi mail cảnh báo lỗi
  }
}

function guiBaoCaoHangNgay() {
  const homNay = new Date();
  const homNayStr = dinhDangNgay(homNay);
  const bc = xayDungBaoCaoNgay(homNay);
  const sheetChamCong = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CHAMCONG);

  bc.danhSach.forEach(function (row) {
    if (row.trangThai === TRANG_THAI.NGHI_KHONG_LY_DO) {
      sheetChamCong.appendRow([homNayStr, row.email, row.hoTen, '', TRANG_THAI.NGHI_KHONG_LY_DO, '', '', '', '', 'Tự động - không chấm công']);
    }
  });

  let hangHtml = '';
  bc.danhSach.forEach(function (row) {
    const mau = row.trangThai === 'Đã chấm công' ? '#27ae60' :
      row.trangThai === TRANG_THAI.NGHI_KHONG_LY_DO ? '#c0392b' : '#e67e22';
    hangHtml += '<tr>' +
      '<td style="padding:6px 10px;border:1px solid #ddd">' + row.hoTen + '</td>' +
      '<td style="padding:6px 10px;border:1px solid #ddd">' + row.email + '</td>' +
      '<td style="padding:6px 10px;border:1px solid #ddd">' + (row.chucDanh || '-') + '</td>' +
      '<td style="padding:6px 10px;border:1px solid #ddd">' + (row.phongBan || '-') + '</td>' +
      '<td style="padding:6px 10px;border:1px solid #ddd;color:' + mau + ';font-weight:bold">' + row.trangThai + '</td>' +
      '<td style="padding:6px 10px;border:1px solid #ddd">' + (row.thoiGian || '-') + '</td>' +
      '</tr>';
  });

  const donCho = layDanhSachDonChoDuyetTho();
  let choDuyetHtml = '';
  if (donCho.length > 0) {
    choDuyetHtml = '<h3 style="color:#c0392b">Đơn nghỉ phép đang chờ duyệt (' + donCho.length + ')</h3><ul>';
    donCho.forEach(function (d) {
      choDuyetHtml += '<li>' + d.hoTen + ' - ' + d.soNgay + ' ngày (từ ' + d.ngayBatDau + ' đến ' + d.ngayKetThuc + ')</li>';
    });
    choDuyetHtml += '</ul>';
  }

  const html =
    '<h2>Báo cáo chấm công ngày ' + homNayStr + ' - ' + TEN_CONG_TY + '</h2>' +
    '<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px">' +
    '<tr style="background:#2c3e50;color:#fff">' +
    '<th style="padding:6px 10px;border:1px solid #ddd">Họ tên</th>' +
    '<th style="padding:6px 10px;border:1px solid #ddd">Email</th>' +
    '<th style="padding:6px 10px;border:1px solid #ddd">Chức danh</th>' +
    '<th style="padding:6px 10px;border:1px solid #ddd">Phòng ban</th>' +
    '<th style="padding:6px 10px;border:1px solid #ddd">Trạng thái</th>' +
    '<th style="padding:6px 10px;border:1px solid #ddd">Thời gian</th>' +
    '</tr>' + hangHtml + '</table>' + choDuyetHtml;

  const config = layConfig();
  const adminEmails = String(config.ADMIN_EMAILS || '').split(',').map(function (s) { return s.trim(); }).filter(String);
  const giamDocEmails = layEmailGiamDoc();
  const nguoiNhan = adminEmails.concat(giamDocEmails).filter(function (email, idx, arr) { return arr.indexOf(email) === idx && email; });

  if (nguoiNhan.length > 0) {
    MailApp.sendEmail({
      to: nguoiNhan.join(','),
      subject: 'Báo cáo chấm công ' + homNayStr + ' - ' + TEN_CONG_TY,
      htmlBody: html
    });
  }
}

function layDanhSachDonChoDuyetTho() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DONNGHIPHEP);
  const data = sheet.getDataRange().getValues();
  const ds = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === 'Chờ duyệt') {
      ds.push({ hoTen: data[i][2], soNgay: data[i][4], ngayBatDau: chuanHoaNgay(data[i][5]), ngayKetThuc: chuanHoaNgay(data[i][6]) });
    }
  }
  return ds;
}

/* ============================================================
 *  BÁO CÁO ĐA THIẾT BỊ HÀNG TUẦN (cho admin & Giám đốc)
 * ============================================================ */

/**
 * Phân tích sheet LichSuThietBi trong khoảng ngày [ngayBatDauStr, ngayKetThucStr]
 * (dạng yyyy-MM-dd), trả về danh sách nhân viên đã dùng từ 2 loại thiết bị
 * khác nhau trở lên để chấm công trong khoảng thời gian đó.
 */
function xayDungBaoCaoThietBiTuan(ngayBatDauStr, ngayKetThucStr) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LICHSUTHIETBI);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();

  // email -> { hoTen, thietBiCount: { 'iPhone - Safari': 3, 'Windows - Chrome': 1 } }
  const tongHop = {};
  for (let i = 1; i < data.length; i++) {
    const ngay = chuanHoaNgay(data[i][0]);
    if (ngay < ngayBatDauStr || ngay > ngayKetThucStr) continue;

    const email = String(data[i][2]).toLowerCase();
    const hoTen = data[i][3];
    const thietBiRutGon = data[i][6] || 'Không xác định';

    if (!tongHop[email]) tongHop[email] = { hoTen: hoTen, thietBiCount: {} };
    tongHop[email].thietBiCount[thietBiRutGon] = (tongHop[email].thietBiCount[thietBiRutGon] || 0) + 1;
  }

  const ketQua = [];
  Object.keys(tongHop).forEach(function (email) {
    const thongTin = tongHop[email];
    const danhSachThietBi = Object.keys(thongTin.thietBiCount);
    if (danhSachThietBi.length > 1) {
      ketQua.push({
        email: email,
        hoTen: thongTin.hoTen,
        soThietBi: danhSachThietBi.length,
        chiTiet: danhSachThietBi.map(function (tb) { return tb + ' (' + thongTin.thietBiCount[tb] + ' lần)'; }).join('; ')
      });
    }
  });

  // Xếp theo số thiết bị nhiều nhất lên đầu
  ketQua.sort(function (a, b) { return b.soThietBi - a.soThietBi; });
  return ketQua;
}

/**
 * Trigger chạy hàng tuần (Config!NGAY_BAO_CAO_THIET_BI, GIO_BAO_CAO_THIET_BI):
 * Tổng hợp và gửi email cho admin + Giám đốc danh sách nhân viên đã dùng
 * nhiều hơn 1 loại thiết bị khác nhau để chấm công trong 7 ngày gần nhất.
 */
function guiBaoCaoThietBiHangTuan() {
  const homNay = new Date();
  const ngayKetThuc = new Date(homNay);
  ngayKetThuc.setDate(ngayKetThuc.getDate() - 1); // đến hết hôm qua
  const ngayBatDau = new Date(ngayKetThuc);
  ngayBatDau.setDate(ngayBatDau.getDate() - 6); // 7 ngày gần nhất

  const ngayBatDauStr = dinhDangNgay(ngayBatDau);
  const ngayKetThucStr = dinhDangNgay(ngayKetThuc);

  const danhSach = xayDungBaoCaoThietBiTuan(ngayBatDauStr, ngayKetThucStr);

  let noiDung;
  if (danhSach.length === 0) {
    noiDung = '<p style="color:#1e8449">Không phát hiện trường hợp nào dùng nhiều thiết bị khác nhau để chấm công trong tuần qua.</p>';
  } else {
    let hangHtml = '';
    danhSach.forEach(function (row) {
      hangHtml += '<tr>' +
        '<td style="padding:6px 10px;border:1px solid #ddd">' + row.hoTen + '</td>' +
        '<td style="padding:6px 10px;border:1px solid #ddd">' + row.email + '</td>' +
        '<td style="padding:6px 10px;border:1px solid #ddd;color:#c0392b;font-weight:bold">' + row.soThietBi + '</td>' +
        '<td style="padding:6px 10px;border:1px solid #ddd">' + row.chiTiet + '</td>' +
        '</tr>';
    });
    noiDung =
      '<p>Phát hiện <b>' + danhSach.length + '</b> trường hợp dùng nhiều hơn 1 thiết bị khác nhau để chấm công trong tuần qua:</p>' +
      '<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px">' +
      '<tr style="background:#2c3e50;color:#fff">' +
      '<th style="padding:6px 10px;border:1px solid #ddd">Họ tên</th>' +
      '<th style="padding:6px 10px;border:1px solid #ddd">Email</th>' +
      '<th style="padding:6px 10px;border:1px solid #ddd">Số thiết bị</th>' +
      '<th style="padding:6px 10px;border:1px solid #ddd">Chi tiết</th>' +
      '</tr>' + hangHtml + '</table>' +
      '<p style="font-size:12px;color:#7f8c8d">Lưu ý: thiết bị được nhận diện qua User-Agent trình duyệt, chỉ mang tính tham khảo, ' +
      'không phải ID phần cứng và có thể không tuyệt đối chính xác (ví dụ 1 người dùng cả điện thoại lẫn máy tính là bình thường, ' +
      'không nhất thiết là gian lận).</p>';
  }

  const html =
    '<h2>Báo cáo dùng nhiều thiết bị chấm công (' + ngayBatDauStr + ' đến ' + ngayKetThucStr + ') - ' + TEN_CONG_TY + '</h2>' +
    noiDung;

  const config = layConfig();
  const adminEmails = String(config.ADMIN_EMAILS || '').split(',').map(function (s) { return s.trim(); }).filter(String);
  const giamDocEmails = layEmailGiamDoc();
  const nguoiNhan = adminEmails.concat(giamDocEmails).filter(function (email, idx, arr) { return arr.indexOf(email) === idx && email; });

  if (nguoiNhan.length > 0) {
    MailApp.sendEmail({
      to: nguoiNhan.join(','),
      subject: 'Báo cáo đa thiết bị chấm công tuần ' + ngayBatDauStr + ' - ' + ngayKetThucStr + ' - ' + TEN_CONG_TY,
      htmlBody: html
    });
  }
}
