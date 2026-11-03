/* =====================================================================
   BIÊN BẢN BÀN GIAO — Quản Lý Dự án Stellar
   Tệp riêng, dán thêm vào Apps Script (cạnh Code.gs). Không cần sửa Code.gs.

   Cài lần đầu: chọn hàm  caiBienBan  ở thanh công cụ → Chạy → cấp quyền.
   Hàm này tạo 4 tab (BB_BienBan, BB_ChiTiet, BB_DanhBa, BB_MauND),
   thư mục Drive "QLDA_BienBan" và tệp mẫu Google Docs "MAU_BIEN_BAN_BAN_GIAO".
   Chạy lại không sao: cái gì có rồi thì giữ nguyên.

   Sửa mẫu in: mở tệp MAU_BIEN_BAN_BAN_GIAO trong thư mục QLDA_BienBan,
   sửa chữ, căn lề, logo tùy ý; chỉ cần giữ nguyên các mã {{...}}.
   ===================================================================== */

var BB_CONG_TY = 'CÔNG TY CỔ PHẦN STELLAR';   // tên đơn vị in ở góc trái biên bản — sửa cho đúng tên pháp nhân
var BB_DIA_DIEM = 'Hà Nội';                     // địa điểm mặc định
var BB_TAB = { bb: 'BB_BienBan', ct: 'BB_ChiTiet', db: 'BB_DanhBa', mau: 'BB_MauND' };
var BB_COLS = {
  BB_BienBan: [
    ['so_bb', 'Số biên bản, tự sinh: BBBG-<mã dự án>-<số thứ tự>'],
    ['ngay', 'Ngày bàn giao (yyyy-mm-dd)'],
    ['ma_da', 'Mã dự án (DuAn.ma)'],
    ['ten_da', 'Tên dự án lúc lập'],
    ['loai', 'Loại bàn giao / tiêu đề phụ'],
    ['ben_giao', 'Đơn vị bên giao'],
    ['nguoi_giao', 'Người giao'],
    ['cv_giao', 'Chức vụ người giao'],
    ['ben_nhan', 'Đơn vị bên nhận'],
    ['nguoi_nhan', 'Người nhận'],
    ['cv_nhan', 'Chức vụ người nhận'],
    ['dia_diem', 'Địa điểm bàn giao'],
    ['noi_dung', 'Nội dung chung'],
    ['ghi_chu', 'Ghi chú cuối biên bản'],
    ['trang_thai', 'Đã lập / Đã ký / Hủy'],
    ['so_muc', 'Số dòng danh mục'],
    ['doc_id', 'ID bản Google Docs'],
    ['pdf_id', 'ID tệp PDF'],
    ['docx_id', 'ID tệp Word'],
    ['nguoi_lap', 'Người lập trên app'],
    ['thoi_diem', 'Thời điểm lưu gần nhất']
  ],
  BB_ChiTiet: [
    ['so_bb', 'Số biên bản'],
    ['stt', 'Số thứ tự dòng'],
    ['cv_id', 'ID hạng mục (CongViec.id), trống nếu nhập tay'],
    ['ten', 'Tên hồ sơ / sản phẩm'],
    ['hinh_thuc', 'Bản cứng / Bản mềm / Bản cứng + bản mềm'],
    ['so_luong', 'Số lượng'],
    ['don_vi', 'Đơn vị tính'],
    ['ghi_chu', 'Ghi chú dòng']
  ],
  BB_DanhBa: [
    ['ten', 'Họ tên (hiện trong danh sách chọn)'],
    ['chuc_vu', 'Chức vụ'],
    ['don_vi', 'Đơn vị / công ty'],
    ['vai', 'Giao hoặc Nhận (hoặc Cả hai)'],
    ['sdt', 'Điện thoại'],
    ['email', 'Email'],
    ['kich_hoat', 'Có / Không']
  ],
  BB_MauND: [
    ['ten_mau', 'Tên mẫu (hiện trong danh sách chọn)'],
    ['loai', 'Tiêu đề phụ in dưới chữ BIÊN BẢN BÀN GIAO'],
    ['noi_dung', 'Đoạn nội dung chung'],
    ['hinh_thuc', 'Hình thức mặc định cho các dòng danh mục'],
    ['ghi_chu', 'Ghi chú cuối biên bản mặc định']
  ]
};
var BB_MAU_MAC_DINH = [
  ['Bàn giao hồ sơ cho CĐT', 'Hồ sơ thiết kế', 'Bên giao bàn giao cho Bên nhận các hồ sơ, sản phẩm thuộc dự án theo danh mục dưới đây.', 'Bản cứng + bản mềm', 'Bên nhận đã kiểm tra số lượng hồ sơ theo danh mục. Biên bản lập thành 02 bản, mỗi bên giữ 01 bản có giá trị như nhau.'],
  ['Bàn giao nội bộ', 'Hồ sơ nội bộ', 'Bàn giao hồ sơ, tài liệu giữa các bộ phận để tiếp tục triển khai dự án theo danh mục dưới đây.', 'Bản mềm', 'Biên bản lập thành 02 bản, mỗi bên giữ 01 bản.'],
  ['Bàn giao bản in', 'Hồ sơ in ấn', 'Bàn giao các bộ hồ sơ đã in theo danh mục dưới đây.', 'Bản cứng', 'Bên nhận đã kiểm tra số bộ, số trang. Biên bản lập thành 02 bản, mỗi bên giữ 01 bản.']
];

/* ---------------- CÀI ĐẶT ---------------- */
function caiBienBan() {
  var ss = bbSs_(), msg = [];
  Object.keys(BB_COLS).forEach(function (name) {
    var s = ss.getSheetByName(name), cols = BB_COLS[name];
    if (s) { msg.push('Giữ tab ' + name); return; }
    s = ss.insertSheet(name);
    s.getRange(1, 1, 2, cols.length).setValues([cols.map(function (c) { return c[0]; }), cols.map(function (c) { return c[1]; })]);
    s.getRange(1, 1, 1, cols.length).setFontWeight('bold').setBackground('#12233F').setFontColor('#FFFFFF');
    s.getRange(2, 1, 1, cols.length).setFontStyle('italic').setFontColor('#5A6A7C');
    s.getRange(3, 1, 998, cols.length).setNumberFormat('@');
    s.setFrozenRows(2);
    msg.push('Tạo tab ' + name);
    if (name === BB_TAB.mau) bbAppend_(name, BB_MAU_MAC_DINH.map(function (r) { return { ten_mau: r[0], loai: r[1], noi_dung: r[2], hinh_thuc: r[3], ghi_chu: r[4] }; }));
    if (name === BB_TAB.db) {
      var ns = bbRead_('NhanSu').filter(function (p) { return p.kich_hoat === 'Có' && String(p.ten).charAt(0) !== '('; });
      bbAppend_(name, ns.map(function (p) { return { ten: p.ten, chuc_vu: '', don_vi: BB_CONG_TY, vai: 'Giao', kich_hoat: 'Có' }; }));
    }
  });
  var props = PropertiesService.getScriptProperties();
  var folder = bbFolder_();
  msg.push('Thư mục lưu: ' + folder.getUrl());
  var tid = props.getProperty('BB_TEMPLATE_ID'), ok = false;
  if (tid) { try { ok = !DriveApp.getFileById(tid).isTrashed(); } catch (e) { ok = false; } }
  if (!ok) {
    tid = bbTaoMau_(folder);
    props.setProperty('BB_TEMPLATE_ID', tid);
    msg.push('Tạo mẫu: https://docs.google.com/document/d/' + tid + '/edit');
  } else msg.push('Giữ mẫu: https://docs.google.com/document/d/' + tid + '/edit');
  Logger.log(msg.join('\n'));
  return msg.join('\n');
}

/* Dùng mẫu Google Docs khác (ví dụ mẫu công ty có sẵn): dán ID tệp vào giữa hai dấu nháy rồi chạy hàm này. */
function doiMauBienBan() {
  var id = '';
  if (!id) throw new Error('Dán ID tệp Google Docs mẫu vào biến id trong hàm doiMauBienBan trước khi chạy.');
  DocumentApp.openById(id);
  PropertiesService.getScriptProperties().setProperty('BB_TEMPLATE_ID', id);
  Logger.log('Đã đổi mẫu biên bản sang ' + id);
}

function bbSs_() { return sh_('DuAn').getParent(); }

function bbFolder_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('BB_FOLDER_ID');
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  var parents = DriveApp.getFileById(bbSs_().getId()).getParents();
  var base = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var it = base.getFoldersByName('QLDA_BienBan');
  var f2 = it.hasNext() ? it.next() : base.createFolder('QLDA_BienBan');
  props.setProperty('BB_FOLDER_ID', f2.getId());
  return f2;
}

function bbSty_(p, o) {
  var t = p.editAsText(), A = DocumentApp.HorizontalAlignment;
  if (o.al) p.setAlignment(A[o.al]);
  if (o.sb !== undefined) p.setSpacingBefore(o.sb);
  if (o.sa !== undefined) p.setSpacingAfter(o.sa);
  if (t.getText().length) {
    if (o.b !== undefined) t.setBold(o.b);
    if (o.i !== undefined) t.setItalic(o.i);
    if (o.fs) t.setFontSize(o.fs);
  }
  return p;
}
/* Ghi nhiều dòng vào một ô bảng, mỗi dòng một đoạn; st(i) trả định dạng cho dòng i */
function bbCellLines_(cell, lines, st) {
  cell.getChild(0).asParagraph().setText(lines[0]);
  for (var i = 1; i < lines.length; i++) cell.appendParagraph(lines[i]);
  for (var j = 0; j < cell.getNumChildren(); j++) bbSty_(cell.getChild(j).asParagraph(), st(j, cell.getNumChildren()));
}

function bbTaoMau_(folder) {
  var doc = DocumentApp.create('MAU_BIEN_BAN_BAN_GIAO'), b = doc.getBody();
  b.setMarginLeft(85).setMarginRight(57).setMarginTop(57).setMarginBottom(57); // 3 cm trái, 2 cm còn lại
  var head = b.appendTable([['', '']]);
  head.setBorderWidth(0); head.setColumnWidth(0, 210);
  bbCellLines_(head.getCell(0, 0), ['{{CONG_TY}}', 'Số: {{SO_BB}}'], function (i) { return { al: 'CENTER', b: i === 0, fs: 12 }; });
  bbCellLines_(head.getCell(0, 1), ['CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', 'Độc lập - Tự do - Hạnh phúc'], function () { return { al: 'CENTER', b: true, fs: 12 }; });
  bbSty_(b.appendParagraph('{{DIA_DIEM}}, {{NGAY_CHU}}'), { al: 'RIGHT', i: true, b: false, sb: 8 });
  bbSty_(b.appendParagraph('BIÊN BẢN BÀN GIAO'), { al: 'CENTER', b: true, i: false, fs: 15, sb: 14 });
  bbSty_(b.appendParagraph('{{LOAI}}'), { al: 'CENTER', b: true, i: false, fs: 13 });
  bbSty_(b.appendParagraph('Dự án: {{DU_AN}}'), { al: 'CENTER', b: false, i: false, fs: 13, sa: 10 });
  var plain = function (txt, bold) { return bbSty_(b.appendParagraph(txt), { al: 'JUSTIFY', b: !!bold, i: false, fs: 13, sa: 4, sb: 0 }); };
  plain('Hôm nay, {{NGAY_CHU}}, tại {{DIA_DIEM}}, chúng tôi gồm:');
  plain('BÊN GIAO: {{BEN_GIAO}}', true);
  plain('Ông/Bà: {{NGUOI_GIAO}}          Chức vụ: {{CV_GIAO}}');
  plain('BÊN NHẬN: {{BEN_NHAN}}', true);
  plain('Ông/Bà: {{NGUOI_NHAN}}          Chức vụ: {{CV_NHAN}}');
  plain('{{NOI_DUNG}}');
  var tb = b.appendTable([['STT', 'Tên hồ sơ, sản phẩm', 'Hình thức', 'Số lượng', 'Ghi chú'], ['{{#STT}}', '{{#TEN}}', '{{#HINH_THUC}}', '{{#SO_LUONG}}', '{{#GHI_CHU}}']]);
  tb.setBorderWidth(1);
  [36, 214, 92, 70, 92].forEach(function (w, i) { tb.setColumnWidth(i, w); });
  for (var r = 0; r < 2; r++) for (var k = 0; k < 5; k++) {
    bbSty_(tb.getCell(r, k).getChild(0).asParagraph(), { al: (k === 1 || k === 4) && r === 1 ? 'LEFT' : 'CENTER', b: r === 0, i: false, fs: 12 });
    if (r === 0) tb.getCell(r, k).setBackgroundColor('#EAEEF3');
  }
  plain('Tổng cộng: {{SO_MUC}} mục.').setSpacingBefore(6);
  plain('{{GHI_CHU}}');
  var sign = b.appendTable([['', '']]);
  sign.setBorderWidth(0);
  [['ĐẠI DIỆN BÊN GIAO', '{{NGUOI_GIAO}}'], ['ĐẠI DIỆN BÊN NHẬN', '{{NGUOI_NHAN}}']].forEach(function (x, c) {
    bbCellLines_(sign.getCell(0, c), [x[0], '(Ký, ghi rõ họ tên)', ' ', ' ', ' ', ' ', x[1]], function (i, n) { return { al: 'CENTER', b: i === 0 || i === n - 1, i: i === 1, fs: 13 }; });
  });
  // đoạn trống mặc định ở đầu thân văn bản
  var first = b.getChild(0);
  if (first.getType() === DocumentApp.ElementType.PARAGRAPH && !first.asParagraph().getText()) first.removeFromParent();
  b.editAsText().setFontFamily('Times New Roman');
  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(folder);
  return doc.getId();
}

/* ---------------- ĐỌC / GHI BẢNG ---------------- */
function bbSheet_(name) {
  var s = bbSs_().getSheetByName(name);
  if (!s) throw new Error('Chưa cài chức năng biên bản. Mở Apps Script, chạy hàm caiBienBan một lần.');
  return s;
}
function bbHeaders_(s) { return s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); }); }
function bbCell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : v;
}
function bbRead_(name) {
  var s = bbSheet_(name), last = s.getLastRow();
  if (last < 3) return [];
  var h = bbHeaders_(s), vals = s.getRange(3, 1, last - 2, h.length).getValues(), out = [];
  vals.forEach(function (r, i) {
    if (r.every(function (x) { return x === '' || x === null; })) return;
    var o = { _row: i + 3 };
    h.forEach(function (k, j) { if (k) o[k] = bbCell_(r[j]); });
    out.push(o);
  });
  return out;
}
function bbAppend_(name, objs) {
  if (!objs.length) return;
  var s = bbSheet_(name), h = bbHeaders_(s), start = Math.max(s.getLastRow() + 1, 3);
  var rng = s.getRange(start, 1, objs.length, h.length);
  rng.setNumberFormat('@');
  rng.setValues(objs.map(function (o) { return h.map(function (k) { return o[k] === undefined || o[k] === null ? '' : String(o[k]); }); }));
}
function bbWriteRow_(name, row, obj) {
  var s = bbSheet_(name), h = bbHeaders_(s), cur = s.getRange(row, 1, 1, h.length).getValues()[0];
  var vals = h.map(function (k, j) { return obj[k] === undefined ? cur[j] : (obj[k] === null ? '' : String(obj[k])); });
  s.getRange(row, 1, 1, h.length).setNumberFormat('@').setValues([vals]);
}
function bbClean_(o) { var c = {}; Object.keys(o).forEach(function (k) { if (k.charAt(0) !== '_') c[k] = o[k]; }); return c; }
function bbIsReview_(u) { return typeof ROLE_REVIEW !== 'undefined' ? ROLE_REVIEW.indexOf(u.vai_tro) >= 0 : (u.vai_tro === 'Trưởng' || u.vai_tro === 'Duyệt'); }
function bbNow_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"); }
function bbStr_(v, max) { return String(v === null || v === undefined ? '' : v).trim().slice(0, max || 500); }

/* ---------------- API ---------------- */
function api_bbInit(token) {
  auth_(token);
  var items = bbRead_(BB_TAB.ct).map(bbClean_), by = {};
  items.forEach(function (it) { (by[it.so_bb] = by[it.so_bb] || []).push(it); });
  var list = bbRead_(BB_TAB.bb).map(function (r) {
    var o = bbClean_(r);
    o.items = (by[o.so_bb] || []).sort(function (a, b) { return Number(a.stt) - Number(b.stt); });
    o.co_file = !!(o.pdf_id || o.docx_id);
    delete o.doc_id; delete o.pdf_id; delete o.docx_id;
    return o;
  });
  return {
    bienBan: list,
    danhBa: bbRead_(BB_TAB.db).filter(function (c) { return c.kich_hoat !== 'Không'; }).map(bbClean_),
    mau: bbRead_(BB_TAB.mau).map(bbClean_),
    cong_ty: BB_CONG_TY,
    dia_diem: BB_DIA_DIEM
  };
}

function api_bbAddContact(token, c) {
  auth_(token);
  var ten = bbStr_(c && c.ten, 120);
  if (!ten) throw new Error('Nhập họ tên người cần thêm.');
  var vai = ['Giao', 'Nhận', 'Cả hai'].indexOf(c.vai) >= 0 ? c.vai : 'Nhận';
  return withLock_(function () {
    var ex = bbRead_(BB_TAB.db).filter(function (x) { return String(x.ten).trim().toLowerCase() === ten.toLowerCase() && String(x.don_vi).trim().toLowerCase() === bbStr_(c.don_vi, 200).toLowerCase(); })[0];
    if (ex) return bbClean_(ex);
    var o = { ten: ten, chuc_vu: bbStr_(c.chuc_vu, 120), don_vi: bbStr_(c.don_vi, 200), vai: vai, sdt: bbStr_(c.sdt, 40), email: bbStr_(c.email, 120), kich_hoat: 'Có' };
    bbAppend_(BB_TAB.db, [o]);
    return o;
  });
}

/* Lưu (tạo mới hoặc sửa) rồi xuất Google Docs, PDF, Word vào Drive */
function api_bbSave(token, p) {
  var u = auth_(token);
  p = p || {};
  var items = (p.items || []).map(function (it, i) {
    return { stt: i + 1, cv_id: bbStr_(it.cv_id, 40), ten: bbStr_(it.ten, 500), hinh_thuc: bbStr_(it.hinh_thuc, 60), so_luong: bbStr_(it.so_luong, 20), don_vi: bbStr_(it.don_vi, 30), ghi_chu: bbStr_(it.ghi_chu, 300) };
  }).filter(function (it) { return it.ten; });
  if (!items.length) throw new Error('Danh mục bàn giao đang trống.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(p.ngay || ''))) throw new Error('Chọn ngày bàn giao.');
  if (!bbStr_(p.nguoi_giao)) throw new Error('Chọn người giao.');
  if (!bbStr_(p.nguoi_nhan)) throw new Error('Chọn người nhận.');
  var pj = bbRead_('DuAn').filter(function (d) { return d.ma === p.ma_da; })[0];
  if (!pj) throw new Error('Không tìm thấy dự án ' + p.ma_da + '.');

  var rec = withLock_(function () {
    var all = bbRead_(BB_TAB.bb), row, so, old = null;
    if (p.so_bb) {
      old = all.filter(function (r) { return r.so_bb === p.so_bb; })[0];
      if (!old) throw new Error('Không tìm thấy biên bản ' + p.so_bb + '.');
      if (!bbIsReview_(u) && old.nguoi_lap !== u.ten) throw new Error('Chỉ người lập hoặc người duyệt được sửa biên bản này.');
      if (old.trang_thai === 'Đã ký' && !bbIsReview_(u)) throw new Error('Biên bản đã ký, chỉ người duyệt được sửa.');
      if (old.ma_da !== p.ma_da) throw new Error('Không đổi dự án của biên bản đã lập. Hãy lập biên bản mới.');
      so = old.so_bb; row = old._row;
    } else {
      var max = 0, re = new RegExp('^BBBG-' + p.ma_da.replace(/[-]/g, '\\-') + '-(\\d+)$');
      all.forEach(function (r) { var m = String(r.so_bb).match(re); if (m) max = Math.max(max, Number(m[1])); });
      so = 'BBBG-' + p.ma_da + '-' + ('00' + (max + 1)).slice(-3);
    }
    var o = {
      so_bb: so, ngay: p.ngay, ma_da: pj.ma, ten_da: pj.ten, loai: bbStr_(p.loai, 200),
      ben_giao: bbStr_(p.ben_giao, 200), nguoi_giao: bbStr_(p.nguoi_giao, 120), cv_giao: bbStr_(p.cv_giao, 120),
      ben_nhan: bbStr_(p.ben_nhan, 200), nguoi_nhan: bbStr_(p.nguoi_nhan, 120), cv_nhan: bbStr_(p.cv_nhan, 120),
      dia_diem: bbStr_(p.dia_diem, 120) || BB_DIA_DIEM, noi_dung: bbStr_(p.noi_dung, 3000), ghi_chu: bbStr_(p.ghi_chu, 2000),
      trang_thai: old ? (old.trang_thai === 'Hủy' ? 'Đã lập' : old.trang_thai) : 'Đã lập',
      so_muc: items.length, nguoi_lap: old ? old.nguoi_lap : u.ten, thoi_diem: bbNow_()
    };
    if (old) bbWriteRow_(BB_TAB.bb, row, o); else bbAppend_(BB_TAB.bb, [o]);
    // thay danh mục: xóa dòng cũ (từ dưới lên) rồi ghi mới
    var ctS = bbSheet_(BB_TAB.ct);
    bbRead_(BB_TAB.ct).filter(function (r) { return r.so_bb === so; }).map(function (r) { return r._row; })
      .sort(function (a, b) { return b - a; }).forEach(function (r) { ctS.deleteRow(r); });
    bbAppend_(BB_TAB.ct, items.map(function (it) { it.so_bb = so; return it; }));
    o._old = old;
    return o;
  });

  var files = bbXuat_(rec, items, pj, rec._old);
  withLock_(function () {
    var r = bbRead_(BB_TAB.bb).filter(function (x) { return x.so_bb === rec.so_bb; })[0];
    if (r) bbWriteRow_(BB_TAB.bb, r._row, files);
  });
  delete rec._old;
  rec.items = items;
  rec.co_file = true;
  return rec;
}

function api_bbStatus(token, so, tt) {
  var u = auth_(token);
  if (['Đã lập', 'Đã ký', 'Hủy'].indexOf(tt) < 0) throw new Error('Trạng thái không hợp lệ.');
  return withLock_(function () {
    var r = bbRead_(BB_TAB.bb).filter(function (x) { return x.so_bb === so; })[0];
    if (!r) throw new Error('Không tìm thấy biên bản ' + so + '.');
    if (tt === 'Hủy' && !bbIsReview_(u)) throw new Error('Chỉ người duyệt được hủy biên bản.');
    if (!bbIsReview_(u) && r.nguoi_lap !== u.ten) throw new Error('Chỉ người lập hoặc người duyệt được đổi trạng thái.');
    bbWriteRow_(BB_TAB.bb, r._row, { trang_thai: tt, thoi_diem: bbNow_() });
    return { trang_thai: tt };
  });
}

/* Trả tệp về trình duyệt dạng base64, để ai đăng nhập app cũng tải được, không cần quyền Drive */
function api_bbFile(token, so, kind) {
  auth_(token);
  var r = bbRead_(BB_TAB.bb).filter(function (x) { return x.so_bb === so; })[0];
  if (!r) throw new Error('Không tìm thấy biên bản ' + so + '.');
  var id = kind === 'docx' ? r.docx_id : r.pdf_id;
  if (!id) throw new Error('Biên bản chưa có tệp. Bấm Sửa rồi Lưu để xuất lại.');
  var f = DriveApp.getFileById(id), blob = f.getBlob();
  return { name: f.getName(), mime: kind === 'docx' ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' : 'application/pdf', b64: Utilities.base64Encode(blob.getBytes()) };
}

/* ---------------- XUẤT TỆP ---------------- */
function bbKhongDau_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D')
    .replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 60);
}
function bbNgayChu_(iso) { var p = String(iso).split('-'); return 'ngày ' + p[2] + ' tháng ' + p[1] + ' năm ' + p[0]; }
function bbRe_(key) { return '\\{\\{' + key.replace(/#/g, '\\#') + '\\}\\}'; }
function bbRep_(v) { return String(v === null || v === undefined ? '' : v).replace(/\\/g, '\\\\').replace(/\$/g, '\\$'); }

function bbProjectFolder_(pj) {
  var root = bbFolder_(), name = pj.ma + '_' + bbKhongDau_(pj.ten), it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

/* Điền một mã {{KEY}} có thể nhiều dòng: dòng đầu thay tại chỗ, các dòng sau thành đoạn mới cùng định dạng */
function bbFillMulti_(body, key, text) {
  var lines = String(text || '').split(/\r?\n/), found = body.findText(bbRe_(key)), guard = 0;
  while (found && guard++ < 20) {
    var el = found.getElement(), par = el.getParent();
    while (par && par.getType() !== DocumentApp.ElementType.PARAGRAPH) par = par.getParent();
    el.asText().replaceText(bbRe_(key), bbRep_(lines[0]));
    if (par) {
      var host = par.getParent(), idx = host.getChildIndex(par);
      for (var i = 1; i < lines.length; i++) {
        var np = par.copy();
        np.setText(lines[i]);
        host.insertParagraph(idx + i, np);
      }
      if (!lines.join('').trim() && !par.getText().trim() && host.getNumChildren() > 1) {
        try { par.removeFromParent(); } catch (e) {}
      }
    }
    found = body.findText(bbRe_(key));
  }
}

function bbXuat_(rec, items, pj, old) {
  var tid = PropertiesService.getScriptProperties().getProperty('BB_TEMPLATE_ID');
  if (!tid) throw new Error('Chưa có mẫu biên bản. Chạy hàm caiBienBan một lần.');
  var folder = bbProjectFolder_(pj);
  if (old) ['doc_id', 'pdf_id', 'docx_id'].forEach(function (k) { if (old[k]) { try { DriveApp.getFileById(old[k]).setTrashed(true); } catch (e) {} } });
  var base = rec.so_bb + '_' + rec.ngay + '_' + bbKhongDau_(pj.ten);
  var copy = DriveApp.getFileById(tid).makeCopy(base, folder);
  var doc = DocumentApp.openById(copy.getId()), body = doc.getBody();

  // bảng danh mục: nhân dòng mẫu có {{#TEN}}
  var tables = body.getTables();
  for (var t = 0; t < tables.length; t++) {
    var tb = tables[t], tr = -1;
    for (var r = 0; r < tb.getNumRows(); r++) if (tb.getRow(r).getText().indexOf('{{#TEN}}') >= 0) { tr = r; break; }
    if (tr < 0) continue;
    var tpl = tb.getRow(tr);
    items.forEach(function (it, i) {
      var row = tb.insertTableRow(tr + 1 + i, tpl.copy());
      var sl = it.so_luong ? it.so_luong + (it.don_vi ? ' ' + it.don_vi : '') : '';
      [['#STT', it.stt], ['#TEN', it.ten], ['#HINH_THUC', it.hinh_thuc], ['#SO_LUONG', sl], ['#DON_VI', it.don_vi], ['#GHI_CHU', it.ghi_chu]]
        .forEach(function (kv) { row.replaceText(bbRe_(kv[0]), bbRep_(kv[1])); });
    });
    tb.removeRow(tr);
  }
  bbFillMulti_(body, 'NOI_DUNG', rec.noi_dung);
  bbFillMulti_(body, 'GHI_CHU', rec.ghi_chu);
  var map = {
    CONG_TY: BB_CONG_TY, SO_BB: rec.so_bb, NGAY: rec.ngay.split('-').reverse().join('/'), NGAY_CHU: bbNgayChu_(rec.ngay),
    DIA_DIEM: rec.dia_diem, LOAI: rec.loai, DU_AN: pj.ten + (pj.ten_thuong_mai ? ' (' + pj.ten_thuong_mai + ')' : ''),
    MA_DA: pj.ma, TEN_THUONG_MAI: pj.ten_thuong_mai || '',
    BEN_GIAO: rec.ben_giao, NGUOI_GIAO: rec.nguoi_giao, CV_GIAO: rec.cv_giao,
    BEN_NHAN: rec.ben_nhan, NGUOI_NHAN: rec.nguoi_nhan, CV_NHAN: rec.cv_nhan,
    SO_MUC: String(items.length), NGUOI_LAP: rec.nguoi_lap
  };
  Object.keys(map).forEach(function (k) { body.replaceText(bbRe_(k), bbRep_(map[k])); });
  doc.saveAndClose();

  var pdf = folder.createFile(copy.getAs('application/pdf').setName(base + '.pdf'));
  var resp = UrlFetchApp.fetch('https://docs.google.com/document/d/' + copy.getId() + '/export?format=docx',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  var docxId = '';
  if (resp.getResponseCode() === 200) docxId = folder.createFile(resp.getBlob().setName(base + '.docx')).getId();
  return { doc_id: copy.getId(), pdf_id: pdf.getId(), docx_id: docxId };
}

/* ---------------- NỐI VÀO CỔNG doPost CỦA Code.gs ----------------
   Trang trên GitHub gọi qua doPost → apiMap_(). Đoạn dưới bổ sung 5 chức năng biên bản
   vào apiMap_ mà không phải sửa Code.gs. Yêu cầu: tệp BienBan nằm DƯỚI Code.gs trong danh sách tệp. */
var BB_API = { api_bbInit: api_bbInit, api_bbSave: api_bbSave, api_bbStatus: api_bbStatus, api_bbFile: api_bbFile, api_bbAddContact: api_bbAddContact };
(function () {
  if (typeof apiMap_ !== 'function' || apiMap_._bb) return;
  var goc = apiMap_;
  apiMap_ = function () { var m = goc(); for (var k in BB_API) m[k] = BB_API[k]; return m; };
  apiMap_._bb = true;
})();

/* Chạy thử phía máy chủ: tạo một biên bản mẫu cho dự án đầu tiên (không cần đăng nhập web) */
function thuBienBan() {
  var pj = bbRead_('DuAn')[0], db = bbRead_(BB_TAB.db);
  var rec = { so_bb: 'BBBG-THU-001', ngay: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), loai: 'Hồ sơ thiết kế', ben_giao: BB_CONG_TY, nguoi_giao: db[0] ? db[0].ten : 'Người giao', cv_giao: '', ben_nhan: 'Đơn vị nhận', nguoi_nhan: 'Người nhận', cv_nhan: '', dia_diem: BB_DIA_DIEM, noi_dung: 'Dòng nội dung thứ nhất.\nDòng nội dung thứ hai.', ghi_chu: 'Biên bản thử, có thể xóa.', nguoi_lap: 'Thử' };
  var files = bbXuat_(rec, [{ stt: 1, ten: 'Hồ sơ thử 1', hinh_thuc: 'Bản cứng', so_luong: '02', don_vi: 'bộ', ghi_chu: '' }, { stt: 2, ten: 'Hồ sơ thử 2', hinh_thuc: 'Bản mềm', so_luong: '01', don_vi: 'file', ghi_chu: 'Gửi qua email' }], pj, null);
  Logger.log('PDF: https://drive.google.com/file/d/' + files.pdf_id + '/view\nDocs: https://docs.google.com/document/d/' + files.doc_id + '/edit');
}
