/**
 * HANH KPI - cầu nối đồng bộ giữa app và Google Sheet (bố cục tab MIX 1).
 * Cách dùng: mở Google Sheet > Tiện ích mở rộng > Apps Script > dán toàn bộ file này,
 * đổi TOKEN bên dưới, rồi Triển khai > Ứng dụng web (Chạy với tư cách: Tôi; Quyền truy cập: Bất kỳ ai).
 *
 * Quy tắc ghi vào bảng khách hàng (cột A tên, B kiểu thu, C COF, D NON, E POS quy đổi):
 *   - Thu kỳ: điền vào dòng trống đầu tiên từ TRÊN xuống (hoặc dòng đã có sẵn tên khách đó).
 *   - Thanh lý: điền từ DƯỚI lên.
 *   - Thu quá hạn: điền từ dưới lên, nằm phía trên các dòng Thanh lý.
 *   - Thu chỉ tính POS (Thanh lý không tính KPI): C và D để trống, cột E ghi số POS quy đổi (số nhập tay); bỏ tick thì trả lại công thức cột E.
 *   - Ngoại lệ: điền từ dưới lên, nằm phía trên các dòng Thu quá hạn. Cột E được ghi số POS quy đổi
 *     (đã nhân hệ số tùy chỉnh trong app); các kiểu khác giữ nguyên công thức ở cột E.
 *   - NTU: app sửa Tiền thu (K), Mindue (L), Hệ số (Q); POS đã thu (J) tự điền bằng POS khi K >= L (cả khi bạn nhập tay trên Sheet); có thể thêm hồ sơ mới vào dòng trống. Hồ sơ được nhận diện bằng số HĐ;
 *     sau mỗi lần ghi, bảng NTU được sắp xếp POS tăng dần từ trên xuống.
 *   - Mọi lệnh ghi đều kiểm tra lại tên khách hoặc số HĐ của dòng; không khớp thì bỏ qua, không ghi nhầm.
 */
var CFG = {
  SHEET: 'MIX 1 NEW',
  TOKEN: 'doi-mat-khau-nay'
};

var CUST = {n: 1, k: 2, cof: 3, non: 4, pe: 5};
/* NTU: app chỉ GHI posThu, tien, he. hd/n/pos/mind chỉ đọc. Không đụng cột R trở đi (bảng kiểu thu). */
var NTU = {hd: 7, n: 8, pos: 9, posThu: 10, tien: 11, mind: 12, he: 17};
var NTU_WRITE = {posThu: 1, tien: 1, he: 1, mind: 1};        /* sửa hồ sơ có sẵn */
var NTU_ADD = {hd: 1, n: 1, pos: 1, mind: 1, posThu: 1, tien: 1, he: 1}; /* thêm hồ sơ mới từ app */
var SETTINGS = {tNON: 'Target tiền NON', tCOF: 'Target tiền COF', tNTU: 'Target NTU', ntuOK: 'NTU CIC OK', dep: 'Số người phụ thuộc'};
var NL = 'Ngoại lệ';

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('HANH KPI').addItem('Sắp xếp NTU theo POS tăng dần', 'sortNtuNow').addToUi();
  } catch (e) {}
}

/* POS đã thu (cột J) tự điền = POS (cột I) khi Tiền thu (K) > 0 và K >= Mindue (L); ngược lại để trống. Chạy khi bạn tự nhập trên Sheet. */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== CFG.SHEET) return;
    var r = e.range;
    if (r.getLastColumn() < 9 || r.getColumn() > 12) return;
    var L = layout_(sh);
    var r1 = Math.max(r.getRow(), L.ntuStart), r2 = Math.min(r.getLastRow(), L.ntuEnd);
    if (r2 >= r1) fillPosThu_(sh, r1, r2);
  } catch (x) {}
}

function fillPosThu_(sh, r1, r2) {
  var n = r2 - r1 + 1;
  if (n <= 0) return;
  var v = sh.getRange(r1, 9, n, 4).getValues();          /* I,J,K,L */
  var f = sh.getRange(r1, 10, n, 1).getFormulasR1C1();   /* không ghi đè nếu J là công thức */
  for (var i = 0; i < n; i++) {
    if (f[i][0]) continue;
    var t = num_(v[i][2]), m = num_(v[i][3]), p = num_(v[i][0]);
    var want = (t !== '' && t > 0 && t >= (m === '' ? 0 : m) && p !== '') ? p : '';
    var cur = num_(v[i][1]);
    if (cur === want) continue;
    var cell = sh.getRange(r1 + i, 10);
    if (want === '') cell.clearContent(); else cell.setValue(want);
  }
}

function sortNtuNow() {
  var sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET);
  sortNtu_(sh, layout_(sh));
}

function doGet(e) {
  try {
    if (!e || !e.parameter || e.parameter.token !== CFG.TOKEN) return out_({ok: false, error: 'auth'});
    var sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET);
    if (!sh) return out_({ok: false, error: 'Không thấy tab ' + CFG.SHEET});
    return out_({ok: true, data: read_(sh)});
  } catch (err) {
    return out_({ok: false, error: String(err && err.message || err)});
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.token !== CFG.TOKEN) return out_({ok: false, error: 'auth'});
    lock.waitLock(20000);
    var sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET);
    if (!sh) return out_({ok: false, error: 'Không thấy tab ' + CFG.SHEET});
    var res = apply_(sh, body.patches || []);
    SpreadsheetApp.flush();
    return out_({ok: true, map: res.map, skipped: res.skipped, data: read_(sh)});
  } catch (err) {
    return out_({ok: false, error: String(err && err.message || err)});
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---- tìm vị trí các khối trong tab ---- */
function layout_(sh) {
  var last = Math.max(sh.getLastRow(), 10);
  var A = sh.getRange(1, 1, last, 1).getDisplayValues();
  var G = sh.getRange(1, 7, last, 1).getDisplayValues();
  var L = {custHead: 0, custStart: 0, custEnd: 0, ntuHead: 0, ntuStart: 0, ntuEnd: 0, labels: {}};
  var i, t;
  for (i = 0; i < last; i++) {
    t = String(A[i][0]).trim();
    if (t === 'TÊN KH ĐÓNG' && !L.custHead) L.custHead = i + 1;
    for (var k in SETTINGS) if (t === SETTINGS[k] && !L.labels[k]) L.labels[k] = i + 1;
    t = String(G[i][0]).trim();
    if (t === 'Số HĐ' && !L.ntuHead) L.ntuHead = i + 1;
  }
  if (!L.custHead || !L.ntuHead) throw new Error('Không nhận ra bố cục tab (thiếu tiêu đề TÊN KH ĐÓNG hoặc Số HĐ)');
  L.custStart = L.custHead + 1;
  L.ntuStart = L.ntuHead + 1;
  L.ntuEnd = L.ntuStart - 1;
  for (i = L.ntuStart; i <= last; i++) {
    if (String(G[i - 1][0]).trim().toUpperCase() === 'TỔNG') { L.ntuEnd = i - 1; break; }
    L.ntuEnd = i;
  }
  /* khối khách hàng: từ dưới tiêu đề đến trước hàng đầu tiên mà cột A và B đều trống, hoặc hàng "TỔNG...", hoặc nhãn cài đặt */
  var Bc = sh.getRange(1, 2, last, 1).getDisplayValues();
  var stopRows = [];
  for (var kk in L.labels) stopRows.push(L.labels[kk]);
  L.custEnd = last;
  for (i = L.custStart; i <= last; i++) {
    t = String(A[i - 1][0]).trim().toUpperCase();
    if (t.indexOf('TỔNG') === 0 || stopRows.indexOf(i) >= 0 || (t === '' && String(Bc[i - 1][0]).trim() === '')) { L.custEnd = i - 1; break; }
  }
  /* tháng: tìm "THÁNG ..." ở các hàng phía trên tiêu đề */
  L.monthCell = null;
  for (i = 1; i < L.custHead; i++) {
    for (var c = 1; c <= 5; c++) {
      var v = String(sh.getRange(i, c).getDisplayValue()).trim();
      if (/^THÁNG\s*\d+/i.test(v)) { L.monthCell = [i, c, v]; break; }
    }
    if (L.monthCell) break;
  }
  return L;
}

function num_(v) {
  if (v === '' || v == null) return '';
  if (typeof v === 'number') return v;
  var s = String(v).replace(/[^\d,.\-]/g, '');
  if (!s) return '';
  if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  else if ((s.match(/\./g) || []).length > 1) s = s.replace(/\./g, '');
  var n = parseFloat(s);
  return isFinite(n) ? n : '';
}

function norm_(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); }

function read_(sh) {
  var L = layout_(sh), data = {month: L.monthCell ? L.monthCell[2] : '', s: {}, cust: [], ntu: []};
  var k, r;
  for (k in SETTINGS) {
    var row = L.labels[k];
    data.s[k] = row ? num_(sh.getRange(row, 2).getValue()) : '';
  }
  if (L.custEnd >= L.custStart) {
    var cn = L.custEnd - L.custStart + 1;
    var cv = sh.getRange(L.custStart, 1, cn, 5).getValues();
    var cfm = sh.getRange(L.custStart, CUST.pe, cn, 1).getFormulasR1C1();
    for (r = 0; r < cv.length; r++) {
      var a = cv[r];
      var kind = String(a[1] || 'Thu kỳ');
      /* "Thu chỉ tính POS": C và D trống, cột E là số (kiểu Ngoại lệ luôn là số, các kiểu khác là số nhập tay không có công thức) */
      var isX = (kind === NL || !cfm[r][0]) && num_(a[4]) !== '' && a[2] === '' && a[3] === '';
      if (String(a[0]).trim() === '' && a[2] === '' && a[3] === '' && !isX) continue;
      data.cust.push({row: L.custStart + r, n: String(a[0]), k: kind, cof: num_(a[2]), non: num_(a[3]), pe: kind === NL || isX ? num_(a[4]) : '', x: isX ? 1 : ''});
    }
  }
  if (L.ntuEnd >= L.ntuStart) {
    var nv = sh.getRange(L.ntuStart, 7, L.ntuEnd - L.ntuStart + 1, 11).getValues();
    for (r = 0; r < nv.length; r++) {
      var b = nv[r]; /* cột G..Q: 0=G hd,1=H n,2=I pos,3=J posThu,4=K tien,5=L mind,10=Q he */
      if (String(b[0]).trim() === '' && String(b[1]).trim() === '') continue;
      data.ntu.push({row: L.ntuStart + r, hd: String(b[0]).trim(), n: String(b[1]), pos: num_(b[2]), posThu: num_(b[3]), tien: num_(b[4]), mind: num_(b[5]), he: num_(b[10])});
    }
  }
  return data;
}

/* ---- chọn dòng trống theo quy tắc ---- */
function grp_(kind) {
  kind = norm_(kind);
  if (/^Thanh lý/i.test(kind)) return 0;
  if (kind === 'Thu quá hạn') return 1;
  if (kind === NL) return 2;
  return 3; /* Thu kỳ */
}

function pickCustRow_(sh, L, kind) {
  var n = L.custEnd - L.custStart + 1;
  if (n <= 0) throw new Error('Bảng khách hàng trống');
  var v = sh.getRange(L.custStart, 1, n, 4).getValues();
  function free(i) { return norm_(v[i][0]) === '' && v[i][2] === '' && v[i][3] === ''; }
  var g = grp_(kind), i;
  if (g === 3) {
    for (i = 0; i < n; i++) if (free(i)) return L.custStart + i;
  } else {
    /* từ dưới lên; Thu quá hạn nằm trên Thanh lý, Ngoại lệ nằm trên Thu quá hạn */
    var limit = n; /* chỉ số (0-based) của dòng đã dùng cao nhất thuộc nhóm thấp hơn */
    for (i = 0; i < n; i++) {
      if (!free(i) && grp_(v[i][1]) < g && i < limit) limit = i;
    }
    for (i = limit - 1; i >= 0; i--) if (free(i)) return L.custStart + i;
    for (i = n - 1; i >= 0; i--) if (free(i)) return L.custStart + i;
  }
  throw new Error('Hết dòng trống trong bảng khách hàng');
}

function peTemplate_(sh, L) {
  for (var r = L.custStart; r <= L.custEnd; r++) {
    var f = sh.getRange(r, CUST.pe).getFormulaR1C1();
    if (f) return f;
  }
  return null;
}

/* Cột E: Ngoại lệ giữ số POS quy đổi do app tính; kiểu khác phải có công thức gốc. */
function finalizeRow_(sh, L, row, f, tpl) {
  var kind = norm_(sh.getRange(row, CUST.k).getValue());
  var c = sh.getRange(row, CUST.pe);
  if (kind === NL) {
    if (f && f.pe !== undefined && f.pe !== '' && f.pe != null) c.setValue(Number(f.pe));
    else if (f && f.pe === '' && tpl) c.setFormulaR1C1(tpl);
  } else {
    /* Thu chỉ tính POS (x=1): C, D trống, E giữ số POS quy đổi do app tính, không đụng công thức. Bỏ tick (x='') thì trả lại công thức E. */
    var hasF = !!c.getFormulaR1C1(), curX;
    if (f && f.x !== undefined && f.x !== null) {
      curX = (f.x === 1 || f.x === '1' || f.x === true);
    } else {
      var cd = sh.getRange(row, CUST.cof, 1, 2).getValues()[0];
      curX = !hasF && num_(c.getValue()) !== '' && cd[0] === '' && cd[1] === '';
    }
    if (curX) {
      if (f && f.pe !== undefined && f.pe !== '' && f.pe != null) c.setValue(Number(f.pe));
      else if (hasF) c.setValue(0);
    } else if (tpl && !hasF) {
      c.setFormulaR1C1(tpl);
    }
  }
}

function setField_(sh, list, row, f, v, isAdd) {
  var col = (list === 'cust' ? CUST : NTU)[f];
  if (!col || f === 'pe') return;
  if (list === 'ntu' && !(isAdd ? NTU_ADD : NTU_WRITE)[f]) return;
  var c = sh.getRange(row, col);
  if (v === '' || v == null) { c.clearContent(); return; }
  if (f === 'hd') { c.setNumberFormat('@'); c.setValue(String(v)); return; }
  if (f === 'n' || f === 'k') { c.setValue(String(v)); return; }
  c.setValue(Number(v));
}

function freeNtuRow_(sh, L) {
  var n = L.ntuEnd - L.ntuStart + 1;
  if (n <= 0) throw new Error('Bảng NTU trống');
  var v = sh.getRange(L.ntuStart, 7, n, 2).getValues();
  for (var i = 0; i < n; i++) if (norm_(v[i][0]) === '' && norm_(v[i][1]) === '') return L.ntuStart + i;
  throw new Error('Hết dòng trống trong bảng NTU');
}

function findNtuRow_(sh, L, hd, hint) {
  hd = norm_(hd);
  if (!hd) return 0;
  var n = L.ntuEnd - L.ntuStart + 1;
  if (n <= 0) return 0;
  var v = sh.getRange(L.ntuStart, 7, n, 1).getValues();
  if (hint >= L.ntuStart && hint <= L.ntuEnd && norm_(v[hint - L.ntuStart][0]) === hd) return hint;
  for (var i = 0; i < n; i++) if (norm_(v[i][0]) === hd) return L.ntuStart + i;
  return 0;
}

function apply_(sh, patches) {
  var L = layout_(sh), map = {}, skipped = 0, i, p, tpl = peTemplate_(sh, L), f, row;
  for (i = 0; i < patches.length; i++) {
    p = patches[i];
    if (p.t === 'u') {
      if (p.l === 'cust') {
        row = p.row;
        if (!(row >= L.custStart && row <= L.custEnd) || (p.exp !== undefined && norm_(sh.getRange(row, 1).getValue()) !== norm_(p.exp))) { skipped++; continue; }
        for (f in p.f) setField_(sh, 'cust', row, f, p.f[f]);
        finalizeRow_(sh, L, row, p.f, tpl);
      } else {
        row = findNtuRow_(sh, L, p.hd, p.row);
        if (!row) { skipped++; continue; }
        for (f in p.f) setField_(sh, 'ntu', row, f, p.f[f]);
      }
    } else if (p.t === 'a') {
      if (p.l === 'ntu') {
        var hd0 = norm_(p.f.hd);
        if (!hd0) { skipped++; continue; }
        var ex = findNtuRow_(sh, L, hd0, 0);
        if (ex) { for (f in p.f) setField_(sh, 'ntu', ex, f, p.f[f], false); map[p.tmp] = ex; continue; } /* đã có: chỉ cập nhật */
        row = freeNtuRow_(sh, L);
        for (f in p.f) setField_(sh, 'ntu', row, f, p.f[f], true);
        map[p.tmp] = row;
        continue;
      }
      row = pickCustRow_(sh, L, p.f.k || 'Thu kỳ');
      sh.getRange(row, CUST.k).setValue(String(p.f.k || 'Thu kỳ'));
      for (f in p.f) setField_(sh, 'cust', row, f, p.f[f]);
      finalizeRow_(sh, L, row, p.f, tpl);
      map[p.tmp] = row;
    } else if (p.t === 'd') {
      if (p.l === 'ntu') {
        row = findNtuRow_(sh, L, p.hd, p.row);
        if (!row) { skipped++; continue; }
        sh.getRange(row, 7, 1, 8).clearContent();  /* G..N */
        sh.getRange(row, 17, 1, 1).clearContent(); /* Q; giữ công thức O, P */
        continue;
      }
      row = p.row;
      if (!(row >= L.custStart && row <= L.custEnd) || (p.exp !== undefined && norm_(sh.getRange(row, 1).getValue()) !== norm_(p.exp))) { skipped++; continue; }
      sh.getRange(row, CUST.cof, 1, 2).clearContent();
      if (p.full) { sh.getRange(row, CUST.n).clearContent(); sh.getRange(row, CUST.k).setValue('Thu kỳ'); }
      finalizeRow_(sh, L, row, {pe: '', x: ''}, tpl);
    } else if (p.t === 's') {
      var lr = L.labels[p.k];
      if (lr) {
        var cell = sh.getRange(lr, 2);
        if (p.v === '' || p.v == null) cell.clearContent(); else cell.setValue(Number(p.v));
      }
    }
  }
  if (patches.length) { fillPosThu_(sh, L.ntuStart, L.ntuEnd); sortNtu_(sh, L); }
  return {map: map, skipped: skipped};
}

/* ---- NTU: POS tăng dần từ trên xuống. Chỉ dời các ô nhập tay (G..N và Q); công thức O, P cùng dòng giữ nguyên. ---- */
function sortNtu_(sh, L) {
  var n = L.ntuEnd - L.ntuStart + 1;
  if (n < 2) return false;
  var rng = sh.getRange(L.ntuStart, 7, n, 11); /* G..Q */
  var fm = rng.getFormulasR1C1();
  var vals = rng.getValues();
  var movable = [0, 1, 2, 3, 4, 5, 6, 7, 10]; /* G,H,I,J,K,L,M,N,Q */
  var r, c;
  for (r = 0; r < n; r++) for (c = 0; c < movable.length; c++) if (fm[r][movable[c]]) return false; /* có công thức trong ô nhập tay: không đụng */
  var items = [];
  for (r = 0; r < n; r++) {
    var used = norm_(vals[r][0]) !== '' || norm_(vals[r][1]) !== '';
    var pos = num_(vals[r][2]);
    items.push({i: r, used: used, pos: pos === '' ? null : pos});
  }
  var sorted = items.slice().sort(function (a, b) {
    if (a.used !== b.used) return a.used ? -1 : 1;
    if (!a.used) return a.i - b.i;
    var ap = a.pos == null, bp = b.pos == null;
    if (ap !== bp) return ap ? 1 : -1;
    if (!ap && a.pos !== b.pos) return a.pos - b.pos;
    return a.i - b.i;
  });
  var same = true;
  for (r = 0; r < n; r++) if (sorted[r].i !== r) { same = false; break; }
  if (same) return false;
  sh.getRange(L.ntuStart, 7, n, 1).setNumberFormat('@'); /* số HĐ giữ dạng chữ */
  var blockA = [], blockQ = [];
  for (r = 0; r < n; r++) {
    var src = vals[sorted[r].i];
    blockA.push(src.slice(0, 8));
    blockQ.push([src[10]]);
  }
  sh.getRange(L.ntuStart, 7, n, 8).setValues(blockA);
  sh.getRange(L.ntuStart, 17, n, 1).setValues(blockQ);
  return true;
}
