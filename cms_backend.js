/*************************************************
 * cms_backend.gs — API backend cho Web Quản Trị
 *  - Hybrid Auth: Google SSO + tài khoản nội bộ (username/password)
 *  - KPI Dashboard tổng hợp (quỹ, tổng vay, đã thu gốc/lãi/phí, counters)
 *  - Phiếu thanh toán linh hoạt: theo kỳ & theo khoảng ngày (tất toán sớm)
 *  - Lịch kỳ sắp tới
 *  - Cấu hình (quỹ vốn, loại thanh toán, Telegram) + Import ở tab Cấu hình
 *  - Telegram Groups (quản lý nhiều nhóm theo mục đích gửi)
 *  - CRM: Interactions, Tasks, Ratings (đánh giá theo ID Khách)
 *  - Phân quyền menu (google/local)
 **************************************************/

/** ====== COMMON UTILS ====== */
function _me_() { return Session.getActiveUser().getEmail() || ''; }
function _now_(){ return new Date(); }
function _toISO_(d){ var y=d.getFullYear(),m=('0'+(d.getMonth()+1)).slice(-2),da=('0'+d.getDate()).slice(-2),h=('0'+d.getHours()).slice(-2),mi=('0'+d.getMinutes()).slice(-2),s=('0'+d.getSeconds()).slice(-2); return y+'-'+m+'-'+da+' '+h+':'+mi+':'+s; }
function _ymd_(d){ return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function _startOfDay_(d){ var x=new Date(d); x.setHours(0,0,0,0); return x; }
function _endOfDay_(d){ var x=new Date(d); x.setHours(23,59,59,999); return x; }
function _endOfPrevDay_(d){ var x=new Date(d); x.setDate(x.getDate()-1); x.setHours(23,59,59,999); return x; }
function _addMonths_(d, n){ var x=new Date(d); x.setMonth(x.getMonth()+Number(n||0)); return x; }
function _monthsDiff_(from, to){ return (to.getFullYear()-from.getFullYear())*12 + (to.getMonth()-from.getMonth()); }
function _toNum_(v){ var n = Number((v||'').toString().replace(/[^\d\.\-]/g,'')); return isNaN(n)?0:n; }
/** Helper: chuyển giá trị ô thành Date (nếu có) */
function toDate_(v){
  if (!v) return null;
  try {
    if (Object.prototype.toString.call(v) === '[object Date]') return v;
    return new Date(v);
  } catch(e){ return null; }
}
function _hash256b64_(s){ var b=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8); return Utilities.base64Encode(b); }

/** ====== SHEET HELPERS ====== */
function _cms_getSheet_(tableKey) {
  var def = BASE_TABLES[tableKey];
  if (!def) {
    var keys = Object.keys(BASE_TABLES || {});
    throw new Error('Sai tableKey: ' + tableKey + ' — Hợp lệ: ' + keys.join(', '));
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(def.sheetName);
  if (!sh) throw new Error('Không thấy sheet: ' + def.sheetName + ' (tableKey=' + tableKey + ')');
  return {sheet: sh, idColName: def.idCol, label: def.label};
}
function _getSheetValues_(sheetName){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Không thấy sheet: ' + sheetName);
  var lastR = sh.getLastRow(), lastC = sh.getLastColumn();
  if (lastR < 1 || lastC < 1) return {headers:[], rows:[]};
  var vals = sh.getRange(1,1,lastR,lastC).getDisplayValues();
  return {headers: vals[0], rows: vals.slice(1), sheet: sh};
}
function _headerMap_(row1) { var map={}; row1.forEach(function(h,i){ map[String(h).trim()] = i; }); return map; }
function _pick_(row, idx){ return (idx!=null && idx>=0) ? row[idx] : ''; }

/** ====== ID helpers (VC000001 / HD000001 / PAY000001) ====== */
function _nextSeqId_(sh, idIdx, prefix, width) {
  var last = sh.getLastRow();
  if (last < 2) return prefix + String(1).padStart(width, '0');
  var vals = sh.getRange(2, idIdx+1, last-1, 1).getDisplayValues();
  var maxN = 0;
  vals.forEach(function(r){
    var s = String(r[0] || '');
    if (s.indexOf(prefix) === 0) {
      var num = Number(s.slice(prefix.length));
      if (!isNaN(num) && num > maxN) maxN = num;
    }
  });
  var n = maxN + 1;
  return prefix + String(n).padStart(width, '0');
}

/** ====== BASIC GUARD (Google SSO / allow-list cũ) ====== */
function _getProps_(){ return PropertiesService.getScriptProperties(); }
function _getAllowed_(){
  try { return JSON.parse(_getProps_().getProperty('ALLOWED') || '[]'); } catch(e){ return []; }
}
function _cms_guard_() {
  var email = Session.getActiveUser().getEmail() || '';
  if (ALLOWED_EMAILS && ALLOWED_EMAILS.length){
    if (ALLOWED_EMAILS.indexOf(email) === -1) throw new Error('Không có quyền: ' + email);
    return;
  }
  var allow = _getAllowed_();
  if (allow.length && allow.indexOf(email) === -1) throw new Error('Không có quyền: ' + email);
}

/** ====== ROLES theo email (menu) ====== */
function _getRoles_(){
  try { return JSON.parse(_getProps_().getProperty('ROLES') || '{}'); } catch(e){ return {}; }
}
function getVisibleMenusForMe(allKeys){
  var me = _me_();
  var roles=_getRoles_();
  var arr = roles[me] || roles['*'] || allKeys || [];
  return (arr.length ? arr : allKeys);
}
function mustHave(menuKey){
  var allowed = getVisibleMenusForMe([]);
  if (allowed.indexOf(menuKey)===-1) throw new Error('Bạn không có quyền truy cập mục: '+menuKey);
  return true;
}

/** ====== HYBRID AUTH (Local account + Google SSO) ====== */
/* Sheet ACCOUNTS: User | Display | PasswordHash | Salt | Menus | Active */
function _accSheet_(){
  var ss=SpreadsheetApp.getActive(); var sh=ss.getSheetByName('ACCOUNTS');
  if(!sh) throw new Error('Thiếu sheet ACCOUNTS (User|Display|PasswordHash|Salt|Menus|Active).');
  var lr=sh.getLastRow(), lc=sh.getLastColumn();
  if(lr<1||lc<1) throw new Error('Sheet ACCOUNTS rỗng.');
  var vals=sh.getRange(1,1,lr,lc).getDisplayValues();
  return {headers: vals[0], rows: vals.slice(1), sheet: sh};
}
function _accMap_(h){ var m={}; h.forEach((x,i)=>m[String(x).trim()]=i); return m; }

/** Cache session token */
function _sessCache_(){ return CacheService.getScriptCache(); }
function _randToken_(){ return Utilities.getUuid().replace(/-/g,''); }
function _expInMinutes_(n){ var d=new Date(); d.setMinutes(d.getMinutes()+n); return d; }

/** Admin tạo/sửa tài khoản nội bộ */
function admin_upsertLocalUser(user, display, passwordPlain, menusCsv, active){
  _cms_guard_();
  user=String(user||'').trim(); if(!user) throw new Error('Thiếu user');

  var s=_accSheet_(); var H=_accMap_(s.headers);
  var iUser=H['User'], iDisp=H['Display'], iHash=H['PasswordHash'], iSalt=H['Salt'], iMenus=H['Menus'], iAct=H['Active'];

  var found=-1;
  for (var i=0;i<s.rows.length;i++){
    if (String(s.rows[i][iUser]).trim().toLowerCase()===user.toLowerCase()){ found=i; break; }
  }

  var salt = (passwordPlain) ? Utilities.getUuid() : (found>=0 ? s.rows[found][iSalt] : Utilities.getUuid());
  var hash = (passwordPlain) ? _hash256b64_(passwordPlain+salt) : (found>=0 ? s.rows[found][iHash] : '');
  var disp = (display!=null)?display:(found>=0?s.rows[found][iDisp]:user);
  var menus= (menusCsv!=null)?menusCsv:(found>=0?s.rows[found][iMenus]:MENU_CONFIG.map(m=>m.key).join(','));
  var act  = (active!=null)?active:(found>=0?s.rows[found][iAct]:'TRUE');

  if(found>=0){
    s.rows[found][iDisp]=disp; s.rows[found][iHash]=hash; s.rows[found][iSalt]=salt;
    s.rows[found][iMenus]=menus; s.rows[found][iAct]=String(act);
    s.sheet.getRange(found+2,1,1,s.headers.length).setValues([s.rows[found]]);
    return {ok:true, action:'update', user:user};
  } else {
    var row=[]; row[iUser]=user; row[iDisp]=disp; row[iHash]=hash; row[iSalt]=salt; row[iMenus]=menus; row[iAct]=String(act);
    for (var c=0;c<s.headers.length;c++){ if (typeof row[c]==='undefined') row[c]=''; }
    s.sheet.appendRow(row);
    return {ok:true, action:'create', user:user};
  }
}
function admin_setLocalUserPassword(user, newPassword){
  _cms_guard_();
  return admin_upsertLocalUser(user, null, newPassword, null, null);
}

/** Local login → token */
function local_login(user, password){
  user=String(user||'').trim(); if(!user) throw new Error('Thiếu user');
  var s=_accSheet_(); var H=_accMap_(s.headers);
  var iUser=H['User'], iDisp=H['Display'], iHash=H['PasswordHash'], iSalt=H['Salt'], iMenus=H['Menus'], iAct=H['Active'];
  var row=null;
  for (var i=0;i<s.rows.length;i++){
    if (String(s.rows[i][iUser]).trim().toLowerCase()===user.toLowerCase()){ row=s.rows[i]; break; }
  }
  if(!row) throw new Error('Sai tài khoản hoặc mật khẩu.');
  if(String(row[iAct]).toLowerCase()==='false') throw new Error('Tài khoản đã bị khoá.');

  var salt=row[iSalt]||''; var expect=row[iHash]||''; var got=_hash256b64_(String(password||'')+salt);
  if (got!==expect) throw new Error('Sai tài khoản hoặc mật khẩu.');

  var token=_randToken_();
  var payload={ user: row[iUser], display: row[iDisp]||row[iUser],
                menus: (row[iMenus]||'').split(',').map(x=>x.trim()).filter(Boolean),
                exp: _toISO_(_expInMinutes_(180)) };
  _sessCache_().put('SESS::'+token, JSON.stringify(payload), 60*180);
  return { ok:true, token: token, display: payload.display, menus: payload.menus };
}
function local_logout(token){
  if(!token) return {ok:true};
  _sessCache_().remove('SESS::'+token);
  return {ok:true};
}
function _sessionFromToken_(token){
  if(!token) return null;
  var txt=_sessCache_().get('SESS::'+token); if(!txt) return null;
  try{ return JSON.parse(txt); }catch(e){ return null; }
}

/** Hybrid guard: Google SSO hoặc token local */
function _hybridGuard_(payload){
  try{ _cms_guard_(); return {mode:'google', principal: _me_()}; }catch(e){}
  var token = payload && (payload.session || payload.token);
  var sess = _sessionFromToken_(token);
  if (sess && Array.isArray(sess.menus)) return {mode:'local', principal: sess.user, menus: sess.menus};
  throw new Error('Bạn chưa đăng nhập hoặc không có quyền.');
}
function mustHaveHybrid(payload, menuKey){
  var ctx=_hybridGuard_(payload||{});
  if (ctx.mode==='local'){
    if (!ctx.menus || ctx.menus.indexOf(menuKey)===-1) throw new Error('Không có quyền vào mục: '+menuKey);
  } else {
    try{ mustHave(menuKey); }catch(e){ /* nếu chưa set role theo email, cho qua */ }
  }
  return true;
}

/** Map tableKey → menuKey */
function _menuKeyForTable_(tbl){
  switch (String(tbl||'').toUpperCase()) {
    case 'CUSTOMERS': return 'customers';
    case 'LOANS':     return 'loans';
    case 'PAYMENTS':  return 'loanmgr';
    default:          return 'dashboard';
  }
}

/** ===================== CRUD (LIST / SAVE / DELETE) ===================== */
function listRecords(payload, page, q) {
  var tbl = (typeof payload === 'string') ? payload : (payload && payload.tableKey);
  var menuKey = _menuKeyForTable_(tbl);
  mustHaveHybrid(payload, menuKey);

  if (typeof payload === 'string') payload = { tableKey: payload, page: page, q: q };
  payload = payload || {};
  var tableKey = payload.tableKey;
  if (!tableKey) throw new Error('Thiếu tableKey khi gọi listRecords.');

  var pageNum  = Math.max(1, Number(payload.page || 1));
  var pageSize = Math.max(1, Number(payload.pageSize || DEFAULT_PAGE_SIZE));
  var kw       = (payload.q || '').toString().trim().toLowerCase();

  var cfg  = _cms_getSheet_(tableKey);
  var sh   = cfg.sheet;

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { total: 0, headers: [], rows: [], label: cfg.label };

  var vals    = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headers = vals[0] || [];
  var dataAll = vals.slice(1);
  var data = dataAll.filter(function(r){ return r.join('').trim().length > 0; });

  if (kw) data = data.filter(function(r){ return r.join(' ').toLowerCase().indexOf(kw) !== -1; });

  var total = data.length;
  var start = (pageNum - 1) * pageSize;
  var rows  = data.slice(start, start + pageSize);

  return { total: total, headers: headers, rows: rows, label: cfg.label };
}

function saveRecord(payload) {
  var tbl = payload && payload.tableKey;
  var menuKey = _menuKeyForTable_(tbl);
  mustHaveHybrid(payload, menuKey);

  payload = payload || {};
  var tableKey = payload.tableKey;
  var record   = payload.record || {};

  var cfg  = _cms_getSheet_(tableKey);
  var sh   = cfg.sheet;
  var last = sh.getLastRow();
  var width= sh.getLastColumn();
  if (last < 1) throw new Error('Sheet chưa có header');

  var headers = sh.getRange(1,1,1,width).getValues()[0];
  var hmap = _headerMap_(headers);

  var rowArr = headers.map(function(h){
    var v = record[h];
    return (typeof v === 'undefined') ? '' : v;
  });

  var idColName = cfg.idColName;
  var idIdx = hmap[idColName];
  if (typeof idIdx === 'undefined') throw new Error('Không thấy cột ID: ' + idColName);

  var idVal = (record[idColName] || '').toString().trim();

  var now = _now_(), me = _me_();
  if (hmap['Cập nhật']  !== undefined) rowArr[hmap['Cập nhật']]  = now;
  if (hmap['Người sửa'] !== undefined) rowArr[hmap['Người sửa']] = me;

  if (!idVal) {
    var prefix = (tableKey === 'CUSTOMERS') ? 'VC' :
                 (tableKey === 'LOANS')     ? 'HD' :
                 (tableKey === 'PAYMENTS')  ? 'PAY' : 'ID';
    var newId = _nextSeqId_(sh, idIdx, prefix, 6);
    rowArr[idIdx] = newId;
    if (hmap['Tạo'] !== undefined) rowArr[hmap['Tạo']] = now;
    sh.appendRow(rowArr);
    return { ok: true, action: 'create', id: newId, record: rowArr };
  } else {
    if (last < 2) {
      if (hmap['Tạo'] !== undefined) rowArr[hmap['Tạo']] = now;
      sh.appendRow(rowArr);
      return { ok: true, action: 'create', id: idVal, record: rowArr };
    }
    var data = sh.getRange(2,1,last-1,width).getValues();
    var foundRow = -1;
    for (var i=0;i<data.length;i++){
      if (String(data[i][idIdx]) === idVal) { foundRow = i+2; break; }
    }
    if (foundRow === -1) {
      if (hmap['Tạo'] !== undefined) rowArr[hmap['Tạo']] = now;
      sh.appendRow(rowArr);
      return { ok: true, action: 'create', id: idVal, record: rowArr };
    }
    sh.getRange(foundRow,1,1,width).setValues([rowArr]);
    return { ok: true, action: 'update', id: idVal, record: rowArr };
  }
}

function deleteRecord(payload) {
  var tbl = payload && payload.tableKey;
  var menuKey = _menuKeyForTable_(tbl);
  mustHaveHybrid(payload, menuKey);

  payload = payload || {};
  var tableKey = payload.tableKey;
  var id       = (payload.id || '').toString().trim();
  if (!id) throw new Error('Thiếu ID');

  var cfg  = _cms_getSheet_(tableKey);
  var sh   = cfg.sheet;
  var last = sh.getLastRow();
  var width= sh.getLastColumn();
  if (last < 2) return {ok:false, msg:'Sheet rỗng'};

  var headers = sh.getRange(1,1,1,width).getValues()[0];
  var hmap = _headerMap_(headers);
  var idIdx = hmap[cfg.idColName];
  var data = sh.getRange(2,1,last-1,width).getValues();

  for (var i=0;i<data.length;i++){
    if (String(data[i][idIdx]) === id) {
      sh.deleteRow(i+2);
      return {ok:true, id:id};
    }
  }
  return {ok:false, msg:'Không tìm thấy ID: '+id};
}


/** ===================== LÃI/GỐC — CÔNG THỨC & PHIẾU ===================== */

function _calcPeriodInfo_(loan, refDate){
  var type  = loan.type;      // 'Cuốn chiếu' | 'Cuối kỳ'
  var P     = loan.P;         // principal
  var r     = loan.r;         // lãi/tháng (decimal)
  var m     = loan.m;         // số kỳ
  var start = loan.start;     // ngày giải ngân
  var cycle = loan.cycle || 1;

  var firstStart = new Date(start);
  var ref = _startOfDay_(refDate);
  var kth = Math.floor(_monthsDiff_(firstStart, ref)/cycle) + 1;
  if (kth < 1) kth = 1;
  if (kth > m) kth = m;

  var periodStart = _addMonths_(firstStart, (kth-1)*cycle);
  var periodEnd   = _endOfPrevDay_(_addMonths_(firstStart, kth*cycle));

  var principalDue = 0, interestDue = 0;

  if (type === 'Cuốn chiếu' || type === LOAN_TYPES.CUON_CHIEU){
    var gPer = P / m;
    var duNoStart = Math.max(0, P - gPer*(kth-1));
    principalDue = gPer;
    interestDue  = duNoStart * r * cycle;
  } else {
    principalDue = (kth===m) ? P : 0;
    interestDue  = P * r * cycle;
  }

  return {
    k: kth,
    start: periodStart,
    end: periodEnd,
    principalDue: principalDue,
    interestDue : interestDue,
    totalDue    : principalDue + interestDue
  };
}

/** Lịch chuẩn theo kỳ (giữ chữ ký cũ) */
function buildSchedule(payload){
  payload = payload || {};
  var type   = payload.type;
  var P      = Number(payload.principal || payload.P || 0);
  var r      = Number(payload.rateMonth || payload.r || 0);
  var m      = Number(payload.months || payload.m || 0);
  var start  = new Date(payload.startDate || payload.start);
  var cycle  = Number(payload.cycle || 1);
  if (!P || !m || isNaN(start)) return {rows:[]};

  var rows = [];
  if (type === LOAN_TYPES.CUON_CHIEU || type === 'Cuốn chiếu') {
    var gocKy = P / m;
    var duNo = P;
    for (var k=1;k<=m;k++){
      var lai = duNo * (r*cycle);
      var goc = gocKy;
      var tong = goc + lai;
      duNo = Math.max(0, duNo - goc);
      rows.push({ky:k, ngay: _addMonths_(start, k*cycle), goc:goc, lai:lai, tong:tong, duNo:duNo});
    }
  } else {
    for (var k=1;k<=m;k++){
      var last = (k===m);
      var lai = P * (r*cycle);
      var goc = last ? P : 0;
      var tong = goc + lai;
      var duNo = last ? 0 : P;
      rows.push({ky:k, ngay: _addMonths_(start, k*cycle), goc:goc, lai:lai, tong:tong, duNo:duNo});
    }
  }
  return {rows: rows};
}
function buildSchedule_api(payload){
  mustHaveHybrid(payload, 'loanmgr');
  return buildSchedule(payload);
}

/** Lấy đối tượng HĐ từ sheet HỢP ĐỒNG theo Mã HĐ */
function _getLoanByCode_(mahd){
  var L = _getSheetValues_('HỢP ĐỒNG');
  var H = _headerMap_(L.headers);
  var idx_id     = H['Mã HĐ'] ?? H['MaHD'];
  var idx_name   = H['Tên KH'] ?? H['TenKH'];
  var idx_amt    = H['Số tiền vay'] ?? H['Gốc'] ?? H['Principal'];
  var idx_rate   = H['Lãi suất (%/tháng)'] ?? H['Lãi/tháng'] ?? H['RatePerMonth'];
  var idx_months = H['Kỳ hạn (tháng)'];
  var idx_type   = H['Hình thức'] ?? H['Loại'] ?? H['LoaiHD'];
  var idx_start  = H['Ngày giải ngân'] ?? H['Bắt đầu'] ?? H['StartDate'];
  var idx_cycle  = H['Chu kỳ (tháng)'];
  var idx_status = H['Trạng thái'];

  var row = (L.rows||[]).find(function(r){
    var code = String(_pick_(r, idx_id)||'').trim();
    return code.toLowerCase() === String(mahd).toLowerCase();
  });
  if (!row) return null;

  return {
    id    : _pick_(row, idx_id),
    name  : _pick_(row, idx_name),
    P     : _toNum_(_pick_(row, idx_amt)),
    r     : Number(_pick_(row, idx_rate)||0)/100,
    m     : Math.max(1, _toNum_(_pick_(row, idx_months))),
    type  : String(_pick_(row, idx_type)||'Cuốn chiếu').trim(),
    start : _toDate_(_pick_(row, idx_start)) || new Date(),
    cycle : Math.max(1, _toNum_(_pick_(row, idx_cycle)) || 1),
    status: String(_pick_(row, idx_status)||'').trim()
  };
}

/** Phiếu theo kỳ (giữ) */
function getLoanSlip(mahd, refISO, forcedK){
  if (!mahd) throw new Error('Thiếu Mã HĐ');
  var loan = _getLoanByCode_(mahd);
  if (!loan) return { ok:false, message:'Không tìm thấy hợp đồng '+mahd };

  var ref = refISO ? _toDate_(refISO) : new Date();
  if (!ref) ref = new Date();

  var period = _calcPeriodInfo_(loan, ref);
  if (forcedK && forcedK>=1 && forcedK<=loan.m){
    var tmpStart = _addMonths_(loan.start, (forcedK-1)*loan.cycle);
    var tmpEnd   = _endOfPrevDay_(_addMonths_(loan.start, forcedK*loan.cycle));
    var tmp = _calcPeriodInfo_(loan, tmpEnd);
    tmp.k = forcedK; tmp.start = tmpStart; tmp.end = tmpEnd;
    period = tmp;
  }

  // thanh toán trong kỳ
  var Pm = _getSheetValues_('THANH TOÁN');
  var PH = _headerMap_(Pm.headers);
  var i_loan = PH['Mã HĐ']; var i_date = PH['Ngày']; var i_amt  = PH['Số tiền']; var i_type = PH['Loại'] ?? PH['PTTT'];

  var paidGoc=0, paidLai=0;
  (Pm.rows||[]).forEach(function(r){
    var code = _pick_(r, i_loan);
    if (String(code||'').trim().toLowerCase() !== String(mahd).toLowerCase()) return;
    var d = _toDate_(_pick_(r, i_date)); if (!d) return;
    if (d < _startOfDay_(period.start) || d > _startOfDay_(period.end)) return;
    var a = _toNum_(_pick_(r, i_amt));
    var t = String(_pick_(r, i_type)||'').toLowerCase();
    if (t==='gốc') paidGoc += a; else if (t==='lãi') paidLai += a;
  });

  var remainGoc = Math.max(0, period.principalDue - paidGoc);
  var remainLai = Math.max(0, period.interestDue  - paidLai);
  var remainTot = remainGoc + remainLai;

  return {
    ok:true,
    loan: {
      id: loan.id, name: loan.name, type: loan.type,
      principal: loan.P, rateMonth: loan.r, months: loan.m, cycle: loan.cycle,
      start: _ymd_(loan.start), status: loan.status
    },
    period: {
      k: period.k,
      from: _ymd_(period.start), to:_ymd_(period.end),
      principalDue: period.principalDue, interestDue: period.interestDue, totalDue: period.totalDue
    },
    paid: { principal: paidGoc, interest: paidLai, total: paidGoc + paidLai },
    remaining: { principal: remainGoc, interest: remainLai, total: remainTot }
  };
}
function getLoanSlip_api(payload){ mustHaveHybrid(payload, 'loanmgr'); return getLoanSlip(payload.mahd, payload.refISO, payload.k); }

/** Phiếu theo khoảng ngày (tất toán sớm) */
function getLoanSlipRange(mahd, fromISO, toISO){
  if (!mahd) throw new Error('Thiếu Mã HĐ');
  var loan = _getLoanByCode_(mahd);
  if (!loan) return { ok:false, message:'Không tìm thấy hợp đồng '+mahd };

  var from = fromISO ? _startOfDay_(_toDate_(fromISO)) : _startOfDay_(loan.start);
  var to   = toISO   ? _endOfDay_(_toDate_(toISO))     : _endOfDay_(new Date());
  if (from>to) { var tmp=from; from=to; to=tmp; }

  // Lãi pro-rata theo ngày: lãi/tháng (r) → lãi/ngày ~ r/30
  var dailyRate = loan.r / 30;
  var principalDue=0, interestDue=0;

  // Tính gốc đến hạn trong khoảng [from..to] theo mốc kỳ
  var gocPer = (loan.type==='Cuốn chiếu' || loan.type===LOAN_TYPES.CUON_CHIEU) ? (loan.P/loan.m) : 0;
  for (var k=1;k<=loan.m;k++){
    var dueEnd = _endOfPrevDay_(_addMonths_(loan.start, k*loan.cycle)); // mốc đến hạn kỳ k
    if (dueEnd>=from && dueEnd<=to){
      if (loan.type==='Cuốn chiếu' || loan.type===LOAN_TYPES.CUON_CHIEU) principalDue += gocPer;
      else if (k===loan.m) principalDue += loan.P; // cuối kỳ
    }
  }

  // Lãi tích luỹ theo ngày (trên dư nợ đầu mỗi kỳ)
  if (loan.type==='Cuốn chiếu' || loan.type===LOAN_TYPES.CUON_CHIEU){
    for (var k=1;k<=loan.m;k++){
      var periodStart = _addMonths_(loan.start, (k-1)*loan.cycle);
      var periodEnd   = _endOfPrevDay_(_addMonths_(loan.start, k*loan.cycle));
      var overlapStart = new Date(Math.max(periodStart, from));
      var overlapEnd   = new Date(Math.min(periodEnd, to));
      if (overlapStart<=overlapEnd){
        var duNoStart = Math.max(0, loan.P - gocPer*(k-1));
        var days = Math.floor( ( _startOfDay_(overlapEnd) - _startOfDay_(overlapStart) ) / (24*3600*1000) ) + 1;
        interestDue += duNoStart * dailyRate * days;
      }
    }
  } else {
    // Cuối kỳ: dư nợ không đổi cho tới kỳ cuối
    var overlapStart2 = from;
    var overlapEnd2   = to;
    var days2 = Math.floor( ( _startOfDay_(overlapEnd2) - _startOfDay_(overlapStart2) ) / (24*3600*1000) ) + 1;
    // Không vượt quá tổng thời gian m*cycle tháng
    var lastDue = _endOfPrevDay_(_addMonths_(loan.start, loan.m*loan.cycle));
    var clippedEnd = new Date(Math.min(overlapEnd2, lastDue));
    if (overlapStart2<=clippedEnd){
      var daysClip = Math.floor( ( _startOfDay_(clippedEnd) - _startOfDay_(overlapStart2) ) / (24*3600*1000) ) + 1;
      interestDue = loan.P * dailyRate * Math.max(0, daysClip);
    } else interestDue = 0;
  }

  // Đã trả trong khoảng
  var Pm = _getSheetValues_('THANH TOÁN');
  var PH = _headerMap_(Pm.headers);
  var i_loan = PH['Mã HĐ']; var i_date = PH['Ngày']; var i_amt  = PH['Số tiền']; var i_type = PH['Loại'] ?? PH['PTTT'];
  var paidGoc=0, paidLai=0;
  (Pm.rows||[]).forEach(function(r){
    if (String(_pick_(r, i_loan)||'').trim().toLowerCase() !== String(mahd).toLowerCase()) return;
    var d = _toDate_(_pick_(r, i_date)); if (!d) return;
    if (d < from || d > to) return;
    var a = _toNum_(_pick_(r, i_amt));
    var t = String(_pick_(r, i_type)||'').toLowerCase();
    if (t==='gốc') paidGoc += a; else if (t==='lãi') paidLai += a;
  });

  var remainGoc = Math.max(0, principalDue - paidGoc);
  var remainLai = Math.max(0, interestDue  - paidLai);

  return {
    ok:true,
    loan: {
      id: loan.id, name: loan.name, type: loan.type,
      principal: loan.P, rateMonth: loan.r, months: loan.m, cycle: loan.cycle,
      start: _ymd_(loan.start), status: loan.status
    },
    range: { from:_ymd_(from), to:_ymd_(to) },
    due: {
      principal: principalDue, interest: interestDue, total: principalDue+interestDue
    },
    paid: {
      principal: paidGoc, interest: paidLai, total: paidGoc+paidLai
    },
    remaining: {
      principal: Math.max(0,remainGoc), interest: Math.max(0,remainLai), total: Math.max(0,remainGoc)+Math.max(0,remainLai)
    }
  };
}
function getLoanSlipRange_api(payload){
  mustHaveHybrid(payload, 'loanmgr');
  return getLoanSlipRange(payload.mahd, payload.fromISO, payload.toISO);
}

/** Kỳ sắp tới */
function getUpcomingSchedule(mahd, refISO, count){
  count = Math.max(1, Number(count||12));
  var loan = _getLoanByCode_(mahd);
  if (!loan) return {rows:[]};

  var rows=[];
  var k0 = _calcPeriodInfo_(loan, refISO ? _toDate_(refISO) : new Date()).k;
  var start = loan.start;
  for (var k=k0; k<=loan.m && rows.length < count; k++){
    var dueEnd = _endOfPrevDay_(_addMonths_(start, k*loan.cycle));
    var info = _calcPeriodInfo_(loan, dueEnd);
    rows.push({k:k, due:_ymd_(dueEnd), goc:info.principalDue, lai:info.interestDue, tong:info.totalDue});
  }
  return { ok:true, rows:rows, currentK:k0 };
}
function getUpcomingSchedule_api(payload){ mustHaveHybrid(payload, 'loanmgr'); return getUpcomingSchedule(payload.mahd, payload.refISO, payload.count); }

/** Nhắc việc hôm nay */
function getRemindersToday(todayISO){
  var today = todayISO ? _startOfDay_(_toDate_(todayISO)) : _startOfDay_(new Date());

  var L = _getSheetValues_('HỢP ĐỒNG');
  var H = _headerMap_(L.headers);
  var idx_id     = H['Mã HĐ'] ?? H['MaHD'];
  var idx_name   = H['Tên KH'] ?? H['TenKH'];
  var idx_amt    = H['Số tiền vay'] ?? H['Gốc'] ?? H['Principal'];
  var idx_rate   = H['Lãi suất (%/tháng)'] ?? H['Lãi/tháng'] ?? H['RatePerMonth'];
  var idx_months = H['Kỳ hạn (tháng)'];
  var idx_type   = H['Hình thức'] ?? H['Loại'] ?? H['LoaiHD'];
  var idx_start  = H['Ngày giải ngân'] ?? H['Bắt đầu'] ?? H['StartDate'];
  var idx_cycle  = H['Chu kỳ (tháng)'];
  var idx_close  = H['Ngày tất toán'] ?? H['Kết thúc'] ?? H['EndDate'];
  var idx_status = H['Trạng thái'];

  var dueToday = [], closeToday = [];

  (L.rows||[]).forEach(function(r){
    var id    = _pick_(r, idx_id);
    var name  = _pick_(r, idx_name);
    var P     = _toNum_(_pick_(r, idx_amt));
    var rmon  = Number(_pick_(r, idx_rate)||0)/100;
    var m     = Math.max(1, _toNum_(_pick_(r, idx_months)));
    var type  = String(_pick_(r, idx_type)||'Cuốn chiếu').trim();
    var start = _toDate_(_pick_(r, idx_start));
    var cycle = Math.max(1, _toNum_(_pick_(r, idx_cycle))||1);
    var status= String(_pick_(r, idx_status)||'').trim();
    var close = _toDate_(_pick_(r, idx_close));

    if (!start) return;
    if (close && _ymd_(close) === _ymd_(today)) {
      closeToday.push({ id:id, name:name, close:_ymd_(close), status:status });
    }
    if (status === 'Hoàn tất') return;

    var cur = _calcPeriodInfo_({type:type,P:P,r:rmon,m:m,start:start,cycle:cycle}, today);
    var dueEnd = _endOfPrevDay_(_addMonths_(start, cur.k * cycle));
    if (_ymd_(dueEnd) === _ymd_(today)) {
      dueToday.push({ id:id, name:name, k:cur.k, due:_ymd_(dueEnd) });
    }
  });

  return { ok:true, dueToday: dueToday, closeToday: closeToday };
}
function getRemindersToday_api(payload){ mustHaveHybrid(payload, 'loanmgr'); return getRemindersToday(payload.todayISO); }

/** ===================== KPI DASHBOARD ===================== */
function getAppConfig(){
  var p = _getProps_();
  var payTypesCsv = p.getProperty('PAY_TYPES') || 'gốc,lãi,phí';
  return {
    fund: Number(p.getProperty('FUND') || 0),
    payTypes: payTypesCsv.split(',').map(s=>String(s||'').trim()).filter(Boolean),
    telegram: {
      bot   : p.getProperty('TG_BOT')  || '',
      chatId: p.getProperty('TG_CHAT') || ''
    }
  };
}
function setAppConfig(payload){
  mustHaveHybrid(payload, 'settings');
  payload = payload || {};
  var p = _getProps_();

  if (typeof payload.fund !== 'undefined'){ p.setProperty('FUND', String(Number(payload.fund || 0))); }
  if (typeof payload.payTypes !== 'undefined'){
    var csv = Array.isArray(payload.payTypes) ? payload.payTypes.join(',') : String(payload.payTypes||'');
    p.setProperty('PAY_TYPES', csv);
  }
  if (payload.telegram){
    if (typeof payload.telegram.bot !== 'undefined'){ p.setProperty('TG_BOT', String(payload.telegram.bot||'')); }
    if (typeof payload.telegram.chatId !== 'undefined'){ p.setProperty('TG_CHAT', String(payload.telegram.chatId||'')); }
  }
  return getAppConfig();
}
function getPaymentTypes(){ return { ok:true, items: getAppConfig().payTypes }; }

/** Tổng hợp KPI */
function getDashboardKpis_api(payload){
  mustHaveHybrid(payload, 'dashboard');

  // loans
  var L = _getSheetValues_('HỢP ĐỒNG'); var LH=_headerMap_(L.headers);
  var idxAmt = LH['Số tiền vay'] ?? LH['Gốc'] ?? LH['Principal'];
  var idxCus = LH['ID Khách'] ?? LH['IDVC'];
  var idxSta = LH['Trạng thái'];
  var idxStart = LH['Ngày giải ngân'] ?? LH['Bắt đầu'] ?? LH['StartDate'];

  var loans = (L.rows||[]).filter(r=> r.join('').trim());
  var totalLoanAmt = loans.reduce((s,r)=> s+_toNum_(_pick_(r,idxAmt)), 0);
  var activeLoans = loans.filter(r=> String(_pick_(r,idxSta)||'') !== 'Hoàn tất');
  var activeCusSet = {};
  activeLoans.forEach(r=>{ var v=String(_pick_(r,idxCus)||''); if (v) activeCusSet[v]=1; });

  // customers
  var C = _getSheetValues_('KHÁCH HÀNG');
  var totalCustomers = (C.rows||[]).filter(r=> r.join('').trim()).length;

  // payments
  var Pm = _getSheetValues_('THANH TOÁN'); var PH=_headerMap_(Pm.headers);
  var idxType = PH['Loại'] ?? PH['PTTT']; var idxAmtP = PH['Số tiền'];
  var sumG=0,sumL=0,sumF=0;
  (Pm.rows||[]).forEach(r=>{
    var t = String(_pick_(r,idxType)||'').toLowerCase();
    var a = _toNum_(_pick_(r,idxAmtP));
    if (t==='gốc') sumG+=a; else if (t==='lãi') sumL+=a; else if (t==='phí') sumF+=a;
  });

  var fund = getAppConfig().fund;
  return {
    ok:true,
    fund: fund,
    totalLoanAmount: totalLoanAmt,
    fundLeft: Math.max(fund-totalLoanAmt,0),
    paidPrincipal: sumG, paidInterest: sumL, paidFee: sumF,
    totalCustomers: totalCustomers,
    activeCustomers: Object.keys(activeCusSet).length,
    totalLoans: loans.length,
    activeLoans: activeLoans.length
  };
}

/** Danh sách HĐ trong tuần + limit */
function listRecentLoans_api(payload){
  mustHaveHybrid(payload, 'dashboard');
  payload = payload || {};
  var limit = Math.max(1, Number(payload.limit||5));

  var L = _getSheetValues_('HỢP ĐỒNG'); var H=_headerMap_(L.headers);
  var idxId=H['Mã HĐ']??H['MaHD'], idxName=H['Tên KH']??H['TenKH']??H['Tên'], idxStart=H['Ngày giải ngân']??H['Bắt đầu']??H['StartDate'], idxAmt=H['Số tiền vay']??H['Gốc']??H['Principal'];
  var now=new Date(); var first=new Date(now), day=(first.getDay()||7); first.setDate(first.getDate()-day+1); first=_startOfDay_(first);
  var last=_endOfDay_(new Date(first)); last.setDate(first.getDate()+6);

  var arr=(L.rows||[]).map(r=>({ id:_pick_(r,idxId), name:_pick_(r,idxName), start:_toDate_(_pick_(r,idxStart)), amount:_toNum_(_pick_(r,idxAmt)) }))
    .filter(o=>o.id && o.start && o.start>=first && o.start<=last)
    .sort((a,b)=> b.start-a.start)
    .slice(0, limit);
  return { ok:true, rows: arr };
}

/** Payments trong tuần + limit */
function listRecentPayments_api(payload){
  mustHaveHybrid(payload, 'dashboard');
  payload = payload || {};
  var limit = Math.max(1, Number(payload.limit||5));

  var P = _getSheetValues_('THANH TOÁN'); var H=_headerMap_(P.headers);
  var idxDate=H['Ngày'], idxType=H['Loại']??H['PTTT'], idxAmt=H['Số tiền'], idxLoan=H['Mã HĐ'];
  var now=new Date(); var first=new Date(now), day=(first.getDay()||7); first.setDate(first.getDate()-day+1); first=_startOfDay_(first);
  var last=_endOfDay_(new Date(first)); last.setDate(first.getDate()+6);

  var arr=(P.rows||[]).map(r=>({ date:_toDate_(_pick_(r,idxDate)), type:_pick_(r,idxType), amount:_toNum_(_pick_(r,idxAmt)), loan:_pick_(r,idxLoan) }))
    .filter(o=>o.date && o.date>=first && o.date<=last)
    .sort((a,b)=> b.date-a.date)
    .slice(0, limit);
  return { ok:true, rows: arr };
}

/** ===================== IMPORT/UPDATE (ở tab Cấu hình sẽ gọi) ===================== */
function importFromFile_api(payload){
  var tbl = payload && payload.tableKey;
  mustHaveHybrid(payload, _menuKeyForTable_(tbl));

  payload = payload || {};
  var fileId = String(payload.fileId||'').trim();
  var sheetName = String(payload.sheetName||'').trim();
  var mode = String(payload.mode||'upsert').toLowerCase();
  if (!fileId) throw new Error('Thiếu fileId');

  var file = DriveApp.getFileById(fileId);
  var ss = SpreadsheetApp.open(file);
  var sh = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
  if (!sh) throw new Error('Không thấy sheet: ' + (sheetName||'(sheet đầu tiên)'));

  var srcVals = sh.getDataRange().getDisplayValues();
  if (!srcVals || srcVals.length<1) return { ok:false, message:'File nguồn rỗng' };
  var srcHeaders = srcVals[0], srcData = srcVals.slice(1);
  var Hsrc = _headerMap_(srcHeaders);

  var cfg = _cms_getSheet_(payload.tableKey);
  var dest = cfg.sheet;
  if (dest.getLastRow()<1 || dest.getLastColumn()<1) throw new Error('Sheet đích chưa có header');

  var destHeaders = dest.getRange(1,1,1,dest.getLastColumn()).getDisplayValues()[0];
  var Hdst = _headerMap_(destHeaders);

  var idColName = cfg.idColName;
  if (mode==='upsert' && typeof Hsrc[idColName]==='undefined'){
    throw new Error('File import thiếu cột ID bắt buộc: '+idColName);
  }

  var last = dest.getLastRow();
  var mapIdToRow = {};
  if (mode==='upsert' && last>=2){
    var existing = dest.getRange(2,1,last-1,destHeaders.length).getValues();
    for (var i=0;i<existing.length;i++){
      var id = String(existing[i][Hdst[idColName]]||'').trim();
      if (id) mapIdToRow[id] = i+2;
    }
  }

  var updated=0, added=0, skipped=0;
  srcData.forEach(function(row){
    if (!row || !row.join('').trim()) { skipped++; return; }
    var out = destHeaders.map(function(h){ return (typeof Hsrc[h]!=='undefined') ? row[Hsrc[h]] : ''; });
    if (mode==='upsert'){
      var id = String(row[Hsrc[idColName]]||'').trim();
      if (!id){ skipped++; return; }
      var at = mapIdToRow[id];
      if (at){ dest.getRange(at,1,1,destHeaders.length).setValues([out]); updated++; }
      else { dest.appendRow(out); added++; }
    } else { dest.appendRow(out); added++; }
  });

  return { ok:true, mode:mode, updated:updated, added:added, skipped:skipped };
}

/** ===================== USER PREFS & ĐỔI MẬT KHẨU ===================== */
function _userProps_(){ return PropertiesService.getUserProperties(); }
function getMyPrefs(){
  var up = _userProps_();
  return { timeFormat: up.getProperty('TIMEFMT') || 'vi-VN', font: up.getProperty('FONT') || 'system' };
}
function setMyPrefs(obj){
  obj = obj || {};
  var up = _userProps_();
  if (obj.timeFormat) up.setProperty('TIMEFMT', String(obj.timeFormat));
  if (obj.font) up.setProperty('FONT', String(obj.font));
  return getMyPrefs();
}
function changeMyPassword(oldPw, newPw){
  oldPw = String(oldPw||''); newPw = String(newPw||'');
  if (!newPw) throw new Error('Mật khẩu mới rỗng.');
  var principal = _me_();
  var s=_accSheet_(); var H=_accMap_(s.headers);
  var iUser=H['User'], iHash=H['PasswordHash'], iSalt=H['Salt'];

  var rowIdx=-1;
  for (var i=0;i<s.rows.length;i++){
    if (String(s.rows[i][iUser]).trim().toLowerCase() === String(principal||'').toLowerCase()){ rowIdx=i; break; }
  }
  if (rowIdx===-1) throw new Error('Không tìm thấy tài khoản nội bộ gắn với email của bạn.');
  var row=s.rows[rowIdx];
  var salt=row[iSalt]||''; var expect=row[iHash]||''; var got=_hash256b64_(oldPw+salt);
  if (got!==expect) throw new Error('Mật khẩu hiện tại không đúng.');
  var newSalt = Utilities.getUuid(); var newHash = _hash256b64_(newPw+newSalt);
  row[iSalt]=newSalt; row[iHash]=newHash;
  s.sheet.getRange(rowIdx+2,1,1,s.headers.length).setValues([row]);
  return { ok:true };
}

/** ===================== TELEGRAM (Text/Ảnh) + GROUPS ===================== */
function _requireTelegramCfg_(){
  var cfg = getAppConfig(); var bot = cfg.telegram.bot, chat = cfg.telegram.chatId;
  if (!bot) throw new Error('Chưa cấu hình Telegram BOT (TG_BOT) trong Cấu hình.');
  return {bot:bot, defaultChat:chat||''};
}
function tg_sendText(payload){
  mustHaveHybrid(payload, 'loanmgr');
  payload = payload || {};
  var txt = String(payload.text||'').trim();
  if (!txt) throw new Error('Thiếu nội dung tin nhắn.');
  var tg = _requireTelegramCfg_(); var chat = payload.chatId || tg.defaultChat;
  if (!chat) throw new Error('Thiếu chat_id. Hãy chọn nhóm trong “Quản lý Telegram”.');

  var url = 'https://api.telegram.org/bot'+tg.bot+'/sendMessage';
  var res = UrlFetchApp.fetch(url, { method:'post', payload:{ chat_id:chat, text:txt, parse_mode:'HTML' }, muteHttpExceptions:true });
  return { ok:(res.getResponseCode()===200), status:res.getResponseCode(), body:res.getContentText() };
}
function tg_sendImageB64(payload){
  mustHaveHybrid(payload, 'loanmgr');
  payload = payload || {};
  var b64 = String(payload.imageB64||'').trim();
  if (!b64) throw new Error('Thiếu imageB64 (base64).');
  var raw = b64.replace(/^data:image\/\w+;base64,/, '');
  var blob = Utilities.newBlob(Utilities.base64Decode(raw), 'image/png', 'slip.png');

  var tg = _requireTelegramCfg_(); var chat = payload.chatId || tg.defaultChat;
  if (!chat) throw new Error('Thiếu chat_id. Hãy chọn nhóm trong “Quản lý Telegram”.');

  var url = 'https://api.telegram.org/bot'+tg.bot+'/sendPhoto';
  var res = UrlFetchApp.fetch(url, { method:'post', payload:{ chat_id: chat, caption:(payload.caption||''), photo: blob }, muteHttpExceptions:true });
  return { ok:(res.getResponseCode()===200), status:res.getResponseCode(), body:res.getContentText() };
}

/** Quản lý nhiều nhóm Telegram (lưu trong ScriptProperties: key TG_GROUPS = JSON array) */
function _getTgGroups_(){
  var p=_getProps_(); try{ return JSON.parse(p.getProperty('TG_GROUPS')||'[]'); }catch(e){ return []; }
}
function _setTgGroups_(arr){ _getProps_().setProperty('TG_GROUPS', JSON.stringify(arr||[])); }
function tg_listGroups(payload){ mustHaveHybrid(payload,'telegram'); return { ok:true, items:_getTgGroups_() }; }
function tg_addGroup(payload){
  mustHaveHybrid(payload,'telegram');
  var name=String(payload.name||'').trim(); var chatId=String(payload.chatId||'').trim(); var purpose=String(payload.purpose||'').trim();
  if(!name||!chatId) throw new Error('Thiếu name/chatId');
  var arr=_getTgGroups_(); if (arr.find(x=>String(x.chatId)===chatId)) throw new Error('chatId đã tồn tại.');
  arr.push({ name:name, chatId:chatId, purpose:purpose||'' });
  _setTgGroups_(arr); return { ok:true, items:arr };
}
function tg_removeGroup(payload){
  mustHaveHybrid(payload,'telegram');
  var chatId=String(payload.chatId||'').trim(); if(!chatId) throw new Error('Thiếu chatId');
  var arr=_getTgGroups_().filter(x=>String(x.chatId)!==chatId); _setTgGroups_(arr); return { ok:true, items:arr };
}
function tg_sendToGroup(payload){
  mustHaveHybrid(payload,'telegram');
  payload=payload||{};
  var chatId=payload.chatId, purpose=payload.purpose, text=payload.text, imageB64=payload.imageB64;
  var arr=_getTgGroups_();
  if (!chatId && purpose){
    var g=arr.find(x=>String(x.purpose||'').toLowerCase()===String(purpose||'').toLowerCase());
    if (g) chatId=g.chatId;
  }
  if (!chatId) throw new Error('Thiếu chatId hoặc purpose.');
  if (imageB64) return tg_sendImageB64({chatId:chatId, imageB64:imageB64, caption:(payload.caption||'')});
  if (text) return tg_sendText({chatId:chatId, text:text});
  throw new Error('Thiếu nội dung gửi (text hoặc imageB64).');
}

/** ===================== CRM ===================== */
function _ensureSheet_(name, headers){
  var ss=SpreadsheetApp.getActive(); var sh=ss.getSheetByName(name);
  if(!sh){ sh=ss.insertSheet(name); sh.getRange(1,1,1,headers.length).setValues([headers]); sh.setFrozenRows(1); }
  return sh;
}
function _crmSheetsEnsure_(){
  _ensureSheet_('CRM_INTERACTIONS', ['ID','ID Khách','Ngày','Kênh','Nội dung','Nhân viên','Tag']);
  _ensureSheet_('CRM_TASKS', ['TaskID','ID Khách','Tiêu đề','Hạn','Trạng thái','Người phụ trách','Ghi chú']);
  _ensureSheet_('CRM_RATINGS', ['ID Khách','Điểm','Tiêu chí','Nhận xét','Cập nhật']);
}
function crm_listInteractions(payload){
  mustHaveHybrid(payload,'crm'); _crmSheetsEnsure_();
  var s=_getSheetValues_('CRM_INTERACTIONS'); return { ok:true, headers:s.headers, rows:s.rows };
}
function crm_saveInteraction(payload){
  mustHaveHybrid(payload,'crm'); _crmSheetsEnsure_();
  var rec=payload&&payload.record||{}; var s=_getSheetValues_('CRM_INTERACTIONS'); var H=_headerMap_(s.headers);
  var id=String(rec['ID']||'').trim();
  var row = s.headers.map(h=> (rec[h]!==undefined)? rec[h] : '');
  if(!id){ var sh=s.sheet; var idIdx=H['ID']; var newId=_nextSeqId_(sh,idIdx,'CRM',6); row[idIdx]=newId; sh.appendRow(row); return {ok:true, action:'create', id:newId}; }
  // update
  var idIdx2=H['ID']; var data=s.sheet.getRange(2,1,s.sheet.getLastRow()-1,s.sheet.getLastColumn()).getValues();
  var found=-1; for (var i=0;i<data.length;i++){ if (String(data[i][idIdx2])===id){ found=i+2; break;} }
  if(found===-1){ s.sheet.appendRow(row); return {ok:true, action:'create', id:id}; }
  s.sheet.getRange(found,1,1,s.headers.length).setValues([row]); return {ok:true, action:'update', id:id};
}
function crm_deleteInteraction(payload){
  mustHaveHybrid(payload,'crm'); var id=String(payload.id||'').trim(); if(!id) throw new Error('Thiếu ID');
  var s=_getSheetValues_('CRM_INTERACTIONS'); var H=_headerMap_(s.headers); var idx=H['ID']; var data=s.sheet.getRange(2,1,s.sheet.getLastRow()-1,s.sheet.getLastColumn()).getValues();
  for (var i=0;i<data.length;i++){ if (String(data[i][idx])===id){ s.sheet.deleteRow(i+2); return {ok:true}; } }
  return {ok:false, msg:'Không tìm thấy ID'};
}
function crm_listTasks(payload){
  mustHaveHybrid(payload,'crm'); _crmSheetsEnsure_(); var s=_getSheetValues_('CRM_TASKS'); return { ok:true, headers:s.headers, rows:s.rows };
}
function crm_saveTask(payload){
  mustHaveHybrid(payload,'crm'); _crmSheetsEnsure_();
  var rec=payload&&payload.record||{}; var s=_getSheetValues_('CRM_TASKS'); var H=_headerMap_(s.headers);
  var id=String(rec['TaskID']||'').trim(); var row = s.headers.map(h=> (rec[h]!==undefined)? rec[h] : '');
  if(!id){ var sh=s.sheet; var idIdx=H['TaskID']; var newId=_nextSeqId_(sh,idIdx,'TASK',6); row[idIdx]=newId; sh.appendRow(row); return {ok:true, action:'create', id:newId}; }
  var idIdx2=H['TaskID']; var data=s.sheet.getRange(2,1,s.sheet.getLastRow()-1,s.sheet.getLastColumn()).getValues();
  var found=-1; for (var i=0;i<data.length;i++){ if (String(data[i][idIdx2])===id){ found=i+2; break;} }
  if(found===-1){ s.sheet.appendRow(row); return {ok:true, action:'create', id:id}; }
  s.sheet.getRange(found,1,1,s.headers.length).setValues([row]); return {ok:true, action:'update', id:id};
}
function crm_setTaskStatus(payload){
  mustHaveHybrid(payload,'crm'); var id=String(payload.id||'').trim(); var status=String(payload.status||'').trim();
  if(!id) throw new Error('Thiếu TaskID');
  var s=_getSheetValues_('CRM_TASKS'); var H=_headerMap_(s.headers); var idx=H['TaskID'], idxSt=H['Trạng thái'];
  var data=s.sheet.getRange(2,1,s.sheet.getLastRow()-1,s.sheet.getLastColumn()).getValues();
  for (var i=0;i<data.length;i++){ if (String(data[i][idx])===id){ data[i][idxSt]=status; s.sheet.getRange(2,1,data.length,s.headers.length).setValues(data); return {ok:true}; } }
  return {ok:false, msg:'Không tìm thấy TaskID'};
}
function crm_deleteTask(payload){
  mustHaveHybrid(payload,'crm'); var id=String(payload.id||'').trim(); if(!id) throw new Error('Thiếu TaskID');
  var s=_getSheetValues_('CRM_TASKS'); var H=_headerMap_(s.headers); var idx=H['TaskID']; var data=s.sheet.getRange(2,1,s.sheet.getLastRow()-1,s.sheet.getLastColumn()).getValues();
  for (var i=0;i<data.length;i++){ if (String(data[i][idx])===id){ s.sheet.deleteRow(i+2); return {ok:true}; } }
  return {ok:false, msg:'Không tìm thấy TaskID'};
}
function crm_saveRating(payload){
  mustHaveHybrid(payload,'crm'); _crmSheetsEnsure_();
  var rec=payload&&payload.record||{}; var s=_getSheetValues_('CRM_RATINGS'); var H=_headerMap_(s.headers);
  var idk=String(rec['ID Khách']||'').trim(); if(!idk) throw new Error('Thiếu ID Khách');
  var rows=s.sheet.getLastRow()>=2 ? s.sheet.getRange(2,1,s.sheet.getLastRow()-1,s.sheet.getLastColumn()).getValues() : [];
  var idxC=H['ID Khách'];
  var found=-1; for (var i=0;i<rows.length;i++){ if (String(rows[i][idxC])===idk){ found=i+2; break; } }
  var row = s.headers.map(h=> (rec[h]!==undefined)? rec[h] : '');
  row[H['Cập nhật']] = new Date();
  if(found===-1){ s.sheet.appendRow(row); return {ok:true, action:'create'}; }
  s.sheet.getRange(found,1,1,s.headers.length).setValues([row]); return {ok:true, action:'update'};
}
function crm_getRatingsByCustomer(payload){
  mustHaveHybrid(payload,'crm'); var idk=String(payload.customerId||'').trim(); if(!idk) throw new Error('Thiếu ID Khách');
  _crmSheetsEnsure_(); var s=_getSheetValues_('CRM_RATINGS'); var H=_headerMap_(s.headers); var idx=H['ID Khách'];
  var rows=(s.rows||[]).filter(r=> String(_pick_(r,idx))===idk);
  return { ok:true, headers:s.headers, rows:rows };
}

/** ===================== WHOAMI ===================== */
function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
function whoami(payload){
  try { _cms_guard_(); return { ok:true, mode:'google', principal:_me_(), menus:getVisibleMenusForMe([]) }; } catch(e){}
  var token = payload && (payload.session || payload.token);
  var sess  = _sessionFromToken_(token);
  if (sess) return { ok:true, mode:'local', principal: sess.user, display: sess.display, menus: sess.menus };
  return { ok:false };
}

/** ===================== TIỆN ÍCH ===================== */
function fillMissingCustomerIDs(){
  var cfg  = _cms_getSheet_('CUSTOMERS'); var sh=cfg.sheet; var last = sh.getLastRow();
  if (last < 2) return {updated:0};

  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  var H = _headerMap_(headers);
  var idIdx = H['ID Khách']; if (idIdx === undefined) throw new Error('Thiếu cột "ID Khách"');

  var vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  var updated = 0;
  for (var i=0;i<vals.length;i++){
    var id = String(vals[i][idIdx]||'').trim();
    if (!id){ var newId = _nextSeqId_(sh, idIdx, 'VC', 6); vals[i][idIdx] = newId; updated++; }
  }
  if (updated>0){ sh.getRange(2,1,last-1,sh.getLastColumn()).setValues(vals); }
  return {updated:updated};
}


/***** ========== DASHBOARD API ========== *****/

/** Đọc 1 giá trị từ SETTINGS|CẤU HÌNH (KEY|VALUE) */
function getConfigValue_(key) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('SETTINGS') || ss.getSheetByName('CẤU HÌNH');
  if (!sh) return null;
  const rng = sh.getRange(1,1, sh.getLastRow(), 2).getValues(); // A:B
  const m = {};
  rng.forEach(r => { if (r[0]) m[String(r[0]).trim()] = r[1]; });
  return (key in m) ? m[key] : null;
}

/** Tuần hiện tại theo GMT+7 (Asia/Ho_Chi_Minh) — tính từ Thứ 2 → CN */
function _weekRange_() {
  const tz = Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh';
  const now = new Date();
  const today = new Date(Utilities.formatDate(now, tz, 'yyyy-MM-dd') + 'T00:00:00');
  const dow = today.getDay() || 7; // CN=0 → 7
  const monday = new Date(today); monday.setDate(today.getDate() - (dow - 1));
  const sunday = new Date(monday); sunday.setDate(monday.getDate() + 6);
  return { tz, start: monday, end: sunday };
}


/** Đọc dữ liệu Dashboard: Quỹ vốn, tổng vay, còn lại, số HĐ, danh sách tuần, thanh toán tuần */
function getDashboardData(){
  const ss = SpreadsheetApp.getActive();
  const loans = ss.getSheetByName('LOANS');
  const pays  = ss.getSheetByName('PAYMENTS');

  const QUY_VON = Number(getConfigValue_('QUY_VON')) || 0;

  let tongSoTienVay = 0;
  let tongSoTienConLai = 0;
  let soHopDong = 0;

  const { tz, start, end } = _weekRange_();
  const wkStartStr = Utilities.formatDate(start, tz, 'yyyy-MM-dd');
  const wkEndStr   = Utilities.formatDate(end,   tz, 'yyyy-MM-dd');

  /** HỢP ĐỒNG TẠO TRONG TUẦN */
  const hdTrongTuan = [];

  /** THANH TOÁN TRONG TUẦN */
  const ttTrongTuan = [];

  if (loans && loans.getLastRow() > 1){
    const lr = loans.getLastRow(), lc = loans.getLastColumn();
    const vals = loans.getRange(1,1,lr,lc).getValues();
    const header = vals.shift();
    const idx = (name) => header.findIndex(h => String(h).trim().toLowerCase() === name.toLowerCase());
    const iMaHD   = Math.max(idx('MaHD'), idx('Mã HĐ'));
    const iNgayGN = Math.max(idx('NgayGiaiNgan'), idx('Ngày giải ngân'));
    const iSoVay  = Math.max(idx('SoTienVay'), idx('Số tiền vay'));
    const iConLai = Math.max(idx('SoTienConLai'), idx('Số tiền vay còn lại'));
    const iTenKH  = Math.max(idx('TenKH'), idx('Tên KH'), idx('Khach hang'), idx('Khách hàng'));

    vals.forEach(r => {
      const ma = (iMaHD>=0)? r[iMaHD] : '';
      const so = Number((iSoVay>=0)? r[iSoVay] : 0) || 0;
      const cl = Number((iConLai>=0)? r[iConLai] : 0) || 0;
      const ten= (iTenKH>=0)? r[iTenKH] : '';
      const d  = toDate_((iNgayGN>=0)? r[iNgayGN] : null);

      if (so) tongSoTienVay += so;
      if (cl) tongSoTienConLai += cl;
      if (ma) soHopDong++;

      // trong tuần (inclusive)
      if (d && d >= start && d <= end){
        hdTrongTuan.push({
          MaHD: ma, TenKH: ten,
          NgayGiaiNgan: Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
          SoTienVay: so
        });
      }
    });
  }

  if (pays && pays.getLastRow() > 1){
    const lr = pays.getLastRow(), lc = pays.getLastColumn();
    const vals = pays.getRange(1,1,lr,lc).getValues();
    const header = vals.shift();
    const idx = (name) => header.findIndex(h => String(h).trim().toLowerCase() === name.toLowerCase());
    const iMaHD   = Math.max(idx('MaHD'), idx('Mã HĐ'));
    const iNgayTT = Math.max(idx('NgayTT'), idx('Ngày thanh toán'));
    const iGoc    = Math.max(idx('GocTra'), idx('Gốc trả'), idx('Goc'));
    const iLai    = Math.max(idx('LaiTra'), idx('Lãi trả'), idx('Lai'));
    const iTong   = Math.max(idx('TongTra'), idx('Tổng trả'), idx('Tong'));

    vals.forEach(r => {
      const ma = (iMaHD>=0)? r[iMaHD] : '';
      const d  = toDate_((iNgayTT>=0)? r[iNgayTT] : null);
      const goc  = Number((iGoc>=0)? r[iGoc] : 0) || 0;
      const lai  = Number((iLai>=0)? r[iLai] : 0) || 0;
      const tong = Number((iTong>=0)? r[iTong] : (goc+lai)) || 0;

      if (d && d >= start && d <= end){
        ttTrongTuan.push({
          MaHD: ma,
          NgayTT: Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
          GocTra: goc, LaiTra: lai, TongTra: tong
        });
      }
    });
  }

  return {
    ok: true,
    tz, week: { start: wkStartStr, end: wkEndStr },
    kpi: {
      QUY_VON,
      TONG_VAY: tongSoTienVay,
      CON_LAI:  tongSoTienConLai,
      SO_HD:    soHopDong
    },
    hdTrongTuan,
    ttTrongTuan
  };
}

/***** ===================== CUSTOMERS API ===================== *****/
/** Sheet & header helpers */
function _getCustomersSheet_() {
  const ss = SpreadsheetApp.getActive();
  // Ưu tiên sheet chuẩn CRM_CUSTOMERS; nếu bạn dùng tên khác, đổi ở đây
  const sh = ss.getSheetByName('CRM_CUSTOMERS') || ss.getSheetByName('CUSTOMERS') || ss.getSheetByName('KHACH_HANG');
  if (!sh) throw new Error('Không tìm thấy sheet CRM_CUSTOMERS');
  return sh;
}
function _hdrIndex_(header, nameVariants) {
  const low = header.map(h => String(h||'').trim().toLowerCase());
  for (let v of nameVariants) {
    const i = low.indexOf(String(v).toLowerCase());
    if (i >= 0) return i;
  }
  return -1;
}
/** Chuẩn hóa bản ghi đọc từ sheet */
function _rowToCustomer_(header, row) {
  // Dò nhiều biến thể nhãn cột
  const iId   = _hdrIndex_(header, ['idvc','id','mã kh','ma kh','ma_kh']);
  const iName = _hdrIndex_(header, ['tên kh','ten kh','khách hàng','khach hang','name']);
  const iPhone= _hdrIndex_(header, ['điện thoại','dien thoai','sdt','phone']);
  const iAddr = _hdrIndex_(header, ['địa chỉ','dia chi','address']);
  const iDob  = _hdrIndex_(header, ['ngày sinh','ngay sinh','dob','birth']);

  return {
    IDVC:   iId   >=0 ? row[iId]   : '',
    TenKH:  iName >=0 ? row[iName] : '',
    Phone:  iPhone>=0 ? row[iPhone]: '',
    Address:iAddr >=0 ? row[iAddr] : '',
    DOB:    iDob  >=0 ? row[iDob]  : ''
  };
}
function _customerToRow_(header, c){
  const arr = new Array(header.length).fill('');
  const map = new Map(header.map((h,i)=>[String(h||'').trim().toLowerCase(), i]));
  function set(colNames, val){
    for (let n of colNames){
      const i = map.get(String(n).toLowerCase());
      if (i>=0){ arr[i] = val; return; }
    }
  }
  set(['idvc','id','mã kh','ma kh','ma_kh'], c.IDVC || '');
  set(['tên kh','ten kh','khách hàng','khach hang','name'], c.TenKH || '');
  set(['điện thoại','dien thoai','sdt','phone'], c.Phone || '');
  set(['địa chỉ','dia chi','address'], c.Address || '');
  set(['ngày sinh','ngay sinh','dob','birth'], c.DOB || '');
  return arr;
}
/** Tạo ID nếu thiếu */
function _genCustomerId_(){
  const t = Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'Asia/Ho_Chi_Minh', 'yyMMddHHmmss');
  return 'KH' + t;
}

/** Liệt kê khách hàng có phân trang + tìm kiếm */
function listCustomersApi(page, pageSize, q){
  page = Math.max(1, Number(page)||1);
  pageSize = Math.min(200, Math.max(5, Number(pageSize)||20));
  q = String(q||'').trim().toLowerCase();

  const sh = _getCustomersSheet_();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return { ok:true, total:0, page, pageSize, rows:[] };

  const data = sh.getRange(1,1,lr,lc).getValues();
  const header = data.shift().map(h=>String(h||'').trim());
  const rows = data.map(r => _rowToCustomer_(header, r));

  let filtered = rows;
  if (q){
    filtered = rows.filter(r => {
      const s = (r.IDVC+'|'+r.TenKH+'|'+r.Phone+'|'+r.Address).toLowerCase();
      return s.includes(q);
    });
  }

  const total = filtered.length;
  const start = (page-1)*pageSize;
  const end = Math.min(total, start+pageSize);
  const pageRows = (start<end) ? filtered.slice(start, end) : [];

  return { ok:true, total, page, pageSize, rows: pageRows };
}

/** Lấy danh sách option cho dropdown Hợp đồng */
function getCustomerOptions(){
  const sh = _getCustomersSheet_();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return { ok:true, options:[] };

  const data = sh.getRange(1,1,lr,lc).getValues();
  const header = data.shift().map(h=>String(h||'').trim());
  const rows = data.map(r => _rowToCustomer_(header, r));
  const options = rows
    .filter(r => r.IDVC && r.TenKH)
    .map(r => ({ value: r.IDVC, label: `${r.TenKH} — ${r.IDVC}${r.Phone ? ' — '+r.Phone : ''}` }));
  return { ok:true, options };
}

/** Thêm/Sửa khách hàng theo IDVC (nếu thiếu thì tự cấp) */
function upsertCustomer(c){
  const sh = _getCustomersSheet_();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 1) throw new Error('CRM_CUSTOMERS chưa có header');

  const header = sh.getRange(1,1,1,lc).getValues()[0].map(h=>String(h||'').trim());
  const id = String(c.IDVC||'').trim() || _genCustomerId_();

  // Tìm dòng theo IDVC
  const idColIdx = _hdrIndex_(header, ['idvc','id','mã kh','ma kh','ma_kh']);
  if (idColIdx < 0) throw new Error('Không tìm thấy cột IDVC/ID/Mã KH');
  const colA = sh.getRange(2, idColIdx+1, Math.max(0, lr-1), 1).getValues().map(r=>String(r[0]||'').trim());
  const rowIndex = colA.findIndex(v => v === id);

  const normalized = {
    IDVC: id,
    TenKH: String(c.TenKH||'').trim(),
    Phone: String(c.Phone||'').trim(),
    Address: String(c.Address||'').trim(),
    DOB: c.DOB || ''
  };
  const outRow = _customerToRow_(header, normalized);

  if (rowIndex >= 0){
    // update
    sh.getRange(2 + rowIndex, 1, 1, header.length).setValues([outRow]);
  } else {
    // append
    sh.getRange(lr+1, 1, 1, header.length).setValues([outRow]);
  }
  return { ok:true, customer: normalized };
}

/** Xoá theo IDVC */
function deleteCustomer(idvc){
  const id = String(idvc||'').trim();
  if (!id) return { ok:false, error:'Thiếu IDVC' };
  const sh = _getCustomersSheet_();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return { ok:true, deleted:false };

  const header = sh.getRange(1,1,1,lc).getValues()[0].map(h=>String(h||'').trim());
  const idColIdx = _hdrIndex_(header, ['idvc','id','mã kh','ma kh','ma_kh']);
  if (idColIdx < 0) throw new Error('Không tìm thấy cột IDVC/ID/Mã KH');

  const col = sh.getRange(2, idColIdx+1, lr-1, 1).getValues().map(r=>String(r[0]||'').trim());
  const rowIndex = col.findIndex(v => v === id);
  if (rowIndex < 0) return { ok:true, deleted:false };

  sh.deleteRow(2 + rowIndex);
  return { ok:true, deleted:true };
}

/***** ===================== LOANS API ===================== *****/
function _getLoansSheet_(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('LOANS') || ss.getSheetByName('HOP_DONG') || ss.getSheetByName('HOPDONG');
  if (!sh) throw new Error('Không tìm thấy sheet LOANS');
  return sh;
}
function _hdrIdx_(header, variants){
  const low = header.map(h => String(h||'').trim().toLowerCase());
  for (let v of variants){
    const i = low.indexOf(String(v).toLowerCase());
    if (i>=0) return i;
  }
  return -1;
}
function _toDate_(v){
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return v;
  try { return new Date(v); } catch(e){ return null; }
}
function _fmtYMD_(d, tz){
  if (!d) return '';
  return Utilities.formatDate(d, tz || (Session.getScriptTimeZone()||'Asia/Ho_Chi_Minh'), 'yyyy-MM-dd');
}
function _rowToLoan_(header, row){
  const iMaHD   = _hdrIdx_(header, ['mahd','mã hđ','ma hd','id']);
  const iIDVC   = _hdrIdx_(header, ['idvc','mã kh','ma kh','ma_kh','id kh']);
  const iTenKH  = _hdrIdx_(header, ['tenkh','tên kh','khach hang','khách hàng','name']);
  const iLoai   = _hdrIdx_(header, ['loaihd','loại hđ','loai hop dong','loan_type','loai']);
  const iNgayGN = _hdrIdx_(header, ['ngaygiaingan','ngày giải ngân','ngay giai ngan','start_date','batdau']);
  const iKyHan  = _hdrIdx_(header, ['kyhan','kỳ hạn','ky han','term']);
  const iLaiSuat= _hdrIdx_(header, ['laisuat','lãi suất','lai suat','interest','lãi/tháng']);
  const iSoVay  = _hdrIdx_(header, ['sotienvay','số tiền vay','so tien vay','principal','goc']);
  const iConLai = _hdrIdx_(header, ['sotienconlai','số tiền vay còn lại','so tien vay con lai','du no','con lai']);
  const iTrangThai=_hdrIdx_(header, ['trangthai','trạng thái','status','state']);
  const iGhiChu = _hdrIdx_(header, ['ghichu','ghi chú','note','notes']);

  return {
    MaHD:   iMaHD>=0   ? row[iMaHD]   : '',
    IDVC:   iIDVC>=0   ? row[iIDVC]   : '',
    TenKH:  iTenKH>=0  ? row[iTenKH]  : '',
    LoaiHD: iLoai>=0   ? row[iLoai]   : '',
    NgayGiaiNgan: _toDate_(iNgayGN>=0? row[iNgayGN] : ''),
    KyHan:  iKyHan>=0  ? row[iKyHan]  : '',
    LaiSuat:iLaiSuat>=0? Number(row[iLaiSuat]||0) : 0, // %/tháng
    SoTienVay: iSoVay>=0? Number(row[iSoVay]||0) : 0,
    SoTienConLai: iConLai>=0? Number(row[iConLai]||0) : 0,
    TrangThai: iTrangThai>=0? row[iTrangThai] : '',
    GhiChu: iGhiChu>=0? row[iGhiChu] : ''
  };
}
function _loanToRow_(header, L){
  const arr = new Array(header.length).fill('');
  const map = new Map(header.map((h,i)=>[String(h||'').trim().toLowerCase(), i]));
  function set(cols, val){
    for (let c of cols){
      const i = map.get(String(c).toLowerCase());
      if (i>=0){ arr[i]=val; return; }
    }
  }
  set(['mahd','mã hđ','ma hd','id'], L.MaHD || '');
  set(['idvc','mã kh','ma kh','ma_kh','id kh'], L.IDVC || '');
  set(['tenkh','tên kh','khach hang','khách hàng','name'], L.TenKH || '');
  set(['loaihd','loại hđ','loai hop dong','loan_type','loai'], L.LoaiHD || '');
  set(['ngaygiaingan','ngày giải ngân','ngay giai ngan','start_date','batdau'], L.NgayGiaiNgan || '');
  set(['kyhan','kỳ hạn','ky han','term'], L.KyHan || '');
  set(['laisuat','lãi suất','lai suat','interest','lãi/tháng'], L.LaiSuat || 0);
  set(['sotienvay','số tiền vay','so tien vay','principal','goc'], L.SoTienVay || 0);
  set(['sotienconlai','số tiền vay còn lại','so tien vay con lai','du no','con lai'], L.SoTienConLai || 0);
  set(['trangthai','trạng thái','status','state'], L.TrangThai || '');
  set(['ghichu','ghi chú','note','notes'], L.GhiChu || '');
  return arr;
}
function _genLoanId_(){
  const tz = Session.getScriptTimeZone()||'Asia/Ho_Chi_Minh';
  const t = Utilities.formatDate(new Date(), tz, 'yyMMddHHmmss');
  return 'HD' + t;
}

/** Liệt kê HĐ (phân trang + tìm kiếm) */
function listLoansApi(page, pageSize, q){
  page = Math.max(1, Number(page)||1);
  pageSize = Math.min(200, Math.max(5, Number(pageSize)||20));
  q = String(q||'').trim().toLowerCase();

  const sh = _getLoansSheet_();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return { ok:true, total:0, page, pageSize, rows:[] };

  const data = sh.getRange(1,1,lr,lc).getValues();
  const header = data.shift().map(h=>String(h||'').trim());
  const rows = data.map(r => _rowToLoan_(header, r));

  let filtered = rows;
  if (q){
    filtered = rows.filter(r => {
      const s = (r.MaHD+'|'+r.IDVC+'|'+r.TenKH+'|'+r.LoaiHD+'|'+r.TrangThai+'|'+r.GhiChu).toLowerCase();
      return s.includes(q);
    });
  }

  const total = filtered.length;
  const start = (page-1)*pageSize;
  const end = Math.min(total, start+pageSize);
  const pageRows = (start<end) ? filtered.slice(start, end) : [];

  // Chuẩn hóa ngày về yyyy-MM-dd để client render
  const tz = Session.getScriptTimeZone()||'Asia/Ho_Chi_Minh';
  pageRows.forEach(r => {
    r.NgayGiaiNganStr = _fmtYMD_(r.NgayGiaiNgan, tz);
  });

  return { ok:true, total, page, pageSize, rows: pageRows };
}

/** Thêm/Sửa Hợp đồng (theo MaHD; nếu trống thì cấp tự động) */
function upsertLoan(L){
  const sh = _getLoansSheet_();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 1) throw new Error('LOANS chưa có header');

  const header = sh.getRange(1,1,1,lc).getValues()[0].map(h=>String(h||'').trim());

  const id = String(L.MaHD||'').trim() || _genLoanId_();
  const tz = Session.getScriptTimeZone()||'Asia/Ho_Chi_Minh';
  const ngay = L.NgayGiaiNgan ? _toDate_(L.NgayGiaiNgan) : null;

  // Map TenKH nếu chưa có: lấy từ IDVC ở sheet KH
  let ten = String(L.TenKH||'').trim();
  if (!ten && L.IDVC){
    try{
      const ss = SpreadsheetApp.getActive();
      const kh = ss.getSheetByName('CRM_CUSTOMERS');
      if (kh && kh.getLastRow()>1){
        const data = kh.getRange(1,1,kh.getLastRow(), kh.getLastColumn()).getValues();
        const h = data.shift().map(x=>String(x||'').trim());
        const idxId = _hdrIdx_(h, ['idvc','mã kh','ma kh','ma_kh']);
        const idxName = _hdrIdx_(h, ['tenkh','tên kh','khach hang','khách hàng','name']);
        if (idxId>=0 && idxName>=0){
          const m = new Map(data.map(r => [String(r[idxId]||'').trim(), r[idxName]]));
          ten = m.get(String(L.IDVC).trim()) || '';
        }
      }
    }catch(e){}
  }

  // Chuẩn hóa record để ghi
  const out = {
    MaHD: id,
    IDVC: String(L.IDVC||'').trim(),
    TenKH: ten,
    LoaiHD: String(L.LoaiHD||'').trim(), // 'Cuốn chiếu' | 'Cuối kỳ' | ...
    NgayGiaiNgan: ngay ? _fmtYMD_(ngay, tz) : '',
    KyHan: String(L.KyHan||'').trim(),   // số/tháng hoặc mô tả
    LaiSuat: Number(L.LaiSuat||0),       // %/tháng
    SoTienVay: Number(L.SoTienVay||0),
    SoTienConLai: Number(L.SoTienConLai||L.SoTienVay||0),
    TrangThai: String(L.TrangThai||'Đang vay').trim(),
    GhiChu: String(L.GhiChu||'').trim()
  };

  const outRow = _loanToRow_(header, out);

  // Tìm theo MaHD
  const idxMa = _hdrIdx_(header, ['mahd','mã hđ','ma hd','id']);
  if (idxMa < 0) throw new Error('Không tìm thấy cột Mã HĐ');
  const col = sh.getRange(2, idxMa+1, Math.max(0,lr-1), 1).getValues().map(r=>String(r[0]||'').trim());
  const rIdx = col.findIndex(v => v === out.MaHD);

  if (rIdx >= 0){
    sh.getRange(2 + rIdx, 1, 1, header.length).setValues([outRow]);
  } else {
    sh.getRange(lr+1, 1, 1, header.length).setValues([outRow]);
  }
  return { ok:true, loan: out };
}

/** Xoá theo MaHD */
function deleteLoan(mahd){
  const id = String(mahd||'').trim();
  if (!id) return { ok:false, error:'Thiếu Mã HĐ' };

  const sh = _getLoansSheet_();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return { ok:true, deleted:false };

  const header = sh.getRange(1,1,1,lc).getValues()[0].map(h=>String(h||'').trim());
  const idxMa = _hdrIdx_(header, ['mahd','mã hđ','ma hd','id']);
  if (idxMa < 0) throw new Error('Không tìm thấy cột Mã HĐ');

  const col = sh.getRange(2, idxMa+1, lr-1, 1).getValues().map(r=>String(r[0]||'').trim());
  const rIdx = col.findIndex(v => v === id);
  if (rIdx < 0) return { ok:true, deleted:false };

  sh.deleteRow(2 + rIdx);
  return { ok:true, deleted:true };
}
