/***** admin_auth.gs — Đăng nhập RIÊNG cho Trang Quản trị cấu hình (KHÔNG dùng chung) *****/

/** Bảng lưu tài khoản local cho Admin (plaintext) */
const AADMIN_SHEET = 'CONFIG_ACCOUNTS';

/** Session token (CacheService Script) */
const AADMIN_SESS_PREFIX = 'AADMIN_SESS::';
const AADMIN_SESS_TTL_MIN = 180;

/** ====== Sheet helpers ====== */
function aadmin__ensureSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(AADMIN_SHEET);
  if (!sh) {
    sh = ss.insertSheet(AADMIN_SHEET);
    sh.getRange(1, 1, 1, 5).setValues([['user', 'password', 'display', 'active', 'menus']]);
    sh.setFrozenRows(1);
  }
  // đảm bảo header chuẩn
  const hdr = sh.getRange(1,1,1,5).getDisplayValues()[0].map(s=>String(s||'').trim().toLowerCase());
  const expect = ['user','password','display','active','menus'];
  let needFix = false;
  for (let i=0;i<5;i++){ if (hdr[i] !== expect[i]) { needFix = true; break; } }
  if (needFix) sh.getRange(1,1,1,5).setValues([['user','password','display','active','menus']]);

  // seed admin/admin nếu chưa có
  const lr = sh.getLastRow();
  if (lr < 2) {
    sh.appendRow(['admin', 'admin', 'Quản trị cấu hình', 'TRUE', 'schema,accounts']);
  } else {
    const vals = sh.getRange(2,1,lr-1,5).getValues();
    const hasAdmin = vals.some(r => String(r[0]||'').trim().toLowerCase() === 'admin');
    if (!hasAdmin) sh.appendRow(['admin', 'admin', 'Quản trị cấu hình', 'TRUE', 'schema,accounts']);
  }
  return sh;
}
function aadmin__readAll_() {
  const sh = aadmin__ensureSheet_();
  const lr = sh.getLastRow();
  if (lr < 2) return [];
  const vals = sh.getRange(2,1,lr-1,5).getValues();
  return vals.map(r => ({
    user: String(r[0]||'').trim(),
    password: String(r[1]||'').trim(),
    display: String(r[2]||'').trim() || String(r[0]||'').trim(),
    active: String(r[3]||'').toUpperCase() !== 'FALSE',
    menus: String(r[4]||'').split(',').map(s=>s.trim()).filter(Boolean)
  })).filter(x => x.user);
}
function aadmin__save_(acc){
  const sh = aadmin__ensureSheet_();
  const lr = sh.getLastRow();
  const rows = lr>1 ? sh.getRange(2,1,lr-1,5).getValues() : [];
  let idx = -1;
  for (let i=0;i<rows.length;i++){
    if (String(rows[i][0]).trim().toLowerCase() === String(acc.user).trim().toLowerCase()) { idx = i; break; }
  }
  const row = [
    acc.user,
    acc.password || '',
    acc.display || acc.user,
    acc.active ? 'TRUE' : 'FALSE',
    (Array.isArray(acc.menus) ? acc.menus : String(acc.menus||'').split(',').map(s=>s.trim()).filter(Boolean)).join(',')
  ];
  if (idx >= 0) {
    sh.getRange(idx+2,1,1,5).setValues([row]);
  } else {
    sh.appendRow(row);
  }
}
function aadmin__delete_(user){
  const sh = aadmin__ensureSheet_();
  const lr = sh.getLastRow();
  if (lr < 2) return;
  const vals = sh.getRange(2,1,lr-1,5).getValues();
  for (let i=0;i<vals.length;i++){
    if (String(vals[i][0]).trim().toLowerCase() === String(user).trim().toLowerCase()){
      sh.deleteRow(i+2);
      return;
    }
  }
}

/** ====== Session helpers ====== */
function aadmin__rand_(){ return Utilities.getUuid().replace(/-/g,''); }
function aadmin__setSess_(token, payload){
  CacheService.getScriptCache().put(AADMIN_SESS_PREFIX + token, JSON.stringify(payload||{}), 60 * AADMIN_SESS_TTL_MIN);
}
function aadmin__getSess_(token){
  if (!token) return null;
  const raw = CacheService.getScriptCache().get(AADMIN_SESS_PREFIX + token);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(_){ return null; }
}
function aadmin__delSess_(token){
  if (!token) return;
  CacheService.getScriptCache().remove(AADMIN_SESS_PREFIX + token);
}

/** ====== Body parse (support google.script.run & webapp) ====== */
function aadmin__parse_(input){
  if (input && typeof input === 'object' && !Array.isArray(input) && !('parameter' in input)) return input;
  const e = input || {};
  if (e.parameter) return e.parameter;
  if (e.postData && e.postData.contents) { try { return JSON.parse(e.postData.contents); } catch(_){ } }
  return {};
}

/** ====== PUBLIC API (gọi từ admin.html) — tất cả prefix aadmin_* ====== */

/** Đăng nhập */
function aadmin_login(payload){
  const body = aadmin__parse_(payload);
  const user = String(body.user || body.username || body.account || '').trim();
  const pass = String(body.password || body.pass || body.pw || '').trim();
  if (!user || !pass) throw new Error('Thiếu tài khoản hoặc mật khẩu.');

  const list = aadmin__readAll_();
  const acc = list.find(x => x.user.toLowerCase() === user.toLowerCase());
  if (!acc) throw new Error('Tài khoản không tồn tại.');
  if (!acc.active) throw new Error('Tài khoản đã bị khoá.');
  if (String(acc.password) !== String(pass)) throw new Error('Sai mật khẩu.');

  const token = aadmin__rand_();
  const payloadOut = { user: acc.user, display: acc.display || acc.user, menus: acc.menus || ['schema'], mode: 'local' };
  aadmin__setSess_(token, payloadOut);

  return { ok:true, token, principal: acc.user, display: acc.display || acc.user, menus: payloadOut.menus, mode:'local' };
}

/** Đăng xuất */
function aadmin_logout(payload){
  const body = aadmin__parse_(payload);
  const t = String(body.session || body.token || '').trim();
  if (t) aadmin__delSess_(t);
  return { ok:true };
}

/** Kiểm tra phiên */
function aadmin_whoami(payload){
  const body = aadmin__parse_(payload);
  const t = String(body.session || body.token || '').trim();
  if (!t) return { ok:false, message:'Chưa đăng nhập' };
  const sess = aadmin__getSess_(t);
  if (!sess) return { ok:false, message:'Phiên hết hạn' };
  return { ok:true, principal: sess.user, display: sess.display, menus: sess.menus, mode: 'local' };
}

/** Liệt kê / tạo / sửa / xoá tài khoản RIÊNG cho trang cấu hình */
function aadmin_listLocalUsers(){ 
  return { ok:true, rows: aadmin__readAll_(), allMenus: ['schema','accounts'] };
}
function aadmin_upsertLocalUser(payload){
  const body = aadmin__parse_(payload);
  const user = String(body.user||'').trim();
  if (!user) throw new Error('Thiếu user');
  const current = aadmin__readAll_().find(x => x.user.toLowerCase() === user.toLowerCase());
  aadmin__save_({
    user,
    password: (body.password!=null && body.password!=='') ? String(body.password) : (current ? current.password : ''),
    display : String(body.display||user),
    active  : (body.active===false || String(body.active).toUpperCase()==='FALSE') ? false : true,
    menus   : Array.isArray(body.menus) ? body.menus : String(body.menus||'').split(',').map(s=>s.trim()).filter(Boolean)
  });
  return { ok:true };
}
function aadmin_deleteLocalUser(payload){
  const body = aadmin__parse_(payload);
  const user = String(body.user||'').trim();
  if (!user) throw new Error('Thiếu user');
  aadmin__delete_(user);
  return { ok:true };
}

/** Trả về menu cho phiên hiện tại (lọc theo quyền nếu muốn) */
function aadmin_getVisibleMenusForSession(payload){
  const body = aadmin__parse_(payload);
  const keys = Array.isArray(body.keys) ? body.keys : [];
  const t = String(body.session || body.token || '').trim();
  const sess = t ? aadmin__getSess_(t) : null;
  if (!sess) return keys;
  const allowed = new Set(sess.menus || keys);
  return keys.filter(k => allowed.has(k) || k === 'schema' || k === 'accounts');
}

/** Tiện ích: seed lại admin/admin (chạy tay nếu cần) */
function aadmin_seedAdmin(){
  const sh = aadmin__ensureSheet_();
  SpreadsheetApp.getUi().alert('Đã đảm bảo có tài khoản: admin / admin trong sheet ' + AADMIN_SHEET);
}


// ==== HÀM runApi MỚI (thay thế bản cũ) ====
function runApi(fn, payload, onOk, onFail){
  const req = withSession(payload || {});
  const ok = (typeof onOk === 'function') ? onOk : (() => {});
  const fail = (typeof onFail === 'function') ? onFail : (err => alert(err?.message || err));

  if (typeof google === 'undefined' || !google.script || !google.script.run){
    console.warn('Google Apps Script runtime không sẵn sàng.');
    fail({ message: 'Không thể kết nối tới Apps Script.' });
    return;
  }
  // Ánh xạ sang tên hàm aadmin_*
  const realFn = API_MAP[fn] || fn;

  google.script.run
    .withSuccessHandler(ok)
    .withFailureHandler(err => {
      const msg = err?.message || err;
      const lower = String(msg || '').toLowerCase();
      if (lower.includes('chưa đăng nhập') || lower.includes('cần đăng nhập') || lower.includes('phiên hết hạn')){
        setToken('');
        setAuthState(false);
        showLoginDialog();
        return;
      }
      fail(err);
    })[realFn](req);
}