/***** server.gs — endpoints Web Admin & Web Quản trị (VI) *****/

const __UI_DEFAULT = { pageSize:20, currencies:['VND','USD','LAK'], loanStatuses:['Đang chạy','Hoàn tất','Tạm dừng'] };
function _uiCfg_(){ try{ return (typeof UI_CFG!=='undefined' && UI_CFG) ? UI_CFG : __UI_DEFAULT; }catch(e){ return __UI_DEFAULT; } }

function _requireAuth_(){
  const email = Session.getActiveUser().getEmail();
  if (ALLOWED_EMAILS && ALLOWED_EMAILS.length){
    if (!ALLOWED_EMAILS.includes(email)) throw new Error('Không có quyền. Email: ' + email);
  }
  return email || 'unknown@user';
}
function _csrfToken_(){
  const cache = CacheService.getUserCache();
  let t = cache.get('csrf'); if (!t){ t = Utilities.getUuid(); cache.put('csrf', t, 3600); }
  return t;
}
function _checkCsrf_(t){
  const cur = CacheService.getUserCache().get('csrf');
  if (!cur || cur!==t) throw new Error('CSRF token không hợp lệ');
}

/*************************************************
 * Router view
 *  - ?view=cms     → Web Quản trị (CMS)
 *  - (mặc định)    → Trang admin cũ quản lý SCHEMA
 **************************************************/
function doGet(e) {
  var view = (e && e.parameter && e.parameter.view) || '';

  if (view === 'cms') {
    var t = HtmlService.createTemplateFromFile('cms_index');
    t.serverEnv = { MENU_CONFIG: MENU_CONFIG, LOAN_TYPES: LOAN_TYPES };
    return t.evaluate()
      .setTitle('Web Quản trị')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var tpl = HtmlService.createTemplateFromFile('admin');
  return tpl.evaluate().setTitle('Web Admin (Quản lý SCHEMA)');
}

// Admin (schema) APIs
function getTableDefs(){ return readSchema_(); }

function api_list(tableKey, query, page, pageSize, sortBy, sortDir){
  _requireAuth_();
  const defs = readSchema_();
  const def = defs[tableKey];
  if (!def) throw new Error('Sai bảng: ' + tableKey);
  const rows = _readAsObjects_(def.sheetName);
  return _filterSearchPaginate_(rows, query, page||1, pageSize||_uiCfg_().pageSize, sortBy||def.idCol, sortDir||'desc');
}

function api_save(tableKey, obj, csrf){
  const who = _requireAuth_(); _checkCsrf_(csrf);
  const def = readSchema_()[tableKey]; if (!def) throw new Error('Sai bảng: ' + tableKey);

  // có thể mở rộng normalize_/validate_ nếu dùng admin này cho nghiệp vụ
  const {row, headers} = _findRowById_(def.sheetName, def.idCol, obj[def.idCol]||'');
  const now = new Date();

  if (row === -1){
    if (!obj[def.idCol]) {
      const prefix = tableKey==='CUSTOMERS' ? 'IDVC' : tableKey==='LOANS' ? 'HD' : 'PAY';
      obj[def.idCol] = _genId_(prefix);
    }
    if (headers.includes('Tạo'))      obj['Tạo'] = now;
    if (headers.includes('Cập nhật')) obj['Cập nhật'] = now;
    if (headers.includes('Người sửa')) obj['Người sửa'] = who;
    _appendObject_(def.sheetName, obj);
    return { ok:true, id: obj[def.idCol], mode:'create' };
  } else {
    if (headers.includes('Cập nhật'))   obj['Cập nhật'] = now;
    if (headers.includes('Người sửa'))  obj['Người sửa'] = who;
    const sh = _getSheet_(def.sheetName);
    const base = _readAsObjects_(def.sheetName).find(r => String(r[def.idCol])===String(obj[def.idCol])) || {};
    _writeRowByIndex_(sh, row, headers, Object.assign(base, obj));
    return { ok:true, id: obj[def.idCol], mode:'update' };
  }
}

function api_delete(tableKey, id, csrf){
  _requireAuth_(); _checkCsrf_(csrf);
  const def = readSchema_()[tableKey];
  const sh = _getSheet_(def.sheetName);
  const {row} = _findRowById_(def.sheetName, def.idCol, id);
  if (row === -1) return { ok:false, message:'Không tìm thấy ID' };
  sh.deleteRow(row);
  return { ok:true };
}
