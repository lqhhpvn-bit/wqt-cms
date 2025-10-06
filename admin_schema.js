/***** admin_schema.gs — API Schema cho Trang Admin (RIÊNG) ******************
 * Lưu schema tại ScriptProperties: AADMIN_SCHEMA_DEFS
 * Cấu trúc:
 * {
 *   "CUSTOMERS": { sheetName, idCol, cols:[], labels:{}, types:{}, required:{}, enums:{}, defaults:{} },
 *   "LOANS":     { ... },
 *   "PAYMENTS":  { ... }
 * }
 *****************************************************************************/

const AADMIN_SCHEMA_KEY = 'AADMIN_SCHEMA_DEFS';

/** ====== Helpers ====== */
function _aadmin_props_(){ return PropertiesService.getScriptProperties(); }
function _aadmin_loadSchema_(){
  try { return JSON.parse(_aadmin_props_().getProperty(AADMIN_SCHEMA_KEY) || '{}'); }
  catch(e){ return {}; }
}
function _aadmin_saveSchema_(obj){
  _aadmin_props_().setProperty(AADMIN_SCHEMA_KEY, JSON.stringify(obj || {}));
}

/** Dựng schema mặc định nếu chưa có */
function _aadmin_defaultSchema_(){
  return {
    CUSTOMERS: {
      sheetName: 'KHÁCH HÀNG',
      idCol: 'ID Khách',
      cols: ['ID Khách','Tên KH','Điện thoại','Địa chỉ','Ngày sinh'],
      labels: {
        'ID Khách':'ID Khách',
        'Tên KH':'Tên KH',
        'Điện thoại':'Điện thoại',
        'Địa chỉ':'Địa chỉ',
        'Ngày sinh':'Ngày sinh'
      },
      types:{}, required:{}, enums:{}, defaults:{}
    },
    LOANS: {
      sheetName: 'HỢP ĐỒNG',
      idCol: 'Mã HĐ',
      cols: ['Mã HĐ','ID Khách','Tên KH','Ngày giải ngân','Kỳ hạn (tháng)','Lãi suất (%/tháng)','Số tiền vay','Trạng thái'],
      labels: {
        'Mã HĐ':'Mã HĐ','ID Khách':'ID Khách','Tên KH':'Tên KH','Ngày giải ngân':'Ngày giải ngân',
        'Kỳ hạn (tháng)':'Kỳ hạn (tháng)','Lãi suất (%/tháng)':'Lãi suất (%/tháng)','Số tiền vay':'Số tiền vay','Trạng thái':'Trạng thái'
      },
      types:{}, required:{}, enums:{}, defaults:{}
    },
    PAYMENTS: {
      sheetName: 'THANH TOÁN',
      idCol: 'Mã giao dịch',
      cols: ['Mã giao dịch','Mã HĐ','Ngày','Loại','Số tiền','Ghi chú'],
      labels: {
        'Mã giao dịch':'Mã giao dịch','Mã HĐ':'Mã HĐ','Ngày':'Ngày','Loại':'Loại','Số tiền':'Số tiền','Ghi chú':'Ghi chú'
      },
      types:{}, required:{}, enums:{}, defaults:{}
    }
  };
}

/** Đảm bảo có schema (nếu chưa có thì seed mặc định dựa trên sheet hiện hữu nếu có) */
function _aadmin_ensureSchema_(){
  let sc = _aadmin_loadSchema_();
  if (Object.keys(sc).length) return sc;

  // Nếu có các sheet, cố gắng đọc header để fill
  const ss = SpreadsheetApp.getActive();
  const defaults = _aadmin_defaultSchema_();
  Object.keys(defaults).forEach(k => {
    const d = defaults[k];
    const sh = ss.getSheetByName(d.sheetName);
    if (sh && sh.getLastRow() >= 1){
      const lastC = sh.getLastColumn();
      const headers = sh.getRange(1,1,1,lastC).getDisplayValues()[0];
      const clean = headers.map(h => String(h||'').trim()).filter(Boolean);
      if (clean.length){
        d.cols = clean.slice();
        d.labels = {};
        clean.forEach(h => d.labels[h] = h);
      }
    }
  });
  _aadmin_saveSchema_(defaults);
  return defaults;
}

/** Quyền: bắt buộc đang đăng nhập ở trang Admin */
function _aadmin_requireLogin_(payload){
  const who = aadmin_whoami(payload);
  if (!who || !who.ok) throw new Error('Chưa đăng nhập');
  return who;
}

/** ===== PUBLIC: Liệt kê schema ===== */
function aadmin_schema_list(payload){
  _aadmin_requireLogin_(payload);
  return _aadmin_ensureSchema_();
}

/** ===== PUBLIC: Thêm cột ===== */
function aadmin_schema_add(payload){
  _aadmin_requireLogin_(payload);
  payload = payload || {};
  const table = String(payload.table || payload.tableKey || '').trim();
  const field = String(payload.field || '').trim();
  if (!table || !field) throw new Error('Thiếu bảng hoặc trường');

  const label = String(payload.label || field);
  const type  = String(payload.type  || 'string');
  const req   = !!payload.required;
  const enums = String(payload.enumValues || '').trim();
  const dft   = String(payload.defaultValue || '');

  const sc = _aadmin_ensureSchema_();
  if (!sc[table]) throw new Error('Sai bảng: '+table);

  const def = sc[table];
  if (!Array.isArray(def.cols)) def.cols = [];
  if (def.cols.indexOf(field) !== -1) throw new Error('Cột đã tồn tại');
  def.cols.push(field);

  def.labels   = def.labels   || {};
  def.types    = def.types    || {};
  def.required = def.required || {};
  def.enums    = def.enums    || {};
  def.defaults = def.defaults || {};

  def.labels[field]   = label;
  def.types[field]    = type;
  def.required[field] = req;
  if (enums)   def.enums[field]    = enums;
  if (dft)     def.defaults[field] = dft;

  _aadmin_saveSchema_(sc);
  return { ok:true };
}

/** ===== PUBLIC: Xoá cột ===== */
function aadmin_schema_delete(payload){
  _aadmin_requireLogin_(payload);
  payload = payload || {};
  const table = String(payload.table || payload.tableKey || '').trim();
  const field = String(payload.field || '').trim();
  if (!table || !field) throw new Error('Thiếu bảng hoặc trường');

  const sc = _aadmin_ensureSchema_();
  if (!sc[table]) throw new Error('Sai bảng: '+table);

  const def = sc[table];
  if (def.idCol && field === def.idCol) throw new Error('Không thể xoá cột ID');
  def.cols = (def.cols || []).filter(c => c !== field);
  if (def.labels)   delete def.labels[field];
  if (def.types)    delete def.types[field];
  if (def.required) delete def.required[field];
  if (def.enums)    delete def.enums[field];
  if (def.defaults) delete def.defaults[field];

  _aadmin_saveSchema_(sc);
  return { ok:true };
}

/** ===== PUBLIC: Áp dụng schema → sheet (ghi header) ===== */
function aadmin_schema_apply(payload){
  _aadmin_requireLogin_(payload);
  payload = payload || {};
  const table = String(payload.table || payload.tableKey || '').trim();
  if (!table) throw new Error('Thiếu bảng');

  const sc = _aadmin_ensureSchema_();
  if (!sc[table]) throw new Error('Sai bảng: '+table);
  const def = sc[table];
  const sheetName = def.sheetName || table;

  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(sheetName);
  if (!sh){
    sh = ss.insertSheet(sheetName);
  }

  const cols = Array.isArray(def.cols) && def.cols.length ? def.cols.slice() : [];
  if (!cols.length) throw new Error('Schema rỗng, chưa có cột để áp dụng.');

  // Ghi header (hàng 1) theo thứ tự def.cols
  sh.clear(); // chỉ sheet này
  sh.getRange(1,1,1,cols.length).setValues([cols]);
  sh.setFrozenRows(1);
  return { ok:true, sheetName, cols };
}
