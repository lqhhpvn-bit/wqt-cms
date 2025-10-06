/***** schema.gs — quản lý SCHEMA & đồng bộ header (chuẩn VI theo UI) ******
 * - Sheet cấu hình:  SCHEMA  (Bảng | Thứ tự | Trường | Nhãn | Kiểu | Bắt buộc | Lựa chọn | Mặc định | Ẩn)
 * - Bảng dữ liệu:    CUSTOMERS, LOANS, PAYMENTS
 * - API cho admin:   api_schema_list / api_schema_add / api_schema_delete / api_schema_apply
 * - Seeder:          seedSchema_()
 ****************************************************************************/

/** ====== BẢNG CƠ SỞ MẶC ĐỊNH ====== */

/** ====== BRIDGE AUTH: dùng nếu admin_auth.gs có _requireAuth(), không có thì bỏ qua ====== */
function mustHaveHybrid(payload, /*scope*/ _scope){
  try {
    if (typeof _requireAuth === 'function') _requireAuth(payload || {});
  } catch (e){
    // nếu chưa có _requireAuth (hoặc chưa login) thì ném lỗi để UI biết
    if (String(e && e.message || '').toLowerCase().includes('đăng nhập')) throw e;
    // còn lại: nuốt lỗi để chạy độc lập (dùng trong môi trường không cần auth)
  }
  return true;
}

/** ====== TẠO / LẤY Sheet SCHEMA ====== */
function _schemaSheet_(){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('SCHEMA');
  if (!sh){
    sh = ss.insertSheet('SCHEMA');
    sh.getRange(1,1,1,9).setValues([['Bảng','Thứ tự','Trường','Nhãn','Kiểu','Bắt buộc','Lựa chọn','Mặc định','Ẩn']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** ====== map tiêu đề EN ↔ VI (để đồng bộ khi có file cũ) ====== */
const _HEAD_MAP_ = {
  'Table':'Bảng', 'Order':'Thứ tự', 'Field':'Trường', 'Label':'Nhãn', 'Type':'Kiểu',
  'Required':'Bắt buộc', 'Enum':'Lựa chọn', 'Default':'Mặc định', 'Hidden':'Ẩn'
};
function _normHeaders_(arr){
  const m = {};
  arr.forEach((h,i)=>{
    const s = String(h||'').trim();
    if (s in _HEAD_MAP_) m[_HEAD_MAP_[s]] = i;
    else m[s] = i;
  });
  return m;
}

/** ====== Đọc SCHEMA về object: { TABLE_KEY: { sheetName, idCol, cols[], labels{} } } ====== */
function readSchema_(){
  const sh = _schemaSheet_();
  const last = sh.getLastRow();
  const heads = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const H = _normHeaders_(heads);

  // chưa có data → trả mặc định: mỗi bảng ít nhất có cột ID
  if (last < 2) {
    const map = {};
    Object.keys(BASE_TABLES).forEach(k=>{
      const id = BASE_TABLES[k].idCol;
      map[k] = { ...BASE_TABLES[k], cols:[id], labels:{ [id]: id } };
    });
    return map;
  }

  const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  const map = {};
  Object.keys(BASE_TABLES).forEach(k=> map[k] = { ...BASE_TABLES[k], cols: [], labels:{} });

  vals.forEach(r=>{
    const t  = String(r[H['Bảng']]||'').trim();
    const f  = String(r[H['Trường']]||'').trim();
    const lb = String(r[H['Nhãn']]||'').trim();
    if (!map[t] || !f) return;
    map[t].cols.push(f);
    if (lb) map[t].labels[f] = lb;
  });

  // đảm bảo cột ID luôn đứng đầu & có nhãn
  Object.keys(map).forEach(k=>{
    const id = BASE_TABLES[k].idCol;
    if (!map[k].cols.includes(id)) map[k].cols.unshift(id);
    if (!map[k].labels[id]) map[k].labels[id] = id;
  });

  return map;
}

function schemaList_(){ return readSchema_(); }

/** ====== Ghi thêm field vào SCHEMA (1 dòng) ====== */
function schemaAddField(table, field, label, type, required, enumVals, defVal){
  table = String(table||'').toUpperCase().trim();
  if (!BASE_TABLES[table]) throw new Error('Bảng không hợp lệ: '+table);
  field = String(field||'').trim();
  if (!field) throw new Error('Thiếu Field');

  const sh = _schemaSheet_();
  const order = Math.max(1, sh.getLastRow()); // dùng lastRow làm thứ tự
  sh.appendRow([
    table, order, field, label || field, (type || 'string'),
    !!required, (enumVals || ''), (defVal || ''), false
  ]);
  return { ok:true };
}

/** ====== Xoá field khỏi SCHEMA ====== */
function schemaDeleteField(table, field){
  const sh = _schemaSheet_();
  const last = sh.getLastRow(); if (last<2) return {ok:true};
  const vals = sh.getRange(2,1,last-1,Math.max(9, sh.getLastColumn())).getValues();
  for (let i=0;i<vals.length;i++){
    if (String(vals[i][0]).toUpperCase()===String(table).toUpperCase() &&
        String(vals[i][2])===String(field)){
      sh.deleteRow(i+2); break;
    }
  }
  return { ok:true };
}

/** ====== Áp dụng SCHEMA: tạo headers vào các sheet dữ liệu nếu thiếu ====== */
function applySchema(){
  const schema = readSchema_();
  const ss = SpreadsheetApp.getActive();

  Object.keys(schema).forEach(k=>{
    const def = schema[k];
    let sh = ss.getSheetByName(def.sheetName);
    if (!sh) sh = ss.insertSheet(def.sheetName);

    const headers = def.cols.length ? def.cols : [BASE_TABLES[k].idCol];
    const cur = sh.getLastColumn()
      ? sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v||'').trim())
      : [];

    if (cur.length===0){
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      sh.setFrozenRows(1);
      return;
    }
    const missing = headers.filter(h=>!cur.includes(h));
    if (missing.length){
      sh.getRange(1,cur.length+1,1,missing.length).setValues([missing]);
    }
  });

  return { ok:true };
}

/** ====== API cho admin UI ====== */
function api_schema_list(payload){
  payload = payload || {};
  mustHaveHybrid(payload, 'schema');
  return readSchema_();
}

function api_schema_add(payload, f, l, y, r, e, d){
  // hỗ trợ kiểu gọi cũ: api_schema_add(table, field, label, type, required, enum, default)
  if (typeof payload !== 'object' || payload === null || Array.isArray(payload)){
    return api_schema_add({ table: payload, field: f, label: l, type: y, required: r, enumValues: e, defaultValue: d });
  }
  mustHaveHybrid(payload, 'schema');

  const table = String(payload.table || payload.tbl || payload.tableKey || '').trim();
  const field = String(payload.field || payload.name || '').trim();
  const label = (payload.label !== undefined && payload.label !== null) ? String(payload.label) : (l || field);
  const type  = String(payload.type  || payload.datatype || y || 'string');

  let reqVal = payload.required;
  if (reqVal === undefined) reqVal = payload.req || payload.isRequired || r;
  const required = (typeof reqVal === 'string') ? !/^(false|0|no|off)$/i.test(reqVal) : !!reqVal;

  let enumVals = payload.enumValues;
  if (enumVals === undefined) enumVals = payload.enum || payload.options || e;
  const enumCsv = (enumVals !== undefined && enumVals !== null) ? String(enumVals) : '';

  let defVal = payload.defaultValue;
  if (defVal === undefined) defVal = payload.default || payload.def || d;
  const defOut = (defVal !== undefined && defVal !== null) ? defVal : '';

  return schemaAddField(table, field, label, type, required, enumCsv, defOut);
}

function api_schema_delete(payload, f){
  // hỗ trợ kiểu gọi cũ: api_schema_delete(table, field)
  if (typeof payload !== 'object' || payload === null || Array.isArray(payload)){
    return api_schema_delete({ table: payload, field: f });
  }
  mustHaveHybrid(payload, 'schema');
  const table = payload.table || payload.tbl || payload.tableKey;
  const field = payload.field || payload.name;
  return schemaDeleteField(table, field);
}

function api_schema_apply(payload){
  payload = payload || {};
  mustHaveHybrid(payload, 'schema');
  return applySchema();
}

/** ====== Seeder SCHEMA mẫu khớp UI & áp dụng sang sheet dữ liệu ====== */
function seedSchema_(){
  const sh = _schemaSheet_();
  if (sh.getLastRow()>1){
    try { SpreadsheetApp.getUi().alert('SCHEMA đã tồn tại.'); } catch(_) {}
    return;
  }
  sh.getRange(1,1,1,9).setValues([['Bảng','Thứ tự','Trường','Nhãn','Kiểu','Bắt buộc','Lựa chọn','Mặc định','Ẩn']]);
  sh.setFrozenRows(1);

  const rows = [
    // KHÁCH HÀNG
    ['CUSTOMERS',1,'ID Khách','ID Khách','string',true,'','','false'],
    ['CUSTOMERS',2,'Mã KH','Mã KH','string',false,'','','false'],
    ['CUSTOMERS',3,'Tên','Tên','string',false,'','','false'],
    ['CUSTOMERS',4,'SĐT','SĐT','string',false,'','','false'],
    ['CUSTOMERS',5,'Email','Email','string',false,'','','false'],
    ['CUSTOMERS',6,'Ghi chú','Ghi chú','string',false,'','','false'],
    ['CUSTOMERS',7,'Tạo','Tạo','date',false,'','','true'],
    ['CUSTOMERS',8,'Cập nhật','Cập nhật','date',false,'','','true'],
    ['CUSTOMERS',9,'Người sửa','Người sửa','string',false,'','','true'],

    // HỢP ĐỒNG
    ['LOANS',1,'Mã HĐ','Mã HĐ','string',true,'','','false'],
    ['LOANS',2,'ID Khách','ID Khách','string',true,'','','false'],
    ['LOANS',3,'Tên KH','Tên KH','string',false,'','','false'],
    ['LOANS',4,'Hình thức','Hình thức','enum',true,'Cuốn chiếu,Cuối kỳ','','false'],
    ['LOANS',5,'Ngày giải ngân','Ngày giải ngân','date',true,'','','false'],
    ['LOANS',6,'Kỳ hạn (tháng)','Kỳ hạn (tháng)','number',true,'','','false'],
    ['LOANS',7,'Lãi suất (%/tháng)','Lãi suất (%/tháng)','number',true,'','','false'],
    ['LOANS',8,'Số tiền vay','Số tiền vay','number',true,'','','false'],
    ['LOANS',9,'Chu kỳ (tháng)','Chu kỳ (tháng)','number',false,'','','false'],
    ['LOANS',10,'Ngày tất toán','Ngày tất toán','date',false,'','','false'],
    ['LOANS',11,'Trạng thái','Trạng thái','enum',false,'Đang chạy,Hoàn tất,Quá hạn,Tạm dừng','','false'],
    ['LOANS',12,'Tạo','Tạo','date',false,'','','true'],
    ['LOANS',13,'Cập nhật','Cập nhật','date',false,'','','true'],
    ['LOANS',14,'Người sửa','Người sửa','string',false,'','','true'],

    // THANH TOÁN
    ['PAYMENTS',1,'PaymentID','PaymentID','string',true,'','','false'],
    ['PAYMENTS',2,'Mã HĐ','Mã HĐ','string',true,'','','false'],
    ['PAYMENTS',3,'Ngày','Ngày','date',true,'','','false'],
    ['PAYMENTS',4,'Số tiền','Số tiền','number',true,'','','false'],
    ['PAYMENTS',5,'Loại','Loại','enum',true,'gốc,lãi,phí','','false'],
    ['PAYMENTS',6,'Ghi chú','Ghi chú','string',false,'','','false'],
    ['PAYMENTS',7,'Tạo','Tạo','date',false,'','','true'],
    ['PAYMENTS',8,'Cập nhật','Cập nhật','date',false,'','','true'],
    ['PAYMENTS',9,'Người sửa','Người sửa','string',false,'','','true'],
  ];
  sh.getRange(2,1,rows.length,9).setValues(rows);

  applySchema();
  try { SpreadsheetApp.getUi().alert('Đã tạo SCHEMA mẫu & áp dụng sang sheet dữ liệu.'); } catch(_) {}
}

/** ====== Việt hoá/đồng bộ header cho các sheet dữ liệu đang dùng tên EN (tùy chọn) ====== */
function renameDataHeadersToVI_(){
  const ss = SpreadsheetApp.getActive();
  // ánh xạ các tên cột cũ (EN) → mới (VI) theo UI
  const map = {
    CUSTOMERS: {
      'IDVC':'ID Khách','SH':'Mã KH','TenKH':'Tên','Phone':'SĐT','Email':'Email','Note':'Ghi chú',
      'CreatedAt':'Tạo','UpdatedAt':'Cập nhật','UpdatedBy':'Người sửa'
    },
    LOANS: {
      'MaHD':'Mã HĐ','IDVC':'ID Khách','TenKH':'Tên KH','LoaiHD':'Hình thức',
      'StartDate':'Ngày giải ngân','TermMonths':'Kỳ hạn (tháng)','RatePerMonth':'Lãi suất (%/tháng)',
      'Principal':'Số tiền vay','CycleMonths':'Chu kỳ (tháng)','EndDate':'Ngày tất toán',
      'Status':'Trạng thái','CreatedAt':'Tạo','UpdatedAt':'Cập nhật','UpdatedBy':'Người sửa'
    },
    PAYMENTS: {
      'PaymentID':'PaymentID','MaHD':'Mã HĐ','PayDate':'Ngày','Amount':'Số tiền','Type':'Loại',
      'Note':'Ghi chú','CreatedAt':'Tạo','UpdatedAt':'Cập nhật','UpdatedBy':'Người sửa','Method':'Ghi chú'
    }
  };

  Object.keys(map).forEach(key=>{
    const sheetName = BASE_TABLES[key].sheetName;
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastColumn()===0) return;
    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const newH = headers.map(h => map[key][String(h).trim()] || h);
    sh.getRange(1,1,1,newH.length).setValues([newH]);
  });

  try { SpreadsheetApp.getUi().alert('Đã đồng bộ header sang tiếng Việt cho CUSTOMERS/LOANS/PAYMENTS.'); } catch(_) {}
}
