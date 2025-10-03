/***** schema.gs — quản lý SCHEMA & đồng bộ header (VI chuẩn theo UI) *****/

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

// map tiêu đề EN ↔ VI
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

// Đọc SCHEMA: trả { TABLE_KEY: {sheetName, idCol, cols[], labels{}} }
function readSchema_(){
  const sh = _schemaSheet_();
  const last = sh.getLastRow();
  const heads = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const H = _normHeaders_(heads);

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
    const t = String(r[H['Bảng']]||'').trim();
    const f = String(r[H['Trường']]||'').trim();
    const lb= String(r[H['Nhãn']]||'').trim();
    if (!map[t] || !f) return;
    map[t].cols.push(f);
    if (lb) map[t].labels[f] = lb;
  });

  Object.keys(map).forEach(k=>{
    const id = BASE_TABLES[k].idCol;
    if (!map[k].cols.includes(id)) map[k].cols.unshift(id);
    if (!map[k].labels[id]) map[k].labels[id] = id;
  });
  return map;
}

function schemaList_(){ return readSchema_(); }

function schemaAddField(table, field, label, type, required, enumVals, defVal){
  table = String(table||'').toUpperCase().trim();
  if (!BASE_TABLES[table]) throw new Error('Bảng không hợp lệ: '+table);
  field = String(field||'').trim();
  if (!field) throw new Error('Thiếu Field');
  const sh = _schemaSheet_();
  const order = sh.getLastRow();
  sh.appendRow([table, order, field, label||field, type||'string', !!required, enumVals||'', defVal||'', false]);
  return { ok:true };
}

function schemaDeleteField(table, field){
  const sh = _schemaSheet_();
  const last = sh.getLastRow(); if (last<2) return {ok:true};
  const vals = sh.getRange(2,1,last-1,9).getValues();
  for (let i=0;i<vals.length;i++){
    if (String(vals[i][0]).toUpperCase()===String(table).toUpperCase() &&
        String(vals[i][2])===String(field)){
      sh.deleteRow(i+2); break;
    }
  }
  return { ok:true };
}

function applySchema(){
  const schema = readSchema_();
  const ss = SpreadsheetApp.getActive();
  Object.keys(schema).forEach(k=>{
    const def = schema[k];
    let sh = ss.getSheetByName(def.sheetName);
    if (!sh) sh = ss.insertSheet(def.sheetName);
    const headers = def.cols.length ? def.cols : [BASE_TABLES[k].idCol];
    const cur = sh.getLastColumn() ? sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v||'').trim()) : [];
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

function api_schema_list(){ return readSchema_(); }
function api_schema_add(t,f,l,y,r,e,d){ return schemaAddField(t,f,l,y,r,e,d); }
function api_schema_delete(t,f){ return schemaDeleteField(t,f); }
function api_schema_apply(){ return applySchema(); }

/* ===== RÀ SOÁT & KHỞI TẠO SCHEMA THEO UI ===== */
function seedSchema_(){
  const sh = _schemaSheet_();
  if (sh.getLastRow()>1){ SpreadsheetApp.getUi().alert('SCHEMA đã tồn tại.'); return; }
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

    // HỢP ĐỒNG — khớp UI
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

    // THANH TOÁN — khớp UI
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
  SpreadsheetApp.getUi().alert('Đã tạo SCHEMA mẫu (theo UI) & áp dụng sang sheet dữ liệu.');
}

/* ===== Việt hoá/đồng bộ header các sheet dữ liệu hiện có (nếu trước đó đang EN) ===== */
function renameDataHeadersToVI_(){
  const ss = SpreadsheetApp.getActive();

  const map = {
    'KHÁCH HÀNG': {
      'IDVC':'ID Khách','SH':'Mã KH','TenKH':'Tên','Phone':'SĐT','Email':'Email','Note':'Ghi chú',
      'CreatedAt':'Tạo','UpdatedAt':'Cập nhật','UpdatedBy':'Người sửa'
    },
    'HỢP ĐỒNG': {
      // ánh xạ phổ biến từ EN → VI mới
      'MaHD':'Mã HĐ','IDVC':'ID Khách','TenKH':'Tên KH','LoaiHD':'Hình thức',
      'StartDate':'Ngày giải ngân','TermMonths':'Kỳ hạn (tháng)','RatePerMonth':'Lãi suất (%/tháng)',
      'Principal':'Số tiền vay','CycleMonths':'Chu kỳ (tháng)','EndDate':'Ngày tất toán',
      'Status':'Trạng thái','CreatedAt':'Tạo','UpdatedAt':'Cập nhật','UpdatedBy':'Người sửa'
    },
    'THANH TOÁN': {
      'PaymentID':'PaymentID','MaHD':'Mã HĐ','PayDate':'Ngày','Amount':'Số tiền','Type':'Loại',
      'Note':'Ghi chú','CreatedAt':'Tạo','UpdatedAt':'Cập nhật','UpdatedBy':'Người sửa','Method':'Ghi chú'
    }
  };

  Object.keys(map).forEach(sheetName=>{
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastColumn()===0) return;
    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const newH = headers.map(h => map[sheetName][String(h).trim()] || h);
    sh.getRange(1,1,1,newH.length).setValues([newH]);
  });

  SpreadsheetApp.getUi().alert('Đã đồng bộ header sang tiếng Việt theo UI (KHÁCH HÀNG/HỢP ĐỒNG/THANH TOÁN).');
}
