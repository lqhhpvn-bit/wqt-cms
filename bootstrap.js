/***** bootstrap.gs — Menu ADMIN & tiện ích Việt hoá *****/

function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('ADMIN')
    .addItem('Mở Web Admin (Schema)', 'openAdmin_')
    .addItem('Mở Web Quản trị (Nghiệp vụ)', 'openApp_')
    .addSeparator()
    .addItem('Cấp ID Khách VCxxxxx cho dòng trống', 'fillMissingCustomerIDs_') // tiện ích
    .addSeparator()
    .addItem('Đổi tên sheet sang tiếng Việt', 'renameSheetsToVI_')
    .addItem('Việt hoá tiêu đề SCHEMA', 'viHoaSchemaHeaders_')
    .addItem('Chuyển Field SCHEMA EN → VI', 'migrateSchemaFieldsToVI_')
    // ⬇️ GỌI WRAPPER MỚI (tránh đệ quy)
    .addItem('Đổi header sheet dữ liệu EN → VI', 'doRenameDataHeadersToVI_')
    .addSeparator()
    .addItem('Tạo SCHEMA mẫu (VI)', 'seedSchema_')
    .addItem('Áp dụng SCHEMA → Sheet', 'applySchema')
    .addToUi();
}

// Gọi backend để cấp ID
function fillMissingCustomerIDs_(){
  var res = fillMissingCustomerIDs();
  SpreadsheetApp.getUi().alert('Đã cấp ID cho ' + (res.updated||0) + ' khách hàng trống ID.');
}

/* --------- Mở giao diện --------- */
function openAdmin_(){
  const html = HtmlService.createTemplateFromFile('admin');
  html.csrf  = (typeof _csrfToken_ === 'function') ? _csrfToken_() : '';
  html.uiCfg = (typeof _uiCfg_      === 'function') ? _uiCfg_()      : { pageSize:20 };
  SpreadsheetApp.getUi().showSidebar(html.evaluate().setTitle('Web Admin (Schema)'));
}
function openApp_(){
  const html = HtmlService.createTemplateFromFile('cms_index');
  html.serverEnv = { MENU_CONFIG: MENU_CONFIG, LOAN_TYPES: LOAN_TYPES };
  SpreadsheetApp.getUi().showSidebar(
    html.evaluate().setTitle('Web Quản trị')
  );
}

/* --------- Đổi tên TAB sheet sang tiếng Việt --------- */
function renameSheetsToVI_(){
  const ss = SpreadsheetApp.getActive();
  const mapping = [
    ['CRM_CUSTOMERS','KHÁCH HÀNG'],
    ['LOANS','HỢP ĐỒNG'],
    ['PAYMENTS','THANH TOÁN']
  ];
  mapping.forEach(([oldName,newName])=>{
    const sh = ss.getSheetByName(oldName);
    if (sh){
      if (ss.getSheetByName(newName)) throw new Error('Đã tồn tại sheet: ' + newName);
      sh.setName(newName);
    }
  });
  SpreadsheetApp.getUi().alert('Đã đổi tên sheet sang tiếng Việt (KHÁCH HÀNG / HỢP ĐỒNG / THANH TOÁN).');
}

/* --------- Việt hoá tiêu đề SCHEMA (hàng 1) --------- */
function viHoaSchemaHeaders_(){
  const sh = _schemaSheet_();
  sh.getRange(1,1,1,9).setValues([['Bảng','Thứ tự','Trường','Nhãn','Kiểu','Bắt buộc','Lựa chọn','Mặc định','Ẩn']]);
  sh.setFrozenRows(1);
  SpreadsheetApp.getUi().alert('Đã Việt hoá tiêu đề SCHEMA.');
}

/* --------- Chuyển Field trong SCHEMA từ EN → VI (nếu có) --------- */
function migrateSchemaFieldsToVI_(){
  const mapByTable = {
    CUSTOMERS: {
      IDVC:'ID Khách', SH:'Mã KH', TenKH:'Tên', Phone:'SĐT', Email:'Email', Note:'Ghi chú',
      CreatedAt:'Tạo', UpdatedAt:'Cập nhật', UpdatedBy:'Người sửa'
    },
    LOANS: {
      MaHD:'Mã HĐ', IDVC:'ID Khách', TenKH:'Tên KH', LoaiHD:'Hình thức',
      StartDate:'Ngày giải ngân', TermMonths:'Kỳ hạn (tháng)', RatePerMonth:'Lãi suất (%/tháng)',
      Principal:'Số tiền vay', CycleMonths:'Chu kỳ (tháng)', EndDate:'Ngày tất toán',
      Status:'Trạng thái', CreatedAt:'Tạo', UpdatedAt:'Cập nhật', UpdatedBy:'Người sửa'
    },
    PAYMENTS: {
      PaymentID:'PaymentID', MaHD:'Mã HĐ', PayDate:'Ngày', Amount:'Số tiền', Type:'Loại',
      Note:'Ghi chú', CreatedAt:'Tạo', UpdatedAt:'Cập nhật', UpdatedBy:'Người sửa'
    }
  };

  const sh = _schemaSheet_();
  const last = sh.getLastRow(); if (last<2){ SpreadsheetApp.getUi().alert('SCHEMA chưa có dữ liệu.'); return; }

  const heads = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const idx = (nameEN, nameVI) => {
    const iEN = heads.findIndex(h => String(h).trim()===nameEN);
    const iVI = heads.findIndex(h => String(h).trim()===nameVI);
    return iVI !== -1 ? iVI : iEN;
  };
  const iTable  = idx('Table','Bảng');
  const iField  = idx('Field','Trường');

  if (iTable === -1 || iField === -1) throw new Error('Không tìm thấy cột "Bảng/Trường" trong SCHEMA.');

  const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  vals.forEach(r=>{
    const t = String(r[iTable]||'').trim();
    const f = String(r[iField]||'').trim();
    if (!mapByTable[t] || !f) return;
    const vi = mapByTable[t][f];
    if (vi && vi!==f) r[iField] = vi;
  });
  sh.getRange(2,1,vals.length,sh.getLastColumn()).setValues(vals);
  SpreadsheetApp.getUi().alert('Đã chuyển Field trong SCHEMA sang tiếng Việt.');
}

/* --------- Wrapper gọi hàm đổi header dữ liệu EN → VI (tránh đệ quy) --------- */
function doRenameDataHeadersToVI_(){
  renameDataHeadersToVI_(); // hàm gốc được định nghĩa trong schema.gs
}
