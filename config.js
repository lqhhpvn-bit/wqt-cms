/***** config.gs — Cấu hình chung cho Web Quản Trị *****/

// (tuỳ chọn) whitelist email. Để mảng rỗng nghĩa là cho tất cả.
const ALLOWED_EMAILS = [
  // "you@example.com"
];

/** Định nghĩa các bảng (sheet) + cột ID + nhãn hiển thị */
const BASE_TABLES = {
  CUSTOMERS: { sheetName: 'KHÁCH HÀNG',  idCol: 'ID Khách',  label: 'Khách hàng' },
  LOANS:     { sheetName: 'HỢP ĐỒNG',    idCol: 'Mã HĐ',     label: 'Hợp đồng'   },
  PAYMENTS:  { sheetName: 'THANH TOÁN',  idCol: 'PaymentID', label: 'Thanh toán' },
  USERS:     { sheetName: 'USERS',       idCol: 'Email',     label: 'Tài khoản'  } // thêm bảng người dùng
};

// Kiểu hợp đồng vay
const LOAN_TYPES = { CUON_CHIEU: 'Cuốn chiếu', CUOI_KY: 'Cuối kỳ' };

/** MENU mới (thứ tự theo yêu cầu) */
const MENU_CONFIG = [
  { key: 'dashboard',  title: 'Tổng quan',                     icon: 'bar-chart-2' },
  { key: 'customers',  title: 'Khách hàng',                    icon: 'users' },
  { key: 'loans',      title: 'Hợp đồng',                      icon: 'file-text' },
  { key: 'loanmgr',    title: 'Quản lý khoản vay/Thanh toán',  icon: 'calculator' }, // gộp
  { key: 'payhistory', title: 'Lịch sử thanh toán',            icon: 'clock' },
  { key: 'crm',        title: 'CRM',                           icon: 'briefcase' },
  { key: 'telegram',   title: 'Nhóm Telegram',                 icon: 'send' },
  { key: 'reports',    title: 'Báo cáo',                       icon: 'pie-chart' },
  { key: 'settings',   title: 'Cấu hình',                      icon: 'sliders' },
  { key: 'account',    title: 'Cài đặt',                       icon: 'settings' }     // hiển thị, font, mật khẩu
];

// Kích thước trang mặc định khi list
const DEFAULT_PAGE_SIZE = 50;
