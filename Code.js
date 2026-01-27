/**
 * ----------------------------------------------------------------------
 * PROJECT: GỬI LINK ZOOM CÁ NHÂN HÓA SAU KHI ĐIỀN FORM
 * AUTHOR: Master T & Trọng
 * DESCRIPTION: Tự động đăng ký Zoom từ Form & Đồng bộ điểm danh vào Sheet
 * ----------------------------------------------------------------------
 */

// --- 1. CẤU HÌNH HỆ THỐNG (Chỉnh trong Script Properties, không hard-code) ---
const CONFIG = {
  // Cấu hình cột trong Google Sheet (Index bắt đầu từ 0: A=0, B=1, C=2...)
  COL_INDEX: {
    EMAIL: 1, // Cột B: Địa chỉ email
    NAME:  2, // Cột C: Họ và tên
    ZALO:  3, // Cột D: Số Zalo
    // Cột ghi kết quả điểm danh (ghi sang cột H, I, J để không đè dữ liệu form)
    RESULT_START: 7 
  }
};

// Các key cần đặt trong Script Properties (Project Settings -> Script properties)
const PROP_KEYS = {
  ACCOUNT_ID: 'ZOOM_ACCOUNT_ID',
  CLIENT_ID: 'ZOOM_CLIENT_ID',
  CLIENT_SECRET: 'ZOOM_CLIENT_SECRET',
  MEETING_ID: 'MEETING_ID'
};

let cachedSettings = null; // cache trong runtime Apps Script

function getSettings() {
  if (cachedSettings) return cachedSettings;
  const props = PropertiesService.getScriptProperties();

  const accountId = props.getProperty(PROP_KEYS.ACCOUNT_ID);
  const clientId = props.getProperty(PROP_KEYS.CLIENT_ID);
  const clientSecret = props.getProperty(PROP_KEYS.CLIENT_SECRET);
  const meetingId = props.getProperty(PROP_KEYS.MEETING_ID); // bắt buộc điền để tránh hard-code

  const missing = [];
  if (!accountId) missing.push(PROP_KEYS.ACCOUNT_ID);
  if (!clientId) missing.push(PROP_KEYS.CLIENT_ID);
  if (!clientSecret) missing.push(PROP_KEYS.CLIENT_SECRET);
  if (!meetingId) missing.push(PROP_KEYS.MEETING_ID);

  if (missing.length) {
    const msg = 'Thiếu Script Properties: ' + missing.join(', ');
    throw new Error(msg);
  }

  cachedSettings = { accountId, clientId, clientSecret, meetingId };
  return cachedSettings;
}

// --- 2. MENU TIỆN ÍCH TRÊN SHEET ---
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Master T Tool')
    .addItem('🔄 Đồng bộ điểm danh Zoom', 'syncAttendance')
    .addToUi();
}

// --- 3. CORE 1: XỬ LÝ KHI CÓ NGƯỜI ĐĂNG KÝ (Real-time) ---
function onFormSubmit(e) {
  try {
    if (!e || !e.values) return;

    // Lấy dữ liệu thô
    const rawName = e.values[CONFIG.COL_INDEX.NAME]; 
    const emailRaw   = e.values[CONFIG.COL_INDEX.EMAIL];
    const rawZalo = e.values[CONFIG.COL_INDEX.ZALO];

    // Chuẩn hóa email và kiểm tra hợp lệ
    const email = (emailRaw || '').toString().trim().toLowerCase();
    if (!isValidEmail(email)) {
      console.error(`Email không hợp lệ, bỏ qua: '${emailRaw}'`);
      return;
    }
    
    // Xử lý Logic Data Cleaning
    const cleanName = standardizeName(rawName);
    
    // Lấy 2 số cuối Zalo (Mặc định '00' nếu lỗi)
    let zaloSuffix = "00";
    if (rawZalo) {
      const strZalo = rawZalo.toString().trim();
      if (strZalo.length >= 2) zaloSuffix = strZalo.slice(-2);
    }

    // Format tên hiển thị Zoom: "26" và "- Nguyễn Văn Minh"
    const zoomFirstName = zaloSuffix;
    const zoomLastName  = `- ${cleanName}`;

    // Gọi API Zoom
    const joinUrl = registerUserToZoom(email, zoomFirstName, zoomLastName);

    // Gửi Email
    if (joinUrl) {
      sendEmailWithUniqueLink(email, cleanName, joinUrl);
    }

  } catch (err) {
    console.error("Lỗi onFormSubmit: " + err.toString());
  }
}

// --- 4. CORE 2: ĐỒNG BỘ ĐIỂM DANH (Post-Meeting) ---
function syncAttendance() {
  const settings = getSettings();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("Chưa có dữ liệu học viên!");
    return;
  }

  // Lấy danh sách Email từ Sheet (Để so khớp)
  // Lấy vùng từ dòng 2 đến dòng cuối, số cột cần lấy dựa trên max index
  const dataRange = sheet.getRange(2, 1, lastRow - 1, CONFIG.COL_INDEX.ZALO + 1); 
  const data = dataRange.getValues();
  
  // Gọi API lấy báo cáo (Hỗ trợ phân trang > 500 người)
  const participants = getZoomReportWithPagination(settings.meetingId);
  
  // Tinh chỉnh dữ liệu báo cáo (Cộng dồn thời gian)
  const reportMap = processReportData(participants);

  // Map dữ liệu xuống từng dòng của Sheet
  const results = data.map(row => {
    const email = row[CONFIG.COL_INDEX.EMAIL];
    const record = reportMap[email];
    
    if (record) {
      // Format giờ vào: HH:mm
      const timeStr = Utilities.formatDate(new Date(record.join_time), "GMT+7", "HH:mm");
      return ["Đã tham gia", record.duration, timeStr];
    } else {
      return ["Vắng", 0, ""];
    }
  });

  // Ghi Batch (Hàng loạt) xuống Sheet -> Tối ưu tốc độ
  // Ghi vào cột E, F, G (Status, Duration, TimeIn)
  sheet.getRange(2, CONFIG.COL_INDEX.RESULT_START, results.length, 3).setValues(results);
  
  SpreadsheetApp.getUi().alert(`Đã đồng bộ xong ${results.length} học viên!`);
}

// --- 5. CÁC HÀM HELPER (API & LOGIC) ---

// Helper: Chuẩn hóa tên Tiếng Việt (Title Case)
function standardizeName(str) {
  if (!str) return "";
  return str.trim().replace(/\s+/g, ' ').toLowerCase().split(' ').map(word => {
    return word.charAt(0).toUpperCase() + word.slice(1);
  }).join(' ');
}

// Helper: API Đăng ký User
function registerUserToZoom(email, firstName, lastName) {
  const settings = getSettings();
  const token = getZoomAccessToken(settings);
  if (!token) return null;

  const url = `https://api.zoom.us/v2/meetings/${settings.meetingId}/registrants`;
  const payload = {
    email: email,
    first_name: firstName,
    last_name: lastName,
    auto_approve: true
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const status = response.getResponseCode();
  const body = response.getContentText();
  const json = parseJsonSafe(body);

  if (!json) {
    console.error(`Zoom register parse error (status ${status}): ${body.slice(0, 400)}`);
    return null;
  }
  if (json.join_url) return json.join_url;

  console.error(`Zoom register failed (status ${status}): ${body.slice(0, 400)}`);
  return null; // Trả về null nếu không có join_url
}

// Helper: API Lấy Report (Vét cạn các trang)
function getZoomReportWithPagination(meetingId) {
  const token = getZoomAccessToken();
  if (!token) return [];

  let allParticipants = [];
  let nextPageToken = "";
  
  do {
    let url = `https://api.zoom.us/v2/report/meetings/${meetingId}/participants?page_size=300`;
    if (nextPageToken) url += `&next_page_token=${nextPageToken}`;

    const options = {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const status = response.getResponseCode();
    const body = response.getContentText();
    const json = parseJsonSafe(body);

    if (json && json.participants) {
      allParticipants = allParticipants.concat(json.participants);
      nextPageToken = json.next_page_token;
    } else {
      console.error(`Zoom report parse/error (status ${status}): ${body.slice(0, 400)}`);
      break;
    }
  } while (nextPageToken);

  return allParticipants;
}

// Helper: Xử lý cộng dồn thời gian từ Report
function processReportData(participants) {
  const map = {};
  participants.forEach(p => {
    const email = p.user_email;
    if (map[email]) {
      map[email].duration += p.duration; // Cộng dồn phút
      // Lấy giờ vào sớm hơn
      if (new Date(p.join_time) < new Date(map[email].join_time)) {
        map[email].join_time = p.join_time;
      }
    } else {
      map[email] = {
        duration: p.duration,
        join_time: p.join_time
      };
    }
  });
  return map;
}

// Helper: Lấy Token OAuth
function getZoomAccessToken(settingsParam) {
  // Lưu Token vào Cache 55 phút để đỡ gọi nhiều lần
  const cache = CacheService.getScriptCache();
  const cachedToken = cache.get('zoom_token');
  if (cachedToken) return cachedToken;

  const settings = settingsParam || getSettings();
  const url = `https://zoom.us/oauth/token?grant_type=account_credentials&account_id=${settings.accountId}`;
  const authBlob = Utilities.base64Encode(settings.clientId + ':' + settings.clientSecret);
  
  const options = {
    method: 'post',
    headers: { 'Authorization': `Basic ${authBlob}` },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const status = response.getResponseCode();
  const body = response.getContentText();
  const json = parseJsonSafe(body);
  
  if (json.access_token) {
    cache.put('zoom_token', json.access_token, 3300); // Cache 55 phút
    return json.access_token;
  } else {
    console.error(`Lỗi lấy Token (status ${status}): ${body}`);
    return null;
  }
}

// Helper: Reset cache và test nhanh token
function resetZoomTokenCache() {
  CacheService.getScriptCache().remove('zoom_token');
}

function testZoomToken() {
  resetZoomTokenCache();
  const token = getZoomAccessToken();
  Logger.log(token ? 'Token OK' : 'Token FAIL');
  return token;
}

// Helper: parse JSON an toàn, tránh crash khi API trả HTML/XML
function parseJsonSafe(body) {
  try {
    return JSON.parse(body);
  } catch (err) {
    return null;
  }
}

// Helper: kiểm tra email cơ bản để tránh 400 Invalid field
function isValidEmail(email) {
  if (!email) return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

// Helper: Gửi Email HTML
function sendEmailWithUniqueLink(email, name, link) {
  const subject = '[Vé tham dự] CHUYÊN ĐỀ: QUY TRÌNH & CÔNG NGHỆ VẬN HÀNH LỚP HỌC ONLINE';
  const template = HtmlService.createTemplateFromFile('EmailTemplate');
  template.name = name;
  template.link = link;
  const htmlBody = template.evaluate().getContent();
  const plainBody =
    `Chào ${name},\n` +
    `Bạn đã đăng ký chuyên đề "Quy trình & Công nghệ vận hành lớp học online".\n` +
    `Link Zoom dành riêng cho bạn: ${link}\n` +
    `Nếu nút trong email không bấm được, hãy dán link này vào trình duyệt.\n` +
    `Hẹn gặp bạn trong lớp!`;
  GmailApp.sendEmail(email, subject, "", {
    htmlBody,
    plainBody,
    from: 'trong@hovantrong.com', // gửi từ alias (cần cấu hình alias trong Gmail trước)
    name: 'Hồ Văn Trọng',
    replyTo: 'trong@hovantrong.com'
  });
}