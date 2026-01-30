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
    PHONE: 7, // Cột H: Số điện thoại
    // Cột ghi kết quả điểm danh (ghi sang cột I, J, K để không đè dữ liệu form)
    RESULT_START: 8 
  }
};

// Các key cần đặt trong Script Properties (Project Settings -> Script properties)
const PROP_KEYS = {
  ACCOUNT_ID: 'ZOOM_ACCOUNT_ID',
  CLIENT_ID: 'ZOOM_CLIENT_ID',
  CLIENT_SECRET: 'ZOOM_CLIENT_SECRET',
  MEETING_ID: 'MEETING_ID',
  // Thông tin lớp học (cho gửi email nhắc nhớ)
  CLASS_NAME: 'CLASS_NAME',
  CLASS_TIME: 'CLASS_TIME',
  CLASS_FORMAT: 'CLASS_FORMAT',
  CLASS_INSTRUCTOR: 'CLASS_INSTRUCTOR'
};

// ===== HƯỚNG DẪN SETUP SCRIPT PROPERTIES =====
// Vào Project Settings → Script properties và điền các key sau:
// 
// ZOOM_ACCOUNT_ID: [Lấy từ Zoom App Marketplace]
// ZOOM_CLIENT_ID: [Lấy từ Zoom App Marketplace]
// ZOOM_CLIENT_SECRET: [Lấy từ Zoom App Marketplace]
// MEETING_ID: [ID của Zoom meeting]
//
// CLASS_NAME: BÍ MẬT VẬN HÀNH LỚP HỌC ONLINE - Tư duy & Công nghệ thực chiến
// CLASS_TIME: 20:30 - 22:00 | Thứ Bảy, ngày 31/01/2026
// CLASS_FORMAT: Trực tuyến qua Zoom
// CLASS_INSTRUCTOR: Hồ Văn Trọng – Chuyên gia IT & Phát triển tâm thức
// ============================================

let cachedSettings = null; // cache trong runtime Apps Script

function getSettings() {
  if (cachedSettings) return cachedSettings;
  const props = PropertiesService.getScriptProperties();

  const accountId = props.getProperty(PROP_KEYS.ACCOUNT_ID);
  const clientId = props.getProperty(PROP_KEYS.CLIENT_ID);
  const clientSecret = props.getProperty(PROP_KEYS.CLIENT_SECRET);
  const meetingId = props.getProperty(PROP_KEYS.MEETING_ID);
  const className = props.getProperty(PROP_KEYS.CLASS_NAME);
  const classTime = props.getProperty(PROP_KEYS.CLASS_TIME);
  const classFormat = props.getProperty(PROP_KEYS.CLASS_FORMAT);
  const classInstructor = props.getProperty(PROP_KEYS.CLASS_INSTRUCTOR);

  const missing = [];
  if (!accountId) missing.push(PROP_KEYS.ACCOUNT_ID);
  if (!clientId) missing.push(PROP_KEYS.CLIENT_ID);
  if (!clientSecret) missing.push(PROP_KEYS.CLIENT_SECRET);
  if (!meetingId) missing.push(PROP_KEYS.MEETING_ID);

  if (missing.length) {
    const msg = 'Thiếu Script Properties: ' + missing.join(', ');
    throw new Error(msg);
  }

  cachedSettings = { 
    accountId, 
    clientId, 
    clientSecret, 
    meetingId,
    className: className || '',
    classTime: classTime || '',
    classFormat: classFormat || '',
    classInstructor: classInstructor || ''
  };
  return cachedSettings;
}

// --- 2. MENU TIỆN ÍCH TRÊN SHEET ---
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Master T Tool')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('� Gửi lời mời tham gia')
      .addItem('📤 Gửi lời mời (từ InviteList)', 'sendInvitationEmails')
      .addItem('📊 Xem tiến trình gửi', 'viewInvitationProgress'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📧 Nhắc nhớ lớp học')
      .addItem('⚡ Gửi ngay', 'sendClassRemindersNow')
      .addItem('⏰ Hẹn giờ gửi (trước 2 giờ)', 'scheduleClassReminders'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🧪 Test gửi email')
      .addItem('✅ Test email xác nhận đăng ký', 'testEmailConfirmation')
      .addItem('⚠️ Test email rate limit', 'testEmailRateLimit')
      .addItem('📝 Test email nhắc nhớ lớp học', 'testEmailReminder')
      .addItem('📨 Test email lời mời', 'testInvitationEmail'))
    .addToUi();
}

// --- 3. CORE 1: XỬ LÝ KHI CÓ NGƯỜI ĐĂNG KÝ (Real-time) ---
function onFormSubmit(e) {
  try {
    if (!e || !e.values) return;

    // Lấy dữ liệu thô
    const rawName = e.values[CONFIG.COL_INDEX.NAME]; 
    const emailRaw   = e.values[CONFIG.COL_INDEX.EMAIL];
    const rawPhone = e.values[CONFIG.COL_INDEX.PHONE];

    // Chuẩn hóa email và kiểm tra hợp lệ
    const email = (emailRaw || '').toString().trim().toLowerCase();
    if (!isValidEmail(email)) {
      console.error(`Email không hợp lệ, bỏ qua: '${emailRaw}'`);
      return;
    }
    
    // Xử lý Logic Data Cleaning
    const cleanName = standardizeName(rawName);
    
    // Lấy 2 số cuối số điện thoại (Mặc định '00' nếu lỗi)
    let phoneSuffix = "00";
    if (rawPhone) {
      const strPhone = rawPhone.toString().trim();
      if (strPhone.length >= 2) phoneSuffix = strPhone.slice(-2);
    }

    // Format tên hiển thị Zoom: "25 - Hồ Văn Trọng"
    const zoomFirstName = phoneSuffix;
    const zoomLastName  = `- ${cleanName}`;

    // Gọi API Zoom
    const joinUrl = registerUserToZoom(email, zoomFirstName, zoomLastName);

    // Kiểm tra nếu gặp lỗi rate limit
    if (joinUrl && joinUrl.error === 'RATE_LIMIT') {
      sendRateLimitEmail(email, cleanName);
      return;
    }

    // Gửi Email
    if (joinUrl) {
      sendEmailWithUniqueLink(email, cleanName, joinUrl, zoomFirstName);
    }

  } catch (err) {
    console.error("Lỗi onFormSubmit: " + err.toString());
  }
}

// --- 4. CORE 2: GỬI EMAIL NHẮC NHỚ LỚP HỌC ---

// Helper: Parse thời gian từ CLASS_TIME
function parseClassStartTime(classTimeString) {
  // Format: "20:30 - 22:00 | Thứ Bảy, ngày 31/01/2026"
  try {
    const parts = classTimeString.split('|');
    if (parts.length < 2) return null;
    
    const timePart = parts[0].trim().split('-')[0].trim(); // "20:30"
    const datePart = parts[1].trim(); // "Thứ Bảy, ngày 31/01/2026"
    
    // Extract date: "ngày 31/01/2026"
    const dateMatch = datePart.match(/ngày\s+(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (!dateMatch) return null;
    
    const day = parseInt(dateMatch[1]);
    const month = parseInt(dateMatch[2]) - 1; // Month is 0-indexed
    const year = parseInt(dateMatch[3]);
    
    // Extract time: "20:30"
    const timeMatch = timePart.match(/(\d{1,2}):(\d{2})/);
    if (!timeMatch) return null;
    
    const hour = parseInt(timeMatch[1]);
    const minute = parseInt(timeMatch[2]);
    
    return new Date(year, month, day, hour, minute, 0);
  } catch (err) {
    console.error('Error parsing class time: ' + err.toString());
    return null;
  }
}

// Helper: Tính thời gian còn lại
function calculateTimeRemaining(startTime) {
  const now = new Date();
  const diff = startTime - now; // milliseconds
  
  if (diff < 0) return 'đã bắt đầu';
  
  const minutes = Math.floor(diff / (1000 * 60));
  const hours = Math.floor(minutes / 60);
  const days = Math.floor(hours / 24);
  
  if (days > 0) {
    return `trong ${days} ngày nữa`;
  } else if (hours > 0) {
    return `trong vài giờ nữa`;
  } else if (minutes > 10) {
    return `trong vài phút nữa`;
  } else {
    return 'ngay bây giờ';
  }
}

// Gửi email nhắc nhớ ngay
function sendClassRemindersNow() {
  const ui = SpreadsheetApp.getUi();
  const settings = getSettings();
  
  // Parse thời gian bắt đầu
  const startTime = parseClassStartTime(settings.classTime);
  if (!startTime) {
    ui.alert('❌ Không thể parse thời gian lớp học từ CLASS_TIME. Vui lòng kiểm tra format!');
    return;
  }
  
  // Tính thời gian còn lại
  const timeRemaining = calculateTimeRemaining(startTime);
  
  // Lấy danh sách email từ Sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    ui.alert("Chưa có dữ liệu học viên!");
    return;
  }
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, CONFIG.COL_INDEX.ZALO + 1);
  const data = dataRange.getValues();
  
  let sentCount = 0;
  let failedCount = 0;
  
  data.forEach(row => {
    const email = row[CONFIG.COL_INDEX.EMAIL];
    const name = row[CONFIG.COL_INDEX.NAME];
    
    if (!email || !isValidEmail(email)) {
      failedCount++;
      return;
    }
    
    try {
      sendClassReminderEmail(email, name, settings.className, settings.classTime, settings.classFormat, settings.classInstructor, timeRemaining);
      sentCount++;
    } catch (err) {
      console.error(`Lỗi gửi email cho ${email}: ${err.toString()}`);
      failedCount++;
    }
  });
  
  ui.alert(`✅ Gửi xong!\n✔️ Thành công: ${sentCount}\n❌ Lỗi: ${failedCount}\n⏰ Thời gian còn lại: ${timeRemaining}`);
}

// Hẹn giờ gửi email trước 2 giờ
function scheduleClassReminders() {
  const ui = SpreadsheetApp.getUi();
  const settings = getSettings();
  
  // Parse thời gian bắt đầu
  const startTime = parseClassStartTime(settings.classTime);
  if (!startTime) {
    ui.alert('❌ Không thể parse thời gian lớp học từ CLASS_TIME. Vui lòng kiểm tra format!');
    return;
  }
  
  // Tính thời gian gửi (trước 2 giờ)
  const sendTime = new Date(startTime.getTime() - 2 * 60 * 60 * 1000);
  const now = new Date();
  
  if (sendTime < now) {
    ui.alert('❌ Thời gian hẹn gửi đã qua! Lớp học sắp bắt đầu hoặc đã bắt đầu.\n\nVui lòng dùng "Gửi ngay" thay thế.');
    return;
  }
  
  // Xóa trigger cũ (nếu có)
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendScheduledClassReminders') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Tạo trigger mới
  ScriptApp.newTrigger('sendScheduledClassReminders')
    .timeBased()
    .at(sendTime)
    .create();
  
  const sendTimeStr = Utilities.formatDate(sendTime, "GMT+7", "HH:mm, dd/MM/yyyy");
  ui.alert(`✅ Đã hẹn giờ gửi email!\n\n⏰ Thời gian gửi: ${sendTimeStr}\n📧 Email sẽ được gửi tự động đến tất cả học viên.`);
}

// Hàm được trigger gọi
function sendScheduledClassReminders() {
  const settings = getSettings();
  const startTime = parseClassStartTime(settings.classTime);
  const timeRemaining = calculateTimeRemaining(startTime);
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return;
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, CONFIG.COL_INDEX.ZALO + 1);
  const data = dataRange.getValues();
  
  data.forEach(row => {
    const email = row[CONFIG.COL_INDEX.EMAIL];
    const name = row[CONFIG.COL_INDEX.NAME];
    
    if (!email || !isValidEmail(email)) return;
    
    try {
      sendClassReminderEmail(email, name, settings.className, settings.classTime, settings.classFormat, settings.classInstructor, timeRemaining);
    } catch (err) {
      console.error(`Lỗi gửi email cho ${email}: ${err.toString()}`);
    }
  });
}

// Legacy function (kept for backward compatibility)
function sendClassReminders() {
  const ui = SpreadsheetApp.getUi();
  
  // Hiển thị dialog nhập thông tin lớp học
  const response = ui.prompt(
    'GỬI NHẮC NHỚ LỚP HỌC',
    'Hãy nhập thông tin dưới đây (format: className|classTime|format|instructor)\n\n' +
    'Ví dụ:\n' +
    'Bí Mật Đằng Sau Một Lớp Học|20:30 - 22:00 | Thứ Bảy, ngày 31/01/2026|Trực tuyến qua Zoom|Hồ Văn Trọng',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.CANCEL) return;
  
  const input = response.getResponseText().trim();
  if (!input) {
    ui.alert('Vui lòng nhập thông tin!');
    return;
  }
  
  // Parse input
  const parts = input.split('|').map(s => s.trim());
  if (parts.length < 4) {
    ui.alert('Sai format! Cần 4 phần tách bằng |');
    return;
  }
  
  const [className, classTime, format, instructor] = parts;
  
  // Lấy danh sách email từ Sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    ui.alert("Chưa có dữ liệu học viên!");
    return;
  }
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, CONFIG.COL_INDEX.ZALO + 1);
  const data = dataRange.getValues();
  
  let sentCount = 0;
  let failedCount = 0;
  
  data.forEach(row => {
    const email = row[CONFIG.COL_INDEX.EMAIL];
    const name = row[CONFIG.COL_INDEX.NAME];
    
    if (!email || !isValidEmail(email)) {
      failedCount++;
      return;
    }
    
    try {
      sendClassReminderEmail(email, name, className, classTime, format, instructor);
      sentCount++;
    } catch (err) {
      console.error(`Lỗi gửi email cho ${email}: ${err.toString()}`);
      failedCount++;
    }
  });
  
  ui.alert(`✅ Gửi xong!\n✔️ Thành công: ${sentCount}\n❌ Lỗi: ${failedCount}`);
}

// --- 4.5. TEST GỬI EMAIL ---
function testEmailConfirmation() {
  const ui = SpreadsheetApp.getUi();
  
  // Nhập email nhận test
  const emailResponse = ui.prompt(
    'TEST: Email xác nhận đăng ký',
    'Nhập địa chỉ email nhận email test:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (emailResponse.getSelectedButton() === ui.Button.CANCEL) return;
  
  const testEmail = emailResponse.getResponseText().trim();
  if (!isValidEmail(testEmail)) {
    ui.alert('Email không hợp lệ!');
    return;
  }
  
  try {
    const testName = 'Test User';
    const testZoomNumber = '25';
    const testLink = 'https://zoom.us/j/123456789';
    sendEmailWithUniqueLink(testEmail, testName, testLink, testZoomNumber);
    ui.alert(`✅ Đã gửi email xác nhận đăng ký đến ${testEmail}`);
  } catch (err) {
    ui.alert(`❌ Lỗi gửi email: ${err.toString()}`);
    console.error(`Lỗi test email: ${err.toString()}`);
  }
}

function testEmailRateLimit() {
  const ui = SpreadsheetApp.getUi();
  
  // Nhập email nhận test
  const emailResponse = ui.prompt(
    'TEST: Email rate limit',
    'Nhập địa chỉ email nhận email test:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (emailResponse.getSelectedButton() === ui.Button.CANCEL) return;
  
  const testEmail = emailResponse.getResponseText().trim();
  if (!isValidEmail(testEmail)) {
    ui.alert('Email không hợp lệ!');
    return;
  }
  
  try {
    const testName = 'Test User';
    sendRateLimitEmail(testEmail, testName);
    ui.alert(`✅ Đã gửi email rate limit đến ${testEmail}`);
  } catch (err) {
    ui.alert(`❌ Lỗi gửi email: ${err.toString()}`);
    console.error(`Lỗi test email: ${err.toString()}`);
  }
}

function testEmailReminder() {
  const ui = SpreadsheetApp.getUi();
  
  // Nhập email nhận test
  const emailResponse = ui.prompt(
    'TEST: Email nhắc nhớ lớp học',
    'Nhập địa chỉ email nhận email test:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (emailResponse.getSelectedButton() === ui.Button.CANCEL) return;
  
  const testEmail = emailResponse.getResponseText().trim();
  if (!isValidEmail(testEmail)) {
    ui.alert('Email không hợp lệ!');
    return;
  }
  
  try {
    const settings = getSettings();
    const testName = 'Test User';
    
    // Lấy thông tin lớp học từ Script Properties
    const className = settings.className;
    const classTime = settings.classTime;
    const classFormat = settings.classFormat;
    const classInstructor = settings.classInstructor;
    
    if (!className || !classTime || !classFormat || !classInstructor) {
      ui.alert('⚠️ Chưa cấu hình thông tin lớp học trong Script Properties.\n\n' +
        'Vui lòng thiết lập:\n' +
        '- CLASS_NAME\n' +
        '- CLASS_TIME\n' +
        '- CLASS_FORMAT\n' +
        '- CLASS_INSTRUCTOR');
      return;
    }
    
    // Tính thời gian còn lại
    const startTime = parseClassStartTime(classTime);
    const timeRemaining = startTime ? calculateTimeRemaining(startTime) : 'trong vài giờ';
    
    sendClassReminderEmail(testEmail, testName, className, classTime, classFormat, classInstructor, timeRemaining);
    ui.alert(`✅ Đã gửi email nhắc nhớ lớp học đến ${testEmail}\n⏰ Thời gian còn lại: ${timeRemaining}`);
  } catch (err) {
    ui.alert(`❌ Lỗi gửi email: ${err.toString()}`);
    console.error(`Lỗi test email: ${err.toString()}`);
  }
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
  
  // Kiểm tra lỗi rate limit
  if (status === 429 || (json.code === 4300 && json.message && json.message.includes("exceeded the daily rate limit"))) {
    console.warn(`Rate limit exceeded for email: ${email}`);
    return { error: 'RATE_LIMIT', email: email, firstName: firstName, lastName: lastName };
  }
  
  if (json.join_url) return json.join_url;

  console.error(`Zoom register failed (status ${status}): ${body.slice(0, 400)}`);
  return null; // Trả về null nếu không có join_url
}

// Helper: API Lấy Report (Vét cạn các trang)
// REMOVED - Hàm này chỉ dùng cho syncAttendance() đã bị xóa

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
function sendEmailWithUniqueLink(email, name, link, zoomNumber) {
  const subject = '[Vé tham dự] CHUYÊN ĐỀ: QUY TRÌNH & CÔNG NGHỆ VẬN HÀNH LỚP HỌC ONLINE';
  const template = HtmlService.createTemplateFromFile('EmailTemplate');
  template.name = name;
  template.link = link;
  template.zoomNumber = zoomNumber || '00'; // Default '00' nếu không có
  const htmlBody = template.evaluate().getContent();
  const plainBody =
    `Chào ${name},\n` +
    `Bạn đã đăng ký chuyên đề "Quy trình & Công nghệ vận hành lớp học online".\n` +
    `Tên Zoom của bạn: ${zoomNumber} - ${name}\n` +
    `Mã số ${zoomNumber} sẽ dùng để quay số trung thưởng trong chương trình.\n` +
    `Link Zoom dành riêng cho bạn: ${link}\n` +
    `Nếu nút trong email không bấm được, hãy dán link này vào trình duyệt.\n` +
    `Hẹn gặp bạn trong lớp!`;
  GmailApp.sendEmail(email, subject, "", {
    htmlBody,
    plainBody,
    from: 'trong@hovantrong.com',
    name: 'Hồ Văn Trọng',
    replyTo: 'trong@hovantrong.com'
  });
}

// Helper: Gửi Email nhắc nhớ lớp học sắp diễn ra
function sendClassReminderEmail(email, name, className, classTime, format, instructor, timeRemaining) {
  const subject = `Nhắc nhớ: Buổi chia sẻ CHUYÊN ĐỀ: QUY TRÌNH & CÔNG NGHỆ VẬN HÀNH LỚP HỌC ONLINE sắp diễn ra`;
  
  // Nếu không truyền timeRemaining, tính mặc định
  if (!timeRemaining) {
    timeRemaining = 'trong vài giờ';
  }
  
  const template = HtmlService.createTemplateFromFile('ClassReminderTemplate');
  template.name = name;
  template.className = className;
  template.classTime = classTime;
  template.format = format;
  template.instructor = instructor;
  template.timeRemaining = timeRemaining;
  
  const htmlBody = template.evaluate().getContent();
  const plainBody =
    `Chào ${name},\n\n` +
    `Lớp "${className}" sắp diễn ra rồi!\n\n` +
    `Thông tin lớp học:\n` +
    `Thời gian: ${classTime}\n` +
    `Hình thức: ${format}\n` +
    `Người chia sẻ: ${instructor}\n\n` +
    `CÁCH VÀO LỚP:\n` +
    `Hãy kiểm tra lại email "Xác nhận đăng ký thành công" mà bạn nhận được lúc đăng ký.\n` +
    `Email đó chứa link Zoom cá nhân của bạn.\n` +
    `Click nút "VÀO LỚP NGAY" hoặc dán link vào trình duyệt.\n\n` +
    `Chuẩn bị tham dự ngay!\n\n` +
    `Trân trọng,\n` +
    `Hồ Văn Trọng`;
    
  GmailApp.sendEmail(email, subject, plainBody, {
    htmlBody,
    from: 'trong@hovantrong.com',
    name: 'Hồ Văn Trọng',
    replyTo: 'trong@hovantrong.com'
  });
}

// Helper: Gửi Email thông báo khi vượt quá giới hạn rate limit
function sendRateLimitEmail(email, name) {
  const subject = '[⚠️ Thông báo] Đạt giới hạn đăng ký - Vui lòng sử dụng email khác';
  const template = HtmlService.createTemplateFromFile('RateLimitEmailTemplate');
  template.name = name;
  template.formLink = 'https://forms.gle/vL8A2nwYpFneRdeW9';
  const htmlBody = template.evaluate().getContent();
  
  const plainBody =
    `Chào ${name},\n\n` +
    `Email này đã đạt giới hạn 3 lần đăng ký trong 24 giờ (quy định của Zoom API).\n\n` +
    `GIẢI PHÁP:\n` +
    `1. Nhanh nhất: Dùng email khác để đăng ký lại\n` +
    `2. Chờ: Thử lại ngày mai sau 24h\n` +
    `3. Liên hệ: 0936 099 625 (Mr. Trọng)\n\n` +
    `Chi tiết xem trong email HTML.\n\n` +
    `Trân trọng,\n` +
    `Hồ Văn Trọng`;
    
  GmailApp.sendEmail(email, subject, "", {
    htmlBody,
    plainBody,
    from: 'trong@hovantrong.com',
    name: 'Hồ Văn Trọng',
    replyTo: 'trong@hovantrong.com'
  });
}

// --- 6. GỬI EMAIL LỜI MỜI THAM GIA (Batch Processing) ---

// Cấu hình cho Invitation
const INVITATION_CONFIG = {
  SHEET_NAME: 'InviteList',        // Tên sheet chứa danh sách mời
  BATCH_SIZE: 20,                   // Số email gửi mỗi batch (tránh timeout)
  DELAY_BETWEEN_EMAILS: 500,        // Độ trễ giữa các email (ms)
  COL_EMAIL: 0,                     // Cột A: Email
  COL_NAME: 1,                      // Cột B: Tên
  COL_STATUS: 2,                    // Cột C: Trạng thái gửi
  COL_SENT_TIME: 3,                 // Cột D: Thời gian gửi
  COL_ERROR: 4                      // Cột E: Lỗi (nếu có)
};

// Hàm chính: Gửi email lời mời từ sheet InviteList
function sendInvitationEmails() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Tìm sheet InviteList
  let sheet = ss.getSheetByName(INVITATION_CONFIG.SHEET_NAME);
  
  if (!sheet) {
    // Tạo sheet mới nếu chưa có
    const createSheet = ui.alert(
      '⚠️ Chưa có sheet "InviteList"',
      'Bạn có muốn tạo sheet "InviteList" mới không?\n\n' +
      'Sheet sẽ có các cột:\n' +
      'A: Email\n' +
      'B: Tên\n' +
      'C: Trạng thái\n' +
      'D: Thời gian gửi\n' +
      'E: Lỗi',
      ui.ButtonSet.YES_NO
    );
    
    if (createSheet === ui.Button.YES) {
      sheet = createInviteListSheet(ss);
      ui.alert('✅ Đã tạo sheet "InviteList"!\n\nVui lòng điền danh sách email và tên, sau đó chạy lại.');
      return;
    } else {
      return;
    }
  }
  
  // Đếm số email chưa gửi
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('❌ Sheet "InviteList" chưa có dữ liệu!\n\nVui lòng điền danh sách email từ dòng 2.');
    return;
  }
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 5);
  const data = dataRange.getValues();
  
  // Đếm số email chưa gửi và đã gửi
  let pendingCount = 0;
  let sentCount = 0;
  
  data.forEach(row => {
    const email = row[INVITATION_CONFIG.COL_EMAIL];
    const status = row[INVITATION_CONFIG.COL_STATUS];
    
    if (email && isValidEmail(email.toString().trim())) {
      if (status === 'Đã gửi' || status === 'SENT') {
        sentCount++;
      } else {
        pendingCount++;
      }
    }
  });
  
  if (pendingCount === 0) {
    ui.alert(`✅ Tất cả email đã được gửi!\n\nTổng số: ${sentCount} email`);
    return;
  }
  
  // Xác nhận trước khi gửi
  const confirm = ui.alert(
    '📨 Xác nhận gửi lời mời',
    `📊 Thống kê:\n` +
    `• Chưa gửi: ${pendingCount} email\n` +
    `• Đã gửi: ${sentCount} email\n\n` +
    `⏱️ Ước tính thời gian: ~${Math.ceil(pendingCount / INVITATION_CONFIG.BATCH_SIZE)} phút\n\n` +
    `Bạn có muốn bắt đầu gửi không?`,
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) {
    return;
  }
  
  // Bắt đầu gửi với progress tracking
  processInvitationBatch();
}

// Tạo sheet InviteList với header
function createInviteListSheet(ss) {
  const sheet = ss.insertSheet(INVITATION_CONFIG.SHEET_NAME);
  
  // Thiết lập header
  const headers = ['Email', 'Tên', 'Trạng thái', 'Thời gian gửi', 'Lỗi'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1e3a8a');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // Thiết lập độ rộng cột
  sheet.setColumnWidth(1, 250); // Email
  sheet.setColumnWidth(2, 200); // Tên
  sheet.setColumnWidth(3, 100); // Trạng thái
  sheet.setColumnWidth(4, 180); // Thời gian gửi
  sheet.setColumnWidth(5, 200); // Lỗi
  
  // Freeze header
  sheet.setFrozenRows(1);
  
  return sheet;
}

// Xử lý gửi email theo batch
function processInvitationBatch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INVITATION_CONFIG.SHEET_NAME);
  
  if (!sheet) {
    console.error('Không tìm thấy sheet InviteList');
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 5);
  const data = dataRange.getValues();
  
  let processedInBatch = 0;
  let totalSent = 0;
  let totalFailed = 0;
  let hasMoreToProcess = false;
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const email = row[INVITATION_CONFIG.COL_EMAIL];
    const name = row[INVITATION_CONFIG.COL_NAME] || 'Anh/Chị';
    const status = row[INVITATION_CONFIG.COL_STATUS];
    
    // Bỏ qua nếu đã gửi hoặc email không hợp lệ
    if (status === 'Đã gửi' || status === 'SENT') {
      totalSent++;
      continue;
    }
    
    if (!email || !isValidEmail(email.toString().trim())) {
      continue;
    }
    
    // Kiểm tra đã đạt batch size chưa
    if (processedInBatch >= INVITATION_CONFIG.BATCH_SIZE) {
      hasMoreToProcess = true;
      break;
    }
    
    // Gửi email
    const rowIndex = i + 2; // Row trong sheet (1-indexed, bắt đầu từ row 2)
    
    try {
      sendInvitationEmail(email.toString().trim(), standardizeName(name));
      
      // Cập nhật trạng thái thành công
      sheet.getRange(rowIndex, INVITATION_CONFIG.COL_STATUS + 1).setValue('Đã gửi');
      sheet.getRange(rowIndex, INVITATION_CONFIG.COL_SENT_TIME + 1).setValue(new Date());
      sheet.getRange(rowIndex, INVITATION_CONFIG.COL_ERROR + 1).setValue('');
      
      // Highlight màu xanh
      sheet.getRange(rowIndex, 1, 1, 5).setBackground('#d1fae5');
      
      processedInBatch++;
      totalSent++;
      
      // Delay giữa các email để tránh rate limit
      if (processedInBatch < INVITATION_CONFIG.BATCH_SIZE) {
        Utilities.sleep(INVITATION_CONFIG.DELAY_BETWEEN_EMAILS);
      }
      
    } catch (err) {
      // Cập nhật trạng thái lỗi
      sheet.getRange(rowIndex, INVITATION_CONFIG.COL_STATUS + 1).setValue('Lỗi');
      sheet.getRange(rowIndex, INVITATION_CONFIG.COL_ERROR + 1).setValue(err.toString().slice(0, 200));
      
      // Highlight màu đỏ
      sheet.getRange(rowIndex, 1, 1, 5).setBackground('#fee2e2');
      
      totalFailed++;
      console.error(`Lỗi gửi email cho ${email}: ${err.toString()}`);
    }
  }
  
  // Nếu còn email chưa gửi, tạo trigger để tiếp tục
  if (hasMoreToProcess) {
    // Xóa trigger cũ nếu có
    deleteTriggerByFunction('processInvitationBatch');
    
    // Tạo trigger mới sau 1 phút để tiếp tục gửi
    ScriptApp.newTrigger('processInvitationBatch')
      .timeBased()
      .after(60 * 1000) // 1 phút
      .create();
    
    console.log(`Batch completed: ${processedInBatch} emails sent. Scheduling next batch...`);
  } else {
    // Xóa trigger nếu đã gửi xong
    deleteTriggerByFunction('processInvitationBatch');
    
    // Gửi thông báo hoàn thành
    console.log(`All invitations sent! Total: ${totalSent} sent, ${totalFailed} failed.`);
  }
  
  // Lưu tiến trình vào Properties
  const props = PropertiesService.getScriptProperties();
  props.setProperty('INVITATION_LAST_UPDATE', new Date().toISOString());
  props.setProperty('INVITATION_TOTAL_SENT', totalSent.toString());
  props.setProperty('INVITATION_TOTAL_FAILED', totalFailed.toString());
}

// Helper: Xóa trigger theo tên function
function deleteTriggerByFunction(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// Xem tiến trình gửi lời mời
function viewInvitationProgress() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INVITATION_CONFIG.SHEET_NAME);
  
  if (!sheet) {
    ui.alert('❌ Chưa có sheet "InviteList"!');
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('❌ Sheet "InviteList" chưa có dữ liệu!');
    return;
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  
  let total = 0;
  let sent = 0;
  let pending = 0;
  let failed = 0;
  let invalid = 0;
  
  data.forEach(row => {
    const email = row[INVITATION_CONFIG.COL_EMAIL];
    const status = row[INVITATION_CONFIG.COL_STATUS];
    
    if (!email) return;
    
    total++;
    
    if (!isValidEmail(email.toString().trim())) {
      invalid++;
      return;
    }
    
    if (status === 'Đã gửi' || status === 'SENT') {
      sent++;
    } else if (status === 'Lỗi') {
      failed++;
    } else {
      pending++;
    }
  });
  
  // Kiểm tra có trigger đang chạy không
  const triggers = ScriptApp.getProjectTriggers();
  const isRunning = triggers.some(t => t.getHandlerFunction() === 'processInvitationBatch');
  
  const props = PropertiesService.getScriptProperties();
  const lastUpdate = props.getProperty('INVITATION_LAST_UPDATE') || 'Chưa có';
  
  ui.alert(
    '📊 Tiến trình gửi lời mời',
    `📧 Tổng số email: ${total}\n` +
    `✅ Đã gửi: ${sent}\n` +
    `⏳ Chưa gửi: ${pending}\n` +
    `❌ Lỗi: ${failed}\n` +
    `⚠️ Email không hợp lệ: ${invalid}\n\n` +
    `🔄 Trạng thái: ${isRunning ? 'Đang xử lý...' : 'Không có batch đang chạy'}\n` +
    `🕐 Cập nhật lần cuối: ${lastUpdate}`,
    ui.ButtonSet.OK
  );
}

// Helper: Gửi email lời mời
function sendInvitationEmail(email, name) {
  const subject = '[THƯ MỜI] Chuyên đề: BÍ MẬT VẬN HÀNH LỚP HỌC ONLINE - Tư duy & Công nghệ thực chiến';
  
  const template = HtmlService.createTemplateFromFile('InvitationEmailTemplate');
  template.name = name || 'Anh/Chị';
  
  const htmlBody = template.evaluate().getContent();
  
  const plainBody =
    `Xin chào ${name},\n\n` +
    `Bạn được mời tham dự buổi chia sẻ chuyên đề "BÍ MẬT VẬN HÀNH LỚP HỌC ONLINE - Tư duy & Công nghệ thực chiến".\n\n` +
    `🎓 MIỄN PHÍ THAM DỰ\n\n` +
    `📌 THÔNG TIN SỰ KIỆN:\n` +
    `• Thời gian: 20:30 - 22:00 | Thứ Bảy, ngày 31/01/2026\n` +
    `• Hình thức: Trực tuyến qua Zoom\n` +
    `• Người chia sẻ: Hồ Văn Trọng – Chuyên gia IT & Phát triển tâm thức\n\n` +
    `🎯 NỘI DUNG CHÍNH - 3 Trụ cột vận hành:\n` +
    `1. Tư duy hệ thống (System Thinking)\n` +
    `2. Công nghệ thực chiến (Tech Stack)\n` +
    `3. Kỹ năng vận hành & Quản trị\n\n` +
    `🎁 QUÀ TẶNG GIÁ TRỊ:\n` +
    `• 02 Giải: Tài khoản ChatGPT Plus (01 tháng)\n` +
    `• 02 Giải: Tài khoản Zoom Pro (03 tháng)\n` +
    `• 02 Giải: Tài khoản Google AI Pro (01 năm)\n\n` +
    `📝 ĐĂNG KÝ NGAY: https://forms.gle/4xKKxYh1REHArHGz6\n\n` +
    `📞 Liên hệ hỗ trợ: 0936 099 625 (Mr. Trọng)\n\n` +
    `Trân trọng,\n` +
    `Hồ Văn Trọng`;
    
  GmailApp.sendEmail(email, subject, plainBody, {
    htmlBody,
    from: 'trong@hovantrong.com',
    name: 'Hồ Văn Trọng',
    replyTo: 'trong@hovantrong.com'
  });
}

// Test gửi email lời mời
function testInvitationEmail() {
  const ui = SpreadsheetApp.getUi();
  
  const emailResponse = ui.prompt(
    'TEST: Email lời mời tham dự',
    'Nhập địa chỉ email nhận email test:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (emailResponse.getSelectedButton() === ui.Button.CANCEL) return;
  
  const testEmail = emailResponse.getResponseText().trim();
  if (!isValidEmail(testEmail)) {
    ui.alert('Email không hợp lệ!');
    return;
  }
  
  try {
    sendInvitationEmail(testEmail, 'Test User');
    ui.alert(`✅ Đã gửi email lời mời đến ${testEmail}`);
  } catch (err) {
    ui.alert(`❌ Lỗi gửi email: ${err.toString()}`);
    console.error(`Lỗi test email: ${err.toString()}`);
  }
}