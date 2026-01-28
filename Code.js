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

    // Kiểm tra nếu gặp lỗi rate limit
    if (joinUrl && joinUrl.error === 'RATE_LIMIT') {
      sendRateLimitEmail(email, cleanName);
      return;
    }

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

// Helper: Gửi Email thông báo khi vượt quá giới hạn rate limit
function sendRateLimitEmail(email, name) {
  const subject = '[⚠️ Thông báo] Đạt giới hạn đăng ký - Vui lòng sử dụng email khác';
  
  const htmlBody = `
    <!DOCTYPE html>
    <html lang="vi">
    <head>
      <meta charset="UTF-8" />
      <meta name="viewport" content="width=device-width, initial-scale=1" />
      <title>Thông báo giới hạn đăng ký</title>
    </head>
    <body style="margin:0; padding:0; background-color:#f5f7fa; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background-color:#f5f7fa; padding:30px 15px;">
        <tr>
          <td align="center">
            <table role="presentation" width="600" cellspacing="0" cellpadding="0" border="0" style="max-width:600px; background:#ffffff; border-radius:16px; box-shadow:0 4px 24px rgba(0,0,0,0.08); overflow:hidden;">
              
              <!-- Header -->
              <tr>
                <td style="background: linear-gradient(135deg, #f97316 0%, #ea580c 100%); padding:35px 40px; text-align:center;">
                  <p style="margin:0 0 8px 0; font-size:13px; color:rgba(255,255,255,0.85); text-transform:uppercase; letter-spacing:1.5px;">⚠️ THÔNG BÁO QUAN TRỌNG</p>
                  <h1 style="margin:0; font-size:23px; font-weight:700; color:#ffffff; line-height:1.3;">Đạt Giới Hạn Đăng Ký<br/>Hôm Nay</h1>
                </td>
              </tr>
              
              <!-- Content -->
              <tr>
                <td style="padding:35px 40px;">
                  <p style="margin:0 0 20px 0; font-size:16px; color:#2d3748; line-height:1.7;">
                    Chào <strong style="color:#f97316;">${name}</strong>,
                  </p>
                  
                  <p style="margin:0 0 20px 0; font-size:15px; color:#4a5568; line-height:1.8;">
                    Hệ thống đã nhận được yêu cầu đăng ký của bạn, nhưng <strong>email này đã đạt giới hạn 3 lần đăng ký trong 24 giờ quy định của Zoom API</strong>.
                  </p>
                  
                  <p style="margin:0 0 20px 0; font-size:15px; color:#4a5568; line-height:1.8;">
                    <strong>⏰ Lý do:</strong> Zoom Pro account giới hạn mỗi email được đăng ký tối đa 3 lần/ngày để bảo vệ hệ thống. Giới hạn này sẽ reset vào 00:00 ngày hôm sau (GMT+7).
                  </p>
                  
                  <!-- Giải pháp -->
                  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background:#ecfdf5; border-radius:12px; border-left:4px solid #16a34a; margin:25px 0;">
                    <tr>
                      <td style="padding:20px 25px;">
                        <p style="margin:0 0 15px 0; font-size:15px; font-weight:700; color:#15803d;">✅ Giải pháp</p>
                        <p style="margin:0 0 10px 0; font-size:14px; color:#2d3748; line-height:1.7;">
                          <strong>Cách 1 (Nhanh nhất):</strong> Dùng email khác để đăng ký lại. Chúng tôi sẽ tạo 1 link Zoom duy nhất cho email mới của bạn.
                        </p>
                        <p style="margin:0 0 10px 0; font-size:14px; color:#2d3748; line-height:1.7;">
                          <strong>Cách 2 (Chờ):</strong> Nếu muốn dùng email này, vui lòng thử lại vào ngày mai (sau 24h).
                        </p>
                        <p style="margin:0; font-size:14px; color:#2d3748; line-height:1.7;">
                          <strong>Cách 3 (Liên hệ):</strong> Gọi/Zalo/Telegram 0936 099 625 (Mr. Trọng) để được hỗ trợ thêm.
                        </p>
                      </td>
                    </tr>
                  </table>
                  
                  <p style="margin:25px 0 15px 0; font-size:15px; color:#2d3748; line-height:1.8;">
                    <strong>Các bước tiếp theo:</strong>
                  </p>
                  
                  <ol style="margin:0 0 25px 0; padding-left:20px; font-size:14px; color:#4a5568; line-height:1.8;">
                    <li style="margin-bottom:10px;">
                      <strong>Nếu chọn email mới:</strong> Vui lòng dùng email khác mà bạn có quyền truy cập và gửi form đăng ký lại.
                    </li>
                    <li style="margin-bottom:10px;">
                      <strong>Xác nhận:</strong> Bạn sẽ nhận email xác nhận link Zoom trong vòng vài giây.
                    </li>
                    <li>
                      <strong>Tham gia:</strong> Sử dụng link đó để tham gia buổi thảo luận vào ngày 31/01/2026.
                    </li>
                  </ol>
                  
                  <!-- CTA Button -->
                  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="margin:30px 0;">
                    <tr>
                      <td align="center">
                        <a href="https://docs.google.com/forms/d/YOUR_FORM_ID/viewform" style="display:inline-block; background: linear-gradient(135deg, #16a34a 0%, #15803d 100%); color:#ffffff; text-decoration:none; padding:14px 40px; border-radius:50px; font-weight:700; font-size:15px; box-shadow:0 4px 15px rgba(22,163,74,0.4);">
                          QUAY LẠI FORM ĐĂNG KÝ
                        </a>
                      </td>
                    </tr>
                  </table>
                  
                  <p style="margin:30px 0 0 0; font-size:13px; color:#718096; line-height:1.6; border-top:1px solid #e2e8f0; padding-top:20px;">
                    <strong>💡 Lưu ý:</strong> Đây là giới hạn của hệ thống Zoom API (Zoom Pro Account). Nếu muốn tăng giới hạn lên 10 lần/ngày hoặc cao hơn, bạn cần nâng cấp lên Zoom Business Account.
                  </p>
                </td>
              </tr>
              
              <!-- Footer -->
              <tr>
                <td style="background:#f7fafc; padding:25px 40px; border-top:1px solid #e2e8f0;">
                  <p style="margin:0; font-size:14px; color:#718096; line-height:1.6;">
                    Trân trọng,<br/>
                    <strong style="color:#4a5568;">Hồ Văn Trọng</strong><br/>
                    <span style="font-size:12px; color:#a0aec0;">Hotline: 0936 099 625</span>
                  </p>
                </td>
              </tr>
              
            </table>
          </td>
        </tr>
      </table>
    </body>
    </html>
  `;
  
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