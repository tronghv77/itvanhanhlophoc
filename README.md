# GỬI LINK ZOOM CÁ NHÂN HÓA SAU KHI ĐIỀN FORM

**GỬI LINK ZOOM CÁ NHÂN HÓA SAU KHI ĐIỀN FORM** là một công cụ tự động hóa mạnh mẽ được xây dựng trên nền tảng Google Apps Script, giúp quản lý quy trình đăng ký và điểm danh cho các lớp học hoặc hội thảo trực tuyến qua Zoom.

Dự án được phát triển bởi **Master T & Trọng**.

## 🚀 Tính năng chính

### 1. Tự động đăng ký Zoom (Real-time)
- **Kích hoạt:** Khi học viên điền Google Form.
- **Xử lý:**
  - Tự động chuẩn hóa tên học viên (Title Case).
  - Đăng ký học viên vào Zoom Webinar/Meeting thông qua API.
  - Tạo liên kết tham gia duy nhất (Unique Join URL) cho từng người.
- **Gửi Email:** Tự động gửi email xác nhận chứa link Zoom riêng biệt cho học viên (sử dụng template HTML chuyên nghiệp).

### 2. Đồng bộ điểm danh (Post-Meeting)
- **Kích hoạt:** Thông qua menu tùy chỉnh trên Google Sheet (`Master T Tool` > `🔄 Đồng bộ điểm danh Zoom`).
- **Xử lý:**
  - Kết nối API Zoom để lấy báo cáo người tham dự (Hỗ trợ phân trang cho lớp đông > 500 người).
  - Tự động cộng dồn thời gian tham gia nếu học viên ra vào nhiều lần.
  - Đối chiếu với danh sách đăng ký trong Sheet.
- **Kết quả:** Cập nhật trạng thái "Đã tham gia" hoặc "Vắng", thời gian tham gia (phút), và giờ vào lớp vào các cột tương ứng trên Sheet.

## 🛠 Yêu cầu hệ thống

1. **Google Workspace:**
   - Google Sheet (Lưu trữ dữ liệu).
   - Google Form (Thu thập đăng ký).
   - Gmail (Gửi thư xác nhận).
2. **Zoom Account:**
   - Tài khoản Zoom Pro/Business trở lên.
   - Tạo ứng dụng **Server-to-Server OAuth** trên [Zoom App Marketplace](https://marketplace.zoom.us/) để lấy Credentials.

## ⚙️ Hướng dẫn cài đặt

### 1. Cấu hình Script Properties
Vào trình soạn thảo Apps Script, chọn **Project Settings** (biểu tượng bánh răng) > **Script Properties** và thêm các key sau:

| Property | Mô tả |
|----------|-------|
| `ZOOM_ACCOUNT_ID` | Account ID từ Zoom App |
| `ZOOM_CLIENT_ID` | Client ID từ Zoom App |
| `ZOOM_CLIENT_SECRET` | Client Secret từ Zoom App |
| `MEETING_ID` | ID của cuộc họp/webinar cần quản lý |

### 2. Cấu trúc Google Sheet
Dữ liệu trong Sheet cần tuân thủ thứ tự cột (Index bắt đầu từ 0):

- **Cột B (Index 1):** Email
- **Cột C (Index 2):** Họ và tên
- **Cột D (Index 3):** Số Zalo
- **Cột H, I, J (Index 7+):** Nơi script sẽ ghi kết quả điểm danh (Status, Duration, Time In).

*Lưu ý: Có thể thay đổi cấu hình này trong biến `CONFIG` tại file `Code.js`.*

### 3. Cài đặt Trigger (Kích hoạt tự động)
Để tính năng đăng ký tự động hoạt động, cần cài đặt Installable Trigger:
1. Vào mục **Triggers** (biểu tượng đồng hồ).
2. Chọn **Add Trigger**.
3. Cấu hình:
   - Function: `onFormSubmit`
   - Event source: `From spreadsheet`
   - Event type: `On form submit`

## 📖 Hướng dẫn sử dụng

1. **Chuẩn bị:** Đảm bảo `MEETING_ID` trong Script Properties là chính xác cho buổi học sắp tới.
2. **Tuyển sinh:** Gửi Google Form cho học viên. Hệ thống sẽ tự động đăng ký và gửi mail.
3. **Kết thúc lớp học:**
   - Mở Google Sheet.
   - Chọn menu **Master T Tool** trên thanh công cụ.
   - Chọn **🔄 Đồng bộ điểm danh Zoom**.
   - Đợi script chạy và xem kết quả cập nhật trực tiếp trên Sheet.

## 📁 Cấu trúc dự án
- `Code.js`: Chứa toàn bộ logic xử lý (API Zoom, xử lý dữ liệu Sheet, gửi mail).
- `EmailTemplate.html`: Mẫu email HTML gửi cho học viên.
- `appsscript.json`: Cấu hình manifest của dự án Apps Script.
