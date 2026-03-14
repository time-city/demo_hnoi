# HƯỚNG DẪN SETUP HỆ THỐNG PHÂN TÍCH ĐIỂM (n8n + Dashboard HTML/JS)

Tài liệu này hướng dẫn từ máy mới hoàn toàn đến khi hệ thống chạy được: n8n hoạt động, cấu hình API model thành công, web dashboard gửi dữ liệu và nhận JSON phản hồi.

---

# 1. CÀI ĐẶT MÔI TRƯỜNG

## 1.1 Cài NodeJS

1. Truy cập https://nodejs.org  
2. Tải bản **LTS**  
3. Cài đặt bình thường (Next → Next → Finish)  
4. Kiểm tra:

Mở Command Prompt / Terminal và chạy:

node -v  
npm -v  

Nếu hiển thị version → OK.

---

# 2. CÀI ĐẶT VÀ CHẠY N8N

## 2.1 Cài n8n global

npm install -g n8n

## 2.2 Chạy n8n

n8n

Nếu thành công sẽ hiện:

Editor is now accessible via:  
http://localhost:5678

Mở trình duyệt và truy cập địa chỉ trên.

---

# 3. IMPORT WORKFLOW

1. Vào n8n  
2. Góc phải trên → Import  
3. Dán file JSON workflow  
4. Nhấn Import  
5. Nhấn Activate để kích hoạt

---

# 4. CẤU HÌNH WEBHOOK

Trong node Webhook:

- Method: GET + POST  
- Path: demo  
- Response Mode: Using "Respond to Webhook"  

Sau khi Activate, endpoint sẽ là:

http://localhost:5678/webhook/demo

Đây là URL mà web dashboard sẽ gọi.

---

# 5. THÊM API KEY VÀ BASE URL CHO MODEL AI

## 5.1 Tạo Credential

1. Menu trái → Credentials  
2. Create Credential  

Nếu dùng OpenAI chính thức → chọn **OpenAI API**  
Nếu dùng server riêng → chọn **OpenAI Compatible API**

## 5.2 Nhập thông tin

- API Key: sk-xxxxxx  
- Base URL: https://yourdomain.com/v1  (nếu dùng server riêng)

Lưu ý:  
- Base URL phải có /v1 ở cuối  
- Nếu dùng OpenAI chính thức thì không cần chỉnh Base URL  

Nhấn Save.

---

# 6. GẮN MODEL VÀO NODE CHAT

Trong workflow:

1. Mở node OpenAI Chat Model hoặc AI Agent  
2. Ở mục Credentials → chọn credential vừa tạo  
3. Ở mục Model → nhập tên model (ví dụ: gpt-4o-mini)  
4. Execute Node để test  

Nếu lỗi 401 → sai API key  
Nếu lỗi model not found → sai tên model  

---

# 7. CHẠY DASHBOARD HTML/CSS/JS

Giả sử bạn có thư mục:

dashboard/
  index.html
  style.css
  app.js

Vào thư mục đó:

cd dashboard

Chạy server tạm:

npx serve

Hoặc:

npx http-server

Sau đó mở:

http://localhost:3000

---

# 8. KẾT NỐI WEB VỚI N8N

Trong file JS, đảm bảo endpoint:

http://localhost:5678/webhook/demo

Ví dụ gọi POST:

fetch("http://localhost:5678/webhook/demo", {
  method: "POST",
  headers: {
    "Content-Type": "application/json"
  },
  body: JSON.stringify(data)
})

Nếu lỗi CORS, chạy n8n với:

Windows:

set N8N_CORS_ENABLED=true  
n8n  

Mac/Linux:

export N8N_CORS_ENABLED=true  
n8n  

---

# 9. LUỒNG HOẠT ĐỘNG HỆ THỐNG

Dashboard gửi JSON  
→ Webhook nhận  
→ Normalize Input  
→ Phân tích chuẩn (Sheet 1 + AI)  
→ Phân tích học sinh (Sheet 2, 3)  
→ Phân tích lớp (Sheet 4)  
→ Merge  
→ Wrap Response  
→ Respond to Webhook trả JSON  

---

# 10. LƯU Ý QUAN TRỌNG

- Không hardcode API key  
- Không public giao diện n8n editor ra Internet  
- Validate JSON trước khi xử lý  
- Khi >500 học sinh: chỉ dùng AI cho phân tích chuẩn, không cho từng học sinh  

---

# 11. TEST HOÀN CHỈNH

1. Chạy n8n  
2. Activate workflow  
3. Chạy dashboard  
4. Gửi dữ liệu mẫu  
5. Kiểm tra response có:

ok: true  

Nếu có → hệ thống hoạt động bình thường.

---

Sau khi hoàn thành toàn bộ bước trên, hệ thống đã sẵn sàng để demo hoặc triển khai production.