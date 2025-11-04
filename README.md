🧩 Dự án: Selenium Excel Login Automation (Python Demo)
📖 Giới thiệu
Dự án minh họa cách thực hiện kiểm thử tự động dựa trên dữ liệu (Data-Driven Testing) bằng Selenium WebDriver (Python).
Script giúp tự động đăng nhập vào trang web demo, đọc thông tin tài khoản từ file Excel, và ghi lại kết quả đăng nhập (Pass/Fail) sau mỗi lần chạy.

⚙️ Tính năng chính
🔹 Tự động kiểm thử chức năng đăng nhập với Selenium WebDriver
🔹 Đọc danh sách tài khoản (username, password) từ file Excel bằng thư viện openpyxl
🔹 Ghi kết quả đăng nhập trở lại file Excel (Pass/Fail)
🔹 Xử lý ngoại lệ, bỏ qua các tài khoản đã đăng nhập thành công
🔹 Có thể chạy trên Chrome hoặc Edge mà không cần cài thủ công driver

🧠 Công nghệ sử dụng
Ngôn ngữ: Python
Thư viện: Selenium, OpenPyXL
Công cụ: Excel, VS Code

📁 Cấu trúc thư mục
Selenium Excel Login Automation/
│
├── data/
│   ├── login_data.xlsx          # File Excel chứa tài khoản kiểm thử
│   └── login_result.xlsx        # File Excel kết quả
├── src/
│   └── selenium_excel_demo.py      # File chính chạy tự động đăng nhập
│
└── README.md                    # Tài liệu mô tả dự án

✅ Quy trình hoạt động
Đọc danh sách tài khoản từ file Excel
Mở trình duyệt và truy cập trang đăng nhập
Nhập username/password → nhấn nút đăng nhập
Ghi lại kết quả kiểm thử vào Excel
Tiếp tục chạy cho đến khi hoàn tất danh sách tài khoản

💡 Mục đích học tập
Dự án được thực hiện nhằm luyện tập các kỹ năng:
Kiểm thử tự động bằng Selenium
Áp dụng mô hình Data-Driven Testing
Kết hợp Python với Excel trong kiểm thử
Xử lý ngoại lệ và ghi log kết quả

👤 Tác giả
Ngoc – Thực tập sinh Tester
Hiện đang học và phát triển kỹ năng Automation Testing với Python & Selenium.
