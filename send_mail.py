import os
import json
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

# Lấy thông tin từ biến bảo mật GitHub Secrets
SENDER_EMAIL = os.environ.get("SENDER_EMAIL")
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD")
RECEIVER_EMAIL = SENDER_EMAIL # Gửi cho chính mình

PROGRESS_FILE = "quiz_progress.json"

def send_daily_report():
    if not SENDER_EMAIL or not SENDER_PASSWORD:
        print("Chưa cấu hình tài khoản gửi email trong GitHub Secrets!")
        return

    # Đọc tiến độ học tập từ file JSON nếu có
    progress_data = {}
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                progress_data = json.load(f)
        except Exception as e:
            print(f"Lỗi đọc file progress: {e}")

    # Tạo nội dung email báo cáo
    subject = "[Tự động] Báo cáo tiến độ ôn tập An Toàn Điện"
    
    html_content = f"""
    <html>
      <body style="font-family: Arial, sans-serif; color: #333;">
        <h2 style="color: #2b6cb0;">📊 Báo cáo tiến độ ôn tập An Toàn Điện mỗi ngày</h2>
        <p>Chào Đức Anh,</p>
        <p>Đây là hệ thống tự động nhắc nhở tiến độ ôn tập thi an toàn điện của bạn lúc 7:00 sáng.</p>
        <hr style="border: none; border-top: 1px solid #eee;" />
        <h3>📈 Dữ liệu tiến độ hiện tại:</h3>
        <pre style="background: #f7fafc; padding: 10px; border-radius: 5px;">{json.dumps(progress_data, ensure_ascii=False, indent=2)}</pre>
        <p style="margin-top: 20px; font-size: 12px; color: #718096;">Hệ thống tự động kích hoạt qua GitHub Actions.</p>
      </body>
    </html>
    """

    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECEIVER_EMAIL
    msg['Subject'] = Header(subject, 'utf-8')
    msg.attach(MIMEText(html_content, 'html', 'utf-8'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, msg.as_string())
        server.quit()
        print("Đã gửi email báo cáo thành công!")
    except Exception as e:
        print(f"Lỗi gửi email: {e}")

if __name__ == "__main__":
    send_daily_report()
