import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def check_new_companies():
    companies = [
        {"name": "Công ty TNHH Nhơn Trạch Phát", "address": "Khu công nghiệp Nhơn Trạch, Đồng Nai", "date": "2026-10-19"},
        {"name": "Công ty CP Long Thành Logistics", "address": "Huyện Long Thành, Đồng Nai", "date": "2026-10-19"},
        {"name": "Công ty TNHH Biên Hòa Phúc", "address": "Thành phố Biên Hòa, Đồng Nai", "date": "2026-10-19"}
    ]
    return companies

def send_email(companies):
    if not companies:
        print("Không có công ty mới.")
        return

    sender_email = os.environ.get("MAIL_USERNAME")
    sender_password = os.environ.get("MAIL_PASSWORD")
    receiver_email = os.environ.get("MAIL_USERNAME")

    html_content = "<h2>Danh sách công ty mới đăng ký (Biên Hòa, Nhơn Trạch, Long Thành, Đồng Nai):</h2><ul>"
    for c in companies:
        html_content += f"<li><b>{c['name']}</b> - Địa chỉ: {c['address']} (Ngày ĐK: {c['date']})</li>"
    html_content += "</ul>"

    message = MIMEMultipart("alternative")
    message["Subject"] = "Báo cáo công ty mới khu vực Đồng Nai"
    message["From"] = sender_email
    message["To"] = receiver_email
    message.attach(MIMEText(html_content, "html"))

    try:
        print(f"Đang kết nối tới Gmail qua cổng 587 cho tài khoản: {sender_email}")
        # Chuyển sang cổng 587 và STARTTLS - cổng tối ưu nhất cho GitHub Actions
        server = smtplib.SMTP("smtp.gmail.com", 587, timeout=30)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        server.quit()
        print("Gửi email thành công!")
    except Exception as e:
        print(f"Lỗi gửi email chi tiết: {e}")

if __name__ == "__main__":
    data = check_new_companies()
    send_email(data)
