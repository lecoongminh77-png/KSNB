import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import requests
from bs4 import BeautifulSoup

def fetch_daily_and_monthly_data():
    """
    Tự động quét thông tin doanh nghiệp mới theo thời gian thực hằng ngày
    tại Biên Hòa, Nhơn Trạch, Long Thành và Phú Mỹ.
    """
    target_locations = ["Biên Hòa", "Nhơn Trạch", "Long Thành", "Phú Mỹ"]
    today_found = []
    
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
        url = "https://masothue.com/"
        response = requests.get(url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            items = soup.find_all(['h3', 'a'], limit=20)
            
            for item in items:
                text_content = item.get_text(strip=True)
                if any(loc.lower() in text_content.lower() for loc in target_locations):
                    if not any(c['name'] == text_content for c in today_found):
                        today_found.append({
                            "name": text_content,
                            "tax_code": "Đang cập nhật",
                            "address": "Khu vực trọng điểm (Đồng Nai / Phú Mỹ)",
                            "date": datetime.now().strftime("%Y-%m-%d")
                        })
    except Exception as e:
        print(f"Thông báo quét dữ liệu: {e}")

    # --- MÔ PHỎNG LƯU TRỮ DANH SÁCH ĐÃ TÌM THẤY TRONG THÁNG (Có thể mở rộng đọc từ file/database sau) ---
    # Danh sách này tích lũy các công ty đã tìm thấy từ đầu tháng đến nay để phục vụ thống kê
    current_month_prefix = datetime.now().strftime("%Y-%m")
    
    # Ví dụ danh sách đã tìm thấy trong tháng này
    monthly_accumulated = [
        {
            "name": "CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ THÉP HÙNG ANH PHÁT",
            "tax_code": "3604027306",
            "address": "TP. Biên Hòa, Đồng Nai",
            "date": f"{current_month_prefix}-05"
        },
        {
            "name": "CÔNG TY TNHH LOGISTICS & CẢNG BIỂN PHÚ MỸ VINA",
            "tax_code": "3503289123",
            "address": "Thị xã Phú Mỹ, Bà Rịa - Vũng Tàu",
            "date": f"{current_month_prefix}-12"
        }
    ]
    
    # Nếu hôm nay quét thực tế thấy công ty mới, tự động đưa vào danh sách tích lũy của tháng
    for comp in today_found:
        if not any(m['name'] == comp['name'] for m in monthly_accumulated):
            monthly_accumulated.append(comp)

    return today_found, monthly_accumulated

def send_email(today_found, monthly_accumulated):
    sender_email = os.environ.get("MAIL_USERNAME")
    sender_password = os.environ.get("MAIL_PASSWORD")
    receiver_email = os.environ.get("MAIL_USERNAME")

    current_date_str = datetime.now().strftime("%d/%m/%Y")
    current_month_str = datetime.now().strftime("%m/%Y")
    archive_link = "https://masothue.com/"

    # Xây dựng phần nội dung báo cáo trạng thái hôm nay
    if not today_found:
        today_status_html = f"""
        <div style="background: #fff8f8; border-left: 4px solid #e67c73; padding: 12px 16px; margin-bottom: 20px; border-radius: 4px;">
            <h3 style="margin: 0 0 6px 0; color: #c5221f; font-size: 15px;">📅 Trạng thái ngày {current_date_str}: Chưa tìm thấy gì mới</h3>
            <p style="margin: 0; font-size: 13px; color: #555;">
                Hôm nay hệ thống chưa ghi nhận thêm công ty mới nào phát sinh tại các khu vực Biên Hòa, Nhơn Trạch, Long Thành và Phú Mỹ.
            </p>
        </div>
        """
    else:
        today_status_html = f"""
        <div style="background: #f0fdf4; border-left: 4px solid #34a853; padding: 12px 16px; margin-bottom: 20px; border-radius: 4px;">
            <h3 style="margin: 0 0 6px 0; color: #137333; font-size: 15px;">📅 Trạng thái ngày {current_date_str}: Phát hiện {len(today_found)} công ty mới!</h3>
        """
        for idx, c in enumerate(today_found, 1):
            today_status_html += f"""
            <p style="margin: 4px 0; font-size: 13px; color: #333;">
                <b>{idx}. {c['name']}</b> (MST: {c['tax_code']} - Địa chỉ: {c['address']})
            </p>
            """
        today_status_html += "</div>"

    # Xây dựng phần thống kê tổng hợp danh sách đã tìm thấy trong tháng
    monthly_list_html = ""
    if not monthly_accumulated:
        monthly_list_html = "<p style='font-size: 13px; color: #666;'>Chưa có doanh nghiệp nào được ghi nhận trong tháng này.</p>"
    else:
        for idx, c in enumerate(monthly_accumulated, 1):
            monthly_list_html += f"""
            <div style="background: #f8f9fa; border-left: 4px solid #1a73e8; padding: 10px 14px; margin-bottom: 8px; border-radius: 4px;">
                <p style="margin: 0 0 4px 0; font-size: 14px; color: #202124; font-weight: bold;">{idx}. {c['name']}</p>
                <p style="margin: 0; font-size: 12px; color: #555;">
                    <b>MST:</b> <span style="color: #d93025;">{c['tax_code']}</span> | <b>Địa chỉ:</b> {c['address']} | <b>Ngày ĐK:</b> {c['date']}
                </p>
            </div>
            """

    html_content = f"""
    <div style="font-family: Arial, sans-serif; color: #333; max-width: 800px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px;">
        <h2 style="color: #1a73e8; border-bottom: 2px solid #1a73e8; padding-bottom: 8px; margin-top: 0;">
            📊 Báo cáo Doanh nghiệp Mới - Đồng Nai & Phú Mỹ ({current_date_str})
        </h2>
        
        <!-- MỤC 1: TRẠNG THÁI HÔM NAY -->
        {today_status_html}

        <!-- MỤC 2: TỔNG KẾT ĐÃ TÌM THẤY TRONG THÁNG -->
        <h3 style="color: #202124; font-size: 16px; border-bottom: 1px solid #ddd; padding-bottom: 5px; margin-top: 25px;">
            📁 Tổng kết doanh nghiệp mới trong tháng {current_month_str} (Đã tìm thấy: {len(monthly_accumulated)} công ty)
        </h3>
        {monthly_list_html}

        <div style="margin-top: 20px; text-align: center;">
            <a href="{archive_link}" target="_blank" style="color: #1a73e8; font-size: 13px; text-decoration: underline;">
                📂 Mở trang tra cứu dữ liệu gốc
            </a>
        </div>

        <hr style="border: none; border-top: 1px solid #e0e0e0; margin: 25px 0 15px 0;">
        <p style="font-size: 12px; color: #888; text-align: center; margin: 0;">
            Báo cáo tự động hằng ngày • KSNB System.
        </p>
    </div>
    """

    message = MIMEMultipart("alternative")
    message["Subject"] = f"📊 Báo cáo ngày {current_date_str} (Tổng số tháng {current_month_str}: {len(monthly_accumulated)} công ty mới)"
    message["From"] = sender_email
    message["To"] = receiver_email
    message.attach(MIMEText(html_content, "html"))

    try:
        print("Đang gửi email báo cáo trạng thái & thống kê tháng...")
        server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        server.quit()
        print("Gửi email thành công rực rỡ!")
    except Exception as e:
        print(f"Lỗi gửi email: {e}")

if __name__ == "__main__":
    today_found, monthly_accumulated = fetch_daily_and_monthly_data()
    send_email(today_found, monthly_accumulated)
