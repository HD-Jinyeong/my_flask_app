import smtplib
from email.mime.text import MIMEText
from email.utils import formataddr

def send_mail(to_email, subject, body):
    from_addr = "noreply@company.com"   # 회사 도메인 주소
    from_name = "HD Hyundai Mipo"

    msg = MIMEText(body, _charset="utf-8")
    msg["From"] = formataddr((from_name, from_addr))
    msg["To"] = to_email
    msg["Subject"] = subject

    smtp_server = "211.193.193.12"  # 회사 SMTP relay 서버
    with smtplib.SMTP(smtp_server) as server:
        server.sendmail(from_addr, [to_email], msg.as_string())
        print("메일 전송 성공 (relay)")
