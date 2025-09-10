import os, boto3, json, smtplib, time
from email.mime.text import MIMEText
from email.utils import formataddr

# 회사 SMTP 서버
SMTP_SERVER = "211.193.193.12"
FROM_ADDR   = "noreply@company.com"
FROM_NAME   = "HD Hyundai Mipo 자동 발송"

# AWS S3 설정 (Render와 동일)
S3_BUCKET = os.getenv("S3_BUCKET")
S3_REGION = os.getenv("S3_REGION", "ap-northeast-2")
AWS_ACCESS_KEY_ID = os.getenv("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = os.getenv("AWS_SECRET_ACCESS_KEY")

print("DEBUG:", S3_BUCKET, S3_REGION, AWS_ACCESS_KEY_ID)  # ✅ 환경변수 확인용

s3 = boto3.client(
    "s3",
    aws_access_key_id=AWS_ACCESS_KEY_ID,
    aws_secret_access_key=AWS_SECRET_ACCESS_KEY,
    region_name=S3_REGION
)

def send_mail(to_addr, subject, body):
    msg = MIMEText(body, _charset="utf-8")
    msg["From"] = formataddr((FROM_NAME, FROM_ADDR))
    msg["To"] = to_addr
    msg["Subject"] = subject

    with smtplib.SMTP(SMTP_SERVER) as server:
        server.sendmail(FROM_ADDR, [to_addr], msg.as_string())
        print(f"메일 전송 성공 → {to_addr}")

def process_s3_files():
    resp = s3.list_objects_v2(Bucket=S3_BUCKET, Prefix="submissions/")
    if "Contents" not in resp:
        print("DEBUG: submissions/ 안에 파일 없음")
        return

    for obj in resp["Contents"]:
        key = obj["Key"]
        if not key.endswith(".json"):
            continue

        data = s3.get_object(Bucket=S3_BUCKET, Key=key)["Body"].read()
        rows = json.loads(data)

        for row in rows:
            # ⚠️ 테스트용: 조건 제거 → 무조건 메일 전송
            subject = f"[HD Hyundai Mipo] {row['project_name']} 제출 확인"
            body = f"""{row['submitter_name']}님,

프로젝트 [{row['project_name']}] 데이터가 접수되었습니다.

감사합니다.
"""
            send_mail(row["submitter_email"], subject, body)

if __name__ == "__main__":
    while True:
        process_s3_files()
        time.sleep(60)  # 1분마다 확인
