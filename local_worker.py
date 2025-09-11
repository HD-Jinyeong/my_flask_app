import os, boto3, json, smtplib, time, datetime
from email.mime.text import MIMEText
from email.utils import formataddr
from dotenv import load_dotenv

load_dotenv()

# ================== 회사 SMTP 서버 ==================
SMTP_SERVER = "211.193.193.12"
FROM_ADDR   = "noreply@company.com"
FROM_NAME   = "HD Hyundai Mipo 자동 발송"

# ================== AWS S3 설정 ==================
S3_BUCKET = os.getenv("S3_BUCKET")
S3_REGION = os.getenv("S3_REGION", "ap-northeast-2")
AWS_ACCESS_KEY_ID = os.getenv("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = os.getenv("AWS_SECRET_ACCESS_KEY")

print("DEBUG ENV:", S3_BUCKET, S3_REGION, AWS_ACCESS_KEY_ID)

s3 = boto3.client(
    "s3",
    aws_access_key_id=AWS_ACCESS_KEY_ID,
    aws_secret_access_key=AWS_SECRET_ACCESS_KEY,
    region_name=S3_REGION
)

# ================== 메일 발송 함수 ==================
def send_mail(to_addr, subject, body):
    msg = MIMEText(body, _charset="utf-8")
    msg["From"] = formataddr((FROM_NAME, FROM_ADDR))
    msg["To"] = to_addr
    msg["Subject"] = subject

    with smtplib.SMTP(SMTP_SERVER) as server:
        server.sendmail(FROM_ADDR, [to_addr], msg.as_string())
        print(f"메일 전송 성공 → {to_addr}")

# ================== S3 파일 처리 ==================
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

        changed = False

        for row in rows:
            last_mail_sent = row.get("last_mail_sent")
            last_updated   = row.get("last_updated")
            force_send     = row.get("force_send", False)

            # ================== 발송 조건 ==================
            if force_send or (not last_mail_sent) or (last_updated and last_updated > last_mail_sent):
                subject = f"[HD Hyundai Mipo] {row['project_name']} 수정 요청드립니다."
                body = f"""{row['submitter_name']}님, 

평소 업무 협조에 감사드립니다.

표제의 건에 관련하여, [{row['project_name']}] 데이터에 수정이 필요합니다.

요청 사유:
{row.get('message', '사유 미입력')}

수정 기한: {row.get('due_date', '날짜 미지정')}까지 부탁드립니다.

감사합니다.
"""
                send_mail(row["submitter_email"], subject, body)

                # ✅ 발송 후 기록 업데이트
                row["last_mail_sent"] = datetime.datetime.now(datetime.UTC).isoformat()
                row["force_send"] = False  # 강제 발송은 1회 처리 후 해제
                changed = True

        # ================== 변경된 경우 JSON 다시 업로드 ==================
        if changed:
            s3.put_object(
                Bucket=S3_BUCKET,
                Key=key,
                Body=json.dumps(rows, ensure_ascii=False, indent=2),
                ContentType="application/json"
            )
            print(f"DEBUG: {key} 업데이트 완료 (메일 발송 기록 반영됨)")


# ================== 메인 루프 ==================
if __name__ == "__main__":
    while True:
        process_s3_files()
        print("DEBUG: 10초 대기중...")
        time.sleep(10)
