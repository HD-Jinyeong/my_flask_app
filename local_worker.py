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

s3 = boto3.client(
    "s3",
    aws_access_key_id=AWS_ACCESS_KEY_ID,
    aws_secret_access_key=AWS_SECRET_ACCESS_KEY,
    region_name=S3_REGION
)

# ================== 담당자 DB ==================
responsibles = [
    {"name": "최현서", "email": "jinyeong@hd.com"},
    {"name": "하태현", "email": "wlsdud5706@naver.com"},
    {"name": "전민수", "email": "wlsdud706@knu.ac.kr"}
]

# ================== 메일 발송 함수 ==================
def send_mail(to_addr, subject, body):
    msg = MIMEText(body, _charset="utf-8")
    msg["From"] = formataddr((FROM_NAME, FROM_ADDR))
    msg["To"] = to_addr
    msg["Subject"] = subject

    with smtplib.SMTP(SMTP_SERVER) as server:
        server.sendmail(FROM_ADDR, [to_addr], msg.as_string())
        print(f"메일 전송 성공 → {to_addr}")

# ================== S3 파일 확인 후 메일 전송 ==================
def process_and_send():
    resp = s3.list_objects_v2(Bucket=S3_BUCKET, Prefix="submissions/")
    if "Contents" not in resp:
        print("DEBUG: submissions/ 없음")
        return

    # 담당자별 미입력 장비 모으기
    pending_by_person = {p["email"]: [] for p in responsibles}

    for obj in resp["Contents"]:
        if not obj["Key"].endswith(".json"):
            continue

        data = s3.get_object(Bucket=S3_BUCKET, Key=obj["Key"])["Body"].read()
        rows = json.loads(data)

        for row in rows:
            if row.get("status") != "done":  # 미입력 상태
                person = row["responsible"]["email"]
                pending_by_person[person].append(
                    (row.get("ship_number"), row.get("category"), row.get("equipment_name"))
                )

    # 담당자별 메일 발송 (중복 없이)
    for person in responsibles:
        email = person["email"]
        tasks = pending_by_person[email]
        if not tasks:
            continue

        subject = "[HD Hyundai Mipo] 미입력 장비 자동 알림"
        body = f"{person['name']}님,\n\n다음 장비 입력이 아직 완료되지 않았습니다:\n\n"
        for ship, cat, eq in tasks:
            body += f"- Ship {ship} / {cat} / {eq}\n"
        body += "\n이 메일은 매일 오후 4시에 자동 발송됩니다.\n감사합니다."

        send_mail(email, subject, body)

# ================== 메인 루프 ==================
if __name__ == "__main__":
    print("Local Worker 실행 중 (매일 16:00 자동 발송) ...")
    while True:
        now = datetime.datetime.now()
        if now.hour == 16 and now.minute == 0:  # 오후 4시 정각
            process_and_send()
            time.sleep(60)  # 중복 발송 방지 (1분 대기)
        time.sleep(5)
