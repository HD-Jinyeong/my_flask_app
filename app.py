import os, io, json, uuid, datetime
from werkzeug.utils import secure_filename
from flask import Flask, request, render_template, redirect, url_for, send_file, flash
import boto3
from openpyxl import Workbook
from docx import Document
import smtplib
from email.mime.text import MIMEText

app = Flask(__name__)
app.secret_key = "secret-key-for-flash"  # flash 메시지용

# 환경변수 (Render → 환경변수 등록 필수)
S3_BUCKET = os.getenv("S3_BUCKET")
S3_REGION = os.getenv("S3_REGION", "ap-northeast-2")
AWS_ACCESS_KEY_ID = os.getenv("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = os.getenv("AWS_SECRET_ACCESS_KEY")
ADMIN_EMAIL = os.getenv("ADMIN_EMAIL")
ADMIN_EMAIL_PASSWORD = os.getenv("ADMIN_EMAIL_PASSWORD")

def s3_client():
    return boto3.client(
        "s3",
        aws_access_key_id=AWS_ACCESS_KEY_ID,
        aws_secret_access_key=AWS_SECRET_ACCESS_KEY,
        region_name=S3_REGION
    )

SUBMISSIONS = []  # 메모리에 저장 (데모용)

def presigned_url(key, expires=3600*24*7):
    s3 = s3_client()
    return s3.generate_presigned_url(
        "get_object",
        Params={"Bucket": S3_BUCKET, "Key": key},
        ExpiresIn=expires
    )

# 이메일 발송 함수
def send_mail(to_email, subject, body):
    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = subject
    msg["From"] = ADMIN_EMAIL
    msg["To"] = to_email
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(ADMIN_EMAIL, ADMIN_EMAIL_PASSWORD)
        server.send_message(msg)

# ================= 사용자 제출 =================

@app.route("/", methods=["GET"])
def home():
    return render_template("form.html")

@app.route("/submit", methods=["POST"])
def submit():
    # 제출자 기본 정보
    submitter_name = request.form.get("submitter_name")
    submitter_email = request.form.get("submitter_email")
    submit_date = request.form.get("submit_date")
    project_name = request.form.get("project_name")

    # 행 단위 데이터
    names = request.form.getlist("equipment_name[]")
    qtys = request.form.getlist("qty[]")
    types = request.form.getlist("type[]")
    certs = request.form.getlist("cert_no[]")
    exgrades = request.form.getlist("ex_proof_grade[]")
    ipgrades = request.form.getlist("ip_grade[]")
    pages = request.form.getlist("page[]")
    files = request.files.getlist("file[]")

    sub_id = str(uuid.uuid4())[:8]
    now = datetime.datetime.utcnow().isoformat() + "Z"

    s3 = s3_client()
    rows = []

    for i in range(len(names)):
        file_url, filename = None, None
        if i < len(files) and files[i] and files[i].filename != "":
            f = files[i]
            safe = secure_filename(f.filename)
            folder = submit_date if submit_date else datetime.datetime.utcnow().strftime("%Y-%m-%d")
            key = f"uploads/{folder}/{sub_id}_{i}_{safe}"
            s3.upload_fileobj(f, S3_BUCKET, key, ExtraArgs={"ContentType": f.mimetype})
            file_url = presigned_url(key)
            filename = safe

        row = {
            "id": sub_id,
            "submitter_name": submitter_name,
            "submitter_email": submitter_email,
            "submit_date": submit_date,
            "project_name": project_name,
            "equipment_name": names[i],
            "qty": qtys[i],
            "type": types[i],
            "cert_no": certs[i],
            "ex_proof_grade": exgrades[i],
            "ip_grade": ipgrades[i],
            "page": pages[i],
            "file": filename,
            "file_url": file_url,
            "timestamp": now
        }
        SUBMISSIONS.append(row)
        rows.append(row)

    # JSON 저장
    s3.put_object(
        Bucket=S3_BUCKET,
        Key=f"submissions/{sub_id}.json",
        Body=json.dumps(rows, ensure_ascii=False).encode("utf-8"),
        ContentType="application/json"
    )

    return redirect(url_for("home"))

# ================= 관리자 페이지 =================

@app.route("/admin")
def admin_dashboard():
    return render_template("admin.html", submissions=SUBMISSIONS)

@app.route("/admin/edit/<id>", methods=["GET", "POST"])
def edit_submission(id):
    s = next((x for x in SUBMISSIONS if x["id"] == id), None)
    if not s:
        return "Not found", 404
    if request.method == "POST":
        s["equipment_name"] = request.form.get("equipment_name")
        s["qty"] = request.form.get("qty")
        flash("수정 완료")
        return redirect(url_for("admin_dashboard"))
    return f"""
    <h3>수정 페이지 (ID: {s['id']})</h3>
    <form method="post">
      장비명: <input type="text" name="equipment_name" value="{s['equipment_name']}"><br>
      수량: <input type="number" name="qty" value="{s['qty']}"><br>
      <button type="submit">저장</button>
    </form>
    """

@app.route("/admin/mail/<id>", methods=["GET"])
def mail_form(id):
    s = next((x for x in SUBMISSIONS if x["id"] == id), None)
    if not s:
        return "Not found", 404
    return render_template("mail_form.html", submission=s)

@app.route("/admin/mail_send/<id>", methods=["POST"])
def mail_send(id):
    s = next((x for x in SUBMISSIONS if x["id"] == id), None)
    if not s:
        return "Not found", 404
    due_date = request.form.get("due_date")
    message = request.form.get("message")
    subject = f"[HD 현대미포] {s['project_name']} 재입력 요청"
    body = f"""
    {s['submitter_name']}님,

    제출하신 프로젝트 [{s['project_name']}] ({s['submit_date']}) 데이터에 수정이 필요합니다.

    요청 사유:
    {message}

    수정 기한: {due_date}

    제출자 이메일: {s['submitter_email']}
    다시 입력 링크: https://내도메인/

    감사합니다.
    """
    send_mail(s["submitter_email"], subject, body)
    flash("재입력 요청 메일 발송 완료")
    return redirect(url_for("admin_dashboard"))

@app.route("/export/excel")
def export_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Submissions"
    ws.append(["id","submitter_name","submitter_email","project_name",
               "equipment_name","qty","file","file_url","timestamp"])
    for s in SUBMISSIONS:
        ws.append([s["id"], s["submitter_name"], s["submitter_email"], s["project_name"],
                   s["equipment_name"], s["qty"], s.get("file",""), s.get("file_url",""), s["timestamp"]])
    stream = io.BytesIO()
    wb.save(stream); stream.seek(0)
    return send_file(stream, as_attachment=True,
                     download_name="submissions.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/admin/export_selected", methods=["POST"])
def export_selected():
    selected_ids = request.form.getlist("selected_ids")
    if not selected_ids:
        return "선택된 항목이 없습니다.", 400
    wb = Workbook()
    ws = wb.active
    ws.title = "Selected"
    ws.append(["id","submitter_name","submitter_email","project_name",
               "equipment_name","qty","file","file_url","timestamp"])
    for s in SUBMISSIONS:
        if s["id"] in selected_ids:
            ws.append([s["id"], s["submitter_name"], s["submitter_email"], s["project_name"],
                       s["equipment_name"], s["qty"], s.get("file",""), s.get("file_url",""), s["timestamp"]])
    stream = io.BytesIO()
    wb.save(stream); stream.seek(0)
    return send_file(stream, as_attachment=True,
                     download_name="selected_submissions.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/health")
def health():
    return "ok", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
