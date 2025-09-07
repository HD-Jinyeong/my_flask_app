import os, io, json, uuid, datetime
from werkzeug.utils import secure_filename
from flask import Flask, request, render_template, redirect, url_for, send_file
import boto3
from openpyxl import Workbook
from docx import Document

app = Flask(__name__)

# Render → Environment에 등록한 값 사용
S3_BUCKET = os.getenv("S3_BUCKET")
S3_REGION = os.getenv("S3_REGION", "ap-northeast-2")
AWS_ACCESS_KEY_ID = os.getenv("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = os.getenv("AWS_SECRET_ACCESS_KEY")

def s3_client():
    return boto3.client(
        "s3",
        aws_access_key_id=AWS_ACCESS_KEY_ID,
        aws_secret_access_key=AWS_SECRET_ACCESS_KEY,
        region_name=S3_REGION
    )

# 메모리에 최근 제출 저장(데모용)
SUBMISSIONS = []

def presigned_url(key, expires=3600*24*7):
    """Private 버킷이라도 접근 가능한 임시 URL 생성(7일 기본)."""
    s3 = s3_client()
    return s3.generate_presigned_url(
        "get_object",
        Params={"Bucket": S3_BUCKET, "Key": key},
        ExpiresIn=expires
    )

@app.route("/", methods=["GET"])
def home():
    return render_template("home.html", title="데이터 수집", submissions=SUBMISSIONS)

@app.route("/submit", methods=["POST"])
def submit():
    name = request.form.get("name")
    memo = request.form.get("memo")
    date = request.form.get("date")  # YYYY-MM-DD
    files = request.files.getlist("files")

    sub_id = str(uuid.uuid4())[:8]
    now = datetime.datetime.utcnow().isoformat() + "Z"

    s3 = s3_client()
    uploaded_files = []

    for f in files:
        if not f or f.filename == "":
            continue
        safe = secure_filename(f.filename)
        # 키 규칙: submissions/날짜/uuid_파일명
        folder = date if date else datetime.datetime.utcnow().strftime("%Y-%m-%d")
        key = f"uploads/{folder}/{sub_id}_{safe}"

        # 업로드
        s3.upload_fileobj(
            f, S3_BUCKET, key,
            ExtraArgs={"ContentType": f.mimetype}  # Content-Type 유지
        )

        url = presigned_url(key)  # 버킷 private이어도 접근 가능
        uploaded_files.append({"filename": safe, "key": key, "url": url})

    submission = {
        "id": sub_id,
        "name": name,
        "memo": memo,
        "date": date,
        "timestamp": now,
        "files": uploaded_files,
    }
    SUBMISSIONS.append(submission)

    # 간단한 영속성: 제출 JSON을 S3에 저장
    s3.put_object(
        Bucket=S3_BUCKET,
        Key=f"submissions/{sub_id}.json",
        Body=json.dumps(submission, ensure_ascii=False).encode("utf-8"),
        ContentType="application/json"
    )

    return redirect(url_for("home"))

@app.route("/export/excel")
def export_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Submissions"
    ws.append(["id", "name", "date", "memo", "timestamp", "filename", "file_url"])

    for s in SUBMISSIONS:
        if s["files"]:
            for f in s["files"]:
                ws.append([s["id"], s["name"], s["date"] or "", s["memo"] or "", s["timestamp"], f["filename"], f["url"]])
        else:
            ws.append([s["id"], s["name"], s["date"] or "", s["memo"] or "", s["timestamp"], "", ""])

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return send_file(stream, as_attachment=True, download_name="submissions.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/export/word")
def export_word():
    doc = Document()
    doc.add_heading("제출 내역", level=1)
    for s in SUBMISSIONS:
        doc.add_paragraph(f"ID: {s['id']}")
        doc.add_paragraph(f"이름: {s['name']}")
        doc.add_paragraph(f"날짜: {s['date'] or '-'}")
        if s["memo"]:
            doc.add_paragraph(f"메모: {s['memo']}")
        doc.add_paragraph(f"시간: {s['timestamp']}")
        if s["files"]:
            doc.add_paragraph("파일:")
            for f in s["files"]:
                # 워드는 하이퍼링크 API가 번거롭지만, URL 텍스트만 써도 자동 인식됨
                doc.add_paragraph(f" - {f['filename']} : {f['url']}")
        doc.add_paragraph("")  # 빈 줄
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return send_file(stream, as_attachment=True, download_name="submissions.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# 헬스체크
@app.route("/health")
def health():
    return "ok", 200
