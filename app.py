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

# 🔹 홈 = form.html 보여주기
@app.route("/", methods=["GET"])
def home():
    return render_template("form.html", submissions=SUBMISSIONS)

# 🔹 제출 처리
@app.route("/submit", methods=["POST"])
def submit():
    # 테이블 형태로 받은 값들 (리스트로 수집)
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
        file_url = None
        filename = None
        if i < len(files) and files[i] and files[i].filename != "":
            f = files[i]
            safe = secure_filename(f.filename)
            folder = datetime.datetime.utcnow().strftime("%Y-%m-%d")
            key = f"uploads/{folder}/{sub_id}_{i}_{safe}"
            s3.upload_fileobj(f, S3_BUCKET, key, ExtraArgs={"ContentType": f.mimetype})
            file_url = presigned_url(key)
            filename = safe

        row = {
            "id": sub_id,
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
        rows.append(row)

    SUBMISSIONS.extend(rows)

    # JSON으로도 S3에 저장
    s3.put_object(
        Bucket=S3_BUCKET,
        Key=f"submissions/{sub_id}.json",
        Body=json.dumps(rows, ensure_ascii=False).encode("utf-8"),
        ContentType="application/json"
    )

    return redirect(url_for("home"))

# 🔹 Excel 내보내기
@app.route("/export/excel")
def export_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Equipments"
    ws.append(["id", "equipment_name", "qty", "type", "cert_no", "ex_proof_grade",
               "ip_grade", "page", "file", "file_url", "timestamp"])

    for s in SUBMISSIONS:
        ws.append([s["id"], s["equipment_name"], s["qty"], s["type"], s["cert_no"],
                   s["ex_proof_grade"], s["ip_grade"], s["page"], s["file"] or "", s["file_url"] or "", s["timestamp"]])

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return send_file(stream, as_attachment=True, download_name="equipments.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# 🔹 Word 내보내기
@app.route("/export/word")
def export_word():
    doc = Document()
    doc.add_heading("제출 장비 내역", level=1)
    for s in SUBMISSIONS:
        doc.add_paragraph(f"EQUIPMENT NAME: {s['equipment_name']}")
        doc.add_paragraph(f"QTY: {s['qty']}")
        doc.add_paragraph(f"TYPE: {s['type']}")
        doc.add_paragraph(f"CERT. NO: {s['cert_no']}")
        doc.add_paragraph(f"EX-PROOF GRADE: {s['ex_proof_grade']}")
        doc.add_paragraph(f"IP GRADE: {s['ip_grade']}")
        doc.add_paragraph(f"PAGE: {s['page']}")
        if s["file_url"]:
            doc.add_paragraph(f"FILE: {s['file']} → {s['file_url']}")
        doc.add_paragraph(f"Timestamp: {s['timestamp']}")
        doc.add_paragraph("")  # 빈 줄

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return send_file(stream, as_attachment=True, download_name="equipments.docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

@app.route("/health")
def health():
    return "ok", 200
