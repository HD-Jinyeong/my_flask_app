# app.py
import os, io, json, uuid, datetime
from werkzeug.utils import secure_filename
from flask import Flask, request, render_template, redirect, url_for, send_file, flash, jsonify
import boto3
from openpyxl import Workbook



app = Flask(__name__)
# app.secret_key = "secret-key-for-flash"  # 환경변수 사용 권장
app.secret_key = os.getenv("SECRET_KEY", "dev-only-change-me")

# ================ 환경변수 ================
S3_BUCKET = os.getenv("S3_BUCKET")
S3_REGION = os.getenv("S3_REGION", "ap-northeast-2")
AWS_ACCESS_KEY_ID = os.getenv("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = os.getenv("AWS_SECRET_ACCESS_KEY")

CATALOG_KEY = "config/equipment_catalog.json"
CONTACTS_KEY = "config/contacts.json"

def s3_client():
    return boto3.client(
        "s3",
        aws_access_key_id=AWS_ACCESS_KEY_ID,
        aws_secret_access_key=AWS_SECRET_ACCESS_KEY,
        region_name=S3_REGION
    )

def presigned_url(key, expires=3600*24*7):
    s3 = s3_client()
    return s3.generate_presigned_url(
        "get_object",
        Params={"Bucket": S3_BUCKET, "Key": key},
        ExpiresIn=expires
    )

def get_json_from_s3(key, default=None):
    s3 = s3_client()
    try:
        obj = s3.get_object(Bucket=S3_BUCKET, Key=key)
        return json.loads(obj["Body"].read())
    except Exception:
        return default if default is not None else {}

def put_json_to_s3(key, data):
    s3 = s3_client()
    s3.put_object(
        Bucket=S3_BUCKET,
        Key=key,
        Body=json.dumps(data, ensure_ascii=False, indent=2),
        ContentType="application/json"
    )

# ================= 사용자 제출 =================
@app.route("/", methods=["GET"])
def home():
    # 단순 폼(카탈로그/연락처는 프런트가 /api에서 로드)
    return render_template("form.html")

@app.route("/submit", methods=["POST"])
def submit():
    submitter_name  = request.form.get("submitter_name")
    submitter_email = request.form.get("submitter_email")
    contact         = request.form.get("contact")
    affiliation     = request.form.get("affiliation")
    submit_date     = request.form.get("submit_date")
    ship_number     = request.form.get("ship_number")  # 변경
    due_date        = request.form.get("due_date")      # 신규

    # 참고/수신자 선택(다중)
    cc_emails = request.form.getlist("cc_emails[]")   # 신규
    to_emails = request.form.getlist("to_emails[]")   # 신규(없으면 submitter로 대체)

    category = request.form.get("category")
    if category == "Other":
        other_category = request.form.get("other_category")
        if other_category:
            category = other_category

    names    = request.form.getlist("equipment_name[]")
    qtys     = request.form.getlist("qty[]")
    makers   = request.form.getlist("maker[]")
    types    = request.form.getlist("type[]")
    certs    = request.form.getlist("cert_no[]")
    exgrades = request.form.getlist("ex_proof_grade[]")
    ipgrades = request.form.getlist("ip_grade[]")
    pages    = request.form.getlist("page[]")
    files    = request.files.getlist("file[]")

    sub_id = str(uuid.uuid4())[:8]
    now    = datetime.datetime.utcnow().isoformat()

    s3 = s3_client()
    rows = []

    max_len = max(len(names), len(qtys), len(makers), len(types), len(certs), len(exgrades), len(ipgrades), len(pages))
    for i in range(max_len):
        name  = names[i]    if i < len(names)    else ""
        qty   = qtys[i]     if i < len(qtys)     else ""
        maker = makers[i]   if i < len(makers)   else ""
        typ   = types[i]    if i < len(types)    else ""
        cert  = certs[i]    if i < len(certs)    else ""
        exg   = exgrades[i] if i < len(exgrades) else ""
        ipg   = ipgrades[i] if i < len(ipgrades) else ""
        page  = pages[i]    if i < len(pages)    else ""

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
            "contact": contact,
            "affiliation": affiliation,
            "submit_date": submit_date,
            "ship_number": ship_number,          # 신규
            "project_name": ship_number,         # 하위호환
            "category": category,
            "equipment_name": name,
            "qty": qty,
            "maker": maker,
            "type": typ,
            "cert_no": cert,
            "ex_proof_grade": exg,
            "ip_grade": ipg,
            "page": page,
            "file": filename,
            "file_url": file_url,
            "timestamp": now,
            # 자동메일 파라미터
            "due_date": due_date,
            "to_emails": to_emails if to_emails else [submitter_email],
            "cc_emails": cc_emails,
            "reminder_status": { "-30": False, "-14": False, "-7": False }
        }
        rows.append(row)

    json_key = f"submissions/{sub_id}.json"
    put_json_to_s3(json_key, rows)
    flash("제출이 완료되었습니다.")
    return redirect(url_for("home"))

# ================ 카탈로그/연락처 API ================
@app.route("/api/catalog")
def api_catalog():
    return jsonify(get_json_from_s3(CATALOG_KEY, default={}))

@app.route("/api/contacts")
def api_contacts():
    return jsonify(get_json_from_s3(CONTACTS_KEY, default=[]))

# ================= 관리자 기능 =================
@app.route("/admin")
def admin_dashboard():
    s3 = s3_client()
    resp = s3.list_objects_v2(Bucket=S3_BUCKET, Prefix="submissions/")
    submissions = []

    if "Contents" in resp:
        for obj in resp["Contents"]:
            if not obj["Key"].endswith(".json"):
                continue
            data = s3.get_object(Bucket=S3_BUCKET, Key=obj["Key"])["Body"].read()
            rows = json.loads(data)
            submissions.extend(rows)

    # ship_number 기준 / category / equipment_name 정렬
    submissions.sort(key=lambda x: (
        x.get("ship_number",""),
        x.get("category",""),
        x.get("equipment_name","")
    ))

    return render_template("admin.html", submissions=submissions)

@app.route("/admin/mail/<id>", methods=["GET"])
def mail_form(id):
    s3 = s3_client()
    key = f"submissions/{id}.json"

    try:
        data = s3.get_object(Bucket=S3_BUCKET, Key=key)["Body"].read()
        rows = json.loads(data)
        submission = rows[0]  # JSON은 리스트이므로 첫 번째 항목
    except Exception:
        return "Not found", 404

    return render_template("mail_form.html", submission=submission)

@app.route("/admin/mail_send/<id>", methods=["POST"])
def mail_send(id):
    due_date = request.form.get("due_date")
    message  = request.form.get("message")
    to_emails = request.form.getlist("to_emails[]")
    cc_emails = request.form.getlist("cc_emails[]")
    now = datetime.datetime.utcnow().isoformat()

    # S3 JSON 불러오기
    s3 = s3_client()
    key = f"submissions/{id}.json"
    obj = s3.get_object(Bucket=S3_BUCKET, Key=key)
    rows = json.loads(obj["Body"].read())

    # 모든 row 업데이트
    for row in rows:
        row["due_date"] = due_date or row.get("due_date")
        row["message"]  = message
        row["to_emails"] = to_emails or row.get("to_emails") or [row.get("submitter_email")]
        row["cc_emails"] = cc_emails or row.get("cc_emails") or []
        row["last_updated"] = now
        row["force_send"] = True   # ✅ 강제 발송 플래그

    # 다시 S3 업로드
    put_json_to_s3(key, rows)

    flash("메일 발송 요청이 저장되었습니다. (로컬 worker가 처리)")
    return redirect(url_for("admin_dashboard"))

@app.route("/admin/edit/<id>", methods=["GET", "POST"])
def edit_submission(id):
    s3 = s3_client()
    data = s3.get_object(Bucket=S3_BUCKET, Key=f"submissions/{id}.json")["Body"].read()
    rows = json.loads(data)

    if request.method == "POST":
        for r in rows:
            r["equipment_name"] = request.form.get("equipment_name", r.get("equipment_name"))
            r["qty"]           = request.form.get("qty", r.get("qty"))
            r["maker"]         = request.form.get("maker", r.get("maker"))
            r["type"]          = request.form.get("type", r.get("type"))
            r["cert_no"]       = request.form.get("cert_no", r.get("cert_no"))
            r["category"]      = request.form.get("category", r.get("category"))
            r["ship_number"]   = request.form.get("ship_number", r.get("ship_number"))
            r["due_date"]      = request.form.get("due_date", r.get("due_date"))
        put_json_to_s3(f"submissions/{id}.json", rows)
        return redirect(url_for("admin_dashboard"))

    s = rows[0]
    return render_template("edit.html", s=s)

def write_grouped_excel(submissions, wb):
    ws = wb.active
    ws.title = "Submissions"
    ws.append([
        "ID","Ship Number","Submitter Name","Email","Category",
        "Equipment Name","QTY","Maker","Type","Cert No.",
        "Ex-proof Grade","IP Grade","Page","File","Due Date"
    ])
    for s in submissions:
        ws.append([
            s.get("id",""),
            s.get("ship_number",""),
            s.get("submitter_name",""),
            s.get("submitter_email",""),
            s.get("category",""),
            s.get("equipment_name",""),
            s.get("qty",""),
            s.get("maker",""),
            s.get("type",""),
            s.get("cert_no",""),
            s.get("ex_proof_grade",""),
            s.get("ip_grade",""),
            s.get("page",""),
            s.get("file",""),
            s.get("due_date","")
        ])

@app.route("/export/excel")
def export_excel():
    s3 = s3_client()
    resp = s3.list_objects_v2(Bucket=S3_BUCKET, Prefix="submissions/")
    submissions = []

    if "Contents" in resp:
        for obj in resp["Contents"]:
            if not obj["Key"].endswith(".json"):
                continue
            data = s3.get_object(Bucket=S3_BUCKET, Key=obj["Key"])["Body"].read()
            rows = json.loads(data)
            submissions.extend(rows)

    wb = Workbook()
    write_grouped_excel(submissions, wb)
    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return send_file(stream, as_attachment=True,
                     download_name="submissions.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/health")
def health():
    return "ok", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
