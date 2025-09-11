import os, io, json, uuid, datetime
from werkzeug.utils import secure_filename
from flask import Flask, request, render_template, redirect, url_for, send_file, flash
import boto3
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = "secret-key-for-flash"

# ================= 환경변수 =================
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

def presigned_url(key, expires=3600*24*7):
    s3 = s3_client()
    return s3.generate_presigned_url(
        "get_object",
        Params={"Bucket": S3_BUCKET, "Key": key},
        ExpiresIn=expires
    )

# ================= 사용자 제출 =================
@app.route("/", methods=["GET"])
def home():
    return render_template("form.html")

@app.route("/submit", methods=["POST"])
def submit():
    submitter_name = request.form.get("submitter_name")
    submitter_email = request.form.get("submitter_email")
    contact = request.form.get("contact")
    affiliation = request.form.get("affiliation")
    submit_date = request.form.get("submit_date")
    project_name = request.form.get("project_name")

    category = request.form.get("category")
    if category == "Other":
        other_category = request.form.get("other_category")
        if other_category:
            category = other_category

    names = request.form.getlist("equipment_name[]")
    qtys = request.form.getlist("qty[]")
    makers = request.form.getlist("maker[]")
    types = request.form.getlist("type[]")
    certs = request.form.getlist("cert_no[]")
    exgrades = request.form.getlist("ex_proof_grade[]")
    ipgrades = request.form.getlist("ip_grade[]")
    pages = request.form.getlist("page[]")
    files = request.files.getlist("file[]")

    sub_id = str(uuid.uuid4())[:8]
    now = datetime.datetime.utcnow().isoformat()

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
            "contact": contact,
            "affiliation": affiliation,
            "submit_date": submit_date,
            "project_name": project_name,
            "category": category,
            "equipment_name": names[i],
            "qty": qtys[i],
            "maker": makers[i],
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

    # JSON → S3 저장
    json_key = f"submissions/{sub_id}.json"
    s3.put_object(
        Bucket=S3_BUCKET,
        Key=json_key,
        Body=json.dumps(rows, ensure_ascii=False, indent=2),
        ContentType="application/json"
    )

    return redirect(url_for("home"))

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
    message = request.form.get("message")
    now = datetime.datetime.utcnow().isoformat()

    # S3 JSON 불러오기
    s3 = s3_client()
    key = f"submissions/{id}.json"
    obj = s3.get_object(Bucket=S3_BUCKET, Key=key)
    rows = json.loads(obj["Body"].read())

    # 모든 row 업데이트
    for row in rows:
        row["due_date"] = due_date
        row["message"] = message
        row["last_updated"] = now
        row["force_send"] = True   # ✅ 강제 발송 플래그

    # 다시 S3 업로드
    s3.put_object(
        Bucket=S3_BUCKET,
        Key=key,
        Body=json.dumps(rows, ensure_ascii=False, indent=2),
        ContentType="application/json"
    )

    flash("메일 발송 요청이 저장되었습니다. (로컬 worker에서 처리)")
    return redirect(url_for("admin_dashboard"))


@app.route("/admin/edit/<id>", methods=["GET", "POST"])
def edit_submission(id):
    s3 = s3_client()
    data = s3.get_object(Bucket=S3_BUCKET, Key=f"submissions/{id}.json")["Body"].read()
    rows = json.loads(data)
    submission = rows[0]  # 첫 번째 row 사용

    if request.method == "POST":
        submission["equipment_name"] = request.form.get("equipment_name")
        submission["qty"] = request.form.get("qty")
        submission["maker"] = request.form.get("maker")
        submission["type"] = request.form.get("type")

        # 다시 S3에 덮어쓰기
        s3.put_object(
            Bucket=S3_BUCKET,
            Key=f"submissions/{id}.json",
            Body=json.dumps([submission], ensure_ascii=False, indent=2),
            ContentType="application/json"
        )
        return redirect(url_for("admin_dashboard"))

    return render_template("edit.html", s=submission)



def write_grouped_excel(submissions, wb):
    ws = wb.active
    ws.title = "Submissions"
    ws.append([
        "ID", "Submitter Name", "Email", "Project", "Category",
        "Equipment Name", "QTY", "Maker", "Type", "Cert No.",
        "Ex-proof Grade", "IP Grade", "Page", "File"
    ])
    for s in submissions:
        ws.append([
            s.get("id",""),
            s.get("submitter_name",""),
            s.get("submitter_email",""),
            s.get("project_name",""),
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
