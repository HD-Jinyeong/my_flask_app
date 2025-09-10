import os, io, json, uuid, datetime
from werkzeug.utils import secure_filename
from flask import Flask, request, render_template, redirect, url_for, send_file, flash
import boto3
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = "secret-key-for-flash"

# 환경변수
S3_BUCKET = os.getenv("S3_BUCKET")
S3_REGION = os.getenv("S3_REGION", "ap-northeast-2")
AWS_ACCESS_KEY_ID = os.getenv("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = os.getenv("AWS_SECRET_ACCESS_KEY")

# 로컬 저장 디렉토리 (JSON 백업용)
LOCAL_SUBMISSION_DIR = "submissions"
os.makedirs(LOCAL_SUBMISSION_DIR, exist_ok=True)

# 메모리 저장
SUBMISSIONS = []

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
        SUBMISSIONS.append(row)
        rows.append(row)

    # JSON 백업
    # JSON 백업 (로컬 저장)
    local_path = os.path.join(LOCAL_SUBMISSION_DIR, f"{sub_id}.json")
    with open(local_path, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)

    # ✅ JSON 파일 S3에도 업로드
    s3 = s3_client()
    s3.upload_file(local_path, S3_BUCKET, f"submissions/{sub_id}.json")

    return redirect(url_for("home"))


# ================= 관리자 기능 =================
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
        s["maker"] = request.form.get("maker")
        s["type"] = request.form.get("type")
        s["cert_no"] = request.form.get("cert_no")
        s["category"] = request.form.get("category")
        flash("수정 완료")
        return redirect(url_for("admin_dashboard"))
    return render_template("edit.html", s=s)

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
    s["due_date"] = due_date
    s["message"] = message
    flash("메일 발송 요청이 기록되었습니다. (Render에서는 실제 전송 안 함)")
    return redirect(url_for("admin_dashboard"))

# ================= Excel Export =================
def write_grouped_excel(submissions, wb):
    ws = wb.active
    ws.title = "Submissions"
    ws.append([
        "ID","Submitter Name","Email","Project","Category",
        "Equipment Name","QTY","Maker","Type","Cert No.",
        "Ex-proof Grade","IP Grade","Page","File"
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
            s.get("file","")
        ])

@app.route("/export/excel")
def export_excel():
    wb = Workbook()
    write_grouped_excel(SUBMISSIONS, wb)
    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return send_file(stream, as_attachment=True,
                     download_name="submissions.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/admin/export_selected", methods=["POST"])
def export_selected():
    selected_ids = request.form.getlist("selected_ids")
    if not selected_ids:
        return "선택된 항목이 없습니다.", 400
    selected = [s for s in SUBMISSIONS if s["id"] in selected_ids]
    wb = Workbook()
    write_grouped_excel(selected, wb)
    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return send_file(stream, as_attachment=True,
                     download_name="selected_submissions.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/health")
def health():
    return "ok", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
