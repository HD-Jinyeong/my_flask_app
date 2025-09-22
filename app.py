import os, io, json, uuid, datetime, random
from werkzeug.utils import secure_filename
from flask import Flask, request, render_template, redirect, url_for, send_file, flash, jsonify
import boto3
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "dev-only-change-me")

# ================ 환경변수 ================
S3_BUCKET = os.getenv("S3_BUCKET")
S3_REGION = os.getenv("S3_REGION", "ap-northeast-2")
AWS_ACCESS_KEY_ID = os.getenv("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = os.getenv("AWS_SECRET_ACCESS_KEY")

# 로컬에서만 admin 페이지 열기
ADMIN_ENABLED = os.getenv("ADMIN_ENABLED", "true").lower() == "true"

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

# ================== 담당자 DB ==================
responsibles = [
    {"name": "최현서", "email": "jinyeong@hd.com", "phone": "010-0000-0000"},
    {"name": "하태현", "email": "wlsdud5706@naver.com", "phone": "010-0000-0000"},
    {"name": "전민수", "email": "wlsdud706@knu.ac.kr", "phone": "010-0000-0000"}
]

# ================= 사용자 제출 =================
@app.route("/", methods=["GET"])
def home():
    ship_number = request.args.get("ship_number")
    catalog_path = os.path.join("config", "equipment_catalog.json")
    catalog = {}
    if os.path.exists(catalog_path):
        with open(catalog_path, "r", encoding="utf-8") as f:
            catalog = json.load(f)
    return render_template("home.html", catalog=catalog, selected_ship=ship_number)

@app.route("/submit", methods=["POST"])
def submit():
    submitter_name  = request.form.get("submitter_name")
    submitter_email = request.form.get("submitter_email")
    contact         = request.form.get("contact")
    affiliation     = request.form.get("affiliation")
    submit_date     = request.form.get("submit_date")
    ship_number     = request.form.get("ship_number")
    due_date        = request.form.get("due_date")

    category = request.form.get("category")
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
    now    = datetime.datetime.now(datetime.timezone.utc).isoformat()

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
            folder = submit_date if submit_date else datetime.datetime.now().strftime("%Y-%m-%d")
            key = f"uploads/{folder}/{sub_id}_{i}_{safe}"
            s3.upload_fileobj(f, S3_BUCKET, key, ExtraArgs={"ContentType": f.mimetype})
            file_url = presigned_url(key)
            filename = safe

        # 담당자 랜덤 배정
        resp = random.choice(responsibles)

        row = {
            "id": sub_id,
            "submitter_name": submitter_name,
            "submitter_email": submitter_email,
            "contact": contact,
            "affiliation": affiliation,
            "submit_date": submit_date,
            "ship_number": ship_number,
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
            "due_date": due_date,
            "responsible": resp,
            "status": "pending"
        }
        rows.append(row)

    json_key = f"submissions/{sub_id}.json"
    put_json_to_s3(json_key, rows)
    flash("제출이 완료되었습니다.")
    return redirect(url_for("home"))

# ================== 장비 수정 ==================
@app.route("/edit/<ship_number>/<category>/<eq>", methods=["GET", "POST"])
def edit_equipment(ship_number, category, eq):
    catalog_path = os.path.join("config", "equipment_catalog.json")
    if not os.path.exists(catalog_path):
        return "Catalog not found", 404

    with open(catalog_path, "r", encoding="utf-8") as f:
        catalog = json.load(f)

    if request.method == "POST":
        qty   = request.form.get("qty")
        maker = request.form.get("maker")
        typ   = request.form.get("type")
        cert  = request.form.get("cert_no")

        # 값 저장
        catalog[category][eq]["qty"] = qty
        catalog[category][eq]["maker"] = maker
        catalog[category][eq]["type"] = typ
        catalog[category][eq]["cert_no"] = cert
        catalog[category][eq]["status"] = "done"

        with open(catalog_path, "w", encoding="utf-8") as f:
            json.dump(catalog, f, ensure_ascii=False, indent=2)

        flash("장비 정보가 업데이트되었습니다.")
        return redirect(url_for("home", ship_number=ship_number))

    info = catalog[category][eq]
    return render_template(
        "edit_equipment.html",
        ship_number=ship_number, category=category, eq=eq, info=info
    )

# ================== Admin 기능 ==================
if ADMIN_ENABLED:

    @app.route("/admin")
    def admin_dashboard():
        s3 = s3_client()
        resp = s3.list_objects_v2(Bucket=S3_BUCKET, Prefix="submissions/")
        submissions = []

        if "Contents" in resp:
            for obj in resp["Contents"]:
                if not obj["Key"].endswith(".json"): continue
                data = s3.get_object(Bucket=S3_BUCKET, Key=obj["Key"])["Body"].read()
                rows = json.loads(data)
                submissions.extend(rows)

        return render_template("admin.html", submissions=submissions)

    @app.route("/admin/manual_mail", methods=["POST"])
    def manual_mail():
        from local_worker import process_and_send
        process_and_send()
        flash("📧 수동 메일 발송 완료")
        return redirect(url_for("admin_dashboard"))

    def write_grouped_excel(submissions, wb):
        ws = wb.active
        ws.title = "Submissions"
        ws.append([
            "ID","Ship Number","Submitter Name","Email","Category",
            "Equipment Name","QTY","Maker","Type","Cert No.",
            "Ex-proof Grade","IP Grade","Page","File","Status","Responsible"
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
                s.get("status",""),
                s.get("responsible",{}).get("name","")
            ])

    @app.route("/export/excel")
    def export_excel():
        s3 = s3_client()
        resp = s3.list_objects_v2(Bucket=S3_BUCKET, Prefix="submissions/")
        submissions = []

        if "Contents" in resp:
            for obj in resp["Contents"]:
                if not obj["Key"].endswith(".json"): continue
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

# ================== Health Check ==================
@app.route("/health")
def health():
    return "ok", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
