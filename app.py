import os, io, json, uuid, datetime, random
from werkzeug.utils import secure_filename
from flask import Flask, request, render_template, redirect, url_for, send_file, flash
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

def get_json_from_file(path, default=None):
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return default if default is not None else {}

def put_json_to_file(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print("[INFO] JSON saved:", path)

# ================== 담당자 DB ==================
responsibles = [
    {"name": "최현서", "email": "jinyeong@hd.com", "phone": "010-0000-0000"},
    {"name": "하태현", "email": "wlsdud5706@naver.com", "phone": "010-0000-0000"},
    {"name": "전민수", "email": "wlsdud706@knu.ac.kr", "phone": "010-0000-0000"}
]

# ================= Ship 별 Due Date =================
SHIP_DUE_DATES = {
    "1": "2025-12-17",
    "2": "2025-12-18",
    "3": "2025-12-19"
}

# ================= 실무 Catalog & EQ List =================
CATALOG_EQUIPMENTS = {
    "Lighting": [
        "Flood Light",
        "Search Light",
        "Navigation Light",
        "Explosion-proof Light"
    ],
    "Switches & Junction Boxes": [
        "Explosion-proof Switch",
        "Explosion-proof Junction Box",
        "Limit Switch",
        "Control Box"
    ],
    "Motor & Machinery": [
        "Explosion-proof Motor",
        "Fan Unit",
        "Pump Unit",
        "Starter Panel"
    ],
    "Communication & Alarm": [
        "Telephone Set",
        "Alarm Bell",
        "Signal Horn",
        "Intercom"
    ],
    "Miscellaneous Equipment": [
        "Heater",
        "Transformer",
        "Battery Charger",
        "Cable Gland"
    ]
}

# ================= Ship별 항목 수량 =================
SHIP_EQUIPMENT_COUNT = {
    "1": 15,
    "2": 15,
    "3": 15
}

def get_or_create_catalog(ship_number, force_reset=False):
    catalog_path = os.path.join("config", f"equipment_catalog_{ship_number}.json")
    print("catalog_path:", catalog_path)

    if force_reset or not os.path.exists(catalog_path):
        print(f"[INFO] Creating new catalog for ship {ship_number}")
        catalog = {}
        total_count = SHIP_EQUIPMENT_COUNT.get(ship_number, 15)

        count = 0
        for category, equipments in CATALOG_EQUIPMENTS.items():
            catalog[category] = {}
            for eq in equipments:
                if count >= total_count:
                    break
                resp = random.choice(responsibles)
                catalog[category][eq] = {
                    "qty": "",
                    "maker": "",
                    "type": "",
                    "cert_no": "",
                    "responsible": resp,
                    "status": "pending",
                    "file": "",
                    "file_url": ""
                }
                count += 1
            if count >= total_count:
                break

        put_json_to_file(catalog_path, catalog)
        return catalog
    else:
        print(f"[INFO] Loading existing catalog for ship {ship_number}")
        return get_json_from_file(catalog_path, default={})

# ================= 홈 =================
@app.route("/", methods=["GET"])
def home():
    ship_number = request.args.get("ship_number")
    catalog = {}
    due_date = None

    if ship_number:
        catalog = get_or_create_catalog(ship_number, force_reset=False)
        due_date = SHIP_DUE_DATES.get(ship_number)

    return render_template("home.html", catalog=catalog, selected_ship=ship_number, due_date=due_date)

# ================= 장비 수정 =================
@app.route("/edit/<ship_number>/<category>/<eq>", methods=["GET", "POST"])
def edit_equipment(ship_number, category, eq):
    catalog_path = os.path.join("config", f"equipment_catalog_{ship_number}.json")
    catalog = get_or_create_catalog(ship_number)

    if request.method == "POST":
        qty   = request.form.get("qty")
        maker = request.form.get("maker")
        typ   = request.form.get("type")
        cert  = request.form.get("cert_no")
        file  = request.files.get("file")

        if category in catalog and eq in catalog[category]:
            catalog[category][eq]["qty"] = qty
            catalog[category][eq]["maker"] = maker
            catalog[category][eq]["type"] = typ
            catalog[category][eq]["cert_no"] = cert
            catalog[category][eq]["status"] = "done"

            if file and file.filename != "":
                safe = secure_filename(file.filename)
                key = f"uploads/edit/{ship_number}_{category}_{eq}_{safe}"
                s3 = s3_client()
                s3.upload_fileobj(file, S3_BUCKET, key, ExtraArgs={"ContentType": file.mimetype})
                catalog[category][eq]["file"] = safe
                catalog[category][eq]["file_url"] = presigned_url(key)

        put_json_to_file(catalog_path, catalog)
        flash("장비 정보가 업데이트되었습니다.")
        return redirect(url_for("home", ship_number=ship_number))

    info = catalog.get(category, {}).get(eq, {})
    return render_template(
        "edit.html",
        ship_number=ship_number, category=category, eq=eq, info=info
    )

# ================== Admin 기능 ==================
if ADMIN_ENABLED:
    @app.route("/admin")
    def admin_dashboard():
        submissions = []
        s3 = s3_client()
        try:
            resp = s3.list_objects_v2(Bucket=S3_BUCKET, Prefix="submissions/")
            if "Contents" in resp:
                for obj in resp["Contents"]:
                    if not obj["Key"].endswith(".json"):
                        continue
                    data = s3.get_object(Bucket=S3_BUCKET, Key=obj["Key"])["Body"].read()
                    rows = json.loads(data)
                    submissions.extend(rows)
        except Exception as e:
            print("[ERROR] Admin load failed:", e)

        return render_template("admin.html", submissions=submissions)

    # 📌 엑셀 내보내기
    @app.route("/export/excel")
    def export_excel():
        wb = Workbook()
        ws = wb.active
        ws.title = "Submissions"
        ws.append([
            "id", "ship_number", "category", "equipment_name", "qty",
            "maker", "type", "cert_no", "status", "responsible", "submitter_name"
        ])

        s3 = s3_client()
        submissions = []
        try:
            resp = s3.list_objects_v2(Bucket=S3_BUCKET, Prefix="submissions/")
            if "Contents" in resp:
                for obj in resp["Contents"]:
                    if obj["Key"].endswith(".json"):
                        data = s3.get_object(Bucket=S3_BUCKET, Key=obj["Key"])["Body"].read()
                        rows = json.loads(data)
                        submissions.extend(rows)
        except Exception as e:
            print("[ERROR] Excel export failed:", e)

        for s in submissions:
            ws.append([
                s.get("id",""),
                s.get("ship_number",""),
                s.get("category",""),
                s.get("equipment_name",""),
                s.get("qty",""),
                s.get("maker",""),
                s.get("type",""),
                s.get("cert_no",""),
                s.get("status",""),
                s.get("responsible",{}).get("name","") if s.get("responsible") else "",
                s.get("submitter_name","")
            ])

        stream = io.BytesIO()
        wb.save(stream); stream.seek(0)
        return send_file(
            stream,
            as_attachment=True,
            download_name="submissions.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # 📌 수동 메일 발송 버튼용 (임시 처리)
    @app.route("/manual_mail", methods=["POST"])
    def manual_mail():
        flash("📧 수동 메일 발송 기능은 아직 구현되지 않았습니다.")
        return redirect(url_for("admin_dashboard"))

# ================= Health Check =================
@app.route("/health")
def health():
    return "ok", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
