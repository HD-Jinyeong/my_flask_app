import os, io, json, uuid, datetime, random, smtplib, time
from urllib.parse import quote
from email.mime.text import MIMEText
from email.utils import formataddr
from werkzeug.utils import secure_filename
from flask import Flask, request, render_template, redirect, url_for, send_file, flash, jsonify, abort
import boto3
from botocore.config import Config  # timeout/retry 설정
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "dev-only-change-me")

# ================ 환경변수 ================
S3_BUCKET = os.getenv("S3_BUCKET")
S3_REGION = os.getenv("S3_REGION", "ap-northeast-2")
AWS_ACCESS_KEY_ID = os.getenv("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = os.getenv("AWS_SECRET_ACCESS_KEY")

# 메일(단순 SMTP) — no-reply 하나로 통일
SMTP_SERVER = os.getenv("SMTP_SERVER", "211.193.193.12")
SMTP_SENDER = os.getenv("SMTP_SENDER", "no-reply@hd.com")
SMTP_FROM_NAME = os.getenv("SMTP_FROM_NAME", "HD Notification")

# 관리자 페이지 on/off
ADMIN_ENABLED = os.getenv("ADMIN_ENABLED", "true").lower() == "true"

# ---------- S3 키 Prefix ----------
CATALOG_PREFIX = os.getenv("CATALOG_PREFIX", "catalog/").rstrip("/") + "/"
CONTACTS_KEY   = CATALOG_PREFIX + "contacts/contacts.json"
ACTIVITY_LOG_KEY = CATALOG_PREFIX + "logs/activity.jsonl"
MAIL_ARCHIVE_PREFIX = CATALOG_PREFIX + "mails/"

# 카탈로그 자동 생성
AUTO_CREATE_CATALOG = os.getenv("AUTO_CREATE_CATALOG", "true").lower() == "true"

# 진단/부트스트랩용 토큰 (선택)
BOOT_TOKEN = os.getenv("BOOT_TOKEN", "")

# ---------- boto3 공통 Config ----------
_BOTO_CONFIG = Config(
    region_name=S3_REGION,
    retries={"max_attempts": 3, "mode": "standard"},
    signature_version="s3v4",
    connect_timeout=5,
    read_timeout=10,
)

def s3_client():
    return boto3.client(
        "s3",
        aws_access_key_id=AWS_ACCESS_KEY_ID,
        aws_secret_access_key=AWS_SECRET_ACCESS_KEY,
        region_name=S3_REGION,
        config=_BOTO_CONFIG
    )

def sts_client():
    return boto3.client(
        "sts",
        aws_access_key_id=AWS_ACCESS_KEY_ID,
        aws_secret_access_key=AWS_SECRET_ACCESS_KEY,
        region_name=S3_REGION,
        config=_BOTO_CONFIG
    )

def presigned_url(key, expires=3600*24*7):
    s3 = s3_client()
    return s3.generate_presigned_url(
        "get_object",
        Params={"Bucket": S3_BUCKET, "Key": key},
        ExpiresIn=expires
    )

# ================ 공통 유틸 ================
def s3_get_json(key, default=None):
    s3 = s3_client()
    try:
        obj = s3.get_object(Bucket=S3_BUCKET, Key=key)
        return json.loads(obj["Body"].read().decode("utf-8"))
    except Exception:
        return default

def s3_put_json(key, data):
    s3 = s3_client()
    try:
        s3.put_object(
            Bucket=S3_BUCKET,
            Key=key,
            Body=json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8"),
            ContentType="application/json",
            CacheControl="no-cache, no-store, must-revalidate"
        )
    except Exception as e:
        print(f"[ERROR] s3_put_json failed key={key}: {e}")
        raise

def append_activity_log(event: dict):
    s3 = s3_client()
    try:
        try:
            obj = s3.get_object(Bucket=S3_BUCKET, Key=ACTIVITY_LOG_KEY)
            old = obj["Body"].read()
        except Exception:
            old = b""
        line = (json.dumps(event, ensure_ascii=False) + "\n").encode("utf-8")
        s3.put_object(
            Bucket=S3_BUCKET,
            Key=ACTIVITY_LOG_KEY,
            Body=old + line,
            ContentType="application/json",
            CacheControl="no-cache, no-store, must-revalidate"
        )
    except Exception as e:
        print("[WARN] activity log append failed:", e)

def get_contacts():
    return s3_get_json(CONTACTS_KEY, default={"list": []})

# ---------- 연락처 정규화 ----------
def _normalize_contact(name, email, phone):
    name  = (name or "").strip()
    email = (email or "").strip().lower()
    phone = (phone or "").strip()
    return name, email, phone

# ================== 담당자 DB(요청대로 갱신) ==================
responsibles = [
    {"name": "최현서", "email": "jinyeong@hd.com",      "phone": "010-0000-0000"},
    {"name": "하태현", "email": "wlsdud5706@naver.com", "phone": "010-0000-0000"},
    {"name": "전민수", "email": "wlsdud5706@knu.ac.kr", "phone": "010-0000-0000"}
]

# ★ 기존 DB(contacts.json & 각 catalog)의 통일된 메일을 정리/치환하기 위한 매핑
RESP_EMAIL_OVERRIDE = {
    "최현서": "jinyeong@hd.com",
    "하태현": "wlsdud5706@naver.com",
    "전민수": "wlsdud5706@knu.ac.kr",
}
RESP_PHONE_OVERRIDE = {
    "최현서": "010-0000-0000",
    "하태현": "010-0000-0000",
    "전민수": "010-0000-0000",
}
OLD_UNIFIED_EMAIL = "jinyeong@hd.com"  # 과거 임시 통일 메일

# ---------- 연락처 중복 방지(개선판: 이름 기준으로 1건만 유지, 표준 이메일 우선) ----------
def dedupe_contacts():
    data = get_contacts()
    lst = data.get("list", [])
    if not lst:
        return
    # 이름 그룹핑
    by_name = {}
    for c in lst:
        n, e, p = _normalize_contact(c.get("name"), c.get("email"), c.get("phone"))
        if not n and not e:
            continue
        if n not in by_name:
            by_name[n] = []
        by_name[n].append({"name": n, "email": e, "phone": p})

    result = []
    for name, items in by_name.items():
        # 1) 표준 이메일(있으면) 선택
        preferred_email = RESP_EMAIL_OVERRIDE.get(name, "").strip().lower()
        chosen = None
        if preferred_email:
            for it in items:
                if it["email"] == preferred_email:
                    chosen = it
                    break
        # 2) 없다면 이메일이 있는 것 중 하나 선택
        if not chosen:
            with_email = [it for it in items if it["email"]]
            if with_email:
                chosen = with_email[0]
        # 3) 그래도 없으면(모두 이메일 없음) 첫 번째
        if not chosen:
            chosen = items[0]

        # 전화번호 보강(비어있으면 다른 항목의 번호로 보강)
        if not chosen.get("phone"):
            for it in items:
                if it.get("phone"):
                    chosen["phone"] = it["phone"]
                    break

        result.append(chosen)

    s3_put_json(CONTACTS_KEY, {"list": result})

def upsert_contact(name, email, phone):
    name, email, phone = _normalize_contact(name, email, phone)
    if not name and not email:
        return
    contacts = get_contacts()
    if "list" not in contacts:
        contacts["list"] = []
    # 동일 이름의 기존 항목을 찾아 업데이트(이메일 없던 항목도 병합)
    updated = False
    for c in contacts["list"]:
        cn, ce, cp = _normalize_contact(c.get("name"), c.get("email"), c.get("phone"))
        if cn == name:
            # 표준 이메일이 있으면 그걸로, 없으면 새 이메일이 있으면 덮어쓰기
            pref = RESP_EMAIL_OVERRIDE.get(name, "").strip().lower()
            if pref:
                c["email"] = pref
            elif email:
                c["email"] = email
            # 전화번호 보강
            if phone:
                c["phone"] = phone
            c["name"] = name
            updated = True
            break
    if not updated:
        # 신규 추가 시도
        contacts["list"].append({"name": name, "email": email, "phone": phone})
    s3_put_json(CONTACTS_KEY, contacts)
    # 마지막에 강제 dedupe(이름당 1건)
    dedupe_contacts()

def seed_contacts():
    try:
        for r in responsibles:
            upsert_contact(r.get("name"), r.get("email"), r.get("phone"))
    except Exception as e:
        print("[WARN] seed_contacts failed:", e)

def cleanup_contacts_unified_email():
    """
    contacts.json 에서 과거 통일 메일(OLD_UNIFIED_EMAIL)을
    '최현서' 외 이름에는 제거하고, dedupe로 재정렬.
    """
    data = get_contacts()
    lst = data.get("list", [])
    cleaned = []
    for c in lst:
        name, email, phone = _normalize_contact(c.get("name"), c.get("email"), c.get("phone"))
        if email == OLD_UNIFIED_EMAIL and name != "최현서":
            # 제거(이후 upsert+dedupe로 표준 이메일로 병합)
            continue
        cleaned.append({"name": name, "email": email, "phone": phone})
    s3_put_json(CONTACTS_KEY, {"list": cleaned})
    # 표준 담당자 3명 업서트 → dedupe로 이름 기준 1건 유지
    for name, new_email in RESP_EMAIL_OVERRIDE.items():
        upsert_contact(name, new_email, RESP_PHONE_OVERRIDE.get(name, ""))

def update_catalog_responsibles():
    """
    모든 ship 카탈로그에서 responsible의 name 기준으로 email/phone을 새 값으로 업데이트.
    """
    s3 = s3_client()
    try:
        resp = s3.list_objects_v2(Bucket=S3_BUCKET, Prefix=CATALOG_PREFIX)
        if "Contents" not in resp:
            return
        for obj in resp["Contents"]:
            k = obj["Key"]
            if not k.endswith(".json"):
                continue
            if "/contacts/" in k or "/logs/" in k or "/mails/" in k:
                continue
            try:
                raw = s3.get_object(Bucket=S3_BUCKET, Key=k)["Body"].read()
                catalog = json.loads(raw.decode("utf-8"))
            except Exception as e:
                print(f"[WARN] catalog load failed: {k} - {e}")
                continue

            changed = False
            for category, eqs in catalog.items():
                for eq_name, info in eqs.items():
                    resp_info = info.get("responsible") or {}
                    name = (resp_info.get("name") or "").strip()
                    if not name:
                        continue
                    new_email = RESP_EMAIL_OVERRIDE.get(name)
                    new_phone = RESP_PHONE_OVERRIDE.get(name)
                    cur_email = (resp_info.get("email") or "").strip().lower()
                    cur_phone = (resp_info.get("phone") or "").strip()
                    if new_email and cur_email != new_email:
                        resp_info["email"] = new_email
                        changed = True
                    if new_phone and cur_phone != new_phone:
                        resp_info["phone"] = new_phone
                        changed = True
                    info["responsible"] = resp_info
                    info["status"] = _recompute_status(info)
            if changed:
                s3_put_json(k, catalog)
    except Exception as e:
        print("[WARN] update_catalog_responsibles failed:", e)

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

# ================= Catalog 유틸 =================
def _catalog_key(ship_number): return f"{CATALOG_PREFIX}equipment_catalog_{ship_number}.json"

def load_catalog(ship_number):
    key = _catalog_key(ship_number)
    return s3_get_json(key, default={})

def create_catalog(ship_number):
    catalog = {}
    total_count = 15
    count = 0
    for category, equipments in CATALOG_EQUIPMENTS.items():
        catalog[category] = {}
        for eq in equipments:
            if count >= total_count: break
            resp = random.choice(responsibles)
            upsert_contact(resp.get("name"), resp.get("email"), resp.get("phone"))
            catalog[category][eq] = {
                "qty": "",
                "maker": "",
                "type": "",
                "cert_no": "",
                "responsible": resp,
                "status": "pending",
                "file": "",
                "file_url": "",
                "file_key": "",
                "submitter_name": "",
                "last_modified": ""
            }
            count += 1
        if count >= total_count: break
    save_catalog(ship_number, catalog)
    return catalog

def get_or_create_catalog(ship_number, force_reset=False):
    if force_reset:
        return create_catalog(ship_number)
    existing = load_catalog(ship_number)
    if existing or not AUTO_CREATE_CATALOG:
        return existing
    return create_catalog(ship_number)

def save_catalog(ship_number, catalog):
    key = _catalog_key(ship_number)
    s3_put_json(key, catalog)

def list_all_submissions():
    submissions = []
    s3 = s3_client()
    try:
        resp = s3.list_objects_v2(Bucket=S3_BUCKET, Prefix=CATALOG_PREFIX)
        if "Contents" in resp:
            for obj in resp["Contents"]:
                k = obj["Key"]
                if not k.endswith(".json"): continue
                if "/contacts/" in k or "/logs/" in k or "/mails/" in k: continue
                data = s3.get_object(Bucket=S3_BUCKET, Key=k)["Body"].read()
                catalog = json.loads(data)
                ship_number = k.split("_")[-1].split(".")[0]
                for category, eqs in catalog.items():
                    for eq_name, eq_info in eqs.items():
                        submissions.append({
                            "ship_number": ship_number,
                            "category": category,
                            "equipment_name": eq_name,
                            "qty": eq_info.get("qty", ""),
                            "maker": eq_info.get("maker", ""),
                            "type": eq_info.get("type", ""),
                            "cert_no": eq_info.get("cert_no", ""),
                            "status": _recompute_status(eq_info),
                            "responsible": eq_info.get("responsible"),
                            "submitter_name": eq_info.get("submitter_name", ""),
                            "file": eq_info.get("file", ""),
                            "file_url": eq_info.get("file_url", ""),
                            "file_key": eq_info.get("file_key", ""),
                            "last_modified": eq_info.get("last_modified", ""),
                            "due_date": SHIP_DUE_DATES.get(ship_number, "")
                        })
    except Exception as e:
        print("[ERROR] list_all_submissions failed:", e)
    return submissions

# 파일 보기 - presigned 리다이렉트
@app.route("/file/<path:key>")
def file_redirect(key):
    if not key:
        abort(404)
    url = presigned_url(key)
    return redirect(url, code=302)

# ================= 응답 캐시 방지 =================
@app.after_request
def add_no_cache_headers(resp):
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp

# ================= 홈 =================
@app.route("/", methods=["GET"])
def home():
    ship_number = request.args.get("ship_number")
    catalog = {}
    due_date = None
    if ship_number:
        catalog = load_catalog(ship_number)
        due_date = SHIP_DUE_DATES.get(ship_number)
        if not catalog and AUTO_CREATE_CATALOG:
            catalog = create_catalog(ship_number)
    return render_template("home.html", catalog=catalog, selected_ship=ship_number, due_date=due_date)

# --------- 완료 판정 유틸 ----------
def _recompute_status(item: dict) -> str:
    fields = ["qty", "maker", "type", "cert_no"]
    filled = all((item.get(k) or "").strip() for k in fields)
    return "done" if filled else "pending"

# ================= 장비 수정 =================
@app.route("/edit/<ship_number>/<category>/<eq>", methods=["GET", "POST"])
def edit(ship_number, category, eq):
    catalog = get_or_create_catalog(ship_number)
    if request.method == "POST":
        qty   = request.form.get("qty")
        maker = request.form.get("maker")
        typ   = request.form.get("type")
        cert  = request.form.get("cert_no")

        resp_name  = request.form.get("resp_name", "").strip()
        resp_email = (request.form.get("resp_email") or "").strip().lower()
        resp_phone = request.form.get("resp_phone", "").strip()

        file  = request.files.get("file")
        submitter_name = request.form.get("submitter_name", "").strip()

        if category in catalog and eq in catalog[category]:
            item = catalog[category][eq]
            if qty is not None:   item["qty"] = qty
            if maker is not None: item["maker"] = maker
            if typ is not None:   item["type"] = typ
            if cert is not None:  item["cert_no"] = cert

            if resp_name or resp_email or resp_phone:
                item["responsible"] = {
                    "name":  resp_name  or item.get("responsible", {}).get("name", ""),
                    "email": resp_email or (item.get("responsible", {}).get("email", "")).strip().lower(),
                    "phone": resp_phone or item.get("responsible", {}).get("phone", "")
                }
                upsert_contact(item["responsible"]["name"], item["responsible"]["email"], item["responsible"]["phone"])

            item["last_modified"] = datetime.datetime.now().isoformat()

            if file and file.filename != "":
                s3 = s3_client()
                safe = secure_filename(file.filename)
                key_file = f"{CATALOG_PREFIX}uploads/edit/{ship_number}_{secure_filename(category)}_{secure_filename(eq)}_{int(datetime.datetime.now().timestamp())}_{safe}"
                s3.upload_fileobj(
                    file, S3_BUCKET, key_file,
                    ExtraArgs={"ContentType": file.mimetype, "CacheControl": "no-cache"}
                )
                item["file"] = safe
                item["file_key"] = key_file
                item["file_url"] = ""

            item["status"] = _recompute_status(item)

        save_catalog(ship_number, catalog)
        append_activity_log({
            "ts": datetime.datetime.now().isoformat(),
            "actor": submitter_name or "admin_or_user",
            "action": "edit",
            "ship": ship_number, "category": category, "equipment": eq,
            "source": "edit_route"
        })

        next_url = request.args.get("next") or request.form.get("next")
        if next_url and next_url.startswith("/"):
            if "?" in next_url:
                next_url = f"{next_url}&_={int(time.time())}"
            else:
                next_url = f"{next_url}?_={int(time.time())}"
            return redirect(next_url)
        return redirect(url_for("home", ship_number=ship_number, _=int(time.time())))

    info = catalog.get(category, {}).get(eq, {})
    return render_template("edit.html", ship_number=ship_number, category=category, eq=eq, info=info)

# ================== Admin 기능 ==================
def _is_incomplete(item: dict) -> bool:
    return _recompute_status(item) != "done"

def _build_missing_report(ship_number: str, catalog: dict):
    lines = []
    to_emails = set()
    total = 0
    by_category = {}
    for category, eqs in catalog.items():
        cat_list = []
        for eq, info in eqs.items():
            if _is_incomplete(info):
                total += 1
                resp = info.get("responsible") or {}
                e = (resp.get("email") or "").strip().lower()
                if e:
                    to_emails.add(e)
                cat_list.append(eq)
        if cat_list:
            by_category[category] = cat_list
            lines.append(f"[{category}]\n" + "\n".join(f"- {x}" for x in cat_list))
    due = SHIP_DUE_DATES.get(ship_number, "")
    body = f"""안녕하세요,

아래 항목들이 {due} 까지 입력되지 않아 안내드립니다.
카테고리별 미입력 장비 리스트는 다음과 같습니다:

{('\n\n'.join(lines)) if lines else '- (없음)'}

번거로우시겠지만 기한 내 입력 부탁드립니다.

감사합니다.
"""
    return sorted(to_emails), body, total, by_category

def send_email_via_smtp(to_emails, cc_emails, subject, body_text):
    from_addr = SMTP_SENDER
    from_name = SMTP_FROM_NAME
    msg = MIMEText(body_text + "\n\n※ 본 메일은 회신 수신되지 않습니다(no-reply).", _charset="utf-8")
    msg["From"] = formataddr((from_name, from_addr))
    if to_emails: msg["To"] = ", ".join(to_emails)
    if cc_emails: msg["Cc"] = ", ".join(cc_emails)
    msg["Subject"] = subject
    recipients = list(dict.fromkeys([*(to_emails or []), *(cc_emails or [])]))
    with smtplib.SMTP(SMTP_SERVER) as server:
        server.sendmail(from_addr, recipients, msg.as_string())

if ADMIN_ENABLED:
    @app.route("/admin")
    def admin_dashboard():
        dedupe_contacts()
        submissions = list_all_submissions()
        contacts = get_contacts()
        ships = sorted({s["ship_number"] for s in submissions})
        incomplete_count = {sh: 0 for sh in ships}
        for sh in ships:
            catalog = load_catalog(sh)
            _, _, cnt, _ = _build_missing_report(sh, catalog if catalog else {})
            incomplete_count[sh] = cnt
        logs = []
        try:
            obj = s3_client().get_object(Bucket=S3_BUCKET, Key=ACTIVITY_LOG_KEY)
            lines = obj["Body"].read().decode("utf-8").strip().splitlines()[-50:]
            for ln in lines:
                try: logs.append(json.loads(ln))
                except Exception: pass
        except Exception: pass
        return render_template(
            "admin.html",
            submissions=submissions,
            contacts=contacts.get("list", []),
            logs=logs,
            ships=ships,
            incomplete_count=incomplete_count,
            SHIP_DUE_DATES=SHIP_DUE_DATES
        )

    @app.route("/admin/ship_mail/<ship_number>", methods=["POST"])
    def send_ship_mail(ship_number):
        catalog = load_catalog(ship_number)
        to_emails, body_text, missing_cnt, by_category = _build_missing_report(ship_number, catalog if catalog else {})
        cc_emails = request.form.getlist("cc_emails")

        if missing_cnt == 0:
            msg = f"Ship {ship_number}: 미입력 항목이 없습니다. 메일을 보내지 않았습니다."
            if request.headers.get("X-Requested-With") == "fetch":
                return jsonify({"ok": False, "message": msg, "missing": 0, "to": [], "cc": cc_emails}), 400
            flash(msg); return redirect(url_for("admin_dashboard", _=int(time.time())))

        if not to_emails:
            msg = f"Ship {ship_number}: 미입력 항목은 있으나 책임자 이메일이 없습니다."
            if request.headers.get("X-Requested-With") == "fetch":
                return jsonify({"ok": False, "message": msg, "missing": missing_cnt, "to": [], "cc": cc_emails}), 400
            flash(msg); return redirect(url_for("admin_dashboard", _=int(time.time())))

        subject = f"[Ship {ship_number}] 미입력 항목 안내 ({SHIP_DUE_DATES.get(ship_number, '')})"
        sent = False; err = None
        try:
            send_email_via_smtp(to_emails, cc_emails, subject, body_text)
            sent = True
        except Exception as e:
            err = str(e); print("[ERROR] SMTP send_email failed:", e)

        archive = {
            "ts": datetime.datetime.now().isoformat(),
            "ship": ship_number,
            "to": to_emails, "cc": cc_emails,
            "subject": subject, "body": body_text,
            "sent": sent, "method": "smtp", "error": err,
            "missing_count": missing_cnt, "by_category": by_category
        }
        s3_put_json(f"{MAIL_ARCHIVE_PREFIX}{ship_number}_bulk_{int(datetime.datetime.now().timestamp())}.json", archive)
        append_activity_log({
            "ts": datetime.datetime.now().isoformat(),
            "actor": "admin",
            "action": "mail_bulk_send",
            "ship": ship_number,
            "result": "ok" if sent else f"fail:{err}"
        })

        if request.headers.get("X-Requested-With") == "fetch":
            return jsonify({"ok": sent, "message": ("전송 완료" if sent else f"전송 실패: {err}"),
                            "missing": missing_cnt, "to": to_emails, "cc": cc_emails}), (200 if sent else 500)
        flash(f"Ship {ship_number}: {'메일 전송 완료' if sent else '메일 전송 실패 - ' + (err or '')}")
        return redirect(url_for("admin_dashboard", _=int(time.time())))

    @app.route("/admin/export_selected", methods=["POST"])
    def export_selected():
        rows = request.form.getlist("rows[]")
        all_items = {(f"{s['ship_number']}|{s['category']}|{s['equipment_name']}"): s
                     for s in list_all_submissions()}
        picked = [all_items[r] for r in rows if r in all_items]
        if not picked:
            flash("선택된 항목이 없습니다.")
            return redirect(url_for("admin_dashboard", _=int(time.time())))
        wb = Workbook(); ws = wb.active; ws.title = "Selected"
        ws.append(["Ship","Category","Equipment","QTY","Maker","Type","Cert No.","Status","Responsible(Name)","Responsible(Email)","Phone","Submitter","Last Modified","Due Date","File"])
        for it in picked:
            resp = it.get("responsible") or {}
            link = ""
            if it.get("file_key"): link = url_for("file_redirect", key=it["file_key"], _external=True)
            elif it.get("file_url"): link = it.get("file_url")
            ws.append([it["ship_number"], it["category"], it["equipment_name"],
                       it.get("qty",""), it.get("maker",""), it.get("type",""),
                       it.get("cert_no",""), it.get("status",""),
                       resp.get("name",""), resp.get("email",""), resp.get("phone",""),
                       it.get("submitter_name",""), it.get("last_modified",""),
                       it.get("due_date",""), link])
        bio = io.BytesIO(); wb.save(bio); bio.seek(0)
        filename = f"selected_export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(bio, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    @app.route("/export/excel")
    def export_excel():
        items = list_all_submissions()
        wb = Workbook(); ws = wb.active; ws.title = "All"
        ws.append(["Ship","Category","Equipment","QTY","Maker","Type","Cert No.","Status","Responsible(Name)","Responsible(Email)","Phone","Submitter","Last Modified","Due Date","File"])
        for it in items:
            resp = it.get("responsible") or {}
            link = ""
            if it.get("file_key"): link = url_for("file_redirect", key=it["file_key"], _external=True)
            elif it.get("file_url"): link = it.get("file_url")
            ws.append([it["ship_number"], it["category"], it["equipment_name"],
                       it.get("qty",""), it.get("maker",""), it.get("type",""),
                       it.get("cert_no",""), it.get("status",""),
                       resp.get("name",""), resp.get("email",""), resp.get("phone",""),
                       it.get("submitter_name",""), it.get("last_modified",""),
                       it.get("due_date",""), link])
        bio = io.BytesIO(); wb.save(bio); bio.seek(0)
        filename = f"export_all_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(bio, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------- 🔎 S3/자격증명 진단용(선택, 토큰 보호) ----------
def _require_token():
    t = request.args.get("token") or request.headers.get("X-Diag-Token") or ""
    if not BOOT_TOKEN or t != BOOT_TOKEN:
        return False
    return True

@app.route("/diag/aws")
def diag_aws():
    if not _require_token():
        return jsonify({"ok": False, "error": "unauthorized"}), 401
    info = {"ok": True, "region": S3_REGION}
    try:
        me = sts_client().get_caller_identity()
        info["identity"] = {
            "Account": me.get("Account"),
            "Arn": me.get("Arn"),
            "UserId": me.get("UserId")
        }
    except Exception as e:
        info["identity_error"] = str(e)
    info["env"] = {
        "S3_BUCKET": S3_BUCKET,
        "CATALOG_PREFIX": CATALOG_PREFIX,
        "AUTO_CREATE_CATALOG": AUTO_CREATE_CATALOG,
        "ADMIN_ENABLED": ADMIN_ENABLED
    }
    return jsonify(info)

@app.route("/diag/s3")
def diag_s3():
    if not _require_token():
        return jsonify({"ok": False, "error": "unauthorized"}), 401
    s3 = s3_client()
    out = {"ok": True, "bucket": S3_BUCKET, "prefix": CATALOG_PREFIX}
    try:
        resp = s3.list_objects_v2(Bucket=S3_BUCKET, Prefix=CATALOG_PREFIX, MaxKeys=5)
        out["list_sample"] = [c["Key"] for c in resp.get("Contents", [])]
    except Exception as e:
        out["list_error"] = str(e)
    return jsonify(out)

@app.route("/diag/s3/key")
def diag_s3_key():
    if not _require_token():
        return jsonify({"ok": False, "error": "unauthorized"}), 401
    key = request.args.get("key")
    if not key:
        return jsonify({"ok": False, "error": "key required"}), 400
    s3 = s3_client()
    out = {"ok": True, "key": key}
    try:
        h = s3.head_object(Bucket=S3_BUCKET, Key=key)
        out["exists"] = True
        out["etag"] = h.get("ETag")
        out["size"] = h.get("ContentLength")
    except Exception as e:
        out["exists"] = False
        out["head_error"] = str(e)
    return jsonify(out)

@app.route("/diag/s3/put_test", methods=["POST"])
def diag_s3_put():
    if not _require_token():
        return jsonify({"ok": False, "error": "unauthorized"}), 401
    key = f"{CATALOG_PREFIX}__ping__/{int(time.time())}.txt"
    s3 = s3_client()
    try:
        s3.put_object(Bucket=S3_BUCKET, Key=key, Body=b"ping", ContentType="text/plain", CacheControl="no-cache")
        obj = s3.get_object(Bucket=S3_BUCKET, Key=key)
        body = obj["Body"].read().decode("utf-8")
        return jsonify({"ok": True, "key": key, "readback": body})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e), "key": key}), 500

# ---------- 🔑 Admin 없이도 1회 초기화 가능한 부트스트랩(선택) ----------
@app.route("/catalog/bootstrap/<ship_number>", methods=["POST"])
def catalog_bootstrap(ship_number):
    token = request.args.get("token") or request.headers.get("X-Boot-Token") or ""
    if not BOOT_TOKEN or token != BOOT_TOKEN:
        return jsonify({"ok": False, "error": "unauthorized"}), 401
    existed = bool(load_catalog(ship_number))
    cat = create_catalog(ship_number)
    return jsonify({"ok": True, "existed_before": existed, "created": True, "key": _catalog_key(ship_number)}), 200

# ================= Health Check =================
@app.route("/health")
def health():
    return "ok", 200

if __name__ == "__main__":
    def _s3_ready():
        return bool(S3_BUCKET and AWS_ACCESS_KEY_ID and AWS_SECRET_ACCESS_KEY and S3_REGION)

    if _s3_ready():
        seed_contacts()
        cleanup_contacts_unified_email()
        update_catalog_responsibles()
        dedupe_contacts()
    else:
        print("[WARN] S3 env not set or partial. Skipping contacts/catalog cleanup.")

    print("[BOOT] S3_BUCKET=", S3_BUCKET, " S3_REGION=", S3_REGION, " PREFIX=", CATALOG_PREFIX, " AUTO_CREATE_CATALOG=", AUTO_CREATE_CATALOG, " ADMIN_ENABLED=", ADMIN_ENABLED)
    app.run(host="0.0.0.0", port=5000, debug=True)
