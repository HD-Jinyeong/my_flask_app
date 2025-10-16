import os, io, json, uuid, datetime, random, smtplib, time
from urllib.parse import quote
from email.mime.text import MIMEText
from email.utils import formataddr
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from flask import Flask, request, render_template, redirect, url_for, send_file, flash, jsonify, abort, session
import boto3
from botocore.config import Config  # timeout/retry 설정
from openpyxl import Workbook
from functools import wraps
from botocore.exceptions import ClientError

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "dev-only-change-me")

# ================ 환경변수 ================
S3_BUCKET = os.getenv("S3_BUCKET")
S3_REGION = os.getenv("S3_REGION", "ap-northeast-2")
AWS_ACCESS_KEY_ID = os.getenv("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = os.getenv("AWS_SECRET_ACCESS_KEY")

# 메일(단순 SMTP)
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
USERS_KEY = CATALOG_PREFIX + "auth/users.json"
INVITES_KEY = CATALOG_PREFIX + "auth/invites.json"
# ✅ 추가: 메일 이벤트 로그 저장 경로(prefix)
MAIL_LOG_PREFIX = CATALOG_PREFIX + "logs/mail/"

# 카탈로그 자동 생성
AUTO_CREATE_CATALOG = os.getenv("AUTO_CREATE_CATALOG", "true").lower() == "true"

# QTY 자동 모드 (읽기만 하고 사용하지 않음)
AUTO_QTY_ENABLED = os.getenv("AUTO_QTY_ENABLED", "true").lower() == "true"

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

# === First-Visit Guard ===
@app.before_request
def _first_visit_guard():
    if session.get("user"):
        return None
    if session.get("first_visit_done"):
        return None
    exempt = {"login", "auth_complete", "static", "health", "file_redirect", "file_inline"}
    if request.endpoint in exempt or (request.path or "").startswith("/static/"):
        return None
    next_path = request.full_path if request.query_string else request.path
    return redirect(url_for("login", next=next_path))

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

# ✅ 추가: 리스트 JSON 전용 get/put + 메일 이벤트 로그 유틸
def _s3_get_json_list(key):
    try:
        obj = s3_client().get_object(Bucket=S3_BUCKET, Key=key)
        body = obj["Body"].read().decode("utf-8")
        data = json.loads(body)
        return data if isinstance(data, list) else []
    except ClientError as e:
        if e.response.get("Error", {}).get("Code") in ("NoSuchKey", "404"):
            return []
        raise
    except Exception:
        return []

def _s3_put_json_list(key, data_list):
    s3_client().put_object(
        Bucket=S3_BUCKET,
        Key=key,
        Body=json.dumps(data_list, ensure_ascii=False).encode("utf-8"),
        ContentType="application/json; charset=utf-8",
        CacheControl="no-cache, no-store, must-revalidate"
    )

def log_mail_event(ship: str, category: str, action: str, result: str, purpose: str = None, extra: dict = None):
    """
    메일 전송 이벤트를 S3에 append
    파일: catalog/logs/mail/{ship}/{category}.json
    항목: { ts, action, result, meta:{purpose, ...} }
    """
    if not (ship and category):
        return False
    key = f"{MAIL_LOG_PREFIX}{ship}/{category}.json"
    logs = _s3_get_json_list(key)
    item = {
        "ts": datetime.datetime.utcnow().isoformat(timespec="seconds") + "Z",
        "action": action,
        "result": result,
        "meta": {}
    }
    if purpose:
        item["meta"]["purpose"] = purpose
    if extra and isinstance(extra, dict):
        item["meta"].update(extra)
    logs.append(item)
    _s3_put_json_list(key, logs)
    return True

def read_mail_logs_grouped(owners_by_ship: dict):
    """
    owners_by_ship[ship] = { category: [owners...] }
    -> logs_by_ship[ship][category] = [...]
    """
    out = {}
    for sh, cats in (owners_by_ship or {}).items():
        out[sh] = {}
        for cat in (cats or {}).keys():
            key = f"{MAIL_LOG_PREFIX}{sh}/{cat}.json"
            lst = _s3_get_json_list(key)
            lst.sort(key=lambda x: x.get("ts", ""), reverse=True)
            out[sh][cat] = lst
    return out

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

def cleanup_bad_logs():
    """category에 '&amp;'가 들어간 잘못된 로그 라인을 제거"""
    s3 = s3_client()
    try:
        obj = s3.get_object(Bucket=S3_BUCKET, Key=ACTIVITY_LOG_KEY)
        lines = obj["Body"].read().decode("utf-8").splitlines()
    except Exception:
        return
    out = []
    removed = 0
    for ln in lines:
        try:
            rec = json.loads(ln)
            cat = (rec.get("category") or "")
            if "&amp;" in cat:
                removed += 1
                continue
        except Exception:
            pass
        out.append(ln)
    if removed:
        try:
            s3.put_object(
                Bucket=S3_BUCKET,
                Key=ACTIVITY_LOG_KEY,
                Body=("\n".join(out) + ("\n" if out else "")).encode("utf-8"),
                ContentType="application/json",
                CacheControl="no-cache, no-store, must-revalidate"
            )
            print(f"[CLEANUP] removed {removed} bad log lines (&amp; in category)")
        except Exception as e:
            print("[WARN] cleanup_bad_logs failed:", e)

def get_contacts():
    return s3_get_json(CONTACTS_KEY, default={"list": []})

def _normalize_contact(name, email, phone):
    name  = (name or "").strip()
    email = (email or "").strip().lower()
    phone = (phone or "").strip()
    return name, email, phone

# ================== 담당자 DB(예시 시드) ==================
responsibles = [
    {"name": "최현서", "email": "jinyeong@hd.com",      "phone": "010-0000-0000"},
    {"name": "하태현", "email": "wlsdud5706@naver.com", "phone": "010-0000-0000"},
    {"name": "전민수", "email": "wlsdud5706@knu.ac.kr", "phone": "010-0000-0000"}
]
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
OLD_UNIFIED_EMAIL = "jinyeong@hd.com"

def dedupe_contacts():
    data = get_contacts()
    lst = data.get("list", [])
    if not lst:
        return
    by_name = {}
    for c in lst:
        n, e, p = _normalize_contact(c.get("name"), c.get("email"), c.get("phone"))
        if not n and not e:
            continue
        by_name.setdefault(n, []).append({"name": n, "email": e, "phone": p})

    result = []
    for name, items in by_name.items():
        preferred_email = RESP_EMAIL_OVERRIDE.get(name, "").strip().lower()
        chosen = None
        if preferred_email:
            for it in items:
                if it["email"] == preferred_email:
                    chosen = it; break
        if not chosen:
            with_email = [it for it in items if it["email"]]
            chosen = with_email[0] if with_email else items[0]
        if not chosen.get("phone"):
            for it in items:
                if it.get("phone"):
                    chosen["phone"] = it["phone"]; break
        result.append(chosen)
    s3_put_json(CONTACTS_KEY, {"list": result})

def upsert_contact(name, email, phone):
    name, email, phone = _normalize_contact(name, email, phone)
    if not name and not email:
        return
    contacts = get_contacts()
    contacts.setdefault("list", [])
    updated = False
    for c in contacts["list"]:
        cn, _, _ = _normalize_contact(c.get("name"), c.get("email"), c.get("phone"))
        if cn == name:
            pref = RESP_EMAIL_OVERRIDE.get(name, "").strip().lower()
            c["email"] = pref or (email or c.get("email",""))
            if phone: c["phone"] = phone
            c["name"] = name
            updated = True
            break
    if not updated:
        contacts["list"].append({"name": name, "email": email, "phone": phone})
    s3_put_json(CONTACTS_KEY, contacts)
    dedupe_contacts()

def seed_contacts():
    try:
        for r in responsibles:
            upsert_contact(r.get("name"), r.get("email"), r.get("phone"))
    except Exception as e:
        print("[WARN] seed_contacts failed:", e)

def cleanup_contacts_unified_email():
    data = get_contacts()
    lst = data.get("list", [])
    cleaned = []
    for c in lst:
        name, email, phone = _normalize_contact(c.get("name"), c.get("email"), c.get("phone"))
        if email == OLD_UNIFIED_EMAIL and name != "최현서":
            continue
        cleaned.append({"name": name, "email": email, "phone": phone})
    s3_put_json(CONTACTS_KEY, {"list": cleaned})
    for name, new_email in RESP_EMAIL_OVERRIDE.items():
        upsert_contact(name, new_email, RESP_PHONE_OVERRIDE.get(name, ""))

def update_catalog_responsibles():
    s3 = s3_client()
    try:
        resp = s3.list_objects_v2(Bucket=S3_BUCKET, Prefix=CATALOG_PREFIX)
        if "Contents" not in resp:
            return
        for obj in resp["Contents"]:
            k = obj["Key"]
            if not k.endswith(".json"): continue
            if "/contacts/" in k or "/logs/" in k or "/mails/" in k or "/auth/" in k: continue
            try:
                raw = s3.get_object(Bucket=S3_BUCKET, Key=k)["Body"].read()
                catalog = json.loads(raw.decode("utf-8"))
            except Exception as e:
                print(f"[WARN] catalog load failed: {k} - {e}")
                continue

            if not isinstance(catalog, dict):
                print(f"[WARN] skip non-dict catalog: {k}")
                continue

            changed = False
            for category, eqs in catalog.items():
                if not isinstance(eqs, dict): continue
                if "__owners__" not in eqs:
                    eqs["__owners__"] = []; changed = True
                if "__status__" not in eqs:
                    eqs["__status__"] = "미입력"; changed = True
                if "__cat_locs__" not in eqs:
                    eqs["__cat_locs__"] = []; changed = True
                if "__cat_photo_key__" not in eqs:
                    eqs["__cat_photo_key__"] = ""; changed = True
                if "__ex_proof__" not in eqs:
                    eqs["__ex_proof__"] = "Unknown"; changed = True

                for eq_name, info in eqs.items():
                    if isinstance(eq_name, str) and eq_name.startswith("__"): continue
                    if not isinstance(info, dict): continue
                    resp_info = info.get("responsible") or {}
                    if not isinstance(resp_info, dict): resp_info = {}
                    name = (resp_info.get("name") or "").strip()
                    if name:
                        new_email = RESP_EMAIL_OVERRIDE.get(name)
                        new_phone = RESP_PHONE_OVERRIDE.get(name)
                        cur_email = (resp_info.get("email") or "").strip().lower()
                        cur_phone = (resp_info.get("phone") or "").strip()
                        if new_email and cur_email != new_email:
                            resp_info["email"] = new_email; changed = True
                        if new_phone and cur_phone != new_phone:
                            resp_info["phone"] = new_phone; changed = True
                        info["responsible"] = resp_info
                    if "__deleted__" not in info:
                        info["__deleted__"] = False; changed = True
                    for k2 in ("ex_proof_grade","ip_grade","location","page"):
                        if k2 not in info:
                            info[k2] = ""; changed = True
                    info["status"] = _recompute_status(info)
            if changed:
                s3_put_json(k, catalog)
    except Exception as e:
        print("[WARN] update_catalog_responsibles failed:", e)

# ================= Ship 별 Due Date =================
SHIP_DUE_DATES = {"1": "2025-12-17", "2": "2025-12-18", "3": "2025-12-19"}

# ================= 실무 Catalog & EQ List =================
CATALOG_EQUIPMENTS = {
    "Lighting": [
        "Flood Light","Search Light","Navigation Light","Explosion-proof Light",
        "Deck Light","Emergency Light","Work Light","Pilot Lamp",
        "LED Panel Light","Area Light"
    ],
    "Switches & Junction Boxes": [
        "Explosion-proof Switch","Explosion-proof Junction Box","Limit Switch","Control Box",
        "Local Control Station","Terminal Box","Push Button Station","Selector Switch",
        "Circuit Breaker Panel","Distribution Box"
    ],
    "Motor & Machinery": [
        "Explosion-proof Motor","Fan Unit","Pump Unit","Starter Panel",
        "Gear Motor","Blower","Compressor","Hydraulic Power Unit",
        "Winch Motor","Conveyor Motor"
    ],
    "Communication & Alarm": [
        "Telephone Set","Alarm Bell","Signal Horn","Intercom",
        "Public Address Amp","Call Point","Beacon","Siren Controller",
        "Talk Back Unit","CCTV Camera"
    ],
    "Miscellaneous Equipment": [
        "Heater","Transformer","Battery Charger","Cable Gland",
        "Light Fitting","Distribution Board","Power Supply","UPS Unit",
        "Inverter","Rectifier"
    ]
}

# ================= Catalog 유틸 =================
def _catalog_key(ship_number): return f"{CATALOG_PREFIX}equipment_catalog_{ship_number}.json"

def _assign_random_category_owners(catalog: dict) -> bool:
    if not isinstance(catalog, dict):
        return False
    contacts = (get_contacts() or {}).get("list", [])
    pool = []
    for c in contacts:
        n, e, p = _normalize_contact(c.get("name"), c.get("email"), c.get("phone"))
        if e: pool.append({"name": n, "email": e, "phone": p})
    if len(pool) < 2:
        return False
    changed = False
    for category, block in catalog.items():
        if not isinstance(block, dict): continue
        if block.get("__owners__"): continue
        picked = random.sample(pool, 2 if len(pool) >= 2 else len(pool))
        block["__owners__"] = picked
        changed = True
    return changed

def load_catalog(ship_number):
    key = _catalog_key(ship_number)
    catalog = s3_get_json(key, default={})
    dirty = False
    if _assign_random_category_owners(catalog):
        dirty = True
    if dirty:
        s3_put_json(key, catalog)
    return catalog

def create_catalog(ship_number):
    catalog = {}
    for category, equipments in CATALOG_EQUIPMENTS.items():
        pick_n = random.randint(7, min(10, len(equipments)))
        picked = random.sample(equipments, pick_n)
        catalog[category] = {"__owners__": [],"__status__": "미입력","__cat_locs__": [],"__cat_photo_key__": "","__ex_proof__": "Unknown"}
        for eq in picked:
            catalog[category][eq] = {
                "qty": "", "maker": "", "type": "", "cert_no": "",
                "responsible": {}, "status": "pending",
                "file": "", "file_url": "", "file_key": "",
                "submitter_name": "", "last_modified": "",
                "photo_key": "", "locs": [],
                "ex_proof_grade": "", "ip_grade": "", "location": "", "page": "",
                "__deleted__": False
            }
    _assign_random_category_owners(catalog)
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

def _ensure_item(ship_number: str, catalog: dict, category: str, eq: str) -> bool:
    created = False
    if not isinstance(catalog, dict):
        return False
    if category not in catalog or not isinstance(catalog.get(category), dict):
        catalog[category] = {"__owners__": [],"__status__": "미입력","__cat_locs__": [],"__cat_photo_key__": "","__ex_proof__": "Unknown"}
        created = True
    if eq != "__CATEGORY__" and eq not in catalog[category]:
        catalog[category][eq] = {
            "qty": "", "maker": "", "type": "", "cert_no": "",
            "responsible": {}, "status": "pending",
            "file": "", "file_url": "", "file_key": "",
            "submitter_name": "", "last_modified": "",
            "photo_key": "", "locs": [],
            "ex_proof_grade": "", "ip_grade": "", "location": "", "page": "",
            "__deleted__": False
        }
        created = True
    if created:
        save_catalog(ship_number, catalog)
    return created

def _ensure_item_extended_fields(item: dict) -> bool:
    changed = False
    for k in ("ex_proof_grade","ip_grade","location","page"):
        if k not in item: item[k] = ""; changed = True
    if "__deleted__" not in item:
        item["__deleted__"] = False; changed = True
    return changed

def _has_any_input(eq_info: dict) -> bool:
    if not isinstance(eq_info, dict):
        return False
    fields = ["qty","maker","type","cert_no","ex_proof_grade","ip_grade","location","page","file_key","file_url","last_modified"]
    return any((eq_info.get(k) or "").strip() for k in fields)

def list_all_submissions():
    submissions = []
    s3 = s3_client()
    try:
        resp = s3.list_objects_v2(Bucket=S3_BUCKET, Prefix=CATALOG_PREFIX)
        if "Contents" in resp:
            for obj in resp["Contents"]:
                k = obj["Key"]
                if not k.endswith(".json"): continue
                if "/contacts/" in k or "/logs/" in k or "/mails/" in k or "/auth/" in k: continue
                data = s3.get_object(Bucket=S3_BUCKET, Key=k)["Body"].read()
                catalog = json.loads(data)
                ship_number = k.split("_")[-1].split(".")[0]
                if not isinstance(catalog, dict): continue
                for category, eqs in catalog.items():
                    if not isinstance(eqs, dict): continue
                    eqs.setdefault("__owners__", []); eqs.setdefault("__status__", "미입력")
                    eqs.setdefault("__cat_locs__", []); eqs.setdefault("__cat_photo_key__", "")
                    eqs.setdefault("__ex_proof__", "Unknown")
                    for eq_name, eq_info in eqs.items():
                        if isinstance(eq_name, str) and eq_name.startswith("__"): continue
                        if not isinstance(eq_info, dict): continue
                        if eq_info.get("__deleted__"): continue
                        if not _has_any_input(eq_info): continue
                        _ensure_item_extended_fields(eq_info)
                        submissions.append({
                            "ship_number": ship_number, "category": category, "equipment_name": eq_name,
                            "qty": eq_info.get("qty",""), "maker": eq_info.get("maker",""), "type": eq_info.get("type",""),
                            "cert_no": eq_info.get("cert_no",""), "status": _recompute_status(eq_info),
                            "responsible": {}, "submitter_name": eq_info.get("submitter_name",""),
                            "file": eq_info.get("file",""), "file_url": eq_info.get("file_url",""), "file_key": eq_info.get("file_key",""),
                            "last_modified": eq_info.get("last_modified",""), "due_date": SHIP_DUE_DATES.get(ship_number,""),
                            "ex_proof_grade": eq_info.get("ex_proof_grade",""), "ip_grade": eq_info.get("ip_grade",""),
                            "location": eq_info.get("location",""), "page": eq_info.get("page","")
                        })
    except Exception as e:
        print("[ERROR] list_all_submissions failed:", e)
    return submissions

def list_deleted_items():
    out = {}
    s3 = s3_client()
    try:
        resp = s3.list_objects_v2(Bucket=S3_BUCKET, Prefix=CATALOG_PREFIX)
        if "Contents" not in resp:
            return out
        for obj in resp["Contents"]:
            k = obj["Key"]
            if not k.endswith(".json"): continue
            if "/contacts/" in k or "/logs/" in k or "/mails/" in k or "/auth/" in k: continue
            data = s3.get_object(Bucket=S3_BUCKET, Key=k)["Body"].read()
            catalog = json.loads(data)
            ship_number = k.split("_")[-1].split(".")[0]
            for category, eqs in (catalog or {}).items():
                if not isinstance(eqs, dict): continue
                for eq, info in eqs.items():
                    if isinstance(eq, str) and eq.startswith("__"): continue
                    if not isinstance(info, dict): continue
                    if info.get("__deleted__"):
                        out.setdefault(ship_number, {}).setdefault(category, []).append(eq)
    except Exception as e:
        print("[ERROR] list_deleted_items failed:", e)
    return out

# 파일 보기
@app.route("/file/<path:key>")
def file_redirect(key):
    if not key: abort(404)
    url = presigned_url(key)
    return redirect(url, code=302)

def _safe_key(key: str) -> str:
    key = (key or "").strip()
    if not key.startswith(CATALOG_PREFIX):
        abort(403)
    return key

@app.route("/file_inline/<path:key>")
def file_inline(key):
    key = _safe_key(key)
    s3 = s3_client()
    try:
        obj = s3.get_object(Bucket=S3_BUCKET, Key=key)
        mime = obj.get("ContentType") or "application/octet-stream"
        data = obj["Body"].read()
        bio = io.BytesIO(data); bio.seek(0)
        resp = send_file(bio, mimetype=mime, as_attachment=False, download_name=os.path.basename(key))
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        return resp
    except Exception as e:
        print("[ERROR] file_inline failed:", e)
        abort(404)

@app.after_request
def add_no_cache_headers(resp):
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp

# ===================== 인증/계정 =====================
def _users_load():
    return s3_get_json(USERS_KEY, default={"users": []})

def _users_save(db):
    s3_put_json(USERS_KEY, db)

def _invites_load():
    return s3_get_json(INVITES_KEY, default={"invites": {}})

def _invites_save(db):
    s3_put_json(INVITES_KEY, db)

def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("user"):
            return redirect(url_for("login", next=request.path))
        return fn(*args, **kwargs)
    return wrapper

@app.route("/auth/invite", methods=["POST"])
def auth_invite():
    if not ADMIN_ENABLED: abort(404)
    email = (request.form.get("email") or "").strip().lower()
    next_url = request.form.get("next") or ""
    if not email:
        return jsonify({"ok": False, "error": "email required"}), 400
    inv = _invites_load()
    token = uuid.uuid4().hex
    inv["invites"][token] = {"email": email, "created": datetime.datetime.now().isoformat(), "next": next_url}
    _invites_save(inv)
    link = url_for("auth_complete", t=token, next=next_url, _external=True)
    subject = "[HD] 계정 생성 안내"
    body = f"다음 링크에서 비밀번호를 설정해 계정을 활성화하세요:\n\n{link}\n\n감사합니다."
    try:
        send_email_via_smtp([email], [], subject, body)
    except Exception as e:
        print("[ERROR] invite mail send failed:", e)
    return jsonify({"ok": True, "token": token, "link": link})

@app.route("/auth/complete", methods=["GET", "POST"])
def auth_complete():
    token = request.args.get("t") or request.form.get("t") or ""
    inv = _invites_load()
    info = inv["invites"].get(token)
    if not info:
        return "유효하지 않은 초대 링크입니다.", 400
    if request.method == "POST":
        pw = request.form.get("password") or ""
        next_url = request.form.get("next") or ""
        if len(pw) < 6:
            flash("비밀번호는 6자 이상이어야 합니다.")
            return render_template("auth_complete.html", email=info["email"], token=token, next=next_url)
        users = _users_load()
        users["users"] = [u for u in users.get("users", []) if u.get("email") != info["email"]]
        users["users"].append({"email": info["email"],"password_hash": generate_password_hash(pw),"created": datetime.datetime.now().isoformat(),"active": True})
        _users_save(users)
        session["user"] = {"email": info["email"]}
        del inv["invites"][token]; _invites_save(inv)
        flash("계정이 생성되었습니다.")
        if next_url and next_url.startswith("/"): return redirect(next_url)
        return redirect(url_for("home"))
    return render_template("auth_complete.html", email=info["email"], token=token, next=info.get("next",""))

@app.route("/login", methods=["GET", "POST"])
def login():
    session["first_visit_done"] = True
    if request.method == "POST":
        email = (request.form.get("email") or "").strip().lower()
        pw = (request.form.get("password") or "")
        users = _users_load().get("users", [])
        user = next((u for u in users if u.get("email")==email and u.get("active")), None)
        if user and check_password_hash(user.get("password_hash",""), pw):
            session["user"] = {"email": email}
            nxt = request.args.get("next") or url_for("home")
            return redirect(nxt)
        flash("이메일 또는 비밀번호가 올바르지 않습니다.")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.pop("user", None)
    return redirect(url_for("login"))

# ================= 홈 =================
@app.route("/", methods=["GET", "POST"])
@login_required
def home():
    if request.method == "POST":
        ship_number = request.form.get("ship_number") or request.args.get("ship_number")
        category = request.form.get("category") or request.args.get("category")
        return redirect(url_for("home", ship_number=ship_number, category=category))

    ship_number = request.args.get("ship_number")
    category = request.args.get("category")
    catalog = {}
    due_date = None
    categories = []
    eqs_in_category = {}
    cat_status = None
    owners = []
    system_ex_proof = "Unknown"
    my_items_in_system = []
    shared_items_in_system = []

    user_email = _current_user_email()

    if ship_number:
        catalog = load_catalog(ship_number)
        due_date = SHIP_DUE_DATES.get(ship_number)
        if not catalog and AUTO_CREATE_CATALOG:
            catalog = create_catalog(ship_number)
        categories = sorted(list(catalog.keys()))
        for cat in categories:
            if isinstance(catalog.get(cat), dict):
                catalog[cat].setdefault("__owners__", [])
                catalog[cat].setdefault("__status__", "미입력")
                catalog[cat].setdefault("__cat_locs__", [])
                catalog[cat].setdefault("__cat_photo_key__", "")
                catalog[cat].setdefault("__ex_proof__", "Unknown")
        if user_email:
            categories = [c for c in categories if _is_owner_of_category(catalog, c, user_email)]

    if ship_number and category and category not in categories:
        category = None

    if ship_number and category and isinstance(catalog.get(category), dict):
        cat_block = catalog.get(category)
        cat_status = cat_block.get("__status__", "미입력")
        owners = cat_block.get("__owners__", [])
        system_ex_proof = cat_block.get("__ex_proof__", "Unknown")
        eqs_in_category = {k: v for k, v in cat_block.items() if not str(k).startswith("__")}

        tmp_shared = []
        for eq_name, info in eqs_in_category.items():
            if not isinstance(info, dict) or info.get("__deleted__"):
                continue
            entered = any([
                (info.get("qty") or "").strip(),
                (info.get("maker") or "").strip(),
                (info.get("type") or "").strip(),
                (info.get("cert_no") or "").strip(),
                (info.get("ex_proof_grade") or "").strip(),
                (info.get("ip_grade") or "").strip(),
                (info.get("location") or "").strip(),
                (info.get("page") or "").strip(),
                (info.get("file_key") or "").strip(),
                (info.get("file_url") or "").strip(),
                (info.get("last_modified") or "").strip(),
            ])
            if entered:
                row = {"eq": eq_name}
                row.update({
                    "qty": info.get("qty",""), "maker": info.get("maker",""), "type": info.get("type",""),
                    "cert_no": info.get("cert_no",""), "ex_proof_grade": info.get("ex_proof_grade",""),
                    "ip_grade": info.get("ip_grade",""), "location": info.get("location",""),
                    "page": info.get("page",""), "file_key": info.get("file_key",""),
                    "file": info.get("file",""), "file_url": info.get("file_url",""),
                    "last_modified": info.get("last_modified","")
                })
                tmp_shared.append(row)
        shared_items_in_system = sorted(tmp_shared, key=lambda x: x.get("last_modified",""), reverse=True)

        if user_email:
            tmp_my = []
            for eq_name, info in eqs_in_category.items():
                if not isinstance(info, dict) or info.get("__deleted__"): continue
                submitter = (info.get("submitter_name") or "").strip().lower()
                if submitter and submitter == user_email:
                    row = {"eq": eq_name}
                    row.update({
                        "qty": info.get("qty",""), "maker": info.get("maker",""), "type": info.get("type",""),
                        "cert_no": info.get("cert_no",""), "ex_proof_grade": info.get("ex_proof_grade",""),
                        "ip_grade": info.get("ip_grade",""), "location": info.get("location",""),
                        "page": info.get("page",""), "file_key": info.get("file_key",""),
                        "file": info.get("file",""), "file_url": info.get("file_url",""),
                        "last_modified": info.get("last_modified","")
                    })
                    tmp_my.append(row)
            my_items_in_system = sorted(tmp_my, key=lambda x: x.get("last_modified",""), reverse=True)

    return render_template(
        "home.html",
        catalog=catalog,
        selected_ship=ship_number,
        due_date=due_date,
        categories=categories,
        selected_category=category,
        eqs_in_category=eqs_in_category,
        cat_status=cat_status,
        owners=owners,
        system_ex_proof=system_ex_proof,
        my_items_in_system=my_items_in_system,
        shared_items_in_system=shared_items_in_system,
        CATALOG_EQUIPMENTS=CATALOG_EQUIPMENTS
    )

def _recompute_status(item: dict) -> str:
    fields = ["qty", "maker", "type", "cert_no"]
    filled = all((item.get(k) or "").strip() for k in fields)
    return "done" if filled else "pending"

# ================= 장비 수정 =================
@app.route("/edit/<ship_number>/<category>/<eq>", methods=["GET", "POST"])
def edit(ship_number, category, eq):
    catalog = get_or_create_catalog(ship_number)
    _ensure_item(ship_number, catalog, category, eq)

    if request.method == "POST":
        qty   = request.form.get("qty")
        maker = request.form.get("maker")
        typ   = request.form.get("type")
        cert  = request.form.get("cert_no")

        ex_grade = request.form.get("ex_proof_grade")
        ip_grade = request.form.get("ip_grade")
        loc_txt  = request.form.get("location")
        page_txt = request.form.get("page")

        file  = request.files.get("file")
        submitter_name = (request.form.get("submitter_name") or "").strip()

        if category in catalog and eq in catalog[category]:
            item = catalog[category][eq]
            if qty is not None: item["qty"] = qty
            if maker is not None: item["maker"] = maker
            if typ is not None:   item["type"] = typ
            if cert is not None:  item["cert_no"] = cert
            if ex_grade is not None: item["ex_proof_grade"] = ex_grade
            if ip_grade is not None: item["ip_grade"] = ip_grade
            if loc_txt  is not None: item["location"] = loc_txt
            if page_txt is not None: item["page"] = page_txt
            if submitter_name:
                item["submitter_name"] = submitter_name
            elif session.get("user",{}).get("email"):
                item["submitter_name"] = session["user"]["email"]
            item["last_modified"] = datetime.datetime.now().isoformat()
            if file and file.filename != "":
                s3 = s3_client()
                safe = secure_filename(file.filename)
                key_file = f"{CATALOG_PREFIX}uploads/edit/{ship_number}_{secure_filename(category)}_{secure_filename(eq)}_{int(datetime.datetime.now().timestamp())}_{safe}"
                s3.upload_fileobj(file, S3_BUCKET, key_file, ExtraArgs={"ContentType": file.mimetype, "CacheControl": "no-cache"})
                item["file"] = safe; item["file_key"] = key_file; item["file_url"] = ""
            item["status"] = _recompute_status(item)

        save_catalog(ship_number, catalog)
        append_activity_log({"ts": datetime.datetime.now().isoformat(),"actor": session.get("user",{}).get("email","guest"),
                             "action": "edit","ship": ship_number, "category": category, "equipment": eq,"source": "edit_route"})

        next_url = request.args.get("next") or request.form.get("next")
        if next_url and next_url.startswith("/"):
            if "?" in next_url: next_url = f"{next_url}&_={int(time.time())}"
            else: next_url = f"{next_url}?_={int(time.time())}"
            return redirect(next_url)
        return redirect(url_for("home", ship_number=ship_number, category=category, _=int(time.time())))

    info = catalog.get(category, {}).get(eq, {})
    return render_template("edit.html", ship_number=ship_number, category=category, eq=eq, info=info)

# ================== 카테고리 담당/상태 ==================
def _send_category_warning(ship_number: str, category: str, owners: list, status_label: str):
    emails = [ (o.get("email") or "").strip().lower() for o in owners if isinstance(o, dict) and (o.get("email")) ]
    emails = [e for e in emails if e]
    if not emails: return
    subject = f"[Ship {ship_number}] '{category}' 카테고리 상태 경고: {status_label}"
    due = SHIP_DUE_DATES.get(ship_number, "")
    body = f"""안녕하세요,

호선 '{ship_number}'의 시스템 '{category}' 상태가 '{status_label}' 로 설정되었습니다.
(기한: {due})

해당 시스템의 장비 입력을 확인해 주세요.

감사합니다.
"""
    try:
        send_email_via_smtp(emails, [], subject, body)
    except Exception as e:
        print("[WARN] category warning mail send failed:", e)

@app.route("/category/owners/update", methods=["POST"])
@login_required
def category_owners_update():
    ship_number = request.form.get("ship_number")
    category = request.form.get("category")
    names  = [request.form.get("name1","").strip(), request.form.get("name2","").strip()]
    emails = [request.form.get("email1","").strip().lower(), request.form.get("email2","").strip().lower()]
    phones = [request.form.get("phone1","").strip(), request.form.get("phone2","").strip()]
    ex_proof = (request.form.get("ex_proof") or "").strip()
    if ex_proof not in ("Y","N","Unknown",""): ex_proof = "Unknown"

    catalog = get_or_create_catalog(ship_number)
    if category not in catalog or not isinstance(catalog.get(category), dict): abort(404)
    owners = []
    for i in range(2):
        if names[i] or emails[i]:
            owners.append({"name": names[i], "email": emails[i], "phone": phones[i]})
            if names[i] or emails[i]: upsert_contact(names[i], emails[i], phones[i])
    catalog[category]["__owners__"] = owners
    catalog[category]["__ex_proof__"] = (ex_proof or catalog[category].get("__ex_proof__","Unknown")) or "Unknown"
    save_catalog(ship_number, catalog)

    append_activity_log({"ts": datetime.datetime.now().isoformat(),"actor": session.get("user",{}).get("email","user"),
                         "action": "category_owners_update","ship": ship_number, "category": category, "equipment": "-",
                         "result": f"ex_proof={catalog[category]['__ex_proof__']}"})
    return redirect(url_for("home", ship_number=ship_number, category=category, _=int(time.time())))

@app.route("/category/status", methods=["POST"])
@login_required
def category_status_set():
    ship_number = request.form.get("ship_number")
    category = request.form.get("category")
    status_label = request.form.get("status")
    if status_label not in ("미입력","미완료","완료"):
        return jsonify({"ok": False, "error": "invalid status"}), 400
    catalog = get_or_create_catalog(ship_number)
    if category not in catalog or not isinstance(catalog.get(category), dict): abort(404)
    catalog[category]["__status__"] = status_label
    save_catalog(ship_number, catalog)
    append_activity_log({"ts": datetime.datetime.now().isoformat(),"actor": session.get("user",{}).get("email","user"),
                         "action": "category_status_set","ship": ship_number, "category": category, "equipment": "-","result": status_label})
    if status_label in ("미입력","미완료"):
        _send_category_warning(ship_number, category, catalog[category].get("__owners__", []), status_label)
    if request.headers.get("X-Requested-With") == "fetch":
        return jsonify({"ok": True, "status": status_label})
    return redirect(url_for("home", ship_number=ship_number, category=category, _=int(time.time())))

# ================== Admin 기능 ==================
def _require_admin():
    if not ADMIN_ENABLED:
        flash("관리자 권한이 없습니다.")
        return redirect(url_for("home"))
    return None

@app.errorhandler(403)
def handle_forbidden(e):
    flash("관리자 권한이 없습니다.")
    return redirect(url_for("home"))


@app.route("/admin")
def admin_dashboard():
    _require_admin()
    dedupe_contacts()
    submissions = list_all_submissions()
    contacts = get_contacts()
    ships = sorted({s["ship_number"] for s in submissions}) or ["1","2","3"]

    incomplete_count = {sh: 0 for sh in ships}
    owners_by_ship = {}
    all_systems_set = set()        # 시스템 중복 제거
    cat_status_by_ship = {}        # ship별 system 상태

    for sh in ships:
        catalog = load_catalog(sh) or {}
        if not catalog:
            catalog = create_catalog(sh)
        owners_by_ship[sh] = {}
        _, _, cnt, _ = _build_missing_report(sh, catalog if catalog else {})
        incomplete_count[sh] = cnt

        for cat, eqs in (catalog or {}).items():
            if isinstance(eqs, dict):
                owners_by_ship[sh][cat] = eqs.get("__owners__", [])
                all_systems_set.add(cat)
                cat_status_by_ship.setdefault(sh, {})[cat] = (eqs.get("__status__") or "미입력")

    systems = sorted(all_systems_set)

    # 최근 액티비티 로그(기존 유지)
    logs = []
    try:
        obj = s3_client().get_object(Bucket=S3_BUCKET, Key=ACTIVITY_LOG_KEY)
        lines = obj["Body"].read().decode("utf-8").strip().splitlines()[-50:]
        for ln in lines:
            try:
                logs.append(json.loads(ln))
            except Exception:
                pass
    except Exception:
        pass

    # ✅ 메일 전송 로그: ship/system 단위로 S3에서 로드
    logs_by_ship = read_mail_logs_grouped(owners_by_ship)

    deleted_by_ship = {}
    return render_template(
        "admin.html",
        submissions=submissions,
        contacts=contacts.get("list", []),
        logs=logs,
        ships=ships,
        incomplete_count=incomplete_count,
        SHIP_DUE_DATES=SHIP_DUE_DATES,
        owners_by_ship=owners_by_ship,
        systems=systems,
        logs_by_ship=logs_by_ship,             # ship -> system -> [mail logs...]
        cat_status_by_ship=cat_status_by_ship,
        deleted_by_ship=deleted_by_ship
    )

def _is_incomplete(item: dict) -> bool:
    return _recompute_status(item) != "done"

def _build_missing_report(ship_number: str, catalog: dict):
    lines = []; to_emails = set(); total = 0; by_category = {}
    if not isinstance(catalog, dict): catalog = {}
    for category, eqs in catalog.items():
        if not isinstance(eqs, dict): continue
        cat_list = []
        for eq, info in eqs.items():
            if str(eq).startswith("__"): continue
            if not isinstance(info, dict): continue
            if info.get("__deleted__"): continue
            if _is_incomplete(info):
                total += 1; cat_list.append(eq)
        if cat_list:
            by_category[category] = cat_list
            lines.append(f"[{category}]\n" + "\n".join(f"- {x}" for x in cat_list))
        owners = eqs.get("__owners__", [])
        for o in owners:
            e = (o.get("email") or "").strip().lower()
            if e: to_emails.add(e)
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
    from_addr = SMTP_SENDER; from_name = SMTP_FROM_NAME
    msg = MIMEText(body_text + "\n\n※ 본 메일은 회신 수신되지 않습니다(no-reply).", _charset="utf-8")
    msg["From"] = formataddr((from_name, from_addr))
    if to_emails: msg["To"] = ", ".join(to_emails)
    if cc_emails: msg["Cc"] = ", ".join(cc_emails)
    msg["Subject"] = subject
    recipients = list(dict.fromkeys([*(to_emails or []), *(cc_emails or [])]))
    with smtplib.SMTP(SMTP_SERVER) as server:
        server.sendmail(from_addr, recipients, msg.as_string())

def _current_user_email() -> str:
    return (session.get("user", {}).get("email") or "").strip().lower()

def _is_owner_of_category(catalog: dict, category: str, email: str) -> bool:
    owners = (catalog.get(category) or {}).get("__owners__", [])
    for o in owners:
        if not isinstance(o, dict): continue
        if (o.get("email") or "").strip().lower() == email: return True
    return False

def cleanup_catalog_amp_keys():
    """카탈로그 파일들에서 카테고리 키에 포함된 '&amp;'를 '&'로 교체"""
    s3 = s3_client()
    try:
        resp = s3.list_objects_v2(Bucket=S3_BUCKET, Prefix=CATALOG_PREFIX)
        if "Contents" not in resp: 
            return
        for obj in resp["Contents"]:
            k = obj["Key"]
            if not k.endswith(".json"): 
                continue
            if "/contacts/" in k or "/logs/" in k or "/mails/" in k or "/auth/" in k:
                continue
            raw = s3.get_object(Bucket=S3_BUCKET, Key=k)["Body"].read()
            catalog = json.loads(raw.decode("utf-8"))
            if not isinstance(catalog, dict):
                continue

            changed = False
            for cat in list(catalog.keys()):
                if "&amp;" in cat:
                    new_cat = cat.replace("&amp;", "&")
                    if new_cat not in catalog:
                        catalog[new_cat] = catalog[cat]
                    del catalog[cat]
                    changed = True
            if changed:
                s3_put_json(k, catalog)
                print(f"[FIX] {k}: category keys '&amp;' -> '&' normalized")
    except Exception as e:
        print("[WARN] cleanup_catalog_amp_keys failed:", e)


@app.route("/admin/ship_mail/<ship_number>", methods=["POST"])
def send_ship_mail(ship_number):
    _require_admin()
    catalog = load_catalog(ship_number)
    to_emails, body_text, missing_cnt, by_category = _build_missing_report(ship_number, catalog if catalog else {})
    cc_emails = request.form.getlist("cc_emails")
    if missing_cnt == 0:
        msg = f"Ship {ship_number}: 미입력 항목이 없습니다. 메일을 보내지 않았습니다."
        if request.headers.get("X-Requested-With") == "fetch":
            return jsonify({"ok": False, "message": msg, "missing": 0, "to": [], "cc": cc_emails}), 400
        flash(msg); return redirect(url_for("admin_dashboard", _=int(time.time())))
    if not to_emails:
        msg = f"Ship {ship_number}: 미입력 항목은 있으나 카테고리 담당자 이메일이 없습니다."
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
        "ts": datetime.datetime.now().isoformat(),"ship": ship_number,"to": to_emails, "cc": cc_emails,
        "subject": subject, "body": body_text,"sent": sent,"method": "smtp","error": err,
        "missing_count": missing_cnt, "by_category": by_category
    }
    s3_put_json(f"{MAIL_ARCHIVE_PREFIX}{ship_number}_bulk_{int(datetime.datetime.now().timestamp())}.json", archive)
    append_activity_log({"ts": datetime.datetime.now().isoformat(),"actor": "admin","action": "mail_bulk_send","ship": ship_number,"result": "ok" if sent else f"fail:{err}"})
    if request.headers.get("X-Requested-With") == "fetch":
        return jsonify({"ok": sent, "message": ("전송 완료" if sent else f"전송 실패: {err}"),
                        "missing": missing_cnt, "to": to_emails, "cc": cc_emails}), (200 if sent else 500)
    flash(f"Ship {ship_number}: {'메일 전송 완료' if sent else '메일 전송 실패 - ' + (err or '')}")
    return redirect(url_for("admin_dashboard", _=int(time.time())))

@app.route("/admin/invite_owner", methods=["POST"])
def admin_invite_owner():
    _require_admin()
    ship = (request.form.get("ship") or "").strip()
    category = (request.form.get("category") or "").strip()
    email = (request.form.get("email") or "").strip().lower()
    if not (ship and category and email):
        return jsonify({"ok": False, "error": "ship/category/email required"}), 400
    cat_block = (load_catalog(ship) or {}).get(category, {})
    first_eq = ""
    if isinstance(cat_block, dict):
        for k in cat_block.keys():
            if isinstance(k, str) and not k.startswith("__"): first_eq = k; break
    next_url = url_for("edit", ship_number=ship, category=category, eq=first_eq) if first_eq else url_for("home", ship_number=ship, category=category)
    inv = _invites_load(); token = uuid.uuid4().hex
    inv["invites"][token] = {"email": email,"created": datetime.datetime.now().isoformat(), "next": next_url}
    _invites_save(inv)
    link = url_for("auth_complete", t=token, next=next_url, _external=True)
    subject = f"[HD] {ship}번선 {category} 담당자 초대"
    body = f"""안녕하세요,

'{category}' 카테고리 작업 페이지 접근을 위해 비밀번호를 생성해 주세요.
아래 링크를 눌러 비밀번호를 설정하면, 로그인 없이 바로 해당 페이지로 이동합니다.

링크: {link}

감사합니다.
"""
    ok = False; err = None
    try:
        send_email_via_smtp([email], [], subject, body); ok = True
    except Exception as e:
        err = str(e)
    append_activity_log({"ts": datetime.datetime.now().isoformat(),"actor": "admin","action": "invite_owner",
                         "ship": ship, "category": category, "equipment": "-","result": "ok" if ok else f"fail:{err}", "target": email})
    # ✅ 메일 이벤트 별도 로그(S3)
    try:
        log_mail_event(ship, category, action="invite", result=("OK" if ok else f"ERROR: {err}"), purpose="invite_owner", extra={"email": email, "by": "admin_click"})
    except Exception as _ex:
        print("[WARN] log_mail_event failed:", _ex)
    return jsonify({"ok": ok, "error": err, "link": link, "ship": ship, "category": category, "email": email, "purpose": "invite_owner"}), (200 if ok else 500)

@app.route("/admin/system_mail", methods=["POST"], endpoint="admin_system_mail")
def admin_system_mail():
    _require_admin()
    ship = (request.form.get("ship") or "").strip()
    category = (request.form.get("category") or "").strip()
    if not (ship and category):
        return jsonify({"ok": False, "error": "ship/category required"}), 400

    catalog = load_catalog(ship) or {}
    block = catalog.get(category)
    if not isinstance(block, dict):
        return jsonify({"ok": False, "error": "category not found"}), 404

    status_label = block.get("__status__", "미입력")
    if status_label not in ("미입력","미완료"):
        return jsonify({"ok": False, "error": f"status '{status_label}' is not target (미입력/미완료만 전송)"}), 400

    owners = block.get("__owners__", [])
    to_emails = [ (o.get("email") or "").strip().lower() for o in owners if isinstance(o, dict) and (o.get("email")) ]
    to_emails = [e for e in to_emails if e]
    if not to_emails:
        return jsonify({"ok": False, "error": "no owner emails"}), 400

    due = SHIP_DUE_DATES.get(ship, "")
    subject = f"[Ship {ship}] '{category}' 입력 상태 안내 ({status_label})"
    body = f"""안녕하세요,

현재 '{category}' 시스템의 입력 상태가 '{status_label}' 상태입니다.
기한({due})까지 입력 부탁드립니다.

감사합니다.
"""
    ok = False; err = None
    try:
        send_email_via_smtp(to_emails, [], subject, body); ok = True
    except Exception as e:
        err = str(e)
    append_activity_log({"ts": datetime.datetime.now().isoformat(),"actor": "admin","action": "system_mail_send",
                         "ship": ship, "category": category, "equipment": "-","result": "ok" if ok else f"fail:{err}"})
    # ✅ 메일 이벤트 별도 로그(S3)
    try:
        log_mail_event(ship, category, action="manual_mail", result=("OK" if ok else f"ERROR: {err}"), purpose="manual_system_mail", extra={"by": "admin_click"})
    except Exception as _ex:
        print("[WARN] log_mail_event failed:", _ex)

    if request.headers.get("X-Requested-With") == "fetch":
        return jsonify({"ok": ok, "error": err, "purpose": "manual_system_mail"})
    flash("시스템 메일 " + ("전송 완료" if ok else ("전송 실패: " + (err or ""))))
    return redirect(url_for("admin_dashboard", _=int(time.time())))

@app.route("/admin/item_delete", methods=["POST"])
def admin_item_delete():
    _require_admin()
    ship = (request.form.get("ship") or "").strip()
    category = (request.form.get("category") or "").strip()
    eq = (request.form.get("eq") or "").strip()
    if not (ship and category and eq):
        return jsonify({"ok": False, "error": "ship/category/eq required"}), 400
    catalog = load_catalog(ship) or {}
    if not isinstance(catalog.get(category), dict) or eq not in catalog[category]:
        return jsonify({"ok": False, "error": "item not found"}), 404
    catalog[category][eq]["__deleted__"] = True
    catalog[category][eq]["last_modified"] = datetime.datetime.now().isoformat()
    save_catalog(ship, catalog)
    append_activity_log({"ts": datetime.datetime.now().isoformat(),"actor": "admin","action": "item_delete",
                         "ship": ship, "category": category, "equipment": eq})
    return jsonify({"ok": True})

@app.route("/admin/item_restore", methods=["POST"])
def admin_item_restore():
    _require_admin()
    ship = (request.form.get("ship") or "").strip()
    category = (request.form.get("category") or "").strip()
    eq = (request.form.get("eq") or "").strip()
    if not (ship and category and eq):
        return jsonify({"ok": False, "error": "ship/category/eq required"}), 400
    catalog = load_catalog(ship) or {}
    if not isinstance(catalog.get(category), dict) or eq not in catalog[category]:
        return jsonify({"ok": False, "error": "item not found"}), 404
    catalog[category][eq]["__deleted__"] = False
    catalog[category][eq]["last_modified"] = datetime.datetime.now().isoformat()
    save_catalog(ship, catalog)
    append_activity_log({"ts": datetime.datetime.now().isoformat(),"actor": "admin","action": "item_restore",
                         "ship": ship, "category": category, "equipment": eq})
    return jsonify({"ok": True})

@app.route("/admin/invite_all_contacts", methods=["POST"])
def admin_invite_all_contacts():
    _require_admin()
    contacts = (get_contacts() or {}).get("list", [])
    emails = []
    for c in contacts:
        _, e, _ = _normalize_contact(c.get("name"), c.get("email"), c.get("phone"))
        if e: emails.append(e)
    emails = sorted(set(emails))
    if not emails:
        return jsonify({"ok": False, "error": "no contact emails"}), 400
    inv = _invites_load(); sent = 0; errs = []
    for e in emails:
        token = uuid.uuid4().hex
        inv["invites"][token] = {"email": e, "created": datetime.datetime.now().isoformat(), "next": "/"}
        try:
            link = url_for("auth_complete", t=token, next="/", _external=True)
            subject = "[HD] 시스템 접근 초대"
            body = f"안녕하세요,\n\n아래 링크에서 비밀번호를 설정하시면 시스템에 접근하실 수 있습니다:\n{link}\n\n감사합니다."
            send_email_via_smtp([e], [], subject, body); sent += 1
        except Exception as ex:
            errs.append(f"{e}:{ex}")
    _invites_save(inv)
    append_activity_log({"ts": datetime.datetime.now().isoformat(),"actor": "admin","action": "invite_all_contacts",
                         "ship": "-", "category": "-", "equipment": "-","result": f"sent={sent}, errors={len(errs)}"})
    if errs:
        return jsonify({"ok": True, "sent": sent, "errors": errs}), 207
    return jsonify({"ok": True, "sent": sent})

@app.route("/admin/catalog_regen/<ship_number>", methods=["POST"])
def admin_catalog_regen(ship_number):
    _require_admin()
    create_catalog(ship_number)
    append_activity_log({"ts": datetime.datetime.now().isoformat(),"actor": "admin","action": "catalog_regen",
                         "ship": ship_number, "category": "-", "equipment": "-","result": "ok"})
    if request.headers.get("X-Requested-With") == "fetch":
        return jsonify({"ok": True})
    flash(f"Ship {ship_number} 카탈로그를 7~10개 랜덤으로 재생성했습니다.")
    return redirect(url_for("admin_dashboard", _=int(time.time())))

# ---------- Admin: Excel Export ----------
@app.route("/admin/export_selected", methods=["POST"], endpoint="export_selected")
def export_selected():
    _require_admin()
    rows = request.form.getlist("rows[]")
    all_items = {(f"{s['ship_number']}|{s['category']}|{s['equipment_name']}"): s for s in list_all_submissions()}
    picked = [all_items[r] for r in rows if r in all_items]
    if not picked:
        flash("선택된 항목이 없습니다.")
        return redirect(url_for("admin_dashboard", _=int(time.time())))
    wb = Workbook(); ws = wb.active; ws.title = "Selected"
    ws.append(["Ship","System(Category)","Equipment","QTY","Maker","Type","Cert No.","EX-PROOF GRADE","IP GRADE","PAGE","LOCATION"])
    for it in picked:
        ws.append([it["ship_number"], it["category"], it["equipment_name"], it.get("qty",""), it.get("maker",""),
                   it.get("type",""), it.get("cert_no",""), it.get("ex_proof_grade",""), it.get("ip_grade",""),
                   it.get("page",""), it.get("location","")])
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    filename = f"selected_export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(bio, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/export/excel", endpoint="export_excel")
def export_excel():
    _require_admin()
    items = list_all_submissions()
    wb = Workbook(); ws = wb.active; ws.title = "All"
    ws.append(["Ship","System(Category)","Equipment","QTY","Maker","Type","Cert No.","EX-PROOF GRADE","IP GRADE","PAGE","LOCATION"])
    for it in items:
        ws.append([it["ship_number"], it["category"], it["equipment_name"], it.get("qty",""), it.get("maker",""),
                   it.get("type",""), it.get("cert_no",""), it.get("ex_proof_grade",""), it.get("ip_grade",""),
                   it.get("page",""), it.get("location","")])
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    filename = f"export_all_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(bio, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ✅ 복구: admin.html에서 쓰는 포인트 관리 링크 엔드포인트
@app.route("/viz/manage/<ship_number>/<category>/<eq>")
def viz_manage(ship_number, category, eq):
    # 포인트 관리 화면 구현 이전: 일단 Edit로 연결
    return redirect(url_for("edit", ship_number=ship_number, category=category, eq=eq))

# ---------- 진단/헬스 ----------
def _require_token():
    t = request.args.get("token") or request.headers.get("X-Boot-Token") or ""
    if not BOOT_TOKEN or t != BOOT_TOKEN:
        return False
    return True

@app.route("/diag/aws")
def diag_aws():
    if not _require_token(): return jsonify({"ok": False, "error": "unauthorized"}), 401
    info = {"ok": True, "region": S3_REGION}
    try:
        me = sts_client().get_caller_identity()
        info["identity"] = {"Account": me.get("Account"), "Arn": me.get("Arn"), "UserId": me.get("UserId")}
    except Exception as e:
        info["identity_error"] = str(e)
    info["env"] = {"S3_BUCKET": S3_BUCKET, "CATALOG_PREFIX": CATALOG_PREFIX, "AUTO_CREATE_CATALOG": AUTO_CREATE_CATALOG, "ADMIN_ENABLED": ADMIN_ENABLED, "AUTO_QTY_ENABLED": AUTO_QTY_ENABLED}
    return jsonify(info)

@app.route("/diag/s3")
def diag_s3():
    if not _require_token(): return jsonify({"ok": False, "error": "unauthorized"}), 401
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
    if not _require_token(): return jsonify({"ok": False, "error": "unauthorized"}), 401
    key = request.args.get("key")
    if not key: return jsonify({"ok": False, "error": "key required"}), 400
    s3 = s3_client()
    out = {"ok": True, "key": key}
    try:
        h = s3.head_object(Bucket=S3_BUCKET, Key=key)
        out["exists"] = True; out["etag"] = h.get("ETag"); out["size"] = h.get("ContentLength")
    except Exception as e:
        out["exists"] = False; out["head_error"] = str(e)
    return jsonify(out)

@app.route("/diag/s3/put_test", methods=["POST"])
def diag_s3_put():
    if not _require_token(): return jsonify({"ok": False, "error": "unauthorized"}), 401
    key = f"{CATALOG_PREFIX}__ping__/{int(time.time())}.txt"
    s3 = s3_client()
    try:
        s3.put_object(Bucket=S3_BUCKET, Key=key, Body=b"ping", ContentType="text/plain", CacheControl="no-cache")
        obj = s3.get_object(Bucket=S3_BUCKET, Key=key)
        body = obj["Body"].read().decode("utf-8")
        return jsonify({"ok": True, "key": key, "readback": body})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e), "key": key}), 500

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
        cleanup_bad_logs()
        cleanup_catalog_amp_keys()
    else:
        print("[WARN] S3 env not set or partial. Skipping contacts/catalog cleanup.")

    print("[BOOT] S3_BUCKET=", S3_BUCKET, " S3_REGION=", S3_REGION, " PREFIX=", CATALOG_PREFIX, " AUTO_CREATE_CATALOG=", AUTO_CREATE_CATALOG, " ADMIN_ENABLED=", ADMIN_ENABLED, " AUTO_QTY_ENABLED=", AUTO_QTY_ENABLED)
    app.run(host="0.0.0.0", port=5000, debug=True)
