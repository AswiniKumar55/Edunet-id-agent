import os, json, re, time, smtplib, random
from pathlib import Path
from datetime import datetime, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import openpyxl
import httpx
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional
from dotenv import load_dotenv

load_dotenv()

# ── Config ────────────────────────────────────────────
EXCEL_FILE  = Path("Edunet ids.xlsx")
STATE_FILE  = Path("state.json")
PUBLIC_DIR  = Path("public")

ADMIN_USER  = os.getenv("ADMIN_USER",  "admin")
ADMIN_PASS  = os.getenv("ADMIN_PASS",  "admin123")
ADMIN_USER2 = os.getenv("ADMIN_USER2", "aiaswinikumar@gmail.com")
ADMIN_PASS2 = os.getenv("ADMIN_PASS2", "Edunet@12345")
GMAIL_USER  = os.getenv("GMAIL_USER", "")
GMAIL_PASS  = os.getenv("GMAIL_PASS", "")
IBM_API_KEY = os.getenv("IBM_API_KEY", "")
IBM_PROJECT = os.getenv("IBM_PROJECT_ID", "")
IBM_WX_URL  = os.getenv("IBM_WX_URL", "https://us-south.ml.cloud.ibm.com")
IBM_MODEL   = os.getenv("IBM_MODEL",  "ibm/granite-3-3-8b-instruct")

# ── In-memory state ───────────────────────────────────
pool:           list = []
history:        list = []
employees:      list = []
deleted_history:list = []

# ── Helpers ───────────────────────────────────────────
def now_ist() -> str:
    from zoneinfo import ZoneInfo
    return datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d/%m/%Y, %I:%M:%S %p")

def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()

def load_state():
    global pool, history, employees, deleted_history
    if STATE_FILE.exists():
        try:
            data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            pool            = data.get("pool",           [])
            history         = data.get("history",        [])
            employees       = data.get("employees",      [])
            deleted_history = data.get("deletedHistory", [])
            return
        except Exception:
            pass
    # Fresh start from Excel
    pool            = parse_excel()
    history         = []
    employees       = []
    deleted_history = []

def save_state():
    STATE_FILE.write_text(
        json.dumps({"pool": pool, "history": history,
                    "employees": employees, "deletedHistory": deleted_history}, indent=2),
        encoding="utf-8"
    )

def parse_excel() -> list:
    if not EXCEL_FILE.exists():
        return []
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).lower().strip() if h else "" for h in rows[0]]

    def find(pattern):
        for i, h in enumerate(headers):
            if re.search(pattern, h):
                return i
        return -1

    e_idx = find(r"mail|email");  e_idx = e_idx if e_idx >= 0 else 1
    p_idx = find(r"pass|pwd");    p_idx = p_idx if p_idx >= 0 else 2
    s_idx = find(r"s\.?no|serial"); s_idx = s_idx if s_idx >= 0 else 0

    result = []
    for i, row in enumerate(rows[1:], 1):
        email = str(row[e_idx]).strip() if row[e_idx] else ""
        pwd   = str(row[p_idx]).strip() if row[p_idx] else ""
        sno   = row[s_idx] if row[s_idx] else i
        if email and email.lower() != "none":
            result.append({
                "sno": sno, "email": email, "password": pwd,
                "used": False, "assignedTo": "", "assignedEmail": "", "date": ""
            })
    return result

# ── Init ──────────────────────────────────────────────
load_state()

# ── FastAPI app ───────────────────────────────────────
app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"],
                   allow_methods=["*"], allow_headers=["*"])

# ── Pydantic models ───────────────────────────────────
class LoginBody(BaseModel):
    username: str
    password: str

class EmployeeBody(BaseModel):
    name:  str
    email: str

class EmployeeUpdate(BaseModel):
    name:  Optional[str] = None
    email: Optional[str] = None

class AllocateBody(BaseModel):
    employeeName:  str
    employeeEmail: str
    count:         int
    assignedBy:    Optional[str] = "Admin"

class DeleteIdsBody(BaseModel):
    emails:    List[str]
    deletedBy: Optional[str] = "Admin"

class RemoveBlankBody(BaseModel):
    deletedBy: Optional[str] = "Admin"

class AddIdsBody(BaseModel):
    ids: List[dict]

class AIBody(BaseModel):
    prompt: str

class SendEmailBody(BaseModel):
    toName:  str
    toEmail: str
    ids:     List[dict]

# ══════════════════════════════════════════════════════
# API ROUTES
# ══════════════════════════════════════════════════════

@app.post("/api/login")
def login(body: LoginBody):
    if body.username == ADMIN_USER and body.password == ADMIN_PASS:
        return {"success": True, "displayName": "Admin"}
    if body.username == ADMIN_USER2 and body.password == ADMIN_PASS2:
        return {"success": True, "displayName": "Aiaswini Kumar"}
    raise HTTPException(401, "Invalid credentials.")

@app.get("/api/stats")
def stats():
    used = sum(1 for x in pool if x["used"])
    return {"total": len(pool), "used": used,
            "available": len(pool) - used, "allocations": len(history)}

@app.get("/api/ids")
def get_ids():
    return {"total": len(pool), "pool": pool}

@app.get("/api/history")
def get_history():
    return history

@app.get("/api/employees")
def get_employees():
    return employees

@app.post("/api/employees")
def add_employee(body: EmployeeBody):
    if not body.name or not body.email:
        raise HTTPException(400, "name and email required")
    if any(e["email"].lower() == body.email.lower() for e in employees):
        raise HTTPException(409, "Employee with this email already exists")
    emp = {"id": int(time.time() * 1000), "name": body.name.strip(),
           "email": body.email.strip().lower()}
    employees.append(emp)
    save_state()
    return {"success": True, "employee": emp}

@app.put("/api/employees/{emp_id}")
def update_employee(emp_id: int, body: EmployeeUpdate):
    emp = next((e for e in employees if e["id"] == emp_id), None)
    if not emp:
        raise HTTPException(404, "Employee not found")
    if body.name:  emp["name"]  = body.name.strip()
    if body.email: emp["email"] = body.email.strip().lower()
    save_state()
    return {"success": True, "employee": emp}

@app.delete("/api/employees/{emp_id}")
def delete_employee(emp_id: int):
    global employees
    before = len(employees)
    employees = [e for e in employees if e["id"] != emp_id]
    if len(employees) == before:
        raise HTTPException(404, "Employee not found")
    save_state()
    return {"success": True}

@app.post("/api/allocate")
def allocate(body: AllocateBody):
    available = [x for x in pool if not x["used"]]
    if len(available) < body.count:
        raise HTTPException(409, f"Only {len(available)} IDs available, requested {body.count}.")
    allocated = available[:body.count]
    ts_ist = now_ist(); ts_iso = now_iso()
    for id_ in allocated:
        id_["used"]          = True
        id_["assignedTo"]    = body.employeeName
        id_["assignedEmail"] = body.employeeEmail
        id_["date"]          = ts_ist
    entry = {
        "id":         int(time.time() * 1000),
        "employee":   body.employeeName,
        "email":      body.employeeEmail,
        "assignedBy": body.assignedBy or "Admin",
        "count":      body.count,
        "date":       ts_ist,
        "dateISO":    ts_iso,
        "ids":        [{"email": x["email"], "password": x["password"]} for x in allocated]
    }
    history.insert(0, entry)
    save_state()
    return {"success": True, "allocated": entry["ids"], "entry": entry}

@app.post("/api/remove-blank-passwords")
def remove_blank(body: RemoveBlankBody):
    global pool
    before = len(pool)
    ts_ist = now_ist(); ts_iso = now_iso()
    for r in pool:
        if r["used"] and (not r["email"] or not str(r.get("password","")).strip()):
            orig = next((h for h in history if any(
                i["email"].lower() == r["email"].lower() for i in (h.get("ids") or [])
            )), None)
            deleted_history.append({
                "id": time.time() + random.random(),
                "edunetEmail": r["email"] or "BLANK",
                "password":    r.get("password") or "BLANK",
                "assignedTo":    r.get("assignedTo",""),
                "assignedEmail": r.get("assignedEmail",""),
                "assignedBy":    orig["assignedBy"] if orig else "Admin",
                "allocationDate":r.get("date",""),
                "deletedBy":     body.deletedBy or "Admin",
                "deletedDate":   ts_ist,
                "deletedDateISO":ts_iso
            })
    pool = [r for r in pool if r["email"] and str(r.get("password","")).strip()]
    for i, r in enumerate(pool): r["sno"] = i + 1
    save_state()
    return {"success": True, "removed": before - len(pool), "total": len(pool)}

@app.post("/api/delete-ids")
def delete_ids(body: DeleteIdsBody):
    global pool
    s = set(e.lower() for e in body.emails)
    before = len(pool)
    ts_ist = now_ist(); ts_iso = now_iso()
    for r in pool:
        if (r.get("email","")).lower() in s and r["used"]:
            orig = next((h for h in history if any(
                i["email"].lower() == r["email"].lower() for i in (h.get("ids") or [])
            )), None)
            deleted_history.append({
                "id": time.time() + random.random(),
                "edunetEmail":   r["email"],
                "password":      r.get("password",""),
                "assignedTo":    r.get("assignedTo",""),
                "assignedEmail": r.get("assignedEmail",""),
                "assignedBy":    orig["assignedBy"] if orig else "Admin",
                "allocationDate":r.get("date",""),
                "deletedBy":     body.deletedBy or "Admin",
                "deletedDate":   ts_ist,
                "deletedDateISO":ts_iso
            })
    pool = [r for r in pool if (r.get("email","")).lower() not in s]
    for i, r in enumerate(pool): r["sno"] = i + 1
    save_state()
    return {"success": True, "removed": before - len(pool), "total": len(pool)}

@app.get("/api/deleted-history")
def get_deleted_history():
    return deleted_history

@app.get("/api/report")
def report(period: str = "", from_: str = "", to: str = ""):
    from datetime import timedelta
    now = datetime.now(timezone.utc)

    if period == "weekly":
        start = datetime(now.year, now.month, now.day, tzinfo=timezone.utc) - timedelta(days=6)
        end   = datetime(now.year, now.month, now.day, 23, 59, 59, tzinfo=timezone.utc)
    elif period == "monthly":
        start = datetime(now.year, now.month, 1, tzinfo=timezone.utc)
        end   = datetime(now.year, now.month, now.day, 23, 59, 59, tzinfo=timezone.utc)
    elif period == "quarterly":
        q = (now.month - 1) // 3
        start = datetime(now.year, q * 3 + 1, 1, tzinfo=timezone.utc)
        end   = datetime(now.year, now.month, now.day, 23, 59, 59, tzinfo=timezone.utc)
    elif from_ and to:
        y1,m1,d1 = map(int, from_.split("-"))
        y2,m2,d2 = map(int, to.split("-"))
        start = datetime(y1, m1, d1, tzinfo=timezone.utc)
        end   = datetime(y2, m2, d2, 23, 59, 59, tzinfo=timezone.utc)
    else:
        start = datetime(1970, 1, 1, tzinfo=timezone.utc)
        end   = datetime(now.year, now.month, now.day, 23, 59, 59, tzinfo=timezone.utc)

    def in_range(iso):
        if not iso: return True
        try:
            d = datetime.fromisoformat(iso.replace("Z","+00:00"))
            return start <= d <= end
        except: return True

    active  = [h for h in history        if in_range(h.get("dateISO"))]
    deleted = [h for h in deleted_history if in_range(h.get("deletedDateISO"))]

    by_assigner = {}
    for h in active:
        k = h.get("assignedBy","Admin")
        if k not in by_assigner:
            by_assigner[k] = {"name": k, "allocations": 0, "idsAssigned": 0}
        by_assigner[k]["allocations"] += 1
        by_assigner[k]["idsAssigned"] += h.get("count", 0)

    return {
        "period":  period or "custom",
        "from":    start.isoformat(), "to": end.isoformat(),
        "active":  active, "deleted": deleted,
        "summary": list(by_assigner.values()),
        "totals":  {
            "allocations": len(active),
            "idsAssigned": sum(h.get("count",0) for h in active),
            "deletedIds":  len(deleted)
        }
    }

@app.post("/api/reset")
def reset():
    for r in pool:
        r["used"] = False; r["assignedTo"] = ""; r["assignedEmail"] = ""; r["date"] = ""
    history.clear()
    save_state()
    return {"success": True}

@app.post("/api/add-ids")
def add_ids(body: AddIdsBody):
    existing = set(p["email"].lower() for p in pool)
    added = skipped = 0
    for item in body.ids:
        email = str(item.get("email","")).strip()
        pwd   = str(item.get("password","")).strip()
        if not email or email.lower() in existing:
            skipped += 1; continue
        existing.add(email.lower())
        pool.append({"sno": len(pool)+1, "email": email, "password": pwd,
                     "used": False, "assignedTo": "", "assignedEmail": "", "date": ""})
        added += 1
    for i, r in enumerate(pool): r["sno"] = i + 1
    save_state()
    return {"success": True, "added": added, "skipped": skipped, "total": len(pool)}

@app.post("/api/reload")
def reload_excel():
    global pool
    try:
        fresh = parse_excel()
        fresh_emails = set(r["email"].lower() for r in fresh)
        for row in fresh:
            existing = next((p for p in pool if p["email"].lower() == row["email"].lower()), None)
            if existing:
                row["used"] = existing["used"]; row["assignedTo"] = existing["assignedTo"]
                row["assignedEmail"] = existing["assignedEmail"]; row["date"] = existing["date"]
        manual = [p for p in pool if p["email"].lower() not in fresh_emails]
        pool = fresh + manual
        for i, r in enumerate(pool): r["sno"] = i + 1
        save_state()
        return {"success": True, "total": len(pool)}
    except Exception as e:
        raise HTTPException(500, str(e))

@app.post("/api/ai")
async def ai_proxy(body: AIBody):
    try:
        async with httpx.AsyncClient(timeout=30) as client:
            token_res = await client.post(
                "https://iam.cloud.ibm.com/identity/token",
                headers={"Content-Type": "application/x-www-form-urlencoded"},
                data=f"grant_type=urn:ibm:params:oauth:grant-type:apikey&apikey={IBM_API_KEY}"
            )
            token = token_res.json().get("access_token")
            if not token:
                raise HTTPException(401, "IAM token failed. Check API key.")
            ai_res = await client.post(
                f"{IBM_WX_URL}/ml/v1/text/generation?version=2023-05-29",
                headers={"Content-Type": "application/json", "Authorization": f"Bearer {token}"},
                json={"model_id": IBM_MODEL, "project_id": IBM_PROJECT,
                      "input": body.prompt,
                      "parameters": {"max_new_tokens": 300, "temperature": 0.1,
                                     "stop_sequences": ["```", "\n\n\n"]}}
            )
            text = ai_res.json().get("results", [{}])[0].get("generated_text", "")
            return {"text": text}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, str(e))

@app.post("/api/send-email")
def send_email(body: SendEmailBody):
    if not body.toName or not body.toEmail or not body.ids:
        raise HTTPException(400, "toName, toEmail and ids are required.")

    cred_lines = "\n".join(
        f"{i+1}. Mail: {id_['email']}  |  Password: {id_['password']}"
        for i, id_ in enumerate(body.ids)
    )
    html = f"""
<div style="font-family:Arial,sans-serif;max-width:600px;margin:auto;border:1px solid #e5e7eb;border-radius:10px;overflow:hidden">
  <div style="background:#1a1f2e;padding:24px 28px">
    <h2 style="color:#fff;margin:0;font-size:20px">&#127891; Edunet IBM SkillsBuild</h2>
    <p style="color:#9ca3af;margin:4px 0 0;font-size:13px">Login Credentials</p>
  </div>
  <div style="padding:28px">
    <p style="font-size:15px;color:#1f2328">Dear <strong>{body.toName}</strong>,</p>
    <p style="font-size:13px;color:#57606a;margin-top:8px">
      Please find your Edunet IBM SkillsBuild login credentials.<br>
      <strong style="color:#dc2626">Do NOT share these credentials with anyone.</strong>
    </p>
    <table style="width:100%;border-collapse:collapse;margin-top:20px;font-size:13px">
      <thead><tr style="background:#f7f8fa">
        <th style="padding:10px 12px;text-align:left;border-bottom:2px solid #e5e7eb;color:#57606a">#</th>
        <th style="padding:10px 12px;text-align:left;border-bottom:2px solid #e5e7eb;color:#57606a">Edunet Mail ID</th>
        <th style="padding:10px 12px;text-align:left;border-bottom:2px solid #e5e7eb;color:#57606a">Password</th>
      </tr></thead>
      <tbody>
        {"".join(f'<tr style="border-bottom:1px solid #f0f1f3"><td style="padding:10px 12px;color:#9ca3af">{i+1}</td><td style="padding:10px 12px;font-family:monospace">{id_["email"]}</td><td style="padding:10px 12px;font-family:monospace">{id_["password"]}</td></tr>' for i, id_ in enumerate(body.ids))}
      </tbody>
    </table>
    <div style="margin-top:24px;padding:16px;background:#f0fdf4;border:1px solid #86efac;border-radius:8px;font-size:13px">
      <strong>Login Portal:</strong> <a href="https://www.edunetworks.in/" style="color:#3b82d4">https://www.edunetworks.in/</a><br>
      <strong>Support:</strong> support@edunetworks.in
    </div>
    <p style="font-size:12px;color:#9ca3af;margin-top:20px">Best regards,<br><strong>Edunet Admin Team</strong></p>
  </div>
  <div style="background:#f7f8fa;padding:12px 28px;border-top:1px solid #e5e7eb;font-size:11px;color:#9ca3af;text-align:center">
    This is an automated message. Please do not reply.
  </div>
</div>"""

    subject = f"Your {len(body.ids)} Edunet IBM SkillsBuild Login Credential{'s' if len(body.ids)>1 else ''}"
    errors = []

    for port, use_ssl in [(465, True), (587, False)]:
        try:
            msg = MIMEMultipart("alternative")
            msg["Subject"] = subject
            msg["From"]    = f"Edunet Admin <{GMAIL_USER}>"
            msg["To"]      = body.toEmail
            msg.attach(MIMEText(f"Dear {body.toName},\n\n{cred_lines}\n\nLogin: https://www.edunetworks.in/\n\nEdunet Admin Team", "plain"))
            msg.attach(MIMEText(html, "html"))

            if use_ssl:
                with smtplib.SMTP_SSL("smtp.gmail.com", port, timeout=30) as s:
                    s.login(GMAIL_USER, GMAIL_PASS)
                    s.sendmail(GMAIL_USER, body.toEmail, msg.as_string())
            else:
                with smtplib.SMTP("smtp.gmail.com", port, timeout=30) as s:
                    s.starttls()
                    s.login(GMAIL_USER, GMAIL_PASS)
                    s.sendmail(GMAIL_USER, body.toEmail, msg.as_string())

            print(f"✅ Email sent via port {port} to {body.toEmail}")
            return {"success": True, "message": f"Email sent to {body.toEmail}"}
        except Exception as e:
            errors.append(f"Port {port}: {e}")
            print(f"❌ Port {port} failed: {e}")

    raise HTTPException(500, " | ".join(errors))

# ── Serve static frontend ─────────────────────────────
if PUBLIC_DIR.exists():
    app.mount("/public", StaticFiles(directory="public"), name="public")

@app.get("/xlsx.full.min.js")
def xlsx_js():
    p = Path("node_modules/xlsx/dist/xlsx.full.min.js")
    if p.exists():
        return FileResponse(p)
    raise HTTPException(404, "xlsx.full.min.js not found")

@app.get("/{full_path:path}")
def spa_fallback(full_path: str):
    index = PUBLIC_DIR / "index.html"
    if index.exists():
        return FileResponse(index)
    raise HTTPException(404, "index.html not found")

@app.get("/")
def root():
    return FileResponse(PUBLIC_DIR / "index.html")
