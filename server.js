require("dotenv").config();
const express    = require("express");
const XLSX       = require("xlsx");
const cors       = require("cors");
const fs         = require("fs");
const path       = require("path");
const nodemailer = require("nodemailer");
const { Resend } = require("resend");

const app  = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: "10mb" }));
app.use(express.static(path.join(__dirname, "public")));

// Serve SheetJS browser build for client-side Excel parsing
app.get("/xlsx.full.min.js", (_req, res) =>
  res.sendFile(path.join(__dirname, "node_modules/xlsx/dist/xlsx.full.min.js"))
);

// ── File paths ────────────────────────────────────────
const EXCEL_FILE = path.join(__dirname, "Edunet ids.xlsx");
const STATE_FILE = path.join(__dirname, "state.json");

// ── Storage: Firestore → JSONBin → local file ─────────
// Priority: Firestore (GCP) > JSONBin (Render free) > local file (dev)
const USE_FIRESTORE = !!process.env.GOOGLE_CLOUD_PROJECT || !!process.env.FIRESTORE_PROJECT_ID;
const JSONBIN_KEY   = process.env.JSONBIN_KEY   || "";   // X-Master-Key from jsonbin.io
const JSONBIN_BIN   = process.env.JSONBIN_BIN   || "";   // Bin ID from jsonbin.io
const USE_JSONBIN   = !USE_FIRESTORE && !!JSONBIN_KEY && !!JSONBIN_BIN;

let db = null;

if (USE_FIRESTORE) {
  const { Firestore } = require("@google-cloud/firestore");
  db = new Firestore({ projectId: process.env.FIRESTORE_PROJECT_ID || process.env.GOOGLE_CLOUD_PROJECT });
  console.log("✅ Using Firestore for storage");
} else if (USE_JSONBIN) {
  console.log("✅ Using JSONBin for persistent storage (Render-safe)");
} else {
  console.log("⚠️  Using local state.json — data will reset on Render sleep/restart.");
  console.log("   → Set JSONBIN_KEY + JSONBIN_BIN env vars on Render for persistence.");
}

// ── Helpers ───────────────────────────────────────────
async function loadState() {
  // 1. Firestore (Cloud Run / GCP)
  if (USE_FIRESTORE) {
    try {
      const doc = await db.collection("edunet").doc("state").get();
      return doc.exists ? doc.data() : null;
    } catch(e) { console.error("Firestore loadState error:", e.message); return null; }
  }

  // 2. JSONBin (free external KV — survives Render sleep)
  if (USE_JSONBIN) {
    try {
      const r = await fetch(`https://api.jsonbin.io/v3/b/${JSONBIN_BIN}/latest`, {
        headers: { "X-Master-Key": JSONBIN_KEY }
      });
      if (r.ok) {
        const j = await r.json();
        return j.record || null;
      }
      console.error("JSONBin loadState HTTP", r.status);
    } catch(e) { console.error("JSONBin loadState error:", e.message); }
    return null;
  }

  // 3. Local file (dev / ephemeral — last resort)
  if (fs.existsSync(STATE_FILE)) {
    try { return JSON.parse(fs.readFileSync(STATE_FILE, "utf8")); } catch(e) {}
  }
  return null;
}

async function saveState() {
  const data = { pool, history, employees, deletedHistory };

  // 1. Firestore
  if (USE_FIRESTORE) {
    try {
      await db.collection("edunet").doc("state").set(data);
    } catch(e) { console.error("Firestore saveState error:", e.message); }
    return;
  }

  // 2. JSONBin
  if (USE_JSONBIN) {
    try {
      const r = await fetch(`https://api.jsonbin.io/v3/b/${JSONBIN_BIN}`, {
        method : "PUT",
        headers: { "Content-Type": "application/json", "X-Master-Key": JSONBIN_KEY },
        body   : JSON.stringify(data)
      });
      if (!r.ok) console.error("JSONBin saveState HTTP", r.status, await r.text());
    } catch(e) { console.error("JSONBin saveState error:", e.message); }
    // Also write local file as hot-cache so reads during the same session are instant
    try { fs.writeFileSync(STATE_FILE, JSON.stringify(data, null, 2)); } catch(_) {}
    return;
  }

  // 3. Local file
  fs.writeFileSync(STATE_FILE, JSON.stringify(data, null, 2));
}

// ── Parse Excel ───────────────────────────────────────
function parseExcel() {
  const wb   = XLSX.readFile(EXCEL_FILE);
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const headers = (rawRows[0] || []).map(h => (h||"").toString().toLowerCase().trim());
  const emailCol = headers.findIndex(h => /mail|email/.test(h));
  const passCol  = headers.findIndex(h => /pass|pwd/.test(h));
  const snoCol   = headers.findIndex(h => /s\.?no|serial/.test(h));
  const eIdx = emailCol >= 0 ? emailCol : 1;
  const pIdx = passCol  >= 0 ? passCol  : 2;
  const sIdx = snoCol   >= 0 ? snoCol   : 0;
  return rawRows.slice(1).map((row, i) => ({
    sno          : row[sIdx] || i + 1,
    email        : (row[eIdx] || "").toString().trim(),
    password     : (row[pIdx] || "").toString().trim(),
    used         : false,
    assignedTo   : "",
    assignedEmail: "",
    date         : ""
  })).filter(r => r.email);
}

// ── Init (async — waits for Firestore on Cloud Run) ───
let pool = [], history = [], employees = [], deletedHistory = [];

async function initState() {
  const saved = await loadState();
  pool           = saved ? (saved.pool            || []) : parseExcel();
  history        = saved ? (saved.history         || []) : [];
  employees      = saved ? (saved.employees       || []) : [];
  deletedHistory = saved ? (saved.deletedHistory  || []) : [];
}

// ═══════════════════════════════════════════════════════
// REST API
// ═══════════════════════════════════════════════════════

// ── Admin login (credentials validated server-side) ───
app.post("/api/login", (req, res) => {
  const { username, password } = req.body;
  const ADMIN_USER  = process.env.ADMIN_USER  || "admin";
  const ADMIN_PASS  = process.env.ADMIN_PASS  || "admin123";
  const ADMIN_USER2 = process.env.ADMIN_USER2 || "aiaswinikumar@gmail.com";
  const ADMIN_PASS2 = process.env.ADMIN_PASS2 || "Edunet@12345";

  if (username === ADMIN_USER && password === ADMIN_PASS) {
    return res.json({ success: true, displayName: "Admin" });
  }
  if (username === ADMIN_USER2 && password === ADMIN_PASS2) {
    return res.json({ success: true, displayName: "Aiaswini Kumar" });
  }
  res.status(401).json({ error: "Invalid credentials." });
});

// Stats
app.get("/api/stats", (_req, res) => {
  const used  = pool.filter(x => x.used).length;
  res.json({ total: pool.length, used, available: pool.length - used, allocations: history.length });
});

// ID pool
app.get("/api/ids", (_req, res) => res.json({ total: pool.length, pool }));

// ── Employees ─────────────────────────────────────────
app.get("/api/employees", (_req, res) => res.json(employees));

app.post("/api/employees", async (req, res) => {
  const { name, email } = req.body;
  if (!name || !email) return res.status(400).json({ error: "name and email required" });
  if (employees.find(e => e.email.toLowerCase() === email.toLowerCase()))
    return res.status(409).json({ error: "Employee with this email already exists" });
  const emp = { id: Date.now(), name: name.trim(), email: email.trim().toLowerCase() };
  employees.push(emp);
  await saveState();
  res.json({ success: true, employee: emp });
});

app.put("/api/employees/:id", async (req, res) => {
  const emp = employees.find(e => e.id === parseInt(req.params.id));
  if (!emp) return res.status(404).json({ error: "Employee not found" });
  if (req.body.name)  emp.name  = req.body.name.trim();
  if (req.body.email) emp.email = req.body.email.trim().toLowerCase();
  await saveState();
  res.json({ success: true, employee: emp });
});

app.delete("/api/employees/:id", async (req, res) => {
  const idx = employees.findIndex(e => e.id === parseInt(req.params.id));
  if (idx === -1) return res.status(404).json({ error: "Employee not found" });
  employees.splice(idx, 1);
  await saveState();
  res.json({ success: true });
});

// ── Allocate ──────────────────────────────────────────
app.post("/api/allocate", async (req, res) => {
  let { employeeName, employeeEmail, count, assignedBy } = req.body;
  count = parseInt(count);
  if (!employeeName || !employeeEmail || !count || count < 1)
    return res.status(400).json({ error: "employeeName, employeeEmail and count are required." });

  const available = pool.filter(x => !x.used);
  if (available.length < count)
    return res.status(409).json({ error: `Only ${available.length} IDs available, requested ${count}.` });

  const allocated = available.slice(0, count);
  const now    = new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" });
  const nowISO = new Date().toISOString();   // for date filtering in reports
  allocated.forEach(id => {
    id.used          = true;
    id.assignedTo    = employeeName;
    id.assignedEmail = employeeEmail;
    id.date          = now;
  });
  const entry = {
    id         : Date.now(),
    employee   : employeeName,
    email      : employeeEmail,
    assignedBy : assignedBy || "Admin",
    count,
    date       : now,
    dateISO    : nowISO,
    ids        : allocated.map(x => ({ email: x.email, password: x.password }))
  };
  history.unshift(entry);
  await saveState();
  res.json({ success: true, allocated: entry.ids, entry });
});

// ── Delete IDs with blank passwords ──────────────────
app.post("/api/remove-blank-passwords", async (req, res) => {
  const { deletedBy } = req.body || {};
  const before  = pool.length;
  const nowISO  = new Date().toISOString();
  const nowDisp = new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" });

  // Archive used broken IDs before removing them
  pool.filter(r => r.used && (!r.email || !r.password || r.password.toString().trim() === "")).forEach(r => {
    const origEntry = history.find(h => h.ids && h.ids.some(i => i.email && i.email.toLowerCase() === (r.email||"").toLowerCase()));
    deletedHistory.push({
      id            : Date.now() + Math.random(),
      edunetEmail   : r.email   || "BLANK",
      password      : r.password|| "BLANK",
      assignedTo    : r.assignedTo    || "",
      assignedEmail : r.assignedEmail || "",
      assignedBy    : origEntry ? (origEntry.assignedBy || "Admin") : "Admin",
      allocationDate: r.date    || "",
      deletedBy     : deletedBy || "Admin",
      deletedDate   : nowDisp,
      deletedDateISO: nowISO
    });
  });

  pool = pool.filter(r =>
    r.email    && r.email.toString().trim()    !== "" &&
    r.password && r.password.toString().trim() !== ""
  );
  pool.forEach((r, i) => { r.sno = i + 1; });
  await saveState();
  res.json({ success: true, removed: before - pool.length, total: pool.length });
});

// ── Delete specific IDs by email ──────────────────────
app.post("/api/delete-ids", async (req, res) => {
  const { emails, deletedBy } = req.body;   // array of email strings
  if (!Array.isArray(emails) || !emails.length)
    return res.status(400).json({ error: "emails array is required." });
  const set     = new Set(emails.map(e => e.toLowerCase()));
  const before  = pool.length;
  const nowISO  = new Date().toISOString();
  const nowDisp = new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" });

  // Archive any history entries whose IDs are being deleted
  pool.filter(r => set.has((r.email||"").toLowerCase()) && r.used).forEach(r => {
    // Find the original allocation entry for this ID
    const origEntry = history.find(h => h.ids && h.ids.some(i => i.email.toLowerCase() === r.email.toLowerCase()));
    deletedHistory.push({
      id           : Date.now() + Math.random(),
      edunetEmail  : r.email,
      password     : r.password,
      assignedTo   : r.assignedTo   || "",
      assignedEmail: r.assignedEmail|| "",
      assignedBy   : origEntry ? (origEntry.assignedBy || "Admin") : "Admin",
      allocationDate: r.date         || "",
      deletedBy    : deletedBy || "Admin",
      deletedDate  : nowDisp,
      deletedDateISO: nowISO
    });
  });

  pool = pool.filter(r => !set.has((r.email||"").toLowerCase()));
  pool.forEach((r, i) => { r.sno = i + 1; });
  await saveState();
  res.json({ success: true, removed: before - pool.length, total: pool.length });
});

// ── Deleted history ────────────────────────────────────
app.get("/api/deleted-history", (_req, res) => res.json(deletedHistory));

// ── Report endpoint ────────────────────────────────────
app.get("/api/report", (req, res) => {
  const { period, from, to } = req.query;
  const now = new Date();

  // Parse a YYYY-MM-DD string as local midnight (avoids UTC offset bug)
  function localDay(str, endOfDay = false) {
    const [y, m, d] = str.split("-").map(Number);
    return endOfDay
      ? new Date(y, m - 1, d, 23, 59, 59, 999)
      : new Date(y, m - 1, d, 0, 0, 0, 0);
  }

  let start, end;

  if (period === "weekly") {
    end   = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
    start = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 6, 0, 0, 0, 0);
  } else if (period === "monthly") {
    start = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);
    end   = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
  } else if (period === "quarterly") {
    const q = Math.floor(now.getMonth() / 3);
    start   = new Date(now.getFullYear(), q * 3, 1, 0, 0, 0, 0);
    end     = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
  } else if (from && to) {
    start = localDay(from, false);
    end   = localDay(to,   true);
  } else {
    // All time
    start = new Date(0);
    end   = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
  }

  // Filter active history
  const active = history.filter(h => {
    const d = h.dateISO ? new Date(h.dateISO) : null;
    if (!d) return true;   // legacy entries without dateISO always included
    return d >= start && d <= end;
  });

  // Filter deleted history
  const deleted = deletedHistory.filter(h => {
    const d = h.deletedDateISO ? new Date(h.deletedDateISO) : null;
    if (!d) return true;
    return d >= start && d <= end;
  });

  // Summary per assignedBy
  const byAssigner = {};
  active.forEach(h => {
    const k = h.assignedBy || "Admin";
    if (!byAssigner[k]) byAssigner[k] = { name: k, allocations: 0, idsAssigned: 0 };
    byAssigner[k].allocations++;
    byAssigner[k].idsAssigned += h.count;
  });

  res.json({
    period  : period || "custom",
    from    : start.toISOString(),
    to      : end.toISOString(),
    active,
    deleted,
    summary : Object.values(byAssigner),
    totals  : {
      allocations : active.length,
      idsAssigned : active.reduce((s, h) => s + h.count, 0),
      deletedIds  : deleted.length
    }
  });
});

// ── Reset all ─────────────────────────────────────────
app.post("/api/reset", async (_req, res) => {
  pool.forEach(r => { r.used=false; r.assignedTo=""; r.assignedEmail=""; r.date=""; });
  history = [];
  await saveState();
  res.json({ success: true });
});

// ── Add more IDs ──────────────────────────────────────
app.post("/api/add-ids", async (req, res) => {
  const { ids } = req.body;           // [{ email, password }]
  if (!Array.isArray(ids) || !ids.length)
    return res.status(400).json({ error: "ids array is required." });

  let added = 0, skipped = 0;
  const existingEmails = new Set(pool.map(p => p.email.toLowerCase()));

  ids.forEach(({ email, password }) => {
    email    = (email    || "").toString().trim();
    password = (password || "").toString().trim();
    if (!email) { skipped++; return; }
    if (existingEmails.has(email.toLowerCase())) { skipped++; return; }
    existingEmails.add(email.toLowerCase());
    pool.push({
      sno          : pool.length + 1,
      email,
      password,
      used         : false,
      assignedTo   : "",
      assignedEmail: "",
      date         : ""
    });
    added++;
  });

  // Re-number sno for the whole pool
  pool.forEach((r, i) => { r.sno = i + 1; });
  await saveState();
  res.json({ success: true, added, skipped, total: pool.length });
});

// ── Reload Excel ──────────────────────────────────────
app.post("/api/reload", async (_req, res) => {
  try {
    const fresh = parseExcel();
    const freshEmails = new Set(fresh.map(r => r.email.toLowerCase()));

    // 1. Preserve used/assigned state for IDs that exist in the Excel file
    fresh.forEach(row => {
      const existing = pool.find(p => p.email.toLowerCase() === row.email.toLowerCase());
      if (existing) {
        row.used          = existing.used;
        row.assignedTo    = existing.assignedTo;
        row.assignedEmail = existing.assignedEmail;
        row.date          = existing.date;
      }
    });

    // 2. Keep any IDs that were manually added (not in the Excel file) — never drop them
    const manualIds = pool.filter(p => !freshEmails.has(p.email.toLowerCase()));

    // 3. Merge: Excel IDs first, then manual additions
    pool = [...fresh, ...manualIds];

    // 4. Re-number sno
    pool.forEach((r, i) => { r.sno = i + 1; });

    await saveState();
    res.json({ success: true, total: pool.length });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── History ───────────────────────────────────────────
app.get("/api/history", (_req, res) => res.json(history));

// ── IBM watsonx AI proxy ──────────────────────────────
app.post("/api/ai", async (req, res) => {
  const { prompt } = req.body;
  const API_KEY    = process.env.IBM_API_KEY;
  const PROJECT_ID = process.env.IBM_PROJECT_ID;
  const WX_URL     = process.env.IBM_WX_URL    || "https://us-south.ml.cloud.ibm.com";
  const MODEL      = process.env.IBM_MODEL     || "ibm/granite-3-3-8b-instruct";

  try {
    // 1. Get IAM token
    const tokenRes = await fetch("https://iam.cloud.ibm.com/identity/token", {
      method : "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body   : `grant_type=urn:ibm:params:oauth:grant-type:apikey&apikey=${API_KEY}`
    });
    const tokenData = await tokenRes.json();
    const token     = tokenData.access_token;
    if (!token) return res.status(401).json({ error: "IAM token failed. Check API key." });

    // 2. Call Granite
    const aiRes = await fetch(`${WX_URL}/ml/v1/text/generation?version=2023-05-29`, {
      method : "POST",
      headers: { "Content-Type": "application/json", "Authorization": "Bearer " + token },
      body   : JSON.stringify({
        model_id  : MODEL,
        project_id: PROJECT_ID,
        input     : prompt,
        parameters: { max_new_tokens: 300, temperature: 0.1, stop_sequences: ["```", "\n\n\n"] }
      })
    });
    const aiData = await aiRes.json();
    const text   = aiData?.results?.[0]?.generated_text || "";
    res.json({ text });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// ── Send Email via Gmail API (HTTPS — never blocked by Render) ──
app.post("/api/send-email", async (req, res) => {
  const { toName, toEmail, ids } = req.body;
  if (!toName || !toEmail || !ids || !ids.length)
    return res.status(400).json({ error: "toName, toEmail and ids are required." });

  const credLines = ids.map((id, i) =>
    `${i + 1}. Mail: ${id.email}  |  Password: ${id.password}`
  ).join("\n");

  const html = `
<div style="font-family:Arial,sans-serif;max-width:600px;margin:auto;border:1px solid #e5e7eb;border-radius:10px;overflow:hidden">
  <div style="background:#1a1f2e;padding:24px 28px">
    <h2 style="color:#fff;margin:0;font-size:20px">&#127891; Edunet IBM SkillsBuild</h2>
    <p style="color:#9ca3af;margin:4px 0 0;font-size:13px">Login Credentials</p>
  </div>
  <div style="padding:28px">
    <p style="font-size:15px;color:#1f2328">Dear <strong>${toName}</strong>,</p>
    <p style="font-size:13px;color:#57606a;margin-top:8px">
      Please find your Edunet IBM SkillsBuild login credentials assigned exclusively to you.<br>
      <strong style="color:#dc2626">Do NOT share these credentials with anyone.</strong>
    </p>
    <table style="width:100%;border-collapse:collapse;margin-top:20px;font-size:13px">
      <thead>
        <tr style="background:#f7f8fa">
          <th style="padding:10px 12px;text-align:left;border-bottom:2px solid #e5e7eb;color:#57606a">#</th>
          <th style="padding:10px 12px;text-align:left;border-bottom:2px solid #e5e7eb;color:#57606a">Edunet Mail ID</th>
          <th style="padding:10px 12px;text-align:left;border-bottom:2px solid #e5e7eb;color:#57606a">Password</th>
        </tr>
      </thead>
      <tbody>
        ${ids.map((id, i) => `
        <tr style="border-bottom:1px solid #f0f1f3">
          <td style="padding:10px 12px;color:#9ca3af">${i + 1}</td>
          <td style="padding:10px 12px;font-family:monospace;color:#1f2328">${id.email}</td>
          <td style="padding:10px 12px;font-family:monospace;color:#1f2328">${id.password}</td>
        </tr>`).join("")}
      </tbody>
    </table>
    <div style="margin-top:24px;padding:16px;background:#f0fdf4;border:1px solid #86efac;border-radius:8px;font-size:13px">
      <strong>Login Portal:</strong> <a href="https://www.edunetworks.in/" style="color:#3b82d4">https://www.edunetworks.in/</a><br>
      <strong>Support:</strong> support@edunetworks.in
    </div>
    <p style="font-size:12px;color:#9ca3af;margin-top:20px">Best regards,<br><strong>Edunet Admin Team</strong></p>
  </div>
  <div style="background:#f7f8fa;padding:12px 28px;border-top:1px solid #e5e7eb;font-size:11px;color:#9ca3af;text-align:center">
    This is an automated message. Please do not reply to this email.
  </div>
</div>`;

  // ── Try Resend first (HTTPS, no port blocking) ──
  if (process.env.RESEND_API_KEY) {
    try {
      const resend = new Resend(process.env.RESEND_API_KEY);
      const from = process.env.RESEND_FROM || `Edunet Admin <onboarding@resend.dev>`;
      await resend.emails.send({
        from,
        to     : toEmail,
        subject: `Your ${ids.length} Edunet IBM SkillsBuild Login Credential${ids.length > 1 ? "s" : ""}`,
        html,
        text   : `Dear ${toName},\n\nYour credentials:\n\n${credLines}\n\nLogin: https://www.edunetworks.in/\n\nEdunet Admin Team`
      });
      console.log("✅ Email sent via Resend to", toEmail);
      return res.json({ success: true, message: `Email sent to ${toEmail}` });
    } catch (e) {
      console.error("❌ Resend failed:", e.message);
      return res.status(500).json({ error: `Resend error: ${e.message}` });
    }
  }

  // ── Fallback: SMTP via nodemailer (ports 465 / 587) ──
  const configs = [
    { port: 465, secure: true },
    { port: 587, secure: false }
  ];

  let lastError = "";
  for (const { port, secure } of configs) {
    try {
      const transporter = nodemailer.createTransport({
        host            : "smtp.gmail.com",
        port,
        secure,
        auth            : { user: process.env.GMAIL_USER, pass: process.env.GMAIL_PASS },
        tls             : { rejectUnauthorized: false },
        connectionTimeout: 10000,
        greetingTimeout : 10000,
        socketTimeout   : 15000
      });

      await transporter.sendMail({
        from   : `"Edunet Admin" <${process.env.GMAIL_USER}>`,
        to     : toEmail,
        subject: `Your ${ids.length} Edunet IBM SkillsBuild Login Credential${ids.length > 1 ? "s" : ""}`,
        html,
        text   : `Dear ${toName},\n\nYour credentials:\n\n${credLines}\n\nLogin: https://www.edunetworks.in/\n\nEdunet Admin Team`
      });

      console.log(`✅ Email sent via SMTP port ${port} to`, toEmail);
      return res.json({ success: true, message: `Email sent to ${toEmail}` });
    } catch (e) {
      lastError = e.message;
      console.error(`❌ SMTP port ${port} failed:`, e.message);
    }
  }

  res.status(500).json({ error: `All email methods failed. Last error: ${lastError}` });
});

// ── SPA fallback ──────────────────────────────────────
app.get("*", (_req, res) =>
  res.sendFile(path.join(__dirname, "public", "index.html"))
);

// ── Start server after state is loaded ────────────────
initState().then(() => {
  app.listen(PORT, () => {
    console.log(`\n✅  Edunet ID Agent  →  http://localhost:${PORT}`);
    console.log(`   Pool : ${pool.length} IDs   |   Employees : ${employees.length}`);
    console.log(`   Storage: ${USE_FIRESTORE ? "Firestore ☁️" : "Local file 📁"}`);
  });
}).catch(e => {
  console.error("❌ Failed to init state:", e.message);
  process.exit(1);
});
