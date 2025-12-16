const express = require("express");
const multer = require("multer");
const os = require("os");
const path = require("path");
const fs = require("fs");
const { spawnSync } = require("child_process");
const XLSX = require("xlsx");
const crypto = require("crypto");

const config = require("./config");
const { appendSentRow, getSentWorkbookBuffer } = require("./excel-log");
const { createTransporter, sendApplicationEmail } = require("./mailer");
const { buildEmail } = require("./template");
const { sleep } = require("./utils");
const { readJson, writeJsonAtomic, ensureDir } = require("./utils");

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const upload = multer({
  dest: path.join(os.tmpdir(), "job-mailer-uploads"),
  limits: {
    fileSize: 12 * 1024 * 1024, // 12MB
  },
});

function normalizeEmail(email) {
  return String(email || "").trim().toLowerCase();
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function parseEmailsFromText(raw) {
  const s = String(raw || "");
  const parts = s
    .split(/[\n,;]+/g)
    .map((x) => normalizeEmail(x))
    .filter(Boolean);
  const out = [];
  const seen = new Set();
  for (const e of parts) {
    if (!isValidEmail(e)) continue;
    if (seen.has(e)) continue;
    seen.add(e);
    out.push(e);
  }
  return out;
}

// -------------------------
// ATS scoring (local heuristic)
// -------------------------
const STOPWORDS = new Set(
  [
    "a","an","the","and","or","but","if","then","else","when","while","to","of","in","on","for","from","with","without",
    "is","are","was","were","be","been","being","as","at","by","we","you","your","our","they","them","their","i","me","my",
    "this","that","these","those","it","its","can","could","should","would","will","may","might","must","also","etc",
    "role","responsibilities","requirements","preferred","experience","years","year","skills","ability","strong","good",
    "work","working","team","teams","communication","develop","development","building","build","design","implement","using",
  ].map((x) => x.toLowerCase()),
);

function tokenize(text) {
  const t = String(text || "").toLowerCase();
  // keep letters/numbers and common tech symbols
  const raw = t.match(/[a-z0-9][a-z0-9+.#/-]{1,}/g) || [];
  return raw
    .map((w) => w.replace(/^[^a-z0-9]+|[^a-z0-9]+$/g, ""))
    .filter((w) => w.length >= 3 && !STOPWORDS.has(w));
}

function topKeywordsFromJd(jd, { limit = 30 } = {}) {
  const tokens = tokenize(jd);
  const freq = new Map();
  for (const t of tokens) freq.set(t, (freq.get(t) || 0) + 1);
  const sorted = Array.from(freq.entries())
    .sort((a, b) => b[1] - a[1])
    .map(([k]) => k);
  return sorted.slice(0, limit);
}

function normalizeWhitespace(s) {
  return String(s || "").replace(/\s+/g, " ").trim();
}

function hasSection(resumeText, name) {
  const v = String(resumeText || "").toLowerCase();
  return v.includes(name);
}

function structureScore(resumeText) {
  let score = 0;
  const v = String(resumeText || "");
  const lower = v.toLowerCase();
  const sections = ["summary", "experience", "skills", "projects", "education"];
  for (const s of sections) if (lower.includes(s)) score += 6;
  // numeric impact
  if (/\b\d{1,3}%\b/.test(v) || /\b\d+\b/.test(v)) score += 6;
  // bullet points
  if (/[\n\r]\s*[-•*]\s+/.test(v)) score += 6;
  return Math.min(30, score);
}

function computeAts({ resumeText, jdText }) {
  const jd = String(jdText || "").trim();
  const resume = String(resumeText || "").trim();
  const kws = topKeywordsFromJd(jd, { limit: 30 });
  const resumeTokens = new Set(tokenize(resume));

  const matched = [];
  const missing = [];
  for (const k of kws) {
    if (resumeTokens.has(k) || resume.toLowerCase().includes(k)) matched.push(k);
    else missing.push(k);
  }

  const ratio = kws.length ? matched.length / kws.length : 0;
  const matchScore = Math.round(ratio * 70);
  const struct = structureScore(resume);
  const score = Math.max(0, Math.min(100, matchScore + struct));

  const suggestions = [];
  if (missing.length) {
    suggestions.push(`Add these missing keywords naturally in Skills/Experience: ${missing.slice(0, 10).join(", ")}`);
  }
  if (!hasSection(resume, "skills")) suggestions.push("Add a dedicated SKILLS section with the exact tech from the JD.");
  if (!hasSection(resume, "experience")) suggestions.push("Add/expand EXPERIENCE with JD-aligned bullet points.");
  if (!hasSection(resume, "projects")) suggestions.push("Add 1–2 PROJECTS relevant to the JD and include tech stack.");
  suggestions.push("Quantify impact: add metrics (%, time saved, latency reduced, users, revenue).");
  suggestions.push("Match job title in your summary headline and tailor first 3 bullets to JD requirements.");

  const keyPoints = [
    ...missing.slice(0, 8).map((k) => `Include “${k}” in a relevant bullet (project/experience) with proof/impact.`),
    "Add 2–3 strong JD-aligned achievements with numbers.",
    "Ensure your most relevant experience appears in the first half of the resume.",
  ].slice(0, 12);

  return {
    score,
    matchedKeywords: matched,
    missingKeywords: missing,
    suggestions,
    keyPoints,
    meta: {
      keywordCount: kws.length,
      matchScore,
      structureScore: struct,
    },
  };
}

function commandExists(cmd) {
  const r = spawnSync("which", [cmd], { stdio: "ignore" });
  return r.status === 0;
}

function extractDocxText(docxPath) {
  // Use system unzip to extract word/document.xml (works on macOS + many linux distros)
  if (!commandExists("unzip")) {
    throw new Error("DOCX parsing needs 'unzip' command. Paste resume text instead.");
  }
  const r = spawnSync("unzip", ["-p", docxPath, "word/document.xml"], { encoding: "utf8", maxBuffer: 20 * 1024 * 1024 });
  if (r.status !== 0) throw new Error("Failed to read DOCX (word/document.xml).");
  const xml = String(r.stdout || "");
  // Remove XML tags, keep spaces
  return normalizeWhitespace(xml.replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " "));
}

function extractPdfText(pdfPath) {
  if (!commandExists("pdftotext")) {
    throw new Error("PDF parsing needs 'pdftotext'. Upload DOCX or paste resume text.");
  }
  const outTxt = `${pdfPath}.txt`;
  const r = spawnSync("pdftotext", [pdfPath, outTxt], { encoding: "utf8" });
  if (r.status !== 0) throw new Error("Failed to parse PDF. Upload DOCX or paste resume text.");
  const txt = fs.readFileSync(outTxt, "utf8");
  fs.unlinkSync(outTxt);
  return normalizeWhitespace(txt);
}

function buildOptimizedResumeText({ originalText, missingKeywords, keyPoints }) {
  const base = String(originalText || "").trim();
  const missing = Array.isArray(missingKeywords) ? missingKeywords : [];
  const points = Array.isArray(keyPoints) ? keyPoints : [];

  const section = [
    "",
    "==============================",
    "ATS OPTIMIZATION (Auto-added)",
    "==============================",
    "",
    missing.length ? `Target keywords to include: ${missing.slice(0, 20).join(", ")}` : "",
    "",
    points.length ? "Key points to add/update:" : "",
    ...points.slice(0, 10).map((p) => `- ${p}`),
    "",
  ]
    .filter((x) => x !== "")
    .join("\n");

  // Also append keywords as a simple "Skills Addendum" line to improve matching.
  const keywordLine = missing.length
    ? `\n\nSkills Addendum: ${missing.slice(0, 25).join(", ")}\n`
    : "\n";

  return `${base}${section}${keywordLine}`.trim() + "\n";
}

function pdfEscape(s) {
  return String(s || "").replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
}

function textToSimplePdfBuffer(text) {
  // Minimal single-file PDF (Helvetica). Good enough for download/printing.
  const lines = String(text || "")
    .replace(/\r/g, "")
    .split("\n")
    .flatMap((l) => {
      // crude wrap at ~95 chars
      const out = [];
      let s = l;
      while (s.length > 95) {
        out.push(s.slice(0, 95));
        s = s.slice(95);
      }
      out.push(s);
      return out;
    });

  const pageHeight = 792; // 11in * 72
  const pageWidth = 612; // 8.5in * 72
  const margin = 48;
  const lineHeight = 12;
  const usable = pageHeight - margin * 2;
  const linesPerPage = Math.max(1, Math.floor(usable / lineHeight));
  const pages = [];
  for (let i = 0; i < lines.length; i += linesPerPage) {
    pages.push(lines.slice(i, i + linesPerPage));
  }

  const objects = [];
  const offsets = [];
  const addObj = (s) => {
    offsets.push(null);
    objects.push(s);
    return objects.length; // 1-based obj number
  };

  const fontObj = addObj("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");

  const pageObjs = [];
  const contentObjs = [];

  for (const pLines of pages) {
    let y = pageHeight - margin;
    const contentLines = [];
    contentLines.push("BT");
    contentLines.push("/F1 10 Tf");
    contentLines.push("1 0 0 1 0 0 Tm");
    for (const l of pLines) {
      contentLines.push(`${margin} ${y} Td`);
      contentLines.push(`(${pdfEscape(l)}) Tj`);
      contentLines.push(`${-margin} 0 Td`);
      y -= lineHeight;
    }
    contentLines.push("ET");
    const stream = contentLines.join("\n");
    const contentObj = addObj(`<< /Length ${Buffer.byteLength(stream, "utf8")} >>\nstream\n${stream}\nendstream`);
    contentObjs.push(contentObj);
  }

  const pagesKids = [];
  for (let i = 0; i < pages.length; i++) {
    const contentObj = contentObjs[i];
    const pageObj = addObj(
      `<< /Type /Page /Parent 0 0 R /MediaBox [0 0 ${pageWidth} ${pageHeight}] /Resources << /Font << /F1 ${fontObj} 0 R >> >> /Contents ${contentObj} 0 R >>`,
    );
    pageObjs.push(pageObj);
    pagesKids.push(`${pageObj} 0 R`);
  }

  const pagesObjNum = addObj(`<< /Type /Pages /Kids [${pagesKids.join(" ")}] /Count ${pagesKids.length} >>`);

  // Patch Parent refs (replace "0 0 R" with pagesObjNum)
  for (const objNum of pageObjs) {
    const idx = objNum - 1;
    objects[idx] = objects[idx].replace("/Parent 0 0 R", `/Parent ${pagesObjNum} 0 R`);
  }

  const catalogObj = addObj(`<< /Type /Catalog /Pages ${pagesObjNum} 0 R >>`);

  let pdf = "%PDF-1.4\n";
  for (let i = 0; i < objects.length; i++) {
    offsets[i] = Buffer.byteLength(pdf, "utf8");
    pdf += `${i + 1} 0 obj\n${objects[i]}\nendobj\n`;
  }
  const xrefStart = Buffer.byteLength(pdf, "utf8");
  pdf += `xref\n0 ${objects.length + 1}\n`;
  pdf += "0000000000 65535 f \n";
  for (let i = 0; i < offsets.length; i++) {
    const off = String(offsets[i]).padStart(10, "0");
    pdf += `${off} 00000 n \n`;
  }
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root ${catalogObj} 0 R >>\nstartxref\n${xrefStart}\n%%EOF\n`;
  return Buffer.from(pdf, "utf8");
}

let lastOptimizedPdfBuffer = null;

function buildSuggestionsPageText({ missingKeywords, keyPoints }) {
  const missing = Array.isArray(missingKeywords) ? missingKeywords : [];
  const points = Array.isArray(keyPoints) ? keyPoints : [];
  const lines = [
    "ATS Optimization Suggestions",
    "",
    missing.length ? `Missing keywords: ${missing.slice(0, 30).join(", ")}` : "Missing keywords: —",
    "",
    "Key points to add/update:",
    ...(points.length ? points.slice(0, 18).map((p) => `- ${p}`) : ["- —"]),
    "",
    "Note: This page is auto-generated. Edit your original resume accordingly.",
    "",
  ];
  return lines.join("\n");
}

function pdfUniteIfAvailable(inputPdfPath, appendPdfBuffer) {
  if (!commandExists("pdfunite")) return null;
  const tmpDir = os.tmpdir();
  const appendPath = path.join(tmpDir, `job-mailer-ats-append-${Date.now()}.pdf`);
  const outPath = path.join(tmpDir, `job-mailer-ats-out-${Date.now()}.pdf`);
  fs.writeFileSync(appendPath, appendPdfBuffer);
  const r = spawnSync("pdfunite", [inputPdfPath, appendPath, outPath], { encoding: "utf8" });
  try {
    if (r.status !== 0) return null;
    const buf = fs.readFileSync(outPath);
    return buf;
  } finally {
    fs.promises.unlink(appendPath).catch(() => {});
    fs.promises.unlink(outPath).catch(() => {});
  }
}

app.get("/api/ats-optimized.pdf", (_req, res) => {
  if (!lastOptimizedPdfBuffer) {
    return res.status(404).json({ ok: false, error: "No optimized resume generated yet." });
  }
  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", 'attachment; filename="optimized-resume.pdf"');
  return res.send(lastOptimizedPdfBuffer);
});

app.post("/api/ats-optimize", upload.single("resume"), async (req, res) => {
  try {
    const jd = String(req.body.jd || "").trim();
    const resumeTextFallback = String(req.body.resumeText || "").trim();
    if (!jd) return res.status(400).json({ ok: false, error: "Job description is required." });

    let resumeText = resumeTextFallback;
    const filePath = req.file?.path || "";
    const orig = String(req.file?.originalname || "").toLowerCase();

    if (!resumeText) {
      if (!filePath) return res.status(400).json({ ok: false, error: "Resume file or resume text is required." });
      if (orig.endsWith(".docx")) resumeText = extractDocxText(filePath);
      else if (orig.endsWith(".pdf")) resumeText = extractPdfText(filePath);
      else return res.status(400).json({ ok: false, error: "Unsupported resume type. Upload PDF/DOCX or paste text." });
    }

    const maxIters = 4;
    let currentText = resumeText;
    let current = computeAts({ resumeText: currentText, jdText: jd });
    let iters = 0;

    while (current.score < 90 && iters < maxIters) {
      iters += 1;
      currentText = buildOptimizedResumeText({
        originalText: currentText,
        missingKeywords: current.missingKeywords,
        keyPoints: current.keyPoints,
      });
      current = computeAts({ resumeText: currentText, jdText: jd });
    }

    const ready = current.score >= 90;
    const suggestionsText = buildSuggestionsPageText({
      missingKeywords: current.missingKeywords,
      keyPoints: current.keyPoints,
    });
    const suggestionsPdf = textToSimplePdfBuffer(suggestionsText);

    // If the user uploaded a PDF, preserve their original resume pages and append suggestions as last page.
    if (filePath && orig.endsWith(".pdf")) {
      const united = pdfUniteIfAvailable(filePath, suggestionsPdf);
      if (united) {
        lastOptimizedPdfBuffer = united;
      } else {
        // Fallback: return original PDF as-is (still better than "only email"), and keep suggestions in UI.
        lastOptimizedPdfBuffer = fs.readFileSync(filePath);
      }
    } else {
      // For DOCX or pasted text, generate a simple PDF from optimized extracted text.
      lastOptimizedPdfBuffer = textToSimplePdfBuffer(currentText);
    }

    return res.json({
      ok: true,
      result: current,
      optimized: {
        iterations: iters,
        ready,
        downloadUrl: "/api/ats-optimized.pdf",
        note: filePath && orig.endsWith(".pdf")
          ? (commandExists("pdfunite")
              ? "Downloaded PDF preserves your original resume and appends suggestions as the last page."
              : "Downloaded PDF preserves your original resume (suggestions could not be appended automatically on this machine).")
          : "Downloaded PDF is generated from extracted resume text + added suggestions.",
      },
    });
  } catch (e) {
    return res.status(500).json({ ok: false, error: String(e?.message || e) });
  } finally {
    if (req.file?.path) fs.promises.unlink(req.file.path).catch(() => {});
  }
});

app.post("/api/ats-score", upload.single("resume"), async (req, res) => {
  try {
    const jd = String(req.body.jd || "").trim();
    const resumeTextFallback = String(req.body.resumeText || "").trim();
    if (!jd) return res.status(400).json({ ok: false, error: "Job description is required." });

    let resumeText = resumeTextFallback;
    const filePath = req.file?.path || "";
    const orig = String(req.file?.originalname || "").toLowerCase();

    if (!resumeText) {
      if (!filePath) return res.status(400).json({ ok: false, error: "Resume file or resume text is required." });
      if (orig.endsWith(".docx")) resumeText = extractDocxText(filePath);
      else if (orig.endsWith(".pdf")) resumeText = extractPdfText(filePath);
      else return res.status(400).json({ ok: false, error: "Unsupported resume type. Upload PDF/DOCX or paste text." });
    }

    const result = computeAts({ resumeText, jdText: jd });
    const note =
      result.meta && result.meta.keywordCount
        ? `Keywords checked: ${result.meta.keywordCount} • Match ${result.meta.matchScore}/70 • Structure ${result.meta.structureScore}/30`
        : "";
    result.meta = { ...(result.meta || {}), note };
    return res.json({ ok: true, result });
  } catch (e) {
    return res.status(500).json({ ok: false, error: String(e?.message || e) });
  } finally {
    if (req.file?.path) fs.promises.unlink(req.file.path).catch(() => {});
  }
});

// -------------------------
// UI Defaults (stored locally)
// -------------------------
const SETTINGS_PATH = path.resolve(config.paths.root, "data", "ui-settings.json");
const DEFAULT_RESUME_PATH = path.resolve(config.paths.root, "data", "default-resume.pdf");

function loadUiSettings() {
  return readJson(SETTINGS_PATH, {});
}

function saveUiSettings(next) {
  writeJsonAtomic(SETTINGS_PATH, next || {});
}

function getEffectiveSettings() {
  const s = loadUiSettings();
  return {
    smtp: {
      host: String(s.smtpHost || config.smtp.host || "").trim(),
      port: Number(s.smtpPort || config.smtp.port || 0) || config.smtp.port,
      secure: s.smtpSecure === undefined ? config.smtp.secure : Boolean(s.smtpSecure),
      user: String(s.smtpUser || config.smtp.user || "").trim(),
      pass: String(s.smtpPass || config.smtp.pass || "").trim(),
    },
    from: {
      email: String(s.fromEmail || config.from.email || s.smtpUser || "").trim(),
      name: String(s.fromName || config.from.name || "").trim(),
    },
    subject: String(s.subject || config.content.subject || "").trim(),
    defaultBody: String(s.defaultBody || "").trim(),
    resumePath: fs.existsSync(DEFAULT_RESUME_PATH) ? DEFAULT_RESUME_PATH : config.paths.resumePath,
    meta: {
      smtpPassSet: Boolean(String(s.smtpPass || "").trim()),
      resumeSet: fs.existsSync(DEFAULT_RESUME_PATH),
    },
  };
}

app.get("/api/settings", (_req, res) => {
  const raw = loadUiSettings();
  const eff = getEffectiveSettings();
  return res.json({
    ok: true,
    settings: {
      smtpHost: String(raw.smtpHost || config.smtp.host || ""),
      smtpPort: raw.smtpPort ?? config.smtp.port,
      smtpSecure: raw.smtpSecure ?? config.smtp.secure,
      smtpUser: String(raw.smtpUser || config.smtp.user || ""),
      // do not return the password
      smtpPassSet: eff.meta.smtpPassSet,
      fromEmail: String(raw.fromEmail || config.from.email || ""),
      fromName: String(raw.fromName || config.from.name || ""),
      subject: String(raw.subject || config.content.subject || ""),
      defaultBody: String(raw.defaultBody || ""),
      resumeSet: eff.meta.resumeSet,
    },
  });
});

app.post("/api/settings", (req, res) => {
  try {
    const prev = loadUiSettings();
    const smtpPassIncoming = String(req.body.smtpPass || "");
    const next = {
      smtpHost: String(req.body.smtpHost || prev.smtpHost || "").trim(),
      smtpPort:
        req.body.smtpPort === null || req.body.smtpPort === undefined || req.body.smtpPort === ""
          ? prev.smtpPort
          : Number(req.body.smtpPort),
      smtpSecure:
        req.body.smtpSecure === undefined || req.body.smtpSecure === null
          ? prev.smtpSecure
          : Boolean(req.body.smtpSecure),
      smtpUser: String(req.body.smtpUser || prev.smtpUser || "").trim(),
      smtpPass: smtpPassIncoming.trim() ? smtpPassIncoming.trim() : String(prev.smtpPass || ""),
      fromEmail: String(req.body.fromEmail || prev.fromEmail || "").trim(),
      fromName: String(req.body.fromName || prev.fromName || "").trim(),
      subject: String(req.body.subject || prev.subject || "").trim(),
      defaultBody: String(req.body.defaultBody || prev.defaultBody || "").trim(),
    };
    saveUiSettings(next);
    return res.json({ ok: true });
  } catch (e) {
    return res.status(500).json({ ok: false, error: String(e?.message || e) });
  }
});

app.post("/api/settings/resume", upload.single("resume"), async (req, res) => {
  try {
    if (!req.file?.path) return res.status(400).json({ ok: false, error: "Resume file is required." });
    ensureDir(path.dirname(DEFAULT_RESUME_PATH));
    await fs.promises.copyFile(req.file.path, DEFAULT_RESUME_PATH);
    return res.json({ ok: true });
  } catch (e) {
    return res.status(500).json({ ok: false, error: String(e?.message || e) });
  } finally {
    if (req.file?.path) fs.promises.unlink(req.file.path).catch(() => {});
  }
});

function normalizeDomain(domain) {
  const d = String(domain || "")
    .trim()
    .toLowerCase()
    .replace(/^https?:\/\//, "")
    .replace(/^www\./, "")
    .replace(/\/.*$/, "");
  return d;
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function bodyToHtml(bodyText) {
  // Basic newline -> <br/> conversion for a simple custom body.
  return `<div style="white-space:pre-wrap;font-family:system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial;">${escapeHtml(
    bodyText,
  )}</div>`;
}

function buildOverriddenEmail({ recipientName, recipientEmail, bodyText }) {
  const rawName = String(recipientName || "").trim();
  const firstName = rawName.replace(/[(),]/g, " ").trim().split(/\s+/)[0] || "";
  const greetingName = firstName || rawName || "Hiring Team";

  const rawBody = String(bodyText || "").trim();
  const bodyHasSignature = (() => {
    if (!rawBody) return false;
    const b = rawBody.toLowerCase();
    return /warm\s+regards/.test(b) || /regards\s*,/.test(b) || /shubham\s+pawar/.test(b);
  })();

  const signatureText = [
    "Warm regards,",
    "Shubham Pawar",
    "MERN Stack Developer | Software Engineer",
    "Immediate Joiner",
  ].join("\n");

  const textParts = [`Hi ${greetingName},`, "", rawBody];
  if (!bodyHasSignature) textParts.push("", signatureText, "");
  else textParts.push("");
  const text = textParts.join("\n");

  const html = `
    <p>Hi ${escapeHtml(greetingName)},</p>
    ${bodyToHtml(rawBody)}
    ${
      bodyHasSignature
        ? ""
        : `<p>
            Warm regards,<br />
            Shubham Pawar<br />
            MERN Stack Developer | Software Engineer<br />
            Immediate Joiner
          </p>`
    }
  `.trim();

  return { text, html };
}

function pickFirstNonEmpty(...vals) {
  for (const v of vals) {
    const s = String(v ?? "").trim();
    if (s) return s;
  }
  return "";
}

function parseRecipientsFromXlsx(filePath) {
  const wb = XLSX.readFile(filePath, { cellDates: true });
  const sheetName = wb.SheetNames[0];
  if (!sheetName) return [];
  const ws = wb.Sheets[sheetName];

  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  // expected columns (case-insensitive):
  // - email / mail
  // - recipient name / name
  // - subject
  // - body
  const out = [];
  for (const r of rows) {
    const email = normalizeEmail(
      pickFirstNonEmpty(
        r.email,
        r.Email,
        r.EMAIL,
        r.mail,
        r.Mail,
        r.MAIL,
        r["email id"],
        r["Email Id"],
        r["EMAIL ID"],
        r["mail id"],
        r["Mail Id"],
        r["MAIL ID"],
        r["email address"],
        r["Email Address"],
        r["EMAIL ADDRESS"],
      ),
    );
    if (!email || !isValidEmail(email)) continue;
    const name = pickFirstNonEmpty(
      r["recipient name"],
      r["Recipient Name"],
      r["RECIPIENT NAME"],
      r["receipnt name"], // common typo
      r["Receipnt Name"],
      r["RECEIPNT NAME"],
      r.name,
      r.Name,
      r.NAME,
    ).trim();
    const subject = pickFirstNonEmpty(r.subject, r.Subject, r.SUBJECT).trim();
    const body = pickFirstNonEmpty(r.body, r.Body, r.BODY).trim();
    out.push({ email, name, subject, body });
  }

  // de-dupe by email (keep first non-empty values)
  const seen = new Map();
  for (const row of out) {
    if (!seen.has(row.email)) {
      seen.set(row.email, row);
      continue;
    }
    const existing = seen.get(row.email);
    if (!existing.name && row.name) existing.name = row.name;
    if (!existing.subject && row.subject) existing.subject = row.subject;
    if (!existing.body && row.body) existing.body = row.body;
  }
  return Array.from(seen.values());
}

function buildTemplateWorkbookBuffer() {
  const header = [["email", "recipient name", "subject", "body"]];
  const sample = [
    ["hr@company.com", "Hiring Team", "", ""],
    ["recruiter@company.com", "Priya", "Application for MERN Stack Developer Role — Immediate Joiner | 3 Yrs Experience", ""],
  ];
  const aoa = header.concat(sample);
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "recipients");
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

const UI_DIR = path.resolve(__dirname, "..", "..", "frontend", "public");
const LOGIN_PATH = path.resolve(UI_DIR, "login.html");

// -------------------------
// Auth (simple local login)
// -------------------------
const AUTH_USER = String(process.env.UI_AUTH_USER || "").trim();
const AUTH_PASS = String(process.env.UI_AUTH_PASS || "").trim();
const AUTH_ENABLED = Boolean(AUTH_USER && AUTH_PASS);
const COOKIE_NAME = "jm_sid";
const sessions = new Map(); // sid -> { createdAt, expiresAt }
const SESSION_TTL_MS = 1000 * 60 * 60 * 12; // 12 hours

function parseCookies(req) {
  const header = req.headers.cookie || "";
  const out = {};
  for (const part of header.split(";")) {
    const idx = part.indexOf("=");
    if (idx === -1) continue;
    const k = part.slice(0, idx).trim();
    const v = part.slice(idx + 1).trim();
    if (!k) continue;
    out[k] = decodeURIComponent(v);
  }
  return out;
}

function createSession() {
  const sid = crypto.randomBytes(24).toString("hex");
  const now = Date.now();
  sessions.set(sid, { createdAt: now, expiresAt: now + SESSION_TTL_MS });
  return sid;
}

function isAuthenticated(req) {
  if (!AUTH_ENABLED) return true;
  const cookies = parseCookies(req);
  const sid = cookies[COOKIE_NAME];
  if (!sid) return false;
  const s = sessions.get(sid);
  if (!s) return false;
  if (Date.now() > s.expiresAt) {
    sessions.delete(sid);
    return false;
  }
  return true;
}

function requireAuth(req, res, next) {
  if (isAuthenticated(req)) return next();
  const isApi = req.path.startsWith("/api/");
  if (isApi) return res.status(401).json({ ok: false, error: "Unauthorized. Please login." });
  return res.redirect("/login");
}

app.get("/health", (_req, res) => res.json({ ok: true }));

app.get("/login", (_req, res) => {
  if (!AUTH_ENABLED) {
    return res
      .status(200)
      .send(
        "Auth is disabled. Set UI_AUTH_USER and UI_AUTH_PASS in .env (or env.example) to enable login.",
      );
  }
  return res.sendFile(LOGIN_PATH);
});

app.post("/api/login", (req, res) => {
  if (!AUTH_ENABLED) return res.json({ ok: true, authEnabled: false });
  const user = String(req.body.user || "").trim();
  const pass = String(req.body.pass || "").trim();
  if (user !== AUTH_USER || pass !== AUTH_PASS) {
    return res.status(401).json({ ok: false, error: "Invalid username or password." });
  }
  const sid = createSession();
  res.setHeader(
    "Set-Cookie",
    `${COOKIE_NAME}=${encodeURIComponent(
      sid,
    )}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${Math.floor(SESSION_TTL_MS / 1000)}`,
  );
  return res.json({ ok: true, authEnabled: true });
});

app.post("/api/logout", (req, res) => {
  const cookies = parseCookies(req);
  const sid = cookies[COOKIE_NAME];
  if (sid) sessions.delete(sid);
  res.setHeader("Set-Cookie", `${COOKIE_NAME}=; Path=/; Max-Age=0; SameSite=Lax`);
  return res.json({ ok: true });
});

// Protect everything (UI + API) except health + login endpoints.
app.use((req, res, next) => {
  if (!AUTH_ENABLED) return next();
  if (req.path === "/health") return next();
  if (req.path === "/login") return next();
  if (req.path === "/api/login") return next();
  return requireAuth(req, res, next);
});

// Serve UI (protected if auth enabled)
app.use(express.static(UI_DIR));

// -------------------------
// HR / Talent lookup (optional)
// -------------------------
const HUNTER_API_KEY = String(process.env.HUNTER_API_KEY || "").trim();
const HR_PROVIDER_DEFAULT = String(process.env.HR_PROVIDER || "hunter").trim().toLowerCase();

// Apollo.io (people database) integration (requires Apollo.io API key)
const APOLLO_API_KEY = String(process.env.APOLLO_API_KEY || "").trim();
const APOLLO_BASE_URL = String(process.env.APOLLO_BASE_URL || "https://api.apollo.io").trim();
const APOLLO_ENDPOINT = String(process.env.APOLLO_ENDPOINT || "/v1/mixed_people/search").trim();
const APOLLO_REVEAL_PHONE_NUMBER = ["1", "true", "yes", "y", "on"].includes(
  String(process.env.APOLLO_REVEAL_PHONE_NUMBER || "").trim().toLowerCase(),
);

function looksLikeApolloGraphOSKey(key) {
  const k = String(key || "").trim();
  return k.startsWith("service:");
}

// Provider status for UI (no secrets returned)
app.get("/api/provider-status", (_req, res) => {
  res.json({
    ok: true,
    providers: {
      hunter: { configured: Boolean(HUNTER_API_KEY) },
      apollo: {
        configured: Boolean(APOLLO_API_KEY),
        looksLikeGraphOS: looksLikeApolloGraphOSKey(APOLLO_API_KEY),
      },
    },
  });
});

// -------------------------
// Company names (saved list + live suggestions)
// -------------------------
const COMPANIES_PATH = path.resolve(config.paths.root, "data", "companies.json");

function loadCompanyNames() {
  const raw = readJson(COMPANIES_PATH, []);
  const arr = Array.isArray(raw) ? raw : raw?.companies;
  const out = [];
  const seen = new Set();
  for (const x of Array.isArray(arr) ? arr : []) {
    const s = String(x || "").trim();
    if (!s) continue;
    const k = s.toLowerCase();
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(s);
  }
  return out.sort((a, b) => a.localeCompare(b));
}

function rememberCompanyName(name) {
  const n = String(name || "").trim();
  if (!n) return;
  ensureDir(path.dirname(COMPANIES_PATH));
  const next = loadCompanyNames();
  const k = n.toLowerCase();
  if (!next.some((x) => x.toLowerCase() === k)) next.push(n);
  next.sort((a, b) => a.localeCompare(b));
  writeJsonAtomic(COMPANIES_PATH, next);
}

app.get("/api/company-names", (_req, res) => {
  return res.json({ ok: true, companies: loadCompanyNames() });
});

app.get("/api/company-suggest", async (req, res) => {
  try {
    const q = String(req.query.query || req.query.q || "").trim();
    if (!q) return res.json({ ok: true, companies: [] });
    const url = `https://autocomplete.clearbit.com/v1/companies/suggest?query=${encodeURIComponent(q)}`;
    const r = await fetch(url, { headers: { Accept: "application/json" } });
    if (!r.ok) return res.json({ ok: true, companies: [] });
    const arr = await r.json().catch(() => []);
    const names = (Array.isArray(arr) ? arr : [])
      .map((x) => String(x?.name || "").trim())
      .filter(Boolean)
      .slice(0, 20);
    return res.json({ ok: true, companies: names });
  } catch (e) {
    return res.json({ ok: true, companies: [] });
  }
});

function isRecruitingRole(s) {
  const v = String(s || "").toLowerCase();
  return (
    v.includes("talent") ||
    v.includes("recruit") ||
    v.includes("hr") ||
    v.includes("human resources") ||
    v.includes("people ops") ||
    v.includes("people operations")
  );
}

async function resolveDomainFromCompany(company) {
  const raw = String(company || "").trim();
  if (!raw) return null;

  function buildQueryVariants(s) {
    const base = String(s || "")
      .trim()
      .replace(/\s+/g, " ")
      .replace(/[(),]/g, " ");

    const tokens = base
      .split(/\s+/g)
      .map((t) => t.trim())
      .filter(Boolean);

    const STOP = new Set([
      "pvt",
      "pvt.",
      "ltd",
      "ltd.",
      "limited",
      "private",
      "inc",
      "inc.",
      "llc",
      "llp",
      "co",
      "co.",
      "company",
      "technologies",
      "technology",
      "solutions",
      "services",
      "systems",
      "group",
      "corp",
      "corp.",
      "corporation",
    ]);

    const strippedTokens = tokens.filter((t) => !STOP.has(t.toLowerCase()));
    const variants = [
      base,
      strippedTokens.join(" "),
      strippedTokens.slice(0, 2).join(" "),
      strippedTokens.slice(0, 1).join(" "),
      tokens.slice(0, 2).join(" "),
      tokens.slice(0, 1).join(" "),
    ]
      .map((x) => String(x || "").trim())
      .filter(Boolean);

    // de-dupe while keeping order
    const out = [];
    const seen = new Set();
    for (const v of variants) {
      const k = v.toLowerCase();
      if (seen.has(k)) continue;
      seen.add(k);
      out.push(v);
    }
    return out.slice(0, 4); // keep it tight (avoid too many network calls)
  }

  const queries = buildQueryVariants(raw);
  for (const q of queries) {
    const url = `https://autocomplete.clearbit.com/v1/companies/suggest?query=${encodeURIComponent(q)}`;
    const r = await fetch(url, { headers: { Accept: "application/json" } });
    if (!r.ok) continue;
    const arr = await r.json().catch(() => []);
    const first = Array.isArray(arr) ? arr[0] : null;
    const domain = normalizeDomain(first?.domain || first?.website || "");
    if (domain) return domain;
  }

  return null;
}

// (Company Finder endpoints removed)

async function hunterDomainSearch(domain) {
  if (!HUNTER_API_KEY) {
    throw new Error("HUNTER_API_KEY is not set on the server.");
  }
  const d = normalizeDomain(domain);
  if (!d) throw new Error("Valid domain is required (example: company.com)");

  const url = `https://api.hunter.io/v2/domain-search?domain=${encodeURIComponent(
    d,
  )}&api_key=${encodeURIComponent(HUNTER_API_KEY)}`;
  const r = await fetch(url);
  const payload = await r.json().catch(() => null);
  if (!r.ok) {
    const msg = payload?.errors?.[0]?.details || payload?.errors?.[0]?.message || "Hunter request failed";
    throw new Error(msg);
  }
  return payload;
}

function recruitingTitleKeywords() {
  return [
    "Talent Acquisition",
    "Recruiter",
    "Recruitment",
    "HR",
    "Human Resources",
    "People Operations",
    "People Ops",
  ];
}

async function apolloPeopleSearch(domain) {
  if (!APOLLO_API_KEY) {
    throw new Error("APOLLO_API_KEY is not set on the server.");
  }
  if (looksLikeApolloGraphOSKey(APOLLO_API_KEY)) {
    throw new Error(
      "APOLLO_API_KEY looks like an Apollo GraphOS (service:...) key. HR Finder Apollo needs an Apollo.io API key.",
    );
  }
  const d = normalizeDomain(domain);
  if (!d) throw new Error("Valid domain is required (example: company.com)");

  // NOTE: Apollo.io APIs and response shapes can vary by plan and may change.
  // This is implemented as a best-effort integration; if your Apollo account uses
  // a different endpoint/shape, set APOLLO_ENDPOINT/APOLLO_BASE_URL and we can adjust mapping.
  const url = `${APOLLO_BASE_URL}${APOLLO_ENDPOINT}`;
  const body = {
    q_organization_domains: d,
    page: 1,
    per_page: 25,
    person_titles: recruitingTitleKeywords(),
    ...(APOLLO_REVEAL_PHONE_NUMBER ? { reveal_phone_number: true } : {}),
  };

  const r = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
      "X-Api-Key": APOLLO_API_KEY,
    },
    body: JSON.stringify(body),
  });
  const payload = await r.json().catch(() => null);
  if (!r.ok) {
    if (r.status === 401) {
      throw new Error(
        "Apollo request failed (401). This usually means the API key is invalid or not an Apollo.io API key.",
      );
    }
    const msg =
      payload?.error ||
      payload?.message ||
      payload?.errors?.[0] ||
      `Apollo request failed (${r.status})`;
    throw new Error(String(msg));
  }
  return payload;
}

function normalizePhone(s) {
  const v = String(s || "").trim();
  if (!v) return "";
  return v;
}

function extractApolloPhone(p) {
  // Apollo response shapes vary by endpoint/plan. We attempt common fields.
  const direct =
    p?.phone_number ||
    p?.phoneNumber ||
    p?.mobile_phone ||
    p?.mobilePhone ||
    p?.mobile_phone_number ||
    p?.mobilePhoneNumber ||
    "";
  const d = normalizePhone(direct);
  if (d) return d;

  const arr = p?.phone_numbers || p?.phoneNumbers || p?.phones || p?.phone_numbers_raw || null;
  if (Array.isArray(arr) && arr.length) {
    for (const x of arr) {
      const cand =
        x?.raw_number ||
        x?.rawNumber ||
        x?.sanitized_number ||
        x?.sanitizedNumber ||
        x?.number ||
        x?.value ||
        x;
      const n = normalizePhone(cand);
      if (n) return n;
    }
  }

  const contact = p?.contact || p?.person || null;
  if (contact) return extractApolloPhone(contact);

  return "";
}

app.get("/api/hr-lookup", async (req, res) => {
  try {
    const company = String(req.query.company || "").trim();
    const domainInput = normalizeDomain(req.query.domain || "");
    const provider = String(req.query.provider || HR_PROVIDER_DEFAULT || "hunter")
      .trim()
      .toLowerCase();

    if (company) {
      // Save for dropdown auto-complete in the UI (best-effort).
      try {
        rememberCompanyName(company);
      } catch {}
    }

    let domain = domainInput;
    if (!domain && company) {
      domain = (await resolveDomainFromCompany(company)) || "";
    }
    if (!domain) {
      return res.status(400).json({
        ok: false,
        error:
          "Provide a company domain (recommended) or a company name (domain will be auto-detected when possible).",
      });
    }

    if (provider === "apollo") {
      const apollo = await apolloPeopleSearch(domain);
      const people =
        apollo?.people || apollo?.contacts || apollo?.data?.people || apollo?.data?.contacts || [];

      const contacts = (Array.isArray(people) ? people : [])
        .map((p) => {
          const email = String(p?.email || p?.email_address || p?.emailAddress || "").trim().toLowerCase();
          if (!isValidEmail(email)) return null;
          const first = p?.first_name || p?.firstName || "";
          const last = p?.last_name || p?.lastName || "";
          const name = String(`${first} ${last}`.trim());
          const position = p?.title || p?.job_title || p?.position || "";
          const phone = extractApolloPhone(p);
          return {
            email,
            name,
            position: String(position || ""),
            seniority: String(p?.seniority || ""),
            phone: phone || null,
            confidence: null,
            source: "apollo",
          };
        })
        .filter(Boolean)
        .slice(0, 25);

      // If Apollo didn't return an org phone, we fall back to the first contact phone (if any)
      const org = apollo?.organization || apollo?.data?.organization || apollo?.account || apollo?.data?.account || null;
      const orgPhone =
        org?.phone_number ||
        org?.phone ||
        org?.phoneNumber ||
        org?.primary_phone ||
        org?.primaryPhone ||
        null;
      const fallbackPhone = contacts.find((c) => c.phone)?.phone || null;

      return res.json({
        ok: true,
        provider,
        company,
        domain,
        contacts,
        mode: "recruiting_only",
        phone: orgPhone ? String(orgPhone) : fallbackPhone ? String(fallbackPhone) : null,
      });
    }

    // Default: Hunter
    const hunter = await hunterDomainSearch(domain);
    const data = hunter?.data || {};
    const emails = Array.isArray(data.emails) ? data.emails : [];
    const org = data.organization || {};
    const phone =
      org.phone_number ||
      org.phone ||
      org.phoneNumber ||
      data.phone_number ||
      data.phone ||
      data.company_phone ||
      null;

    const allContacts = emails
      .filter((e) => isValidEmail(e?.value))
      .map((e) => {
        const firstName = e?.first_name || "";
        const lastName = e?.last_name || "";
        const fullName = `${firstName} ${lastName}`.trim();
        const position = e?.position || e?.department || "";
        const seniority = e?.seniority || "";
        return {
          email: String(e.value).toLowerCase(),
          name: fullName,
          position,
          seniority,
          confidence: e?.confidence ?? null,
          source: "hunter",
        };
      })
      .slice(0, 50);

    const recruitingContacts = allContacts
      .filter((c) => isRecruitingRole(`${c.position} ${c.seniority}`))
      .slice(0, 25);

    const contacts =
      recruitingContacts.length > 0 ? recruitingContacts : allContacts.slice(0, 25);

    return res.json({
      ok: true,
      provider: "hunter",
      company,
      domain,
      contacts,
      mode: recruitingContacts.length > 0 ? "recruiting_only" : "all_emails_fallback",
      phone: phone ? String(phone) : null,
    });
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e?.message || e) });
  }
});

// Downloadable Excel template
app.get("/api/template.xlsx", (_req, res) => {
  console.log("[ui] template download: /api/template.xlsx");
  const buf = buildTemplateWorkbookBuffer();
  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
  res.setHeader("Content-Disposition", 'attachment; filename="job-mailer-template.xlsx"');
  res.send(buf);
});

// Alias (in case you prefer a shorter URL)
app.get("/template.xlsx", (_req, res) => {
  console.log("[ui] template download: /template.xlsx");
  const buf = buildTemplateWorkbookBuffer();
  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
  res.setHeader("Content-Disposition", 'attachment; filename="job-mailer-template.xlsx"');
  res.send(buf);
});

// Download sent email log (Excel)
app.get("/api/sent.xlsx", (_req, res) => {
  console.log("[ui] sent log download: /api/sent.xlsx");
  const buf = getSentWorkbookBuffer(config.paths.sentXlsx);
  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
  res.setHeader("Content-Disposition", 'attachment; filename="job-mailer-sent.xlsx"');
  res.send(buf);
});

// Alias (in case you prefer a shorter URL)
app.get("/sent.xlsx", (_req, res) => {
  console.log("[ui] sent log download: /sent.xlsx");
  const buf = getSentWorkbookBuffer(config.paths.sentXlsx);
  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
  res.setHeader("Content-Disposition", 'attachment; filename="job-mailer-sent.xlsx"');
  res.send(buf);
});

app.post("/api/send", upload.single("resume"), async (req, res) => {
  const toEmail = normalizeEmail(req.body.email);
  const toName = String(req.body.name || "").trim();
  const subjectOverride = String(req.body.subject || "").trim();
  const bodyOverride = String(req.body.body || "").trim();

  if (!toEmail || !isValidEmail(toEmail)) {
    return res.status(400).json({ ok: false, error: "Valid email is required." });
  }

  const eff = getEffectiveSettings();
  const subject = subjectOverride || eff.subject || config.content.subject;

  // Decide content: override only if user provided body.
  let text;
  let html;
  const defaultBody = eff.defaultBody;
  const bodyToUse = bodyOverride || defaultBody;
  if (bodyToUse) {
    const overridden = buildOverriddenEmail({
      recipientName: toName,
      recipientEmail: toEmail,
      bodyText: bodyToUse,
    });
    text = overridden.text;
    html = overridden.html;
  } else {
    const built = buildEmail({
      recipientName: toName,
      recipientEmail: toEmail,
      subject,
    });
    text = built.text;
    html = built.html;
  }

  const resumePath = req.file?.path ? req.file.path : eff.resumePath;

  try {
    const transporter = await createTransporter({ smtp: eff.smtp, from: eff.from });

    // Reuse sender but with our custom text/html when bodyOverride is present.
    const info = bodyOverride
      ? await transporter.sendMail({
          from: eff.from.name ? `"${eff.from.name}" <${eff.from.email}>` : eff.from.email,
          to: toEmail,
          subject,
          text,
          html,
          attachments: [
            {
              filename: req.file?.originalname || path.basename(resumePath),
              path: resumePath,
            },
          ],
        })
      : await sendApplicationEmail({
          transporter,
          from: eff.from,
          toEmail,
          toName,
          subject,
          resumePath,
        });

    try {
      appendSentRow(config.paths.sentXlsx, {
        email: toEmail,
        name: toName,
        subject,
        error: "",
      });
    } catch (e) {
      console.error("[excel-log] Failed to log sent email:", e?.message || e);
    }

    res.json({
      ok: true,
      toEmail,
      subject,
      messageId: info.messageId,
      response: info.response,
      usedDefaults: {
        subject: !subjectOverride,
        body: !bodyOverride,
        resume: !req.file,
      },
    });
  } catch (e) {
    try {
      appendSentRow(config.paths.sentXlsx, {
        email: toEmail,
        name: toName,
        subject,
        error: String(e?.message || e),
      });
    } catch (logErr) {
      console.error("[excel-log] Failed to log failed email:", logErr?.message || logErr);
    }
    res.status(500).json({ ok: false, error: String(e?.message || e) });
  } finally {
    // Clean up uploaded file if present.
    if (req.file?.path) {
      fs.promises.unlink(req.file.path).catch(() => {});
    }
  }
});

// Bulk send from Excel:
// - excel is required
// - resume is optional (applies to all rows)
// - for each row, subject/body/name can override; otherwise defaults apply
app.post(
  "/api/send-bulk",
  upload.fields([
    { name: "excel", maxCount: 1 },
    { name: "resume", maxCount: 1 },
  ]),
  async (req, res) => {
    const excelFile = req.files?.excel?.[0];
    const resumeFile = req.files?.resume?.[0];
    if (!excelFile?.path) {
      return res.status(400).json({ ok: false, error: "Excel (.xlsx) file is required." });
    }

    console.log(
      `[ui] bulk send requested: excel=${excelFile.originalname} (${excelFile.size} bytes) resume=${
        resumeFile?.originalname || "(default)"
      }`,
    );

    let rows = [];
    try {
      rows = parseRecipientsFromXlsx(excelFile.path);
    } catch (e) {
      return res.status(400).json({
        ok: false,
        error: `Failed to read Excel. Make sure it's a valid .xlsx with columns: email, recipient name, subject, body. (${String(
          e?.message || e,
        )})`,
      });
    } finally {
      fs.promises.unlink(excelFile.path).catch(() => {});
    }

    if (!rows.length) {
      if (resumeFile?.path) fs.promises.unlink(resumeFile.path).catch(() => {});
      return res.status(400).json({
        ok: false,
        error:
          "No valid rows found. Ensure your sheet has an 'email' (or 'mail') column with valid emails.",
      });
    }

    console.log(`[ui] bulk parsed rows: ${rows.length}`);

    const eff = getEffectiveSettings();
    const resumePath = resumeFile?.path ? resumeFile.path : eff.resumePath;
    const transporter = await createTransporter({ smtp: eff.smtp, from: eff.from });

    const results = [];
    for (const r of rows) {
      const subject = r.subject || eff.subject || config.content.subject;
      const bodyOverride = r.body || eff.defaultBody || "";

      let text;
      let html;
      if (bodyOverride) {
        const overridden = buildOverriddenEmail({
          recipientName: r.name,
          recipientEmail: r.email,
          bodyText: bodyOverride,
        });
        text = overridden.text;
        html = overridden.html;
      } else {
        const built = buildEmail({
          recipientName: r.name,
          recipientEmail: r.email,
          subject,
        });
        text = built.text;
        html = built.html;
      }

      try {
        console.log(`[ui] bulk sending -> ${r.email}`);
        const info = await transporter.sendMail({
          from: eff.from.name ? `"${eff.from.name}" <${eff.from.email}>` : eff.from.email,
          to: r.email,
          subject,
          text,
          html,
          attachments: [
            {
              filename: resumeFile?.originalname || path.basename(resumePath),
              path: resumePath,
            },
          ],
        });
        console.log(`[ui] bulk sent OK -> ${r.email} (messageId=${info.messageId || "n/a"})`);
        try {
          appendSentRow(config.paths.sentXlsx, {
            email: r.email,
            name: String(r.name || ""),
            subject,
            error: "",
          });
        } catch (logErr) {
          console.error("[excel-log] Failed to log bulk sent email:", logErr?.message || logErr);
        }
        results.push({ email: r.email, ok: true, messageId: info.messageId, response: info.response });
      } catch (e) {
        console.error(`[ui] bulk send FAILED -> ${r.email}: ${String(e?.message || e)}`);
        try {
          appendSentRow(config.paths.sentXlsx, {
            email: r.email,
            name: String(r.name || ""),
            subject,
            error: String(e?.message || e),
          });
        } catch (logErr) {
          console.error("[excel-log] Failed to log bulk failed email:", logErr?.message || logErr);
        }
        results.push({ email: r.email, ok: false, error: String(e?.message || e) });
      }

      await sleep(config.behavior.delayMsBetweenEmails);
    }

    if (resumeFile?.path) fs.promises.unlink(resumeFile.path).catch(() => {});

    const sent = results.filter((x) => x.ok).length;
    const failed = results.length - sent;
    res.json({ ok: true, total: results.length, sent, failed, results });
  },
);

// Bulk send from direct copy/paste list:
// - emails is required (comma/newline separated)
// - resume optional (applies to all)
app.post("/api/send-list", upload.single("resume"), async (req, res) => {
  const emailsRaw = String(req.body.emails || "").trim();
  const emails = parseEmailsFromText(emailsRaw);
  if (!emails.length) {
    if (req.file?.path) fs.promises.unlink(req.file.path).catch(() => {});
    return res.status(400).json({
      ok: false,
      error: "No valid emails found. Paste comma/newline-separated emails.",
    });
  }

  console.log(`[ui] list send requested: emails=${emails.length} resume=${req.file?.originalname || "(default)"}`);

  const eff = getEffectiveSettings();
  const resumePath = req.file?.path ? req.file.path : eff.resumePath;
  const transporter = await createTransporter({ smtp: eff.smtp, from: eff.from });

  const results = [];
  for (const email of emails) {
    const subject = eff.subject || config.content.subject;
    const bodyOverride = eff.defaultBody || "";
    let text;
    let html;
    if (bodyOverride) {
      const overridden = buildOverriddenEmail({
        recipientName: "",
        recipientEmail: email,
        bodyText: bodyOverride,
      });
      text = overridden.text;
      html = overridden.html;
    } else {
      const built = buildEmail({
        recipientName: "",
        recipientEmail: email,
        subject,
      });
      text = built.text;
      html = built.html;
    }
    try {
      console.log(`[ui] list sending -> ${email}`);
      const info = await transporter.sendMail({
        from: eff.from.name ? `"${eff.from.name}" <${eff.from.email}>` : eff.from.email,
        to: email,
        subject,
        text,
        html,
        attachments: [
          {
            filename: req.file?.originalname || path.basename(resumePath),
            path: resumePath,
          },
        ],
      });
      try {
        appendSentRow(config.paths.sentXlsx, {
          email,
          name: "",
          subject,
          error: "",
        });
      } catch (logErr) {
        console.error("[excel-log] Failed to log list sent email:", logErr?.message || logErr);
      }
      results.push({ email, ok: true, messageId: info.messageId, response: info.response });
    } catch (e) {
      try {
        appendSentRow(config.paths.sentXlsx, {
          email,
          name: "",
          subject,
          error: String(e?.message || e),
        });
      } catch (logErr) {
        console.error("[excel-log] Failed to log list failed email:", logErr?.message || logErr);
      }
      console.error(`[ui] list send FAILED -> ${email}: ${String(e?.message || e)}`);
      results.push({ email, ok: false, error: String(e?.message || e) });
    }

    await sleep(config.behavior.delayMsBetweenEmails);
  }

  if (req.file?.path) fs.promises.unlink(req.file.path).catch(() => {});

  const sent = results.filter((x) => x.ok).length;
  const failed = results.length - sent;
  return res.json({ ok: true, total: results.length, sent, failed, results });
});

// Render (and most PaaS) requires binding to 0.0.0.0 and the port provided in $PORT.
const HOST = String(process.env.HOST || process.env.UI_HOST || "0.0.0.0");
const PORT = Number(process.env.PORT || process.env.UI_PORT || 4545);
const server = app.listen(PORT, HOST, () => {
  const shownHost = HOST === "0.0.0.0" ? "localhost" : HOST;
  console.log(`UI running at http://${shownHost}:${PORT}`);
  console.log(`Env loaded from: ${config.meta?.loadedEnvFile || "(unknown)"}`);
  console.log(
    `Auth enabled: ${AUTH_ENABLED ? "yes" : "no"}${
      AUTH_ENABLED ? "" : " (set UI_AUTH_USER/UI_AUTH_PASS to enable)"
    }`,
  );
});

server.on("error", (err) => {
  console.error("UI server failed to start:", err?.message || err);
  process.exitCode = 1;
});


