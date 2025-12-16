const fs = require("fs");
const path = require("path");
const dotenv = require("dotenv");

// backend/src -> backend -> project root
const ROOT = path.resolve(__dirname, "..", "..");

// Prefer local secrets in .env, but allow running from env.example if .env is missing.
const ENV_PATH_PRIMARY = path.resolve(ROOT, ".env");
const ENV_PATH_FALLBACK = path.resolve(ROOT, "env.example");

let loadedEnvFile = null;
if (fs.existsSync(ENV_PATH_PRIMARY)) {
  dotenv.config({ path: ENV_PATH_PRIMARY });
  loadedEnvFile = ENV_PATH_PRIMARY;
} else if (fs.existsSync(ENV_PATH_FALLBACK)) {
  dotenv.config({ path: ENV_PATH_FALLBACK });
  loadedEnvFile = ENV_PATH_FALLBACK;
} else {
  // As a last resort, attempt default dotenv behavior (.env in cwd)
  dotenv.config();
  loadedEnvFile = "(dotenv default)";
}

function env(name, fallback) {
  const v = process.env[name];
  return v === undefined || v === "" ? fallback : v;
}

function envBool(name, fallback) {
  const v = env(name, "");
  if (v === "") return fallback;
  return ["1", "true", "yes", "y", "on"].includes(String(v).toLowerCase());
}

function envInt(name, fallback) {
  const v = env(name, "");
  if (v === "") return fallback;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n : fallback;
}

module.exports = {
  meta: {
    loadedEnvFile,
  },
  paths: {
    root: ROOT,
    recipientsCsv: path.resolve(ROOT, "data/recipients.csv"),
    sentJson: path.resolve(ROOT, "data/sent.json"),
    sentXlsx: path.resolve(ROOT, env("SENT_XLSX_PATH", "data/sent.xlsx")),
    resumePath: path.resolve(ROOT, env("RESUME_PATH", "assets/Shubham_Pawar_3Yr.pdf")),
  },
  smtp: {
    host: env("SMTP_HOST", ""),
    port: envInt("SMTP_PORT", 465),
    secure: envBool("SMTP_SECURE", true),
    user: env("SMTP_USER", ""),
    pass: env("SMTP_PASS", ""),
  },
  from: {
    email: env("FROM_EMAIL", env("SMTP_USER", "")),
    name: env("FROM_NAME", "Shubham Pawar"),
  },
  schedule: {
    cron: env("SCHEDULE_CRON", "0 10 * * *"),
    timezone: env("TIMEZONE", "Asia/Kolkata"),
  },
  behavior: {
    delayMsBetweenEmails: envInt("DELAY_MS_BETWEEN_EMAILS", 1500),
    dryRun: envBool("DRY_RUN", false),
  },
  content: {
    subject: env(
      "SUBJECT",
      "Application for MERN Stack Developer Role — Immediate Joiner | 3 Yrs Experience",
    ),
  },
};


