const path = require("path");

const { readJson } = require("./utils");

function guessGreetingFromEmail(email) {
  const e = String(email || "").toLowerCase();
  const hrHints = ["hr", "hiring", "recruit", "talent", "peopleops", "people-ops"];
  if (hrHints.some((h) => e.includes(h))) return "Hiring Team";
  return "Hiring Team";
}

function loadDefaultBodyFromUiSettings() {
  try {
    // backend/src -> backend -> project root
    const root = path.resolve(__dirname, "..", "..");
    const settingsPath = path.resolve(root, "data", "ui-settings.json");
    const s = readJson(settingsPath, {});
    return String(s?.defaultBody || "").trim();
  } catch {
    return "";
  }
}

function bodyAlreadyHasSignature(bodyText) {
  const b = String(bodyText || "");
  if (!b.trim()) return false;
  const v = b.toLowerCase();
  // tolerate newlines/spaces between words
  return /warm\s+regards/.test(v) || /regards\s*,/.test(v) || /shubham\s+pawar/.test(v);
}

function firstNameOnly(fullName) {
  const name = String(fullName || "").trim();
  if (!name) return "";
  // Keep generic greetings as-is.
  if (name.toLowerCase() === "hiring team") return "Hiring Team";
  // Strip common punctuation and take first token.
  const cleaned = name.replace(/[(),]/g, " ").trim();
  const first = cleaned.split(/\s+/)[0] || "";
  return first || name;
}

function buildEmail({ recipientName, recipientEmail, subject }) {
  const name = String(recipientName || "").trim();
  const greetingName = firstNameOnly(name) || guessGreetingFromEmail(recipientEmail);

  const defaultBody = loadDefaultBodyFromUiSettings();
  const signatureText = [
    "Warm regards,",
    "Shubham Pawar",
    "MERN Stack Developer | Software Engineer",
    "Immediate Joiner",
  ].join("\n");

  const shouldAddSignature = !bodyAlreadyHasSignature(defaultBody);
  const textParts = [`Hi ${greetingName},`, "", defaultBody];
  if (shouldAddSignature) textParts.push("", signatureText, "");
  else textParts.push("");
  const text = textParts.join("\n").trim() + "\n";

  const signatureHtml = `
    <p>
      Warm regards,<br />
      Shubham Pawar<br />
      MERN Stack Developer | Software Engineer<br />
      Immediate Joiner
    </p>
  `.trim();

  const html = `
    <p>Hi ${escapeHtml(greetingName)},</p>
    <p>${bodyToHtml(defaultBody)}</p>
    ${shouldAddSignature ? signatureHtml : ""}
  `.trim();

  return { subject, text, html };
}

function bodyToHtml(bodyText) {
  return String(bodyText || "")
    .split("\n")
    .map((line) => escapeHtml(line))
    .join("<br />");
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

module.exports = { buildEmail };


