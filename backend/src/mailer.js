const fs = require("fs");
const path = require("path");
const nodemailer = require("nodemailer");
const { buildEmail } = require("./template");

function assertSmtpConfig(smtp) {
  const missing = [];
  if (!smtp.host) missing.push("SMTP_HOST");
  if (!smtp.port) missing.push("SMTP_PORT");
  if (!smtp.user) missing.push("SMTP_USER");
  if (!smtp.pass) missing.push("SMTP_PASS");
  if (missing.length) {
    throw new Error(`Missing SMTP config: ${missing.join(", ")} (set these in .env)`);
  }
}

function isReachabilityError(err) {
  const code = String(err?.code || "").toUpperCase();
  return [
    "ENETUNREACH",
    "EHOSTUNREACH",
    "ENETDOWN",
    "ECONNREFUSED",
    "ETIMEDOUT",
    "ENOTFOUND",
    "EAI_AGAIN",
  ].includes(code);
}

function formatEndpoint(smtp) {
  return `${smtp.host}:${smtp.port} (secure=${smtp.secure ? "true" : "false"})`;
}

function createNodeMailerTransport({ smtp, from }) {
  return nodemailer.createTransport({
    host: smtp.host,
    port: smtp.port,
    secure: smtp.secure,
    auth: {
      user: smtp.user,
      pass: smtp.pass,
    },
  });
}

async function createTransporter({ smtp, from }) {
  assertSmtpConfig(smtp);
  if (!from?.email) throw new Error("Missing FROM_EMAIL (or SMTP_USER) in .env");

  const primary = createNodeMailerTransport({ smtp, from });
  try {
    await primary.verify();
    return primary;
  } catch (err) {
    // Common case: some networks/ISPs block outbound SMTP submission ports.
    // If user configured 587/STARTTLS, retry once on 465/SSL which often works.
    const canTry465 =
      Number(smtp.port) === 587 && smtp.secure === false && isReachabilityError(err);

    if (canTry465) {
      const fallbackSmtp = { ...smtp, port: 465, secure: true };
      const fallback = createNodeMailerTransport({ smtp: fallbackSmtp, from });
      try {
        await fallback.verify();
        console.warn(
          `[smtp] Primary ${formatEndpoint(smtp)} failed (${err?.code || "UNKNOWN"}). Using fallback ${formatEndpoint(
            fallbackSmtp,
          )}.`,
        );
        return fallback;
      } catch (err2) {
        throw new Error(
          `SMTP connection failed.\n- Primary: ${formatEndpoint(smtp)} -> ${err?.code || "UNKNOWN"}: ${
            err?.message || err
          }\n- Fallback: ${formatEndpoint(fallbackSmtp)} -> ${err2?.code || "UNKNOWN"}: ${
            err2?.message || err2
          }\n\nThis is a network reachability problem (not an auth/password issue). If you're on VPN/corporate Wi‑Fi, or your ISP blocks SMTP ports, switch networks or use an email provider API (SendGrid/Mailgun/Resend) instead of direct SMTP.`,
          { cause: err2 },
        );
      }
    }

    if (String(err?.code || "").toUpperCase() === "ENETUNREACH") {
      throw new Error(
        `SMTP connection failed: ${formatEndpoint(smtp)} -> ENETUNREACH: ${
          err?.message || err
        }\n\nYour network cannot reach the SMTP server on that port. Try:\n- Set SMTP_PORT=465 and SMTP_SECURE=true (SSL)\n- Disable VPN / try a different Wi‑Fi/network\n- If on a hosted environment, use an email API provider (HTTP) instead of SMTP`,
        { cause: err },
      );
    }

    throw new Error(
      `SMTP connection failed: ${formatEndpoint(smtp)} -> ${err?.code || "UNKNOWN"}: ${
        err?.message || err
      }`,
      { cause: err },
    );
  }
}

async function sendApplicationEmail({
  transporter,
  from,
  toEmail,
  toName,
  subject,
  resumePath,
}) {
  const { text, html } = buildEmail({
    recipientName: toName,
    recipientEmail: toEmail,
    subject,
  });

  const attachments = [];
  if (resumePath) {
    const abs = path.resolve(resumePath);
    if (!fs.existsSync(abs)) {
      throw new Error(
        `Resume not found at ${abs}. Put your PDF there or set RESUME_PATH in .env`,
      );
    }
    attachments.push({
      filename: path.basename(abs),
      path: abs,
    });
  }

  return await transporter.sendMail({
    from: from.name ? `"${from.name}" <${from.email}>` : from.email,
    to: toEmail,
    subject,
    text,
    html,
    attachments,
  });
}

module.exports = { createTransporter, sendApplicationEmail };


