const config = require("./config");
const { loadRecipients } = require("./recipients");
const { loadSentLog, isSent, markSent, markError } = require("./sent-log");
const { appendSentRow } = require("./excel-log");
const { sleep } = require("./utils");
const { createTransporter, sendApplicationEmail } = require("./mailer");

async function sendPending({ source = "manual" } = {}) {
  const recipients = loadRecipients(config.paths.recipientsCsv);
  if (!recipients.length) {
    console.log(
      "No recipients found in recipients.csv. Add lines like 'email,name' (header optional).",
    );
    return { sent: 0, pending: 0, recipients: 0 };
  }
  const sentLog = loadSentLog(config.paths.sentJson);

  const pending = recipients.filter((r) => !isSent(sentLog, r.email));
  if (!pending.length) {
    console.log("No pending recipients. (All emails already sent)");
    return { sent: 0, pending: 0 };
  }

  console.log(`Pending recipients: ${pending.length}`);
  if (config.behavior.dryRun) {
    console.log("DRY_RUN=true — not sending emails. Would send to:");
    for (const r of pending) console.log(`- ${r.email}${r.name ? ` (${r.name})` : ""}`);
    return { sent: 0, pending: pending.length, dryRun: true };
  }

  const transporter = await createTransporter({ smtp: config.smtp, from: config.from });
  let sentCount = 0;

  for (const r of pending) {
    console.log(`Sending -> ${r.email}${r.name ? ` (${r.name})` : ""}`);
    try {
      const info = await sendApplicationEmail({
        transporter,
        from: config.from,
        toEmail: r.email,
        toName: r.name,
        subject: config.content.subject,
        resumePath: config.paths.resumePath,
      });
      markSent(config.paths.sentJson, r.email, {
        messageId: info.messageId,
        response: info.response,
        source,
      });
      try {
        appendSentRow(config.paths.sentXlsx, {
          email: r.email,
          name: r.name || "",
          subject: config.content.subject,
          error: "",
        });
      } catch (e) {
        console.error("[excel-log] Failed to log sent email:", e?.message || e);
      }
      sentCount += 1;
      console.log(`Sent OK: ${r.email} (messageId: ${info.messageId || "n/a"})`);
    } catch (err) {
      markError(config.paths.sentJson, r.email, {
        error: String(err?.message || err),
        source,
      });
      try {
        appendSentRow(config.paths.sentXlsx, {
          email: r.email,
          name: r.name || "",
          subject: config.content.subject,
          error: String(err?.message || err),
        });
      } catch (e) {
        console.error("[excel-log] Failed to log failed email:", e?.message || e);
      }
      console.error(`Send FAILED: ${r.email} -> ${err?.message || err}`);
    }

    await sleep(config.behavior.delayMsBetweenEmails);
  }

  console.log("Done.");
  return { sent: sentCount, pending: pending.length };
}

module.exports = { sendPending };


