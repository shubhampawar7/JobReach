const chokidar = require("chokidar");
const config = require("./config");
const { loadRecipients } = require("./recipients");
const { loadSentLog, isSent, markSent, markError } = require("./sent-log");
const { sleep } = require("./utils");
const { createTransporter, sendApplicationEmail } = require("./mailer");

function setFromRecipients(recipients) {
  return new Set(recipients.map((r) => r.email));
}

async function startWatcher() {
  console.log("Watching for new recipients in:", config.paths.recipientsCsv);
  console.log("DRY_RUN:", config.behavior.dryRun);

  let recipients = loadRecipients(config.paths.recipientsCsv);
  let known = setFromRecipients(recipients);

  const transporter = await createTransporter({ smtp: config.smtp, from: config.from });

  let debounceTimer = null;
  const onChange = () => {
    if (debounceTimer) clearTimeout(debounceTimer);
    debounceTimer = setTimeout(async () => {
      try {
        const next = loadRecipients(config.paths.recipientsCsv);
        const nextSet = setFromRecipients(next);

        const addedEmails = [];
        for (const e of nextSet) if (!known.has(e)) addedEmails.push(e);

        known = nextSet;
        recipients = next;

        if (!addedEmails.length) return;

        const sentLog = loadSentLog(config.paths.sentJson);
        const addedRecipients = recipients.filter((r) => addedEmails.includes(r.email));
        const toSend = addedRecipients.filter((r) => !isSent(sentLog, r.email));
        if (!toSend.length) return;

        console.log(`New recipients added: ${toSend.length}`);
        for (const r of toSend) {
          if (config.behavior.dryRun) {
            console.log(`DRY_RUN=true — would send -> ${r.email}`);
            continue;
          }
          try {
            console.log(`Sending (watch) -> ${r.email}`);
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
              source: "watch",
            });
            console.log(`Sent (watch) OK: ${r.email}`);
          } catch (err) {
            markError(config.paths.sentJson, r.email, {
              error: String(err?.message || err),
              source: "watch",
            });
            console.error(`Send FAILED (watch): ${r.email} -> ${err?.message || err}`);
          }
          await sleep(config.behavior.delayMsBetweenEmails);
        }
      } catch (err) {
        console.error("Watcher error:", err?.message || err);
      }
    }, 400);
  };

  const watcher = chokidar.watch(config.paths.recipientsCsv, { ignoreInitial: true });
  watcher.on("add", onChange).on("change", onChange);
  return watcher;
}

module.exports = { startWatcher };


