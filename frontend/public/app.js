const $ = (sel) => document.querySelector(sel);

const form = $("#sendForm");
const statusEl = $("#status");
const sendBtn = $("#sendBtn");
const spinner = $(".btnSpinner");
const resetBtn = $("#resetBtn");
const subjectInput = $("#subject");
const bodyInput = $("#body");

const dropzone = $("#dropzone");
const resumeInput = $("#resume");
const filePill = $("#filePill");
const fileName = $("#fileName");
const clearFile = $("#clearFile");

const excelInput = $("#excel");
const bulkResumeInput = $("#bulkResume");
const bulkSendBtn = $("#bulkSendBtn");
const logoutBtn = $("#logoutBtn");
const bulkModeSel = $("#bulkMode");
const bulkExcelWrap = $("#bulkExcelWrap");
const bulkPasteWrap = $("#bulkPasteWrap");
const bulkEmailsTa = $("#bulkEmails");

// Tabs
const tabSend = $("#tabSend");
const tabHr = $("#tabHr");
const tabDefaults = $("#tabDefaults");
const panelSend = $("#panelSend");
const panelHr = $("#panelHr");
const panelDefaults = $("#panelDefaults");
const panelSide = $("#panelSide");

// HR finder
const hrSearchBtn = $("#hrSearchBtn");
const hrResults = $("#hrResults");
const providerSel = $("#provider");
let lastHrContacts = [];
let lastHrPhone = "";

const companyInput = $("#company");
const companyDropdown = $("#companyDropdown");
const domainInput = $("#domain");

function debounce(fn, waitMs) {
  let t = null;
  return (...args) => {
    if (t) clearTimeout(t);
    t = setTimeout(() => fn(...args), waitMs);
  };
}

function uniqStrings(arr) {
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
  return out;
}

let savedCompanyNames = [];
let liveCompanyNames = [];
let taActiveIndex = -1;

function getAllCompanyNames() {
  return uniqStrings([...(savedCompanyNames || []), ...(liveCompanyNames || [])]);
}

function ensureDropdownShell() {
  if (!companyDropdown) return null;
  // Use an inner wrapper so we can keep borders fixed while list scrolls.
  if (!companyDropdown.querySelector(".typeaheadMenuInner")) {
    companyDropdown.innerHTML = `<div class="typeaheadMenuInner"></div><div class="typeaheadMeta">Type to search, or pick a company.</div>`;
  }
  return companyDropdown.querySelector(".typeaheadMenuInner");
}

function closeCompanyDropdown() {
  if (!companyDropdown) return;
  companyDropdown.classList.add("hidden");
  taActiveIndex = -1;
}

function openCompanyDropdown() {
  if (!companyDropdown) return;
  companyDropdown.classList.remove("hidden");
}

function setActiveItem(idx) {
  const inner = ensureDropdownShell();
  if (!inner) return;
  const items = Array.from(inner.querySelectorAll(".typeaheadItem"));
  if (!items.length) {
    taActiveIndex = -1;
    return;
  }
  taActiveIndex = Math.max(0, Math.min(idx, items.length - 1));
  items.forEach((el, i) => el.classList.toggle("active", i === taActiveIndex));
  const active = items[taActiveIndex];
  if (active) active.scrollIntoView({ block: "nearest" });
}

function commitCompanyValue(v) {
  if (!companyInput) return;
  companyInput.value = String(v || "");
  closeCompanyDropdown();
}

function renderCompanyDropdown({ forceOpen = false } = {}) {
  if (!companyDropdown) return;
  const inner = ensureDropdownShell();
  if (!inner) return;

  const q = String(companyInput?.value || "").trim().toLowerCase();
  const all = getAllCompanyNames();
  const filtered = q
    ? all.filter((n) => String(n).toLowerCase().includes(q))
    : all;

  const shown = filtered.slice(0, 80);

  if (!shown.length) {
    inner.innerHTML = `<div class="typeaheadMeta" style="border-top:none;background:transparent;padding:10px 10px">No matches. Keep typing to use a custom name.</div>`;
    taActiveIndex = -1;
  } else {
    inner.innerHTML = shown
      .map(
        (name) =>
          `<button type="button" class="typeaheadItem" role="option" data-value="${escapeHtml(
            name,
          )}">${escapeHtml(name)}</button>`,
      )
      .join("");
    taActiveIndex = -1;
  }

  // Open only when focusing/typing.
  if (forceOpen) openCompanyDropdown();
}

async function loadSavedCompanyNames() {
  try {
    const res = await fetch("/api/company-names");
    const data = await res.json().catch(() => ({}));
    if (!res.ok || !data.ok) return;
    savedCompanyNames = uniqStrings(data.companies || []);
    renderCompanyDropdown();
  } catch {
    // ignore
  }
}

const fetchCompanySuggestDebounced = debounce(async () => {
  const q = String(companyInput?.value || "").trim();
  if (q.length < 2) {
    liveCompanyNames = [];
    renderCompanyDropdown({ forceOpen: true });
    return;
  }
  try {
    const res = await fetch(`/api/company-suggest?query=${encodeURIComponent(q)}`);
    const data = await res.json().catch(() => ({}));
    if (!res.ok || !data.ok) return;
    liveCompanyNames = uniqStrings(data.companies || []);
    renderCompanyDropdown({ forceOpen: true });
  } catch {
    // ignore
  }
}, 180);

companyInput?.addEventListener("focus", () => {
  renderCompanyDropdown({ forceOpen: true });
});

companyInput?.addEventListener("input", () => {
  renderCompanyDropdown({ forceOpen: true });
  fetchCompanySuggestDebounced();
});

companyInput?.addEventListener("keydown", (e) => {
  if (!companyDropdown || companyDropdown.classList.contains("hidden")) {
    if (e.key === "ArrowDown") {
      renderCompanyDropdown({ forceOpen: true });
      setActiveItem(0);
      e.preventDefault();
    }
    return;
  }

  const inner = ensureDropdownShell();
  const items = inner ? Array.from(inner.querySelectorAll(".typeaheadItem")) : [];
  if (!items.length) return;

  if (e.key === "ArrowDown") {
    setActiveItem((taActiveIndex < 0 ? -1 : taActiveIndex) + 1);
    e.preventDefault();
  } else if (e.key === "ArrowUp") {
    setActiveItem((taActiveIndex < 0 ? items.length : taActiveIndex) - 1);
    e.preventDefault();
  } else if (e.key === "Enter") {
    if (taActiveIndex >= 0 && items[taActiveIndex]) {
      commitCompanyValue(items[taActiveIndex].getAttribute("data-value") || "");
      e.preventDefault();
    }
  } else if (e.key === "Escape") {
    closeCompanyDropdown();
    e.preventDefault();
  }
});

companyDropdown?.addEventListener("mousedown", (e) => {
  const btn = e.target?.closest?.(".typeaheadItem");
  if (!btn) return;
  e.preventDefault(); // prevent input blur before click
  commitCompanyValue(btn.getAttribute("data-value") || "");
});

document.addEventListener("mousedown", (e) => {
  const within =
    e.target?.closest?.("#companyTypeahead") || e.target?.closest?.("#companyDropdown") || null;
  if (!within) closeCompanyDropdown();
});

async function initProviderStatus() {
  try {
    const res = await fetch("/api/provider-status");
    const data = await res.json().catch(() => ({}));
    if (!providerSel) return;
    const apolloOpt = providerSel.querySelector('option[value="apollo"]');
    if (!apolloOpt) return;
    const apollo = data?.providers?.apollo || {};

    if (!apollo.configured) {
      apolloOpt.disabled = true;
      apolloOpt.textContent = "Apollo (set APOLLO_API_KEY)";
      return;
    }
    if (apollo.looksLikeGraphOS) {
      apolloOpt.disabled = true;
      apolloOpt.textContent = "Apollo (GraphOS key detected — needs Apollo.io key)";
      return;
    }
  } catch {
    // ignore
  }
}

function toast(type, title, msg, { timeoutMs = 3500 } = {}) {
  let wrap = document.querySelector(".toastWrap");
  if (!wrap) {
    wrap = document.createElement("div");
    wrap.className = "toastWrap";
    document.body.appendChild(wrap);
  }

  const el = document.createElement("div");
  el.className = `toast ${type === "bad" ? "bad" : "good"}`;
  el.innerHTML = `<div class="toastTitle">${escapeHtml(title)}</div><div class="toastMsg">${escapeHtml(
    msg || "",
  )}</div>`;
  wrap.appendChild(el);

  const remove = () => {
    el.classList.add("toastOut");
    setTimeout(() => el.remove(), 170);
  };

  setTimeout(remove, timeoutMs);
  el.addEventListener("click", remove);
}

function setStatus(type, html) {
  statusEl.classList.remove("empty", "good", "bad");
  statusEl.classList.add(type);
  statusEl.innerHTML = html;
}

function setLoading(isLoading) {
  sendBtn.disabled = isLoading;
  if (isLoading) spinner.classList.remove("hidden");
  else spinner.classList.add("hidden");
}

function setTab(which) {
  const isSend = which === "send";
  const isHr = which === "hr";
  const isDefaults = which === "defaults";
  tabSend?.classList.toggle("active", isSend);
  tabHr?.classList.toggle("active", isHr);
  tabDefaults?.classList.toggle("active", isDefaults);
  panelSend?.classList.toggle("hidden", !isSend);
  panelHr?.classList.toggle("hidden", !isHr);
  panelDefaults?.classList.toggle("hidden", !isDefaults);
  // Keep the right-side status visible for send; hide for other tabs to give space.
  panelSide?.classList.toggle("hidden", !isSend);
}

tabSend?.addEventListener("click", () => setTab("send"));
tabHr?.addEventListener("click", () => setTab("hr"));
tabDefaults?.addEventListener("click", () => setTab("defaults"));

initProviderStatus();
loadSavedCompanyNames();

// -------------------------
// Defaults tab (settings)
// -------------------------
const defSmtpHost = $("#defSmtpHost");
const defSmtpPort = $("#defSmtpPort");
const defSmtpSecure = $("#defSmtpSecure");
const defSmtpUser = $("#defSmtpUser");
const defSmtpPass = $("#defSmtpPass");
const defFromEmail = $("#defFromEmail");
const defFromName = $("#defFromName");
const defSubject = $("#defSubject");
const defBody = $("#defBody");
const defResume = $("#defResume");
const defUploadResumeBtn = $("#defUploadResumeBtn");
const defSaveBtn = $("#defSaveBtn");
const defStatus = $("#defStatus");

function setDefStatus(type, html) {
  if (!defStatus) return;
  defStatus.classList.remove("empty", "good", "bad");
  defStatus.classList.add(type);
  defStatus.innerHTML = html;
}

async function loadDefaultsIntoUI() {
  if (!defSmtpHost) return;
  try {
    const res = await fetch("/api/settings");
    const data = await res.json().catch(() => ({}));
    if (!res.ok || !data.ok) throw new Error(data.error || `Request failed (${res.status})`);
    const s = data.settings || {};
    defSmtpHost.value = s.smtpHost || "";
    defSmtpPort.value = s.smtpPort ? String(s.smtpPort) : "";
    defSmtpSecure.value = String(Boolean(s.smtpSecure));
    defSmtpUser.value = s.smtpUser || "";
    defFromEmail.value = s.fromEmail || "";
    defFromName.value = s.fromName || "";
    defSubject.value = s.subject || "";
    defBody.value = s.defaultBody || "";
    defSmtpPass.value = "";

    // Also apply defaults into Send tab (only if user hasn't typed overrides).
    if (subjectInput && !String(subjectInput.value || "").trim() && s.subject) {
      subjectInput.value = String(s.subject || "");
    }
    if (bodyInput && !String(bodyInput.value || "").trim() && s.defaultBody) {
      bodyInput.value = String(s.defaultBody || "");
    }

    setDefStatus(
      "empty",
      `Loaded. Saved password: <strong>${s.smtpPassSet ? "yes" : "no"}</strong>. Resume uploaded: <strong>${
        s.resumeSet ? "yes" : "no"
      }</strong>.`,
    );
  } catch (e) {
    setDefStatus("bad", `<strong>Failed to load.</strong><br/>${escapeHtml(String(e?.message || e))}`);
  }
}

defSaveBtn?.addEventListener("click", async () => {
  try {
    setDefStatus("empty", "Saving…");
    const payload = {
      smtpHost: String(defSmtpHost?.value || "").trim(),
      smtpPort: Number(String(defSmtpPort?.value || "").trim() || 0) || null,
      smtpSecure: String(defSmtpSecure?.value || "false") === "true",
      smtpUser: String(defSmtpUser?.value || "").trim(),
      smtpPass: String(defSmtpPass?.value || ""),
      fromEmail: String(defFromEmail?.value || "").trim(),
      fromName: String(defFromName?.value || "").trim(),
      subject: String(defSubject?.value || "").trim(),
      defaultBody: String(defBody?.value || "").trim(),
    };
    const res = await fetch("/api/settings", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok || !data.ok) {
      const err = data.error || `Request failed (${res.status})`;
      setDefStatus("bad", `<strong>Save failed.</strong><br/>${escapeHtml(err)}`);
      toast("bad", "Save failed", err);
      return;
    }
    defSmtpPass.value = "";
    setDefStatus("good", "<strong>Saved.</strong> Defaults updated.");
    toast("good", "Saved", "Defaults updated");
  } catch (e) {
    const msg = String(e?.message || e);
    setDefStatus("bad", `<strong>Error.</strong><br/>${escapeHtml(msg)}`);
    toast("bad", "Error", msg);
  }
});

defUploadResumeBtn?.addEventListener("click", async () => {
  const f = defResume?.files?.[0];
  if (!f) {
    toast("bad", "Missing file", "Choose a PDF resume first.");
    return;
  }
  if (!String(f.name || "").toLowerCase().endsWith(".pdf")) {
    toast("bad", "Invalid file", "Resume must be a PDF.");
    return;
  }
  try {
    setDefStatus("empty", "Uploading resume…");
    const fd = new FormData();
    fd.set("resume", f);
    const res = await fetch("/api/settings/resume", { method: "POST", body: fd });
    const data = await res.json().catch(() => ({}));
    if (!res.ok || !data.ok) {
      const err = data.error || `Request failed (${res.status})`;
      setDefStatus("bad", `<strong>Upload failed.</strong><br/>${escapeHtml(err)}`);
      toast("bad", "Upload failed", err);
      return;
    }
    setDefStatus("good", "<strong>Uploaded.</strong> Default resume updated.");
    toast("good", "Uploaded", "Default resume updated");
    await loadDefaultsIntoUI();
  } catch (e) {
    const msg = String(e?.message || e);
    setDefStatus("bad", `<strong>Error.</strong><br/>${escapeHtml(msg)}`);
    toast("bad", "Error", msg);
  }
});

loadDefaultsIntoUI();

function updateFileUI() {
  const f = resumeInput.files && resumeInput.files[0];
  if (!f) {
    filePill.classList.add("hidden");
    fileName.textContent = "";
    return;
  }
  filePill.classList.remove("hidden");
  fileName.textContent = `${f.name} (${Math.round(f.size / 1024)} KB)`;
}

resumeInput.addEventListener("change", updateFileUI);
clearFile.addEventListener("click", () => {
  resumeInput.value = "";
  updateFileUI();
});

function prevent(e) {
  e.preventDefault();
  e.stopPropagation();
}

["dragenter", "dragover"].forEach((evt) => {
  dropzone.addEventListener(evt, (e) => {
    prevent(e);
    dropzone.classList.add("drag");
  });
});

["dragleave", "drop"].forEach((evt) => {
  dropzone.addEventListener(evt, (e) => {
    prevent(e);
    dropzone.classList.remove("drag");
  });
});

dropzone.addEventListener("drop", (e) => {
  const f = e.dataTransfer.files && e.dataTransfer.files[0];
  if (!f) return;
  if (!f.name.toLowerCase().endsWith(".pdf")) {
    setStatus("bad", "<strong>Resume must be a PDF.</strong>");
    return;
  }
  const dt = new DataTransfer();
  dt.items.add(f);
  resumeInput.files = dt.files;
  updateFileUI();
});

resetBtn.addEventListener("click", () => {
  form.reset();
  resumeInput.value = "";
  updateFileUI();
  statusEl.className = "status empty";
  statusEl.textContent = "Fill the form and click Send Email.";
});

form.addEventListener("submit", async (e) => {
  e.preventDefault();

  const email = $("#email").value.trim();
  if (!email) {
    setStatus("bad", "<strong>Email is required.</strong>");
    return;
  }

  setLoading(true);
  setStatus("empty", "Sending…");

  const fd = new FormData();
  fd.set("email", email);
  fd.set("name", $("#name").value.trim());
  fd.set("subject", $("#subject").value.trim());
  fd.set("body", $("#body").value.trim());

  const f = resumeInput.files && resumeInput.files[0];
  if (f) fd.set("resume", f);

  try {
    const res = await fetch("/api/send", { method: "POST", body: fd });
    const data = await res.json().catch(() => ({}));
    if (!res.ok || !data.ok) {
      const err = data.error || `Request failed (${res.status})`;
      setStatus("bad", `<strong>Failed.</strong><br/>${escapeHtml(err)}`);
      toast("bad", "Email failed", err);
      return;
    }

    const defaults = data.usedDefaults || {};
    setStatus(
      "good",
      `<strong>Sent!</strong><br/>
      To: <code>${escapeHtml(data.toEmail)}</code><br/>
      Subject: <code>${escapeHtml(data.subject || "")}</code><br/>
      <div style="margin-top:10px;color:rgba(255,255,255,.75)">
        Used defaults: subject=${defaults.subject ? "yes" : "no"}, body=${
        defaults.body ? "yes" : "no"
      }, resume=${defaults.resume ? "yes" : "no"}
      </div>`,
    );
    toast("good", "Email sent", data.toEmail);
  } catch (err) {
    const msg = String(err?.message || err);
    setStatus("bad", `<strong>Error.</strong><br/>${escapeHtml(msg)}`);
    toast("bad", "Error", msg);
  } finally {
    setLoading(false);
  }
});

logoutBtn?.addEventListener("click", async () => {
  try {
    await fetch("/api/logout", { method: "POST" });
  } catch {}
  window.location.href = "/login";
});

function renderHrResults(contacts, meta = {}) {
  lastHrContacts = Array.isArray(contacts) ? contacts : [];
  const phones = (lastHrContacts || [])
    .map((c) => String(c?.phone || "").trim())
    .filter(Boolean);
  const uniqPhones = Array.from(new Set(phones));
  lastHrPhone = meta.phone || uniqPhones.join("\n") || "";
  if (!lastHrContacts.length) {
    hrResults.className = "status empty";
    hrResults.innerHTML = "No HR / Talent contacts found.";
    return;
  }

  const cards = lastHrContacts
    .map((c) => {
      const name = c.name ? escapeHtml(c.name) : "Hiring Team";
      const pos = c.position ? escapeHtml(c.position) : "HR / Talent";
      const email = c.email ? escapeHtml(c.email) : "—";
      const phone = c.phone ? escapeHtml(c.phone) : "";
      const sendName = escapeHtml(String(c.name || "Hiring Team"));
      const conf =
        c.confidence === null || c.confidence === undefined ? "—" : escapeHtml(String(c.confidence));
      return `
        <div class="hrCard">
          <div class="hrName">${name}</div>
          <div class="hrRole">${pos}</div>
          <div class="hrEmailRow">
            <code class="hrEmail" title="${email}">${email}</code>
            <button class="iconBtn js-copy-email" type="button" data-email="${email}" title="Copy email">
              <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                <path d="M9 9h10v10H9V9Z" stroke="currentColor" stroke-width="2" />
                <path d="M5 15H4a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v1" stroke="currentColor" stroke-width="2" />
              </svg>
            </button>
            <button class="iconBtn js-send-hr-email" type="button" data-email="${email}" data-name="${sendName}" title="Send email">
              <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                <path d="M22 2L11 13" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                <path d="M22 2L15 22L11 13L2 9L22 2Z" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
              </svg>
            </button>
          </div>
          ${
            phone
              ? `<div class="hrEmailRow" style="margin-top:8px">
                  <code class="hrEmail" title="${phone}">${phone}</code>
                  <button class="iconBtn js-copy-phone" type="button" data-phone="${phone}" title="Copy phone">
                    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                      <path d="M9 9h10v10H9V9Z" stroke="currentColor" stroke-width="2" />
                      <path d="M5 15H4a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v1" stroke="currentColor" stroke-width="2" />
                    </svg>
                  </button>
                </div>`
              : ""
          }
          <div class="hrBottomRow">
            <span class="hrBadge">Confidence</span>
            <span class="hrBadge">${conf}</span>
          </div>
        </div>
      `;
    })
    .join("");

  hrResults.className = "status";
  hrResults.innerHTML = `
    <div style="margin-bottom:10px;color:rgba(255,255,255,.75)">
      <div class="hrMetaRow">
        <div>
          Found <strong>${lastHrContacts.length}</strong> contacts for
          <code>${escapeHtml(meta.domain || meta.company || "—")}</code>
        </div>
        ${
          meta.phone
            ? `<div style="color:rgba(255,255,255,.78)">Company phone: <code>${escapeHtml(
                meta.phone,
              )}</code></div>`
            : ""
        }
      </div>
      ${
        meta.mode === "all_emails_fallback"
          ? `<div style="margin-top:6px;color:rgba(255,211,109,.9)"><strong>Note:</strong> HR/TA roles not available for this domain; showing all discovered emails.</div>`
          : ""
      }
    </div>
    <div class="hrCards">${cards}</div>
  `;
}

hrResults?.addEventListener("click", async (e) => {
  const btn = e.target?.closest?.(".js-copy-email");
  const phoneBtn = e.target?.closest?.(".js-copy-phone");
  const sendBtn = e.target?.closest?.(".js-send-hr-email");
  if (!btn && !phoneBtn && !sendBtn) return;

  if (sendBtn) {
    const email = String(sendBtn.getAttribute("data-email") || "").trim();
    const name = String(sendBtn.getAttribute("data-name") || "").trim();
    if (!email || email === "—") {
      toast("bad", "Send failed", "No email found");
      return;
    }

    const oldText = sendBtn.getAttribute("data-old-text") || "";
    if (!oldText) sendBtn.setAttribute("data-old-text", sendBtn.innerHTML);
    sendBtn.disabled = true;
    sendBtn.style.opacity = "0.7";
    sendBtn.innerHTML = `<span style="font-size:12px;font-weight:800">…</span>`;

    try {
      const fd = new FormData();
      fd.set("email", email);
      fd.set("name", name);
      // subject/body empty => defaults
      fd.set("subject", "");
      fd.set("body", "");

      const res = await fetch("/api/send", { method: "POST", body: fd });
      const data = await res.json().catch(() => ({}));
      if (!res.ok || !data.ok) {
        const err = data.error || `Request failed (${res.status})`;
        toast("bad", "Send failed", err);
        return;
      }
      toast("good", "Email sent", email);
    } catch (err) {
      const msg = String(err?.message || err);
      toast("bad", "Send failed", msg);
    } finally {
      const html = sendBtn.getAttribute("data-old-text");
      if (html) sendBtn.innerHTML = html;
      sendBtn.disabled = false;
      sendBtn.style.opacity = "";
    }
    return;
  }

  const val = btn
    ? btn.getAttribute("data-email") || ""
    : phoneBtn
      ? phoneBtn.getAttribute("data-phone") || ""
      : "";
  const label = btn ? "email" : "phone";

  if (!val || val === "—") {
    toast("bad", "Copy failed", `No ${label} to copy`);
    return;
  }
  try {
    await navigator.clipboard.writeText(val);
    toast("good", "Copied", val);
  } catch {
    toast("bad", "Copy failed", "Browser blocked clipboard. Copy manually.");
  }
});

hrSearchBtn?.addEventListener("click", async () => {
  const company = ($("#company")?.value || "").trim();
  const domain = ($("#domain")?.value || "").trim();
  const provider = (providerSel?.value || "hunter").trim();

  if (!company && !domain) {
    toast("bad", "Missing input", "Enter company name or domain.");
    return;
  }

  hrResults.className = "status empty";
  hrResults.innerHTML = "Searching…";

  try {
    const qs = new URLSearchParams();
    if (company) qs.set("company", company);
    if (domain) qs.set("domain", domain);
    if (provider) qs.set("provider", provider);
    const res = await fetch(`/api/hr-lookup?${qs.toString()}`);
    const data = await res.json().catch(() => ({}));
    if (!res.ok || !data.ok) {
      const err = data.error || `Request failed (${res.status})`;
      hrResults.className = "status bad";
      hrResults.innerHTML = `<strong>Failed.</strong><br/>${escapeHtml(err)}`;
      toast("bad", "HR lookup failed", err);
      return;
    }
    renderHrResults(data.contacts || [], {
      domain: data.domain,
      company: data.company,
      mode: data.mode,
      phone: data.phone,
    });
    toast("good", "HR lookup complete", `${(data.contacts || []).length} contacts found`);
  } catch (e) {
    const msg = String(e?.message || e);
    hrResults.className = "status bad";
    hrResults.innerHTML = `<strong>Error.</strong><br/>${escapeHtml(msg)}`;
    toast("bad", "Error", msg);
  }
});

bulkSendBtn?.addEventListener("click", async () => {
  const mode = String(bulkModeSel?.value || "excel");

  setLoading(true);
  setStatus("empty", "Sending bulk emails… (this may take a bit)");

  const fd = new FormData();
  if (mode === "excel") {
    const excel = excelInput?.files?.[0];
    if (!excel) {
      setLoading(false);
      setStatus("bad", "<strong>Excel (.xlsx) is required for bulk send.</strong>");
      return;
    }
    if (!excel.name.toLowerCase().endsWith(".xlsx")) {
      setLoading(false);
      setStatus("bad", "<strong>Please upload a valid .xlsx file.</strong>");
      return;
    }
    fd.set("excel", excel);
  } else {
    const raw = String(bulkEmailsTa?.value || "").trim();
    if (!raw) {
      setLoading(false);
      setStatus("bad", "<strong>Please paste at least one email.</strong>");
      return;
    }
    fd.set("emails", raw);
  }

  const bulkResume = bulkResumeInput?.files?.[0];
  if (bulkResume) fd.set("resume", bulkResume);

  try {
    const res = await fetch(mode === "excel" ? "/api/send-bulk" : "/api/send-list", {
      method: "POST",
      body: fd,
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok || !data.ok) {
      const err = data.error || `Request failed (${res.status})`;
      setStatus("bad", `<strong>Bulk send failed.</strong><br/>${escapeHtml(err)}`);
      toast("bad", "Bulk failed", err);
      return;
    }

    const failedLines =
      (data.results || [])
        .filter((r) => !r.ok)
        .slice(0, 8)
        .map((r) => `<li><code>${escapeHtml(r.email)}</code> — ${escapeHtml(r.error || "failed")}</li>`)
        .join("") || "";

    const sentLines =
      (data.results || [])
        .filter((r) => r.ok)
        .slice(0, 8)
        .map(
          (r) =>
            `<li><code>${escapeHtml(r.email)}</code> — <span style="color:rgba(109,255,181,.9)">sent</span></li>`,
        )
        .join("") || "";

    setStatus(
      data.failed ? "bad" : "good",
      `<strong>Bulk done.</strong><br/>
      Total: <code>${data.total}</code> | Sent: <code>${data.sent}</code> | Failed: <code>${data.failed}</code>
      ${
        sentLines
          ? `<div style="margin-top:10px;color:rgba(255,255,255,.75)"><strong>Sent (sample):</strong><ul style="margin:6px 0 0 18px">${sentLines}</ul></div>`
          : ""
      }
      ${
        failedLines
          ? `<div style="margin-top:10px;color:rgba(255,255,255,.75)"><strong>Some failures:</strong><ul style="margin:6px 0 0 18px">${failedLines}</ul></div>`
          : ""
      }`,
    );
    toast(
      data.failed ? "bad" : "good",
      "Bulk complete",
      `Sent ${data.sent}/${data.total} (${data.failed} failed)`,
      { timeoutMs: 4500 },
    );
  } catch (err) {
    const msg = String(err?.message || err);
    setStatus("bad", `<strong>Error.</strong><br/>${escapeHtml(msg)}`);
    toast("bad", "Error", msg);
  } finally {
    setLoading(false);
  }
});

function setBulkMode(mode) {
  const m = String(mode || "excel");
  bulkExcelWrap?.classList.toggle("hidden", m !== "excel");
  bulkPasteWrap?.classList.toggle("hidden", m !== "paste");
}

setBulkMode(bulkModeSel?.value || "excel");
bulkModeSel?.addEventListener("change", () => setBulkMode(bulkModeSel?.value || "excel"));

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}


