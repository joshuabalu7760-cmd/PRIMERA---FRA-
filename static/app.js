const state = {
  uploadId: null,
  rows: [],
  reports: [],
  analysisId: null,
  filter: "all",
  selected: new Set(),
  authId: null,
  templates: [],
  activeTemplate: null,
};
const $ = (selector) => document.querySelector(selector);

function showToast(message, isError = false) {
  const toast = $("#toast");
  toast.textContent = message;
  toast.classList.toggle("error", isError);
  toast.classList.add("show");
  window.clearTimeout(showToast.timeout);
  showToast.timeout = window.setTimeout(() => toast.classList.remove("show"), 4200);
}

async function postJson(url, body) {
  const response = await fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || "The request could not be completed.");
  return data;
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[character]);
}

function setStep(step) {
  document.querySelectorAll(".step").forEach((item) => item.classList.toggle("active", Number(item.dataset.step) <= step));
}

function matchingReport(row) {
  const cell = (row.report_cell || "").toLowerCase();
  return cell ? state.reports.find((report) => cell.includes(report.name.toLowerCase()) || report.name.toLowerCase().includes(cell)) : null;
}

function renderImportRows() {
  const body = $("#import-body");
  body.replaceChildren();
  for (const row of state.rows) {
    const matched = matchingReport(row);
    if (matched && !row.report_id) row.report_id = matched.id;
    const options = ['<option value="">Choose embedded PDF…</option>', ...state.reports.map((report) => `<option value="${report.id}" ${row.report_id === report.id ? "selected" : ""}>${escapeHtml(report.name)}${report.has_work === true ? " · works found" : report.has_work === false ? " · no works" : " · review"}</option>`)].join("");
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><span class="property-name">${escapeHtml(row.property)}</span><span class="agent-name">${escapeHtml(row.postcode)}</span></td>
      <td><span class="property-name">${escapeHtml(row.managing_agent)}</span><span class="agent-name">${escapeHtml(row.recipient)}</span></td>
      <td><select class="report-select" aria-label="FRA PDF for ${escapeHtml(row.property)}">${options}</select></td>
      <td><span class="date-cell">${escapeHtml(row.inspection_date || "Not recorded")}</span><span class="agent-name">Due ${escapeHtml(row.due_date || "not recorded")}</span></td>
      <td><span class="notes-cell" title="${escapeHtml(row.notes)}">${escapeHtml(row.notes || "—")}</span></td>`;
    tr.querySelector(".report-select").addEventListener("change", (event) => { row.report_id = event.target.value; });
    body.append(tr);
  }
  $("#row-total").textContent = `${state.rows.length} PROPERTIES · ${state.reports.length} EMBEDDED PDF${state.reports.length === 1 ? "" : "S"}`;
  if (!state.reports.length) {
    $("#attachment-note").textContent = "No embedded PDF was detected in this workbook. Make sure each PDF is inserted as an embedded file object, not just displayed as an icon or linked from a shared drive.";
    $("#attachment-note").classList.add("warning-note");
  } else {
    $("#attachment-note").textContent = "Assign the correct PDF to every property. Rows with a matching filename in the FRA Report cell are preselected; check each match before continuing.";
    $("#attachment-note").classList.remove("warning-note");
  }
}

async function uploadSpreadsheet(file) {
  if (!file) return;
  state.uploadId = null;
  state.analysisId = null;
  state.rows = [];
  state.reports = [];
  state.selected.clear();
  $("#results-section").classList.add("hidden");
  $("#email-panel").classList.add("hidden");
  $("#integration-errors").classList.add("hidden");
  $("#file-status").textContent = `Reading ${file.name}…`;
  $("#mapping-error").textContent = "";
  const form = new FormData();
  form.append("spreadsheet", file);
  try {
    const response = await fetch("/api/spreadsheet", { method: "POST", body: form });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Could not read the workbook.");
    state.uploadId = data.upload_id;
    state.rows = data.rows;
    state.reports = data.reports;
    state.rows.forEach((row) => { row.report_id = matchingReport(row)?.id || ""; });
    $("#row-total").textContent = `${data.row_count} PROPERTIES`;
    $("#file-status").textContent = `${file.name} · ${data.row_count} property rows · first worksheet`;
    $("#mapping-area").classList.remove("hidden");
    renderImportRows();
    setStep(1);
    $("#mapping-area").scrollIntoView({ behavior: "smooth", block: "nearest" });
  } catch (error) {
    $("#file-status").textContent = "";
    showToast(error.message, true);
  }
}

$("#spreadsheet").addEventListener("change", (event) => uploadSpreadsheet(event.target.files[0]));
const dropzone = $("#dropzone");
dropzone.addEventListener("dragover", (event) => { event.preventDefault(); dropzone.classList.add("dragover"); });
dropzone.addEventListener("dragleave", () => dropzone.classList.remove("dragover"));
dropzone.addEventListener("drop", (event) => { event.preventDefault(); dropzone.classList.remove("dragover"); uploadSpreadsheet(event.dataTransfer.files[0]); });

$("#analyze-button").addEventListener("click", async () => {
  const button = $("#analyze-button");
  $("#mapping-error").textContent = "";
  button.disabled = true;
  button.firstElementChild.textContent = "Analyzing PDFs…";
  try {
    const reportMap = Object.fromEntries(state.rows.map((row) => [row.row_number, row.report_id]));
    const data = await postJson("/api/analyze", { upload_id: state.uploadId, report_map: reportMap });
    state.analysisId = data.analysis_id;
    state.rows = data.results.map((row) => ({ ...row, status: row.report?.has_work === true ? "works" : row.report?.has_work === false ? "clear" : "review" }));
    state.selected = new Set(state.rows.filter((row) => row.status === "works" && row.report && row.recipient).map((row) => row.row_number));
    renderResults();
    $("#results-section").classList.remove("hidden");
    setStep(2);
    $("#results-section").scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (error) {
    $("#mapping-error").textContent = error.message;
  } finally {
    button.disabled = false;
    button.firstElementChild.textContent = "Analyze reports";
  }
});

function visibleRows() {
  const search = $("#search").value.trim().toLowerCase();
  return state.rows.filter((row) => {
    const fitsFilter = state.filter === "all" || row.status === state.filter;
    const fitsSearch = !search || `${row.property} ${row.postcode} ${row.managing_agent} ${row.recipient}`.toLowerCase().includes(search);
    return fitsFilter && fitsSearch;
  });
}

function renderResults() {
  const body = $("#results-body");
  body.replaceChildren();
  const rows = visibleRows();
  $("#empty-state").classList.toggle("hidden", rows.length > 0);
  for (const row of rows) {
    const tr = document.createElement("tr");
    const checked = state.selected.has(row.row_number);
    const recipient = row.recipient ? escapeHtml(row.recipient) : '<span class="missing-report">No email recorded</span>';
    const report = row.report ? `<span class="report-name" title="${escapeHtml(row.report.name)}"><span class="pdf-icon">PDF</span>${escapeHtml(row.report.name)}</span>` : '<span class="missing-report">No PDF assigned</span>';
    tr.innerHTML = `<td><input type="checkbox" class="row-check" aria-label="Select ${escapeHtml(row.property)}" ${checked ? "checked" : ""} ${!row.report || !row.recipient || row.status !== "works" ? "disabled" : ""}></td>
      <td><span class="property-name">${escapeHtml(row.property)}</span><span class="agent-name">${escapeHtml(row.postcode)}</span></td>
      <td><span class="property-name">${escapeHtml(row.managing_agent)}</span><span class="agent-name">${recipient}</span></td>
      <td>${report}</td>
      <td><select class="status-select ${row.status}" aria-label="Works status for ${escapeHtml(row.property)}"><option value="works" ${row.status === "works" ? "selected" : ""}>Works found</option><option value="review" ${row.status === "review" ? "selected" : ""}>Review report</option><option value="clear" ${row.status === "clear" ? "selected" : ""}>No works</option></select><span class="evidence">${escapeHtml(row.report?.evidence || "Assign a report to assess works.")}</span></td>
      <td><span class="date-cell">${escapeHtml(row.inspection_date || "—")}</span><span class="agent-name">Due ${escapeHtml(row.due_date || "—")}</span></td>
      <td><span class="notes-cell" title="${escapeHtml(row.notes)}">${escapeHtml(row.notes || "—")}</span></td>`;
    tr.querySelector(".row-check").addEventListener("change", (event) => {
      if (event.target.checked) state.selected.add(row.row_number);
      else state.selected.delete(row.row_number);
      updateSelectionNote();
    });
    tr.querySelector(".status-select").addEventListener("change", (event) => {
      row.status = event.target.value;
      if (row.status !== "works") state.selected.delete(row.row_number);
      else if (row.report && row.recipient) state.selected.add(row.row_number);
      renderResults();
    });
    body.append(tr);
  }
  const counts = { all: state.rows.length, works: 0, review: 0, clear: 0 };
  state.rows.forEach((row) => counts[row.status]++);
  $("#count-all").textContent = counts.all;
  $("#count-works").textContent = counts.works;
  $("#count-review").textContent = counts.review;
  $("#count-clear").textContent = counts.clear;
  $("#summary-stats").innerHTML = `<div><strong>${counts.works}</strong><span>works found</span></div><div><strong>${counts.review}</strong><span>need review</span></div><div><strong>${state.rows.filter((row) => !row.report).length}</strong><span>report missing</span></div>`;
  updateSelectionNote();
}

function updateSelectionNote() {
  const count = state.selected.size;
  $("#selection-note").textContent = `${count} propert${count === 1 ? "y" : "ies"} selected`;
}

document.querySelectorAll(".filter").forEach((button) => button.addEventListener("click", () => {
  state.filter = button.dataset.filter;
  document.querySelectorAll(".filter").forEach((filter) => filter.classList.toggle("active", filter === button));
  renderResults();
}));
$("#search").addEventListener("input", renderResults);
$("#select-visible").addEventListener("change", (event) => {
  visibleRows().filter((row) => row.status === "works" && row.report && row.recipient).forEach((row) => {
    if (event.target.checked) state.selected.add(row.row_number);
    else state.selected.delete(row.row_number);
  });
  renderResults();
});

$("#open-email").addEventListener("click", () => {
  const ready = state.rows.filter((row) => state.selected.has(row.row_number) && row.status === "works" && row.report && row.recipient);
  if (!ready.length) return showToast("Select at least one works-found property with a report and email address.", true);
  $("#send-count").textContent = ready.length;
  $("#send-list").innerHTML = ready.map((row) => `<div>${escapeHtml(row.property)}<span>${escapeHtml(row.recipient)} · ${escapeHtml(row.report.name)}</span></div>`).join("");
  $("#email-panel").classList.remove("hidden");
  $("#email-panel").scrollIntoView({ behavior: "smooth", block: "start" });
  setStep(3);
});
$("#back-review").addEventListener("click", () => { $("#email-panel").classList.add("hidden"); setStep(2); $("#results-section").scrollIntoView({ behavior: "smooth", block: "start" }); });

async function refreshTemplates(selectedId = "") {
  const response = await fetch("/api/templates");
  const data = await response.json();
  state.templates = data.templates || [];
  const select = $("#email-template");
  select.replaceChildren(new Option(state.templates.length ? "Choose a template…" : "Import a template to begin", ""));
  state.templates.forEach((template) => select.add(new Option(template.name, template.id)));
  if (selectedId && state.templates.some((template) => template.id === selectedId)) select.value = selectedId;
  else if (state.templates.length === 1) select.value = state.templates[0].id;
  $("#delete-template").disabled = !select.value;
  await showTemplatePreview(select.value);
}

async function showTemplatePreview(templateId) {
  const frame = $("#template-preview");
  state.activeTemplate = null;
  if (!templateId) {
    frame.classList.add("hidden");
    return;
  }
  try {
    const response = await fetch(`/api/templates/${templateId}`);
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Could not load this template.");
    state.activeTemplate = data;
    const content = data.content_type === "html" ? data.body : `<pre>${escapeHtml(data.body)}</pre>`;
    frame.srcdoc = `<!doctype html><meta charset="utf-8"><style>body{font:14px Arial,sans-serif;color:#24352d;padding:16px;line-height:1.5}pre{white-space:pre-wrap;font:inherit}</style>${content}`;
    frame.classList.remove("hidden");
    $("#delete-template").disabled = false;
  } catch (error) {
    showToast(error.message, true);
  }
}

$("#email-template").addEventListener("change", async (event) => {
  $("#delete-template").disabled = !event.target.value;
  await showTemplatePreview(event.target.value);
});
$("#template-files").addEventListener("change", async (event) => {
  const files = [...event.target.files];
  if (!files.length) return;
  const form = new FormData();
  files.forEach((file) => form.append("templates", file));
  try {
    const response = await fetch("/api/templates", { method: "POST", body: form });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Could not import templates.");
    await refreshTemplates(data.imported?.[0]?.id || "");
    if (data.errors?.length) showToast(`Imported ${data.imported.length}; ${data.errors.join(" ")}`, data.imported.length === 0);
    else showToast(`Imported ${data.imported.length} Outlook template${data.imported.length === 1 ? "" : "s"}.`);
  } catch (error) {
    showToast(error.message, true);
  } finally {
    event.target.value = "";
  }
});
$("#delete-template").addEventListener("click", async () => {
  const templateId = $("#email-template").value;
  if (!templateId || !window.confirm("Remove this template from the app?")) return;
  const response = await fetch(`/api/templates/${templateId}`, { method: "DELETE" });
  const data = await response.json();
  if (!response.ok) return showToast(data.error || "Could not remove template.", true);
  await refreshTemplates();
  showToast("Template removed.");
});

$("#connect-button").addEventListener("click", async () => {
  const button = $("#connect-button");
  button.disabled = true;
  if (state.authId) {
    try {
      await postJson("/api/auth/disconnect", { auth_id: state.authId });
      state.authId = null;
      $("#connection-status").textContent = "Not connected";
      $("#auth-message").textContent = "Disconnected from Koritsu. Connect again whenever you need to send.";
      button.firstElementChild.textContent = "Connect Microsoft 365";
    } catch (error) {
      $("#auth-message").textContent = error.message;
    } finally {
      button.disabled = false;
    }
    return;
  }
  $("#auth-message").textContent = "Starting Microsoft sign-in…";
  try {
    const data = await postJson("/api/auth/start", {
      client_id: $("#graph-client-id").value.trim(),
      tenant_id: $("#graph-tenant-id").value.trim(),
    });
    state.authId = data.auth_id;
    $("#auth-message").textContent = data.message;
    pollAuthStatus(data.auth_id);
  } catch (error) {
    $("#auth-message").textContent = error.message;
    button.disabled = false;
  }
});

async function pollAuthStatus(authId) {
  try {
    const response = await fetch(`/api/auth/status/${authId}`);
    const data = await response.json();
    if (data.status === "connected") {
      state.authId = authId;
      $("#connection-status").textContent = "Connected to Microsoft 365";
      $("#auth-message").textContent = "Connected. Disconnect whenever you want to remove Koritsu’s local session.";
      $("#connect-button").firstElementChild.textContent = "Disconnect Microsoft 365";
      $("#connect-button").disabled = false;
      return;
    }
    if (data.status === "error") {
      state.authId = null;
      $("#connection-status").textContent = "Not connected";
      $("#auth-message").textContent = data.error || "Microsoft sign-in failed.";
      $("#connect-button").firstElementChild.textContent = "Connect Microsoft 365";
      $("#connect-button").disabled = false;
      return;
    }
    window.setTimeout(() => pollAuthStatus(authId), 2500);
  } catch (error) {
    $("#auth-message").textContent = error.message;
    $("#connect-button").disabled = false;
  }
}

$("#send-button").addEventListener("click", async () => {
  const button = $("#send-button");
  $("#send-error").textContent = "";
  if (!state.activeTemplate) return $("#send-error").textContent = "Choose an Outlook template first.";
  if (!state.authId) return $("#send-error").textContent = "Connect to Microsoft 365 first.";
  button.disabled = true;
  button.firstElementChild.textContent = "Sending…";
  const rowNumbers = state.rows.filter((row) => state.selected.has(row.row_number) && row.status === "works").map((row) => row.row_number);
  try {
    const result = await postJson("/api/send", {
      analysis_id: state.analysisId,
      auth_id: state.authId,
      template_id: state.activeTemplate.id,
      row_numbers: rowNumbers,
      reviewed_statuses: Object.fromEntries(state.rows.map((row) => [row.row_number, row.status])),
    });
    if (result.failures?.length) {
      result.sent.forEach((rowNumber) => state.selected.delete(rowNumber));
      const remaining = state.rows.filter((row) => state.selected.has(row.row_number));
      $("#send-count").textContent = remaining.length;
      $("#send-list").innerHTML = remaining.map((row) => `<div>${escapeHtml(row.property)}<span>${escapeHtml(row.recipient)} · ${escapeHtml(row.report.name)}</span></div>`).join("");
      $("#send-error").textContent = `${result.sent.length} sent; ${result.failures.length} failed. ${result.failures.map((failure) => failure.error).join("; ")}`;
      button.disabled = false;
      button.firstElementChild.textContent = "Retry failed reports";
      updateSelectionNote();
    } else {
      showToast(`${result.sent.length} report email${result.sent.length === 1 ? "" : "s"} sent.`);
      button.firstElementChild.textContent = "Batch sent";
    }
  } catch (error) {
    $("#send-error").textContent = error.message;
    button.disabled = false;
    button.firstElementChild.textContent = "Send selected reports";
  }
});

refreshTemplates().catch((error) => showToast(error.message, true));