const prepareForm = document.getElementById("prepare-form");
const reviewForm = document.getElementById("review-form");
const reviewPanel = document.getElementById("review-panel");
const preparePanel = document.getElementById("prepare-panel");
const reviewDirect = document.getElementById("review-direct");
const refreshReviewButton = document.getElementById("refresh-review-button");
const reviewLink = document.getElementById("review-link");
const draftLink = document.getElementById("draft-link");
const downloadLink = document.getElementById("download-link");
const exceptionsLink = document.getElementById("exceptions-link");
const newJobButton = document.getElementById("new-job-button");
const statusBox = document.getElementById("status-box");
const health = document.getElementById("health");
const storageKey = "document-csv-parser:weekly-job-id";
const reviewRowsElement = document.getElementById("review-rows");
const reviewCount = document.getElementById("review-count");
const reviewSearch = document.getElementById("review-search");
const reviewUnchecked = document.getElementById("review-unchecked");
const reviewEmpty = document.getElementById("review-empty");
const reviewFile = document.getElementById("catalyst_review");
const reviewButton = document.getElementById("review-button");
const sourceReviewLink = document.getElementById("source-review-link");
const sourceReviewFileField = document.getElementById("source-review-file-field");
const sourceReviewFile = document.getElementById("source_review");
const includeMobile = document.getElementById("include_mobile");
const includeM2m = document.getElementById("include_m2m");
let activeJobId = null;
let pollingTimer = null;
let requestPending = false;
let reviewRows = null;
let loadedReviewJobId = null;
let reviewSelections = {};
let reviewEditable = false;
let currentJobStatus = null;

function syncOutputSelection() {
  includeMobile.setCustomValidity(includeMobile.checked || includeM2m.checked ? "" : "Select at least one report.");
  document.getElementById("output_services").value = includeMobile.checked && includeM2m.checked
    ? "both" : includeM2m.checked ? "m2m" : "mobile";
  for (const [id, selected] of [["mobile-template-fields", includeMobile.checked], ["m2m-template-fields", includeM2m.checked]]) {
    const fields = document.getElementById(id);
    fields.hidden = !selected;
    for (const input of fields.querySelectorAll("input")) {
      input.disabled = !selected;
      input.required = selected;
    }
  }
}

function normalizeJobId(jobId) {
  if (!jobId) return jobId;
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(jobId)
    ? jobId.replaceAll("-", "").toLowerCase()
    : jobId;
}

function reviewStorageKey(jobId) {
  return `document-csv-parser:weekly-review:${jobId}`;
}

function resetReviewView() {
  reviewRows = null;
  loadedReviewJobId = null;
  reviewSelections = {};
  reviewRowsElement.replaceChildren();
  reviewCount.textContent = "Loading Catalyst rows...";
  reviewSearch.value = "";
  reviewUnchecked.checked = false;
  reviewEmpty.hidden = true;
  reviewButton.disabled = true;
  reviewDirect.hidden = false;
  refreshReviewButton.hidden = true;
  currentJobStatus = null;
}

function saveReviewSelections() {
  try { window.localStorage.setItem(reviewStorageKey(activeJobId), JSON.stringify(reviewSelections)); }
  catch (_error) { /* storage optional */ }
}

function restoreReviewSelections(jobId) {
  try {
    const saved = JSON.parse(window.localStorage.getItem(reviewStorageKey(jobId)) || "{}");
    return saved && typeof saved === "object" && !Array.isArray(saved) ? saved : {};
  } catch (_error) { return {}; }
}

function updateReviewProgress() {
  const total = reviewRows?.length || 0;
  const completed = reviewRows?.filter((row) => Boolean(reviewSelections[row.master_row])).length || 0;
  if (reviewRows) reviewCount.textContent = `${completed} of ${total} Catalyst statuses checked`;
  if (reviewRows && currentJobStatus === "waiting_review" && !requestPending) {
    setStatus(`Master draft ready. ${completed} of ${total} Catalyst statuses checked.`, "ok");
  }
  reviewButton.disabled = !reviewEditable || requestPending ||
    (!reviewFile.files.length && (!reviewRows || completed !== total));
}

function filterReviewRows() {
  const query = reviewSearch.value.trim().toLocaleLowerCase();
  let visible = 0;
  for (const row of reviewRowsElement.rows) {
    const matches = row.dataset.search.includes(query) &&
      (!reviewUnchecked.checked || !reviewSelections[row.dataset.masterRow]);
    row.hidden = !matches;
    if (matches) visible += 1;
  }
  reviewEmpty.hidden = visible !== 0 || !reviewRows?.length;
}

function renderReviewRows() {
  const fragment = document.createDocumentFragment();
  for (const row of reviewRows) {
    const tr = document.createElement("tr");
    tr.dataset.masterRow = row.master_row;
    tr.dataset.search = [row.quote_num, row.order_number, row.ca_name, row.service_group]
      .join(" ").toLocaleLowerCase();
    const identity = document.createElement("td");
    const quote = document.createElement("strong");
    quote.textContent = row.quote_num || "(no quote)";
    const order = document.createElement("small");
    order.textContent = row.order_number || `Master row ${row.master_row}`;
    identity.append(quote, order);
    const customer = document.createElement("td");
    customer.textContent = row.ca_name || "-";
    const service = document.createElement("td");
    service.textContent = row.service_group || "-";
    const statusCell = document.createElement("td");
    const select = document.createElement("select");
    select.dataset.masterRow = row.master_row;
    select.setAttribute("aria-label", `Catalyst status for quote ${row.quote_num}, order ${row.order_number}`);
    for (const [value, label] of [
      ["", "Select status"], ["In Progress", "In Progress"],
      ["COMPLETE", "Complete"], ["CANCEL", "Cancelled"],
    ]) {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = label;
      select.append(option);
    }
    select.value = reviewSelections[row.master_row] || "";
    select.disabled = !reviewEditable;
    statusCell.append(select);
    tr.append(identity, customer, service, statusCell);
    fragment.append(tr);
  }
  reviewRowsElement.replaceChildren(fragment);
  updateReviewProgress();
  filterReviewRows();
}

async function loadReviewData(job) {
  if (loadedReviewJobId === job.job_id || !job.review_data_url) return;
  let payload;
  try {
    payload = await requestJson(job.review_data_url, { cache: "no-store" });
  } catch (error) {
    reviewCount.textContent = "Direct review unavailable for this job";
    if (job.status !== "review_failed") setStatus(error.message, "bad");
    return;
  }
  if (activeJobId !== normalizeJobId(job.job_id)) return;
  reviewRows = payload.rows;
  loadedReviewJobId = job.job_id;
  const saved = restoreReviewSelections(job.job_id);
  reviewSelections = {};
  for (const row of reviewRows) {
    if (["In Progress", "COMPLETE", "CANCEL"].includes(saved[row.master_row])) {
      reviewSelections[row.master_row] = saved[row.master_row];
    }
  }
  renderReviewRows();
}

function setStatus(message, kind = "") {
  statusBox.className = `status-box${kind ? ` ${kind}` : ""}`;
  statusBox.textContent = message;
}

function setFormDisabled(form, disabled) {
  for (const control of form.elements) control.disabled = disabled;
}

function rememberJob(jobId) {
  activeJobId = normalizeJobId(jobId);
  try { window.localStorage.setItem(storageKey, activeJobId); } catch (_error) { /* storage optional */ }
}

function forgetJob() {
  if (activeJobId) {
    try { window.localStorage.removeItem(reviewStorageKey(activeJobId)); } catch (_error) { /* storage optional */ }
  }
  activeJobId = null;
  try { window.localStorage.removeItem(storageKey); } catch (_error) { /* storage optional */ }
  resetReviewView();
}

function stopPolling() {
  if (pollingTimer !== null) window.clearTimeout(pollingTimer);
  pollingTimer = null;
}

async function requestJson(url, options = {}) {
  const response = await fetch(url, { ...options, headers: { Accept: "application/json", ...(options.headers || {}) } });
  const text = await response.text();
  let payload;
  try { payload = text ? JSON.parse(text) : {}; } catch (_error) { throw new Error(`Server returned HTTP ${response.status} without JSON.`); }
  if (!response.ok) {
    const detail = payload.detail;
    const message = Array.isArray(detail)
      ? detail.map((item) => {
        const location = Array.isArray(item.loc)
          ? item.loc.filter((part) => part !== "body").join(".")
          : "";
        const reason = item.msg || JSON.stringify(item);
        return location ? `${location}: ${reason}` : reason;
      }).join("; ")
      : detail && typeof detail === "object"
        ? JSON.stringify(detail)
        : detail || payload.error || `HTTP ${response.status}`;
    const error = new Error(message);
    error.status = response.status;
    throw error;
  }
  return payload;
}

function renderJob(job) {
  const status = job.status;
  currentJobStatus = status;
  const preparing = ["uploading", "queued", "preparing", "refreshing_review"].includes(status);
  const finalizing = ["uploading_review", "queued_finalization", "finalizing", "generating"].includes(status);
  setFormDisabled(prepareForm, true);
  reviewPanel.hidden = !["waiting_review", "refreshing_review", "uploading_review", "queued_finalization", "finalizing", "generating", "review_failed", "succeeded"].includes(status);
  preparePanel.hidden = !reviewPanel.hidden;
  refreshReviewButton.hidden = !job.review_refresh_url;
  if (job.review_refresh_url) refreshReviewButton.dataset.url = job.review_refresh_url;
  reviewDirect.hidden = Boolean(job.review_refresh_url) || status === "refreshing_review";
  reviewLink.href = job.review_url || "#";
  sourceReviewLink.href = job.source_review_url || "#";
  sourceReviewLink.hidden = !job.source_review_url;
  sourceReviewFileField.hidden = !job.source_review_url;
  draftLink.href = job.draft_url || "#";
  draftLink.hidden = !job.draft_url;
  setFormDisabled(reviewForm, !["waiting_review", "review_failed"].includes(status));
  reviewEditable = ["waiting_review", "review_failed"].includes(status);
  updateReviewProgress();
  const exceptionsUrl = job.final_exceptions_url || job.preparation_exceptions_url;
  exceptionsLink.classList.toggle("visible", Boolean(exceptionsUrl));
  if (exceptionsUrl) exceptionsLink.href = exceptionsUrl;
  downloadLink.classList.toggle("visible", status === "succeeded");
  if (status === "succeeded") downloadLink.href = job.download_url;
  newJobButton.hidden = preparing || finalizing;
  if (["failed", "review_failed"].includes(status)) setStatus(job.error || "Weekly job failed.", "bad");
  else if (status === "waiting_review" && job.review_refresh_url) {
    setStatus("This job has an older review format. Refresh Catalyst Review to enter statuses directly.");
  }
  else if (status === "waiting_review") {
    const pending = job.preparation_audit?.pending_catalyst_rows;
    setStatus(`Master draft ready. ${pending ?? "Pending"} Catalyst statuses need review.`, "ok");
  }
  else if (status === "succeeded") {
    const issueCount = job.final_audit?.issue_counts
      ? Object.values(job.final_audit.issue_counts).reduce((a, b) => a + b, 0)
      : 0;
    const warning = job.warnings?.length ? ` ${job.warnings.join(" ")}` : "";
    setStatus(`Weekly reports ready. ${job.output_files?.length || 0} files. ${issueCount} audit exceptions.${warning}`, warning ? "" : "ok");
  }
  else setStatus(job.stage || "Processing...");
}

async function pollJob() {
  stopPolling();
  if (!activeJobId) return;
  try {
    const job = await requestJson(`/api/weekly/jobs/${activeJobId}`, { cache: "no-store" });
    renderJob(job);
    if (["waiting_review", "review_failed"].includes(job.status) && !job.review_refresh_url) {
      await loadReviewData(job);
    }
    if (!["waiting_review", "succeeded", "failed", "review_failed"].includes(job.status)) {
      pollingTimer = window.setTimeout(pollJob, 3000);
    }
  } catch (error) {
    setStatus(`Cannot read weekly job status: ${error.message}. Retrying...`, "bad");
    pollingTimer = window.setTimeout(pollJob, 5000);
  }
}

prepareForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (requestPending || activeJobId) return;
  syncOutputSelection();
  if (!prepareForm.reportValidity()) return;
  requestPending = true;
  const data = new FormData(prepareForm);
  setFormDisabled(prepareForm, true);
  setStatus("Uploading weekly source files...");
  const jobId = window.crypto.randomUUID();
  data.set("job_id", jobId);
  rememberJob(jobId);
  try {
    await requestJson("/api/weekly/jobs", { method: "POST", body: data });
    await pollJob();
  } catch (error) {
    setStatus(error.message, "bad");
    if ([400, 413, 422].includes(error.status)) {
      forgetJob();
      setFormDisabled(prepareForm, false);
      syncOutputSelection();
    } else {
      // A gateway can fail after the upload was accepted; keep the ID for polling.
      pollingTimer = window.setTimeout(pollJob, 5000);
    }
  } finally {
    requestPending = false;
    updateReviewProgress();
  }
});

refreshReviewButton.addEventListener("click", async () => {
  if (requestPending || !activeJobId || !refreshReviewButton.dataset.url) return;
  requestPending = true;
  refreshReviewButton.disabled = true;
  setStatus("Refreshing Catalyst review from saved inputs...");
  try {
    await requestJson(refreshReviewButton.dataset.url, { method: "POST" });
    await pollJob();
  } catch (error) {
    setStatus(error.message, "bad");
    await pollJob();
  } finally {
    requestPending = false;
    refreshReviewButton.disabled = false;
    updateReviewProgress();
  }
});

reviewForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (requestPending || !activeJobId || reviewButton.disabled) return;
  requestPending = true;
  const data = new FormData(reviewForm);
  if (reviewFile.files.length) {
    data.delete("reviewed_statuses");
  } else {
    data.delete("catalyst_review");
    data.set("reviewed_statuses", JSON.stringify(reviewSelections));
  }
  if (!sourceReviewFile.files.length) data.delete("source_review");
  if (!document.getElementById("slide_review").files.length) data.delete("slide_review");
  if (!document.getElementById("pivot_review").files.length) data.delete("pivot_review");
  if (!document.getElementById("presentation_review").files.length) data.delete("presentation_review");
  setFormDisabled(reviewForm, true);
  setStatus("Submitting Catalyst review...");
  try {
    await requestJson(`/api/weekly/jobs/${activeJobId}/finalize`, { method: "POST", body: data });
    await pollJob();
  } catch (error) {
    setStatus(error.message, "bad");
    await pollJob();
  } finally {
    requestPending = false;
    updateReviewProgress();
  }
});

reviewRowsElement.addEventListener("change", (event) => {
  const select = event.target;
  if (!(select instanceof HTMLSelectElement)) return;
  if (select.value) reviewSelections[select.dataset.masterRow] = select.value;
  else delete reviewSelections[select.dataset.masterRow];
  saveReviewSelections();
  updateReviewProgress();
  filterReviewRows();
});
reviewSearch.addEventListener("input", filterReviewRows);
reviewUnchecked.addEventListener("change", filterReviewRows);
reviewFile.addEventListener("change", updateReviewProgress);
sourceReviewFile.addEventListener("change", updateReviewProgress);

newJobButton.addEventListener("click", () => {
  stopPolling();
  forgetJob();
  prepareForm.reset();
  reviewForm.reset();
  reviewPanel.hidden = true;
  preparePanel.hidden = false;
  exceptionsLink.classList.remove("visible");
  downloadLink.classList.remove("visible");
  newJobButton.hidden = true;
  setFormDisabled(prepareForm, false);
  syncOutputSelection();
  setStatus("Ready.");
});

includeMobile.addEventListener("change", syncOutputSelection);
includeM2m.addEventListener("change", syncOutputSelection);
syncOutputSelection();

const now = new Date();
document.getElementById("report_date").value = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
document.getElementById("report_month").value = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
fetch("/health", { cache: "no-store" })
  .then((response) => { health.textContent = response.ok ? "API status: ready" : "API status: unavailable"; })
  .catch(() => { health.textContent = "API status: unavailable"; });
try { activeJobId = normalizeJobId(window.localStorage.getItem(storageKey)); } catch (_error) { /* storage optional */ }
if (activeJobId) pollJob();
