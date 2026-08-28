const form = document.getElementById("upload-form");
const button = document.getElementById("submit-button");
const statusBox = document.getElementById("status-box");
const downloadLink = document.getElementById("download-link");
const health = document.getElementById("health");
const activeJobStorageKey = "document-csv-parser:pipeline-active-job";
let progressTimer = null;
let pageIsUnloading = false;

window.addEventListener("pagehide", () => {
  pageIsUnloading = true;
  stopProgressTimer();
});

function loadActiveJob() {
  try {
    const value = window.localStorage.getItem(activeJobStorageKey);
    if (!value) {
      return null;
    }

    const job = JSON.parse(value);
    if (
      typeof job.statusUrl !== "string" ||
      !job.statusUrl.startsWith("/jobs/pipeline/") ||
      !Number.isFinite(job.startedAt)
    ) {
      window.localStorage.removeItem(activeJobStorageKey);
      return null;
    }
    return job;
  } catch (_error) {
    return null;
  }
}

function saveActiveJob(job) {
  try {
    window.localStorage.setItem(activeJobStorageKey, JSON.stringify(job));
  } catch (_error) {
    // Polling still works in the current page when storage is unavailable.
  }
}

function clearActiveJob() {
  try {
    window.localStorage.removeItem(activeJobStorageKey);
  } catch (_error) {
    // Nothing else is required when storage is unavailable.
  }
}

function createJobId() {
  if (window.crypto && typeof window.crypto.randomUUID === "function") {
    return window.crypto.randomUUID().replaceAll("-", "");
  }

  const bytes = new Uint8Array(16);
  window.crypto.getRandomValues(bytes);
  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;
  return Array.from(bytes, (value) => value.toString(16).padStart(2, "0")).join("");
}

function formatDuration(totalSeconds) {
  const seconds = Math.max(0, Math.round(Number(totalSeconds) || 0));
  const minutes = Math.floor(seconds / 60);
  const remainingSeconds = seconds % 60;

  if (minutes === 0) {
    return `${remainingSeconds} detik`;
  }

  return `${minutes} menit ${remainingSeconds} detik`;
}

function processingStage(seconds) {
  if (seconds < 5) {
    return "Uploading files";
  }
  if (seconds < 30) {
    return "Generating Daily Tracking";
  }
  if (seconds < 180) {
    return "Generating Ongoing Tracking";
  }
  if (seconds < 300) {
    return "Generating iPhone Tracking";
  }

  return "Packaging output ZIP";
}

function startProgressTimer(startedAt) {
  if (progressTimer !== null) {
    window.clearInterval(progressTimer);
  }

  const updateProgress = () => {
    const seconds = Math.round((Date.now() - startedAt) / 1000);
    statusBox.textContent = `${processingStage(seconds)}... ${formatDuration(seconds)}`;
  };
  updateProgress();
  progressTimer = window.setInterval(updateProgress, 1000);
}

function stopProgressTimer() {
  if (progressTimer !== null) {
    window.clearInterval(progressTimer);
    progressTimer = null;
  }
}

class TransientGatewayError extends Error {
  constructor(message, status) {
    super(message);
    this.name = "TransientGatewayError";
    this.status = status;
  }
}

function isTransientGatewayStatus(status) {
  return [408, 425, 429, 500, 502, 503, 504, 520, 521, 522, 523, 524].includes(status);
}

async function readJsonResponse(response) {
  const contentType = response.headers.get("content-type") || "";
  const text = await response.text();

  if (contentType.includes("application/json")) {
    return text ? JSON.parse(text) : {};
  }

  if (text.trim().startsWith("<")) {
    if (isTransientGatewayStatus(response.status)) {
      throw new TransientGatewayError(`Public gateway returned HTTP ${response.status}. Retrying...`, response.status);
    }

    throw new Error("The public gateway returned an HTML page instead of API JSON. Please wait a moment, then try again or check the generated output.");
  }

  throw new Error(text || `Unexpected server response: HTTP ${response.status}`);
}

function sleep(milliseconds) {
  return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      Accept: "application/json",
      ...(options.headers || {}),
    },
  });
  const payload = await readJsonResponse(response);
  if (!response.ok) {
    const error = new Error(payload.detail || payload.error || `Request failed: HTTP ${response.status}`);
    error.status = response.status;
    throw error;
  }
  return payload;
}

async function waitForJob(statusUrl) {
  const startedAt = Date.now();
  let transientFailures = 0;
  const registrationGracePeriodMs = 10 * 60 * 1000;

  while (true) {
    let payload;
    try {
      payload = await fetchJson(statusUrl, { cache: "no-store" });
      transientFailures = 0;
    } catch (error) {
      const retryWindowMs = 60 * 60 * 1000;
      const waitingForRegistration =
        error.status === 404 && Date.now() - startedAt < registrationGracePeriodMs;
      const canRetry =
        waitingForRegistration ||
        (error instanceof TransientGatewayError &&
          Date.now() - startedAt < retryWindowMs &&
          transientFailures < 240);

      if (!canRetry) {
        throw error;
      }

      transientFailures += 1;
      await sleep(Math.min(15000, 3000 + transientFailures * 1000));
      continue;
    }

    if (payload.status === "succeeded") {
      return payload;
    }
    if (payload.status === "failed") {
      throw new Error(payload.error || "Pipeline failed.");
    }

    await sleep(3000);
  }
}

async function monitorActiveJob(job) {
  button.disabled = true;
  downloadLink.classList.remove("visible");
  downloadLink.removeAttribute("href");
  statusBox.className = "status-box";
  startProgressTimer(job.startedAt);

  try {
    const payload = await waitForJob(job.statusUrl);
    stopProgressTimer();
    statusBox.className = "status-box ok";
    statusBox.textContent = `Generated ${payload.output_files.join(", ")} in ${formatDuration(payload.elapsed_seconds)}.`;
    downloadLink.href = payload.download_url;
    downloadLink.classList.add("visible");
  } catch (error) {
    if (!pageIsUnloading) {
      clearActiveJob();
      stopProgressTimer();
      statusBox.className = "status-box bad";
      statusBox.textContent = error.message || "Pipeline failed.";
    }
  } finally {
    if (!pageIsUnloading) {
      button.disabled = false;
    }
  }
}

fetch("/health")
  .then((response) => (response.ok ? response.json() : Promise.reject()))
  .then(() => {
    health.textContent = "API status: ready";
  })
  .catch(() => {
    health.textContent = "API status: unavailable";
  });

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  downloadLink.classList.remove("visible");
  downloadLink.removeAttribute("href");
  statusBox.className = "status-box";
  button.disabled = true;

  const started = Date.now();
  const jobId = createJobId();
  const activeJob = {
    statusUrl: `/jobs/pipeline/${jobId}`,
    startedAt: started,
    filename: "daily-pipeline-output.zip",
  };
  saveActiveJob(activeJob);
  startProgressTimer(started);

  try {
    const body = new FormData(form);
    body.set("job_id", jobId);
    body.set("refresh_template", document.getElementById("refresh_template").checked ? "true" : "false");
    body.set("ongoing_with_pivot", document.getElementById("ongoing_with_pivot").checked ? "true" : "false");

    const queued = await fetchJson("/convert/pipeline/upload/jobs", {
      method: "POST",
      body,
    });
    activeJob.statusUrl = queued.status_url;
    activeJob.filename = queued.filename;
    saveActiveJob(activeJob);
    await monitorActiveJob(activeJob);
  } catch (error) {
    if (!pageIsUnloading) {
      clearActiveJob();
      stopProgressTimer();
      statusBox.className = "status-box bad";
      statusBox.textContent = error.message || "Pipeline failed.";
    }
  } finally {
    stopProgressTimer();
    if (!pageIsUnloading) {
      button.disabled = false;
    }
  }
});

downloadLink.addEventListener("click", clearActiveJob);

const activeJob = loadActiveJob();
if (activeJob) {
  monitorActiveJob(activeJob);
}
