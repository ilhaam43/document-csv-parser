const form = document.getElementById("upload-form");
const button = document.getElementById("submit-button");
const statusBox = document.getElementById("status-box");
const downloadLink = document.getElementById("download-link");
const health = document.getElementById("health");

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
  if (seconds < 20) {
    return "Preparing workbook";
  }
  if (seconds < 60) {
    return "Cleaning and mapping data";
  }
  if (seconds < 180) {
    return "Refreshing Excel pivots";
  }

  return "Saving output workbook";
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
      throw new TransientGatewayError(`Public tunnel returned HTTP ${response.status}. Retrying...`, response.status);
    }

    throw new Error("The public tunnel returned an HTML error page instead of API JSON. The request may have timed out while Excel was still processing. Please wait a moment, then try again or check the generated output.");
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
    throw new Error(payload.detail || payload.error || `Request failed: HTTP ${response.status}`);
  }
  return payload;
}

async function waitForJob(statusUrl) {
  const startedAt = Date.now();
  let transientFailures = 0;

  while (true) {
    let payload;
    try {
      payload = await fetchJson(statusUrl, { cache: "no-store" });
      transientFailures = 0;
    } catch (error) {
      const retryWindowMs = 60 * 60 * 1000;
      const canRetry =
        error instanceof TransientGatewayError &&
        Date.now() - startedAt < retryWindowMs &&
        transientFailures < 240;

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
      throw new Error(payload.error || "Conversion failed.");
    }

    await sleep(3000);
  }
}

const jobMonitor = window.createPersistentJobMonitor({
  storageKey: "document-csv-parser:report-1-active-job",
  statusPrefix: "/jobs/report-1/",
  button,
  statusBox,
  downloadLink,
  processingStage,
  formatDuration,
  waitForJob,
  successMessage: (payload) => `Generated ${payload.filename} in ${formatDuration(payload.elapsed_seconds)}.`,
  fallbackError: "Conversion failed.",
});

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
  const started = jobMonitor.startSubmission();

  try {
    const body = new FormData(form);
    body.set("refresh_template", document.getElementById("refresh_template").checked ? "true" : "false");

    const queued = await fetchJson("/convert/upload/jobs", {
      method: "POST",
      body,
    });
    await jobMonitor.trackQueuedJob(queued, started);
  } catch (error) {
    jobMonitor.fail(error);
  }
});
