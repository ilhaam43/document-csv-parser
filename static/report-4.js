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
    return "Uploading IDE files";
  }
  if (seconds < 20) {
    return "Reading IDE dashboard";
  }
  if (seconds < 40) {
    return "Mapping data from the previous workbook";
  }
  if (seconds < 60) {
    return "Applying collabs OTC/MRC fallback";
  }
  if (seconds < 90) {
    return "Updating ALL ORDER";
  }
  if (seconds < 180) {
    return "Refreshing IDE pivots";
  }

  return "Validating and saving workbook";
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

  if (text.trim().startsWith("<") && isTransientGatewayStatus(response.status)) {
    throw new TransientGatewayError(`Public gateway returned HTTP ${response.status}. Retrying...`, response.status);
  }

  if (text.trim().startsWith("<")) {
    throw new Error("The public gateway returned an HTML page instead of API JSON.");
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
  storageKey: "document-csv-parser:report-4-active-job",
  statusPrefix: "/jobs/report-4/",
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
    const queued = await fetchJson("/convert/report-4/upload/jobs", {
      method: "POST",
      body,
    });
    await jobMonitor.trackQueuedJob(queued, started);
  } catch (error) {
    jobMonitor.fail(error);
  }
});
