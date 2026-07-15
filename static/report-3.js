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
    return "Preparing iPhone workbook";
  }
  if (seconds < 60) {
    return "Mapping iPhone order data";
  }
  if (seconds < 180) {
    return "Refreshing iPhone pivots";
  }

  return "Saving output workbook";
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
  const timer = window.setInterval(() => {
    const seconds = Math.round((Date.now() - started) / 1000);
    statusBox.textContent = `${processingStage(seconds)}... ${formatDuration(seconds)}`;
  }, 1000);

  try {
    const body = new FormData(form);

    const queued = await fetchJson("/convert/report-3/upload/jobs", {
      method: "POST",
      body,
    });
    statusBox.textContent = `Job queued. Processing ${queued.filename}...`;
    const payload = await waitForJob(queued.status_url);

    statusBox.className = "status-box ok";
    statusBox.textContent = `Generated ${payload.filename} in ${formatDuration(payload.elapsed_seconds)}.`;
    downloadLink.href = payload.download_url;
    downloadLink.classList.add("visible");
  } catch (error) {
    statusBox.className = "status-box bad";
    statusBox.textContent = error.message || "Conversion failed.";
  } finally {
    window.clearInterval(timer);
    button.disabled = false;
  }
});
