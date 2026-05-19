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

    const response = await fetch("/convert/report-3/upload", {
      method: "POST",
      body,
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.detail || "Conversion failed.");
    }

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
