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
    body.set("refresh_template", document.getElementById("refresh_template").checked ? "true" : "false");
    body.set("ongoing_with_pivot", document.getElementById("ongoing_with_pivot").checked ? "true" : "false");

    const response = await fetch("/convert/pipeline/upload", {
      method: "POST",
      body,
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.detail || "Pipeline failed.");
    }

    statusBox.className = "status-box ok";
    statusBox.textContent = `Generated ${payload.output_files.join(", ")} in ${formatDuration(payload.elapsed_seconds)}.`;
    downloadLink.href = payload.download_url;
    downloadLink.classList.add("visible");
  } catch (error) {
    statusBox.className = "status-box bad";
    statusBox.textContent = error.message || "Pipeline failed.";
  } finally {
    window.clearInterval(timer);
    button.disabled = false;
  }
});
