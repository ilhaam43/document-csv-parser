const form = document.getElementById("upload-form");
const button = document.getElementById("submit-button");
const statusBox = document.getElementById("status-box");
const downloadLink = document.getElementById("download-link");
const health = document.getElementById("health");

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
    statusBox.textContent = `Processing ongoing workbook... ${seconds}s`;
  }, 1000);

  try {
    const body = new FormData(form);
    body.set("with_pivot", document.getElementById("with_pivot").checked ? "true" : "false");

    const response = await fetch("/convert/report-2/upload", {
      method: "POST",
      body,
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.detail || "Conversion failed.");
    }

    statusBox.className = "status-box ok";
    statusBox.textContent = `Generated ${payload.filename} in ${payload.elapsed_seconds}s.`;
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
