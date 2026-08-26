const form = document.getElementById("email-form");
const button = document.getElementById("submit-button");
const statusBox = document.getElementById("status-box");
const onedriveLink = document.getElementById("onedrive-link");
const health = document.getElementById("health");

async function readJson(response) {
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.detail || `Request failed: HTTP ${response.status}`);
  }
  return payload;
}

fetch("/health")
  .then((response) => (response.ok ? response.json() : Promise.reject()))
  .then(() => { health.textContent = "API status: ready"; })
  .catch(() => { health.textContent = "API status: unavailable"; });

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  button.disabled = true;
  onedriveLink.classList.remove("visible");
  statusBox.className = "status-box";
    statusBox.textContent = "Uploading Report 1 workbook and creating screenshots...";

  try {
    const body = new FormData(form);
    body.set("dry_run", document.getElementById("dry_run").checked ? "true" : "false");
    const result = await readJson(await fetch("/email-report-1/send", { method: "POST", body }));

    statusBox.className = "status-box ok";
    statusBox.textContent = result.message;
    if (Array.isArray(result.onedrive) && result.onedrive[0] && result.onedrive[0].web_url) {
      onedriveLink.href = result.onedrive[0].web_url;
      onedriveLink.classList.add("visible");
    }
  } catch (error) {
    statusBox.className = "status-box bad";
    statusBox.textContent = error.message || "Email workflow failed.";
  } finally {
    button.disabled = false;
  }
});
