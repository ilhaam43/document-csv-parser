const form = document.getElementById("email-form");
const button = document.getElementById("submit-button");
const statusBox = document.getElementById("status-box");
const imageLink = document.getElementById("image-link");
const health = document.getElementById("health");
async function readJson(response) { const payload = await response.json().catch(() => ({})); if (!response.ok) throw new Error(payload.detail || `Request failed: HTTP ${response.status}`); return payload; }
fetch("/health").then((response) => response.ok ? response.json() : Promise.reject()).then(() => { health.textContent = "API status: ready"; }).catch(() => { health.textContent = "API status: unavailable"; });
form.addEventListener("submit", async (event) => { event.preventDefault(); button.disabled = true; imageLink.classList.remove("visible"); statusBox.className = "status-box"; statusBox.textContent = "Uploading Report 3 workbook and creating screenshot..."; try { const body = new FormData(form); body.set("dry_run", document.getElementById("dry_run").checked ? "true" : "false"); const result = await readJson(await fetch("/email-report-3/send", { method: "POST", body })); statusBox.className = "status-box ok"; statusBox.textContent = result.message; if (Array.isArray(result.public_images) && result.public_images[0]) { imageLink.href = result.public_images[0]; imageLink.classList.add("visible"); } } catch (error) { statusBox.className = "status-box bad"; statusBox.textContent = error.message || "Report 3 email workflow failed."; } finally { button.disabled = false; } });
