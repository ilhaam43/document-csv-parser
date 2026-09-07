(function () {
  "use strict";

  window.createPersistentJobMonitor = function createPersistentJobMonitor(config) {
    const {
      storageKey,
      statusPrefix,
      button,
      statusBox,
      downloadLink,
      processingStage,
      formatDuration,
      waitForJob,
      successMessage,
      fallbackError,
    } = config;
    let progressTimer = null;
    let pageIsUnloading = false;

    window.addEventListener("pagehide", () => {
      pageIsUnloading = true;
      stopProgress();
    });

    function clearStoredJob() {
      try {
        window.localStorage.removeItem(storageKey);
      } catch (_error) {
        // Current-page polling still works when browser storage is unavailable.
      }
    }

    function loadStoredJob() {
      try {
        const value = window.localStorage.getItem(storageKey);
        if (!value) {
          return null;
        }

        const job = JSON.parse(value);
        if (
          typeof job.statusUrl !== "string" ||
          !job.statusUrl.startsWith(statusPrefix) ||
          !Number.isFinite(job.startedAt)
        ) {
          clearStoredJob();
          return null;
        }
        return job;
      } catch (_error) {
        clearStoredJob();
        return null;
      }
    }

    function saveStoredJob(job) {
      try {
        window.localStorage.setItem(storageKey, JSON.stringify(job));
      } catch (_error) {
        // Current-page polling still works when browser storage is unavailable.
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

    function stopProgress() {
      if (progressTimer !== null) {
        window.clearInterval(progressTimer);
        progressTimer = null;
      }
    }

    function startProgress(startedAt) {
      stopProgress();
      const update = () => {
        const seconds = Math.round((Date.now() - startedAt) / 1000);
        statusBox.textContent = `${processingStage(seconds)}... ${formatDuration(seconds)}`;
      };
      update();
      progressTimer = window.setInterval(update, 1000);
    }

    function hideDownload() {
      downloadLink.classList.remove("visible");
      downloadLink.removeAttribute("href");
    }

    function fail(error) {
      if (pageIsUnloading) {
        return;
      }
      clearStoredJob();
      stopProgress();
      statusBox.className = "status-box bad";
      statusBox.textContent = error.message || fallbackError;
      button.disabled = false;
    }

    async function monitor(job) {
      button.disabled = true;
      hideDownload();
      statusBox.className = "status-box";
      startProgress(job.startedAt);

      try {
        const payload = await waitForJob(job.statusUrl);
        stopProgress();
        statusBox.className = "status-box ok";
        statusBox.textContent = successMessage(payload);
        downloadLink.href = payload.download_url;
        downloadLink.classList.add("visible");
      } catch (error) {
        if (!pageIsUnloading) {
          fail(error);
        }
      } finally {
        stopProgress();
        if (!pageIsUnloading) {
          button.disabled = false;
        }
      }
    }

    function startSubmission(pendingJob = null) {
      clearStoredJob();
      hideDownload();
      statusBox.className = "status-box";
      button.disabled = true;
      const startedAt = Date.now();
      if (pendingJob) {
        saveStoredJob({
          statusUrl: pendingJob.statusUrl,
          startedAt,
          filename: pendingJob.filename || "",
        });
      }
      startProgress(startedAt);
      return startedAt;
    }

    async function trackQueuedJob(queued, startedAt) {
      const job = {
        statusUrl: queued.status_url,
        startedAt,
        filename: queued.filename,
      };
      saveStoredJob(job);
      await monitor(job);
    }

    downloadLink.addEventListener("click", clearStoredJob);

    const storedJob = loadStoredJob();
    if (storedJob) {
      void monitor(storedJob);
    }

    return {
      createJobId,
      fail,
      startSubmission,
      trackQueuedJob,
    };
  };
})();
