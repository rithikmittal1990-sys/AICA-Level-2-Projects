(() => {
  const STAGES = [
    "uploaded",
    "extracting",
    "classifying",
    "mapping",
    "review",
    "generating",
    "validating",
  ];

  const fileInput = document.getElementById("file-input");
  const dropzone = document.getElementById("dropzone");
  const filenameRow = document.getElementById("filename");
  const filenameValue = document.getElementById("filename-value");
  const clearFile = document.getElementById("clear-file");
  const generateButton = document.getElementById("generate");
  const progressPanel = document.getElementById("progress-panel");
  const stepItems = [...document.querySelectorAll("#steps li")];
  const reviewPanel = document.getElementById("review-panel");
  const reviewBody = document.getElementById("review-body");
  const reviewMeta = document.getElementById("review-meta");
  const yearCurrent = document.getElementById("year-current");
  const yearPrevious = document.getElementById("year-previous");
  const approveAll = document.getElementById("approve-all");
  const approveEligible = document.getElementById("approve-eligible");
  const approveGenerate = document.getElementById("approve-generate");
  const resultPanel = document.getElementById("result-panel");
  const validationBadge = document.getElementById("validation-badge");
  const downloadLink = document.getElementById("download");
  const issuesPanel = document.getElementById("issues-panel");
  const issuesList = document.getElementById("issues-list");

  let selectedFile = null;
  let pollTimer = null;
  let activeJobId = null;
  let reviewState = null;

  fileInput.addEventListener("change", () => {
    setFile(fileInput.files[0] || null);
  });

  clearFile.addEventListener("click", () => {
    fileInput.value = "";
    setFile(null);
  });

  ["dragenter", "dragover"].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.add("is-dragover");
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.remove("is-dragover");
    });
  });

  dropzone.addEventListener("drop", (event) => {
    const file = event.dataTransfer?.files?.[0];
    if (file) {
      setFile(file);
    }
  });

  generateButton.addEventListener("click", () => {
    if (selectedFile) {
      void generate(selectedFile);
    }
  });

  yearCurrent.addEventListener("change", () => {
    void saveYears();
  });
  yearPrevious.addEventListener("change", () => {
    void saveYears();
  });
  if (approveAll) {
    approveAll.addEventListener("click", () => {
      void approveAllRows();
    });
  }
  approveEligible.addEventListener("click", () => {
    void approveEligibleRows();
  });
  approveGenerate.addEventListener("click", () => {
    void submitApproval();
  });

  function setFile(file) {
    if (file && hasWrongExcelExtension(file)) {
      selectedFile = null;
      fileInput.value = "";
      showIssues({
        errors: [
          {
            message:
              "This file ends with .xlxs — the correct extension is .xlsx. Rename the file and try again.",
          },
        ],
        warnings: [],
      });
      filenameRow.hidden = true;
      generateButton.disabled = true;
      return;
    }

    selectedFile = file && isSupportedUpload(file) ? file : file ? null : null;
    if (file && !isSupportedUpload(file)) {
      selectedFile = null;
      fileInput.value = "";
      showIssues({
        errors: [{ message: "Please upload a trial balance (.xlsx) or PDF file." }],
        warnings: [],
      });
    } else {
      hideIssues();
    }

    if (selectedFile) {
      filenameValue.textContent = selectedFile.name;
      filenameRow.hidden = false;
      generateButton.disabled = false;
    } else {
      filenameRow.hidden = true;
      generateButton.disabled = true;
    }
  }

  function hasWrongExcelExtension(file) {
    const name = (file.name || "").toLowerCase();
    return name.endsWith(".xlxs") || name.endsWith(".xslx");
  }

  function isSupportedUpload(file) {
    return isPdf(file) || isExcel(file);
  }

  function isExcel(file) {
    const name = (file.name || "").toLowerCase();
    const type = (file.type || "").toLowerCase();
    return (
      name.endsWith(".xlsx") ||
      name.endsWith(".xlsm") ||
      type.includes("spreadsheetml") ||
      type.includes("excel") ||
      type === "application/vnd.ms-excel"
    );
  }

  function isPdf(file) {
    const name = (file.name || "").toLowerCase();
    return file.type === "application/pdf" || name.endsWith(".pdf");
  }

  async function generate(file) {
    stopPolling();
    activeJobId = null;
    reviewState = null;
    generateButton.disabled = true;
    generateButton.textContent = "Generating…";
    progressPanel.hidden = false;
    reviewPanel.hidden = true;
    resultPanel.hidden = true;
    downloadLink.hidden = true;
    hideIssues();
    setProgress("uploaded", "processing");

    const body = new FormData();
    body.append("file", file, file.name);

    try {
      const uploaded = await readJson(await fetch("/upload", { method: "POST", body }));
      const jobId = uploaded.job_id;
      if (!jobId) {
        throw new Error("The server did not return a job id.");
      }
      activeJobId = jobId;
      if (uploaded.status === "failed") {
        finishFailed(uploaded);
        return;
      }
      await pollJob(jobId);
    } catch (error) {
      finishFailed({
        status: "failed",
        errors: [{ message: publicMessage(error) }],
        warnings: [],
      });
    }
  }

  async function pollJob(jobId) {
    return new Promise((resolve, reject) => {
      const tick = async () => {
        try {
          const status = await readJson(await fetch(`/status/${encodeURIComponent(jobId)}`));
          setProgress(status.stage, status.status);

          if (status.status === "review_required") {
            await loadReview(jobId);
            resolve();
            return;
          }
          if (status.status === "completed") {
            await finishCompleted(jobId, status);
            resolve();
            return;
          }
          if (status.status === "failed") {
            finishFailed(status);
            resolve();
            return;
          }
          pollTimer = window.setTimeout(tick, 400);
        } catch (error) {
          reject(error);
        }
      };
      void tick();
    });
  }

  async function loadReview(jobId) {
    setProgress("review", "review_required");
    const review = await readJson(
      await fetch(`/review/${encodeURIComponent(jobId)}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: "{}",
      })
    );
    reviewState = review;
    renderReview(review);
    generateButton.disabled = !selectedFile;
    generateButton.textContent = "Generate Excel";
  }

  function renderReview(review) {
    reviewPanel.hidden = false;
    yearCurrent.value = (review.financial_year && review.financial_year.current) || "";
    yearPrevious.value = (review.financial_year && review.financial_year.previous) || "";
    const summary = review.summary || {};
    reviewMeta.textContent = `${summary.total || 0} mappings · ${summary.needs_review || 0} need review · threshold ${(Number(review.threshold || 0) * 100).toFixed(0)}%`;
    reviewBody.innerHTML = "";
    (review.items || []).forEach((item) => {
      reviewBody.append(reviewRow(item));
    });
  }

  function reviewRow(item) {
    const row = document.createElement("tr");
    row.dataset.itemId = item.item_id;
    row.className = rowClass(item.status);
    row.append(
      cell(item.source_field || "—"),
      valueCell(item),
      cell(item.source_page == null ? "—" : String(item.source_page)),
      cell(item.schedule_iii_category || "—"),
      cell(item.excel_destination || "—"),
      cell(formatConfidence(item.confidence)),
      statusCell(item.status),
      periodCell(item),
      noteCell(item),
      actionsCell(item)
    );
    return row;
  }

  function cell(text) {
    const td = document.createElement("td");
    td.textContent = text;
    return td;
  }

  function valueCell(item) {
    const td = document.createElement("td");
    const input = document.createElement("input");
    input.type = "text";
    input.value = item.extracted_value == null ? "" : String(item.extracted_value);
    input.addEventListener("change", () => {
      void patchItem(item.item_id, { extracted_value: input.value === "" ? null : input.value });
    });
    td.append(input);
    return td;
  }

  function statusCell(status) {
    const td = document.createElement("td");
    const pill = document.createElement("span");
    pill.className = "status-pill";
    pill.dataset.status = status || "pending";
    pill.textContent = formatStatus(status);
    td.append(pill);
    return td;
  }

  function periodCell(item) {
    const td = document.createElement("td");
    const select = document.createElement("select");
    ["current", "previous", "note"].forEach((period) => {
      const option = document.createElement("option");
      option.value = period;
      option.textContent = period === "note" ? "Note" : period === "previous" ? "Previous year" : "Current year";
      if (item.period === period) {
        option.selected = true;
      }
      select.append(option);
    });
    select.addEventListener("change", () => {
      void patchItem(item.item_id, { period: select.value });
    });
    td.append(select);
    return td;
  }

  function noteCell(item) {
    const td = document.createElement("td");
    const input = document.createElement("input");
    input.type = "text";
    input.value = item.note_number == null ? "" : String(item.note_number);
    input.addEventListener("change", () => {
      void patchItem(item.item_id, { note_number: input.value === "" ? null : input.value });
    });
    td.append(input);
    return td;
  }

  function actionsCell(item) {
    const td = document.createElement("td");
    const wrap = document.createElement("div");
    wrap.className = "row-actions";
    const approve = document.createElement("button");
    approve.type = "button";
    approve.className = "is-approve";
    approve.textContent = "Approve";
    approve.addEventListener("click", () => {
      void patchItem(item.item_id, { status: "approved" });
    });
    const reject = document.createElement("button");
    reject.type = "button";
    reject.className = "is-reject";
    reject.textContent = "Reject";
    reject.addEventListener("click", () => {
      void patchItem(item.item_id, { status: "rejected" });
    });
    wrap.append(approve, reject);
    td.append(wrap);
    return td;
  }

  async function patchItem(itemId, changes) {
    if (!activeJobId) {
      return;
    }
    const review = await readJson(
      await fetch(`/review/${encodeURIComponent(activeJobId)}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ items: [{ item_id: itemId, ...changes }] }),
      })
    );
    reviewState = review;
    renderReview(review);
  }

  async function saveYears() {
    if (!activeJobId) {
      return;
    }
    const review = await readJson(
      await fetch(`/review/${encodeURIComponent(activeJobId)}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          financial_year: {
            current: yearCurrent.value,
            previous: yearPrevious.value,
          },
        }),
      })
    );
    reviewState = review;
    renderReview(review);
  }

  async function approveEligibleRows() {
    if (!reviewState) {
      return;
    }
    const items = (reviewState.items || [])
      .filter((item) => item.status === "pending")
      .map((item) => ({ item_id: item.item_id, status: "approved" }));
    if (!items.length || !activeJobId) {
      return;
    }
    const review = await readJson(
      await fetch(`/review/${encodeURIComponent(activeJobId)}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ items }),
      })
    );
    reviewState = review;
    renderReview(review);
  }

  async function approveAllRows() {
    if (!reviewState || !activeJobId) {
      return;
    }
    const items = (reviewState.items || [])
      .filter((item) => item.status !== "approved")
      .map((item) => ({ item_id: item.item_id, status: "approved" }));
    if (!items.length) {
      return;
    }
    if (approveAll) {
      approveAll.disabled = true;
      approveAll.textContent = "Approving…";
    }
    try {
      const review = await readJson(
        await fetch(`/review/${encodeURIComponent(activeJobId)}`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ items }),
        })
      );
      reviewState = review;
      renderReview(review);
    } finally {
      if (approveAll) {
        approveAll.disabled = false;
        approveAll.textContent = "✓ Approve all";
      }
    }
  }

  async function submitApproval() {
    if (!activeJobId) {
      return;
    }
    approveGenerate.disabled = true;
    setProgress("generating", "processing");
    try {
      const generated = await readJson(
        await fetch(`/approve/${encodeURIComponent(activeJobId)}`, { method: "POST" })
      );
      if (generated.status === "failed") {
        finishFailed(generated);
        return;
      }
      if (generated.status === "review_required") {
        await loadReview(activeJobId);
        showIssues({
          errors: [{ message: "Mappings below the confidence threshold still need a decision." }],
          warnings: [],
        });
        return;
      }
      await finishCompleted(activeJobId, generated);
    } catch (error) {
      setProgress("review", "review_required");
      showIssues({ errors: [{ message: publicMessage(error) }], warnings: [] });
    } finally {
      approveGenerate.disabled = false;
    }
  }

  async function finishCompleted(jobId, status) {
    setProgress("completed", "completed");
    reviewPanel.hidden = true;
    let validationStatus = status.validation_status || "WARNING";
    let report = { errors: status.errors || [], warnings: status.warnings || [] };

    try {
      const validation = await readJson(await fetch(`/validation/${encodeURIComponent(jobId)}`));
      validationStatus = validation.validation_status || validation.validation?.status || validationStatus;
      report = {
        errors: [
          ...(status.errors || []),
          ...((validation.validation && validation.validation.errors) || []),
        ],
        warnings: [
          ...(status.warnings || []),
          ...((validation.validation && validation.validation.warnings) || []),
        ],
      };
    } catch {
      // Status payload is enough to show a summary if validation cannot be loaded.
    }

    showValidation(validationStatus);
    downloadLink.href = `/download/${encodeURIComponent(jobId)}`;
    downloadLink.hidden = false;
    generateButton.disabled = !selectedFile;
    generateButton.textContent = "Generate Excel";
    showIssues(report);
  }

  function finishFailed(payload) {
    setProgress(payload.stage || "failed", "failed");
    reviewPanel.hidden = true;
    showValidation("FAILED");
    downloadLink.hidden = true;
    generateButton.disabled = !selectedFile;
    generateButton.textContent = "Generate Excel";
    showIssues({
      errors: payload.errors || [{ message: "Document processing failed." }],
      warnings: payload.warnings || [],
    });
  }

  function setProgress(stage, status) {
    const activeIndex =
      stage === "completed" || status === "completed"
        ? STAGES.length
        : Math.max(0, STAGES.indexOf(stage));
    const failed = status === "failed";

    stepItems.forEach((item, index) => {
      item.classList.remove("is-done", "is-active", "is-error");
      if (failed && index === Math.min(activeIndex, STAGES.length - 1) && stage !== "completed") {
        item.classList.add("is-error");
        return;
      }
      if (index < activeIndex) {
        item.classList.add("is-done");
      } else if (index === activeIndex && !failed) {
        item.classList.add("is-active");
      }
    });
  }

  function showValidation(status) {
    resultPanel.hidden = false;
    const label = status === "FAILED" ? "FAILED" : status;
    validationBadge.textContent = label;
    validationBadge.dataset.status = label;
  }

  function showIssues(report) {
    const errors = uniqueIssues(report.errors);
    const warnings = uniqueIssues(report.warnings);
    if (!errors.length && !warnings.length) {
      hideIssues();
      return;
    }

    issuesPanel.hidden = false;
    issuesList.innerHTML = "";
    if (errors.length) {
      issuesList.append(issueGroup("Errors", errors, "error"));
    }
    if (warnings.length) {
      issuesList.append(issueGroup("Warnings", warnings, "warning"));
    }
  }

  function hideIssues() {
    issuesPanel.hidden = true;
    issuesList.innerHTML = "";
  }

  function issueGroup(title, items, kind) {
    const section = document.createElement("div");
    section.className = `issue-group ${kind}`;
    const heading = document.createElement("h3");
    heading.textContent = title;
    const list = document.createElement("ul");
    items.forEach((item) => {
      const row = document.createElement("li");
      row.textContent = issueText(item);
      list.append(row);
    });
    section.append(heading, list);
    return section;
  }

  function uniqueIssues(items) {
    const seen = new Set();
    const result = [];
    (items || []).forEach((item) => {
      const text = issueText(item);
      if (text && !seen.has(text)) {
        seen.add(text);
        result.push(item);
      }
    });
    return result;
  }

  function issueText(item) {
    if (!item) {
      return "";
    }
    if (typeof item === "string") {
      return item;
    }
    return String(item.message || item.detail || "").trim();
  }

  async function readJson(response) {
    let payload = {};
    try {
      payload = await response.json();
    } catch {
      payload = {};
    }
    delete payload.traceback;
    delete payload.exception;
    delete payload.stack;
    if (!response.ok) {
      throw new Error(payload.message || "The request could not be completed.");
    }
    return payload;
  }

  function publicMessage(error) {
    const text = error && error.message ? String(error.message) : "";
    if (!text || /traceback|exception at /i.test(text)) {
      return "Document processing failed.";
    }
    return text;
  }

  function formatConfidence(value) {
    if (value == null || Number.isNaN(Number(value))) {
      return "—";
    }
    return `${Math.round(Number(value) * 100)}%`;
  }

  function formatStatus(status) {
    if (status === "needs_review") {
      return "Needs review";
    }
    return status || "pending";
  }

  function rowClass(status) {
    if (status === "needs_review") {
      return "needs-review";
    }
    if (status === "approved") {
      return "is-approved";
    }
    if (status === "rejected") {
      return "is-rejected";
    }
    return "";
  }

  function stopPolling() {
    if (pollTimer) {
      window.clearTimeout(pollTimer);
      pollTimer = null;
    }
  }
})();
