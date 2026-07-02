(() => {
  const $ = (id) => document.getElementById(id);
  const state = { files: { source: null, template: null }, profile: "generic" };

  function useDefaultTemplate() {
    const toggle = $("use-default-template");
    return state.profile === "esf" && toggle && toggle.checked;
  }

  function updateTemplateUi() {
    const actions = $("template-actions");
    const area = document.querySelector('.drop-area[data-target="template"]');
    const placeholder = area.querySelector(".drop-placeholder");
    const info = $("template-info");
    const defaultMode = useDefaultTemplate();

    actions.classList.toggle("hidden", state.profile !== "esf");
    area.classList.toggle("disabled", defaultMode);
    if (defaultMode) {
      area.classList.add("has-file");
      placeholder.classList.add("hidden");
      info.classList.remove("hidden");
      info.textContent = "✓ 内置 FSEC ESF 2026 v2.2.4 空白模板";
    } else if (state.files.template) {
      area.classList.add("has-file");
      placeholder.classList.add("hidden");
      info.classList.remove("hidden");
      info.textContent = `✓ ${state.files.template.name} (${(state.files.template.size / 1048576).toFixed(1)} MB)`;
    } else {
      area.classList.remove("has-file");
      placeholder.classList.remove("hidden");
      info.classList.add("hidden");
      info.textContent = "";
    }
  }

  // Mode selection
  document.querySelectorAll('input[name="profile"]').forEach((inp) => {
    inp.addEventListener("change", () => {
      state.profile = inp.value;
      if (state.profile === "esf") {
        const defaultTemplate = $("use-default-template");
        if (defaultTemplate) defaultTemplate.checked = true;
        const noImages = $("no-images");
        if (noImages) noImages.checked = false;
        const keepTemplateImages = $("keep-template-images");
        if (keepTemplateImages) keepTemplateImages.checked = true;
      }
      document.querySelectorAll(".mode-card").forEach((el) => {
        el.classList.toggle("active", el.querySelector("input").checked);
      });
      updateTemplateUi();
    });
  });

  $("use-default-template").addEventListener("change", updateTemplateUi);

  // Sliders
  ["context-threshold", "fuzzy-threshold", "image-margin"].forEach((id) => {
    const el = $(id);
    if (!el) return;
    el.addEventListener("input", () => {
      $(id + "-val").textContent = parseFloat(el.value).toFixed(2);
    });
  });

  // File dropzones
  function bindDrop(field) {
    const input = $(`${field}-file`);
    const area = document.querySelector(`.drop-area[data-target="${field}"]`);
    const placeholder = area.querySelector(".drop-placeholder");
    const info = $(`${field}-info`);

    function setFile(f) {
      state.files[field] = f;
      if (field === "template") {
        const defaultTemplate = $("use-default-template");
        if (defaultTemplate) defaultTemplate.checked = false;
      }
      area.classList.add("has-file");
      placeholder.classList.add("hidden");
      info.classList.remove("hidden");
      info.textContent = `✓ ${f.name} (${(f.size / 1048576).toFixed(1)} MB)`;
      if (field === "template") updateTemplateUi();
    }

    input.addEventListener("change", () => { if (input.files[0]) setFile(input.files[0]); });
    area.addEventListener("dragover", (e) => { e.preventDefault(); area.classList.add("over"); });
    area.addEventListener("dragleave", () => area.classList.remove("over"));
    area.addEventListener("drop", (e) => {
      e.preventDefault();
      area.classList.remove("over");
      const f = e.dataTransfer.files[0];
      if (f) { setFile(f); const dt = new DataTransfer(); dt.items.add(f); input.files = dt.files; }
    });
  }
  bindDrop("source");
  bindDrop("template");

  // Run
  $("run-btn").addEventListener("click", async () => {
    const status = $("status-text");
    const defaultTemplate = useDefaultTemplate();
    if (!state.files.source) { status.textContent = "请选择旧版文件"; status.style.color = "var(--danger)"; return; }
    if (!defaultTemplate && !state.files.template) { status.textContent = "请选择新版模板"; status.style.color = "var(--danger)"; return; }

    status.textContent = "上传中…";
    status.style.color = "var(--text-secondary)";

    const fd = new FormData();
    fd.set("profile", state.profile);
    fd.set("use_default_template", defaultTemplate ? "1" : "0");
    fd.set("overwrite", $("overwrite").checked ? "1" : "0");
    fd.set("no_images", $("no-images").checked ? "1" : "0");
    fd.set("keep_template_images", $("keep-template-images").checked ? "1" : "0");
    fd.set("markdown", "0");
    fd.set("source", state.files.source);
    if (!defaultTemplate) fd.set("template", state.files.template);
    fd.set("context_threshold", $("context-threshold").value);
    fd.set("fuzzy_threshold", $("fuzzy-threshold").value);
    fd.set("image_margin", $("image-margin").value);
    fd.set("cross_sheet", $("cross-sheet").checked ? "1" : "0");
    fd.set("filter_status", $("filter-status").checked ? "1" : "0");
    fd.set("keep_instructional", $("keep-instructional").checked ? "1" : "0");

    $("progress-section").classList.remove("hidden");
    $("result-section").classList.add("hidden");
    $("log").textContent = "";
    $("progress-fill").style.width = "0%";
    $("run-btn").disabled = true;

    try {
      const r = await fetch("/api/migrate", { method: "POST", body: fd });
      const data = await r.json();
      if (!r.ok) throw new Error(data.error || "上传失败");
      subscribe(data.job_id);
      status.textContent = "处理中…";
    } catch (e) {
      status.textContent = e.message;
      status.style.color = "var(--danger)";
      $("run-btn").disabled = false;
    }
  });

  function subscribe(jobId) {
    const es = new EventSource(`/api/job/events?job_id=${encodeURIComponent(jobId)}`);
    es.onmessage = (ev) => {
      let d;
      try { d = JSON.parse(ev.data); } catch { return; }
      if (d.type === "progress") {
        const pct = Math.round((d.current / Math.max(d.total, 1)) * 100);
        $("progress-fill").style.width = pct + "%";
        $("progress-stage").textContent = `${d.stage} (${d.current}/${d.total})`;
        $("log").textContent += `[${d.current}/${d.total}] ${d.stage}\n`;
        $("log").scrollTop = $("log").scrollHeight;
      } else if (d.type === "log") {
        $("log").textContent += d.message + "\n";
        $("log").scrollTop = $("log").scrollHeight;
      } else if (d.type === "done") {
        $("progress-fill").style.width = "100%";
        $("progress-stage").textContent = `完成 · ${d.elapsed}s`;
        showResult(jobId, d);
        es.close();
        $("run-btn").disabled = false;
        $("status-text").textContent = "";
      } else if (d.type === "error") {
        $("log").textContent += "错误:\n" + d.message + "\n";
        $("progress-stage").textContent = "出错";
        $("progress-fill").style.background = "var(--danger)";
        es.close();
        $("run-btn").disabled = false;
        $("status-text").textContent = "出错";
        $("status-text").style.color = "var(--danger)";
      }
    };
  }

  function showResult(jobId, d) {
    $("result-section").classList.remove("hidden");
    $("kpi-row").innerHTML = [
      { v: d.filled, l: "已填充" },
      { v: d.skipped, l: "需确认" },
      { v: d.images, l: "图片" },
      { v: d.elapsed + "s", l: "用时" },
    ].map((k) => `<div class="kpi"><div class="val">${k.v}</div><div class="lbl">${k.l}</div></div>`).join("");

    $("downloads").innerHTML = (d.downloads || [])
      .map((f) => `<a class="dl-btn" href="/api/job/download?job_id=${encodeURIComponent(jobId)}&key=${encodeURIComponent(f.key)}" download="${f.filename}">${f.name}</a>`)
      .join("");
  }

  updateTemplateUi();
})();
