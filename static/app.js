(() => {
  const I18N = {
    ru: {
      logo: "Движок рассылок для стоматологий",
      subtitle: "Персонализированные рассылки отзывов для стоматологий",
      upload_title: "Загрузите список пациентов",
      upload_hint: "Excel, CSV или TXT. Нужны колонки с именем и телефоном.",
      drop_text: "Перетащите файл",
      or_browse: "или выберите",
      msg_title: "Текст сообщения",
      msg_hint: "Плейсхолдер",
      msg_hint2: "— имя пациента.",
      reset_btn: "Сбросить шаблон",
      run_title: "Запуск",
      run_hint: "Демо-режим: сообщения не отправляются, только отчёт.",
      run_btn: "Обработать список",
      processing: "Обработка…",
      result_title: "Результат",
      preview_title: "Превью",
      th_patient: "Пациент",
      th_phone: "Телефон",
      download_btn: "Скачать отчёт (.xlsx)",
      err_filetype: "Нужен файл .xlsx, .xls, .csv или .txt",
      err_server: "Ошибка сервера",
      err_generic: "Не удалось обработать файл",
      stat_total: "Всего",
      stat_ok: "Обработано",
      stat_skip: "Пропущено",
      template: "Здравствуйте, {name}!\nСпасибо, что посетили нашу клинику!\nНам очень важно знать ваше мнение. Если у вас есть минутка, пожалуйста, оставьте отзыв о вашем посещении в Google Maps. Это поможет другим пациентам сделать выбор и поможет нам становиться ещё лучше.\nОставить отзыв:\nhttps://g.page/ВАША_ССЫЛКА/review",
    },
    en: {
      logo: "Dental Campaign Engine",
      subtitle: "Personalized review request campaigns for dental clinics",
      upload_title: "Upload patient list",
      upload_hint: "Excel, CSV or TXT. Columns for name and phone required.",
      drop_text: "Drop file here",
      or_browse: "or browse",
      msg_title: "Message text",
      msg_hint: "Placeholder",
      msg_hint2: "— patient name.",
      reset_btn: "Reset template",
      run_title: "Run",
      run_hint: "Demo mode: messages are not sent, only a report is generated.",
      run_btn: "Process list",
      processing: "Processing…",
      result_title: "Result",
      preview_title: "Preview",
      th_patient: "Patient",
      th_phone: "Phone",
      download_btn: "Download report (.xlsx)",
      err_filetype: "Please use .xlsx, .xls, .csv or .txt",
      err_server: "Server error",
      err_generic: "Failed to process the file",
      stat_total: "Total",
      stat_ok: "Processed",
      stat_skip: "Skipped",
      template: "Hello, {name}!\nThank you for visiting our clinic!\nYour feedback means a lot to us. If you have a minute, please leave a review of your visit on Google Maps. It helps other patients choose and helps us improve.\nLeave a review:\nhttps://g.page/YOUR_LINK/review",
    },
    de: {
      logo: "Dental Campaign Engine",
      subtitle: "Personalisierte Bewertungsanfragen für Zahnarztpraxen",
      upload_title: "Patientenliste hochladen",
      upload_hint: "Excel, CSV oder TXT. Spalten für Name und Telefon erforderlich.",
      drop_text: "Datei hierher ziehen",
      or_browse: "oder auswählen",
      msg_title: "Nachrichtentext",
      msg_hint: "Platzhalter",
      msg_hint2: "— Patientenname.",
      reset_btn: "Vorlage zurücksetzen",
      run_title: "Start",
      run_hint: "Demo-Modus: Nachrichten werden nicht gesendet, nur ein Bericht wird erstellt.",
      run_btn: "Liste verarbeiten",
      processing: "Verarbeitung…",
      result_title: "Ergebnis",
      preview_title: "Vorschau",
      th_patient: "Patient",
      th_phone: "Telefon",
      download_btn: "Bericht herunterladen (.xlsx)",
      err_filetype: "Bitte .xlsx, .xls, .csv oder .txt verwenden",
      err_server: "Serverfehler",
      err_generic: "Datei konnte nicht verarbeitet werden",
      stat_total: "Gesamt",
      stat_ok: "Verarbeitet",
      stat_skip: "Übersprungen",
      template: "Guten Tag, {name}!\nVielen Dank für Ihren Besuch in unserer Praxis!\nIhr Feedback ist uns sehr wichtig. Wenn Sie eine Minute Zeit haben, hinterlassen Sie bitte eine Bewertung Ihres Besuchs bei Google Maps. Das hilft anderen Patienten bei der Wahl und uns, noch besser zu werden.\nBewertung abgeben:\nhttps://g.page/IHR_LINK/review",
    },
  };

  let lang = localStorage.getItem("dce_lang") || "ru";
  const t = () => I18N[lang];

  const fileInput = document.getElementById("file-input");
  const dropzone = document.getElementById("dropzone");
  const browseBtn = document.getElementById("browse-btn");
  const fileSelected = document.getElementById("file-selected");
  const fileNameEl = document.getElementById("file-name");
  const removeFileBtn = document.getElementById("remove-file");
  const messageTemplate = document.getElementById("message-template");
  const resetTemplateBtn = document.getElementById("reset-template");
  const runBtn = document.getElementById("run-btn");
  const btnText = runBtn.querySelector(".btn-text");
  const btnLoader = runBtn.querySelector(".btn-loader");
  const resultsSection = document.getElementById("results-section");
  const statsEl = document.getElementById("stats");
  const previewWrap = document.getElementById("preview-wrap");
  const previewTableBody = document.querySelector("#preview-table tbody");
  const downloadBtn = document.getElementById("download-btn");
  const errorBanner = document.getElementById("error-banner");
  const errorText = document.getElementById("error-text");
  const errorClose = document.getElementById("error-close");

  let selectedFile = null;

  function applyLang() {
    document.querySelectorAll("[data-i18n]").forEach((el) => {
      const key = el.getAttribute("data-i18n");
      if (t()[key] !== undefined) el.textContent = t()[key];
    });
    document.querySelectorAll(".lang-btn").forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.lang === lang);
    });
    messageTemplate.value = t().template;
    document.documentElement.lang = lang;
    localStorage.setItem("dce_lang", lang);
  }

  document.querySelectorAll(".lang-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      lang = btn.dataset.lang;
      applyLang();
    });
  });

  applyLang();

  function setFile(file) {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["xlsx", "xls", "csv", "txt"].includes(ext)) {
      showError(t().err_filetype);
      return;
    }
    selectedFile = file;
    fileNameEl.textContent = file.name;
    dropzone.querySelector(".dropzone-content").hidden = true;
    fileSelected.hidden = false;
    runBtn.disabled = false;
    hideError();
    resultsSection.hidden = true;
  }

  function clearFile() {
    selectedFile = null;
    fileInput.value = "";
    dropzone.querySelector(".dropzone-content").hidden = false;
    fileSelected.hidden = true;
    runBtn.disabled = true;
  }

  browseBtn.addEventListener("click", (e) => { e.stopPropagation(); fileInput.click(); });
  dropzone.addEventListener("click", () => { if (!selectedFile) fileInput.click(); });
  fileInput.addEventListener("change", () => { if (fileInput.files[0]) setFile(fileInput.files[0]); });
  removeFileBtn.addEventListener("click", (e) => { e.stopPropagation(); clearFile(); });

  ["dragenter", "dragover"].forEach((evt) => {
    dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.add("dragover"); });
  });
  ["dragleave", "drop"].forEach((evt) => {
    dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.remove("dragover"); });
  });
  dropzone.addEventListener("drop", (e) => {
    const file = e.dataTransfer.files[0];
    if (file) setFile(file);
  });

  resetTemplateBtn.addEventListener("click", () => { messageTemplate.value = t().template; });

  function showError(msg) { errorText.textContent = msg || ""; errorBanner.hidden = !msg; }
  function hideError() { errorBanner.hidden = true; }
  errorClose.addEventListener("click", hideError);

  runBtn.addEventListener("click", async () => {
    if (!selectedFile) return;
    hideError();
    resultsSection.hidden = true;
    runBtn.disabled = true;
    btnText.hidden = true;
    btnLoader.hidden = false;

    const formData = new FormData();
    formData.append("file", selectedFile);
    formData.append("message_template", messageTemplate.value);

    try {
      const res = await fetch("/api/process", { method: "POST", body: formData });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || t().err_server);
      renderResults(data);
    } catch (err) {
      showError(err.message || t().err_generic);
    } finally {
      runBtn.disabled = false;
      btnText.hidden = false;
      btnLoader.hidden = true;
    }
  });

  function renderResults(data) {
    const tr = t();
    statsEl.innerHTML =
      '<div class="stat-box muted"><div class="value">' + data.total + '</div><div class="label">' + tr.stat_total + '</div></div>' +
      '<div class="stat-box success"><div class="value">' + data.processed + '</div><div class="label">' + tr.stat_ok + '</div></div>' +
      '<div class="stat-box warning"><div class="value">' + data.skipped + '</div><div class="label">' + tr.stat_skip + '</div></div>';

    if (data.preview && data.preview.length) {
      previewTableBody.innerHTML = data.preview.map(function (row) {
        return "<tr><td>" + esc(row["Пациент"]) + "</td><td>" + esc(row["Телефон"]) + "</td><td>" + esc(row["SMS"]) + "</td><td>" + esc(row["WhatsApp"]) + "</td></tr>";
      }).join("");
      previewWrap.hidden = false;
    } else {
      previewWrap.hidden = true;
    }

    if (data.report_file) {
      downloadBtn.href = "/api/download/" + data.report_file;
      downloadBtn.hidden = false;
    } else {
      downloadBtn.hidden = true;
    }

    resultsSection.hidden = false;
    resultsSection.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }

  function esc(str) {
    const d = document.createElement("div");
    d.textContent = str == null ? "" : str;
    return d.innerHTML;
  }
})();
