document.addEventListener("DOMContentLoaded", () => {
  const state = { db: [], selected: [], query: "" };
  
  const $ = (sel) => document.querySelector(sel);
  const searchInput = $("#search");
  const resultsEl = $("#results");
  const selectedEl = $("#selected");
  const countEl = $("#count");
  const reportListEl = $("#reportList");
  const emptyNote = $("#emptyNote");
  const dateEl = $("#date");
  
  const inputPosition = $("#inputPosition");
  const inputRank = $("#inputRank");
  const reportPosition = $("#reportPosition");
  const reportRank = $("#reportRank");

  // Мгновенное обновление предпросмотра при вводе текста
  if (inputPosition) {
    inputPosition.addEventListener("input", (e) => {
      reportPosition.textContent = e.target.value.trim() || "Главный специалист-эксперт ОТО";
    });
  }
  if (inputRank) {
    inputRank.addEventListener("input", (e) => {
      reportRank.textContent = e.target.value.trim() || "капитан полиции Иванов И.И.";
    });
  }

  // Установка текущей даты
  const today = new Date();
  if (dateEl) {
    dateEl.textContent = `«${String(today.getDate()).padStart(2,"0")}» ${
      ["января","февраля","марта","апреля","мая","июня","июля","августа","сентября","октября","ноября","декабря"][today.getMonth()]
    } ${today.getFullYear()} г.`;
  }

  // --- ЗАГРУЗКА И ЧТЕНИЕ EXCEL ---
  async function loadExcel() {
    try {
      const response = await fetch("card.xlsx");
      if (!response.ok) throw new Error("Файл не найден");
      
      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false, defval: "" });
      
      const out = [];
      // Пропускаем первую строку (заголовки), начинаем со второй (индекс 1)
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        const cName = (r[0] || "").toString().trim();
        const cInv  = (r[1] || "").toString().trim();
        const pName = (r[2] || "").toString().trim();
        const pInv  = (r[3] || "").toString().trim();

        if (!cName) continue; // Пропускаем пустые строки
        if (/^итого$/i.test(cName)) continue; // Пропускаем строку "Итого"

        // УНИКАЛЬНЫЙ ID = номер строки в Excel. Это исключает дубликаты.
        out.push({ id: i, cName, cInv, pName, pInv });
      }
      state.db = out;
      renderResults();
    } catch (error) {
      console.error(error);
      if (resultsEl) {
        resultsEl.innerHTML = `
          <li class="empty-hint" style="color: #dc2626;">
            Ошибка чтения card.xlsx.<br><br>
            Убедитесь, что файл лежит в корне репозитория.
          </li>`;
      }
    }
  }

  // --- ОТРИСОВКА И ПОИСК ---
  const norm = (s) => (s || "").toString().toLowerCase();
  const escapeHtml = (s) => (s || "").toString().replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");

  function filteredItems() {
    const q = norm(state.query).trim();
    
    // Берем всю базу
    let items = state.db;
    
    // Отфильтровываем те, чей ID уже есть в выбранных
    const selectedIds = new Set(state.selected.map(s => s.id));
    items = items.filter(it => !selectedIds.has(it.id));

    if (!q) return items.slice(0, 300);
    return items.filter(it =>
      norm(it.cName).includes(q) ||
      norm(it.cInv).includes(q) ||
      norm(it.pName).includes(q) ||
      norm(it.pInv).includes(q)
    ).slice(0, 500);
  }

  function renderResults() {
    if (!resultsEl) return;
    const items = filteredItems();
    resultsEl.innerHTML = "";
    if (!items.length) {
      resultsEl.innerHTML = `<li class="empty-hint">Ничего не найдено</li>`;
      return;
    }
    
    items.forEach(it => {
      const li = document.createElement("li");
      li.innerHTML = `
        <div><strong>${escapeHtml(it.cName)}</strong></div>
        <div class="muted">Инв. № ${escapeHtml(it.cInv || "—")} → ${escapeHtml(it.pName)} (${escapeHtml(it.pInv || "—")})</div>
      `;
      li.addEventListener("click", () => addItem(it));
      resultsEl.appendChild(li);
    });
  }

  function addItem(it) {
    state.selected.push(it);
    renderSelected();
    renderReport();
    renderResults(); // Обновляем поиск, чтобы выбранный исчез
  }

  function removeItem(idx) {
    state.selected.splice(idx, 1);
    renderSelected();
    renderReport();
    renderResults(); // Обновляем поиск, чтобы удаленный вернулся
  }

  function renderSelected() {
    if (!selectedEl) return;
    countEl.textContent = state.selected.length;
    selectedEl.innerHTML = "";
    state.selected.forEach((it, idx) => {
      const li = document.createElement("li");
      li.innerHTML = `
        <div>
          <div><strong>${escapeHtml(it.cName)}</strong> — ${escapeHtml(it.cInv || "—")}</div>
          <div class="muted">${escapeHtml(it.pName)} (${escapeHtml(it.pInv || "—")})</div>
        </div>
        <button title="Удалить">×</button>
      `;
      li.querySelector("button").addEventListener("click", () => removeItem(idx));
      selectedEl.appendChild(li);
    });
  }

  // --- ФОРМИРОВАНИЕ РАПОРТА ---
  function reportLine(it) {
    return `${it.cName} инв. № ${it.cInv || "—"} к ${it.pName} инв. № ${it.pInv || "—"}`;
  }

  function renderReport() {
    if (!reportListEl) return;
    reportListEl.innerHTML = "";
    if (!state.selected.length) {
      emptyNote.style.display = "";
    } else {
      emptyNote.style.display = "none";
      state.selected.forEach(it => {
        const li = document.createElement("li");
        li.textContent = reportLine(it) + ";";
        reportListEl.appendChild(li);
      });
    }
  }

  // --- СКАЧИВАНИЕ DOCX ---
  async function exportDocx() {
    if (!state.selected.length) return alert("Сначала добавьте картриджи.");
    const { Document, Packer, Paragraph, TextRun, AlignmentType } = window.docx;
    const FONT = "Times New Roman";
    const SIZE = 28;

    const p = (text, opts = {}) => new Paragraph({
      alignment: opts.align || AlignmentType.JUSTIFIED,
      spacing: { line: 360, after: opts.after ?? 120 },
      indent: opts.indent,
      children: [ new TextRun({ text, font: FONT, size: SIZE, bold: opts.bold }) ]
    });

    const children = [];
    children.push(p("Начальнику МО МВД России «Городецкий»", { align: AlignmentType.RIGHT, after: 0 }));
    children.push(p("полковнику полиции", { align: AlignmentType.RIGHT, after: 0 }));
    children.push(p("Трошину Н.С.", { align: AlignmentType.RIGHT, after: 240 }));
    children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [ new TextRun({ text: "Рапорт", font: FONT, size: 32, bold: true }) ] }));
    children.push(p("Прошу Вас поручить службе тылового обеспечения заправить следующие картриджи:", { indent: { firstLine: 720 }, after: 160 }));
    
    state.selected.forEach((it, i) => {
      children.push(p(`${i + 1}. ${reportLine(it)};`, { indent: { left: 720, hanging: 360 }, after: 120 }));
    });

    children.push(p("", { after: 240 }));
    children.push(p(reportPosition.textContent, { align: AlignmentType.LEFT, after: 0 }));
    children.push(p("МО МВД России «Городецкий»", { align: AlignmentType.LEFT, after: 0 }));
    children.push(p(reportRank.textContent, { align: AlignmentType.LEFT, after: 0 }));
    children.push(p(dateEl.textContent, { align: AlignmentType.LEFT, after: 0 }));

    const doc = new Document({
      styles: { default: { document: { run: { font: FONT, size: SIZE } } } },
      sections: [{ properties: { page: { margin: { top: 1134, bottom: 1134, left: 1701, right: 850 } } }, children }]
    });
    
    const blob = await Packer.toBlob(doc);
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `Рапорт_заправка_${Date.now()}.docx`;
    a.click();
  }

  // --- СКАЧИВАНИЕ PDF ---
  function exportPdf() {
    if (!state.selected.length) return alert("Сначала добавьте картриджи.");
    window.html2pdf().set({
      margin: 0,
      filename: `Рапорт_заправка_${Date.now()}.pdf`,
      image: { type: "jpeg", quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true, backgroundColor: "#ffffff" },
      jsPDF: { unit: "mm", format: "a4", orientation: "portrait" }
    }).from(document.getElementById("report")).save();
  }

  // --- СОБЫТИЯ ---
  if (searchInput) {
    searchInput.addEventListener("input", (e) => { state.query = e.target.value; renderResults(); });
  }
  
  const clearBtn = $("#clearBtn");
  if (clearBtn) {
    clearBtn.addEventListener("click", () => { state.selected = []; renderSelected(); renderReport(); renderResults(); });
  }

  const docxBtn = $("#docxBtn");
  if (docxBtn) {
    docxBtn.addEventListener("click", exportDocx);
  }

  const pdfBtn = $("#pdfBtn");
  if (pdfBtn) {
    pdfBtn.addEventListener("click", exportPdf);
  }

  // Запускаем загрузку Excel при старте
  loadExcel();
});
