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
  const inputFio = $("#inputFio");
  
  const reportPosition = $("#reportPosition");
  const reportRankAndFio = $("#reportRankAndFio");

  // Синхронизация полей ввода с предпросмотром
  function updatePreviewText() {
    reportPosition.textContent = inputPosition.value.trim() || "Инженер ГИТСиЗИ";
    const rank = inputRank.value.trim();
    const fio = inputFio.value.trim();
    reportRankAndFio.textContent = `${rank} ${fio}`.trim() || "мл. лейтенант внутренней службы Иванов И.И.";
  }

  inputPosition.addEventListener("input", updatePreviewText);
  inputRank.addEventListener("input", updatePreviewText);
  inputFio.addEventListener("input", updatePreviewText);

  // Установка текущей даты (в формате «19» октября 2026 г.)
  const today = new Date();
  const day = String(today.getDate()).padStart(2, "0");
  const year = today.getFullYear();
  const months = ["января","февраля","марта","апреля","мая","июня","июля","августа","сентября","октября","ноября","декабря"];
  const month = months[today.getMonth()];
  dateEl.textContent = `«${day}» ${month} ${year} г.`;

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
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        const cName = (r[0] || "").toString().trim();
        const cInv  = (r[1] || "").toString().trim();
        const pName = (r[2] || "").toString().trim();
        const pInv  = (r[3] || "").toString().trim();

        if (!cName) continue;
        if (/^итого$/i.test(cName)) continue;

        out.push({ id: i, cName, cInv, pName, pInv });
      }
      state.db = out;
      renderResults();
    } catch (error) {
      console.error(error);
      resultsEl.innerHTML = `<li class="empty-hint" style="color: #dc2626;">Ошибка чтения card.xlsx. Убедитесь, что файл лежит в корне репозитория.</li>`;
    }
  }

  // --- ОТРИСОВКА И ПОИСК ---
  const norm = (s) => (s || "").toString().toLowerCase();
  const escapeHtml = (s) => (s || "").toString().replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");

  function filteredItems() {
    const q = norm(state.query).trim();
    let items = state.db;
    
    // Скрываем уже выбранные
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
    renderResults();
  }

  function removeItem(idx) {
    state.selected.splice(idx, 1);
    renderSelected();
    renderReport();
    renderResults();
  }

  function renderSelected() {
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

  // --- СКАЧИВАНИЕ DOCX (С ТОЧНЫМ ФОРМАТИРОВАНИЕМ) ---
  async function exportDocx() {
    if (!state.selected.length) return alert("Сначала добавьте картриджи.");
    const { Document, Packer, Paragraph, TextRun, AlignmentType, convertMillimetersToTwip } = window.docx;
    const FONT = "Times New Roman";
    const SIZE = 28; // 14 pt

    // Вспомогательная функция для создания параграфа
    const p = (text, opts = {}) => new Paragraph({
      alignment: opts.align || AlignmentType.JUSTIFIED,
      spacing: { line: 360, before: 0, after: 0 }, // Одинарный интервал
      indent: opts.indent,
      children: [ new TextRun({ text: text, font: FONT, size: SIZE, bold: opts.bold }) ]
    });

    const children = [];

    // 1. Шапка (справа, без отступа)
    children.push(p("Начальнику МО МВД", { align: AlignmentType.RIGHT, indent: { left: 0, firstLine: 0 } }));
    children.push(p("России «Городецкий»", { align: AlignmentType.RIGHT, indent: { left: 0, firstLine: 0 } }));
    children.push(p("полковнику полиции", { align: AlignmentType.RIGHT, indent: { left: 0, firstLine: 0 } }));
    children.push(p("Трошину Н.С.", { align: AlignmentType.RIGHT, indent: { left: 0, firstLine: 0 }, after: 240 })); // Отступ после шапки

    // 2. Заголовок (по центру, жирный)
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 240 },
      children: [ new TextRun({ text: "Рапорт", font: FONT, size: SIZE, bold: true }) ]
    }));

    // 3. Вступление (по ширине, отступ первой строки 1.25 см)
    children.push(p("Прошу Вас поручить службе тылового обеспечения заправить следующие картриджи:", {
      indent: { firstLine: convertMillimetersToTwip(12.5) }, after: 120
    }));

    // 4. Список картриджей (по ширине, отступ первой строки 1.25 см)
    state.selected.forEach((it, i) => {
      children.push(p(`${i + 1}. ${reportLine(it)};`, {
        indent: { firstLine: convertMillimetersToTwip(12.5) }, after: 0
      }));
    });

    // 5. Подпись (отступ слева 11.25 см, без отступа первой строки)
    children.push(p("", { after: 240 })); // Пустая строка
    
    // Функция для строк подписи
    const pSign = (text) => new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { line: 360, before: 0, after: 0 },
      indent: { left: convertMillimetersToTwip(112.5), firstLine: 0 },
      children: [ new TextRun({ text: text, font: FONT, size: SIZE }) ]
    });

    children.push(pSign(inputPosition.value.trim() || "Инженер ГИТСиЗИ"));
    children.push(pSign("МО МВД России «Городецкий»"));
    children.push(pSign(`${inputRank.value.trim()} ${inputFio.value.trim()}`));
    children.push(pSign(dateEl.textContent));

    // Создание документа
    const doc = new Document({
      styles: { default: { document: { run: { font: FONT, size: SIZE } } } },
      sections: [{ 
        properties: { 
          page: { 
            // Стандартные поля страницы Word (2 см сверху/снизу, 1.5 см справа, 3 см слева)
            margin: { 
              top: convertMillimetersToTwip(20), 
              bottom: convertMillimetersToTwip(20), 
              right: convertMillimetersToTwip(15), 
              left: convertMillimetersToTwip(30) 
            } 
          } 
        }, 
        children 
      }]
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
  searchInput.addEventListener("input", (e) => { state.query = e.target.value; renderResults(); });
  $("#clearBtn").addEventListener("click", () => { state.selected = []; renderSelected(); renderReport(); renderResults(); });
  $("#docxBtn").addEventListener("click", exportDocx);
  $("#pdfBtn").addEventListener("click", exportPdf);

  // Запускаем загрузку Excel при старте
  loadExcel();
});
