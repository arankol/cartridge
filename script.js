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
  const reportRank = $("#reportRank");
  const reportFio = $("#reportFio");

  // Синхронизация полей ввода
  function updatePreviewText() {
    reportPosition.textContent = inputPosition.value.trim() || "Инженер ГИТСиЗИ";
    reportRank.textContent = inputRank.value.trim() || "мл. лейтенант внутренней службы";
    reportFio.textContent = inputFio.value.trim() || "Иванов И.И.";
  }

  if(inputPosition) inputPosition.addEventListener("input", updatePreviewText);
  if(inputRank) inputRank.addEventListener("input", updatePreviewText);
  if(inputFio) inputFio.addEventListener("input", updatePreviewText);

  // Дата
  const today = new Date();
  const day = String(today.getDate()).padStart(2, "0");
  const year = today.getFullYear();
  const months = ["января","февраля","марта","апреля","мая","июня","июля","августа","сентября","октября","ноября","декабря"];
  const month = months[today.getMonth()];
  if(dateEl) dateEl.textContent = `«${day}» ${month} ${year} г.`;

  // --- ЗАГРУЗКА EXCEL ---
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

        if (!cName || /^итого$/i.test(cName)) continue;
        out.push({ id: i, cName, cInv, pName, pInv });
      }
      state.db = out;
      renderResults();
    } catch (error) {
      console.error(error);
      if(resultsEl) resultsEl.innerHTML = `<li class="empty-hint" style="color: #dc2626;">Ошибка чтения card.xlsx.</li>`;
    }
  }

  // --- ПОИСК ---
  const norm = (s) => (s || "").toString().toLowerCase();
  const escapeHtml = (s) => (s || "").toString().replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");

  function filteredItems() {
    const q = norm(state.query).trim();
    let items = state.db;
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
    if(!resultsEl) return;
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
    if(!selectedEl) return;
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

  // --- ФОРМИРОВАНИЕ ТЕКСТА ---
  function reportLine(it) {
    // Используем обычные пробелы, так как выравнивать по ширине мы больше не будем
    return `${it.cName} инв. № ${it.cInv || "—"} к ${it.pName} инв. № ${it.pInv || "—"}`;
  }

  // --- ОТРИСОВКА ПРЕДПРОСМОТРА ---
  function renderReport() {
    if(!reportListEl) return;
    reportListEl.innerHTML = "";
    if (!state.selected.length) {
      emptyNote.style.display = "";
    } else {
      emptyNote.style.display = "none";
      state.selected.forEach((it, idx) => {
        const div = document.createElement("div");
        const isLast = idx === state.selected.length - 1;
        const punctuation = isLast ? "." : ";";
        
        // Выравниваем по левому краю и делаем висячий отступ (номер слева, текст ровно)
        div.style.textAlign = "left"; 
        div.style.paddingLeft = "1.25cm";
        div.style.textIndent = "-1.25cm";
        div.style.marginBottom = "0";
        div.textContent = `${idx + 1}. ${reportLine(it)}${punctuation}`;
        reportListEl.appendChild(div);
      });
    }
  }

  // --- СКАЧИВАНИЕ DOCX ---
  async function exportDocx() {
    if (!state.selected.length) return alert("Сначала добавьте картриджи.");
    
    const docx = window.docx;
    const { Document, Packer, Paragraph, TextRun, AlignmentType } = docx;
    const FONT = "Times New Roman";
    const SIZE = 28; // 14 pt

    // 1 см = 567 твипов (twip)
    const CM_1_25 = 709; 
    const CM_11_25 = 6378;

    const children = [];

    // 1. Шапка (Отступ слева 11.25 см)
    const pHeader = (text) => new Paragraph({
      alignment: AlignmentType.LEFT,
      indent: { left: CM_11_25 },
      spacing: { line: 240, before: 0, after: 0 },
      children: [ new TextRun({ text, font: FONT, size: SIZE }) ]
    });
    children.push(pHeader("Начальнику МО МВД"));
    children.push(pHeader("России «Городецкий»"));
    children.push(pHeader("полковнику полиции"));
    children.push(pHeader("Трошину Н.С."));
    children.push(new Paragraph({ spacing: { line: 240 } })); // Пустая строка

    // 2. Заголовок
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: 240, after: 240 },
      children: [ new TextRun({ text: "Рапорт", font: FONT, size: SIZE, bold: true }) ]
    }));

    // 3. Вступление (По ширине, отступ первой строки 1.25 см)
    children.push(new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      indent: { firstLine: CM_1_25 },
      spacing: { line: 240, after: 120 },
      children: [ new TextRun({ text: "Прошу Вас поручить службе тылового обеспечения заправить следующие картриджи:", font: FONT, size: SIZE }) ]
    }));

    // 4. Список картриджей (По ЛЕВОМУ краю с висячим отступом, чтобы не было дыр)
    state.selected.forEach((it, i) => {
      const isLast = i === state.selected.length - 1;
      const punctuation = isLast ? "." : ";";
      
      children.push(new Paragraph({
        alignment: AlignmentType.LEFT, // Выравнивание по левому краю
        indent: { left: CM_1_25, hanging: CM_1_25 }, // Висячий отступ
        spacing: { line: 240, after: 0 },
        children: [ new TextRun({ text: `${i + 1}. ${reportLine(it)}${punctuation}`, font: FONT, size: SIZE }) ]
      }));
    });

    children.push(new Paragraph({ spacing: { line: 240 } }));
    children.push(new Paragraph({ spacing: { line: 240 } }));

    // 5. Подпись
    const pSign = (text) => new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { line: 240, after: 0 },
      children: [ new TextRun({ text, font: FONT, size: SIZE }) ]
    });

    children.push(pSign(inputPosition.value.trim() || "Инженер ГИТСиЗИ"));
    children.push(pSign("МО МВД России «Городецкий»"));
    children.push(pSign(`${inputRank.value.trim() || "мл. лейтенант внутренней службы"} \t\t\t\t ${inputFio.value.trim() || "Иванов И.И."}`)); // Отступ через табы
    children.push(pSign(dateEl.textContent));

    // Настройка документа
    const doc = new Document({
      styles: { default: { document: { run: { font: FONT, size: SIZE } } } },
      sections: [{ 
        properties: { 
          page: { margin: { top: 1134, bottom: 1134, right: 850, left: 1701 } } 
        }, 
        children 
      }]
    });
    
    try {
      const blob = await Packer.toBlob(doc);
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `Рапорт_заправка_${Date.now()}.docx`;
      a.click();
    } catch (e) {
      console.error(e);
      alert("Ошибка при создании файла DOCX.");
    }
  }

  // --- СКАЧИВАНИЕ PDF ---
  function exportPdf() {
    if (!state.selected.length) return alert("Сначала добавьте картриджи.");
    
    const element = document.getElementById("report");
    
    window.html2pdf().set({
      margin: 0,
      filename: `Рапорт_заправка_${Date.now()}.pdf`,
      image: { type: "jpeg", quality: 1.0 },
      html2canvas: { 
        scale: 3, 
        useCORS: true, 
        backgroundColor: "#ffffff",
        wordSpacing: 0 
      },
      jsPDF: { unit: "mm", format: "a4", orientation: "portrait" }
    }).from(element).save();
  }

  // --- СОБЫТИЯ ---
  if(searchInput) searchInput.addEventListener("input", (e) => { state.query = e.target.value; renderResults(); });
  const clearBtn = $("#clearBtn");
  if(clearBtn) clearBtn.addEventListener("click", () => { state.selected = []; renderSelected(); renderReport(); renderResults(); });
  
  const docxBtn = $("#docxBtn");
  if(docxBtn) docxBtn.addEventListener("click", exportDocx);
  
  const pdfBtn = $("#pdfBtn");
  if(pdfBtn) pdfBtn.addEventListener("click", exportPdf);

  loadExcel();
});
