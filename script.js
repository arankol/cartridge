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

  function updatePreviewText() {
    reportPosition.textContent = inputPosition.value.trim() || "Инженер ГИТСиЗИ";
    reportRank.textContent = inputRank.value.trim() || "мл. лейтенант внутренней службы";
    reportFio.textContent = inputFio.value.trim() || "Иванов И.И.";
  }

  if(inputPosition) inputPosition.addEventListener("input", updatePreviewText);
  if(inputRank) inputRank.addEventListener("input", updatePreviewText);
  if(inputFio) inputFio.addEventListener("input", updatePreviewText);

  const today = new Date();
  const day = String(today.getDate()).padStart(2, "0");
  const year = today.getFullYear();
  const months = ["января","февраля","марта","апреля","мая","июня","июля","августа","сентября","октября","ноября","декабря"];
  const month = months[today.getMonth()];
  if(dateEl) dateEl.textContent = `«${day}» ${month} ${year} г.`;

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

  function reportLine(it) {
    const cName = it.cName;
    const cInv = it.cInv || "—";
    const pName = it.pName;
    const pInv = it.pInv || "—";
    return `${cName} инв. № ${cInv} к ${pName} инв. № ${pInv}`;
  }

  function renderReport() {
    if(!reportListEl) return;
    reportListEl.innerHTML = "";
    if (!state.selected.length) {
      emptyNote.style.display = "";
    } else {
      emptyNote.style.display = "none";
      state.selected.forEach((it, idx) => {
        const li = document.createElement("li");
        const isLast = idx === state.selected.length - 1;
        const punctuation = isLast ? "." : ";";
        li.textContent = reportLine(it) + punctuation;
        reportListEl.appendChild(li);
      });
    }
  }

  async function exportDocx() {
    if (!state.selected.length) return alert("Сначала добавьте картриджи.");
    
    const docx = window.docx;
    const { Document, Packer, Paragraph, TextRun, AlignmentType, TabStopType, LevelFormat } = docx;
    const FONT = "Times New Roman";
    const SIZE = 28; // 14pt

    const INDENT_11_25_CM = 6378; // 11.25 см
    const INDENT_1_25_CM = 709;   // 1.25 см
    const RIGHT_TAB_POS = 9070;

    const children = [];

    // 1. Шапка
    const pHeader = (text) => new Paragraph({
      alignment: AlignmentType.LEFT,
      indent: { left: INDENT_11_25_CM, firstLine: 0 },
      spacing: { before: 0, after: 0, line: 240 },
      children: [ new TextRun({ text, font: FONT, size: SIZE }) ]
    });
    children.push(pHeader("Начальнику МО МВД"));
    children.push(pHeader("России «Городецкий»"));
    children.push(pHeader("полковнику полиции"));
    children.push(pHeader("Трошину Н.С."));
    children.push(new Paragraph({ spacing: { after: 0, line: 240 } }));

    // 2. Заголовок "Рапорт" (Просто текст, без жирности и стилей заголовка)
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0, line: 240 },
      children: [ 
        new TextRun({ 
          text: "Рапорт", 
          font: FONT, 
          size: SIZE, 
          bold: false 
        }) 
      ]
    }));
    children.push(new Paragraph({ spacing: { after: 0, line: 240 } }));

    // 3. Вступление
    children.push(new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      indent: { left: 0, right: 0, firstLine: INDENT_1_25_CM },
      spacing: { before: 0, after: 0, line: 240 },
      children: [ new TextRun({ text: "Прошу Вас поручить службе тылового обеспечения заправить следующие картриджи:", font: FONT, size: SIZE }) ]
    }));

    // 4. Список картриджей
    state.selected.forEach((it, i) => {
      const isLast = i === state.selected.length - 1;
      const punctuation = isLast ? "." : ";";
      
      children.push(new Paragraph({
        text: reportLine(it) + punctuation,
        numbering: {
          reference: "custom-numbering",
          level: 0,
          instance: 0
        }
      }));
    });

    children.push(new Paragraph({ spacing: { after: 0, line: 240 } }));
    children.push(new Paragraph({ spacing: { after: 0, line: 240 } }));

    // 5. Подпись
    const pSign = (text) => new Paragraph({
      alignment: AlignmentType.LEFT,
      indent: { left: 0, right: 0, firstLine: 0 },
      spacing: { before: 0, after: 0, line: 240 },
      children: [ new TextRun({ text, font: FONT, size: SIZE }) ]
    });

    children.push(pSign(inputPosition.value.trim() || "Инженер ГИТСиЗИ"));
    children.push(pSign("МО МВД России «Городецкий»"));
    
    children.push(new Paragraph({
      alignment: AlignmentType.LEFT,
      indent: { left: 0, right: 0, firstLine: 0 },
      spacing: { before: 0, after: 0, line: 240 },
      tabStops: [{ type: TabStopType.RIGHT, position: RIGHT_TAB_POS }],
      children: [ new TextRun({ text: (inputRank.value.trim() || "мл. лейтенант внутренней службы") + "\t" + (inputFio.value.trim() || "Иванов И.И."), font: FONT, size: SIZE }) ]
    }));

    children.push(pSign(dateEl.textContent));

    // Настройка нумерации Word
    const numberingConfig = {
      config: [
        {
          reference: "custom-numbering",
          levels: [
            {
              level: 0,
              format: LevelFormat.DECIMAL,
              text: "%1.",
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: { left: 0, firstLine: INDENT_1_25_CM }, 
                  alignment: AlignmentType.JUSTIFIED,
                  spacing: { before: 0, after: 0, line: 240 },
                },
                run: {
                  font: FONT,
                  size: SIZE
                }
              }
            }
          ]
        }
      ]
    };

    const doc = new Document({
      numbering: numberingConfig,
      styles: {
        default: {
          document: { run: { font: FONT, size: SIZE } }
        }
      },
      sections: [{ 
        properties: { page: { margin: { top: 1134, bottom: 1134, right: 850, left: 1701 } } }, 
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
      alert("Ошибка при создании DOCX. Проверьте консоль.");
    }
  }

  if(searchInput) searchInput.addEventListener("input", (e) => { state.query = e.target.value; renderResults(); });
  const clearBtn = $("#clearBtn");
  if(clearBtn) clearBtn.addEventListener("click", () => { state.selected = []; renderSelected(); renderReport(); renderResults(); });
  
  const docxBtn = $("#docxBtn");
  if(docxBtn) docxBtn.addEventListener("click", exportDocx);

  loadExcel();
});
