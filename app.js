/* global XLSX, pdfjsLib */
(function () {
  const state = {
    workpackage: [],
    baseEin: [],
    baseEinWorkbook: null,
    master: [],
    mats: [],
    greenEdits: { master: new Set(), mats: new Set() }
  };

  const CONFIG = {
    teamResolverRules: [
      { contains: "#1 ENGINE", team: "G1" },
      { contains: "LEFT HAND ENGINE", team: "G1" },
      { contains: "NO1 ENG", team: "G1" },
      { contains: "#2 ENGINE", team: "G2" },
      { contains: "RIGHT HAND ENGINE", team: "G2" },
      { contains: "NO2 ENG", team: "G2" }
    ],
    requiredMasterFields: ["team", "anexo", "irCt"],
    requiredMatsFields: ["pn", "descripcion", "qty"],
    manualKeywords: ["MANUAL", "TBD"],
    matEligibleValue: "SI"
  };

  const NON_EDITABLE_FIELDS = new Set([
    "status",
    "manualReasons",
    "num",
    "wo",
    "title"
  ]);

  const statusBox = document.getElementById("statusBox");
  const masterSummary = document.getElementById("masterSummary");
  const matsSummary = document.getElementById("matsSummary");
  const masterTable = document.getElementById("masterTable");
  const matsTable = document.getElementById("matsTable");

  function log(message) {
    const now = new Date().toLocaleTimeString();
    statusBox.textContent += `[${now}] ${message}\n`;
    statusBox.scrollTop = statusBox.scrollHeight;
  }

  function clearLog() {
    statusBox.textContent = "";
  }

  function normalize(value) {
    return String(value || "").trim();
  }

  function normalizeUpper(value) {
    return normalize(value).toUpperCase();
  }

  function normalizeForMatch(value) {
    return normalizeUpper(value)
      .replace(/\s*-\s*/g, "-")
      .replace(/\s*\/\s*/g, "/")
      .replace(/\s+/g, " ")
      .trim();
  }

  function normalizeTitleDisplay(value) {
    let t = normalize(value).replace(/\s+/g, " ");
    // Compacta separadores en codigos tecnicos (ej. ALT - 25 - 0038 -> ALT-25-0038)
    t = t.replace(/([A-Za-z0-9])\s*-\s*([A-Za-z0-9])/g, "$1-$2");
    t = t.replace(/([A-Za-z0-9])\s*\/\s*([A-Za-z0-9])/g, "$1/$2");
    return t.trim();
  }

  function containsText(value) {
    const v = normalize(value);
    if (!v) return false;
    if (!Number.isNaN(Number(v))) return false;
    return /[A-Za-zÁÉÍÓÚáéíóúÑñ]/.test(v);
  }

  function extractTextUntilDelimiter(text) {
    const t = normalize(text);
    const idx = t.indexOf(" ");
    return idx > -1 ? t.slice(0, idx).trim() : t;
  }

  function hasSpecificPattern(value) {
    const v = normalizeUpper(value);
    return /^ALT-\d{2}-\d{4}-[A-Z]{2}/.test(v) || /^ZL-\d{3}-\d{2}-\d-[A-Z]{2}/.test(v);
  }

  function hasPriorityPattern(value) {
    return /^[A-Za-z]{3}-\d{2}-\d{4}-[A-Za-z]{2}/.test(normalize(value));
  }

  function hasPnPattern(value) {
    const v = normalizeForMatch(value);
    return v.includes("PN:") && (/PN:\d{3}-\d{4}-\d{2}/.test(v) || /PN:\d{7}-\d{2}/.test(v));
  }

  function extractNPattern(value) {
    const v = normalizeForMatch(value);
    const match = v.match(/\d{6}-[A-Za-z]\d(?:-\d)?(?:-[A-Za-z]{2}(?:\/\d)?)?/);
    return match ? match[0] : "";
  }

  function resolveTeamFromDescription(description) {
    const d = normalizeUpper(description);
    if (!d) return "";
    for (const rule of CONFIG.teamResolverRules) {
      if (d.includes(normalizeUpper(rule.contains))) return normalize(rule.team);
    }
    return "";
  }

  function containsManualKeyword(value) {
    const v = normalizeUpper(value);
    return CONFIG.manualKeywords.some((keyword) => v.includes(normalizeUpper(keyword)));
  }

  function pickBestBaseMatch(title, baseEin) {
    let bestIndex = -1;
    let bestScore = { bucket: 0, length: 0 };
    const t = normalize(title);
    const tMatch = normalizeForMatch(title);
    if (!t) return bestIndex;

    if (hasPnPattern(t)) {
      const p = extractNPattern(t);
      if (p) {
        const idx = baseEin.findIndex((row) => normalizeUpper(row.criterio).includes(normalizeUpper(p)));
        if (idx > -1) return idx;
      }
    }

    for (let i = 0; i < baseEin.length; i += 1) {
      const key = normalize(baseEin[i].criterio);
      if (!key || key.length < 3) continue;
      if (!tMatch.includes(normalizeForMatch(key))) continue;

      let bucket = 1;
      if (hasSpecificPattern(key)) bucket = 2;
      if (hasPriorityPattern(key)) bucket = 3;

      if (bucket > bestScore.bucket || (bucket === bestScore.bucket && key.length > bestScore.length)) {
        bestIndex = i;
        bestScore = { bucket, length: key.length };
      }
    }
    return bestIndex;
  }

  function evaluateMasterRow(row) {
    const reasons = [];
    const teamUpper = normalizeUpper(row.team);

    if (teamUpper === "GX") {
      const resolved = resolveTeamFromDescription(row.title);
      if (resolved) row.team = resolved;
      else reasons.push("TEAM=GX sin resolucion automatica");
    } else if (teamUpper === "LX") {
      reasons.push("TEAM=LX requiere control manual");
    }

    ["team", "amtoss", "remarks", "anexo", "irCt", "matSiNo"].forEach((field) => {
      if (containsManualKeyword(row[field])) reasons.push(`Campo ${field} contiene palabra de control manual`);
    });

    CONFIG.requiredMasterFields.forEach((field) => {
      if (!normalize(row[field])) reasons.push(`Campo obligatorio vacio: ${field}`);
    });

    row.manualReasons = reasons;
    row.status = reasons.length ? "CONTROL MANUAL" : "OK";
  }

  function evaluateMatsRow(row) {
    const reasons = [];
    CONFIG.requiredMatsFields.forEach((field) => {
      if (!normalize(row[field])) reasons.push(`Campo obligatorio vacio: ${field}`);
    });
    if (!normalize(row.team)) reasons.push("TEAM vacio");
    row.manualReasons = reasons;
    row.status = reasons.length ? "CONTROL MANUAL" : "OK";
  }

  function buildMaster(workpackage, baseEin) {
    const rows = [];

    for (let i = 0; i < workpackage.length; i += 1) {
      const wp = workpackage[i];
      const title = normalize(wp.title);
      if (!title) continue;

      const woRaw = normalize(wp.wo);
      const wo = containsText(woRaw) ? extractTextUntilDelimiter(title) : woRaw;
      const matchIndex = pickBestBaseMatch(title, baseEin);

      if (matchIndex === -1) {
        rows.push({
          wpIndex: i,
          num: normalize(wp.num),
          wo,
          title: title.slice(0, 200),
          team: "",
          amtoss: "CONTROL MANUAL",
          remarks: "",
          anexo: "",
          irCt: "",
          matSiNo: "",
          criterioBusqueda: "SIN CRITERIO",
          status: "CONTROL MANUAL",
          manualReasons: ["Sin coincidencia de criterio en Base de datos"]
        });
        continue;
      }

      const base = baseEin[matchIndex];
      const row = {
        wpIndex: i,
        num: normalize(wp.num),
        wo,
        title: title.slice(0, 200),
        team: normalize(base.team),
        amtoss: normalize(base.amtoss),
        remarks: normalize(base.remarks),
        anexo: normalize(base.anexo),
        irCt: normalize(base.irCt),
        matSiNo: normalize(base.matSiNo),
        criterioBusqueda: normalize(base.criterio),
        status: "OK",
        manualReasons: []
      };

      evaluateMasterRow(row);
      rows.push(row);
    }

    return rows;
  }

  function buildMats(workpackage, baseEin, masterRows) {
    const rows = [];
    const matEligible = normalizeUpper(CONFIG.matEligibleValue);
    const masterByWp = new Map();
    masterRows.forEach((m) => {
      if (!masterByWp.has(m.wpIndex)) masterByWp.set(m.wpIndex, m);
    });

    for (let i = 0; i < workpackage.length; i += 1) {
      const wp = workpackage[i];
      const title = normalize(wp.title);
      if (!title) continue;

      const woRaw = normalize(wp.wo);
      const wo = containsText(woRaw) ? extractTextUntilDelimiter(title) : woRaw;
      const matchIndex = pickBestBaseMatch(title, baseEin);

      if (matchIndex === -1) {
        rows.push({
          wpIndex: i,
          num: normalize(wp.num),
          wo,
          title: title.slice(0, 200),
          team: "CONTROL MANUAL",
          pn: "",
          descripcion: "",
          qty: "",
          type: "",
          notas: "",
          matSiNo: "",
          criterioBusqueda: "SIN CRITERIO",
          status: "CONTROL MANUAL",
          manualReasons: ["Sin coincidencia de criterio en Base de datos"]
        });
        continue;
      }

      const selectedKey = normalizeUpper(baseEin[matchIndex].criterio);
      const candidates = baseEin.filter((row) => normalizeUpper(row.criterio) === selectedKey);
      const withMat = candidates.filter((row) => normalizeUpper(row.matSiNo) === matEligible);
      if (withMat.length === 0) continue;

      for (const base of withMat) {
        const masterRow = masterByWp.get(i);
        const row = {
          wpIndex: i,
          num: normalize(wp.num),
          wo,
          title: title.slice(0, 200),
          team: normalize(masterRow ? masterRow.team : base.team),
          pn: normalize(base.pn),
          descripcion: normalize(base.descripcion),
          qty: normalize(base.qty),
          type: normalize(base.type),
          notas: normalize(base.notas),
          matSiNo: normalize(base.matSiNo),
          criterioBusqueda: normalize(base.criterio),
          status: "OK",
          manualReasons: []
        };

        if (normalizeUpper(row.team) === "GX") {
          const resolved = resolveTeamFromDescription(title);
          if (resolved) row.team = resolved;
        }

        evaluateMatsRow(row);
        rows.push(row);
      }
    }

    return rows;
  }

  function getStats(rows) {
    const total = rows.length;
    const manual = rows.filter((r) => r.status === "CONTROL MANUAL").length;
    return { total, manual, automatico: total - manual };
  }

  function badge(status) {
    if (status === "CONTROL MANUAL") return '<span class="badge manual">CONTROL MANUAL</span>';
    return '<span class="badge ok">OK</span>';
  }

  const HEADER_LABELS = {
    num: "Nº", wo: "WO", title: "TITLE", team: "TEAM",
    amtoss: "AMTOSS", remarks: "REMARKS", anexo: "ANEXO",
    irCt: "IR/CT", matSiNo: "MAT SI/NO",
    criterioBusqueda: "CRITERIO BÚSQUEDA", status: "STATUS",
    manualReasons: "RAZONES CONTROL MANUAL",
    pn: "PN", descripcion: "DESCRIPCIÓN", qty: "QTY",
    type: "TYPE", notas: "NOTAS"
  };

  const activeFilters = { master: null, mats: null }; // { field, value } | null

  function getDatasetRows(dataset) {
    return dataset === "master" ? state.master : state.mats;
  }

  function renderTable(table, rows, datasetName) {
    if (!rows.length) {
      table.innerHTML = "<tr><td>No hay datos.</td></tr>";
      return;
    }

    const greenSet = state.greenEdits[datasetName] || new Set();
    const allRows = getDatasetRows(datasetName);
    const headers = Object.keys(rows[0]).filter((h) => h !== "wpIndex");

    const COL_WIDTHS = { num: "38px", title: "180px" };
    const colgroup = `<colgroup>${headers.map((h) => COL_WIDTHS[h] ? `<col style="width:${COL_WIDTHS[h]}">` : "<col>").join("")}</colgroup>`;

    // Fila de filtros tipo Excel (dentro del thead para que sticky funcione correctamente)
    const af = activeFilters[datasetName];
    const filterCells = headers.map((field) => {
      if (field === "manualReasons") return "<th></th>";
      const unique = [...new Set(allRows.map((r) => normalize(String(r[field] ?? ""))))]
        .filter((v) => v !== "")
        .sort((a, b) => a.localeCompare(b, "es", { numeric: true }));
      const active = af && af.field === field;
      const activeClass = active ? " class=\"active-filter\"" : "";
      const options = '<option value="">— todos —</option>' +
        unique.map((v) => `<option value="${v}"${active && af.value === v ? " selected" : ""}>${v}</option>`).join("");
      return `<th><select${activeClass} data-filter-dataset="${datasetName}" data-filter-field="${field}">${options}</select></th>`;
    }).join("");

    const thead = `<thead><tr>${headers.map((h) => `<th>${HEADER_LABELS[h] ?? h}</th>`).join("")}</tr><tr class="filter-row">${filterCells}</tr></thead>`;

    const tbody = rows.map((row, rowIdx) => {
      // rowIdx here is position in the FILTERED view; we need the original index for green tracking
      // We reconstruct originalIdx by finding this row object in allRows
      const originalIdx = allRows.indexOf(row);
      const cls = row.status === "CONTROL MANUAL" ? "manual" : "";
      const tds = headers.map((field) => {
        if (field === "status") return `<td title="${row.status}">${badge(row.status)}</td>`;
        if (field === "manualReasons") { const v = (row.manualReasons || []).join(" | "); return `<td title="${v}">${v}</td>`; }

        const val = String(row[field] ?? "");
        const editable = !NON_EDITABLE_FIELDS.has(field);
        if (!editable) return `<td title="${val}">${val}</td>`;

        const isGreen = greenSet.has(`${originalIdx}-${field}`);
        const cellClass = `editable${isGreen ? " green-marked" : ""}`;
        return `<td title="${val}" contenteditable="true" class="${cellClass}" data-editable="1" data-dataset="${datasetName}" data-row="${originalIdx}" data-field="${field}">${val}</td>`;
      }).join("");

      return `<tr class="${cls}">${tds}</tr>`;
    }).join("");

    table.innerHTML = `${colgroup}${thead}<tbody>${tbody}</tbody>`;

    // Attach filter change listeners
    table.querySelectorAll("tr.filter-row select").forEach((sel) => {
      sel.addEventListener("change", () => {
        const ds = sel.dataset.filterDataset;
        const fld = sel.dataset.filterField;
        const val = sel.value;
        if (!val) {
          // Only clear if this field was the active filter
          if (activeFilters[ds] && activeFilters[ds].field === fld) activeFilters[ds] = null;
        } else {
          activeFilters[ds] = { field: fld, value: val };
        }
        const tbl = ds === "master" ? masterTable : matsTable;
        renderTableFiltered(tbl, ds);
      });
    });
  }

  function applyRowWindow(table, rowLimit) {
    const wrapper = table.closest(".table-wrap");
    if (!wrapper) return;

    const header = table.querySelector("thead");
    const bodyRows = Array.from(table.querySelectorAll("tbody tr"));
    if (!header || bodyRows.length === 0) {
      wrapper.style.maxHeight = "";
      return;
    }

    const visibleRows = bodyRows.slice(0, rowLimit);
    const rowsHeight = visibleRows.reduce((sum, row) => sum + row.offsetHeight, 0);
    const totalHeight = header.offsetHeight + rowsHeight + 2;
    wrapper.style.maxHeight = `${Math.ceil(totalHeight)}px`;
  }

  // ═══════════════════════════════════════════════════════════
  //  FILTROS INLINE Y EDICIÓN EN BLOQUE POR TABLA
  // ═══════════════════════════════════════════════════════════

  function getVisibleRows(dataset) {
    const rows = getDatasetRows(dataset);
    const f = activeFilters[dataset];
    if (!f) return rows;
    return rows.filter((r) => normalizeUpper(String(r[f.field] ?? "")) === normalizeUpper(f.value));
  }

  function renderTableFiltered(table, dataset) {
    renderTable(table, getVisibleRows(dataset), dataset);
  }

  // Poblar selectores de columna para los bulk-bars de cada tabla
  function populateFieldSelectors() {
    function buildOptions(rows) {
      if (!rows.length) return '<option value="">— columna —</option>';
      const fields = Object.keys(rows[0]).filter((f) => f !== "wpIndex" && f !== "status" && f !== "manualReasons");
      return '<option value="">— columna —</option>' +
        fields.map((f) => `<option value="${f}">${HEADER_LABELS[f] ?? f}</option>`).join("");
    }
    document.getElementById("masterBulkField").innerHTML = buildOptions(state.master);
    document.getElementById("matsBulkField").innerHTML = buildOptions(state.mats);
  }

  function applyBulkEdit(dataset) {
    const field = document.getElementById(`${dataset}BulkField`).value;
    const fromVal = normalize(document.getElementById(`${dataset}BulkFrom`).value);
    const toVal = normalize(document.getElementById(`${dataset}BulkTo`).value);

    if (!field) { log("Selecciona una columna para la edición en bloque."); return; }

    const rows = getDatasetRows(dataset);
    const greenSet = state.greenEdits[dataset];
    let count = 0;

    rows.forEach((row, rowIdx) => {
      const current = normalize(String(row[field] ?? ""));
      if (fromVal !== "" && normalizeUpper(current) !== normalizeUpper(fromVal)) return;
      row[field] = toVal;
      greenSet.add(`${rowIdx}-${field}`);
      count++;
    });

    const table = dataset === "master" ? masterTable : matsTable;
    renderTableFiltered(table, dataset);
    refreshSummaries();

    const fromDesc = fromVal !== "" ? `"${fromVal}"` : "cualquier valor";
    log(`Edición en bloque en ${dataset.toUpperCase()}: campo "${HEADER_LABELS[field] ?? field}", ${fromDesc} → "${toVal}". ${count} fila(s) modificada(s) y marcadas en verde.`);
  }

  function refreshSummaries() {
    const m = getStats(state.master);
    const t = getStats(state.mats);
    masterSummary.textContent = `Total MASTER: ${m.total} | Automáticos: ${m.automatico} | Control manual: ${m.manual}`;
    matsSummary.textContent = `Total MAT: ${t.total} | Automáticos: ${t.automatico} | Control manual: ${t.manual}`;
  }

  function refreshTables() {
    renderTableFiltered(masterTable, "master");
    renderTableFiltered(matsTable, "mats");
    refreshSummaries();
    populateFieldSelectors();
  }

  function handleCellFocus(event) {
    const td = event.target;
    if (!(td instanceof HTMLElement)) return;
    if (!td.matches("td[data-editable='1']")) return;
    td.dataset.originalValue = td.textContent;
  }

  function handleCellEdit(event) {
    const td = event.target;
    if (!(td instanceof HTMLElement)) return;
    if (!td.matches("td[data-editable='1']")) return;

    const datasetName = td.dataset.dataset;
    const rowIdx = Number(td.dataset.row);
    const field = td.dataset.field;

    const rows = datasetName === "master" ? state.master : state.mats;
    if (!rows[rowIdx]) return;

    const rawVal = normalize(td.textContent);
    const newVal = field === "amtoss" ? rawVal.replace(/,\s+/g, ",") : rawVal;
    const oldVal = normalize(td.dataset.originalValue ?? "");

    rows[rowIdx][field] = newVal;

    // Solo re-renderizar si el valor cambió (evita destruir el TD justo antes del contextmenu)
    if (newVal !== oldVal) {
      if (datasetName === "master") {
        evaluateMasterRow(rows[rowIdx]);
      } else {
        evaluateMatsRow(rows[rowIdx]);
      }
      refreshTables();
    }
  }

  function excelDateToText(value) {
    if (value instanceof Date && !Number.isNaN(value.valueOf())) {
      const yyyy = value.getFullYear();
      const mm = String(value.getMonth() + 1).padStart(2, "0");
      const dd = String(value.getDate()).padStart(2, "0");
      return `${yyyy}-${mm}-${dd}`;
    }
    return normalize(value);
  }

  function parseWorkpackageRowsFromMatrix(matrix) {
    if (!matrix.length) return [];

    const headers = matrix[0].map((h) => normalizeUpper(h));
    const idxNo = headers.findIndex((h) => h === "NO." || h === "NO" || h === "Nº" || h === "ITEM NO.");
    const idxWo = headers.findIndex((h) => h === "W/O" || h === "WO" || h === "W/O NUMBER");
    const idxAta = headers.findIndex((h) => h === "ATA");
    const idxDesc = headers.findIndex((h) => h.includes("MAINTENANCE EVENT DESCRIPTION") || h.includes("ORIGINATING DATA") || h.includes("DESCRIPTION"));

    if (idxNo === -1 || idxWo === -1 || idxDesc === -1) {
      throw new Error("Formato de Workpackage no reconocido.");
    }

    const rows = [];
    for (let i = 1; i < matrix.length; i += 1) {
      const r = matrix[i];
      const title = normalize(r[idxDesc]);
      if (!title) continue;
      rows.push({
        num: normalize(r[idxNo]),
        wo: normalize(r[idxWo]),
        ata: idxAta > -1 ? excelDateToText(r[idxAta]) : "",
        title: normalizeTitleDisplay(title)
      });
    }
    return rows;
  }

  function parseWorkpackageWorkbook(workbook) {
    const firstSheet = workbook.SheetNames[0];
    const ws = workbook.Sheets[firstSheet];
    const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: "" });
    return parseWorkpackageRowsFromMatrix(matrix);
  }

  function extractWoFromPdfCell(rawWo) {
    const text = normalize(rawWo);
    if (!text) return "";

    const noSeq = text.replace(/SEQ\.?\s*\d+/gi, " ");
    const noApostrophe = noSeq.replace(/'/g, "");
    const match8 = noApostrophe.match(/\b\d{8}\b/);
    if (match8) return match8[0];

    const tokens = noSeq.match(/\b[A-Z]{3}\b/g) || [];
    const candidate = tokens.find((t) => t !== "SEQ" && t !== "ADD");
    if (candidate) return candidate;

    return normalize(noSeq.split(/\s+/)[0]);
  }

  function groupPdfItemsByRow(items, tolerance) {
    const sorted = [...items].sort((a, b) => (Math.abs(b.y - a.y) > 0.01 ? b.y - a.y : a.x - b.x));
    const rows = [];
    for (const it of sorted) {
      const last = rows[rows.length - 1];
      if (!last || Math.abs(last.y - it.y) > tolerance) {
        rows.push({ y: it.y, items: [it] });
      } else {
        last.items.push(it);
      }
    }
    rows.forEach((r) => r.items.sort((a, b) => a.x - b.x));
    return rows;
  }

  function detectPdfHeaderColumns(row) {
    const columns = {};
    for (const it of row.items) {
      const text = normalizeUpper(it.text);
      if (!text) continue;
      if (text.includes("ITEM") || text.includes("NO.")) {
        columns.itemNo = Math.min(columns.itemNo ?? Number.POSITIVE_INFINITY, it.x);
      }
      if (text.includes("W/O") || text.includes("WO NUMBER") || text.includes("W/O NUMBER")) {
        columns.wo = Math.min(columns.wo ?? Number.POSITIVE_INFINITY, it.x);
      }
      if (text.includes("ORIGINATING")) {
        columns.originating = Math.min(columns.originating ?? Number.POSITIVE_INFINITY, it.x);
      }
      if (text.includes("REASON")) {
        columns.reason = Math.min(columns.reason ?? Number.POSITIVE_INFINITY, it.x);
      }
    }
    if (!Number.isFinite(columns.itemNo) || !Number.isFinite(columns.wo) || !Number.isFinite(columns.originating) || !Number.isFinite(columns.reason)) {
      return null;
    }
    return columns;
  }

  function buildColumnBoundaries(columns) {
    const ordered = [
      { key: "itemNo", x: columns.itemNo },
      { key: "wo", x: columns.wo },
      { key: "originating", x: columns.originating },
      { key: "reason", x: columns.reason }
    ].sort((a, b) => a.x - b.x);

    const boundaries = [];
    for (let i = 0; i < ordered.length - 1; i += 1) {
      boundaries.push((ordered[i].x + ordered[i + 1].x) / 2);
    }
    return { ordered, boundaries };
  }

  function assignItemToColumn(x, ordered, boundaries) {
    for (let i = 0; i < boundaries.length; i += 1) {
      if (x < boundaries[i]) return ordered[i].key;
    }
    return ordered[ordered.length - 1].key;
  }

  function parseWorkpackagePdfRows(rawRows) {
    const parsed = [];
    let current = null;

    for (const row of rawRows) {
      const itemNoRaw = normalize(row.itemNo);
      const itemMatch = itemNoRaw.match(/^(\d+)\.?$/);
      const startsNew = Boolean(itemMatch);
      if (startsNew) {
        if (current) parsed.push(current);
        current = {
          num: itemMatch[1],
          woRaw: normalize(row.wo),
          title: normalize(row.originating),
          reason: normalize(row.reason)
        };
        continue;
      }
      if (!current) continue;

      if (normalize(row.wo)) current.woRaw = `${current.woRaw} ${normalize(row.wo)}`.trim();
      if (normalize(row.originating)) current.title = `${current.title} ${normalize(row.originating)}`.trim();
      if (normalize(row.reason)) current.reason = `${current.reason} ${normalize(row.reason)}`.trim();
    }

    if (current) parsed.push(current);

    return parsed
      .filter((r) => normalizeUpper(`${r.reason} ${r.title} ${r.woRaw}`).includes("ADD"))
      .map((r) => ({
        num: r.num,
        wo: extractWoFromPdfCell(r.woRaw),
        ata: "",
        title: normalize(r.title)
      }))
      .filter((r) => r.num && r.title);
  }

  function isPdfNoiseLine(lineUpper) {
    if (!lineUpper) return true;
    if (lineUpper.includes("BASE MAINTENANCE CHANGES")) return true;
    if (lineUpper.includes("CONTROL DOCUMENT")) return true;
    if (lineUpper.includes("REGISTRATION:")) return true;
    if (lineUpper.includes("CHECK REFERENCE")) return true;
    if (lineUpper.includes("REVISION DATE")) return true;
    if (lineUpper.includes("PAGE NO")) return true;
    if (lineUpper.includes("ITEM NO.") && lineUpper.includes("ORIGINATING DATA")) return true;
    if (lineUpper.includes("W/O NUMBER") && lineUpper.includes("REASON")) return true;
    if (lineUpper.includes("SEE REV. NO")) return true;
    if (lineUpper.includes("AER LINGUS")) return true;
    if (lineUpper === "NO. REV. NO") return true;
    if (/^\d+\s+OF\s+\d+$/i.test(lineUpper)) return true;
    if (/^ISSUE\s+\d+$/i.test(lineUpper)) return true;
    if (lineUpper === "AIRCRAFT") return true;
    if (lineUpper === "ITEM SEE") return true;
    if (/\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\b/.test(lineUpper) && /\d{4}/.test(lineUpper) && lineUpper.includes(" OF ")) return true;
    if (lineUpper.includes("ACKNOWLEDGEMENT")) return true;
    if (lineUpper.startsWith("ALL THE AMENDMENTS")) return true;
    if (lineUpper.includes("___")) return true;
    return false;
  }

  function parseWorkpackagePdfFromLines(lines) {
    const cleanedLines = lines
      .map((rawLine) => normalize(rawLine).replace(/\s+/g, " "))
      .filter((line) => line && !isPdfNoiseLine(normalizeUpper(line)) && !/^SEQ\.?\s*\d+/i.test(normalizeUpper(line)));

    const entries = cleanedLines.map((line) => {
      const itemMatch = line.match(/^(\d+)\.\s+(.+)$/);
      if (!itemMatch) return { kind: "text", line };

      const num = itemMatch[1];
      const rest = normalize(itemMatch[2]);
      let woRaw = rest;
      let reason = "";
      const dateMatch = rest.match(/\b\d{2}\/\d{2}\/\d{4}\b/);
      if (dateMatch && typeof dateMatch.index === "number") {
        woRaw = normalize(rest.slice(0, dateMatch.index));
        reason = normalize(rest.slice(dateMatch.index + dateMatch[0].length));
      } else {
        const reasonMatch = rest.match(/\b(ADD|CANX)\b/i);
        if (reasonMatch && typeof reasonMatch.index === "number") {
          woRaw = normalize(rest.slice(0, reasonMatch.index));
          reason = normalize(rest.slice(reasonMatch.index));
        }
      }
      return { kind: "item", num, woRaw, reason, prefix: [], suffix: [] };
    });

    const itemPositions = [];
    for (let i = 0; i < entries.length; i += 1) {
      if (entries[i].kind === "item") itemPositions.push(i);
    }
    if (!itemPositions.length) return [];

    // Assign lines before the first item as its prefix (captures title text that appears above the item row)
    const firstPos = itemPositions[0];
    const beforeFirstLines = entries.slice(0, firstPos).filter((e) => e.kind === "text").map((e) => e.line);
    if (beforeFirstLines.length) {
      entries[firstPos].prefix.push(...beforeFirstLines);
    }

    for (let i = 0; i < itemPositions.length - 1; i += 1) {
      const currPos = itemPositions[i];
      const nextPos = itemPositions[i + 1];
      const gapLines = entries.slice(currPos + 1, nextPos).filter((e) => e.kind === "text").map((e) => e.line);
      if (!gapLines.length) continue;
      const split = Math.ceil(gapLines.length / 2);
      entries[currPos].suffix.push(...gapLines.slice(0, split));
      entries[nextPos].prefix.push(...gapLines.slice(split));
    }

    // Assign lines after the last item as its suffix (captures title text that appears below the item row)
    const lastPos = itemPositions[itemPositions.length - 1];
    const afterLastLines = entries.slice(lastPos + 1).filter((e) => e.kind === "text").map((e) => e.line);
    if (afterLastLines.length) {
      entries[lastPos].suffix.push(...afterLastLines.slice(0, Math.ceil(afterLastLines.length / 2)));
    }

    return itemPositions
      .map((pos) => {
        const e = entries[pos];
        return {
          num: e.num,
          woRaw: e.woRaw,
          reason: e.reason,
          title: normalize([...e.prefix, ...e.suffix].join(" "))
        };
      })
      .filter((r) => normalizeUpper(r.reason).includes("ADD"))
      .map((r) => ({
        num: r.num,
        wo: extractWoFromPdfCell(r.woRaw),
        ata: "",
        title: normalizeTitleDisplay(r.title)
      }))
      .filter((r) => r.num && r.title);
  }

  async function parseWorkpackagePdf(file) {
    if (typeof pdfjsLib === "undefined") {
      throw new Error("No se pudo cargar la libreria PDF.");
    }

    if (!pdfjsLib.GlobalWorkerOptions.workerSrc) {
      pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
    }

    const bytes = new Uint8Array(await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = () => reject(new Error("No se pudo leer el fichero PDF."));
      reader.readAsArrayBuffer(file);
    }));
    const doc = await pdfjsLib.getDocument({ data: bytes }).promise;
    const rawRows = [];
    const parsedFromLines = [];

    for (let pageNum = 1; pageNum <= doc.numPages; pageNum += 1) {
      const page = await doc.getPage(pageNum);
      const text = await page.getTextContent();
      const items = text.items
        .map((it) => ({
          text: normalize(it.str),
          x: it.transform[4],
          y: it.transform[5]
        }))
        .filter((it) => it.text);

      if (!items.length) continue;

      const rows = groupPdfItemsByRow(items, 2.5);
      const pageLineTexts = [];
      rows.forEach((row) => {
        const lineText = row.items.map((x) => normalize(x.text)).join(" ").replace(/\s+/g, " ").trim();
        if (lineText) pageLineTexts.push(lineText);
      });
      const pageParsed = parseWorkpackagePdfFromLines(pageLineTexts);
      if (pageParsed.length) parsedFromLines.push(...pageParsed);

      let headerIdx = -1;
      let headerColumns = null;
      for (let i = 0; i < rows.length; i += 1) {
        const joined = rows[i].items.map((x) => normalizeUpper(x.text)).join(" ");
        if (joined.includes("ITEM") && joined.includes("REASON") && joined.includes("ORIGINATING")) {
          const detected = detectPdfHeaderColumns(rows[i]);
          if (detected) {
            headerIdx = i;
            headerColumns = detected;
            break;
          }
        }
      }
      if (headerIdx === -1 || !headerColumns) continue;

      const { ordered, boundaries } = buildColumnBoundaries(headerColumns);
      for (let i = headerIdx + 1; i < rows.length; i += 1) {
        const line = { itemNo: "", wo: "", originating: "", reason: "" };
        for (const it of rows[i].items) {
          const key = assignItemToColumn(it.x, ordered, boundaries);
          line[key] = `${line[key]} ${it.text}`.trim();
        }
        if (!line.itemNo && !line.wo && !line.originating && !line.reason) continue;
        rawRows.push(line);
      }
    }

    let parsed = parsedFromLines;
    if (!parsed.length) {
      parsed = parseWorkpackagePdfRows(rawRows);
    }
    if (!parsed.length) {
      throw new Error("No se detectaron registros validos en PDF (Reason debe contener ADD).");
    }
    return parsed;
  }

  function parseBaseEinWorkbook(workbook) {
    const sheetName = workbook.SheetNames.includes("BASE EIN") ? "BASE EIN" : workbook.SheetNames[0];
    const ws = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: "" });
    if (matrix.length < 2) return [];

    const rows = [];
    for (let i = 1; i < matrix.length; i += 1) {
      const r = matrix[i];
      if (r.every((v) => normalize(v) === "")) continue;
      rows.push({
        criterio: normalize(r[0]),
        title: normalize(r[1]),
        team: normalize(r[2]),
        amtoss: normalize(r[3]),
        remarks: normalize(r[4]),
        anexo: normalize(r[5]),
        irCt: normalize(r[6]),
        matSiNo: normalize(r[7]),
        pn: normalize(r[8]),
        descripcion: normalize(r[9]),
        qty: normalize(r[10]),
        type: normalize(r[11]),
        notas: normalize(r[12])
      });
    }
    return rows;
  }

  async function parseWorkbookFromFile(file) {
    const buffer = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = () => reject(new Error("No se pudo leer el fichero. Comprueba que no está abierto en Excel u otro programa."));
      reader.readAsArrayBuffer(file);
    });
    return XLSX.read(buffer, { type: "array", cellDates: true });
  }

  // ═══════════════════════════════════════════════════════════
  //  VLG — Parser de Workpackage
  //  Columnas: A=Seq.No., B=Work Step/Event Number, C=Title
  //  CRITERIO = parte del título antes del primer "/"
  // ═══════════════════════════════════════════════════════════

  function parseWorkpackageVlgFromMatrix(matrix) {
    if (!matrix.length) return [];

    // Buscar la fila de cabecera: puede estar en cualquier posición
    // Se detecta buscando la fila que contenga "SEQ" y ("WORK STEP" o "EVENT NUMBER") y "TITLE"
    let headerRowIdx = -1;
    let idxNo = -1, idxWo = -1, idxTitle = -1;

    for (let i = 0; i < matrix.length; i++) {
      const headers = matrix[i].map((h) => normalizeUpper(String(h)));
      const ni = headers.findIndex((h) => h.includes("SEQ") || h === "NO." || h === "NO");
      const wi = headers.findIndex((h) => h.includes("WORK STEP") || h.includes("EVENT NUMBER") || h.includes("W/O"));
      const ti = headers.findIndex((h) => h.includes("TITLE") || h.includes("DESCRIPTION"));
      if (ni !== -1 && wi !== -1 && ti !== -1) {
        headerRowIdx = i;
        idxNo = ni; idxWo = wi; idxTitle = ti;
        break;
      }
    }

    if (headerRowIdx === -1) {
      throw new Error("Formato de Workpackage VLG no reconocido. Se esperan columnas Seq.No., Work Step/Event Number y Title.");
    }

    const rows = [];
    for (let i = headerRowIdx + 1; i < matrix.length; i++) {
      const r = matrix[i];
      const title = normalize(r[idxTitle]);
      if (!title) continue;
      rows.push({
        num: normalize(r[idxNo]),
        wo: normalize(r[idxWo]),
        ata: "",
        title: normalizeTitleDisplay(title)
      });
    }
    return rows;
  }

  function parseWorkpackageVlgWorkbook(workbook) {
    // VLG Workpackage puede estar en cualquier hoja; preferimos la que contenga "TASK"
    const sheetName = workbook.SheetNames.find((n) => normalizeUpper(n).includes("TASK"))
      ?? workbook.SheetNames.find((n) => normalizeUpper(n).includes("WORKPACKAGE"))
      ?? workbook.SheetNames[0];
    const ws = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    return parseWorkpackageVlgFromMatrix(matrix);
  }

  // ═══════════════════════════════════════════════════════════
  //  VLG — Parser de BASE VLG
  //  Misma estructura de columnas que BASE EIN (A=BÚSQUEDA … M=NOTAS)
  // ═══════════════════════════════════════════════════════════

  function parseBaseVlgWorkbook(workbook) {
    const sheetName = workbook.SheetNames.find((n) => normalizeUpper(n).includes("BASE VLG"))
      ?? workbook.SheetNames[0];
    const ws = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: "" });
    if (matrix.length < 2) return [];

    const rows = [];
    for (let i = 1; i < matrix.length; i++) {
      const r = matrix[i];
      if (r.every((v) => normalize(v) === "")) continue;
      rows.push({
        criterio:    normalize(r[0]),
        title:       normalize(r[1]),
        team:        normalize(r[2]),
        amtoss:      normalize(r[3]),
        remarks:     normalize(r[4]),
        anexo:       normalize(r[5]),
        irCt:        normalize(r[6]),
        matSiNo:     normalize(r[7]),
        pn:          normalize(r[8]),
        descripcion: normalize(r[9]),
        qty:         normalize(r[10]),
        type:        normalize(r[11]),
        notas:       normalize(r[12])
      });
    }
    return rows;
  }

  // ═══════════════════════════════════════════════════════════
  //  VLG — Extracción de CRITERIO desde título
  //  El criterio es la parte antes del primer "/"
  // ═══════════════════════════════════════════════════════════

  function extractVlgCriterio(title) {
    const t = normalize(title);
    const slashIdx = t.indexOf("/");
    if (slashIdx > 0) return t.slice(0, slashIdx).trim();
    return t;
  }

  // ═══════════════════════════════════════════════════════════
  //  VLG — Algoritmo de matching contra BASE VLG
  //  Basado en ExtraerTodasReferencias + EsPatronAlfanumericoConGuiones
  //  Selecciona la coincidencia más larga con prioridad para patrones específicos
  // ═══════════════════════════════════════════════════════════

  function isAlphanumericWithDashes(value) {
    // Patrón alfanumérico con guiones: letras y dígitos separados por guiones
    return /^[A-Z0-9]+(-[A-Z0-9]+)+$/i.test(normalize(value));
  }

  function isSpecificVlgPattern(value) {
    const v = normalizeUpper(value);
    // Patrón específico: longitud >= 8 y contiene letras y números
    return v.length >= 8 && /[A-Z]/.test(v) && /[0-9]/.test(v);
  }

  function extractAllReferences(text) {
    // Extrae todas las posibles referencias del texto (tokens alfanuméricos con guiones)
    const t = normalizeUpper(text);
    const tokens = t.match(/[A-Z0-9]+(?:-[A-Z0-9]+)+/g) || [];
    // Añadir también el texto completo antes del primer espacio como candidato
    const firstToken = t.split(/\s+/)[0];
    if (firstToken && !tokens.includes(firstToken)) tokens.unshift(firstToken);
    return [...new Set(tokens)];
  }

  function pickBestBaseMatchVlg(title, baseVlg) {
    // Replica exacta de la macro VLG: InStr(criterio_wp, criterio_bd) > 0
    // criterio_wp = parte antes del primer "/" (o título completo si no hay "/")
    const criterio = extractVlgCriterio(title);
    const criterioUp = normalizeUpper(criterio);

    let bestIndex = -1;
    let bestScore = { priority: false, length: 0 };

    for (let i = 0; i < baseVlg.length; i++) {
      const key = normalizeUpper(normalize(baseVlg[i].criterio));
      if (!key || key.length < 3) continue;

      // InStr(criterio_wp, criterio_bd): criterio de BD contenido en criterio del WP
      if (!criterioUp.includes(key)) continue;

      const isPriority = isSpecificVlgPattern(key);
      const keyLen = key.length;

      if (!bestScore.priority && isPriority) {
        bestIndex = i;
        bestScore = { priority: true, length: keyLen };
      } else if (isPriority === bestScore.priority && keyLen > bestScore.length) {
        bestIndex = i;
        bestScore = { priority: isPriority, length: keyLen };
      }
    }

    return bestIndex;
  }

  // ═══════════════════════════════════════════════════════════
  //  VLG — buildMasterVlg
  // ═══════════════════════════════════════════════════════════

  function buildMasterVlg(workpackage, baseVlg) {
    const rows = [];

    for (let i = 0; i < workpackage.length; i++) {
      const wp = workpackage[i];
      const title = normalize(wp.title);
      if (!title) continue;

      const wo = normalize(wp.wo);
      const matchIndex = pickBestBaseMatchVlg(title, baseVlg);

      if (matchIndex === -1) {
        rows.push({
          wpIndex: i,
          num: normalize(wp.num),
          wo,
          title: title.slice(0, 200),
          team: "",
          amtoss: "CONTROL MANUAL",
          remarks: "",
          anexo: "",
          irCt: "",
          matSiNo: "",
          criterioBusqueda: "SIN CRITERIO",
          status: "CONTROL MANUAL",
          manualReasons: ["Sin coincidencia de criterio en Base de datos"]
        });
        continue;
      }

      const base = baseVlg[matchIndex];
      const row = {
        wpIndex: i,
        num: normalize(wp.num),
        wo,
        title: title.slice(0, 200),
        team: normalize(base.team),
        amtoss: normalize(base.amtoss),
        remarks: normalize(base.remarks),
        anexo: normalize(base.anexo),
        irCt: normalize(base.irCt),
        matSiNo: normalize(base.matSiNo),
        criterioBusqueda: normalize(base.criterio),
        status: "OK",
        manualReasons: []
      };

      // Reusar evaluateMasterRow (misma lógica de campos obligatorios y keywords)
      evaluateMasterRow(row);
      rows.push(row);
    }

    return rows;
  }

  // ═══════════════════════════════════════════════════════════
  //  VLG — buildMatsVlg
  //  Sin condición P/N en el matching (diferencia clave con EIN)
  // ═══════════════════════════════════════════════════════════

  // Replica exacta de ExtraerPatronReferencia VBA (solo para VLG MAT)
  // Pre-procesa el título del WP para extraer el token de referencia antes de buscar en BD
  function extractPatronReferencia(texto) {
    const texto2 = (texto || "").trim();
    if (!texto2) return "";
    const palabras = texto2.split(" ");
    for (const palabra of palabras) {
      const p = palabra.trim();
      if (p.length < 11) continue;
      if (!/^[A-Za-z]/.test(p)) continue;
      if (!p.includes("-")) continue;
      if ((p.match(/-/g) || []).length < 2) continue;
      if (p.includes(" ")) continue;
      const partes = p.split("-");
      if (partes.length < 3) continue;
      const parte0 = partes[0];
      if (!parte0[0].match(/[A-Za-z]/) || parte0.length < 2) continue;
      if (/[0-9]/.test(parte0.slice(1))) return p.toUpperCase();
    }
    return "";
  }

  function buildMatsVlg(workpackage, baseVlg, masterRows) {
    const rows = [];
    const matEligible = normalizeUpper(CONFIG.matEligibleValue);
    const masterByWp = new Map();
    masterRows.forEach((m) => { if (!masterByWp.has(m.wpIndex)) masterByWp.set(m.wpIndex, m); });

    for (let i = 0; i < workpackage.length; i++) {
      const wp = workpackage[i];
      const title = normalize(wp.title);
      if (!title) continue;

      const wo = normalize(wp.wo);
      // Replicar ExtraerPatronReferencia VBA: buscar en BD con el patrón extraído si existe
      const patron = extractPatronReferencia(wp.title);
      const titleForSearch = patron ? patron : title;
      const matchIndex = pickBestBaseMatchVlg(titleForSearch, baseVlg);

      if (matchIndex === -1) {
        rows.push({
          wpIndex: i,
          num: normalize(wp.num),
          wo,
          title: title.slice(0, 200),
          team: "CONTROL MANUAL",
          pn: "",
          descripcion: "",
          qty: "",
          type: "",
          notas: "",
          matSiNo: "",
          criterioBusqueda: "SIN CRITERIO",
          status: "CONTROL MANUAL",
          manualReasons: ["Sin coincidencia de criterio en Base de datos"]
        });
        continue;
      }

      // VLG: matching solo por criterio (sin condición P/N)
      const selectedKey = normalizeUpper(baseVlg[matchIndex].criterio);
      const candidates = baseVlg.filter((r) => normalizeUpper(r.criterio) === selectedKey);
      const withMat = candidates.filter((r) => normalizeUpper(r.matSiNo) === matEligible);
      if (withMat.length === 0) continue;

      for (const base of withMat) {
        const masterRow = masterByWp.get(i);
        const row = {
          wpIndex: i,
          num: normalize(wp.num),
          wo,
          title: title.slice(0, 200),
          team: normalize(masterRow ? masterRow.team : base.team),
          pn: normalize(base.pn),
          descripcion: normalize(base.descripcion),
          qty: normalize(base.qty),
          type: normalize(base.type),
          notas: normalize(base.notas),
          matSiNo: normalize(base.matSiNo),
          criterioBusqueda: normalize(base.criterio),
          status: "OK",
          manualReasons: []
        };
        evaluateMatsRow(row);
        rows.push(row);
      }
    }

    return rows;
  }

  async function loadBaseEinFromProject() {
    const url = encodeURI("../Base de datos EIN.xlsx");
    const res = await fetch(url, { cache: "no-store" });
    if (!res.ok) throw new Error("No se pudo abrir 'Base de datos EIN.xlsx' desde el proyecto.");
    const buffer = await res.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array", cellDates: true });
    state.baseEinWorkbook = wb;
    return parseBaseEinWorkbook(wb);
  }

  async function exportRowsToXlsx(rows, sheetName, fileName, greenSet) {
    if (!rows.length) {
      throw new Error(`No hay datos para exportar ${sheetName}.`);
    }

    const HEADER_LABELS = {
      num: "Nº", wo: "WO", title: "TITLE", team: "TEAM",
      amtoss: "AMTOSS", remarks: "REMARKS", anexo: "ANEXO",
      irCt: "IR/CT", matSiNo: "MAT SI/NO",
      criterioBusqueda: "CRITERIO BÚSQUEDA", status: "STATUS",
      manualReasons: "RAZONES CONTROL MANUAL",
      pn: "PN", descripcion: "DESCRIPCIÓN", qty: "QTY",
      type: "TYPE", notas: "NOTAS"
    };

    // Anchos en caracteres aproximados (como en la tabla visual)
    const COL_CHAR_WIDTHS = {
      num: 5, wo: 14, title: 28, team: 7,
      amtoss: 16, remarks: 20, anexo: 10,
      irCt: 7, matSiNo: 10,
      criterioBusqueda: 22, status: 14,
      manualReasons: 30,
      pn: 16, descripcion: 22, qty: 5,
      type: 8, notas: 18
    };

    const fields = Object.keys(rows[0]).filter((h) => h !== "wpIndex");

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(sheetName);

    // Anchos de columna
    ws.columns = fields.map((f) => ({ width: COL_CHAR_WIDTHS[f] ?? 14 }));

    // Fila de encabezado
    const headerRow = ws.addRow(fields.map((f) => HEADER_LABELS[f] ?? f));
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, size: 9, color: { argb: "FF102A2D" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE6F2F4" } };
      cell.border = {
        bottom: { style: "thin", color: { argb: "FFD8E3E6" } }
      };
      cell.alignment = { vertical: "middle", wrapText: false };
    });
    headerRow.height = 16;

    // Filas de datos
    rows.forEach((row, rowIdx) => {
      const values = fields.map((f) => {
        if (f === "manualReasons") return (row.manualReasons || []).join(" | ");
        return row[f] ?? "";
      });
      const dataRow = ws.addRow(values);
      dataRow.height = 14;

      const isManual = row.status === "CONTROL MANUAL";

      dataRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const field = fields[colNumber - 1];
        const key = `${rowIdx}-${field}`;
        const isGreen = greenSet && greenSet.has(key);

        cell.font = { size: 9 };
        cell.alignment = { vertical: "middle", wrapText: false };
        cell.border = {
          bottom: { style: "hair", color: { argb: "FFD8E3E6" } }
        };

        if (isGreen) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD4F7D4" } };
        } else if (isManual) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF3F3" } };
        }

        // Badge STATUS
        if (field === "status") {
          if (row.status === "OK") {
            cell.font = { size: 9, bold: true, color: { argb: "FF2D6A4F" } };
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDFF5E8" } };
          } else if (row.status === "CONTROL MANUAL") {
            cell.font = { size: 9, bold: true, color: { argb: "FFC1121F" } };
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFE2E5" } };
          }
        }
      });
    });

    // Descargar
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    a.click();
    URL.revokeObjectURL(url);
  }

  function getCliente() {
    return document.getElementById("clienteSelect").value; // "EIN" | "VLG"
  }

  async function loadWorkpackage(file, cliente) {
    if (typeof XLSX === "undefined") throw new Error("No se pudo cargar el parser de Excel (SheetJS).");
    const rawName = normalize(file.name);
    const ext = rawName.includes(".") ? rawName.slice(rawName.lastIndexOf(".")).toLowerCase() : "";
    const isPdf = ext === ".pdf" || normalizeUpper(file.type).includes("PDF");

    if (cliente === "VLG") {
      if (isPdf) throw new Error("El Workpackage VLG debe ser un fichero Excel (.xlsx), no PDF.");
      log("Leyendo Workpackage-Summary VLG...");
      const wpWb = await parseWorkbookFromFile(file);
      state.workpackage = parseWorkpackageVlgWorkbook(wpWb);
      log(`Workpackage VLG cargado: ${state.workpackage.length} filas.`);
    } else {
      if (isPdf) {
        log("Leyendo Workpackage-Summary.pdf...");
        state.workpackage = await parseWorkpackagePdf(file);
        log(`Workpackage PDF cargado: ${state.workpackage.length} filas utiles (Reason contiene ADD).`);
      } else {
        log("Leyendo Workpackage-Summary.xlsx...");
        const wpWb = await parseWorkbookFromFile(file);
        state.workpackage = parseWorkpackageWorkbook(wpWb);
        log(`Workpackage XLSX cargado: ${state.workpackage.length} filas utiles.`);
      }
    }
  }

  async function handleExecuteMaster() {
    clearLog();
    const cliente = getCliente();
    const input = document.getElementById("workpackageFile");
    const file = input.files && input.files[0];
    if (!file) throw new Error("Selecciona un fichero Workpackage-Summary.");

    const baseInput = document.getElementById("baseEinFile");
    const baseFile = baseInput.files && baseInput.files[0];
    if (!baseFile) throw new Error(`Selecciona el fichero Base de datos ${cliente}.xlsx.`);

    await loadWorkpackage(file, cliente);

    log(`Leyendo Base de datos ${cliente}.xlsx...`);
    const baseWb = await parseWorkbookFromFile(baseFile);
    state.baseEinWorkbook = baseWb;

    if (cliente === "VLG") {
      state.baseEin = parseBaseVlgWorkbook(baseWb);
      log(`BASE VLG cargada: ${state.baseEin.length} filas.`);
      state.master = buildMasterVlg(state.workpackage, state.baseEin);
    } else {
      state.baseEin = parseBaseEinWorkbook(baseWb);
      log(`BASE EIN cargada: ${state.baseEin.length} filas.`);
      state.master = buildMaster(state.workpackage, state.baseEin);
    }

    state.greenEdits.master = new Set();
    activeFilters.master = null;
    renderTableFiltered(masterTable, "master");
    refreshSummaries();
    populateFieldSelectors();
    document.getElementById("execMatsBtn").disabled = false;

    log(`MASTER generado: ${state.master.length} filas.`);
    log(`PASO 2 completado. Revisa el MASTER y ejecuta MAT.`);
  }

  async function handleExecuteMats() {
    const cliente = getCliente();
    if (!state.master.length) throw new Error("Ejecuta primero MASTER.");

    if (cliente === "VLG") {
      state.mats = buildMatsVlg(state.workpackage, state.baseEin, state.master);
    } else {
      state.mats = buildMats(state.workpackage, state.baseEin, state.master);
    }

    state.greenEdits.mats = new Set();
    activeFilters.mats = null;
    renderTableFiltered(matsTable, "mats");
    refreshSummaries();
    populateFieldSelectors();

    log(`MAT generado: ${state.mats.length} filas.`);
    log("PASO 3 completado. Revisa los resultados y edita celdas si es necesario.");
  }

  async function handleUpdateBaseEin(dataset) {
    const cliente = getCliente();
    const greenSet = state.greenEdits[dataset];
    const baseLabel = cliente === "VLG" ? "BASE VLG" : "BASE EIN";

    if (!greenSet.size) {
      log(`No hay celdas verdes en ${dataset.toUpperCase()} para actualizar ${baseLabel}.`);
      return;
    }
    if (!state.baseEinWorkbook) {
      throw new Error(`${baseLabel} no cargada. Ejecuta primero MASTER.`);
    }

    // Columnas (idénticas en BASE EIN y BASE VLG)
    // MASTER: TEAM=2, AMTOSS=3, REMARKS=4, ANEXO=5, IR/CT=6, MAT SI/NO=7
    // MAT:    TEAM=2, MAT SI/NO=7, P/N=8, DESC=9, QTY=10, TYPE=11, NOTAS=12
    const FIELD_TO_COL = {
      criterioBusqueda: 0,
      team: 2, amtoss: 3, remarks: 4, anexo: 5, irCt: 6, matSiNo: 7,
      pn: 8, descripcion: 9, qty: 10, type: 11, notas: 12
    };

    const rows = dataset === "master" ? state.master : state.mats;
    const wb = state.baseEinWorkbook;
    const sheetName = wb.SheetNames.find((n) => normalizeUpper(n).includes(baseLabel))
      ?? (wb.SheetNames.includes("BASE EIN") ? "BASE EIN" : wb.SheetNames[0]);
    const ws = wb.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

    const rowsToSync = new Map();
    for (const key of greenSet) {
      const dashIdx = key.indexOf("-");
      const rowIdx = Number(key.slice(0, dashIdx));
      const field = key.slice(dashIdx + 1);
      if (FIELD_TO_COL[field] === undefined) continue;
      if (!rowsToSync.has(rowIdx)) rowsToSync.set(rowIdx, new Set());
      rowsToSync.get(rowIdx).add(field);
    }

    let nActualizadas = 0, nCreadas = 0, nSaltadas = 0;

    for (const [rowIdx, fields] of rowsToSync) {
      const appRow = rows[rowIdx];
      if (!appRow) continue;

      const criterio = normalizeUpper(appRow.criterioBusqueda);
      if (!criterio || criterio === "SIN CRITERIO") { nSaltadas++; continue; }

      const matchingRows = [];
      for (let i = 1; i < matrix.length; i++) {
        if (normalizeUpper(normalize(String(matrix[i][0] ?? ""))) !== criterio) continue;
        // EIN MAT: condición adicional de P/N. VLG MAT: solo criterio.
        if (cliente === "EIN" && dataset === "mats") {
          const basePn = normalizeUpper(normalize(String(matrix[i][8] ?? "")));
          const appPn = normalizeUpper(normalize(appRow.pn ?? ""));
          if (appPn && basePn !== appPn) continue;
        }
        matchingRows.push(i);
      }

      if (matchingRows.length > 0) {
        // Coincidencia encontrada: sobrescribir en todas las filas coincidentes
        // No escribir criterioBusqueda (col 0) en actualizaciones — solo en creaciones
        for (const i of matchingRows) {
          for (const field of fields) {
            if (field === "criterioBusqueda") continue;
            matrix[i][FIELD_TO_COL[field]] = appRow[field] ?? "";
          }
        }
        nActualizadas++;
      } else {
        // Sin coincidencia: crear fila nueva solo si criterioBusqueda está marcado en verde
        if (!fields.has("criterioBusqueda")) { nSaltadas++; continue; }
        const newRow = Array(13).fill("");
        for (const field of fields) {
          newRow[FIELD_TO_COL[field]] = appRow[field] ?? "";
        }
        matrix.push(newRow);
        nCreadas++;
      }
    }

    if (nActualizadas === 0 && nCreadas === 0) {
      log(`No se encontraron coincidencias en ${baseLabel}. Saltadas: ${nSaltadas}.`);
      return;
    }

    const newWs = XLSX.utils.aoa_to_sheet(matrix);
    wb.Sheets[sheetName] = newWs;
    XLSX.writeFile(wb, `Base de datos ${cliente}.xlsx`);
    log(`${baseLabel} actualizada. Filas actualizadas: ${nActualizadas} | Creadas: ${nCreadas} | Saltadas: ${nSaltadas}. Sustituye el fichero original con el descargado.`);
  }

  document.getElementById("execMasterBtn").addEventListener("click", async () => {
    try {
      await handleExecuteMaster();
    } catch (error) {
      log(`ERROR: ${error.message}`);
      console.error(error);
    }
  });

  document.getElementById("execMatsBtn").addEventListener("click", async () => {
    try {
      await handleExecuteMats();
    } catch (error) {
      log(`ERROR: ${error.message}`);
      console.error(error);
    }
  });

  document.getElementById("updateMasterBaseBtn").addEventListener("click", async () => {
    try {
      await handleUpdateBaseEin("master");
    } catch (error) {
      log(`ERROR: ${error.message}`);
      console.error(error);
    }
  });

  document.getElementById("updateMatsBaseBtn").addEventListener("click", async () => {
    try {
      await handleUpdateBaseEin("mats");
    } catch (error) {
      log(`ERROR: ${error.message}`);
      console.error(error);
    }
  });

  document.getElementById("downloadMasterXlsxBtn").addEventListener("click", async () => {
    try {
      await exportRowsToXlsx(state.master, "MASTER", "MASTER-output.xlsx", state.greenEdits.master);
    } catch (error) {
      log(`ERROR: ${error.message}`);
    }
  });

  document.getElementById("downloadMatsXlsxBtn").addEventListener("click", async () => {
    try {
      await exportRowsToXlsx(state.mats, "MATS", "MATS-output.xlsx", state.greenEdits.mats);
    } catch (error) {
      log(`ERROR: ${error.message}`);
    }
  });

  // Menú contextual para marcar/desmarcar celda en verde (clic derecho)
  let contextMenu = null;

  function removeContextMenu() {
    if (contextMenu) { contextMenu.remove(); contextMenu = null; }
  }

  function handleCellContextMenu(event) {
    const td = event.target;
    if (!(td instanceof HTMLElement)) return;
    if (!td.matches("td[data-editable='1']")) return;

    event.preventDefault();
    removeContextMenu();

    const datasetName = td.dataset.dataset;
    const rowIdx = td.dataset.row;
    const field = td.dataset.field;
    if (!datasetName || rowIdx === undefined || !field) return;

    const isGreen = td.classList.contains("green-marked");
    const menu = document.createElement("div");
    menu.style.cssText = `position:fixed;z-index:9999;background:#fff;border:1px solid #ccc;border-radius:4px;box-shadow:2px 2px 6px rgba(0,0,0,0.2);padding:4px 0;font-size:13px;`;
    menu.style.left = `${event.clientX}px`;
    menu.style.top = `${event.clientY}px`;

    const option = document.createElement("div");
    option.textContent = isGreen ? "Quitar marca (no actualizar BD)" : "Marcar para actualizar BD";
    option.style.cssText = "padding:6px 16px;cursor:pointer;white-space:nowrap;";
    option.addEventListener("mouseenter", () => { option.style.background = "#f0f0f0"; });
    option.addEventListener("mouseleave", () => { option.style.background = ""; });
    option.addEventListener("click", () => {
      if (isGreen) {
        state.greenEdits[datasetName].delete(`${rowIdx}-${field}`);
        td.classList.remove("green-marked");
      } else {
        state.greenEdits[datasetName].add(`${rowIdx}-${field}`);
        td.classList.add("green-marked");
      }
      removeContextMenu();
    });

    menu.appendChild(option);
    document.body.appendChild(menu);
    contextMenu = menu;
  }

  document.addEventListener("mousedown", (e) => { if (contextMenu && !contextMenu.contains(e.target)) removeContextMenu(); });
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") removeContextMenu(); });

  masterTable.addEventListener("focus", handleCellFocus, true);
  matsTable.addEventListener("focus", handleCellFocus, true);
  masterTable.addEventListener("blur", handleCellEdit, true);
  matsTable.addEventListener("blur", handleCellEdit, true);
  masterTable.addEventListener("contextmenu", handleCellContextMenu);
  matsTable.addEventListener("contextmenu", handleCellContextMenu);

  // Actualizar etiqueta de base de datos al cambiar de cliente
  document.getElementById("clienteSelect").addEventListener("change", () => {
    const cliente = getCliente();
    document.getElementById("baseLabel").textContent = `Base de datos ${cliente} (.xlsx)`;
    // Limpiar fichero seleccionado y estado al cambiar cliente
    document.getElementById("baseEinFile").value = "";
    document.getElementById("workpackageFile").value = "";
    state.workpackage = [];
    state.baseEin = [];
    state.baseEinWorkbook = null;
    state.master = [];
    state.mats = [];
    state.greenEdits = { master: new Set(), mats: new Set() };
    masterTable.innerHTML = "";
    matsTable.innerHTML = "";
    clearLog();
    log(`Cliente cambiado a ${cliente}. Selecciona el Workpackage-Summary y la Base de datos ${cliente}.`);
  });

  // Edición en bloque — listeners por tabla
  document.getElementById("masterBulkApplyBtn").addEventListener("click", () => applyBulkEdit("master"));
  document.getElementById("matsBulkApplyBtn").addEventListener("click", () => applyBulkEdit("mats"));

  clearLog();
  log("App lista. Selecciona el cliente, el Workpackage-Summary y la Base de datos, luego pulsa 'Ejecutar MASTER'.");
})();
