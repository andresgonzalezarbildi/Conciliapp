"use strict";

const DEFAULT_CONFIG = Object.freeze({
  dateTolerance: 1,
  amountAbsTolerance: 0.10,
  amountPercentTolerance: 0,
  maxGroupSize: 8,
  maxCombinations: 25000,
  requireSameSign: true,
  invertBetweenTables: false,
  searchOneToOne: true,
  searchOneToMany: true,
  searchManyToOne: true,
  considerDescription: true,
  autoThreshold: 70,
  possibleThreshold: 55,
  ignoreCase: true,
  ignoreAccents: true,
  ignorePunctuation: true,
  ignoredWords: "pago, recibo, transferencia",
  excludeReconciled: false
});

const MAX_FILE_SIZE = 25 * 1024 * 1024;
const PAGE_SIZE = 20;
const DATE_MS = 86400000;
const ACCEPTED_EXTENSIONS = ["xlsx", "xls", "csv"];
const FIELD_LABELS = {
  date: "Fecha",
  description: "Descripción",
  amount: "Monto",
  debit: "Débito",
  credit: "Crédito",
  status: "Estado (opcional)"
};

const state = {
  step: 1,
  maxVisitedStep: 1,
  sources: {
    system: createEmptySource("system", "Sistema contable"),
    bank: createEmptySource("bank", "Caja o banco")
  },
  config: { ...DEFAULT_CONFIG },
  results: createEmptyResults(),
  processing: { cancelled: false, running: false },
  review: {
    tab: "confirmed",
    search: "",
    type: "all",
    sort: "score-desc",
    page: 1,
    selectedSystem: new Set(),
    selectedBank: new Set(),
    editingId: null
  }
};

function createEmptySource(key, label) {
  return {
    key,
    label,
    name: "",
    file: null,
    workbook: null,
    sheetNames: [],
    selectedSheet: "",
    matrix: [],
    headerRowIndex: 0,
    headers: [],
    rows: [],
    formatMode: "auto",
    detectedFormat: "signed",
    positiveMeaning: "debit",
    splitConvention: "preserve",
    allowBoth: false,
    mapping: { date: "", description: "", amount: "", debit: "", credit: "", status: "" },
    movements: [],
    invalidRows: [],
    validationErrors: [],
    validationWarnings: [],
    isValid: false
  };
}

function createEmptyResults() {
  return {
    reconciliations: [],
    processingAt: null,
    evaluatedCombinations: 0,
    combinationLimitReached: false,
    nextId: 1
  };
}

const dom = {};

document.addEventListener("DOMContentLoaded", initializeApplication);

function initializeApplication() {
  cacheDom();
  bindGlobalEvents();
  bindSourceEditor("system");
  bindSourceEditor("bank");
  populateConfigForm();
  renderProgress();
  refreshIcons();
  if (typeof XLSX === "undefined") {
    showToast("No se pudo cargar el componente de Excel", "Verifique la conexión a Internet y vuelva a abrir el archivo.", "error", 9000);
  }
}

function cacheDom() {
  dom.panels = [...document.querySelectorAll(".step-panel")];
  dom.progress = [...document.querySelectorAll("#progressSteps li")];
  dom.configForm = document.getElementById("configForm");
  dom.processingBar = document.getElementById("processingBar");
  dom.processingPercent = document.getElementById("processingPercent");
  dom.processingMessage = document.getElementById("processingMessage");
  dom.processingDetail = document.getElementById("processingDetail");
  dom.reviewContent = document.getElementById("reviewContent");
  dom.summaryCards = document.getElementById("summaryCards");
  dom.pagination = document.getElementById("pagination");
  dom.formatsDialog = document.getElementById("formatsDialog");
  dom.detailDialog = document.getElementById("detailDialog");
  dom.editGroupDialog = document.getElementById("editGroupDialog");
  dom.toastRegion = document.getElementById("toastRegion");
  dom.continueButtons = {
    system: document.getElementById("continueSystemBtn"),
    bank: document.getElementById("continueBankBtn")
  };
}

function bindGlobalEvents() {
  document.getElementById("startBtn").addEventListener("click", () => goToStep(2));
  document.querySelectorAll("[data-go-step]").forEach(button => {
    button.addEventListener("click", () => goToStep(Number(button.dataset.goStep)));
  });
  document.querySelectorAll("[data-open-formats]").forEach(button => {
    button.addEventListener("click", () => dom.formatsDialog.showModal());
  });
  dom.progress.forEach(item => {
    item.addEventListener("click", () => {
      const target = Number(item.dataset.step);
      if (target <= state.maxVisitedStep && target !== 5) goToStep(target);
    });
  });
  document.getElementById("newReconciliationBtn").addEventListener("click", confirmNewReconciliation);
  document.getElementById("finishNewBtn").addEventListener("click", confirmNewReconciliation);
  dom.continueButtons.system.addEventListener("click", () => goToStep(3));
  dom.continueButtons.bank.addEventListener("click", () => {
    renderConfigSummary();
    goToStep(4);
  });
  document.getElementById("resetConfigBtn").addEventListener("click", () => {
    state.config = { ...DEFAULT_CONFIG };
    populateConfigForm();
    showToast("Parámetros restablecidos", "Se recuperaron los valores iniciales.", "success");
  });
  dom.configForm.addEventListener("input", readConfigForm);
  dom.configForm.addEventListener("change", readConfigForm);
  document.getElementById("processBtn").addEventListener("click", startReconciliation);
  document.getElementById("cancelProcessBtn").addEventListener("click", () => { state.processing.cancelled = true; });
  document.getElementById("reprocessBtn").addEventListener("click", () => goToStep(4));
  document.getElementById("continueExportBtn").addEventListener("click", () => {
    renderExportSummary();
    goToStep(7);
  });
  document.getElementById("downloadExcelBtn").addEventListener("click", exportWorkbook);
  document.querySelectorAll("[data-review-tab]").forEach(button => {
    button.addEventListener("click", () => {
      state.review.tab = button.dataset.reviewTab;
      state.review.page = 1;
      renderReview();
    });
  });
  document.getElementById("reviewSearch").addEventListener("input", event => {
    state.review.search = event.target.value.trim().toLowerCase();
    state.review.page = 1;
    renderReviewContent();
  });
  document.getElementById("typeFilter").addEventListener("change", event => {
    state.review.type = event.target.value;
    state.review.page = 1;
    renderReviewContent();
  });
  document.getElementById("sortResults").addEventListener("change", event => {
    state.review.sort = event.target.value;
    state.review.page = 1;
    renderReviewContent();
  });
  document.querySelectorAll("[data-close-detail]").forEach(button => button.addEventListener("click", () => dom.detailDialog.close()));
  document.querySelectorAll("[data-close-group]").forEach(button => button.addEventListener("click", () => dom.editGroupDialog.close()));
  document.getElementById("saveGroupBtn").addEventListener("click", saveEditedGroup);
  window.addEventListener("keydown", event => {
    if (event.key === "Escape") {
      if (dom.detailDialog.open) dom.detailDialog.close();
      if (dom.editGroupDialog.open) dom.editGroupDialog.close();
    }
  });
}

function refreshIcons(root = document) {
  if (window.lucide) window.lucide.createIcons({ nodes: root.querySelectorAll("[data-lucide]") });
}

function goToStep(step) {
  if (step === 3 && !state.sources.system.isValid) return;
  if (step >= 4 && (!state.sources.system.isValid || !state.sources.bank.isValid)) return;
  if (step >= 6 && !state.results.processingAt) return;
  state.step = step;
  if (step !== 5) state.maxVisitedStep = Math.max(state.maxVisitedStep, step);
  dom.panels.forEach(panel => panel.classList.toggle("active", Number(panel.dataset.panel) === step));
  renderProgress();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function renderProgress() {
  dom.progress.forEach(item => {
    const step = Number(item.dataset.step);
    item.classList.toggle("active", step === state.step);
    item.classList.toggle("completed", step < state.step || (step < state.maxVisitedStep && step !== state.step));
    item.style.cursor = step <= state.maxVisitedStep && step !== 5 ? "pointer" : "default";
    item.setAttribute("aria-current", step === state.step ? "step" : "false");
  });
}

function bindSourceEditor(sourceKey) {
  const editor = document.querySelector(`[data-source-editor="${sourceKey}"]`);
  const source = state.sources[sourceKey];
  const dropZone = editor.querySelector("[data-drop-zone]");
  const fileInput = editor.querySelector("[data-file-input]");
  editor._refs = {
    dropZone,
    fileInput,
    fileSummary: editor.querySelector("[data-file-summary]"),
    importWorkspace: editor.querySelector("[data-import-workspace]"),
    sheetSelect: editor.querySelector("[data-sheet-select]"),
    detectionNote: editor.querySelector("[data-detection-note]"),
    mappingGrid: editor.querySelector("[data-mapping-grid]"),
    validationBox: editor.querySelector("[data-validation-box]"),
    previewTable: editor.querySelector("[data-preview-table]"),
    nameInput: editor.querySelector("[data-source-name]"),
    formatMode: editor.querySelector("[data-format-mode]"),
    positiveMeaning: editor.querySelector("[data-positive-meaning]"),
    splitConvention: editor.querySelector("[data-split-convention]"),
    allowBoth: editor.querySelector("[data-allow-both]")
  };
  const refs = editor._refs;
  dropZone.addEventListener("click", () => fileInput.click());
  dropZone.addEventListener("keydown", event => {
    if (event.key === "Enter" || event.key === " ") { event.preventDefault(); fileInput.click(); }
  });
  ["dragenter", "dragover"].forEach(type => dropZone.addEventListener(type, event => {
    event.preventDefault();
    dropZone.classList.add("dragging");
  }));
  ["dragleave", "drop"].forEach(type => dropZone.addEventListener(type, event => {
    event.preventDefault();
    dropZone.classList.remove("dragging");
  }));
  dropZone.addEventListener("drop", event => {
    const file = event.dataTransfer.files[0];
    if (file) loadSourceFile(sourceKey, file);
  });
  fileInput.addEventListener("change", event => {
    const file = event.target.files[0];
    if (file) loadSourceFile(sourceKey, file);
  });
  refs.nameInput.addEventListener("input", event => { source.name = event.target.value.trim(); });
  refs.formatMode.addEventListener("change", event => {
    source.formatMode = event.target.value;
    toggleFormatOptions(editor, source);
    if (source.workbook) { autoMapColumns(source); renderSourceEditor(sourceKey); }
  });
  refs.positiveMeaning.addEventListener("change", event => { source.positiveMeaning = event.target.value; validateSource(sourceKey); });
  refs.splitConvention.addEventListener("change", event => { source.splitConvention = event.target.value; validateSource(sourceKey); });
  refs.allowBoth.addEventListener("change", event => { source.allowBoth = event.target.checked; validateSource(sourceKey); });
  refs.sheetSelect.addEventListener("change", event => {
    source.selectedSheet = event.target.value;
    extractSelectedSheet(source);
    autoMapColumns(source);
    renderSourceEditor(sourceKey);
  });
  toggleFormatOptions(editor, source);
}

function toggleFormatOptions(editor, source) {
  const effective = source.formatMode === "auto" ? source.detectedFormat : source.formatMode;
  editor.querySelectorAll(".signed-option").forEach(node => node.classList.toggle("hidden", effective !== "signed"));
  editor.querySelectorAll(".split-option").forEach(node => node.classList.toggle("hidden", effective !== "split"));
}

async function loadSourceFile(sourceKey, file) {
  const source = state.sources[sourceKey];
  const extension = file.name.split(".").pop().toLowerCase();
  if (!ACCEPTED_EXTENSIONS.includes(extension)) {
    showToast("Formato no admitido", "Seleccione un archivo XLSX, XLS o CSV.", "error");
    return;
  }
  if (file.size > MAX_FILE_SIZE) {
    showToast("El archivo supera 25 MB", "Puede intentarlo, pero se recomienda reducirlo para evitar bloqueos.", "error", 7000);
  }
  if (typeof XLSX === "undefined") {
    showToast("Lector de Excel no disponible", "Vuelva a abrir la aplicación con conexión a Internet.", "error");
    return;
  }
  try {
    const data = await file.arrayBuffer();
    const isCsv = extension === "csv";
    const input = isCsv ? new TextDecoder("utf-8").decode(data) : data;
    const workbook = XLSX.read(input, {
      type: isCsv ? "string" : "array",
      cellDates: true,
      dense: false,
      raw: isCsv,
      codepage: 65001
    });
    if (!workbook.SheetNames.length) throw new Error("El archivo no contiene hojas legibles.");
    Object.assign(source, {
      file,
      workbook,
      sheetNames: [...workbook.SheetNames],
      selectedSheet: workbook.SheetNames[0],
      invalidRows: [],
      validationErrors: [],
      validationWarnings: [],
      isValid: false
    });
    extractSelectedSheet(source);
    autoMapColumns(source);
    renderSourceEditor(sourceKey);
    showToast("Archivo cargado", `${file.name} se leyó correctamente.`, "success");
  } catch (error) {
    console.error(error);
    showToast("No se pudo leer el archivo", error.message || "El contenido no parece ser un Excel o CSV válido.", "error", 8000);
  }
}

function extractSelectedSheet(source) {
  const sheet = source.workbook.Sheets[source.selectedSheet];
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "", blankrows: false });
  const headerRowIndex = matrix.findIndex(row => Array.isArray(row) && row.some(value => String(value).trim() !== ""));
  if (headerRowIndex < 0) {
    source.matrix = [];
    source.headers = [];
    source.rows = [];
    return;
  }
  const rawHeaders = matrix[headerRowIndex];
  const width = Math.max(rawHeaders.length, ...matrix.slice(headerRowIndex + 1).map(row => row.length));
  const headers = Array.from({ length: width }, (_, index) => {
    const value = String(rawHeaders[index] ?? "").trim();
    return value || `Columna ${columnLetter(index + 1)}`;
  });
  source.matrix = matrix;
  source.headerRowIndex = headerRowIndex;
  source.headers = headers;
  source.rows = matrix.slice(headerRowIndex + 1).map((values, index) => ({
    excelRow: headerRowIndex + index + 2,
    values: Array.from({ length: width }, (_, column) => values[column] ?? "")
  })).filter(row => row.values.some(value => String(value).trim() !== ""));
}

function autoMapColumns(source) {
  const normalized = source.headers.map(header => normalizeHeader(header));
  const find = aliases => {
    let index = normalized.findIndex(header => aliases.includes(header));
    if (index < 0) index = normalized.findIndex(header => aliases.some(alias => header.includes(alias)));
    return index < 0 ? "" : String(index);
  };
  const detected = {
    date: find(["fecha", "date", "fec", "fecha movimiento", "fecha valor"]),
    description: find(["descripcion", "description", "concepto", "detalle", "glosa", "referencia", "movimiento"]),
    amount: find(["monto", "importe", "amount", "saldo movimiento", "valor"]),
    debit: find(["debito", "debe", "debit", "egreso", "retiro", "cargo"]),
    credit: find(["credito", "haber", "credit", "ingreso", "deposito", "abono"]),
    status: find(["estado", "status", "conciliado", "situacion"])
  };
  const hasSplit = detected.debit !== "" && detected.credit !== "";
  const hasSigned = detected.amount !== "";
  source.detectedFormat = hasSplit && !hasSigned ? "split" : "signed";
  const effective = source.formatMode === "auto" ? source.detectedFormat : source.formatMode;
  source.mapping = {
    date: detected.date,
    description: detected.description,
    amount: effective === "signed" ? detected.amount : "",
    debit: effective === "split" ? detected.debit : "",
    credit: effective === "split" ? detected.credit : "",
    status: detected.status
  };
}

function renderSourceEditor(sourceKey) {
  const source = state.sources[sourceKey];
  const editor = document.querySelector(`[data-source-editor="${sourceKey}"]`);
  const refs = editor._refs;
  refs.nameInput.value = source.name;
  refs.formatMode.value = source.formatMode;
  refs.positiveMeaning.value = source.positiveMeaning;
  refs.splitConvention.value = source.splitConvention;
  refs.allowBoth.checked = source.allowBoth;
  toggleFormatOptions(editor, source);
  if (!source.file) return;
  refs.fileSummary.classList.remove("hidden");
  refs.importWorkspace.classList.remove("hidden");
  refs.fileSummary.innerHTML = `<div><i data-lucide="file-spreadsheet"></i><span><strong>${escapeHtml(source.file.name)}</strong><small>${formatFileSize(source.file.size)} · ${source.rows.length.toLocaleString("es-UY")} filas detectadas</small></span></div><button class="button button-ghost button-small" type="button" data-replace-file><i data-lucide="replace"></i> Reemplazar</button>`;
  refs.fileSummary.querySelector("[data-replace-file]").addEventListener("click", () => refs.fileInput.click());
  refs.sheetSelect.innerHTML = source.sheetNames.map(name => `<option value="${escapeAttribute(name)}"${name === source.selectedSheet ? " selected" : ""}>${escapeHtml(name)}</option>`).join("");
  const effective = getEffectiveFormat(source);
  refs.detectionNote.innerHTML = source.formatMode === "auto"
    ? `Formato detectado: <strong>${effective === "signed" ? "Monto con signo" : "Débito y Crédito"}</strong>`
    : `Formato seleccionado manualmente: <strong>${effective === "signed" ? "Monto con signo" : "Débito y Crédito"}</strong>`;
  renderMappingGrid(source, refs.mappingGrid);
  renderPreview(source, refs.previewTable);
  validateSource(sourceKey);
  refreshIcons(editor);
}

function renderMappingGrid(source, container) {
  const effective = getEffectiveFormat(source);
  const fields = effective === "signed"
    ? ["date", "description", "amount", "status"]
    : ["date", "description", "debit", "credit", "status"];
  container.innerHTML = fields.map(field => {
    const required = field !== "status";
    const options = [`<option value="">${required ? "Seleccionar columna" : "No utilizar"}</option>`]
      .concat(source.headers.map((header, index) => `<option value="${index}"${source.mapping[field] === String(index) ? " selected" : ""}>${escapeHtml(header)}</option>`));
    return `<label class="field ${required ? "required" : ""}"><span>${FIELD_LABELS[field]}</span><select data-map-field="${field}">${options.join("")}</select></label>`;
  }).join("");
  container.querySelectorAll("[data-map-field]").forEach(select => {
    select.addEventListener("change", event => {
      source.mapping[event.target.dataset.mapField] = event.target.value;
      validateSource(source.key);
    });
  });
}

function renderPreview(source, table) {
  const rows = source.rows.slice(0, 10);
  table.innerHTML = `<thead><tr><th class="row-number">Fila</th>${source.headers.map(header => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead><tbody>${rows.map(row => `<tr><td class="row-number">${row.excelRow}</td>${row.values.map(value => `<td title="${escapeAttribute(displayOriginalValue(value))}">${escapeHtml(displayOriginalValue(value) || "—")}</td>`).join("")}</tr>`).join("")}</tbody>`;
}

function validateSource(sourceKey) {
  const source = state.sources[sourceKey];
  const errors = [];
  const warnings = [];
  const invalidRows = [];
  const movements = [];
  const effective = getEffectiveFormat(source);
  if (!source.headers.length) errors.push("No se encontró una fila de encabezados válida.");
  if (source.mapping.date === "") errors.push("No se encontró una columna de fecha válida.");
  if (source.mapping.description === "") errors.push("Debe seleccionar una columna para la descripción.");
  if (effective === "signed" && source.mapping.amount === "") errors.push("Debe seleccionar una columna para el monto.");
  if (effective === "split" && (source.mapping.debit === "" || source.mapping.credit === "")) errors.push("Debe seleccionar las columnas Débito y Crédito.");
  if (effective === "split" && source.mapping.debit !== "" && source.mapping.debit === source.mapping.credit) errors.push("Las columnas Débito y Crédito no pueden ser la misma.");
  const assigned = Object.entries(source.mapping).filter(([, value]) => value !== "");
  const duplicate = assigned.find(([field, value], index) => assigned.some(([otherField, otherValue], otherIndex) => otherIndex !== index && otherValue === value && fieldsAreIncompatible(field, otherField)));
  if (duplicate) errors.push("No se puede asignar una misma columna a varios campos incompatibles.");
  if (!errors.length) {
    for (const row of source.rows) {
      const parsed = normalizeSourceRow(source, row, effective);
      if (parsed.errors.length) {
        invalidRows.push({ source: source.label, sheet: source.selectedSheet, row: row.excelRow, errors: parsed.errors.join(" "), values: [...row.values] });
      } else if (!parsed.skip) {
        movements.push(parsed.movement);
      }
    }
    if (!movements.length) errors.push("No hay al menos una fila válida para procesar.");
    if (invalidRows.length) warnings.push(`Hay ${invalidRows.length} ${invalidRows.length === 1 ? "fila que no pudo interpretarse" : "filas que no pudieron interpretarse"}.`);
  }
  source.validationErrors = errors;
  source.validationWarnings = warnings;
  source.invalidRows = invalidRows;
  source.movements = movements;
  source.isValid = errors.length === 0 && movements.length > 0;
  renderValidation(source);
  dom.continueButtons[sourceKey].disabled = !source.isValid;
}

function fieldsAreIncompatible(a, b) {
  if (a === b) return false;
  return true;
}

function normalizeSourceRow(source, row, effectiveFormat) {
  const errors = [];
  const rawDate = valueAt(row, source.mapping.date);
  const rawDescription = valueAt(row, source.mapping.description);
  const date = parseDateValue(rawDate);
  const description = String(rawDescription ?? "").trim();
  if (rawDate === "" || rawDate === null || rawDate === undefined) errors.push(`La fila ${row.excelRow} no contiene fecha.`);
  else if (!date) errors.push(`La fila ${row.excelRow} contiene una fecha no reconocida.`);
  if (!description) errors.push(`La fila ${row.excelRow} no contiene descripción.`);
  let signedAmount = null;
  let originalAmount = "";
  let movementType = "";
  if (effectiveFormat === "signed") {
    const rawAmount = valueAt(row, source.mapping.amount);
    const parsedAmount = parseAmount(rawAmount);
    originalAmount = rawAmount;
    if (rawAmount === "" || rawAmount === null || rawAmount === undefined) errors.push(`La fila ${row.excelRow} no contiene monto.`);
    else if (parsedAmount === null) errors.push(`La fila ${row.excelRow} contiene un monto no reconocido.`);
    else {
      signedAmount = source.positiveMeaning === "debit" ? parsedAmount : -parsedAmount;
      movementType = signedAmount >= 0 ? "debit" : "credit";
    }
  } else {
    const rawDebit = valueAt(row, source.mapping.debit);
    const rawCredit = valueAt(row, source.mapping.credit);
    const debitEmpty = isBlank(rawDebit);
    const creditEmpty = isBlank(rawCredit);
    const debit = debitEmpty ? 0 : parseAmount(rawDebit);
    const credit = creditEmpty ? 0 : parseAmount(rawCredit);
    originalAmount = `Débito: ${displayOriginalValue(rawDebit)} · Crédito: ${displayOriginalValue(rawCredit)}`;
    if (debitEmpty && creditEmpty) errors.push(`La fila ${row.excelRow} no contiene débito ni crédito.`);
    if (!debitEmpty && debit === null) errors.push(`La fila ${row.excelRow} contiene un débito no reconocido.`);
    if (!creditEmpty && credit === null) errors.push(`La fila ${row.excelRow} contiene un crédito no reconocido.`);
    if (!debitEmpty && !creditEmpty && !source.allowBoth) errors.push(`La fila ${row.excelRow} contiene débito y crédito simultáneamente.`);
    if (debit !== null && credit !== null && !(debitEmpty && creditEmpty)) {
      // En columnas separadas, Débito/Crédito indica el tipo de movimiento; el
      // signo de la celda sigue indicando si el importe suma o resta.
      const originalSignedAmount = (debit || 0) + (credit || 0);
      signedAmount = source.splitConvention === "invert" ? -originalSignedAmount : originalSignedAmount;
      movementType = !debitEmpty && creditEmpty ? "debit" : debitEmpty && !creditEmpty ? "credit" : "mixed";
    }
  }
  const status = source.mapping.status === "" ? "" : String(valueAt(row, source.mapping.status) ?? "").trim();
  const skip = state.config.excludeReconciled && isReconciledStatus(status);
  return {
    errors,
    skip,
    movement: errors.length ? null : {
      id: `${source.key}-${String(row.excelRow).padStart(6, "0")}`,
      source: source.key,
      row: row.excelRow,
      date,
      dateKey: toDateKey(date),
      originalDate: rawDate,
      description,
      originalDescription: rawDescription,
      amount: roundMoney(signedAmount),
      originalAmount,
      type: movementType,
      status,
      rawValues: [...row.values]
    }
  };
}

function renderValidation(source) {
  const editor = document.querySelector(`[data-source-editor="${source.key}"]`);
  const box = editor._refs.validationBox;
  let type = "success";
  let icon = "circle-check";
  let title = `${source.movements.length.toLocaleString("es-UY")} movimientos válidos`;
  let body = "El mapeo está listo para continuar.";
  let details = [];
  if (source.validationErrors.length) {
    type = "error"; icon = "circle-x"; title = "El formato necesita correcciones";
    body = source.validationErrors[0]; details = source.validationErrors.slice(1);
  } else if (source.validationWarnings.length) {
    type = "warning"; icon = "triangle-alert"; title = `${source.movements.length.toLocaleString("es-UY")} movimientos válidos con advertencias`;
    body = source.validationWarnings[0];
  }
  box.innerHTML = `<div class="validation-message ${type}"><i data-lucide="${icon}"></i><div><strong>${escapeHtml(title)}</strong><p>${escapeHtml(body)}</p>${details.length ? `<ul>${details.map(item => `<li>${escapeHtml(item)}</li>`).join("")}</ul>` : ""}</div>${source.invalidRows.length ? `<button class="button button-secondary button-small" type="button" data-download-invalid><i data-lucide="download"></i> Reporte de errores</button>` : ""}</div>`;
  const invalidButton = box.querySelector("[data-download-invalid]");
  if (invalidButton) invalidButton.addEventListener("click", () => downloadInvalidRows(source));
  refreshIcons(box);
}

function downloadInvalidRows(source) {
  const headers = ["Origen", "Hoja", "Fila", "Error", ...source.headers];
  const rows = source.invalidRows.map(item => [item.source, item.sheet, item.row, item.errors, ...item.values]);
  const csv = [headers, ...rows].map(row => row.map(csvEscape).join(";")).join("\r\n");
  downloadBlob(new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8" }), `errores_${source.key}_${toFileTimestamp()}.csv`);
}

function getEffectiveFormat(source) {
  return source.formatMode === "auto" ? source.detectedFormat : source.formatMode;
}

function valueAt(row, mappingValue) {
  return mappingValue === "" ? "" : row.values[Number(mappingValue)];
}

function isBlank(value) {
  return value === "" || value === null || value === undefined || (typeof value === "string" && value.trim() === "");
}

function parseAmount(input) {
  if (typeof input === "number") return Number.isFinite(input) ? input : null;
  if (typeof input !== "string") return null;
  let value = input.trim();
  if (!value) return null;
  const negativeParentheses = /^\s*\(.*\)\s*$/.test(value);
  value = value.replace(/[()]/g, "").replace(/[^0-9,\.\-+]/g, "").replace(/^\+/, "");
  if (!value || !/[0-9]/.test(value)) return null;
  const sign = value.startsWith("-") || negativeParentheses ? -1 : 1;
  value = value.replace(/[+-]/g, "");
  const lastComma = value.lastIndexOf(",");
  const lastDot = value.lastIndexOf(".");
  if (lastComma >= 0 && lastDot >= 0) {
    const decimal = lastComma > lastDot ? "," : ".";
    const thousands = decimal === "," ? /\./g : /,/g;
    value = value.replace(thousands, "").replace(decimal, ".");
  } else {
    const separator = lastComma >= 0 ? "," : lastDot >= 0 ? "." : "";
    if (separator) {
      const parts = value.split(separator);
      const decimalDigits = parts[parts.length - 1].length;
      const hasRepeatedSeparators = parts.length > 2;
      if (hasRepeatedSeparators) {
        if (decimalDigits === 2) value = `${parts.slice(0, -1).join("")}.${parts.at(-1)}`;
        else value = parts.join("");
      } else if (decimalDigits === 3 && parts[0].length >= 1 && parts[0].length <= 3) {
        value = parts.join("");
      } else {
        value = value.replace(separator, ".");
      }
    }
  }
  const number = Number(value);
  return Number.isFinite(number) ? sign * Math.abs(number) : null;
}

function parseDateValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return startOfDay(value);
  if (typeof value === "number" && Number.isFinite(value)) {
    if (window.XLSX?.SSF?.parse_date_code) {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (parsed) return new Date(Date.UTC(parsed.y, parsed.m - 1, parsed.d));
    }
    if (value > 20000 && value < 100000) return new Date(Date.UTC(1899, 11, 30 + Math.floor(value)));
    return null;
  }
  if (typeof value !== "string") return null;
  const text = value.trim();
  if (!text) return null;
  const match = text.match(/^(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})(?:\s+.*)?$/);
  if (match) {
    let year;
    let month;
    let day;
    if (match[1].length === 4) {
      year = Number(match[1]); month = Number(match[2]); day = Number(match[3]);
    } else {
      day = Number(match[1]); month = Number(match[2]); year = Number(match[3]);
      if (year < 100) year += year < 50 ? 2000 : 1900;
    }
    const date = new Date(Date.UTC(year, month - 1, day));
    if (date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day) return date;
    return null;
  }
  const timestamp = Date.parse(text);
  return Number.isNaN(timestamp) ? null : startOfDay(new Date(timestamp));
}

function startOfDay(date) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
}

function normalizeHeader(value) {
  return String(value ?? "").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]+/g, " ").trim();
}

function normalizeDescription(value, config = state.config) {
  let text = String(value ?? "").trim();
  if (config.ignoreCase) text = text.toLowerCase();
  if (config.ignoreAccents) text = text.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  if (config.ignorePunctuation) text = text.replace(/[^\p{L}\p{N}\s]/gu, " ");
  const ignored = String(config.ignoredWords || "").split(",").map(word => word.trim()).filter(Boolean);
  if (ignored.length) {
    const ignoredSet = new Set(ignored.map(word => config.ignoreAccents ? word.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase() : word.toLowerCase()));
    text = text.split(/\s+/).filter(token => !ignoredSet.has(token.toLowerCase())).join(" ");
  }
  return text.replace(/\s+/g, " ").trim();
}

function descriptionSimilarity(a, b) {
  const left = normalizeDescription(a);
  const right = normalizeDescription(b);
  if (!left && !right) return 1;
  if (!left || !right) return 0;
  if (left === right) return 1;
  const leftTokens = new Set(left.split(" ").filter(Boolean));
  const rightTokens = new Set(right.split(" ").filter(Boolean));
  const intersection = [...leftTokens].filter(token => rightTokens.has(token)).length;
  const union = new Set([...leftTokens, ...rightTokens]).size;
  const jaccard = union ? intersection / union : 0;
  const maxLength = Math.max(left.length, right.length);
  const levenshtein = maxLength ? 1 - levenshteinDistance(left, right) / maxLength : 1;
  const containment = left.includes(right) || right.includes(left) ? Math.min(left.length, right.length) / maxLength : 0;
  const leftTokenList = left.split(" ").filter(token => token.length > 1);
  const rightTokenList = right.split(" ").filter(token => token.length > 1);
  const shorterTokens = leftTokenList.length <= rightTokenList.length ? leftTokenList : rightTokenList;
  const longerTokens = leftTokenList.length <= rightTokenList.length ? rightTokenList : leftTokenList;
  const phrases = new Set(longerTokens);
  for (let size = 2; size <= 3; size++) {
    for (let index = 0; index <= longerTokens.length - size; index++) phrases.add(longerTokens.slice(index, index + size).join(""));
  }
  const fuzzyCoverage = shorterTokens.length
    ? shorterTokens.reduce((sum, token) => {
      const best = [...phrases].reduce((maximum, phrase) => {
        const length = Math.max(token.length, phrase.length);
        const similarity = length ? 1 - levenshteinDistance(token, phrase) / length : 1;
        return Math.max(maximum, similarity);
      }, 0);
      return sum + best;
    }, 0) / shorterTokens.length
    : 0;
  const compactLeft = left.replace(/\s+/g, "");
  const compactRight = right.replace(/\s+/g, "");
  const shorterCompact = compactLeft.length <= compactRight.length ? compactLeft : compactRight;
  const longerCompact = compactLeft.length <= compactRight.length ? compactRight : compactLeft;
  const compactContainment = shorterCompact.length >= 5 && longerCompact.includes(shorterCompact) ? .82 : 0;
  const baseSimilarity = jaccard * .5 + levenshtein * .35 + containment * .15;
  return clamp(Math.max(baseSimilarity, fuzzyCoverage * .85, compactContainment), 0, 1);
}

function levenshteinDistance(a, b) {
  if (a.length < b.length) [a, b] = [b, a];
  let previous = Array.from({ length: b.length + 1 }, (_, index) => index);
  for (let i = 1; i <= a.length; i++) {
    const current = [i];
    for (let j = 1; j <= b.length; j++) {
      current[j] = Math.min(current[j - 1] + 1, previous[j] + 1, previous[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    previous = current;
  }
  return previous[b.length];
}

function extractReferences(description) {
  return new Set((String(description).match(/\d{3,}/g) || []).map(value => value.replace(/^0+/, "") || "0"));
}

function referenceScore(descriptionsA, descriptionsB) {
  const refsA = new Set(descriptionsA.flatMap(value => [...extractReferences(value)]));
  const refsB = new Set(descriptionsB.flatMap(value => [...extractReferences(value)]));
  if (!refsA.size || !refsB.size) return 0;
  return [...refsA].some(value => refsB.has(value)) ? 10 : 0;
}

function isGenericDescription(description) {
  const normalized = normalizeDescription(description);
  const tokens = normalized.split(" ").filter(Boolean);
  return tokens.length <= 1 || normalized.length < 5 || ["varios", "ajuste", "movimiento", "deposito", "caja", "banco"].includes(normalized);
}

function isReconciledStatus(status) {
  return /conciliad|reconcil|matched|confirmad|completad|\bok\b/i.test(String(status));
}

function populateConfigForm() {
  Object.entries(state.config).forEach(([key, value]) => {
    const input = dom.configForm.elements.namedItem(key);
    if (!input) return;
    if (input.type === "checkbox") input.checked = Boolean(value);
    else input.value = value;
  });
  updateCostWarning();
}

function readConfigForm() {
  const data = new FormData(dom.configForm);
  const numberFields = ["dateTolerance", "amountAbsTolerance", "amountPercentTolerance", "maxGroupSize", "maxCombinations", "autoThreshold", "possibleThreshold"];
  const checkboxFields = ["requireSameSign", "invertBetweenTables", "searchOneToOne", "searchOneToMany", "searchManyToOne", "considerDescription", "ignoreCase", "ignoreAccents", "ignorePunctuation", "excludeReconciled"];
  numberFields.forEach(key => { state.config[key] = Number(data.get(key)); });
  checkboxFields.forEach(key => { state.config[key] = data.has(key); });
  state.config.ignoredWords = String(data.get("ignoredWords") || "");
  if (state.config.possibleThreshold > state.config.autoThreshold) {
    state.config.possibleThreshold = state.config.autoThreshold;
    dom.configForm.elements.namedItem("possibleThreshold").value = state.config.possibleThreshold;
  }
  updateCostWarning();
}

function updateCostWarning() {
  const warning = document.getElementById("costWarning");
  const complexity = state.config.maxGroupSize * state.config.maxCombinations;
  const risky = state.config.maxGroupSize > 8 || state.config.maxCombinations > 100000 || complexity > 800000;
  warning.classList.toggle("hidden", !risky);
  if (risky) warning.querySelector("span").textContent = "Estos límites pueden producir una búsqueda lenta. La aplicación detendrá la evaluación al alcanzar el máximo configurado.";
}

function renderConfigSummary() {
  const container = document.getElementById("sourceConfigSummary");
  container.innerHTML = [state.sources.system, state.sources.bank].map(source => `<div class="source-pill"><span><i data-lucide="${source.key === "system" ? "database" : "landmark"}"></i></span><div><strong>${escapeHtml(source.name || source.label)}</strong><small>${source.movements.length.toLocaleString("es-UY")} movimientos · ${getEffectiveFormat(source) === "signed" ? "Monto con signo" : "Débito y Crédito"} · ${source.invalidRows.length} errores omitidos</small></div></div>`).join("");
  refreshIcons(container);
}

function resetApplication() {
  state.step = 1;
  state.maxVisitedStep = 1;
  state.sources.system = createEmptySource("system", "Sistema contable");
  state.sources.bank = createEmptySource("bank", "Caja o banco");
  state.config = { ...DEFAULT_CONFIG };
  state.results = createEmptyResults();
  state.processing = { cancelled: false, running: false };
  state.review = { tab: "confirmed", search: "", type: "all", sort: "score-desc", page: 1, selectedSystem: new Set(), selectedBank: new Set(), editingId: null };
  document.querySelectorAll("[data-source-editor]").forEach(editor => {
    editor.querySelector("[data-source-name]").value = "";
    editor.querySelector("[data-file-input]").value = "";
    editor.querySelector("[data-format-mode]").value = "auto";
    editor.querySelector("[data-positive-meaning]").value = "debit";
    editor.querySelector("[data-split-convention]").value = "preserve";
    editor.querySelector("[data-allow-both]").checked = false;
    editor.querySelector("[data-file-summary]").classList.add("hidden");
    editor.querySelector("[data-import-workspace]").classList.add("hidden");
    toggleFormatOptions(editor, state.sources[editor.dataset.sourceEditor]);
  });
  dom.continueButtons.system.disabled = true;
  dom.continueButtons.bank.disabled = true;
  document.getElementById("reviewSearch").value = "";
  document.getElementById("typeFilter").value = "all";
  document.getElementById("sortResults").value = "score-desc";
  populateConfigForm();
  goToStep(1);
  showToast("Datos eliminados", "La aplicación quedó lista para una nueva conciliación.", "success");
}

function confirmNewReconciliation() {
  const hasData = state.sources.system.file || state.sources.bank.file || state.results.processingAt;
  if (!hasData || window.confirm("Se eliminarán de la memoria todos los archivos y resultados cargados. ¿Desea continuar?")) resetApplication();
}

async function startReconciliation() {
  if (state.processing.running) return;
  readConfigForm();
  validateSource("system");
  validateSource("bank");
  if (!state.sources.system.isValid || !state.sources.bank.isValid) {
    showToast("No se puede iniciar", "Revise el mapeo y la validación de ambas tablas.", "error");
    return;
  }
  state.processing = { cancelled: false, running: true };
  state.results = createEmptyResults();
  setProcessingProgress(2, "Preparando y normalizando las tablas…", "Validando movimientos");
  goToStep(5);
  try {
    await yieldToMain();
    if (state.processing.cancelled) return cancelProcessing();
    setProcessingProgress(10, "Buscando coincidencias uno a uno…", `${state.sources.system.movements.length + state.sources.bank.movements.length} movimientos`);
    if (state.config.searchOneToOne) await reconcileOneToOne();
    if (state.processing.cancelled) return cancelProcessing();
    setProcessingProgress(48, "Buscando agrupaciones de movimientos…", "Uno a varios");
    if (state.config.searchOneToMany) await reconcileGroups("one-to-many");
    if (state.processing.cancelled) return cancelProcessing();
    setProcessingProgress(73, "Buscando agrupaciones de movimientos…", "Varios a uno");
    if (state.config.searchManyToOne) await reconcileGroups("many-to-one");
    if (state.processing.cancelled) return cancelProcessing();
    setProcessingProgress(93, "Consolidando resultados…", "Calculando pendientes y totales");
    await yieldToMain();
    state.results.processingAt = new Date();
    state.processing.running = false;
    setProcessingProgress(100, "Conciliación completada", `${state.results.reconciliations.length} propuestas generadas`);
    await delay(280);
    state.review.tab = "confirmed";
    state.review.page = 1;
    state.review.selectedSystem.clear();
    state.review.selectedBank.clear();
    renderReview();
    goToStep(6);
    if (state.results.combinationLimitReached) {
      showToast("Búsqueda limitada", `Se alcanzó el máximo de ${state.config.maxCombinations.toLocaleString("es-UY")} combinaciones. Puede ampliarlo y reprocesar.`, "error", 8500);
    }
  } catch (error) {
    console.error(error);
    state.processing.running = false;
    showToast("No se pudo completar la conciliación", error.message || "Ocurrió un error durante el procesamiento.", "error", 9000);
    goToStep(4);
  }
}

function cancelProcessing() {
  state.processing.running = false;
  state.results = createEmptyResults();
  showToast("Procesamiento cancelado", "No se conservaron resultados parciales.", "error");
  goToStep(4);
}

function setProcessingProgress(percent, message, detail) {
  const value = clamp(Math.round(percent), 0, 100);
  dom.processingBar.style.width = `${value}%`;
  dom.processingPercent.textContent = `${value}%`;
  dom.processingMessage.textContent = message;
  dom.processingDetail.textContent = detail;
  dom.processingBar.parentElement.setAttribute("aria-valuenow", String(value));
}

async function reconcileOneToOne() {
  const system = state.sources.system.movements;
  const bank = state.sources.bank.movements;
  const candidates = [];
  let checked = 0;
  const total = Math.max(1, system.length * bank.length);
  for (const left of system) {
    for (const right of bank) {
      checked++;
      if (!movementsCanMatch(left, right)) continue;
      const calculated = calculateReconciliation([left], [right], "one-to-one");
      if (!calculated.amountWithinTolerance || !calculated.dateWithinTolerance) continue;
      if (calculated.score < state.config.possibleThreshold) continue;
      candidates.push(calculated);
      if (checked % 5000 === 0) {
        setProcessingProgress(10 + 34 * (checked / total), "Buscando coincidencias uno a uno…", `${checked.toLocaleString("es-UY")} comparaciones`);
        await yieldToMain();
        if (state.processing.cancelled) return;
      }
    }
  }
  const bySystem = groupBy(candidates, item => item.systemIds[0]);
  const byBank = groupBy(candidates, item => item.bankIds[0]);
  candidates.forEach(candidate => {
    const leftAlternatives = (bySystem.get(candidate.systemIds[0]) || []).filter(item => item !== candidate && item.score >= candidate.score - 5);
    const rightAlternatives = (byBank.get(candidate.bankIds[0]) || []).filter(item => item !== candidate && item.score >= candidate.score - 5);
    candidate.ambiguous = leftAlternatives.length > 0 || rightAlternatives.length > 0;
    candidate.alternativeCount = leftAlternatives.length + rightAlternatives.length;
    if (candidate.ambiguous) {
      candidate.score = Math.max(0, candidate.score - Math.min(15, 6 + candidate.alternativeCount * 2));
      candidate.reasons.push(`${candidate.alternativeCount} alternativa(s) con puntuación similar`);
    }
  });
  candidates.sort((a, b) => Number(b.exact) - Number(a.exact) || b.score - a.score || a.dateDifference - b.dateDifference);
  const usedSystem = new Set();
  const usedBank = new Set();
  for (const candidate of candidates) {
    if (usedSystem.has(candidate.systemIds[0]) || usedBank.has(candidate.bankIds[0])) continue;
    if (candidate.score < state.config.possibleThreshold) continue;
    candidate.status = candidate.score >= state.config.autoThreshold && !candidate.ambiguous ? "confirmed" : "possible";
    addReconciliation(candidate);
    usedSystem.add(candidate.systemIds[0]);
    usedBank.add(candidate.bankIds[0]);
  }
}

async function reconcileGroups(type) {
  const reserved = getReservedIds();
  const systemRemaining = state.sources.system.movements.filter(item => !reserved.system.has(item.id));
  const bankRemaining = state.sources.bank.movements.filter(item => !reserved.bank.has(item.id));
  const anchors = type === "one-to-many" ? systemRemaining : bankRemaining;
  const pool = type === "one-to-many" ? bankRemaining : systemRemaining;
  const candidates = [];
  const stageStart = type === "one-to-many" ? 48 : 73;
  const stageSpan = type === "one-to-many" ? 23 : 18;
  for (let index = 0; index < anchors.length; index++) {
    if (state.results.evaluatedCombinations >= state.config.maxCombinations) {
      state.results.combinationLimitReached = true;
      break;
    }
    const anchor = anchors[index];
    const compatible = pool
      .filter(item => groupMemberCanMatchAnchor(anchor, item, type))
      .sort((a, b) => Math.abs(Math.abs(comparisonAmount(a)) - Math.abs(comparisonAmount(anchor))) - Math.abs(Math.abs(comparisonAmount(b)) - Math.abs(comparisonAmount(anchor))))
      .slice(0, 24);
    if (compatible.length >= 2) {
      const found = enumerateGroupCandidates(anchor, compatible, type);
      candidates.push(...found);
    }
    if (index % 5 === 0 || index === anchors.length - 1) {
      setProcessingProgress(stageStart + stageSpan * ((index + 1) / Math.max(1, anchors.length)), "Buscando agrupaciones de movimientos…", `${state.results.evaluatedCombinations.toLocaleString("es-UY")} combinaciones evaluadas`);
      await yieldToMain();
      if (state.processing.cancelled) return;
    }
  }
  const byAnchor = groupBy(candidates, candidate => type === "one-to-many" ? candidate.systemIds[0] : candidate.bankIds[0]);
  candidates.forEach(candidate => {
    const anchorId = type === "one-to-many" ? candidate.systemIds[0] : candidate.bankIds[0];
    const alternatives = (byAnchor.get(anchorId) || []).filter(item => item !== candidate && item.score >= candidate.score - 5);
    candidate.ambiguous = alternatives.length > 0;
    candidate.alternativeCount = alternatives.length;
    if (candidate.ambiguous) {
      candidate.score = Math.max(0, candidate.score - Math.min(18, 8 + alternatives.length * 2));
      candidate.reasons.push(`Existen ${alternatives.length + 1} combinaciones posibles para el mismo importe`);
    }
  });
  candidates.sort((a, b) => b.score - a.score || a.totalMembers - b.totalMembers || a.dateDifference - b.dateDifference);
  const currentlyReserved = getReservedIds();
  for (const candidate of candidates) {
    if (candidate.systemIds.some(id => currentlyReserved.system.has(id)) || candidate.bankIds.some(id => currentlyReserved.bank.has(id))) continue;
    if (candidate.score < state.config.possibleThreshold) continue;
    candidate.status = candidate.score >= state.config.autoThreshold && !candidate.ambiguous ? "confirmed" : "possible";
    addReconciliation(candidate);
    candidate.systemIds.forEach(id => currentlyReserved.system.add(id));
    candidate.bankIds.forEach(id => currentlyReserved.bank.add(id));
  }
}

function enumerateGroupCandidates(anchor, compatible, type) {
  const results = [];
  const maxSize = Math.min(state.config.maxGroupSize, compatible.length);
  const target = Math.abs(comparisonAmount(anchor));
  const tolerance = amountToleranceFor(target);
  const members = compatible.map(item => ({ item, absoluteAmount: Math.abs(comparisonAmount(item)) }));
  function visit(start, selected, sum) {
    if (state.results.evaluatedCombinations >= state.config.maxCombinations) return;
    if (selected.length >= 2) {
      state.results.evaluatedCombinations++;
      if (Math.abs(target - sum) <= tolerance + .005) {
        const selectedMovements = selected.map(entry => entry.item);
        const systemMovements = type === "one-to-many" ? [anchor] : selectedMovements;
        const bankMovements = type === "one-to-many" ? selectedMovements : [anchor];
        const calculated = calculateReconciliation(systemMovements, bankMovements, type);
        if (calculated.amountWithinTolerance && calculated.dateWithinTolerance && calculated.score >= state.config.possibleThreshold) results.push(calculated);
      }
    }
    if (selected.length >= maxSize || start >= members.length || sum > target + tolerance + .005) return;
    for (let index = start; index < members.length; index++) {
      if (state.results.evaluatedCombinations >= state.config.maxCombinations) return;
      const next = members[index];
      visit(index + 1, [...selected, next], sum + next.absoluteAmount);
    }
  }
  visit(0, [], 0);
  return results;
}

function movementsCanMatch(systemMovement, bankMovement) {
  const dateDifference = daysBetween(systemMovement.date, bankMovement.date);
  if (dateDifference > state.config.dateTolerance) return false;
  if (state.config.requireSameSign && Math.sign(comparisonAmount(systemMovement)) !== Math.sign(comparisonAmount(bankMovement))) return false;
  const difference = Math.abs(comparisonAmount(systemMovement) - comparisonAmount(bankMovement));
  return difference <= amountToleranceFor(Math.max(Math.abs(comparisonAmount(systemMovement)), Math.abs(comparisonAmount(bankMovement)))) + .005;
}

function groupMemberCanMatchAnchor(anchor, member, type) {
  if (daysBetween(anchor.date, member.date) > state.config.dateTolerance) return false;
  const anchorAmount = type === "one-to-many" ? comparisonAmount(anchor) : comparisonAmount(anchor);
  const memberAmount = comparisonAmount(member);
  if (state.config.requireSameSign && Math.sign(anchorAmount) !== Math.sign(memberAmount)) return false;
  return Math.abs(memberAmount) <= Math.abs(anchorAmount) + amountToleranceFor(Math.abs(anchorAmount)) + .005;
}

function comparisonAmount(movement) {
  if (movement.source === "bank" && state.config.invertBetweenTables) return -movement.amount;
  return movement.amount;
}

function calculateReconciliation(systemMovements, bankMovements, type = inferType(systemMovements, bankMovements)) {
  const totalSystem = roundMoney(systemMovements.reduce((sum, movement) => sum + comparisonAmount(movement), 0));
  const totalBankComparison = roundMoney(bankMovements.reduce((sum, movement) => sum + comparisonAmount(movement), 0));
  const totalBankOriginal = roundMoney(bankMovements.reduce((sum, movement) => sum + movement.amount, 0));
  const difference = roundMoney(totalSystem - totalBankComparison);
  const tolerance = amountToleranceFor(Math.max(Math.abs(totalSystem), Math.abs(totalBankComparison)));
  const amountWithinTolerance = Math.abs(difference) <= tolerance + .005;
  const amountScore = amountWithinTolerance
    ? tolerance <= .005 ? (Math.abs(difference) <= .005 ? 45 : 0) : 45 - Math.min(5, 5 * Math.abs(difference) / Math.max(tolerance, .01))
    : 0;
  const dateDifferences = systemMovements.flatMap(system => bankMovements.map(bank => daysBetween(system.date, bank.date)));
  const dateDifference = dateDifferences.length ? Math.max(...dateDifferences) : 0;
  const averageDateDifference = dateDifferences.length ? dateDifferences.reduce((sum, value) => sum + value, 0) / dateDifferences.length : 0;
  const dateWithinTolerance = dateDifference <= state.config.dateTolerance;
  const dateScore = dateWithinTolerance
    ? state.config.dateTolerance === 0 ? 25 : 25 * Math.max(.35, 1 - averageDateDifference / (state.config.dateTolerance + 1))
    : 0;
  const systemDescription = systemMovements.map(item => item.description).join(" ");
  const bankDescription = bankMovements.map(item => item.description).join(" ");
  const similarity = state.config.considerDescription ? descriptionSimilarity(systemDescription, bankDescription) : 1;
  const descriptionScore = 20 * similarity;
  const refsScore = referenceScore(systemMovements.map(item => item.description), bankMovements.map(item => item.description));
  const totalMembers = systemMovements.length + bankMovements.length;
  let penalty = 0;
  if (type !== "one-to-one") penalty += Math.max(0, totalMembers - 3) * 2;
  if (systemMovements.some(item => isGenericDescription(item.description)) || bankMovements.some(item => isGenericDescription(item.description))) penalty += 3;
  if (dateDifference > 0) penalty += Math.min(5, dateDifference);
  const score = Math.round(clamp(amountScore + dateScore + descriptionScore + refsScore - penalty, 0, 100));
  const exact = Math.abs(difference) <= .005 && dateWithinTolerance && (!state.config.considerDescription || similarity >= .9);
  const reasons = [
    Math.abs(difference) <= .005 ? "Los importes coinciden exactamente" : `La diferencia de importe (${formatMoney(Math.abs(difference))}) está dentro de la tolerancia`,
    dateDifference === 0 ? "Las fechas coinciden" : `Las fechas difieren hasta ${dateDifference} día(s)`,
    state.config.considerDescription ? `Similitud de descripciones: ${Math.round(similarity * 100)}%` : "La comparación de descripciones está desactivada"
  ];
  if (refsScore) reasons.push("Se encontraron referencias numéricas coincidentes");
  if (type !== "one-to-one") reasons.push(`La suma de ${type === "one-to-many" ? bankMovements.length : systemMovements.length} movimientos coincide con el otro lado`);
  return {
    id: null,
    type,
    systemIds: systemMovements.map(item => item.id),
    bankIds: bankMovements.map(item => item.id),
    systemMovements,
    bankMovements,
    totalSystem,
    totalBank: totalBankComparison,
    totalBankOriginal,
    difference,
    amountTolerance: tolerance,
    amountWithinTolerance,
    dateWithinTolerance,
    dateDifference,
    descriptionSimilarity: similarity,
    score,
    exact,
    ambiguous: false,
    alternativeCount: 0,
    totalMembers,
    status: "possible",
    criterion: exact ? "Monto exacto, fecha compatible y descripción coincidente" : type === "one-to-one" ? "Coincidencia probable uno a uno" : "Agrupación por suma de importes",
    reasons,
    observation: "",
    createdAt: new Date(),
    manual: type === "manual"
  };
}

function addReconciliation(reconciliation) {
  reconciliation.id = `CON-${String(state.results.nextId++).padStart(5, "0")}`;
  state.results.reconciliations.push(reconciliation);
}

function amountToleranceFor(reference) {
  return Math.max(state.config.amountAbsTolerance, Math.abs(reference) * state.config.amountPercentTolerance / 100);
}

function inferType(systemMovements, bankMovements) {
  if (systemMovements.length === 1 && bankMovements.length === 1) return "one-to-one";
  if (systemMovements.length === 1) return "one-to-many";
  if (bankMovements.length === 1) return "many-to-one";
  return "manual";
}

function getReservedIds() {
  const system = new Set();
  const bank = new Set();
  state.results.reconciliations.filter(item => item.status === "confirmed" || item.status === "possible").forEach(item => {
    item.systemIds.forEach(id => system.add(id));
    item.bankIds.forEach(id => bank.add(id));
  });
  return { system, bank };
}

function getConfirmedIds() {
  const system = new Set();
  const bank = new Set();
  state.results.reconciliations.filter(item => item.status === "confirmed").forEach(item => {
    item.systemIds.forEach(id => system.add(id));
    item.bankIds.forEach(id => bank.add(id));
  });
  return { system, bank };
}

function getPendingMovements(sourceKey) {
  const reserved = getReservedIds()[sourceKey];
  return state.sources[sourceKey].movements.filter(item => !reserved.has(item.id));
}

function calculateSummary() {
  const totalSystem = state.sources.system.movements.length;
  const totalBank = state.sources.bank.movements.length;
  const total = totalSystem + totalBank;
  const confirmedIds = getConfirmedIds();
  const reservedIds = getReservedIds();
  const confirmedCount = confirmedIds.system.size + confirmedIds.bank.size;
  const possibleCount = (reservedIds.system.size - confirmedIds.system.size) + (reservedIds.bank.size - confirmedIds.bank.size);
  const pendingSystem = getPendingMovements("system");
  const pendingBank = getPendingMovements("bank");
  const pendingCount = pendingSystem.length + pendingBank.length;
  const confirmedReconciliations = state.results.reconciliations.filter(item => item.status === "confirmed");
  const reconciledAmount = confirmedReconciliations.reduce((sum, item) => sum + Math.abs(item.totalSystem), 0);
  const pendingSystemAmount = pendingSystem.reduce((sum, item) => sum + comparisonAmount(item), 0);
  const pendingBankAmount = pendingBank.reduce((sum, item) => sum + comparisonAmount(item), 0);
  const pendingDifference = roundMoney(pendingSystemAmount - pendingBankAmount);
  return {
    total,
    totalSystem,
    totalBank,
    confirmedCount,
    possibleCount,
    pendingCount,
    pendingSystem,
    pendingBank,
    reconciledAmount: roundMoney(reconciledAmount),
    pendingDifference,
    percentage: total ? confirmedCount / total * 100 : 0,
    confirmedReconciliations: confirmedReconciliations.length,
    possibleReconciliations: state.results.reconciliations.filter(item => item.status === "possible").length
  };
}

function renderReview() {
  const summary = calculateSummary();
  const cards = [
    ["Movimientos", summary.total.toLocaleString("es-UY"), ""],
    ["Conciliados", summary.confirmedCount.toLocaleString("es-UY"), "success"],
    ["Posibles", summary.possibleCount.toLocaleString("es-UY"), "warning"],
    ["Pendientes", summary.pendingCount.toLocaleString("es-UY"), "danger"],
    ["Importe conciliado", formatMoney(summary.reconciledAmount), "success"],
    ["Diferencia pendiente", formatMoney(summary.pendingDifference), summary.pendingDifference ? "danger" : "success"],
    ["Avance", `${formatDecimal(summary.percentage, 1)}%`, "success"]
  ];
  dom.summaryCards.innerHTML = cards.map(([label, value, type]) => `<div class="summary-card ${type}"><span title="${escapeAttribute(label)}">${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`).join("");
  document.querySelectorAll("[data-review-tab]").forEach(button => button.classList.toggle("active", button.dataset.reviewTab === state.review.tab));
  document.querySelector('[data-tab-count="confirmed"]').textContent = summary.confirmedReconciliations;
  document.querySelector('[data-tab-count="possible"]').textContent = summary.possibleReconciliations;
  document.querySelector('[data-tab-count="pending"]').textContent = summary.pendingCount;
  renderReviewContent();
}

function renderReviewContent() {
  if (state.review.tab === "pending") {
    renderPendingReview();
    return;
  }
  let items = state.results.reconciliations.filter(item => item.status === state.review.tab);
  if (state.review.type !== "all") items = items.filter(item => item.type === state.review.type || (state.review.type === "manual" && item.manual));
  if (state.review.search) items = items.filter(item => reconciliationSearchText(item).includes(state.review.search));
  items = sortReconciliations(items, state.review.sort);
  const pageCount = Math.max(1, Math.ceil(items.length / PAGE_SIZE));
  state.review.page = Math.min(state.review.page, pageCount);
  const pageItems = items.slice((state.review.page - 1) * PAGE_SIZE, state.review.page * PAGE_SIZE);
  if (!pageItems.length) {
    dom.reviewContent.innerHTML = `<div class="empty-state"><div><i data-lucide="${state.review.tab === "confirmed" ? "circle-check" : "circle-help"}"></i><h3>${state.review.tab === "confirmed" ? "No hay conciliaciones confirmadas" : "No hay posibles conciliaciones"}</h3><p>${state.review.search ? "No se encontraron resultados para la búsqueda o los filtros actuales." : state.review.tab === "confirmed" ? "Las coincidencias aprobadas aparecerán aquí." : "No quedaron casos dudosos para revisar."}</p></div></div>`;
    dom.pagination.innerHTML = "";
    refreshIcons(dom.reviewContent);
    return;
  }
  const isPossible = state.review.tab === "possible";
  dom.reviewContent.innerHTML = `<div class="table-scroll"><table class="data-table result-table"><thead><tr><th>ID / Estado</th><th>Tipo</th><th>Sistema contable</th><th>Caja o banco</th><th>Total sistema</th><th>Total caja/banco</th><th>Diferencia</th><th>Confianza</th><th>${isPossible ? "Observación / Acciones" : "Acciones"}</th></tr></thead><tbody>${pageItems.map(item => renderReconciliationRow(item, isPossible)).join("")}</tbody></table></div>`;
  bindResultTableEvents();
  renderPagination(items.length, pageCount);
  refreshIcons(dom.reviewContent);
}

function renderReconciliationRow(item, isPossible) {
  const systemDescription = summarizeMovements(item.systemMovements);
  const bankDescription = summarizeMovements(item.bankMovements);
  const actions = isPossible
    ? `<div class="tabular-actions"><input class="observation-input" data-observation="${item.id}" value="${escapeAttribute(item.observation)}" placeholder="Agregar observación"><button class="table-action approve" data-action="approve" data-id="${item.id}"><i data-lucide="check"></i>Aprobar</button><button class="table-action" data-action="edit" data-id="${item.id}"><i data-lucide="list-restart"></i>Editar</button><button class="table-action reject" data-action="reject" data-id="${item.id}"><i data-lucide="x"></i>Rechazar</button><button class="table-action" data-action="detail" data-id="${item.id}"><i data-lucide="info"></i>Motivo</button></div>`
    : `<div class="tabular-actions"><button class="table-action" data-action="detail" data-id="${item.id}"><i data-lucide="eye"></i>Ver detalle</button><button class="table-action reject" data-action="unmatch" data-id="${item.id}"><i data-lucide="unlink"></i>Quitar conciliación</button></div>`;
  return `<tr class="result-row ${item.status}"><td><strong>${item.id}</strong><br>${statusBadge(item.status)}</td><td><span class="type-badge">${typeLabel(item)}</span></td><td class="descriptions"><strong>${escapeHtml(systemDescription.title)}</strong><span>${escapeHtml(systemDescription.subtitle)}</span></td><td class="descriptions"><strong>${escapeHtml(bankDescription.title)}</strong><span>${escapeHtml(bankDescription.subtitle)}</span></td><td class="amount">${formatMoney(item.totalSystem)}</td><td class="amount">${formatMoney(item.totalBank)}</td><td class="amount ${item.difference ? "negative" : ""}">${formatMoney(item.difference)}</td><td><span class="score-badge ${scoreClass(item.score)}">${item.score}</span></td><td>${actions}</td></tr>`;
}

function bindResultTableEvents() {
  dom.reviewContent.querySelectorAll("[data-action]").forEach(button => {
    button.addEventListener("click", () => handleReconciliationAction(button.dataset.action, button.dataset.id));
  });
  dom.reviewContent.querySelectorAll("[data-observation]").forEach(input => {
    input.addEventListener("change", () => {
      const item = findReconciliation(input.dataset.observation);
      if (item) item.observation = input.value.trim();
    });
  });
}

function handleReconciliationAction(action, id) {
  const item = findReconciliation(id);
  if (!item) return;
  if (action === "detail") return openReconciliationDetail(item);
  if (action === "edit") return openEditGroup(item);
  if (action === "approve") {
    item.status = "confirmed";
    item.observation = document.querySelector(`[data-observation="${id}"]`)?.value.trim() || item.observation;
    showToast("Conciliación aprobada", `${id} pasó a conciliados.`, "success");
  }
  if (action === "unmatch") {
    item.status = "rejected";
    item.observation = item.observation || "Conciliación quitada manualmente";
    showToast("Conciliación quitada", "Sus movimientos volvieron a la lista de pendientes.", "error");
  }
  if (action === "reject") {
    item.status = "rejected";
    item.observation = document.querySelector(`[data-observation="${id}"]`)?.value.trim() || item.observation;
    showToast("Propuesta rechazada", "Sus movimientos volvieron a la lista de pendientes.", "error");
  }
  state.review.page = 1;
  renderReview();
}

function renderPendingReview() {
  let system = getPendingMovements("system");
  let bank = getPendingMovements("bank");
  if (state.review.search) {
    system = system.filter(item => movementSearchText(item).includes(state.review.search));
    bank = bank.filter(item => movementSearchText(item).includes(state.review.search));
  }
  system = sortPendingMovements(system, state.review.sort);
  bank = sortPendingMovements(bank, state.review.sort);
  const pageCount = Math.max(1, Math.ceil(Math.max(system.length, bank.length) / PAGE_SIZE));
  state.review.page = Math.min(state.review.page, pageCount);
  const start = (state.review.page - 1) * PAGE_SIZE;
  const systemPage = system.slice(start, start + PAGE_SIZE);
  const bankPage = bank.slice(start, start + PAGE_SIZE);
  const selectedSystemMovements = movementsFromIds("system", state.review.selectedSystem);
  const selectedBankMovements = movementsFromIds("bank", state.review.selectedBank);
  dom.reviewContent.innerHTML = `<div class="pending-wrap"><div class="manual-banner"><div><strong>Conciliación manual</strong><small>Seleccione al menos un movimiento de cada lado. La selección se conserva al cambiar de página.</small></div><button id="createManualBtn" class="button button-primary button-small" type="button" ${selectedSystemMovements.length && selectedBankMovements.length ? "" : "disabled"}><i data-lucide="link"></i> Crear conciliación (${selectedSystemMovements.length} ↔ ${selectedBankMovements.length})</button></div><div class="manual-grid">${renderPendingSide("system", systemPage, system)}${renderPendingSide("bank", bankPage, bank)}</div></div>`;
  dom.reviewContent.querySelectorAll("[data-pending-select]").forEach(checkbox => {
    checkbox.addEventListener("change", () => {
      const sourceKey = checkbox.dataset.source;
      const selection = sourceKey === "system" ? state.review.selectedSystem : state.review.selectedBank;
      if (checkbox.checked) selection.add(checkbox.value); else selection.delete(checkbox.value);
      renderPendingReview();
    });
  });
  document.getElementById("createManualBtn")?.addEventListener("click", createManualReconciliation);
  renderPagination(Math.max(system.length, bank.length), pageCount);
  refreshIcons(dom.reviewContent);
}

function renderPendingSide(sourceKey, pageItems, allItems) {
  const label = sourceKey === "system" ? (state.sources.system.name || "Sistema contable") : (state.sources.bank.name || "Caja o banco");
  const selection = sourceKey === "system" ? state.review.selectedSystem : state.review.selectedBank;
  const total = allItems.reduce((sum, item) => sum + comparisonAmount(item), 0);
  return `<section class="pending-side"><h3>${escapeHtml(label)} <span>${allItems.length} pendientes</span></h3><div class="table-scroll"><table class="data-table pending-table"><thead><tr><th></th><th>Fila</th><th>Fecha</th><th>Descripción</th><th>Tipo</th><th class="amount">Monto</th></tr></thead><tbody>${pageItems.length ? pageItems.map(item => `<tr class="${selection.has(item.id) ? "selected" : ""}"><td><input type="checkbox" data-pending-select data-source="${sourceKey}" value="${item.id}" ${selection.has(item.id) ? "checked" : ""} aria-label="Seleccionar fila ${item.row}"></td><td>${item.row}</td><td>${formatDate(item.date)}</td><td>${escapeHtml(item.description)}</td><td><span class="type-badge">${movementTypeLabel(item.type)}</span></td><td class="amount ${item.amount < 0 ? "negative" : ""}">${formatMoney(item.amount)}</td></tr>`).join("") : `<tr><td colspan="6"><div class="empty-state"><div><p>No hay movimientos pendientes en esta página.</p></div></div></td></tr>`}</tbody></table></div><div class="pending-total"><span>Total pendiente visible</span><strong>${formatMoney(total)}</strong></div></section>`;
}

function createManualReconciliation() {
  const systemMovements = movementsFromIds("system", state.review.selectedSystem);
  const bankMovements = movementsFromIds("bank", state.review.selectedBank);
  if (!systemMovements.length || !bankMovements.length) return;
  const reconciliation = calculateReconciliation(systemMovements, bankMovements, "manual");
  reconciliation.status = "confirmed";
  reconciliation.manual = true;
  reconciliation.ambiguous = false;
  reconciliation.criterion = "Conciliación creada manualmente por el usuario";
  reconciliation.reasons.push("La selección fue confirmada manualmente");
  reconciliation.observation = "Conciliación manual";
  addReconciliation(reconciliation);
  state.review.selectedSystem.clear();
  state.review.selectedBank.clear();
  showToast("Conciliación manual creada", `${reconciliation.id} fue agregada a conciliados.`, "success");
  renderReview();
}

function openReconciliationDetail(item) {
  document.getElementById("detailTitle").textContent = `Detalle ${item.id}`;
  document.getElementById("detailSubtitle").textContent = `${typeLabel(item)} · ${item.criterion}`;
  const content = document.getElementById("detailContent");
  content.innerHTML = `<div class="detail-grid">${renderDetailSide("Sistema contable", item.systemMovements)}${renderDetailSide("Caja o banco", item.bankMovements)}<div class="detail-summary"><div><span>Total sistema</span><strong>${formatMoney(item.totalSystem)}</strong></div><div><span>Total caja/banco</span><strong>${formatMoney(item.totalBank)}</strong></div><div><span>Diferencia</span><strong>${formatMoney(item.difference)}</strong></div><div><span>Confianza</span><strong>${item.score}/100</strong></div><div><span>Estado</span><strong>${item.status === "confirmed" ? "Conciliado" : "Posible"}</strong></div></div><div class="reason-list"><h3>Por qué se propuso</h3><ul>${item.reasons.map(reason => `<li>${escapeHtml(reason)}</li>`).join("")}${item.ambiguous ? "<li>La combinación es ambigua y requiere aprobación manual.</li>" : ""}</ul></div></div>`;
  refreshIcons(content);
  dom.detailDialog.showModal();
}

function renderDetailSide(title, movements) {
  return `<section class="detail-side"><h3>${escapeHtml(title)} · ${movements.length} movimiento(s)</h3>${movements.map(item => `<article class="detail-movement"><header><span>Fila ${item.row} · ${formatDate(item.date)}</span><strong class="amount ${item.amount < 0 ? "negative" : ""}">${formatMoney(item.amount)}</strong></header><p>${escapeHtml(item.description)}</p><small>${movementTypeLabel(item.type)}${item.status ? ` · Estado original: ${escapeHtml(item.status)}` : ""}</small></article>`).join("")}</section>`;
}

function openEditGroup(item) {
  state.review.editingId = item.id;
  const reservedByOthers = { system: new Set(), bank: new Set() };
  state.results.reconciliations.filter(other => other.id !== item.id && (other.status === "confirmed" || other.status === "possible")).forEach(other => {
    other.systemIds.forEach(id => reservedByOthers.system.add(id));
    other.bankIds.forEach(id => reservedByOthers.bank.add(id));
  });
  const availableSystem = state.sources.system.movements.filter(movement => !reservedByOthers.system.has(movement.id));
  const availableBank = state.sources.bank.movements.filter(movement => !reservedByOthers.bank.has(movement.id));
  const container = document.getElementById("editGroupContent");
  container.innerHTML = renderEditSide("system", availableSystem, new Set(item.systemIds)) + renderEditSide("bank", availableBank, new Set(item.bankIds));
  container.querySelectorAll("[data-group-select]").forEach(checkbox => checkbox.addEventListener("change", updateEditGroupTotals));
  updateEditGroupTotals();
  dom.editGroupDialog.showModal();
}

function renderEditSide(sourceKey, movements, selected) {
  const label = sourceKey === "system" ? "Sistema contable" : "Caja o banco";
  return `<section class="pending-side"><h3>${label}<span>${movements.length} disponibles</span></h3><div class="table-scroll"><table class="data-table pending-table"><thead><tr><th></th><th>Fila</th><th>Fecha</th><th>Descripción</th><th class="amount">Monto</th></tr></thead><tbody>${movements.map(item => `<tr><td><input data-group-select data-source="${sourceKey}" type="checkbox" value="${item.id}" ${selected.has(item.id) ? "checked" : ""}></td><td>${item.row}</td><td>${formatDate(item.date)}</td><td>${escapeHtml(item.description)}</td><td class="amount ${item.amount < 0 ? "negative" : ""}">${formatMoney(item.amount)}</td></tr>`).join("")}</tbody></table></div></section>`;
}

function updateEditGroupTotals() {
  const selected = getEditGroupSelections();
  const systemMovements = movementsFromIds("system", selected.system);
  const bankMovements = movementsFromIds("bank", selected.bank);
  const totalSystem = systemMovements.reduce((sum, item) => sum + comparisonAmount(item), 0);
  const totalBank = bankMovements.reduce((sum, item) => sum + comparisonAmount(item), 0);
  const difference = roundMoney(totalSystem - totalBank);
  document.getElementById("editGroupTotals").innerHTML = `<div><span>Total sistema</span><strong>${formatMoney(totalSystem)}</strong></div><div><span>Total caja / banco</span><strong>${formatMoney(totalBank)}</strong></div><div><span>Diferencia</span><strong>${formatMoney(difference)}</strong></div>`;
  document.getElementById("saveGroupBtn").disabled = !systemMovements.length || !bankMovements.length;
}

function getEditGroupSelections() {
  const selected = { system: new Set(), bank: new Set() };
  document.querySelectorAll("[data-group-select]:checked").forEach(input => selected[input.dataset.source].add(input.value));
  return selected;
}

function saveEditedGroup() {
  const current = findReconciliation(state.review.editingId);
  if (!current) return;
  const selected = getEditGroupSelections();
  const systemMovements = movementsFromIds("system", selected.system);
  const bankMovements = movementsFromIds("bank", selected.bank);
  if (!systemMovements.length || !bankMovements.length) return;
  const updated = calculateReconciliation(systemMovements, bankMovements);
  Object.assign(current, updated, {
    id: current.id,
    status: "possible",
    observation: current.observation,
    ambiguous: false,
    alternativeCount: 0,
    criterion: "Agrupación editada manualmente; requiere aprobación",
    manual: true
  });
  current.reasons.push("La agrupación original fue modificada manualmente");
  dom.editGroupDialog.close();
  showToast("Agrupación actualizada", "Revise los nuevos totales y apruebe la propuesta.", "success");
  renderReview();
}

function getEditGroupSelectionIds(sourceKey) {
  return new Set([...document.querySelectorAll(`[data-group-select][data-source="${sourceKey}"]:checked`)].map(input => input.value));
}

function movementsFromIds(sourceKey, ids) {
  const set = ids instanceof Set ? ids : new Set(ids);
  return state.sources[sourceKey].movements.filter(item => set.has(item.id));
}

function findReconciliation(id) {
  return state.results.reconciliations.find(item => item.id === id);
}

function renderPagination(totalItems, pageCount) {
  if (totalItems <= PAGE_SIZE) {
    dom.pagination.innerHTML = `<span>${totalItems.toLocaleString("es-UY")} registro(s)</span>`;
    return;
  }
  const pages = paginationPages(state.review.page, pageCount);
  dom.pagination.innerHTML = `<span>${totalItems.toLocaleString("es-UY")} registros · Página ${state.review.page} de ${pageCount}</span><div class="pagination-controls"><button type="button" data-page="${state.review.page - 1}" ${state.review.page === 1 ? "disabled" : ""} aria-label="Página anterior"><i data-lucide="chevron-left"></i></button>${pages.map(page => page === "…" ? `<span>…</span>` : `<button type="button" data-page="${page}" class="${page === state.review.page ? "active" : ""}">${page}</button>`).join("")}<button type="button" data-page="${state.review.page + 1}" ${state.review.page === pageCount ? "disabled" : ""} aria-label="Página siguiente"><i data-lucide="chevron-right"></i></button></div>`;
  dom.pagination.querySelectorAll("[data-page]").forEach(button => button.addEventListener("click", () => {
    state.review.page = Number(button.dataset.page);
    renderReviewContent();
  }));
  refreshIcons(dom.pagination);
}

function paginationPages(current, count) {
  if (count <= 7) return Array.from({ length: count }, (_, index) => index + 1);
  const values = new Set([1, count, current, current - 1, current + 1].filter(value => value >= 1 && value <= count));
  const sorted = [...values].sort((a, b) => a - b);
  const result = [];
  sorted.forEach((value, index) => {
    if (index && value - sorted[index - 1] > 1) result.push("…");
    result.push(value);
  });
  return result;
}

function sortReconciliations(items, sort) {
  const copy = [...items];
  const firstDate = item => Math.min(...[...item.systemMovements, ...item.bankMovements].map(movement => movement.date.getTime()));
  if (sort === "score-asc") copy.sort((a, b) => a.score - b.score);
  else if (sort === "date-desc") copy.sort((a, b) => firstDate(b) - firstDate(a));
  else if (sort === "date-asc") copy.sort((a, b) => firstDate(a) - firstDate(b));
  else if (sort === "difference-desc") copy.sort((a, b) => Math.abs(b.difference) - Math.abs(a.difference));
  else copy.sort((a, b) => b.score - a.score);
  return copy;
}

function sortPendingMovements(items, sort) {
  const copy = [...items];
  if (sort === "date-asc") copy.sort((a, b) => a.date - b.date || a.row - b.row);
  else if (sort === "difference-desc") copy.sort((a, b) => Math.abs(b.amount) - Math.abs(a.amount));
  else copy.sort((a, b) => b.date - a.date || a.row - b.row);
  return copy;
}

function reconciliationSearchText(item) {
  return [item.id, item.type, item.criterion, item.score, item.difference, item.observation, ...item.systemMovements.flatMap(m => [m.row, m.dateKey, m.description, m.amount]), ...item.bankMovements.flatMap(m => [m.row, m.dateKey, m.description, m.amount])].join(" ").toLowerCase();
}

function movementSearchText(item) {
  return [item.id, item.row, item.dateKey, formatDate(item.date), item.description, item.amount, item.type].join(" ").toLowerCase();
}

function summarizeMovements(movements) {
  if (movements.length === 1) return { title: movements[0].description, subtitle: `Fila ${movements[0].row} · ${formatDate(movements[0].date)}` };
  return { title: `${movements.length} movimientos`, subtitle: `${movements.map(item => `F${item.row}`).join(", ")} · ${formatDateRange(movements)}` };
}

function typeLabel(item) {
  if (item.manual || item.type === "manual") return "Manual";
  return { "one-to-one": "Uno a uno", "one-to-many": "Uno a varios", "many-to-one": "Varios a uno" }[item.type] || item.type;
}

function statusBadge(status) {
  return status === "confirmed"
    ? `<span class="status-badge confirmed"><i data-lucide="check"></i> Conciliado</span>`
    : `<span class="status-badge possible"><i data-lucide="circle-help"></i> Posible</span>`;
}

function scoreClass(score) {
  return score >= state.config.autoThreshold ? "high" : score >= state.config.possibleThreshold ? "medium" : "low";
}

function movementTypeLabel(type) {
  return type === "debit" ? "Débito" : type === "credit" ? "Crédito" : "Mixto";
}

function renderExportSummary() {
  const summary = calculateSummary();
  document.getElementById("exportSummary").innerHTML = [
    ["Conciliaciones", summary.confirmedReconciliations],
    ["Posibles", summary.possibleReconciliations],
    ["Pendientes", summary.pendingCount],
    ["Importe conciliado", formatMoney(summary.reconciledAmount)],
    ["Diferencia pendiente", formatMoney(summary.pendingDifference)],
    ["Porcentaje conciliado", `${formatDecimal(summary.percentage, 1)}%`]
  ].map(([label, value]) => `<div class="export-stat"><span>${escapeHtml(label)}</span><strong>${escapeHtml(String(value))}</strong></div>`).join("");
}

function exportWorkbook() {
  if (typeof XLSX === "undefined") {
    showToast("No se puede generar el Excel", "El componente de Excel no está disponible.", "error");
    return;
  }
  try {
    const workbook = XLSX.utils.book_new();
    workbook.Props = {
      Title: "Resultado de conciliación contable",
      Subject: "Conciliación de movimientos",
      Author: "Aplicación de conciliación contable",
      CreatedDate: new Date()
    };
    const summary = calculateSummary();
    const processingDate = state.results.processingAt || new Date();
    const summaryRows = [
      ["RESUMEN DE CONCILIACIÓN CONTABLE", ""],
      ["Fecha y hora de procesamiento", formatDateTime(processingDate)],
      ["Tabla del sistema", state.sources.system.name || state.sources.system.label],
      ["Tabla de caja o banco", state.sources.bank.name || state.sources.bank.label],
      ["Movimientos del sistema", summary.totalSystem],
      ["Movimientos de caja o banco", summary.totalBank],
      ["Movimientos conciliados", summary.confirmedCount],
      ["Movimientos en posibles conciliaciones", summary.possibleCount],
      ["Movimientos pendientes", summary.pendingCount],
      ["Importe conciliado", summary.reconciledAmount],
      ["Diferencia pendiente", summary.pendingDifference],
      ["Porcentaje de conciliación", summary.percentage / 100],
      ["", ""],
      ["PARÁMETROS UTILIZADOS", ""],
      ["Tolerancia de fechas (días)", state.config.dateTolerance],
      ["Tolerancia absoluta de monto", state.config.amountAbsTolerance],
      ["Tolerancia porcentual de monto", state.config.amountPercentTolerance / 100],
      ["Máximo por agrupación", state.config.maxGroupSize],
      ["Máximo de combinaciones", state.config.maxCombinations],
      ["Comparar signos obligatoriamente", yesNo(state.config.requireSameSign)],
      ["Convención invertida entre tablas", yesNo(state.config.invertBetweenTables)],
      ["Nivel automático", state.config.autoThreshold],
      ["Nivel posible", state.config.possibleThreshold],
      ["Combinaciones evaluadas", state.results.evaluatedCombinations],
      ["Límite de combinaciones alcanzado", yesNo(state.results.combinationLimitReached)]
    ];
    const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
    styleSummarySheet(summarySheet);
    appendSheet(workbook, summarySheet, "Resumen");

    const reconciliationHeaders = ["ID de conciliación", "Tipo", "Filas del sistema", "Filas de caja o banco", "Fechas del sistema", "Fechas de caja o banco", "Descripciones del sistema", "Descripciones de caja o banco", "Monto del sistema", "Monto de caja o banco", "Monto original caja o banco", "Diferencia", "Puntaje", "Estado", "Criterio", "Observaciones", "Fecha y hora de procesamiento"];
    const confirmed = state.results.reconciliations.filter(item => item.status === "confirmed");
    const possible = state.results.reconciliations.filter(item => item.status === "possible");
    const reconciliationSheet = createStyledDataSheet(reconciliationHeaders, confirmed.map(reconciliationExportRow), { fill: "E9F5ED", numericColumns: [8, 9, 10, 11, 12] });
    appendSheet(workbook, reconciliationSheet, "Conciliaciones");
    const possibleSheet = createStyledDataSheet(reconciliationHeaders, possible.map(reconciliationExportRow), { fill: "FFF7DC", numericColumns: [8, 9, 10, 11, 12] });
    appendSheet(workbook, possibleSheet, "Posibles conciliaciones");

    const pendingHeaders = ["Fila original", "Fecha", "Descripción", "Monto firmado", "Tipo", "Estado original"];
    const pendingSystemSheet = createStyledDataSheet(pendingHeaders, summary.pendingSystem.map(pendingExportRow), { fill: "FBEEEE", numericColumns: [3] });
    appendSheet(workbook, pendingSystemSheet, "Pendientes del sistema");
    const pendingBankSheet = createStyledDataSheet(pendingHeaders, summary.pendingBank.map(pendingExportRow), { fill: "FBEEEE", numericColumns: [3] });
    appendSheet(workbook, pendingBankSheet, "Pendientes de caja o banco");

    appendSheet(workbook, createOriginalDataSheet(state.sources.system), "Datos originales del sistema");
    appendSheet(workbook, createOriginalDataSheet(state.sources.bank), "Datos originales caja o banco");

    const allErrors = [...state.sources.system.invalidRows, ...state.sources.bank.invalidRows];
    if (allErrors.length) {
      const errorRows = allErrors.map(item => [item.source, item.sheet, item.row, item.errors, item.values.map(displayOriginalValue).join(" | ")]);
      appendSheet(workbook, createStyledDataSheet(["Origen", "Hoja", "Fila", "Error", "Datos originales"], errorRows, { fill: "FBEEEE", numericColumns: [2] }), "Errores de importación");
    }
    const output = XLSX.write(workbook, { compression: true, bookType: "xlsx", type: "array", cellStyles: true });
    downloadBlob(
      new Blob([output], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
      `conciliacion_${toFileTimestamp()}.xlsx`
    );
    showToast("Excel generado", "El libro de conciliación se descargó correctamente.", "success");
  } catch (error) {
    console.error(error);
    showToast("No se pudo generar el Excel", error.message || "Revise los datos y vuelva a intentarlo.", "error", 9000);
  }
}

function reconciliationExportRow(item) {
  return [
    item.id,
    typeLabel(item),
    item.systemMovements.map(movement => movement.row).join(", "),
    item.bankMovements.map(movement => movement.row).join(", "),
    item.systemMovements.map(movement => formatDate(movement.date)).join(" | "),
    item.bankMovements.map(movement => formatDate(movement.date)).join(" | "),
    item.systemMovements.map(movement => movement.description).join(" | "),
    item.bankMovements.map(movement => movement.description).join(" | "),
    item.totalSystem,
    item.totalBank,
    item.totalBankOriginal,
    item.difference,
    item.score,
    item.status === "confirmed" ? "Conciliado" : "Posible",
    item.criterion,
    item.observation,
    formatDateTime(state.results.processingAt || item.createdAt)
  ];
}

function pendingExportRow(item) {
  return [item.row, formatDate(item.date), item.description, item.amount, movementTypeLabel(item.type), item.status];
}

function createOriginalDataSheet(source) {
  const rows = [source.headers, ...source.rows.map(row => row.values.map(value => value instanceof Date ? formatDate(value) : value))];
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  applyTableSheetStyle(worksheet, { fill: "FFFFFF", numericColumns: [] });
  return worksheet;
}

function createStyledDataSheet(headers, rows, options = {}) {
  const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  applyTableSheetStyle(worksheet, options);
  return worksheet;
}

function applyTableSheetStyle(worksheet, { fill = "FFFFFF", numericColumns = [] } = {}) {
  const range = worksheet["!ref"] ? XLSX.utils.decode_range(worksheet["!ref"]) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  const headerStyle = {
    fill: { fgColor: { rgb: "164E55" } },
    font: { bold: true, color: { rgb: "FFFFFF" } },
    alignment: { vertical: "center", wrapText: true },
    border: { bottom: { style: "thin", color: { rgb: "A9B7B9" } } }
  };
  for (let column = range.s.c; column <= range.e.c; column++) {
    const cell = worksheet[XLSX.utils.encode_cell({ r: 0, c: column })];
    if (cell) cell.s = headerStyle;
  }
  for (let row = 1; row <= range.e.r; row++) {
    for (let column = range.s.c; column <= range.e.c; column++) {
      const address = XLSX.utils.encode_cell({ r: row, c: column });
      const cell = worksheet[address];
      if (!cell) continue;
      cell.s = {
        fill: { fgColor: { rgb: row % 2 ? fill : mixHexWithWhite(fill, .55) } },
        alignment: { vertical: "top", wrapText: column === 4 || column === 5 || column === 6 || column === 7 },
        border: { bottom: { style: "hair", color: { rgb: "DDE3E4" } } }
      };
      if (numericColumns.includes(column)) {
        cell.z = column === 12 ? "0" : "#,##0.00;[Red]-#,##0.00";
      }
    }
  }
  worksheet["!autofilter"] = { ref: XLSX.utils.encode_range({ r: 0, c: range.s.c }, { r: range.e.r, c: range.e.c }) };
  worksheet["!freeze"] = { xSplit: 0, ySplit: 1, topLeftCell: "A2", activePane: "bottomLeft", state: "frozen" };
  worksheet["!rows"] = [{ hpt: 28 }];
  worksheet["!cols"] = estimateColumnWidths(worksheet, range);
}

function styleSummarySheet(worksheet) {
  const range = XLSX.utils.decode_range(worksheet["!ref"]);
  worksheet["!cols"] = [{ wch: 40 }, { wch: 28 }];
  worksheet["!freeze"] = { xSplit: 0, ySplit: 1, topLeftCell: "A2", activePane: "bottomLeft", state: "frozen" };
  for (let row = range.s.r; row <= range.e.r; row++) {
    const labelCell = worksheet[XLSX.utils.encode_cell({ r: row, c: 0 })];
    const valueCell = worksheet[XLSX.utils.encode_cell({ r: row, c: 1 })];
    if (!labelCell) continue;
    const section = row === 0 || row === 13;
    labelCell.s = section
      ? { fill: { fgColor: { rgb: "164E55" } }, font: { bold: true, color: { rgb: "FFFFFF" }, sz: row === 0 ? 14 : 11 }, alignment: { vertical: "center" } }
      : { fill: { fgColor: { rgb: row % 2 ? "F4F7F7" : "FFFFFF" } }, font: { bold: true, color: { rgb: "405153" } }, border: { bottom: { style: "hair", color: { rgb: "DDE3E4" } } } };
    if (valueCell) valueCell.s = section
      ? { fill: { fgColor: { rgb: "164E55" } } }
      : { fill: { fgColor: { rgb: row % 2 ? "F4F7F7" : "FFFFFF" } }, border: { bottom: { style: "hair", color: { rgb: "DDE3E4" } } } };
  }
  if (worksheet.B10) worksheet.B10.z = "#,##0.00;[Red]-#,##0.00";
  if (worksheet.B11) worksheet.B11.z = "#,##0.00;[Red]-#,##0.00";
  if (worksheet.B12) worksheet.B12.z = "0.0%";
  if (worksheet.B16) worksheet.B16.z = "#,##0.00";
  if (worksheet.B17) worksheet.B17.z = "0.00%";
}

function estimateColumnWidths(worksheet, range) {
  return Array.from({ length: range.e.c - range.s.c + 1 }, (_, index) => {
    let maximum = 10;
    for (let row = range.s.r; row <= Math.min(range.e.r, 200); row++) {
      const cell = worksheet[XLSX.utils.encode_cell({ r: row, c: index })];
      const text = cell ? String(cell.v ?? "") : "";
      maximum = Math.max(maximum, Math.min(60, text.length + 2));
    }
    return { wch: maximum };
  });
}

function appendSheet(workbook, worksheet, name) {
  XLSX.utils.book_append_sheet(workbook, worksheet, name.slice(0, 31));
}

function showToast(title, message, type = "success", duration = 4200) {
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.innerHTML = `<i data-lucide="${type === "success" ? "circle-check" : "triangle-alert"}"></i><div><strong>${escapeHtml(title)}</strong><span>${escapeHtml(message)}</span></div>`;
  dom.toastRegion.appendChild(toast);
  refreshIcons(toast);
  window.setTimeout(() => toast.remove(), duration);
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function displayOriginalValue(value) {
  if (value instanceof Date) return formatDate(value);
  return String(value ?? "");
}

function formatMoney(value) {
  return new Intl.NumberFormat("es-UY", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(Number(value) || 0);
}

function formatDecimal(value, digits = 2) {
  return new Intl.NumberFormat("es-UY", { minimumFractionDigits: digits, maximumFractionDigits: digits }).format(Number(value) || 0);
}

function formatDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return new Intl.DateTimeFormat("es-UY", { timeZone: "UTC", day: "2-digit", month: "2-digit", year: "numeric" }).format(date);
}

function formatDateTime(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return new Intl.DateTimeFormat("es-UY", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit", second: "2-digit" }).format(date);
}

function formatDateRange(movements) {
  const dates = movements.map(item => item.date).sort((a, b) => a - b);
  return dates[0].getTime() === dates.at(-1).getTime() ? formatDate(dates[0]) : `${formatDate(dates[0])}–${formatDate(dates.at(-1))}`;
}

function formatFileSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${formatDecimal(bytes / 1024, 1)} KB`;
  return `${formatDecimal(bytes / (1024 * 1024), 1)} MB`;
}

function toDateKey(date) {
  return date.toISOString().slice(0, 10);
}

function toFileTimestamp() {
  const date = new Date();
  const pad = value => String(value).padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}_${pad(date.getHours())}-${pad(date.getMinutes())}`;
}

function daysBetween(a, b) {
  return Math.abs(Math.round((a.getTime() - b.getTime()) / DATE_MS));
}

function roundMoney(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}

function clamp(value, minimum, maximum) {
  return Math.min(maximum, Math.max(minimum, value));
}

function groupBy(items, keyFunction) {
  const groups = new Map();
  items.forEach(item => {
    const key = keyFunction(item);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(item);
  });
  return groups;
}

function columnLetter(number) {
  let value = number;
  let result = "";
  while (value > 0) {
    value--;
    result = String.fromCharCode(65 + value % 26) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function csvEscape(value) {
  const text = displayOriginalValue(value);
  return /[;"\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function yesNo(value) {
  return value ? "Sí" : "No";
}

function mixHexWithWhite(hex, ratio) {
  const clean = hex.replace("#", "");
  const channels = [0, 2, 4].map(index => parseInt(clean.slice(index, index + 2), 16));
  return channels.map(channel => Math.round(channel + (255 - channel) * ratio).toString(16).padStart(2, "0")).join("").toUpperCase();
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>'"]/g, character => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" })[character]);
}

function escapeAttribute(value) {
  return escapeHtml(value).replace(/`/g, "&#96;");
}

function delay(milliseconds) {
  return new Promise(resolve => setTimeout(resolve, milliseconds));
}

function yieldToMain() {
  return delay(0);
}

window.ReconciliationApp = Object.freeze({
  DEFAULT_CONFIG,
  parseAmount,
  parseDateValue,
  normalizeDescription,
  descriptionSimilarity,
  calculateReconciliation,
  enumerateGroupCandidates,
  normalizeSourceRow,
  getState: () => state
});
