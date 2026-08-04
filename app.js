"use strict";

const APP_BUILD = "2026.08.04-1";
const STATE_SCHEMA_VERSION = 1;
const STATE_SHEET_NAME = "Estado ConciliApp";
const STATE_CHUNK_SIZE = 30000;
const LOCAL_DATABASE_NAME = "ConciliAppLocal";
const LOCAL_DATABASE_VERSION = 1;
const LOCAL_STATE_STORE = "reconciliations";
const LOCAL_STATE_KEY = "current";

const DEFAULT_CONFIG = Object.freeze({
  dateTolerance: 1,
  amountAbsTolerance: 0.10,
  amountPercentTolerance: 0,
  maxGroupSize: 100,
  maxPairComparisons: 2000000,
  maxCombinations: 25000,
  requireSameSign: true,
  invertBetweenTables: false,
  searchOneToOne: true,
  searchOneToMany: true,
  searchManyToOne: true,
  allowMixedGroupSigns: false,
  searchInternalOffsets: false,
  considerDescription: true,
  autoThreshold: 70,
  possibleThreshold: 55,
  ignoreCase: true,
  ignoreAccents: true,
  ignorePunctuation: true,
  ignoredWords: "pago, recibo, transferencia",
  excludeReconciled: false,
  relaxedDescriptionPriority: false
});

const MAX_FILE_SIZE = 25 * 1024 * 1024;
const PAGE_SIZE = 50;
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
const HEADER_ALIASES = Object.freeze({
  date: ["fecha", "date", "fec", "fecha movimiento", "fecha valor"],
  description: ["descripcion", "description", "concepto", "detalle", "glosa", "referencia", "movimiento"],
  amount: ["monto", "importe", "amount", "saldo movimiento", "valor"],
  debit: ["debito", "debe", "debit", "egreso", "retiro", "cargo", "pagos"],
  credit: ["credito", "haber", "credit", "ingreso", "deposito", "abono", "ingresos"],
  status: ["estado", "status", "conciliado", "situacion"]
});

const state = {
  step: 1,
  maxVisitedStep: 1,
  sources: {
    system: createEmptySource("system", "Sistema contable"),
    bank: createEmptySource("bank", "Caja o banco")
  },
  config: { ...DEFAULT_CONFIG },
  results: createEmptyResults(),
  processing: { cancelled: false, running: false, worker: null, jobId: null },
  review: {
    tab: "confirmed",
    search: "",
    type: "all",
    sort: "score-desc",
    page: 1,
    selectedSystem: new Set(),
    selectedBank: new Set(),
    editingId: null,
    editAvailableSystem: [],
    editAvailableBank: [],
    editSelectedSystem: new Set(),
    editSelectedBank: new Set(),
    editSearchSystem: "",
    editSearchBank: "",
    rejectedSignatures: new Set(),
    rejectedProposals: [],
    periodFilter: { from: "", to: "", appliedAt: null }
  },
  workspace: createWorkspaceState(),
  accountTransfer: createAccountTransferState(),
  transferLog: [],
  persistence: { restoring: false, saveTimer: null, saveErrorShown: false, lastSavedAt: null }
};

function createWorkspaceState(id = "", name = "") {
  return {
    id: id || `workspace-${Date.now()}-${Math.random().toString(16).slice(2)}`,
    name: String(name || "").trim()
  };
}

function createAccountTransferState() {
  return {
    currentSnapshot: null,
    destinationSnapshot: null,
    destinationFileName: "",
    currentSearch: "",
    destinationSearch: "",
    selectedCurrent: new Set(),
    selectedDestination: new Set(),
    dragPayload: null
  };
}

function createEmptySource(key, label) {
  return {
    key,
    label,
    name: "",
    file: null,
    workbook: null,
    sheetNames: [],
    selectedSheet: "",
    importRangeInfo: null,
    matrix: [],
    headerRowIndex: 0,
    headerRowNumber: 1,
    dataStartRow: 2,
    dataEndRow: "",
    headers: [],
    rows: [],
    detectedBlocks: [],
    columnBlock: "all",
    dateFrom: "",
    dateTo: "",
    reportPeriod: null,
    periodSource: "",
    excludedDescriptions: "saldo inicial, totales, total",
    filteredRowsCount: 0,
    formatMode: "auto",
    detectedFormat: "signed",
    positiveMeaning: "debit",
    splitConvention: "preserve",
    splitConventionLocked: false,
    allowBoth: false,
    mapping: { date: "", description: "", amount: "", debit: "", credit: "", status: "" },
    movements: [],
    invalidRows: [],
    validationErrors: [],
    validationWarnings: [],
    isValid: false,
    restoredState: false,
    restoredRowCount: 0
  };
}

function createEmptyResults() {
  return {
    reconciliations: [],
    processingAt: null,
    evaluatedPairs: 0,
    candidatePairs: 0,
    evaluatedCombinations: 0,
    pairLimitReached: false,
    combinationLimitReached: false,
    limitedGroupAnchors: 0,
    engineMode: "",
    retryPasses: [],
    nextId: 1
  };
}

const dom = {};

document.addEventListener("DOMContentLoaded", initializeApplication);

async function initializeApplication() {
  console.info(`ConciliApp ${APP_BUILD}`);
  cacheDom();
  bindDateDisplayInputs();
  bindGlobalEvents();
  bindSourceEditor("system");
  bindSourceEditor("bank");
  populateConfigForm();
  renderProgress();
  refreshIcons();
  if (typeof XLSX === "undefined") {
    showToast("No se pudo cargar el componente de Excel", "Verifique la conexión a Internet y vuelva a abrir el archivo.", "error", 9000);
  }
  await restorePersistedStateOnStartup();
}

function bindDateDisplayInputs() {
  document.querySelectorAll("[data-date-display]").forEach(input => {
    input.addEventListener("input", () => {
      const digits = input.value.replace(/\D/g, "").slice(0, 8);
      input.value = [digits.slice(0, 2), digits.slice(2, 4), digits.slice(4, 8)].filter(Boolean).join("/");
      input.setCustomValidity("");
    });
    input.addEventListener("blur", () => {
      const parsed = parseDisplayDateInput(input.value);
      if (parsed === null) {
        input.setCustomValidity("Use el formato dd/mm/aaaa o dd/mm/aa e indique una fecha válida.");
        return;
      }
      input.setCustomValidity("");
      if (parsed) input.value = formatDateInput(parsed);
    });
  });
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
  dom.retryDialog = document.getElementById("retryDialog");
  dom.retryForm = document.getElementById("retryForm");
  dom.periodTrimDialog = document.getElementById("periodTrimDialog");
  dom.periodTrimForm = document.getElementById("periodTrimForm");
  dom.accountTransferDialog = document.getElementById("accountTransferDialog");
  dom.accountTransferContent = document.getElementById("accountTransferContent");
  dom.accountTransferInput = document.getElementById("accountTransferInput");
  dom.savedAccountSelect = document.getElementById("savedAccountSelect");
  dom.toastRegion = document.getElementById("toastRegion");
  dom.continueButtons = {
    system: document.getElementById("continueSystemBtn"),
    bank: document.getElementById("continueBankBtn")
  };
}

function bindGlobalEvents() {
  document.getElementById("startBtn").addEventListener("click", () => goToStep(2));
  document.getElementById("resumeReconciliationBtn").addEventListener("click", () => document.getElementById("resumeReconciliationInput").click());
  document.getElementById("resumeReconciliationInput").addEventListener("change", async event => {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (file) await loadPreviousReconciliationFile(file);
  });
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
  document.getElementById("cancelProcessBtn").addEventListener("click", requestProcessingCancellation);
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
      syncReviewControls();
      renderReview();
    });
  });
  document.getElementById("reviewSearch").addEventListener("input", event => {
    state.review.search = event.target.value.trim().toLowerCase();
    state.review.page = 1;
    syncReviewControls();
    renderReviewContent();
    refreshIcons(document.querySelector(".review-toolbar"));
  });
  document.getElementById("typeFilter").addEventListener("change", event => {
    state.review.type = event.target.value;
    state.review.page = 1;
    syncReviewControls();
    renderReviewContent();
    refreshIcons(document.querySelector(".review-toolbar"));
  });
  document.getElementById("sortResults").addEventListener("change", event => {
    state.review.sort = event.target.value;
    state.review.page = 1;
    renderReviewContent();
  });
  document.getElementById("approveAllPossibleBtn").addEventListener("click", approveAllPossible);
  document.getElementById("rejectAllPossibleBtn").addEventListener("click", rejectAllPossible);
  document.getElementById("retryPendingBtn").addEventListener("click", openRetryDialog);
  document.getElementById("trimPeriodBtn").addEventListener("click", openPeriodTrimDialog);
  document.getElementById("manageAccountsBtn").addEventListener("click", openAccountTransferDialog);
  document.getElementById("loadAccountTransferBtn").addEventListener("click", () => dom.accountTransferInput.click());
  dom.accountTransferInput.addEventListener("change", async event => {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (file) await loadDestinationAccountFile(file);
  });
  document.getElementById("loadSavedAccountBtn").addEventListener("click", loadSelectedSavedAccount);
  document.getElementById("moveToDestinationBtn").addEventListener("click", () => moveSelectedAccountMovements("current", "destination"));
  document.getElementById("moveToCurrentBtn").addEventListener("click", () => moveSelectedAccountMovements("destination", "current"));
  document.getElementById("saveAccountTransfersBtn").addEventListener("click", saveAccountTransfers);
  document.getElementById("openDestinationAccountBtn").addEventListener("click", openDestinationAccount);
  document.querySelectorAll("[data-close-account-transfer]").forEach(button => button.addEventListener("click", () => dom.accountTransferDialog.close()));
  document.querySelectorAll("[data-close-retry]").forEach(button => button.addEventListener("click", () => dom.retryDialog.close()));
  document.querySelectorAll("[data-close-period]").forEach(button => button.addEventListener("click", () => dom.periodTrimDialog.close()));
  dom.retryForm.addEventListener("submit", event => {
    event.preventDefault();
    const options = readRetryOptions();
    dom.retryDialog.close();
    retryPendingReconciliation(options);
  });
  dom.periodTrimForm.addEventListener("input", updatePeriodTrimPreview);
  dom.periodTrimForm.addEventListener("submit", applyPeriodTrim);
  document.getElementById("restorePeriodBtn").addEventListener("click", restorePeriodExclusions);
  document.querySelectorAll("[data-close-detail]").forEach(button => button.addEventListener("click", () => dom.detailDialog.close()));
  document.querySelectorAll("[data-close-group]").forEach(button => button.addEventListener("click", () => dom.editGroupDialog.close()));
  document.getElementById("saveGroupBtn").addEventListener("click", saveEditedGroup);
  document.getElementById("approveGroupBtn").addEventListener("click", approveEditedGroup);
  document.getElementById("rejectGroupBtn").addEventListener("click", rejectEditedGroup);
  window.addEventListener("keydown", event => {
    if (event.key === "Escape") {
      if (dom.detailDialog.open) dom.detailDialog.close();
      if (dom.editGroupDialog.open) dom.editGroupDialog.close();
      if (dom.retryDialog.open) dom.retryDialog.close();
      if (dom.periodTrimDialog.open) dom.periodTrimDialog.close();
      if (dom.accountTransferDialog.open) dom.accountTransferDialog.close();
    }
  });
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden" && !state.persistence.restoring) persistCurrentState();
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
    headerRow: editor.querySelector("[data-header-row]"),
    dataStartRow: editor.querySelector("[data-data-start-row]"),
    dataEndRow: editor.querySelector("[data-data-end-row]"),
    columnBlockField: editor.querySelector("[data-column-block-field]"),
    columnBlock: editor.querySelector("[data-column-block]"),
    dateFrom: editor.querySelector("[data-date-from]"),
    dateTo: editor.querySelector("[data-date-to]"),
    excludedDescriptions: editor.querySelector("[data-excluded-descriptions]"),
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
  refs.nameInput.addEventListener("input", event => { source.name = event.target.value.trim(); scheduleStatePersistence(900); });
  refs.formatMode.addEventListener("change", event => {
    source.formatMode = event.target.value;
    toggleFormatOptions(editor, source);
    if (source.workbook) { autoMapColumns(source); renderSourceEditor(sourceKey); }
    else scheduleStatePersistence();
  });
  refs.positiveMeaning.addEventListener("change", event => { source.positiveMeaning = event.target.value; validateSource(sourceKey); });
  refs.splitConvention.addEventListener("change", event => {
    source.splitConvention = event.target.value;
    source.splitConventionLocked = true;
    validateSource(sourceKey);
    renderImportDetectionNote(source, refs.detectionNote);
  });
  refs.allowBoth.addEventListener("change", event => { source.allowBoth = event.target.checked; validateSource(sourceKey); });
  refs.sheetSelect.addEventListener("change", event => {
    source.selectedSheet = event.target.value;
    extractSelectedSheet(source);
    autoMapColumns(source);
    renderSourceEditor(sourceKey);
  });
  refs.headerRow.addEventListener("change", event => {
    const next = clamp(Math.trunc(Number(event.target.value) || 1), 1, Math.max(1, source.matrix.length));
    source.headerRowNumber = next;
    source.dataStartRow = next + 1;
    applySelectedTableRange(source);
    autoMapColumns(source);
    renderSourceEditor(sourceKey);
  });
  refs.dataStartRow.addEventListener("change", event => {
    source.dataStartRow = clamp(Math.trunc(Number(event.target.value) || source.headerRowNumber + 1), 1, Math.max(1, source.matrix.length));
    applySelectedTableRange(source);
    renderSourceEditor(sourceKey);
  });
  refs.dataEndRow.addEventListener("change", event => {
    const value = String(event.target.value || "").trim();
    source.dataEndRow = value ? clamp(Math.trunc(Number(value)), source.dataStartRow, Math.max(source.dataStartRow, source.matrix.length)) : "";
    applySelectedTableRange(source);
    renderSourceEditor(sourceKey);
  });
  refs.columnBlock.addEventListener("change", event => {
    source.columnBlock = event.target.value;
    autoMapColumns(source);
    renderSourceEditor(sourceKey);
  });
  refs.dateFrom.addEventListener("change", event => applySourceDateFilter(sourceKey, "dateFrom", event.target));
  refs.dateTo.addEventListener("change", event => applySourceDateFilter(sourceKey, "dateTo", event.target));
  refs.excludedDescriptions.addEventListener("change", event => { source.excludedDescriptions = event.target.value.trim(); validateSource(sourceKey); });
  toggleFormatOptions(editor, source);
}

function applySourceDateFilter(sourceKey, field, input) {
  const parsed = parseDisplayDateInput(input.value);
  if (parsed === null) {
    input.setCustomValidity("Use el formato dd/mm/aaaa o dd/mm/aa e indique una fecha válida.");
    input.reportValidity();
    return;
  }
  input.setCustomValidity("");
  if (parsed) input.value = formatDateInput(parsed);
  const source = state.sources[sourceKey];
  source[field] = parsed;
  validateSource(sourceKey);
  renderImportDetectionNote(source, document.querySelector(`[data-source-editor="${sourceKey}"]`)._refs.detectionNote);
}

function toggleFormatOptions(editor, source) {
  const effective = source.formatMode === "auto" ? source.detectedFormat : source.formatMode;
  editor.querySelectorAll(".signed-option").forEach(node => node.classList.toggle("hidden", effective !== "signed"));
  editor.querySelectorAll(".split-option").forEach(node => node.classList.toggle("hidden", effective !== "split"));
}

async function loadSourceFile(sourceKey, file) {
  const source = state.sources[sourceKey];
  const editorName = document.querySelector(`[data-source-editor="${sourceKey}"] [data-source-name]`)?.value.trim();
  if (editorName) source.name = editorName;
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
      cellDates: false,
      dense: false,
      raw: isCsv,
      codepage: 65001
    });
    if (!workbook.SheetNames.length) throw new Error("El archivo no contiene hojas legibles.");
    clearReconciliationResultsForSourceChange();
    Object.assign(source, {
      file,
      workbook,
      sheetNames: [...workbook.SheetNames],
      selectedSheet: selectInitialSheet(workbook),
      importRangeInfo: null,
      headerRowNumber: 1,
      dataStartRow: 2,
      dataEndRow: "",
      detectedBlocks: [],
      columnBlock: "all",
      dateFrom: "",
      dateTo: "",
      reportPeriod: null,
      periodSource: "",
      filteredRowsCount: 0,
      invalidRows: [],
      validationErrors: [],
      validationWarnings: [],
      isValid: false,
      restoredState: false,
      restoredRowCount: 0
    });
    extractSelectedSheet(source);
    autoMapColumns(source);
    renderSourceEditor(sourceKey);
    scheduleStatePersistence();
    showToast("Archivo cargado", `${file.name} se leyó correctamente.`, "success");
  } catch (error) {
    console.error(error);
    showToast("No se pudo leer el archivo", error.message || "El contenido no parece ser un Excel o CSV válido.", "error", 8000);
  }
}

function extractSelectedSheet(source) {
  const sheet = source.workbook.Sheets[source.selectedSheet];
  const importRangeInfo = findEffectiveWorksheetRange(sheet);
  const matrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: true,
    defval: "",
    blankrows: true,
    range: importRangeInfo.range
  });
  const headerRowIndex = detectHeaderRow(matrix);
  source.importRangeInfo = importRangeInfo;
  if (headerRowIndex < 0) {
    source.matrix = [];
    source.headers = [];
    source.rows = [];
    return;
  }
  source.matrix = matrix;
  source.headerRowNumber = headerRowIndex + 1;
  source.dataStartRow = headerRowIndex + 2;
  source.dataEndRow = "";
  source.columnBlock = "all";
  source.reportPeriod = detectReportPeriod(matrix, headerRowIndex);
  if (source.reportPeriod) {
    source.dateFrom = source.reportPeriod.from;
    source.dateTo = source.reportPeriod.to;
    source.periodSource = "archivo";
  } else if (source.key === "bank") {
    const system = state.sources.system;
    const from = system.reportPeriod?.from || system.dateFrom;
    const to = system.reportPeriod?.to || system.dateTo;
    if (from && to) {
      source.dateFrom = from;
      source.dateTo = to;
      source.periodSource = "sistema";
    }
  }
  applySelectedTableRange(source);
}

function findEffectiveWorksheetRange(sheet) {
  const declaredReference = sheet?.["!ref"] || "A1:A1";
  let declaredRange;
  try {
    declaredRange = XLSX.utils.decode_range(declaredReference);
  } catch {
    declaredRange = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  }
  let lastSubstantiveRow = 0;
  let lastSubstantiveColumn = 0;
  let substantiveCells = 0;

  // Algunos Excel guardan ceros o formato hasta la fila 1.048.576. Recorrer
  // las celdas dispersas evita materializar ese rango como un millón de filas.
  for (const address in sheet) {
    if (address[0] === "!") continue;
    const cell = sheet[address];
    if (!isSubstantiveWorksheetCell(cell)) continue;
    let position;
    try {
      position = XLSX.utils.decode_cell(address);
    } catch {
      continue;
    }
    lastSubstantiveRow = Math.max(lastSubstantiveRow, position.r);
    lastSubstantiveColumn = Math.max(lastSubstantiveColumn, position.c);
    substantiveCells++;
  }

  const range = {
    s: { r: 0, c: 0 },
    e: { r: lastSubstantiveRow, c: lastSubstantiveColumn }
  };
  return {
    range,
    declaredReference,
    effectiveReference: XLSX.utils.encode_range(range),
    declaredRows: declaredRange.e.r - declaredRange.s.r + 1,
    effectiveRows: lastSubstantiveRow + 1,
    ignoredTrailingRows: Math.max(0, declaredRange.e.r - lastSubstantiveRow),
    substantiveCells
  };
}

function isSubstantiveWorksheetCell(cell) {
  if (!cell || cell.v === null || cell.v === undefined || cell.v === "") return false;
  if (typeof cell.v === "number") return Number.isFinite(cell.v) ? Math.abs(cell.v) > 1e-12 : true;
  return true;
}

function applySelectedTableRange(source) {
  if (!source.matrix.length) return;
  const headerRowIndex = clamp(source.headerRowNumber - 1, 0, source.matrix.length - 1);
  const rawHeaders = source.matrix[headerRowIndex] || [];
  const endRowIndex = source.dataEndRow ? Math.min(Number(source.dataEndRow), source.matrix.length) : source.matrix.length;
  const dataStartIndex = clamp(source.dataStartRow - 1, 0, source.matrix.length);
  const relevantStartIndex = Math.min(headerRowIndex, dataStartIndex);
  let width = rawHeaders.length;
  for (let index = relevantStartIndex; index < endRowIndex; index++) {
    const row = source.matrix[index];
    if (Array.isArray(row) && row.length > width) width = row.length;
  }
  const headers = Array.from({ length: width }, (_, index) => {
    const value = String(rawHeaders[index] ?? "").trim();
    return value || `Columna ${columnLetter(index + 1)}`;
  });
  source.headerRowIndex = headerRowIndex;
  source.headers = headers;
  source.detectedBlocks = detectColumnBlocks(source);
  if (source.detectedBlocks.length >= 2 && source.columnBlock === "all") source.columnBlock = "choose";
  if (!["all", "choose"].includes(source.columnBlock) && !source.detectedBlocks.some(block => block.value === source.columnBlock)) source.columnBlock = "all";
  source.rows = source.matrix.slice(dataStartIndex, endRowIndex).map((values, index) => ({
    excelRow: dataStartIndex + index + 1,
    values: Array.from({ length: width }, (_, column) => values[column] ?? "")
  })).filter(row => row.values.some(value => String(value).trim() !== ""));
}

function autoMapColumns(source) {
  const normalized = source.headers.map(header => normalizeHeader(header));
  const selectedBlock = source.detectedBlocks.find(block => block.value === source.columnBlock);
  const allowed = normalized.map((_, index) => index).filter(index => !selectedBlock || (index >= selectedBlock.start && index <= selectedBlock.end));
  const find = aliases => {
    let index = allowed.find(column => aliases.includes(normalized[column]));
    if (index === undefined) index = allowed.find(column => aliases.some(alias => normalized[column].includes(alias)));
    return index === undefined ? "" : String(index);
  };
  const detected = {
    date: find(HEADER_ALIASES.date),
    description: find(HEADER_ALIASES.description),
    amount: find(HEADER_ALIASES.amount),
    debit: find(HEADER_ALIASES.debit),
    credit: find(HEADER_ALIASES.credit),
    status: find(HEADER_ALIASES.status)
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
  maybeAutoDetectSplitConvention(source);
}

function maybeAutoDetectSplitConvention(source) {
  if (source.splitConventionLocked || getEffectiveFormat(source) !== "split" || source.mapping.debit === "" || source.mapping.credit === "") return;
  const amounts = { debit: [], credit: [] };
  source.rows.slice(0, 1000).forEach(row => {
    ["debit", "credit"].forEach(field => {
      const raw = valueAt(row, source.mapping[field]);
      if (isBlank(raw)) return;
      const parsed = parseAmount(raw);
      if (parsed !== null && Math.abs(parsed) > .005) amounts[field].push(parsed);
    });
  });
  const values = [...amounts.debit, ...amounts.credit];
  if (!values.length || !values.every(value => value >= 0)) {
    source.splitConvention = "preserve";
    return;
  }
  const debitHeader = normalizeHeader(source.headers[Number(source.mapping.debit)] || "");
  const creditHeader = normalizeHeader(source.headers[Number(source.mapping.credit)] || "");
  const ledgerHeaders = /(^| )debe($| )/.test(debitHeader) || /(^| )haber($| )/.test(creditHeader);
  // En un extracto de banco o una caja expresada como Ingresos/Pagos, el
  // Crédito/Ingreso aumenta la caja y el Débito/Pago la disminuye. En un mayor
  // con Debe/Haber se conserva la convención contable tradicional.
  source.splitConvention = source.key === "bank" && !ledgerHeaders ? "credit-positive" : "debit-positive";
}

function detectHeaderRow(matrix) {
  let bestIndex = -1;
  let bestScore = -1;
  const limit = Math.min(matrix.length, 80);
  for (let index = 0; index < limit; index++) {
    const row = Array.isArray(matrix[index]) ? matrix[index] : [];
    if (!row.some(value => String(value ?? "").trim() !== "")) continue;
    const score = headerRowScore(row);
    if (score > bestScore) { bestScore = score; bestIndex = index; }
  }
  return bestIndex;
}

function headerRowScore(row) {
  const normalized = (row || []).map(normalizeHeader);
  const has = field => normalized.some(value => HEADER_ALIASES[field].some(alias => value === alias || value.includes(alias)));
  return (has("date") ? 6 : 0) + (has("description") ? 6 : 0) + (has("amount") ? 5 : 0) + (has("debit") ? 4 : 0) + (has("credit") ? 4 : 0) + Math.min(3, normalized.filter(Boolean).length / 3);
}

function selectInitialSheet(workbook) {
  for (const name of workbook.SheetNames) {
    const sheet = workbook.Sheets[name];
    if (!sheet?.["!ref"]) continue;
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const lastRow = Math.min(range.e.r, range.s.r + 79);
    const lastColumn = Math.min(range.e.c, range.s.c + 49);
    for (let row = range.s.r; row <= lastRow; row++) {
      const values = [];
      for (let column = range.s.c; column <= lastColumn; column++) values.push(sheet[XLSX.utils.encode_cell({ r: row, c: column })]?.v ?? "");
      if (headerRowScore(values) >= 15) return name;
    }
  }
  return workbook.SheetNames[0];
}

function detectReportPeriod(matrix, headerRowIndex) {
  const datePattern = /\b\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}\b/g;
  for (const row of matrix.slice(0, Math.max(0, headerRowIndex))) {
    for (const value of row || []) {
      const matches = String(value ?? "").match(datePattern) || [];
      if (matches.length < 2) continue;
      const from = parseDateValue(matches[0]);
      const to = parseDateValue(matches[1]);
      if (from && to && from <= to) return { from: toDateKey(from), to: toDateKey(to) };
    }
  }
  return null;
}

function detectColumnBlocks(source) {
  const normalized = source.headers.map(normalizeHeader);
  const dateColumns = normalized.map((value, index) => HEADER_ALIASES.date.some(alias => value === alias || value.includes(alias)) ? index : -1).filter(index => index >= 0);
  const blocks = [];
  dateColumns.forEach((start, position) => {
    const nextDate = dateColumns[position + 1] ?? normalized.length;
    const inside = (aliases, from = start + 1) => {
      for (let index = from; index < nextDate; index++) if (aliases.some(alias => normalized[index] === alias || normalized[index].includes(alias))) return index;
      return -1;
    };
    const description = inside(HEADER_ALIASES.description);
    const amount = inside(HEADER_ALIASES.amount);
    const debit = inside(HEADER_ALIASES.debit);
    const credit = inside(HEADER_ALIASES.credit);
    if (description < 0 || (amount < 0 && (debit < 0 || credit < 0))) return;
    let end = Math.max(description, amount, debit, credit);
    const rawHeader = source.matrix[source.headerRowIndex] || [];
    while (end + 1 < nextDate && !isBlank(rawHeader[end + 1])) end++;
    const title = findBlockTitle(source.matrix, source.headerRowIndex, start, end);
    const range = `${columnLetter(start + 1)}–${columnLetter(end + 1)}`;
    blocks.push({ start, end, value: `${start}:${end}`, label: `${range}${title ? ` · ${title}` : ""}` });
  });
  return blocks;
}

function findBlockTitle(matrix, headerRowIndex, start, end) {
  for (let row = headerRowIndex - 1; row >= 0; row--) {
    for (let column = start; column <= end; column++) {
      const value = String(matrix[row]?.[column] ?? "").trim();
      if (value) return value.slice(0, 50);
    }
  }
  return "";
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
  const detectedRows = source.restoredRowCount || source.rows.length;
  const restoredLabel = source.restoredState ? " · restaurado desde una conciliación guardada" : "";
  refs.fileSummary.innerHTML = `<div><i data-lucide="file-spreadsheet"></i><span><strong>${escapeHtml(source.file.name)}</strong><small>${formatFileSize(source.file.size)} · ${detectedRows.toLocaleString("es-UY")} filas detectadas${restoredLabel}</small></span></div><button class="button button-ghost button-small" type="button" data-replace-file><i data-lucide="replace"></i> Reemplazar</button>`;
  refs.fileSummary.querySelector("[data-replace-file]").addEventListener("click", () => refs.fileInput.click());
  refs.sheetSelect.innerHTML = source.sheetNames.map(name => `<option value="${escapeAttribute(name)}"${name === source.selectedSheet ? " selected" : ""}>${escapeHtml(name)}</option>`).join("");
  refs.headerRow.value = source.headerRowNumber;
  refs.dataStartRow.value = source.dataStartRow;
  refs.dataEndRow.value = source.dataEndRow;
  refs.dateFrom.value = formatDateInput(source.dateFrom);
  refs.dateTo.value = formatDateInput(source.dateTo);
  refs.excludedDescriptions.value = source.excludedDescriptions;
  refs.columnBlockField.classList.toggle("hidden", source.detectedBlocks.length < 2);
  refs.columnBlock.innerHTML = source.detectedBlocks.length >= 2
    ? [`<option value="choose"${source.columnBlock === "choose" ? " selected" : ""}>Seleccionar tabla</option>`, ...source.detectedBlocks.map(block => `<option value="${block.value}"${source.columnBlock === block.value ? " selected" : ""}>${escapeHtml(block.label)}</option>`), `<option value="all"${source.columnBlock === "all" ? " selected" : ""}>Todas las columnas (avanzado)</option>`].join("")
    : `<option value="all">Todas las columnas</option>`;
  renderImportDetectionNote(source, refs.detectionNote);
  renderMappingGrid(source, refs.mappingGrid);
  renderPreview(source, refs.previewTable);
  validateSource(sourceKey);
  refreshIcons(editor);
}

function renderImportDetectionNote(source, container) {
  const effective = getEffectiveFormat(source);
  const format = effective === "signed" ? "Monto con signo" : "Débito y Crédito";
  const splitConvention = effective === "split" ? ` · Convención${source.splitConventionLocked ? "" : " detectada"} <strong>${escapeHtml(splitConventionLabel(source.splitConvention))}</strong>` : "";
  const period = source.dateFrom && source.dateTo ? ` · Período <strong>${formatDateInput(source.dateFrom)}–${formatDateInput(source.dateTo)}</strong>${source.periodSource === "sistema" ? " sugerido desde el sistema" : source.periodSource === "archivo" ? " detectado en el archivo" : ""}` : "";
  const ignoredRows = source.importRangeInfo?.ignoredTrailingRows > 0
    ? ` · Se ignoraron <strong>${source.importRangeInfo.ignoredTrailingRows.toLocaleString("es-UY")}</strong> filas finales con ceros o formato residual`
    : "";
  container.innerHTML = `${source.formatMode === "auto" ? "Formato detectado" : "Formato seleccionado"}: <strong>${format}</strong>${splitConvention} · Encabezados en fila <strong>${source.headerRowNumber}</strong>${period}${ignoredRows}`;
}

function splitConventionLabel(value) {
  return {
    preserve: "Conservar signo escrito",
    invert: "Invertir signo escrito",
    "debit-positive": "Débito + / Crédito −",
    "credit-positive": "Débito − / Crédito +"
  }[value] || String(value || "");
}

function renderMappingGrid(source, container) {
  const effective = getEffectiveFormat(source);
  const fields = effective === "signed"
    ? ["date", "description", "amount", "status"]
    : ["date", "description", "debit", "credit", "status"];
  container.innerHTML = fields.map(field => {
    const required = field !== "status";
    const options = [`<option value="">${required ? "Seleccionar columna" : "No utilizar"}</option>`]
      .concat(source.headers.map((header, index) => `<option value="${index}"${source.mapping[field] === String(index) ? " selected" : ""}>${columnLetter(index + 1)} · ${escapeHtml(header)}</option>`));
    return `<label class="field ${required ? "required" : ""}"><span>${FIELD_LABELS[field]}</span><select data-map-field="${field}">${options.join("")}</select></label>`;
  }).join("");
  container.querySelectorAll("[data-map-field]").forEach(select => {
    select.addEventListener("change", event => {
      source.mapping[event.target.dataset.mapField] = event.target.value;
      maybeAutoDetectSplitConvention(source);
      const editor = document.querySelector(`[data-source-editor="${source.key}"]`);
      editor._refs.splitConvention.value = source.splitConvention;
      renderImportDetectionNote(source, editor._refs.detectionNote);
      validateSource(source.key);
    });
  });
}

function renderPreview(source, table) {
  const rows = source.rows.slice(0, 10);
  const mappedColumns = new Set(Object.values(source.mapping).filter(value => value !== "").map(Number));
  table.innerHTML = `<thead><tr><th class="row-number">Fila</th>${source.headers.map((header, index) => `<th class="${mappedColumns.has(index) ? "mapped-column" : ""}"><small>${columnLetter(index + 1)}</small>${escapeHtml(header)}</th>`).join("")}</tr></thead><tbody>${rows.map(row => `<tr><td class="row-number">${row.excelRow}</td>${row.values.map((value, index) => { const displayed = displayPreviewValue(source, value, index); return `<td class="${mappedColumns.has(index) ? "mapped-column" : ""}" title="${escapeAttribute(displayed)}">${escapeHtml(displayed || "—")}</td>`; }).join("")}</tr>`).join("")}</tbody>`;
}

function displayPreviewValue(source, value, columnIndex) {
  if (source.mapping.date === String(columnIndex)) {
    const date = parseDateValue(value);
    if (date) return formatDate(date);
  }
  return displayOriginalValue(value);
}

function validateSource(sourceKey) {
  const source = state.sources[sourceKey];
  if (source.restoredState && !source.workbook) {
    source.isValid = source.movements.length > 0;
    source.validationErrors = source.isValid ? [] : ["El estado guardado no contiene movimientos válidos."];
    source.validationWarnings = source.isValid ? ["Datos restaurados. Para cambiar hoja o columnas, vuelva a cargar el archivo original."] : [];
    renderValidation(source);
    dom.continueButtons[sourceKey].disabled = !source.isValid;
    scheduleStatePersistence();
    return;
  }
  const errors = [];
  const warnings = [];
  const invalidRows = [];
  const movements = [];
  let filteredRowsCount = 0;
  const effective = getEffectiveFormat(source);
  if (!source.headers.length) errors.push("No se encontró una fila de encabezados válida.");
  if (source.detectedBlocks.length >= 2 && source.columnBlock === "choose") errors.push("Se detectaron varias tablas en la hoja. Debe seleccionar cuál desea importar.");
  if (source.dateFrom && source.dateTo && source.dateFrom > source.dateTo) errors.push("La fecha inicial del filtro no puede ser posterior a la fecha final.");
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
      if (parsed.skip) {
        filteredRowsCount++;
      } else if (parsed.errors.length) {
        invalidRows.push({ source: source.label, sheet: source.selectedSheet, row: row.excelRow, errors: parsed.errors.join(" "), values: [...row.values] });
      } else {
        movements.push(parsed.movement);
      }
    }
    if (!movements.length) errors.push("No hay al menos una fila válida para procesar.");
    if (invalidRows.length) warnings.push(`Hay ${invalidRows.length} ${invalidRows.length === 1 ? "fila que no pudo interpretarse" : "filas que no pudieron interpretarse"}.`);
  }
  source.validationErrors = errors;
  source.validationWarnings = warnings;
  source.invalidRows = invalidRows;
  source.filteredRowsCount = filteredRowsCount;
  source.movements = movements;
  source.isValid = errors.length === 0 && movements.length > 0;
  renderValidation(source);
  dom.continueButtons[sourceKey].disabled = !source.isValid;
  scheduleStatePersistence();
}

function fieldsAreIncompatible(a, b) {
  if (a === b) return false;
  return true;
}

function normalizeSourceRow(source, row, effectiveFormat) {
  const errors = [];
  const rawDate = valueAt(row, source.mapping.date);
  const rawDescription = valueAt(row, source.mapping.description);
  const mappedAmounts = effectiveFormat === "signed"
    ? [valueAt(row, source.mapping.amount)]
    : [valueAt(row, source.mapping.debit), valueAt(row, source.mapping.credit)];
  const mappedAmountsAreEmpty = effectiveFormat === "split"
    ? mappedAmounts.every(value => {
      if (isBlank(value)) return true;
      const parsed = parseAmount(value);
      return parsed !== null && Math.abs(parsed) <= .005;
    })
    : mappedAmounts.every(isBlank);
  if (isBlank(rawDate) && isBlank(rawDescription) && mappedAmountsAreEmpty) return { errors: [], skip: true, skipReason: "empty-mapped-fields", movement: null };
  const excludedDescriptions = new Set(String(source.excludedDescriptions || "").split(",").map(normalizeHeader).filter(Boolean));
  if (excludedDescriptions.has(normalizeHeader(rawDescription))) return { errors: [], skip: true, skipReason: "excluded-description", movement: null };
  const date = parseDateValue(rawDate);
  const description = String(rawDescription ?? "").trim();
  if (date && !isDateWithinSourceFilter(date, source)) return { errors: [], skip: true, skipReason: "outside-date-range", movement: null };
  if (rawDate === "" || rawDate === null || rawDate === undefined) errors.push(`La fila ${row.excelRow} no contiene fecha.`);
  else if (!date) errors.push(`La fila ${row.excelRow} contiene una fecha no reconocida.`);
  if (!description) errors.push(`La fila ${row.excelRow} no contiene descripción.`);
  let signedAmount = null;
  let originalAmount = "";
  let movementType = "";
  let debitAmount = 0;
  let creditAmount = 0;
  if (effectiveFormat === "signed") {
    const rawAmount = valueAt(row, source.mapping.amount);
    const parsedAmount = parseAmount(rawAmount);
    originalAmount = rawAmount;
    if (rawAmount === "" || rawAmount === null || rawAmount === undefined) errors.push(`La fila ${row.excelRow} no contiene monto.`);
    else if (parsedAmount === null) errors.push(`La fila ${row.excelRow} contiene un monto no reconocido.`);
    else {
      signedAmount = source.positiveMeaning === "debit" ? parsedAmount : -parsedAmount;
      movementType = signedAmount >= 0 ? "debit" : "credit";
      debitAmount = movementType === "debit" ? Math.abs(signedAmount) : 0;
      creditAmount = movementType === "credit" ? Math.abs(signedAmount) : 0;
    }
  } else {
    const rawDebit = valueAt(row, source.mapping.debit);
    const rawCredit = valueAt(row, source.mapping.credit);
    const debitProvided = !isBlank(rawDebit);
    const creditProvided = !isBlank(rawCredit);
    const debit = debitProvided ? parseAmount(rawDebit) : 0;
    const credit = creditProvided ? parseAmount(rawCredit) : 0;
    const debitActive = debit !== null && Math.abs(debit) > .005;
    const creditActive = credit !== null && Math.abs(credit) > .005;
    originalAmount = `Débito: ${displayOriginalValue(rawDebit)} · Crédito: ${displayOriginalValue(rawCredit)}`;
    if (!debitProvided && !creditProvided) errors.push(`La fila ${row.excelRow} no contiene débito ni crédito.`);
    if (debitProvided && debit === null) errors.push(`La fila ${row.excelRow} contiene un débito no reconocido.`);
    if (creditProvided && credit === null) errors.push(`La fila ${row.excelRow} contiene un crédito no reconocido.`);
    if (debit !== null && credit !== null && !debitActive && !creditActive) errors.push(`La fila ${row.excelRow} no contiene un débito o crédito distinto de cero.`);
    if (debitActive && creditActive && !source.allowBoth) errors.push(`La fila ${row.excelRow} contiene débito y crédito simultáneamente.`);
    if (debit !== null && credit !== null && (debitActive || creditActive)) {
      // La dirección contable puede venir dada por el signo de la celda o por
      // la columna donde está el importe. Se conserva el valor original aparte.
      const writtenAmount = (debit || 0) + (credit || 0);
      if (source.splitConvention === "debit-positive") signedAmount = Math.abs(debit || 0) - Math.abs(credit || 0);
      else if (source.splitConvention === "credit-positive") signedAmount = -Math.abs(debit || 0) + Math.abs(credit || 0);
      else signedAmount = source.splitConvention === "invert" ? -writtenAmount : writtenAmount;
      movementType = debitActive && !creditActive ? "debit" : !debitActive && creditActive ? "credit" : "mixed";
      debitAmount = debitActive ? Math.abs(debit || 0) : 0;
      creditAmount = creditActive ? Math.abs(credit || 0) : 0;
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
      debitAmount: roundMoney(debitAmount),
      creditAmount: roundMoney(creditAmount),
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
  let body = source.filteredRowsCount
    ? `El mapeo está listo. Se omitieron ${source.filteredRowsCount.toLocaleString("es-UY")} filas por rango, período, otra tabla o descripción excluida.`
    : "El mapeo está listo para continuar.";
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

function isDateWithinSourceFilter(date, source) {
  const key = toDateKey(date);
  if (source.dateFrom && key < source.dateFrom) return false;
  if (source.dateTo && key > source.dateTo) return false;
  return true;
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
  const genericTokens = new Set(["cont", "contable", "gasto", "gastos", "factura", "credito", "debito", "pago", "pagos", "recibo", "transferencia", "movimiento", "ajuste", "caja", "banco", "cuota", "cuotas", "proveedor"]);
  const sharedBusinessToken = [...leftTokens].some(token => token.length >= 3 && /\p{L}/u.test(token) && !genericTokens.has(token) && rightTokens.has(token));
  return clamp(Math.max(baseSimilarity, fuzzyCoverage * .85, compactContainment, sharedBusinessToken ? .82 : 0), 0, 1);
}

function groupedDescriptionSimilarity(systemMovements, bankMovements) {
  if (systemMovements.length === 1 && bankMovements.length === 1) return descriptionSimilarity(systemMovements[0].description, bankMovements[0].description);
  const left = systemMovements.map(item => item.description).slice(0, 200);
  const right = bankMovements.map(item => item.description).slice(0, 200);
  const coverage = (items, others) => items.reduce((sum, description) => {
    let best = 0;
    for (const other of others) best = Math.max(best, descriptionSimilarity(description, other));
    return sum + best;
  }, 0) / Math.max(1, items.length);
  return (coverage(left, right) + coverage(right, left)) / 2;
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
  const numberFields = ["dateTolerance", "amountAbsTolerance", "amountPercentTolerance", "maxGroupSize", "maxPairComparisons", "maxCombinations", "autoThreshold", "possibleThreshold"];
  const checkboxFields = ["requireSameSign", "invertBetweenTables", "searchOneToOne", "searchOneToMany", "searchManyToOne", "allowMixedGroupSigns", "searchInternalOffsets", "considerDescription", "ignoreCase", "ignoreAccents", "ignorePunctuation", "excludeReconciled"];
  numberFields.forEach(key => { state.config[key] = Number(data.get(key)); });
  checkboxFields.forEach(key => { state.config[key] = data.has(key); });
  state.config.ignoredWords = String(data.get("ignoredWords") || "");
  if (state.config.possibleThreshold > state.config.autoThreshold) {
    state.config.possibleThreshold = state.config.autoThreshold;
    dom.configForm.elements.namedItem("possibleThreshold").value = state.config.possibleThreshold;
  }
  updateCostWarning();
  scheduleStatePersistence();
}

function updateCostWarning() {
  const warning = document.getElementById("costWarning");
  const estimate = estimateIndexedPairWorkload(state.sources.system.movements, state.sources.bank.movements, state.config);
  const complexity = Math.min(state.config.maxGroupSize, 32) * state.config.maxCombinations;
  const limitRisk = estimate.indexedPairs > state.config.maxPairComparisons;
  const largeTables = estimate.totalMovements > 30000;
  const risky = limitRisk || largeTables || state.config.maxGroupSize > 250 || state.config.maxCombinations > 500000 || complexity > 16000000;
  warning.classList.toggle("hidden", !risky);
  if (risky) {
    warning.querySelector("span").textContent = limitRisk
      ? `La estimación de ${estimate.indexedPairs.toLocaleString("es-UY")} parejas supera el límite configurado. La búsqueda se detendrá antes de bloquear el equipo.`
      : largeTables
        ? `Se procesarán ${estimate.totalMovements.toLocaleString("es-UY")} movimientos. El trabajo se ejecutará por ventanas de fecha y, cuando el navegador lo permita, en segundo plano.`
        : "Los intentos configurados se aplican a cada movimiento base y pueden producir una búsqueda más lenta. Si uno llega al tope, el motor continúa con el siguiente.";
  }
  renderWorkloadEstimate(estimate);
}

function estimateIndexedPairWorkload(systemMovements, bankMovements, config = state.config) {
  if (config.ignoreDates) {
    const bankBySign = new Map();
    bankMovements.forEach(movement => {
      const sign = config.requireSameSign ? Math.sign(comparisonAmountForConfig(movement, config)) : 0;
      bankBySign.set(sign, (bankBySign.get(sign) || 0) + 1);
    });
    let indexedPairs = 0;
    systemMovements.forEach(movement => {
      const sign = config.requireSameSign ? Math.sign(comparisonAmountForConfig(movement, config)) : 0;
      indexedPairs += bankBySign.get(sign) || 0;
    });
    return {
      totalMovements: systemMovements.length + bankMovements.length,
      naivePairs: systemMovements.length * bankMovements.length,
      indexedPairs,
      reduction: systemMovements.length && bankMovements.length ? 1 - indexedPairs / Math.max(1, systemMovements.length * bankMovements.length) : 0
    };
  }
  const tolerance = Math.max(0, Math.trunc(config.dateTolerance || 0));
  const buckets = new Map();
  bankMovements.forEach(movement => {
    const day = Math.round(movement.date.getTime() / DATE_MS);
    const sign = config.requireSameSign ? Math.sign(comparisonAmountForConfig(movement, config)) : 0;
    const key = `${day}|${sign}`;
    buckets.set(key, (buckets.get(key) || 0) + 1);
  });
  let indexedPairs = 0;
  systemMovements.forEach(movement => {
    const day = Math.round(movement.date.getTime() / DATE_MS);
    const sign = config.requireSameSign ? Math.sign(comparisonAmountForConfig(movement, config)) : 0;
    for (let offset = -tolerance; offset <= tolerance; offset++) indexedPairs += buckets.get(`${day + offset}|${sign}`) || 0;
  });
  return {
    totalMovements: systemMovements.length + bankMovements.length,
    naivePairs: systemMovements.length * bankMovements.length,
    indexedPairs,
    reduction: systemMovements.length && bankMovements.length ? 1 - indexedPairs / Math.max(1, systemMovements.length * bankMovements.length) : 0
  };
}

function comparisonAmountForConfig(movement, config) {
  return movement.source === "bank" && config.invertBetweenTables ? -movement.amount : movement.amount;
}

function renderWorkloadEstimate(estimate = estimateIndexedPairWorkload(state.sources.system.movements, state.sources.bank.movements, state.config)) {
  const container = document.getElementById("workloadEstimate");
  if (!container) return;
  const ratio = estimate.indexedPairs / Math.max(1, state.config.maxPairComparisons);
  const level = ratio > 1 || estimate.totalMovements > 100000 ? "danger" : ratio > .6 || estimate.indexedPairs > 500000 || estimate.totalMovements > 30000 ? "warning" : "safe";
  const reduction = estimate.naivePairs ? `${formatDecimal(estimate.reduction * 100, 1)}% menos` : "Sin datos";
  container.className = `workload-estimate ${level}`;
  container.innerHTML = [
    ["Movimientos", estimate.totalMovements.toLocaleString("es-UY"), `${state.sources.system.movements.length.toLocaleString("es-UY")} ↔ ${state.sources.bank.movements.length.toLocaleString("es-UY")}`],
    ["Comparación ingenua", estimate.naivePairs.toLocaleString("es-UY"), "No se ejecutará de esta forma"],
    ["Candidatos por fecha/signo", estimate.indexedPairs.toLocaleString("es-UY"), reduction],
    ["Ejecución", typeof Worker !== "undefined" ? "Web Worker" : "Modo compatible", `Tope uno a uno: ${state.config.maxPairComparisons.toLocaleString("es-UY")}`]
  ].map(([label, value, detail]) => `<div class="workload-stat"><span>${label}</span><strong>${value}</strong><small>${detail}</small></div>`).join("");
}

function renderConfigSummary() {
  const container = document.getElementById("sourceConfigSummary");
  container.innerHTML = [state.sources.system, state.sources.bank].map(source => {
    const selectedBlock = source.detectedBlocks.find(block => block.value === source.columnBlock);
    const period = source.dateFrom && source.dateTo ? ` · ${formatDateInput(source.dateFrom)}–${formatDateInput(source.dateTo)}` : "";
    return `<div class="source-pill"><span><i data-lucide="${source.key === "system" ? "database" : "landmark"}"></i></span><div><strong>${escapeHtml(source.name || source.label)}</strong><small>${source.movements.length.toLocaleString("es-UY")} movimientos · ${escapeHtml(source.selectedSheet)}${selectedBlock ? ` · ${escapeHtml(selectedBlock.label)}` : ""}${period} · ${source.invalidRows.length} errores</small></div></div>`;
  }).join("");
  updateCostWarning();
  refreshIcons(container);
}

function resetApplication() {
  try {
    const previousSnapshot = createStateSnapshot();
    if (snapshotHasProgress(previousSnapshot) && window.indexedDB) persistSnapshot(previousSnapshot, workspaceStorageKey(previousSnapshot.workspace.id)).catch(() => {});
  } catch {}
  clearPersistedState().catch(() => {});
  if (state.persistence.saveTimer) window.clearTimeout(state.persistence.saveTimer);
  state.persistence.saveTimer = null;
  state.step = 1;
  state.maxVisitedStep = 1;
  // Los controles del editor conservan referencias a estos objetos. Se limpian
  // en el lugar para que sus eventos sigan apuntando al estado vigente.
  Object.assign(state.sources.system, createEmptySource("system", "Sistema contable"));
  Object.assign(state.sources.bank, createEmptySource("bank", "Caja o banco"));
  state.config = { ...DEFAULT_CONFIG };
  state.results = createEmptyResults();
  terminateProcessingWorker();
  state.processing = { cancelled: false, running: false, worker: null, jobId: null };
  state.review = {
    tab: "confirmed", search: "", type: "all", sort: "score-desc", page: 1,
    selectedSystem: new Set(), selectedBank: new Set(), editingId: null,
    editAvailableSystem: [], editAvailableBank: [],
    editSelectedSystem: new Set(), editSelectedBank: new Set(),
    editSearchSystem: "", editSearchBank: "",
    rejectedSignatures: new Set(), rejectedProposals: [],
    periodFilter: { from: "", to: "", appliedAt: null }
  };
  state.workspace = createWorkspaceState();
  state.accountTransfer = createAccountTransferState();
  state.transferLog = [];
  state.persistence.lastSavedAt = null;
  updateLocalSaveStatus("Sin progreso guardado");
  document.querySelectorAll("[data-source-editor]").forEach(editor => {
    editor.querySelector("[data-source-name]").value = "";
    editor.querySelector("[data-file-input]").value = "";
    editor.querySelector("[data-format-mode]").value = "auto";
    editor.querySelector("[data-positive-meaning]").value = "debit";
    editor.querySelector("[data-split-convention]").value = "preserve";
    editor.querySelector("[data-allow-both]").checked = false;
    editor.querySelector("[data-header-row]").value = "";
    editor.querySelector("[data-data-start-row]").value = "";
    editor.querySelector("[data-data-end-row]").value = "";
    editor.querySelector("[data-date-from]").value = "";
    editor.querySelector("[data-date-to]").value = "";
    editor.querySelector("[data-excluded-descriptions]").value = "saldo inicial, totales, total";
    editor.querySelector("[data-file-summary]").classList.add("hidden");
    editor.querySelector("[data-import-workspace]").classList.add("hidden");
    toggleFormatOptions(editor, state.sources[editor.dataset.sourceEditor]);
  });
  dom.continueButtons.system.disabled = true;
  dom.continueButtons.bank.disabled = true;
  document.getElementById("reviewSearch").value = "";
  document.getElementById("exportFileName").value = "";
  document.getElementById("typeFilter").value = "all";
  document.getElementById("sortResults").value = "score-desc";
  syncReviewControls();
  populateConfigForm();
  goToStep(1);
  showToast("Nueva conciliación", "La cuenta anterior quedó guardada localmente y la aplicación está lista para comenzar otra.", "success");
}

function confirmNewReconciliation() {
  const hasData = state.sources.system.file || state.sources.bank.file || state.results.processingAt;
  if (!hasData || window.confirm("La cuenta actual quedará guardada localmente y se abrirá una conciliación nueva. ¿Desea continuar?")) resetApplication();
}

function clearReconciliationResultsForSourceChange() {
  terminateProcessingWorker();
  state.processing = { cancelled: false, running: false, worker: null, jobId: null };
  state.config = { ...DEFAULT_CONFIG };
  state.results = createEmptyResults();
  state.review.tab = "confirmed";
  state.review.search = "";
  state.review.type = "all";
  state.review.sort = "score-desc";
  state.review.page = 1;
  state.review.selectedSystem.clear();
  state.review.selectedBank.clear();
  state.review.editingId = null;
  state.review.editAvailableSystem = [];
  state.review.editAvailableBank = [];
  state.review.editSelectedSystem.clear();
  state.review.editSelectedBank.clear();
  state.review.editSearchSystem = "";
  state.review.editSearchBank = "";
  state.review.periodFilter = { from: "", to: "", appliedAt: null };
  clearRejectedProposals();
  const search = document.getElementById("reviewSearch");
  if (search) search.value = "";
  if (dom.configForm) populateConfigForm();
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
  const preserveResults = Boolean(state.results.processingAt || state.results.reconciliations.length);
  const systemInput = preserveResults ? getPendingMovements("system") : state.sources.system.movements;
  const bankInput = preserveResults ? getPendingMovements("bank") : state.sources.bank.movements;
  if (preserveResults && systemInput.length + bankInput.length < 2) {
    showToast("No hay pendientes suficientes", "Las conciliaciones existentes se conservaron sin cambios.", "success");
    renderReview();
    goToStep(6);
    return;
  }
  const estimate = estimateIndexedPairWorkload(systemInput, bankInput, state.config);
  state.processing = { cancelled: false, running: true, worker: null, jobId: `job-${Date.now()}-${Math.random().toString(16).slice(2)}`, preserveResults, mode: preserveResults ? "pending" : "full" };
  if (!preserveResults) state.results = createEmptyResults();
  setProcessingProgress(2, "Preparando índices de búsqueda…", `${estimate.indexedPairs.toLocaleString("es-UY")} parejas candidatas estimadas`);
  goToStep(5);
  try {
    await yieldToMain();
    if (state.processing.cancelled) return cancelProcessing();
    const engineResult = await runReconciliationEngine({
      system: systemInput,
      bank: bankInput,
      config: { ...state.config },
      estimatedPairs: estimate.indexedPairs,
      excludedSignatures: getRejectedSignatureList()
    }, { start: 2, end: 78, prefix: "Primera pasada" });
    if (state.processing.cancelled || engineResult.cancelled) return cancelProcessing();
    const primaryResult = hydrateEngineResult(engineResult);
    if (preserveResults) {
      primaryResult.reconciliations.forEach(item => {
        item.reprocessedPending = true;
        item.criterion = `Reprocesamiento de pendientes · ${item.criterion}`;
        item.reasons.push("Las conciliaciones anteriores se conservaron y sólo se analizaron los movimientos pendientes");
        addReconciliation(item);
      });
      mergeEngineMetrics(engineResult);
      state.results.engineMode = `${state.results.engineMode || "indexado"} + reprocesamiento de pendientes`;
    } else {
      state.results = primaryResult;
    }
    const automaticPass = await runAutomaticPendingPass();
    if (state.processing.cancelled || automaticPass.cancelled) return cancelProcessing();
    state.results.processingAt = new Date();
    state.processing.running = false;
    terminateProcessingWorker();
    setProcessingProgress(100, "Conciliación completada", `${preserveResults ? primaryResult.reconciliations.length + automaticPass.found : state.results.reconciliations.length} propuestas nuevas · ${automaticPass.found} encontradas en la pasada flexible automática`);
    await delay(280);
    state.review.tab = "confirmed";
    state.review.page = 1;
    state.review.selectedSystem.clear();
    state.review.selectedBank.clear();
    renderReview();
    goToStep(6);
    if (state.results.pairLimitReached) showToast("Búsqueda uno a uno acotada", "El cupo de candidatos se repartió entre todos los movimientos para no omitir los últimos. Puede ampliarlo y reprocesar los pendientes.", "error", 8500);
    if (state.results.combinationLimitReached) showToast("Búsqueda de sumas acotada", `${state.results.limitedGroupAnchors.toLocaleString("es-UY")} movimiento(s) alcanzaron el tope individual; el motor continuó con todos los siguientes.`, "error", 8500);
  } catch (error) {
    console.error(error);
    state.processing.running = false;
    terminateProcessingWorker();
    showToast("No se pudo completar la conciliación", error.message || "Ocurrió un error durante el procesamiento.", "error", 9000);
    goToStep(preserveResults ? 6 : 4);
  }
}

async function runAutomaticPendingPass() {
  const systemPending = getPendingMovements("system");
  const bankPending = getPendingMovements("bank");
  if (systemPending.length + bankPending.length < 2) return { found: 0, cancelled: false };
  const relaxedConfig = {
    ...state.config,
    dateTolerance: Math.max(state.config.dateTolerance, 5),
    amountAbsTolerance: Math.max(state.config.amountAbsTolerance, 1),
    amountPercentTolerance: Math.max(state.config.amountPercentTolerance, .1),
    possibleThreshold: Math.min(state.config.possibleThreshold, 50),
    autoThreshold: 101,
    minimumDescriptionSimilarity: .3,
    relaxedDescriptionPriority: true,
    forcePossible: true
  };
  const estimate = estimateIndexedPairWorkload(systemPending, bankPending, relaxedConfig);
  const engineResult = await runReconciliationEngine({
    system: systemPending,
    bank: bankPending,
    config: relaxedConfig,
    estimatedPairs: estimate.indexedPairs,
    excludedSignatures: getRejectedSignatureList()
  }, { start: 79, end: 98, prefix: "Segunda pasada automática" });
  if (engineResult.cancelled) return { found: 0, cancelled: true };
  const hydrated = hydrateEngineResult(engineResult).reconciliations;
  hydrated.forEach(item => {
    item.relaxedPass = true;
    item.automaticRelaxedPass = true;
    item.status = "possible";
    item.criterion = `Pasada flexible automática · ${item.criterion}`;
    item.reasons.push("Segunda pasada automática limitada a movimientos que seguían pendientes");
    item.reasons.push(`Tolerancias flexibles: ${relaxedConfig.dateTolerance} día(s), ${formatMoney(relaxedConfig.amountAbsTolerance)} absolutos y ${formatDecimal(relaxedConfig.amountPercentTolerance, 2)}%`);
    addReconciliation(item);
  });
  mergeEngineMetrics(engineResult);
  state.results.engineMode = `${state.results.engineMode || "indexado"} + pasada flexible automática`;
  state.results.retryPasses.push({
    mode: "automatic",
    at: new Date(),
    dateTolerance: relaxedConfig.dateTolerance,
    amountAbsTolerance: relaxedConfig.amountAbsTolerance,
    amountPercentTolerance: relaxedConfig.amountPercentTolerance,
    minimumDescriptionSimilarity: relaxedConfig.minimumDescriptionSimilarity * 100,
    possibleThreshold: relaxedConfig.possibleThreshold,
    ignoreDates: false,
    found: hydrated.length
  });
  return { found: hydrated.length, cancelled: false };
}

function requestProcessingCancellation() {
  if (!state.processing.running || state.processing.cancelled) return;
  state.processing.cancelled = true;
  if (state.processing.worker && state.processing.jobId) state.processing.worker.postMessage({ type: "cancel", jobId: state.processing.jobId });
  setProcessingProgress(Number(dom.processingPercent.textContent.replace("%", "")) || 0, "Cancelando procesamiento…", "Deteniendo el motor de conciliación");
}

function cancelProcessing() {
  const preserveResults = Boolean(state.processing.preserveResults);
  terminateProcessingWorker();
  state.processing.running = false;
  if (preserveResults) {
    state.review.tab = "pending";
    renderReview();
    showToast("Reanálisis cancelado", "Las conciliaciones anteriores se conservaron sin cambios.", "error");
    goToStep(6);
  } else {
    state.results = createEmptyResults();
    showToast("Procesamiento cancelado", "No se conservaron resultados parciales.", "error");
    goToStep(4);
  }
}

function openPeriodTrimDialog() {
  if (state.processing.running) return;
  const period = state.review.periodFilter || { from: "", to: "" };
  dom.periodTrimForm.elements.namedItem("from").value = formatDateInput(period.from);
  dom.periodTrimForm.elements.namedItem("to").value = formatDateInput(period.to);
  document.getElementById("restorePeriodBtn").disabled = !period.from && !period.to;
  updatePeriodTrimPreview();
  dom.periodTrimDialog.showModal();
  refreshIcons(dom.periodTrimDialog);
}

function readPeriodTrimForm() {
  const rawFrom = String(dom.periodTrimForm.elements.namedItem("from").value || "").trim();
  const rawTo = String(dom.periodTrimForm.elements.namedItem("to").value || "").trim();
  const from = parseDisplayDateInput(rawFrom);
  const to = parseDisplayDateInput(rawTo);
  return {
    from: from || "",
    to: to || "",
    invalid: Boolean((rawFrom && from === null) || (rawTo && to === null))
  };
}

function periodTrimCounts(period) {
  const system = getUnreservedMovements("system");
  const bank = getUnreservedMovements("bank");
  const excludedSystem = system.filter(item => movementIsOutsidePeriod(item, period)).length;
  const excludedBank = bank.filter(item => movementIsOutsidePeriod(item, period)).length;
  return {
    system: system.length,
    bank: bank.length,
    excludedSystem,
    excludedBank,
    excluded: excludedSystem + excludedBank,
    active: system.length + bank.length - excludedSystem - excludedBank
  };
}

function updatePeriodTrimPreview() {
  const period = readPeriodTrimForm();
  if (period.invalid) {
    document.getElementById("periodTrimPreview").innerHTML = `<strong>Fecha incompleta o no válida</strong><br><small>Utilice siempre el formato dd/mm/aaaa.</small>`;
    return;
  }
  const counts = periodTrimCounts(period);
  const range = period.from || period.to
    ? `${period.from ? `desde ${formatDateInput(period.from)}` : "sin fecha inicial"} · ${period.to ? `hasta ${formatDateInput(period.to)}` : "sin fecha final"}`
    : "todavía no se indicó un período";
  document.getElementById("periodTrimPreview").innerHTML = `<strong>${counts.active.toLocaleString("es-UY")}</strong> movimientos sin conciliar quedarían activos · <strong>${counts.excluded.toLocaleString("es-UY")}</strong> excluidos (${counts.excludedSystem.toLocaleString("es-UY")} del sistema y ${counts.excludedBank.toLocaleString("es-UY")} de caja/banco)<br><small>${escapeHtml(range)}</small>`;
}

function applyPeriodTrim(event) {
  event.preventDefault();
  const period = readPeriodTrimForm();
  if (period.invalid) {
    showToast("Fecha no válida", "Use el formato dd/mm/aaaa o dd/mm/aa e indique fechas existentes.", "error");
    return;
  }
  if (!period.from && !period.to) {
    showToast("Indique un período", "Complete al menos la fecha inicial o la fecha final.", "error");
    return;
  }
  if (period.from && period.to && period.from > period.to) {
    showToast("Período no válido", "La fecha inicial no puede ser posterior a la fecha final.", "error");
    return;
  }
  const counts = periodTrimCounts(period);
  state.review.periodFilter = { from: period.from, to: period.to, appliedAt: new Date() };
  state.review.selectedSystem.clear();
  state.review.selectedBank.clear();
  state.review.page = 1;
  dom.periodTrimDialog.close();
  renderReview();
  showToast("Período aplicado", `${counts.excluded.toLocaleString("es-UY")} movimientos quedaron excluidos sin ser eliminados.`, "success", 7000);
}

function restorePeriodExclusions() {
  const excluded = getExcludedMovements("system").length + getExcludedMovements("bank").length;
  state.review.periodFilter = { from: "", to: "", appliedAt: null };
  state.review.page = 1;
  dom.periodTrimDialog.close();
  renderReview();
  showToast("Período restaurado", `${excluded.toLocaleString("es-UY")} movimientos volvieron a Pendientes.`, "success");
}

function getCurrentWorkspaceName() {
  const exportName = document.getElementById("exportFileName")?.value.trim().replace(/\.xlsx$/i, "");
  const name = exportName || state.workspace?.name || state.sources.bank.name || state.sources.system.name || "Conciliación sin nombre";
  if (state.workspace) state.workspace.name = name;
  return name;
}

function cloneSnapshot(snapshot) {
  return JSON.parse(JSON.stringify(snapshot));
}

function normalizeTransferSnapshot(snapshot, fallbackName = "Conciliación") {
  if (!snapshot || Number(snapshot.schemaVersion) !== STATE_SCHEMA_VERSION || !snapshotHasProgress(snapshot)) throw new Error("La conciliación no contiene un estado compatible.");
  const normalized = cloneSnapshot(snapshot);
  normalized.workspace ||= {};
  normalized.workspace.id ||= `workspace-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  normalized.workspace.name ||= String(normalized.exportFileName || fallbackName).replace(/\.xlsx$/i, "");
  normalized.sources ||= {};
  for (const sourceKey of ["system", "bank"]) {
    normalized.sources[sourceKey] ||= {};
    normalized.sources[sourceKey].movements ||= [];
  }
  normalized.results ||= createEmptyResults();
  normalized.results.reconciliations ||= [];
  normalized.transferLog ||= [];
  normalized.review ||= {};
  normalized.review.periodFilter ||= { from: "", to: "", appliedAt: null };
  return normalized;
}

function snapshotWorkspaceName(snapshot) {
  return String(snapshot?.workspace?.name || snapshot?.exportFileName || snapshot?.sources?.bank?.name || snapshot?.sources?.system?.name || "Conciliación").replace(/\.xlsx$/i, "");
}

function snapshotReservedIds(snapshot) {
  const reserved = { system: new Set(), bank: new Set() };
  (snapshot?.results?.reconciliations || []).filter(item => item.status === "confirmed" || item.status === "possible").forEach(item => {
    (item.systemIds || []).forEach(id => reserved.system.add(id));
    (item.bankIds || []).forEach(id => reserved.bank.add(id));
  });
  return reserved;
}

function snapshotMovementOutsidePeriod(movement, snapshot) {
  const period = snapshot?.review?.periodFilter || {};
  const dateKey = movement.dateKey || toDateKey(reviveStoredDate(movement.date));
  return Boolean((period.from && dateKey < period.from) || (period.to && dateKey > period.to));
}

function snapshotPendingMovements(snapshot, sourceKey) {
  const reserved = snapshotReservedIds(snapshot)[sourceKey];
  return (snapshot?.sources?.[sourceKey]?.movements || []).filter(item => !reserved.has(item.id) && !snapshotMovementOutsidePeriod(item, snapshot));
}

function findCrossAccountMatch(movement, otherSnapshot) {
  if (!otherSnapshot) return null;
  const movementDate = movement.dateKey || toDateKey(reviveStoredDate(movement.date));
  const amount = Math.abs(Number(movement.amount) || 0);
  let best = null;
  for (const sourceKey of ["system", "bank"]) {
    for (const candidate of otherSnapshot.sources?.[sourceKey]?.movements || []) {
      const candidateDate = candidate.dateKey || toDateKey(reviveStoredDate(candidate.date));
      if (movementDate !== candidateDate || Math.abs(amount - Math.abs(Number(candidate.amount) || 0)) > .01) continue;
      const similarity = descriptionSimilarity(movement.description, candidate.description);
      if (similarity < .45) continue;
      if (!best || similarity > best.similarity) best = { candidate, sourceKey, similarity };
    }
  }
  return best;
}

async function openAccountTransferDialog() {
  state.accountTransfer = createAccountTransferState();
  state.accountTransfer.currentSnapshot = normalizeTransferSnapshot(createStateSnapshot(), getCurrentWorkspaceName());
  await populateSavedAccountSelect();
  renderAccountTransferDialog();
  dom.accountTransferDialog.showModal();
  refreshIcons(dom.accountTransferDialog);
}

async function populateSavedAccountSelect() {
  if (!dom.savedAccountSelect) return;
  dom.savedAccountSelect.innerHTML = `<option value="">Seleccione una cuenta local</option>`;
  if (!window.indexedDB) return;
  try {
    const entries = await readSavedWorkspaceSnapshots();
    entries
      .filter(entry => entry.snapshot?.workspace?.id !== state.workspace?.id)
      .sort((left, right) => snapshotWorkspaceName(left.snapshot).localeCompare(snapshotWorkspaceName(right.snapshot), "es"))
      .forEach(entry => {
        const option = document.createElement("option");
        option.value = String(entry.key);
        option.textContent = snapshotWorkspaceName(entry.snapshot);
        dom.savedAccountSelect.append(option);
      });
  } catch (error) {
    console.error(error);
  }
}

async function loadSelectedSavedAccount() {
  const key = dom.savedAccountSelect.value;
  if (!key) {
    showToast("Seleccione una cuenta", "Elija una conciliación guardada antes de abrirla.", "error");
    return;
  }
  try {
    const snapshot = await readPersistedSnapshot(key);
    state.accountTransfer.destinationSnapshot = normalizeTransferSnapshot(snapshot, "Cuenta destino");
    state.accountTransfer.destinationFileName = snapshotWorkspaceName(snapshot);
    state.accountTransfer.selectedDestination.clear();
    renderAccountTransferDialog();
  } catch (error) {
    console.error(error);
    showToast("No se pudo abrir la cuenta", error.message || "La conciliación guardada no está disponible.", "error");
  }
}

async function loadDestinationAccountFile(file) {
  if (!/\.xlsx$/i.test(file.name || "")) {
    showToast("Archivo no admitido", "Seleccione un XLSX exportado previamente por ConciliApp.", "error");
    return;
  }
  if (typeof XLSX === "undefined") {
    showToast("Lector de Excel no disponible", "Vuelva a abrir la aplicación con conexión a Internet.", "error");
    return;
  }
  try {
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false, dense: false });
    const snapshot = normalizeTransferSnapshot(extractStateSnapshotFromWorkbook(workbook), file.name.replace(/\.xlsx$/i, ""));
    if (snapshot.workspace.id === state.accountTransfer.currentSnapshot?.workspace?.id) snapshot.workspace.id = `workspace-${Date.now()}-${Math.random().toString(16).slice(2)}`;
    snapshot.workspace.name ||= file.name.replace(/\.xlsx$/i, "");
    state.accountTransfer.destinationSnapshot = snapshot;
    state.accountTransfer.destinationFileName = file.name.replace(/\.xlsx$/i, "");
    state.accountTransfer.selectedDestination.clear();
    renderAccountTransferDialog();
    showToast("Cuenta destino cargada", `${snapshotWorkspaceName(snapshot)} está disponible para mover pendientes.`, "success");
  } catch (error) {
    console.error(error);
    showToast("No se pudo cargar la cuenta", error.message || "El archivo no contiene una conciliación compatible.", "error", 8000);
  }
}

function renderAccountTransferDialog() {
  const transfer = state.accountTransfer;
  const current = transfer.currentSnapshot;
  const destination = transfer.destinationSnapshot;
  document.getElementById("moveToDestinationBtn").disabled = !destination || !transfer.selectedCurrent.size;
  document.getElementById("moveToCurrentBtn").disabled = !destination || !transfer.selectedDestination.size;
  document.getElementById("saveAccountTransfersBtn").disabled = !destination;
  document.getElementById("openDestinationAccountBtn").disabled = !destination;
  if (!current) {
    dom.accountTransferContent.innerHTML = `<div class="account-transfer-placeholder">No hay una conciliación actual para administrar.</div>`;
    return;
  }
  if (!destination) {
    dom.accountTransferContent.innerHTML = `${renderAccountPanel("current", current, transfer.currentSearch)}<div class="account-transfer-placeholder"><div><strong>Cargue la conciliación de la otra cuenta</strong><p>Puede elegir una cuenta guardada localmente o un XLSX exportado. Después podrá arrastrar los movimientos pendientes.</p></div></div>`;
  } else {
    dom.accountTransferContent.innerHTML = `${renderAccountPanel("current", current, transfer.currentSearch)}${renderAccountPanel("destination", destination, transfer.destinationSearch)}<div class="transfer-history-note">Al mover una fila se conserva su fecha, importe, descripción y número de fila original. La aplicación le asigna un ID nuevo en la cuenta destino y registra de qué cuenta provino.</div>`;
  }
  bindAccountTransferContentEvents();
  refreshIcons(dom.accountTransferDialog);
}

function renderAccountPanel(origin, snapshot, search) {
  const selected = origin === "current" ? state.accountTransfer.selectedCurrent : state.accountTransfer.selectedDestination;
  const otherSnapshot = origin === "current" ? state.accountTransfer.destinationSnapshot : state.accountTransfer.currentSnapshot;
  const query = String(search || "").trim().toLowerCase();
  const groups = ["system", "bank"].map(sourceKey => {
    const all = snapshotPendingMovements(snapshot, sourceKey);
    const visible = all.filter(item => !query || movementSearchText({ ...item, date: reviveStoredDate(item.date) }).includes(query));
    const rows = visible.length ? visible.map(item => {
      const key = `${sourceKey}:${item.id}`;
      const date = reviveStoredDate(item.date);
      const crossMatch = findCrossAccountMatch(item, otherSnapshot);
      const matchBadge = crossMatch ? `<em class="cross-account-match" title="Misma fecha e importe en ${escapeAttribute(snapshotWorkspaceName(otherSnapshot))}: ${escapeAttribute(crossMatch.candidate.description)}"><i data-lucide="badge-alert"></i> Posible otra cuenta</em>` : "";
      return `<div class="account-movement-row ${selected.has(key) ? "selected" : ""} ${crossMatch ? "cross-account-suggested" : ""}" draggable="true" data-account-movement data-origin="${origin}" data-source="${sourceKey}" data-id="${escapeAttribute(item.id)}"><input type="checkbox" data-account-select value="${escapeAttribute(key)}" ${selected.has(key) ? "checked" : ""} aria-label="Seleccionar movimiento"><small>${date ? formatDate(date) : escapeHtml(item.dateKey || "")}</small><span class="account-movement-description" title="${escapeAttribute(item.description)}"><span>${escapeHtml(item.description)}</span>${matchBadge}</span><strong class="account-movement-amount ${Number(item.amount) < 0 ? "negative" : ""}">${formatMoney(Number(item.amount) || 0)}</strong></div>`;
    }).join("") : `<div class="account-empty-group">No hay pendientes para este filtro.</div>`;
    return `<section class="account-movement-group"><div class="account-movement-group-heading"><strong>${sourceKey === "system" ? "Sistema contable" : "Caja o banco"}</strong><span>${visible.length.toLocaleString("es-UY")} de ${all.length.toLocaleString("es-UY")}</span></div><div class="account-movement-list">${rows}</div></section>`;
  }).join("");
  const totalPending = snapshotPendingMovements(snapshot, "system").length + snapshotPendingMovements(snapshot, "bank").length;
  return `<section class="account-panel" data-account-drop="${origin}"><div class="account-panel-heading"><div><h3>${escapeHtml(snapshotWorkspaceName(snapshot))}</h3><span>${origin === "current" ? "Cuenta actual" : "Cuenta destino"}</span></div><span class="account-panel-count">${totalPending.toLocaleString("es-UY")} pendientes</span></div><label class="search-field"><i data-lucide="search"></i><input data-account-search="${origin}" type="search" value="${escapeAttribute(search || "")}" placeholder="Filtrar movimientos"></label>${groups}</section>`;
}

function bindAccountTransferContentEvents() {
  dom.accountTransferContent.querySelectorAll("[data-account-select]").forEach(checkbox => {
    checkbox.addEventListener("change", event => {
      const row = event.target.closest("[data-account-movement]");
      const selected = row.dataset.origin === "current" ? state.accountTransfer.selectedCurrent : state.accountTransfer.selectedDestination;
      const key = `${row.dataset.source}:${row.dataset.id}`;
      if (event.target.checked) selected.add(key);
      else selected.delete(key);
      renderAccountTransferDialog();
    });
  });
  dom.accountTransferContent.querySelectorAll("[data-account-movement]").forEach(row => {
    row.addEventListener("click", event => {
      if (event.target.matches("input")) return;
      const checkbox = row.querySelector("[data-account-select]");
      checkbox.checked = !checkbox.checked;
      checkbox.dispatchEvent(new Event("change", { bubbles: true }));
    });
    row.addEventListener("dragstart", event => {
      const payload = { origin: row.dataset.origin, sourceKey: row.dataset.source, id: row.dataset.id };
      state.accountTransfer.dragPayload = payload;
      row.classList.add("dragging");
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", JSON.stringify(payload));
    });
    row.addEventListener("dragend", () => {
      row.classList.remove("dragging");
      state.accountTransfer.dragPayload = null;
      dom.accountTransferContent.querySelectorAll(".is-drop-target").forEach(panel => panel.classList.remove("is-drop-target"));
    });
  });
  dom.accountTransferContent.querySelectorAll("[data-account-search]").forEach(input => {
    input.addEventListener("input", event => {
      if (event.target.dataset.accountSearch === "current") state.accountTransfer.currentSearch = event.target.value;
      else state.accountTransfer.destinationSearch = event.target.value;
      renderAccountTransferDialog();
    });
  });
  dom.accountTransferContent.querySelectorAll("[data-account-drop]").forEach(panel => {
    panel.addEventListener("dragover", event => {
      const payload = state.accountTransfer.dragPayload;
      if (!payload || payload.origin === panel.dataset.accountDrop || !state.accountTransfer.destinationSnapshot) return;
      event.preventDefault();
      event.dataTransfer.dropEffect = "move";
      panel.classList.add("is-drop-target");
    });
    panel.addEventListener("dragleave", event => {
      if (!panel.contains(event.relatedTarget)) panel.classList.remove("is-drop-target");
    });
    panel.addEventListener("drop", event => {
      event.preventDefault();
      panel.classList.remove("is-drop-target");
      let payload = state.accountTransfer.dragPayload;
      if (!payload) {
        try { payload = JSON.parse(event.dataTransfer.getData("text/plain")); } catch { payload = null; }
      }
      if (!payload || payload.origin === panel.dataset.accountDrop) return;
      moveSingleAccountMovement(payload.origin, panel.dataset.accountDrop, payload.sourceKey, payload.id);
    });
  });
}

function transferredMovementId(destinationSnapshot, sourceKey) {
  const existing = new Set((destinationSnapshot.sources?.[sourceKey]?.movements || []).map(item => item.id));
  let candidate;
  do candidate = `${sourceKey}-transfer-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
  while (existing.has(candidate));
  return candidate;
}

function moveSnapshotMovement(originSnapshot, destinationSnapshot, sourceKey, movementId) {
  if (!originSnapshot || !destinationSnapshot || !["system", "bank"].includes(sourceKey)) return false;
  const pendingIds = new Set(snapshotPendingMovements(originSnapshot, sourceKey).map(item => item.id));
  if (!pendingIds.has(movementId)) return false;
  const movements = originSnapshot.sources[sourceKey].movements;
  const index = movements.findIndex(item => item.id === movementId);
  if (index < 0) return false;
  const [movement] = movements.splice(index, 1);
  const fromName = snapshotWorkspaceName(originSnapshot);
  const toName = snapshotWorkspaceName(destinationSnapshot);
  const movedAt = new Date().toISOString();
  const moved = {
    ...movement,
    id: transferredMovementId(destinationSnapshot, sourceKey),
    source: sourceKey,
    transferOriginId: movement.transferOriginId || movement.id,
    transferOriginAccount: movement.transferOriginAccount || fromName,
    transferHistory: [
      ...(Array.isArray(movement.transferHistory) ? movement.transferHistory : []),
      { fromWorkspaceId: originSnapshot.workspace.id, fromAccount: fromName, toWorkspaceId: destinationSnapshot.workspace.id, toAccount: toName, sourceKey, at: movedAt }
    ]
  };
  destinationSnapshot.sources[sourceKey].movements.push(moved);
  const transferEntry = {
    id: `transfer-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
    at: movedAt,
    sourceKey,
    movementOriginId: movement.transferOriginId || movement.id,
    movementDestinationId: moved.id,
    row: movement.row,
    date: movement.date,
    dateKey: movement.dateKey,
    description: movement.description,
    amount: movement.amount,
    fromWorkspaceId: originSnapshot.workspace.id,
    fromAccount: fromName,
    toWorkspaceId: destinationSnapshot.workspace.id,
    toAccount: toName
  };
  originSnapshot.transferLog ||= [];
  destinationSnapshot.transferLog ||= [];
  originSnapshot.transferLog.push({ ...transferEntry, direction: "out" });
  destinationSnapshot.transferLog.push({ ...transferEntry, direction: "in" });
  destinationSnapshot.sources[sourceKey].movements.sort((left, right) => String(left.dateKey || left.date).localeCompare(String(right.dateKey || right.date)) || Number(left.row || 0) - Number(right.row || 0));
  originSnapshot.sources[sourceKey].restoredRowCount = originSnapshot.sources[sourceKey].movements.length;
  destinationSnapshot.sources[sourceKey].restoredRowCount = destinationSnapshot.sources[sourceKey].movements.length;
  originSnapshot.savedAt = movedAt;
  destinationSnapshot.savedAt = movedAt;
  return true;
}

function moveSingleAccountMovement(origin, target, sourceKey, movementId) {
  const originSnapshot = origin === "current" ? state.accountTransfer.currentSnapshot : state.accountTransfer.destinationSnapshot;
  const destinationSnapshot = target === "current" ? state.accountTransfer.currentSnapshot : state.accountTransfer.destinationSnapshot;
  if (!moveSnapshotMovement(originSnapshot, destinationSnapshot, sourceKey, movementId)) {
    showToast("No se pudo mover", "El movimiento ya no está pendiente o fue modificado.", "error");
    return;
  }
  state.accountTransfer.selectedCurrent.clear();
  state.accountTransfer.selectedDestination.clear();
  renderAccountTransferDialog();
}

function moveSelectedAccountMovements(origin, target) {
  if (!state.accountTransfer.destinationSnapshot) return;
  const selected = origin === "current" ? state.accountTransfer.selectedCurrent : state.accountTransfer.selectedDestination;
  const originSnapshot = origin === "current" ? state.accountTransfer.currentSnapshot : state.accountTransfer.destinationSnapshot;
  const destinationSnapshot = target === "current" ? state.accountTransfer.currentSnapshot : state.accountTransfer.destinationSnapshot;
  let moved = 0;
  for (const key of [...selected]) {
    const separator = key.indexOf(":");
    const sourceKey = key.slice(0, separator);
    const movementId = key.slice(separator + 1);
    if (moveSnapshotMovement(originSnapshot, destinationSnapshot, sourceKey, movementId)) moved++;
  }
  selected.clear();
  state.accountTransfer.selectedCurrent.clear();
  state.accountTransfer.selectedDestination.clear();
  renderAccountTransferDialog();
  if (moved) showToast("Movimientos preparados", `${moved} movimiento(s) se moverán al guardar ambas cuentas.`, "success");
}

async function commitAccountTransfer(openDestination) {
  const currentSnapshot = state.accountTransfer.currentSnapshot;
  const destinationSnapshot = state.accountTransfer.destinationSnapshot;
  if (!currentSnapshot || !destinationSnapshot) return;
  currentSnapshot.workspace.name = snapshotWorkspaceName(currentSnapshot);
  destinationSnapshot.workspace.name = snapshotWorkspaceName(destinationSnapshot);
  currentSnapshot.exportFileName ||= currentSnapshot.workspace.name;
  destinationSnapshot.exportFileName ||= destinationSnapshot.workspace.name;
  const targetSnapshot = openDestination ? destinationSnapshot : currentSnapshot;
  const otherSnapshot = openDestination ? currentSnapshot : destinationSnapshot;
  try {
    await persistSnapshot(otherSnapshot, workspaceStorageKey(otherSnapshot.workspace.id));
    await persistSnapshot(targetSnapshot, workspaceStorageKey(targetSnapshot.workspace.id));
    await persistSnapshot(targetSnapshot);
    dom.accountTransferDialog.close();
    restoreStateSnapshot(targetSnapshot);
    showToast(openDestination ? "Cuenta destino abierta" : "Cuentas actualizadas", openDestination ? `${snapshotWorkspaceName(targetSnapshot)} quedó como conciliación activa.` : "Los movimientos fueron guardados en ambas conciliaciones.", "success", 7000);
  } catch (error) {
    console.error(error);
    showToast("No se pudieron guardar las cuentas", error.message || "El navegador no permitió completar el guardado.", "error", 8000);
  }
}

function saveAccountTransfers() {
  commitAccountTransfer(false);
}

function openDestinationAccount() {
  commitAccountTransfer(true);
}

function openRetryDialog() {
  if (state.processing.running) return;
  const systemPending = getPendingMovements("system");
  const bankPending = getPendingMovements("bank");
  if (systemPending.length + bankPending.length < 2) {
    showToast("No hay suficientes pendientes", "No quedan movimientos suficientes para realizar otra búsqueda.", "error");
    return;
  }
  const form = dom.retryForm;
  form.elements.namedItem("dateTolerance").value = Math.max(state.config.dateTolerance, 30);
  form.elements.namedItem("amountAbsTolerance").value = Math.max(state.config.amountAbsTolerance, 1);
  form.elements.namedItem("amountPercentTolerance").value = Math.max(state.config.amountPercentTolerance, .1);
  form.elements.namedItem("minimumDescriptionSimilarity").value = 55;
  form.elements.namedItem("possibleThreshold").value = Math.min(state.config.possibleThreshold, 45);
  form.elements.namedItem("searchGroups").checked = state.config.searchOneToMany || state.config.searchManyToOne;
  form.elements.namedItem("ignoreDates").checked = false;
  document.getElementById("retryPendingSummary").innerHTML = `<strong>${systemPending.length.toLocaleString("es-UY")}</strong> pendientes del sistema · <strong>${bankPending.length.toLocaleString("es-UY")}</strong> de caja o banco · <strong>${state.review.rejectedSignatures.size.toLocaleString("es-UY")}</strong> agrupaciones rechazadas que no se repetirán`;
  dom.retryDialog.showModal();
  refreshIcons(dom.retryDialog);
}

function readRetryOptions() {
  const data = new FormData(dom.retryForm);
  return {
    dateTolerance: clamp(Math.trunc(Number(data.get("dateTolerance")) || 0), 0, 3660),
    amountAbsTolerance: Math.max(0, Number(data.get("amountAbsTolerance")) || 0),
    amountPercentTolerance: clamp(Number(data.get("amountPercentTolerance")) || 0, 0, 100),
    minimumDescriptionSimilarity: clamp(Number(data.get("minimumDescriptionSimilarity")) || 0, 0, 100),
    possibleThreshold: clamp(Number(data.get("possibleThreshold")) || 0, 0, 100),
    searchGroups: data.get("searchGroups") === "on",
    ignoreDates: data.get("ignoreDates") === "on"
  };
}

async function retryPendingReconciliation(options = {}) {
  if (state.processing.running) return;
  let systemPending = getPendingMovements("system");
  let bankPending = getPendingMovements("bank");
  if (systemPending.length + bankPending.length < 2) {
    showToast("No hay suficientes pendientes", "No quedan movimientos suficientes para realizar otra búsqueda.", "error");
    return;
  }
  const retryOptions = {
    dateTolerance: Math.max(state.config.dateTolerance, 30),
    amountAbsTolerance: Math.max(state.config.amountAbsTolerance, 1),
    amountPercentTolerance: Math.max(state.config.amountPercentTolerance, .1),
    minimumDescriptionSimilarity: 55,
    possibleThreshold: Math.min(state.config.possibleThreshold, 45),
    searchGroups: true,
    ignoreDates: false,
    ...options
  };
  const relaxedConfig = {
    ...state.config,
    dateTolerance: retryOptions.dateTolerance,
    amountAbsTolerance: retryOptions.amountAbsTolerance,
    amountPercentTolerance: retryOptions.amountPercentTolerance,
    possibleThreshold: retryOptions.possibleThreshold,
    autoThreshold: 101,
    minimumDescriptionSimilarity: retryOptions.minimumDescriptionSimilarity / 100,
    relaxedDescriptionPriority: true,
    allowLowDescriptionDatedOneToOne: true,
    forcePossible: true,
    searchOneToMany: state.config.searchOneToMany && retryOptions.searchGroups,
    searchManyToOne: state.config.searchManyToOne && retryOptions.searchGroups
  };
  const estimate = estimateIndexedPairWorkload(systemPending, bankPending, relaxedConfig);
  state.processing = {
    cancelled: false,
    running: true,
    worker: null,
    jobId: `retry-${Date.now()}-${Math.random().toString(16).slice(2)}`,
    preserveResults: true,
    mode: "advanced-retry"
  };
  setProcessingProgress(2, "Buscando nuevamente en los pendientes…", `${systemPending.length + bankPending.length} movimientos · parámetros personalizados`);
  goToStep(5);
  try {
    await yieldToMain();
    if (state.processing.cancelled) return cancelProcessing();
    const engineResult = await runReconciliationEngine({
      system: systemPending,
      bank: bankPending,
      config: relaxedConfig,
      estimatedPairs: estimate.indexedPairs,
      excludedSignatures: getRejectedSignatureList()
    }, { start: 2, end: retryOptions.ignoreDates ? 68 : 98, prefix: "Búsqueda avanzada" });
    if (state.processing.cancelled || engineResult.cancelled) return cancelProcessing();
    const hydrated = hydrateEngineResult(engineResult).reconciliations;
    hydrated.forEach(item => {
      item.relaxedPass = true;
      item.advancedRetry = true;
      item.status = "possible";
      item.criterion = `Búsqueda avanzada de pendientes · ${item.criterion}`;
      item.reasons.push("Búsqueda configurada manualmente y limitada a movimientos pendientes");
      item.reasons.push(`Parámetros: ${relaxedConfig.dateTolerance} día(s), tolerancia ${formatMoney(relaxedConfig.amountAbsTolerance)} y ${formatDecimal(relaxedConfig.amountPercentTolerance, 2)}%`);
      addReconciliation(item);
    });
    mergeEngineMetrics(engineResult);

    const dateAgnostic = [];
    if (retryOptions.ignoreDates) {
      systemPending = getPendingMovements("system");
      bankPending = getPendingMovements("bank");
      if (systemPending.length && bankPending.length) {
        const fullDateSpan = movementDateSpanDays([...systemPending, ...bankPending]);
        const noDateConfig = {
          ...relaxedConfig,
          dateTolerance: 0,
          ignoreDates: true,
          searchOneToOne: true,
          searchOneToMany: false,
          searchManyToOne: false,
          searchInternalOffsets: false,
          forcePossible: true
        };
        const noDateEstimate = estimateIndexedPairWorkload(systemPending, bankPending, noDateConfig);
        const noDateResult = await runReconciliationEngine({
          system: systemPending,
          bank: bankPending,
          config: noDateConfig,
          estimatedPairs: noDateEstimate.indexedPairs,
          excludedSignatures: getRejectedSignatureList()
        }, { start: 70, end: 98, prefix: "Búsqueda sin límite de fecha" });
        if (state.processing.cancelled || noDateResult.cancelled) return cancelProcessing();
        for (const item of hydrateEngineResult(noDateResult).reconciliations) {
          item.relaxedPass = true;
          item.advancedRetry = true;
          item.dateAgnosticPass = true;
          item.status = "possible";
          item.criterion = `Coincidencia sin límite de fecha · ${item.criterion}`;
          item.reasons.push(`La diferencia de fecha se ignoró deliberadamente; el rango inspeccionado fue de ${fullDateSpan} día(s)`);
          item.reasons.push(`Se exigió al menos ${retryOptions.minimumDescriptionSimilarity}% de similitud o una referencia numérica compartida`);
          addReconciliation(item);
          dateAgnostic.push(item);
        }
        mergeEngineMetrics(noDateResult);
      }
    }
    state.results.processingAt = new Date();
    state.results.engineMode = `${state.results.engineMode || "indexado"} + búsqueda avanzada`;
    state.results.retryPasses.push({
      mode: "advanced",
      at: new Date(),
      dateTolerance: relaxedConfig.dateTolerance,
      amountAbsTolerance: relaxedConfig.amountAbsTolerance,
      amountPercentTolerance: relaxedConfig.amountPercentTolerance,
      minimumDescriptionSimilarity: retryOptions.minimumDescriptionSimilarity,
      possibleThreshold: relaxedConfig.possibleThreshold,
      autoThreshold: relaxedConfig.autoThreshold,
      ignoreDates: retryOptions.ignoreDates,
      found: hydrated.length + dateAgnostic.length
    });
    state.processing.running = false;
    terminateProcessingWorker();
    const possibleCount = hydrated.length + dateAgnostic.length;
    setProcessingProgress(100, "Búsqueda avanzada completada", `${possibleCount} posibles conciliaciones nuevas`);
    await delay(280);
    state.review.tab = possibleCount ? "possible" : "pending";
    state.review.page = 1;
    state.review.selectedSystem.clear();
    state.review.selectedBank.clear();
    renderReview();
    goToStep(6);
    if (possibleCount) showToast("Búsqueda completada", `Se agregaron ${possibleCount} propuestas a Posibles.`, "success", 6500);
    else showToast("Sin nuevas coincidencias", "La búsqueda no encontró relaciones que cumplan los parámetros elegidos.", "error", 6500);
  } catch (error) {
    console.error(error);
    state.processing.running = false;
    terminateProcessingWorker();
    state.review.tab = "pending";
    renderReview();
    goToStep(6);
    showToast("No se pudo reanalizar", error.message || "Las conciliaciones anteriores se conservaron.", "error", 9000);
  }
}

function mergeEngineMetrics(engineResult) {
  state.results.candidatePairs += engineResult.candidatePairs || 0;
  state.results.evaluatedPairs += engineResult.evaluatedPairs || 0;
  state.results.evaluatedCombinations += engineResult.evaluatedCombinations || 0;
  state.results.limitedGroupAnchors += engineResult.limitedGroupAnchors || 0;
  state.results.pairLimitReached ||= Boolean(engineResult.pairLimitReached);
  state.results.combinationLimitReached ||= Boolean(engineResult.combinationLimitReached);
}

function movementDateSpanDays(movements) {
  let minimum = Infinity;
  let maximum = -Infinity;
  for (const movement of movements) {
    const timestamp = movement.date.getTime();
    minimum = Math.min(minimum, timestamp);
    maximum = Math.max(maximum, timestamp);
  }
  return Number.isFinite(minimum) ? Math.max(0, Math.ceil((maximum - minimum) / DATE_MS)) : 0;
}

function reconciliationSignatureFromIds(systemIds = [], bankIds = []) {
  return `${[...systemIds].sort().join("|")}::${[...bankIds].sort().join("|")}`;
}

function reconciliationSignature(item) {
  return reconciliationSignatureFromIds(item.systemIds, item.bankIds);
}

function getRejectedSignatureList() {
  return [...state.review.rejectedSignatures];
}

function rememberRejectedProposal(item, reason) {
  const signature = reconciliationSignature(item);
  state.review.rejectedSignatures.add(signature);
  if (!state.review.rejectedProposals.some(entry => entry.signature === signature)) {
    state.review.rejectedProposals.push({
      signature,
      systemIds: [...item.systemIds],
      bankIds: [...item.bankIds],
      reason,
      at: new Date()
    });
  }
}

function clearRejectedProposals() {
  state.review.rejectedSignatures.clear();
  state.review.rejectedProposals.length = 0;
}

function terminateProcessingWorker() {
  if (state.processing?.worker) {
    if (state.processing.worker._objectUrl) URL.revokeObjectURL(state.processing.worker._objectUrl);
    state.processing.worker.terminate();
    state.processing.worker = null;
  }
}

function hydrateEngineResult(engineResult) {
  const systemById = new Map(state.sources.system.movements.map(movement => [movement.id, movement]));
  const bankById = new Map(state.sources.bank.movements.map(movement => [movement.id, movement]));
  const reconciliations = engineResult.reconciliations.map(item => ({
    ...item,
    systemMovements: item.systemIds.map(id => systemById.get(id)).filter(Boolean),
    bankMovements: item.bankIds.map(id => bankById.get(id)).filter(Boolean),
    createdAt: new Date(item.createdAt)
  }));
  return {
    ...createEmptyResults(),
    ...engineResult,
    reconciliations,
    nextId: reconciliations.length + 1
  };
}

function setProcessingProgress(percent, message, detail) {
  const value = clamp(Math.round(percent), 0, 100);
  dom.processingBar.style.width = `${value}%`;
  dom.processingPercent.textContent = `${value}%`;
  dom.processingMessage.textContent = message;
  dom.processingDetail.textContent = detail;
  dom.processingBar.parentElement.setAttribute("aria-valuenow", String(value));
}

async function runReconciliationEngine(payload, progressRange = {}) {
  const cleanPayload = {
    ...payload,
    system: payload.system.map(toEngineMovement),
    bank: payload.bank.map(toEngineMovement)
  };
  const progressStart = Number.isFinite(progressRange.start) ? progressRange.start : 0;
  const progressEnd = Number.isFinite(progressRange.end) ? progressRange.end : 100;
  const progressPrefix = progressRange.prefix ? `${progressRange.prefix} · ` : "";
  const progress = update => setProcessingProgress(
    progressStart + (progressEnd - progressStart) * clamp(Number(update.percent) || 0, 0, 100) / 100,
    `${progressPrefix}${update.message}`,
    update.detail
  );
  if (typeof Worker !== "undefined" && typeof Blob !== "undefined" && typeof URL !== "undefined" && typeof URL.createObjectURL === "function") {
    try {
      const workerSource = `${reconciliationEngine.toString()}\n(${reconciliationWorkerRuntime.toString()})();`;
      const objectUrl = URL.createObjectURL(new Blob([workerSource], { type: "text/javascript" }));
      const worker = new Worker(objectUrl);
      worker._objectUrl = objectUrl;
      state.processing.worker = worker;
      return await new Promise((resolve, reject) => {
        worker.onmessage = event => {
          const message = event.data || {};
          if (message.jobId !== state.processing.jobId) return;
          if (message.type === "progress") progress(message.update);
          if (message.type === "done") {
            terminateProcessingWorker();
            resolve({ ...message.result, engineMode: "worker" });
          }
          if (message.type === "error") {
            terminateProcessingWorker();
            reject(new Error(message.error || "El Web Worker no pudo completar la conciliación."));
          }
        };
        worker.onerror = event => {
          terminateProcessingWorker();
          reject(new Error(event.message || "Error en el motor de conciliación."));
        };
        worker.postMessage({ type: "start", jobId: state.processing.jobId, payload: cleanPayload });
      });
    } catch (error) {
      console.warn("No se pudo iniciar el Web Worker; se utilizará el modo compatible.", error);
      terminateProcessingWorker();
    }
  }
  const result = await reconciliationEngine(cleanPayload, progress, () => state.processing.cancelled);
  return { ...result, engineMode: "main-thread-indexed" };
}

function toEngineMovement(movement) {
  return {
    id: movement.id,
    source: movement.source,
    row: movement.row,
    dateTime: movement.date.getTime(),
    description: movement.description,
    amount: movement.amount,
    type: movement.type
  };
}

function reconciliationWorkerRuntime() {
  const cancelled = new Set();
  self.onmessage = async event => {
    const message = event.data || {};
    if (message.type === "cancel") {
      cancelled.add(message.jobId);
      return;
    }
    if (message.type !== "start") return;
    try {
      const result = await reconciliationEngine(
        message.payload,
        update => self.postMessage({ type: "progress", jobId: message.jobId, update }),
        () => cancelled.has(message.jobId)
      );
      self.postMessage({ type: "done", jobId: message.jobId, result });
    } catch (error) {
      self.postMessage({ type: "error", jobId: message.jobId, error: error?.message || String(error) });
    } finally {
      cancelled.delete(message.jobId);
    }
  };
}

async function reconciliationEngine(payload, emitProgress = () => {}, shouldCancel = () => false) {
  const DAY_MS = 86400000;
  const config = payload.config;
  const excludedSignatures = new Set(payload.excludedSignatures || []);
  const metrics = {
    candidatePairs: 0,
    evaluatedPairs: 0,
    evaluatedCombinations: 0,
    pairLimitReached: false,
    combinationLimitReached: false,
    limitedGroupAnchors: 0
  };
  let nextId = 1;
  const reconciliations = [];
  const pause = () => new Promise(resolve => setTimeout(resolve, 0));
  const clampValue = (value, minimum, maximum) => Math.min(maximum, Math.max(minimum, value));
  const round = value => Math.round((Number(value) + Number.EPSILON) * 100) / 100;
  const comparisonAmount = movement => movement.source === "bank" && config.invertBetweenTables ? -movement.amount : movement.amount;
  const amountTolerance = reference => Math.max(config.amountAbsTolerance, Math.abs(reference) * config.amountPercentTolerance / 100);
  const signKey = movement => config.requireSameSign ? Math.sign(movement.comparisonAmount) : 0;
  const dayKey = movement => String(movement.day);
  const dateKey = movement => `${movement.day}|${signKey(movement)}`;
  const amountKey = movement => `${movement.day}|${signKey(movement)}|${Math.round(movement.comparisonAmount * 100)}`;
  const formatAmount = value => new Intl.NumberFormat("es-UY", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(Number(value) || 0);

  function normalizeText(value) {
    let text = String(value ?? "").trim();
    if (config.ignoreCase) text = text.toLowerCase();
    if (config.ignoreAccents) text = text.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    if (config.ignorePunctuation) text = text.replace(/[^\p{L}\p{N}\s]/gu, " ");
    const ignored = String(config.ignoredWords || "").split(",").map(word => word.trim()).filter(Boolean);
    if (ignored.length) {
      const ignoredSet = new Set(ignored.map(word => {
        let normalized = config.ignoreAccents ? word.normalize("NFD").replace(/[\u0300-\u036f]/g, "") : word;
        return config.ignoreCase ? normalized.toLowerCase() : normalized;
      }));
      text = text.split(/\s+/).filter(token => !ignoredSet.has(config.ignoreCase ? token.toLowerCase() : token)).join(" ");
    }
    return text.replace(/\s+/g, " ").trim();
  }

  function levenshtein(a, b) {
    if (a.length < b.length) [a, b] = [b, a];
    let previous = Array.from({ length: b.length + 1 }, (_, index) => index);
    for (let i = 1; i <= a.length; i++) {
      const current = [i];
      for (let j = 1; j <= b.length; j++) current[j] = Math.min(current[j - 1] + 1, previous[j] + 1, previous[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      previous = current;
    }
    return previous[b.length];
  }

  function similarity(a, b) {
    const left = normalizeText(a);
    const right = normalizeText(b);
    if (!left && !right) return 1;
    if (!left || !right) return 0;
    if (left === right) return 1;
    const leftTokens = new Set(left.split(" ").filter(Boolean));
    const rightTokens = new Set(right.split(" ").filter(Boolean));
    const intersection = [...leftTokens].filter(token => rightTokens.has(token)).length;
    const union = new Set([...leftTokens, ...rightTokens]).size;
    const jaccard = union ? intersection / union : 0;
    const maximumLength = Math.max(left.length, right.length);
    const editSimilarity = maximumLength ? 1 - levenshtein(left, right) / maximumLength : 1;
    const containment = left.includes(right) || right.includes(left) ? Math.min(left.length, right.length) / maximumLength : 0;
    const leftList = left.split(" ").filter(token => token.length > 1);
    const rightList = right.split(" ").filter(token => token.length > 1);
    const shorter = leftList.length <= rightList.length ? leftList : rightList;
    const longer = leftList.length <= rightList.length ? rightList : leftList;
    const phrases = new Set(longer);
    for (let size = 2; size <= 3; size++) for (let index = 0; index <= longer.length - size; index++) phrases.add(longer.slice(index, index + size).join(""));
    const fuzzyCoverage = shorter.length ? shorter.reduce((sum, token) => {
      let best = 0;
      for (const phrase of phrases) {
        const length = Math.max(token.length, phrase.length);
        best = Math.max(best, length ? 1 - levenshtein(token, phrase) / length : 1);
      }
      return sum + best;
    }, 0) / shorter.length : 0;
    const compactLeft = left.replace(/\s+/g, "");
    const compactRight = right.replace(/\s+/g, "");
    const shorterCompact = compactLeft.length <= compactRight.length ? compactLeft : compactRight;
    const longerCompact = compactLeft.length <= compactRight.length ? compactRight : compactLeft;
    const compactContainment = shorterCompact.length >= 5 && longerCompact.includes(shorterCompact) ? .82 : 0;
    const genericTokens = new Set(["cont", "contable", "gasto", "gastos", "factura", "credito", "debito", "pago", "pagos", "recibo", "transferencia", "movimiento", "ajuste", "caja", "banco", "cuota", "cuotas", "proveedor"]);
    const sharedBusinessToken = [...leftTokens].some(token => token.length >= 3 && /\p{L}/u.test(token) && !genericTokens.has(token) && rightTokens.has(token));
    const businessTokenBoost = sharedBusinessToken ? (config.relaxedDescriptionPriority ? .96 : .82) : 0;
    return clampValue(Math.max(jaccard * .5 + editSimilarity * .35 + containment * .15, fuzzyCoverage * .85, compactContainment, businessTokenBoost), 0, 1);
  }

  function groupSimilarity(systemMovements, bankMovements) {
    if (systemMovements.length === 1 && bankMovements.length === 1) return similarity(systemMovements[0].description, bankMovements[0].description);
    const left = systemMovements.map(item => item.description).slice(0, 200);
    const right = bankMovements.map(item => item.description).slice(0, 200);
    const coverage = (items, others) => items.reduce((sum, description) => {
      let best = 0;
      for (const other of others) best = Math.max(best, similarity(description, other));
      return sum + best;
    }, 0) / Math.max(1, items.length);
    return (coverage(left, right) + coverage(right, left)) / 2;
  }

  function references(description) {
    return new Set((String(description).match(/\d{3,}/g) || []).map(value => value.replace(/^0+/, "") || "0"));
  }

  function referencePoints(descriptionsA, descriptionsB) {
    const left = new Set(descriptionsA.flatMap(value => [...references(value)]));
    const right = new Set(descriptionsB.flatMap(value => [...references(value)]));
    if (!left.size || !right.size) return 0;
    return [...left].some(value => right.has(value)) ? 10 : 0;
  }

  function genericDescription(description) {
    const normalized = normalizeText(description);
    const tokens = normalized.split(" ").filter(Boolean);
    return tokens.length <= 1 || normalized.length < 5 || ["varios", "ajuste", "movimiento", "deposito", "caja", "banco"].includes(normalized);
  }

  function prepareMovement(movement) {
    const dateTime = Number(movement.dateTime);
    const prepared = { ...movement, dateTime, day: Math.round(dateTime / DAY_MS) };
    prepared.comparisonAmount = comparisonAmount(prepared);
    return prepared;
  }

  function calculate(systemMovements, bankMovements, type) {
    const totalSystem = round(systemMovements.reduce((sum, movement) => sum + movement.comparisonAmount, 0));
    const totalBank = round(bankMovements.reduce((sum, movement) => sum + movement.comparisonAmount, 0));
    const totalBankOriginal = round(bankMovements.reduce((sum, movement) => sum + movement.amount, 0));
    const difference = round(totalSystem - totalBank);
    const tolerance = amountTolerance(Math.max(Math.abs(totalSystem), Math.abs(totalBank)));
    const amountWithinTolerance = Math.abs(difference) <= tolerance + .005;
    const amountScore = amountWithinTolerance
      ? tolerance <= .005 ? (Math.abs(difference) <= .005 ? 40 : 0) : 40 - Math.min(5, 5 * Math.abs(difference) / Math.max(tolerance, .01))
      : 0;
    let maximumDateDifference = 0;
    let totalDateDifference = 0;
    let dateComparisonCount = 0;
    for (const systemMovement of systemMovements) {
      for (const bankMovement of bankMovements) {
        const differenceInDays = Math.abs(Math.round((systemMovement.dateTime - bankMovement.dateTime) / DAY_MS));
        maximumDateDifference = Math.max(maximumDateDifference, differenceInDays);
        totalDateDifference += differenceInDays;
        dateComparisonCount++;
      }
    }
    const averageDateDifference = dateComparisonCount ? totalDateDifference / dateComparisonCount : 0;
    const ignoreDates = Boolean(config.ignoreDates);
    const dateWithinTolerance = ignoreDates || maximumDateDifference <= config.dateTolerance;
    const dateScore = ignoreDates ? 0 : dateWithinTolerance ? config.dateTolerance === 0 ? 20 : 20 * Math.max(.35, 1 - averageDateDifference / (config.dateTolerance + 1)) : 0;
    const descriptionSimilarity = config.considerDescription ? groupSimilarity(systemMovements, bankMovements) : 1;
    const refsScore = referencePoints(systemMovements.map(item => item.description), bankMovements.map(item => item.description));
    const totalMembers = systemMovements.length + bankMovements.length;
    const systemSigns = new Set(systemMovements.map(item => Math.sign(item.comparisonAmount)).filter(Boolean));
    const bankSigns = new Set(bankMovements.map(item => Math.sign(item.comparisonAmount)).filter(Boolean));
    const mixedGroupSigns = type !== "one-to-one" && (systemSigns.size > 1 || bankSigns.size > 1);
    const protectedGroupDescription = type !== "one-to-one" && config.considerDescription && descriptionSimilarity >= .82;
    const descriptionScore = 30 * (protectedGroupDescription ? 1 : descriptionSimilarity);
    let penalty = 0;
    if (type !== "one-to-one" && !protectedGroupDescription) penalty += Math.min(12, Math.log2(Math.max(1, totalMembers - 2)) * 2);
    if (!protectedGroupDescription && (systemMovements.some(item => genericDescription(item.description)) || bankMovements.some(item => genericDescription(item.description)))) penalty += 3;
    if (!ignoreDates && maximumDateDifference > 0) penalty += Math.min(5, maximumDateDifference);
    if (mixedGroupSigns) penalty += 18;
    let score = Math.round(clampValue(amountScore + dateScore + descriptionScore + refsScore - penalty, 0, 100));
    const exactAmountAndDate = Math.abs(difference) <= .005 && maximumDateDifference === 0;
    const exactAmountCompatibleDate = Math.abs(difference) <= .005 && dateWithinTolerance;
    if (exactAmountAndDate) score = Math.max(score, type === "one-to-one" ? 90 : 85);
    else if (exactAmountCompatibleDate) score = Math.max(score, type === "one-to-one" ? 85 : 80);
    const exact = Math.abs(difference) <= .005 && dateWithinTolerance && (!config.considerDescription || descriptionSimilarity >= .9 || protectedGroupDescription);
    const reasons = [
      Math.abs(difference) <= .005 ? "Los importes coinciden exactamente" : `La diferencia de importe (${formatAmount(Math.abs(difference))}) está dentro de la tolerancia`,
      ignoreDates ? `La fecha se ignoró en esta búsqueda (${maximumDateDifference} día(s) de diferencia)` : maximumDateDifference === 0 ? "Las fechas coinciden" : `Las fechas difieren hasta ${maximumDateDifference} día(s)`,
      config.considerDescription ? `Similitud de descripciones: ${Math.round(descriptionSimilarity * 100)}%` : "La comparación de descripciones está desactivada"
    ];
    if (refsScore) reasons.push("Se encontraron referencias numéricas coincidentes");
    if (type !== "one-to-one") reasons.push(`La suma de ${type === "one-to-many" ? bankMovements.length : systemMovements.length} movimientos coincide con el otro lado`);
    if (mixedGroupSigns) reasons.push("La agrupación mezcla Débitos y Créditos; requiere revisión manual aunque el neto coincida");
    if (protectedGroupDescription) reasons.push("La descripción comercial coincide; no se aplicó penalización por cantidad de movimientos");
    if (exactAmountAndDate && descriptionSimilarity < .82) reasons.push("Monto y fecha exactos tuvieron prioridad sobre la diferencia de descripción");
    return {
      id: null,
      type,
      systemIds: systemMovements.map(item => item.id),
      bankIds: bankMovements.map(item => item.id),
      totalSystem,
      totalBank,
      totalBankOriginal,
      difference,
      amountTolerance: tolerance,
      amountWithinTolerance,
      dateWithinTolerance,
      dateDifference: maximumDateDifference,
      descriptionSimilarity,
      referenceScore: refsScore,
      protectedGroupDescription,
      mixedGroupSigns,
      exactAmountAndDate,
      score,
      exact,
      ambiguous: false,
      alternativeCount: 0,
      totalMembers,
      status: "possible",
      criterion: exact ? "Monto exacto, fecha compatible y descripción coincidente" : type === "one-to-one" ? "Coincidencia probable uno a uno" : "Agrupación por suma de importes",
      reasons,
      observation: "",
      createdAt: Date.now(),
      manual: false
    };
  }

  function buildIndex(movements, keyBuilder) {
    const index = new Map();
    movements.forEach(movement => {
      const key = keyBuilder(movement);
      if (!index.has(key)) index.set(key, []);
      index.get(key).push(movement);
    });
    return index;
  }

  function lowerBoundByAmount(movements, target) {
    let low = 0;
    let high = movements.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (movements[middle].comparisonAmount < target) low = middle + 1;
      else high = middle;
    }
    return low;
  }

  function nearestAmountCursor(movements, target, margin) {
    let right = lowerBoundByAmount(movements, target);
    let left = right - 1;
    return () => {
      const leftItem = left >= 0 && (!Number.isFinite(margin) || movements[left].comparisonAmount >= target - margin) ? movements[left] : null;
      const rightItem = right < movements.length && (!Number.isFinite(margin) || movements[right].comparisonAmount <= target + margin) ? movements[right] : null;
      if (!leftItem && !rightItem) return null;
      if (!rightItem || (leftItem && Math.abs(leftItem.comparisonAmount - target) <= Math.abs(rightItem.comparisonAmount - target))) {
        left--;
        return leftItem;
      }
      right++;
      return rightItem;
    };
  }

  function reserve(reconciliation, sets) {
    reconciliation.systemIds.forEach(id => sets.system.add(id));
    reconciliation.bankIds.forEach(id => sets.bank.add(id));
  }

  function proposalSignature(reconciliation) {
    return `${[...reconciliation.systemIds].sort().join("|")}::${[...reconciliation.bankIds].sort().join("|")}`;
  }

  function candidateIsAllowed(reconciliation) {
    if (excludedSignatures.has(proposalSignature(reconciliation))) return false;
    const minimumDescriptionSimilarity = Number(config.minimumDescriptionSimilarity) || 0;
    const descriptionBelowMinimum = minimumDescriptionSimilarity > 0
      && reconciliation.descriptionSimilarity < minimumDescriptionSimilarity
      && !(reconciliation.referenceScore > 0);
    if (descriptionBelowMinimum) {
      const datedOneToOneFallback = Boolean(config.allowLowDescriptionDatedOneToOne)
        && !config.ignoreDates
        && reconciliation.type === "one-to-one"
        && reconciliation.amountWithinTolerance
        && reconciliation.dateWithinTolerance;
      if (!datedOneToOneFallback) return false;
      if (!reconciliation.lowDescriptionFallback) {
        reconciliation.lowDescriptionFallback = true;
        reconciliation.reasons.push("La descripción quedó por debajo del mínimo; se mostró porque monto y fecha alcanzaron el puntaje configurado");
      }
    }
    return true;
  }

  function addReconciliation(reconciliation) {
    if (!candidateIsAllowed(reconciliation)) return false;
    if (config.forcePossible) reconciliation.status = "possible";
    reconciliation.id = `CON-${String(nextId++).padStart(5, "0")}`;
    reconciliations.push(reconciliation);
    return true;
  }

  function applyAmbiguity(candidates, anchorSelector, penaltyMaximum, penaltyBase) {
    const grouped = new Map();
    candidates.forEach(candidate => {
      const key = anchorSelector(candidate);
      if (!grouped.has(key)) grouped.set(key, []);
      grouped.get(key).push(candidate);
    });
    candidates.forEach(candidate => {
      const alternatives = (grouped.get(anchorSelector(candidate)) || []).filter(item => item !== candidate && item.score >= candidate.score - 5);
      candidate.ambiguous = alternatives.length > 0;
      candidate.alternativeCount = alternatives.length;
      if (candidate.ambiguous) {
        candidate.score = Math.max(0, candidate.score - Math.min(penaltyMaximum, penaltyBase + alternatives.length * 2));
        candidate.reasons.push(`${alternatives.length} alternativa(s) con puntuación similar`);
      }
    });
  }

  function oneToOneQuality(candidate) {
    return candidate.score * 1000
      + Math.round((candidate.descriptionSimilarity || 0) * 20000)
      + Number(candidate.referenceScore || 0) * 250
      + (candidate.exactAmountAndDate ? 3000 : 0)
      - Number(candidate.dateDifference || 0) * 500;
  }

  function selectOneToOneGlobally(candidates) {
    const byNode = new Map();
    const addNodeEdge = (node, candidate) => {
      if (!byNode.has(node)) byNode.set(node, []);
      byNode.get(node).push(candidate);
    };
    candidates.forEach(candidate => {
      addNodeEdge(`s:${candidate.systemIds[0]}`, candidate);
      addNodeEdge(`b:${candidate.bankIds[0]}`, candidate);
    });
    const visitedNodes = new Set();
    const selected = [];
    for (const startNode of byNode.keys()) {
      if (visitedNodes.has(startNode)) continue;
      const queue = [startNode];
      const componentCandidates = new Set();
      const systemIds = new Set();
      const bankIds = new Set();
      while (queue.length) {
        const node = queue.pop();
        if (visitedNodes.has(node)) continue;
        visitedNodes.add(node);
        if (node.startsWith("s:")) systemIds.add(node.slice(2));
        else bankIds.add(node.slice(2));
        for (const candidate of byNode.get(node) || []) {
          componentCandidates.add(candidate);
          const systemNode = `s:${candidate.systemIds[0]}`;
          const bankNode = `b:${candidate.bankIds[0]}`;
          if (!visitedNodes.has(systemNode)) queue.push(systemNode);
          if (!visitedNodes.has(bankNode)) queue.push(bankNode);
        }
      }
      const edges = [...componentCandidates];
      const leftIsSystem = systemIds.size <= bankIds.size;
      const leftIds = [...(leftIsSystem ? systemIds : bankIds)];
      const rightIds = [...(leftIsSystem ? bankIds : systemIds)];
      if (leftIds.length <= 12 && rightIds.length <= 18 && edges.length <= 90) {
        const rightIndex = new Map(rightIds.map((id, index) => [id, index]));
        const edgesByLeft = new Map(leftIds.map(id => [id, []]));
        edges.forEach(candidate => {
          const leftId = leftIsSystem ? candidate.systemIds[0] : candidate.bankIds[0];
          const rightId = leftIsSystem ? candidate.bankIds[0] : candidate.systemIds[0];
          if (rightIndex.has(rightId)) edgesByLeft.get(leftId)?.push({ candidate, rightIndex: rightIndex.get(rightId) });
        });
        const memo = new Map();
        const solve = (leftIndex, usedMask) => {
          if (leftIndex >= leftIds.length) return { value: 0, chosen: [] };
          const key = `${leftIndex}|${usedMask}`;
          if (memo.has(key)) return memo.get(key);
          let best = solve(leftIndex + 1, usedMask);
          for (const edge of edgesByLeft.get(leftIds[leftIndex]) || []) {
            const bit = 1 << edge.rightIndex;
            if (usedMask & bit) continue;
            const tail = solve(leftIndex + 1, usedMask | bit);
            const value = 1000000 + oneToOneQuality(edge.candidate) + tail.value;
            if (value > best.value) best = { value, chosen: [edge.candidate, ...tail.chosen] };
          }
          memo.set(key, best);
          return best;
        };
        selected.push(...solve(0, 0).chosen);
      } else {
        const usedSystem = new Set();
        const usedBank = new Set();
        edges.sort((a, b) => oneToOneQuality(b) - oneToOneQuality(a));
        for (const candidate of edges) {
          const systemId = candidate.systemIds[0];
          const bankId = candidate.bankIds[0];
          if (usedSystem.has(systemId) || usedBank.has(bankId)) continue;
          usedSystem.add(systemId);
          usedBank.add(bankId);
          selected.push(candidate);
        }
      }
    }
    return selected;
  }

  function applySelectedOneToOneAmbiguity(selected, allCandidates) {
    const bySystem = new Map();
    const byBank = new Map();
    allCandidates.forEach(candidate => {
      const systemId = candidate.systemIds[0];
      const bankId = candidate.bankIds[0];
      if (!bySystem.has(systemId)) bySystem.set(systemId, []);
      if (!byBank.has(bankId)) byBank.set(bankId, []);
      bySystem.get(systemId).push(candidate);
      byBank.get(bankId).push(candidate);
    });
    selected.forEach(candidate => {
      const quality = oneToOneQuality(candidate);
      const alternatives = new Set([
        ...(bySystem.get(candidate.systemIds[0]) || []),
        ...(byBank.get(candidate.bankIds[0]) || [])
      ].filter(item => item !== candidate && oneToOneQuality(item) >= quality - 1800));
      candidate.ambiguous = alternatives.size > 0;
      candidate.alternativeCount = alternatives.size;
      if (candidate.ambiguous) {
        candidate.score = Math.max(0, candidate.score - Math.min(12, 4 + alternatives.size * 2));
        candidate.reasons.push(`${alternatives.size} alternativa(s) globales con calidad similar`);
      } else if (allCandidates.length > selected.length) {
        candidate.reasons.push("Se eligió dentro de la asignación global que mejor combina fecha, descripción e importe");
      }
    });
  }

  async function checkpoint(percent, message, detail, force = false) {
    if (!force && (metrics.candidatePairs + metrics.evaluatedCombinations) % 2048 !== 0) return false;
    emitProgress({ percent, message, detail });
    await pause();
    return shouldCancel();
  }

  const system = payload.system.map(prepareMovement);
  const bank = payload.bank.map(prepareMovement);
  const reserved = { system: new Set(), bank: new Set() };

  if (config.searchOneToOne) {
    emitProgress({ percent: 7, message: config.ignoreDates ? "Construyendo índice por signo e importe…" : "Construyendo índices por fecha, signo e importe…", detail: `${system.length + bank.length} movimientos` });
    const bankByExactAmount = config.ignoreDates ? new Map() : buildIndex(bank, amountKey);
    const bankBySignAndAmount = buildIndex(bank, signKey);
    bankBySignAndAmount.forEach(items => items.sort((a, b) => a.comparisonAmount - b.comparisonAmount));
    const candidates = [];
    const estimated = Math.max(1, payload.estimatedPairs || system.length);
    const candidatesPerMovement = Math.max(1, Math.ceil(Math.max(1, config.maxPairComparisons) / Math.max(1, system.length)));
    for (let leftIndex = 0; leftIndex < system.length; leftIndex++) {
      const left = system[leftIndex];
      const seenForLeft = new Set();
      let inspectedForLeft = 0;
      let stopCurrentMovement = false;
      const inspect = right => {
        if (seenForLeft.has(right.id)) return false;
        if (inspectedForLeft >= candidatesPerMovement) {
          metrics.pairLimitReached = true;
          stopCurrentMovement = true;
          return false;
        }
        seenForLeft.add(right.id);
        inspectedForLeft++;
        metrics.candidatePairs++;
        const difference = Math.abs(left.comparisonAmount - right.comparisonAmount);
        if (difference <= amountTolerance(Math.max(Math.abs(left.comparisonAmount), Math.abs(right.comparisonAmount))) + .005) {
          metrics.evaluatedPairs++;
          const calculated = calculate([left], [right], "one-to-one");
          if (calculated.amountWithinTolerance && calculated.dateWithinTolerance && calculated.score >= config.possibleThreshold && candidateIsAllowed(calculated)) candidates.push(calculated);
        }
        return true;
      };
      if (config.ignoreDates) {
        const sorted = bankBySignAndAmount.get(signKey(left)) || [];
        const percentage = Math.max(0, Number(config.amountPercentTolerance) || 0) / 100;
        const percentageMargin = percentage >= 1 ? Infinity : Math.abs(left.comparisonAmount) * percentage / Math.max(1e-9, 1 - percentage);
        const margin = Math.max(Number(config.amountAbsTolerance) || 0, percentageMargin) + .01;
        const nextClosest = nearestAmountCursor(sorted, left.comparisonAmount, margin);
        let right;
        while ((right = nextClosest()) && !stopCurrentMovement) {
          const counted = inspect(right);
          if (counted && metrics.candidatePairs % 2048 === 0) {
            const progress = 8 + 46 * Math.min(1, metrics.candidatePairs / estimated);
            if (await checkpoint(progress, "Buscando coincidencias uno a uno sin fecha…", `${metrics.candidatePairs.toLocaleString("es-UY")} candidatos por signo e importe`, true)) return { cancelled: true, reconciliations: [], ...metrics };
          }
        }
      } else {
        for (let offset = -config.dateTolerance; offset <= config.dateTolerance && !stopCurrentMovement; offset++) {
          const exactKey = `${left.day + offset}|${signKey(left)}|${Math.round(left.comparisonAmount * 100)}`;
          for (const right of bankByExactAmount.get(exactKey) || []) {
            const counted = inspect(right);
            if (counted && metrics.candidatePairs % 2048 === 0) {
              const progress = 8 + 46 * Math.min(1, metrics.candidatePairs / estimated);
              if (await checkpoint(progress, "Buscando coincidencias uno a uno…", `${metrics.candidatePairs.toLocaleString("es-UY")} candidatos por fecha/signo`, true)) return { cancelled: true, reconciliations: [], ...metrics };
            }
            if (stopCurrentMovement) break;
          }
        }
        const sorted = bankBySignAndAmount.get(signKey(left)) || [];
        const percentage = Math.max(0, Number(config.amountPercentTolerance) || 0) / 100;
        const percentageMargin = percentage >= 1 ? Infinity : Math.abs(left.comparisonAmount) * percentage / Math.max(1e-9, 1 - percentage);
        const margin = Math.max(Number(config.amountAbsTolerance) || 0, percentageMargin) + .01;
        const nextClosest = nearestAmountCursor(sorted, left.comparisonAmount, margin);
        let right;
        while ((right = nextClosest()) && !stopCurrentMovement) {
          if (Math.abs(left.day - right.day) > config.dateTolerance) continue;
          const counted = inspect(right);
          if (counted && metrics.candidatePairs % 2048 === 0) {
            const progress = 8 + 46 * Math.min(1, metrics.candidatePairs / estimated);
            if (await checkpoint(progress, "Buscando coincidencias uno a uno…", `${metrics.candidatePairs.toLocaleString("es-UY")} candidatos por fecha, signo e importe`, true)) return { cancelled: true, reconciliations: [], ...metrics };
          }
        }
      }
      if (shouldCancel()) return { cancelled: true, reconciliations: [], ...metrics };
      if (leftIndex % 100 === 0) {
        emitProgress({ percent: 8 + 46 * ((leftIndex + 1) / Math.max(1, system.length)), message: config.ignoreDates ? "Buscando coincidencias uno a uno sin fecha…" : "Buscando coincidencias uno a uno…", detail: `${leftIndex + 1} de ${system.length} movimientos indexados` });
        await pause();
      }
    }
    const globallySelected = selectOneToOneGlobally(candidates);
    applySelectedOneToOneAmbiguity(globallySelected, candidates);
    globallySelected.sort((a, b) => oneToOneQuality(b) - oneToOneQuality(a));
    for (const candidate of globallySelected) {
      if (reserved.system.has(candidate.systemIds[0]) || reserved.bank.has(candidate.bankIds[0]) || candidate.score < config.possibleThreshold) continue;
      candidate.status = candidate.score >= config.autoThreshold && !candidate.ambiguous && !candidate.mixedGroupSigns ? "confirmed" : "possible";
      if (addReconciliation(candidate)) reserve(candidate, reserved);
    }
  }

  async function enumerateGroups(anchor, compatible, type) {
    const results = [];
    const mixedSigns = Boolean(config.allowMixedGroupSigns);
    const target = mixedSigns ? anchor.comparisonAmount : Math.abs(anchor.comparisonAmount);
    const tolerance = amountTolerance(Math.abs(target));
    const maximumSelected = Math.min(Math.max(2, config.maxGroupSize), compatible.length);
    const members = compatible.map(item => ({ item, searchAmount: mixedSigns ? item.comparisonAmount : Math.abs(item.comparisonAmount) }));
    const stack = [{ start: 0, selected: [], sum: 0 }];
    let anchorAttempts = 0;
    while (stack.length && anchorAttempts < config.maxCombinations) {
      const current = stack.pop();
      if (current.selected.length >= 2) {
        anchorAttempts++;
        metrics.evaluatedCombinations++;
        if (Math.abs(target - current.sum) <= tolerance + .005) {
          const selectedMovements = current.selected.map(entry => entry.item);
          const systemMovements = type === "one-to-many" ? [anchor] : selectedMovements;
          const bankMovements = type === "one-to-many" ? selectedMovements : [anchor];
          const calculated = calculate(systemMovements, bankMovements, type);
          if (calculated.amountWithinTolerance && calculated.dateWithinTolerance && calculated.score >= config.possibleThreshold) results.push(calculated);
        }
        if (anchorAttempts % 1024 === 0) {
          if (await checkpoint(type === "one-to-many" ? 70 : 88, "Buscando agrupaciones por ventanas de fecha…", `${metrics.evaluatedCombinations.toLocaleString("es-UY")} intentos totales`, true)) return { results, cancelled: true, limitReached: false };
        }
      }
      if (current.selected.length >= maximumSelected || current.start >= members.length || (!mixedSigns && current.sum > target + tolerance + .005)) continue;
      for (let index = members.length - 1; index >= current.start; index--) {
        const next = members[index];
        const nextSum = current.sum + next.searchAmount;
        if (mixedSigns || nextSum <= target + tolerance + .005) stack.push({ start: index + 1, selected: [...current.selected, next], sum: nextSum });
      }
    }
    const limitReached = stack.length > 0 && anchorAttempts >= config.maxCombinations;
    if (limitReached) {
      metrics.combinationLimitReached = true;
      metrics.limitedGroupAnchors++;
    }
    return { results, cancelled: false, limitReached };
  }

  function findBulkGroups(anchor, compatible, type) {
    const results = [];
    const mixedSigns = Boolean(config.allowMixedGroupSigns);
    const target = mixedSigns ? anchor.comparisonAmount : Math.abs(anchor.comparisonAmount);
    const tolerance = amountTolerance(Math.abs(target));
    const maximumMembers = Math.max(2, Math.trunc(config.maxGroupSize || 2));
    const related = compatible.filter(member => similarity(anchor.description, member.description) >= .7 || referencePoints([anchor.description], [member.description]) > 0);
    const variants = [
      { label: "mismo día y descripción relacionada", members: related.filter(member => member.day === anchor.day) },
      { label: "ventana de fecha y descripción relacionada", members: related },
      { label: "todos los movimientos del mismo día", members: compatible.filter(member => member.day === anchor.day) },
      { label: "todos los movimientos de la ventana de fecha", members: compatible }
    ];
    const seen = new Set();
    for (const variant of variants) {
      const members = variant.members;
      if (members.length < 2 || members.length > maximumMembers) continue;
      const signature = members.map(item => item.id).sort().join("|");
      if (seen.has(signature)) continue;
      seen.add(signature);
      metrics.evaluatedCombinations++;
      const sum = round(members.reduce((total, member) => total + (mixedSigns ? member.comparisonAmount : Math.abs(member.comparisonAmount)), 0));
      if (Math.abs(target - sum) > tolerance + .005) continue;
      const systemMovements = type === "one-to-many" ? [anchor] : members;
      const bankMovements = type === "one-to-many" ? members : [anchor];
      const calculated = calculate(systemMovements, bankMovements, type);
      if (!calculated.amountWithinTolerance || !calculated.dateWithinTolerance || calculated.score < config.possibleThreshold) continue;
      calculated.bulk = true;
      calculated.criterion = "Agrupación masiva por suma exacta";
      calculated.reasons.push(`Se tomó el conjunto completo: ${variant.label}`);
      calculated.reasons.push("La agrupación masiva se comprobó sin enumerar todas sus subcombinaciones");
      results.push(calculated);
    }
    return results;
  }

  async function searchGroups(type, progressStart, progressEnd) {
    if (shouldCancel()) return true;
    const systemRemaining = system.filter(item => !reserved.system.has(item.id));
    const bankRemaining = bank.filter(item => !reserved.bank.has(item.id));
    const anchors = type === "one-to-many" ? systemRemaining : bankRemaining;
    const pool = type === "one-to-many" ? bankRemaining : systemRemaining;
    const mixedSigns = Boolean(config.allowMixedGroupSigns);
    const poolByDate = buildIndex(pool, mixedSigns ? dayKey : dateKey);
    const candidates = [];
    for (let anchorIndex = 0; anchorIndex < anchors.length; anchorIndex++) {
      const anchor = anchors[anchorIndex];
      const compatible = [];
      for (let offset = -config.dateTolerance; offset <= config.dateTolerance; offset++) {
        const key = mixedSigns ? String(anchor.day + offset) : `${anchor.day + offset}|${signKey(anchor)}`;
        for (const member of poolByDate.get(key) || []) {
          if (mixedSigns || Math.abs(member.comparisonAmount) <= Math.abs(anchor.comparisonAmount) + amountTolerance(Math.abs(anchor.comparisonAmount)) + .005) compatible.push(member);
        }
      }
      compatible.sort((a, b) => similarity(anchor.description, b.description) - similarity(anchor.description, a.description)
        || Math.abs(anchor.comparisonAmount - a.comparisonAmount) - Math.abs(anchor.comparisonAmount - b.comparisonAmount));
      if (compatible.length >= 2) {
        const bulkGroups = findBulkGroups(anchor, compatible, type);
        for (const group of bulkGroups) candidates.push(group);
        const hasStrongBulkGroup = bulkGroups.some(candidate => candidate.score >= config.autoThreshold);
        if (!hasStrongBulkGroup) {
          const subsetPoolSize = Math.min(compatible.length, Math.max(24, Math.min(32, config.maxGroupSize)));
          const enumerated = await enumerateGroups(anchor, compatible.slice(0, subsetPoolSize), type);
          for (const group of enumerated.results) candidates.push(group);
          if (enumerated.cancelled) return true;
        }
      }
      if (anchorIndex % 20 === 0 || anchorIndex === anchors.length - 1) {
        emitProgress({ percent: progressStart + (progressEnd - progressStart) * ((anchorIndex + 1) / Math.max(1, anchors.length)), message: "Buscando agrupaciones por ventanas de fecha…", detail: `${anchorIndex + 1} de ${anchors.length} movimientos base · ${metrics.limitedGroupAnchors} acotados` });
        await pause();
        if (shouldCancel()) return true;
      }
    }
    const eligibleCandidates = candidates.filter(candidateIsAllowed);
    applyAmbiguity(eligibleCandidates, candidate => type === "one-to-many" ? candidate.systemIds[0] : candidate.bankIds[0], 18, 8);
    eligibleCandidates.sort((a, b) => b.score - a.score || a.totalMembers - b.totalMembers || a.dateDifference - b.dateDifference);
    for (const candidate of eligibleCandidates) {
      if (candidate.systemIds.some(id => reserved.system.has(id)) || candidate.bankIds.some(id => reserved.bank.has(id)) || candidate.score < config.possibleThreshold) continue;
      candidate.status = candidate.score >= config.autoThreshold && !candidate.ambiguous && !candidate.mixedGroupSigns ? "confirmed" : "possible";
      if (addReconciliation(candidate)) reserve(candidate, reserved);
    }
    return false;
  }

  function internalDescriptionSimilarity(movements) {
    if (movements.length < 2) return 0;
    let total = 0;
    for (let index = 0; index < movements.length; index++) {
      let best = 0;
      for (let other = 0; other < movements.length; other++) {
        if (index !== other) best = Math.max(best, similarity(movements[index].description, movements[other].description));
      }
      total += best;
    }
    return total / movements.length;
  }

  function internalBusinessKey(description) {
    const genericTokens = new Set(["cont", "contable", "gasto", "gastos", "factura", "credito", "debito", "pago", "pagos", "recibo", "transferencia", "movimiento", "ajuste", "caja", "banco", "cuota", "cuotas", "proveedor", "ingreso", "egreso"]);
    const tokens = normalizeText(description).split(" ").filter(token => token.length >= 3 && /\p{L}/u.test(token) && !genericTokens.has(token));
    return tokens.slice(0, 2).join("|");
  }

  function internalOffsetEvidence(movements) {
    const reversalWords = ["devolucion", "reversa", "reversion", "anulacion", "anulado", "retorno", "estorno", "contrapartida"];
    const normalized = movements.map(item => normalizeText(item.description));
    const hasReversalWord = normalized.some(description => reversalWords.some(word => description.includes(word)));
    let sharedReference = false;
    let relatedOppositeSigns = false;
    for (let left = 0; left < movements.length; left++) {
      for (let right = left + 1; right < movements.length; right++) {
        if (Math.sign(movements[left].comparisonAmount) === Math.sign(movements[right].comparisonAmount)) continue;
        if (referencePoints([movements[left].description], [movements[right].description])) sharedReference = true;
        const leftKey = internalBusinessKey(movements[left].description);
        const rightKey = internalBusinessKey(movements[right].description);
        if ((leftKey && leftKey === rightKey) || similarity(movements[left].description, movements[right].description) >= .62) relatedOppositeSigns = true;
      }
    }
    return {
      strong: sharedReference || (hasReversalWord && relatedOppositeSigns),
      sharedReference,
      hasReversalWord,
      relatedOppositeSigns
    };
  }

  function calculateInternal(movements, sourceKey, criterion) {
    const net = round(movements.reduce((sum, movement) => sum + movement.comparisonAmount, 0));
    let reference = 0;
    let minimumDay = Infinity;
    let maximumDay = -Infinity;
    for (const movement of movements) {
      reference = Math.max(reference, Math.abs(movement.comparisonAmount));
      minimumDay = Math.min(minimumDay, movement.day);
      maximumDay = Math.max(maximumDay, movement.day);
    }
    const tolerance = amountTolerance(reference);
    const amountWithinTolerance = Math.abs(net) <= tolerance + .005;
    const dateDifference = Number.isFinite(minimumDay) ? maximumDay - minimumDay : 0;
    const dateWithinTolerance = dateDifference <= config.dateTolerance;
    const descriptionSimilarity = config.considerDescription ? internalDescriptionSimilarity(movements) : 1;
    const evidence = internalOffsetEvidence(movements);
    let refsScore = 0;
    for (let index = 0; index < movements.length && !refsScore; index++) {
      for (let other = index + 1; other < movements.length && !refsScore; other++) {
        refsScore = referencePoints([movements[index].description], [movements[other].description]);
      }
    }
    const amountScore = amountWithinTolerance ? (Math.abs(net) <= .005 ? 45 : 40) : 0;
    const dateScore = dateWithinTolerance ? (dateDifference === 0 ? 25 : 25 * Math.max(.35, 1 - dateDifference / (config.dateTolerance + 1))) : 0;
    const penalty = Math.min(12, Math.log2(Math.max(1, movements.length - 1)) * 2)
      + (movements.some(item => genericDescription(item.description)) ? 3 : 0)
      + (evidence.strong ? 0 : 15);
    let score = Math.round(clampValue(amountScore + dateScore + 20 * descriptionSimilarity + refsScore - penalty, 0, 100));
    if (!evidence.strong) score = Math.min(score, Math.max(config.possibleThreshold, config.autoThreshold - 1));
    const label = sourceKey === "system" ? "sistema contable" : "caja o banco";
    const reasons = [
      Math.abs(net) <= .005 ? "Los Débitos y Créditos seleccionados dejan un neto de cero" : `El neto (${formatAmount(net)}) está dentro de la tolerancia`,
      dateDifference === 0 ? "Los movimientos son del mismo día" : `Las fechas abarcan ${dateDifference} día(s)`,
      config.considerDescription ? `Similitud interna de descripciones: ${Math.round(descriptionSimilarity * 100)}%` : "La comparación de descripciones está desactivada",
      `La compensación ocurre dentro de ${label}; no necesita un movimiento del otro lado`
    ];
    if (refsScore) reasons.push("Se encontraron referencias numéricas coincidentes");
    if (evidence.strong) reasons.push("La reversa tiene evidencia semántica o referencias compartidas");
    else reasons.push("No hay evidencia suficiente de reversa o anulación; no se aprobará automáticamente");
    return {
      id: null,
      type: sourceKey === "system" ? "internal-system" : "internal-bank",
      systemIds: sourceKey === "system" ? movements.map(item => item.id) : [],
      bankIds: sourceKey === "bank" ? movements.map(item => item.id) : [],
      totalSystem: sourceKey === "system" ? net : 0,
      totalBank: sourceKey === "bank" ? net : 0,
      totalBankOriginal: sourceKey === "bank" ? round(movements.reduce((sum, movement) => sum + movement.amount, 0)) : 0,
      difference: net,
      amountTolerance: tolerance,
      amountWithinTolerance,
      dateWithinTolerance,
      dateDifference,
      descriptionSimilarity,
      internalAutoEligible: evidence.strong,
      score,
      exact: Math.abs(net) <= .005 && dateWithinTolerance && evidence.strong && (!config.considerDescription || descriptionSimilarity >= .7),
      ambiguous: false,
      alternativeCount: 0,
      totalMembers: movements.length,
      status: "possible",
      criterion,
      reasons,
      observation: "",
      createdAt: Date.now(),
      manual: false
    };
  }

  async function searchInternalOffsets(sourceKey, movements, reservedSet, progressStart, progressEnd) {
    const remaining = movements.filter(item => !reservedSet.has(item.id) && Math.sign(item.comparisonAmount));
    if (remaining.length < 2) return false;
    const exactIndex = buildIndex(remaining, item => `${item.day}|${Math.sign(item.comparisonAmount)}|${Math.round(Math.abs(item.comparisonAmount) * 100)}`);
    const relatedIndex = buildIndex(remaining, item => `${item.day}|${Math.sign(item.comparisonAmount)}|${Math.round(Math.abs(item.comparisonAmount) * 100)}|${internalBusinessKey(item.description)}`);
    const pairCandidates = [];
    const seenPairs = new Set();
    for (let index = 0; index < remaining.length; index++) {
      const anchor = remaining[index];
      if (anchor.comparisonAmount > 0) {
        const matches = [];
        const matchIds = new Set();
        const appendMatches = items => {
          const start = items.length ? index % items.length : 0;
          for (let step = 0; step < items.length; step++) {
            if (matches.length >= 8) break;
            const item = items[(start + step) % items.length];
            if (!matchIds.has(item.id)) { matchIds.add(item.id); matches.push(item); }
          }
        };
        const businessKey = internalBusinessKey(anchor.description);
        for (let offset = -config.dateTolerance; offset <= config.dateTolerance; offset++) {
          const key = `${anchor.day + offset}|-1|${Math.round(Math.abs(anchor.comparisonAmount) * 100)}`;
          if (businessKey) appendMatches(relatedIndex.get(`${key}|${businessKey}`) || []);
          appendMatches(exactIndex.get(key) || []);
        }
        matches.sort((a, b) => similarity(anchor.description, b.description) - similarity(anchor.description, a.description)
          || Math.abs(anchor.day - a.day) - Math.abs(anchor.day - b.day));
        // Evita materializar miles de parejas equivalentes cuando un archivo
        // contiene muchos importes repetidos en la misma fecha.
        for (const other of matches.slice(0, 4)) {
          const signature = [anchor.id, other.id].sort().join("|");
          if (anchor.id === other.id || seenPairs.has(signature)) continue;
          seenPairs.add(signature);
          const candidate = calculateInternal([anchor, other], sourceKey, "Compensación interna de Débito y Crédito");
          if (candidate.amountWithinTolerance && candidate.dateWithinTolerance && candidate.score >= config.possibleThreshold && candidateIsAllowed(candidate)) pairCandidates.push(candidate);
        }
      }
      if (index % 250 === 0) {
        emitProgress({ percent: progressStart + (progressEnd - progressStart) * .45 * ((index + 1) / remaining.length), message: "Buscando compensaciones internas…", detail: `${index + 1} de ${remaining.length} movimientos en ${sourceKey === "system" ? "sistema" : "caja/banco"}` });
        await pause();
        if (shouldCancel()) return true;
      }
    }
    const candidatesByMovement = new Map();
    pairCandidates.forEach(candidate => [...candidate.systemIds, ...candidate.bankIds].forEach(id => {
      if (!candidatesByMovement.has(id)) candidatesByMovement.set(id, []);
      candidatesByMovement.get(id).push(candidate);
    }));
    pairCandidates.forEach(candidate => {
      const alternatives = new Set();
      [...candidate.systemIds, ...candidate.bankIds].forEach(id => (candidatesByMovement.get(id) || []).forEach(other => {
        if (other !== candidate && other.score >= candidate.score - 5) alternatives.add(other);
      }));
      candidate.ambiguous = alternatives.size > 0;
      candidate.alternativeCount = alternatives.size;
      if (candidate.ambiguous) {
        candidate.score = Math.max(0, candidate.score - Math.min(15, 6 + alternatives.size * 2));
        candidate.reasons.push(`${alternatives.size} compensación(es) alternativa(s) con puntuación similar`);
      }
    });
    pairCandidates.sort((a, b) => b.score - a.score || a.dateDifference - b.dateDifference);
    for (const candidate of pairCandidates) {
      const ids = sourceKey === "system" ? candidate.systemIds : candidate.bankIds;
      if (ids.some(id => reservedSet.has(id)) || candidate.score < config.possibleThreshold) continue;
      candidate.status = candidate.score >= config.autoThreshold && !candidate.ambiguous && candidate.internalAutoEligible ? "confirmed" : "possible";
      if (addReconciliation(candidate)) ids.forEach(id => reservedSet.add(id));
    }

    const bulkRemaining = movements.filter(item => !reservedSet.has(item.id) && Math.sign(item.comparisonAmount));
    const buckets = new Map();
    bulkRemaining.forEach(item => {
      const businessKey = internalBusinessKey(item.description);
      if (!businessKey) return;
      const key = `${item.day}|${businessKey}`;
      if (!buckets.has(key)) buckets.set(key, []);
      buckets.get(key).push(item);
    });
    for (const members of buckets.values()) {
      if (members.length < 2 || members.length > config.maxGroupSize) continue;
      const signs = new Set(members.map(item => Math.sign(item.comparisonAmount)));
      if (!signs.has(1) || !signs.has(-1)) continue;
      const candidate = calculateInternal(members, sourceKey, "Compensación interna agrupada por fecha y descripción");
      metrics.evaluatedCombinations++;
      if (!candidate.amountWithinTolerance || !candidate.dateWithinTolerance || candidate.score < config.possibleThreshold) continue;
      candidate.bulk = true;
      candidate.reasons.push("Se comprobó el conjunto completo sin enumerar subcombinaciones");
      candidate.status = candidate.score >= config.autoThreshold && candidate.internalAutoEligible ? "confirmed" : "possible";
      if (addReconciliation(candidate)) members.forEach(item => reservedSet.add(item.id));
    }
    emitProgress({ percent: progressEnd, message: "Compensaciones internas comprobadas…", detail: `${sourceKey === "system" ? "Sistema" : "Caja/banco"}: ${remaining.length.toLocaleString("es-UY")} movimientos inspeccionados` });
    await pause();
    return shouldCancel();
  }

  if (config.searchOneToMany && await searchGroups("one-to-many", 56, 74)) return { cancelled: true, reconciliations: [], ...metrics };
  if (config.searchManyToOne && await searchGroups("many-to-one", 75, 92)) return { cancelled: true, reconciliations: [], ...metrics };
  if (config.searchInternalOffsets && await searchInternalOffsets("system", system, reserved.system, 92, 94)) return { cancelled: true, reconciliations: [], ...metrics };
  if (config.searchInternalOffsets && await searchInternalOffsets("bank", bank, reserved.bank, 94, 96)) return { cancelled: true, reconciliations: [], ...metrics };
  emitProgress({ percent: 96, message: "Consolidando resultados…", detail: `${reconciliations.length} conciliaciones propuestas` });
  await pause();
  return { cancelled: false, reconciliations, ...metrics };
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
      for (const candidate of found) candidates.push(candidate);
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
    ? tolerance <= .005 ? (Math.abs(difference) <= .005 ? 40 : 0) : 40 - Math.min(5, 5 * Math.abs(difference) / Math.max(tolerance, .01))
    : 0;
  let dateDifference = 0;
  let totalDateDifference = 0;
  let dateComparisonCount = 0;
  for (const systemMovement of systemMovements) {
    for (const bankMovement of bankMovements) {
      const differenceInDays = daysBetween(systemMovement.date, bankMovement.date);
      dateDifference = Math.max(dateDifference, differenceInDays);
      totalDateDifference += differenceInDays;
      dateComparisonCount++;
    }
  }
  const averageDateDifference = dateComparisonCount ? totalDateDifference / dateComparisonCount : 0;
  const dateWithinTolerance = dateDifference <= state.config.dateTolerance;
  const dateScore = dateWithinTolerance
    ? state.config.dateTolerance === 0 ? 20 : 20 * Math.max(.35, 1 - averageDateDifference / (state.config.dateTolerance + 1))
    : 0;
  const similarity = state.config.considerDescription ? groupedDescriptionSimilarity(systemMovements, bankMovements) : 1;
  const refsScore = referenceScore(systemMovements.map(item => item.description), bankMovements.map(item => item.description));
  const totalMembers = systemMovements.length + bankMovements.length;
  const protectedGroupDescription = type !== "one-to-one" && state.config.considerDescription && similarity >= .82;
  const descriptionScore = 30 * (protectedGroupDescription ? 1 : similarity);
  let penalty = 0;
  if (type !== "one-to-one" && !protectedGroupDescription) penalty += Math.min(12, Math.log2(Math.max(1, totalMembers - 2)) * 2);
  if (!protectedGroupDescription && (systemMovements.some(item => isGenericDescription(item.description)) || bankMovements.some(item => isGenericDescription(item.description)))) penalty += 3;
  if (dateDifference > 0) penalty += Math.min(5, dateDifference);
  let score = Math.round(clamp(amountScore + dateScore + descriptionScore + refsScore - penalty, 0, 100));
  const exactAmountAndDate = Math.abs(difference) <= .005 && dateDifference === 0;
  const exactAmountCompatibleDate = Math.abs(difference) <= .005 && dateWithinTolerance;
  if (exactAmountAndDate) score = Math.max(score, type === "one-to-one" ? 90 : 85);
  else if (exactAmountCompatibleDate) score = Math.max(score, type === "one-to-one" ? 85 : 80);
  const exact = Math.abs(difference) <= .005 && dateWithinTolerance && (!state.config.considerDescription || similarity >= .9 || protectedGroupDescription);
  const reasons = [
    Math.abs(difference) <= .005 ? "Los importes coinciden exactamente" : `La diferencia de importe (${formatMoney(Math.abs(difference))}) está dentro de la tolerancia`,
    dateDifference === 0 ? "Las fechas coinciden" : `Las fechas difieren hasta ${dateDifference} día(s)`,
    state.config.considerDescription ? `Similitud de descripciones: ${Math.round(similarity * 100)}%` : "La comparación de descripciones está desactivada"
  ];
  if (refsScore) reasons.push("Se encontraron referencias numéricas coincidentes");
  if (type !== "one-to-one") reasons.push(`La suma de ${type === "one-to-many" ? bankMovements.length : systemMovements.length} movimientos coincide con el otro lado`);
  if (protectedGroupDescription) reasons.push("La descripción comercial coincide; no se aplicó penalización por cantidad de movimientos");
  if (exactAmountAndDate && similarity < .82) reasons.push("Monto y fecha exactos tuvieron prioridad sobre la diferencia de descripción");
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
    protectedGroupDescription,
    exactAmountAndDate,
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

function internalSelectionCanReconcile(movements) {
  if (movements.length < 2) return false;
  const amounts = movements.map(comparisonAmount).filter(amount => Math.abs(amount) > .005);
  if (!amounts.some(amount => amount > 0) || !amounts.some(amount => amount < 0)) return false;
  const net = roundMoney(amounts.reduce((sum, amount) => sum + amount, 0));
  let reference = 0;
  for (const amount of amounts) reference = Math.max(reference, Math.abs(amount));
  return Math.abs(net) <= amountToleranceFor(reference) + .005;
}

function manualSelectionCanReconcile(systemMovements, bankMovements) {
  if (systemMovements.length && bankMovements.length) return true;
  if (systemMovements.length) return internalSelectionCanReconcile(systemMovements);
  if (bankMovements.length) return internalSelectionCanReconcile(bankMovements);
  return false;
}

function calculateInternalReconciliation(movements, sourceKey, manual = false) {
  const net = roundMoney(movements.reduce((sum, movement) => sum + comparisonAmount(movement), 0));
  let reference = 0;
  let minimumTimestamp = Infinity;
  let maximumTimestamp = -Infinity;
  for (const movement of movements) {
    reference = Math.max(reference, Math.abs(comparisonAmount(movement)));
    const timestamp = movement.date.getTime();
    minimumTimestamp = Math.min(minimumTimestamp, timestamp);
    maximumTimestamp = Math.max(maximumTimestamp, timestamp);
  }
  const tolerance = amountToleranceFor(reference);
  const dateDifference = Number.isFinite(minimumTimestamp) ? Math.round((maximumTimestamp - minimumTimestamp) / DATE_MS) : 0;
  let similarity = 0;
  if (!state.config.considerDescription) similarity = 1;
  else if (movements.length >= 2) {
    similarity = movements.reduce((sum, movement, index) => {
      let best = 0;
      movements.forEach((other, otherIndex) => {
        if (index !== otherIndex) best = Math.max(best, descriptionSimilarity(movement.description, other.description));
      });
      return sum + best;
    }, 0) / movements.length;
  }
  const amountWithinTolerance = Math.abs(net) <= tolerance + .005;
  const dateWithinTolerance = dateDifference <= state.config.dateTolerance;
  const amountScore = amountWithinTolerance ? (Math.abs(net) <= .005 ? 45 : 40) : 0;
  const dateScore = dateWithinTolerance ? (dateDifference === 0 ? 25 : 25 * Math.max(.35, 1 - dateDifference / (state.config.dateTolerance + 1))) : 0;
  const penalty = Math.min(12, Math.log2(Math.max(1, movements.length - 1)) * 2);
  const score = Math.round(clamp(amountScore + dateScore + 20 * similarity - penalty, 0, 100));
  const label = sourceKey === "system" ? "sistema contable" : "caja o banco";
  return {
    id: null,
    type: sourceKey === "system" ? "internal-system" : "internal-bank",
    systemIds: sourceKey === "system" ? movements.map(item => item.id) : [],
    bankIds: sourceKey === "bank" ? movements.map(item => item.id) : [],
    systemMovements: sourceKey === "system" ? movements : [],
    bankMovements: sourceKey === "bank" ? movements : [],
    totalSystem: sourceKey === "system" ? net : 0,
    totalBank: sourceKey === "bank" ? net : 0,
    totalBankOriginal: sourceKey === "bank" ? roundMoney(movements.reduce((sum, movement) => sum + movement.amount, 0)) : 0,
    difference: net,
    amountTolerance: tolerance,
    amountWithinTolerance,
    dateWithinTolerance,
    dateDifference,
    descriptionSimilarity: similarity,
    score,
    exact: Math.abs(net) <= .005,
    ambiguous: false,
    alternativeCount: 0,
    totalMembers: movements.length,
    status: manual ? "confirmed" : "possible",
    criterion: manual ? `Compensación interna creada manualmente en ${label}` : `Compensación interna en ${label}`,
    reasons: [
      Math.abs(net) <= .005 ? "Los Débitos y Créditos seleccionados dejan un neto de cero" : `El neto (${formatMoney(net)}) está dentro de la tolerancia`,
      dateDifference === 0 ? "Los movimientos son del mismo día" : `Las fechas abarcan ${dateDifference} día(s)`,
      state.config.considerDescription ? `Similitud interna de descripciones: ${Math.round(similarity * 100)}%` : "La comparación de descripciones está desactivada",
      manual ? "La selección fue confirmada manualmente" : "La compensación se encontró dentro de una sola tabla"
    ],
    observation: manual ? "Compensación interna manual" : "",
    createdAt: new Date(),
    manual
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
  if (!bankMovements.length && systemMovements.length >= 2) return "internal-system";
  if (!systemMovements.length && bankMovements.length >= 2) return "internal-bank";
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

function getUnreservedMovements(sourceKey) {
  const reserved = getReservedIds()[sourceKey];
  return state.sources[sourceKey].movements.filter(item => !reserved.has(item.id));
}

function movementIsOutsidePeriod(movement, period = state.review.periodFilter) {
  const from = String(period?.from || "");
  const to = String(period?.to || "");
  if (!from && !to) return false;
  const dateKey = movement.dateKey || toDateKey(movement.date);
  return Boolean((from && dateKey < from) || (to && dateKey > to));
}

function getExcludedMovements(sourceKey) {
  return getUnreservedMovements(sourceKey).filter(item => movementIsOutsidePeriod(item));
}

function getPendingMovements(sourceKey) {
  return getUnreservedMovements(sourceKey).filter(item => !movementIsOutsidePeriod(item));
}

function calculateSummary() {
  const excludedSystem = getExcludedMovements("system");
  const excludedBank = getExcludedMovements("bank");
  const totalSystem = state.sources.system.movements.length - excludedSystem.length;
  const totalBank = state.sources.bank.movements.length - excludedBank.length;
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
  const pendingAbsoluteAmount = roundMoney([...pendingSystem, ...pendingBank].reduce((sum, item) => sum + Math.abs(comparisonAmount(item)), 0));
  const excludedIds = {
    system: new Set(excludedSystem.map(item => item.id)),
    bank: new Set(excludedBank.map(item => item.id))
  };
  const activeAbsoluteAmount = roundMoney(
    state.sources.system.movements.filter(item => !excludedIds.system.has(item.id)).reduce((sum, item) => sum + Math.abs(comparisonAmount(item)), 0)
    + state.sources.bank.movements.filter(item => !excludedIds.bank.has(item.id)).reduce((sum, item) => sum + Math.abs(comparisonAmount(item)), 0)
  );
  const confirmedAbsoluteAmount = roundMoney(
    state.sources.system.movements.filter(item => confirmedIds.system.has(item.id)).reduce((sum, item) => sum + Math.abs(comparisonAmount(item)), 0)
    + state.sources.bank.movements.filter(item => confirmedIds.bank.has(item.id)).reduce((sum, item) => sum + Math.abs(comparisonAmount(item)), 0)
  );
  const incomingTransfers = (state.transferLog || []).filter(item => item.direction === "in");
  const outgoingTransfers = (state.transferLog || []).filter(item => item.direction === "out");
  return {
    total,
    totalSystem,
    totalBank,
    confirmedCount,
    possibleCount,
    pendingCount,
    pendingSystem,
    pendingBank,
    excludedSystem,
    excludedBank,
    excludedCount: excludedSystem.length + excludedBank.length,
    reconciledAmount: roundMoney(reconciledAmount),
    pendingDifference,
    pendingAbsoluteAmount,
    activeAbsoluteAmount,
    confirmedAbsoluteAmount,
    percentage: total ? confirmedCount / total * 100 : 0,
    amountPercentage: activeAbsoluteAmount ? confirmedAbsoluteAmount / activeAbsoluteAmount * 100 : 0,
    incomingTransfers,
    outgoingTransfers,
    transferCount: incomingTransfers.length + outgoingTransfers.length,
    confirmedReconciliations: confirmedReconciliations.length,
    possibleReconciliations: state.results.reconciliations.filter(item => item.status === "possible").length
  };
}

function syncReviewControls() {
  const typeFilter = document.getElementById("typeFilter");
  const sortFilter = document.getElementById("sortResults");
  const typeLabelElement = document.getElementById("typeFilterLabel");
  const searchInput = document.getElementById("reviewSearch");
  const approveAllButton = document.getElementById("approveAllPossibleBtn");
  const rejectAllButton = document.getElementById("rejectAllPossibleBtn");
  const retryPendingButton = document.getElementById("retryPendingBtn");
  const trimPeriodButton = document.getElementById("trimPeriodBtn");
  const isPending = state.review.tab === "pending";
  const typeOptions = isPending
    ? [["all", "Débitos y créditos"], ["debit", "Sólo débitos"], ["credit", "Sólo créditos"]]
    : [["all", "Todos"], ["one-to-one", "Uno a uno"], ["one-to-many", "Uno a varios"], ["many-to-one", "Varios a uno"], ["internal", "Compensación interna"], ["manual", "Creada manualmente"], ["manual-approved", "Posible aprobada manualmente"]];
  const sortOptions = isPending
    ? [["date-desc", "Fecha más reciente"], ["date-asc", "Fecha más antigua"], ["difference-desc", "Mayor importe"]]
    : [["score-desc", "Mayor confianza"], ["score-asc", "Menor confianza"], ["date-desc", "Fecha más reciente"], ["date-asc", "Fecha más antigua"], ["difference-desc", "Mayor diferencia"]];
  if (!typeOptions.some(([value]) => value === state.review.type)) state.review.type = "all";
  if (!sortOptions.some(([value]) => value === state.review.sort)) state.review.sort = isPending ? "date-desc" : "score-desc";
  typeFilter.innerHTML = typeOptions.map(([value, label]) => `<option value="${value}"${value === state.review.type ? " selected" : ""}>${label}</option>`).join("");
  sortFilter.innerHTML = sortOptions.map(([value, label]) => `<option value="${value}"${value === state.review.sort ? " selected" : ""}>${label}</option>`).join("");
  typeLabelElement.textContent = isPending ? "Tipo de movimiento" : "Tipo de conciliación";
  searchInput.placeholder = isPending ? "Filtrar pendientes por fecha, descripción o importe" : "Buscar fecha, descripción, importe o ID";
  const filteredPossibleCount = state.review.tab === "possible" ? getFilteredReconciliations("possible").length : 0;
  approveAllButton.classList.toggle("hidden", state.review.tab !== "possible" || !filteredPossibleCount);
  approveAllButton.innerHTML = `<i data-lucide="check-check"></i> Aprobar filtrados (${filteredPossibleCount})`;
  rejectAllButton.classList.toggle("hidden", state.review.tab !== "possible" || !filteredPossibleCount);
  rejectAllButton.innerHTML = `<i data-lucide="x-circle"></i> Rechazar filtrados (${filteredPossibleCount})`;
  const pendingCount = getPendingMovements("system").length + getPendingMovements("bank").length;
  const excludedCount = getExcludedMovements("system").length + getExcludedMovements("bank").length;
  retryPendingButton.classList.toggle("hidden", !isPending || pendingCount < 2 || state.processing.running);
  retryPendingButton.innerHTML = `<i data-lucide="scan-search"></i> Búsqueda avanzada (${pendingCount})`;
  trimPeriodButton.classList.toggle("hidden", !isPending || state.processing.running);
  trimPeriodButton.innerHTML = `<i data-lucide="calendar-range"></i> ${excludedCount ? `Período · ${excludedCount} excluidos` : "Recortar período"}`;
}

function approveAllPossible() {
  const possible = getFilteredReconciliations("possible");
  if (!possible.length) return;
  dom.reviewContent.querySelectorAll("[data-observation]").forEach(input => {
    const item = findReconciliation(input.dataset.observation);
    if (item) item.observation = input.value.trim();
  });
  const ambiguousCount = possible.filter(item => item.ambiguous).length;
  const detail = ambiguousCount ? ` ${ambiguousCount} tienen alternativas marcadas como ambiguas.` : "";
  if (!window.confirm(`Se aprobarán ${possible.length} conciliaciones visibles según la búsqueda y el filtro actuales.${detail} ¿Desea continuar?`)) return;
  possible.forEach(item => {
    item.status = "confirmed";
    item.manuallyApproved = true;
    item.manuallyApprovedAt = new Date();
    if (!item.observation) item.observation = "Aprobada mediante acción masiva";
  });
  state.review.page = 1;
  showToast("Conciliaciones aprobadas", `${possible.length} propuestas filtradas pasaron a conciliados.`, "success");
  renderReview();
}

function rejectAllPossible() {
  const possible = getFilteredReconciliations("possible");
  if (!possible.length) return;
  dom.reviewContent.querySelectorAll("[data-observation]").forEach(input => {
    const item = findReconciliation(input.dataset.observation);
    if (item) item.observation = input.value.trim();
  });
  if (!window.confirm(`Se rechazarán ${possible.length} conciliaciones visibles según la búsqueda y el filtro actuales. Sus movimientos volverán a Pendientes. ¿Desea continuar?`)) return;
  possible.forEach(item => {
    rememberRejectedProposal(item, "Propuesta rechazada mediante acción masiva");
    item.status = "rejected";
    if (!item.observation) item.observation = "Rechazada mediante acción masiva";
  });
  state.review.page = 1;
  showToast("Propuestas rechazadas", `${possible.length} propuestas filtradas volvieron a pendientes.`, "error");
  renderReview();
}

function renderReview() {
  syncReviewControls();
  const summary = calculateSummary();
  const cards = [
    ["Movimientos activos", summary.total.toLocaleString("es-UY"), ""],
    ["Conciliados", summary.confirmedCount.toLocaleString("es-UY"), "success"],
    ["Posibles", summary.possibleCount.toLocaleString("es-UY"), "warning"],
    ["Pendientes", summary.pendingCount.toLocaleString("es-UY"), "danger"],
    ["Excluidos", summary.excludedCount.toLocaleString("es-UY"), summary.excludedCount ? "warning" : ""],
    ["Importe conciliado", formatMoney(summary.reconciledAmount), "success"],
    ["Pendiente absoluto", formatMoney(summary.pendingAbsoluteAmount), summary.pendingAbsoluteAmount ? "danger" : "success"],
    ["Diferencia neta", formatMoney(summary.pendingDifference), summary.pendingDifference ? "danger" : "success"],
    ["Avance por filas", `${formatDecimal(summary.percentage, 1)}%`, "success"],
    ["Avance por importe", `${formatDecimal(summary.amountPercentage, 1)}%`, "success"]
  ];
  dom.summaryCards.innerHTML = cards.map(([label, value, type]) => `<div class="summary-card ${type}"><span title="${escapeAttribute(label)}">${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`).join("");
  document.querySelectorAll("[data-review-tab]").forEach(button => button.classList.toggle("active", button.dataset.reviewTab === state.review.tab));
  document.querySelector('[data-tab-count="confirmed"]').textContent = summary.confirmedReconciliations;
  document.querySelector('[data-tab-count="possible"]').textContent = summary.possibleReconciliations;
  document.querySelector('[data-tab-count="pending"]').textContent = summary.pendingCount;
  renderReviewContent();
  refreshIcons(document.querySelector(".review-toolbar"));
  scheduleStatePersistence();
}

function renderReviewContent() {
  if (state.review.tab === "pending") {
    renderPendingReview();
    return;
  }
  let items = getFilteredReconciliations(state.review.tab);
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
  dom.reviewContent.innerHTML = `<div class="table-scroll result-table-scroll"><table class="data-table result-table"><colgroup><col class="result-col-id"><col class="result-col-type"><col class="result-col-description"><col class="result-col-description"><col class="result-col-total"><col class="result-col-total"><col class="result-col-difference"><col class="result-col-score"><col class="result-col-actions"></colgroup><thead><tr><th>ID / Estado</th><th>Tipo</th><th>Sistema contable</th><th>Caja o banco</th><th>Total sistema</th><th>Total caja/banco</th><th>Diferencia</th><th>Confianza</th><th>${isPossible ? "Observación / Acciones" : "Acciones"}</th></tr></thead><tbody>${pageItems.map(item => renderReconciliationRow(item, isPossible)).join("")}</tbody></table></div>`;
  bindResultTableEvents();
  renderPagination(items.length, pageCount);
  refreshIcons(dom.reviewContent);
}

function getFilteredReconciliations(status) {
  let items = state.results.reconciliations.filter(item => item.status === status);
  if (state.review.type !== "all") items = items.filter(item => item.type === state.review.type
    || (state.review.type === "manual" && item.manual)
    || (state.review.type === "manual-approved" && item.manuallyApproved)
    || (state.review.type === "internal" && item.type.startsWith("internal-")));
  if (state.review.search) items = items.filter(item => reconciliationSearchText(item).includes(state.review.search));
  return items;
}

function renderReconciliationRow(item, isPossible) {
  const observationControl = `<label class="observation-control"><span class="visually-hidden">Observación que se exporta</span><input class="observation-input" data-observation="${item.id}" value="${escapeAttribute(item.observation)}" placeholder="Observación (se exporta)" title="Esta observación se incluye en el Excel final"></label>`;
  const observationNote = item.observation ? `<span class="observation-note" title="${escapeAttribute(item.observation)}"><i data-lucide="message-square-text"></i> Observación</span>` : "";
  const actions = isPossible
    ? `<div class="tabular-actions result-actions-possible">${observationControl}<button class="table-action approve" type="button" data-action="approve" data-id="${item.id}"><i data-lucide="check"></i>Aprobar</button><button class="table-action" type="button" data-action="review" data-id="${item.id}" title="Ver el motivo y cambiar los movimientos"><i data-lucide="list-checks"></i>Revisar</button><button class="table-action reject" type="button" data-action="reject" data-id="${item.id}"><i data-lucide="x"></i>Rechazar</button></div>`
    : `<div class="tabular-actions result-actions-confirmed">${observationNote}<button class="table-action" type="button" data-action="detail" data-id="${item.id}"><i data-lucide="eye"></i>Detalle</button><button class="table-action reject" type="button" data-action="unmatch" data-id="${item.id}"><i data-lucide="unlink"></i>Quitar</button></div>`;
  const retryLabel = item.dateAgnosticPass ? "Sin fecha" : item.automaticRelaxedPass ? "Flexible auto" : item.advancedRetry ? "Búsqueda avanzada" : item.relaxedPass ? "Reanálisis" : "";
  const manualApprovalLabel = item.manuallyApproved ? `<span class="manual-approval-badge">Aprobada manualmente</span>` : "";
  return `<tr class="result-row ${item.status}"><td data-label="ID / Estado"><strong>${item.id}</strong>${statusBadge(item.status)}</td><td data-label="Tipo"><span class="type-badge">${typeLabel(item)}</span>${manualApprovalLabel}${retryLabel ? `<span class="retry-badge">${retryLabel}</span>` : ""}</td><td data-label="Sistema contable" class="descriptions">${renderMovementSummaryCell(item.systemMovements)}</td><td data-label="Caja o banco" class="descriptions">${renderMovementSummaryCell(item.bankMovements)}</td><td data-label="Total sistema" class="amount">${formatMoney(item.totalSystem)}</td><td data-label="Total caja/banco" class="amount">${formatMoney(item.totalBank)}</td><td data-label="Diferencia" class="amount ${item.difference ? "negative" : ""}">${formatMoney(item.difference)}</td><td data-label="Confianza"><span class="score-badge ${scoreClass(item.score)}">${item.score}</span></td><td data-label="Acciones">${actions}</td></tr>`;
}

function bindResultTableEvents() {
  dom.reviewContent.querySelectorAll("[data-action]").forEach(button => {
    button.addEventListener("click", () => handleReconciliationAction(button.dataset.action, button.dataset.id));
  });
  dom.reviewContent.querySelectorAll("[data-observation]").forEach(input => {
    input.addEventListener("input", () => {
      const item = findReconciliation(input.dataset.observation);
      if (item) {
        item.observation = input.value.trim();
        scheduleStatePersistence(900);
      }
    });
  });
}

function handleReconciliationAction(action, id) {
  const item = findReconciliation(id);
  if (!item) return;
  if (action === "detail") return openReconciliationDetail(item);
  if (action === "review" || action === "edit") return openEditGroup(item);
  if (action === "approve") {
    item.status = "confirmed";
    item.manuallyApproved = true;
    item.manuallyApprovedAt = new Date();
    item.observation = document.querySelector(`[data-observation="${id}"]`)?.value.trim() || item.observation;
    showToast("Conciliación aprobada", `${id} pasó a conciliados.`, "success");
  }
  if (action === "unmatch") {
    rememberRejectedProposal(item, "Conciliación confirmada quitada por el usuario");
    item.status = "rejected";
    item.observation = item.observation || "Conciliación quitada manualmente";
    showToast("Conciliación quitada", "Sus movimientos volvieron a la lista de pendientes.", "error");
  }
  if (action === "reject") {
    rememberRejectedProposal(item, "Propuesta rechazada por el usuario");
    item.status = "rejected";
    item.observation = document.querySelector(`[data-observation="${id}"]`)?.value.trim() || item.observation;
    showToast("Propuesta rechazada", "Sus movimientos volvieron a la lista de pendientes.", "error");
  }
  state.review.page = 1;
  renderReview();
}

function renderPendingReview() {
  const allPendingSystem = getPendingMovements("system");
  const allPendingBank = getPendingMovements("bank");
  const pendingIds = { system: new Set(allPendingSystem.map(item => item.id)), bank: new Set(allPendingBank.map(item => item.id)) };
  [...state.review.selectedSystem].forEach(id => { if (!pendingIds.system.has(id)) state.review.selectedSystem.delete(id); });
  [...state.review.selectedBank].forEach(id => { if (!pendingIds.bank.has(id)) state.review.selectedBank.delete(id); });
  let system = allPendingSystem;
  let bank = allPendingBank;
  if (state.review.search) {
    system = system.filter(item => movementSearchText(item).includes(state.review.search));
    bank = bank.filter(item => movementSearchText(item).includes(state.review.search));
  }
  if (state.review.type !== "all") {
    system = system.filter(item => item.type === state.review.type);
    bank = bank.filter(item => item.type === state.review.type);
  }
  system = sortPendingMovements(system, state.review.sort);
  bank = sortPendingMovements(bank, state.review.sort);
  const selectedSystemMovements = movementsFromIds("system", state.review.selectedSystem);
  const selectedBankMovements = movementsFromIds("bank", state.review.selectedBank);
  const selectedDifference = roundMoney(selectedSystemMovements.reduce((sum, item) => sum + comparisonAmount(item), 0) - selectedBankMovements.reduce((sum, item) => sum + comparisonAmount(item), 0));
  const selectionReady = manualSelectionCanReconcile(selectedSystemMovements, selectedBankMovements);
  dom.reviewContent.innerHTML = `<div class="pending-wrap"><div class="manual-banner"><div><strong>Conciliación manual</strong><small>Haga clic en cualquier parte de una fila. Puede cruzar ambos lados o seleccionar sólo Débitos y Créditos del mismo lado cuando su neto sea cero.</small><span id="manualSelectionDifference">Diferencia seleccionada: <b>${formatMoney(selectedDifference)}</b></span></div><button id="createManualBtn" class="button button-primary button-small" type="button" ${selectionReady ? "" : "disabled"}><i data-lucide="link"></i> Crear conciliación (${selectedSystemMovements.length} ↔ ${selectedBankMovements.length})</button></div><div class="manual-grid">${renderPendingSide("system", system)}${renderPendingSide("bank", bank)}</div></div>`;
  dom.reviewContent.querySelectorAll("[data-pending-select]").forEach(checkbox => {
    checkbox.addEventListener("click", event => event.stopPropagation());
    checkbox.addEventListener("change", () => {
      const sourceKey = checkbox.dataset.source;
      const selection = sourceKey === "system" ? state.review.selectedSystem : state.review.selectedBank;
      if (checkbox.checked) selection.add(checkbox.value); else selection.delete(checkbox.value);
      checkbox.closest("tr")?.classList.toggle("selected", checkbox.checked);
      checkbox.closest("tr")?.setAttribute("aria-selected", String(checkbox.checked));
      updatePendingSelectionSummary();
    });
  });
  dom.reviewContent.querySelectorAll("[data-pending-row]").forEach(row => {
    const toggle = () => {
      const checkbox = row.querySelector("[data-pending-select]");
      checkbox.checked = !checkbox.checked;
      checkbox.dispatchEvent(new Event("change", { bubbles: true }));
    };
    row.addEventListener("click", event => { if (!event.target.closest("input, button, a")) toggle(); });
    row.addEventListener("keydown", event => { if (event.key === "Enter" || event.key === " ") { event.preventDefault(); toggle(); } });
  });
  dom.reviewContent.querySelectorAll("[data-select-filtered]").forEach(button => button.addEventListener("click", () => {
    const sourceKey = button.dataset.selectFiltered;
    const items = sourceKey === "system" ? system : bank;
    const selection = sourceKey === "system" ? state.review.selectedSystem : state.review.selectedBank;
    items.forEach(item => selection.add(item.id));
    renderPendingReview();
  }));
  dom.reviewContent.querySelectorAll("[data-clear-selection]").forEach(button => button.addEventListener("click", () => {
    const selection = button.dataset.clearSelection === "system" ? state.review.selectedSystem : state.review.selectedBank;
    selection.clear();
    renderPendingReview();
  }));
  document.getElementById("createManualBtn")?.addEventListener("click", createManualReconciliation);
  dom.pagination.innerHTML = `<span>${system.length.toLocaleString("es-UY")} pendientes visibles del sistema · ${bank.length.toLocaleString("es-UY")} de caja o banco</span>`;
  refreshIcons(dom.reviewContent);
}

function renderPendingSide(sourceKey, allItems) {
  const label = sourceKey === "system" ? (state.sources.system.name || "Sistema contable") : (state.sources.bank.name || "Caja o banco");
  const selection = sourceKey === "system" ? state.review.selectedSystem : state.review.selectedBank;
  const selectedMovements = movementsFromIds(sourceKey, selection);
  return `<section class="pending-side"><div class="pending-side-heading"><h3>${escapeHtml(label)} <span>${allItems.length} visibles</span></h3><div><button class="table-action" data-select-filtered="${sourceKey}" type="button"><i data-lucide="list-checks"></i> Seleccionar filtrados</button><button class="table-action" data-clear-selection="${sourceKey}" type="button">Limpiar lado</button></div></div><div class="table-scroll pending-scroll"><table class="data-table pending-table"><colgroup><col class="pending-col-check"><col class="pending-col-row"><col class="pending-col-date"><col class="pending-col-description"><col class="pending-col-type"><col class="pending-col-amount"></colgroup><thead><tr><th></th><th>Fila</th><th>Fecha</th><th>Descripción</th><th>Tipo</th><th class="amount">Monto</th></tr></thead><tbody>${allItems.length ? allItems.map(item => `<tr data-pending-row data-source="${sourceKey}" data-id="${item.id}" tabindex="0" aria-selected="${selection.has(item.id)}" class="${selection.has(item.id) ? "selected" : ""}"><td><input type="checkbox" data-pending-select data-source="${sourceKey}" value="${item.id}" ${selection.has(item.id) ? "checked" : ""} aria-label="Seleccionar fila ${item.row}"></td><td>${item.row}</td><td>${formatDate(item.date)}</td><td class="pending-description">${renderPendingDescription(item)}</td><td><span class="type-badge">${movementTypeLabel(item.type)}</span></td><td class="amount ${item.amount < 0 ? "negative" : ""}">${formatMoney(item.amount)}</td></tr>`).join("") : `<tr><td colspan="6"><div class="empty-state"><div><p>No hay movimientos para los filtros actuales.</p></div></div></td></tr>`}</tbody></table></div><div class="pending-total pending-selected-total" id="pendingTotals-${sourceKey}">${renderSelectedMovementTotals(selectedMovements)}</div></section>`;
}

function renderSelectedMovementTotals(movements) {
  const debit = roundMoney(movements.filter(item => item.type === "debit").reduce((sum, item) => sum + comparisonAmount(item), 0));
  const credit = roundMoney(movements.filter(item => item.type === "credit").reduce((sum, item) => sum + comparisonAmount(item), 0));
  const net = roundMoney(movements.reduce((sum, item) => sum + comparisonAmount(item), 0));
  return `<span><b>${movements.length}</b> seleccionados</span><span>Débito <strong>${formatMoney(debit)}</strong></span><span>Crédito <strong>${formatMoney(credit)}</strong></span><span>Neto <strong>${formatMoney(net)}</strong></span>`;
}

function updatePendingSelectionSummary() {
  const systemMovements = movementsFromIds("system", state.review.selectedSystem);
  const bankMovements = movementsFromIds("bank", state.review.selectedBank);
  const systemTotal = systemMovements.reduce((sum, item) => sum + comparisonAmount(item), 0);
  const bankTotal = bankMovements.reduce((sum, item) => sum + comparisonAmount(item), 0);
  const difference = roundMoney(systemTotal - bankTotal);
  const createButton = document.getElementById("createManualBtn");
  createButton.disabled = !manualSelectionCanReconcile(systemMovements, bankMovements);
  createButton.innerHTML = `<i data-lucide="link"></i> Crear conciliación (${systemMovements.length} ↔ ${bankMovements.length})`;
  document.getElementById("manualSelectionDifference").innerHTML = `Diferencia seleccionada: <b>${formatMoney(difference)}</b>`;
  document.getElementById("pendingTotals-system").innerHTML = renderSelectedMovementTotals(systemMovements);
  document.getElementById("pendingTotals-bank").innerHTML = renderSelectedMovementTotals(bankMovements);
  refreshIcons(createButton);
}

function createManualReconciliation() {
  const systemMovements = movementsFromIds("system", state.review.selectedSystem);
  const bankMovements = movementsFromIds("bank", state.review.selectedBank);
  if (!manualSelectionCanReconcile(systemMovements, bankMovements)) return;
  const reconciliation = systemMovements.length && bankMovements.length
    ? calculateReconciliation(systemMovements, bankMovements, "manual")
    : calculateInternalReconciliation(systemMovements.length ? systemMovements : bankMovements, systemMovements.length ? "system" : "bank", true);
  reconciliation.status = "confirmed";
  reconciliation.manual = true;
  reconciliation.ambiguous = false;
  if (systemMovements.length && bankMovements.length) {
    reconciliation.criterion = "Conciliación creada manualmente por el usuario";
    reconciliation.reasons.push("La selección fue confirmada manualmente");
    reconciliation.observation = "Conciliación manual";
  }
  state.review.rejectedSignatures.delete(reconciliationSignature(reconciliation));
  addReconciliation(reconciliation);
  state.review.selectedSystem.clear();
  state.review.selectedBank.clear();
  showToast("Conciliación manual creada", `${reconciliation.id} fue agregada a conciliados.`, "success");
  renderReview();
}

function openReconciliationDetail(item) {
  renderReconciliationDetail(item);
  dom.detailDialog.showModal();
}

function renderReconciliationDetail(item) {
  document.getElementById("detailTitle").textContent = `Detalle ${item.id}`;
  document.getElementById("detailSubtitle").textContent = `${typeLabel(item)} · ${item.criterion}`;
  const content = document.getElementById("detailContent");
  content.innerHTML = `<div class="detail-grid">${renderDetailSide("Sistema contable", item.systemMovements, "system", item)}${renderDetailSide("Caja o banco", item.bankMovements, "bank", item)}<div class="detail-summary"><div><span>Total sistema</span><strong>${formatMoney(item.totalSystem)}</strong></div><div><span>Total caja/banco</span><strong>${formatMoney(item.totalBank)}</strong></div><div><span>Diferencia</span><strong>${formatMoney(item.difference)}</strong></div><div><span>Confianza</span><strong>${item.score}/100</strong></div><div><span>Estado</span><strong>${item.status === "confirmed" ? "Conciliado" : "Posible"}</strong></div></div><div class="reason-list"><h3>Por qué se propuso</h3><ul>${item.reasons.map(reason => `<li>${escapeHtml(reason)}</li>`).join("")}${item.ambiguous ? "<li>La combinación es ambigua y requiere aprobación manual.</li>" : ""}</ul></div><label class="field detail-observation"><span>Observación <small>se incluye en el Excel exportado</small></span><textarea data-detail-observation placeholder="Agregar una nota para esta conciliación">${escapeHtml(item.observation)}</textarea></label></div>`;
  content.querySelector("[data-detail-observation]")?.addEventListener("input", event => {
    item.observation = event.target.value.trim();
    scheduleStatePersistence(900);
  });
  content.querySelectorAll("[data-detail-remove]").forEach(button => button.addEventListener("click", () => {
    removeMovementFromReconciliationDetail(item, button.dataset.source, button.dataset.detailRemove);
  }));
  refreshIcons(content);
}

function renderDetailSide(title, movements, sourceKey, reconciliation) {
  const allowRemoval = reconciliation.status === "possible";
  return `<section class="detail-side"><h3>${escapeHtml(title)} · ${movements.length} movimiento(s)</h3>${movements.length ? movements.map(item => {
    const canRemove = allowRemoval && canRemoveMovementFromReconciliation(reconciliation, sourceKey, item.id);
    const removeButton = allowRemoval ? `<button class="detail-remove-button" data-detail-remove="${escapeAttribute(item.id)}" data-source="${sourceKey}" type="button" ${canRemove ? "" : "disabled"} title="${canRemove ? "Quitar este movimiento de la propuesta" : "La propuesta debe conservar movimientos válidos en ambos lados"}"><i data-lucide="trash-2"></i> Quitar</button>` : "";
    return `<article class="detail-movement"><header><span>Fila ${item.row} · ${formatDate(item.date)}</span><div class="detail-movement-heading-actions"><strong class="amount ${item.amount < 0 ? "negative" : ""}">${formatMoney(item.amount)}</strong>${removeButton}</div></header><p>${escapeHtml(item.description)}</p><small>${movementTypeLabel(item.type)}${item.status ? ` · Estado original: ${escapeHtml(item.status)}` : ""}</small></article>`;
  }).join("") : `<div class="detail-movement"><p>Sin contraparte: los movimientos se compensan dentro de la otra tabla.</p></div>`}</section>`;
}

function canRemoveMovementFromReconciliation(reconciliation, sourceKey, movementId) {
  const systemMovements = sourceKey === "system" ? reconciliation.systemMovements.filter(item => item.id !== movementId) : reconciliation.systemMovements;
  const bankMovements = sourceKey === "bank" ? reconciliation.bankMovements.filter(item => item.id !== movementId) : reconciliation.bankMovements;
  return manualSelectionCanReconcile(systemMovements, bankMovements);
}

function removeMovementFromReconciliationDetail(reconciliation, sourceKey, movementId) {
  if (reconciliation.status !== "possible" || !canRemoveMovementFromReconciliation(reconciliation, sourceKey, movementId)) return;
  const systemMovements = sourceKey === "system" ? reconciliation.systemMovements.filter(item => item.id !== movementId) : reconciliation.systemMovements;
  const bankMovements = sourceKey === "bank" ? reconciliation.bankMovements.filter(item => item.id !== movementId) : reconciliation.bankMovements;
  const removedMovement = (sourceKey === "system" ? reconciliation.systemMovements : reconciliation.bankMovements).find(item => item.id === movementId);
  const originalSignature = reconciliationSignature(reconciliation);
  const updated = systemMovements.length && bankMovements.length
    ? calculateReconciliation(systemMovements, bankMovements)
    : calculateInternalReconciliation(systemMovements.length ? systemMovements : bankMovements, systemMovements.length ? "system" : "bank");
  const observation = reconciliation.observation;
  rememberRejectedProposal(reconciliation, "Agrupación modificada desde el detalle");
  state.review.rejectedSignatures.delete(reconciliationSignature(updated));
  Object.assign(reconciliation, updated, {
    id: reconciliation.id,
    status: "possible",
    observation,
    ambiguous: false,
    alternativeCount: 0,
    criterion: "Agrupación ajustada desde el detalle; requiere aprobación",
    manual: true
  });
  reconciliation.reasons.push(`Se quitó manualmente la fila ${removedMovement?.row ?? "seleccionada"} desde el detalle`);
  if (originalSignature === reconciliationSignature(reconciliation)) state.review.rejectedSignatures.delete(originalSignature);
  showToast("Movimiento quitado", "La propuesta y sus totales fueron recalculados.", "success");
  renderReview();
  renderReconciliationDetail(reconciliation);
}

function openEditGroup(item) {
  state.review.editingId = item.id;
  const reservedByOthers = { system: new Set(), bank: new Set() };
  state.results.reconciliations.filter(other => other.id !== item.id && (other.status === "confirmed" || other.status === "possible")).forEach(other => {
    other.systemIds.forEach(id => reservedByOthers.system.add(id));
    other.bankIds.forEach(id => reservedByOthers.bank.add(id));
  });
  const currentSystemIds = new Set(item.systemIds);
  const currentBankIds = new Set(item.bankIds);
  state.review.editAvailableSystem = state.sources.system.movements.filter(movement => !reservedByOthers.system.has(movement.id) && (currentSystemIds.has(movement.id) || !movementIsOutsidePeriod(movement)));
  state.review.editAvailableBank = state.sources.bank.movements.filter(movement => !reservedByOthers.bank.has(movement.id) && (currentBankIds.has(movement.id) || !movementIsOutsidePeriod(movement)));
  state.review.editSelectedSystem = new Set(item.systemIds);
  state.review.editSelectedBank = new Set(item.bankIds);
  state.review.editSearchSystem = "";
  state.review.editSearchBank = "";
  document.getElementById("editGroupTitle").textContent = `Revisar ${item.id}`;
  document.getElementById("editGroupSubtitle").textContent = `${typeLabel(item)} · ${item.criterion}`;
  renderEditGroupReason(item);
  renderEditGroupContent();
  document.getElementById("editGroupObservation").value = item.observation || "";
  dom.editGroupDialog.showModal();
}

function renderEditGroupReason(item) {
  const reasons = Array.isArray(item.reasons) ? item.reasons : [];
  document.getElementById("editGroupReason").innerHTML = `<div class="proposal-review-overview"><div><span>Confianza</span><strong>${item.score}/100</strong></div><div><span>Diferencia</span><strong>${formatMoney(item.difference)}</strong></div><div><span>Estado</span><strong>${item.status === "confirmed" ? "Conciliado" : "Posible"}</strong></div></div><div class="reason-list"><h3>Por qué se propuso</h3><ul>${reasons.length ? reasons.map(reason => `<li>${escapeHtml(reason)}</li>`).join("") : "<li>La propuesta fue generada por los criterios de monto, fecha y descripción configurados.</li>"}${item.ambiguous ? "<li>Existen alternativas y requiere revisión manual.</li>" : ""}</ul></div>`;
  refreshIcons(document.getElementById("editGroupReason"));
}

function getFilteredEditMovements(sourceKey) {
  const movements = sourceKey === "system" ? state.review.editAvailableSystem : state.review.editAvailableBank;
  const search = sourceKey === "system" ? state.review.editSearchSystem : state.review.editSearchBank;
  const selected = sourceKey === "system" ? state.review.editSelectedSystem : state.review.editSelectedBank;
  const filtered = search ? movements.filter(item => movementSearchText(item).includes(search)) : [...movements];
  return filtered.sort((left, right) => Number(selected.has(right.id)) - Number(selected.has(left.id)));
}

function renderEditGroupContent() {
  const container = document.getElementById("editGroupContent");
  container.innerHTML = renderEditSide("system") + renderEditSide("bank");
  container.querySelectorAll("[data-edit-search]").forEach(input => input.addEventListener("input", () => {
    const sourceKey = input.dataset.editSearch;
    if (sourceKey === "system") state.review.editSearchSystem = input.value.trim().toLowerCase();
    else state.review.editSearchBank = input.value.trim().toLowerCase();
    renderEditGroupContent();
    const nextInput = document.querySelector(`[data-edit-search="${sourceKey}"]`);
    nextInput?.focus();
    nextInput?.setSelectionRange(nextInput.value.length, nextInput.value.length);
  }));
  container.querySelectorAll("[data-group-select]").forEach(checkbox => {
    checkbox.addEventListener("click", event => event.stopPropagation());
    checkbox.addEventListener("change", () => {
      const selection = checkbox.dataset.source === "system" ? state.review.editSelectedSystem : state.review.editSelectedBank;
      if (checkbox.checked) selection.add(checkbox.value); else selection.delete(checkbox.value);
      checkbox.closest("tr")?.classList.toggle("selected", checkbox.checked);
      checkbox.closest("tr")?.setAttribute("aria-selected", String(checkbox.checked));
      updateEditGroupTotals();
    });
  });
  container.querySelectorAll("[data-group-row]").forEach(row => {
    const toggle = () => {
      const checkbox = row.querySelector("[data-group-select]");
      checkbox.checked = !checkbox.checked;
      checkbox.dispatchEvent(new Event("change", { bubbles: true }));
    };
    row.addEventListener("click", event => { if (!event.target.closest("input, button, a")) toggle(); });
    row.addEventListener("keydown", event => { if (event.key === "Enter" || event.key === " ") { event.preventDefault(); toggle(); } });
  });
  container.querySelectorAll("[data-edit-select-filtered]").forEach(button => button.addEventListener("click", () => {
    const sourceKey = button.dataset.editSelectFiltered;
    const selection = sourceKey === "system" ? state.review.editSelectedSystem : state.review.editSelectedBank;
    getFilteredEditMovements(sourceKey).forEach(item => selection.add(item.id));
    syncEditSelectionDom(sourceKey);
    updateEditGroupTotals();
  }));
  container.querySelectorAll("[data-edit-clear]").forEach(button => button.addEventListener("click", () => {
    const sourceKey = button.dataset.editClear;
    const selection = sourceKey === "system" ? state.review.editSelectedSystem : state.review.editSelectedBank;
    selection.clear();
    syncEditSelectionDom(sourceKey);
    updateEditGroupTotals();
  }));
  updateEditGroupTotals();
  refreshIcons(container);
}

function syncEditSelectionDom(sourceKey) {
  const selection = sourceKey === "system" ? state.review.editSelectedSystem : state.review.editSelectedBank;
  document.querySelectorAll(`[data-group-select][data-source="${sourceKey}"]`).forEach(checkbox => {
    const selected = selection.has(checkbox.value);
    checkbox.checked = selected;
    checkbox.closest("tr")?.classList.toggle("selected", selected);
    checkbox.closest("tr")?.setAttribute("aria-selected", String(selected));
  });
}

function renderEditSide(sourceKey) {
  const label = sourceKey === "system" ? "Sistema contable" : "Caja o banco";
  const movements = getFilteredEditMovements(sourceKey);
  const selected = sourceKey === "system" ? state.review.editSelectedSystem : state.review.editSelectedBank;
  const selectedMovements = movementsFromIds(sourceKey, selected);
  const search = sourceKey === "system" ? state.review.editSearchSystem : state.review.editSearchBank;
  return `<section class="pending-side edit-pending-side"><h3>${label}<span>${movements.length} visibles</span></h3><div class="edit-side-tools"><label class="search-field"><i data-lucide="search"></i><input data-edit-search="${sourceKey}" type="search" value="${escapeAttribute(search)}" placeholder="Filtrar este lado"></label><div><button class="table-action" data-edit-select-filtered="${sourceKey}" type="button"><i data-lucide="list-checks"></i> Seleccionar filtrados</button><button class="table-action" data-edit-clear="${sourceKey}" type="button">Limpiar</button></div></div><div class="table-scroll" data-edit-scroll="${sourceKey}"><table class="data-table pending-table"><colgroup><col class="pending-col-check"><col class="pending-col-row"><col class="pending-col-date"><col class="pending-col-description"><col class="pending-col-type"><col class="pending-col-amount"></colgroup><thead><tr><th></th><th>Fila</th><th>Fecha</th><th>Descripción</th><th>Tipo</th><th class="amount">Monto</th></tr></thead><tbody>${movements.length ? movements.map(item => `<tr data-group-row tabindex="0" aria-selected="${selected.has(item.id)}" class="${selected.has(item.id) ? "selected" : ""}"><td><input data-group-select data-source="${sourceKey}" type="checkbox" value="${item.id}" ${selected.has(item.id) ? "checked" : ""}></td><td>${item.row}</td><td>${formatDate(item.date)}</td><td class="pending-description">${renderPendingDescription(item)}</td><td><span class="type-badge">${movementTypeLabel(item.type)}</span></td><td class="amount ${item.amount < 0 ? "negative" : ""}">${formatMoney(item.amount)}</td></tr>`).join("") : `<tr><td colspan="6">No hay movimientos para este filtro.</td></tr>`}</tbody></table></div><div class="pending-total pending-selected-total edit-selected-total" id="editSideTotals-${sourceKey}">${renderSelectedMovementTotals(selectedMovements)}</div></section>`;
}

function updateEditGroupTotals() {
  const selected = getEditGroupSelections();
  const systemMovements = movementsFromIds("system", selected.system);
  const bankMovements = movementsFromIds("bank", selected.bank);
  const totalSystem = systemMovements.reduce((sum, item) => sum + comparisonAmount(item), 0);
  const totalBank = bankMovements.reduce((sum, item) => sum + comparisonAmount(item), 0);
  const difference = roundMoney(totalSystem - totalBank);
  document.getElementById("editGroupTotals").innerHTML = `<div>${renderEditSideTotals("Sistema", systemMovements)}</div><div>${renderEditSideTotals("Caja / banco", bankMovements)}</div><div class="difference-total"><span>Diferencia neta</span><strong>${formatMoney(difference)}</strong><small>${systemMovements.length} ↔ ${bankMovements.length} movimientos</small></div>`;
  const systemFooter = document.getElementById("editSideTotals-system");
  const bankFooter = document.getElementById("editSideTotals-bank");
  if (systemFooter) systemFooter.innerHTML = renderSelectedMovementTotals(systemMovements);
  if (bankFooter) bankFooter.innerHTML = renderSelectedMovementTotals(bankMovements);
  const selectionIsValid = manualSelectionCanReconcile(systemMovements, bankMovements);
  document.getElementById("saveGroupBtn").disabled = !selectionIsValid;
  document.getElementById("approveGroupBtn").disabled = !selectionIsValid;
}

function renderEditSideTotals(label, movements) {
  const debit = roundMoney(movements.filter(item => item.type === "debit").reduce((sum, item) => sum + comparisonAmount(item), 0));
  const credit = roundMoney(movements.filter(item => item.type === "credit").reduce((sum, item) => sum + comparisonAmount(item), 0));
  const net = roundMoney(movements.reduce((sum, item) => sum + comparisonAmount(item), 0));
  return `<span>${label} · ${movements.length}</span><small>Débito <b>${formatMoney(debit)}</b> · Crédito <b>${formatMoney(credit)}</b></small><strong>${formatMoney(net)}</strong>`;
}

function getEditGroupSelections() {
  return { system: new Set(state.review.editSelectedSystem), bank: new Set(state.review.editSelectedBank) };
}

function commitEditedGroup(targetStatus) {
  const current = findReconciliation(state.review.editingId);
  if (!current) return false;
  const selected = getEditGroupSelections();
  const systemMovements = movementsFromIds("system", selected.system);
  const bankMovements = movementsFromIds("bank", selected.bank);
  if (!manualSelectionCanReconcile(systemMovements, bankMovements)) return false;
  const originalSignature = reconciliationSignature(current);
  const selectedSignature = reconciliationSignatureFromIds([...selected.system], [...selected.bank]);
  const selectionChanged = originalSignature !== selectedSignature;
  const observation = document.getElementById("editGroupObservation").value.trim();
  if (selectionChanged) {
    const updated = systemMovements.length && bankMovements.length
      ? calculateReconciliation(systemMovements, bankMovements)
      : calculateInternalReconciliation(systemMovements.length ? systemMovements : bankMovements, systemMovements.length ? "system" : "bank");
    rememberRejectedProposal(current, "Agrupación reemplazada al cambiar sus movimientos");
    state.review.rejectedSignatures.delete(reconciliationSignature(updated));
    Object.assign(current, updated, {
      id: current.id,
      status: targetStatus,
      observation,
      ambiguous: false,
      alternativeCount: 0,
      criterion: targetStatus === "confirmed" ? "Agrupación editada y aprobada manualmente" : "Agrupación editada manualmente; requiere aprobación",
      manual: true
    });
    current.reasons.push("La agrupación original fue modificada manualmente");
  } else {
    current.status = targetStatus;
    current.observation = observation;
  }
  if (targetStatus === "confirmed") {
    current.manuallyApproved = true;
    current.manuallyApprovedAt = new Date();
    current.ambiguous = false;
    current.alternativeCount = 0;
    if (!current.reasons.includes("La propuesta fue aprobada manualmente durante la revisión")) current.reasons.push("La propuesta fue aprobada manualmente durante la revisión");
  }
  dom.editGroupDialog.close();
  renderReview();
  return true;
}

function saveEditedGroup() {
  if (!commitEditedGroup("possible")) return;
  showToast("Propuesta guardada", "Los movimientos, totales y la observación quedaron actualizados.", "success");
}

function approveEditedGroup() {
  const current = findReconciliation(state.review.editingId);
  if (!current || !commitEditedGroup("confirmed")) return;
  showToast("Conciliación aprobada", `${current.id} pasó a conciliados.`, "success");
}

function rejectEditedGroup() {
  const current = findReconciliation(state.review.editingId);
  if (!current) return;
  current.observation = document.getElementById("editGroupObservation").value.trim();
  rememberRejectedProposal(current, "Propuesta rechazada desde la revisión detallada");
  current.status = "rejected";
  dom.editGroupDialog.close();
  state.review.page = 1;
  renderReview();
  showToast("Propuesta rechazada", "Sus movimientos volvieron a la lista de pendientes.", "error");
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
  const firstDate = item => {
    let minimum = Infinity;
    for (const movement of item.systemMovements) minimum = Math.min(minimum, movement.date.getTime());
    for (const movement of item.bankMovements) minimum = Math.min(minimum, movement.date.getTime());
    return minimum;
  };
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
  return [item.id, item.type, item.manuallyApproved ? "posible aprobada manualmente" : "", item.criterion, item.score, item.difference, item.observation, ...item.systemMovements.flatMap(m => [m.row, m.dateKey, m.description, m.amount]), ...item.bankMovements.flatMap(m => [m.row, m.dateKey, m.description, m.amount])].join(" ").toLowerCase();
}

function movementSearchText(item) {
  return [item.id, item.row, item.dateKey, formatDate(item.date), item.description, item.amount, item.type, item.transferOriginAccount].join(" ").toLowerCase();
}

function renderPendingDescription(item) {
  const origin = item.transferOriginAccount
    ? `<small class="movement-origin-badge"><i data-lucide="arrow-left-right"></i> Movido desde ${escapeHtml(item.transferOriginAccount)}</small>`
    : "";
  return `<span title="${escapeAttribute(item.description)}">${escapeHtml(item.description)}</span>${origin}`;
}

function summarizeMovements(movements) {
  if (!movements.length) return { title: "Sin contraparte", subtitle: "Compensación dentro de la otra tabla" };
  if (movements.length === 1) return { title: movements[0].description, subtitle: `Fila ${movements[0].row} · ${formatDate(movements[0].date)}` };
  return { title: `${movements.length} movimientos`, subtitle: `${movements.map(item => `F${item.row}`).join(", ")} · ${formatDateRange(movements)}` };
}

function renderMovementSummaryCell(movements) {
  const summary = summarizeMovements(movements);
  return `<strong title="${escapeAttribute(summary.title)}">${escapeHtml(summary.title)}</strong><span class="movement-references">${escapeHtml(summary.subtitle)}</span>`;
}

function typeLabel(item) {
  if (item.type === "internal-system") return "Compensación interna · Sistema";
  if (item.type === "internal-bank") return "Compensación interna · Caja/banco";
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

function createStateSnapshot() {
  const serializeMovement = movement => ({
    id: movement.id,
    source: movement.source,
    row: movement.row,
    date: movement.date instanceof Date ? movement.date.toISOString() : movement.date,
    dateKey: movement.dateKey,
    originalDate: movement.originalDate,
    description: movement.description,
    originalDescription: movement.originalDescription,
    amount: movement.amount,
    debitAmount: movementDebitCredit(movement).debit,
    creditAmount: movementDebitCredit(movement).credit,
    originalAmount: movement.originalAmount,
    type: movement.type,
    status: movement.status,
    rawValues: Array.isArray(movement.rawValues) ? movement.rawValues : [],
    transferOriginId: movement.transferOriginId || "",
    transferOriginAccount: movement.transferOriginAccount || "",
    transferHistory: Array.isArray(movement.transferHistory) ? movement.transferHistory : []
  });
  const serializeSource = source => ({
    key: source.key,
    label: source.label,
    name: source.name,
    fileName: source.file?.name || source.fileName || `${source.label}.xlsx`,
    fileSize: Number(source.file?.size) || 0,
    selectedSheet: source.selectedSheet,
    matrix: source.matrix,
    headers: source.headers,
    previewRows: source.rows.slice(0, 10),
    restoredRowCount: source.restoredRowCount || source.rows.length || source.movements.length,
    importRangeInfo: source.importRangeInfo,
    headerRowIndex: source.headerRowIndex,
    headerRowNumber: source.headerRowNumber,
    dataStartRow: source.dataStartRow,
    dataEndRow: source.dataEndRow,
    detectedBlocks: source.detectedBlocks,
    columnBlock: source.columnBlock,
    dateFrom: source.dateFrom,
    dateTo: source.dateTo,
    reportPeriod: source.reportPeriod,
    periodSource: source.periodSource,
    excludedDescriptions: source.excludedDescriptions,
    filteredRowsCount: source.filteredRowsCount,
    formatMode: source.formatMode,
    detectedFormat: source.detectedFormat,
    positiveMeaning: source.positiveMeaning,
    splitConvention: source.splitConvention,
    splitConventionLocked: source.splitConventionLocked,
    allowBoth: source.allowBoth,
    mapping: { ...source.mapping },
    movements: source.movements.map(serializeMovement),
    invalidRows: source.invalidRows,
    validationWarnings: source.validationWarnings
  });
  const reconciliations = state.results.reconciliations.map(item => {
    const { systemMovements, bankMovements, ...serialized } = item;
    return {
      ...serialized,
      createdAt: item.createdAt instanceof Date ? item.createdAt.toISOString() : item.createdAt,
      approvedAt: item.approvedAt instanceof Date ? item.approvedAt.toISOString() : item.approvedAt,
      manuallyApprovedAt: item.manuallyApprovedAt instanceof Date ? item.manuallyApprovedAt.toISOString() : item.manuallyApprovedAt
    };
  });
  return {
    schemaVersion: STATE_SCHEMA_VERSION,
    appBuild: APP_BUILD,
    savedAt: new Date().toISOString(),
    workspace: {
      id: state.workspace?.id || createWorkspaceState().id,
      name: getCurrentWorkspaceName()
    },
    transferLog: (state.transferLog || []).map(item => ({ ...item, at: item.at instanceof Date ? item.at.toISOString() : item.at })),
    exportFileName: document.getElementById("exportFileName")?.value.trim() || "",
    config: { ...state.config },
    sources: {
      system: serializeSource(state.sources.system),
      bank: serializeSource(state.sources.bank)
    },
    results: {
      ...state.results,
      processingAt: state.results.processingAt instanceof Date ? state.results.processingAt.toISOString() : state.results.processingAt,
      retryPasses: (state.results.retryPasses || []).map(pass => ({ ...pass, at: pass.at instanceof Date ? pass.at.toISOString() : pass.at })),
      reconciliations
    },
    review: {
      tab: state.review.tab,
      search: state.review.search,
      type: state.review.type,
      sort: state.review.sort,
      periodFilter: {
        ...state.review.periodFilter,
        appliedAt: state.review.periodFilter?.appliedAt instanceof Date ? state.review.periodFilter.appliedAt.toISOString() : state.review.periodFilter?.appliedAt
      },
      rejectedSignatures: [...state.review.rejectedSignatures],
      rejectedProposals: state.review.rejectedProposals.map(item => ({ ...item, at: item.at instanceof Date ? item.at.toISOString() : item.at }))
    }
  };
}

function snapshotHasProgress(snapshot) {
  return Boolean(snapshot?.sources?.system?.movements?.length
    || snapshot?.sources?.bank?.movements?.length
    || snapshot?.results?.reconciliations?.length);
}

function reviveStoredDate(value) {
  if (!value) return null;
  const date = value instanceof Date ? value : new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function restoreStateSnapshot(snapshot) {
  if (!snapshot || Number(snapshot.schemaVersion) !== STATE_SCHEMA_VERSION) throw new Error("La versión del estado guardado no es compatible con esta aplicación.");
  if (!snapshotHasProgress(snapshot)) throw new Error("El archivo no contiene una conciliación guardada.");
  state.persistence.restoring = true;
  terminateProcessingWorker();
  state.processing = { cancelled: false, running: false, worker: null, jobId: null };
  state.config = { ...DEFAULT_CONFIG, ...(snapshot.config || {}) };
  state.workspace = createWorkspaceState(snapshot.workspace?.id, snapshot.workspace?.name || snapshot.exportFileName || "");
  state.accountTransfer = createAccountTransferState();
  state.transferLog = (snapshot.transferLog || []).map(item => ({ ...item, at: reviveStoredDate(item.at) || item.at }));
  const restoreSource = (sourceKey, label) => {
    const saved = snapshot.sources?.[sourceKey] || {};
    const movements = (saved.movements || []).map(item => ({ ...item, date: reviveStoredDate(item.date) }));
    const selectedSheet = saved.selectedSheet || "Datos restaurados";
    Object.assign(state.sources[sourceKey], createEmptySource(sourceKey, label), saved, {
      file: { name: saved.fileName || `${label}.xlsx`, size: Number(saved.fileSize) || 0 },
      workbook: null,
      sheetNames: [selectedSheet],
      selectedSheet,
      matrix: Array.isArray(saved.matrix) ? saved.matrix : [],
      headers: Array.isArray(saved.headers) ? saved.headers : [],
      rows: Array.isArray(saved.previewRows) ? saved.previewRows : [],
      movements,
      invalidRows: Array.isArray(saved.invalidRows) ? saved.invalidRows : [],
      validationErrors: [],
      validationWarnings: ["Datos restaurados. Para cambiar hoja o columnas, vuelva a cargar el archivo original."],
      isValid: movements.length > 0,
      restoredState: true,
      restoredRowCount: Number(saved.restoredRowCount) || movements.length
    });
  };
  restoreSource("system", "Sistema contable");
  restoreSource("bank", "Caja o banco");
  const systemById = new Map(state.sources.system.movements.map(item => [item.id, item]));
  const bankById = new Map(state.sources.bank.movements.map(item => [item.id, item]));
  const savedResults = snapshot.results || {};
  const reconciliations = (savedResults.reconciliations || []).map(item => ({
    ...item,
    createdAt: reviveStoredDate(item.createdAt) || new Date(),
    approvedAt: reviveStoredDate(item.approvedAt),
    manuallyApprovedAt: reviveStoredDate(item.manuallyApprovedAt),
    systemMovements: (item.systemIds || []).map(id => systemById.get(id)).filter(Boolean),
    bankMovements: (item.bankIds || []).map(id => bankById.get(id)).filter(Boolean)
  }));
  state.results = {
    ...createEmptyResults(),
    ...savedResults,
    processingAt: reviveStoredDate(savedResults.processingAt)
      || (reconciliations.length ? reviveStoredDate(snapshot.savedAt) || new Date() : null),
    retryPasses: (savedResults.retryPasses || []).map(pass => ({ ...pass, at: reviveStoredDate(pass.at) })),
    reconciliations,
    nextId: Math.max(Number(savedResults.nextId) || 1, reconciliations.length + 1)
  };
  const savedReview = snapshot.review || {};
  state.review = {
    tab: ["confirmed", "possible", "pending"].includes(savedReview.tab) ? savedReview.tab : "confirmed",
    search: String(savedReview.search || ""),
    type: String(savedReview.type || "all"),
    sort: String(savedReview.sort || "score-desc"),
    page: 1,
    selectedSystem: new Set(),
    selectedBank: new Set(),
    editingId: null,
    editAvailableSystem: [],
    editAvailableBank: [],
    editSelectedSystem: new Set(),
    editSelectedBank: new Set(),
    editSearchSystem: "",
    editSearchBank: "",
    periodFilter: {
      from: String(savedReview.periodFilter?.from || ""),
      to: String(savedReview.periodFilter?.to || ""),
      appliedAt: reviveStoredDate(savedReview.periodFilter?.appliedAt)
    },
    rejectedSignatures: new Set(savedReview.rejectedSignatures || []),
    rejectedProposals: (savedReview.rejectedProposals || []).map(item => ({ ...item, at: reviveStoredDate(item.at) }))
  };
  populateConfigForm();
  renderSourceEditor("system");
  renderSourceEditor("bank");
  renderConfigSummary();
  document.getElementById("reviewSearch").value = state.review.search;
  document.getElementById("exportFileName").value = String(snapshot.exportFileName || "").replace(/\.xlsx$/i, "");
  const targetStep = state.results.processingAt || reconciliations.length
    ? 6
    : state.sources.system.isValid && state.sources.bank.isValid ? 4
      : state.sources.system.isValid ? 3 : 2;
  state.maxVisitedStep = Math.max(targetStep, targetStep === 6 ? 7 : targetStep);
  if (targetStep === 6) renderReview();
  goToStep(targetStep);
  state.persistence.restoring = false;
  state.persistence.lastSavedAt = reviveStoredDate(snapshot.savedAt) || new Date();
  updateLocalSaveStatus("Progreso restaurado");
}

function createApplicationStateSheet(snapshot) {
  const json = JSON.stringify(snapshot);
  const rows = [["CONCILIAPP_STATE", STATE_SCHEMA_VERSION], ["Parte", "Contenido"]];
  for (let offset = 0, part = 1; offset < json.length; offset += STATE_CHUNK_SIZE, part++) rows.push([part, json.slice(offset, offset + STATE_CHUNK_SIZE)]);
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  worksheet["!cols"] = [{ wch: 12 }, { wch: 80 }];
  return worksheet;
}

function extractStateSnapshotFromWorkbook(workbook) {
  const worksheet = workbook.Sheets?.[STATE_SHEET_NAME];
  if (!worksheet) throw new Error("Este Excel no contiene el estado interno de una conciliación exportada por la aplicación.");
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true, defval: "" });
  if (rows[0]?.[0] !== "CONCILIAPP_STATE") throw new Error("La hoja de estado no tiene un formato reconocido.");
  const json = rows.slice(2).sort((left, right) => Number(left[0]) - Number(right[0])).map(row => String(row[1] || "")).join("");
  return JSON.parse(json);
}

function hideWorkbookSheet(workbook, sheetName) {
  const index = workbook.SheetNames.indexOf(sheetName);
  if (index < 0) return;
  workbook.Workbook ||= {};
  workbook.Workbook.Sheets ||= [];
  while (workbook.Workbook.Sheets.length < workbook.SheetNames.length) workbook.Workbook.Sheets.push({});
  workbook.Workbook.Sheets[index] = { ...workbook.Workbook.Sheets[index], Hidden: 1 };
}

async function loadPreviousReconciliationFile(file) {
  if (!/\.xlsx$/i.test(file.name || "")) {
    showToast("Archivo no admitido", "Seleccione un XLSX exportado previamente por esta aplicación.", "error");
    return;
  }
  if (typeof XLSX === "undefined") {
    showToast("Lector de Excel no disponible", "Vuelva a abrir la aplicación con conexión a Internet.", "error");
    return;
  }
  try {
    updateLocalSaveStatus("Abriendo conciliación…");
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false, dense: false });
    const snapshot = extractStateSnapshotFromWorkbook(workbook);
    snapshot.workspace ||= {};
    snapshot.workspace.id ||= `workspace-${Date.now()}-${Math.random().toString(16).slice(2)}`;
    snapshot.workspace.name ||= normalizeExportFileName(file.name).replace(/\.xlsx$/i, "");
    restoreStateSnapshot(snapshot);
    const loadedFileName = normalizeExportFileName(file.name).replace(/\.xlsx$/i, "");
    document.getElementById("exportFileName").value = loadedFileName;
    const restoredSnapshot = createStateSnapshot();
    await persistSnapshot(restoredSnapshot);
    await persistSnapshot(restoredSnapshot, workspaceStorageKey(restoredSnapshot.workspace.id));
    showToast("Conciliación restaurada", "Se recuperaron los conciliados, posibles, pendientes y parámetros guardados.", "success", 7000);
  } catch (error) {
    console.error(error);
    state.persistence.restoring = false;
    updateLocalSaveStatus("Guardado local automático");
    showToast("No se pudo restaurar", error.message || "El archivo no es una conciliación compatible.", "error", 9000);
  }
}

function openLocalStateDatabase() {
  return new Promise((resolve, reject) => {
    if (!window.indexedDB) return reject(new Error("IndexedDB no está disponible."));
    const request = window.indexedDB.open(LOCAL_DATABASE_NAME, LOCAL_DATABASE_VERSION);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(LOCAL_STATE_STORE)) request.result.createObjectStore(LOCAL_STATE_STORE);
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("No se pudo abrir el almacenamiento local."));
  });
}

function workspaceStorageKey(workspaceId) {
  return `workspace:${String(workspaceId || "").trim()}`;
}

async function persistSnapshot(snapshot, key = LOCAL_STATE_KEY) {
  const database = await openLocalStateDatabase();
  await new Promise((resolve, reject) => {
    const transaction = database.transaction(LOCAL_STATE_STORE, "readwrite");
    transaction.objectStore(LOCAL_STATE_STORE).put(snapshot, key);
    transaction.oncomplete = resolve;
    transaction.onerror = () => reject(transaction.error || new Error("No se pudo guardar el progreso."));
    transaction.onabort = () => reject(transaction.error || new Error("Se canceló el guardado local."));
  });
  database.close();
}

async function readPersistedSnapshot(key = LOCAL_STATE_KEY) {
  const database = await openLocalStateDatabase();
  const snapshot = await new Promise((resolve, reject) => {
    const transaction = database.transaction(LOCAL_STATE_STORE, "readonly");
    const request = transaction.objectStore(LOCAL_STATE_STORE).get(key);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error || new Error("No se pudo leer el progreso local."));
  });
  database.close();
  return snapshot;
}

async function clearPersistedState() {
  if (!window.indexedDB) return;
  const database = await openLocalStateDatabase();
  await new Promise((resolve, reject) => {
    const transaction = database.transaction(LOCAL_STATE_STORE, "readwrite");
    transaction.objectStore(LOCAL_STATE_STORE).delete(LOCAL_STATE_KEY);
    transaction.oncomplete = resolve;
    transaction.onerror = () => reject(transaction.error || new Error("No se pudo borrar el progreso local."));
  });
  database.close();
}

async function readSavedWorkspaceSnapshots() {
  const database = await openLocalStateDatabase();
  const entries = await new Promise((resolve, reject) => {
    const transaction = database.transaction(LOCAL_STATE_STORE, "readonly");
    const store = transaction.objectStore(LOCAL_STATE_STORE);
    const keyRequest = store.getAllKeys();
    const valueRequest = store.getAll();
    transaction.oncomplete = () => resolve((keyRequest.result || []).map((key, index) => ({ key, snapshot: valueRequest.result?.[index] })).filter(entry => String(entry.key).startsWith("workspace:") && snapshotHasProgress(entry.snapshot)));
    transaction.onerror = () => reject(transaction.error || new Error("No se pudieron leer las cuentas guardadas."));
    transaction.onabort = () => reject(transaction.error || new Error("Se canceló la lectura de cuentas guardadas."));
  });
  database.close();
  return entries;
}

function scheduleStatePersistence(delay = 650) {
  if (state.persistence.restoring || !window.indexedDB) return;
  if (state.persistence.saveTimer) window.clearTimeout(state.persistence.saveTimer);
  updateLocalSaveStatus("Cambios pendientes…");
  state.persistence.saveTimer = window.setTimeout(() => {
    state.persistence.saveTimer = null;
    persistCurrentState();
  }, delay);
}

async function persistCurrentState() {
  if (state.persistence.restoring || !window.indexedDB) return;
  try {
    const snapshot = createStateSnapshot();
    if (!snapshotHasProgress(snapshot)) {
      await clearPersistedState();
      updateLocalSaveStatus("Sin progreso guardado");
      return;
    }
    await persistSnapshot(snapshot);
    await persistSnapshot(snapshot, workspaceStorageKey(snapshot.workspace.id));
    state.persistence.lastSavedAt = new Date();
    state.persistence.saveErrorShown = false;
    updateLocalSaveStatus("Progreso guardado localmente");
  } catch (error) {
    console.error(error);
    updateLocalSaveStatus("No se pudo guardar");
    if (!state.persistence.saveErrorShown) {
      state.persistence.saveErrorShown = true;
      showToast("Guardado local no disponible", "El progreso sigue abierto, pero este navegador no permitió guardarlo automáticamente.", "error", 8000);
    }
  }
}

async function restorePersistedStateOnStartup() {
  if (!window.indexedDB) {
    updateLocalSaveStatus("Guardado local no disponible");
    return;
  }
  try {
    const snapshot = await readPersistedSnapshot();
    if (!snapshotHasProgress(snapshot)) return;
    restoreStateSnapshot(snapshot);
    showToast("Progreso recuperado", "La conciliación se restauró automáticamente después de recargar la página.", "success", 7000);
  } catch (error) {
    console.error(error);
    state.persistence.restoring = false;
    updateLocalSaveStatus("No se pudo restaurar");
  }
}

function updateLocalSaveStatus(message) {
  const element = document.getElementById("localSaveStatus");
  if (element) element.textContent = message;
}

function renderExportSummary() {
  const summary = calculateSummary();
  document.getElementById("exportSummary").innerHTML = [
    ["Conciliaciones", summary.confirmedReconciliations],
    ["Posibles", summary.possibleReconciliations],
    ["Pendientes", summary.pendingCount],
    ["Excluidos por período", summary.excludedCount],
    ["Movimientos recibidos", summary.incomingTransfers.length],
    ["Movimientos enviados", summary.outgoingTransfers.length],
    ["Importe conciliado", formatMoney(summary.reconciledAmount)],
    ["Pendiente absoluto", formatMoney(summary.pendingAbsoluteAmount)],
    ["Diferencia neta", formatMoney(summary.pendingDifference)],
    ["Avance por filas", `${formatDecimal(summary.percentage, 1)}%`],
    ["Avance por importe", `${formatDecimal(summary.amountPercentage, 1)}%`]
  ].map(([label, value]) => `<div class="export-stat"><span>${escapeHtml(label)}</span><strong>${escapeHtml(String(value))}</strong></div>`).join("");
  const fileNameInput = document.getElementById("exportFileName");
  if (!fileNameInput.value.trim()) fileNameInput.value = `conciliacion_${toFileTimestamp()}`;
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
    const retryPasses = Array.isArray(state.results.retryPasses) ? state.results.retryPasses : [];
    const lastRetry = retryPasses.at(-1);
    const summaryRows = [
      ["RESUMEN DE CONCILIACIÓN CONTABLE", ""],
      ["Fecha y hora de procesamiento", formatDateTime(processingDate)],
      ["Tabla del sistema", state.sources.system.name || state.sources.system.label],
      ["Tabla de caja o banco", state.sources.bank.name || state.sources.bank.label],
      ["Movimientos activos del sistema", summary.totalSystem],
      ["Movimientos activos de caja o banco", summary.totalBank],
      ["Movimientos excluidos por período", summary.excludedCount],
      ["Período activo", state.review.periodFilter?.from || state.review.periodFilter?.to ? `${state.review.periodFilter.from ? `Desde ${formatDateInput(state.review.periodFilter.from)}` : "Sin fecha inicial"} · ${state.review.periodFilter.to ? `Hasta ${formatDateInput(state.review.periodFilter.to)}` : "Sin fecha final"}` : "Sin recorte"],
      ["Movimientos conciliados", summary.confirmedCount],
      ["Movimientos en posibles conciliaciones", summary.possibleCount],
      ["Movimientos pendientes", summary.pendingCount],
      ["Movimientos recibidos desde otras cuentas", summary.incomingTransfers.length],
      ["Movimientos enviados a otras cuentas", summary.outgoingTransfers.length],
      ["Importe conciliado", summary.reconciledAmount],
      ["Importe pendiente absoluto", summary.pendingAbsoluteAmount],
      ["Diferencia pendiente neta", summary.pendingDifference],
      ["Porcentaje de conciliación por filas", summary.percentage / 100],
      ["Porcentaje de conciliación por importe", summary.amountPercentage / 100],
      ["", ""],
      ["PARÁMETROS UTILIZADOS", ""],
      ["Tolerancia de fechas (días)", state.config.dateTolerance],
      ["Tolerancia absoluta de monto", state.config.amountAbsTolerance],
      ["Tolerancia porcentual de monto", state.config.amountPercentTolerance / 100],
      ["Movimientos máximos por lado en una agrupación", state.config.maxGroupSize],
      ["Tope de candidatos uno a uno", state.config.maxPairComparisons],
      ["Intentos de suma por movimiento", state.config.maxCombinations],
      ["Comparar signos obligatoriamente", yesNo(state.config.requireSameSign)],
      ["Convención invertida entre tablas", yesNo(state.config.invertBetweenTables)],
      ["Permitir Débitos y Créditos en una agrupación", yesNo(state.config.allowMixedGroupSigns)],
      ["Buscar compensaciones dentro de una tabla", yesNo(state.config.searchInternalOffsets)],
      ["Convención del sistema", splitConventionLabel(state.sources.system.splitConvention)],
      ["Convención de caja o banco", splitConventionLabel(state.sources.bank.splitConvention)],
      ["Nivel automático", state.config.autoThreshold],
      ["Nivel posible", state.config.possibleThreshold],
      ["Motor de procesamiento", String(state.results.engineMode).includes("worker") ? "Web Worker indexado" : "Motor indexado compatible"],
      ["Parejas candidatas inspeccionadas", state.results.candidatePairs],
      ["Parejas puntuadas", state.results.evaluatedPairs],
      ["Combinaciones agrupadas evaluadas", state.results.evaluatedCombinations],
      ["Movimientos base acotados", state.results.limitedGroupAnchors],
      ["Límite de parejas alcanzado", yesNo(state.results.pairLimitReached)],
      ["Límite de agrupaciones alcanzado", yesNo(state.results.combinationLimitReached)],
      ["Pasadas flexibles sobre pendientes", retryPasses.length],
      ["Agrupaciones rechazadas recordadas", state.review.rejectedSignatures.size],
      ["Último reanálisis flexible", lastRetry ? formatDateTime(lastRetry.at) : "No realizado"],
      ["Parámetros del último reanálisis", lastRetry ? `${lastRetry.dateTolerance} días · tolerancia ${formatMoney(lastRetry.amountAbsTolerance)} + ${formatDecimal(lastRetry.amountPercentTolerance || 0, 2)}% · descripción ${formatDecimal(lastRetry.minimumDescriptionSimilarity || 0, 0)}% · sin fecha ${yesNo(lastRetry.ignoreDates)} · ${lastRetry.found} coincidencias` : ""]
    ];
    const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
    styleSummarySheet(summarySheet);
    appendSheet(workbook, summarySheet, "Resumen");

    const reconciliationHeaders = ["ID de conciliación", "Tipo", "Filas del sistema", "Filas de caja o banco", "Fechas del sistema", "Fechas de caja o banco", "Descripciones del sistema", "Descripciones de caja o banco", "Débito del sistema", "Crédito del sistema", "Débito de caja o banco", "Crédito de caja o banco", "Diferencia", "Puntaje", "Estado", "Criterio", "Observaciones", "Fecha y hora de procesamiento"];
    const confirmed = state.results.reconciliations.filter(item => item.status === "confirmed");
    const possible = state.results.reconciliations.filter(item => item.status === "possible");
    const reconciliationSheet = createStyledDataSheet(reconciliationHeaders, confirmed.map(reconciliationExportRow), { fill: "E9F5ED", numericColumns: [8, 9, 10, 11, 12], integerColumns: [13] });
    appendSheet(workbook, reconciliationSheet, "Conciliaciones");
    const possibleSheet = createStyledDataSheet(reconciliationHeaders, possible.map(reconciliationExportRow), { fill: "FFF7DC", numericColumns: [8, 9, 10, 11, 12], integerColumns: [13] });
    appendSheet(workbook, possibleSheet, "Posibles conciliaciones");

    const pendingHeaders = ["Fila original", "Fecha", "Descripción", "Débito", "Crédito", "Tipo", "Estado original"];
    const pendingSystemSheet = createStyledDataSheet(pendingHeaders, summary.pendingSystem.map(pendingExportRow), { fill: "FBEEEE", numericColumns: [3, 4] });
    appendSheet(workbook, pendingSystemSheet, "Pendientes del sistema");
    const pendingBankSheet = createStyledDataSheet(pendingHeaders, summary.pendingBank.map(pendingExportRow), { fill: "FBEEEE", numericColumns: [3, 4] });
    appendSheet(workbook, pendingBankSheet, "Pendientes de caja o banco");

    if ((state.transferLog || []).length) {
      const transferHeaders = ["Dirección", "Fecha y hora", "Cuenta origen", "Cuenta destino", "Origen del movimiento", "Fila original", "Fecha del movimiento", "Descripción", "Débito", "Crédito", "ID original", "ID en destino"];
      const transferRows = state.transferLog.map(transferExportRow);
      appendSheet(workbook, createStyledDataSheet(transferHeaders, transferRows, { fill: "EDF5FA", numericColumns: [8, 9] }), "Movimientos transferidos");
    }

    if (summary.excludedCount) {
      const excludedHeaders = ["Origen", ...pendingHeaders, "Motivo", "Período aplicado"];
      const periodLabel = state.review.periodFilter?.from || state.review.periodFilter?.to
        ? `${state.review.periodFilter.from ? `Desde ${formatDateInput(state.review.periodFilter.from)}` : "Sin fecha inicial"} · ${state.review.periodFilter.to ? `Hasta ${formatDateInput(state.review.periodFilter.to)}` : "Sin fecha final"}`
        : "";
      const excludedRows = [
        ...summary.excludedSystem.map(item => excludedPeriodExportRow("Sistema contable", item, periodLabel)),
        ...summary.excludedBank.map(item => excludedPeriodExportRow("Caja o banco", item, periodLabel))
      ];
      appendSheet(workbook, createStyledDataSheet(excludedHeaders, excludedRows, { fill: "F3F0EB", numericColumns: [4, 5] }), "Excluidos por período");
    }

    appendSheet(workbook, createOriginalDataSheet(state.sources.system), "Datos originales del sistema");
    appendSheet(workbook, createOriginalDataSheet(state.sources.bank), "Datos originales caja o banco");

    const allErrors = [...state.sources.system.invalidRows, ...state.sources.bank.invalidRows];
    if (allErrors.length) {
      const errorRows = allErrors.map(item => [item.source, item.sheet, item.row, item.errors, item.values.map(displayOriginalValue).join(" | ")]);
      appendSheet(workbook, createStyledDataSheet(["Origen", "Hoja", "Fila", "Error", "Datos originales"], errorRows, { fill: "FBEEEE", numericColumns: [2] }), "Errores de importación");
    }
    const snapshot = createStateSnapshot();
    appendSheet(workbook, createApplicationStateSheet(snapshot), STATE_SHEET_NAME);
    hideWorkbookSheet(workbook, STATE_SHEET_NAME);
    const output = XLSX.write(workbook, { compression: true, bookType: "xlsx", type: "array", cellStyles: true });
    const exportFileNameInput = document.getElementById("exportFileName");
    const exportFileName = normalizeExportFileName(exportFileNameInput.value);
    exportFileNameInput.value = exportFileName.replace(/\.xlsx$/i, "");
    downloadBlob(
      new Blob([output], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
      exportFileName
    );
    showToast("Excel generado", "El libro de conciliación se descargó correctamente.", "success");
    scheduleStatePersistence();
  } catch (error) {
    console.error(error);
    showToast("No se pudo generar el Excel", error.message || "Revise los datos y vuelva a intentarlo.", "error", 9000);
  }
}

function reconciliationExportRow(item) {
  const systemTotals = debitCreditTotals(item.systemMovements);
  const bankTotals = debitCreditTotals(item.bankMovements);
  return [
    item.id,
    typeLabel(item),
    item.systemMovements.map(movement => movement.row).join(", "),
    item.bankMovements.map(movement => movement.row).join(", "),
    item.systemMovements.map(movement => formatDate(movement.date)).join(" | "),
    item.bankMovements.map(movement => formatDate(movement.date)).join(" | "),
    item.systemMovements.map(movement => movement.description).join(" | "),
    item.bankMovements.map(movement => movement.description).join(" | "),
    systemTotals.debit,
    systemTotals.credit,
    bankTotals.debit,
    bankTotals.credit,
    item.difference,
    item.score,
    item.status === "confirmed" ? item.manuallyApproved ? "Conciliado · posible aprobada manualmente" : "Conciliado" : "Posible",
    item.criterion,
    item.observation,
    formatDateTime(state.results.processingAt || item.createdAt)
  ];
}

function pendingExportRow(item) {
  const amounts = movementDebitCredit(item);
  return [item.row, formatDate(item.date), item.description, amounts.debit, amounts.credit, movementTypeLabel(item.type), item.status];
}

function transferExportRow(item) {
  const movement = {
    amount: Number(item.amount) || 0,
    debitAmount: Number(item.amount) >= 0 ? Math.abs(Number(item.amount) || 0) : 0,
    creditAmount: Number(item.amount) < 0 ? Math.abs(Number(item.amount) || 0) : 0,
    type: Number(item.amount) >= 0 ? "debit" : "credit"
  };
  const totals = movementDebitCredit(movement);
  return [
    item.direction === "in" ? "Recibido" : "Enviado",
    formatDateTime(reviveStoredDate(item.at) || new Date()),
    item.fromAccount || "",
    item.toAccount || "",
    item.sourceKey === "system" ? "Sistema contable" : "Caja o banco",
    item.row || "",
    item.dateKey ? formatDateInput(item.dateKey) : formatDate(reviveStoredDate(item.date)),
    item.description || "",
    totals.debit,
    totals.credit,
    item.movementOriginId || "",
    item.movementDestinationId || ""
  ];
}

function excludedPeriodExportRow(sourceLabel, item, periodLabel) {
  return [sourceLabel, ...pendingExportRow(item), "Fuera del período activo; no participa en búsquedas ni pendientes", periodLabel];
}

function movementDebitCredit(movement) {
  const hasSeparatedAmounts = Number.isFinite(Number(movement.debitAmount)) && Number.isFinite(Number(movement.creditAmount));
  if (hasSeparatedAmounts) return { debit: Math.abs(Number(movement.debitAmount) || 0), credit: Math.abs(Number(movement.creditAmount) || 0) };
  if (movement.type === "debit") return { debit: Math.abs(Number(movement.amount) || 0), credit: 0 };
  if (movement.type === "credit") return { debit: 0, credit: Math.abs(Number(movement.amount) || 0) };
  return Number(movement.amount) >= 0
    ? { debit: Math.abs(Number(movement.amount) || 0), credit: 0 }
    : { debit: 0, credit: Math.abs(Number(movement.amount) || 0) };
}

function debitCreditTotals(movements) {
  return movements.reduce((totals, movement) => {
    const amounts = movementDebitCredit(movement);
    totals.debit = roundMoney(totals.debit + amounts.debit);
    totals.credit = roundMoney(totals.credit + amounts.credit);
    return totals;
  }, { debit: 0, credit: 0 });
}

function createOriginalDataSheet(source) {
  const dateColumn = source.mapping.date === "" ? -1 : Number(source.mapping.date);
  const rows = source.matrix.map((row, rowIndex) => row.map((value, columnIndex) => {
    if (rowIndex >= source.dataStartRow - 1 && columnIndex === dateColumn) {
      const date = parseDateValue(value);
      if (date) return formatDate(date);
    }
    return value instanceof Date ? formatDate(value) : value;
  }));
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  applyTableSheetStyle(worksheet, { fill: "FFFFFF", numericColumns: [] });
  return worksheet;
}

function createStyledDataSheet(headers, rows, options = {}) {
  const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  applyTableSheetStyle(worksheet, options);
  return worksheet;
}

function applyTableSheetStyle(worksheet, { fill = "FFFFFF", numericColumns = [], integerColumns = [] } = {}) {
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
      if (integerColumns.includes(column)) cell.z = "0";
      else if (numericColumns.includes(column)) cell.z = "#,##0.00;[Red]-#,##0.00";
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
    const section = row === 0 || row === 19;
    labelCell.s = section
      ? { fill: { fgColor: { rgb: "164E55" } }, font: { bold: true, color: { rgb: "FFFFFF" }, sz: row === 0 ? 14 : 11 }, alignment: { vertical: "center" } }
      : { fill: { fgColor: { rgb: row % 2 ? "F4F7F7" : "FFFFFF" } }, font: { bold: true, color: { rgb: "405153" } }, border: { bottom: { style: "hair", color: { rgb: "DDE3E4" } } } };
    if (valueCell) valueCell.s = section
      ? { fill: { fgColor: { rgb: "164E55" } } }
      : { fill: { fgColor: { rgb: row % 2 ? "F4F7F7" : "FFFFFF" } }, border: { bottom: { style: "hair", color: { rgb: "DDE3E4" } } } };
  }
  if (worksheet.B14) worksheet.B14.z = "#,##0.00;[Red]-#,##0.00";
  if (worksheet.B15) worksheet.B15.z = "#,##0.00;[Red]-#,##0.00";
  if (worksheet.B16) worksheet.B16.z = "#,##0.00;[Red]-#,##0.00";
  if (worksheet.B17) worksheet.B17.z = "0.0%";
  if (worksheet.B18) worksheet.B18.z = "0.0%";
  if (worksheet.B22) worksheet.B22.z = "#,##0.00";
  if (worksheet.B23) worksheet.B23.z = "0.00%";
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
  while (dom.toastRegion.children.length >= 2) dom.toastRegion.firstElementChild.remove();
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

function normalizeExportFileName(value) {
  let name = String(value || "").trim().replace(/\.xlsx$/i, "");
  name = name.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "-").replace(/[. ]+$/g, "").trim().slice(0, 120);
  return `${name || `conciliacion_${toFileTimestamp()}`}.xlsx`;
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

function formatDateInput(value) {
  const match = String(value || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return match ? `${match[3]}/${match[2]}/${match[1]}` : String(value || "");
}

function parseDisplayDateInput(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  const compact = text.match(/^\d{6}$|^\d{8}$/)
    ? text
    : text.match(/^\d{2}[\/\-.]\d{2}[\/\-.](?:\d{2}|\d{4})$/)
      ? text.replace(/\D/g, "")
      : "";
  if (!compact) return null;
  const day = Number(compact.slice(0, 2));
  const month = Number(compact.slice(2, 4));
  let year = Number(compact.slice(4));
  if (compact.length === 6) year += year < 50 ? 2000 : 1900;
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null;
  return `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
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
  APP_BUILD,
  DEFAULT_CONFIG,
  parseAmount,
  parseDateValue,
  normalizeDescription,
  descriptionSimilarity,
  calculateReconciliation,
  calculateInternalReconciliation,
  internalSelectionCanReconcile,
  maybeAutoDetectSplitConvention,
  enumerateGroupCandidates,
  reconciliationEngine,
  retryPendingReconciliation,
  cancelProcessing,
  loadSourceFile,
  loadPreviousReconciliationFile,
  renderReview,
  normalizeSourceRow,
  movementDebitCredit,
  debitCreditTotals,
  reconciliationExportRow,
  pendingExportRow,
  createStateSnapshot,
  restoreStateSnapshot,
  createApplicationStateSheet,
  extractStateSnapshotFromWorkbook,
  normalizeTransferSnapshot,
  snapshotPendingMovements,
  moveSnapshotMovement,
  persistSnapshot,
  readPersistedSnapshot,
  readSavedWorkspaceSnapshots,
  clearPersistedState,
  exportWorkbook,
  getState: () => state
});
