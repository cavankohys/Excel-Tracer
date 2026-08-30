
const COLORS = {
  hardcode: "#0000FF",
  offsheet: "#008000",
  formula: "#000000"
};

let currentAddress = "";
let traceSourceAddress = "";
let lastTrace = [];
let lastTraceLabel = "";
let selectionHandlerRegistered = false;
let workbookEventHandlersRegistered = false;
let working = false;
let internalWorkbookMutationDepth = 0;
let calculationRefreshTimer = null;
let calculationInfo = { mode: "", state: "", iterative: null };
let logicState = createEmptyLogicState();
let areaState = createEmptyAreaState();
let lastAuditFindings = [];
let lastAuditMeta = null;
let activeAuditSeverity = "all";
let activeAuditCategory = "all";
let analysisMode = "dependencies";

const EXPLORER_LIMITS = {
  maxDepth: 20,
  maxNodes: 500,
  maxSourceCells: 2000
};

let explorer = createEmptyExplorerState();

function createEmptyExplorerState() {
  return {
    active: false,
    relation: "",
    sourceAddress: "",
    rootId: "",
    activeId: "",
    nodes: new Map(),
    history: [],
    sheetVisibility: new Map(),
    hiddenResultCount: 0,
    navigationToken: 0,
    sourceCellCount: 1,
    boundaryNotice: "",
    calculationNotice: "",
    stale: false,
    changeGeneration: 0
  };
}

function createEmptyLogicState() {
  return { active: false, sourceAddress: "", sourceSheet: "", formula: "", formulaR1C1: "", nodes: new Map(), rootId: "", activeId: "" };
}

function createEmptyAreaState() {
  return { active: false, sourceAddress: "", sourceSheet: "", items: [], activeIndex: 0, exactOutputMapping: true };
}


const $ = (id) => document.getElementById(id);

const SHORTCUT_ACTIONS = [
  { id: "OpenModelTracer", label: "Explorer", hint: "Open pane", safeKey: "M" },
  { id: "ExplorePrecedents", label: "Precedents", hint: "Trace upstream", safeKey: "P" },
  { id: "ExploreDependents", label: "Dependents", hint: "Trace downstream", safeKey: "D" },
  { id: "ExplorerBack", label: "Back", hint: "Navigation history", safeKey: "Left" },
  { id: "RecalculateWorkbook", label: "Recalculate", hint: "Refresh workbook values", safeKey: "R" },
  { id: "OpenFormulaLogic", label: "Formula Logic", hint: "Break down active formula", safeKey: "L" },
  { id: "ReviewCalculationArea", label: "Calculation Area", hint: "Review selected block", safeKey: "A" },
  { id: "ColorHardcodeBlue", label: "Blue hardcode", hint: "Model format", safeKey: "B" },
  { id: "ColorOffsheetGreen", label: "Green off-sheet", hint: "Model format", safeKey: "G" },
  { id: "ColorFormulaBlack", label: "Black formula", hint: "Model format", safeKey: "K" }
];
let shortcutSettingsSupported = false;
let shortcutCurrentValues = {};


Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("Excel only", "error");
    return;
  }

  configureDeploymentUi();
  registerShortcutActions();
  wireButtons();
  restoreSettings();
  configureApiSupport();
  await refreshShortcutSummary();
  await refreshSelection();
  await refreshCalculationStatus();
  await registerSelectionHandler();
  await registerWorkbookEventHandlers();
  setStatus("", "");
});


function configureDeploymentUi() {
  try {
    const mode = new URLSearchParams(window.location.search).get("shortcuts");
    const box = $("shortcut-box");
    const headerButton = $("shortcut-settings-btn");
    if (mode === "0") {
      if (box) box.classList.add("hidden");
      if (headerButton) headerButton.classList.add("hidden");
    }
  } catch {}
}

function wireButtons() {
  $("tab-trace").addEventListener("click", () => switchTab("trace"));
  $("tab-format").addEventListener("click", () => switchTab("format"));
  $("tab-audit").addEventListener("click", () => switchTab("audit"));
  $("shortcut-settings-btn").addEventListener("click", openShortcutSettings);
  $("shortcut-configure-inline").addEventListener("click", openShortcutSettings);
  $("shortcut-close").addEventListener("click", closeShortcutSettings);
  $("shortcut-cancel").addEventListener("click", closeShortcutSettings);
  $("shortcut-profile-safe").addEventListener("click", () => loadShortcutProfile("safe"));
  $("shortcut-profile-banking").addEventListener("click", () => loadShortcutProfile("banking"));
  $("shortcut-reset").addEventListener("click", resetShortcutDefaults);
  $("shortcut-check").addEventListener("click", () => checkShortcutConflicts(false));
  $("shortcut-apply").addEventListener("click", applyShortcutSettings);
  $("shortcut-modal").addEventListener("click", (event) => {
    if (event.target === $("shortcut-modal")) closeShortcutSettings();
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !$("shortcut-modal").classList.contains("hidden")) {
      event.preventDefault();
      closeShortcutSettings();
    }
  });

  $("refresh-btn").addEventListener("click", async () => {
    resetExplorer();
    await refreshSelection(true);
  });
  $("explore-precedents").addEventListener("click", () => startExplorer("precedents"));
  $("explore-dependents").addEventListener("click", () => startExplorer("dependents"));
  $("explorer-source-btn").addEventListener("click", goToExplorerSource);
  $("explorer-back-btn").addEventListener("click", explorerGoBack);
  $("explorer-refresh-trace-btn").addEventListener("click", () => refreshExplorerTrace(true));
  $("copy-btn").addEventListener("click", copyExplorerTree);
  $("recalculate-btn").addEventListener("click", recalculateWorkbook);
  $("formula-logic-btn").addEventListener("click", openFormulaLogic);
  $("logic-close-btn").addEventListener("click", closeFormulaLogic);
  $("logic-source-btn").addEventListener("click", goToLogicSource);
  $("logic-tree").addEventListener("keydown", handleLogicKeydown);
  $("logic-tree").addEventListener("click", handleLogicClick);
  $("calculation-area-btn").addEventListener("click", reviewCalculationArea);
  $("area-close-btn").addEventListener("click", closeCalculationArea);
  $("area-source-btn").addEventListener("click", goToAreaSource);
  $("area-results").addEventListener("keydown", handleAreaKeydown);
  $("area-results").addEventListener("click", handleAreaClick);
  $("explorer-tree").addEventListener("keydown", handleExplorerKeydown);
  $("explorer-tree").addEventListener("click", handleExplorerClick);

  $("format-selection").addEventListener("click", () => formatScope("selection"));
  $("format-sheet").addEventListener("click", () => formatScope("sheet"));
  $("format-workbook").addEventListener("click", () => formatScope("workbook"));

  $("manual-blue").addEventListener("click", () => setSelectedFontColor(COLORS.hardcode, "Blue"));
  $("manual-green").addEventListener("click", () => setSelectedFontColor(COLORS.offsheet, "Green"));
  $("manual-black").addEventListener("click", () => setSelectedFontColor(COLORS.formula, "Black"));

  $("audit-workbook").addEventListener("click", () => auditScope("workbook"));
  $("audit-sheet").addEventListener("click", () => auditScope("sheet"));
  $("audit-selection").addEventListener("click", () => auditScope("selection"));
  $("copy-audit").addEventListener("click", copyAuditFindings);
  $("export-audit").addEventListener("click", exportAuditReport);
  $("audit-category-filter").addEventListener("change", (event) => {
    activeAuditCategory = event.target.value;
    renderAuditFindings();
  });
  document.querySelectorAll("#severity-filter .filter-tab").forEach((button) => {
    button.addEventListener("click", () => {
      activeAuditSeverity = button.dataset.severity || "all";
      document.querySelectorAll("#severity-filter .filter-tab").forEach((x) => x.classList.toggle("active", x === button));
      renderAuditFindings();
    });
  });
  $("toggle-checks").addEventListener("click", () => {
    const detail = $("audit-checks-detail");
    const hidden = detail.classList.toggle("hidden");
    $("toggle-checks").textContent = hidden ? "Show" : "Hide";
  });

  $("reset-local-formulas").addEventListener("change", saveSettings);
  $("include-text-hardcodes").addEventListener("change", saveSettings);
}

function switchTab(which) {
  const trace = which === "trace";
  const format = which === "format";
  const audit = which === "audit";
  $("tab-trace").classList.toggle("active", trace);
  $("tab-format").classList.toggle("active", format);
  $("tab-audit").classList.toggle("active", audit);
  $("panel-trace").classList.toggle("hidden", !trace);
  $("panel-format").classList.toggle("hidden", !format);
  $("panel-audit").classList.toggle("hidden", !audit);
}

function showAnalysisMode(mode = "dependencies") {
  analysisMode = ["dependencies", "logic", "area"].includes(mode) ? mode : "dependencies";
  $("explorer-card").classList.toggle("hidden", analysisMode !== "dependencies");
  $("formula-logic-card").classList.toggle("hidden", analysisMode !== "logic");
  $("calculation-area-card").classList.toggle("hidden", analysisMode !== "area");
}

function saveSettings() {
  try {
    localStorage.setItem("modelTracer.resetLocal", $("reset-local-formulas").checked ? "1" : "0");
    localStorage.setItem("modelTracer.includeText", $("include-text-hardcodes").checked ? "1" : "0");
  } catch {}
}

function restoreSettings() {
  try {
    const reset = localStorage.getItem("modelTracer.resetLocal");
    const text = localStorage.getItem("modelTracer.includeText");
    if (reset !== null) $("reset-local-formulas").checked = reset === "1";
    if (text !== null) $("include-text-hardcodes").checked = text === "1";
  } catch {}
}

function configureApiSupport() {
  const supports = {
    precedents: Office.context.requirements.isSetSupported("ExcelApi", "1.12"),
    dependents: Office.context.requirements.isSetSupported("ExcelApi", "1.13"),
    spill: Office.context.requirements.isSetSupported("ExcelApi", "1.12"),
    calculationState: Office.context.requirements.isSetSupported("ExcelApi", "1.9"),
    workbookEvents: Office.context.requirements.isSetSupported("ExcelApi", "1.7")
  };

  $("explore-precedents").disabled = !supports.precedents;
  $("explore-dependents").disabled = !supports.dependents;

  const missing = [];
  if (!supports.precedents) missing.push("precedent exploration (ExcelApi 1.12)");
  if (!supports.dependents) missing.push("dependent exploration (ExcelApi 1.13)");

  if (missing.length) {
    const note = $("api-note");
    note.textContent = "Unsupported in this Excel version: " + missing.join(", ") + ".";
    note.classList.remove("hidden");
  }

  shortcutSettingsSupported = Boolean(
    Office.actions &&
    Office.actions.getShortcuts &&
    Office.actions.replaceShortcuts &&
    Office.actions.areShortcutsInUse &&
    Office.context.requirements.isSetSupported("KeyboardShortcuts", "1.1") &&
    Office.context.requirements.isSetSupported("SharedRuntime", "1.1")
  );
}

async function registerSelectionHandler() {
  if (selectionHandlerRegistered) return;
  try {
    await Excel.run(async (context) => {
      context.workbook.onSelectionChanged.add(async () => {
        if (!working && !explorer.active) await refreshSelection(false);
      });
      await context.sync();
      selectionHandlerRegistered = true;
    });
  } catch {
    // Refresh button remains available if selection events are unsupported.
  }
}


async function registerWorkbookEventHandlers() {
  if (workbookEventHandlersRegistered) return;
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) return;

  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.onChanged.add(handleWorksheetChanged);
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.13")) {
        sheets.onFormulaChanged.add(handleWorksheetFormulaChanged);
      }
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.8")) {
        sheets.onCalculated.add(handleWorksheetCalculated);
      }
      await context.sync();
      workbookEventHandlersRegistered = true;
    });
  } catch {
    // The add-in remains functional without live invalidation on older hosts.
  }
}

function handleWorksheetChanged(event) {
  if (internalWorkbookMutationDepth > 0) return;

  // ExcelApi 1.13 provides a formula-specific event. On newer hosts, data-only changes
  // invalidate values/audit results but do not force dependency-branch rebuilds.
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.13")) markExplorerTraceStale();
  markAuditStale();
  refreshActiveExplorerPreviewSoon();
  scheduleCalculationRefresh();
}

function handleWorksheetFormulaChanged(event) {
  if (internalWorkbookMutationDepth > 0) return;
  markExplorerTraceStale();
  markAuditStale();
  scheduleCalculationRefresh();
}

function markExplorerTraceStale() {
  if (!explorer.active) return;
  explorer.stale = true;
  explorer.changeGeneration += 1;
  refreshExplorerWarning();
  const refreshButton = $("explorer-refresh-trace-btn");
  if (refreshButton) refreshButton.classList.remove("hidden");
}

function markAuditStale() {
  if (!lastAuditMeta) return;
  lastAuditMeta.stale = true;
  const badge = $("audit-status-badge");
  if (badge) {
    badge.textContent = "Workbook changed since audit";
    badge.className = "audit-status review";
  }
}

function refreshActiveExplorerPreviewSoon() {
  if (!explorer.active || !explorer.activeId) return;
  const node = explorer.nodes.get(explorer.activeId);
  if (!node || node.nodeType !== "cell") return;
  node.loadedPreview = false;
  ensureExplorerPreview(node, { force: true }).catch(() => {});
}

function handleWorksheetCalculated() {
  scheduleCalculationRefresh();
  refreshActiveExplorerPreviewSoon();
}

function scheduleCalculationRefresh() {
  if (calculationRefreshTimer) clearTimeout(calculationRefreshTimer);
  calculationRefreshTimer = setTimeout(() => {
    calculationRefreshTimer = null;
    refreshCalculationStatus(true).catch(() => {});
  }, 180);
}

async function getCalculationSnapshot(context = null) {
  const loadSnapshot = async (ctx) => {
    const application = ctx.workbook.application;
    application.load("calculationMode");
    const supports19 = Office.context.requirements.isSetSupported("ExcelApi", "1.9");
    if (supports19) {
      application.load("calculationState");
      application.iterativeCalculation.load("enabled,maxIteration,maxChange");
    }
    await ctx.sync();
    return {
      mode: String(application.calculationMode || ""),
      state: supports19 ? String(application.calculationState || "") : "",
      iterative: supports19 ? {
        enabled: Boolean(application.iterativeCalculation.enabled),
        maxIteration: Number(application.iterativeCalculation.maxIteration || 0),
        maxChange: Number(application.iterativeCalculation.maxChange || 0)
      } : null
    };
  };

  if (context) return loadSnapshot(context);
  return Excel.run(async (ctx) => loadSnapshot(ctx));
}

function calculationNotice(info) {
  if (!info) return "";
  const bits = [];
  const mode = String(info.mode || "").toLowerCase().replace(/[^a-z]/g, "");
  const state = String(info.state || "").toLowerCase();
  if (mode === "manual") bits.push("Manual calculation");
  if (mode.includes("automaticexcept") || mode.includes("excepttables")) bits.push("Automatic except data tables");
  if (state.includes("pending")) bits.push("values pending recalculation");
  if (state.includes("calculating")) bits.push("calculation in progress");
  return bits.join(" · ");
}

async function refreshCalculationStatus(silent = false) {
  try {
    calculationInfo = await getCalculationSnapshot();
    const notice = calculationNotice(calculationInfo);
    const box = $("calc-status");
    const text = $("calc-status-text");
    if (!box || !text) return calculationInfo;

    if (notice) {
      text.textContent = `Values may be stale. ${notice}`;
      box.classList.remove("hidden");
    } else {
      text.textContent = "";
      box.classList.add("hidden");
    }

    if (explorer.active) {
      explorer.calculationNotice = notice ? `Values may be stale. ${notice}.` : "";
      refreshExplorerWarning();
    }
    return calculationInfo;
  } catch (error) {
    if (!silent) setStatus("Calculation status unavailable", "error");
    return calculationInfo;
  }
}

async function recalculateWorkbook(event) {
  setWorking(true, "Recalculating");
  internalWorkbookMutationDepth += 1;
  try {
    await Excel.run(async (context) => {
      context.workbook.application.calculate("Recalculate");
      await context.sync();
    });
    await refreshCalculationStatus(true);
    if (explorer.active && explorer.activeId) {
      const node = explorer.nodes.get(explorer.activeId);
      if (node) await ensureExplorerPreview(node, { force: true });
    }
  } catch (error) {
    showExplorerWarning(`Could not recalculate workbook: ${normalizeError(error)}`);
  } finally {
    internalWorkbookMutationDepth = Math.max(0, internalWorkbookMutationDepth - 1);
    setWorking(false);
    try { event?.completed?.(); } catch {}
  }
}

async function refreshSelection(clearResults = true) {
  try {
    const snapshot = await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      const first = range.getCell(0, 0);
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      range.load(["address", "rowCount", "columnCount"]);
      first.load(["formulas", "values"]);
      sheet.load("name");
      await context.sync();

      return {
        address: qualifyAddress(sheet.name, localAddress(range.address)),
        cellCount: Math.max(1, Number(range.rowCount || 1) * Number(range.columnCount || 1)),
        formula: first.formulas?.[0]?.[0],
        value: first.values?.[0]?.[0]
      };
    });

    updateSourceCard(snapshot.address, snapshot.formula, snapshot.value, snapshot.cellCount);

    if (clearResults && !explorer.active) {
      renderExplorerEmpty("No active trace");
    }
  } catch (error) {
    showError(normalizeError(error));
  }
}


/* ----------------------------
   Formula Explorer
----------------------------- */


function flattenWorkbookAddresses(addresses) {
  const out = [];
  const seen = new Set();

  for (const raw of addresses) {
    for (const token of splitExcelAddressList(raw)) {
      const trimmed = token.trim();
      if (!trimmed) continue;
      const key = trimmed.toUpperCase();
      if (!seen.has(key)) {
        seen.add(key);
        out.push(trimmed);
      }
    }
  }
  return out;
}

function splitExcelAddressList(text) {
  const parts = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];

    if (ch === "'") {
      if (inQuotes && text[i + 1] === "'") {
        current += "''";
        i++;
        continue;
      }
      inQuotes = !inQuotes;
      current += ch;
      continue;
    }

    if (ch === "," && !inQuotes) {
      parts.push(current);
      current = "";
    } else {
      current += ch;
    }
  }

  if (current) parts.push(current);
  return parts;
}



async function startExplorer(relation) {
  if (!["precedents", "dependents"].includes(relation)) return;
  showAnalysisMode("dependencies");

  setWorking(true, relation === "precedents" ? "Loading precedents" : "Loading dependents");
  try {
    const snapshot = await Excel.run(async (context) => {
      const workbook = context.workbook;
      const range = workbook.getSelectedRange();
      const first = range.getCell(0, 0);
      const sheet = workbook.worksheets.getActiveWorksheet();
      const sheets = workbook.worksheets;

      range.load(["address", "rowCount", "columnCount"]);
      first.load(["formulas", "values"]);
      sheet.load("name");
      sheets.load("items/name,items/visibility");
      const spillParent = Office.context.requirements.isSetSupported("ExcelApi", "1.12")
        ? first.getSpillParentOrNullObject()
        : null;
      const spillRange = Office.context.requirements.isSetSupported("ExcelApi", "1.12")
        ? first.getSpillingToRangeOrNullObject()
        : null;
      if (spillParent) spillParent.load(["isNullObject", "address"]);
      if (spillRange) spillRange.load(["isNullObject", "address"]);
      const calc = await getCalculationSnapshot(context);

      const cellCount = Math.max(1, Number(range.rowCount || 1) * Number(range.columnCount || 1));
      if (cellCount > EXPLORER_LIMITS.maxSourceCells) {
        throw new Error(`Selection contains ${cellCount.toLocaleString()} cells. Model Tracer limits multi-cell tracing to ${EXPLORER_LIMITS.maxSourceCells.toLocaleString()} source cells to protect Excel performance. Narrow the selection and retry.`);
      }

      return {
        sourceAddress: qualifyAddress(sheet.name, localAddress(range.address)),
        sourceCellCount: cellCount,
        sourceFormula: first.formulas?.[0]?.[0],
        sourceValue: first.values?.[0]?.[0],
        spillParent: spillParent && !spillParent.isNullObject ? qualifyAddress(sheet.name, localAddress(spillParent.address)) : "",
        spillRange: spillRange && !spillRange.isNullObject ? qualifyAddress(sheet.name, localAddress(spillRange.address)) : "",
        calculation: calc,
        sheets: sheets.items.map((item) => ({
          name: item.name,
          visibility: String(item.visibility || "Visible")
        }))
      };
    });

    explorer = createEmptyExplorerState();
    explorer.active = true;
    explorer.relation = relation;
    explorer.sourceAddress = snapshot.sourceAddress;
    explorer.sourceCellCount = snapshot.sourceCellCount;
    explorer.boundaryNotice = relation === "dependents"
      ? "Dependent tracing is limited to the current workbook. Excel cannot discover formulas in other workbook files that may depend on this source."
      : "";
    calculationInfo = snapshot.calculation || calculationInfo;
    explorer.calculationNotice = calculationNotice(calculationInfo)
      ? `Values may be stale. ${calculationNotice(calculationInfo)}.`
      : "";
    explorer.sheetVisibility = new Map(snapshot.sheets.map((item) => [item.name.toLowerCase(), item.visibility]));

    const isSelectionRoot = snapshot.sourceCellCount > 1;
    const root = createExplorerNode({
      address: snapshot.sourceAddress,
      parentId: "",
      depth: 0,
      relationClass: isSelectionRoot ? "selection" : "source",
      formula: isSelectionRoot ? null : snapshot.sourceFormula,
      value: isSelectionRoot ? null : snapshot.sourceValue,
      loadedPreview: true,
      nodeType: isSelectionRoot ? "selection" : "cell",
      cellCount: snapshot.sourceCellCount,
      spillParent: isSelectionRoot ? "" : snapshot.spillParent,
      spillRange: isSelectionRoot ? "" : snapshot.spillRange
    });

    explorer.rootId = root.id;
    explorer.activeId = root.id;
    explorer.nodes.set(root.id, root);

    updateSourceCard(snapshot.sourceAddress, snapshot.sourceFormula, snapshot.sourceValue, snapshot.sourceCellCount);

    await expandExplorerNode(root.id, { keepSelection: true });
    root.expanded = true;
    renderExplorer();
    updateExplorerPreview(root);

    const tree = $("explorer-tree");
    tree.focus({ preventScroll: true });
  } catch (error) {
    if (isInvalidSelection(error)) {
      renderExplorerError("Select one contiguous cell or rectangular range in Excel, then run Precedents or Dependents again.");
    } else {
      renderExplorerError(normalizeError(error));
    }
  } finally {
    setWorking(false);
  }
}

function createExplorerNode({
  address,
  parentId,
  depth,
  relationClass = "",
  formula = null,
  value = null,
  loadedPreview = false,
  nodeType = "cell",
  cellCount = 1,
  externalWorkbook = "",
  spillParent = "",
  spillRange = ""
}) {
  const id = `node-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
  const parent = parentId ? explorer.nodes.get(parentId) : null;

  if (nodeType === "external") {
    return {
      id,
      address: externalWorkbook || address,
      parentId,
      depth,
      relationClass: "external",
      hidden: false,
      cycle: false,
      children: [],
      expanded: false,
      loading: false,
      formula: null,
      value: null,
      loadedPreview: true,
      nodeType,
      cellCount: 0,
      externalWorkbook: externalWorkbook || address,
      spillParent: "",
      spillRange: "",
      previewGeneration: explorer.changeGeneration,
      childrenGeneration: explorer.changeGeneration
    };
  }

  const parsed = parseQualifiedAddress(address);
  const sheetName = parsed.sheetName || (parent ? parseQualifiedAddress(parent.address).sheetName : "");

  let relation = relationClass;
  if (!relation && parent) {
    const parentSheet = parseQualifiedAddress(parent.address).sheetName;
    relation = sheetName && parentSheet && sheetName.toLowerCase() !== parentSheet.toLowerCase()
      ? "off-sheet"
      : "same-sheet";
  }

  const hidden = isExplorerSheetHidden(sheetName);
  const cycle = nodeType === "cell" && isAddressInAncestorChain(parentId, address);

  return {
    id,
    address,
    parentId,
    depth,
    relationClass: hidden ? "hidden" : (cycle ? "cycle" : relation),
    hidden,
    cycle,
    children: null,
    expanded: false,
    loading: false,
    formula,
    value,
    loadedPreview,
    nodeType,
    cellCount,
    externalWorkbook,
    spillParent,
    spillRange,
    previewGeneration: explorer.changeGeneration,
    childrenGeneration: -1
  };
}

function isAddressInAncestorChain(parentId, address) {
  const target = normalizeAddressKey(address);
  let cursor = parentId;
  while (cursor) {
    const node = explorer.nodes.get(cursor);
    if (!node) break;
    if (normalizeAddressKey(node.address) === target) return true;
    cursor = node.parentId;
  }
  return false;
}

function normalizeAddressKey(address) {
  return String(address || "").replace(/\$/g, "").toUpperCase();
}

function isExplorerSheetHidden(sheetName) {
  if (!sheetName) return false;
  const visibility = explorer.sheetVisibility.get(sheetName.toLowerCase());
  return visibility && visibility.toLowerCase() !== "visible";
}

async function expandExplorerNode(nodeId, options = {}) {
  const node = explorer.nodes.get(nodeId);
  if (!node || node.cycle || node.nodeType === "external" || node.depth >= EXPLORER_LIMITS.maxDepth) return;

  if (Array.isArray(node.children) && node.childrenGeneration === explorer.changeGeneration) {
    node.expanded = true;
    renderExplorer();
    return;
  }

  if (Array.isArray(node.children) && node.childrenGeneration !== explorer.changeGeneration) {
    for (const childId of node.children) removeExplorerSubtree(childId);
    node.children = null;
  }

  if (explorer.nodes.size >= EXPLORER_LIMITS.maxNodes) {
    showExplorerWarning(`Explorer stopped loading new nodes at ${EXPLORER_LIMITS.maxNodes} nodes to protect workbook performance.`);
    return;
  }

  node.loading = true;
  renderExplorer();

  try {
    const addresses = await getDirectRelatedAddresses(node.address, explorer.relation);
    let externalWorkbooks = [];

    if (explorer.relation === "precedents") {
      if (node.nodeType === "selection") {
        externalWorkbooks = await getExternalWorkbookNamesForRange(node.address);
      } else if (node.loadedPreview) {
        externalWorkbooks = extractExternalWorkbookNames(node.formula);
      } else {
        externalWorkbooks = await getExternalWorkbookNamesForRange(node.address);
      }
    }

    const unique = [];
    const seen = new Set();

    for (const address of addresses) {
      const key = normalizeAddressKey(address);
      if (!key || seen.has(key)) continue;
      seen.add(key);
      unique.push(address);
    }

    if (node.nodeType === "selection") {
      unique.sort(compareExplorerAddressesByWorksheet);
    }

    node.children = [];

    for (const address of unique) {
      if (explorer.nodes.size >= EXPLORER_LIMITS.maxNodes) break;
      const child = createExplorerNode({
        address,
        parentId: node.id,
        depth: node.depth + 1
      });
      explorer.nodes.set(child.id, child);
      node.children.push(child.id);
      if (child.hidden) explorer.hiddenResultCount += 1;
    }

    for (const workbookName of externalWorkbooks) {
      if (explorer.nodes.size >= EXPLORER_LIMITS.maxNodes) break;
      const child = createExplorerNode({
        address: `[${workbookName}]`,
        parentId: node.id,
        depth: node.depth + 1,
        nodeType: "external",
        externalWorkbook: workbookName
      });
      explorer.nodes.set(child.id, child);
      node.children.push(child.id);
    }

    node.expanded = true;
    node.childrenGeneration = explorer.changeGeneration;

    refreshExplorerWarning();

    if (!options.keepSelection && node.children.length) {
      await setExplorerActive(node.children[0], { navigate: true, recordHistory: true });
      return;
    }

    renderExplorer();
  } catch (error) {
    if (isItemNotFound(error)) {
      node.children = [];
      node.expanded = true;
      node.childrenGeneration = explorer.changeGeneration;
      renderExplorer();
    } else {
      showExplorerWarning(`Could not expand ${node.address}: ${normalizeError(error)}`);
    }
  } finally {
    node.loading = false;
    renderExplorer();
  }
}


function removeExplorerSubtree(nodeId) {
  const node = explorer.nodes.get(nodeId);
  if (!node) return;
  if (Array.isArray(node.children)) node.children.forEach(removeExplorerSubtree);
  explorer.nodes.delete(nodeId);
}

async function getDirectRelatedAddresses(address, relation) {
  return Excel.run(async (context) => {
    const parsed = parseQualifiedAddress(address);
    const sheet = parsed.sheetName
      ? context.workbook.worksheets.getItem(parsed.sheetName)
      : context.workbook.worksheets.getActiveWorksheet();

    const range = sheet.getRange(parsed.cellAddress);
    let related;

    try {
      related = relation === "precedents"
        ? range.getDirectPrecedents()
        : range.getDirectDependents();
      related.load("addresses");
      await context.sync();
    } catch (error) {
      if (isItemNotFound(error)) return [];
      throw error;
    }

    return flattenWorkbookAddresses(related.addresses || []);
  });
}

async function getExternalWorkbookNamesForRange(address) {
  return Excel.run(async (context) => {
    const parsed = parseQualifiedAddress(address);
    const sheet = parsed.sheetName
      ? context.workbook.worksheets.getItem(parsed.sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(parsed.cellAddress);
    range.load("formulas");
    await context.sync();
    return extractExternalWorkbookNames(range.formulas || []);
  });
}

function extractExternalWorkbookNames(formulas) {
  const names = new Set();
  const scan = (value) => {
    if (Array.isArray(value)) {
      value.forEach(scan);
      return;
    }
    if (typeof value !== "string" || !value.startsWith("=")) return;

    const formula = removeQuotedText(value);
    const regex = /\[([^\]]+\.(?:xlsx|xlsm|xlsb|xls|xlam|csv))\]/gi;
    let match;
    while ((match = regex.exec(formula)) !== null) {
      if (match[1]) names.add(match[1]);
    }
  };
  scan(formulas);
  return [...names].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
}

function compareExplorerAddressesByWorksheet(a, b) {
  const pa = parseQualifiedAddress(a);
  const pb = parseQualifiedAddress(b);
  const sheetCompare = String(pa.sheetName || "").localeCompare(String(pb.sheetName || ""), undefined, { sensitivity: "base" });
  if (sheetCompare !== 0) return sheetCompare;
  return String(pa.cellAddress || "").localeCompare(String(pb.cellAddress || ""), undefined, { numeric: true, sensitivity: "base" });
}

function getVisibleExplorerNodes() {
  if (!explorer.active || !explorer.rootId) return [];
  const out = [];

  const visit = (id) => {
    const node = explorer.nodes.get(id);
    if (!node) return;
    out.push(node);
    if (node.expanded && Array.isArray(node.children)) {
      node.children.forEach(visit);
    }
  };

  visit(explorer.rootId);
  return out;
}

function renderExplorer() {
  const tree = $("explorer-tree");

  if (!explorer.active || !explorer.rootId) {
    renderExplorerEmpty("No active trace");
    return;
  }

  const visible = getVisibleExplorerNodes();
  tree.className = "explorer-tree";
  tree.innerHTML = "";

  $("explorer-title").textContent = explorer.relation === "precedents" ? "Precedents" : "Dependents";
  $("explorer-count").textContent = Math.max(0, visible.length - 1);
  $("explorer-source-btn").disabled = false;
  $("explorer-back-btn").disabled = explorer.history.length === 0;
  $("copy-btn").disabled = visible.length <= 1;
  $("explorer-refresh-trace-btn").classList.toggle("hidden", !explorer.stale);

  let lastRootGroup = null;

  visible.forEach((node) => {
    if (explorer.sourceCellCount > 1 && node.parentId === explorer.rootId) {
      const parsedGroup = node.nodeType === "external" ? null : parseQualifiedAddress(node.address);
      const groupName = node.nodeType === "external"
        ? "External workbooks"
        : (parsedGroup?.sheetName || "Current sheet");
      if (groupName !== lastRootGroup) {
        const header = document.createElement("div");
        header.className = "explorer-group-header";
        header.textContent = groupName;
        tree.appendChild(header);
        lastRootGroup = groupName;
      }
    }

    const row = document.createElement("div");
    row.id = `explorer-row-${node.id}`;
    row.className = [
      "explorer-row",
      node.id === explorer.rootId ? "root" : "",
      node.id === explorer.activeId ? "active" : "",
      node.hidden ? "hidden-sheet" : "",
      node.cycle ? "cycle" : "",
      node.nodeType === "external" ? "external-boundary" : ""
    ].filter(Boolean).join(" ");
    row.setAttribute("role", "treeitem");
    row.setAttribute("aria-level", String(node.depth + 1));
    row.setAttribute("aria-selected", node.id === explorer.activeId ? "true" : "false");
    if (Array.isArray(node.children) && node.children.length) {
      row.setAttribute("aria-expanded", node.expanded ? "true" : "false");
    }
    row.dataset.nodeId = node.id;
    row.style.setProperty("--depth", String(node.depth));

    const toggle = document.createElement("span");
    toggle.className = `tree-toggle${node.cycle || node.nodeType === "external" ? "" : " expandable"}`;
    toggle.dataset.nodeId = node.id;
    toggle.dataset.action = "toggle";
    if (node.loading) toggle.textContent = "…";
    else if (node.nodeType === "external") toggle.textContent = "";
    else if (node.cycle) toggle.textContent = "!";
    else if (node.depth >= EXPLORER_LIMITS.maxDepth) toggle.textContent = "";
    else if (Array.isArray(node.children) && node.children.length === 0) toggle.textContent = "";
    else if (Array.isArray(node.children)) toggle.textContent = node.expanded ? "▾" : "▸";
    else toggle.textContent = "▸";

    const spillOutput = node.spillParent && normalizeAddressKey(node.spillParent) !== normalizeAddressKey(node.address);
    const spillFormula = !spillOutput && node.spillRange && isFormulaValue(node.formula);
    const main = document.createElement("span");
    main.className = "explorer-row-main";

    const parsed = node.nodeType === "external" ? { sheetName: "", cellAddress: "" } : parseQualifiedAddress(node.address);
    const addressLine = document.createElement("div");
    addressLine.className = "explorer-row-address";
    if (node.nodeType === "external") {
      addressLine.textContent = node.externalWorkbook;
    } else if (node.nodeType === "selection" && node.id === explorer.rootId) {
      addressLine.textContent = `Selection, ${Number(node.cellCount || 0).toLocaleString()} cells`;
    } else {
      addressLine.textContent = parsed.cellAddress || node.address;
    }

    const sub = document.createElement("div");
    sub.className = "explorer-row-sub";
    if (node.nodeType === "external") {
      sub.textContent = "Workbook boundary";
    } else if (node.nodeType === "selection" && node.id === explorer.rootId) {
      sub.textContent = node.address;
    } else if (spillOutput) {
      sub.textContent = `Spilled from ${node.spillParent}`;
    } else if (spillFormula) {
      sub.textContent = `Spills to ${node.spillRange}`;
    } else {
      sub.textContent = parsed.sheetName || "Current sheet";
    }

    main.append(addressLine, sub);

    const meta = document.createElement("span");
    meta.className = "explorer-row-meta";
    if (node.nodeType === "external") meta.textContent = "EXTERNAL";
    else if (node.nodeType === "selection" && node.id === explorer.rootId) meta.textContent = "SELECTION";
    else if (spillOutput) meta.textContent = "SPILL OUTPUT";
    else if (spillFormula) meta.textContent = "SPILL FORMULA";
    else if (node.id === explorer.rootId) meta.textContent = "SOURCE";
    else if (node.hidden) meta.textContent = "HIDDEN";
    else if (node.cycle) meta.textContent = "CYCLE";
    else if (node.relationClass === "off-sheet") meta.textContent = "OFF-SHEET";
    else meta.textContent = "";
    meta.classList.toggle("off-sheet", node.relationClass === "off-sheet");
    meta.classList.toggle("external", node.nodeType === "external");
    meta.classList.toggle("warning", node.hidden || spillOutput || spillFormula);
    meta.classList.toggle("error", node.cycle);

    row.append(toggle, main, meta);
    tree.appendChild(row);
  });

  tree.setAttribute("aria-activedescendant", explorer.activeId ? `explorer-row-${explorer.activeId}` : "");
  scrollActiveExplorerRowIntoView();
}

function renderExplorerEmpty(message) {
  const tree = $("explorer-tree");
  tree.className = "explorer-tree empty";
  tree.innerHTML = `<div class="empty-state">${escapeHtml(message)}</div>`;
  $("explorer-title").textContent = "No active trace";
  $("explorer-count").textContent = "0";
  $("explorer-source-btn").disabled = true;
  $("explorer-back-btn").disabled = true;
  $("copy-btn").disabled = true;
  $("explorer-refresh-trace-btn").classList.add("hidden");
  $("explorer-preview").classList.add("hidden");
  clearExplorerWarning();
}

function renderExplorerError(message) {
  explorer = createEmptyExplorerState();
  const tree = $("explorer-tree");
  tree.className = "explorer-tree";
  tree.innerHTML = `<div class="error-box">${escapeHtml(message)}</div>`;
  $("explorer-title").textContent = "Explorer error";
  $("explorer-count").textContent = "0";
  $("explorer-source-btn").disabled = true;
  $("explorer-back-btn").disabled = true;
  $("copy-btn").disabled = true;
  $("explorer-refresh-trace-btn").classList.add("hidden");
}

function resetExplorer() {
  explorer = createEmptyExplorerState();
  renderExplorerEmpty("No active trace");
}

async function handleExplorerKeydown(event) {
  if (!explorer.active) return;

  const visible = getVisibleExplorerNodes();
  const index = visible.findIndex((node) => node.id === explorer.activeId);
  if (index < 0) return;

  if (event.key === "ArrowDown") {
    event.preventDefault();
    if (index < visible.length - 1) {
      await setExplorerActive(visible[index + 1].id, { navigate: true, recordHistory: true });
    }
    return;
  }

  if (event.key === "ArrowUp") {
    event.preventDefault();
    if (index > 0) {
      await setExplorerActive(visible[index - 1].id, { navigate: true, recordHistory: true });
    }
    return;
  }

  if (event.key === "ArrowRight") {
    event.preventDefault();
    const node = visible[index];

    if (node.cycle || node.nodeType === "external") return;

    if (!Array.isArray(node.children)) {
      await expandExplorerNode(node.id, { keepSelection: true });
      return;
    }

    if (!node.expanded && node.children.length) {
      node.expanded = true;
      renderExplorer();
      return;
    }

    if (node.expanded && node.children.length) {
      await setExplorerActive(node.children[0], { navigate: true, recordHistory: true });
    }
    return;
  }

  if (event.key === "ArrowLeft") {
    event.preventDefault();
    const node = visible[index];

    if (node.expanded && Array.isArray(node.children) && node.children.length) {
      node.expanded = false;
      renderExplorer();
      return;
    }

    if (node.parentId) {
      await setExplorerActive(node.parentId, { navigate: true, recordHistory: true });
    }
    return;
  }

  if (event.key === "Enter") {
    event.preventDefault();
    await navigateToExplorerNode(explorer.nodes.get(explorer.activeId));
    try {
      if (Office.addin && Office.addin.hide) await Office.addin.hide();
    } catch {}
    return;
  }

  if (event.key === "Backspace") {
    event.preventDefault();
    await explorerGoBack();
    return;
  }

  if (event.key.toLowerCase() === "r" && !event.ctrlKey && !event.metaKey && !event.altKey) {
    event.preventDefault();
    await refreshExplorerTrace(true);
    return;
  }

  if (event.key === "Escape") {
    event.preventDefault();
    if (explorer.history.length) {
      await explorerGoBack();
      return;
    }
    if (explorer.activeId !== explorer.rootId) {
      await setExplorerActive(explorer.rootId, { navigate: true, recordHistory: false });
      return;
    }
    await navigateToExplorerNode(explorer.nodes.get(explorer.rootId));
    try {
      if (Office.addin && Office.addin.hide) await Office.addin.hide();
    } catch {}
  }
}

async function handleExplorerClick(event) {
  const toggle = event.target.closest("[data-action='toggle']");
  if (toggle) {
    const nodeId = toggle.dataset.nodeId;
    const node = explorer.nodes.get(nodeId);
    if (!node || node.cycle || node.nodeType === "external") return;

    if (!Array.isArray(node.children)) {
      await expandExplorerNode(nodeId, { keepSelection: true });
    } else {
      node.expanded = !node.expanded;
      renderExplorer();
    }
    $("explorer-tree").focus({ preventScroll: true });
    return;
  }

  const row = event.target.closest(".explorer-row");
  if (!row) return;

  await setExplorerActive(row.dataset.nodeId, { navigate: true, recordHistory: true });
  $("explorer-tree").focus({ preventScroll: true });
}

async function setExplorerActive(nodeId, options = {}) {
  let next = explorer.nodes.get(nodeId);
  let current = explorer.nodes.get(explorer.activeId);
  if (!next) return;

  // A workbook edit invalidates loaded dependency children. Before the next keyboard jump,
  // refresh only the branch that contains the requested node and preserve the same address when possible.
  if (explorer.stale && options.navigate && next.parentId) {
    next = await refreshStaleBranchForNavigation(next);
    current = explorer.nodes.get(explorer.activeId);
    if (!next) return;
  }

  if (options.recordHistory && current && current.id !== next.id) {
    explorer.history.push(current.id);
    if (explorer.history.length > 100) explorer.history.shift();
  }

  explorer.activeId = next.id;
  renderExplorer();

  // Navigation already selects the cell and refreshes formula/value/spill preview in one Excel.run.
  // Avoid a separate preview read before every arrow-key jump.
  if (options.navigate) {
    await navigateToExplorerNode(next);
  } else {
    await ensureExplorerPreview(next, { force: next.previewGeneration !== explorer.changeGeneration });
  }
}

async function refreshStaleBranchForNavigation(node) {
  if (!node?.parentId) return node;
  const parent = explorer.nodes.get(node.parentId);
  if (!parent || parent.cycle || parent.nodeType === "external") return node;

  const desiredKey = normalizeAddressKey(node.address);
  const oldChildren = Array.isArray(parent.children) ? [...parent.children] : [];
  for (const childId of oldChildren) removeExplorerSubtree(childId);
  parent.children = null;
  parent.childrenGeneration = -1;
  parent.expanded = true;
  explorer.history = explorer.history.filter((id) => explorer.nodes.has(id));
  if (!explorer.nodes.has(explorer.activeId)) explorer.activeId = parent.id;

  await expandExplorerNode(parent.id, { keepSelection: true });
  const refreshedParent = explorer.nodes.get(parent.id);
  if (!refreshedParent) return null;
  const replacementId = (refreshedParent.children || []).find((id) => {
    const candidate = explorer.nodes.get(id);
    return candidate && normalizeAddressKey(candidate.address) === desiredKey;
  });
  return replacementId ? explorer.nodes.get(replacementId) : refreshedParent;
}

async function navigateToExplorerNode(node) {
  if (!node) return;

  if (node.nodeType === "external") {
    updateExplorerPreview(node);
    return;
  }

  if (node.hidden) {
    showExplorerWarning(`Cannot jump to ${node.address}. The worksheet is hidden.`);
    return;
  }

  const token = ++explorer.navigationToken;

  try {
    const preview = await Excel.run(async (context) => {
      const parsed = parseQualifiedAddress(node.address);
      const sheet = parsed.sheetName
        ? context.workbook.worksheets.getItem(parsed.sheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      sheet.activate();
      const range = sheet.getRange(parsed.cellAddress);
      const first = range.getCell(0, 0);
      first.load(["formulas", "values"]);
      let spillParent = null;
      let spillRange = null;
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.12")) {
        spillParent = first.getSpillParentOrNullObject();
        spillRange = first.getSpillingToRangeOrNullObject();
        spillParent.load(["isNullObject", "address"]);
        spillRange.load(["isNullObject", "address"]);
      }
      range.select();
      await context.sync();

      return {
        formula: first.formulas?.[0]?.[0],
        value: first.values?.[0]?.[0],
        spillParent: spillParent && !spillParent.isNullObject ? qualifyAddress(parsed.sheetName || sheet.name, localAddress(spillParent.address)) : "",
        spillRange: spillRange && !spillRange.isNullObject ? qualifyAddress(parsed.sheetName || sheet.name, localAddress(spillRange.address)) : ""
      };
    });

    if (token !== explorer.navigationToken) return;

    const spillChanged = node.spillParent !== (preview.spillParent || "") || node.spillRange !== (preview.spillRange || "");
    node.formula = preview.formula;
    node.value = preview.value;
    node.spillParent = preview.spillParent || "";
    node.spillRange = preview.spillRange || "";
    node.loadedPreview = true;
    node.previewGeneration = explorer.changeGeneration;
    if (spillChanged) renderExplorer();
    updateExplorerPreview(node);
  } catch (error) {
    showExplorerWarning(`Could not navigate to ${node.address}: ${normalizeError(error)}`);
  }
}

async function ensureExplorerPreview(node, options = {}) {
  if (node?.nodeType === "external" || node?.nodeType === "selection") {
    updateExplorerPreview(node);
    return;
  }

  if (!node || (node.loadedPreview && !options.force && node.previewGeneration === explorer.changeGeneration)) {
    if (node) updateExplorerPreview(node);
    return;
  }

  try {
    const preview = await Excel.run(async (context) => {
      const parsed = parseQualifiedAddress(node.address);
      const sheet = parsed.sheetName
        ? context.workbook.worksheets.getItem(parsed.sheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = sheet.getRange(parsed.cellAddress).getCell(0, 0);
      range.load(["formulas", "values"]);
      let spillParent = null;
      let spillRange = null;
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.12")) {
        spillParent = range.getSpillParentOrNullObject();
        spillRange = range.getSpillingToRangeOrNullObject();
        spillParent.load(["isNullObject", "address"]);
        spillRange.load(["isNullObject", "address"]);
      }
      await context.sync();

      return {
        formula: range.formulas?.[0]?.[0],
        value: range.values?.[0]?.[0],
        spillParent: spillParent && !spillParent.isNullObject ? qualifyAddress(parsed.sheetName || sheet.name, localAddress(spillParent.address)) : "",
        spillRange: spillRange && !spillRange.isNullObject ? qualifyAddress(parsed.sheetName || sheet.name, localAddress(spillRange.address)) : ""
      };
    });

    const spillChanged = node.spillParent !== (preview.spillParent || "") || node.spillRange !== (preview.spillRange || "");
    node.formula = preview.formula;
    node.value = preview.value;
    node.spillParent = preview.spillParent || "";
    node.spillRange = preview.spillRange || "";
    node.loadedPreview = true;
    node.previewGeneration = explorer.changeGeneration;
    if (spillChanged) renderExplorer();
    updateExplorerPreview(node);
  } catch {
    updateExplorerPreview(node);
  }
}

function updateExplorerPreview(node) {
  if (!node) {
    $("explorer-preview").classList.add("hidden");
    return;
  }

  const preview = $("explorer-preview");
  const chip = $("preview-kind");

  if (node.nodeType === "external") {
    chip.className = "relation-label external";
    chip.textContent = "EXTERNAL";
    $("preview-address").textContent = node.externalWorkbook;
    $("preview-formula").textContent = `External workbook: ${node.externalWorkbook}`;
    $("preview-value").textContent = "Cross-workbook tracing unavailable";
    preview.classList.remove("hidden");
    return;
  }

  if (node.nodeType === "selection" && node.id === explorer.rootId) {
    chip.className = "relation-label selection";
    chip.textContent = "SELECTION";
    $("preview-address").textContent = node.address;
    $("preview-formula").textContent = `${Number(node.cellCount || 0).toLocaleString()} source cells`;
    $("preview-value").textContent = "Results grouped by worksheet";
    preview.classList.remove("hidden");
    return;
  }

  const spillParentDifferent = node.spillParent && normalizeAddressKey(node.spillParent) !== normalizeAddressKey(node.address);
  if (spillParentDifferent) {
    chip.className = "relation-label spill";
    chip.textContent = "SPILL OUTPUT";
    $("preview-address").textContent = node.address;
    $("preview-formula").textContent = `Spilled from ${node.spillParent}`;
    const value = node.value;
    $("preview-value").textContent = value === null || typeof value === "undefined" || value === ""
      ? "Displayed value: —"
      : `Displayed value: ${String(value)}`;
    preview.classList.remove("hidden");
    return;
  }

  if (node.spillRange && typeof node.formula === "string" && node.formula.startsWith("=")) {
    chip.className = "relation-label spill";
    chip.textContent = "SPILL FORMULA";
    $("preview-address").textContent = node.address;
    $("preview-formula").textContent = node.formula;
    $("preview-value").textContent = `Spills to ${node.spillRange}`;
    preview.classList.remove("hidden");
    return;
  }

  chip.className = `relation-label ${node.relationClass || ""}`.trim();
  chip.textContent =
    node.id === explorer.rootId ? "SOURCE" :
    node.hidden ? "HIDDEN" :
    node.cycle ? "CYCLE" :
    node.relationClass === "off-sheet" ? "OFF-SHEET" :
    "SAME-SHEET";

  $("preview-address").textContent = node.address;

  const formula = node.formula;
  if (typeof formula === "string" && formula.startsWith("=")) {
    $("preview-formula").textContent = formula;
  } else {
    $("preview-formula").textContent = "Hardcode / value";
  }

  const value = node.value;
  $("preview-value").textContent =
    value === null || typeof value === "undefined" || value === ""
      ? "Displayed value: —"
      : `Displayed value: ${String(value)}`;

  preview.classList.remove("hidden");
}

function updateSourceCard(address, formula, value, cellCount = 1) {
  currentAddress = address;

  if (cellCount > 1) {
    $("source-address").textContent = `Selection, ${Number(cellCount).toLocaleString()} cells`;
    $("source-formula").textContent = address;
    $("source-formula").classList.add("muted");
    return;
  }

  $("source-address").textContent = address;

  if (typeof formula === "string" && formula.startsWith("=")) {
    $("source-formula").textContent = formula;
    $("source-formula").classList.remove("muted");
  } else {
    $("source-formula").textContent = `Value: ${value ?? ""}`;
    $("source-formula").classList.add("muted");
  }
}

async function goToExplorerSource() {
  if (!explorer.rootId) return;
  await setExplorerActive(explorer.rootId, { navigate: true, recordHistory: true });
  $("explorer-tree").focus({ preventScroll: true });
}

async function explorerGoBack() {
  if (!explorer.history.length) return;

  let previousId = explorer.history.pop();
  while (previousId && !explorer.nodes.has(previousId) && explorer.history.length) {
    previousId = explorer.history.pop();
  }

  if (!previousId || !explorer.nodes.has(previousId)) {
    renderExplorer();
    return;
  }

  await setExplorerActive(previousId, { navigate: true, recordHistory: false });
  $("explorer-tree").focus({ preventScroll: true });
}

function scrollActiveExplorerRowIntoView() {
  if (!explorer.activeId) return;
  const row = document.getElementById(`explorer-row-${explorer.activeId}`);
  if (row) row.scrollIntoView({ block: "nearest" });
}

function showExplorerWarning(message) {
  const warning = $("explorer-warning");
  warning.textContent = message;
  warning.classList.remove("hidden");
}

function clearExplorerWarning() {
  const warning = $("explorer-warning");
  warning.textContent = "";
  warning.classList.add("hidden");
}


function refreshExplorerWarning() {
  if (!explorer.active) {
    clearExplorerWarning();
    return;
  }

  const notices = [];
  if (explorer.stale) notices.push("Model changed · loaded trace branches refresh as you navigate. Press R for a full trace refresh.");
  if (explorer.hiddenResultCount) {
    notices.push(`${explorer.hiddenResultCount} result${explorer.hiddenResultCount === 1 ? "" : "s"} on hidden sheet${explorer.hiddenResultCount === 1 ? "" : "s"}.`);
  }
  if (explorer.boundaryNotice) notices.push(explorer.boundaryNotice);
  if (explorer.calculationNotice) notices.push(explorer.calculationNotice);

  if (notices.length) showExplorerWarning(notices.join(" "));
  else clearExplorerWarning();
}

async function refreshExplorerTrace(full = false) {
  if (!explorer.active || !explorer.rootId) return;
  const targetId = full ? explorer.rootId : explorer.activeId;
  const target = explorer.nodes.get(targetId);
  if (!target) return;

  if (full) {
    for (const childId of target.children || []) removeExplorerSubtree(childId);
    target.children = null;
    explorer.hiddenResultCount = 0;
    explorer.history = [];
  } else if (Array.isArray(target.children)) {
    for (const childId of target.children) removeExplorerSubtree(childId);
    target.children = null;
  }

  target.loadedPreview = target.nodeType !== "cell";
  target.previewGeneration = -1;
  target.childrenGeneration = -1;
  explorer.stale = false;
  $("explorer-refresh-trace-btn").classList.add("hidden");

  if (target.nodeType === "cell") await ensureExplorerPreview(target, { force: true });
  await expandExplorerNode(target.id, { keepSelection: true });
  target.expanded = true;
  renderExplorer();
  refreshExplorerWarning();
  $("explorer-tree").focus({ preventScroll: true });
}

async function copyExplorerTree() {
  const visible = getVisibleExplorerNodes();
  if (visible.length <= 1) return;

  const relationLabel = explorer.relation === "precedents" ? "Precedents" : "Dependents";
  const lines = [
    `${relationLabel} explorer`,
    `Source: ${explorer.sourceAddress}`,
    ...visible.slice(1).map((node) => `${"  ".repeat(Math.max(0, node.depth - 1))}- ${node.nodeType === "external" ? `External workbook boundary: ${node.externalWorkbook}` : node.address}`)
  ];

  try {
    await navigator.clipboard.writeText(lines.join("\n"));
    const button = $("copy-btn");
    const old = button.textContent;
    button.textContent = "Copied";
    setTimeout(() => button.textContent = old, 1000);
  } catch {
    showExplorerWarning("Clipboard access was blocked by Excel.");
  }
}

async function openModelTracerCommand(event) {
  try {
    if (Office.addin && Office.addin.showAsTaskpane) {
      await Office.addin.showAsTaskpane();
    }
    switchTab("trace");
    if (explorer.active && explorer.activeId) {
      renderExplorer();
      const tree = $("explorer-tree");
      tree.focus({ preventScroll: true });
    }
  } catch (error) {
    showExplorerWarning(`Could not open Model Tracer: ${normalizeError(error)}`);
  } finally {
    try { event?.completed?.(); } catch {}
  }
}

async function launchExplorerCommand(relation, event) {
  try {
    if (Office.addin && Office.addin.showAsTaskpane) {
      await Office.addin.showAsTaskpane();
    }
    switchTab("trace");
    await startExplorer(relation);
  } catch (error) {
    showExplorerWarning(`Could not open ${relation}: ${normalizeError(error)}`);
  } finally {
    try { event?.completed?.(); } catch {}
  }
}

async function explorerBackCommand(event) {
  try {
    if (Office.addin && Office.addin.showAsTaskpane) {
      await Office.addin.showAsTaskpane();
    }
    switchTab("trace");
    await explorerGoBack();
  } catch (error) {
    showExplorerWarning(`Could not go back: ${normalizeError(error)}`);
  } finally {
    try { event?.completed?.(); } catch {}
  }
}

/* ----------------------------
   Financial-model formatting
----------------------------- */

async function formatScope(scope) {
  const resetLocal = $("reset-local-formulas").checked;
  const includeText = $("include-text-hardcodes").checked;

  const label =
    scope === "selection" ? "selected range" :
    scope === "sheet" ? "current sheet" :
    "workbook";

  setFormatSummary(`Formatting ${label}…`, "");
  setWorking(true, "Formatting");

  try {
    const totals = await Excel.run(async (context) => {
      let jobs = [];

      if (scope === "selection") {
        const range = context.workbook.getSelectedRange();
        range.load(["formulas", "values", "rowCount", "columnCount", "address"]);
        jobs.push({ range, name: "Selection" });
        await context.sync();
      } else if (scope === "sheet") {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        const range = sheet.getUsedRangeOrNullObject(true);
        range.load(["isNullObject", "formulas", "values", "rowCount", "columnCount", "address"]);
        jobs.push({ range, sheet });
        await context.sync();
        if (range.isNullObject) return emptyTotals();
      } else {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();

        for (const sheet of sheets.items) {
          const range = sheet.getUsedRangeOrNullObject(true);
          range.load(["isNullObject", "formulas", "values", "rowCount", "columnCount", "address"]);
          jobs.push({ range, sheet });
        }
        await context.sync();
        jobs = jobs.filter((job) => !job.range.isNullObject);
      }

      const totals = emptyTotals();

      for (const job of jobs) {
        const stats = applyFinancialModelFormatting(
          job.range,
          job.range.formulas || [],
          job.range.values || [],
          resetLocal,
          includeText
        );
        totals.hardcodes += stats.hardcodes;
        totals.offsheets += stats.offsheets;
        totals.localFormulas += stats.localFormulas;
        totals.textHardcodes += stats.textHardcodes;
        totals.untouched += stats.untouched;
        totals.cells += stats.cells;
        totals.ranges += 1;
      }

      await context.sync();
      return totals;
    });

    const parts = [
      `${totals.hardcodes} hardcode${totals.hardcodes === 1 ? "" : "s"} → blue`,
      `${totals.offsheets} off-sheet formula${totals.offsheets === 1 ? "" : "s"} → green`
    ];
    if (resetLocal) parts.push(`${totals.localFormulas} same-sheet formula${totals.localFormulas === 1 ? "" : "s"} → black`);
    if (includeText && totals.textHardcodes) parts.push(`${totals.textHardcodes} text hardcode${totals.textHardcodes === 1 ? "" : "s"} → blue`);

    setFormatSummary(`Done. ${parts.join("; ")}`, "success");
  } catch (error) {
    const message = normalizeError(error);
    if (scope === "selection" && /multiple|areas|selected range/i.test(message)) {
      setFormatSummary("Select one contiguous range and try again. Multi-area selections are not supported yet.", "error");
    } else {
      setFormatSummary(`Could not format ${label}: ${message}`, "error");
    }
  } finally {
    setWorking(false);
  }
}

function emptyTotals() {
  return {
    hardcodes: 0,
    offsheets: 0,
    localFormulas: 0,
    textHardcodes: 0,
    untouched: 0,
    cells: 0,
    ranges: 0
  };
}

function applyFinancialModelFormatting(range, formulas, values, resetLocal, includeText) {
  const stats = emptyTotals();
  const rows = formulas.length;

  for (let r = 0; r < rows; r++) {
    const cols = formulas[r]?.length || 0;
    let runColor = null;
    let runStart = -1;

    const flush = (endExclusive) => {
      if (runColor === null || runStart < 0 || endExclusive <= runStart) return;
      const run = range.getCell(r, runStart).getResizedRange(0, endExclusive - runStart - 1);
      run.format.font.color = runColor;
      runColor = null;
      runStart = -1;
    };

    for (let c = 0; c < cols; c++) {
      stats.cells += 1;

      const formula = formulas[r][c];
      const value = values?.[r]?.[c];
      const classification = classifyCell(formula, value, resetLocal, includeText);

      if (classification.kind === "hardcode") stats.hardcodes += 1;
      else if (classification.kind === "offsheet") stats.offsheets += 1;
      else if (classification.kind === "local") stats.localFormulas += 1;
      else if (classification.kind === "textHardcode") stats.textHardcodes += 1;
      else stats.untouched += 1;

      const nextColor = classification.color;

      if (nextColor !== runColor) {
        flush(c);
        if (nextColor) {
          runColor = nextColor;
          runStart = c;
        }
      }
    }

    flush(cols);
  }

  return stats;
}

function classifyCell(formula, value, resetLocal, includeText) {
  if (typeof formula === "string" && formula.startsWith("=")) {
    if (formulaReferencesOtherSheet(formula)) {
      return { kind: "offsheet", color: COLORS.offsheet };
    }
    if (resetLocal) {
      return { kind: "local", color: COLORS.formula };
    }
    return { kind: "untouched", color: null };
  }

  // Excel dates are numeric serial values, so they naturally fall into this branch.
  if (typeof value === "number" || typeof value === "boolean") {
    return { kind: "hardcode", color: COLORS.hardcode };
  }

  if (includeText && typeof value === "string" && value.trim() !== "") {
    return { kind: "textHardcode", color: COLORS.hardcode };
  }

  return { kind: "untouched", color: null };
}

function formulaReferencesOtherSheet(formula) {
  // Remove quoted text so a formula such as ="Hello!" is not misclassified.
  let outsideStrings = "";
  let inString = false;

  for (let i = 0; i < formula.length; i++) {
    const ch = formula[i];

    if (ch === '"') {
      if (inString && formula[i + 1] === '"') {
        i++;
        continue;
      }
      inString = !inString;
      continue;
    }

    if (!inString) outsideStrings += ch;
  }

  return outsideStrings.includes("!");
}

async function setSelectedFontColor(color, label) {
  setWorking(true, "Formatting");
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.font.color = color;
      await context.sync();
    });
    setFormatSummary(`${label} font applied to the selected range.`, "success");
  } catch (error) {
    setFormatSummary(`Could not apply ${label.toLowerCase()}: ${normalizeError(error)}`, "error");
  } finally {
    setWorking(false);
  }
}

function setFormatSummary(text, mode = "") {
  const el = $("format-summary");
  el.textContent = text;
  el.className = `inline-result ${mode}`.trim();
}


/* ----------------------------
   Keyboard shortcut actions
----------------------------- */

function registerShortcutActions() {
  try {
    if (!Office.actions || !Office.actions.associate) return;

    Office.actions.associate("ColorHardcodeBlue", (event) => {
      Promise.resolve(applyShortcutColor(COLORS.hardcode))
        .finally(() => { try { event?.completed?.(); } catch {} });
    });
    Office.actions.associate("ColorOffsheetGreen", (event) => {
      Promise.resolve(applyShortcutColor(COLORS.offsheet))
        .finally(() => { try { event?.completed?.(); } catch {} });
    });
    Office.actions.associate("ColorFormulaBlack", (event) => {
      Promise.resolve(applyShortcutColor(COLORS.formula))
        .finally(() => { try { event?.completed?.(); } catch {} });
    });

    Office.actions.associate("OpenModelTracer", (event) => openModelTracerCommand(event));
    Office.actions.associate("ExplorePrecedents", (event) => launchExplorerCommand("precedents", event));
    Office.actions.associate("ExploreDependents", (event) => launchExplorerCommand("dependents", event));
    Office.actions.associate("ExplorerBack", (event) => explorerBackCommand(event));
    Office.actions.associate("RecalculateWorkbook", (event) => recalculateWorkbook(event));
    Office.actions.associate("OpenFormulaLogic", (event) => analysisModeCommand("logic", event));
    Office.actions.associate("ReviewCalculationArea", (event) => analysisModeCommand("area", event));
  } catch {}
}

function applyShortcutColor(color) {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.format.font.color = color;
    await context.sync();
  });
}



/* ----------------------------
   User-remappable keyboard shortcuts
----------------------------- */

function shortcutPlatform() {
  const platform = String(Office.context.platform || "").toLowerCase();
  if (platform.includes("mac")) return "mac";
  if (platform.includes("online")) return "web";
  if (platform.includes("pc") || platform.includes("win")) return "windows";
  const ua = String(navigator.userAgent || "").toLowerCase();
  if (ua.includes("mac")) return "mac";
  return "windows";
}

function shortcutProfile(name) {
  const platform = shortcutPlatform();
  const modifier = name === "banking"
    ? (platform === "mac" ? "Command+Option" : "Ctrl+Alt")
    : (platform === "mac" ? "Command+Shift+Option" : "Ctrl+Shift+Alt");
  return Object.fromEntries(SHORTCUT_ACTIONS.map((action) => [action.id, `${modifier}+${action.safeKey}`]));
}

function shortcutEditorValues() {
  const values = {};
  SHORTCUT_ACTIONS.forEach((action) => {
    const input = $(`shortcut-input-${action.id}`);
    if (input) values[action.id] = input.value.trim().replace(/\s*\+\s*/g, "+");
  });
  return values;
}

function setShortcutEditorValues(values) {
  SHORTCUT_ACTIONS.forEach((action) => {
    const input = $(`shortcut-input-${action.id}`);
    if (input && Object.prototype.hasOwnProperty.call(values, action.id)) input.value = values[action.id] || "";
    setShortcutRowState(action.id, "", "");
  });
}

function setShortcutRowState(actionId, text, stateClass) {
  const state = $(`shortcut-state-${actionId}`);
  if (!state) return;
  state.textContent = text || "—";
  state.className = `shortcut-state ${stateClass || ""}`.trim();
}

function renderShortcutEditor(values = {}) {
  const editor = $("shortcut-editor");
  if (!editor) return;
  editor.innerHTML = "";
  SHORTCUT_ACTIONS.forEach((action) => {
    const row = document.createElement("div");
    row.className = "shortcut-edit-row";

    const label = document.createElement("div");
    label.className = "shortcut-edit-label";
    const strong = document.createElement("strong");
    strong.textContent = action.label;
    const small = document.createElement("small");
    small.textContent = action.hint;
    label.append(strong, small);

    const input = document.createElement("input");
    input.id = `shortcut-input-${action.id}`;
    input.className = "shortcut-input";
    input.type = "text";
    input.autocomplete = "off";
    input.spellcheck = false;
    input.setAttribute("aria-label", `${action.label} shortcut`);
    input.value = values[action.id] || "";
    input.disabled = !shortcutSettingsSupported;
    input.addEventListener("input", () => setShortcutRowState(action.id, "Changed", "active"));

    const status = document.createElement("span");
    status.id = `shortcut-state-${action.id}`;
    status.className = "shortcut-state";
    status.textContent = "—";

    row.append(label, input, status);
    editor.appendChild(row);
  });
}

async function refreshShortcutSummary() {
  const summary = $("shortcut-summary-list");
  if (!summary) return;

  if (!shortcutSettingsSupported) {
    summary.innerHTML = '<div class="mini-note">Shortcut remapping is unavailable on this Excel build. Ribbon KeyTips and task-pane controls remain available.</div>';
    return;
  }

  try {
    shortcutCurrentValues = await Office.actions.getShortcuts();
    const ids = ["ColorHardcodeBlue", "ColorOffsheetGreen", "ColorFormulaBlack"];
    const labels = {
      ColorHardcodeBlue: "Blue hardcode",
      ColorOffsheetGreen: "Green off-sheet",
      ColorFormulaBlack: "Black formula"
    };
    summary.innerHTML = "";
    ids.forEach((id) => {
      const row = document.createElement("div");
      row.className = "shortcut-row";
      const label = document.createElement("span");
      label.textContent = labels[id];
      const key = document.createElement("kbd");
      key.textContent = shortcutCurrentValues[id] || "Conflict / unavailable";
      row.append(label, key);
      summary.appendChild(row);
    });
  } catch (error) {
    summary.innerHTML = `<div class="mini-note">Could not read shortcut bindings: ${escapeHtml(normalizeError(error))}</div>`;
  }
}

async function openShortcutSettings() {
  const modal = $("shortcut-modal");
  if (!modal) return;
  modal.classList.remove("hidden");

  const support = $("shortcut-support-note");
  const controls = ["shortcut-profile-safe", "shortcut-profile-banking", "shortcut-reset", "shortcut-check", "shortcut-apply"];
  controls.forEach((id) => { if ($(id)) $(id).disabled = !shortcutSettingsSupported; });

  if (!shortcutSettingsSupported) {
    support.textContent = "Shortcut remapping is unavailable in this Excel build. Alt KeyTips and task-pane navigation remain available.";
    support.className = "shortcut-support-note warn";
    renderShortcutEditor({});
    return;
  }

  support.textContent = "Remapping is supported. Microsoft 365 saves custom bindings for your signed-in account on this platform.";
  support.className = "shortcut-support-note ok";
  setShortcutMessage("Loading current bindings…", "");

  try {
    shortcutCurrentValues = await Office.actions.getShortcuts();
    renderShortcutEditor(shortcutCurrentValues);
    markCurrentShortcutRows();
    setShortcutMessage("Edit a binding, choose a profile, or check conflicts before applying.", "");
    const first = $("shortcut-editor")?.querySelector("input");
    first?.focus({ preventScroll: true });
  } catch (error) {
    renderShortcutEditor({});
    setShortcutMessage(`Could not load shortcuts: ${normalizeError(error)}`, "error");
  }
}

function closeShortcutSettings() {
  $("shortcut-modal")?.classList.add("hidden");
}

function loadShortcutProfile(name) {
  if (!shortcutSettingsSupported) return;
  setShortcutEditorValues(shortcutProfile(name));
  setShortcutMessage(name === "banking"
    ? "Banking / PF profile loaded. Run conflict check before applying."
    : "Safe default profile loaded. Run conflict check or apply directly.", "");
}

async function resetShortcutDefaults() {
  if (!shortcutSettingsSupported) return;
  const reset = Object.fromEntries(SHORTCUT_ACTIONS.map((action) => [action.id, null]));
  setShortcutBusy(true);
  try {
    await Office.actions.replaceShortcuts(reset);
    shortcutCurrentValues = await Office.actions.getShortcuts();
    renderShortcutEditor(shortcutCurrentValues);
    markCurrentShortcutRows();
    await refreshShortcutSummary();
    setShortcutMessage("Default shortcuts restored.", "success");
  } catch (error) {
    setShortcutMessage(`Could not reset shortcuts: ${normalizeError(error)}`, "error");
  } finally {
    setShortcutBusy(false);
  }
}

function shortcutDuplicates(values) {
  const seen = new Map();
  const duplicates = new Set();
  for (const [id, raw] of Object.entries(values)) {
    const key = String(raw || "").toUpperCase();
    if (!key) continue;
    if (seen.has(key)) {
      duplicates.add(id);
      duplicates.add(seen.get(key));
    } else {
      seen.set(key, id);
    }
  }
  return duplicates;
}

async function checkShortcutConflicts(silent = false) {
  if (!shortcutSettingsSupported) return { conflicts: 0, duplicates: 0 };
  const values = shortcutEditorValues();
  const missing = Object.entries(values).filter(([, value]) => !value).map(([id]) => id);
  if (missing.length) {
    missing.forEach((id) => setShortcutRowState(id, "Required", "conflict"));
    if (!silent) setShortcutMessage("Assign a shortcut to every action or reset defaults.", "error");
    return { conflicts: 0, duplicates: 0, invalid: true };
  }

  const duplicates = shortcutDuplicates(values);
  SHORTCUT_ACTIONS.forEach((action) => setShortcutRowState(action.id, "Checking", ""));
  duplicates.forEach((id) => setShortcutRowState(id, "Duplicate", "duplicate"));

  const uniqueKeys = [...new Set(Object.values(values))];
  try {
    const report = await Office.actions.areShortcutsInUse(uniqueKeys);
    const inUse = new Map(report.map((item) => [String(item.shortcut).toUpperCase(), Boolean(item.inUse)]));
    let conflicts = 0;
    SHORTCUT_ACTIONS.forEach((action) => {
      if (duplicates.has(action.id)) return;
      const key = values[action.id];
      if (inUse.get(String(key).toUpperCase())) {
        conflicts += 1;
        setShortcutRowState(action.id, "Conflict", "conflict");
      } else {
        setShortcutRowState(action.id, "Available", "available");
      }
    });
    if (!silent) {
      if (duplicates.size) setShortcutMessage("Resolve duplicate Model Tracer bindings before applying.", "error");
      else if (conflicts) setShortcutMessage(`${conflicts} binding${conflicts === 1 ? "" : "s"} may conflict with Excel or another add-in. You can still apply them; Office may ask which action should win.`, "warning");
      else setShortcutMessage("No conflicts detected.", "success");
    }
    return { conflicts, duplicates: duplicates.size, invalid: false };
  } catch (error) {
    if (!silent) setShortcutMessage(`Could not check conflicts: ${normalizeError(error)}`, "error");
    return { conflicts: 0, duplicates: duplicates.size, invalid: true };
  }
}

async function applyShortcutSettings() {
  if (!shortcutSettingsSupported) return;
  const values = shortcutEditorValues();
  const check = await checkShortcutConflicts(true);
  if (check.invalid || check.duplicates) {
    setShortcutMessage("Fix missing or duplicate bindings before applying.", "error");
    return;
  }

  setShortcutBusy(true);
  try {
    await Office.actions.replaceShortcuts(values);
    shortcutCurrentValues = await Office.actions.getShortcuts();
    setShortcutEditorValues(shortcutCurrentValues);
    markCurrentShortcutRows();
    await refreshShortcutSummary();
    setShortcutMessage(check.conflicts
      ? "Shortcuts applied. One or more conflicts may trigger an Office choice dialog when first used."
      : "Shortcuts applied.", check.conflicts ? "warning" : "success");
  } catch (error) {
    setShortcutMessage(`Could not apply shortcuts: ${normalizeError(error)}. Check that each combination uses supported modifier keys.`, "error");
  } finally {
    setShortcutBusy(false);
  }
}

function markCurrentShortcutRows() {
  SHORTCUT_ACTIONS.forEach((action) => {
    const value = shortcutCurrentValues[action.id];
    setShortcutRowState(action.id, value ? "Active" : "Conflict", value ? "active" : "conflict");
  });
}

function setShortcutBusy(isBusy) {
  ["shortcut-profile-safe", "shortcut-profile-banking", "shortcut-reset", "shortcut-check", "shortcut-apply", "shortcut-cancel"].forEach((id) => {
    const el = $(id);
    if (el) el.disabled = isBusy;
  });
  document.querySelectorAll(".shortcut-input").forEach((input) => { input.disabled = isBusy || !shortcutSettingsSupported; });
}

function setShortcutMessage(text, mode = "") {
  const message = $("shortcut-message");
  if (!message) return;
  message.textContent = text;
  message.className = `shortcut-message ${mode}`.trim();
}


async function analysisModeCommand(mode, event) {
  try {
    if (Office.addin && Office.addin.showAsTaskpane) await Office.addin.showAsTaskpane();
    switchTab("trace");
    if (mode === "logic") await openFormulaLogic();
    else await reviewCalculationArea();
  } catch (error) {
    showExplorerWarning(`Could not open ${mode === "logic" ? "Formula Logic" : "Calculation Area"}: ${normalizeError(error)}`);
  } finally {
    try { event?.completed?.(); } catch {}
  }
}

/* ----------------------------
   Formula Logic
----------------------------- */

async function openFormulaLogic() {
  setWorking(true, "Reading formula");
  try {
    const snapshot = await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      range.load(["address", "rowCount", "columnCount", "formulas", "formulasR1C1"]);
      sheet.load("name");
      await context.sync();
      const cells = Number(range.rowCount || 1) * Number(range.columnCount || 1);
      if (cells !== 1) throw new Error("Formula Logic works on one cell at a time. Select a single formula cell.");
      return {
        address: qualifyAddress(sheet.name, localAddress(range.address)),
        sheetName: sheet.name,
        formula: range.formulas?.[0]?.[0],
        formulaR1C1: range.formulasR1C1?.[0]?.[0]
      };
    });

    if (!isFormulaValue(snapshot.formula)) {
      throw new Error("The selected cell does not contain a formula.");
    }

    logicState = createEmptyLogicState();
    logicState.active = true;
    logicState.sourceAddress = snapshot.address;
    logicState.sourceSheet = snapshot.sheetName;
    logicState.formula = String(snapshot.formula);
    logicState.formulaR1C1 = String(snapshot.formulaR1C1 || "");

    const descriptor = parseFormulaLogicExpression(logicState.formula.slice(1), "Formula");
    const root = materializeLogicNode(descriptor, "", 0);
    logicState.rootId = root.id;
    logicState.activeId = root.id;

    showAnalysisMode("logic");
    renderLogicTree();
    updateLogicPreview(root);
    $("logic-tree").focus({ preventScroll: true });
  } catch (error) {
    showAnalysisMode("logic");
    $("logic-tree").innerHTML = `<div class="error-box">${escapeHtml(normalizeError(error))}</div>`;
    $("logic-note").classList.add("hidden");
  } finally {
    setWorking(false);
  }
}

function parseFormulaLogicExpression(expression, labelHint = "") {
  let text = String(expression || "").trim();
  if (!text) return { type: "value", label: labelHint || "Value", expression: "" };

  if (isWrappedBySingleParentheses(text)) {
    return {
      type: "group",
      label: labelHint || "Group",
      expression: text,
      children: [parseFormulaLogicExpression(text.slice(1, -1), "Expression")]
    };
  }

  const fn = getWholeFunctionCall(text);
  if (fn) {
    const args = splitTopLevelFormulaArgs(fn.args);
    const labels = formulaArgumentLabels(fn.name, args.length);
    return {
      type: "function",
      label: labelHint ? `${labelHint}: ${fn.name}` : fn.name,
      expression: text,
      children: args.map((arg, i) => parseFormulaLogicExpression(arg, labels[i] || `Arg ${i + 1}`))
    };
  }

  const operator = findTopLevelFormulaOperator(text);
  if (operator) {
    const left = text.slice(0, operator.index);
    const right = text.slice(operator.index + operator.op.length);
    return {
      type: "operator",
      label: labelHint ? `${labelHint}: ${operator.op}` : operator.op,
      expression: text,
      children: [
        parseFormulaLogicExpression(left, "Left"),
        parseFormulaLogicExpression(right, "Right")
      ]
    };
  }

  const ref = resolveFormulaReference(text, logicState.sourceSheet);
  if (ref.isReference) {
    return {
      type: ref.external ? "external" : "reference",
      label: labelHint || "Reference",
      expression: text,
      address: ref.address,
      navigable: !ref.external && Boolean(ref.address)
    };
  }

  return { type: "value", label: labelHint || "Value", expression: text };
}

function isWrappedBySingleParentheses(text) {
  if (!text.startsWith("(") || !text.endsWith(")")) return false;
  let depth = 0;
  let inString = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') {
      if (inString && text[i + 1] === '"') { i++; continue; }
      inString = !inString;
      continue;
    }
    if (inString) continue;
    if (ch === "(") depth++;
    if (ch === ")") depth--;
    if (depth === 0 && i < text.length - 1) return false;
  }
  return depth === 0;
}

function getWholeFunctionCall(text) {
  const match = /^([A-Z_][A-Z0-9._]*)\s*\(/i.exec(text);
  if (!match) return null;
  const openIndex = text.indexOf("(", match[0].length - 1);
  let depth = 0;
  let inString = false;
  for (let i = openIndex; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') {
      if (inString && text[i + 1] === '"') { i++; continue; }
      inString = !inString;
      continue;
    }
    if (inString) continue;
    if (ch === "(") depth++;
    else if (ch === ")") {
      depth--;
      if (depth === 0) {
        if (text.slice(i + 1).trim()) return null;
        return { name: match[1].toUpperCase(), args: text.slice(openIndex + 1, i) };
      }
    }
  }
  return null;
}

function splitTopLevelFormulaArgs(text) {
  const parts = [];
  let start = 0;
  let depth = 0;
  let inString = false;
  let inSheetQuote = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"' && !inSheetQuote) {
      if (inString && text[i + 1] === '"') { i++; continue; }
      inString = !inString;
      continue;
    }
    if (ch === "'" && !inString) {
      if (inSheetQuote && text[i + 1] === "'") { i++; continue; }
      inSheetQuote = !inSheetQuote;
      continue;
    }
    if (inString || inSheetQuote) continue;
    if (ch === "(") depth++;
    else if (ch === ")") depth--;
    else if ((ch === "," || ch === ";") && depth === 0) {
      parts.push(text.slice(start, i).trim());
      start = i + 1;
    }
  }
  parts.push(text.slice(start).trim());
  return parts;
}

function formulaArgumentLabels(name, count) {
  const fixed = {
    IF: ["Condition", "True", "False"],
    IFERROR: ["Value", "Fallback"],
    IFNA: ["Value", "Fallback"],
    XLOOKUP: ["Lookup value", "Lookup array", "Return array", "If not found", "Match mode", "Search mode"],
    XMATCH: ["Lookup value", "Lookup array", "Match mode", "Search mode"],
    MATCH: ["Lookup value", "Lookup array", "Match type"],
    INDEX: ["Array", "Row", "Column", "Area"],
    CHOOSE: ["Index", "Value 1", "Value 2", "Value 3"],
    FILTER: ["Array", "Include", "If empty"],
    SORT: ["Array", "Sort index", "Sort order", "By column"],
    UNIQUE: ["Array", "By column", "Exactly once"]
  };
  if (fixed[name]) return fixed[name].slice(0, count);
  if (["SUMIFS", "COUNTIFS", "AVERAGEIFS", "MAXIFS", "MINIFS"].includes(name)) {
    const labels = [];
    if (name !== "COUNTIFS") labels.push(name === "AVERAGEIFS" ? "Average range" : name === "SUMIFS" ? "Sum range" : "Result range");
    while (labels.length < count) {
      const pair = Math.floor((labels.length - (name === "COUNTIFS" ? 0 : 1)) / 2) + 1;
      labels.push(labels.length % 2 === (name === "COUNTIFS" ? 0 : 1) ? `Criteria range ${pair}` : `Criteria ${pair}`);
    }
    return labels;
  }
  return Array.from({ length: count }, (_, i) => `Arg ${i + 1}`);
}

function findTopLevelFormulaOperator(text) {
  const groups = [
    ["<=", ">=", "<>", "=", "<", ">"],
    ["&"],
    ["+", "-"],
    ["*", "/"],
    ["^"]
  ];

  for (const ops of groups) {
    let depth = 0;
    let inString = false;
    let inSheetQuote = false;
    for (let i = text.length - 1; i >= 0; i--) {
      const ch = text[i];
      if (ch === '"' && !inSheetQuote) { inString = !inString; continue; }
      if (ch === "'" && !inString) { inSheetQuote = !inSheetQuote; continue; }
      if (inString || inSheetQuote) continue;
      if (ch === ")") { depth++; continue; }
      if (ch === "(") { depth--; continue; }
      if (depth !== 0) continue;

      for (const op of ops) {
        const start = i - op.length + 1;
        if (start < 0 || text.slice(start, i + 1) !== op) continue;
        if ((op === "+" || op === "-") && isUnaryFormulaOperator(text, start)) continue;
        if (!text.slice(0, start).trim() || !text.slice(i + 1).trim()) continue;
        return { op, index: start };
      }
    }
  }
  return null;
}

function isUnaryFormulaOperator(text, index) {
  let i = index - 1;
  while (i >= 0 && /\s/.test(text[i])) i--;
  return i < 0 || "(,+-*/^&=<>".includes(text[i]);
}

function resolveFormulaReference(expression, sourceSheet) {
  let text = String(expression || "").trim();
  if (text.startsWith("@")) text = text.slice(1);
  if (/\[[^\]]+\.(?:xlsx|xlsm|xlsb|xls|xlam|csv)\]/i.test(text)) {
    return { isReference: true, external: true, address: "" };
  }
  const refRe = /^(?:(?:'(?:(?:'')|[^'])+'|[A-Za-z_][A-Za-z0-9_. ]*)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?#?$/i;
  if (!refRe.test(text)) return { isReference: false, external: false, address: "" };
  text = text.replace(/#$/, "");
  const parsed = parseQualifiedAddress(text);
  if (parsed.sheetName && parsed.sheetName.includes(":")) return { isReference: true, external: false, address: "" };
  return {
    isReference: true,
    external: false,
    address: parsed.sheetName ? qualifyAddress(parsed.sheetName, parsed.cellAddress) : qualifyAddress(sourceSheet, parsed.cellAddress)
  };
}

function materializeLogicNode(descriptor, parentId, depth) {
  const id = `logic-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
  const node = {
    id,
    parentId,
    depth,
    type: descriptor.type,
    label: descriptor.label,
    expression: descriptor.expression,
    address: descriptor.address || "",
    navigable: Boolean(descriptor.navigable),
    children: [],
    expanded: depth < 2
  };
  logicState.nodes.set(id, node);
  for (const childDescriptor of descriptor.children || []) {
    const child = materializeLogicNode(childDescriptor, id, depth + 1);
    node.children.push(child.id);
  }
  return node;
}

function getVisibleLogicNodes() {
  const out = [];
  const visit = (id) => {
    const node = logicState.nodes.get(id);
    if (!node) return;
    out.push(node);
    if (node.expanded) node.children.forEach(visit);
  };
  if (logicState.rootId) visit(logicState.rootId);
  return out;
}

function renderLogicTree() {
  const tree = $("logic-tree");
  if (!logicState.active) return;
  const visible = getVisibleLogicNodes();
  tree.innerHTML = "";
  $("logic-count").textContent = visible.length;
  $("logic-title").textContent = logicState.sourceAddress;

  visible.forEach((node) => {
    const row = document.createElement("div");
    row.id = `logic-row-${node.id}`;
    row.className = `explorer-row logic-row${node.id === logicState.activeId ? " active" : ""}`;
    row.dataset.nodeId = node.id;
    row.style.setProperty("--depth", String(node.depth));
    row.setAttribute("role", "treeitem");
    row.setAttribute("aria-selected", node.id === logicState.activeId ? "true" : "false");

    const toggle = document.createElement("span");
    toggle.className = `tree-toggle${node.children.length ? " expandable" : ""}`;
    toggle.dataset.nodeId = node.id;
    toggle.dataset.action = "logic-toggle";
    toggle.textContent = node.children.length ? (node.expanded ? "▾" : "▸") : "";

    const main = document.createElement("span");
    main.className = "explorer-row-main";
    const addressLine = document.createElement("div");
    addressLine.className = "explorer-row-address";
    addressLine.textContent = node.label;
    const sub = document.createElement("div");
    sub.className = "explorer-row-sub";
    sub.textContent = node.type === "reference" ? node.address : compactFormulaExpression(node.expression);
    main.append(addressLine, sub);

    const meta = document.createElement("span");
    meta.className = "explorer-row-meta";
    meta.textContent = node.type.toUpperCase();
    row.append(toggle, main, meta);
    tree.appendChild(row);
  });
  tree.setAttribute("aria-activedescendant", logicState.activeId ? `logic-row-${logicState.activeId}` : "");
  document.getElementById(`logic-row-${logicState.activeId}`)?.scrollIntoView({ block: "nearest" });
}

function compactFormulaExpression(text) {
  const s = String(text || "").replace(/\s+/g, " ");
  return s.length > 95 ? `${s.slice(0, 92)}…` : s;
}

async function setLogicActive(nodeId, navigate = true) {
  const node = logicState.nodes.get(nodeId);
  if (!node) return;
  logicState.activeId = nodeId;
  renderLogicTree();
  updateLogicPreview(node);
  if (navigate && node.navigable) await navigateToAddress(node.address);
}

function updateLogicPreview(node) {
  if (!node) return;
  $("logic-preview").classList.remove("hidden");
  const chip = $("logic-preview-kind");
  chip.className = `relation-label${node.type === "reference" ? " same-sheet" : node.type === "external" ? " external" : ""}`;
  chip.textContent = node.type.toUpperCase();
  $("logic-preview-address").textContent = node.navigable ? node.address : logicState.sourceAddress;
  $("logic-preview-expression").textContent = node.expression || node.label;
  $("logic-preview-r1c1").textContent = `A1: ${logicState.formula}; R1C1: ${logicState.formulaR1C1 || "—"}`;
}

async function handleLogicKeydown(event) {
  if (!logicState.active) return;
  const visible = getVisibleLogicNodes();
  const index = visible.findIndex((x) => x.id === logicState.activeId);
  if (index < 0) return;
  if (event.key === "ArrowDown") {
    event.preventDefault();
    if (index < visible.length - 1) await setLogicActive(visible[index + 1].id, true);
  } else if (event.key === "ArrowUp") {
    event.preventDefault();
    if (index > 0) await setLogicActive(visible[index - 1].id, true);
  } else if (event.key === "ArrowRight") {
    event.preventDefault();
    const node = visible[index];
    if (node.children.length) {
      if (!node.expanded) { node.expanded = true; renderLogicTree(); }
      else await setLogicActive(node.children[0], true);
    }
  } else if (event.key === "ArrowLeft") {
    event.preventDefault();
    const node = visible[index];
    if (node.expanded && node.children.length) { node.expanded = false; renderLogicTree(); }
    else if (node.parentId) await setLogicActive(node.parentId, false);
  } else if (event.key === "Enter") {
    event.preventDefault();
    const node = visible[index];
    if (node.navigable) await navigateToAddress(node.address);
  } else if (event.key === "Escape") {
    event.preventDefault();
    await closeFormulaLogic();
  }
}

async function handleLogicClick(event) {
  const toggle = event.target.closest("[data-action='logic-toggle']");
  if (toggle) {
    const node = logicState.nodes.get(toggle.dataset.nodeId);
    if (node?.children.length) { node.expanded = !node.expanded; renderLogicTree(); }
    $("logic-tree").focus({ preventScroll: true });
    return;
  }
  const row = event.target.closest(".logic-row");
  if (!row) return;
  await setLogicActive(row.dataset.nodeId, true);
  $("logic-tree").focus({ preventScroll: true });
}

async function goToLogicSource() {
  if (!logicState.sourceAddress) return;
  await navigateToAddress(logicState.sourceAddress);
  $("logic-tree").focus({ preventScroll: true });
}

async function closeFormulaLogic() {
  if (logicState.sourceAddress) await navigateToAddress(logicState.sourceAddress);
  logicState = createEmptyLogicState();
  showAnalysisMode("dependencies");
  if (explorer.active) $("explorer-tree").focus({ preventScroll: true });
}

async function navigateToAddress(address) {
  if (!address) return;
  try {
    await Excel.run(async (context) => {
      const parsed = parseQualifiedAddress(address);
      const sheet = parsed.sheetName
        ? context.workbook.worksheets.getItem(parsed.sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      sheet.activate();
      sheet.getRange(parsed.cellAddress).select();
      await context.sync();
    });
  } catch (error) {
    const note = $("logic-note");
    if (note) {
      note.textContent = `Could not navigate to ${address}: ${normalizeError(error)}`;
      note.classList.remove("hidden");
    }
  }
}

/* ----------------------------
   Calculation Area Review
----------------------------- */

const CALC_AREA_LIMITS = { maxCells: 2000, maxOutsideDependentsForMapping: 500 };

async function reviewCalculationArea() {
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.13")) {
    showAreaError("Calculation Area review requires ExcelApi 1.13 for dependable input/output tracing.");
    return;
  }

  setWorking(true, "Reviewing area");
  try {
    const result = await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheet = workbook.worksheets.getActiveWorksheet();
      const range = workbook.getSelectedRange();
      sheet.load("name");
      range.load(["address", "rowCount", "columnCount", "formulas"]);
      await context.sync();

      const cellCount = Number(range.rowCount || 1) * Number(range.columnCount || 1);
      if (cellCount > CALC_AREA_LIMITS.maxCells) {
        throw new Error(`Calculation Area review is limited to ${CALC_AREA_LIMITS.maxCells.toLocaleString()} selected cells to protect Excel performance.`);
      }

      const sourceAddress = qualifyAddress(sheet.name, localAddress(range.address));
      const formulaCells = [];
      const origin = parseRangeOrigin(range.address);
      for (let r = 0; r < range.formulas.length; r++) {
        for (let c = 0; c < (range.formulas[r]?.length || 0); c++) {
          if (!isFormulaValue(range.formulas[r][c])) continue;
          formulaCells.push({
            address: qualifyAddress(sheet.name, cellAddressFromOffset(origin.row, origin.col, r, c)),
            formula: String(range.formulas[r][c])
          });
        }
      }

      const externalWorkbooks = extractExternalWorkbookNames(range.formulas || []);

      let precedents = [];
      let aggregateDependents = [];
      try {
        const p = range.getDirectPrecedents();
        p.load("addresses");
        await context.sync();
        precedents = flattenWorkbookAddresses(p.addresses || []);
      } catch (error) {
        if (!isItemNotFound(error)) throw error;
      }
      try {
        const d = range.getDirectDependents();
        d.load("addresses");
        await context.sync();
        aggregateDependents = flattenWorkbookAddresses(d.addresses || []);
      } catch (error) {
        if (!isItemNotFound(error)) throw error;
      }

      const boundaryInputs = precedents.filter((a) => !addressInsideQualifiedRange(a, sourceAddress));
      const outsideDependents = dedupeAddresses(aggregateDependents.filter((a) => !addressInsideQualifiedRange(a, sourceAddress)));
      const outputMap = new Map();
      let exactOutputMapping = outsideDependents.length <= CALC_AREA_LIMITS.maxOutsideDependentsForMapping;

      // Attribute each outside dependent back to the selected cells that directly feed it.
      // This avoids probing every formula inside a large schedule and preserves the professional
      // "inside output -> outside consumer" view users expect from a calculation-area review.
      if (exactOutputMapping && outsideDependents.length) {
        const probes = [];
        try {
          for (const dependentAddress of outsideDependents) {
            const parsed = parseQualifiedAddress(dependentAddress);
            const dependentRange = workbook.worksheets.getItem(parsed.sheetName).getRange(parsed.cellAddress);
            const pres = dependentRange.getDirectPrecedents();
            pres.load("addresses");
            probes.push({ dependentAddress, pres });
          }
          await context.sync();

          for (const probe of probes) {
            const insideSources = flattenWorkbookAddresses(probe.pres.addresses || [])
              .filter((address) => addressInsideQualifiedRange(address, sourceAddress));
            for (const source of insideSources) {
              const key = normalizeAddressKey(source);
              if (!outputMap.has(key)) outputMap.set(key, { address: source, downstream: [] });
              outputMap.get(key).downstream.push(probe.dependentAddress);
            }
          }
        } catch (error) {
          // Direct-dependency APIs can still return ItemNotFound for unusual formula constructs.
          // Inputs/calculations remain valid; use the aggregate outside boundary as a transparent fallback.
          exactOutputMapping = false;
          outputMap.clear();
        }
      }

      const outputs = Array.from(outputMap.values()).map((item) => ({
        address: item.address,
        downstream: dedupeAddresses(item.downstream)
      }));

      return {
        sourceAddress,
        sourceSheet: sheet.name,
        inputs: dedupeAddresses(boundaryInputs),
        calculations: formulaCells,
        outputs,
        downstream: dedupeAddresses(outsideDependents),
        externalWorkbooks,
        exactOutputMapping
      };
    });

    areaState = createEmptyAreaState();
    areaState.active = true;
    areaState.sourceAddress = result.sourceAddress;
    areaState.sourceSheet = result.sourceSheet;
    areaState.exactOutputMapping = result.exactOutputMapping;

    const items = [];
    for (const address of result.inputs) items.push({ group: "Inputs", address, detail: "Precedent outside selected block" });
    for (const workbookName of result.externalWorkbooks || []) items.push({
      group: "External",
      address: `[${workbookName}]`,
      detail: "External workbook input boundary",
      nonNavigable: true
    });
    for (const item of result.calculations) items.push({ group: "Calculations", address: item.address, detail: compactFormulaExpression(item.formula) });
    if (result.exactOutputMapping) {
      for (const item of result.outputs) items.push({ group: "Outputs", address: item.address, detail: `Feeds ${item.downstream.slice(0, 3).join(", ")}${item.downstream.length > 3 ? "…" : ""}` });
    } else {
      for (const address of result.downstream) items.push({ group: "Outputs", address, detail: "Downstream dependent outside block", downstreamOnly: true });
    }
    areaState.items = items;
    areaState.activeIndex = 0;

    showAnalysisMode("area");
    $("area-title").textContent = result.sourceAddress;
    const areaNotices = [];
    if (!result.exactOutputMapping) {
      areaNotices.push(`The selected block has ${result.downstream.length.toLocaleString()} outside dependent cells, above the ${CALC_AREA_LIMITS.maxOutsideDependentsForMapping.toLocaleString()} attribution limit or Excel could not resolve one of the dependency probes. Inputs and calculations are complete; outputs show the downstream boundary instead. Narrow the selection for exact source-cell attribution.`);
    }
    if (result.externalWorkbooks?.length) {
      areaNotices.push(`External workbook boundary. Cross-workbook tracing unavailable.`);
    }
    if (areaNotices.length) {
      $("area-note").textContent = areaNotices.join(" ");
      $("area-note").classList.remove("hidden");
    } else {
      $("area-note").classList.add("hidden");
    }
    renderAreaResults();
    $("area-results").focus({ preventScroll: true });
  } catch (error) {
    showAreaError(normalizeError(error));
  } finally {
    setWorking(false);
  }
}

function dedupeAddresses(addresses) {
  const seen = new Set();
  const out = [];
  for (const address of addresses || []) {
    const key = normalizeAddressKey(address);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    out.push(address);
  }
  return out.sort(compareExplorerAddressesByWorksheet);
}

function addressInsideQualifiedRange(address, containerAddress) {
  const a = parseQualifiedAddress(address);
  const c = parseQualifiedAddress(containerAddress);
  if ((a.sheetName || "").toLowerCase() !== (c.sheetName || "").toLowerCase()) return false;
  const aCell = parseA1Cell(a.cellAddress.split(":")[0]);
  const bounds = parseA1Bounds(c.cellAddress);
  if (!aCell || !bounds) return false;
  return aCell.row >= bounds.top && aCell.row <= bounds.bottom && aCell.col >= bounds.left && aCell.col <= bounds.right;
}

function parseA1Cell(address) {
  const match = /^\$?([A-Z]{1,3})\$?(\d+)$/i.exec(String(address || ""));
  if (!match) return null;
  return { row: Number(match[2]), col: columnLettersToNumber(match[1]) };
}

function parseA1Bounds(address) {
  const parts = String(address || "").replace(/\$/g, "").split(":");
  const first = parseA1Cell(parts[0]);
  const last = parseA1Cell(parts[1] || parts[0]);
  if (!first || !last) return null;
  return {
    top: Math.min(first.row, last.row), bottom: Math.max(first.row, last.row),
    left: Math.min(first.col, last.col), right: Math.max(first.col, last.col)
  };
}

function renderAreaResults() {
  const container = $("area-results");
  container.innerHTML = "";
  $("area-count").textContent = areaState.items.length;
  if (!areaState.items.length) {
    container.innerHTML = '<div class="empty-state">No boundary items</div>';
    return;
  }

  let lastGroup = "";
  areaState.items.forEach((item, index) => {
    if (item.group !== lastGroup) {
      const header = document.createElement("div");
      header.className = "explorer-group-header";
      header.textContent = item.group;
      container.appendChild(header);
      lastGroup = item.group;
    }
    const row = document.createElement("div");
    row.className = `area-row${index === areaState.activeIndex ? " active" : ""}`;
    row.dataset.index = String(index);
    row.setAttribute("role", "option");
    row.setAttribute("aria-selected", index === areaState.activeIndex ? "true" : "false");
    row.innerHTML = `<span class="area-row-main"><strong>${escapeHtml(item.address)}</strong><small>${escapeHtml(item.detail)}</small></span><span class="jump"></span>`;
    container.appendChild(row);
  });
  container.querySelector(`.area-row[data-index="${areaState.activeIndex}"]`)?.scrollIntoView({ block: "nearest" });
}

async function setAreaActive(index, navigate = true) {
  if (index < 0 || index >= areaState.items.length) return;
  areaState.activeIndex = index;
  renderAreaResults();
  if (navigate && !areaState.items[index].nonNavigable) await navigateToAddress(areaState.items[index].address);
}

async function handleAreaKeydown(event) {
  if (!areaState.active) return;
  if (event.key === "ArrowDown") {
    event.preventDefault();
    await setAreaActive(Math.min(areaState.items.length - 1, areaState.activeIndex + 1), true);
  } else if (event.key === "ArrowUp") {
    event.preventDefault();
    await setAreaActive(Math.max(0, areaState.activeIndex - 1), true);
  } else if (event.key === "Enter") {
    event.preventDefault();
    const item = areaState.items[areaState.activeIndex];
    if (item && !item.nonNavigable) await navigateToAddress(item.address);
  } else if (event.key === "Escape") {
    event.preventDefault();
    await closeCalculationArea();
  }
}

async function handleAreaClick(event) {
  const row = event.target.closest(".area-row");
  if (!row) return;
  await setAreaActive(Number(row.dataset.index), true);
  $("area-results").focus({ preventScroll: true });
}

async function goToAreaSource() {
  if (areaState.sourceAddress) await navigateToAddress(areaState.sourceAddress);
  $("area-results").focus({ preventScroll: true });
}

async function closeCalculationArea() {
  if (areaState.sourceAddress) await navigateToAddress(areaState.sourceAddress);
  areaState = createEmptyAreaState();
  showAnalysisMode("dependencies");
  if (explorer.active) $("explorer-tree").focus({ preventScroll: true });
}

function showAreaError(message) {
  areaState = createEmptyAreaState();
  showAnalysisMode("area");
  $("area-note").textContent = message;
  $("area-note").classList.remove("hidden");
  $("area-results").innerHTML = "";
  $("area-count").textContent = "0";
}

/* ----------------------------
   Professional model audit
----------------------------- */

const AUDIT_SEVERITY_ORDER = { critical: 0, high: 1, medium: 2, low: 3 };
const EXCEL_ERROR_RE = /^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|SPILL!|CALC!|FIELD!|GETTING_DATA|CONNECT!|BLOCKED!|UNKNOWN!|BUSY!)/i;
const EXTERNAL_BOOK_RE = /\[[^\]]+\.(?:xlsx|xlsm|xlsb|xls|xlam|csv)\]/i;
const WHOLE_COLUMN_RE = /(?:^|[^A-Z0-9_])\$?[A-Z]{1,3}:\$?[A-Z]{1,3}(?=$|[^A-Z0-9_])/i;
const VOLATILE_FUNCTIONS = ["OFFSET", "INDIRECT", "TODAY", "NOW", "RAND", "RANDBETWEEN", "CELL", "INFO"];
const DYNAMIC_ARRAY_FUNCTIONS = [
  "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY", "TAKE", "DROP",
  "CHOOSECOLS", "CHOOSEROWS", "HSTACK", "VSTACK", "TOROW", "TOCOL", "WRAPROWS", "WRAPCOLS", "EXPAND"
];
const MAX_AUDIT_SPILL_PROBES = 500;

async function auditScope(scope) {
  const label =
    scope === "workbook" ? "entire workbook" :
    scope === "sheet" ? "current sheet" :
    "selected range";

  resetAuditFilters();
  setAuditSummary(`Auditing ${label}…`, "");
  setWorking(true, "Auditing");

  try {
    const result = await Excel.run(async (context) => {
      const workbook = context.workbook;
      const supportsNamedFormula = Office.context.requirements.isSetSupported("ExcelApi", "1.7");
      const sheets = workbook.worksheets;
      sheets.load("items/name,items/visibility");

      let jobs = [];
      let names = null;

      if (scope === "selection") {
        const sheet = workbook.worksheets.getActiveWorksheet();
        sheet.load("name,visibility");
        const range = workbook.getSelectedRange();
        range.load(["formulas", "formulasR1C1", "values", "numberFormat", "rowCount", "columnCount", "address"]);
        jobs.push({ sheet, range });
      } else if (scope === "sheet") {
        const sheet = workbook.worksheets.getActiveWorksheet();
        sheet.load("name,visibility");
        const range = sheet.getUsedRangeOrNullObject(true);
        range.load(["isNullObject", "formulas", "formulasR1C1", "values", "numberFormat", "rowCount", "columnCount", "address"]);
        jobs.push({ sheet, range });
      } else {
        await context.sync();
        for (const sheet of sheets.items) {
          const range = sheet.getUsedRangeOrNullObject(true);
          range.load(["isNullObject", "formulas", "formulasR1C1", "values", "numberFormat", "rowCount", "columnCount", "address"]);
          jobs.push({ sheet, range });
        }

        names = workbook.names;
        if (supportsNamedFormula) {
          names.load("items/name,items/type,items/value,items/visible,items/formula");
        } else {
          names.load("items/name,items/type,items/value,items/visible");
        }
      }

      await context.sync();

      const calculation = await getCalculationSnapshot(context);
      const spillRangesByJob = new Map();
      const supportsSpillAudit = Office.context.requirements.isSetSupported("ExcelApi", "1.12");
      if (supportsSpillAudit) {
        const spillProbes = [];
        let spillProbeCount = 0;
        for (const job of jobs) {
          if (job.range.isNullObject || spillProbeCount >= MAX_AUDIT_SPILL_PROBES) continue;
          const jobCells = Number(job.range.rowCount || 0) * Number(job.range.columnCount || 0);
          if (jobCells > 250000) continue;
          const formulas = job.range.formulas || [];
          for (let r = 0; r < formulas.length && spillProbeCount < MAX_AUDIT_SPILL_PROBES; r++) {
            for (let c = 0; c < (formulas[r]?.length || 0) && spillProbeCount < MAX_AUDIT_SPILL_PROBES; c++) {
              const formula = formulas[r]?.[c];
              if (!isFormulaValue(formula) || !looksLikeDynamicArrayFormula(formula)) continue;
              const parent = job.range.getCell(r, c);
              const spilling = parent.getSpillingToRangeOrNullObject();
              spilling.load("isNullObject,address");
              spillProbes.push({ job, r, c, spilling });
              spillProbeCount += 1;
            }
          }
        }
        if (spillProbes.length) {
          await context.sync();
          for (const probe of spillProbes) {
            if (probe.spilling.isNullObject) continue;
            const parsed = parseQualifiedAddress(probe.spilling.address);
            const bounds = parseA1Bounds(parsed.cellAddress);
            if (!bounds) continue;
            const origin = parseRangeOrigin(probe.job.range.address);
            const parentRow = origin.row + probe.r;
            const parentCol = origin.col + probe.c;
            const ranges = spillRangesByJob.get(probe.job) || [];
            ranges.push({
              sheetName: probe.job.sheet.name,
              top: bounds.top,
              bottom: bounds.bottom,
              left: bounds.left,
              right: bounds.right,
              parentRow,
              parentCol
            });
            spillRangesByJob.set(probe.job, ranges);
          }
        }
      }

      const findings = analyzeCalculationControls(calculation);
      const meta = {
        scope,
        scopeLabel: label,
        scannedCells: 0,
        formulaCells: 0,
        hardcodeCells: 0,
        blankCells: 0,
        scannedSheets: 0,
        hiddenSheets: 0,
        namedItems: 0,
        skippedLargeSheets: [],
        calculation,
        spillAware: supportsSpillAudit,
        spillProbeLimitReached: supportsSpillAudit && Array.from(spillRangesByJob.values()).reduce((n, ranges) => n + ranges.length, 0) >= MAX_AUDIT_SPILL_PROBES,
        stale: false,
        auditTime: new Date().toISOString()
      };

      const MAX_AUDIT_CELLS_PER_SHEET = 250000;

      for (const job of jobs) {
        if (job.range.isNullObject) continue;

        const cells = (job.range.rowCount || 0) * (job.range.columnCount || 0);
        if (scope === "workbook" && cells > MAX_AUDIT_CELLS_PER_SHEET) {
          meta.skippedLargeSheets.push(job.sheet.name);
          findings.push(makeFinding({
            severity: "medium",
            category: "Performance & complexity",
            check: "SCAN_LIMIT",
            title: "Very large sheet not fully scanned",
            source: job.sheet.name,
            detail: `${job.sheet.name} contains approximately ${cells.toLocaleString()} used cells, above the full-workbook audit safety limit.`,
            recommendation: "Run a focused current-sheet or selected-range audit on the model area that matters."
          }));
          continue;
        }

        const analyzed = analyzeAuditRange(
          job.sheet.name,
          job.range.address,
          job.range.formulas || [],
          job.range.formulasR1C1 || [],
          job.range.values || [],
          job.range.numberFormat || [],
          spillRangesByJob.get(job) || []
        );

        findings.push(...analyzed.findings);
        meta.scannedCells += analyzed.stats.cells;
        meta.formulaCells += analyzed.stats.formulas;
        meta.hardcodeCells += analyzed.stats.hardcodes;
        meta.blankCells += analyzed.stats.blanks;
        meta.scannedSheets += 1;
      }

      if (scope === "workbook") {
        meta.hiddenSheets = sheets.items.filter((sheet) => String(sheet.visibility).toLowerCase() !== "visible").length;
        if (names) {
          meta.namedItems = names.items.length;
          findings.push(...analyzeNamedItems(names.items, supportsNamedFormula));
        }
      }

      return { findings: dedupeAndSortFindings(findings), meta };
    });

    lastAuditFindings = result.findings;
    lastAuditMeta = result.meta;

    updateAuditDashboard();
    renderAuditFindings();

    const counts = countAuditSeverities(lastAuditFindings);
    const status = deriveAuditStatus(counts);

    setAuditSummary(
      `${status.label}. Scanned ${result.meta.formulaCells.toLocaleString()} formulas across ${result.meta.scannedSheets} sheet${result.meta.scannedSheets === 1 ? "" : "s"}${result.meta.skippedLargeSheets.length ? `; ${result.meta.skippedLargeSheets.length} very large sheet${result.meta.skippedLargeSheets.length === 1 ? "" : "s"} require focused review` : ""}.`,
      status.className === "clean" ? "success" : ""
    );
  } catch (error) {
    setAuditSummary(`Could not audit ${label}: ${normalizeError(error)}`, "error");
    setAuditStatus("Audit failed", "critical");
  } finally {
    setWorking(false);
  }
}

function analyzeAuditRange(sheetName, rangeAddress, formulas, formulasR1C1, values, numberFormats, spillRanges = []) {
  const findings = [];
  const rows = formulas.length;
  const cols = rows ? (formulas[0]?.length || 0) : 0;
  const origin = parseRangeOrigin(rangeAddress);
  const stats = { cells: rows * cols, formulas: 0, hardcodes: 0, blanks: 0 };

  const formulaAt = (r, c) => {
    if (r < 0 || c < 0 || r >= rows || c >= cols) return false;
    return isFormulaValue(formulas[r]?.[c]);
  };

  const r1c1At = (r, c) => (r >= 0 && c >= 0 && r < rows && c < cols) ? formulasR1C1[r]?.[c] : null;
  const nfAt = (r, c) => (r >= 0 && c >= 0 && r < rows && c < cols) ? numberFormats[r]?.[c] : null;

  const addCellFinding = (r, c, data) => {
    findings.push(makeFinding({
      ...data,
      sheet: sheetName,
      address: qualifyAddress(sheetName, cellAddressFromOffset(origin.row, origin.col, r, c)),
      source: sheetName
    }));
  };

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const formula = formulas[r]?.[c];
      const r1c1 = formulasR1C1[r]?.[c];
      const value = values[r]?.[c];
      const isFormula = isFormulaValue(formula);
      const absoluteRow = origin.row + r;
      const absoluteCol = origin.col + c;
      const isSpillOutput = isSpillOutputCell(spillRanges, absoluteRow, absoluteCol);
      const isBlank = !isFormula && !isSpillOutput && (value === "" || value === null || typeof value === "undefined");
      const isHardcode = !isFormula && !isSpillOutput && (typeof value === "number" || typeof value === "boolean");

      if (isSpillOutput) continue;
      if (isFormula) stats.formulas += 1;
      else if (isHardcode) stats.hardcodes += 1;
      else if (isBlank) stats.blanks += 1;

      const leftFormula = formulaAt(r, c - 1);
      const rightFormula = formulaAt(r, c + 1);
      const aboveFormula = formulaAt(r - 1, c);
      const belowFormula = formulaAt(r + 1, c);

      const leftPattern = r1c1At(r, c - 1);
      const rightPattern = r1c1At(r, c + 1);
      const abovePattern = r1c1At(r - 1, c);
      const belowPattern = r1c1At(r + 1, c);

      const strongHorizontalSeries = leftFormula && rightFormula && leftPattern === rightPattern;
      const strongVerticalSeries = aboveFormula && belowFormula && abovePattern === belowPattern;

      if (isHardcode && (strongHorizontalSeries || strongVerticalSeries)) {
        addCellFinding(r, c, {
          severity: "high",
          category: "Series consistency",
          check: "HARD_CODE_BREAK",
          title: "Hardcode interrupts a copied formula series",
          detail: `Value ${String(value)} sits between matching formula patterns ${strongHorizontalSeries ? "horizontally" : "vertically"}.`,
          recommendation: "Confirm the override is intentional; otherwise restore the copied formula."
        });
      }

      if (isBlank && (strongHorizontalSeries || strongVerticalSeries)) {
        addCellFinding(r, c, {
          severity: "high",
          category: "Series consistency",
          check: "BLANK_BREAK",
          title: "Blank interrupts a copied formula series",
          detail: `A blank cell sits between matching formulas ${strongHorizontalSeries ? "horizontally" : "vertically"}.`,
          recommendation: "Confirm the blank is intentional; otherwise restore the missing formula."
        });
      }

      if (!isFormula) continue;

      const displayed = String(value ?? "");
      if (EXCEL_ERROR_RE.test(displayed)) {
        addCellFinding(r, c, {
          severity: "critical",
          category: "Formula integrity",
          check: "FORMULA_ERROR",
          title: "Formula returns an Excel error",
          detail: `${displayed}; ${formula}`,
          recommendation: "Trace precedents and resolve the source error before relying on downstream outputs."
        });
      }

      if (String(formula).toUpperCase().includes("#REF!")) {
        addCellFinding(r, c, {
          severity: "critical",
          category: "Formula integrity",
          check: "BROKEN_REF",
          title: "Formula contains a broken #REF! reference",
          detail: String(formula),
          recommendation: "Repair the deleted or moved range reference."
        });
      }

      if (hasDirectSelfReference(String(r1c1 || ""))) {
        addCellFinding(r, c, {
          severity: "critical",
          category: "Formula integrity",
          check: "DIRECT_SELF_REFERENCE",
          title: "Direct self-reference / circular-reference risk",
          detail: String(formula),
          recommendation: "Verify whether iterative calculation is intentional; otherwise remove the self-reference."
        });
      }

      const consistency = strongestFormulaConsistencyEvidence(formulas, formulasR1C1, r, c);
      if (consistency && String(r1c1 || "") !== consistency.expectedPattern) {
        addCellFinding(r, c, {
          severity: "high",
          category: "Formula integrity",
          check: "FORMULA_OUTLIER",
          title: "Formula inconsistency in copied series",
          detail: `${String(formula)}; expected R1C1 pattern from ${consistency.supportCount} surrounding ${consistency.axis} cells: ${consistency.expectedPattern}`,
          recommendation: "Compare with the surrounding copied formulas and confirm whether the reference or anchoring change is intentional."
        });
      }

      if (EXTERNAL_BOOK_RE.test(removeQuotedText(String(formula)))) {
        addCellFinding(r, c, {
          severity: "high",
          category: "Links & names",
          check: "EXTERNAL_WORKBOOK_LINK",
          title: "External workbook link",
          detail: String(formula),
          recommendation: "Confirm the external source is controlled, current and expected to remain available."
        });
      }

      const currentFormat = normalizeNumberFormat(nfAt(r, c));
      const leftFormat = normalizeNumberFormat(nfAt(r, c - 1));
      const rightFormat = normalizeNumberFormat(nfAt(r, c + 1));
      const aboveFormat = normalizeNumberFormat(nfAt(r - 1, c));
      const belowFormat = normalizeNumberFormat(nfAt(r + 1, c));

      const horizontalFormatOutlier =
        strongHorizontalSeries && leftFormat && leftFormat === rightFormat && currentFormat && currentFormat !== leftFormat;
      const verticalFormatOutlier =
        strongVerticalSeries && aboveFormat && aboveFormat === belowFormat && currentFormat && currentFormat !== aboveFormat;

      if (horizontalFormatOutlier || verticalFormatOutlier) {
        addCellFinding(r, c, {
          severity: "medium",
          category: "Series consistency",
          check: "NUMBER_FORMAT_OUTLIER",
          title: "Number format differs inside a copied formula series",
          detail: `Cell format: ${currentFormat}; surrounding format: ${horizontalFormatOutlier ? leftFormat : aboveFormat}`,
          recommendation: "Confirm the display-unit or number-format exception is intentional."
        });
      }

      const formulaOutsideStrings = removeQuotedText(String(formula));
      const volatile = VOLATILE_FUNCTIONS.filter((fn) => new RegExp(`\\b${fn}\\s*\\(`, "i").test(formulaOutsideStrings));
      if (volatile.length) {
        addCellFinding(r, c, {
          severity: "medium",
          category: "Performance & complexity",
          check: "VOLATILE_FUNCTION",
          title: "Volatile / indirect formula",
          detail: `${volatile.join(", ")}; ${formula}`,
          recommendation: "Confirm the volatility is necessary because these functions can make large models slower or harder to audit."
        });
      }

      if (WHOLE_COLUMN_RE.test(formulaOutsideStrings)) {
        addCellFinding(r, c, {
          severity: "medium",
          category: "Performance & complexity",
          check: "WHOLE_COLUMN_REFERENCE",
          title: "Whole-column reference",
          detail: String(formula),
          recommendation: "Consider a bounded range where practical, especially in calculation-heavy models."
        });
      }

      const functionCount = countFormulaFunctions(formulaOutsideStrings);
      if (String(formula).length > 350 || functionCount > 15) {
        addCellFinding(r, c, {
          severity: "low",
          category: "Performance & complexity",
          check: "COMPLEX_FORMULA",
          title: "Unusually long or complex formula",
          detail: `${String(formula).length} characters; approximately ${functionCount} function calls`,
          recommendation: "Consider breaking the calculation into transparent helper rows if this is a core model calculation."
        });
      }
    }
  }

  return { findings, stats };
}

function looksLikeDynamicArrayFormula(formula) {
  const clean = removeQuotedText(String(formula || ""));
  return DYNAMIC_ARRAY_FUNCTIONS.some((fn) => new RegExp(`\\b${fn}\\s*\\(`, "i").test(clean));
}

function isSpillOutputCell(spillRanges, row, col) {
  for (const range of spillRanges || []) {
    if (row < range.top || row > range.bottom || col < range.left || col > range.right) continue;
    if (row === range.parentRow && col === range.parentCol) return false;
    return true;
  }
  return false;
}

function analyzeCalculationControls(info) {
  const findings = [];
  const mode = String(info?.mode || "");
  const state = String(info?.state || "");
  const modeKey = mode.toLowerCase().replace(/[^a-z]/g, "");
  const stateKey = state.toLowerCase();

  if (modeKey === "manual") {
    findings.push(makeFinding({
      severity: "medium",
      category: "Workbook controls",
      check: "MANUAL_CALC_MODE",
      title: "Excel is in Manual calculation mode",
      source: "Excel calculation settings",
      detail: "Displayed values can be stale until Excel is recalculated.",
      recommendation: "Confirm Manual mode is intentional. Recalculate before reviewing value-based outputs."
    }));
  } else if (modeKey.includes("automaticexcept") || modeKey.includes("excepttables")) {
    findings.push(makeFinding({
      severity: "low",
      category: "Workbook controls",
      check: "AUTO_EXCEPT_TABLES",
      title: "Calculation mode is Automatic Except Data Tables",
      source: "Excel calculation settings",
      detail: `Calculation mode: ${mode}`,
      recommendation: "Confirm this performance setting is intentional and that data-table outputs are refreshed before review."
    }));
  }

  if (stateKey && !stateKey.includes("done")) {
    findings.push(makeFinding({
      severity: "medium",
      category: "Workbook controls",
      check: "CALCULATION_NOT_DONE",
      title: "Excel calculation is not complete",
      source: "Excel calculation state",
      detail: `Calculation state at audit start: ${state}`,
      recommendation: "Recalculate or allow Excel to finish calculating before relying on value-sensitive findings."
    }));
  }

  if (info?.iterative?.enabled) {
    findings.push(makeFinding({
      severity: "medium",
      category: "Workbook controls",
      check: "ITERATIVE_CALCULATION",
      title: "Iterative calculation is enabled",
      source: "Excel calculation settings",
      detail: `Maximum iterations: ${info.iterative.maxIteration ?? "—"}; maximum change: ${info.iterative.maxChange ?? "—"}`,
      recommendation: "Confirm intentional circularity is documented and the iteration/tolerance settings are appropriate for the model."
    }));
  }

  return findings;
}

function strongestFormulaConsistencyEvidence(formulas, formulasR1C1, row, col) {
  const horizontal = formulaConsistencyEvidence(formulas, formulasR1C1, row, col, "row");
  const vertical = formulaConsistencyEvidence(formulas, formulasR1C1, row, col, "column");
  if (!horizontal) return vertical;
  if (!vertical) return horizontal;
  return horizontal.supportCount >= vertical.supportCount ? horizontal : vertical;
}

function formulaConsistencyEvidence(formulas, formulasR1C1, row, col, axis) {
  const rows = formulas.length;
  const cols = rows ? (formulas[0]?.length || 0) : 0;
  const sample = [];
  let hasBefore = false;
  let hasAfter = false;

  for (let distance = 1; distance <= 3; distance++) {
    for (const direction of [-1, 1]) {
      const rr = axis === "row" ? row : row + direction * distance;
      const cc = axis === "row" ? col + direction * distance : col;
      if (rr < 0 || cc < 0 || rr >= rows || cc >= cols) continue;
      const formula = formulas[rr]?.[cc];
      if (!isFormulaValue(formula)) continue;
      const pattern = String(formulasR1C1[rr]?.[cc] || "");
      if (!pattern) continue;
      sample.push({ pattern, direction, distance });
      if (direction < 0) hasBefore = true;
      if (direction > 0) hasAfter = true;
    }
  }

  if (!hasBefore || !hasAfter || sample.length < 2) return null;
  const counts = new Map();
  for (const item of sample) counts.set(item.pattern, (counts.get(item.pattern) || 0) + 1);
  const ranked = Array.from(counts.entries()).sort((a, b) => b[1] - a[1]);
  if (!ranked.length) return null;
  const [expectedPattern, supportCount] = ranked[0];
  const onBothSides = sample.some((item) => item.direction < 0 && item.pattern === expectedPattern)
    && sample.some((item) => item.direction > 0 && item.pattern === expectedPattern);
  const share = supportCount / sample.length;
  if (!onBothSides || supportCount < 2 || share < 0.67) return null;
  return { axis, expectedPattern, supportCount, sampleCount: sample.length };
}

function analyzeNamedItems(items, supportsFormula) {
  const findings = [];

  for (const item of items || []) {
    const value = String(item.value ?? "");
    const formula = supportsFormula ? String(item.formula ?? "") : "";
    const type = String(item.type ?? "");

    if (type.toLowerCase() === "error" || EXCEL_ERROR_RE.test(value) || formula.toUpperCase().includes("#REF!")) {
      findings.push(makeFinding({
        severity: "critical",
        category: "Links & names",
        check: "INVALID_DEFINED_NAME",
        title: "Invalid defined name",
        source: `Defined name: ${item.name}`,
        detail: formula || value || type,
        recommendation: "Repair or delete the defined name if it is no longer required."
      }));
    }

    if (supportsFormula && EXTERNAL_BOOK_RE.test(removeQuotedText(formula))) {
      findings.push(makeFinding({
        severity: "high",
        category: "Links & names",
        check: "EXTERNAL_NAME_LINK",
        title: "Defined name references an external workbook",
        source: `Defined name: ${item.name}`,
        detail: formula,
        recommendation: "Confirm the external dependency is intentional and controlled."
      }));
    }
  }

  return findings;
}

function makeFinding(data) {
  return {
    severity: data.severity || "low",
    category: data.category || "Other",
    check: data.check || "REVIEW",
    title: data.title || "Review item",
    address: data.address || "",
    sheet: data.sheet || "",
    source: data.source || data.address || "",
    detail: data.detail || "",
    recommendation: data.recommendation || ""
  };
}

function dedupeAndSortFindings(findings) {
  const seen = new Set();
  const out = [];

  for (const finding of findings) {
    const key = [finding.check, finding.address, finding.source, finding.detail].join("|");
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(finding);
  }

  out.sort((a, b) => {
    const severity = AUDIT_SEVERITY_ORDER[a.severity] - AUDIT_SEVERITY_ORDER[b.severity];
    if (severity) return severity;
    const category = a.category.localeCompare(b.category);
    if (category) return category;
    return String(a.address || a.source).localeCompare(String(b.address || b.source));
  });

  return out;
}

function isFormulaValue(value) {
  return typeof value === "string" && value.startsWith("=");
}

function removeQuotedText(formula) {
  let result = "";
  let inString = false;

  for (let i = 0; i < formula.length; i++) {
    const ch = formula[i];
    if (ch === '"') {
      if (inString && formula[i + 1] === '"') {
        i++;
        continue;
      }
      inString = !inString;
      continue;
    }
    if (!inString) result += ch;
  }

  return result;
}

function hasDirectSelfReference(formulaR1C1) {
  if (!formulaR1C1) return false;
  return /(^|[^A-Z0-9_!])R(?:\[0\])?C(?:\[0\])?(?!\[|\d|[A-Z0-9_])/i.test(removeQuotedText(formulaR1C1));
}

function normalizeNumberFormat(format) {
  if (Array.isArray(format)) return String(format[0] ?? "");
  return String(format ?? "");
}

function countFormulaFunctions(formula) {
  const matches = String(formula).match(/\b[A-Z][A-Z0-9._]*\s*\(/gi);
  return matches ? matches.length : 0;
}

function parseRangeOrigin(address) {
  const local = localAddress(address).split(":")[0].replace(/\$/g, "");
  const match = /^([A-Z]+)(\d+)$/i.exec(local);
  if (!match) return { row: 1, col: 1 };
  return { row: Number(match[2]), col: columnLettersToNumber(match[1]) };
}

function columnLettersToNumber(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function columnNumberToLetters(n) {
  let s = "";
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

function cellAddressFromOffset(baseRow, baseCol, rowOffset, colOffset) {
  return `${columnNumberToLetters(baseCol + colOffset)}${baseRow + rowOffset}`;
}

function countAuditSeverities(findings) {
  return (findings || []).reduce((acc, finding) => {
    acc[finding.severity] = (acc[finding.severity] || 0) + 1;
    return acc;
  }, { critical: 0, high: 0, medium: 0, low: 0 });
}

function deriveAuditStatus(counts) {
  if (counts.critical > 0) return { label: "Critical issues found", className: "critical" };
  if (counts.high > 0) return { label: "High-priority review required", className: "high" };
  if (counts.medium > 0) return { label: "Review items found", className: "review" };
  if (counts.low > 0) return { label: "Minor review items found", className: "review" };
  return { label: "No material flags from included checks", className: "clean" };
}

function setAuditStatus(text, mode) {
  const badge = $("audit-status-badge");
  badge.textContent = text;
  badge.className = `audit-status ${mode || "neutral"}`;
}

function updateAuditDashboard() {
  const counts = countAuditSeverities(lastAuditFindings);
  const status = deriveAuditStatus(counts);

  $("kpi-critical").textContent = counts.critical;
  $("kpi-high").textContent = counts.high;
  $("kpi-review").textContent = counts.medium + counts.low;
  $("kpi-formulas").textContent = lastAuditMeta ? lastAuditMeta.formulaCells.toLocaleString() : "—";
  setAuditStatus(status.label, status.className);

  if (!lastAuditMeta) {
    $("audit-meta").textContent = "Run an audit to view findings.";
    return;
  }

  const metaParts = [
    `${lastAuditMeta.scannedSheets} sheet${lastAuditMeta.scannedSheets === 1 ? "" : "s"} scanned`,
    `${lastAuditMeta.scannedCells.toLocaleString()} cells reviewed`,
    `${lastAuditMeta.hardcodeCells.toLocaleString()} numeric/boolean hardcodes`
  ];

  if (lastAuditMeta.calculation?.mode) metaParts.push(`Calculation: ${lastAuditMeta.calculation.mode}`);
  if (lastAuditMeta.stale) metaParts.push("Workbook changed since audit");

  if (lastAuditMeta.scope === "workbook") {
    metaParts.push(`${lastAuditMeta.namedItems} defined names`);
    if (lastAuditMeta.hiddenSheets) metaParts.push(`${lastAuditMeta.hiddenSheets} hidden sheet${lastAuditMeta.hiddenSheets === 1 ? "" : "s"}`);
  }

  $("audit-meta").textContent = metaParts.join("; ");
}

function resetAuditFilters() {
  activeAuditSeverity = "all";
  activeAuditCategory = "all";
  document.querySelectorAll("#severity-filter .filter-tab").forEach((button) => {
    button.classList.toggle("active", button.dataset.severity === "all");
  });
  if ($("audit-category-filter")) $("audit-category-filter").value = "all";
}

function getFilteredAuditFindings() {
  return lastAuditFindings.filter((finding) => {
    const severityOk = activeAuditSeverity === "all" || finding.severity === activeAuditSeverity;
    const categoryOk = activeAuditCategory === "all" || finding.category === activeAuditCategory;
    return severityOk && categoryOk;
  });
}

function renderAuditFindings() {
  const container = $("audit-results");
  const filtered = getFilteredAuditFindings();

  $("copy-audit").disabled = lastAuditFindings.length === 0;
  $("export-audit").disabled = lastAuditFindings.length === 0 || !lastAuditMeta;

  if (!lastAuditMeta) {
    $("audit-count").textContent = "No audit run.";
    container.className = "audit-results empty";
    container.innerHTML = '<div class="empty-state">No findings</div>';
    return;
  }

  $("audit-count").textContent =
    activeAuditSeverity === "all" && activeAuditCategory === "all"
      ? `${lastAuditFindings.length} finding${lastAuditFindings.length === 1 ? "" : "s"}`
      : `${filtered.length} of ${lastAuditFindings.length} findings shown`;

  if (!filtered.length) {
    container.className = "audit-results empty";
    container.innerHTML = '<div class="audit-empty-note">No findings match the current filters.</div>';
    return;
  }

  container.className = "audit-results";
  container.innerHTML = "";

  filtered.forEach((finding) => {
    const item = document.createElement(finding.address ? "button" : "div");
    item.className = `audit-item${finding.address ? " clickable" : ""}`;
    if (finding.address) {
      item.type = "button";
      item.addEventListener("click", () => jumpToAddress(finding.address));
    }

    const severity = document.createElement("span");
    severity.className = `severity-badge ${finding.severity}`;
    severity.textContent = finding.severity;

    const middle = document.createElement("span");
    const category = document.createElement("div");
    category.className = "audit-category";
    category.textContent = finding.category;

    const title = document.createElement("div");
    title.className = "audit-title";
    title.textContent = finding.title;

    const source = document.createElement("div");
    source.className = "audit-address";
    source.textContent = finding.address || finding.source || "Workbook";

    const detail = document.createElement("div");
    detail.className = "audit-detail";
    detail.textContent = finding.detail;

    const recommendation = document.createElement("div");
    recommendation.className = "audit-recommendation";
    recommendation.textContent = finding.recommendation;

    middle.append(category, title, source, detail, recommendation);

    const jump = document.createElement("span");
    jump.className = "jump";
    jump.textContent = "";

    item.append(severity, middle, jump);
    container.appendChild(item);
  });
}

function setAuditSummary(text, mode = "") {
  const el = $("audit-summary");
  el.textContent = text;
  el.className = `inline-result ${mode}`.trim();
}

async function copyAuditFindings() {
  if (!lastAuditFindings.length) return;

  const lines = [
    ["Severity", "Category", "Location", "Check", "Detail", "Review"].join("\t"),
    ...lastAuditFindings.map((f) => [
      f.severity.toUpperCase(),
      f.category,
      f.address || f.source,
      f.title,
      f.detail,
      f.recommendation
    ].join("\t"))
  ];

  try {
    await navigator.clipboard.writeText(lines.join("\n"));
    const btn = $("copy-audit");
    const old = btn.textContent;
    btn.textContent = "Copied";
    setTimeout(() => btn.textContent = old, 1000);
  } catch {
    setAuditSummary("Clipboard access was blocked by Excel.", "error");
  }
}

async function exportAuditReport() {
  if (!lastAuditMeta || !lastAuditFindings.length) return;

  setWorking(true, "Building report");
  internalWorkbookMutationDepth += 1;
  try {
    const reportName = await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheets = workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const baseName = "Model Audit";
      const existing = new Set(sheets.items.map((sheet) => sheet.name.toLowerCase()));
      let name = baseName;
      let n = 2;
      while (existing.has(name.toLowerCase())) name = `${baseName} ${n++}`;

      const sheet = sheets.add(name);
      const counts = countAuditSeverities(lastAuditFindings);
      const status = deriveAuditStatus(counts);

      const summary = [
        ["MODEL AUDIT REPORT", ""],
        ["Audit status", status.label],
        ["Scope", lastAuditMeta.scopeLabel],
        ["Sheets scanned", lastAuditMeta.scannedSheets],
        ["Cells reviewed", lastAuditMeta.scannedCells],
        ["Formula cells", lastAuditMeta.formulaCells],
        ["Critical / High / Medium / Low", `${counts.critical} / ${counts.high} / ${counts.medium} / ${counts.low}`]
      ];

      sheet.getRangeByIndexes(0, 0, summary.length, 2).values = summary;
      sheet.getRange("A1:B1").merge(false);
      sheet.getRange("A1").format.font.bold = true;
      sheet.getRange("A1").format.font.size = 16;
      sheet.getRange("A2:A7").format.font.bold = true;

      const headers = [["Severity", "Category", "Location", "Check", "Detail", "Suggested review"]];
      sheet.getRange("A9:F9").values = headers;
      sheet.getRange("A9:F9").format.font.bold = true;
      sheet.getRange("A9:F9").format.fill.color = "#E9EEF5";

      const rows = lastAuditFindings.map((f) => [
        f.severity.toUpperCase(),
        f.category,
        f.address || f.source,
        f.title,
        f.detail,
        f.recommendation
      ]);

      if (rows.length) {
        const dataRange = sheet.getRangeByIndexes(9, 0, rows.length, 6);
        dataRange.values = rows;
        dataRange.format.wrapText = true;
        dataRange.format.verticalAlignment = "Top";
      }

      sheet.getRange("A:A").format.columnWidth = 72;
      sheet.getRange("B:B").format.columnWidth = 120;
      sheet.getRange("C:C").format.columnWidth = 150;
      sheet.getRange("D:D").format.columnWidth = 180;
      sheet.getRange("E:F").format.columnWidth = 300;
      sheet.getRangeByIndexes(0, 0, 9 + rows.length, 6).format.autofitRows();

      sheet.activate();
      sheet.getRange("A1").select();
      await context.sync();
      return name;
    });

    setAuditSummary(`Created report sheet: ${reportName}`, "success");
  } catch (error) {
    setAuditSummary(`Could not create report sheet: ${normalizeError(error)}`, "error");
  } finally {
    internalWorkbookMutationDepth = Math.max(0, internalWorkbookMutationDepth - 1);
    setWorking(false);
  }
}

/* ----------------------------
   Address / shared helpers
----------------------------- */

function parseQualifiedAddress(address) {
  const bang = findUnquotedBang(address);
  if (bang === -1) {
    return { sheetName: "", cellAddress: address.replace(/\$/g, "") };
  }

  let sheet = address.slice(0, bang).trim();
  let cell = address.slice(bang + 1).trim();

  if (sheet.startsWith("'") && sheet.endsWith("'")) {
    sheet = sheet.slice(1, -1).replace(/''/g, "'");
  }

  return { sheetName: sheet, cellAddress: cell.replace(/\$/g, "") };
}

function findUnquotedBang(text) {
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    if (text[i] === "'") {
      if (inQuotes && text[i + 1] === "'") {
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (text[i] === "!" && !inQuotes) {
      return i;
    }
  }
  return -1;
}

function qualifyAddress(sheetName, cellAddress) {
  const safeSheet = /[\s,'!]/.test(sheetName)
    ? `'${sheetName.replace(/'/g, "''")}'`
    : sheetName;
  return `${safeSheet}!${cellAddress}`;
}

function localAddress(address) {
  const parsed = parseQualifiedAddress(address);
  return parsed.cellAddress;
}

function isItemNotFound(error) {
  return error?.code === "ItemNotFound" ||
         error?.name === "ItemNotFound" ||
         String(error?.message || "").includes("ItemNotFound");
}

function isInvalidSelection(error) {
  return error?.code === "InvalidSelection" ||
         error?.name === "InvalidSelection" ||
         String(error?.message || "").toLowerCase().includes("invalid selection");
}

function normalizeError(error) {
  if (!error) return "Unknown error.";
  if (error.debugInfo?.errorLocation) {
    return `${error.message || error.code} (${error.debugInfo.errorLocation})`;
  }
  return error.message || error.code || String(error);
}

function showError(message) {
  if ($("explorer-tree")) {
    renderExplorerError(message);
    return;
  }
}

function setWorking(isWorking, label = "") {
  working = isWorking;

  document.querySelectorAll(
    ".command-btn, .scope-btn, .manual-btn, #recalculate-btn, #export-audit"
  ).forEach((button) => {
    if (isWorking) {
      button.dataset.wasDisabled = button.disabled ? "1" : "0";
      button.disabled = true;
    } else if (button.dataset.wasDisabled === "0") {
      button.disabled = false;
    }
  });

  if (isWorking) {
    setStatus(label || "Working", "working");
  } else {
    setStatus("", "");
  }
}

function setStatus(text, mode = "") {
  const status = $("status-pill");
  if (!status) return;
  status.textContent = text || "";
  status.className = `status-text ${mode || ""}`.trim();
  status.classList.toggle("hidden", !text);
}

function escapeHtml(text) {
  return String(text).replace(/[&<>"']/g, (ch) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;"
  }[ch]));
}
