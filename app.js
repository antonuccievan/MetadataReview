const THEME_STORAGE_KEY = "metadata-review-theme";
const SPELL_DICTIONARY_AFF_SOURCES = [
  "./dictionaries/en_US.aff",
  "https://cdn.jsdelivr.net/npm/dictionary-en-us@3.0.0/index.aff"
];
const SPELL_DICTIONARY_DIC_SOURCES = [
  "./dictionaries/en_US.dic",
  "https://cdn.jsdelivr.net/npm/dictionary-en-us@3.0.0/index.dic"
];

const FINANCE_ACCOUNTING_TERMS = new Set([
  "account","accounts","accounting","accrual","accruals","accrued","adjusting","adjustment","adjustments","amortization","amortize","amortized",
  "ap","ar","asset","assets","audit","audited","auditor","auditors","balance","balances","bank","banking","benchmark","book","booked","books",
  "budget","budgeted","budgeting","capex","capital","capitalization","capitalize","cash","cashflow","cashflows","chart","close","closing","coa",
  "collection","collections","compliance","consolidate","consolidated","consolidation","contra","controller","controllers","cost","costing","costs",
  "credit","credits","currency","current","debit","debits","deferred","depreciate","depreciated","depreciation","disclosure","disclosures","ebit",
  "ebitda","equity","expense","expenses","fair","finance","financial","financing","fiscal","forecast","forecasting","forecasts","fraud","gaap",
  "general","gl","goodwill","gross","impairment","income","incurred","indexation","indirect","inflation","interest","inventory","invoice","invoiced",
  "invoices","journal","journals","land","lease","leases","ledger","liability","liabilities","liquidation","loan","loans","longterm","margin",
  "materiality","measure","measurement","monthend","net","noncash","note","notes","operating","opex","otherincome","otherexpense","outflow","owner",
  "owners","ownership","payable","payables","payroll","period","periodic","posting","postings","ppe","prepaid","price","pricing","procurement",
  "profit","profitability","provision","provisions","purchase","purchases","ratio","ratios","receivable","receivables","reconcile","reconciled",
  "reconciliation","reconciliations","recognition","reserve","reserves","residual","restate","restated","restatement","retained","retention","revaluation",
  "revenue","revenues","rollforward","scenario","segment","segments","sellside","sga","soa","solvency","statement","statements","subledger","subsidiary",
  "tax","taxable","taxation","throughput","trading","transaction","transactions","treasury","trial","turnover","unearned","variance","variances","vendor",
  "vendors","workingcapital","writeoff","writeoffs","yearend","building","buildings"
]);

const EVERYDAY_ENGLISH_TERMS = new Set([
  "about","above","across","after","again","against","almost","along","already","also","always","among","another","answer","around","because",
  "before","behind","below","better","between","beyond","billion","business","buyer","calendar","called","cannot","change","children","client",
  "company","complete","country","customer","daily","dashboard","data","department","detail","different","document","during","effective","email",
  "employee","estimate","example","family","final","future","growth","group","health","history","holiday","important","include","increase","industry",
  "information","inside","issue","items","language","large","latest","letter","little","manage","manager","market","meeting","member","message",
  "million","minute","modern","money","monthly","morning","network","number","office","online","option","outside","payment","people","period",
  "personal","policy","project","quality","quarter","question","recent","record","report","result","review","safety","school","service","simple",
  "since","small","social","solution","source","special","standard","status","street","strong","support","system","target","team","today","tomorrow",
  "total","travel","update","value","version","weekly","within","without","world","yearly"
]);

const SPELL_CONFIDENCE_THRESHOLDS = {
  high: 0.8,
  medium: 0.6
};

const state = {
  workbook: null,
  dataRows: [],
  rows: [],
  headers: [],
  sourceHeaders: [],
  collapsed: new Set(),
  hierarchyRoots: [],
  loading: false,
  headerRowNumber: null,
  headerStartColumnNumber: 2,
  parentColumnNumber: 3,
  childColumnNumber: 4,
  reportType: "review",
  reportColumns: new Set(),
  spellStatusFilter: "all",
  filteredRows: [],
  findingsByRowId: new Map(),
  spellChecker: null,
  spellCheckerReady: false,
  spellAssessmentCache: new Map()
};

const fileInput = document.getElementById("fileInput");
const sheetSelect = document.getElementById("sheetSelect");
const tableWrap = document.getElementById("tableWrap");
const stats = document.getElementById("stats");
const expandAllBtn = document.getElementById("expandAllBtn");
const collapseAllBtn = document.getElementById("collapseAllBtn");
const themeToggleBtn = document.getElementById("themeToggleBtn");
const parentColumnSelect = document.getElementById("parentColumnSelect");
const childColumnSelect = document.getElementById("childColumnSelect");
const reportSelect = document.getElementById("reportSelect");
const spellOptionsWrap = document.getElementById("spellOptionsWrap");
const spellColumnSelect = document.getElementById("spellColumnSelect");
const reportColumnsLabel = document.getElementById("reportColumnsLabel");
const spellScorecard = document.getElementById("spellScorecard");
const scoreAllBtn = document.getElementById("scoreAllBtn");
const scorePassBtn = document.getElementById("scorePassBtn");
const scoreFailBtn = document.getElementById("scoreFailBtn");

fileInput?.addEventListener("change", handleFileUpload);
sheetSelect?.addEventListener("change", loadSheetRows);
expandAllBtn?.addEventListener("click", () => {
  state.collapsed.clear();
  renderTable();
});
collapseAllBtn?.addEventListener("click", () => {
  state.collapsed = new Set(state.rows.filter((r) => r.hasChildren).map((r) => r.id));
  state.rows.filter((r) => r.level === 0).forEach((r) => state.collapsed.delete(r.id));
  renderTable();
});
parentColumnSelect?.addEventListener("change", handleHierarchyColumnChange);
childColumnSelect?.addEventListener("change", handleHierarchyColumnChange);
reportSelect?.addEventListener("change", handleReportTypeChange);
spellColumnSelect?.addEventListener("change", () => {
  const selectedOptions = [...spellColumnSelect.selectedOptions];
  state.reportColumns.clear();
  selectedOptions.forEach((option) => {
    const columnIndex = Number(option.value);
    if (Number.isInteger(columnIndex) && columnIndex >= 0 && columnIndex < state.sourceHeaders.length) {
      state.reportColumns.add(state.sourceHeaders[columnIndex]);
    }
  });
  renderTable();
});

function setSpellStatusFilter(nextFilter) {
  state.spellStatusFilter = nextFilter;
  renderTable();
}

scoreAllBtn?.addEventListener("click", () => setSpellStatusFilter("all"));
scorePassBtn?.addEventListener("click", () => setSpellStatusFilter("pass"));
scoreFailBtn?.addEventListener("click", () => setSpellStatusFilter("fail"));

function updateScorecardButtons(passCount, failCount) {
  if (!scoreAllBtn || !scorePassBtn || !scoreFailBtn) return;
  scoreAllBtn.textContent = `All: ${passCount + failCount}`;
  scorePassBtn.textContent = `Pass: ${passCount}`;
  scoreFailBtn.textContent = `Fail: ${failCount}`;

  scoreAllBtn.classList.toggle("active", state.spellStatusFilter === "all");
  scorePassBtn.classList.toggle("active", state.spellStatusFilter === "pass");
  scoreFailBtn.classList.toggle("active", state.spellStatusFilter === "fail");
}

function isMisspelledToken(token) {
  return assessSpelling(token).misspelled;
}

function markMisspelledWords(input) {
  const text = String(input ?? "");
  const parts = text.split(/(\b[A-Za-z']+\b)/g);
  return parts
    .map((part) => {
      if (/^\b[A-Za-z']+\b$/.test(part) && isMisspelledToken(part)) {
        return `<span class="misspelled-token">${escapeHtml(part)}</span>`;
      }
      return escapeHtml(part);
    })
    .join("");
}

themeToggleBtn?.addEventListener("click", toggleTheme);
initializeTheme();
initializeSpellChecker();

tableWrap?.addEventListener(
  "wheel",
  (event) => {
    if (Math.abs(event.deltaY) <= Math.abs(event.deltaX)) return;
    if (!event.shiftKey) return;

    tableWrap.scrollLeft += event.deltaY;
    event.preventDefault();
  },
  { passive: false }
);

function setLoading(loading, message = "Processing workbook...") {
  state.loading = loading;
  if (loading) {
    tableWrap.innerHTML = `
      <div class="loading-wrap" role="status" aria-live="polite">
        <span class="spinner" aria-hidden="true"></span>
        <span>${escapeHtml(message)}</span>
      </div>
    `;
  }
}

function waitForPaint() {
  return new Promise((resolve) => requestAnimationFrame(() => resolve()));
}

async function initializeSpellChecker() {
  if (typeof window.Typo !== "function") {
    return;
  }

  try {
    const [affData, dicData] = await Promise.all([
      fetchFirstAvailableText(SPELL_DICTIONARY_AFF_SOURCES),
      fetchFirstAvailableText(SPELL_DICTIONARY_DIC_SOURCES)
    ]);

    if (!affData || !dicData) return;

    state.spellChecker = new window.Typo("en_US", affData, dicData, { platform: "any" });
    state.spellCheckerReady = true;
    state.spellAssessmentCache.clear();

    if (state.reportType === "spell") {
      renderTable();
    }
  } catch {
    // Fallback behavior is handled in assessSpelling when dictionary loading fails.
  }
}

async function fetchFirstAvailableText(urls) {
  for (const url of urls) {
    try {
      const response = await fetch(url, { cache: "force-cache" });
      if (response.ok) {
        return response.text();
      }
    } catch {
      // Try the next source.
    }
  }
  return null;
}

function initializeTheme() {
  const savedTheme = localStorage.getItem(THEME_STORAGE_KEY);
  const prefersDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
  const theme = savedTheme || (prefersDark ? "dark" : "light");
  applyTheme(theme);
}

function toggleTheme() {
  const currentTheme = document.documentElement.getAttribute("data-theme") || "light";
  const nextTheme = currentTheme === "dark" ? "light" : "dark";
  applyTheme(nextTheme);
  localStorage.setItem(THEME_STORAGE_KEY, nextTheme);
}

function applyTheme(theme) {
  document.documentElement.setAttribute("data-theme", theme);
  const icon = theme === "dark" ? "🌙" : "☀️";
  const label = theme === "dark" ? "Switch to light mode" : "Switch to dark mode";
  if (!themeToggleBtn) return;
  themeToggleBtn.textContent = `${icon} ${theme === "dark" ? "Dark" : "Light"}`;
  themeToggleBtn.setAttribute("aria-label", label);
  themeToggleBtn.setAttribute("title", label);
}

async function handleFileUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    setLoading(true, "Reading workbook...");
    await waitForPaint();

    const data = await file.arrayBuffer();
    state.workbook = XLSX.read(data, { type: "array", cellStyles: true });

    if (!sheetSelect) throw new Error("Sheet selector is unavailable in this page layout.");

    sheetSelect.innerHTML = '<option value="">Select a sheet</option>';
    state.workbook.SheetNames.forEach((name) => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      sheetSelect.append(option);
    });
    sheetSelect.disabled = false;

    const firstSheet = state.workbook.SheetNames[0];
    if (firstSheet) {
      sheetSelect.value = firstSheet;
      await loadSheetRows();
    } else {
      setLoading(false);
      tableWrap.innerHTML = '<div class="empty">Workbook contains no sheets.</div>';
    }
  } catch (error) {
    setLoading(false);
    tableWrap.innerHTML = `<div class="empty">Unable to read workbook: ${escapeHtml(error?.message || "unknown error")}</div>`;
  }
}

function findHeaderRowIndex(grid) {
  return grid.findIndex((row) => String(row[1] ?? "").trim() !== "");
}

function populateHierarchyColumnSelects(sourceHeaders) {
  if (!parentColumnSelect || !childColumnSelect) {
    throw new Error("Hierarchy selectors are unavailable in this page layout.");
  }

  const options = sourceHeaders
    .map((header, index) => {
      const columnNumber = state.headerStartColumnNumber + index;
      return `<option value="${columnNumber}">Column ${columnNumber}: ${escapeHtml(header)}</option>`;
    })
    .join("");

  parentColumnSelect.innerHTML = options;
  childColumnSelect.innerHTML = options;

  const fallbackParent = state.headerStartColumnNumber + 1;
  const fallbackChild = state.headerStartColumnNumber + 2;
  const lastAvailableColumn = state.headerStartColumnNumber + sourceHeaders.length - 1;

  state.parentColumnNumber = Math.min(fallbackParent, lastAvailableColumn);
  state.childColumnNumber = Math.min(fallbackChild, lastAvailableColumn);

  parentColumnSelect.value = String(state.parentColumnNumber);
  childColumnSelect.value = String(state.childColumnNumber);
  parentColumnSelect.disabled = false;
  childColumnSelect.disabled = false;
}

function populateReportColumns(sourceHeaders) {
  if (!spellColumnSelect || !reportSelect) {
    throw new Error("Report controls are unavailable in this page layout.");
  }

  if (sourceHeaders.length === 0) {
    spellColumnSelect.innerHTML = "";
    reportSelect.disabled = true;
    return;
  }

  reportSelect.disabled = false;

  spellColumnSelect.innerHTML = sourceHeaders
    .map((header, index) => `<option value="${index}" ${state.reportColumns.has(header) ? "selected" : ""}>${escapeHtml(header)}</option>`)
    .join("");

  reportSelect.value = state.reportType;
  syncReportParamVisibility();
}

function syncReportParamVisibility() {
  if (!spellOptionsWrap || !spellScorecard) return;
  const showColumnReport = state.reportType === "spell" || state.reportType === "space";
  spellOptionsWrap.hidden = !showColumnReport;
  spellScorecard.hidden = !showColumnReport;
  if (reportColumnsLabel) {
    reportColumnsLabel.textContent =
      state.reportType === "space" ? "Space Check columns (multi-select)" : "Spell check columns (multi-select)";
  }
  if (spellColumnSelect) {
    spellColumnSelect.disabled = !showColumnReport;
  }
}

function handleReportTypeChange() {
  if (!reportSelect) return;
  state.reportType = reportSelect.value;
  if (state.reportType === "review") {
    state.spellStatusFilter = "all";
    state.reportColumns.clear();
    if (spellColumnSelect) {
      [...spellColumnSelect.options].forEach((option) => {
        option.selected = false;
      });
    }
  }
  syncReportParamVisibility();
  renderTable();
}

function hasInternalSpace(value) {
  return /\s/.test(String(value ?? ""));
}

function handleHierarchyColumnChange() {
  if (state.dataRows.length === 0) return;

  const selectedParentColumn = Number(parentColumnSelect.value);
  const selectedChildColumn = Number(childColumnSelect.value);

  if (!Number.isFinite(selectedParentColumn) || !Number.isFinite(selectedChildColumn)) return;

  state.parentColumnNumber = selectedParentColumn;
  state.childColumnNumber = selectedChildColumn;
  runQuery();
}

function buildEntries(sourceRows) {
  return sourceRows.map((entry) => {
    const parentValue = normalizeHierarchyValue(entry.sourceRow[state.parentColumnNumber - 1]);
    const childValue = normalizeHierarchyValue(entry.sourceRow[state.childColumnNumber - 1]);

    return {
      ...entry,
      level: 0,
      parentId: null,
      hasChildren: false,
      isVisible: true,
      hierarchyLabel: childValue || parentValue || `(Row ${entry.sheetRowNumber})`,
      parentKey: parentValue,
      childKey: childValue,
      hierarchyPath: ""
    };
  });
}

function runQuery() {
  if (state.dataRows.length === 0) return;
  state.rows = buildEntries(state.dataRows);
  rebuildHierarchyFromParentChild();
}

async function loadSheetRows() {
  if (!sheetSelect) return;
  const sheetName = sheetSelect.value;
  if (!sheetName || !state.workbook) return;

  try {
    setLoading(true, `Processing sheet \"${sheetName}\"...`);
    await waitForPaint();

    const worksheet = state.workbook.Sheets[sheetName];
    const grid = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: "",
      raw: false,
      blankrows: false
    });

    if (grid.length === 0) {
      setLoading(false);
      tableWrap.innerHTML = '<div class="empty">The selected sheet is empty.</div>';
      return;
    }

    const headerRowIndex = findHeaderRowIndex(grid);
    if (headerRowIndex < 0) {
      setLoading(false);
      tableWrap.innerHTML = '<div class="empty">No header row found: column B is empty for all rows.</div>';
      return;
    }

    const headerRow = grid[headerRowIndex] || [];
    const headerStartColumnIndex = 1;

    state.headerRowNumber = headerRowIndex + 1;
    state.headerStartColumnNumber = headerStartColumnIndex + 1;
    state.parentColumnNumber = state.headerStartColumnNumber + 1;
    state.childColumnNumber = state.headerStartColumnNumber + 2;

    const sourceHeaders = headerRow
      .slice(headerStartColumnIndex)
      .map((h, i) => String(h ?? "").trim() || `Column ${state.headerStartColumnNumber + i}`);

    if (sourceHeaders.length === 0) {
      setLoading(false);
      tableWrap.innerHTML = '<div class="empty">The selected sheet has no columns after column B.</div>';
      return;
    }

    state.sourceHeaders = sourceHeaders;
    state.headers = ["Hierarchy", ...sourceHeaders];

    const missingReportColumns = [...state.reportColumns].filter((header) => !sourceHeaders.includes(header));
    missingReportColumns.forEach((header) => state.reportColumns.delete(header));

    populateHierarchyColumnSelects(sourceHeaders);
    populateReportColumns(sourceHeaders);

    const dataRows = grid.slice(headerRowIndex + 1).map((row, index) => {
      const mapped = {};
      sourceHeaders.forEach((header, colIndex) => {
        mapped[header] = row[colIndex + headerStartColumnIndex] ?? "";
      });

      return {
        id: `row-${index}`,
        originalIndex: index,
        sourceRow: row,
        sheetRowNumber: headerRowIndex + index + 2,
        row: mapped
      };
    });

    state.dataRows = dataRows;
    state.rows = buildEntries(dataRows.slice(0, 10));

    if (expandAllBtn) expandAllBtn.disabled = false;
    if (collapseAllBtn) collapseAllBtn.disabled = false;

    rebuildHierarchyFromParentChild();
    runQuery();
  } catch (error) {
    setLoading(false);
    tableWrap.innerHTML = `<div class="empty">Unable to process sheet: ${escapeHtml(error?.message || "unknown error")}</div>`;
  }
}

function normalizeHierarchyValue(value) {
  return String(value ?? "").trim();
}

function rebuildHierarchyFromParentChild() {
  state.collapsed.clear();

  const rowsByChild = new Map();
  state.rows.forEach((entry) => {
    if (!entry.childKey) return;
    if (!rowsByChild.has(entry.childKey)) rowsByChild.set(entry.childKey, []);
    rowsByChild.get(entry.childKey).push(entry);
  });

  state.rows.forEach((entry) => {
    entry.parentId = null;
    entry.level = 0;
    if (!entry.parentKey) return;

    const possibleParents = rowsByChild.get(entry.parentKey) || [];
    const parent = possibleParents.find((candidate) => candidate.id !== entry.id);
    if (parent) {
      entry.parentId = parent.id;
    }
  });

  const childrenByParent = new Map();
  state.rows.forEach((entry) => {
    const key = entry.parentId || "ROOT";
    if (!childrenByParent.has(key)) childrenByParent.set(key, []);
    childrenByParent.get(key).push(entry);
  });

  state.rows.forEach((entry) => {
    entry.hasChildren = (childrenByParent.get(entry.id) || []).length > 0;
  });

  const roots = childrenByParent.get("ROOT") || [];
  state.hierarchyRoots = roots.map((r) => r.id);

  const rowById = new Map(state.rows.map((r) => [r.id, r]));
  const visiting = new Set();

  const assignLevel = (entry) => {
    if (!entry.parentId) {
      entry.level = 0;
      return;
    }
    if (visiting.has(entry.id)) {
      entry.parentId = null;
      entry.level = 0;
      return;
    }

    visiting.add(entry.id);
    const parent = rowById.get(entry.parentId);
    if (!parent) {
      entry.parentId = null;
      entry.level = 0;
    } else {
      assignLevel(parent);
      entry.level = parent.level + 1;
    }
    visiting.delete(entry.id);
  };

  state.rows.forEach(assignLevel);

  state.rows.forEach((entry) => {
    const path = [];
    const seen = new Set();
    let current = entry;
    while (current && !seen.has(current.id)) {
      seen.add(current.id);
      path.push(current.hierarchyLabel);
      current = current.parentId ? rowById.get(current.parentId) : null;
    }

    if (current) path.push("[Cycle]");
    entry.hierarchyPath = path.reverse().join(" > ");
  });

  const ordered = [];
  const visited = new Set();

  const visit = (node) => {
    if (!node || visited.has(node.id)) return;
    visited.add(node.id);
    ordered.push(node);

    const children = (childrenByParent.get(node.id) || []).slice().sort((a, b) => a.originalIndex - b.originalIndex);
    children.forEach(visit);
  };

  roots.slice().sort((a, b) => a.originalIndex - b.originalIndex).forEach(visit);
  state.rows.forEach(visit);
  state.rows = ordered;

  if (expandAllBtn) expandAllBtn.disabled = false;
  if (collapseAllBtn) collapseAllBtn.disabled = false;

  setLoading(false);
  renderTable();
}

function tokenize(value) {
  return String(value ?? "")
    .split(/[^A-Za-z']+/)
    .map((token) => token.trim())
    .filter(Boolean);
}

function getTokenCandidates(normalized) {
  return [
    normalized,
    normalized.replace(/'(s|d|ll|re|ve|m|t)$/i, ""),
    normalized.replace(/ies$/i, "y"),
    normalized.replace(/es$/i, ""),
    normalized.replace(/s$/i, "")
  ].filter((value, index, values) => value.length > 2 && values.indexOf(value) === index);
}

function editDistance(a, b) {
  const matrix = Array.from({ length: a.length + 1 }, () => Array(b.length + 1).fill(0));
  for (let i = 0; i <= a.length; i += 1) matrix[i][0] = i;
  for (let j = 0; j <= b.length; j += 1) matrix[0][j] = j;

  for (let i = 1; i <= a.length; i += 1) {
    for (let j = 1; j <= b.length; j += 1) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      matrix[i][j] = Math.min(matrix[i - 1][j] + 1, matrix[i][j - 1] + 1, matrix[i - 1][j - 1] + cost);
    }
  }
  return matrix[a.length][b.length];
}

function findClosestLexiconWord(token) {
  const referenceWords = new Set([...FINANCE_ACCOUNTING_TERMS, ...EVERYDAY_ENGLISH_TERMS]);
  let closestWord = "";
  let smallestDistance = Number.POSITIVE_INFINITY;

  referenceWords.forEach((word) => {
    const distance = editDistance(token, word);
    if (distance < smallestDistance) {
      smallestDistance = distance;
      closestWord = word;
    }
  });

  return { word: closestWord, distance: smallestDistance };
}

function getConfidenceLabel(score) {
  if (score >= SPELL_CONFIDENCE_THRESHOLDS.high) return "High";
  if (score >= SPELL_CONFIDENCE_THRESHOLDS.medium) return "Medium";
  return "Low";
}

function assessSpelling(token) {
  if (!token) {
    return { misspelled: false, confidenceScore: 0, confidenceLabel: "Low", suggestion: "" };
  }
  if (/\d/.test(token)) {
    return { misspelled: false, confidenceScore: 0, confidenceLabel: "Low", suggestion: "" };
  }

  const cleaned = token.replace(/^'+|'+$/g, "");
  const normalized = cleaned.toLowerCase();
  if (!normalized) {
    return { misspelled: false, confidenceScore: 0, confidenceLabel: "Low", suggestion: "" };
  }

  if (state.spellAssessmentCache.has(normalized)) {
    return state.spellAssessmentCache.get(normalized);
  }

  const fallbackResult = { misspelled: false, confidenceScore: 0, confidenceLabel: "Low", suggestion: "" };

  if (normalized.length <= 2) {
    state.spellAssessmentCache.set(normalized, fallbackResult);
    return fallbackResult;
  }

  const businessCandidates = getTokenCandidates(normalized);
  const inBusinessLexicon = businessCandidates.some((candidate) => FINANCE_ACCOUNTING_TERMS.has(candidate));
  const inEverydayLexicon = businessCandidates.some((candidate) => EVERYDAY_ENGLISH_TERMS.has(candidate));
  const inSupplementalLexicon = inBusinessLexicon || inEverydayLexicon;
  const dictionaryAvailable = state.spellCheckerReady && Boolean(state.spellChecker);

  if (dictionaryAvailable) {
    const spellCandidates = [
      cleaned,
      normalized,
      normalized.charAt(0).toUpperCase() + normalized.slice(1),
      ...businessCandidates
    ].filter((value, index, values) => value && values.indexOf(value) === index);

    if (spellCandidates.some((candidate) => state.spellChecker.check(candidate))) {
      state.spellAssessmentCache.set(normalized, fallbackResult);
      return fallbackResult;
    }
  }

  if (inSupplementalLexicon) {
    state.spellAssessmentCache.set(normalized, fallbackResult);
    return fallbackResult;
  }

  let suggestion = "";
  let distance = Number.POSITIVE_INFINITY;

  if (dictionaryAvailable) {
    const suggestions = state.spellChecker.suggest(cleaned, 5) || [];
    suggestion = suggestions[0] || "";
    if (suggestion) {
      distance = editDistance(normalized, suggestion.toLowerCase());
    }
  }

  if (!suggestion) {
    const nearest = findClosestLexiconWord(normalized);
    if (nearest.word) {
      suggestion = nearest.word.charAt(0).toUpperCase() + nearest.word.slice(1);
      distance = nearest.distance;
    }
  }

  let confidenceScore = 0.55;
  if (normalized.length >= 5) confidenceScore += 0.1;
  if (distance <= 1) confidenceScore += 0.28;
  else if (distance === 2) confidenceScore += 0.2;
  else if (distance === 3) confidenceScore += 0.1;
  if (!dictionaryAvailable) confidenceScore -= 0.1;
  confidenceScore = Math.max(0.35, Math.min(0.98, confidenceScore));

  const result = {
    misspelled: true,
    confidenceScore,
    confidenceLabel: getConfidenceLabel(confidenceScore),
    suggestion
  };
  state.spellAssessmentCache.set(normalized, result);
  return result;
}

function evaluateSpellRow(entry, selectedColumns) {
  const issues = [];
  selectedColumns.forEach((header) => {
    const value = String(entry.row[header] ?? "");
    if (!value.trim()) return;

    const assessments = [];
    [...new Set(tokenize(value))].forEach((token) => {
      const assessment = assessSpelling(token);
      if (!assessment.misspelled) return;
      const confidencePercent = `${Math.round(assessment.confidenceScore * 100)}%`;
      const suggestionText = assessment.suggestion ? `, suggestion: ${assessment.suggestion}` : "";
      assessments.push(`${token} (${assessment.confidenceLabel} ${confidencePercent}${suggestionText})`);
    });

    if (assessments.length > 0) {
      issues.push(`${header}: ${assessments.join(", ")}`);
    }
  });

  return issues;
}

function evaluateSpaceRow(entry, selectedColumns) {
  const issues = [];
  selectedColumns.forEach((header) => {
    const value = String(entry.row[header] ?? "");
    if (!value) return;
    if (hasInternalSpace(value)) {
      issues.push(`${header}: contains spaces`);
    }
  });

  return issues;
}

function applyReportFilter(rows) {
  const selectedColumns = [...state.reportColumns].filter((header) => state.sourceHeaders.includes(header));
  if ((state.reportType !== "spell" && state.reportType !== "space") || selectedColumns.length === 0) {
    return {
      rows,
      findingsByRowId: new Map(),
      statusByRowId: new Map(),
      passCount: rows.length,
      failCount: 0,
      selectedColumns
    };
  }

  const findingsByRowId = new Map();
  const statusByRowId = new Map();
  let passCount = 0;
  let failCount = 0;

  rows.forEach((entry) => {
    const issues = state.reportType === "space" ? evaluateSpaceRow(entry, selectedColumns) : evaluateSpellRow(entry, selectedColumns);
    const status = issues.length > 0 ? "Fail" : "Pass";
    statusByRowId.set(entry.id, status);
    if (issues.length > 0) {
      findingsByRowId.set(entry.id, issues);
      failCount += 1;
    } else {
      passCount += 1;
    }
  });

  const filteredRows = rows.filter((entry) => {
    const status = statusByRowId.get(entry.id);
    if (state.spellStatusFilter === "pass") return status === "Pass";
    if (state.spellStatusFilter === "fail") return status === "Fail";
    return true;
  });

  return { rows: filteredRows, findingsByRowId, statusByRowId, passCount, failCount, selectedColumns };
}

function renderTable() {
  const rowById = new Map(state.rows.map((r) => [r.id, r]));

  state.rows.forEach((entry) => {
    let current = entry.parentId;
    let hidden = false;
    while (current) {
      if (state.collapsed.has(current)) {
        hidden = true;
        break;
      }
      current = rowById.get(current)?.parentId || null;
    }
    entry.isVisible = !hidden;
  });

  const visibleRows = state.rows.filter((r) => r.isVisible);
  const { rows: reportedRows, findingsByRowId, statusByRowId, passCount, failCount, selectedColumns } = applyReportFilter(visibleRows);
  state.filteredRows = reportedRows;
  state.findingsByRowId = findingsByRowId;

  if (state.reportType === "spell" || state.reportType === "space") {
    updateScorecardButtons(passCount, failCount);
  }

  const rowsToRender = reportedRows;
  const visibleCount = rowsToRender.length;
  const maxDepth = Math.max(...rowsToRender.map((r) => r.level), 0);
  const reportSummary =
    state.reportType === "spell"
      ? `Spell check · ${selectedColumns.length} column(s)`
      : state.reportType === "space"
        ? `Space Check · ${selectedColumns.length} column(s)`
        : "Review";

  stats.innerHTML = `
    <span>Headers: <strong>row ${state.headerRowNumber ?? "?"}</strong></span>
    <span>Total rows: <strong>${state.rows.length}</strong></span>
    <span>Shown rows: <strong>${visibleCount}</strong></span>
    <span>Max depth: <strong>${rowsToRender.length > 0 ? maxDepth + 1 : 0}</strong></span>
    <span>Grouping source: <strong>column ${state.parentColumnNumber} (parent) → column ${state.childColumnNumber} (child)</strong></span>
    <span>Report: <strong>${escapeHtml(reportSummary)}</strong></span>
  `;

  const isColumnReport = state.reportType === "spell" || state.reportType === "space";
  const sourceHeaders = isColumnReport ? selectedColumns : state.headers.slice(1);
  const headerSet = isColumnReport ? ["Hierarchy", "Status", "Report Findings", ...sourceHeaders] : state.headers;
  const headerCells = headerSet.map((h, i) => `<th class="${i === 0 ? "sticky-col" : ""}">${escapeHtml(h)}</th>`).join("");

  const bodyRows = rowsToRender
    .map((entry) => {
      const hierarchyCell = `
        <td class="sticky-col">
          <div class="hierarchy-cell" title="${escapeHtml(entry.hierarchyPath)}">
            <span class="indent-guide" style="--indent:${entry.level * 18}px"></span>
            ${
              entry.hasChildren
                ? `
              <button class="twisty" data-id="${entry.id}" data-action="expand" aria-label="Expand row">+</button>
              <button class="twisty" data-id="${entry.id}" data-action="collapse" aria-label="Collapse row">−</button>
            `
                : '<span class="twisty placeholder">•</span>'
            }
            <span>${escapeHtml(entry.hierarchyLabel)}</span>
            <span class="pill">L${entry.level + 1}</span>
          </div>
        </td>
      `;

      const findingCell = !isColumnReport ? "" : `<td>${escapeHtml((state.findingsByRowId.get(entry.id) || []).join(" | "))}</td>`;

      const statusCell = !isColumnReport ? "" : `<td>${statusByRowId.get(entry.id) || "Pass"}</td>`;

      const rowCells = sourceHeaders
        .map((header) => {
          const cellValue = String(entry.row[header] ?? "");
          if (state.reportType === "spell") {
            return `<td>${markMisspelledWords(cellValue)}</td>`;
          }
          return `<td>${escapeHtml(cellValue)}</td>`;
        })
        .join("");

      return `<tr>${hierarchyCell}${statusCell}${findingCell}${rowCells}</tr>`;
    })
    .join("");

  tableWrap.innerHTML = `
    <table>
      <thead><tr>${headerCells}</tr></thead>
      <tbody>${bodyRows || `<tr><td colspan="${headerSet.length}"><div class="empty">No rows match the active report filters.</div></td></tr>`}</tbody>
    </table>
  `;

  tableWrap.querySelectorAll(".twisty[data-id]").forEach((button) => {
    button.addEventListener("click", () => {
      const id = button.getAttribute("data-id");
      const action = button.getAttribute("data-action");
      if (!id) return;

      if (action === "expand") {
        state.collapsed.delete(id);
      } else if (action === "collapse") {
        state.collapsed.add(id);
      } else if (state.collapsed.has(id)) {
        state.collapsed.delete(id);
      } else {
        state.collapsed.add(id);
      }
      renderTable();
    });
  });
}

function escapeHtml(input) {
  return String(input)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
