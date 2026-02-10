const THEME_STORAGE_KEY = "metadata-review-theme";

const BASIC_DICTIONARY_WORDS = new Set([
  "a","about","above","across","after","again","against","all","also","am","an","and","any","are","as","at","be","been","before","being","below","between","both","but","by",
  "can","case","check","child","column","columns","company","data","date","description","document","documents","down","each","email","entry","error","errors","file","files","for","from","group","groups","has","have","he","her","here","his","how","i","id","if","in","into","is","it","its","item","items","job","key","keys","label","line","list","lower","made","may","metadata","mode","name","new","no","not","number","of","on","one","or","order","other","our","out","over","parent","path","proper","record","records","report","reports","review","row","rows","run","same","select","sheet","should","show","simple","spell","start","status","string","table","text","that","the","their","them","there","these","this","to","type","under","upper","up","use","value","values","was","we","when","where","which","with","work","workflow","you","your"
]);

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
  findingsByRowId: new Map()
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
  return looksMisspelled(token);
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
  const showSpell = state.reportType === "spell";
  spellOptionsWrap.hidden = !showSpell;
  spellScorecard.hidden = !showSpell;
  if (spellColumnSelect) {
    spellColumnSelect.disabled = !showSpell;
  }
}

function handleReportTypeChange() {
  if (!reportSelect) return;
  state.reportType = reportSelect.value;
  if (state.reportType !== "spell") {
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

function looksMisspelled(token) {
  if (!token) return false;
  if (/\d/.test(token)) return false;
  if (token.length <= 2) return false;
  if (/^[A-Z]{2,6}$/.test(token)) return false;

  const cleaned = token.replace(/^'+|'+$/g, "").toLowerCase();
  if (cleaned.length <= 2) return false;
  return !BASIC_DICTIONARY_WORDS.has(cleaned);
}

function evaluateSpellRow(entry, selectedColumns) {
  const issues = [];
  selectedColumns.forEach((header) => {
    const value = String(entry.row[header] ?? "");
    if (!value.trim()) return;

    const misspelled = [...new Set(tokenize(value).filter(looksMisspelled))];
    if (misspelled.length > 0) {
      issues.push(`${header}: ${misspelled.join(", ")}`);
    }
  });

  return issues;
}

function applyReportFilter(rows) {
  const selectedColumns = [...state.reportColumns].filter((header) => state.sourceHeaders.includes(header));
  if (state.reportType !== "spell" || selectedColumns.length === 0) {
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
    const issues = evaluateSpellRow(entry, selectedColumns);
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

  if (state.reportType === "spell") {
    updateScorecardButtons(passCount, failCount);
  }

  const rowsToRender = reportedRows;
  const visibleCount = rowsToRender.length;
  const maxDepth = Math.max(...rowsToRender.map((r) => r.level), 0);
  const reportSummary = state.reportType === "spell" ? `Spell check · ${selectedColumns.length} column(s)` : "Review";

  stats.innerHTML = `
    <span>Headers: <strong>row ${state.headerRowNumber ?? "?"}</strong></span>
    <span>Total rows: <strong>${state.rows.length}</strong></span>
    <span>Shown rows: <strong>${visibleCount}</strong></span>
    <span>Max depth: <strong>${rowsToRender.length > 0 ? maxDepth + 1 : 0}</strong></span>
    <span>Grouping source: <strong>column ${state.parentColumnNumber} (parent) → column ${state.childColumnNumber} (child)</strong></span>
    <span>Report: <strong>${escapeHtml(reportSummary)}</strong></span>
  `;

  const sourceHeaders = state.reportType === "spell" ? selectedColumns : state.headers.slice(1);
  const headerSet = state.reportType === "spell" ? ["Hierarchy", "Status", "Report Findings", ...sourceHeaders] : state.headers;
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

      const findingCell =
        state.reportType !== "spell"
          ? ""
          : `<td>${escapeHtml((state.findingsByRowId.get(entry.id) || []).join(" | "))}</td>`;

      const statusCell = state.reportType !== "spell" ? "" : `<td>${statusByRowId.get(entry.id) || "Pass"}</td>`;

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
