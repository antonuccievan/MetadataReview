const state = {
  workbook: null,
  rows: [],
  headers: [],
  collapsed: new Set(),
  hierarchyRoots: [],
  loading: false
};

const fileInput = document.getElementById("fileInput");
const sheetSelect = document.getElementById("sheetSelect");
const tableWrap = document.getElementById("tableWrap");
const stats = document.getElementById("stats");
const expandAllBtn = document.getElementById("expandAllBtn");
const collapseAllBtn = document.getElementById("collapseAllBtn");
const resetBtn = document.getElementById("resetBtn");

fileInput.addEventListener("change", handleFileUpload);
sheetSelect.addEventListener("change", loadSheetRows);
expandAllBtn.addEventListener("click", () => {
  state.collapsed.clear();
  renderTable();
});
collapseAllBtn.addEventListener("click", () => {
  state.collapsed = new Set(state.rows.filter((r) => r.hasChildren).map((r) => r.id));
  state.rows.filter((r) => r.level === 0).forEach((r) => state.collapsed.delete(r.id));
  renderTable();
});
resetBtn.addEventListener("click", () => {
  if (state.rows.length === 0) return;
  rebuildHierarchyFromParentChild();
});

tableWrap.addEventListener(
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

async function handleFileUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  setLoading(true, "Reading workbook...");
  await waitForPaint();

  const data = await file.arrayBuffer();
  state.workbook = XLSX.read(data, { type: "array", cellStyles: true });

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
}

async function loadSheetRows() {
  const sheetName = sheetSelect.value;
  if (!sheetName || !state.workbook) return;

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

  const headerRow = grid[0] || [];
  const firstHeaderCell = String(headerRow[0] ?? "").trim();
  const headerOffset = firstHeaderCell ? 0 : 1;

  const sourceHeaders = headerRow
    .slice(headerOffset)
    .map((h, i) => (String(h ?? "").trim() || `Column ${headerOffset + i + 1}`));

  state.headers = ["Hierarchy", ...sourceHeaders];

  const dataRows = grid.slice(1).map((row, index) => {
    const mapped = {};
    sourceHeaders.forEach((header, colIndex) => {
      mapped[header] = row[colIndex + headerOffset] ?? "";
    });

    const parentValue = normalizeHierarchyValue(row[2]);
    const childValue = normalizeHierarchyValue(row[3]);

    return {
      id: `row-${index}`,
      originalIndex: index,
      row: mapped,
      level: 0,
      parentId: null,
      hasChildren: false,
      isVisible: true,
      hierarchyLabel: childValue || parentValue || `(Row ${index + 2})`,
      parentKey: parentValue,
      childKey: childValue,
      hierarchyPath: ""
    };
  });

  state.rows = dataRows;
  rebuildHierarchyFromParentChild();
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
    let current = entry;
    while (current) {
      path.push(current.hierarchyLabel);
      current = current.parentId ? rowById.get(current.parentId) : null;
    }
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

  expandAllBtn.disabled = false;
  collapseAllBtn.disabled = false;
  resetBtn.disabled = false;

  setLoading(false);
  renderTable();
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

  const visibleCount = state.rows.filter((r) => r.isVisible).length;
  const maxDepth = Math.max(...state.rows.map((r) => r.level), 0);

  stats.innerHTML = `
    <span>Headers: <strong>row 1</strong></span>
    <span>Total rows: <strong>${state.rows.length}</strong></span>
    <span>Visible rows: <strong>${visibleCount}</strong></span>
    <span>Max depth: <strong>${maxDepth + 1}</strong></span>
    <span>Grouping source: <strong>column 3 (parent) → column 4 (child)</strong></span>
  `;

  const headerCells = state.headers.map((h, i) => `<th class="${i === 0 ? "sticky-col" : ""}">${escapeHtml(h)}</th>`).join("");
  const sourceHeaders = state.headers.slice(1);

  const bodyRows = state.rows
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

      const rowCells = sourceHeaders
        .map((header) => `<td>${escapeHtml(String(entry.row[header] ?? ""))}</td>`)
        .join("");

      return `<tr class="${entry.isVisible ? "" : "hidden"}">${hierarchyCell}${rowCells}</tr>`;
    })
    .join("");

  tableWrap.innerHTML = `
    <table>
      <thead><tr>${headerCells}</tr></thead>
      <tbody>${bodyRows}</tbody>
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
  return input
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
