const state = {
  workbook: null,
  rows: [],
  headers: [],
  hierarchyColumn: "",
  collapsed: new Set(),
  tree: []
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
  rebuildTreeFromCurrentRows();
});

async function handleFileUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;

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

  if (state.workbook.SheetNames.length === 1) {
    sheetSelect.value = state.workbook.SheetNames[0];
    loadSheetRows();
  }
}

function loadSheetRows() {
  const sheetName = sheetSelect.value;
  if (!sheetName || !state.workbook) return;

  const worksheet = state.workbook.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    raw: false,
    blankrows: false
  });

  if (grid.length === 0) {
    tableWrap.innerHTML = '<div class="empty">The selected sheet is empty.</div>';
    return;
  }

  state.headers = (grid[0] || []).map((h, i) => (h || `Column ${i + 1}`).toString().trim());
  const dataRows = grid.slice(1).map((row) => {
    const mapped = {};
    state.headers.forEach((header, index) => {
      mapped[header] = row[index] ?? "";
    });
    return mapped;
  });

  state.rows = dataRows.map((row, index) => ({
    id: `row-${index}`,
    originalIndex: index,
    row,
    level: 0,
    parentId: null,
    hasChildren: false,
    isVisible: true
  }));

  rebuildTreeFromCurrentRows();
}

function rebuildTreeFromCurrentRows() {
  state.hierarchyColumn = state.headers[0] || "Column 1";
  state.collapsed.clear();

  if (!state.hierarchyColumn || state.headers.length === 0) {
    tableWrap.innerHTML = '<div class="empty">Unable to parse headers from the first row.</div>';
    return;
  }

  const stack = [];
  state.rows.forEach((entry) => {
    const value = String(entry.row[state.hierarchyColumn] ?? "");
    const level = detectLevel(value);

    entry.level = level;
    while (stack.length > level) {
      stack.pop();
    }

    entry.parentId = stack.length ? stack[stack.length - 1] : null;
    stack[level] = entry.id;
  });

  const childrenByParent = new Map();
  state.rows.forEach((entry) => {
    const key = entry.parentId || "ROOT";
    if (!childrenByParent.has(key)) childrenByParent.set(key, []);
    childrenByParent.get(key).push(entry.id);
  });

  state.rows.forEach((entry) => {
    entry.hasChildren = (childrenByParent.get(entry.id) || []).length > 0;
  });

  state.tree = state.rows;

  expandAllBtn.disabled = false;
  collapseAllBtn.disabled = false;
  resetBtn.disabled = false;

  renderTable();
}

function detectLevel(value) {
  const tabs = (value.match(/^\t+/) || [""])[0].length;
  if (tabs > 0) return tabs;

  const leadingSpaces = (value.match(/^ +/) || [""])[0].length;
  if (leadingSpaces > 0) return Math.floor(leadingSpaces / 2);

  const bulletPattern = value.match(/^(?:\.|-|\*|•)+\s*/);
  if (bulletPattern) return bulletPattern[0].replace(/\s/g, "").length;

  return 0;
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
    <span>Grouping source: <strong>indent in ${state.hierarchyColumn}</strong></span>
  `;

  const headerCells = state.headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("");
  const bodyRows = state.rows
    .map((entry) => {
      const rowCells = state.headers
        .map((header) => {
          const value = String(entry.row[header] ?? "");

          if (header !== state.hierarchyColumn) {
            return `<td>${escapeHtml(value)}</td>`;
          }

          const twisty = entry.hasChildren
            ? `<button class="twisty" data-id="${entry.id}" aria-label="Toggle row">${state.collapsed.has(entry.id) ? "▶" : "▼"}</button>`
            : '<span class="twisty placeholder">•</span>';

          return `
            <td>
              <div class="hierarchy-cell">
                <span class="indent-guide" style="--indent:${entry.level * 18}px"></span>
                ${twisty}
                <span>${escapeHtml(value.trimStart())}</span>
                <span class="pill">L${entry.level + 1}</span>
              </div>
            </td>`;
        })
        .join("");

      return `<tr class="${entry.isVisible ? "" : "hidden"}">${rowCells}</tr>`;
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
      if (!id) return;

      if (state.collapsed.has(id)) {
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
