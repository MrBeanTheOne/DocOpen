// Workbook (xlsx/csv) UI: grid rendering, sheet CRUD, selection ranges,
// copy/paste, row/column ops, sorting, and DOM<->model sync. Classic script;
// shares renderer.js globals (currentDocument, selectedCell, editMode, ...)
// the same way formula.js shares window.docFormula.

const WORKBOOK_DISPLAY_ROWS = {
  narrow: 30,
  wide: 34,
  full: 38
};
const WORKBOOK_DISPLAY_COLUMNS = {
  narrow: 10,
  wide: 12,
  full: 16
};

function renderWorkbook(workbookData) {
  wordEditor = null;
  savedWordRange = null;

  // CSV raw view: the serialized text instead of the grid (view mode only).
  // syncWorkbookFromDom is grid-guarded, so the model survives the detour.
  if (!editMode && workbookData.rawView && workbookData.extension === ".csv") {
    sheetPanel.hidden = true;
    selectedCell = null;
    selectedRange = null;
    rangeAnchor = null;
    updateSelectionLabel();
    const pre = document.createElement("pre");
    pre.className = `text-document document-surface width-${getDocumentLayout(workbookData)}`;
    pre.textContent = csvTextForSheet(workbookData.sheets[selectedSheetIndex] || workbookData.sheets[0]);
    viewer.replaceChildren(pre);
    return;
  }

  sheetPanel.hidden = workbookData.sheets.length === 0;
  sheetList.replaceChildren(
    ...workbookData.sheets.map((sheet, index) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = index === selectedSheetIndex ? "sheet-tab is-active" : "sheet-tab";
      button.textContent = sheet.name;
      button.title = "Double-click to rename, right-click to delete";
      button.addEventListener("click", () => {
        syncWorkbookFromDom();
        selectedSheetIndex = index;
        selectedCell = null;
        selectedRange = null;
        rangeAnchor = null;
        renderWorkbook(workbookData);
      });
      button.addEventListener("dblclick", () => {
        renameSheet(index);
      });
      button.addEventListener("contextmenu", (event) => {
        event.preventDefault();
        deleteSheet(index);
      });
      return button;
    })
  );

  if (workbookData.extension !== ".csv") {
    // CSV files only ever save their first sheet, so don't offer more.
    const addSheetButton = document.createElement("button");
    addSheetButton.type = "button";
    addSheetButton.className = "sheet-add-button";
    addSheetButton.textContent = "+ Add sheet";
    addSheetButton.addEventListener("click", () => {
      addSheet();
    });
    sheetList.appendChild(addSheetButton);
  }

  const sheet = workbookData.sheets[selectedSheetIndex];
  if (!sheet) {
    selectedCell = null;
    updateSelectionLabel();
    viewer.replaceChildren(emptyBlock("This sheet is empty."));
    return;
  }

  const workbookShell = document.createElement("section");
  workbookShell.className = `workbook-document document-surface width-${getDocumentLayout(workbookData)}`;

  const tableWrap = document.createElement("div");
  tableWrap.className = "table-wrap";

  ensureSheetLayout(sheet);

  const table = document.createElement("table");
  const actualColumnCount = getSheetColumnCount(sheet);
  const displayColumnCount = getWorkbookDisplayColumnCount(sheet, actualColumnCount);
  const displayRowCount = getWorkbookDisplayRowCount(sheet);
  ensureSheetLayout(sheet, actualColumnCount);
  const rowHeaderWidth = getRowHeaderWidth(displayRowCount);
  table.className = "sheet-table";
  table.style.setProperty("--row-header-width", `${rowHeaderWidth}px`);

  const colgroup = document.createElement("colgroup");
  const rowHeaderCol = document.createElement("col");
  rowHeaderCol.className = "row-header-col";
  rowHeaderCol.style.width = `${rowHeaderWidth}px`;
  colgroup.appendChild(rowHeaderCol);

  for (let columnIndex = 0; columnIndex < displayColumnCount; columnIndex += 1) {
    const col = document.createElement("col");
    col.dataset.columnIndex = String(columnIndex);
    col.style.width = `${getColumnWidth(sheet, columnIndex)}px`;
    colgroup.appendChild(col);
  }

  table.appendChild(colgroup);

  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  const cornerHeader = document.createElement("th");
  cornerHeader.className = "corner-header";
  headerRow.appendChild(cornerHeader);

  for (let columnIndex = 0; columnIndex < displayColumnCount; columnIndex += 1) {
    const th = document.createElement("th");
    th.className = "column-header";
    th.dataset.columnIndex = String(columnIndex);
    th.textContent = columnName(columnIndex);
    th.style.width = `${getColumnWidth(sheet, columnIndex)}px`;
    th.appendChild(createResizeHandle("column", columnIndex));
    headerRow.appendChild(th);
  }

  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (let rowIndex = 0; rowIndex < displayRowCount; rowIndex += 1) {
    const row = sheet.rows[rowIndex] || [];
    const tr = document.createElement("tr");
    tr.dataset.rowIndex = String(rowIndex);
    tr.style.height = `${getRowHeight(sheet, rowIndex)}px`;
    if (rowIndex >= sheet.rows.length) {
      tr.classList.add("is-filler-row");
    }
    const rowHeader = document.createElement("th");
    rowHeader.className = "row-header";
    rowHeader.dataset.rowIndex = String(rowIndex);
    rowHeader.textContent = String(rowIndex + 1);
    rowHeader.style.width = `${rowHeaderWidth}px`;
    rowHeader.appendChild(createResizeHandle("row", rowIndex));
    tr.appendChild(rowHeader);

    for (let columnIndex = 0; columnIndex < displayColumnCount; columnIndex += 1) {
      const td = document.createElement("td");
      const rawValue = row[columnIndex] ?? "";
      if (window.docFormula.isFormula(rawValue)) {
        td.textContent = window.docFormula.evaluateCellDisplay(sheet, rowIndex, columnIndex);
        td.dataset.formula = rawValue;
        td.classList.add("is-formula");
      } else {
        td.textContent = rawValue;
      }
      td.contentEditable = editMode ? "true" : "false";
      td.dataset.rowIndex = String(rowIndex);
      td.dataset.columnIndex = String(columnIndex);
      td.style.width = `${getColumnWidth(sheet, columnIndex)}px`;
      td.style.height = `${getRowHeight(sheet, rowIndex)}px`;
      if (selectedCell?.rowIndex === rowIndex && selectedCell?.columnIndex === columnIndex) {
        td.classList.add("is-selected");
      }
      if (isCellInSelectedRange(rowIndex, columnIndex)) {
        td.classList.add("is-range-selected");
      }
      if (rowIndex >= sheet.rows.length || columnIndex >= actualColumnCount) {
        td.classList.add("is-filler-cell");
      }
      td.addEventListener("pointerdown", (event) => {
        if (event.button !== 0) {
          return;
        }

        event.preventDefault();
        selectWorkbookCell(rowIndex, columnIndex, { extend: event.shiftKey });
        isSelectingRange = editMode;
        td.focus();
      });
      td.addEventListener("pointerenter", () => {
        if (!isSelectingRange || !rangeAnchor) {
          return;
        }

        selectWorkbookCell(rowIndex, columnIndex, { extend: true });
      });
      td.addEventListener("focus", () => {
        if (!selectedCell) {
          selectWorkbookCell(rowIndex, columnIndex);
        }
        if (editMode && td.dataset.formula) {
          // Edit the raw formula, not its computed value.
          td.textContent = td.dataset.formula;
          delete td.dataset.formula;
          td.classList.remove("is-formula");
        }
        if (editMode) {
          // pointerdown's preventDefault (needed for range-drag) suppresses the
          // browser's own caret placement, which strands the caret at the start.
          placeCaretAtEnd(td);
        }
      });
      td.addEventListener("blur", () => {
        if (!editMode || currentDocument?.kind !== "workbook") {
          return;
        }
        syncWorkbookFromDom();
        refreshFormulaCells();
      });
      td.addEventListener("keydown", (event) => {
        if (!editMode) {
          return;
        }
        let rowDelta = 0;
        let columnDelta = 0;
        if (event.key === "Enter") {
          rowDelta = event.shiftKey ? -1 : 1;
        } else if (event.key === "Tab") {
          columnDelta = event.shiftKey ? -1 : 1;
        } else {
          return;
        }
        event.preventDefault();
        moveWorkbookSelection(rowIndex + rowDelta, columnIndex + columnDelta);
      });
      td.addEventListener("click", (event) => {
        event.preventDefault();
      });
      td.addEventListener("input", () => {
        setDirty(true);
      });
      tr.appendChild(td);
    }

    tbody.appendChild(tr);
  }

  table.appendChild(tbody);
  tableWrap.appendChild(table);
  workbookShell.appendChild(tableWrap);
  viewer.replaceChildren(workbookShell);
}

function addSheet() {
  syncWorkbookFromDom();
  const sheets = currentDocument.sheets;
  let counter = sheets.length + 1;
  while (sheets.some((sheet) => sheet.name === `Sheet${counter}`)) {
    counter += 1;
  }
  sheets.push({ name: `Sheet${counter}`, rows: [[""]], columnWidths: [120], rowHeights: [30] });
  selectedSheetIndex = sheets.length - 1;
  selectedCell = null;
  selectedRange = null;
  rangeAnchor = null;
  setDirty(true);
  renderWorkbook(currentDocument);
  setStatus("Sheet added");
}

async function renameSheet(index) {
  const sheet = currentDocument?.sheets[index];
  if (!sheet) {
    return;
  }

  syncWorkbookFromDom();
  const name = await showPrompt({
    title: "Rename sheet",
    message: "Enter a new sheet name.",
    defaultValue: sheet.name,
    primaryLabel: "Rename"
  });
  if (name === null) {
    return;
  }

  const trimmed = name.trim().slice(0, 31); // Excel's sheet-name limits
  if (!trimmed || /[\\/?*[\]:]/.test(trimmed)) {
    setStatus("Invalid sheet name");
    return;
  }
  if (currentDocument.sheets.some((other, otherIndex) => otherIndex !== index && other.name.toLowerCase() === trimmed.toLowerCase())) {
    setStatus("Sheet name already used");
    return;
  }

  sheet.name = trimmed;
  setDirty(true);
  renderWorkbook(currentDocument);
}

async function deleteSheet(index) {
  const sheets = currentDocument?.sheets;
  if (!sheets || !sheets[index]) {
    return;
  }
  if (sheets.length === 1) {
    setStatus("Workbooks need at least one sheet");
    return;
  }

  const confirmed = await showConfirm({
    title: "Delete sheet",
    message: `Delete "${sheets[index].name}" and all of its cells?`,
    tone: "warning",
    primaryLabel: "Delete"
  });
  if (!confirmed) {
    return;
  }

  syncWorkbookFromDom();
  sheets.splice(index, 1);
  if (index < selectedSheetIndex) {
    selectedSheetIndex -= 1;
  } else {
    selectedSheetIndex = Math.min(selectedSheetIndex, sheets.length - 1);
  }
  selectedCell = null;
  selectedRange = null;
  rangeAnchor = null;
  setDirty(true);
  renderWorkbook(currentDocument);
  setStatus("Sheet deleted");
}

function syncWorkbookFromDom() {
  if (!currentDocument || currentDocument.kind !== "workbook") {
    return;
  }

  const sheet = currentDocument.sheets[selectedSheetIndex];
  if (!sheet) {
    return;
  }

  const cells = viewer.querySelectorAll("td[data-row-index][data-column-index]");
  if (cells.length === 0) {
    // The grid for this sheet is not rendered yet (e.g. sync fired on open,
    // before renderWorkbook). There is nothing to read, so don't wipe the rows.
    return;
  }

  const nextRows = [];
  let lastContentRow = -1;
  let lastContentColumn = -1;

  cells.forEach((cell) => {
    const rowIndex = Number(cell.dataset.rowIndex);
    const columnIndex = Number(cell.dataset.columnIndex);
    // Formula cells display their computed value; the raw formula lives in data-formula.
    const value = cell.dataset.formula ?? cell.innerText;
    nextRows[rowIndex] ||= [];
    nextRows[rowIndex][columnIndex] = value;

    if (String(value).trim() !== "") {
      lastContentRow = Math.max(lastContentRow, rowIndex);
      lastContentColumn = Math.max(lastContentColumn, columnIndex);
    }
  });

  const rowCount = Math.max(1, lastContentRow + 1);
  const columnCount = Math.max(1, lastContentColumn + 1);
  sheet.rows = Array.from({ length: rowCount }, (_unused, rowIndex) => {
    const row = nextRows[rowIndex] || [];
    return Array.from({ length: columnCount }, (_unusedColumn, columnIndex) => row[columnIndex] || "");
  });
  sheet.columnWidths = (sheet.columnWidths || []).slice(0, columnCount);
  sheet.rowHeights = (sheet.rowHeights || []).slice(0, rowCount);
}

// Recompute the displayed value of every formula cell without rebuilding the
// grid (a rebuild would steal focus mid-edit).
function refreshFormulaCells() {
  const sheet = currentDocument?.sheets[selectedSheetIndex];
  if (!sheet) {
    return;
  }

  viewer.querySelectorAll("td[data-row-index][data-column-index]").forEach((td) => {
    if (td === document.activeElement) {
      return;
    }

    const rowIndex = Number(td.dataset.rowIndex);
    const columnIndex = Number(td.dataset.columnIndex);
    const raw = sheet.rows[rowIndex]?.[columnIndex] ?? "";
    if (window.docFormula.isFormula(raw)) {
      td.dataset.formula = raw;
      td.classList.add("is-formula");
      td.textContent = window.docFormula.evaluateCellDisplay(sheet, rowIndex, columnIndex);
    } else if (td.dataset.formula) {
      delete td.dataset.formula;
      td.classList.remove("is-formula");
      td.textContent = raw;
    }
  });
}

function placeCaretAtEnd(element) {
  const range = document.createRange();
  range.selectNodeContents(element);
  range.collapse(false);
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);
}

function moveWorkbookSelection(rowIndex, columnIndex) {
  const target = viewer.querySelector(`td[data-row-index="${rowIndex}"][data-column-index="${columnIndex}"]`);
  if (!target) {
    return;
  }

  selectWorkbookCell(rowIndex, columnIndex);
  target.focus();
}

function createResizeHandle(type, index) {
  const handle = document.createElement("span");
  handle.className = type === "column" ? "column-resize-handle" : "row-resize-handle";
  handle.dataset.resizeType = type;
  handle.dataset.resizeIndex = String(index);
  handle.addEventListener("pointerdown", (event) => {
    if (!currentDocument || currentDocument.kind !== "workbook") {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    syncWorkbookFromDom();
    const sheet = currentDocument.sheets[selectedSheetIndex];
    ensureSheetLayout(sheet);

    resizeState = {
      type,
      index,
      startX: event.clientX,
      startY: event.clientY,
      startSize: type === "column" ? getColumnWidth(sheet, index) : getRowHeight(sheet, index)
    };
    document.body.classList.add("is-resizing-sheet");
    handle.setPointerCapture?.(event.pointerId);
  });

  return handle;
}

function resizeActiveSheet(event) {
  const sheet = currentDocument?.sheets[selectedSheetIndex];
  if (!sheet) {
    return;
  }

  ensureSheetLayout(sheet);

  if (resizeState.type === "column") {
    const width = clamp(resizeState.startSize + event.clientX - resizeState.startX, 48, 640);
    sheet.columnWidths[resizeState.index] = width;
    applyColumnWidth(resizeState.index, width);
  } else {
    const height = clamp(resizeState.startSize + event.clientY - resizeState.startY, 22, 220);
    sheet.rowHeights[resizeState.index] = height;
    applyRowHeight(resizeState.index, height);
  }
}

function applyColumnWidth(columnIndex, width) {
  viewer
    .querySelectorAll(`[data-column-index="${columnIndex}"]`)
    .forEach((element) => {
      element.style.width = `${width}px`;
    });

  const col = viewer.querySelector(`col[data-column-index="${columnIndex}"]`);
  if (col) {
    col.style.width = `${width}px`;
  }
}

function applyRowHeight(rowIndex, height) {
  viewer
    .querySelectorAll(`[data-row-index="${rowIndex}"]`)
    .forEach((element) => {
      element.style.height = `${height}px`;
    });
}

function ensureSheetLayout(sheet, columnCount = 1) {
  sheet.columnWidths ||= [];
  sheet.rowHeights ||= [];

  for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
    sheet.columnWidths[columnIndex] ||= defaultColumnWidth(sheet, columnIndex);
  }

  for (let rowIndex = 0; rowIndex < sheet.rows.length; rowIndex += 1) {
    sheet.rowHeights[rowIndex] ||= 30;
  }
}

function getSheetColumnCount(sheet) {
  return Math.max(1, ...sheet.rows.map((row) => row.length), sheet.columnWidths?.length || 0);
}

function getWorkbookDisplayColumnCount(sheet, actualColumnCount = getSheetColumnCount(sheet)) {
  const layout = getDocumentLayout();
  return Math.max(actualColumnCount, WORKBOOK_DISPLAY_COLUMNS[layout] || WORKBOOK_DISPLAY_COLUMNS.narrow);
}

function getWorkbookDisplayRowCount(sheet) {
  const layout = getDocumentLayout();
  return Math.max(sheet.rows.length || 1, WORKBOOK_DISPLAY_ROWS[layout] || WORKBOOK_DISPLAY_ROWS.narrow);
}

function getColumnWidth(sheet, columnIndex) {
  return sheet.columnWidths?.[columnIndex] || defaultColumnWidth(sheet, columnIndex);
}

function getRowHeight(sheet, rowIndex) {
  return sheet.rowHeights?.[rowIndex] || 30;
}

function defaultColumnWidth(sheet, columnIndex) {
  let widest = columnName(columnIndex).length;

  sheet.rows.forEach((row) => {
    widest = Math.max(widest, String(row[columnIndex] ?? "").length);
  });

  return clamp(widest * 7 + 34, 92, 280);
}

function getRowHeaderWidth(rowCount) {
  const digitCount = Math.max(2, String(Math.max(1, rowCount)).length);
  return clamp(digitCount * 8 + 18, 34, 58);
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, Math.round(value)));
}

function selectWorkbookCell(rowIndex, columnIndex, options = {}) {
  const nextCell = { rowIndex, columnIndex };
  if (!options.extend || !rangeAnchor) {
    rangeAnchor = nextCell;
  }

  selectedCell = nextCell;
  selectedRange = normalizeRange(rangeAnchor, nextCell);
  updateSelectionLabel();
  updateRenderedRangeSelection();
}

function normalizeRange(startCell, endCell) {
  return {
    startRow: Math.min(startCell.rowIndex, endCell.rowIndex),
    endRow: Math.max(startCell.rowIndex, endCell.rowIndex),
    startColumn: Math.min(startCell.columnIndex, endCell.columnIndex),
    endColumn: Math.max(startCell.columnIndex, endCell.columnIndex)
  };
}

function isCellInSelectedRange(rowIndex, columnIndex) {
  return Boolean(
    selectedRange &&
      rowIndex >= selectedRange.startRow &&
      rowIndex <= selectedRange.endRow &&
      columnIndex >= selectedRange.startColumn &&
      columnIndex <= selectedRange.endColumn
  );
}

function updateRenderedRangeSelection() {
  viewer.querySelectorAll("td[data-row-index][data-column-index]").forEach((cell) => {
    const rowIndex = Number(cell.dataset.rowIndex);
    const columnIndex = Number(cell.dataset.columnIndex);
    cell.classList.toggle("is-selected", selectedCell?.rowIndex === rowIndex && selectedCell?.columnIndex === columnIndex);
    cell.classList.toggle("is-range-selected", isCellInSelectedRange(rowIndex, columnIndex));
  });
}

function canCopyWorkbookRange() {
  return Boolean(currentDocument?.kind === "workbook" && selectedRange);
}

function canPasteWorkbookRange() {
  return Boolean(currentDocument?.kind === "workbook" && editMode && selectedRange);
}

// Delete/Backspace only hijacks multi-cell ranges; a single focused cell keeps
// native text editing.
function canClearWorkbookRange() {
  return Boolean(currentDocument?.kind === "workbook" && editMode && selectedRange && isMultiCellRange(selectedRange));
}

function clearWorkbookRange() {
  syncWorkbookFromDom();
  const sheet = currentDocument.sheets[selectedSheetIndex];
  const range = selectedRange || normalizeRange(selectedCell, selectedCell);
  for (let rowIndex = range.startRow; rowIndex <= range.endRow; rowIndex += 1) {
    for (let columnIndex = range.startColumn; columnIndex <= range.endColumn; columnIndex += 1) {
      if (sheet.rows[rowIndex] && columnIndex < sheet.rows[rowIndex].length) {
        sheet.rows[rowIndex][columnIndex] = "";
      }
    }
  }
  setDirty(true);
  renderWorkbook(currentDocument);
}

async function copyWorkbookSelection() {
  const text = getSelectedWorkbookText();
  try {
    await navigator.clipboard.writeText(text);
    setStatus("Cells copied");
  } catch {
    setStatus("Copy failed");
  }
}

function getSelectedWorkbookText() {
  if (!selectedRange) {
    return "";
  }

  syncWorkbookFromDom();
  const sheet = currentDocument.sheets[selectedSheetIndex];
  const rows = [];
  for (let rowIndex = selectedRange.startRow; rowIndex <= selectedRange.endRow; rowIndex += 1) {
    const values = [];
    for (let columnIndex = selectedRange.startColumn; columnIndex <= selectedRange.endColumn; columnIndex += 1) {
      values.push(sheet.rows[rowIndex]?.[columnIndex] || "");
    }
    rows.push(values.join("\t"));
  }

  return rows.join("\n");
}

async function pasteWorkbookSelection() {
  let text = "";
  try {
    text = await navigator.clipboard.readText();
  } catch {
    setStatus("Copy failed");
    return;
  }

  const matrix = parseClipboardTable(text);
  if (matrix.length === 0 || matrix[0].length === 0) {
    return;
  }

  syncWorkbookFromDom();
  const sheet = currentDocument.sheets[selectedSheetIndex];
  const targetRange = selectedRange || normalizeRange(selectedCell, selectedCell);

  if (matrix.length === 1 && matrix[0].length === 1 && isMultiCellRange(targetRange)) {
    for (let rowIndex = targetRange.startRow; rowIndex <= targetRange.endRow; rowIndex += 1) {
      for (let columnIndex = targetRange.startColumn; columnIndex <= targetRange.endColumn; columnIndex += 1) {
        setSheetCellValue(sheet, rowIndex, columnIndex, matrix[0][0]);
      }
    }
  } else {
    matrix.forEach((row, rowOffset) => {
      row.forEach((value, columnOffset) => {
        setSheetCellValue(sheet, targetRange.startRow + rowOffset, targetRange.startColumn + columnOffset, value);
      });
    });
    selectedRange = normalizeRange(
      { rowIndex: targetRange.startRow, columnIndex: targetRange.startColumn },
      { rowIndex: targetRange.startRow + matrix.length - 1, columnIndex: targetRange.startColumn + matrix[0].length - 1 }
    );
    selectedCell = { rowIndex: selectedRange.startRow, columnIndex: selectedRange.startColumn };
    rangeAnchor = selectedCell;
  }

  setDirty(true);
  renderWorkbook(currentDocument);
  updateSelectionLabel();
  setStatus("Cells pasted");
}

function parseClipboardTable(text) {
  return String(text || "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/\n$/, "")
    .split("\n")
    .filter((row) => row.length > 0)
    .map((row) => row.split("\t"));
}

function setSheetCellValue(sheet, rowIndex, columnIndex, value) {
  sheet.columnWidths ||= [];
  sheet.rowHeights ||= [];

  while (sheet.rows.length <= rowIndex) {
    sheet.rows.push([]);
  }

  sheet.rows.forEach((row) => {
    while (row.length <= columnIndex) {
      row.push("");
    }
  });

  sheet.rows[rowIndex][columnIndex] = value;
  while (sheet.columnWidths.length <= columnIndex) {
    sheet.columnWidths.push(120);
  }
  while (sheet.rowHeights.length <= rowIndex) {
    sheet.rowHeights.push(30);
  }
}

function isMultiCellRange(range) {
  return range.startRow !== range.endRow || range.startColumn !== range.endColumn;
}

function updateSelectionLabel() {
  if (!selectionLabel) {
    return;
  }

  if (!currentDocument || currentDocument.kind !== "workbook" || !selectedCell || !selectedRange) {
    selectionLabel.textContent = "No cell";
    return;
  }

  const start = `${columnName(selectedRange.startColumn)}${selectedRange.startRow + 1}`;
  const end = `${columnName(selectedRange.endColumn)}${selectedRange.endRow + 1}`;
  selectionLabel.textContent = start === end ? start : `${start}:${end}`;
}

async function runSheetAction(action) {
  if (!currentDocument || currentDocument.kind !== "workbook") {
    return;
  }

  syncWorkbookFromDom();

  const sheet = currentDocument.sheets[selectedSheetIndex];
  if (!sheet) {
    return;
  }

  const actualColumnCount = getSheetColumnCount(sheet);
  const columnCount = Math.max(actualColumnCount, selectedCell ? selectedCell.columnIndex + 1 : 1);
  const rowIndex = selectedRange?.startRow ?? selectedCell?.rowIndex ?? Math.max(0, sheet.rows.length - 1);
  const columnIndex = selectedRange?.startColumn ?? selectedCell?.columnIndex ?? Math.max(0, columnCount - 1);

  if (action === "add-row") {
    const insertAt = Math.min(rowIndex + 1, sheet.rows.length);
    sheet.rows.splice(insertAt, 0, Array(columnCount).fill(""));
    sheet.rowHeights.splice(insertAt, 0, 30);
    selectedCell = { rowIndex: insertAt, columnIndex: 0 };
  } else if (action === "delete-row" && sheet.rows.length > 0 && rowIndex < sheet.rows.length) {
    sheet.rows.splice(rowIndex, 1);
    sheet.rowHeights.splice(rowIndex, 1);
    selectedCell = sheet.rows.length > 0 ? { rowIndex: Math.max(0, rowIndex - 1), columnIndex: 0 } : null;
  } else if (action === "add-column") {
    const insertAt = Math.min(columnIndex + 1, columnCount);
    ensureSheetRows(sheet, columnCount);
    sheet.rows.forEach((row) => row.splice(insertAt, 0, ""));
    sheet.columnWidths.splice(insertAt, 0, 120);
    selectedCell = { rowIndex: Math.max(0, rowIndex), columnIndex: insertAt };
  } else if (action === "delete-column" && actualColumnCount > 1 && columnIndex < actualColumnCount) {
    sheet.rows.forEach((row) => row.splice(columnIndex, 1));
    sheet.columnWidths.splice(columnIndex, 1);
    selectedCell = { rowIndex: Math.max(0, rowIndex), columnIndex: Math.max(0, columnIndex - 1) };
  } else if (action === "clear-cell" && selectedCell && sheet.rows[selectedCell.rowIndex]) {
    sheet.rows[selectedCell.rowIndex][selectedCell.columnIndex] = "";
  } else if (action === "sort-asc" || action === "sort-desc") {
    if (!selectedCell) {
      setStatus("Select a column to sort");
      return;
    }
    // Sorting moves rows but never rewrites cell refs (ranges like A1:A10 can't
    // be rewritten soundly once rows interleave — Excel doesn't try either), so
    // any formula with a ref will point at the wrong row afterwards. Warn first.
    if (sheetHasCellRefFormulas(sheet)) {
      const confirmed = await showConfirm({
        title: "Sort with formulas?",
        message: "This sheet contains formulas with cell references. Sorting moves rows without updating references, so those formulas may point at the wrong cells afterwards.",
        tone: "warning",
        primaryLabel: "Sort anyway"
      });
      if (!confirmed) {
        return;
      }
    }
    sortSheetByColumn(sheet, columnIndex, action === "sort-asc" ? 1 : -1);
  }

  if (selectedCell) {
    rangeAnchor = selectedCell;
    selectedRange = normalizeRange(selectedCell, selectedCell);
  } else {
    rangeAnchor = null;
    selectedRange = null;
  }

  setDirty(true);
  renderWorkbook(currentDocument);
}

function sheetHasCellRefFormulas(sheet) {
  return sheet.rows.some((row) =>
    row.some((value) => typeof value === "string" && value.startsWith("=") && /[A-Z]+\d+/i.test(value))
  );
}

// ponytail: sorts every row (no header detection); formula refs are guarded by
// a confirm dialog in runSheetAction rather than rewritten (ranges can't be).
function sortSheetByColumn(sheet, columnIndex, direction) {
  ensureSheetLayout(sheet);
  const zipped = sheet.rows.map((row, index) => ({ row, height: sheet.rowHeights[index] || 30 }));
  zipped.sort((left, right) => {
    const leftText = String(left.row[columnIndex] ?? "").trim();
    const rightText = String(right.row[columnIndex] ?? "").trim();
    if (leftText === "" || rightText === "") {
      return leftText === rightText ? 0 : leftText === "" ? 1 : -1; // empties last either way
    }
    return direction * compareCellValues(leftText, rightText);
  });
  sheet.rows = zipped.map((entry) => entry.row);
  sheet.rowHeights = zipped.map((entry) => entry.height);
}

function csvTextForSheet(sheet) {
  if (!sheet) {
    return "";
  }
  const field = (value) => {
    const text = String(value ?? "");
    return /[",\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
  };
  return sheet.rows.map((row) => row.map(field).join(",")).join("\n");
}

function compareCellValues(leftText, rightText) {
  const leftNumber = Number(leftText);
  const rightNumber = Number(rightText);
  if (!Number.isNaN(leftNumber) && !Number.isNaN(rightNumber)) {
    return leftNumber - rightNumber;
  }
  return leftText.localeCompare(rightText, undefined, { numeric: true, sensitivity: "base" });
}

function ensureSheetRows(sheet, columnCount) {
  if (sheet.rows.length === 0) {
    sheet.rows.push(Array(columnCount).fill(""));
    return;
  }

  sheet.rows.forEach((row) => {
    while (row.length < columnCount) {
      row.push("");
    }
  });
}

function columnName(index) {
  let name = "";
  let value = index + 1;

  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - remainder) / 26);
  }

  return name;
}
