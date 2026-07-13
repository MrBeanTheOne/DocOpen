const openButton = document.querySelector("#openButton");
const newButton = document.querySelector("#newButton");
const newMenu = document.querySelector("#newMenu");
const newWordButton = document.querySelector("#newWordButton");
const newWorkbookButton = document.querySelector("#newWorkbookButton");
const newMarkdownButton = document.querySelector("#newMarkdownButton");
const newTextButton = document.querySelector("#newTextButton");
const newCsvButton = document.querySelector("#newCsvButton");
const newImageButton = document.querySelector("#newImageButton");
const editButton = document.querySelector("#editButton");
const rawViewButton = document.querySelector("#rawViewButton");
const formatJsonButton = document.querySelector("#formatJsonButton");
const editButtonLabel = document.querySelector("#editButtonLabel");
const saveButton = document.querySelector("#saveButton");
const saveButtonLabel = document.querySelector("#saveButtonLabel");
const saveAsButton = document.querySelector("#saveAsButton");
const revertButton = document.querySelector("#revertButton");
const sidebarToggle = document.querySelector("#sidebarToggle");
const sidebarToggleGlyph = document.querySelector("#sidebarToggleGlyph");
const sidebarRail = document.querySelector("#sidebarRail");
const appShell = document.querySelector("#appShell");
const dropZone = document.querySelector("#dropZone");
const viewer = document.querySelector("#viewer");
const tabStrip = document.querySelector("#tabStrip");
const documentKind = document.querySelector("#documentKind");
const currentFilePanel = document.querySelector("#currentFilePanel");
const fileName = document.querySelector("#fileName");
const fileKindIcon = document.querySelector("#fileKindIcon");
const fileMeta = document.querySelector("#fileMeta");
const documentNotices = document.querySelector("#documentNotices");
const documentInfoPanel = document.querySelector("#documentInfoPanel");
const documentInfoGrid = document.querySelector("#documentInfoGrid");
const backupPanel = document.querySelector("#backupPanel");
const backupList = document.querySelector("#backupList");
const clearBackupsButton = document.querySelector("#clearBackupsButton");
const recentPanel = document.querySelector("#recentPanel");
const recentList = document.querySelector("#recentList");
const clearRecentButton = document.querySelector("#clearRecentButton");
const documentTitle = document.querySelector("#documentTitle");
const modeLabel = document.querySelector("#modeLabel");
const statusText = document.querySelector("#statusText");
const dirtyStatus = document.querySelector("#dirtyStatus");
const toast = document.querySelector("#toast");
const sheetPanel = document.querySelector("#sheetPanel");
const sheetList = document.querySelector("#sheetList");
const editToolbar = document.querySelector("#editToolbar");
const selectionLabel = document.querySelector("#selectionLabel");
const layoutSeg = document.querySelector("#layoutSeg");
const fontFamilySelect = document.querySelector("#fontFamilySelect");
const fontSizeSelect = document.querySelector("#fontSizeSelect");
const fontColorInput = document.querySelector("#fontColorInput");
const highlightColorInput = document.querySelector("#highlightColorInput");
const imageInput = document.querySelector("#imageInput");
const imageColorInput = document.querySelector("#imageColorInput");
const imageSizeSelect = document.querySelector("#imageSizeSelect");
const paintFontSelect = document.querySelector("#paintFontSelect");
const fillToleranceSelect = document.querySelector("#fillToleranceSelect");
const findBar = document.querySelector("#findBar");
const findInput = document.querySelector("#findInput");
const replaceInput = document.querySelector("#replaceInput");
const replaceButton = document.querySelector("#replaceButton");
const replaceAllButton = document.querySelector("#replaceAllButton");
const findCloseButton = document.querySelector("#findCloseButton");
const promptBar = document.querySelector("#promptBar");
const promptLabel = document.querySelector("#promptLabel");
const promptInput = document.querySelector("#promptInput");
const promptOkButton = document.querySelector("#promptOkButton");
const promptCancelButton = document.querySelector("#promptCancelButton");
const tableTools = document.querySelector("#tableTools");
const dialogOverlay = document.querySelector("#dialogOverlay");
const appDialog = document.querySelector("#appDialog");
const dialogIcon = document.querySelector("#dialogIcon");
const dialogTitle = document.querySelector("#dialogTitle");
const dialogMessage = document.querySelector("#dialogMessage");
const dialogInput = document.querySelector("#dialogInput");
const dialogPrimaryButton = document.querySelector("#dialogPrimaryButton");
const dialogSecondaryButton = document.querySelector("#dialogSecondaryButton");
const dialogCancelButton = document.querySelector("#dialogCancelButton");

const RECENT_FILES_KEY = "docopen.recentFiles";
const BACKUP_HISTORY_KEY = "docopen.backupHistory";
const MAX_BACKUP_HISTORY = 18;
const DEFAULT_DOCUMENT_LAYOUT = "narrow";
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

const TOAST_MESSAGES = new Map([
  ["Open", "Opened"],
  ["Open canceled", "Open canceled"],
  ["Could not open file", "Could not open file"],
  ["Path copied", "Path copied"],
  ["Copy failed", "Could not copy path"],
  ["Read-only", "Read-only document"],
  ["Save canceled", "Save canceled"],
  ["Save failed", "Save failed"],
  ["Backup missing", "Backup missing"],
  ["Cells copied", "Cells copied"],
  ["Cells pasted", "Cells pasted"],
  ["Cells cut", "Cells cut"],
  ["Created", "Created"],
  ["Saved", "Saved"],
  ["Saved with backup", "Saved with backup"],
  ["Reverted", "Reverted"]
]);

let currentDocument = null;
let tabs = [];
let activeTabIndex = -1;
let selectedSheetIndex = 0;
let selectedCell = null;
let selectedRange = null;
let rangeAnchor = null;
let isSelectingRange = false;
let editMode = false;
let isDirty = false;
let wordEditor = null;
let savedWordRange = null;
let imageCanvas = null;
let paintTool = "brush";
let paintShapeFill = false;
let paintSelection = null; // canvas-space rect { x, y, width, height }
let paintSelectionBox = null; // marquee overlay element (lives in the paint shell)
let paintTextEditor = null; // in-place text box: { element, x, y, size, color }
let paintUndoStack = [];
let paintRedoStack = [];
let paintZoomReset = null;
let inlinePromptResolver = null;
let lastSavedSnapshot = null;
let resizeState = null;
let documentLayout = DEFAULT_DOCUMENT_LAYOUT;
let dragDepth = 0;
let toastTimer = null;
let activeDialog = null;
let recentFiles = loadRecentFiles();
let backupHistory = loadBackupHistory();

sidebarToggleGlyph.textContent = "<";
sidebarRail.setAttribute("aria-hidden", "true");
renderRecentFiles();
renderBackupHistory();
window.documentOpener.syncRecentFiles?.(recentFiles);

openButton.addEventListener("click", async () => {
  closeNewMenu();
  setStatus("Opening...");
  const result = await window.documentOpener.openDocument();
  handleOpenResult(result);
});

newButton.addEventListener("click", (event) => {
  event.stopPropagation();
  toggleNewMenu();
});

newWordButton.addEventListener("click", async () => {
  await createNewDocument("word");
});

newWorkbookButton.addEventListener("click", async () => {
  await createNewDocument("workbook");
});

newMarkdownButton.addEventListener("click", async () => {
  await createNewDocument("markdown");
});

newTextButton.addEventListener("click", async () => {
  await createNewDocument("text");
});

newCsvButton.addEventListener("click", async () => {
  await createNewDocument("csv");
});

newImageButton.addEventListener("click", async () => {
  closeNewMenu();
  const spec = await showPrompt({
    title: "Image size",
    message: "Enter the canvas size as width x height in pixels.",
    defaultValue: "1280x720",
    primaryLabel: "Create"
  });
  if (spec === null) {
    return;
  }

  const match = /^\s*(\d+)\s*[x×,\s]\s*(\d+)\s*$/i.exec(spec);
  if (!match) {
    setStatus("Enter a size like 1280x720");
    return;
  }

  await createNewDocument("image", {
    width: clampInt(match[1], 8, 4096),
    height: clampInt(match[2], 8, 4096)
  });
});

editButton.addEventListener("click", () => {
  setEditMode(!editMode);
});

saveButton.addEventListener("click", async () => {
  await saveCurrentDocument(false);
});

saveAsButton.addEventListener("click", async () => {
  await saveCurrentDocument(true);
});

revertButton.addEventListener("click", () => {
  revertToSavedSnapshot();
});

// Clicking a section title folds the section (Document info, Recently opened).
document.querySelectorAll(".panel-toggle").forEach((toggle) => {
  const fold = () => toggle.closest("section").classList.toggle("is-collapsed");
  toggle.addEventListener("click", fold);
  toggle.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      fold();
    }
  });
});

clearRecentButton.addEventListener("click", () => {
  recentFiles = [];
  saveRecentFiles();
  renderRecentFiles();
});

clearBackupsButton.addEventListener("click", () => {
  if (!currentDocument?.filePath) {
    backupHistory = [];
  } else {
    backupHistory = backupHistory.filter((item) => !sameFilePath(item.sourcePath, currentDocument.filePath));
  }
  saveBackupHistory();
  renderBackupHistory();
});

layoutSeg.addEventListener("click", (event) => {
  const button = event.target.closest("[data-layout]");
  if (!button) {
    return;
  }

  setDocumentLayout(button.dataset.layout);
});

sidebarToggle.addEventListener("click", () => {
  toggleSidebar();
});

sidebarRail.addEventListener("click", () => {
  toggleSidebar(false);
});

function toggleSidebar(forceCollapsed) {
  const currentCollapsed = appShell.classList.contains("sidebar-collapsed");
  const nextCollapsed = typeof forceCollapsed === "boolean" ? forceCollapsed : !currentCollapsed;
  appShell.classList.toggle("sidebar-collapsed", nextCollapsed);
  sidebarRail.setAttribute("aria-hidden", String(!nextCollapsed));
  sidebarToggle.setAttribute("aria-expanded", String(!nextCollapsed));
  sidebarToggle.setAttribute("aria-label", nextCollapsed ? "Expand sidebar" : "Collapse sidebar");
  sidebarToggleGlyph.textContent = nextCollapsed ? ">" : "<";
}

function toggleNewMenu() {
  const nextHidden = !newMenu.hidden;
  newMenu.hidden = nextHidden;
  newButton.setAttribute("aria-expanded", String(!nextHidden));
}

function closeNewMenu() {
  newMenu.hidden = true;
  newButton.setAttribute("aria-expanded", "false");
}

editToolbar.addEventListener("click", (event) => {
  const button = event.target.closest("button");
  if (!button || !editMode || !currentDocument) {
    return;
  }

  if (button.dataset.command) {
    runWordCommand(button.dataset.command);
  } else if (button.dataset.block) {
    runWordBlockCommand(button.dataset.block);
  } else if (button.dataset.wordAction) {
    runWordAction(button.dataset.wordAction);
  } else if (button.dataset.sheetAction) {
    runSheetAction(button.dataset.sheetAction);
  } else if (button.dataset.imageTool || button.dataset.imageAction) {
    commitPaintText(); // a pending text box lands before any other paint action
    handleImageToolbarButton(button);
  }
});

function handleImageToolbarButton(button) {
  if (button.dataset.imageTool) {
    paintTool = button.dataset.imageTool;
    updatePaintToolButtons();
  } else if (button.dataset.imageAction === "undo") {
    undoPaint();
  } else if (button.dataset.imageAction === "redo") {
    redoPaint();
  } else if (button.dataset.imageAction === "fill-shape") {
    paintShapeFill = !paintShapeFill;
    button.classList.toggle("is-active", paintShapeFill);
  } else if (["flip-h", "flip-v", "rotate"].includes(button.dataset.imageAction)) {
    transformPaintCanvas(button.dataset.imageAction);
  } else if (button.dataset.imageAction === "resize") {
    resizePaintCanvas();
  } else if (button.dataset.imageAction === "cut-selection") {
    cutPaintSelection();
  } else if (button.dataset.imageAction === "crop-selection") {
    cropPaintSelection();
  } else if (button.dataset.imageAction === "clear" && imageCanvas) {
    pushPaintUndo();
    const context = imageCanvas.getContext("2d");
    context.globalCompositeOperation = "source-over";
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, imageCanvas.width, imageCanvas.height);
    setDirty(true);
  }
}

editToolbar.addEventListener("mousedown", (event) => {
  if (event.target.closest("button")) {
    event.preventDefault();
  }
});

editToolbar.addEventListener("change", (event) => {
  if (!editMode || !currentDocument || currentDocument.kind !== "word") {
    return;
  }
  // Paint color/size are read live from the inputs on each stroke; no handler needed.

  if (event.target === fontFamilySelect) {
    runWordCommand("fontName", fontFamilySelect.value);
  } else if (event.target === fontSizeSelect) {
    runWordCommand("fontSize", fontSizeSelect.value);
  } else if (event.target === fontColorInput) {
    runWordCommand("foreColor", fontColorInput.value);
  } else if (event.target === highlightColorInput) {
    runWordCommand("hiliteColor", highlightColorInput.value);
  }
});

rawViewButton.addEventListener("click", () => {
  if (!currentDocument || editMode || !supportsRawView(currentDocument)) {
    return;
  }

  currentDocument.rawView = !currentDocument.rawView;
  if (currentDocument.kind === "workbook") {
    renderWorkbook(currentDocument);
  } else {
    renderMarkdown(currentDocument);
  }
  updateRawViewControls();
});

formatJsonButton.addEventListener("click", () => {
  const textarea = viewer.querySelector(".markdown-editor");
  if (!textarea || !isJsonDocument(currentDocument)) {
    return;
  }

  try {
    textarea.value = JSON.stringify(JSON.parse(textarea.value), null, 2);
    textarea.dispatchEvent(new Event("input", { bubbles: true }));
    setStatus("Formatted");
  } catch (error) {
    showAlert({
      title: "Invalid JSON",
      message: `The document is not valid JSON: ${error.message}`,
      tone: "danger",
      primaryLabel: "OK"
    });
  }
});

findBar.addEventListener("mousedown", (event) => {
  if (event.target.closest("button")) {
    event.preventDefault();
  }
});

findInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    findInDocument(event.shiftKey);
  }
});

replaceInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    replaceCurrentMatch();
  }
});

replaceButton.addEventListener("click", () => replaceCurrentMatch());
replaceAllButton.addEventListener("click", () => replaceAllMatches());
findCloseButton.addEventListener("click", () => toggleFindBar(false));

promptBar.addEventListener("mousedown", (event) => {
  if (event.target.closest("button")) {
    event.preventDefault();
  }
});
promptOkButton.addEventListener("click", () => resolveInlinePrompt(promptInput.value));
promptCancelButton.addEventListener("click", () => resolveInlinePrompt(null));
promptInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    resolveInlinePrompt(promptInput.value);
  } else if (event.key === "Escape") {
    event.preventDefault();
    resolveInlinePrompt(null);
  }
});

dialogPrimaryButton.addEventListener("click", () => resolveAppDialog("primary"));
dialogSecondaryButton.addEventListener("click", () => resolveAppDialog("secondary"));
dialogCancelButton.addEventListener("click", () => resolveAppDialog("cancel"));
dialogInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    resolveAppDialog("primary");
  } else if (event.key === "Escape") {
    event.preventDefault();
    resolveAppDialog("cancel");
  }
});

imageInput.addEventListener("change", () => {
  const file = imageInput.files?.[0];
  imageInput.value = "";
  if (!file || !wordEditor || currentDocument?.kind !== "word") {
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    const imageHtml = `<img src="${reader.result}" style="max-width:100%;height:auto;" />`;
    focusEditorAtSavedRange();
    const inserted = document.execCommand("insertHTML", false, imageHtml);
    if (!inserted) {
      wordEditor.insertAdjacentHTML("beforeend", imageHtml);
    }
    captureWordSelection();
    syncCurrentDocumentFromDom();
    setDirty(true);
  };
  reader.readAsDataURL(file);
});

document.addEventListener("keydown", async (event) => {
  if (event.key === "Escape" && activeDialog) {
    event.preventDefault();
    resolveAppDialog("cancel");
    return;
  }

  if (event.key === "Escape") {
    closeNewMenu();
    isSelectingRange = false;
    toggleFindBar(false);
    resolveInlinePrompt(null);
    clearPaintSelection();
  }

  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "f" && FINDABLE_KINDS.has(currentDocument?.kind)) {
    event.preventDefault();
    toggleFindBar(true);
  } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
    event.preventDefault();
    if (!canSaveCurrentDocument()) {
      return;
    }
    await saveCurrentDocument(false);
  } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "z" && editMode && currentDocument?.kind === "image") {
    event.preventDefault();
    if (event.shiftKey) {
      redoPaint();
    } else {
      undoPaint();
    }
  } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "y" && editMode && currentDocument?.kind === "image") {
    event.preventDefault();
    redoPaint();
  } else if ((event.ctrlKey || event.metaKey) && event.key === "0" && editMode && currentDocument?.kind === "image") {
    event.preventDefault();
    paintZoomReset?.();
  } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "w") {
    event.preventDefault();
    if (tabs.length > 0) {
      await closeTab(activeTabIndex);
    }
  } else if (event.ctrlKey && event.key === "Tab") {
    event.preventDefault();
    if (tabs.length > 1) {
      activateTab((activeTabIndex + (event.shiftKey ? tabs.length - 1 : 1)) % tabs.length);
    }
  } else if ((event.ctrlKey || event.metaKey) && ["c", "x", "v"].includes(event.key.toLowerCase()) && editMode && currentDocument?.kind === "image" && !event.target.closest("input, textarea, [contenteditable]")) {
    const key = event.key.toLowerCase();
    if (key !== "v" && !paintSelection) {
      return; // nothing to copy; leave the event alone
    }
    event.preventDefault();
    if (key === "v") {
      pastePaintClipboard();
    } else if (key === "x") {
      await copyPaintSelection();
      cutPaintSelection();
    } else {
      copyPaintSelection();
    }
  } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c" && canCopyWorkbookRange()) {
    event.preventDefault();
    await copyWorkbookSelection();
  } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "v" && canPasteWorkbookRange()) {
    event.preventDefault();
    await pasteWorkbookSelection();
  } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "x" && canPasteWorkbookRange()) {
    event.preventDefault();
    await copyWorkbookSelection();
    clearWorkbookRange();
    setStatus("Cells cut");
  } else if ((event.key === "Delete" || event.key === "Backspace") && paintSelection && editMode && currentDocument?.kind === "image" && !event.target.closest("input, textarea, [contenteditable]")) {
    event.preventDefault();
    cutPaintSelection();
  } else if ((event.key === "Delete" || event.key === "Backspace") && canClearWorkbookRange()) {
    event.preventDefault();
    clearWorkbookRange();
  }
});

document.addEventListener("click", (event) => {
  if (!newMenu.hidden && !newMenu.contains(event.target) && event.target !== newButton) {
    closeNewMenu();
  }
});

document.addEventListener("pointermove", (event) => {
  if (!resizeState) {
    return;
  }

  event.preventDefault();
  resizeActiveSheet(event);
});

document.addEventListener("pointerup", () => {
  if (!resizeState) {
    isSelectingRange = false;
    return;
  }

  resizeState = null;
  document.body.classList.remove("is-resizing-sheet");
  setDirty(true);
});

window.documentOpener.onOpenedFromSystem(async (result) => {
  handleOpenResult(result);
});

// Tray "New Document" submenu.
window.documentOpener.onNewDocumentRequest?.(async (kind) => {
  await createNewDocument(kind);
});

// A screenshot arrives as a fresh image tab, already in edit mode for annotation.
window.documentOpener.onScreenshot?.(async (payload) => {
  if (typeof payload?.dataUrl !== "string" || !payload.dataUrl.startsWith("data:image/")) {
    return;
  }

  const stamp = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  await createNewDocument("image", {
    dataUrl: payload.dataUrl,
    width: payload.width,
    height: payload.height,
    fileName: `Screenshot ${stamp.getFullYear()}-${pad(stamp.getMonth() + 1)}-${pad(stamp.getDate())} ${pad(stamp.getHours())}${pad(stamp.getMinutes())}${pad(stamp.getSeconds())}.png`
  });
  setStatus("Screenshot captured (also on clipboard)");
});

// The window was closed with unsaved changes; confirm with the in-app dialog.
window.documentOpener.onRequestClose?.(async () => {
  snapshotActiveTab();
  const dirtyTabs = tabs.filter((tab) => tab.isDirty);
  if (dirtyTabs.length === 0) {
    window.documentOpener.confirmClose();
    return;
  }

  const result = await showAppDialog({
    title: "Unsaved changes",
    message:
      dirtyTabs.length === 1
        ? "You have unsaved changes. Save them before closing?"
        : `You have unsaved changes in ${dirtyTabs.length} tabs. Save them before closing?`,
    tone: "warning",
    primaryLabel: "Save",
    secondaryLabel: "Don't save",
    cancelLabel: "Cancel"
  });

  if (result.action === "cancel") {
    window.documentOpener.cancelClose?.(); // Also forgets a pending tray-Quit intent.
    return; // Keep the window open.
  }
  if (result.action === "secondary") {
    window.documentOpener.confirmClose(); // Discard and close.
    return;
  }

  for (const tab of dirtyTabs) {
    const index = tabs.indexOf(tab);
    if (index === -1) {
      continue;
    }
    if (index !== activeTabIndex) {
      activateTab(index);
    }
    await saveCurrentDocument(false);
    if (isDirty) {
      return; // Save canceled or failed — stay open on this tab.
    }
  }
  window.documentOpener.confirmClose();
});

dropZone.addEventListener("dragenter", (event) => {
  if (!hasDraggedFiles(event)) {
    return;
  }

  event.preventDefault();
  dragDepth += 1;
  dropZone.classList.add("is-dragging");
});

dropZone.addEventListener("dragover", (event) => {
  if (!hasDraggedFiles(event)) {
    return;
  }

  event.preventDefault();
  event.dataTransfer.dropEffect = "copy";
  dropZone.classList.add("is-dragging");
});

dropZone.addEventListener("dragleave", () => {
  dragDepth = Math.max(0, dragDepth - 1);
  if (dragDepth === 0) {
    dropZone.classList.remove("is-dragging");
  }
});

dropZone.addEventListener("drop", async (event) => {
  event.preventDefault();
  dragDepth = 0;
  dropZone.classList.remove("is-dragging");

  const file = event.dataTransfer.files[0];
  if (!file) {
    return;
  }

  setStatus("Opening...");
  const result = await window.documentOpener.openDroppedFile(file);
  handleOpenResult(result);
});

function handleOpenResult(result, options = {}) {
  if (!result || result.canceled) {
    setStatus("Ready");
    return;
  }

  if (!result.ok) {
    // Keep whatever is open; a failed open shouldn't nuke the current tab.
    setStatus("Could not open file");
    showAlert({
      title: "Could not open file",
      message: result.error || "The document could not be opened.",
      tone: "danger",
      primaryLabel: "OK"
    });
    return;
  }

  result.layoutWidth ||= documentLayout;
  if (result.kind === "word") {
    result.wordWidth ||= result.layoutWidth;
  }

  // If the file is already open, focus (and refresh, when clean) its tab
  // instead of opening a duplicate.
  const existingIndex = options.reuseTab
    ? -1
    : tabs.findIndex((tab) => tab.document.filePath && result.filePath && sameFilePath(tab.document.filePath, result.filePath));
  if (existingIndex !== -1) {
    snapshotActiveTab();
    activeTabIndex = existingIndex;
    if (!tabs[existingIndex].isDirty) {
      tabs[existingIndex] = createDocumentTab(result);
    }
    presentActiveTab();
    rememberRecentDocument(result);
    setStatus("Open");
    return;
  }

  const tab = createDocumentTab(result);
  if (options.reuseTab && tabs[activeTabIndex]) {
    tabs[activeTabIndex] = tab;
  } else {
    snapshotActiveTab();
    tabs.push(tab);
    activeTabIndex = tabs.length - 1;
  }
  presentActiveTab();
  rememberRecentDocument(result);
  setStatus("Open");
}

function createDocumentTab(documentData) {
  return {
    document: documentData,
    editMode: false,
    isDirty: false,
    lastSavedSnapshot: cloneDocument(documentData),
    selectedSheetIndex: 0,
    selectedCell: null,
    selectedRange: null,
    rangeAnchor: null
  };
}

// Fold the live editing state back into the active tab before leaving it.
function snapshotActiveTab() {
  const tab = tabs[activeTabIndex];
  if (!tab) {
    return;
  }

  syncCurrentDocumentFromDom();
  Object.assign(tab, {
    document: currentDocument,
    editMode,
    isDirty,
    lastSavedSnapshot,
    selectedSheetIndex,
    selectedCell,
    selectedRange,
    rangeAnchor
  });
}

// Load the active tab's state into the working globals and render it.
function presentActiveTab() {
  const tab = tabs[activeTabIndex] || null;
  wordEditor = null;
  savedWordRange = null;
  imageCanvas = null;
  toggleFindBar(false);

  if (!tab) {
    currentDocument = null;
    selectedSheetIndex = 0;
    selectedCell = null;
    selectedRange = null;
    rangeAnchor = null;
    setEditMode(false, { sync: false });
    setDirty(false);
    renderError("No document open");
    renderTabStrip();
    return;
  }

  currentDocument = tab.document;
  selectedSheetIndex = tab.selectedSheetIndex;
  selectedCell = tab.selectedCell;
  selectedRange = tab.selectedRange;
  rangeAnchor = tab.rangeAnchor;
  lastSavedSnapshot = tab.lastSavedSnapshot;
  fileName.textContent = currentDocument.fileName;
  fileName.title = currentDocument.fileName;
  documentTitle.textContent = currentDocument.fileName;
  applyDocumentIdentity(currentDocument);
  enableDocumentActions(true);
  setEditMode(tab.editMode, { sync: false }); // renders the document
  setDirty(tab.isDirty);
  renderTabStrip();
}

function activateTab(index) {
  if (index === activeTabIndex || !tabs[index]) {
    return;
  }

  snapshotActiveTab();
  activeTabIndex = index;
  presentActiveTab();
}

async function closeTab(index) {
  const tab = tabs[index];
  if (!tab) {
    return;
  }

  if (tab.isDirty) {
    const confirmed = await showConfirm({
      title: "Unsaved changes",
      message: `Close "${tab.document.fileName}" without saving?`,
      tone: "warning",
      primaryLabel: "Close tab",
      cancelLabel: "Keep open"
    });
    if (!confirmed) {
      return;
    }
  }

  tabs.splice(index, 1);
  if (index < activeTabIndex) {
    activeTabIndex -= 1;
    renderTabStrip();
  } else if (index === activeTabIndex) {
    activeTabIndex = Math.min(index, tabs.length - 1);
    presentActiveTab();
  } else {
    renderTabStrip();
  }
  window.documentOpener.notifyDirty?.(tabs.some((item) => item.isDirty));
}

function renderTabStrip() {
  if (!tabStrip) {
    return;
  }

  tabStrip.hidden = tabs.length === 0;
  document.body.classList.toggle("has-tabs", tabs.length > 0);
  tabStrip.replaceChildren(
    ...tabs.map((tab, index) => {
      const identity = documentIdentity(tab.document);
      const button = document.createElement("button");
      button.type = "button";
      button.className = index === activeTabIndex ? "doc-tab is-active" : "doc-tab";
      button.setAttribute("role", "tab");
      button.setAttribute("aria-selected", String(index === activeTabIndex));
      button.title = tab.document.filePath || tab.document.fileName;
      button.addEventListener("click", () => activateTab(index));
      button.addEventListener("auxclick", (event) => {
        if (event.button === 1) {
          event.preventDefault();
          closeTab(index);
        }
      });

      const icon = document.createElement("span");
      icon.className = `tab-kind-icon ${identity.className}`;
      icon.textContent = identity.icon;

      const name = document.createElement("span");
      name.className = "tab-name";
      name.textContent = tab.document.fileName;

      const dirtyDot = document.createElement("span");
      dirtyDot.className = "tab-dirty";
      dirtyDot.textContent = "•";
      dirtyDot.hidden = !tab.isDirty;

      const close = document.createElement("span");
      close.className = "tab-close";
      close.textContent = "×";
      close.title = "Close tab";
      close.setAttribute("role", "button");
      close.setAttribute("aria-label", `Close ${tab.document.fileName}`);
      close.addEventListener("click", (event) => {
        event.stopPropagation();
        closeTab(index);
      });

      button.append(icon, name, dirtyDot, close);
      return button;
    })
  );
}

async function createNewDocument(kind, creatorOptions) {
  closeNewMenu();

  const blankCreators = {
    word: createBlankWordDocument,
    markdown: createBlankMarkdownDocument,
    text: createBlankTextDocument,
    workbook: createBlankWorkbookDocument,
    csv: createBlankCsvDocument,
    image: createBlankImageDocument
  };
  const documentData = (blankCreators[kind] || createBlankWorkbookDocument)(creatorOptions);
  snapshotActiveTab();
  const tab = createDocumentTab(documentData);
  tab.editMode = true;
  tab.isDirty = true;
  tabs.push(tab);
  activeTabIndex = tabs.length - 1;
  presentActiveTab();
  setStatus("Created");
}

function createBlankWordDocument() {
  return {
    ok: true,
    kind: "word",
    filePath: "",
    fileName: "Untitled document.docx",
    extension: ".docx",
    html: "<h1>Untitled document</h1><p></p>",
    warnings: [],
    isNew: true,
    layoutWidth: documentLayout,
    wordWidth: documentLayout
  };
}

function createBlankMarkdownDocument() {
  return {
    ok: true,
    kind: "markdown",
    filePath: "",
    fileName: "Untitled.md",
    extension: ".md",
    text: "# Untitled\n\n",
    isNew: true,
    layoutWidth: documentLayout
  };
}

function createBlankTextDocument() {
  return {
    ok: true,
    kind: "text",
    filePath: "",
    fileName: "Untitled.txt",
    extension: ".txt",
    text: "",
    isNew: true,
    layoutWidth: documentLayout
  };
}

function createBlankWorkbookDocument() {
  return {
    ok: true,
    kind: "workbook",
    filePath: "",
    fileName: "Untitled workbook.xlsx",
    extension: ".xlsx",
    isNew: true,
    layoutWidth: documentLayout,
    sheets: [
      {
        name: "Sheet1",
        rows: [[""]],
        columnWidths: [120],
        rowHeights: [30]
      }
    ]
  };
}

function createBlankCsvDocument() {
  const workbookData = createBlankWorkbookDocument();
  workbookData.fileName = "Untitled.csv";
  workbookData.extension = ".csv";
  return workbookData;
}

function createBlankImageDocument(options = {}) {
  return {
    ok: true,
    kind: "image",
    filePath: "",
    fileName: options.fileName || "Untitled image.png",
    extension: ".png",
    isNew: true,
    dataUrl: options.dataUrl || "",
    width: options.width || 1280,
    height: options.height || 720,
    layoutWidth: documentLayout
  };
}

function renderWord(docData) {
  sheetPanel.hidden = true;
  sheetList.replaceChildren();
  const safeHtml = sanitizeWordHtml(docData.html || "<p>This document did not contain previewable text.</p>");
  if (safeHtml !== docData.html) {
    docData.html = safeHtml;
    docData.safetyWarnings = ["Potentially unsafe document content was removed from the preview."];
  }
  renderDocumentNotices(docData);

  const article = document.createElement("article");
  article.className = `word-document document-surface width-${getDocumentLayout(docData)}`;
  article.contentEditable = editMode ? "true" : "false";
  article.spellcheck = true;
  article.innerHTML = safeHtml;
  article.addEventListener("focus", () => {
    selectedCell = null;
    selectedRange = null;
    rangeAnchor = null;
    captureWordSelection();
    updateSelectionLabel();
  });
  article.addEventListener("mouseup", captureWordSelection);
  article.addEventListener("keyup", captureWordSelection);
  article.addEventListener("input", () => {
    captureWordSelection();
    setDirty(true);
  });

  wordEditor = article;

  viewer.replaceChildren(article);
}

function renderMarkdown(docData) {
  wordEditor = null;
  savedWordRange = null;
  sheetPanel.hidden = true;
  sheetList.replaceChildren();
  selectedCell = null;
  selectedRange = null;
  rangeAnchor = null;
  updateSelectionLabel();

  if (editMode) {
    const editor = document.createElement("textarea");
    editor.className = `markdown-editor document-surface width-${getDocumentLayout(docData)}`;
    editor.spellcheck = true;
    editor.value = docData.text || "";
    editor.addEventListener("input", () => {
      docData.text = editor.value;
      setDirty(true);
    });
    viewer.replaceChildren(editor);
    editor.focus();
    return;
  }

  if (docData.kind === "text" || docData.rawView) {
    const pre = document.createElement("pre");
    pre.className = `text-document document-surface width-${getDocumentLayout(docData)}`;
    // JSON defaults to a pretty, highlighted view; the Raw toggle (or
    // unparseable JSON) falls back to the text exactly as on disk.
    const highlighted = isJsonDocument(docData) && !docData.rawView ? highlightJson(docData.text || "") : null;
    if (highlighted !== null) {
      pre.classList.add("json-view");
      pre.innerHTML = highlighted;
    } else {
      pre.textContent = docData.text || "";
    }
    viewer.replaceChildren(pre);
    return;
  }

  const article = document.createElement("article");
  article.className = `markdown-document document-surface width-${getDocumentLayout(docData)}`;
  article.innerHTML = markdownToHtml(docData.text || "");
  viewer.replaceChildren(article);
}

// Pretty-print + tokenize JSON into highlight spans. Returns null when the
// text isn't valid JSON so callers fall back to the plain view.
function highlightJson(text) {
  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch {
    return null;
  }

  const pretty = JSON.stringify(parsed, null, 2);
  const token = /("(?:\\.|[^"\\])*")(\s*:)?|\b(?:true|false)\b|\bnull\b|-?\b\d+(?:\.\d+)?(?:[eE][+-]?\d+)?\b/g;
  let html = "";
  let last = 0;
  let match;
  while ((match = token.exec(pretty))) {
    html += escapeHtml(pretty.slice(last, match.index));
    if (match[2]) {
      html += `<span class="json-key">${escapeHtml(match[1])}</span>${match[2]}`;
    } else {
      const value = match[0];
      const cls = match[1] ? "json-string" : value === "null" ? "json-null" : value === "true" || value === "false" ? "json-boolean" : "json-number";
      html += `<span class="${cls}">${escapeHtml(value)}</span>`;
    }
    last = match.index + match[0].length;
  }
  html += escapeHtml(pretty.slice(last));
  return html;
}

let imageMenuCleanup = null;

function closeImageContextMenu() {
  if (imageMenuCleanup) {
    imageMenuCleanup();
    imageMenuCleanup = null;
  }
}

// Copies the whole image (edited canvas state when painting, otherwise the
// document) to the OS clipboard.
async function copyWholeImage() {
  if (currentDocument?.kind !== "image") {
    return;
  }

  let result = null;
  if (imageCanvas) {
    result = await window.documentOpener.copyImageToClipboard(imageCanvas.toDataURL("image/png"));
  } else if (currentDocument.dataUrl) {
    result = await window.documentOpener.copyImageToClipboard(currentDocument.dataUrl);
  } else if (currentDocument.filePath) {
    result = await window.documentOpener.copyImageFileToClipboard(currentDocument.filePath);
  }
  setStatus(result?.ok ? "Image copied to clipboard" : "Could not copy this image");
}

function showImageContextMenu(event) {
  event.preventDefault();
  closeImageContextMenu();

  const menu = document.createElement("div");
  menu.className = "new-menu";
  menu.style.position = "fixed";
  menu.style.left = `${event.clientX}px`;
  menu.style.top = `${event.clientY}px`;
  menu.style.right = "auto";
  menu.style.width = "170px";

  const copyItem = document.createElement("button");
  copyItem.type = "button";
  copyItem.className = "new-menu-item";
  copyItem.style.gridTemplateColumns = "1fr";
  copyItem.style.padding = "0 12px";
  copyItem.textContent = "Copy image";
  copyItem.addEventListener("click", async () => {
    closeImageContextMenu();
    await copyWholeImage();
  });
  menu.appendChild(copyItem);
  document.body.appendChild(menu);

  // Keep the menu inside the window near the edges.
  const rect = menu.getBoundingClientRect();
  if (rect.right > window.innerWidth) {
    menu.style.left = `${Math.max(0, window.innerWidth - rect.width - 4)}px`;
  }
  if (rect.bottom > window.innerHeight) {
    menu.style.top = `${Math.max(0, window.innerHeight - rect.height - 4)}px`;
  }

  const dismiss = (dismissEvent) => {
    if (!menu.contains(dismissEvent.target)) {
      closeImageContextMenu();
    }
  };
  const onKey = (keyEvent) => {
    if (keyEvent.key === "Escape") {
      closeImageContextMenu();
    }
  };
  document.addEventListener("pointerdown", dismiss, true);
  document.addEventListener("keydown", onKey, true);
  imageMenuCleanup = () => {
    document.removeEventListener("pointerdown", dismiss, true);
    document.removeEventListener("keydown", onKey, true);
    menu.remove();
  };
}

function renderImageDocument(imageData) {
  wordEditor = null;
  savedWordRange = null;
  toggleFindBar(false);
  selectedCell = null;
  selectedRange = null;
  rangeAnchor = null;
  updateSelectionLabel();
  sheetPanel.hidden = true;
  sheetList.replaceChildren();

  if (editMode) {
    renderImagePaint(imageData);
    return;
  }

  imageCanvas = null;
  const shell = document.createElement("div");
  shell.className = "image-document";

  const image = document.createElement("img");
  image.className = "image-frame";
  image.alt = imageData.fileName;
  image.addEventListener("load", () => image.classList.add("is-loaded"));
  image.addEventListener("contextmenu", showImageContextMenu);
  image.src = imageData.dataUrl || imageData.fileUrl;

  let zoom = null; // null = fit to view

  const applyZoom = (nextZoom, anchorX, anchorY) => {
    if (!image.naturalWidth) {
      return;
    }

    const rect = image.getBoundingClientRect();
    zoom = Math.min(8, Math.max(0.1, nextZoom));
    image.classList.add("is-zoomed");
    image.style.width = `${image.naturalWidth * zoom}px`;

    // Keep the point under the cursor stable while zooming.
    const nextRect = image.getBoundingClientRect();
    const anchorRatioX = rect.width ? (anchorX - rect.left) / rect.width : 0.5;
    const anchorRatioY = rect.height ? (anchorY - rect.top) / rect.height : 0.5;
    viewer.scrollLeft += nextRect.left - rect.left + anchorRatioX * (nextRect.width - rect.width);
    viewer.scrollTop += nextRect.top - rect.top + anchorRatioY * (nextRect.height - rect.height);
  };

  shell.addEventListener(
    "wheel",
    (event) => {
      if (!event.ctrlKey) {
        return;
      }

      event.preventDefault();
      const factor = event.deltaY < 0 ? 1.12 : 1 / 1.12;
      const currentZoom = zoom ?? image.clientWidth / image.naturalWidth;
      applyZoom(currentZoom * factor, event.clientX, event.clientY);
    },
    { passive: false }
  );

  image.addEventListener("dblclick", () => {
    zoom = null;
    image.classList.remove("is-zoomed");
    image.style.width = "";
  });

  shell.appendChild(image);
  viewer.replaceChildren(shell);
}

function renderImagePaint(imageData) {
  const shell = document.createElement("div");
  shell.className = "image-document is-painting";

  const canvas = document.createElement("canvas");
  canvas.className = "paint-canvas";
  canvas.addEventListener("contextmenu", showImageContextMenu);
  const context = canvas.getContext("2d");

  const selectionBox = document.createElement("div");
  selectionBox.className = "paint-selection";
  selectionBox.hidden = true;

  // Bottom-right drag grip: resizes the canvas like MS Paint's corner handle.
  const resizeGrip = document.createElement("div");
  resizeGrip.className = "paint-resize-grip";
  resizeGrip.title = "Drag to resize canvas";
  const resizeGhost = document.createElement("div");
  resizeGhost.className = "paint-resize-ghost";
  resizeGhost.hidden = true;

  const positionResizeGrip = () => {
    resizeGrip.style.left = `${canvas.offsetLeft + canvas.clientWidth}px`;
    resizeGrip.style.top = `${canvas.offsetTop + canvas.clientHeight}px`;
  };

  let gripDrag = null;
  resizeGrip.addEventListener("pointerdown", (event) => {
    if (event.button !== 0) {
      return;
    }
    event.preventDefault();
    commitPaintText();
    gripDrag = {
      startX: event.clientX,
      startY: event.clientY,
      width: canvas.width,
      height: canvas.height,
      scale: canvas.clientWidth / canvas.width || 1
    };
    resizeGrip.setPointerCapture?.(event.pointerId);
  });
  resizeGrip.addEventListener("pointermove", (event) => {
    if (!gripDrag) {
      return;
    }
    const width = clampInt(gripDrag.width + (event.clientX - gripDrag.startX) / gripDrag.scale, 8, 4096);
    const height = clampInt(gripDrag.height + (event.clientY - gripDrag.startY) / gripDrag.scale, 8, 4096);
    gripDrag.next = { width, height };
    resizeGhost.hidden = false;
    resizeGhost.style.left = `${canvas.offsetLeft}px`;
    resizeGhost.style.top = `${canvas.offsetTop}px`;
    resizeGhost.style.width = `${width * gripDrag.scale}px`;
    resizeGhost.style.height = `${height * gripDrag.scale}px`;
    setStatus(`${width} × ${height}`);
  });
  const endGripDrag = (commit) => {
    if (!gripDrag) {
      return;
    }
    resizeGhost.hidden = true;
    if (commit && gripDrag.next) {
      applyPaintCanvasSize(gripDrag.next.width, gripDrag.next.height);
    }
    gripDrag = null;
  };
  resizeGrip.addEventListener("pointerup", () => endGripDrag(true));
  resizeGrip.addEventListener("pointercancel", () => endGripDrag(false));

  // Repositions the grip on zoom, canvas resize, image load, and layout changes.
  new ResizeObserver(positionResizeGrip).observe(canvas);

  const prepareSurface = (width, height, source) => {
    canvas.width = width;
    canvas.height = height;
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, width, height);
    if (source) {
      context.drawImage(source, 0, 0);
    }
  };

  const source = imageData.dataUrl || imageData.fileUrl;
  if (source) {
    const image = new Image();
    image.onload = () => prepareSurface(image.naturalWidth, image.naturalHeight, image);
    image.src = source;
    // Give the canvas its expected size up front so layout doesn't jump.
    prepareSurface(imageData.width || 1280, imageData.height || 720);
  } else {
    prepareSurface(imageData.width || 1280, imageData.height || 720);
  }

  let stroke = null; // active drag: { startX, startY, before: ImageData }
  let zoom = null; // null = fit to view via CSS max-width

  const canvasPoint = (event) => {
    const rect = canvas.getBoundingClientRect();
    return {
      x: ((event.clientX - rect.left) * canvas.width) / rect.width,
      y: ((event.clientY - rect.top) * canvas.height) / rect.height
    };
  };

  // Drawing coordinates always map through the rendered rect, so zoom only
  // changes the CSS size — strokes stay pixel-accurate at any zoom level.
  shell.addEventListener(
    "wheel",
    (event) => {
      if (!event.ctrlKey) {
        return;
      }

      event.preventDefault();
      commitPaintText(); // zoom moves the canvas box; land pending text first
      const rect = canvas.getBoundingClientRect();
      if (!rect.width) {
        return;
      }
      const factor = event.deltaY < 0 ? 1.12 : 1 / 1.12;
      zoom = Math.min(8, Math.max(0.1, (zoom ?? rect.width / canvas.width) * factor));
      canvas.classList.add("is-zoomed");
      canvas.style.width = `${canvas.width * zoom}px`;

      // Keep the point under the cursor stable while zooming.
      const nextRect = canvas.getBoundingClientRect();
      const anchorRatioX = (event.clientX - rect.left) / rect.width;
      const anchorRatioY = rect.height ? (event.clientY - rect.top) / rect.height : 0.5;
      viewer.scrollLeft += nextRect.left - rect.left + anchorRatioX * (nextRect.width - rect.width);
      viewer.scrollTop += nextRect.top - rect.top + anchorRatioY * (nextRect.height - rect.height);
      positionPaintSelectionBox();
    },
    { passive: false }
  );

  paintZoomReset = () => {
    zoom = null;
    canvas.classList.remove("is-zoomed");
    canvas.style.width = "";
    positionPaintSelectionBox();
  };

  canvas.addEventListener("pointerdown", (event) => {
    if (event.button !== 0) {
      return;
    }

    event.preventDefault();
    const point = canvasPoint(event);

    if (paintTextEditor) {
      commitPaintText(); // the click that commits doesn't also draw, like MS Paint
      return;
    }

    if (paintTool === "select") {
      if (paintSelectionContains(point)) {
        // Lift the selected pixels so the drag moves them (MS Paint style).
        const rect = paintSelection;
        pushPaintUndo(); // snapshots the pre-move state (and clears the marquee)
        const lifted = context.getImageData(rect.x, rect.y, rect.width, rect.height);
        clearPaintRect(context, rect);
        const base = context.getImageData(0, 0, canvas.width, canvas.height);
        context.putImageData(lifted, rect.x, rect.y); // visually unchanged until dragged
        paintSelection = rect;
        positionPaintSelectionBox();
        stroke = { startX: point.x, startY: point.y, select: true, move: { rect, lifted, base } };
      } else {
        clearPaintSelection();
        stroke = { startX: point.x, startY: point.y, select: true };
      }
      canvas.setPointerCapture?.(event.pointerId);
      return;
    }

    clearPaintSelection();

    if (paintTool === "picker") {
      const pixel = context.getImageData(
        Math.min(canvas.width - 1, Math.max(0, Math.floor(point.x))),
        Math.min(canvas.height - 1, Math.max(0, Math.floor(point.y))),
        1,
        1
      ).data;
      imageColorInput.value = `#${[pixel[0], pixel[1], pixel[2]].map((v) => v.toString(16).padStart(2, "0")).join("")}`;
      paintTool = "brush"; // like MS Paint, picking hands you back the brush
      updatePaintToolButtons();
      return;
    }

    if (paintTool === "text") {
      openPaintTextEditor(point.x, point.y);
      return;
    }

    const before = context.getImageData(0, 0, canvas.width, canvas.height);

    if (paintTool === "fill") {
      pushPaintUndo(before);
      floodFill(context, Math.floor(point.x), Math.floor(point.y), imageColorInput.value);
      setDirty(true);
      return;
    }

    stroke = { startX: point.x, startY: point.y, before };
    canvas.setPointerCapture?.(event.pointerId);
    context.lineCap = "round";
    context.lineJoin = "round";
    context.lineWidth = Number(imageSizeSelect.value) || 4;
    if (paintTool === "eraser") {
      // Erase to transparency on PNG; JPEG has no alpha, so erase to white there.
      const isJpeg = imageData.extension === ".jpg" || imageData.extension === ".jpeg";
      context.globalCompositeOperation = isJpeg ? "source-over" : "destination-out";
      context.strokeStyle = "#ffffff";
    } else {
      context.globalCompositeOperation = "source-over";
      context.strokeStyle = imageColorInput.value;
    }

    if (paintTool === "brush" || paintTool === "eraser") {
      context.beginPath();
      context.moveTo(point.x, point.y);
      context.lineTo(point.x + 0.01, point.y);
      context.stroke();
    }
  });

  canvas.addEventListener("pointermove", (event) => {
    if (!stroke) {
      if (paintTool === "select") {
        canvas.style.cursor = paintSelectionContains(canvasPoint(event)) ? "move" : "";
      }
      return;
    }

    let point = canvasPoint(event);
    if (stroke.move) {
      const { rect, lifted, base } = stroke.move;
      const x = Math.round(Math.min(canvas.width - rect.width, Math.max(0, rect.x + point.x - stroke.startX)));
      const y = Math.round(Math.min(canvas.height - rect.height, Math.max(0, rect.y + point.y - stroke.startY)));
      context.putImageData(base, 0, 0);
      context.putImageData(lifted, x, y);
      paintSelection = { x, y, width: rect.width, height: rect.height };
      positionPaintSelectionBox();
      return;
    }

    if (stroke.select) {
      paintSelection = normalizePaintRect(stroke.startX, stroke.startY, point.x, point.y);
      positionPaintSelectionBox();
      return;
    }

    if (paintTool === "brush" || paintTool === "eraser") {
      context.lineTo(point.x, point.y);
      context.stroke();
      return;
    }

    if (event.shiftKey) {
      point = constrainShapePoint(paintTool, stroke.startX, stroke.startY, point.x, point.y);
    }

    // Shape preview: restore the pre-drag pixels, then draw the shape so far.
    context.putImageData(stroke.before, 0, 0);
    drawPaintShape(context, paintTool, stroke.startX, stroke.startY, point.x, point.y, paintShapeFill);
  });

  canvas.addEventListener("pointerup", () => {
    if (!stroke) {
      return;
    }

    if (stroke.select) {
      if (stroke.move) {
        setDirty(true); // the moved pixels are already stamped at the drop position
      } else if (!paintSelection || paintSelection.width < 2 || paintSelection.height < 2) {
        // A click without a real drag just deselects.
        clearPaintSelection();
      }
      stroke = null;
      return;
    }

    pushPaintUndo(stroke.before);
    stroke = null;
    setDirty(true);
  });

  imageCanvas = canvas;
  paintSelection = null;
  paintSelectionBox = selectionBox;
  paintTextEditor = null; // any prior text box died with its shell
  paintUndoStack = [];
  paintRedoStack = [];
  updatePaintToolButtons();
  shell.appendChild(canvas);
  shell.appendChild(selectionBox);
  shell.appendChild(resizeGhost);
  shell.appendChild(resizeGrip);
  viewer.replaceChildren(shell);
  positionResizeGrip();
}

function normalizePaintRect(x1, y1, x2, y2) {
  const left = Math.max(0, Math.floor(Math.min(x1, x2)));
  const top = Math.max(0, Math.floor(Math.min(y1, y2)));
  const right = Math.min(imageCanvas.width, Math.ceil(Math.max(x1, x2)));
  const bottom = Math.min(imageCanvas.height, Math.ceil(Math.max(y1, y2)));
  return { x: left, y: top, width: Math.max(0, right - left), height: Math.max(0, bottom - top) };
}

// The marquee is a shell-positioned overlay div, so it maps canvas-space
// coordinates through the rendered (possibly zoomed) canvas box.
function positionPaintSelectionBox() {
  if (!paintSelection || !paintSelectionBox || !imageCanvas) {
    return;
  }
  const scaleX = imageCanvas.clientWidth / imageCanvas.width;
  const scaleY = imageCanvas.clientHeight / imageCanvas.height;
  paintSelectionBox.hidden = false;
  paintSelectionBox.style.left = `${imageCanvas.offsetLeft + paintSelection.x * scaleX}px`;
  paintSelectionBox.style.top = `${imageCanvas.offsetTop + paintSelection.y * scaleY}px`;
  paintSelectionBox.style.width = `${paintSelection.width * scaleX}px`;
  paintSelectionBox.style.height = `${paintSelection.height * scaleY}px`;
}

function clearPaintSelection() {
  paintSelection = null;
  if (paintSelectionBox) {
    paintSelectionBox.hidden = true;
  }
}

function paintSelectionContains(point) {
  return Boolean(
    paintSelection &&
      point.x >= paintSelection.x &&
      point.x <= paintSelection.x + paintSelection.width &&
      point.y >= paintSelection.y &&
      point.y <= paintSelection.y + paintSelection.height
  );
}

// JPEG has no alpha, so cleared areas go white; PNG keeps transparency.
function clearPaintRect(context, rect) {
  if (currentDocument?.extension === ".jpg" || currentDocument?.extension === ".jpeg") {
    context.globalCompositeOperation = "source-over";
    context.fillStyle = "#ffffff";
    context.fillRect(rect.x, rect.y, rect.width, rect.height);
  } else {
    context.clearRect(rect.x, rect.y, rect.width, rect.height);
  }
}

function cutPaintSelection() {
  if (!paintSelection || !imageCanvas || !viewer.contains(imageCanvas)) {
    return;
  }

  const rect = paintSelection;
  pushPaintUndo(); // also clears the marquee
  clearPaintRect(imageCanvas.getContext("2d"), rect);
  setDirty(true);
}

// Copies the selection to the OS clipboard as a PNG, so it pastes into other apps.
async function copyPaintSelection() {
  if (!paintSelection || !imageCanvas) {
    return;
  }
  const { x, y, width, height } = paintSelection;
  const block = document.createElement("canvas");
  block.width = width;
  block.height = height;
  block.getContext("2d").putImageData(imageCanvas.getContext("2d").getImageData(x, y, width, height), 0, 0);
  await window.documentOpener.copyImageToClipboard(block.toDataURL("image/png"));
  setStatus("Selection copied");
}

// Pastes the OS clipboard image at the top-left as a fresh selection so it can
// be dragged into place. Accepts images copied from any app.
async function pastePaintClipboard() {
  if (!imageCanvas || !viewer.contains(imageCanvas)) {
    return;
  }

  const dataUrl = await window.documentOpener.readImageFromClipboard();
  if (!dataUrl) {
    setStatus("Clipboard has no image");
    return;
  }
  if (!imageCanvas || !viewer.contains(imageCanvas)) {
    return;
  }

  const image = new Image();
  await new Promise((resolve, reject) => {
    image.onload = resolve;
    image.onerror = reject;
    image.src = dataUrl;
  });

  pushPaintUndo();
  const context = imageCanvas.getContext("2d");
  context.globalCompositeOperation = "source-over";
  context.drawImage(image, 0, 0);
  paintSelection = {
    x: 0,
    y: 0,
    width: Math.min(image.naturalWidth, imageCanvas.width),
    height: Math.min(image.naturalHeight, imageCanvas.height)
  };
  positionPaintSelectionBox();
  paintTool = "select";
  updatePaintToolButtons();
  setDirty(true);
}

function cropPaintSelection() {
  if (!paintSelection || !imageCanvas || !viewer.contains(imageCanvas)) {
    return;
  }

  const { x, y, width, height } = paintSelection;
  const kept = imageCanvas.getContext("2d").getImageData(x, y, width, height);
  pushPaintUndo(); // also clears the marquee
  imageCanvas.width = width;
  imageCanvas.height = height;
  imageCanvas.getContext("2d").putImageData(kept, 0, 0);
  paintZoomReset?.();
  syncPaintDocumentSize();
  setDirty(true);
}

function updatePaintToolButtons() {
  editToolbar.querySelectorAll("[data-image-tool]").forEach((toolButton) => {
    toolButton.classList.toggle("is-active", toolButton.dataset.imageTool === paintTool);
  });
}

function drawPaintShape(context, tool, startX, startY, endX, endY, filled) {
  context.beginPath();
  if (tool === "line") {
    context.moveTo(startX, startY);
    context.lineTo(endX, endY);
  } else if (tool === "rect") {
    context.rect(Math.min(startX, endX), Math.min(startY, endY), Math.abs(endX - startX), Math.abs(endY - startY));
  } else if (tool === "ellipse") {
    context.ellipse((startX + endX) / 2, (startY + endY) / 2, Math.abs(endX - startX) / 2, Math.abs(endY - startY) / 2, 0, 0, Math.PI * 2);
  }
  if (filled && tool !== "line") {
    context.fillStyle = context.strokeStyle;
    context.fill();
  } else {
    context.stroke();
  }
}

// Shift-drag: lines snap to 45-degree steps, rect/ellipse become square/circle.
function constrainShapePoint(tool, startX, startY, endX, endY) {
  const deltaX = endX - startX;
  const deltaY = endY - startY;
  if (tool === "line") {
    const angle = Math.round(Math.atan2(deltaY, deltaX) / (Math.PI / 4)) * (Math.PI / 4);
    const length = Math.hypot(deltaX, deltaY);
    return { x: startX + Math.cos(angle) * length, y: startY + Math.sin(angle) * length };
  }
  const side = Math.min(Math.abs(deltaX), Math.abs(deltaY));
  return {
    x: startX + (deltaX < 0 ? -side : side),
    y: startY + (deltaY < 0 ? -side : side)
  };
}

// MS Paint-style text: an editable overlay box at the click point; typing is
// WYSIWYG (same font, color, and zoom scale as the canvas) and the text is
// drawn onto the canvas when the box loses focus. Escape cancels.
function openPaintTextEditor(x, y) {
  const shell = imageCanvas?.closest(".image-document");
  if (!shell) {
    return;
  }

  // ponytail: text size rides the brush-size select (x4).
  const size = Math.max(12, (Number(imageSizeSelect.value) || 4) * 4);
  const font = paintFontSelect.value || "Arial";
  const scale = imageCanvas.clientWidth / imageCanvas.width;
  const element = document.createElement("div");
  element.className = "paint-text-editor";
  element.contentEditable = "plaintext-only";
  element.spellcheck = false;
  element.style.left = `${imageCanvas.offsetLeft + x * scale}px`;
  element.style.top = `${imageCanvas.offsetTop + y * scale}px`;
  element.style.color = imageColorInput.value;
  element.style.font = `${size * scale}px ${font}`;

  paintTextEditor = { element, x, y, size, font, color: imageColorInput.value };

  element.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      event.preventDefault();
      paintTextEditor = null;
      element.remove();
    }
    // Keep typing (and Ctrl+Z etc.) inside the box, away from the global shortcuts.
    event.stopPropagation();
  });
  element.addEventListener("blur", () => commitPaintText());

  shell.appendChild(element);
  element.focus();
}

function commitPaintText() {
  if (!paintTextEditor) {
    return;
  }

  const { element, x, y, size, font, color } = paintTextEditor;
  paintTextEditor = null;
  const text = element.innerText.replace(/\n+$/, "");
  const box = element.getBoundingClientRect(); // read before removal
  element.remove();
  if (!text.trim() || !imageCanvas || !viewer.contains(imageCanvas)) {
    return;
  }

  pushPaintUndo();
  const context = imageCanvas.getContext("2d");
  context.globalCompositeOperation = "source-over";
  context.fillStyle = color;
  context.font = `${size}px ${font || "Arial"}`;
  context.textBaseline = "top";
  // The box is resizable, so mirror its wrapping and clipping in canvas space.
  const scale = imageCanvas.clientWidth / imageCanvas.width || 1;
  const maxWidth = box.width / scale + 0.5;
  const maxHeight = box.height / scale + 0.5;
  // 1.2 line-height with half-leading above, matching the overlay's CSS line boxes.
  wrapPaintTextLines(context, text, maxWidth).forEach((line, index) => {
    const top = size * 0.1 + index * size * 1.2;
    if (top < maxHeight) {
      context.fillText(line, x, y + top);
    }
  });
  setDirty(true);
}

// Greedy word wrap matching the overlay's pre-wrap layout. ponytail: a single
// word wider than the box overflows instead of char-breaking, same as the CSS.
function wrapPaintTextLines(context, text, maxWidth) {
  const lines = [];
  for (const paragraph of text.split("\n")) {
    let line = "";
    for (const word of paragraph.split(" ")) {
      const candidate = line ? `${line} ${word}` : word;
      if (!line || context.measureText(candidate).width <= maxWidth) {
        line = candidate;
      } else {
        lines.push(line);
        line = word;
      }
    }
    lines.push(line);
  }
  return lines;
}

function transformPaintCanvas(action) {
  if (!imageCanvas || !viewer.contains(imageCanvas)) {
    return;
  }

  pushPaintUndo();
  const source = document.createElement("canvas");
  source.width = imageCanvas.width;
  source.height = imageCanvas.height;
  source.getContext("2d").drawImage(imageCanvas, 0, 0);

  const context = imageCanvas.getContext("2d");
  if (action === "rotate") {
    imageCanvas.width = source.height;
    imageCanvas.height = source.width;
  } else {
    imageCanvas.width = source.width; // reassigning clears the canvas
  }
  context.save();
  if (action === "rotate") {
    context.translate(imageCanvas.width, 0);
    context.rotate(Math.PI / 2);
  } else if (action === "flip-h") {
    context.translate(imageCanvas.width, 0);
    context.scale(-1, 1);
  } else {
    context.translate(0, imageCanvas.height);
    context.scale(1, -1);
  }
  context.drawImage(source, 0, 0);
  context.restore();
  syncPaintDocumentSize();
  setDirty(true);
}

async function resizePaintCanvas() {
  if (!imageCanvas || !viewer.contains(imageCanvas)) {
    return;
  }

  const spec = await askInline("Canvas size", `${imageCanvas.width}x${imageCanvas.height}`, {
    message: "Enter width x height in pixels. Content is kept at the top-left.",
    primaryLabel: "Resize"
  });
  if (!spec || !imageCanvas || !viewer.contains(imageCanvas)) {
    return;
  }

  const match = /^\s*(\d+)\s*[x×,\s]\s*(\d+)\s*$/i.exec(spec);
  if (!match) {
    setStatus("Enter a size like 1280x720");
    return;
  }

  applyPaintCanvasSize(clampInt(match[1], 8, 4096), clampInt(match[2], 8, 4096));
}

// Content stays anchored at the top-left; new area is transparent on PNG,
// white on JPEG (no alpha there).
function applyPaintCanvasSize(width, height) {
  if (!imageCanvas || !viewer.contains(imageCanvas) || (width === imageCanvas.width && height === imageCanvas.height)) {
    return;
  }

  pushPaintUndo();
  const source = document.createElement("canvas");
  source.width = imageCanvas.width;
  source.height = imageCanvas.height;
  source.getContext("2d").drawImage(imageCanvas, 0, 0);

  imageCanvas.width = width;
  imageCanvas.height = height;
  const context = imageCanvas.getContext("2d");
  if (currentDocument?.extension === ".jpg" || currentDocument?.extension === ".jpeg") {
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, width, height);
  }
  context.drawImage(source, 0, 0);
  syncPaintDocumentSize();
  setDirty(true);
}

function syncPaintDocumentSize() {
  if (currentDocument?.kind === "image" && imageCanvas) {
    currentDocument.width = imageCanvas.width;
    currentDocument.height = imageCanvas.height;
  }
}

function floodFill(context, startX, startY, hexColor) {
  const { width, height } = context.canvas;
  if (startX < 0 || startY < 0 || startX >= width || startY >= height) {
    return;
  }

  const image = context.getImageData(0, 0, width, height);
  const data = image.data;
  const fillRed = parseInt(hexColor.slice(1, 3), 16);
  const fillGreen = parseInt(hexColor.slice(3, 5), 16);
  const fillBlue = parseInt(hexColor.slice(5, 7), 16);
  const startOffset = (startY * width + startX) * 4;
  const targetRed = data[startOffset];
  const targetGreen = data[startOffset + 1];
  const targetBlue = data[startOffset + 2];
  const targetAlpha = data[startOffset + 3];
  if (targetRed === fillRed && targetGreen === fillGreen && targetBlue === fillBlue && targetAlpha === 255) {
    return;
  }

  const TOLERANCE = Number(fillToleranceSelect.value) || 0;
  const visited = new Uint8Array(width * height);
  const stack = [startY * width + startX];
  while (stack.length > 0) {
    const index = stack.pop();
    if (visited[index]) {
      continue;
    }
    visited[index] = 1;

    const offset = index * 4;
    if (
      Math.abs(data[offset] - targetRed) > TOLERANCE ||
      Math.abs(data[offset + 1] - targetGreen) > TOLERANCE ||
      Math.abs(data[offset + 2] - targetBlue) > TOLERANCE ||
      Math.abs(data[offset + 3] - targetAlpha) > TOLERANCE
    ) {
      continue;
    }

    data[offset] = fillRed;
    data[offset + 1] = fillGreen;
    data[offset + 2] = fillBlue;
    data[offset + 3] = 255;

    const x = index % width;
    if (x > 0) {
      stack.push(index - 1);
    }
    if (x < width - 1) {
      stack.push(index + 1);
    }
    if (index >= width) {
      stack.push(index - width);
    }
    if (index < width * (height - 1)) {
      stack.push(index + width);
    }
  }

  context.putImageData(image, 0, 0);
}

// ponytail: undo/redo history is 10 full-canvas snapshots each (~37 MB at
// 1280x720) and resets when the canvas re-renders (edit toggle, tab switch);
// move to compressed/tiled snapshots if memory or history depth ever bites.
function pushPaintUndo(snapshot) {
  if (!imageCanvas) {
    return;
  }

  const image = snapshot || imageCanvas.getContext("2d").getImageData(0, 0, imageCanvas.width, imageCanvas.height);
  paintUndoStack.push(image);
  if (paintUndoStack.length > 10) {
    paintUndoStack.shift();
  }
  paintRedoStack = []; // a fresh action invalidates the redo trail
  clearPaintSelection(); // every mutation routes through here, so the marquee never goes stale
}

function undoPaint() {
  movePaintHistory(paintUndoStack, paintRedoStack);
}

function redoPaint() {
  movePaintHistory(paintRedoStack, paintUndoStack);
}

function movePaintHistory(fromStack, toStack) {
  if (fromStack.length === 0 || !imageCanvas || !viewer.contains(imageCanvas)) {
    return;
  }

  const context = imageCanvas.getContext("2d");
  toStack.push(context.getImageData(0, 0, imageCanvas.width, imageCanvas.height));
  if (toStack.length > 10) {
    toStack.shift();
  }

  const image = fromStack.pop();
  // Rotate/resize snapshots can differ in size — restore the dimensions too.
  if (image.width !== imageCanvas.width || image.height !== imageCanvas.height) {
    imageCanvas.width = image.width;
    imageCanvas.height = image.height;
    syncPaintDocumentSize();
  }
  context.putImageData(image, 0, 0);
  clearPaintSelection();
  setDirty(true);
}

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

function renderPdf(pdfData) {
  wordEditor = null;
  savedWordRange = null;
  toggleFindBar(false);
  selectedCell = null;
  selectedRange = null;
  rangeAnchor = null;
  updateSelectionLabel();
  sheetPanel.hidden = true;
  sheetList.replaceChildren();

  const frameShell = document.createElement("div");
  frameShell.className = "pdf-document";

  const frame = document.createElement("iframe");
  frame.className = "pdf-frame";
  frame.title = pdfData.fileName;
  frame.src = pdfData.fileUrl;

  frameShell.appendChild(frame);
  viewer.replaceChildren(frameShell);
}

function renderError(message) {
  fileName.textContent = "Nothing open";
  fileName.title = "";
  fileMeta.textContent = "No document selected.";
  documentTitle.textContent = "No document selected";
  applyDocumentIdentity(null);
  wordEditor = null;
  toggleFindBar(false);
  lastSavedSnapshot = null;
  selectedCell = null;
  selectedRange = null;
  rangeAnchor = null;
  sheetPanel.hidden = true;
  sheetList.replaceChildren();
  enableDocumentActions(false);
  viewer.replaceChildren(emptyBlock(message));
}

function emptyBlock(message) {
  const block = document.createElement("div");
  block.className = "empty-state";
  const heading = document.createElement("h3");
  heading.textContent = message;
  block.appendChild(heading);
  return block;
}

const SAFE_WORD_TAGS = new Set([
  "a",
  "b",
  "blockquote",
  "br",
  "div",
  "em",
  "font",
  "h1",
  "h2",
  "h3",
  "i",
  "img",
  "li",
  "ol",
  "p",
  "s",
  "span",
  "strike",
  "strong",
  "table",
  "tbody",
  "td",
  "th",
  "thead",
  "tr",
  "u",
  "ul"
]);
const REMOVE_WORD_TAGS = new Set(["script", "style", "iframe", "object", "embed", "svg", "math", "meta", "link", "base"]);
const SAFE_WORD_ATTRIBUTES = new Set(["align", "alt", "class", "color", "face", "href", "size", "src", "style", "title"]);
const SAFE_STYLE_PROPERTIES = new Set([
  "background-color",
  "color",
  "font-family",
  "font-size",
  "font-style",
  "font-weight",
  "height",
  "max-width",
  "text-align",
  "text-decoration",
  "width"
]);

function sanitizeWordHtml(html) {
  const template = document.createElement("template");
  template.innerHTML = String(html || "");
  sanitizeWordChildren(template.content);
  return template.innerHTML;
}

function sanitizeWordChildren(parent) {
  Array.from(parent.childNodes).forEach((node) => {
    if (node.nodeType === Node.COMMENT_NODE) {
      node.remove();
      return;
    }

    if (node.nodeType !== Node.ELEMENT_NODE) {
      return;
    }

    const tagName = node.tagName.toLowerCase();
    if (REMOVE_WORD_TAGS.has(tagName)) {
      node.remove();
      return;
    }

    if (!SAFE_WORD_TAGS.has(tagName)) {
      node.replaceWith(...Array.from(node.childNodes));
      sanitizeWordChildren(parent);
      return;
    }

    sanitizeWordAttributes(node);
    sanitizeWordChildren(node);
  });
}

function sanitizeWordAttributes(element) {
  Array.from(element.attributes).forEach((attribute) => {
    const name = attribute.name.toLowerCase();
    const value = attribute.value;

    if (name.startsWith("on") || !SAFE_WORD_ATTRIBUTES.has(name)) {
      element.removeAttribute(attribute.name);
      return;
    }

    if (name === "style") {
      const safeStyle = sanitizeStyle(value);
      if (safeStyle) {
        element.setAttribute("style", safeStyle);
      } else {
        element.removeAttribute(attribute.name);
      }
      return;
    }

    if ((name === "href" || name === "src") && !isSafeWordUrl(value, name, element.tagName.toLowerCase())) {
      element.removeAttribute(attribute.name);
    }
  });
}

function sanitizeStyle(style) {
  return String(style || "")
    .split(";")
    .map((declaration) => {
      const [rawName, ...rawValue] = declaration.split(":");
      const name = rawName?.trim().toLowerCase();
      const value = rawValue.join(":").trim();
      if (!name || !value || !SAFE_STYLE_PROPERTIES.has(name) || /url\s*\(|expression|javascript:|-moz-binding/i.test(value)) {
        return "";
      }
      return `${name}: ${value}`;
    })
    .filter(Boolean)
    .join("; ");
}

function isSafeWordUrl(value, attributeName, tagName) {
  const url = String(value || "").trim();
  if (!url) {
    return false;
  }

  if (attributeName === "src" && tagName === "img") {
    return /^data:image\/(?:png|jpe?g|gif);base64,/i.test(url);
  }

  if (attributeName !== "href") {
    return false;
  }

  try {
    const parsed = new URL(url);
    return ["http:", "https:", "mailto:"].includes(parsed.protocol);
  } catch {
    return false;
  }
}

function renderDocumentNotices(documentData) {
  if (!documentNotices) {
    return;
  }

  documentNotices.replaceChildren();
  if (!documentData) {
    documentNotices.hidden = true;
    return;
  }

  const notices = [];
  if (Array.isArray(documentData.safetyWarnings)) {
    notices.push(...documentData.safetyWarnings.map((message) => ({ kind: "warning", message })));
  }
  if (Array.isArray(documentData.warnings) && documentData.warnings.length > 0) {
    notices.push({
      kind: "warning",
      message: `${documentData.warnings.length} conversion warning${documentData.warnings.length === 1 ? "" : "s"}`,
      title: documentData.warnings.join("\n")
    });
  }
  if (documentData.lastBackupPath) {
    notices.push({
      kind: "info",
      message: `Backup created: ${baseName(documentData.lastBackupPath)}`,
      title: documentData.lastBackupPath
    });
  }
  if (documentData.lastSaveMode === "temp-replace") {
    notices.push({
      kind: "info",
      message: "Safe save completed"
    });
  }

  documentNotices.hidden = notices.length === 0;
  notices.forEach((notice) => {
    const item = document.createElement("div");
    item.className = `document-notice is-${notice.kind}`;
    item.textContent = notice.message;
    if (notice.title) {
      item.title = notice.title;
    }
    documentNotices.appendChild(item);
  });
}

function renderDocumentInfo(documentData) {
  if (!documentInfoPanel || !documentInfoGrid) {
    return;
  }

  documentInfoGrid.replaceChildren();
  documentInfoPanel.hidden = !documentData;
  if (!documentData) {
    return;
  }

  const identity = documentIdentity(documentData);
  const rows = [
    ["Type", identity.label],
    ["Size", documentData.size ? formatFileSize(documentData.size) : documentData.isNew ? "Not saved" : "-"],
    ["Modified", documentData.modifiedAt ? formatDateTime(documentData.modifiedAt) : documentData.isNew ? "Not saved" : "-"],
    ["Backups", backupCountForDocument(documentData)]
  ];

  rows.forEach(([label, value]) => {
    const term = document.createElement("dt");
    term.textContent = label;
    const detail = document.createElement("dd");
    detail.textContent = String(value);
    documentInfoGrid.append(term, detail);
  });
}

function renderBackupHistory(documentData = currentDocument) {
  if (!backupPanel || !backupList || !clearBackupsButton) {
    return;
  }

  backupList.replaceChildren();
  const backups = backupsForDocument(documentData);
  backupPanel.hidden = !documentData || backups.length === 0;
  clearBackupsButton.hidden = backups.length === 0;
  if (!documentData || backups.length === 0) {
    return;
  }

  backups.forEach((backup) => {
    const item = document.createElement("button");
    item.type = "button";
    item.className = "backup-item";
    item.title = backup.backupPath;
    item.addEventListener("click", async () => {
      const result = await window.documentOpener.revealPath(backup.backupPath);
      if (!result?.ok) {
        setStatus("Backup missing");
        await showAlert({
          title: "Backup missing",
          message: result?.error || "The backup file could not be found.",
          tone: "warning",
          primaryLabel: "OK"
        });
        backupHistory = backupHistory.filter((itemData) => itemData.backupPath !== backup.backupPath);
        saveBackupHistory();
        renderBackupHistory();
      }
    });

    const icon = document.createElement("span");
    icon.className = "backup-icon";
    icon.textContent = "BAK";

    const copy = document.createElement("span");
    copy.className = "backup-copy";
    const name = document.createElement("strong");
    name.textContent = baseName(backup.backupPath);
    const meta = document.createElement("span");
    meta.textContent = backup.createdAt ? formatDateTime(backup.createdAt) : "Backup";
    copy.append(name, meta);
    item.append(icon, copy);
    backupList.appendChild(item);
  });
}

function recordBackup(documentData, backupPath) {
  if (!documentData?.filePath || !backupPath) {
    return;
  }

  const entry = {
    backupPath,
    sourcePath: documentData.filePath,
    sourceName: documentData.fileName,
    kind: documentData.kind,
    extension: documentData.extension || extensionFromPath(documentData.filePath),
    createdAt: new Date().toISOString()
  };

  backupHistory = [
    entry,
    ...backupHistory.filter((item) => item.backupPath !== backupPath)
  ].slice(0, MAX_BACKUP_HISTORY);
  saveBackupHistory();
}

function backupsForDocument(documentData) {
  if (!documentData?.filePath) {
    return [];
  }

  return backupHistory.filter((item) => sameFilePath(item.sourcePath, documentData.filePath));
}

function backupCountForDocument(documentData) {
  const count = backupsForDocument(documentData).length;
  return `${count} backup${count === 1 ? "" : "s"}`;
}

function markdownToHtml(markdown) {
  const lines = String(markdown || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  const html = [];
  let paragraph = [];
  let list = null;

  const flushParagraph = () => {
    if (paragraph.length === 0) {
      return;
    }
    html.push(`<p>${renderMarkdownInline(paragraph.join(" "))}</p>`);
    paragraph = [];
  };

  const flushList = () => {
    if (!list) {
      return;
    }
    html.push(`<${list.type}>${list.items.map(renderMarkdownListItem).join("")}</${list.type}>`);
    list = null;
  };

  const flushBlocks = () => {
    flushParagraph();
    flushList();
  };

  const pushListItem = (type, text) => {
    flushParagraph();
    if (list && list.type !== type) {
      flushList();
    }
    if (!list) {
      list = { type, items: [] };
    }
    list.items.push(text);
  };

  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];

    if (/^\s*```/.test(line)) {
      flushBlocks();
      const codeLines = [];
      i += 1;
      while (i < lines.length && !/^\s*```/.test(lines[i])) {
        codeLines.push(lines[i]);
        i += 1;
      }
      html.push(`<pre><code>${escapeHtml(codeLines.join("\n"))}</code></pre>`);
      continue;
    }

    if (/^\s*$/.test(line)) {
      flushBlocks();
      continue;
    }

    if (/^\s*([-*_])(?:\s*\1){2,}\s*$/.test(line)) {
      flushBlocks();
      html.push("<hr />");
      continue;
    }

    const heading = /^(#{1,6})\s+(.+?)\s*#*\s*$/.exec(line);
    if (heading) {
      flushBlocks();
      html.push(`<h${heading[1].length}>${renderMarkdownInline(heading[2])}</h${heading[1].length}>`);
      continue;
    }

    if (line.includes("|") && i + 1 < lines.length && /^\s*\|?[\s:|-]*-[\s:|-]*\|?\s*$/.test(lines[i + 1])) {
      flushBlocks();
      const header = splitMarkdownTableRow(line);
      const aligns = splitMarkdownTableRow(lines[i + 1]).map(markdownCellAlign);
      const rows = [];
      i += 2;
      while (i < lines.length && lines[i].includes("|") && !/^\s*$/.test(lines[i])) {
        rows.push(splitMarkdownTableRow(lines[i]));
        i += 1;
      }
      i -= 1;
      html.push(renderMarkdownTable(header, aligns, rows));
      continue;
    }

    const quote = /^\s*>\s?(.*)$/.exec(line);
    if (quote) {
      flushBlocks();
      html.push(`<blockquote>${renderMarkdownInline(quote[1])}</blockquote>`);
      continue;
    }

    const orderedItem = /^\s*\d+[.)]\s+(.+)$/.exec(line);
    if (orderedItem) {
      pushListItem("ol", orderedItem[1]);
      continue;
    }

    const unorderedItem = /^\s*[-*+]\s+(.+)$/.exec(line);
    if (unorderedItem) {
      pushListItem("ul", unorderedItem[1]);
      continue;
    }

    flushList();
    paragraph.push(line.trim());
  }

  flushBlocks();

  return html.length > 0 ? html.join("") : "<p>This Markdown file is empty.</p>";
}

function renderMarkdownListItem(item) {
  const task = /^\[([ xX])\]\s+(.*)$/.exec(item);
  if (task) {
    const checked = task[1].toLowerCase() === "x" ? " checked" : "";
    return `<li class="task-item"><input type="checkbox" disabled${checked} /> ${renderMarkdownInline(task[2])}</li>`;
  }
  return `<li>${renderMarkdownInline(item)}</li>`;
}

function splitMarkdownTableRow(line) {
  let cells = line.trim();
  if (cells.startsWith("|")) {
    cells = cells.slice(1);
  }
  if (cells.endsWith("|")) {
    cells = cells.slice(0, -1);
  }
  return cells.split("|").map((cell) => cell.trim());
}

function markdownCellAlign(separator) {
  const left = separator.startsWith(":");
  const right = separator.endsWith(":");
  if (left && right) {
    return "center";
  }
  return right ? "right" : left ? "left" : "";
}

function renderMarkdownTable(header, aligns, rows) {
  const alignAttr = (index) => (aligns[index] ? ` style="text-align:${aligns[index]}"` : "");
  const head = header.map((cell, index) => `<th${alignAttr(index)}>${renderMarkdownInline(cell)}</th>`).join("");
  const body = rows
    .map((row) => `<tr>${header.map((_cell, index) => `<td${alignAttr(index)}>${renderMarkdownInline(row[index] || "")}</td>`).join("")}</tr>`)
    .join("");
  return `<table class="md-table"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
}

function renderMarkdownInline(value) {
  return String(value)
    .split(/(`[^`]+`)/)
    .map((part) => {
      if (part.length >= 2 && part.startsWith("`") && part.endsWith("`")) {
        return `<code>${escapeHtml(part.slice(1, -1))}</code>`;
      }
      return formatMarkdownText(part);
    })
    .join("");
}

function formatMarkdownText(text) {
  let html = escapeHtml(text);
  html = html.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, (_match, alt, url) => {
    const safeUrl = safePreviewImage(url);
    return safeUrl ? `<img src="${escapeHtml(safeUrl)}" alt="${alt}" />` : escapeHtml(alt);
  });
  html = html.replace(/~~([^~]+)~~/g, "<s>$1</s>");
  html = html.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");
  html = html.replace(/__([^_]+)__/g, "<strong>$1</strong>");
  html = html.replace(/\*([^*]+)\*/g, "<em>$1</em>");
  html = html.replace(/\b_([^_]+)_\b/g, "<em>$1</em>");
  html = html.replace(/\[([^\]]+)\]\(([^)]+)\)/g, (_match, label, url) => {
    const safeUrl = safePreviewLink(url);
    return safeUrl ? `<a href="${escapeHtml(safeUrl)}">${label}</a>` : label;
  });
  return html;
}

function safePreviewImage(value) {
  const link = String(value || "").trim();
  if (/^data:image\//i.test(link)) {
    return link;
  }
  try {
    const parsed = new URL(link);
    return ["http:", "https:"].includes(parsed.protocol) ? link : null;
  } catch {
    return null;
  }
}

function safePreviewLink(value) {
  const link = String(value || "").trim();
  try {
    const parsed = new URL(link);
    return ["http:", "https:", "mailto:"].includes(parsed.protocol) ? link : null;
  } catch {
    return null;
  }
}

function htmlToPlainText(html) {
  const template = document.createElement("template");
  template.innerHTML = String(html || "");
  return template.content.textContent || "";
}

// ponytail: covers the tags the Word editor produces (headings, emphasis,
// lists, links, tables, blockquotes); anything else falls back to its text.
function htmlToMarkdown(html) {
  const template = document.createElement("template");
  template.innerHTML = String(html || "");
  return `${markdownBlocks(template.content).replace(/\n{3,}/g, "\n\n").trim()}\n`;
}

const MARKDOWN_BLOCK_TAGS = new Set(["p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "table", "blockquote", "pre"]);

function markdownBlocks(parent) {
  let out = "";
  parent.childNodes.forEach((node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.nodeValue.trim();
      if (text) {
        out += `${text}\n\n`;
      }
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      return;
    }

    const tag = node.tagName.toLowerCase();
    const heading = /^h([1-6])$/.exec(tag);
    if (heading) {
      out += `${"#".repeat(Number(heading[1]))} ${markdownInline(node).trim()}\n\n`;
    } else if (tag === "blockquote") {
      out += `> ${markdownInline(node).trim()}\n\n`;
    } else if (tag === "ul" || tag === "ol") {
      out += `${markdownList(node, tag === "ol")}\n`;
    } else if (tag === "table") {
      out += `${markdownTable(node)}\n`;
    } else if (tag === "pre") {
      out += `\`\`\`\n${node.textContent.replace(/\n$/, "")}\n\`\`\`\n\n`;
    } else if (tag === "br") {
      out += "\n";
    } else if (Array.from(node.children).some((child) => MARKDOWN_BLOCK_TAGS.has(child.tagName.toLowerCase()))) {
      out += markdownBlocks(node); // container element — unwrap its blocks
    } else {
      const inline = markdownInline(node).trim();
      if (inline) {
        out += `${inline}\n\n`;
      }
    }
  });
  return out;
}

function markdownInline(parent) {
  let out = "";
  parent.childNodes.forEach((node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      out += node.nodeValue.replace(/\s+/g, " ");
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      return;
    }

    const tag = node.tagName.toLowerCase();
    const inner = markdownInline(node);
    const trimmed = inner.trim();
    if (tag === "strong" || tag === "b") {
      out += trimmed ? ` **${trimmed}** ` : "";
    } else if (tag === "em" || tag === "i") {
      out += trimmed ? ` *${trimmed}* ` : "";
    } else if (tag === "s" || tag === "strike") {
      out += trimmed ? ` ~~${trimmed}~~ ` : "";
    } else if (tag === "code") {
      out += `\`${node.textContent}\``;
    } else if (tag === "a") {
      const href = node.getAttribute("href") || "";
      out += `[${trimmed || href}](${href})`;
    } else if (tag === "img") {
      out += node.getAttribute("alt") ? `![${node.getAttribute("alt")}]()` : "";
    } else if (tag === "br") {
      out += "\n";
    } else {
      out += inner;
    }
  });
  return out.replace(/ {2,}/g, " ");
}

function markdownList(list, ordered, indent = "") {
  let out = "";
  let index = 0;
  Array.from(list.children).forEach((item) => {
    if (item.tagName.toLowerCase() !== "li") {
      return;
    }
    index += 1;
    const itemText = item.cloneNode(true);
    itemText.querySelectorAll("ul, ol").forEach((nested) => nested.remove());
    out += `${indent}${ordered ? `${index}.` : "-"} ${markdownInline(itemText).trim()}\n`;
    item.querySelectorAll(":scope > ul, :scope > ol").forEach((nested) => {
      out += markdownList(nested, nested.tagName.toLowerCase() === "ol", `${indent}  `);
    });
  });
  return out;
}

function markdownTable(table) {
  const rows = Array.from(table.querySelectorAll("tr"))
    .map((tr) => Array.from(tr.querySelectorAll("td, th")).map((cell) => markdownInline(cell).trim().replace(/\|/g, "\\|")))
    .filter((row) => row.length > 0);
  if (rows.length === 0) {
    return "";
  }

  const width = Math.max(...rows.map((row) => row.length));
  const line = (cells) => `| ${Array.from({ length: width }, (_unused, i) => cells[i] || "").join(" | ")} |`;
  return [line(rows[0]), `| ${Array(width).fill("---").join(" | ")} |`, ...rows.slice(1).map(line), ""].join("\n");
}

function kindForExtension(extension) {
  if (extension === ".docx") {
    return "word";
  }
  if (extension === ".md") {
    return "markdown";
  }
  if ([".txt", ".log", ".json"].includes(extension)) {
    return "text";
  }
  if ([".xlsx", ".csv"].includes(extension)) {
    return "workbook";
  }
  if ([".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".svg"].includes(extension)) {
    return "image";
  }
  return null;
}

function warningSummary(warnings) {
  if (!warnings || warnings.length === 0) {
    return "Word document";
  }

  return `Word document - ${warnings.length} preview warning${warnings.length === 1 ? "" : "s"}`;
}

function setStatus(text) {
  statusText.textContent = text;
  const toastMessage = TOAST_MESSAGES.get(text);

  if (toastMessage) {
    showToast(toastMessage);
  }
}

function showToast(message) {
  if (!toast) {
    return;
  }

  window.clearTimeout(toastTimer);
  toast.textContent = message;
  toast.classList.add("show");
  toastTimer = window.setTimeout(() => {
    toast.classList.remove("show");
  }, 2200);
}

function showAlert(options) {
  return showAppDialog({
    title: options.title,
    message: options.message,
    tone: options.tone || "info",
    primaryLabel: options.primaryLabel || "OK"
  }).then(() => true);
}

function showConfirm(options) {
  return showAppDialog({
    title: options.title,
    message: options.message,
    tone: options.tone || "info",
    primaryLabel: options.primaryLabel || "Continue",
    cancelLabel: options.cancelLabel || "Cancel"
  }).then((result) => result.action === "primary");
}

function showPrompt(options) {
  return showAppDialog({
    title: options.title,
    message: options.message || "",
    tone: options.tone || "info",
    primaryLabel: options.primaryLabel || "OK",
    cancelLabel: options.cancelLabel || "Cancel",
    input: true,
    defaultValue: options.defaultValue || "",
    inputLabel: options.inputLabel || options.title
  }).then((result) => (result.action === "primary" ? result.value : null));
}

function showAppDialog(options) {
  if (!dialogOverlay) {
    return Promise.resolve({ action: "primary", value: "" });
  }

  if (activeDialog) {
    resolveAppDialog("cancel");
  }

  const previousFocus = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  dialogTitle.textContent = options.title || "DocOpen";
  dialogMessage.textContent = options.message || "";
  dialogMessage.hidden = !options.message;
  dialogIcon.textContent = dialogIconForTone(options.tone);
  appDialog.dataset.tone = options.tone || "info";

  dialogPrimaryButton.textContent = options.primaryLabel || "OK";
  dialogSecondaryButton.hidden = !options.secondaryLabel;
  dialogSecondaryButton.textContent = options.secondaryLabel || "";
  dialogCancelButton.hidden = !options.cancelLabel;
  dialogCancelButton.textContent = options.cancelLabel || "";

  dialogInput.hidden = !options.input;
  dialogInput.value = options.defaultValue || "";
  dialogInput.setAttribute("aria-label", options.inputLabel || "Value");

  dialogOverlay.hidden = false;
  document.body.classList.add("has-dialog");

  return new Promise((resolve) => {
    activeDialog = { resolve, previousFocus };
    window.requestAnimationFrame(() => {
      if (options.input) {
        dialogInput.focus();
        dialogInput.select();
      } else {
        dialogPrimaryButton.focus();
      }
    });
  });
}

function resolveAppDialog(action) {
  if (!activeDialog) {
    return;
  }

  const { resolve, previousFocus } = activeDialog;
  const value = dialogInput.value;
  activeDialog = null;
  dialogOverlay.hidden = true;
  document.body.classList.remove("has-dialog");
  resolve({ action, value });
  previousFocus?.focus?.();
}

function dialogIconForTone(tone) {
  if (tone === "warning") {
    return "!";
  }
  if (tone === "danger") {
    return "x";
  }
  return "i";
}

function setEditMode(nextEditMode, options = {}) {
  if (options.sync !== false) {
    syncCurrentDocumentFromDom();
  }
  editMode = Boolean(nextEditMode && currentDocument && isEditableDocument(currentDocument));
  document.body.classList.toggle("is-editing", editMode);
  updateModeLabel();
  editButtonLabel.textContent = editMode ? "Done" : "Edit";
  editButton.classList.toggle("is-active", editMode);
  editButton.setAttribute("aria-pressed", String(editMode));
  editToolbar.hidden = !editMode;
  updateDocumentActionVisibility();
  updateLayoutControls();

  if (!currentDocument) {
    editButton.disabled = true;
    editToolbar.hidden = true;
    return;
  }

  editButton.disabled = !isEditableDocument(currentDocument);
  updateToolbarForDocument();

  if (currentDocument.kind === "word") {
    renderWord(currentDocument);
  } else if (currentDocument.kind === "markdown" || currentDocument.kind === "text") {
    renderMarkdown(currentDocument);
  } else if (currentDocument.kind === "workbook") {
    // Note: no syncWorkbookFromDom() here — the top of this function already
    // syncs when sync is enabled. Syncing again on open (sync:false) would read
    // the previous document's grid into the freshly-loaded one.
    renderWorkbook(currentDocument);
  } else if (currentDocument.kind === "pdf") {
    renderPdf(currentDocument);
  } else if (currentDocument.kind === "image") {
    renderImageDocument(currentDocument);
  }
}

function setDirty(nextDirty) {
  isDirty = Boolean(nextDirty);
  if (tabs[activeTabIndex]) {
    tabs[activeTabIndex].isDirty = isDirty;
  }
  dirtyStatus.hidden = !isDirty;
  saveButton.disabled = !canSaveCurrentDocument() || !isDirty;
  revertButton.disabled = !isEditableDocument(currentDocument) || !isDirty;
  updateSaveButtonLabel();
  updateDocumentActionVisibility();
  updateWindowTitle();
  // The window-close guard cares about ANY dirty tab, not just the active one.
  window.documentOpener.notifyDirty?.(tabs.some((tab) => tab.isDirty));
  renderTabStrip();

  if (!currentDocument) {
    return;
  }

  fileMeta.textContent = `${documentSummary(currentDocument)}${isDirty ? " - unsaved changes" : ""}`;
  renderDocumentNotices(currentDocument);
}

function updateWindowTitle() {
  const name = currentDocument?.fileName;
  document.title = name ? `${isDirty ? "* " : ""}${name} - DocOpen` : "DocOpen";
}

function enableDocumentActions(isEnabled) {
  const canEdit = isEnabled && isEditableDocument(currentDocument);
  updateDocumentActionVisibility();
  editButton.disabled = !canEdit;
  saveButton.disabled = !canEdit || !isDirty;
  saveAsButton.disabled = !canEdit;
  revertButton.disabled = !canEdit || !isDirty;
  updateSaveButtonLabel();
}

function updateDocumentActionVisibility() {
  const shouldShow = isEditableDocument(currentDocument);
  [editButton, saveButton, saveAsButton, revertButton].forEach((button) => {
    button.hidden = !shouldShow;
  });
  document.body.classList.toggle("has-editable-document", shouldShow);
  updateLayoutControls();
}

function updateModeLabel() {
  if (!currentDocument) {
    modeLabel.textContent = "Read-only preview";
  } else if (!isEditableDocument(currentDocument)) {
    modeLabel.textContent = "Read-only document";
  } else {
    modeLabel.textContent = editMode ? "Editing enabled" : "Preview";
  }
}

async function saveCurrentDocument(saveAs) {
  if (!canSaveCurrentDocument()) {
    return;
  }

  syncCurrentDocumentFromDom();

  const payload = buildSavePayload();
  if (!payload) {
    setStatus("Read-only");
    return;
  }

  const shouldSaveAs = saveAs || !payload.filePath;
  setStatus(shouldSaveAs ? "Saving as..." : "Saving...");

  const result = shouldSaveAs
    ? await window.documentOpener.saveDocumentAs(payload)
    : await window.documentOpener.saveDocument(payload);

  if (!result || result.canceled) {
    setStatus("Save canceled");
    return;
  }

  if (!result.ok) {
    setStatus("Save failed");
    await showAlert({
      title: "Save failed",
      message: result.error || "The document could not be saved.",
      tone: "danger",
      primaryLabel: "OK"
    });
    return;
  }

  const previousExtension = currentDocument.extension;
  currentDocument.filePath = result.filePath;
  currentDocument.fileName = result.fileName;
  currentDocument.extension = extensionFromPath(result.filePath) || currentDocument.extension;

  // Saved into another document family (e.g. .md exported as .docx)? Reload
  // from disk so the matching viewer/editor takes over.
  const savedKind = kindForExtension(currentDocument.extension);
  if (savedKind && savedKind !== currentDocument.kind) {
    setStatus(result.backupPath ? "Saved with backup" : "Saved");
    const reopened = await window.documentOpener.openRecentFile(result.filePath);
    handleOpenResult(reopened, { reuseTab: true });
    return;
  }
  currentDocument.size = result.size ?? currentDocument.size;
  currentDocument.modifiedAt = result.modifiedAt || currentDocument.modifiedAt;
  currentDocument.lastBackupPath = result.backupPath || null;
  currentDocument.lastSaveMode = result.saveMode || null;
  currentDocument.isNew = false;
  recordBackup(currentDocument, result.backupPath);
  lastSavedSnapshot = cloneDocument(currentDocument);
  fileName.textContent = result.fileName;
  fileName.title = result.fileName;
  documentTitle.textContent = result.fileName;
  applyDocumentIdentity(currentDocument);
  rememberRecentDocument(currentDocument);
  setDirty(false);
  if (currentDocument.kind === "workbook" && previousExtension !== currentDocument.extension) {
    renderWorkbook(currentDocument); // .csv <-> .xlsx changes the Add-sheet affordance
  }
  setStatus(result.backupPath ? "Saved with backup" : "Saved");
}

function buildSavePayload() {
  if (currentDocument.kind === "word") {
    const html = sanitizeWordHtml(wordEditor?.innerHTML || currentDocument.html || "");
    return {
      kind: "word",
      filePath: currentDocument.filePath,
      defaultFileName: currentDocument.fileName || "Untitled document.docx",
      extension: currentDocument.extension,
      html,
      // text/markdown let Save As export to .txt/.md.
      text: wordEditor?.innerText || htmlToPlainText(html),
      markdown: htmlToMarkdown(html)
    };
  }

  if (currentDocument.kind === "markdown" || currentDocument.kind === "text") {
    const payload = {
      kind: currentDocument.kind,
      filePath: currentDocument.filePath,
      defaultFileName: currentDocument.fileName || (currentDocument.kind === "text" ? "Untitled.txt" : "Untitled.md"),
      extension: currentDocument.extension,
      text: currentDocument.text || ""
    };
    if (currentDocument.kind === "markdown") {
      // html lets Save As export to .docx via the existing Word writer.
      payload.html = markdownToHtml(currentDocument.text || "");
    }
    return payload;
  }

  if (currentDocument.kind === "workbook") {
    return {
      kind: "workbook",
      filePath: currentDocument.filePath,
      defaultFileName: currentDocument.fileName || "Untitled workbook.xlsx",
      extension: currentDocument.extension,
      sheets: currentDocument.sheets
    };
  }

  if (currentDocument.kind === "image") {
    return {
      kind: "image",
      filePath: currentDocument.filePath,
      defaultFileName: currentDocument.fileName || "Untitled image.png",
      extension: currentDocument.extension,
      dataUrl: currentDocument.dataUrl || "",
      sourcePath: currentDocument.filePath || ""
    };
  }

  return null;
}

function syncCurrentDocumentFromDom() {
  if (!currentDocument) {
    return;
  }

  if (currentDocument.kind === "word" && wordEditor) {
    currentDocument.html = sanitizeWordHtml(wordEditor.innerHTML);
    currentDocument.text = wordEditor.innerText;
  } else if (currentDocument.kind === "markdown" || currentDocument.kind === "text") {
    const markdownEditor = viewer.querySelector(".markdown-editor");
    if (markdownEditor) {
      currentDocument.text = markdownEditor.value;
    }
  } else if (currentDocument.kind === "workbook") {
    syncWorkbookFromDom();
  } else if (currentDocument.kind === "image" && imageCanvas && viewer.contains(imageCanvas)) {
    commitPaintText(); // an open text box counts as content — land it before snapshotting
    currentDocument.dataUrl = imageCanvas.toDataURL(
      currentDocument.extension === ".jpg" || currentDocument.extension === ".jpeg" ? "image/jpeg" : "image/png"
    );
  }
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

function hasDraggedFiles(event) {
  return Array.from(event.dataTransfer?.types || []).includes("Files");
}

function sheetSummary(workbookData) {
  if (workbookData.extension === ".csv") {
    return "CSV sheet";
  }

  return `${workbookData.sheets.length} sheet${workbookData.sheets.length === 1 ? "" : "s"}`;
}

function pdfSummary(pdfData) {
  return `PDF document - read-only${pdfData.size ? ` - ${formatFileSize(pdfData.size)}` : ""}`;
}

function documentSummary(documentData) {
  if (!documentData) {
    return "";
  }

  if (documentData.kind === "word") {
    return warningSummary(documentData.warnings);
  }

  if (documentData.kind === "markdown") {
    return "Markdown document";
  }

  if (documentData.kind === "text") {
    if (documentData.extension === ".json") {
      return "JSON file";
    }
    if (documentData.extension === ".log") {
      return "Log file";
    }
    return "Text file";
  }

  if (documentData.kind === "image") {
    const suffix = documentData.size ? ` - ${formatFileSize(documentData.size)}` : "";
    return isEditableImage(documentData) ? `Image${suffix}` : `Image - read-only${suffix}`;
  }

  if (documentData.kind === "workbook") {
    return sheetSummary(documentData);
  }

  if (documentData.kind === "pdf") {
    return pdfSummary(documentData);
  }

  return "Document";
}

function isEditableDocument(documentData) {
  return Boolean(
    documentData &&
      (documentData.kind === "word" ||
        documentData.kind === "workbook" ||
        documentData.kind === "markdown" ||
        documentData.kind === "text" ||
        isEditableImage(documentData))
  );
}

// Canvas edits save as PNG/JPEG; other image formats stay view-only.
function isEditableImage(documentData) {
  return Boolean(
    documentData?.kind === "image" &&
      (documentData.isNew || [".png", ".jpg", ".jpeg"].includes(documentData.extension))
  );
}

function canSaveCurrentDocument() {
  return isEditableDocument(currentDocument);
}

function formatFileSize(bytes) {
  if (bytes < 1024) {
    return `${bytes} B`;
  }

  if (bytes < 1024 * 1024) {
    return `${Math.round(bytes / 1024)} KB`;
  }

  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function formatDateTime(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "-";
  }

  return date.toLocaleString(undefined, {
    year: "numeric",
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit"
  });
}

function applyDocumentIdentity(documentData) {
  const identity = documentIdentity(documentData);
  document.body.dataset.documentKind = identity.bodyKind;
  currentFilePanel.hidden = !documentData;
  documentKind.hidden = !documentData;
  documentKind.textContent = identity.icon;
  documentKind.title = identity.label;
  documentKind.setAttribute("aria-label", identity.label);
  documentKind.className = `document-kind ${identity.className}`;
  fileKindIcon.textContent = identity.icon;
  fileKindIcon.className = `file-kind-icon ${identity.className}`;
  renderDocumentNotices(documentData);
  renderDocumentInfo(documentData);
  renderBackupHistory(documentData);
  updateLayoutControls();
}

function documentIdentity(documentData) {
  if (!documentData) {
    return {
      bodyKind: "empty",
      className: "kind-empty",
      icon: "-",
      label: "No file"
    };
  }

  if (documentData.kind === "word") {
    return {
      bodyKind: "word",
      className: "kind-word",
      icon: "DOC",
      label: "Word"
    };
  }

  if (documentData.kind === "markdown") {
    return {
      bodyKind: "markdown",
      className: "kind-markdown",
      icon: "MD",
      label: "Markdown"
    };
  }

  if (documentData.kind === "text") {
    return {
      bodyKind: "text",
      className: "kind-text",
      icon: documentData.extension === ".json" ? "{ }" : documentData.extension === ".log" ? "LOG" : "TXT",
      label: "Text"
    };
  }

  if (documentData.kind === "image") {
    return {
      bodyKind: "image",
      className: "kind-image",
      icon: "IMG",
      label: isEditableImage(documentData) ? "Image" : "Image read-only"
    };
  }

  if (documentData.kind === "workbook" && documentData.extension === ".csv") {
    return {
      bodyKind: "csv",
      className: "kind-csv",
      icon: "CSV",
      label: "CSV"
    };
  }

  if (documentData.kind === "workbook") {
    return {
      bodyKind: "workbook",
      className: "kind-workbook",
      icon: "XLS",
      label: "Workbook"
    };
  }

  if (documentData.kind === "pdf") {
    return {
      bodyKind: "pdf",
      className: "kind-pdf",
      icon: "PDF",
      label: "PDF read-only"
    };
  }

  return {
    bodyKind: "document",
    className: "kind-empty",
    icon: "DOC",
    label: "Document"
  };
}

function loadRecentFiles() {
  try {
    const parsed = JSON.parse(window.localStorage.getItem(RECENT_FILES_KEY) || "[]");
    if (!Array.isArray(parsed)) {
      return [];
    }

    return parsed.filter((item) => item && typeof item.filePath === "string" && typeof item.fileName === "string");
  } catch {
    return [];
  }
}

function saveRecentFiles() {
  try {
    window.localStorage.setItem(RECENT_FILES_KEY, JSON.stringify(recentFiles));
  } catch {
    // Recent files are helpful, but never worth blocking document work.
  }
  window.documentOpener.syncRecentFiles?.(recentFiles);
}

function loadBackupHistory() {
  try {
    const parsed = JSON.parse(window.localStorage.getItem(BACKUP_HISTORY_KEY) || "[]");
    if (!Array.isArray(parsed)) {
      return [];
    }

    return parsed
      .filter((item) => item && typeof item.backupPath === "string" && typeof item.sourcePath === "string")
      .slice(0, MAX_BACKUP_HISTORY);
  } catch {
    return [];
  }
}

function saveBackupHistory() {
  try {
    window.localStorage.setItem(BACKUP_HISTORY_KEY, JSON.stringify(backupHistory.slice(0, MAX_BACKUP_HISTORY)));
  } catch {
    // Backup history is a convenience; actual backup files remain on disk.
  }
}

function rememberRecentDocument(documentData) {
  if (!documentData?.filePath || !documentData?.fileName) {
    return;
  }

  const nextEntry = {
    filePath: documentData.filePath,
    fileName: documentData.fileName,
    kind: documentData.kind,
    extension: documentData.extension || extensionFromPath(documentData.filePath),
    summary: documentSummary(documentData),
    openedAt: Date.now()
  };

  recentFiles = [nextEntry, ...recentFiles.filter((item) => !sameFilePath(item.filePath, nextEntry.filePath))];
  saveRecentFiles();
  renderRecentFiles();
}

function removeRecentFile(filePath) {
  const nextFiles = recentFiles.filter((item) => !sameFilePath(item.filePath, filePath));
  if (nextFiles.length === recentFiles.length) {
    return;
  }

  recentFiles = nextFiles;
  saveRecentFiles();
  renderRecentFiles();
}

function renderRecentFiles() {
  if (!recentPanel || !recentList) {
    return;
  }

  recentList.replaceChildren();
  clearRecentButton.hidden = recentFiles.length === 0;

  if (recentFiles.length === 0) {
    const empty = document.createElement("div");
    empty.className = "recent-empty";
    empty.textContent = "No recent files";
    recentList.appendChild(empty);
    return;
  }

  const items = recentFiles.map((recentFile) => {
    const button = document.createElement("button");
    const identity = documentIdentity(recentFile);
    button.type = "button";
    button.className = "recent-file";
    button.title = recentFile.filePath;
    button.addEventListener("click", () => {
      openRecentDocument(recentFile);
    });

    const icon = document.createElement("span");
    icon.className = `recent-file-icon ${identity.className}`;
    icon.textContent = identity.icon;

    const copy = document.createElement("span");
    copy.className = "recent-file-copy";

    const name = document.createElement("strong");
    name.textContent = recentFile.fileName;

    const meta = document.createElement("span");
    meta.textContent = recentFile.summary || identity.label;

    copy.append(name, meta);
    button.append(icon, copy);
    return button;
  });

  recentList.append(...items);
}

async function openRecentDocument(recentFile) {
  setStatus("Opening...");
  const result = await window.documentOpener.openRecentFile(recentFile.filePath);
  if (!result?.ok) {
    removeRecentFile(recentFile.filePath);
  }
  handleOpenResult(result);
}

function sameFilePath(left, right) {
  return String(left).toLowerCase() === String(right).toLowerCase();
}

function updateSaveButtonLabel() {
  if (!currentDocument || !isEditableDocument(currentDocument)) {
    saveButtonLabel.textContent = "Save";
    saveButton.classList.remove("is-clean", "is-dirty");
    return;
  }

  saveButtonLabel.textContent = isDirty ? "Save" : "Saved";
  saveButton.classList.toggle("is-clean", !isDirty);
  saveButton.classList.toggle("is-dirty", isDirty);
}

function extensionFromPath(filePath) {
  const match = filePath.match(/\.[^.\\\/]+$/);
  return match ? match[0].toLowerCase() : "";
}

function baseName(filePath) {
  return String(filePath || "").split(/[\\/]/).pop() || String(filePath || "");
}

function updateToolbarForDocument() {
  const hasVisibleToolset =
    currentDocument &&
    (currentDocument.kind === "word" || currentDocument.kind === "workbook" || currentDocument.kind === "image");
  editToolbar.hidden = !editMode || !hasVisibleToolset;
  const toolsets = editToolbar.querySelectorAll("[data-toolset]");
  toolsets.forEach((toolset) => {
    toolset.hidden = !currentDocument || toolset.dataset.toolset !== currentDocument.kind;
  });
  updateSelectionLabel();
  updateLayoutControls();
  updateRawViewControls();
}

// Raw view is a VIEW-mode alternate rendering: markdown preview <-> raw text,
// CSV grid <-> raw CSV text, JSON pretty/highlighted <-> text as on disk.
function supportsRawView(documentData) {
  if (!documentData) {
    return false;
  }
  return (
    documentData.kind === "markdown" ||
    isJsonDocument(documentData) ||
    (documentData.kind === "workbook" && documentData.extension === ".csv")
  );
}

function isJsonDocument(documentData) {
  return documentData?.kind === "text" && documentData.extension === ".json";
}

function updateRawViewControls() {
  rawViewButton.hidden = editMode || !supportsRawView(currentDocument);
  rawViewButton.setAttribute("aria-pressed", String(Boolean(currentDocument?.rawView)));
  formatJsonButton.hidden = !editMode || !isJsonDocument(currentDocument);
}

function setDocumentLayout(width) {
  if (!currentDocument || !["narrow", "wide", "full"].includes(width)) {
    return;
  }

  syncCurrentDocumentFromDom();
  documentLayout = width;
  currentDocument.layoutWidth = width;
  if (currentDocument.kind === "word") {
    currentDocument.wordWidth = width;
    renderWord(currentDocument);
  } else if (currentDocument.kind === "markdown" || currentDocument.kind === "text") {
    renderMarkdown(currentDocument);
  } else if (currentDocument.kind === "workbook") {
    renderWorkbook(currentDocument);
  } else if (currentDocument.kind === "pdf") {
    renderPdf(currentDocument);
  } else if (currentDocument.kind === "image") {
    renderImageDocument(currentDocument);
  }
  updateLayoutControls();
}

function getDocumentLayout(documentData = currentDocument) {
  return documentData?.layoutWidth || documentData?.wordWidth || documentLayout || DEFAULT_DOCUMENT_LAYOUT;
}

function updateLayoutControls() {
  if (!layoutSeg) {
    return;
  }

  layoutSeg.hidden = !currentDocument || currentDocument.kind === "image" || currentDocument.kind === "pdf";
  const activeWidth = getDocumentLayout();
  layoutSeg.querySelectorAll("[data-layout]").forEach((button) => {
    button.setAttribute("aria-pressed", String(button.dataset.layout === activeWidth));
  });
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

function captureWordSelection() {
  if (!wordEditor) {
    savedWordRange = null;
    return;
  }

  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) {
    return;
  }

  const range = selection.getRangeAt(0);
  if (wordEditor.contains(range.commonAncestorContainer)) {
    savedWordRange = range.cloneRange();
  }

  updateTableToolsVisibility();
}

function updateTableToolsVisibility() {
  if (!tableTools) {
    return;
  }

  const inTable = Boolean(editMode && currentDocument?.kind === "word" && currentTableContext());
  tableTools.hidden = !inTable;
}

function restoreWordSelection() {
  if (!savedWordRange) {
    return;
  }

  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(savedWordRange);
}

// Focus the editor with a guaranteed-valid caret: the saved range if it still
// lives inside the editor, otherwise the end of the document. Needed for insert
// commands that run after the inline prompt bar stole focus.
function focusEditorAtSavedRange() {
  if (!wordEditor) {
    return;
  }

  wordEditor.focus();
  const selection = window.getSelection();

  if (savedWordRange && wordEditor.contains(savedWordRange.commonAncestorContainer)) {
    selection.removeAllRanges();
    selection.addRange(savedWordRange);
    return;
  }

  const range = document.createRange();
  range.selectNodeContents(wordEditor);
  range.collapse(false);
  selection.removeAllRanges();
  selection.addRange(range);
}

function runWordCommand(command, value = null) {
  if (!wordEditor || currentDocument.kind !== "word") {
    return;
  }

  wordEditor.focus();
  restoreWordSelection();
  document.execCommand(command, false, value);
  captureWordSelection();
  syncCurrentDocumentFromDom();
  setDirty(true);
}

function runWordBlockCommand(blockName) {
  if (!wordEditor || currentDocument.kind !== "word") {
    return;
  }

  wordEditor.focus();
  restoreWordSelection();
  document.execCommand("formatBlock", false, blockName);
  captureWordSelection();
  syncCurrentDocumentFromDom();
  setDirty(true);
}

function runWordAction(action) {
  if (!wordEditor || currentDocument.kind !== "word") {
    return;
  }

  if (action === "link") {
    insertLink();
  } else if (action === "unlink") {
    runWordCommand("unlink");
  } else if (action === "image") {
    imageInput.click();
  } else if (action === "table") {
    insertTable();
  } else if (action === "find") {
    toggleFindBar();
  } else if (action === "row-add" || action === "row-del" || action === "col-add" || action === "col-del") {
    runTableAction(action);
  }
}

// Find the table cell the caret sits in, using the live selection or the last
// captured range. Returns null when the caret is not inside an editor table.
function currentTableContext() {
  const selection = window.getSelection();
  let node = selection && selection.rangeCount > 0 ? selection.getRangeAt(0).commonAncestorContainer : null;
  if (!node || !wordEditor.contains(node)) {
    node = savedWordRange ? savedWordRange.commonAncestorContainer : null;
  }
  if (!node || !wordEditor.contains(node)) {
    return null;
  }

  const element = node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement;
  const cell = element?.closest("td, th");
  if (!cell || !wordEditor.contains(cell)) {
    return null;
  }

  const row = cell.closest("tr");
  const table = cell.closest("table");
  if (!row || !table || !wordEditor.contains(table)) {
    return null;
  }

  const cellIndex = Array.prototype.indexOf.call(row.children, cell);
  return { cell, row, table, cellIndex };
}

function runTableAction(action) {
  const context = currentTableContext();
  if (!context) {
    setStatus("Place the cursor in a table");
    return;
  }

  const { row, table, cellIndex } = context;

  if (action === "row-add") {
    const newRow = document.createElement("tr");
    for (let index = 0; index < row.children.length; index += 1) {
      const td = document.createElement("td");
      td.innerHTML = "<br />";
      newRow.appendChild(td);
    }
    row.after(newRow);
  } else if (action === "row-del") {
    if (table.rows.length <= 1) {
      table.remove();
    } else {
      row.remove();
    }
  } else if (action === "col-add") {
    Array.from(table.rows).forEach((tableRow) => {
      const td = document.createElement("td");
      td.innerHTML = "<br />";
      const reference = tableRow.children[cellIndex];
      if (reference) {
        reference.after(td);
      } else {
        tableRow.appendChild(td);
      }
    });
  } else if (action === "col-del") {
    const columnCount = Math.max(...Array.from(table.rows).map((tableRow) => tableRow.children.length));
    if (columnCount <= 1) {
      table.remove();
    } else {
      Array.from(table.rows).forEach((tableRow) => {
        tableRow.children[cellIndex]?.remove();
      });
    }
  }

  wordEditor.focus();
  captureWordSelection();
  syncCurrentDocumentFromDom();
  setDirty(true);
}

async function insertLink() {
  captureWordSelection();
  const url = await askInline("Link address", "https://", {
    message: "Enter the destination URL for the selected text.",
    primaryLabel: "Insert"
  });
  if (!url) {
    return;
  }

  focusEditorAtSavedRange();
  const selection = window.getSelection();
  const hasText = selection && !selection.isCollapsed && wordEditor.contains(selection.anchorNode);
  if (hasText) {
    document.execCommand("createLink", false, url);
  } else {
    document.execCommand("insertHTML", false, `<a href="${escapeHtml(url)}">${escapeHtml(url)}</a>`);
  }

  captureWordSelection();
  syncCurrentDocumentFromDom();
  setDirty(true);
}

async function insertTable() {
  captureWordSelection();
  const spec = await askInline("Table size", "3x3", {
    message: "Enter rows x columns.",
    primaryLabel: "Insert"
  });
  if (!spec) {
    return;
  }

  const match = /^\s*(\d+)\s*[x×,\s]\s*(\d+)\s*$/i.exec(spec);
  if (!match) {
    setStatus("Enter a size like 3x3");
    return;
  }

  const rows = clampInt(match[1], 1, 50);
  const columns = clampInt(match[2], 1, 20);
  focusEditorAtSavedRange();

  let html = '<table class="doc-table"><tbody>';
  for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
    html += "<tr>";
    for (let columnIndex = 0; columnIndex < columns; columnIndex += 1) {
      html += "<td><br /></td>";
    }
    html += "</tr>";
  }
  html += "</tbody></table><p><br /></p>";

  const inserted = document.execCommand("insertHTML", false, html);
  if (!inserted || !wordEditor.querySelector("table")) {
    wordEditor.insertAdjacentHTML("beforeend", html);
  }
  captureWordSelection();
  syncCurrentDocumentFromDom();
  setDirty(true);
}

function toggleFindBar(forceShow) {
  if (!findBar) {
    return;
  }

  const show = typeof forceShow === "boolean" ? forceShow : findBar.hidden;
  findBar.hidden = !show;

  if (show) {
    resolveInlinePrompt(null);
    findInput.focus();
    findInput.select();
  }
}

// Concatenate the editor's text nodes into one string plus an index map, so
// searches can match across formatting boundaries (e.g. a bold-split word).
function buildTextMap(root = wordEditor) {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes = [];
  let text = "";
  while (walker.nextNode()) {
    const node = walker.currentNode;
    nodes.push({ node, start: text.length });
    text += node.nodeValue;
  }
  return { text, nodes };
}

function locateInMap(map, index) {
  for (const entry of map.nodes) {
    if (index <= entry.start + entry.node.nodeValue.length) {
      return { node: entry.node, offset: index - entry.start };
    }
  }
  const last = map.nodes[map.nodes.length - 1];
  return { node: last.node, offset: last.node.nodeValue.length };
}

function rangeFromIndices(map, startIndex, endIndex) {
  const start = locateInMap(map, startIndex);
  const end = locateInMap(map, endIndex);
  const range = document.createRange();
  range.setStart(start.node, start.offset);
  range.setEnd(end.node, end.offset);
  return range;
}

function selectEditorRange(range) {
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);
  const anchor = range.startContainer.nodeType === Node.ELEMENT_NODE ? range.startContainer : range.startContainer.parentElement;
  anchor?.scrollIntoView({ block: "center", inline: "nearest" });
}

let findCursor = { term: null, index: 0 };

const FINDABLE_KINDS = new Set(["word", "markdown", "text", "workbook"]);

// Shared cursor-based substring search: returns the match index or -1,
// wrapping around when nothing is found past the cursor.
function nextMatchIndex(haystack, needle, backwards) {
  if (findCursor.term !== needle) {
    findCursor = { term: needle, index: 0 };
  }

  let index;
  if (backwards) {
    index = haystack.lastIndexOf(needle, Math.max(0, findCursor.index - needle.length - 1));
    if (index === -1) {
      index = haystack.lastIndexOf(needle);
    }
  } else {
    index = haystack.indexOf(needle, findCursor.index);
    if (index === -1) {
      index = haystack.indexOf(needle);
    }
  }

  if (index !== -1) {
    findCursor.index = backwards ? index : index + needle.length;
  }
  return index;
}

function findInDocument(backwards) {
  const term = findInput.value;
  if (!term || !currentDocument) {
    return;
  }

  if (currentDocument.kind === "workbook") {
    findInWorkbook(term, backwards);
    return;
  }

  const textarea = viewer.querySelector(".markdown-editor");
  if (textarea) {
    findInTextarea(textarea, term, backwards);
    return;
  }

  const root = wordEditor || viewer.querySelector(".markdown-document, .text-document");
  if (!root) {
    setStatus("No matches");
    return;
  }

  const map = buildTextMap(root);
  if (map.nodes.length === 0) {
    setStatus("No matches");
    return;
  }

  const index = nextMatchIndex(map.text.toLowerCase(), term.toLowerCase(), backwards);
  if (index === -1) {
    setStatus("No matches");
    return;
  }

  selectEditorRange(rangeFromIndices(map, index, index + term.length));
}

function findInTextarea(textarea, term, backwards) {
  const index = nextMatchIndex(textarea.value.toLowerCase(), term.toLowerCase(), backwards);
  if (index === -1) {
    setStatus("No matches");
    return;
  }

  textarea.setSelectionRange(index, index + term.length);
  const lineHeight = parseFloat(getComputedStyle(textarea).lineHeight) || 20;
  const line = textarea.value.slice(0, index).split("\n").length - 1;
  textarea.scrollTop = Math.max(0, line * lineHeight - textarea.clientHeight / 2);
}

// ponytail: searches the raw model value, so formula cells match on their
// "=..." text rather than the computed result.
function findInWorkbook(term, backwards) {
  syncWorkbookFromDom();
  const sheet = currentDocument.sheets[selectedSheetIndex];
  if (!sheet) {
    return;
  }

  const needle = term.toLowerCase();
  const matches = [];
  sheet.rows.forEach((row, rowIndex) => {
    row.forEach((value, columnIndex) => {
      if (String(value ?? "").toLowerCase().includes(needle)) {
        matches.push({ rowIndex, columnIndex });
      }
    });
  });

  if (matches.length === 0) {
    setStatus("No matches");
    return;
  }

  const position = (cell) => cell.rowIndex * 100000 + cell.columnIndex;
  const current = selectedCell ? position(selectedCell) : -1;
  let target;
  if (backwards) {
    target = [...matches].reverse().find((cell) => position(cell) < current) || matches[matches.length - 1];
  } else {
    target = matches.find((cell) => position(cell) > current) || matches[0];
  }

  selectWorkbookCell(target.rowIndex, target.columnIndex);
  viewer
    .querySelector(`td[data-row-index="${target.rowIndex}"][data-column-index="${target.columnIndex}"]`)
    ?.scrollIntoView({ block: "center", inline: "nearest" });
  setStatus(`${matches.indexOf(target) + 1} of ${matches.length} matches`);
}

function replaceCurrentMatch() {
  if (!editMode || !currentDocument) {
    return;
  }

  const term = findInput.value;
  if (!term) {
    return;
  }

  if (currentDocument.kind === "workbook") {
    replaceInWorkbookCell(term);
    return;
  }

  const textarea = viewer.querySelector(".markdown-editor");
  if (textarea) {
    replaceInTextarea(textarea, term);
    return;
  }

  if (!wordEditor) {
    return;
  }

  const selection = window.getSelection();
  const matchesTerm =
    selection &&
    selection.rangeCount > 0 &&
    !selection.isCollapsed &&
    wordEditor.contains(selection.anchorNode) &&
    selection.toString().toLowerCase() === term.toLowerCase();

  if (matchesTerm) {
    const range = selection.getRangeAt(0);
    const replacementNode = document.createTextNode(replaceInput.value);
    range.deleteContents();
    range.insertNode(replacementNode);
    const caret = document.createRange();
    caret.setStartAfter(replacementNode);
    caret.collapse(true);
    selection.removeAllRanges();
    selection.addRange(caret);
    findCursor = { term: null, index: 0 };
    captureWordSelection();
    syncCurrentDocumentFromDom();
    setDirty(true);
  }

  findInDocument(false);
}

function replaceInTextarea(textarea, term) {
  const start = textarea.selectionStart;
  const end = textarea.selectionEnd;
  const selected = textarea.value.slice(start, end);
  if (selected.toLowerCase() === term.toLowerCase() && start !== end) {
    textarea.value = textarea.value.slice(0, start) + replaceInput.value + textarea.value.slice(end);
    textarea.dispatchEvent(new Event("input", { bubbles: true }));
    findCursor = { term: term.toLowerCase(), index: start + replaceInput.value.length };
  }

  findInDocument(false);
}

function replaceInWorkbookCell(term) {
  const sheet = currentDocument.sheets[selectedSheetIndex];
  const row = sheet && selectedCell ? sheet.rows[selectedCell.rowIndex] : null;
  if (row) {
    const value = String(row[selectedCell.columnIndex] ?? "");
    const index = value.toLowerCase().indexOf(term.toLowerCase());
    if (index !== -1) {
      row[selectedCell.columnIndex] = value.slice(0, index) + replaceInput.value + value.slice(index + term.length);
      setDirty(true);
      renderWorkbook(currentDocument);
    }
  }

  findInDocument(false);
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function replaceAllMatches() {
  if (!editMode || !currentDocument) {
    return;
  }

  const term = findInput.value;
  if (!term) {
    return;
  }

  const pattern = new RegExp(escapeRegExp(term), "gi");
  const replacement = replaceInput.value;
  let replaced = 0;

  if (currentDocument.kind === "workbook") {
    syncWorkbookFromDom();
    const sheet = currentDocument.sheets[selectedSheetIndex];
    if (!sheet) {
      return;
    }
    sheet.rows.forEach((row) => {
      row.forEach((value, columnIndex) => {
        const text = String(value ?? "");
        const matches = text.match(pattern);
        if (matches) {
          replaced += matches.length;
          row[columnIndex] = text.replace(pattern, replacement);
        }
      });
    });
    if (replaced > 0) {
      setDirty(true);
      renderWorkbook(currentDocument);
    }
  } else if (viewer.querySelector(".markdown-editor")) {
    const textarea = viewer.querySelector(".markdown-editor");
    const matches = textarea.value.match(pattern);
    if (matches) {
      replaced = matches.length;
      textarea.value = textarea.value.replace(pattern, replacement);
      textarea.dispatchEvent(new Event("input", { bubbles: true }));
    }
  } else if (wordEditor) {
    const needle = term.toLowerCase();
    // Rebuild the map after each replacement (DOM shifts); track a search
    // offset so a replacement containing the term isn't re-matched.
    let searchFrom = 0;
    for (;;) {
      const map = buildTextMap();
      const index = map.text.toLowerCase().indexOf(needle, searchFrom);
      if (index === -1) {
        break;
      }
      const range = rangeFromIndices(map, index, index + needle.length);
      range.deleteContents();
      range.insertNode(document.createTextNode(replacement));
      searchFrom = index + replacement.length;
      replaced += 1;
    }
    if (replaced > 0) {
      captureWordSelection();
      syncCurrentDocumentFromDom();
      setDirty(true);
    }
  }

  findCursor = { term: null, index: 0 };
  setStatus(replaced > 0 ? `Replaced ${replaced} match${replaced === 1 ? "" : "es"}` : "No matches");
}

function askInline(label, defaultValue = "", options = {}) {
  toggleFindBar(false);
  resolveInlinePrompt(null);
  return showPrompt({
    title: label,
    message: options.message || "",
    defaultValue,
    inputLabel: label,
    primaryLabel: options.primaryLabel || "OK",
    cancelLabel: "Cancel"
  });
}

function resolveInlinePrompt(value) {
  if (!inlinePromptResolver) {
    return;
  }

  const resolve = inlinePromptResolver;
  inlinePromptResolver = null;
  promptBar.hidden = true;
  resolve(value);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function clampInt(value, min, max) {
  return Math.max(min, Math.min(max, Math.round(Number(value)) || min));
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

function revertToSavedSnapshot() {
  if (!lastSavedSnapshot) {
    return;
  }

  currentDocument = cloneDocument(lastSavedSnapshot);
  if (tabs[activeTabIndex]) {
    tabs[activeTabIndex].document = currentDocument;
  }
  selectedSheetIndex = 0;
  selectedCell = null;
  selectedRange = null;
  rangeAnchor = null;
  fileName.textContent = currentDocument.fileName;
  fileName.title = currentDocument.fileName;
  documentTitle.textContent = currentDocument.fileName;
  applyDocumentIdentity(currentDocument);
  updateLayoutControls();
  setDirty(false);

  if (currentDocument.kind === "word") {
    renderWord(currentDocument);
  } else if (currentDocument.kind === "markdown" || currentDocument.kind === "text") {
    renderMarkdown(currentDocument);
  } else if (currentDocument.kind === "workbook") {
    renderWorkbook(currentDocument);
  } else if (currentDocument.kind === "pdf") {
    renderPdf(currentDocument);
  } else if (currentDocument.kind === "image") {
    renderImageDocument(currentDocument);
  }

  setStatus("Reverted");
}

function cloneDocument(documentData) {
  return JSON.parse(JSON.stringify(documentData));
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

// --- Crash recovery --------------------------------------------------------
// Every 30s the dirty tabs are snapshotted to userData/autosave.json; a normal
// quit deletes the file (main.js before-quit), so finding one on launch means
// the last session crashed. Worst case loses 30s of typing.

const AUTOSAVE_INTERVAL_MS = 30_000;
let autosaveHadData = false;

async function autosaveTick() {
  snapshotActiveTab();
  const dirtyTabs = tabs.filter((tab) => tab.isDirty);
  if (dirtyTabs.length === 0) {
    if (autosaveHadData) {
      autosaveHadData = false;
      window.documentOpener.clearAutosave();
    }
    return;
  }

  autosaveHadData = true;
  await window.documentOpener.writeAutosave({
    savedAt: new Date().toISOString(),
    documents: dirtyTabs.map((tab) => ({ document: tab.document, editMode: tab.editMode }))
  });
}

async function restoreAutosave() {
  const data = await window.documentOpener.readAutosave();
  const documents = Array.isArray(data?.documents) ? data.documents : [];
  if (documents.length === 0) {
    return;
  }

  const count = documents.length;
  const confirmed = await showConfirm({
    title: "Restore unsaved work?",
    message: `DocOpen closed unexpectedly with ${count} unsaved document${count === 1 ? "" : "s"}. Restore ${count === 1 ? "it" : "them"}?`,
    tone: "warning",
    primaryLabel: "Restore",
    cancelLabel: "Discard"
  });
  if (!confirmed) {
    window.documentOpener.clearAutosave();
    return;
  }

  for (const entry of documents) {
    if (!entry?.document?.kind) {
      continue;
    }
    snapshotActiveTab();
    const tab = createDocumentTab(entry.document);
    // ponytail: lastSavedSnapshot is the restored (dirty) state, so Revert
    // returns here rather than to the on-disk file after a restore.
    tab.isDirty = true;
    tab.editMode = Boolean(entry.editMode) && isEditableDocument(entry.document);
    tabs.push(tab);
    activeTabIndex = tabs.length - 1;
  }
  presentActiveTab();
  setStatus("Restored unsaved work");
}

setInterval(autosaveTick, AUTOSAVE_INTERVAL_MS);
restoreAutosave();
