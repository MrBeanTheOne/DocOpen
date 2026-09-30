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
// The welcome empty state ships in index.html; keep a copy so closing the
// last tab restores it instead of a bare heading.
const emptyStateHtml = viewer.innerHTML;
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
    closeSettings();
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
  } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "o") {
    event.preventDefault();
    openButton.click();
  } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "n") {
    event.preventDefault();
    toggleNewMenu();
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
  viewer.innerHTML = emptyStateHtml;
}

function emptyBlock(message) {
  const block = document.createElement("div");
  block.className = "empty-state";
  const heading = document.createElement("h3");
  heading.textContent = message;
  block.appendChild(heading);
  return block;
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

function kindForExtension(extension) {
  if (extension === ".docx") {
    return "word";
  }
  if ([".md", ".markdown"].includes(extension)) {
    return "markdown";
  }
  if ([".txt", ".log", ".json", ".jsonl", ".yaml", ".yml", ".xml", ".ini", ".cfg", ".conf"].includes(extension)) {
    return "text";
  }
  if ([".xlsx", ".csv"].includes(extension)) {
    return "workbook";
  }
  if ([".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".svg", ".ico", ".avif"].includes(extension)) {
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
    modeLabel.textContent = "Nothing open";
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

function clampInt(value, min, max) {
  return Math.max(min, Math.min(max, Math.round(Number(value)) || min));
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

// --- Settings --------------------------------------------------------------
// localStorage is the source of truth (same deal as recent files); main gets a
// copy of the ones it enforces through settings:sync. Changes apply on tick —
// the dialog has no Save button.

const SETTINGS_KEY = "docopen.settings";
const DEFAULT_SETTINGS = { closeToTray: true, checkUpdatesOnLaunch: true };

const settingsButton = document.querySelector("#settingsButton");
const settingsOverlay = document.querySelector("#settingsOverlay");
const settingsCloseButton = document.querySelector("#settingsCloseButton");
const closeToTraySetting = document.querySelector("#closeToTraySetting");
const checkUpdatesSetting = document.querySelector("#checkUpdatesSetting");

const settings = loadSettings();

function loadSettings() {
  try {
    const parsed = JSON.parse(window.localStorage.getItem(SETTINGS_KEY) || "{}");
    return { ...DEFAULT_SETTINGS, ...(parsed && typeof parsed === "object" ? parsed : {}) };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

function saveSettings() {
  try {
    window.localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
  } catch {
    // A settings write failing is not worth interrupting document work over.
  }
  return window.documentOpener.syncSettings?.(settings);
}

function openSettings() {
  closeToTraySetting.checked = settings.closeToTray;
  checkUpdatesSetting.checked = settings.checkUpdatesOnLaunch;
  settingsOverlay.hidden = false;
  document.body.classList.add("has-dialog");
  settingsCloseButton.focus();
}

function closeSettings() {
  settingsOverlay.hidden = true;
  document.body.classList.remove("has-dialog");
}

settingsButton?.addEventListener("click", openSettings);
settingsCloseButton?.addEventListener("click", closeSettings);
settingsOverlay?.addEventListener("click", (event) => {
  if (event.target === settingsOverlay) {
    closeSettings();
  }
});

closeToTraySetting?.addEventListener("change", () => {
  settings.closeToTray = closeToTraySetting.checked;
  saveSettings();
});

checkUpdatesSetting?.addEventListener("change", () => {
  settings.checkUpdatesOnLaunch = checkUpdatesSetting.checked;
  saveSettings();
});

// Main starts on the defaults, so mirror the stored values at every launch.
saveSettings();

// --- Auto-update -----------------------------------------------------------
// Main relays electron-updater's events here; the sidebar foot shows where the
// download is and turns into a Restart button once the installer is staged.

const updateText = document.querySelector("#updateText");
const updateButton = document.querySelector("#updateButton");
const updateBar = document.querySelector("#updateBar");
const updateBarFill = document.querySelector("#updateBarFill");

let updateState = "idle";
let updateIdleLabel = "DocOpen";
let updatePendingVersion = "";
let updateIdleTimer = null;

function setUpdateIdle() {
  applyUpdateStatus({ state: "idle" });
}

function applyUpdateStatus(status) {
  window.clearTimeout(updateIdleTimer);
  updateState = status.state;
  updateButton.hidden = false;
  updateButton.textContent = "Check";
  updateButton.disabled = false;
  updateBar.hidden = true;
  updateText.title = "";
  updateText.classList.remove("is-ready");

  if (status.version) {
    updatePendingVersion = status.version;
  }

  switch (status.state) {
    case "checking":
      updateText.textContent = "Checking for updates...";
      updateButton.disabled = true;
      break;
    case "downloading": {
      const percent = typeof status.percent === "number" ? status.percent : 0;
      updateText.textContent = `Downloading v${updatePendingVersion} - ${percent}%`;
      updateButton.hidden = true;
      updateBar.hidden = false;
      updateBarFill.style.width = `${percent}%`;
      break;
    }
    case "ready":
      updateText.textContent = `v${updatePendingVersion} ready to install`;
      updateText.classList.add("is-ready");
      updateButton.textContent = "Restart";
      showToast(`DocOpen v${updatePendingVersion} is ready - restart to install`);
      break;
    case "current":
      updateText.textContent = "Up to date";
      updateIdleTimer = window.setTimeout(setUpdateIdle, 4000);
      break;
    case "error":
      updateText.textContent = "Update check failed";
      // The reason is too long for the foot; keep it on hover.
      updateText.title = status.message || "";
      updateButton.textContent = "Retry";
      break;
    case "dev":
      updateText.textContent = "Dev build - updates off";
      updateButton.hidden = true;
      break;
    default:
      updateText.textContent = updateIdleLabel;
  }
}

updateButton?.addEventListener("click", async () => {
  if (updateState === "ready") {
    if (tabs.some((tab) => tab.isDirty)) {
      const go = await showConfirm({
        title: "Restart to install?",
        message: "You have unsaved changes. They will be lost when DocOpen restarts.",
        tone: "warning",
        primaryLabel: "Restart anyway"
      });
      if (!go) {
        return;
      }
    }
    window.documentOpener.installUpdate?.();
    return;
  }
  window.documentOpener.checkForUpdates?.();
});

window.documentOpener.onUpdateStatus?.(applyUpdateStatus);
window.documentOpener.appVersion?.().then((version) => {
  updateIdleLabel = `DocOpen v${version}`;
  if (updateState === "idle") {
    setUpdateIdle();
  }
});
setUpdateIdle();
if (settings.checkUpdatesOnLaunch) {
  window.documentOpener.checkForUpdates?.();
}
