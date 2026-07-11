const path = require("node:path");
const fs = require("node:fs/promises");
const { app, BrowserWindow, clipboard, dialog, ipcMain, nativeImage, shell } = require("electron");
const { autoUpdater } = require("electron-updater");
const { openDocument, SUPPORTED_EXTENSIONS } = require("./documentReader");
const { saveDocument } = require("./documentWriter");

let mainWindow;
let pendingFilePath = findSupportedFileArg(process.argv);
let documentDirty = false;
let allowClose = false;

const gotSingleInstanceLock = app.requestSingleInstanceLock();

if (!gotSingleInstanceLock) {
  app.quit();
}

app.on("second-instance", (_event, argv) => {
  const filePath = findSupportedFileArg(argv);
  if (!filePath) {
    return;
  }

  if (mainWindow) {
    if (mainWindow.isMinimized()) {
      mainWindow.restore();
    }
    mainWindow.focus();
    openPathInWindow(filePath);
  } else {
    pendingFilePath = filePath;
  }
});

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1180,
    height: 780,
    minWidth: 900,
    minHeight: 620,
    backgroundColor: "#101114",
    title: "DocOpen",
    icon: path.join(__dirname, "assets", "icon.ico"),
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false
    }
  });

  mainWindow.loadFile(path.join(__dirname, "index.html"));

  mainWindow.on("close", (event) => {
    if (allowClose || !documentDirty) {
      return;
    }

    // Let the renderer handle it with the in-app styled dialog, then close via
    // the "document:closeConfirmed" message once the user has decided.
    event.preventDefault();
    mainWindow.webContents.send("document:requestClose");
  });

  mainWindow.webContents.once("did-finish-load", () => {
    if (pendingFilePath) {
      openPathInWindow(pendingFilePath);
      pendingFilePath = null;
    }
  });
}

app.whenReady().then(() => {
  createWindow();

  if (app.isPackaged) {
    // Downloads in the background, notifies the user, installs on quit.
    autoUpdater.checkForUpdatesAndNotify().catch(() => {});
  }

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

ipcMain.on("document:dirty", (_event, isDirty) => {
  documentDirty = Boolean(isDirty);
});

ipcMain.on("document:closeConfirmed", () => {
  allowClose = true;
  if (mainWindow) {
    mainWindow.close();
  }
});

ipcMain.handle("document:open", async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: "Open a document",
    properties: ["openFile"],
    filters: [
      { name: "Documents", extensions: ["docx", "xlsx", "csv", "pdf", "md", "txt", "log", "json", "png", "jpg", "jpeg", "gif", "webp", "bmp", "svg"] },
      { name: "Word", extensions: ["docx"] },
      { name: "Excel", extensions: ["xlsx", "csv"] },
      { name: "Markdown", extensions: ["md"] },
      { name: "Text", extensions: ["txt", "log", "json"] },
      { name: "Images", extensions: ["png", "jpg", "jpeg", "gif", "webp", "bmp", "svg"] },
      { name: "PDF", extensions: ["pdf"] }
    ]
  });

  if (result.canceled || result.filePaths.length === 0) {
    return { canceled: true };
  }

  return openVerifiedPath(result.filePaths[0]);
});

ipcMain.handle("document:openPath", async (_event, filePath) => {
  return openVerifiedPath(filePath);
});

ipcMain.handle("document:save", async (_event, payload) => {
  const validation = validateSavePayload(payload);
  if (!validation.ok) {
    return validation;
  }

  return saveDocument(validation.payload);
});

ipcMain.handle("document:saveAs", async (_event, payload) => {
  const validation = validateSavePayload(payload, { allowMissingPath: true });
  if (!validation.ok) {
    return validation;
  }

  const safePayload = validation.payload;
  const defaultExtension = safePayload.kind === "word" ? "docx" : safePayload.extension?.replace(".", "") || "xlsx";
  const result = await dialog.showSaveDialog(mainWindow, {
    title: "Save document",
    defaultPath: safePayload.filePath || safePayload.defaultFileName || `Untitled.${defaultExtension}`,
    filters: getSaveFilters(safePayload.kind, defaultExtension)
  });

  if (result.canceled || !result.filePath) {
    return { canceled: true };
  }

  const savePayload = {
    ...safePayload,
    filePath: result.filePath,
    extension: path.extname(result.filePath).toLowerCase() || `.${defaultExtension}`
  };
  const saveValidation = validateSavePayload(savePayload);
  if (!saveValidation.ok) {
    return saveValidation;
  }

  return saveDocument(saveValidation.payload);
});

// Crash-recovery snapshot: renderer writes dirty tabs periodically, reads it
// back on launch, and a normal quit wipes it — so it only survives a crash.
function autosaveFilePath() {
  return path.join(app.getPath("userData"), "autosave.json");
}

ipcMain.handle("autosave:write", async (_event, data) => {
  try {
    await fs.writeFile(autosaveFilePath(), JSON.stringify(data));
    return { ok: true };
  } catch (error) {
    return { ok: false, error: error.message };
  }
});

ipcMain.handle("autosave:read", async () => {
  try {
    return JSON.parse(await fs.readFile(autosaveFilePath(), "utf8"));
  } catch {
    return null;
  }
});

ipcMain.handle("autosave:clear", async () => {
  await fs.unlink(autosaveFilePath()).catch(() => {});
  return { ok: true };
});

app.on("before-quit", () => {
  try {
    require("node:fs").unlinkSync(autosaveFilePath());
  } catch {}
});

ipcMain.handle("clipboard:writeImage", (_event, dataUrl) => {
  if (typeof dataUrl !== "string" || !dataUrl.startsWith("data:image/")) {
    return { ok: false };
  }
  clipboard.writeImage(nativeImage.createFromDataURL(dataUrl));
  return { ok: true };
});

ipcMain.handle("clipboard:readImage", () => {
  const image = clipboard.readImage();
  return image.isEmpty() ? null : image.toDataURL();
});

ipcMain.handle("document:revealPath", async (_event, filePath) => {
  const validation = await validateRevealPath(filePath);
  if (!validation.ok) {
    return validation;
  }

  shell.showItemInFolder(validation.filePath);
  return { ok: true };
});

function getSaveFilters(kind, defaultExtension) {
  if (kind === "word") {
    return [
      { name: "Word document", extensions: ["docx"] },
      { name: "Markdown", extensions: ["md"] },
      { name: "Text file", extensions: ["txt"] }
    ];
  }

  if (kind === "markdown") {
    return [
      { name: "Markdown", extensions: ["md"] },
      { name: "Word document", extensions: ["docx"] },
      { name: "Text file", extensions: ["txt"] }
    ];
  }

  if (kind === "text") {
    return [
      { name: "Current format", extensions: [defaultExtension] },
      { name: "Text file", extensions: ["txt"] },
      { name: "Log file", extensions: ["log"] },
      { name: "JSON", extensions: ["json"] },
      { name: "Markdown", extensions: ["md"] }
    ];
  }

  if (kind === "image") {
    return [
      { name: "PNG image", extensions: ["png"] },
      { name: "JPEG image", extensions: ["jpg", "jpeg"] }
    ];
  }

  return [
    { name: "Excel workbook", extensions: ["xlsx"] },
    { name: "CSV", extensions: ["csv"] },
    { name: "Current format", extensions: [defaultExtension] }
  ];
}

async function openPathInWindow(filePath) {
  const result = await openVerifiedPath(filePath);
  mainWindow.webContents.send("document:openedFromSystem", result);
}

function findSupportedFileArg(argv) {
  return argv.find((arg) => SUPPORTED_EXTENSIONS.has(path.extname(arg).toLowerCase()));
}

async function openVerifiedPath(filePath) {
  const validation = await validateOpenPath(filePath);
  if (!validation.ok) {
    return validation;
  }

  return openDocument(validation.filePath);
}

async function validateOpenPath(filePath) {
  if (typeof filePath !== "string" || filePath.trim() === "") {
    return { ok: false, error: "No document path was provided." };
  }

  const resolvedPath = path.resolve(filePath);
  const extension = path.extname(resolvedPath).toLowerCase();
  if (!SUPPORTED_EXTENSIONS.has(extension)) {
    return { ok: false, error: "Unsupported file type. Open a document, text, PDF, or image file." };
  }

  try {
    const stats = await fs.stat(resolvedPath);
    if (!stats.isFile()) {
      return { ok: false, error: "Choose a document file, not a folder." };
    }
  } catch {
    return { ok: false, error: "The document could not be found." };
  }

  return { ok: true, filePath: resolvedPath };
}

async function validateRevealPath(filePath) {
  if (typeof filePath !== "string" || filePath.trim() === "") {
    return { ok: false, error: "No path was provided." };
  }

  const resolvedPath = path.resolve(filePath);
  try {
    await fs.access(resolvedPath);
  } catch {
    return { ok: false, error: "The file could not be found." };
  }

  return { ok: true, filePath: resolvedPath };
}

function validateSavePayload(payload, options = {}) {
  if (!payload || typeof payload !== "object") {
    return { ok: false, error: "The save request was invalid." };
  }

  if (payload.kind === "pdf") {
    return { ok: false, error: "This file type is read-only in this version." };
  }

  if (!["word", "workbook", "markdown", "text", "image"].includes(payload.kind)) {
    return { ok: false, error: "Unsupported document type." };
  }

  const requestedExtension = path.extname(payload.filePath || "").toLowerCase() || payload.extension || "";
  const extension = String(requestedExtension).toLowerCase();
  const allowedExtensions = {
    word: [".docx", ".md", ".txt"],
    workbook: [".xlsx", ".csv"],
    markdown: [".md", ".docx", ".txt"],
    text: [".txt", ".log", ".json", ".md"]
  };
  if (allowedExtensions[payload.kind] && extension && !allowedExtensions[payload.kind].includes(extension)) {
    return {
      ok: false,
      error: `This document can only be saved as ${allowedExtensions[payload.kind].join(", ")} files.`
    };
  }

  if (payload.kind === "image" && extension && ![".png", ".jpg", ".jpeg"].includes(extension)) {
    return { ok: false, error: "Images can only be saved as .png or .jpg files." };
  }

  if (!options.allowMissingPath && (typeof payload.filePath !== "string" || payload.filePath.trim() === "")) {
    return { ok: false, error: "Choose a save location before saving." };
  }

  const safePayload = {
    ...payload,
    filePath: payload.filePath ? path.resolve(payload.filePath) : "",
    extension: extension || defaultExtensionForKind(payload.kind)
  };

  if (payload.kind === "word" || payload.kind === "markdown" || payload.kind === "text") {
    safePayload.html = String(payload.html || "");
    safePayload.text = String(payload.text || "");
    safePayload.markdown = String(payload.markdown || "");
  } else if (payload.kind === "image") {
    const dataUrl = typeof payload.dataUrl === "string" && /^data:image\/(png|jpeg);base64,/.test(payload.dataUrl) ? payload.dataUrl : "";
    const sourcePath = typeof payload.sourcePath === "string" && payload.sourcePath.trim() !== "" ? path.resolve(payload.sourcePath) : "";
    if (!dataUrl && !sourcePath) {
      return { ok: false, error: "There is no image data to save." };
    }
    safePayload.dataUrl = dataUrl;
    safePayload.sourcePath = sourcePath;
  } else {
    safePayload.sheets = Array.isArray(payload.sheets) ? payload.sheets : [];
  }

  return { ok: true, payload: safePayload };
}

function defaultExtensionForKind(kind) {
  if (kind === "word") {
    return ".docx";
  }
  if (kind === "markdown") {
    return ".md";
  }
  if (kind === "text") {
    return ".txt";
  }
  if (kind === "image") {
    return ".png";
  }
  return ".xlsx";
}
