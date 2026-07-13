const { contextBridge, ipcRenderer, webUtils } = require("electron");

contextBridge.exposeInMainWorld("documentOpener", {
  openDocument: () => ipcRenderer.invoke("document:open"),
  openRecentFile: (filePath) => ipcRenderer.invoke("document:openPath", filePath),
  openDroppedFile: (file) => {
    const filePath = webUtils.getPathForFile(file);
    return ipcRenderer.invoke("document:openPath", filePath);
  },
  saveDocument: (payload) => ipcRenderer.invoke("document:save", payload),
  saveDocumentAs: (payload) => ipcRenderer.invoke("document:saveAs", payload),
  copyImageToClipboard: (dataUrl) => ipcRenderer.invoke("clipboard:writeImage", dataUrl),
  copyImageFileToClipboard: (filePath) => ipcRenderer.invoke("clipboard:writeImageFromPath", filePath),
  readImageFromClipboard: () => ipcRenderer.invoke("clipboard:readImage"),
  revealPath: (filePath) => ipcRenderer.invoke("document:revealPath", filePath),
  onOpenedFromSystem: (callback) => {
    const listener = (_event, result) => callback(result);
    ipcRenderer.on("document:openedFromSystem", listener);
    return () => ipcRenderer.removeListener("document:openedFromSystem", listener);
  },
  notifyDirty: (isDirty) => ipcRenderer.send("document:dirty", isDirty),
  writeAutosave: (data) => ipcRenderer.invoke("autosave:write", data),
  readAutosave: () => ipcRenderer.invoke("autosave:read"),
  clearAutosave: () => ipcRenderer.invoke("autosave:clear"),
  onRequestClose: (callback) => {
    const listener = () => callback();
    ipcRenderer.on("document:requestClose", listener);
    return () => ipcRenderer.removeListener("document:requestClose", listener);
  },
  confirmClose: () => ipcRenderer.send("document:closeConfirmed"),
  cancelClose: () => ipcRenderer.send("document:closeCancelled"),
  captureScreenshot: (mode, delayMs) => ipcRenderer.invoke("screenshot:capture", mode, delayMs),
  onScreenshot: (callback) => {
    const listener = (_event, payload) => callback(payload);
    ipcRenderer.on("screenshot:new", listener);
    return () => ipcRenderer.removeListener("screenshot:new", listener);
  },
  onNewDocumentRequest: (callback) => {
    const listener = (_event, kind) => callback(kind);
    ipcRenderer.on("document:newRequest", listener);
    return () => ipcRenderer.removeListener("document:newRequest", listener);
  },
  syncRecentFiles: (files) => ipcRenderer.send("recent:sync", files)
});
