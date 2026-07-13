const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("snip", {
  done: (rect) => ipcRenderer.send("screenshot:regionDone", rect),
  cancel: () => ipcRenderer.send("screenshot:regionCancel")
});
