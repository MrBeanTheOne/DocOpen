// Screenshot harness: drives the real app through key states and captures PNGs.
// Run: npx electron <this file>  (cwd = repo root)
const path = require("path");
const fs = require("fs");
const { app, BrowserWindow } = require("electron");
require(path.join(process.cwd(), "src", "main.js"));

const OUT = path.join(__dirname, "shots");
fs.mkdirSync(OUT, { recursive: true });

const STATES = [
  ["01-empty", `(async () => { await window.documentOpener.clearAutosave(); if (activeDialog) resolveAppDialog("cancel"); })()`],
  ["02-new-menu", `(async () => { document.querySelector("#newButton").click(); })()`],
  ["03-word-edit", `(async () => {
    document.querySelector("#newMenu").hidden = true;
    await createNewDocument("word");
    wordEditor.innerHTML = "<h1>Quarterly Report</h1><p>Alpha <strong>bravo</strong> charlie <em>delta</em> echo foxtrot golf hotel.</p><ul><li>One</li><li>Two</li></ul>";
  })()`],
  ["04-workbook", `(async () => {
    await createNewDocument("workbook");
    const t = (r,c,v) => { const cell = document.querySelector('td[data-row-index="'+r+'"][data-column-index="'+c+'"]'); cell.dispatchEvent(new PointerEvent("pointerdown",{bubbles:true,button:0})); document.dispatchEvent(new PointerEvent("pointerup",{bubbles:true})); document.execCommand("insertText", false, v); };
    t(0,0,"Item"); t(0,1,"Qty"); t(1,0,"Widget"); t(1,1,"42"); t(2,0,"Gadget"); t(2,1,"7");
    t(3,1,"=B2+B3");
    const blur = document.querySelector('td[data-row-index="5"][data-column-index="0"]');
    blur.dispatchEvent(new PointerEvent("pointerdown",{bubbles:true,button:0}));
  })()`],
  ["05-markdown-view", `(async () => {
    await createNewDocument("markdown");
    document.querySelector(".markdown-editor").value = "# DocOpen\\n\\nA **markdown** preview with a [link](https://example.com).\\n\\n- alpha\\n- bravo\\n\\n> quoted line\\n\\n\\u0060\\u0060\\u0060\\ncode block\\n\\u0060\\u0060\\u0060";
    document.querySelector(".markdown-editor").dispatchEvent(new Event("input",{bubbles:true}));
    setEditMode(false);
  })()`],
  ["06-image-paint", `(async () => {
    await createNewDocument("image", { width: 640, height: 400 });
    await new Promise(r => setTimeout(r, 80));
    const c = document.querySelector(".paint-canvas").getContext("2d");
    c.fillStyle = "#3b82f6"; c.fillRect(60, 60, 180, 120);
    c.strokeStyle = "#ef4444"; c.lineWidth = 6; c.beginPath(); c.arc(400, 200, 80, 0, 7); c.stroke();
  })()`],
  ["07-find-bar", `(async () => {
    activateTab(0);
    toggleFindBar(true);
  })()`],
  ["08-dialog", `(async () => {
    toggleFindBar(false);
    showAlert({ title: "Delete backup?", message: "This removes the backup copy from disk. The original document is not touched.", tone: "danger", primaryLabel: "Delete", cancelLabel: "Cancel" });
  })()`],
  ["09-json-view", `(async () => {
    if (activeDialog) resolveAppDialog("cancel");
    await createNewDocument("text");
    currentDocument.extension = ".json";
    currentDocument.fileName = "config.json";
    document.querySelector(".markdown-editor").value = JSON.stringify({name:"docopen",count:3,ok:true,tags:["a","b"]});
    document.querySelector(".markdown-editor").dispatchEvent(new Event("input",{bubbles:true}));
    setEditMode(false);
  })()`],
  ["10-sidebar-collapsed", `(async () => { document.querySelector("#sidebarToggle").click(); })()`],
  ["11-sidebar-restored", `(async () => { document.querySelector("#sidebarRail").click(); })()`]
];

app.whenReady().then(async () => {
  let win;
  for (let attempt = 0; attempt < 50 && !win; attempt += 1) {
    win = BrowserWindow.getAllWindows()[0];
    if (!win) await new Promise((r) => setTimeout(r, 100));
  }
  if (!win) { console.error("no window"); app.exit(1); return; }
  if (win.webContents.isLoading()) {
    await new Promise((r) => win.webContents.once("did-finish-load", r));
  }
  win.setSize(1440, 900);
  await new Promise((r) => setTimeout(r, 600));

  for (const [name, script] of STATES) {
    try {
      await win.webContents.executeJavaScript(script, true);
    } catch (e) {
      console.error(name, "script error:", e.message);
    }
    await new Promise((r) => setTimeout(r, 250));
    const image = await win.webContents.capturePage();
    fs.writeFileSync(path.join(OUT, name + ".png"), image.toPNG());
    console.log("shot:", name);
  }
  await win.webContents.executeJavaScript(`window.documentOpener.clearAutosave()`, true).catch(() => {});
  app.exit(0);
});
