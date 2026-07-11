// Drives the real renderer inside Electron to check workbook DOM wiring
// (cell editing, formula display swap, caret placement). Run: npm run ui-check
const { app, BrowserWindow } = require("electron");
require("../src/main.js");

const TEST = `(async () => {
  const out = { failures: [] };
  const check = (condition, message) => { if (!condition) out.failures.push(message); };

  const cell = (r, c) => document.querySelector('td[data-row-index="' + r + '"][data-column-index="' + c + '"]');
  const click = (r, c) => {
    cell(r, c).dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0 }));
    document.dispatchEvent(new PointerEvent("pointerup", { bubbles: true }));
  };
  const type = (r, c, text) => {
    click(r, c);
    document.execCommand("insertText", false, text);
  };
  const pressKey = (r, c, key, shiftKey = false) => {
    cell(r, c).dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key, shiftKey }));
  };
  const caretInfo = (r, c) => {
    const selection = getSelection();
    if (!selection.rangeCount) return { inCell: false, atEnd: false };
    const range = selection.getRangeAt(0);
    const el = cell(r, c);
    const end = document.createRange();
    end.selectNodeContents(el);
    end.collapse(false);
    return {
      inCell: el.contains(range.startContainer),
      atEnd: range.collapsed && range.compareBoundaryPoints(Range.START_TO_START, end) === 0
    };
  };

  // A leftover crash-recovery file (e.g. from real app use) would pop a restore
  // dialog and desync the dialog clicks below — clear it before starting.
  await window.documentOpener.clearAutosave();
  if (activeDialog) resolveAppDialog("cancel");

  await createNewDocument("workbook");

  // 1. Clean click-type flow ending with a formula.
  type(0, 0, "5");
  type(0, 1, "3");
  type(0, 2, "=A1+B1");
  click(1, 0); // blur the formula cell
  out.formulaDisplay = cell(0, 2).textContent;
  out.formulaRaw = cell(0, 2).dataset.formula || null;
  out.modelRow0 = JSON.parse(JSON.stringify(currentDocument.sheets[0].rows[0] || []));
  check(out.formulaDisplay === "8", "formula display expected 8, got " + JSON.stringify(out.formulaDisplay));
  check(out.formulaRaw === "=A1+B1", "formula raw expected =A1+B1, got " + JSON.stringify(out.formulaRaw));

  // 2. Clicking into a cell with text puts the caret at the end.
  click(0, 0);
  const caret = caretInfo(0, 0);
  check(caret.inCell && caret.atEnd, "caret should be at end of clicked cell, got " + JSON.stringify(caret));

  // 3. Clicking into a formula cell swaps to the raw formula, caret at end.
  click(0, 2);
  out.focusedFormulaText = cell(0, 2).textContent;
  const formulaCaret = caretInfo(0, 2);
  check(out.focusedFormulaText === "=A1+B1", "focused formula cell should show raw formula, got " + JSON.stringify(out.focusedFormulaText));
  check(formulaCaret.inCell && formulaCaret.atEnd, "caret should be at end of formula text, got " + JSON.stringify(formulaCaret));
  click(1, 1); // blur again -> back to computed
  check(cell(0, 2).textContent === "8", "formula cell should show computed value after blur, got " + JSON.stringify(cell(0, 2).textContent));

  // 4. Enter-key navigation flow: type, Enter, type, Enter, formula.
  type(2, 0, "10");
  pressKey(2, 0, "Enter");
  type(3, 0, "20");
  pressKey(3, 0, "Enter");
  type(4, 0, "=SUM(A3:A4)");
  pressKey(4, 0, "Enter");
  check(cell(4, 0).textContent === "30", "Enter-flow formula expected 30, got " + JSON.stringify(cell(4, 0).textContent));
  check(document.activeElement === cell(5, 0), "Enter should move focus down a row");

  // 5. Editing a referenced cell recomputes dependents on blur.
  type(0, 0, "0"); // caret at end of "5" -> "50"
  click(1, 0);
  check(cell(0, 0).textContent === "50", "typing at caret should append, got " + JSON.stringify(cell(0, 0).textContent));
  check(cell(0, 2).textContent === "53", "dependent formula should recompute to 53, got " + JSON.stringify(cell(0, 2).textContent));

  // 5. Tabs isolate documents and switching preserves state.
  await createNewDocument("text");
  check(tabs.length === 2, "expected 2 tabs after creating a second document, got " + tabs.length);
  check(currentDocument.kind === "text", "active tab should be the new text document");
  const textEditor = document.querySelector(".markdown-editor");
  textEditor.value = "tab two text";
  textEditor.dispatchEvent(new Event("input", { bubbles: true }));
  activateTab(0);
  check(currentDocument.kind === "workbook", "switching back should restore the workbook tab");
  check(cell(0, 0).textContent === "50", "workbook cell should survive the tab switch, got " + JSON.stringify(cell(0, 0).textContent));
  check(document.querySelectorAll(".doc-tab").length === 2, "tab strip should show 2 tabs");
  activateTab(1);
  check(document.querySelector(".markdown-editor").value === "tab two text", "text content should survive the tab switch");
  const closePromise = closeTab(1);
  await new Promise((resolve) => setTimeout(resolve, 80));
  document.querySelector("#dialogPrimaryButton").click();
  await closePromise;
  check(tabs.length === 1, "closing a tab should remove it, got " + tabs.length);
  check(currentDocument.kind === "workbook", "closing the active tab should activate its neighbor");

  // 6. htmlToMarkdown basics for Save As .md exports.
  const md = htmlToMarkdown("<h1>Title</h1><p><strong>bold</strong> and <em>it</em></p><ul><li>one</li><li>two</li></ul>");
  check(
    md.includes("# Title") && md.includes("**bold**") && md.includes("*it*") && md.includes("- one"),
    "htmlToMarkdown should convert headings/emphasis/lists, got " + JSON.stringify(md)
  );

  // 7. Paint select / cut / crop on a fresh image canvas.
  await createNewDocument("image", { width: 100, height: 80 });
  await new Promise((resolve) => setTimeout(resolve, 50)); // let the canvas get layout
  const paint = document.querySelector(".paint-canvas");
  const paintCtx = paint.getContext("2d");
  paintCtx.fillStyle = "#ff0000";
  paintCtx.fillRect(12, 12, 6, 6);
  const paintAt = (x, y) => {
    const rect = paint.getBoundingClientRect();
    return { clientX: rect.left + (x / paint.width) * rect.width, clientY: rect.top + (y / paint.height) * rect.height };
  };
  const paintPointer = (type, x, y) =>
    paint.dispatchEvent(new PointerEvent(type, { bubbles: true, button: 0, ...paintAt(x, y) }));
  const dragSelect = () => {
    paintPointer("pointerdown", 10, 10);
    paintPointer("pointermove", 50, 40);
    paintPointer("pointerup", 50, 40);
  };
  const near = (a, b) => Math.abs(a - b) <= 2; // px rounding through the rendered rect

  document.querySelector('[data-image-tool="select"]').click();
  dragSelect();
  check(
    paintSelection && near(paintSelection.x, 10) && near(paintSelection.y, 10) && near(paintSelection.width, 40) && near(paintSelection.height, 30),
    "select drag should set a ~40x30 selection at ~10,10, got " + JSON.stringify(paintSelection)
  );
  check(!document.querySelector(".paint-selection").hidden, "selection marquee should be visible after a drag");

  document.querySelector('[data-image-action="cut-selection"]').click();
  check(paintCtx.getImageData(25, 25, 1, 1).data[3] === 0, "cut should clear the selected area to transparent");
  check(paintCtx.getImageData(5, 5, 1, 1).data[3] === 255, "cut should leave pixels outside the selection alone");
  check(paintSelection === null, "cut should clear the selection");
  undoPaint();
  check(paintCtx.getImageData(25, 25, 1, 1).data[3] === 255, "undo should restore the cut pixels");

  dragSelect();
  document.querySelector('[data-image-action="crop-selection"]').click();
  check(near(paint.width, 40) && near(paint.height, 30), "crop should resize the canvas to the selection, got " + paint.width + "x" + paint.height);
  check(near(currentDocument.width, 40) && near(currentDocument.height, 30), "crop should sync the document size");
  const croppedPixel = paint.getContext("2d").getImageData(4, 4, 1, 1).data;
  check(croppedPixel[0] === 255 && croppedPixel[1] === 0 && croppedPixel[2] === 0, "crop should keep content relative to the selection origin");
  undoPaint();
  check(paint.width === 100 && paint.height === 80, "undo should restore the pre-crop canvas size, got " + paint.width + "x" + paint.height);

  // 8. Dragging inside the selection moves the lifted pixels.
  dragSelect(); // ~ (10,10) 40x30, red block at (12,12)-(18,18)
  paintPointer("pointerdown", 20, 20);
  paintPointer("pointermove", 45, 45);
  paintPointer("pointerup", 45, 45);
  check(
    paintSelection && near(paintSelection.x, 35) && near(paintSelection.y, 35),
    "move drag should shift the selection by the drag delta, got " + JSON.stringify(paintSelection)
  );
  const movedPixel = paintCtx.getImageData(39, 39, 1, 1).data;
  check(movedPixel[0] === 255 && movedPixel[1] === 0 && movedPixel[2] === 0, "moved selection should carry its pixels");
  check(paintCtx.getImageData(14, 14, 1, 1).data[3] === 0, "moving should leave a transparent hole at the source");

  // 9. Copy / paste round-trip via the real OS clipboard (clobbers it; dev-check only).
  await copyPaintSelection();
  await pastePaintClipboard();
  check(paintSelection && paintSelection.x === 0 && paintSelection.y === 0, "paste should select the pasted block at the origin");
  check(paintTool === "select", "paste should hand back the select tool");
  const pastedPixel = paintCtx.getImageData(4, 4, 1, 1).data;
  check(pastedPixel[0] === 255 && pastedPixel[1] === 0 && pastedPixel[2] === 0, "pasted block should contain the copied pixels");
  undoPaint();
  check(paintCtx.getImageData(4, 4, 1, 1).data[1] === 255, "undo should restore the white pixel under the pasted block");

  // 9b. In-place text tool: overlay box, commit on click-away, Escape cancels.
  document.querySelector('[data-image-tool="text"]').click();
  paintPointer("pointerdown", 10, 10);
  let textBox = document.querySelector(".paint-text-editor");
  check(Boolean(textBox), "clicking with the text tool should open an overlay text box");
  if (textBox) {
    textBox.innerText = "Hi";
    paintPointer("pointerdown", 90, 70); // click elsewhere commits
    check(!document.querySelector(".paint-text-editor"), "clicking away should close the text box");
    const region = paintCtx.getImageData(10, 10, 30, 24).data;
    let inked = false;
    for (let i = 0; i < region.length; i += 4) {
      if (region[i + 3] > 0 && (region[i] !== 255 || region[i + 1] !== 255 || region[i + 2] !== 255)) {
        inked = true;
        break;
      }
    }
    check(inked, "committed text should leave non-white pixels near the click point");
    undoPaint();
  }
  paintPointer("pointerdown", 10, 10);
  textBox = document.querySelector(".paint-text-editor");
  if (textBox) {
    textBox.innerText = "discard me";
    textBox.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));
    check(!document.querySelector(".paint-text-editor"), "Escape should cancel the text box");
  }

  // 9c. Resizable text box: a narrowed box wraps the committed text.
  paintPointer("pointerdown", 10, 10);
  textBox = document.querySelector(".paint-text-editor");
  if (textBox) {
    check(getComputedStyle(textBox).resize === "both", "text box should have a native resize grip");
    textBox.style.width = "50px"; // what the grip drag would set
    textBox.innerText = "hello world";
    paintPointer("pointerdown", 90, 70); // commit
    const wrapRegion = paintCtx.getImageData(10, 31, 20, 10).data;
    let wrapInk = false;
    for (let i = 0; i < wrapRegion.length; i += 4) {
      if (wrapRegion[i + 3] > 0 && (wrapRegion[i] !== 255 || wrapRegion[i + 1] !== 255 || wrapRegion[i + 2] !== 255)) {
        wrapInk = true;
        break;
      }
    }
    check(wrapInk, "narrowed text box should wrap the second word onto a second line");
    undoPaint();
  }

  // 10. Sidebar: location foot removed, section titles collapse their panels.
  check(!document.querySelector(".sidebar-foot"), "location foot should be gone from the sidebar");
  const recentToggle = document.querySelector("#recentPanel .panel-toggle");
  recentToggle.click();
  check(document.querySelector("#recentPanel").classList.contains("is-collapsed"), "clicking the title should collapse Recently opened");
  check(getComputedStyle(document.querySelector("#recentList")).display === "none", "collapsed recent list should be hidden");
  recentToggle.click();
  check(getComputedStyle(document.querySelector("#recentList")).display !== "none", "clicking the title again should expand the list");
  const infoToggle = document.querySelector("#documentInfoPanel .panel-toggle");
  infoToggle.click();
  check(document.querySelector("#documentInfoPanel").classList.contains("is-collapsed"), "clicking the title should collapse Document info");
  infoToggle.click();

  // 11. Find & Replace across word, textarea, and workbook documents.
  await createNewDocument("word");
  wordEditor.innerHTML = "<p>foo bar foo</p>";
  findInput.value = "bar";
  findInDocument(false);
  check(getSelection().toString() === "bar", "word find should select the match, got " + JSON.stringify(getSelection().toString()));
  findInput.value = "foo";
  replaceInput.value = "baz";
  replaceAllMatches();
  check(/baz bar baz/.test(wordEditor.innerText), "word replace-all should replace both matches, got " + JSON.stringify(wordEditor.innerText));

  await createNewDocument("text");
  const findTextarea = document.querySelector(".markdown-editor");
  findTextarea.value = "alpha beta alpha";
  findTextarea.dispatchEvent(new Event("input", { bubbles: true }));
  findInput.value = "beta";
  findInDocument(false);
  check(findTextarea.selectionStart === 6 && findTextarea.selectionEnd === 10, "textarea find should select the match range, got " + findTextarea.selectionStart + "-" + findTextarea.selectionEnd);
  findInput.value = "alpha";
  replaceInput.value = "gamma";
  replaceAllMatches();
  check(findTextarea.value === "gamma beta gamma", "textarea replace-all should replace both matches, got " + JSON.stringify(findTextarea.value));

  activateTab(0); // back to the workbook
  findInput.value = "20";
  findInDocument(false);
  check(selectedCell && selectedCell.rowIndex === 3 && selectedCell.columnIndex === 0, "workbook find should select the matching cell, got " + JSON.stringify(selectedCell));
  replaceInput.value = "99";
  replaceCurrentMatch();
  check(currentDocument.sheets[0].rows[3][0] === "99", "workbook replace should rewrite the cell, got " + JSON.stringify(currentDocument.sheets[0].rows[3][0]));
  const ctrlF = new KeyboardEvent("keydown", { key: "f", ctrlKey: true, bubbles: true, cancelable: true });
  document.dispatchEvent(ctrlF);
  check(!document.querySelector("#findBar").hidden, "Ctrl+F should open the find bar on a workbook");
  document.querySelector("#findCloseButton").click();

  // 12. Sorting a sheet with formula refs asks for confirmation first.
  check(sheetHasCellRefFormulas({ rows: [["=A1+B1"]] }), "ref formulas should be detected");
  check(!sheetHasCellRefFormulas({ rows: [["=1+1", "plain"]] }), "formulas without refs should not trip the guard");
  click(0, 0);
  const rowsBeforeSort = JSON.stringify(currentDocument.sheets[0].rows);
  const sortPromise = runSheetAction("sort-asc");
  await new Promise((resolve) => setTimeout(resolve, 80));
  check(Boolean(activeDialog), "sorting a sheet with formulas should open the confirm dialog");
  document.querySelector("#dialogCancelButton").click();
  await sortPromise;
  check(JSON.stringify(currentDocument.sheets[0].rows) === rowsBeforeSort, "cancelling the sort confirm should leave rows untouched");

  // 13. Crash-recovery autosave: restore flow, periodic write, cleanup.
  const tabCountBeforeRestore = tabs.length;
  await window.documentOpener.writeAutosave({
    savedAt: "test",
    documents: [{ document: { kind: "text", fileName: "restored.txt", extension: ".txt", text: "restored text" }, editMode: true }]
  });
  const restorePromise = restoreAutosave();
  await new Promise((resolve) => setTimeout(resolve, 80));
  document.querySelector("#dialogPrimaryButton").click();
  await restorePromise;
  check(tabs.length === tabCountBeforeRestore + 1, "restore should add a tab for the recovered document");
  check(currentDocument.text === "restored text", "restore should present the recovered content");
  check(isDirty, "a restored document should be dirty");
  await autosaveTick();
  const autosaved = await window.documentOpener.readAutosave();
  check(Array.isArray(autosaved?.documents) && autosaved.documents.length > 0, "autosaveTick should write the dirty tabs");
  await window.documentOpener.clearAutosave();
  check((await window.documentOpener.readAutosave()) === null, "clearAutosave should remove the snapshot");

  // 14. Paint: corner resize grip, text font picker, fill tolerance.
  await createNewDocument("image", { width: 100, height: 80 });
  await new Promise((resolve) => setTimeout(resolve, 50)); // let the canvas get layout
  const paint2 = document.querySelector(".paint-canvas");
  const grip = document.querySelector(".paint-resize-grip");
  check(Boolean(grip), "paint canvas should have a resize grip");
  if (grip) {
    const gripRect = grip.getBoundingClientRect();
    const gripScale = paint2.clientWidth / paint2.width || 1;
    const startWidth = paint2.width;
    const startHeight = paint2.height;
    const gripPointer = (type, dx, dy) =>
      grip.dispatchEvent(new PointerEvent(type, { bubbles: true, button: 0, clientX: gripRect.left + 6 + dx, clientY: gripRect.top + 6 + dy }));
    gripPointer("pointerdown", 0, 0);
    gripPointer("pointermove", 20 * gripScale, 10 * gripScale);
    gripPointer("pointerup", 20 * gripScale, 10 * gripScale);
    check(
      Math.abs(paint2.width - (startWidth + 20)) <= 2 && Math.abs(paint2.height - (startHeight + 10)) <= 2,
      "grip drag should resize the canvas by the drag delta, got " + paint2.width + "x" + paint2.height + " from " + startWidth + "x" + startHeight
    );
    check(currentDocument.width === paint2.width, "grip resize should sync the document size");
    undoPaint();
    check(paint2.width === startWidth && paint2.height === startHeight, "undo should restore the pre-grip-resize size");
  }

  const paint2Ctx = paint2.getContext("2d");
  document.querySelector("#paintFontSelect").value = "Georgia";
  document.querySelector('[data-image-tool="text"]').click();
  const paintAt2 = (x, y) => {
    const rect = paint2.getBoundingClientRect();
    return { clientX: rect.left + (x / paint2.width) * rect.width, clientY: rect.top + (y / paint2.height) * rect.height };
  };
  paint2.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0, ...paintAt2(10, 10) }));
  const fontBox = document.querySelector(".paint-text-editor");
  check(fontBox && fontBox.style.font.includes("Georgia"), "text box should use the picked font, got " + JSON.stringify(fontBox?.style.font));
  fontBox?.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));
  document.querySelector("#paintFontSelect").value = "Arial";

  // Tolerance: a #f8f8f8 patch on white differs by 7/channel — an exact fill
  // stays inside the patch, a 96 fill bleeds into the white background.
  paint2Ctx.fillStyle = "#ffffff";
  paint2Ctx.fillRect(0, 0, paint2.width, paint2.height);
  paint2Ctx.fillStyle = "#f8f8f8";
  paint2Ctx.fillRect(0, 0, 20, 20);
  document.querySelector("#imageColorInput").value = "#ff0000";
  document.querySelector("#fillToleranceSelect").value = "0";
  document.querySelector('[data-image-tool="fill"]').click();
  paint2.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0, ...paintAt2(10, 10) }));
  check(paint2Ctx.getImageData(10, 10, 1, 1).data[0] === 255 && paint2Ctx.getImageData(10, 10, 1, 1).data[1] === 0, "exact fill should recolor the patch");
  check(paint2Ctx.getImageData(40, 40, 1, 1).data[1] === 255, "exact fill should not bleed into the near-white background");
  undoPaint();
  document.querySelector("#fillToleranceSelect").value = "96";
  paint2.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0, ...paintAt2(10, 10) }));
  check(paint2Ctx.getImageData(40, 40, 1, 1).data[1] === 0, "tolerant fill should bleed across the 7-value difference");
  document.querySelector("#fillToleranceSelect").value = "32";

  // 15. Raw view toggles (markdown/JSON/CSV) + JSON pretty view + Format.
  await createNewDocument("markdown");
  document.querySelector(".markdown-editor").value = "# Hello Raw";
  document.querySelector(".markdown-editor").dispatchEvent(new Event("input", { bubbles: true }));
  setEditMode(false);
  check(Boolean(document.querySelector(".markdown-document")), "markdown view mode should render the preview");
  check(!document.querySelector("#rawViewButton").hidden, "raw view button should show for markdown in view mode");
  document.querySelector("#rawViewButton").click();
  const mdPre = document.querySelector(".text-document");
  check(mdPre && mdPre.textContent === "# Hello Raw", "markdown raw view should show the raw text, got " + JSON.stringify(mdPre?.textContent));
  document.querySelector("#rawViewButton").click();
  check(Boolean(document.querySelector(".markdown-document")), "toggling raw off should restore the preview");

  await createNewDocument("text");
  currentDocument.extension = ".json";
  currentDocument.fileName = "test.json";
  const jsonRaw = '{"name":"docopen","count":3,"ok":true,"nothing":null}';
  document.querySelector(".markdown-editor").value = jsonRaw;
  document.querySelector(".markdown-editor").dispatchEvent(new Event("input", { bubbles: true }));
  setEditMode(false);
  const jsonPre = document.querySelector(".text-document");
  check(Boolean(jsonPre?.classList.contains("json-view")), "JSON view mode should default to the pretty view");
  check(jsonPre?.querySelectorAll(".json-key").length === 4, "pretty JSON should highlight 4 keys, got " + jsonPre?.querySelectorAll(".json-key").length);
  check(Boolean(jsonPre?.textContent.includes('  "name": "docopen"')), "pretty JSON should be indented");
  document.querySelector("#rawViewButton").click();
  check(document.querySelector(".text-document")?.textContent === jsonRaw, "JSON raw view should show the text exactly as typed");
  setEditMode(true);
  check(!document.querySelector("#formatJsonButton").hidden, "Format button should show for JSON in edit mode");
  document.querySelector("#formatJsonButton").click();
  check(
    document.querySelector(".markdown-editor").value === JSON.stringify(JSON.parse(jsonRaw), null, 2),
    "Format should pretty-print the JSON in place, got " + JSON.stringify(document.querySelector(".markdown-editor").value)
  );

  await createNewDocument("csv");
  setEditMode(false);
  currentDocument.sheets[0].rows = [["a", "b,c"], ["1", "2"]];
  document.querySelector("#rawViewButton").click();
  check(
    document.querySelector(".text-document")?.textContent === 'a,"b,c"\\n1,2',
    "CSV raw view should serialize with quoting, got " + JSON.stringify(document.querySelector(".text-document")?.textContent)
  );
  document.querySelector("#rawViewButton").click();
  check(Boolean(cell(0, 0)), "toggling raw off should restore the CSV grid");

  // Leave no crash-recovery bait behind: app.exit() skips before-quit cleanup.
  await window.documentOpener.clearAutosave();

  return JSON.stringify(out, null, 2);
})()`;

app.whenReady().then(async () => {
  let win;
  for (let attempt = 0; attempt < 50 && !win; attempt += 1) {
    win = BrowserWindow.getAllWindows()[0];
    if (!win) {
      await new Promise((resolve) => setTimeout(resolve, 100));
    }
  }
  if (!win) {
    console.error("ui-check: no window appeared");
    app.exit(1);
    return;
  }
  if (win.webContents.isLoading()) {
    await new Promise((resolve) => win.webContents.once("did-finish-load", resolve));
  }

  try {
    const result = await win.webContents.executeJavaScript(TEST, true);
    const parsed = JSON.parse(result);
    console.log(result);
    if (parsed.failures.length > 0) {
      console.error(`ui-check FAILED (${parsed.failures.length} failure${parsed.failures.length === 1 ? "" : "s"})`);
      app.exit(1);
      return;
    }
    console.log("ui-check passed");
    app.exit(0);
  } catch (error) {
    console.error("ui-check errored:", error);
    app.exit(1);
  }
});
