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
