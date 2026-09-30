// Image viewing + MS Paint-style canvas editor: tools, selection marquee,
// clipboard, transforms, flood fill, text boxes, undo/redo, context menu.
// Classic script; owns the paint state globals and shares the rest with
// renderer.js (viewer, currentDocument, setDirty, ...).

let imageCanvas = null;
let paintTool = "brush";
let paintShapeFill = false;
let paintSelection = null; // canvas-space rect { x, y, width, height }
let paintSelectionBox = null; // marquee overlay element (lives in the paint shell)
let paintTextEditor = null; // in-place text box: { element, x, y, size, color }
let paintUndoStack = [];
let paintRedoStack = [];
let paintZoomReset = null;

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

  if (canSaveCurrentDocument()) {
    const saveItem = document.createElement("button");
    saveItem.type = "button";
    saveItem.className = "new-menu-item";
    saveItem.style.gridTemplateColumns = "1fr";
    saveItem.style.padding = "0 12px";
    saveItem.textContent = "Save as…";
    saveItem.addEventListener("click", async () => {
      closeImageContextMenu();
      await saveCurrentDocument(true);
    });
    menu.appendChild(saveItem);
  }

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
