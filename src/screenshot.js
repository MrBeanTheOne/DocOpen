const path = require("node:path");
const fs = require("node:fs/promises");
const { app, BrowserWindow, desktopCapturer, ipcMain, nativeImage, screen } = require("electron");

// Screenshot engine: capture-first, crop-second (same trick Snipping Tool uses —
// the screen is frozen at the moment of invocation and the overlay can't
// photobomb its own capture).

async function captureAllDisplays() {
  const displays = screen.getAllDisplays();
  const thumbnailSize = {
    width: Math.max(...displays.map((d) => Math.round(d.size.width * d.scaleFactor))),
    height: Math.max(...displays.map((d) => Math.round(d.size.height * d.scaleFactor)))
  };
  const sources = await desktopCapturer.getSources({ types: ["screen"], thumbnailSize });

  return displays.map((display, index) => {
    const source =
      sources.find((s) => String(s.display_id) === String(display.id)) || sources[index] || sources[0];
    return { display, image: source.thumbnail };
  });
}

// ponytail: offsets assume a uniform scale factor across displays; mixed-DPI
// multi-monitor layouts may show small seams. Map through DIP space if it matters.
function compositeCaptures(captures) {
  const parts = captures.map(({ display, image }) => ({
    x: Math.round(display.bounds.x * display.scaleFactor),
    y: Math.round(display.bounds.y * display.scaleFactor),
    image
  }));
  const minX = Math.min(...parts.map((p) => p.x));
  const minY = Math.min(...parts.map((p) => p.y));
  const width = Math.max(...parts.map((p) => p.x + p.image.getSize().width)) - minX;
  const height = Math.max(...parts.map((p) => p.y + p.image.getSize().height)) - minY;

  const out = Buffer.alloc(width * height * 4);
  for (const part of parts) {
    const size = part.image.getSize();
    const bitmap = part.image.toBitmap();
    const offsetX = part.x - minX;
    const offsetY = part.y - minY;
    for (let row = 0; row < size.height; row += 1) {
      bitmap.copy(out, ((offsetY + row) * width + offsetX) * 4, row * size.width * 4, (row + 1) * size.width * 4);
    }
  }

  return nativeImage.createFromBitmap(out, { width, height });
}

// Fullscreen frozen-image overlays, one per display; the user drags a rectangle
// on whichever screen they want. Resolves with the cropped image, or null on Escape.
async function selectRegion(captures) {
  const tempDir = app.getPath("temp");
  const tempFiles = [];
  const overlays = [];

  const cleanup = () => {
    // Listeners first: destroying an overlay fires "closed" synchronously,
    // which would otherwise emit a cancel that races the real result.
    ipcMain.removeAllListeners("screenshot:regionDone");
    ipcMain.removeAllListeners("screenshot:regionCancel");
    overlays.forEach((w) => !w.isDestroyed() && w.destroy());
    tempFiles.forEach((f) => fs.unlink(f).catch(() => {}));
  };

  for (const { display, image } of captures) {
    const tempFile = path.join(tempDir, `docopen-snip-${display.id}.png`);
    await fs.writeFile(tempFile, image.toPNG());
    tempFiles.push(tempFile);
  }

  return new Promise((resolve) => {
    ipcMain.once("screenshot:regionDone", (_event, rect) => {
      const capture = captures.find((c) => String(c.display.id) === String(rect.displayId));
      cleanup();
      if (!capture || rect.width < 3 || rect.height < 3) {
        resolve(null);
        return;
      }
      const size = capture.image.getSize();
      const x = Math.max(0, Math.min(rect.x, size.width - 1));
      const y = Math.max(0, Math.min(rect.y, size.height - 1));
      resolve(
        capture.image.crop({
          x,
          y,
          width: Math.min(rect.width, size.width - x),
          height: Math.min(rect.height, size.height - y)
        })
      );
    });
    ipcMain.once("screenshot:regionCancel", () => {
      cleanup();
      resolve(null);
    });

    captures.forEach(({ display, image }, index) => {
      const overlay = new BrowserWindow({
        x: display.bounds.x,
        y: display.bounds.y,
        width: display.bounds.width,
        height: display.bounds.height,
        frame: false,
        resizable: false,
        movable: false,
        skipTaskbar: true,
        alwaysOnTop: true,
        enableLargerThanScreen: true,
        backgroundColor: "#000000",
        webPreferences: {
          preload: path.join(__dirname, "overlay-preload.js"),
          contextIsolation: true,
          nodeIntegration: false
        }
      });
      overlay.setAlwaysOnTop(true, "screen-saver");
      overlay.on("closed", () => {
        // User closed an overlay some other way (Alt+F4): treat as cancel.
        if (overlays.every((w) => w.isDestroyed())) {
          ipcMain.emit("screenshot:regionCancel");
        }
      });
      overlay.loadFile(path.join(__dirname, "overlay.html"), {
        query: {
          img: `file:///${tempFiles[index].replace(/\\/g, "/")}`,
          displayId: String(display.id),
          pixelWidth: String(image.getSize().width),
          pixelHeight: String(image.getSize().height)
        }
      });
      overlays.push(overlay);
    });

    const cursorDisplay = screen.getDisplayNearestPoint(screen.getCursorScreenPoint());
    const focused = overlays[captures.findIndex((c) => c.display.id === cursorDisplay.id)];
    (focused || overlays[0])?.focus();
  });
}

// mode: "region" | "fullscreen" | "all". Returns a nativeImage or null (cancelled).
async function captureScreenshot(mode, delayMs = 0) {
  if (delayMs > 0) {
    await new Promise((r) => setTimeout(r, delayMs));
  }

  const captures = await captureAllDisplays();

  if (mode === "all") {
    return captures.length === 1 ? captures[0].image : compositeCaptures(captures);
  }

  if (mode === "fullscreen") {
    const active = screen.getDisplayNearestPoint(screen.getCursorScreenPoint());
    return (captures.find((c) => c.display.id === active.id) || captures[0]).image;
  }

  return selectRegion(captures);
}

module.exports = { captureScreenshot };
