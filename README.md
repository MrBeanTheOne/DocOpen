<div align="center">
  <img src="src/assets/logo.png" alt="DocOpen" width="96" height="96">

  # DocOpen

  **A dead-simple, dark-mode document editor for Windows.**

  Word, Excel, CSV, PDF, Markdown, JSON, plain text, and images — one fast Electron app, no cloud, no account.

</div>

---

## Features

- **Tabs** — open several documents at once; `Ctrl+Tab` / `Ctrl+Shift+Tab` to cycle, `Ctrl+W` to close.
- **Word documents** — rich text editing (bold/italic/underline, colors, alignment, lists, links, inline images, tables) that saves back to a clean `.docx`.
- **Spreadsheets** — cell editing, row/column add & delete, resizable headers, and live formulas (`SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `+ - * /`) for `.xlsx` and `.csv`.
- **Images** — a built-in paint editor: brush, eraser, eyedropper, text, shapes, flood fill, flip/rotate, resize, a select/cut/crop/move tool, undo/redo, and OS-clipboard copy-paste.
- **Markdown, JSON & text** — live preview, and a raw-view toggle to see (or hand-edit) the exact source, with one-click JSON pretty-printing.
- **PDF** — quick read-only preview.
- **Find & replace** — one `Ctrl+F` bar that works across every document kind, including inside spreadsheet formulas.
- **Screenshot tool** — capture a region, a window, or the whole screen (`Ctrl+Alt+S`, or from the tray), with an optional delay. Shots land on the clipboard and open as an editable image tab.
- **Runs in the tray** — closing the window frees the memory-heavy renderer instead of just hiding it; the app keeps living in the tray for quick reopen, recent files, and screenshots.
- **Autosave & crash recovery** — dirty tabs are snapshotted every 30 seconds and offered back if the app ever closes uncleanly.
- **Safe saves** — every save drops a timestamped `.bak` next to the original first; **Revert** restores the last saved/opened state in the session.
- **Auto-updates** — new versions ship as GitHub releases and install themselves via `electron-updater`.

## Supported formats

| Format | Extension(s) | Editing |
|---|---|---|
| Word | `.docx` | Full rich-text editor |
| Excel | `.xlsx` | Cells, rows/columns, formulas |
| CSV | `.csv` | Cells, rows/columns |
| Markdown | `.md` | Text editor + rendered preview |
| Text / logs / JSON | `.txt`, `.log`, `.json` | Text editor (JSON gets pretty-print) |
| Images | `.png`, `.jpg`/`.jpeg` | Full paint editor |
| Images (other) | `.gif`, `.webp`, `.bmp`, `.svg` | View only |
| PDF | `.pdf` | Read-only preview |

Legacy `.doc` and `.xls` are intentionally out of scope.

## Install & run

```powershell
npm install
npm start
```

Windows file associations (Open with → DocOpen) are wired up in `package.json` for all supported extensions once the app is installed via the packaged build below.

## Building an installer

```powershell
npm run dist:win     # NSIS installer in dist/, unsigned local build
```

Tagged pushes (`v*`) run `.github/workflows/release.yml`, which builds the installer and publishes it as a GitHub release; installed copies pick up new releases automatically on launch.

## Keyboard shortcuts

| Shortcut | Action |
|---|---|
| `Ctrl+O` | Open a file |
| `Ctrl+N` | New document |
| `Ctrl+S` | Save |
| `Ctrl+W` | Close tab |
| `Ctrl+Tab` / `Ctrl+Shift+Tab` | Next / previous tab |
| `Ctrl+F` | Find & replace |
| `Ctrl+Alt+S` | Screenshot: region |
| `Ctrl+Z` / `Ctrl+Shift+Z` or `Ctrl+Y` | Undo / redo (image edit) |
| `Ctrl+C` / `Ctrl+X` / `Ctrl+V` | Copy / cut / paste (image selection & spreadsheet ranges) |

## Verify

```powershell
npm run smoke       # headless: create/open/edit/save/reopen for every format, incl. PDF read-only
npm run ui-check    # drives the real app through its UI interactions
```

`scripts/ui-shots.js` (`npx electron scripts/ui-shots.js`) drives the app through key visual states and writes PNGs to `scripts/shots/` for manual visual QA — useful after any styling change.
