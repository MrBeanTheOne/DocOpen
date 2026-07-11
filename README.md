# DocOpen

A small dark-mode Electron app for opening modern Word, Excel, CSV, and PDF files.

## Run

```powershell
npm install
npm start
```

## Verify

```powershell
npm run smoke
```

The smoke check creates temporary `.docx`, `.xlsx`, `.csv`, and `.pdf` files. It confirms editable formats can open, edit, save, and reopen, and that PDF opens read-only.

## Supported Now

- Word: `.docx`
- Excel: `.xlsx`
- CSV: `.csv`
- PDF: `.pdf` read-only preview

Legacy `.doc` and `.xls` files are intentionally not included in this first slice.

## Editing

- Toggle Edit before changing document content; the edit toolbar appears only in edit mode.
- Word editing supports text, headings, bold, italic, underline, colors, alignment, lists, undo/redo, find & replace (Ctrl+F, matches across formatting), hyperlinks, inline images, and tables, then saves into a clean `.docx`.
- Inside a table, use the row/column tool buttons to add or delete rows and columns; deleting the last row or column removes the table.
- Hyperlinks save with a real Word `Hyperlink` style, and external image URLs are fetched and embedded on save.
- Excel and CSV editing supports direct cell edits, add/delete row, add/delete column, and clear cell.
- Excel grids support drag resizing column headers and row headers; `.xlsx` saves those dimensions.
- PDF files are opened in read-only mode only.
- Save creates a timestamped `.bak` copy beside the original file before overwriting.
- Revert restores the last saved/opened version in the current session.
- The left sidebar can collapse to keep the document grid front and center.

## File Association Ready

The app is wired to open files passed by Windows when using Open with or a packaged file association. The package metadata includes `.docx`, `.xlsx`, `.csv`, and `.pdf` associations for a later installer/build step.
