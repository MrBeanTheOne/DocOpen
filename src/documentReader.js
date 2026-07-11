const path = require("node:path");
const fs = require("node:fs/promises");
const { pathToFileURL } = require("node:url");
const ExcelJS = require("exceljs");
const mammoth = require("mammoth");

const TEXT_EXTENSIONS = new Set([".txt", ".log", ".json"]);
const IMAGE_EXTENSIONS = new Set([".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".svg"]);
const SUPPORTED_EXTENSIONS = new Set([".docx", ".xlsx", ".csv", ".pdf", ".md", ...TEXT_EXTENSIONS, ...IMAGE_EXTENSIONS]);

async function openDocument(filePath) {
  const extension = path.extname(filePath).toLowerCase();

  if (!SUPPORTED_EXTENSIONS.has(extension)) {
    return {
      ok: false,
      error: "Unsupported file type. Open a document, text, PDF, or image file."
    };
  }

  try {
    if (extension === ".docx") {
      return await openWordDocument(filePath);
    }

    if (extension === ".pdf") {
      return await openPdfDocument(filePath);
    }

    if (extension === ".md") {
      return await openMarkdownDocument(filePath);
    }

    if (TEXT_EXTENSIONS.has(extension)) {
      return await openTextDocument(filePath, extension);
    }

    if (IMAGE_EXTENSIONS.has(extension)) {
      return await openImageDocument(filePath, extension);
    }

    return await openWorkbook(filePath, extension);
  } catch (error) {
    return {
      ok: false,
      error: error instanceof Error ? error.message : "The document could not be opened."
    };
  }
}

async function openPdfDocument(filePath) {
  const metadata = await fileMetadata(filePath);

  return {
    ok: true,
    kind: "pdf",
    filePath,
    fileName: path.basename(filePath),
    extension: ".pdf",
    fileUrl: pathToFileURL(filePath).href,
    ...metadata
  };
}

async function openWordDocument(filePath) {
  const metadata = await fileMetadata(filePath);
  const result = await mammoth.convertToHtml({
    path: filePath
  });

  return {
    ok: true,
    kind: "word",
    filePath,
    fileName: path.basename(filePath),
    extension: ".docx",
    ...metadata,
    html: result.value,
    warnings: result.messages.map((message) => message.message)
  };
}

async function openMarkdownDocument(filePath) {
  const metadata = await fileMetadata(filePath);
  const text = await fs.readFile(filePath, "utf8");

  return {
    ok: true,
    kind: "markdown",
    filePath,
    fileName: path.basename(filePath),
    extension: ".md",
    ...metadata,
    text
  };
}

async function openTextDocument(filePath, extension) {
  const metadata = await fileMetadata(filePath);
  const text = await fs.readFile(filePath, "utf8");

  return {
    ok: true,
    kind: "text",
    filePath,
    fileName: path.basename(filePath),
    extension,
    ...metadata,
    text
  };
}

async function openImageDocument(filePath, extension) {
  const metadata = await fileMetadata(filePath);

  return {
    ok: true,
    kind: "image",
    filePath,
    fileName: path.basename(filePath),
    extension,
    fileUrl: pathToFileURL(filePath).href,
    ...metadata
  };
}

async function openWorkbook(filePath, extension) {
  const workbook = new ExcelJS.Workbook();

  if (extension === ".csv") {
    const metadata = await fileMetadata(filePath);
    const worksheet = await workbook.csv.readFile(filePath);

    return {
      ok: true,
      kind: "workbook",
      filePath,
      fileName: path.basename(filePath),
      extension,
      ...metadata,
      sheets: [worksheetToPreview(worksheet)]
    };
  }

  const metadata = await fileMetadata(filePath);
  await workbook.xlsx.readFile(filePath);

  return {
    ok: true,
    kind: "workbook",
    filePath,
    fileName: path.basename(filePath),
    extension,
    ...metadata,
    sheets: workbook.worksheets.map((worksheet) => worksheetToPreview(worksheet))
  };
}

async function fileMetadata(filePath) {
  const stats = await fs.stat(filePath);
  return {
    size: stats.size,
    modifiedAt: stats.mtime.toISOString()
  };
}

function worksheetToPreview(worksheet) {
  const rows = [];
  const rowHeights = [];

  worksheet.eachRow({ includeEmpty: true }, (row) => {
    const values = [];
    rowHeights[row.number - 1] = pointsToPixels(row.height);

    row.eachCell({ includeEmpty: true }, (cell, columnNumber) => {
      values[columnNumber - 1] = formatCellValue(cell.value);
    });

    rows[row.number - 1] = values;
  });

  return {
    name: worksheet.name || "Sheet 1",
    rows: trimTrailingEmptyRows(rows.map((row) => row || [])),
    columnWidths: worksheet.columns.map((column) => excelWidthToPixels(column.width)),
    rowHeights
  };
}

function excelWidthToPixels(width) {
  if (!width) {
    return null;
  }

  return Math.max(48, Math.round(width * 7 + 5));
}

function pointsToPixels(height) {
  if (!height) {
    return null;
  }

  return Math.max(22, Math.round(height / 0.75));
}

function formatCellValue(value) {
  if (value == null) {
    return "";
  }

  if (value instanceof Date) {
    return value.toLocaleDateString();
  }

  if (typeof value === "object") {
    if ("text" in value) {
      return value.text;
    }

    // Formulas the in-app evaluator understands stay live as "=..." text;
    // anything fancier keeps showing Excel's cached result.
    if ("formula" in value && value.formula && isSupportedFormula(value.formula)) {
      return `=${value.formula}`;
    }

    if ("result" in value) {
      return formatCellValue(value.result);
    }

    if (Array.isArray(value.richText)) {
      return value.richText.map((item) => item.text).join("");
    }

    if ("hyperlink" in value && "text" in value) {
      return value.text;
    }

    return JSON.stringify(value);
  }

  return String(value);
}

// Mirrors what src/formula.js can evaluate: refs, arithmetic, and the five
// supported functions. Keep the two in sync when the evaluator grows.
function isSupportedFormula(formula) {
  const text = String(formula);
  if (!/^[\sA-Za-z0-9$:+\-*/(),.]*$/.test(text)) {
    return false;
  }

  const words = text.match(/[A-Za-z][A-Za-z0-9]*/g) || [];
  return words.every(
    (word) => /^[A-Za-z]{1,3}\d+$/.test(word) || /^(SUM|AVERAGE|MIN|MAX|COUNT)$/i.test(word)
  );
}

function trimTrailingEmptyRows(rows) {
  let lastContentRow = rows.length - 1;

  while (lastContentRow >= 0) {
    const hasContent = rows[lastContentRow].some((cell) => String(cell).trim() !== "");
    if (hasContent) {
      break;
    }
    lastContentRow -= 1;
  }

  return rows.slice(0, lastContentRow + 1);
}

module.exports = {
  openDocument,
  SUPPORTED_EXTENSIONS
};
