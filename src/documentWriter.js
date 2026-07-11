const fs = require("node:fs/promises");
const path = require("node:path");
const ExcelJS = require("exceljs");
const JSZip = require("jszip");

async function saveDocument(payload) {
  if (!payload || !payload.filePath) {
    return {
      ok: false,
      error: "Choose a save location before saving."
    };
  }

  try {
    const backupPath = await createBackup(payload.filePath);
    const tempPath = await createTempSavePath(payload.filePath, payload.extension);

    try {
      const targetExtension = (path.extname(payload.filePath) || payload.extension || "").toLowerCase();
      const isDocumentKind = payload.kind === "word" || payload.kind === "markdown" || payload.kind === "text";

      if (payload.kind === "workbook") {
        await saveWorkbook({
          ...payload,
          filePath: tempPath
        });
      } else if (payload.kind === "image") {
        await saveImageDocument(tempPath, payload);
      } else if (isDocumentKind && targetExtension === ".docx") {
        // Word natively, or a Markdown doc exported through the Word writer.
        await saveWordDocument(tempPath, payload.html || "", payload.text || "");
      } else if (isDocumentKind && targetExtension === ".md" && payload.kind === "word") {
        await saveMarkdownDocument(tempPath, payload.markdown || payload.text || "");
      } else if (isDocumentKind) {
        await saveMarkdownDocument(tempPath, payload.text || "");
      } else {
        return {
          ok: false,
          error: "Unsupported document type."
        };
      }

      await replaceWithTempFile(tempPath, payload.filePath);
    } finally {
      await removeTempFile(tempPath);
    }

    const stats = await fs.stat(payload.filePath);

    return {
      ok: true,
      filePath: payload.filePath,
      fileName: path.basename(payload.filePath),
      size: stats.size,
      modifiedAt: stats.mtime.toISOString(),
      backupPath,
      saveMode: "temp-replace"
    };
  } catch (error) {
    return {
      ok: false,
      error: error instanceof Error ? error.message : "The document could not be saved."
    };
  }
}

async function createBackup(filePath) {
  try {
    await fs.access(filePath);
  } catch {
    return null;
  }

  const parsed = path.parse(filePath);
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const backupPath = path.join(parsed.dir, `${parsed.name}.${stamp}.bak${parsed.ext}`);
  await fs.copyFile(filePath, backupPath);
  return backupPath;
}

async function createTempSavePath(filePath, extension) {
  const parsed = path.parse(filePath);
  const suffix = path.extname(filePath) || extension || ".tmp";

  for (let attempt = 0; attempt < 20; attempt += 1) {
    const token = `${Date.now()}-${process.pid}-${Math.random().toString(36).slice(2, 8)}`;
    const tempPath = path.join(parsed.dir, `.docopen-${parsed.name}-${token}.tmp${suffix}`);
    try {
      const handle = await fs.open(tempPath, "wx");
      await handle.close();
      return tempPath;
    } catch (error) {
      if (!error || error.code !== "EEXIST") {
        throw error;
      }
    }
  }

  throw new Error("Could not create a temporary save file.");
}

async function replaceWithTempFile(tempPath, targetPath) {
  await fs.copyFile(tempPath, targetPath);
}

async function removeTempFile(tempPath) {
  try {
    await fs.rm(tempPath, { force: true });
  } catch {
    // A leftover temp file is less important than preserving the user's save result.
  }
}

async function saveImageDocument(filePath, payload) {
  // ponytail: Save As with no edits copies the original bytes as-is, even if
  // the chosen extension differs; convert via nativeImage if that ever bites.
  if (!payload.dataUrl) {
    if (!payload.sourcePath) {
      throw new Error("There is no image data to save.");
    }
    await fs.copyFile(payload.sourcePath, filePath);
    return;
  }

  const wantsJpeg = payload.extension === ".jpg" || payload.extension === ".jpeg";
  const match = /^data:image\/(png|jpeg);base64,(.+)$/.exec(payload.dataUrl);
  if (match && (match[1] === "jpeg") === wantsJpeg) {
    await fs.writeFile(filePath, Buffer.from(match[2], "base64"));
    return;
  }

  // Canvas encoded one format but the user picked the other in Save As.
  const { nativeImage } = require("electron");
  const image = nativeImage.createFromDataURL(payload.dataUrl);
  if (image.isEmpty()) {
    throw new Error("The image data could not be read.");
  }
  await fs.writeFile(filePath, wantsJpeg ? image.toJPEG(92) : image.toPNG());
}

async function saveWorkbook(payload) {
  const workbook = new ExcelJS.Workbook();

  const isCsv = payload.extension === ".csv";
  payload.sheets.forEach((sheet, index) => {
    const worksheet = workbook.addWorksheet(sheet.name || `Sheet ${index + 1}`);
    sheet.rows.forEach((row) => {
      worksheet.addRow(isCsv ? row : row.map(workbookCellValue));
    });
    applySheetLayout(worksheet, sheet);
  });

  if (payload.extension === ".csv") {
    await workbook.csv.writeFile(payload.filePath, {
      sheetName: payload.sheets[0]?.name || "Sheet 1"
    });
    return;
  }

  await workbook.xlsx.writeFile(payload.filePath);
}

async function saveMarkdownDocument(filePath, text) {
  await fs.writeFile(filePath, String(text || ""), "utf8");
}

// Grid cells are edited as strings; store "=..." as live formulas and numeric
// strings as numbers so Excel treats saved cells natively. CSV keeps raw text.
function workbookCellValue(value) {
  if (typeof value !== "string") {
    return value;
  }

  const trimmed = value.trim();
  if (trimmed.startsWith("=") && trimmed.length > 1) {
    return { formula: trimmed.slice(1) };
  }
  if (trimmed !== "" && !Number.isNaN(Number(trimmed))) {
    return Number(trimmed);
  }
  return value;
}

function applySheetLayout(worksheet, sheet) {
  if (Array.isArray(sheet.columnWidths)) {
    sheet.columnWidths.forEach((width, index) => {
      if (width) {
        worksheet.getColumn(index + 1).width = pixelsToExcelWidth(width);
      }
    });
  }

  if (Array.isArray(sheet.rowHeights)) {
    sheet.rowHeights.forEach((height, index) => {
      if (height) {
        worksheet.getRow(index + 1).height = pixelsToPoints(height);
      }
    });
  }
}

function pixelsToExcelWidth(width) {
  return Math.max(4, Math.round(((Number(width) - 5) / 7) * 100) / 100);
}

function pixelsToPoints(height) {
  return Math.max(12, Math.round(Number(height) * 0.75 * 100) / 100);
}

async function saveWordDocument(filePath, html, text) {
  const zip = new JSZip();
  const inlinedHtml = html ? await inlineRemoteImages(html) : "";
  const blocks = inlinedHtml ? htmlToWordBlocks(inlinedHtml) : textToWordBlocks(text);
  const ctx = { rels: createRelRegistry(), media: [], imageExts: new Set() };
  const bodyXml =
    blocks.length === 0
      ? renderParagraph({ segments: [{ text: "" }] }, ctx)
      : blocks.map((block) => renderBlock(block, ctx)).join("");

  // Word finds styles.xml via a "styles" relationship, not a body reference.
  ctx.rels.add("styles", "styles.xml");

  zip.file(
    "[Content_Types].xml",
    xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        ${[...ctx.imageExts]
          .map((ext) => `<Default Extension="${ext}" ContentType="${imageContentType(ext)}"/>`)
          .join("")}
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
      </Types>`)
  );
  zip.folder("_rels").file(
    ".rels",
    xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
          Target="word/document.xml"/>
      </Relationships>`)
  );

  const wordFolder = zip.folder("word");
  wordFolder.file(
    "document.xml",
    xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
          ${bodyXml}
          <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
          </w:sectPr>
        </w:body>
      </w:document>`)
  );
  wordFolder.file(
    "styles.xml",
    xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:type="character" w:styleId="Hyperlink">
          <w:name w:val="Hyperlink"/>
          <w:rPr>
            <w:color w:val="0563C1"/>
            <w:u w:val="single"/>
          </w:rPr>
        </w:style>
      </w:styles>`)
  );
  wordFolder.folder("_rels").file("document.xml.rels", ctx.rels.toXml());
  ctx.media.forEach((item) => {
    wordFolder.folder("media").file(item.name, item.buffer);
  });

  const buffer = await zip.generateAsync({ type: "nodebuffer" });
  await fs.writeFile(filePath, buffer);
}

function textToWordBlocks(text) {
  return String(text || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => ({
      segments: [{ text: line }]
    }));
}

function htmlToWordBlocks(html) {
  const source = String(html);
  const blocks = [];
  const tablePattern = /<table[\s\S]*?<\/table>/gi;
  let lastIndex = 0;
  let match;

  while ((match = tablePattern.exec(source))) {
    blocks.push(...paragraphsFromHtml(source.slice(lastIndex, match.index)));
    blocks.push(parseTableBlock(match[0]));
    lastIndex = tablePattern.lastIndex;
  }

  blocks.push(...paragraphsFromHtml(source.slice(lastIndex)));

  if (blocks.length > 0) {
    return blocks;
  }

  const fallbackText = stripHtml(source).trim();
  return fallbackText ? textToWordBlocks(fallbackText) : [];
}

function paragraphsFromHtml(html) {
  const blocks = [];
  const normalized = String(html)
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|h1|h2|h3|li)>/gi, "</$1>\n");
  const blockPattern = /<(h1|h2|h3|p|div|li)([^>]*)>([\s\S]*?)<\/\1>/gi;
  let match;

  while ((match = blockPattern.exec(normalized))) {
    const tagName = match[1].toLowerCase();
    const blockAttributes = parseTagAttributes(match[2] || "");
    const segments = parseInlineSegments(match[3]);
    const hasContent = segments.some((segment) => segment.image || (segment.text && segment.text.trim() !== ""));

    if (hasContent) {
      blocks.push({
        tagName,
        alignment: alignmentFromAttributes(blockAttributes),
        segments: tagName === "li" ? [{ text: "- " }, ...segments] : segments
      });
    }
  }

  if (blocks.length > 0) {
    return blocks;
  }

  const fallbackText = stripHtml(normalized).trim();
  return fallbackText ? textToWordBlocks(fallbackText) : [];
}

function parseTableBlock(tableHtml) {
  const rows = [];
  const rowPattern = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  let rowMatch;

  while ((rowMatch = rowPattern.exec(tableHtml))) {
    const cells = [];
    const cellPattern = /<(t[dh])[^>]*>([\s\S]*?)<\/\1>/gi;
    let cellMatch;

    while ((cellMatch = cellPattern.exec(rowMatch[1]))) {
      cells.push({ segments: parseInlineSegments(cellMatch[2]) });
    }

    if (cells.length > 0) {
      rows.push(cells);
    }
  }

  return { type: "table", rows };
}

function parseInlineSegments(html) {
  const segments = [];
  const state = {
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    color: null,
    backgroundColor: null,
    fontFamily: null,
    fontSize: null,
    link: null
  };
  const tokenPattern = /<\/?[^>]+>|[^<]+/g;
  let match;

  while ((match = tokenPattern.exec(html))) {
    const token = match[0];

    if (token.startsWith("<")) {
      const tagName = token.match(/^<\s*([a-z0-9]+)/i)?.[1]?.toLowerCase();
      if (tagName === "img" && !token.startsWith("</")) {
        const attributes = parseTagAttributes(token);
        if (isSafeImageSource(attributes.src)) {
          segments.push({ image: { dataUrl: attributes.src }, link: state.link });
        }
        continue;
      }
      updateInlineState(token, state);
      continue;
    }

    const text = decodeHtml(stripHtml(token));
    if (text) {
      segments.push({
        text,
        bold: state.bold,
        italic: state.italic,
        underline: state.underline,
        strike: state.strike,
        color: state.color,
        backgroundColor: state.backgroundColor,
        fontFamily: state.fontFamily,
        fontSize: state.fontSize,
        link: state.link
      });
    }
  }

  return segments.length > 0 ? segments : [{ text: decodeHtml(stripHtml(html)) }];
}

function updateInlineState(token, state) {
  const closing = token.startsWith("</");
  const tagMatch = token.match(/^<\/?\s*([a-z0-9]+)/i);
  const tagName = tagMatch?.[1]?.toLowerCase();
  const attributes = parseTagAttributes(token);

  if (tagName === "strong" || tagName === "b") {
    state.bold = !closing;
  } else if (tagName === "em" || tagName === "i") {
    state.italic = !closing;
  } else if (tagName === "u") {
    state.underline = !closing;
  } else if (tagName === "s" || tagName === "strike" || tagName === "del") {
    state.strike = !closing;
  } else if (tagName === "font") {
    state.color = closing ? null : normalizeColor(attributes.color) || state.color;
    state.fontFamily = closing ? null : cleanFontFamily(attributes.face) || state.fontFamily;
    state.fontSize = closing ? null : fontSizeValueToHalfPoints(attributes.size) || state.fontSize;
  } else if (tagName === "a") {
    state.link = closing ? null : safeLink(attributes.href);
  } else if (tagName === "span") {
    const styles = parseStyleAttribute(attributes.style);
    state.color = closing ? null : normalizeColor(styles.color) || state.color;
    state.backgroundColor = closing ? null : normalizeColor(styles["background-color"]) || state.backgroundColor;
    state.fontFamily = closing ? null : cleanFontFamily(styles["font-family"]) || state.fontFamily;
    state.fontSize = closing ? null : cssFontSizeToHalfPoints(styles["font-size"]) || state.fontSize;
  }
}

function renderBlock(block, ctx) {
  if (block.type === "table") {
    return renderTable(block, ctx);
  }

  return renderParagraph(block, ctx);
}

function renderParagraph(block, ctx) {
  const paragraphProperties = paragraphPropertiesFor(block);
  const inner = renderInline(block.segments || [], block.tagName, ctx);
  return `<w:p>${paragraphProperties}${inner}</w:p>`;
}

function renderInline(segments, tagName, ctx) {
  let out = "";
  let index = 0;

  while (index < segments.length) {
    const link = segments[index].link;

    if (link) {
      let runs = "";
      while (index < segments.length && segments[index].link === link) {
        runs += renderSegment(segments[index], tagName, ctx);
        index += 1;
      }
      const relId = ctx.rels.add("hyperlink", link, "External");
      out += `<w:hyperlink r:id="${relId}">${runs}</w:hyperlink>`;
    } else {
      out += renderSegment(segments[index], tagName, ctx);
      index += 1;
    }
  }

  return out;
}

function renderSegment(segment, tagName, ctx) {
  if (segment.image) {
    return renderImage(segment.image, ctx);
  }

  if (segment.link) {
    return wordRun({ ...segment, styleId: "Hyperlink", underline: true, color: segment.color || "0563C1" }, tagName);
  }

  return wordRun(segment, tagName);
}

function renderTable(block, ctx) {
  const rows = block.rows || [];
  const columnCount = Math.max(1, ...rows.map((row) => row.length));
  const gridWidth = Math.floor(9360 / columnCount);
  const grid = Array.from({ length: columnCount }, () => `<w:gridCol w:w="${gridWidth}"/>`).join("");
  const borders = ["top", "left", "bottom", "right", "insideH", "insideV"]
    .map((edge) => `<w:${edge} w:val="single" w:sz="4" w:space="0" w:color="808080"/>`)
    .join("");
  const tableProperties = `<w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>${borders}</w:tblBorders></w:tblPr>`;

  const body = rows
    .map((row) => {
      let cells = "";
      for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
        const cell = row[columnIndex];
        const segments = cell && cell.segments && cell.segments.length > 0 ? cell.segments : [{ text: "" }];
        const paragraph = renderParagraph({ segments }, ctx);
        cells += `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>${paragraph}</w:tc>`;
      }
      return `<w:tr>${cells}</w:tr>`;
    })
    .join("");

  return `<w:tbl>${tableProperties}<w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`;
}

function renderImage(image, ctx) {
  const { buffer, ext } = decodeDataUrl(image.dataUrl);
  if (!buffer) {
    return "";
  }

  const { width, height } = imageDimensions(buffer, ext);
  const { cx, cy } = fitImageEmu(width, height);
  const index = ctx.media.length + 1;
  const name = `image${index}.${ext}`;
  ctx.media.push({ name, buffer });
  ctx.imageExts.add(ext);
  const relId = ctx.rels.add("image", `media/${name}`);

  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${index}" name="Image ${index}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${index}" name="Image ${index}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

function paragraphPropertiesFor(block) {
  const properties = [];

  if (block.tagName === "h1") {
    properties.push('<w:spacing w:before="240" w:after="120"/>');
  }

  if (block.tagName === "h2" || block.tagName === "h3") {
    properties.push('<w:spacing w:before="180" w:after="90"/>');
  }

  if (block.alignment) {
    properties.push(`<w:jc w:val="${block.alignment}"/>`);
  }

  return properties.length > 0 ? `<w:pPr>${properties.join("")}</w:pPr>` : "";
}

function wordRun(segment, tagName) {
  const properties = [];

  if (segment.styleId) {
    properties.push(`<w:rStyle w:val="${segment.styleId}"/>`);
  }

  if (segment.bold || tagName === "h1" || tagName === "h2" || tagName === "h3") {
    properties.push("<w:b/>");
  }

  if (segment.italic) {
    properties.push("<w:i/>");
  }

  if (segment.underline) {
    properties.push('<w:u w:val="single"/>');
  }

  if (segment.strike) {
    properties.push("<w:strike/>");
  }

  if (segment.fontFamily) {
    const family = escapeXml(segment.fontFamily);
    properties.push(`<w:rFonts w:ascii="${family}" w:hAnsi="${family}"/>`);
  }

  if (segment.color) {
    properties.push(`<w:color w:val="${segment.color}"/>`);
  }

  if (segment.backgroundColor) {
    properties.push(`<w:shd w:val="clear" w:color="auto" w:fill="${segment.backgroundColor}"/>`);
  }

  if (segment.fontSize) {
    properties.push(`<w:sz w:val="${segment.fontSize}"/>`);
  }

  if (tagName === "h1") {
    properties.push('<w:sz w:val="32"/>');
  } else if (tagName === "h2") {
    properties.push('<w:sz w:val="28"/>');
  }

  const runProperties = properties.length > 0 ? `<w:rPr>${properties.join("")}</w:rPr>` : "";
  return `<w:r>${runProperties}<w:t xml:space="preserve">${escapeXml(segment.text)}</w:t></w:r>`;
}

function parseTagAttributes(token) {
  const attributes = {};
  const attributePattern = /([a-zA-Z:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/g;
  let match;

  while ((match = attributePattern.exec(token))) {
    attributes[match[1].toLowerCase()] = match[2] ?? match[3] ?? match[4] ?? "";
  }

  return attributes;
}

function parseStyleAttribute(style) {
  const styles = {};
  String(style || "")
    .split(";")
    .forEach((declaration) => {
      const [rawName, ...rawValue] = declaration.split(":");
      const name = rawName?.trim().toLowerCase();
      const value = rawValue.join(":").trim();
      if (name && value) {
        styles[name] = value;
      }
    });
  return styles;
}

function alignmentFromAttributes(attributes) {
  const styles = parseStyleAttribute(attributes.style);
  const alignment = String(attributes.align || styles["text-align"] || "").toLowerCase();

  if (alignment === "center") {
    return "center";
  }

  if (alignment === "right") {
    return "right";
  }

  if (alignment === "justify") {
    return "both";
  }

  return null;
}

function normalizeColor(value) {
  const color = String(value || "").trim();
  if (!color) {
    return null;
  }

  if (color.startsWith("#")) {
    const hex = color.slice(1);
    if (/^[0-9a-f]{3}$/i.test(hex)) {
      return hex
        .split("")
        .map((char) => `${char}${char}`)
        .join("")
        .toUpperCase();
    }
    return /^[0-9a-f]{6}$/i.test(hex) ? hex.toUpperCase() : null;
  }

  const rgbMatch = color.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
  if (rgbMatch) {
    return [rgbMatch[1], rgbMatch[2], rgbMatch[3]]
      .map((part) => clampColor(Number(part)).toString(16).padStart(2, "0"))
      .join("")
      .toUpperCase();
  }

  const namedColors = {
    black: "000000",
    blue: "0000FF",
    cyan: "00FFFF",
    green: "008000",
    grey: "808080",
    gray: "808080",
    magenta: "FF00FF",
    orange: "FFA500",
    purple: "800080",
    red: "FF0000",
    white: "FFFFFF",
    yellow: "FFFF00"
  };

  return namedColors[color.toLowerCase()] || null;
}

function clampColor(value) {
  return Math.max(0, Math.min(255, Math.round(value)));
}

function cleanFontFamily(value) {
  return String(value || "")
    .split(",")[0]
    .replace(/['"]/g, "")
    .trim() || null;
}

function fontSizeValueToHalfPoints(value) {
  const sizeMap = {
    1: 16,
    2: 20,
    3: 24,
    4: 28,
    5: 36,
    6: 48,
    7: 72
  };
  const numericValue = Number(value);
  return sizeMap[numericValue] || null;
}

function cssFontSizeToHalfPoints(value) {
  const size = String(value || "").trim().toLowerCase();
  if (!size) {
    return null;
  }

  const numericValue = Number.parseFloat(size);
  if (!Number.isFinite(numericValue)) {
    return null;
  }

  if (size.endsWith("px")) {
    return Math.max(12, Math.round(numericValue * 1.5));
  }

  if (size.endsWith("pt")) {
    return Math.max(12, Math.round(numericValue * 2));
  }

  return fontSizeValueToHalfPoints(numericValue);
}

function stripHtml(value) {
  return String(value).replace(/<[^>]*>/g, "");
}

function decodeHtml(value) {
  return String(value)
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&apos;/g, "'");
}

function xml(value) {
  return value.replace(/\n\s+/g, " ").trim();
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function createRelRegistry() {
  const items = [];
  const typeUrls = {
    hyperlink: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
  };

  return {
    add(type, target, mode) {
      const id = `rId${items.length + 1}`;
      items.push({ id, type, target, mode });
      return id;
    },
    toXml() {
      return xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          ${items
            .map(
              (item) =>
                `<Relationship Id="${item.id}" Type="${typeUrls[item.type]}" Target="${escapeXml(item.target)}"${
                  item.mode ? ` TargetMode="${item.mode}"` : ""
                }/>`
            )
            .join("")}
        </Relationships>`);
    }
  };
}

// Fetch any http(s) <img> sources and inline them as data URLs so they embed
// into the .docx like pasted images. Unreachable images are left as-is (dropped).
async function inlineRemoteImages(html) {
  const urls = [...String(html).matchAll(/<img[^>]+src="(https?:\/\/[^"]+)"/gi)].map((match) => match[1]);
  let result = String(html);

  for (const url of [...new Set(urls)]) {
    try {
      const response = await fetch(url);
      if (!response.ok) {
        continue;
      }
      const buffer = Buffer.from(await response.arrayBuffer());
      const mime = response.headers.get("content-type") || "image/png";
      result = result.split(url).join(`data:${mime};base64,${buffer.toString("base64")}`);
    } catch {
      // Leave the original URL in place; renderImage will drop it safely.
    }
  }

  return result;
}

function safeLink(value) {
  const link = String(value || "").trim();
  if (!link) {
    return null;
  }

  try {
    const parsed = new URL(link);
    return ["http:", "https:", "mailto:"].includes(parsed.protocol) ? link : null;
  } catch {
    return null;
  }
}

function isSafeImageSource(value) {
  return /^data:image\/(?:png|jpe?g|gif);base64,/i.test(String(value || "").trim());
}

function imageContentType(ext) {
  if (ext === "jpg" || ext === "jpeg") {
    return "image/jpeg";
  }
  if (ext === "gif") {
    return "image/gif";
  }
  return "image/png";
}

function decodeDataUrl(dataUrl) {
  const match = /^data:([^;,]+)?(;base64)?,(.*)$/is.exec(String(dataUrl || ""));
  if (!match || !match[2]) {
    return { buffer: null, ext: "png" };
  }

  const mime = (match[1] || "image/png").toLowerCase();
  const ext = mime.includes("jpeg") || mime.includes("jpg")
    ? "jpg"
    : mime.includes("gif")
      ? "gif"
      : mime.includes("png")
        ? "png"
        : null;

  if (!ext) {
    return { buffer: null, ext: "png" };
  }

  try {
    return { buffer: Buffer.from(match[3], "base64"), ext };
  } catch {
    return { buffer: null, ext };
  }
}

function fitImageEmu(width, height) {
  const perPixel = 9525;
  const maxWidth = 5943600; // 6.5in content width in EMU
  let cx = Math.round(width * perPixel);
  let cy = Math.round(height * perPixel);

  if (cx > maxWidth) {
    cy = Math.round((cy * maxWidth) / cx);
    cx = maxWidth;
  }

  return { cx: Math.max(1, cx), cy: Math.max(1, cy) };
}

function imageDimensions(buffer, ext) {
  try {
    if (ext === "png" && buffer.length >= 24 && buffer.toString("ascii", 1, 4) === "PNG") {
      return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
    }

    if (ext === "gif" && buffer.length >= 10 && buffer.toString("ascii", 0, 3) === "GIF") {
      return { width: buffer.readUInt16LE(6), height: buffer.readUInt16LE(8) };
    }

    if ((ext === "jpg" || ext === "jpeg") && buffer[0] === 0xff && buffer[1] === 0xd8) {
      let offset = 2;
      while (offset + 9 < buffer.length) {
        if (buffer[offset] !== 0xff) {
          break;
        }
        const marker = buffer[offset + 1];
        if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
          return { height: buffer.readUInt16BE(offset + 5), width: buffer.readUInt16BE(offset + 7) };
        }
        offset += 2 + buffer.readUInt16BE(offset + 2);
      }
    }
  } catch {
    // Fall through to the default size below.
  }

  return { width: 400, height: 300 };
}

module.exports = {
  saveDocument
};
