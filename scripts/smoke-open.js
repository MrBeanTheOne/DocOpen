const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const ExcelJS = require("exceljs");
const JSZip = require("jszip");
const { openDocument } = require("../src/documentReader");
const { saveDocument } = require("../src/documentWriter");
const formula = require("../src/formula");

function checkFormulaEvaluator() {
  const sheet = {
    rows: [
      ["1", "2", "=SUM(A1:B1)"],
      ["=A1*B1+3", "text", "=C1/0"],
      ["=NOPE(1)", "=A2", "=B2"]
    ]
  };
  assert(formula.evaluateCellDisplay(sheet, 0, 2) === "3", "Expected =SUM(A1:B1) to be 3");
  assert(formula.evaluateCellDisplay(sheet, 1, 0) === "5", "Expected =A1*B1+3 to be 5");
  assert(formula.evaluateCellDisplay(sheet, 1, 2) === "#DIV/0!", "Expected division by zero to show #DIV/0!");
  assert(formula.evaluateCellDisplay(sheet, 2, 0) === "#NAME?", "Expected unknown function to show #NAME?");
  assert(formula.evaluateCellDisplay(sheet, 2, 1) === "5", "Expected formula chains to evaluate");
  assert(formula.evaluateCellDisplay(sheet, 2, 2) === "text", "Expected a bare ref to mirror text");
  assert(formula.evaluateCellDisplay({ rows: [["=AVERAGE(A2:A4, 10)"], ["4"], [""], ["x"]] }, 0, 0) === "7", "Expected AVERAGE to skip blanks and text");
  const cyclic = { rows: [["=B1", "=A1"]] };
  assert(formula.evaluateCellDisplay(cyclic, 0, 0) === "#REF!", "Expected a formula cycle to show #REF!");
}

async function main() {
  checkFormulaEvaluator();
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "docopen-"));
  const xlsxPath = path.join(tempDir, "sample.xlsx");
  const csvPath = path.join(tempDir, "sample.csv");
  const docxPath = path.join(tempDir, "sample.docx");
  const pdfPath = path.join(tempDir, "sample.pdf");
  const markdownPath = path.join(tempDir, "sample.md");
  const textPath = path.join(tempDir, "sample.txt");
  const jsonPath = path.join(tempDir, "sample.json");
  const imagePath = path.join(tempDir, "sample.png");

  await writeSampleWorkbook(xlsxPath);
  await fs.writeFile(csvPath, "Name,Amount\nCoffee,4.25\nNotebook,12.00\n", "utf8");
  await writeSampleDocx(docxPath);
  await writeSamplePdf(pdfPath);
  await fs.writeFile(markdownPath, "# Notes\n\n- Coffee\n- Notebook\n", "utf8");
  await fs.writeFile(textPath, "plain notes\nline two\n", "utf8");
  await fs.writeFile(jsonPath, '{"name":"Coffee","amount":4.25}\n', "utf8");
  await fs.writeFile(
    imagePath,
    Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwADhgGAWjR9awAAAABJRU5ErkJggg==", "base64")
  );

  const workbook = await openDocument(xlsxPath);
  assert(workbook.ok, "Expected sample .xlsx to open");
  assert(workbook.kind === "workbook", "Expected .xlsx to be a workbook");
  assert(workbook.size > 0, "Expected .xlsx metadata to include file size");
  assert(workbook.modifiedAt, "Expected .xlsx metadata to include modified time");
  assert(workbook.sheets[0].rows[1][0] === "Coffee", "Expected .xlsx cell A2 to match");
  workbook.sheets[0].rows[1][0] = "Tea";
  workbook.sheets[0].rows.push(["Total", "=SUM(B2:B3)"]);
  workbook.sheets[0].columnWidths = [180, 96];
  workbook.sheets[0].rowHeights = [30, 44, 30];
  const savedWorkbook = await saveDocument(workbook);
  assert(savedWorkbook.ok, "Expected sample .xlsx to save");
  assert(savedWorkbook.saveMode === "temp-replace", "Expected .xlsx save to use temp-replace mode");
  assert(savedWorkbook.size > 0, "Expected saved .xlsx result to include file size");
  assert(savedWorkbook.modifiedAt, "Expected saved .xlsx result to include modified time");
  assert(savedWorkbook.backupPath, "Expected existing .xlsx save to report a backup path");
  assert(await fileExists(savedWorkbook.backupPath), "Expected reported .xlsx backup to exist");
  assert((await listDocOpenTempFiles(tempDir)).length === 0, "Expected .xlsx temp save files to be cleaned up");
  const workbookBackup = await openDocument(savedWorkbook.backupPath);
  assert(workbookBackup.sheets[0].rows[1][0] === "Coffee", "Expected .xlsx backup to preserve previous value");
  const reopenedWorkbook = await openDocument(xlsxPath);
  assert(reopenedWorkbook.sheets[0].rows[1][0] === "Tea", "Expected saved .xlsx cell A2 to match");
  assert(reopenedWorkbook.sheets[0].columnWidths[0] === 180, "Expected saved .xlsx column width to match");
  assert(reopenedWorkbook.sheets[0].rowHeights[1] === 44, "Expected saved .xlsx row height to match");
  assert(reopenedWorkbook.sheets[0].rows[3][1] === "=SUM(B2:B3)", "Expected saved formula to round-trip as a formula");
  const rawWorkbook = new ExcelJS.Workbook();
  await rawWorkbook.xlsx.readFile(xlsxPath);
  assert(typeof rawWorkbook.worksheets[0].getCell("B2").value === "number", "Expected numeric string cells to save as numbers");
  assert(rawWorkbook.worksheets[0].getCell("B4").value?.formula === "SUM(B2:B3)", "Expected formula cells to save as live formulas");

  const csv = await openDocument(csvPath);
  assert(csv.ok, "Expected sample .csv to open");
  assert(csv.sheets[0].rows[2][1] === "12", "Expected .csv cell B3 to match");
  csv.sheets[0].rows[2][1] = "14";
  const savedCsv = await saveDocument(csv);
  assert(savedCsv.ok, "Expected sample .csv to save");
  assert(savedCsv.saveMode === "temp-replace", "Expected .csv save to use temp-replace mode");
  assert((await listDocOpenTempFiles(tempDir)).length === 0, "Expected .csv temp save files to be cleaned up");
  const reopenedCsv = await openDocument(csvPath);
  assert(reopenedCsv.sheets[0].rows[2][1] === "14", "Expected saved .csv cell B3 to match");

  const word = await openDocument(docxPath);
  assert(word.ok, "Expected sample .docx to open");
  assert(word.size > 0, "Expected .docx metadata to include file size");
  assert(word.modifiedAt, "Expected .docx metadata to include modified time");
  assert(word.html.includes("Hello from Word"), "Expected .docx text to be converted");
  const savedWord = await saveDocument({
    ...word,
    html:
      '<h1>Hello from edited Word</h1><p style="text-align: center;"><font face="Georgia" color="#62a8f2" size="5"><strong>Bold note</strong></font> <span style="background-color: #e8b765;"><s>marked</s></span></p>',
    text: "Hello from edited Word"
  });
  assert(savedWord.ok, "Expected sample .docx to save");
  assert(savedWord.saveMode === "temp-replace", "Expected .docx save to use temp-replace mode");
  assert((await listDocOpenTempFiles(tempDir)).length === 0, "Expected .docx temp save files to be cleaned up");
  const savedWordXml = await readDocxDocumentXml(docxPath);
  assert(savedWordXml.includes("<w:rFonts"), "Expected saved .docx to include font family formatting");
  assert(savedWordXml.includes('<w:color w:val="62A8F2"'), "Expected saved .docx to include text color formatting");
  assert(savedWordXml.includes('<w:shd w:val="clear"'), "Expected saved .docx to include highlight formatting");
  assert(savedWordXml.includes("<w:strike/>"), "Expected saved .docx to include strikethrough formatting");
  assert(savedWordXml.includes('<w:jc w:val="center"/>'), "Expected saved .docx to include paragraph alignment");
  const reopenedWord = await openDocument(docxPath);
  assert(reopenedWord.html.includes("Hello from edited Word"), "Expected saved .docx text to match");
  assert(reopenedWord.html.includes("Bold note"), "Expected saved .docx formatted text to match");

  const pngDataUrl =
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwADhgGAWjR9awAAAABJRU5ErkJggg==";
  const savedRichWord = await saveDocument({
    ...word,
    html:
      '<p>Visit <a href="https://example.com">the site</a> for more.</p>' +
      '<table><tbody><tr><td>A1</td><td>B1</td></tr><tr><td>A2</td><td>B2</td></tr></tbody></table>' +
      `<p><img src="${pngDataUrl}" /></p>`,
    text: ""
  });
  assert(savedRichWord.ok, "Expected rich .docx to save");
  const richXml = await readDocxDocumentXml(docxPath);
  assert(richXml.includes("<w:hyperlink"), "Expected saved .docx to include a hyperlink");
  assert(richXml.includes('<w:rStyle w:val="Hyperlink"/>'), "Expected hyperlink runs to use the Hyperlink style");
  assert(richXml.includes("<w:tbl>"), "Expected saved .docx to include a table");
  assert(richXml.includes("<w:drawing>"), "Expected saved .docx to include an inline image drawing");
  const richRels = await readDocxPart(docxPath, "word/_rels/document.xml.rels");
  assert(richRels.includes("hyperlink"), "Expected saved .docx rels to include the hyperlink relationship");
  assert(richRels.includes("styles.xml"), "Expected saved .docx rels to reference the styles part");
  assert(richRels.includes("media/image1.png"), "Expected saved .docx rels to reference the embedded image");
  const stylesXml = await readDocxPart(docxPath, "word/styles.xml");
  assert(stylesXml.includes('w:styleId="Hyperlink"'), "Expected styles.xml to define the Hyperlink style");
  const mediaEntry = await readDocxHasEntry(docxPath, "word/media/image1.png");
  assert(mediaEntry, "Expected saved .docx to embed the image part");
  const reopenedRichWord = await openDocument(docxPath);
  assert(reopenedRichWord.html.includes("<table"), "Expected reopened .docx to render the table");
  assert(reopenedRichWord.html.includes("<a "), "Expected reopened .docx to render the hyperlink");
  assert(reopenedRichWord.html.includes("<img"), "Expected reopened .docx to render the image");

  const unsafeWord = await openDocument(docxPath);
  const savedUnsafeWord = await saveDocument({
    ...unsafeWord,
    html:
      '<p><a href="javascript:alert(1)" onclick="alert(2)">Unsafe link</a>' +
      '<img src="javascript:alert(3)" />' +
      '<img src="data:text/html;base64,PHNjcmlwdD5iYWQ8L3NjcmlwdD4=" /></p>',
    text: ""
  });
  assert(savedUnsafeWord.ok, "Expected unsafe .docx fixture to save");
  const unsafeXml = await readDocxDocumentXml(docxPath);
  const unsafeRels = await readDocxPart(docxPath, "word/_rels/document.xml.rels");
  assert(!unsafeXml.includes("<w:hyperlink"), "Expected unsafe hyperlink to be removed from .docx XML");
  assert(!unsafeRels.includes("javascript:"), "Expected unsafe URL to be removed from .docx relationships");
  assert(!unsafeRels.includes("media/image"), "Expected unsafe images to be removed from .docx relationships");

  const pdf = await openDocument(pdfPath);
  assert(pdf.ok, "Expected sample .pdf to open");
  assert(pdf.kind === "pdf", "Expected .pdf to be a PDF document");
  assert(pdf.fileUrl.startsWith("file:///"), "Expected .pdf to expose a file URL");

  const markdown = await openDocument(markdownPath);
  assert(markdown.ok, "Expected sample .md to open");
  assert(markdown.kind === "markdown", "Expected .md to be a Markdown document");
  assert(markdown.text.includes("# Notes"), "Expected .md text to load");
  const savedMarkdown = await saveDocument({
    ...markdown,
    text: "# Edited notes\n\n- Tea\n"
  });
  assert(savedMarkdown.ok, "Expected sample .md to save");
  assert(savedMarkdown.saveMode === "temp-replace", "Expected .md save to use temp-replace mode");
  assert(savedMarkdown.backupPath, "Expected existing .md save to report a backup path");
  assert((await fs.readFile(savedMarkdown.backupPath, "utf8")).includes("- Coffee"), "Expected .md backup to preserve previous text");
  const reopenedMarkdown = await openDocument(markdownPath);
  assert(reopenedMarkdown.text.includes("- Tea"), "Expected saved .md text to match");

  const text = await openDocument(textPath);
  assert(text.ok, "Expected sample .txt to open");
  assert(text.kind === "text", "Expected .txt to be a text document");
  assert(text.text.includes("plain notes"), "Expected .txt text to load");
  const savedText = await saveDocument({
    ...text,
    text: "edited notes\n"
  });
  assert(savedText.ok, "Expected sample .txt to save");
  assert(savedText.backupPath, "Expected existing .txt save to report a backup path");
  const reopenedText = await openDocument(textPath);
  assert(reopenedText.text.includes("edited notes"), "Expected saved .txt text to match");

  const json = await openDocument(jsonPath);
  assert(json.ok, "Expected sample .json to open");
  assert(json.kind === "text", "Expected .json to be a text document");
  assert(json.text.includes('"Coffee"'), "Expected .json text to load");

  // Cross-format saves: word -> .md/.txt, markdown -> .docx.
  const wordAsMdPath = path.join(tempDir, "word-export.md");
  const savedWordMd = await saveDocument({
    kind: "word",
    filePath: wordAsMdPath,
    extension: ".md",
    html: "<h1>Title</h1>",
    text: "Title plain",
    markdown: "# Title\n\n**bold** exported\n"
  });
  assert(savedWordMd.ok, "Expected word-to-markdown save to succeed");
  const reopenedWordMd = await openDocument(wordAsMdPath);
  assert(reopenedWordMd.kind === "markdown", "Expected exported .md to reopen as markdown");
  assert(reopenedWordMd.text.includes("**bold** exported"), "Expected word-to-markdown export to use the markdown payload");

  const wordAsTxtPath = path.join(tempDir, "word-export.txt");
  const savedWordTxt = await saveDocument({
    kind: "word",
    filePath: wordAsTxtPath,
    extension: ".txt",
    html: "<p>ignored</p>",
    text: "plain words only",
    markdown: "# ignored"
  });
  assert(savedWordTxt.ok, "Expected word-to-text save to succeed");
  assert((await fs.readFile(wordAsTxtPath, "utf8")).includes("plain words only"), "Expected word-to-text export to use the text payload");

  const mdAsDocxPath = path.join(tempDir, "md-export.docx");
  const savedMdDocx = await saveDocument({
    kind: "markdown",
    filePath: mdAsDocxPath,
    extension: ".docx",
    html: "<h1>Doc from md</h1><p><strong>bold md</strong></p>",
    text: "# Doc from md\n\n**bold md**\n"
  });
  assert(savedMdDocx.ok, "Expected markdown-to-word save to succeed");
  const reopenedMdDocx = await openDocument(mdAsDocxPath);
  assert(reopenedMdDocx.kind === "word", "Expected exported .docx to reopen as word");
  assert(reopenedMdDocx.html.includes("Doc from md"), "Expected markdown-to-word export to carry the converted html");

  const image = await openDocument(imagePath);
  assert(image.ok, "Expected sample .png to open");
  assert(image.kind === "image", "Expected .png to be an image document");
  assert(image.fileUrl.startsWith("file:///"), "Expected .png to expose a file URL");
  // A 2x2 red PNG standing in for canvas.toDataURL output.
  const paintedDataUrl =
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAEklEQVR4nGP8z8DwnwEJMKFzAEEsAQ73wCPUAAAAAElFTkSuQmCC";
  const savedImage = await saveDocument({ ...image, extension: ".png", dataUrl: paintedDataUrl });
  assert(savedImage.ok, "Expected painted image to save");
  assert(savedImage.backupPath, "Expected existing image save to report a backup path");
  assert((await listDocOpenTempFiles(tempDir)).length === 0, "Expected image temp save files to be cleaned up");
  const reopenedImage = await openDocument(imagePath);
  assert(reopenedImage.ok, "Expected saved image to reopen");
  assert(reopenedImage.size > image.size, "Expected saved image bytes to replace the original");
  const copiedImagePath = path.join(tempDir, "copy.png");
  const copiedImage = await saveDocument({ kind: "image", filePath: copiedImagePath, extension: ".png", sourcePath: imagePath });
  assert(copiedImage.ok, "Expected image Save As without edits to copy the source file");
  assert(await fileExists(copiedImagePath), "Expected copied image to exist");

  console.log("Smoke check passed: opened .pdf read-only and opened, edited, saved, and reopened .docx, .xlsx, .csv, .md, .txt, .json, and .png samples.");
}

async function writeSampleWorkbook(filePath) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Expenses");
  worksheet.addRows([
    ["Name", "Amount"],
    ["Coffee", 4.25],
    ["Notebook", 12]
  ]);
  await workbook.xlsx.writeFile(filePath);
}

async function writeSampleDocx(filePath) {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
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
  zip.folder("word").file(
    "document.xml",
    xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r>
              <w:t>Hello from Word</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>`)
  );

  const buffer = await zip.generateAsync({ type: "nodebuffer" });
  await fs.writeFile(filePath, buffer);
}

async function writeSamplePdf(filePath) {
  const pageStream = "BT /F1 24 Tf 72 720 Td (Hello from PDF) Tj ET";
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    `<< /Length ${Buffer.byteLength(pageStream, "ascii")} >>\nstream\n${pageStream}\nendstream`
  ];
  let content = "%PDF-1.4\n";
  const offsets = [0];

  objects.forEach((object, index) => {
    offsets.push(Buffer.byteLength(content, "ascii"));
    content += `${index + 1} 0 obj\n${object}\nendobj\n`;
  });

  const xrefOffset = Buffer.byteLength(content, "ascii");
  content += `xref\n0 ${objects.length + 1}\n`;
  content += "0000000000 65535 f \n";
  offsets.slice(1).forEach((offset) => {
    content += `${String(offset).padStart(10, "0")} 00000 n \n`;
  });
  content += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`;

  await fs.writeFile(filePath, Buffer.from(content, "ascii"));
}

async function readDocxDocumentXml(filePath) {
  const zip = await JSZip.loadAsync(await fs.readFile(filePath));
  return zip.file("word/document.xml").async("string");
}

async function readDocxPart(filePath, partPath) {
  const zip = await JSZip.loadAsync(await fs.readFile(filePath));
  return zip.file(partPath).async("string");
}

async function readDocxHasEntry(filePath, partPath) {
  const zip = await JSZip.loadAsync(await fs.readFile(filePath));
  return Boolean(zip.file(partPath));
}

async function fileExists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function listDocOpenTempFiles(dirPath) {
  const entries = await fs.readdir(dirPath);
  return entries.filter((entry) => entry.startsWith(".docopen-"));
}

function xml(value) {
  return value.replace(/\n\s+/g, " ").trim();
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
