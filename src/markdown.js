// Markdown <-> HTML conversion for the markdown preview and cross-format
// Save As (md -> docx, docx -> md), plus escapeHtml. Pure transforms over
// strings/DOM templates; classic script, no state.

function markdownToHtml(markdown) {
  const lines = String(markdown || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  const html = [];
  let paragraph = [];
  let list = null;

  const flushParagraph = () => {
    if (paragraph.length === 0) {
      return;
    }
    html.push(`<p>${renderMarkdownInline(paragraph.join(" "))}</p>`);
    paragraph = [];
  };

  const flushList = () => {
    if (!list) {
      return;
    }
    html.push(`<${list.type}>${list.items.map(renderMarkdownListItem).join("")}</${list.type}>`);
    list = null;
  };

  const flushBlocks = () => {
    flushParagraph();
    flushList();
  };

  const pushListItem = (type, text) => {
    flushParagraph();
    if (list && list.type !== type) {
      flushList();
    }
    if (!list) {
      list = { type, items: [] };
    }
    list.items.push(text);
  };

  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];

    if (/^\s*```/.test(line)) {
      flushBlocks();
      const codeLines = [];
      i += 1;
      while (i < lines.length && !/^\s*```/.test(lines[i])) {
        codeLines.push(lines[i]);
        i += 1;
      }
      html.push(`<pre><code>${escapeHtml(codeLines.join("\n"))}</code></pre>`);
      continue;
    }

    if (/^\s*$/.test(line)) {
      flushBlocks();
      continue;
    }

    if (/^\s*([-*_])(?:\s*\1){2,}\s*$/.test(line)) {
      flushBlocks();
      html.push("<hr />");
      continue;
    }

    const heading = /^(#{1,6})\s+(.+?)\s*#*\s*$/.exec(line);
    if (heading) {
      flushBlocks();
      html.push(`<h${heading[1].length}>${renderMarkdownInline(heading[2])}</h${heading[1].length}>`);
      continue;
    }

    if (line.includes("|") && i + 1 < lines.length && /^\s*\|?[\s:|-]*-[\s:|-]*\|?\s*$/.test(lines[i + 1])) {
      flushBlocks();
      const header = splitMarkdownTableRow(line);
      const aligns = splitMarkdownTableRow(lines[i + 1]).map(markdownCellAlign);
      const rows = [];
      i += 2;
      while (i < lines.length && lines[i].includes("|") && !/^\s*$/.test(lines[i])) {
        rows.push(splitMarkdownTableRow(lines[i]));
        i += 1;
      }
      i -= 1;
      html.push(renderMarkdownTable(header, aligns, rows));
      continue;
    }

    const quote = /^\s*>\s?(.*)$/.exec(line);
    if (quote) {
      flushBlocks();
      html.push(`<blockquote>${renderMarkdownInline(quote[1])}</blockquote>`);
      continue;
    }

    const orderedItem = /^\s*\d+[.)]\s+(.+)$/.exec(line);
    if (orderedItem) {
      pushListItem("ol", orderedItem[1]);
      continue;
    }

    const unorderedItem = /^\s*[-*+]\s+(.+)$/.exec(line);
    if (unorderedItem) {
      pushListItem("ul", unorderedItem[1]);
      continue;
    }

    flushList();
    paragraph.push(line.trim());
  }

  flushBlocks();

  return html.length > 0 ? html.join("") : "<p>This Markdown file is empty.</p>";
}

function renderMarkdownListItem(item) {
  const task = /^\[([ xX])\]\s+(.*)$/.exec(item);
  if (task) {
    const checked = task[1].toLowerCase() === "x" ? " checked" : "";
    return `<li class="task-item"><input type="checkbox" disabled${checked} /> ${renderMarkdownInline(task[2])}</li>`;
  }
  return `<li>${renderMarkdownInline(item)}</li>`;
}

function splitMarkdownTableRow(line) {
  let cells = line.trim();
  if (cells.startsWith("|")) {
    cells = cells.slice(1);
  }
  if (cells.endsWith("|")) {
    cells = cells.slice(0, -1);
  }
  return cells.split("|").map((cell) => cell.trim());
}

function markdownCellAlign(separator) {
  const left = separator.startsWith(":");
  const right = separator.endsWith(":");
  if (left && right) {
    return "center";
  }
  return right ? "right" : left ? "left" : "";
}

function renderMarkdownTable(header, aligns, rows) {
  const alignAttr = (index) => (aligns[index] ? ` style="text-align:${aligns[index]}"` : "");
  const head = header.map((cell, index) => `<th${alignAttr(index)}>${renderMarkdownInline(cell)}</th>`).join("");
  const body = rows
    .map((row) => `<tr>${header.map((_cell, index) => `<td${alignAttr(index)}>${renderMarkdownInline(row[index] || "")}</td>`).join("")}</tr>`)
    .join("");
  return `<table class="md-table"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
}

function renderMarkdownInline(value) {
  return String(value)
    .split(/(`[^`]+`)/)
    .map((part) => {
      if (part.length >= 2 && part.startsWith("`") && part.endsWith("`")) {
        return `<code>${escapeHtml(part.slice(1, -1))}</code>`;
      }
      return formatMarkdownText(part);
    })
    .join("");
}

function formatMarkdownText(text) {
  let html = escapeHtml(text);
  html = html.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, (_match, alt, url) => {
    const safeUrl = safePreviewImage(url);
    return safeUrl ? `<img src="${escapeHtml(safeUrl)}" alt="${alt}" />` : escapeHtml(alt);
  });
  html = html.replace(/~~([^~]+)~~/g, "<s>$1</s>");
  html = html.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");
  html = html.replace(/__([^_]+)__/g, "<strong>$1</strong>");
  html = html.replace(/\*([^*]+)\*/g, "<em>$1</em>");
  html = html.replace(/\b_([^_]+)_\b/g, "<em>$1</em>");
  html = html.replace(/\[([^\]]+)\]\(([^)]+)\)/g, (_match, label, url) => {
    const safeUrl = safePreviewLink(url);
    return safeUrl ? `<a href="${escapeHtml(safeUrl)}">${label}</a>` : label;
  });
  return html;
}

function safePreviewImage(value) {
  const link = String(value || "").trim();
  if (/^data:image\//i.test(link)) {
    return link;
  }
  try {
    const parsed = new URL(link);
    return ["http:", "https:"].includes(parsed.protocol) ? link : null;
  } catch {
    return null;
  }
}

function safePreviewLink(value) {
  const link = String(value || "").trim();
  try {
    const parsed = new URL(link);
    return ["http:", "https:", "mailto:"].includes(parsed.protocol) ? link : null;
  } catch {
    return null;
  }
}

function htmlToPlainText(html) {
  const template = document.createElement("template");
  template.innerHTML = String(html || "");
  return template.content.textContent || "";
}

// ponytail: covers the tags the Word editor produces (headings, emphasis,
// lists, links, tables, blockquotes); anything else falls back to its text.
function htmlToMarkdown(html) {
  const template = document.createElement("template");
  template.innerHTML = String(html || "");
  return `${markdownBlocks(template.content).replace(/\n{3,}/g, "\n\n").trim()}\n`;
}

const MARKDOWN_BLOCK_TAGS = new Set(["p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "table", "blockquote", "pre"]);

function markdownBlocks(parent) {
  let out = "";
  parent.childNodes.forEach((node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.nodeValue.trim();
      if (text) {
        out += `${text}\n\n`;
      }
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      return;
    }

    const tag = node.tagName.toLowerCase();
    const heading = /^h([1-6])$/.exec(tag);
    if (heading) {
      out += `${"#".repeat(Number(heading[1]))} ${markdownInline(node).trim()}\n\n`;
    } else if (tag === "blockquote") {
      out += `> ${markdownInline(node).trim()}\n\n`;
    } else if (tag === "ul" || tag === "ol") {
      out += `${markdownList(node, tag === "ol")}\n`;
    } else if (tag === "table") {
      out += `${markdownTable(node)}\n`;
    } else if (tag === "pre") {
      out += `\`\`\`\n${node.textContent.replace(/\n$/, "")}\n\`\`\`\n\n`;
    } else if (tag === "br") {
      out += "\n";
    } else if (Array.from(node.children).some((child) => MARKDOWN_BLOCK_TAGS.has(child.tagName.toLowerCase()))) {
      out += markdownBlocks(node); // container element — unwrap its blocks
    } else {
      const inline = markdownInline(node).trim();
      if (inline) {
        out += `${inline}\n\n`;
      }
    }
  });
  return out;
}

function markdownInline(parent) {
  let out = "";
  parent.childNodes.forEach((node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      out += node.nodeValue.replace(/\s+/g, " ");
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      return;
    }

    const tag = node.tagName.toLowerCase();
    const inner = markdownInline(node);
    const trimmed = inner.trim();
    if (tag === "strong" || tag === "b") {
      out += trimmed ? ` **${trimmed}** ` : "";
    } else if (tag === "em" || tag === "i") {
      out += trimmed ? ` *${trimmed}* ` : "";
    } else if (tag === "s" || tag === "strike") {
      out += trimmed ? ` ~~${trimmed}~~ ` : "";
    } else if (tag === "code") {
      out += `\`${node.textContent}\``;
    } else if (tag === "a") {
      const href = node.getAttribute("href") || "";
      out += `[${trimmed || href}](${href})`;
    } else if (tag === "img") {
      out += node.getAttribute("alt") ? `![${node.getAttribute("alt")}]()` : "";
    } else if (tag === "br") {
      out += "\n";
    } else {
      out += inner;
    }
  });
  return out.replace(/ {2,}/g, " ");
}

function markdownList(list, ordered, indent = "") {
  let out = "";
  let index = 0;
  Array.from(list.children).forEach((item) => {
    if (item.tagName.toLowerCase() !== "li") {
      return;
    }
    index += 1;
    const itemText = item.cloneNode(true);
    itemText.querySelectorAll("ul, ol").forEach((nested) => nested.remove());
    out += `${indent}${ordered ? `${index}.` : "-"} ${markdownInline(itemText).trim()}\n`;
    item.querySelectorAll(":scope > ul, :scope > ol").forEach((nested) => {
      out += markdownList(nested, nested.tagName.toLowerCase() === "ol", `${indent}  `);
    });
  });
  return out;
}

function markdownTable(table) {
  const rows = Array.from(table.querySelectorAll("tr"))
    .map((tr) => Array.from(tr.querySelectorAll("td, th")).map((cell) => markdownInline(cell).trim().replace(/\|/g, "\\|")))
    .filter((row) => row.length > 0);
  if (rows.length === 0) {
    return "";
  }

  const width = Math.max(...rows.map((row) => row.length));
  const line = (cells) => `| ${Array.from({ length: width }, (_unused, i) => cells[i] || "").join(" | ")} |`;
  return [line(rows[0]), `| ${Array(width).fill("---").join(" | ")} |`, ...rows.slice(1).map(line), ""].join("\n");
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
