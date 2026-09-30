// Word-document HTML sanitizer: allow-lists for tags/attributes/styles and
// URL vetting. Everything rendered into or saved from the Word editor goes
// through sanitizeWordHtml. Classic script, no state.

const SAFE_WORD_TAGS = new Set([
  "a",
  "b",
  "blockquote",
  "br",
  "div",
  "em",
  "font",
  "h1",
  "h2",
  "h3",
  "i",
  "img",
  "li",
  "ol",
  "p",
  "s",
  "span",
  "strike",
  "strong",
  "table",
  "tbody",
  "td",
  "th",
  "thead",
  "tr",
  "u",
  "ul"
]);
const REMOVE_WORD_TAGS = new Set(["script", "style", "iframe", "object", "embed", "svg", "math", "meta", "link", "base"]);
const SAFE_WORD_ATTRIBUTES = new Set(["align", "alt", "class", "color", "face", "href", "size", "src", "style", "title"]);
const SAFE_STYLE_PROPERTIES = new Set([
  "background-color",
  "color",
  "font-family",
  "font-size",
  "font-style",
  "font-weight",
  "height",
  "max-width",
  "text-align",
  "text-decoration",
  "width"
]);

function sanitizeWordHtml(html) {
  const template = document.createElement("template");
  template.innerHTML = String(html || "");
  sanitizeWordChildren(template.content);
  return template.innerHTML;
}

function sanitizeWordChildren(parent) {
  Array.from(parent.childNodes).forEach((node) => {
    if (node.nodeType === Node.COMMENT_NODE) {
      node.remove();
      return;
    }

    if (node.nodeType !== Node.ELEMENT_NODE) {
      return;
    }

    const tagName = node.tagName.toLowerCase();
    if (REMOVE_WORD_TAGS.has(tagName)) {
      node.remove();
      return;
    }

    if (!SAFE_WORD_TAGS.has(tagName)) {
      node.replaceWith(...Array.from(node.childNodes));
      sanitizeWordChildren(parent);
      return;
    }

    sanitizeWordAttributes(node);
    sanitizeWordChildren(node);
  });
}

function sanitizeWordAttributes(element) {
  Array.from(element.attributes).forEach((attribute) => {
    const name = attribute.name.toLowerCase();
    const value = attribute.value;

    if (name.startsWith("on") || !SAFE_WORD_ATTRIBUTES.has(name)) {
      element.removeAttribute(attribute.name);
      return;
    }

    if (name === "style") {
      const safeStyle = sanitizeStyle(value);
      if (safeStyle) {
        element.setAttribute("style", safeStyle);
      } else {
        element.removeAttribute(attribute.name);
      }
      return;
    }

    if ((name === "href" || name === "src") && !isSafeWordUrl(value, name, element.tagName.toLowerCase())) {
      element.removeAttribute(attribute.name);
    }
  });
}

function sanitizeStyle(style) {
  return String(style || "")
    .split(";")
    .map((declaration) => {
      const [rawName, ...rawValue] = declaration.split(":");
      const name = rawName?.trim().toLowerCase();
      const value = rawValue.join(":").trim();
      if (!name || !value || !SAFE_STYLE_PROPERTIES.has(name) || /url\s*\(|expression|javascript:|-moz-binding/i.test(value)) {
        return "";
      }
      return `${name}: ${value}`;
    })
    .filter(Boolean)
    .join("; ");
}

function isSafeWordUrl(value, attributeName, tagName) {
  const url = String(value || "").trim();
  if (!url) {
    return false;
  }

  if (attributeName === "src" && tagName === "img") {
    return /^data:image\/(?:png|jpe?g|gif);base64,/i.test(url);
  }

  if (attributeName !== "href") {
    return false;
  }

  try {
    const parsed = new URL(url);
    return ["http:", "https:", "mailto:"].includes(parsed.protocol);
  } catch {
    return false;
  }
}
