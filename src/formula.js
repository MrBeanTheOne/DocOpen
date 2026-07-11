// Tiny spreadsheet formula evaluator shared by the renderer (script tag) and
// the smoke test (require). Supports =A1-style refs, + - * / ( ), and
// SUM/AVERAGE/MIN/MAX/COUNT over ranges like A1:B3.
(function (globalScope) {
  const FUNCTIONS = {
    SUM: (values) => values.reduce((sum, value) => sum + value, 0),
    AVERAGE: (values) => (values.length === 0 ? 0 : values.reduce((sum, value) => sum + value, 0) / values.length),
    MIN: (values) => (values.length === 0 ? 0 : Math.min(...values)),
    MAX: (values) => (values.length === 0 ? 0 : Math.max(...values)),
    COUNT: (values) => values.length
  };

  function isFormula(value) {
    return typeof value === "string" && value.startsWith("=") && value.trim().length > 1;
  }

  // Display text for any cell: formulas evaluate, everything else passes through.
  function evaluateCellDisplay(sheet, rowIndex, columnIndex) {
    try {
      return formatValue(evaluateCell(sheet, rowIndex, columnIndex, new Set()));
    } catch (error) {
      return error.code || "#ERROR!";
    }
  }

  function evaluateCell(sheet, rowIndex, columnIndex, visiting) {
    const raw = sheet.rows[rowIndex]?.[columnIndex] ?? "";
    if (!isFormula(raw)) {
      return String(raw);
    }

    const key = `${rowIndex},${columnIndex}`;
    if (visiting.has(key)) {
      throw formulaError("#REF!");
    }

    visiting.add(key);
    try {
      const body = raw.slice(1).trim();
      const bareRef = parseBareRef(body);
      if (bareRef) {
        // A plain =A1 mirrors the referenced cell, including text.
        return evaluateCell(sheet, bareRef.row, bareRef.column, visiting);
      }
      return evaluateFormula(sheet, body, visiting);
    } finally {
      visiting.delete(key);
    }
  }

  function evaluateFormula(sheet, source, visiting) {
    const tokens = tokenize(source);
    let position = 0;

    const peek = () => tokens[position];
    const take = (type) => {
      const token = tokens[position];
      if (!token || token.type !== type) {
        throw formulaError("#ERROR!");
      }
      position += 1;
      return token;
    };

    function parseExpression() {
      let value = parseTerm();
      while (peek() && (peek().type === "+" || peek().type === "-")) {
        const operator = tokens[position].type;
        position += 1;
        const right = parseTerm();
        value = operator === "+" ? value + right : value - right;
      }
      return value;
    }

    function parseTerm() {
      let value = parseFactor();
      while (peek() && (peek().type === "*" || peek().type === "/")) {
        const operator = tokens[position].type;
        position += 1;
        const right = parseFactor();
        value = operator === "*" ? value * right : value / right;
      }
      return value;
    }

    function parseFactor() {
      const token = peek();
      if (!token) {
        throw formulaError("#ERROR!");
      }
      if (token.type === "-") {
        position += 1;
        return -parseFactor();
      }
      if (token.type === "+") {
        position += 1;
        return parseFactor();
      }
      if (token.type === "number") {
        position += 1;
        return token.value;
      }
      if (token.type === "(") {
        position += 1;
        const value = parseExpression();
        take(")");
        return value;
      }
      if (token.type === "ident") {
        position += 1;
        if (peek() && peek().type === "(") {
          const fn = FUNCTIONS[token.value];
          if (!fn) {
            throw formulaError("#NAME?");
          }
          position += 1;
          const values = [];
          if (peek() && peek().type !== ")") {
            values.push(...parseArgument());
            while (peek() && peek().type === ",") {
              position += 1;
              values.push(...parseArgument());
            }
          }
          take(")");
          return fn(values);
        }
        return numberAtRef(parseCellRef(token.value));
      }
      throw formulaError("#ERROR!");
    }

    // One function argument: a range (numeric cells only, like Excel) or an expression.
    function parseArgument() {
      const token = peek();
      if (token && token.type === "ident" && tokens[position + 1]?.type === ":") {
        const start = parseCellRef(token.value);
        position += 2;
        const end = parseCellRef(take("ident").value);
        const values = [];
        for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row += 1) {
          for (let column = Math.min(start.column, end.column); column <= Math.max(start.column, end.column); column += 1) {
            const value = numericValueAt(row, column);
            if (value !== null) {
              values.push(value);
            }
          }
        }
        return values;
      }
      return [parseExpression()];
    }

    function numberAtRef(ref) {
      const value = evaluateCell(sheet, ref.row, ref.column, visiting);
      if (typeof value === "number") {
        return value;
      }
      const trimmed = cleanCellText(value);
      if (trimmed === "") {
        return 0;
      }
      const parsed = Number(trimmed);
      if (Number.isNaN(parsed)) {
        throw formulaError("#VALUE!");
      }
      return parsed;
    }

    function numericValueAt(row, column) {
      const value = evaluateCell(sheet, row, column, visiting);
      if (typeof value === "number") {
        return value;
      }
      const trimmed = cleanCellText(value);
      if (trimmed === "") {
        return null;
      }
      const parsed = Number(trimmed);
      return Number.isNaN(parsed) ? null : parsed;
    }

    const result = parseExpression();
    if (position < tokens.length) {
      throw formulaError("#ERROR!");
    }
    return result;
  }

  function tokenize(source) {
    const tokens = [];
    let index = 0;

    while (index < source.length) {
      const char = source[index];
      if (/\s/.test(char)) {
        index += 1;
        continue;
      }
      if (/[0-9.]/.test(char)) {
        const match = /^\d*\.?\d+/.exec(source.slice(index));
        if (!match) {
          throw formulaError("#ERROR!");
        }
        tokens.push({ type: "number", value: Number(match[0]) });
        index += match[0].length;
        continue;
      }
      if (/[A-Za-z$]/.test(char)) {
        const match = /^[$A-Za-z][$A-Za-z0-9]*/.exec(source.slice(index));
        tokens.push({ type: "ident", value: match[0].replace(/\$/g, "").toUpperCase() });
        index += match[0].length;
        continue;
      }
      if ("+-*/():,".includes(char)) {
        tokens.push({ type: char });
        index += 1;
        continue;
      }
      throw formulaError("#ERROR!");
    }

    return tokens;
  }

  function parseBareRef(body) {
    const match = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(body);
    if (!match) {
      return null;
    }
    return { row: Number(match[2]) - 1, column: columnIndexFromName(match[1].toUpperCase()) };
  }

  function parseCellRef(ident) {
    const match = /^([A-Z]{1,3})(\d+)$/.exec(ident);
    if (!match) {
      throw formulaError("#NAME?");
    }
    return { row: Number(match[2]) - 1, column: columnIndexFromName(match[1]) };
  }

  function columnIndexFromName(name) {
    let index = 0;
    for (const char of name) {
      index = index * 26 + (char.charCodeAt(0) - 64);
    }
    return index - 1;
  }

  // contentEditable can leave zero-width characters behind; strip them before
  // deciding a cell is "not a number".
  function cleanCellText(text) {
    return text.replace(/[\u200B-\u200D\uFEFF]/g, "").trim();
  }

  function formatValue(value) {
    if (typeof value !== "number") {
      return value;
    }
    if (!Number.isFinite(value)) {
      return "#DIV/0!";
    }
    return String(Math.round(value * 1e10) / 1e10);
  }

  function formulaError(code) {
    const error = new Error(code);
    error.code = code;
    return error;
  }

  const api = { isFormula, evaluateCellDisplay };
  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  } else {
    globalScope.docFormula = api;
  }
})(typeof window !== "undefined" ? window : globalThis);
