export const arrayTextFormat = {
  isFormat(input) {
    if (typeof input !== "string") return false;
    const trimmed = input.trim();
    return trimmed.startsWith("{") && trimmed.endsWith("}");
  },
  toJS(input) {
    return this.parseArrayText(input.trim());
  },
  parseArrayText(str) {
    const trimmed = str.trim();
    if (!trimmed.startsWith("{") || !trimmed.endsWith("}")) {
      return this.parseLiteral(trimmed);
    }

    const inner = trimmed.slice(1, -1).trim();
    if (!inner) return [];

    const rows = this.splitTopLevel(inner, ";");
    const parsedRows = rows.map((row) =>
      this.splitTopLevel(row, ",").map((cell) => {
        const cleanCell = cell.trim();
        if (cleanCell.startsWith("{") && cleanCell.endsWith("}")) {
          return this.parseArrayText(cleanCell);
        }
        return this.parseLiteral(cleanCell);
      })
    );

    // Unwrap single-row literals (e.g., "{10,20,30}" ->)
    if (parsedRows.length === 1) {
      return parsedRows.at(0);
    }
    // Unwrap single-column nested rows (e.g., "{{10,20,30};{40,50,60}}" -> [[10,20,30], [40,50,60]])
    if (parsedRows.every((r) => Array.isArray(r) && r.length === 1 && Array.isArray(r.at(0)))) {
      return parsedRows.map((r) => r.at(0));
    }
    return parsedRows;
  },
  splitTopLevel(str, delimiter) {
    const parts = [];
    let current = "";
    let inQuotes = false;
    let braceDepth = 0;

    for (let i = 0; i < str.length; i++) {
      const ch = str[i];
      if (ch === '"' && str[i - 1] !== "\\") {
        inQuotes = !inQuotes;
      } else if (!inQuotes) {
        if (ch === "{") braceDepth++;
        else if (ch === "}") braceDepth--;
      }

      if (ch === delimiter && !inQuotes && braceDepth === 0) {
        parts.push(current.trim());
        current = "";
      } else {
        current += ch;
      }
    }
    parts.push(current.trim());
    return parts;
  },
  parseLiteral(token) {
    if (token.startsWith('"') && token.endsWith('"')) {
      return token.slice(1, -1).replace(/\\"/g, '"');
    }
    if (token.toUpperCase() === "TRUE") return true;
    if (token.toUpperCase() === "FALSE") return false;
    if (!isNaN(token) && token !== "") return Number(token);
    return token;
  }
};
