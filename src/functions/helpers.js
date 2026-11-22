import { fromFormats } from "./input_formats/index.js";

function unwrapSingleCell(rawInput) {
  if (
    Array.isArray(rawInput) &&
    rawInput.length === 1 &&
    Array.isArray(rawInput.at(0)) &&
    rawInput.at(0).length === 1
  ) {
    return rawInput.at(0).at(0);
  }
  return rawInput;
}

export function parseInput(rawInput, format = "Auto") {
  const unwrapped = unwrapSingleCell(rawInput);

  if (
    unwrapped &&
    typeof unwrapped === "object" &&
    !Array.isArray(unwrapped) &&
    unwrapped.type === "Entity"
  ) {
    return unwrapEntity(unwrapped);
  }

  const normalizedFormat = format ? String(format).trim().toLowerCase() : "auto";

  if (normalizedFormat !== "auto") {
    const entry = Object.entries(fromFormats).find(
      ([key]) => key.toLowerCase() === normalizedFormat
    );
    if (!entry) {
      throw new Error(
        `Unknown format: "${format}". Available: ${Object.keys(fromFormats).join(", ")}`
      );
    }
    const [, parser] = entry;
    const valueToParse = normalizedFormat === "array" ? rawInput : unwrapped;
    return parser.toJS(valueToParse);
  }

  for (const parser of Object.values(fromFormats)) {
    if (parser.isFormat(unwrapped)) {
      const parsed = parser.toJS(unwrapped);
      if (typeof parsed === "string" && parsed !== unwrapped) {
        return parseInput(parsed, "Auto");
      }
      return parsed;
    }
  }
  return unwrapped;
}

export function applyHeaders(data) {
  if (!Array.isArray(data) || data.length === 0 || !Array.isArray(data.at(0))) {
    return data;
  }

  const colCount = data.at(0).length;

  // 2 columns: use first column as keys -> { key: value }
  if (colCount === 2) {
    const obj = {};
    for (const row of data) {
      if (Array.isArray(row) && row.length >= 2) {
        const key = String(row.at(0) ?? "").trim();
        if (key !== "") {
          obj[key] = row.at(1);
        }
      }
    }
    return obj;
  }

  // More than 2 columns: use first row as keys -> [{ header1: val1, ... }, ...]
  if (colCount > 2) {
    const headerRow = data.at(0).map((h) => String(h ?? "").trim());
    return data.slice(1).map((row) => {
      const record = {};
      headerRow.forEach((key, idx) => {
        if (key !== "") {
          record[key] = Array.isArray(row) ? row[idx] : undefined;
        }
      });
      return record;
    });
  }

  return data;
}

export function unwrapEntity(cellValue) {
  if (!cellValue || typeof cellValue !== "object") return cellValue;
  if (Array.isArray(cellValue)) return cellValue.map(unwrapEntity);

  if (cellValue.type === "Entity" && cellValue.properties) {
    const keys = Object.keys(cellValue.properties);
    const isIndexedArray =
      keys.length > 0 && keys.every((k, idx) => String(idx) === k || /^\d+$/.test(k));

    if (cellValue._isArray || isIndexedArray) {
      return keys
        .sort((a, b) => Number(a) - Number(b))
        .map((k) => unwrapEntity(cellValue.properties[k]));
    }

    const result = {};
    for (const [key, val] of Object.entries(cellValue.properties)) {
      result[key] = unwrapEntity(val);
    }
    return result;
  }

  if (cellValue.type === "Array" && Array.isArray(cellValue.elements)) {
    return cellValue.elements.map((row) => row.map(unwrapEntity));
  }

  return cellValue.basicValue !== undefined ? cellValue.basicValue : cellValue;
}

export function toExcelValue(val, entityName = null) {
  if (val === null || val === undefined) return "";
  if (typeof val === "string" || typeof val === "number" || typeof val === "boolean") {
    return val;
  }

  if (Array.isArray(val)) {
    const properties = {};
    val.forEach((item, idx) => {
      properties[String(idx)] = toExcelValue(item);
    });
    return {
      type: "Entity",
      text: entityName || val._defaultName || "Array",
      properties
    };
  }

  if (typeof val === "object") {
    const properties = {};
    const keys = Object.keys(val);
    for (const key of keys) {
      properties[key] = toExcelValue(val[key], key);
    }
    const defaultText =
      val._defaultName ||
      (keys.length ? `{${keys.slice(0, 3).join(", ")}${keys.length > 3 ? "..." : ""}}` : "{}");
    return {
      type: "Entity",
      text: entityName || defaultText,
      properties
    };
  }

  return String(val);
}

export function resolvePath(data, path) {
  if (path === undefined || path === null || path === "") return data;
  const tokens = String(path)
    .replace(/\[(\d+)\]/g, ".$1")
    .split(".")
    .filter(Boolean);

  let current = data;
  for (const token of tokens) {
    if (current === null || current === undefined) return undefined;
    current = current[token];
  }
  return current;
}

function isPlainObject(val) {
  return val !== null && typeof val === "object" && !Array.isArray(val);
}

/**
 * Expands a single object into 1 or more flat row maps up to `maxLevels`.
 */
function expandObjectRows(obj, padValue, maxLevels, currentLevel = 1, prefix = "") {
  let rows = [{}];

  for (const [key, val] of Object.entries(obj)) {
    const colKey = prefix ? `${prefix}.${key}` : key;

    if (currentLevel >= maxLevels && (isPlainObject(val) || Array.isArray(val))) {
      const entityCell = toExcelValue(val, key);
      rows.forEach((r) => {
        r[colKey] = entityCell;
      });
      continue;
    }

    if (isPlainObject(val)) {
      const subRows = expandObjectRows(val, padValue, maxLevels, currentLevel + 1, colKey);
      const nextRows = [];
      for (const parentRow of rows) {
        subRows.forEach((subRow, sIdx) => {
          const merged = {};
          for (const [pk, pv] of Object.entries(parentRow)) {
            merged[pk] = sIdx === 0 || padValue === null ? pv : padValue;
          }
          Object.assign(merged, subRow);
          nextRows.push(merged);
        });
      }
      rows = nextRows;
    } else if (Array.isArray(val)) {
      if (val.length === 0) {
        rows.forEach((r) => {
          r[colKey] = "";
        });
        continue;
      }

      const elementRowGroups = val.map((item) => {
        if (isPlainObject(item)) {
          return expandObjectRows(item, padValue, maxLevels, currentLevel + 1, colKey);
        }
        if (Array.isArray(item)) {
          return [{ [colKey]: toExcelValue(item, key) }];
        }
        return [{ [colKey]: item === null || item === undefined ? "" : item }];
      });

      const expandedSubRows = elementRowGroups.flat();
      const nextRows = [];
      for (const parentRow of rows) {
        expandedSubRows.forEach((subRow, sIdx) => {
          const merged = {};
          for (const [pk, pv] of Object.entries(parentRow)) {
            merged[pk] = sIdx === 0 || padValue === null ? pv : padValue;
          }
          Object.assign(merged, subRow);
          nextRows.push(merged);
        });
      }
      rows = nextRows;
    } else {
      const scalarVal = val === null || val === undefined ? "" : val;
      rows.forEach((r) => {
        r[colKey] = scalarVal;
      });
    }
  }

  return rows;
}

/**
 * Recursively expands nested non-object arrays up to `maxLevels`.
 */
function expandArrayGrid(arr, fill, maxLevels, currentLevel = 1) {
  if (currentLevel >= maxLevels) {
    return arr.map((item) => [toExcelValue(item)]);
  }

  const expandedRows = [];
  for (const row of arr) {
    if (!Array.isArray(row)) {
      expandedRows.push([isPlainObject(row) ? toExcelValue(row) : (row ?? fill)]);
      continue;
    }

    if (currentLevel + 1 >= maxLevels) {
      expandedRows.push(
        row.map((cell) =>
          cell === null || cell === undefined
            ? fill
            : Array.isArray(cell) || isPlainObject(cell)
            ? toExcelValue(cell)
            : cell
        )
      );
    } else {
      const subGrid = expandArrayGrid(row, fill, maxLevels, currentLevel + 1);
      expandedRows.push(...subGrid);
    }
  }

  const maxCols = Math.max(1, ...expandedRows.map((r) => r.length));
  return expandedRows.map((r) =>
    r.length < maxCols ? [...r, ...Array(maxCols - r.length).fill(fill)] : r
  );
}

export function flattenData(data, padValue = null, levels = 1) {
  if (data === null || data === undefined) return [[""]];

  // Scalar
  if (!isPlainObject(data) && !Array.isArray(data)) {
    return [[data]];
  }

  // Array input
  if (Array.isArray(data)) {
    if (data.length === 0) return [[""]];

    // Array of objects -> union of keys as headers (row 0), values in object index rows
    if (data.some(isPlainObject)) {
      const allRowMaps = [];
      const headerSet = new Set();

      for (const item of data) {
        if (isPlainObject(item)) {
          const expanded = expandObjectRows(item, padValue, levels, 1, "");
          for (const rowMap of expanded) {
            Object.keys(rowMap).forEach((k) => headerSet.add(k));
            allRowMaps.push(rowMap);
          }
        } else {
          headerSet.add("value");
          allRowMaps.push({ value: toExcelValue(item) });
        }
      }

      const headers = Array.from(headerSet);
      const grid = [headers];
      for (const rowMap of allRowMaps) {
        grid.push(headers.map((h) => (rowMap[h] !== undefined ? rowMap[h] : "")));
      }
      return grid;
    }

    // Non-object array (1D, 2D, 3D, etc.) -> expand according to `levels`
    const fill = padValue !== null && padValue !== undefined ? padValue : "";
    return expandArrayGrid(data, fill, levels, 1);
  }

  // Single Object -> keys as headers (first row), values as second row (+ child rows if arrays expanded)
  const rowMaps = expandObjectRows(data, padValue, levels, 1, "");
  const headerSet = new Set();
  rowMaps.forEach((rm) => Object.keys(rm).forEach((k) => headerSet.add(k)));

  const headers = Array.from(headerSet);
  if (headers.length === 0) return [[""]];

  const grid = [headers];
  for (const rowMap of rowMaps) {
    grid.push(headers.map((h) => (rowMap[h] !== undefined ? rowMap[h] : "")));
  }
  return grid;
}
