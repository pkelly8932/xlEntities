/* global CustomFunctions */
import {
  parseInput,
  applyHeaders,
  toExcelValue,
  resolvePath,
  flattenData
} from "./helpers.js";

/**
 * Creates an Excel Entity card or value from JSON, XML, array text, or a range.
 * @customfunction
 * @param {any[][]} input Raw JSON, XML, array literal string, or cell range.
 * @param {string} [name] Optional text displayed in the cell next to the entity icon.
 * @param {boolean} [headers] If true, uses first column (2 cols) or first row (>2 cols) as keys. Default is false.
 * @param {string} [format] Force a specific parser (e.g., "json", "xml", "arrayText", "array") or "Auto".
 * @returns {any}
 */
export function NEW(input, name, headers, format) {
  const entityName =
    name !== null && name !== undefined && String(name).trim() !== "" ? String(name) : null;
  const useHeaders = headers ?? false;
  const selectedFormat = format ?? "Auto";

  try {
    let data = parseInput(input, selectedFormat);
    if (useHeaders) {
      data = applyHeaders(data);
    }
    return toExcelValue(data, entityName);
  } catch (err) {
    throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, err.message);
  }
}
CustomFunctions.associate("NEW", NEW);

/**
 * Extracts a nested value by dot/bracket path from an Entity or raw input.
 * @customfunction
 * @param {any} source Entity cell or raw data input.
 * @param {string} [path] Dot or bracket path (e.g., "user.address.city" or "items.name").
 * @returns {any}
 */
export function GET(source, path) {
  const data = parseInput(source);
  const target = resolvePath(data, path);
  if (target === undefined) {
    throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable, "Path not found");
  }
  const tokens = path ? String(path).replace(/\[(\d+)\]/g, ".$1").split(".").filter(Boolean) : [];
  const lastKey = tokens.length > 0 ? tokens[tokens.length - 1] : null;
  return toExcelValue(target, lastKey);
}
CustomFunctions.associate("GET", GET);

/**
 * Returns the keys or indices of an Entity or object.
 * @customfunction
 * @param {any} source Entity cell or raw data input.
 * @param {boolean} [spill] If true (default), spills vertically. If false, wraps in an additional array level.
 * @returns {any[][]}
 */
export function KEYS(source, spill) {
  const shouldSpill = spill ?? true;
  const data = parseInput(source);

  if (data === null || typeof data !== "object") {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "Source is not an object or array"
    );
  }

  const keys = Object.keys(data);
  const column = keys.length ? keys.map((k) => [k]) : [[""]];
  return shouldSpill ? column : [[column]];
}
CustomFunctions.associate("KEYS", KEYS);

/**
 * Flattens an Entity, object, or array into a tabular grid up to a specified depth.
 * @customfunction
 * @param {any} source Entity cell or raw data input.
 * @param {any} [pad_value] Value to fill sibling columns in child rows (default null copies parent row values).
 * @param {number} [levels] Nesting depth to expand (default 1; deeper values stay as Entities).
 * @returns {any[][]}
 */
export function FLATTEN(source, pad_value, levels) {
  const padVal = pad_value === undefined ? null : pad_value;
  const maxLevels = levels ?? 1;
  const data = parseInput(source);
  return flattenData(data, padVal, maxLevels);
}
CustomFunctions.associate("FLATTEN", FLATTEN);
