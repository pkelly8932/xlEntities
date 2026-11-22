// Match all .js files in the current directory EXCEPT:
// - index.js
// - any file starting with "_" (e.g., _template.js, _disabled.js)
const context = require.context("./", false, /^(?!\.\/(?:index\.js|_)).*\.js$/);

const entries = [];

context.keys().forEach((filePath) => {
  const fileName = filePath.replace(/^\.\/|\.js$/g, "");

  // Match a 2-digit priority prefix followed by an underscore (e.g., "20_json")
  const match = fileName.match(/^(\d{2})_(.+)$/);
  const priority = match ? Number(match[1]) : 99;
  const formatName = match ? match[2] : fileName;

  const mod = context(filePath);
  const parser =
    mod.default || Object.values(mod).find((exp) => exp && typeof exp.isFormat === "function");

  if (parser) {
    entries.push({ priority, formatName, parser });
  }
});

// Sort ascending by priority, then ascending alphabetically by formatName
entries.sort((a, b) => {
  if (a.priority !== b.priority) {
    return a.priority - b.priority;
  }
  return a.formatName.localeCompare(b.formatName);
});

export const fromFormats = {};

for (const { formatName, parser } of entries) {
  fromFormats[formatName] = parser;
}
