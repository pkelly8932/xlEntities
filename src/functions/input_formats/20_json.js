export const jsonFormat = {
  isFormat(input) {
    if (typeof input !== "string") return false;
    const trimmed = input.trim();
    if (
      !(trimmed.startsWith("{") && trimmed.endsWith("}")) &&
      !(trimmed.startsWith("[") && trimmed.endsWith("]"))
    ) {
      return false;
    }
    try {
      JSON.parse(trimmed);
      return true;
    } catch {
      return false;
    }
  },
  toJS(input) {
    return JSON.parse(input.trim());
  }
};
