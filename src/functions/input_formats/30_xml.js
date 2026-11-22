export const xmlFormat = {
  isFormat(input) {
    if (typeof input !== "string") return false;
    const trimmed = input.trim();
    return trimmed.startsWith("<") && trimmed.endsWith(">") && /^<[a-zA-Z_][\w:.-]*/.test(trimmed);
  },
  toJS(input) {
    const cleaned = input
      .trim()
      .replace(/<\?xml[\s\S]*?\?>/gi, "")
      .replace(/<!--[\s\S]*?-->/g, "")
      .trim();

    const root = this.parseElement(cleaned, { pos: 0 });
    if (!root) {
      throw new Error("Invalid XML input");
    }

    // D3: If root has a primitive/null value (e.g. <ping>ok</ping>), keep { ping: "ok" }
    if (root.value === null || typeof root.value !== "object") {
      const result = { [root.tagName]: root.value };
      Object.defineProperty(result, "_defaultName", { value: root.tagName, enumerable: false });
      return result;
    }

    // D1 & D2: Unwrap the root tag to its object/array value and attach root.tagName as fallback label
    const unwrapped = root.value;
    Object.defineProperty(unwrapped, "_defaultName", { value: root.tagName, enumerable: false });
    return unwrapped;
  },
  parseElement(xml, state) {
    while (state.pos < xml.length && /\s/.test(xml[state.pos])) state.pos++;
    if (xml[state.pos] !== "<" || xml[state.pos + 1] === "/") return null;

    const tagClose = xml.indexOf(">", state.pos);
    if (tagClose === -1) throw new Error("Unclosed XML tag");

    const rawTag = xml.slice(state.pos + 1, tagClose).trim();
    const selfClosing = rawTag.endsWith("/");
    const tagContent = selfClosing ? rawTag.slice(0, -1).trim() : rawTag;

    const spaceIdx = tagContent.search(/\s/);
    const tagName = spaceIdx === -1 ? tagContent : tagContent.slice(0, spaceIdx);
    const attrStr = spaceIdx === -1 ? "" : tagContent.slice(spaceIdx + 1);

    const obj = {};
    const attrRegex = /([a-zA-Z_][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
    let attrMatch;
    while ((attrMatch = attrRegex.exec(attrStr)) !== null) {
      obj[`@${attrMatch[1]}`] = attrMatch[2] ?? attrMatch[3];
    }

    state.pos = tagClose + 1;
    if (selfClosing) {
      return { tagName, value: Object.keys(obj).length > 0 ? obj : null };
    }

    let textBuffer = "";
    let hasChildElements = false;

    while (state.pos < xml.length) {
      if (xml.startsWith(`</${tagName}`, state.pos)) {
        const endClose = xml.indexOf(">", state.pos);
        if (endClose === -1) throw new Error(`Unclosed closing tag </${tagName}>`);
        state.pos = endClose + 1;

        if (!hasChildElements) {
          const text = textBuffer.trim();
          const parsed = text === "" ? null : !isNaN(text) ? Number(text) : text;
          const value = Object.keys(obj).length > 0 ? { ...obj, "#text": parsed } : parsed;
          return { tagName, value };
        }

        // Collapse single-property array wrappers (e.g., <items><item>A</item><item>B</item></items> -> ["A", "B"])
        const keys = Object.keys(obj);
        if (keys.length === 1 && !keys[0].startsWith("@")) {
          const childKey = keys[0];
          const childVal = obj[childKey];

          // Repeated child elements (e.g., <items><item>A</item><item>B</item></items>)
          if (Array.isArray(childVal)) {
            return { tagName, value: childVal };
          }

          // Single child element where parent tag is plural of child tag (e.g., <items><item>A</item></items>)
          const isPluralWrapper =
            tagName.toLowerCase() === `${childKey.toLowerCase()}s` ||
            (childKey.toLowerCase().endsWith("y") &&
              tagName.toLowerCase() === `${childKey.slice(0, -1).toLowerCase()}ies`);

          if (isPluralWrapper) {
            return { tagName, value: [childVal] };
          }
        }


        return { tagName, value: obj };
      }

      if (xml[state.pos] === "<") {
        hasChildElements = true;
        const child = this.parseElement(xml, state);
        if (child) {
          if (obj[child.tagName] !== undefined) {
            if (!Array.isArray(obj[child.tagName])) {
              obj[child.tagName] = [obj[child.tagName]];
            }
            obj[child.tagName].push(child.value);
          } else {
            obj[child.tagName] = child.value;
          }
        }
      } else {
        textBuffer += xml[state.pos++];
      }
    }

    throw new Error(`Missing closing tag </${tagName}>`);
  }
};
