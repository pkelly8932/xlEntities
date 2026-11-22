# xlEntities - Structured Data & Entity Custom Functions for Excel

**xlEntities** is a functions-only Excel JavaScript add-in that lets you create, inspect, query, and flatten structured data (JSON, XML, Excel array-literal text, and worksheet ranges) directly inside Excel cells using native **Excel Entity Cards**.

No task pane, no custom ribbon buttons, and no external backend ... just lightweight custom functions under the `ENTITY` namespace.

---

## Features

* **Multi-Format Input Parsing:** Automatically detects or explicitly parses **JSON**, **XML**, **Excel array-literal strings** (`{1,2;3,4}`), and **2D cell ranges**.
* **Interactive Entity Cards:** Stores nested objects and arrays inside single cells as inspectable Excel Entity data types, using property keys and XML tags as entity labels.
* **Dot & Bracket Path Queries:** Extract deeply nested properties or array elements with `ENTITY.GET`.
* **Flexible Tabular Flattening:** Expand nested objects, uneven arrays of objects, and multi-dimensional arrays into spilled Excel tables with `ENTITY.FLATTEN`, controlling expansion depth (`levels`) and child-row padding (`pad_value`).
* **Nested Array Support:** Includes a `spill` option on `ENTITY.KEYS` for wrapping output into nested arrays inside wrapper formulas or Excel Beta/Insider builds.
* **Extensible Parser Architecture:** Drop new `NN_formatName.js` parser modules into `src/functions/input_formats/` to automatically register new input formats at build time.

---

## Custom Functions Reference

All functions are registered under the **`ENTITY`** namespace.

### 1. `=ENTITY.NEW(input, [name], [headers], [format])`
Parses raw structured text or a cell range into an Excel Entity card (or scalar value).

| Parameter | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| **`input`** | `Range \| String` | *(Required)* | A cell range, JSON string, XML string, or Excel array literal (`{...}`). |
| **`name`** | `String` | Auto | Text displayed next to the entity icon in the cell. Falls back to the root XML tag name or `"Array"`. |
| **`headers`** | `Boolean` | `FALSE` | When `TRUE` on a 2D array/range:<br>• **2 columns:** Uses column 1 as keys and column 2 as values.<br>• **>2 columns:** Uses row 1 as column headers and maps subsequent rows into an array of objects. |
| **`format`** | `String` | `"Auto"` | Force a specific parser (`"array"`, `"json"`, `"arrayText"`, `"xml"`) or use `"Auto"` detection. |

#### Examples
```excel
=ENTITY.NEW("{""id"":101,""user"":{""name"":""Alice"",""city"":""Boise""}}", "Account")
=ENTITY.NEW("<items><item>Keyboard</item><item>Mouse</item></items>", "Cart")
=ENTITY.NEW(A1:B4, "Config", TRUE)
=ENTITY.NEW(A1:D10, "Orders", TRUE)
=ENTITY.NEW("{""ok"":true}", "Status", FALSE, "json")
```

---

### 2. `=ENTITY.GET(source, [path])`
Extracts a scalar value or sub-entity from an Entity card or raw structured string using dot and bracket notation.

| Parameter | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| **`source`** | `Entity \| String` | *(Required)* | An Entity cell or raw JSON/XML/array-text string. |
| **`path`** | `String` | `""` | Dot/bracket path to the target property (e.g., `"user.city"`, `"orders[0].sku"`, or `"items.2"`). |

#### Examples
```excel
=ENTITY.GET(A1, "user.city")        --> "Boise"
=ENTITY.GET(A1, "orders[0].sku")    --> "KB-1"
=ENTITY.GET(A1, "user")             --> Sub-entity card labeled "user"
```

---

### 3. `=ENTITY.KEYS(source, [spill])`
Returns the property keys (for objects) or zero-based indices (for arrays) of an Entity.

| Parameter | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| **`source`** | `Entity \| String` | *(Required)* | An Entity cell or raw structured string. |
| **`spill`** | `Boolean` | `TRUE` | When `TRUE`, spills keys vertically down a column. When `FALSE`, wraps the result in an outer array level to create a nested array. |

> **Note on `spill = FALSE` (Nested Arrays):**
> Native nested array cell rendering is currently only available in **Microsoft 365 Beta / Insider** builds. In general-release versions of Excel, placing `=ENTITY.KEYS(A1, FALSE)` directly in a standalone cell will display `#VALUE!`, though it can still be used inside wrapper formulas (such as `=INDEX(ENTITY.KEYS(A1, FALSE), 1)` or `=ROWS(ENTITY.KEYS(A1, FALSE))`).

---

### 4. `=ENTITY.FLATTEN(source, [pad_value], [levels])`
Flattens an Entity, object, or array into a spilled tabular range up to a specified nesting depth.

| Parameter | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| **`source`** | `Entity \| String` | *(Required)* | An Entity cell, array, or raw structured string. |
| **`pad_value`** | `Any` | `null` | Value used to fill sibling columns when child array rows are expanded. When omitted (`null`), parent row values are copied down into child rows. |
| **`levels`** | `Number` | `1` | Number of nested levels to expand into columns/rows. Nested objects or arrays deeper than `levels` remain interactive Entity cards. |

#### Flattening Rules
* **1D Arrays:** Spill vertically in ascending index order with no index column.
* **Multi-Dimensional Arrays:** `levels = 1` spills outer elements vertically; `levels = 2` expands the second dimension across columns (padding ragged rows with `pad_value` or `""`).
* **Single Objects:** Keys form the header row (row 1) and values form row 2.
* **Arrays of Objects:** Uses the union of keys across all objects as column headers (`""` where a key is absent on a given object).
* **Nested Arrays on Objects:** Expands each array element into an additional child row, filling parent columns with `pad_value` (or copying parent values when `pad_value` is `null`).

#### Examples
```excel
=ENTITY.FLATTEN(A1)             --> Expands 1 level; copies parent values down child rows
=ENTITY.FLATTEN(A1, "---")      --> Expands 1 level; fills sibling cells in child rows with "---"
=ENTITY.FLATTEN(A1, "", 2)      --> Expands 2 levels deep; leaves level 3+ structures as Entity cards
=ENTITY.FLATTEN(F10, , 2)       --> Unpacks a 2D array entity into a spilled 2D worksheet range
```

---

## Project Structure

```text
├── manifest.xml                      # Functions-only Office Add-in manifest
├── webpack.config.js                 # Webpack & CustomFunctionsMetadataPlugin config
└── src/
    └── functions/
        ├── functions.html            # Invisible runtime host page
        ├── functions.js              # Excel custom function entry points (NEW, GET, KEYS, FLATTEN)
        ├── helpers.js                # Entity serialization, path resolution, and flattening engine
        └── input_formats/
            ├── index.js              # Dynamic Webpack loader (sorts ascending by NN_ prefix)
            ├── 10_array.js           # 2D range / array parser
            ├── 20_json.js            # JSON parser
            ├── 30_arrayText.js       # Excel array literal ({a,b;c,d}) parser
            └── 30_xml.js             # Pure-JS XML parser
```

### Adding Custom Input Formats
`src/functions/input_formats/index.js` uses `require.context` to automatically register any `.js` file in `input_formats/` whose name does not start with `_` or equal `index.js`:

1. Create a file named `NN_formatName.js` (e.g., `25_csv.js`), where `NN` is a 2-digit priority number (lower numbers evaluate first during `"Auto"` detection).
2. Export an object implementing `isFormat(input)` and `toJS(input)`:
   ```javascript
   export const csvFormat = {
     isFormat(input) {
       return typeof input === "string" && input.includes(",");
     },
     toJS(input) {
       // Return parsed JS array or object
     }
   };
   ```
3. Rebuild the project (`npm run build`). The format is automatically available to `"Auto"` detection and as `=ENTITY.NEW(input, , , "csv")`.

---

## Development & Deployment

### Local Testing
```bash
cd [path]                        # Navigate to the xlEntities directory
npm install yo generator-office  # Install Yeoman and the official generator-office package
npm run validate                 # Validate manifest.xml against Office XML schemas
npm start                        # Start local HTTPS dev server and sideload into Excel
```

### Self-Hosting (e.g., GitHub Pages)
1. Update these URLs in `manifest.xml` to your hosted HTTPS URL:
   - `Functions.Script.Url`
   - `Functions.Metadata.Url`
   - `Functions.Page.Url`
   - `DefaultSettings > SourceLocation`
3. Run `npm run build` and publish the contents of `dist/` to your static host.
4. Sideload `manifest.xml` into Excel via **Home > Add-ins > More Add-ins > Upload My Add-in** (or by using a shared network catalog).

### License:
This project is licensed under the PolyForm Perimeter License 1.0.1. You are free to use, modify, and self-host this add-in for personal or internal company use, provided you do not sell or offer it as a competing product.
