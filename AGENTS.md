# AGENTS.md

## What this is
Chrome extension (Manifest V3) that injects Excel rows into arbitrary web forms, including React/Radix/Shadcn comboboxes and virtualized tables. Plain JavaScript, no build step, no package manager, no tests, no lint config.

## Load / verify
- Load unpacked via `chrome://extensions` → Developer mode → Load unpacked (repo root).
- There is no `npm`/build/test/lint command. Verify manually in Chrome against a real target page.
- After editing `content.js` or `manifest.json`: reload the extension on `chrome://extensions`, then refresh the target tab. `content.js` guards with `window.__SHEETINJECT_LOADED__`, so an already-injected tab keeps running the old copy.
- Toolbar click toggles the panel. If the tab has no content script, `background.js` injects `lib/xlsx.mini.min.js` then `content.js`; `content.js` depends on the global `window.XLSX`.
- Use "Simple Test (3 Rows)" before "Inject All" when checking injection changes.

## Structure
- `content.js` (~2.2k lines) is effectively the whole app: IIFE, state machine, page scanner, selector engine, Excel parsing/export, DOM injection, and Shadow DOM UI.
- `background.js` only handles toolbar clicks / lazy injection. Message types are `SHEETINJECT_PING` and `SHEETINJECT_TOGGLE`; changing them requires updating both files. It also injects `lib/react_bridge.js` into `world: 'MAIN'` before `content.js`.
- `lib/react_bridge.js` is the MAIN-world React bridge. Content scripts run isolated and cannot see `__reactFiber$` expandos, so React mode talks to this bridge through `window.postMessage` (`siReactRequest` / `siReactResponse`). The bridge sets `data-si-react-bridge` on `<html>` as a cross-world marker. Every request carries `config` from `state.reactConfig`; the bridge merges it with its own defaults and uses it for the anchor selector, row/update/country/sous prop names, and sous option keys. Keep option objects inside the bridge: `onSelectSousPosition` needs the real option object from the page realm.
- `lib/xlsx.mini.min.js` is vendored SheetJS. Do not edit or replace casually.
- `README.txt` is end-user install/usage docs, not developer documentation.

## Implementation invariants
- UI lives in a Shadow DOM host with id `sheetinject-extension-host`; the hover highlight is a separate `#sheetinject-hover-ring`. Scans and drops must ignore elements inside the extension host.
- Mappings are either a normal CSS selector (scoped to a row) or an indexed template containing `{N}`, e.g. `[data-testid="input-quantity-{1}"]`. `{N}` is a base index; `resolveIndexedSelector` adds the target row offset at injection time. Keep this scheme when adding selector logic.
- Rows are handled two ways: indexed template selectors, or a repeating `rowSelector` plus `addRowSelector`. Virtualized tables use row height + scroll container probing; the add-row click loop waits up to 40 frames.
- Mappings and selectors persist per host in `chrome.storage.local` under `sheetinject_v2_profile_<hostname>`. React per-column modes and React hook config persist separately under `sheetinject_v2_react_<hostname>`. Schema changes need a new key or migration.
- Injection modes: `state.injectionMode` is `simulated` (default), `native`, `raw`, or `react`. Simulated uses focus/clicks/popup selection/readonly handling; native ("Direct + events") uses native setters plus `input`/`change`/`blur` but skips UI steps, so React state, calculations, and virtualized rows persist; raw writes `el.value`/`textContent` without events and will not survive React re-renders or virtualized scrolling; react ("React state (updateRow)") finds the table component through React Fiber and calls its `updateRow`/`fetchSousPositions`/`onSelectSousPosition` props directly instead of touching the DOM. Keep simulated as the default for cascading Radix fields.
- In React mode each Excel column has a per-column method stored in `state.reactColumnModes[header]`: `react` (default, use `updateRow`/React props) or `simulated` (use the mapped page control through `injectValue`, e.g. searchable comboboxes that must be typed into). The UI exposes this as a React/Human select in section 2; no automatic fallback between the two.
- React mode maps Excel headers to actual keys of `owner.rows[0]`, with header→field aliases editable in the UI (`field: label, label` per line; defaults include algex-style headers and Arabic page labels such as `البند التعريفي` → `categoryCode`, `البند التعريفي الفرعي` → `subCategoryCode`, `الوحدة` → `unit`, `بلد المنشأ` → `countryOfOrigin`, `البلد المورد` → `countryOfExport`). All page-specific hooks are editable in the UI: fiber anchor selector, rows/update/country/sous prop names, sous option keys, row-id field/header labels, and the category/sub/sous/country role field names. Country indexes are resolved at runtime from `owner[reactConfig.countriesProp]`, not hardcoded; conversion is chosen by value shape (single index vs comma-separated list), not by role field name, so either country field can hold either form. It takes the row id from an `id` column or the numbered DOM control, waits for the sous-positions prop before selecting, never calls `ensureRowExistsForIndex`/Add Row, verifies each write against the live row object, and logs warnings when the update did not apply. All React prop access goes through `reactBridge()`, never through DOM events or direct fiber reads in `content.js`.
- Simulated and native injection go through native prototype setters and dispatch `input`/`change` (plus clicks/focus for comboboxes and checkboxes) so React-controlled controls update. Do not replace them with direct `el.value = ...` writes.
- Readonly "anti-autofill" inputs are handled by temporarily removing `readonly` only in simulated mode (see `setNativeInputText`).
- Comboboxes are Radix/Shadcn-oriented: trigger detection uses `role="combobox"`, `aria-haspopup="listbox"`, `aria-haspopup="dialog"`; popup harvesting and selection support searchable single-select, multi-select, and cascading fields (category → subcategory → unit). Normalization is intentionally Arabic-aware (hamza/ta-marbuta and Arabic-Indic digits). Preserve `normalizeText`/`normalizeLoose` behavior.
- Injection order is dependency-based (`orderEntriesByDependency`), not spreadsheet order; cascading fields rely on this.

## Style
- Vanilla ES2020-style browser JS, 2-space indent, single quotes, semicolons; no modules/transpilation/bundler.
- Keep new code inside the existing IIFE and follow the section-comment layout rather than adding new files.
- No comments unless explaining a non-obvious browser/DOM workaround (existing section markers are fine).
