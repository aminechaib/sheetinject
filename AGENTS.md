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
- `background.js` only handles toolbar clicks / lazy injection. Message types are `SHEETINJECT_PING` and `SHEETINJECT_TOGGLE`; changing them requires updating both files.
- `lib/xlsx.mini.min.js` is vendored SheetJS. Do not edit or replace casually.
- `README.txt` is end-user install/usage docs, not developer documentation.

## Implementation invariants
- UI lives in a Shadow DOM host with id `sheetinject-extension-host`; the hover highlight is a separate `#sheetinject-hover-ring`. Scans and drops must ignore elements inside the extension host.
- Mappings are either a normal CSS selector (scoped to a row) or an indexed template containing `{N}`, e.g. `[data-testid="input-quantity-{1}"]`. `{N}` is a base index; `resolveIndexedSelector` adds the target row offset at injection time. Keep this scheme when adding selector logic.
- Rows are handled two ways: indexed template selectors, or a repeating `rowSelector` plus `addRowSelector`. Virtualized tables use row height + scroll container probing; the add-row click loop waits up to 40 frames.
- Mappings and selectors persist per host in `chrome.storage.local` under `sheetinject_v2_profile_<hostname>`. Schema changes need a new key or migration.
- Injection modes: `state.injectionMode` is `simulated` (default), `native`, or `raw`. Simulated uses focus/clicks/popup selection/readonly handling; native ("Direct + events") uses native setters plus `input`/`change`/`blur` but skips UI steps, so React state, calculations, and virtualized rows persist; raw writes `el.value`/`textContent` without events and will not survive React re-renders or virtualized scrolling. Keep simulated as the default for cascading Radix fields.
- Simulated and native injection go through native prototype setters and dispatch `input`/`change` (plus clicks/focus for comboboxes and checkboxes) so React-controlled controls update. Do not replace them with direct `el.value = ...` writes.
- Readonly "anti-autofill" inputs are handled by temporarily removing `readonly` only in simulated mode (see `setNativeInputText`).
- Comboboxes are Radix/Shadcn-oriented: trigger detection uses `role="combobox"`, `aria-haspopup="listbox"`, `aria-haspopup="dialog"`; popup harvesting and selection support searchable single-select, multi-select, and cascading fields (category → subcategory → unit). Normalization is intentionally Arabic-aware (hamza/ta-marbuta and Arabic-Indic digits). Preserve `normalizeText`/`normalizeLoose` behavior.
- Injection order is dependency-based (`orderEntriesByDependency`), not spreadsheet order; cascading fields rely on this.

## Style
- Vanilla ES2020-style browser JS, 2-space indent, single quotes, semicolons; no modules/transpilation/bundler.
- Keep new code inside the existing IIFE and follow the section-comment layout rather than adding new files.
- No comments unless explaining a non-obvious browser/DOM workaround (existing section markers are fine).
