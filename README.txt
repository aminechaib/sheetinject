SheetInject — Chrome Extension (Manifest V3)
==========================================

How to install in Google Chrome:
1. Unzip this folder to a permanent location on your computer.
2. Open Chrome and navigate to: chrome://extensions
3. Turn ON "Developer mode" in the top-right corner.
4. Click "Load unpacked" and select this unzipped folder.
5. Pin "SheetInject" to your Chrome toolbar.

Smart Workflow on any website:
1. Open the website with your multi-row inputs, selects, or Radix/Shadcn comboboxes.
2. Click the SheetInject icon in the Chrome toolbar. It automatically detects how
   many rows exist on the page, detects the Add Row button, and lists table fields.
3. Click "Scan Page & Lists" to harvest the exact allowed dropdown/combobox options,
   then click "Export Template .xlsx" to download a ready-to-fill Excel file with
   matching headers and an "Allowed Options" sheet.
4. Fill your Excel rows (up to 600+ rows) and drop the .xlsx file into SheetInject.
5. Under "Multi-Row Repeater", choose Injection Mode: "Simulated (human-like)"
   for cascading React/comboboxes, "Direct + events (page state)" for fast direct
   injection that still updates React state, calculations, and virtualized rows,
   or "Raw value (no page state)" for DOM-only values (resets on scroll and skips
   calculations).
6. Click "Auto-Match" (or drag columns onto fields), click "Simple Test (3 Rows)",
   then click "Inject All" for instant native DOM injection.
