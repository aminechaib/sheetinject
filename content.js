(() => {
  const HOST_ID = 'sheetinject-extension-host';
  const RING_ID = 'sheetinject-hover-ring';

  if (window.__SHEETINJECT_LOADED__) {
    const existing = document.getElementById(HOST_ID);
    if (existing) {
      existing.style.display = existing.style.display === 'none' ? 'block' : 'none';
    }
    return;
  }
  window.__SHEETINJECT_LOADED__ = true;

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg && msg.type === 'SHEETINJECT_PING') {
      sendResponse({ mounted: true });
      return true;
    }
    if (msg && msg.type === 'SHEETINJECT_TOGGLE') {
      const host = document.getElementById(HOST_ID);
      if (host) {
        host.style.display = host.style.display === 'none' ? 'block' : 'none';
      }
      sendResponse({ ok: true });
      return true;
    }
    return false;
  });

  const DEFAULT_REACT_CONFIG = {
    anchorSelector: 'textarea[data-testid^="textarea-description-"]',
    rowsProp: 'rows',
    updateRowProp: 'updateRow',
    countriesProp: 'countries',
    fetchSousProp: 'fetchSousPositions',
    sousPositionsProp: 'sousPositionsData',
    selectSousProp: 'onSelectSousPosition',
    sousPositionKey: 'sousPosition',
    subCategoryKey: 'subCategoryCode',
    idField: 'id',
    idHeaderLabels: 'id, rowid, row id',
    categoryField: 'categoryCode',
    subCategoryField: 'subCategoryCode',
    sousPositionField: 'sousPosition',
    countryExportField: 'countryOfExport',
    countryImportField: 'countryOfOrigin',
    aliasesText: [
      'categoryCode: input1, category, category code, البند التعريفي',
      'subCategoryCode: input2, sub category, subcategory, sub category code, subcategory code, sous position, sousposition, البند التعريفي الفرعي',
      'productDescription: textarea, description, product description, تسمية المنتج',
      'unit: unit, الوحدة',
      'countryOfOrigin: origin index, originindex, country of origin, بلد المنشأ',
      'countryOfExport: countries, country of export, البلد المورد',
      'quantity: value1, quantity, qty, الكمية المطلوب استرادها',
      'price: value2, price, unit price, سعر الوحدة بالدولار'
    ].join('\n')
  };

  // --- State ---
  const state = {
    workbook: null,
    fileName: '',
    sheetNames: [],
    activeSheet: '',
    headers: [],
    rows: [],
    rowSelector: '',
    addRowSelector: '',
    mappings: {}, // { [excelHeader]: selectorOrTemplate }
    reactColumnModes: {}, // { [excelHeader]: 'react' | 'simulated' }
    reactConfig: { ...DEFAULT_REACT_CONFIG },
    pace: 'balanced', // 'instant' | 'balanced' | 'safe'
    injectionMode: 'simulated',
    startRowOffset: 0,
    isRunning: false,
    isScanning: false,
    stopRequested: false,
    pickMode: null, // null | { kind: 'column', header } | { kind: 'addRow' }
    dragPayload: null,
    showLists: false,
    showFields: false,
    scan: {
      detectedRows: 0,
      rowHeight: 52,
      fields: [], // { label, kind, selector, templateSelector, options: [], disabled: boolean }
      addRowAutoDetected: false
    },
    logs: [],
    progress: { current: 0, total: 0, ok: 0, warn: 0, err: 0, elapsedMs: 0 }
  };

  const STORAGE_KEY = 'sheetinject_v2_profile_' + location.hostname;
  const REACT_STORAGE_KEY = 'sheetinject_v2_react_' + location.hostname;
  const COMBOBOX_SEL = '[role="combobox"], [aria-haspopup="listbox"], [aria-haspopup="dialog"]';
  const FORM_CONTROL_SEL = 'input:not([type="file"]):not([type="hidden"]):not([role="searchbox"]), select, textarea, [contenteditable="true"], ' + COMBOBOX_SEL;

  // --- Highlight Ring on Host Page ---
  const hoverRing = document.createElement('div');
  hoverRing.id = RING_ID;
  Object.assign(hoverRing.style, {
    position: 'fixed',
    zIndex: '2147483645',
    pointerEvents: 'none',
    border: '2px dashed #2563EB',
    backgroundColor: 'rgba(37, 99, 235, 0.10)',
    borderRadius: '6px',
    display: 'none',
    transition: 'all 60ms ease-out',
    boxSizing: 'border-box'
  });
  const hoverTag = document.createElement('div');
  Object.assign(hoverTag.style, {
    position: 'absolute',
    top: '-24px',
    left: '0',
    background: '#0F172A',
    color: '#FFFFFF',
    fontSize: '11px',
    fontFamily: 'system-ui, sans-serif',
    fontWeight: '600',
    padding: '2px 8px',
    borderRadius: '4px',
    whiteSpace: 'nowrap',
    boxShadow: '0 2px 6px rgba(0,0,0,0.2)'
  });
  hoverRing.appendChild(hoverTag);
  document.documentElement.appendChild(hoverRing);

  function hideRing() {
    hoverRing.style.display = 'none';
  }

  function showRingForElement(el, label) {
    if (!el || !(el instanceof Element)) {
      hideRing();
      return;
    }
    const rect = el.getBoundingClientRect();
    hoverRing.style.display = 'block';
    hoverRing.style.left = rect.left + 'px';
    hoverRing.style.top = rect.top + 'px';
    hoverRing.style.width = rect.width + 'px';
    hoverRing.style.height = rect.height + 'px';
    hoverTag.textContent = label;
  }

  // --- Host + Shadow DOM ---
  const host = document.createElement('div');
  host.id = HOST_ID;
  Object.assign(host.style, {
    position: 'fixed',
    top: '20px',
    right: '20px',
    zIndex: '2147483647',
    width: '410px',
    maxHeight: 'calc(100vh - 40px)',
    fontFamily: 'system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif'
  });
  const shadow = host.attachShadow({ mode: 'open' });
  document.documentElement.appendChild(host);

  const styleEl = document.createElement('style');
  styleEl.textContent = `
    * { box-sizing: border-box; }
    .si-panel {
      background: #FFFFFF;
      border: 1px solid #CBD5E1;
      border-radius: 12px;
      box-shadow: 0 20px 45px -12px rgba(15, 23, 42, 0.25);
      color: #0F172A;
      display: flex;
      flex-direction: column;
      max-height: calc(100vh - 40px);
      overflow: hidden;
      font-size: 13px;
      line-height: 1.4;
      direction: ltr;
      text-align: left;
    }
    .si-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 12px 14px;
      background: #0F172A;
      color: #FFFFFF;
      cursor: grab;
      user-select: none;
    }
    .si-header:active { cursor: grabbing; }
    .si-brand { font-weight: 700; font-size: 14px; letter-spacing: -0.01em; }
    .si-header-actions { display: flex; gap: 6px; }
    .si-icon-btn {
      background: rgba(255,255,255,0.12);
      border: none;
      color: #FFFFFF;
      border-radius: 6px;
      padding: 4px 8px;
      font-size: 11px;
      font-weight: 600;
      cursor: pointer;
    }
    .si-icon-btn:hover { background: rgba(255,255,255,0.22); }
    .si-body {
      padding: 14px;
      overflow-y: auto;
      display: flex;
      flex-direction: column;
      gap: 12px;
    }
    .si-scan-card {
      background: #F8FAFC;
      border: 1px solid #E2E8F0;
      border-radius: 8px;
      padding: 10px 11px;
    }
    .si-section-title {
      font-size: 11px;
      font-weight: 700;
      color: #475569;
      margin: 0 0 6px 0;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    .si-dropzone {
      border: 1.5px dashed #94A3B8;
      border-radius: 8px;
      padding: 14px 12px;
      text-align: center;
      background: #F8FAFC;
      cursor: pointer;
      transition: border-color 140ms, background 140ms;
    }
    .si-dropzone.active {
      border-color: #2563EB;
      background: #EFF6FF;
    }
    .si-btn {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 5px;
      border-radius: 7px;
      font-size: 12px;
      font-weight: 600;
      padding: 7px 10px;
      cursor: pointer;
      border: 1px solid #CBD5E1;
      background: #FFFFFF;
      color: #0F172A;
      transition: background 120ms, border-color 120ms;
    }
    .si-btn:hover:not(:disabled) { background: #F1F5F9; }
    .si-btn:disabled { opacity: 0.5; cursor: not-allowed; }
    .si-btn-primary {
      background: #2563EB;
      border-color: #2563EB;
      color: #FFFFFF;
    }
    .si-btn-primary:hover:not(:disabled) { background: #1D4ED8; }
    .si-btn-emerald {
      background: #059669;
      border-color: #059669;
      color: #FFFFFF;
    }
    .si-btn-emerald:hover:not(:disabled) { background: #047857; }
    .si-btn-soft-emerald {
      background: #ECFDF5;
      border-color: #A7F3D0;
      color: #047857;
    }
    .si-btn-soft-emerald:hover:not(:disabled) { background: #D1FAE5; }
    .si-btn-danger {
      background: #DC2626;
      border-color: #DC2626;
      color: #FFFFFF;
    }
    .si-col-list {
      display: flex;
      flex-direction: column;
      gap: 6px;
      max-height: 200px;
      overflow-y: auto;
    }
    .si-col-row {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
      padding: 6px 8px;
      border: 1px solid #E2E8F0;
      border-radius: 7px;
      background: #FFFFFF;
    }
    .si-col-chip {
      display: flex;
      flex-direction: column;
      flex: 1;
      min-width: 0;
      cursor: grab;
      padding: 2px 4px;
      border-radius: 4px;
      background: #F8FAFC;
      border: 1px solid #E2E8F0;
    }
    .si-col-chip:active { cursor: grabbing; }
    .si-col-name {
      font-weight: 600;
      font-size: 12px;
      color: #0F172A;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .si-col-sample {
      font-size: 11px;
      color: #64748B;
      font-family: ui-monospace, monospace;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .si-input {
      width: 100%;
      padding: 6px 8px;
      border: 1px solid #CBD5E1;
      border-radius: 6px;
      font-size: 12px;
      font-family: ui-monospace, monospace;
      color: #0F172A;
      background: #FFFFFF;
    }
    .si-select {
      padding: 5px 8px;
      border: 1px solid #CBD5E1;
      border-radius: 6px;
      font-size: 12px;
      color: #0F172A;
      background: #FFFFFF;
    }
    .si-progress-track {
      height: 8px;
      background: #E2E8F0;
      border-radius: 999px;
      overflow: hidden;
    }
    .si-progress-fill {
      height: 100%;
      background: #059669;
      width: 0%;
      transition: width 100ms linear;
    }
    .si-log-box {
      max-height: 125px;
      overflow-y: auto;
      border: 1px solid #E2E8F0;
      border-radius: 7px;
      background: #F8FAFC;
      padding: 6px 8px;
      font-family: ui-monospace, monospace;
      font-size: 11px;
      display: flex;
      flex-direction: column;
      gap: 3px;
    }
    .si-lists-box {
      margin-top: 8px;
      max-height: 155px;
      overflow-y: auto;
      border-top: 1px solid #E2E8F0;
      padding-top: 8px;
      display: flex;
      flex-direction: column;
      gap: 6px;
    }
    .si-grid-2 {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
    }
    .si-muted { color: #64748B; font-size: 11px; }
    .si-link-btn {
      background: none;
      border: none;
      padding: 0;
      color: #2563EB;
      font-size: 11px;
      font-weight: 600;
      cursor: pointer;
      text-align: left;
      margin-top: 6px;
    }
    .si-badge-ok { color: #047857; font-weight: 600; font-size: 11px; }
    .si-badge-warn { color: #B45309; font-weight: 600; font-size: 11px; }
    .si-badge-err { color: #B91C1C; font-weight: 600; font-size: 11px; }
  `;
  shadow.appendChild(styleEl);

  const panel = document.createElement('div');
  panel.className = 'si-panel';
  shadow.appendChild(panel);

  // --- Normalization & Matching Helpers (Arabic + Latin Accents + Digits) ---
  function normalizeText(str) {
    return String(str || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[\u064B-\u065F\u0670]/g, '')
      .replace(/[أإآ]/g, 'ا')
      .replace(/ة/g, 'ه')
      .replace(/ى/g, 'ي')
      .replace(/[\u0660-\u0669]/g, (d) => String(d.charCodeAt(0) - 0x0660))
      .replace(/[\u06F0-\u06F9]/g, (d) => String(d.charCodeAt(0) - 0x06f0))
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  }

  function normalizeLoose(str) {
    return normalizeText(str).replace(/[\-_()\[\],.;:/\\|]+/g, ' ').replace(/\s+/g, ' ').trim();
  }

  function escapeCss(str) {
    if (window.CSS && CSS.escape) return CSS.escape(str);
    return String(str).replace(/([ #;?%&,.+*~\':"!^$[\]()=>|/@])/g, '\\$1');
  }

  function escapeAttrVal(str) {
    return String(str).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  }

  function isElementVisible(el) {
    if (!el || host.contains(el)) return false;
    const rect = el.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0;
  }

  function findFormTarget(x, y) {
    const els = document.elementsFromPoint(x, y);
    for (const el of els) {
      if (host.contains(el) || el === host || el === hoverRing) continue;
      const direct = el.closest && el.closest(FORM_CONTROL_SEL);
      if (direct && !host.contains(direct)) return direct;
      const inner = el.querySelector && el.querySelector(FORM_CONTROL_SEL);
      if (inner && el.getBoundingClientRect().height < 64) return inner;
    }
    return null;
  }

  function findButtonTarget(x, y) {
    const els = document.elementsFromPoint(x, y);
    for (const el of els) {
      if (host.contains(el) || el === host || el === hoverRing) continue;
      const btn = el.closest('button, [role="button"], input[type="button"], input[type="submit"], a');
      if (btn) return btn;
    }
    return null;
  }

  function buildElementSelector(el) {
    if (el.dataset && el.dataset.testid) {
      return '[data-testid="' + escapeAttrVal(el.dataset.testid) + '"]';
    }
    if (el.dataset && el.dataset.field) {
      return '[data-field="' + escapeAttrVal(el.dataset.field) + '"]';
    }
    if (el.id && !/\d{2,}|[_-]\d+$/.test(el.id) && !el.id.includes(':')) {
      return '#' + escapeCss(el.id);
    }
    const parts = [];
    let cur = el;
    while (cur && cur.nodeType === 1 && cur !== document.body) {
      let seg = cur.tagName.toLowerCase();
      if (cur.id && !/\d{2,}|[_-]\d+$/.test(cur.id) && !cur.id.includes(':')) {
        seg = '#' + escapeCss(cur.id);
        parts.unshift(seg);
        break;
      }
      const parent = cur.parentElement;
      if (parent) {
        const siblings = Array.from(parent.children).filter((c) => c.tagName === cur.tagName);
        if (siblings.length > 1) {
          seg += ':nth-of-type(' + (siblings.indexOf(cur) + 1) + ')';
        }
      }
      parts.unshift(seg);
      cur = parent;
    }
    return parts.join(' > ');
  }

  function deriveIndexedTemplateSelector(el) {
    const attrs = ['data-testid', 'name', 'id'];
    for (const attr of attrs) {
      const val = el.getAttribute(attr);
      if (!val || val.includes(':')) continue;

      const trailing = /^(.*?)([-_.]?)(\d+)$/.exec(val);
      if (trailing) {
        const [, prefix, sep, numStr] = trailing;
        const n = parseInt(numStr, 10);
        // Allow indices up to 9999 so even picking Row 2 or Row 5 maps properly (normalized against base index)
        if (n >= 0 && prefix.length >= 2 && !/^f_$/i.test(prefix + sep)) {
          return '[' + attr + '="' + escapeAttrVal(prefix + sep) + '{' + n + '}"]';
        }
      }

      const embedded = /^(.*?)(\[|[-_.])(\d+)(\].*|[-_.].*)$/.exec(val);
      if (embedded) {
        const [, pre, openSep, numStr, post] = embedded;
        const n = parseInt(numStr, 10);
        if (n >= 0 && pre.length >= 2) {
          return '[' + attr + '="' + escapeAttrVal(pre + openSep) + '{' + n + '}' + escapeAttrVal(post) + '"]';
        }
      }
    }
    return '';
  }

  function detectRepeatingRowContext(targetEl) {
    const indexedTemplate = deriveIndexedTemplateSelector(targetEl);
    if (indexedTemplate) {
      return {
        rowSelector: '',
        cellSelector: indexedTemplate
      };
    }

    let cur = targetEl.parentElement;
    while (cur && cur !== document.body && cur !== document.documentElement) {
      const parent = cur.parentElement;
      if (parent) {
        const sameTagSiblings = Array.from(parent.children).filter(
          (c) => c.tagName === cur.tagName && c.querySelector(FORM_CONTROL_SEL)
        );
        const controlsInsideCur = cur.querySelectorAll(FORM_CONTROL_SEL).length;
        const isExplicitRow =
          cur.tagName === 'TR' ||
          cur.getAttribute('role') === 'row' ||
          cur.hasAttribute('data-row') ||
          cur.hasAttribute('data-row-index');

        const isLikelySingleGridRow =
          sameTagSiblings.length >= 1 &&
          controlsInsideCur >= 2 &&
          controlsInsideCur <= 30 &&
          parent.querySelectorAll(FORM_CONTROL_SEL).length === controlsInsideCur;

        if (isExplicitRow || sameTagSiblings.length >= 2 || isLikelySingleGridRow) {
          const parentSel = buildElementSelector(parent);
          const rowSel = parentSel + ' > ' + cur.tagName.toLowerCase();
          const relParts = [];
          let inner = targetEl;
          while (inner && inner !== cur) {
            let seg = inner.tagName.toLowerCase();
            if (inner.dataset && inner.dataset.field) {
              seg += '[data-field="' + escapeAttrVal(inner.dataset.field) + '"]';
              relParts.unshift(seg);
              break;
            }
            const tid = inner.getAttribute('data-testid');
            if (tid && !/\d+/.test(tid)) {
              seg += '[data-testid="' + escapeAttrVal(tid) + '"]';
              relParts.unshift(seg);
              break;
            }
            const nameAttr = inner.getAttribute('name');
            if (nameAttr && !/\d+/.test(nameAttr)) {
              seg += '[name="' + escapeAttrVal(nameAttr) + '"]';
              relParts.unshift(seg);
              break;
            }
            const p = inner.parentElement;
            if (p) {
              const sibs = Array.from(p.children).filter((c) => c.tagName === inner.tagName);
              if (sibs.length > 1) {
                seg += ':nth-of-type(' + (sibs.indexOf(inner) + 1) + ')';
              }
            }
            relParts.unshift(seg);
            inner = p;
          }
          return {
            rowSelector: rowSel,
            cellSelector: relParts.join(' > ')
          };
        }
      }
      cur = parent;
    }

    return {
      rowSelector: '',
      cellSelector: buildElementSelector(targetEl)
    };
  }

  function resolveIndexedSelector(template, domRowIndex) {
    return template.replace(/\{(\d+)\}/g, (_m, base) => String(Number(base) + domRowIndex));
  }

  function hasIndexPlaceholder(sel) {
    return /\{\d+\}/.test(sel);
  }

  // --- Virtualized Container & Smart Table Scanner ---
  function findRowContainerOfControl(el) {
    let cur = el.parentElement;
    while (cur && cur !== document.body) {
      const count = cur.querySelectorAll(FORM_CONTROL_SEL).length;
      if (count >= 3 && count <= 25) {
        return cur;
      }
      cur = cur.parentElement;
    }
    return null;
  }

  function findScrollableAncestor(el) {
    let cur = el ? el.parentElement : null;
    while (cur && cur !== document.body && cur !== document.documentElement) {
      const style = getComputedStyle(cur);
      if (
        (style.overflowY === 'auto' || style.overflowY === 'scroll') &&
        cur.clientHeight > 100 &&
        cur.scrollHeight >= cur.clientHeight
      ) {
        return cur;
      }
      cur = cur.parentElement;
    }
    return null;
  }

  function extractGridHeaderLabel(controlEl, rowEl) {
    if (rowEl) {
      // Find direct cell child of rowEl containing controlEl
      let cellEl = controlEl;
      while (cellEl && cellEl.parentElement !== rowEl) {
        cellEl = cellEl.parentElement;
      }
      if (cellEl && cellEl.parentElement === rowEl) {
        const colIndex = Array.from(rowEl.children).indexOf(cellEl);
        if (colIndex >= 0) {
          // Check sticky header sibling above virtual inner wrapper or table thead
          const scrollWrap = findScrollableAncestor(rowEl);
          const candidateContainers = [
            rowEl.parentElement && rowEl.parentElement.previousElementSibling,
            scrollWrap && scrollWrap.querySelector('.sticky'),
            rowEl.closest('table') && rowEl.closest('table').querySelector('thead tr')
          ].filter(Boolean);

          for (const hdr of candidateContainers) {
            if (hdr.children && hdr.children.length === rowEl.children.length) {
              const hdrText = (hdr.children[colIndex].textContent || '').replace(/\s+/g, ' ').trim();
              if (hdrText) return hdrText;
            }
          }
        }
      }
    }

    if (controlEl.id) {
      const lbl = document.querySelector('label[for="' + escapeAttrVal(controlEl.id) + '"]');
      if (lbl && lbl.textContent.trim()) return lbl.textContent.trim();
    }
    const ariaLabel = controlEl.getAttribute('aria-label');
    if (ariaLabel && ariaLabel.trim()) return ariaLabel.trim();

    const placeholder = controlEl.getAttribute('placeholder');
    if (placeholder && placeholder.trim()) return placeholder.trim();

    const tid = controlEl.getAttribute('data-testid');
    if (tid) {
      return tid.replace(/^(button|select|input|textarea|checkbox)-/i, '').replace(/[-_]\d+$/, '');
    }

    if (controlEl.matches(COMBOBOX_SEL)) {
      const txt = (controlEl.textContent || '').trim();
      if (txt && txt.length <= 40) return txt;
    }

    const nameAttr = controlEl.getAttribute('name');
    if (nameAttr) return nameAttr;

    return 'Field';
  }

  function autoDetectAddRowButton() {
    const btns = Array.from(document.querySelectorAll('button, [role="button"], a')).filter(
      (b) => isElementVisible(b) && !b.matches(COMBOBOX_SEL)
    );
    const match = btns.find((b) => {
      const tid = (b.getAttribute('data-testid') || '').toLowerCase();
      const txt = normalizeText(b.textContent);
      if (tid.includes('add-row') || tid.includes('add_row') || tid.includes('addrow')) return true;
      return (
        txt.includes('add row') ||
        txt.includes('add item') ||
        txt.includes('add line') ||
        txt.includes('new row') ||
        txt.includes('اضافه صف') ||
        txt.includes('اضافه سطر') ||
        txt.includes('صف جديد') ||
        txt === '+ add'
      );
    });
    return match ? buildElementSelector(match) : '';
  }

  function detectRowCountAndFirstRowControls() {
    const allControls = Array.from(document.querySelectorAll(FORM_CONTROL_SEL)).filter(isElementVisible);

    // 1. Group numbered row controls by their row index (e.g., button-category-1, textarea-description-1, input-quantity-1)
    const controlsByRowNum = new Map();
    for (const el of allControls) {
      const tid = el.getAttribute('data-testid') || el.getAttribute('id') || '';
      const m = /^(.*?)([-_.])(\d+)$/.exec(tid);
      if (m) {
        const idxNum = parseInt(m[3], 10);
        if (!controlsByRowNum.has(idxNum)) controlsByRowNum.set(idxNum, []);
        controlsByRowNum.get(idxNum).push(el);
      }
    }

    let firstRowEl = null;
    let rowOneControls = [];
    let detectedRows = 0;
    let rowHeight = 52;

    if (controlsByRowNum.size > 0) {
      const sortedIndices = Array.from(controlsByRowNum.keys()).sort((a, b) => a - b);
      const firstIdx = sortedIndices[0];
      rowOneControls = controlsByRowNum.get(firstIdx) || [];
      if (rowOneControls[0]) {
        firstRowEl = findRowContainerOfControl(rowOneControls[0]);
      }
      detectedRows = Math.max(sortedIndices.length, sortedIndices[sortedIndices.length - 1]);
    } else {
      // Fallback: check tbody tr or repeating row containers
      const tableRows = Array.from(document.querySelectorAll('tbody tr')).filter(
        (tr) => !host.contains(tr) && tr.querySelector(FORM_CONTROL_SEL)
      );
      if (tableRows.length > 0) {
        detectedRows = tableRows.length;
        firstRowEl = tableRows[0];
        rowOneControls = Array.from(firstRowEl.querySelectorAll(FORM_CONTROL_SEL)).filter(isElementVisible);
      } else {
        rowOneControls = allControls;
        detectedRows = allControls.length > 0 ? 1 : 0;
      }
    }

    // Check if inside a virtualized scroll container (e.g., @tanstack/react-virtual)
    if (firstRowEl) {
      const h = firstRowEl.getBoundingClientRect().height;
      if (h >= 28 && h <= 160) rowHeight = Math.round(h);
      const virtualInner = firstRowEl.parentElement;
      if (virtualInner && virtualInner.style && virtualInner.style.position === 'relative') {
        const totalVirtualPx = parseFloat(virtualInner.style.height || '0') || virtualInner.offsetHeight;
        if (totalVirtualPx > 0 && rowHeight > 0) {
          const virtualCount = Math.round(totalVirtualPx / rowHeight);
          if (virtualCount > detectedRows) detectedRows = virtualCount;
        }
      }
    }

    return { detectedRows, rowHeight, firstRowEl, rowOneControls };
  }

  async function scanPageStructure(harvestComboboxLists) {
    state.isScanning = true;
    render();

    try {
      const { detectedRows, rowHeight, firstRowEl, rowOneControls } = detectRowCountAndFirstRowControls();
      state.scan.rowHeight = rowHeight;

      const autoAddSel = autoDetectAddRowButton();
      if (autoAddSel && (!state.addRowSelector || !document.querySelector(state.addRowSelector))) {
        state.addRowSelector = autoAddSel;
        state.scan.addRowAutoDetected = true;
      }

      const scannedFields = [];
      for (let idx = 0; idx < rowOneControls.length; idx++) {
        const el = rowOneControls[idx];
        const ctx = detectRepeatingRowContext(el);

        let kind = 'input';
        let options = [];
        const isDisabled = Boolean(el.disabled || el.getAttribute('aria-disabled') === 'true');

        if (el.tagName === 'SELECT') {
          kind = 'select';
          options = Array.from(el.options)
            .filter((o) => o.value !== '')
            .map((o) => o.text.trim() || o.value);
        } else if (el.matches(COMBOBOX_SEL) && el.tagName !== 'INPUT') {
          kind = 'combobox';
          if (harvestComboboxLists && !isDisabled) {
            const popup = await openComboboxPopup(el);
            if (popup) {
              await waitMs(120);
              const items = extractPopupItems(popup);
              const rawTexts = items
                .map((i) => (i.textContent || '').replace(/\s+/g, ' ').trim())
                .filter(Boolean);
              options = Array.from(new Set(rawTexts));
              await closeComboboxPopup(el);
              await nextFrame();
              await waitMs(60);
            }
          } else {
            const prev = state.scan.fields.find((f) => f.templateSelector === ctx.cellSelector);
            if (prev) options = prev.options;
          }
        } else if (el.tagName === 'INPUT') {
          if (el.type === 'checkbox' || el.type === 'radio') kind = 'checkbox';
          else if (el.type === 'number' || el.getAttribute('inputmode') === 'decimal') kind = 'number';
          else if (el.type === 'date') kind = 'date';
        } else if (el.tagName === 'TEXTAREA') {
          kind = 'textarea';
        }

        scannedFields.push({
          label: extractGridHeaderLabel(el, firstRowEl),
          kind,
          selector: buildElementSelector(el),
          templateSelector: ctx.cellSelector,
          options,
          disabled: isDisabled
        });
      }

      state.scan.detectedRows = detectedRows;
      state.scan.fields = scannedFields;
      if (harvestComboboxLists) {
        const listCount = scannedFields.filter((f) => f.options.length > 0).length;
        addLog('Scanned ' + detectedRows + ' row(s), ' + scannedFields.length + ' table fields, ' + listCount + ' populated lists', 'ok');
      }
    } finally {
      state.isScanning = false;
      render();
    }
  }

  function exportWebsiteTemplateXlsx() {
    if (!window.XLSX || state.scan.fields.length === 0) return;

    const headers = state.scan.fields.map((f) => f.label);
    const exampleRow = {};
    for (const f of state.scan.fields) {
      if (f.options.length > 0) exampleRow[f.label] = f.options[0];
      else if (f.kind === 'number') exampleRow[f.label] = '10';
      else if (f.kind === 'date') exampleRow[f.label] = new Date().toISOString().slice(0, 10);
      else if (f.kind === 'checkbox') exampleRow[f.label] = 'No';
      else exampleRow[f.label] = '';
    }

    const wb = window.XLSX.utils.book_new();
    const dataWs = window.XLSX.utils.json_to_sheet([exampleRow], { header: headers });
    window.XLSX.utils.book_append_sheet(wb, dataWs, 'Data Entry');

    const listFields = state.scan.fields.filter((f) => f.options.length > 0);
    if (listFields.length > 0) {
      const maxLen = Math.max(...listFields.map((f) => f.options.length));
      const aoa = [listFields.map((f) => f.label)];
      for (let r = 0; r < maxLen; r++) {
        aoa.push(listFields.map((f) => f.options[r] || ''));
      }
      const optWs = window.XLSX.utils.aoa_to_sheet(aoa);
      window.XLSX.utils.book_append_sheet(wb, optWs, 'Allowed Options');
    }

    const hostSlug = (location.hostname || 'website').replace(/[^a-z0-9]+/gi, '_');
    window.XLSX.writeFile(wb, 'sheetinject_template_' + hostSlug + '.xlsx');
    addLog('Exported Excel template with ' + headers.length + ' columns', 'ok');
    render();
  }

  // --- Native DOM + Readonly Anti-Autofill + Radix Combobox (Searchable & Multi-Select) Engine ---
  const TRUTH = new Set(['true', '1', 'yes', 'y', 'checked', 'on', 'x', 'exempt', 'نعم', 'صح']);

  function triggerClickSequence(el) {
    const baseOpts = { bubbles: true, cancelable: true, view: window, button: 0, buttons: 1 };
    const ptrOpts = { ...baseOpts, pointerId: 1, pointerType: 'mouse', isPrimary: true };
    el.dispatchEvent(new PointerEvent('pointerover', ptrOpts));
    el.dispatchEvent(new MouseEvent('mouseover', baseOpts));
    el.dispatchEvent(new PointerEvent('pointerdown', ptrOpts));
    el.dispatchEvent(new MouseEvent('mousedown', baseOpts));
    el.dispatchEvent(new PointerEvent('pointerup', ptrOpts));
    el.dispatchEvent(new MouseEvent('mouseup', baseOpts));
    el.click();
  }

  async function setNativeInputText(el, val, opts) {
    const skipBlur = Boolean(opts && opts.skipBlur);
    const skipUi = Boolean(opts && opts.skipUi);

    // Handle anti-autofill readonly="" inputs (e.g. textarea-description-1, input-quantity-1, input-price-1)
    if (!skipUi) {
      el.focus();
      if (el.hasAttribute('readonly')) {
        triggerClickSequence(el);
        await nextFrame();
        if (el.hasAttribute('readonly')) {
          el.removeAttribute('readonly');
        }
      }
    }

    const proto = el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
    if (el._valueTracker) el._valueTracker.setValue(el.value === val ? val + '__' : el.value);

    let finalVal = val;
    if (el.tagName === 'INPUT' && (el.type === 'number' || el.getAttribute('inputmode') === 'decimal')) {
      const normalizedDigits = String(val)
        .replace(/[\u0660-\u0669]/g, (d) => String(d.charCodeAt(0) - 0x0660))
        .replace(/[\u06F0-\u06F9]/g, (d) => String(d.charCodeAt(0) - 0x06f0));
      finalVal = normalizedDigits.replace(/[^0-9.-]/g, '');
    }

    if (setter) setter.call(el, finalVal);
    else el.value = finalVal;

    el.dispatchEvent(new Event('input', { bubbles: true, cancelable: true }));
    el.dispatchEvent(new Event('change', { bubbles: true, cancelable: true }));
    if (!skipBlur && !skipUi) {
      el.dispatchEvent(new Event('blur', { bubbles: true, cancelable: true }));
    }
  }

  function dispatchValueEvents(el) {
    el.dispatchEvent(new Event('input', { bubbles: true, cancelable: true }));
    el.dispatchEvent(new Event('change', { bubbles: true, cancelable: true }));
  }

  function setDirectComboboxValue(el, val, raw) {
    if ('value' in el) {
      if (raw) {
        el.value = val;
      } else {
        const setter = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(el), 'value')?.set;
        if (setter) setter.call(el, val);
        else el.value = val;
      }
    }
    el.textContent = val;
    if (!raw) dispatchValueEvents(el);
    return { ok: true };
  }

  function setRawValue(el, rawVal) {
    const val = rawVal == null ? '' : String(rawVal);

    if (el.isContentEditable) {
      el.textContent = val;
      return { ok: true };
    }

    if (el.tagName === 'SELECT') {
      const opts = Array.from(el.options);
      const targetNorm = normalizeText(val);
      const match =
        opts.find((o) => o.value === val) ||
        opts.find((o) => o.text.trim() === val) ||
        opts.find((o) => normalizeText(o.value) === targetNorm || normalizeText(o.text) === targetNorm);
      if (!match) return { ok: false, warn: 'No option matched "' + val + '"' };
      el.value = match.value;
      return { ok: true };
    }

    if (el.tagName === 'INPUT' && (el.type === 'checkbox' || el.type === 'radio')) {
      el.checked = TRUTH.has(normalizeText(val));
      return { ok: true };
    }

    if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
      el.value = val;
      return { ok: true };
    }

    if (el.matches(COMBOBOX_SEL)) {
      return setDirectComboboxValue(el, val, true);
    }

    return { ok: false, warn: 'Unsupported element' };
  }

  function clickAddRow(el) {
    if (state.injectionMode === 'simulated') triggerClickSequence(el);
    else el.click();
  }

  function nextFrame() {
    return new Promise((r) => requestAnimationFrame(() => r()));
  }

  function waitMs(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }

  function extractPopupItems(popup) {
    if (!popup) return [];
    const popupRect = popup.getBoundingClientRect();

    // 1. Semantic ARIA / cmdk / Radix collection items (do NOT return early if < 3 items, as a chip/tab may share [data-value])
    const semantic = Array.from(
      popup.querySelectorAll(
        '[role="option"], [role="menuitemcheckbox"], [cmdk-item], [data-radix-collection-item], [data-value], ul li button, ul li, [role="listbox"] button, [role="listbox"] > div'
      )
    ).filter((el) => {
      if (el.hasAttribute('disabled') || el.getAttribute('aria-disabled') === 'true') return false;
      const r = el.getBoundingClientRect();
      return r.width > 0 && r.height > 0 && r.height <= 96;
    });

    // 2. Generic Radix Dialog list items (plain <div> / <label> / <button> / <li> rows, including multi-select checkbox rows)
    const rawCandidates = Array.from(popup.querySelectorAll('div, label, button, li, a')).filter((el) => {
      if (el === popup || host.contains(el)) return false;
      if (el.hasAttribute('disabled') || el.getAttribute('aria-disabled') === 'true') return false;
      if (el.getAttribute('aria-label') === 'Close') return false;
      // Only exclude containers wrapping text/search inputs — ALLOW checkbox/radio rows!
      if (el.querySelector('input:not([type="checkbox"]):not([type="radio"]):not([type="hidden"]), textarea, select, [role="searchbox"]')) {
        return false;
      }
      const txt = (el.textContent || '').replace(/\s+/g, ' ').trim();
      if (!txt || txt.length > 220) return false;
      const rect = el.getBoundingClientRect();
      if (rect.width < 24 || rect.height < 16 || rect.height > 72) return false;
      const cs = getComputedStyle(el);
      if (cs.overflowY === 'auto' || cs.overflowY === 'scroll') return false;
      // Ignore small inline-flex selected-country chip badges inside multi-select chip bars (width < 45% of popup when popup is wide)
      if (
        (cs.display === 'inline-flex' || cs.display === 'inline-block') &&
        popupRect.width > 220 &&
        rect.width < popupRect.width * 0.45 &&
        el.tagName !== 'BUTTON'
      ) {
        return false;
      }
      return true;
    });

    // Prune parent containers that wrap >= 2 inner item rows (keeps single row <div> even if it has a text + code sub-div!)
    const candSet = new Set(rawCandidates);
    const genericRows = rawCandidates.filter((el) => {
      let containedCount = 0;
      let wrapsEqualText = false;
      const elNorm = normalizeText(el.textContent);
      for (const other of candSet) {
        if (other !== el && el.contains(other)) {
          containedCount++;
          if (normalizeText(other.textContent) === elNorm) {
            wrapsEqualText = true;
          }
          if (containedCount >= 2) return false;
        }
      }
      // If a parent wrapper has the exact same text as a single inner candidate and parent is not a button/label/role-option, prefer the inner element unless parent is cursor-pointer and child is not
      if (wrapsEqualText && containedCount === 1) {
        const cs = getComputedStyle(el);
        const childSame = rawCandidates.find((other) => other !== el && el.contains(other));
        if (childSame && (getComputedStyle(childSame).cursor === 'pointer' || cs.cursor !== 'pointer')) {
          return false;
        }
      }
      return true;
    });

    // Prefer semantic items if there are at least 2 of them and generic didn't find a bigger list; otherwise merge keeping order
    if (semantic.length >= 2 && semantic.length >= genericRows.length) {
      return semantic;
    }

    const merged = [];
    const seen = new Set();
    for (const el of semantic) {
      if (!seen.has(el)) {
        seen.add(el);
        merged.push(el);
      }
    }
    for (const el of genericRows) {
      if (!seen.has(el)) {
        seen.add(el);
        merged.push(el);
      }
    }
    return merged;
  }

  function getItemText(el) {
    return (el.textContent || '').replace(/\s+/g, ' ').trim();
  }

  function normalizeCompact(str) {
    return normalizeText(str).replace(/[\s\-._/\\|()]+/g, '');
  }

  function findMatchingPopupItem(items, targetNorm) {
    if (!items || items.length === 0 || !targetNorm) return null;
    const targetLoose = normalizeLoose(targetNorm);
    const targetCompact = normalizeCompact(targetNorm);

    const exactVal = items.find((o) => normalizeText(o.getAttribute('data-value')) === targetNorm);
    if (exactVal) return exactVal;

    const exactText = items.find((o) => normalizeText(getItemText(o)) === targetNorm);
    if (exactText) return exactText;

    const looseExact = items.find((o) => {
      const lTxt = normalizeLoose(getItemText(o));
      const lDv = normalizeLoose(o.getAttribute('data-value'));
      return lTxt === targetLoose || (lDv && lDv === targetLoose);
    });
    if (looseExact) return looseExact;

    const startsWithMatch = items.find((o) => {
      const txt = normalizeText(getItemText(o));
      const dv = normalizeText(o.getAttribute('data-value'));
      const cTxt = normalizeCompact(getItemText(o));
      const cDv = normalizeCompact(o.getAttribute('data-value'));
      return (
        txt.startsWith(targetNorm) ||
        (dv && dv.startsWith(targetNorm)) ||
        (targetCompact && (cTxt.startsWith(targetCompact) || (cDv && cDv.startsWith(targetCompact))))
      );
    });
    if (startsWithMatch) return startsWithMatch;

    const includesMatches = items.filter((o) => {
      const txt = normalizeLoose(getItemText(o));
      const dv = normalizeLoose(o.getAttribute('data-value'));
      const cTxt = normalizeCompact(getItemText(o));
      const cDv = normalizeCompact(o.getAttribute('data-value'));
      return (
        (txt && txt.includes(targetLoose)) ||
        (dv && dv.includes(targetLoose)) ||
        (targetCompact && ((cTxt && cTxt.includes(targetCompact)) || (cDv && cDv.includes(targetCompact))))
      );
    });
    if (includesMatches.length > 0) {
      includesMatches.sort((a, b) => getItemText(a).length - getItemText(b).length);
      return includesMatches[0];
    }

    return null;
  }

  function getOpenPopupForTrigger(triggerEl) {
    const controlsId = triggerEl.getAttribute('aria-controls');
    if (controlsId) {
      const byId = document.getElementById(controlsId);
      if (byId && byId.getAttribute('data-state') !== 'closed' && isElementVisible(byId)) {
        return byId;
      }
      // Do not fall back to unrelated popups while aria-controls hasn't mounted or is still closed
      return null;
    }
    const openPopup = document.querySelector(
      '[role="dialog"][data-state="open"], [role="listbox"][data-state="open"], [data-radix-popper-content-wrapper] [role="dialog"][data-state="open"]'
    );
    if (openPopup && !host.contains(openPopup) && isElementVisible(openPopup)) {
      return openPopup;
    }
    return null;
  }

  async function openComboboxPopup(triggerEl) {
    const alreadyOpen =
      triggerEl.getAttribute('aria-expanded') === 'true' ||
      triggerEl.getAttribute('data-state') === 'open';
    if (alreadyOpen) {
      const existing = getOpenPopupForTrigger(triggerEl);
      if (existing) return existing;
    }

    triggerClickSequence(triggerEl);

    for (let attempt = 0; attempt < 35; attempt++) {
      await nextFrame();
      const popup = getOpenPopupForTrigger(triggerEl);
      if (popup) return popup;
    }
    return null;
  }

  async function closeComboboxPopup(triggerEl) {
    // Allow 2 frames for Radix auto-close after an option click before checking if popup is still open
    await nextFrame();
    await nextFrame();
    const popup = getOpenPopupForTrigger(triggerEl);
    const expanded = triggerEl.getAttribute('aria-expanded') === 'true' || triggerEl.getAttribute('data-state') === 'open';
    if (!popup && !expanded) return;

    const active = document.activeElement && !host.contains(document.activeElement)
      ? document.activeElement
      : document;
    active.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', keyCode: 27, bubbles: true }));
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', keyCode: 27, bubbles: true }));
    await waitMs(45);

    if (getOpenPopupForTrigger(triggerEl) || triggerEl.getAttribute('aria-expanded') === 'true' || triggerEl.getAttribute('data-state') === 'open') {
      triggerClickSequence(triggerEl);
      await waitMs(50);
    }
  }

  function isPopupItemAlreadyChecked(itemEl) {
    if (!itemEl) return false;
    if (itemEl.getAttribute('aria-checked') === 'true' || itemEl.getAttribute('data-state') === 'checked') {
      return true;
    }
    const nativeCb = itemEl.querySelector('input[type="checkbox"]');
    if (nativeCb && nativeCb.checked) return true;
    const roleCb = itemEl.querySelector('[role="checkbox"]');
    if (roleCb && (roleCb.getAttribute('aria-checked') === 'true' || roleCb.getAttribute('data-state') === 'checked')) {
      return true;
    }
    return false;
  }

  function findScrollableInsidePopup(popup) {
    const all = [popup, ...Array.from(popup.querySelectorAll('div, ul'))];
    return all.find((el) => {
      const cs = getComputedStyle(el);
      return (cs.overflowY === 'auto' || cs.overflowY === 'scroll') && el.scrollHeight > el.clientHeight + 20;
    }) || null;
  }

  async function selectSingleComboboxValue(triggerEl, token, isMultiMode) {
    const targetNorm = normalizeText(token);
    if (!targetNorm) return { ok: true };

    let popup = await openComboboxPopup(triggerEl);
    if (!popup) {
      return { ok: false, warn: 'Combobox dialog did not open' };
    }

    // Wait 2 frames for popup animation & initial search input render
    await nextFrame();
    await nextFrame();

    const findSearchInput = (p) =>
      p.querySelector('input:not([type="hidden"]):not([type="checkbox"]):not([type="radio"])');

    let searchInput = findSearchInput(popup);
    let typedToken = '';

    if (searchInput) {
      await setNativeInputText(searchInput, token, { skipBlur: true });
      typedToken = token;
      await waitMs(150);
    }

    let match = null;
    let lastItems = [];
    let retriedCompactSearch = false;
    const maxPolls = 50; // Up to ~3.0s for slow cascading / HS subcategory API lookups

    for (let poll = 0; poll < maxPolls; poll++) {
      popup = getOpenPopupForTrigger(triggerEl) || popup;

      // Re-detect searchInput if popup rendered loading spinner first, or restore if async re-render cleared it
      const liveInput = findSearchInput(popup);
      if (liveInput && (!searchInput || (typedToken && normalizeText(liveInput.value) === ''))) {
        searchInput = liveInput;
        await setNativeInputText(searchInput, typedToken || token, { skipBlur: true });
        typedToken = typedToken || token;
        await waitMs(150);
      }

      lastItems = extractPopupItems(popup);
      match = findMatchingPopupItem(lastItems, targetNorm);
      if (match) break;

      // If formatted code like "3926.90.9700" yielded no match after ~600ms, retry search with compact digits "3926909700"
      if (searchInput && !retriedCompactSearch && poll >= 10) {
        const compactCode = String(token).replace(/[\s\-._/]+/g, '');
        if (/^\d{3,}$/.test(compactCode) && compactCode !== token) {
          retriedCompactSearch = true;
          typedToken = compactCode;
          await setNativeInputText(searchInput, compactCode, { skipBlur: true });
          await waitMs(180);
          continue;
        }
      }

      if (!searchInput && lastItems.length > 0 && poll >= 3) break;
      await waitMs(60);
    }

    // If not matched yet and the popup has a scrollable list container (e.g. virtualized country list), scroll to find it
    if (!match) {
      const scroller = findScrollableInsidePopup(popup);
      if (scroller) {
        const origTop = scroller.scrollTop;
        const step = Math.max(120, Math.round(scroller.clientHeight * 0.75));
        for (let top = 0; top <= scroller.scrollHeight; top += step) {
          scroller.scrollTop = top;
          await nextFrame();
          lastItems = extractPopupItems(popup);
          match = findMatchingPopupItem(lastItems, targetNorm);
          if (match) break;
        }
        if (!match) scroller.scrollTop = origTop;
      }
    }

    if (!match) {
      const samplePreview = lastItems
        .slice(0, 3)
        .map((i) => getItemText(i))
        .filter(Boolean)
        .join(', ');
      return {
        ok: false,
        warn:
          'No option matched "' +
          token +
          '"' +
          (lastItems.length > 0 ? ' (saw ' + lastItems.length + ' items: ' + samplePreview + ')' : ' (0 items detected)')
      };
    }

    // In multi-select dialogs (e.g. Country of Origin), don't uncheck an already-selected country!
    if (isMultiMode && isPopupItemAlreadyChecked(match)) {
      return { ok: true };
    }

    match.scrollIntoView({ block: 'nearest' });
    await nextFrame();
    await nextFrame();

    // Guard against stale/detached node if React re-rendered the filtered list while scrolling
    if (!match.isConnected) {
      popup = getOpenPopupForTrigger(triggerEl) || popup;
      const freshItems = extractPopupItems(popup);
      match = findMatchingPopupItem(freshItems, targetNorm) || match;
    }

    const beforeTriggerNorm = normalizeText(triggerEl.textContent);
    const innerCb = match.querySelector('input[type="checkbox"], [role="checkbox"]');
    const beforeChecked = innerCb
      ? (innerCb.checked || innerCb.getAttribute('aria-checked') === 'true' || innerCb.getAttribute('data-state') === 'checked')
      : false;

    triggerClickSequence(match);
    await waitMs(85);

    if (innerCb) {
      const afterChecked =
        innerCb.checked || innerCb.getAttribute('aria-checked') === 'true' || innerCb.getAttribute('data-state') === 'checked';
      if (afterChecked === beforeChecked) {
        triggerClickSequence(innerCb);
        await waitMs(75);
      }
    } else if (!isMultiMode) {
      // Verify single-select actually registered (popup closed OR trigger label updated); if still stuck open with unchanged label, retry click cleanly
      const stillOpen = Boolean(getOpenPopupForTrigger(triggerEl));
      const unchanged = normalizeText(triggerEl.textContent) === beforeTriggerNorm;
      if (stillOpen && unchanged) {
        const livePopup = getOpenPopupForTrigger(triggerEl) || popup;
        const freshMatch = findMatchingPopupItem(extractPopupItems(livePopup), targetNorm);
        if (freshMatch) {
          const clickTarget = freshMatch.querySelector('span, button') || freshMatch;
          triggerClickSequence(clickTarget);
          await waitMs(95);
        }
      }
    }

    await waitMs(50);
    return { ok: true };
  }

  async function injectCombobox(triggerEl, rawVal) {
    const trimmed = String(rawVal || '').trim();
    if (!trimmed) return { ok: true };
    const targetNorm = normalizeText(trimmed);

    if (normalizeText(triggerEl.textContent) === targetNorm) {
      return { ok: true };
    }

    const tidLower = (triggerEl.getAttribute('data-testid') || '').toLowerCase();
    const isCascadingChild =
      tidLower.includes('subcategory') ||
      tidLower.includes('sub_category') ||
      tidLower.includes('unit');

    // Brief settle pause on cascading child controls so parent state transitions & loading/disabled flags kick in before probing
    if (isCascadingChild) {
      await waitMs(120);
    }

    // Wait up to 4.0s if temporarily disabled while a parent cascading combobox (e.g. Category -> Subcategory -> Unit) resolves
    for (let w = 0; w < 160; w++) {
      if (!triggerEl.disabled && triggerEl.getAttribute('aria-disabled') !== 'true') break;
      await waitMs(25);
    }

    if (triggerEl.disabled || triggerEl.getAttribute('aria-disabled') === 'true') {
      return { ok: false, warn: 'Combobox is disabled (select parent field first)' };
    }

    const isLikelyMulti = tidLower.includes('countryoforigin') || tidLower.includes('origin');

    const hasMultiSep = /[,;|\n]/.test(trimmed);
    if (!hasMultiSep) {
      const beforeNorm = normalizeText(triggerEl.textContent);
      let res = await selectSingleComboboxValue(triggerEl, trimmed, isLikelyMulti);
      await closeComboboxPopup(triggerEl);

      // Automatic single retry for cascading child controls or missed popup clicks
      if (!res.ok || (!isLikelyMulti && normalizeText(triggerEl.textContent) === beforeNorm && !beforeNorm.includes(targetNorm))) {
        await waitMs(260);
        // Re-wait in case child briefly re-disabled while fetching
        for (let w = 0; w < 60; w++) {
          if (!triggerEl.disabled && triggerEl.getAttribute('aria-disabled') !== 'true') break;
          await waitMs(25);
        }
        res = await selectSingleComboboxValue(triggerEl, trimmed, isLikelyMulti);
        await closeComboboxPopup(triggerEl);
      }

      await waitMs(40);
      return res;
    }

    // Try exact whole value first (in case an option naturally contains a comma), otherwise split for multi-select (e.g. Country of Origin)
    const tokens = trimmed
      .split(/[,;|\n]+/)
      .map((s) => s.trim())
      .filter(Boolean);

    const firstWholeTry = await selectSingleComboboxValue(triggerEl, trimmed, isLikelyMulti);
    if (firstWholeTry.ok) {
      await closeComboboxPopup(triggerEl);
      await waitMs(40);
      return firstWholeTry;
    }

    const missed = [];
    for (const tok of tokens) {
      const r = await selectSingleComboboxValue(triggerEl, tok, true);
      if (!r.ok) missed.push(r.warn || tok);
    }

    await closeComboboxPopup(triggerEl);
    await waitMs(40);

    if (missed.length > 0) {
      return { ok: false, warn: missed.join('; ') };
    }
    return { ok: true };
  }

  async function injectValue(el, rawVal) {
    const val = rawVal == null ? '' : String(rawVal);
    const mode = state.injectionMode;

    if (mode === 'raw') {
      return setRawValue(el, val);
    }

    const noUi = mode === 'native';

    if (
      el.matches(COMBOBOX_SEL) &&
      el.tagName !== 'INPUT' &&
      el.tagName !== 'SELECT'
    ) {
      if (noUi) return setDirectComboboxValue(el, val, false);
      return injectCombobox(el, val);
    }

    if (el.tagName === 'SELECT') {
      const trimmed = val.trim();
      const targetNorm = normalizeText(trimmed);
      const opts = Array.from(el.options);
      let match =
        opts.find((o) => o.value === trimmed) ||
        opts.find((o) => o.text.trim() === trimmed) ||
        opts.find((o) => normalizeText(o.value) === targetNorm || normalizeText(o.text) === targetNorm) ||
        (targetNorm ? opts.find((o) => normalizeText(o.text).startsWith(targetNorm) || normalizeText(o.text).includes(targetNorm)) : null);

      if (!match) {
        return { ok: false, warn: 'No option matched "' + trimmed + '"' };
      }

      const setter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value')?.set;
      if (el._valueTracker) el._valueTracker.setValue('__si_reset__');
      if (setter) setter.call(el, match.value);
      else el.value = match.value;

      el.dispatchEvent(new Event('input', { bubbles: true, cancelable: true }));
      el.dispatchEvent(new Event('change', { bubbles: true, cancelable: true }));
      return { ok: true };
    }

    if (el.tagName === 'INPUT' && (el.type === 'checkbox' || el.type === 'radio')) {
      const want = TRUTH.has(normalizeText(val));
      if (el.checked !== want) {
        const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'checked')?.set;
        if (el._valueTracker) el._valueTracker.setValue(String(!want));
        if (setter) setter.call(el, want);
        else el.checked = want;
        if (!noUi) el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
        el.dispatchEvent(new Event('input', { bubbles: true, cancelable: true }));
        el.dispatchEvent(new Event('change', { bubbles: true, cancelable: true }));
      }
      return { ok: true };
    }

    if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
      await setNativeInputText(el, val, noUi ? { skipUi: true } : undefined);
      return { ok: true };
    }

    if (el.isContentEditable) {
      el.textContent = val;
      el.dispatchEvent(new Event('input', { bubbles: true, cancelable: true }));
      if (noUi) {
        el.dispatchEvent(new Event('change', { bubbles: true, cancelable: true }));
        el.dispatchEvent(new Event('blur', { bubbles: true, cancelable: true }));
      }
      return { ok: true };
    }

    return { ok: false, warn: 'Unsupported element' };
  }

  function reactAliases() {
    const text = String(state.reactConfig.aliasesText || '');
    const parsed = text.split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => {
        const idx = line.indexOf(':');
        if (idx < 0) return null;
        const field = line.slice(0, idx).trim();
        if (!field) return null;
        const labels = line.slice(idx + 1).split(',').map((v) => v.trim()).filter(Boolean);
        return labels.length ? { field, labels: labels.map(normalizeLoose) } : null;
      })
      .filter(Boolean);

    return parsed.length ? parsed : DEFAULT_REACT_CONFIG.aliasesText.split(/\r?\n/).map((line) => {
      const idx = line.indexOf(':');
      return {
        field: line.slice(0, idx).trim(),
        labels: line.slice(idx + 1).split(',').map((v) => normalizeLoose(v.trim())).filter(Boolean)
      };
    });
  }

  function reactFieldName(header) {
    const raw = String(header || '').trim();
    const key = normalizeLoose(raw);
    const hit = reactAliases().find((entry) => entry.labels.includes(key));
    return hit ? hit.field : raw;
  }

  let reactRequestSeq = 0;

  function reactBridge(action, payload, timeoutMs) {
    return new Promise((resolve, reject) => {
      if (!document.documentElement.hasAttribute('data-si-react-bridge')) {
        reject(new Error('React bridge not injected — reload extension and refresh tab'));
        return;
      }

      const id = 'si_' + Date.now() + '_' + ++reactRequestSeq;
      const timer = setTimeout(() => {
        window.removeEventListener('message', onResponse);
        reject(new Error('React bridge timeout (' + action + ')'));
      }, timeoutMs || 12000);

      function onResponse(event) {
        if (event.source !== window) return;
        const detail = event.data;
        if (!detail || detail.siReactResponse !== true || detail.id !== id) return;
        clearTimeout(timer);
        window.removeEventListener('message', onResponse);
        resolve(detail.result || { ok: false, warn: 'empty bridge response' });
      }

      window.addEventListener('message', onResponse);
      window.postMessage({ siReactRequest: true, id, action, config: { ...state.reactConfig }, ...payload }, '*');
    });
  }

  function reactDomRowId(targetDomIndex) {
    const selectors = Object.values(state.mappings).filter((sel) => sel && hasIndexPlaceholder(sel));
    const candidates = selectors.map((sel) => resolveIndexedSelector(sel, targetDomIndex));
    candidates.push('[data-testid$="-' + (targetDomIndex + 1) + '"]');

    for (const sel of candidates) {
      const el = document.querySelector(sel);
      if (!el) continue;
      for (const attr of ['data-testid', 'id', 'name']) {
        const val = el.getAttribute(attr);
        const m = val && /(?:^|[-_.])(\d+)$/.exec(val);
        if (m) return Number(m[1]);
      }
    }
    return null;
  }

  function reactFieldKey(header, fields) {
    const alias = reactFieldName(header);
    const list = Array.isArray(fields) ? fields : [];
    if (list.includes(alias)) return alias;
    const hNorm = normalizeLoose(header).replace(/\s/g, '');
    const match = list.find((k) => normalizeLoose(k).replace(/\s/g, '') === hNorm);
    return match || alias;
  }

  async function reactUpdateField(rowId, field, value) {
    const res = await reactBridge('update', { rowId, field, value });
    await waitMs(10);
    return Boolean(res.ok && (res.value === null || res.value === String(value ?? '')));
  }

  function isSimulatedColumn(header) {
    return state.reactColumnModes[header] === 'simulated';
  }

  async function injectMappedColumn(header, value, targetDomIndex, rowScope) {
    const rawSel = state.mappings[header];
    if (!rawSel) return { ok: false, warn: 'no mapped control' };

    let el = null;
    if (hasIndexPlaceholder(rawSel)) {
      el = document.querySelector(resolveIndexedSelector(rawSel, targetDomIndex));
    } else if (rowScope) {
      el = rowScope.querySelector(rawSel);
    }
    if (!el) return { ok: false, warn: 'mapped control not found' };

    return injectValue(el, String(value ?? ''));
  }

  async function injectReactRow(rowData, targetDomIndex, ownerInfo, rowScope) {
    const warnings = [];
    let written = 0;
    const cfg = state.reactConfig;

    const idHeader = Object.keys(rowData).find((h) =>
      cfg.idHeaderLabels.split(',').map((v) => normalizeText(v.trim())).includes(normalizeText(h)));
    let rowId = idHeader ? Number(rowData[idHeader]) : reactDomRowId(targetDomIndex);
    if (!Number.isFinite(rowId)) rowId = targetDomIndex + 1;

    const entries = Object.entries(rowData).filter(([h]) => String(h || '').trim() && h !== idHeader);
    const fieldOf = (h) => reactFieldKey(h, ownerInfo.fields);
    const categoryEntry = entries.find(([h]) => fieldOf(h) === cfg.categoryField);
    const subEntry = entries.find(([h]) => [cfg.subCategoryField, cfg.sousPositionField].includes(fieldOf(h)));

    if (categoryEntry) {
      const res = isSimulatedColumn(categoryEntry[0])
        ? await injectMappedColumn(categoryEntry[0], categoryEntry[1], targetDomIndex, rowScope)
        : { ok: await reactUpdateField(rowId, cfg.categoryField, categoryEntry[1]) };
      if (res.ok) written++;
      else warnings.push(categoryEntry[0] + ': ' + (res.warn || 'updateRow did not apply'));
    }

    if (subEntry && isSimulatedColumn(subEntry[0])) {
      const res = await injectMappedColumn(subEntry[0], subEntry[1], targetDomIndex, rowScope);
      if (res.ok) written++;
      else warnings.push(subEntry[0] + ': ' + (res.warn || 'not filled'));
    } else if (subEntry && ownerInfo.fetch && ownerInfo.select) {
      const fetched = await reactBridge('fetchSous', { rowId, category: categoryEntry ? categoryEntry[1] : '' });
      if (!fetched.ok) warnings.push(subEntry[0] + ': ' + fetched.warn);

      const deadline = Date.now() + 10000;
      let found = false;
      while (Date.now() < deadline) {
        const res = await reactBridge('options', { rowId }, 3000);
        if (res.ok && (res.options || []).some((item) =>
          String(item[state.reactConfig.sousPositionKey] ?? item[state.reactConfig.subCategoryKey] ?? '') === String(subEntry[1]))) {
          found = true;
          break;
        }
        await waitMs(100);
      }

      if (found) {
        const res = await reactBridge('selectSous', { rowId, sousPosition: subEntry[1] });
        if (res.ok) {
          written++;
          await waitMs(10);
        } else {
          warnings.push(subEntry[0] + ': ' + res.warn);
        }
      } else {
        warnings.push(subEntry[0] + ': sous-position not found');
      }
    }

    for (const [header, rawVal] of entries) {
      const field = fieldOf(header);
      if (field === cfg.categoryField || field === cfg.subCategoryField || field === cfg.sousPositionField) continue;

      let val = rawVal == null ? '' : String(rawVal);
      if ([cfg.countryExportField, cfg.countryImportField].includes(field)) {
        const parts = val.split(',').map((v) => v.trim()).filter(Boolean);
        if (parts.length > 1 && parts.every((v) => /^\d+$/.test(v))) {
          const res = await reactBridge('countryCodes', { codes: val });
          if (res.ok) val = res.codes;
        } else if (/^\d+$/.test(val)) {
          const res = await reactBridge('countryCode', { code: val });
          if (res.ok) val = res.code;
        }
      }

      const res = isSimulatedColumn(header)
        ? await injectMappedColumn(header, val, targetDomIndex, rowScope)
        : { ok: await reactUpdateField(rowId, field, val) };
      if (res.ok) written++;
      else warnings.push(header + ': ' + (res.warn || 'updateRow did not apply'));
    }

    return { written, warnings };
  }

  function addLog(text, level) {
    state.logs.unshift({ text, level: level || 'ok' });
    if (state.logs.length > 60) state.logs.length = 60;
  }

  async function ensureRowExistsForIndex(targetDomIndex, mappedEntries, scrollContainer) {
    let addBtn = state.addRowSelector ? document.querySelector(state.addRowSelector) : null;
    if (!addBtn) {
      const autoSel = autoDetectAddRowButton();
      if (autoSel) {
        state.addRowSelector = autoSel;
        addBtn = document.querySelector(autoSel);
      }
    }

    // Strategy A: Indexed template selector (works for virtualized React tables & numbered data-testid)
    const indexedEntry = mappedEntries.find(([, sel]) => hasIndexPlaceholder(sel));
    if (indexedEntry) {
      const probeSel = resolveIndexedSelector(indexedEntry[1], targetDomIndex);

      if (document.querySelector(probeSel)) return { ok: true };

      // If virtualized scroll container exists, scroll to target position first to see if row is already in React state
      if (scrollContainer) {
        const rowH = state.scan.rowHeight || 52;
        scrollContainer.scrollTop = Math.max(0, targetDomIndex * rowH - 120);
        await nextFrame();
        await nextFrame();
        if (document.querySelector(probeSel)) return { ok: true };
      }

      if (!addBtn) {
        return { ok: false, reason: 'Missing DOM row #' + (targetDomIndex + 1) + ' & no Add Row button mapped' };
      }

      clickAddRow(addBtn);

      for (let tries = 0; tries < 40; tries++) {
        await nextFrame();
        if (scrollContainer) {
          const rowH = state.scan.rowHeight || 52;
          scrollContainer.scrollTop = Math.max(0, targetDomIndex * rowH - 120);
        }
        if (document.querySelector(probeSel)) return { ok: true };
      }
      return { ok: false, reason: 'Row #' + (targetDomIndex + 1) + ' (' + probeSel + ') did not mount' };
    }

    // Strategy B: Repeating row container selector
    if (state.rowSelector) {
      let domRows = document.querySelectorAll(state.rowSelector);
      if (targetDomIndex >= domRows.length) {
        if (!addBtn) return { ok: false, reason: 'Missing DOM row & no Add Row button mapped' };
        const missing = targetDomIndex - domRows.length + 1;
        for (let k = 0; k < missing; k++) clickAddRow(addBtn);
        for (let tries = 0; tries < 40; tries++) {
          await nextFrame();
          if (document.querySelectorAll(state.rowSelector).length > targetDomIndex) {
            return { ok: true };
          }
        }
        return { ok: false, reason: 'DOM row container did not appear' };
      }
    }

    return { ok: true };
  }

  function orderEntriesByDependency(entries) {
    const weight = (sel) => {
      const s = String(sel).toLowerCase();
      if (s.includes('category') && !s.includes('sub')) return 1;
      if (s.includes('subcategory') || s.includes('sub_category')) return 2;
      if (s.includes('description')) return 3;
      if (s.includes('unit')) return 4;
      if (s.includes('countryoforigin') || s.includes('origin')) return 5;
      if (s.includes('countryofexport') || s.includes('export')) return 6;
      if (s.includes('quantity') || s.includes('qty')) return 7;
      if (s.includes('price')) return 8;
      return 10;
    };
    return [...entries].sort((a, b) => weight(a[1]) - weight(b[1]));
  }

  async function runInjection(limitCount) {
    const isReactMode = state.injectionMode === 'react';
    const rawMapped = Object.entries(state.mappings).filter(([, sel]) => Boolean(sel));
    if ((!isReactMode && rawMapped.length === 0) || state.rows.length === 0) return;
    const mappedEntries = orderEntriesByDependency(rawMapped);

    state.isRunning = true;
    state.stopRequested = false;
    state.logs = [];

    const total = limitCount ? Math.min(state.rows.length, limitCount) : state.rows.length;
    state.progress = { current: 0, total, ok: 0, warn: 0, err: 0, elapsedMs: 0 };
    render();

    let reactOwnerInfo = null;
    if (isReactMode) {
      try {
        const res = await reactBridge('inspect');
        if (!res.ok) throw new Error(res.warn);
        reactOwnerInfo = res;
        addLog('React owner: ' + res.rows + ' rows [' + (res.fields || []).join(', ') + ']', 'ok');
      } catch (err) {
        addLog('React owner: ' + err.message, 'err');
      }
    }

    // Locate virtual scroll container once if present
    const firstProbe = mappedEntries.find(([, sel]) => hasIndexPlaceholder(sel));
    let scrollContainer = null;
    if (firstProbe) {
      const sampleEl = document.querySelector(resolveIndexedSelector(firstProbe[1], state.startRowOffset));
      const rowEl = sampleEl ? findRowContainerOfControl(sampleEl) : null;
      if (rowEl) scrollContainer = findScrollableAncestor(rowEl);
    }

    const startedAt = performance.now();
    const batchSize = state.pace === 'instant' ? 15 : state.pace === 'balanced' ? 5 : 1;
    const delay = state.pace === 'instant' ? 0 : state.pace === 'balanced' ? 10 : 35;

    for (let i = 0; i < total; i++) {
      if (state.stopRequested) {
        addLog('Stopped by user at row ' + (i + 1), 'warn');
        break;
      }

      const targetDomIndex = state.startRowOffset + i;
      if (!isReactMode) {
        const ensure = await ensureRowExistsForIndex(targetDomIndex, mappedEntries, scrollContainer);
        if (!ensure.ok) {
          state.progress.err++;
          addLog('Row ' + (i + 1) + ': ' + ensure.reason, 'err');
          state.progress.current = i + 1;
          renderProgressOnly();
          continue;
        }
      }

      const rowScope = state.rowSelector
        ? document.querySelectorAll(state.rowSelector)[targetDomIndex] || document
        : document;

      const rowData = state.rows[i];
      let rowWarnings = [];
      let written = 0;

      if (isReactMode) {
        try {
          if (!reactOwnerInfo) throw new Error('React owner not available');
          const res = await injectReactRow(rowData, targetDomIndex, reactOwnerInfo, rowScope);
          written = res.written;
          rowWarnings = res.warnings;
        } catch (err) {
          rowWarnings.push(err.message);
        }
      } else {
        for (const [header, rawSel] of mappedEntries) {
          let el = null;
          if (hasIndexPlaceholder(rawSel)) {
            const resolvedSel = resolveIndexedSelector(rawSel, targetDomIndex);
            el = document.querySelector(resolvedSel);
          } else {
            el = rowScope.querySelector(rawSel);
          }

          if (!el) {
            rowWarnings.push(header + ' missing');
            continue;
          }
          const res = await injectValue(el, rowData[header] ?? '');
          if (res.ok) written++;
          else if (res.warn) rowWarnings.push(header + ': ' + res.warn);
        }
      }

      if (rowWarnings.length > 0) {
        state.progress.warn++;
        addLog('Row ' + (i + 1) + ': ' + rowWarnings.join('; '), 'warn');
      } else {
        state.progress.ok++;
        if (i < 5 || (i + 1) % 50 === 0 || i === total - 1) {
          addLog('Row ' + (i + 1) + ': ' + written + ' fields injected', 'ok');
        }
      }

      state.progress.current = i + 1;
      state.progress.elapsedMs = Math.round(performance.now() - startedAt);

      if ((i + 1) % batchSize === 0 || i === total - 1) {
        renderProgressOnly();
        if (delay > 0) await waitMs(delay);
        else await nextFrame();
      }
    }

    state.progress.elapsedMs = Math.round(performance.now() - startedAt);
    state.isRunning = false;
    await scanPageStructure(false);
  }

  // --- Spreadsheet Parsing ---
  function handleFile(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = new Uint8Array(ev.target.result);
        const wb = window.XLSX.read(data, { type: 'array', cellDates: true });
        state.workbook = wb;
        state.fileName = file.name;
        state.sheetNames = wb.SheetNames || [];
        loadSheet(state.sheetNames[0]);
      } catch (err) {
        addLog('Parse failed: ' + err.message, 'err');
        render();
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function loadSheet(sheetName) {
    if (!state.workbook || !sheetName) return;
    state.activeSheet = sheetName;
    const ws = state.workbook.Sheets[sheetName];
    const aoa = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
    if (!aoa || aoa.length === 0) {
      state.headers = [];
      state.rows = [];
      render();
      return;
    }
    state.headers = aoa[0].map((h, i) => String(h || '').trim() || 'Column ' + (i + 1));
    const parsedRows = [];
    for (let r = 1; r < aoa.length; r++) {
      const row = aoa[r];
      if (!Array.isArray(row) || !row.some((v) => String(v || '').trim() !== '')) continue;
      const obj = {};
      state.headers.forEach((h, c) => {
        obj[h] = row[c] == null ? '' : String(row[c]).trim();
      });
      parsedRows.push(obj);
    }
    state.rows = parsedRows;
    loadSavedMappings();
    render();
  }

  function saveMappings() {
    try {
      chrome.storage.local.set({
        [STORAGE_KEY]: {
          rowSelector: state.rowSelector,
          addRowSelector: state.addRowSelector,
          mappings: state.mappings
        }
      });
    } catch (_e) {}
  }

  function saveReactConfig() {
    try {
      chrome.storage.local.set({
        [REACT_STORAGE_KEY]: {
          reactColumnModes: state.reactColumnModes,
          reactConfig: state.reactConfig
        }
      });
    } catch (_e) {}
  }

  function loadSavedMappings() {
    try {
      chrome.storage.local.get([STORAGE_KEY], (res) => {
        const saved = res && res[STORAGE_KEY];
        if (saved) {
          if (saved.rowSelector) state.rowSelector = saved.rowSelector;
          if (saved.addRowSelector) state.addRowSelector = saved.addRowSelector;
          if (saved.mappings) state.mappings = { ...saved.mappings };
          render();
        }
      });
      chrome.storage.local.get([REACT_STORAGE_KEY], (res) => {
        const saved = res && res[REACT_STORAGE_KEY];
        if (saved) {
          if (saved.reactColumnModes) state.reactColumnModes = { ...saved.reactColumnModes };
          if (saved.reactConfig) {
            const cfg = { ...DEFAULT_REACT_CONFIG, ...saved.reactConfig };
            cfg.aliasesText = String(cfg.aliasesText || '').replace(
              'countryOfExport: origin index, originindex, country of export, بلد المنشأ\ncountryOfOrigin: countries, country of origin, البلد المورد',
              'countryOfOrigin: origin index, originindex, country of origin, بلد المنشأ\ncountryOfExport: countries, country of export, البلد المورد'
            );
            state.reactConfig = cfg;
          }
          render();
        }
      });
    } catch (_e) {}
  }

  function autoMatchOnPage() {
    let matchedCount = 0;
    for (const header of state.headers) {
      const hNorm = normalizeText(header);

      const scannedMatch = state.scan.fields.find((f) => {
        const lNorm = normalizeText(f.label);
        return lNorm === hNorm || lNorm.includes(hNorm) || hNorm.includes(lNorm);
      });
      if (scannedMatch && scannedMatch.templateSelector) {
        state.mappings[header] = scannedMatch.templateSelector;
        matchedCount++;
        continue;
      }

      const controls = Array.from(document.querySelectorAll(FORM_CONTROL_SEL)).filter(isElementVisible);
      let bestEl = null;
      for (const el of controls) {
        const fieldName = normalizeText(
          (el.getAttribute('data-field') || '') + ' ' +
          (el.getAttribute('data-testid') || '') + ' ' +
          (el.getAttribute('name') || '') + ' ' +
          (el.getAttribute('placeholder') || '') + ' ' +
          (el.getAttribute('aria-label') || '') + ' ' +
          (el.matches(COMBOBOX_SEL) ? el.textContent || '' : '')
        );

        if (fieldName && (fieldName === hNorm || fieldName.includes(hNorm) || hNorm.includes(fieldName))) {
          bestEl = el;
          break;
        }
      }
      if (bestEl) {
        const ctx = detectRepeatingRowContext(bestEl);
        if (ctx.rowSelector && !state.rowSelector) state.rowSelector = ctx.rowSelector;
        state.mappings[header] = ctx.cellSelector;
        matchedCount++;
      }
    }
    saveMappings();
    addLog('Auto-matched ' + matchedCount + ' fields on page', 'ok');
    render();
  }

  // --- Global Drag-and-Drop / Click-to-Pick onto Host Webpage ---
  document.addEventListener('dragover', (e) => {
    if (!state.dragPayload) return;
    e.preventDefault();
    const target =
      state.dragPayload.kind === 'addRow'
        ? findButtonTarget(e.clientX, e.clientY)
        : findFormTarget(e.clientX, e.clientY);
    if (target) {
      const lbl =
        state.dragPayload.kind === 'addRow'
          ? 'Drop to set Add Row button'
          : 'Drop to map "' + state.dragPayload.header + '"';
      showRingForElement(target, lbl);
    } else {
      hideRing();
    }
  }, true);

  document.addEventListener('drop', (e) => {
    if (!state.dragPayload) return;
    const payload = state.dragPayload;
    state.dragPayload = null;
    hideRing();

    const target =
      payload.kind === 'addRow'
        ? findButtonTarget(e.clientX, e.clientY)
        : findFormTarget(e.clientX, e.clientY);
    if (!target) return;

    e.preventDefault();
    e.stopPropagation();

    if (payload.kind === 'addRow') {
      state.addRowSelector = buildElementSelector(target);
    } else {
      const ctx = detectRepeatingRowContext(target);
      if (ctx.rowSelector) state.rowSelector = ctx.rowSelector;
      state.mappings[payload.header] = ctx.cellSelector;
    }
    saveMappings();
    render();
  }, true);

  document.addEventListener('mousemove', (e) => {
    if (!state.pickMode) return;
    const target =
      state.pickMode.kind === 'addRow'
        ? findButtonTarget(e.clientX, e.clientY)
        : findFormTarget(e.clientX, e.clientY);
    if (target) {
      const lbl =
        state.pickMode.kind === 'addRow'
          ? 'Click to bind Add Row button'
          : 'Click to bind "' + state.pickMode.header + '"';
      showRingForElement(target, lbl);
    } else {
      hideRing();
    }
  }, true);

  document.addEventListener('click', (e) => {
    if (!state.pickMode) return;
    if (host.contains(e.target)) return;
    e.preventDefault();
    e.stopPropagation();
    const mode = state.pickMode;
    state.pickMode = null;
    hideRing();

    const target =
      mode.kind === 'addRow'
        ? findButtonTarget(e.clientX, e.clientY)
        : findFormTarget(e.clientX, e.clientY);
    if (target) {
      if (mode.kind === 'addRow') {
        state.addRowSelector = buildElementSelector(target);
      } else {
        const ctx = detectRepeatingRowContext(target);
        if (ctx.rowSelector) state.rowSelector = ctx.rowSelector;
        state.mappings[mode.header] = ctx.cellSelector;
      }
      saveMappings();
    }
    render();
  }, true);

  // --- Panel Dragging ---
  let dragWin = null;
  window.addEventListener('mousemove', (e) => {
    if (!dragWin) return;
    host.style.left = Math.max(8, Math.min(window.innerWidth - 418, e.clientX - dragWin.ox)) + 'px';
    host.style.top = Math.max(8, Math.min(window.innerHeight - 60, e.clientY - dragWin.oy)) + 'px';
    host.style.right = 'auto';
  });
  window.addEventListener('mouseup', () => { dragWin = null; });

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function renderProgressOnly() {
    const fill = shadow.getElementById('si-prog-fill');
    const txt = shadow.getElementById('si-prog-text');
    if (fill && state.progress.total > 0) {
      fill.style.width = Math.round((state.progress.current / state.progress.total) * 100) + '%';
    }
    if (txt) {
      txt.textContent =
        state.progress.current + ' / ' + state.progress.total +
        ' rows · ' + state.progress.elapsedMs + 'ms';
    }
  }

  function render() {
    const mappedCount = Object.values(state.mappings).filter(Boolean).length;
    const pct = state.progress.total > 0 ? Math.round((state.progress.current / state.progress.total) * 100) : 0;
    const listFields = state.scan.fields.filter((f) => f.options.length > 0);
    const comboboxCount = state.scan.fields.filter((f) => f.kind === 'combobox' || f.kind === 'select').length;
    const rowsToCreate = state.rows.length > 0 ? Math.max(0, state.rows.length - state.scan.detectedRows) : 0;

    panel.innerHTML = `
      <div class="si-header" id="si-drag-handle">
        <span class="si-brand">SheetInject</span>
        <div class="si-header-actions">
          <button class="si-icon-btn" id="si-close-btn">Hide</button>
        </div>
      </div>
      <div class="si-body">
        <div class="si-scan-card">
          <div class="si-section-title">
            <span>Page Detector &amp; Select Lists</span>
            <span class="si-muted">${state.scan.detectedRows} row(s) · ${state.scan.fields.length} fields</span>
          </div>
          <div class="si-muted" style="margin-bottom:8px;">
            ${comboboxCount} select/combobox lists · Add Row ${state.addRowSelector ? 'ready' : 'not found yet'}
          </div>
          <div class="si-grid-2">
            <button class="si-btn" id="si-scan-btn" ${state.isRunning || state.isScanning ? 'disabled' : ''}>${state.isScanning ? 'Scanning Lists...' : 'Scan Page & Lists'}</button>
            <button class="si-btn si-btn-soft-emerald" id="si-export-tpl-btn" ${state.scan.fields.length === 0 || state.isRunning ? 'disabled' : ''}>Export Template .xlsx</button>
          </div>
          ${state.scan.fields.length > 0 ? `
            <div style="display:flex;justify-content:space-between;gap:8px;">
              <button class="si-link-btn" id="si-toggle-fields-btn">${state.showFields ? 'Hide' : 'View'} detected fields (${state.scan.fields.length})</button>
              ${listFields.length > 0 ? '<button class="si-link-btn" id="si-toggle-lists-btn">' + (state.showLists ? 'Hide' : 'View') + ' lists (' + listFields.length + ')</button>' : ''}
            </div>
            ${state.showFields ? `
              <div class="si-lists-box">
                ${state.scan.fields.map((f, idx) => '<div><strong>' + (idx + 1) + '. ' + escapeHtml(f.label) + '</strong> <span class="si-muted">(' + f.kind + (f.disabled ? ', disabled until parent picked' : '') + ')</span></div>').join('')}
              </div>
            ` : ''}
            ${state.showLists && listFields.length > 0 ? `
              <div class="si-lists-box">
                ${listFields.map((lf) => '<div><strong>' + escapeHtml(lf.label) + '</strong> <span class="si-muted">(' + lf.options.length + ')</span><div class="si-muted" style="font-family:ui-monospace,monospace;">' + escapeHtml(lf.options.slice(0, 15).join(', ')) + (lf.options.length > 15 ? ' ...' : '') + '</div></div>').join('')}
              </div>
            ` : ''}
          ` : ''}
        </div>

        <div>
          <div class="si-section-title">
            <span>1. Excel / CSV Spreadsheet</span>
            ${state.rows.length ? '<span class="si-muted">' + state.rows.length + ' rows</span>' : ''}
          </div>
          <div class="si-dropzone" id="si-file-drop">
            <input type="file" id="si-file-input" accept=".xlsx,.xls,.csv,.ods" style="display:none" />
            ${state.fileName
              ? '<strong>' + escapeHtml(state.fileName) + '</strong><div class="si-muted">' + state.headers.length + ' columns · Drop another file to replace</div>'
              : '<strong>Drop .xlsx or .csv file here</strong><div class="si-muted">or click to browse from computer</div>'}
          </div>
          ${state.rows.length > 0 ? `
            <div class="si-muted" style="margin-top:5px;">
              ${rowsToCreate > 0 ? 'Page has ' + state.scan.detectedRows + ' row(s) — ' + rowsToCreate + ' more will auto-spawn.' : 'Page has ' + state.scan.detectedRows + ' rows (covers all ' + state.rows.length + ' Excel rows).'}
            </div>
          ` : ''}
          ${state.sheetNames.length > 1 ? `
            <div style="margin-top:6px;display:flex;align-items:center;gap:6px;">
              <span class="si-muted">Sheet:</span>
              <select class="si-select" id="si-sheet-sel" style="flex:1">
                ${state.sheetNames.map((s) => '<option value="' + escapeHtml(s) + '"' + (s === state.activeSheet ? ' selected' : '') + '>' + escapeHtml(s) + '</option>').join('')}
              </select>
            </div>
          ` : ''}
        </div>

        ${state.headers.length > 0 ? `
          <div>
            <div class="si-section-title">
              <span>2. Drag Columns onto Inputs / Comboboxes</span>
              <button class="si-btn" id="si-automatch-btn" style="padding:3px 8px;font-size:11px;">Auto-Match</button>
            </div>
            ${state.injectionMode === 'react' ? '<div class="si-muted" style="margin-bottom:6px;">Column method: React = updateRow · Human = type/click into page control</div>' : ''}
            <div class="si-col-list">
              ${state.headers.map((h) => {
                const sample = state.rows[0] ? state.rows[0][h] : '';
                const mapped = state.mappings[h];
                const colMode = state.reactColumnModes[h] || 'react';
                const isPicking = state.pickMode && state.pickMode.kind === 'column' && state.pickMode.header === h;
                return `
                  <div class="si-col-row">
                    <div class="si-col-chip" draggable="true" data-col-drag="${escapeHtml(h)}">
                      <span class="si-col-name">${escapeHtml(h)}</span>
                      <span class="si-col-sample">${mapped ? 'Mapped: ' + escapeHtml(mapped) : 'e.g. ' + escapeHtml(sample || 'empty')}</span>
                    </div>
                    ${state.injectionMode === 'react' ? `
                      <select class="si-select" data-col-mode="${escapeHtml(h)}" title="Injection method for this column" style="padding:3px 4px;font-size:10px;">
                        <option value="react"${colMode === 'react' ? ' selected' : ''}>React</option>
                        <option value="simulated"${colMode === 'simulated' ? ' selected' : ''}>Human</option>
                      </select>
                    ` : ''}
                    <button class="si-btn" data-col-pick="${escapeHtml(h)}" style="padding:5px 8px;font-size:11px;">${isPicking ? 'Click Page...' : 'Pick'}</button>
                    ${mapped ? '<button class="si-btn" data-col-clear="' + escapeHtml(h) + '" style="padding:5px 7px;font-size:11px;">×</button>' : ''}
                  </div>
                `;
              }).join('')}
            </div>
          </div>

          <div>
            <div class="si-section-title">
              <span>3. Multi-Row Repeater (Up to 600+ Rows)</span>
            </div>
            <div style="display:flex;flex-direction:column;gap:6px;">
              <input class="si-input" id="si-row-sel" placeholder="Row container selector (optional if {1} indexed)" value="${escapeHtml(state.rowSelector)}" />
              <div style="display:flex;gap:6px;">
                <input class="si-input" id="si-add-sel" placeholder="Add Row button selector" value="${escapeHtml(state.addRowSelector)}" />
                <button class="si-btn" id="si-pick-add" draggable="true" style="white-space:nowrap;">${state.pickMode && state.pickMode.kind === 'addRow' ? 'Click Button...' : 'Drag / Pick'}</button>
              </div>
              <div>
                <span class="si-muted">Injection Mode</span>
                <select class="si-select" id="si-mode-sel" style="width:100%;margin-top:2px;">
                  <option value="simulated"${state.injectionMode === 'simulated' ? ' selected' : ''}>Simulated (human-like)</option>
                  <option value="native"${state.injectionMode === 'native' ? ' selected' : ''}>Direct + events (page state)</option>
                  <option value="raw"${state.injectionMode === 'raw' ? ' selected' : ''}>Raw value (no page state)</option>
                  <option value="react"${state.injectionMode === 'react' ? ' selected' : ''}>React state (updateRow)</option>
                </select>
                <div class="si-muted" style="margin-top:2px;">React state uses the page's updateRow method; Excel headers are used as field names.</div>
              </div>
              ${state.injectionMode === 'react' ? `
                <div style="grid-column:1/-1;border-top:1px solid #334155;padding-top:8px;">
                  <div class="si-section-title"><span>React page hooks (editable)</span></div>
                  <div class="si-grid-2">
                    <div><span class="si-muted">Anchor selector</span><input class="si-input" data-react-config="anchorSelector" value="${escapeHtml(state.reactConfig.anchorSelector)}" style="margin-top:2px;" /></div>
                    <div><span class="si-muted">Rows prop</span><input class="si-input" data-react-config="rowsProp" value="${escapeHtml(state.reactConfig.rowsProp)}" style="margin-top:2px;" /></div>
                    <div><span class="si-muted">Update method</span><input class="si-input" data-react-config="updateRowProp" value="${escapeHtml(state.reactConfig.updateRowProp)}" style="margin-top:2px;" /></div>
                    <div><span class="si-muted">Countries prop</span><input class="si-input" data-react-config="countriesProp" value="${escapeHtml(state.reactConfig.countriesProp)}" style="margin-top:2px;" /></div>
                    <div><span class="si-muted">Fetch sous method</span><input class="si-input" data-react-config="fetchSousProp" value="${escapeHtml(state.reactConfig.fetchSousProp)}" style="margin-top:2px;" /></div>
                    <div><span class="si-muted">Sous positions prop</span><input class="si-input" data-react-config="sousPositionsProp" value="${escapeHtml(state.reactConfig.sousPositionsProp)}" style="margin-top:2px;" /></div>
                    <div><span class="si-muted">Select sous method</span><input class="si-input" data-react-config="selectSousProp" value="${escapeHtml(state.reactConfig.selectSousProp)}" style="margin-top:2px;" /></div>
                    <div><span class="si-muted">Option keys (sous / sub)</span><input class="si-input" id="si-react-option-keys" value="${escapeHtml(state.reactConfig.sousPositionKey + ', ' + state.reactConfig.subCategoryKey)}" style="margin-top:2px;" /></div>
                    <div><span class="si-muted">Row id field</span><input class="si-input" data-react-config="idField" value="${escapeHtml(state.reactConfig.idField)}" style="margin-top:2px;" /></div>
                    <div><span class="si-muted">Id header labels</span><input class="si-input" id="si-react-id-headers" value="${escapeHtml(state.reactConfig.idHeaderLabels)}" style="margin-top:2px;" /></div>
                    <div><span class="si-muted">Category field</span><input class="si-input" data-react-config="categoryField" value="${escapeHtml(state.reactConfig.categoryField)}" style="margin-top:2px;" /></div>
                    <div><span class="si-muted">Sub / sous fields</span><input class="si-input" id="si-react-sub-fields" value="${escapeHtml(state.reactConfig.subCategoryField + ', ' + state.reactConfig.sousPositionField)}" style="margin-top:2px;" /></div>
                    <div><span class="si-muted">Export / import fields</span><input class="si-input" id="si-react-country-fields" value="${escapeHtml(state.reactConfig.countryExportField + ', ' + state.reactConfig.countryImportField)}" style="margin-top:2px;" /></div>
                  </div>
                  <div style="margin-top:8px;">
                    <span class="si-muted">Header → field aliases (one per line: field: label, label)</span>
                    <textarea class="si-input" data-react-config="aliasesText" rows="7" style="margin-top:2px;width:100%;resize:vertical;font-size:11px;">${escapeHtml(state.reactConfig.aliasesText)}</textarea>
                  </div>
                  <button class="si-btn" id="si-react-config-reset" style="margin-top:6px;padding:4px 8px;font-size:11px;">Reset defaults</button>
                </div>
              ` : ''}
              <div class="si-grid-2">
                <div>
                  <span class="si-muted">Injection Pace</span>
                  <select class="si-select" id="si-pace-sel" style="width:100%;margin-top:2px;">
                    <option value="instant"${state.pace === 'instant' ? ' selected' : ''}>Instant (15/frame)</option>
                    <option value="balanced"${state.pace === 'balanced' ? ' selected' : ''}>Balanced (10ms)</option>
                    <option value="safe"${state.pace === 'safe' ? ' selected' : ''}>Heavy DOM (35ms)</option>
                  </select>
                </div>
                <div>
                  <span class="si-muted">Start Row Index</span>
                  <input class="si-input" id="si-offset-inp" type="number" min="0" value="${state.startRowOffset}" style="margin-top:2px;" />
                </div>
              </div>
            </div>
          </div>

          <div>
            <div class="si-section-title">
              <span>4. Test &amp; Batch Inject</span>
              <span class="si-muted">${mappedCount} mapped</span>
            </div>
            ${state.isRunning ? `
              <button class="si-btn si-btn-danger" id="si-stop-btn" style="width:100%">Stop Injection</button>
            ` : `
              <div class="si-grid-2">
                <button class="si-btn si-btn-primary" id="si-test-btn" ${mappedCount === 0 ? 'disabled' : ''}>Simple Test (3 Rows)</button>
                <button class="si-btn si-btn-emerald" id="si-run-btn" ${mappedCount === 0 ? 'disabled' : ''}>Inject All (${state.rows.length})</button>
              </div>
            `}

            ${state.progress.total > 0 ? `
              <div style="margin-top:8px;">
                <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
                  <span class="si-muted" id="si-prog-text">${state.progress.current} / ${state.progress.total} rows · ${state.progress.elapsedMs}ms</span>
                  <span class="si-muted">${pct}%</span>
                </div>
                <div class="si-progress-track">
                  <div class="si-progress-fill" id="si-prog-fill" style="width:${pct}%"></div>
                </div>
              </div>
            ` : ''}
          </div>
        ` : ''}

        ${state.logs.length > 0 ? `
          <div class="si-log-box">
            ${state.logs.map((l) => '<div class="si-badge-' + l.level + '">' + escapeHtml(l.text) + '</div>').join('')}
          </div>
        ` : ''}
      </div>
    `;

    // Wire events inside Shadow DOM
    const dragHandle = shadow.getElementById('si-drag-handle');
    if (dragHandle) {
      dragHandle.addEventListener('mousedown', (e) => {
        if (e.target.closest('button')) return;
        const rect = host.getBoundingClientRect();
        dragWin = { ox: e.clientX - rect.left, oy: e.clientY - rect.top };
      });
    }

    const closeBtn = shadow.getElementById('si-close-btn');
    if (closeBtn) {
      closeBtn.addEventListener('click', () => {
        host.style.display = 'none';
        hideRing();
      });
    }

    const scanBtn = shadow.getElementById('si-scan-btn');
    if (scanBtn) {
      scanBtn.addEventListener('click', () => scanPageStructure(true));
    }

    const exportTplBtn = shadow.getElementById('si-export-tpl-btn');
    if (exportTplBtn) {
      exportTplBtn.addEventListener('click', exportWebsiteTemplateXlsx);
    }

    const toggleFieldsBtn = shadow.getElementById('si-toggle-fields-btn');
    if (toggleFieldsBtn) {
      toggleFieldsBtn.addEventListener('click', () => {
        state.showFields = !state.showFields;
        render();
      });
    }

    const toggleListsBtn = shadow.getElementById('si-toggle-lists-btn');
    if (toggleListsBtn) {
      toggleListsBtn.addEventListener('click', () => {
        state.showLists = !state.showLists;
        render();
      });
    }

    const dropzone = shadow.getElementById('si-file-drop');
    const fileInput = shadow.getElementById('si-file-input');
    if (dropzone && fileInput) {
      dropzone.addEventListener('click', () => fileInput.click());
      fileInput.addEventListener('change', (e) => handleFile(e.target.files[0]));
      dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('active');
      });
      dropzone.addEventListener('dragleave', () => dropzone.classList.remove('active'));
      dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('active');
        if (e.dataTransfer && e.dataTransfer.files[0]) {
          handleFile(e.dataTransfer.files[0]);
        }
      });
    }

    const sheetSel = shadow.getElementById('si-sheet-sel');
    if (sheetSel) {
      sheetSel.addEventListener('change', (e) => loadSheet(e.target.value));
    }

    const autoMatchBtn = shadow.getElementById('si-automatch-btn');
    if (autoMatchBtn) {
      autoMatchBtn.addEventListener('click', autoMatchOnPage);
    }

    shadow.querySelectorAll('[data-col-drag]').forEach((el) => {
      el.addEventListener('dragstart', (e) => {
        const header = el.getAttribute('data-col-drag');
        state.dragPayload = { kind: 'column', header };
        if (e.dataTransfer) {
          e.dataTransfer.setData('text/plain', header);
          e.dataTransfer.effectAllowed = 'link';
        }
      });
      el.addEventListener('dragend', () => {
        state.dragPayload = null;
        hideRing();
      });
    });

    shadow.querySelectorAll('[data-col-pick]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const header = btn.getAttribute('data-col-pick');
        state.pickMode =
          state.pickMode && state.pickMode.header === header
            ? null
            : { kind: 'column', header };
        hideRing();
        render();
      });
    });

    shadow.querySelectorAll('[data-col-clear]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const header = btn.getAttribute('data-col-clear');
        delete state.mappings[header];
        saveMappings();
        render();
      });
    });

    shadow.querySelectorAll('[data-col-mode]').forEach((sel) => {
      sel.addEventListener('change', (e) => {
        const header = sel.getAttribute('data-col-mode');
        state.reactColumnModes[header] = e.target.value;
        saveReactConfig();
      });
    });

    const pickAddBtn = shadow.getElementById('si-pick-add');
    if (pickAddBtn) {
      pickAddBtn.addEventListener('dragstart', (e) => {
        state.dragPayload = { kind: 'addRow' };
        if (e.dataTransfer) e.dataTransfer.setData('text/plain', '__add_row__');
      });
      pickAddBtn.addEventListener('dragend', () => {
        state.dragPayload = null;
        hideRing();
      });
      pickAddBtn.addEventListener('click', () => {
        state.pickMode =
          state.pickMode && state.pickMode.kind === 'addRow' ? null : { kind: 'addRow' };
        hideRing();
        render();
      });
    }

    const rowSelInp = shadow.getElementById('si-row-sel');
    if (rowSelInp) {
      rowSelInp.addEventListener('input', (e) => {
        state.rowSelector = e.target.value;
        saveMappings();
      });
    }

    const addSelInp = shadow.getElementById('si-add-sel');
    if (addSelInp) {
      addSelInp.addEventListener('input', (e) => {
        state.addRowSelector = e.target.value;
        saveMappings();
      });
    }

    const paceSel = shadow.getElementById('si-pace-sel');
    if (paceSel) {
      paceSel.addEventListener('change', (e) => { state.pace = e.target.value; });
    }

    const modeSel = shadow.getElementById('si-mode-sel');
    if (modeSel) {
      modeSel.addEventListener('change', (e) => { state.injectionMode = e.target.value; render(); });
    }

    shadow.querySelectorAll('[data-react-config]').forEach((inp) => {
      inp.addEventListener('change', (e) => {
        state.reactConfig[e.target.getAttribute('data-react-config')] = e.target.value;
        saveReactConfig();
      });
    });

    const optionKeysInp = shadow.getElementById('si-react-option-keys');
    if (optionKeysInp) {
      optionKeysInp.addEventListener('change', (e) => {
        const parts = String(e.target.value).split(',').map((v) => v.trim()).filter(Boolean);
        if (parts[0]) state.reactConfig.sousPositionKey = parts[0];
        if (parts[1]) state.reactConfig.subCategoryKey = parts[1];
        saveReactConfig();
      });
    }

    const idHeadersInp = shadow.getElementById('si-react-id-headers');
    if (idHeadersInp) {
      idHeadersInp.addEventListener('change', (e) => {
        state.reactConfig.idHeaderLabels = e.target.value;
        saveReactConfig();
      });
    }

    const subFieldsInp = shadow.getElementById('si-react-sub-fields');
    if (subFieldsInp) {
      subFieldsInp.addEventListener('change', (e) => {
        const parts = String(e.target.value).split(',').map((v) => v.trim()).filter(Boolean);
        if (parts[0]) state.reactConfig.subCategoryField = parts[0];
        if (parts[1]) state.reactConfig.sousPositionField = parts[1];
        saveReactConfig();
      });
    }

    const countryFieldsInp = shadow.getElementById('si-react-country-fields');
    if (countryFieldsInp) {
      countryFieldsInp.addEventListener('change', (e) => {
        const parts = String(e.target.value).split(',').map((v) => v.trim()).filter(Boolean);
        if (parts[0]) state.reactConfig.countryExportField = parts[0];
        if (parts[1]) state.reactConfig.countryImportField = parts[1];
        saveReactConfig();
      });
    }

    const resetReactCfgBtn = shadow.getElementById('si-react-config-reset');
    if (resetReactCfgBtn) {
      resetReactCfgBtn.addEventListener('click', () => {
        state.reactConfig = { ...DEFAULT_REACT_CONFIG };
        saveReactConfig();
        render();
      });
    }

    const offsetInp = shadow.getElementById('si-offset-inp');
    if (offsetInp) {
      offsetInp.addEventListener('input', (e) => {
        state.startRowOffset = Math.max(0, parseInt(e.target.value || '0', 10) || 0);
      });
    }

    const testBtn = shadow.getElementById('si-test-btn');
    if (testBtn) testBtn.addEventListener('click', () => runInjection(3));

    const runBtn = shadow.getElementById('si-run-btn');
    if (runBtn) runBtn.addEventListener('click', () => runInjection());

    const stopBtn = shadow.getElementById('si-stop-btn');
    if (stopBtn) stopBtn.addEventListener('click', () => { state.stopRequested = true; });
  }

  // Initial passive page scan on mount
  scanPageStructure(false);
})();
