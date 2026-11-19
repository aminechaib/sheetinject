(() => {
  if (window.__SHEETINJECT_REACT_BRIDGE__) return;
  window.__SHEETINJECT_REACT_BRIDGE__ = true;

  document.documentElement.setAttribute('data-si-react-bridge', '1');

  const DEFAULT_CONFIG = {
    anchorSelector: 'textarea[data-testid^="textarea-description-"]',
    rowsProp: 'rows',
    idField: 'id',
    updateRowProp: 'updateRow',
    countriesProp: 'countries',
    fetchSousProp: 'fetchSousPositions',
    sousPositionsProp: 'sousPositionsData',
    selectSousProp: 'onSelectSousPosition',
    sousPositionKey: 'sousPosition',
    subCategoryKey: 'subCategoryCode'
  };

  function fiberOf(el) {
    const key = Object.keys(el).find(
      (k) => k.startsWith('__reactFiber$') || k.startsWith('__reactInternalInstance$')
    );
    return key ? el[key] : null;
  }

  function findOwner(cfg) {
    const anchored = cfg.anchorSelector
      ? Array.from(document.querySelectorAll(cfg.anchorSelector))
      : [];
    const anchors = [
      ...anchored,
      ...document.querySelectorAll('input, textarea, select, [role="combobox"]'),
      ...document.querySelectorAll('[data-testid]')
    ].filter(Boolean);

    let best = null;
    for (const anchor of anchors) {
      let fiber = fiberOf(anchor);
      let guard = 0;
      while (fiber && guard++ < 80) {
        const props = fiber.memoizedProps;
        if (props && Array.isArray(props[cfg.rowsProp]) && typeof props[cfg.updateRowProp] === 'function') {
          if (!best || props[cfg.rowsProp].length > best.rows.length) {
            best = { rows: props[cfg.rowsProp], props };
          }
          break;
        }
        fiber = fiber.return;
      }
    }
    if (!best) throw new Error('React table owner not found');
    return best;
  }

  window.addEventListener('message', (event) => {
    if (event.source !== window) return;
    const detail = event.data;
    if (!detail || detail.siReactRequest !== true || !detail.id) return;

    const cfg = { ...DEFAULT_CONFIG, ...(detail.config || {}) };
    let result;
    try {
      const owner = findOwner(cfg);
      const props = owner.props;
      const rows = owner.rows;
      const rowId = Number(detail.rowId);

      if (detail.action === 'inspect') {
        result = {
          ok: true,
          rows: rows.length,
          fields: rows[0] ? Object.keys(rows[0]) : [],
          countries: Array.isArray(props[cfg.countriesProp]),
          fetch: typeof props[cfg.fetchSousProp] === 'function',
          select: typeof props[cfg.selectSousProp] === 'function'
        };
      } else if (detail.action === 'update') {
        props[cfg.updateRowProp](rowId, detail.field, String(detail.value ?? ''));
        const row = rows.find((r) => Number(r && r[cfg.idField]) === rowId);
        result = { ok: true, value: row ? String(row[detail.field] ?? '') : null };
      } else if (detail.action === 'countryCode') {
        const list = props[cfg.countriesProp] || [];
        const country = list[Number(detail.code) - 1];
        result = country && country.code
          ? { ok: true, code: country.code }
          : { ok: false, warn: 'country index not found' };
      } else if (detail.action === 'countryCodes') {
        const list = props[cfg.countriesProp] || [];
        const codes = String(detail.codes ?? '')
          .split(',')
          .map((v) => v.trim())
          .filter(Boolean)
          .map((v) => list[Number(v) - 1]?.code)
          .filter(Boolean);
        result = { ok: true, codes: codes.join(',') };
      } else if (detail.action === 'fetchSous') {
        props[cfg.fetchSousProp](rowId, String(detail.category ?? ''));
        result = { ok: true };
      } else if (detail.action === 'options') {
        const options = props[cfg.sousPositionsProp]?.[rowId] || [];
        result = {
          ok: true,
          options: options.map((item) => ({
            [cfg.sousPositionKey]: item[cfg.sousPositionKey],
            [cfg.subCategoryKey]: item[cfg.subCategoryKey]
          }))
        };
      } else if (detail.action === 'selectSous') {
        const options = props[cfg.sousPositionsProp]?.[rowId] || [];
        const option = options.find(
          (item) => String(item[cfg.sousPositionKey] ?? item[cfg.subCategoryKey] ?? '') === String(detail.sousPosition)
        );
        if (!option) {
          result = { ok: false, warn: 'sous-position not found' };
        } else {
          props[cfg.selectSousProp](rowId, option);
          result = { ok: true };
        }
      } else {
        result = { ok: false, warn: 'unknown action' };
      }
    } catch (err) {
      result = { ok: false, warn: err.message || String(err) };
    }

    window.postMessage({ siReactResponse: true, id: detail.id, result }, '*');
  });
})();
