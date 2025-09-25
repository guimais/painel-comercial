(function () {
  const queryOne = (selector, scope = document) => scope.querySelector(selector);
  const queryAll = (selector, scope = document) => Array.from(scope.querySelectorAll(selector));
  const listen = (element, eventName, handler, options) => {
    if (element) {
      element.addEventListener(eventName, handler, options);
    }
  };

  const STORAGE_KEY = "comercial-demo-state-v1";
  const NOTES_STORAGE_KEY = "comercial-demo-notes-v1";

  const tableElement = queryOne('.tabela-comercial');
  const bodyElement = queryOne('tbody', tableElement);
  const headElement = queryOne('thead', tableElement);
  const cardElement = queryOne('.card');
  const observationsSection = queryOne('.observations');
  const observationsGrid = queryOne('#observations-grid');

  if (!tableElement || !bodyElement || !headElement || !cardElement) {
    return;
  }

  const columnIndex = {
    lead: 0,
    company: 1,
    stage: 2,
    lastContact: 3,
    positiveReplies: 4,
    activities: 5,
    cadence: 6,
    owner: 7,
    nextStep: 8,
    status: 9
  };

  const STALLED_THRESHOLD_DAYS = 7;
  const safeStorage = (() => {
    try {
      return window.localStorage;
    } catch (error) {
      return null;
    }
  })();

  const loadObservationNotes = () => {
    if (!safeStorage) return {};
    try {
      const raw = safeStorage.getItem(NOTES_STORAGE_KEY) || '{}';
      const parsed = JSON.parse(raw);
      return parsed && typeof parsed === 'object' ? parsed : {};
    } catch (error) {
      return {};
    }
  };

  let observationNotes = loadObservationNotes();

  const saveObservationNotes = () => {
    if (!safeStorage) return;
    try {
      safeStorage.setItem(NOTES_STORAGE_KEY, JSON.stringify(observationNotes));
    } catch (error) {
      // storage not available
    }
  };

  const scheduleObservationPersist = (() => {
    let timer;
    return () => {
      clearTimeout(timer);
      timer = setTimeout(saveObservationNotes, 240);
    };
  })();

  const observationCards = new Map();
  let observationsFallback = null;
  const HTML_ESCAPE = {
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "\"": "&quot;",
    "'": "&#39;"
  };

  const escapeHtml = (value = "") => String(value).replace(/[&<>"\u0027]/g, (char) => HTML_ESCAPE[char] || char);

  const slugify = (value = "") => {
    return String(value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "") || "referencia";
  };

  const NOTE_FOCUS_TIMEOUT = 1600;

  const renderObservationPill = (icon, value) => {
    const label = (value || '').trim();
    if (!label) {
      return '';
    }
    return `<span class="observation-pill"><i class="${icon}" aria-hidden="true"></i>${escapeHtml(label)}</span>`;
  };

  const toInteger = (value) => {
    const match = String(value).match(/-?\d+/);
    return match ? parseInt(match[0], 10) : 0;
  };

  const toDays = (value) => {
    const match = String(value).match(/(\d+)\s*d(?:ias?)?/i);
    if (match) {
      return parseInt(match[1], 10);
    }
    const parsed = toInteger(value);
    return Number.isNaN(parsed) ? 0 : parsed;
  };

  const countActivities = (value) => String(value)
    .split(/[,+]/)
    .map(toInteger)
    .reduce((total, current) => (Number.isFinite(current) ? total + current : total), 0);

  const scoreLead = (row) => {
    const replies = toInteger(cellText(row, columnIndex.positiveReplies));
    const activities = countActivities(cellText(row, columnIndex.activities));
    const cadence = toDays(cellText(row, columnIndex.cadence));
    return replies * 3 + activities + cadence / 2;
  };

  const cellText = (row, index) => (row.cells[index]?.innerText || row.cells[index]?.textContent || '').trim();
  const rowStage = (row) => (row.dataset.etapa || queryOne('td:nth-child(3)', row)?.innerText || '').replace(/\s+/g, ' ').trim();
  const rowStatus = (row) => (row.dataset.status || queryOne('td:last-child', row)?.innerText || '').replace(/\s+/g, ' ').trim();
  const tableRows = () => queryAll('.tabela-comercial tbody tr');
  const getVisibleRows = () => tableRows().filter((row) => row.style.display !== 'none');

  const VISIBLE_ROW_LIMIT = 5;

  const updateTableScrollBounds = () => {
    if (!tableWrapper) return;
    const allRows = tableRows();
    if (!allRows.length) {
      tableWrapper.style.removeProperty('max-height');
      return;
    }
    const visibleRows = getVisibleRows();
    const referenceRow = visibleRows[0] || allRows[0];
    if (!referenceRow) {
      tableWrapper.style.removeProperty('max-height');
      return;
    }
    const rowHeight = referenceRow.getBoundingClientRect().height;
    if (!rowHeight) {
      tableWrapper.style.removeProperty('max-height');
      return;
    }
    if (visibleRows.length <= VISIBLE_ROW_LIMIT) {
      tableWrapper.style.removeProperty('max-height');
      return;
    }
    const headerHeight = headElement?.getBoundingClientRect().height || 0;
    const maxHeight = Math.round(headerHeight + rowHeight * VISIBLE_ROW_LIMIT);
    tableWrapper.style.maxHeight = String(maxHeight) + 'px';
  };


  const SORT_ICONS = {
    neutral: '\u2195',
    asc: '\u2191',
    desc: '\u2193'
  };

  const ARROW_CHARS = String.fromCharCode(0x2195, 0x2191, 0x2193);
  const ARROW_REGEX = new RegExp('[' + ARROW_CHARS + ']', 'g');

  const toolbarSection = document.createElement('section');
  toolbarSection.className = 'toolbar';
  toolbarSection.innerHTML = `
    <div class="kpis" role="region" aria-label="Indicadores do funil">
      <div class="kpi" data-kpi="total">
        <div class="kpi-icon" aria-hidden="true"><i class="ri-group-line"></i></div>
        <div class="kpi-body">
          <span class="kpi-label">Leads</span>
          <strong class="kpi-value">0</strong>
          <span class="kpi-meta">Visiveis</span>
        </div>
      </div>
      <div class="kpi" data-kpi="negociacao">
        <div class="kpi-icon" aria-hidden="true"><i class="ri-shake-hands-line"></i></div>
        <div class="kpi-body">
          <span class="kpi-label">Negociacao</span>
          <strong class="kpi-value">0</strong>
          <span class="kpi-meta">Em andamento</span>
        </div>
      </div>
      <div class="kpi" data-kpi="ganho">
        <div class="kpi-icon" aria-hidden="true"><i class="ri-trophy-line"></i></div>
        <div class="kpi-body">
          <span class="kpi-label">Ganhos</span>
          <strong class="kpi-value">0</strong>
          <span class="kpi-meta">Fechados</span>
        </div>
      </div>
      <div class="kpi" data-kpi="perdido">
        <div class="kpi-icon" aria-hidden="true"><i class="ri-close-circle-line"></i></div>
        <div class="kpi-body">
          <span class="kpi-label">Perdidos</span>
          <strong class="kpi-value">0</strong>
          <span class="kpi-meta">Encerrados</span>
        </div>
      </div>
      <div class="kpi" data-kpi="tx">
        <div class="kpi-icon" aria-hidden="true"><i class="ri-pie-chart-line"></i></div>
        <div class="kpi-body">
          <span class="kpi-label">Conversao</span>
          <strong class="kpi-value">0%</strong>
          <span class="kpi-meta">Taxa de ganho</span>
        </div>
      </div>
      <div class="kpi" data-kpi="const">
        <div class="kpi-icon" aria-hidden="true"><i class="ri-time-line"></i></div>
        <div class="kpi-body">
          <span class="kpi-label">Constancia</span>
          <strong class="kpi-value">0d</strong>
          <span class="kpi-meta">Dias medios</span>
        </div>
      </div>
      <div class="kpi" data-kpi="positivas">
        <div class="kpi-icon" aria-hidden="true"><i class="ri-mail-check-line"></i></div>
        <div class="kpi-body">
          <span class="kpi-label">Resp. positivas</span>
          <strong class="kpi-value">0</strong>
          <span class="kpi-meta">Acumulado</span>
        </div>
      </div>
      <div class="kpi" data-kpi="acoes">
        <div class="kpi-icon" aria-hidden="true"><i class="ri-bar-chart-line"></i></div>
        <div class="kpi-body">
          <span class="kpi-label">Atividades</span>
          <strong class="kpi-value">0</strong>
          <span class="kpi-meta">Media por lead</span>
        </div>
      </div>
    </div>

    <div class="filters" role="region" aria-label="Filtros">
      <div class="group group-search">
        <i class="ri-search-line" aria-hidden="true"></i>
        <input id="f-q" type="search" placeholder="Buscar (lead, empresa, responsavel, proxima acao)" aria-label="Buscar" />
      </div>
      <select id="f-etapa" aria-label="Filtrar por etapa"><option value="">Etapa (todas)</option></select>
      <select id="f-status" aria-label="Filtrar por status"><option value="">Status (todos)</option></select>
      <select id="f-resp" aria-label="Filtrar por responsavel"><option value="">Responsavel (todos)</option></select>
      <button id="btn-clear" class="btn ghost" title="Limpar filtros (F)">Limpar</button>
      <div class="split" aria-hidden="true"></div>
      <button id="btn-highlight" class="btn primary" title="Destaques do semestre (G)">Destaques</button>
      <button id="btn-export" class="btn" title="Exportar CSV (E)">Exportar</button>
      <button id="btn-copy" class="btn" title="Copiar tabela">Copiar</button>
      <span class="filters-hotkeys" aria-hidden="true">
        <kbd>F</kbd> limpa · <kbd>G</kbd> destaques · <kbd>E</kbd> exporta
      </span>
      <span class="result-count" aria-live="polite"></span>
    </div>
  `;

  const tableWrapper = queryOne('.table-responsive', cardElement);
  if (tableWrapper) {
    cardElement.insertBefore(toolbarSection, tableWrapper);
  }

  const insightsSection = document.createElement('section');
  insightsSection.className = 'insights-panel hidden';
  insightsSection.innerHTML = `
    <div class="insights-header">
      <span class="insights-title"><i class="ri-lightbulb-line" aria-hidden="true"></i> Insights rapidos</span>
      <button type="button" class="insights-toggle" aria-expanded="true">Ocultar</button>
    </div>
    <div class="insights-body"></div>
  `;
  if (tableWrapper) {
    cardElement.insertBefore(insightsSection, tableWrapper);
  } else {
    cardElement.appendChild(insightsSection);
  }

  const insightsBody = queryOne('.insights-body', insightsSection);
  const insightsToggle = queryOne('.insights-toggle', insightsSection);

  const insightsStyles = document.createElement('style');
  insightsStyles.textContent = [
    '.insights-panel{margin:14px 0 18px;padding:16px 18px;border-radius:14px;border:1px solid var(--border);background:linear-gradient(135deg, rgba(18,130,162,.16), rgba(0,31,84,.32));display:grid;gap:14px;}',
    '.insights-panel.hidden{display:none;}',
    '.insights-panel.collapsed .insights-body{display:none;}',
    '.insights-header{display:flex;align-items:center;justify-content:space-between;gap:12px;}',
    '.insights-title{display:inline-flex;align-items:center;gap:10px;font-size:12px;font-weight:700;letter-spacing:.4px;text-transform:uppercase;color:var(--ink);}',
    '.insights-title i{font-size:18px;color:var(--blue);}',
    '.insights-toggle{background:transparent;border:1px solid var(--border);color:var(--ink);font-size:12px;font-weight:600;padding:6px 14px;border-radius:999px;cursor:pointer;transition:background .2s ease,color .2s ease;}',
    '.insights-toggle:hover{background:rgba(18,130,162,.18);color:var(--tone-cream);}',
    '.insights-body{display:grid;gap:12px;}',
    '.insights-grid{display:grid;gap:12px;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));}',
    '.insight-card{padding:14px;border-radius:12px;background:rgba(0,31,84,.32);border:1px solid rgba(18,130,162,.28);display:grid;gap:10px;}',
    '.insight-card header{display:flex;align-items:center;justify-content:space-between;gap:8px;}',
    '.insight-card strong{font-size:12px;font-weight:700;letter-spacing:.3px;text-transform:uppercase;color:color-mix(in srgb, var(--ink) 78%, var(--blue) 22%);}',
    '.insight-count{font-size:22px;font-weight:700;color:var(--ink);}',
    '.insight-list{list-style:none;margin:0;padding:0;display:grid;gap:6px;}',
    '.insight-list li{display:grid;gap:2px;}',
    '.insight-list .lead-name{font-weight:700;color:var(--ink);font-size:13px;}',
    '.insight-list .lead-meta{font-size:12px;color:color-mix(in srgb, var(--ink) 60%, var(--blue) 40%);}',
    '.insight-empty{font-size:12px;color:color-mix(in srgb, var(--ink) 62%, var(--blue) 38%);}'
  ].join('');
  document.head.appendChild(insightsStyles);
  const highlightModal = document.createElement('div');
  highlightModal.className = 'highlight-modal hidden';
  highlightModal.innerHTML = [
    '    <article class="highlight-card" role="dialog" aria-modal="true" aria-labelledby="highlight-dialog-title">',
    '      <header class="highlight-head">',
    '        <div class="highlight-headline">',
    '          <span class="highlight-title" id="highlight-dialog-title"><i class="ri-rocket-line" aria-hidden="true"></i>Top 3 do mes</span>',
    '          <span class="highlight-period" aria-live="polite"></span>',
    '        </div>',
    '        <button type="button" class="highlight-close" aria-label="Fechar painel de destaques"><i class="ri-close-line"></i></button>',
    '      </header>',
    '      <p class="highlight-caption">Ranking baseado em respostas positivas, atividades e constancia.</p>',
    '      <div class="highlight-body"></div>',
    '    </article>'
  ].join('');
  document.body.appendChild(highlightModal);

  const highlightBody = queryOne('.highlight-body', highlightModal);
  const highlightPeriod = queryOne('.highlight-period', highlightModal);
  const highlightClose = queryOne('.highlight-close', highlightModal);

  const highlightStyles = document.createElement('style');
  highlightStyles.textContent = [
    'body.highlight-open{overflow:hidden;}',
    '.card{transition:transform .24s ease, filter .24s ease, opacity .24s ease;}',
    '.highlight-modal{position:fixed;inset:0;display:flex;align-items:center;justify-content:center;padding:28px;background:rgba(4,12,24,.68);backdrop-filter:blur(12px) saturate(120%);z-index:999;opacity:0;pointer-events:none;transition:opacity .24s ease;}',
    '.highlight-modal.hidden{display:none;}',
    '.highlight-modal.visible{opacity:1;pointer-events:auto;}',
    '.highlight-card{position:relative;width:min(580px,95vw);border-radius:24px;padding:34px 32px 32px;background:linear-gradient(170deg, rgba(18,130,162,.42), rgba(0,31,84,.9));border:1px solid rgba(254,252,251,.2);box-shadow:0 30px 90px rgba(0,0,0,.55);display:grid;gap:22px;color:var(--ink);overflow:hidden;}',
    '.highlight-card::before{content:"";position:absolute;inset:0;border-radius:inherit;background:linear-gradient(140deg, rgba(18,130,162,.24), transparent 55%);opacity:.9;pointer-events:none;}',
    '.highlight-head{position:relative;display:flex;align-items:flex-start;justify-content:space-between;gap:16px;z-index:1;}',
    '.highlight-headline{display:flex;flex-direction:column;gap:6px;}',
    '.highlight-title{display:inline-flex;align-items:center;gap:12px;font-size:15px;font-weight:800;letter-spacing:.3px;text-transform:uppercase;}',
    '.highlight-title i{font-size:22px;color:var(--blue);}',
    '.highlight-period{font-size:12px;letter-spacing:.4px;text-transform:uppercase;color:color-mix(in srgb, var(--ink) 62%, var(--blue) 38%);}',
    '.highlight-caption{position:relative;z-index:1;margin:0;font-size:13px;color:color-mix(in srgb, var(--ink) 72%, var(--blue) 28%);}',
    '.highlight-close{position:relative;z-index:1;background:rgba(10,17,40,.58);border:1px solid rgba(254,252,251,.28);color:var(--ink);width:40px;height:40px;border-radius:50%;display:grid;place-items:center;cursor:pointer;transition:transform .2s ease, background .2s ease, color .2s ease;box-shadow:0 12px 28px rgba(0,0,0,.38);}',
    '.highlight-close:hover,.highlight-close:focus-visible{background:rgba(18,130,162,.32);color:var(--tone-cream);transform:translateY(-1px);}',
    '.highlight-body{position:relative;z-index:1;display:grid;gap:18px;}',
    '.highlight-list{list-style:none;margin:0;padding:0;display:grid;gap:16px;}',
    '.highlight-list li{position:relative;display:grid;gap:12px;padding:20px;border-radius:18px;background:rgba(0,31,84,.55);border:1px solid rgba(254,252,251,.18);box-shadow:0 20px 40px rgba(0,0,0,.38);overflow:hidden;}',
    '.highlight-list li::after{content:"";position:absolute;inset:0;border-radius:inherit;border:1px solid rgba(254,252,251,.1);opacity:.8;pointer-events:none;}',
    '.highlight-rank{position:relative;z-index:1;font-size:13px;font-weight:700;letter-spacing:.3px;text-transform:uppercase;color:color-mix(in srgb, var(--ink) 74%, var(--blue) 26%);display:inline-flex;align-items:center;gap:8px;}',
    '.highlight-rank i{font-size:18px;color:var(--blue);}',
    '.highlight-lead{position:relative;z-index:1;display:flex;align-items:center;justify-content:space-between;gap:12px;}',
    '.highlight-lead strong{font-size:21px;font-weight:800;}',
    '.highlight-score{font-size:22px;font-weight:800;color:var(--tone-cream);}',
    '.highlight-meta{position:relative;z-index:1;display:flex;flex-wrap:wrap;gap:10px;font-size:12px;color:color-mix(in srgb, var(--ink) 62%, var(--blue) 38%);}',
    '.highlight-meta span{display:inline-flex;align-items:center;gap:6px;padding:6px 12px;border-radius:999px;background:rgba(18,130,162,.22);border:1px solid rgba(254,252,251,.14);}',
    '.highlight-meta span i{font-size:14px;color:var(--tone-cream);opacity:.86;}',
    '.highlight-empty{position:relative;z-index:1;font-size:12px;color:color-mix(in srgb, var(--ink) 62%, var(--blue) 38%);text-align:center;padding:20px;border-radius:16px;border:1px solid rgba(254,252,251,.12);background:rgba(10,17,40,.42);}',
    '.card.modal-blur{transform:scale(.98);opacity:.82;}',
    '.card.modal-blur .toolbar, .card.modal-blur .insights-panel, .card.modal-blur .table-responsive, .card.modal-blur .nota{filter:blur(6px) brightness(.75);pointer-events:none;}'
  ].join('');
  document.head.appendChild(highlightStyles);

  listen(insightsToggle, 'click', () => {
    const collapsed = insightsSection.classList.toggle('collapsed');
    if (insightsToggle) {
      insightsToggle.textContent = collapsed ? 'Mostrar' : 'Ocultar';
      insightsToggle.setAttribute('aria-expanded', String(!collapsed));
    }
  });


  const sortableStyles = document.createElement('style');
  sortableStyles.textContent = [
    'th.sortable{cursor:pointer;user-select:none}',
    'th.sortable .arrow{margin-left:8px;font-size:12px;opacity:.55;transition:opacity .2s ease,transform .2s ease;display:inline-block;min-width:12px;text-align:center}',
    'th.sortable.sort-asc .arrow{opacity:1;transform:translateY(-1px)}',
    'th.sortable.sort-desc .arrow{opacity:1;transform:translateY(1px)}'
  ].join('');
  document.head.appendChild(sortableStyles);

  const uniqueSorted = (values) => [...new Set(values)].filter(Boolean).sort((a, b) => a.localeCompare(b));

  const stageFilter = queryOne('#f-etapa', toolbarSection);
  const statusFilter = queryOne('#f-status', toolbarSection);
  const ownerFilter = queryOne('#f-resp', toolbarSection);
  const searchInput = queryOne('#f-q', toolbarSection);
  const clearButton = queryOne('#btn-clear', toolbarSection);
  const highlightButton = queryOne('#btn-highlight', toolbarSection);
  const exportButton = queryOne('#btn-export', toolbarSection);
  const copyButton = queryOne('#btn-copy', toolbarSection);
  const resultSummary = queryOne('.result-count', toolbarSection);
  let lastHighlightRows = [];

  uniqueSorted(tableRows().map(rowStage)).forEach((stage) => {
    const option = document.createElement('option');
    option.value = stage;
    option.textContent = stage;
    stageFilter?.appendChild(option);
  });

  uniqueSorted(tableRows().map(rowStatus)).forEach((status) => {
    const option = document.createElement('option');
    option.value = status;
    option.textContent = status;
    statusFilter?.appendChild(option);
  });

  uniqueSorted(tableRows().map((row) => cellText(row, columnIndex.owner))).forEach((owner) => {
    const option = document.createElement('option');
    option.value = owner;
    option.textContent = owner;
    ownerFilter?.appendChild(option);
  });

  const headerCells = queryAll('th', headElement);
  headerCells.forEach((headerCell, index) => {
    headerCell.classList.add('sortable');
    const arrow = document.createElement('span');
    arrow.className = 'arrow';
    arrow.textContent = SORT_ICONS.neutral;
    headerCell.appendChild(arrow);
    listen(headerCell, 'click', () => {
      const currentIndex = Number(tableElement.dataset.sortCol);
      const currentDirection = tableElement.dataset.sortDir || null;
      const nextDirection = currentIndex === index && currentDirection === 'asc' ? 'desc' : 'asc';
      sortBy(index, nextDirection);
      persistState();
    });
  });

  const sortBy = (index, direction = 'asc') => {
    const factor = direction === 'asc' ? 1 : -1;
    const sortedRows = tableRows().sort((rowA, rowB) => {
      let valueA = cellText(rowA, index);
      let valueB = cellText(rowB, index);

      if (index === columnIndex.cadence) {
        valueA = toDays(valueA);
        valueB = toDays(valueB);
      } else if (index === columnIndex.positiveReplies) {
        valueA = toInteger(valueA);
        valueB = toInteger(valueB);
      } else {
        valueA = valueA.toLowerCase();
        valueB = valueB.toLowerCase();
      }

      if (valueA < valueB) return -1 * factor;
      if (valueA > valueB) return 1 * factor;
      return 0;
    });

    bodyElement.append(...sortedRows);
    buildObservationCards();
    tableElement.dataset.sortCol = String(index);
    tableElement.dataset.sortDir = direction;

    headerCells.forEach((headerCell, headerIndex) => {
      headerCell.classList.remove('sort-asc', 'sort-desc');
      const arrow = queryOne('.arrow', headerCell);
      if (!arrow) return;
      if (headerIndex === index) {
        const activeClass = direction === 'asc' ? 'sort-asc' : 'sort-desc';
        headerCell.classList.add(activeClass);
        arrow.textContent = direction === 'asc' ? SORT_ICONS.asc : SORT_ICONS.desc;
      } else {
        arrow.textContent = SORT_ICONS.neutral;
      }
    });
  };

  const debounce = (fn, delay = 160) => {
    let timer;
    return (...args) => {
      clearTimeout(timer);
      timer = setTimeout(() => fn(...args), delay);
    };
  };

  listen(window, 'resize', debounce(updateTableScrollBounds, 180));

  const updateIndicators = () => {
    const visibleRows = tableRows().filter((row) => row.style.display !== 'none');
    const totalVisible = visibleRows.length;
    const wonCount = visibleRows.filter((row) => /ganho/i.test(rowStatus(row))).length;
    const lostCount = visibleRows.filter((row) => /perdido/i.test(rowStatus(row))).length;
    const negotiationCount = visibleRows.filter((row) => /negocia/i.test(rowStatus(row))).length;
    const cadenceSum = visibleRows.reduce((total, row) => total + toDays(cellText(row, columnIndex.cadence)), 0);
    const cadenceAverage = totalVisible ? Math.round(cadenceSum / totalVisible) : 0;
    const repliesTotal = visibleRows.reduce((total, row) => total + toInteger(cellText(row, columnIndex.positiveReplies)), 0);
    const activitiesTotal = visibleRows.reduce((total, row) => total + countActivities(cellText(row, columnIndex.activities)), 0);
    const conversionRate = totalVisible ? Math.round((wonCount / totalVisible) * 100) : 0;
    const activitiesAverageRaw = totalVisible ? activitiesTotal / totalVisible : 0;
    const activitiesAverage = totalVisible
      ? (activitiesAverageRaw >= 10 ? Math.round(activitiesAverageRaw).toString() : activitiesAverageRaw.toFixed(1).replace('.', ',').replace(/,0$/, ''))
      : '0';

    setKpiValue('total', totalVisible);
    setKpiValue('negociacao', negotiationCount);
    setKpiValue('ganho', wonCount);
    setKpiValue('perdido', lostCount);
    setKpiValue('tx', `${conversionRate}%`);
    setKpiValue('const', `${cadenceAverage}d`);
    setKpiValue('positivas', repliesTotal);
    setKpiValue('acoes', activitiesAverage);
  };

  const setKpiValue = (key, value) => {
    const element = queryOne(`.kpi[data-kpi="${key}"] .kpi-value`, toolbarSection);
    if (element) {
      element.textContent = String(value);
    }
  };

  const renderResultSummary = (visibleCount) => {
    if (!resultSummary) return;
    if (!Number.isFinite(visibleCount) || visibleCount < 0) {
      resultSummary.textContent = '';
      return;
    }
    const label = visibleCount === 0
      ? 'Nenhum lead visivel'
      : visibleCount === 1
        ? '1 lead visivel'
        : `${visibleCount} leads visiveis`;
    resultSummary.textContent = label;
  };

  const createInsightCard = (title, value, listHtml) => `
    <article class="insight-card">
      <header>
        <strong>${title}</strong>
        <span class="insight-count">${value}</span>
      </header>
      <ul class="insight-list">${listHtml}</ul>
    </article>
  `;

  const updateInsightsPanel = () => {
    if (!insightsBody || !insightsSection) return;

    const visibleRows = getVisibleRows();
    if (!visibleRows.length) {
      insightsSection.classList.add('hidden');
      insightsBody.innerHTML = '<p class="insight-empty">Nenhum lead visivel no filtro atual.</p>';
      return;
    }

    insightsSection.classList.remove('hidden');

    const repliesSorted = [...visibleRows].sort((a, b) =>
      toInteger(cellText(b, columnIndex.positiveReplies)) - toInteger(cellText(a, columnIndex.positiveReplies))
    );
    const activitiesSorted = [...visibleRows].sort((a, b) =>
      countActivities(cellText(b, columnIndex.activities)) - countActivities(cellText(a, columnIndex.activities))
    );
    const stalledSorted = visibleRows
      .filter((row) => toDays(cellText(row, columnIndex.cadence)) >= STALLED_THRESHOLD_DAYS)
      .sort((a, b) => toDays(cellText(b, columnIndex.cadence)) - toDays(cellText(a, columnIndex.cadence)));

    const repliesPeak = repliesSorted[0] ? `${toInteger(cellText(repliesSorted[0], columnIndex.positiveReplies))} resp.` : '0';
    const activitiesPeak = activitiesSorted[0]
      ? `${countActivities(cellText(activitiesSorted[0], columnIndex.activities))} atv.`
      : '0';
    const stalledLabel = stalledSorted.length ? `${stalledSorted.length}` : '0';

    const cards = [];
    cards.push(createInsightCard(
      'Respostas positivas',
      repliesPeak,
      renderLeadItems(repliesSorted, (row) => `${toInteger(cellText(row, columnIndex.positiveReplies))} resp.`)
    ));
    cards.push(createInsightCard(
      'Atividades registradas',
      activitiesPeak,
      renderLeadItems(activitiesSorted, (row) => `${countActivities(cellText(row, columnIndex.activities))} atv.`)
    ));
    cards.push(createInsightCard(
      'Leads estagnados',
      stalledLabel,
      renderLeadItems(stalledSorted, (row) => `${toDays(cellText(row, columnIndex.cadence))}d sem contato`)
    ));

    const highlightRows = lastHighlightRows.length ? lastHighlightRows : repliesSorted.slice(0, 3);
    if (highlightRows.length) {
      const scoredHighlight = highlightRows
        .map((row) => ({ row, score: scoreLead(row) }))
        .sort((a, b) => b.score - a.score)
        .map((item) => item.row);
      cards.push(createInsightCard(
        'Ultimo Top 3',
        `${Math.min(3, highlightRows.length)}`,
        renderLeadItems(scoredHighlight, (row) => `${Math.round(scoreLead(row))} pts`)
      ));
    }

    insightsBody.innerHTML = `<div class="insights-grid">${cards.join('')}</div>`;
  };


  const renderLeadItems = (rows, metaBuilder) => {
    if (!rows.length) {
      return '<li class="insight-empty">Nenhum lead</li>';
    }
    return rows.slice(0, 3)
      .map((row) => {
        const name = cellText(row, columnIndex.lead);
        const meta = metaBuilder ? metaBuilder(row) : '';
        const metaHtml = meta ? `<span class="lead-meta">${meta}</span>` : '';
        return `<li><span class="lead-name">${name}</span>${metaHtml}</li>`;
      })
      .join('');
  };

  const rankLabels = ['Top 1', 'Top 2', 'Top 3'];
  const rankIcons = ['ri-trophy-line', 'ri-medal-line', 'ri-award-line'];
  const formatHighlightPeriod = () => {
    try {
      const formatted = new Intl.DateTimeFormat('pt-BR', { month: 'long', year: 'numeric' }).format(new Date());
      return formatted.charAt(0).toUpperCase() + formatted.slice(1);
    } catch (error) {
      return '';
    }
  };

  const renderHighlightModal = (entries) => {
    if (!entries.length) {
      return '<div class="highlight-empty">Nenhum destaque disponivel.</div>';
    }
    const items = entries.map(({ row, score }, index) => {
      const replies = toInteger(cellText(row, columnIndex.positiveReplies));
      const activitiesSummary = cellText(row, columnIndex.activities);
      const cadenceDays = toDays(cellText(row, columnIndex.cadence));
      const ownerName = cellText(row, columnIndex.owner);
      const stageName = rowStage(row);
      const icon = rankIcons[index] || 'ri-trophy-line';
      const label = rankLabels[index] || `Top ${index + 1}`;
      const leadName = cellText(row, columnIndex.lead);
      return `
        <li>
          <span class="highlight-rank"><i class="${icon}"></i>${label}</span>
          <div class="highlight-lead">
            <strong>${leadName}</strong>
            <span class="highlight-score">${Math.round(score)}</span>
          </div>
          <div class="highlight-meta">
            <span><i class="ri-compass-line"></i>${stageName}</span>
            <span><i class="ri-mail-check-line"></i>${replies} resp.</span>
            <span><i class="ri-bar-chart-line"></i>${activitiesSummary}</span>
            <span><i class="ri-time-line"></i>${cadenceDays}d cadencia</span>
            <span><i class="ri-user-line"></i>${ownerName}</span>
          </div>
        </li>
      `;
    }).join('');
    return `<ul class="highlight-list">${items}</ul>`;
  };

  const updateObservationVisibility = () => {
    if (!observationsGrid) return;
    if (!observationCards.size) {
      if (observationsFallback) {
        observationsFallback.hidden = false;
      }
      return;
    }
    const visibleRows = getVisibleRows();
    const visibleIds = new Set(visibleRows.map((row) => row.id));
    let visibleCount = 0;

    observationCards.forEach(({ item }) => {
      if (!visibleIds.size) {
        item.style.display = 'none';
      } else if (visibleIds.has(item.dataset.leadId)) {
        item.style.display = '';
        visibleCount += 1;
      } else {
        item.style.display = 'none';
      }
    });

    if (observationsFallback) {
      observationsFallback.hidden = visibleCount > 0;
    }
  };

  const buildObservationCards = () => {
    if (!observationsGrid) return;

    observationsGrid.innerHTML = '';
    observationCards.clear();
    observationsFallback = null;

    const rows = tableRows();
    if (!rows.length) {
      const emptyItem = document.createElement('li');
      emptyItem.className = 'observation-empty observation-empty-message';
      emptyItem.textContent = 'Nenhum lead cadastrado.';
      emptyItem.setAttribute('role', 'listitem');
      observationsGrid.appendChild(emptyItem);
      return;
    }

    const fragment = document.createDocumentFragment();
    const usedIds = new Set();
    let needsSave = false;

    rows.forEach((row) => {
      const leadName = cellText(row, columnIndex.lead);
      if (!leadName) return;

      let rowId = row.id;
      if (rowId && usedIds.has(rowId)) {
        rowId = '';
      }
      if (!rowId) {
        const baseId = `lead-${slugify(leadName)}`;
        let candidate = baseId;
        let counter = 1;
        while (usedIds.has(candidate)) {
          candidate = `${baseId}-${counter}`;
          counter += 1;
        }
        rowId = candidate;
        row.id = rowId;
      }
      usedIds.add(rowId);

      const company = cellText(row, columnIndex.company);
      const owner = cellText(row, columnIndex.owner);
      const nextStep = cellText(row, columnIndex.nextStep);
      const stageName = rowStage(row);
      const statusName = rowStatus(row);
      const cadenceLabel = cellText(row, columnIndex.cadence);
      const lastContact = cellText(row, columnIndex.lastContact);

      const item = document.createElement('li');
      item.className = 'observation-item';
      item.dataset.leadId = rowId;
      item.setAttribute('role', 'listitem');

      const noteId = `note-${rowId}`;

      const metaPills = [
        renderObservationPill('ri-map-pin-2-line', stageName),
        renderObservationPill('ri-flag-2-line', statusName),
        renderObservationPill('ri-calendar-event-line', lastContact),
        renderObservationPill('ri-time-line', cadenceLabel),
        renderObservationPill('ri-user-voice-line', owner),
        renderObservationPill('ri-compass-3-line', nextStep)
      ].filter(Boolean).join('');

      item.innerHTML = `
        <article class="observation-card">
          <header class="observation-head">
            <div class="observation-title">
              <span class="observation-lead-name">${escapeHtml(leadName)}</span>
              <span class="observation-company">${escapeHtml(company)}</span>
            </div>
            <button type="button" class="observation-jump" data-target="${rowId}">
              <i class="ri-focus-2-line" aria-hidden="true"></i>
              Ver na tabela
            </button>
          </header>
          <div class="observation-meta">${metaPills}</div>
          <div class="observation-note-group">
            <label class="observation-note-label" for="${noteId}">Observacoes</label>
            <textarea class="observation-note" id="${noteId}" data-lead-id="${rowId}" aria-label="Observacoes para ${escapeHtml(leadName)}" placeholder="Anote direcionamentos e percepcoes sobre ${escapeHtml(leadName)}"></textarea>
          </div>
        </article>
      `;

      const textarea = queryOne('.observation-note', item);
      if (textarea) {
        const storedNote = observationNotes[rowId] ?? observationNotes[leadName];
        if (typeof storedNote === 'string') {
          textarea.value = storedNote;
          if (observationNotes[leadName] && !observationNotes[rowId]) {
            observationNotes[rowId] = storedNote;
            delete observationNotes[leadName];
            needsSave = true;
          }
        }
        listen(textarea, 'input', () => {
          const noteValue = textarea.value;
          if (noteValue && noteValue.trim()) {
            observationNotes[rowId] = noteValue;
          } else {
            delete observationNotes[rowId];
          }
          scheduleObservationPersist();
        });
      }

      const jumpButton = queryOne('.observation-jump', item);
      if (jumpButton) {
        listen(jumpButton, 'click', () => {
          const targetRow = document.getElementById(rowId);
          if (!targetRow) return;
          targetRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
          targetRow.classList.add('note-focus');
          const timerId = Number(targetRow.dataset.noteFocusTimer || 0);
          if (timerId) {
            clearTimeout(timerId);
          }
          const timeoutHandle = setTimeout(() => {
            targetRow.classList.remove('note-focus');
            delete targetRow.dataset.noteFocusTimer;
          }, NOTE_FOCUS_TIMEOUT);
          targetRow.dataset.noteFocusTimer = String(timeoutHandle);
        });
      }

      fragment.appendChild(item);
      observationCards.set(rowId, { item, row });
    });

    observationsGrid.appendChild(fragment);

    if (needsSave) {
      saveObservationNotes();
    }

    observationsFallback = document.createElement('li');
    observationsFallback.className = 'observation-empty observation-empty-message';
    observationsFallback.setAttribute('role', 'listitem');
    observationsFallback.textContent = 'Nenhum lead visivel com os filtros atuais.';
    observationsFallback.hidden = true;
    observationsGrid.appendChild(observationsFallback);

    updateObservationVisibility();
  };
  const openHighlightModal = (entries) => {
    if (!highlightModal || !highlightBody) return;
    highlightBody.innerHTML = renderHighlightModal(entries);
    highlightModal.classList.remove('hidden');
    requestAnimationFrame(() => highlightModal.classList.add('visible'));
    if (highlightPeriod) {
      highlightPeriod.textContent = formatHighlightPeriod();
    }
    document.body.classList.add('highlight-open');
    cardElement.classList.add('modal-blur');
    highlightButton?.setAttribute('aria-expanded', 'true');
    highlightClose?.focus({ preventScroll: true });
  };

  const closeHighlightModal = () => {
    if (!highlightModal) return;
    document.body.classList.remove('highlight-open');
    if (!highlightModal.classList.contains('visible')) {
      highlightModal.classList.add('hidden');
      cardElement.classList.remove('modal-blur');
      highlightButton?.setAttribute('aria-expanded', 'false');
      highlightButton?.focus({ preventScroll: true });
      return;
    }
    highlightModal.classList.remove('visible');
    const handle = () => {
      highlightModal.classList.add('hidden');
      highlightModal.removeEventListener('transitionend', handle);
    };
    highlightModal.addEventListener('transitionend', handle, { once: true });
    cardElement.classList.remove('modal-blur');
    highlightButton?.setAttribute('aria-expanded', 'false');
    highlightButton?.focus({ preventScroll: true });
  };

  const isHighlightModalVisible = () => highlightModal?.classList.contains('visible');

  listen(highlightClose, 'click', closeHighlightModal);
  listen(highlightModal, 'click', (event) => {
    if (event.target === highlightModal) {
      closeHighlightModal();
    }
  });

  listen(document, 'keydown', (event) => {
    const key = event.key?.toLowerCase();

    if (key === 'escape' && isHighlightModalVisible()) {
      closeHighlightModal();
      return;
    }

    const activeElement = document.activeElement;
    const isTyping = activeElement && /^(input|textarea)$/i.test(activeElement.tagName);
    const isEditable = activeElement && activeElement.isContentEditable;
    if (isTyping || isEditable || event.metaKey || event.ctrlKey || event.altKey) {
      return;
    }

    if (key === 'f') {
      event.preventDefault();
      clearButton?.click();
    } else if (key === 'g') {
      event.preventDefault();
      highlightButton?.click();
    } else if (key === 'e') {
      event.preventDefault();
      exportButton?.click();
    }
  });
  const applyFilters = () => {
    const queryValue = (searchInput?.value || '').trim().toLowerCase();
    const stageValue = stageFilter?.value || '';
    const statusValue = statusFilter?.value || '';
    const ownerValue = ownerFilter?.value || '';

    let visibleCount = 0;

    tableRows().forEach((row) => {
      const searchableText = [
        cellText(row, columnIndex.lead),
        cellText(row, columnIndex.company),
        cellText(row, columnIndex.owner),
        cellText(row, columnIndex.nextStep)
      ].join(' ').toLowerCase();

      const matchesQuery = !queryValue || searchableText.includes(queryValue);
      const matchesStage = !stageValue || rowStage(row) === stageValue;
      const matchesStatus = !statusValue || rowStatus(row) === statusValue;
      const matchesOwner = !ownerValue || cellText(row, columnIndex.owner) === ownerValue;

      const shouldDisplay = matchesQuery && matchesStage && matchesStatus && matchesOwner;
      row.style.display = shouldDisplay ? '' : 'none';
      if (shouldDisplay) {
        visibleCount += 1;
      }
    });

    renderResultSummary(visibleCount);
    updateIndicators();
    updateInsightsPanel();
    updateTableScrollBounds();
    updateObservationVisibility();
    if (isHighlightModalVisible()) { closeHighlightModal(); }
  };

  const applyFiltersDebounced = debounce(() => {
    applyFilters();
    persistState();
  });

  listen(searchInput, 'input', applyFiltersDebounced);
  listen(stageFilter, 'change', () => { applyFilters(); persistState(); });
  listen(statusFilter, 'change', () => { applyFilters(); persistState(); });
  listen(ownerFilter, 'change', () => { applyFilters(); persistState(); });

  listen(clearButton, 'click', () => {
    if (searchInput) searchInput.value = '';
    if (stageFilter) stageFilter.value = '';
    if (statusFilter) statusFilter.value = '';
    if (ownerFilter) ownerFilter.value = '';
    applyFilters();
    persistState();
  });

  listen(highlightButton, 'click', () => {
    const visibleRows = getVisibleRows();
    if (!visibleRows.length) {
      lastHighlightRows = [];
      updateInsightsPanel();
      closeHighlightModal();
      return;
    }
    const scored = visibleRows
      .map((row) => ({ row, score: scoreLead(row) }))
      .sort((a, b) => b.score - a.score)
      .slice(0, 3);

    tableRows().forEach((row) => row.classList.remove('topline'));
    scored.forEach(({ row }) => row.classList.add('topline'));
    lastHighlightRows = scored.map((item) => item.row);
    updateInsightsPanel();
    openHighlightModal(scored);
  });

  const exportVisibleRows = () => {
    const visibleRows = tableRows().filter((row) => row.style.display !== 'none');
    const headers = queryAll('.tabela-comercial thead th').map((header) => header.innerText.replace(ARROW_REGEX, '').trim());
    const dataLines = visibleRows.map((row) =>
      Array.from(row.cells)
        .map((cell) => `"${cell.innerText.replace(/"/g, '""')}"`)
        .join(',')
    );
    const csvContent = [headers.join(','), ...dataLines].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const downloadLink = document.createElement('a');
    downloadLink.href = URL.createObjectURL(blob);
    downloadLink.download = 'tabela-comercial-demo.csv';
    downloadLink.click();
    URL.revokeObjectURL(downloadLink.href);
  };

  const copyVisibleRows = async () => {
    const visibleRows = tableRows().filter((row) => row.style.display !== 'none');
    const headers = queryAll('.tabela-comercial thead th').map((header) => header.innerText.replace(ARROW_REGEX, '').trim());
    const dataLines = visibleRows.map((row) =>
      Array.from(row.cells)
        .map((cell) => cell.innerText)
        .join('\t')
    );
    const clipboardText = [headers.join('\t'), ...dataLines].join('\n');
    try {
      await navigator.clipboard.writeText(clipboardText);
    } catch (error) {
      console.error('Clipboard copy failed', error);
    }
  };

  listen(exportButton, 'click', exportVisibleRows);
  listen(copyButton, 'click', copyVisibleRows);

  const persistState = () => {
    const state = {
      query: searchInput?.value || '',
      stage: stageFilter?.value || '',
      status: statusFilter?.value || '',
      owner: ownerFilter?.value || '',
      sortColumn: tableElement.dataset.sortCol || null,
      sortDirection: tableElement.dataset.sortDir || null
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  };

  const restoreState = () => {
    try {
      const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
      if (saved.query != null && searchInput) searchInput.value = saved.query;
      if (saved.stage != null && stageFilter) stageFilter.value = saved.stage;
      if (saved.status != null && statusFilter) statusFilter.value = saved.status;
      if (saved.owner != null && ownerFilter) ownerFilter.value = saved.owner;
      if (saved.sortColumn != null) {
        const column = parseInt(saved.sortColumn, 10);
        const direction = saved.sortDirection === 'desc' ? 'desc' : 'asc';
        sortBy(column, direction);
      }
    } catch (error) {
      console.error('Estado do filtro nao pode ser restaurado', error);
    }
  };

  buildObservationCards();
  restoreState();
  applyFilters();
  persistState();
})();



// === PLUS SAFE v4 (gated: Ctrl+Shift+P) ===
;(() => {
  let PLUS_ON = false;
  let WIRED = false;
  const STORAGE_LOSS = "comercial-loss-reasons-v1";

  const lossMap = (() => {
    try { return JSON.parse(localStorage.getItem(STORAGE_LOSS) || "{}"); } catch(e){ return {}; }
  })();
  const saveLoss = () => { try { localStorage.setItem(STORAGE_LOSS, JSON.stringify(lossMap)); } catch(e){} };

  const table = document.querySelector('.tabela-comercial');
  if (!table) return;
  const body = table.querySelector('tbody');
  const head = table.querySelector('thead');
  const columnIndex = { lead:0, company:1, stage:2, lastContact:3, positiveReplies:4, activities:5, cadence:6, owner:7, nextStep:8, status:9 };
  const cellText = (row, i) => (row.cells[i]?.innerText || '').trim();
  const rowStatus = (row) => (row.dataset.status || (row.querySelector('td:last-child .status')?.textContent || '')).trim();
  const tableRows = () => Array.from(body.querySelectorAll('tr'));
  const visibleRows = () => tableRows().filter(r => r.style.display !== 'none');
  const headers = () => Array.from(head.querySelectorAll('th')).map(th => th.textContent.trim());

  // Hint (no layout touch)
  const hintId='plus-hint';
  const setHint = (on) => {
    let el = document.getElementById(hintId);
    if (!el){ el=document.createElement('div'); el.id=hintId; el.style.cssText='position:fixed;bottom:12px;right:14px;font:600 12px/1 system-ui,Segoe UI,Roboto,Arial;color:#9fb9ff;opacity:.75;user-select:none;pointer-events:none;z-index:2147483647'; document.body.appendChild(el); }
    el.textContent = on ? 'PLUS ativo (Ctrl+Shift+P desliga • E: CSV • Shift+E: XLS)' : 'Ctrl+Shift+P ativa recursos PLUS';
  };
  setHint(false);

  // Keybinds
  document.addEventListener('keydown', (e) => {
    if (e.ctrlKey && e.shiftKey && (e.key==='P' || e.key==='p')) {
      PLUS_ON = !PLUS_ON; setHint(PLUS_ON); if (PLUS_ON && !WIRED) wireOnce();
    }
    if (!PLUS_ON) return;
    if (e.key === 'E' || e.key === 'e') {
      e.preventDefault();
      if (e.shiftKey) exportXLS(); else exportCSV();
    }
  });

  // CSV Export
  const escapeCSV = (v) => ('"' + String(v).replaceAll('"','""') + '"');
  const toCSV = (rows) => {
    const H = headers();
    const lines = [H.join(',')];
    rows.forEach(r => {
      const arr = H.map((_,i) => escapeCSV(cellText(r,i)));
      // motivo de perda (coluna extra)
      const leadKey = r.id || cellText(r,0);
      const reason = (lossMap[leadKey]?.reason || '').trim();
      lines.push(arr.concat([escapeCSV(reason)]).join(','));
    });
    return lines.join('\n');
  };
  const download = (content, filename, mime) => {
    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href=url; a.download=filename; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
  };
  const exportCSV = () => {
    const csv = toCSV(visibleRows());
    // inclui cabeçalho da coluna extra
    const withHeader = ['"'+headers().join('","')+'","Motivo de Perda"'].join('\n') + '\n' + csv.split('\n').slice(1).join('\n');
    download(withHeader, 'leads_filtrados.csv', 'text/csv;charset=utf-8;');
  };

  // XML Spreadsheet 2003 (.xls) — abre em qualquer Excel
  const exportXLS = () => {
    const H = headers();
    const rows = visibleRows().map(r => {
      const base = H.map((_,i)=>cellText(r,i));
      const key = r.id || cellText(r,0);
      base.push((lossMap[key]?.reason || '').trim());
      return base;
    });
    const allHeaders = H.concat(['Motivo de Perda']);
    const xmlHeader = '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>';
    const open = '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">';
    const close = '</Workbook>';
    const sheetOpen = '<Worksheet ss:Name="Leads"><Table>';
    const sheetClose = '</Table></Worksheet>';
    const row = (cells) => '<Row>' + cells.map(v => `<Cell><Data ss:Type="String">${String(v).replace(/[<&>]/g, s=>({'<':'&lt;','>':'&gt;','&':'&amp;'}[s]))}</Data></Cell>`).join('') + '</Row>';
    const xml = [xmlHeader, open, sheetOpen, row(allHeaders)]
      .concat(rows.map(r => row(r)))
      .concat([sheetClose, close]).join('');
    download(xml, 'leads_filtrados.xls', 'application/vnd.ms-excel');
  };

  function wireOnce(){
    if (WIRED) return; WIRED = true;
    // Dblclick editors (only when PLUS_ON)
    body.addEventListener('dblclick', (e) => {
      if (!PLUS_ON) return;
      const td = e.target.closest('td'); if (!td) return;
      const row = td.parentElement;
      const idx = Array.prototype.indexOf.call(row.children, td);

      // Próxima ação
      if (idx === columnIndex.nextStep) {
        const original = td.textContent.trim();
        td.setAttribute('contenteditable','true'); td.focus();
        const end = (commit)=>{ td.removeAttribute('contenteditable'); if(!commit) td.textContent=original; };
        td.addEventListener('keydown', (ev)=>{ if(ev.key==='Enter'){ev.preventDefault(); td.blur();} if(ev.key==='Escape'){ev.preventDefault(); end(false);} });
        td.addEventListener('blur', ()=> end(true), {once:true});
      }

      // Responsável (autocomplete com nomes existentes)
      if (idx === columnIndex.owner) {
        let dl = document.getElementById('owner-options'); if(!dl){dl=document.createElement('datalist');dl.id='owner-options';document.body.appendChild(dl);}
        const set = new Set(tableRows().map(r=>cellText(r,columnIndex.owner)).filter(Boolean));
        dl.innerHTML = Array.from(set).map(n=>`<option value="${n}">`).join('');
        const input=document.createElement('input'); input.type='text'; input.value=td.textContent.trim(); input.setAttribute('list','owner-options');
        Object.assign(input.style,{position:'absolute',inset:'0',width:'100%',height:'100%',background:'transparent',color:'inherit',font:'inherit',border:'1px solid rgba(255,255,255,.18)',borderRadius:'8px',padding:'4px 8px',outline:'none'});
        const prev=td.style.position; if(!prev) td.style.position='relative'; td.appendChild(input); input.focus(); input.select();
        const cleanup=()=>{input.remove(); td.style.position=prev;};
        const commit=()=>{const v=input.value.trim(); if(v) td.textContent=v;};
        input.addEventListener('keydown',(ev)=>{ if(ev.key==='Enter'){ev.preventDefault(); commit(); cleanup();} if(ev.key==='Escape'){ev.preventDefault(); cleanup();} });
        input.addEventListener('blur', ()=>{commit(); cleanup();}, {once:true});
      }

      // Status + Motivo de perda
      if (idx === columnIndex.status) {
        const current = rowStatus(row);
        const label = window.prompt('Status (Pendente | Em negociação | Ganho | Perdido):', current);
        if (!label) return;
        let pill = td.querySelector('.status'); if (!pill){ pill=document.createElement('span'); pill.className='status'; td.innerHTML=''; td.appendChild(pill); }
        pill.className='status';
        if (/^Pendente$/i.test(label)) { pill.classList.add('pendente'); pill.textContent='Pendente'; row.dataset.status='Pendente'; }
        else if (/negocia/i.test(label)) { pill.classList.add('negociacao'); pill.textContent='Em negociação'; row.dataset.status='Em negociação'; }
        else if (/^Ganho$/i.test(label)) { pill.classList.add('ganho'); pill.innerHTML='<i class="ri-check-line" aria-hidden="true"></i> Ganho'; row.dataset.status='Ganho'; }
        else if (/Perdido/i.test(label)) { pill.classList.add('perdido'); pill.innerHTML='<i class="ri-close-line" aria-hidden="true"></i> Perdido'; row.dataset.status='Perdido'; }
        else { pill.textContent=label.trim(); row.dataset.status=label.trim(); }

        // motivo de perda
        const leadKey = row.id || cellText(row, 0);
        if (/Perdido/i.test(row.dataset.status)) {
          const prev = lossMap[leadKey]?.reason || '';
          const reason = window.prompt('Motivo de perda (opcional):', prev) || '';
          lossMap[leadKey] = { reason };
          saveLoss();
        }
        // atualizar cards
        refreshLossCards();
      }
    });
  }

  // Sincroniza motivo de perda no Painel de Observações (sem alterar o template)
  const refreshLossCards = () => {
    const grid = document.getElementById('observations-grid') || document.querySelector('.observations #observations-grid');
    if (!grid) return;
    Array.from(grid.querySelectorAll('.observation-card')).forEach((card) => {
      const nameEl = card.querySelector('.lead-title, header h4, .observation-title');
      const id = card.getAttribute('data-row-id') || (nameEl ? nameEl.textContent.trim() : '');
      const old = card.querySelector('.observation-loss'); if (old) old.remove();
      // tenta status no card; se não houver, consulta a tabela
      let statusText = (card.querySelector('.status')?.textContent || '').trim();
      if (!statusText) {
        const row = tableRows().find(r => (r.id||cellText(r,0)) === id || cellText(r,0) === id);
        statusText = row ? rowStatus(row) : '';
      }
      const reason = (lossMap[id]?.reason || '').trim();
      if (/Perdido/i.test(statusText) && reason) {
        const div = document.createElement('div');
        div.className = 'observation-loss';
        div.style.cssText = 'margin:8px 0 4px;font-size:12px;opacity:.9;color:var(--muted,#94a3b8)';
        div.innerHTML = '<strong>Motivo de perda:</strong> ' + reason.replace(/[<>]/g,'');
        card.appendChild(div);
      }
    });
  };
})();