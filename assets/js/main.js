(() => {
	const STATE_STORAGE_KEY = 'comercial-demo-state-v1';
	const NOTES_STORAGE_KEY = 'comercial-demo-notes-v1';
	const LOSS_REASONS_STORAGE_KEY = 'comercial-loss-reasons-v1';
	const STALLED_THRESHOLD_DAYS = 7;
	const VISIBLE_ROW_LIMIT = 5;
	const NOTE_SAVE_DELAY = 240;
	const FILTER_DELAY = 160;
	const RESIZE_DELAY = 180;
	const MODAL_TRANSITION_DURATION = 240;
	const ROW_FOCUS_DURATION = 1600;
	const BUTTON_FEEDBACK_DURATION = 1800;

	const COLUMN = {
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

	const SORT_ICONS = {
		neutral: '\u2195',
		asc: '\u2191',
		desc: '\u2193'
	};

	const ARIA_SORT_VALUES = {
		asc: 'ascending',
		desc: 'descending'
	};

	const RANKING_BADGES = [
		{ label: 'Top 1', icon: 'ri-trophy-line' },
		{ label: 'Top 2', icon: 'ri-medal-line' },
		{ label: 'Top 3', icon: 'ri-award-line' }
	];

	const STATUS_STYLES = [
		{ pattern: /^pendente$/i, label: 'Pendente', className: 'pendente', icon: '' },
		{ pattern: /negocia/i, label: 'Em negociação', className: 'negociacao', icon: '' },
		{ pattern: /^ganho$/i, label: 'Ganho', className: 'ganho', icon: 'ri-check-line' },
		{ pattern: /perdido/i, label: 'Perdido', className: 'perdido', icon: 'ri-close-line' }
	];

	const PLUS_HINT_TEXT = {
		active: 'PLUS ativo · duplo clique edita próxima ação, responsável e status · Ctrl + Shift + P desliga',
		inactive: 'Ctrl + Shift + P ativa recursos PLUS'
	};

	const HTML_ESCAPE_MAP = {
		'&': '&amp;',
		'<': '&lt;',
		'>': '&gt;',
		'"': '&quot;',
		"'": '&#39;'
	};

	const KPI_DEFINITIONS = [
		{ key: 'total', icon: 'ri-group-line', label: 'Leads', initialValue: '0', meta: 'Visíveis' },
		{ key: 'negociacao', icon: 'ri-shake-hands-line', label: 'Negociação', initialValue: '0', meta: 'Em andamento' },
		{ key: 'ganho', icon: 'ri-trophy-line', label: 'Ganhos', initialValue: '0', meta: 'Fechados' },
		{ key: 'perdido', icon: 'ri-close-circle-line', label: 'Perdidos', initialValue: '0', meta: 'Encerrados' },
		{ key: 'tx', icon: 'ri-pie-chart-line', label: 'Conversão', initialValue: '0%', meta: 'Taxa de ganho' },
		{ key: 'const', icon: 'ri-time-line', label: 'Constância', initialValue: '0d', meta: 'Dias médios' },
		{ key: 'positivas', icon: 'ri-mail-check-line', label: 'Resp. positivas', initialValue: '0', meta: 'Acumulado' },
		{ key: 'acoes', icon: 'ri-bar-chart-line', label: 'Atividades', initialValue: '0', meta: 'Média por lead' }
	];

	const TOOLBAR_TEMPLATE = `
		<div class="kpis" role="region" aria-label="Indicadores do funil">
			${KPI_DEFINITIONS.map(({ key, icon, label, initialValue, meta }) => `
				<div class="kpi" data-kpi="${key}">
					<div class="kpi-icon" aria-hidden="true"><i class="${icon}"></i></div>
					<div class="kpi-body">
						<span class="kpi-label">${label}</span>
						<strong class="kpi-value">${initialValue}</strong>
						<span class="kpi-meta">${meta}</span>
					</div>
				</div>
			`).join('')}
		</div>
		<div class="filters" role="region" aria-label="Filtros">
			<div class="group">
				<i class="ri-search-line" aria-hidden="true"></i>
				<input id="f-q" type="search" placeholder="Buscar lead, empresa, responsável ou próxima ação" aria-label="Buscar leads" autocomplete="off">
			</div>
			<select id="f-etapa" aria-label="Filtrar por etapa"><option value="">Etapa (todas)</option></select>
			<select id="f-status" aria-label="Filtrar por status"><option value="">Status (todos)</option></select>
			<select id="f-resp" aria-label="Filtrar por responsável"><option value="">Responsável (todos)</option></select>
			<button type="button" id="btn-clear" class="btn ghost" title="Limpar filtros (F)">Limpar</button>
			<div class="split" aria-hidden="true"></div>
			<button type="button" id="btn-highlight" class="btn primary" title="Destaques do semestre (G)" aria-haspopup="dialog" aria-expanded="false">Destaques</button>
			<button type="button" id="btn-export" class="btn" title="Exportar CSV (E) ou Excel (Shift + E)">Exportar</button>
			<button type="button" id="btn-copy" class="btn" title="Copiar tabela" data-label="Copiar">Copiar</button>
			<span class="filters-hotkeys" aria-hidden="true"><kbd>F</kbd> limpa · <kbd>G</kbd> destaques · <kbd>E</kbd> exporta</span>
			<span class="result-count" aria-live="polite"></span>
		</div>
	`;

	const INSIGHTS_TEMPLATE = `
		<div class="insights-header">
			<span class="insights-title"><i class="ri-lightbulb-line" aria-hidden="true"></i> Insights rápidos</span>
			<button type="button" class="insights-toggle" aria-expanded="true">Ocultar</button>
		</div>
		<div class="insights-body"></div>
	`;

	const HIGHLIGHT_TEMPLATE = `
		<article class="highlight-card" role="dialog" aria-modal="true" aria-labelledby="highlight-dialog-title">
			<header class="highlight-head">
				<div class="highlight-headline">
					<h2 class="highlight-title" id="highlight-dialog-title"><i class="ri-rocket-line" aria-hidden="true"></i> Top 3 do mês</h2>
					<span class="highlight-period"></span>
				</div>
				<button type="button" class="highlight-close" aria-label="Fechar painel de destaques"><i class="ri-close-line" aria-hidden="true"></i></button>
			</header>
			<p class="highlight-caption">Ranking baseado em respostas positivas, atividades e constância.</p>
			<div class="highlight-body"></div>
		</article>
	`;

	const findElement = (selector, scope = document) => scope.querySelector(selector);

	const table = findElement('.tabela-comercial');
	const card = findElement('.card');
	const tableHead = table?.tHead;
	const tableBody = table?.tBodies[0];

	if (!table || !card || !tableHead || !tableBody) {
		return;
	}

	const tableWrapper = findElement('.table-responsive', card);
	const observationsList = findElement('#observations-grid');
	const headerCells = Array.from(tableHead.rows[0]?.cells || []);
	const toolbar = findElement('.toolbar', card) || document.createElement('div');
	const insightsPanel = document.createElement('section');
	const highlightModal = document.createElement('div');
	const plusHint = document.createElement('div');
	const ownerOptions = document.createElement('datalist');

	toolbar.className = 'toolbar';
	toolbar.innerHTML = TOOLBAR_TEMPLATE;
	insightsPanel.className = 'insights-panel is-hidden';
	insightsPanel.setAttribute('aria-label', 'Insights rápidos');
	insightsPanel.innerHTML = INSIGHTS_TEMPLATE;
	highlightModal.className = 'highlight-modal is-hidden';
	highlightModal.innerHTML = HIGHLIGHT_TEMPLATE;
	plusHint.className = 'plus-hint';
	plusHint.setAttribute('aria-hidden', 'true');
	ownerOptions.id = 'owner-options';

	const searchInput = findElement('#f-q', toolbar);
	const stageFilter = findElement('#f-etapa', toolbar);
	const statusFilter = findElement('#f-status', toolbar);
	const ownerFilter = findElement('#f-resp', toolbar);
	const clearButton = findElement('#btn-clear', toolbar);
	const highlightButton = findElement('#btn-highlight', toolbar);
	const exportButton = findElement('#btn-export', toolbar);
	const copyButton = findElement('#btn-copy', toolbar);
	const resultCount = findElement('.result-count', toolbar);
	const selectFilters = [stageFilter, statusFilter, ownerFilter];
	const insightsBody = findElement('.insights-body', insightsPanel);
	const insightsToggle = findElement('.insights-toggle', insightsPanel);
	const highlightBody = findElement('.highlight-body', highlightModal);
	const highlightPeriod = findElement('.highlight-period', highlightModal);
	const highlightCloseButton = findElement('.highlight-close', highlightModal);

	const reducedMotionQuery = window.matchMedia('(prefers-reduced-motion: reduce)');
	const browserStorage = (() => {
		try {
			const testKey = '__storage-test__';
			window.localStorage.setItem(testKey, testKey);
			window.localStorage.removeItem(testKey);
			return window.localStorage;
		} catch {
			return null;
		}
	})();

	const observationItems = new Map();
	const openObservationIds = new Set();
	const rowFocusTimers = new WeakMap();
	let observationNotes = {};
	let lossReasons = {};
	let lastHighlightRows = [];
	let observationsEmptyMessage = null;
	let isPlusModeActive = false;
	let modalHideTimer = 0;
	let buttonFeedbackTimer = 0;

	const escapeHtml = (value = '') => String(value).replace(/[&<>"']/g, (character) => HTML_ESCAPE_MAP[character]);

	const normalizeText = (value = '') => String(value)
		.normalize('NFD')
		.replace(/[\u0300-\u036f]/g, '')
		.toLowerCase();

	const slugify = (value) => normalizeText(value)
		.replace(/[^a-z0-9]+/g, '-')
		.replace(/^-+|-+$/g, '') || 'referencia';

	const toInteger = (value) => {
		const numberMatch = String(value).match(/-?\d+/);
		return numberMatch ? parseInt(numberMatch[0], 10) : 0;
	};

	const debounce = (callback, delay) => {
		let timer;
		return (...callbackArguments) => {
			clearTimeout(timer);
			timer = setTimeout(() => callback(...callbackArguments), delay);
		};
	};

	const formatAverage = (value) => {
		if (value >= 10) {
			return String(Math.round(value));
		}
		return value.toFixed(1).replace('.', ',').replace(/,0$/, '');
	};

	const formatCurrentMonth = () => {
		const monthLabel = new Intl.DateTimeFormat('pt-BR', { month: 'long', year: 'numeric' }).format(new Date());
		return monthLabel.charAt(0).toUpperCase() + monthLabel.slice(1);
	};

	const cellText = (row, columnIndex) => (row.cells[columnIndex]?.textContent || '').replace(/\s+/g, ' ').trim();
	const getAllRows = () => Array.from(tableBody.rows);
	const getVisibleRows = () => getAllRows().filter((row) => !row.hidden);
	const getLeadName = (row) => cellText(row, COLUMN.lead);
	const getCompany = (row) => cellText(row, COLUMN.company);
	const getStage = (row) => row.dataset.etapa || cellText(row, COLUMN.stage);
	const getStatus = (row) => row.dataset.status || cellText(row, COLUMN.status);
	const getOwner = (row) => cellText(row, COLUMN.owner);
	const getNextStep = (row) => cellText(row, COLUMN.nextStep);
	const getPositiveReplies = (row) => toInteger(cellText(row, COLUMN.positiveReplies));
	const getCadenceDays = (row) => toInteger(cellText(row, COLUMN.cadence));
	const isLostLead = (row) => /perdido/i.test(getStatus(row));

	const getActivitiesCount = (row) => cellText(row, COLUMN.activities)
		.split(/[,+]/)
		.reduce((total, activity) => total + toInteger(activity), 0);

	const getLeadScore = (row) => getPositiveReplies(row) * 3 + getActivitiesCount(row) + getCadenceDays(row) / 2;

	const getLastContactTime = (row) => {
		const timeElement = row.cells[COLUMN.lastContact]?.querySelector('time');
		if (timeElement?.dateTime) {
			return Date.parse(timeElement.dateTime);
		}
		const [day, month, year] = cellText(row, COLUMN.lastContact).split('/').map(Number);
		return new Date(year, month - 1, day).getTime() || 0;
	};

	const getLossReason = (row) => {
		const savedReason = lossReasons[row.id]?.reason || lossReasons[getLeadName(row)]?.reason || '';
		return savedReason.trim();
	};

	const getActiveLossReason = (row) => (isLostLead(row) ? getLossReason(row) : '');

	const sortRowsDescending = (rows, readValue) => [...rows].sort((firstRow, secondRow) => readValue(secondRow) - readValue(firstRow));

	const SORT_VALUE_READERS = {
		[COLUMN.lastContact]: getLastContactTime,
		[COLUMN.positiveReplies]: getPositiveReplies,
		[COLUMN.activities]: getActivitiesCount,
		[COLUMN.cadence]: getCadenceDays
	};

	const readStoredJson = (storageKey) => {
		if (!browserStorage) {
			return {};
		}
		try {
			const parsedValue = JSON.parse(browserStorage.getItem(storageKey) || '{}');
			return parsedValue && typeof parsedValue === 'object' ? parsedValue : {};
		} catch {
			return {};
		}
	};

	const writeStoredJson = (storageKey, value) => {
		if (!browserStorage) {
			return;
		}
		try {
			browserStorage.setItem(storageKey, JSON.stringify(value));
		} catch {
			return;
		}
	};

	const saveNotesLater = debounce(() => writeStoredJson(NOTES_STORAGE_KEY, observationNotes), NOTE_SAVE_DELAY);

	const saveState = () => writeStoredJson(STATE_STORAGE_KEY, {
		query: searchInput.value,
		stage: stageFilter.value,
		status: statusFilter.value,
		owner: ownerFilter.value,
		sortColumn: table.dataset.sortColumn ?? null,
		sortDirection: table.dataset.sortDirection ?? null
	});

	const restoreState = () => {
		const savedState = readStoredJson(STATE_STORAGE_KEY);
		const sortColumn = parseInt(savedState.sortColumn, 10);
		searchInput.value = savedState.query || '';
		stageFilter.value = savedState.stage || '';
		statusFilter.value = savedState.status || '';
		ownerFilter.value = savedState.owner || '';
		if (headerCells[sortColumn]) {
			sortTable(sortColumn, savedState.sortDirection === 'desc' ? 'desc' : 'asc');
		}
	};

	const assignRowIds = () => {
		const usedIds = new Set();
		getAllRows().forEach((row) => {
			const baseId = `lead-${slugify(getLeadName(row))}`;
			let rowId = row.id && !usedIds.has(row.id) ? row.id : baseId;
			let suffix = 1;
			while (usedIds.has(rowId)) {
				rowId = `${baseId}-${suffix}`;
				suffix += 1;
			}
			row.id = rowId;
			usedIds.add(rowId);
		});
	};

	const fillSelectOptions = (selectElement, values) => {
		const selectedValue = selectElement.value;
		const placeholderOption = selectElement.options[0];
		const uniqueValues = [...new Set(values.filter(Boolean))].sort((first, second) => first.localeCompare(second, 'pt-BR'));
		selectElement.replaceChildren(placeholderOption, ...uniqueValues.map((value) => new Option(value, value)));
		selectElement.value = uniqueValues.includes(selectedValue) ? selectedValue : '';
	};

	const refreshFilterOptions = () => {
		const rows = getAllRows();
		fillSelectOptions(stageFilter, rows.map(getStage));
		fillSelectOptions(statusFilter, rows.map(getStatus));
		fillSelectOptions(ownerFilter, rows.map(getOwner));
	};

	const rowMatchesFilters = (row, filters) => {
		const searchableText = normalizeText([getLeadName(row), getCompany(row), getOwner(row), getNextStep(row)].join(' '));
		return (!filters.query || searchableText.includes(filters.query))
			&& (!filters.stage || getStage(row) === filters.stage)
			&& (!filters.status || getStatus(row) === filters.status)
			&& (!filters.owner || getOwner(row) === filters.owner);
	};

	const applyFilters = () => {
		const currentFilters = {
			query: normalizeText(searchInput.value.trim()),
			stage: stageFilter.value,
			status: statusFilter.value,
			owner: ownerFilter.value
		};
		getAllRows().forEach((row) => {
			row.hidden = !rowMatchesFilters(row, currentFilters);
		});
		const visibleRows = getVisibleRows();
		updateResultCount(visibleRows.length);
		updateKpis(visibleRows);
		updateInsights(visibleRows);
		updateTableHeight();
		updateObservationVisibility();
		closeHighlightModal();
	};

	const applyFiltersAndSave = () => {
		applyFilters();
		saveState();
	};

	const clearFilters = () => {
		searchInput.value = '';
		selectFilters.forEach((selectElement) => {
			selectElement.value = '';
		});
		applyFiltersAndSave();
	};

	const prepareSortableHeaders = () => {
		headerCells.forEach((headerCell) => {
			const sortArrow = document.createElement('span');
			headerCell.dataset.label = headerCell.textContent.trim();
			headerCell.classList.add('sortable');
			headerCell.tabIndex = 0;
			headerCell.setAttribute('aria-sort', 'none');
			sortArrow.className = 'arrow';
			sortArrow.setAttribute('aria-hidden', 'true');
			sortArrow.textContent = SORT_ICONS.neutral;
			headerCell.append(sortArrow);
		});
	};

	const sortTable = (columnIndex, direction) => {
		const readValue = SORT_VALUE_READERS[columnIndex] || ((row) => cellText(row, columnIndex));
		const directionFactor = direction === 'asc' ? 1 : -1;
		const sortedRows = getAllRows().sort((firstRow, secondRow) => {
			const firstValue = readValue(firstRow);
			const secondValue = readValue(secondRow);
			const comparison = typeof firstValue === 'number'
				? firstValue - secondValue
				: firstValue.localeCompare(secondValue, 'pt-BR', { sensitivity: 'base' });
			return comparison * directionFactor;
		});

		tableBody.append(...sortedRows);
		table.dataset.sortColumn = String(columnIndex);
		table.dataset.sortDirection = direction;

		headerCells.forEach((headerCell, headerIndex) => {
			const isActiveColumn = headerIndex === columnIndex;
			headerCell.classList.toggle('sort-asc', isActiveColumn && direction === 'asc');
			headerCell.classList.toggle('sort-desc', isActiveColumn && direction === 'desc');
			headerCell.setAttribute('aria-sort', isActiveColumn ? ARIA_SORT_VALUES[direction] : 'none');
			findElement('.arrow', headerCell).textContent = isActiveColumn ? SORT_ICONS[direction] : SORT_ICONS.neutral;
		});

		reorderObservationItems();
	};

	const sortByHeader = (headerCell) => {
		const columnIndex = headerCells.indexOf(headerCell);
		if (columnIndex < 0) {
			return;
		}
		const isSameColumn = table.dataset.sortColumn === String(columnIndex);
		const nextDirection = isSameColumn && table.dataset.sortDirection === 'asc' ? 'desc' : 'asc';
		sortTable(columnIndex, nextDirection);
		saveState();
	};

	const setKpiValue = (kpiKey, value) => {
		findElement(`.kpi[data-kpi="${kpiKey}"] .kpi-value`, toolbar).textContent = String(value);
	};

	const updateKpis = (visibleRows) => {
		const visibleCount = visibleRows.length;
		const countByStatus = (statusPattern) => visibleRows.filter((row) => statusPattern.test(getStatus(row))).length;
		const sumOf = (readValue) => visibleRows.reduce((total, row) => total + readValue(row), 0);
		const averageOf = (readValue) => (visibleCount ? sumOf(readValue) / visibleCount : 0);
		const wonCount = countByStatus(/ganho/i);
		const conversionRate = visibleCount ? Math.round((wonCount / visibleCount) * 100) : 0;

		setKpiValue('total', visibleCount);
		setKpiValue('negociacao', countByStatus(/negocia/i));
		setKpiValue('ganho', wonCount);
		setKpiValue('perdido', countByStatus(/perdido/i));
		setKpiValue('tx', `${conversionRate}%`);
		setKpiValue('const', `${Math.round(averageOf(getCadenceDays))}d`);
		setKpiValue('positivas', sumOf(getPositiveReplies));
		setKpiValue('acoes', formatAverage(averageOf(getActivitiesCount)));
	};

	const updateResultCount = (visibleCount) => {
		const countLabels = {
			0: 'Nenhum lead visível',
			1: '1 lead visível'
		};
		resultCount.textContent = countLabels[visibleCount] || `${visibleCount} leads visíveis`;
	};

	const renderInsightLeads = (rows, describeRow) => {
		if (!rows.length) {
			return '<li class="insight-empty">Nenhum lead</li>';
		}
		return rows.slice(0, 3).map((row) => `
			<li>
				<span class="lead-name">${escapeHtml(getLeadName(row))}</span>
				<span class="lead-meta">${escapeHtml(describeRow(row))}</span>
			</li>
		`).join('');
	};

	const renderInsightCard = (title, highlightValue, rows, describeRow) => `
		<article class="insight-card">
			<header>
				<strong>${title}</strong>
				<span class="insight-count">${escapeHtml(highlightValue)}</span>
			</header>
			<ul class="insight-list">${renderInsightLeads(rows, describeRow)}</ul>
		</article>
	`;

	const updateInsights = (visibleRows) => {
		insightsPanel.classList.toggle('is-hidden', !visibleRows.length);
		if (!visibleRows.length) {
			insightsBody.innerHTML = '';
			return;
		}

		const describeReplies = (row) => `${getPositiveReplies(row)} resp.`;
		const describeActivities = (row) => `${getActivitiesCount(row)} atv.`;
		const describeCadence = (row) => `${getCadenceDays(row)}d sem contato`;
		const describeScore = (row) => `${Math.round(getLeadScore(row))} pts`;
		const rowsByReplies = sortRowsDescending(visibleRows, getPositiveReplies);
		const rowsByActivities = sortRowsDescending(visibleRows, getActivitiesCount);
		const stalledRows = sortRowsDescending(visibleRows.filter((row) => getCadenceDays(row) >= STALLED_THRESHOLD_DAYS), getCadenceDays);
		const topRows = sortRowsDescending(lastHighlightRows.length ? lastHighlightRows : rowsByReplies.slice(0, 3), getLeadScore);

		insightsBody.innerHTML = `
			<div class="insights-grid">
				${renderInsightCard('Respostas positivas', describeReplies(rowsByReplies[0]), rowsByReplies, describeReplies)}
				${renderInsightCard('Atividades registradas', describeActivities(rowsByActivities[0]), rowsByActivities, describeActivities)}
				${renderInsightCard('Leads estagnados', stalledRows.length, stalledRows, describeCadence)}
				${renderInsightCard('Último Top 3', topRows.length, topRows, describeScore)}
			</div>
		`;
	};

	const toggleInsights = () => {
		const isCollapsed = insightsPanel.classList.toggle('is-collapsed');
		insightsToggle.textContent = isCollapsed ? 'Mostrar' : 'Ocultar';
		insightsToggle.setAttribute('aria-expanded', String(!isCollapsed));
	};

	const updateTableHeight = () => {
		if (!tableWrapper) {
			return;
		}
		const visibleRows = getVisibleRows();
		const rowHeight = visibleRows[0]?.getBoundingClientRect().height || 0;
		if (visibleRows.length <= VISIBLE_ROW_LIMIT || !rowHeight) {
			tableWrapper.style.removeProperty('max-height');
			return;
		}
		const captionHeight = table.caption?.getBoundingClientRect().height || 0;
		const headerHeight = tableHead.getBoundingClientRect().height;
		tableWrapper.style.maxHeight = `${Math.round(captionHeight + headerHeight + rowHeight * VISIBLE_ROW_LIMIT)}px`;
	};

	const renderObservationPill = (icon, value) => {
		if (!value) {
			return '';
		}
		return `<span class="observation-pill"><i class="${icon}" aria-hidden="true"></i>${escapeHtml(value)}</span>`;
	};

	const getStoredNote = (rowId, leadName) => {
		if (typeof observationNotes[rowId] === 'string') {
			return observationNotes[rowId];
		}
		if (typeof observationNotes[leadName] !== 'string') {
			return '';
		}
		observationNotes[rowId] = observationNotes[leadName];
		delete observationNotes[leadName];
		saveNotesLater();
		return observationNotes[rowId];
	};

	const createObservationItem = (row) => {
		const rowId = row.id;
		const leadName = getLeadName(row);
		const noteId = `note-${rowId}`;
		const storedNote = getStoredNote(rowId, leadName);
		const lossReason = getActiveLossReason(row);
		const isOpen = openObservationIds.has(rowId);
		const observationItem = document.createElement('li');
		const metaPills = [
			['ri-map-pin-2-line', getStage(row)],
			['ri-flag-2-line', getStatus(row)],
			['ri-calendar-event-line', cellText(row, COLUMN.lastContact)],
			['ri-time-line', cellText(row, COLUMN.cadence)],
			['ri-user-voice-line', getOwner(row)],
			['ri-compass-3-line', getNextStep(row)]
		].map(([icon, value]) => renderObservationPill(icon, value)).join('');
		const lossReasonMarkup = lossReason
			? `<p class="observation-loss"><strong>Motivo de perda:</strong> ${escapeHtml(lossReason)}</p>`
			: '';

		observationItem.className = 'observation-item';
		observationItem.classList.toggle('is-open', isOpen);
		observationItem.classList.toggle('has-note', Boolean(storedNote.trim()));
		observationItem.dataset.leadId = rowId;
		observationItem.innerHTML = `
			<article class="observation-card">
				<header class="observation-head">
					<div class="observation-title" role="button" tabindex="0" aria-expanded="${isOpen}">
						<span class="observation-lead-name">${escapeHtml(leadName)}</span>
						<span class="observation-company">${escapeHtml(getCompany(row))}</span>
					</div>
					<button type="button" class="observation-jump" data-target="${rowId}">
						<i class="ri-focus-2-line" aria-hidden="true"></i>
						Ver na tabela
					</button>
				</header>
				<div class="observation-meta">${metaPills}</div>
				${lossReasonMarkup}
				<div class="observation-note-group">
					<label class="observation-note-label" for="${noteId}">Observações</label>
					<textarea class="observation-note" id="${noteId}" data-lead-id="${rowId}" placeholder="Anote direcionamentos e percepções sobre ${escapeHtml(leadName)}"></textarea>
				</div>
			</article>
		`;
		findElement('.observation-note', observationItem).value = storedNote;
		return observationItem;
	};

	const buildObservationList = () => {
		if (!observationsList) {
			return;
		}
		const rows = getAllRows().filter(getLeadName);
		observationItems.clear();
		rows.forEach((row) => observationItems.set(row.id, createObservationItem(row)));
		observationsEmptyMessage = document.createElement('li');
		observationsEmptyMessage.className = 'observation-empty';
		observationsEmptyMessage.textContent = rows.length ? 'Nenhum lead visível com os filtros atuais.' : 'Nenhum lead cadastrado.';
		observationsList.replaceChildren(...observationItems.values(), observationsEmptyMessage);
		updateObservationVisibility();
	};

	const reorderObservationItems = () => {
		if (!observationsList || !observationsEmptyMessage) {
			return;
		}
		const orderedItems = getAllRows().map((row) => observationItems.get(row.id)).filter(Boolean);
		observationsList.replaceChildren(...orderedItems, observationsEmptyMessage);
	};

	const updateObservationVisibility = () => {
		if (!observationsEmptyMessage) {
			return;
		}
		let visibleCount = 0;
		observationItems.forEach((observationItem, rowId) => {
			const isVisible = document.getElementById(rowId)?.hidden === false;
			observationItem.hidden = !isVisible;
			visibleCount += Number(isVisible);
		});
		observationsEmptyMessage.hidden = visibleCount > 0;
	};

	const toggleObservationItem = (observationItem) => {
		const rowId = observationItem.dataset.leadId;
		const isOpening = !observationItem.classList.contains('is-open');
		observationItem.classList.toggle('is-open', isOpening);
		findElement('.observation-title', observationItem).setAttribute('aria-expanded', String(isOpening));
		if (isOpening) {
			openObservationIds.add(rowId);
		} else {
			openObservationIds.delete(rowId);
		}
	};

	const saveObservationNote = (noteField) => {
		const rowId = noteField.dataset.leadId;
		const hasNote = Boolean(noteField.value.trim());
		if (hasNote) {
			observationNotes[rowId] = noteField.value;
		} else {
			delete observationNotes[rowId];
		}
		noteField.closest('.observation-item').classList.toggle('has-note', hasNote);
		saveNotesLater();
	};

	const focusTableRow = (rowId) => {
		const targetRow = document.getElementById(rowId);
		if (!targetRow) {
			return;
		}
		clearTimeout(rowFocusTimers.get(targetRow));
		targetRow.classList.remove('note-focus');
		targetRow.scrollIntoView({ behavior: reducedMotionQuery.matches ? 'auto' : 'smooth', block: 'center' });
		requestAnimationFrame(() => targetRow.classList.add('note-focus'));
		rowFocusTimers.set(targetRow, setTimeout(() => targetRow.classList.remove('note-focus'), ROW_FOCUS_DURATION));
	};

	const renderHighlightList = (rankedRows) => {
		if (!rankedRows.length) {
			return '<p class="highlight-empty">Nenhum destaque disponível.</p>';
		}
		const listItems = rankedRows.map(({ row, score }, position) => {
			const { label, icon } = RANKING_BADGES[position];
			return `
				<li>
					<span class="highlight-rank"><i class="${icon}" aria-hidden="true"></i>${label}</span>
					<div class="highlight-lead">
						<strong>${escapeHtml(getLeadName(row))}</strong>
						<span class="highlight-score">${Math.round(score)} pts</span>
					</div>
					<div class="highlight-meta">
						<span><i class="ri-compass-line" aria-hidden="true"></i>${escapeHtml(getStage(row))}</span>
						<span><i class="ri-mail-check-line" aria-hidden="true"></i>${getPositiveReplies(row)} resp.</span>
						<span><i class="ri-bar-chart-line" aria-hidden="true"></i>${escapeHtml(cellText(row, COLUMN.activities))}</span>
						<span><i class="ri-time-line" aria-hidden="true"></i>${getCadenceDays(row)}d de constância</span>
						<span><i class="ri-user-line" aria-hidden="true"></i>${escapeHtml(getOwner(row))}</span>
					</div>
				</li>
			`;
		}).join('');
		return `<ol class="highlight-list">${listItems}</ol>`;
	};

	const isHighlightModalOpen = () => highlightModal.classList.contains('is-visible');

	const openHighlightModal = (rankedRows) => {
		clearTimeout(modalHideTimer);
		highlightBody.innerHTML = renderHighlightList(rankedRows);
		highlightPeriod.textContent = formatCurrentMonth();
		highlightModal.classList.remove('is-hidden');
		void highlightModal.offsetWidth;
		highlightModal.classList.add('is-visible');
		document.body.classList.add('highlight-open');
		card.classList.add('modal-blur');
		highlightButton.setAttribute('aria-expanded', 'true');
		highlightCloseButton.focus({ preventScroll: true });
	};

	const closeHighlightModal = () => {
		if (!isHighlightModalOpen()) {
			return;
		}
		highlightModal.classList.remove('is-visible');
		document.body.classList.remove('highlight-open');
		card.classList.remove('modal-blur');
		highlightButton.setAttribute('aria-expanded', 'false');
		highlightButton.focus({ preventScroll: true });
		modalHideTimer = setTimeout(() => highlightModal.classList.add('is-hidden'), MODAL_TRANSITION_DURATION);
	};

	const showHighlights = () => {
		const rankedRows = getVisibleRows()
			.map((row) => ({ row, score: getLeadScore(row) }))
			.sort((firstEntry, secondEntry) => secondEntry.score - firstEntry.score)
			.slice(0, RANKING_BADGES.length);
		lastHighlightRows = rankedRows.map(({ row }) => row);
		getAllRows().forEach((row) => row.classList.toggle('topline', lastHighlightRows.includes(row)));
		updateInsights(getVisibleRows());
		openHighlightModal(rankedRows);
	};

	const getExportData = () => {
		const headerLabels = [...headerCells.map((headerCell) => headerCell.dataset.label), 'Motivo de perda'];
		const rowValues = getVisibleRows().map((row) => [
			...Array.from(row.cells, (cell, columnIndex) => cellText(row, columnIndex)),
			getActiveLossReason(row)
		]);
		return [headerLabels, ...rowValues];
	};

	const downloadFile = (content, fileName, mimeType) => {
		const fileUrl = URL.createObjectURL(new Blob([content], { type: mimeType }));
		const downloadLink = document.createElement('a');
		downloadLink.href = fileUrl;
		downloadLink.download = fileName;
		document.body.append(downloadLink);
		downloadLink.click();
		downloadLink.remove();
		setTimeout(() => URL.revokeObjectURL(fileUrl), 1000);
	};

	const exportCsv = () => {
		const csvContent = getExportData()
			.map((rowValues) => rowValues.map((value) => `"${value.replace(/"/g, '""')}"`).join(';'))
			.join('\r\n');
		downloadFile(`\uFEFF${csvContent}`, 'leads-filtrados.csv', 'text/csv;charset=utf-8;');
	};

	const exportExcel = () => {
		const xmlRows = getExportData()
			.map((rowValues) => `<Row>${rowValues.map((value) => `<Cell><Data ss:Type="String">${escapeHtml(value)}</Data></Cell>`).join('')}</Row>`)
			.join('');
		const workbook = [
			'<?xml version="1.0" encoding="UTF-8"?>',
			'<?mso-application progid="Excel.Sheet"?>',
			'<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">',
			`<Worksheet ss:Name="Leads"><Table>${xmlRows}</Table></Worksheet>`,
			'</Workbook>'
		].join('');
		downloadFile(workbook, 'leads-filtrados.xls', 'application/vnd.ms-excel');
	};

	const exportTable = (useExcelFormat) => {
		if (useExcelFormat) {
			exportExcel();
		} else {
			exportCsv();
		}
	};

	const showButtonFeedback = (button, message) => {
		clearTimeout(buttonFeedbackTimer);
		button.textContent = message;
		buttonFeedbackTimer = setTimeout(() => {
			button.textContent = button.dataset.label;
		}, BUTTON_FEEDBACK_DURATION);
	};

	const copyTable = async () => {
		const clipboardText = getExportData().map((rowValues) => rowValues.join('\t')).join('\n');
		try {
			await navigator.clipboard.writeText(clipboardText);
			showButtonFeedback(copyButton, 'Copiado!');
		} catch {
			showButtonFeedback(copyButton, 'Não foi possível copiar');
		}
	};

	const updatePlusHint = () => {
		plusHint.textContent = isPlusModeActive ? PLUS_HINT_TEXT.active : PLUS_HINT_TEXT.inactive;
	};

	const togglePlusMode = () => {
		isPlusModeActive = !isPlusModeActive;
		updatePlusHint();
	};

	const refreshAfterEdit = () => {
		refreshFilterOptions();
		buildObservationList();
		applyFiltersAndSave();
	};

	const editNextStep = (cell) => {
		const originalText = cell.textContent.trim();
		let shouldSave = true;
		const handleEditorKeys = (event) => {
			if (event.key !== 'Enter' && event.key !== 'Escape') {
				return;
			}
			event.preventDefault();
			shouldSave = event.key === 'Enter';
			cell.blur();
		};
		cell.contentEditable = 'true';
		cell.addEventListener('keydown', handleEditorKeys);
		cell.addEventListener('blur', () => {
			const editedText = cell.textContent.trim();
			cell.removeAttribute('contenteditable');
			cell.removeEventListener('keydown', handleEditorKeys);
			cell.textContent = shouldSave && editedText ? editedText : originalText;
			refreshAfterEdit();
		}, { once: true });
		cell.focus();
	};

	const editOwner = (cell) => {
		const originalText = cell.textContent.trim();
		const ownerInput = document.createElement('input');
		const ownerNames = [...new Set(getAllRows().map(getOwner).filter(Boolean))];
		let shouldSave = true;

		ownerOptions.replaceChildren(...ownerNames.map((ownerName) => new Option(ownerName, ownerName)));
		ownerInput.type = 'text';
		ownerInput.className = 'cell-editor';
		ownerInput.value = originalText;
		ownerInput.setAttribute('list', ownerOptions.id);
		ownerInput.setAttribute('aria-label', 'Editar responsável');
		ownerInput.addEventListener('keydown', (event) => {
			if (event.key !== 'Enter' && event.key !== 'Escape') {
				return;
			}
			event.preventDefault();
			shouldSave = event.key === 'Enter';
			ownerInput.blur();
		});
		ownerInput.addEventListener('blur', () => {
			const newOwner = ownerInput.value.trim();
			ownerInput.remove();
			cell.classList.remove('is-editing');
			cell.textContent = shouldSave && newOwner ? newOwner : originalText;
			refreshAfterEdit();
		}, { once: true });

		cell.classList.add('is-editing');
		cell.append(ownerInput);
		ownerInput.focus();
		ownerInput.select();
	};

	const editStatus = (cell) => {
		const row = cell.parentElement;
		const typedStatus = window.prompt('Status (Pendente | Em negociação | Ganho | Perdido):', getStatus(row))?.trim();
		if (!typedStatus) {
			return;
		}
		const statusStyle = STATUS_STYLES.find(({ pattern }) => pattern.test(typedStatus));
		const statusLabel = statusStyle?.label || typedStatus;
		const statusClasses = ['status', statusStyle?.className].filter(Boolean).join(' ');
		const statusIcon = statusStyle?.icon ? `<i class="${statusStyle.icon}" aria-hidden="true"></i> ` : '';

		cell.innerHTML = `<span class="${statusClasses}">${statusIcon}${escapeHtml(statusLabel)}</span>`;
		row.dataset.status = statusLabel;

		if (isLostLead(row)) {
			const currentReason = getLossReason(row);
			const typedReason = window.prompt('Motivo de perda (opcional):', currentReason) ?? currentReason;
			lossReasons[row.id] = { reason: typedReason.trim() };
			writeStoredJson(LOSS_REASONS_STORAGE_KEY, lossReasons);
		}

		refreshAfterEdit();
	};

	const CELL_EDITORS = {
		[COLUMN.nextStep]: editNextStep,
		[COLUMN.owner]: editOwner,
		[COLUMN.status]: editStatus
	};

	const isActivationKey = (event) => event.key === 'Enter' || event.key === ' ';

	const handleTableDoubleClick = (event) => {
		const cell = event.target.closest('td');
		const startEditing = cell && CELL_EDITORS[cell.cellIndex];
		if (!isPlusModeActive || !startEditing || cell.isContentEditable || cell.classList.contains('is-editing')) {
			return;
		}
		startEditing(cell);
	};

	const handleHeaderClick = (event) => {
		const headerCell = event.target.closest('th');
		if (headerCell) {
			sortByHeader(headerCell);
		}
	};

	const handleHeaderKeydown = (event) => {
		const headerCell = event.target.closest('th');
		if (!headerCell || !isActivationKey(event)) {
			return;
		}
		event.preventDefault();
		sortByHeader(headerCell);
	};

	const handleObservationClick = (event) => {
		const jumpButton = event.target.closest('.observation-jump');
		const observationHead = event.target.closest('.observation-head');
		if (jumpButton) {
			focusTableRow(jumpButton.dataset.target);
			return;
		}
		if (observationHead) {
			toggleObservationItem(observationHead.closest('.observation-item'));
		}
	};

	const handleObservationKeydown = (event) => {
		const observationTitle = event.target.closest('.observation-title');
		if (!observationTitle || !isActivationKey(event)) {
			return;
		}
		event.preventDefault();
		toggleObservationItem(observationTitle.closest('.observation-item'));
	};

	const handleObservationInput = (event) => {
		if (event.target.matches('.observation-note')) {
			saveObservationNote(event.target);
		}
	};

	const handleModalBackdropClick = (event) => {
		if (event.target === highlightModal) {
			closeHighlightModal();
		}
	};

	const handleGlobalShortcuts = (event) => {
		const pressedKey = event.key?.toLowerCase();
		const activeElement = document.activeElement;
		const isTyping = activeElement?.matches('input, textarea, select') || activeElement?.isContentEditable;
		const shortcutActions = {
			f: clearFilters,
			g: showHighlights,
			e: () => exportTable(event.shiftKey)
		};

		if (!pressedKey) {
			return;
		}
		if (event.ctrlKey && event.shiftKey && pressedKey === 'p') {
			event.preventDefault();
			togglePlusMode();
			return;
		}
		if (pressedKey === 'escape' && isHighlightModalOpen()) {
			closeHighlightModal();
			return;
		}
		if (isTyping || event.metaKey || event.ctrlKey || event.altKey || !shortcutActions[pressedKey]) {
			return;
		}
		event.preventDefault();
		shortcutActions[pressedKey]();
	};

	const mountInterface = () => {
		if (!toolbar.isConnected) {
			card.insertBefore(toolbar, tableWrapper);
		}
		card.insertBefore(insightsPanel, tableWrapper);
		document.body.append(highlightModal, plusHint, ownerOptions);
	};

	const bindEvents = () => {
		searchInput.addEventListener('input', debounce(applyFiltersAndSave, FILTER_DELAY));
		selectFilters.forEach((selectElement) => selectElement.addEventListener('change', applyFiltersAndSave));
		clearButton.addEventListener('click', clearFilters);
		highlightButton.addEventListener('click', showHighlights);
		exportButton.addEventListener('click', (event) => exportTable(event.shiftKey));
		copyButton.addEventListener('click', copyTable);
		insightsToggle.addEventListener('click', toggleInsights);
		highlightCloseButton.addEventListener('click', closeHighlightModal);
		highlightModal.addEventListener('click', handleModalBackdropClick);
		tableHead.addEventListener('click', handleHeaderClick);
		tableHead.addEventListener('keydown', handleHeaderKeydown);
		tableBody.addEventListener('dblclick', handleTableDoubleClick);
		observationsList?.addEventListener('click', handleObservationClick);
		observationsList?.addEventListener('keydown', handleObservationKeydown);
		observationsList?.addEventListener('input', handleObservationInput);
		document.addEventListener('keydown', handleGlobalShortcuts);
		window.addEventListener('resize', debounce(updateTableHeight, RESIZE_DELAY));
		document.fonts?.ready.then(updateTableHeight);
	};

	observationNotes = readStoredJson(NOTES_STORAGE_KEY);
	lossReasons = readStoredJson(LOSS_REASONS_STORAGE_KEY);

	mountInterface();
	prepareSortableHeaders();
	assignRowIds();
	refreshFilterOptions();
	buildObservationList();
	restoreState();
	applyFilters();
	updatePlusHint();
	bindEvents();
})();