let dataset = [];
let headerRow = [];
let activeTool = 'mass';
let selectedOperators = [];
let availableOperators = [];
let availableTrips = [];
let operatorColumnIndex = -1;
let tripColumnIndex = -1;
let toastTimer = null;
let toolSwitchTimer = null;

const COL = {
	SPX: 2,
	TO_STATUS: 4
};

const fileInput = document.getElementById('fileInput');
const uploadDropzone = document.getElementById('uploadDropzone');
const chooseFileBtn = document.getElementById('chooseFileBtn');
const replaceFileBtn = document.getElementById('replaceFileBtn');
const removeFileBtn = document.getElementById('removeFileBtn');
const closeLoadedCardBtn = document.getElementById('closeLoadedCardBtn');
const metaFileName = document.getElementById('metaFileName');
const metaFileSize = document.getElementById('metaFileSize');
const metaRows = document.getElementById('metaRows');
const metaTrip = document.getElementById('metaTrip');
const tripInput = document.getElementById('tripInput');
const tripCombobox = document.getElementById('tripCombobox');
const tripDropdown = document.getElementById('tripDropdown');
const tripHelper = document.getElementById('tripHelper');
const slotSelect = document.getElementById('slotSelect');
const reportDateInput = document.getElementById('reportDate');
const reportSummary = document.getElementById('reportSummary');
const tabMassUpload = document.getElementById('tabMassUpload');
const tabTOReport = document.getElementById('tabTOReport');
const massUploadPanel = document.getElementById('massUploadPanel');
const toReportPanel = document.getElementById('toReportPanel');
const operatorInput = document.getElementById('operatorInput');
const operatorCombobox = document.getElementById('operatorCombobox');
const operatorDropdown = document.getElementById('operatorDropdown');
const selectedOperatorTags = document.getElementById('selectedOperatorTags');
const operatorCountChip = document.getElementById('operatorCountChip');
const operatorDetectedCount = document.getElementById('operatorDetectedCount');
const tripDetectedCount = document.getElementById('tripDetectedCount');
const selectAllOperatorsBtn = document.getElementById('selectAllOperatorsBtn');
const clearOperatorsBtn = document.getElementById('clearOperatorsBtn');
const generateMassBtn = document.getElementById('generateMassBtn');
const copyTrackingBtn = document.getElementById('copyTrackingBtn');
const generateTOBtn = document.getElementById('generateTOBtn');
const generateAllBtn = document.getElementById('generateAllBtn');
const totalDataEl = document.getElementById('totalData');
const filteredDataEl = document.getElementById('filteredData');
const previewBody = document.getElementById('previewBody');
const errorMsg = document.getElementById('errorMsg');
const successMsg = document.getElementById('successMsg');
const warningMsg = document.getElementById('warningMsg');
const progressList = document.getElementById('progressList');
const toast = document.getElementById('toast');
const themeToggle = document.getElementById('themeToggle');
const prefersReducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

function escapeHtml(value) {
	return String(value ?? '')
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
		.replaceAll('"', '&quot;')
		.replaceAll("'", '&#39;');
}

function hideToast() {
	if (toastTimer) {
		clearTimeout(toastTimer);
	}
	toast.className = 'toast';
	toast.innerHTML = '';
}

function showToast(messageOrConfig, type = 'success') {
	if (toastTimer) {
		clearTimeout(toastTimer);
	}

	const config = typeof messageOrConfig === 'object' && messageOrConfig !== null
		? messageOrConfig
		: { message: String(messageOrConfig ?? ''), type };

	const toastType = config.type || 'success';
	const title = config.title || (toastType === 'error' ? 'Error' : 'Success');
	const message = config.message || '';
	const details = Array.isArray(config.details) ? config.details : [];
	const duration = typeof config.duration === 'number' ? config.duration : 4000;
	const isProcessing = toastType === 'processing';
	const icon = isProcessing
		? '<span class="toast-spinner" aria-hidden="true"></span>'
		: `<span class="toast-icon" aria-hidden="true">${escapeHtml(config.icon || (toastType === 'error' ? '✖' : '✔'))}</span>`;

	const detailsHtml = details.length
		? `<ul class="toast-steps">${details.map((item) => `<li>${escapeHtml(item)}</li>`).join('')}</ul>`
		: '';

	toast.innerHTML = `
		<div class="toast-card">
			<div class="toast-head">
				<div class="toast-head-left">${icon}<strong>${escapeHtml(title)}</strong></div>
				<button type="button" class="toast-close" aria-label="Close notification">✕</button>
			</div>
			${message ? `<div class="toast-message">${escapeHtml(message)}</div>` : ''}
			${detailsHtml}
		</div>
	`;

	toast.className = `toast show toast-${toastType}`;

	if (duration > 0) {
		toastTimer = setTimeout(() => {
			hideToast();
		}, duration);
	}
}

function showProcessingToast(stepMessages) {
	showToast({
		type: 'processing',
		title: 'Processing report...',
		message: 'Please wait while we generate the report.',
		details: stepMessages,
		duration: 0
	});
}

function showReportSuccessToast(messages) {
	showToast({
		type: 'success',
		title: 'Report Generated Successfully',
		message: messages.join(' • '),
		duration: 4000,
		icon: '✔'
	});
}

function showReportErrorToast() {
	showToast({
		type: 'error',
		title: 'Report Generation Failed',
		message: 'Something went wrong while generating the report.',
		duration: 4000,
		icon: '✖'
	});
}

function showWarningToast(message) {
	showToast({
		type: 'warning',
		title: 'Warning',
		message,
		duration: 4000,
		icon: '⚠'
	});
}

function resetMessages() {
	errorMsg.style.display = 'none';
	successMsg.style.display = 'none';
	warningMsg.style.display = 'none';
}

function setError(message) {
	resetMessages();
	errorMsg.textContent = message;
	errorMsg.style.display = 'block';
	showToast(message, 'error');
}

function setSuccess(message) {
	resetMessages();
	successMsg.textContent = message;
	successMsg.style.display = 'block';
	showToast(message, 'success');
}

function setWarning(message) {
	warningMsg.textContent = message;
	warningMsg.style.display = 'block';
}

function clearWarning() {
	warningMsg.style.display = 'none';
}

function normalize(value) {
	return String(value ?? '').trim();
}

function formatFileSize(bytes) {
	if (!bytes) {
		return '0 B';
	}
	const units = ['B', 'KB', 'MB', 'GB'];
	let size = bytes;
	let unit = 0;
	while (size >= 1024 && unit < units.length - 1) {
		size /= 1024;
		unit += 1;
	}
	return `${unit === 0 ? Math.round(size) : size.toFixed(2)} ${units[unit]}`;
}

function setUploaderState(state) {
	uploadDropzone.setAttribute('data-state', state);
}

function resetLoadedMeta() {
	metaFileName.textContent = '-';
	metaFileSize.textContent = '-';
	metaRows.textContent = '-';
	metaTrip.textContent = '-';
}

function updateLoadedMeta(file, rowCount, tripDetected = '') {
	metaFileName.textContent = file?.name || '-';
	metaFileSize.textContent = file ? formatFileSize(file.size) : '-';
	metaRows.textContent = file ? `${rowCount} rows` : '-';
	metaTrip.textContent = tripDetected || '-';
}

function setProgress(stepKey) {
	Array.from(progressList.querySelectorAll('li')).forEach((item) => {
		item.classList.remove('active', 'done');
	});

	const order = ['filter', 'mass', 'to'];
	const stepIndex = order.indexOf(stepKey);
	Array.from(progressList.querySelectorAll('li')).forEach((item) => {
		const itemIndex = order.indexOf(item.dataset.step);
		if (itemIndex < stepIndex) {
			item.classList.add('done');
		} else if (itemIndex === stepIndex) {
			item.classList.add('active');
		}
	});
}

function clearProgress() {
	Array.from(progressList.querySelectorAll('li')).forEach((item) => {
		item.classList.remove('active', 'done');
	});
}

function animateStatValue(element, targetValue) {
	const safeTarget = Number.isFinite(targetValue) ? Math.max(0, targetValue) : 0;
	const currentValue = Number.parseInt(element.dataset.currentValue || element.textContent || '0', 10) || 0;

	if (prefersReducedMotion || currentValue === safeTarget) {
		element.textContent = String(safeTarget);
		element.dataset.currentValue = String(safeTarget);
		return;
	}

	const startTime = performance.now();
	const duration = 360;
	const delta = safeTarget - currentValue;

	const tick = (now) => {
		const elapsed = now - startTime;
		const progress = Math.min(1, elapsed / duration);
		const eased = 1 - (1 - progress) * (1 - progress);
		const nextValue = Math.round(currentValue + delta * eased);
		element.textContent = String(nextValue);
		if (progress < 1) {
			requestAnimationFrame(tick);
		} else {
			element.textContent = String(safeTarget);
			element.dataset.currentValue = String(safeTarget);
			element.classList.remove('value-updated');
			void element.offsetWidth;
			element.classList.add('value-updated');
		}
	};

	requestAnimationFrame(tick);
}

function setButtonLoading(button, isLoading, loadingText = 'Processing...') {
	if (!button) {
		return;
	}
	if (isLoading) {
		if (!button.dataset.originalText) {
			button.dataset.originalText = button.textContent;
		}
		button.textContent = loadingText;
		button.disabled = true;
		button.classList.add('is-loading');
		return;
	}

	button.disabled = false;
	button.classList.remove('is-loading');
	if (button.dataset.originalText) {
		button.textContent = button.dataset.originalText;
	}
}

function applyTheme(theme) {
	const isDark = theme === 'dark';
	document.body.classList.toggle('dark-mode', isDark);
	themeToggle.textContent = isDark ? '☀️' : '🌙';
}

function initTheme() {
	const saved = localStorage.getItem('mass-upload-theme') || 'light';
	applyTheme(saved);
}

function setActiveTool(tool) {
	if (tool === activeTool) {
		refreshPreview();
		return;
	}

	if (toolSwitchTimer) {
		clearTimeout(toolSwitchTimer);
		toolSwitchTimer = null;
	}

	const nextPanel = tool === 'mass' ? massUploadPanel : toReportPanel;
	const currentPanel = activeTool === 'mass' ? massUploadPanel : toReportPanel;

	activeTool = tool;
	tabMassUpload.classList.toggle('active', tool === 'mass');
	tabTOReport.classList.toggle('active', tool === 'to');

	if (prefersReducedMotion) {
		massUploadPanel.classList.toggle('hidden', tool !== 'mass');
		toReportPanel.classList.toggle('hidden', tool !== 'to');
		refreshPreview();
		return;
	}

	if (currentPanel !== nextPanel && !currentPanel.classList.contains('hidden')) {
		currentPanel.classList.add('is-hiding');
		toolSwitchTimer = setTimeout(() => {
			currentPanel.classList.add('hidden');
			currentPanel.classList.remove('is-hiding');
			nextPanel.classList.remove('hidden');
			nextPanel.classList.add('is-showing');
			requestAnimationFrame(() => {
				nextPanel.classList.remove('is-showing');
			});
			toolSwitchTimer = null;
		}, 190);
	} else {
		massUploadPanel.classList.toggle('hidden', tool !== 'mass');
		toReportPanel.classList.toggle('hidden', tool !== 'to');
		nextPanel.classList.add('is-showing');
		requestAnimationFrame(() => {
			nextPanel.classList.remove('is-showing');
		});
	}

	refreshPreview();
}

function populateSlots() {
	for (let i = 1; i <= 10; i += 1) {
		const option = document.createElement('option');
		option.value = `Slot ${i}`;
		option.textContent = `Slot ${i}`;
		slotSelect.appendChild(option);
	}
}

function getTodayInputDate() {
	const now = new Date();
	const year = String(now.getFullYear());
	const month = String(now.getMonth() + 1).padStart(2, '0');
	const day = String(now.getDate()).padStart(2, '0');
	return `${year}-${month}-${day}`;
}

function formatDateDDMMYYYY(selectedDate) {
	if (!selectedDate) {
		return '';
	}
	const normalizedDate = normalize(selectedDate);

	if (/^\d{2}-\d{2}-\d{4}$/.test(normalizedDate)) {
		return normalizedDate;
	}

	if (!/^\d{4}-\d{2}-\d{2}$/.test(normalizedDate)) {
		return '';
	}

	const [year, month, day] = normalizedDate.split('-');
	return `${day}-${month}-${year}`;
}

function getSelectedReportDateFormatted() {
	return formatDateDDMMYYYY(reportDateInput.value);
}

function updateReportSummary() {
	const trip = normalize(tripInput.value) || '-';
	const slot = slotSelect.value || '-';
	const reportDate = getSelectedReportDateFormatted() || '-';
	reportSummary.textContent = `Trip: ${trip} | Slot: ${slot} | Report Date: ${reportDate}`;
}

function updateOperatorCounter() {
	operatorCountChip.textContent = `Operator dipilih: ${selectedOperators.length}`;
}

function updateOperatorDetectedCounter() {
	operatorDetectedCount.textContent = `Operators detected: ${availableOperators.length}`;
}

function updateTripDetectedCounter() {
	tripDetectedCount.textContent = `Trips detected: ${availableTrips.length}`;
}

function getColumnIndexByName(columnName) {
	const normalizedColumnName = normalize(columnName).toLowerCase();
	return headerRow.findIndex((header) => normalize(header).toLowerCase() === normalizedColumnName);
}

function getColumnIndexByNames(columnNames) {
	for (const columnName of columnNames) {
		const index = getColumnIndexByName(columnName);
		if (index !== -1) {
			return index;
		}
	}
	return -1;
}

function extractUniqueOperatorsFromDataset() {
	const index = getColumnIndexByNames(['Operator']);
	if (index === -1) {
		return { hasOperatorColumn: false, operators: [], columnIndex: -1 };
	}

	const operators = Array.from(
		new Set(
			dataset
				.map((row) => normalize(row[index]))
				.filter(Boolean)
		)
	);

	return {
		hasOperatorColumn: true,
		operators,
		columnIndex: index
	};
}

function extractUniqueTripsFromDataset() {
	const index = getColumnIndexByNames(['Line Hual Trip Number', 'LineHaul Trip Number']);
	if (index === -1) {
		return { hasTripColumn: false, trips: [], columnIndex: -1 };
	}

	const trips = Array.from(
		new Set(
			dataset
				.map((row) => normalize(row[index]))
				.filter(Boolean)
		)
	);

	return {
		hasTripColumn: true,
		trips,
		columnIndex: index
	};
}

function setTripDropdownOptions(trips) {
	availableTrips = [...trips];
	renderTripDropdown();
}

function getMissingRequiredColumns() {
	const missing = [];
	if (operatorColumnIndex === -1) {
		missing.push('Operator');
	}
	if (tripColumnIndex === -1) {
		missing.push('Line Hual Trip Number');
	}
	return missing;
}

function updateRequiredColumnsWarning(showToastMessage = false) {
	const missingColumns = getMissingRequiredColumns();
	if (!missingColumns.length) {
		clearWarning();
		return;
	}

	const message = `Missing column in uploaded file: ${missingColumns.join(' / ')}`;
	setWarning(message);
	if (showToastMessage) {
		showWarningToast(message);
	}
}

function resetOperatorState() {
	selectedOperators = [];
	availableOperators = [];
	operatorColumnIndex = -1;
	operatorInput.value = '';
	updateOperatorCounter();
	updateOperatorDetectedCounter();
	renderSelectedOperatorTags();
	renderOperatorDropdown();
}

function resetTripState() {
	availableTrips = [];
	tripColumnIndex = -1;
	tripInput.value = '';
	setTripDropdownOptions([]);
	tripHelper.textContent = '';
	updateTripDetectedCounter();
}

function rebuildOperatorOptionsFromDataset() {
	const extracted = extractUniqueOperatorsFromDataset();
	availableOperators = extracted.operators;
	operatorColumnIndex = extracted.columnIndex;

	updateOperatorDetectedCounter();
	renderOperatorDropdown();
}

function rebuildTripOptionsFromDataset() {
	const extracted = extractUniqueTripsFromDataset();
	availableTrips = extracted.trips;
	tripColumnIndex = extracted.columnIndex;
	setTripDropdownOptions(availableTrips);

	if (availableTrips.length === 1) {
		tripInput.value = availableTrips[0];
		tripHelper.textContent = 'Trip detected automatically.';
	} else {
		tripInput.value = '';
		tripHelper.textContent = '';
	}

	updateTripDetectedCounter();
	renderTripDropdown();
}

function renderTripDropdown() {
	const keyword = normalize(tripInput.value).toLowerCase();
	const shouldOpen = document.activeElement === tripInput || Boolean(keyword);
	if (!shouldOpen) {
		tripDropdown.classList.remove('open');
		tripDropdown.classList.remove('open-up');
		return;
	}

	const filteredTrips = availableTrips.filter((trip) => trip.toLowerCase().includes(keyword));
	tripDropdown.innerHTML = '';

	if (!filteredTrips.length) {
		const empty = document.createElement('div');
		empty.className = 'operator-option-empty';
		empty.textContent = tripColumnIndex === -1
			? 'Missing column in uploaded file: Line Hual Trip Number'
			: 'Trip tidak ditemukan.';
		tripDropdown.appendChild(empty);
		positionTripDropdown();
		tripDropdown.classList.add('open');
		return;
	}

	filteredTrips.forEach((trip) => {
		const option = document.createElement('button');
		option.type = 'button';
		option.className = 'operator-option';
		option.textContent = trip;
		option.addEventListener('click', () => {
			tripInput.value = trip;
			tripDropdown.classList.remove('open');
			refreshPreview();
			updateReportSummary();
			tripInput.focus();
		});
		tripDropdown.appendChild(option);
	});

	positionTripDropdown();
	tripDropdown.classList.add('open');
}

function positionTripDropdown() {
	const viewportMargin = 12;
	const minPanelHeight = 120;
	const maxPanelHeight = 240;
	const comboboxRect = tripCombobox.getBoundingClientRect();
	const optionCount = Math.max(1, tripDropdown.childElementCount);
	const estimatedHeight = Math.min(maxPanelHeight, optionCount * 43 + 12);

	const spaceBelow = window.innerHeight - comboboxRect.bottom - viewportMargin;
	const spaceAbove = comboboxRect.top - viewportMargin;
	const shouldOpenUp = spaceBelow < estimatedHeight && spaceAbove > spaceBelow;

	tripDropdown.classList.toggle('open-up', shouldOpenUp);

	const availableSpace = shouldOpenUp ? spaceAbove : spaceBelow;
	const computedMaxHeight = Math.max(minPanelHeight, Math.min(maxPanelHeight, availableSpace));
	tripDropdown.style.maxHeight = `${computedMaxHeight}px`;
}

function closeTripDropdown() {
	tripDropdown.classList.remove('open');
	tripDropdown.classList.remove('open-up');
}

function renderSelectedOperatorTags() {
	selectedOperatorTags.innerHTML = '';
	if (!selectedOperators.length) {
		const empty = document.createElement('span');
		empty.className = 'selected-tag-empty';
		empty.textContent = 'Belum ada operator dipilih.';
		selectedOperatorTags.appendChild(empty);
		return;
	}

	selectedOperators.forEach((operator) => {
		const tag = document.createElement('span');
		tag.className = 'selected-tag';
		tag.textContent = operator;

		const removeBtn = document.createElement('button');
		removeBtn.type = 'button';
		removeBtn.className = 'tag-remove';
		removeBtn.textContent = '✕';
		removeBtn.addEventListener('click', () => {
			selectedOperators = selectedOperators.filter((item) => item !== operator);
			renderSelectedOperatorTags();
			renderOperatorDropdown();
			updateOperatorCounter();
			refreshPreview();
		});

		tag.appendChild(removeBtn);
		selectedOperatorTags.appendChild(tag);
	});
}

function renderOperatorDropdown() {
	const keyword = normalize(operatorInput.value).toLowerCase();
	const shouldOpen = document.activeElement === operatorInput || Boolean(keyword);
	if (!shouldOpen) {
		operatorDropdown.classList.remove('open');
		return;
	}

	const filtered = availableOperators.filter((operator) => {
		return operator.toLowerCase().includes(keyword) && !selectedOperators.includes(operator);
	});

	operatorDropdown.innerHTML = '';
	if (!filtered.length) {
		const empty = document.createElement('div');
		empty.className = 'operator-option-empty';
		empty.textContent = operatorColumnIndex === -1
			? 'Missing column in uploaded file: Operator'
			: availableOperators.length
				? 'Operator tidak ditemukan.'
				: 'Belum ada operator terdeteksi dari dataset.';
		operatorDropdown.appendChild(empty);
		operatorDropdown.classList.add('open');
		return;
	}

	filtered.forEach((operator) => {
		const option = document.createElement('button');
		option.type = 'button';
		option.className = 'operator-option';
		option.textContent = operator;
		option.addEventListener('click', () => {
			selectedOperators.push(operator);
			operatorInput.value = '';
			renderSelectedOperatorTags();
			renderOperatorDropdown();
			updateOperatorCounter();
			refreshPreview();
			operatorInput.focus();
		});
		operatorDropdown.appendChild(option);
	});

	operatorDropdown.classList.add('open');
}

function closeOperatorDropdown() {
	operatorDropdown.classList.remove('open');
}

function clearSourceFile() {
	dataset = [];
	headerRow = [];
	resetOperatorState();
	resetTripState();
	fileInput.value = '';
	resetLoadedMeta();
	setUploaderState('empty');
	refreshPreview();
	clearProgress();
	resetMessages();
	clearWarning();
}

async function loadFile(file) {
	if (!file) {
		return;
	}

	resetMessages();
	clearWarning();
	clearProgress();
	resetOperatorState();
	resetTripState();
	setUploaderState('loading');

	try {
		const extension = file.name.toLowerCase().split('.').pop();
		const aoa = extension === 'csv' ? await parseCSV(file) : await parseExcel(file);
		if (!aoa.length || aoa.length < 2) {
			throw new Error('Data kosong atau header tidak ditemukan.');
		}

		headerRow = aoa[0].slice(0, 35);
		dataset = aoa.slice(1).map((row) => row.slice(0, 35));
		rebuildOperatorOptionsFromDataset();
		rebuildTripOptionsFromDataset();
		updateRequiredColumnsWarning(true);
		const detectedTrip = normalize(tripInput.value);
		updateLoadedMeta(file, dataset.length, detectedTrip);
		setUploaderState('loaded');
		refreshPreview();
		setSuccess('Source file loaded successfully.');
		showToast('Source file loaded successfully.', 'success');
	} catch (error) {
		dataset = [];
		headerRow = [];
		resetOperatorState();
		resetTripState();
		resetLoadedMeta();
		setUploaderState('empty');
		setError(`Gagal membaca file: ${error.message}`);
	}
}

async function parseExcel(file) {
	const buffer = await file.arrayBuffer();
	const workbook = XLSX.read(buffer, { type: 'array' });
	const sheetName = workbook.SheetNames[0];
	const sheet = workbook.Sheets[sheetName];
	return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
}

async function parseCSV(file) {
	const text = await file.text();
	const workbook = XLSX.read(text, { type: 'string', FS: ',' });
	const sheetName = workbook.SheetNames[0];
	const sheet = workbook.Sheets[sheetName];
	return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
}

function filterDatasetByTrip(tripNumber) {
	if (tripColumnIndex === -1) {
		return [];
	}
	const normalizedTrip = normalize(tripNumber);
	return dataset.filter((row) => normalize(row[tripColumnIndex]) === normalizedTrip);
}

function filterDatasetByOperator(rows, operators) {
	if (!operators.length || operatorColumnIndex === -1) {
		return [];
	}
	return rows.filter((row) => operators.includes(normalize(row[operatorColumnIndex])));
}

function createBorderStyle(color = 'D1D5DB') {
	return {
		top: { style: 'thin', color: { rgb: color } },
		bottom: { style: 'thin', color: { rgb: color } },
		left: { style: 'thin', color: { rgb: color } },
		right: { style: 'thin', color: { rgb: color } }
	};
}

function ensureCell(worksheet, rowIndex, colIndex) {
	const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
	if (!worksheet[cellAddress]) {
		worksheet[cellAddress] = { t: 's', v: '' };
	}
	if (!worksheet[cellAddress].s) {
		worksheet[cellAddress].s = {};
	}
	return worksheet[cellAddress];
}

function calculateColumnWidths(rows, minWidth = 18, maxWidth = 28) {
	if (!rows.length) {
		return [];
	}
	const colCount = Math.max(...rows.map((row) => row.length));
	const widths = [];
	for (let col = 0; col < colCount; col += 1) {
		let maxLen = 0;
		for (let row = 0; row < rows.length; row += 1) {
			const value = rows[row]?.[col] ?? '';
			const len = String(value).trim().length;
			if (len > maxLen) {
				maxLen = len;
			}
		}
		const width = Math.max(minWidth, Math.min(maxWidth, maxLen + 2));
		widths.push({ wch: width });
	}
	return widths;
}

function applyStandardHeaderStyle(worksheet, colCount) {
	const border = createBorderStyle('CBD5E1');
	for (let col = 0; col < colCount; col += 1) {
		const cell = ensureCell(worksheet, 0, col);
		cell.s = {
			...(cell.s || {}),
			font: { bold: true },
			alignment: { horizontal: 'center', vertical: 'center' },
			fill: { patternType: 'solid', fgColor: { rgb: 'E2E8F0' } },
			border
		};
	}
	worksheet['!rows'] = [{ hpx: 28 }];
}

function applyBordersAndZebra(worksheet, rowCount, colCount, options = {}) {
	const {
		evenRowColor = 'F8FAFC',
		oddRowColor = 'FFFFFF',
		alignments = {}
	} = options;
	const border = createBorderStyle('D1D5DB');

	for (let row = 1; row < rowCount; row += 1) {
		const isEvenDisplayRow = (row + 1) % 2 === 0;
		const fillColor = isEvenDisplayRow ? evenRowColor : oddRowColor;

		for (let col = 0; col < colCount; col += 1) {
			const cell = ensureCell(worksheet, row, col);
			const alignment = alignments[col] || { horizontal: 'left', vertical: 'center' };
			cell.s = {
				...(cell.s || {}),
				alignment,
				fill: { patternType: 'solid', fgColor: { rgb: fillColor } },
				border
			};
		}
	}
}

function applyFreezeAndFilter(worksheet, colCount) {
	if (colCount <= 0) {
		return;
	}
	const endCol = XLSX.utils.encode_col(colCount - 1);
	worksheet['!autofilter'] = { ref: `A1:${endCol}1` };
	worksheet['!freeze'] = {
		xSplit: 0,
		ySplit: 1,
		topLeftCell: 'A2',
		activePane: 'bottomLeft',
		state: 'frozen'
	};
}

function formatMassUploadWorksheet(worksheet, sheetData) {
	const rowCount = sheetData.length;
	const colCount = sheetData[0]?.length || 0;
	worksheet['!cols'] = calculateColumnWidths(sheetData, 18, 28);
	applyStandardHeaderStyle(worksheet, colCount);
	applyBordersAndZebra(worksheet, rowCount, colCount, {
		alignments: {
			0: { horizontal: 'center', vertical: 'center' },
			1: { horizontal: 'center', vertical: 'center' },
			2: { horizontal: 'center', vertical: 'center' },
			3: { horizontal: 'center', vertical: 'center' },
			4: { horizontal: 'center', vertical: 'center' }
		}
	});
	applyFreezeAndFilter(worksheet, colCount);
}

function formatDataWorksheet(worksheet, rows) {
	const rowCount = rows.length;
	const colCount = rows[0]?.length || 0;
	worksheet['!cols'] = calculateColumnWidths(rows, 18, 28);
	applyStandardHeaderStyle(worksheet, colCount);
	applyBordersAndZebra(worksheet, rowCount, colCount);
	applyFreezeAndFilter(worksheet, colCount);
}

function formatPivotWorksheet(worksheet, pivotRows) {
	const rowCount = pivotRows.length;
	const colCount = pivotRows[0]?.length || 0;
	const border = createBorderStyle('D1D5DB');
	worksheet['!cols'] = calculateColumnWidths(pivotRows, 12, 28);
	applyStandardHeaderStyle(worksheet, colCount);

	for (let col = 0; col < colCount; col += 1) {
		const headerCell = ensureCell(worksheet, 0, col);
		headerCell.s = {
			...(headerCell.s || {}),
			font: { bold: true, color: { rgb: '1D4ED8' } },
			alignment: { horizontal: 'center', vertical: 'center' },
			fill: { patternType: 'solid', fgColor: { rgb: 'DBEAFE' } },
			border
		};
	}

	for (let row = 1; row < rowCount; row += 1) {
		const isGrandTotal = row === rowCount - 1;
		for (let col = 0; col < colCount; col += 1) {
			const cell = ensureCell(worksheet, row, col);
			const baseStyle = {
				alignment: { horizontal: col === 0 ? 'left' : 'center', vertical: 'center' },
				border,
				fill: {
					patternType: 'solid',
					fgColor: { rgb: row % 2 === 0 ? 'F8FAFC' : 'FFFFFF' }
				}
			};

			if (isGrandTotal) {
				baseStyle.font = { bold: true, color: { rgb: '1D4ED8' } };
				baseStyle.fill = { patternType: 'solid', fgColor: { rgb: 'DBEAFE' } };
				baseStyle.border = {
					...border,
					top: { style: 'thin', color: { rgb: '94A3B8' } }
				};
			}

			cell.s = {
				...(cell.s || {}),
				...baseStyle
			};
		}
	}

	applyFreezeAndFilter(worksheet, colCount);
}

function refreshPreview() {
	previewBody.innerHTML = '';
	const trip = normalize(tripInput.value);
	const byTrip = trip ? filterDatasetByTrip(trip) : [];
	const filteredRows = activeTool === 'mass' ? filterDatasetByOperator(byTrip, selectedOperators) : byTrip;

	animateStatValue(totalDataEl, dataset.length);
	animateStatValue(filteredDataEl, filteredRows.length);

	if (!filteredRows.length) {
		const row = document.createElement('tr');
		row.innerHTML = '<td colspan="2">Belum ada data.</td>';
		previewBody.appendChild(row);
		return;
	}

	filteredRows.forEach((rowData, index) => {
		const row = document.createElement('tr');
		row.innerHTML = `<td>${index + 1}</td><td>${normalize(rowData[COL.SPX])}</td>`;
		previewBody.appendChild(row);
	});
}

function createPivotData(rows) {
	const pivotMap = {};
	rows.forEach((row) => {
		const status = normalize(row[COL.TO_STATUS]) || 'Unknown';
		const tracking = normalize(row[COL.SPX]);
		if (!tracking) {
			return;
		}
		pivotMap[status] = (pivotMap[status] || 0) + 1;
	});

	const result = [['TO Status', 'Count SPX Tracking Number']];
	let grandTotal = 0;
	Object.entries(pivotMap).forEach(([status, count]) => {
		result.push([status, count]);
		grandTotal += count;
	});
	result.push(['Grand Total', grandTotal]);
	return result;
}

function generateMassUploadFile(rows) {
	const data = rows.map((row) => [normalize(row[COL.SPX]), '', '', '', '']);
	const sheetData = [['SPX Tracking Number', 'Weight(KG)', 'Length(cm)', 'Width(cm)', 'Height(cm)'], ...data];
	const workbook = XLSX.utils.book_new();
	const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
	formatMassUploadWorksheet(worksheet, sheetData);
	XLSX.utils.book_append_sheet(workbook, worksheet, '工作表1');

	const datePart = getSelectedReportDateFormatted();
	const fileName = `Mass Upload Bulky - ${slotSelect.value} - ${datePart}.xlsx`;
	XLSX.writeFile(workbook, fileName);
}

function generateTOManagementReport(rows) {
	const workbook = XLSX.utils.book_new();
	const dataSheetRows = [headerRow.slice(0, 35), ...rows.map((row) => row.slice(0, 35))];
	const dataSheet = XLSX.utils.aoa_to_sheet(dataSheetRows);
	formatDataWorksheet(dataSheet, dataSheetRows);
	XLSX.utils.book_append_sheet(workbook, dataSheet, 'Data');

	const datePart = getSelectedReportDateFormatted();
	const pivotRows = createPivotData(rows);
	const pivotSheet = XLSX.utils.aoa_to_sheet(pivotRows);
	formatPivotWorksheet(pivotSheet, pivotRows);
	XLSX.utils.book_append_sheet(workbook, pivotSheet, `Pivot - ${slotSelect.value} - ${datePart}`);

	const fileName = `TO Management Report - ${slotSelect.value} - ${datePart}.xlsx`;
	XLSX.writeFile(workbook, fileName);
}

function validateBeforeGenerate(requireOperator = false) {
	if (!dataset.length) {
		setError('Please upload source file first.');
		return null;
	}
	const trip = normalize(tripInput.value);
	if (!trip) {
		setError('LineHaul Trip Number wajib diisi.');
		return null;
	}

	const reportDate = getSelectedReportDateFormatted();
	if (!reportDate) {
		setWarning('Please select a report date first.');
		showWarningToast('Please select a report date first.');
		return null;
	}

	if (tripColumnIndex === -1) {
		const message = 'Missing column in uploaded file: Line Hual Trip Number';
		setWarning(message);
		showWarningToast(message);
		return null;
	}

	const tripRows = filterDatasetByTrip(trip);
	if (!tripRows.length) {
		setError('Trip tidak ditemukan di data.');
		setWarning('Trip tidak ditemukan di data.');
		return null;
	}

	if (requireOperator && operatorColumnIndex === -1) {
		const message = 'Missing column in uploaded file: Operator';
		setWarning(message);
		showWarningToast(message);
		return null;
	}

	if (requireOperator && !selectedOperators.length) {
		setError('Operator belum dipilih.');
		return null;
	}

	return tripRows;
}

function copyTrackingNumbers(rows) {
	const text = rows.map((row) => normalize(row[COL.SPX])).filter(Boolean).join('\n');
	if (!text) {
		setError('Tidak ada tracking number untuk dicopy.');
		return;
	}
	navigator.clipboard.writeText(text)
		.then(() => setSuccess(`Berhasil copy ${rows.length} tracking number.`))
		.catch((error) => setError(`Gagal copy ke clipboard: ${error.message}`));
}

async function generateAllReports() {
	resetMessages();
	clearWarning();
	setButtonLoading(generateAllBtn, true, 'Generating Reports...');
	showProcessingToast([
		'Filtering dataset',
		'Generating Mass Upload file',
		'Generating TO Management Report'
	]);
	try {
		const tripRows = validateBeforeGenerate(true);
		if (!tripRows) {
			clearProgress();
			return;
		}

		const massRows = filterDatasetByOperator(tripRows, selectedOperators);
		if (!massRows.length) {
			setError('Trip tidak ditemukan di data atau operator tidak sesuai.');
			clearProgress();
			return;
		}

		setProgress('filter');
		await new Promise((resolve) => setTimeout(resolve, 120));

		setProgress('mass');
		generateMassUploadFile(massRows);
		await new Promise((resolve) => setTimeout(resolve, 120));

		setProgress('to');
		generateTOManagementReport(tripRows);
		setSuccess('Dua report berhasil digenerate dan diunduh otomatis.');
		showReportSuccessToast([
			'Mass Upload file downloaded',
			'TO Management Report downloaded'
		]);
	} catch (error) {
		setError(`Gagal generate report: ${error.message}`);
		showReportErrorToast();
	} finally {
		setButtonLoading(generateAllBtn, false);
	}
}

function bindKeyboardShortcuts() {
	document.addEventListener('keydown', (event) => {
		const isTyping = ['INPUT', 'TEXTAREA', 'SELECT'].includes(document.activeElement?.tagName);
		if (isTyping) {
			return;
		}

		if (event.ctrlKey && !event.shiftKey && event.key === 'Enter') {
			event.preventDefault();
			generateAllBtn.click();
			return;
		}

		if (event.ctrlKey && event.shiftKey && event.key.toLowerCase() === 'm') {
			event.preventDefault();
			setActiveTool('mass');
			generateMassBtn.click();
			return;
		}

		if (event.ctrlKey && event.shiftKey && event.key.toLowerCase() === 't') {
			event.preventDefault();
			setActiveTool('to');
			generateTOBtn.click();
		}
	});
}

function bindUploadEvents() {
	uploadDropzone.addEventListener('dragover', (event) => {
		event.preventDefault();
		uploadDropzone.classList.add('dragover');
	});

	uploadDropzone.addEventListener('dragleave', () => {
		uploadDropzone.classList.remove('dragover');
	});

	uploadDropzone.addEventListener('drop', (event) => {
		event.preventDefault();
		uploadDropzone.classList.remove('dragover');
		const file = event.dataTransfer.files[0];
		if (file) {
			loadFile(file);
		}
	});

	fileInput.addEventListener('change', (event) => {
		const file = event.target.files[0];
		if (file) {
			loadFile(file);
		}
	});

	chooseFileBtn.addEventListener('click', () => fileInput.click());
	replaceFileBtn.addEventListener('click', () => fileInput.click());
	removeFileBtn.addEventListener('click', () => {
		clearSourceFile();
		setSuccess('Source file removed.');
	});
	closeLoadedCardBtn.addEventListener('click', () => {
		clearSourceFile();
		setSuccess('Source file removed.');
	});
}

function init() {
	populateSlots();
	reportDateInput.value = getTodayInputDate();
	updateOperatorCounter();
	updateOperatorDetectedCounter();
	updateTripDetectedCounter();
	setTripDropdownOptions([]);
	renderSelectedOperatorTags();
	refreshPreview();
	updateReportSummary();
	bindUploadEvents();
	bindKeyboardShortcuts();
	setUploaderState('empty');
	resetLoadedMeta();
	initTheme();

	tabMassUpload.addEventListener('click', () => setActiveTool('mass'));
	tabTOReport.addEventListener('click', () => setActiveTool('to'));

	operatorInput.addEventListener('input', renderOperatorDropdown);
	operatorInput.addEventListener('focus', renderOperatorDropdown);
	operatorInput.addEventListener('keydown', (event) => {
		if (event.key === 'Enter') {
			event.preventDefault();
			const first = operatorDropdown.querySelector('.operator-option');
			if (first) {
				first.click();
			}
		}
	});

	document.addEventListener('click', (event) => {
		if (!operatorCombobox.contains(event.target)) {
			closeOperatorDropdown();
		}
		if (!tripCombobox.contains(event.target)) {
			closeTripDropdown();
		}
	});

	selectAllOperatorsBtn.addEventListener('click', () => {
		selectedOperators = [...availableOperators];
		renderSelectedOperatorTags();
		renderOperatorDropdown();
		updateOperatorCounter();
		refreshPreview();
	});

	clearOperatorsBtn.addEventListener('click', () => {
		selectedOperators = [];
		renderSelectedOperatorTags();
		renderOperatorDropdown();
		updateOperatorCounter();
		refreshPreview();
	});

	tripInput.addEventListener('change', () => {
		refreshPreview();
		updateReportSummary();
	});
	tripInput.addEventListener('input', () => {
		renderTripDropdown();
		refreshPreview();
		updateReportSummary();
	});
	tripInput.addEventListener('focus', renderTripDropdown);
	tripInput.addEventListener('keydown', (event) => {
		if (event.key === 'Enter') {
			event.preventDefault();
			const first = tripDropdown.querySelector('.operator-option');
			if (first) {
				first.click();
			}
		}
	});

	window.addEventListener('resize', () => {
		if (tripDropdown.classList.contains('open')) {
			positionTripDropdown();
		}
	});

	document.addEventListener('scroll', () => {
		if (tripDropdown.classList.contains('open')) {
			positionTripDropdown();
		}
	}, true);

	slotSelect.addEventListener('change', updateReportSummary);
	reportDateInput.addEventListener('change', updateReportSummary);

	generateMassBtn.addEventListener('click', () => {
		resetMessages();
		clearWarning();
		setButtonLoading(generateMassBtn, true, 'Generating Excel...');
		showProcessingToast([
			'Filtering dataset',
			'Generating Mass Upload file'
		]);
		try {
			const tripRows = validateBeforeGenerate(true);
			if (!tripRows) {
				return;
			}
			const massRows = filterDatasetByOperator(tripRows, selectedOperators);
			if (!massRows.length) {
				setError('Trip tidak ditemukan di data atau operator tidak sesuai.');
				return;
			}
			generateMassUploadFile(massRows);
			setSuccess('Mass Upload file berhasil digenerate.');
			showReportSuccessToast(['Mass Upload file downloaded']);
		} catch (error) {
			setError(`Gagal generate report: ${error.message}`);
			showReportErrorToast();
		} finally {
			setButtonLoading(generateMassBtn, false);
		}
	});

	copyTrackingBtn.addEventListener('click', () => {
		const tripRows = validateBeforeGenerate(true);
		if (!tripRows) {
			return;
		}
		const massRows = filterDatasetByOperator(tripRows, selectedOperators);
		if (!massRows.length) {
			setError('Trip tidak ditemukan di data atau operator tidak sesuai.');
			return;
		}
		copyTrackingNumbers(massRows);
	});

	generateTOBtn.addEventListener('click', () => {
		resetMessages();
		clearWarning();
		setButtonLoading(generateTOBtn, true, 'Generating Report...');
		showProcessingToast([
			'Filtering dataset',
			'Generating TO Management Report'
		]);
		try {
			const tripRows = validateBeforeGenerate(false);
			if (!tripRows) {
				return;
			}
			generateTOManagementReport(tripRows);
			setSuccess('TO Management Report berhasil digenerate.');
			showReportSuccessToast(['TO Management Report downloaded']);
		} catch (error) {
			setError(`Gagal generate report: ${error.message}`);
			showReportErrorToast();
		} finally {
			setButtonLoading(generateTOBtn, false);
		}
	});

	generateAllBtn.addEventListener('click', generateAllReports);

	toast.addEventListener('click', (event) => {
		if (event.target.classList.contains('toast-close')) {
			hideToast();
		}
	});

	themeToggle.addEventListener('click', () => {
		const isDark = document.body.classList.contains('dark-mode');
		const nextTheme = isDark ? 'light' : 'dark';
		applyTheme(nextTheme);
		localStorage.setItem('mass-upload-theme', nextTheme);
	});
}

init();
