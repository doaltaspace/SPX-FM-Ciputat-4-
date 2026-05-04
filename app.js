let dataset = [];
let headerRow = [];
let activeTool = 'mass';
let selectedOperators = [];
let availableOperators = [];
let availableTrips = [];
let operatorColumnIndex = -1;
let tripColumnIndex = -1;
let quantityColumnIndex = -1;
let massUploadData = [];
let massUploadDraftData = [];
let bulkyTotalQty = 0;
let toastTimer = null;
let toastHideTimer = null;
let toolSwitchTimer = null;
let notificationAudioContext = null;
let lastNotificationSoundAt = 0;
let notificationLogEntries = [];
let prealertUploadedPdfFile = null;
let customErrorAudioElement = null;
let dbToFiles = [];
let dbSjFile = null;
let linehaulTemplateData = {
	staTime: '-',
	ataTime: '-',
	stdTime: '-',
	stfTime: '-',
	atdTime: '-',
	sealCode: '-'
};
const modalCloseTimers = new WeakMap();
// Optional custom error sound file (example: 'sounds/oink.mp3').
const ERROR_NOTIFICATION_SOUND_URL = 'sounds/error.mp3';

// API endpoint for automatic draft creation (host your backend service URL here).
const PREALERT_DRAFT_API_URL = '';
// Force direct Gmail API flow and skip Apps Script/backend endpoint.
const PREALERT_FORCE_DIRECT_GMAIL_API = true;
const PREALERT_OAUTH_CONFIG_KEY = 'prealert_oauth_config_v1';
const runtimeOAuthConfig = (() => {
	if (typeof window === 'undefined' || !window.localStorage) {
		return {};
	}

	try {
		const raw = window.localStorage.getItem(PREALERT_OAUTH_CONFIG_KEY);
		return raw ? JSON.parse(raw) : {};
	} catch (error) {
		return {};
	}
})();

// Optional Google OAuth client info for direct Gmail API fallback.
const PREALERT_GOOGLE_CLIENT_ID = String(runtimeOAuthConfig.clientId || '').trim();
const PREALERT_GOOGLE_CLIENT_SECRET = String(runtimeOAuthConfig.clientSecret || '').trim();
// Optional refresh token for automatic access-token renewal (direct Gmail API fallback).
const PREALERT_GOOGLE_REFRESH_TOKEN = String(runtimeOAuthConfig.refreshToken || '').trim();
// Optional bearer token bootstrap for first request.
const PREALERT_DRAFT_API_TOKEN = String(runtimeOAuthConfig.accessToken || '').trim();
// Optional: force draft recipient (recommended for team shared mailbox).
const PREALERT_DRAFT_RECIPIENT = '';
const PREALERT_TOKEN_CACHE_KEY = 'prealert_gmail_oauth_cache_v1';
// Default recipients for Pre-Alert email.
const PREALERT_TO_RECIPIENTS = [
	'panggih.harto@spxexpress.com',
	'afrid.febrian@spxexpress.com',
	'a.rosyid@spxexpress.com',
	'hendry.saputra@spxexpress.com',
	'bas_benediktus.manurung@spx-external.com',
	'doddy.laksono@spx-external.com',
	'bas_m.anugrah@spx-external.com',
	'fg_nana.suryana@spx-external.com',
	'bas_dimas.pambudi@shopee-external.com'
];
const PREALERT_CC_RECIPIENTS = [
	{ name: 'Ahmad Syafiq Salafi', email: 'ahmad.salafi@spxexpress.com' },
	{ name: 'Aris Rahman Saputra', email: 'aris.saputra@spxexpress.com' },
	{ name: 'Atikah Suharti', email: 'atikah.suharti@shopee-xpress.com' },
	{ name: 'Audry Nissa Oktawina', email: 'audry.oktawina@spxexpress.com' },
	{ name: 'Ega Apriliawan', email: 'ega.apriliawan@spxexpress.com' },
	{ name: 'Egi Hendriko', email: 'egi.hendriko@spxexpress.com' },
	{ name: 'Eki Maulana', email: 'eki.maulana@spxexpress.com' },
	{ name: 'Ermi Nurwiati', email: 'ermi.nurwiati@spxexpress.com' },
	{ name: 'Hartias Rizalina', email: 'hartias.rizalina@spxexpress.com' },
	{ name: 'Inka Novianti Tarigan', email: 'inka.tarigan@spxexpress.com' },
	{ name: 'Saepul Bahri', email: 'ipi_saepul.bahri@spx-external.com' },
	{ name: 'Irawan Tri Atmodjo', email: 'irawan.tri@spxexpress.com' },
	{ name: 'Rifky Alfandi', email: 'rifky.alfandi@spxexpress.com' },
	{ name: 'Salma Farah Yuliani', email: 'salma.yuliani@spxexpress.com' },
	{ name: 'Siska Febriyanti', email: 'siska.febriyanti@spxexpress.com' },
	{ name: 'Sugama Sugama', email: 'ugbm_sugama.sugama@spx-external.com' },
	{ name: 'Virgiawan Bachtiar', email: 'virgiawan.bachtiar@spxexpress.com' },
	{ name: 'Feizal Umar Syah', email: 'feizal.syah@spx-external.com' },
	{ name: 'Wahyu Fitri Astuti', email: 'wahyu.astuti@spxexpress.com' }
];
let prealertRuntimeAccessToken = PREALERT_DRAFT_API_TOKEN;

function loadCachedPrealertAccessToken() {
	if (typeof window === 'undefined' || !window.localStorage) {
		return '';
	}

	try {
		const raw = window.localStorage.getItem(PREALERT_TOKEN_CACHE_KEY);
		if (!raw) {
			return '';
		}

		const cached = JSON.parse(raw);
		const accessToken = String(cached?.accessToken || '').trim();
		const expiresAt = Number(cached?.expiresAt || 0);
		if (!accessToken || !expiresAt) {
			return '';
		}

		// Keep a small buffer so token is refreshed before it actually expires.
		if (Date.now() >= (expiresAt - 60000)) {
			return '';
		}

		return accessToken;
	} catch (error) {
		return '';
	}
}

function saveCachedPrealertAccessToken(accessToken, expiresInSeconds) {
	if (typeof window === 'undefined' || !window.localStorage) {
		return;
	}

	try {
		const ttlSeconds = Math.max(Number(expiresInSeconds || 3600), 120);
		const expiresAt = Date.now() + (ttlSeconds * 1000);
		window.localStorage.setItem(PREALERT_TOKEN_CACHE_KEY, JSON.stringify({
			accessToken: String(accessToken || ''),
			expiresAt
		}));
	} catch (error) {
		// Ignore storage errors.
	}
}

function clearCachedPrealertAccessToken() {
	if (typeof window === 'undefined' || !window.localStorage) {
		return;
	}

	try {
		window.localStorage.removeItem(PREALERT_TOKEN_CACHE_KEY);
	} catch (error) {
		// Ignore storage errors.
	}
}

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
const dateDisplayText = document.getElementById('dateDisplayText');
const reportSummary = document.getElementById('reportSummary');
const tabMassUpload = document.getElementById('tabMassUpload');
const tabTOReport = document.getElementById('tabTOReport');
const massUploadPanel = document.getElementById('massUploadPanel');
const toReportPanel = document.getElementById('toReportPanel');
const operatorInput = document.getElementById('operatorInput');
const operatorCombobox = document.getElementById('operatorCombobox');
const operatorDropdown = document.getElementById('operatorDropdown');
const selectedOperatorTags = document.getElementById('selectedOperatorTags');
const operatorDetectedCount = document.getElementById('operatorDetectedCount');
const tripDetectedCount = document.getElementById('tripDetectedCount');

const copyTrackingBtn = document.getElementById('copyTrackingBtn');
const copyLinehaulReportBtn = document.getElementById('copyLinehaulReportBtn');
const copyBulkyReportBtn = document.getElementById('copyBulkyReportBtn');
const copyHourlyReportBtn = document.getElementById('copyHourlyReportBtn');
const copyPickupReportTemplateBtn = document.getElementById('copyPickupReportTemplateBtn');

const generateAllBtn = document.getElementById('generateAllBtn');
const generateAllInlineBtn = document.getElementById('generateAllInlineBtn');
const generateEmailFab = document.getElementById('generateEmailFab');
const floatingMenu = document.getElementById('floatingMenu');
const floatingMenuToggle = document.getElementById('floatingMenuToggle');
const floatingMenuDropdown = document.getElementById('floatingMenuDropdown');
const totalDataEl = document.getElementById('totalData');
const filteredDataEl = document.getElementById('filteredData');

const toStatusSummary = document.getElementById('toStatusSummary');
const toStatusSummaryBody = document.getElementById('toStatusSummaryBody');
const viewAllPreviewBtn = document.getElementById('viewAllPreviewBtn');
const errorMsg = document.getElementById('errorMsg');
const successMsg = document.getElementById('successMsg');
const warningMsg = document.getElementById('warningMsg');
const progressList = document.getElementById('progressList');
const toast = document.getElementById('toast');
const notificationLogBtn = document.getElementById('notificationLogBtn');
const notificationLogCount = document.getElementById('notificationLogCount');
const notificationLogOverlay = document.getElementById('notificationLogOverlay');
const notificationLogCloseBtn = document.getElementById('notificationLogCloseBtn');
const notificationLogClearBtn = document.getElementById('notificationLogClearBtn');
const notificationLogList = document.getElementById('notificationLogList');
const notificationLogEmpty = document.getElementById('notificationLogEmpty');
const themeToggle = document.getElementById('themeToggle');
const prealertPdfInput = document.getElementById('prealertPdfInput');
const prealertUploadBox = document.getElementById('prealertUploadBox');
const prealertUploadInfo = document.getElementById('prealertUploadInfo');
const prealertUploadStage = document.getElementById('prealertUploadStage');
const prealertUploadProgressFill = document.getElementById('prealertUploadProgressFill');
const prealertUploadStateIcon = document.getElementById('prealertUploadStateIcon');
const prealertUploadStateTitle = document.getElementById('prealertUploadStateTitle');
const prealertUploadStateDesc = document.getElementById('prealertUploadStateDesc');
const emailSubjectEl = document.getElementById('emailSubject');
const emailBodyEl = document.getElementById('emailBody');
const generateGmailDraftBtn = document.getElementById('generateGmailDraftBtn');
const downloadPrealertReportBtn = document.getElementById('downloadPrealertReportBtn');
const downloadPrealertFirstPageJpgBtn = document.getElementById('downloadPrealertFirstPageJpgBtn');
const driverNameEl = document.getElementById('driverName');
const plateNumberEl = document.getElementById('plateNumber');
const destinationEl = document.getElementById('destination');
const tripNumberEl = document.getElementById('tripNumber');
const tripTimeEl = document.getElementById('tripTime');
const ltNumberEl = document.getElementById('ltNumber');
const hvQtyEl = document.getElementById('hvQty');
const totalToQtyEl = document.getElementById('totalToQty');
const liquidationQtyEl = document.getElementById('liquidationQty');
const orderQtyEl = document.getElementById('orderQty');
const massUploadModal = document.getElementById('massUploadModal');
const closeMassUploadModalBtn = document.getElementById('closeMassUploadModalBtn');
const massUploadConfirmOverlay = document.getElementById('massUploadConfirmOverlay');
const massUploadConfirmCancelBtn = document.getElementById('massUploadConfirmCancelBtn');
const massUploadConfirmYesBtn = document.getElementById('massUploadConfirmYesBtn');
const massUploadEditorBody = document.getElementById('massUploadEditorBody');
const massUploadStatus = document.getElementById('massUploadStatus');
const fixMassUploadDimensionsBtn = document.getElementById('fixMassUploadDimensionsBtn');
const applyMassUploadChangesBtn = document.getElementById('applyMassUploadChangesBtn');
const cancelMassUploadChangesBtn = document.getElementById('cancelMassUploadChangesBtn');
const hubSelect = document.getElementById('hubSelect');
const hubCustomInput = document.getElementById('hubCustomInput');
const hubSelectWrapper = document.getElementById('hubSelectWrapper');
const headerSubtitle = document.getElementById('headerSubtitle');
const prefersReducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

/**
 * Returns the currently selected hub name.
 * If "Lainnya" is chosen and the user typed a custom name, that name is returned.
 * Otherwise defaults to "Ciputat 4 First Mile Hub".
 */
function getSelectedHubName() {
	if (!hubSelect) {
		return 'Ciputat 4 First Mile Hub';
	}

	if (hubSelect.value === '__custom__') {
		const custom = normalize(hubCustomInput?.value);
		return custom || 'First Mile Hub';
	}

	return hubSelect.value || 'Ciputat 4 First Mile Hub';
}

/**
 * Returns an abbreviated hub name suitable for email subjects and short references.
 * e.g. "Ciputat 4 First Mile Hub" -> "CIPUTAT 4 FIRST MILE"
 */
function getHubShortName() {
	const full = getSelectedHubName();
	return full.replace(/\s*Hub\s*$/i, '').trim();
}

/**
 * Returns the AMH-prefix version: "Ciputat 4 AMH" style.
 * Extracts hub identifier before "First Mile" if present.
 */
function getHubAmhName() {
	const full = getSelectedHubName();
	const match = full.match(/^(.+?)\s*First\s*Mile/i);
	return match ? `${match[1].trim()} AMH` : full.replace(/\s*Hub\s*$/i, '').trim() + ' AMH';
}

function escapeHtml(value) {
	return String(value ?? '')
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
		.replaceAll('"', '&quot;')
		.replaceAll("'", '&#39;');
}

function getNotificationAudioContext() {
	if (notificationAudioContext) {
		return notificationAudioContext;
	}

	const AudioContextClass = window.AudioContext || window.webkitAudioContext;
	if (!AudioContextClass) {
		return null;
	}

	notificationAudioContext = new AudioContextClass();
	return notificationAudioContext;
}

function getCustomErrorAudioElement() {
	const soundUrl = normalize(ERROR_NOTIFICATION_SOUND_URL);
	if (!soundUrl) {
		return null;
	}

	if (!customErrorAudioElement || !customErrorAudioElement.src.includes(soundUrl)) {
		customErrorAudioElement = new Audio(soundUrl);
		customErrorAudioElement.preload = 'auto';
	}

	return customErrorAudioElement;
}

function unlockNotificationAudio() {
	const audioContext = getNotificationAudioContext();
	if (audioContext?.state === 'suspended') {
		audioContext.resume().catch(() => {});
	}

	const customAudio = getCustomErrorAudioElement();
	if (!customAudio) {
		return;
	}

	customAudio.muted = true;
	customAudio.currentTime = 0;
	const warmupPromise = customAudio.play();
	if (warmupPromise && typeof warmupPromise.then === 'function') {
		warmupPromise
			.then(() => {
				customAudio.pause();
				customAudio.currentTime = 0;
				customAudio.muted = false;
			})
			.catch(() => {
				customAudio.muted = false;
			});
	}
}

function playNotificationSound(type) {
	if (!['success', 'error', 'warning'].includes(type)) {
		return;
	}

	const now = Date.now();
	if (now - lastNotificationSoundAt < 220) {
		return;
	}

	if (type === 'error') {
		const customAudio = getCustomErrorAudioElement();
		if (customAudio) {
			try {
				customAudio.pause();
				customAudio.currentTime = 0;
				customAudio.volume = 1;
				customAudio.muted = false;
				const playPromise = customAudio.play();
				if (playPromise && typeof playPromise.catch === 'function') {
					playPromise.catch(() => {});
				}
				lastNotificationSoundAt = now;
				return;
			} catch (error) {
				// Fall back to synthesized error tone below.
			}
		}
	}

	const audioContext = getNotificationAudioContext();
	if (!audioContext) {
		return;
	}

	lastNotificationSoundAt = now;

	if (audioContext.state === 'suspended') {
		audioContext.resume().catch(() => {});
	}

	const baseTime = audioContext.currentTime;
	const gainNode = audioContext.createGain();
	gainNode.connect(audioContext.destination);
	gainNode.gain.setValueAtTime(0.0001, baseTime);

	if (type === 'success') {
		const firstOsc = audioContext.createOscillator();
		const secondOsc = audioContext.createOscillator();
		firstOsc.type = 'sine';
		secondOsc.type = 'triangle';
		firstOsc.frequency.setValueAtTime(740, baseTime);
		secondOsc.frequency.setValueAtTime(988, baseTime + 0.09);

		firstOsc.connect(gainNode);
		secondOsc.connect(gainNode);

		gainNode.gain.linearRampToValueAtTime(0.075, baseTime + 0.02);
		gainNode.gain.linearRampToValueAtTime(0.055, baseTime + 0.1);
		gainNode.gain.exponentialRampToValueAtTime(0.0001, baseTime + 0.28);

		firstOsc.start(baseTime);
		firstOsc.stop(baseTime + 0.14);
		secondOsc.start(baseTime + 0.09);
		secondOsc.stop(baseTime + 0.28);
		return;
	}

	if (type === 'warning') {
		const warningOscA = audioContext.createOscillator();
		const warningOscB = audioContext.createOscillator();
		warningOscA.type = 'square';
		warningOscB.type = 'square';
		warningOscA.frequency.setValueAtTime(720, baseTime);
		warningOscB.frequency.setValueAtTime(720, baseTime + 0.16);
		warningOscA.connect(gainNode);
		warningOscB.connect(gainNode);

		gainNode.gain.linearRampToValueAtTime(0.09, baseTime + 0.01);
		gainNode.gain.exponentialRampToValueAtTime(0.0001, baseTime + 0.12);
		gainNode.gain.setValueAtTime(0.0001, baseTime + 0.13);
		gainNode.gain.linearRampToValueAtTime(0.085, baseTime + 0.17);
		gainNode.gain.exponentialRampToValueAtTime(0.0001, baseTime + 0.3);

		warningOscA.start(baseTime);
		warningOscA.stop(baseTime + 0.12);
		warningOscB.start(baseTime + 0.16);
		warningOscB.stop(baseTime + 0.3);
		return;
	}

	const emitOinkFallback = () => {
		const synthBaseTime = audioContext.currentTime;
		const synthGainNode = audioContext.createGain();
		synthGainNode.connect(audioContext.destination);
		synthGainNode.gain.setValueAtTime(0.0001, synthBaseTime);

		// Build a very clear "oink-oink" style error sound.
		const oinkTimes = [synthBaseTime, synthBaseTime + 0.18];

		for (const startAt of oinkTimes) {
			const oscMain = audioContext.createOscillator();
			const oscSub = audioContext.createOscillator();
			const formantFilter = audioContext.createBiquadFilter();
			const noiseFilter = audioContext.createBiquadFilter();
			const mixGain = audioContext.createGain();
			const noiseGain = audioContext.createGain();

			oscMain.type = 'square';
			oscSub.type = 'triangle';
			oscMain.frequency.setValueAtTime(260, startAt);
			oscMain.frequency.exponentialRampToValueAtTime(170, startAt + 0.05);
			oscMain.frequency.exponentialRampToValueAtTime(95, startAt + 0.14);
			oscSub.frequency.setValueAtTime(190, startAt);
			oscSub.frequency.exponentialRampToValueAtTime(110, startAt + 0.14);
			oscSub.detune.setValueAtTime(-16, startAt);

			formantFilter.type = 'bandpass';
			formantFilter.frequency.setValueAtTime(640, startAt);
			formantFilter.Q.setValueAtTime(2.8, startAt);

			mixGain.gain.setValueAtTime(0.0001, startAt);
			mixGain.gain.linearRampToValueAtTime(0.42, startAt + 0.01);
			mixGain.gain.linearRampToValueAtTime(0.22, startAt + 0.06);
			mixGain.gain.exponentialRampToValueAtTime(0.0001, startAt + 0.16);

			noiseFilter.type = 'highpass';
			noiseFilter.frequency.setValueAtTime(560, startAt);
			noiseFilter.Q.setValueAtTime(0.7, startAt);

			noiseGain.gain.setValueAtTime(0.0001, startAt);
			noiseGain.gain.linearRampToValueAtTime(0.08, startAt + 0.01);
			noiseGain.gain.exponentialRampToValueAtTime(0.0001, startAt + 0.11);

			const noiseBuffer = audioContext.createBuffer(1, Math.floor(audioContext.sampleRate * 0.18), audioContext.sampleRate);
			const noiseData = noiseBuffer.getChannelData(0);
			for (let i = 0; i < noiseData.length; i += 1) {
				noiseData[i] = (Math.random() * 2 - 1) * 0.9;
			}
			const noiseSource = audioContext.createBufferSource();
			noiseSource.buffer = noiseBuffer;

			oscMain.connect(formantFilter);
			oscSub.connect(formantFilter);
			formantFilter.connect(mixGain);
			mixGain.connect(synthGainNode);

			noiseSource.connect(noiseFilter);
			noiseFilter.connect(noiseGain);
			noiseGain.connect(synthGainNode);

			oscMain.start(startAt);
			oscSub.start(startAt + 0.003);
			noiseSource.start(startAt);
			oscMain.stop(startAt + 0.16);
			oscSub.stop(startAt + 0.16);
			noiseSource.stop(startAt + 0.16);
		}

		synthGainNode.gain.linearRampToValueAtTime(0.78, synthBaseTime + 0.012);
		synthGainNode.gain.linearRampToValueAtTime(0.62, synthBaseTime + 0.18);
		synthGainNode.gain.exponentialRampToValueAtTime(0.0001, synthBaseTime + 0.44);
	};

	emitOinkFallback();
}

function hideToast() {
	if (toastTimer) {
		clearTimeout(toastTimer);
		toastTimer = null;
	}

	if (toastHideTimer) {
		clearTimeout(toastHideTimer);
		toastHideTimer = null;
	}

	if (!toast.classList.contains('show')) {
		toast.className = 'toast';
		toast.innerHTML = '';
		return;
	}

	toast.classList.remove('show');
	toast.classList.add('is-hiding');

	toastHideTimer = setTimeout(() => {
		toast.className = 'toast';
		toast.innerHTML = '';
		toastHideTimer = null;
	}, 320);
}

function formatNotificationTime(timestamp) {
	return new Date(timestamp).toLocaleTimeString('id-ID', {
		hour: '2-digit',
		minute: '2-digit',
		second: '2-digit'
	});
}

function updateNotificationLogCount() {
	if (!notificationLogCount) {
		return;
	}

	const total = notificationLogEntries.length;
	notificationLogCount.textContent = total > 99 ? '99+' : String(total);
	notificationLogCount.classList.toggle('is-hidden', total === 0);
}

function renderNotificationLog() {
	if (!notificationLogList || !notificationLogEmpty) {
		return;
	}

	if (!notificationLogEntries.length) {
		notificationLogList.innerHTML = '';
		notificationLogEmpty.style.display = 'block';
		updateNotificationLogCount();
		return;
	}

	notificationLogEmpty.style.display = 'none';
	notificationLogList.innerHTML = notificationLogEntries.map((entry) => {
		const detailsHtml = entry.details.length
			? `<div class="notification-log-details">${entry.details.map((detail) => `<div>${escapeHtml(detail)}</div>`).join('')}</div>`
			: '';

		return `
			<li class="notification-log-item log-${escapeHtml(entry.type)}">
				<div class="notification-log-meta">
					<span>${escapeHtml(entry.label)}</span>
					<time>${formatNotificationTime(entry.timestamp)}</time>
				</div>
				<div class="notification-log-title">${escapeHtml(entry.title)}</div>
				${entry.message ? `<div class="notification-log-message">${escapeHtml(entry.message)}</div>` : ''}
				${detailsHtml}
			</li>
		`;
	}).join('');

	updateNotificationLogCount();
}

function addNotificationLogEntry(config) {
	const type = config.type || 'info';
	const labels = {
		success: 'Success',
		error: 'Error',
		warning: 'Warning',
		processing: 'Processing',
		info: 'Info'
	};

	notificationLogEntries.unshift({
		type,
		label: labels[type] || 'Info',
		title: config.title || (type === 'error' ? 'Error' : 'Notification'),
		message: config.message || '',
		details: Array.isArray(config.details) ? config.details.map((item) => String(item ?? '')) : [],
		timestamp: Date.now()
	});

	if (notificationLogEntries.length > 120) {
		notificationLogEntries = notificationLogEntries.slice(0, 120);
	}

	renderNotificationLog();
}

function openNotificationLog() {
	if (!notificationLogOverlay) {
		return;
	}

	openModalOverlay(notificationLogOverlay);
}

function closeNotificationLog() {
	if (!notificationLogOverlay) {
		return;
	}

	closeModalOverlay(notificationLogOverlay);
}

function clearNotificationLog() {
	notificationLogEntries = [];
	renderNotificationLog();
}

function showToast(messageOrConfig, type = 'success') {
	if (toastTimer) {
		clearTimeout(toastTimer);
		toastTimer = null;
	}

	if (toastHideTimer) {
		clearTimeout(toastHideTimer);
		toastHideTimer = null;
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
		: `<span class="toast-icon" aria-hidden="true">${escapeHtml(config.icon || (toastType === 'error' ? '\u2716' : '\u2714'))}</span>`;

	addNotificationLogEntry({
		type: toastType,
		title,
		message,
		details
	});

	const detailsHtml = details.length
		? `<ul class="toast-steps">${details.map((item) => `<li>${escapeHtml(item)}</li>`).join('')}</ul>`
		: '';

	toast.innerHTML = `
		<div class="toast-card">
			<div class="toast-head">
				<div class="toast-head-left">${icon}<strong>${escapeHtml(title)}</strong></div>
				<button type="button" class="toast-close" aria-label="Close notification"><svg width="14" height="14" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" fill="rgba(0,0,0,0.08)" stroke="currentColor" stroke-width="2"/><path d="M15 9L9 15M9 9l6 6" stroke="currentColor" stroke-width="2.6" stroke-linecap="round"/></svg></button>
			</div>
			${message ? `<div class="toast-message">${escapeHtml(message)}</div>` : ''}
			${detailsHtml}
		</div>
	`;

	toast.className = `toast toast-${toastType}`;
	void toast.offsetWidth;
	toast.classList.add('show');
	playNotificationSound(toastType);

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
		icon: '\u2714'
	});
}

function showReportErrorToast() {
	showToast({
		type: 'error',
		title: 'Report Generation Failed',
		message: 'Something went wrong while generating the report.',
		duration: 4000,
		icon: '\u2716'
	});
}

function showWarningToast(message) {
	showToast({
		type: 'warning',
		title: 'Warning',
		message,
		duration: 4000,
		icon: '\u26A0'
	});
}

function resetMessages() {
	if (errorMsg) {
		errorMsg.style.display = 'none';
	}
	if (successMsg) {
		successMsg.style.display = 'none';
	}
	if (warningMsg) {
		warningMsg.style.display = 'none';
	}
}

function setError(message) {
	showToast(message, 'error');
}

function setSuccess(message) {
	showToast(message, 'success');
}

function setWarning(message) {
	if (warningMsg) {
		warningMsg.textContent = message;
	}
}

function clearWarning() {
	if (warningMsg) {
		warningMsg.style.display = 'none';
	}
}

function updateBodyScrollLock() {
	const hasOpenModal = massUploadModal.classList.contains('open')
		|| massUploadModal.classList.contains('is-closing')
		|| massUploadConfirmOverlay.classList.contains('open')
		|| massUploadConfirmOverlay.classList.contains('is-closing')
		|| notificationLogOverlay.classList.contains('open')
		|| notificationLogOverlay.classList.contains('is-closing');

	document.body.style.overflow = hasOpenModal ? 'hidden' : '';
}

function openModalOverlay(overlay) {
	const activeTimer = modalCloseTimers.get(overlay);
	if (activeTimer) {
		clearTimeout(activeTimer);
		modalCloseTimers.delete(overlay);
	}

	overlay.classList.remove('is-closing');
	overlay.classList.add('open');
	overlay.setAttribute('aria-hidden', 'false');
	updateBodyScrollLock();
}

function closeModalOverlay(overlay) {
	if (!overlay.classList.contains('open') && !overlay.classList.contains('is-closing')) {
		return;
	}

	overlay.classList.remove('open');
	overlay.classList.add('is-closing');
	overlay.setAttribute('aria-hidden', 'true');

	const activeTimer = modalCloseTimers.get(overlay);
	if (activeTimer) {
		clearTimeout(activeTimer);
	}

	const closeTimer = setTimeout(() => {
		overlay.classList.remove('is-closing');
		modalCloseTimers.delete(overlay);
		updateBodyScrollLock();
	}, 220);

	modalCloseTimers.set(overlay, closeTimer);
}

function openMassUploadModal() {
	openModalOverlay(massUploadModal);
	setMassUploadStatus('Edit dimensions, then click Apply Changes.', 'info');
}

function closeMassUploadModal() {
	closeModalOverlay(massUploadConfirmOverlay);
	closeModalOverlay(massUploadModal);
	setMassUploadStatus('Ready to edit dimensions.', 'info');
}

function openMassUploadCloseConfirmation() {
	openModalOverlay(massUploadConfirmOverlay);
}

function closeMassUploadCloseConfirmation() {
	closeModalOverlay(massUploadConfirmOverlay);
}

function normalize(value) {
	return String(value ?? '').trim();
}

function resetLinehaulTemplateData() {
	linehaulTemplateData = {
		staTime: '-',
		ataTime: '-',
		stdTime: '-',
		stfTime: '-',
		atdTime: '-',
		sealCode: '-'
	};
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
	if (!progressList) return;
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
	if (!progressList) return;
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

function isPdfFile(file) {
	if (!file) {
		return false;
	}
	const type = normalize(file.type).toLowerCase();
	const name = normalize(file.name).toLowerCase();
	return type === 'application/pdf' || name.endsWith('.pdf');
}

function setPrealertUploadInfo(message, tone = 'info') {
	if (!prealertUploadInfo) {
		return;
	}

	prealertUploadInfo.textContent = message;
	prealertUploadInfo.classList.remove('success', 'error');
	if (tone === 'success' || tone === 'error') {
		prealertUploadInfo.classList.add(tone);
	}
}

function setPrealertUploadState(state = 'idle', progress = 0, stageText = 'Ready', detailText = 'Pilih file untuk mulai parsing.') {
	if (prealertUploadBox) {
		prealertUploadBox.dataset.uploadState = state;
	}

	if (prealertUploadStage) {
		prealertUploadStage.textContent = stageText;
		prealertUploadStage.className = 'prealert-upload-stage';
		if (state === 'uploading') {
			prealertUploadStage.classList.add('state-uploading');
		}
		if (state === 'parsing') {
			prealertUploadStage.classList.add('state-parsing');
		}
		if (state === 'success') {
			prealertUploadStage.classList.add('state-success');
		}
		if (state === 'error') {
			prealertUploadStage.classList.add('state-error');
		}
	}

	if (prealertUploadProgressFill) {
		const safeProgress = Math.max(0, Math.min(100, Number(progress) || 0));
		prealertUploadProgressFill.style.width = `${safeProgress}%`;
		const progressBar = prealertUploadProgressFill.parentElement;
		if (progressBar) {
			progressBar.setAttribute('aria-valuenow', String(Math.round(safeProgress)));
		}
	}

	if (prealertUploadStateTitle) {
		prealertUploadStateTitle.textContent = stageText;
	}

	if (prealertUploadStateDesc) {
		prealertUploadStateDesc.textContent = detailText;
	}

	if (prealertUploadStateIcon) {
		if (state === 'success') {
			prealertUploadStateIcon.innerHTML = '<svg width="22" height="22" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" fill="#C8E6C9" stroke="#2D2D2D" stroke-width="2.2"/><polyline points="8 12.5 11 15.5 16.5 9" fill="none" stroke="#2D2D2D" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/></svg>';
		} else if (state === 'error') {
			prealertUploadStateIcon.innerHTML = '<svg width="22" height="22" viewBox="0 0 24 24"><path d="M12 2L1.5 20h21z" fill="#FFF9C4" stroke="#2D2D2D" stroke-width="2.2" stroke-linejoin="round"/><line x1="12" y1="9" x2="12" y2="14" stroke="#E74C3C" stroke-width="2.6" stroke-linecap="round"/><circle cx="12" cy="17" r="1.2" fill="#E74C3C"/></svg>';
		} else if (state === 'uploading' || state === 'parsing') {
			prealertUploadStateIcon.innerHTML = '<svg class="spin-icon" width="22" height="22" viewBox="0 0 24 24"><circle cx="12" cy="12" r="9" fill="none" stroke="#FFE0B2" stroke-width="3"/><path d="M12 3a9 9 0 0 1 9 9" fill="none" stroke="#FF6A00" stroke-width="3" stroke-linecap="round"/></svg>';
		} else {
			prealertUploadStateIcon.innerHTML = '<svg width="22" height="22" viewBox="0 0 24 24"><path d="M14 2H7a3 3 0 0 0-3 3v14a3 3 0 0 0 3 3h10a3 3 0 0 0 3-3V8z" fill="#FFF3E0" stroke="#2D2D2D" stroke-width="2" stroke-linejoin="round"/><path d="M14 2v6a2 2 0 0 0 2 2h4" fill="#FFD966" stroke="#2D2D2D" stroke-width="1.8" stroke-linejoin="round"/></svg>';
		}
	}
}

function resetPrealertExtractionTable() {
	const fields = [
		driverNameEl,
		plateNumberEl,
		destinationEl,
		tripNumberEl,
		tripTimeEl,
		ltNumberEl,
		hvQtyEl,
		totalToQtyEl,
		liquidationQtyEl,
		orderQtyEl
	];

	fields.forEach((field) => {
		if (field) {
			field.textContent = '-';
		}
	});

	resetLinehaulTemplateData();

	syncPrealertTripFromReportSlot();
	updatePrealertEmailPreview();
}

function getTextValue(element) {
	return normalize(element?.textContent) || '-';
}

function getReportSelectionContext() {
	return {
		trip: getSlotTripNumber(slotSelect?.value),
		slot: normalize(slotSelect?.value) || '-',
		reportDate: getSelectedReportDateFormatted() || '-'
	};
}

function getSlotTripNumber(slotValue) {
	const match = normalize(slotValue).match(/(\d+)/);
	return match?.[1] || '-';
}

function syncPrealertTripFromReportSlot() {
	if (!tripNumberEl) {
		return;
	}
	tripNumberEl.textContent = getSlotTripNumber(slotSelect?.value);
}

function updateDateDisplay() {
	if (!dateDisplayText || !reportDateInput) return;
	const val = reportDateInput.value;
	if (!val || !/^\d{4}-\d{2}-\d{2}$/.test(val)) {
		dateDisplayText.textContent = '—';
		return;
	}
	const bulan = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
	const [y, m, d] = val.split('-');
	dateDisplayText.textContent = `${parseInt(d, 10)} ${bulan[parseInt(m, 10) - 1]} ${y}`;
}

function getIndonesianReportDate() {
	const rawDate = normalize(reportDateInput?.value);
	if (!/^\d{4}-\d{2}-\d{2}$/.test(rawDate)) {
		return '-';
	}

	const [year, month, day] = rawDate.split('-');
	const monthNames = [
		'Januari',
		'Februari',
		'Maret',
		'April',
		'Mei',
		'Juni',
		'Juli',
		'Agustus',
		'September',
		'Oktober',
		'November',
		'Desember'
	];

	const monthIndex = Number.parseInt(month, 10) - 1;
	const monthLabel = monthNames[monthIndex] || month;
	return `${day}-${monthLabel}-${year}`;
}

function getIndonesianReportDateWithDayUppercase() {
	const rawDate = normalize(reportDateInput?.value);
	if (!/^\d{4}-\d{2}-\d{2}$/.test(rawDate)) {
		return '-';
	}

	const [year, month, day] = rawDate.split('-');
	const dayNames = [
		'MINGGU',
		'SENIN',
		'SELASA',
		'RABU',
		'KAMIS',
		'JUMAT',
		'SABTU'
	];
	const monthNames = [
		'JANUARI',
		'FEBRUARI',
		'MARET',
		'APRIL',
		'MEI',
		'JUNI',
		'JULI',
		'AGUSTUS',
		'SEPTEMBER',
		'OKTOBER',
		'NOVEMBER',
		'DESEMBER'
	];

	const monthNumber = Number.parseInt(month, 10);
	const dayNumber = Number.parseInt(day, 10);
	const yearNumber = Number.parseInt(year, 10);
	const dateObject = new Date(yearNumber, monthNumber - 1, dayNumber);
	const dayLabel = dayNames[dateObject.getDay()] || '-';
	const monthLabel = monthNames[monthNumber - 1] || month;
	const formattedDay = Number.isFinite(dayNumber) ? String(dayNumber) : day;

	return `${dayLabel} ${formattedDay} ${monthLabel} ${year}`;
}

function getIndonesianReportDateWithDayTitleCase() {
	const rawDate = normalize(reportDateInput?.value);
	if (!/^\d{4}-\d{2}-\d{2}$/.test(rawDate)) {
		return '-';
	}

	const [year, month, day] = rawDate.split('-');
	const dayNames = [
		'Minggu',
		'Senin',
		'Selasa',
		'Rabu',
		'Kamis',
		'Jumat',
		'Sabtu'
	];
	const monthNames = [
		'Januari',
		'Februari',
		'Maret',
		'April',
		'Mei',
		'Juni',
		'Juli',
		'Agustus',
		'September',
		'Oktober',
		'November',
		'Desember'
	];

	const monthNumber = Number.parseInt(month, 10);
	const dayNumber = Number.parseInt(day, 10);
	const yearNumber = Number.parseInt(year, 10);
	const dateObject = new Date(yearNumber, monthNumber - 1, dayNumber);
	const dayLabel = dayNames[dateObject.getDay()] || '-';
	const monthLabel = monthNames[monthNumber - 1] || month;

	return `${dayLabel}, ${day} ${monthLabel} ${year}`;
}

function getIndonesianReportDateWithDayTitleCaseNatural() {
	const rawDate = normalize(reportDateInput?.value);
	if (!/^\d{4}-\d{2}-\d{2}$/.test(rawDate)) {
		return '-';
	}

	const [year, month, day] = rawDate.split('-');
	const dayNames = [
		'Minggu',
		'Senin',
		'Selasa',
		'Rabu',
		'Kamis',
		'Jumat',
		'Sabtu'
	];
	const monthNames = [
		'Januari',
		'Februari',
		'Maret',
		'April',
		'Mei',
		'Juni',
		'Juli',
		'Agustus',
		'September',
		'Oktober',
		'November',
		'Desember'
	];

	const monthNumber = Number.parseInt(month, 10);
	const dayNumber = Number.parseInt(day, 10);
	const yearNumber = Number.parseInt(year, 10);
	const dateObject = new Date(yearNumber, monthNumber - 1, dayNumber);
	const dayLabel = dayNames[dateObject.getDay()] || '-';
	const monthLabel = monthNames[monthNumber - 1] || month;
	const formattedDay = Number.isFinite(dayNumber) ? String(dayNumber) : day;

	return `${dayLabel}, ${formattedDay} ${monthLabel} ${year}`;
}

function getLinehaulQtyFromReportData() {
	const candidates = [
		getTextValue(orderQtyEl),
		getTextValue(totalToQtyEl),
		getTextValue(hvQtyEl)
	];

	const found = candidates.find((value) => value && value !== '-');
	if (found) {
		return found;
	}

	const fallback = normalize(filteredDataEl?.textContent || totalDataEl?.textContent);
	return fallback || '-';
}

function getBulkyQtyFromReportData() {
	const qty = normalize(filteredDataEl?.textContent);
	return qty || '0';
}

function toTitleCaseWords(value) {
	const safeValue = normalize(value);
	if (!safeValue || safeValue === '-') {
		return '-';
	}

	return safeValue
		.toLowerCase()
		.replace(/(^|[\s/-])([a-z])/g, (match, separator, letter) => `${separator}${letter.toUpperCase()}`);
}

function extractClockTime(value) {
	const matched = String(value || '').match(/(\d{1,2}):(\d{2})(?::\d{2})?/);
	if (!matched) {
		return '-';
	}

	const hour = Number.parseInt(matched[1], 10);
	const minute = Number.parseInt(matched[2], 10);
	if (!Number.isFinite(hour) || !Number.isFinite(minute) || hour < 0 || hour > 23 || minute < 0 || minute > 59) {
		return '-';
	}

	return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
}

function extractTimeByLabelFromLines(lines, labelRegex, maxLookAhead = 8) {
	if (!Array.isArray(lines)) {
		return '-';
	}

	const boundaryRegex = /\b(?:Origin|Destination|ATA|STA|STD|STF|ATD|Waktu\s*Segel|Kode\s*Segel|Nama\s*Driver|Nomor\s*Polisi|Nama\s*Vendor|Jumlah\s*TO|#\s*Nomor\s*TO)\b/i;

	for (let index = 0; index < lines.length; index += 1) {
		const currentLine = normalize(lines[index]);
		if (!currentLine || !labelRegex.test(currentLine)) {
			continue;
		}

		const inlineDetected = extractClockTime(currentLine);
		if (inlineDetected !== '-') {
			return inlineDetected;
		}

		for (let offset = 1; offset <= maxLookAhead && (index + offset) < lines.length; offset += 1) {
			const nextLine = normalize(lines[index + offset]);
			if (!nextLine || /^[:\-]+$/.test(nextLine)) {
				continue;
			}

			const detected = extractClockTime(nextLine);
			if (detected !== '-') {
				return detected;
			}

			if (offset > 1 && boundaryRegex.test(nextLine) && !labelRegex.test(nextLine)) {
				break;
			}
		}
	}

	return '-';
}

function pickPreferredClockValue(values = []) {
	for (const value of values) {
		const parsed = extractClockTime(value);
		if (parsed !== '-') {
			return parsed;
		}
	}
	return '-';
}

function parseClockToMinutes(clockValue) {
	const clock = extractClockTime(clockValue);
	if (clock === '-') {
		return null;
	}

	const [hour, minute] = clock.split(':').map((part) => Number.parseInt(part, 10));
	if (!Number.isFinite(hour) || !Number.isFinite(minute)) {
		return null;
	}

	return hour * 60 + minute;
}

function formatMinutesAsClock(totalMinutes) {
	if (!Number.isFinite(totalMinutes)) {
		return '-';
	}

	const minutesInDay = 24 * 60;
	const normalized = ((totalMinutes % minutesInDay) + minutesInDay) % minutesInDay;
	const hour = Math.floor(normalized / 60);
	const minute = normalized % 60;
	return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
}

function shiftClockByMinutes(clockValue, deltaMinutes) {
	const baseMinutes = parseClockToMinutes(clockValue);
	if (baseMinutes === null) {
		return '-';
	}

	return formatMinutesAsClock(baseMinutes + Number(deltaMinutes || 0));
}

function formatClockWithEarly(actualClock, compareClock) {
	const formattedActual = extractClockTime(actualClock);
	if (formattedActual === '-') {
		return '-';
	}

	const actualMinutes = parseClockToMinutes(formattedActual);
	const compareMinutes = parseClockToMinutes(compareClock);
	if (actualMinutes === null || compareMinutes === null) {
		return formattedActual;
	}

	if (actualMinutes < compareMinutes) {
		return `${formattedActual} Early`;
	}

	return formattedActual;
}

function formatAtdWithEarly(atdClock, compareClock) {
	return formatClockWithEarly(atdClock, compareClock);
}

function buildLinehaulReportTemplateText() {
	const { slot } = getReportSelectionContext();
	const slotNumber = getSlotTripNumber(slot);
	const destination = toTitleCaseWords(getTextValue(destinationEl));
	const driverName = toTitleCaseWords(getTextValue(driverNameEl));
	const plateNumber = getTextValue(plateNumberEl);
	const ltNumber = getTextValue(ltNumberEl);
	const totalToQty = getTextValue(totalToQtyEl);
	const orderQty = getTextValue(orderQtyEl);
	const hvQty = getTextValue(hvQtyEl);
	const etaBaseTime = linehaulTemplateData.stdTime !== '-' ? linehaulTemplateData.stdTime : linehaulTemplateData.staTime;
	const etaTime = shiftClockByMinutes(etaBaseTime, -30);
	const ataTime = formatClockWithEarly(linehaulTemplateData.ataTime, etaTime);
	const stdTime = extractClockTime(linehaulTemplateData.stdTime);
	const atdCompareClock = linehaulTemplateData.stfTime !== '-' ? linehaulTemplateData.stfTime : linehaulTemplateData.stdTime;
	const atdTime = formatAtdWithEarly(linehaulTemplateData.atdTime, atdCompareClock);
	const sealCode = normalize(linehaulTemplateData.sealCode || '-').toUpperCase() || '-';
	const dateLabel = getIndonesianReportDateWithDayTitleCaseNatural();

	return [
		'Daily Report',
		getSelectedHubName(),
		dateLabel,
		'',
		`Slot : ${slotNumber}`,
		'',
		`Destination: ${destination}`,
		`ETA : ${etaTime}`,
		`ATA : ${ataTime}`,
		`STD : ${stdTime}`,
		`ATD : ${atdTime}`,
		'',
		`Nama Driver : ${driverName}`,
		`No Polisi : ${plateNumber}`,
		`No Seal : ${sealCode}`,
		`No. Linehaul Trip  : ${ltNumber}`,
		`Qty of TO : ${totalToQty}`,
		`Qty of Parcel : ${orderQty}`,
		`HV TO Quantity : ${hvQty}`,
		'',
		'Terima-Kasih "Good-Luck"'
	].join('\n');
}

function buildBulkyReportTemplateText() {
	const { slot } = getReportSelectionContext();
	const slotNumber = getSlotTripNumber(slot);
	const bulkyToCount = normalize(filteredDataEl?.textContent) || '0';
	const bulkyTotalQtyValue = bulkyTotalQty || 0;
	const dateLabel = getIndonesianReportDateWithDayTitleCase();

	return [
		'Mass Upload Bulky Measurement',
		getSelectedHubName(),
		dateLabel,
		'',
		`Slot : ${slotNumber}`,
		`Total of TO Bulky  : ${bulkyToCount}`,
		`Total of Qty Bulky : ${bulkyTotalQtyValue}`,
		'',
		'Terima-kasih~'
	].join('\n');
}

function buildHourlyReportTemplateText() {
	const dateLabel = getIndonesianReportDateWithDayTitleCase();
	return [
		'Hourly Performance Dialogue Dashboard',
		`${getSelectedHubName()}`,
		dateLabel
	].join('\n');
}

function buildPickupReportTemplateText() {
	const dateLabel = getIndonesianReportDateWithDayTitleCaseNatural();
	const hubName = getSelectedHubName();

	return [
		'*Pick Up Success Rate*',
		`*${hubName}*`,
		`*${dateLabel}*`,
		'',
		'Thanks ya, Team.',
		'Untuk yang *success rate-nya masih kurang*, yuk dibantu tingkatkan lagi dan dimaksimalkan follow up-nya di lapangan. Yang sudah *achieve*, tetap jaga konsistensinya ya biar performanya stabil.',
		'',
		'Kalau ada kendala di lapangan, langsung infoin aja ya supaya bisa cepat dibantu cari solusinya.',
		'',
		'Terima kasih buat effort-nya, Team.🙏🏻🚀'
	].join('\n');
}

async function copyReportTemplateText(templateType) {
	let text = '';
	let label = '';

	if (templateType === 'linehaul') {
		text = buildLinehaulReportTemplateText();
		label = 'Linehaul Report';
	}

	if (templateType === 'bulky') {
		text = buildBulkyReportTemplateText();
		label = 'Bulky Report';
	}

	if (templateType === 'hourly') {
		text = buildHourlyReportTemplateText();
		label = 'Hourly Report';
	}

	if (templateType === 'pickup') {
		text = buildPickupReportTemplateText();
		label = 'Pickup Report';
	}

	if (!text) {
		setError('Template report tidak valid.');
		return;
	}

	try {
		await navigator.clipboard.writeText(text);
		setSuccess(`${label} berhasil di-copy.`);
	} catch (error) {
		setError(`Gagal copy ${label}: ${error.message}`);
	}
}

function buildPrealertSubject() {
	const { slot } = getReportSelectionContext();
	const destination = getTextValue(destinationEl);
	const subjectDestination = destination === '-' ? 'DESTINATION' : destination.toUpperCase();
	const slotTripNumber = getSlotTripNumber(slot);
	const formattedDate = getIndonesianReportDate().replaceAll('-', ' ');

	return `PRE ALERT - AMH ${getHubShortName().toUpperCase()} TO ${subjectDestination} - TRIP ${slotTripNumber} - ${formattedDate}`;
}

function buildPrealertBody() {
	const { trip, slot, reportDate } = getReportSelectionContext();
	const destination = getTextValue(destinationEl);
	const driver = getTextValue(driverNameEl);
	const plate = getTextValue(plateNumberEl);
	const ltNumber = getTextValue(ltNumberEl);
	const tripTime = getTextValue(tripTimeEl);
	const hvQty = getTextValue(hvQtyEl);
	const totalToQty = getTextValue(totalToQtyEl);
	const liquidationQty = getTextValue(liquidationQtyEl);
	const orderQty = getTextValue(orderQtyEl);

	const rows = [
		['Nama Driver', driver],
		['No Polisi', plate],
		['Next Destination', destination],
		['Trip', trip],
		['Time', tripTime],
		['LT Number', ltNumber],
		['HV TO Quantity', hvQty],
		['Total TO Quantity', totalToQty],
		['Liquidation TO Quantity', liquidationQty],
		['Total Order Quantity', orderQty]
	];

	const htmlRows = rows
		.map(([label, value]) => `
			<tr>
				<td style="border:1px solid #d1d5db;padding:8px 10px;font-size:13px;font-weight:600;background:#f8fafc;width:40%;vertical-align:top;">${escapeHtml(label)}</td>
				<td style="border:1px solid #d1d5db;padding:8px 10px;font-size:13px;color:#0f172a;vertical-align:top;">${escapeHtml(value)}</td>
			</tr>
		`)
		.join('');

	const html = `
		<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;line-height:1.55;color:#0f172a;max-width:620px;">
			<p style="margin:0 0 4px;">Dear All</p>
			<p style="margin:0 0 10px;">Berikut Terlampir Surat Jalan From ${escapeHtml(getHubAmhName())} To ${escapeHtml(destination)}</p>
			<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;max-width:560px;border-collapse:collapse;margin:6px 0 10px;background:#ffffff;border:1px solid #d1d5db;">
				<thead>
					<tr>
						<th colspan="2" style="border:1px solid #d1d5db;padding:10px 12px;background:#8dc63f;color:#ffffff;text-align:center;font-size:14px;font-weight:700;letter-spacing:0.04em;">${escapeHtml(getHubShortName().toUpperCase())}</th>
					</tr>
				</thead>
				<tbody>
					${htmlRows}
				</tbody>
			</table>
			<p style="margin:0 0 6px;font-size:12px;line-height:1.4;color:#334155;"><b>Noted:</b> Apabila dalam waktu 3 jam setelah shipment tiba tidak terdapat feedback dari pihak penerima, maka seluruh tanggung jawab terkait paket dianggap telah diterima oleh next station.</p>
			<p style="margin:0;">Terima-kasih</p>
		</div>
	`;

	const previewHtmlRows = rows
		.map(([label, value]) => `
			<tr>
				<td style="border:2px solid #2D2D2D;padding:5px 8px;font-size:13px;font-weight:700;background:#8dc63f;color:#2D2D2D;width:40%;vertical-align:top;text-transform:uppercase;letter-spacing:0.02em;">${escapeHtml(label)}</td>
				<td style="border:2px solid #2D2D2D;padding:5px 8px;font-size:13px;color:#2D2D2D;vertical-align:top;background:#e8f5d9;">${escapeHtml(value)}</td>
			</tr>
		`)
		.join('');

	const previewHtml = `
		<div style="font-family:Fredoka,Segoe UI,Arial,sans-serif;font-size:14px;line-height:1.4;color:#2D2D2D;max-width:620px;">
			<p style="margin:0 0 4px;">Dear All</p>
			<p style="margin:0 0 8px;">Berikut Terlampir Surat Jalan From ${escapeHtml(getHubAmhName())} To ${escapeHtml(destination)}</p>
			<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;max-width:560px;border-collapse:collapse;margin:6px 0 10px;background:#ffffff;border:2.5px solid #2D2D2D;border-radius:14px;overflow:hidden;box-shadow:4px 4px 0 #2D2D2D;">
				<thead>
					<tr>
						<th colspan="2" style="border:2px solid #2D2D2D;padding:10px 12px;background:#5a9e2f;color:#ffffff;text-align:center;font-size:15px;font-weight:800;letter-spacing:0.06em;text-transform:uppercase;text-shadow:1px 1px 0 rgba(0,0,0,0.15);">${escapeHtml(getHubShortName().toUpperCase())}</th>
					</tr>
				</thead>
				<tbody>
					${previewHtmlRows}
				</tbody>
			</table>
			<p style="margin:0 0 6px;font-size:12px;line-height:1.4;color:#5A5A5A;"><b>Noted:</b> Apabila dalam waktu 3 jam setelah shipment tiba tidak terdapat feedback dari pihak penerima, maka seluruh tanggung jawab terkait paket dianggap telah diterima oleh next station.</p>
			<p style="margin:0;">Terima-kasih</p>
		</div>
	`;

	const text = [
		'Dear All',
		'',
		`Berikut Terlampir Surat Jalan From ${getHubAmhName()} To ${destination}`,
		'',
		`=== ${getHubShortName().toUpperCase()} ===`,
		...rows.map(([label, value]) => `${label}: ${value}`),
		'',
		'Noted: Apabila dalam waktu 3 jam setelah shipment tiba tidak terdapat feedback dari pihak penerima, maka seluruh tanggung jawab terkait paket dianggap telah diterima oleh next station.',
		'',
		'Terima-kasih'
	].join('\n');

	return {
		html,
		previewHtml,
		text,
		reportDate,
		slot
	};
}

function updatePrealertEmailPreview() {
	if (emailSubjectEl) {
		emailSubjectEl.textContent = buildPrealertSubject();
	}

	if (emailBodyEl) {
		emailBodyEl.innerHTML = buildPrealertBody().previewHtml;
	}
}

function sanitizePrealertRecipient(rawRecipient) {
	const value = normalize(rawRecipient);
	if (!value) {
		return '';
	}

	if (/^https?:\/\//i.test(value)) {
		return '';
	}

	const isEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
	return isEmail ? value : '';
}

function isValidEmailAddress(value) {
	return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}

function formatMailboxAddress(name, email) {
	const safeEmail = String(email || '').trim();
	if (!safeEmail) {
		return '';
	}

	const safeName = String(name || '').trim().replace(/"/g, '\\"');
	return safeName ? `"${safeName}" <${safeEmail}>` : safeEmail;
}

function buildPrealertRecipients() {
	const forcedRecipient = sanitizePrealertRecipient(PREALERT_DRAFT_RECIPIENT);
	const toSet = new Set(
		PREALERT_TO_RECIPIENTS
			.map((item) => String(item || '').trim().toLowerCase())
			.filter((email) => isValidEmailAddress(email))
	);
	if (forcedRecipient) {
		toSet.add(forcedRecipient.toLowerCase());
	}

	const toEmails = Array.from(toSet);
	const ccEntries = PREALERT_CC_RECIPIENTS
		.map((item) => ({
			name: String(item?.name || '').trim(),
			email: String(item?.email || '').trim().toLowerCase()
		}))
		.filter((item) => isValidEmailAddress(item.email));

	const toHeader = toEmails.length ? toEmails.join(', ') : 'undisclosed-recipients:;';
	const ccHeader = ccEntries.length
		? ccEntries.map((item) => formatMailboxAddress(item.name, item.email)).join(', ')
		: '';

	return {
		toEmails,
		ccEmails: ccEntries.map((item) => item.email),
		toHeader,
		ccHeader,
		toQuery: toEmails.join(','),
		ccQuery: ccEntries.map((item) => item.email).join(',')
	};
}

async function requestPrealertDraftViaApi(payload) {
	if (!PREALERT_DRAFT_API_URL) {
		throw new Error('API URL belum diisi di PREALERT_DRAFT_API_URL');
	}

	const apiUrl = String(PREALERT_DRAFT_API_URL || '').trim();
	if (/github\.io/i.test(apiUrl) || /doaltaspace\.github\.io/i.test(apiUrl)) {
		throw new Error('PREALERT_DRAFT_API_URL harus endpoint backend API, bukan URL frontend GitHub Pages.');
	}

	if (/script\.google(?:usercontent)?\.com/i.test(apiUrl)) {
		throw new Error('Mode ini non-Apps Script. Kosongkan PREALERT_DRAFT_API_URL atau pakai backend API Anda sendiri (bukan script.googleusercontent/script.google.com).');
	}

	const headers = {
		'Content-Type': 'application/json'
	};

	if (PREALERT_DRAFT_API_TOKEN) {
		headers.Authorization = `Bearer ${PREALERT_DRAFT_API_TOKEN}`;
	}

	const response = await fetch(PREALERT_DRAFT_API_URL, {
		method: 'POST',
		headers,
		body: JSON.stringify(payload)
	});

	const rawText = await response.text();
	let result = {};
	try {
		result = rawText ? JSON.parse(rawText) : {};
	} catch (error) {
		result = {
			success: response.ok,
			message: rawText || `HTTP ${response.status}`
		};
	}

	if (!response.ok || result.success === false) {
		throw new Error(result.message || `HTTP ${response.status}`);
	}

	return result;
}

function encodeUtf8ToBase64Url(value) {
	const bytes = new TextEncoder().encode(String(value || ''));
	let binary = '';
	bytes.forEach((byte) => {
		binary += String.fromCharCode(byte);
	});

	return btoa(binary)
		.replace(/\+/g, '-')
		.replace(/\//g, '_')
		.replace(/=+$/g, '');
}

function isRevokedOrExpiredTokenMessage(message) {
	return /expired or revoked|invalid_grant|invalid credentials|invalid[_ ]token/i.test(String(message || ''));
}

function buildDirectApiTokenHelpMessage(rawMessage) {
	if (!isRevokedOrExpiredTokenMessage(rawMessage)) {
		return rawMessage;
	}

	return 'Token Gmail API kadaluarsa/revoked. Generate refresh token baru (offline access + prompt consent), update PREALERT_GOOGLE_REFRESH_TOKEN, lalu coba lagi.';
}

async function requestPrealertDraftDirectGmail(payload) {
	if (!prealertRuntimeAccessToken) {
		prealertRuntimeAccessToken = loadCachedPrealertAccessToken();
	}

	if (!prealertRuntimeAccessToken && !PREALERT_GOOGLE_REFRESH_TOKEN) {
		throw new Error('Bearer token belum diisi');
	}

	const toHeader = payload.toHeader || payload.recipient || PREALERT_DRAFT_RECIPIENT || 'undisclosed-recipients:;';
	const ccHeader = payload.ccHeader || '';
	const mimeHeaders = [
		`To: ${toHeader}`,
		`Subject: ${payload.subject}`,
		'MIME-Version: 1.0',
		'Content-Type: text/html; charset="UTF-8"'
	];

	if (ccHeader) {
		mimeHeaders.splice(1, 0, `Cc: ${ccHeader}`);
	}

	const mime = [
		...mimeHeaders,
		'',
		payload.htmlBody || payload.textBody || ''
	].join('\r\n');

	const raw = encodeUtf8ToBase64Url(mime);

	const createDraft = async (accessToken) => fetch('https://gmail.googleapis.com/gmail/v1/users/me/drafts', {
		method: 'POST',
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${accessToken}`
		},
		body: JSON.stringify({ message: { raw } })
	});

	const refreshAccessToken = async () => {
		if (!PREALERT_GOOGLE_CLIENT_ID || !PREALERT_GOOGLE_CLIENT_SECRET || !PREALERT_GOOGLE_REFRESH_TOKEN) {
			throw new Error('Refresh token belum diisi (PREALERT_GOOGLE_REFRESH_TOKEN)');
		}

		const tokenResp = await fetch('https://oauth2.googleapis.com/token', {
			method: 'POST',
			headers: {
				'Content-Type': 'application/x-www-form-urlencoded'
			},
			body: new URLSearchParams({
				client_id: PREALERT_GOOGLE_CLIENT_ID,
				client_secret: PREALERT_GOOGLE_CLIENT_SECRET,
				refresh_token: PREALERT_GOOGLE_REFRESH_TOKEN,
				grant_type: 'refresh_token'
			}).toString()
		});

		const tokenText = await tokenResp.text();
		let tokenData = {};
		try {
			tokenData = tokenText ? JSON.parse(tokenText) : {};
		} catch (error) {
			tokenData = {};
		}

		if (!tokenResp.ok || !tokenData.access_token) {
			const tokenError = tokenData.error_description || tokenData.error || `Token refresh gagal (HTTP ${tokenResp.status})`;
			throw new Error(buildDirectApiTokenHelpMessage(tokenError));
		}

		prealertRuntimeAccessToken = tokenData.access_token;
		saveCachedPrealertAccessToken(prealertRuntimeAccessToken, tokenData.expires_in);
		return prealertRuntimeAccessToken;
	};

	if (!prealertRuntimeAccessToken && PREALERT_GOOGLE_REFRESH_TOKEN) {
		await refreshAccessToken();
	}

	let response = await createDraft(prealertRuntimeAccessToken);

	const text = await response.text();
	let result = {};
	try {
		result = text ? JSON.parse(text) : {};
	} catch (error) {
		result = {};
	}

	const needRefresh = response.status === 401
		|| response.status === 403
		|| /invalid[_ ]token|invalid credentials|expired/i.test(result.error?.message || '');
	if (needRefresh && PREALERT_GOOGLE_REFRESH_TOKEN) {
		clearCachedPrealertAccessToken();
		prealertRuntimeAccessToken = '';
		const newToken = await refreshAccessToken();
		response = await createDraft(newToken);
		const retryText = await response.text();
		try {
			result = retryText ? JSON.parse(retryText) : {};
		} catch (error) {
			result = {};
		}
	}

	if (!response.ok) {
		const message = result.error?.message || `HTTP ${response.status}`;
		throw new Error(buildDirectApiTokenHelpMessage(message));
	}

	return {
		success: true,
		message: 'Draft berhasil dibuat via Gmail API token.',
		draftId: result.id,
		draftUrl: result.id ? `https://mail.google.com/mail/u/0/#drafts?compose=${encodeURIComponent(result.id)}` : ''
	};
}

async function generateGmailDraft() {
	const subject = buildPrealertSubject();
	const body = buildPrealertBody();
	const recipients = buildPrealertRecipients();
	const autoDraftConfigured = Boolean(PREALERT_DRAFT_API_URL)
		|| Boolean(PREALERT_GOOGLE_REFRESH_TOKEN)
		|| prealertRuntimeAccessToken.startsWith('ya29.');
	const payload = {
		subject,
		htmlBody: body.html,
		textBody: body.text,
		source: getSelectedHubName(),
		recipient: recipients.toEmails[0] || undefined,
		to: recipients.toEmails,
		cc: recipients.ccEmails,
		toHeader: recipients.toHeader,
		ccHeader: recipients.ccHeader
	};
	const directModeEnabled = PREALERT_FORCE_DIRECT_GMAIL_API;

	const openManualDraftCompose = async () => {
		const query = new URLSearchParams({
			view: 'cm',
			fs: '1',
			su: subject
		});

		if (recipients.toQuery) {
			query.set('to', recipients.toQuery);
		}

		if (recipients.ccQuery) {
			query.set('cc', recipients.ccQuery);
		}

		const gmailUrl = `https://mail.google.com/mail/?${query.toString()}`;

		window.open(gmailUrl, '_blank', 'noopener');

		// Gmail compose URL only supports plain text body, so we copy HTML and ask user to paste.
		if (navigator.clipboard && window.ClipboardItem) {
			try {
				const item = new ClipboardItem({
					'text/html': new Blob([body.html], { type: 'text/html' }),
					'text/plain': new Blob([body.text], { type: 'text/plain' })
				});
				await navigator.clipboard.write([item]);
				setSuccess('Gmail draft dibuka. Template tabel sudah di-copy. Paste (Ctrl+V) di body email agar tampil tabel.');
				return;
			} catch (error) {
				// Fallback to plain text clipboard below.
			}
		}

		// Legacy clipboard fallback for browsers that block ClipboardItem rich HTML.
		try {
			const hidden = document.createElement('div');
			hidden.innerHTML = body.html;
			hidden.style.position = 'fixed';
			hidden.style.left = '-99999px';
			hidden.style.top = '0';
			document.body.appendChild(hidden);

			const range = document.createRange();
			range.selectNodeContents(hidden);
			const selection = window.getSelection();
			selection?.removeAllRanges();
			selection?.addRange(range);

			const copied = document.execCommand('copy');
			selection?.removeAllRanges();
			hidden.remove();

			if (copied) {
				setSuccess('Gmail draft dibuka. Template tabel berhasil di-copy. Paste (Ctrl+V) di body email.');
				return;
			}
		} catch (error) {
			// Continue to plain text fallback.
		}

		if (navigator.clipboard) {
			try {
				await navigator.clipboard.writeText(body.text);
				setWarning('Gmail draft dibuka. Browser hanya mengizinkan copy teks, jadi tabel tidak otomatis.');
				return;
			} catch (error) {
				setWarning('Gmail draft dibuka. Silakan copy manual dari Email Preview.');
			}
		}
	};

	if (PREALERT_DRAFT_API_URL && !directModeEnabled) {
		setButtonLoading(generateGmailDraftBtn, true, 'Creating Draft...');
		try {
			const result = await requestPrealertDraftViaApi(payload);
			if (result.draftUrl) {
				window.open(result.draftUrl, '_blank', 'noopener');
			} else {
				window.open('https://mail.google.com/mail/#drafts', '_blank', 'noopener');
			}

			setSuccess(result.message || `Draft Gmail otomatis berhasil dibuat${result.recipient ? ` untuk ${result.recipient}` : ''}.`);
			return;
		} catch (error) {
			if (prealertRuntimeAccessToken.startsWith('ya29.') || PREALERT_GOOGLE_REFRESH_TOKEN) {
				try {
					const fallbackDirectResult = await requestPrealertDraftDirectGmail(payload);
					if (fallbackDirectResult.draftUrl) {
						window.open(fallbackDirectResult.draftUrl, '_blank', 'noopener');
					} else {
						window.open('https://mail.google.com/mail/#drafts', '_blank', 'noopener');
					}
					setSuccess(fallbackDirectResult.message || 'Draft Gmail otomatis berhasil dibuat.');
					return;
				} catch (directError) {
					setWarning(`Auto draft via API gagal (${error.message}) dan direct Gmail API gagal (${directError.message}). Dialihkan ke compose manual.`);
					await openManualDraftCompose();
					return;
				}
			} else {
				setWarning(`Auto draft via API gagal (${error.message}). Dialihkan ke compose manual.`);
				await openManualDraftCompose();
				return;
			}
		} finally {
			setButtonLoading(generateGmailDraftBtn, false);
		}
	} else if (prealertRuntimeAccessToken.startsWith('ya29.') || PREALERT_GOOGLE_REFRESH_TOKEN) {
		setButtonLoading(generateGmailDraftBtn, true, 'Creating Draft...');
		try {
			const result = await requestPrealertDraftDirectGmail(payload);
			if (result.draftUrl) {
				window.open(result.draftUrl, '_blank', 'noopener');
			} else {
				window.open('https://mail.google.com/mail/#drafts', '_blank', 'noopener');
			}
			setSuccess(result.message || 'Draft Gmail otomatis berhasil dibuat.');
			return;
		} catch (error) {
			setWarning(`Direct Gmail API gagal (${error.message}). Dialihkan ke compose manual.`);
			await openManualDraftCompose();
			return;
		} finally {
			setButtonLoading(generateGmailDraftBtn, false);
		}
	}

	if (autoDraftConfigured) {
		setError('Auto draft aktif, tapi pembuatan draft belum berhasil. Compose manual dibatalkan supaya tidak perlu paste manual.');
		return;
	}

	await openManualDraftCompose();
}

function downloadPrealertReport() {
	const subject = buildPrealertSubject();
	const body = buildPrealertBody();
	const reportHtml = `<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>${escapeHtml(subject)}</title>
  <style>
    body { font-family: Arial, sans-serif; color: #111827; padding: 20px; }
    table { width: 100%; max-width: 520px; border-collapse: collapse; margin-top: 8px; }
    th, td { border: 1px solid #111827; padding: 6px 8px; font-size: 13px; }
    th { background: #8dc63f; text-align: center; }
    td:first-child { width: 46%; font-weight: 600; }
  </style>
</head>
<body>
  <h3>${escapeHtml(subject)}</h3>
  ${body.html}
</body>
</html>`;

	const blob = new Blob([reportHtml], { type: 'text/html;charset=utf-8' });
	const url = URL.createObjectURL(blob);
	const anchor = document.createElement('a');
	anchor.href = url;
	anchor.download = `Pre-Alert-${getIndonesianReportDate()}.html`;
	document.body.appendChild(anchor);
	anchor.click();
	anchor.remove();
	URL.revokeObjectURL(url);

	setSuccess('Report Pre-Alert berhasil diunduh.');
}

function sanitizeFileNamePart(value, fallback = '-') {
	const normalized = normalize(value) || fallback;
	const safe = normalized
		.replace(/[\\/:*?"<>|]/g, ' ')
		.replace(/\s+/g, ' ')
		.trim();

	return safe || fallback;
}

function buildPrealertFirstPageJpgFilename() {
	const slotNumber = getSlotTripNumber(slotSelect?.value);
	const plate = getTextValue(plateNumberEl);
	const reportDate = getSelectedReportDateFormatted() || formatDateDDMMYYYY(getTodayInputDate()) || '-';

	const safeSlot = sanitizeFileNamePart(slotNumber, '-');
	const safePlate = sanitizeFileNamePart(plate, '-');
	const safeDate = sanitizeFileNamePart(reportDate, '-');

	return `Surat Jalan Slot ${safeSlot} - ${safePlate} - ${getHubShortName()} to Transit Point Depok DC - ${safeDate}.jpg`;
}

async function downloadPrealertFirstPageJpg() {
	const selectedFile = prealertUploadedPdfFile || prealertPdfInput?.files?.[0];
	if (!selectedFile) {
		setWarning('Upload PDF terlebih dahulu sebelum download JPG Surat Jalan.');
		return;
	}

	if (!isPdfFile(selectedFile)) {
		setWarning('File yang dipilih bukan PDF. Silakan upload PDF terlebih dahulu.');
		return;
	}

	if (typeof pdfjsLib === 'undefined' || !pdfjsLib?.getDocument) {
		setError('PDF.js belum tersedia, tidak bisa membuat JPG dari PDF.');
		return;
	}

	setButtonLoading(downloadPrealertFirstPageJpgBtn, true, 'Rendering JPG...');
	try {
		if (pdfjsLib.GlobalWorkerOptions) {
			pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
		}

		const buffer = await selectedFile.arrayBuffer();
		const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;

		let targetPageNumber = null;
		const hubNameForPdf = getSelectedHubName();
		const hubPdfPattern = new RegExp(hubNameForPdf.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s*'), 'i');
		for (let i = 1; i <= pdf.numPages; i += 1) {
			const candidatePage = await pdf.getPage(i);
			const textContent = await candidatePage.getTextContent();
			const pageText = textContent.items.map((item) => item.str).join(' ');
			if (/Origin/i.test(pageText) && hubPdfPattern.test(pageText)) {
				targetPageNumber = i;
				break;
			}
		}

		// Fallback: try any page with "Origin" if hub-specific match failed
		if (!targetPageNumber) {
			for (let i = 1; i <= pdf.numPages; i += 1) {
				const candidatePage = await pdf.getPage(i);
				const textContent = await candidatePage.getTextContent();
				const pageText = textContent.items.map((item) => item.str).join(' ');
				if (/Origin/i.test(pageText) && /First\s*Mile\s*Hub/i.test(pageText)) {
					targetPageNumber = i;
					break;
				}
			}
		}

		if (!targetPageNumber) {
			throw new Error(`Tidak ditemukan halaman dengan "Origin : ${hubNameForPdf}" di PDF ini.`);
		}

		const page = await pdf.getPage(targetPageNumber);
		const viewport = page.getViewport({ scale: 2 });
		const canvas = document.createElement('canvas');
		const context = canvas.getContext('2d', { alpha: false });

		canvas.width = Math.max(1, Math.floor(viewport.width));
		canvas.height = Math.max(1, Math.floor(viewport.height));

		if (!context) {
			throw new Error('Canvas context tidak tersedia.');
		}

		await page.render({ canvasContext: context, viewport }).promise;

		const jpgBlob = await new Promise((resolve, reject) => {
			canvas.toBlob((blob) => {
				if (!blob) {
					reject(new Error('Gagal membuat file JPG.'));
					return;
				}
				resolve(blob);
			}, 'image/jpeg', 0.92);
		});

		const objectUrl = URL.createObjectURL(jpgBlob);
		const anchor = document.createElement('a');
		anchor.href = objectUrl;
		anchor.download = buildPrealertFirstPageJpgFilename();
		document.body.appendChild(anchor);
		anchor.click();
		anchor.remove();
		URL.revokeObjectURL(objectUrl);

		setSuccess(`JPG halaman ${targetPageNumber} (Origin: ${getSelectedHubName()}) berhasil diunduh.`);
	} catch (error) {
		setError(`Gagal membuat JPG Surat Jalan: ${error.message}`);
	} finally {
		setButtonLoading(downloadPrealertFirstPageJpgBtn, false);
	}
}

function readMatch(text, patterns) {
	for (const pattern of patterns) {
		const match = text.match(pattern);
		if (match?.[1]) {
			return normalize(match[1]) || '-';
		}
	}
	return '-';
}

function extractAfter(label, text, maxLength = 50) {
	const source = String(text || '');
	const index = source.indexOf(label);

	if (index === -1) {
		return '-';
	}

	const slice = source.substring(index + label.length, index + label.length + maxLength);
	const value = slice.split(/[\n\r]/)[0].trim();

	return value || '-';
}

function buildLinesFromTextItems(items) {
	const prepared = items
		.map((item) => ({
			str: normalize(item.str),
			x: Number(item.transform?.[4] || 0),
			y: Number(item.transform?.[5] || 0)
		}))
		.filter((entry) => entry.str);

	prepared.sort((a, b) => {
		if (Math.abs(a.y - b.y) > 1.5) {
			return b.y - a.y;
		}
		return a.x - b.x;
	});

	const groups = [];
	prepared.forEach((entry) => {
		const lastGroup = groups[groups.length - 1];
		if (!lastGroup || Math.abs(lastGroup.y - entry.y) > 2.8) {
			groups.push({ y: entry.y, items: [entry] });
			return;
		}
		lastGroup.items.push(entry);
	});

	return groups
		.map((group) => group.items
			.sort((a, b) => a.x - b.x)
			.map((item) => item.str)
			.join(' ')
			.replace(/\s+/g, ' ')
			.trim())
		.filter(Boolean);
}

function extractLineValue(lines, labelRegex, options = {}) {
	const maxLookAhead = options.maxLookAhead ?? 2;
	const skipRegex = options.skipRegex || /(Surat Jalan|Nama Line Haul Trip|Origin|ATA|STD|STA|Kode Segel|Schedule\/Adhoc|Nama Vendor|Notes|PIC Gudang|#\s+Nomor TO|Jmlh|Berat|TO Type|DG Type)/i;

	for (let index = 0; index < lines.length; index += 1) {
		const line = lines[index];
		if (!labelRegex.test(line)) {
			continue;
		}

		const inlineMatch = line.match(new RegExp(`${labelRegex.source}\\s*:?\\s*(.+)$`, 'i'));
		if (inlineMatch?.[1]) {
			const inlineValue = normalize(inlineMatch[1]);
			if (inlineValue && !skipRegex.test(inlineValue)) {
				return inlineValue;
			}
		}

		for (let offset = 1; offset <= maxLookAhead; offset += 1) {
			const nextLine = normalize(lines[index + offset]);
			if (!nextLine || skipRegex.test(nextLine)) {
				continue;
			}
			if (/^[:\-]+$/.test(nextLine)) {
				continue;
			}
			if (/^\d{1,4}[\/.-]\d{1,2}[\/.-]\d{1,4}$/.test(nextLine)) {
				continue;
			}
			return nextLine;
		}
	}

	return '-';
}

function extractQuantitiesFromLines(lines) {
	const blockText = lines.join('\n');
	const vendorPattern = /Nama\s*Vendor[\s\S]{0,160}?(\d+)\s+(\d+)\s+(\d+)\s+[\d.,]+\s+(\d+)/gi;
	let best = null;
	let match = vendorPattern.exec(blockText);

	while (match) {
		const totalTo = Number.parseInt(match[1], 10);
		const hv = Number.parseInt(match[2], 10);
		const paket = Number.parseInt(match[3], 10);
		if (!best || totalTo > best.totalTo) {
			best = { totalTo, hv, paket };
		}
		match = vendorPattern.exec(blockText);
	}

	if (best) {
		return {
			hvQty: String(best.hv),
			totalToQty: String(best.totalTo),
			orderQty: String(best.paket)
		};
	}

	const hvQty = extractLineValue(lines, /Jumlah\s*TO\s*HV/i, { maxLookAhead: 3 }).match(/\d+/)?.[0] || '-';
	const totalToQty = extractLineValue(lines, /Jumlah\s*TO(?!\s*HV)/i, { maxLookAhead: 3 }).match(/\d+/)?.[0] || '-';
	const orderQty = extractLineValue(lines, /Jumlah\s*Paket/i, { maxLookAhead: 3 }).match(/\d+/)?.[0] || '-';

	return { hvQty, totalToQty, orderQty };
}

function extractLineHaulDataFromLines(lines) {
	const textBlock = lines.join('\n');
	const compactText = textBlock.replace(/\s+/g, ' ').trim();

	const driverRegexMatch = compactText.match(/Nama\s*Driver\s*:?\s*([A-Za-z\s'.-]{3,}?)(?=\s+Nomor\s*Polisi|\s+Jumlah\s*TO|\s+Schedule\/Adhoc|$)/i);
	const plateRegexMatch = compactText.match(/Nomor\s*Polisi\s*:?\s*([A-Z]{1,2}\s*\d{1,4}\s*[A-Z]{0,3})/i);

	const rawDriver = driverRegexMatch?.[1] || extractLineValue(lines, /Nama\s*Driver/i, { maxLookAhead: 1 });
	const cleanedDriver = normalize(rawDriver)
		.replace(/Nomor\s*Polisi[\s\S]*$/i, '')
		.replace(/\s{2,}/g, ' ')
		.trim();

	const plateValue = plateRegexMatch?.[1] || extractLineValue(lines, /Nomor\s*Polisi|No\.?\s*Polisi/i, { maxLookAhead: 2 });
	const destination = extractLineValue(lines, /\bDestination\b/i, { maxLookAhead: 2 });
	const timeValue = extractLineValue(lines, /Waktu\s*Segel/i, { maxLookAhead: 2 });
	const staLineValue = extractLineValue(lines, /\bSTA\b|Jadwal\s*Kedatangan/i, {
		maxLookAhead: 3,
		skipRegex: /(Surat Jalan|Nama Line Haul Trip|Origin|Kode Segel|Schedule\/Adhoc|Nama Vendor|Notes|PIC Gudang|#\s+Nomor TO|Jmlh|Berat|TO Type|DG Type)/i
	});
	const staTimeFromLines = extractTimeByLabelFromLines(lines, /\bSTA\b|Jadwal\s*Kedatangan/i);
	const ataLineValue = extractLineValue(lines, /\bATA\b|Actual\s*Kedatangan/i, {
		maxLookAhead: 3,
		skipRegex: /(Surat Jalan|Nama Line Haul Trip|Origin|Kode Segel|Schedule\/Adhoc|Nama Vendor|Notes|PIC Gudang|#\s+Nomor TO|Jmlh|Berat|TO Type|DG Type)/i
	});
	const ataTimeFromLines = extractTimeByLabelFromLines(lines, /\bATA\b|Actual\s*Kedatangan/i);
	const stdLineValue = extractLineValue(lines, /\bSTD\b|Jadwal\s*Keberangkatan/i, {
		maxLookAhead: 3,
		skipRegex: /(Surat Jalan|Nama Line Haul Trip|Origin|Kode Segel|Schedule\/Adhoc|Nama Vendor|Notes|PIC Gudang|#\s+Nomor TO|Jmlh|Berat|TO Type|DG Type)/i
	});
	const stdTimeFromLines = extractTimeByLabelFromLines(lines, /\bSTD\b|Jadwal\s*Keberangkatan/i);
	const stfLineValue = extractLineValue(lines, /\bSTF\b/i, {
		maxLookAhead: 3,
		skipRegex: /(Surat Jalan|Nama Line Haul Trip|Origin|Kode Segel|Schedule\/Adhoc|Nama Vendor|Notes|PIC Gudang|#\s+Nomor TO|Jmlh|Berat|TO Type|DG Type)/i
	});
	const stfTimeFromLines = extractTimeByLabelFromLines(lines, /\bSTF\b/i);

	const staTime = pickPreferredClockValue([
		staTimeFromLines,
		staLineValue,
		...readAllMatches(compactText, [
			/(?:\bSTA\b|\bJadwal\s*Kedatangan\b)[\s\S]{0,80}?(\d{1,2}:\d{2}(?::\d{2})?)/i
		])
	]);
	const ataTime = pickPreferredClockValue([
		ataTimeFromLines,
		ataLineValue,
		...readAllMatches(compactText, [
			/(?:\bATA\b|\bActual\s*Kedatangan\b)[\s\S]{0,80}?(\d{1,2}:\d{2}(?::\d{2})?)/i
		])
	]);
	const stdTime = pickPreferredClockValue([
		stdTimeFromLines,
		stdLineValue,
		...readAllMatches(compactText, [
			/(?:\bSTD\b|\bJadwal\s*Keberangkatan\b)[\s\S]{0,80}?(\d{1,2}:\d{2}(?::\d{2})?)/i
		])
	]);
	const stfTime = pickPreferredClockValue([
		stfTimeFromLines,
		stfLineValue,
		...readAllMatches(compactText, [
			/\bSTF\b[\s\S]{0,80}?(\d{1,2}:\d{2}(?::\d{2})?)/i
		])
	]);
	const atdTime = pickPreferredClockValue([
		extractTimeByLabelFromLines(lines, /Waktu\s*Segel|\bATD\b/i),
		timeValue,
		...readAllMatches(compactText, [
			/\bWaktu\s*Segel\b[^\d]*(?:\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}\s*)?(\d{1,2}:\d{2}(?::\d{2})?)/i,
			/\bATD(?:\s*\([^)]*\))?\s*[:\-]?\s*(?:\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}\s*)?(\d{1,2}:\d{2}(?::\d{2})?)/i
		])
	]);
	const sealCode = normalize(pickLastNonDash(readAllMatches(compactText, [
		/\bKode\s*Segel\s*[:\-]?\s*([A-Z0-9\-/]{3,})/i,
		/\bNo\.?\s*Seal\s*[:\-]?\s*([A-Z0-9\-/]{3,})/i,
		/\bSeal\s*Code\s*[:\-]?\s*([A-Z0-9\-/]{3,})/i
	]))).toUpperCase() || '-';
	const ltMatch = textBlock.match(/LT0[A-Z0-9]+/i);
	const quantity = extractQuantitiesFromLines(lines);

	return {
		driver: cleanedDriver || '-',
		plate: plateValue.match(/[A-Z]{1,2}\s*\d{1,4}\s*[A-Z]{0,3}/i)?.[0] || plateValue || '-',
		destination: destination || '-',
		time: atdTime || '-',
		ltNumber: ltMatch ? ltMatch[0].toUpperCase() : '-',
		staTime: staTime || '-',
		ataTime: ataTime || '-',
		stdTime: stdTime || '-',
		stfTime: stfTime || '-',
		atdTime: atdTime || '-',
		sealCode: sealCode || '-',
		hvQty: quantity.hvQty || '-',
		totalToQty: quantity.totalToQty || '-',
		orderQty: quantity.orderQty || '-'
	};
}

function readAllMatches(text, patterns) {
	const results = [];
	for (const pattern of patterns) {
		const flags = pattern.flags.includes('g') ? pattern.flags : `${pattern.flags}g`;
		const regex = new RegExp(pattern.source, flags);
		let match = regex.exec(text);
		while (match) {
			if (match[1]) {
				results.push(normalize(match[1]));
			}
			match = regex.exec(text);
		}
	}
	return results.filter(Boolean);
}

function pickLastNonDash(values) {
	for (let index = values.length - 1; index >= 0; index -= 1) {
		const value = normalize(values[index]);
		if (value && value !== '-') {
			return value;
		}
	}
	return '-';
}

function pickMaxNumeric(values) {
	let maxValue = null;
	values.forEach((value) => {
		const numeric = Number.parseInt(String(value).replace(/[^0-9]/g, ''), 10);
		if (Number.isFinite(numeric)) {
			maxValue = maxValue === null ? numeric : Math.max(maxValue, numeric);
		}
	});
	return maxValue === null ? '-' : String(maxValue);
}

function extractDataFromText(text) {
	const compactText = String(text || '').replace(/\s+/g, ' ').trim();

	const driver = pickLastNonDash(readAllMatches(compactText, [
		/\bNama\s*Driver\s*[:\-]?\s*([A-Za-z.\s'-]{3,})/i,
		/\bDriver\s*[:\-]?\s*([A-Za-z.\s'-]{3,})/i
	]));

	const plate = pickLastNonDash(readAllMatches(compactText, [
		/\bNomor\s*Polisi\s*[:\-]?\s*([A-Z]{1,2}\s*\d{1,4}\s*[A-Z]{0,3})/i,
		/\bNo\.?\s*Polisi\s*[:\-]?\s*([A-Z0-9\s-]{4,})/i,
		/\bPlate\s*Number\s*[:\-]?\s*([A-Z0-9\s-]{4,})/i
	]));

	const destination = pickLastNonDash(readAllMatches(compactText, [
		/\bNext\s*Destination\s*[:\-]?\s*([A-Za-z0-9.\s\-/]+)/i,
		/\bDestination\s*[:\-]?\s*([A-Za-z0-9.\s\-/]+)/i
	]));

	const trip = pickLastNonDash(readAllMatches(compactText, [
		/\bTrip\s*(?:Number|No\.?|#)?\s*[:\-]?\s*([A-Z0-9\-/]+)/i
	]));

	const time = pickLastNonDash(readAllMatches(compactText, [
		/\bWaktu\s*Segel\b[^\d]*(?:\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4})?\s*(\d{1,2}:\d{2}(?::\d{2})?)/i,
		/\bTime\b\s*[:\-]?\s*(\d{1,2}:\d{2}(?::\d{2})?)/i
	]));

	const ltNumber = pickLastNonDash(readAllMatches(compactText, [
		/\b(LT0[A-Z0-9\-/]+)\b/i,
		/\b(LT[A-Z0-9\-/]{3,})\b/i
	]));

	let hvQty = pickMaxNumeric(readAllMatches(compactText, [
		/\bJumlah\s*TO\s*HV\s*[:\-]?\s*(\d+)\b/i,
		/\bTO\s*HV\s*[:\-]?\s*(\d+)\b/i
	]));

	let totalToQty = pickMaxNumeric(readAllMatches(compactText, [
		/\bJumlah\s*TO\b(?!\s*HV)\s*[:\-]?\s*(\d+)\b/i,
		/\bTotal\s*TO\s*[:\-]?\s*(\d+)\b/i
	]));

	let orderQty = pickMaxNumeric(readAllMatches(compactText, [
		/\bJumlah\s*Paket\s*[:\-]?\s*(\d+)\b/i,
		/\bTotal\s*(?:Order|Paket)\s*(?:Quantity)?\s*[:\-]?\s*(\d+)\b/i
	]));

	const liquidationQty = pickMaxNumeric(readAllMatches(compactText, [
		/\bLiquidation\s*(?:TO\s*Quantity)?\s*[:\-]?\s*(\d+)\b/i,
		/\bTO\s*Type\b[^\n\r]*?\bLiquidation\b[^\d]*(\d+)\b/i
	]));

	const vendorSummary = compactText.match(/Nama\s*Vendor[\s\S]{0,120}?([0-9]+)\s+([0-9]+)\s+([0-9]+)\s+[0-9.,]+\s+([0-9]+)/i);
	if (vendorSummary) {
		totalToQty = String(Number.parseInt(vendorSummary[1], 10));
		hvQty = String(Number.parseInt(vendorSummary[2], 10));
		orderQty = String(Number.parseInt(vendorSummary[3], 10));
	}

	return {
		driver,
		plate,
		destination,
		trip,
		time,
		ltNumber,
		hvQty,
		totalToQty,
		liquidationQty,
		orderQty
	};
}

function applyExtractedData(data) {
	if (!data) {
		return;
	}

	linehaulTemplateData = {
		staTime: extractClockTime(data.staTime),
		ataTime: extractClockTime(data.ataTime),
		stdTime: extractClockTime(data.stdTime),
		stfTime: extractClockTime(data.stfTime),
		atdTime: extractClockTime(data.atdTime || data.time),
		sealCode: normalize(data.sealCode || '-').toUpperCase() || '-'
	};

	if (driverNameEl) driverNameEl.textContent = data.driver;
	if (plateNumberEl) plateNumberEl.textContent = data.plate;
	if (destinationEl) destinationEl.textContent = data.destination;
	if (tripNumberEl) tripNumberEl.textContent = data.trip;
	if (tripTimeEl) tripTimeEl.textContent = data.time;
	if (ltNumberEl) ltNumberEl.textContent = data.ltNumber;
	if (hvQtyEl) hvQtyEl.textContent = data.hvQty;
	if (totalToQtyEl) totalToQtyEl.textContent = data.totalToQty;
	if (liquidationQtyEl) liquidationQtyEl.textContent = data.liquidationQty;
	if (orderQtyEl) orderQtyEl.textContent = data.orderQty;

	// Auto-set trip input from Surat Jalan LT Number
	const ltVal = normalize(data.ltNumber);
	if (ltVal && ltVal !== '-') {
		tripInput.value = ltVal;
		tripHelper.textContent = `LT Number dari Surat Jalan: ${ltVal}`;
		refreshPreview();
		updateReportSummary();
	}

	updatePrealertEmailPreview();
}

async function parsePDF(file, onProgress) {
	if (typeof pdfjsLib === 'undefined' || !pdfjsLib?.getDocument) {
		throw new Error('PDF.js belum tersedia.');
	}

	if (pdfjsLib.GlobalWorkerOptions) {
		pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
	}

	const buffer = await file.arrayBuffer();
	const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
	const pages = [];
	const debugLines = [];
	if (typeof onProgress === 'function') {
		onProgress(20, `Membaca ${pdf.numPages} halaman PDF...`);
	}

	for (let pageIndex = 1; pageIndex <= pdf.numPages; pageIndex += 1) {
		const page = await pdf.getPage(pageIndex);
		const textContent = await page.getTextContent();
		const lines = buildLinesFromTextItems(textContent.items);
		pages.push(lines);
		debugLines.push(`--- PAGE ${pageIndex} ---`);
		lines.slice(0, 30).forEach((line) => debugLines.push(line));

		if (typeof onProgress === 'function') {
			const currentProgress = 20 + (pageIndex / Math.max(1, pdf.numPages)) * 55;
			onProgress(currentProgress, `Parsing halaman ${pageIndex}/${pdf.numPages}...`);
		}
	}

 	const pdfText = pages.flat().join(' ');
	console.log('PDF LINES:', debugLines.join('\n'));

	console.log('PDF TEXT:', pdfText);
	let extracted = null;
	let bestScore = -1;
	const candidates = [];

	pages.forEach((lines) => {
		const candidate = extractLineHaulDataFromLines(lines);
		candidates.push(candidate);
		const scoreParts = [
			candidate.driver !== '-' ? 1 : 0,
			candidate.plate !== '-' ? 1 : 0,
			candidate.destination !== '-' ? 1 : 0,
			candidate.time !== '-' ? 1 : 0,
			candidate.staTime !== '-' ? 1 : 0,
			candidate.ataTime !== '-' ? 1 : 0,
			candidate.stdTime !== '-' ? 1 : 0,
			candidate.stfTime !== '-' ? 1 : 0,
			candidate.atdTime !== '-' ? 1 : 0,
			candidate.sealCode !== '-' ? 1 : 0,
			candidate.ltNumber !== '-' ? 1 : 0,
			Number.isFinite(Number.parseInt(candidate.totalToQty, 10)) ? 1 : 0,
			Number.isFinite(Number.parseInt(candidate.orderQty, 10)) ? 1 : 0,
			Number.isFinite(Number.parseInt(candidate.hvQty, 10)) ? 1 : 0
		];
		const score = scoreParts.reduce((sum, value) => sum + value, 0);
		if (score > bestScore) {
			bestScore = score;
			extracted = candidate;
		}
	});

	if (!extracted) {
		extracted = extractLineHaulDataFromLines(pages.flat());
		candidates.push(extracted);
	}

	const mergedExtraction = {
		driver: extracted.driver || '-',
		plate: extracted.plate || '-',
		destination: extracted.destination || '-',
		time: pickLastNonDash(candidates.map((item) => item.atdTime || item.time || '-')),
		staTime: pickLastNonDash(candidates.map((item) => item.staTime || '-')),
		ataTime: pickLastNonDash(candidates.map((item) => item.ataTime || '-')),
		stdTime: pickLastNonDash(candidates.map((item) => item.stdTime || '-')),
		stfTime: pickLastNonDash(candidates.map((item) => item.stfTime || '-')),
		atdTime: pickLastNonDash(candidates.map((item) => item.atdTime || item.time || '-')),
		sealCode: pickLastNonDash(candidates.map((item) => item.sealCode || '-')),
		ltNumber: pickLastNonDash(candidates.map((item) => item.ltNumber || '-')),
		hvQty: pickMaxNumeric(candidates.map((item) => item.hvQty || '-')),
		totalToQty: pickMaxNumeric(candidates.map((item) => item.totalToQty || '-')),
		orderQty: pickMaxNumeric(candidates.map((item) => item.orderQty || '-'))
	};

	if (mergedExtraction.driver === '-') {
		mergedExtraction.driver = pickLastNonDash(candidates.map((item) => item.driver || '-'));
	}
	if (mergedExtraction.plate === '-') {
		mergedExtraction.plate = pickLastNonDash(candidates.map((item) => item.plate || '-'));
	}
	if (mergedExtraction.destination === '-') {
		mergedExtraction.destination = pickLastNonDash(candidates.map((item) => item.destination || '-'));
	}

	if (typeof onProgress === 'function') {
		onProgress(90, 'Mengisi data extraction ke tabel...');
	}

	applyExtractedData({
		driver: mergedExtraction.driver,
		plate: mergedExtraction.plate,
		destination: mergedExtraction.destination,
		trip: getSlotTripNumber(slotSelect?.value),
		time: mergedExtraction.time,
		staTime: mergedExtraction.staTime,
		ataTime: mergedExtraction.ataTime,
		stdTime: mergedExtraction.stdTime,
		stfTime: mergedExtraction.stfTime,
		atdTime: mergedExtraction.atdTime,
		sealCode: mergedExtraction.sealCode,
		ltNumber: mergedExtraction.ltNumber,
		hvQty: mergedExtraction.hvQty,
		totalToQty: mergedExtraction.totalToQty,
		liquidationQty: '-',
		orderQty: mergedExtraction.orderQty
	});

	if (typeof onProgress === 'function') {
		onProgress(100, 'Selesai memproses PDF.');
	}
}

async function handlePrealertPdfUpload(file) {
	if (!file) {
		prealertUploadedPdfFile = null;
		setPrealertUploadInfo('Belum ada file PDF dipilih.', 'info');
		setPrealertUploadState('idle', 0, 'Ready', 'Pilih file untuk mulai parsing.');
		resetPrealertExtractionTable();
		return;
	}

	if (!isPdfFile(file)) {
		prealertUploadedPdfFile = null;
		if (prealertPdfInput) {
			prealertPdfInput.value = '';
		}
		setPrealertUploadInfo('File harus berformat PDF (.pdf).', 'error');
		setPrealertUploadState('error', 0, 'Invalid File', 'Gunakan file PDF (.pdf).');
		showWarningToast('File harus berformat PDF (.pdf).');
		resetPrealertExtractionTable();
		return;
	}

	setPrealertUploadState('uploading', 10, 'Uploading', `File terdeteksi: ${file.name}`);
	setPrealertUploadInfo(`Membaca PDF: ${file.name}`, 'info');
	prealertUploadedPdfFile = file;

	try {
		console.log('PDF uploaded:', file.name);
		setPrealertUploadState('parsing', 18, 'Parsing', `Memproses ${file.name}`);
		await parsePDF(file, (progress, message) => {
			setPrealertUploadState('parsing', progress, 'Parsing', message);
			setPrealertUploadInfo(`${message} (${Math.round(progress)}%)`, 'info');
		});
		setPrealertUploadState('success', 100, 'Done', `${file.name} selesai diproses.`);
		setPrealertUploadInfo(`File ter-upload: ${file.name}`, 'success');
		setSuccess('PDF berhasil diproses dan data berhasil diekstrak');
	} catch (error) {
		resetPrealertExtractionTable();
		setPrealertUploadState('error', 100, 'Failed', `Gagal memproses ${file.name}`);
		setPrealertUploadInfo(`Gagal memproses PDF: ${error.message}`, 'error');
		setError(`Gagal memproses PDF: ${error.message}`);
	}
}

function setMassUploadStatus(message, tone = 'info') {
	if (!massUploadStatus) {
		return;
	}

	massUploadStatus.textContent = message;
	massUploadStatus.className = 'peek-status';

	if (tone === 'processing') {
		massUploadStatus.classList.add('status-processing');
	} else if (tone === 'success') {
		massUploadStatus.classList.add('status-success');
	} else if (tone === 'warning') {
		massUploadStatus.classList.add('status-warning');
	}
}

function applyTheme(theme) {
	const isDark = theme === 'dark';
	document.body.classList.toggle('dark-mode', isDark);
	themeToggle.innerHTML = isDark ? '<svg class="icon-theme" width="20" height="20" viewBox="0 0 24 24"><circle cx="12" cy="12" r="5" fill="#FFD966" stroke="#2D2D2D" stroke-width="2.2"/><line x1="12" y1="1" x2="12" y2="4" stroke="#FFD966" stroke-width="2.6" stroke-linecap="round"/><line x1="12" y1="20" x2="12" y2="23" stroke="#FFD966" stroke-width="2.6" stroke-linecap="round"/><line x1="4.2" y1="4.2" x2="6.3" y2="6.3" stroke="#FFD966" stroke-width="2.4" stroke-linecap="round"/><line x1="17.7" y1="17.7" x2="19.8" y2="19.8" stroke="#FFD966" stroke-width="2.4" stroke-linecap="round"/><line x1="1" y1="12" x2="4" y2="12" stroke="#FFD966" stroke-width="2.6" stroke-linecap="round"/><line x1="20" y1="12" x2="23" y2="12" stroke="#FFD966" stroke-width="2.6" stroke-linecap="round"/><line x1="4.2" y1="19.8" x2="6.3" y2="17.7" stroke="#FFD966" stroke-width="2.4" stroke-linecap="round"/><line x1="17.7" y1="6.3" x2="19.8" y2="4.2" stroke="#FFD966" stroke-width="2.4" stroke-linecap="round"/></svg>' : '<svg class="icon-theme" width="20" height="20" viewBox="0 0 24 24"><path d="M21 12.79A9 9 0 1 1 11.21 3a7 7 0 0 0 9.79 9.79z" fill="#FFD966" stroke="#2D2D2D" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/><circle cx="9" cy="10" r="1" fill="#2D2D2D"/><circle cx="14" cy="9" r="0.7" fill="#2D2D2D"/></svg>';
}

function initTheme() {
	const saved = localStorage.getItem('mass-upload-theme') || 'light';
	applyTheme(saved);
}

function setActiveTool(tool) {
	activeTool = tool;
	tabMassUpload?.classList.toggle('active', tool === 'mass');
	tabTOReport?.classList.toggle('active', tool === 'to');

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
	reportSummary.textContent = `LT: ${trip} | Slot: ${slot} | Report Date: ${reportDate}`;
	syncPrealertTripFromReportSlot();
	updatePrealertEmailPreview();
}

function onHubChanged() {
	const hubName = getSelectedHubName();

	// Update header subtitle
	if (headerSubtitle) {
		headerSubtitle.textContent = `Kijang ${hubName.replace(/\s*Hub\s*$/i, '')}`;
	}

	// Update footer
	const footer = document.querySelector('.app-footer');
	if (footer) {
		footer.textContent = `Copyright by AMH ${hubName.replace(/\s*First\s*Mile\s*Hub\s*$/i, '').trim()} Hub`;
	}

	// Update email preview
	updatePrealertEmailPreview();
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

function extractQuantityColumnIndexFromDataset() {
	return getColumnIndexByNames(['Total Quantity', 'Quantity', 'Qty', 'Total Paket', 'Total Order Quantity', 'Total Qty']);
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

	const message = `Missing column in uploaded file: ${missingColumns.join(' / ')}. LT Number akan diambil dari Surat Jalan.`;
	setWarning(message);
	if (showToastMessage) {
		showWarningToast(message);
	}
}

function resetOperatorState() {
	selectedOperators = [];
	availableOperators = [];
	operatorColumnIndex = -1;
	quantityColumnIndex = -1;
	bulkyTotalQty = 0;
	operatorInput.value = '';
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
	quantityColumnIndex = extractQuantityColumnIndexFromDataset();

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
			? 'Kolom Trip tidak ditemukan. LT Number dari Surat Jalan digunakan.'
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


	selectedOperators.forEach((operator) => {
		const tag = document.createElement('span');
		tag.className = 'selected-tag';
		tag.textContent = operator;

		const removeBtn = document.createElement('button');
		removeBtn.type = 'button';
		removeBtn.className = 'tag-remove';
		removeBtn.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24"><circle cx="12" cy="12" r="11" fill="#FF4D4D" stroke="#2D2D2D" stroke-width="2"/><path d="M15 9L9 15M9 9l6 6" stroke="#2D2D2D" stroke-width="2.5" stroke-linecap="round"/></svg>';
		removeBtn.addEventListener('click', () => {
			selectedOperators = selectedOperators.filter((item) => item !== operator);
			renderSelectedOperatorTags();
			renderOperatorDropdown();
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
			? 'File TO Management belum diunggah. Unggah file untuk melanjutkan.'
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
	massUploadData = [];
	massUploadDraftData = [];
	closeMassUploadModal();
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
	} catch (error) {
		dataset = [];
		headerRow = [];
		massUploadData = [];
		massUploadDraftData = [];
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

function getMassFilteredRows() {
	const trip = normalize(tripInput.value);
	if (!trip) {
		return [];
	}
	const tripRows = filterDatasetByTrip(trip);
	if (!tripRows.length) {
		return [];
	}
	return filterDatasetByOperator(tripRows, selectedOperators);
}

function syncMassUploadData(rows) {
	const previousMap = new Map(
		massUploadData.map((item) => [item.tracking, item])
	);

	massUploadData = rows
		.map((row) => normalize(row[COL.SPX]))
		.filter(Boolean)
		.map((tracking) => {
			const previous = previousMap.get(tracking);
			return {
				tracking,
				weight: previous?.weight ?? '',
				length: previous?.length ?? '',
				width: previous?.width ?? '',
				height: previous?.height ?? ''
			};
		});
}

function renderMassUploadEditorTable() {
	massUploadEditorBody.innerHTML = '';
	if (!massUploadDraftData.length) {
		const emptyRow = document.createElement('div');
		emptyRow.className = 'editor-empty-row';
		emptyRow.textContent = 'Belum ada data.';
		massUploadEditorBody.appendChild(emptyRow);
		return;
	}

	massUploadDraftData.forEach((item, index) => {
		const cells = [
			`<div class="editor-cell"><span class="mass-editor-label">${escapeHtml(item.tracking)}</span></div>`,
			`<div class="editor-cell"><input type="text" inputmode="decimal" class="mass-editor-input" data-index="${index}" data-col-index="0" data-field="weight" value="${escapeHtml(item.weight)}" /></div>`,
			`<div class="editor-cell"><input type="text" inputmode="decimal" class="mass-editor-input" data-index="${index}" data-col-index="1" data-field="length" value="${escapeHtml(item.length)}" /></div>`,
			`<div class="editor-cell"><input type="text" inputmode="decimal" class="mass-editor-input" data-index="${index}" data-col-index="2" data-field="width" value="${escapeHtml(item.width)}" /></div>`,
			`<div class="editor-cell"><input type="text" inputmode="decimal" class="mass-editor-input" data-index="${index}" data-col-index="3" data-field="height" value="${escapeHtml(item.height)}" /></div>`
		];

		const rowFragment = document.createRange().createContextualFragment(cells.join(''));
		massUploadEditorBody.appendChild(rowFragment);
	});
}

function openMassUploadEditor() {
	const rows = getMassFilteredRows();
	if (!rows.length) {
		setWarning('Tidak ada data filtered untuk diedit.');
		showWarningToast('Tidak ada data filtered untuk diedit.');
		return;
	}

	syncMassUploadData(rows);
	massUploadDraftData = massUploadData.map((item) => ({ ...item }));
	renderMassUploadEditorTable();
	openMassUploadModal();
}

function focusMassEditorCell(rowIndex, colIndex) {
	const cell = massUploadEditorBody.querySelector(
		`.mass-editor-input[data-index="${rowIndex}"][data-col-index="${colIndex}"]`
	);
	if (cell) {
		cell.focus();
		cell.select();
	}
}

function setActiveMassEditorCell(inputElement) {
	massUploadEditorBody.querySelectorAll('.mass-editor-input.is-active').forEach((cell) => {
		cell.classList.remove('is-active');
	});
	inputElement.classList.add('is-active');
}

function handleMassEditorKeydown(event) {
	const target = event.target;
	if (!(target instanceof HTMLInputElement) || !target.classList.contains('mass-editor-input')) {
		return;
	}

	const rowIndex = Number.parseInt(target.dataset.index || '-1', 10);
	const colIndex = Number.parseInt(target.dataset.colIndex || '-1', 10);
	if (!Number.isInteger(rowIndex) || !Number.isInteger(colIndex)) {
		return;
	}

	if (event.key === 'Enter') {
		event.preventDefault();
		const nextRow = Math.min(rowIndex + 1, massUploadDraftData.length - 1);
		focusMassEditorCell(nextRow, colIndex);
		return;
	}

	if (event.key === 'Tab') {
		event.preventDefault();
		const direction = event.shiftKey ? -1 : 1;
		let nextRow = rowIndex;
		let nextCol = colIndex + direction;

		if (nextCol > 3) {
			nextCol = 0;
			nextRow = Math.min(rowIndex + 1, massUploadDraftData.length - 1);
		}

		if (nextCol < 0) {
			nextCol = 3;
			nextRow = Math.max(rowIndex - 1, 0);
		}

		focusMassEditorCell(nextRow, nextCol);
	}
}

function normalizePastedDimensionValue(rawValue) {
	const cleaned = String(rawValue ?? '').trim();
	if (cleaned === '') {
		return null;
	}

	const numericValue = Number(cleaned);
	if (Number.isFinite(numericValue)) {
		return String(numericValue);
	}

	return cleaned;
}

function handleMassEditorPaste(event) {
	const target = event.target;
	if (!(target instanceof HTMLInputElement) || !target.classList.contains('mass-editor-input')) {
		return;
	}

	const startRow = Number.parseInt(target.dataset.index || '-1', 10);
	const startCol = Number.parseInt(target.dataset.colIndex || '-1', 10);
	if (!Number.isInteger(startRow) || startRow < 0 || startRow >= massUploadDraftData.length) {
		return;
	}
	if (!Number.isInteger(startCol) || startCol < 0 || startCol > 3) {
		return;
	}

	const clipboardText = event.clipboardData?.getData('text/plain') || '';
	if (!clipboardText.trim()) {
		return;
	}

	event.preventDefault();

	const parsedRows = clipboardText
		.replace(/\r/g, '')
		.split('\n')
		.filter((line) => line.length > 0)
		.map((line) => line.split('\t'));

	if (!parsedRows.length) {
		return;
	}

	const dimensionFields = ['weight', 'length', 'width', 'height'];
	let pastedRowCount = 0;

	for (let rowOffset = 0; rowOffset < parsedRows.length; rowOffset += 1) {
		const targetRowIndex = startRow + rowOffset;
		if (targetRowIndex >= massUploadDraftData.length) {
			break;
		}

		const sourceCells = parsedRows[rowOffset];
		let rowUpdated = false;

		for (let sourceColIndex = 0; sourceColIndex < sourceCells.length; sourceColIndex += 1) {
			const targetColIndex = startCol + sourceColIndex;
			if (targetColIndex > 3) {
				break;
			}

			const normalizedValue = normalizePastedDimensionValue(sourceCells[sourceColIndex]);
			if (normalizedValue === null) {
				continue;
			}

			const field = dimensionFields[targetColIndex];
			massUploadDraftData[targetRowIndex][field] = normalizedValue;
			rowUpdated = true;
		}

		if (rowUpdated) {
			pastedRowCount += 1;
		}
	}

	if (!pastedRowCount) {
		showWarningToast('Tidak ada data yang berhasil ditempel.');
		return;
	}

	renderMassUploadEditorTable();
	focusMassEditorCell(startRow, startCol);
	showToast(`${pastedRowCount} rows pasted successfully.`, 'success');
}

function shuffleDimensions(values) {
	const shuffled = [...values];
	for (let index = shuffled.length - 1; index > 0; index -= 1) {
		const randomIndex = Math.floor(Math.random() * (index + 1));
		const temp = shuffled[index];
		shuffled[index] = shuffled[randomIndex];
		shuffled[randomIndex] = temp;
	}
	return shuffled;
}

function getDimensionRepairReason(item) {
	const weightNormalized = normalize(item.weight).toUpperCase();
	if (weightNormalized === 'FIX_ME') {
		return 'fix_me';
	}
	if (
		normalize(item.weight) === '' ||
		normalize(item.length) === '' ||
		normalize(item.width) === '' ||
		normalize(item.height) === ''
	) {
		return 'empty';
	}
	return null;
}

function fixMassUploadDimensions() {
	if (!massUploadDraftData.length) {
		showWarningToast('Tidak ada data untuk diperbaiki.');
		return { total: 0, fixMeCount: 0, emptyCount: 0 };
	}

	let fixMeCount = 0;
	let emptyCount = 0;

	massUploadDraftData = massUploadDraftData.map((item) => {
		const reason = getDimensionRepairReason(item);
		if (!reason) {
			return item;
		}

		if (reason === 'fix_me') {
			fixMeCount += 1;
		} else {
			emptyCount += 1;
		}

		const [lengthValue, widthValue, heightValue] = shuffleDimensions([20, 25, 25]);

		return {
			...item,
			weight: '2.08',
			length: String(lengthValue),
			width: String(widthValue),
			height: String(heightValue)
		};
	});

	const total = fixMeCount + emptyCount;

	if (!total) {
		showWarningToast('Tidak ada baris FIX_ME atau kosong yang perlu diperbaiki.');
		return { total: 0, fixMeCount: 0, emptyCount: 0 };
	}

	const details = [];
	if (fixMeCount > 0) {
		details.push(`${fixMeCount} FIX_ME`);
	}
	if (emptyCount > 0) {
		details.push(`${emptyCount} kosong`);
	}

	renderMassUploadEditorTable();
	showToast({
		type: 'warning',
		title: 'Dimensions repaired',
		message: `${total} rows fixed (${details.join(', ')})`
	});

	return { total, fixMeCount, emptyCount };
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
	if (toStatusSummaryBody) {
		toStatusSummaryBody.innerHTML = '';
	}
	const trip = normalize(tripInput.value);
	const byTrip = trip ? filterDatasetByTrip(trip) : [];
	const filteredRows = filterDatasetByOperator(byTrip, selectedOperators);

	if (toStatusSummary) {
		toStatusSummary.classList.remove('hidden');
	}



	syncMassUploadData(filteredRows);

	if (viewAllPreviewBtn) {
		viewAllPreviewBtn.disabled = !filteredRows.length;
	}

	animateStatValue(totalDataEl, byTrip.length);
	animateStatValue(filteredDataEl, filteredRows.length);

	// Calculate bulky total quantity
	if (quantityColumnIndex !== -1) {
		bulkyTotalQty = filteredRows.reduce((sum, row) => {
			const val = Number.parseFloat(String(row[quantityColumnIndex] || '0').replace(/,/g, ''));
			return sum + (Number.isNaN(val) ? 0 : val);
		}, 0);
	} else {
		bulkyTotalQty = filteredRows.length; // fallback to count if no quantity column found
	}

	const statusMap = new Map();
	byTrip.forEach((rowData) => {
		const status = normalize(rowData[COL.TO_STATUS]) || 'Unknown';
		statusMap.set(status, (statusMap.get(status) || 0) + 1);
	});

	if (!statusMap.size) {
		const emptyStatusRow = document.createElement('tr');
		emptyStatusRow.innerHTML = '<td colspan="2">Belum ada data.</td>';
		toStatusSummaryBody?.appendChild(emptyStatusRow);
	} else {
		Array.from(statusMap.entries())
			.sort((a, b) => b[1] - a[1])
			.forEach(([status, count]) => {
				const row = document.createElement('tr');
				row.innerHTML = `<td>${escapeHtml(status)}</td><td>${count}</td>`;
				toStatusSummaryBody?.appendChild(row);
			});
	}


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
	syncMassUploadData(rows);
	const massDataMap = new Map(massUploadData.map((item) => [item.tracking, item]));
	const data = rows.map((row) => {
		const tracking = normalize(row[COL.SPX]);
		const edited = massDataMap.get(tracking);
		return [
			tracking,
			edited?.weight ?? '',
			edited?.length ?? '',
			edited?.width ?? '',
			edited?.height ?? ''
		];
	});
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
		setError('Upload Surat Jalan terlebih dahulu untuk mendapatkan LT Number.');
		return null;
	}

	const reportDate = getSelectedReportDateFormatted();
	if (!reportDate) {
		setWarning('Please select a report date first.');
		showWarningToast('Please select a report date first.');
		return null;
	}

	if (tripColumnIndex === -1) {
		const message = 'Kolom Line Hual Trip Number tidak ditemukan. Pastikan file TO Management sudah benar.';
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
		const message = 'File TO Management belum diunggah. Unggah file untuk melanjutkan.';
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
	setButtonLoading(generateAllInlineBtn, true, 'Generating Reports...');
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
		setButtonLoading(generateAllInlineBtn, false);
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

		if (event.ctrlKey && event.shiftKey && event.key.toLowerCase() === 't') {
			event.preventDefault();
			setActiveTool('to');

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
	updateDateDisplay();
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
	document.addEventListener('pointerdown', unlockNotificationAudio, { once: true });
	renderNotificationLog();
	setPrealertUploadInfo('Belum ada file PDF dipilih.', 'info');
	setPrealertUploadState('idle', 0, 'Ready', 'Pilih file untuk mulai parsing.');
	syncPrealertTripFromReportSlot();
	updatePrealertEmailPreview();
	const setFloatingMenuOpen = (isOpen) => {
		if (!floatingMenu || !floatingMenuToggle || !floatingMenuDropdown) {
			return;
		}

		floatingMenu.classList.toggle('open', isOpen);
		floatingMenuToggle.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
		floatingMenuDropdown.setAttribute('aria-hidden', isOpen ? 'false' : 'true');
	};

	if (floatingMenuToggle) {
		floatingMenuToggle.addEventListener('click', () => {
			const isOpen = floatingMenu?.classList.contains('open');
			setFloatingMenuOpen(!isOpen);
		});
	}

	document.querySelectorAll('.tab').forEach((tab) => {
		tab.addEventListener('click', () => {
			document.querySelectorAll('.tab').forEach((t) => t.classList.remove('active'));
			tab.classList.add('active');

			const target = tab.dataset.tab;

			document.querySelectorAll('.page').forEach((p) => p.classList.remove('active'));

			if (target === 'report') {
				document.getElementById('reportPage')?.classList.add('active');
			}

			if (target === 'prealert') {
				document.getElementById('prealertPage')?.classList.add('active');
			}

			if (target === 'hourly') {
				document.getElementById('hourlyPage')?.classList.add('active');
			}
		});
	});

	if (prealertPdfInput) {
		prealertPdfInput.addEventListener('change', async (event) => {
			const file = event.target.files?.[0];
			if (!file) {
				return;
			}
			await handlePrealertPdfUpload(file);
		});
	}

	if (prealertUploadBox) {
		prealertUploadBox.addEventListener('dragover', (event) => {
			event.preventDefault();
			prealertUploadBox.classList.add('dragover');
		});

		prealertUploadBox.addEventListener('dragleave', () => {
			prealertUploadBox.classList.remove('dragover');
		});

		prealertUploadBox.addEventListener('drop', async (event) => {
			event.preventDefault();
			prealertUploadBox.classList.remove('dragover');
			const file = event.dataTransfer?.files?.[0];
			if (!file) {
				return;
			}

			if (prealertPdfInput && typeof DataTransfer !== 'undefined') {
				const dataTransfer = new DataTransfer();
				dataTransfer.items.add(file);
				prealertPdfInput.files = dataTransfer.files;
			}

			await handlePrealertPdfUpload(file);
		});
	}

	generateGmailDraftBtn?.addEventListener('click', () => {
		generateGmailDraft();
	});

	downloadPrealertReportBtn?.addEventListener('click', () => {
		downloadPrealertReport();
	});

	downloadPrealertFirstPageJpgBtn?.addEventListener('click', () => {
		downloadPrealertFirstPageJpg();
	});

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

	operatorCombobox.addEventListener('click', (event) => {
		if (event.target === operatorCombobox || event.target === selectedOperatorTags) {
			operatorInput.focus();
		}
	});

	document.addEventListener('click', (event) => {
		if (!operatorCombobox.contains(event.target)) {
			closeOperatorDropdown();
		}
		if (!tripCombobox.contains(event.target)) {
			closeTripDropdown();
		}
		if (floatingMenu && !floatingMenu.contains(event.target)) {
			setFloatingMenuOpen(false);
		}
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

	viewAllPreviewBtn.addEventListener('click', openMassUploadEditor);

	massUploadEditorBody.addEventListener('input', (event) => {
		const target = event.target;
		if (!(target instanceof HTMLInputElement) || !target.classList.contains('mass-editor-input')) {
			return;
		}

		const index = Number.parseInt(target.dataset.index || '-1', 10);
		const field = target.dataset.field;
		if (!Number.isInteger(index) || index < 0 || index >= massUploadDraftData.length) {
			return;
		}
		if (!['weight', 'length', 'width', 'height'].includes(field)) {
			return;
		}

		massUploadDraftData[index][field] = normalize(target.value);
		setMassUploadStatus('Unsaved changes. Click Apply Changes to save.', 'warning');
	});

	massUploadEditorBody.addEventListener('focusin', (event) => {
		const target = event.target;
		if (target instanceof HTMLInputElement && target.classList.contains('mass-editor-input')) {
			setActiveMassEditorCell(target);
		}
	});

	massUploadEditorBody.addEventListener('focusout', (event) => {
		const target = event.target;
		if (target instanceof HTMLInputElement && target.classList.contains('mass-editor-input')) {
			target.classList.remove('is-active');
		}
	});

	massUploadEditorBody.addEventListener('keydown', handleMassEditorKeydown);

	massUploadEditorBody.addEventListener('paste', handleMassEditorPaste);

	fixMassUploadDimensionsBtn.addEventListener('click', () => {
		setMassUploadStatus('Fixing dimensions...', 'processing');
		const result = fixMassUploadDimensions();
		if (result.total > 0) {
			const parts = [];
			if (result.fixMeCount > 0) {
				parts.push(`${result.fixMeCount} FIX_ME`);
			}
			if (result.emptyCount > 0) {
				parts.push(`${result.emptyCount} kosong`);
			}
			setMassUploadStatus(`${result.total} rows fixed (${parts.join(', ')}). Review and apply changes.`, 'success');
			return;
		}
		setMassUploadStatus('No FIX_ME or empty rows detected to fix.', 'warning');
	});

	applyMassUploadChangesBtn.addEventListener('click', () => {
		setMassUploadStatus('Applying changes...', 'processing');
		setButtonLoading(applyMassUploadChangesBtn, true, 'Applying...');
		massUploadData = massUploadDraftData.map((item) => ({ ...item }));
		showToast({
			type: 'success',
			title: 'Success',
			message: 'Changes applied successfully.'
		});
		setMassUploadStatus('Changes applied successfully.', 'success');
		setTimeout(() => {
			setButtonLoading(applyMassUploadChangesBtn, false);
			closeMassUploadModal();
		}, 1000);
	});

	cancelMassUploadChangesBtn?.addEventListener('click', closeMassUploadModal);
	closeMassUploadModalBtn.addEventListener('click', closeMassUploadModal);
	massUploadConfirmCancelBtn.addEventListener('click', closeMassUploadCloseConfirmation);
	massUploadConfirmYesBtn.addEventListener('click', closeMassUploadModal);

	massUploadModal.addEventListener('click', (event) => {
		if (event.target === massUploadModal) {
			closeMassUploadModal();
		}
	});

	massUploadConfirmOverlay.addEventListener('click', (event) => {
		if (event.target === massUploadConfirmOverlay) {
			closeMassUploadCloseConfirmation();
		}
	});

	notificationLogOverlay.addEventListener('click', (event) => {
		if (event.target === notificationLogOverlay) {
			closeNotificationLog();
		}
	});

	document.addEventListener('keydown', (event) => {
		if (event.key === 'Escape' && floatingMenu?.classList.contains('open')) {
			setFloatingMenuOpen(false);
			return;
		}

		if (event.key === 'Escape' && notificationLogOverlay.classList.contains('open')) {
			closeNotificationLog();
			return;
		}

		if (event.key === 'Escape' && massUploadConfirmOverlay.classList.contains('open')) {
			closeMassUploadCloseConfirmation();
			return;
		}

		if (event.key === 'Escape' && massUploadModal.classList.contains('open')) {
			closeMassUploadModal();
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
	reportDateInput.addEventListener('change', () => {
		updateReportSummary();
		updateDateDisplay();
	});

	// Click anywhere on date box to open date picker
	const dateWrapper = document.getElementById('dateDisplayWrapper');
	if (dateWrapper && reportDateInput) {
		dateWrapper.addEventListener('click', (e) => {
			if (e.target === reportDateInput) return; // already handled natively
			try {
				reportDateInput.showPicker();
			} catch {
				reportDateInput.focus();
				reportDateInput.click();
			}
		});
	}

	// Hub select handler
	if (hubSelect) {
		hubSelect.addEventListener('change', () => {
			if (hubSelect.value === '__custom__') {
				hubSelectWrapper?.classList.add('hub-custom-active');
				if (hubCustomInput) {
					hubCustomInput.style.display = '';
					hubCustomInput.focus();
				}
			} else {
				hubSelectWrapper?.classList.remove('hub-custom-active');
				if (hubCustomInput) {
					hubCustomInput.style.display = 'none';
					hubCustomInput.value = '';
				}
			}
			onHubChanged();
		});
	}

	if (hubCustomInput) {
		hubCustomInput.addEventListener('input', () => {
			onHubChanged();
		});

		// Allow pressing Escape to go back to dropdown
		hubCustomInput.addEventListener('keydown', (e) => {
			if (e.key === 'Escape') {
				hubSelect.value = 'Ciputat 4 First Mile Hub';
				hubSelectWrapper?.classList.remove('hub-custom-active');
				hubCustomInput.style.display = 'none';
				hubCustomInput.value = '';
				onHubChanged();
			}
		});
	}

	copyLinehaulReportBtn?.addEventListener('click', () => {
		copyReportTemplateText('linehaul');
	});

	copyBulkyReportBtn?.addEventListener('click', () => {
		copyReportTemplateText('bulky');
	});

	copyHourlyReportBtn?.addEventListener('click', () => {
		copyReportTemplateText('hourly');
	});

	copyPickupReportTemplateBtn?.addEventListener('click', () => {
		copyReportTemplateText('pickup');
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



	generateAllBtn.addEventListener('click', generateAllReports);
	generateAllInlineBtn?.addEventListener('click', generateAllReports);
	generateEmailFab?.addEventListener('click', () => {
		setFloatingMenuOpen(false);
		generateGmailDraft();
	});
	generateAllBtn.addEventListener('click', () => {
		setFloatingMenuOpen(false);
	});
	notificationLogBtn.addEventListener('click', openNotificationLog);
	notificationLogCloseBtn.addEventListener('click', closeNotificationLog);
	notificationLogClearBtn.addEventListener('click', clearNotificationLog);

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

// ========================================
// DATABASE UPLOAD SECTION - Logic
// ========================================

const dbToFileInput = document.getElementById('dbToFileInput');
const dbToDropzone = document.getElementById('dbToDropzone');
const dbToChooseBtn = document.getElementById('dbToChooseBtn');
const dbToAddMoreBtn = document.getElementById('dbToAddMoreBtn');
const dbToLoadedFiles = document.getElementById('dbToLoadedFiles');
const dbToFileCount = document.getElementById('dbToFileCount');

const dbSjFileInput = document.getElementById('dbSjFileInput');
const dbSjDropzone = document.getElementById('dbSjDropzone');
const dbSjChooseBtn = document.getElementById('dbSjChooseBtn');
const dbSjReplaceBtn = document.getElementById('dbSjReplaceBtn');
const dbSjClearBtn = document.getElementById('dbSjClearBtn');
const dbSjLoadedFiles = document.getElementById('dbSjLoadedFiles');
const dbSjFileCount = document.getElementById('dbSjFileCount');

function dbFormatFileSize(bytes) {
	if (!bytes) return '0 B';
	const units = ['B', 'KB', 'MB', 'GB'];
	let size = bytes;
	let unit = 0;
	while (size >= 1024 && unit < units.length - 1) { size /= 1024; unit++; }
	return `${unit === 0 ? Math.round(size) : size.toFixed(1)} ${units[unit]}`;
}

function dbFileChipIconSvg(fileName) {
	const ext = String(fileName || '').split('.').pop().toLowerCase();
	if (ext === 'pdf') {
		return '<svg viewBox="0 0 24 24"><path d="M14 2H7a3 3 0 0 0-3 3v14a3 3 0 0 0 3 3h10a3 3 0 0 0 3-3V8z" fill="#FFCDD2" stroke="#C62828" stroke-width="2" stroke-linejoin="round"/><path d="M14 2v6a2 2 0 0 0 2 2h4" fill="#EF9A9A" stroke="#C62828" stroke-width="1.8" stroke-linejoin="round"/></svg>';
	}
	return '<svg viewBox="0 0 24 24"><rect x="3" y="2" width="18" height="20" rx="3.5" fill="#E8F5E9" stroke="#2D2D2D" stroke-width="2"/><line x1="7" y1="8" x2="17" y2="8" stroke="#66BB6A" stroke-width="2" stroke-linecap="round"/><line x1="7" y1="12" x2="14" y2="12" stroke="#43A047" stroke-width="2" stroke-linecap="round"/></svg>';
}

function renderDbToFiles() {
	if (!dbToLoadedFiles || !dbToFileCount || !dbToDropzone) return;

	dbToFileCount.textContent = `${dbToFiles.length} file`;
	dbToDropzone.setAttribute('data-state', dbToFiles.length > 0 ? 'loaded' : 'empty');

	dbToLoadedFiles.innerHTML = dbToFiles.map((f, i) => `
		<div class="db-file-chip">
			<span class="db-file-chip-icon">${dbFileChipIconSvg(f.name)}</span>
			<span class="db-file-chip-name" title="${escapeHtml(f.name)}">${escapeHtml(f.name)}</span>
			<span class="db-file-chip-size">${dbFormatFileSize(f.size)}</span>
			<button type="button" class="db-file-chip-remove" data-db-to-remove="${i}" aria-label="Remove ${escapeHtml(f.name)}">&times;</button>
		</div>
	`).join('');

	// Update stats
	updateDbToStats();
}

function updateDbToStats() {
	const statTO = document.getElementById('dbToStatTO');
	const statOperators = document.getElementById('dbToStatOperators');
	const statLT = document.getElementById('dbToStatLT');
	const statQty = document.getElementById('dbToStatQty');

	if (statOperators) {
		statOperators.textContent = availableOperators.length > 0 ? availableOperators.length.toLocaleString('id-ID') : '-';
	}
	
	if (statTO || statLT || statQty) {
		if (dataset.length > 0 && typeof headerRow !== 'undefined') {
			const uniqueTOs = new Set();
			const uniqueLTs = new Set();
			const uniqueTrackings = new Set();
			
			const toIndex = getColumnIndexByNames(['TO Number', 'TO_Number', 'Nomor TO', 'TO number', 'to number', 'TO No', 'TO_No']);
			const ltIndex = getColumnIndexByNames(['Line Hual Trip Number', 'LineHaul Trip Number', 'LT Number']);
			const trackingIndex = getColumnIndexByNames(['SPX Tracking Number', 'Tracking Number', 'Resi']);
			
			for (const row of dataset) {
				const to = toIndex !== -1 ? String(row[toIndex] || '').trim() : '';
				if (to) uniqueTOs.add(to);
				
				const lt = ltIndex !== -1 ? String(row[ltIndex] || '').trim() : '';
				if (lt) uniqueLTs.add(lt);
				
				const tracking = trackingIndex !== -1 ? String(row[trackingIndex] || '').trim() : '';
				if (tracking) uniqueTrackings.add(tracking);
			}
			
			if (statTO) statTO.textContent = uniqueTOs.size > 0 ? uniqueTOs.size.toLocaleString('id-ID') : '-';
			if (statLT) statLT.textContent = uniqueLTs.size > 0 ? uniqueLTs.size.toLocaleString('id-ID') : '-';
			if (statQty) statQty.textContent = uniqueTrackings.size > 0 ? uniqueTrackings.size.toLocaleString('id-ID') : '-';
		} else {
			if (statTO) statTO.textContent = '-';
			if (statLT) statLT.textContent = '-';
			if (statQty) statQty.textContent = '-';
		}
	}
}

function renderDbSjFile() {
	if (!dbSjLoadedFiles || !dbSjFileCount || !dbSjDropzone) return;

	dbSjFileCount.textContent = dbSjFile ? '1 file' : '0 file';
	dbSjDropzone.setAttribute('data-state', dbSjFile ? 'loaded' : 'empty');

	if (!dbSjFile) {
		dbSjLoadedFiles.innerHTML = '';
		updateDbSjExtracted();
		return;
	}

	dbSjLoadedFiles.innerHTML = '';
}

function updateDbSjExtracted() {
	const container = document.getElementById('dbSjExtracted');
	if (!container) return;

	const driver = getTextValue(driverNameEl);
	const plate = getTextValue(plateNumberEl);
	const destination = getTextValue(destinationEl);
	const ltNumber = getTextValue(ltNumberEl);
	const totalToQty = getTextValue(totalToQtyEl);
	const orderQty = getTextValue(orderQtyEl);

	const hasData = dbSjFile && (driver !== '-' || plate !== '-' || destination !== '-');

	if (!hasData) {
		container.classList.remove('has-data');
		container.innerHTML = '';
		return;
	}

	container.classList.add('has-data');
	container.innerHTML = `
		<div class="db-sj-info-item span-full" style="background:var(--primary); border:var(--cartoon-border); padding:8px 12px; margin-bottom:4px; display:flex; flex-direction:row; justify-content:space-between; align-items:center;">
			<span class="db-sj-info-label" style="color:#FFF; font-size:11px;">LT Number</span>
			<span class="db-sj-info-value" style="color:#FFF; font-size:15px;">${escapeHtml(ltNumber)}</span>
		</div>
		<div class="db-sj-info-item">
			<span class="db-sj-info-label">Nama Driver</span>
			<span class="db-sj-info-value" title="${escapeHtml(driver)}">${escapeHtml(driver)}</span>
		</div>
		<div class="db-sj-info-item">
			<span class="db-sj-info-label">No Polisi</span>
			<span class="db-sj-info-value">${escapeHtml(plate)}</span>
		</div>
		<div class="db-sj-info-item span-full">
			<span class="db-sj-info-label">Next Destination</span>
			<span class="db-sj-info-value" title="${escapeHtml(destination)}">${escapeHtml(destination)}</span>
		</div>
		<div class="db-sj-info-item span-full">
			<span class="db-sj-info-label">Total Order Quantity</span>
			<span class="db-sj-info-value">${escapeHtml(orderQty)}</span>
		</div>
	`;
}

async function extractDbFiles(fileList) {
	let finalFiles = [];
	let hasRar = false;

	for (let i = 0; i < fileList.length; i++) {
		const f = fileList[i];
		const ext = f.name.toLowerCase().split('.').pop();
		
		if (['xlsx', 'csv', 'xls'].includes(ext)) {
			finalFiles.push(f);
		} else if (ext === 'zip') {
			try {
				if (typeof JSZip === 'undefined') {
					console.warn('JSZip library not loaded');
					continue;
				}
				const zip = new JSZip();
				const contents = await zip.loadAsync(f);
				for (const [filename, fileData] of Object.entries(contents.files)) {
					if (!fileData.dir && !filename.startsWith('__MACOSX/')) {
						const innerExt = filename.toLowerCase().split('.').pop();
						if (['xlsx', 'csv', 'xls'].includes(innerExt)) {
							const blob = await fileData.async('blob');
							const extractedFile = new File([blob], filename, { type: blob.type, lastModified: fileData.date.getTime() });
							finalFiles.push(extractedFile);
						}
					}
				}
			} catch (err) {
				console.error('Failed to extract ZIP:', err);
				showWarningToast(`Gagal mengekstrak ${f.name}: File korup atau format tidak didukung.`);
			}
		} else if (ext === 'rar') {
			hasRar = true;
		}
	}

	if (hasRar) {
		showWarningToast('Format .rar belum didukung otomatis, mohon gunakan format .zip untuk mengekstrak data.');
	}

	return finalFiles;
}

async function handleDbToFilesAdded(fileList) {
	if (dbToDropzone) dbToDropzone.setAttribute('data-state', 'loading');

	const newFiles = await extractDbFiles(fileList);

	if (!newFiles.length) {
		if (dbToDropzone) dbToDropzone.setAttribute('data-state', dbToFiles.length ? 'loaded' : 'empty');
		showWarningToast('Hanya file Excel (.xlsx, .xls), CSV (.csv), dan ZIP (.zip) yang didukung.');
		return;
	}

	// Add to list (avoid exact duplicate names)
	const existingNames = new Set(dbToFiles.map(f => f.name));
	let addedCount = 0;
	for (const f of newFiles) {
		if (!existingNames.has(f.name)) {
			dbToFiles.push(f);
			existingNames.add(f.name);
			addedCount++;
		}
	}

	if (addedCount === 0) {
		if (dbToDropzone) dbToDropzone.setAttribute('data-state', dbToFiles.length ? 'loaded' : 'empty');
		showToast({
			type: 'error',
			title: 'File Duplikat',
			message: 'Semua file yang dipilih sudah ada di daftar upload.',
			duration: 4000
		});
		return;
	}

	// Auto-merge and load all TO files into the existing Global Source pipeline
	await mergeAndLoadToFiles();

	// Re-render AFTER merge so stats reflect the fully-loaded dataset
	renderDbToFiles();

	showToast({
		type: 'success',
		title: 'TO Management Files',
		message: `${addedCount} file ditambahkan. Total: ${dbToFiles.length} file.`,
		duration: 3000
	});
}

async function mergeAndLoadToFiles() {
	if (!dbToFiles.length) {
		return;
	}

	try {
		resetMessages();
		clearWarning();
		clearProgress();
		resetOperatorState();
		resetTripState();
		setUploaderState('loading');

		let allHeadersSet = new Set();
		let fileObjectsData = [];
		let totalRowsBefore = 0;

		for (const file of dbToFiles) {
			const extension = file.name.toLowerCase().split('.').pop();
			const aoa = extension === 'csv' ? await parseCSV(file) : await parseExcel(file);

			if (!aoa.length || aoa.length < 2) continue;

			const fileHeader = aoa[0].map(h => String(h || '').trim());
			const fileData = aoa.slice(1);

			fileHeader.forEach(h => {
				if (h) allHeadersSet.add(h);
			});

			totalRowsBefore += fileData.length;
			fileObjectsData.push({ header: fileHeader, data: fileData });
		}

		if (fileObjectsData.length === 0) {
			throw new Error('Tidak ada data valid di file yang diupload.');
		}

		let mergedHeaders = Array.from(allHeadersSet);
		let mergedData = [];

		for (const fileObj of fileObjectsData) {
			const indexMap = fileObj.header.map(h => mergedHeaders.indexOf(h));

			for (const row of fileObj.data) {
				const mergedRow = new Array(mergedHeaders.length).fill('');
				for (let i = 0; i < row.length; i++) {
					const destIndex = indexMap[i];
					if (destIndex !== -1) {
						mergedRow[destIndex] = row[i];
					}
				}
				mergedData.push(mergedRow);
			}
		}

		if (!mergedHeaders.length || !mergedData.length) {
			throw new Error('Tidak ada data valid di file yang diupload.');
		}

		headerRow = mergedHeaders;
		dataset = mergedData;
		rebuildOperatorOptionsFromDataset();
		rebuildTripOptionsFromDataset();
		updateRequiredColumnsWarning(true);

		const detectedTrip = normalize(tripInput.value);
		updateLoadedMeta(
			{ name: `${dbToFiles.length} file(s) merged`, size: dbToFiles.reduce((s, f) => s + f.size, 0) },
			dataset.length,
			detectedTrip
		);
		setUploaderState('loaded');
		refreshPreview();
	} catch (error) {
		dataset = [];
		headerRow = [];
		massUploadData = [];
		massUploadDraftData = [];
		resetOperatorState();
		resetTripState();
		resetLoadedMeta();
		setUploaderState('empty');
		setError(`Gagal membaca file TO: ${error.message}`);
	}
}

async function handleDbSjFileAdded(file) {
	if (!file) return;

	const ext = file.name.toLowerCase().split('.').pop();
	if (ext !== 'pdf') {
		showWarningToast('Hanya file PDF yang didukung untuk Surat Jalan.');
		return;
	}

	if (dbSjDropzone) dbSjDropzone.setAttribute('data-state', 'loading');

	dbSjFile = file;

	// Also feed into the existing Prealert PDF pipeline
	if (prealertPdfInput && typeof DataTransfer !== 'undefined') {
		const dt = new DataTransfer();
		dt.items.add(file);
		prealertPdfInput.files = dt.files;
	}

	await handlePrealertPdfUpload(file);
	updateDbSjExtracted();
	renderDbSjFile();

	showToast({
		type: 'success',
		title: 'Surat Jalan Uploaded',
		message: `File "${file.name}" berhasil diupload dan diekstrak.`,
		duration: 3000
	});
}

function bindDatabaseUploadEvents() {
	if (!dbToDropzone || !dbSjDropzone) return;

	// === TO Management ===
	dbToDropzone.addEventListener('dragover', (e) => { e.preventDefault(); dbToDropzone.classList.add('dragover'); });
	dbToDropzone.addEventListener('dragleave', () => dbToDropzone.classList.remove('dragover'));
	dbToDropzone.addEventListener('drop', (e) => {
		e.preventDefault();
		dbToDropzone.classList.remove('dragover');
		if (e.dataTransfer?.files?.length) handleDbToFilesAdded(e.dataTransfer.files);
	});
	dbToDropzone.addEventListener('click', (e) => {
		if (e.target.closest('.db-file-chip-remove') || e.target.closest('.db-action-btn')) return;
		if (dbToDropzone.getAttribute('data-state') === 'empty') dbToFileInput.click();
	});

	dbToFileInput.addEventListener('change', (e) => {
		if (e.target.files?.length) handleDbToFilesAdded(e.target.files);
		dbToFileInput.value = '';
	});

	dbToChooseBtn.addEventListener('click', (e) => { e.stopPropagation(); dbToFileInput.click(); });
	dbToAddMoreBtn.addEventListener('click', (e) => { e.stopPropagation(); dbToFileInput.click(); });

	// Remove individual TO file chip
	dbToLoadedFiles.addEventListener('click', async (e) => {
		const removeBtn = e.target.closest('[data-db-to-remove]');
		if (!removeBtn) return;
		const idx = parseInt(removeBtn.getAttribute('data-db-to-remove'), 10);
		if (Number.isFinite(idx) && idx >= 0 && idx < dbToFiles.length) {
			dbToDropzone.setAttribute('data-state', 'loading');
			dbToFiles.splice(idx, 1);
			if (dbToFiles.length > 0) {
				await mergeAndLoadToFiles();
				// Re-render again so stats reflect updated dataset
				renderDbToFiles();
			} else {
				clearSourceFile();
				dbToDropzone.setAttribute('data-state', 'empty');
				renderDbToFiles();
			}
		}
	});

	// === Surat Jalan ===
	dbSjDropzone.addEventListener('dragover', (e) => { e.preventDefault(); dbSjDropzone.classList.add('dragover'); });
	dbSjDropzone.addEventListener('dragleave', () => dbSjDropzone.classList.remove('dragover'));
	dbSjDropzone.addEventListener('drop', (e) => {
		e.preventDefault();
		dbSjDropzone.classList.remove('dragover');
		if (e.dataTransfer?.files?.[0]) handleDbSjFileAdded(e.dataTransfer.files[0]);
	});
	dbSjDropzone.addEventListener('click', (e) => {
		if (e.target.closest('.db-action-btn')) return;
		if (dbSjDropzone.getAttribute('data-state') === 'empty') dbSjFileInput.click();
	});

	dbSjFileInput.addEventListener('change', (e) => {
		if (e.target.files?.[0]) handleDbSjFileAdded(e.target.files[0]);
		dbSjFileInput.value = '';
	});

	dbSjChooseBtn.addEventListener('click', (e) => { e.stopPropagation(); dbSjFileInput.click(); });
	dbSjReplaceBtn.addEventListener('click', (e) => { e.stopPropagation(); dbSjFileInput.click(); });
	dbSjClearBtn.addEventListener('click', (e) => {
		e.stopPropagation();
		dbSjFile = null;
		renderDbSjFile();
		showToast({ type: 'success', title: 'Removed', message: 'Surat Jalan file dihapus.', duration: 2500 });
	});

	// === More Toggle Button ===
	bindDbMoreToggle();

	// === Extra Upload Cards ===
	bindDbExtraUploadCards();
}

// ========================================
// DATABASE MORE TOGGLE — expand/collapse
// ========================================

function bindDbMoreToggle() {
	const toggleBtn = document.getElementById('dbMoreToggleBtn');
	const moreContent = document.getElementById('dbMoreContent');
	if (!toggleBtn || !moreContent) return;

	toggleBtn.addEventListener('click', () => {
		const isOpen = moreContent.classList.contains('is-open');
		const labelEl = toggleBtn.querySelector('.database-more-label');
		
		if (isOpen) {
			moreContent.style.overflow = 'hidden';
			moreContent.classList.remove('is-open');
			moreContent.setAttribute('aria-hidden', 'true');
			toggleBtn.setAttribute('aria-expanded', 'false');
			if (labelEl) labelEl.textContent = 'More';
		} else {
			moreContent.classList.add('is-open');
			moreContent.setAttribute('aria-hidden', 'false');
			toggleBtn.setAttribute('aria-expanded', 'true');
			if (labelEl) labelEl.textContent = 'Collapse';
			
			// Wait for the slide-down transition (0.4s) to complete before allowing overflow visible for shadows
			setTimeout(() => {
				if (moreContent.classList.contains('is-open')) {
					moreContent.style.overflow = 'visible';
				}
			}, 400);
		}
	});
}

// ========================================
// EXTRA DATABASE UPLOAD CARDS — Logic
// ========================================

// Storage for uploaded files per extra card
const dbExtraFiles = {
	order: [],
	pickupAssign: [],
	pickupMgmt: [],
	pickupPerf: [],
	receiveMgmt: []
};

// Merged/parsed data for each extra card (Dynamic Header Matching)
const dbExtraMergedData = {
	order: { header: [], data: [] },
	pickupAssign: { header: [], data: [] },
	pickupMgmt: { header: [], data: [] },
	pickupPerf: { header: [], data: [] },
	receiveMgmt: { header: [], data: [] }
};

async function mergeAndLoadExtraFiles(cardKey) {
	const files = dbExtraFiles[cardKey] || [];
	if (!files.length) {
		dbExtraMergedData[cardKey] = { header: [], data: [] };
		return;
	}

	try {
		let allHeadersSet = new Set();
		let fileObjectsData = [];

		for (const file of files) {
			const extension = file.name.toLowerCase().split('.').pop();
			const aoa = extension === 'csv' ? await parseCSV(file) : await parseExcel(file);

			if (!aoa.length || aoa.length < 2) continue;

			const fileHeader = aoa[0].map(h => String(h || '').trim());
			const fileData = aoa.slice(1);

			fileHeader.forEach(h => {
				if (h) allHeadersSet.add(h);
			});

			fileObjectsData.push({ header: fileHeader, data: fileData });
		}

		if (fileObjectsData.length === 0) {
			dbExtraMergedData[cardKey] = { header: [], data: [] };
			return;
		}

		const mergedHeaders = Array.from(allHeadersSet);
		const mergedData = [];

		for (const fileObj of fileObjectsData) {
			const indexMap = fileObj.header.map(h => mergedHeaders.indexOf(h));

			for (const row of fileObj.data) {
				const mergedRow = new Array(mergedHeaders.length).fill('');
				for (let i = 0; i < row.length; i++) {
					const destIndex = indexMap[i];
					if (destIndex !== -1) {
						mergedRow[destIndex] = row[i];
					}
				}
				mergedData.push(mergedRow);
			}
		}

		dbExtraMergedData[cardKey] = { header: mergedHeaders, data: mergedData };
	} catch (error) {
		console.error(`Error merging ${cardKey} files:`, error);
		dbExtraMergedData[cardKey] = { header: [], data: [] };
	}
}

// Configuration for each extra card
const dbExtraCardConfig = [
	{
		key: 'order',
		fileInputId: 'dbOrderFileInput',
		dropzoneId: 'dbOrderDropzone',
		chooseBtnId: 'dbOrderChooseBtn',
		addMoreBtnId: 'dbOrderAddMoreBtn',
		loadedFilesId: 'dbOrderLoadedFiles',
		fileCountId: 'dbOrderFileCount',
		label: 'Order Management'
	},
	{
		key: 'pickupAssign',
		fileInputId: 'dbPickupAssignFileInput',
		dropzoneId: 'dbPickupAssignDropzone',
		chooseBtnId: 'dbPickupAssignChooseBtn',
		addMoreBtnId: 'dbPickupAssignAddMoreBtn',
		loadedFilesId: 'dbPickupAssignLoadedFiles',
		fileCountId: 'dbPickupAssignFileCount',
		label: 'Pick Up Assignment'
	},
	{
		key: 'pickupMgmt',
		fileInputId: 'dbPickupMgmtFileInput',
		dropzoneId: 'dbPickupMgmtDropzone',
		chooseBtnId: 'dbPickupMgmtChooseBtn',
		addMoreBtnId: 'dbPickupMgmtAddMoreBtn',
		loadedFilesId: 'dbPickupMgmtLoadedFiles',
		fileCountId: 'dbPickupMgmtFileCount',
		label: 'Pick Up Management'
	},
	{
		key: 'pickupPerf',
		fileInputId: 'dbPickupPerfFileInput',
		dropzoneId: 'dbPickupPerfDropzone',
		chooseBtnId: 'dbPickupPerfChooseBtn',
		addMoreBtnId: 'dbPickupPerfAddMoreBtn',
		loadedFilesId: 'dbPickupPerfLoadedFiles',
		fileCountId: 'dbPickupPerfFileCount',
		label: 'Pick Up Performance Report'
	},
	{
		key: 'receiveMgmt',
		fileInputId: 'dbReceiveMgmtFileInput',
		dropzoneId: 'dbReceiveMgmtDropzone',
		chooseBtnId: 'dbReceiveMgmtChooseBtn',
		addMoreBtnId: 'dbReceiveMgmtAddMoreBtn',
		loadedFilesId: 'dbReceiveMgmtLoadedFiles',
		fileCountId: 'dbReceiveMgmtFileCount',
		label: 'Receive Management'
	}
];

function renderDbExtraFiles(cardKey) {
	const config = dbExtraCardConfig.find(c => c.key === cardKey);
	if (!config) return;

	const loadedFilesEl = document.getElementById(config.loadedFilesId);
	const fileCountEl = document.getElementById(config.fileCountId);
	const dropzoneEl = document.getElementById(config.dropzoneId);
	const files = dbExtraFiles[cardKey] || [];

	if (!loadedFilesEl || !fileCountEl || !dropzoneEl) return;

	fileCountEl.textContent = `${files.length} file`;
	dropzoneEl.setAttribute('data-state', files.length > 0 ? 'loaded' : 'empty');

	loadedFilesEl.innerHTML = files.map((f, i) => `
		<div class="db-file-chip">
			<span class="db-file-chip-icon">${dbFileChipIconSvg(f.name)}</span>
			<span class="db-file-chip-name" title="${escapeHtml(f.name)}">${escapeHtml(f.name)}</span>
			<span class="db-file-chip-size">${dbFormatFileSize(f.size)}</span>
			<button type="button" class="db-file-chip-remove" data-db-extra-remove="${cardKey}:${i}" aria-label="Remove ${escapeHtml(f.name)}">&times;</button>
		</div>
	`).join('');
}

async function handleDbExtraFilesAdded(cardKey, fileList) {
	const config = dbExtraCardConfig.find(c => c.key === cardKey);
	if (!config) return;

	const dropzone = document.getElementById(config.dropzoneId);
	if (dropzone) dropzone.setAttribute('data-state', 'loading');

	const newFiles = await extractDbFiles(fileList);

	if (!newFiles.length) {
		if (dropzone) dropzone.setAttribute('data-state', (dbExtraFiles[cardKey] && dbExtraFiles[cardKey].length) ? 'loaded' : 'empty');
		showWarningToast('Hanya file Excel (.xlsx, .xls), CSV (.csv), dan ZIP (.zip) yang didukung.');
		return;
	}

	const existingNames = new Set((dbExtraFiles[cardKey] || []).map(f => f.name));
	let addedCount = 0;
	for (const f of newFiles) {
		if (!existingNames.has(f.name)) {
			dbExtraFiles[cardKey].push(f);
			existingNames.add(f.name);
			addedCount++;
		}
	}

	if (addedCount === 0) {
		if (dropzone) dropzone.setAttribute('data-state', (dbExtraFiles[cardKey] && dbExtraFiles[cardKey].length) ? 'loaded' : 'empty');
		showToast({
			type: 'error',
			title: 'File Duplikat',
			message: 'Semua file yang dipilih sudah ada di daftar upload.',
			duration: 4000
		});
		return;
	}

	await mergeAndLoadExtraFiles(cardKey);
	renderDbExtraFiles(cardKey);

	showToast({
		type: 'success',
		title: config.label,
		message: `${addedCount} file ditambahkan. Total: ${dbExtraFiles[cardKey].length} file.`,
		duration: 3000
	});
}

function bindDbExtraUploadCards() {
	for (const config of dbExtraCardConfig) {
		const fileInput = document.getElementById(config.fileInputId);
		const dropzone = document.getElementById(config.dropzoneId);
		const chooseBtn = document.getElementById(config.chooseBtnId);
		const addMoreBtn = document.getElementById(config.addMoreBtnId);
		const loadedFilesEl = document.getElementById(config.loadedFilesId);

		if (!fileInput || !dropzone) continue;

		// Drag & drop
		dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
		dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
		dropzone.addEventListener('drop', (e) => {
			e.preventDefault();
			dropzone.classList.remove('dragover');
			if (e.dataTransfer?.files?.length) handleDbExtraFilesAdded(config.key, e.dataTransfer.files);
		});

		// Click on dropzone (empty state)
		dropzone.addEventListener('click', (e) => {
			if (e.target.closest('.db-file-chip-remove') || e.target.closest('.db-action-btn')) return;
			if (dropzone.getAttribute('data-state') === 'empty') fileInput.click();
		});

		// File input change
		fileInput.addEventListener('change', (e) => {
			if (e.target.files?.length) handleDbExtraFilesAdded(config.key, e.target.files);
			fileInput.value = '';
		});

		// Choose / Add More buttons
		if (chooseBtn) chooseBtn.addEventListener('click', (e) => { e.stopPropagation(); fileInput.click(); });
		if (addMoreBtn) addMoreBtn.addEventListener('click', (e) => { e.stopPropagation(); fileInput.click(); });

		// Remove individual file chip
		if (loadedFilesEl) {
			loadedFilesEl.addEventListener('click', async (e) => {
				const removeBtn = e.target.closest('[data-db-extra-remove]');
				if (!removeBtn) return;
				const [key, idxStr] = (removeBtn.getAttribute('data-db-extra-remove') || '').split(':');
				const idx = parseInt(idxStr, 10);
				if (key && dbExtraFiles[key] && Number.isFinite(idx) && idx >= 0 && idx < dbExtraFiles[key].length) {
					const dz = document.getElementById(dbExtraCardConfig.find(c => c.key === key)?.dropzoneId);
					if (dz) dz.setAttribute('data-state', 'loading');
					dbExtraFiles[key].splice(idx, 1);
					await mergeAndLoadExtraFiles(key);
					renderDbExtraFiles(key);
				}
			});
		}
	}
}

init();
bindDatabaseUploadEvents();

// ========================================
// BULKY MEASUREMENT REPORT — JPEG Generator
// ========================================

const generateBulkyMeasurementBtn = document.getElementById('generateBulkyMeasurementBtn');
const downloadBulkyMeasurementBtn = document.getElementById('downloadBulkyMeasurementBtn');
const copyBulkyMeasurementBtn = document.getElementById('copyBulkyMeasurementBtn');
const closeBulkyMeasurementBtn = document.getElementById('closeBulkyMeasurementBtn');
const bulkyMeasurementModal = document.getElementById('bulkyMeasurementModal');
const bmReportPreviewWrap = document.getElementById('bmReportPreviewWrap');

let bmLastRenderedCanvas = null;

function getToNumberColumnIndex() {
	return getColumnIndexByNames(['TO Number', 'TO_Number', 'Nomor TO', 'TO number', 'to number', 'TO No', 'TO_No']);
}

function buildBulkyMeasurementData() {
	const trip = normalize(tripInput.value);
	if (!trip) return null;

	const tripRows = filterDatasetByTrip(trip);
	if (!tripRows.length) return null;

	const bulkyRows = filterDatasetByOperator(tripRows, selectedOperators);
	if (!bulkyRows.length) return null;

	const toColIdx = getToNumberColumnIndex();
	if (toColIdx === -1) return null;

	// Group by TO Number
	const toMap = new Map();
	bulkyRows.forEach(row => {
		const toNum = normalize(row[toColIdx]);
		if (!toNum) return;
		toMap.set(toNum, (toMap.get(toNum) || 0) + 1);
	});

	const entries = [];
	let totalQty = 0;
	let totalSuccess = 0;
	toMap.forEach((count, toNum) => {
		entries.push({ toNumber: toNum, qty: count, success: count });
		totalQty += count;
		totalSuccess += count;
	});

	const achievement = totalQty > 0 ? (totalSuccess / totalQty) * 100 : 0;

	return { entries, totalQty, totalSuccess, achievement };
}

async function renderSuratJalanToCanvas() {
	const selectedFile = dbSjFile || prealertUploadedPdfFile || prealertPdfInput?.files?.[0];
	if (!selectedFile) return null;

	if (typeof pdfjsLib === 'undefined' || !pdfjsLib?.getDocument) return null;

	if (pdfjsLib.GlobalWorkerOptions) {
		pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
	}

	const buffer = await selectedFile.arrayBuffer();
	const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;

	// Find the page with Origin + hub name
	let targetPageNumber = null;
	const hubNameForPdf = getSelectedHubName();
	const hubPdfPattern = new RegExp(hubNameForPdf.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s*'), 'i');

	for (let i = 1; i <= pdf.numPages; i++) {
		const candidatePage = await pdf.getPage(i);
		const textContent = await candidatePage.getTextContent();
		const pageText = textContent.items.map(item => item.str).join(' ');
		if (/Origin/i.test(pageText) && hubPdfPattern.test(pageText)) {
			targetPageNumber = i;
			break;
		}
	}

	if (!targetPageNumber) {
		for (let i = 1; i <= pdf.numPages; i++) {
			const candidatePage = await pdf.getPage(i);
			const textContent = await candidatePage.getTextContent();
			const pageText = textContent.items.map(item => item.str).join(' ');
			if (/Origin/i.test(pageText) && /First\s*Mile\s*Hub/i.test(pageText)) {
				targetPageNumber = i;
				break;
			}
		}
	}

	if (!targetPageNumber) targetPageNumber = 1;

	const page = await pdf.getPage(targetPageNumber);
	// Increase scale for much sharper PDF render
	const viewport = page.getViewport({ scale: 5.0 });
	const canvas = document.createElement('canvas');
	const context = canvas.getContext('2d', { alpha: false });
	canvas.width = Math.floor(viewport.width);
	canvas.height = Math.floor(viewport.height);
	if (!context) return null;

	await page.render({ canvasContext: context, viewport }).promise;
	return canvas;
}

function roundRect(ctx, x, y, w, h, r) {
	ctx.beginPath();
	ctx.moveTo(x + r, y);
	ctx.lineTo(x + w - r, y);
	ctx.quadraticCurveTo(x + w, y, x + w, y + r);
	ctx.lineTo(x + w, y + h - r);
	ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
	ctx.lineTo(x + r, y + h);
	ctx.quadraticCurveTo(x, y + h, x, y + h - r);
	ctx.lineTo(x, y + r);
	ctx.quadraticCurveTo(x, y, x + r, y);
	ctx.closePath();
}

function drawBulkyMeasurementCanvas(sjCanvas, measurementData, slotLabel, hubName, reportDateText) {
	const CANVAS_W = 750;
	const PAD = 30;
	const INNER_W = CANVAS_W - PAD * 2;

	// Determine Surat Jalan dimensions (cropping top ~41% to give bottom padding)
	const cropRatio = 0.41;
	const srcW = sjCanvas ? sjCanvas.width : 1000;
	const srcH = sjCanvas ? Math.min(sjCanvas.height, Math.round(srcW * cropRatio)) : 380;
	const sjImgRatio = srcH / srcW;
	
	const sjDrawW = INNER_W; // Full width, no horizontal padding
	const sjDrawH = Math.round(sjDrawW * sjImgRatio);
	const sjSectionH = sjDrawH; // Card height tightly wraps the image

	// Pre-calculate heights & positions for overlap
	const SLOT_BADGE_H = 38; 
	const SJ_SECTION_TOP = 85; // Pushed down to give badge room at top
	const slotBadgeY = SJ_SECTION_TOP - SLOT_BADGE_H - 12; // Floating above the card

	const MI_BADGE_H = 38; 
	const TABLE_TOP = SJ_SECTION_TOP + sjSectionH + 70;
	// Keep spacious top gap for SJ shadow, but reduce bottom gap so table is close like before
	const MI_BADGE_TOP = SJ_SECTION_TOP + sjSectionH + 24;

	const ROW_H = 34; // Ramping (slimmer rows)
	const SUPER_HEADER_H = 42; // Slimmer super header
	const HEADER_H = 36; // Slimmer header
	const dataRows = measurementData.entries.length;
	const TABLE_H = SUPER_HEADER_H + HEADER_H + dataRows * ROW_H + ROW_H + ROW_H; // super header + header + data + total + achievement
	const TABLE_BOTTOM = TABLE_TOP + TABLE_H;

	const FOOTER_H = 60;
	const CANVAS_H = TABLE_BOTTOM + FOOTER_H + 20;
	const SHADOW_OFFSET = 8; // For outer cartoon shadow

	// Scale up by 3x for high-resolution output
	const SCALE = 3;
	const canvas = document.createElement('canvas');
	canvas.width = (CANVAS_W + SHADOW_OFFSET) * SCALE;
	canvas.height = (CANVAS_H + SHADOW_OFFSET) * SCALE;
	const ctx = canvas.getContext('2d');
	ctx.scale(SCALE, SCALE);

	// ---- OUTER SHADOW ----
	ctx.fillStyle = '#2D2D2D';
	roundRect(ctx, SHADOW_OFFSET, SHADOW_OFFSET, CANVAS_W, CANVAS_H, 32);
	ctx.fill();

	// ---- BACKGROUND (golden/orange) ----
	ctx.fillStyle = '#FFC93C';
	roundRect(ctx, 0, 0, CANVAS_W, CANVAS_H, 32);
	ctx.fill();

	// Outer dark border
	ctx.strokeStyle = '#2D2D2D';
	ctx.lineWidth = 4;
	roundRect(ctx, 2, 2, CANVAS_W - 4, CANVAS_H - 4, 30);
	ctx.stroke();

	// Inner border line (double border effect)
	ctx.strokeStyle = 'rgba(0,0,0,0.1)';
	ctx.lineWidth = 2;
	roundRect(ctx, 12, 12, CANVAS_W - 24, CANVAS_H - 24, 24);
	ctx.stroke();

	// ---- SURAT JALAN CARD ----
	const sjCardX = PAD;
	const sjCardY = SJ_SECTION_TOP;
	const sjCardW = INNER_W;
	const sjCardH = sjSectionH;

	// Shadow
	ctx.fillStyle = '#2D2D2D';
	roundRect(ctx, sjCardX + 6, sjCardY + 6, sjCardW, sjCardH, 24);
	ctx.fill();

	// Draw the cropped Surat Jalan image as the card itself
	if (sjCanvas) {
		ctx.save();
		roundRect(ctx, sjCardX, sjCardY, sjCardW, sjCardH, 24);
		ctx.clip();
		
		// Fill white background base
		ctx.fillStyle = '#FFFFFF';
		ctx.fillRect(sjCardX, sjCardY, sjCardW, sjCardH);

		// Draw only the cropped portion, filling the entire card
		ctx.drawImage(sjCanvas, 0, 0, srcW, srcH, sjCardX, sjCardY, sjCardW, sjCardH);
		ctx.restore();

		// Card outer border
		ctx.strokeStyle = '#2D2D2D';
		ctx.lineWidth = 3;
		roundRect(ctx, sjCardX, sjCardY, sjCardW, sjCardH, 24);
		ctx.stroke();
	} else {
		// Fallback white card background
		ctx.fillStyle = '#FFFFFF';
		roundRect(ctx, sjCardX, sjCardY, sjCardW, sjCardH, 24);
		ctx.fill();
		ctx.strokeStyle = '#2D2D2D';
		ctx.lineWidth = 3;
		roundRect(ctx, sjCardX, sjCardY, sjCardW, sjCardH, 24);
		ctx.stroke();
	}

	// ---- SLOT BADGE (Floating above SJ Card) ----
	const slotText = slotLabel.toUpperCase();
	ctx.font = 'bold 18px Fredoka, sans-serif';
	const slotTextW = ctx.measureText(slotText).width;
	const slotBadgeW = slotTextW + 60; // Padding
	const slotBadgeX = PAD + 8; // Left aligned, inset from padding

	// ---- DATE BADGE (Top Right) ----
	ctx.save();
	ctx.font = 'bold 14px Fredoka, sans-serif';
	const DATE_BADGE_H = 28;
	const dateTextW = ctx.measureText(reportDateText).width;
	const dateBadgeW = dateTextW + 24; // Slightly smaller padding
	const dateBadgeX = CANVAS_W - PAD - 8 - dateBadgeW; // Symmetrical inset from right padding
	const dateBadgeY = slotBadgeY + (SLOT_BADGE_H - DATE_BADGE_H) / 2; // Center vertically with SLOT badge

	// Shadow
	ctx.fillStyle = '#2D2D2D';
	roundRect(ctx, dateBadgeX + 3, dateBadgeY + 3, dateBadgeW, DATE_BADGE_H, 14);
	ctx.fill();

	// Background
	ctx.fillStyle = '#FFFFFF';
	roundRect(ctx, dateBadgeX, dateBadgeY, dateBadgeW, DATE_BADGE_H, 14); 
	ctx.fill();

	// Border
	ctx.strokeStyle = '#2D2D2D';
	ctx.lineWidth = 2.5;
	roundRect(ctx, dateBadgeX, dateBadgeY, dateBadgeW, DATE_BADGE_H, 14);
	ctx.stroke();

	// Text
	ctx.fillStyle = '#2D2D2D';
	ctx.textAlign = 'center';
	ctx.textBaseline = 'middle';
	ctx.fillText(reportDateText, dateBadgeX + dateBadgeW / 2, dateBadgeY + DATE_BADGE_H / 2 - 1);
	ctx.restore();

	// Shadow
	ctx.fillStyle = '#2D2D2D';
	roundRect(ctx, slotBadgeX + 4, slotBadgeY + 4, slotBadgeW, SLOT_BADGE_H, 19);
	ctx.fill();

	ctx.fillStyle = '#4CAF50';
	roundRect(ctx, slotBadgeX, slotBadgeY, slotBadgeW, SLOT_BADGE_H, 19); 
	ctx.fill();
	ctx.strokeStyle = '#2D2D2D';
	ctx.lineWidth = 3;
	roundRect(ctx, slotBadgeX, slotBadgeY, slotBadgeW, SLOT_BADGE_H, 19);
	ctx.stroke();

	ctx.fillStyle = '#FFFFFF';
	ctx.textAlign = 'center';
	ctx.textBaseline = 'middle';
	ctx.fillText(slotText, slotBadgeX + slotBadgeW / 2, slotBadgeY + SLOT_BADGE_H / 2 - 1);

	// ---- TABLE ----
	const tableX = PAD;
	const tableW = INNER_W;
	// Make QTY and Success columns the same width, giving remaining space to the first column
	const col2W = 195;
	const col3W = 195;
	const col1W = tableW - col2W - col3W; // 300

	// Shadow
	ctx.fillStyle = '#2D2D2D';
	roundRect(ctx, tableX + 6, TABLE_TOP + 6, tableW, TABLE_H, 24);
	ctx.fill();

	// Draw table background
	ctx.save();
	roundRect(ctx, tableX, TABLE_TOP, tableW, TABLE_H, 24);
	ctx.clip();

	// Header background (covers both super header and header)
	ctx.fillStyle = '#E3F2FD'; // Soft classy blue to contrast with the orange badge
	ctx.fillRect(tableX, TABLE_TOP, tableW, SUPER_HEADER_H + HEADER_H);

	ctx.fillStyle = '#2D2D2D';
	ctx.textAlign = 'center';
	ctx.textBaseline = 'middle';

	// Super Header Text (centered in the full cell height)
	ctx.font = 'bold 18px Fredoka, sans-serif';
	const superHeaderTextY = TABLE_TOP + SUPER_HEADER_H / 2 - 1; 
	ctx.fillText('Count of SPX Tracking Number', tableX + col1W / 2, superHeaderTextY);
	// col2 and col3 are merged in the super header
	ctx.fillText('Measurement Info', tableX + col1W + (col2W + col3W) / 2, superHeaderTextY);

	// Header Text
	ctx.font = 'bold 18px Fredoka, sans-serif';
	const headerTextY = TABLE_TOP + SUPER_HEADER_H + HEADER_H / 2 - 1;
	ctx.fillText('TO Number', tableX + col1W / 2, headerTextY);
	ctx.fillText('Quantity', tableX + col1W + col2W / 2, headerTextY);
	ctx.fillText('Success', tableX + col1W + col2W + col3W / 2, headerTextY);

	// Data rows
	let rowY = TABLE_TOP + SUPER_HEADER_H + HEADER_H;
	measurementData.entries.forEach((entry, idx) => {
		const isEven = idx % 2 === 0;
		ctx.fillStyle = isEven ? '#FFF8E1' : '#FFFFFF';
		ctx.fillRect(tableX, rowY, tableW, ROW_H);

		ctx.fillStyle = '#2D2D2D';
		ctx.font = '600 16px Fredoka, sans-serif';
		ctx.textAlign = 'center'; // Center TO Number for symmetry
		ctx.fillText(entry.toNumber, tableX + col1W / 2, rowY + ROW_H / 2 + 1);
		ctx.fillText(String(entry.qty), tableX + col1W + col2W / 2, rowY + ROW_H / 2 + 1);
		ctx.fillText(String(entry.success), tableX + col1W + col2W + col3W / 2, rowY + ROW_H / 2 + 1);

		rowY += ROW_H;
	});

	// TOTAL row (yellow)
	ctx.fillStyle = '#FFD966';
	ctx.fillRect(tableX, rowY, tableW, ROW_H);
	ctx.fillStyle = '#2D2D2D';
	ctx.font = 'bold 18px Fredoka, sans-serif';
	ctx.textAlign = 'center';
	ctx.fillText('TOTAL', tableX + col1W / 2, rowY + ROW_H / 2 + 1);
	ctx.fillText(String(measurementData.totalQty), tableX + col1W + col2W / 2, rowY + ROW_H / 2 + 1);
	ctx.fillText(String(measurementData.totalSuccess), tableX + col1W + col2W + col3W / 2, rowY + ROW_H / 2 + 1);
	rowY += ROW_H;

	// Achievement row (green)
	ctx.fillStyle = '#4CAF50';
	ctx.fillRect(tableX, rowY, tableW, ROW_H);
	ctx.fillStyle = '#FFFFFF';
	ctx.font = 'bold 18px Fredoka, sans-serif';
	ctx.textAlign = 'center';
	ctx.fillText('Achievements', tableX + col1W / 2, rowY + ROW_H / 2 + 1);
	const achievementText = measurementData.achievement.toFixed(2) + '%';
	ctx.fillText(achievementText, tableX + col1W + (col2W + col3W) / 2, rowY + ROW_H / 2 + 1);

	ctx.restore();

	// Table grid lines
	ctx.strokeStyle = '#2D2D2D';
	ctx.lineWidth = 1.5;

	// Horizontal lines
	let lineY = TABLE_TOP + SUPER_HEADER_H; // line between super header and header
	ctx.beginPath();
	ctx.moveTo(tableX, lineY);
	ctx.lineTo(tableX + tableW, lineY);
	ctx.stroke();

	lineY = TABLE_TOP + SUPER_HEADER_H + HEADER_H; // line between header and data
	for (let i = 0; i <= dataRows + 2; i++) {
		ctx.beginPath();
		ctx.moveTo(tableX, lineY);
		ctx.lineTo(tableX + tableW, lineY);
		ctx.stroke();
		lineY += ROW_H;
	}

	// Vertical lines
	ctx.beginPath();
	// line 1 (between col1 and col2) - goes all the way up and down
	ctx.moveTo(tableX + col1W, TABLE_TOP);
	ctx.lineTo(tableX + col1W, TABLE_TOP + TABLE_H);
	ctx.stroke();

	ctx.beginPath();
	// line 2 (between col2 and col3) - starts below super header, stops before achievement
	ctx.moveTo(tableX + col1W + col2W, TABLE_TOP + SUPER_HEADER_H);
	ctx.lineTo(tableX + col1W + col2W, TABLE_TOP + SUPER_HEADER_H + HEADER_H + dataRows * ROW_H + ROW_H);
	ctx.stroke();

	// Table outer border again (on top of clip)
	ctx.strokeStyle = '#2D2D2D';
	ctx.lineWidth = 3;
	roundRect(ctx, tableX, TABLE_TOP, tableW, TABLE_H, 24);
	ctx.stroke();

	// ---- MEASUREMENT INFO BADGE (Floating above Table) ----
	const miText = 'MEASUREMENT INFO';
	ctx.font = 'bold 18px Fredoka, sans-serif';
	const miTextW = ctx.measureText(miText).width;
	const miBadgeW = miTextW + 60; // Padding
	const miBadgeX = PAD + 8; // Left aligned, inset from padding

	// Shadow
	ctx.fillStyle = '#2D2D2D';
	roundRect(ctx, miBadgeX + 4, MI_BADGE_TOP + 4, miBadgeW, MI_BADGE_H, 19);
	ctx.fill();

	ctx.fillStyle = '#FF8C42';
	roundRect(ctx, miBadgeX, MI_BADGE_TOP, miBadgeW, MI_BADGE_H, 19); 
	ctx.fill();
	ctx.strokeStyle = '#2D2D2D';
	ctx.lineWidth = 3;
	roundRect(ctx, miBadgeX, MI_BADGE_TOP, miBadgeW, MI_BADGE_H, 19);
	ctx.stroke();

	ctx.fillStyle = '#FFFFFF';
	ctx.textAlign = 'center';
	ctx.textBaseline = 'middle';
	ctx.fillText(miText, miBadgeX + miBadgeW / 2, MI_BADGE_TOP + MI_BADGE_H / 2 - 1);

	// ---- FOOTER BADGE ----
	const footerText = hubName;
	ctx.font = 'bold 20px Fredoka, sans-serif';
	const footerTextW = ctx.measureText(footerText).width;
	const FOOTER_BADGE_H = 44;
	const footerBadgeW = footerTextW + 60; // Padding
	const footerBadgeX = (CANVAS_W - footerBadgeW) / 2;
	const footerBadgeY = CANVAS_H - FOOTER_BADGE_H - 16; // Mepet ke bawah (dekat dengan inner border)

	// Shadow
	ctx.fillStyle = '#2D2D2D';
	roundRect(ctx, footerBadgeX + 4, footerBadgeY + 4, footerBadgeW, FOOTER_BADGE_H, 22);
	ctx.fill();

	// Background
	ctx.fillStyle = '#FFFFFF';
	roundRect(ctx, footerBadgeX, footerBadgeY, footerBadgeW, FOOTER_BADGE_H, 22); 
	ctx.fill();

	// Border
	ctx.strokeStyle = '#2D2D2D';
	ctx.lineWidth = 3;
	roundRect(ctx, footerBadgeX, footerBadgeY, footerBadgeW, FOOTER_BADGE_H, 22);
	ctx.stroke();

	// Text
	ctx.fillStyle = '#2D2D2D';
	ctx.textAlign = 'center';
	ctx.textBaseline = 'middle';
	ctx.fillText(footerText, CANVAS_W / 2, footerBadgeY + FOOTER_BADGE_H / 2 - 1);

	return canvas;
}

function openBulkyMeasurementModal() {
	if (!bulkyMeasurementModal) return;
	bulkyMeasurementModal.classList.remove('is-closing');
	bulkyMeasurementModal.classList.add('open');
	bulkyMeasurementModal.setAttribute('aria-hidden', 'false');
	document.body.style.overflow = 'hidden';
	if (typeof window._sfx?.modalOpen === 'function') window._sfx.modalOpen();
}

function closeBulkyMeasurementModal() {
	if (!bulkyMeasurementModal) return;
	bulkyMeasurementModal.classList.add('is-closing');
	bulkyMeasurementModal.classList.remove('open');
	if (typeof window._sfx?.modalClose === 'function') window._sfx.modalClose();
	setTimeout(() => {
		bulkyMeasurementModal.classList.remove('is-closing');
		bulkyMeasurementModal.setAttribute('aria-hidden', 'true');
		document.body.style.overflow = '';
	}, 250);
}

async function generateBulkyMeasurementReport() {
	// Validate
	if (!dataset.length) {
		setError('Upload file TO Management terlebih dahulu.');
		return;
	}

	const trip = normalize(tripInput.value);
	if (!trip) {
		setError('Upload Surat Jalan terlebih dahulu untuk mendapatkan LT Number.');
		return;
	}

	if (!selectedOperators.length) {
		setError('Pilih operator bulky terlebih dahulu.');
		return;
	}

	const toColIdx = getToNumberColumnIndex();
	if (toColIdx === -1) {
		setError('Kolom TO Number tidak ditemukan di file TO Management. Pastikan ada kolom "TO Number".');
		return;
	}

	const measurementData = buildBulkyMeasurementData();
	if (!measurementData || !measurementData.entries.length) {
		setError('Tidak ada data TO Number ditemukan untuk operator yang dipilih.');
		return;
	}

	setButtonLoading(generateBulkyMeasurementBtn, true, 'Generating...');
	showProcessingToast(['Rendering Surat Jalan', 'Building Measurement Report']);

	try {
		// Render surat jalan PDF page to canvas
		let sjCanvas = null;
		try {
			sjCanvas = await renderSuratJalanToCanvas();
		} catch (err) {
			console.warn('Could not render Surat Jalan:', err);
		}

		const slotLabel = `SLOT - ${getSlotTripNumber(slotSelect?.value)}`;
		const hubName = getSelectedHubName();

		let reportDateText = getIndonesianReportDateWithDayTitleCaseNatural();
		if (reportDateText === '-') {
			const [year, month, day] = getTodayInputDate().split('-');
			const d = new Date(parseInt(year, 10), parseInt(month, 10) - 1, parseInt(day, 10));
			const dayNames = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
			const monthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
			reportDateText = `${dayNames[d.getDay()]}, ${parseInt(day, 10)} ${monthNames[parseInt(month, 10) - 1]} ${year}`;
		}

		// Draw the full report on canvas
		const reportCanvas = drawBulkyMeasurementCanvas(sjCanvas, measurementData, slotLabel, hubName, reportDateText);
		bmLastRenderedCanvas = reportCanvas;

		// Show in modal preview
		if (bmReportPreviewWrap) {
			bmReportPreviewWrap.innerHTML = '';
			bmReportPreviewWrap.appendChild(reportCanvas);
		}

		openBulkyMeasurementModal();
		showReportSuccessToast(['Bulky Measurement Report generated']);
	} catch (error) {
		setError(`Gagal generate Bulky Measurement Report: ${error.message}`);
		showReportErrorToast();
	} finally {
		setButtonLoading(generateBulkyMeasurementBtn, false);
	}
}

function downloadBulkyMeasurementPng() {
	if (!bmLastRenderedCanvas) {
		setError('Generate report terlebih dahulu.');
		return;
	}

	const slotNumber = getSlotTripNumber(slotSelect?.value);
	const datePart = getSelectedReportDateFormatted() || formatDateDDMMYYYY(getTodayInputDate()) || 'report';
	const fileName = `Bulky Measurement - Slot ${slotNumber} - ${datePart}.png`;

	bmLastRenderedCanvas.toBlob((blob) => {
		if (!blob) {
			setError('Gagal membuat file PNG.');
			return;
		}

		const url = URL.createObjectURL(blob);
		const anchor = document.createElement('a');
		anchor.href = url;
		anchor.download = fileName;
		document.body.appendChild(anchor);
		anchor.click();
		anchor.remove();
		URL.revokeObjectURL(url);

		showToast({
			type: 'success',
			title: 'Downloaded',
			message: `${fileName} berhasil didownload.`,
			duration: 3000
		});
	}, 'image/png');
}

async function copyBulkyMeasurementImage() {
	if (!bmLastRenderedCanvas) {
		setError('Generate report terlebih dahulu.');
		return;
	}

	try {
		const blob = await new Promise((resolve) => bmLastRenderedCanvas.toBlob(resolve, 'image/png'));
		if (!blob) {
			setError('Gagal membuat gambar untuk disalin.');
			return;
		}

		await navigator.clipboard.write([
			new ClipboardItem({ 'image/png': blob })
		]);

		showToast({
			type: 'success',
			title: 'Copied!',
			message: 'Gambar report berhasil disalin ke clipboard.',
			duration: 3000
		});
	} catch (err) {
		console.error('Copy failed:', err);
		setError('Gagal menyalin gambar. Pastikan browser mendukung fitur ini.');
	}
}

// Event bindings
if (generateBulkyMeasurementBtn) {
	generateBulkyMeasurementBtn.addEventListener('click', generateBulkyMeasurementReport);
}

if (downloadBulkyMeasurementBtn) {
	downloadBulkyMeasurementBtn.addEventListener('click', downloadBulkyMeasurementPng);
}

if (copyBulkyMeasurementBtn) {
	copyBulkyMeasurementBtn.addEventListener('click', copyBulkyMeasurementImage);
}

if (closeBulkyMeasurementBtn) {
	closeBulkyMeasurementBtn.addEventListener('click', closeBulkyMeasurementModal);
}

if (bulkyMeasurementModal) {
	bulkyMeasurementModal.addEventListener('click', (e) => {
		if (e.target === bulkyMeasurementModal) closeBulkyMeasurementModal();
	});

	document.addEventListener('keydown', (e) => {
		if (e.key === 'Escape' && bulkyMeasurementModal.classList.contains('open')) {
			closeBulkyMeasurementModal();
		}
	});
}

// Pre Alert Toggle
// Prealert toggle removed, now handled by folder tabs

// Pickup Report Export Logic
const copyPickupReportBtn = document.getElementById('copyPickupReportBtn');
const downloadPickupReportBtn = document.getElementById('downloadPickupReportBtn');
const pickupTableContainer = document.getElementById('pickupTableContainer');

async function generatePickupReportCanvas() {
	if (!pickupTableContainer) return null;
	try {
		// Use html2canvas with scale 3 or 4 for super sharp quality
		return await html2canvas(pickupTableContainer, { scale: 4, backgroundColor: null });
	} catch (err) {
		console.error("html2canvas error:", err);
		return null;
	}
}

if (copyPickupReportBtn) {
	copyPickupReportBtn.addEventListener('click', async () => {
		const btn = copyPickupReportBtn;
		const originalText = btn.innerHTML;
		btn.innerHTML = 'COPYING...';
		btn.disabled = true;
		
		const canvas = await generatePickupReportCanvas();
		if (!canvas) {
			setError('Gagal membuat gambar untuk disalin.');
			btn.innerHTML = originalText;
			btn.disabled = false;
			return;
		}
		
		try {
			const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
			await navigator.clipboard.write([
				new ClipboardItem({ 'image/png': blob })
			]);
			showToast({ type: 'success', title: 'Copied!', message: 'Pickup Report disalin dengan kualitas super tajam.', duration: 3000 });
		} catch (err) {
			console.error('Copy failed:', err);
			setError('Gagal menyalin. Pastikan browser mendukung fitur ini.');
		}
		btn.innerHTML = originalText;
		btn.disabled = false;
	});
}

if (downloadPickupReportBtn) {
	downloadPickupReportBtn.addEventListener('click', async () => {
		const btn = downloadPickupReportBtn;
		const originalText = btn.innerHTML;
		btn.innerHTML = 'DOWNLOADING...';
		btn.disabled = true;

		const canvas = await generatePickupReportCanvas();
		if (!canvas) {
			setError('Gagal membuat gambar untuk di-download.');
			btn.innerHTML = originalText;
			btn.disabled = false;
			return;
		}
		
		canvas.toBlob((blob) => {
			const url = URL.createObjectURL(blob);
			const a = document.createElement('a');
			a.href = url;
			a.download = `Pickup_Success_Rate_${getTodayInputDate()}.png`;
			document.body.appendChild(a);
			a.click();
			document.body.removeChild(a);
			URL.revokeObjectURL(url);
			
			showToast({ type: 'success', title: 'Downloaded', message: 'Pickup Report berhasil didownload.', duration: 3000 });
			btn.innerHTML = originalText;
			btn.disabled = false;
		}, 'image/png');
	});
}
