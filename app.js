let dataset = [];
let headerRow = [];
let activeTool = 'mass';
let selectedOperators = [];
let availableOperators = [];
let availableTrips = [];
let operatorColumnIndex = -1;
let tripColumnIndex = -1;
let massUploadData = [];
let massUploadDraftData = [];
let toastTimer = null;
let toastHideTimer = null;
let toolSwitchTimer = null;
let notificationAudioContext = null;
let lastNotificationSoundAt = 0;
let notificationLogEntries = [];
let prealertUploadedPdfFile = null;
let customErrorAudioElement = null;
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
const copyTrackingBtn = document.getElementById('copyTrackingBtn');
const copyLinehaulReportBtn = document.getElementById('copyLinehaulReportBtn');
const copyBulkyReportBtn = document.getElementById('copyBulkyReportBtn');
const copyHourlyReportBtn = document.getElementById('copyHourlyReportBtn');
const generateTOBtn = document.getElementById('generateTOBtn');
const generateAllBtn = document.getElementById('generateAllBtn');
const generateAllInlineBtn = document.getElementById('generateAllInlineBtn');
const generateEmailFab = document.getElementById('generateEmailFab');
const floatingMenu = document.getElementById('floatingMenu');
const floatingMenuToggle = document.getElementById('floatingMenuToggle');
const floatingMenuDropdown = document.getElementById('floatingMenuDropdown');
const totalDataEl = document.getElementById('totalData');
const filteredDataEl = document.getElementById('filteredData');
const previewHeadLeft = document.getElementById('previewHeadLeft');
const previewHeadRight = document.getElementById('previewHeadRight');
const previewBody = document.getElementById('previewBody');
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
const prefersReducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

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
	const etaTime = shiftClockByMinutes(linehaulTemplateData.staTime, -30);
	const ataTime = formatClockWithEarly(linehaulTemplateData.ataTime, etaTime);
	const stdTime = extractClockTime(linehaulTemplateData.stdTime);
	const atdCompareClock = linehaulTemplateData.stfTime !== '-' ? linehaulTemplateData.stfTime : linehaulTemplateData.stdTime;
	const atdTime = formatAtdWithEarly(linehaulTemplateData.atdTime, atdCompareClock);
	const sealCode = normalize(linehaulTemplateData.sealCode || '-').toUpperCase() || '-';
	const dateLabel = getIndonesianReportDateWithDayTitleCaseNatural();

	return [
		'Daily Report Ciputat 4 First Mile Hub',
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
		`Qty Of To : ${totalToQty}`,
		`Qty Of Parcel : ${orderQty}`,
		`Hv To Quantity : ${hvQty}`,
		'',
		'Terima-Kasih "Good-Luck"'
	].join('\n');
}

function buildBulkyReportTemplateText() {
	const { slot } = getReportSelectionContext();
	const slotNumber = getSlotTripNumber(slot);
	const bulkyQty = getBulkyQtyFromReportData();
	const dateLabel = getIndonesianReportDateWithDayTitleCase();

	return [
		'Mass Upload Bulky Ciputat 4 First Mile Hub',
		`Slot ${slotNumber} : ${bulkyQty}`,
		dateLabel
	].join('\n');
}

function buildHourlyReportTemplateText() {
	const dateLabel = getIndonesianReportDateWithDayTitleCase();
	return [
		'Hourly Performance Dialogue Dashboard',
		'Ciputat 4 First Mile Hub',
		dateLabel
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

	return `PRE ALERT - AMH CIPUTAT 4 FIRST MILE TO ${subjectDestination} - TRIP ${slotTripNumber} - ${formattedDate}`;
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
			<p style="margin:0 0 10px;">Dear All</p>
			<p style="margin:0 0 12px;">Berikut Terlampir Surat Jalan From Ciputat 4 AMH To ${escapeHtml(destination)}</p>
			<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;max-width:560px;border-collapse:collapse;margin:8px 0 12px;background:#ffffff;border:1px solid #d1d5db;">
				<thead>
					<tr>
						<th colspan="2" style="border:1px solid #d1d5db;padding:10px 12px;background:#8dc63f;color:#ffffff;text-align:center;font-size:14px;font-weight:700;letter-spacing:0.04em;">CIPUTAT 4 FIRST MILE</th>
					</tr>
				</thead>
				<tbody>
					${htmlRows}
				</tbody>
			</table>
			<p style="margin:0 0 10px;font-size:12px;line-height:1.55;color:#334155;"><b>Noted:</b> Apabila dalam waktu 3 jam setelah shipment tiba tidak terdapat feedback dari pihak penerima, maka seluruh tanggung jawab terkait paket dianggap telah diterima oleh next station.</p>
			<p style="margin:0;">Terima-kasih</p>
		</div>
	`;

	const previewHtmlRows = rows
		.map(([label, value]) => `
			<tr>
				<td style="border:2px solid #2D2D2D;padding:9px 12px;font-size:13px;font-weight:700;background:#8dc63f;color:#2D2D2D;width:40%;vertical-align:top;text-transform:uppercase;letter-spacing:0.02em;">${escapeHtml(label)}</td>
				<td style="border:2px solid #2D2D2D;padding:9px 12px;font-size:13px;color:#2D2D2D;vertical-align:top;background:#e8f5d9;">${escapeHtml(value)}</td>
			</tr>
		`)
		.join('');

	const previewHtml = `
		<div style="font-family:Fredoka,Segoe UI,Arial,sans-serif;font-size:14px;line-height:1.55;color:#2D2D2D;max-width:620px;">
			<p style="margin:0 0 10px;">Dear All</p>
			<p style="margin:0 0 12px;">Berikut Terlampir Surat Jalan From Ciputat 4 AMH To ${escapeHtml(destination)}</p>
			<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;max-width:560px;border-collapse:collapse;margin:8px 0 12px;background:#ffffff;border:2.5px solid #2D2D2D;border-radius:14px;overflow:hidden;box-shadow:4px 4px 0 #2D2D2D;">
				<thead>
					<tr>
						<th colspan="2" style="border:2px solid #2D2D2D;padding:10px 12px;background:#5a9e2f;color:#ffffff;text-align:center;font-size:15px;font-weight:800;letter-spacing:0.06em;text-transform:uppercase;text-shadow:1px 1px 0 rgba(0,0,0,0.15);">CIPUTAT 4 FIRST MILE</th>
					</tr>
				</thead>
				<tbody>
					${previewHtmlRows}
				</tbody>
			</table>
			<p style="margin:0 0 10px;font-size:12px;line-height:1.55;color:#5A5A5A;"><b>Noted:</b> Apabila dalam waktu 3 jam setelah shipment tiba tidak terdapat feedback dari pihak penerima, maka seluruh tanggung jawab terkait paket dianggap telah diterima oleh next station.</p>
			<p style="margin:0;">Terima-kasih</p>
		</div>
	`;

	const text = [
		'Dear All',
		'',
		`Berikut Terlampir Surat Jalan From Ciputat 4 AMH To ${destination}`,
		'',
		'=== CIPUTAT 4 FIRST MILE ===',
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
		source: 'Ciputat 4 First Mile Hub',
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

	return `Surat Jalan Slot ${safeSlot} - ${safePlate} - Ciputat 4 First Mile to Transit Point Depok DC - ${safeDate}.jpg`;
}

async function downloadPrealertFirstPageJpg() {
	const selectedFile = prealertUploadedPdfFile || prealertPdfInput?.files?.[0];
	if (!selectedFile) {
		setWarning('Upload PDF terlebih dahulu sebelum download JPG halaman 1.');
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
		const targetPageNumber = 2;
		if (pdf.numPages < targetPageNumber) {
			throw new Error(`PDF hanya memiliki ${pdf.numPages} halaman. Tidak ada halaman 2.`);
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

		setSuccess('JPG halaman 2 berhasil diunduh.');
	} catch (error) {
		setError(`Gagal membuat JPG halaman 2: ${error.message}`);
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
	reportSummary.textContent = `Trip: ${trip} | Slot: ${slot} | Report Date: ${reportDate}`;
	syncPrealertTripFromReportSlot();
	updatePrealertEmailPreview();
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
		removeBtn.innerHTML = '<svg width="11" height="11" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" fill="rgba(0,0,0,0.1)" stroke="currentColor" stroke-width="2.2"/><path d="M15 9L9 15M9 9l6 6" stroke="currentColor" stroke-width="2.8" stroke-linecap="round"/></svg>';
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

function fixMassUploadDimensions() {
	if (!massUploadDraftData.length) {
		showWarningToast('Tidak ada data untuk diperbaiki.');
		return 0;
	}

	let repairedCount = 0;

	massUploadDraftData = massUploadDraftData.map((item) => {
		if (normalize(item.weight).toUpperCase() !== 'FIX_ME') {
			return item;
		}

		const [lengthValue, widthValue, heightValue] = shuffleDimensions([20, 25, 25]);
		repairedCount += 1;

		return {
			...item,
			weight: '2.08',
			length: String(lengthValue),
			width: String(widthValue),
			height: String(heightValue)
		};
	});

	if (!repairedCount) {
		showWarningToast('Tidak ada baris FIX_ME yang perlu diperbaiki.');
		return 0;
	}

	renderMassUploadEditorTable();
	showToast({
		type: 'warning',
		title: 'Dimensions repaired',
		message: `${repairedCount} rows fixed successfully`
	});

	return repairedCount;
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
	if (toStatusSummaryBody) {
		toStatusSummaryBody.innerHTML = '';
	}
	const trip = normalize(tripInput.value);
	const byTrip = trip ? filterDatasetByTrip(trip) : [];
	const filteredRows = filterDatasetByOperator(byTrip, selectedOperators);

	if (toStatusSummary) {
		toStatusSummary.classList.remove('hidden');
	}

	if (previewHeadLeft && previewHeadRight) {
		previewHeadLeft.textContent = 'No';
		previewHeadLeft.style.width = '70px';
		previewHeadRight.textContent = 'SPX Tracking Number';
	}

	syncMassUploadData(filteredRows);

	if (viewAllPreviewBtn) {
		viewAllPreviewBtn.disabled = !filteredRows.length;
	}

	animateStatValue(totalDataEl, byTrip.length);
	animateStatValue(filteredDataEl, filteredRows.length);

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

	if (!filteredRows.length) {
		const row = document.createElement('tr');
		row.innerHTML = '<td colspan="2">Belum ada data.</td>';
		previewBody.appendChild(row);
		return;
	}

	filteredRows.forEach((rowData, index) => {
		const row = document.createElement('tr');
		row.innerHTML = `<td>${index + 1}</td><td>${escapeHtml(normalize(rowData[COL.SPX]))}</td>`;
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
		const repairedCount = fixMassUploadDimensions();
		if (repairedCount > 0) {
			setMassUploadStatus(`${repairedCount} rows fixed. Review and apply changes.`, 'success');
			return;
		}
		setMassUploadStatus('No FIX_ME rows detected to fix.', 'warning');
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
	reportDateInput.addEventListener('change', updateReportSummary);

	copyLinehaulReportBtn?.addEventListener('click', () => {
		copyReportTemplateText('linehaul');
	});

	copyBulkyReportBtn?.addEventListener('click', () => {
		copyReportTemplateText('bulky');
	});

	copyHourlyReportBtn?.addEventListener('click', () => {
		copyReportTemplateText('hourly');
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

init();
