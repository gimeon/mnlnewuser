// ============================================================================
// 설정
// ============================================================================

// 엑셀 컬럼명 -> 내부 필드 매핑. 후보를 배열로 주면 순서대로 시도.
const COLUMN_MAP = {
  organization:     ['소속기관', '기관', '기관명', 'organization'],
  department:       ['부서', '부서명', 'department'],
  name:             ['이름', '성명', 'name'],
  title:            ['직책', '직급', 'title', 'position'],
  phone:            ['전화번호', '휴대폰', '연락처', 'phone', 'mobile'],
  registrationTime: ['전화번호 등록 일시', '가입일', '가입일자', '등록일', '등록일시'],
};

const REPORT_HEADER_LABEL = '모바일 뉴스리더 제공';
const HISTORY_KEY = 'dailyReportHistory_v1';
const LAST_AUTHOR_KEY = 'dailyReportLastAuthor_v1';
const HISTORY_MAX_ITEMS = 10;
const MAX_FILE_SIZE_MB = 15;

// ============================================================================
// DOM 참조
// ============================================================================

const $ = (id) => document.getElementById(id);
const STATUS_EL         = $('status');
const RESULT_SECTION    = $('resultSection');
const RESULT_TEXT       = $('resultText');
const DEBUG_AREA        = $('debugArea');
const DROPZONE          = $('dropzone');
const FILE_INPUT        = $('fileInput');
const COPY_BTN          = $('copyBtn');
const SAVE_RECORD_BTN   = $('saveRecordBtn');
const HISTORY_LIST      = $('historyList');
const LAST_REPORT_DISP  = $('lastReportDisplay');
const FILE_LOADED       = $('fileLoaded');
const LOADED_FILE_NAME  = $('loadedFileName');
const LOADED_FILE_META  = $('loadedFileMeta');
const REMOVE_FILE_BTN   = $('removeFileBtn');
const SAVE_POPOVER      = $('savePopover');
const POPOVER_AUTHOR    = $('popoverAuthorInput');
const POPOVER_HINT      = $('popoverHint');
const POPOVER_CANCEL    = $('popoverCancelBtn');
const POPOVER_CONFIRM   = $('popoverConfirmBtn');

// 세션 상태
let currentRows = null;     // 최근 파싱된 엑셀 행들
let currentReport = null;   // 최근 생성된 보고 텍스트/엔트리
let isBusy = false;         // 파일 로드/파싱 중이면 true

function setBusy(flag) {
  isBusy = flag;
  document.body.classList.toggle('is-busy', flag);
  [FILE_INPUT, REMOVE_FILE_BTN].forEach((el) => { if (el) el.disabled = flag; });
  if (flag) {
    SAVE_RECORD_BTN.disabled = true;
    COPY_BTN.disabled = true;
  } else {
    const hasReport = !!currentReport;
    SAVE_RECORD_BTN.disabled = !hasReport;
    COPY_BTN.disabled = !hasReport;
  }
  const startBtn = document.getElementById('startChangeBtn');
  const endBtn = document.getElementById('endChangeBtn');
  if (startBtn) startBtn.disabled = flag;
  if (endBtn) endBtn.disabled = flag;
  document.querySelectorAll('input[name="startPoint"]').forEach((el) => {
    el.disabled = flag || (el.value === 'lastReport' && !getLastReportTime());
  });
  document.querySelectorAll('input[name="endPoint"]').forEach((el) => {
    el.disabled = flag;
  });
}

// 브라우저가 status 메시지를 실제로 페인트할 수 있도록 다음 프레임까지 양보
function yieldToUI() {
  return new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 0)));
}

// ============================================================================
// 시간 포맷 헬퍼
// ============================================================================

const pad = (n) => String(n).padStart(2, '0');

const WEEKDAYS_KO = ['일', '월', '화', '수', '목', '금', '토'];
function formatLocalFull(d) {
  const dow = WEEKDAYS_KO[d.getDay()];
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}(${dow}) ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function formatDuration(startMs, endMs) {
  const diffMin = Math.max(0, Math.floor((endMs - startMs) / 60000));
  const hours = Math.floor(diffMin / 60);
  const minutes = diffMin % 60;
  if (hours === 0) return `${minutes}분`;
  return `${hours}시간 ${minutes}분`;
}

// ============================================================================
// 저장 기록 — 우선 /api/reports (공유 D1), 실패 시 localStorage 폴백
// ============================================================================

let _historyCache = [];

async function fetchHistoryFromApi() {
  const res = await fetch('/api/reports', { headers: { 'Accept': 'application/json' } });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return await res.json();
}

function loadHistoryLocal() {
  try {
    const raw = localStorage.getItem(HISTORY_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch { return []; }
}

async function loadHistory() {
  try {
    _historyCache = await fetchHistoryFromApi();
  } catch (err) {
    console.warn('[history] API 실패, localStorage로 폴백:', err.message);
    _historyCache = loadHistoryLocal();
  }
  return _historyCache;
}

async function saveHistoryEntry(entry) {
  try {
    const res = await fetch('/api/reports', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(entry),
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    await loadHistory();
    return;
  } catch (err) {
    console.warn('[history] API 저장 실패, localStorage로 폴백:', err.message);
  }
  // Fallback
  try {
    const list = loadHistoryLocal();
    list.unshift(entry);
    if (list.length > HISTORY_MAX_ITEMS) list.length = HISTORY_MAX_ITEMS;
    localStorage.setItem(HISTORY_KEY, JSON.stringify(list));
    _historyCache = list;
  } catch {}
}

async function clearAllHistory() {
  try {
    const res = await fetch('/api/reports', { method: 'DELETE' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    await loadHistory();
    return;
  } catch (err) {
    console.warn('[history] API 삭제 실패, localStorage로 폴백:', err.message);
  }
  try { localStorage.removeItem(HISTORY_KEY); } catch {}
  _historyCache = [];
}

function getLastReportTime() {
  return _historyCache.length > 0 ? new Date(_historyCache[0].createdAt) : null;
}

// ============================================================================
// 집계 시작점 / 종료점 선택
// ============================================================================

const CUSTOM_START_INPUT = $('customStartTime');
const CUSTOM_END_INPUT   = $('customEndTime');
const CUSTOM_START_WRAP  = document.querySelector('.custom-start-wrap');
const CUSTOM_END_WRAP    = document.querySelector('.custom-end-wrap');
const START_POPOVER      = $('startPopover');
const END_POPOVER        = $('endPopover');
const START_CHANGE_BTN   = $('startChangeBtn');
const END_CHANGE_BTN     = $('endChangeBtn');
const START_POINT_TEXT   = $('startPointText');
const START_POINT_HINT   = $('startPointHint');
const END_POINT_TEXT     = $('endPointText');
const END_POINT_HINT     = $('endPointHint');
const PERIOD_SUMMARY     = $('periodSummary');

let currentFileMeta = null; // { name, lastModified }

const START_LABELS = {
  lastReport: '마지막 보고 시각부터',
  endZero:    '종료일 0시부터',
  last24h:    '종료점 24시간 전부터',
  custom:     '직접 설정',
};
const END_LABELS = {
  fileModified: '파일 최종 수정일까지',
  now:          '현재 시각까지',
  custom:       '직접 설정',
};

function getSelectedStartOption() {
  const el = document.querySelector('input[name="startPoint"]:checked');
  return el ? el.value : null;
}
function getSelectedEndOption() {
  const el = document.querySelector('input[name="endPoint"]:checked');
  return el ? el.value : null;
}

function getEndPoint() {
  const opt = getSelectedEndOption();
  if (opt === 'fileModified') return currentFileMeta ? new Date(currentFileMeta.lastModified) : null;
  if (opt === 'now') return new Date();
  if (opt === 'custom') {
    const d = customEndPicker.selectedDates && customEndPicker.selectedDates[0];
    return d ? new Date(d) : null;
  }
  return null;
}

function getStartPoint(endDate) {
  const opt = getSelectedStartOption();
  if (opt === 'lastReport') return getLastReportTime();
  if (opt === 'endZero') {
    if (!endDate) return null;
    const d = new Date(endDate); d.setHours(0, 0, 0, 0); return d;
  }
  if (opt === 'last24h') {
    if (!endDate) return null;
    return new Date(endDate.getTime() - 24 * 3600 * 1000);
  }
  if (opt === 'custom') {
    const d = customStartPicker.selectedDates && customStartPicker.selectedDates[0];
    return d ? new Date(d) : null;
  }
  return null;
}

// flatpickr 인스턴스 (시작점 / 종료점 직접 설정용)
const customStartPicker = flatpickr(CUSTOM_START_INPUT, {
  locale: 'ko', enableTime: true, time_24hr: true,
  dateFormat: 'Y-m-d(D) H:i', allowInput: true, minuteIncrement: 10,
  onChange: () => {
    updateStartPointSummary();
    updatePeriodSummary();
    if (currentRows) renderReport();
    if (getSelectedStartOption() === 'custom') closeAllPeriodPopovers();
  },
});
const customEndPicker = flatpickr(CUSTOM_END_INPUT, {
  locale: 'ko', enableTime: true, time_24hr: true,
  dateFormat: 'Y-m-d(D) H:i', allowInput: true, minuteIncrement: 10,
  onChange: () => {
    updateEndPointSummary();
    updateStartPointSummary();
    updateStartPointHints();
    updatePeriodSummary();
    if (currentRows) renderReport();
    if (getSelectedEndOption() === 'custom') closeAllPeriodPopovers();
  },
});

// ---- 요약 라인 업데이트 ----
function updateStartPointSummary() {
  if (!START_POINT_TEXT) return;
  const opt = getSelectedStartOption();
  const endDate = getEndPoint();
  let hint = '';
  if (opt === 'lastReport') {
    const last = getLastReportTime();
    hint = last ? formatLocalFull(last) : '아직 저장된 보고 없음';
  } else if (opt === 'endZero') {
    if (endDate) { const d = new Date(endDate); d.setHours(0, 0, 0, 0); hint = formatLocalFull(d); }
    else hint = '종료점 필요';
  } else if (opt === 'last24h') {
    hint = endDate ? formatLocalFull(new Date(endDate.getTime() - 24 * 3600 * 1000)) : '종료점 필요';
  } else if (opt === 'custom') {
    const d = customStartPicker.selectedDates[0];
    hint = d ? formatLocalFull(d) : '시각 선택 필요';
  }
  START_POINT_TEXT.textContent = START_LABELS[opt] || '-';
  START_POINT_HINT.textContent = hint;
  START_POINT_HINT.classList.remove('accent-latest');
}

function updateEndPointSummary() {
  if (!END_POINT_TEXT) return;
  const opt = getSelectedEndOption();
  let hint = '';
  if (opt === 'fileModified') {
    hint = currentFileMeta ? formatLocalFull(new Date(currentFileMeta.lastModified)) : '파일 업로드 필요';
  } else if (opt === 'now') {
    hint = formatLocalFull(new Date());
  } else if (opt === 'custom') {
    const d = customEndPicker.selectedDates[0];
    hint = d ? formatLocalFull(d) : '시각 선택 필요';
  }
  END_POINT_TEXT.textContent = END_LABELS[opt] || '-';
  END_POINT_HINT.textContent = hint;
}

// ---- 팝오버 내부 hint ----
function updateStartPointHints() {
  const endDate = getEndPoint();
  const last = getLastReportTime();
  if (last) {
    LAST_REPORT_DISP.innerHTML = `<span class="latest-badge">마지막</span> ${escapeHtml(formatLocalFull(last))}`;
  } else {
    LAST_REPORT_DISP.textContent = '(아직 저장된 보고 없음)';
  }
  const endZeroEl = $('endZeroDisplay');
  const last24hEl = $('last24hDisplay');
  if (endZeroEl) {
    if (endDate) { const d = new Date(endDate); d.setHours(0, 0, 0, 0); endZeroEl.textContent = `(${formatLocalFull(d)})`; }
    else endZeroEl.textContent = '';
  }
  if (last24hEl) {
    last24hEl.textContent = endDate ? `(${formatLocalFull(new Date(endDate.getTime() - 24 * 3600 * 1000))})` : '';
  }
}

function updateEndPointHints() {
  const fileModEl = $('fileModifiedDisplay');
  const nowEl = $('nowDisplay');
  if (fileModEl) fileModEl.textContent = currentFileMeta ? `(${formatLocalFull(new Date(currentFileMeta.lastModified))})` : '(파일 없음)';
  if (nowEl) nowEl.textContent = `(${formatLocalFull(new Date())})`;
}

function updatePeriodSummary() {
  if (!PERIOD_SUMMARY) return;
  const endDate = getEndPoint();
  const startDate = getStartPoint(endDate);
  PERIOD_SUMMARY.classList.remove('error');
  if (!endDate) {
    PERIOD_SUMMARY.innerHTML = '<em>왼쪽의 업로드 영역에 엑셀 파일을 올려주세요</em>';
    return;
  }
  if (!startDate) {
    PERIOD_SUMMARY.innerHTML = '<em>집계 시작점을 설정해 주세요</em>';
    return;
  }
  if (startDate.getTime() >= endDate.getTime()) {
    PERIOD_SUMMARY.textContent = '시작 시각이 종료 시각보다 늦거나 같습니다.';
    PERIOD_SUMMARY.classList.add('error');
    return;
  }
  const dur = formatDuration(startDate.getTime(), endDate.getTime());
  PERIOD_SUMMARY.innerHTML = `<strong>${escapeHtml(formatLocalFull(startDate))}</strong> ~ <strong>${escapeHtml(formatLocalFull(endDate))}</strong> (${escapeHtml(dur)})`;
}

// ---- 팝오버 토글 (시작점/종료점 상호 배타) ----
let openPeriodPopover = null;
let periodOutsideHandler = null;

function openStartPopover() {
  closeAllPeriodPopovers();
  START_POPOVER.classList.remove('hidden');
  START_CHANGE_BTN.classList.add('is-open');
  START_CHANGE_BTN.textContent = '닫기';
  openPeriodPopover = 'start';
  attachPeriodOutside();
}
function openEndPopover() {
  closeAllPeriodPopovers();
  END_POPOVER.classList.remove('hidden');
  END_CHANGE_BTN.classList.add('is-open');
  END_CHANGE_BTN.textContent = '닫기';
  openPeriodPopover = 'end';
  attachPeriodOutside();
}
function closeAllPeriodPopovers() {
  START_POPOVER.classList.add('hidden');
  END_POPOVER.classList.add('hidden');
  START_CHANGE_BTN.classList.remove('is-open');
  END_CHANGE_BTN.classList.remove('is-open');
  START_CHANGE_BTN.textContent = '변경';
  END_CHANGE_BTN.textContent = '변경';
  openPeriodPopover = null;
  detachPeriodOutside();
}
function attachPeriodOutside() {
  periodOutsideHandler = (e) => {
    const inside = START_POPOVER.contains(e.target) || END_POPOVER.contains(e.target)
                || START_CHANGE_BTN.contains(e.target) || END_CHANGE_BTN.contains(e.target);
    if (inside) return;
    closeAllPeriodPopovers();
  };
  setTimeout(() => {
    document.addEventListener('mousedown', periodOutsideHandler);
    document.addEventListener('keydown', periodEscHandler);
  }, 0);
}
function detachPeriodOutside() {
  if (periodOutsideHandler) {
    document.removeEventListener('mousedown', periodOutsideHandler);
    periodOutsideHandler = null;
  }
  document.removeEventListener('keydown', periodEscHandler);
}
function periodEscHandler(e) { if (e.key === 'Escape') closeAllPeriodPopovers(); }

START_CHANGE_BTN.addEventListener('click', (e) => {
  e.stopPropagation();
  if (openPeriodPopover === 'start') closeAllPeriodPopovers();
  else openStartPopover();
});
END_CHANGE_BTN.addEventListener('click', (e) => {
  e.stopPropagation();
  if (openPeriodPopover === 'end') closeAllPeriodPopovers();
  else openEndPopover();
});

// ---- 직접 설정 인풋 가시성 ----
function toggleCustomStartVisibility() {
  if (!CUSTOM_START_WRAP) return;
  CUSTOM_START_WRAP.classList.toggle('is-active', getSelectedStartOption() === 'custom');
  if (getSelectedStartOption() === 'custom' && !customStartPicker.selectedDates[0]) {
    setTimeout(() => customStartPicker.open(), 0);
  }
}
function toggleCustomEndVisibility() {
  if (!CUSTOM_END_WRAP) return;
  CUSTOM_END_WRAP.classList.toggle('is-active', getSelectedEndOption() === 'custom');
  if (getSelectedEndOption() === 'custom' && !customEndPicker.selectedDates[0]) {
    setTimeout(() => customEndPicker.open(), 0);
  }
}

// ---- 라디오 변경 리스너 ----
document.querySelectorAll('input[name="startPoint"]').forEach((el) => {
  el.addEventListener('change', () => {
    toggleCustomStartVisibility();
    updateStartPointSummary();
    updateStartPointHints();
    updatePeriodSummary();
    if (currentRows) renderReport();
    if (!(getSelectedStartOption() === 'custom' && !customStartPicker.selectedDates[0])) closeAllPeriodPopovers();
  });
});
document.querySelectorAll('input[name="endPoint"]').forEach((el) => {
  el.addEventListener('change', () => {
    toggleCustomEndVisibility();
    updateEndPointSummary();
    updateStartPointSummary();
    updateStartPointHints();
    updateEndPointHints();
    updatePeriodSummary();
    if (currentRows) renderReport();
    if (!(getSelectedEndOption() === 'custom' && !customEndPicker.selectedDates[0])) closeAllPeriodPopovers();
  });
});

function refreshStartPointUI() {
  const last = getLastReportTime();
  const lastRadio = document.querySelector('input[name="startPoint"][value="lastReport"]');
  if (last) lastRadio.disabled = false;
  else {
    lastRadio.disabled = true;
    if (lastRadio.checked) document.querySelector('input[name="startPoint"][value="last24h"]').checked = true;
  }
  if (!document.querySelector('input[name="startPoint"]:checked')) {
    const target = last ? 'lastReport' : 'last24h';
    document.querySelector(`input[name="startPoint"][value="${target}"]`).checked = true;
  }
}

function refreshEndPointUI() {
  // '파일 최종 수정일까지'는 항상 기본/선택 가능. 파일 없으면 요약에 힌트만 표시.
  const fileRadio = document.querySelector('input[name="endPoint"][value="fileModified"]');
  fileRadio.disabled = false;
  if (!document.querySelector('input[name="endPoint"]:checked')) {
    fileRadio.checked = true;
  }
}

function refreshPeriodUI() {
  refreshStartPointUI();
  refreshEndPointUI();
  updateStartPointHints();
  updateEndPointHints();
  updateStartPointSummary();
  updateEndPointSummary();
  updatePeriodSummary();
}

// 'now' 옵션 사용 중일 땐 30초마다 UI 갱신
setInterval(() => {
  if (getSelectedEndOption() === 'now') {
    updateEndPointHints();
    updateEndPointSummary();
    updateStartPointSummary();
    updatePeriodSummary();
  }
}, 30_000);

// ============================================================================
// 상태 메시지
// ============================================================================

function setStatus(message, variant = 'info') {
  STATUS_EL.textContent = message;
  STATUS_EL.className = `status is-${variant}`;
  STATUS_EL.classList.remove('hidden');
}


// ============================================================================
// 파일 업로드 (드롭존/파일 선택)
// ============================================================================

DROPZONE.addEventListener('click', () => { if (!isBusy) FILE_INPUT.click(); });
DROPZONE.addEventListener('keydown', (e) => {
  if (isBusy) return;
  if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); FILE_INPUT.click(); }
});
FILE_INPUT.addEventListener('change', (e) => {
  if (isBusy) return;
  if (e.target.files && e.target.files[0]) handleFile(e.target.files[0]);
});
['dragenter', 'dragover'].forEach((evt) => {
  DROPZONE.addEventListener(evt, (e) => {
    e.preventDefault(); e.stopPropagation();
    if (!isBusy) DROPZONE.classList.add('is-dragging');
  });
});
['dragleave', 'drop'].forEach((evt) => {
  DROPZONE.addEventListener(evt, (e) => { e.preventDefault(); e.stopPropagation(); DROPZONE.classList.remove('is-dragging'); });
});
DROPZONE.addEventListener('drop', (e) => {
  if (isBusy) return;
  const file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
  if (file) handleFile(file);
});

// 클립보드 붙여넣기 지원 (Finder에서 파일 복사 → Cmd+V)
document.addEventListener('paste', (e) => {
  if (isBusy) return;
  if (!e.clipboardData) return;
  const files = Array.from(e.clipboardData.files || []);
  const xlsx = files.find((f) => /\.xlsx?$/i.test(f.name) || /spreadsheet|excel/i.test(f.type));
  if (xlsx) {
    e.preventDefault();
    handleFile(xlsx);
  } else if (files.length > 0) {
    setStatus('붙여넣기된 파일이 엑셀(.xlsx/.xls)이 아닙니다.', 'error');
  }
});

async function handleFile(file) {
  if (isBusy) return;
  const sizeMB = file.size / 1024 / 1024;
  if (sizeMB > MAX_FILE_SIZE_MB) {
    setStatus(`파일이 너무 큽니다 (${sizeMB.toFixed(1)}MB). 최대 ${MAX_FILE_SIZE_MB}MB까지 처리할 수 있어요.`, 'error');
    return;
  }
  setBusy(true);
  try {
    const sizeMBstr = sizeMB.toFixed(1);
    setStatus(`① 파일 읽는 중… (${file.name}, ${sizeMBstr}MB)`, 'info');
    await yieldToUI();
    const buffer = await file.arrayBuffer();

    setStatus(`② 엑셀 분석 중… (${sizeMBstr}MB · 10~30초 소요 가능, 이 동안 다른 액션은 잠시 막혀요)`, 'info');
    await yieldToUI();
    // 파일 메타 저장 (기본 종료점이 '파일 최종 수정일까지'로 이미 맞춰짐)
    currentFileMeta = { name: file.name, lastModified: file.lastModified };
    refreshPeriodUI();
    await parseAndRender(buffer);
    showFileLoaded(file);
  } catch (err) {
    console.error(err);
    setStatus(`파일을 읽을 수 없습니다: ${err.message}`, 'error');
  } finally {
    setBusy(false);
  }
}

async function parseAndRender(buffer) {
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: false });
  if (!rows.length) { setStatus('엑셀에 데이터가 없습니다.', 'error'); return; }
  currentRows = rows;
  await yieldToUI();
  renderReport();
}

// ============================================================================
// 파일 제거 버튼 — 업로드된 파일 정보 초기화
// ============================================================================

REMOVE_FILE_BTN.addEventListener('click', () => {
  if (isBusy) return;
  const hasReport = !!currentReport;
  if (hasReport) {
    if (!confirm('파일을 제거하면 현재 생성된 보고 문장도 함께 초기화됩니다. 진행할까요?')) return;
  }
  currentRows = null;
  currentReport = null;
  currentFileMeta = null;
  FILE_LOADED.classList.add('hidden');
  DROPZONE.classList.remove('hidden');
  FILE_INPUT.value = '';
  refreshPeriodUI();
  resetResultToEmpty();
  setStatus('파일이 제거되었습니다.', 'info');
});

function resetResultToEmpty() {
  RESULT_SECTION.classList.add('is-empty');
  RESULT_TEXT.textContent = '엑셀 파일을 업로드하고 집계 기간을 설정하면 여기에 일간 보고 문장이 생성됩니다.';
  DEBUG_AREA.innerHTML = '';
  SAVE_RECORD_BTN.disabled = true;
  COPY_BTN.disabled = true;
}

function showFileLoaded(file) {
  LOADED_FILE_NAME.textContent = file.name;
  const sizeMB = (file.size / 1024 / 1024).toFixed(1);
  let metaText = `${sizeMB}MB`;
  if (file.lastModified) {
    metaText += ` · 최종 수정일 ${formatLocalFull(new Date(file.lastModified))}`;
  }
  LOADED_FILE_META.textContent = metaText;
  DROPZONE.classList.add('hidden');
  FILE_LOADED.classList.remove('hidden');
}

// ============================================================================
// 엑셀 → 정규화
// ============================================================================

function pickField(row, candidates) {
  for (const key of candidates) {
    if (row[key] !== undefined && row[key] !== null && row[key] !== '') return row[key];
  }
  return null;
}

function normalizeRow(row) {
  const rawPhone = pickField(row, COLUMN_MAP.phone);
  return {
    organization: pickField(row, COLUMN_MAP.organization),
    department:   pickField(row, COLUMN_MAP.department),
    name:         pickField(row, COLUMN_MAP.name),
    title:        pickField(row, COLUMN_MAP.title),
    phoneTail:    extractPhoneTail(rawPhone),
    registeredAt: parseDateTime(pickField(row, COLUMN_MAP.registrationTime)),
  };
}

function extractPhoneTail(v) {
  if (!v) return '';
  const digits = String(v).replace(/\D/g, '');
  return digits.length >= 4 ? digits.slice(-4) : digits;
}

function parseDateTime(value) {
  if (value === null || value === undefined || value === '') return null;
  if (value instanceof Date && !isNaN(value)) return value;
  if (typeof value === 'number') {
    const ms = Math.round((value - 25569) * 86400 * 1000);
    const d = new Date(ms);
    if (!isNaN(d)) return d;
  }
  if (typeof value === 'string') {
    const trimmed = value.trim();
    const m = trimmed.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
    if (m) {
      const [, y, mo, d, hh, mm, ss] = m;
      return new Date(Number(y), Number(mo) - 1, Number(d), Number(hh || 0), Number(mm || 0), Number(ss || 0));
    }
    const parsed = new Date(trimmed);
    if (!isNaN(parsed)) return parsed;
  }
  return null;
}

// ============================================================================
// 대표자 선정
// ============================================================================

function rankIndex(title, organization) {
  const category = window.resolveRankCategory(organization);
  const list = window.RANK_PRIORITY[category] || window.RANK_PRIORITY['일반'];
  if (!title) return Number.POSITIVE_INFINITY;
  const t = String(title).trim();
  const idx = list.indexOf(t);
  if (idx !== -1) return idx;
  for (let i = 0; i < list.length; i++) if (t.includes(list[i])) return i;
  return Number.POSITIVE_INFINITY;
}

function pickRepresentative(group) {
  const organization = group[0].organization;
  const titled = group.filter((m) => m.title && String(m.title).trim());
  const pool = titled.length > 0 ? titled : group;

  let bestByRank = null, bestIdx = Number.POSITIVE_INFINITY;
  for (const m of pool) {
    const r = rankIndex(m.title, organization);
    if (r < bestIdx) { bestByRank = m; bestIdx = r; }
  }
  if (bestByRank && bestIdx !== Number.POSITIVE_INFINITY) return bestByRank;

  let latest = pool[0];
  for (let i = 1; i < pool.length; i++) {
    const c = pool[i];
    if (!c.registeredAt) continue;
    if (!latest.registeredAt || c.registeredAt.getTime() > latest.registeredAt.getTime()) latest = c;
  }
  return latest;
}

function groupBy(rows, key) {
  const map = new Map();
  for (const r of rows) {
    const k = r[key] || '(기관 미상)';
    if (!map.has(k)) map.set(k, []);
    map.get(k).push(r);
  }
  return map;
}

// ============================================================================
// 보고서 생성
// ============================================================================

function buildReport(rows, startDate, endDate) {
  const normalized = rows.map(normalizeRow);
  const startMs = startDate.getTime();
  const endMs = endDate.getTime();

  const newRows = normalized.filter((r) =>
    r.registeredAt && r.registeredAt.getTime() > startMs && r.registeredAt.getTime() <= endMs
  );
  const snapshotRows = normalized.filter((r) =>
    r.registeredAt && r.registeredAt.getTime() <= endMs
  );
  const afterByOrg = groupBy(snapshotRows, 'organization');

  if (newRows.length === 0) {
    return { text: `▲ ${REPORT_HEADER_LABEL} (집계 기간 내 신규 가입자 없음)`, entries: [], newCount: 0 };
  }

  const newByOrg = groupBy(newRows, 'organization');
  const entries = [];
  for (const [org, members] of newByOrg) {
    const rep = pickRepresentative(members);
    const after = (afterByOrg.get(org) || []).length;
    const newCount = members.length;
    const before = after - newCount;
    const repFull = [rep.department, rep.name, rep.title].filter(Boolean).join(' ');
    const countClause = newCount > 1 ? `등 ${newCount}명` : `${newCount}명`;
    const body = [org, repFull, countClause].filter(Boolean).join(' ');
    entries.push({
      line: `- ${body}(${before}명 -> ${after}명)`,
      organization: org, department: rep.department, name: rep.name, title: rep.title,
      newCount, before, after,
      members,  // 이 기관의 신규 가입자 전체 (대표자 포함)
      representative: rep,
    });
  }

  // 총 인원(after) 많은 기관 순으로 정렬. 동률이면 신규 많은 순 → 기관명 순.
  entries.sort((a, b) => (b.after - a.after) || (b.newCount - a.newCount) || a.organization.localeCompare(b.organization, 'ko'));

  return {
    text: `▲ ${REPORT_HEADER_LABEL}(${entries.length}개처)\n${entries.map((e) => e.line).join('\n')}`,
    entries,
    newCount: newRows.length,
  };
}

function renderReport() {
  if (!currentRows) return;
  const endDate = getEndPoint();
  if (!endDate) { setStatus('집계 종료점을 설정해 주세요.', 'error'); return; }
  const startDate = getStartPoint(endDate);
  if (!startDate) { setStatus('집계 시작점을 설정해 주세요.', 'error'); return; }
  if (startDate.getTime() >= endDate.getTime()) {
    setStatus('시작 시각이 종료 시각보다 늦거나 같습니다.', 'error');
    return;
  }
  const report = buildReport(currentRows, startDate, endDate);
  currentReport = { ...report, startDate, endDate };

  RESULT_SECTION.classList.remove('is-empty');
  RESULT_TEXT.textContent = report.text;
  SAVE_RECORD_BTN.disabled = false;
  COPY_BTN.disabled = false;
  renderDebugTable(report.entries);

  const windowLabel = `${formatLocalFull(startDate)} ~ ${formatLocalFull(endDate)}`;
  const duration = formatDuration(startDate.getTime(), endDate.getTime());
  if (report.newCount === 0) {
    setStatus(`집계 기간: ${windowLabel} (${duration}), 신규 없음 (전체 ${currentRows.length}행)`, 'info');
  } else {
    setStatus(`집계 기간: ${windowLabel} (${duration}), 신규 ${report.newCount}명 / ${report.entries.length}개 기관 집계 완료`, 'success');
  }
}

function renderDebugTable(entries) {
  if (!entries.length) { DEBUG_AREA.innerHTML = ''; return; }
  const rowsHtml = entries.flatMap((e) => {
    return e.members.map((m, i) => {
      const isFirst = i === 0;
      const isRep = (m === e.representative);
      const tsLabel = m.registeredAt ? formatLocalFull(m.registeredAt) : '';
      const orgCell = isFirst
        ? `<td rowspan="${e.members.length}" class="org-cell"><strong>${escapeHtml(e.organization)}</strong><br/><small>${e.before} → ${e.after}</small></td>`
        : '';
      return `
        <tr>
          ${orgCell}
          <td>${escapeHtml(m.department || '')}</td>
          <td>${escapeHtml(m.name || '')}${isRep ? ' <span class="rep-badge">대표</span>' : ''}</td>
          <td>${escapeHtml(m.title || '')}</td>
          <td class="mono-cell">${escapeHtml(m.phoneTail || '')}</td>
          <td class="mono-cell">${escapeHtml(tsLabel)}</td>
        </tr>`;
    });
  }).join('');
  DEBUG_AREA.innerHTML = `
    <table class="debug-table">
      <thead><tr><th>기관 (before → after)</th><th>부서</th><th>이름</th><th>직책</th><th>전화번호</th><th>가입 시각</th></tr></thead>
      <tbody>${rowsHtml}</tbody>
    </table>`;
}

// ============================================================================
// 복사 / 기록 저장 / 기록 렌더
// ============================================================================

COPY_BTN.addEventListener('click', async () => {
  try {
    await navigator.clipboard.writeText(RESULT_TEXT.textContent);
    COPY_BTN.textContent = '복사됨 ✓';
    COPY_BTN.classList.add('copied');
    setTimeout(() => { COPY_BTN.textContent = '복사'; COPY_BTN.classList.remove('copied'); }, 1500);
  } catch {
    setStatus('클립보드 복사에 실패했습니다. 수동으로 드래그하여 복사하세요.', 'error');
  }
});

// -------- 보고 기록 저장 팝오버 --------

let popoverOutsideHandler = null;

function openSavePopover() {
  if (!currentReport) { setStatus('저장할 보고가 없습니다.', 'error'); return; }
  const lastAuthor = localStorage.getItem(LAST_AUTHOR_KEY) || '';
  POPOVER_AUTHOR.value = lastAuthor;
  POPOVER_HINT.innerHTML = `기준 시각(<strong>${formatLocalFull(currentReport.endDate)}</strong>)으로 저장됩니다.`;
  SAVE_POPOVER.classList.remove('hidden');
  SAVE_RECORD_BTN.classList.add('is-pressed');
  // 다음 tick에 포커스 (팝오버 열자마자 outside-click 판정을 피하기 위해)
  setTimeout(() => POPOVER_AUTHOR.focus(), 0);

  // 바깥 클릭 / Esc 닫기
  popoverOutsideHandler = (e) => {
    if (SAVE_POPOVER.contains(e.target) || SAVE_RECORD_BTN.contains(e.target)) return;
    closeSavePopover();
  };
  document.addEventListener('mousedown', popoverOutsideHandler);
  document.addEventListener('keydown', popoverEscHandler);
}

function closeSavePopover() {
  SAVE_POPOVER.classList.add('hidden');
  SAVE_RECORD_BTN.classList.remove('is-pressed');
  if (popoverOutsideHandler) {
    document.removeEventListener('mousedown', popoverOutsideHandler);
    popoverOutsideHandler = null;
  }
  document.removeEventListener('keydown', popoverEscHandler);
}

function popoverEscHandler(e) {
  if (e.key === 'Escape') closeSavePopover();
}

async function confirmSavePopover() {
  const author = (POPOVER_AUTHOR.value || '').trim();
  if (!author) {
    POPOVER_AUTHOR.focus();
    POPOVER_AUTHOR.style.borderColor = '#b32020';
    return;
  }
  POPOVER_AUTHOR.style.borderColor = '';
  localStorage.setItem(LAST_AUTHOR_KEY, author);
  const entry = {
    // createdAt = 기준 시각(endDate). "마지막 보고 시각부터" 옵션이 이 값을 읽는다.
    createdAt: currentReport.endDate.toISOString(),
    savedAt: new Date().toISOString(),
    author,
    startAt: currentReport.startDate.toISOString(),
    endAt: currentReport.endDate.toISOString(),
    reportText: currentReport.text,
    newCount: currentReport.newCount,
    orgCount: currentReport.entries.length,
  };
  POPOVER_CONFIRM.disabled = true;
  POPOVER_CONFIRM.textContent = '저장 중…';
  try {
    await saveHistoryEntry(entry);
    renderHistoryList();
    refreshPeriodUI();
    closeSavePopover();
    setStatus(`저장됐습니다. (${entry.author} · 기준 ${formatLocalFull(new Date(entry.createdAt))})`, 'success');
  } finally {
    POPOVER_CONFIRM.disabled = false;
    POPOVER_CONFIRM.textContent = '저장';
  }
}

SAVE_RECORD_BTN.addEventListener('click', (e) => {
  e.stopPropagation();
  if (SAVE_POPOVER.classList.contains('hidden')) openSavePopover();
  else closeSavePopover();
});
POPOVER_CANCEL.addEventListener('click', closeSavePopover);
POPOVER_CONFIRM.addEventListener('click', confirmSavePopover);
POPOVER_AUTHOR.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') { e.preventDefault(); confirmSavePopover(); }
});


let historyShowAll = false;

function renderHistoryList() {
  const list = _historyCache;
  if (list.length === 0) {
    HISTORY_LIST.innerHTML = '<p class="empty">아직 저장된 기록이 없습니다.</p>';
    return;
  }
  const visible = historyShowAll ? list : list.slice(0, 2);
  const itemsHtml = visible.map((e, idx) => {
    const winStart = formatLocalFull(new Date(e.startAt));
    const winEnd = formatLocalFull(new Date(e.endAt));
    const duration = formatDuration(new Date(e.startAt).getTime(), new Date(e.endAt).getTime());
    const savedTs = e.savedAt ? formatLocalFull(new Date(e.savedAt)) : '';
    const savedFooter = savedTs ? `<div class="history-footer">이 보고는 ${escapeHtml(savedTs)}에 저장되었습니다.</div>` : '';
    const winEndHtml = escapeHtml(winEnd);
    const latestBadge = idx === 0 ? '<span class="latest-badge">마지막</span>' : '';
    return `
      <div class="history-item${idx === 0 ? ' is-latest' : ''}">
        <div class="history-meta">
          ${latestBadge}<strong>${escapeHtml(e.author)}</strong> 집계 기간: ${escapeHtml(winStart)} ~ ${winEndHtml} (${escapeHtml(duration)}), ${e.newCount}명 / ${e.orgCount}개 기관
        </div>
        <details class="history-toggle">
          <summary></summary>
          <pre class="history-text">${escapeHtml(e.reportText)}</pre>
        </details>
        ${savedFooter}
      </div>`;
  }).join('');

  let moreBtn = '';
  if (list.length > 2) {
    const label = historyShowAll ? '접기' : `전체 보기 (${list.length - 2}개 더)`;
    moreBtn = `<button id="historyShowAllBtn" class="history-show-all-btn" type="button">${label}</button>`;
  }
  HISTORY_LIST.innerHTML = itemsHtml + moreBtn;
  const btn = $('historyShowAllBtn');
  if (btn) btn.addEventListener('click', () => { historyShowAll = !historyShowAll; renderHistoryList(); });
  return;
}

// ============================================================================
// 이벤트 바인딩 / 초기화
// ============================================================================

document.querySelectorAll('input[name="duration"]').forEach((el) => {
  el.addEventListener('change', () => {
    toggleCustomStartVisibility();
    updateStartPointSummary();
    updatePeriodSummary();
    if (currentRows) renderReport();
    // 직접 설정은 날짜 선택 전까지 펼친 상태 유지 (아래 customStartPicker.onChange에서 닫음)
    const opt = getSelectedDurationOption();
    if (!(opt === 'custom' && !customStartPicker.selectedDates[0])) {
      collapseDuration();
    }
  });
});

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// 초기 렌더 — API에서 기록 불러온 뒤 UI 업데이트
(async () => {
  await loadHistory();
  refreshPeriodUI();
  toggleCustomStartVisibility();
  toggleCustomEndVisibility();
  renderHistoryList();
})();
