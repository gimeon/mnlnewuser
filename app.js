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
// 관리자 페이지의 엑셀 다운로드 엔드포인트. 회사망 브라우저 세션이 있어야 접근 가능.
const ADMIN_EXPORT_URL = 'https://mng.yna.co.kr/mng/idgrouplist/exceldownload';
const HISTORY_KEY = 'dailyReportHistory_v1';
const LAST_AUTHOR_KEY = 'dailyReportLastAuthor_v1';
const HISTORY_MAX_ITEMS = 50;
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
const REF_TIME_INPUT    = $('referenceTime');
const RESET_NOW_BTN     = $('resetNowBtn');
const AUTO_DL_BTN       = $('autoDownloadBtn');
const AUTO_DL_HINT      = $('autoDownloadHint');
const HISTORY_LIST      = $('historyList');
const CLEAR_HISTORY_BTN = $('clearHistoryBtn');
const LAST_REPORT_DISP  = $('lastReportDisplay');
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
  // AUTO_DL_BTN은 "(준비중)" 상태로 항상 disabled — setBusy 토글 대상에서 제외
  [SAVE_RECORD_BTN, COPY_BTN, RESET_NOW_BTN, CLEAR_HISTORY_BTN, FILE_INPUT].forEach((el) => { if (el) el.disabled = flag; });
  REF_TIME_INPUT.disabled = flag;
  if (flag) refTimePicker._input.setAttribute('readonly', true);
  else refTimePicker._input.removeAttribute('readonly');
  document.querySelectorAll('input[name="duration"]').forEach((el) => { el.disabled = flag || (el.value === 'lastReport' && !getLastReportTime()); });
}

// 브라우저가 status 메시지를 실제로 페인트할 수 있도록 다음 프레임까지 양보
function yieldToUI() {
  return new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 0)));
}

// ============================================================================
// 시간 포맷 헬퍼
// ============================================================================

const pad = (n) => String(n).padStart(2, '0');

function formatLocalFull(d) {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function formatDuration(startMs, endMs) {
  const diffMin = Math.max(0, Math.floor((endMs - startMs) / 60000));
  const hours = Math.floor(diffMin / 60);
  const minutes = diffMin % 60;
  if (hours === 0) return `${minutes}분`;
  return `${hours}시간 ${minutes}분`;
}

// ============================================================================
// 저장 기록 (localStorage)
// ============================================================================

function loadHistory() {
  try {
    const raw = localStorage.getItem(HISTORY_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch { return []; }
}

function saveHistoryEntry(entry) {
  const list = loadHistory();
  list.unshift(entry);
  if (list.length > HISTORY_MAX_ITEMS) list.length = HISTORY_MAX_ITEMS;
  localStorage.setItem(HISTORY_KEY, JSON.stringify(list));
  return list;
}

function getLastReportTime() {
  const list = loadHistory();
  return list.length > 0 ? new Date(list[0].createdAt) : null;
}

// ============================================================================
// 기준 시각 / 기간 선택
// ============================================================================

const CURRENT_TIME_BADGE = $('currentTimeBadge');
function showCurrentBadge() { CURRENT_TIME_BADGE.classList.remove('hidden'); }
function hideCurrentBadge() { CURRENT_TIME_BADGE.classList.add('hidden'); }

// flatpickr 인스턴스 (yyyy-mm-dd hh:mm 포맷으로 datetime 선택)
const refTimePicker = flatpickr(REF_TIME_INPUT, {
  locale: 'ko',
  enableTime: true,
  time_24hr: true,
  dateFormat: 'Y-m-d H:i',
  allowInput: true,
  minuteIncrement: 10,
  defaultDate: new Date(),
  onChange: () => { hideCurrentBadge(); if (currentRows) renderReport(); },
});

// 달력 하단에 빠른 시각 프리셋 버튼을 주입한다.
refTimePicker.config.onReady.push((dates, str, fp) => {
  if (fp.calendarContainer.querySelector('.fp-time-presets')) return;
  const presets = [
    { label: '지금', apply: (d) => { const n = new Date(); d.setHours(n.getHours(), n.getMinutes(), 0, 0); } },
    { label: '00:00', h: 0,  m: 0 },
    { label: '09:00', h: 9,  m: 0 },
    { label: '13:00', h: 13, m: 0 },
    { label: '16:00', h: 16, m: 0 },
    { label: '18:00', h: 18, m: 0 },
  ];
  const wrap = document.createElement('div');
  wrap.className = 'fp-time-presets';
  presets.forEach((p) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'fp-preset-btn';
    btn.textContent = p.label;
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      const cur = fp.selectedDates[0] ? new Date(fp.selectedDates[0]) : new Date();
      if (p.apply) p.apply(cur);
      else cur.setHours(p.h, p.m, 0, 0);
      fp.setDate(cur, true);
    });
    wrap.appendChild(btn);
  });
  fp.calendarContainer.appendChild(wrap);
});
refTimePicker.config.onReady.forEach((fn) => fn(refTimePicker.selectedDates, '', refTimePicker));

function initReferenceTime() {
  refTimePicker.setDate(new Date(), false);
  showCurrentBadge();
}
// 페이지 로드 시엔 기본값이 현재 시각이므로 배지 표시
showCurrentBadge();

function getReferenceDate() {
  const d = refTimePicker.selectedDates && refTimePicker.selectedDates[0];
  return d ? new Date(d) : new Date();
}

function getSelectedDurationOption() {
  const el = document.querySelector('input[name="duration"]:checked');
  return el ? el.value : null;
}

function computeStartDate(refDate) {
  const opt = getSelectedDurationOption();
  if (opt === 'lastReport') {
    return getLastReportTime();
  }
  if (opt === 'todayMidnight') {
    const d = new Date(refDate);
    d.setHours(0, 0, 0, 0);
    return d;
  }
  if (opt === 'last24h') {
    return new Date(refDate.getTime() - 24 * 3600 * 1000);
  }
  return null;
}

function refreshDurationUI() {
  const last = getLastReportTime();
  const lastReportRadio = document.querySelector('input[name="duration"][value="lastReport"]');
  if (last) {
    lastReportRadio.disabled = false;
    LAST_REPORT_DISP.innerHTML = `(마지막 보고: <span class="accent-latest">${escapeHtml(formatLocalFull(last))}</span>)`;
  } else {
    lastReportRadio.disabled = true;
    LAST_REPORT_DISP.textContent = '(아직 저장된 보고 없음)';
    if (lastReportRadio.checked) {
      document.querySelector('input[name="duration"][value="last24h"]').checked = true;
    }
  }
  // 기본 선택: 기록 있으면 lastReport, 없으면 last24h
  const anyChecked = document.querySelector('input[name="duration"]:checked');
  if (!anyChecked) {
    const target = last ? 'lastReport' : 'last24h';
    document.querySelector(`input[name="duration"][value="${target}"]`).checked = true;
  }
}

// ============================================================================
// 상태 메시지
// ============================================================================

function setStatus(message, variant = 'info') {
  STATUS_EL.textContent = message;
  STATUS_EL.className = `status is-${variant}`;
  STATUS_EL.classList.remove('hidden');
}

function setAutoDlHint(message, variant = '') {
  AUTO_DL_HINT.textContent = message;
  AUTO_DL_HINT.className = `field-hint ${variant}`;
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
    await parseAndRender(buffer);
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
// 자동 다운로드 (관리자 API → fetch → parse)
// ============================================================================

let downloadController = null;

AUTO_DL_BTN.addEventListener('click', async () => {
  if (downloadController) {
    downloadController.abort();
    downloadController = null;
    AUTO_DL_BTN.textContent = '관리자에서 자동 다운로드 & 분석';
    setAutoDlHint('취소됨', '');
    return;
  }
  if (isBusy) return;
  setBusy(true);

  downloadController = new AbortController();
  AUTO_DL_BTN.disabled = false;   // 다운로드 중에는 취소 가능해야 하므로 다시 활성화
  AUTO_DL_BTN.textContent = '취소';
  const startTs = Date.now();
  const interval = setInterval(() => {
    const s = Math.floor((Date.now() - startTs) / 1000);
    setAutoDlHint(`다운로드 진행 중… (${Math.floor(s / 60)}분 ${s % 60}초) · 보통 5분 내외 소요`, '');
  }, 1000);

  try {
    setStatus('관리자 서버에서 엑셀을 받아오는 중입니다. 최대 5~6분이 걸릴 수 있어요.', 'info');
    const res = await fetch(ADMIN_EXPORT_URL, {
      method: 'GET',
      credentials: 'include',
      signal: downloadController.signal,
    });

    if (!res.ok) throw new Error(`HTTP ${res.status} ${res.statusText}`);

    const contentType = res.headers.get('Content-Type') || '';
    if (/text\/html/i.test(contentType)) {
      throw new Error('HTML 응답이 돌아왔습니다. 세션이 만료됐을 수 있어요 — 관리자 페이지에 다시 로그인한 뒤 시도해 주세요.');
    }

    const buffer = await res.arrayBuffer();
    setAutoDlHint(`${(buffer.byteLength / 1024 / 1024).toFixed(1)}MB 받음. 분석 중…`, '');
    await yieldToUI();
    await parseAndRender(buffer);
    setAutoDlHint('자동 다운로드 완료.', 'success');
  } catch (err) {
    console.error(err);
    const msg = (err && err.name === 'AbortError') ? '사용자가 취소했습니다.' : err.message || String(err);
    // CORS로 차단된 경우 TypeError로 나오는 경우가 많음
    const isCors = err && err.name === 'TypeError';
    if (isCors) {
      setAutoDlHint('자동 다운로드 차단됨 — 브라우저 CORS 정책. 아래 "새 탭에서 열기" 링크로 받거나, 업로드 방식을 쓰세요.', 'error');
      setStatus('자동 다운로드 실패 (CORS). 새 탭에서 직접 받으시거나 파일 업로드로 진행해 주세요.', 'error');
    } else {
      setAutoDlHint(`자동 다운로드 실패: ${msg}`, 'error');
      setStatus(`자동 다운로드 실패: ${msg}`, 'error');
    }
  } finally {
    clearInterval(interval);
    downloadController = null;
    AUTO_DL_BTN.textContent = '관리자에서 자동 다운로드 & 분석';
    setBusy(false);
  }
});

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
  const endDate = getReferenceDate();
  const startDate = computeStartDate(endDate);
  if (!startDate) {
    setStatus('집계 기간 시작점을 선택해 주세요.', 'error');
    return;
  }
  if (startDate.getTime() >= endDate.getTime()) {
    setStatus('시작 시각이 기준(종료) 시각보다 이후입니다. 선택을 확인해 주세요.', 'error');
    return;
  }
  const report = buildReport(currentRows, startDate, endDate);
  currentReport = { ...report, startDate, endDate };

  RESULT_SECTION.classList.remove('hidden');
  RESULT_TEXT.textContent = report.text;
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

function confirmSavePopover() {
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
  saveHistoryEntry(entry);
  renderHistoryList();
  refreshDurationUI();
  closeSavePopover();
  setStatus(`저장됐습니다. (${entry.author} · 기준 ${formatLocalFull(new Date(entry.createdAt))})`, 'success');
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

CLEAR_HISTORY_BTN.addEventListener('click', () => {
  if (!confirm('저장된 모든 보고 기록을 삭제할까요? 되돌릴 수 없습니다.')) return;
  localStorage.removeItem(HISTORY_KEY);
  renderHistoryList();
  refreshDurationUI();
});

function renderHistoryList() {
  const list = loadHistory();
  if (list.length === 0) {
    HISTORY_LIST.innerHTML = '<p class="empty">아직 저장된 기록이 없습니다.</p>';
    return;
  }
  HISTORY_LIST.innerHTML = list.map((e, idx) => {
    const winStart = formatLocalFull(new Date(e.startAt));
    const winEnd = formatLocalFull(new Date(e.endAt));
    const duration = formatDuration(new Date(e.startAt).getTime(), new Date(e.endAt).getTime());
    const savedTs = e.savedAt ? formatLocalFull(new Date(e.savedAt)) : '';
    const savedFooter = savedTs ? `<div class="history-footer">이 보고는 ${escapeHtml(savedTs)}에 저장되었습니다.</div>` : '';
    const winEndHtml = idx === 0
      ? `<span class="accent-latest">${escapeHtml(winEnd)}</span>`
      : escapeHtml(winEnd);
    return `
      <div class="history-item">
        <div class="history-meta">
          <strong>${escapeHtml(e.author)}</strong> 집계 기간: ${escapeHtml(winStart)} ~ ${winEndHtml} (${escapeHtml(duration)}), ${e.newCount}명 / ${e.orgCount}개 기관
        </div>
        <details class="history-toggle">
          <summary></summary>
          <pre class="history-text">${escapeHtml(e.reportText)}</pre>
        </details>
        ${savedFooter}
      </div>`;
  }).join('');
}

// ============================================================================
// 이벤트 바인딩 / 초기화
// ============================================================================

RESET_NOW_BTN.addEventListener('click', () => { initReferenceTime(); if (currentRows) renderReport(); });
document.querySelectorAll('input[name="duration"]').forEach((el) => {
  el.addEventListener('change', () => { if (currentRows) renderReport(); });
});

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// 초기 렌더
refreshDurationUI();
renderHistoryList();
