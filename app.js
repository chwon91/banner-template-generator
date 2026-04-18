/* ============================================================
   그립 배너 템플릿 자동 생성기
   ============================================================ */

'use strict';

// ── 상수 ──────────────────────────────────────────────────
const MONTH_NAMES = [
  'January','February','March','April','May','June',
  'July','August','September','October','November','December',
];
const DAY_NAMES_KO = ['일','월','화','수','목','금','토'];

const URL_RULES = [
  { prefix: 'EP_',  label: '이벤트 페이지', example: 'gripshow://openUrl?url=...&close=false&theme=black&share=true&login=false' },
  { prefix: 'LP_',  label: '라이브예고',    example: 'gripshow://live/{liveId}' },
  { prefix: 'ST_',  label: '소식/스토리',   example: 'gripshow://story/{storyId}' },
  { prefix: '쇼츠_', label: '쇼츠',         example: 'gripshow://shorts/{ID}', note: 'URL 없이 제목 마지막 _ 세그먼트에서 ID 추출' },
];

// ── 전역 상태 ──────────────────────────────────────────────
const state = {
  // 날짜 선택
  selectedDate: null,   // { year, month(0-indexed), day }
  calViewYear: 0,
  calViewMonth: 0,

  // 탭1: 라이브배너
  userMap: {},
  metabaseLoaded: false,
  parsedRows: [],
  mainTitles: {},

  // 탭2: 영상 배너
  videoRows: [],

  // 탭3: 이미지 배너
  imageRows: [],
};

// ── 유틸 ──────────────────────────────────────────────────
function pad2(n) {
  return String(n).padStart(2, '0');
}

function showAlert(containerId, message, type = 'warning') {
  const el = document.getElementById(containerId);
  if (el) el.innerHTML = `<div class="alert alert-${type}">${message}</div>`;
}

function clearAlert(containerId) {
  const el = document.getElementById(containerId);
  if (el) el.innerHTML = '';
}

function escHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function getKSTToday() {
  const now = new Date();
  const utcMs = now.getTime() + now.getTimezoneOffset() * 60000;
  const kst = new Date(utcMs + 9 * 3600000);
  return { year: kst.getFullYear(), month: kst.getMonth(), day: kst.getDate() };
}

// ── 캘린더 ────────────────────────────────────────────────
function renderCalendar() {
  const { calViewYear: year, calViewMonth: month } = state;
  const kst = getKSTToday();
  const firstDow = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();

  let html = `
    <div class="cal-wrapper">
      <div class="cal-header">
        <button id="calPrev" class="cal-nav">&#8249;</button>
        <span class="cal-month-year">${MONTH_NAMES[month]} ${year}</span>
        <button id="calNext" class="cal-nav">&#8250;</button>
      </div>
      <div class="cal-grid">
        <div class="cal-weekday cal-wd-sun">Su</div>
        <div class="cal-weekday">Mo</div>
        <div class="cal-weekday">Tu</div>
        <div class="cal-weekday">We</div>
        <div class="cal-weekday">Th</div>
        <div class="cal-weekday">Fr</div>
        <div class="cal-weekday cal-wd-sat">Sa</div>
  `;

  for (let i = 0; i < firstDow; i++) {
    html += `<div class="cal-day cal-day--empty"></div>`;
  }

  for (let d = 1; d <= daysInMonth; d++) {
    const dow = (firstDow + d - 1) % 7;
    const isToday = year === kst.year && month === kst.month && d === kst.day;
    const isSel = state.selectedDate &&
      year === state.selectedDate.year &&
      month === state.selectedDate.month &&
      d === state.selectedDate.day;

    const classes = [
      'cal-day',
      dow === 0 ? 'cal-day--sun' : '',
      dow === 6 ? 'cal-day--sat' : '',
      isToday  ? 'cal-day--today' : '',
      isSel    ? 'cal-day--selected' : '',
    ].filter(Boolean).join(' ');

    html += `<div class="${classes}" data-day="${d}">${d}</div>`;
  }

  html += `</div>`; // .cal-grid

  if (state.selectedDate) {
    const { year: sy, month: sm, day: sd } = state.selectedDate;
    const dow = new Date(sy, sm, sd).getDay();
    html += `
      <div class="cal-selected-display">
        선택: <strong>${sy}.${pad2(sm + 1)}.${pad2(sd)} (${DAY_NAMES_KO[dow]})</strong>
      </div>
    `;
  }

  html += `</div>`; // .cal-wrapper

  const container = document.getElementById('calContainer');
  container.innerHTML = html;

  container.querySelector('#calPrev').addEventListener('click', () => {
    state.calViewMonth--;
    if (state.calViewMonth < 0) { state.calViewMonth = 11; state.calViewYear--; }
    renderCalendar();
  });

  container.querySelector('#calNext').addEventListener('click', () => {
    state.calViewMonth++;
    if (state.calViewMonth > 11) { state.calViewMonth = 0; state.calViewYear++; }
    renderCalendar();
  });

  container.querySelectorAll('.cal-day[data-day]').forEach(cell => {
    cell.addEventListener('click', () => {
      state.selectedDate = {
        year: state.calViewYear,
        month: state.calViewMonth,
        day: parseInt(cell.dataset.day, 10),
      };
      renderCalendar();
    });
  });
}

// ── 편성표 파싱 ────────────────────────────────────────────
function parseScheduleGrid(text, selectedDate) {
  const lines = text.split('\n').map(l => l.trimEnd()).filter(l => l.trim() !== '');
  if (lines.length === 0) return null;

  const VALID_POS = new Set(['2', '3', '4', '5', '6']);
  let posMap = {};   // { colIndex(number): posString }
  let headerIdx = -1;

  for (let i = 0; i < lines.length; i++) {
    const cells = lines[i].split('\t');
    const found = [];
    cells.forEach((c, idx) => {
      if (VALID_POS.has(c.trim())) found.push({ idx, pos: c.trim() });
    });
    if (found.length >= 2) {
      headerIdx = i;
      found.forEach(({ idx, pos }) => { posMap[idx] = pos; });
      break;
    }
  }

  if (headerIdx === -1) return null;

  const dateStr = `${selectedDate.month + 1}/${selectedDate.day}`;
  const lbLines = [];

  for (let i = headerIdx + 1; i < lines.length; i++) {
    const cells = lines[i].split('\t');
    const timeRaw = (cells[0] ?? '').trim();
    if (!/^\d{1,2}:\d{2}$/.test(timeRaw)) continue;

    const [hStr, mStr] = timeRaw.split(':');
    const time = `${pad2(parseInt(hStr, 10))}:${pad2(parseInt(mStr, 10))}`;

    Object.entries(posMap).forEach(([colIdxStr, pos]) => {
      const gripper = (cells[parseInt(colIdxStr)] ?? '').trim();
      if (gripper) {
        lbLines.push(`LB_${pos}_${dateStr} ${time} ${gripper}`);
      }
    });
  }

  // 시간 오름차순, 동일 시간 내 위치 오름차순 정렬
  lbLines.sort((a, b) => {
    const parse = s => {
      const m = s.match(/^LB_(\d+)_\S+\s+(\d+:\d+)/);
      return m ? { pos: +m[1], time: m[2] } : { pos: 9, time: '99:99' };
    };
    const pa = parse(a), pb = parse(b);
    if (pa.time < pb.time) return -1;
    if (pa.time > pb.time) return 1;
    return pa.pos - pb.pos;
  });

  return lbLines;
}

// ── LB_ 파싱 ──────────────────────────────────────────────
function parseLB(line, year) {
  line = line.trim();
  if (!line) return null;

  const regex = /^LB_(\d+)_((\d{1,2})\/(\d{1,2}))\s+(\d{1,2}:\d{2})\s+(.+)$/;
  const m = line.match(regex);

  if (!m) return { raw: line, parseError: true };

  const position    = m[1];
  const dateShort   = m[2];
  const month       = pad2(parseInt(m[3], 10));
  const day         = pad2(parseInt(m[4], 10));
  const dateFull    = `${year}.${month}.${day}`;
  const timeParts   = m[5].split(':');
  const hh          = pad2(parseInt(timeParts[0], 10));
  const mm          = pad2(parseInt(timeParts[1], 10));
  const time          = `${hh}:${mm}`;
  const timeFormatted = `${hh}:${mm}:00`;
  const gripperName = m[6].trim();

  return { raw: line, position, dateShort, dateFull, time, timeFormatted, gripperName, parseError: false };
}

// ── USER_ID 조회 ───────────────────────────────────────────
function lookupUserId(gripperName, userMap) {
  return userMap[gripperName] ?? null;
}

// ── URL 변환 ───────────────────────────────────────────────
function convertUrl(title, originalUrl) {
  if (!title || !originalUrl) return { scheme: '', error: '제목 또는 URL 없음' };

  const prefix = title.split('_')[0].toUpperCase();
  const url = originalUrl.trim();

  if (prefix === 'EP') {
    const encoded = encodeURIComponent(url);
    return { scheme: `gripshow://openUrl?url=${encoded}&close=false&theme=black&share=true&login=false`, error: null };
  }
  if (prefix === 'LP') {
    const liveMatch = url.match(/grip\.show\/live\/([^/?#]+)/);
    if (!liveMatch) return { scheme: '', error: 'LP URL에서 live ID를 추출하지 못했습니다.' };
    return { scheme: `gripshow://live/${liveMatch[1]}`, error: null };
  }
  if (prefix === 'ST') {
    const storyMatch = url.match(/grip\.show\/story\/([^/?#]+)/);
    if (!storyMatch) return { scheme: '', error: 'ST URL에서 story ID를 추출하지 못했습니다.' };
    return { scheme: `gripshow://story/${storyMatch[1]}`, error: null };
  }
  if (prefix === '쇼츠') {
    const urlMatch = url && url.match(/grip\.show\/shorts\/([^/?#]+)/);
    if (urlMatch) return { scheme: `gripshow://shorts/${urlMatch[1]}`, error: null };
    const parts = title.split('_');
    const id = parts[parts.length - 1]?.trim();
    if (!id) return { scheme: '', error: '쇼츠 ID를 추출하지 못했습니다. (형식: 쇼츠_그리퍼명_ID)' };
    return { scheme: `gripshow://shorts/${id}`, error: null };
  }

  return { scheme: '', error: `알 수 없는 말머리: ${prefix || '(없음)'}` };
}

// ── 연속 방송 중복 제거 ────────────────────────────────────
function timeToMinutes(hhmm) {
  const [h, m] = hhmm.split(':').map(Number);
  return h * 60 + m;
}

function toHHMMSS(totalSecs) {
  return `${pad2(Math.floor(totalSecs / 3600) % 24)}:${pad2(Math.floor((totalSecs % 3600) / 60))}:${pad2(totalSecs % 60)}`;
}

function applyConsecutiveDedup(rows) {
  const runState = {};
  const result = [];

  rows.forEach(row => {
    if (row.parseError || !['2','3','4','5','6'].includes(row.position)) {
      result.push(row);
      return;
    }
    const key = `${row.position}|${row.gripperName}|${row.dateFull}`;
    const run = runState[key];

    if (run && timeToMinutes(row.time) - timeToMinutes(run.lastTime) === 30) {
      run.lastTime = row.time;
      run.keptRow.runEndTime = toHHMMSS(timeToMinutes(row.time) * 60 + 29 * 60 + 59);
    } else {
      row.runEndTime = null;
      runState[key] = { keptRow: row, lastTime: row.time };
      result.push(row);
    }
  });

  return result;
}

// ── SheetJS 날짜/숫자 자동변환 방지 ───────────────────────
function forceStringCells(ws, aoaData) {
  aoaData.forEach((row, r) => {
    row.forEach((val, c) => {
      const cellAddr = XLSX.utils.encode_cell({ r, c });
      if (ws[cellAddr]) {
        ws[cellAddr].t = 's';
        ws[cellAddr].v = String(val ?? '');
      }
    });
  });
}

// ── URL 규칙 카드 렌더링 ───────────────────────────────────
function renderUrlRulesCard(containerId) {
  const rows = URL_RULES.map(r => `
    <tr>
      <td><code>${escHtml(r.prefix)}</code></td>
      <td>${escHtml(r.label)}</td>
      <td>
        <code>${escHtml(r.example)}</code>
        ${r.note ? `<span class="rule-note">— ${escHtml(r.note)}</span>` : ''}
      </td>
    </tr>
  `).join('');

  document.getElementById(containerId).innerHTML = `
    <div class="card card--info">
      <div class="card-title">URL 변환 규칙 안내</div>
      <table class="info-table">
        <thead>
          <tr><th>말머리</th><th>의미</th><th>변환 예시</th></tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

// ── LB 테이블 렌더링 ───────────────────────────────────────
function renderLBTable() {
  const rows = state.parsedRows;
  const container = document.getElementById('lbPreview');

  if (rows.length === 0) {
    container.innerHTML = '<div class="empty-state">파싱 결과가 없습니다.</div>';
    return;
  }

  const total = rows.length;
  const errors = rows.filter(r => r.parseError || !lookupUserId(r.gripperName, state.userMap)).length;
  const ok = total - errors;

  let html = `
    <div class="parse-stats">
      <span class="stat-chip total">전체 ${total}건</span>
      <span class="stat-chip ok">정상 ${ok}건</span>
      ${errors > 0 ? `<span class="stat-chip error">오류 ${errors}건</span>` : ''}
    </div>
    <div class="table-wrapper">
    <table>
      <thead>
        <tr>
          <th>#</th>
          <th>제목 (원본)</th>
          <th>사용자 코드</th>
          <th>메인타이틀</th>
          <th>날짜</th>
          <th>시작시간</th>
        </tr>
      </thead>
      <tbody>
  `;

  rows.forEach((row, i) => {
    if (row.parseError) {
      html += `
        <tr class="row-error">
          <td>—</td>
          <td colspan="5">
            <span style="color:#dc2626;font-weight:500;">파싱 실패:</span>
            <code style="font-size:12px;margin-left:6px;">${escHtml(row.raw)}</code>
          </td>
        </tr>
      `;
      return;
    }

    const userId = lookupUserId(row.gripperName, state.userMap);
    const userCell = userId
      ? `<span class="user-id-ok">${escHtml(userId)}</span>`
      : `<span class="user-id-missing">❌ 매핑 없음</span>`;
    const rowClass = userId ? '' : 'row-error';

    const isPos2 = row.position === '2';
    const savedTitle = state.mainTitles[i] ?? '';
    const inputHtml = isPos2
      ? `<input class="inline-input" type="text" data-index="${i}" value="${escHtml(savedTitle)}" placeholder="메인타이틀 입력">`
      : `<input class="inline-input" type="text" disabled placeholder="">`;

    const cardNum = row.templateIndex !== undefined && row.templateIndex !== null
      ? `<span class="template-idx">${row.templateIndex + 1}</span>`
      : '';

    html += `
      <tr class="${rowClass}">
        <td style="color:#9ca3af;font-size:12px;">${cardNum}</td>
        <td><code style="font-size:12px;">${escHtml(row.raw)}</code></td>
        <td>${userCell}</td>
        <td>${inputHtml}</td>
        <td>${escHtml(row.dateFull)}</td>
        <td>${escHtml(row.timeFormatted)}</td>
      </tr>
    `;
  });

  html += '</tbody></table></div>';
  container.innerHTML = html;

  container.querySelectorAll('.inline-input[data-index]').forEach(input => {
    input.addEventListener('input', () => {
      state.mainTitles[parseInt(input.dataset.index, 10)] = input.value;
    });
  });
}

// ── 시작 카드 번호 읽기 ────────────────────────────────────
function getStartCardNum() {
  const input = document.getElementById('startCardNumInput');
  if (!input || !input.value.trim()) return null;
  const val = parseInt(input.value.trim(), 10);
  return isNaN(val) ? null : val;
}

// ── LB 등록 템플릿 다운로드 ───────────────────────────────
const LB_HEADER = [
  '콘텐츠 카드 번호',
  '날짜 (yyyy.mm.dd)',
  '방송 시작 시간 (hh:mm:ss)',
  '방송 종료 시간 (hh:mm:ss)',
];

function downloadLBTemplate(pos) {
  if (state.parsedRows.length === 0) {
    showAlert('lbAlert', '먼저 [파싱 & 미리보기]를 실행해주세요.', 'warning');
    return;
  }

  const filteredRows = state.parsedRows.filter(r => !r.parseError && r.position === pos);
  if (filteredRows.length === 0) {
    showAlert('lbAlert', `LB_${pos} 위치의 방송이 없습니다.`, 'warning');
    return;
  }

  const startNum = getStartCardNum();
  const dataRows = filteredRows.map(row => {
    const cardNum = startNum !== null && row.templateIndex !== null
      ? String(startNum + row.templateIndex)
      : '';
    const endTime = row.runEndTime ? `'${row.runEndTime}` : '';
    return [cardNum, `'${row.dateFull}`, `'${row.timeFormatted}`, endTime];
  });

  const wb = XLSX.utils.book_new();
  const wsData = [LB_HEADER, ...dataRows];
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  forceStringCells(ws, wsData);
  XLSX.utils.book_append_sheet(wb, ws, `LB_${pos}`);
  XLSX.writeFile(wb, `라이브배너등록 템플릿_LB${pos}.xlsx`);
  clearAlert('lbAlert');
}

// ── 배너 탭(영상/이미지) 공통 초기화 ──────────────────────
function initBannerTab(tabId, stateKey, downloadBtnId, fileName, sheetName) {
  const addBtn      = document.getElementById(`${tabId}-addRow`);
  const tbody       = document.getElementById(`${tabId}-tbody`);
  const downloadBtn = document.getElementById(downloadBtnId);
  const alertDiv    = `${tabId}-alert`;

  function renderRows() {
    const rows = state[stateKey];
    tbody.innerHTML = '';

    if (rows.length === 0) {
      tbody.innerHTML = `<tr><td colspan="4" class="empty-state">[행 추가] 버튼을 눌러 데이터를 입력하세요.</td></tr>`;
      return;
    }

    rows.forEach((row, i) => {
      const { scheme, error } = convertUrl(row.title, row.url);
      const schemeDisplay = error
        ? `<span class="scheme-cell error">변환 불가: ${escHtml(error)}</span>`
        : `<span class="scheme-cell">${escHtml(scheme)}</span>`;

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>
          <input class="inline-input" type="text" data-field="title" data-index="${i}"
            value="${escHtml(row.title)}" placeholder="예) EP_사전내일_완내스오빠_0410" style="min-width:200px;">
        </td>
        <td>
          <input class="inline-input" type="text" data-field="url" data-index="${i}"
            value="${escHtml(row.url)}" placeholder="https://www.grip.show/..." style="min-width:240px;">
        </td>
        <td>${schemeDisplay}</td>
        <td style="white-space:nowrap;">
          <button class="btn btn-secondary btn-sm" data-copy="${i}" style="margin-right:4px;">복사</button>
          <button class="btn btn-danger btn-sm" data-delete="${i}">삭제</button>
        </td>
      `;
      tbody.appendChild(tr);
    });

    tbody.querySelectorAll('.inline-input').forEach(input => {
      input.addEventListener('input', () => {
        const idx = parseInt(input.dataset.index, 10);
        state[stateKey][idx][input.dataset.field] = input.value;
        const { scheme, error } = convertUrl(state[stateKey][idx].title, state[stateKey][idx].url);
        const schemeCell = input.closest('tr').querySelector('td:nth-child(3)');
        schemeCell.innerHTML = error
          ? `<span class="scheme-cell error">변환 불가: ${escHtml(error)}</span>`
          : `<span class="scheme-cell">${escHtml(scheme)}</span>`;
      });
    });

    tbody.querySelectorAll('[data-copy]').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.dataset.copy, 10);
        const orig = state[stateKey][idx];
        state[stateKey].splice(idx + 1, 0, { title: orig.title, url: orig.url });
        renderRows();
      });
    });

    tbody.querySelectorAll('[data-delete]').forEach(btn => {
      btn.addEventListener('click', () => {
        state[stateKey].splice(parseInt(btn.dataset.delete, 10), 1);
        renderRows();
      });
    });
  }

  addBtn.addEventListener('click', () => {
    state[stateKey].push({ title: '', url: '' });
    renderRows();
    const inputs = tbody.querySelectorAll('input[data-field="title"]');
    if (inputs.length > 0) inputs[inputs.length - 1].focus();
  });

  downloadBtn.addEventListener('click', () => {
    const rows = state[stateKey].filter(r => r.title || r.url);
    if (rows.length === 0) {
      showAlert(alertDiv, '다운로드할 데이터가 없습니다. 행을 추가하고 입력해주세요.', 'warning');
      return;
    }
    const header = ['제목', '스킴'];
    const dataRows = rows.map(row => {
      const { scheme, error } = convertUrl(row.title, row.url);
      return [row.title, error ? `[변환 불가] ${error}` : scheme];
    });
    const wb = XLSX.utils.book_new();
    const wsData = [header, ...dataRows];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    forceStringCells(ws, wsData);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, fileName);
    clearAlert(alertDiv);
  });

  renderRows();
}

// ── 초기화 ────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {

  // 캘린더 초기화 (KST 오늘)
  const kst = getKSTToday();
  state.selectedDate = { year: kst.year, month: kst.month, day: kst.day };
  state.calViewYear  = kst.year;
  state.calViewMonth = kst.month;
  renderCalendar();

  // ① 탭 전환
  const tabBtns = [...document.querySelectorAll('.tab-btn')];
  let activeBtn   = tabBtns.find(b => b.classList.contains('active'));
  let activePanel = document.querySelector('.tab-panel.active');

  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      if (btn === activeBtn) return;
      activeBtn.classList.remove('active');
      activePanel.classList.remove('active');
      btn.classList.add('active');
      const panel = document.getElementById(btn.dataset.tab);
      panel.classList.add('active');
      activeBtn   = btn;
      activePanel = panel;
    });
  });

  // ② 편성표 → LB_ 변환
  document.getElementById('btnConvertSchedule').addEventListener('click', () => {
    clearAlert('lbAlert');

    if (!state.selectedDate) {
      showAlert('lbAlert', '날짜를 먼저 선택해주세요.', 'warning');
      return;
    }

    const text = document.getElementById('scheduleText').value;
    if (!text.trim()) {
      showAlert('lbAlert', '편성표 텍스트를 붙여넣어 주세요.', 'warning');
      return;
    }

    const lines = parseScheduleGrid(text, state.selectedDate);
    if (!lines) {
      showAlert('lbAlert', '편성표 형식을 인식하지 못했습니다. 헤더(2~6)와 시간열이 포함되었는지 확인해주세요.', 'error');
      return;
    }
    if (lines.length === 0) {
      showAlert('lbAlert', '방송 데이터가 없습니다. (모든 셀이 비어있음)', 'warning');
      return;
    }

    document.getElementById('lbText').value = lines.join('\n');
    showAlert('lbAlert', `${lines.length}건의 LB_ 텍스트를 생성했습니다. 확인 후 [파싱 & 미리보기]를 클릭하세요.`, 'info');
  });

  // ③ Metabase xlsx 업로드
  document.getElementById('metabaseFile').addEventListener('change', function (e) {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    const fileNameEl = document.getElementById('metabaseFileName');
    fileNameEl.innerHTML = '<span style="color:#6b7280;">파일 읽는 중...</span>';
    fileNameEl.classList.remove('loaded');

    state.userMap = {};
    state.metabaseLoaded = false;

    const results = [];
    let pending = files.length;

    files.forEach(file => {
      const reader = new FileReader();
      reader.onload = function (ev) {
        try {
          const data = new Uint8Array(ev.target.result);
          const wb   = XLSX.read(data, { type: 'array' });
          const ws   = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

          if (rows.length === 0) {
            results.push({ name: file.name, count: 0, error: '데이터 없음' });
          } else {
            const firstKeys = Object.keys(rows[0]);
            const nameKey = firstKeys.find(k => k.toUpperCase() === 'USER_NAME');
            const idKey   = firstKeys.find(k => k.toUpperCase() === 'USER_ID');

            if (!nameKey || !idKey) {
              results.push({ name: file.name, count: 0, error: 'USER_NAME/USER_ID 컬럼 없음' });
            } else {
              let count = 0;
              rows.forEach(row => {
                const name = String(row[nameKey]).trim();
                const id   = String(row[idKey]).trim();
                if (name && id) { state.userMap[name] = id; count++; }
              });
              results.push({ name: file.name, count, error: null });
            }
          }
        } catch (err) {
          results.push({ name: file.name, count: 0, error: err.message });
        }

        pending--;
        if (pending > 0) return;

        const totalCount = Object.keys(state.userMap).length;
        const hasAnyError = results.some(r => r.error);

        if (totalCount > 0) {
          state.metabaseLoaded = true;
          clearAlert('lbAlert');
        }

        const listHtml = results.map(r =>
          r.error
            ? `<div class="file-item file-item--error">✗ ${escHtml(r.name)} — ${escHtml(r.error)}</div>`
            : `<div class="file-item file-item--ok">✓ ${escHtml(r.name)} (${r.count}명)</div>`
        ).join('');

        fileNameEl.innerHTML = listHtml + `<div class="file-total">총 ${totalCount}명 로드됨</div>`;
        fileNameEl.classList.add('loaded');

        if (hasAnyError && totalCount === 0) {
          showAlert('lbAlert', '파일을 불러오지 못했습니다. 오류 내용을 확인해주세요.', 'error');
        }
      };
      reader.readAsArrayBuffer(file);
    });
  });

  // ④ 파싱 & 미리보기
  document.getElementById('btnParse').addEventListener('click', () => {
    clearAlert('lbAlert');

    if (!state.metabaseLoaded) {
      showAlert('lbAlert', 'Metabase 파일을 먼저 업로드해주세요.', 'warning');
      return;
    }
    if (!state.selectedDate) {
      showAlert('lbAlert', '날짜를 먼저 선택해주세요.', 'warning');
      return;
    }

    const year    = state.selectedDate.year;
    const rawText = document.getElementById('lbText').value;
    const lines   = rawText.split('\n').filter(l => l.trim() !== '');

    if (lines.length === 0) {
      showAlert('lbAlert', 'LB_ 텍스트를 입력해주세요.', 'warning');
      return;
    }

    const deduped = applyConsecutiveDedup(lines.map(line => parseLB(line, year)));

    // templateIndex 부여 (카드 번호 계산용)
    let idx = 0;
    deduped.forEach(row => {
      row.templateIndex = row.parseError ? null : idx++;
    });

    state.parsedRows = deduped;
    state.mainTitles = {};
    renderLBTable();
  });

  // ⑤ 라이브 템플릿 다운로드
  document.getElementById('btnDownloadLive').addEventListener('click', () => {
    if (state.parsedRows.length === 0) {
      showAlert('lbAlert', '먼저 [파싱 & 미리보기]를 실행해주세요.', 'warning');
      return;
    }

    const header = ['제목', '사용자 코드 (그리퍼 코드)', '메인타이틀'];
    const dataRows = [];

    state.parsedRows.forEach((row, i) => {
      if (row.parseError) return;
      dataRows.push([row.raw, lookupUserId(row.gripperName, state.userMap) ?? '', state.mainTitles[i] ?? '']);
    });

    const wb = XLSX.utils.book_new();
    const wsData = [header, ...dataRows];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    forceStringCells(ws, wsData);
    XLSX.utils.book_append_sheet(wb, ws, '라이브 템플릿');
    XLSX.writeFile(wb, '라이브 템플릿.xlsx');
  });

  // ⑥ 라이브배너등록 템플릿 — 위치별 다운로드
  const LB_POSITIONS = ['2', '3', '4', '5', '6'];

  LB_POSITIONS.forEach(pos => {
    document.querySelector(`[data-lb-pos="${pos}"]`).addEventListener('click', () => {
      downloadLBTemplate(pos);
    });
  });

  // ⑦ 전체 다운로드 (LB_2~6)
  document.getElementById('btnDownloadAllLB').addEventListener('click', () => {
    if (state.parsedRows.length === 0) {
      showAlert('lbAlert', '먼저 [파싱 & 미리보기]를 실행해주세요.', 'warning');
      return;
    }

    const availablePos = LB_POSITIONS.filter(pos =>
      state.parsedRows.some(r => !r.parseError && r.position === pos)
    );

    if (availablePos.length === 0) {
      showAlert('lbAlert', '다운로드할 LB 위치 데이터가 없습니다.', 'warning');
      return;
    }

    availablePos.forEach((pos, i) => {
      setTimeout(() => downloadLBTemplate(pos), i * 400);
    });
  });

  // URL 규칙 카드 렌더링 (탭2, 탭3)
  renderUrlRulesCard('url-rules-video');
  renderUrlRulesCard('url-rules-image');

  // 탭2: 영상 배너
  initBannerTab('video', 'videoRows', 'btnDownloadVideo', '영상 템플릿.xlsx', '영상 템플릿');

  // 탭3: 이미지 배너
  initBannerTab('image', 'imageRows', 'btnDownloadImage', '이미지 템플릿.xlsx', '이미지 템플릿');
});
