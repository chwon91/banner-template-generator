/* ============================================================
   그립 배너 템플릿 자동 생성기
   ============================================================ */

'use strict';

// ── 전역 상태 ──────────────────────────────────────────────
const state = {
  // 탭1: 라이브배너
  userMap: {},          // { "그리퍼명": "user_id" }
  metabaseLoaded: false,
  parsedRows: [],       // parseLB 결과 배열
  mainTitles: {},       // index -> 메인타이틀 입력값

  // 탭2: 영상 배너
  videoRows: [],        // [{ title, url }]

  // 탭3: 이미지 배너
  imageRows: [],        // [{ title, url }]
};

// ── URL 변환 규칙 데이터 ────────────────────────────────────
const URL_RULES = [
  { prefix: 'EP_',  label: '이벤트 페이지', example: 'gripshow://openUrl?url=...&close=false&theme=black&share=true&login=false' },
  { prefix: 'LP_',  label: '라이브예고',    example: 'gripshow://live/{liveId}' },
  { prefix: 'ST_',  label: '소식/스토리',   example: 'gripshow://story/{storyId}' },
  { prefix: '쇼츠_', label: '쇼츠',         example: 'gripshow://shorts/{ID}', note: 'URL 없이 제목 마지막 _ 세그먼트에서 ID 추출' },
];

// ── 유틸 ──────────────────────────────────────────────────
function pad2(n) {
  return String(n).padStart(2, '0');
}

function showAlert(containerId, message, type = 'warning') {
  const container = document.getElementById(containerId);
  if (!container) return;
  container.innerHTML = `<div class="alert alert-${type}">${message}</div>`;
}

function clearAlert(containerId) {
  const container = document.getElementById(containerId);
  if (container) container.innerHTML = '';
}

function escHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ── 1. LB_ 파싱 ────────────────────────────────────────────
/**
 * parseLB("LB_2_4/10 11:00 완내스오빠", 2026)
 * → { raw, position, dateShort, dateFull, time, timeFormatted, gripperName }
 */
function parseLB(line, year) {
  line = line.trim();
  if (!line) return null;

  // LB_[위치]_[월/일] [시간] [그리퍼명]
  // e.g. LB_2_4/10 11:00 완내스오빠
  const regex = /^LB_(\d+)_((\d{1,2})\/(\d{1,2}))\s+(\d{1,2}:\d{2})\s+(.+)$/;
  const m = line.match(regex);

  if (!m) {
    return { raw: line, parseError: true };
  }

  const position    = m[1];
  const dateShort   = m[2];            // "4/10"
  const month       = pad2(parseInt(m[3], 10));
  const day         = pad2(parseInt(m[4], 10));
  const dateFull    = `${year}.${month}.${day}`;
  // 시간 제로패딩: "8:00" → "08:00", "11:00" → "11:00"
  const timeParts     = m[5].split(':');
  const hh            = pad2(parseInt(timeParts[0], 10));
  const mm            = pad2(parseInt(timeParts[1], 10));
  const time          = `${hh}:${mm}`;
  const timeFormatted = `${hh}:${mm}:00`;
  const gripperName = m[6].trim();

  return {
    raw: line,
    position,
    dateShort,
    dateFull,
    time,
    timeFormatted,
    gripperName,
    parseError: false,
  };
}

// ── 2. USER_ID 조회 ────────────────────────────────────────
function lookupUserId(gripperName, userMap) {
  if (!gripperName) return null;
  return userMap[gripperName] ?? null;
}

// ── 3. URL 변환 ────────────────────────────────────────────
function convertUrl(title, originalUrl) {
  if (!title || !originalUrl) return { scheme: '', error: '제목 또는 URL 없음' };

  const prefix = title.split('_')[0].toUpperCase();
  const url = originalUrl.trim();

  if (prefix === 'EP') {
    // 이벤트 페이지
    const encoded = encodeURIComponent(url);
    return { scheme: `gripshow://openUrl?url=${encoded}&close=false&theme=black&share=true&login=false`, error: null };
  }

  if (prefix === 'LP') {
    // 라이브예고: https://www.grip.show/live/{id}
    const liveMatch = url.match(/grip\.show\/live\/([^/?#]+)/);
    if (!liveMatch) return { scheme: '', error: 'LP URL에서 live ID를 추출하지 못했습니다.' };
    return { scheme: `gripshow://live/${liveMatch[1]}`, error: null };
  }

  if (prefix === 'ST') {
    // 소식/스토리: https://link.grip.show/story/{id}
    const storyMatch = url.match(/grip\.show\/story\/([^/?#]+)/);
    if (!storyMatch) return { scheme: '', error: 'ST URL에서 story ID를 추출하지 못했습니다.' };
    return { scheme: `gripshow://story/${storyMatch[1]}`, error: null };
  }

  if (prefix === '쇼츠') {
    // 쇼츠: URL에서 추출 시도, 없으면 제목 마지막 _ 세그먼트 사용
    // 예) 쇼츠_또진이네_x3n21d03 → gripshow://shorts/x3n21d03
    const urlMatch = url && url.match(/grip\.show\/shorts\/([^/?#]+)/);
    if (urlMatch) return { scheme: `gripshow://shorts/${urlMatch[1]}`, error: null };
    const parts = title.split('_');
    const id = parts[parts.length - 1]?.trim();
    if (!id) return { scheme: '', error: '쇼츠 ID를 추출하지 못했습니다. (형식: 쇼츠_그리퍼명_ID)' };
    return { scheme: `gripshow://shorts/${id}`, error: null };
  }

  return { scheme: '', error: `알 수 없는 말머리: ${prefix || '(없음)'}` };
}

// ── LB_6 중복 제거 (동일 셀러 30분 연속 → 첫 카드만 유지) ─
function timeToMinutes(hhmm) {
  const [h, m] = hhmm.split(':').map(Number);
  return h * 60 + m;
}

function deduplicateLB6(rows) {
  const result = [];
  rows.forEach(row => {
    if (result.length === 0) { result.push(row); return; }
    const prev = result[result.length - 1];
    const isSameSeller = row.gripperName === prev.gripperName;
    const isSameDate   = row.dateFull === prev.dateFull;
    const diff         = timeToMinutes(row.time) - timeToMinutes(prev.time);
    if (isSameSeller && isSameDate && diff === 30) return; // 중복 제거
    result.push(row);
  });
  return result;
}

// 전체 parsedRows에서 LB_6 중복 제거 (파싱 시점에 적용)
// 쌍이 제거된 유지 행에 lb6HasPair = true 마킹
function applyLB6Dedup(rows) {
  const lastKeptTime = {}; // "gripperName|dateFull" → 마지막 유지된 time
  const lastKeptRow  = {}; // "gripperName|dateFull" → 마지막 유지된 row 참조
  return rows.filter(row => {
    if (row.parseError || row.position !== '6') return true;
    const key = `${row.gripperName}|${row.dateFull}`;
    const prevTime = lastKeptTime[key];
    if (prevTime !== undefined && timeToMinutes(row.time) - timeToMinutes(prevTime) === 30) {
      if (lastKeptRow[key]) lastKeptRow[key].lb6HasPair = true; // 앞 행에 플래그
      return false;
    }
    lastKeptTime[key] = row.time;
    lastKeptRow[key]  = row;
    return true;
  });
}

// LB_6 종료시간 계산: 시작시간 + 59분 59초
function calcEndTime(hhmmss) {
  const [h, m, s] = hhmmss.split(':').map(Number);
  const total = h * 3600 + m * 60 + s + 59 * 60 + 59;
  return `${pad2(Math.floor(total / 3600) % 24)}:${pad2(Math.floor((total % 3600) / 60))}:${pad2(total % 60)}`;
}

// ── SheetJS 날짜/숫자 자동변환 방지 ─────────────────────
/**
 * aoa_to_sheet 이후 모든 셀을 명시적으로 's' 타입 지정.
 * 날짜(yyyy.mm.dd), 시간(hh:mm:ss) 등이 숫자로 변환되는 것을 방지.
 */
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
          <tr>
            <th>말머리</th>
            <th>의미</th>
            <th>변환 예시</th>
          </tr>
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

    // 메인타이틀: position === "2" 일 때만 활성화
    const isPos2 = row.position === '2';
    const savedTitle = state.mainTitles[i] ?? '';
    const inputHtml = isPos2
      ? `<input class="inline-input" type="text" data-index="${i}" value="${escHtml(savedTitle)}" placeholder="메인타이틀 입력">`
      : `<input class="inline-input" type="text" disabled placeholder="">`;

    html += `
      <tr class="${rowClass}">
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

  // 메인타이틀 입력값 실시간 저장
  container.querySelectorAll('.inline-input[data-index]').forEach(input => {
    input.addEventListener('input', () => {
      state.mainTitles[parseInt(input.dataset.index, 10)] = input.value;
    });
  });
}

// ── 탭2/3 공통: 배너 행 관리 ──────────────────────────────
function initBannerTab(tabId, stateKey, downloadBtnId, fileName, sheetName) {
  const addBtn      = document.getElementById(`${tabId}-addRow`);
  const tbody       = document.getElementById(`${tabId}-tbody`);
  const downloadBtn = document.getElementById(downloadBtnId);
  const alertDiv    = `${tabId}-alert`;

  function renderRows() {
    const rows = state[stateKey];
    tbody.innerHTML = '';

    if (rows.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="4" class="empty-state">
            [행 추가] 버튼을 눌러 데이터를 입력하세요.
          </td>
        </tr>
      `;
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
          <input class="inline-input" type="text"
            data-field="title" data-index="${i}"
            value="${escHtml(row.title)}"
            placeholder="예) EP_사전내일_완내스오빠_0410"
            style="min-width:200px;">
        </td>
        <td>
          <input class="inline-input" type="text"
            data-field="url" data-index="${i}"
            value="${escHtml(row.url)}"
            placeholder="https://www.grip.show/..."
            style="min-width:240px;">
        </td>
        <td>${schemeDisplay}</td>
        <td style="white-space:nowrap;">
          <button class="btn btn-secondary btn-sm" data-copy="${i}" style="margin-right:4px;">복사</button>
          <button class="btn btn-danger btn-sm" data-delete="${i}">삭제</button>
        </td>
      `;
      tbody.appendChild(tr);
    });

    // 이벤트 바인딩
    tbody.querySelectorAll('.inline-input').forEach(input => {
      input.addEventListener('input', () => {
        const idx   = parseInt(input.dataset.index, 10);
        const field = input.dataset.field;
        state[stateKey][idx][field] = input.value;
        // 스킴 셀만 갱신
        const { scheme, error } = convertUrl(
          state[stateKey][idx].title,
          state[stateKey][idx].url
        );
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
    // 마지막 행 title input에 포커스
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

  // 초기 렌더
  renderRows();
}

// ── 초기화 ────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  // 연도 기본값: 현재 연도
  document.getElementById('yearInput').value = new Date().getFullYear();

  // ① 탭 전환 (activeBtn/Panel 추적으로 최적화)
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

  // ② Metabase xlsx 업로드 — 다중 파일 지원, 결과 합산
  document.getElementById('metabaseFile').addEventListener('change', function (e) {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    const fileNameEl = document.getElementById('metabaseFileName');
    fileNameEl.innerHTML = '<span style="color:#6b7280;">파일 읽는 중...</span>';
    fileNameEl.classList.remove('loaded');

    state.userMap = {};
    state.metabaseLoaded = false;

    const results = [];   // { name, count, error }
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
            const nameKey   = firstKeys.find(k => k.toUpperCase() === 'USER_NAME');
            const idKey     = firstKeys.find(k => k.toUpperCase() === 'USER_ID');

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

        // 모든 파일 처리 완료
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

        fileNameEl.innerHTML = listHtml
          + `<div class="file-total">총 ${totalCount}명 로드됨</div>`;
        fileNameEl.classList.add('loaded');

        if (hasAnyError && totalCount === 0) {
          showAlert('lbAlert', '파일을 불러오지 못했습니다. 오류 내용을 확인해주세요.', 'error');
        }
      };
      reader.readAsArrayBuffer(file);
    });
  });

  // ③ 파싱 & 미리보기
  document.getElementById('btnParse').addEventListener('click', () => {
    clearAlert('lbAlert');

    if (!state.metabaseLoaded) {
      showAlert('lbAlert', 'Metabase 파일을 먼저 업로드해주세요.', 'warning');
      return;
    }

    const year    = parseInt(document.getElementById('yearInput').value, 10) || new Date().getFullYear();
    const rawText = document.getElementById('lbText').value;
    const lines   = rawText.split('\n').filter(l => l.trim() !== '');

    if (lines.length === 0) {
      showAlert('lbAlert', 'LB_ 텍스트를 입력해주세요.', 'warning');
      return;
    }

    state.parsedRows = applyLB6Dedup(lines.map(line => parseLB(line, year)));
    state.mainTitles = {};
    renderLBTable();
  });

  // ④ 라이브 템플릿 다운로드
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

  // ⑤ 라이브배너등록 템플릿 — 위치별(LB_2~LB_6) 다운로드
  const LB_POSITIONS = ['2', '3', '4', '5', '6'];
  const LB_HEADER = [
    '콘텐츠 카드 번호',
    '날짜 (yyyy.mm.dd)',
    '방송 시작 시간 (hh:mm:ss)',
    '방송 종료 시간 (hh:mm:ss)',
  ];

  LB_POSITIONS.forEach(pos => {
    document.querySelector(`[data-lb-pos="${pos}"]`).addEventListener('click', () => {
      if (state.parsedRows.length === 0) {
        showAlert('lbAlert', '먼저 [파싱 & 미리보기]를 실행해주세요.', 'warning');
        return;
      }
      let filteredRows = state.parsedRows.filter(row => !row.parseError && row.position === pos);
      if (pos === '6') filteredRows = deduplicateLB6(filteredRows);
      const dataRows = filteredRows.map(row => {
        const endTime = (pos === '6' && row.lb6HasPair) ? `'${calcEndTime(row.timeFormatted)}` : '';
        return ['', `'${row.dateFull}`, `'${row.timeFormatted}`, endTime];
      });

      if (dataRows.length === 0) {
        showAlert('lbAlert', `LB_${pos} 위치의 방송이 없습니다.`, 'warning');
        return;
      }
      const wb = XLSX.utils.book_new();
      const wsData = [LB_HEADER, ...dataRows];
      const ws = XLSX.utils.aoa_to_sheet(wsData);
      forceStringCells(ws, wsData);
      XLSX.utils.book_append_sheet(wb, ws, `LB_${pos}`);
      XLSX.writeFile(wb, `라이브배너등록 템플릿_LB${pos}.xlsx`);
      clearAlert('lbAlert');
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
