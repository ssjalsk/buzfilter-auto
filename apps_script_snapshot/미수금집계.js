/**
 * 미수금집계.js
 * 업무시트 E열='입금확인' + C열='언론' → 언론 매출 통계 미수 리스트 자동 갱신
 * - G열: 업무시트 해당 행 바로가기 링크 (H열 왼쪽 빈 열)
 * - S열이 숫자(누적 입금액)면 실잔액 = N - S 표시
 * - E열 변경 감지 트리거 (설치형)로 자동 갱신
 */

const MISU_CFG = {
  WORK_SHEET:  '업무시트',
  STATS_SHEET: '언론 매출 통계',
  HEADER_ROW:  35,
  START_ROW:   36,
  MAX_ROWS:    40,
  COL_G:  7,   // G: 바로가기 링크 (H열 왼쪽)
  COL_H:  8,   // H: 고객명
  COL_I:  9,   // I: 미수금 (실잔액)
  COL_J: 10,   // J: VAT포함
};

/* 숫자 판별: S열에 숫자가 저장된 경우만 부분입금으로 처리 */
function _isMisuNum_(v) {
  if (v === null || v === undefined || v === '') return false;
  if (typeof v === 'number') return !isNaN(v) && v > 0;
  const s = String(v).trim();
  if (/초과입금|부족/.test(s)) return false;
  const n = parseFloat(s.replace(/,/g,''));
  return !isNaN(n) && n > 0;
}

function _misuNum_(v) {
  if (!_isMisuNum_(v)) return 0;
  if (typeof v === 'number') return Math.round(v);
  return Math.round(parseFloat(String(v).replace(/,/g,'')));
}

// ─── 핵심: 미수금 목록 갱신 ──────────────────────────────
function refreshMisuList() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const wsWork  = ss.getSheetByName(MISU_CFG.WORK_SHEET);
  const wsStats = ss.getSheetByName(MISU_CFG.STATS_SHEET);
  if (!wsWork || !wsStats) return;

  const SHEET_ID = ss.getId();
  const WORK_GID = wsWork.getSheetId();

  const lastRow = wsWork.getLastRow();
  if (lastRow < 2) return;
  const data = wsWork.getRange(2, 1, lastRow - 1, 20).getValues();

  // C='언론' + E='입금확인' 필터
  const misuList = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (String(row[2] || '').trim() !== '언론')      continue;
    if (String(row[4] || '').trim() !== '입금확인')  continue;

    const rowNum    = i + 2;
    const name      = String(row[1]  || '').trim();
    const totalCost = Math.round(Number(String(row[13]||'').replace(/,/g,''))||0); // N열
    const sVal      = row[18]; // S열 (누적 입금액 또는 텍스트)
    const paidSoFar = _misuNum_(sVal);
    const amount    = Math.max(0, totalCost - paidSoFar);

    const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit#gid=${WORK_GID}&range=A${rowNum}`;
    misuList.push({ name, amount, totalCost, paidSoFar, url });
  }

  const { HEADER_ROW, START_ROW, MAX_ROWS, COL_G, COL_H, COL_I, COL_J } = MISU_CFG;

  // H/I/J열은 수식(FILTER)으로 자동 관리 → G열(바로가기 링크)만 처리
  // 기존 G열 클리어
  wsStats.getRange(HEADER_ROW, COL_G, MAX_ROWS + 2, 1).clearContent();

  if (misuList.length === 0) return;

  // G열: 바로가기 링크만 업데이트
  for (let i = 0; i < misuList.length && i < MAX_ROWS; i++) {
    const { url } = misuList[i];
    const r = START_ROW + i;
    wsStats.getRange(r, COL_G).setFormula(`=HYPERLINK("${url}","→ 이동")`);
  }
}

// ─── 설치형 onEdit 트리거 등록 (최초 1회) ────────────────
function installMisuTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onEditMisuCheck_')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onEditMisuCheck_')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ 자동 갱신 트리거 설치 완료\n' +
    '업무시트 E열 변경 시 미수금 목록 자동 갱신됩니다.'
  );
}

// ─── E열 변경 감지 핸들러 ────────────────────────────────
function onEditMisuCheck_(e) {
  if (!e) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== MISU_CFG.WORK_SHEET) return;
  if (e.range.getColumn() !== 5) return;

  const newVal = String(e.value    || '').trim();
  const oldVal = String(e.oldValue || '').trim();

  if (newVal === '입금확인' || oldVal === '입금확인' ||
      newVal === '송출완료' || oldVal === '송출완료') {
    refreshMisuList();
  }
}

// ─── 메뉴 등록은 이메일.js onOpen() 통합 관리 ─────────────
