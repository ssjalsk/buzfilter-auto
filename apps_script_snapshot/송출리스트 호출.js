/************************************************************
 * Project: 한용범
 * File   : 송출리스트 호출.gs
 * Trigger: From spreadsheet → On edit → HYB_Trigger (설치형)
 * Note   : 반려 이벤트 ON. 새 실행사에는 "빈 슬롯(B empty)"로 회차 재배정.
 *          비건뉴스: 매체사+소재 기준 3섹션(특수/일반/GEO) 자동 분기
 *          공감신문: 반려 시 A열(회차)만 남기고 B~H 전부 삭제
 ************************************************************/

// 이피알몰 외부 시트 ID
const EPR_SHEET_ID = '1kZrvcCaIOBFnJpReGmqGC1bVnFAYH1Jx-ESY-dTS54s';
const EPR_KEY      = '__EPR__';  // EXECUTOR_MAP 내부 식별자

// 박병욱 외부 시트 ID
const BBW_SHEET_ID = '1_XesHNaj4MG1_7Vg0EsjokwvcAbrUN--JEmfDMnVidI';
const BBW_KEY      = '__BBW__';  // EXECUTOR_MAP 내부 식별자

const HYB = {
  PROJECT_NAME: '한용범',
  MAIN_SHEET: '업무시트',
  LOG_SHEET: 'HYB_LOG',

  COL: {
    SUBJECT: 1,        // A 기사주제
    STATUS:  5,        // E 진행상태
    SEND_DATE: 7,      // G 송출일
    MEDIA: 9,          // I 매체사
    EXECUTOR: 10,      // J 실행사
    MATERIAL: 11,      // K 소재
    ROUND: 12,         // L 회차
    LINK: 13,          // M 기사링크
    SALES: 14,         // N 매출금액 (수식 → 최종반려 시 0 덮어씀)
    COST: 15,          // O 원가
    EXECUTOR_CACHE: 27, // AA 실행사_캐시
    TASK_ID: 29        // AC 고유번호 (실행사 시트 칸 메모와 연결) [2026-10-16]
  },

  EXECUTOR_MAP: {
    '플래나':   '플래나 송출',
    '한미마':   '한미마 송출',
    '정진욱':   '공감신문',
    '이디엘':   '이디엘 송출',
    '비건':     '비건뉴스',
    '이피알몰': EPR_KEY,      // ← 이피알몰 외부 시트
    '박병욱':   BBW_KEY       // ← 박병욱 외부 시트
  },

  TARGET_LAYOUT: {
    '플래나 송출': { A_round:1, B_subject:2, C_date:3, D_link:4, E_media:5, F_material:6, G_cost:7, H_flag:8, linkCol:4 },
    '한미마 송출': { A_round:1, B_subject:2, C_date:3, D_media:4, E_material:5, F_link:6, G_cost:7, H_flag:8, linkCol:6 },
    '공감신문'   : { A_round:1, B_subject:2, C_date:3, D_media:4, E_material:5, F_link:6, G_cost:7, H_flag:8, linkCol:6 },
    '이디엘 송출': { A_round:1, B_subject:2, C_date:3, D_media:4, E_material:5, F_link:6, G_cost:7, H_flag:8, linkCol:6 },
    [EPR_KEY]:    { A_round:1, B_subject:2, C_date:3, D_media:4, E_material:5, F_link:6, G_cost:7, H_flag:8, linkCol:6 },
    [BBW_KEY]:    { A_round:1, B_subject:2, C_date:3, D_media:4, E_material:5, F_link:6, G_cost:7, H_flag:8, linkCol:6 }
    // 비건뉴스는 3섹션 구조라 VEGAN_SECTION_LAYOUT에 별도 정의
  },

  LOG_ENABLED: true,
  STABILIZE_DELAY_MS: 5000,
  LOCK_WAIT_MS: 60000   // [2026-10-16] 20초 → 60초 (연속 입력 때 대기열이 길어 동기화를 포기하던 문제)
};

/**
 * 실행사 시트 가져오기
 * - 이피알몰(EPR_KEY)이면 외부 스프레드시트에서 첫 번째 탭 반환
 * - 그 외는 현재 스프레드시트 내 탭 반환
 */
function getTargetSheet_(name) {
  if (name === EPR_KEY) {
    try {
      return SpreadsheetApp.openById(EPR_SHEET_ID).getSheets()[0];
    } catch (e) {
      log_('ERROR', '이피알몰 외부 시트 열기 실패', { msg: e.message });
      return null;
    }
  }
  if (name === BBW_KEY) {
    try {
      return SpreadsheetApp.openById(BBW_SHEET_ID).getSheets()[0];
    } catch (e) {
      log_('ERROR', '박병욱 외부 시트 열기 실패', { msg: e.message });
      return null;
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

// ─── 비건뉴스 3섹션 레이아웃 ────────────────────────────────────
// 섹션별: 회차(A_round) / 기사주제(B_subject) / 송출일(C_date) / 링크(D_link) / 진행현황(E_flag)
const VEGAN_SECTION_LAYOUT = {
  '특수': { A_round:1,  B_subject:2,  C_date:3,  D_link:4,  E_flag:5  },
  '일반': { A_round:8,  B_subject:9,  C_date:10, D_link:11, E_flag:12 },
  'GEO':  { A_round:14, B_subject:15, C_date:16, D_link:17, E_flag:18 }
};

/** 매체사+소재 기반 비건뉴스 섹션 결정 */
function getVeganSection_(media, material) {
  if (media === '구글비건') return 'GEO';
  if (material === '일반')  return '일반';
  return '특수';
}

// ============== [2026-10-16] 고유번호·슬롯 관리 ==============
/**
 * 업무마다 고유번호(업무시트 AC열)를 붙이고, 실행사 시트의 그 칸 B열(주제) 메모에 'HYB:<번호>@<업무시트 행>' 로 기록.
 * 수정·반려·링크 입력 때 회차가 아니라 고유번호로 정확히 그 칸을 찾음
 * → 회차 번호가 여러 달·충전 구간에 겹쳐도, 다른 실행사 회차가 L열에 남아 있어도 다른 고객 칸을 건드리지 않음.
 * 고유번호가 없는 예전 건: '회차+주제'(지금 실행사 회차일 때만) → '주제+송출일' 이 맞는 칸만 인정. 주제가 다른 칸은 절대 덮어쓰지 않음.
 */
const ID_TAG = 'HYB:';
const DATA_START_ROW = 3;   // 실행사 시트: 1행 제목, 2행 머리글

// 회차 자동 채우기: 갱신(월초·월요일·충전) 때 이만큼 번호 칸 확보, 평소엔 남은 칸이 SLOT_LOW 이하이면 같은 만큼 이어 붙임
const SLOT_CFG = {
  '플래나 송출': { target: 150 },
  '이디엘 송출': { target: 15 },
  '한미마 송출': { target: 30 },
  [BBW_KEY]:     { target: 30 },
  [EPR_KEY]:     { target: 20 }
  // 공감신문: 회차 갱신·자동 채우기 없음 (58회차 계약, 번호는 대표님이 관리) — 반려로 비운 칸 재사용만
  // 비건뉴스: 섹션별 결제 단위(100회·30회)라 자동 채우기 제외
};
const SLOT_LOW = 5;

function _newTaskId_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyMMddHHmmss') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
}
function _parseTag_(note) {
  const m = String(note || '').match(/HYB:([^@\s]+)(?:@(\d+))?/);
  return m ? { id: m[1], row: m[2] ? Number(m[2]) : 0 } : null;
}
function _setTag_(sheet, rowIndex, subjectCol, id, mainRow) {
  sheet.getRange(rowIndex, subjectCol).setNote(id ? `${ID_TAG}${id}@${mainRow}` : '');
}

/** 업무시트 행의 고유번호 { id, dup } — 없으면 새로 만듦. dup=같은 번호가 다른 행에도 있음(행 복사) */
function _taskIdOf_(main, row) {
  const cell = main.getRange(row, HYB.COL.TASK_ID);
  let id = String(cell.getValue() || '').trim();
  if (!id) { id = _newTaskId_(); cell.setValue(id); return { id, dup: false }; }
  const last = main.getLastRow();
  const ids  = main.getRange(2, HYB.COL.TASK_ID, last - 1, 1).getValues();
  const dup  = ids.some((r, i) => i + 2 !== row && String(r[0]).trim() === id);
  return { id, dup };
}
/** 행 복사로 번호가 겹쳤는데 이 행이 원래 주인이 아니면 → 새 번호 + L·실행사기억 비움 (새 업무로 처리) */
function _regenTaskId_(main, row) {
  const id = _newTaskId_();
  main.getRange(row, HYB.COL.TASK_ID).setValue(id);
  main.getRange(row, HYB.COL.ROUND).setValue('');
  main.getRange(row, HYB.COL.EXECUTOR_CACHE).setValue('');
  log_('INFO', '고유번호 중복(행 복사) → 새 번호', { row, id });
  return id;
}

/** 실행사 시트 슬롯 읽기: vals(DATA_START_ROW부터), tags(B열 메모 파싱) */
function _readSlots_(sheet, layout) {
  const last = sheet.getLastRow();
  const n = last - DATA_START_ROW + 1;
  if (n <= 0) return { vals: [], tags: [] };
  const maxCol = Math.max(layout.A_round, layout.B_subject, layout.C_date, layout.G_cost || 1, layout.H_flag || 1, layout.E_flag || 1, layout.D_link || 1);
  const vals  = sheet.getRange(DATA_START_ROW, 1, n, maxCol).getValues();
  const notes = sheet.getRange(DATA_START_ROW, layout.B_subject, n, 1).getNotes();
  return { vals, tags: notes.map(r => _parseTag_(r[0])) };
}
const _rowOf_ = i => i + DATA_START_ROW;
const _cellStr_ = (vals, i, col) => String(vals[i][col - 1] === null || vals[i][col - 1] === undefined ? '' : vals[i][col - 1]).trim();

function _isRejectedSlot_(sheetName, layout, vals, i) {
  if (layout.E_flag) return _cellStr_(vals, i, layout.E_flag) === '반려';          // 비건 섹션
  if (sheetName === '이디엘 송출') return _cellStr_(vals, i, layout.G_cost) === '반려';
  return layout.H_flag ? _cellStr_(vals, i, layout.H_flag) === '반려' : false;
}

/** 현재 회차 구간 시작 인덱스: 빈 칸 첫 번호가 1이면(막 갱신) 그 칸, 아니면 마지막으로 1회차가 쓰인 칸. 공감·비건=0(구간 없음) */
function _cycleStartIdx_(sheetName, layout, vals) {
  if (!SLOT_CFG[sheetName] || SLOT_CFG[sheetName].cap) return 0;
  let lastUsed = -1;
  for (let i = vals.length - 1; i >= 0; i--) if (_cellStr_(vals, i, layout.B_subject)) { lastUsed = i; break; }
  if (lastUsed + 1 < vals.length && _cellStr_(vals, lastUsed + 1, layout.A_round) === '1') return lastUsed + 1;
  for (let i = lastUsed; i >= 0; i--) {
    if (_cellStr_(vals, i, layout.A_round) === '1' && _cellStr_(vals, i, layout.B_subject)) return i;
  }
  return 0;
}

/**
 * 이 업무가 이미 들어간 칸 찾기 → 인덱스 또는 -1
 *  1) 고유번호 일치 (행 복사로 번호가 겹친 경우엔 기록된 업무시트 행이 같거나 주제+송출일이 같을 때만)
 *  2) [예전 건] 회차+주제 일치 — roundTrusted(지금 실행사의 회차일 때)만, 고유번호 없는 칸만, 송출일 가까운 칸
 *  3) [예전 건] 주제+송출일 일치 — 고유번호 없는 칸만
 */
function _findTaskSlot_(S, layout, { id, dup, mainRow, roundTrusted, subject, sendDateStr }) {
  const { vals, tags } = S;
  const d = normalizeDateString_(sendDateStr);
  if (id) {
    for (let i = vals.length - 1; i >= 0; i--) {
      const t = tags[i];
      if (!t || t.id !== id) continue;
      if (!dup || t.row === mainRow ||
          (_cellStr_(vals, i, layout.B_subject) === subject && normalizeDateString_(vals[i][layout.C_date - 1]) === d)) return i;
    }
  }
  if (roundTrusted) {
    let best = -1, bestGap = Infinity;
    for (let i = 0; i < vals.length; i++) {
      if (tags[i]) continue;
      if (_cellStr_(vals, i, layout.A_round) !== String(roundTrusted)) continue;
      if (_cellStr_(vals, i, layout.B_subject) !== subject) continue;
      const gap = Math.abs(_dayNum_(vals[i][layout.C_date - 1]) - _dayNum_(sendDateStr));
      if (gap <= bestGap) { best = i; bestGap = gap; }
    }
    if (best >= 0) return best;
  }
  for (let i = vals.length - 1; i >= 0; i--) {
    if (tags[i]) continue;
    if (_cellStr_(vals, i, layout.B_subject) === subject && normalizeDateString_(vals[i][layout.C_date - 1]) === d) return i;
  }
  return -1;
}
function _dayNum_(v) {
  const s = normalizeDateString_(v); const p = s.split('.');
  if (p.length !== 3) return 0;
  const t = new Date(+p[0], +p[1] - 1, +p[2]).getTime();
  return isNaN(t) ? 0 : Math.round(t / 86400000);
}

/** 새로 배정할 칸: 현재 회차 구간 안의 첫 빈 칸 (공감신문은 반려로 비워진 칸 포함 맨 위부터) */
function _newSlotIdx_(sheetName, layout, S) {
  const start = _cycleStartIdx_(sheetName, layout, S.vals);
  for (let i = start; i < S.vals.length; i++) {
    if (!_cellStr_(S.vals, i, layout.B_subject)) return i;
  }
  return -1;
}

/**
 * 업무 배정/갱신 공통: 이 업무 칸이 있으면 그 칸, 없으면 새 칸.
 * 이 업무의 반려 칸은 현재 회차 구간 안일 때만 재사용 (대표님 기준: 같은 업무가 돌아오면 그 반려 회차 자리 사용,
 * 단 지난달·지난 충전 구간 칸은 쓰지 않음)
 */
function _resolveAssignSlot_(sheetName, layout, S, key) {
  let i = _findTaskSlot_(S, layout, key);
  if (i >= 0 && _isRejectedSlot_(sheetName, layout, S.vals, i) && i < _cycleStartIdx_(sheetName, layout, S.vals)) i = -1;
  if (i < 0) i = _newSlotIdx_(sheetName, layout, S);
  return i;
}

// ============== 설치형 트리거 진입 ==============
function HYB_Trigger(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();

    // ── 한미마 K열 충전 감지 → 회차 갱신 (빈 칸 1부터 + 30칸 확보) ──────────────
    if (sheet.getName() === '한미마 송출') {
      const col = e.range.getColumn();
      const row = e.range.getRow();
      if (col === 11 && row >= 2) {          // K열(11) = 충전 금액
        const val = Number(String(e.range.getValue() || '').replace(/[,\s]/g, ''));
        if (val > 0) {
          _ensureSlots_(sheet, '한미마 송출', HYB.TARGET_LAYOUT['한미마 송출'], true);
          log_('INFO', '한미마 충전 리셋', { chargedRow: row, amount: val });
        }
      }
      return;
    }

    if (sheet.getName() !== HYB.MAIN_SHEET) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row <= 1) return;

    // 감시 컬럼: E/G/I/J/K/M
    const watch = [HYB.COL.STATUS, HYB.COL.SEND_DATE, HYB.COL.MEDIA, HYB.COL.EXECUTOR, HYB.COL.MATERIAL, HYB.COL.LINK];
    if (!watch.includes(col)) return;

    // [2026-10-16 속도] 예전: I/J/K 변경 시 무조건 5초 대기(실행사 자동 배정 기다림)
    //  → 실행사는 아래에서 매체단가표로 직접 계산하므로 새 업무는 대기 없음.
    //    이미 배정된 업무의 매체사·소재를 바꿀 때만 2초 대기 (I·K를 연달아 고칠 때 중간 실행사로 잘못 이관 방지)
    const isIJK = [HYB.COL.MEDIA, HYB.COL.EXECUTOR, HYB.COL.MATERIAL].includes(col);
    if (isIJK && col !== HYB.COL.EXECUTOR && String(sheet.getRange(row, HYB.COL.EXECUTOR_CACHE).getValue() || '').trim()) {
      Utilities.sleep(2000);
    }

    // [2026-10-16] 송출 동기화 전용 잠금(문서 잠금) — 정산 자동매칭(스크립트 잠금)이 오래 걸려도 서로 막지 않음
    const lock = LockService.getDocumentLock();
    lock.waitLock(HYB.LOCK_WAIT_MS);

    // [2026-10-16] 같은 행을 연달아 처리할 때 값이 그대로면 건너뜀 (대기열 단축)
    const sigBefore = _hybRowSig_(sheet, row);
    if (col !== HYB.COL.STATUS) {
      const cache = CacheService.getScriptCache();
      if (cache.get('hyb:' + row) === sigBefore) { lock.releaseLock(); return; }
    }

    // 값 읽기
    const subject     = String(sheet.getRange(row, HYB.COL.SUBJECT).getValue()       || '').trim();
    const sendDate    = sheet.getRange(row, HYB.COL.SEND_DATE).getValue();
    const sendDateStr = normalizeDateString_(sendDate);
    const media       = String(sheet.getRange(row, HYB.COL.MEDIA).getValue()          || '').trim();
    let   execNow     = String(sheet.getRange(row, HYB.COL.EXECUTOR).getValue()       || '').trim();
    const material    = String(sheet.getRange(row, HYB.COL.MATERIAL).getValue()       || '').trim();
    const link        = String(sheet.getRange(row, HYB.COL.LINK).getValue()           || '').trim();
    let   cost        = sheet.getRange(row, HYB.COL.COST).getValue();
    // 매체사·소재를 바꾼 경우: 업무시트 자동배정(onEdit)과 같은 규칙으로 실행사·원가를 바로 계산 (기다리지 않음)
    if (col === HYB.COL.MEDIA || col === HYB.COL.MATERIAL) {
      const auto = _autoExecutor_(media, material);
      if (auto) { execNow = auto.exec; cost = auto.cost; }
    }
    const roundNow    = String(sheet.getRange(row, HYB.COL.ROUND).getValue()          || '').trim();
    const execCache   = String(sheet.getRange(row, HYB.COL.EXECUTOR_CACHE).getValue() || '').trim();

    const targetSheetName = HYB.EXECUTOR_MAP[execNow] || '';
    const isVegan         = (targetSheetName === '비건뉴스');
    const task = _taskIdOf_(sheet, row);
    // L열 회차는 '지금 실행사의 회차'로 확인될 때만 사용 (다른 실행사 회차가 남아 엉뚱한 칸을 찾는 버그 방지)
    const roundTrusted = (execCache && execCache === execNow) ? roundNow : '';
    const ctx = { row, subject, sendDateStr, media, material, cost, link, task, roundTrusted, linkOnly: col === HYB.COL.LINK };

    // 0) E열(진행상태) = '반려' → 최종 반려 처리 (재배정 없음)
    if (col === HYB.COL.STATUS) {
      const statusVal = String(sheet.getRange(row, HYB.COL.STATUS).getValue() || '').trim();
      if (statusVal === '반려') {
        handleFinalReject_({ ...ctx, sheet, roundNow, execNow, execCache });
      }
      lock.releaseLock();
      return;
    }

    // 1) 실행사 변경 감지 → 반려 이벤트
    const isExecChanged = !!execCache && execCache !== execNow;
    if (isExecChanged) {
      handleRejectAndReassign_({ ...ctx, oldExecutor: execCache, newExecutor: execNow, roundNow });
      lock.releaseLock();
      return;
    }

    // 실행사 기억은 없는데 실행사가 비었거나 목록에 없음 → 남은 L열 회차 비움
    if (!targetSheetName) {
      if (roundNow && !execCache) sheet.getRange(row, HYB.COL.ROUND).setValue('');
      log_('WARN', col === HYB.COL.LINK ? 'Link edit skipped: executor not set' : 'Base upsert skipped: executor not set', { row });
      lock.releaseLock();
      return;
    }

    // 2) 링크 편집 / 3) 기본 업서트(G/I/J/K) — 둘 다 이 업무 칸을 찾아 갱신(없으면 새 칸 배정)
    if (isVegan) {
      handleVeganUpsert_(ctx);
    } else {
      handleBaseUpsert_({ ...ctx, targetSheetName, execName: execNow });
    }
    try { CacheService.getScriptCache().put('hyb:' + row, _hybRowSig_(sheet, row), 21600); } catch (e) {}
    lock.releaseLock();

  } catch (err) {
    log_('ERROR','HYB_Trigger failed',{message: err && err.message, stack: err && err.stack});
    try { LockService.getDocumentLock().releaseLock(); } catch (e) {}
  }
}

/** 송출 동기화에 쓰이는 값 묶음 (주제·송출일·매체사·실행사·소재·회차·링크·원가·실행사기억·고유번호) */
function _hybRowSig_(sheet, row) {
  const v = sheet.getRange(row, 1, 1, HYB.COL.TASK_ID).getValues()[0];
  const c = HYB.COL;
  return JSON.stringify([c.SUBJECT, c.SEND_DATE, c.MEDIA, c.EXECUTOR, c.MATERIAL, c.ROUND, c.LINK, c.COST, c.EXECUTOR_CACHE, c.TASK_ID]
    .map(k => String(v[k - 1] === null || v[k - 1] === undefined ? '' : v[k - 1]).trim()));
}

// ============== 기본 업서트 (플래나·한미마·이디엘·공감·외부 시트) ==============
function handleBaseUpsert_({ row, subject, sendDateStr, media, material, cost, link, task, roundTrusted, targetSheetName, linkOnly, execName }) {
  const main   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HYB.MAIN_SHEET);
  const target = getTargetSheet_(targetSheetName);
  const layout = HYB.TARGET_LAYOUT[targetSheetName];
  if (!target || !layout) { log_('ERROR','Target missing',{targetSheetName}); return; }

  // [속도] 실행사 시트는 한 번만 읽음 (플래나 월 리셋 확인도 읽은 값으로)
  let S = _readSlots_(target, layout);
  if (targetSheetName === '플래나 송출' && _plannerNeedsReset_(S, layout)) {
    _ensureSlots_(target, targetSheetName, layout, true);
    log_('INFO', '플래나 월 리셋', {});
    S = _readSlots_(target, layout);
  }

  let key = { id: task.id, dup: task.dup, mainRow: row, roundTrusted, subject, sendDateStr };
  let i = _findTaskSlot_(S, layout, key);
  if (i < 0 && task.dup) {               // 행 복사본 → 새 업무로
    key = { ...key, id: _regenTaskId_(main, row), dup: false, roundTrusted: '' };
  }
  const existing = i >= 0 && !_isRejectedSlot_(targetSheetName, layout, S.vals, i);
  i = _resolveAssignSlot_(targetSheetName, layout, S, key);
  if (i < 0) { log_('WARN','No free slot',{targetSheetName}); return; }
  const rowIndex = _rowOf_(i);

  // [속도] 회차를 가장 먼저 업무시트에 표시 (담당자가 바로 보고 접수할 수 있게)
  const roundVal = _cellStr_(S.vals, i, layout.A_round);
  main.getRange(row, HYB.COL.ROUND).setValue(roundVal);
  const execVal = execName || String(main.getRange(row, HYB.COL.EXECUTOR).getValue() || '').trim();
  if (execVal) main.getRange(row, HYB.COL.EXECUTOR_CACHE).setValue(execVal);
  SpreadsheetApp.flush();

  if (linkOnly && existing) {
    // 링크만 입력: 이미 있는 이 업무 칸의 링크만 갱신 (담당자가 직접 적은 금액·'결제 완' 등 보존)
    target.getRange(rowIndex, layout.linkCol).setValue(link || '');
  } else {
    // [속도] B~H 한 번에 쓰기 (새 링크 없으면 기존 링크 유지)
    const cur = S.vals[i].slice(0, 8);
    while (cur.length < 8) cur.push('');
    const put = (col, v) => { cur[col - 1] = v; };
    put(layout.B_subject, subject || '');
    put(layout.C_date, normalizeDateString_(sendDateStr) || '');
    put(layout.linkCol === 4 ? layout.E_media : layout.D_media, media || '');
    put(layout.linkCol === 4 ? layout.F_material : layout.E_material, material || '');
    if (cost !== undefined) put(layout.G_cost, cost || '');
    if (layout.H_flag) put(layout.H_flag, '');
    if (link) put(layout.linkCol, link);
    target.getRange(rowIndex, 2, 1, 7).setValues([cur.slice(1, 8)]);
  }
  _setTag_(target, rowIndex, layout.B_subject, key.id, row);

  // [속도] 회차 칸 정리·채우기는 필요할 때만
  if (_slotsNeedWork_(targetSheetName, layout, S, i)) _ensureSlots_(target, targetSheetName, layout, false);
  log_('INFO','Base upsert OK',{ targetSheetName, rowIndex, roundVal, id: key.id });
}

/** 매체사+소재 → 매체단가표 최저 원가 실행사 (업무시트 onEdit 자동배정과 같은 규칙). 없으면 null */
function _autoExecutor_(media, material) {
  if (!media || !material || material === '수작업') return null;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('매체단가표');
  if (!sh) return null;
  const last = sh.getLastRow(); if (last < 2) return null;
  const nz = v => String(v === null || v === undefined ? '' : v).replace(/\s+/g, ' ').trim();
  const money = v => { const n = parseFloat(String(v === null || v === undefined ? '' : v).replace(/[^\d.-]/g, '')); return isNaN(n) ? 0 : Math.round(n); };
  let best = null;
  for (const r of sh.getRange(2, 1, last - 1, 4).getValues()) {
    if (nz(r[0]) !== nz(media) || nz(r[3]) !== nz(material)) continue;
    const c = money(r[2]);
    // 동점이면 표에서 아래쪽 실행사 — 업무시트 자동배정(reduce: a.원가 < b.원가 ? a : b)과 똑같이
    if (c > 0 && (!best || c <= best.cost)) best = { exec: nz(r[1]), cost: c };
  }
  return best;
}

/** 플래나: 읽어 둔 칸 값으로 월 리셋 필요 여부 (마지막 사용 칸의 송출일 월 ≠ 이번 달) */
function _plannerNeedsReset_(S, layout) {
  const now = new Date();
  for (let i = S.vals.length - 1; i >= 0; i--) {
    if (!_cellStr_(S.vals, i, layout.B_subject)) continue;
    const p = (normalizeDateString_(S.vals[i][layout.C_date - 1]) || '').split('.');
    if (p.length === 3 && !isNaN(parseInt(p[0], 10))) {
      return parseInt(p[0], 10) !== now.getFullYear() || parseInt(p[1], 10) !== now.getMonth() + 1;
    }
  }
  return true;
}

/** 배정 후 빈 칸 번호가 차례대로이고 SLOT_LOW 보다 많이 남았으면 다시 읽지 않음 */
function _slotsNeedWork_(sheetName, layout, S, assignedIdx) {
  if (!SLOT_CFG[sheetName]) return false;
  let lastUsed = assignedIdx;
  for (let k = S.vals.length - 1; k > lastUsed; k--) if (_cellStr_(S.vals, k, layout.B_subject)) { lastUsed = k; break; }
  const tail = S.vals.length - (lastUsed + 1);
  if (tail <= SLOT_LOW) return true;
  const first = parseInt(_cellStr_(S.vals, lastUsed + 1, layout.A_round), 10);
  if (isNaN(first)) return true;
  for (let k = 1; k < tail; k++) {
    if (parseInt(_cellStr_(S.vals, lastUsed + 1 + k, layout.A_round), 10) !== first + k) return true;
  }
  return false;
}

// ============== 비건뉴스 전용 업서트 ==============
function handleVeganUpsert_({ row, subject, sendDateStr, media, material, link, task, roundTrusted }) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const main  = ss.getSheetByName(HYB.MAIN_SHEET);
  const vSheet= ss.getSheetByName('비건뉴스');
  if (!vSheet) { log_('ERROR','비건뉴스 시트 없음'); return; }

  const section = getVeganSection_(media, material);
  const layout  = VEGAN_SECTION_LAYOUT[section];
  let key = { id: task.id, dup: task.dup, mainRow: row, roundTrusted, subject, sendDateStr };

  // 섹션이 바뀐 경우(같은 비건 실행사 안에서 매체사·소재 변경): 예전 섹션의 이 업무 칸은 반려 처리
  for (const [sec, lay] of Object.entries(VEGAN_SECTION_LAYOUT)) {
    if (sec === section) continue;
    const S2 = _readSlots_(vSheet, lay);
    const j = _findTaskSlot_(S2, lay, { ...key, roundTrusted: '' });
    if (j >= 0 && S2.tags[j] && _cellStr_(S2.vals, j, lay.E_flag) !== '반려') {
      vSheet.getRange(_rowOf_(j), lay.E_flag).setValue('반려');
      key.roundTrusted = '';
      log_('INFO','Vegan 섹션 변경 → 이전 섹션 반려',{ from: sec, to: section, row: _rowOf_(j) });
    }
  }

  const S = _readSlots_(vSheet, layout);
  let i = _findTaskSlot_(S, layout, key);
  if (i < 0 && task.dup) key = { ...key, id: _regenTaskId_(main, row), dup: false, roundTrusted: '' };
  if (i < 0) i = _veganNewSlotIdx_(layout, S, subject);
  if (i < 0) { log_('WARN','Vegan upsert: 빈 슬롯 없음',{section}); return; }
  const targetRow = _rowOf_(i);

  vSheet.getRange(targetRow, layout.B_subject).setValue(subject || '');
  vSheet.getRange(targetRow, layout.C_date).setValue(normalizeDateString_(sendDateStr) || '');
  if (link) vSheet.getRange(targetRow, layout.D_link).setValue(link);
  _setTag_(vSheet, targetRow, layout.B_subject, key.id, row);

  // 반려 슬롯 재사용 시 진행현황 → '대체완료' (같은 업무가 돌아온 경우는 반려 표시만 지움)
  const eFlag = _cellStr_(S.vals, i, layout.E_flag);
  if (eFlag === '반려') {
    const own = S.tags[i] && S.tags[i].id === key.id;
    vSheet.getRange(targetRow, layout.E_flag).setValue(own ? '' : '대체완료');
  }

  const roundVal = String(vSheet.getRange(targetRow, layout.A_round).getValue() || '').trim();
  main.getRange(row, HYB.COL.ROUND).setValue(roundVal);
  main.getRange(row, HYB.COL.EXECUTOR_CACHE).setValue(main.getRange(row, HYB.COL.EXECUTOR).getValue());
  log_('INFO','Vegan upsert OK',{section, targetRow, roundVal, eFlag, id: key.id});
}

/** 비건 새 칸: 같은 주제 반려 칸 → 아무 반려 칸(수량 보전 — 대체) → 회차 있는 빈 칸 → 완전 빈 칸 */
function _veganNewSlotIdx_(layout, S, subject) {
  const { vals } = S;
  for (let i = 0; i < vals.length; i++)
    if (_cellStr_(vals, i, layout.E_flag) === '반려' && _cellStr_(vals, i, layout.B_subject) === subject) return i;
  for (let i = 0; i < vals.length; i++)
    if (_cellStr_(vals, i, layout.E_flag) === '반려') return i;
  for (let i = 0; i < vals.length; i++)
    if (_cellStr_(vals, i, layout.A_round) && !_cellStr_(vals, i, layout.B_subject)) return i;
  for (let i = 0; i < vals.length; i++)
    if (!_cellStr_(vals, i, layout.A_round) && !_cellStr_(vals, i, layout.B_subject)) return i;
  return -1;
}

/** 비건 반려: 이 업무 칸(고유번호 → 회차+주제 → 주제+날짜)만. 못 찾으면 아무것도 안 함 */
function applyVeganReject_(vSheet, media, material, key) {
  const order = [getVeganSection_(media, material), ...Object.keys(VEGAN_SECTION_LAYOUT)];
  for (const section of order) {
    const layout = VEGAN_SECTION_LAYOUT[section];
    const S = _readSlots_(vSheet, layout);
    const i = _findTaskSlot_(S, layout, section === order[0] ? key : { ...key, roundTrusted: '' });
    if (i < 0) continue;
    if (_cellStr_(S.vals, i, layout.E_flag) !== '대체완료') {
      vSheet.getRange(_rowOf_(i), layout.E_flag).setValue('반려');
      log_('INFO','Vegan reject OK',{section, row: _rowOf_(i)});
    }
    return;
  }
  log_('WARN','Vegan reject: 행 없음',{ subject: key.subject });
}

function applyBaseFields_(sheet, layout, rowIndex, { subject, sendDateStr, media, material, cost }) {
  if (subject !== undefined)     sheet.getRange(rowIndex, layout.B_subject).setValue(subject || '');
  if (sendDateStr !== undefined) sheet.getRange(rowIndex, layout.C_date).setValue(normalizeDateString_(sendDateStr) || '');
  if (layout.linkCol === 4) {
    sheet.getRange(rowIndex, layout.E_media).setValue(media || '');
    sheet.getRange(rowIndex, layout.F_material).setValue(material || '');
    if (cost !== undefined) sheet.getRange(rowIndex, layout.G_cost).setValue(cost || '');
  } else {
    sheet.getRange(rowIndex, layout.D_media).setValue(media || '');
    sheet.getRange(rowIndex, layout.E_material).setValue(material || '');
    if (cost !== undefined) sheet.getRange(rowIndex, layout.G_cost).setValue(cost || '');
  }
  // 반려 슬롯 재사용 시 H열(반려 플래그) 초기화
  if (layout.H_flag) sheet.getRange(rowIndex, layout.H_flag).setValue('');
}

/** 실행사 시트에서 이 업무 칸을 찾아 반려 규칙 적용 (못 찾으면 아무 칸도 건드리지 않음) */
function _rejectTaskSlot_(sheetName, key, media, material) {
  const sheet = getTargetSheet_(sheetName);
  if (!sheet) { log_('ERROR','반려: 실행사 시트 열기 실패',{sheetName}); return; }
  if (sheetName === '비건뉴스') { applyVeganReject_(sheet, media, material, key); return; }
  const layout = HYB.TARGET_LAYOUT[sheetName];
  if (!layout) { log_('ERROR','반려: 레이아웃 없음',{sheetName}); return; }
  const S = _readSlots_(sheet, layout);
  const i = _findTaskSlot_(S, layout, key);
  if (i < 0) { log_('WARN','Reject skipped: 행 없음',{ sheetName, subject: key.subject, round: key.roundTrusted }); return; }
  applyRejectRule_(sheetName, layout, sheet, _rowOf_(i));
}

// ============== 최종 반려 (재배정 없음) ==============
/**
 * E열(진행상태) = '반려' 입력 시 호출
 * - M열(링크) 있으면 안전장치로 스킵
 * - 업무시트 N열(매출), O열(원가) → 0 처리
 * - J열 실행사 시트에서 이 업무 칸만 반려 처리 (재배정 없음)
 * - EXECUTOR_CACHE·ROUND 초기화 (고유번호는 유지 → 같은 실행사로 다시 오면 그 반려 칸 재사용)
 */
function handleFinalReject_({ row, sheet, subject, sendDateStr, media, material, roundNow, execNow, execCache, task }) {
  const mVal = String(sheet.getRange(row, HYB.COL.LINK).getValue() || '').trim();
  if (mVal) { log_('WARN', '최종반려 스킵: M열 링크 존재', { row }); return; }

  sheet.getRange(row, HYB.COL.SALES).setValue(0);  // N열 수식 덮어씀
  sheet.getRange(row, HYB.COL.COST).setValue(0);   // O열

  if (!execNow) { log_('INFO', '최종반려: 실행사 없음 → N/O만 0처리', { row }); return; }
  const sheetName = HYB.EXECUTOR_MAP[execNow];
  if (!sheetName) {
    log_('WARN', '최종반려: 실행사 매핑 없음', { execNow, validKeys: Object.keys(HYB.EXECUTOR_MAP).join(',') });
    return;
  }

  const roundTrusted = (execCache && execCache === execNow) ? roundNow : '';
  _rejectTaskSlot_(sheetName, { id: task.id, dup: task.dup, mainRow: row, roundTrusted, subject, sendDateStr }, media, material);

  sheet.getRange(row, HYB.COL.EXECUTOR_CACHE).setValue('');
  sheet.getRange(row, HYB.COL.ROUND).setValue('');
  log_('INFO', '최종반려 완료', { row, executor: execNow, sheetName });
}

// ============== 반려 + 이관 ==============
function handleRejectAndReassign_({ row, subject, sendDateStr, media, material, cost, link, oldExecutor, newExecutor, roundNow, task }) {
  const main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HYB.MAIN_SHEET);
  const prevName = HYB.EXECUTOR_MAP[oldExecutor];
  const nextName = HYB.EXECUTOR_MAP[newExecutor];

  // ── 1) 이전 시트에서 이 업무 칸 반려 (L열 회차는 이전 실행사 것이므로 이전 시트에서만 사용) ──
  if (prevName) {
    _rejectTaskSlot_(prevName, { id: task.id, dup: task.dup, mainRow: row, roundTrusted: roundNow, subject, sendDateStr }, media, material);
  } else {
    log_('WARN','Reject: oldExecutor 매핑 없음',{oldExecutor});
  }

  // 이전 실행사 회차는 여기서 항상 비움
  main.getRange(row, HYB.COL.ROUND).setValue('');

  // ── 2) 새 실행사가 없거나 목록에 없음 → 반려만 하고 종료 (L열·실행사기억 모두 비움) ──
  if (!nextName) {
    main.getRange(row, HYB.COL.EXECUTOR_CACHE).setValue('');
    log_('INFO','Reject only (no reassign)',{prevName});
    return;
  }

  // ── 3) 새 시트에 배정 (이 업무의 같은 구간 반려 칸이 있으면 그 자리, 없으면 새 칸) ──
  const ctx = { row, subject, sendDateStr, media, material, cost, link, task, roundTrusted: '' };
  if (nextName === '비건뉴스') handleVeganUpsert_(ctx);
  else handleBaseUpsert_({ ...ctx, targetSheetName: nextName, execName: newExecutor });

  main.getRange(row, HYB.COL.EXECUTOR_CACHE).setValue(newExecutor);
  log_('INFO','Reject+Reassign done',{from: prevName, to: nextName});
}

// ─── 반려 규칙 (시트별) ────────────────────────────────────────
function applyRejectRule_(sheetName, layout, sheet, rowIndex) {
  if (!rowIndex) { log_('WARN','Reject skipped: 행 없음',{sheetName}); return; }
  switch (sheetName) {
    case '플래나 송출':
    case '한미마 송출':
    case EPR_KEY:           // 이피알몰: 한미마와 동일 반려 규칙
    case BBW_KEY:           // 박병욱: 한미마와 동일 반려 규칙
      sheet.getRange(rowIndex, layout.G_cost).setValue(0);
      if (layout.H_flag) { try { sheet.getRange(rowIndex, layout.H_flag).setValue('반려'); } catch (e) {} }
      break;
    case '이디엘 송출':
      sheet.getRange(rowIndex, layout.G_cost).setValue('반려');
      break;
    case '공감신문':
      // ★ A열(회차)만 남기고 B~H 전부 삭제 → 새 건 배정 (고유번호 메모도 지움 → 빈 칸으로 재사용)
      const roundBackup = sheet.getRange(rowIndex, layout.A_round).getValue();
      const lastCol     = layout.H_flag || layout.G_cost;
      for (let c = layout.B_subject; c <= lastCol; c++) {
        sheet.getRange(rowIndex, c).setValue('');
      }
      sheet.getRange(rowIndex, layout.B_subject).setNote('');
      sheet.getRange(rowIndex, layout.A_round).setValue(roundBackup);
      break;
    default:
      log_('WARN','Reject: 알 수 없는 시트',{sheetName});
  }
}


// ============== 공통 유틸 ==============
function normalizeDateString_(val) {
  if (val === '' || val === null || val === undefined) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, 'Asia/Seoul', 'yyyy.MM.dd');
  }
  const s = String(val).trim();
  if (/^\d{4}\.\d{1,2}\.\d{1,2}$/.test(s)) {
    const p = s.split('.').map(x => String(x).trim()).filter(Boolean);
    if (p.length === 3) return `${p[0].padStart(4,'0')}.${p[1].padStart(2,'0')}.${p[2].padStart(2,'0')}`;
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy.MM.dd');
  return s;
}

function log_(level, msg, obj) {
  if (!HYB.LOG_ENABLED && level !== 'ERROR') return;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let s = ss.getSheetByName(HYB.LOG_SHEET);
    if (!s) {
      s = ss.insertSheet(HYB.LOG_SHEET);
      s.getRange(1,1,1,6).setValues([['ts','level','message','json','user','project']]);
    }
    s.appendRow([
      Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
      level, msg,
      obj ? JSON.stringify(obj) : '',
      Session.getActiveUser().getEmail ? Session.getActiveUser().getEmail() : '',
      HYB.PROJECT_NAME
    ]);
  } catch (e) {}
}

// ─── 설치형 트리거 등록 (최초 1회 또는 트리거 초기화 시) ─────────
function installHYBTrigger() {
  // 기존 HYB_Trigger 전부 삭제 후 재등록
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'HYB_Trigger')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('HYB_Trigger')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ HYB_Trigger 설치 완료\n' +
    '업무시트 G/I/J/K/M열 변경 시 실행사 시트 자동 동기화됩니다.\n' +
    '비건뉴스 자동 연동도 활성화되었습니다.'
  );
}

/** 현재 등록된 설치형 트리거 목록 확인 (디버깅용) */
function listInstalledTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const info = triggers.map(t => `${t.getHandlerFunction()} (${t.getEventType()})`);
  SpreadsheetApp.getUi().alert('등록된 트리거:\n' + (info.join('\n') || '없음'));
}

// ============== 회차 자동 갱신·채우기 ==============

/**
 * [2026-10-16] 빈 칸 회차 번호 관리 (실행사별 칸 수: SLOT_CFG)
 *  - reset=true (플래나 매월·이디엘 매주 월요일·한미마 충전): 마지막 사용 칸 아래 빈 칸을 1부터 다시 매기고, target 칸까지 확보
 *  - reset=false (배정할 때마다): 빈 칸 번호를 첫 빈 칸 번호부터 순서대로 정리(예전 30칸 제한으로 남은 '51,52…' 같은 튀는 번호 바로잡음),
 *    남은 칸이 SLOT_LOW 이하이면 target 칸만큼 이어 붙임
 *  - 공감신문·비건뉴스는 대상 아님 (회차 번호 건드리지 않음)
 *  - 사용된 칸(B열 주제 있음)은 절대 건드리지 않음. 시트 줄이 모자라면 줄 추가
 */
function _ensureSlots_(sheet, sheetName, layout, reset) {
  const cfg = SLOT_CFG[sheetName];
  if (!cfg || !sheet || !layout) return;
  const aCol = layout.A_round, bCol = layout.B_subject;
  const last = Math.max(sheet.getLastRow(), DATA_START_ROW);
  const vals = sheet.getRange(DATA_START_ROW, 1, last - DATA_START_ROW + 1, Math.max(aCol, bCol)).getValues();

  let lastUsed = -1;
  for (let i = vals.length - 1; i >= 0; i--) if (String(vals[i][bCol - 1] || '').trim()) { lastUsed = i; break; }
  const tailStart = lastUsed + 1;                       // vals 인덱스
  const tailCount = vals.length - tailStart;
  const num = v => { const n = parseInt(String(v || '').replace(/[^\d]/g, ''), 10); return isNaN(n) ? 0 : n; };
  const lastUsedRound = lastUsed >= 0 ? num(vals[lastUsed][aCol - 1]) : 0;

  let n0;
  if (reset) n0 = 1;
  else n0 = (tailCount > 0 && num(vals[tailStart][aCol - 1])) || (lastUsedRound + 1);

  let want = tailCount;
  if (reset) want = Math.max(tailCount, cfg.target);
  else if (tailCount <= SLOT_LOW) want = tailCount + cfg.target;
  if (cfg.cap) want = Math.min(want, Math.max(0, cfg.cap - n0 + 1));
  if (want <= 0) return;

  const firstRow = DATA_START_ROW + tailStart;
  const needLast = firstRow + want - 1;
  const maxRows  = sheet.getMaxRows();
  if (needLast > maxRows) sheet.insertRowsAfter(maxRows, needLast - maxRows);

  const nums = [];
  for (let k = 0; k < want; k++) nums.push([n0 + k]);
  const cur = sheet.getRange(firstRow, aCol, want, 1).getValues();
  const changed = cur.some((r, k) => num(r[0]) !== nums[k][0] || String(r[0]).trim() === '');
  if (changed) sheet.getRange(firstRow, aCol, want, 1).setValues(nums);

  if (changed || reset) log_('INFO', '_ensureSlots_', { sheetName, reset, firstRow, from: n0, to: n0 + want - 1, before: tailCount });
}

/** (예전 함수 — 호환용) 빈 칸 회차 재번호화 → _ensureSlots_ 로 대체 */
function _renumberEmptySlots_(sheet, layout, startRow) {
  const name = Object.keys(HYB.TARGET_LAYOUT).find(k => HYB.TARGET_LAYOUT[k] === layout) || '';
  _ensureSlots_(sheet, name, layout, true);
}

/**
 * 플래나 월 리셋 체크:
 * 마지막으로 배정된 행의 날짜가 현재 월과 다르면 회차 갱신(1부터 + 150칸 확보).
 */
function _checkPlannerMonthReset_(sheet, layout) {
  const now      = new Date();
  const nowYear  = now.getFullYear();
  const nowMonth = now.getMonth() + 1;   // 1~12

  const last = sheet.getLastRow();
  if (last < 2) return;

  const aCol = layout.A_round;
  const bCol = layout.B_subject;
  const cCol = layout.C_date;
  const maxC = Math.max(aCol, bCol, cCol);
  const vals = sheet.getRange(2, 1, last - 1, maxC).getValues();

  // 마지막으로 사용된 행(B열 있음)의 날짜 확인
  let lastYear = null, lastMonth = null;
  for (let i = vals.length - 1; i >= 0; i--) {
    const sub     = String(vals[i][bCol - 1] || '').trim();
    const dateVal = vals[i][cCol - 1];
    if (sub && dateVal) {
      const normalized = normalizeDateString_(dateVal);
      const parts = (normalized || '').split('.');
      if (parts.length === 3 && !isNaN(parseInt(parts[0], 10))) {
        lastYear  = parseInt(parts[0], 10);
        lastMonth = parseInt(parts[1], 10);
        break;
      }
    }
  }

  const needReset =
    lastYear === null ||
    lastYear !== nowYear ||
    lastMonth !== nowMonth;

  if (needReset) {
    _ensureSlots_(sheet, '플래나 송출', layout, true);
    log_('INFO', '플래나 월 리셋', { nowYear, nowMonth, lastYear, lastMonth });
  }
}

/**
 * 이디엘 주간 리셋 — 매주 월요일 오전 6시 트리거로 실행.
 * 이디엘 송출 시트의 빈 칸 회차를 1부터 + 15칸 확보.
 */
function resetIdelWeekly_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('이디엘 송출');
  if (!sheet) {
    log_('WARN', 'resetIdelWeekly_: 이디엘 송출 시트 없음');
    return;
  }
  _ensureSlots_(sheet, '이디엘 송출', HYB.TARGET_LAYOUT['이디엘 송출'], true);
  log_('INFO', '이디엘 주간 리셋 완료');
}

/**
 * 이디엘 주간 리셋 트리거 설치 (최초 1회).
 * 매주 월요일 오전 6시에 resetIdelWeekly_ 실행.
 */
function installIdelWeeklyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'resetIdelWeekly_')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('resetIdelWeekly_')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(6)
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ 이디엘 주간 리셋 트리거 설치 완료!\n' +
    '매주 월요일 오전 6시에 이디엘 송출 빈 슬롯 회차가 1부터 자동 갱신됩니다.'
  );
}
