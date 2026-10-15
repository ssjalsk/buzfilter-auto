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
    EXECUTOR_CACHE: 27 // AA 실행사_캐시
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
  LOCK_WAIT_MS: 20000
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

// ============== 설치형 트리거 진입 ==============
function HYB_Trigger(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();

    // ── 한미마 K열 충전 감지 → 빈 슬롯 재번호화 ──────────────
    if (sheet.getName() === '한미마 송출') {
      const col = e.range.getColumn();
      const row = e.range.getRow();
      if (col === 11 && row >= 2) {          // K열(11) = 충전 금액
        const val = Number(String(e.range.getValue() || '').replace(/[,\s]/g, ''));
        if (val > 0) {
          _renumberEmptySlots_(sheet, HYB.TARGET_LAYOUT['한미마 송출'], row + 1);
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

    // I/J/K 변경은 5초 안정화 후 처리
    const isIJK = [HYB.COL.MEDIA, HYB.COL.EXECUTOR, HYB.COL.MATERIAL].includes(col);
    if (isIJK) Utilities.sleep(HYB.STABILIZE_DELAY_MS);

    const lock = LockService.getScriptLock();
    lock.waitLock(HYB.LOCK_WAIT_MS);

    // 값 읽기
    const subject     = String(sheet.getRange(row, HYB.COL.SUBJECT).getValue()       || '').trim();
    const sendDate    = sheet.getRange(row, HYB.COL.SEND_DATE).getValue();
    const sendDateStr = normalizeDateString_(sendDate);
    const media       = String(sheet.getRange(row, HYB.COL.MEDIA).getValue()          || '').trim();
    const execNow     = String(sheet.getRange(row, HYB.COL.EXECUTOR).getValue()       || '').trim();
    const material    = String(sheet.getRange(row, HYB.COL.MATERIAL).getValue()       || '').trim();
    const link        = String(sheet.getRange(row, HYB.COL.LINK).getValue()           || '').trim();
    const cost        = sheet.getRange(row, HYB.COL.COST).getValue();
    const roundNow    = String(sheet.getRange(row, HYB.COL.ROUND).getValue()          || '').trim();
    const execCache   = String(sheet.getRange(row, HYB.COL.EXECUTOR_CACHE).getValue() || '').trim();

    const targetSheetName = HYB.EXECUTOR_MAP[execNow] || '';
    const isVegan         = (targetSheetName === '비건뉴스');

    // 0) E열(진행상태) = '반려' → 최종 반려 처리 (재배정 없음)
    if (col === HYB.COL.STATUS) {
      const statusVal = String(sheet.getRange(row, HYB.COL.STATUS).getValue() || '').trim();
      if (statusVal === '반려') {
        handleFinalReject_({ row, sheet, subject, sendDateStr, media, material, roundNow, execNow });
      }
      lock.releaseLock();
      return;
    }

    // 1) 실행사 변경 감지 → 반려 이벤트
    const isExecChanged = !!execCache && execCache !== execNow;
    if (isExecChanged) {
      // 이전 미디어 정보를 캐시에서 읽을 수 없으므로 현재 media/material 사용
      // (반려 시 이전 섹션 탐색은 회차로 먼저 시도)
      handleRejectAndReassign_({
        row, subject, sendDateStr, media, material, cost, link,
        oldExecutor: execCache,
        newExecutor: execNow,
        roundNow
      });
      lock.releaseLock();
      return;
    }

    // 2) 링크만 편집
    if (col === HYB.COL.LINK) {
      if (!targetSheetName) { log_('WARN','Link edit skipped: executor not set',{row}); lock.releaseLock(); return; }
      if (isVegan) {
        handleVeganLinkUpsert_({ row, subject, sendDateStr, link, media, material });
      } else {
        handleLinkUpsert_({ row, subject, sendDateStr, link, targetSheetName });
      }
      lock.releaseLock();
      return;
    }

    // 3) 기본 업서트(G/I/J/K)
    if (!targetSheetName) { log_('WARN','Base upsert skipped: executor not set',{row}); lock.releaseLock(); return; }
    if (isVegan) {
      handleVeganUpsert_({ row, subject, sendDateStr, media, material });
    } else {
      handleBaseUpsert_({ row, subject, sendDateStr, media, material, cost, targetSheetName });
    }
    lock.releaseLock();

  } catch (err) {
    log_('ERROR','HYB_Trigger failed',{message: err && err.message, stack: err && err.stack});
    try { LockService.getScriptLock().releaseLock(); } catch (e) {}
  }
}

// ============== 기본/링크 업서트 (기존 시트) ==============
function handleBaseUpsert_({ row, subject, sendDateStr, media, material, cost, targetSheetName }) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const main   = ss.getSheetByName(HYB.MAIN_SHEET);
  const target = getTargetSheet_(targetSheetName);   // ← 외부 시트 지원
  const layout = HYB.TARGET_LAYOUT[targetSheetName];
  if (!target || !layout) { log_('ERROR','Target missing',{targetSheetName}); return; }

  // ── 플래나 월 리셋 체크 ──────────────────────────────────────
  if (targetSheetName === '플래나 송출') {
    _checkPlannerMonthReset_(target, layout);
  }

  const resolved = resolveRowForBase_(target, layout, {
    roundNow: String(main.getRange(row, HYB.COL.ROUND).getValue() || '').trim(),
    subject, sendDateStr
  });
  if (!resolved) { log_('INFO','Base upsert: no row to update',{targetSheetName}); return; }

  applyBaseFields_(target, layout, resolved, { subject, sendDateStr, media, material, cost });

  const roundVal = String(target.getRange(resolved, layout.A_round).getValue() || '').trim();
  main.getRange(row, HYB.COL.ROUND).setValue(roundVal);
  // ★ 버그수정: getActiveSheet() 대신 main으로 고정
  const execNow = main.getRange(row, HYB.COL.EXECUTOR).getValue();
  if (execNow) main.getRange(row, HYB.COL.EXECUTOR_CACHE).setValue(execNow);
}

function handleLinkUpsert_({ row, subject, sendDateStr, link, targetSheetName }) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const main   = ss.getSheetByName(HYB.MAIN_SHEET);
  const target = getTargetSheet_(targetSheetName);   // ← 외부 시트 지원
  const layout = HYB.TARGET_LAYOUT[targetSheetName];
  if (!target || !layout) { log_('ERROR','Target missing',{targetSheetName}); return; }

  const resolved = resolveRowForBase_(target, layout, {
    roundNow: String(main.getRange(row, HYB.COL.ROUND).getValue() || '').trim(),
    subject, sendDateStr
  });
  if (!resolved) { log_('INFO','Link upsert: no row to update',{targetSheetName}); return; }

  target.getRange(resolved, layout.linkCol).setValue(link || '');

  const roundVal = String(target.getRange(resolved, layout.A_round).getValue() || '').trim();
  main.getRange(row, HYB.COL.ROUND).setValue(roundVal);
  const execNow = main.getRange(row, HYB.COL.EXECUTOR).getValue();
  if (execNow) main.getRange(row, HYB.COL.EXECUTOR_CACHE).setValue(execNow);
}

// ============== 비건뉴스 전용 업서트 ==============
function handleVeganUpsert_({ row, subject, sendDateStr, media, material }) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const main  = ss.getSheetByName(HYB.MAIN_SHEET);
  const vSheet= ss.getSheetByName('비건뉴스');
  if (!vSheet) { log_('ERROR','비건뉴스 시트 없음'); return; }

  const section   = getVeganSection_(media, material);
  const layout    = VEGAN_SECTION_LAYOUT[section];
  const roundNow  = String(main.getRange(row, HYB.COL.ROUND).getValue() || '').trim();
  const targetRow = resolveVeganRow_(vSheet, layout, roundNow, subject, sendDateStr);

  if (!targetRow) {
    log_('WARN','Vegan upsert: 빈 슬롯 없음',{section}); return;
  }

  vSheet.getRange(targetRow, layout.B_subject).setValue(subject || '');
  vSheet.getRange(targetRow, layout.C_date).setValue(normalizeDateString_(sendDateStr) || '');

  // 반려 슬롯 재사용 시 E열(반려 플래그) → '대체완료'로 변경
  const eFlag = String(vSheet.getRange(targetRow, layout.E_flag).getValue() || '').trim();
  if (eFlag === '반려') {
    vSheet.getRange(targetRow, layout.E_flag).setValue('대체완료');
  }

  // 회차 역동기화 (L열)
  const roundVal = String(vSheet.getRange(targetRow, layout.A_round).getValue() || '').trim();
  if (roundVal) main.getRange(row, HYB.COL.ROUND).setValue(roundVal);
  main.getRange(row, HYB.COL.EXECUTOR_CACHE).setValue(
    main.getRange(row, HYB.COL.EXECUTOR).getValue()
  );

  log_('INFO','Vegan upsert OK',{section, targetRow, roundVal, eFlag});
}

function handleVeganLinkUpsert_({ row, subject, sendDateStr, link, media, material }) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const main  = ss.getSheetByName(HYB.MAIN_SHEET);
  const vSheet= ss.getSheetByName('비건뉴스');
  if (!vSheet) { log_('ERROR','비건뉴스 시트 없음'); return; }

  const section   = getVeganSection_(media, material);
  const layout    = VEGAN_SECTION_LAYOUT[section];
  const roundNow  = String(main.getRange(row, HYB.COL.ROUND).getValue() || '').trim();
  const targetRow = resolveVeganRow_(vSheet, layout, roundNow, subject, sendDateStr);

  if (!targetRow) {
    log_('WARN','Vegan link upsert: 행 없음',{section}); return;
  }

  vSheet.getRange(targetRow, layout.D_link).setValue(link || '');

  const roundVal = String(vSheet.getRange(targetRow, layout.A_round).getValue() || '').trim();
  if (roundVal) main.getRange(row, HYB.COL.ROUND).setValue(roundVal);
  const execNow = main.getRange(row, HYB.COL.EXECUTOR).getValue();
  if (execNow) main.getRange(row, HYB.COL.EXECUTOR_CACHE).setValue(execNow);

  log_('INFO','Vegan link upsert OK',{section, targetRow});
}

/** 비건뉴스 행 탐색:
 *  1) A열(회차) 일치
 *  2) 반려 슬롯 우선 (E열='반려') ← 수량 보전을 위해 반려 회차 재활용
 *  3) 주제+날짜 일치
 *  4) 빈 슬롯 (회차 있고 주제 없음)
 *  5) 완전 빈 행 (회차도 주제도 없음)
 */
function resolveVeganRow_(sheet, layout, roundNow, subject, sendDateStr) {
  const last   = sheet.getLastRow(); if (last < 2) return null;
  const maxCol = Math.max(layout.A_round, layout.B_subject, layout.C_date, layout.D_link, layout.E_flag);
  const vals   = sheet.getRange(2, 1, last - 1, maxCol).getValues();

  // 1) 회차로 찾기 (기존 회차가 지정된 경우 최우선) — B열 주제 있는 슬롯만 선택
  if (roundNow) {
    for (let i = vals.length - 1; i >= 0; i--) {
      const round = String(vals[i][layout.A_round - 1]);
      const sub   = String(vals[i][layout.B_subject - 1] || '').trim();
      if (round === String(roundNow) && sub) return i + 2;
    }
  }
  // 2) 반려 슬롯 우선: 같은 주제의 반려 슬롯 먼저 → 없으면 아무 반려 슬롯
  // ★ 버그수정: 주제 미매칭으로 다른 건 반려 슬롯에 덮어쓰는 오탐 방지
  for (let i = 0; i < vals.length; i++) {
    const sub  = String(vals[i][layout.B_subject - 1] || '').trim();
    const flag = String(vals[i][layout.E_flag - 1]    || '').trim();
    if (flag === '반려' && sub === subject) return i + 2;
  }
  // 같은 주제 반려 슬롯 없을 때만 아무 반려 슬롯 (기존 동작 보존)
  for (let i = 0; i < vals.length; i++) {
    const flag = String(vals[i][layout.E_flag - 1] || '').trim();
    if (flag === '반려') return i + 2;
  }
  // 3) 주제+날짜 일치
  const d = normalizeDateString_(sendDateStr);
  for (let i = 0; i < vals.length; i++) {
    const sub = String(vals[i][layout.B_subject - 1] || '').trim();
    const dt  = normalizeDateString_(vals[i][layout.C_date - 1]);
    if (sub === subject && dt === d) return i + 2;
  }
  // 4) 빈 슬롯: 회차 있고 주제 없는 행
  for (let i = 0; i < vals.length; i++) {
    const round = String(vals[i][layout.A_round - 1] || '').trim();
    const sub   = String(vals[i][layout.B_subject - 1] || '').trim();
    if (round && !sub) return i + 2;
  }
  // 5) 완전 빈 행: 회차도 주제도 없는 행 (새로 추가된 빈 행)
  for (let i = 0; i < vals.length; i++) {
    const round = String(vals[i][layout.A_round - 1] || '').trim();
    const sub   = String(vals[i][layout.B_subject - 1] || '').trim();
    if (!round && !sub) return i + 2;
  }
  return null;
}

/** 비건뉴스 반려 적용: 소재 기준으로 섹션 특정 후 해당 섹션 E열에 '반려' 표기 */
function applyVeganReject_(vSheet, media, material, roundNow, subject, sendDateStr) {
  // handleVeganUpsert_ 와 동일하게 소재로 섹션 결정 → 섹션 간 오탐 방지
  const section  = getVeganSection_(media, material);
  const layout   = VEGAN_SECTION_LAYOUT[section];
  const targetRow = resolveVeganRow_(vSheet, layout, roundNow, subject, sendDateStr);
  if (!targetRow) {
    log_('WARN','Vegan reject: 행 없음',{section, roundNow, subject});
    return;
  }
  const flagVal = String(vSheet.getRange(targetRow, layout.E_flag).getValue() || '').trim();
  if (flagVal !== '대체완료') {  // 이미 완료된 건은 건드리지 않음
    vSheet.getRange(targetRow, layout.E_flag).setValue('반려');
    log_('INFO','Vegan reject OK',{section, targetRow});
  }
}

// 기본/링크에서 사용할 대상행 탐색
function resolveRowForBase_(sheet, layout, { roundNow, subject, sendDateStr }) {
  const last   = sheet.getLastRow(); if (last < 2) return null;
  const maxCol = Math.max(layout.G_cost, layout.H_flag || layout.G_cost);
  const vals   = sheet.getRange(2, 1, last - 1, maxCol).getValues();

  // 1) 회차로 찾기 — B열에 주제가 있는 슬롯만 (빈 슬롯을 회차로 잘못 선택 방지)
  // 회차 탐색은 "이미 배정된 슬롯 업데이트" 용도. 빈 슬롯은 step 3에서 처리.
  if (roundNow) {
    for (let i = vals.length - 1; i >= 0; i--) {
      const round = String(vals[i][layout.A_round - 1]);
      const sub   = String(vals[i][layout.B_subject - 1] || '').trim();
      if (round === String(roundNow) && sub) return i + 2;
    }
  }
  // 2) (주제+날짜) 완전일치
  const d = normalizeDateString_(sendDateStr);
  const exact = [];
  for (let i = 0; i < vals.length; i++) {
    const sub = String(vals[i][layout.B_subject - 1] || '').trim();
    const dt  = normalizeDateString_(vals[i][layout.C_date - 1]);
    if (sub === subject && dt === d) exact.push(i + 2);
  }
  if (exact.length) return exact[exact.length - 1];

  // 3) 빈 슬롯(B empty)
  for (let i = 0; i < vals.length; i++) {
    const sub = String(vals[i][layout.B_subject - 1] || '').trim();
    if (!sub) return i + 2;
  }
  return null;
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

// ============== 최종 반려 (재배정 없음) ==============
/**
 * E열(진행상태) = '반려' 입력 시 호출
 * - M열(링크) 있으면 안전장치로 스킵
 * - 업무시트 N열(매출), O열(원가) → 0 처리
 * - J열 실행사 시트에서 해당 슬롯 반려 처리 (재배정 없음)
 * - EXECUTOR_CACHE 초기화
 */
function handleFinalReject_({ row, sheet, subject, sendDateStr, media, material, roundNow, execNow }) {
  // 1. M열(링크) 체크 — 링크 있으면 이미 기사 나간 상태 → 스킵
  const mVal = String(sheet.getRange(row, HYB.COL.LINK).getValue() || '').trim();
  if (mVal) {
    log_('WARN', '최종반려 스킵: M열 링크 존재', { row });
    return;
  }

  // 2. 업무시트 N열(매출) = 0, O열(원가) = 0
  sheet.getRange(row, HYB.COL.SALES).setValue(0);  // N열 수식 덮어씀
  sheet.getRange(row, HYB.COL.COST).setValue(0);   // O열

  // 3. 실행사 없으면 N/O 처리만 하고 종료
  if (!execNow) {
    log_('INFO', '최종반려: 실행사 없음 → N/O만 0처리', { row });
    return;
  }

  // 4. 실행사 시트 반려 처리 (기존 규칙 그대로, 재배정 없음)
  const sheetName = HYB.EXECUTOR_MAP[execNow];
  if (!sheetName) {
    log_('WARN', '최종반려: 실행사 매핑 없음', { execNow, validKeys: Object.keys(HYB.EXECUTOR_MAP).join(',') });
    return;
  }

  log_('INFO', '최종반려: 실행사 시트 탐색 시작', { execNow, sheetName, roundNow, subject, sendDateStr });

  if (sheetName === '비건뉴스') {
    const vSheet = getTargetSheet_(sheetName);
    if (vSheet) applyVeganReject_(vSheet, media, material, roundNow, subject, sendDateStr);
  } else {
    const targetSheet  = getTargetSheet_(sheetName);
    const targetLayout = HYB.TARGET_LAYOUT[sheetName];
    if (!targetSheet) { log_('ERROR', '최종반려: 실행사 시트 열기 실패', { sheetName }); return; }
    if (!targetLayout) { log_('ERROR', '최종반려: 레이아웃 없음', { sheetName }); return; }
    const targetRow = resolveRowForRejectPrev_(targetSheet, targetLayout, { roundNow, subject, sendDateStr }, sheetName);
    log_('INFO', '최종반려: targetRow 탐색 결과', { sheetName, targetRow });
    applyRejectRule_(sheetName, targetLayout, targetSheet, targetRow);
  }

  // 5. EXECUTOR_CACHE + ROUND 초기화
  // ROUND도 함께 초기화: 이후 J열 실행사 변경 시 이전 회차로 잘못 탐색하는 버그 방지
  sheet.getRange(row, HYB.COL.EXECUTOR_CACHE).setValue('');
  sheet.getRange(row, HYB.COL.ROUND).setValue('');

  log_('INFO', '최종반려 완료', { row, executor: execNow, sheetName });
}

// ============== 반려 + 이관 ==============
function handleRejectAndReassign_({ row, subject, sendDateStr, media, material, cost, link, oldExecutor, newExecutor, roundNow }) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const main = ss.getSheetByName(HYB.MAIN_SHEET);

  const prevName = HYB.EXECUTOR_MAP[oldExecutor];
  const nextName = HYB.EXECUTOR_MAP[newExecutor];
  if (!prevName) { log_('WARN','Reject: oldExecutor 매핑 없음',{oldExecutor}); return; }

  const prevSheet  = getTargetSheet_(prevName);      // ← 외부 시트 지원
  const prevLayout = HYB.TARGET_LAYOUT[prevName];  // 비건뉴스면 undefined

  // ── 1) 이전 시트에서 반려 적용 ──────────────────────────────
  if (prevName === '비건뉴스') {
    if (prevSheet) applyVeganReject_(prevSheet, media, material, roundNow, subject, sendDateStr);
  } else {
    if (prevSheet && prevLayout) {
      const prevRow = resolveRowForRejectPrev_(prevSheet, prevLayout, { roundNow, subject, sendDateStr }, prevName);
      applyRejectRule_(prevName, prevLayout, prevSheet, prevRow);
    }
  }

  // ── 2) 새 시트 배정이 없으면 (실행사 제거) 종료 ─────────────
  if (!nextName) {
    // 캐시 초기화
    main.getRange(row, HYB.COL.EXECUTOR_CACHE).setValue('');
    log_('INFO','Reject only (no reassign)',{prevName});
    return;
  }

  // ── 3) 새 시트에 배정 ────────────────────────────────────────
  // 이전 실행사 회차 잔존 방지: 새 실행사 배정 전 L열 초기화
  main.getRange(row, HYB.COL.ROUND).setValue('');
  if (nextName === '비건뉴스') {
    // 비건뉴스로 이관: 섹션 결정 후 빈 슬롯에 배정
    handleVeganUpsert_({ row, subject, sendDateStr, media, material });
    if (link) handleVeganLinkUpsert_({ row, subject, sendDateStr, link, media, material });
  } else {
    const nextSheet  = getTargetSheet_(nextName);   // ← 외부 시트 지원
    const nextLayout = HYB.TARGET_LAYOUT[nextName];
    if (!nextSheet || !nextLayout) { log_('ERROR','Next sheet/layout missing',{nextName}); return; }

    const nextRow = resolveRowForAssignNext_(nextSheet, nextLayout, { subject, sendDateStr });
    if (!nextRow) { log_('INFO','Next: 빈 슬롯 없음',{nextName}); return; }

    applyBaseFields_(nextSheet, nextLayout, nextRow, { subject, sendDateStr, media, material, cost });
    // 반려 슬롯 재활용 시 H열 반려 표시 초기화
    if (nextLayout.H_flag) nextSheet.getRange(nextRow, nextLayout.H_flag).setValue('');
    if (link) nextSheet.getRange(nextRow, nextLayout.linkCol).setValue(link);

    // 회차 동기화
    // ★ 버그수정: newRound가 없어도 L열을 반드시 업데이트 → 이전 실행사 회차 잔존 방지
    const newRound = String(nextSheet.getRange(nextRow, nextLayout.A_round).getValue() || '').trim();
    main.getRange(row, HYB.COL.ROUND).setValue(newRound);
  }

  main.getRange(row, HYB.COL.EXECUTOR_CACHE).setValue(newExecutor);
  log_('INFO','Reject+Reassign done',{from: prevName, to: nextName});
}

// 이전 시트: 회차 → (주제+날짜) → 주제만(폴백)
// sheetName: 플래나 송출/한미마 송출/이디엘 송출은 날짜 기준 최신 행 선택
function resolveRowForRejectPrev_(sheet, layout, { roundNow, subject, sendDateStr }, sheetName) {
  const last   = sheet.getLastRow(); if (last < 2) return null;
  const maxCol = Math.max(layout.G_cost, layout.H_flag || layout.G_cost);
  const vals   = sheet.getRange(2, 1, last - 1, maxCol).getValues();

  const useDateSort = ['플래나 송출', '한미마 송출', '이디엘 송출'].includes(sheetName);

  // 1) 회차 일치 + B열(주제) 있는 행만 — 빈 슬롯 잘못 반려 방지
  if (roundNow) {
    const cands = [];
    for (let i = 0; i < vals.length; i++) {
      const sub = String(vals[i][layout.B_subject - 1] || '').trim();
      if (String(vals[i][layout.A_round - 1]) === String(roundNow) && sub) cands.push(i + 2);
    }
    if (cands.length) {
      return useDateSort ? _latestByDate_(cands, vals, layout.C_date) : cands[cands.length - 1];
    }
    log_('INFO', 'resolveRowForRejectPrev_: 회차 탐색 실패', { roundNow, subject });
  }

  // 2) 주제+날짜 완전일치
  const d = normalizeDateString_(sendDateStr);
  const exact = [];
  for (let i = 0; i < vals.length; i++) {
    const sub = String(vals[i][layout.B_subject - 1] || '').trim();
    const dt  = normalizeDateString_(vals[i][layout.C_date - 1]);
    if (sub === subject && dt === d) exact.push(i + 2);
  }
  if (exact.length) return exact[exact.length - 1];

  // 3) 폴백: 주제만 일치 — 플래나/한미마/이디엘은 날짜 기준 최신, 나머지는 마지막 행
  const subOnly = [];
  for (let i = 0; i < vals.length; i++) {
    const sub = String(vals[i][layout.B_subject - 1] || '').trim();
    if (sub === subject) subOnly.push(i + 2);
  }
  if (subOnly.length) {
    const result = useDateSort ? _latestByDate_(subOnly, vals, layout.C_date) : subOnly[subOnly.length - 1];
    log_('INFO', 'resolveRowForRejectPrev_: 주제만 매칭(폴백)', { subject, matchedRow: result });
    return result;
  }

  log_('WARN', 'resolveRowForRejectPrev_: 행 없음', { roundNow, subject, sendDateStr });
  return null;
}

// C열(날짜) 기준 가장 최근 행 반환
function _latestByDate_(rowNums, vals, cDateCol) {
  let best = rowNums[0];
  let bestDate = normalizeDateString_(vals[rowNums[0] - 2][cDateCol - 1]);
  for (let i = 1; i < rowNums.length; i++) {
    const dt = normalizeDateString_(vals[rowNums[i] - 2][cDateCol - 1]);
    if (dt > bestDate) { bestDate = dt; best = rowNums[i]; }
  }
  return best;
}

// 새 시트: 반려 슬롯 우선 → 빈 슬롯 → (주제+날짜)
function resolveRowForAssignNext_(sheet, layout, { subject, sendDateStr }) {
  const last   = sheet.getLastRow(); if (last < 2) return null;
  const maxCol = Math.max(layout.G_cost, layout.H_flag || layout.G_cost);
  const vals   = sheet.getRange(2, 1, last - 1, maxCol).getValues();

  // 1) 같은 subject의 반려된 슬롯 우선
  if (layout.H_flag) {
    for (let i = 0; i < vals.length; i++) {
      const sub  = String(vals[i][layout.B_subject - 1] || '').trim();
      const flag = String(vals[i][layout.H_flag - 1]    || '').trim();
      if (sub === subject && flag === '반려') return i + 2;
    }
  }
  for (let i = 0; i < vals.length; i++) {
    const sub  = String(vals[i][layout.B_subject - 1] || '').trim();
    const cost = String(vals[i][layout.G_cost - 1]    || '').trim();
    if (sub === subject && cost === '반려') return i + 2;
  }
  // 2) 빈 슬롯
  for (let i = 0; i < vals.length; i++) {
    const sub = String(vals[i][layout.B_subject - 1] || '').trim();
    if (!sub) return i + 2;
  }
  // 3) (주제+날짜) 완전일치
  const d = normalizeDateString_(sendDateStr);
  const exact = [];
  for (let i = 0; i < vals.length; i++) {
    const sub = String(vals[i][layout.B_subject - 1] || '').trim();
    const dt  = normalizeDateString_(vals[i][layout.C_date - 1]);
    if (sub === subject && dt === d) exact.push(i + 2);
  }
  if (exact.length) return exact[exact.length - 1];
  return null;
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
      // ★ A열(회차)만 남기고 B~H 전부 삭제 → 새 건 배정
      const roundBackup = sheet.getRange(rowIndex, layout.A_round).getValue();
      const lastCol     = layout.H_flag || layout.G_cost;
      for (let c = layout.B_subject; c <= lastCol; c++) {
        sheet.getRange(rowIndex, c).setValue('');
      }
      // A열은 반드시 유지 (혹시 지워졌을 경우 복원)
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

// ============== 회차 자동 리셋 ==============

/**
 * 빈 슬롯(B_subject 비어있는 행)의 A_round를 1부터 재번호화.
 * 이미 사용된 행(B_subject 값 있음)은 절대 건드리지 않음.
 *
 * @param {Sheet}   sheet      - 대상 시트
 * @param {object}  layout     - TARGET_LAYOUT 항목
 * @param {number}  [startRow=2] - 이 행(포함)부터 탐색 시작 (한미마 충전 시 충전행+1)
 */
function _renumberEmptySlots_(sheet, layout, startRow) {
  const last = sheet.getLastRow();
  const aCol = layout.A_round;
  const bCol = layout.B_subject;
  const from = startRow || 2;
  if (last < from) return;

  const count = last - from + 1;
  const vals  = sheet.getRange(from, 1, count, Math.max(aCol, bCol)).getValues();

  let counter = 0;
  const updates = [];
  for (let i = 0; i < vals.length; i++) {
    const sub = String(vals[i][bCol - 1] || '').trim();
    if (!sub) {
      counter++;
      updates.push({ rowIndex: from + i, newNum: counter });
    }
  }

  // 일괄 쓰기
  for (const { rowIndex, newNum } of updates) {
    sheet.getRange(rowIndex, aCol).setValue(newNum);
  }

  log_('INFO', '_renumberEmptySlots_', { from, totalUpdated: updates.length, lastNum: counter });
}

/**
 * 플래나 월 리셋 체크:
 * 마지막으로 배정된 행의 날짜가 현재 월과 다르면 빈 슬롯 재번호화.
 * 아직 배정된 행이 없을 때도 재번호화.
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
  // ★ 버그수정: new Date('yyyy.MM.dd') 는 파싱 실패 → normalizeDateString_ 활용으로 교체
  let lastYear = null, lastMonth = null;
  for (let i = vals.length - 1; i >= 0; i--) {
    const sub     = String(vals[i][bCol - 1] || '').trim();
    const dateVal = vals[i][cCol - 1];
    if (sub && dateVal) {
      const normalized = normalizeDateString_(dateVal);
      // normalizeDateString_ 반환값은 'yyyy.MM.dd' 형식이므로 직접 분리
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
    _renumberEmptySlots_(sheet, layout, 2);
    log_('INFO', '플래나 월 리셋', { nowYear, nowMonth, lastYear, lastMonth });
  }
}

/**
 * 이디엘 주간 리셋 — 매주 월요일 오전 6시 트리거로 실행.
 * 이디엘 송출 시트의 빈 슬롯 A열을 1부터 재번호화.
 */
function resetIdelWeekly_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('이디엘 송출');
  if (!sheet) {
    log_('WARN', 'resetIdelWeekly_: 이디엘 송출 시트 없음');
    return;
  }
  const layout = HYB.TARGET_LAYOUT['이디엘 송출'];
  _renumberEmptySlots_(sheet, layout, 2);
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
