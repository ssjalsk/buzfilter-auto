/**
 * 리뷰 업무시트 — 드롭다운 + I·J열 자동채움
 *
 * [드롭다운]
 *   C열(플랫폼)   ← '리뷰실행사' B열 중복제거
 *   F열(상품분류) ← '리뷰실행사' C열 중복제거
 *   G열(실행사)   ← '리뷰실행사' A열 중복제거
 *
 * [자동채움]  C+F+G 일치 시
 *   I열(실행비) ← '리뷰실행사' D열
 *   J열(택배비) ← '리뷰실행사' E열
 *
 * [지출예정금액 L열]
 *   수식 =(H+I+J)*K 으로 직접 입력 → 트리거가 건드리지 않음
 */

const REVIEW_CONF = {
  WORK_SHEET: '리뷰 업무시트',
  EXEC_SHEET: '리뷰실행사',
  DATA_ROW:   2,

  // ── 리뷰 업무시트 열 ──────────────────────────────────────────
  COL: {
    STATUS:    4,   // D 진행상태
    PLATFORM:  3,   // C 플랫폼
    CATEGORY:  6,   // F 상품분류
    EXECUTOR:  7,   // G 실행사
    EXEC_FEE:  9,   // I 실행비  (자동채움)
    DELIVERY:  10,  // J 택배비  (자동채움)
    COST:      12,  // L 지출예정금액
  },

  STATUS_ACTIVE: '진행 중',         // (참고용) 지출 시작 단계
  STATUS_REQUEST: '리뷰 요청',      // [2026-10-07] 이 상태일 때만 L열 = 수식, 그 외 상태는 L열 = 0
  SETTLE_SHEET:  '종합 정산시트',
  SETTLE_BUDGET_CELL: 'P5',         // 리뷰 지출예정 합계 셀

  // ── 리뷰실행사 시트 열 ────────────────────────────────────────
  EXEC_COL: {
    EXECUTOR:  1,   // A 실행사
    PLATFORM:  2,   // B 플랫폼
    CATEGORY:  3,   // C 상품분류
    EXEC_FEE:  4,   // D 실행비
    DELIVERY:  5,   // E 택배비
  }
};

/** 리뷰 매출(Q) 입력 시 선충전 자동 차감 — 선충전잔액 정리 전까지 false (REVIEW_Trigger 참고) */
const REVIEW_PREPAID_ON_ENTRY = false;

/* ───── 유틸 ──────────────────────────────────────────────────── */
function _rv_(v){ return (v===null||v===undefined)?'':String(v).trim(); }

function _uniqueList_(sheet, col, startRow){
  const last = sheet.getLastRow();
  if(last < startRow) return [];
  const vals = sheet.getRange(startRow, col, last - startRow + 1, 1).getValues();
  const seen = new Set();
  for(const [v] of vals){ const s=_rv_(v); if(s) seen.add(s); }
  return [...seen].sort();
}

function _applyDropdown_(workSheet, startRow, col, list){
  if(!list || !list.length) return;
  const maxRows = Math.max(workSheet.getMaxRows(), 3000);
  const range   = workSheet.getRange(startRow, col, maxRows - startRow + 1, 1);
  const rule    = SpreadsheetApp.newDataValidation()
                    .requireValueInList(list, true)
                    .setAllowInvalid(true)
                    .build();
  range.setDataValidation(rule);
}

/* ───── 드롭다운 갱신 (내부용, alert 없음) ───────────────────── */
function _refreshReviewDropdowns_(execSheet){
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const work = ss.getSheetByName(REVIEW_CONF.WORK_SHEET);
  if(!work) return;

  const exec = execSheet || ss.getSheetByName(REVIEW_CONF.EXEC_SHEET);
  if(!exec) return;

  const dr = REVIEW_CONF.DATA_ROW;
  const ec = REVIEW_CONF.EXEC_COL;

  const execList     = _uniqueList_(exec, ec.EXECUTOR, dr);
  const platformList = _uniqueList_(exec, ec.PLATFORM, dr);
  const categoryList = _uniqueList_(exec, ec.CATEGORY, dr);

  _applyDropdown_(work, dr, REVIEW_CONF.COL.PLATFORM,  platformList);  // C열
  _applyDropdown_(work, dr, REVIEW_CONF.COL.CATEGORY,  categoryList);  // F열
  _applyDropdown_(work, dr, REVIEW_CONF.COL.EXECUTOR,  execList);      // G열
}

/* ───── 드롭다운 설치 (메뉴용, alert 포함) ───────────────────── */
function installReviewDropdowns(){
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const work = ss.getSheetByName(REVIEW_CONF.WORK_SHEET);
  const exec = ss.getSheetByName(REVIEW_CONF.EXEC_SHEET);

  if(!work){ SpreadsheetApp.getUi().alert('❌ "리뷰 업무시트" 를 찾을 수 없습니다.'); return; }
  if(!exec){ SpreadsheetApp.getUi().alert('❌ "리뷰실행사" 시트를 찾을 수 없습니다.');  return; }

  const dr = REVIEW_CONF.DATA_ROW;
  const ec = REVIEW_CONF.EXEC_COL;

  const execList     = _uniqueList_(exec, ec.EXECUTOR, dr);
  const platformList = _uniqueList_(exec, ec.PLATFORM, dr);
  const categoryList = _uniqueList_(exec, ec.CATEGORY, dr);

  _applyDropdown_(work, dr, REVIEW_CONF.COL.PLATFORM,  platformList);  // C열
  _applyDropdown_(work, dr, REVIEW_CONF.COL.CATEGORY,  categoryList);  // F열
  _applyDropdown_(work, dr, REVIEW_CONF.COL.EXECUTOR,  execList);      // G열

  try{
    SpreadsheetApp.getUi().alert(
      '✅ 드롭다운 설치 완료\n' +
      '  C열(플랫폼):   ' + platformList.length + '개\n' +
      '  F열(상품분류): ' + categoryList.length + '개\n' +
      '  G열(실행사):   ' + execList.length     + '개'
    );
  } catch(e){}
}

/* ───── 실행비·택배비 조회 ────────────────────────────────────── */
function _lookupFees_(execSheet, executor, platform, category){
  const last = execSheet.getLastRow();
  if(last < REVIEW_CONF.DATA_ROW) return null;

  const ec   = REVIEW_CONF.EXEC_COL;
  const maxC = Math.max(ec.EXECUTOR, ec.PLATFORM, ec.CATEGORY, ec.EXEC_FEE, ec.DELIVERY);
  const vals = execSheet.getRange(
    REVIEW_CONF.DATA_ROW, 1,
    last - REVIEW_CONF.DATA_ROW + 1, maxC
  ).getValues();

  for(const row of vals){
    if(_rv_(row[ec.EXECUTOR-1]) === executor &&
       _rv_(row[ec.PLATFORM-1]) === platform  &&
       _rv_(row[ec.CATEGORY-1]) === category){
      return {
        execFee:  row[ec.EXEC_FEE-1],
        delivery: row[ec.DELIVERY-1]
      };
    }
  }
  return null;
}

/* ───── I·J열 채우기 (한 행) ─────────────────────────────────── */
function _fillFeeRow_(workSheet, targetRow, fees){
  const c = REVIEW_CONF.COL;
  workSheet.getRange(targetRow, c.EXEC_FEE).setValue(fees.execFee);
  workSheet.getRange(targetRow, c.DELIVERY).setValue(fees.delivery);
  // L열(지출예정금액)은 수식이므로 건드리지 않음
}

/* ───── 리뷰실행사 단가 변경 → 해당 조합 I·J 일괄 갱신 ─────── */
function _syncFeesByExecRow_(execSheet, changedRow){
  const ec      = REVIEW_CONF.EXEC_COL;
  const maxEC   = Math.max(ec.EXECUTOR, ec.PLATFORM, ec.CATEGORY, ec.EXEC_FEE, ec.DELIVERY);
  const execVals= execSheet.getRange(changedRow, 1, 1, maxEC).getValues()[0];

  const executor = _rv_(execVals[ec.EXECUTOR-1]);
  const platform = _rv_(execVals[ec.PLATFORM-1]);
  const category = _rv_(execVals[ec.CATEGORY-1]);
  const execFee  = execVals[ec.EXEC_FEE-1];
  const delivery = execVals[ec.DELIVERY-1];
  if(!executor || !platform || !category) return;

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const workSheet = ss.getSheetByName(REVIEW_CONF.WORK_SHEET);
  if(!workSheet) return;

  const c    = REVIEW_CONF.COL;
  const dr   = REVIEW_CONF.DATA_ROW;
  const last = workSheet.getLastRow();
  if(last < dr) return;

  const maxWC = Math.max(c.PLATFORM, c.CATEGORY, c.EXECUTOR);
  const vals  = workSheet.getRange(dr, 1, last - dr + 1, maxWC).getValues();

  for(let i = 0; i < vals.length; i++){
    const r = vals[i];
    if(_rv_(r[c.EXECUTOR-1]) !== executor) continue;
    if(_rv_(r[c.PLATFORM-1]) !== platform)  continue;
    if(_rv_(r[c.CATEGORY-1]) !== category)  continue;
    _fillFeeRow_(workSheet, dr + i, { execFee, delivery });
  }
}

/* ───── onEdit 트리거 ─────────────────────────────────────────── */
/**
 * [리뷰실행사]   D(실행비)/E(택배비) 변경 → 해당 조합 I·J 일괄 갱신
 * [리뷰 업무시트] C/F/G 변경           → I·J 자동채움
 * L열은 수식이므로 트리거가 절대 건드리지 않음
 */
function REVIEW_Trigger(e){
  if(!e || !e.range) return;
  const sheet     = e.range.getSheet();
  const sheetName = sheet.getName();
  const col       = e.range.getColumn();
  const row       = e.range.getRow();

  // ── 리뷰실행사 변경 → 종류에 따라 처리 ──────────────────────
  if(sheetName === REVIEW_CONF.EXEC_SHEET){
    if(row < REVIEW_CONF.DATA_ROW) return;
    const ec = REVIEW_CONF.EXEC_COL;

    if(col === ec.EXEC_FEE || col === ec.DELIVERY){
      // D/E열(실행비·택배비) → I·J 일괄 갱신
      _syncFeesByExecRow_(sheet, row);
    } else if(col === ec.EXECUTOR || col === ec.PLATFORM || col === ec.CATEGORY){
      // A/B/C열(실행사·플랫폼·상품분류) 추가/변경 → 드롭다운 자동 갱신
      _refreshReviewDropdowns_(sheet);
    }
    return;
  }

  // ── 리뷰 업무시트: D열 진행상태 변경 → L열 처리 + 선충전 처리 ────
  // Q열(매출)은 수식이라 onEdit 미발동 → D열 '입금확인' 시점으로 트리거 이동
  if(sheetName === REVIEW_CONF.WORK_SHEET && col === REVIEW_CONF.COL.STATUS){
    if(row >= REVIEW_CONF.DATA_ROW){
      // [2026-10-07] 여러 행을 한꺼번에 바꿔도(붙여넣기·끌어 채우기) 모든 행의 L열 처리
      const nRows = e.range.getNumRows();
      for(let r = row; r < row + nRows; r++) _handleStatusChange_(sheet, r);
      const statusVal = String(e.range.getValue() || '').trim();
      if(statusVal === '입금확인'){
        _checkPrepaidCredit_Review_(sheet, row);
      }
    }
    return;
  }

  // ── [2026-10-15] 리뷰 업무시트: B(고객명)·Q(매출) 입력 → 선충전 잔액 자동 차감 ──
  // ⚠️ 현재 꺼둠(REVIEW_PREPAID_ON_ENTRY=false): 리뷰 고객 선충전잔액에 예전 이중 적립분(리뷰 입금이 선충전에도 쌓임)이
  //    남아 있어, 켜면 아직 돈 안 낸 리뷰 건이 그 잔액으로 입금완료 처리될 수 있음. 선충전 정리 후 true 로 켤 것.
  if(REVIEW_PREPAID_ON_ENTRY && sheetName === REVIEW_CONF.WORK_SHEET && (col === 2 || col === 17) && e.range.getNumColumns() === 1){
    if(row >= REVIEW_CONF.DATA_ROW){
      const nRows = e.range.getNumRows();
      for(let r = row; r < row + nRows; r++) _checkPrepaidCredit_Review_(sheet, r);
    }
    return;
  }

  // ── 리뷰 업무시트: C/F/G 변경 → I·J 자동채움 ────────────────
  if(sheetName !== REVIEW_CONF.WORK_SHEET) return;
  const dr = REVIEW_CONF.DATA_ROW;
  if(row < dr) return;

  const c = REVIEW_CONF.COL;
  if(![c.PLATFORM, c.CATEGORY, c.EXECUTOR].includes(col)) return;

  // 행의 C·F·G 읽기
  const maxC    = Math.max(c.PLATFORM, c.CATEGORY, c.EXECUTOR);
  const rowVals = sheet.getRange(row, 1, 1, maxC).getValues()[0];
  const executor = _rv_(rowVals[c.EXECUTOR-1]);
  const platform = _rv_(rowVals[c.PLATFORM-1]);
  const category = _rv_(rowVals[c.CATEGORY-1]);

  if(!executor || !platform || !category) return;

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const execSheet = ss.getSheetByName(REVIEW_CONF.EXEC_SHEET);
  if(!execSheet) return;

  const fees = _lookupFees_(execSheet, executor, platform, category);
  if(!fees) return;

  _fillFeeRow_(sheet, row, fees);
}

/* ───── D열 진행상태 변경 → L열 지출예정금액 처리 ────────────── */
/**
 * [2026-10-07 대표님 기준]
 * D열 '리뷰 요청'      → L열 = (H+I+J)*K 수식 (지출 예정 금액 산출)
 * 그 외 상태(진행 중·리뷰 전달 완·작업 완료 등) → L열 = 0 (이미 지출된 건, 예정 집계 제외)
 * D열을 비우면 L열은 건드리지 않음
 * (예전: '진행 중'일 때만 0, 한 번 0이 되면 '리뷰 요청'으로 되돌려도 0 고정)
 */
function _handleStatusChange_(sheet, row){
  const statusVal = String(sheet.getRange(row, REVIEW_CONF.COL.STATUS).getValue() || '').trim();
  if(!statusVal) return;
  const costCell  = sheet.getRange(row, REVIEW_CONF.COL.COST);

  // 띄어쓰기 차이('리뷰요청')도 같은 상태로 인정
  if(statusVal.replace(/\s/g, '') === REVIEW_CONF.STATUS_REQUEST.replace(/\s/g, '')){
    costCell.setFormula(`=(H${row}+I${row}+J${row})*K${row}`);
  } else {
    costCell.setValue(0);
  }
}

/* ───── 종합 정산시트 P5 수식 설치 (최초 1회) ────────────────── */
/**
 * 종합 정산시트 P5에 리뷰 업무시트 L열 합계 수식 설정.
 * L열은 D열 진행상태에 따라 0 또는 =(H+I+J)*K 로 관리되므로
 * P5는 SUM만 하면 자동으로 예정 지출 합계가 집계됨.
 */
function installReviewBudgetFormula(){
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const settle = ss.getSheetByName(REVIEW_CONF.SETTLE_SHEET);
  if(!settle){
    SpreadsheetApp.getUi().alert('❌ 종합 정산시트를 찾을 수 없습니다.');
    return;
  }
  settle.getRange(REVIEW_CONF.SETTLE_BUDGET_CELL)
        .setFormula("=SUM('리뷰 업무시트'!L2:L)");
  SpreadsheetApp.getUi().alert(
    '✅ 종합 정산시트 ' + REVIEW_CONF.SETTLE_BUDGET_CELL + '에 리뷰 지출예정 합계 수식 설치 완료!\n' +
    '리뷰 업무시트 D열(진행 중) 변경 시 L열이 자동으로 0 처리되어 실시간 반영됩니다.'
  );
}

/* ───── 트리거 설치 ───────────────────────────────────────────── */
function installReviewTrigger(){
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'REVIEW_Trigger')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('REVIEW_Trigger')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ 리뷰 자동채움 트리거 설치 완료!\n' +
    '리뷰 업무시트 C/F/G 입력 시 I(실행비)·J(택배비) 자동채움\n' +
    '리뷰실행사 D/E 수정 시 해당 조합 I·J 일괄 갱신\n' +
    'L열(지출예정금액)은 수식으로 자동계산됩니다.'
  );
}
