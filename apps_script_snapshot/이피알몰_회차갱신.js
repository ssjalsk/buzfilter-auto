/**
 * 이피알몰 진행 내역 시트
 * K열(충전금액) 입력 시 빈 슬롯 A열(회차) 자동 갱신
 * ─ 자동 배포 버전: 설치 후 시트를 열면 상단 메뉴에서 트리거 설치 가능
 */

var LAYOUT = {
  headerRows: 2,
  A_round:    1,
  B_subject:  2,
  K_charge:   11
};
var SLOT_COUNT = 30;

// 메인 시트 ID — 이 스프레드시트에서는 회차갱신 로직이 절대 실행되면 안 됨
var MAIN_SS_ID_EPR = '1mOV-HmlODZaxPohiVFay9_Rnh31-vhcPXF_b-5-EjB0';

/* ── 메뉴: 외부 시트 배포 버전 전용 (메인 시트에서는 이메일.js의 onOpen() 사용) ── */
function onOpen_EPR_() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ 회차 갱신')
    .addItem('▶ 회차 갱신 트리거 설치 (최초 1회)', 'installTrigger')
    .addItem('🔄 지금 즉시 회차 갱신', 'testRenumber')
    .addToUi();
}

/* ── 설치형 트리거 핸들러 ── */
function onKEdit(e) {
  if (!e || !e.range) return;
  // 실행사 전용 — 메인 시트에서는 절대 실행 방지
  if (SpreadsheetApp.getActive().getId() === MAIN_SS_ID_EPR) return;
  var sheet = e.range.getSheet();
  var col   = e.range.getColumn();
  var row   = e.range.getRow();
  if (col !== LAYOUT.K_charge) return;
  if (row <= LAYOUT.headerRows) return;
  var val = Number(String(e.range.getValue() || '').replace(/[,\s₩]/g, ''));
  if (!(val > 0)) return;
  _renumberEmptySlots_(sheet);
}

/* ── 빈 슬롯(B열 비어있는 행) A열에 1부터 번호 부여 ── */
function _renumberEmptySlots_(sheet) {
  var lastRow  = Math.max(sheet.getLastRow(), LAYOUT.headerRows + SLOT_COUNT);
  var dataRows = lastRow - LAYOUT.headerRows;
  if (dataRows <= 0) return;
  var vals = sheet.getRange(LAYOUT.headerRows + 1, 1, dataRows, LAYOUT.B_subject).getValues();
  var counter = 1;
  var updates = [];
  for (var i = 0; i < vals.length; i++) {
    var bVal = String(vals[i][LAYOUT.B_subject - 1] || '').trim();
    if (bVal === '') {
      updates.push({ row: LAYOUT.headerRows + 1 + i, num: counter++ });
      if (counter > SLOT_COUNT) break;
    }
  }
  for (var j = 0; j < updates.length; j++) {
    sheet.getRange(updates[j].row, LAYOUT.A_round).setValue(updates[j].num);
  }
  Logger.log('✅ 빈 슬롯 회차 갱신: ' + updates.length + '개');
}

/* ── 트리거 설치 (메뉴에서 실행) ── */
function installTrigger() {
  // 메인 시트에서는 설치 차단
  if (SpreadsheetApp.getActive().getId() === MAIN_SS_ID_EPR) {
    SpreadsheetApp.getUi().alert('⚠️ 이 트리거는 이피알몰 외부 시트에서만 설치 가능합니다.\n메인 시트에서는 설치할 수 없습니다.');
    return;
  }
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'onKEdit'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onKEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert('✅ 트리거 설치 완료!\nK열에 충전금액 입력 시 빈 슬롯 회차가 자동 갱신됩니다.');
}

/* ── 수동 즉시 갱신 (메뉴에서 실행) ── */
function testRenumber() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  _renumberEmptySlots_(sheet);
  SpreadsheetApp.getUi().alert('✅ 회차 갱신 완료');
}
