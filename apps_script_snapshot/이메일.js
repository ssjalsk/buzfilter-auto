/***** 설정값 *****/
const CONF = {
  SHEET_NAME: '업무시트',   // 시트명
  COL_STATUS: 5,            // E: 상태
  COL_DATE: 7,              // G: 송출일자
  COL_TIME: 8,              // H: 송출시간
  COL_SENT: 28,             // AB: 예약알림전송 (AA=27은 EXECUTOR_CACHE 전용)
  TIMEZONE: 'Asia/Seoul',
  // 여기에 알림받을 이메일 주소 입력 (예: 본인 Gmail)
  ALERT_EMAIL: 'ssjalsk@gmail.com'
};

/***** 메뉴 생성 *****/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('알림 자동화')
    .addItem('① 트리거 설치(1분 간격)', 'installMinutelyTrigger')
    .addItem('② 수동 한 번 실행(테스트)', 'checkAndNotify')
    .addItem('③ 테스트 메일 발송', 'sendSampleEmail')
    .addToUi();
  ui.createMenu('📊 알리고 자동화')
    .addItem('미수금 목록 수동 갱신', 'refreshMisuList')
    .addSeparator()
    .addItem('⚙️ 미수금 트리거 설치 (최초 1회)', 'installMisuTrigger')
    .addSeparator()
    .addItem('⚙️ 송출연동 트리거 설치 (최초 1회)', 'installHYBTrigger')
    .addItem('⚙️ 입금 자동매칭 트리거 설치 (최초 1회)', 'installSettleTrigger')
    .addItem('⚙️ 고객DB 드롭다운 자동갱신 트리거 설치 (최초 1회)', 'installCustomerDbTrigger')
    .addSeparator()
    .addItem('📋 리뷰 드롭다운 설치/갱신 (C·F·G열)', 'installReviewDropdowns')
    .addItem('⚙️ 리뷰 자동계산 트리거 설치 (최초 1회)', 'installReviewTrigger')
    .addItem('📋 리뷰 지출예정 합계 수식 설치 (P5, 최초 1회)', 'installReviewBudgetFormula')
    .addSeparator()
    .addItem('⚙️ 이디엘 주간 리셋 트리거 설치 (최초 1회)', 'installIdelWeeklyTrigger')
    .addSeparator()
    .addItem('📊 분리정산 통계 수식 설치 (최초 1회)', 'installDivisionStatsFormulas')
    .addItem('🗑️ 분리정산 STAT 트리거 삭제', 'removeStatTrigger')
    .addItem('🗑️ onKEdit 트리거 삭제 (실행사 외부시트 전용 — 메인에서 제거)', 'removeKEditTrigger_')
    .addItem('🔍 설치된 트리거 확인', 'listInstalledTriggers')
    .addToUi();
}

/***** 1분 주기 트리거 설치 *****/
function installMinutelyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkAndNotify')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('checkAndNotify')
    .timeBased()
    .everyMinutes(1)
    .create();

  SpreadsheetApp.getUi().alert('✅ 1분 간격 트리거 설치 완료');
}

/***** 메인 실행 로직 *****/
function checkAndNotify() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONF.SHEET_NAME);
  if (!sh) throw new Error(`시트 "${CONF.SHEET_NAME}" 를 찾을 수 없습니다.`);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const rng = sh.getRange(2, 1, lastRow - 1, CONF.COL_SENT);
  const values = rng.getValues();
  const now = new Date();
  const nowKST = toKST(now, CONF.TIMEZONE);
  const updates = [];

  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    const status = String(row[CONF.COL_STATUS - 1] || '').trim();
    if (status !== '송출예약') continue;

    const dateCell = row[CONF.COL_DATE - 1];
    const timeCell = row[CONF.COL_TIME - 1];
    const sentCell = row[CONF.COL_SENT - 1];
    if (sentCell) continue; // 이미 전송됨

    const scheduled = buildDateTimeKST(dateCell, timeCell, CONF.TIMEZONE);
    if (!scheduled) continue;

    if (nowKST.getTime() >= scheduled.getTime()) {
      const rowIndex = r + 2;
      const client = sh.getRange(rowIndex, 2).getValue();   // B
      const channel = sh.getRange(rowIndex, 4).getValue();  // D
      const media = sh.getRange(rowIndex, 11).getValue();   // K
      const link = sh.getRange(rowIndex, 13).getValue();    // M: 링크 (L=회차, M=링크)

      const subject = `📢 [송출예약 알림] ${client || '무제'} (${channel || ''})`;
      const body = [
        '🕐 송출예약 알림 도착',
        '',
        `• 프로젝트: ${client || '-'}`,
        `• 채널: ${channel || '-'} ${media ? '/ ' + media : ''}`,
        `• 예약시각: ${fmtDate(scheduled)}`,
        link ? `• 링크: ${link}` : '',
        '',
        '자동알림 시스템 - 업무시트'
      ].join('\n');

      const ok = sendEmail(CONF.ALERT_EMAIL, subject, body);
      if (ok) updates.push({ r: rowIndex, c: CONF.COL_SENT, v: new Date() });
    }
  }

  // 전송 완료 표시
  updates.forEach(u => sh.getRange(u.r, u.c).setValue(u.v));
}

/***** Gmail 발송 *****/
function sendEmail(to, subject, body) {
  try {
    GmailApp.sendEmail(to, subject, body);
    return true;
  } catch (e) {
    console.error('메일 발송 실패:', e);
    return false;
  }
}

/***** 도우미 함수들 *****/
function buildDateTimeKST(dateCell, timeCell, tz) {
  if (!dateCell) return null;
  let d = (dateCell instanceof Date) ? new Date(dateCell) : new Date(dateCell);
  if (isNaN(d.getTime())) return null;

  let hours = 9, minutes = 0;
  if (timeCell) {
    if (timeCell instanceof Date) {
      hours = timeCell.getHours();
      minutes = timeCell.getMinutes();
    } else {
      const m = String(timeCell).match(/^(\d{1,2}):(\d{2})$/);
      if (m) {
        hours = parseInt(m[1], 10);
        minutes = parseInt(m[2], 10);
      }
    }
  }
  const y = d.getFullYear(), mo = d.getMonth(), da = d.getDate();
  const dt = new Date(y, mo, da, hours, minutes, 0, 0);
  return toKST(dt, tz);
}

function toKST(date, tz) {
  const s = Utilities.formatDate(new Date(date), tz, 'yyyy/MM/dd HH:mm:ss');
  const p = s.split(/[/ :]/).map(Number);
  return new Date(p[0], p[1] - 1, p[2], p[3], p[4], p[5]);
}
function fmtDate(d) {
  return Utilities.formatDate(d, CONF.TIMEZONE, 'yyyy-MM-dd HH:mm');
}
function sendSampleEmail() {
  sendEmail(CONF.ALERT_EMAIL, '🔔 테스트 알림', '이 메일이 도착했다면 설정이 잘 된 것입니다.');
}

/***** onKEdit 트리거 삭제 — 메인 시트에 잘못 설치된 트리거 제거용 *****/
function removeKEditTrigger_() {
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onKEdit');
  if (triggers.length === 0) {
    SpreadsheetApp.getUi().alert('ℹ️ onKEdit 트리거가 없습니다. (이미 삭제됐거나 설치된 적 없음)');
    return;
  }
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  SpreadsheetApp.getUi().alert('✅ onKEdit 트리거 ' + triggers.length + '개 삭제 완료\n메인 시트에서 회차갱신 함수가 더 이상 실행되지 않습니다.');
}
