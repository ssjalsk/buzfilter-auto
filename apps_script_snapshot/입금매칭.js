/**
 * 입금 자동 매칭 — Cloud Function 연동
 * 종합 정산시트 G열(입금자명) + H열(미발행) 또는 K열(부가세포함) 입력 시 처리
 *
 * [수정 2026-09-16]
 * - 입력 순서 독립: G열→금액, 금액→G열 어느 순서든 모두 처리
 * - G열, H열, K열 중 어느 것이든 편집 시 → G열과 금액 둘 다 있으면 실행
 * - 금액 없으면 대기 (나중에 금액 입력 시 자동 처리)
 */

// ── 배포 후 여기에 실제 URL 입력 ────────────────────────────────
const MATCH_FUNC_URL    = 'https://asia-northeast3-aligo-automation.cloudfunctions.net/match-payment';
const MATCH_FUNC_SECRET = 'aligo-secret-2026';
// ─────────────────────────────────────────────────────────────────

/**
 * Cloud Function 호출
 */
function callMatchPayment_(rowIndex, payer, amountK, amountH, mode) {
  if (!MATCH_FUNC_URL || MATCH_FUNC_URL.startsWith('YOUR_')) {
    console.warn('MATCH_FUNC_URL 미설정 — Cloud Function 배포 후 URL을 입력하세요.');
    return;
  }

  try {
    const payload = JSON.stringify({
      row:      rowIndex,
      payer:    payer,
      amount_k: amountK || 0,
      amount_h: amountH || 0,
      mode:     mode || 'match',   // 'delta' = 이미 매칭된 행의 금액 변경 (차액 처리)
      secret:   MATCH_FUNC_SECRET
    });

    const options = {
      method:      'post',
      contentType: 'application/json',
      payload:     payload,
      muteHttpExceptions: true
    };

    const resp = UrlFetchApp.fetch(MATCH_FUNC_URL, options);
    const code = resp.getResponseCode();
    const body = JSON.parse(resp.getContentText() || '{}');

    if (code === 200 && body.ok) {
      if (body.mode === 'delta') {
        console.log(`🔁 금액 변경 처리 (행 ${rowIndex}): ${body.action}${body.diff ? ' / 차액 ' + body.diff : ''}`);
      } else if (body.matched) {
        const detail = `processed:${body.processed}, prepaid:${body.prepaid}`;
        console.log(`✅ 입금 매칭 완료: "${payer}" → "${body.matched}" (행 ${rowIndex}) [${detail}]`);
        // 계산서 체크 시트: 부가세포함 금액 일치 시 선발행 → 영수발행 변경
        _updateInvoiceSheet_(body.matched, amountK, amountH);
      } else {
        console.log(`⚠️ 매칭 고객 없음: "${payer}" — F열 비워둠`);
      }
    } else {
      console.error(`❌ Cloud Function 오류 [${code}]:`, body.error || body);
    }
  } catch (err) {
    console.error('callMatchPayment_ 예외:', err.message);
  }
}

/**
 * 설치형 트리거 핸들러 — 종합 정산시트 G/H/K열 감시
 * (단순 onEdit은 UrlFetchApp 불가 → 설치형 트리거 필요)
 *
 * [처리 조건] 입력 순서 무관 — 아래 조건 모두 충족 시 실행:
 *   1. 종합 정산시트의 E(5), G(7), H(8), K(11)열 중 하나 편집
 *   2. E열(구분) = '매출'
 *   3. G열(입금자명) 비어있지 않음
 *   4. H열(미발행) 또는 K열(부가세포함) 중 하나 이상 값 있음
 *   → 매출·입금자명·금액 3가지를 어떤 순서로 입력해도, 마지막 칸 입력 순간 1회 실행
 *
 * [2026-10-07] E열 추가: 입금자명·금액을 먼저 쓰고 '매출'을 마지막에 고르는 경우 대응.
 *   E열 편집은 아래일 때만 실행 (과거 행 재처리 방지):
 *   - E열 한 칸만 편집 (여러 칸 붙여넣기·끌어 채우기 제외)
 *   - 빈칸 → '매출' 로 처음 입력 (이미 '매출'인 행을 다시 고르는 경우 제외)
 */
function onSettleEdit(e) {
  if (!e) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== '종합 정산시트') return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 4) return;

  // E(5), G(7), H(8), K(11)열 편집만 처리
  if (col !== 5 && col !== 7 && col !== 8 && col !== 11) return;

  // [2026-10-07] E열 편집: 한 칸 + 빈칸→'매출' 첫 입력일 때만
  if (col === 5) {
    if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
    const newVal = String((e.value !== undefined && e.value !== null) ? e.value : sheet.getRange(row, 5).getValue()).trim();
    const oldVal = String((e.oldValue !== undefined && e.oldValue !== null) ? e.oldValue : '').trim();
    if (newVal !== '매출' || oldVal !== '') return;
  }

  // E열(5) = '매출'인 행만 처리
  const eVal = String(sheet.getRange(row, 5).getValue() || '').trim();
  if (eVal !== '매출') return;

  // ── [2026-10-06] 잠금: 정산 입금 처리는 한 번에 하나씩 ──────────
  // 같은 편집에 트리거가 두 번 실행되거나 여러 행이 동시에 입력돼도 이중 처리 방지.
  // 잠금을 잡은 뒤 F열·금액을 다시 읽으므로, 먼저 처리된 결과를 보고 판단함.
  const lock = LockService.getScriptLock();
  try { lock.waitLock(120000); }
  catch (err) { console.error(`❌ 행 ${row}: 입금 처리 잠금 대기 초과 — 다시 입력해 주세요.`); return; }

  try {
    const existingAdv = String(sheet.getRange(row, 6).getValue() || '').trim();
    const payer       = String(sheet.getRange(row, 7).getValue() || '').trim();
    const kRaw        = sheet.getRange(row, 11).getValue();  // K열: 부가세포함금액
    const hRaw        = sheet.getRange(row, 8).getValue();   // H열: 미발행금액
    const amountK     = kRaw ? Number(String(kRaw).replace(/[^\d.]/g, '')) || 0 : 0;
    const amountH     = hRaw ? Number(String(hRaw).replace(/[^\d.]/g, '')) || 0 : 0;

    // ── 이미 광고주 매칭된 행 ────────────────────────────────
    // 예전: 무조건 스킵 → 금액을 고쳐도 반영 안 됨
    // 지금: 금액(H/K) 수정이고 자동매칭 처리 기록(AC열)이 있을 때만 차액 처리 요청
    //       (늘어남 → 차액만 추가 처리 / 줄어듦 → 표시만). 기록 없는 행은 예전처럼 스킵.
    if (existingAdv) {
      const record = String(sheet.getRange(row, 29).getValue() || '').trim();  // AC열
      if ((col === 8 || col === 11) && record && !existingAdv.startsWith('⚠️')) {
        console.log(`🔁 행 ${row}: 금액 변경 감지 — 처리기록 ${record}, K열 ${amountK}, H열 ${amountH}`);
        callMatchPayment_(row, payer, amountK, amountH, 'delta');
      }
      return;
    }

    // ── 입력 순서 무관: G열과 금액 둘 다 있을 때만 실행 ──────────
    if (!payer) {
      console.log(`ℹ️ 행 ${row}: G열(입금자명) 없음 — 입금자명 입력 후 처리됩니다.`);
      return;
    }
    if (amountK === 0 && amountH === 0) {
      console.log(`ℹ️ 행 ${row}: 금액(H/K열) 없음 — 금액 입력 후 처리됩니다.`);
      return;
    }

    console.log(`🔄 행 ${row}: 입금 매칭 시작 — 입금자: "${payer}", K열: ${amountK}, H열: ${amountH}`);
    callMatchPayment_(row, payer, amountK, amountH, 'match');
  } finally {
    lock.releaseLock();
  }
}

/**
 * 트리거 설치 (최초 1회 실행)
 */
function installSettleTrigger() {
  // 기존 트리거 중복 방지
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onSettleEdit')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onSettleEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert('✅ 정산시트 입금 자동매칭 트리거 설치 완료!');
}

/**
 * 수동 테스트용 — Apps Script 편집기에서 직접 실행
 */
function testMatchPayment() {
  const testRow     = 5;
  const testPayer   = '홍길동';
  const testAmountK = 132000;
  callMatchPayment_(testRow, testPayer, testAmountK, 0);
  SpreadsheetApp.getUi().alert('테스트 완료 — 로그에서 결과 확인');
}
