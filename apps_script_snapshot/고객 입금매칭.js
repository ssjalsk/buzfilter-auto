/***** 스택(remaining) 기반 정산 실행부 – H/I 통합 *****/
/**
 * S열 개선: 부분입금 시 텍스트 메시지 대신 "누적 입금액(숫자)" 저장
 * → 추가 입금 시 N - S = 실잔액으로 정확히 차감
 */

/* ───────── 유틸 ───────── */
function _norm(v){return (v===null||v===undefined)?'':String(v).replace(/\s+/g,' ').trim();}
function _num(v){
  if(v===null||v===undefined)return 0;
  if(typeof v==='number')return isNaN(v)?0:Math.round(v);
  const s=String(v).replace(/[,\s₩원￦]/g,'').replace(/[^\d.-]/g,'');
  const n=parseFloat(s); return isNaN(n)?0:Math.round(n);
}
function _dateKey(v){
  if(v instanceof Date&&!isNaN(v))return v.getTime();
  const s=_norm(v); if(!s) return 9e15;
  const m=s.match(/^(\d{4})[.\-\/\s](\d{1,2})[.\-\/\s](\d{1,2})$/);
  if(m){ const dt=new Date(+m[1],+m[2]-1,+m[3]); return isNaN(dt)?9e15:dt.getTime(); }
  const dt2=new Date(s); return isNaN(dt2)?9e15:dt2.getTime();
}
function _setCell(sh,r,c,v){ sh.getRange(r,c).setValue(v); }
function _clearCell(sh,r,c){ sh.getRange(r,c).clearContent(); }

/* Q='입금완료'이면 M 텍스트 존재 시 E='송출완료' + 미수금 목록 즉시 갱신 */
function _updateStatusIfPublished_(sheet,row){
  const E=5,M=13,Q=17;
  if (_norm(sheet.getRange(row,Q).getValue())!=='입금완료') return;
  if (_norm(sheet.getRange(row,M).getValue())) {
    sheet.getRange(row,E).setValue('송출완료');
    // 프로그래밍 변경은 onEditMisuCheck_ 트리거가 감지 못하므로 직접 갱신
    try { if(sheet.getName()==='업무시트') refreshMisuList(); } catch(e) {}
  }
}

/**
 * 고객 드롭다운 갱신 — 고객_DB 기반
 * 업무시트 B열 + 종합 정산시트 F열 동시 적용
 * 고객_DB 없으면 업무시트 B열 기존 값 fallback
 */
function ensureDropdownSource_(){
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const work = ss.getSheetByName('업무시트');
  const sum  = ss.getSheetByName('종합 정산시트');
  const db   = ss.getSheetByName('고객_DB');
  if(!work || !sum) return null;

  // ── 고객_DB에서 담당자명/사업자명 읽기 ──────────────────
  const uniq = new Map();
  if(db) {
    const dbVals = db.getDataRange().getValues();
    const header = dbVals[0] || [];
    const nameIdx = header.indexOf('담당자명');
    const bizIdx  = header.indexOf('사업자명');
    for(let i = 1; i < dbVals.length; i++){
      const r = dbVals[i];
      if(!r.some(v => v)) continue;
      const name = nameIdx >= 0 ? _norm(r[nameIdx]) : '';
      const biz  = bizIdx  >= 0 ? _norm(r[bizIdx])  : '';
      if(name) uniq.set(name, true);
      else if(biz) uniq.set(biz, true);
    }
  }

  // ── fallback: 업무시트 B열 기존 값 ──────────────────────
  if(!uniq.size){
    const lastW = work.getLastRow();
    if(lastW >= 2){
      const vals = work.getRange(2, 2, lastW-1, 1).getValues();
      for(const v of vals){ const k = _norm(v[0]); if(k) uniq.set(k, true); }
    }
  }

  const list = [...uniq.keys()].sort();
  if(!list.length) return null;

  // ── 종합 정산시트 AA열(27)에 목록 저장 ──────────────────
  const SUM_AA = 27, SUM_F = 6, F_START = 4;
  const HEADER = '__고객목록(자동)__';
  if(_norm(sum.getRange(1, SUM_AA).getValue()) !== HEADER)
    sum.getRange(1, SUM_AA).setValue(HEADER);
  const clearLen = Math.max(sum.getLastRow(), list.length + 1);
  if(clearLen > 1) sum.getRange(2, SUM_AA, clearLen-1, 1).clearContent();
  sum.getRange(2, SUM_AA, list.length, 1).setValues(list.map(x => [x]));

  const src   = sum.getRange(2, SUM_AA, Math.max(1, sum.getMaxRows()-1), 1);
  const rule  = SpreadsheetApp.newDataValidation()
                  .requireValueInRange(src, true).setAllowInvalid(true).build();

  // ── 종합 정산시트 F열 드롭다운 ───────────────────────────
  const lastF = Math.max(sum.getMaxRows(), 2000);
  sum.getRange(F_START, SUM_F, lastF - F_START + 1, 1).setDataValidation(rule);

  // ── 업무시트 B열 드롭다운 ────────────────────────────────
  const lastB = Math.max(work.getMaxRows(), 3000);
  work.getRange(2, 2, lastB - 1, 1).setDataValidation(rule);

  // ── 리뷰 업무시트 B열 드롭다운 ───────────────────────────
  const review = ss.getSheetByName('리뷰 업무시트');
  if(review) {
    const lastR = Math.max(review.getMaxRows(), 3000);
    review.getRange(2, 2, lastR - 1, 1).setDataValidation(rule);
  }

  // ── 계산서 체크 시트 A열 드롭다운 ────────────────────────
  const invoice = ss.getSheetByName('계산서 체크');
  if(invoice) {
    const lastI = Math.max(invoice.getMaxRows(), 3000);
    invoice.getRange(2, 1, lastI - 1, 1).setDataValidation(rule);
  }

  return src;
}

/**
 * onChange 트리거 핸들러
 * 고객_DB 행 추가/변경(API 포함) 시 자동 드롭다운 갱신
 */
function onCustomerDbChange(e) {
  try {
    ensureDropdownSource_();
  } catch(err) {
    console.error('onCustomerDbChange 오류:', err.message);
  }
}

/**
 * 설치형 onChange 트리거 등록 (최초 1회 실행)
 * Apps Script 편집기에서 직접 실행
 */
function installCustomerDbTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onCustomerDbChange')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onCustomerDbChange')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create();

  SpreadsheetApp.getUi().alert('✅ 고객DB → 드롭다운 자동갱신 트리거 설치 완료!\n고객_DB에 새 행 추가 시 업무시트 B열 + 종합 정산시트 F열이 자동으로 갱신됩니다.');
}

/* 금액 안전 추출 */
function _getEditedAmount_(sum, row, col, editedValue){
  const a=_num(editedValue); if(a>0) return a;
  const b=_num(sum.getRange(row,col).getValue()); if(b>0) return b;
  const c=_num(sum.getRange(row,col).getDisplayValue()); return c>0?c:0;
}

/* ───────── 메인: H/I 통합 스택 정산 ───────── */
function APPLY_RECON_STACK_(row,col,editedValue){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sum=ss.getSheetByName('종합 정산시트');
  const work=ss.getSheetByName('업무시트');
  if(!sum||!work) return;

  const lock=LockService.getDocumentLock();
  try{ lock.waitLock(30000); }catch(e){ return; }

  try{
    ensureDropdownSource_();

    const SUM={F:6,G:7,H:8,I:9};
    const W={B:2,E:5,F:6,M:13,N:14,Q:17,R:18,S:19};
    const TOL=1;
    if(row<4 || (col!==SUM.H && col!==SUM.I)) return;

    const advertiser=_norm(sum.getRange(row,SUM.F).getValue());
    if(!advertiser) return;
    // [2026-10-06] 자동매칭(Cloud Function)이 이미 처리한 행(AC열 처리기록 있음)은 건너뜀.
    //  → 금액 수정 시 여기서 '전체 금액'을 다시 처리하던 이중 처리 경로 차단.
    //    차액 처리는 onSettleEdit(입금매칭.js) → Cloud Function delta 모드가 전담.
    if(_norm(sum.getRange(row,29).getValue())) return;
    const payer=_norm(sum.getRange(row,SUM.G).getValue());

    const amount=_getEditedAmount_(sum,row,col,editedValue);
    if(!(amount>0)) return;

    // ── 후보 수집: B==advertiser && Q!='입금완료' ──────────
    // S열에 누적 입금액이 있으면 need = N - S (실잔액)
    const last=work.getLastRow(); if(last<2) return;
    const data=work.getRange(2,1,last-1,W.S).getValues();

    const cands=[];
    for(let i=0;i<data.length;i++){
      const r=data[i];
      if(_norm(r[W.B-1])!==advertiser) continue;
      if(_norm(r[W.Q-1])==='입금완료') continue;

      const totalCost  = _num(r[W.N-1]); if(totalCost<=0) continue;
      const alreadyPaid= _num(r[W.S-1]); // S열: 숫자면 누적 입금액, 텍스트면 0
      const need       = Math.max(0, totalCost - alreadyPaid);

      // S에 기록된 금액이 N 이상 → 완납 미처리건 자동 마감
      if(need<=0){
        _setCell(work,i+2,W.Q,'입금완료');
        if(payer) _setCell(work,i+2,W.R,payer);
        _clearCell(work,i+2,W.S);
        _updateStatusIfPublished_(work,i+2);
        continue;
      }
      const dKey=_dateKey(r[W.F-1]);
      cands.push({row:i+2, need, dKey, alreadyPaid, totalCost, st:'work'});
    }

    // ── 리뷰 업무시트 후보 추가 ──────────────────────────────
    const RV={B:2,Q:17,T:20,U:21,W:23,O:15};
    const review=ss.getSheetByName('리뷰 업무시트');
    if(review){
      const lastR=review.getLastRow();
      if(lastR>=2){
        const dataR=review.getRange(2,1,lastR-1,RV.W).getValues();
        for(let i=0;i<dataR.length;i++){
          const r=dataR[i];
          if(_norm(r[RV.B-1])!==advertiser) continue;
          if(_norm(r[RV.T-1])==='입금완료') continue;
          const totalCost=_num(r[RV.Q-1]); if(totalCost<=0) continue;
          const dKey=_dateKey(r[RV.O-1]);
          cands.push({row:i+2, need:totalCost, dKey, alreadyPaid:0, totalCost, st:'review'});
        }
      }
    }
    if(!cands.length) return;

    // 날짜 오름차순(FIFO)
    cands.sort((a,b)=>(a.dKey-b.dKey)||(a.row-b.row));
    // ⚠️ S 일괄 초기화 없음 — 누적 추적 보존

    let remaining=amount;
    let lastProcessedRow=null;

    for(const c of cands){
      const sh=c.st==='review'?review:work;
      const cQ=c.st==='review'?RV.T:W.Q;
      const cR=c.st==='review'?RV.U:W.R;
      if(remaining >= c.need - TOL){
        // ── 완납 ──────────────────────────────────────────
        _setCell(sh,c.row,cQ,'입금완료');
        if(payer) _setCell(sh,c.row,cR,payer);
        if(c.st==='review') sh.getRange(c.row,RV.W).clearContent();
        else _clearCell(work,c.row,W.S);
        if(c.st!=='review') _updateStatusIfPublished_(work,c.row);
        remaining -= c.need;
        lastProcessedRow=c.row;
        if(remaining <= TOL){ remaining=0; break; }
      }else{
        if(c.st==='review'){
          // ── 리뷰: W열에 부족금액 표시 ────────────────
          const shortage=c.need-remaining;
          _setCell(review,c.row,RV.W,'₩'+shortage.toLocaleString()+' 부족');
          if(payer) _setCell(review,c.row,RV.U,payer);
        }else{
          // ── 업무시트: S에 누적 입금액(숫자) 저장 ─────
          const newPaid=c.alreadyPaid+remaining;
          _setCell(work,c.row,W.S,newPaid);
          if(payer) _setCell(work,c.row,W.R,payer);
        }
        lastProcessedRow=c.row;
        remaining=0;
        break;
      }
    }

    // ── 초과입금 처리: 선충전잔액에 추가 ────
    if(remaining > TOL){
      _addToPrepaid_(advertiser, Math.round(remaining));
    }

    SpreadsheetApp.flush();
  }finally{
    try{ lock.releaseLock(); }catch(e){}
  }
}

/* ───── 선충전 잔액 자동 입금완료 처리 ─────────────────────── */
/**
 * 업무시트 N열(금액) 입력 시 자동 실행:
 * 1. 고객_DB G열(선충전 잔액) 확인
 * 2. 잔액 >= N열 금액 → Q열='입금완료', DB 잔액 차감, S열 초기화
 * 3. 잔액 < N열 금액  → S열에 '₩XXX 부족' 표시 (입금완료 처리 안 함)
 * 4. 잔액 없는 고객   → 아무것도 안 함 (기존 방식 유지)
 */
function _checkPrepaidCredit_(workSheet, row) {
  const W = { B: 2, N: 14, Q: 17, R: 18, S: 19 };
  const DB_PREPAID_COL = 7; // G열
  const DB_ALIAS_COL   = 5; // E열 (별칭 — 통장 입금자명)

  const customer  = _norm(workSheet.getRange(row, W.B).getValue());
  const amount    = _num(workSheet.getRange(row, W.N).getValue());
  const payStatus = _norm(workSheet.getRange(row, W.Q).getValue());

  // 고객명/금액 없거나 이미 입금완료면 스킵
  if (!customer || !(amount > 0) || payStatus === '입금완료') return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const db = ss.getSheetByName('고객_DB');
  if (!db) return;

  // 고객_DB에서 일치하는 행 찾기 (담당자명 또는 사업자명)
  const dbVals  = db.getDataRange().getValues();
  const header  = dbVals[0] || [];
  const nameIdx = header.indexOf('담당자명');
  const bizIdx  = header.indexOf('사업자명');
  const aliasIdx= header.indexOf('별칭');  // E열: 통장 입금자명

  let dbRowNum  = -1;
  let prepaid   = 0;
  let payerName = '';

  for (let i = 1; i < dbVals.length; i++) {
    const r        = dbVals[i];
    const name     = nameIdx  >= 0 ? _norm(r[nameIdx])  : '';
    const biz      = bizIdx   >= 0 ? _norm(r[bizIdx])   : '';
    const rawAlias = aliasIdx >= 0 ? _norm(r[aliasIdx]) : '';

    // 별칭은 쉼표 구분 목록 → 각각 비교
    const aliases  = rawAlias ? rawAlias.split(',').map(a => a.trim()).filter(a => a) : [];
    const isMatch  = name === customer || biz === customer || aliases.includes(customer);

    if (isMatch) {
      dbRowNum = i + 1; // 1-indexed
      prepaid  = _num(r[DB_PREPAID_COL - 1]); // G열 = index 6

      // 입금자명: 별칭 첫 번째 값 → 없으면 담당자명 → 없으면 사업자명
      payerName = (aliases[0] || name || biz);
      break;
    }
  }

  // DB에 고객 없거나 선충전 잔액 없으면 스킵
  if (dbRowNum < 0 || !(prepaid > 0)) return;

  if (prepaid >= amount) {
    // ── 완납: 입금자명 기입 + 입금완료 처리 + 고객_DB 잔액 차감 ──
    if (payerName) _setCell(workSheet, row, W.R, payerName); // R열: 입금자명
    _setCell(workSheet, row, W.Q, '입금완료');               // Q열: 입금완료
    workSheet.getRange(row, W.S).clearContent();
    _updateStatusIfPublished_(workSheet, row);
    db.getRange(dbRowNum, DB_PREPAID_COL).setValue(prepaid - amount);
  } else {
    // ── 잔액 부족: 선충전 전액 차감 + S열에 처리된 금액(숫자) 저장 ──
    // prepaid = 선충전으로 처리된 금액 → 실입금 처리 시 already-paid로 인식
    db.getRange(dbRowNum, DB_PREPAID_COL).setValue(0);
    _setCell(workSheet, row, W.S, prepaid);
  }
}

/* ───── 리뷰 업무시트 선충전 잔액 자동 입금완료 처리 ──────────── */
/**
 * 리뷰 업무시트 Q열(매출금액) 입력 시 자동 실행:
 * 1. 고객_DB G열(선충전 잔액) 확인
 * 2. 잔액 >= Q열 금액 → T열='입금완료', DB 잔액 차감, W열 초기화
 * 3. 잔액 < Q열 금액  → W열에 '₩XXX 부족' 표시 (입금완료 처리 안 함)
 * 4. 잔액 없는 고객   → 아무것도 안 함
 */
function _checkPrepaidCredit_Review_(reviewSheet, row) {
  const R = { B: 2, Q: 17, T: 20, U: 21, W: 23 };
  const DB_PREPAID_COL = 7; // G열
  const DB_ALIAS_COL   = 5; // E열 (별칭 — 통장 입금자명)

  // Q열이 수식일 수 있으므로 flush()로 최신값 보장 후 읽기
  SpreadsheetApp.flush();

  const customer  = _norm(reviewSheet.getRange(row, R.B).getValue());
  const amount    = _num(reviewSheet.getRange(row, R.Q).getValue());
  const payStatus = _norm(reviewSheet.getRange(row, R.T).getValue());

  if (!customer || !(amount > 0) || payStatus === '입금완료') return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const db = ss.getSheetByName('고객_DB');
  if (!db) return;

  const dbVals  = db.getDataRange().getValues();
  const header  = dbVals[0] || [];
  const nameIdx = header.indexOf('담당자명');
  const bizIdx  = header.indexOf('사업자명');
  const aliasIdx= header.indexOf('별칭');  // E열: 통장 입금자명

  let dbRowNum  = -1;
  let prepaid   = 0;
  let payerName = '';

  for (let i = 1; i < dbVals.length; i++) {
    const r        = dbVals[i];
    const name     = nameIdx  >= 0 ? _norm(r[nameIdx])  : '';
    const biz      = bizIdx   >= 0 ? _norm(r[bizIdx])   : '';
    const rawAlias = aliasIdx >= 0 ? _norm(r[aliasIdx]) : '';

    const aliases  = rawAlias ? rawAlias.split(',').map(a => a.trim()).filter(a => a) : [];
    const isMatch  = name === customer || biz === customer || aliases.includes(customer);

    if (isMatch) {
      dbRowNum = i + 1;
      prepaid  = _num(r[DB_PREPAID_COL - 1]);

      // 입금자명: 별칭 첫 번째 값 → 없으면 담당자명 → 없으면 사업자명
      payerName = (aliases[0] || name || biz);
      break;
    }
  }

  if (dbRowNum < 0 || !(prepaid > 0)) return;

  if (prepaid >= amount) {
    // ── 완납: 입금자명 기입 + 입금완료 처리 + 고객_DB 잔액 차감 ──
    if (payerName) _setCell(reviewSheet, row, R.U, payerName); // U열: 입금자명
    _setCell(reviewSheet, row, R.T, '입금완료');               // T열: 입금완료
    reviewSheet.getRange(row, R.W).clearContent();
    db.getRange(dbRowNum, DB_PREPAID_COL).setValue(prepaid - amount);
  } else {
    // ── 잔액 부족: W열에 부족금액 표시 ───────────────────────────
    const shortage = amount - prepaid;
    _setCell(reviewSheet, row, R.W, '₩' + shortage.toLocaleString() + ' 부족');
  }
}

/* ───── 선충전 디버그 (Apps Script 에디터에서 직접 실행) ───────────
 * ROW_TO_TEST에 업무시트 실제 행 번호 입력 후 실행 → 로그 확인
 * 실행: Apps Script 에디터 → 함수 선택 → debugPrepaidCredit → ▶ 실행
 */
/* ───── 계산서 체크 시트: 영수발행 자동변경 ────────────────────────
 * 종합 정산시트 입금 매칭 완료 후 호출.
 * 조건: A열=광고주 일치 + D열(부가세포함금액)=입금액 일치 + F열='선발행' + H열=미완료
 * 충족 시 F열을 '영수발행'으로 변경.
 *
 * @param {string} advertiser  - 매칭된 광고주명
 * @param {number} amountK     - 부가세포함 입금액 (K열, 0이면 H열 입금)
 * @param {number} amountH     - 미발행 입금액 (H열, 0이면 K열 입금)
 */
function _updateInvoiceSheet_(advertiser, amountK, amountH) {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const invoice = ss.getSheetByName('계산서 체크');
    if (!invoice) return;

    const last = invoice.getLastRow();
    if (last < 2) return;

    // A~H열 읽기
    const data  = invoice.getRange(2, 1, last - 1, 8).getValues();
    const advN  = _norm(advertiser);
    const TOL   = 1; // 1원 허용 오차

    // 비교 기준금액: K열(부가세포함) 우선, 없으면 H열(미발행)*1.1 환산
    const compareAmount = amountK > 0 ? amountK : Math.round(amountH * 1.1);
    if (!(compareAmount > 0)) return;

    for (let i = 0; i < data.length; i++) {
      const r       = data[i];
      const rowAdv  = _norm(r[0]);             // A열: 광고주
      const dAmount = _num(r[3]);              // D열: 부가세포함금액
      const fVal    = String(r[5] || '').trim(); // F열: 선발행여부
      const hVal    = r[7];                    // H열: 완료여부 (체크박스 true/false)

      if (rowAdv !== advN) continue;
      if (fVal !== '선발행') continue;
      if (hVal === true || hVal === 'TRUE') continue; // 완료 건 제외
      if (Math.abs(dAmount - compareAmount) > TOL) continue;

      invoice.getRange(i + 2, 6).setValue('영수발행');
      Logger.log(`✅ 계산서 체크: "${advertiser}" ${compareAmount}원 → 영수발행 변경 (${i+2}행)`);
      break; // 첫 번째 일치 건만 처리
    }
  } catch(err) {
    Logger.log('_updateInvoiceSheet_ 오류: ' + err.message);
    // 오류 시 무시 — 메인 입금 처리에 영향 없도록
  }
}

/* ───── 선충전 디버그 (Apps Script 에디터에서 직접 실행) ───────────
 * ROW_TO_TEST에 업무시트 실제 행 번호 입력 후 실행 → 로그 확인
 * 실행: Apps Script 에디터 → 함수 선택 → debugPrepaidCredit → ▶ 실행
 */
function debugPrepaidCredit() {
  const ROW_TO_TEST = 2; // ← 테스트할 업무시트 행 번호로 변경

  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const work = ss.getSheetByName('업무시트');
  const db   = ss.getSheetByName('고객_DB');

  if (!work) { Logger.log('❌ 업무시트 없음'); return; }
  if (!db)   { Logger.log('❌ 고객_DB 없음'); return; }

  const customer  = _norm(work.getRange(ROW_TO_TEST, 2).getValue());
  const amount    = _num( work.getRange(ROW_TO_TEST, 14).getValue());
  const payStatus = _norm(work.getRange(ROW_TO_TEST, 17).getValue());

  Logger.log(`=== 업무시트 ${ROW_TO_TEST}행 ===`);
  Logger.log(`B열(고객명): "${customer}"`);
  Logger.log(`N열(금액):   ${amount}`);
  Logger.log(`Q열(입금여부): "${payStatus}"`);

  if (!customer)          { Logger.log('⛔ STOP: 고객명 없음'); return; }
  if (!(amount > 0))      { Logger.log('⛔ STOP: 금액 없음 or 0'); return; }
  if (payStatus==='입금완료'){ Logger.log('⛔ STOP: 이미 입금완료'); return; }

  Logger.log('✅ 1차 조건 통과');

  const dbVals  = db.getDataRange().getValues();
  const header  = dbVals[0] || [];
  Logger.log(`고객_DB 헤더: ${JSON.stringify(header)}`);

  const nameIdx  = header.indexOf('담당자명');
  const bizIdx   = header.indexOf('사업자명');
  const aliasIdx = header.indexOf('별칭');
  Logger.log(`nameIdx=${nameIdx}, bizIdx=${bizIdx}, aliasIdx=${aliasIdx}`);

  let found = false;
  for (let i = 1; i < dbVals.length; i++) {
    const r    = dbVals[i];
    const name = nameIdx >= 0 ? _norm(r[nameIdx]) : '';
    const biz  = bizIdx  >= 0 ? _norm(r[bizIdx])  : '';
    if (!name && !biz) continue;
    if (name === customer || biz === customer) {
      const prepaid   = _num(r[6]); // G열 index 6
      const rawAlias  = aliasIdx >= 0 ? _norm(r[aliasIdx]) : '';
      const payerName = rawAlias.split(',')[0].trim() || name || biz;
      Logger.log(`✅ DB 매칭 성공 — DB행: ${i+1}, 담당자명: "${name}", 사업자명: "${biz}"`);
      Logger.log(`   선충전잔액(G열): ${prepaid}`);
      Logger.log(`   별칭: "${rawAlias}" → 입금자명: "${payerName}"`);
      if (!(prepaid > 0)) Logger.log('⛔ STOP: 선충전 잔액 없음 (0)');
      else if (prepaid >= amount) Logger.log(`✅ 처리 예정: 입금완료 + 잔액 ${prepaid} → ${prepaid - amount}`);
      else Logger.log(`⚠️ 잔액 부족: ${prepaid} < ${amount}, 부족금액: ${amount - prepaid}`);
      found = true;
      break;
    }
  }

  if (!found) {
    Logger.log(`⛔ STOP: 고객_DB에서 "${customer}" 매칭 실패`);
    Logger.log('--- DB 전체 담당자명/사업자명 목록 ---');
    for (let i = 1; i < dbVals.length; i++) {
      const r    = dbVals[i];
      const name = nameIdx >= 0 ? _norm(r[nameIdx]) : '';
      const biz  = bizIdx  >= 0 ? _norm(r[bizIdx])  : '';
      if (name || biz) Logger.log(`  DB행 ${i+1}: 담당자명="${name}", 사업자명="${biz}"`);
    }
  }
}

/* ───── 선충전잔액 추가 ─────────────────────────────────────────────
 * 고객_DB G열(선충전잔액)에 amount를 더한다.
 * APPLY_RECON_STACK_ 초과입금 처리에서 호출.
 */
function _addToPrepaid_(advertiser, amount) {
  if (!advertiser || !(amount > 0)) return;
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const db  = ss.getSheetByName('고객_DB');
    if (!db) return;

    const dbVals    = db.getDataRange().getValues();
    const header    = dbVals[0] || [];
    const nameIdx   = header.indexOf('담당자명');
    const bizIdx    = header.indexOf('사업자명');
    const prepaidIdx = header.indexOf('선충전잔액');
    if (prepaidIdx < 0) return;

    const advN = _norm(advertiser);
    for (let i = 1; i < dbVals.length; i++) {
      const r    = dbVals[i];
      const name = nameIdx >= 0 ? _norm(r[nameIdx]) : '';
      const biz  = bizIdx  >= 0 ? _norm(r[bizIdx])  : '';
      if (name === advN || biz === advN) {
        const current = _num(r[prepaidIdx]);
        db.getRange(i + 1, prepaidIdx + 1).setValue(current + amount);
        Logger.log(`✅ _addToPrepaid_: "${advertiser}" 선충전 ${current} + ${amount} = ${current + amount}`);
        return;
      }
    }
    Logger.log(`⚠️ _addToPrepaid_: "${advertiser}" 고객_DB 매칭 실패`);
  } catch(err) {
    Logger.log('_addToPrepaid_ 오류: ' + err.message);
  }
}
