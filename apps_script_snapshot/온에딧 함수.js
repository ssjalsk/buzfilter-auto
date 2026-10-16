/***********************
 * onEdit (통합 트리거 + 안전보강)
 ***********************/
function onEdit(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const nRows = e.range.getNumRows();
  const nCols = e.range.getNumColumns();

  /* ───────── 업무시트 ───────── */
  if (sheetName === "업무시트" && row > 1) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const 단가시트 = ss.getSheetByName("매체단가표");

    // B: 고객명 → F: 의뢰일 자동
    if (col === 2) {
      const dateCell = sheet.getRange(row, 6); // F
      const hasName = e.value && String(e.value).trim() !== "";
      if (hasName) {
        if (!dateCell.getValue()) {
          const today = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy. M. dd");
          dateCell.setValue(today);
        }
      } else {
        dateCell.clearContent();
      }
      if (typeof ensureDropdownSource_ === 'function') ensureDropdownSource_();
      return;
    }

    // R(입금자명) 편집 → Q(입금여부) '승격만' 허용(강등 금지)
    if (col === 18) {
      const qCell = sheet.getRange(row, 17);
      const prevQ = String(qCell.getValue() || '').trim();
      const rVal  = String(sheet.getRange(row, col).getValue() || '').trim();
      if (prevQ !== '입금완료' && rVal) qCell.setValue('입금완료'); // 승격만
      _updateStatusIfPublished_(sheet, row);
      return;
    }

    // M 입력 시 상태 보강 (M있음 → Q에 따라 상태 결정, 진행상태 제한 없음)
    if (col === 13) {
      const E = 5, Q = 17;
      const payQ = String(sheet.getRange(row, Q).getValue() || '').trim();
      const mVal = (e && e.value) ? String(e.value).trim() : String(sheet.getRange(row, 13).getValue() || '').trim();
      if (mVal !== '') {
        if (payQ === '입금완료') {
          sheet.getRange(row, E).setValue('송출완료');
        } else {
          sheet.getRange(row, E).setValue('입금확인');
          // ⚠️ setValue는 onEdit을 발동시키지 않으므로 선충전 처리 직접 호출
          _checkPrepaidCredit_(sheet, row);
        }
      }
      _updateStatusIfPublished_(sheet, row);
      return;
    }

    // E(진행상태) → '입금확인' 변경 시 → 선충전 잔액 자동 입금완료 처리
    // N열은 수식(price_mirror 참조)이라 onEdit 미발동 → E열 트리거로 처리
    // ※ M열 입력 시 setValue로 E열 변경 → onEdit 미발동 → M열 블록에서 직접 호출
    // ※ 사용자가 E열 드롭다운 직접 선택 시에만 이 블록 실행
    if (col === 5) {
      const statusVal = String(sheet.getRange(row, 5).getValue() || '').trim(); // e.value보다 안전
      if (statusVal === '입금확인') {
        _checkPrepaidCredit_(sheet, row);
      }
      return;
    }

    // 매체사/소재 → 실행사/원가 자동
    if (col === 9 || col === 11) {
      const 매체사 = String(sheet.getRange(row, 9).getValue() || '').trim();
      const 소재   = String(sheet.getRange(row, 11).getValue() || '').trim();
      if (!매체사 || !소재 || 소재 === "수작업") return;

      const last = Math.max(단가시트.getLastRow() - 1, 0);
      if (last === 0) return;
      const 단가표 = 단가시트.getRange(2, 1, last, 4).getValues();
      const 후보 = 단가표
        .filter(r => _norm(r[0]) === _norm(매체사) && _norm(r[3]) === _norm(소재))
        .map(r => ({ 실행사: r[1], 원가: _num(r[2]) }))
        .filter(r => r.원가 > 0);
      if (!후보.length) return;

      // [2026-10-16] 최저가 동점(서로 다른 실행사 2곳 이상) → 자동 배정하지 않고 담당자가 J열에서 선택
      const 최저가 = Math.min(...후보.map(r => r.원가));
      const 동점 = [...new Set(후보.filter(r => r.원가 === 최저가).map(r => _norm(r.실행사)))].filter(x => x);
      if (동점.length >= 2) {
        sheet.getRange(row, 10).setValue('');        // J 비움 → 선택 대기
        sheet.getRange(row, 15).setValue(최저가);     // O (동점이라 원가는 같음)
        _markExecutorTie_(sheet, row, 동점, 최저가);
        return;
      }

      const 최저 = 후보.reduce((a, b) => a.원가 < b.원가 ? a : b);
      sheet.getRange(row, 10).setValue(최저.실행사); // J
      sheet.getRange(row, 15).setValue(최저.원가);   // O
      _clearExecutorTie_(sheet, row);
    }

    // [2026-10-16] J열(실행사)을 직접 고르면 → 동점 선택 표시 해제 (원래 자유 입력 칸으로)
    if (col === 10 && String(sheet.getRange(row, 10).getValue() || '').trim()) {
      _clearExecutorTie_(sheet, row);
    }
    return;
  }

  /* ───────── 계산서 체크 시트 ───────── */
  if (sheetName === '계산서 체크' && row > 1) {

    // A열(광고주) 변경 시 → 고객_DB에서 사업자명 B열 자동채우기
    if (col === 1) {
      const customerName = _norm(sheet.getRange(row, 1).getValue());
      if (customerName) {
        const ss2 = SpreadsheetApp.getActiveSpreadsheet();
        const db2 = ss2.getSheetByName('고객_DB');
        if (db2) {
          const dbVals2  = db2.getDataRange().getValues();
          const header2  = dbVals2[0] || [];
          const nameIdx2 = header2.indexOf('담당자명');
          const bizIdx2  = header2.indexOf('사업자명');
          const aliasIdx2= header2.indexOf('별칭');

          for (let i = 1; i < dbVals2.length; i++) {
            const r2       = dbVals2[i];
            const name2    = nameIdx2  >= 0 ? _norm(r2[nameIdx2])  : '';
            const biz2     = bizIdx2   >= 0 ? _norm(r2[bizIdx2])   : '';
            const rawAlias2= aliasIdx2 >= 0 ? String(r2[aliasIdx2] || '') : '';
            const aliases2 = rawAlias2.split(',').map(a => _norm(a.trim())).filter(a => a);

            if (name2 === customerName || biz2 === customerName || aliases2.includes(customerName)) {
              const bizName = bizIdx2 >= 0 ? String(r2[bizIdx2] || '').trim() : '';
              if (bizName) sheet.getRange(row, 2).setValue(bizName); // B열: 사업자명
              break;
            }
          }
        }
      }
    }

    // B열·C열·H열 변경 시 → 시트 탭 색상 업데이트
    // B열+C열 모두 채워진 행에 H열(완료여부) 체크 안 된 건이 있으면 탭 빨간색
    if (col === 2 || col === 3 || col === 8) {
      _updateCalcSheetTabColor_(sheet);
    }

    return;
  }

  /* ───────── 종합 정산시트 ───────── */
  // ⚠️ 단순 onEdit은 UrlFetchApp 사용 불가 → G/K/H열 Cloud Function 호출은
  //    설치형 트리거(onSettleEdit)가 전담. 여기서는 드롭다운·집계만 처리.
  if (sheetName === '종합 정산시트') {
    // H/I 한 경로로 처리 (단일 셀만)
    if (row >= 4 && (col === 8 || col === 9)) {
      if (nRows !== 1 || nCols !== 1) return;
      APPLY_RECON_STACK_(row, col, e && e.value);
      return;
    }

    // F(광고주) 드롭다운 연결
    if (col === 6 && row >= 4) {
      const srcRange = (typeof ensureDropdownSource_ === 'function') ? ensureDropdownSource_() : null;
      if (srcRange) {
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInRange(srcRange, true)
          .setAllowInvalid(false)
          .setHelpText('업무시트의 고객명 목록에서 선택하세요.')
          .build();
        sheet.getRange(row, 6).setDataValidation(rule);
      }
      return;
    }

    // 상단 집계 영역 변경 시
    if (row === 3 && col >= 14 && col <= 17) {
      if (typeof 집계_자동화 === 'function') 집계_자동화();
    }
    return;
  }
}

// ─── 계산서 체크 탭 색상 업데이트 ────────────────────────────────
/**
 * '계산서 체크' 시트에서
 * B열(사업자명) + C열(일반금액) 모두 채워진 행에
 * H열(완료여부) 체크박스가 FALSE인 건이 1개라도 있으면 탭 빨간색,
 * 전부 체크되어 있으면 탭 기본색으로 복원.
 *
 * @param {Sheet} sheet - '계산서 체크' 시트 객체
 */
function _updateCalcSheetTabColor_(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      sheet.setTabColor(null);
      return;
    }

    // B(2)~H(8)열 한 번에 읽기
    const vals = sheet.getRange(2, 2, lastRow - 1, 7).getValues();
    // vals[i][0]=B열, vals[i][1]=C열, vals[i][6]=H열

    let hasUnchecked = false;
    for (let i = 0; i < vals.length; i++) {
      const bVal = String(vals[i][0] || '').trim();       // B: 사업자명
      const cVal = vals[i][1];                            // C: 일반금액
      const hVal = vals[i][6];                            // H: 완료여부 체크박스

      // B열과 C열이 모두 채워진 행인지 확인
      const bFilled = bVal !== '';
      const cFilled = cVal !== '' && cVal !== null && cVal !== undefined;

      if (bFilled && cFilled && hVal !== true) {
        hasUnchecked = true;
        break;
      }
    }

    sheet.setTabColor(hasUnchecked ? '#ff0000' : null);
  } catch (e) {
    // 탭 색상 오류는 다른 기능에 영향 주지 않도록 조용히 처리
  }
}

/* ───── [2026-10-16] 최저가 동점 실행사 선택 ─────────────────────────────
 * 동점이면 J열에 동점 실행사만 선택 목록 + 연노랑 + 메모. 목록 밖 실행사도 직접 입력 가능(경고만).
 * 실행사를 고르거나 동점이 아니게 되면 원래대로(선택 목록 없음, 색·메모 제거).
 * 송출 동기화(HYB_Trigger)도 동점이면 배정하지 않고 J열 선택을 기다림 (_autoExecutor_ 의 tie).
 */
const TIE_NOTE_MARK = '⚖️ 최저가 동점';
function _markExecutorTie_(sheet, row, executors, cost) {
  const cell = sheet.getRange(row, 10);
  cell.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(executors, true).setAllowInvalid(true).build());
  cell.setBackground('#fff2cc');
  cell.setNote(`${TIE_NOTE_MARK} — 실행사를 선택해 주세요\n${executors.join(' / ')} (원가 ${Number(cost).toLocaleString()}원)`);
}
function _clearExecutorTie_(sheet, row) {
  const cell = sheet.getRange(row, 10);
  if (!String(cell.getNote() || '').startsWith(TIE_NOTE_MARK)) return;   // 동점 표시한 칸만 되돌림
  cell.setDataValidation(null);
  cell.setBackground(null);
  cell.setNote('');
}
