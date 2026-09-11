import os, json, re, functions_framework
from google.oauth2.service_account import Credentials
import gspread

SPREADSHEET_ID        = "1mOV-HmlODZaxPohiVFay9_Rnh31-vhcPXF_b-5-EjB0"  # 테스트 시트 — 고객_DB/업무시트/종합 정산시트 전부
SETTLE_SPREADSHEET_ID = "1mOV-HmlODZaxPohiVFay9_Rnh31-vhcPXF_b-5-EjB0"  # 동일
DB_SHEET       = "고객_DB"
WORK_SHEET     = "업무시트"
SETTLE_SHEET   = "종합 정산시트"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]


def get_gc():
    raw = os.environ.get("GOOGLE_CREDENTIALS", "")
    if not raw:
        raise RuntimeError("GOOGLE_CREDENTIALS 환경변수 없음")
    cred_dict = json.loads(raw)
    creds = Credentials.from_service_account_info(cred_dict, scopes=SCOPES)
    return gspread.authorize(creds)


def norm(s):
    return re.sub(r"[\s\-_·・()\[\]]", "", str(s)).lower()


def parse_amount(val):
    """금액 문자열에서 정수 추출 (쉼표/원/공백 등 제거)"""
    cleaned = re.sub(r"[^\d.]", "", str(val))
    try:
        return int(float(cleaned)) if cleaned else 0
    except Exception:
        return 0


def load_customer_db(gc):
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet(DB_SHEET)
    rows = ws.get_all_values()
    if len(rows) < 2:
        return []
    headers = rows[0]
    result = []
    for r in rows[1:]:
        if not any(r):
            continue
        d = {headers[i]: (r[i] if i < len(r) else "") for i in range(len(headers))}
        result.append(d)
    return result


def _get_unpaid_amounts(gc, adv_n):
    """업무시트 + 리뷰 업무시트에서 해당 고객의 미수 청구액 목록 반환"""
    amounts = []
    try:
        ws = gc.open_by_key(SPREADSHEET_ID).worksheet(WORK_SHEET)
        rows = ws.get_all_values()
        for row in rows[1:]:
            if len(row) < 14: continue
            if norm(row[1]) != adv_n: continue          # B열: 담당자명
            if str(row[4]).strip() != '입금확인': continue  # E열: 진행상태
            if len(row) > 17 and str(row[17]).strip(): continue  # R열: 이미 처리
            n_val = parse_amount(row[13] if len(row) > 13 else "")
            s_val = parse_amount(row[18] if len(row) > 18 else "")
            remaining = n_val - s_val
            if remaining > 0:
                amounts.append(remaining)
    except Exception:
        pass
    try:
        rws = gc.open_by_key(SPREADSHEET_ID).worksheet("리뷰 업무시트")
        rows = rws.get_all_values()
        for row in rows[1:]:
            if len(row) < 17: continue
            if norm(row[1]) != adv_n: continue          # B열: 담당자명
            t_val = str(row[19]).strip() if len(row) > 19 else ""
            if t_val == '입금완료': continue             # T열: 이미 완료
            q_val = parse_amount(row[16] if len(row) > 16 else "")
            if q_val > 0:
                amounts.append(q_val)
    except Exception:
        pass
    return amounts


def _amount_score(check_amount, amounts):
    """입금액과 미수 청구액 목록의 일치도 점수(0~100) 반환"""
    if not amounts or check_amount <= 0:
        return 0
    total = sum(amounts)
    # 개별 청구액 완전 일치
    if check_amount in amounts:
        return 100
    # 합계 완전 일치
    if check_amount == total:
        return 90
    # 개별 10% 이내
    for amt in amounts:
        if amt > 0 and abs(check_amount - amt) / amt <= 0.1:
            return 70
    # 합계 10% 이내
    if total > 0 and abs(check_amount - total) / total <= 0.1:
        return 60
    return 0


def match_customer(payer, customers, amount_k=0, amount_h=0, gc=None):
    pn = norm(payer)
    if not pn:
        return None

    # 1차: 완전 일치 (별칭, 담당자명, 사업자명, 대표자명)
    for cust in customers:
        aliases = [a.strip() for a in str(cust.get("별칭", "")).split(",") if a.strip()]
        fields  = aliases + [cust.get("담당자명",""), cust.get("사업자명",""), cust.get("대표자명","")]
        for f in fields:
            if norm(f) == pn:
                return cust.get("담당자명","") or cust.get("사업자명","")

    # 2차: 부분 포함 → 후보 전부 수집
    candidates = []
    for cust in customers:
        aliases = [a.strip() for a in str(cust.get("별칭", "")).split(",") if a.strip()]
        fields  = aliases + [cust.get("담당자명",""), cust.get("사업자명",""), cust.get("대표자명","")]
        for f in fields:
            fn = norm(f)
            if fn and (fn in pn or pn in fn):
                candidates.append(cust)
                break

    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0].get("담당자명","") or candidates[0].get("사업자명","")

    # 3차: 중복 후보 → 업무시트·리뷰 업무시트 입금액 기준 2차 검수
    if gc and (amount_k > 0 or amount_h > 0):
        check_amount = amount_h if amount_h > 0 else round(amount_k / 1.1)
        best_name, best_score = None, -1
        for cust in candidates:
            adv_name = cust.get("담당자명","") or cust.get("사업자명","")
            unpaid   = _get_unpaid_amounts(gc, norm(adv_name))
            score    = _amount_score(check_amount, unpaid)
            if score > best_score:
                best_score = score
                best_name  = adv_name
        if best_name and best_score >= 60:
            return best_name

    # fallback: 첫 번째 후보 반환
    return candidates[0].get("담당자명","") or candidates[0].get("사업자명","")


def write_advertiser(gc, row_idx, advertiser):
    ws = gc.open_by_key(SETTLE_SPREADSHEET_ID).worksheet(SETTLE_SHEET)
    ws.update_cell(row_idx, 6, advertiser)


def add_prepayment(gc, advertiser, amount_n):
    """
    고객_DB 선충전잔액 컬럼에 amount_n(부가세 제외 금액) 추가.
    반환: True=성공, False=컬럼없음 or 고객없음
    """
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet(DB_SHEET)
    rows = ws.get_all_values()
    if len(rows) < 2:
        return False
    headers = rows[0]
    if "선충전잔액" not in headers:
        return False

    pre_col_idx = headers.index("선충전잔액")          # 0-based
    adv_n = norm(advertiser)

    for i, row in enumerate(rows[1:], start=2):
        담당자_idx = headers.index("담당자명") if "담당자명" in headers else -1
        사업자_idx = headers.index("사업자명") if "사업자명" in headers else -1
        담당자 = (row[담당자_idx] if 담당자_idx >= 0 and len(row) > 담당자_idx else "")
        사업자 = (row[사업자_idx] if 사업자_idx >= 0 and len(row) > 사업자_idx else "")
        if norm(담당자) == adv_n or norm(사업자) == adv_n:
            current = parse_amount(row[pre_col_idx] if len(row) > pre_col_idx else "")
            ws.update_cell(i, pre_col_idx + 1, current + amount_n)
            return True
    return False


def update_work_sheet(gc, advertiser, payer, amount_k=0, amount_h=0):
    """
    업무시트 입금 처리 — N열(부가세 제외) 기준 금액 순차 할당.

    파라미터:
      advertiser : 매칭된 광고주명 (고객_DB 담당자명/사업자명)
      payer      : 입금자명 (R열에 기록)
      amount_k   : 종합 정산시트 K열 부가세포함 입금액
      amount_h   : 종합 정산시트 H열 미발행 입금액 (부가세 없음, 원금 그대로)

    처리 흐름:
      1. remaining_n 계산
         - amount_h > 0 → remaining_n = amount_h (부가세 없음)
         - amount_k > 0 → remaining_n = round(amount_k / 1.1) (부가세 제외)
      2. 처리된 행(R열 채워짐) 중 S열 잔액 → remaining_n에 합산, 해당 S열 초기화
      3. 미처리 행(R열 비어있음) 위→아래 순서로 FIFO 할당
         - remaining_n >= needed → 완전처리 (Q=입금완료, R=입금자명, M있으면 E=송출완료)
         - remaining_n < needed  → 부분처리 (S열에 누적값 기록, 중단)
      4. 완전처리 후 remaining_n 남으면 → 마지막 처리된 행 S열에 승계금액 기록
      5. 처리할 미처리 행 자체가 없으면 → 고객_DB 선충전잔액에 추가

    반환: {"processed": 완전처리행수, "prepaid": 선충전추가금액}
    """
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet(WORK_SHEET)
    all_vals = ws.get_all_values()
    if len(all_vals) < 2:
        return {"processed": 0, "prepaid": 0}

    adv_n = norm(advertiser)

    # ── 입금 금액 계산 ────────────────────────────────────────────────
    # H열(미발행): 부가세 없이 원금 그대로
    # K열(발행): 부가세 포함 → /1.1로 제외
    if amount_h > 0:
        remaining_n = amount_h
    elif amount_k > 0:
        remaining_n = round(amount_k / 1.1)
    else:
        remaining_n = 0

    # ── 처리된 행 S열 잔액(승계금액) 수집 및 합산 ──────────────────
    # 컬럼 인덱스 (0-based): B=1, E=4, M=12, N=13, Q=16, R=17, S=18
    prepaid_rows = []  # R열 채워진 행 중 S열에 잔액 있는 행
    if remaining_n > 0:
        for i, row in enumerate(all_vals[1:], start=2):
            b_val = norm(row[1] if len(row) > 1 else "")
            if b_val != adv_n:
                continue
            r_val = row[17].strip() if len(row) > 17 else ""
            if not r_val:
                continue  # 미처리 행 제외 (처리된 행만 확인)
            s_val = parse_amount(row[18] if len(row) > 18 else "")
            if s_val > 0:
                prepaid_rows.append({"row": i, "s_val": s_val})

        # 승계잔액을 remaining_n에 합산
        total_balance = sum(p["s_val"] for p in prepaid_rows)
        remaining_n += total_balance

    # ── 미처리 행 수집 ───────────────────────────────────────────────
    candidates = []
    for i, row in enumerate(all_vals[1:], start=2):
        b_val = norm(row[1] if len(row) > 1 else "")
        if b_val != adv_n:
            continue
        r_val = row[17].strip() if len(row) > 17 else ""  # R열: 입금자명
        if r_val:
            continue  # 이미 입금 처리된 행 제외
        n_val = parse_amount(row[13] if len(row) > 13 else "")  # N열: 부가세 제외 금액
        if n_val <= 0:
            continue  # 금액 없는 행 제외
        m_val = row[12].strip() if len(row) > 12 else ""       # M열: 링크
        s_val = parse_amount(row[18] if len(row) > 18 else "")  # S열: 기존 부분입금 누적
        candidates.append({
            "row":   i,
            "n_val": n_val,
            "m_val": m_val,
            "s_val": s_val,
        })

    # ── 순차 할당 ────────────────────────────────────────────────────
    updates          = []
    processed        = 0
    last_processed_row = None

    # 승계잔액을 소비했으므로 해당 행 S열 초기화
    for p in prepaid_rows:
        updates.append({"range": f"S{p['row']}", "values": [[""]]})

    for c in candidates:
        if remaining_n <= 0:
            break

        needed = c["n_val"] - c["s_val"]  # 완전처리에 필요한 잔액

        if needed <= 0:
            # S열이 이미 충족된 예외 케이스 → 완전처리
            updates.append({"range": f"R{c['row']}", "values": [[payer]]})
            updates.append({"range": f"Q{c['row']}", "values": [["입금완료"]]})
            updates.append({"range": f"S{c['row']}", "values": [[""]]})
            if c["m_val"]:
                updates.append({"range": f"E{c['row']}", "values": [["송출완료"]]})
            processed += 1
            last_processed_row = c["row"]
            continue

        if remaining_n >= needed:
            # 완전처리
            updates.append({"range": f"R{c['row']}", "values": [[payer]]})
            updates.append({"range": f"Q{c['row']}", "values": [["입금완료"]]})
            updates.append({"range": f"S{c['row']}", "values": [[""]]})  # 임시 초기화 (아래서 덮어씀)
            if c["m_val"]:
                updates.append({"range": f"E{c['row']}", "values": [["송출완료"]]})
            remaining_n -= needed
            processed += 1
            last_processed_row = c["row"]
        else:
            # 부분처리: S열에 누적금액 기록
            new_s = c["s_val"] + remaining_n
            updates.append({"range": f"S{c['row']}", "values": [[str(new_s)]]})
            remaining_n = 0

    # ── 완전처리 후 남은 승계금액 → 마지막 처리된 행 S열에 기록 ────
    if remaining_n > 0 and last_processed_row is not None:
        # 이미 S="" 업데이트가 있으면 제거 후 잔액으로 덮어쓰기
        updates = [u for u in updates if u["range"] != f"S{last_processed_row}"]
        updates.append({"range": f"S{last_processed_row}", "values": [[str(remaining_n)]]})

    if updates:
        ws.batch_update(updates)

    # ── 선충전: 처리할 미처리 행 자체가 없는 경우 ───────────────────
    prepaid = 0
    if remaining_n > 0 and last_processed_row is None:
        # 미처리 행이 아예 없을 때만 선충전잔액으로 처리
        ok = add_prepayment(gc, advertiser, remaining_n)
        if ok:
            prepaid = remaining_n

    return {"processed": processed, "prepaid": prepaid}


@functions_framework.http
def match_payment(request):
    if request.method == "OPTIONS":
        return ("", 204, {"Access-Control-Allow-Origin": "*",
                          "Access-Control-Allow-Methods": "POST",
                          "Access-Control-Allow-Headers": "Content-Type"})

    data   = request.get_json(silent=True) or {}
    secret = os.environ.get("FUNC_SECRET", "aligo-secret-2026")
    if data.get("secret") != secret:
        return (json.dumps({"ok": False, "error": "unauthorized"}), 403,
                {"Content-Type": "application/json"})

    row      = data.get("row")
    payer    = str(data.get("payer", "")).strip()
    amount_k = parse_amount(data.get("amount_k", 0))  # K열 부가세포함
    amount_h = parse_amount(data.get("amount_h", 0))  # H열 미발행 (부가세 없음)

    if not row or not payer:
        return (json.dumps({"ok": False, "error": "row/payer 필수"}), 400,
                {"Content-Type": "application/json"})

    try:
        gc = get_gc()

        # amount_k=0, amount_h=0이면 종합 정산시트에서 직접 읽기 (G열 먼저 입력 시 대비)
        if amount_k == 0 and amount_h == 0:
            try:
                settle_ws = gc.open_by_key(SETTLE_SPREADSHEET_ID).worksheet(SETTLE_SHEET)
                k_raw = settle_ws.cell(int(row), 11).value  # K열=11번
                h_raw = settle_ws.cell(int(row), 8).value   # H열=8번
                amount_k = parse_amount(k_raw or 0)
                amount_h = parse_amount(h_raw or 0)
            except Exception:
                pass  # 읽기 실패 시 0으로 유지

        customers = load_customer_db(gc)
        matched   = match_customer(payer, customers, amount_k=amount_k, amount_h=amount_h, gc=gc)

        if matched:
            write_advertiser(gc, int(row), matched)
            result = update_work_sheet(gc, matched, payer, amount_k, amount_h)
            return (json.dumps({
                "ok":        True,
                "matched":   matched,
                "processed": result["processed"],
                "prepaid":   result["prepaid"],
            }), 200, {"Content-Type": "application/json"})
        else:
            return (json.dumps({"ok": True, "matched": None,
                                "msg": f"'{payer}' 매칭 고객 없음 (DB: {len(customers)}명)"}), 200,
                    {"Content-Type": "application/json"})

    except Exception as e:
        return (json.dumps({"ok": False, "error": str(e)}), 500,
                {"Content-Type": "application/json"})
