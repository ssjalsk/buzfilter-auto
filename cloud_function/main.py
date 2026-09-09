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
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet(DB_SHEET)  # 원본에서 읽기만
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


def match_customer(payer, customers):
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
    # 2차: 부분 포함
    for cust in customers:
        aliases = [a.strip() for a in str(cust.get("별칭", "")).split(",") if a.strip()]
        fields  = aliases + [cust.get("담당자명",""), cust.get("사업자명",""), cust.get("대표자명","")]
        for f in fields:
            fn = norm(f)
            if fn and (fn in pn or pn in fn):
                return cust.get("담당자명","") or cust.get("사업자명","")
    return None


def write_advertiser(gc, row_idx, advertiser):
    ws = gc.open_by_key(SETTLE_SPREADSHEET_ID).worksheet(SETTLE_SHEET)
    ws.update_cell(row_idx, 6, advertiser)


def add_prepayment(gc, advertiser, amount_n):
    """
    고객_DB 선충전잔액 컬럼에 amount_n(부가세 제외 금액) 추가.
    반환: True=성공, False=컬럼없음 or 고객없음
    """
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet(DB_SHEET)  # 원본에서 읽기만
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


def update_work_sheet(gc, advertiser, payer, amount_k=0):
    """
    업무시트 입금 처리 — N열(부가세 제외) 기준 금액 순차 할당.

    파라미터:
      advertiser : 매칭된 광고주명 (고객_DB 담당자명/사업자명)
      payer      : 입금자명 (R열에 기록)
      amount_k   : 종합 정산시트 K열 부가세포함 입금액

    처리 흐름:
      1. remaining_n = round(amount_k / 1.1)  ← 부가세 제외 잔액으로 변환
      2. 해당 광고주 미처리 행(R열 비어있음) 위→아래 순서로 수집
      3. 각 행: needed = N열 - S열(기존 부분입금 누적)
         - remaining_n >= needed → 완전처리 (Q=입금완료, R=입금자명, S=초기화, M있으면 E=송출완료)
         - remaining_n < needed  → 부분처리 (S열에 새 누적값 기록, 중단)
      4. 처리 후 remaining_n 남으면 → 고객_DB 선충전잔액에 추가

    반환: {"processed": 완전처리행수, "prepaid": 선충전추가금액}
    """
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet(WORK_SHEET)
    all_vals = ws.get_all_values()
    if len(all_vals) < 2:
        return {"processed": 0, "prepaid": 0}

    adv_n = norm(advertiser)
    # 부가세 제외 잔액 (N열과 동일 단위)
    remaining_n = round(amount_k / 1.1) if amount_k > 0 else 0

    # ── 해당 광고주 미처리 행 수집 ──────────────────────────────────
    # 컬럼 인덱스 (0-based): B=1, E=4, M=12, N=13, Q=16, R=17, S=18
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
        m_val = row[12].strip() if len(row) > 12 else ""          # M열: 링크
        s_val = parse_amount(row[18] if len(row) > 18 else "")    # S열: 기존 부분입금 누적
        candidates.append({
            "row":   i,
            "n_val": n_val,
            "m_val": m_val,
            "s_val": s_val,
        })

    # ── 순차 할당 ────────────────────────────────────────────────────
    updates  = []
    processed = 0

    for c in candidates:
        if remaining_n <= 0:
            break

        needed = c["n_val"] - c["s_val"]  # 이 행을 완전처리하는 데 필요한 잔액

        if needed <= 0:
            # S열이 이미 꽉 찼으면(예외 케이스) → 완전처리만
            updates.append({"range": f"R{c['row']}", "values": [[payer]]})
            updates.append({"range": f"Q{c['row']}", "values": [["입금완료"]]})
            updates.append({"range": f"S{c['row']}", "values": [[""]]})
            if c["m_val"]:
                updates.append({"range": f"E{c['row']}", "values": [["송출완료"]]})
            processed += 1
            continue

        if remaining_n >= needed:
            # 완전처리
            updates.append({"range": f"R{c['row']}", "values": [[payer]]})
            updates.append({"range": f"Q{c['row']}", "values": [["입금완료"]]})
            updates.append({"range": f"S{c['row']}", "values": [[""]]})
            if c["m_val"]:
                updates.append({"range": f"E{c['row']}", "values": [["송출완료"]]})
            remaining_n -= needed
            processed += 1
        else:
            # 부분처리: S열에 누적금액 기록
            new_s = c["s_val"] + remaining_n
            updates.append({"range": f"S{c['row']}", "values": [[str(new_s)]]})
            remaining_n = 0

    if updates:
        ws.batch_update(updates)

    # ── 선충전: 남은 잔액 → 고객_DB 선충전잔액 ──────────────────────
    prepaid = 0
    if remaining_n > 0:
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
    amount_k = parse_amount(data.get("amount_k", 0))  # K열 부가세포함 금액

    if not row or not payer:
        return (json.dumps({"ok": False, "error": "row/payer 필수"}), 400,
                {"Content-Type": "application/json"})

    try:
        gc        = get_gc()

        # amount_k=0이면 종합 정산시트 K열 직접 읽기 (G열 먼저 입력 시 대비)
        if amount_k == 0:
            try:
                settle_ws = gc.open_by_key(SETTLE_SPREADSHEET_ID).worksheet(SETTLE_SHEET)
                k_raw = settle_ws.cell(int(row), 11).value  # K열=11번
                amount_k = parse_amount(k_raw or 0)
            except Exception:
                pass  # 읽기 실패 시 0으로 유지

        customers = load_customer_db(gc)
        matched   = match_customer(payer, customers)

        if matched:
            write_advertiser(gc, int(row), matched)
            result = update_work_sheet(gc, matched, payer, amount_k)
            return (json.dumps({
                "ok":       True,
                "matched":  matched,
                "processed": result["processed"],
                "prepaid":  result["prepaid"],
            }), 200, {"Content-Type": "application/json"})
        else:
            return (json.dumps({"ok": True, "matched": None,
                                "msg": f"'{payer}' 매칭 고객 없음 (DB: {len(customers)}명)"}), 200,
                    {"Content-Type": "application/json"})

    except Exception as e:
        return (json.dumps({"ok": False, "error": str(e)}), 500,
                {"Content-Type": "application/json"})
