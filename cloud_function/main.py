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


def _parse_s_val(val):
    """업무시트 S열 읽기 전용: 숫자만 누적입금액으로 인정.
    '₩XXX 부족' '초과입금 X' 같은 상태 텍스트는 0 반환 — parse_amount와 구분."""
    s = str(val).strip()
    if not s:
        return 0
    if "부족" in s or "초과" in s:
        return 0
    return parse_amount(s)


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


_CORP = re.compile(r"주식회사|유한회사|유한책임회사|\(주\)|\(유\)|㈜|㈲|（주）|（유）")


def norm_corp(s):
    """회사형태 표기('(주)', '주식회사', '㈜', '유한회사', '(유)')를 뗀 비교용 문자열"""
    return norm(_CORP.sub("", str(s)))


def _same(a, b):
    """완전 일치: 그대로 비교 또는 회사형태 표기를 뗀 비교 중 하나라도 같으면 일치"""
    na, nb = norm(a), norm(b)
    if na and na == nb:
        return True
    ca, cb = norm_corp(a), norm_corp(b)
    return bool(ca) and ca == cb


def _overlap(payer, field):
    """부분 일치 길이 (한쪽이 다른 쪽에 포함될 때 짧은 쪽 길이). 0이면 불일치.
    ※ 회사형태를 뗀 비교는 부분일치에 쓰지 않음 — '주식회사 미소' → '미소'처럼 짧아져 엉뚱한 고객과 걸리는 것 방지"""
    p, f = norm(payer), norm(field)
    if p and f and (f in p or p in f):
        return min(len(p), len(f))
    return 0


def _fields(cust):
    """매칭 비교 대상: 별칭(쉼표 구분) + 담당자명 + 사업자명 + 대표자명"""
    aliases = [a.strip() for a in str(cust.get("별칭", "")).split(",") if a.strip()]
    return aliases + [cust.get("담당자명",""), cust.get("사업자명",""), cust.get("대표자명","")]


def _cust_name(cust):
    return cust.get("담당자명","") or cust.get("사업자명","")


def _pick_by_unpaid(gc, candidates, amount_k, amount_h):
    """
    [2026-09-30] 입금자명이 여러 고객과 겹칠 때(예: 같은 사업자의 병원/치과 담당자 분리)
    업무시트·리뷰 업무시트 미입금 금액으로 '확실히 한 명'을 고를 수 있을 때만 반환.
    최고 점수(60점 이상) 후보가 딱 한 명이 아니면 None → 수동확인 처리.
    """
    if not gc or (amount_k <= 0 and amount_h <= 0):
        return None
    check_amount = amount_h if amount_h > 0 else round(amount_k / 1.1)
    scored = []
    for cust in candidates:
        unpaid = _get_unpaid_amounts(gc, norm(_cust_name(cust)))
        scored.append((_amount_score(check_amount, unpaid), cust))
    best = max(s for s, _ in scored)
    top  = [c for s, c in scored if s == best]
    if best >= 60 and len(top) == 1:
        return top[0]
    return None


def match_customer(payer, customers, amount_k=0, amount_h=0, gc=None):
    """
    반환: {"status": "matched" | "manual" | "none", "cust": 고객dict|None, "candidates": [고객dict]}
      matched : 한 명으로 확정
      manual  : 여러 고객과 겹치는데 미입금 금액으로도 구분 불가 → 추측하지 않고 수동확인
      none    : 일치하는 고객 없음
    """
    pn = norm(payer)
    if not pn:
        return {"status": "none", "cust": None, "candidates": []}

    # 1차: 완전 일치 — 겹치는 고객이 있는지 끝까지 전부 수집 (예전: 위쪽 고객 하나로 즉시 확정)
    #      '(주)지엠이엠' 과 '지엠이엠' 처럼 회사형태 표기만 다른 것도 같은 이름으로 봄
    exact = [c for c in customers if any(_same(payer, f) for f in _fields(c))]
    if len(exact) == 1:
        return {"status": "matched", "cust": exact[0], "candidates": exact}
    if len(exact) > 1:
        pick = _pick_by_unpaid(gc, exact, amount_k, amount_h)
        if pick:
            return {"status": "matched", "cust": pick, "candidates": exact}
        return {"status": "manual", "cust": None, "candidates": exact}

    # 2차: 부분 포함 → 후보 전부 수집 (고객별 가장 길게 겹친 길이)
    scored = []
    for cust in customers:
        ov = max((_overlap(payer, f) for f in _fields(cust)), default=0)
        if ov:
            scored.append((ov, cust))

    if not scored:
        return {"status": "none", "cust": None, "candidates": []}
    # 더 길게 겹친 후보 우선 (예: '애드닷 환불' → '애드'(2자) 보다 '애드닷'(3자))
    top_ov     = max(ov for ov, _ in scored)
    candidates = [c for ov, c in scored if ov == top_ov]
    if len(candidates) == 1:
        return {"status": "matched", "cust": candidates[0], "candidates": candidates}

    # 3차: 그래도 여러 명 → 미입금 금액 기준 검수 (예전: 구분 안 되면 첫 번째 후보로 추측)
    pick = _pick_by_unpaid(gc, candidates, amount_k, amount_h)
    if pick:
        return {"status": "matched", "cust": pick, "candidates": candidates}
    return {"status": "manual", "cust": None, "candidates": candidates}


def write_advertiser(gc, row_idx, advertiser):
    ws = gc.open_by_key(SETTLE_SPREADSHEET_ID).worksheet(SETTLE_SHEET)
    ws.update_cell(row_idx, 6, advertiser)


MANUAL_MARK = "⚠️수동확인"


def write_manual_check(gc, row_idx, names):
    """
    여러 고객과 겹쳐 자동 판단 불가 → F열에 수동확인 표시 + 메모로 후보 안내.
    업무시트 입금처리·선충전 적립은 하지 않음.
    """
    ws = gc.open_by_key(SETTLE_SPREADSHEET_ID).worksheet(SETTLE_SHEET)
    ws.update_cell(row_idx, 6, MANUAL_MARK)
    ws.update_note(f"F{row_idx}",
                   "입금자명이 여러 고객과 겹쳐 자동으로 판단하지 못했습니다.\n"
                   "후보: " + " / ".join(names) + "\n"
                   "업무시트 미입금 건 확인 후 F열을 지우고 올바른 광고주를 선택해 주세요.")


def _get_prepaid_info(gc, advertiser):
    """
    고객_DB에서 선충전잔액 조회.
    반환: (balance, db_row_1indexed, db_col_1indexed)
          고객 없거나 컬럼 없으면 (0, -1, -1)
    """
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet(DB_SHEET)
    rows = ws.get_all_values()
    if len(rows) < 2:
        return 0, -1, -1
    headers = rows[0]
    if "선충전잔액" not in headers:
        return 0, -1, -1
    pre_col_idx = headers.index("선충전잔액")
    adv_n = norm(advertiser)
    담당자_idx = headers.index("담당자명") if "담당자명" in headers else -1
    사업자_idx = headers.index("사업자명") if "사업자명" in headers else -1
    for i, row in enumerate(rows[1:], start=2):
        담당자 = row[담당자_idx] if 담당자_idx >= 0 and len(row) > 담당자_idx else ""
        사업자 = row[사업자_idx] if 사업자_idx >= 0 and len(row) > 사업자_idx else ""
        if norm(담당자) == adv_n or norm(사업자) == adv_n:
            balance = parse_amount(row[pre_col_idx] if len(row) > pre_col_idx else "")
            return balance, i, pre_col_idx + 1
    return 0, -1, -1


def add_prepayment(gc, advertiser, amount_n):
    """
    고객_DB 선충전잔액 컬럼에 amount_n(부가세 제외 금액) 추가.
    반환: True=성공, False=컬럼없음 or 고객없음
    """
    balance, db_row, db_col = _get_prepaid_info(gc, advertiser)
    if db_row < 0:
        return False
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet(DB_SHEET)
    ws.update_cell(db_row, db_col, balance + amount_n)
    return True


def _build_name_set(advertiser, cust_info):
    """
    업무시트 B열 매칭용 norm된 이름 집합 반환.
    고객_DB의 담당자명, 사업자명, 별칭 전체를 포함해
    B열에 별칭이 입력된 경우도 올바르게 매칭.
    """
    names = set()
    names.add(norm(advertiser))
    if cust_info:
        for field in ["담당자명", "사업자명", "대표자명"]:
            v = norm(cust_info.get(field, ""))
            if v:
                names.add(v)
        for alias in str(cust_info.get("별칭", "")).split(","):
            v = norm(alias.strip())
            if v:
                names.add(v)
    names.discard("")
    return names


def update_work_sheet(gc, advertiser, payer, amount_k=0, amount_h=0, cust_info=None):
    """
    업무시트 입금 처리 — N열(부가세 제외) 기준 금액 순차 할당.

    파라미터:
      advertiser : 매칭된 광고주명 (고객_DB 담당자명/사업자명)
      payer      : 입금자명 (R열에 기록)
      amount_k   : 종합 정산시트 K열 부가세포함 입금액
      amount_h   : 종합 정산시트 H열 미발행 입금액 (부가세 없음, 원금 그대로)
      cust_info  : 고객_DB 행 dict (담당자명/사업자명/별칭 포함) — 별칭 B열 매칭용

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

    # B열 매칭: 담당자명, 사업자명, 별칭 모두 포함
    name_set = _build_name_set(advertiser, cust_info)

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
            if b_val not in name_set:
                continue
            r_val = row[17].strip() if len(row) > 17 else ""
            if not r_val:
                continue  # 미처리 행 제외 (처리된 행만 확인)
            s_val = _parse_s_val(row[18] if len(row) > 18 else "")
            if s_val > 0:
                prepaid_rows.append({"row": i, "s_val": s_val})

        # 승계잔액을 remaining_n에 합산
        total_balance = sum(p["s_val"] for p in prepaid_rows)
        remaining_n += total_balance

    # ── 미처리 행 수집 ───────────────────────────────────────────────
    candidates = []
    for i, row in enumerate(all_vals[1:], start=2):
        b_val = norm(row[1] if len(row) > 1 else "")
        if b_val not in name_set:
            continue
        r_val = row[17].strip() if len(row) > 17 else ""  # R열: 입금자명
        if r_val:
            continue  # 이미 입금 처리된 행 제외
        n_val = parse_amount(row[13] if len(row) > 13 else "")  # N열: 부가세 제외 금액
        if n_val <= 0:
            continue  # 금액 없는 행 제외
        m_val = row[12].strip() if len(row) > 12 else ""       # M열: 링크
        s_val = _parse_s_val(row[18] if len(row) > 18 else "")  # S열: 기존 부분입금 누적
        candidates.append({
            "row":   i,
            "n_val": n_val,
            "m_val": m_val,
            "s_val": s_val,
        })

    # ── 순차 할당 ────────────────────────────────────────────────────
    updates   = []
    processed = 0

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
            # 부분처리 — remaining_n만큼 S열에 누적
            new_s = c["s_val"] + remaining_n
            updates.append({"range": f"S{c['row']}", "values": [[new_s]]})
            remaining_n = 0
            break

    if updates:
        ws.batch_update(updates)

    # ── 완전처리 후 남은 금액(초과입금) → 선충전잔액에 추가 ──────────
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
        res       = match_customer(payer, customers, amount_k=amount_k, amount_h=amount_h, gc=gc)

        if res["status"] == "manual":
            names = [_cust_name(c) for c in res["candidates"]]
            write_manual_check(gc, int(row), names)
            return (json.dumps({"ok": True, "matched": None, "manual": True, "candidates": names,
                                "msg": f"'{payer}' 여러 고객과 겹침 → 수동확인"}), 200,
                    {"Content-Type": "application/json"})

        if res["status"] == "matched":
            # 매칭된 고객_DB 행을 그대로 전달 (이름으로 다시 찾지 않음 → 사업자명 겹쳐도 정확)
            cust_info = res["cust"]
            matched   = _cust_name(cust_info)
            write_advertiser(gc, int(row), matched)
            result = update_work_sheet(gc, matched, payer, amount_k, amount_h, cust_info)
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
