import os, json, re, time, functions_framework
from google.oauth2.service_account import Credentials
import gspread

SPREADSHEET_ID        = "1mOV-HmlODZaxPohiVFay9_Rnh31-vhcPXF_b-5-EjB0"  # 테스트 시트 — 고객_DB/업무시트/종합 정산시트 전부
SETTLE_SPREADSHEET_ID = "1mOV-HmlODZaxPohiVFay9_Rnh31-vhcPXF_b-5-EjB0"  # 동일
DB_SHEET       = "고객_DB"
WORK_SHEET     = "업무시트"
SETTLE_SHEET   = "종합 정산시트"
REVIEW_SHEET   = "리뷰 업무시트"
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
    # [2026-10-15] 숫자·쉼표·₩·원만 있을 때만 금액. 담당자 메모('6/8 10건 접수…' → 681,020원으로 읽히던 문제)는 0
    if not re.fullmatch(r"[₩￦\s\d,.원\-]+", s):
        return 0
    return parse_amount(s)


def _date_key(val):
    """의뢰일 '2026. 9. 7' / '2026-09-07' → (2026, 9, 7). 없거나 못 읽으면 맨 뒤로"""
    m = re.match(r"\s*(\d{4})\s*[.\-/]\s*(\d{1,2})\s*[.\-/]\s*(\d{1,2})", str(val or ""))
    return (int(m[1]), int(m[2]), int(m[3])) if m else (9999, 12, 31)


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
            s_val = _parse_s_val(row[18] if len(row) > 18 else "")
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
            w_val = _parse_s_val(row[22] if len(row) > 22 else "")   # W열: 부분입금 누적
            if q_val - w_val > 0:
                amounts.append(q_val - w_val)
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
RECORD_COL  = 29   # 종합 정산시트 AC열: 자동매칭이 실제로 처리한 금액(부가세 제외) 기록 — 금액 변경 시 차액 계산 기준


def net_amount(amount_k, amount_h):
    """입금액(부가세 제외): H열(미발행) 우선, 없으면 K열/1.1"""
    if amount_h > 0:
        return amount_h
    if amount_k > 0:
        return round(amount_k / 1.1)
    return 0


def write_record(gc, row_idx, amount_n):
    ws = gc.open_by_key(SETTLE_SPREADSHEET_ID).worksheet(SETTLE_SHEET)
    ws.update_cell(row_idx, RECORD_COL, amount_n)


def mark_amount_cell(gc, row_idx, note, warn):
    """K열(입금액)에 메모 + 경고 시 연한 빨간 배경 (처리 결과를 시트에서 바로 확인)"""
    ws = gc.open_by_key(SETTLE_SPREADSHEET_ID).worksheet(SETTLE_SHEET)
    ws.update_note(f"K{row_idx}", note)
    color = {"red": 1.0, "green": 0.85, "blue": 0.85} if warn else {"red": 1.0, "green": 1.0, "blue": 1.0}
    ws.format(f"K{row_idx}", {"backgroundColor": color})


def process_amount_change(gc, row_idx, payer, amount_k, amount_h):
    """
    [2026-10-06] 이미 자동매칭된 행의 금액(H/K)이 수정됐을 때.
      - AC열 기록 없음 → 아무것도 안 함 (기능 도입 전 입력 / 광고주 수동선택 행 — 기존 동작 유지)
      - 늘어남 → 늘어난 차액만 같은 고객 대상으로 업무시트 재매칭 (완료/승계/선충전)
      - 줄어듦 → 자동 되돌리기 안 함, K열 메모·빨간 표시로 수동확인 요청
    """
    ws  = gc.open_by_key(SETTLE_SPREADSHEET_ID).worksheet(SETTLE_SHEET)
    adv = str(ws.cell(row_idx, 6).value or "").strip()
    rec = str(ws.cell(row_idx, RECORD_COL).value or "").strip()
    if not adv or adv.startswith("⚠️") or not rec:
        return {"action": "skip"}

    recorded = parse_amount(rec)
    new_n    = net_amount(amount_k, amount_h)
    diff     = new_n - recorded
    if diff == 0:
        return {"action": "same"}

    if diff < 0:
        mark_amount_cell(gc, row_idx,
                         f"⚠️ 금액 감소 — 자동 처리하지 않았습니다.\n"
                         f"처리된 금액(부가세 제외): {recorded:,}원 → 바뀐 금액: {new_n:,}원 (차액 {diff:,}원)\n"
                         f"업무시트 입금완료 건/선충전잔액을 확인해 수동으로 조정해 주세요.", warn=True)
        return {"action": "decrease", "diff": diff}

    customers = load_customer_db(gc)
    cust_info = next((c for c in customers
                      if c.get("담당자명","") == adv or c.get("사업자명","") == adv), None)
    if not cust_info:
        mark_amount_cell(gc, row_idx,
                         f"⚠️ 금액 증가 {diff:,}원 — 광고주 '{adv}' 를 고객_DB에서 찾지 못해 처리하지 않았습니다.", warn=True)
        return {"action": "no_customer", "diff": diff}

    # 기록을 먼저 갱신 → 처리 중 오류가 나도 차액이 두 번 들어가지 않음 (오류 시 메모로 수동확인 안내)
    write_record(gc, row_idx, new_n)
    try:
        result = update_work_sheet(gc, adv, payer, 0, diff, cust_info)   # amount_h 자리 = 부가세 제외 차액
    except Exception as e:
        mark_amount_cell(gc, row_idx,
                         f"⚠️ 금액 증가 {diff:,}원 추가 처리 중 오류 — 수동확인 필요\n({e})", warn=True)
        raise
    if result.get("price_pending"):
        # [2026-10-15] 단가표 로딩 실패 → 아무것도 처리 안 됨. 기록을 되돌려 K열 재입력 시 다시 처리되게 함
        write_record(gc, row_idx, recorded)
        mark_amount_cell(gc, row_idx,
                         f"⚠️ 금액 증가 {diff:,}원 — 단가표(업무시트 N열 금액)를 불러오지 못해 처리하지 않았습니다.\n"
                         f"업무시트 {result['price_pending']}행 금액이 보이는지 확인 후 K열 금액을 다시 입력해 주세요.", warn=True)
        return {"action": "price_pending", "diff": diff, "rows": result["price_pending"]}
    mark_amount_cell(gc, row_idx,
                     f"✅ 금액 증가 반영: {recorded:,}원 → {new_n:,}원 (부가세 제외)\n"
                     f"차액 {diff:,}원 추가 처리 — 완료 {result['processed']}건, 선충전 {result['prepaid']:,}원",
                     warn=False)
    return {"action": "increase", "diff": diff, **result}


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


def write_price_pending(gc, row_idx, advertiser, work_rows):
    """
    [2026-10-15] 고객은 찾았지만 업무시트 금액(N열, 외부 단가표 수식)이 안 불러와져 배정 불가.
    F열 수동확인 표시 + 메모. 업무시트 입금처리·선충전 적립은 하지 않음.
    F열을 지우고 K(또는 H)열 금액을 다시 입력하면 자동매칭이 다시 실행됨.
    """
    ws = gc.open_by_key(SETTLE_SPREADSHEET_ID).worksheet(SETTLE_SHEET)
    ws.update_cell(row_idx, 6, MANUAL_MARK)
    ws.update_note(f"F{row_idx}",
                   f"고객: {advertiser}\n"
                   f"업무시트 금액(N열)을 단가표에서 불러오지 못해 입금 처리를 멈췄습니다. (업무시트 {', '.join(map(str, work_rows))}행)\n"
                   "업무시트에 금액이 보이면 → 이 F열을 지우고 K(또는 H)열 금액을 다시 입력해 주세요. 자동으로 다시 처리됩니다.")


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
    if "선충전잔액" in headers:
        pre_col_idx = headers.index("선충전잔액")
    elif len(headers) > 6:
        # [2026-10-16] 제목 칸(G1)이 지워지거나 0으로 덮여도 G열 고정 위치로 계속 기록 (남은 입금액이 사라지지 않게)
        pre_col_idx = 6
    else:
        return 0, -1, -1
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


PRICE_SHEET = "price_mirror"   # 업무시트 N열 단가 수식이 참조하는 탭 (외부 단가표 IMPORTRANGE)


def _price_pending_rows(all_vals, name_set):
    """이 고객의 미입금 행 중 '매체사(I)·소재(K)는 있는데 금액(N)만 빈' 행 번호.
    단가표(IMPORTRANGE) 새로고침 순간에는 N열 수식이 전부 빈칸으로 읽힘 → 그대로 두면 이 행을 건너뛰고
    뒤쪽 행이나 선충전에 잘못 배정됨 (2026-10-08 플랜원컴퍼니 3920행)."""
    rows = []
    for i, row in enumerate(all_vals[1:], start=2):
        if norm(row[1] if len(row) > 1 else "") not in name_set:
            continue
        if (row[17].strip() if len(row) > 17 else ""):          # R열 입금자 있음 = 처리된 행
            continue
        if (row[16].strip() if len(row) > 16 else "") == "입금완료":
            continue
        if (row[4].strip() if len(row) > 4 else "") == "반려":
            continue
        n = row[13].strip() if len(row) > 13 else ""
        i_val = row[8].strip() if len(row) > 8 else ""
        k_val = row[10].strip() if len(row) > 10 else ""
        if not n and i_val and k_val:
            rows.append(i)
    return rows


def _price_sheet_loaded(gc):
    """price_mirror(외부 단가표 불러오기)가 정상 표시 중인지: A4='매체사' + 아래 매체 목록 존재"""
    try:
        vals = gc.open_by_key(SPREADSHEET_ID).worksheet(PRICE_SHEET).get("A4:A8")
        col = [r[0].strip() if r else "" for r in vals]
        return len(col) >= 3 and col[0] == "매체사" and sum(1 for c in col[1:] if c) >= 2
    except Exception:
        return False


def update_work_sheet(gc, advertiser, payer, amount_k=0, amount_h=0, cust_info=None):
    """
    입금 처리 — 업무시트 + 리뷰 업무시트 미입금 건에 의뢰일 오래된 순(FIFO)으로 배정.

    파라미터:
      advertiser : 매칭된 광고주명 (고객_DB 담당자명/사업자명)
      payer      : 입금자명 (R열 / 리뷰 U열에 기록)
      amount_k   : 종합 정산시트 K열 부가세포함 입금액
      amount_h   : 종합 정산시트 H열 미발행 입금액 (부가세 없음, 원금 그대로)
      cust_info  : 고객_DB 행 dict (담당자명/사업자명/별칭 포함) — 별칭 B열 매칭용

    처리 흐름:
      1. 입금액(부가세 제외) = H열 그대로 / K열 ÷ 1.1
      2. 미입금 건 수집
         - 업무시트: B 일치, R(입금자) 빈칸, Q≠입금완료, E≠반려, N>0 / 기존 부분입금 = S열(숫자만)
         - 리뷰 업무시트: B 일치, T≠입금완료, Q(매출)>0 / 기존 부분입금 = W열(숫자만)
      3. 의뢰일 오래된 순으로 배정
         - 남은 입금 >= 필요액 → 완납 (업무: Q=입금완료·R=입금자·S 비움·링크 있으면 E=송출완료 / 리뷰: T=입금완료·U=입금자·W 비움)
         - 남은 입금 <  필요액 → 부분입금 (S 또는 W열에 누적 금액 숫자로 기록, 입금자명은 쓰지 않음) 후 중단
      4. 다 배정하고 남은 금액 → 고객_DB 선충전잔액에 추가

    [2026-10-15 정산 전체 점검]
      - S열 메모(글자)를 금액으로 읽지 않음 (_parse_s_val)
      - 이미 입금완료(Q)인데 R이 빈 행, 반려 행에는 배정하지 않음
      - 리뷰 업무시트 포함 (예전 정산 방식 APPLY_RECON_STACK_ 과 같은 범위)
      - 처리된 행의 S열을 '승계금액'으로 다시 더하던 동작 제거 (메모·부분입금 기록이 돈으로 두 번 계산되던 원인)

    반환: {"processed": 완납 건수, "prepaid": 선충전 추가 금액} / 단가표 로딩 실패 시 {"price_pending": [행]}
    """
    ss = gc.open_by_key(SPREADSHEET_ID)
    ws = ss.worksheet(WORK_SHEET)
    all_vals = ws.get_all_values()

    # B열 매칭: 담당자명, 사업자명, 대표자명, 별칭 모두 포함
    name_set = _build_name_set(advertiser, cust_info)

    # ── [2026-10-15] 단가표 로딩 중 금액 빈칸 대비 ──────────────────
    # 금액만 빈 미입금 행이 있으면 몇 초 뒤 다시 읽음.
    #  - 다시 읽어서 채워지면 → 정상 진행
    #  - 단가표는 정상인데 여전히 빈칸 → 단가표에 없는 매체(원래 빈칸) → 예전처럼 건너뛰고 진행
    #  - 끝까지 단가표가 안 불러와짐 → 아무것도 쓰지 않고 price_pending 반환 (호출부에서 수동확인 표시)
    pending = _price_pending_rows(all_vals, name_set)
    for attempt in range(3):
        if not pending:
            break
        if attempt >= 1 and _price_sheet_loaded(gc):
            break
        time.sleep(3)
        all_vals = ws.get_all_values()
        pending  = _price_pending_rows(all_vals, name_set)
    if pending and not _price_sheet_loaded(gc):
        return {"processed": 0, "prepaid": 0, "price_pending": pending}

    # ── 입금 금액 (부가세 제외) ──────────────────────────────────────
    remaining_n = net_amount(amount_k, amount_h)
    if remaining_n <= 0:
        return {"processed": 0, "prepaid": 0}

    cell = lambda row, i: (str(row[i]).strip() if len(row) > i else "")

    # ── 미입금 건 수집 ───────────────────────────────────────────────
    # 업무시트 (0-based): B=1, E=4, F=5(의뢰일), M=12, N=13, Q=16, R=17, S=18
    candidates = []
    for i, row in enumerate(all_vals[1:], start=2):
        if norm(cell(row, 1)) not in name_set:
            continue
        if cell(row, 17):                    # R열 입금자 있음 = 처리된 행
            continue
        if cell(row, 16) == "입금완료":       # Q열 입금완료 (R만 비어 있는 수동 처리 행 포함)
            continue
        if cell(row, 4) == "반려":
            continue
        n_val = parse_amount(cell(row, 13))
        if n_val <= 0:
            continue                          # 금액 없는 행 제외
        candidates.append({"sheet": "work", "row": i, "need_total": n_val,
                           "paid": _parse_s_val(cell(row, 18)), "m_val": cell(row, 12),
                           "date": _date_key(cell(row, 5))})

    # 리뷰 업무시트 (0-based): B=1, N=13(의뢰일), Q=16(매출), T=19(입금여부), U=20(입금자명), W=22(승계금액)
    try:
        rws = ss.worksheet(REVIEW_SHEET)
        r_vals = rws.get_all_values()
    except Exception:
        rws, r_vals = None, []
    for i, row in enumerate(r_vals[1:], start=2):
        if norm(cell(row, 1)) not in name_set:
            continue
        if cell(row, 19) == "입금완료":
            continue
        q_val = parse_amount(cell(row, 16))
        if q_val <= 0:
            continue
        candidates.append({"sheet": "review", "row": i, "need_total": q_val,
                           "paid": _parse_s_val(cell(row, 22)), "m_val": "",
                           "date": _date_key(cell(row, 13))})

    # 의뢰일 오래된 순 (같은 날이면 업무시트 → 리뷰, 위 행부터)
    candidates.sort(key=lambda c: (c["date"], 0 if c["sheet"] == "work" else 1, c["row"]))

    # ── 순차 배정 ────────────────────────────────────────────────────
    w_up, r_up = [], []
    processed  = 0
    for c in candidates:
        if remaining_n <= 0:
            break
        needed = c["need_total"] - c["paid"]
        if c["sheet"] == "work":
            Qc, Rc, Pc, up = "Q", "R", "S", w_up
        else:
            Qc, Rc, Pc, up = "T", "U", "W", r_up

        if needed <= 0 or remaining_n >= needed:
            # 완납 (needed<=0: 부분입금 누적이 이미 금액 이상인 예외 → 돈 쓰지 않고 마감)
            up.append({"range": f"{Rc}{c['row']}", "values": [[payer]]})
            up.append({"range": f"{Qc}{c['row']}", "values": [["입금완료"]]})
            up.append({"range": f"{Pc}{c['row']}", "values": [[""]]})
            if c["m_val"]:
                up.append({"range": f"E{c['row']}", "values": [["송출완료"]]})
            remaining_n -= max(needed, 0)
            processed += 1
        else:
            # 부분입금 — 누적 금액만 숫자로 기록 (입금자명은 완납 때만)
            up.append({"range": f"{Pc}{c['row']}", "values": [[c["paid"] + remaining_n]]})
            remaining_n = 0
            break

    if w_up:
        ws.batch_update(w_up)
    if r_up and rws:
        rws.batch_update(r_up)

    # ── 다 배정하고 남은 금액(초과입금·미입금 건 없음) → 선충전잔액에 추가 ──
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

        # [2026-10-06] 이미 매칭된 행의 금액 수정 → 차액 처리 모드
        if data.get("mode") == "delta":
            res = process_amount_change(gc, int(row), payer, amount_k, amount_h)
            return (json.dumps({"ok": True, "mode": "delta", **res}), 200,
                    {"Content-Type": "application/json"})

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

        # [2026-10-16] 같은 정산 행 이중 처리 방지: 이미 광고주(F)가 채워졌거나 처리기록(AC)이 있으면 다시 처리하지 않음
        #  (어떤 경로로든 같은 입금 행에 매칭 요청이 두 번 오면, 두 번째는 미입금 건이 없어 입금액 전액이 선충전으로 쌓였음)
        try:
            settle_ws = gc.open_by_key(SETTLE_SPREADSHEET_ID).worksheet(SETTLE_SHEET)
            f_now  = str(settle_ws.cell(int(row), 6).value or "").strip()
            ac_now = str(settle_ws.cell(int(row), RECORD_COL).value or "").strip()
        except Exception:
            f_now = ac_now = ""
        if f_now or ac_now:
            return (json.dumps({"ok": True, "matched": None, "skipped": True,
                                "msg": f"행 {row}: 이미 처리된 입금 행 (F='{f_now}', AC='{ac_now}') — 중복 요청 무시"}), 200,
                    {"Content-Type": "application/json"})

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
            if result.get("price_pending"):
                # [2026-10-15] 단가표가 안 불러와져 금액을 모름 → 업무시트·선충전 아무것도 안 건드림
                write_price_pending(gc, int(row), matched, result["price_pending"])
                return (json.dumps({"ok": True, "matched": None, "manual": True,
                                    "msg": f"'{payer}' → {matched}: 단가표 로딩 실패로 수동확인 (업무시트 {result['price_pending']}행)"}),
                        200, {"Content-Type": "application/json"})
            write_record(gc, int(row), net_amount(amount_k, amount_h))   # 처리 금액 기록 (금액 변경 시 차액 기준)
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
