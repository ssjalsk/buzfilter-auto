import os, json, re, functions_framework
from google.oauth2.service_account import Credentials
import gspread

SPREADSHEET_ID        = "1OJkg679B09qvW5hAY_vT35KD0dl5435peGszwv55Fzs"  # 알리고미디어 (고객_DB, 업무시트)
SETTLE_SPREADSHEET_ID = "1mOV-HmlODZaxPohiVFay9_Rnh31-vhcPXF_b-5-EjB0"  # 알리고미디어의 테스트 시트 (종합 정산시트)
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


def update_work_sheet(gc, advertiser, payer):
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet(WORK_SHEET)
    all_vals = ws.get_all_values()
    if len(all_vals) < 2:
        return
    adv_n   = norm(advertiser)
    updates = []
    for i, row in enumerate(all_vals[1:], start=2):
        b = norm(row[1] if len(row) > 1 else "")
        if b != adv_n:
            continue
        e     = row[4].strip() if len(row) > 4 else ""
        r_val = row[17].strip() if len(row) > 17 else ""
        if e == "입금확인":
            updates.append({"range": f"E{i}", "values": [["송출완료"]]})
        if not r_val:
            updates.append({"range": f"R{i}", "values": [[payer]]})
    if updates:
        ws.batch_update(updates)


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

    row   = data.get("row")
    payer = str(data.get("payer", "")).strip()
    if not row or not payer:
        return (json.dumps({"ok": False, "error": "row/payer 필수"}), 400,
                {"Content-Type": "application/json"})

    try:
        gc        = get_gc()
        customers = load_customer_db(gc)
        matched   = match_customer(payer, customers)

        if matched:
            write_advertiser(gc, int(row), matched)
            update_work_sheet(gc, matched, payer)
            return (json.dumps({"ok": True, "matched": matched}), 200,
                    {"Content-Type": "application/json"})
        else:
            return (json.dumps({"ok": True, "matched": None,
                                "msg": f"'{payer}' 매칭 고객 없음 (DB: {len(customers)}명)"}), 200,
                    {"Content-Type": "application/json"})

    except Exception as e:
        return (json.dumps({"ok": False, "error": str(e)}), 500,
                {"Content-Type": "application/json"})
