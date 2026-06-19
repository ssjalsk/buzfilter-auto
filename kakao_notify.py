import os
import re
import json
import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import pytz

KST = pytz.timezone("Asia/Seoul")
SHEET_ID = "1OJkg679B09qvW5hAY_vT35KD0dl5435peGszwv55Fzs"
SHEET_TAB = "업무시트"
TOKEN_EXPIRE_DATE = datetime(2026, 8, 15)  # 리프레시 토큰 만료일 (발급 후 60일)


def refresh_access_token(rest_api_key, refresh_token):
    resp = requests.post(
        "https://kauth.kakao.com/oauth/token",
        data={
            "grant_type": "refresh_token",
            "client_id": rest_api_key,
            "refresh_token": refresh_token,
        },
    )
    data = resp.json()
    if "access_token" not in data:
        raise RuntimeError(f"토큰 갱신 실패: {data}")
    return data["access_token"]


def send_kakao_message(access_token, text):
    resp = requests.post(
        "https://kapi.kakao.com/v2/api/talk/memo/default/send",
        headers={"Authorization": f"Bearer {access_token}"},
        data={
            "template_object": json.dumps({
                "object_type": "text",
                "text": text,
                "link": {"web_url": "", "mobile_web_url": ""},
            })
        },
    )
    return resp.json()


def parse_date(value):
    s = str(value).strip()
    m = re.search(r"(\d{4})[.\-/\s]+(\d{1,2})[.\-/\s]+(\d{1,2})", s)
    if not m:
        return None
    y, mo, d = (int(x) for x in m.groups())
    try:
        return datetime(y, mo, d).date()
    except Exception:
        return None


def parse_time(value):
    s = str(value).strip()
    for fmt in ("%H:%M:%S", "%H:%M", "%I:%M %p", "%I:%M%p"):
        try:
            return datetime.strptime(s, fmt).time()
        except Exception:
            continue
    return None


def main():
    now = datetime.now(KST)
    print(f"현재 시각(KST): {now.strftime('%Y-%m-%d %H:%M')}")

    creds_json = json.loads(os.environ["GOOGLE_CREDENTIALS"])
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_json, scope)
    client = gspread.authorize(creds)

    sheet = client.open_by_key(SHEET_ID).worksheet(SHEET_TAB)
    rows = sheet.get_all_values()

    matched = []
    for i, row in enumerate(rows[1:], start=2):
        # 짧은 행을 8칸으로 패딩 (get_all_values는 빈 열을 생략할 수 있음)
        row = list(row) + [""] * max(0, 8 - len(row))

        topic = row[0].strip()    # A열
        date_val = row[6].strip()  # G열
        time_val = row[7].strip()  # H열

        if not topic:
            continue

        if not date_val or not time_val:
            print(f"  스킵: 행{i} [{topic}] 날짜/시간 비어있음 (date='{date_val}', time='{time_val}')")
            continue

        row_date = parse_date(date_val)
        row_time = parse_time(time_val)

        if row_date is None:
            print(f"  스킵: 행{i} [{topic}] 날짜 파싱 실패 (date='{date_val}')")
            continue
        if row_time is None:
            print(f"  스킵: 행{i} [{topic}] 시간 파싱 실패 (time='{time_val}')")
            continue

        scheduled = datetime.combine(row_date, row_time)
        diff_sec = (now.replace(tzinfo=None) - scheduled).total_seconds()
        print(f"  검사: 행{i} [{topic}] {date_val} {time_val} → diff={int(diff_sec)}s")

        if 0 <= diff_sec < 600:  # 예약시간 기준 0~10분 이내 (10분 cron 주기에 맞춤)
            matched.append(topic)
            print(f"  매칭: 행{i} [{topic}] {date_val} {time_val} (diff={int(diff_sec)}s)")

    # 토큰 만료 D-10 이내 경고 (매일 09:00에 한 번만)
    days_left = (TOKEN_EXPIRE_DATE - now.replace(tzinfo=None)).days
    if days_left <= 10 and now.hour == 9 and now.minute < 10:
        access_token = refresh_access_token(
            os.environ["KAKAO_REST_API_KEY"],
            os.environ["KAKAO_REFRESH_TOKEN"],
        )
        send_kakao_message(access_token, f"⚠️ 카카오 알림봇 토큰이 {days_left}일 후 만료됩니다. 알리고에게 갱신 요청하세요!")
        print(f"토큰 만료 경고 전송: D-{days_left}")

    if not matched:
        print("알림 없음")
        return

    access_token = refresh_access_token(
        os.environ["KAKAO_REST_API_KEY"],
        os.environ["KAKAO_REFRESH_TOKEN"],
    )

    for topic in matched:
        msg = f"{topic} 노출 체크 해야됩니다."
        result = send_kakao_message(access_token, msg)
        print(f"전송 결과: {result}")


if __name__ == "__main__":
    main()
