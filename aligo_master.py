import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import anthropic
from datetime import datetime
import os
import json

# --- 1. 설정 ---
ANTHROPIC_API_KEY = st.secrets["ANTHROPIC_API_KEY"]
SHEET_URL = "https://docs.google.com/spreadsheets/d/1CtD6VVtmiQNz90mKJFfuPq8-LMowLHg3NZPnoqwpISE/"

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

# --- 2. 구글 시트 연결 ---
def get_sheet(worksheet_name):
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
        from oauth2client.service_account import ServiceAccountCredentials
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client_gs = gspread.authorize(creds)
        return client_gs.open_by_url(SHEET_URL).worksheet(worksheet_name)
    except Exception as e:
        st.error(f"❌ 시트 연결 실패 ({worksheet_name}): {e}")
        return None

# --- 3. 마지막 데이터 행 찾기 ---
def find_last_data_row(sheet):
    all_values = sheet.get_all_values()
    last_row = 2
    for i, row in enumerate(all_values):
        if len(row) > 1 and row[1].strip() != '':
            last_row = i + 1
    return last_row + 1

# --- 4. 열별 개별 입력 (수식 열 건드리지 않음) ---
def insert_row_safe(sheet, start_row, rows_data):
    for i, row in enumerate(rows_data):
        current_row = start_row + i
        sheet.update(f"B{current_row}", [[row[0]]], value_input_option='USER_ENTERED')
        sheet.update(f"C{current_row}", [[row[1]]], value_input_option='USER_ENTERED')
        sheet.update(f"D{current_row}", [[row[2]]], value_input_option='USER_ENTERED')
        sheet.update(f"E{current_row}", [[row[3]]], value_input_option='USER_ENTERED')
        sheet.update(f"F{current_row}", [[row[4]]], value_input_option='USER_ENTERED')
        sheet.update(f"H{current_row}", [[row[5]]], value_input_option='USER_ENTERED')
        sheet.update(f"I{current_row}", [[row[6]]], value_input_option='USER_ENTERED')
        sheet.update(f"K{current_row}", [[row[7]]], value_input_option='USER_ENTERED')

# --- 5. UI ---
st.set_page_config(page_title="버즈필터 발주 자동입력", layout="wide")
st.title("🤖 버즈필터 발주 자동 장부 입력")
st.subheader("📊 발주서 엑셀을 업로드하면 장부에 자동으로 입력합니다.")

uploaded_file = st.file_uploader("발주서 엑셀 파일 선택 (.xlsx)", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = [col.split('(')[0].strip() for col in df.columns]

    st.write("📂 업로드 데이터 미리보기:", df.head())
    st.write(f"총 {len(df)}건 발주 데이터 확인")

    if st.button("🚀 장부 자동입력 시작"):

        with st.spinner("마진계산기 불러오는 중..."):
            margin_sheet = get_sheet("2. 버즈필터 마진 계산기")
            if margin_sheet is None:
                st.stop()

            all_margin_data = margin_sheet.get_all_values()
            margin_headers = all_margin_data[1]
            margin_rows = all_margin_data[2:]
            calc_df = pd.DataFrame(margin_rows, columns=margin_headers)
            calc_df = calc_df[calc_df['제품명'].str.strip() != '']

            st.success(f"✅ 마진계산기 로드 완료 ({len(calc_df)}개 상품)")

        with st.spinner("AI가 상품 매칭 중..."):
            today = datetime.now()
            rows_to_add = []
            match_results = []

            for idx, row in df.iterrows():
                raw_name = str(row.get('상품명+옵션+개수', ''))
                quantity = int(row.get('수량', 1))
                price = int(row.get('가격', 0))
                channel = str(row.get('판매처', '쿠팡'))

                prompt = f"""
너는 상품 매칭 전문가야.
아래 발주서 상품명과 가장 일치하는 제품의 정보를 찾아줘.

발주서 상품명: {raw_name}

상품 리스트:
{calc_df[['브랜드', '제품명', '상품코드 표']].to_string(index=False)}

반드시 아래 형식으로만 답해줘. 다른 말은 절대 하지마:
상품코드: [코드값]
브랜드: [브랜드값]

못 찾겠으면:
상품코드: 미등록
브랜드: 미등록
"""
                response = client.messages.create(
                    model="claude-haiku-4-5-20251001",
                    max_tokens=100,
                    messages=[{"role": "user", "content": prompt}]
                )

                result_text = response.content[0].text.strip()
                matched_code = "미등록"
                matched_brand = "미등록"

                for line in result_text.split('\n'):
                    if '상품코드:' in line:
                        matched_code = line.split('상품코드:')[1].strip()
                    if '브랜드:' in line:
                        matched_brand = line.split('브랜드:')[1].strip()

                match_results.append({
                    '상품명': raw_name,
                    '매칭 브랜드': matched_brand,
                    '매칭 코드': matched_code,
                    '판매처': channel,
                    '가격': price,
                    '수량': quantity
                })

                new_row = [
                    str(today.year),
                    str(today.month),
                    str(today.day),
                    matched_brand,
                    matched_code,
                    channel,
                    price,
                    quantity,
                ]
                rows_to_add.append(new_row)

            st.session_state['rows_to_add'] = rows_to_add
            st.session_state['match_results'] = match_results
            st.session_state['ready_to_insert'] = True

        st.success("✅ AI 매칭 완료! 아래 결과를 확인해주세요.")

    if st.session_state.get('ready_to_insert'):
        result_df = pd.DataFrame(st.session_state['match_results'])
        st.write("🔍 AI 매칭 결과 확인:")
        st.dataframe(result_df)

        unmatched = result_df[result_df['매칭 코드'] == '미등록']
        if len(unmatched) > 0:
            st.warning(f"⚠️ {len(unmatched)}건 매칭 실패 - 수동 확인 필요")

        if st.button("✅ 확인했습니다. 장부에 최종 입력합니다."):
            with st.spinner("장부 입력 중..."):
                try:
                    ledger_sheet = get_sheet("2. 버즈필터 장부")
                    if ledger_sheet is None:
                        st.stop()

                    start_row = find_last_data_row(ledger_sheet)
                    st.info(f"📍 {start_row}행부터 입력 시작")

                    insert_row_safe(
                        ledger_sheet,
                        start_row,
                        st.session_state['rows_to_add']
                    )

                    st.success(f"🎉 총 {len(st.session_state['rows_to_add'])}건 장부 입력 완료!")
                    st.balloons()

                    st.session_state['ready_to_insert'] = False
                    st.session_state['rows_to_add'] = []
                    st.session_state['match_results'] = []

                except Exception as e:
                    st.error(f"❌ 입력 실패: {e}")
