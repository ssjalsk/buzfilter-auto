import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import anthropic
from datetime import datetime
import os
import json
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# =====================================================
# 공통 설정
# =====================================================
def get_anthropic_client():
    try:
        return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
    except:
        api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        return anthropic.Anthropic(api_key=api_key)

def get_sheet(worksheet_name):
    SHEET_URL = "https://docs.google.com/spreadsheets/d/1CtD6VVtmiQNz90mKJFfuPq8-LMowLHg3NZPnoqwpISE/"
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        try:
            creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        except:
            BASE_DIR = os.path.dirname(os.path.abspath(__file__))
            JSON_FILE = os.path.join(BASE_DIR, 'service_account.json')
            creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, scope)
        client_gs = gspread.authorize(creds)
        return client_gs.open_by_url(SHEET_URL).worksheet(worksheet_name)
    except Exception as e:
        st.error(f"❌ 시트 연결 실패 ({worksheet_name}): {e}")
        return None

def find_last_data_row(sheet):
    all_values = sheet.get_all_values()
    last_row = 2
    for i, row in enumerate(all_values):
        if len(row) > 1 and row[1].strip() != '':
            last_row = i + 1
    return last_row + 1

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

# =====================================================
# 페이지 설정
# =====================================================
st.set_page_config(page_title="버즈필터 자동화", page_icon="🤖", layout="wide")

# =====================================================
# 사이드바 메뉴
# =====================================================
with st.sidebar:
    st.markdown("## 📋 메뉴")
    st.markdown("---")
    menu = st.radio(
        "",
        options=["🏭 버즈필터 발주", "📝 리뷰 입력"],
        label_visibility="collapsed"
    )
    st.markdown("---")
    st.caption("버즈필터 업무 자동화 시스템")
    st.caption("© 2025 알리고미디어")

# =====================================================
# 페이지 1: 버즈필터 발주
# =====================================================
if menu == "🏭 버즈필터 발주":
    st.title("🤖 버즈필터 발주 자동 장부 입력")
    st.subheader("📊 발주서 엑셀을 업로드하면 장부에 자동으로 입력합니다.")

    uploaded_file = st.file_uploader("발주서 엑셀 파일 선택 (.xlsx)", type=['xlsx'])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        df.columns = [col.split('(')[0].strip() for col in df.columns]

        st.write("📂 업로드 데이터 미리보기:", df.head())
        st.write(f"총 {len(df)}건 발주 데이터 확인")

        if st.button("🚀 장부 자동입력 시작"):
            client = get_anthropic_client()

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

# =====================================================
# 페이지 2: 리뷰 입력
# =====================================================
elif menu == "📝 리뷰 입력":
    st.title("📝 리뷰 엑셀 자동 변환기")
    st.subheader("리뷰 텍스트 파일을 업로드하면 엑셀 파일로 자동 변환합니다.")

    def parse_reviews(text):
        # 구분 형식 4가지: 1 / 1. / 1) / (1)
        # 숫자만 있는 줄, 또는 숫자+마침표/괄호 형태의 줄을 기준으로 분리
        delimiter = re.compile(
            r'^\s*(?:\((\d+)\)|(\d+)[.\)]|(\d+))\s*$',
            re.MULTILINE
        )

        # 구분자 위치와 번호를 먼저 다 찾기
        markers = []
        for m in delimiter.finditer(text):
            num = int(m.group(1) or m.group(2) or m.group(3))
            markers.append((num, m.start(), m.end()))

        if not markers:
            return []

        reviews = []
        for i, (num, start, end) in enumerate(markers):
            # 이 리뷰의 내용: 구분자 끝 ~ 다음 구분자 시작
            if i + 1 < len(markers):
                content_raw = text[end:markers[i + 1][1]]
            else:
                content_raw = text[end:]

            content = content_raw.strip()
            if content:
                reviews.append((num, content))

        return sorted(reviews, key=lambda x: x[0])

    def create_excel(reviews):
        wb = Workbook()
        ws = wb.active
        ws.title = "리뷰"

        header_fill = PatternFill("solid", start_color="FF6B35", end_color="FF6B35")
        header_font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left_wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)
        thin = Side(style="thin", color="CCCCCC")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for cell_addr, label in {"A1": "번호", "B1": "별점", "C1": "리뷰 내용"}.items():
            cell = ws[cell_addr]
            cell.value = label
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center
            cell.border = border

        ws.column_dimensions["A"].width = 8
        ws.column_dimensions["B"].width = 8
        ws.column_dimensions["C"].width = 70
        ws.row_dimensions[1].height = 30

        alt_fill = PatternFill("solid", start_color="FFF5F0", end_color="FFF5F0")
        normal_font = Font(name="Arial", size=10)

        for i, (num, content) in enumerate(reviews, start=2):
            row_fill = alt_fill if i % 2 == 0 else None

            a_cell = ws.cell(row=i, column=1, value=num)
            a_cell.font = Font(name="Arial", size=10, bold=True)
            a_cell.alignment = center
            a_cell.border = border
            if row_fill:
                a_cell.fill = row_fill

            b_cell = ws.cell(row=i, column=2, value="")
            b_cell.border = border
            if row_fill:
                b_cell.fill = row_fill

            c_cell = ws.cell(row=i, column=3, value=content)
            c_cell.font = normal_font
            c_cell.alignment = left_wrap
            c_cell.border = border
            if row_fill:
                c_cell.fill = row_fill

            lines = content.count('\n') + 1
            ws.row_dimensions[i].height = max(40, min(lines * 18, 200))

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    tab1, tab2 = st.tabs(["📁 파일 업로드", "✏️ 텍스트 직접 입력"])

    with tab1:
        uploaded_txt = st.file_uploader(
            "리뷰 텍스트 파일 업로드 (.txt)",
            type=["txt"],
            help="번호로 구분된 리뷰 텍스트 파일을 업로드해주세요."
        )
        if uploaded_txt:
            text_to_process = uploaded_txt.read().decode("utf-8", errors="ignore")
            st.success(f"✅ 파일 업로드 완료: {uploaded_txt.name}")

            reviews = parse_reviews(text_to_process)
            if reviews:
                st.markdown(f"### 📊 **{len(reviews)}개** 리뷰 감지됨")
                with st.expander("👀 변환 미리보기", expanded=True):
                    for num, content in reviews[:5]:
                        st.markdown(f"**{num}번 리뷰**")
                        st.text(content[:200] + ("..." if len(content) > 200 else ""))
                        st.divider()
                    if len(reviews) > 5:
                        st.info(f"... 외 {len(reviews) - 5}개 리뷰")

                excel_data = create_excel(reviews)
                st.download_button(
                    label="⬇️ 엑셀 파일 다운로드",
                    data=excel_data,
                    file_name="리뷰목록.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
            else:
                st.error("❌ 리뷰를 파싱할 수 없습니다. 텍스트 형식을 확인해주세요.")

    with tab2:
        manual_text = st.text_area(
            "리뷰 내용을 직접 붙여넣기",
            height=300,
            placeholder='1.\n"리뷰 내용..."\n\n2.\n"리뷰 내용..."'
        )
        if manual_text.strip():
            reviews = parse_reviews(manual_text)
            if reviews:
                st.markdown(f"### 📊 **{len(reviews)}개** 리뷰 감지됨")
                with st.expander("👀 변환 미리보기", expanded=True):
                    for num, content in reviews[:5]:
                        st.markdown(f"**{num}번 리뷰**")
                        st.text(content[:200] + ("..." if len(content) > 200 else ""))
                        st.divider()
                    if len(reviews) > 5:
                        st.info(f"... 외 {len(reviews) - 5}개 리뷰")

                excel_data = create_excel(reviews)
                st.download_button(
                    label="⬇️ 엑셀 파일 다운로드",
                    data=excel_data,
                    file_name="리뷰목록.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
            else:
                st.error("❌ 리뷰를 파싱할 수 없습니다.")

    with st.expander("📖 사용법"):
        st.markdown("""
        **텍스트 파일 형식:**
        - 각 리뷰는 `1.` `2.` 또는 `1)` `2)` 형식으로 번호를 붙여주세요
        - 번호 다음에 줄바꿈 후 리뷰 내용 작성
        - 리뷰 사이는 빈 줄로 구분

        **엑셀 출력:**
        - **A열**: 리뷰 번호
        - **B열**: 별점 (비워둠)
        - **C열**: 리뷰 내용
        """)
