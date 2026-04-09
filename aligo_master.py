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

def get_anthropic_client():
    try:
        return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
    except:
        return anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))

def get_sheet(worksheet_name):
    SHEET_URL = "https://docs.google.com/spreadsheets/d/1CtD6VVtmiQNz90mKJFfuPq8-LMowLHg3NZPnoqwpISE/"
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        try:
            creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        except:
            BASE_DIR = os.path.dirname(os.path.abspath(__file__))
            creds = ServiceAccountCredentials.from_json_keyfile_name(os.path.join(BASE_DIR, 'service_account.json'), scope)
        client_gs = gspread.authorize(creds)
        return client_gs.open_by_url(SHEET_URL).worksheet(worksheet_name)
    except Exception as e:
        st.error(f"시트 연결 실패 ({worksheet_name}): {e}")
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
        r = start_row + i
        sheet.update(f"B{r}", [[row[0]]], value_input_option='USER_ENTERED')
        sheet.update(f"C{r}", [[row[1]]], value_input_option='USER_ENTERED')
        sheet.update(f"D{r}", [[row[2]]], value_input_option='USER_ENTERED')
        sheet.update(f"E{r}", [[row[3]]], value_input_option='USER_ENTERED')
        sheet.update(f"F{r}", [[row[4]]], value_input_option='USER_ENTERED')
        sheet.update(f"H{r}", [[row[5]]], value_input_option='USER_ENTERED')
        sheet.update(f"I{r}", [[row[6]]], value_input_option='USER_ENTERED')
        sheet.update(f"K{r}", [[row[7]]], value_input_option='USER_ENTERED')

def generate_quote_pdf(quote_data, stamp_path=None):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    fr = os.path.join(BASE_DIR, 'NotoSansKR-Regular.ttf')
    fb_path = os.path.join(BASE_DIR, 'NotoSansKR-Bold.ttf')
    has_kor = os.path.exists(fr) and os.path.exists(fb_path)
    if has_kor:
        pdfmetrics.registerFont(TTFont('KR', fr))
        pdfmetrics.registerFont(TTFont('KR-B', fb_path))
        fn, fb = 'KR', 'KR-B'
    else:
        fn, fb = 'Helvetica', 'Helvetica-Bold'

    buf = io.BytesIO()
    w, h = A4
    c = canvas.Canvas(buf, pagesize=A4)

    LG = colors.HexColor("#F2F2F2")
    MG = colors.HexColor("#CCCCCC")
    DG = colors.HexColor("#404040")
    TBG = colors.HexColor("#D9D9D9")
    ML, MR = 20*mm, w-20*mm
    PW = MR - ML

    items = quote_data["items"]
    is_tax = quote_data["tax_type"] == "발행"
    sup = sum(int(it["수량"])*int(it["단가"]) for it in items)
    vat = int(sup*0.1) if is_tax else 0
    tot = sup + vat

    y = h - 18*mm
    c.setFont(fb, 28)
    c.setFillColor(colors.black)
    c.drawCentredString(w/2, y, "견   적   서")
    y -= 10*mm
    c.setStrokeColor(colors.black)
    c.setLineWidth(1.5)
    c.line(ML, y, MR, y)
    y -= 8*mm

    bt = y
    rcx, rcw = ML+PW*0.5, PW*0.5
    c.setFont(fb, 18); c.setFillColor(colors.HexColor("#1a5fa8"))
    c.drawString(ML+5*mm, bt-12*mm, "Aligo")
    c.setFont(fb, 14)
    c.drawString(ML+5*mm, bt-20*mm, "Media")

    if stamp_path and os.path.exists(stamp_path):
        try:
            c.drawImage(stamp_path, rcx-24*mm, bt-33*mm, width=22*mm, height=22*mm, mask='auto')
        except: pass

    c.setFillColor(DG)
    c.rect(rcx, bt-6*mm, rcw, 6*mm, fill=1, stroke=0)
    c.setFillColor(colors.white); c.setFont(fb, 10)
    c.drawCentredString(rcx+rcw/2, bt-4.5*mm, "공  급  자")

    srows = [("등록번호","161-22-02310","대표자","박철규"),("상  호","알리고미디어","",""),
             ("주  소","서울 마포구 양화로64, 8층","",""),("연락처","010-9469-2381","",""),
             ("업  태","전문, 서비스업","종 목","광고대행업")]
    rh = 5.5*mm
    for i,(k1,v1,k2,v2) in enumerate(srows):
        ry = bt-6*mm-(i+1)*rh
        c.setFillColor(LG if i%2==0 else colors.white)
        c.rect(rcx, ry, rcw, rh, fill=1, stroke=0)
        c.setStrokeColor(MG); c.setLineWidth(0.5)
        c.rect(rcx, ry, rcw, rh, fill=0, stroke=1)
        c.setFillColor(colors.black)
        c.setFont(fb,8); c.drawString(rcx+2*mm, ry+1.5*mm, k1)
        c.setFont(fn,8); c.drawString(rcx+16*mm, ry+1.5*mm, v1)
        if k2:
            c.setFont(fb,8); c.drawString(rcx+rcw*0.62, ry+1.5*mm, k2)
            c.setFont(fn,8); c.drawString(rcx+rcw*0.62+10*mm, ry+1.5*mm, v2)

    y = bt-6*mm-len(srows)*rh-5*mm
    c.setStrokeColor(MG); c.setLineWidth(0.5)
    c.setFillColor(LG); c.rect(ML, y-6*mm, 28*mm, 6*mm, fill=1, stroke=1)
    c.setFillColor(colors.black); c.setFont(fb,9)
    c.drawCentredString(ML+14*mm, y-4.5*mm, "견  적  일")
    c.setFillColor(colors.white); c.rect(ML+28*mm, y-6*mm, PW/2-28*mm, 6*mm, fill=1, stroke=1)
    c.setFillColor(colors.black); c.setFont(fn,9)
    c.drawString(ML+30*mm, y-4.5*mm, quote_data["date"])
    ax = ML+PW/2
    c.setFillColor(colors.white); c.rect(ax, y-6*mm, PW/2, 6*mm, fill=1, stroke=1)
    c.setFillColor(colors.black); c.setFont(fb,8)
    c.drawCentredString(ax+PW/4, y-4.5*mm, "IBK기업은행")
    y -= 6*mm
    c.setFillColor(colors.white); c.rect(ML, y-6*mm, PW/2, 6*mm, fill=1, stroke=1)
    c.setFillColor(colors.black); c.setFont(fb,10)
    c.drawCentredString(ML+PW/4, y-4.5*mm, f"{quote_data['client']}  귀하")
    ax2 = ML+PW/2
    c.setFillColor(colors.white); c.rect(ax2, y-6*mm, PW/2, 6*mm, fill=1, stroke=1)
    c.setFont(fn,7.5); c.setFillColor(colors.black)
    c.drawCentredString(ax2+PW/4, y-4.5*mm, "208-174145-04-018 박철규 (알리고 미디어)")
    y -= 8*mm

    c.setFillColor(TBG); c.rect(ML, y-12*mm, PW*0.38, 12*mm, fill=1, stroke=1)
    c.setFillColor(colors.black); c.setFont(fb,11)
    c.drawCentredString(ML+PW*0.19, y-6*mm, "합계금액")
    c.setFont(fn,8)
    c.drawCentredString(ML+PW*0.19, y-10*mm, "(부가세 포함)" if is_tax else "(VAT 미포함)")
    c.setFillColor(LG); c.rect(ML+PW*0.38, y-12*mm, PW*0.47, 12*mm, fill=1, stroke=1)
    c.setFillColor(colors.black); c.setFont(fb,11)
    c.drawCentredString(ML+PW*0.615, y-7*mm, "진행 상품 상세 내역")
    c.setFillColor(colors.white); c.rect(ML+PW*0.85, y-12*mm, PW*0.15, 12*mm, fill=1, stroke=1)
    c.setFillColor(colors.black); c.setFont(fb,9)
    c.drawCentredString(ML+PW*0.925, y-7*mm, f"₩{tot:,}")
    y -= 14*mm

    cx = [ML, ML+12*mm, ML+70*mm, ML+105*mm, ML+118*mm, ML+136*mm, ML+154*mm]
    cw = [12*mm,58*mm,35*mm,13*mm,18*mm,18*mm]; cw.append(MR-cx[-1])
    labels = ["NO","품목","구성","수량","단가","공급가액(VAT별도)","비고"]
    c.setFillColor(TBG); c.rect(ML, y-6*mm, PW, 6*mm, fill=1, stroke=1)
    c.setFillColor(colors.black); c.setFont(fb,8)
    for i,lbl in enumerate(labels):
        c.drawCentredString(cx[i]+cw[i]/2, y-4.5*mm, lbl)
        if i>0: c.setLineWidth(0.5); c.line(cx[i],y,cx[i],y-6*mm)
    y -= 6*mm

    max_r = max(len(items),10); rh2=6*mm
    for i in range(max_r):
        c.setFillColor(LG if i%2==0 else colors.white)
        c.rect(ML, y-rh2, PW, rh2, fill=1, stroke=0)
        c.setStrokeColor(MG); c.setLineWidth(0.3)
        c.rect(ML, y-rh2, PW, rh2, fill=0, stroke=1)
        c.setFillColor(colors.black); c.setFont(fn,8)
        c.drawCentredString(cx[0]+cw[0]/2, y-4.5*mm, str(i+1))
        if i < len(items):
            it=items[i]; qty=int(it["수량"]); price=int(it["단가"]); sp=qty*price
            c.drawString(cx[1]+2*mm, y-4.5*mm, str(it.get("품목","")))
            c.drawString(cx[2]+2*mm, y-4.5*mm, str(it.get("구성","")))
            c.drawCentredString(cx[3]+cw[3]/2, y-4.5*mm, f"{qty:,}")
            c.drawRightString(cx[4]+cw[4]-1*mm, y-4.5*mm, f"{price:,}")
            c.drawRightString(cx[5]+cw[5]-1*mm, y-4.5*mm, f"{sp:,}")
            c.drawCentredString(cx[6]+cw[6]/2, y-4.5*mm, str(it.get("비고","")))
        else:
            c.drawRightString(cx[5]+cw[5]-1*mm, y-4.5*mm, "0")
        for j in range(1,len(cx)): c.setLineWidth(0.3); c.line(cx[j],y,cx[j],y-rh2)
        y -= rh2

    sums = [("공급가액 합계",sup),("세  액 (VAT)",vat),("합  계(부가세 포함)",tot)] if is_tax else [("공급가액 합계 (VAT 미발행)",sup)]
    for lbl,amt in sums:
        c.setFillColor(TBG); c.rect(ML, y-7*mm, PW, 7*mm, fill=1, stroke=1)
        c.setFillColor(colors.black); c.setFont(fb,9)
        c.drawCentredString(ML+PW*0.5, y-4.8*mm, lbl)
        c.drawRightString(cx[5]+cw[5]-1*mm, y-4.8*mm, f"{amt:,}")
        y -= 7*mm

    y -= 5*mm
    c.setFont(fb,9); c.setFillColor(colors.black)
    c.drawString(ML, y, "▶ 입금 계좌번호 : IBK기업은행 208-174145-04-018 박철규 (알리고 미디어)")
    y -= 6*mm
    c.drawString(ML, y, f"▶ 비  고 : {quote_data.get('memo','')}")
    c.save(); buf.seek(0)
    return buf

# =====================================================
st.set_page_config(page_title="버즈필터 자동화", page_icon="🤖", layout="wide")

with st.sidebar:
    st.markdown("## 📋 메뉴")
    st.markdown("---")
    menu = st.radio("", options=["🏭 버즈필터 발주", "📝 리뷰 입력", "📄 견적서 생성"], label_visibility="collapsed")
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
                ms = get_sheet("2. 버즈필터 마진 계산기")
                if ms is None: st.stop()
                amd = ms.get_all_values()
                calc_df = pd.DataFrame(amd[2:], columns=amd[1])
                calc_df = calc_df[calc_df['제품명'].str.strip() != '']
                st.success(f"✅ 마진계산기 로드 완료 ({len(calc_df)}개 상품)")
            with st.spinner("AI가 상품 매칭 중..."):
                today = datetime.now()
                rows_to_add, match_results = [], []

                # ── 텍스트 정규화 함수 ──
                def normalize(text):
                    import unicodedata
                    text = str(text).lower()
                    text = unicodedata.normalize('NFC', text)
                    # 공백 통일
                    text = re.sub(r'\s+', '', text)
                    # 특수문자 제거
                    text = re.sub(r'[^\w가-힣]', '', text)
                    return text

                # ── 1차: 텍스트 유사도로 후보 좁히기 ──
                def get_candidates(raw_name, calc_df, top_n=10):
                    norm_raw = normalize(raw_name)
                    scores = []
                    for _, r in calc_df.iterrows():
                        norm_prod = normalize(str(r.get('제품명','')))
                        norm_brand = normalize(str(r.get('브랜드','')))
                        combined = norm_brand + norm_prod

                        # 공통 글자 비율 계산
                        score = 0
                        for ch in norm_raw:
                            if ch in combined:
                                score += 1
                        score = score / max(len(norm_raw), 1)
                        scores.append((score, r))

                    scores.sort(key=lambda x: x[0], reverse=True)
                    top = [r for s, r in scores[:top_n] if s > 0]
                    return top if top else [r for _, r in scores[:top_n]]

                for idx, row in df.iterrows():
                    raw = str(row.get('상품명+옵션+개수',''))
                    qty = int(row.get('수량',1))
                    price = int(row.get('가격',0))
                    ch = str(row.get('판매처','쿠팡'))

                    # 후보 상품만 추려서 AI에 전달 (전체 목록 대신)
                    candidates = get_candidates(raw, calc_df, top_n=10)
                    cand_df = pd.DataFrame(candidates)

                    prompt = f"""너는 공기청정기/필터 상품 매칭 전문가야.

[중요 규칙]
- 띄어쓰기, 공백 차이는 무시해 (예: "헤파필터" = "헤파 필터" = "헤 파필터")
- 브랜드명이 다르면 절대 매칭하지 마
- 모델명/시리즈명이 핵심이야 (예: ACL-120Z0, CDH, R톨 등)
- 필터 종류가 다르면 다른 상품이야 (헤파≠탈취≠기능성)
- 확신이 없으면 반드시 미등록으로 답해

발주서 상품명: {raw}

후보 상품 리스트 (브랜드 / 제품명 / 상품코드):
{cand_df[['브랜드','제품명','상품코드 표']].to_string(index=False) if '브랜드' in cand_df.columns else calc_df[['브랜드','제품명','상품코드 표']].head(10).to_string(index=False)}

반드시 아래 형식으로만 답해. 다른 말 절대 금지:
상품코드: [코드값]
브랜드: [브랜드값]

매칭 불가 시:
상품코드: 미등록
브랜드: 미등록"""

                    resp = client.messages.create(
                        model="claude-haiku-4-5-20251001",
                        max_tokens=100,
                        messages=[{"role":"user","content":prompt}]
                    )
                    rt = resp.content[0].text.strip()
                    mc, mb = "미등록", "미등록"
                    for line in rt.split('\n'):
                        if '상품코드:' in line: mc = line.split('상품코드:')[1].strip()
                        if '브랜드:' in line: mb = line.split('브랜드:')[1].strip()
                    match_results.append({'상품명':raw,'매칭 브랜드':mb,'매칭 코드':mc,'판매처':ch,'가격':price,'수량':qty})
                    rows_to_add.append([str(today.year),str(today.month),str(today.day),mb,mc,ch,price,qty])
                st.session_state['rows_to_add'] = rows_to_add
                st.session_state['match_results'] = match_results
                st.session_state['ready_to_insert'] = True
            st.success("✅ AI 매칭 완료!")
        if st.session_state.get('ready_to_insert'):
            rdf = pd.DataFrame(st.session_state['match_results'])
            st.write("🔍 AI 매칭 결과:"); st.dataframe(rdf)
            unm = rdf[rdf['매칭 코드']=='미등록']
            if len(unm) > 0: st.warning(f"⚠️ {len(unm)}건 매칭 실패")
            if st.button("✅ 확인했습니다. 장부에 최종 입력합니다."):
                with st.spinner("장부 입력 중..."):
                    try:
                        ls = get_sheet("2. 버즈필터 장부")
                        if ls is None: st.stop()
                        sr = find_last_data_row(ls)
                        st.info(f"📍 {sr}행부터 입력 시작")
                        insert_row_safe(ls, sr, st.session_state['rows_to_add'])
                        st.success(f"🎉 총 {len(st.session_state['rows_to_add'])}건 입력 완료!")
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
        delim = re.compile(r'^\s*(?:\((\d+)\)|(\d+)[.\)]|(\d+))\s*$', re.MULTILINE)
        markers = [(int(m.group(1) or m.group(2) or m.group(3)), m.start(), m.end()) for m in delim.finditer(text)]
        if not markers: return []
        reviews = []
        for i,(num,start,end) in enumerate(markers):
            raw = text[end:markers[i+1][1]] if i+1<len(markers) else text[end:]
            content = raw.strip()
            if content: reviews.append((num, content))
        return sorted(reviews, key=lambda x: x[0])

    def create_excel(reviews):
        wb = Workbook(); ws = wb.active; ws.title = "리뷰"
        hf = PatternFill("solid", start_color="FF6B35", end_color="FF6B35")
        hfont = Font(bold=True, color="FFFFFF", name="Arial", size=11)
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        lw = Alignment(horizontal="left", vertical="top", wrap_text=True)
        thin = Side(style="thin", color="CCCCCC")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for addr, lbl in {"A1":"번호","B1":"별점","C1":"리뷰 내용"}.items():
            cell=ws[addr]; cell.value=lbl; cell.fill=hf; cell.font=hfont; cell.alignment=center; cell.border=border
        ws.column_dimensions["A"].width=8; ws.column_dimensions["B"].width=8; ws.column_dimensions["C"].width=70
        ws.row_dimensions[1].height=30
        af = PatternFill("solid", start_color="FFF5F0", end_color="FFF5F0")
        for i,(num,content) in enumerate(reviews, start=2):
            rf = af if i%2==0 else None
            a=ws.cell(row=i,column=1,value=num); a.font=Font(name="Arial",size=10,bold=True); a.alignment=center; a.border=border
            if rf: a.fill=rf
            b=ws.cell(row=i,column=2,value=""); b.border=border
            if rf: b.fill=rf
            cc=ws.cell(row=i,column=3,value=content); cc.font=Font(name="Arial",size=10); cc.alignment=lw; cc.border=border
            if rf: cc.fill=rf
            ws.row_dimensions[i].height=max(40,min(content.count('\n')*18+18,200))
        out=io.BytesIO(); wb.save(out); out.seek(0); return out

    tab1, tab2 = st.tabs(["📁 파일 업로드", "✏️ 텍스트 직접 입력"])
    with tab1:
        utxt = st.file_uploader("리뷰 텍스트 파일 (.txt)", type=["txt"])
        if utxt:
            txt = utxt.read().decode("utf-8", errors="ignore")
            st.success(f"✅ {utxt.name} 업로드 완료")
            revs = parse_reviews(txt)
            if revs:
                st.markdown(f"### 📊 **{len(revs)}개** 리뷰 감지됨")
                with st.expander("👀 미리보기", expanded=True):
                    for num,content in revs[:5]:
                        st.markdown(f"**{num}번 리뷰**"); st.text(content[:200]+("..." if len(content)>200 else "")); st.divider()
                    if len(revs)>5: st.info(f"... 외 {len(revs)-5}개")
                st.download_button("⬇️ 엑셀 다운로드", create_excel(revs), "리뷰목록.xlsx", use_container_width=True, type="primary")
            else:
                st.error("❌ 리뷰를 파싱할 수 없습니다.")
    with tab2:
        mt = st.text_area("리뷰 내용 붙여넣기", height=300, placeholder='1\n리뷰 내용...\n\n2\n리뷰 내용...')
        if mt.strip():
            revs = parse_reviews(mt)
            if revs:
                st.markdown(f"### 📊 **{len(revs)}개** 리뷰 감지됨")
                with st.expander("👀 미리보기", expanded=True):
                    for num,content in revs[:5]:
                        st.markdown(f"**{num}번 리뷰**"); st.text(content[:200]+("..." if len(content)>200 else "")); st.divider()
                st.download_button("⬇️ 엑셀 다운로드", create_excel(revs), "리뷰목록.xlsx", use_container_width=True, type="primary")
            else:
                st.error("❌ 리뷰를 파싱할 수 없습니다.")

# =====================================================
# 페이지 3: 견적서 생성
# =====================================================
elif menu == "📄 견적서 생성":
    st.title("📄 견적서 자동 생성")
    st.subheader("정보를 입력하면 PDF 견적서를 자동으로 만들어드립니다.")

    col1, col2, col3 = st.columns(3)
    with col1: client_name = st.text_input("고객사명", placeholder="예) 지케이라이프")
    with col2: quote_date = st.text_input("견적일", value=datetime.now().strftime("%Y. %m. %d"))
    with col3: tax_type = st.radio("계산서 발행 여부", ["발행", "미발행"], horizontal=True)
    memo = st.text_input("비고 (선택)", placeholder="예) 패키지 할인 포함")

    st.markdown("---")
    st.markdown("### 📋 항목 입력")
    st.caption("➕ 항목 추가 버튼으로 행을 늘리고, 🗑️ 버튼으로 삭제하세요.")

    if 'quote_items' not in st.session_state:
        st.session_state.quote_items = [{"품목":"","구성":"","수량":1,"단가":0,"비고":""}]

    hcols = st.columns([3,2,1,2,2,1])
    for col,lbl in zip(hcols,["**품목**","**구성**","**수량**","**단가(원)**","**공급가액**","**삭제**"]):
        col.markdown(lbl)

    to_del = []
    for i,item in enumerate(st.session_state.quote_items):
        cols = st.columns([3,2,1,2,2,1])
        item["품목"] = cols[0].text_input(f"p{i}", value=item["품목"], label_visibility="collapsed", placeholder="예) 쿠팡 리뷰")
        item["구성"] = cols[1].text_input(f"g{i}", value=item["구성"], label_visibility="collapsed", placeholder="예) 실행비")
        item["수량"] = cols[2].number_input(f"q{i}", value=item["수량"], min_value=0, label_visibility="collapsed")
        item["단가"] = cols[3].number_input(f"u{i}", value=item["단가"], min_value=0, step=100, label_visibility="collapsed")
        sp = item["수량"] * item["단가"]
        cols[4].markdown(f"<div style='padding:8px 0;font-weight:bold;'>₩{sp:,}</div>", unsafe_allow_html=True)
        if cols[5].button("🗑️", key=f"d{i}"): to_del.append(i)

    for i in sorted(to_del, reverse=True): st.session_state.quote_items.pop(i)
    if to_del: st.rerun()

    if st.button("➕ 항목 추가"):
        st.session_state.quote_items.append({"품목":"","구성":"","수량":1,"단가":0,"비고":""})
        st.rerun()

    st.markdown("---")
    valid = [it for it in st.session_state.quote_items if it["품목"].strip()]
    sup = sum(it["수량"]*it["단가"] for it in valid)
    vat = int(sup*0.1) if tax_type=="발행" else 0
    tot = sup + vat

    mc1,mc2,mc3 = st.columns(3)
    mc1.metric("공급가액 합계", f"₩{sup:,}")
    if tax_type == "발행":
        mc2.metric("세액 (VAT 10%)", f"₩{vat:,}")
        mc3.metric("최종 합계 (부가세 포함)", f"₩{tot:,}")
    else:
        mc2.metric("계산서", "미발행")
        mc3.metric("최종 합계 (VAT 없음)", f"₩{tot:,}")

    st.markdown("---")
    if st.button("📄 견적서 PDF 생성", type="primary", use_container_width=True):
        if not client_name.strip():
            st.error("❌ 고객사명을 입력해주세요!")
        elif not valid:
            st.error("❌ 항목을 최소 1개 이상 입력해주세요!")
        else:
            with st.spinner("PDF 생성 중..."):
                BASE_DIR = os.path.dirname(os.path.abspath(__file__))
                stamp_path = os.path.join(BASE_DIR, '직인_투명.png')
                qd = {"date":quote_date,"client":client_name,"tax_type":tax_type,"memo":memo,"items":valid}
                try:
                    pdf_buf = generate_quote_pdf(qd, stamp_path)
                    fname = f"견적서_{client_name}_{quote_date.replace('. ','').replace('.','')}.pdf"
                    st.success("✅ 견적서 PDF 생성 완료!")
                    st.download_button("⬇️ PDF 다운로드", pdf_buf, fname,
                                       mime="application/pdf", use_container_width=True, type="primary")
                except Exception as e:
                    st.error(f"❌ PDF 생성 실패: {e}")
                    st.info("💡 NotoSansKR-Regular.ttf, NotoSansKR-Bold.ttf 파일이 같은 폴더에 있는지 확인해주세요!")
