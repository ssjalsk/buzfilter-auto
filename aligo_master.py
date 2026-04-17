# -*- coding: utf-8 -*-
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
import requests
import zipfile
import base64
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

def extract_qty_from_text(text):
    match = re.search(r'/\s*(\d+)\s*(세트|개)', str(text))
    if match:
        return int(match.group(1))
    return 1

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
    c.setFont(fb, 28); c.setFillColor(colors.black)
    c.drawCentredString(w/2, y, "견   적   서")
    y -= 10*mm
    c.setStrokeColor(colors.black); c.setLineWidth(1.5); c.line(ML, y, MR, y)
    y -= 8*mm
    bt = y
    rcx, rcw = ML+PW*0.5, PW*0.5
    c.setFont(fb, 18); c.setFillColor(colors.HexColor("#1a5fa8"))
    c.drawString(ML+5*mm, bt-12*mm, "Aligo")
    c.setFont(fb, 14); c.drawString(ML+5*mm, bt-20*mm, "Media")
    if stamp_path and os.path.exists(stamp_path):
        try: c.drawImage(stamp_path, rcx-24*mm, bt-33*mm, width=22*mm, height=22*mm, mask='auto')
        except: pass
    c.setFillColor(DG); c.rect(rcx, bt-6*mm, rcw, 6*mm, fill=1, stroke=0)
    c.setFillColor(colors.white); c.setFont(fb, 10)
    c.drawCentredString(rcx+rcw/2, bt-4.5*mm, "공  급  자")
    srows = [("등록번호","161-22-02310","대표자","박철규"),("상  호","알리고미디어","",""),
             ("주  소","서울 마포구 양화로64, 8층","",""),("연락처","010-9469-2381","",""),
             ("업  태","전문, 서비스업","종 목","광고대행업")]
    rh = 5.5*mm
    for i,(k1,v1,k2,v2) in enumerate(srows):
        ry = bt-6*mm-(i+1)*rh
        c.setFillColor(LG if i%2==0 else colors.white); c.rect(rcx, ry, rcw, rh, fill=1, stroke=0)
        c.setStrokeColor(MG); c.setLineWidth(0.5); c.rect(rcx, ry, rcw, rh, fill=0, stroke=1)
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
    c.setFillColor(colors.black); c.setFont(fn,9); c.drawString(ML+30*mm, y-4.5*mm, quote_data["date"])
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
    c.setFont(fn,8); c.drawCentredString(ML+PW*0.19, y-10*mm, "(부가세 포함)" if is_tax else "(VAT 미포함)")
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
        c.setFillColor(LG if i%2==0 else colors.white); c.rect(ML, y-rh2, PW, rh2, fill=1, stroke=0)
        c.setStrokeColor(MG); c.setLineWidth(0.3); c.rect(ML, y-rh2, PW, rh2, fill=0, stroke=1)
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

def deploy_to_netlify(html_content, site_id, token, extra_files=None):
    try:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr('index.html', html_content.encode('utf-8'))
            if extra_files:
                for filename, file_bytes in extra_files.items():
                    zf.writestr(filename, file_bytes)
        zip_buffer.seek(0)
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/zip'
        }
        response = requests.post(
            f'https://api.netlify.com/api/v1/sites/{site_id}/deploys',
            headers=headers,
            data=zip_buffer.getvalue(),
            timeout=60
        )
        if response.status_code in [200, 201]:
            return True, "성공"
        else:
            return False, f"오류 코드: {response.status_code}\n{response.text[:300]}"
    except Exception as e:
        return False, str(e)

# =====================================================
# 리뷰 파싱 (기존 유지)
# =====================================================
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

# =====================================================
# 엑셀 생성 (기존 유지 + 리뷰 생성기에서도 재사용)
# =====================================================
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

# =====================================================
# [신규] 리뷰 AI 생성 함수 - 배치 분할 방식
# =====================================================
def generate_reviews_with_claude(client, product_info, selling_points, review_count, char_count, image_data=None, progress_callback=None):
    import random

    BATCH_SIZE = 20

    persona_pool = [
        "20대 초반 여성 대학생", "20대 중반 직장 여성", "20대 후반 직장 여성",
        "20대 초반 남성 대학생", "20대 후반 남성 직장인",
        "30대 초반 주부", "30대 중반 워킹맘", "30대 후반 주부",
        "30대 초반 남성 직장인", "30대 후반 남성 직장인",
        "30대 초반 여성 자영업자", "30대 후반 여성 자영업자",
        "40대 초반 주부", "40대 중반 주부", "40대 후반 주부",
        "40대 초반 남성 회사원", "40대 중반 남성 회사원", "40대 후반 남성 회사원",
        "40대 여성 자영업자", "40대 워킹맘",
        "50대 초반 주부", "50대 후반 주부",
        "50대 초반 남성", "50대 후반 남성",
        "60대 초반 여성", "60대 후반 여성",
        "20대 여성 프리랜서", "30대 여성 교사", "40대 여성 간호사",
        "20대 남성 군인", "30대 남성 공무원", "40대 남성 자영업자",
        "30대 싱글 여성", "30대 신혼 여성", "40대 싱글 남성",
        "20대 후반 여성 간호사", "30대 여성 약사", "40대 여성 교사",
        "50대 남성 공무원", "60대 남성 은퇴자",
        "20대 남성 배달기사", "30대 남성 IT개발자", "40대 남성 의사",
        "20대 여성 헤어디자이너", "30대 여성 요가강사", "40대 여성 영양사",
        "50대 여성 교사", "60대 여성 주부", "30대 워킹대디", "40대 워킹대디"
    ]

    random.shuffle(persona_pool)
    all_personas = [persona_pool[i % len(persona_pool)] for i in range(review_count)]

    all_reviews = []

    batches = []
    for start in range(0, review_count, BATCH_SIZE):
        end = min(start + BATCH_SIZE, review_count)
        batches.append((start, end))

    for batch_idx, (start, end) in enumerate(batches):
        batch_num = batch_idx + 1
        batch_personas = all_personas[start:end]
        batch_count = end - start
        global_start_num = start + 1

        persona_text = "\n".join([
            f"{global_start_num + i}번 리뷰: {p}"
            for i, p in enumerate(batch_personas)
        ])

        prev_summary = ""
        if all_reviews:
            recent = all_reviews[-20:]
            prev_lines = []
            for num, content in recent:
                first_line = content.split('\n')[0][:60]
                prev_lines.append(f"- {num}번: {first_line}...")
            prev_summary = "\n[이미 작성된 리뷰 도입부 (절대 유사하게 쓰지 말 것)]\n" + "\n".join(prev_lines)

        user_content = []
        if image_data and batch_idx == 0:
            user_content.append({
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": image_data["media_type"],
                    "data": image_data["data"]
                }
            })

        prompt = f"""너는 실제 구매자처럼 자연스러운 한국어 리뷰를 쓰는 전문 작가야.

[제품 정보]
{product_info}

[소구점 / 강조할 내용]
{selling_points if selling_points else "없음 (제품 정보 기반으로 자유롭게 작성)"}

[작성 조건]
- 이번 배치: {global_start_num}번 ~ {end}번 리뷰 (총 {batch_count}개)
- 리뷰당 글자 수: 약 {char_count}자 내외
- 각 리뷰는 아래 페르소나에 맞게 말투와 내용을 다르게 작성

[페르소나 배정]
{persona_text}
{prev_summary}

[필수 규칙]
1. 번호는 {global_start_num}부터 시작, 각 리뷰는 "번호." 한 줄 후 리뷰 내용
   예시:
   {global_start_num}.
   리뷰 내용...

   {global_start_num+1}.
   리뷰 내용...

2. 그림 이모지 절대 사용 금지 (🌈 ☂️ ❤️ ⭐ 등 유니코드 이모티콘 전부 금지)
3. 텍스트 감성 표현 자연스럽게 허용 (ㅎㅎ, ㅋㅋ, ~!, ~~, !!, ㅠㅠ 등)
4. 리뷰 간 표현·문장구조·도입부·마무리 절대 중복 금지
5. 페르소나에 맞는 실제 사람 말투 사용
   - 20대: 가볍고 솔직한 톤, 줄임말 가능
   - 30-40대: 실용적이고 구체적인 톤
   - 50-60대: 차분하고 정중한 톤
6. 구매 동기, 사용 경험, 구체적 디테일이 각 리뷰마다 달라야 함

정확히 {batch_count}개 리뷰를 작성해줘. 설명이나 부연 없이 리뷰만 출력해.
"""

        user_content.append({"type": "text", "text": prompt})

        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=8000,
            messages=[{"role": "user", "content": user_content}]
        )

        raw_text = response.content[0].text.strip()
        batch_reviews = parse_generated_reviews(raw_text)

        if batch_reviews:
            all_reviews.extend(batch_reviews)

        if progress_callback:
            progress_callback(batch_idx + 1, len(batches), len(all_reviews))

    return all_reviews

# =====================================================
# 리뷰 텍스트 파싱 (AI 생성 결과용 - 숫자. 형식)
# =====================================================
def parse_generated_reviews(text):
    lines = text.split('\n')
    reviews = []
    current_num = None
    current_lines = []

    for line in lines:
        stripped = line.strip()
        num_match = re.match(r'^(\d+)[.\)]?\s*$', stripped)
        if num_match:
            if current_num is not None and current_lines:
                content = '\n'.join(current_lines).strip()
                if content:
                    reviews.append((current_num, content))
            current_num = int(num_match.group(1))
            current_lines = []
        else:
            if current_num is not None:
                current_lines.append(line)

    if current_num is not None and current_lines:
        content = '\n'.join(current_lines).strip()
        if content:
            reviews.append((current_num, content))

    return sorted(reviews, key=lambda x: x[0])


# =====================================================
# Streamlit 앱 시작
# =====================================================
st.set_page_config(page_title="버즈필터 자동화", page_icon="🤖", layout="wide")

with st.sidebar:
    st.markdown("## 📋 메뉴")
    st.markdown("---")
    menu = st.radio("", options=[
        "🏭 버즈필터 발주",
        "✍️ 리뷰 생성",
        "📝 리뷰 입력",
        "📄 견적서 생성",
        "🌐 홈페이지 자동 개선"
    ], label_visibility="collapsed")
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
                for idx, row in df.iterrows():
                    raw = str(row.get('상품명+옵션+개수', ''))
                    qty = extract_qty_from_text(raw)
                    price = int(row.get('가격', 0))
                    ch = str(row.get('판매처', '쿠팡'))

                    # ✅ 버그 수정: 마크다운 금지 + 유사 매칭 강화
                    prompt = f"""너는 상품 매칭 전문가야.
발주서 상품명: {raw}
상품 리스트:
{calc_df[['브랜드','제품명','상품코드 표']].to_string(index=False)}

규칙:
1. 브랜드명이나 핵심 키워드가 겹치면 매칭해라 (정확히 일치 안해도 됨)
2. 예: "다이슨 공기청정기 시리즈" → 다이슨 관련 상품 중 가장 유사한 것 1개
3. 예: "위닉스 뽀송" → 위닉스 뽀송 관련 상품 1개
4. 가장 유사한 것 1개만 선택해라
5. 마크다운 절대 사용 금지. 코드값에 **, *, _ 같은 기호 절대 붙이지 마라
6. 정말 모르겠으면 미등록

반드시 아래 형식으로만 답해줘:
상품코드: [코드값]
브랜드: [브랜드값]
못 찾겠으면:
상품코드: 미등록
브랜드: 미등록"""

                    resp = client.messages.create(model="claude-haiku-4-5-20251001", max_tokens=100, messages=[{"role":"user","content":prompt}])
                    rt = resp.content[0].text.strip()
                    mc, mb = "미등록", "미등록"
                    for line in rt.split('\n'):
                        # ✅ 버그 수정: ** 등 마크다운 기호 제거
                        if '상품코드:' in line:
                            mc = line.split('상품코드:')[1].strip().replace('**', '').replace('*', '').replace('_', '').strip()
                        if '브랜드:' in line:
                            mb = line.split('브랜드:')[1].strip().replace('**', '').replace('*', '').replace('_', '').strip()

                    match_results.append({'상품명': raw, '매칭 브랜드': mb, '매칭 코드': mc, '판매처': ch, '가격': price, '수량(파싱)': qty})
                    rows_to_add.append([f"{today.year}년", f"{today.month}월", f"{today.day}일", mb, mc, ch, price, qty])
                st.session_state['rows_to_add'] = rows_to_add
                st.session_state['match_results'] = match_results
                st.session_state['ready_to_insert'] = True
            st.success("✅ AI 매칭 완료!")

        if st.session_state.get('ready_to_insert'):
            rdf = pd.DataFrame(st.session_state['match_results'])
            st.write("🔍 AI 매칭 결과")
            st.dataframe(rdf)
            no_slash = rdf[~rdf['상품명'].str.contains('/', na=False)]
            if len(no_slash) > 0:
                st.warning(f"⚠️ {len(no_slash)}건 수량 파싱 불가 → 수량 1로 처리")
                st.dataframe(no_slash[['상품명', '수량(파싱)']])
            unm = rdf[rdf['매칭 코드'] == '미등록']
            if len(unm) > 0:
                st.warning(f"⚠️ {len(unm)}건 상품 매칭 실패")
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
# 페이지 2: 리뷰 생성 (신규)
# =====================================================
elif menu == "✍️ 리뷰 생성":
    st.title("✍️ AI 리뷰 자동 생성기")
    st.subheader("제품 정보를 입력하면 자연스럽고 다양한 리뷰를 생성해드립니다.")

    if 'generated_reviews' not in st.session_state:
        st.session_state.generated_reviews = []
    if 'review_edit_mode' not in st.session_state:
        st.session_state.review_edit_mode = False

    st.markdown("### 📦 STEP 1 — 제품 정보 입력")

    col_left, col_right = st.columns([2, 1])
    with col_left:
        product_info = st.text_area(
            "제품 정보 (제품명, 카테고리, 특징 등)",
            height=150,
            placeholder="예)\n제품명: 콜라겐 마스크팩\n카테고리: 스킨케어\n특징: 저자극 성분, 수분 집중 케어, 붙임성 좋음, 개별 포장"
        )
        selling_points = st.text_area(
            "소구점 / 강조할 내용 (선택)",
            height=100,
            placeholder="예) 피부 흡수력, 아침에 쓰기 좋음, 가성비, 선물용으로 좋다는 점 강조"
        )
    with col_right:
        product_image = st.file_uploader(
            "제품 이미지 (선택)",
            type=["jpg", "jpeg", "png", "webp"],
            help="이미지를 함께 주면 Claude가 더 정확하게 리뷰를 작성합니다."
        )
        if product_image:
            st.image(product_image, caption="업로드된 이미지", use_container_width=True)

    st.markdown("### ⚙️ STEP 2 — 리뷰 설정")
    col1, col2 = st.columns(2)
    with col1:
        review_count = st.number_input("리뷰 개수", min_value=1, max_value=100, value=10, step=1)
    with col2:
        char_count = st.number_input("리뷰당 글자 수 (약)", min_value=50, max_value=500, value=150, step=10)

    st.markdown("---")

    total_batches = max(1, (review_count + 19) // 20)
    st.caption(f"💡 {review_count}개 요청 시 20개씩 {total_batches}번 나눠 생성됩니다.")

    if st.button("🚀 리뷰 생성 시작", type="primary", use_container_width=True):
        if not product_info.strip():
            st.error("❌ 제품 정보를 입력해주세요!")
        else:
            try:
                ai_client = get_anthropic_client()

                image_data = None
                if product_image:
                    product_image.seek(0)
                    img_bytes = product_image.read()
                    img_b64 = base64.b64encode(img_bytes).decode('utf-8')
                    ext = product_image.name.split('.')[-1].lower()
                    media_type_map = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png", "webp": "image/webp"}
                    image_data = {
                        "media_type": media_type_map.get(ext, "image/jpeg"),
                        "data": img_b64
                    }

                progress_bar = st.progress(0)
                status_text = st.empty()

                def update_progress(current_batch, total_b, total_generated):
                    pct = int((current_batch / total_b) * 100)
                    progress_bar.progress(pct)
                    status_text.text(f"⏳ 배치 {current_batch}/{total_b} 완료 — 현재까지 {total_generated}개 생성됨")

                status_text.text(f"🚀 리뷰 생성 시작 (총 {total_batches}번 배치 호출)...")

                parsed = generate_reviews_with_claude(
                    client=ai_client,
                    product_info=product_info,
                    selling_points=selling_points,
                    review_count=review_count,
                    char_count=char_count,
                    image_data=image_data,
                    progress_callback=update_progress
                )

                progress_bar.progress(100)
                status_text.empty()

                st.session_state.generated_reviews = list(parsed)
                st.session_state.review_edit_mode = True
                st.success(f"✅ 총 {len(parsed)}개 리뷰 생성 완료!")

            except Exception as e:
                st.error(f"❌ 생성 실패: {e}")

    if st.session_state.review_edit_mode and st.session_state.generated_reviews:
        st.markdown("---")
        st.markdown(f"### 📋 STEP 3 — 결과 확인 및 수정 ({len(st.session_state.generated_reviews)}개)")
        st.caption("각 리뷰를 직접 수정할 수 있습니다. 수정 후 아래 [저장 및 엑셀 다운로드] 버튼을 눌러주세요.")

        updated_reviews = []
        for i, (num, content) in enumerate(st.session_state.generated_reviews):
            with st.expander(f"리뷰 {num}번", expanded=(i < 3)):
                edited = st.text_area(
                    f"리뷰 {num} 내용",
                    value=content,
                    height=150,
                    key=f"review_edit_{i}",
                    label_visibility="collapsed"
                )
                updated_reviews.append((num, edited))

        st.markdown("---")
        st.markdown("### 💾 STEP 4 — 저장 및 다운로드")

        col_save, col_reset = st.columns([3, 1])
        with col_save:
            if st.button("⬇️ 저장 및 엑셀 다운로드", type="primary", use_container_width=True):
                st.session_state.generated_reviews = updated_reviews
                excel_data = create_excel(updated_reviews)
                fname = f"리뷰_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                st.download_button(
                    label="📥 엑셀 파일 다운로드 클릭",
                    data=excel_data,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
                st.success("✅ 엑셀 파일이 준비되었습니다!")

        with col_reset:
            if st.button("🔄 초기화", use_container_width=True):
                st.session_state.generated_reviews = []
                st.session_state.review_edit_mode = False
                st.rerun()

        with st.expander("📄 텍스트 전체 보기 (복사용)"):
            full_text = ""
            for num, content in st.session_state.generated_reviews:
                full_text += f"{num}.\n{content}\n\n"
            st.text_area("전체 리뷰 텍스트", value=full_text.strip(), height=400, label_visibility="collapsed")

# =====================================================
# 페이지 3: 리뷰 입력 (기존 유지)
# =====================================================
elif menu == "📝 리뷰 입력":
    st.title("📝 리뷰 엑셀 자동 변환기")
    st.subheader("리뷰 텍스트 파일을 업로드하면 엑셀 파일로 자동 변환합니다.")

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
# 페이지 4: 견적서 생성 (기존 유지)
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

# =====================================================
# 페이지 5: 홈페이지 자동 개선 (기존 유지)
# =====================================================
elif menu == "🌐 홈페이지 자동 개선":
    st.title("🌐 홈페이지 자동 개선 + 자동 배포")
    st.subheader("HTML과 이미지를 업로드하면 Claude가 수정하고 Netlify에 자동 배포합니다.")

    try:
        NETLIFY_TOKEN = st.secrets["NETLIFY_TOKEN"]
        NETLIFY_SITE_ID = st.secrets["NETLIFY_SITE_ID"]
        netlify_ready = True
        st.success("✅ Netlify 연결 준비 완료 — 버튼 한 번으로 자동 배포됩니다.")
    except:
        netlify_ready = False
        st.error("❌ Netlify 미연결 — secrets.toml에 아래 두 줄을 추가해주세요.")
        st.code("""NETLIFY_TOKEN = "발급받은_토큰"\nNETLIFY_SITE_ID = "57340c83-2554-459c-9a49-b29fbdb9b0c0" """, language="toml")

    st.markdown("---")
    st.markdown("### 📂 STEP 1 — index.html 업로드")
    uploaded_html = st.file_uploader("index.html 업로드", type=["html","htm"], key="html_upload")
    st.markdown("### 🖼️ STEP 2 — 이미지 파일 업로드")
    st.caption("image_4.png, pr_chat.png, review_chat.png 등 홈페이지에 쓰이는 이미지를 모두 올려주세요.")
    uploaded_images = st.file_uploader(
        "이미지 파일 (여러 개 동시 선택 가능)",
        type=["png","jpg","jpeg","gif","webp","ico"],
        accept_multiple_files=True,
        key="image_upload"
    )
    if uploaded_images:
        st.success(f"✅ 이미지 {len(uploaded_images)}개: {', '.join([f.name for f in uploaded_images])}")

    st.markdown("---")
    st.markdown("### ✅ STEP 3 — 개선 항목 선택")
    col1, col2, col3 = st.columns(3)
    with col1: check_mobile = st.checkbox("📱 모바일 최적화", value=True)
    with col2: check_responsive = st.checkbox("📐 반응형 디자인", value=True)
    with col3: check_seo = st.checkbox("🔍 구글 SEO", value=True)
    check_extra = st.text_area("📝 추가 요청사항 (선택)", placeholder="예) 버튼 색상을 더 눈에 띄게 / CTA 문구 강하게", height=80)
    st.markdown("---")

    if uploaded_html:
        html_content = uploaded_html.read().decode("utf-8", errors="ignore")
        if not any([check_mobile, check_responsive, check_seo, check_extra.strip()]):
            st.warning("⚠️ 개선 항목을 최소 1개 이상 선택해주세요.")
        else:
            if st.button("🚀 Claude 수정 + Netlify 자동 배포", type="primary", use_container_width=True):
                check_list = []
                if check_mobile: check_list.append("1. 모바일 최적화 (터치 타겟 44px, iOS 폰트 방지, 햄버거 메뉴, 모바일 CTA)")
                if check_responsive: check_list.append("2. 반응형 디자인 (브레이크포인트 3단계, clamp() 폰트, 그리드 자동 전환)")
                if check_seo: check_list.append("3. 구글 SEO (og:image 절대경로, LocalBusiness 구조화 데이터, main/address 태그)")
                if check_extra.strip(): check_list.append(f"4. 추가 요청: {check_extra.strip()}")

                prompt = f"""너는 웹 개발 전문가야. 아래 HTML을 분석하고 수정해서 완성된 HTML 코드만 반환해줘.
[개선 항목]
{chr(10).join(check_list)}
[주의사항]
- 수정된 HTML 전체 코드만 반환 (설명 없이 <!DOCTYPE html>부터 시작)
- 기존 디자인·색상·브랜드 정체성 유지
- 기존 기능(카운터 애니메이션, 관리자 시크릿 클릭) 유지
- 한국어 텍스트 수정 금지
[원본 HTML]
{html_content}"""

                progress_bar = st.progress(0)
                status_text = st.empty()
                try:
                    status_text.text("🤖 Claude가 분석 및 수정 중... (30초~1분 소요)")
                    progress_bar.progress(20)
                    ai_client = get_anthropic_client()
                    response = ai_client.messages.create(
                        model="claude-opus-4-5",
                        max_tokens=16000,
                        messages=[{"role": "user", "content": prompt}]
                    )
                    improved_html = response.content[0].text.strip()
                    if improved_html.startswith("```"):
                        lines = improved_html.split("\n")
                        improved_html = "\n".join(lines[1:-1] if lines[-1].strip() == "```" else lines[1:])
                    progress_bar.progress(60)
                    status_text.text("✅ Claude 수정 완료! Netlify 배포 중...")
                    extra_files = {}
                    if uploaded_images:
                        for img_file in uploaded_images:
                            img_file.seek(0)
                            extra_files[img_file.name] = img_file.read()
                    if netlify_ready:
                        success, result = deploy_to_netlify(improved_html, NETLIFY_SITE_ID, NETLIFY_TOKEN, extra_files)
                        progress_bar.progress(100)
                        if success:
                            status_text.text("🎉 완료!")
                            st.success("🎉 수정 완료 + Netlify 자동 배포 성공!")
                            st.balloons()
                            st.link_button("🔗 aligomedia.co.kr 확인하기", "https://aligomedia.co.kr")
                        else:
                            st.error(f"❌ Netlify 배포 실패\n\n{result}")
                    else:
                        progress_bar.progress(100)
                        status_text.text("✅ 수정 완료")
                    st.session_state["improved_html"] = improved_html
                    st.session_state["original_html"] = html_content
                    st.session_state["improvement_done"] = True
                except Exception as e:
                    st.error(f"❌ 오류 발생: {e}")
    else:
        st.info("👆 STEP 1에서 index.html을 업로드하면 시작할 수 있어요!")

    if st.session_state.get("improvement_done"):
        improved_html = st.session_state["improved_html"]
        original_html = st.session_state["original_html"]
        st.markdown("---")
        col_before, col_after = st.columns(2)
        with col_before:
            st.markdown("#### 📄 수정 전")
            st.metric("파일 크기", f"{len(original_html):,}자")
            with st.expander("원본 코드 보기"):
                st.code(original_html[:1500] + "...", language="html")
        with col_after:
            st.markdown("#### ✅ 수정 후")
            st.metric("파일 크기", f"{len(improved_html):,}자", delta=f"{len(improved_html)-len(original_html):+,}자")
            with st.expander("수정된 코드 보기"):
                st.code(improved_html[:1500] + "...", language="html")
        st.markdown("---")
        st.download_button(
            label="⬇️ 수정된 index.html 다운로드",
            data=improved_html.encode("utf-8"),
            file_name="index.html",
            mime="text/html",
            use_container_width=True
        )
        if st.button("🔄 처음부터 다시", use_container_width=True):
            st.session_state["improvement_done"] = False
            st.session_state["improved_html"] = ""
            st.rerun()
