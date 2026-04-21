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
    if not rows_data:
        return
    col_map = [('B', 0), ('C', 1), ('D', 2), ('E', 3), ('F', 4), ('H', 5), ('I', 6), ('K', 7)]
    updates = []
    for i, row in enumerate(rows_data):
        r = start_row + i
        for col, idx in col_map:
            updates.append({'range': f'{col}{r}', 'values': [[row[idx]]]})
    sheet.batch_update(updates, value_input_option='USER_ENTERED')

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

def analyze_images_with_claude(client, image_data_list):
    if not image_data_list:
        return ""
    content = []
    for img in image_data_list:
        content.append({"type": "image", "source": {"type": "base64", "media_type": img["media_type"], "data": img["data"]}})
    content.append({"type": "text", "text": """이 제품 이미지들을 분석해서 리뷰 작성에 활용할 수 있도록 아래 항목을 상세하게 설명해줘.
1. 제품 외관 및 디자인 (색상, 형태, 크기감, 패키징)
2. 제품에 표시된 텍스트, 로고, 브랜드명, 성분 등
3. 제품의 재질감, 질감, 마감 느낌
4. 이미지에서 보이는 특징적인 요소들
5. 전반적인 제품 분위기
리뷰 작가가 실제로 제품을 써본 것처럼 묘사할 수 있도록 구체적으로 써줘. 설명만 출력해."""})
    response = client.messages.create(model="claude-sonnet-4-6", max_tokens=1000, messages=[{"role": "user", "content": content}])
    return response.content[0].text.strip()

def generate_reviews_with_claude(client, product_info, selling_points, review_count, char_count, image_data_list=None, progress_callback=None):
    import random
    BATCH_SIZE = 20
    image_description = ""
    if image_data_list:
        image_description = analyze_images_with_claude(client, image_data_list)
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
        persona_text = "\n".join([f"{global_start_num + i}번 리뷰: {p}" for i, p in enumerate(batch_personas)])
        prev_summary = ""
        if all_reviews:
            recent = all_reviews[-20:]
            prev_lines = []
            for num, content in recent:
                first_line = content.split('\n')[0][:60]
                prev_lines.append(f"- {num}번: {first_line}...")
            prev_summary = "\n[이미 작성된 리뷰 도입부 (절대 유사하게 쓰지 말 것)]\n" + "\n".join(prev_lines)
        image_section = ""
        if image_description:
            image_section = f"\n[제품 이미지 분석 결과]\n{image_description}\n"
        prompt = f"""너는 실제 구매자처럼 자연스러운 한국어 리뷰를 쓰는 전문 작가야.
[제품 정보]
{product_info}
[소구점]
{selling_points if selling_points else "없음"}
{image_section}
[작성 조건]
- 이번 배치: {global_start_num}번 ~ {end}번 리뷰 (총 {batch_count}개)
- 리뷰당 글자 수: 약 {char_count}자 내외
[페르소나 배정]
{persona_text}
{prev_summary}
[필수 규칙]
1. 번호는 {global_start_num}부터 시작, 각 리뷰는 "번호." 한 줄 후 리뷰 내용
2. 그림 이모지 절대 사용 금지
3. 텍스트 감성 표현 자연스럽게 허용 (ㅎㅎ, ㅋㅋ 등)
4. 리뷰 간 표현·문장구조 절대 중복 금지
5. 페르소나에 맞는 실제 사람 말투 사용
정확히 {batch_count}개 리뷰를 작성해줘. 설명이나 부연 없이 리뷰만 출력해."""
        response = client.messages.create(model="claude-sonnet-4-6", max_tokens=8000, messages=[{"role": "user", "content": prompt}])
        raw_text = response.content[0].text.strip()
        batch_reviews = parse_generated_reviews(raw_text)
        if batch_reviews:
            all_reviews.extend(batch_reviews)
        if progress_callback:
            progress_callback(batch_idx + 1, len(batches), len(all_reviews))
    return all_reviews

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

def parse_match_response(text):
    mc, mb = "미등록", "미등록"
    for line in text.split('\n'):
        line = line.strip()
        if re.search(r'상품코드\s*(?:표)?\s*:', line):
            val = re.split(r'상품코드\s*(?:표)?\s*:', line, maxsplit=1)[1]
            val = re.sub(r'[*_`]', '', val).strip()
            if val and val != '미등록':
                mc = val
            elif val == '미등록':
                mc = '미등록'
        if re.search(r'브랜드\s*:', line):
            val = re.split(r'브랜드\s*:', line, maxsplit=1)[1]
            val = re.sub(r'[*_`]', '', val).strip()
            if val and val != '미등록':
                mb = val
            elif val == '미등록':
                mb = '미등록'
    return mc, mb


# ==================== 함소아 보고서 관련 ====================

HAMSOA_SHEET_URL = "https://docs.google.com/spreadsheets/d/1yozxvC3iXhCkbC3yXf5ad6PHaaZC3yQEvEXhpRcnGwc/"

def get_hamsoa_sheet(worksheet_name):
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        try:
            creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        except:
            BASE_DIR = os.path.dirname(os.path.abspath(__file__))
            creds = ServiceAccountCredentials.from_json_keyfile_name(
                os.path.join(BASE_DIR, 'service_account.json'), scope)
        client_gs = gspread.authorize(creds)
        return client_gs.open_by_url(HAMSOA_SHEET_URL).worksheet(worksheet_name)
    except Exception as e:
        st.error(f"함소아 시트 연결 실패 ({worksheet_name}): {e}")
        return None

def find_current_block_start(rows, no_col=0):
    """A열 번호가 1로 재시작되는 마지막 위치 반환"""
    last_start = 0
    prev_num = 0
    for i, row in enumerate(rows):
        if len(row) > no_col:
            cell = str(row[no_col]).strip()
            try:
                num = int(cell)
                if num == 1 and prev_num >= 5:
                    last_start = i
                prev_num = num
            except:
                pass
    return last_start

def _get_field(rec, keys):
    for k in keys:
        if k in rec and str(rec[k]).strip():
            return str(rec[k]).strip()
    return ''

def parse_competitor_sheet(sheet):
    all_values = sheet.get_all_values()
    if not all_values:
        return [], {}

    # 헤더 행 찾기
    header_row_idx = 0
    for i, row in enumerate(all_values):
        joined = ' '.join(str(c) for c in row)
        if ('경쟁사' in joined or 'NO' in joined.upper()) and '매체사' in joined and '발행' in joined:
            header_row_idx = i
            break

    headers = [str(h).strip() for h in all_values[header_row_idx]]
    data_rows = all_values[header_row_idx + 1:]

    block_start = find_current_block_start(data_rows, 0)
    current_rows = data_rows[block_start:]

    records = []
    for row in current_rows:
        if len(row) > 1 and str(row[0]).strip().isdigit():
            record = {}
            for j, h in enumerate(headers):
                record[h] = row[j].strip() if j < len(row) else ''
            records.append(record)

    # 병원별 통계
    hospital_counts = {}
    strategy_by_hospital = {}
    for rec in records:
        comp = _get_field(rec, ['경쟁사', '병원', '업체'])
        if not comp:
            continue
        hospital_counts[comp] = hospital_counts.get(comp, 0) + 1
        strategy = _get_field(rec, ['전략유형', '전략 유형', '전략'])
        if comp not in strategy_by_hospital:
            strategy_by_hospital[comp] = {}
        if strategy:
            strategy_by_hospital[comp][strategy] = strategy_by_hospital[comp].get(strategy, 0) + 1

    return records, {
        'hospital_counts': hospital_counts,
        'strategy_by_hospital': strategy_by_hospital,
        'headers': headers,
    }

def parse_hamsoa_sheet(sheet):
    all_values = sheet.get_all_values()
    if not all_values:
        return [], []

    articles = []
    billing = []

    article_header_idx = -1
    billing_header_idx = -1

    for i, row in enumerate(all_values):
        joined = ' '.join(str(c) for c in row)
        if '발행일' in joined and '구분' in joined and '제목' in joined and article_header_idx < 0:
            article_header_idx = i
        if '매체사' in joined and '분류' in joined and ('건당' in joined or '견적' in joined) and billing_header_idx < 0:
            billing_header_idx = i

    if article_header_idx >= 0:
        art_headers = [str(h).strip() for h in all_values[article_header_idx]]
        end_idx = billing_header_idx if 0 < billing_header_idx > article_header_idx else len(all_values)
        for i, row in enumerate(all_values[article_header_idx + 1:end_idx], start=article_header_idx + 1):
            if len(row) > 1 and str(row[0]).strip() and str(row[0]).strip() not in ['발행일', '']:
                record = {}
                for j, h in enumerate(art_headers):
                    record[h] = row[j].strip() if j < len(row) else ''
                articles.append(record)

    if billing_header_idx >= 0:
        bill_headers = [str(h).strip() for h in all_values[billing_header_idx]]
        for row in all_values[billing_header_idx + 1:]:
            if len(row) > 1 and str(row[0]).strip():
                cell0 = str(row[0]).strip().upper()
                if 'TOTAL' in cell0 or 'TOAL' in cell0:
                    continue
                record = {}
                has_data = False
                for j, h in enumerate(bill_headers):
                    val = row[j].strip() if j < len(row) else ''
                    record[h] = val
                    if val and j > 0:
                        has_data = True
                if has_data:
                    billing.append(record)

    return articles, billing

def make_bar_chart(hospital_counts, hamsoa_count):
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt

    try:
        plt.rcParams['font.family'] = 'Malgun Gothic'
    except:
        try:
            plt.rcParams['font.family'] = 'NanumGothic'
        except:
            pass
    plt.rcParams['axes.unicode_minus'] = False

    color_map = {
        '함소아 한의원': '#4472C4',
        '자생한방병원': '#FF4444',
        '폴리한의원': '#FFC000',
        '아이누리한의원': '#70AD47',
        '꽃피는 한의원': '#ED7D31',
        '해아림한의원': '#4BACC6',
        '헤아림한의원': '#4BACC6',
    }

    all_data = {'함소아 한의원': hamsoa_count}
    all_data.update(hospital_counts)

    sorted_names = ['함소아 한의원'] + sorted(
        [h for h in hospital_counts], key=lambda x: hospital_counts.get(x, 0), reverse=True)
    counts = [all_data.get(h, 0) for h in sorted_names]
    colors = [color_map.get(h, '#888888') for h in sorted_names]

    fig, ax = plt.subplots(figsize=(11, 5))
    bars = ax.bar(range(len(sorted_names)), counts, color=colors, width=0.5, edgecolor='white')

    ax.set_xticks(range(len(sorted_names)))
    ax.set_xticklabels(sorted_names, fontsize=10)
    ax.set_title("각 병원별 보도자료 배포 수량", fontsize=13, fontweight='bold', pad=15)
    ax.yaxis.grid(True, linestyle='--', alpha=0.5)
    ax.set_axisbelow(True)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    # 범례
    legend_patches = [plt.Rectangle((0, 0), 1, 1, color=color_map.get(n, '#888888')) for n in sorted_names]
    ax.legend(legend_patches, sorted_names, loc='upper right', fontsize=8, ncol=2)

    for bar, count in zip(bars, counts):
        if count > 0:
            ax.text(bar.get_x() + bar.get_width() / 2., bar.get_height() + 0.3,
                    str(count), ha='center', va='bottom', fontsize=10, fontweight='bold')

    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    buf.seek(0)
    plt.close()
    return buf

def make_strategy_chart(strategy_by_hospital):
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt

    try:
        plt.rcParams['font.family'] = 'Malgun Gothic'
    except:
        try:
            plt.rcParams['font.family'] = 'NanumGothic'
        except:
            pass
    plt.rcParams['axes.unicode_minus'] = False

    if not strategy_by_hospital:
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.text(0.5, 0.5, '데이터 없음', ha='center', va='center', fontsize=14)
        ax.axis('off')
        buf = io.BytesIO()
        plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
        buf.seek(0)
        plt.close()
        return buf

    all_strategies = set()
    for strats in strategy_by_hospital.values():
        all_strategies.update(strats.keys())
    strategy_list = sorted(all_strategies)

    hospitals = list(strategy_by_hospital.keys())
    x = list(range(len(hospitals)))
    width = 0.7 / max(len(strategy_list), 1)

    strat_colors = {
        '브랜드 강화형': '#4472C4',
        '질환 타깃형': '#FF4444',
        '시술 중심형': '#FFC000',
        '마케팅 형': '#70AD47',
        '마케팅형': '#70AD47',
    }

    fig, ax = plt.subplots(figsize=(11, 5))
    for i, strategy in enumerate(strategy_list):
        counts = [strategy_by_hospital[h].get(strategy, 0) for h in hospitals]
        offset = (i - len(strategy_list) / 2 + 0.5) * width
        color = strat_colors.get(strategy, f'C{i}')
        ax.bar([xi + offset for xi in x], counts, width=width * 0.9, label=strategy, color=color)

    ax.set_xticks(x)
    ax.set_xticklabels(hospitals, fontsize=9)
    ax.set_title("경쟁사 전략유형 별 집계 현황", fontsize=13, fontweight='bold', pad=15)
    ax.legend(loc='upper right', fontsize=9)
    ax.yaxis.grid(True, linestyle='--', alpha=0.5)
    ax.set_axisbelow(True)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    buf.seek(0)
    plt.close()
    return buf

def _set_cell_bg(cell, hex_color):
    from pptx.oxml.ns import qn
    from lxml import etree
    tc = cell._tc
    tcPr = tc.find(qn('a:tcPr'))
    if tcPr is None:
        tcPr = etree.SubElement(tc, qn('a:tcPr'))
    for child in list(tcPr):
        tag = child.tag
        if 'Fill' in tag or 'fill' in tag:
            tcPr.remove(child)
    solidFill = etree.SubElement(tcPr, qn('a:solidFill'))
    srgbClr = etree.SubElement(solidFill, qn('a:srgbClr'))
    srgbClr.set('val', hex_color.replace('#', '').upper())
    tcPr.insert(0, solidFill)

def _style_cell(cell, text, bg_hex=None, fg_hex=None, font_size=8,
                bold=False, center=False, italic=False):
    from pptx.util import Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor

    cell.text = ""
    tf = cell.text_frame
    tf.word_wrap = True
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.CENTER if center else PP_ALIGN.LEFT

    run = para.add_run()
    run.text = str(text)
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    if fg_hex:
        h = fg_hex.replace('#', '')
        run.font.color.rgb = RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    if bg_hex:
        _set_cell_bg(cell, bg_hex)

def parse_jasaeng_strategy(raw_text):
    """자생한방병원 전략 데이터 파싱
    형식:
        질환 타겟형
        비염 1
        일자목 1
        시술 중심형
        ...
    반환: strategies dict, order list, totals dict, grand_total int
    """
    strategies = {}
    strategy_order = []
    current_strategy = None

    for line in raw_text.strip().split('\n'):
        line = line.strip()
        if not line:
            continue
        # 줄 끝 숫자 추출 (공백 유무 상관없이: "잔기침1" or "비염 1" 모두 처리)
        m = re.match(r'^(.+?)\s*(\d+)\s*$', line)
        if m:
            keyword, count = m.group(1).strip(), int(m.group(2))
            if current_strategy is not None:
                strategies[current_strategy].append((keyword, count))
            continue
        # 전략유형 헤더 (숫자로 끝나지 않는 줄)
        current_strategy = line
        if current_strategy not in strategies:
            strategies[current_strategy] = []
            strategy_order.append(current_strategy)

    totals = {s: sum(c for _, c in items) for s, items in strategies.items()}
    grand_total = sum(totals.values())
    return strategies, strategy_order, totals, grand_total


def generate_quantity_analysis(hospital_counts, hamsoa_count, jasaeng_total, report_date):
    """병원별 발행 수량 비교 분석 텍스트 생성 (이미지 1 형식)
    반환: list of str (각 불릿 텍스트)
    """
    others = {h: c for h, c in hospital_counts.items()
              if '자생' not in h and '함소아' not in h}
    combined = list(others.items())
    combined.append(('함소아한의원', hamsoa_count))
    combined = sorted(combined, key=lambda x: x[1], reverse=True)

    positive = [(h, c) for h, c in combined if c > 0]
    zero_hosps = [h for h, c in combined if c == 0]

    texts = []

    # 불릿 1: 자생 1위
    texts.append(
        f"{report_date} 기준, 자생한방병원은 총 {jasaeng_total}건의 보도자료를 언론에 배포하며, "
        f"전체 병원 중 가장 많은 기사 발행 수를 기록했습니다."
    )

    # 불릿 2: 나머지 병원들 순서대로
    if positive:
        first_h, first_c = positive[0]
        parts2 = [f"{first_h}은 총 {first_c}건을 발행하며 그 뒤를 이었고"]
        for h, c in positive[1:]:
            parts2.append(f"{h}이 총 {c}건 발행 하였으며")
        t2 = ", ".join(parts2)
        if zero_hosps:
            t2 += f", {', '.join(zero_hosps)}은 발행 수량이 0건으로 나타났습니다."
        else:
            t2 += "."
        texts.append(t2)

    # 불릿 3: 자생 vs 함소아 비율
    if jasaeng_total > 0 and hamsoa_count > 0:
        ratio = round(hamsoa_count / jasaeng_total * 100, 1)
        times = round(jasaeng_total / hamsoa_count)
        texts.append(
            f"자생한방병원은 총 {jasaeng_total}건, 함소아한의원은 {hamsoa_count}건의 언론보도를 발행하였으며, "
            f"함소아한의원의 발행 수는 자생한방병원의 약 {ratio}% 수준으로, "
            f"양 기관 간 보도자료 운영 규모에서 {times}배 이상의 격차를 보였습니다."
        )

    return texts


def generate_strategy_analysis(jasaeng_strategies, jasaeng_order, jasaeng_totals, jasaeng_grand_total,
                                competitor_records, hospital_counts):
    """전략유형별 세부 분석 텍스트 생성 (이미지 2 형식)
    반환: list of str (각 불릿 텍스트)
    """
    texts = []

    # 자생한방병원 전략 breakdown
    strat_parts = []
    for s in jasaeng_order:
        t = jasaeng_totals.get(s, 0)
        strat_parts.append(f"'{s}'이 {t}건")
    jasaeng_strat_str = ", ".join(strat_parts)

    texts.append(
        f"총 {jasaeng_grand_total}건의 자생한방병원 기사 중, {jasaeng_strat_str}으로 확인되며, "
        f"질환 관련 정보성 기사와 브랜드 이미지 제고를 중심으로 전략적 콘텐츠 배포를 함께 다뤄진 것으로 분석됩니다. "
        f"특히 이슈가 발생한 키워드를 중점으로 칼럼 기사를 배포하여 최대한 많은 사람들에게 기사를 "
        f"노출시킬 수 있도록 집중한 것으로 확인됩니다."
    )

    # 경쟁사 스프레드시트 병원별 전략유형 집계
    hosp_strategies = {}
    for rec in competitor_records:
        h = _get_field(rec, ['병원', '경쟁사', 'hospital', '병 원'])
        s = _get_field(rec, ['전략유형', '전략 유형', 'strategy', '유형'])
        if h and s:
            hosp_strategies.setdefault(h, {})
            hosp_strategies[h][s] = hosp_strategies[h].get(s, 0) + 1

    sorted_hosps = sorted(hospital_counts.items(), key=lambda x: x[1], reverse=True)
    active = [(h, c) for h, c in sorted_hosps if c > 0 and '자생' not in h]
    inactive = [h for h, c in sorted_hosps if c == 0 and '자생' not in h]

    for h, c in active:
        strat = hosp_strategies.get(h, {})
        if strat:
            sorted_strat = sorted(strat.items(), key=lambda x: x[1], reverse=True)
            strat_desc = ", ".join(f"{s}이 {n}건" for s, n in sorted_strat)
            main_s = sorted_strat[0][0]
            texts.append(
                f"{h}은 총 {c}건의 기사 중 {strat_desc}으로 구성되어, "
                f"'{main_s}' 전략을 중심으로 운영한 것으로 분석됩니다."
            )
        else:
            texts.append(f"{h}은 총 {c}건의 기사를 발행하였습니다.")

    if inactive:
        inactive_str = ", ".join(inactive)
        texts.append(
            f"반면, {inactive_str}은 집계된 기사가 없는 상태로, 언론 노출 자체가 미비합니다."
        )

    return texts


def generate_hamsoa_ppt(competitor_records, meta, hamsoa_articles, billing_data,
                         report_date, report_month, hamsoa_article_count):
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN

    NAVY_RGB = RGBColor(0x1A, 0x2B, 0x4A)
    WHITE_RGB = RGBColor(0xFF, 0xFF, 0xFF)
    GRAY_RGB = RGBColor(0xAA, 0xAA, 0xAA)
    NAVY_HEX = "1A2B4A"
    BEIGE_HEX = "F5E3BA"
    WHITE_HEX = "FFFFFF"
    ALT_HEX = "F2F2F2"

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    def new_slide():
        return prs.slides.add_slide(blank)

    def set_bg(slide, hex_color):
        bg = slide.background
        fill = bg.fill
        fill.solid()
        h = hex_color.replace('#', '')
        fill.fore_color.rgb = RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

    def txt(slide, text, left, top, width, height,
            size=11, bold=False, rgb=None, align=PP_ALIGN.LEFT, italic=False):
        box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
        tf = box.text_frame
        tf.word_wrap = True
        para = tf.paragraphs[0]
        para.alignment = align
        run = para.add_run()
        run.text = str(text)
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        if rgb:
            run.font.color.rgb = rgb
        return box

    def aligo_mark(slide):
        txt(slide, "ALIGO", 12.1, 0.1, 1.1, 0.28, size=8, rgb=GRAY_RGB, align=PP_ALIGN.RIGHT)
        txt(slide, "MEDIA", 12.1, 0.35, 1.1, 0.28, size=8, rgb=GRAY_RGB, align=PP_ALIGN.RIGHT)

    def section_title(slide, title_text):
        txt(slide, title_text, 0.5, 0.35, 12.5, 0.55, size=14, bold=True, rgb=NAVY_RGB)
        line = slide.shapes.add_connector(1, Inches(0.5), Inches(0.95), Inches(12.8), Inches(0.95))
        line.line.color.rgb = NAVY_RGB
        line.line.width = Pt(1.2)

    def bullet_box(slide, bullets, left, top, width, height):
        box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(0xEF, 0xF6, 0xFF)
        box.line.color.rgb = RGBColor(0xC5, 0xD8, 0xF0)
        box.line.width = Pt(0.75)
        tf = box.text_frame
        tf.word_wrap = True
        for i, bullet_text in enumerate(bullets):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.space_before = Pt(5)
            run = para.add_run()
            run.text = f"➡  {bullet_text}"
            run.font.size = Pt(9)
            run.font.color.rgb = NAVY_RGB

    def add_pic(slide, img_buf, left, top, width, height):
        img_buf.seek(0)
        slide.shapes.add_picture(img_buf, Inches(left), Inches(top), Inches(width), Inches(height))

    def add_table(slide, data, col_widths, left, top, max_height=5.8):
        rows_n = len(data)
        cols_n = len(data[0]) if data else 1
        row_h = min(max_height / rows_n, 0.55)
        tbl_h = row_h * rows_n
        tbl = slide.shapes.add_table(
            rows_n, cols_n,
            Inches(left), Inches(top),
            Inches(sum(col_widths)), Inches(tbl_h)
        ).table
        for ci, cw in enumerate(col_widths):
            tbl.columns[ci].width = Inches(cw)
        for ri, row_data in enumerate(data):
            for ci, val in enumerate(row_data):
                cell = tbl.cell(ri, ci)
                is_header = ri == 0
                is_total = ri == rows_n - 1 and str(data[ri][0]).upper().startswith('TOTAL')
                is_alt = ri % 2 == 0 and not is_header

                if is_header:
                    bg, fg = BEIGE_HEX, NAVY_HEX
                elif is_total:
                    bg, fg = NAVY_HEX, WHITE_HEX
                elif is_alt:
                    bg, fg = ALT_HEX, "333333"
                else:
                    bg, fg = WHITE_HEX, "333333"

                # 예정 행 빨간 글씨
                if not is_header and '예정' in str(val):
                    fg = "CC0000"

                center_cols = {0, 1, 2, 3, 5, 6}  # 보통 날짜/숫자 컬럼
                _style_cell(cell, val, bg_hex=bg, fg_hex=fg,
                            font_size=9 if is_header else 8,
                            bold=is_header or is_total,
                            center=(ci in center_cols))
        return tbl

    # ── SLIDE 1: Cover ──────────────────────────────────────
    s1 = new_slide()
    set_bg(s1, NAVY_HEX)
    txt(s1, "ALIGO", 11.9, 0.18, 1.2, 0.3, size=9, rgb=GRAY_RGB, align=PP_ALIGN.RIGHT)
    txt(s1, "MEDIA", 11.9, 0.46, 1.2, 0.3, size=9, rgb=GRAY_RGB, align=PP_ALIGN.RIGHT)
    txt(s1, "함소아 한의원", 0.8, 1.8, 11.7, 2.0,
        size=58, bold=True, rgb=WHITE_RGB, align=PP_ALIGN.CENTER)
    txt(s1, "REPORT", 0.5, 4.9, 8.0, 1.6, size=62, bold=True, rgb=WHITE_RGB)
    txt(s1, report_date, 7.0, 5.8, 6.0, 0.55, size=14, bold=True, rgb=WHITE_RGB, align=PP_ALIGN.RIGHT)
    txt(s1, "종합 리포트", 7.0, 6.32, 6.0, 0.4, size=11, rgb=WHITE_RGB, align=PP_ALIGN.RIGHT)

    # ── SLIDE 2: Contents ────────────────────────────────────
    s2 = new_slide()
    set_bg(s2, "F0F0F0")
    aligo_mark(s2)
    txt(s2, f"{report_month} 종합 리포트", 0.7, 0.32, 8.0, 0.42, size=12, bold=True, rgb=NAVY_RGB)
    txt(s2, "CONTENTS", 0.7, 0.7, 11.5, 1.5, size=72, bold=True, rgb=NAVY_RGB)

    contents_items = [
        ("01", "언론보도 발행 수량 비교 (함소아 vs 경쟁사)"),
        ("02", "경쟁사 기사 발행 정보 상세 분석"),
        ("03", "병원별 기사 발행 현황 요약"),
        ("04", "함소아 기사 발행 현황 요약"),
        ("05", "브릿지경제 지면보도 현황"),
        ("06", "함소아 기사 집행 금액 정산표"),
    ]
    for i, (num, label) in enumerate(contents_items):
        col_idx = i // 3
        row_idx = i % 3
        bx = 0.7 + col_idx * 6.4
        by = 2.55 + row_idx * 1.25

        nb = s2.shapes.add_shape(1, Inches(bx), Inches(by), Inches(0.52), Inches(0.42))
        nb.fill.solid()
        nb.fill.fore_color.rgb = NAVY_RGB
        nb.line.fill.background()
        np_ = nb.text_frame.paragraphs[0]
        np_.alignment = PP_ALIGN.CENTER
        nr = np_.add_run()
        nr.text = num
        nr.font.size = Pt(12)
        nr.font.bold = True
        nr.font.color.rgb = WHITE_RGB

        txt(s2, label, bx + 0.62, by + 0.02, 5.6, 0.42, size=11, rgb=NAVY_RGB)

    # ── SLIDE 3: 언론보도 발행 수량 비교 ─────────────────────
    s3 = new_slide()
    aligo_mark(s3)
    section_title(s3, "1. 언론보도 발행 수량 비교 (함소아 vs 경쟁사)")

    hospital_counts = meta.get('hospital_counts', {})
    chart_buf = make_bar_chart(hospital_counts, hamsoa_article_count)
    add_pic(s3, chart_buf, 0.4, 1.05, 9.3, 4.6)

    total_comp = sum(hospital_counts.values())
    top_h = max(hospital_counts, key=hospital_counts.get) if hospital_counts else ''
    top_c = hospital_counts.get(top_h, 0)

    bullets_s3 = []
    if top_h:
        bullets_s3.append(f"{report_date} 기준, {top_h}은(는) 총 {top_c}건으로 전체 병원 중 가장 많은 기사 발행 수를 기록했습니다.")
    bullets_s3.append(f"함소아한의원은 총 {hamsoa_article_count}건의 언론보도를 발행하였습니다.")
    for h, c in sorted(hospital_counts.items(), key=lambda x: x[1], reverse=True)[:2]:
        if h != top_h:
            bullets_s3.append(f"{h}: {c}건 발행")
    bullet_box(s3, bullets_s3[:3], 9.85, 1.05, 3.2, 4.6)

    # ── SLIDES 4+: 경쟁사 기사 발행 정보 상세 분석 ──────────
    COMP_KEY_MAP = {
        '경쟁사': ['경쟁사', '병원'],
        '매체사': ['매체사', '매체'],
        '발행일': ['발행일자', '발행일', '날짜'],
        '주요키워드': ['메인키워드', '주요키워드', '키워드', '메인 키워드'],
        '제목': ['제목', '기사제목'],
        '월간 검색량': ['월간 검색량', '월간검색량', '검색량'],
    }
    comp_header_row = ['NO.', '경쟁사', '매체사', '발행일', '주요키워드', '제목', '월간 검색량']
    comp_col_w = [0.45, 1.2, 1.2, 0.95, 1.0, 5.6, 1.0]

    ROWS_PER = 10
    total_rec = len(competitor_records)
    n_comp_slides = max(1, -(-total_rec // ROWS_PER))

    for si in range(n_comp_slides):
        sld = new_slide()
        aligo_mark(sld)
        section_title(sld, f"2. 경쟁사 기사 발행 정보 상세 분석 ({si + 1})")

        batch = competitor_records[si * ROWS_PER: (si + 1) * ROWS_PER]
        table_data = [comp_header_row]
        for k, rec in enumerate(batch):
            table_data.append([
                str(si * ROWS_PER + k + 1),
                _get_field(rec, COMP_KEY_MAP['경쟁사']),
                _get_field(rec, COMP_KEY_MAP['매체사']),
                _get_field(rec, COMP_KEY_MAP['발행일']),
                _get_field(rec, COMP_KEY_MAP['주요키워드']),
                _get_field(rec, COMP_KEY_MAP['제목']),
                _get_field(rec, COMP_KEY_MAP['월간 검색량']),
            ])
        add_table(sld, table_data, comp_col_w, left=0.35, top=1.1)

    # 경쟁사 분석 텍스트 슬라이드
    s_comp_txt = new_slide()
    aligo_mark(s_comp_txt)
    section_title(s_comp_txt, f"2. 경쟁사 기사 발행 정보 상세 분석 ({n_comp_slides + 1})")

    top_hospitals = sorted(hospital_counts.items(), key=lambda x: x[1], reverse=True)
    analysis_bullets = [f"총 {total_comp}건의 경쟁사 기사 중, {top_h}이(가) 총 {top_c}건으로 가장 높은 기사 발행 빈도를 보였습니다." if top_h else "데이터를 분석하고 있습니다."]
    for h, c in top_hospitals[:4]:
        if c > 0:
            strats = meta.get('strategy_by_hospital', {}).get(h, {})
            top_s = max(strats, key=strats.get) if strats else ''
            analysis_bullets.append(f"{h}: {c}건 ({top_s} 중심)" if top_s else f"{h}: {c}건 발행")
    bullet_box(s_comp_txt, analysis_bullets[:5], 0.5, 1.2, 12.3, 3.5)

    # ── 병원별 기사 발행 현황 ────────────────────────────────
    s_hosp = new_slide()
    aligo_mark(s_hosp)
    section_title(s_hosp, "3. 병원별 기사 발행 현황 요약")

    strat_chart_buf = make_strategy_chart(meta.get('strategy_by_hospital', {}))
    add_pic(s_hosp, strat_chart_buf, 0.4, 1.05, 9.5, 5.0)

    hosp_bullets = []
    for h, strats in sorted(meta.get('strategy_by_hospital', {}).items(),
                             key=lambda x: sum(x[1].values()), reverse=True)[:4]:
        total = sum(strats.values())
        top_s = max(strats, key=strats.get) if strats else ''
        hosp_bullets.append(f"{h}: 총 {total}건 ({top_s} {strats.get(top_s, 0)}건)")
    if hosp_bullets:
        bullet_box(s_hosp, hosp_bullets, 10.05, 1.05, 3.1, 5.0)

    # ── 함소아 기사 발행 현황 ────────────────────────────────
    ART_KEYS = ['발행일', '구분', '제목', '매체사', '메인 키워드', '검색량', '진행 현황', '본문 요약']
    ART_KEY_ALT = ['발행일', '구분', '제목', '매체사', '메인키워드', '검색량', '진행현황', '본문요약']
    ART_HDR = ['발행일', '구분', '제목', '매체사', '메인 키워드', '검색량', '진행현황', '본문 요약']
    ART_COL_W = [0.95, 0.75, 2.5, 1.05, 1.05, 0.65, 0.85, 4.7]

    ART_ROWS = 7
    total_arts = len(hamsoa_articles)
    n_art_slides = max(1, -(-total_arts // ART_ROWS))

    for si in range(n_art_slides):
        sld = new_slide()
        aligo_mark(sld)
        section_title(sld, f"4. 함소아 기사 발행 현황 요약 ({si + 1})")

        batch = hamsoa_articles[si * ART_ROWS: (si + 1) * ART_ROWS]
        table_data = [ART_HDR]
        for art in batch:
            row = []
            for k, k_alt in zip(ART_KEYS, ART_KEY_ALT):
                val = art.get(k, art.get(k_alt, art.get(k.replace(' ', ''), '')))
                row.append(val)
            table_data.append(row)
        add_table(sld, table_data, ART_COL_W, left=0.35, top=1.1)

    # 함소아 기사 분석 텍스트
    s_art_txt = new_slide()
    aligo_mark(s_art_txt)
    section_title(s_art_txt, f"4. 함소아 기사 발행 현황 요약 ({n_art_slides + 1})")

    completed = [a for a in hamsoa_articles
                 if '완료' in str(a.get('진행 현황', a.get('진행현황', '')))]
    planned = [a for a in hamsoa_articles
               if '예정' in str(a.get('진행 현황', a.get('진행현황', '')))]

    art_bullets = [
        f"{report_date} 기준, 총 {total_arts}건의 기사 중 {len(completed)}건은 발행 완료입니다.(기획기사 및 명의칼럼 포함)",
    ]
    if planned:
        art_bullets.append(f"예정된 기사: {len(planned)}건")

    keyword_counter = {}
    for a in hamsoa_articles:
        kw = a.get('메인 키워드', a.get('메인키워드', ''))
        if kw:
            keyword_counter[kw] = keyword_counter.get(kw, 0) + 1

    bullet_box(s_art_txt, art_bullets, 0.5, 1.2, 12.3, 3.5)

    # ── 브릿지경제 지면보도 ──────────────────────────────────
    s_bridge = new_slide()
    aligo_mark(s_bridge)
    section_title(s_bridge, "5. 함소아 브릿지경제 지면보도 현황")

    bridge_arts = [a for a in hamsoa_articles
                   if '브릿지' in str(a.get('매체사', ''))]

    bridge_data = [['NO.', '게재일', '키워드', '기사 제목']]
    for i, art in enumerate(bridge_arts):
        bridge_data.append([
            str(i + 1),
            art.get('발행일', ''),
            art.get('메인 키워드', art.get('메인키워드', '')),
            art.get('제목', ''),
        ])

    if len(bridge_data) > 1:
        add_table(s_bridge, bridge_data, [0.5, 1.3, 1.5, 9.2], left=0.35, top=1.1, max_height=4.0)
    else:
        txt(s_bridge, "브릿지경제 지면보도 데이터가 없습니다. (매체사 열에 '브릿지경제'로 입력된 항목이 있으면 자동 집계됩니다.)",
            0.5, 2.0, 12.0, 1.0, size=11, rgb=NAVY_RGB)

    # ── 함소아 집행 금액 정산표 ──────────────────────────────
    s_bill = new_slide()
    aligo_mark(s_bill)
    section_title(s_bill, "6. 함소아 기사 집행 금액 정산표")

    if billing_data:
        bill_hdr = ['매체사', '분류', '개재 건수', '건당 견적', '매체별 합산']
        bill_table_data = [bill_hdr]
        total_amount = 0

        for rec in billing_data:
            vals = list(rec.values())
            row = (vals + [''] * 5)[:5]
            bill_table_data.append(row)
            try:
                amt_str = str(vals[4] if len(vals) > 4 else (vals[-1] if vals else '')).replace(',', '').replace('원', '').strip()
                if amt_str.isdigit():
                    total_amount += int(amt_str)
            except:
                pass

        bill_table_data.append(['TOTAL', f"기획기사 {len(billing_data) - 1}건 + 보고서 집계",
                                 f"{len(billing_data)}건", '', f"{total_amount:,}원" if total_amount else ''])

        add_table(s_bill, bill_table_data, [2.8, 2.5, 1.3, 1.3, 1.6], left=1.8, top=1.1, max_height=5.5)

        txt(s_bill, "※ 해당 금액은 VAT 별도 기준입니다",
            1.8, 1.1 + 0.55 * len(bill_table_data) + 0.15, 9.5, 0.35,
            size=9, italic=True, rgb=RGBColor(0x88, 0x88, 0x88))
    else:
        txt(s_bill, "정산 데이터가 없습니다.",
            0.5, 2.0, 12.0, 0.5, size=11, rgb=NAVY_RGB)

    # 저장
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ==================== Streamlit 앱 ====================

st.set_page_config(page_title="버즈필터 자동화", page_icon="🤖", layout="wide")

with st.sidebar:
    st.markdown("## 📋 메뉴")
    st.markdown("---")
    menu = st.radio("", options=[
        "🏭 버즈필터 발주",
        "✍️ 리뷰 생성",
        "📝 리뷰 입력",
        "📄 견적서 생성",
        "🌐 홈페이지 자동 개선",
        "📊 함소아 보고서",
    ], label_visibility="collapsed")
    st.markdown("---")
    st.caption("버즈필터 업무 자동화 시스템")
    st.caption("© 2025 알리고미디어")

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
                product_col = '상품코드 표' if '상품코드 표' in calc_df.columns else '상품코드'
                display_df = calc_df[['브랜드', '제품명', product_col]].copy()
                display_df = display_df.rename(columns={product_col: '상품코드'})
                for idx, row in df.iterrows():
                    raw = str(row.get('상품명+옵션+개수', ''))
                    qty = extract_qty_from_text(raw)
                    price = int(row.get('가격', 0))
                    ch = str(row.get('판매처', '쿠팡'))
                    prompt = f"""너는 상품 매칭 전문가야.
발주서 상품명: {raw}
상품 리스트 (브랜드 / 제품명 / 상품코드):
{display_df.to_string(index=False)}
규칙:
1. 브랜드명이나 핵심 키워드가 겹치면 매칭해라
2. 가장 유사한 것 1개만 선택해라
3. 마크다운 절대 사용 금지
4. 정말 모르겠으면 미등록
반드시 아래 형식으로만 답해줘:
상품코드: [값]
브랜드: [값]"""
                    resp = client.messages.create(model="claude-haiku-4-5-20251001", max_tokens=150,
                                                  messages=[{"role": "user", "content": prompt}])
                    rt = resp.content[0].text.strip()
                    mc, mb = parse_match_response(rt)
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
        product_info = st.text_area("제품 정보 (제품명, 카테고리, 특징 등)", height=150, placeholder="예)\n제품명: 콜라겐 마스크팩\n카테고리: 스킨케어\n특징: 저자극 성분, 수분 집중 케어")
        selling_points = st.text_area("소구점 / 강조할 내용 (선택)", height=100, placeholder="예) 피부 흡수력, 아침에 쓰기 좋음, 가성비")
    with col_right:
        product_images = st.file_uploader("제품 이미지 (선택, 여러 장 가능)", type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True)
        if product_images:
            for img in product_images:
                st.image(img, caption=img.name, use_container_width=True)
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
                image_data_list = []
                media_type_map = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png", "webp": "image/webp"}
                if product_images:
                    for product_image in product_images:
                        product_image.seek(0)
                        img_bytes = product_image.read()
                        img_b64 = base64.b64encode(img_bytes).decode('utf-8')
                        ext = product_image.name.split('.')[-1].lower()
                        image_data_list.append({"media_type": media_type_map.get(ext, "image/jpeg"), "data": img_b64})
                progress_bar = st.progress(0)
                status_text = st.empty()
                def update_progress(current_batch, total_b, total_generated):
                    pct = int((current_batch / total_b) * 100)
                    progress_bar.progress(pct)
                    status_text.text(f"⏳ 배치 {current_batch}/{total_b} 완료 — 현재까지 {total_generated}개 생성됨")
                status_text.text(f"🚀 리뷰 생성 시작 (총 {total_batches}번 배치 호출)...")
                parsed = generate_reviews_with_claude(
                    client=ai_client, product_info=product_info, selling_points=selling_points,
                    review_count=review_count, char_count=char_count,
                    image_data_list=image_data_list, progress_callback=update_progress)
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
        updated_reviews = []
        for i, (num, content) in enumerate(st.session_state.generated_reviews):
            with st.expander(f"리뷰 {num}번", expanded=(i < 3)):
                edited = st.text_area(f"리뷰 {num} 내용", value=content, height=150, key=f"review_edit_{i}", label_visibility="collapsed")
                updated_reviews.append((num, edited))
        st.markdown("---")
        st.markdown("### 💾 STEP 4 — 저장 및 다운로드")
        col_save, col_reset = st.columns([3, 1])
        with col_save:
            if st.button("⬇️ 저장 및 엑셀 다운로드", type="primary", use_container_width=True):
                st.session_state.generated_reviews = updated_reviews
                excel_data = create_excel(updated_reviews)
                fname = f"리뷰_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                st.download_button(label="📥 엑셀 파일 다운로드 클릭", data=excel_data, file_name=fname,
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True, type="primary")
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

elif menu == "📝 리뷰 입력":
    st.title("📝 리뷰 엑셀 자동 변환기")
    st.subheader("리뷰 텍스트 파일을 업로드하면 엑셀 파일로 자동 변환합니다.")
    tab1, tab2 = st.tabs(["📁 파일 업로드", "✏️ 텍스트 직접 입력"])
    with tab1:
        utxt = st.file_uploader("리뷰 텍스트 파일 (.txt)", type=["txt"])
        if utxt:
            txt_content = utxt.read().decode("utf-8", errors="ignore")
            st.success(f"✅ {utxt.name} 업로드 완료")
            revs = parse_reviews(txt_content)
            if revs:
                st.markdown(f"### 📊 **{len(revs)}개** 리뷰 감지됨")
                with st.expander("👀 미리보기", expanded=True):
                    for num, content in revs[:5]:
                        st.markdown(f"**{num}번 리뷰**"); st.text(content[:200] + ("..." if len(content) > 200 else "")); st.divider()
                    if len(revs) > 5: st.info(f"... 외 {len(revs) - 5}개")
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
                    for num, content in revs[:5]:
                        st.markdown(f"**{num}번 리뷰**"); st.text(content[:200] + ("..." if len(content) > 200 else "")); st.divider()
                st.download_button("⬇️ 엑셀 다운로드", create_excel(revs), "리뷰목록.xlsx", use_container_width=True, type="primary")
            else:
                st.error("❌ 리뷰를 파싱할 수 없습니다.")

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
    if 'quote_items' not in st.session_state:
        st.session_state.quote_items = [{"품목": "", "구성": "", "수량": 1, "단가": 0, "비고": ""}]
    hcols = st.columns([3, 2, 1, 2, 2, 1])
    for col, lbl in zip(hcols, ["**품목**", "**구성**", "**수량**", "**단가(원)**", "**공급가액**", "**삭제**"]):
        col.markdown(lbl)
    to_del = []
    for i, item in enumerate(st.session_state.quote_items):
        cols = st.columns([3, 2, 1, 2, 2, 1])
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
        st.session_state.quote_items.append({"품목": "", "구성": "", "수량": 1, "단가": 0, "비고": ""})
        st.rerun()
    st.markdown("---")
    valid = [it for it in st.session_state.quote_items if it["품목"].strip()]
    sup = sum(it["수량"] * it["단가"] for it in valid)
    vat = int(sup * 0.1) if tax_type == "발행" else 0
    tot = sup + vat
    mc1, mc2, mc3 = st.columns(3)
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
                qd = {"date": quote_date, "client": client_name, "tax_type": tax_type, "memo": memo, "items": valid}
                try:
                    pdf_buf = generate_quote_pdf(qd, stamp_path)
                    fname = f"견적서_{client_name}_{quote_date.replace('. ', '').replace('.', '')}.pdf"
                    st.success("✅ 견적서 PDF 생성 완료!")
                    st.download_button("⬇️ PDF 다운로드", pdf_buf, fname, mime="application/pdf", use_container_width=True, type="primary")
                except Exception as e:
                    st.error(f"❌ PDF 생성 실패: {e}")

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
        st.error("❌ Netlify 미연결 — secrets.toml에 NETLIFY_TOKEN과 NETLIFY_SITE_ID를 추가해주세요.")
    st.markdown("---")
    st.markdown("### 📂 STEP 1 — index.html 업로드")
    uploaded_html = st.file_uploader("index.html 업로드", type=["html", "htm"], key="html_upload")
    st.markdown("### 🖼️ STEP 2 — 이미지 파일 업로드")
    uploaded_images = st.file_uploader("이미지 파일 (여러 개 동시 선택 가능)", type=["png", "jpg", "jpeg", "gif", "webp", "ico"], accept_multiple_files=True, key="image_upload")
    if uploaded_images:
        st.success(f"✅ 이미지 {len(uploaded_images)}개: {', '.join([f.name for f in uploaded_images])}")
    st.markdown("---")
    st.markdown("### ✅ STEP 3 — 개선 항목 선택")
    col1, col2, col3 = st.columns(3)
    with col1: check_mobile = st.checkbox("📱 모바일 최적화", value=True)
    with col2: check_responsive = st.checkbox("📐 반응형 디자인", value=True)
    with col3: check_seo = st.checkbox("🔍 구글 SEO", value=True)
    check_extra = st.text_area("📝 추가 요청사항 (선택)", placeholder="예) 버튼 색상을 더 눈에 띄게", height=80)
    st.markdown("---")
    if uploaded_html:
        html_content = uploaded_html.read().decode("utf-8", errors="ignore")
        if not any([check_mobile, check_responsive, check_seo, check_extra.strip()]):
            st.warning("⚠️ 개선 항목을 최소 1개 이상 선택해주세요.")
        else:
            if st.button("🚀 Claude 수정 + Netlify 자동 배포", type="primary", use_container_width=True):
                check_list = []
                if check_mobile: check_list.append("1. 모바일 최적화")
                if check_responsive: check_list.append("2. 반응형 디자인")
                if check_seo: check_list.append("3. 구글 SEO")
                if check_extra.strip(): check_list.append(f"4. 추가 요청: {check_extra.strip()}")
                prompt = f"""너는 웹 개발 전문가야. 아래 HTML을 분석하고 수정해서 완성된 HTML 코드만 반환해줘.
[개선 항목]
{chr(10).join(check_list)}
[주의사항]
- 수정된 HTML 전체 코드만 반환 (설명 없이 <!DOCTYPE html>부터 시작)
- 기존 디자인·색상·브랜드 정체성 유지
- 한국어 텍스트 수정 금지
[원본 HTML]
{html_content}"""
                progress_bar = st.progress(0)
                status_text = st.empty()
                try:
                    status_text.text("🤖 Claude가 분석 및 수정 중...")
                    progress_bar.progress(20)
                    ai_client = get_anthropic_client()
                    response = ai_client.messages.create(model="claude-opus-4-5", max_tokens=16000, messages=[{"role": "user", "content": prompt}])
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
            st.metric("파일 크기", f"{len(improved_html):,}자", delta=f"{len(improved_html) - len(original_html):+,}자")
            with st.expander("수정된 코드 보기"):
                st.code(improved_html[:1500] + "...", language="html")
        st.markdown("---")
        st.download_button(label="⬇️ 수정된 index.html 다운로드", data=improved_html.encode("utf-8"), file_name="index.html", mime="text/html", use_container_width=True)
        if st.button("🔄 처음부터 다시", use_container_width=True):
            st.session_state["improvement_done"] = False
            st.session_state["improved_html"] = ""
            st.rerun()

elif menu == "📊 함소아 보고서":
    st.title("📊 함소아한의원 보고서 분석 텍스트 생성")

    # ── STEP 1: 기본 설정 ──
    st.markdown("### STEP 1. 기본 설정")
    col1, col2 = st.columns(2)
    with col1:
        report_date = st.text_input("보고서 기준일",
                                    value=datetime.now().strftime("%Y년 %m월 %d일"),
                                    placeholder="예) 2026년 3월 20일")
        comp_sheet_name = st.text_input("경쟁사 시트 탭 이름", value="경쟁사 동향 분석 시트")
    with col2:
        report_month = st.text_input("보고 월",
                                     value=datetime.now().strftime("%Y년 %m월"),
                                     placeholder="예) 2026년 3월")
        mgmt_sheet_name = st.text_input("함소아 관리 시트 탭 이름", value="함소아 한의원 관리 시트")

    hamsoa_manual = st.number_input(
        "함소아 기사 발행 건수 (0이면 시트에서 자동 계산)",
        min_value=0, value=0)

    st.markdown("---")

    # ── STEP 2: 자생한방병원 전략 데이터 입력 ──
    st.markdown("### STEP 2. 자생한방병원 전략 데이터 입력")
    st.caption("전략유형 이름을 먼저 쓰고, 그 아래에 '키워드 건수' 형식으로 입력. 건수 없는 줄은 유형 헤더로 처리됩니다.")

    jasaeng_raw = st.text_area(
        "자생한방병원 전략 데이터",
        height=300,
        placeholder="""질환 타겟형
비염 1
일자목 1
혈당스파이크 24

시술 중심형

마케팅형
척추건강 증진사업 21
공동모금회 백미 29

브랜드 강화형
동작침법 52
발목 염좌 18""",
        help="숫자가 있는 줄 = 항목, 숫자 없는 줄 = 전략유형 헤더"
    )

    st.markdown("---")

    # ── STEP 3: 생성 버튼 ──
    if st.button("✨ 분석 텍스트 생성", type="primary", use_container_width=True):

        if not jasaeng_raw.strip():
            st.error("자생한방병원 전략 데이터를 입력해주세요.")
            st.stop()

        jasaeng_strategies, jasaeng_order, jasaeng_totals, jasaeng_grand_total = \
            parse_jasaeng_strategy(jasaeng_raw)

        with st.spinner("경쟁사 시트 불러오는 중..."):
            comp_sheet = get_hamsoa_sheet(comp_sheet_name)
            if comp_sheet is None:
                st.error(f"'{comp_sheet_name}' 시트를 찾을 수 없습니다.")
                st.stop()
            competitor_records, meta = parse_competitor_sheet(comp_sheet)

        with st.spinner("함소아 관리 시트 불러오는 중..."):
            mgmt_sheet = get_hamsoa_sheet(mgmt_sheet_name)
            if mgmt_sheet is None:
                st.error(f"'{mgmt_sheet_name}' 시트를 찾을 수 없습니다.")
                st.stop()
            hamsoa_articles, billing_data = parse_hamsoa_sheet(mgmt_sheet)

        # 함소아 기사 수: 수동 입력 우선, 없으면 시트 전체 행 수
        if hamsoa_manual > 0:
            hamsoa_count = hamsoa_manual
        else:
            hamsoa_count = len(hamsoa_articles)
        hospital_counts = meta.get('hospital_counts', {})

        st.success(f"✅ 자생 {jasaeng_grand_total}건 | 경쟁사 기사 {len(competitor_records)}건 | 함소아 {hamsoa_count}건 로드 완료")

        # 함소아 시트 파싱 결과 디버그
        if hamsoa_articles:
            with st.expander(f"📋 함소아 기사 현황 ({hamsoa_count}건) — 컬럼 확인", expanded=False):
                df_art = pd.DataFrame(hamsoa_articles)
                st.write("감지된 컬럼:", list(df_art.columns))
                # 진행현황 컬럼 찾아서 분포 보여주기
                status_col = next((c for c in df_art.columns if '진행' in c or '현황' in c), None)
                if status_col:
                    st.write(f"'{status_col}' 값 분포:", df_art[status_col].value_counts().to_dict())
                st.dataframe(df_art.head(10))
        else:
            st.warning("⚠️ 함소아 기사를 시트에서 읽지 못했습니다. '발행일', '구분', '제목' 컬럼이 있는 헤더 행이 있는지 확인해주세요.")

        # 자생 전략 집계 확인
        with st.expander("📊 자생한방병원 전략 집계 확인", expanded=True):
            cols = st.columns(len(jasaeng_order) if jasaeng_order else 1)
            for i, s in enumerate(jasaeng_order):
                t = jasaeng_totals.get(s, 0)
                items = jasaeng_strategies.get(s, [])
                with cols[i % len(cols)]:
                    st.metric(s, f"{t}건")
                    if items:
                        st.caption(" / ".join(f"{kw}({cnt})" for kw, cnt in items))
            st.markdown(f"**총 합계: {jasaeng_grand_total}건**")

        # ── 분석 텍스트 1: 발행 수량 비교 ──
        st.markdown("### 📝 분석 텍스트 1 — 발행 수량 비교")
        qty_texts = generate_quantity_analysis(
            hospital_counts, hamsoa_count, jasaeng_grand_total, report_date)
        qty_output = "\n\n".join(f"➡  {t}" for t in qty_texts)
        st.text_area("복사해서 PPT에 붙여넣으세요", value=qty_output, height=200, key="qty_out")

        # ── 분석 텍스트 2: 전략유형별 세부 분석 ──
        st.markdown("### 📝 분석 텍스트 2 — 전략유형별 세부 분석")
        strat_texts = generate_strategy_analysis(
            jasaeng_strategies, jasaeng_order, jasaeng_totals, jasaeng_grand_total,
            competitor_records, hospital_counts)
        strat_output = "\n\n".join(f"➡  {t}" for t in strat_texts)
        st.text_area("복사해서 PPT에 붙여넣으세요", value=strat_output, height=280, key="strat_out")

        # ── 데이터 미리보기 ──
        with st.expander("📋 경쟁사 기사 미리보기", expanded=False):
            if competitor_records:
                st.dataframe(pd.DataFrame(competitor_records).head(20))
                if hospital_counts:
                    st.write("병원별 건수:", hospital_counts)
            else:
                st.warning("⚠️ 경쟁사 데이터 없음")

        with st.expander("📋 함소아 기사 미리보기", expanded=False):
            if hamsoa_articles:
                st.dataframe(pd.DataFrame(hamsoa_articles))
            else:
                st.warning("⚠️ 함소아 기사 데이터 없음")
