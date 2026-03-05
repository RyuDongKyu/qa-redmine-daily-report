import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import requests
from datetime import datetime, timedelta
import os
import json
import time

# ==========================================
# 1. 환경 설정
# ==========================================
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
SHEET_ID = os.getenv("SHEET_ID")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")
RECIPIENT_EMAIL = os.getenv("RECIPIENT_EMAIL")
JSON_KEY_FILE = "service_key.json"

# ==========================================
# 2. 데이터 추출 (KST 시간 보정 + 제외 로직)
# ==========================================
def get_yesterday_issues():
    print("🌐 구글 시트 데이터 추출 중...")
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
        g_client = gspread.authorize(creds)
        sheet = g_client.open_by_key(SHEET_ID).worksheet("통합_issues")
        all_data = sheet.get_all_values()
        rows = all_data[1:]
    except Exception as e:
        print(f"❌ 구글 시트 접속 실패: {e}")
        return "", []

    # [수정] 날짜 계산 로직: 월요일이면 금~일 데이터 추출
    utc_now = datetime.utcnow()
    kst_now = utc_now + timedelta(hours=9)
    current_weekday = kst_now.weekday()  # 0:월, 1:화, ..., 6:일

    target_dates = []
    if current_weekday == 0:  # 월요일(0)인 경우
        print("📅 오늘은 월요일입니다. 금, 토, 일 데이터를 추출합니다.")
        # 금(3일 전), 토(2일 전), 일(1일 전)
        for i in [3, 2, 1]:
            target_dates.append(kst_now - timedelta(days=i))
    else:
        # 평일인 경우 어제(1일 전) 데이터만 추출
        target_dates.append(kst_now - timedelta(days=1))

    # 비교를 위한 날짜 포맷 리스트 생성 (2024-05-20 및 2024. 5. 20. 형태 모두 대응)
    target_formats = []
    for d in target_dates:
        dash = d.strftime('%Y-%m-%d')
        dot = d.strftime('%Y. %m. %d.').replace('. 0', '. ')
        if dot.startswith('0'): dot = dot[1:]
        target_formats.append(dash)
        target_formats.append(dot)

    # 메일 제목 등에 표시될 날짜 문자열 결정
    if len(target_dates) > 1:
        date_display = f"{target_dates[0].strftime('%Y-%m-%d')} ~ {target_dates[-1].strftime('%Y-%m-%d')}"
    else:
        date_display = target_dates[0].strftime('%Y-%m-%d')

    print(f"📅 대상 기간: {date_display}")

    filtered_rows = []
    for row in rows:
        try:
            # 1. 날짜 확인 (AJ열 = 인덱스 35)
            input_time = row[35].strip() if len(row) > 35 else ""
            
            # [수정] 대상 날짜 리스트 중 하나라도 시트 날짜에 포함되어 있는지 확인
            is_target_day = any(fmt in input_time for fmt in target_formats)
            
            if is_target_day:
                # 2. 필수 조건 확인 (42열 값 존재 여부)
                col_42_val = row[41].strip() if len(row) > 41 else ""
                
                # 42열(AP)에 값이 없으면 제외
                if not col_42_val:
                    continue

                filtered_rows.append({
                    "no": row[0].strip(),
                    "category": row[1].strip() if len(row) > 1 else "미분류",
                    "type": row[3].strip() if len(row) > 3 else "",
                    "status": row[5].strip() if len(row) > 5 else "",
                    "priority": row[6].strip() if len(row) > 6 else "",
                    "title": row[7].strip() if len(row) > 7 else "",
                    "registrar": row[8].strip() if len(row) > 8 else "",
                    "manager": row[9].strip() if len(row) > 9 else "",
                    "date": input_time[:10],
                    "content": " | ".join([row[i].strip() for i in range(27, 32) if len(row) > i and row[i].strip()])
                })
        except: continue
        
    print(f"📝 필터링 후 추출된 이슈 수: {len(filtered_rows)}건")
    return date_display, filtered_rows

# ==========================================
# 3. 수동 리포트 생성기 (AI 실패 시 작동)
# ==========================================
def generate_manual_report(date_str, issues, error_msg=""):
    print("⚠️ AI 생성 실패. 수동 리포트 모드로 전환합니다.")
    
    grouped = {}
    for issue in issues:
        cat = issue['category']
        if cat not in grouped: grouped[cat] = []
        grouped[cat].append(issue)

    html = f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #333;">
        <h2>안녕하세요, {date_str} QA 등록 이슈 리포트입니다.</h2>
        <p style="color: red; font-size: 12px;">※ AI 서버 연결 불안정으로 인해 수동 생성된 리포트입니다. (사유: {error_msg})</p>
    """

    for cat, items in grouped.items():
        html += f"<h3 style='border-bottom: 2px solid #555; padding-bottom: 5px; margin-top: 30px;'>📂 {cat}</h3>"
        html += """
        <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
            <thead>
                <tr style="background-color: #f2f2f2; text-align: left;">
                    <th style="border: 1px solid #ddd; padding: 8px;">번호</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">등록일</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">상태</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">유형</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">우선순위</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">제목</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">등록자</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">담당자</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">비고</th>
                </tr>
            </thead>
            <tbody>
        """
        for item in items:
            html += f"""
                <tr>
                    <td style="border: 1px solid #ddd; padding: 8px;"><a href="https://projects.rsupport.com/issues/{item['no']}">#{item['no']}</a></td>
                    <td style="border: 1px solid #ddd; padding: 8px;">{item['date']}</td>
                    <td style="border: 1px solid #ddd; padding: 8px;">{item['status']}</td>
                    <td style="border: 1px solid #ddd; padding: 8px;">{item['type']}</td>
                    <td style="border: 1px solid #ddd; padding: 8px;">{item['priority']}</td>
                    <td style="border: 1px solid #ddd; padding: 8px;">{item['title']}</td>
                    <td style="border: 1px solid #ddd; padding: 8px;">{item['registrar']}</td>
                    <td style="border: 1px solid #ddd; padding: 8px;">{item['manager']}</td>
                    <td style="border: 1px solid #ddd; padding: 8px;">{item['content']}</td>
                </tr>
            """
        html += "</tbody></table>"
    html += "</body></html>"
    return html

# ==========================================
# 4. AI 리포트 시도 (실패 시 수동 전환)
# ==========================================
def ask_gemini(date_str, issues):
    # [수정됨] 필수 준수 사항을 강력하게 명시
    prompt = f"""
    당신은 'Redmine Daily Report Agent'입니다. 
    아래 [작성 원칙 v9.5]와 [인라인 HTML 가이드]를 반드시 **100% 준수**하여 본문을 작성하세요.

    [작성 원칙 v9.5 - 필수 준수 사항]
    1. 인사말: "안녕하세요, {date_str} QA 등록 이슈 리포트입니다."로 시작할 것.
    2. 그룹화: 'category'별로 섹션을 나눌 것. (예: <h3 class='cat-title'>📂 프로젝트명</h3>)
    3. 테이블 순서: 번호(#no), 등록일, 상태, 유형, 우선순위, 제목, 등록자, 담당자, 요약(AI) 순서로 컬럼을 배치할 것.
    4. 요약(AI) 처리: 'content'를 **반드시 한국어 두 문장**으로 핵심만 요약하여 '요약(AI)' 컬럼에 넣을 것.
    5. 링크 생성: 번호(#no)에는 반드시 <a href="https://projects.rsupport.com/issues/{{no}}">#{{no}}</a> 링크를 적용할 것.
    6. 데이터 변형 금지: 제목, 번호(#no), 등록일, 상태, 유형, 우선순위, 제목, 등록자, 담당자 등의 텍스트는 원문 그대로 유지할 것.

    [인라인 HTML 가이드 - 필수 적용]
    - <table style="width:100%; border-collapse:collapse; font-family:'Malgun Gothic',sans-serif; font-size:12px; border:1px solid #ddd;">
    - <th style="background-color:#f2f2f2; border:1px solid #ddd; padding:8px; font-weight:bold; text-align:center;">
    - <td style="border:1px solid #ddd; padding:8px; text-align:left;">
    - <td style="border:1px solid #ddd; padding:8px; text-align:center;"> (번호(#no), 등록일, 상태, 유형, 우선순위, 제목, 등록자, 담당자, 요약(AI))

    데이터: {json.dumps(issues, ensure_ascii=False)}
    """

    #candidate_models = ["gemini-3-pro-preview", "gemini-3-flash-preview", "gemini-2.5-flash", "gemini-2.5-flash-lite"]
    candidate_models = ["gemini-3.1-pro-preview", "gemini-3-flash-preview", "gemini-2.5-flash", "gemini-2.5-flash-lite"]
    headers = {'Content-Type': 'application/json'}
    data = {"contents": [{"parts": [{"text": prompt}]}]}
    
    last_error = ""

    for model in candidate_models:
        # 모델명 에러(404) 방지를 위해 정확한 엔드포인트 사용 확인
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={GEMINI_API_KEY}"
        try:
            print(f"🤖 AI 호출 시도: {model} ...")
            res = requests.post(url, headers=headers, json=data, timeout=120)
            
            if res.status_code == 200:
                print(f"✅ AI 리포트 생성 성공! (모델: {model})")
                
                raw_text = res.json()['candidates'][0]['content']['parts'][0]['text']
                clean_html = raw_text.replace('```html', '').replace('```', '').strip()
                
                # [수정] 인사말 다음 줄에 표시하기 위해 div 태그로 변경하고 여백 조절
                success_msg = f"<div style='color: #0052cc; font-size: 12px; font-weight: bold; margin-top: 10px; margin-bottom: 20px;'>✅ AI 분석 완료 (사용 모델: {model})</div>"
                
                # [핵심] </h2> 태그(인사말 끝) 바로 뒤에 문구를 삽입하여 다음 줄에 출력되게 함
                if "</h2>" in clean_html:
                    final_html = clean_html.replace("</h2>", f"</h2>{success_msg}")
                else:
                    # 인사말 태그가 없을 경우 맨 앞에 삽입
                    final_html = success_msg + clean_html
                    
                return final_html

            elif res.status_code == 429:
                time.sleep(5)
            else:
                last_error = f"{model} Error ({res.status_code})"
        except Exception as e:
            last_error = str(e)
            continue

    # 모든 모델 시도 실패 시 수동 리포트 반환
    return generate_manual_report(date_str, issues, last_error)

# ==========================================
# 5. 메일 발송 (다중 수신자 지원 수정됨)
# ==========================================
def send_email(subject, body):
    # [수정] 이메일 주소가 콤마(,)로 구분되어 있을 경우 리스트로 변환
    if "," in RECIPIENT_EMAIL:
        recipient_list = [addr.strip() for addr in RECIPIENT_EMAIL.split(',')]
    else:
        recipient_list = [RECIPIENT_EMAIL.strip()]

    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    # 헤더에는 보기 좋게 콤마로 합쳐서 표시 (예: "a@test.com, b@test.com")
    msg['To'] = ", ".join(recipient_list) 
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'html'))
    
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as s:
            s.starttls()
            s.login(SENDER_EMAIL, SENDER_PASSWORD)
            # send_message는 헤더(To)에 적힌 모든 수신자에게 자동으로 발송합니다.
            s.send_message(msg)
        print(f"✅ 메일 발송 완료 (수신자: {len(recipient_list)}명)")
    except Exception as e:
        print(f"❌ 메일 발송 실패: {e}")

if __name__ == "__main__":
    y_date, issues = get_yesterday_issues()
    if issues:
        final_html = ask_gemini(y_date, issues)
        send_email(f"[일일보고] {y_date} QA 등록 이슈 현황", final_html)
    else:
        send_email(f"[일일보고] {y_date} QA 등록 이슈 없음", f"<h3>{y_date} 등록된 QA 이슈가 없습니다.</h3>")
