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

    # [수정됨] GitHub 서버(UTC) 시간을 한국 시간(KST)으로 변환 후 어제 날짜 계산
    # UTC 현재 시간 가져오기
    utc_now = datetime.utcnow()
    # KST = UTC + 9시간
    kst_now = utc_now + timedelta(hours=9)
    # KST 기준 어제 날짜
    target_date = kst_now - timedelta(days=1)

    target_dash = target_date.strftime('%Y-%m-%d')
    target_dot = target_date.strftime('%Y. %m. %d.').replace('. 0', '. ')
    if target_dot.startswith('0'): target_dot = target_dot[1:]
    
    print(f"📅 한국 시간 기준 어제 날짜: {target_dash}")

    filtered_rows = []
    for row in rows:
        try:
            # 1. 날짜 확인 (AJ열 = 인덱스 35)
            input_time = row[35].strip() if len(row) > 35 else ""
            
            if target_dash in input_time or target_dot in input_time:
                # 2. 필수 조건 확인 (42열 값 존재 여부)
                col_42_val = row[41].strip() if len(row) > 41 else ""
                
                # 42열(AP)에 값이 없으면 제외
                if not col_42_val:
                    continue

                # 3. 외부 유입 확인 (등록자 공란)
                qa_reg = row[24].strip() if len(row) > 24 else "" 
                dev_reg = row[25].strip() if len(row) > 25 else "" 

                if not qa_reg and not dev_reg:
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
    return target_dash, filtered_rows

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
        <h2>안녕하세요, {date_str} QA가 등록한 레드마인 목록입니다.</h2>
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
    1. 인사말: "안녕하세요, {date_str} QA가 등록한 레드마인 목록입니다."로 시작할 것.
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
    
    candidate_models = ["gemini-2.5-flash", "gemini-2.5-flash-lite", "gemini-pro"]
    headers = {'Content-Type': 'application/json'}
    data = {"contents": [{"parts": [{"text": prompt}]}]}
    
    last_error = ""

    for model in candidate_models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={GEMINI_API_KEY}"
        try:
            print(f"🤖 AI 호출 시도: {model} ...")
            res = requests.post(url, headers=headers, json=data, timeout=30)
            
            if res.status_code == 200:
                print("✅ AI 리포트 생성 성공!")
                return res.json()['candidates'][0]['content']['parts'][0]['text'].replace('```html', '').replace('```', '').strip()
            elif res.status_code == 429:
                time.sleep(5)
            else:
                last_error = f"{model} Error ({res.status_code})"
        except Exception as e:
            last_error = str(e)
            continue

    return generate_manual_report(date_str, issues, last_error)

# ==========================================
# 5. 메일 발송
# ==========================================
def send_email(subject, body):
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECIPIENT_EMAIL
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'html'))
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as s:
            s.starttls()
            s.login(SENDER_EMAIL, SENDER_PASSWORD)
            s.send_message(msg)
        print("✅ 메일 발송 완료")
    except Exception as e:
        print(f"❌ 메일 발송 실패: {e}")

if __name__ == "__main__":
    y_date, issues = get_yesterday_issues()
    if issues:
        final_html = ask_gemini(y_date, issues)
        send_email(f"[일일보고] {y_date} QA 레드마인 등록 현황", final_html)
    else:
        send_email(f"[일일보고] {y_date} QA 레드마인 등록 없음", f"<h3>{y_date} 자 등록된 이슈가 없습니다.</h3>")
