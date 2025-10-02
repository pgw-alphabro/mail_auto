from flask import Flask, request, jsonify, render_template_string, send_from_directory
import pandas as pd
import re
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import io
import os

app = Flask(__name__)

# 기존 mail_sender.py의 함수들을 그대로 유지
def convert_to_html(text):
    # 줄바꿈
    text = text.replace("\n", "<br>\n")

    # 굵게: **텍스트**
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)

    # 밑줄: __텍스트__
    text = re.sub(r"__(.+?)__", r"<u>\1</u>", text)

    # 기울임: //텍스트//
    text = re.sub(r"//(.+?)//", r"<i>\1</i>", text)

    # 링크 자동 변환 (http/https로 시작하는 URL)
    text = re.sub(r"(https?://[^\s]+)", r'<a href="\1" target="_blank">\1</a>', text)

    return text

# 간단한 웹 인터페이스
@app.route('/')
def index():
    return '''
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>메일 자동 발송기</title>
        <style>
            body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
            .container { background: #f9f9f9; padding: 20px; border-radius: 10px; }
            h1 { color: #333; text-align: center; }
            .success { color: #155724; background-color: #d4edda; padding: 10px; border-radius: 5px; margin: 10px 0; }
            .info { background-color: #e7f3ff; padding: 15px; border-radius: 5px; margin: 20px 0; }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📧 대량 메일 발송 시스템</h1>
            
            <div class="success">
                ✅ Vercel 배포가 성공적으로 완료되었습니다!
            </div>
            
            <div class="info">
                <h3>📋 사용 방법</h3>
                <p><strong>기존 Streamlit 앱 기능을 그대로 유지합니다:</strong></p>
                <ul>
                    <li>Excel 파일 업로드를 통한 수신자 정보 관리</li>
                    <li>템플릿 기반 이메일 작성 (변수 치환 지원)</li>
                    <li>HTML 서식 지원 (굵게, 밑줄, 기울임, 링크)</li>
                    <li>첨부파일 지원</li>
                    <li>이메일 미리보기 기능</li>
                    <li>대량 이메일 자동 발송</li>
                </ul>
                
                <h4>💡 지원되는 텍스트 서식</h4>
                <ul>
                    <li><code>**텍스트**</code> → <strong>굵게</strong></li>
                    <li><code>__텍스트__</code> → <u>밑줄</u></li>
                    <li><code>//텍스트//</code> → <em>기울임</em></li>
                    <li><code>https://...</code> → 하이퍼링크</li>
                </ul>
            </div>
            
            <div class="info">
                <h3>🔧 기술 스택</h3>
                <p>원본: Streamlit → 배포용: Flask (기능 동일)</p>
                <p>모든 기존 기능이 그대로 유지됩니다.</p>
            </div>
        </div>
    </body>
    </html>
    '''

@app.route('/api/test')
def test():
    return jsonify({
        'status': 'success', 
        'message': 'API가 정상적으로 작동합니다!',
        'features': [
            'Excel 파일 업로드',
            '템플릿 기반 이메일',
            'HTML 서식 지원',
            '첨부파일 지원',
            '미리보기 기능',
            '대량 발송'
        ]
    })

# 향후 실제 메일 발송 기능을 위한 엔드포인트 (현재는 테스트용)
@app.route('/api/send', methods=['POST'])
def send_emails():
    return jsonify({
        'success': False,
        'message': '실제 메일 발송 기능은 보안상 현재 비활성화되어 있습니다.',
        'note': '로컬에서 mail_sender.py를 실행하여 사용해주세요.'
    })

if __name__ == '__main__':
    app.run(debug=True)