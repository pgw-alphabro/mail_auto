from flask import Flask, request, jsonify, render_template_string, send_file
import pandas as pd
import re
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import io
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

# 기존 mail_sender.py의 convert_to_html 함수 그대로 유지
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

# HTML 템플릿 (기존 Streamlit UI와 동일한 기능)
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>메일 자동 발송기</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .main-container {
            max-width: 1400px;
            margin: 0 auto;
            display: grid;
            grid-template-columns: 350px 1fr;
            gap: 30px;
            min-height: calc(100vh - 40px);
        }
        
        .sidebar {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
            border-radius: 20px;
            padding: 30px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            height: fit-content;
            position: sticky;
            top: 20px;
        }
        
        .container {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
            border-radius: 20px;
            padding: 40px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
        }
        
        h1 {
            color: #2d3748;
            text-align: center;
            margin-bottom: 40px;
            font-size: 2.5rem;
            font-weight: 700;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }
        
        .sidebar h3 {
            color: #4a5568;
            margin-bottom: 15px;
            font-size: 1.2rem;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .form-section {
            background: #f8fafc;
            border-radius: 15px;
            padding: 25px;
            margin-bottom: 25px;
            border: 1px solid #e2e8f0;
        }
        
        .form-section h3 {
            color: #2d3748;
            margin-bottom: 20px;
            font-size: 1.3rem;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .form-group {
            margin-bottom: 20px;
        }
        
        label {
            display: block;
            margin-bottom: 8px;
            font-weight: 600;
            color: #4a5568;
            font-size: 0.95rem;
        }
        
        input[type="text"], textarea, input[type="file"] {
            width: 100%;
            padding: 12px 16px;
            border: 2px solid #e2e8f0;
            border-radius: 10px;
            font-size: 14px;
            transition: all 0.3s ease;
            background: white;
        }
        
        input[type="text"]:focus, textarea:focus {
            outline: none;
            border-color: #667eea;
            box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
        }
        
        textarea {
            height: 180px;
            resize: vertical;
            font-family: 'Monaco', 'Menlo', monospace;
            line-height: 1.5;
        }
        
        .button-group {
            display: flex;
            gap: 15px;
            flex-wrap: wrap;
            margin-top: 30px;
        }
        
        button {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 14px 28px;
            border: none;
            border-radius: 10px;
            cursor: pointer;
            font-size: 15px;
            font-weight: 600;
            transition: all 0.3s ease;
            box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
        }
        
        button:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 25px rgba(102, 126, 234, 0.4);
        }
        
        .btn-success {
            background: linear-gradient(135deg, #48bb78 0%, #38a169 100%);
            box-shadow: 0 4px 15px rgba(72, 187, 120, 0.3);
        }
        
        .btn-success:hover {
            box-shadow: 0 8px 25px rgba(72, 187, 120, 0.4);
        }
        
        .btn-preview {
            background: linear-gradient(135deg, #ed8936 0%, #dd6b20 100%);
            box-shadow: 0 4px 15px rgba(237, 137, 54, 0.3);
        }
        
        .btn-preview:hover {
            box-shadow: 0 8px 25px rgba(237, 137, 54, 0.4);
        }
        .info-box {
            background: linear-gradient(135deg, #ebf8ff 0%, #bee3f8 100%);
            border: 1px solid #90cdf4;
            border-radius: 15px;
            padding: 20px;
            margin: 25px 0;
            box-shadow: 0 4px 15px rgba(144, 205, 244, 0.2);
        }
        
        .info-box ul {
            margin: 10px 0 0 20px;
        }
        
        .info-box li {
            margin: 8px 0;
            color: #2d3748;
        }
        
        .preview-box {
            background: #f7fafc;
            border: 2px solid #e2e8f0;
            border-radius: 15px;
            padding: 25px;
            margin: 25px 0;
            display: none;
            box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        }
        
        .preview-box h3 {
            color: #2d3748;
            margin-bottom: 20px;
            font-size: 1.3rem;
        }
        
        .error {
            color: #c53030;
            background: linear-gradient(135deg, #fed7d7 0%, #feb2b2 100%);
            border: 1px solid #fc8181;
            padding: 15px;
            border-radius: 10px;
            margin: 15px 0;
            font-weight: 500;
        }
        
        .success {
            color: #22543d;
            background: linear-gradient(135deg, #c6f6d5 0%, #9ae6b4 100%);
            border: 1px solid #68d391;
            padding: 15px;
            border-radius: 10px;
            margin: 15px 0;
            font-weight: 500;
        }
        
        .warning {
            color: #744210;
            background: linear-gradient(135deg, #fefcbf 0%, #faf089 100%);
            border: 1px solid #f6e05e;
            padding: 15px;
            border-radius: 10px;
            margin: 15px 0;
            font-weight: 500;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 15px 0;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        }
        
        th, td {
            border: 1px solid #e2e8f0;
            padding: 12px 16px;
            text-align: left;
        }
        
        th {
            background: linear-gradient(135deg, #edf2f7 0%, #e2e8f0 100%);
            font-weight: 600;
            color: #2d3748;
        }
        
        .loading {
            display: none;
            text-align: center;
            padding: 30px;
            color: #4a5568;
            font-size: 1.1rem;
        }
        
        .file-upload-area {
            border: 2px dashed #cbd5e0;
            border-radius: 15px;
            padding: 30px;
            text-align: center;
            transition: all 0.3s ease;
            background: #f7fafc;
        }
        
        .file-upload-area:hover {
            border-color: #667eea;
            background: #edf2f7;
        }
        
        .format-table {
            background: white;
            border-radius: 10px;
            overflow: hidden;
            margin-top: 15px;
        }
        
        .format-table th {
            background: #4a5568;
            color: white;
        }
        
        @media (max-width: 1024px) {
            .main-container {
                grid-template-columns: 1fr;
                gap: 20px;
            }
            
            .sidebar {
                position: static;
            }
        }
    </style>
</head>
<body>
    <div class="main-container">
        <!-- 사이드바 -->
        <div class="sidebar">
            <h3>📎 첨부파일</h3>
            <div class="form-group">
                <input type="file" id="attachment" name="attachment">
                <small style="color: #718096; margin-top: 5px; display: block;">선택사항</small>
            </div>
            
            <h3 style="margin-top: 30px;">💡 서식 가이드</h3>
            <p style="color: #4a5568; margin-bottom: 15px; font-size: 0.9rem;">이메일 본문에 서식을 지정할 수 있습니다:</p>
            <table class="format-table">
                <tr><th>입력 형식</th><th>결과</th></tr>
                <tr><td><code>**텍스트**</code></td><td><strong>굵게</strong></td></tr>
                <tr><td><code>__텍스트__</code></td><td><u>밑줄</u></td></tr>
                <tr><td><code>//텍스트//</code></td><td><em>기울임</em></td></tr>
                <tr><td><code>https://...</code></td><td>링크</td></tr>
            </table>
            
            <button class="btn-success" onclick="downloadSample()" style="width: 100%; margin-top: 25px;">
                📥 샘플 파일 다운로드
            </button>
        </div>

        <!-- 메인 컨텐츠 -->
        <div class="container">
            <h1>📧 대량 메일 발송</h1>
            
            <div class="info-box">
                <h3 style="margin-bottom: 15px; color: #2b6cb0;">📋 사용 방법</h3>
                <p style="margin-bottom: 10px;">EXCEL 파일을 업로드하면 수신자 정보에 따라 이메일을 자동 발송합니다.</p>
                <ul>
                    <li><strong>필수값</strong>: 이름, 이메일</li>
                    <li><strong>선택값</strong>: {일자}, {장소}, {내용}, {상세내용} 등 자유롭게 사용 가능</li>
                </ul>
            </div>

            <div class="form-section">
                <h3>📁 파일 업로드</h3>
                <div class="file-upload-area">
                    <input type="file" id="excel_file" accept=".xlsx,.xls" onchange="handleFileUpload()" style="margin-bottom: 10px;">
                    <p style="color: #718096; font-size: 0.9rem;">Excel 파일(.xlsx, .xls)을 선택해주세요</p>
                </div>
                <div id="file-preview"></div>
            </div>

            <div class="form-section">
                <h3>✏️ 이메일 작성</h3>
                <div class="form-group">
                    <label for="subject">제목</label>
                    <input type="text" id="subject" placeholder="이메일 제목을 입력해주세요" value="제목">
                </div>

                <div class="form-group">
                    <label for="body">본문</label>
                    <textarea id="body" placeholder="이메일 본문을 입력하세요...">안녕하세요 {이름}님,
제출하신 서류를 검토한 결과 아래 항목에 대해 보완이 필요합니다.

▶ 내용: {내용}
▶ 상세 내용: {상세내용}

감사합니다.</textarea>
                </div>

                <div class="button-group">
                    <button class="btn-preview" onclick="previewEmail()">🔍 미리보기</button>
                    <button onclick="sendEmails()">📨 이메일 발송</button>
                </div>
            </div>

            <div id="preview" class="preview-box">
                <h3>🔍 이메일 미리보기 (첫 번째 수신자 기준)</h3>
                <div id="preview-content"></div>
            </div>

            <div class="loading" id="loading">
                <p>📨 이메일을 발송 중입니다...</p>
            </div>

            <div id="result"></div>
        </div>
    </div></body>

    <script>
        let uploadedData = null;

        function downloadSample() {
            const csvContent = "이름,이메일,내용,상세내용\\n홍길동,test1@example.com,참가신청서 누락,참가신청서를 제출해주세요\\n김영희,test2@example.com,팀원 확인서 미제출,팀원 확인서를 제출해주세요";
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement("a");
            const url = URL.createObjectURL(blob);
            link.setAttribute("href", url);
            link.setAttribute("download", "sample_format.csv");
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }

        async function handleFileUpload() {
            const fileInput = document.getElementById('excel_file');
            const file = fileInput.files[0];
            
            if (!file) return;

            const formData = new FormData();
            formData.append('excel_file', file);

            try {
                const response = await fetch('/api/upload', {
                    method: 'POST',
                    body: formData
                });
                
                const result = await response.json();
                
                if (result.success) {
                    uploadedData = result.data;
                    displayFilePreview(result.preview);
                } else {
                    document.getElementById('result').innerHTML = `<div class="error">${result.error}</div>`;
                }
            } catch (error) {
                document.getElementById('result').innerHTML = `<div class="error">파일 업로드 중 오류: ${error.message}</div>`;
            }
        }

        function displayFilePreview(preview) {
            const previewDiv = document.getElementById('file-preview');
            previewDiv.innerHTML = `
                <div class="success">
                    <h4>📋 엑셀 미리보기</h4>
                    ${preview}
                </div>
            `;
        }

        async function previewEmail() {
            if (!uploadedData) {
                alert('먼저 엑셀 파일을 업로드해주세요.');
                return;
            }

            const subject = document.getElementById('subject').value;
            const body = document.getElementById('body').value;

            try {
                const response = await fetch('/api/preview', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({
                        subject: subject,
                        body: body,
                        data: uploadedData
                    })
                });
                
                const result = await response.json();
                
                if (result.success) {
                    document.getElementById('preview-content').innerHTML = `
                        <p><strong>제목:</strong> ${result.preview_subject}</p>
                        <div><strong>본문:</strong></div>
                        <div style="border: 1px solid #ddd; padding: 10px; margin-top: 10px;">
                            ${result.preview_body}
                        </div>
                    `;
                    document.getElementById('preview').style.display = 'block';
                } else {
                    document.getElementById('result').innerHTML = `<div class="error">${result.error}</div>`;
                }
            } catch (error) {
                document.getElementById('result').innerHTML = `<div class="error">미리보기 중 오류: ${error.message}</div>`;
            }
        }

        async function sendEmails() {
            if (!uploadedData) {
                alert('먼저 엑셀 파일을 업로드해주세요.');
                return;
            }

            const subject = document.getElementById('subject').value;
            const body = document.getElementById('body').value;
            const attachmentFile = document.getElementById('attachment').files[0];

            const formData = new FormData();
            formData.append('subject', subject);
            formData.append('body', body);
            formData.append('data', JSON.stringify(uploadedData));
            
            if (attachmentFile) {
                formData.append('attachment', attachmentFile);
            }

            document.getElementById('loading').style.display = 'block';
            document.getElementById('result').innerHTML = '';

            try {
                const response = await fetch('/api/send', {
                    method: 'POST',
                    body: formData
                });
                
                const result = await response.json();
                
                document.getElementById('loading').style.display = 'none';
                
                if (result.success) {
                    document.getElementById('result').innerHTML = `<div class="success">총 ${result.sent_count}건의 이메일을 성공적으로 보냈습니다.</div>`;
                } else {
                    document.getElementById('result').innerHTML = `<div class="error">${result.error}</div>`;
                }
            } catch (error) {
                document.getElementById('loading').style.display = 'none';
                document.getElementById('result').innerHTML = `<div class="error">이메일 발송 중 오류: ${error.message}</div>`;
            }
        }
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/api/upload', methods=['POST'])
def upload_file():
    try:
        excel_file = request.files.get('excel_file')
        if not excel_file:
            return jsonify({'success': False, 'error': '파일이 업로드되지 않았습니다.'})
        
        # 엑셀 파일 읽기
        df = pd.read_excel(io.BytesIO(excel_file.read()))
        
        # 필수 컬럼 확인
        required_columns = {"이름", "이메일"}
        missing_required = required_columns - set(df.columns)
        if missing_required:
            return jsonify({'success': False, 'error': f'필수 컬럼 누락: {", ".join(missing_required)}'})
        
        # 데이터프레임을 HTML 테이블로 변환 (미리보기용)
        preview_html = df.head().to_html(classes='table', table_id='preview-table')
        
        return jsonify({
            'success': True,
            'data': df.to_dict('records'),
            'preview': preview_html
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': f'파일 처리 중 오류: {str(e)}'})

@app.route('/api/preview', methods=['POST'])
def preview_email():
    try:
        data = request.json
        subject_template = data.get('subject', '')
        body_template = convert_to_html(data.get('body', ''))
        records = data.get('data', [])
        
        if not records:
            return jsonify({'success': False, 'error': '데이터가 없습니다.'})
        
        # 첫 번째 행으로 미리보기 생성
        first_row = records[0]
        
        try:
            preview_subject = subject_template.format(**first_row)
            preview_body = body_template.format(**first_row)
            
            return jsonify({
                'success': True,
                'preview_subject': preview_subject,
                'preview_body': preview_body
            })
        except KeyError as e:
            return jsonify({'success': False, 'error': f'템플릿 변수 오류: {str(e)}'})
            
    except Exception as e:
        return jsonify({'success': False, 'error': f'미리보기 생성 중 오류: {str(e)}'})

@app.route('/api/send', methods=['POST'])
def send_emails():
    try:
        subject_template = request.form.get('subject', '')
        body_template = convert_to_html(request.form.get('body', ''))
        records = eval(request.form.get('data', '[]'))  # JSON 파싱
        
        if not records:
            return jsonify({'success': False, 'error': '데이터가 없습니다.'})
        
        # 이메일 설정 (기존 mail_sender.py와 동일)
        sender_email = 'mvptest.kr@gmail.com'
        sender_password = 'tyft tvur rkwg uics'
        
        # SMTP 서버 연결
        smtp_server = 'smtp.gmail.com'
        smtp_port = 587
        
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        
        sent_count = 0
        
        # 첨부파일 처리
        attachment_file = request.files.get('attachment')
        attachment_bytes = None
        attachment_name = None
        if attachment_file and attachment_file.filename:
            attachment_bytes = attachment_file.read()
            attachment_name = secure_filename(attachment_file.filename)
        
        # 각 행에 대해 이메일 발송 (기존 로직과 동일)
        for row_dict in records:
            try:
                subject = subject_template.format(**row_dict)
                body = body_template.format(**row_dict)
            except KeyError as e:
                continue  # 변수가 없는 경우 건너뛰기
            
            # 이메일 메시지 생성
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = row_dict['이메일']
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'html'))
            
            # 첨부파일 추가
            if attachment_bytes:
                part = MIMEApplication(attachment_bytes, Name=attachment_name)
                part['Content-Disposition'] = f'attachment; filename="{attachment_name}"'
                msg.attach(part)
            
            # 이메일 발송
            server.send_message(msg)
            sent_count += 1
        
        server.quit()
        
        return jsonify({
            'success': True,
            'sent_count': sent_count
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': f'이메일 발송 중 오류: {str(e)}'})

if __name__ == '__main__':
    app.run(debug=True)
