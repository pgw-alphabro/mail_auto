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
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            background: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .sidebar {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 20px;
        }
        h1 {
            color: #333;
            text-align: center;
            margin-bottom: 30px;
        }
        .form-group {
            margin-bottom: 20px;
        }
        label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
            color: #555;
        }
        input[type="text"], textarea, input[type="file"] {
            width: 100%;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 5px;
            font-size: 14px;
            box-sizing: border-box;
        }
        textarea {
            height: 200px;
            resize: vertical;
        }
        button {
            background-color: #007bff;
            color: white;
            padding: 12px 30px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 16px;
            margin: 10px 5px 10px 0;
        }
        button:hover {
            background-color: #0056b3;
        }
        .btn-success {
            background-color: #28a745;
        }
        .btn-success:hover {
            background-color: #218838;
        }
        .info-box {
            background-color: #e7f3ff;
            border: 1px solid #b3d9ff;
            border-radius: 5px;
            padding: 15px;
            margin: 20px 0;
        }
        .preview-box {
            background-color: #f8f9fa;
            border: 1px solid #dee2e6;
            border-radius: 5px;
            padding: 15px;
            margin: 20px 0;
            display: none;
        }
        .error {
            color: #dc3545;
            background-color: #f8d7da;
            border: 1px solid #f5c6cb;
            padding: 10px;
            border-radius: 5px;
            margin: 10px 0;
        }
        .success {
            color: #155724;
            background-color: #d4edda;
            border: 1px solid #c3e6cb;
            padding: 10px;
            border-radius: 5px;
            margin: 10px 0;
        }
        .warning {
            color: #856404;
            background-color: #fff3cd;
            border: 1px solid #ffeaa7;
            padding: 10px;
            border-radius: 5px;
            margin: 10px 0;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 10px 0;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
        }
        .loading {
            display: none;
            text-align: center;
            padding: 20px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📧 대량 메일 발송</h1>
        
        <div class="info-box">
            <p>EXCEL 파일을 업로드하면 수신자 정보에 따라 이메일을 자동 발송합니다.</p>
            <ul>
                <li><strong>필수값 (꼭 넣어주세요)</strong>: 이름, 이메일</li>
                <li><strong>선택값</strong>: {일자}, {장소}, {내용}, {상세내용} 등 자유롭게 사용 가능</li>
            </ul>
        </div>

        <div class="sidebar">
            <h3>📎 첨부파일 (선택사항)</h3>
            <input type="file" id="attachment" name="attachment">
            
            <h3 style="margin-top: 20px;">💡 글꼴 도구 사용하기</h3>
            <p>아래 문법을 이용해 이메일 본문에 서식을 지정할 수 있습니다.</p>
            <table>
                <tr><th>입력 형식</th><th>의미</th></tr>
                <tr><td><code>**텍스트**</code></td><td><strong>굵게</strong></td></tr>
                <tr><td><code>__텍스트__</code></td><td><u>밑줄</u></td></tr>
                <tr><td><code>//텍스트//</code></td><td><em>기울임</em></td></tr>
                <tr><td><code>https://...</code></td><td>하이퍼링크</td></tr>
            </table>
        </div>

        <button class="btn-success" onclick="downloadSample()">📥 샘플 파일 다운로드</button>

        <div class="form-group">
            <label for="excel_file">📁 파일 업로드</label>
            <input type="file" id="excel_file" accept=".xlsx,.xls" onchange="handleFileUpload()">
        </div>

        <div id="file-preview"></div>

        <div class="form-group">
            <label for="subject">✏️ 이메일 제목 및 본문 템플릿</label>
            <input type="text" id="subject" placeholder="제목을 입력해주세요" value="제목">
        </div>

        <div class="form-group">
            <label for="body">이메일 본문</label>
            <textarea id="body" placeholder="이메일 본문을 입력하세요">안녕하세요 {이름}님,
제출하신 서류를 검토한 결과 아래 항목에 대해 보완이 필요합니다.

▶ 내용: {내용}
▶ 상세 내용: {상세내용}

감사합니다.</textarea>
        </div>

        <button onclick="previewEmail()">🔍 이메일 미리보기</button>
        <button onclick="sendEmails()">📨 이메일 보내기</button>

        <div id="preview" class="preview-box">
            <h3>🔍 이메일 미리보기 (첫 번째 수신자 기준)</h3>
            <div id="preview-content"></div>
        </div>

        <div class="loading" id="loading">
            <p>📨 이메일을 발송 중입니다...</p>
        </div>

        <div id="result"></div>
    </div>

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
