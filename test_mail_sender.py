import streamlit as st
import pandas as pd
import re
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# 페이지 레이아웃
st.set_page_config(page_title="메일 자동 발송기", layout="wide")

# ------------------------ Sidebar 구성 ------------------------
st.sidebar.markdown("## 📎 첨부파일 (선택사항)")
attachment_file = st.sidebar.file_uploader("첨부할 파일을 선택하세요", type=None)

st.sidebar.markdown(""" ## 💡 글꼴 도구 사용하기
    - 아래 문법을 이용해 이메일 본문에 서식을 지정할 수 있습니다.  
    - 일반 문장처럼 작성해 주세요. 줄바꿈도 엔터로만 하면 됩니다.

    ### ✍️ 지원되는 텍스트 서식

    | 입력 형식          | 의미         |
    |-------------------|--------------|
    | `**텍스트**`       | **굵게**     |
    | `__텍스트__`       | _밑줄_       |
    | `//텍스트//`       | *기울임*     |
    | `https://...`      | 하이퍼링크   |

    ---
                    """)

# ------------------------ 이메일 작성 메인 UI ------------------------
st.title("📧 대량 메일 발송")

st.markdown("""
EXCEL 파일을 업로드하면 수신자 정보에 따라 이메일을 자동 발송합니다.  
- **필수값 (꼭 넣어주세요)**: `이름`, `이메일`  
- **선택값**: `{일자}`, `{장소}`, `{내용}`, `{상세내용}` 등 자유롭게 사용 가능  
""")

# 샘플 EXCEL 다운로드
sample_xlsx = """이름,이메일,내용,상세내용
홍길동,test1@example.com,참가신청서 누락,참가신청서를 제출해주세요
김영희,test2@example.com,팀원 확인서 미제출,팀원 확인서를 제출해주세요
"""
st.download_button(
    label="📥 샘플 파일 다운로드",
    data=sample_xlsx,
    file_name="sample_format.xlsx",
    mime="text/xlsx"
)
st.markdown("---") 
st.subheader("📁 파일 업로드")
uploaded_file = st.file_uploader("내용을 포함한 엑셀 파일(.xlsx)을 업로드하세요", type="xlsx")

sender_email = 'mvptest.kr@gmail.com'
sender_password = 'tyft tvur rkwg uics'

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

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.text("📋 엑셀 미리보기")
        st.dataframe(df)
    except Exception as e:
        st.error(f"❌ 엑셀 파일을 읽는 중 오류 발생: {e}")
    st.markdown("---") 
    st.subheader("💡 글꼴 도구 사용하기")
    with st.expander("(클릭하여 보기)"):
        st.markdown("""
    - 아래 문법을 이용해 이메일 본문에 서식을 지정할 수 있습니다.  
    - 일반 문장처럼 작성해 주세요. 줄바꿈도 엔터로만 하면 됩니다.

    ### ✍️ 지원되는 텍스트 서식

    | 입력 형식          | 의미         | 예시 입력                    |
    |-------------------|--------------|------------------------------|
    | `**텍스트**`       | **굵게**     | `**보완 필요**`              |
    | `__텍스트__`       | _밑줄_       | `__필수 제출__`              |
    | `//텍스트//`       | *기울임*     | `//추가 설명//`              |
    | `https://...`      | 하이퍼링크   | `https://example.com`        |

    ---
                    """)
    st.markdown("---")  
    st.markdown("## ✏️ 이메일 제목 및 본문 템플릿")

    subject_template = st.text_input(
        "제목을 입력해주세요",
        value="제목"
    )

    body_plain_text = st.text_area(
        "이메일 본문",
        height=350,
        value="""[작성 예시]
안녕하세요 {이름}님,
제출하신 서류를 검토한 결과 아래 항목에 대해 보완이 필요합니다.

▶ 내용: {내용}
▶ 상세 내용: {상세내용}

감사합니다."""
    )

    # HTML 변환 적용
    body_template = convert_to_html(body_plain_text)

    # 템플릿 변수 확인
    required_columns = {"이름", "이메일"}
    missing_required = required_columns - set(df.columns)
    if missing_required:
        st.error(f"❌ 필수 컬럼 누락: {', '.join(missing_required)}")
    else:
        def extract_placeholders(template):
            return set(re.findall(r"{(\w+)}", template))

        placeholders = extract_placeholders(subject_template) | extract_placeholders(body_template)
        csv_columns = set(df.columns)
        missing_placeholders = placeholders - csv_columns

        if missing_placeholders:
            st.warning(f"⚠️ 업로드한 파일에 다음의 값이 존재하지 않습니다. : {', '.join(missing_placeholders)}")

        # 미리보기
        st.markdown("---")  
        st.subheader("🔍 이메일 미리보기 (첫 번째 수신자 기준)")
        first_row = df.iloc[0].to_dict()
        try:
            preview_subject = subject_template.format(**first_row)
            preview_body = body_template.format(**first_row)

            st.text(f"제목: {preview_subject}")
            st.markdown(preview_body, unsafe_allow_html=True)

            with st.expander("🔍 HTML 원본 보기"):
                st.code(preview_body, language='html')

        except KeyError as e:
            st.error(f"⚠️ 오류: {e}")

        # 이메일 발송 버튼
        st.markdown("---")  
        if st.button("📨 이메일 보내기"):
            try:
                smtp_server = 'smtp.gmail.com'
                smtp_port = 587

                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
                server.login(sender_email, sender_password)

                sent_count = 0

                # 첨부파일 읽기
                attachment_bytes = None
                attachment_name = None
                if attachment_file is not None:
                    attachment_bytes = attachment_file.read()
                    attachment_name = attachment_file.name

                for _, row in df.iterrows():
                    row_dict = row.to_dict()
                    try:
                        subject = subject_template.format(**row_dict)
                        body = body_template.format(**row_dict)
                    except KeyError as e:
                        st.error(f"❌ {row_dict.get('이름', '수신자')}에게 메일 생성 실패 - 누락된 변수: {e}")
                        continue

                    msg = MIMEMultipart()
                    msg['From'] = sender_email
                    msg['To'] = row_dict['이메일']
                    msg['Subject'] = subject
                    msg.attach(MIMEText(body, 'html'))

                    if attachment_bytes:
                        part = MIMEApplication(attachment_bytes, Name=attachment_name)
                        part['Content-Disposition'] = f'attachment; filename="{attachment_name}"'
                        msg.attach(part)

                    server.send_message(msg)
                    sent_count += 1
                    st.write(f"✅ {row_dict['이름']} ({row_dict['이메일']}) 에게 발송 완료")

                server.quit()
                st.success(f"총 {sent_count}건의 이메일을 성공적으로 보냈습니다.")

            except Exception as e:
                st.error(f"❌ 이메일 전송 중 에러 발생: {e}")
