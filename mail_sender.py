import streamlit as st
import pandas as pd
import re
import smtplib
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# 페이지 레이아웃
st.set_page_config(page_title="하이웍스 메일 자동 발송기", layout="wide")

# ------------------------ Sidebar 구성 ------------------------
# SMTP 서버 설정을 먼저 가져와서 제목에 반영
smtp_server_temp = st.sidebar.selectbox(
    "SMTP 서버",
    options=["smtp.hiworks.com", "smtp.gmail.com", "smtp.naver.com"],
    index=0,
    help="SMTP 서버 선택",
    key="smtp_server_temp"
)

# SMTP 서버에 따른 제목 변경
if smtp_server_temp == "smtp.gmail.com":
    st.sidebar.markdown("## 🔐 Gmail 계정 설정")
    email_help = "Gmail 주소를 입력하세요 (예: user@gmail.com)"
    password_help = "Gmail 앱 비밀번호를 입력하세요 (16자리)"
elif smtp_server_temp == "smtp.naver.com":
    st.sidebar.markdown("## 🔐 네이버 계정 설정")
    email_help = "네이버 메일 주소를 입력하세요 (예: user@naver.com)"
    password_help = "네이버 계정 비밀번호를 입력하세요"
else:  # hiworks
    st.sidebar.markdown("## 🔐 하이웍스 계정 설정")
    email_help = "하이웍스 이메일 주소를 입력하세요"
    password_help = "하이웍스 앱 비밀번호를 입력하세요 (2단계 인증 시)"

sender_email = st.sidebar.text_input(
    "발신 이메일",
    value="",
    help=email_help
)

sender_password = st.sidebar.text_input(
    "비밀번호",
    type="password",
    help=password_help
)

st.sidebar.markdown("---")

# SMTP 서버 설정
st.sidebar.markdown("## ⚙️ SMTP 설정")
smtp_server = smtp_server_temp  # 위에서 선택한 서버 사용

smtp_port = st.sidebar.selectbox(
    "포트 선택",
    options=[587, 465, 25],
    index=0,
    help="587: STARTTLS (권장) / 465: SSL / 25: 일반"
)

# SMTP 서버별 안내 메시지
if smtp_server == "smtp.gmail.com":
    st.sidebar.info("""
    📧 **Gmail 설정 안내**
    1. Google 계정 → 보안 → 2단계 인증 활성화
    2. 앱 비밀번호 생성 (16자리)
    3. 위에 앱 비밀번호 입력 (일반 비밀번호 아님)
    """)
elif smtp_server == "smtp.naver.com":
    st.sidebar.info("""
    📧 **네이버 설정 안내**
    1. 네이버 메일 주소와 비밀번호 입력
    2. 2단계 인증 없어도 사용 가능
    3. 포트 587 권장
    """)
else:  # hiworks
    st.sidebar.info("""
    📧 **하이웍스 설정 안내**
    1. 하이웍스 웹메일 → 보안 설정 확인
    2. 앱 비밀번호 생성 (있다면)
    3. 관리자에게 SMTP 사용 권한 문의
    """)

# 발송 속도 조절
st.sidebar.markdown("## 🚀 발송 속도 설정")
send_delay = st.sidebar.slider(
    "메일 간 대기 시간 (초)",
    min_value=0,
    max_value=10,
    value=2,
    help="스팸 방지를 위해 2초 이상 권장"
)

batch_size = st.sidebar.number_input(
    "배치당 발송 수",
    min_value=1,
    max_value=100,
    value=10,
    help="이 수만큼 발송 후 긴 대기 시간 적용"
)

batch_delay = st.sidebar.slider(
    "배치 대기 시간 (초)",
    min_value=0,
    max_value=60,
    value=10,
    help=f"{batch_size}통 발송 후 대기 시간"
)

st.sidebar.markdown("---")

st.sidebar.markdown("## 📎 첨부파일 (선택사항)")
attachment_file = st.sidebar.file_uploader("첨부할 파일을 선택하세요", type=None)

st.sidebar.markdown("""## 💡 글꼴 도구 사용하기
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
st.title("📧 하이웍스 대량 메일 발송")

st.markdown("""
하이웍스 계정으로 EXCEL 파일 기반 이메일을 자동 발송합니다.  
- **필수값 (꼭 넣어주세요)**: `이름`, `이메일`  
- **선택값**: `{일자}`, `{장소}`, `{내용}`, `{상세내용}` 등 자유롭게 사용 가능  
""")

# 샘플 EXCEL 다운로드
sample_csv = """이름,이메일,내용,상세내용
홍길동,test1@example.com,참가신청서 누락,참가신청서를 제출해주세요
김영희,test2@example.com,팀원 확인서 미제출,팀원 확인서를 제출해주세요
이철수,test3@example.com,결제 확인 필요,결제 내역을 확인해주세요
"""

st.download_button(
    label="📥 샘플 파일 다운로드 (CSV)",
    data=sample_csv,
    file_name="sample_format.csv",
    mime="text/csv"
)

st.markdown("---") 

# HTML 변환 함수
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

st.subheader("📁 파일 업로드")
uploaded_file = st.file_uploader(
    "수신자 정보가 포함된 파일을 업로드하세요",
    type=['xlsx', 'csv']
)

if uploaded_file:
    try:
        # 파일 확장자에 따라 읽기
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        st.text("📋 파일 미리보기")
        st.dataframe(df)
        
        st.info(f"📊 총 수신자: {len(df)}명")
        
    except Exception as e:
        st.error(f"❌ 파일을 읽는 중 오류 발생: {e}")
        st.stop()
    
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
        value="{이름}님, 서류 보완 안내"
    )

    body_plain_text = st.text_area(
        "이메일 본문",
        height=350,
        value="""안녕하세요 **{이름}**님,

제출하신 서류를 검토한 결과 아래 항목에 대해 보완이 필요합니다.

▶ __내용__: {내용}
▶ __상세 내용__: {상세내용}

빠른 시일 내에 보완해주시기 바랍니다.

감사합니다."""
    )

    # HTML 변환 적용
    body_template = convert_to_html(body_plain_text)

    # 템플릿 변수 확인
    required_columns = {"이름", "이메일"}
    missing_required = required_columns - set(df.columns)
    
    if missing_required:
        st.error(f"❌ 필수 컬럼 누락: {', '.join(missing_required)}")
        st.stop()
    
    def extract_placeholders(template):
        return set(re.findall(r"{(\w+)}", template))

    placeholders = extract_placeholders(subject_template) | extract_placeholders(body_template)
    csv_columns = set(df.columns)
    missing_placeholders = placeholders - csv_columns

    if missing_placeholders:
        st.warning(f"⚠️ 업로드한 파일에 다음 컬럼이 없습니다: {', '.join(missing_placeholders)}")
        st.info("💡 누락된 컬럼은 빈 값으로 처리됩니다.")

    # 미리보기
    st.markdown("---")  
    st.subheader("🔍 이메일 미리보기 (첫 번째 수신자 기준)")
    first_row = df.iloc[0].fillna('').to_dict()
    
    try:
        preview_subject = subject_template.format(**first_row)
        preview_body = body_template.format(**first_row)

        st.text(f"📮 받는 사람: {first_row.get('이메일', 'N/A')}")
        st.text(f"📝 제목: {preview_subject}")
        st.markdown("**📄 본문:**")
        st.markdown(preview_body, unsafe_allow_html=True)

        with st.expander("🔍 HTML 원본 보기"):
            st.code(preview_body, language='html')

    except KeyError as e:
        st.error(f"⚠️ 템플릿 변수 오류: {e}")
        st.stop()

    # 발송 전 확인사항
    st.markdown("---")
    st.subheader("✅ 발송 전 확인사항")
    
    col1, col2 = st.columns(2)
    with col1:
        st.info(f"""
        **📊 발송 정보**
        - 총 수신자: {len(df)}명
        - 발신자: {sender_email}
        - 첨부파일: {'✅ 있음' if attachment_file else '❌ 없음'}
        """)
    
    with col2:
        st.info(f"""
        **⏱️ 예상 소요 시간**
        - 메일당 대기: {send_delay}초
        - 배치 대기: {batch_delay}초/{batch_size}통
        - 총 예상 시간: 약 {int((len(df) * send_delay + (len(df) // batch_size) * batch_delay) / 60)}분
        """)
    
    # 이메일 발송 버튼
    st.markdown("---")
    
    if not sender_email or '@' not in sender_email:
        st.error("❌ 사이드바에서 발신 이메일을 입력해주세요!")
        st.stop()
    
    if not sender_password:
        st.error("❌ 사이드바에서 비밀번호를 입력해주세요!")
        st.stop()
    
    if st.button("📨 이메일 발송 시작", type="primary"):
        try:
            
            # 진행 상황 표시
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # SMTP 연결
            server_name = "Gmail" if "gmail" in smtp_server else "네이버" if "naver" in smtp_server else "하이웍스"
            status_text.text(f"🔌 {server_name} 서버에 연결 중...")
            st.info(f"연결 시도: {smtp_server}:{smtp_port}")
            
            if smtp_port == 587:
                import ssl
                context = ssl.create_default_context()
                context.set_ciphers('DEFAULT@SECLEVEL=1')
                context.check_hostname = False
                context.verify_mode = ssl.CERT_NONE
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls(context=context)
            else:  # 465
                import ssl
                context = ssl.create_default_context()
                context.set_ciphers('DEFAULT@SECLEVEL=1')
                context.check_hostname = False
                context.verify_mode = ssl.CERT_NONE
                server = smtplib.SMTP_SSL(smtp_server, smtp_port, context=context)
            
            status_text.text(f"🔐 {server_name} 로그인 시도 중...")
            st.info(f"로그인 시도: {sender_email}")
            st.info(f"비밀번호 길이: {len(sender_password)}자, 특수문자(!): {'!' in sender_password}")
            
            # 여러 방식으로 로그인 시도
            login_success = False
            
            # 1. 원본 비밀번호로 시도
            try:
                server.login(sender_email, sender_password)
                login_success = True
                st.success("✅ 원본 비밀번호로 로그인 성공")
            except smtplib.SMTPAuthenticationError as e:
                st.warning(f"원본 비밀번호 실패: {str(e)[:100]}")
            
            if not login_success:
                raise Exception("로그인 실패")
                
            status_text.text("✅ 로그인 성공!")
            time.sleep(1)

            sent_count = 0
            failed_count = 0
            failed_list = []

            # 첨부파일 읽기
            attachment_bytes = None
            attachment_name = None
            if attachment_file is not None:
                attachment_bytes = attachment_file.read()
                attachment_name = attachment_file.name

            # 발송 시작
            total_count = len(df)
            
            for idx, (_, row) in enumerate(df.iterrows(), start=1):
                row_dict = row.fillna('').to_dict()
                
                try:
                    # 템플릿 치환
                    subject = subject_template.format(**row_dict)
                    body = body_template.format(**row_dict)
                    
                    # 메시지 생성
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

                    # 발송
                    server.send_message(msg)
                    sent_count += 1
                    
                    status_text.text(f"✅ [{idx}/{total_count}] {row_dict.get('이름', '수신자')} ({row_dict['이메일']}) 발송 완료")
                    
                    # 진행률 업데이트
                    progress_bar.progress(idx / total_count)
                    
                    # 배치 대기
                    if idx % batch_size == 0 and idx < total_count:
                        status_text.text(f"⏸️ 배치 대기 중... ({batch_delay}초)")
                        time.sleep(batch_delay)
                    else:
                        time.sleep(send_delay)
                    
                except Exception as e:
                    failed_count += 1
                    failed_list.append({
                        '이름': row_dict.get('이름', 'N/A'),
                        '이메일': row_dict.get('이메일', 'N/A'),
                        '오류': str(e)
                    })
                    status_text.text(f"❌ [{idx}/{total_count}] {row_dict.get('이름', '수신자')} 발송 실패")
                    time.sleep(send_delay)

            # 연결 종료
            server.quit()
            progress_bar.progress(1.0)
            
            # 결과 표시
            st.markdown("---")
            st.success(f"🎉 발송 완료! 성공: {sent_count}건 / 실패: {failed_count}건")
            
            if failed_list:
                st.error("❌ 발송 실패 목록:")
                failed_df = pd.DataFrame(failed_list)
                st.dataframe(failed_df)

        except smtplib.SMTPAuthenticationError as e:
            st.error(f"❌ 로그인 실패: {e}")
            st.error("이메일 또는 비밀번호를 확인해주세요.")
            st.info("💡 2단계 인증 사용 시 앱 비밀번호를 입력해야 합니다.")
            st.code(f"상세 오류: {str(e)}")
        except smtplib.SMTPException as e:
            st.error(f"❌ SMTP 오류: {e}")
            st.info("💡 하이웍스 관리자에게 SMTP 사용 권한을 확인해주세요.")
            st.code(f"상세 오류: {str(e)}")
        except Exception as e:
            st.error(f"❌ 예상치 못한 오류 발생: {e}")
            st.code(f"상세 오류: {str(e)}")

else:
    st.info("👆 파일을 업로드하여 시작하세요!")