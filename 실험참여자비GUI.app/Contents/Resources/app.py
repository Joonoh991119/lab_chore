"""
실험참여자비 양식 자동 입력 GUI  —  CSNL lab_chore
실행: .app 더블클릭  또는  streamlit run app.py
의존: pip install streamlit openpyxl
"""
import os, re, shutil
import openpyxl
import streamlit as st

WORK_DIR = os.environ.get("LAB_CHORE_DIR", os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_FILE = os.path.join(WORK_DIR, "실험참여자비 양식(중견).xlsx")

BANK_LIST = [
    "국민은행","기업은행","신한은행","우리은행","하나은행",
    "농협은행","SC제일은행","씨티은행","카카오뱅크","토스뱅크",
    "케이뱅크","부산은행","대구은행","경남은행","광주은행",
    "전북은행","제주은행","산업은행","수협은행",
    "새마을금고","신협","우체국","저축은행","기타",
]

st.set_page_config(page_title="실험참여자비 양식 입력", page_icon="🔬", layout="centered")
st.title("🔬 실험참여자비 양식 자동 입력")
st.caption(f"템플릿 폴더: `{WORK_DIR}`")

if not os.path.exists(TEMPLATE_FILE):
    st.error(
        f"⚠️ 템플릿 파일을 찾을 수 없습니다: `{TEMPLATE_FILE}`\n\n"
        "**이 .app 파일을 템플릿 Excel 파일들과 같은 폴더에 넣어 주세요.**"
    )
    st.stop()

st.subheader("👤 참가자 정보")
col1, col2 = st.columns(2)
with col1: name  = st.text_input("이름 *", placeholder="예: 홍길동")
with col2: inst  = st.text_input("소속 *", value="서울대학교")
col3, col4 = st.columns(2)
with col3: jid   = st.text_input("주민등록번호 *", placeholder="XXXXXX-XXXXXXX")
with col4: email = st.text_input("이메일 *", placeholder="user@snu.ac.kr")

st.subheader("🏦 계좌 정보")
col5, col6 = st.columns(2)
with col5: bank    = st.selectbox("은행명 *", options=BANK_LIST, index=2)
with col6: account = st.text_input("계좌번호 *", placeholder="110-545-811341")
holder = st.text_input("예금주", placeholder="비워두면 이름과 동일")

st.subheader("📋 활용 정보")
col7, col8 = st.columns([1, 2])
with col7: amount_str = st.text_input("지급액 * (원)", placeholder="90000")
with col8: date_str   = st.text_input("활용일자 (선택)", placeholder="2026.03.19~03.20")

st.divider()

if st.button("✅  양식 저장", type="primary", use_container_width=True):
    errors = []
    if not name.strip():   errors.append("이름을 입력하세요.")
    if not inst.strip():   errors.append("소속을 입력하세요.")
    if not jid.strip():    errors.append("주민등록번호를 입력하세요.")
    elif not re.match(r"^\d{6}-\d{7}$", jid.strip()):
        errors.append("주민등록번호 형식: XXXXXX-XXXXXXX")
    if not email.strip():  errors.append("이메일을 입력하세요.")
    if not account.strip(): errors.append("계좌번호를 입력하세요.")
    if not amount_str.strip(): errors.append("지급액을 입력하세요.")
    else:
        try:    amount = int(amount_str.strip().replace(",", ""))
        except: errors.append("지급액은 숫자로 입력하세요.")

    if errors:
        for e in errors: st.error(e)
        st.stop()

    amount = int(amount_str.strip().replace(",", ""))
    final_holder = holder.strip() or name.strip()
    output_name  = f"실험참여자비 양식_{name.strip()}.xlsx"
    output_path  = os.path.join(WORK_DIR, output_name)

    try:
        shutil.copy2(TEMPLATE_FILE, output_path)
        wb = openpyxl.load_workbook(output_path)
        ws = wb.active
        ws["B16"] = name.strip()
        ws["D16"] = inst.strip()
        ws["E16"] = jid.strip()
        ws["F16"] = email.strip()
        ws["G16"] = bank
        ws["I16"] = account.strip()
        ws["L16"] = final_holder
        ws["D19"] = amount
        if date_str.strip(): ws["C10"] = date_str.strip()
        wb.save(output_path)
        st.success(f"✅ 저장 완료: **{output_name}**")
        st.info(f"📂 `{output_path}`")
        with st.expander("📄 입력 내용 확인"):
            st.table({
                "항목": ["이름","소속","주민등록번호","이메일","은행명","계좌번호","예금주","지급액","활용일자"],
                "값":   [name.strip(),inst.strip(),jid.strip(),email.strip(),
                         bank,account.strip(),final_holder,f"{amount:,}원",
                         date_str.strip() or "(미입력)"],
            })
    except Exception as e:
        st.error(f"저장 실패: {e}")
