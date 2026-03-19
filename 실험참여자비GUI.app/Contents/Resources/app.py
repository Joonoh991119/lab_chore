"""
실험참여자비 양식 자동 입력 GUI  —  CSNL lab_chore  v2.0
실행: .app 더블클릭  또는  streamlit run app.py
의존: pip install streamlit openpyxl streamlit-drawable-canvas Pillow
"""
import os, re, shutil, io, datetime, tempfile
import openpyxl
from openpyxl.drawing.image import Image as XLImage
import streamlit as st
from PIL import Image

WORK_DIR = os.environ.get("LAB_CHORE_DIR", os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_FILE = os.path.join(WORK_DIR, "실험참여자비 양식(중견).xlsx")

BANK_LIST = [
    "국민은행","기업은행","신한은행","우리은행","하나은행",
    "농협은행","SC제일은행","씨티은행","카카오뱅크","토스뱅크",
    "케이뱅크","부산은행","대구은행","경남은행","광주은행",
    "전북은행","제주은행","산업은행","수협은행",
    "새마을금고","신협","우체국","저축은행","기타",
]

# ── 페이지 설정 ──────────────────────────────────────────────────────────
st.set_page_config(page_title="실험참여자비 양식 입력", page_icon="🔬", layout="centered")
st.title("🔬 실험참여자비 양식 자동 입력")
st.caption(f"템플릿 폴더: `{WORK_DIR}`")

if not os.path.exists(TEMPLATE_FILE):
    st.error(
        f"⚠️ 템플릿 파일을 찾을 수 없습니다: `{TEMPLATE_FILE}`\n\n"
        "**이 .app 파일을 템플릿 Excel 파일들과 같은 폴더에 넣어 주세요.**"
    )
    st.stop()

# ── 1. 참가자 정보 ──────────────────────────────────────────────────────
st.subheader("👤 참가자 정보")
col1, col2 = st.columns(2)
with col1: name  = st.text_input("이름 *", placeholder="예: 홍길동")
with col2: inst  = st.text_input("소속 *", value="서울대학교")
col3, col4 = st.columns(2)
with col3: jid   = st.text_input("주민등록번호 *", placeholder="XXXXXX-XXXXXXX")
with col4: email = st.text_input("이메일 *", placeholder="user@snu.ac.kr")

# ── 2. 계좌 정보 ────────────────────────────────────────────────────────
st.subheader("🏦 계좌 정보")
col5, col6 = st.columns(2)
with col5: bank    = st.selectbox("은행명 *", options=BANK_LIST, index=2)
with col6: account = st.text_input("계좌번호 *", placeholder="110-545-811341")
holder = st.text_input("예금주", placeholder="비워두면 이름과 동일")

# ── 3. 활용 정보 ────────────────────────────────────────────────────────
st.subheader("📋 활용 정보")
col7, col8 = st.columns([1, 2])
with col7:
    amount_str = st.text_input("지급액 * (원)", placeholder="90000")
with col8:
    date_str = st.text_input("활용일자", placeholder="2026.03.19~03.20")

# 참여 시간
st.markdown("**⏱ 참여 시간**")
col9, col10, col11 = st.columns(3)
with col9:
    start_time = st.time_input("시작 시간", value=datetime.time(14, 0), step=1800)
with col10:
    end_time   = st.time_input("종료 시간", value=datetime.time(15, 0), step=1800)
with col11:
    # 총 참여 시간 (시간 단위, 소수점 가능)
    duration_h = st.number_input(
        "총 참여 시간 (시간)",
        min_value=0.5, max_value=24.0, value=1.0, step=0.5,
        help="세션 수 × 1회 참여시간. 예) 6회 × 1시간 = 6"
    )

# ── 4. 전자서명 ─────────────────────────────────────────────────────────
st.subheader("✍️ 전자서명")
st.caption("아래 캔버스에 서명하세요. 서명이 엑셀 파일의 '수령인 서명' 칸에 자동으로 삽입됩니다.")

# streamlit-drawable-canvas 임포트 (없으면 안내 메시지)
try:
    from streamlit_drawable_canvas import st_canvas

    sig_canvas = st_canvas(
        fill_color="rgba(255,255,255,1)",
        stroke_width=2,
        stroke_color="#111111",
        background_color="#ffffff",
        height=130,
        width=420,
        drawing_mode="freedraw",
        key="signature_canvas",
        display_toolbar=False,
    )

    col_sig1, col_sig2 = st.columns([1, 4])
    with col_sig1:
        clear_sig = st.button("🗑 지우기", key="clear_sig")
    if clear_sig:
        # canvas key를 바꿔서 초기화
        st.session_state["_sig_reset"] = not st.session_state.get("_sig_reset", False)
        st.rerun()

    sig_image_data = sig_canvas.image_data  # numpy RGBA array or None
    has_signature = (
        sig_image_data is not None
        and sig_image_data.sum() > 0
        and not (sig_image_data[:, :, :3] == 255).all()
    )

    if has_signature:
        st.success("✅ 서명이 입력되었습니다.")
    else:
        st.info("ℹ️ 서명을 그려주세요 (선택사항 — 비워두면 서명 없이 저장됩니다).")

    CANVAS_AVAILABLE = True

except ImportError:
    st.warning(
        "⚠️ `streamlit-drawable-canvas` 패키지가 설치되지 않았습니다.\n\n"
        "터미널에서 다음을 실행하세요:\n```\npip install streamlit-drawable-canvas Pillow\n```\n"
        "설치 후 앱을 재시작하면 서명 기능이 활성화됩니다."
    )
    sig_image_data = None
    has_signature = False
    CANVAS_AVAILABLE = False

# ── 5. 저장 ─────────────────────────────────────────────────────────────
st.divider()
if st.button("✅  양식 저장", type="primary", use_container_width=True):

    # 유효성 검사
    errors = []
    if not name.strip():    errors.append("이름을 입력하세요.")
    if not inst.strip():    errors.append("소속을 입력하세요.")
    if not jid.strip():     errors.append("주민등록번호를 입력하세요.")
    elif not re.match(r"^\d{6}-\d{7}$", jid.strip()):
        errors.append("주민등록번호 형식: XXXXXX-XXXXXXX")
    if not email.strip():   errors.append("이메일을 입력하세요.")
    if not account.strip(): errors.append("계좌번호를 입력하세요.")
    if not amount_str.strip(): errors.append("지급액을 입력하세요.")
    else:
        try:    amount = int(amount_str.strip().replace(",", ""))
        except: errors.append("지급액은 숫자로 입력하세요.")
    if end_time <= start_time:
        errors.append("종료 시간이 시작 시간보다 늦어야 합니다.")

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

        # 참가자 정보
        ws["B16"] = name.strip()
        ws["D16"] = inst.strip()
        ws["E16"] = jid.strip()
        ws["F16"] = email.strip()
        ws["G16"] = bank
        ws["I16"] = account.strip()
        ws["L16"] = final_holder
        ws["D19"] = amount

        # 활용일자
        if date_str.strip():
            ws["C10"] = date_str.strip()

        # 참여 시간 (G10 = 시작, I10 = 종료, B11 = 총 시간)
        ws["G10"] = start_time
        ws["I10"] = end_time
        # B11: 총 참여 시간 (시간, 정수 or 소수)
        ws["B11"] = int(duration_h) if duration_h == int(duration_h) else duration_h

        # 전자서명 이미지 삽입 (B17 = "수령인 서명" 칸)
        if has_signature and sig_image_data is not None:
            import numpy as np
            # RGBA numpy → PIL → 흰 배경 합성 → PNG bytes
            sig_pil  = Image.fromarray(sig_image_data.astype("uint8"), "RGBA")
            bg       = Image.new("RGB", sig_pil.size, (255, 255, 255))
            bg.paste(sig_pil, mask=sig_pil.split()[3])

            sig_bytes = io.BytesIO()
            bg.save(sig_bytes, format="PNG")
            sig_bytes.seek(0)

            xl_img        = XLImage(sig_bytes)
            xl_img.width  = 160   # 픽셀 너비
            xl_img.height = 55    # 픽셀 높이
            xl_img.anchor = "B17"
            ws.add_image(xl_img)

        wb.save(output_path)

        st.success(f"✅ 저장 완료: **{output_name}**")
        st.info(f"📂 `{output_path}`")

        with st.expander("📄 입력 내용 확인"):
            st.table({
                "항목": ["이름","소속","주민등록번호","이메일","은행명",
                         "계좌번호","예금주","지급액","활용일자",
                         "시작시간","종료시간","총참여시간","서명"],
                "값":   [name.strip(), inst.strip(), jid.strip(), email.strip(), bank,
                         account.strip(), final_holder, f"{amount:,}원",
                         date_str.strip() or "(미입력)",
                         start_time.strftime("%H:%M"), end_time.strftime("%H:%M"),
                         f"{duration_h}시간",
                         "✅ 서명 포함" if has_signature else "—"],
            })

    except Exception as e:
        st.error(f"저장 실패: {e}")
        import traceback; st.code(traceback.format_exc())
