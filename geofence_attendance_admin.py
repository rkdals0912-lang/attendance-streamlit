import math
from datetime import datetime
from io import BytesIO

import pandas as pd
import requests
import streamlit as st

try:
    from streamlit_geolocation import streamlit_geolocation
except ImportError:
    streamlit_geolocation = None


st.set_page_config(
    page_title="출퇴근 반경",
    page_icon="📍",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .block-container {max-width: 1180px; padding-top: 1.6rem; padding-bottom: 3rem;}
    .hero {
        padding: 1.6rem 1.8rem;
        border: 1px solid rgba(49, 51, 63, 0.15);
        border-radius: 18px;
        margin-bottom: 1rem;
    }
    .eyebrow {font-size: .78rem; font-weight: 800; letter-spacing: .12em; opacity: .65;}
    .hero h1 {margin: .25rem 0 .3rem 0; font-size: 2.15rem;}
    .hero p {margin: 0; opacity: .72; font-size: .98rem;}
    .step-card {
        border: 1px solid rgba(49, 51, 63, 0.14);
        border-radius: 16px;
        padding: 1rem 1.1rem;
        min-height: 152px;
    }
    .step-no {font-size: .78rem; font-weight: 800; opacity: .55; margin-bottom: .4rem;}
    .step-title {font-size: 1.15rem; font-weight: 800; margin-bottom: .45rem;}
    .muted {opacity: .68; font-size: .92rem; line-height: 1.55;}
    .status-card {
        border: 1px solid rgba(49, 51, 63, 0.14);
        border-radius: 16px;
        padding: 1.2rem 1.25rem;
        min-height: 132px;
    }
    .status-label {font-size: .8rem; font-weight: 700; opacity: .6;}
    .status-value {font-size: 1.65rem; font-weight: 800; margin-top: .35rem;}
    .status-sub {font-size: .88rem; opacity: .65; margin-top: .2rem;}
    div[data-testid="stMetric"] {
        border: 1px solid rgba(49, 51, 63, 0.14);
        padding: 14px 16px;
        border-radius: 14px;
    }
    div[data-testid="stDataFrame"] {border-radius: 14px; overflow: hidden;}
    .footer-note {opacity: .55; font-size: .82rem; margin-top: 1.5rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------- helpers ----------
def init_state():
    defaults = {
       "",
    "김도균",
    "안동인",
    "Suresh(수렛)",
    "Chamara(차마르)",
    "Nadun(나둔)",
    "Dinusha(디누샤)",
    "양준석",
    "황한빈",
    "Nuwan(누완)",
    "Puspika(푸스피카)",
    "Sardorbek(사도르벡)",
    "Prasad Kumara(프라삿 쿠마르)",
    "양민석",
    "Amila(아밀라)",
    "Maduranga(마두랑가)",
    "Mahesh(마헷)",
    "Jin Feng(김봉)",
    "구진서",
    "Punsara(푼사라)",
    "이종현",
    "Sanjaya(산자야)",
    "CHARITHA(자릿)",
    "송용원",
    "Ram(람)",
    "Damith(다미스)",
    "박경열",
    "Jayanath(자야나스)",
    "Athula(아뚤라)",
    "Samith(사미드)",
        "company_lat": None,
        "company_lon": None,
        "current_lat": None,
        "current_lon": None,
        "current_accuracy": None,
        "radius_m": 100,
        "attendance_status": "퇴근",
        "clock_in": None,
        "clock_out": None,
        "records": [],
        "last_inside": None,
        "auto_detect": True,
        "notify": True,
        "webhook": "",
        "is_admin": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def haversine_m(lat1, lon1, lat2, lon2):
    r = 6371000
    p1 = math.radians(lat1)
    p2 = math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def push_webhook(record):
    url = st.session_state.webhook.strip()
    if not url:
        return
    try:
        requests.post(url, json=record, timeout=5)
    except Exception:
        pass


def add_record(action, method="수동"):
    now = datetime.now()
    record = {
        "작업자": st.session_state.worker or "미선택",
        "구분": action,
        "일시": now.strftime("%Y-%m-%d %H:%M:%S"),
        "방식": method,
        "위도": st.session_state.current_lat,
        "경도": st.session_state.current_lon,
    }
    st.session_state.records.insert(0, record)
    push_webhook(record)
    return now


def clock_in(method="수동"):
    if st.session_state.attendance_status == "출근":
        return
    now = add_record("출근", method)
    st.session_state.clock_in = now
    st.session_state.clock_out = None
    st.session_state.attendance_status = "출근"


def clock_out(method="수동"):
    if st.session_state.attendance_status == "퇴근":
        return
    now = add_record("퇴근", method)
    st.session_state.clock_out = now
    st.session_state.attendance_status = "퇴근"


def excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="출퇴근기록")
    output.seek(0)
    return output.getvalue()


init_state()

# ---------- sidebar ----------
with st.sidebar:
    st.markdown("### 👤 사용자")
    st.caption("작업자를 선택하고 출퇴근 상태를 확인합니다.")

    worker_options = ["", "김민수", "이서연", "박준호", "최지우"]
    current_idx = worker_options.index(st.session_state.worker) if st.session_state.worker in worker_options else 0
    worker = st.selectbox(
        "작업자",
        worker_options,
        index=current_idx,
        format_func=lambda x: x or "이름을 선택하세요",
    )
    st.session_state.worker = worker

    st.divider()
    st.markdown("### 🔐 관리자")

    if not st.session_state.is_admin:
        admin_password = st.text_input(
            "관리자 비밀번호",
            type="password",
            placeholder="비밀번호 입력",
        )
        if st.button("관리자 로그인", use_container_width=True):
            secret_pw = st.secrets.get("ADMIN_PASSWORD", "")
            if not secret_pw:
                st.error("Streamlit Secrets에 ADMIN_PASSWORD가 설정되지 않았습니다.")
            elif admin_password == secret_pw:
                st.session_state.is_admin = True
                st.success("관리자 모드로 전환했습니다.")
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")
    else:
        st.success("✅ 관리자 모드")

        st.session_state.radius_m = st.slider(
            "인식 반경 (m)", 50, 500, int(st.session_state.radius_m), 10,
            help="회사 건물 크기와 GPS 오차를 고려해 100~200m를 권장합니다.",
        )
        st.session_state.auto_detect = st.toggle("자동 감지", value=st.session_state.auto_detect)
        st.session_state.notify = st.toggle("출퇴근 알림", value=st.session_state.notify)
        st.session_state.webhook = st.text_input(
            "구글 시트 웹훅 URL (선택)",
            value=st.session_state.webhook,
            placeholder="https://script.google.com/macros/s/.../exec",
        )
        st.caption("웹훅이 있으면 출퇴근 기록을 JSON POST로 전송합니다.")

        if st.button("관리자 로그아웃", use_container_width=True):
            st.session_state.is_admin = False
            st.rerun()

        st.divider()
        if st.button("🗑️ 기록 전체 초기화", use_container_width=True):
            st.session_state.records = []
            st.session_state.clock_in = None
            st.session_state.clock_out = None
            st.session_state.attendance_status = "퇴근"
            st.session_state.last_inside = None
            st.success("기록을 초기화했습니다.")

# ---------- header ----------
st.markdown(
    """
    <div class="hero">
      <div class="eyebrow">GEOFENCE ATTENDANCE</div>
      <h1>📍 출퇴근 반경</h1>
      <p>회사 위치와 현재 위치의 거리를 기준으로 출퇴근 상태를 간단하게 기록합니다.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------- onboarding steps ----------
step1, step2 = st.columns(2)
with step1:
    st.markdown(
        """
        <div class="step-card">
          <div class="step-no">1단계</div>
          <div class="step-title">이름을 선택하세요</div>
          <div class="muted">이 기기를 사용하는 작업자 본인의 이름을 선택하세요. 이후 출퇴근 기록에 이 이름이 함께 저장됩니다.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.session_state.worker:
        st.success(f"작업자: {st.session_state.worker}")
    else:
        st.warning("왼쪽 설정에서 작업자를 선택하세요.")

with step2:
    st.markdown(
        """
        <div class="step-card">
          <div class="step-no">2단계</div>
          <div class="step-title">회사 위치를 등록하세요</div>
          <div class="muted">현재 위치를 회사 위치로 저장하면 이 지점을 기준으로 설정 반경 안팎을 판단합니다.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.write("")

# ---------- geolocation ----------
geo_col, company_col = st.columns([1.15, 1])

with geo_col:
    st.subheader("📡 현재 위치")
    if streamlit_geolocation is not None:
        location = streamlit_geolocation()
        if location and location.get("latitude") is not None:
            st.session_state.current_lat = float(location["latitude"])
            st.session_state.current_lon = float(location["longitude"])
            st.session_state.current_accuracy = location.get("accuracy")
            st.success("브라우저에서 현재 위치를 확인했습니다.")
        else:
            st.info("위치 권한을 허용한 뒤 위치 버튼을 눌러주세요.")
    else:
        st.warning("`streamlit-geolocation`이 설치되지 않아 수동 위치 입력 모드로 동작합니다.")
        c1, c2 = st.columns(2)
        with c1:
            lat = st.number_input("현재 위도", value=37.5665, format="%.6f")
        with c2:
            lon = st.number_input("현재 경도", value=126.9780, format="%.6f")
        if st.button("현재 위치 적용", use_container_width=True):
            st.session_state.current_lat = float(lat)
            st.session_state.current_lon = float(lon)
            st.session_state.current_accuracy = None

    if st.session_state.current_lat is not None:
        st.caption(
            f"위도 {st.session_state.current_lat:.6f} · 경도 {st.session_state.current_lon:.6f}"
            + (f" · 정확도 약 {st.session_state.current_accuracy:.0f}m" if st.session_state.current_accuracy else "")
        )

with company_col:
    st.subheader("🏢 회사 위치")
    if st.session_state.company_lat is None:
        st.info("아직 회사 위치가 등록되지 않았습니다.")
    else:
        st.success(
            f"등록됨 · {st.session_state.company_lat:.6f}, {st.session_state.company_lon:.6f}"
        )

    if st.session_state.is_admin:
        if st.button("📌 현재 위치를 회사로 저장", type="primary", use_container_width=True):
            if st.session_state.current_lat is None:
                st.error("먼저 현재 위치를 확인해주세요.")
            else:
                st.session_state.company_lat = st.session_state.current_lat
                st.session_state.company_lon = st.session_state.current_lon
                st.success("현재 위치를 회사 위치로 저장했습니다.")
                st.rerun()
    else:
        st.caption("🔒 회사 위치 변경은 관리자만 가능합니다.")

# ---------- distance / auto detection ----------
distance = None
inside = None
if (
    st.session_state.current_lat is not None
    and st.session_state.company_lat is not None
):
    distance = haversine_m(
        st.session_state.current_lat,
        st.session_state.current_lon,
        st.session_state.company_lat,
        st.session_state.company_lon,
    )
    inside = distance <= st.session_state.radius_m

    if st.session_state.auto_detect:
        prev = st.session_state.last_inside
        if prev is False and inside is True:
            clock_in("자동")
            if st.session_state.notify:
                st.toast("회사 반경에 진입해 자동 출근 처리했습니다.", icon="✅")
        elif prev is True and inside is False:
            clock_out("자동")
            if st.session_state.notify:
                st.toast("회사 반경을 벗어나 자동 퇴근 처리했습니다.", icon="👋")
        st.session_state.last_inside = inside

# ---------- status cards ----------
st.divider()
st.subheader("오늘의 출퇴근 상태")

now = datetime.now()
status_text = "회사 반경 안" if inside is True else ("회사 반경 밖" if inside is False else "위치 확인 대기")
status_icon = "🟢" if inside is True else ("🟠" if inside is False else "⚪")

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(
        f"<div class='status-card'><div class='status-label'>현재 상태</div><div class='status-value'>{status_icon} {status_text}</div><div class='status-sub'>{f'{distance:.0f}m 거리' if distance is not None else '회사 위치와 현재 위치를 등록하세요'}</div></div>",
        unsafe_allow_html=True,
    )
with c2:
    in_text = st.session_state.clock_in.strftime("%H:%M:%S") if st.session_state.clock_in else "—"
    st.markdown(
        f"<div class='status-card'><div class='status-label'>출근</div><div class='status-value'>{in_text}</div><div class='status-sub'>{now.strftime('%Y-%m-%d')}</div></div>",
        unsafe_allow_html=True,
    )
with c3:
    out_text = st.session_state.clock_out.strftime("%H:%M:%S") if st.session_state.clock_out else "—"
    st.markdown(
        f"<div class='status-card'><div class='status-label'>퇴근</div><div class='status-value'>{out_text}</div><div class='status-sub'>현재 {st.session_state.attendance_status} 상태</div></div>",
        unsafe_allow_html=True,
    )
with c4:
    worked = "0시간 0분"
    if st.session_state.clock_in:
        end = st.session_state.clock_out or now
        mins = max(0, int((end - st.session_state.clock_in).total_seconds() // 60))
        worked = f"{mins // 60}시간 {mins % 60}분"
    st.markdown(
        f"<div class='status-card'><div class='status-label'>누적 근무</div><div class='status-value'>{worked}</div><div class='status-sub'>현재 세션 기준</div></div>",
        unsafe_allow_html=True,
    )

st.write("")
btn1, btn2 = st.columns(2)
with btn1:
    if st.button("✅ 수동 출근", use_container_width=True, disabled=st.session_state.attendance_status == "출근"):
        if not st.session_state.worker:
            st.error("먼저 작업자를 선택해주세요.")
        else:
            clock_in("수동")
            st.rerun()
with btn2:
    if st.button("👋 수동 퇴근", use_container_width=True, disabled=st.session_state.attendance_status == "퇴근"):
        clock_out("수동")
        st.rerun()

# ---------- recent records ----------
st.divider()
head1, head2 = st.columns([3, 1])
with head1:
    st.subheader("🧾 최근 기록")
    st.caption("최근 출퇴근 기록을 확인하고 엑셀로 내보낼 수 있습니다.")

records_df = pd.DataFrame(st.session_state.records)
if not records_df.empty:
    if st.session_state.is_admin:
        visible_df = records_df.copy()
        with head2:
            st.download_button(
                "⬇️ 전체 Excel",
                data=excel_bytes(visible_df),
                file_name=f"attendance_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        st.caption("관리자 모드: 전체 작업자의 기록을 표시합니다.")
    else:
        if st.session_state.worker:
            visible_df = records_df[records_df["작업자"] == st.session_state.worker].copy()
        else:
            visible_df = records_df.iloc[0:0].copy()
        with head2:
            st.button("🔒 관리자 전용", disabled=True, use_container_width=True)
        st.caption("일반 사용자는 선택한 작업자의 기록만 확인할 수 있습니다.")

    if not visible_df.empty:
        st.dataframe(
            visible_df[["작업자", "구분", "일시", "방식", "위도", "경도"]],
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("표시할 기록이 없습니다.")
else:
    with head2:
        st.button("⬇️ 엑셀로 내보내기", disabled=True, use_container_width=True)
    st.info("아직 기록이 없습니다. 수동 출퇴근을 사용하거나 회사 반경을 드나들면 기록이 쌓입니다.")

st.markdown(
    "<div class='footer-note'>※ 브라우저 위치 정보는 사용자의 위치 권한 허용이 필요하며 GPS 정확도에 따라 오차가 발생할 수 있습니다.</div>",
    unsafe_allow_html=True,
)
