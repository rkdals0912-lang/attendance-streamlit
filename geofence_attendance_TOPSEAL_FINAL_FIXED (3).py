
import base64
import hashlib
import hmac
import json
import math
import time
import uuid
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

try:
    from streamlit_geolocation import streamlit_geolocation
except ImportError:
    streamlit_geolocation = None

try:
    import extra_streamlit_components as stx
except ImportError:
    stx = None

st.set_page_config(
    page_title="TOPSEAL 출퇴근 관리",
    page_icon="📍",
    layout="wide",
    initial_sidebar_state="expanded",
)

COOKIE_NAME = "topseal_device_binding"

st.markdown("""
<style>
.block-container {max-width:1180px;padding-top:1.1rem;padding-bottom:3rem;}
.brand {font-size:2.2rem;font-weight:900;letter-spacing:.08em;}
.hero {padding:1.35rem 1.6rem;border:1px solid rgba(49,51,63,.14);border-radius:18px;margin:.6rem 0 1rem;}
.hero h1 {margin:.15rem 0 .3rem;font-size:2rem;}
.hero p {margin:0;opacity:.7;}
.card {border:1px solid rgba(49,51,63,.14);border-radius:16px;padding:1rem 1.1rem;}
.label {font-size:.8rem;font-weight:700;opacity:.6;}
.value {font-size:1.5rem;font-weight:800;margin-top:.3rem;}
.sub {font-size:.86rem;opacity:.65;margin-top:.25rem;}
</style>
""", unsafe_allow_html=True)

def secret(name, default=None):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default

def init_state():
    defaults = {
        "worker": "",
        "group": "",
        "device_id": "",
        "registered": False,
        "validated": False,
        "admin": False,
        "lat": None,
        "lon": None,
        "accuracy": None,
        "status": "퇴근",
        "clock_in": "",
        "clock_out": "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

def api_url():
    return str(secret("GOOGLE_APPS_SCRIPT_URL", "https://script.google.com/macros/s/AKfycbwf75Jdw9UJLlUGmCrfQToIbAwaUib8ldnKeQc2nS1JwTMwH4GfYfwMFOJhgpEKGvsq/exec")).strip()

def api_token():
    return str(secret("API_TOKEN", "topseal_attendance_2026_secure_api_9K7mP4x2")).strip()

def device_secret():
    return str(secret("DEVICE_SECRET", "topseal_device_binding_2026_X8qL3mN7vR5kT2pA")).strip()

def company_settings():
    try:
        return (
            float(secret("COMPANY_LAT", 36.9366)),
            float(secret("COMPANY_LON", 127.5348)),
            int(secret("GEOFENCE_RADIUS_M", 150)),
        )
    except Exception:
        return None, None, 150

def encode_binding(payload):
    raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode()
    raw64 = base64.urlsafe_b64encode(raw).decode().rstrip("=")
    sig = hmac.new(device_secret().encode(), raw64.encode(), hashlib.sha256).hexdigest()
    return f"{raw64}.{sig}"

def decode_binding(token):
    if not token or "." not in token or not device_secret():
        return None
    try:
        raw64, sig = token.rsplit(".", 1)
        expected = hmac.new(device_secret().encode(), raw64.encode(), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(sig, expected):
            return None
        padded = raw64 + "=" * (-len(raw64) % 4)
        return json.loads(base64.urlsafe_b64decode(padded).decode())
    except Exception:
        return None

def api_post(action, **payload):
    if not api_url() or not api_token():
        return {"ok": False, "message": "Google Sheets API 설정이 없습니다."}
    try:
        r = requests.post(
            api_url(),
            json={"action": action, "token": api_token(), **payload},
            timeout=15,
        )
        r.raise_for_status()
        data = r.json()
        return data if isinstance(data, dict) else {"ok": False, "message": "서버 응답 오류"}
    except Exception as e:
        return {"ok": False, "message": f"서버 연결 실패: {e}"}

def haversine_m(lat1, lon1, lat2, lon2):
    r = 6371000
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp/2)**2 + math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2
    return 2*r*math.asin(math.sqrt(a))

def excel_bytes(df):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="근태기록")
    out.seek(0)
    return out.getvalue()

cookie_manager = stx.CookieManager() if stx is not None else None

# 쿠키 복원
if cookie_manager and device_secret():
    cookies = cookie_manager.get_all() or {}
    binding = decode_binding(cookies.get(COOKIE_NAME, ""))
    if binding:
        st.session_state.group = binding.get("group", "")
        st.session_state.worker = binding.get("worker", "")
        st.session_state.device_id = binding.get("device_id", "")
        st.session_state.registered = bool(st.session_state.device_id)

# 서버 검증
if st.session_state.registered:
    check = api_post(
        "validate_device",
        worker=st.session_state.worker,
        device_id=st.session_state.device_id,
    )
    st.session_state.validated = bool(check.get("ok"))

with st.sidebar:
    st.markdown("### 👤 사용자")
    if st.session_state.registered:
        st.write(f"**작업반:** {st.session_state.group}")
        st.write(f"**작업자:** {st.session_state.worker}")
        if st.session_state.validated:
            st.success("✅ 등록 기기")
        else:
            st.error("❌ 등록 확인 실패")
    else:
        st.caption("최초 1회 기기 등록이 필요합니다.")

    st.divider()
    st.markdown("### 🔐 관리자")

    if not st.session_state.admin:
        pw = st.text_input("관리자 비밀번호", type="password")
        if st.button("관리자 로그인", use_container_width=True):
            expected = str(secret("ADMIN_PASSWORD", ""))
            if expected and hmac.compare_digest(pw, expected):
                st.session_state.admin = True
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")
    else:
        st.success("관리자 모드")

        with st.expander("👥 직원 관리", expanded=False):
            tab1, tab2, tab3 = st.tabs(["신규 추가", "수정/퇴사", "기기 해제"])

            with tab1:
                new_group = st.selectbox("작업반", ["제조반", "포장반"], key="new_group")
                new_name = st.text_input("이름", key="new_name")
                new_pin = st.text_input("휴대폰 마지막 4자리", max_chars=4, key="new_pin")
                if st.button("➕ 직원 추가", use_container_width=True):
                    result = api_post(
                        "add_worker",
                        group=new_group,
                        worker=new_name.strip(),
                        pin=new_pin.strip(),
                    )
                    st.success(result.get("message")) if result.get("ok") else st.error(result.get("message"))

            with tab2:
                workers_result = api_post("list_workers", include_inactive=True)
                workers = workers_result.get("workers", []) if workers_result.get("ok") else []
                names = [w.get("작업자") for w in workers]
                selected = st.selectbox("직원", names, index=None, placeholder="직원 선택", key="edit_worker")
                if selected:
                    row = next((w for w in workers if w.get("작업자") == selected), {})
                    edit_group = st.selectbox("작업반 변경", ["제조반", "포장반"],
                                              index=0 if row.get("작업반") == "제조반" else 1)
                    edit_pin = st.text_input("새 휴대폰 마지막 4자리(변경 시만)", max_chars=4)
                    c1, c2 = st.columns(2)
                    with c1:
                        if st.button("직원 정보 수정", use_container_width=True):
                            result = api_post(
                                "update_worker",
                                worker=selected,
                                group=edit_group,
                                pin=edit_pin.strip(),
                            )
                            st.success(result.get("message")) if result.get("ok") else st.error(result.get("message"))
                    with c2:
                        if st.button("퇴사 처리", use_container_width=True):
                            result = api_post("deactivate_worker", worker=selected)
                            st.success(result.get("message")) if result.get("ok") else st.error(result.get("message"))

            with tab3:
                active = api_post("list_workers", include_inactive=False)
                workers = active.get("workers", []) if active.get("ok") else []
                active_names = [w.get("작업자") for w in workers if w.get("작업자")]
                target = st.selectbox("기기 등록 해제할 직원", active_names, index=None)
                if st.button("📱 기기 등록 해제", use_container_width=True, disabled=not target):
                    result = api_post("unregister_worker", worker=target)
                    st.success(result.get("message")) if result.get("ok") else st.error(result.get("message"))

        if st.button("관리자 로그아웃", use_container_width=True):
            st.session_state.admin = False
            st.rerun()

logo = Path("topseal_logo.png")
if logo.exists():
    _, c, _ = st.columns([1, 1.15, 1])
    with c:
        st.image(str(logo), use_container_width=True)
else:
    st.markdown('<div class="brand">TOPSEAL</div>', unsafe_allow_html=True)

st.markdown("""
<div class="hero">
<div class="eyebrow">SECURE GEOFENCE ATTENDANCE</div>
<h1>📍 출퇴근 관리</h1>
<p>등록 기기, 중앙 DB, 회사 GPS 반경을 함께 확인합니다.</p>
</div>
""", unsafe_allow_html=True)

required = []
if stx is None: required.append("extra-streamlit-components")
if streamlit_geolocation is None: required.append("streamlit-geolocation")
if not api_url(): required.append("GOOGLE_APPS_SCRIPT_URL")
if not api_token(): required.append("API_TOKEN")
if not device_secret(): required.append("DEVICE_SECRET")

company_lat, company_lon, radius_m = company_settings()
if company_lat is None or company_lon is None:
    required.append("COMPANY_LAT / COMPANY_LON")

if required:
    st.error("설정 필요: " + ", ".join(required))
    st.stop()

# 최초 등록
if not st.session_state.registered:
    st.subheader("📱 최초 1회 기기 등록")
    st.info("등록 후에는 매일 이름을 다시 선택하지 않습니다.")

    active = api_post("list_workers", include_inactive=False)

    if not active.get("ok"):
        st.error(f"직원 목록 API 오류: {active.get('message', '알 수 없는 오류')}")
        with st.expander("API 응답 확인"):
            st.json(active)
        st.stop()

    workers = active.get("workers", [])
    workers = [w for w in workers if str(w.get("상태", "")).strip().upper() == "ACTIVE"]

    if not workers:
        st.error("직원관리 시트에서 ACTIVE 직원을 불러오지 못했습니다.")
        st.caption("직원관리 시트의 헤더가 작업반 | 작업자 | PIN | 상태 순서인지, 상태가 ACTIVE인지 확인하세요.")
        with st.expander("API 응답 확인"):
            st.json(active)
        st.stop()

    st.success(f"직원 목록 연결 완료 · {len(workers)}명")

    group = st.selectbox(
        "1. 작업반",
        ["제조반", "포장반"],
        index=None,
        placeholder="작업반 선택",
    )

    group_workers = [
        w for w in workers
        if str(w.get("작업반", "")).strip() == str(group or "").strip()
    ]

    worker_names = [
        str(w.get("작업자", "")).strip()
        for w in group_workers
        if str(w.get("작업자", "")).strip()
    ]

    worker = st.selectbox(
        "2. 이름",
        worker_names,
        index=None,
        placeholder="본인 이름 선택" if group else "먼저 작업반 선택",
        disabled=not group,
    )

    if group and not worker_names:
        st.warning(f"{group}에 ACTIVE 상태로 등록된 직원이 없습니다.")

    phone4 = st.text_input(
        "3. 휴대폰 번호 마지막 4자리",
        type="password",
        max_chars=4,
    )
    agree = st.checkbox("이 휴대폰을 본인의 출퇴근 전용 기기로 등록합니다.")

    can_register = bool(
        worker
        and phone4.isdigit()
        and len(phone4) == 4
        and agree
    )

    if st.button(
        "🔒 이 휴대폰 등록",
        type="primary",
        use_container_width=True,
        disabled=not can_register,
    ):
        new_device_id = str(uuid.uuid4())
        result = api_post(
            "register_device",
            group=group,
            worker=worker,
            pin=phone4,
            device_id=new_device_id,
        )

        if result.get("ok"):
            payload = result.get("binding") or {
                "group": group,
                "worker": worker,
                "device_id": new_device_id,
            }
            token = encode_binding(payload)

            if cookie_manager is None:
                st.error("기기 저장 기능을 사용할 수 없습니다. extra-streamlit-components 설치를 확인하세요.")
                st.stop()

            cookie_manager.set(
                COOKIE_NAME,
                token,
                expires_at=datetime.now() + timedelta(days=365),
            )
            st.success("등록 완료")
            time.sleep(0.7)
            st.rerun()
        else:
            st.error(result.get("message", "기기 등록에 실패했습니다."))

    st.stop()

if not st.session_state.validated:
    st.error("중앙 DB에서 이 기기를 확인하지 못했습니다. 관리자에게 문의하세요.")
    st.stop()

# 서버 상태
state = api_post("get_state", worker=st.session_state.worker, device_id=st.session_state.device_id)
if state.get("ok"):
    st.session_state.status = state.get("status", "퇴근")
    st.session_state.clock_in = state.get("clock_in", "")
    st.session_state.clock_out = state.get("clock_out", "")

st.markdown(f"""
<div class="card">
<div class="label">등록 사용자</div>
<div class="value">👤 {st.session_state.worker}</div>
<div class="sub">{st.session_state.group} · 등록 기기 인증 완료</div>
</div>
""", unsafe_allow_html=True)

st.subheader("📡 현재 위치")
location = streamlit_geolocation()
if location and location.get("latitude") is not None:
    st.session_state.lat = float(location["latitude"])
    st.session_state.lon = float(location["longitude"])
    st.session_state.accuracy = location.get("accuracy")
    st.success("현재 위치 확인 완료")
else:
    st.info("위치 권한을 허용한 뒤 위치 확인 버튼을 눌러주세요.")

distance = None
inside = None
if st.session_state.lat is not None:
    distance = haversine_m(
        st.session_state.lat, st.session_state.lon,
        company_lat, company_lon
    )
    inside = distance <= radius_m

st.divider()
st.subheader("오늘의 출퇴근 상태")

loc = "회사 반경 안" if inside is True else ("회사 반경 밖" if inside is False else "위치 대기")
icon = "🟢" if inside is True else ("🟠" if inside is False else "⚪")

c1, c2, c3 = st.columns(3)
with c1:
    st.markdown(f"<div class='card'><div class='label'>현재 위치</div><div class='value'>{icon} {loc}</div><div class='sub'>{f'회사까지 {distance:.0f}m' if distance is not None else 'GPS 확인 필요'} · 반경 {radius_m}m</div></div>", unsafe_allow_html=True)
with c2:
    st.markdown(f"<div class='card'><div class='label'>출근</div><div class='value'>{st.session_state.clock_in or '—'}</div><div class='sub'>오늘 출근 시각</div></div>", unsafe_allow_html=True)
with c3:
    st.markdown(f"<div class='card'><div class='label'>퇴근</div><div class='value'>{st.session_state.clock_out or '—'}</div><div class='sub'>현재 {st.session_state.status}</div></div>", unsafe_allow_html=True)

b1, b2 = st.columns(2)
with b1:
    if st.button("✅ 출근하기", type="primary", use_container_width=True,
                 disabled=inside is not True or st.session_state.status == "출근"):
        result = api_post(
            "attendance",
            worker=st.session_state.worker,
            device_id=st.session_state.device_id,
            attendance_action="출근",
            latitude=st.session_state.lat,
            longitude=st.session_state.lon,
            distance_m=round(distance,1),
        )
        st.success(result.get("message")) if result.get("ok") else st.error(result.get("message"))
        if result.get("ok"): st.rerun()

with b2:
    if st.button("👋 퇴근하기", use_container_width=True,
                 disabled=inside is not True or st.session_state.status != "출근"):
        result = api_post(
            "attendance",
            worker=st.session_state.worker,
            device_id=st.session_state.device_id,
            attendance_action="퇴근",
            latitude=st.session_state.lat,
            longitude=st.session_state.lon,
            distance_m=round(distance,1),
        )
        st.success(result.get("message")) if result.get("ok") else st.error(result.get("message"))
        if result.get("ok"): st.rerun()

if inside is False:
    st.warning("회사 반경 밖에서는 출퇴근 기록을 할 수 없습니다.")

st.divider()
st.subheader("🧾 근태 현황")

if st.session_state.admin:
    group_view = st.radio("반 선택", ["제조반", "포장반"], horizontal=True)
    date_text = st.date_input("조회 날짜", value=datetime.now().date()).strftime("%Y-%m-%d")
    result = api_post("daily_roster", group=group_view, date=date_text)
else:
    result = api_post("list_records", worker=st.session_state.worker, limit=100)

if result.get("ok"):
    rows = result.get("records", [])
    df = pd.DataFrame(rows)
    if not df.empty:
        st.dataframe(df, use_container_width=True, hide_index=True)
        if st.session_state.admin:
            st.download_button(
                "⬇️ Excel 다운로드",
                data=excel_bytes(df),
                file_name=f"TOPSEAL_{group_view}_{date_text}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.info("표시할 기록이 없습니다.")
else:
    st.error(result.get("message"))

st.caption(
    "미출근 직원은 자동 결근 처리하지 않고 Blank로 유지합니다."
)
