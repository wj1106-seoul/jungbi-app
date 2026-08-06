# -*- coding: utf-8 -*-
"""
app.py - 정비사업 입찰공고 수집기 사내 웹앱 (Streamlit)

실행 방법:
    streamlit run app.py
그 다음 같은 사내망에 있는 다른 PC에서 http://<이 PC의 IP>:8501 로 접속하면 됩니다.
"""
import os
import json
from datetime import datetime
from pathlib import Path

import streamlit as st
import pandas as pd

import collector_core as core

st.set_page_config(page_title="정비사업 입찰공고 수집기", page_icon="🏗️", layout="wide")

BASE_DIR = Path(__file__).resolve().parent

# ------------------------------------------------------------------
# 디자인 (커스텀 CSS)
# ------------------------------------------------------------------
st.markdown(
    """
    <style>
    .stApp { background-color: #0E1117; }
    .main .block-container { padding-top: 2rem; max-width: 1200px; }

    .app-header {
        background: linear-gradient(135deg, #1F3864 0%, #2E5090 100%);
        padding: 28px 32px;
        border-radius: 14px;
        margin-bottom: 24px;
        box-shadow: 0 4px 18px rgba(0,0,0,0.25);
    }
    .app-header h1 {
        color: #FFFFFF; font-size: 28px; margin: 0 0 6px 0; font-weight: 700;
    }
    .app-header p {
        color: #CBD8EE; font-size: 14px; margin: 0;
    }

    div[data-testid="stMetric"] {
        background-color: #1A2233;
        border: 1px solid #2A3550;
        border-radius: 10px;
        padding: 14px 18px;
    }
    div[data-testid="stMetricLabel"] { color: #9FB3D1; }
    div[data-testid="stMetricValue"] { color: #FFFFFF; }

    .stButton > button[kind="primary"] {
        background-color: #C0392B;
        border: none;
        font-weight: 600;
        border-radius: 8px;
    }
    .stButton > button[kind="primary"]:hover {
        background-color: #A93226;
    }

    section[data-testid="stSidebar"] {
        background-color: #131722;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ------------------------------------------------------------------
# 세션 상태 초기화
# ------------------------------------------------------------------
if "cfg" not in st.session_state:
    st.session_state.cfg = core.load_config()
if "log_lines" not in st.session_state:
    st.session_state.log_lines = []
if "last_result" not in st.session_state:
    st.session_state.last_result = None


def push_log(msg: str):
    st.session_state.log_lines.append(msg)


# ------------------------------------------------------------------
# 사이드바 - 실행 옵션 + 필터 편집
# ------------------------------------------------------------------
st.sidebar.header("⚙️ 실행 옵션")

def get_service_key() -> str:
    """로컬(.env)과 Streamlit Cloud(secrets) 양쪽 모두 지원."""
    key = os.environ.get("SERVICE_KEY", "")
    if key:
        return key
    try:
        return st.secrets.get("SERVICE_KEY", "")
    except Exception:
        return ""


service_key = get_service_key()
if not service_key:
    st.sidebar.error(
        "SERVICE_KEY가 설정되지 않았습니다.\n"
        "로컬 실행: .env 파일 확인 / Streamlit Cloud: Secrets 설정 확인"
    )

period_mode = st.sidebar.radio("조회 기간", ["직전 영업일만", "최근 N일"], index=0)
yesterday_only = period_mode == "직전 영업일만"
days_back = 1
if not yesterday_only:
    days_back = st.sidebar.number_input("최근 며칠", min_value=1, max_value=30, value=3)

download_attachments = st.sidebar.checkbox("첨부파일(공고문/지침서) 다운로드", value=True)

st.sidebar.divider()
st.sidebar.header("🔧 필터 규칙 편집")
st.sidebar.caption("한 줄에 키워드 하나씩 입력하세요. 저장하면 즉시 다음 실행부터 반영됩니다.")

with st.sidebar.expander("발주기관 포함 키워드", expanded=False):
    inst_text = st.text_area(
        "institution_include", value="\n".join(st.session_state.cfg["include_institution_keywords"]),
        height=150, label_visibility="collapsed",
    )

with st.sidebar.expander("공고명 포함 키워드", expanded=False):
    title_inc_text = st.text_area(
        "title_include", value="\n".join(st.session_state.cfg["include_title_keywords"]),
        height=150, label_visibility="collapsed",
    )

with st.sidebar.expander("공고명 제외 키워드", expanded=False):
    title_exc_text = st.text_area(
        "title_exclude", value="\n".join(st.session_state.cfg["exclude_title_keywords"]),
        height=150, label_visibility="collapsed",
    )

if st.sidebar.button("💾 필터 규칙 저장", use_container_width=True):
    cfg = st.session_state.cfg
    cfg["include_institution_keywords"] = [l.strip() for l in inst_text.splitlines() if l.strip()]
    cfg["include_title_keywords"] = [l.strip() for l in title_inc_text.splitlines() if l.strip()]
    cfg["exclude_title_keywords"] = [l.strip() for l in title_exc_text.splitlines() if l.strip()]
    core.save_config(cfg)
    st.session_state.cfg = cfg
    st.sidebar.success("저장되었습니다.")

# ------------------------------------------------------------------
# 메인 화면
# ------------------------------------------------------------------
st.markdown(
    """
    <div class="app-header">
        <h1>🏗️ 정비사업 입찰공고 수집기</h1>
        <p>누리장터 민간입찰공고 중 정비사업 관련 기술용역 공고만 자동으로 걸러서 보여줍니다.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

col_run, col_status = st.columns([1, 4])
with col_run:
    run_clicked = st.button("🚀 지금 수집하기", type="primary", use_container_width=True)

status_placeholder = st.empty()
log_box = st.expander("실행 로그", expanded=run_clicked)

if run_clicked:
    st.session_state.log_lines = []
    log_area = log_box.empty()

    def progress_cb(msg: str):
        push_log(msg)
        log_area.code("\n".join(st.session_state.log_lines[-200:]))

    with st.spinner("공고를 조회하고 필터링하는 중입니다..."):
        result = core.run_pipeline(
            service_key=service_key,
            cfg=st.session_state.cfg,
            yesterday_only=yesterday_only,
            days_back=days_back,
            download_attachments=download_attachments,
            progress_cb=progress_cb,
        )
    st.session_state.last_result = result

result = st.session_state.last_result

if result is None:
    status_placeholder.info("아직 실행하지 않았습니다. '지금 수집하기'를 눌러 시작하세요.")
elif not result.ok:
    status_placeholder.warning(result.message)
else:
    m1, m2, m3 = st.columns(3)
    m1.metric("기준일", f"{result.title_date:%Y-%m-%d}")
    m2.metric("API 조회 건수", f"{result.raw_count:,}건")
    m3.metric("필터 통과 건수", f"{len(result.filtered_df):,}건")

    st.write("")
    df_display = result.filtered_df.drop(columns=["폴더명"], errors="ignore")
    st.dataframe(df_display, use_container_width=True, height=480)

    if result.excel_path and os.path.exists(result.excel_path):
        with open(result.excel_path, "rb") as f:
            st.download_button(
                "📥 엑셀 파일 다운로드",
                data=f.read(),
                file_name=os.path.basename(result.excel_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

st.divider()
with st.expander("📜 최근 실행 로그 파일 보기"):
    log_files = sorted(core.LOG_DIR.glob("collector_*.log"), reverse=True)
    if not log_files:
        st.caption("아직 로그 파일이 없습니다.")
    else:
        chosen = st.selectbox("로그 파일 선택", [f.name for f in log_files])
        chosen_path = core.LOG_DIR / chosen
        st.code(chosen_path.read_text(encoding="utf-8")[-8000:])
