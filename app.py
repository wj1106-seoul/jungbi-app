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
import altair as alt

import collector_core as core

st.set_page_config(page_title="정비사업 입찰공고 수집기", page_icon="🏗️", layout="wide")

BASE_DIR = Path(__file__).resolve().parent

# ------------------------------------------------------------------
# 디자인 (커스텀 CSS) - 사내 인트라넷(EGA) 톤에 맞춘 오렌지 테마
# ------------------------------------------------------------------
BRAND_ORANGE = "#E8590C"
BRAND_ORANGE_DARK = "#C94E0B"
BRAND_ORANGE_LIGHT = "#FFF1E8"

st.markdown(
    f"""
    <style>
    .stApp {{ background-color: #F4F5F7; }}
    .main .block-container {{ padding-top: 1.5rem; max-width: 1200px; }}

    .app-header {{
        background-color: #FFFFFF;
        border-bottom: 4px solid {BRAND_ORANGE};
        padding: 20px 28px;
        border-radius: 10px;
        margin-bottom: 24px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.06);
        display: flex;
        align-items: center;
        gap: 16px;
    }}
    .app-header .logo {{
        font-size: 30px; font-weight: 800; color: {BRAND_ORANGE};
        letter-spacing: 1px; line-height: 1;
    }}
    .app-header .title-block h1 {{
        color: #222222; font-size: 22px; margin: 0 0 4px 0; font-weight: 700;
    }}
    .app-header .title-block p {{
        color: #666666; font-size: 13px; margin: 0;
    }}

    div[data-testid="stMetric"] {{
        background-color: #FFFFFF;
        border: 1px solid #EEEEEE;
        border-left: 4px solid {BRAND_ORANGE};
        border-radius: 8px;
        padding: 14px 18px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.04);
    }}
    div[data-testid="stMetricLabel"] {{ color: #888888; }}
    div[data-testid="stMetricValue"] {{ color: #222222; }}

    .stButton > button[kind="primary"] {{
        background-color: {BRAND_ORANGE};
        border: none;
        font-weight: 600;
        border-radius: 6px;
        color: #FFFFFF;
    }}
    .stButton > button[kind="primary"]:hover {{
        background-color: {BRAND_ORANGE_DARK};
    }}
    section[data-testid="stSidebar"] .stButton > button {{
        background-color: {BRAND_ORANGE_LIGHT};
        color: {BRAND_ORANGE_DARK} !important;
        border: 1px solid {BRAND_ORANGE};
        font-weight: 600;
    }}
    section[data-testid="stSidebar"] .stButton > button:hover {{
        background-color: {BRAND_ORANGE};
        color: #FFFFFF !important;
    }}

    section[data-testid="stSidebar"] {{
        background-color: #FFFFFF;
        border-right: 1px solid #EEEEEE;
    }}
    section[data-testid="stSidebar"] * {{
        color: #333333;
    }}
    section[data-testid="stSidebar"] input,
    section[data-testid="stSidebar"] textarea,
    section[data-testid="stSidebar"] div[data-baseweb="input"] input,
    section[data-testid="stSidebar"] div[data-baseweb="select"] * {{
        color: #222222 !important;
        -webkit-text-fill-color: #222222 !important;
    }}
    section[data-testid="stSidebar"] textarea,
    section[data-testid="stSidebar"] div[data-baseweb="input"],
    section[data-testid="stSidebar"] div[data-baseweb="select"] > div,
    section[data-testid="stSidebar"] div[data-baseweb="base-input"] {{
        background-color: #FAFAFA !important;
        border: 1px solid #DDDDDD !important;
    }}
    section[data-testid="stSidebar"] div[data-testid="stNumberInput"] div,
    section[data-testid="stSidebar"] div[data-testid="stNumberInput"] button {{
        background-color: #FAFAFA !important;
        color: #222222 !important;
    }}
    section[data-testid="stSidebar"] div[data-testid="stNumberInput"] input {{
        background-color: #FAFAFA !important;
        color: #222222 !important;
        -webkit-text-fill-color: #222222 !important;
    }}
    section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 {{
        color: {BRAND_ORANGE} !important;
    }}
    section[data-testid="stSidebar"] .stAlert p {{
        color: inherit;
    }}

    div[data-testid="stExpander"] {{
        background-color: #FFFFFF;
        border: 1px solid #EEEEEE !important;
        border-radius: 8px;
    }}
    div[data-testid="stExpander"] summary {{
        color: #333333 !important;
    }}

    div[data-testid="stDataFrame"] {{
        border: 1px solid #EEEEEE;
        border-radius: 8px;
    }}

    .stAlert {{ border-radius: 8px; }}

    div[data-baseweb="input"] input,
    div[data-baseweb="select"] *,
    .stNumberInput input,
    .stTextInput input {{
        color: #222222 !important;
        -webkit-text-fill-color: #222222 !important;
    }}
    div[data-baseweb="input"],
    div[data-baseweb="select"] > div,
    div[data-baseweb="base-input"] {{
        background-color: #FAFAFA !important;
        border: 1px solid #DDDDDD !important;
    }}
    div[data-testid="stCaptionContainer"] {{
        color: #666666 !important;
    }}

    /* 드롭다운(select) 목록을 열었을 때 뜨는 팝업 - 화면 최상단에 별도로 그려지는 부분 */
    ul[role="listbox"] {{
        background-color: #FFFFFF !important;
    }}
    ul[role="listbox"] li {{
        background-color: #FFFFFF !important;
        color: #222222 !important;
    }}
    ul[role="listbox"] li:hover {{
        background-color: {BRAND_ORANGE_LIGHT} !important;
    }}

    /* 모바일(좁은 화면) 최적화 */
    @media (max-width: 640px) {{
        .main .block-container {{
            padding-left: 1rem;
            padding-right: 1rem;
            padding-top: 1rem;
        }}
        .app-header {{
            flex-direction: column;
            align-items: flex-start;
            padding: 16px 18px;
            gap: 8px;
        }}
        .app-header .logo {{
            font-size: 22px;
        }}
        .app-header .title-block h1 {{
            font-size: 17px;
        }}
        .app-header .title-block p {{
            font-size: 12px;
        }}
        div[data-testid="stMetric"] {{
            padding: 10px 12px;
        }}
        div[data-testid="stMetricValue"] {{
            font-size: 20px;
        }}
    }}
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
        <div class="logo">EGA</div>
        <div class="title-block">
            <h1>정비사업 입찰공고 수집기</h1>
            <p>누리장터 민간입찰공고 중 정비사업 관련 기술용역 공고만 자동으로 걸러서 보여줍니다.</p>
        </div>
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

    search_col, cat_col = st.columns([3, 1])
    with search_col:
        search_text = st.text_input(
            "🔍 검색 (발주기관, 지역, 공고명, 비고 등에서 검색)",
            value="",
            placeholder="예: 서울, CM, 강남구, 신탁 ...",
        )
    with cat_col:
        cat_options = ["전체"] + sorted(df_display["구분"].dropna().unique().tolist())
        cat_selected = st.selectbox("구분 필터", cat_options)

    filtered_view = df_display.copy()
    if cat_selected != "전체":
        filtered_view = filtered_view[filtered_view["구분"] == cat_selected]
    if search_text.strip():
        keyword = search_text.strip()
        mask = filtered_view.astype(str).apply(
            lambda col: col.str.contains(keyword, case=False, na=False)
        ).any(axis=1)
        filtered_view = filtered_view[mask]

    st.caption(f"검색 결과: {len(filtered_view):,}건 / 전체 {len(df_display):,}건")
    st.dataframe(filtered_view, use_container_width=True, height=480)

    with st.expander("📊 통계 보기", expanded=False):
        def horizontal_bar(series: pd.Series, label_name: str):
            chart_df = series.reset_index()
            chart_df.columns = [label_name, "건수"]
            chart = (
                alt.Chart(chart_df)
                .mark_bar(color=BRAND_ORANGE)
                .encode(
                    x=alt.X("건수:Q", axis=alt.Axis(tickMinStep=1, format="d")),
                    y=alt.Y(f"{label_name}:N", sort="-x", title=None),
                    tooltip=[label_name, "건수"],
                )
                .properties(height=max(120, 32 * len(chart_df)))
            )
            st.altair_chart(chart, use_container_width=True)

        chart_col1, chart_col2 = st.columns(2)
        with chart_col1:
            st.caption("구분별 건수")
            cat_counts = df_display["구분"].replace("", "미분류").value_counts()
            horizontal_bar(cat_counts, "구분")
        with chart_col2:
            st.caption("지역별 건수 (상위 10개)")
            region_counts = (
                df_display["지역"].replace("", pd.NA).dropna().value_counts().head(10)
            )
            if region_counts.empty:
                st.caption("지역 정보가 있는 공고가 없습니다.")
            else:
                horizontal_bar(region_counts, "지역")

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
