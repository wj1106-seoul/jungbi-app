# -*- coding: utf-8 -*-
"""
app.py - 정비사업 입찰공고 수집기 사내 웹앱 (Streamlit)

실행 방법:
    streamlit run app.py
그 다음 같은 사내망에 있는 다른 PC에서 http://<이 PC의 IP>:8501 로 접속하면 됩니다.
"""
import os
import json
import shutil
import base64
import subprocess
import zipfile
import tempfile
import re
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import altair as alt

import collector_core as core
import realestate_core as re_core
import report_realestate as rr_report

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

    /* 탭("오늘의 공고 수집" / "조합별 이력 조회") 글씨 크게 */
    button[data-baseweb="tab"] {{
        height: auto;
        padding: 12px 20px;
    }}
    button[data-baseweb="tab"] p {{
        font-size: 18px !important;
        font-weight: 700 !important;
    }}
    button[data-baseweb="tab"][aria-selected="true"] p {{
        color: {BRAND_ORANGE} !important;
    }}

    /* 탭(오늘의 공고 수집 / 조합별 이력 조회) 글씨 크게 */
    div[data-baseweb="tab-list"] button {{
        padding: 10px 22px !important;
    }}
    div[data-baseweb="tab-list"] button * {{
        font-size: 20px !important;
        font-weight: 800 !important;
    }}
    div[data-baseweb="tab-list"] button[aria-selected="true"] * {{
        color: {BRAND_ORANGE} !important;
    }}
    div[data-baseweb="tab-highlight"] {{
        background-color: {BRAND_ORANGE} !important;
        height: 3px !important;
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
# 로그인 (관리자 / 팀원 2단계 비밀번호)
# ------------------------------------------------------------------
def get_secret(name: str, default: str = "") -> str:
    """로컬(.env)과 Streamlit Cloud(secrets) 양쪽 모두 지원."""
    val = os.environ.get(name, "")
    if val:
        return val
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


ADMIN_PASSWORD = get_secret("ADMIN_PASSWORD")
TEAM_PASSWORD = get_secret("TEAM_PASSWORD")

if "auth_role" not in st.session_state:
    st.session_state.auth_role = None

if st.session_state.auth_role is None:
    st.markdown(
        f"""
        <style>
        .stApp {{ background-color: #D9D9D9; }}

        /* 로그인 화면일 때만: 위쪽 여백 확보 */
        div[data-testid="stAppViewContainer"] .main .block-container {{
            padding-top: 6vh;
        }}

        /* 로그인 카드 전체(로고+입력창+버튼)를 하나의 흰 박스로 묶음 */
        .st-key-login_box {{
            background: #FFFFFF !important;
            border-top: 6px solid {BRAND_ORANGE};
            box-shadow: 0 2px 12px rgba(0,0,0,0.12);
            padding: 45px 40px 35px 40px;
        }}

        .login-logo {{
            text-align: center;
            font-size: 46px;
            font-weight: 800;
            color: {BRAND_ORANGE};
            letter-spacing: -1px;
            line-height: 1;
            margin-bottom: 6px;
        }}
        .login-subtitle {{
            text-align: center;
            font-size: 14px;
            color: #444444;
            margin-bottom: 24px;
        }}

        /* 비밀번호 입력창 박스 스타일 (여러 Streamlit 버전 대응) */
        .st-key-login_box div[data-testid="stTextInput"] > div,
        .st-key-login_box div[data-baseweb="input"],
        .st-key-login_box div[data-baseweb="base-input"] {{
            background-color: #FFFFFF !important;
            border: 1px solid #CCCCCC !important;
            border-radius: 0 !important;
            box-shadow: none !important;
        }}
        .st-key-login_box input {{
            padding: 12px 14px !important;
            font-size: 15px !important;
            color: #333333 !important;
            -webkit-text-fill-color: #333333 !important;
            background-color: transparent !important;
        }}

        /* Login 버튼 */
        .st-key-login_box div[data-testid="stFormSubmitButton"] button {{
            background-color: {BRAND_ORANGE} !important;
            color: #FFFFFF !important;
            border: none !important;
            border-radius: 0 !important;
            font-weight: 700 !important;
            font-size: 16px !important;
            padding: 12px 0 !important;
            margin-top: 6px;
        }}
        .st-key-login_box div[data-testid="stFormSubmitButton"] button:hover {{
            background-color: {BRAND_ORANGE_DARK} !important;
        }}
        /* st.form 자체의 기본 테두리/배경 제거 (구버전 Streamlit 대비 안전장치) */
        .st-key-login_box div[data-testid="stForm"] {{
            border: none !important;
            background: transparent !important;
            padding: 0 !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    _, login_col, _ = st.columns([1, 1.2, 1])
    with login_col:
        with st.container(key="login_box"):
            st.markdown(
                """
                <div class="login-logo">EGA</div>
                <div class="login-subtitle">정비사업 입찰공고 수집기</div>
                """,
                unsafe_allow_html=True,
            )
            with st.form("login_form"):
                pw_input = st.text_input(
                    "비밀번호", type="password", placeholder="PASSWORD", label_visibility="collapsed"
                )
                submitted = st.form_submit_button("Login", use_container_width=True)


    if submitted:
        if ADMIN_PASSWORD and pw_input == ADMIN_PASSWORD:
            st.session_state.auth_role = "admin"
            st.rerun()
        elif TEAM_PASSWORD and pw_input == TEAM_PASSWORD:
            st.session_state.auth_role = "team"
            st.rerun()
        elif not ADMIN_PASSWORD and not TEAM_PASSWORD:
            st.error("비밀번호가 설정되지 않았습니다. 관리자에게 문의해주세요.")
        else:
            st.error("비밀번호가 올바르지 않습니다.")
    st.stop()

is_admin = st.session_state.auth_role == "admin"

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
# 첨부파일 미리보기 (다운로드 없이 내용 확인)
# ------------------------------------------------------------------
def _local_tag(tag: str) -> str:
    """XML 태그의 네임스페이스를 떼고 태그 이름만 반환. 예: '{ns}t' -> 't'"""
    return tag.rsplit("}", 1)[-1]


def _convert_hwp_to_html(path: Path) -> str:
    """구버전 HWP(바이너리, v5) 파일을 표/이미지가 살아있는 HTML로 변환 (hwp5html 사용)."""
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            result = subprocess.run(
                ["hwp5html", "--output", tmpdir, str(path)],
                capture_output=True,
                timeout=40,
            )
        except FileNotFoundError:
            raise RuntimeError("hwp5html 실행 파일을 찾을 수 없습니다.")
        except subprocess.TimeoutExpired:
            raise RuntimeError("파일 변환이 너무 오래 걸려 중단했습니다.")
        if result.returncode != 0:
            err = result.stderr.decode("utf-8", errors="ignore").strip()
            raise RuntimeError(err or "hwp5html 변환에 실패했습니다.")

        index_path = Path(tmpdir) / "index.xhtml"
        css_path = Path(tmpdir) / "styles.css"
        bindata_dir = Path(tmpdir) / "bindata"
        if not index_path.exists():
            raise RuntimeError("변환 결과(HTML)를 찾을 수 없습니다.")

        html_content = index_path.read_text(encoding="utf-8", errors="ignore")
        css_content = css_path.read_text(encoding="utf-8", errors="ignore") if css_path.exists() else ""

        # bindata/파일명 형태의 이미지 경로를 base64로 바꿔 이미지가 바로 보이게 함
        def _inline_image(match: "re.Match") -> str:
            rel_name = Path(match.group(1)).name
            img_path = bindata_dir / rel_name
            if img_path.exists():
                ext = img_path.suffix.lstrip(".").lower() or "png"
                mime = "image/jpeg" if ext in ("jpg", "jpeg") else f"image/{ext}"
                b64 = base64.b64encode(img_path.read_bytes()).decode("utf-8")
                return f'src="data:{mime};base64,{b64}"'
            return match.group(0)

        html_content = re.sub(r'src="bindata/([^"]+)"', _inline_image, html_content)

        extra_style = (
            "body{font-family:'Malgun Gothic','맑은 고딕',sans-serif; padding:16px; "
            "background:#FFFFFF; color:#222222;} "
            "table{border-collapse:collapse;} td,th{border:1px solid #999999; padding:4px 6px;}"
        )
        style_block = f"<style>{css_content}\n{extra_style}</style>"
        if "<head>" in html_content:
            html_content = html_content.replace("<head>", f"<head>{style_block}", 1)
        else:
            html_content = style_block + html_content
        return html_content


def _parse_hwpx_blocks(path: Path):
    """HWPX(zip+xml) 파일을 문단/표 블록 리스트로 파싱. [('text', str) | ('table', rows), ...]"""
    blocks = []

    def _collect_text_in(elem, out: list):
        for child in elem:
            tag = _local_tag(child.tag)
            if tag == "t":
                if child.text:
                    out.append(child.text)
            elif tag == "tbl":
                continue  # 표는 별도 블록으로 처리하므로 텍스트 수집에서는 건너뜀
            else:
                _collect_text_in(child, out)

    def _extract_table(tbl_elem):
        rows = []
        for tr in tbl_elem:
            if _local_tag(tr.tag) != "tr":
                continue
            row_cells = []
            for tc in tr:
                if _local_tag(tc.tag) != "tc":
                    continue
                cell_texts: list = []
                _collect_text_in(tc, cell_texts)
                row_cells.append(" ".join(cell_texts).strip())
            if row_cells:
                rows.append(row_cells)
        return rows

    def _walk(elem):
        for child in elem:
            tag = _local_tag(child.tag)
            if tag == "p":
                para_texts: list = []
                for run in child:
                    if _local_tag(run.tag) != "run":
                        continue
                    for rc in run:
                        rtag = _local_tag(rc.tag)
                        if rtag == "t":
                            if rc.text:
                                para_texts.append(rc.text)
                        elif rtag == "tbl":
                            rows = _extract_table(rc)
                            if rows:
                                blocks.append(("table", rows))
                        else:
                            nested: list = []
                            _collect_text_in(rc, nested)
                            para_texts.extend(nested)
                text = "".join(para_texts).strip()
                if text:
                    blocks.append(("text", text))
            else:
                _walk(child)

    try:
        with zipfile.ZipFile(path) as zf:
            section_names = sorted(
                n for n in zf.namelist()
                if "section" in n.lower() and n.lower().endswith(".xml")
            )
            for name in section_names:
                root = ET.fromstring(zf.read(name))
                _walk(root)
    except Exception as e:
        raise RuntimeError(f"hwpx 파일을 읽는 중 문제가 발생했습니다: {e}")

    return blocks


def _render_document_blocks(blocks):
    """문단/표 블록 리스트를 원본 문서에 가깝게, 꾸밈없이 렌더링."""
    if not blocks:
        st.info("추출된 내용이 없습니다. 다운로드하여 확인해주세요.")
        return

    def _esc(s):
        return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    parts = [
        '<div style="background:#FFFFFF; border:1px solid #E5E5E5; border-radius:8px; '
        'padding:28px 32px; max-height:650px; overflow-y:auto; '
        'font-family:\'Malgun Gothic\',\'맑은 고딕\',sans-serif; font-size:14.5px; '
        'line-height:1.8; color:#222222;">'
    ]
    for kind, content in blocks:
        if kind == "text":
            safe = _esc(content)
            parts.append(f'<p style="margin:0 0 14px 0;">{safe}</p>')
        elif kind == "table":
            rows_html = []
            for row in content:
                cells_html = "".join(
                    f'<td style="border:1px solid #999999; padding:6px 10px; '
                    f'font-size:13.5px; text-align:left;">{_esc(c).strip()}</td>'
                    for c in row
                )
                rows_html.append(f"<tr>{cells_html}</tr>")
            table_html = (
                '<table style="border-collapse:collapse; margin:6px 0 20px 0; width:100%;">'
                + "".join(rows_html)
                + "</table>"
            )
            parts.append(table_html)
    parts.append("</div>")
    st.markdown("".join(parts), unsafe_allow_html=True)


def render_file_preview(path: Path):
    """지원하는 형식이면 화면에 미리보기를 렌더링, 아니면 안내 문구만 표시."""
    suffix = path.suffix.lower()
    try:
        if suffix == ".pdf":
            data = path.read_bytes()
            b64 = base64.b64encode(data).decode("utf-8")
            pdf_html = f"""
            <div id="pdf-container" style="border:1px solid #E5E5E5; border-radius:8px;
                box-shadow:0 1px 4px rgba(0,0,0,0.06); height:700px; overflow-y:auto;
                background:#F5F5F5; padding:8px; text-align:center;">
              <div id="pdf-status" style="color:#888; padding:20px;">PDF를 불러오는 중...</div>
            </div>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
            <script>
              pdfjsLib.GlobalWorkerOptions.workerSrc =
                'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

              const base64Data = "{b64}";
              const raw = atob(base64Data);
              const bytes = new Uint8Array(raw.length);
              for (let i = 0; i < raw.length; i++) {{ bytes[i] = raw.charCodeAt(i); }}

              const container = document.getElementById('pdf-container');
              const statusEl = document.getElementById('pdf-status');

              pdfjsLib.getDocument({{ data: bytes }}).promise.then(function(pdf) {{
                statusEl.remove();
                let renderNext = function(pageNum) {{
                  if (pageNum > pdf.numPages) return;
                  pdf.getPage(pageNum).then(function(page) {{
                    const scale = 1.3;
                    const viewport = page.getViewport({{ scale: scale }});
                    const canvas = document.createElement('canvas');
                    canvas.width = viewport.width;
                    canvas.height = viewport.height;
                    canvas.style.display = 'block';
                    canvas.style.margin = '0 auto 10px auto';
                    canvas.style.boxShadow = '0 1px 3px rgba(0,0,0,0.15)';
                    container.appendChild(canvas);
                    const ctx = canvas.getContext('2d');
                    page.render({{ canvasContext: ctx, viewport: viewport }}).promise.then(function() {{
                      renderNext(pageNum + 1);
                    }});
                  }});
                }};
                renderNext(1);
              }}).catch(function(err) {{
                statusEl.innerText = 'PDF 미리보기를 표시할 수 없습니다: ' + err.message;
              }});
            </script>
            """
            components.html(pdf_html, height=710, scrolling=True)
        elif suffix in (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"):
            st.image(str(path), use_container_width=True)
        elif suffix == ".hwp":
            with st.spinner("HWP 파일을 문서 형태로 변환하는 중... (표/이미지 포함)"):
                html_content = _convert_hwp_to_html(path)
            components.html(html_content, height=650, scrolling=True)
        elif suffix == ".hwpx":
            with st.spinner("HWPX 파일을 문서 형태로 변환하는 중... (표 포함)"):
                blocks = _parse_hwpx_blocks(path)
            _render_document_blocks(blocks)
        else:
            st.info("이 파일 형식은 미리보기를 지원하지 않습니다. 다운로드하여 확인해주세요.")
    except Exception as e:
        st.warning(f"미리보기를 만들 수 없습니다: {e}")


def render_attachment_browser(attachment_dir_path: Path, key_prefix: str, zip_file_name: str):
    """첨부파일 zip 전체 다운로드 + 파일별 미리보기/개별 다운로드 UI를 렌더링."""
    zip_base = str(BASE_DIR / "output" / f"{key_prefix}_첨부파일_zip")
    zip_path = shutil.make_archive(zip_base, "zip", root_dir=str(attachment_dir_path))
    with open(zip_path, "rb") as f:
        st.download_button(
            "⬇️ 전체 한 번에 다운로드 (zip)",
            data=f.read(),
            file_name=zip_file_name,
            mime="application/zip",
            use_container_width=True,
            key=f"{key_prefix}_zip_dl",
        )
    st.divider()
    st.caption("파일별로 👁 미리보기(다운로드 없이 내용 확인) 또는 ⬇️ 개별 다운로드를 받으실 수 있습니다.")

    if "preview_target" not in st.session_state:
        st.session_state.preview_target = None

    folders = sorted([f for f in attachment_dir_path.iterdir() if f.is_dir()])
    if not folders:
        st.caption("다운로드된 첨부파일이 없습니다.")

    for folder in folders:
        files = sorted([f for f in folder.glob("*") if f.is_file()])
        if not files:
            continue
        st.markdown(f"**{folder.name}**")
        for f in files:
            row_name, row_preview, row_download = st.columns([6, 1.3, 1.3])
            row_name.write(f"📄 {f.name}")
            if row_preview.button("👁 미리보기", key=f"prevbtn_{key_prefix}_{folder.name}_{f.name}"):
                st.session_state.preview_target = str(f)
            with open(f, "rb") as fh:
                row_download.download_button(
                    "⬇️ 다운로드",
                    data=fh.read(),
                    file_name=f.name,
                    key=f"dl_{key_prefix}_{folder.name}_{f.name}",
                )
            if st.session_state.preview_target == str(f):
                with st.container(border=True):
                    pc_title, pc_close = st.columns([8, 1])
                    pc_title.markdown(f"**🔎 미리보기 — {f.name}**")
                    if pc_close.button("닫기", key=f"closeprev_{key_prefix}_{folder.name}_{f.name}"):
                        st.session_state.preview_target = None
                        st.rerun()
                    else:
                        render_file_preview(f)
        st.write("")


# ------------------------------------------------------------------
# 사이드바 - 실행 옵션 + 필터 편집
# ------------------------------------------------------------------
st.sidebar.header("⚙️ 실행 옵션")

role_label = "🔑 관리자" if is_admin else "👤 팀원"
st.sidebar.caption(f"현재 로그인: {role_label}")
if st.sidebar.button("로그아웃", use_container_width=True):
    st.session_state.auth_role = None
    st.rerun()
st.sidebar.divider()

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

period_mode = st.sidebar.radio("조회 기간", ["직전 영업일만", "최근 N일", "특정 기간 검색"], index=0)
yesterday_only = period_mode == "직전 영업일만"
days_back = 1
search_start_date = None
search_end_date = None
if period_mode == "최근 N일":
    days_back = st.sidebar.number_input("최근 며칠", min_value=1, max_value=30, value=3)
elif period_mode == "특정 기간 검색":
    default_start = datetime(datetime.now().year, 1, 1).date()
    default_end = datetime.now().date()
    date_range = st.sidebar.date_input(
        "조회할 기간",
        value=(default_start, default_end),
        max_value=datetime.now().date(),
    )
    if isinstance(date_range, (tuple, list)) and len(date_range) == 2:
        search_start_date, search_end_date = date_range
    else:
        st.sidebar.warning("시작일과 종료일을 모두 선택해주세요.")

download_attachments = st.sidebar.checkbox("첨부파일(공고문/지침서) 다운로드", value=True)

if is_admin:
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

tab1, tab2, tab3, tab4 = st.tabs([
    "📅 오늘의 공고 수집", "🏢 조합별 이력 조회", "🏘️ 실거래가 조사", "📊 월간 분석 보고서",
])

with tab1:
    col_run, col_status = st.columns([1, 4])
    with col_run:
        run_clicked = st.button("🚀 지금 수집하기", type="primary", use_container_width=True)

    status_placeholder = st.empty()
    log_box = st.expander("실행 로그", expanded=run_clicked)

    if run_clicked:
        if period_mode == "특정 기간 검색" and (search_start_date is None or search_end_date is None):
            st.error("먼저 사이드바에서 조회할 시작일과 종료일을 모두 선택해주세요.")
            st.stop()

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
                start_date=search_start_date,
                end_date=search_end_date,
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

        attachment_dir_path = core.DEFAULT_ATTACHMENT_DIR
        if attachment_dir_path.exists() and any(attachment_dir_path.iterdir()):
            with st.expander("📎 공고문/지침서 다운로드 · 미리보기", expanded=False):
                render_attachment_browser(
                    attachment_dir_path,
                    key_prefix="today",
                    zip_file_name="공고문_첨부파일.zip",
                )

with tab2:
    st.markdown(
        """
        <p style="color:#555555; font-size:14px; margin-top:-8px;">
        특정 발주기관(조합/신탁사 등) 이름으로, 지정한 기간 전체에서 기술용역(CM/PM/설계/감리 등) 공고를 모두 찾아드립니다.
        </p>
        """,
        unsafe_allow_html=True,
    )

    hist_col1, hist_col2, hist_col3 = st.columns([2, 1, 1])
    with hist_col1:
        institution_keyword = st.text_input(
            "발주기관 이름(일부만 입력해도 됩니다)",
            placeholder="예: 마포로5구역제2지구, OO신탁, OO조합 ...",
        )
    with hist_col2:
        years_back = st.number_input("조회 기간(년)", min_value=0.1, max_value=5.0, value=1.0, step=0.5)
    with hist_col3:
        hist_biz_types = st.multiselect(
            "업무구분", ["용역", "공사", "물품", "기타"], default=["용역"],
        )

    hist_download_attachments = st.checkbox("첨부파일(공고문/지침서) 다운로드", value=True, key="hist_dl_att")

    est_calls = core.estimate_institution_search_calls(years_back, hist_biz_types or ["용역"])
    st.caption(
        f"⏱️ 예상 API 호출 횟수: 약 {est_calls}회 — 기간이 길수록 오래 걸립니다 "
        f"(1년 기준 대략 5~15분, 5년이면 수십 분 이상 걸릴 수 있어요)."
    )

    hist_run = st.button("🔎 이력 조회하기", type="primary", use_container_width=True)
    hist_status = st.empty()
    hist_log_box = st.expander("조회 로그", expanded=hist_run)

    if hist_run:
        if not institution_keyword.strip():
            hist_status.warning("발주기관 이름(키워드)을 먼저 입력해주세요.")
        else:
            hist_log_lines = []
            hist_log_area = hist_log_box.empty()

            def hist_progress_cb(msg: str):
                hist_log_lines.append(msg)
                hist_log_area.code("\n".join(hist_log_lines[-300:]))

            with st.spinner(f"'{institution_keyword}' 관련 공고를 최근 {years_back}년치 검색 중입니다... 시간이 걸릴 수 있어요."):
                hist_result = core.search_by_institution(
                    institution_keyword,
                    service_key=service_key,
                    cfg=st.session_state.cfg,
                    years_back=years_back,
                    biz_types=hist_biz_types or ["용역"],
                    download_attachments=hist_download_attachments,
                    progress_cb=hist_progress_cb,
                )
            st.session_state.hist_result = hist_result

    hist_result = st.session_state.get("hist_result")

    if hist_result is None:
        hist_status.info("발주기관 이름을 입력하고 '이력 조회하기'를 눌러 시작하세요.")
    elif not hist_result.ok:
        hist_status.warning(hist_result.message)
    else:
        hc1, hc2 = st.columns(2)
        hc1.metric("API 조회 건수", f"{hist_result.raw_count:,}건")
        hc2.metric("필터 통과 건수", f"{len(hist_result.filtered_df):,}건")

        st.write("")
        st.dataframe(hist_result.filtered_df, use_container_width=True, height=480)

        if hist_result.excel_path and os.path.exists(hist_result.excel_path):
            with open(hist_result.excel_path, "rb") as f:
                st.download_button(
                    "📥 엑셀 파일 다운로드",
                    data=f.read(),
                    file_name=os.path.basename(hist_result.excel_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        if hist_result.attachment_dir:
            hist_attach_dir = Path(hist_result.attachment_dir)
            if hist_attach_dir.exists() and any(hist_attach_dir.iterdir()):
                with st.expander("📎 공고문/지침서 다운로드 · 미리보기", expanded=False):
                    render_attachment_browser(
                        hist_attach_dir,
                        key_prefix="hist",
                        zip_file_name="공고문_첨부파일.zip",
                    )

with tab3:
    st.subheader("🏘️ 실거래가 조사")
    st.caption("국토교통부 실거래가 공개 API를 이용해 부동산 매매 실거래가를 조회합니다.")

    molit_key = get_secret("MOLIT_SERVICE_KEY")
    if not molit_key:
        st.warning(
            "실거래가 조회용 인증키(MOLIT_SERVICE_KEY)가 설정되어 있지 않습니다. "
            "관리자에게 Streamlit Cloud Secrets 설정을 요청해주세요."
        )

    subtab_apt, subtab_biz = st.tabs(["🏠 아파트", "🏬 상업·업무용"])

    with subtab_apt:
        st.markdown("**🏠 전국 아파트 실거래가 조회**")
        st.caption("아파트 매매 실거래가를 조회합니다.")

        re_col1, re_col2, re_col3, re_col4 = st.columns([1.2, 1.5, 1, 0.8])
        with re_col1:
            re_sido = st.selectbox("시·도", list(re_core.REGION_CODES.keys()), index=list(re_core.REGION_CODES.keys()).index("경기도"), key="re_sido")
        with re_col2:
            sigungu_options = list(re_core.REGION_CODES.get(re_sido, {}).keys())
            default_idx = sigungu_options.index("성남시 분당구") if "성남시 분당구" in sigungu_options else 0
            re_sigungu = st.selectbox("시·군·구", sigungu_options, index=default_idx, key="re_sigungu")
        with re_col3:
            re_months = st.selectbox("조회기간(개월)", [1, 3, 6, 12, 24, 36, 60], index=3, key="re_months")
        with re_col4:
            re_dong = st.text_input("동 검색(선택)", key="re_dong", placeholder="예: 정자동")

        re_run = st.button("🔍 실거래가 조회", type="primary", key="re_run_btn")

        if "re_df" not in st.session_state:
            st.session_state.re_df = pd.DataFrame()
            st.session_state.re_meta = {}

        if re_run:
            lawd_cd = re_core.REGION_CODES.get(re_sido, {}).get(re_sigungu, "")
            if not molit_key:
                st.error("인증키가 없어 조회할 수 없습니다.")
            elif not lawd_cd:
                st.error("시·도/시·군·구를 다시 선택해주세요.")
            else:
                re_status = st.empty()

                def _re_progress(msg):
                    re_status.info(msg)

                try:
                    with st.spinner("실거래가 조회 중..."):
                        df_re = re_core.fetch_transactions(
                            molit_key, lawd_cd, re_months, re_dong, progress_cb=_re_progress
                        )
                    re_status.empty()
                    if df_re.empty:
                        st.info("조건에 맞는 실거래가가 없습니다.")
                        st.session_state.re_df = pd.DataFrame()
                    else:
                        st.session_state.re_df = df_re
                        st.session_state.re_meta = {
                            "sido": re_sido, "sigungu": re_sigungu, "dong": re_dong, "months": re_months,
                        }
                        st.success(f"{len(df_re):,}건 조회 완료")
                except Exception as e:
                    re_status.empty()
                    st.error(f"조회 중 오류가 발생했습니다: {e}")

        df_re_display = st.session_state.get("re_df", pd.DataFrame())
        if not df_re_display.empty:
            count = len(df_re_display)
            avg_price = int(df_re_display["거래금액(원)"].mean())
            avg_py_price = int(df_re_display["평당가(원)"].mean())
            max_price = int(df_re_display["거래금액(원)"].max())
            min_price = int(df_re_display["거래금액(원)"].min())

            kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
            kpi1.metric("총 거래건수", f"{count:,}건")
            kpi2.metric("평균 거래가", f"{avg_price:,}원")
            kpi3.metric("평균 평당가", f"{avg_py_price:,}원")
            kpi4.metric("최고 거래가", f"{max_price:,}원")
            kpi5.metric("최저 거래가", f"{min_price:,}원")

            st.dataframe(
                df_re_display[["법정동", "단지명", "계약일", "거래금액(원)", "전용면적(㎡)", "전용면적(평)", "평당가(원)", "층"]],
                use_container_width=True,
                hide_index=True,
            )

            meta = st.session_state.re_meta
            try:
                excel_bytes = re_core.build_excel_bytes(
                    df_re_display, meta.get("sido", ""), meta.get("sigungu", ""),
                    meta.get("dong", ""), meta.get("months", ""),
                )
                dong_part = f"_{meta.get('dong')}" if meta.get("dong") else ""
                file_name = f"{meta.get('sido')}_{meta.get('sigungu')}{dong_part}_아파트_실거래가.xlsx"
                st.download_button(
                    "⬇️ 엑셀 보고서 다운로드",
                    data=excel_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
            except Exception as e:
                st.warning(f"엑셀 생성 중 문제가 발생했습니다: {e}")
        else:
            st.caption("시·도 → 시·군·구 → 기간을 선택한 뒤 조회 버튼을 눌러주세요.")

    with subtab_biz:
        st.markdown("**🏬 상업·업무용 부동산 실거래가 조회**")
        st.caption(
            "상업·업무용 부동산(근린생활시설, 업무시설 등) 매매 실거래가를 조회합니다. "
            "아파트처럼 단지명이 없고, 건물유형/건물주용도로 구분됩니다."
        )

        cm_col1, cm_col2 = st.columns([1.2, 1.5])
        with cm_col1:
            cm_sido = st.selectbox(
                "시·도", list(re_core.REGION_CODES.keys()),
                index=list(re_core.REGION_CODES.keys()).index("서울특별시"), key="cm_sido",
            )
        with cm_col2:
            cm_sigungu_options = list(re_core.REGION_CODES.get(cm_sido, {}).keys())
            cm_default_idx = cm_sigungu_options.index("강남구") if "강남구" in cm_sigungu_options else 0
            cm_sigungu = st.selectbox("시·군·구", cm_sigungu_options, index=cm_default_idx, key="cm_sigungu")

        cm_col3, cm_col4 = st.columns([1.5, 1])
        with cm_col3:
            cm_date_range = st.date_input(
                "조회 기간",
                value=(datetime.now().date() - timedelta(days=90), datetime.now().date()),
                max_value=datetime.now().date(),
                key="cm_date_range",
            )
        with cm_col4:
            cm_dong = st.text_input("동 검색(선택, 콤마로 여러 개)", key="cm_dong", placeholder="예: 역삼동, 삼성동")

        cm_use_all = st.checkbox("용도 필터 없이 상업·업무용 전체 조회", value=False, key="cm_use_all")
        if cm_use_all:
            cm_use_filter = None
        else:
            cm_use_keywords = st.text_input(
                "건물주용도 필터(콤마로 구분, 부분일치)", value=", ".join(re_core.DEFAULT_USE_FILTER),
                key="cm_use_keywords",
            )
            cm_use_filter = [u.strip() for u in cm_use_keywords.split(",") if u.strip()] or None

        cm_run = st.button("🔍 실거래가 조회", type="primary", key="cm_run_btn")

        if "cm_df" not in st.session_state:
            st.session_state.cm_df = pd.DataFrame()
            st.session_state.cm_meta = {}

        if cm_run:
            cm_lawd_cd = re_core.REGION_CODES.get(cm_sido, {}).get(cm_sigungu, "")
            if not molit_key:
                st.error("인증키가 없어 조회할 수 없습니다.")
            elif not cm_lawd_cd:
                st.error("시·도/시·군·구를 다시 선택해주세요.")
            elif not isinstance(cm_date_range, (tuple, list)) or len(cm_date_range) != 2:
                st.error("조회 기간의 시작일과 종료일을 모두 선택해주세요.")
            else:
                cm_start_date, cm_end_date = cm_date_range
                cm_status = st.empty()

                def _cm_progress(msg):
                    cm_status.info(msg)

                try:
                    with st.spinner("상업·업무용 실거래가 조회 중..."):
                        df_cm = re_core.fetch_commercial_transactions(
                            molit_key, cm_lawd_cd, cm_start_date, cm_end_date,
                            dong_filter=cm_dong, use_filter=cm_use_filter, progress_cb=_cm_progress,
                        )
                    cm_status.empty()
                    if df_cm.empty:
                        st.info("조건에 맞는 실거래가가 없습니다.")
                        st.session_state.cm_df = pd.DataFrame()
                    else:
                        st.session_state.cm_df = df_cm
                        st.session_state.cm_meta = {
                            "sido": cm_sido, "sigungu": cm_sigungu, "dong": cm_dong,
                            "start_date": cm_start_date, "end_date": cm_end_date, "use_filter": cm_use_filter,
                        }
                        st.success(f"{len(df_cm):,}건 조회 완료")
                except re_core.ServiceKeyError as e:
                    cm_status.empty()
                    st.error(f"인증키 문제로 조회할 수 없습니다: {e}")
                except Exception as e:
                    cm_status.empty()
                    st.error(f"조회 중 오류가 발생했습니다: {e}")

        df_cm_display = st.session_state.get("cm_df", pd.DataFrame())
        if not df_cm_display.empty:
            cm_count = len(df_cm_display)
            cm_avg_price = int(df_cm_display["거래금액(원)"].mean())
            cm_py_valid = df_cm_display[df_cm_display["평당가(원)"] > 0]["평당가(원)"]
            cm_avg_py_price = int(cm_py_valid.mean()) if not cm_py_valid.empty else 0
            cm_max_price = int(df_cm_display["거래금액(원)"].max())
            cm_min_price = int(df_cm_display["거래금액(원)"].min())

            cm_kpi1, cm_kpi2, cm_kpi3, cm_kpi4, cm_kpi5 = st.columns(5)
            cm_kpi1.metric("총 거래건수", f"{cm_count:,}건")
            cm_kpi2.metric("평균 거래가", f"{cm_avg_price:,}원")
            cm_kpi3.metric("평균 평당가", f"{cm_avg_py_price:,}원" if cm_avg_py_price else "산정불가")
            cm_kpi4.metric("최고 거래가", f"{cm_max_price:,}원")
            cm_kpi5.metric("최저 거래가", f"{cm_min_price:,}원")

            st.dataframe(
                df_cm_display[[
                    "법정동", "지번", "건물유형", "건물주용도", "계약일", "거래금액(원)",
                    "건물면적(㎡)", "대지면적(㎡)", "평당가(원)", "층",
                ]],
                use_container_width=True,
                hide_index=True,
            )

            cm_meta = st.session_state.cm_meta
            try:
                cm_excel_bytes = re_core.build_commercial_excel_bytes(
                    df_cm_display, cm_meta.get("sido", ""), cm_meta.get("sigungu", ""),
                    cm_meta.get("dong", ""), cm_meta.get("start_date"), cm_meta.get("end_date"),
                    cm_meta.get("use_filter"),
                )
                cm_dong_part = f"_{cm_meta.get('dong')}" if cm_meta.get("dong") else ""
                cm_file_name = f"{cm_meta.get('sido')}_{cm_meta.get('sigungu')}{cm_dong_part}_상업업무용_실거래가.xlsx"
                st.download_button(
                    "⬇️ 엑셀 보고서 다운로드",
                    data=cm_excel_bytes,
                    file_name=cm_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    key="cm_download_btn",
                )
            except Exception as e:
                st.warning(f"엑셀 생성 중 문제가 발생했습니다: {e}")
        else:
            st.caption("시·도 → 시·군·구 → 기간을 선택한 뒤 조회 버튼을 눌러주세요.")

with tab4:
    st.subheader("📊 부동산 실거래가 분석 보고서 (관리자 전용)")

    if not is_admin:
        st.info("이 기능은 관리자 로그인 시에만 사용할 수 있습니다.")
    else:
        st.caption(
            "실거래가 원본 데이터가 담긴 '원본데이터' 시트를 포함한 엑셀을 업로드하면, "
            "동별·용도별 요약, 업무 지번별 분기추이, 근린생활 면적구간 분석, 층효과 비교 4개 "
            "집계시트를 다시 계산해드립니다. 원본데이터 시트는 그대로 보존됩니다."
        )

        uploaded_re = st.file_uploader(
            "실거래가 분석 엑셀 업로드 (.xlsx, '원본데이터' 시트 포함)", type=["xlsx"], key="re_report_upload"
        )

        area_threshold = st.number_input(
            "업무시설 '표준면적' 기준 (이 값 미만만 지번별 분기추이에 포함, ㎡)",
            min_value=10.0, max_value=200.0, value=50.0, step=5.0, key="re_report_area_threshold",
        )

        if "re_report_excel_bytes" not in st.session_state:
            st.session_state.re_report_excel_bytes = None

        gen_clicked = st.button(
            "🔄 보고서 생성", type="primary", key="re_report_gen_btn",
            disabled=uploaded_re is None,
        )

        if gen_clicked and uploaded_re is not None:
            input_bytes = uploaded_re.read()
            re_report_status = st.empty()

            def _re_report_progress(msg):
                re_report_status.info(msg)

            try:
                with st.spinner("집계시트 재계산 중..."):
                    excel_bytes = rr_report.refresh_realestate_excel_bytes(
                        input_bytes, area_threshold=area_threshold, progress_cb=_re_report_progress
                    )
                    df_preview = rr_report.build_master_df(input_bytes)
                re_report_status.empty()
                st.session_state.re_report_excel_bytes = excel_bytes
                st.session_state.re_report_file_stem = Path(uploaded_re.name).stem
                st.success(f"생성 완료! (거래 {len(df_preview):,}건)")
            except Exception as e:
                re_report_status.empty()
                st.error(f"보고서 생성 중 오류가 발생했습니다: {e}")

        if st.session_state.get("re_report_excel_bytes"):
            st.divider()
            stem = st.session_state.get("re_report_file_stem", "실거래가_분석")
            st.download_button(
                "⬇️ 갱신된 엑셀 다운로드",
                data=st.session_state.re_report_excel_bytes,
                file_name=f"{stem}_갱신.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="re_report_excel_dl",
            )

        with st.expander("ℹ️ 알려진 한계 (참고)"):
            st.markdown(
                "- **층효과_비교의 '업무' 상관계수** 1개 값만 예시 파일과 정확히 재현되지 않았습니다. "
                "표본이 적을 때 이상치 하나로도 상관계수가 크게 흔들리는 값이라, 다른 핵심 집계표(동별 요약, "
                "면적구간분석, 지번별 분기추이 등)보다 중요도가 낮다고 보고 우선 진행했습니다.\n"
                "- **업무 층 구간**은 5~10층/11~15층/16~20층만 집계합니다(1~4층·21층 이상은 표본이 매우 "
                "적어 제외). 필요하시면 구간을 조정해드릴 수 있어요.\n"
                "- 이 자동화는 실거래가 조회 탭에서 받은 데이터를 '원본데이터' 시트 형태(법정동/지번/건물유형/"
                "건물주용도/계약일/거래금액(원)/건물면적(㎡)/대지면적(㎡)/평당가(원)/층/비고 컬럼)로 정리한 "
                "파일을 전제로 합니다."
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
