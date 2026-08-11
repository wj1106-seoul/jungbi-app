# -*- coding: utf-8 -*-
"""
collector_core.py
------------------------------------------------------------------
누리장터 정비사업 입찰공고 수집기 - 핵심 엔진 (UI/스케줄러에서 공통으로 가져다 쓰는 모듈)

[v3 반영 (2026-08-07 기준)]
  - 공고 상태 판정을 API의 "공고구분" 필드 기반으로 변경 (변경/취소/철회/무효만 제외,
    재공고·정정공고는 정상 수집). 필드가 없을 때만 공고명 기반 예비 판정 사용.
  - 공고명 제외 규칙에 "공사" 발주 건 제외 규칙 추가 (감리/설계/심의 등 딸린 용역은 예외)
  - 분류 로직 전면 개선: CM > PM > 감리 > 설계 > 엔지니어링 우선순위, 설계는 등장 분야를
    "전기+통신"처럼 자동 조합
  - 협력업체 공고 중 "기타"로 분류된 건은 첨부 공고문 본문에서 "용역명"을 찾아 재분류
  - 공고번호 기준 중복 제거 외에, 공고명+발주기관 내용 기준 2차 중복 제거 추가

원본 collector_v2.py 대비 이번 웹앱 버전에서 유지되는 것:
  1) SERVICE_KEY를 .env / Streamlit secrets에서 읽음 (보안)
  2) 필터 키워드를 config.json으로 분리 -> 웹 화면에서 편집 가능
  3) API 호출 실패 시 지수 백오프 재시도
  4) 실행 로그 파일 + progress_cb로 화면에 실시간 표시
  5) run_pipeline() 하나로 Streamlit 앱/CLI/스케줄러 어디서든 재사용
"""

import os
import re
import time
import html
import json
import logging
import subprocess
import urllib.parse
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable, Optional

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ============================== 경로/로그 설정 ===============================

BASE_DIR = Path(__file__).resolve().parent
LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)
CONFIG_PATH = BASE_DIR / "config.json"
DEFAULT_EXCEL_PATH = BASE_DIR / "output" / "정비사업입찰공고_수집.xlsx"
DEFAULT_ATTACHMENT_DIR = BASE_DIR / "output" / "공고문_다운로드"


def get_logger() -> logging.Logger:
    logger = logging.getLogger("collector")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    fh = logging.FileHandler(LOG_DIR / f"collector_{datetime.now():%Y%m%d}.log", encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(fh)
    return logger


logger = get_logger()


def _emit(progress_cb: Optional[Callable[[str], None]], msg: str, level: str = "info"):
    """로그 파일 + (있으면) 화면 콜백에 동시에 메시지를 남김"""
    getattr(logger, level, logger.info)(msg)
    if progress_cb:
        try:
            progress_cb(msg)
        except Exception:
            pass


# ============================== curl 기반 네트워크 함수 (원본 동일) ===============================

CURL_PATH = os.environ.get("CURL_PATH", "curl")


def curl_get_json(url: str, params: dict, timeout: int = 30):
    params = dict(params)
    service_key = params.pop("serviceKey", "")
    query = urllib.parse.urlencode(params, safe="")
    if service_key:
        full_url = f"{url}?serviceKey={service_key}&{query}" if query else f"{url}?serviceKey={service_key}"
    else:
        full_url = f"{url}?{query}"

    try:
        result = subprocess.run(
            [CURL_PATH, "-s", "-S", "--max-time", str(timeout), full_url],
            capture_output=True,
            timeout=timeout + 10,
        )
    except FileNotFoundError:
        raise RuntimeError("curl 실행파일을 찾을 수 없습니다 (PATH 확인 필요).")
    except subprocess.TimeoutExpired:
        raise RuntimeError("curl 요청이 시간 초과되었습니다.")

    if result.returncode != 0:
        err_text = result.stderr.decode("utf-8", errors="ignore").strip()
        raise RuntimeError(f"curl 실행 오류(code={result.returncode}): {err_text[:200]}")

    raw_text = result.stdout.decode("utf-8", errors="ignore")
    try:
        data = json.loads(raw_text)
    except ValueError:
        return raw_text, None
    return raw_text, data


def curl_download(url: str, dest_path: str, timeout: int = 30) -> bool:
    try:
        result = subprocess.run(
            [CURL_PATH, "-s", "-S", "-L", "--max-time", str(timeout), "-o", dest_path, url],
            capture_output=True,
            timeout=timeout + 10,
        )
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False
    if result.returncode != 0:
        return False
    return os.path.exists(dest_path) and os.path.getsize(dest_path) > 0


def curl_get_json_with_retry(url, params, timeout=30, retries=3, backoff_base=2, progress_cb=None):
    """일시적 네트워크 오류(타임아웃 등)에 대해 지수 백오프로 재시도"""
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            return curl_get_json(url, params, timeout=timeout)
        except RuntimeError as e:
            last_err = e
            if attempt < retries:
                wait = backoff_base ** attempt
                _emit(progress_cb, f"  [재시도 {attempt}/{retries - 1}] {e} -> {wait}초 후 재시도", "warning")
                time.sleep(wait)
    raise last_err


# ============================== 필터/분류 설정 (config.json으로 분리) ===============================
# 발주기관/공고명 포함·제외 키워드는 웹 화면에서 편집 가능하도록 config.json에 둡니다.
# (공고구분 판정, 분류 우선순위 알고리즘 등 "엔진 로직"에 가까운 것들은 아래에서 모듈 상수로 둡니다.)

DEFAULT_CONFIG = {
    "biz_types": ["용역"],
    "include_institution_keywords": [
        "재개발", "재건축", "정비사업", "도심복합", "주택정비", "소규모재건축", "가로주택", "소규모주택",
        "자산신탁", "토지신탁", "부동산신탁", "코리아신탁", "대한토지신탁", "하나자산신탁",
        "케이비부동산신탁", "한국자산신탁",
        "추진위원회", "준비위원회", "추진준비위", "주민대표회의", "운영위원회",
    ],
    # [v3] 공고구분 필드가 없을 때만 쓰는 예비(공고명 기반) 제외 키워드. 재공고는 일부러 제외 목록에 없음(수집 대상).
    "fallback_exclude_status_keywords": ["변경공고", "변경 공고", "취소공고", "취소 공고"],
    "include_title_keywords": [
        "CM", "PCM", "건설사업관리",
        "PM", "P.M", "사업관리",
        "설계자", "설계업자", "건축설계", "기본설계", "실시설계",
        "감리",
        "환경영향평가", "재해영향평가", "교육영향평가", "교육환경평가",
        "교통영향평가", "지하안전", "건축물안전영향평가", "전력계통영향평가",
        "친환경",
        "토목설계", "토목",
        "흙막이", "석면해체", "해체계획", "인허가", "경관심의", "기반시설",
        "소방", "전기", "통신",
    ],
    # [v3] 인테리어/통신보안/사업비컨설팅/도장·방수 등 마감·솔루션 계열 제외 키워드 추가, "도시계획" 계열 통합
    "exclude_title_keywords": [
        "정비사업전문관리", "시공사", "시공자",
        "세무", "회계", "법", "변호사", "소송", "매도청구",
        "이주관리", "범죄예방", "토지등소유자", "전체회의", "총회", "홍보", "채용",
        "분양", "책임매입", "매각", "주택관리업자", "사업시행자", "우선협상",
        "정비계획", "CCTV", "횡단보도", "감정평가", "조합설립", "청산",
        "보류지", "상가", "보험", "폐기물", "공사비", "국공유지", "무상양도", "기부대양여",
        "공공디자인", "도시계획",
        "BF인증", "장애물 없는", "설계공모", "현상설계",
        "석면사전조사", "석면측정", "농도측정", "농도 측정", "HUG보증", "지장물",
        "내진설계", "내풍설계", "풍동실험", "지적측량", "토질",
        "관리처분", "관리계획수립", "원인자부담금", "음식물",
        "인테리어", "실내건축", "가구설계",
        "통신보안", "보안솔루션", "보안 솔루션", "정보보안", "홈네트워크 보안",
        "사업비절감", "사업비 절감", "정산", "환급", "원가절감",
        "도료", "도장", "방수", "조명", "사인물",
        "제연", "부속실",
    ],
    "exclude_unless_keyword": {"경관": "경관심의"},
    "exclude_title_phrases": ["임대주택 매각"],
    "unclear_coop_note": "수동확인",
}


def load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                loaded = json.load(f)
            # 이전 버전(config.json)에 없던 v3 신규 키가 있으면 기본값으로 채워 넣음(하위 호환)
            merged = dict(DEFAULT_CONFIG)
            merged.update(loaded)
            for key, default_val in DEFAULT_CONFIG.items():
                if key not in loaded:
                    merged[key] = default_val
            return merged
        except (json.JSONDecodeError, OSError):
            logger.warning("config.json 파싱 실패 - 기본값으로 복구합니다.")
    save_config(DEFAULT_CONFIG)
    return json.loads(json.dumps(DEFAULT_CONFIG))


def save_config(cfg: dict) -> None:
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ============================== API 호출용 내부 설정 ===============================

URL_PREFIX_CANDIDATES = [
    "https://apis.data.go.kr/1230000/ao/PrvtBidNtceService",
    "http://apis.data.go.kr/1230000/ao/PrvtBidNtceService",
]

OPERATION_BY_BIZ = {
    "용역": "getPrvtBidPblancListInfoServc",
    "물품": "getPrvtBidPblancListInfoThng",
    "공사": "getPrvtBidPblancListInfoCnstwk",
    "기타": "getPrvtBidPblancListInfoEtc",
}

CHUNK_DAYS = 3
NUM_OF_ROWS = 500

# [v3] 공고구분(공고종류) 필드 후보 - 일반/변경/취소/재공고 판정용
NOTICE_KIND_FIELDS = [
    "ntceKindNm", "bidNtceKindNm", "pblancKindNm", "ntceDivNm",
    "bidNtceDivNm", "rgstTyNm", "ntceSeCd", "ntceSeNm",
]

# 공고구분 값이 이 중 하나에 해당하면 제외 (재공고·정정공고는 여기 없음 = 수집됨)
EXCLUDE_NOTICE_KINDS = ["변경", "취소", "철회", "무효"]

FIELD_CANDIDATES = {
    "공고명": ["bidNtceNm", "pblancNm", "ntceNm"],
    "공고구분": NOTICE_KIND_FIELDS,
    "발주기관": ["ntceInsttNm", "dminsttNm", "ordersInsttNm", "orderInsttNm"],
    "지역": ["prtcptPsblRgnNm", "rgnNm", "areaNm"],
    "공고일시": ["bidNtceDt", "nticeDt", "ntceDt", "pblancDt", "rgstDt"],
    "추정가격": ["asignBdgtAmt", "presmptPrce"],
    "기초금액": ["bssamt", "baseAmt", "presmptPrce"],
    "투찰마감": ["bidClseDt", "bidClseDate", "opengDt"],
    "공고번호": ["bidNtceNo", "pblancNo"],
    "공고차수": ["bidNtceOrd", "pblancOrd"],
}

# [v3] "OO공사 용역"처럼 실제로는 공사 발주인 건을 제외. 단, 공사에 딸린 감리/설계/심의 용역은 살림.
CONSTRUCTION_HINT = "공사"
CONSTRUCTION_ALLOW = [
    "감리", "설계", "계획서", "심의", "평가", "진단", "점검",
    "용역관리", "사업관리", "건설사업관리",
]

# [v3] 분류 알고리즘용 상수 (설계 세부내역 조합, 엔지니어링 키워드)
DESIGN_FIELDS = [
    ("정보통신", "통신"), ("통신", "통신"), ("전기", "전기"), ("소방", "소방"),
    ("기계", "기계설비"), ("설비", "기계설비"),
    ("토목", "토목"), ("조경", "조경"), ("건축", "건축"),
]
ENGINEERING_KEYWORDS = [
    "교통영향평가", "환경영향평가", "재해영향평가", "교육환경평가", "교육영향평가",
    "친환경평가", "친환경인증", "경관심의", "지하안전", "전력계통영향평가", "전력계통",
    "건축물안전영향평가", "해체계획서", "해체계획",
]

# [v3] 협력업체 공고 본문에서 용역명을 찾기 위한 라벨
SERVICE_NAME_LABELS = ["용역명", "용역명칭", "과업명", "용역및공사명", "공사및용역명", "입찰건명"]
PROJECT_NAME_LABELS = ["사업명", "사업지명"]
_NEXT_LABEL_CUT = r"(?=\s*(?:[가-힣]|\d+)\s*[.)]\s*[가-힣\s]{1,8}\s*[:：])"


def pick_field(item: dict, candidates: list, default=""):
    for c in candidates:
        if c in item and item[c] not in (None, ""):
            return item[c]
    return default


def clean_text(s: str) -> str:
    if not s:
        return ""
    s = html.unescape(str(s))
    s = s.replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


# ---------------------------- API 호출 부분 ----------------------------

def get_previous_business_day(reference_dt: datetime) -> datetime:
    day = reference_dt.replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=1)
    while day.weekday() >= 5:
        day -= timedelta(days=1)
    return day


def make_date_chunks(yesterday_only: bool, days_back: int):
    if yesterday_only:
        target_day = get_previous_business_day(datetime.now())
        start_dt = target_day
        end_dt = target_day + timedelta(days=1, minutes=-1)
        return [(start_dt, end_dt)]

    end_dt = datetime.now()
    start_dt = end_dt - timedelta(days=days_back)
    chunks = []
    cur = start_dt
    while cur < end_dt:
        chunk_end = min(cur + timedelta(days=CHUNK_DAYS), end_dt)
        chunks.append((cur, chunk_end))
        cur = chunk_end
    return chunks


def fetch_one_page(operation: str, begin_dt: datetime, end_dt: datetime, page_no: int,
                    service_key: str, progress_cb=None):
    params = {
        "serviceKey": service_key,
        "type": "json",
        "inqryDiv": "1",
        "inqryBgnDt": begin_dt.strftime("%Y%m%d%H%M"),
        "inqryEndDt": end_dt.strftime("%Y%m%d%H%M"),
        "pageNo": page_no,
        "numOfRows": NUM_OF_ROWS,
    }

    last_error = None
    for prefix in URL_PREFIX_CANDIDATES:
        url = f"{prefix}/{operation}"
        try:
            raw_text, data = curl_get_json_with_retry(url, params, timeout=45, progress_cb=progress_cb)
        except RuntimeError as e:
            last_error = f"{e} (주소: {prefix})"
            continue

        if data is None:
            last_error = f"JSON 파싱 실패 (주소: {prefix}) - 응답 앞부분: {raw_text[:300]}"
            continue

        if "response" not in data:
            last_error = f"예상치 못한 응답 구조 (주소: {prefix}) - 응답 앞부분: {raw_text[:300]}"
            continue

        header = data["response"].get("header", {})
        result_code = header.get("resultCode")

        if result_code != "00":
            last_error = f"API 오류 resultCode={result_code}, msg={header.get('resultMsg')} (주소: {prefix})"
            continue

        if prefix in URL_PREFIX_CANDIDATES:
            URL_PREFIX_CANDIDATES.remove(prefix)
            URL_PREFIX_CANDIDATES.insert(0, prefix)

        body = data.get("response", {}).get("body", {})
        items = body.get("items", [])
        if isinstance(items, dict):
            items = [items]
        total_count = int(body.get("totalCount", 0) or 0)
        return items, total_count

    raise RuntimeError(last_error or "알 수 없는 오류 (모든 주소 후보 실패)")


def fetch_all(biz_type: str, service_key: str, yesterday_only: bool, days_back: int, progress_cb=None):
    operation = OPERATION_BY_BIZ[biz_type]
    all_items = []

    for begin_dt, end_dt in make_date_chunks(yesterday_only, days_back):
        page_no = 1
        while True:
            try:
                items, total_count = fetch_one_page(operation, begin_dt, end_dt, page_no,
                                                      service_key, progress_cb=progress_cb)
            except RuntimeError as e:
                _emit(progress_cb, f"  [!] {biz_type} 조회 중 오류(최종 실패): {e}", "error")
                break

            if not items:
                break
            all_items.extend(items)

            if page_no * NUM_OF_ROWS >= total_count:
                break
            page_no += 1
            time.sleep(0.2)

    for it in all_items:
        it["_업무구분"] = biz_type

    return all_items


def collect_raw(cfg: dict, service_key: str, yesterday_only: bool, days_back: int, progress_cb=None):
    all_rows = []
    for biz_type in cfg["biz_types"]:
        _emit(progress_cb, f"조회 중: [{biz_type}] 전체 (날짜 범위 내) ...")
        rows = fetch_all(biz_type, service_key, yesterday_only, days_back, progress_cb=progress_cb)
        _emit(progress_cb, f"  -> {len(rows)}건 조회됨")
        all_rows.extend(rows)
    return all_rows


# ---------------------------- 필터링 / 분류 로직 (v3) ----------------------------

def get_notice_kind(item: dict) -> str:
    """API 응답에서 공고구분(공고종류) 값을 꺼냄. 필드가 없으면 빈 문자열(=판정 불가)."""
    return clean_text(pick_field(item, NOTICE_KIND_FIELDS))


def is_registered_notice(item: dict, title: str, cfg: dict) -> bool:
    """
    1순위: API의 공고구분 필드로 판정. '변경/취소/철회/무효'만 제외하고
           '일반공고·재공고·정정공고'는 모두 수집합니다.
    2순위: 공고구분 필드가 응답에 없을 때만, 공고명 기반으로 예비 판정합니다.
    """
    kind = get_notice_kind(item)
    if kind:
        return not any(bad in kind for bad in EXCLUDE_NOTICE_KINDS)

    t = title.replace(" ", "")
    fallback_kws = cfg.get("fallback_exclude_status_keywords", [])
    return not any(kw.replace(" ", "") in t for kw in fallback_kws)


def institution_passes(institution: str, cfg: dict) -> bool:
    return any(kw in institution for kw in cfg["include_institution_keywords"])


def title_excluded(title: str, cfg: dict) -> bool:
    """공백을 지운 문자열로 비교(띄어쓰기 표기 차이 무시) + 공사 발주 건 제외 규칙 포함"""
    t = title.replace(" ", "")

    if any(kw.replace(" ", "") in t for kw in cfg["exclude_title_keywords"]):
        return True
    if any(phrase.replace(" ", "") in t for phrase in cfg["exclude_title_phrases"]):
        return True
    for bad_kw, unless_kw in cfg["exclude_unless_keyword"].items():
        if bad_kw in t and unless_kw not in t:
            return True
    if CONSTRUCTION_HINT in t and not any(a in t for a in CONSTRUCTION_ALLOW):
        return True
    return False


def title_included(title: str, cfg: dict) -> bool:
    return any(kw in title for kw in cfg["include_title_keywords"])


def classify_title(title: str):
    """
    공고명으로 (구분, 세부내역)을 결정. 매칭 안 되면 ("기타", "").
    우선순위: CM > PM > 감리 > 설계 > 엔지니어링.
    설계 세부내역은 등장 분야를 모아 "전기+통신"처럼 자동 조합.
    """
    t = re.sub(r"[\s,·]", "", title)

    if any(kw in t for kw in ["건설사업관리", "PCM", "CM"]):
        return "CM", "건설사업관리"
    if any(kw in t for kw in ["사업관리", "P.M", "PM"]):
        return "PM", "사업관리"

    if "감리" in t:
        return "감리", ("석면해체" if "석면" in t else "감리")

    if "설계" in t:
        if "설계자" in t or "설계업자" in t:
            return "설계", "설계자"
        found = {}
        for kw, name in DESIGN_FIELDS:
            pos = t.find(kw)
            if pos >= 0 and (name not in found or pos < found[name]):
                found[name] = pos
        fields = [name for name, _ in sorted(found.items(), key=lambda x: x[1])]
        return "설계", ("+".join(fields) if fields else "설계")

    for kw in ENGINEERING_KEYWORDS:
        if kw in t:
            return "엔지니어링", kw

    return "기타", ""


def extract_region(institution: str, title: str) -> str:
    text = institution + " " + title
    m = re.search(r"[가-힣]+시\s?[가-힣]+구", text)
    if m:
        return normalize_region_name(m.group(0).replace("  ", " "))
    m = re.search(r"[가-힣]+\d*동\d*가?\s*\d+(-\d+)?번지", text)
    if m:
        return m.group(0)
    m = re.search(r"[가-힣0-9]+구역", text)
    if m:
        return m.group(0)
    m = re.search(r"[가-힣]+\d*동(?=\s|$)", text)
    if m:
        return m.group(0)
    return ""


def format_notice_datetime(raw) -> str:
    if not raw:
        return ""
    digits = re.sub(r"\D", "", str(raw))
    if len(digits) >= 8:
        return f"{digits[0:4]}-{digits[4:6]}-{digits[6:8]}"
    return str(raw)


def make_content_key(title: str, institution: str) -> str:
    """
    중복 판정용 키. 공고명에서 공백·괄호·쉼표 등 표기 차이를 모두 제거하고,
    끝에 붙는 '입찰공고 / 공고 / 입찰' 같은 상투어도 떼어낸 뒤 발주기관명과 합침.
    """
    t = re.sub(r"[\s\(\)\[\]{}·,\.\-_/]", "", str(title))
    t = re.sub(r"(입찰재공고|입찰공고|재공고|입찰|공고)$", "", t)
    inst = re.sub(r"\s", "", str(institution))
    return f"{t}|{inst}"


def apply_filters_and_classify(raw_items: list, cfg: dict):
    records = []
    for it in raw_items:
        title = clean_text(pick_field(it, FIELD_CANDIDATES["공고명"]))
        institution = clean_text(pick_field(it, FIELD_CANDIDATES["발주기관"]))

        if not title:
            continue
        if not is_registered_notice(it, title, cfg):
            continue
        if not institution_passes(institution, cfg):
            continue
        if title_excluded(title, cfg):
            continue

        note = ""
        is_coop = "협력업체" in title
        has_clear_keyword = title_included(title, cfg)

        if is_coop:
            if not has_clear_keyword:
                note = cfg["unclear_coop_note"]
        else:
            if not has_clear_keyword:
                continue

        category, detail = classify_title(title)
        region = extract_region(institution, title)

        note_parts = []
        if note:
            note_parts.append(note)
        kind = get_notice_kind(it)
        for mark in ("재공고", "정정"):
            if mark in kind or mark in title.replace(" ", ""):
                note_parts.append(f"[{mark if mark != '정정' else '정정공고'}]")
                break
        if category and category != "기타":
            tag = f"[{category}/{detail}]" if detail else f"[{category}]"
            note_parts.append(tag)
        remark = " ".join(note_parts)

        notice_dt = format_notice_datetime(pick_field(it, FIELD_CANDIDATES["공고일시"]))
        folder_digits = re.sub(r"\D", "", notice_dt)[:8] if notice_dt else ""

        rec = {
            "폴더명": folder_digits,
            "구분": category,
            "세부내역": detail,
            "공고일시": notice_dt,
            "지역": region,
            "공고명": title,
            "발주기관": institution,
            "업체명": "",
            "건축연면적(㎡)": "",
            "건축연면적(평)": "",
            "대지면적(㎡)": "",
            "대지면적(평)": "",
            "구역면적(㎡)": "",
            "구역면적(평)": "",
            "입찰금액(원)": "",
            "평단가(원)": "",
            "비고": remark,
            "_공고번호": pick_field(it, FIELD_CANDIDATES["공고번호"]),
            "_공고차수": pick_field(it, FIELD_CANDIDATES["공고차수"]),
            "_업무구분": it.get("_업무구분", ""),
            "_구분": category,
            "_세부내역": detail,
            "_사업유형": extract_project_type(title + " " + institution),
            "_raw": str(it),
            "_raw_dict": it,
        }
        records.append(rec)

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame(records)

    # 1차: 공고번호+차수 기준 중복 제거
    df["_고유키"] = (
        df["_공고번호"].astype(str) + "_" + df["_공고차수"].astype(str) + "_" + df["_업무구분"]
    )
    df = df.drop_duplicates(subset="_고유키", keep="first")

    # 2차: 내용(공고명+발주기관) 기준 중복 제거 - 표기 차이로 중복 등록된 건 정리
    df["_내용키"] = df.apply(lambda r: make_content_key(r["공고명"], r["발주기관"]), axis=1)
    df = df.drop_duplicates(subset="_내용키", keep="first")

    return df


# ---------------------------- 첨부파일(공고문/지침서) 다운로드 ----------------------------

def sanitize_filename(name: str, max_len: int = 80) -> str:
    name = clean_text(name)
    name = re.sub(r'[\\/:*?"<>|]', "", name)
    name = name.strip(" .")
    if len(name) > max_len:
        name = name[:max_len]
    return name or "제목없음"


def find_urls_in_item(item: dict):
    found = []
    for key, value in item.items():
        if key.startswith("_"):
            continue
        if isinstance(value, str) and value.strip().lower().startswith("http"):
            found.append((key, value.strip()))
    return found


def guess_extension(url: str, content_type: str, content_bytes: bytes = b"") -> str:
    if content_bytes:
        head = content_bytes[:8]
        if head.startswith(b"%PDF"):
            return ".pdf"
        if head.startswith(b"PK\x03\x04"):
            return ".hwpx"
        if head.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"):
            return ".hwp"

    for ext in (".pdf", ".hwp", ".hwpx", ".zip", ".doc", ".docx", ".xls", ".xlsx"):
        if url.lower().endswith(ext):
            return ext

    if content_type:
        ct = content_type.lower()
        if "pdf" in ct:
            return ".pdf"
        if "zip" in ct:
            return ".zip"
        if "hwp" in ct:
            return ".hwp"

    return ".dat"


def extract_project_type(text: str) -> str:
    order = [
        "소규모재건축", "가로주택정비사업", "가로주택",
        "도심복합개발", "도심복합",
        "재건축", "재개발", "주택정비사업", "주택정비",
    ]
    for kw in order:
        if kw in text:
            return kw
    return ""


# ---------------------------- 첨부파일 텍스트 추출 (지역/면적 정보용) ----------------------------

def extract_text_from_pdf(file_path: str) -> str:
    try:
        from pypdf import PdfReader
    except ImportError:
        try:
            from PyPDF2 import PdfReader
        except ImportError:
            return ""
    try:
        reader = PdfReader(file_path)
        return "\n".join((page.extract_text() or "") for page in reader.pages)
    except Exception:
        return ""


def extract_text_from_hwp(file_path: str) -> str:
    try:
        import olefile
    except ImportError:
        return ""
    try:
        ole = olefile.OleFileIO(file_path)
        stream_name = None
        for name in ole.listdir():
            if "/".join(name).lower().endswith("prvtext"):
                stream_name = name
                break
        if not stream_name:
            return ""
        data = ole.openstream(stream_name).read()
        return data.decode("utf-16le", errors="ignore")
    except Exception:
        return ""


def extract_text_from_hwpx(file_path: str) -> str:
    try:
        import zipfile
        chunks = []
        with zipfile.ZipFile(file_path) as z:
            for name in z.namelist():
                if name.startswith("Contents/section") and name.endswith(".xml"):
                    with z.open(name) as f:
                        chunks.append(f.read().decode("utf-8", errors="ignore"))
        return "\n".join(chunks)
    except Exception:
        return ""


def extract_text_from_file(file_path: str) -> str:
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pdf":
        return extract_text_from_pdf(file_path)
    if ext == ".hwp":
        return extract_text_from_hwp(file_path)
    if ext == ".hwpx":
        return extract_text_from_hwpx(file_path)
    return ""


# ------------- [v3] 협력업체 공고 재분류: 공고문 본문에서 용역명 추출 -------------

def _fuzzy(label: str) -> str:
    """라벨 글자 사이 공백 허용 ('용 역 명' 표기 대응)"""
    return r"\s*".join(re.escape(ch) for ch in label)


def _grab_label_value(flat: str, labels: list, max_len: int = 120) -> str:
    for lab in labels:
        m = re.search(_fuzzy(lab) + r"\s*[:：]\s*", flat)
        if m:
            seg = flat[m.end():m.end() + max_len]
            cut = re.search(_NEXT_LABEL_CUT, seg)
            if cut:
                seg = seg[:cut.start()]
            return re.sub(r"\s+", " ", seg).strip(" .,·-")
    return ""


def _title_after_notice_no(flat: str) -> str:
    """'○○조합 공고 제2026-007호' 바로 뒤에 붙는 공고 제목을 뽑음"""
    m = re.search(r"공\s*고\s*제?\s*[0-9A-Za-z\s\-–—]{1,20}?호", flat)
    if not m:
        return ""
    seg = flat[m.end():m.end() + 90]
    m2 = re.search(r"입찰\s*(재)?공고", seg)
    if m2:
        seg = seg[:m2.end()]
    return re.sub(r"\s+", " ", seg).strip(" .,·-")


def extract_service_name(text: str):
    """공고문 본문에서 (용역명, 출처)를 뽑음. 못 찾으면 ("", "")."""
    if not text:
        return "", ""
    flat = re.sub(r"\s+", " ", text)

    svc = _grab_label_value(flat, SERVICE_NAME_LABELS)
    if svc:
        proj = _grab_label_value(flat, PROJECT_NAME_LABELS)
        if proj and proj in svc:
            svc = svc.replace(proj, "").strip(" .,·-")
        if 4 <= len(svc) <= 80:
            return svc, "용역명"

    title = _title_after_notice_no(flat)
    if title and 4 <= len(title) <= 80:
        return title, "공고제목"
    return "", ""


def classify_from_document(text: str, cfg: dict):
    """
    공고문 본문 기반 분류. (구분, 세부내역, 용역명) 반환. 실패하면 ("기타", "", "").
    용역명 -> 공고제목 순으로 시도하고, classify_title() 결과가 '기타'면 버림
    (title_excluded()로 오추출을 한 번 더 걸러냄).
    """
    if not text:
        return "기타", "", ""
    flat = re.sub(r"\s+", " ", text)

    for getter in (
        lambda: _grab_label_value(flat, SERVICE_NAME_LABELS),
        lambda: _title_after_notice_no(flat),
    ):
        name = getter()
        if not name or not (4 <= len(name) <= 80):
            continue
        proj = _grab_label_value(flat, PROJECT_NAME_LABELS)
        if proj and proj in name:
            name = name.replace(proj, "").strip(" .,·-")
        if title_excluded(name, cfg):
            continue
        cat, det = classify_title(name)
        if cat != "기타":
            return cat, det, name
    return "기타", "", ""


def extract_location_snippet(text: str) -> str:
    m = re.search(r"(사업지|사업)?\s*위\s*치\s*[:：]?", text)
    if m:
        start = m.end()
        return text[start:start + 80]
    return ""


REGION_NAME_NORMALIZE = {
    "서울특별시": "서울시",
    "부산광역시": "부산시",
    "대구광역시": "대구시",
    "인천광역시": "인천시",
    "광주광역시": "광주시",
    "대전광역시": "대전시",
    "울산광역시": "울산시",
    "세종특별자치시": "세종시",
}


def normalize_region_name(region: str) -> str:
    for full, short in REGION_NAME_NORMALIZE.items():
        if region.startswith(full):
            return region.replace(full, short, 1)
    return region


def find_region_in_text(text: str) -> str:
    snippet = extract_location_snippet(text)
    search_targets = [snippet, text] if snippet else [text]

    for target in search_targets:
        m = re.search(
            r"([가-힣]+도)?\s*([가-힣]+(?:특별자치시|특별시|광역시|자치시|시|군))(\s*[가-힣0-9]+(?:구|읍|면))?",
            target,
        )
        if m:
            parts = [p.strip() for p in m.groups() if p]
            if parts:
                return normalize_region_name(" ".join(parts))
    return ""


AREA_LABELS = {
    "구역면적(㎡)": ["구역면적", "정비구역면적", "사업구역면적"],
    "대지면적(㎡)": ["대지면적", "부지면적", "사업부지면적"],
    "건축연면적(㎡)": ["건축연면적", "신축연면적", "연면적", "총연면적"],
}


def _fuzzy_label_pattern(label: str) -> str:
    return r"\s*".join(re.escape(ch) for ch in label)


SQM_TO_PYEONG = 3.305785


def format_area_number(value: float) -> str:
    return f"{round(value):,}"


def extract_area_info(text: str) -> dict:
    result = {}
    for col_m2, labels in AREA_LABELS.items():
        for label in labels:
            pattern = _fuzzy_label_pattern(label) + r"\s*[:：\(]?\s*([\d,]+\.?\d*)\s*(?:제곱미터|㎡|m2|m²)"
            m = re.search(pattern, text)
            if m:
                value_m2 = float(m.group(1).replace(",", ""))
                col_py = col_m2.replace("(㎡)", "(평)")
                result[col_m2] = format_area_number(value_m2)
                result[col_py] = format_area_number(value_m2 / SQM_TO_PYEONG)
                break
    return result


def download_attachments_for_row(row, attachment_dir: str):
    """
    반환값: (처리 결과 메모, 문서에서 찾은 지역 또는 빈 문자열, 면적 정보 dict, 합쳐진 문서 본문 텍스트)
    """
    item = row.get("_raw_dict", {})
    if not isinstance(item, dict):
        return "원본 데이터 없음", "", {}, ""

    urls = find_urls_in_item(item)
    if not urls:
        return "다운로드 링크 없음(URL 필드 미발견)", "", {}, ""

    institution = row.get("발주기관", "")
    project_type = row.get("_사업유형", "")
    content_label = row.get("_세부내역") or row.get("_구분") or ""

    folder_name = sanitize_filename(f"{institution}_{row['공고명']}")
    folder_path = os.path.join(attachment_dir, folder_name)
    os.makedirs(folder_path, exist_ok=True)

    doc_labels = ["공고문", "지침서"]

    results = []
    region_found = ""
    area_found = {}
    doc_texts = []
    for idx, (field_name, url) in enumerate(urls, start=1):
        download_url = url
        tmp_path = os.path.join(folder_path, f"_tmp_dl_{idx}.dat")

        ok = curl_download(download_url, tmp_path, timeout=30)
        if not ok and download_url.startswith("http://"):
            download_url = "https://" + download_url[len("http://"):]
            ok = curl_download(download_url, tmp_path, timeout=30)

        if not ok:
            results.append(f"{field_name} 실패(다운로드 오류)")
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass
            continue

        with open(tmp_path, "rb") as f:
            content_head = f.read()

        ext = guess_extension(download_url, "", content_head)
        doc_label = doc_labels[idx - 1] if idx <= len(doc_labels) else f"첨부{idx}"

        name_parts = [p for p in [institution, project_type, content_label, doc_label] if p]
        file_name = sanitize_filename(" ".join(name_parts)) + ext

        file_path = os.path.join(folder_path, file_name)
        try:
            os.replace(tmp_path, file_path)
            results.append(f"{field_name} -> {file_name}")
        except OSError as e:
            results.append(f"{field_name} 저장실패({e})")
            continue

        doc_text = extract_text_from_file(file_path)
        if doc_text:
            doc_texts.append(doc_text)
        if not region_found:
            region_found = find_region_in_text(doc_text)
        if not area_found:
            area_found = extract_area_info(doc_text)

    return "; ".join(results), region_found, area_found, "\n".join(doc_texts)


# ---------------------------- 엑셀 출력 ----------------------------

FULL_COLS = [
    "폴더명", "연번", "구분", "세부내역", "공고일시", "지역", "공고명", "발주기관", "업체명",
    "건축연면적(㎡)", "건축연면적(평)", "대지면적(㎡)", "대지면적(평)",
    "구역면적(㎡)", "구역면적(평)", "입찰금액(원)", "평단가(원)", "비고",
]

HEADER_LABELS = {
    "건축연면적(㎡)": "건축연면적\n(㎡)",
    "건축연면적(평)": "건축연면적\n(평)",
    "대지면적(㎡)": "대지면적\n(㎡)",
    "대지면적(평)": "대지면적\n(평)",
    "구역면적(㎡)": "구역면적\n(㎡)",
    "구역면적(평)": "구역면적\n(평)",
    "입찰금액(원)": "입찰금액\n(원)",
    "평단가(원)": "평단가\n(원)",
}

COL_WIDTHS = {
    "폴더명": 11, "연번": 6, "구분": 10, "세부내역": 14, "공고일시": 17, "지역": 14,
    "공고명": 42, "발주기관": 30, "업체명": 22,
    "건축연면적(㎡)": 10, "건축연면적(평)": 10, "대지면적(㎡)": 10, "대지면적(평)": 10,
    "구역면적(㎡)": 10, "구역면적(평)": 10, "입찰금액(원)": 13, "평단가(원)": 12, "비고": 20,
}

HEADER_FILL = PatternFill("solid", start_color="1F3864", end_color="1F3864")
HEADER_FONT = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(name="맑은 고딕", bold=True, size=13, color="1F3864")
BODY_FONT = Font(name="맑은 고딕", size=10)
UNCLEAR_FONT = Font(name="맑은 고딕", size=10, bold=True, color="9C5700")
UNCLEAR_FILL = PatternFill("solid", start_color="FFEB9C", end_color="FFEB9C")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_excel(ws, data_row_count: int):
    n_cols = len(FULL_COLS)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    title_cell = ws.cell(row=1, column=1)
    title_cell.font = TITLE_FONT
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 24

    ws.row_dimensions[2].height = 32
    for col_idx, col_name in enumerate(FULL_COLS, start=1):
        cell = ws.cell(row=2, column=col_idx)
        cell.value = HEADER_LABELS.get(col_name, col_name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER

    for r in range(3, 3 + data_row_count):
        ws.row_dimensions[r].height = 34
        note_value = ws.cell(row=r, column=FULL_COLS.index("비고") + 1).value
        is_unclear = note_value and "수동확인" in str(note_value)

        for col_idx, col_name in enumerate(FULL_COLS, start=1):
            cell = ws.cell(row=r, column=col_idx)
            cell.border = BORDER
            if is_unclear:
                cell.fill = UNCLEAR_FILL
                cell.font = UNCLEAR_FONT
            else:
                cell.font = BODY_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for col_idx, col_name in enumerate(FULL_COLS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = COL_WIDTHS.get(col_name, 15)

    ws.freeze_panes = "A3"


def write_excel(df_out: pd.DataFrame, excel_path: str, title_date: datetime):
    os.makedirs(os.path.dirname(excel_path) or ".", exist_ok=True)
    title_str = title_date.strftime("%Y-%m-%d")
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, header=True, startrow=1, sheet_name="Sheet1")
        ws = writer.sheets["Sheet1"]
        ws.cell(row=1, column=1, value=f"■ 정비사업 입찰공고문 ({title_str})")
        style_excel(ws, len(df_out))


# ---------------------------- 파이프라인 (엔진의 최종 진입점) ----------------------------

@dataclass
class RunResult:
    raw_count: int = 0
    filtered_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    excel_path: Optional[str] = None
    title_date: Optional[datetime] = None
    attachment_logs: list = field(default_factory=list)
    reclassified_count: int = 0
    errors: list = field(default_factory=list)
    ok: bool = True
    message: str = ""


def run_pipeline(
    *,
    service_key: str,
    cfg: Optional[dict] = None,
    yesterday_only: bool = True,
    days_back: int = 1,
    download_attachments: bool = True,
    excel_path: Optional[str] = None,
    attachment_dir: Optional[str] = None,
    progress_cb: Optional[Callable[[str], None]] = None,
) -> RunResult:
    """엔진 전체를 한 번에 실행. Streamlit 앱/CLI/스케줄러가 모두 이 함수 하나만 호출하면 됩니다."""
    cfg = cfg or load_config()
    excel_path = str(excel_path or DEFAULT_EXCEL_PATH)
    attachment_dir = str(attachment_dir or DEFAULT_ATTACHMENT_DIR)

    if not service_key:
        return RunResult(ok=False, message="SERVICE_KEY가 설정되지 않았습니다. .env 파일을 확인해주세요.")

    raw_items = collect_raw(cfg, service_key, yesterday_only, days_back, progress_cb=progress_cb)
    if not raw_items:
        return RunResult(ok=False, message="API 조회 결과 자체가 없습니다. (기간/키를 확인해보세요)")

    df = apply_filters_and_classify(raw_items, cfg)
    if df.empty:
        return RunResult(
            raw_count=len(raw_items), ok=False,
            message=f"조회된 {len(raw_items)}건 중 필터 통과한 공고가 없습니다.",
        )

    df = df.reset_index(drop=True)
    df.insert(0, "연번", range(1, len(df) + 1))

    attachment_logs = []
    reclassified_count = 0
    if download_attachments:
        _emit(progress_cb, f"--- 첨부파일(공고문/지침서) 다운로드 시도 ({len(df)}건) ---")
        os.makedirs(attachment_dir, exist_ok=True)
        for idx, row in df.iterrows():
            result, region_from_doc, area_info, doc_text = download_attachments_for_row(row, attachment_dir)
            if region_from_doc:
                df.at[idx, "지역"] = region_from_doc
            for col_name, value in area_info.items():
                df.at[idx, col_name] = value

            # [v3] 협력업체 공고 중 "기타"로 남은 건을 공고문 본문 용역명으로 재분류
            if df.at[idx, "구분"] == "기타" and "협력업체" in str(row["공고명"]) and doc_text:
                cat, det, svc_name = classify_from_document(doc_text, cfg)
                if cat != "기타":
                    df.at[idx, "구분"] = cat
                    df.at[idx, "세부내역"] = det
                    if svc_name and svc_name not in str(row["공고명"]):
                        df.at[idx, "공고명"] = f"{row['공고명']} ({svc_name})"
                    old = str(df.at[idx, "비고"] or "")
                    df.at[idx, "비고"] = (old + f" [문서추출/{cat}/{det}]").strip()
                    reclassified_count += 1
                    _emit(progress_cb, f"    · 문서 기반 재분류: {cat}/{det} <- \"{svc_name}\"")

            area_note = f", 면적정보: {area_info}" if area_info else ""
            log_line = f"  - {row['공고명']}: {result}{area_note}"
            attachment_logs.append(log_line)
            _emit(progress_cb, log_line)

        if reclassified_count:
            _emit(progress_cb, f"  [문서추출] 협력업체 공고 {reclassified_count}건을 공고문 본문으로 재분류했습니다.")

    # 재분류로 연번이 비지 않도록 마지막에 다시 매김
    df = df.reset_index(drop=True)
    df["연번"] = range(1, len(df) + 1)

    df_out = df[FULL_COLS].copy()
    title_date = get_previous_business_day(datetime.now()) if yesterday_only else datetime.now()

    write_excel(df_out, excel_path, title_date)
    _emit(progress_cb, f"=== API 조회 {len(raw_items)}건 중 필터 통과 {len(df_out)}건 -> 엑셀 갱신 완료: {excel_path} ===")

    return RunResult(
        raw_count=len(raw_items),
        filtered_df=df_out,
        excel_path=excel_path,
        title_date=title_date,
        attachment_logs=attachment_logs,
        reclassified_count=reclassified_count,
        ok=True,
        message=f"필터 통과 {len(df_out)}건",
    )


# ---------------------------- [신규] 특정 발주기관 이력 조회 ----------------------------
# "A조합 5년치 전부 정리해줘"처럼, 하루/최근N일이 아니라 특정 발주기관을 기준으로
# 훨씬 긴 기간을 뒤져서 모아보는 별도 기능. 서버가 발주기관으로 걸러주지 않기 때문에
# 그 기간 전체 데이터를 CHUNK_DAYS(3일) 단위로 잘게 쪼개 반복 조회합니다.
# (기간이 길수록 호출 횟수가 많아져 시간이 오래 걸릴 수 있습니다)

@dataclass
class InstitutionSearchResult:
    raw_count: int = 0
    filtered_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    excel_path: Optional[str] = None
    ok: bool = True
    message: str = ""


def estimate_institution_search_calls(years_back: float, biz_types: list) -> int:
    """조회 전에 대략 몇 번 API를 호출하게 될지 미리 알려주기 위한 추정치"""
    total_days = int(365 * years_back)
    chunk_count = max(1, -(-total_days // CHUNK_DAYS))  # 올림 나눗셈
    return chunk_count * len(biz_types)


def search_by_institution(
    keyword: str,
    service_key: str,
    cfg: dict,
    *,
    years_back: float = 1,
    biz_types: Optional[list] = None,
    excel_dir: Optional[str] = None,
    progress_cb: Optional[Callable[[str], None]] = None,
) -> InstitutionSearchResult:
    """
    특정 발주기관(조합/신탁사 등) 이름이 포함된 공고를 지정 기간(years_back년) 동안
    전부 찾아서, 지금 쓰는 것과 동일한 기술용역(CM/PM/설계/감리 등) 필터를 적용해 반환.
    """
    keyword = keyword.strip()
    if not keyword:
        return InstitutionSearchResult(ok=False, message="발주기관 이름(키워드)을 입력해주세요.")
    if not service_key:
        return InstitutionSearchResult(ok=False, message="SERVICE_KEY가 설정되지 않았습니다.")

    biz_types = biz_types or cfg.get("biz_types", ["용역"])
    end_dt = datetime.now()
    start_dt = end_dt - timedelta(days=int(365 * years_back))

    chunks = []
    cur = start_dt
    while cur < end_dt:
        chunk_end = min(cur + timedelta(days=CHUNK_DAYS), end_dt)
        chunks.append((cur, chunk_end))
        cur = chunk_end

    total_chunks = len(chunks) * len(biz_types)
    _emit(
        progress_cb,
        f"조회 기간: {start_dt:%Y-%m-%d} ~ {end_dt:%Y-%m-%d} "
        f"({len(chunks)}개 구간 x {len(biz_types)}개 업무구분 = 총 {total_chunks}회 조회 예정)",
    )

    all_items = []
    done = 0
    for biz_type in biz_types:
        operation = OPERATION_BY_BIZ[biz_type]
        for b, e in chunks:
            done += 1
            if done % 5 == 0 or done == total_chunks:
                _emit(progress_cb, f"  [{done}/{total_chunks}] [{biz_type}] {b:%Y-%m-%d}~{e:%Y-%m-%d} 조회 중...")
            page_no = 1
            while True:
                try:
                    items, total_count = fetch_one_page(operation, b, e, page_no, service_key, progress_cb=progress_cb)
                except RuntimeError as ex:
                    _emit(progress_cb, f"    [!] 오류(이 구간 건너뜀): {ex}", "warning")
                    break
                if not items:
                    break
                for it in items:
                    it["_업무구분"] = biz_type
                all_items.extend(items)
                if page_no * NUM_OF_ROWS >= total_count:
                    break
                page_no += 1
                time.sleep(0.2)

    _emit(progress_cb, f"총 {len(all_items)}건 조회 완료. 발주기관 '{keyword}' 포함 여부로 필터링 중...")

    records = []
    for it in all_items:
        title = clean_text(pick_field(it, FIELD_CANDIDATES["공고명"]))
        institution = clean_text(pick_field(it, FIELD_CANDIDATES["발주기관"]))
        if not title or not institution:
            continue
        if keyword not in institution:
            continue
        if not is_registered_notice(it, title, cfg):
            continue
        if title_excluded(title, cfg):
            continue

        is_coop = "협력업체" in title
        has_clear_keyword = title_included(title, cfg)
        note = ""
        if is_coop:
            if not has_clear_keyword:
                note = cfg["unclear_coop_note"]
        else:
            if not has_clear_keyword:
                continue

        category, detail = classify_title(title)
        region = extract_region(institution, title)
        notice_dt = format_notice_datetime(pick_field(it, FIELD_CANDIDATES["공고일시"]))

        note_parts = []
        if note:
            note_parts.append(note)
        kind = get_notice_kind(it)
        for mark in ("재공고", "정정"):
            if mark in kind or mark in title.replace(" ", ""):
                note_parts.append(f"[{mark if mark != '정정' else '정정공고'}]")
                break
        if category and category != "기타":
            tag = f"[{category}/{detail}]" if detail else f"[{category}]"
            note_parts.append(tag)

        records.append({
            "공고일시": notice_dt,
            "구분": category,
            "세부내역": detail,
            "지역": region,
            "공고명": title,
            "발주기관": institution,
            "업무구분": it.get("_업무구분", ""),
            "비고": " ".join(note_parts),
            "_공고번호": pick_field(it, FIELD_CANDIDATES["공고번호"]),
            "_공고차수": pick_field(it, FIELD_CANDIDATES["공고차수"]),
        })

    if not records:
        return InstitutionSearchResult(
            raw_count=len(all_items), ok=False,
            message=f"조회 기간 내 '{keyword}' 관련 기술용역 공고를 찾지 못했습니다.",
        )

    df = pd.DataFrame(records)
    df["_고유키"] = df["_공고번호"].astype(str) + "_" + df["_공고차수"].astype(str)
    df = df.drop_duplicates(subset="_고유키", keep="first")
    df = df.drop(columns=["_공고번호", "_공고차수", "_고유키"])
    df = df.sort_values("공고일시", ascending=False).reset_index(drop=True)

    excel_dir = excel_dir or str(DEFAULT_EXCEL_PATH.parent)
    os.makedirs(excel_dir, exist_ok=True)
    safe_keyword = re.sub(r'[\\/:*?"<>|]', "", keyword)[:40] or "조회결과"
    excel_path = os.path.join(excel_dir, f"{safe_keyword}_이력조회.xlsx")
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")

    _emit(progress_cb, f"=== 필터 통과 {len(df)}건 -> 엑셀 저장 완료: {excel_path} ===")

    return InstitutionSearchResult(
        raw_count=len(all_items),
        filtered_df=df,
        excel_path=excel_path,
        ok=True,
        message=f"{len(df)}건 찾음",
    )
