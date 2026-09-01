# -*- coding: utf-8 -*-
"""
news_core.py - 키워드 기반 뉴스 스크랩 (구글 뉴스 RSS)

구글 뉴스는 별도 API 키 없이 무료 RSS 피드로 검색 결과를 받아올 수 있습니다.
(추후 네이버 검색 API 키를 발급받으면 fetch_naver_news()를 추가해 함께 사용할 수 있도록
 구조를 열어두었습니다.)
"""
import time
import re
import xml.etree.ElementTree as ET
from datetime import datetime
from html.parser import HTMLParser

import pandas as pd
import requests

GOOGLE_NEWS_RSS_URL = "https://news.google.com/rss/search"

DEFAULT_KEYWORDS = [
    "정비사업",
    "사옥 신축",
    "사옥 투자",
    "물류센터 투자",
    "오피스 빌딩 매입",
    "공장 부지 매입",
]

# 제목에 이 단어들이 포함되면 부동산/개발·민간투자와 무관한 기사로 보고 제외
DEFAULT_EXCLUDE_KEYWORDS = [
    # 사건사고
    "화재", "사고", "폭발", "붕괴", "사망", "부상", "파업",
    # 법적 분쟁
    "논란", "검찰", "구속", "기소", "재판", "소송", "고소", "판결",
    # 주식·실적
    "실적", "주가", "배당", "상장", "공모주", "IPO",
    # "정비"의 다른 뜻 (차량·장비 정비 등, 재건축·재개발과 무관)
    "차량정비", "정비창", "군수", "카센터",
    # 공공기관 건물 (민간분야만 원하므로 관공서 관련 제외)
    "관공서", "정부청사", "시청사", "구청사", "공공청사",
    # 인사·채용
    "채용", "인사발령", "승진",
    # 가십·연예
    "연예", "드라마", "예능",
]


def fetch_google_news(keyword: str, max_items: int = 15, days_back: int = 7) -> list:
    """구글 뉴스 RSS에서 키워드로 뉴스를 검색해서 리스트(dict)로 반환."""
    query = keyword
    if days_back:
        query = f"{keyword} when:{days_back}d"

    params = {"q": query, "hl": "ko", "gl": "KR", "ceid": "KR:ko"}
    resp = requests.get(GOOGLE_NEWS_RSS_URL, params=params, timeout=20)
    resp.raise_for_status()

    root = ET.fromstring(resp.content)
    items = []
    for item_el in root.findall(".//item")[:max_items]:
        title = (item_el.findtext("title") or "").strip()
        link = (item_el.findtext("link") or "").strip()
        pub_date_raw = (item_el.findtext("pubDate") or "").strip()

        source_el = item_el.find("source")
        source_name = source_el.text.strip() if (source_el is not None and source_el.text) else ""

        # 구글 뉴스 제목은 보통 "기사제목 - 언론사명" 형태라, 중복되는 언론사명을 제목에서 정리
        if source_name and title.endswith(f" - {source_name}"):
            title = title[: -(len(source_name) + 3)].strip()

        pub_dt = None
        for fmt in ("%a, %d %b %Y %H:%M:%S %Z", "%a, %d %b %Y %H:%M:%S %z"):
            try:
                pub_dt = datetime.strptime(pub_date_raw, fmt)
                break
            except ValueError:
                continue

        items.append(
            {
                "키워드": keyword,
                "제목": title,
                "언론사": source_name,
                "날짜": pub_dt.strftime("%Y-%m-%d") if pub_dt else pub_date_raw,
                "_날짜정렬용": pub_dt,
                "링크": link,
            }
        )
    return items


def collect_news(
    keywords: list,
    max_items_per_keyword: int = 15,
    days_back: int = 7,
    exclude_keywords: list = None,
    logger=None,
) -> pd.DataFrame:
    """여러 키워드에 대해 구글 뉴스를 수집하고, 제외 키워드 필터링 + 중복 제거 + 최신순 정렬된 DataFrame 반환."""
    exclude_keywords = exclude_keywords or []
    all_items = []
    excluded_count = 0
    for kw in keywords:
        if logger:
            logger.log(f"'{kw}' 뉴스 검색 중...")
        try:
            items = fetch_google_news(kw, max_items=max_items_per_keyword, days_back=days_back)
            kept_items = []
            for it in items:
                title_lower = it["제목"].lower()
                if any(ex.strip().lower() in title_lower for ex in exclude_keywords if ex.strip()):
                    excluded_count += 1
                    continue
                kept_items.append(it)
            all_items.extend(kept_items)
            if logger:
                skipped = len(items) - len(kept_items)
                msg = f"  -> {len(kept_items)}건 수집"
                if skipped:
                    msg += f" (제외 키워드로 {skipped}건 걸러냄)"
                logger.log(msg)
        except Exception as e:
            if logger:
                logger.log(f"  [!] '{kw}' 검색 중 오류: {e}")
        time.sleep(0.3)  # 구글 서버에 너무 빠르게 연속 요청하지 않도록 약간의 간격

    df = pd.DataFrame(all_items)
    if df.empty:
        return df

    # 같은 기사(같은 링크)가 여러 키워드에 동시에 걸리는 경우 중복 제거
    df = df.drop_duplicates(subset=["링크"]).copy()
    df = df.sort_values("_날짜정렬용", ascending=False, na_position="last")
    df = df.drop(columns=["_날짜정렬용"]).reset_index(drop=True)
    df.insert(0, "연번", range(1, len(df) + 1))
    return df


def build_news_excel_bytes(df: pd.DataFrame) -> bytes:
    """뉴스 결과를 엑셀 바이트로 변환 (다운로드 버튼용)."""
    import io

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="뉴스스크랩")
        ws = writer.sheets["뉴스스크랩"]
        widths = {"A": 6, "B": 14, "C": 45, "D": 16, "E": 12, "F": 50, "G": 12}
        for col, w in widths.items():
            ws.column_dimensions[col].width = w
    return buf.getvalue()


class _MetaDescriptionParser(HTMLParser):
    """기사 페이지의 <meta name="description"> / <meta property="og:description"> 값을 추출."""

    def __init__(self):
        super().__init__()
        self.description = ""
        self.title_fallback = ""
        self._in_title = False

    def handle_starttag(self, tag, attrs):
        attrs_dict = dict(attrs)
        if tag == "meta":
            name = (attrs_dict.get("name") or "").lower()
            prop = (attrs_dict.get("property") or "").lower()
            content = attrs_dict.get("content") or ""
            if content and (name == "description" or prop in ("og:description", "twitter:description")):
                if not self.description:
                    self.description = content.strip()
        elif tag == "title":
            self._in_title = True

    def handle_data(self, data):
        if self._in_title and not self.title_fallback:
            self.title_fallback = data.strip()

    def handle_endtag(self, tag):
        if tag == "title":
            self._in_title = False


def fetch_article_summary(url: str, timeout: float = 6.0, logger=None) -> str:
    """기사 실제 페이지에 접속해서 메타 요약(검색엔진용 소개문)을 가져옴.
    구글 뉴스 링크는 실제 언론사 사이트로 리다이렉트되는데, 사이트에 따라
    요약을 못 가져올 수 있고, 이 경우 빈 문자열을 반환함(오류로 처리하지 않음)."""
    try:
        resp = requests.get(
            url,
            timeout=timeout,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"},
            allow_redirects=True,
        )
        final_url = resp.url
        # [진단] 구글 뉴스 링크가 실제 언론사 사이트로 안 넘어가고 구글 도메인에 그대로 머물러 있는지 확인
        if logger:
            stayed_on_google = "news.google.com" in final_url or "google.com" in final_url
            logger.log(
                f"    [진단] 최종 도달 주소: {final_url[:100]} "
                f"{'(⚠️ 구글에 머물러 있음, 실제 기사로 못 넘어감)' if stayed_on_google else '(언론사 사이트 도달함)'}"
            )
        if resp.status_code != 200:
            if logger:
                logger.log(f"    [진단] 상태코드 {resp.status_code}로 실패")
            return ""
        html_head = resp.text[:20000]
        parser = _MetaDescriptionParser()
        parser.feed(html_head)
        summary = parser.description or ""
        summary = re.sub(r"\s+", " ", summary).strip()
        if logger and not summary:
            logger.log(f"    [진단] 페이지는 받았지만 meta description 태그를 못 찾음")
        return summary[:200]  # 너무 길면 200자로 자름
    except Exception as e:
        if logger:
            logger.log(f"    [진단] 접속 자체 실패: {e}")
        return ""


class _ArticleBodyParser(HTMLParser):
    """기사 페이지에서 본문으로 보이는 <p> 태그들의 텍스트를 모아 추출."""

    def __init__(self):
        super().__init__()
        self.paragraphs = []
        self._in_p = False
        self._in_skip_tag = False  # script/style 안의 텍스트는 무시
        self._buf = []

    def handle_starttag(self, tag, attrs):
        if tag == "p":
            self._in_p = True
            self._buf = []
        elif tag in ("script", "style"):
            self._in_skip_tag = True

    def handle_endtag(self, tag):
        if tag == "p" and self._in_p:
            text = "".join(self._buf).strip()
            if len(text) > 15:  # 너무 짧은(버튼 라벨 등) 조각은 제외
                self.paragraphs.append(text)
            self._in_p = False
        elif tag in ("script", "style"):
            self._in_skip_tag = False

    def handle_data(self, data):
        if self._in_p and not self._in_skip_tag:
            self._buf.append(data)


def extract_article_text(html: str, max_chars: int = 3000) -> str:
    """HTML에서 본문 <p> 태그 텍스트만 모아 하나의 문자열로 반환."""
    parser = _ArticleBodyParser()
    try:
        parser.feed(html)
    except Exception:
        pass
    full_text = "\n".join(parser.paragraphs)
    full_text = re.sub(r"[ \t]+", " ", full_text).strip()
    return full_text[:max_chars]


GEMINI_API_URL_TEMPLATE = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"


def summarize_with_gemini(article_text: str, api_key: str, model: str = "gemini-2.5-flash", timeout: float = 30.0) -> str:
    """기사 본문 텍스트를 Gemini에게 보내 2~3문장 한국어 핵심요약을 받아옴."""
    if not article_text.strip():
        return ""
    prompt = (
        "다음은 뉴스 기사 본문입니다. 이 기사의 핵심 내용을 한국어로 2~3문장으로만 "
        "간결하게 요약해 주세요. 다른 설명이나 서두 없이 요약 내용만 답하세요.\n\n"
        f"기사 본문:\n{article_text}"
    )
    url = GEMINI_API_URL_TEMPLATE.format(model=model)
    payload = {"contents": [{"role": "user", "parts": [{"text": prompt}]}]}
    try:
        resp = requests.post(
            url,
            params={"key": api_key},
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=timeout,
        )
        if resp.status_code != 200:
            return f"[요약 실패: 상태코드 {resp.status_code}]"
        data = resp.json()
        candidates = data.get("candidates", [])
        if not candidates:
            return "[요약 실패: 응답에 결과 없음]"
        parts = candidates[0].get("content", {}).get("parts", [])
        text = "".join(p.get("text", "") for p in parts).strip()
        return text
    except Exception as e:
        return f"[요약 실패: {e}]"


def add_ai_summaries(
    df: pd.DataFrame,
    gemini_api_key: str,
    max_articles: int = 20,
    model: str = "gemini-2.5-flash",
    logger=None,
) -> pd.DataFrame:
    """DataFrame의 각 기사 실제 본문을 가져와 Gemini로 진짜 핵심요약을 생성해 '요약' 컬럼에 추가."""
    if df.empty:
        return df
    df = df.copy()
    summaries = []
    n = min(len(df), max_articles)
    for i, row in df.iterrows():
        if i >= max_articles:
            summaries.append("")
            continue
        if logger:
            logger.log(f"  ({i + 1}/{n}) 기사 본문 읽는 중: {row['제목'][:30]}...")
        try:
            resp = requests.get(
                row["링크"],
                timeout=8.0,
                headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"},
                allow_redirects=True,
            )
            final_url = resp.url
            if logger:
                stayed_on_google = "google.com" in final_url
                logger.log(
                    f"    [진단] 도달 주소: {final_url[:90]} "
                    f"{'(⚠️ 구글에 머물러 있음)' if stayed_on_google else '(언론사 사이트 도달)'}"
                )
            if resp.status_code != 200:
                summaries.append("")
                continue
            article_text = extract_article_text(resp.text)
            if not article_text or len(article_text) < 50:
                if logger:
                    logger.log(f"    [진단] 본문 추출 실패 또는 너무 짧음(길이={len(article_text)})")
                summaries.append("")
                continue
            summary = summarize_with_gemini(article_text, gemini_api_key, model=model)
            summaries.append(summary)
        except Exception as e:
            if logger:
                logger.log(f"    [진단] 처리 중 오류: {e}")
            summaries.append("")
        time.sleep(0.2)
    df["요약"] = summaries
    return df
