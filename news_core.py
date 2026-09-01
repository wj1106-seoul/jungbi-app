# -*- coding: utf-8 -*-
"""
news_core.py - 키워드 기반 뉴스 스크랩 (구글 뉴스 RSS)

구글 뉴스는 별도 API 키 없이 무료 RSS 피드로 검색 결과를 받아올 수 있습니다.
(추후 네이버 검색 API 키를 발급받으면 fetch_naver_news()를 추가해 함께 사용할 수 있도록
 구조를 열어두었습니다.)
"""
import time
import xml.etree.ElementTree as ET
from datetime import datetime

import pandas as pd
import requests

GOOGLE_NEWS_RSS_URL = "https://news.google.com/rss/search"

DEFAULT_KEYWORDS = [
    "정비사업",
    "사옥 신축",
    "사옥 투자",
    "물류센터 투자",
    "오피스 사옥",
    "공장",
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


def collect_news(keywords: list, max_items_per_keyword: int = 15, days_back: int = 7, logger=None) -> pd.DataFrame:
    """여러 키워드에 대해 구글 뉴스를 수집하고, 중복 제거 및 최신순 정렬된 DataFrame 반환."""
    all_items = []
    for kw in keywords:
        if logger:
            logger.log(f"'{kw}' 뉴스 검색 중...")
        try:
            items = fetch_google_news(kw, max_items=max_items_per_keyword, days_back=days_back)
            all_items.extend(items)
            if logger:
                logger.log(f"  -> {len(items)}건 수집")
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
        widths = {"A": 6, "B": 14, "C": 55, "D": 16, "E": 12, "F": 60}
        for col, w in widths.items():
            ws.column_dimensions[col].width = w
    return buf.getvalue()
