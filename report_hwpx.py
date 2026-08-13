# -*- coding: utf-8 -*-
"""
report_hwpx.py - 정비사업 분석 보고서(HWPX, 한글 파일) 생성

report_docx.py와 동일한 8개 섹션 구조를 python-hwpx 라이브러리로 만듭니다.
build_report_data(dict)를 받아 .hwpx 파일을 bytes로 반환합니다.
"""
from hwpx.document import HwpxDocument

NAVY_HEX = "1F3864"


def _pct(v, digits=1):
    if v is None:
        return ""
    return f"{v * 100:.{digits}f}%"


def _num(v):
    if v is None or v == "":
        return ""
    try:
        return f"{round(float(v)):,}"
    except (TypeError, ValueError):
        return str(v)


def _num1(v, digits=1):
    if v is None or v == "":
        return ""
    try:
        return f"{float(v):.{digits}f}"
    except (TypeError, ValueError):
        return str(v)


def _add_bullet(doc, text):
    doc.add_paragraph(text)
    idx = len(doc.paragraphs) - 1
    doc.set_list_format(paragraph_index=idx, kind="bullet", bullet_char="•")


def _add_table(doc, headers, rows):
    n_rows = len(rows) + 1
    n_cols = len(headers)
    tbl = doc.add_table(rows=n_rows, cols=n_cols)

    for j, h in enumerate(headers):
        tbl.set_cell_text(0, j, str(h))
        tbl.set_cell_shading(0, j, NAVY_HEX)

    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            tbl.set_cell_text(i, j, str(val))

    return tbl


def build_report_hwpx_bytes(data: dict) -> bytes:
    """report_data(dict)를 받아 HWPX(한글) 문서를 만들고 bytes로 반환."""
    doc = HwpxDocument.new()
    ins = data["인사이트"]

    # ---------- 표지 ----------
    doc.add_heading("2026년 정비사업 입찰공고 분석 보고서", level=1)
    doc.add_paragraph("지역·유형별 동향, 유찰·투명성, 발주주체 구조, 가격, 사업 파이프라인")
    doc.add_paragraph(
        f"분석 대상: {data['period_start']} ~ {data['period_end']} "
        f"입찰공고 {data['총공고수']:,}건 · 사업장 {data['총사업장수']:,}개소"
    )
    doc.add_paragraph(f"작성일: {data['generated_at']}  (자동 생성)")

    # ---------- 1. 개요 ----------
    doc.add_heading("1. 개요", level=1)
    doc.add_paragraph(
        "이 보고서는 정비사업(재건축·재개발·가로주택정비 등) 협력업체 선정 입찰공고 데이터를 정리·분석한 "
        f"결과를 종합한 것입니다. 원본 데이터는 응찰업체 단위로 기록되어 있어(총 {data['총응찰행수']:,}행), "
        f"공고 단위({data['총공고수']:,}건)와 사업장 단위({data['총사업장수']:,}개소)로 각각 재집계하여 분석했습니다."
    )
    _add_bullet(doc, f"공고({data['총공고수']:,}건): 같은 사업장이 설계·감리·교통영향평가 등 서로 다른 용역을 각각 발주한 개별 건")
    _add_bullet(doc, f"사업장({data['총사업장수']:,}개소): 같은 아파트·조합이 여러 공고를 냈어도 1건으로 처리한 실제 사업지 수")
    doc.add_paragraph(
        f"데이터 기간은 {data['period_start']}부터 {data['period_end']}까지이며, "
        "이후 신규 공고가 누적되는 대로 자동 반영되는 구조입니다."
    )

    # ---------- 2. 지역·유형별 동향 ----------
    doc.add_heading("2. 지역·유형별 동향", level=1)
    doc.add_heading("2.1 사업유형별 월별 추이", level=2)
    rows = [
        [r["월"], _num(r["재건축"]), _num(r["재개발계"]), _num(r["가로주택정비"]),
         _num(r["소규모정비계"]), _num(r["기타"]), _num(r["사업장수"])]
        for r in data["지역유형추이"]
    ]
    _add_table(doc, ["월", "재건축", "재개발계", "가로주택정비", "소규모정비계", "기타", "사업장수"], rows)

    doc.add_heading("2.2 지역 분포", level=2)
    top3 = ", ".join(f"{r['시도']}({_num(r['사업장수'])}개소)" for r in data["발주주체_시도별"][:3])
    doc.add_paragraph(
        f"시도별로는 {top3} 순으로 사업장이 많았습니다. 전체 {data['총사업장수']}개소 중 발주주체(조합/추진위원회/신탁사) "
        "구성은 지역별로 뚜렷한 차이를 보입니다."
    )

    # ---------- 3. 유찰률과 결과공개 실태 ----------
    doc.add_heading("3. 유찰률과 결과공개 실태", level=1)
    doc.add_paragraph(
        "공고 결과는 크게 세 가지로 나뉩니다: 정상 개찰, 유찰·취소(실제 무산), 결과미공개(개찰은 했으나 결과를 "
        "공표하지 않음). 이 구분이 중요한 이유는, 실제 시장 실패율과 정보 투명성 문제가 전혀 다른 성격이기 때문입니다."
    )

    doc.add_heading("3.1 실제 유찰·취소율", level=2)
    doc.add_paragraph(
        f"전체 {data['총공고수']}건 중 유찰·취소는 {ins['전체유찰취소율']}%에 불과합니다. "
        "정비사업 입찰시장 자체는 안정적으로 작동하고 있는 것으로 보입니다."
    )

    doc.add_heading("3.2 결과미공개율", level=2)
    doc.add_paragraph(
        f"반면 결과미공개율은 {ins['전체결과미공개율']}%입니다. 사업유형별로는 "
        f"{ins['결과미공개율최고유형']}({ins['결과미공개율최고유형_비율']}%)이 특히 높게 나타났습니다."
    )
    rows = [
        [r["구분"], _num(r["전체공고수"]), _num(r["유찰·취소"]), _pct(r["유찰·취소율"]),
         _num(r["결과미공개"]), _pct(r["결과미공개율"]), _num1(r["평균참여업체수"])]
        for r in data["유찰_사업유형별"]
    ]
    _add_table(doc, ["구분", "전체공고수", "유찰·취소", "유찰·취소율", "결과미공개", "결과미공개율", "평균참여업체수"], rows)

    doc.add_heading("3.3 발주주체별 투명성", level=2)
    doc.add_paragraph(
        f"발주주체별 결과미공개율은 {ins['가장투명한발주주체']}이 가장 낮고(가장 투명), "
        f"{ins['가장불투명한발주주체']}이 가장 높게(가장 불투명) 나타났습니다."
    )
    rows = [
        [r["구분"], _num(r["전체공고수"]), _pct(r["유찰·취소율"]), _pct(r["결과미공개율"]), _num1(r["평균참여업체수"])]
        for r in data["유찰_발주주체별"]
    ]
    _add_table(doc, ["발주주체", "전체공고수", "유찰·취소율", "결과미공개율", "평균참여업체수"], rows)

    doc.add_heading("3.4 가로주택정비×조합 결과미공개 상세", level=2)
    doc.add_paragraph("2건 이상 발주한 조합 중 결과미공개율이 높은 상위 사례입니다.")
    rows = [
        [r["발주기관"], _num(r["전체공고수"]), _num(r["비공개건수"]), _pct(r["비공개율"])]
        for r in data["유찰_가로주택조합상세"] if r["발주기관"] != "합계"
    ]
    _add_table(doc, ["발주기관(조합)", "전체공고수", "비공개건수", "비공개율"], rows)

    # ---------- 4. 발주주체별 지역 분포 ----------
    doc.add_heading("4. 발주주체별 지역 분포", level=1)
    doc.add_paragraph("사업장 단위로 볼 때, 조합·추진위원회·신탁사 비중은 지역별로 편차가 큽니다.")
    rows = [
        [r["시도"], _num(r["조합"]), _num(r["추진위원회"]), _num(r["신탁사"]), _num(r["기타"]),
         _num(r["사업장수"]), _pct(r["신탁사비중"])]
        for r in data["발주주체_시도별"]
    ]
    _add_table(doc, ["시도", "조합", "추진위원회", "신탁사", "기타", "사업장수", "신탁사비중"], rows)

    doc.add_paragraph("서울 자치구 단위 상위 10구:")
    rows = [
        [r["자치구"], _num(r["조합"]), _num(r["추진위원회"]), _num(r["신탁사"]), _num(r["기타"]),
         _num(r["사업장수"]), _pct(r["신탁사비중"])]
        for r in data["발주주체_자치구별"]
    ]
    _add_table(doc, ["자치구", "조합", "추진위원회", "신탁사", "기타", "사업장수", "신탁사비중"], rows)

    # ---------- 5. 신탁사별 관리 사업장 현황 ----------
    doc.add_heading("5. 신탁사별 관리 사업장 현황", level=1)
    doc.add_paragraph(f"총 {len(data['신탁사별건수'])}개 신탁사가 사업장을 관리하고 있습니다.")
    rows = [[r["신탁사"], _num(r["사업장수"])] for r in data["신탁사별건수"]]
    _add_table(doc, ["신탁사", "사업장수"], rows)

    # ---------- 6. 가격 분석 ----------
    doc.add_heading("6. 가격 분석", level=1)
    doc.add_paragraph(
        "입찰금액에는 ㎡당 단가(설계자 선정 등, 100만원 미만)와 계약 총액(100만원 이상) 두 체계가 섞여 있어, "
        "자릿수가 비슷한 값끼리만 비교했습니다."
    )

    doc.add_heading("6.1 설계자 선정 단가 — 지역별", level=2)
    rows = [[r["시도"], _num(r["응찰수"]), _num(r["평균단가"])] for r in data["가격_단가_시도별"]]
    _add_table(doc, ["시도", "응찰수", "평균단가(원/㎡)"], rows)

    doc.add_heading("6.2 설계자 선정 단가 — 사업유형별", level=2)
    doc.add_paragraph(
        f"사업유형별로는 {ins['최고단가사업유형']}의 평균단가가 {_num(ins['최고단가사업유형_값'])}원/㎡로 "
        "가장 높게 나타났습니다."
    )
    rows = [[r["사업유형"], _num(r["응찰수"]), _num(r["평균단가"])] for r in data["가격_단가_유형별"]]
    _add_table(doc, ["사업유형", "응찰수", "평균단가(원/㎡)"], rows)

    doc.add_heading("6.3 계약 총액 — 구분별 규모", level=2)
    rows = [[r["구분"], _num(r["응찰수"]), _num(r["평균총액"])] for r in data["가격_총액_구분별"]]
    _add_table(doc, ["구분", "응찰수", "평균총액(원)"], rows)

    # ---------- 7. 사업장 파이프라인 전망 ----------
    doc.add_heading("7. 사업장 파이프라인 전망", level=1)
    doc.add_paragraph(
        f"전체 {data['총사업장수']}개 사업장 중 {ins['후속대기사업장수']}개소({ins['후속대기비율']}%)가 "
        "설계·엔지니어링 공고만 낸 상태로, 감리·토목 등 후속 단계 공고를 아직 내지 않았습니다."
    )
    rows = [[r["상태"], _num(r["사업장수"])] for r in data["파이프라인_진행단계"]]
    _add_table(doc, ["상태", "사업장수"], rows)

    doc.add_paragraph("월별 신규 진입 사업장과 기존 사업장의 후속 공고 분리:")
    rows = [
        [r["월"], _num(r["신규사업장"]), _num(r["후속공고"]), _num(r["전체공고"]), _pct(r["신규비중"])]
        for r in data["파이프라인_월별신규후속"]
    ]
    _add_table(doc, ["월", "신규 사업장", "후속 공고", "전체 공고", "신규 비중"], rows)

    # ---------- 8. 종합 시사점 ----------
    doc.add_heading("8. 종합 시사점", level=1)
    _add_bullet(
        doc,
        f"정비사업 입찰시장 자체는 안정적입니다. 순수 유찰·취소율은 {ins['전체유찰취소율']}%에 불과해 "
        "시장 실패보다는 정보 공개 관행이 더 큰 이슈입니다.",
    )
    _add_bullet(
        doc,
        f"결과미공개({ins['전체결과미공개율']}%)는 {ins['결과미공개율최고유형']}에서 특히 높게"
        f"({ins['결과미공개율최고유형_비율']}%) 나타났습니다.",
    )
    _add_bullet(
        doc,
        f"발주주체별로는 {ins['가장투명한발주주체']}이 가장 투명하고, {ins['가장불투명한발주주체']}이 "
        "상대적으로 결과 비공개가 많았습니다.",
    )
    _add_bullet(
        doc,
        f"{ins['최고단가사업유형']}의 ㎡당 단가가 재건축·재개발 평균보다 약 "
        f"{ins['최고단가_기준대비프리미엄pct']}% 높게 나타났습니다.",
    )
    if ins.get("최저단가지역_값"):
        ratio = ins["최고단가지역_값"] / ins["최저단가지역_값"]
        _add_bullet(
            doc,
            f"설계 단가는 지역별로 최대 {ratio:.1f}배 차이가 납니다"
            f"({ins['최저단가지역']} vs {ins['최고단가지역']}).",
        )
    _add_bullet(
        doc,
        f"{ins['후속대기사업장수']}개 사업장({ins['후속대기비율']}%)이 후속 공고 대기 중입니다. "
        "향후 감리·시공 관련 공고가 이어질 잠재 물량으로 해석할 수 있습니다.",
    )
    doc.add_paragraph(
        "※ 본 보고서는 정리된 입찰공고 데이터를 기반으로 자동 생성되었습니다. 해석 문장은 계산된 수치를 "
        "템플릿에 대입한 것으로, 세부 맥락(특정 조합명 등)은 원본 데이터를 함께 확인해 주세요."
    )

    return doc.to_bytes()
