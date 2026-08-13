# -*- coding: utf-8 -*-
"""
report_hwpx.py - 정비사업 분석 보고서(HWPX, 한글 파일) 생성

[가독성 개선 - v2]
처음 버전은 제목이 본문과 똑같은 서식으로 나오는 문제가 있었습니다. 원인은 세 가지였습니다.
  1) add_paragraph()의 기본값(inherit_style=True)이 "바로 앞 문단의 스타일을 그대로 물려받는"
     방식이라, 제목(개요 스타일) 다음에 오는 본문 문단들이 제목 스타일을 계속 물려받고 있었음.
     -> 본문/불릿은 style="바탕글", inherit_style=False로 명시적으로 리셋.
  2) 이 라이브러리가 기본 제공하는 "개요 1/2" 스타일 자체가 글자모양은 바탕글과 동일해서
     (굵기/크기/색 구분이 없음), 스타일만 바꿔서는 눈에 띄지 않았음.
     -> ensure_run_style()로 굵고 큰 남색 문자서식을 직접 만들어 제목에 적용.
  3) 표 헤더 셀의 배경색(음영)이 라이브러리 버그로 alpha=0(완전 투명)으로 저장되어 안 보였음.
     -> 문서를 만든 뒤 후처리로 alpha 값을 보정.
"""
import re

from hwpx.document import HwpxDocument

NAVY = "#1F3864"
WHITE = "#FFFFFF"


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


class ReportBuilder:
    """제목/본문/표 서식을 일관되게 적용하기 위한 헬퍼."""

    def __init__(self, doc: HwpxDocument):
        self.doc = doc
        self.h1_char = doc.ensure_run_style(bold=True, size=16, color=NAVY)
        self.h2_char = doc.ensure_run_style(bold=True, size=13, color=NAVY)
        self.title_char = doc.ensure_run_style(bold=True, size=20, color=NAVY)
        self.subtitle_char = doc.ensure_run_style(italic=True, size=12)
        self.note_char = doc.ensure_run_style(italic=True, size=9, color="#666666")
        self.table_header_char = doc.ensure_run_style(bold=True, color=WHITE)

    def _set_run_char(self, paragraph, char_pr_id):
        for run in paragraph.runs:
            run.char_pr_id_ref = char_pr_id

    def title(self, text):
        p = self.doc.add_paragraph(text, style="바탕글", inherit_style=False)
        self._set_run_char(p, self.title_char)

    def subtitle(self, text):
        p = self.doc.add_paragraph(text, style="바탕글", inherit_style=False)
        self._set_run_char(p, self.subtitle_char)

    def heading1(self, text):
        self.doc.add_heading(text, level=1, char_pr_id_ref=self.h1_char)

    def heading2(self, text):
        self.doc.add_heading(text, level=2, char_pr_id_ref=self.h2_char)

    def body(self, text):
        self.doc.add_paragraph(text, style="바탕글", inherit_style=False)

    def note(self, text):
        p = self.doc.add_paragraph(text, style="바탕글", inherit_style=False)
        self._set_run_char(p, self.note_char)

    def bullet(self, text):
        self.doc.add_paragraph(text, style="바탕글", inherit_style=False)
        idx = len(self.doc.paragraphs) - 1
        self.doc.set_list_format(paragraph_index=idx, kind="bullet", bullet_char="•")

    def table(self, headers, rows):
        n_rows = len(rows) + 1
        n_cols = len(headers)
        tbl = self.doc.add_table(rows=n_rows, cols=n_cols)

        for j, h in enumerate(headers):
            tbl.set_cell_text(0, j, str(h))
            tbl.set_cell_shading(0, j, NAVY)
            cell = tbl.cell(0, j)
            for para in cell.paragraphs:
                self._set_run_char(para, self.table_header_char)

        for i, row in enumerate(rows, start=1):
            for j, val in enumerate(row):
                tbl.set_cell_text(i, j, str(val))

        return tbl


def _fix_shading_alpha(hwpx_bytes: bytes) -> bytes:
    """라이브러리가 셀 배경색을 alpha=0(투명)으로 저장하는 버그를 보정.
    우리가 지정한 색(NAVY)의 winBrush만 alpha를 완전 불투명(255)으로 바꿔줍니다."""
    import zipfile
    import io

    color_hex = NAVY.lstrip("#").upper()
    src = io.BytesIO(hwpx_bytes)
    out = io.BytesIO()

    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "Contents/header.xml":
                text = data.decode("utf-8")
                pattern = re.compile(
                    r'(<hc:winBrush faceColor="#' + color_hex + r'"[^>]*?)alpha="0"',
                    re.IGNORECASE,
                )
                text = pattern.sub(r'\1alpha="255"', text)
                data = text.encode("utf-8")
            zout.writestr(item, data)

    return out.getvalue()


def build_report_hwpx_bytes(data: dict) -> bytes:
    """report_data(dict)를 받아 HWPX(한글) 문서를 만들고 bytes로 반환."""
    doc = HwpxDocument.new()
    b = ReportBuilder(doc)
    ins = data["인사이트"]

    b.title("2026년 정비사업 입찰공고 분석 보고서")
    b.subtitle("지역·유형별 동향, 유찰·투명성, 발주주체 구조, 가격, 사업 파이프라인")
    b.body(
        f"분석 대상: {data['period_start']} ~ {data['period_end']} "
        f"입찰공고 {data['총공고수']:,}건 · 사업장 {data['총사업장수']:,}개소"
    )
    b.body(f"작성일: {data['generated_at']}  (자동 생성)")

    b.heading1("1. 개요")
    b.body(
        "이 보고서는 정비사업(재건축·재개발·가로주택정비 등) 협력업체 선정 입찰공고 데이터를 정리·분석한 "
        f"결과를 종합한 것입니다. 원본 데이터는 응찰업체 단위로 기록되어 있어(총 {data['총응찰행수']:,}행), "
        f"공고 단위({data['총공고수']:,}건)와 사업장 단위({data['총사업장수']:,}개소)로 각각 재집계하여 분석했습니다."
    )
    b.bullet(f"공고({data['총공고수']:,}건): 같은 사업장이 설계·감리·교통영향평가 등 서로 다른 용역을 각각 발주한 개별 건")
    b.bullet(f"사업장({data['총사업장수']:,}개소): 같은 아파트·조합이 여러 공고를 냈어도 1건으로 처리한 실제 사업지 수")
    b.body(
        f"데이터 기간은 {data['period_start']}부터 {data['period_end']}까지이며, "
        "이후 신규 공고가 누적되는 대로 자동 반영되는 구조입니다."
    )

    b.heading1("2. 지역·유형별 동향")
    b.heading2("2.1 사업유형별 월별 추이")
    rows = [
        [r["월"], _num(r["재건축"]), _num(r["재개발계"]), _num(r["가로주택정비"]),
         _num(r["소규모정비계"]), _num(r["기타"]), _num(r["사업장수"])]
        for r in data["지역유형추이"]
    ]
    b.table(["월", "재건축", "재개발계", "가로주택정비", "소규모정비계", "기타", "사업장수"], rows)

    b.heading2("2.2 지역 분포")
    top3 = ", ".join(f"{r['시도']}({_num(r['사업장수'])}개소)" for r in data["발주주체_시도별"][:3])
    b.body(
        f"시도별로는 {top3} 순으로 사업장이 많았습니다. 전체 {data['총사업장수']}개소 중 발주주체(조합/추진위원회/신탁사) "
        "구성은 지역별로 뚜렷한 차이를 보입니다."
    )

    b.heading1("3. 유찰률과 결과공개 실태")
    b.body(
        "공고 결과는 크게 세 가지로 나뉩니다: 정상 개찰, 유찰·취소(실제 무산), 결과미공개(개찰은 했으나 결과를 "
        "공표하지 않음). 이 구분이 중요한 이유는, 실제 시장 실패율과 정보 투명성 문제가 전혀 다른 성격이기 때문입니다."
    )

    b.heading2("3.1 실제 유찰·취소율")
    b.body(
        f"전체 {data['총공고수']}건 중 유찰·취소는 {ins['전체유찰취소율']}%에 불과합니다. "
        "정비사업 입찰시장 자체는 안정적으로 작동하고 있는 것으로 보입니다."
    )

    b.heading2("3.2 결과미공개율")
    b.body(
        f"반면 결과미공개율은 {ins['전체결과미공개율']}%입니다. 사업유형별로는 "
        f"{ins['결과미공개율최고유형']}({ins['결과미공개율최고유형_비율']}%)이 특히 높게 나타났습니다."
    )
    rows = [
        [r["구분"], _num(r["전체공고수"]), _num(r["유찰·취소"]), _pct(r["유찰·취소율"]),
         _num(r["결과미공개"]), _pct(r["결과미공개율"]), _num1(r["평균참여업체수"])]
        for r in data["유찰_사업유형별"]
    ]
    b.table(["구분", "전체공고수", "유찰·취소", "유찰·취소율", "결과미공개", "결과미공개율", "평균참여업체수"], rows)

    b.heading2("3.3 발주주체별 투명성")
    b.body(
        f"발주주체별 결과미공개율은 {ins['가장투명한발주주체']}이 가장 낮고(가장 투명), "
        f"{ins['가장불투명한발주주체']}이 가장 높게(가장 불투명) 나타났습니다."
    )
    rows = [
        [r["구분"], _num(r["전체공고수"]), _pct(r["유찰·취소율"]), _pct(r["결과미공개율"]), _num1(r["평균참여업체수"])]
        for r in data["유찰_발주주체별"]
    ]
    b.table(["발주주체", "전체공고수", "유찰·취소율", "결과미공개율", "평균참여업체수"], rows)

    b.heading2("3.4 가로주택정비×조합 결과미공개 상세")
    b.body("2건 이상 발주한 조합 중 결과미공개율이 높은 상위 사례입니다.")
    rows = [
        [r["발주기관"], _num(r["전체공고수"]), _num(r["비공개건수"]), _pct(r["비공개율"])]
        for r in data["유찰_가로주택조합상세"] if r["발주기관"] != "합계"
    ]
    b.table(["발주기관(조합)", "전체공고수", "비공개건수", "비공개율"], rows)

    b.heading1("4. 발주주체별 지역 분포")
    b.body("사업장 단위로 볼 때, 조합·추진위원회·신탁사 비중은 지역별로 편차가 큽니다.")
    rows = [
        [r["시도"], _num(r["조합"]), _num(r["추진위원회"]), _num(r["신탁사"]), _num(r["기타"]),
         _num(r["사업장수"]), _pct(r["신탁사비중"])]
        for r in data["발주주체_시도별"]
    ]
    b.table(["시도", "조합", "추진위원회", "신탁사", "기타", "사업장수", "신탁사비중"], rows)

    b.body("서울 자치구 단위 상위 10구:")
    rows = [
        [r["자치구"], _num(r["조합"]), _num(r["추진위원회"]), _num(r["신탁사"]), _num(r["기타"]),
         _num(r["사업장수"]), _pct(r["신탁사비중"])]
        for r in data["발주주체_자치구별"]
    ]
    b.table(["자치구", "조합", "추진위원회", "신탁사", "기타", "사업장수", "신탁사비중"], rows)

    b.heading1("5. 신탁사별 관리 사업장 현황")
    b.body(f"총 {len(data['신탁사별건수'])}개 신탁사가 사업장을 관리하고 있습니다.")
    rows = [[r["신탁사"], _num(r["사업장수"])] for r in data["신탁사별건수"]]
    b.table(["신탁사", "사업장수"], rows)

    b.heading1("6. 가격 분석")
    b.body(
        "입찰금액에는 ㎡당 단가(설계자 선정 등, 100만원 미만)와 계약 총액(100만원 이상) 두 체계가 섞여 있어, "
        "자릿수가 비슷한 값끼리만 비교했습니다."
    )

    b.heading2("6.1 설계자 선정 단가 — 지역별")
    rows = [[r["시도"], _num(r["응찰수"]), _num(r["평균단가"])] for r in data["가격_단가_시도별"]]
    b.table(["시도", "응찰수", "평균단가(원/㎡)"], rows)

    b.heading2("6.2 설계자 선정 단가 — 사업유형별")
    b.body(
        f"사업유형별로는 {ins['최고단가사업유형']}의 평균단가가 {_num(ins['최고단가사업유형_값'])}원/㎡로 "
        "가장 높게 나타났습니다."
    )
    rows = [[r["사업유형"], _num(r["응찰수"]), _num(r["평균단가"])] for r in data["가격_단가_유형별"]]
    b.table(["사업유형", "응찰수", "평균단가(원/㎡)"], rows)

    b.heading2("6.3 계약 총액 — 구분별 규모")
    rows = [[r["구분"], _num(r["응찰수"]), _num(r["평균총액"])] for r in data["가격_총액_구분별"]]
    b.table(["구분", "응찰수", "평균총액(원)"], rows)

    b.heading1("7. 사업장 파이프라인 전망")
    b.body(
        f"전체 {data['총사업장수']}개 사업장 중 {ins['후속대기사업장수']}개소({ins['후속대기비율']}%)가 "
        "설계·엔지니어링 공고만 낸 상태로, 감리·토목 등 후속 단계 공고를 아직 내지 않았습니다."
    )
    rows = [[r["상태"], _num(r["사업장수"])] for r in data["파이프라인_진행단계"]]
    b.table(["상태", "사업장수"], rows)

    b.body("월별 신규 진입 사업장과 기존 사업장의 후속 공고 분리:")
    rows = [
        [r["월"], _num(r["신규사업장"]), _num(r["후속공고"]), _num(r["전체공고"]), _pct(r["신규비중"])]
        for r in data["파이프라인_월별신규후속"]
    ]
    b.table(["월", "신규 사업장", "후속 공고", "전체 공고", "신규 비중"], rows)

    b.heading1("8. 종합 시사점")
    b.bullet(
        f"정비사업 입찰시장 자체는 안정적입니다. 순수 유찰·취소율은 {ins['전체유찰취소율']}%에 불과해 "
        "시장 실패보다는 정보 공개 관행이 더 큰 이슈입니다."
    )
    b.bullet(
        f"결과미공개({ins['전체결과미공개율']}%)는 {ins['결과미공개율최고유형']}에서 특히 높게"
        f"({ins['결과미공개율최고유형_비율']}%) 나타났습니다."
    )
    b.bullet(
        f"발주주체별로는 {ins['가장투명한발주주체']}이 가장 투명하고, {ins['가장불투명한발주주체']}이 "
        "상대적으로 결과 비공개가 많았습니다."
    )
    b.bullet(
        f"{ins['최고단가사업유형']}의 ㎡당 단가가 재건축·재개발 평균보다 약 "
        f"{ins['최고단가_기준대비프리미엄pct']}% 높게 나타났습니다."
    )
    if ins.get("최저단가지역_값"):
        ratio = ins["최고단가지역_값"] / ins["최저단가지역_값"]
        b.bullet(
            f"설계 단가는 지역별로 최대 {ratio:.1f}배 차이가 납니다"
            f"({ins['최저단가지역']} vs {ins['최고단가지역']})."
        )
    b.bullet(
        f"{ins['후속대기사업장수']}개 사업장({ins['후속대기비율']}%)이 후속 공고 대기 중입니다. "
        "향후 감리·시공 관련 공고가 이어질 잠재 물량으로 해석할 수 있습니다."
    )
    b.note(
        "※ 본 보고서는 정리된 입찰공고 데이터를 기반으로 자동 생성되었습니다. 해석 문장은 계산된 수치를 "
        "템플릿에 대입한 것으로, 세부 맥락(특정 조합명 등)은 원본 데이터를 함께 확인해 주세요."
    )

    raw_bytes = doc.to_bytes()
    return _fix_shading_alpha(raw_bytes)
