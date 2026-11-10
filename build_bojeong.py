# -*- coding: utf-8 -*-
"""보정서(등록면허세 납부확인서 제출) HWP·PDF·Word 생성기.

2026. 11. 5.자 보정명령(2026. 11. 9. 송달): 등록면허세 7,200원 납부 후 납세번호가 기재된
납부확인서 제출. build_hwp.py 의 서식(여백·글꼴·서명란)을 그대로 쓴다.

사용법: python build_bojeong.py   → 서류/보정서_등록면허세/ 에 .hwp·PDF/·Word/ 생성
"""
import os

from build_hwp import (C, BaseBuilder, DocxBuilder, HwpBuilder, PdfBuilder,
                       hangul_from_word)

B = {
    "사건번호": "2026카임○○○",
    "사건명": "주택임차권등기",
    "보정명령일": "2026. 11. 5.",
    "송달일": "2026. 11. 9.",
    "등록면허세": "7,200",
    "납부처": "원주시",
    # 납부확인서에 적힌 값을 넣으면 보정서에 함께 적힌다. 비워 두면 그 줄을 뺀다
    "납부일": "",          # 예: "2026. 11. 10."
    "납세번호": "",        # 예: 납부확인서의 「납세번호」 또는 「전자납부번호」
    "작성일": "2026. 11. 10.",
}

NAME = "보정서_등록면허세_납부확인서"


# 민사소송규칙 제4조 제2항: A4, 여백 위 45·아래 30·좌우 20mm(build_hwp.MARGIN),
# 글자 12포인트 이상, 줄 간격 200%(1.5줄) 이상
class Pdf200(PdfBuilder):
    def __init__(self):
        super().__init__()
        for st in self.styles.values():
            st.leading = st.fontSize * 2.0


class Docx15(DocxBuilder):
    line = 1.5   # Word 「1.5줄」


class Hwp200(HwpBuilder):
    def _p(self, text, kind, align, indent=0, hang=0, before=0, after=0, keep=False):
        super()._p(text, kind, align, indent, hang, before, after, keep)
        para = self.d.paragraphs[-1]
        self.st.apply_paragraph_format(paragraphs=[para], line_spacing_percent=200)


def doc_bojeong(b):
    b.title("보  정  서")
    # 표(10.5pt) 대신 12pt 본문으로 당사자 표시
    # 띄어쓰기가 PDF에서 합쳐지지 않게 줄바꿈 없는 공백(NBSP)으로 칸을 맞춘다 (항목명 = 한글 4자 폭)
    sp = " "
    b.p(f"사{sp*4}건{sp*4}{B['사건번호']}{sp*2}{B['사건명']}", after=0)
    b.p(f"신{sp}청{sp}인{sp*4}{C['신청인']}", after=0)
    b.p(f"피신청인{sp*4}{C['피신청인']}", after=10)
    b.p(f"위 사건에 관하여 신청인은 {B['보정명령일']}자 보정명령({B['송달일']} 송달)에 따라 "
        "다음과 같이 보정합니다.")
    b.sub("보 정 내 용")
    b.p(f"1. 보정명령의 흠결사항 제1항에 따라 등록면허세 금 {B['등록면허세']}원을 "
        f"{B['납부처']}에 납부하였습니다.", hang=0)
    detail = [f"세    목 : 등록면허세(등록분)", f"납부금액 : 금 {B['등록면허세']}원"]
    if B["납부일"]:
        detail.append(f"납 부 일 : {B['납부일']}")
    if B["납세번호"]:
        detail.append(f"납세번호 : {B['납세번호']}")
    b.items(detail, indent=8)
    b.p("2. 이에 납세번호가 기재된 「취득세(등록면허세) 납부확인서」를 별첨과 같이 제출합니다.")
    b.h("첨 부 서 류")
    b.items(["1. 취득세(등록면허세) 납부확인서    1통"], indent=4)
    # 서명란: 1쪽에 들어가도록 빈 줄 대신 문단 간격으로 띄운다
    b.keep_last()
    b._p(B["작성일"], "body", "CENTER", before=18, keep=True)
    b._p(f"신청인{sp*2}{C['신청인']}{sp*4}(인)", "body", "RIGHT", before=12, keep=True)
    b._p(f"{C['법원']} 귀중", "sub", "CENTER", before=24)


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    out = os.path.join(here, "서류", "보정서_등록면허세")
    pdf_out, word_out = os.path.join(out, "PDF"), os.path.join(out, "Word")
    os.makedirs(pdf_out, exist_ok=True)
    os.makedirs(word_out, exist_ok=True)
    for cls, folder, ext in ((Pdf200, pdf_out, ".pdf"), (Docx15, word_out, ".docx")):
        b = cls()
        doc_bojeong(b)
        b.save(os.path.join(folder, NAME + ext))
    if not hangul_from_word(word_out, out, None, names=[NAME]):
        b = Hwp200()
        doc_bojeong(b)
        b.save(os.path.join(out, NAME + ".hwp"))
    print("saved to", out)


if __name__ == "__main__":
    main()
