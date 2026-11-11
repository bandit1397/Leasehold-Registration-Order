# -*- coding: utf-8 -*-
"""2차 보정서(임차범위 호실·면적 기재, 별지 목록·도면 정정) HWP·PDF·Word 생성기.

2026. 11. 11.자 보정명령: 「계약서에 근거하여 신청취지 임차범위란의 호실과 면적을 기재하고,
별지 목록과 도면도 그에 맞게 작성하여 제출」.

임차범위는 build_hwp.C["임차범위"] 한 곳에서 정하며, 신청취지·별지 1·별지 2에 똑같이 들어간다.

사용법: python build_bojeong2.py   → 서류/보정서_임차범위/ 에 .hwp·PDF/·Word/ 생성
"""
import os

from build_hwp import (C, V, VARIANTS, DocxBuilder, HwpBuilder, PdfBuilder, doc_schedule1,
                       doc_schedule2, hangul_from_word, 갑)
from build_bojeong import Docx15, Hwp200, Pdf200

B = {
    "사건번호": "2026카임○○○",
    "사건명": "주택임차권등기",
    "보정명령일": "2026. 11. 11.",
    "송달일": "2026. 11. 11.",      # 전자소송 「송달일」과 다르면 고칠 것 (기한 = 송달일 + 7일)
    "작성일": "2026. 11. 11.",      # 제출하는 날로 고칠 것 (정정 별지 2의 작성일에도 들어감)
    "종전임차범위": "제2층 제202호 전유부분 전부 (전유면적 12.37㎡)",
    "공부203면적": "31.41㎡",
}

NAME = "01_보정서_임차범위_정정"
SCHEDULES = [("02_별지1_부동산및임차목적물의표시_정정", doc_schedule1),
             ("03_별지2_위치표시도면_정정", doc_schedule2)]


def doc_bojeong2(b):
    sp = " "
    b.title("보  정  서")
    b.p(f"사{sp*4}건{sp*4}{B['사건번호']}{sp*2}{B['사건명']}", after=0)
    b.p(f"신{sp}청{sp}인{sp*4}{C['신청인']}", after=0)
    b.p(f"피신청인{sp*4}{C['피신청인']}", after=10)
    b.p(f"위 사건에 관하여 신청인은 {B['보정명령일']}자 보정명령({B['송달일']} 송달)에 따라 "
        "다음과 같이 보정합니다.")
    b.sub("보 정 내 용")

    b.h("1. 신청취지 임차범위의 정정")
    b.p(f"신청취지 6항 임차범위 「{B['종전임차범위']}」를 다음과 같이 정정합니다.")
    b.p(f"「{C['임차범위']}」", indent=4)
    b.p("이에 따라 정정된 신청취지 전체는 다음과 같습니다.")
    b.table([
        ["항목", "기재 내용"],
        ["1. 임대차계약일자", C["계약일"]],
        ["2. 임차보증금액 / 차임", f"보증금 금 {C['보증금']}원 / 차임 월 금 {C['차임']}원"],
        ["3. 주민등록일자", C["전입일"]],
        ["4. 점유개시일자", C["전입일"]],
        ["5. 확정일자", C["전입일"]],
        ["6. 임차범위", C["임차범위"]],
    ], weights=[30, 70])

    b.h("2. 계약서에 근거한 호실과 면적")
    b.p(f"가. 임대차계약서({갑('계약서')})는 임차목적물을 「나동 제203호」로 표시하고, 면적은 "
        "「공부상 전용면적 또는 연면적(실측면적)」에 따르도록 정하고 있습니다.", hang=2)
    b.p("나. 계약서의 「제203호」는 현장 출입문에 붙은 호실 번호이며 공부상 호실 번호가 아닙니다. "
        f"신청인이 {C['전입일']} 인도받아 현재까지 점유하는 주택은 건축물현황도에 「202호~502호」로 표기된 "
        f"부분, 즉 공부상 제2층 제202호 전유부분 전체입니다({갑('대장')}, {갑('사진')}, {갑('별지2')}). "
        "이 건물 2층은 현장 제202호부터 제207호까지 모든 호실의 번호가 공부상 호실 번호와 다릅니다.", hang=2)
    b.p("다. 따라서 계약서가 정한 기준인 공부상 전용면적에 따른 임차목적물의 면적은 공부상 제2층 "
        "제202호의 전유면적 12.37㎡입니다(집합건축물대장 및 등기사항전부증명서 표제부, "
        f"{갑('등기202')}, {갑('대장')}).", hang=2)
    b.p(f"라. 공부상 제2층 제203호(전유면적 {B['공부203면적']})는 신청인이 점유하는 부분과 위치와 크기가 "
        "다른 별개의 전유부분이고, 이미 다른 임차인의 주택임차권등기가 마쳐져 있습니다"
        f"({갑('등기203')}). 계약서의 「제203호」를 공부상 제203호로 보아 임차범위를 정하면 신청인이 "
        "점유하지 않는 부분에 등기가 이루어지게 됩니다.", hang=2)
    b.p("마. 이에 임차범위에는 등기가 이루어질 공부상 호실과 면적을 적고, 임대차계약서상 표시를 괄호로 "
        "함께 적어 계약서와 등기 목적물의 대응관계가 드러나도록 하였습니다.", hang=2)

    b.h("3. 별지 목록과 도면의 정정")
    b.p("별지 1(부동산 및 임차목적물의 표시)의 4항 임차범위를 위 1항과 같이 고치고 계약서의 면적 기준을 "
        "비고에 적었으며, 별지 2(임차목적물 위치 표시 도면)에도 같은 임차범위와 계약서상 표시를 적은 "
        "정정본을 제출합니다. 종전 별지 1·2는 이 정정본으로 갈음합니다.")

    b.h("첨 부 서 류")
    b.items(["1. 별지 1  부동산 및 임차목적물의 표시(정정)    1통",
             "2. 별지 2  임차목적물 위치 표시 도면(정정)    1통"], indent=4)
    b.keep_last()
    b._p(B["작성일"], "body", "CENTER", before=18, keep=True)
    b._p(f"신청인{sp*2}{C['신청인']}{sp*4}(인)", "body", "RIGHT", before=12, keep=True)
    b._p(f"{C['법원']} 귀중", "sub", "CENTER", before=24)


def main():
    V.update(VARIANTS[0][1])          # 제출용 증거 번호(사실확인서 없음)
    C["작성일"] = B["작성일"]          # 정정 별지 2의 작성일
    here = os.path.dirname(os.path.abspath(__file__))
    out = os.path.join(here, "서류", "보정서_임차범위")
    pdf_out, word_out = os.path.join(out, "PDF"), os.path.join(out, "Word")
    os.makedirs(pdf_out, exist_ok=True)
    os.makedirs(word_out, exist_ok=True)
    docs = [(NAME, doc_bojeong2, Pdf200, Docx15, Hwp200)] + \
           [(n, fn, PdfBuilder, DocxBuilder, HwpBuilder) for n, fn in SCHEDULES]
    for name, fn, pdf_cls, docx_cls, _ in docs:
        for cls, folder, ext in ((pdf_cls, pdf_out, ".pdf"), (docx_cls, word_out, ".docx")):
            b = cls()
            fn(b)
            b.save(os.path.join(folder, name + ext))
    if not hangul_from_word(word_out, out, None, names=[n for n, *_ in docs]):
        for name, fn, _, _, hwp_cls in docs:
            b = hwp_cls()
            fn(b)
            b.save(os.path.join(out, name + ".hwp"))
    print("saved to", out)


if __name__ == "__main__":
    main()
