# -*- coding: utf-8 -*-
"""임차범위에 관한 신청인 의견서(임차범위 문구 재정정) — 간단본 HWP·PDF·Word 생성기.

2026. 11. 18. 법원(민사신청과) 연락: 2층에 공부상 203호가 따로 있으므로 임차범위를
「2층 제202호 전유부분 전부 12.37㎡」로 괄호 없이 적어 범위를 확실히 할 것.
계약서 등의 「203호」는 현재 현장에서 실제로 운영되는 203호(＝공부상 202호)로서
공부상 203호와 다르다는 점만 짧게 밝힌다.

임차범위는 build_hwp.C["임차범위"] 한 곳에서 정한다.
사용법: python build_bojeong3.py   → 서류/보정서_임차범위_재정정/ 에 .hwp·PDF/·Word/ 생성
"""
import os

import build_hwp
from build_hwp import (C, V, VARIANTS, DocxBuilder, HwpBuilder, PdfBuilder, doc_schedule1,
                       doc_schedule2, hangul_from_word, 갑)
from build_bojeong import Docx15, Hwp200, Pdf200
import build_bojeong2 as b2

B = {
    "사건번호": "2026카임○○○",
    "사건명": "주택임차권등기",
    "작성일": "2026. 11. 18.",     # 제출하는 날로 고칠 것 (정정 별지 2의 작성일에도 들어감)
    "종전임차범위": "제2층 제202호 전유부분 전부 12.37㎡ (임대차계약서상 표시: 나동 제203호)",
}

NAME = "01_임차범위에_관한_신청인_의견서"
HERE = os.path.dirname(os.path.abspath(__file__))


def doc_bojeong3(b):
    sp = " "
    b.title("임차범위에 관한 신청인 의견서")
    b.p(f"사{sp*4}건{sp*4}{B['사건번호']}{sp*2}{B['사건명']}", after=0)
    b.p(f"신{sp}청{sp}인{sp*4}{C['신청인']}", after=0)
    b.p(f"피신청인{sp*4}{C['피신청인']}", after=10)
    b.p("위 사건에 관하여 신청인은 임차범위를 명확히 하라는 귀원의 안내에 따라 임차범위에 관한 의견을 "
        "다음과 같이 제출합니다.")

    b.h("1. 임차범위에 관한 의견")
    b.p("신청취지 6항 임차범위는 다음과 같이 기재함이 정확하다고 생각합니다.")
    b.items([f"종전 : 「{B['종전임차범위']}」",
             f"정정 : **「{C['임차범위']}」**"], indent=4)
    b.p("2층에 공부상 제203호가 따로 있으므로, 혼동이 없도록 공부상 호실과 전유면적만 적었습니다.")

    b.h("2. 현장 표시 「제203호」에 관하여")
    b.p(f"가. 「나동 제203호」는 **현재 현장에서 실제로 운영되는 호실 번호(현장 표시)**이고, 임대차계약서"
        f"({갑('계약서')})도 이 현장 표시를 따라 적은 것으로서 **공부상 제2층 제203호와는 다른 것**입니다.", hang=2)
    b.p("나. 엘리베이터는 2층 복도의 왼쪽 끝에 있으며, **엘리베이터에서 내려 복도를 바라보는 방향**을 "
        "기준으로 위치는 다음과 같습니다(별지 2 사진 1·2).", hang=2)
    b.p(f"  1) 현장 제203호 : **복도를 따라 끝까지 들어간 곳의 왼쪽 호실**로, **공부상 2층 제202호 전유부분 "
        f"전부(12.37㎡)**를 나누지 않고 쓰는 주택이며 신청인이 {C['전입일']}부터 점유하고 있습니다.", hang=4)
    b.p(f"  2) 공부상 제2층 제203호({b2.B['공부203면적']}) : 엘리베이터에서 내리자마자 **바로 오른쪽** 부분으로, "
        f"현장 제206호·제207호로 나뉘어 쓰이고 다른 임차인의 임차권등기({b2.B['기존임차권등기']})가 있습니다.",
        hang=4)
    b.p("다. 이처럼 두 호실은 복도의 서로 반대쪽 끝에 있으므로, **현장 표시 「나동 제203호」는 "
        "공부상 제2층 제203호가 아닙니다.**", hang=2)
    b.table([
        ["공부상 호실", "현장 호실", "위치 (엘리베이터에서\n내려 복도를 볼 때)", "비고"],
        ["**2층 제202호 (12.37㎡)**", "**제203호**", "**복도 끝까지 간 곳**\n**왼쪽**", "**이 사건 임차목적물**\n**(신청인 점유)**"],
        [f"2층 제203호 ({b2.B['공부203면적']})", "제206호, 제207호", "내리자마자\n바로 오른쪽", "임차목적물 아님"],
    ], weights=[26, 22, 26, 26])

    b.h("3. 결론")
    b.p("별지 1·2와 임차목적물 확대 도면의 임차범위도 위와 같이 고쳐 함께 제출하니(종전 별지는 이 정정본으로 "
        f"갈음), 이 사건 임차범위를 **「{C['임차범위']}」**로 하여 주택임차권등기를 명하여 주시기 바랍니다.")

    b.h("첨 부 서 류")
    b.items(["1. 별지 1  부동산 및 임차목적물의 표시(정정)    1통",
             "2. 별지 2  임차목적물 위치 표시 도면(정정)    1통",
             "3. 임차목적물 확대 도면    1통"], indent=4)
    b.keep_last()
    b._p(B["작성일"], "body", "CENTER", before=4, keep=True)
    b._p(f"신청인{sp*2}{C['신청인']}{sp*4}(인)", "body", "RIGHT", before=2, keep=True)
    b._p(f"{C['법원']} 귀중", "sub", "CENTER", before=6)


ZOOM3_IMAGE = os.path.join(HERE, "자료", "확대도면_202호_3차.jpg")


def doc_zoom3(b):
    """3차 확대 도면: 203호 설명은 의견서에 있으므로 빼고, 공부상 202호가 임차목적물이라는 것만 적는다."""
    b.title("임차목적물 확대 도면")
    b.center(f"**(공부상) {C['건물']} 제2층 제202호**")
    b.picture(ZOOM3_IMAGE, width_mm=86)
    b.h("설 명 란")
    b.table([
        ["구분", "내용"],
        ["임차목적물", "**공부상 제2층 제202호 전유부분 전부 (전유면적 12.37㎡)**"],
        ["임차범위", f"**{C['임차범위']}**\n= 빨간색 실선 내부 전체 (확장형 발코니 · 원룸(주방) · 현관 · 욕실)"],
        ["실제 점유", f"신청인이 {C['전입일']} 인도받아 나누지 않은 전체를 **단독 점유** 중"],
        ["점 표시를 하지\n않은 이유", "임차범위가 건물의 일부가 아니라 **공부상 제202호 전유부분 전부**이므로, ㄱ·ㄴ·ㄷ 등의 "
                         "점 표시 없이 **전유부분 외곽 경계(빨간색 실선)** 그대로 표시함"],
        ["도면 출처 · 치수", f"원주시 발급 건축물현황도({C['현황도발급일']}) 2층 평면도 중 원룸형 \"A\" 부분 확대 / "
                          "가로 3,600 × 깊이 1,200 + 2,500 + 2,000mm (위치·경계 설명용, 면적은 대장의 12.37㎡ 기준)"],
    ], weights=[22, 78])
    b.note("※ 빨간색 실선·붉은 글씨·위쪽 제목과 「확장형 발코니 깊이」 표시는 신청인이 설명을 위해 덧붙인 것이고, "
           "파란 상자의 숫자는 그 자리에 인쇄된 도면 치수(단위 mm)를 읽기 쉽게 크게 다시 쓴 것임. 나머지는 "
           "건축물현황도 그대로임.")


SCHEDULES = [("02_별지1_부동산및임차목적물의표시_재정정", doc_schedule1),
             ("03_별지2_위치표시도면_재정정", doc_schedule2),
             ("04_임차목적물_확대도면_재정정", doc_zoom3)]


def main():
    V.update(VARIANTS[0][1])          # 제출용 증거 번호(사실확인서 없음)
    C["작성일"] = B["작성일"]
    C["호실표시이름"] = "현장 표시"     # 별지 2 머리말
    # 확대 도면: 그림 안 「＝ 현장 표시 나동 제203호」 상자 없이, 제목 띠는 공부상 202호만
    b2.make_zoom_image(ZOOM3_IMAGE, side_box=False, t1="공부상 제2층 제202호  (임차목적물)")
    b2.make_plan2_marked()
    build_hwp.PLAN_IMAGES[0] = b2.PLAN1_MARKED
    build_hwp.PLAN_IMAGES[1] = b2.PLAN2_MARKED
    out = os.path.join(HERE, "서류", "보정서_임차범위_재정정")
    pdf_out, word_out = os.path.join(out, "PDF"), os.path.join(out, "Word")
    os.makedirs(pdf_out, exist_ok=True)
    os.makedirs(word_out, exist_ok=True)
    docs = [(NAME, doc_bojeong3, Pdf200, Docx15, Hwp200)] + \
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
