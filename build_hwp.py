# -*- coding: utf-8 -*-
"""주택임차권등기명령 신청서류 일체 HWP·PDF 생성기.

전제: 공부상 ○○빌 나동 제202호 전유부분 전체 = 계약서·현장 표시 제203호 전체.
확인되지 않은 사항은 모두 000 으로 남긴다.

사용법: python build_hwp.py [출력폴더]
필요: pip install python-hwpx reportlab python-docx   (PDF 글꼴: Windows 바탕·맑은 고딕)
"""
import os
import re
import sys
from xml.sax.saxutils import escape

from hwpx import HwpxDocument

MM = 283.465  # 1mm = 283.465 HWPUNIT
FONT_DIR = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
PLAN_IMAGES = [
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "자료", name)
    for name in ("별지2_사진1_공부상호실.jpg", "별지2_사진2_현재호실.jpg")
]

# ---------------------------------------------------------------------------
# 공통 기재사항 (수정 시 여기만 고치면 전 서류에 반영)
# ---------------------------------------------------------------------------
C = {
    "신청인": "○○○",
    "피신청인": "△△△",
    "법원": "춘천지방법원 원주지원",
    "경매사건": "2026타경○○",
    "소재지": "강원특별자치도 원주시 000",
    "도로명": "000",
    "건물": "○○빌 나동",
    "목적물": "공부상 ○○빌 나동 제202호 전유부분 전체(계약서 및 현장 출입문 표시 제203호)",
    "보증금": "5,000,000",
    "차임": "230,000",
    "관리비": "80,000",
    "연체월수": "4",
    "연체차임": "920,000",
    "2기차임": "460,000",
    "연체관리비": "320,000",
    "연체합계": "1,240,000",
    "잔액": "3,760,000",
    "잔액_차임만": "4,080,000",
    "전입일": "2026. 2. 23.",
    "종료일": "2026. 10. 31.",
    "작성일": "2026. 11. 3.",      # 임차권등기명령 신청일 (신청서·별지·정산서 등)
    "발송일": "2026. 11. 2.",      # 내용증명 발송일
    "회신기한": "2026. 11. 9.",    # 내용증명의 정산내역 회신·반환 기한
}


def image_size(path):
    from reportlab.lib.utils import ImageReader
    return ImageReader(path).getSize()


class BaseBuilder:
    """서류 본문 함수(doc_*)가 쓰는 공통 인터페이스. 출력 형식은 하위 클래스가 정한다."""

    def title(self, text):
        self._p(text, "title", "CENTER", after=14)

    def sub(self, text, align="CENTER"):
        self._p(text, "sub", align, before=6, after=6)

    def h(self, text):
        self._p(text, "head", "JUSTIFY", before=8, after=2, keep=True)

    def p(self, text, indent=0, hang=0):
        self._p(text, "body", "JUSTIFY", indent=indent, hang=hang)

    def items(self, lines, indent=4):
        for line in lines:
            self.p(line, indent=indent)

    def center(self, text):
        self._p(text, "body", "CENTER")

    def right(self, text):
        self._p(text, "body", "RIGHT")

    def note(self, text):
        self._p(text, "note", "JUSTIFY")

    def small(self, text):
        self._p(text, "small", "JUSTIFY")

    def blank(self):
        self._p("", "body", "LEFT")

    def caption(self, text):
        self._p(text, "head", "CENTER", before=6, after=2, keep=True)

    def sign(self, who="신청인", name=C["신청인"], date=C["작성일"], court=True):
        self.blank()
        self.center(date)
        self.blank()
        self.right(f"{who}  {name}  (서명 또는 날인)")
        if court:
            self.blank()
            self._p(f"{C['법원']} 귀중", "sub", "CENTER", before=12)


class HwpBuilder(BaseBuilder):
    def __init__(self):
        self.d = HwpxDocument.new()
        st = self.d.styles
        self.st = st
        self.chars = {
            "title": st.ensure_run(bold=True, size=20, font="바탕"),
            "sub": st.ensure_run(bold=True, size=14, font="바탕"),
            "head": st.ensure_run(bold=True, size=12, font="바탕"),
            "body": st.ensure_run(size=12, font="바탕"),
            "small": st.ensure_run(size=10, font="바탕"),
            "note": st.ensure_run(size=10, font="바탕", color="#C00000"),
        }
        self.d.page.set_margins(left=int(20 * MM), right=int(20 * MM),
                                top=int(20 * MM), bottom=int(15 * MM))
        self._next_break = False

    def _p(self, text, kind, align, indent=0, hang=0, before=0, after=0,
           keep=False):
        para = self.d.add_paragraph(text, char_pr_id_ref=self.chars[kind])
        fmt = dict(alignment=align, line_spacing_percent=160,
                   indent_left_mm=indent, first_line_indent_mm=hang,
                   spacing_before_pt=before, spacing_after_pt=after)
        if keep:
            fmt["keep_with_next"] = True
        if self._next_break:
            fmt["page_break_before"] = True
            self._next_break = False
        self.st.apply_paragraph_format(paragraphs=[para], **fmt)

    def page_break(self):
        self._next_break = True

    def table(self, rows, weights=None, header=True):
        t = self.d.add_table(len(rows), len(rows[0]), width=int(170 * MM))
        if weights:
            t.set_column_widths(weights)
        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                t.set_cell_text(r, c, val)
                if header and r == 0:
                    t.set_cell_shading(r, c, "#E7E6E6")

    def picture(self, path, width_mm):
        w, h = image_size(path)
        with open(path, "rb") as f:
            self.d.add_picture(f.read(), "jpg", width_mm=width_mm,
                               height_mm=width_mm * h / w, align="CENTER")

    def save(self, path):
        self.d.save_to_path(path)


class PdfBuilder(BaseBuilder):
    _fonts_ready = False

    def __init__(self):
        from reportlab.lib.colors import HexColor
        from reportlab.lib.styles import ParagraphStyle
        if not PdfBuilder._fonts_ready:
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont
            pdfmetrics.registerFont(TTFont(
                "Batang", os.path.join(FONT_DIR, "batang.ttc"), subfontIndex=0))
            pdfmetrics.registerFont(TTFont(
                "MalgunBd", os.path.join(FONT_DIR, "malgunbd.ttf")))
            PdfBuilder._fonts_ready = True
        spec = {  # 글꼴, 크기 — 바탕에는 굵은꼴이 없어 제목은 맑은 고딕 굵게
            "title": ("MalgunBd", 20), "sub": ("MalgunBd", 14),
            "head": ("MalgunBd", 12), "body": ("Batang", 12),
            "small": ("Batang", 10), "note": ("Batang", 10),
        }
        self.styles = {
            k: ParagraphStyle(k, fontName=f, fontSize=s, leading=s * 1.6,
                              wordWrap="CJK",
                              textColor=HexColor("#C00000" if k == "note" else "#000000"))
            for k, (f, s) in spec.items()
        }
        self.cell = ParagraphStyle("cell", fontName="Batang", fontSize=10.5,
                                   leading=15, wordWrap="CJK")
        self.flow = []

    def _p(self, text, kind, align, indent=0, hang=0, before=0, after=0,
           keep=False):
        from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
        from reportlab.lib.units import mm
        from reportlab.platypus import Paragraph
        aligns = {"CENTER": TA_CENTER, "JUSTIFY": TA_JUSTIFY,
                  "LEFT": TA_LEFT, "RIGHT": TA_RIGHT}
        base = self.styles[kind]
        style = base.clone(base.name, alignment=aligns[align],
                           leftIndent=indent * mm, firstLineIndent=hang * mm,
                           spaceBefore=before, spaceAfter=after,
                           keepWithNext=keep)
        self.flow.append(Paragraph(escape(text) or "&nbsp;", style))

    def page_break(self):
        from reportlab.platypus import PageBreak
        self.flow.append(PageBreak())

    def table(self, rows, weights=None, header=True):
        from reportlab.lib.colors import HexColor, black
        from reportlab.lib.units import mm
        from reportlab.platypus import Paragraph, Spacer, Table, TableStyle
        weights = weights or [1] * len(rows[0])
        widths = [170 * mm * w / sum(weights) for w in weights]
        data = [[Paragraph(escape(v).replace("\n", "<br/>") or "&nbsp;", self.cell)
                 for v in row] for row in rows]
        t = Table(data, colWidths=widths, repeatRows=1 if header else 0)
        cmds = [("GRID", (0, 0), (-1, -1), 0.5, black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]
        if header:
            cmds.append(("BACKGROUND", (0, 0), (-1, 0), HexColor("#E7E6E6")))
        t.setStyle(TableStyle(cmds))
        self.flow += [Spacer(1, 3), t, Spacer(1, 3)]

    def picture(self, path, width_mm):
        from reportlab.lib.units import mm
        from reportlab.platypus import Image
        w, h = image_size(path)
        self.flow.append(Image(path, width=width_mm * mm,
                               height=width_mm * mm * h / w))

    def save(self, path):
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import mm
        from reportlab.platypus import SimpleDocTemplate

        def page_no(canvas, doc):
            canvas.setFont("Batang", 9)
            canvas.drawCentredString(A4[0] / 2, 8 * mm, f"- {doc.page} -")

        SimpleDocTemplate(path, pagesize=A4, leftMargin=20 * mm,
                          rightMargin=20 * mm, topMargin=20 * mm,
                          bottomMargin=15 * mm, title=os.path.basename(path)[:-4],
                          author=C["신청인"]).build(
            self.flow, onFirstPage=page_no, onLaterPages=page_no)


# 빈칸 표시(000, 00. 00., 000000-0000000 등). 5,000,000 같은 금액은 제외.
BLANK_RE = re.compile(r"(?<![\d,])0{2,}(?:[.\-]\s?0{2,})*(?![\d,])")


class DocxBuilder(BaseBuilder):
    """Word(.docx) 출력 — 빈칸(000)은 노란 형광펜으로 표시한다."""

    def __init__(self):
        from docx import Document
        from docx.shared import Mm
        self.d = Document()
        sec = self.d.sections[0]
        sec.page_width, sec.page_height = Mm(210), Mm(297)
        sec.left_margin = sec.right_margin = sec.top_margin = Mm(20)
        sec.bottom_margin = Mm(15)
        # Word 가 한글과 숫자 사이에 자동으로 넣는 간격("제 202 호")을 끈다
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        ppr = self.d.styles["Normal"].element.get_or_add_pPr()
        for tag in ("w:autoSpaceDE", "w:autoSpaceDN"):
            el = OxmlElement(tag)
            el.set(qn("w:val"), "0")
            ppr.append(el)
        self.spec = {  # 크기, 굵게, 색
            "title": (20, True, None), "sub": (14, True, None),
            "head": (12, True, None), "body": (12, False, None),
            "small": (10, False, None), "note": (10, False, "C00000"),
        }
        self._next_break = False
        self._first = True

    @staticmethod
    def _runs(para, text, size, bold=False, color=None):
        from docx.enum.text import WD_COLOR_INDEX
        from docx.oxml.ns import qn
        from docx.shared import Pt, RGBColor
        pos = 0
        pieces = []
        for m in BLANK_RE.finditer(text):
            pieces += [(text[pos:m.start()], False), (m.group(), True)]
            pos = m.end()
        pieces.append((text[pos:], False))
        for chunk, blank in pieces:
            if not chunk:
                continue
            run = para.add_run(chunk)
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.name = "바탕"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "바탕")
            if color:
                run.font.color.rgb = RGBColor.from_string(color)
            if blank:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW

    def _para(self):
        # 새 문서의 첫 문단은 빈 문단을 재사용하지 않고 그대로 추가한다
        para = self.d.add_paragraph()
        if self._next_break:
            para.paragraph_format.page_break_before = True
            self._next_break = False
        return para

    def _p(self, text, kind, align, indent=0, hang=0, before=0, after=0,
           keep=False):
        from docx.enum.text import WD_ALIGN_PARAGRAPH as A
        from docx.shared import Mm, Pt
        para = self._para()
        pf = para.paragraph_format
        pf.alignment = {"CENTER": A.CENTER, "JUSTIFY": A.JUSTIFY,
                        "LEFT": A.LEFT, "RIGHT": A.RIGHT}[align]
        pf.left_indent = Mm(indent)
        pf.first_line_indent = Mm(hang)
        pf.space_before, pf.space_after = Pt(before), Pt(after)
        pf.line_spacing = 1.3
        pf.keep_with_next = keep
        size, bold, color = self.spec[kind]
        self._runs(para, text, size, bold, color)

    def page_break(self):
        self._next_break = True

    def table(self, rows, weights=None, header=True):
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Mm
        weights = weights or [1] * len(rows[0])
        t = self.d.add_table(rows=len(rows), cols=len(rows[0]))
        t.style = "Table Grid"
        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                cell = t.cell(r, c)
                cell.width = Mm(170 * weights[c] / sum(weights))
                for i, line in enumerate(val.split("\n")):
                    para = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
                    para.paragraph_format.space_after = 0
                    self._runs(para, line, 10.5)
                if header and r == 0:
                    shd = OxmlElement("w:shd")
                    shd.set(qn("w:val"), "clear")
                    shd.set(qn("w:fill"), "E7E6E6")
                    cell._tc.get_or_add_tcPr().append(shd)
        self.d.add_paragraph()

    def picture(self, path, width_mm):
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Mm
        para = self._para()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.add_run().add_picture(path, width=Mm(width_mm))

    def save(self, path):
        self.d.core_properties.author = C["신청인"]
        self.d.save(path)


# ---------------------------------------------------------------------------
# 서류별 본문
# ---------------------------------------------------------------------------
def doc_application(b):
    b.title("주택임차권등기명령 신청서")
    b.table([
        ["구분", "내용"],
        ["사건번호", "2026카임000 (접수 시 법원이 부여)"],
        ["신청인(임차인)", f"성명: {C['신청인']} / 주민등록번호: 000000-0000000\n"
                       "주소: 000\n송달장소: 000 / 연락처: 000-0000-0000"],
        ["피신청인(임대인)", f"성명: {C['피신청인']}\n주소: 000 (등기사항전부증명서상 소유자 주소)\n"
                         "연락처: 000-0000-0000"],
    ], weights=[25, 75])

    b.sub("신 청 취 지")
    b.p(f"별지 1 목록 기재 건물, 즉 {C['목적물']}에 관하여 아래와 같은 주택임차권등기를 명한다.")
    b.center("아    래")
    b.table([
        ["항목", "기재 내용"],
        ["1. 임대차계약일자", "2026. 00. 00."],
        ["2. 임차보증금액 / 차임", f"보증금 금 {C['보증금']}원 / 차임 월 금 {C['차임']}원\n(관리비 월 금 {C['관리비']}원 별도)"],
        ["3. 주민등록일자", C["전입일"]],
        ["4. 점유개시일자", C["전입일"]],
        ["5. 확정일자", C["전입일"]],
        ["6. 임차범위", "공부상 나동 제202호 전유부분 전체 (전유면적 000㎡)\n"
                    "※ 계약서·현장 출입문 표시 제203호와 동일한 공간"],
        ["7. 임대차 종료일", f"{C['종료일']} (차임 연체를 이유로 한 피신청인의 해지)"],
    ], weights=[30, 70])
    b.p("신청비용은 피신청인이 부담한다.")
    b.p("라는 결정을 구합니다.")

    b.sub("신 청 이 유")
    b.h("1. 임대차계약의 체결")
    b.p(f"신청인은 2026. 00. 00. 공인중개사 000의 중개로 피신청인과 별지 1 기재 건물에 관하여 "
        f"임차보증금 금 {C['보증금']}원, 월차임 금 {C['차임']}원(관리비 월 금 {C['관리비']}원 별도), "
        "임대차기간 2026. 00. 00.부터 0000. 00. 00.까지로 하는 임대차계약을 체결하고, "
        "보증금 전액을 지급하였습니다.", hang=4)
    b.h("2. 임차목적물의 표시 — 공부상 제202호 전유부분 전체")
    b.p("위 임대차계약서와 현장 출입문에는 임차목적물이 「나동 제203호」로 표시되어 있습니다. "
        "그러나 건축물대장·건축물현황도 및 등기사항전부증명서를 대조한 결과, 신청인이 계약 당시 "
        "안내받아 계약하고 인도받은 공간은 공부상 나동 제202호 전유부분 전체와 위치·구조·면적이 "
        "동일한 하나의 주거공간임이 최종 확인되었습니다. 신청인은 공부상 제202호의 일부만을 임차한 "
        "것이 아니라 공부상 제202호 전유부분 전체를 그대로 임차하였고, 공부상 제203호를 임차하거나 "
        "점유한 사실은 없습니다. 그 경위는 별첨 신청이유 보충서와 같습니다.", hang=4)
    b.h("3. 대항력 및 우선변제권의 취득")
    b.p(f"신청인은 {C['전입일']} 위 목적물을 인도받아 점유를 개시하고, 같은 날 주민등록 전입신고를 "
        "마쳤으며, 같은 날 임대차계약서에 확정일자를 받았습니다. 위 대항요건 및 확정일자는 "
        f"목적물에 관한 경매개시결정 기입등기일(2026. 4. 00.) 이전에 갖추어졌습니다.", hang=4)
    b.h("4. 임대차의 종료")
    b.p(f"신청인은 2026. 00.분부터 2026. 00.분까지 {C['연체월수']}개월분 차임 합계 금 {C['연체차임']}원을 "
        f"연체하였고, 이는 2기의 차임액 금 {C['2기차임']}원을 초과합니다. 피신청인은 2026. 10. 26. "
        "신청인에게 「2026. 10. 30.까지 밀린 월세를 납부하지 않으면 계약을 해지한다」는 취지의 "
        "문자메시지를 보냈고, 신청인은 위 기한까지 이를 납부하지 못하였습니다. 따라서 위 임대차계약은 "
        f"민법 제640조 및 임대차계약서 제000조에 따라 {C['종료일']} 종료되었습니다.", hang=4)
    b.h("5. 임차보증금의 미반환")
    b.p(f"임대차가 종료되었음에도 신청인은 임차보증금 금 {C['보증금']}원을 전혀 반환받지 못하였습니다. "
        f"미납 차임 금 {C['연체차임']}원과 미납 관리비 금 {C['연체관리비']}원(합계 금 {C['연체합계']}원)을 "
        f"모두 공제하더라도 반환받지 못한 보증금은 금 {C['잔액']}원이고, 관리비 공제가 인정되지 않을 "
        f"경우에는 금 {C['잔액_차임만']}원입니다. 정확한 미납액은 000 확인 후 보정하겠습니다. "
        f"신청인은 {C['발송일']} 피신청인에게 내용증명우편으로 보증금 반환을 청구하였습니다"
        "(갑 제13호증).", hang=4)
    b.h("6. 경매절차와의 관계")
    b.p(f"위 건물에 관하여 {C['법원']} {C['경매사건']} 부동산경매절차가 진행 중이며, 신청인은 배당요구 "
        "종기 내에 배당요구를 마쳤습니다(배당요구 접수일 000). 본 신청은 위 배당요구를 그대로 유지하면서, "
        "신청인이 주거를 옮기더라도 이미 취득한 대항력과 우선변제권을 유지하기 위한 것입니다.", hang=4)
    b.h("7. 결론")
    b.p("이에 주택임대차보호법 제3조의3에 따라 신청취지와 같은 결정을 구합니다.", hang=4)

    b.sub("소 명 방 법 및 첨 부 서 류", align="LEFT")
    b.items([
        "1. 건물등기사항전부증명서(나동 제202호)          1통",
        "2. 임대차계약증서 사본(확정일자 날인면 포함)      1통",
        "3. 주민등록초본(주소변동사항 포함)                1통",
        "4. 건축물대장 및 건축물현황도                      각 1통",
        "5. 신청이유 보충서                                 1통",
        "6. 별지 1 부동산 및 임차목적물의 표시            1통",
        "7. 별지 2 임차목적물 위치 표시 도면              1통",
        "8. 증거설명서 및 갑 제1호증 내지 갑 제13호증     각 1통",
        "9. 송달료·인지·등록면허세·등기신청수수료 납부서  각 1통",
    ])
    b.sign()


def doc_supplement(b):
    b.title("신청이유 보충서")
    b.p(f"사    건 : 2026카임000 주택임차권등기명령 신청")
    b.p(f"신 청 인 : {C['신청인']}")
    b.p(f"피신청인 : {C['피신청인']}")
    b.blank()
    b.p("위 사건에 관하여 신청인은 임차목적물의 동일성과 임대차 종료 사유를 아래와 같이 보충합니다.")

    b.h("1. 계약 체결 경위")
    b.p(f"신청인은 2026. 00. 00. 중개보조원 000의 안내로 피신청인 소유 {C['건물']} 2층을 임차하였습니다. "
        "현장 2층에는 제201호부터 제207호까지 7개의 출입문이 호실 번호와 함께 표시되어 있었고, "
        "신청인에게 안내된 공간은 출입문에 제203호로 표시된 독립 원룸이었습니다. "
        "임대차계약서에도 목적물이 「나동 제203호」로 기재되었습니다.", hang=4)

    b.h("2. 공부상 호실과 현장 호실의 대응관계")
    b.items([
        "가. 공부(건축물대장·등기부)상 2층에는 제201호부터 제204호까지 4개 호실만 존재합니다.",
        "나. 현장에서는 제201호부터 제207호까지 7개 호실 번호가 사용되고 있습니다.",
        "다. 공부상 제201호가 현장에서 제201호와 제202호로 나뉘어 표시·사용되면서, 공부상 제202호는 "
        "공간의 분할 없이 그대로 현장 출입문에 제203호로 표시되었습니다.",
        "라. 따라서 신청인이 계약한 현장 제203호는 공부상 제202호의 일부가 아니라, 공부상 제202호 "
        "전유부분 전체와 위치·구조·경계·면적이 동일한 하나의 주거공간입니다.",
        "마. 2층 평면도(별지 2)에는 신청인이 점유하는 공간이 원룸형 \"A\"(202호~502호)로 표기되어 "
        "있고, 나머지 공간은 201호~501호, 203호~503호, 204호~504호로 표기되어 있어, 이 공간이 "
        "공부상 제202호임이 도면 자체로 드러납니다.",
        "바. 위 사실은 건축물현황도(발급일 2026. 00. 00.), 등기사항전부증명서상 제202호 전유면적 "
        "000㎡와 현장 실측 면적 000㎡의 대조로 확인됩니다.",
    ])

    b.h("3. 다른 임차인 관련 정황")
    b.p("공부상 제203호에 관하여 임차권등기를 마친 다른 임차인의 방은 현장에서 제205호로 표시되어 "
        "사용되고 있습니다. 이는 호실 표시의 불일치가 신청인에게만 생긴 개별적 착오가 아니라 "
        "건물 전체의 관행적 표시 방식에서 비롯된 것임을 보여줍니다(갑 제11호증, 확보 경로 000).", hang=4)

    b.h("4. 계속 점유")
    b.p(f"신청인은 계약 당시 위 불일치를 알지 못하였고, {C['전입일']}부터 현재까지 현장 제203호, "
        "즉 공부상 제202호 전유부분 전체를 거실·주방·욕실을 포함하여 단독으로 계속 점유하고 있습니다.", hang=4)

    b.h("5. 차임 연체액이 2기 차임액에 달한 사실")
    b.items([
        f"가. 월차임은 금 {C['차임']}원이고, 관리비 월 금 {C['관리비']}원은 차임과 별도로 약정되었습니다.",
        f"나. 연체 차임은 2026. 00.분부터 2026. 00.분까지 {C['연체월수']}개월분 합계 금 {C['연체차임']}원으로, "
        f"2기 차임액 금 {C['2기차임']}원을 초과합니다. 위 금액에는 관리비·연체료가 포함되어 있지 않습니다.",
        "다. 피신청인은 2026. 10. 26. 문자메시지로 2026. 10. 30.까지 미납 시 해지한다고 통지하였고, "
        f"기한이 지나 임대차는 {C['종료일']} 종료되었습니다(갑 제4호증).",
    ])

    b.h("6. 경매절차에서의 표시 정정 경위")
    b.p(f"신청인은 {C['법원']} {C['경매사건']} 경매절차에서 배당요구를 마친 뒤, 목적물의 호실 표시를 "
        "정정하는 보정서와 점유공간 표시 도면을 제출하였습니다(접수일·접수번호 000). 만일 종전 보정서에 "
        "목적물이 제202호의 「일부」로 표시되어 있다면, 이는 당시 확인이 완료되지 않았던 데 따른 것이며, "
        "최종 확인 결과에 따라 신청인은 목적물을 공부상 제202호 전유부분 전체로 정정하여 소명합니다. "
        "경매법원에도 같은 취지의 보정서를 추가로 제출할 예정입니다(제출일 000).", hang=4)

    b.h("7. 소명의 취지")
    b.p("신청인이 임차한 목적물은 공부상 나동 제202호 전유부분 전체이며, 계약서 및 현장의 「제203호」는 "
        "같은 공간을 가리키는 현장 호실 표시에 불과합니다. 따라서 임차권등기는 공부상 나동 제202호 "
        "전유부분 전체에 관하여 이루어져야 함을 소명합니다.", hang=4)
    b.sign()


def doc_schedule1(b):
    b.title("별지 1  부동산 및 임차목적물의 표시")
    b.h("1. 1동의 건물의 표시")
    b.items([
        f"소재지 : {C['소재지']}",
        f"도로명주소 : {C['도로명']}",
        f"건물명칭 및 번호 : {C['건물']}",
        "구조·지붕·층수·면적 : 000 (등기사항전부증명서 표제부 기재를 그대로 옮겨 적음)",
    ])
    b.h("2. 전유부분의 건물의 표시")
    b.items([
        "건물번호 : 제2층 제202호",
        "구조 : 000",
        "면적 : 000㎡",
    ])
    b.h("3. 대지권의 표시")
    b.items([
        "토지의 표시 : 000",
        "대지권의 종류 : 소유권대지권",
        "대지권의 비율 : 000분의 000",
    ])
    b.h("4. 임차목적물(임차범위)")
    b.p("위 제2층 제202호 전유부분 전체 (전유면적 000㎡)", indent=4)
    b.h("5. 비고")
    b.p("임대차계약서 및 현장 출입문에는 「제203호」로 표시되어 있으나, 이는 공부상 나동 제202호 "
        "전유부분 전체와 동일한 공간을 가리키는 현장 호실 표시임. 끝.", indent=4)
    b.blank()
    b.note("※ 1~3항은 최신 등기사항전부증명서의 표제부 기재를 한 글자도 바꾸지 않고 옮겨 적으십시오.")


def doc_schedule2(b):
    b.title("별지 2  임차목적물 위치 표시 도면")
    b.center(f"{C['건물']} 2층 평면도 — 빨간색 실선 내부 전체가 임차목적물")
    b.caption("[사진 1]  공부상 호실 표시")
    b.picture(PLAN_IMAGES[0], width_mm=128)
    b.caption("[사진 2]  현재 현장 호실 표시")
    b.picture(PLAN_IMAGES[1], width_mm=128)
    b.page_break()
    b.h("1. 공부상 호실과 현장 호실의 대응")
    b.table([
        ["공부상 호실 (사진 1)", "현장 호실 (사진 2)", "비고"],
        ["제201호", "제201호, 제202호",
         "공부상 1개 호실을 현장에서 2개로 나누어 사용 (사진 2 주황색 선)"],
        ["제202호 (빨간색)", "제203호 (빨간색)",
         "임차목적물 — 나누지 않고 전체를 그대로 사용, 호실 번호만 바뀜"],
        ["제203호, 제204호", "제204호 ~ 제207호",
         "공부상 2개 호실을 현장에서 4개로 나누어 사용 (사진 2 주황색 선)"],
    ], weights=[24, 24, 52])
    b.h("2. 임차목적물의 표시")
    b.table([
        ["구분", "내용"],
        ["임차목적물", "사진 1·2의 빨간색 실선으로 둘러싼 부분 전체 (두 사진의 경계가 같음)"],
        ["도면상 호실 표기", "원룸형 \"A\" (202호~502호) → 2층은 공부상 제202호"],
        ["현장·계약서 표시", "제203호"],
        ["위치", "2층 도면 우측 상단 끝, 기준선 X3~X4 사이"],
        ["구성", "확장형 발코니, 원룸(주방 포함), 현관, 욕실 (아래쪽 창고는 제외)"],
        ["치수(도면 중심선 기준)", "가로 3,600mm × 세로 5,700mm\n"
                               "(발코니 1,200 + 원룸 2,500 + 현관·욕실 2,000)"],
        ["등기상 전유면적", "000㎡"],
        ["파란색 실선", "다른 호실 — 임차목적물 아님"],
        ["도면 출처", "000 (발급일 2026. 00. 00.)"],
    ], weights=[28, 72])
    b.blank()
    b.p("본 도면 사진 1·2에서 빨간색 실선으로 표시한 부분 전체는 신청인 ○○○이 2026. 2. 23. 인도받아 "
        "점유하고 있는 현장 출입문 표시 제203호로서, 도면에 원룸형 \"A\"(202호~502호)로 표기된 "
        "공부상 ○○빌 나동 제202호 전유부분 전체에 해당하는 임차목적물임을 표시한다.")
    b.items([
        f"작성일 : {C['작성일']}",
        f"작성자 : 신청인 {C['신청인']}  (서명 또는 날인)",
    ], indent=0)


def doc_settlement(b):
    b.title("보증금 및 미납차임·관리비 정산 내역서")
    b.p(f"신청인 : {C['신청인']}    피신청인 : {C['피신청인']}")
    b.h("1. 임차보증금")
    b.table([
        ["항목", "금액", "비고"],
        ["임차보증금 총액", f"{C['보증금']}원", "임대차계약서"],
        ["반환받은 금액", "0원", f"{C['작성일']} 현재"],
        ["미반환 보증금", f"{C['보증금']}원", ""],
    ], weights=[35, 30, 35])
    b.h("2. 미납 차임 (관리비 제외)")
    b.table([
        ["항목", "금액", "산정"],
        ["월 차임", f"{C['차임']}원", "계약서 제000조"],
        ["미납 기간", f"{C['연체월수']}개월", "2026. 00.분 ~ 2026. 00.분"],
        ["미납 차임 합계", f"{C['연체차임']}원", f"{C['차임']} × {C['연체월수']}"],
        ["2기 차임액", f"{C['2기차임']}원", f"{C['차임']} × 2 → 연체액이 2배 초과"],
    ], weights=[35, 30, 35])
    b.h("3. 미납 관리비 (차임과 별도)")
    b.table([
        ["항목", "금액", "산정"],
        ["월 관리비", f"{C['관리비']}원", "계약서·부과내역"],
        ["미납 관리비 합계", f"{C['연체관리비']}원", f"{C['관리비']} × {C['연체월수']}"],
    ], weights=[35, 30, 35])
    b.h("4. 정산")
    b.table([
        ["항목", "금액"],
        ["임차보증금", f"{C['보증금']}원"],
        ["(−) 미납 차임", f"{C['연체차임']}원"],
        ["(=) 소계 (관리비 공제 불인정 시 반환잔액)", f"{C['잔액_차임만']}원"],
        ["(−) 미납 관리비", f"{C['연체관리비']}원"],
        ["(=) 신청인 주장 반환잔액", f"{C['잔액']}원"],
    ], weights=[65, 35])
    b.h("5. 확인이 필요한 사항")
    b.items([
        f"가. 신청인 측이 파악한 미납 총액 1,200,000원과 위 항목별 합계 {C['연체합계']}원 사이에 "
        "40,000원의 차이가 있음. 월별 미납 내역으로 확정할 것(000).",
        f"나. 위 차이는 해지의 효력에 영향이 없음(차임만으로 {C['연체차임']}원 > 2기 {C['2기차임']}원).",
        "다. 다툼이 있는 공제 항목(연체료·원상복구비 등): 000",
    ])
    b.h("6. 신청인의 입장")
    b.p("신청인은 확정된 미납 차임과 정당한 관리비는 보증금에서 정산할 의사가 있으나, 그 밖의 공제 항목과 "
        "최종 반환액에 관하여 합의가 이루어지지 않았고 보증금은 전혀 반환되지 않았습니다.")
    b.sign(court=False)


def doc_affidavit(b):
    b.title("사 실 확 인 서")
    b.table([
        ["확인인", "내용"],
        ["성명", "000"],
        ["생년월일", "0000. 00. 00."],
        ["주소", "000"],
        ["연락처", "000-0000-0000"],
        ["신청인과의 관계", "000 (예: 이웃 거주자, 중개보조원, 관리인)"],
    ], weights=[30, 70])
    b.blank()
    b.p(f"본인은 {C['소재지']} 소재 {C['건물']} 2층의 현장 출입문에 「제203호」로 표시된 주거공간에 관하여 "
        "아래 사항을 직접 보고 들어 알고 있으므로 이를 확인합니다.")
    b.items([
        "1. 확인 기간 : 0000. 00. 00. ~ 0000. 00. 00.",
        "2. 현장 제203호의 표시와 구조 : 000",
        f"3. 신청인 {C['신청인']}이 {C['전입일']} 이후 위 공간 전체에 거주한 사실 : 000",
        "4. 현장 제203호가 공부상 제202호 전체에 해당한다는 사실(직접 아는 범위에서만) : 000",
        "5. 그 밖의 사항 : 000",
    ])
    b.p("위 내용은 본인이 직접 경험한 사실이며, 추측이나 전해 들은 내용은 기재하지 않았습니다.")
    b.sign(who="확인인", name="000", court=False)
    b.note("※ 확인인의 신분증 사본(또는 인감증명서)을 첨부하십시오.")


def doc_evidence(b):
    b.title("증 거 설 명 서")
    b.p("사건 : 2026카임000 주택임차권등기명령 신청")
    b.p(f"신청인 : {C['신청인']}    피신청인 : {C['피신청인']}")
    b.blank()
    b.table([
        ["호증", "서증명", "작성일", "입증취지"],
        ["갑 제1호증", "임대차계약서(확정일자 날인)", "2026. 00. 00.", f"보증금 {C['보증금']}원·차임 {C['차임']}원·관리비 {C['관리비']}원 약정, 확정일자 {C['전입일']}"],
        ["갑 제2호증", "보증금 이체내역", "2026. 00. 00.", "보증금 전액 지급"],
        ["갑 제3호증", "주민등록초본", "2026. 00. 00.", f"{C['전입일']} 전입 및 계속 거주"],
        ["갑 제4호증", "2026. 10. 26.자 문자메시지 출력본", "2026. 10. 26.", "차임 연체를 이유로 한 해지 통지"],
        ["갑 제5호증", "차임·관리비 납부·미납 내역", "2026. 00. 00.", f"차임 {C['연체차임']}원·관리비 {C['연체관리비']}원 미납, 차임과 관리비 구분"],
        ["갑 제6호증", "등기사항전부증명서(나동 202호)", "2026. 00. 00.", "목적물 표시, 전유면적, 소유자, 근저당·경매개시결정"],
        ["갑 제7호증", "건축물대장·건축물현황도", "2026. 00. 00.", "공부상 2층 4개 호실 구성, 제202호 위치·면적"],
        ["갑 제8호증", "현장 사진(출입문·내부)", "2026. 00. 00.", "현장 203호 표시와 점유공간 전체"],
        ["갑 제9호증", "별지 2 위치 표시 도면", "2026. 00. 00.", "공부상 제202호 전체 = 현장 제203호"],
        ["갑 제10호증", f"배당요구서·보정서({C['경매사건']})", "2026. 00. 00.", "배당요구 및 표시 정정 경위"],
        ["갑 제11호증", "다른 임차인 임차권등기 기록", "2026. 00. 00.", "건물 전체의 호실 표시 불일치 관행"],
        ["갑 제12호증", "사실확인서", "2026. 00. 00.", "현장 점유 및 호실 대응관계"],
        ["갑 제13호증", "내용증명 및 우체국 접수증", C["발송일"], "계약 종료 확인 및 보증금 반환 청구"],
    ], weights=[14, 28, 16, 42])
    b.note("※ 확보하지 못한 증거는 목록에서 삭제하고 번호를 다시 매기십시오.")
    b.sign(court=True)


def doc_checklist(b):
    b.title("제출서류 목록 및 체크리스트")
    b.h("1. 법정 첨부서류 (임차권등기명령 절차에 관한 규칙 제3조)")
    b.table([
        ["확인", "서류", "비고"],
        ["□", "건물 등기사항전부증명서(나동 제202호)", "최신 발급, 소유자 △△△ 확인"],
        ["□", "임대차계약증서 사본", "확정일자 날인면 포함"],
        ["□", "점유개시일·주민등록일 소명자료", "주민등록초본, 전입세대열람내역"],
        ["□", "주거용 사용 증명자료", "등기부상 용도가 주거가 아닐 때만"],
    ], weights=[8, 52, 40])
    b.h("2. 실무 첨부서류")
    b.table([
        ["확인", "서류", "비고"],
        ["□", "주택임차권등기명령 신청서", "원본 1 + 부본(피신청인 수)"],
        ["□", "신청이유 보충서", "202호 전체 = 현장 203호"],
        ["□", "별지 1 부동산 및 임차목적물의 표시", "등기부와 글자 대조"],
        ["□", "별지 2 위치 표시 도면", "「일부」·「사선」 표시 없는 도면"],
        ["□", "건축물대장·건축물현황도", "2층 4개 호실 확인"],
        ["□", "보증금 이체내역", "500만 원"],
        ["□", "정산 내역서·미납 내역", "차임과 관리비 분리"],
        ["□", "2026. 10. 26. 문자 캡처", "발신번호·날짜 보이게 전체"],
        ["□", "현장 사진", "출입문 203호, 내부, 계량기"],
        ["□", "다른 임차인 임차권등기 기록", "확보 시"],
        ["□", "경매 배당요구서·보정서 사본", C["경매사건"]],
        ["□", "사실확인서·증거설명서", "확보 시"],
        ["□", "내용증명 사본·우체국 접수증", f"{C['발송일']} 발송분 (갑 제13호증)"],
        ["□", "위임장·가족관계증명서", "가족이 대리 접수할 때만"],
    ], weights=[8, 52, 40])
    b.h("3. 비용 (접수 창구·위택스에서 최종 확정)")
    b.table([
        ["항목", "금액", "비고"],
        ["인지", "2,000원", "전자소송 1,800원"],
        ["송달료", "약 33,000원", "당사자 2명 × 3회분"],
        ["등록면허세(지방교육세 포함)", "7,200원", "위택스 납부"],
        ["등기신청수수료", "000원", "접수 시 확인"],
    ], weights=[40, 25, 35])
    b.p("※ 위 비용은 주택임대차보호법 제3조의3 제8항에 따라 임대인에게 청구할 수 있습니다.")
    b.h("4. 접수 전 반드시 확인")
    b.items([
        "① 등기부상 소유자가 여전히 △△△인지 (이미 제3자에게 이전되었으면 피신청인·전략 변경)",
        "② 등기부상 제202호 전유면적 = 실측·현황도 면적 (000㎡)",
        "③ 미납 총액 1,200,000원과 1,240,000원의 차이 40,000원 확정",
        f"④ 경매 {C['경매사건']} 배당요구는 철회하지 말 것, 보정서 202호 「전체」로 추가 정정",
        "⑤ 임차권등기 기입을 등기부로 확인한 뒤에 이사·전출할 것",
        "⑥ 2026. 4. 경매개시결정 이후의 임차권등기는 낙찰자에게 인수되지 않음 — 배당요구가 핵심",
        "⑦ 2017년 근저당권보다 선순위가 될 수 없음 — 배당 시 소액임차인 최우선변제(주임법 제8조) 주장",
        "⑧ 서류 안의 000을 모두 증빙으로 채울 것",
    ])


def doc_notice(b):
    b.title("내 용 증 명")
    b.table([
        ["구분", "내용"],
        ["수신인", f"{C['피신청인']} (주소: 000)"],
        ["발신인", f"{C['신청인']} (주소: 000, 연락처: 000-0000-0000)"],
        ["제목", "임대차계약 종료 확인 및 임차보증금 반환 청구"],
    ], weights=[20, 80])
    b.blank()
    b.p(f"1. (임대차계약) 본인은 2026. 00. 00. 귀하와 {C['소재지']} {C['건물']} 현장 표시 "
        f"제203호에 관하여 보증금 {C['보증금']}원, 월차임 {C['차임']}원, 관리비 월 {C['관리비']}원으로 "
        f"임대차계약을 체결하고 보증금 전액을 지급하였으며, {C['전입일']} 입주·전입신고를 하고 "
        "확정일자를 받았습니다.", hang=4)
    b.p("2. (계약의 종료) 귀하는 2026. 10. 26. 본인에게 「2026. 10. 30.까지 밀린 월세를 납부하지 "
        "않으면 계약을 해지한다」는 문자메시지를 보냈고, 본인은 그 기한까지 차임을 납부하지 "
        f"못하였습니다. 따라서 위 임대차계약은 귀하의 해지 통지에 따라 {C['종료일']} 종료되었음을 "
        "확인합니다.", hang=4)
    b.p(f"3. (보증금 반환 청구) 계약이 종료되었으므로, 임차보증금 {C['보증금']}원에서 본인이 납부하지 "
        f"못한 차임 {C['연체차임']}원({C['연체월수']}개월분)을 공제한 금 {C['잔액_차임만']}원의 반환을 "
        "청구합니다. 관리비 등 그 밖에 공제할 금액이 있다면 그 항목·금액·근거를 서면으로 알려 "
        "주시기 바랍니다. 확인 후 정당한 금액은 정산하겠습니다. 위 정산 내역의 회신과 보증금 반환을 "
        f"{C['회신기한']}까지 하여 주시기 바랍니다.", hang=4)
    b.p("4. (목적물의 표시와 인도) 본인이 임차·점유하는 목적물은 계약서와 현장 출입문에 제203호로 "
        "표시된 공간으로서, 건축물대장·등기부상 나동 제202호 전유부분 전체와 같은 곳입니다. 본인은 "
        "보증금을 반환받는 것과 동시에 위 목적물을 인도하겠습니다.", hang=4)
    b.p("5. (임차권등기명령 신청) 본인은 주거를 옮기더라도 이미 취득한 대항력과 우선변제권을 유지하기 "
        f"위하여, 주택임대차보호법 제3조의3에 따라 {C['작성일']} {C['법원']}에 임차권등기명령을 신청할 "
        "예정입니다. 이는 보증금 반환 협의와 별개로 권리를 보전하기 위한 절차이며, 그 신청·등기 비용은 "
        "같은 조 제8항에 따라 귀하에게 청구합니다. 보증금이 반환되면 임차권등기 말소에 협력하겠습니다.",
        hang=4)
    b.p(f"6. (경매절차) 본인은 {C['법원']} {C['경매사건']} 부동산경매절차에서 배당요구를 마쳤으며, "
        "그 절차에서도 임차인으로서의 권리를 행사하겠습니다.", hang=4)
    b.sign(who="발신인", date=C["발송일"], court=False)


def doc_poa(b):
    b.title("위  임  장")
    b.note("※ 신청인 본인이 아닌 가족 등이 접수할 때만 사용합니다. 비변호사 대리의 허용 여부와 "
           "소송대리허가신청 필요 여부는 접수 전 법원 민사신청과에 확인하십시오.")
    b.table([
        ["구분", "내용"],
        ["수임인 성명", "000"],
        ["주민등록번호", "000000-0000000"],
        ["주소", "000"],
        ["신청인과의 관계", "000 (가족관계증명서 첨부)"],
    ], weights=[30, 70])
    b.blank()
    b.p(f"위임인 {C['신청인']}은 위 사람을 대리인으로 정하고, {C['피신청인']}을 피신청인으로 하는 "
        f"{C['법원']} 주택임차권등기명령 신청사건에 관하여 신청서 제출, 보정, 송달료 납부·환급, "
        "결정문 및 등기필 관련 서류의 수령에 관한 모든 권한을 위임합니다.")
    b.sign(who="위임인", court=True)


DOCS = [
    ("01_주택임차권등기명령_신청서", doc_application),
    ("02_신청이유_보충서", doc_supplement),
    ("03_별지1_부동산및임차목적물의표시", doc_schedule1),
    ("04_별지2_위치표시도면", doc_schedule2),
    ("05_보증금_차임_정산내역서", doc_settlement),
    ("06_사실확인서", doc_affidavit),
    ("07_증거설명서", doc_evidence),
    ("08_제출서류목록_체크리스트", doc_checklist),
    ("09_내용증명", doc_notice),
    ("10_위임장_해당시", doc_poa),
]


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "서류")
    pdf_out = os.path.join(out, "PDF")
    word_out = os.path.join(out, "Word")
    os.makedirs(pdf_out, exist_ok=True)
    os.makedirs(word_out, exist_ok=True)
    combined_name = "00_임차권등기명령_신청서류_전체합본"
    for builder_cls, folder, ext in ((HwpBuilder, out, ".hwp"),
                                     (PdfBuilder, pdf_out, ".pdf"),
                                     (DocxBuilder, word_out, ".docx")):
        for name, fn in DOCS:
            b = builder_cls()
            fn(b)
            b.save(os.path.join(folder, name + ext))
        combined = builder_cls()
        for i, (_, fn) in enumerate(DOCS):
            if i:
                combined.page_break()
            fn(combined)
        combined.save(os.path.join(folder, combined_name + ext))
        if builder_cls is HwpBuilder:
            combined.save(os.path.join(folder, combined_name + ".hwpx"))
    print("saved to", out, pdf_out, word_out)


if __name__ == "__main__":
    main()
