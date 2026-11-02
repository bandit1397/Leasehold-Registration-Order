# -*- coding: utf-8 -*-
"""주택임차권등기명령 신청서류 일체 HWP·PDF 생성기.

전제: 공부상 ○○빌 나동 제202호 전유부분 전체 = 계약서·현장 표시 제203호 전체.
확인되지 않은 사항은 모두 000 으로 남긴다.

사용법: python build_hwp.py [출력폴더]
필요: pip install python-hwpx reportlab python-docx pywin32   (PDF 글꼴: Windows 바탕·맑은 고딕)
HWP: 한글 프로그램이 설치되어 있으면 Word 출력을 한글로 열어 저장한다(가장 정확).
     한글이 없으면 python-hwpx 로 직접 만든다(빈칸 노란 표시 없음).
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
    # --- 개인정보: 실제 값은 개인정보.json 에 둔다 (GitHub 에 올리지 않음). 여기는 가림 값 ---
    "신청인": "○○○",
    "신청인주민번호": "＿＿＿＿＿＿-＿＿＿＿＿＿＿",
    "신청인주소": "강원특별자치도 원주시 ○○면 ○○리 ○○○ (○○길 ○○, ○○빌) 나동 203호",  # 주민등록상 주소
    "연락처": "010-○○○○-○○○○",
    "피신청인": "△△△",
    # 아래 값이 "＿" 밑줄이면 인쇄 후 손으로 쓰는 칸, "" 이면 문장에서 그 부분을 뺀다
    "피신청인주소": "＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿",
    "피신청인연락처": "＿＿＿-＿＿＿＿-＿＿＿＿",
    "중개사": "",                         # 공인중개사 이름(상호). 없으면 "공인중개사의 중개로"
    "확인인성명": "＿＿＿＿＿＿",          # 06 사실확인서 (확인인이 직접 씀)
    "확인인생년월일": "＿＿＿＿. ＿＿. ＿＿.",
    "확인인주소": "＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿",
    "확인인연락처": "＿＿＿-＿＿＿＿-＿＿＿＿",
    "확인인관계": "＿＿＿＿＿＿＿＿＿＿",  # 예: 이웃 거주자, 관리인
    "수임인성명": "＿＿＿＿＿＿",          # 10 위임장
    "수임인주민번호": "＿＿＿＿＿＿-＿＿＿＿＿＿＿",
    "수임인주소": "＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿",
    "수임인관계": "＿＿＿＿＿＿ (가족관계증명서 첨부)",
    "법원": "춘천지방법원 원주지원",
    "경매사건": "2026타경○○",
    "경매채권자": "○○새마을금고",
    "소재지": "강원특별자치도 원주시 ○○면 ○○리 ○○○",
    "도로명": "강원특별자치도 원주시 ○○면 ○○길 ○○",
    "건물": "○○빌 나동",
    # --- 서류 확인 후 채울 사항 ---
    "사건번호": "2026카임        ",       # 접수 때 법원이 부여 → 빈칸으로 인쇄
    "경매개시등기일": "2026. 4.",         # 경매개시결정 기입등기일 (등기부 갑구, 예: 2026. 4. 15.)
    "연체시작": "2026. 7.",              # 연체 차임 첫 달 (예: 2026. 7.)
    "연체끝": "2026. 10.",                # 연체 차임 마지막 달
    "해지조항": "",                       # 계약서의 차임 연체 해지 조항 번호 (예: "7"). 없으면 조항 번호를 빼고 씀
    "배당요구접수일": "",                 # 없으면 괄호를 뺌
    "종전소명서제출일": "2026. 10. 30.",  # 「임차인 배당요구 관련 소명자료 제출서」를 경매계에 낸 날
    "경매계": "",                         # 예: "3" → 경매3계. 없으면 법원 이름만
    "표제부_1동": "＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿",                  # 별지 1: 등기사항전부증명서 표제부 그대로
    "전유구조": "＿＿＿＿＿＿＿＿＿＿",
    "토지표시": "＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿",
    "대지권비율": "＿＿＿＿＿분의 ＿＿＿＿",
    "등기신청수수료": "3,000",
    "배당요구종기": "2026. 6. 24.",
    "계약일": "2026. 2. 23.",
    "계약기간": "2026. 2. 23.부터 2027. 1. 22.까지",
    "종전소명서": "임차인 배당요구 관련 소명자료 제출서",
    "보증금": "5,000,000",
    "차임": "240,000",
    "관리비": "80,000",
    "연체월수": "4",
    "연체차임": "960,000",
    "2기차임": "480,000",
    "연체관리비": "320,000",
    "연체합계": "1,280,000",
    "잔액": "3,720,000",
    "잔액_차임만": "4,040,000",
    "전입일": "2026. 2. 23.",
    "종료일": "2026. 10. 31.",
    "작성일": "2026. 11. 3.",      # 임차권등기명령 신청일 (신청서·별지·정산서 등)
    "발송일": "2026. 11. 3.",      # 내용증명 발송일
    "회신기한": "2026. 11. 5.",    # 내용증명의 보증금 반환·동시 인도 기한
    "현황도발급일": "2026. 11. 2.",  # 원주시 발급 건축물대장·현황도
    "서증발급일": "2026. 11. 3.",    # 증거설명서의 발급·출력일 (계약서·문자·내용증명·대장 외)
}

# 개인정보.json (이 PC에만 있음) 이 있으면 위 가림 값을 실제 값으로 바꾼다
PRIVATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "개인정보.json")
if os.path.exists(PRIVATE):
    import json
    with open(PRIVATE, encoding="utf-8") as _f:
        C.update({k: v for k, v in json.load(_f).items() if not k.startswith("_")})

# 서류 구성: 사실확인서(증거)·위임장 포함 여부. main() 이 A안·B안을 차례로 만든다
V = {"사실확인서": True, "위임장": True}
VARIANTS = [
    ("A_사실확인서·위임장_포함", {"사실확인서": True, "위임장": True}),
    ("B_사실확인서·위임장_제외", {"사실확인서": False, "위임장": False}),
]

# 증거 순서. 사실확인서가 없으면 그 뒤 번호가 하나씩 당겨진다
EVIDENCE = ["계약서", "이체내역", "초본", "문자", "미납내역", "등기202", "대장", "사진", "별지2",
            "경매서류", "사실확인서", "내용증명", "등기203"]


def ev_keys():
    return [k for k in EVIDENCE if k != "사실확인서" or V["사실확인서"]]


def 갑(key):
    return f"갑 제{ev_keys().index(key) + 1}호증"


def image_size(path):
    from reportlab.lib.utils import ImageReader
    return ImageReader(path).getSize()


class BaseBuilder:
    """서류 본문 함수(doc_*)가 쓰는 공통 인터페이스. 출력 형식은 하위 클래스가 정한다."""

    def title(self, text):
        self._p(text, "title", "CENTER", after=14)

    def sub(self, text, align="CENTER"):
        self._p(text, "sub", align, before=6, after=6, keep=True)

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
        self.right(f"{who}  {name}")
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
        t.autofit = False
        # 한글은 칸 너비(tcW)가 아니라 표 격자(gridCol)를 따르므로 격자에도 비율을 넣는다
        for col, w in zip(t._tbl.tblGrid.gridCol_lst, weights):
            col.w = Mm(170 * w / sum(weights))
        # 셀 안쪽 여백도 명시한다 (없으면 한글이 위·아래 여백을 따로 넣어 행이 높아짐)
        mar = OxmlElement("w:tblCellMar")
        for side, val in (("top", 0), ("left", 108), ("bottom", 0), ("right", 108)):
            el = OxmlElement(f"w:{side}")
            el.set(qn("w:w"), str(val))
            el.set(qn("w:type"), "dxa")
            mar.append(el)
        t._tbl.tblPr.append(mar)
        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                cell = t.cell(r, c)
                cell.width = Mm(170 * weights[c] / sum(weights))
                for i, line in enumerate(val.split("\n")):
                    para = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
                    para.paragraph_format.space_after = 0
                    para.paragraph_format.line_spacing = 1.0  # 명시하지 않으면 한글이 160%로 연다
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
        ["사건번호", C["사건번호"]],
        ["신청인(임차인)", f"성명: {C['신청인']} / 주민등록번호: {C['신청인주민번호']}\n"
                       f"주소: {C['신청인주소']}\n연락처: {C['연락처']}"],
        ["피신청인(임대인)", f"성명: {C['피신청인']}\n주소: {C['피신청인주소']}\n"
                         f"연락처: {C['피신청인연락처']}"],
    ], weights=[25, 75])

    b.sub("신 청 취 지")
    # 신청취지는 결정문·등기부에 그대로 옮겨지므로 등기 사항만 적는다 (203호·관리비·종료일은 신청이유에)
    b.p("별지 1 목록 기재 건물에 관하여 아래와 같은 주택임차권등기를 명한다.")
    b.center("아    래")
    b.table([
        ["항목", "기재 내용"],
        ["1. 임대차계약일자", C["계약일"]],
        ["2. 임차보증금액 / 차임", f"보증금 금 {C['보증금']}원 / 차임 월 금 {C['차임']}원"],
        ["3. 주민등록일자", C["전입일"]],
        ["4. 점유개시일자", C["전입일"]],
        ["5. 확정일자", C["전입일"]],
        ["6. 임차범위", "나동 제202호 전유부분 전부 (전유면적 12.37㎡)"],
    ], weights=[30, 70])
    b.p("신청비용은 피신청인이 부담한다.")
    b.p("라는 결정을 구합니다.")

    b.sub("신 청 이 유")
    b.h("1. 임대차계약의 체결")
    b.p(f"신청인은 {C['계약일']} 공인중개사{' ' + C['중개사'] if C['중개사'] else ''}의 중개로 피신청인과 별지 1 기재 건물에 관하여 "
        f"임차보증금 금 {C['보증금']}원, 월차임 금 {C['차임']}원(관리비 월 금 {C['관리비']}원 별도), "
        f"임대차기간 {C['계약기간']}로 하는 임대차계약을 체결하고, "
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
        f"목적물에 관한 경매개시결정 기입등기일({C['경매개시등기일']}) 이전에 갖추어졌습니다. 신청인의 "
        "주민등록상 주소와 확정일자 받은 계약서에는 현장 표시에 따라 「나동 203호」로 기재되어 "
        "있으나, 이는 공부상 나동 제202호 전유부분 전체를 가리킵니다.", hang=4)
    b.h("4. 임대차의 종료")
    b.p(f"신청인은 {C['연체시작']}분부터 {C['연체끝']}분까지 {C['연체월수']}개월분 차임 합계 금 {C['연체차임']}원을 "
        f"연체하였고, 이는 2기의 차임액 금 {C['2기차임']}원을 초과합니다. 피신청인은 2026. 10. 26. "
        "신청인에게 「2026. 10. 30.까지 밀린 월세를 납부하지 않으면 계약을 해지한다」는 취지의 "
        "문자메시지를 보냈고, 신청인은 위 기한까지 이를 납부하지 못하였습니다. 따라서 위 임대차계약은 "
        f"민법 제640조 및 임대차계약서{' 제' + C['해지조항'] + '조' if C['해지조항'] else ''}에 따라 {C['종료일']} 종료되었습니다.", hang=4)
    b.h("5. 임차보증금의 미반환")
    b.p(f"임대차가 종료되었음에도 신청인은 임차보증금 금 {C['보증금']}원을 전혀 반환받지 못하였습니다. "
        f"미납 차임 금 {C['연체차임']}원과 미납 관리비 금 {C['연체관리비']}원(합계 금 {C['연체합계']}원)을 "
        f"모두 공제하더라도 반환받지 못한 보증금은 금 {C['잔액']}원이고, 관리비 공제가 인정되지 않을 "
        f"경우에는 금 {C['잔액_차임만']}원입니다. "
        f"신청인은 {C['발송일']} 피신청인에게 내용증명우편으로 보증금 반환을 청구하였습니다"
        f"({갑('내용증명')}).", hang=4)
    b.h("6. 경매절차와의 관계")
    b.p(f"위 건물에 관하여 {C['법원']} {C['경매사건']} 부동산경매절차가 진행 중이며, 신청인은 배당요구 "
        f"종기({C['배당요구종기']}) 내에 배당요구를 마쳤습니다{'(배당요구 접수일 ' + C['배당요구접수일'] + ')' if C['배당요구접수일'] else ''}. "
        "본 신청은 위 배당요구를 그대로 유지하면서, "
        "신청인이 주거를 옮기더라도 이미 취득한 대항력과 우선변제권을 유지하기 위한 것입니다.", hang=4)
    b.h("7. 결론")
    b.p("이에 주택임대차보호법 제3조의3에 따라 신청취지와 같은 결정을 구합니다. 피신청인이 위 건물을 "
        "제3자에게 매도할 가능성이 있으므로, 신속한 결정과 등기 촉탁을 바랍니다.", hang=4)

    b.sub("소 명 방 법 및 첨 부 서 류", align="LEFT")
    b.items([
        "1. 건물등기사항전부증명서(나동 제202호)          1통",
        "2. 임대차계약증서 사본(확정일자 날인면 포함)      1통",
        "3. 주민등록초본(주소변동사항 포함)                1통",
        "4. 건축물대장 및 건축물현황도                      각 1통",
        "5. 신청이유 보충서                                 1통",
        "6. 별지 1 부동산 및 임차목적물의 표시            1통",
        "7. 별지 2 임차목적물 위치 표시 도면              1통",
        f"8. 증거설명서 및 갑 제1호증 내지 갑 제{len(ev_keys())}호증     각 1통",
        "   (위 1~4호 서류는 갑 제6·1·3·7호증과 같음)",
        "9. 송달료·인지·등록면허세·등기신청수수료 납부서  각 1통",
    ])
    b.sign()


def doc_supplement(b):
    b.title("신청이유 보충서")
    b.p(f"사    건 : {C['사건번호']} 주택임차권등기명령 신청")
    b.p(f"신 청 인 : {C['신청인']}")
    b.p(f"피신청인 : {C['피신청인']}")
    b.blank()
    b.p("위 사건에 관하여 신청인은 임차목적물의 동일성과 임대차 종료 사유를 아래와 같이 보충합니다.")

    b.h("1. 계약 체결 경위")
    b.p(f"신청인은 {C['계약일']} 공인중개사{' ' + C['중개사'] if C['중개사'] else ''}의 중개로 피신청인 소유 {C['건물']} 2층을 임차하였습니다. "
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
        f"바. 위 사실은 원주시 발급 건축물현황도(발급일 {C['현황도발급일']})와 집합건축물대장 전유부상 제202호 전유면적 "
        "12.37㎡의 대조로 확인됩니다.",
        "사. 이와 같은 공부상 호실과 현장 호실의 불일치는 신청인의 호실에만 있는 것이 아니라, 현재 "
        "2층에서 주거로 사용되는 현장 제202호부터 제207호까지 모든 호실에 공통으로 나타납니다. 건물 "
        "전체에 현장 호실 번호가 공부와 다르게 붙어 있기 때문이며, 신청인의 계약서·주민등록에 적힌 "
        "「제203호」도 같은 원인에서 비롯된 것입니다.",
    ])

    b.h("3. 계속 점유")
    b.p(f"신청인은 계약 당시 위 불일치를 알지 못하였고, {C['전입일']}부터 현재까지 현장 제203호, "
        "즉 공부상 제202호 전유부분 전체를 원룸(주방 포함)·현관·욕실·발코니를 포함하여 단독으로 계속 "
        "점유하고 있습니다.", hang=4)

    b.h("4. 차임 연체액이 2기 차임액에 달한 사실")
    b.items([
        f"가. 월차임은 금 {C['차임']}원이고, 관리비 월 금 {C['관리비']}원은 차임과 별도로 약정되었습니다.",
        f"나. 연체 차임은 {C['연체시작']}분부터 {C['연체끝']}분까지 {C['연체월수']}개월분 합계 금 {C['연체차임']}원으로, "
        f"2기 차임액 금 {C['2기차임']}원을 초과합니다. 위 금액에는 관리비·연체료가 포함되어 있지 않습니다.",
        "다. 피신청인은 2026. 10. 26. 문자메시지로 2026. 10. 30.까지 미납 시 해지한다고 통지하였고, "
        f"기한이 지나 임대차는 {C['종료일']} 종료되었습니다(갑 제4호증).",
    ])

    b.h("5. 경매절차에서의 표시 정정 경위")
    b.p(f"신청인은 {C['법원']} {C['경매사건']} 경매절차에서 배당요구 종기({C['배당요구종기']}) 전에 "
        f"배당요구를 마친 뒤, 「{C['종전소명서']}」와 점유 부분 표시 도면을 제출하였습니다{'(제출일 ' + C['종전소명서제출일'] + ')' if C['종전소명서제출일'] else ''}. "
        "위 소명자료 제출서에는 목적물을 「공부상 나동 202호 중 일부」로 표시하였습니다. 이는 점유 부분이 "
        "공부상 202호에 해당한다는 점까지 확인한 상태에서, 그 범위가 202호의 일부인지 전체인지를 "
        "건축물현황도의 호실 표기와 정밀하게 대조하지 못한 채 작성한 데 따른 것입니다. "
        f"그 후 {C['현황도발급일']} 집합건축물대장(전유부)과 건축물현황도를 새로 발급받아 평면도(별지 2)의 "
        "「202호~502호」 표기, 공부상 경계 및 전유면적(12.37㎡)과 대조한 결과, 신청인이 점유하는 "
        "공간은 나뉜 적이 없는 공부상 제202호 전유부분 전체와 일치함을 확인하였습니다. 이에 신청인은 "
        "목적물을 공부상 제202호 전유부분 전체로 정정하여 소명하며, 경매법원에도 같은 취지의 "
        f"「임차목적물 표시 정정 소명서」를 제출하였습니다(제출일 {C['작성일']}).", hang=4)

    b.h("6. 소명의 취지")
    b.p("신청인이 임차한 목적물은 공부상 나동 제202호 전유부분 전체이며, 계약서 및 현장의 「제203호」는 "
        "같은 공간을 가리키는 현장 호실 표시에 불과합니다. 따라서 임차권등기는 공부상 나동 제202호 "
        "전유부분 전체에 관하여 이루어져야 함을 소명합니다. 한편 공부상 제203호에 관하여는 이미 다른 "
        f"임차인의 임차권등기가 마쳐져 있으나({갑('등기203')}), 신청인의 등기는 공부상 제202호에 관한 것이므로 "
        "그 등기와 목적물이 겹치지 않습니다.", hang=4)
    b.sign()


def doc_schedule1(b):
    b.title("별지 1  부동산 및 임차목적물의 표시")
    b.h("1. 1동의 건물의 표시")
    b.items([
        f"소재지 : {C['소재지']}",
        f"도로명주소 : {C['도로명']}",
        f"건물명칭 및 번호 : {C['건물']}",
        f"구조·지붕·층수·면적 : {C['표제부_1동']}",
    ])
    b.h("2. 전유부분의 건물의 표시")
    b.items([
        "건물번호 : 제2층 제202호",
        f"구조 : {C['전유구조']}",
        "면적 : 12.37㎡",
    ])
    b.h("3. 대지권의 표시")
    b.items([
        f"토지의 표시 : {C['토지표시']}",
        "대지권의 종류 : 소유권대지권",
        f"대지권의 비율 : {C['대지권비율']}",
    ])
    b.h("4. 임차목적물(임차범위)")
    b.p("위 제2층 제202호 전유부분 전부 (전유면적 12.37㎡)", indent=4)
    b.h("5. 비고")
    b.p("임대차계약서 및 현장 출입문에는 「제203호」로 표시되어 있으나, 이는 공부상 나동 제202호 "
        "전유부분 전체와 동일한 공간을 가리키는 현장 호실 표시임. 끝.", indent=4)
    b.blank()


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
         "공부상 2개 호실을 현장에서 4개로 나누어 사용 (사진 2 주황색 선). "
         f"공부상 제203호에는 다른 임차인의 임차권등기가 마쳐져 있음 ({갑('등기203')})"],
    ], weights=[24, 24, 52])
    b.note("※ 현재 2층에서 주거로 사용되는 현장 제202호부터 제207호까지 모든 호실이 공부상 호실과 번호가 "
           "다르며, 신청인의 계약서·주민등록상 「제203호」도 같은 원인에서 비롯된 현장 호실 표시임.")
    b.h("2. 임차목적물의 표시")
    b.table([
        ["구분", "내용"],
        ["임차목적물", "사진 1·2의 빨간색 실선으로 둘러싼 부분 전체 (두 사진의 경계가 같음)"],
        ["도면상 호실 표기", "원룸형 \"A\" (202호~502호) → 2층은 공부상 제202호"],
        ["현장·계약서 표시", "제203호"],
        ["위치", "2층 도면 우측 상단 끝, 기준선 X3~X4 사이"],
        ["구성", "확장형 발코니, 원룸(주방 포함), 현관, 욕실 (아래쪽 창고는 제외)"],
        ["도면 치수(벽 중심선 기준)", "원룸(주방 포함) 3,600mm × 2,500mm, 현관·욕실 부분 깊이 2,000mm "
                                    "(욕실 1,500mm × 2,000mm)\n"
                                    "※ 발코니(깊이 1,200mm)는 전유면적에 포함되지 않음. 벽 중심선 치수이므로 "
                                    "벽체 안쪽 기준인 전유면적보다 다소 크게 계산됨"],
        ["전유면적(집합건축물대장)", "12.37㎡"],
        ["파란색 실선", "다른 호실 — 임차목적물 아님"],
        ["도면 출처", f"원주시 발급 건축물현황도 (발급일 {C['현황도발급일']})"],
    ], weights=[28, 72])
    b.blank()
    b.p("본 도면 사진 1·2에서 빨간색 실선으로 표시한 부분 전체는 신청인 ○○○이 2026. 2. 23. 인도받아 "
        "점유하고 있는 현장 출입문 표시 제203호로서, 도면에 원룸형 \"A\"(202호~502호)로 표기된 "
        "공부상 ○○빌 나동 제202호 전유부분 전체에 해당하는 임차목적물임을 표시한다.")
    b.blank()
    b.items([
        f"작성일 : {C['작성일']}",
        f"작성자 : 신청인 {C['신청인']}",
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
        ["월 차임", f"{C['차임']}원", f"임대차계약서{' 제' + C['해지조항'] + '조' if C['해지조항'] else ''}"],
        ["미납 기간", f"{C['연체월수']}개월", f"{C['연체시작']}분 ~ {C['연체끝']}분"],
        ["미납 차임 합계", f"{C['연체차임']}원", f"{C['차임']} × {C['연체월수']}"],
        ["2기 차임액", f"{C['2기차임']}원", f"{C['차임']} × 2 → 미납 차임 {C['연체차임']}원이 이를 초과"],
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
    b.h("5. 신청인의 입장")
    b.p("신청인은 확정된 미납 차임과 정당한 관리비는 보증금에서 정산할 의사가 있으나, 그 밖의 공제 항목과 "
        "최종 반환액에 관하여 합의가 이루어지지 않았고 보증금은 전혀 반환되지 않았습니다.")
    b.sign(court=False)


def doc_affidavit(b):
    b.title("사 실 확 인 서")
    b.table([
        ["확인인", "내용"],
        ["성명", C["확인인성명"]],
        ["생년월일", C["확인인생년월일"]],
        ["주소", C["확인인주소"]],
        ["연락처", C["확인인연락처"]],
        ["신청인과의 관계", C["확인인관계"]],
    ], weights=[30, 70])
    b.blank()
    b.p(f"본인은 {C['소재지']} 소재 {C['건물']} 2층의 현장 출입문에 「제203호」로 표시된 주거공간에 관하여 "
        "아래 사항을 직접 보고 들어 알고 있으므로 이를 확인합니다.")
    b.items([
        "1. 확인 기간 : ＿＿＿＿. ＿＿. ＿＿. ~ ＿＿＿＿. ＿＿. ＿＿.",
        "2. 현장 제203호의 표시와 구조 : ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿",
        f"3. 신청인 {C['신청인']}이 {C['전입일']} 이후 위 공간 전체에 거주한 사실 : ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿",
        "4. 현장 제203호가 공부상 제202호 전체에 해당한다는 사실(직접 아는 범위에서만) : ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿",
        "5. 그 밖의 사항 : ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿",
    ])
    b.p("위 내용은 본인이 직접 경험한 사실이며, 추측이나 전해 들은 내용은 기재하지 않았습니다.")
    b.sign(who="확인인", name=C["확인인성명"], court=False)


def doc_evidence(b):
    b.title("증 거 설 명 서")
    b.p(f"사건 : {C['사건번호']} 주택임차권등기명령 신청")
    b.p(f"신청인 : {C['신청인']}    피신청인 : {C['피신청인']}")
    b.blank()
    d = C["서증발급일"]
    rows = {
        "계약서": ("임대차계약서(확정일자 날인)", C["계약일"],
                f"보증금 {C['보증금']}원·차임 {C['차임']}원·관리비 {C['관리비']}원 약정, 확정일자 {C['전입일']}"),
        "이체내역": ("보증금 이체내역", d, "보증금 전액 지급"),
        "초본": ("주민등록초본", d, f"{C['전입일']} 전입 및 계속 거주"),
        "문자": ("2026. 10. 26.자 문자메시지 출력본", "2026. 10. 26.", "차임 연체를 이유로 한 해지 통지"),
        "미납내역": ("차임·관리비 납부·미납 내역", d,
                 f"차임 {C['연체차임']}원·관리비 {C['연체관리비']}원 미납, 차임과 관리비 구분"),
        "등기202": ("등기사항전부증명서(나동 202호)", d, "목적물 표시, 전유면적, 소유자, 근저당·경매개시결정"),
        "대장": ("건축물대장·건축물현황도", C["현황도발급일"], "공부상 2층 4개 호실 구성, 제202호 위치·면적"),
        "사진": ("현장 사진(출입문·내부)", d, "현장 203호 표시와 점유공간 전체"),
        "별지2": ("별지 2 위치 표시 도면", d, "공부상 제202호 전체 = 현장 제203호"),
        # 경매 서류는 제출 시기가 달라 가지번호로 나눈다 (종전 소명서 → 이번 정정 소명서)
        "경매서류": [
            (f"배당요구서({C['경매사건']})", f"{C['배당요구종기']} 이전", "배당요구 종기 내 배당요구"),
            (C["종전소명서"], C["종전소명서제출일"], "목적물을 「202호 중 일부」로 표시한 경위"),
            ("임차목적물 표시 정정 소명서", C["작성일"], "목적물을 「202호 전유부분 전체」로 정정"),
        ],
        "사실확인서": ("사실확인서", d, "현장 점유 및 호실 대응관계"),
        "내용증명": ("내용증명 및 우체국 접수증", C["발송일"], "계약 종료 확인 및 보증금 반환 청구"),
        "등기203": ("등기사항전부증명서(나동 203호)", d,
                  "공부상 제203호의 다른 임차인 임차권등기 — 신청인 등기(제202호)와 겹치지 않음"),
    }
    b.table([["호증", "서증명", "작성일", "입증취지"]]
            + [row for k in ev_keys() for row in (
                [[f"{갑(k)}의{i}", *r] for i, r in enumerate(rows[k], 1)] if isinstance(rows[k], list)
                else [[갑(k), *rows[k]]])],
            weights=[17, 27, 15, 41])
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
        ["□", "경매 배당요구서·소명자료 제출서·정정 소명서 사본", f"{C['경매사건']} ({갑('경매서류')}의1~3)"],
        ["□", "사실확인서·증거설명서" if V["사실확인서"] else "증거설명서",
         f"사실확인서 = {갑('사실확인서')}, 확인인 신분증 사본 첨부" if V["사실확인서"] else ""],
        ["□", "내용증명 사본·우체국 접수증", f"{C['발송일']} 발송분 ({갑('내용증명')})"],
        ["□", "등기사항전부증명서(나동 제203호)", f"을구 임차권등기 보이게, 다른 임차인 주민등록번호 뒷자리 가림 ({갑('등기203')})"],
    ] + ([["□", "위임장·가족관계증명서", "가족이 대리 접수할 때만"]] if V["위임장"] else []),
        weights=[8, 52, 40])
    b.h("3. 비용 (접수 창구·위택스에서 최종 확정)")
    b.table([
        ["항목", "금액", "비고"],
        ["인지", "2,000원", "전자소송 1,800원"],
        ["송달료", "약 33,000원", "당사자 2명 × 3회분"],
        ["등록면허세(지방교육세 포함)", "7,200원", "위택스 납부"],
        ["등기신청수수료", f"{C['등기신청수수료']}원", "접수 시 확인"],
    ], weights=[40, 25, 35])
    b.p("※ 위 비용은 주택임대차보호법 제3조의3 제8항에 따라 임대인에게 청구할 수 있습니다.")
    b.h("4. 접수 전 반드시 확인")
    b.items([
        "① 등기부상 소유자가 여전히 △△△인지 (이미 제3자에게 이전되었으면 피신청인·전략 변경)",
        "② 등기부상 제202호 전유면적이 집합건축물대장(12.37㎡)과 같은지 대조",
        f"③ 월 차임 {C['차임']}원·관리비 {C['관리비']}원(월 합계 320,000원)과 미납 4개월분이 계약서·통장 내역과 맞는지 확인",
        f"④ 경매 {C['경매사건']} 배당요구는 철회하지 말 것, 정정 소명서(서류 11)로 202호 「전체」 정정",
        "⑤ 임차권등기 기입을 등기부로 확인한 뒤에 이사·전출할 것",
        "⑥ 2026. 4. 경매개시결정 이후의 임차권등기는 낙찰자에게 인수되지 않음 — 배당요구가 핵심",
        "⑦ 2017년 근저당권보다 선순위가 될 수 없음 — 배당 시 소액임차인 최우선변제(주임법 제8조) 주장",
        "⑧ 밑줄 칸(＿＿)을 모두 채울 것 — 주민등록번호, 임대인 주소·연락처, 별지 1 표제부(등기부 그대로)",
    ])


def doc_notice(b):
    b.title("내 용 증 명")
    b.table([
        ["수신인", f"{C['피신청인']} (주소: {C['피신청인주소']})"],
        ["발신인", f"{C['신청인']} (주소: {C['신청인주소']}, 연락처: {C['연락처']})"],
        ["제목", "임대차계약 종료 확인 및 임차보증금 반환 청구"],
    ], weights=[20, 80], header=False)
    b.p(f"1. (임대차계약) 본인은 {C['계약일']} 귀하와 {C['소재지']} {C['건물']} 현장 표시 "
        f"제203호에 관하여 보증금 {C['보증금']}원, 월차임 {C['차임']}원, 관리비 월 {C['관리비']}원으로 "
        f"임대차계약을 체결하고 보증금 전액을 지급하였으며, {C['전입일']} 입주·전입신고를 하고 "
        "확정일자를 받았습니다.", hang=4)
    b.p("2. (계약의 종료) 귀하는 2026. 10. 26. 본인에게 「2026. 10. 30.까지 밀린 월세를 납부하지 "
        "않으면 계약을 해지한다」는 문자메시지를 보냈고, 본인은 그 기한까지 차임을 납부하지 "
        f"못하였습니다. 따라서 위 임대차계약은 귀하의 해지 통지에 따라 {C['종료일']} 종료되었음을 "
        "확인합니다.", hang=4)
    b.p(f"3. (보증금 반환 청구) 계약이 종료되어 귀하의 보증금 반환 의무는 이미 이행기가 되었습니다. "
        f"임차보증금 {C['보증금']}원에서 본인이 납부하지 못한 차임 {C['연체차임']}원"
        f"({C['연체월수']}개월분)을 공제한 금 {C['잔액_차임만']}원(관리비 미납분이 확인되면 이를 추가로 "
        f"공제한 금액)을 늦어도 {C['회신기한']}까지 반환하여 주시기 바랍니다. 위 날짜는 반환을 요청하는 "
        "날짜일 뿐, 반환 시기를 늦추어 드리거나 본인의 권리 행사를 미루겠다는 뜻이 아닙니다.", hang=4)
    b.p("4. (목적물의 표시와 인도) 본인이 임차·점유하는 목적물은 계약서와 현장 출입문에 제203호로 "
        "표시된 공간으로서, 건축물대장·등기부상 나동 제202호 전유부분 전체와 같은 곳입니다. 귀하가 "
        f"{C['회신기한']}까지 위 보증금을 반환하면, 본인은 반환받는 것과 동시에 열쇠를 넘기는 등 위 "
        "목적물을 인도하겠습니다.", hang=4)
    # 임차권등기명령 신청 사실은 일부러 적지 않는다 (사전 고지 의무 없음, 매도 서두름 방지)
    b.p("5. (권리 보전) 본인은 위 반환 요청과 관계없이 보증금 반환을 확보하기 위하여 필요한 법적 절차를 "
        "진행할 수 있으며, 그 비용은 귀하에게 청구하겠습니다.", hang=4)
    b.p(f"6. (경매절차) 본인은 {C['법원']} {C['경매사건']} 부동산경매절차에서 배당요구를 마쳤으며, "
        "그 절차에서도 임차인으로서의 권리를 행사하겠습니다.", hang=4)
    b.right(f"{C['발송일']}    발신인  {C['신청인']}")  # 한 쪽에 맞추려 한 줄로


def doc_poa(b):
    b.title("위  임  장")
    b.table([
        ["구분", "내용"],
        ["수임인 성명", C["수임인성명"]],
        ["주민등록번호", C["수임인주민번호"]],
        ["주소", C["수임인주소"]],
        ["신청인과의 관계", C["수임인관계"]],
    ], weights=[30, 70])
    b.blank()
    b.p(f"위임인 {C['신청인']}은 위 사람을 대리인으로 정하고, 피신청인 {C['피신청인']}에 대한 "
        f"{C['법원']} 주택임차권등기명령 신청사건에 관하여 신청서 제출, 보정, 송달료 납부·환급, "
        "결정문 및 등기필 관련 서류의 수령에 관한 모든 권한을 위임합니다.")
    b.sign(who="위임인", court=True)


def doc_auction_correction(b):
    """경매법원(경매계)에 내는 정정 소명서. 경매기록은 소유자도 열람할 수 있으므로
    계약 종료·임차권등기 이야기는 쓰지 않는다."""
    b.title("임차목적물 표시 정정 소명서")
    b.center("(임차인 배당요구 관련)")
    b.table([
        ["사    건", f"{C['경매사건']} 부동산임의경매"],
        ["채 권 자", C["경매채권자"]],
        ["채 무 자\n(소 유 자)", C["피신청인"]],
        ["제 출 인\n(임 차 인)", f"성명: {C['신청인']} (연락처: {C['연락처']})\n주소: {C['신청인주소']}\n"
                            "(현장 표시 나동 203호 / 공부상 나동 202호 전유부분 전체)"],
    ], weights=[20, 80], header=False)
    b.p(f"위 사건에 관하여 제출인(임차인)은 「{C['종전소명서']}」({'제출일 ' + C['종전소명서제출일'] + ', ' if C['종전소명서제출일'] else ''}이하 "
        "「종전 소명서」)에서 임차목적물을 「공부상 나동 202호 중 일부」로 표시하였으나, 이를 아래와 "
        "같이 정정하고 소명자료를 제출합니다.")

    b.h("1. 정정 내용")
    b.table([
        ["구분", "종전 소명서", "정정"],
        ["임차목적물", "공부상 나동 202호 중 일부\n(도면 표시 부분)", "공부상 나동 202호 전유부분 전체"],
        ["현장 호실 표시", "나동 203호", "나동 203호 (변동 없음)"],
        ["점유 부분 표시", "소명자료 5(건축물현황도) 위에\n「일부」로 표시", "이 소명서 별지 사진의\n빨간색 표시(전체)로 대체"],
    ], weights=[22, 39, 39])

    b.h("2. 정정 사유")
    b.items([
        "가. 종전 소명서는 제출인의 점유 부분이 공부상 202호에 해당한다는 점까지 확인한 상태에서, 그 "
        "범위가 202호의 일부인지 전체인지를 건축물현황도의 호실 표기와 정밀하게 대조하지 못한 채 작성한 "
        f"것입니다. 제출인은 종전 소명서 제출 후 {C['현황도발급일']} 집합건축물대장(전유부)과 "
        "건축물현황도를 새로 발급받아 대조하였고, 그 결과 점유 부분이 공부상 202호 전유부분 전체"
        "(전유면적 12.37㎡)와 일치함을 확인하였습니다.",
        "나. 2층 평면도에는 공부상 4개 호실이 201호~501호, 202호~502호, 203호~503호, 204호~504호로 "
        "표기되어 있고, 제출인이 점유하는 공간은 원룸형 \"A\"(202호~502호)로 표기되어 있습니다"
        "(별지 사진 1).",
        "다. 현장에서는 공부상 201호가 201호·202호로, 공부상 203호·204호가 204호~207호로 나뉘어 "
        "사용되고 있으나, 공부상 202호는 나뉘지 않은 채 출입문 표시만 203호로 바뀌어 사용되고 "
        "있습니다(별지 사진 2).",
        "라. 이와 같은 공부상 호실과 현장 호실의 불일치는 제출인의 호실에만 있는 것이 아니라, 현재 "
        "2층에서 주거로 사용되는 현장 202호부터 207호까지 모든 호실에 공통으로 나타납니다. 건물 "
        "전체에 현장 호실 번호가 공부와 다르게 붙어 있기 때문이며, 제출인의 계약서·주민등록에 적힌 "
        "「203호」도 같은 원인에서 비롯된 것입니다.",
        "마. 따라서 제출인이 점유하는 현장 203호는 공부상 202호의 일부가 아니라, 원룸(주방 포함)·현관·"
        "욕실로 이루어진 공부상 202호 전유부분 전체(전유면적 12.37㎡)이며, 제출인이 이를 이에 딸린 "
        "확장형 발코니와 함께 단독으로 점유하고 있습니다.",
    ])

    b.h("3. 변동 없는 사항")
    b.p(f"계약일({C['계약일']}), 임대차보증금(금 {C['보증금']}원), "
        f"인도·전입신고·확정일자({C['전입일']}), 배당요구(배당요구 종기 {C['배당요구종기']} 이전) 등 "
        "종전 소명서의 나머지 기재 사항은 변동이 없습니다.", hang=4)

    b.h("4. 요청사항")
    b.items([
        "가. 제출인의 배당요구를 공부상 나동 202호 전유부분 전체에 관한 임차인의 배당요구로 보아 "
        "주시기 바랍니다.",
        "나. 주택임대차보호법 제3조의2에 따른 우선변제권 및 같은 법 제8조에 따른 소액임차인 "
        "최우선변제권 해당 여부를 판단하실 때 위 정정 내용을 반영하여 주시기 바랍니다.",
        "다. 종전 소명서의 소명자료 5(건축물현황도)는 그대로 원용하되, 그 위에 「일부」로 표시한 "
        "점유 부분 표시는 이 소명서의 별지로 대체하여 주시기 바랍니다.",
    ])

    b.h("소명자료")
    b.items([
        "1. 2층 평면도 비교 사진 (공부상 호실 표시 / 현재 현장 호실 표시) — 별지",
        f"2. 집합건축물대장(전유부, 나동 제202호) 및 건축물현황도 ({C['현황도발급일']} 발급)   각 1통",
    ])
    b.p("※ 종전 소명서의 소명자료 5(건축물현황도)·6(건축물대장·등기사항전부증명서)은 그대로 원용하며, "
        "위 2는 전유면적과 호실 경계를 확인하기 위하여 새로 발급받은 것입니다.", indent=4)
    b.blank()
    b.center(C["작성일"])
    b.right(f"위 제출인(임차인)  {C['신청인']}")
    b._p(f"{C['법원']} {'경매' + C['경매계'] + '계 ' if C['경매계'] else ''}귀중", "sub", "CENTER", before=12)

    b.page_break()
    b.title("별지  2층 평면도 비교 사진")
    b.caption("[사진 1]  공부상 호실 표시")
    b.picture(PLAN_IMAGES[0], width_mm=120)
    b.caption("[사진 2]  현재 현장 호실 표시")
    b.picture(PLAN_IMAGES[1], width_mm=120)
    b.small("빨간색 실선: 제출인 점유 부분(공부상 202호 전유부분 전체 = 현장 203호) / "
            "파란색 실선: 다른 호실 / 주황색 선: 현장에서 나누어 사용하는 경계. "
            "사진 1의 호실 번호는 도면에 인쇄된 호실 표기(201호~501호, 202호~502호, 203호~503호, "
            "204호~504호)에 따라 표시한 것입니다.")


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
    ("11_경매법원_임차목적물_표시정정_소명서", doc_auction_correction),
]


def active_docs():
    """현재 V(사실확인서·위임장 포함 여부)에 맞는 서류 목록."""
    skip = set()
    if not V["사실확인서"]:
        skip.add(doc_affidavit)
    if not V["위임장"]:
        skip.add(doc_poa)
    return [(name, fn) for name, fn in DOCS if fn not in skip]


def hangul_from_word(word_dir, out, combined_name):
    """설치된 한글 프로그램으로 Word 출력을 열어 한글 원본(.hwp)으로 저장한다.

    빈칸 노란 표시·표 열 너비·사진이 Word와 같게 옮겨진다. 한글은 Word의 1.3배 줄 간격을
    조금 넓게 그리므로 본문 줄 간격을 125%로 맞춰 쪽수를 Word와 비슷하게 한다.
    한글이 없거나 자동화가 안 되면 False 를 돌려준다.
    """
    try:
        import win32com.client
        hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
    except Exception as e:
        print("한글 자동화를 쓸 수 없음 → python-hwpx 로 만듦:", e)
        return False
    try:
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    except Exception:
        pass  # 보안 모듈이 없으면 한글이 파일 접근 허용 여부를 물어본다
    try:
        names = [name for name, _ in active_docs()] + [combined_name]
        failed = []
        for name in names:
            hwp.Open(os.path.join(word_dir, name + ".docx"), "", "forceopen:true")
            hwp.Run("SelectAll")
            ps = hwp.HParameterSet.HParaShape
            hwp.HAction.GetDefault("ParagraphShape", ps.HSet)
            ps.LineSpacingType, ps.LineSpacing = 0, 125
            hwp.HAction.Execute("ParagraphShape", ps.HSet)
            hwp.Run("Cancel")
            # 표를 글자처럼 취급해야 제목 앞쪽으로 표가 떠서 배치되지 않는다
            ctrl = hwp.HeadCtrl
            while ctrl:
                if ctrl.CtrlID == "tbl":
                    prop = ctrl.Properties
                    prop.SetItem("TreatAsChar", True)
                    ctrl.Properties = prop
                ctrl = ctrl.Next
            if not hwp.SaveAs(os.path.join(out, name + ".hwp"), "HWP", ""):
                failed.append(name)
            if name == combined_name:
                hwp.SaveAs(os.path.join(out, name + ".hwpx"), "HWPX", "")
            print(f"{name}.hwp  {hwp.PageCount}쪽")
            hwp.Clear(1)
    finally:
        hwp.Quit()
    if failed:
        # 한글 창에서 열려 있는 파일은 잠겨 있어 덮어쓰지 못한다
        sys.exit("저장 실패 (한글에서 열려 있으면 닫고 다시 실행): " + ", ".join(failed))
    return True


def build(out):
    pdf_out = os.path.join(out, "PDF")
    word_out = os.path.join(out, "Word")
    os.makedirs(pdf_out, exist_ok=True)
    os.makedirs(word_out, exist_ok=True)
    docs = active_docs()
    combined_name = "00_임차권등기명령_신청서류_전체합본"
    for builder_cls, folder, ext in ((PdfBuilder, pdf_out, ".pdf"),
                                     (DocxBuilder, word_out, ".docx"),
                                     (HwpBuilder, out, ".hwp")):
        if builder_cls is HwpBuilder and hangul_from_word(word_out, out, combined_name):
            continue
        for name, fn in docs:
            b = builder_cls()
            fn(b)
            b.save(os.path.join(folder, name + ext))
        combined = builder_cls()
        for i, (_, fn) in enumerate(docs):
            if i:
                combined.page_break()
            fn(combined)
        combined.save(os.path.join(folder, combined_name + ext))
        if builder_cls is HwpBuilder:
            combined.save(os.path.join(folder, combined_name + ".hwpx"))
    print("saved to", out)


def main():
    root = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "서류")
    for folder, variant in VARIANTS:
        V.update(variant)
        print(f"=== {folder} ===")
        build(os.path.join(root, folder))


if __name__ == "__main__":
    main()
