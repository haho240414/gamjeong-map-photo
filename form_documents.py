#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""한글(HWP) 감정평가 서식 `위치도(34-580)`·`건물개황도(34-581)`를 Word로 재현한다.

서식 치수는 원본 .hwp(A4 세로, 여백 위 2.0·아래 1.0·좌우 1.5cm, 표 폭 17.98cm)의 표 구조에서 읽은 값이다.
이미지(지도·개황도)는 호출하는 쪽이 만들어 넘긴다.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

import os
FONT = os.getenv("FORM_FONT") or "맑은 고딕"
TABLE_W = 17.98

# --- 로우레벨 ---------------------------------------------------------------


def set_font(run, size: float = 10, bold: bool = False, underline: bool = False) -> None:
    run.font.name = FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.underline = underline


def borders(cell, **edges) -> None:
    """edges: top/bottom/left/right = 'single'|'dotted'|'nil' (선 굵기 6=0.75pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcb = tcPr.find(qn("w:tcBorders"))
    if tcb is None:
        tcb = OxmlElement("w:tcBorders")
        tcPr.append(tcb)
    for edge in ("top", "left", "bottom", "right"):
        val = edges.get(edge, "nil")
        el = tcb.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            tcb.append(el)
        el.set(qn("w:val"), "nil" if val == "nil" else val)
        if val != "nil":
            el.set(qn("w:sz"), "6")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "000000")


def cell_margins(cell, left=0.1, right=0.05, top=0.03, bottom=0.03) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for k, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{k}")
        el.set(qn("w:w"), str(round(v / 2.54 * 1440)))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tcPr.append(mar)


def fixed_table(table, widths_cm: list[float]) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    tblPr = table._tbl.tblPr
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblPr.append(lay)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(round(sum(widths_cm) / 2.54 * 1440)))
    tblW.set(qn("w:type"), "dxa")
    grid = table._tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths_cm):
        gc.set(qn("w:w"), str(round(w / 2.54 * 1440)))
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            pass
    for ci, w in enumerate(widths_cm):
        for row in table.rows:
            row.cells[ci].width = Cm(w)


def row_height(row, cm: float, exact: bool = True) -> None:
    row.height = Cm(cm)
    row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY if exact else WD_ROW_HEIGHT_RULE.AT_LEAST


def para(cell_or_doc, text: str = "", size: float = 10, bold: bool = False, underline: bool = False,
         align=WD_ALIGN_PARAGRAPH.LEFT, first: bool = False, space_after: float = 0, space_before: float = 0):
    """셀의 첫 문단을 재사용(first=True)하거나 새 문단을 추가한다."""
    p = cell_or_doc.paragraphs[0] if first else cell_or_doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing = 1.0
    if text:
        set_font(p.add_run(text), size, bold, underline)
    return p


def new_document() -> Document:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin, sec.bottom_margin = Cm(2.0), Cm(1.0)
    sec.left_margin = sec.right_margin = Cm(1.5)
    sec.header_distance = sec.footer_distance = Cm(0.5)
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    st.font.size = Pt(10)
    return doc


def title(doc, text: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(10)
    set_font(p.add_run(text), 20, True, True)


def header_row(row, ids: tuple[str, str, str], merge_a, merge_b, merge_c) -> None:
    """감정번호 / 작성일 / 작성자 줄(서식처럼 밑줄). ids는 값(비워 두면 서식 그대로)."""
    nb = "\u00a0"
    labels = [f"감정번호: {ids[0]}", f"작성일: {ids[1]}", f"작성자: {ids[2]}"]
    defaults = [f"감정번호:{nb * 38}", f"작성일:{nb * 8}년{nb * 4}월{nb * 4}일{nb * 6}", f"작성자:{nb * 20}(인){nb * 2}"]
    for cell, label, default, i in zip((merge_a, merge_b, merge_c), labels, defaults, range(3)):
        text = label if ids[i] else default
        borders(cell)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
        cell_margins(cell, left=0.18)
        p = para(cell, "", first=True)
        set_font(p.add_run(text), 10, underline=True)


def footer_block(doc, reg_text: str, barcode: Path | None) -> None:
    """등록번호 줄 + 바코드는 쪽 아래(바닥글)에 둔다."""
    sec = doc.sections[0]
    sec.footer_distance = Cm(0.35)
    ft = sec.footer
    p = ft.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    pPr = p._p.get_or_add_pPr()
    pb = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "dotted")
    top.set(qn("w:sz"), "8")
    top.set(qn("w:space"), "1")
    top.set(qn("w:color"), "000000")
    pb.append(top)
    pPr.append(pb)
    set_font(p.add_run(reg_text), 8)
    if barcode and Path(barcode).exists():
        bp = ft.add_paragraph()
        bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        bp.paragraph_format.space_after = Pt(0)
        bp.add_run().add_picture(str(barcode), width=Cm(4.0), height=Cm(1.0))


# --- 위치도 (34-580) --------------------------------------------------------


def build_location_form(out: Path, overview_img: Path, detail_img: Path, barcode: Path | None = None,
                        ids: tuple[str, str, str] = ("", "", ""), symbol: str = "",
                        scale_top: str = "N.S") -> Path:
    doc = new_document()
    title(doc, "위   치   도")
    widths = [1.65, 2.37, 1.77, 6.20, 5.99]
    tbl = doc.add_table(rows=4, cols=5)
    fixed_table(tbl, widths)
    heights = [1.08, 0.88, 10.16, 10.84]
    for r, h in zip(tbl.rows, heights):
        row_height(r, h)
    r0 = tbl.rows[0].cells
    header_row(tbl.rows[0], ids, r0[0].merge(r0[2]), r0[3], r0[4])
    r1 = tbl.rows[1].cells
    for cell, text, boxed in ((r1[0], "기 호", True), (r1[1], symbol, True)):
        borders(cell, top="single", bottom="single", left="single", right="single")
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        para(cell, text, 10, align=WD_ALIGN_PARAGRAPH.CENTER, first=True)
    merged = r1[2].merge(r1[4])
    borders(merged)
    para(merged, "", first=True)

    def map_cell(row, caption, img, max_h_cm, extra=""):
        c = tbl.rows[row].cells[0].merge(tbl.rows[row].cells[4])
        borders(c, top="single", bottom="single", left="single", right="single")
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        cell_margins(c, 0.08, 0.08, 0.1, 0.05)
        p = para(c, "", align=WD_ALIGN_PARAGRAPH.CENTER, first=True, space_after=2)
        set_font(p.add_run(caption), 11, True, True)
        ip = para(c, "", align=WD_ALIGN_PARAGRAPH.CENTER)
        from PIL import Image

        w, h = Image.open(img).size
        max_w = 17.6
        scale = min(max_w / w, max_h_cm / h)
        ip.add_run().add_picture(str(img), width=Cm(w * scale), height=Cm(h * scale))
        if extra:
            sp = para(c, "", align=WD_ALIGN_PARAGRAPH.RIGHT)
            set_font(sp.add_run(extra), 10)

    map_cell(2, "부 근 위 치", overview_img, 8.6, f"S : {scale_top}")
    map_cell(3, "지적 및 건물 위치도", detail_img, 9.9)
    footer_block(doc, "등록번호 : 34-580   (2010274)   (2018.09.03.)", barcode)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    return out


# --- 건물개황도 (34-581) ----------------------------------------------------


def build_building_form(out: Path, pages: list[dict], barcode: Path | None = None,
                        ids: tuple[str, str, str] = ("", "", "")) -> Path:
    """pages: [{'sketch': Path, 'scale': '1/400', 'rows': [(기호, [산출근거 줄...], 구조, 용도), ...]}]"""
    doc = new_document()
    widths = [1.71, 4.29, 5.80, 0.19, 2.85, 3.14]
    for pi, page in enumerate(pages):
        if pi:
            doc.add_page_break()
        title(doc, "건   물   개   황   도")
        tbl = doc.add_table(rows=10, cols=6)
        fixed_table(tbl, widths)
        for r in tbl.rows:
            row_height(r, 0.88)
        row_height(tbl.rows[2], 15.10)
        r0 = tbl.rows[0].cells
        header_row(tbl.rows[0], ids, r0[0].merge(r0[1]), r0[2].merge(r0[3]), r0[4].merge(r0[5]))
        r1 = tbl.rows[1].cells
        for cell, text in ((r1[0], "기 호"), (r1[1], page.get("symbol", ""))):
            borders(cell, top="single", bottom="single", left="single", right="single")
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            para(cell, text, 10, align=WD_ALIGN_PARAGRAPH.CENTER, first=True)
        rest = r1[2].merge(r1[5])
        borders(rest)
        rest.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        para(rest, f"축적 : {page.get('scale', '1/   ')}", 10, align=WD_ALIGN_PARAGRAPH.RIGHT, first=True)
        sk = tbl.rows[2].cells[0].merge(tbl.rows[2].cells[5])
        borders(sk, top="single", bottom="single", left="single", right="single")
        cell_margins(sk, 0.02, 0.02, 0.02, 0.02)
        sk.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = para(sk, "", align=WD_ALIGN_PARAGRAPH.CENTER, first=True)
        p.add_run().add_picture(str(page["sketch"]), width=Cm(17.9), height=Cm(15.0))
        # 표 머리
        h = tbl.rows[3].cells
        cells = [h[0], h[1].merge(h[2]), h[3].merge(h[4]), h[5]]
        for c, text in zip(cells, ("기 호", "면 적 산 출 근 거", "구조", "용도")):
            borders(c, top="single", bottom="single", left="single", right="single")
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            para(c, text, 10, align=WD_ALIGN_PARAGRAPH.CENTER, first=True)
        rows = page["rows"]
        for i in range(6):
            r = tbl.rows[4 + i].cells
            cells = [r[0], r[1].merge(r[2]), r[3].merge(r[4]), r[5]]
            data = rows[i] if i < len(rows) else ("", [], "", "")
            sym, basis, struct, use = data
            for c, kind in zip(cells, ("sym", "basis", "struct", "use")):
                borders(c, top="single", bottom="single", left="single", right="single")
                c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                cell_margins(c, 0.12, 0.05, 0.02, 0.02)
            para(cells[0], sym, 10, align=WD_ALIGN_PARAGRAPH.CENTER, first=True)
            lines = basis if isinstance(basis, list) else ([basis] if basis else [])
            size = 8.5 if len(lines) <= 1 else 7.5
            para(cells[1], lines[0] if lines else "", size, first=True)
            for extra in lines[1:]:
                para(cells[1], extra, size)
            para(cells[2], struct, 8.5, align=WD_ALIGN_PARAGRAPH.CENTER, first=True)
            use_lines = use if isinstance(use, list) else ([use] if use else [])
            usize = 8.5 if len(use_lines) <= 1 else 7
            para(cells[3], use_lines[0] if use_lines else "", usize, align=WD_ALIGN_PARAGRAPH.CENTER, first=True)
            for extra in use_lines[1:]:
                para(cells[3], extra, usize, align=WD_ALIGN_PARAGRAPH.CENTER)
    footer_block(doc, "등록번호 : 34-581   (2010030)   (2018.09.03.)", barcode)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    return out
