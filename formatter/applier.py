# formatter/applier.py
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

_ALIGNMENT_MAP = {
    "CENTER":  WD_ALIGN_PARAGRAPH.CENTER,
    "JUSTIFY": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "LEFT":    WD_ALIGN_PARAGRAPH.LEFT,
    "RIGHT":   WD_ALIGN_PARAGRAPH.RIGHT,
}

_LEVEL_ORDER = ["h1", "h2", "h3", "h4", "h5"]


def _max_preset_level(template: dict) -> str:
    for lvl in reversed(_LEVEL_ORDER):
        if lvl in template:
            return lvl
    return "body"


def _set_run_font_all_slots(run, font_name: str):
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), font_name)
    run.font.name = font_name


def _disable_snap_to_grid(para):
    pPr = para._element.get_or_add_pPr()
    snap = pPr.find(qn("w:snapToGrid"))
    if snap is None:
        snap = OxmlElement("w:snapToGrid")
        pPr.append(snap)
    snap.set(qn("w:val"), "0")


def _set_outline_level(para, level_num):
    pPr = para._element.get_or_add_pPr()
    ol = pPr.find(qn("w:outlineLvl"))
    if level_num is None:
        if ol is not None:
            pPr.remove(ol)
        return
    if ol is None:
        ol = OxmlElement("w:outlineLvl")
        pPr.append(ol)
    ol.set(qn("w:val"), str(level_num))


def apply_format_to_paragraph(para, level: str, template: dict):
    key = level.lower() if level != "Body" else "body"

    if key.startswith("h") and key[1:].isdigit():
        _set_outline_level(para, int(key[1:]) - 1)
    else:
        _set_outline_level(para, None)

    if key not in template:
        hn_key = _max_preset_level(template)
        font_name = template[hn_key]["font_name"]
        for run in para.runs:
            _set_run_font_all_slots(run, font_name)
        return

    cfg = template[key]
    font_name    = cfg["font_name"]
    font_size    = Pt(cfg["font_size"])
    bold         = cfg.get("bold")
    alignment    = _ALIGNMENT_MAP.get(cfg.get("alignment", "JUSTIFY"),
                                      WD_ALIGN_PARAGRAPH.JUSTIFY)
    indent_pt    = cfg.get("indent_first_line", 0)
    line_spacing = cfg.get("line_spacing", 1.5)
    space_before = Pt(cfg.get("space_before", 0))
    space_after  = Pt(cfg.get("space_after", 0))

    para.alignment = alignment
    pf = para.paragraph_format
    pf.first_line_indent = Pt(indent_pt)
    pf.line_spacing = line_spacing
    pf.space_before = space_before
    pf.space_after  = space_after
    _disable_snap_to_grid(para)

    for run in para.runs:
        _set_run_font_all_slots(run, font_name)
        run.font.size = font_size
        if bold is True:
            run.bold = True
        elif bold is False:
            run.bold = False


def apply_font_only(para, font_name: str):
    for run in para.runs:
        _set_run_font_all_slots(run, font_name)


def center_tables_and_images(doc: Document):
    """对表格和图片段落施加居中。原本就居中的不变（写入等价值，不影响其他格式）。"""
    for tbl in doc.tables:
        tblPr = tbl._element.find(qn("w:tblPr"))
        if tblPr is not None:
            tblInd = tblPr.find(qn("w:tblInd"))
            if tblInd is not None:
                tblPr.remove(tblInd)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tblPr = tbl._element.find(qn("w:tblPr"))
        if tblPr is not None:
            jc = tblPr.find(qn("w:jc"))
            if jc is None:
                jc = OxmlElement("w:jc")
                tblPr.append(jc)
            jc.set(qn("w:val"), "center")

    drawing_tag = qn("w:drawing")
    pict_tag = qn("w:pict")
    for para in doc.paragraphs:
        has_media = False
        for el in para._element.iter():
            if el.tag == drawing_tag or el.tag == pict_tag:
                has_media = True
                break
        if has_media:
            pPr = para._element.find(qn("w:pPr"))
            if pPr is not None:
                ind = pPr.find(qn("w:ind"))
                if ind is not None:
                    pPr.remove(ind)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER


def apply_page_margins(doc: Document, template: dict):
    margins = template.get("page_margins_cm", {})
    for section in doc.sections:
        if "top" in margins:
            section.top_margin = Cm(margins["top"])
        if "bottom" in margins:
            section.bottom_margin = Cm(margins["bottom"])
        if "left" in margins:
            section.left_margin = Cm(margins["left"])
        if "right" in margins:
            section.right_margin = Cm(margins["right"])
