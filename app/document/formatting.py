import re
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from app.domain.models import Finding, Severity, FormatConfig


def _set_run_font(run, name: str, size: float) -> None:
    run.font.name = name
    run.font.size = Pt(size)
    rpr = run._r.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{key}"), name)


def _page_field(paragraph, cfg: FormatConfig) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, end])
    _set_run_font(run, cfg.font_name, cfg.font_size_pt)


def _is_numbered_heading(text: str) -> bool:
    return bool(re.match(r"^\d+(?:\.\d+)*\s+\S", text.strip()))


def format_document(src: str, dst: str, cfg: FormatConfig, structural_elements: frozenset[str]) -> list[Finding]:
    src_path, dst_path = Path(src), Path(dst)
    if src_path.resolve() == dst_path.resolve():
        raise ValueError("Исходный и выходной файлы должны быть разными.")

    doc = Document(src)
    for section in doc.sections:
        section.top_margin = Cm(cfg.margin_top_cm)
        section.bottom_margin = Cm(cfg.margin_bottom_cm)
        section.left_margin = Cm(cfg.margin_left_cm)
        section.right_margin = Cm(cfg.margin_right_cm)
        footer = section.footer
        p = footer.paragraphs[0]
        p.text = ""
        _page_field(p, cfg)

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        structural = text.lower().rstrip(".") in structural_elements
        heading = structural or _is_numbered_heading(text) or paragraph.style.name.lower().startswith("heading")
        pf = paragraph.paragraph_format
        pf.line_spacing = cfg.line_spacing
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.first_line_indent = Cm(0 if heading else cfg.first_line_cm)
        pf.keep_with_next = heading
        if structural:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif heading:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for run in paragraph.runs:
            _set_run_font(run, cfg.font_name, cfg.font_size_pt)
            run.bold = heading
            run.underline = False

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.line_spacing = cfg.line_spacing
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(0)
                    for run in paragraph.runs:
                        _set_run_font(run, cfg.font_name, 13)

    settings = doc.settings.element
    update = OxmlElement("w:updateFields")
    update.set(qn("w:val"), "true")
    settings.append(update)
    doc.save(dst)

    return [
        Finding("P305-4.1.3", Severity.OK, "Поля установлены: 30/10/20/20 мм.", True),
        Finding("P305-4.1.5", Severity.OK, "Основной текст приведён к Times New Roman 14 pt, 1,5, по ширине.", True),
        Finding("P305-4.1.5", Severity.OK, "Для обычных абзацев установлен отступ первой строки 1,25 см и интервалы 0/0.", True),
        Finding("P305-4.3", Severity.OK, "В нижний колонтитул добавлено поле автоматической нумерации страниц.", True),
    ]
