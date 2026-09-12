from docx import Document
from app.domain.models import Finding, Severity, FormatConfig
from app.document.analyzer import analyze_document


def _near(actual: float, expected: float, tolerance: float = 0.05) -> bool:
    return abs(actual - expected) <= tolerance


def check_document(path: str, cfg: FormatConfig, structural_elements: frozenset[str] = frozenset()) -> list[Finding]:
    doc = Document(path)
    out: list[Finding] = []

    for i, section in enumerate(doc.sections, 1):
        fields = [
            ("левое поле", section.left_margin.cm, cfg.margin_left_cm),
            ("правое поле", section.right_margin.cm, cfg.margin_right_cm),
            ("верхнее поле", section.top_margin.cm, cfg.margin_top_cm),
            ("нижнее поле", section.bottom_margin.cm, cfg.margin_bottom_cm),
        ]
        for name, actual, expected in fields:
            ok = _near(actual, expected)
            out.append(Finding("P305-4.1.3", Severity.OK if ok else Severity.ERROR,
                               f"Раздел {i}: {name} {actual:.2f} см; требуется {expected:.2f} см.", location=f"section {i}"))

    paragraphs = analyze_document(path, structural_elements)
    double_spaces = sum("  " in p.text for p in paragraphs if p.text)
    out.append(Finding("P305-4.1.13", Severity.OK if not double_spaces else Severity.WARNING,
                       f"Абзацев с двойными пробелами: {double_spaces}."))

    nonempty = 0
    wrong_font = 0
    wrong_indent = 0
    wrong_spacing = 0
    for paragraph in doc.paragraphs:
        if not paragraph.text.strip():
            continue
        nonempty += 1
        pf = paragraph.paragraph_format
        if pf.first_line_indent is not None and not _near(pf.first_line_indent.cm, cfg.first_line_cm, 0.08):
            text = paragraph.text.strip().lower().rstrip(".")
            if text not in structural_elements:
                wrong_indent += 1
        if pf.space_before is not None and abs(pf.space_before.pt - cfg.space_before_pt) > 0.1:
            wrong_spacing += 1
        if pf.space_after is not None and abs(pf.space_after.pt - cfg.space_after_pt) > 0.1:
            wrong_spacing += 1
        for run in paragraph.runs:
            if run.font.name and run.font.name != cfg.font_name:
                wrong_font += 1

    out.append(Finding("P305-4.1.5", Severity.OK if not wrong_font else Severity.WARNING,
                       f"Фрагментов с явно отличающимся шрифтом: {wrong_font} (проверено непустых абзацев: {nonempty})."))
    out.append(Finding("P305-4.1.5", Severity.OK if not wrong_indent else Severity.WARNING,
                       f"Непустых абзацев с отличающимся абзацным отступом: {wrong_indent}."))
    out.append(Finding("P305-4.1.5", Severity.OK if not wrong_spacing else Severity.WARNING,
                       f"Абзацев с отличающимися интервалами до/после: {wrong_spacing}."))
    return out
