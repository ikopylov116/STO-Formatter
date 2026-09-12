import re
from dataclasses import dataclass
from docx import Document


@dataclass(frozen=True)
class ParagraphInfo:
    index: int
    text: str
    structural: bool
    numbered_heading: bool


_NUMBERED = re.compile(r"^\d+(?:\.\d+)*\s+\S")


def is_structural(text: str, structural_elements: frozenset[str]) -> bool:
    return text.strip().lower().rstrip(".") in structural_elements


def is_numbered_heading(text: str) -> bool:
    return bool(_NUMBERED.match(text.strip()))


def analyze_document(path: str, structural_elements: frozenset[str]) -> list[ParagraphInfo]:
    doc = Document(path)
    result: list[ParagraphInfo] = []
    for i, paragraph in enumerate(doc.paragraphs, 1):
        text = paragraph.text.strip()
        result.append(ParagraphInfo(i, text, is_structural(text, structural_elements), is_numbered_heading(text)))
    return result
