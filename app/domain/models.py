from dataclasses import dataclass
from enum import Enum


class Severity(str, Enum):
    OK = "ok"
    WARNING = "warning"
    ERROR = "error"


@dataclass(frozen=True)
class Finding:
    rule_id: str
    severity: Severity
    message: str
    fixed: bool = False
    location: str = ""


@dataclass(frozen=True)
class FormatConfig:
    font_name: str = "Times New Roman"
    font_size_pt: float = 14
    line_spacing: float = 1.5
    first_line_cm: float = 1.25
    space_before_pt: float = 0
    space_after_pt: float = 0
    margin_left_cm: float = 3.0
    margin_right_cm: float = 1.0
    margin_top_cm: float = 2.0
    margin_bottom_cm: float = 2.0
