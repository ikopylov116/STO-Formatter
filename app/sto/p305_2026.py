from app.domain.models import FormatConfig


class P3052026Profile:
    id = "P_305.2026"
    title = "СТО P_305.2026"

    @property
    def config(self) -> FormatConfig:
        return FormatConfig()

    @property
    def structural_elements(self) -> frozenset[str]:
        return frozenset({
            "реферат", "аннотация", "оглавление", "содержание",
            "нормативные ссылки", "определения, обозначения и сокращения",
            "введение", "заключение", "выводы", "заключение/выводы",
            "список использованных источников", "приложения",
        })
