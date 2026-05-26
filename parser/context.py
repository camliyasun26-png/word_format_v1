# parser/context.py
class ChapterContext:
    def __init__(self):
        self._headings: list[tuple[str, str]] = []
        self._in_chapter = False

    def push(self, level: str, text: str):
        if level == "H1":
            self._headings = [(level, text)]
            self._in_chapter = True
        else:
            self._headings.append((level, text))
            if level in ("H2", "H3", "H4", "H5"):
                self._in_chapter = True

    def has_chapter(self) -> bool:
        return self._in_chapter

    def last_heading(self) -> tuple[str, str] | None:
        return self._headings[-1] if self._headings else None

    def last_heading_of_type(self, level: str) -> str | None:
        for lvl, txt in reversed(self._headings):
            if lvl == level:
                return txt
        return None

    def coexisting_bracket_types(self) -> set[str]:
        types = set()
        for _, txt in self._headings:
            if txt.startswith(("(", "（")):
                types.add("case1")
        return types
