# parser/rules.py
import re
from parser.context import ChapterContext

_LPAR = r"[(\（]"
_RPAR = r"[)\）]"
_DOT  = r"[.．]"

_RE_H1 = re.compile(r"^\s*(第\s*[0-9０-９一二三四五六七八九十百千]+\s*章)")
_RE_H2 = re.compile(r"^\s*([一二三四五六七八九十百千]+\s*、)")
_RE_H3_BRACKET = re.compile(r"^\s*[(\（]\s*[一二三四五六七八九十百千]+\s*[)\）]")
_RE_H3_SINGLE  = re.compile(r"^\s*[0-9０-９]+\s*[^\w,，。；;)）(（]\s*(?![0-9０-９])")
_RE_H4_TWO    = re.compile(r"^\s*[0-9０-９]+\s*[^\w,，。；;)）(（]\s*[0-9０-９]+\s*(?:[^\w,，。；;)）(（]\s*)?(?![0-9０-９])")
_RE_H5_THREE  = re.compile(r"^\s*[0-9０-９]+\s*[^\w,，。；;)）(（]\s*[0-9０-９]+\s*[^\w,，。；;)）(（]\s*[0-9０-９]+")


def _has_any_bold(para) -> bool:
    return any(getattr(r, "bold", False) for r in para.runs)


def _bold_prefix(para) -> str:
    result = []
    started = False
    for r in para.runs:
        text = r.text
        if not started and (not text or text.isspace()):
            result.append(text)
            continue
        if getattr(r, "bold", False):
            started = True
            result.append(text)
        else:
            break
    return "".join(result)


def _module1_level(text: str, ctx: ChapterContext) -> tuple[str, str] | None:
    if _RE_H1.match(text):
        return ("H1", "high")
    if _RE_H2.match(text):
        return ("H2", "high")
    if _RE_H3_BRACKET.match(text):
        return ("H3", "high")
    if _RE_H5_THREE.match(text):
        return ("H5", "high")
    if _RE_H4_TWO.match(text):
        return ("H4", "high")
    if _RE_H3_SINGLE.match(text):
        prev = None
        for lvl, txt in reversed(ctx._headings):
            is_single = _RE_H3_SINGLE.match(txt) and not _RE_H4_TWO.match(txt) and not _RE_H5_THREE.match(txt)
            if is_single:
                continue
            prev = (lvl, txt)
            break
        if prev and prev[0] == "H3" and _RE_H3_BRACKET.match(prev[1]):
            return ("H4", "medium")
        return ("H3", "medium")
    return None


def classify_paragraph(para, ctx: ChapterContext) -> dict:
    text = para.text.strip()

    if not _has_any_bold(para):
        return {
            "original_text": text,
            "detected_level": "Body",
            "confidence": "high",
            "matched_rule": "前置规则·无加粗",
            "is_split": False,
        }

    prefix = _bold_prefix(para)

    result = _module4(text, prefix, ctx)
    if result:
        return result

    m1 = _module1_level(prefix, ctx)
    if m1:
        level, conf = m1
        return {
            "original_text": text,
            "detected_level": level,
            "confidence": conf,
            "matched_rule": f"模块一·{level}",
            "is_split": False,
        }

    result = _module2(text, prefix, ctx)
    if result:
        return result

    return {
        "original_text": text,
        "detected_level": "Body",
        "confidence": "low",
        "matched_rule": "模块三·兜底",
        "is_split": False,
    }


def _module4(text: str, prefix: str, ctx: ChapterContext) -> dict | None:
    _RE_DIGIT_START = re.compile(r"^[0-9０-９一二三四五六七八九十百千]")
    _RE_COLON = re.compile(r"[：:]")
    if not _RE_DIGIT_START.match(prefix):
        return None
    colon_match = _RE_COLON.search(text)
    if not colon_match:
        return None

    colon_pos = colon_match.start()
    before = text[:colon_pos].strip()
    after  = text[colon_pos + 1:].strip()

    m1 = _module1_level(before, ctx)
    if m1 is None:
        return {
            "original_text": text,
            "detected_level": "Body",
            "confidence": "low",
            "matched_rule": "模块四·正则失配",
            "is_split": False,
        }

    level, inner_conf = m1
    if len(after) > 34:
        return {
            "original_text": text,
            "detected_level": level,
            "confidence": inner_conf,
            "matched_rule": "模块四·拆分",
            "is_split": True,
            "split_title": before,
            "split_body": after,
        }
    else:
        return {
            "original_text": text,
            "detected_level": level,
            "confidence": inner_conf,
            "matched_rule": "模块四·整体标题",
            "is_split": False,
        }


def _module2(text: str, prefix: str, ctx: ChapterContext) -> dict | None:
    _RE_CASE1 = re.compile(r"^\s*[(\（]\s*[0-9a-zA-ZαβγδεζηθιοΑΒΓΔΕΖΗΘΙΟ]+\s*[)\）]")
    _RE_CASE2 = re.compile(r"^\s*[0-9a-zA-ZαβγδεζηθιοΑΒΓΔΕΖΗΘΙΟ]+\s*[)\）]")

    if not ctx.has_chapter():
        return None

    def _last_non_matching(pattern):
        for lvl, txt in reversed(ctx._headings):
            if not pattern.match(txt):
                return (lvl, txt)
        return None

    if _RE_CASE1.match(prefix):
        prev = _last_non_matching(_RE_CASE1)
        level = _shift_down(prev[0]) if prev else "H4"
        return {
            "original_text": text,
            "detected_level": level,
            "confidence": "medium",
            "matched_rule": "模块二·case1",
            "is_split": False,
        }

    if _RE_CASE2.match(prefix):
        prev = _last_non_matching(_RE_CASE2)
        level = _shift_down(prev[0]) if prev else "H5"
        return {
            "original_text": text,
            "detected_level": level,
            "confidence": "medium",
            "matched_rule": "模块二·case2",
            "is_split": False,
        }

    return None


_LEVEL_ORDER = ["H1", "H2", "H3", "H4", "H5"]

def _shift_down(level: str) -> str:
    if level not in _LEVEL_ORDER:
        return "H5"
    idx = _LEVEL_ORDER.index(level)
    return _LEVEL_ORDER[min(idx + 1, len(_LEVEL_ORDER) - 1)]
