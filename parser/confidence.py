# parser/confidence.py

_LOW_RULES = {"模块三·兜底", "模块四·正则失配", "前置规则·无上下文回退"}
_MEDIUM_RULES = {"模块二·case1", "模块二·case2"}


def explain_confidence(result: dict) -> str:
    rule = result.get("matched_rule", "")
    if rule in _LOW_RULES:
        return "low"
    if rule in _MEDIUM_RULES:
        return "medium"
    if rule in ("模块四·拆分", "模块四·整体标题"):
        return result.get("confidence", "high")
    if result.get("_context_dependent"):
        return "medium"
    return result.get("confidence", "high")
