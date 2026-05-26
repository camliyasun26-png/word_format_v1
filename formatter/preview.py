# formatter/preview.py
import html as html_lib

_ALIGNMENT_MAP = {
    "CENTER": "center",
    "JUSTIFY": "justify",
    "LEFT": "left",
    "RIGHT": "right",
}

_LEVEL_ORDER = ["h1", "h2", "h3", "h4", "h5"]


def _max_preset_level(template: dict) -> str:
    for lvl in reversed(_LEVEL_ORDER):
        if lvl in template:
            return lvl
    return "body"


def render_paragraph_html(text: str, level: str, template: dict) -> str:
    key = level.lower() if level != "Body" else "body"

    if key not in template:
        # Hj 超出预设 Hn：仅用 Hn 的字体，其余样式不应用预设格式
        hn_key = _max_preset_level(template)
        font_name = template[hn_key].get("font_name", "仿宋_GB2312")
        escaped = html_lib.escape(text)
        return f'<p style="font-family:\'{font_name}\',FangSong,serif;margin:0;padding:2px 0">{escaped}</p>'

    cfg = template[key]
    font_name = cfg.get("font_name", "仿宋_GB2312")
    font_size_pt = cfg.get("font_size", 12)
    bold = cfg.get("bold", False)
    alignment = _ALIGNMENT_MAP.get(cfg.get("alignment", "JUSTIFY"), "justify")
    indent_pt = cfg.get("indent_first_line", 24)
    line_spacing = cfg.get("line_spacing", 1.5)

    styles = [
        f"font-family:'{font_name}',FangSong,serif",
        f"font-size:{font_size_pt}pt",
        f"font-weight:{'bold' if bold else 'normal'}",
        f"text-align:{alignment}",
        f"text-indent:{indent_pt}pt",
        f"line-height:{line_spacing}",
        "margin:0",
        "padding:2px 0",
    ]
    style_str = ";".join(styles)
    escaped = html_lib.escape(text)
    return f'<p style="{style_str}">{escaped}</p>'
