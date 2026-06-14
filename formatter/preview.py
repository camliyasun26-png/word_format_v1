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
    text = text or ""
    # 处理公式级别：用 LaTeX display 模式渲染
    if level == "Equation":
        eq_config = template.get("equation", {})
        alignment = _ALIGNMENT_MAP.get(eq_config.get("alignment", "RIGHT"), "right")
        escaped = html_lib.escape(text)
        return (
            f'<p style="text-align:{alignment};margin:2px 0;padding:2px 0">'
            f'<span class="katex-formula" data-latex="{escaped}">\\({escaped}\\)</span>'
            f'</p>'
        )
    
    key = level.lower() if level != "Body" else "body"

    # 空段落：输出最小高度占位，避免在格式化预览中撑出大量空白
    if not text.strip():
        return '<p style="margin:0;padding:0;line-height:0.8em;font-size:8pt"> </p>'

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
