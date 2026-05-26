# tests/test_preview.py
from formatter.preview import render_paragraph_html
from formatter.template_loader import load_template

TPL = load_template("templates/report_cn.yaml")

def test_h1_renders_center_bold():
    html = render_paragraph_html("第一章 总体情况", "H1", TPL)
    assert 'text-align:center' in html
    assert 'font-weight:bold' in html
    assert '仿宋' in html or 'FangSong' in html

def test_body_renders_indent():
    html = render_paragraph_html("普通正文内容", "Body", TPL)
    assert 'text-indent' in html
    assert 'font-weight:bold' not in html

def test_h2_renders_justify():
    html = render_paragraph_html("一、背景", "H2", TPL)
    assert 'text-align:justify' in html
    assert 'font-weight:bold' in html

def test_output_is_p_tag():
    html = render_paragraph_html("文字", "Body", TPL)
    assert html.startswith("<p ") or html.startswith("<p>")
