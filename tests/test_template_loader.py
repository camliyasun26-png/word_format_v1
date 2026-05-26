# tests/test_template_loader.py
import pytest
from formatter.template_loader import load_template, list_templates

def test_load_template_returns_required_keys():
    tpl = load_template("templates/report_cn.yaml")
    for level in ("h1", "h2", "h3", "h4", "body"):
        assert level in tpl, f"Missing level: {level}"
        assert "font_name" in tpl[level]
        assert "font_size" in tpl[level]

def test_load_template_page_margins():
    tpl = load_template("templates/report_cn.yaml")
    assert "page_margins_cm" in tpl
    assert tpl["page_margins_cm"]["top"] == 2.54

def test_list_templates_finds_yaml(tmp_path):
    (tmp_path / "a.yaml").write_text("name: a\n")
    (tmp_path / "b.yaml").write_text("name: b\n")
    result = list_templates(str(tmp_path))
    assert set(result) == {"a.yaml", "b.yaml"}
