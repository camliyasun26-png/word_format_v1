# tests/test_rules.py
import pytest
from unittest.mock import MagicMock
from parser.rules import classify_paragraph
from parser.context import ChapterContext

def make_para(text: str, bold_runs: list[str]) -> MagicMock:
    para = MagicMock()
    para.text = text
    runs = []
    for t in bold_runs:
        r = MagicMock()
        r.text = t
        r.bold = True
        runs.append(r)
    remaining = text
    for t in bold_runs:
        remaining = remaining.replace(t, "", 1)
    if remaining:
        r = MagicMock()
        r.text = remaining
        r.bold = False
        runs.append(r)
    para.runs = runs
    return para

def test_no_bold_run_returns_body():
    para = make_para("普通段落文字", [])
    ctx = ChapterContext()
    result = classify_paragraph(para, ctx)
    assert result["detected_level"] == "Body"
    assert result["confidence"] == "high"

def test_chapter_heading_h1():
    para = make_para("第一章 总体情况", ["第一章"])
    ctx = ChapterContext()
    result = classify_paragraph(para, ctx)
    assert result["detected_level"] == "H1"
    assert result["confidence"] == "high"

def test_chinese_ordinal_h2():
    para = make_para("一、背景", ["一、"])
    ctx = ChapterContext()
    ctx.push("H1", "第一章")
    result = classify_paragraph(para, ctx)
    assert result["detected_level"] == "H2"

def test_bracketed_chinese_h3():
    para = make_para("（一）基本情况", ["（一）"])
    ctx = ChapterContext()
    ctx.push("H1", "第一章")
    result = classify_paragraph(para, ctx)
    assert result["detected_level"] == "H3"

def test_single_digit_dot_no_context_h3():
    para = make_para("1. 说明", ["1."])
    ctx = ChapterContext()
    ctx.push("H1", "第一章")
    result = classify_paragraph(para, ctx)
    assert result["detected_level"] == "H3"
    assert result["confidence"] == "medium"

def test_single_digit_dot_after_yi_h4():
    para = make_para("1. 说明", ["1."])
    ctx = ChapterContext()
    ctx.push("H1", "第一章")
    ctx.push("H3", "（一）")
    result = classify_paragraph(para, ctx)
    assert result["detected_level"] == "H4"
    assert result["confidence"] == "medium"

def test_two_level_digits_h4():
    para = make_para("1.1 内容", ["1.1"])
    ctx = ChapterContext()
    ctx.push("H1", "第一章")
    result = classify_paragraph(para, ctx)
    assert result["detected_level"] == "H4"

def test_three_level_digits_h5():
    para = make_para("1.1.1 内容", ["1.1.1"])
    ctx = ChapterContext()
    ctx.push("H1", "第一章")
    result = classify_paragraph(para, ctx)
    assert result["detected_level"] == "H5"
