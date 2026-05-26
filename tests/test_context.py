# tests/test_context.py
from parser.context import ChapterContext

def test_no_h1_returns_no_context():
    ctx = ChapterContext()
    assert ctx.has_chapter() is False

def test_push_h1_creates_chapter():
    ctx = ChapterContext()
    ctx.push("H1", "第一章")
    assert ctx.has_chapter() is True

def test_last_heading_in_chapter():
    ctx = ChapterContext()
    ctx.push("H1", "第一章")
    ctx.push("H2", "一、")
    assert ctx.last_heading() == ("H2", "一、")

def test_new_h1_resets_chapter():
    ctx = ChapterContext()
    ctx.push("H1", "第一章")
    ctx.push("H2", "一、")
    ctx.push("H1", "第二章")
    assert ctx.last_heading() == ("H1", "第二章")

def test_last_heading_of_type():
    ctx = ChapterContext()
    ctx.push("H1", "第一章")
    ctx.push("H3", "（一）")
    ctx.push("H4", "1.")
    assert ctx.last_heading_of_type("H3") == "（一）"
    assert ctx.last_heading_of_type("H2") is None
