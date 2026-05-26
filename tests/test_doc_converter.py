# tests/test_doc_converter.py
import os
import pytest
from converter.doc_converter import is_doc, to_docx, to_doc

def test_is_doc_true():
    assert is_doc("report.doc") is True

def test_is_doc_false():
    assert is_doc("report.docx") is False

def test_is_doc_false_for_docx():
    assert is_doc("file.DOCX") is False
