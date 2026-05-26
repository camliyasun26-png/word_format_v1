# tests/test_confidence.py
from parser.confidence import explain_confidence

def test_high_confidence_module1():
    result = {"matched_rule": "模块一·H1", "detected_level": "H1"}
    assert explain_confidence(result) == "high"

def test_medium_confidence_module2():
    result = {"matched_rule": "模块二·case1", "detected_level": "H4"}
    assert explain_confidence(result) == "medium"

def test_medium_confidence_single_dot():
    result = {"matched_rule": "模块一·H3", "detected_level": "H3",
              "_context_dependent": True}
    assert explain_confidence(result) == "medium"

def test_low_confidence_module3():
    result = {"matched_rule": "模块三·兜底", "detected_level": "Body"}
    assert explain_confidence(result) == "low"

def test_low_confidence_module4_miss():
    result = {"matched_rule": "模块四·正则失配", "detected_level": "Body"}
    assert explain_confidence(result) == "low"

def test_high_confidence_module4_success():
    result = {"matched_rule": "模块四·拆分", "detected_level": "H2",
              "confidence": "high"}
    assert explain_confidence(result) == "high"
