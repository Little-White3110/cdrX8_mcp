"""第三十八波（结构体建文档 / Fill 工厂 / 命名箭头 / 边界构造器）纯逻辑测试。"""

from __future__ import annotations

import inspect

from cdrx8_mcp import ops


def test_apply_fill_registered():
    assert "apply_fill" in ops.SHAPE_OPS


def test_arrowhead_named_registered():
    assert "arrowhead_named" in ops.DOC_ACTIONS


def test_ltrb_constructors_registered():
    for kind in ("rect_ltrb", "ellipse_ltrb", "polygon_full"):
        assert kind in ops.CREATABLE


def test_create_document_accepts_options():
    sig = inspect.signature(ops.create_document)
    assert "options" in sig.parameters


def test_open_document_accepts_options():
    sig = inspect.signature(ops.open_document)
    assert "options" in sig.parameters
