"""第三十五波（单位换算/颜色工厂/加页）的纯逻辑测试。"""

from __future__ import annotations

import pytest

from cdrx8_mcp import ops
from cdrx8_mcp.errors import CdrBadArgument


def test_wave35_doc_actions_registered():
    for act in ("convert_units", "save_settings", "restore_settings",
                "reset_settings", "resolve_links"):
        assert act in ops.DOC_ACTIONS


def test_convert_units_requires_all_params():
    with pytest.raises(CdrBadArgument):
        ops.doc_ops(None, "convert_units", value=1.0)
    with pytest.raises(CdrBadArgument):
        ops.doc_ops(None, "convert_units", value=1.0, from_unit=3)


def test_color_object_actions_include_factories():
    for act in ("object_model", "object_from_string"):
        assert act in ops.COLOR_OBJECT_ACTIONS


def test_color_model_factory_rejects_unknown_model():
    with pytest.raises(CdrBadArgument):
        ops.color_object_ops("object_model", model="zzz", values=[1])


def test_color_from_string_requires_string():
    with pytest.raises(CdrBadArgument):
        ops.color_object_ops("object_from_string", color2=None)
