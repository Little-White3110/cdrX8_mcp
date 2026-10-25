"""颜色对象 / 描边对象的纯逻辑测试（不依赖 CorelDRAW）。"""

from __future__ import annotations

import pytest

from cdrx8_mcp import ops
from cdrx8_mcp.errors import CdrBadArgument


def test_color_object_actions_cover_models():
    for model in ("rgb", "cmyk", "gray", "hls", "hsb", "lab", "cmy", "yiq"):
        assert model in ops.COLOR_CONVERT_CALLS
        assert model in ops.COLOR_ASSIGN_CALLS
        assert model in ops.COLOR_COMPONENT_ARITY
    assert "registration" in ops.COLOR_ASSIGN_CALLS
    assert "bw" in ops.COLOR_ASSIGN_CALLS


def test_color_component_arity_matches_values():
    assert ops.COLOR_COMPONENT_ARITY["cmyk"] == 4
    assert ops.COLOR_COMPONENT_ARITY["rgb"] == 3
    assert ops.COLOR_COMPONENT_ARITY["gray"] == 1
    assert ops.COLOR_COMPONENT_ARITY["registration"] == 0


def test_color_ops_rejects_unknown_action_before_com():
    with pytest.raises(CdrBadArgument):
        ops.color_ops(None, "nope", color="#000000")


def test_color_assign_rejects_bad_arity_before_com():
    # 分量个数校验发生在造 COM 对象之前，所以 doc=None 也能安全报错
    with pytest.raises(CdrBadArgument):
        ops.color_ops(None, "object_assign", model="cmyk", values=[1, 2])


def test_color_convert_rejects_unknown_model_before_com():
    with pytest.raises(CdrBadArgument):
        ops.color_ops(None, "object_convert", color="#000000", model="xyz")


def test_outline_actions_and_props():
    for act in ("info", "set", "convert_to_object", "copy", "compare"):
        assert act in ops.OUTLINE_ACTIONS
    for key in ("width", "nib_stretch", "nib_angle", "miter_limit", "justification"):
        assert key in ops.OUTLINE_PROPS


def test_outline_ops_rejects_unknown_action_before_com():
    with pytest.raises(CdrBadArgument):
        ops.outline_ops(None, None, {}, "nope")
