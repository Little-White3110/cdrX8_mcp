"""矩形对象 / 文字三态 / 曲线段的纯逻辑测试（不依赖 CorelDRAW）。"""

from __future__ import annotations

import pytest

from cdrx8_mcp import ops
from cdrx8_mcp.errors import CdrBadArgument


def test_tristate_uses_cdr_enum():
    # X8 的 cdrTriState：true = -1，false = 0（传 True 会被当成非法值把粗体关掉）
    assert ops._tristate(True) == -1
    assert ops._tristate(False) == 0
    assert ops._tristate(None) == 0


def test_underline_is_not_tristate():
    # 下划线形参是样式枚举：0=无，1..5=有（-1 等于无）
    assert ops._underline(True) == 1
    assert ops._underline(False) == 0


def test_rect_actions_registered():
    for act in ("info", "offset", "inflate", "union", "intersect", "contains",
                "copy", "create_rectangle", "create_ellipse"):
        assert act in ops.RECT_ACTIONS
    assert set(ops.RECT_ACTIONS) == set(ops.RECT_READ_ACTIONS) | {
        "create_rectangle", "create_ellipse"}


def test_rect_ops_rejects_unknown_action_before_com():
    with pytest.raises(CdrBadArgument):
        ops.rect_ops(None, "nope")


def test_curve_segment_is_creatable():
    assert "curve_segment" in ops.CREATABLE
