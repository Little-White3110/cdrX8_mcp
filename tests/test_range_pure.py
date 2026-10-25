"""对齐/范围批量的纯逻辑测试（不依赖 CorelDRAW）。"""

from __future__ import annotations

import pytest

from cdrx8_mcp import ops
from cdrx8_mcp.errors import CdrBadArgument


def test_range_ops_registered():
    for op in ("unlock_all", "ungroup_all", "break_apart_ex", "clear_effect",
               "range_info", "range_add", "range_remove"):
        assert op in ops.SHAPE_OPS


def test_align_rejects_unknown_target():
    with pytest.raises(CdrBadArgument):
        ops.align_shapes(None, horizontal="left", to="moon")


def test_align_point_requires_coordinates():
    with pytest.raises(CdrBadArgument):
        ops.align_shapes(None, horizontal="left", to="point")


def test_align_shape_requires_reference():
    with pytest.raises(CdrBadArgument):
        ops.align_shapes(None, horizontal="left", to="shape")
    with pytest.raises(CdrBadArgument):
        ops.align_shapes(None, horizontal="left", to="range")


def test_align_requires_axis():
    with pytest.raises(CdrBadArgument):
        ops.align_shapes(None, horizontal="none", vertical="none", to="page")
