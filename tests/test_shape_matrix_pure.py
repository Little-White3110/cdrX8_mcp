"""第三十九波（形状矩阵 / 扭曲 / 文字内嵌 / 范围复制）纯逻辑测试。"""

from __future__ import annotations

import inspect

from cdrx8_mcp import ops


def test_wave39_shape_ops_registered():
    for op in ("matrix", "is_type", "in_front_of", "is_on_shape", "distort",
               "place_text", "duplicate_range", "clone_range"):
        assert op in ops.SHAPE_OPS


def test_shape_op_accepts_params():
    sig = inspect.signature(ops.shape_op)
    assert "params" in sig.parameters
