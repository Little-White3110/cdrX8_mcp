"""吸附点纯逻辑测试：不依赖 CorelDRAW，只验证参数守卫与辅助函数。"""

from __future__ import annotations

import pytest

from cdrx8_mcp import ops
from cdrx8_mcp.errors import CdrBadArgument


def test_snap_type_name_covers_mask():
    assert ops._snap_type_name(1) == "user"
    assert ops._snap_type_name(2) == "object"
    assert ops._snap_type_name(4) == "bbox"
    assert ops._snap_type_name(8) == "edge"
    assert ops._snap_type_name(16) == "free"
    assert ops._snap_type_name(32) == "auto"


def test_snap_type_name_unknown_and_none():
    assert ops._snap_type_name(64) == "type_64"
    assert ops._snap_type_name(None) is None


def test_snap_kind_counts_groups_by_kind():
    counts = ops._snap_kind_counts([
        {"kind": "object"}, {"kind": "object"}, {"kind": "bbox"}, {"kind": None, "type": 16},
    ])
    assert counts == {"object": 2, "bbox": 1, "16": 1}


def test_snap_actions_are_consistent():
    # 读动作必须是全动作的子集
    assert set(ops.SNAP_READ_ACTIONS) <= set(ops.SNAP_ACTIONS)
    for act in ("list", "add_user", "move", "range", "bbox", "edge", "user", "near"):
        assert act in ops.SNAP_ACTIONS
    assert "auto_snap" in ops.SNAP_RANGE_OPS
    assert "change_direction" in ops.SNAP_RANGE_OPS


def test_snap_ops_rejects_unknown_action_without_touching_com():
    # 动作校验发生在进入 COM/doc 之前，所以 doc=None 也能安全报错
    with pytest.raises(CdrBadArgument):
        ops.snap_ops(None, "nope")


def test_snap_range_requires_indices():
    with pytest.raises(CdrBadArgument):
        ops._snap_range(None, [])
