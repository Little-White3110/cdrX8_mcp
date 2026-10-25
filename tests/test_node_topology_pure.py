"""第三十六波（节点拓扑/曲线合并）的纯逻辑测试。"""

from __future__ import annotations

from cdrx8_mcp import ops


def test_node_topology_actions_registered():
    for act in ("neighbors", "distance", "average", "join_nodes"):
        assert act in ops.NODE_ACTIONS
    # 变更类动作要登记在 NODE_MUTATIONS 里（决定是否走 undo 组）
    for act in ("average", "join_nodes"):
        assert act in ops.NODE_MUTATIONS


def test_node_readonly_actions_not_mutations():
    # 纯读动作不应被当成变更
    for act in ("neighbors", "distance"):
        assert act not in ops.NODE_MUTATIONS


def test_curve_merge_actions_registered():
    for act in ("merge_curve", "weld_curve", "is_on_curve", "curve_copy",
                "curve_info"):
        assert act in ops.PATH_ACTIONS
