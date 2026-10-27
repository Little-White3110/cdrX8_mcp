# -*- coding: utf-8 -*-
"""cdr_help 目录与示例的防漂移测试（纯 Python，不需要 CorelDRAW）。"""

from cdrx8_mcp import ops, server


def test_catalog_covers_all_tools():
    """目录必须覆盖所有已注册的 MCP 工具（避免加了工具忘了登记）。"""
    registered = set(server.mcp._tool_manager._tools)
    catalog = set(server._TOOL_CATALOG)
    assert registered == catalog, f"未登记：{registered - catalog}"


def test_example_tools_exist():
    """示例表里的工具名必须真实存在。"""
    for name in server._ACTION_EXAMPLES:
        assert name in server._TOOL_CATALOG, f"示例表里有未知工具：{name}"


def test_example_actions_exist_in_catalog():
    """每个示例的 action 必须出现在该工具的 action 清单里（默认用法除外）。"""
    for tool, table in server._ACTION_EXAMPLES.items():
        entry = server._catalog_entry(tool, server._TOOL_CATALOG[tool])
        actions = set(entry.get("actions") or [])
        if not actions:
            continue  # 无 action 表的工具（如 cdr_export）不校验
        for action in table:
            if action == "__call__":
                continue
            assert action in actions, f"{tool} 的示例 action {action!r} 不在目录里"


def test_status_tool_count_matches_catalog():
    """cdr_status 的工具数必须与目录一致（唯一真源）。"""
    st = server.cdr_status()
    assert st.get("tools") == len(server._TOOL_CATALOG)


def test_help_detail_shape():
    """cdr_help 的三种调用形态都应返回可用结构。"""
    top = server.cdr_help()
    assert top["tool_count"] >= 31 and isinstance(top["workflow"], list)

    one = server.cdr_help(tool="shape_ops")
    assert one["tool"] == "cdr_shape_ops" and len(one["actions"]) >= 80
    assert isinstance(one.get("examples"), list) and one["examples"]

    ex = server.cdr_help(tool="cdr_shape_ops", action="extrude")
    assert ex["example"]["example"].startswith("cdr_shape_ops(")

    bad = server.cdr_help(tool="nope")
    assert bad["error"] == "unknown_tool" and bad["available"]

    miss = server.cdr_help(tool="shape_ops", action="no_such_action")
    assert miss["example"]["example"] is None
    assert miss["example"]["available_examples"]
