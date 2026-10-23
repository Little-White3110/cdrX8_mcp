"""会话与快照令牌的单元测试。"""

from __future__ import annotations

import pytest

from cdrx8_mcp.errors import CdrDocumentNotFound
from cdrx8_mcp.session import SessionRegistry, snapshot_token


def test_token_is_deterministic():
    a = snapshot_token("未命名-1", 0, 1, 3)
    b = snapshot_token("未命名-1", 0, 1, 3)
    assert a == b and len(a) == 16


@pytest.mark.parametrize(
    "args",
    [
        ("未命名-2", 0, 1, 3),   # 换文档
        ("未命名-1", 1, 1, 3),   # revision 变了
        ("未命名-1", 0, 2, 3),   # 页数变了（用户手动加页）
        ("未命名-1", 0, 1, 4),   # 图形数变了（用户手动加对象）
    ],
)
def test_token_changes_with_any_input(args):
    assert snapshot_token(*args) != snapshot_token("未命名-1", 0, 1, 3)


def test_registry_reuses_session_by_name():
    reg = SessionRegistry()
    s1 = reg.get_or_create("A")
    s2 = reg.get_or_create("A")
    assert s1 is s2 and s1.sid == s2.sid


def test_registry_assigns_distinct_ids_and_bumps():
    reg = SessionRegistry()
    a = reg.get_or_create("A")
    b = reg.get_or_create("B")
    assert a.sid != b.sid
    assert a.revision == 0
    assert a.bump() == 1
    assert a.bump() == 2


def test_registry_by_sid_and_unknown():
    reg = SessionRegistry()
    a = reg.get_or_create("A")
    assert reg.by_sid(a.sid) is a
    with pytest.raises(CdrDocumentNotFound):
        reg.by_sid("nope")


def test_registry_rename_keeps_revision():
    reg = SessionRegistry()
    a = reg.get_or_create("未命名-1")
    a.bump()
    reg.rename("未命名-1", "成品.cdr")
    assert a.name == "成品.cdr"
    assert a.revision == 1
    assert reg.get_or_create("成品.cdr") is a
