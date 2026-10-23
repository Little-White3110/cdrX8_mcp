"""纯逻辑单元测试：不依赖 CorelDRAW，可在任何平台运行。"""

from __future__ import annotations

import os

import pytest

from cdrx8_mcp.constants import FILTER_PNG, FILTER_SVG, PROGIDS, TEXT_TYPE, UNIT_MM, UNIT_NAMES
from cdrx8_mcp.errors import CdrBadArgument, CdrFileExists, CdrPathDenied
from cdrx8_mcp.safety import allowed_roots, env_flag, resolve_read_path, resolve_write_path


def test_constants_match_measured_x8_values():
    # 这三个数字是 X8 实测标定结果，改动会导致几何 / 导出错乱
    assert UNIT_MM == 3, "实测：3 才是毫米；2 是英尺，4 是厘米"
    assert UNIT_NAMES[2] == "ft"
    assert UNIT_NAMES[4] == "cm"
    assert FILTER_PNG == 802
    assert FILTER_SVG == 1345
    assert TEXT_TYPE == 6, "实测：X8 的艺术字与段落文字 Type 都是 6"
    assert PROGIDS[0] == "CorelDRAW.Application"
    assert "CorelDRAW.Application.18" in PROGIDS


def test_allowed_roots_empty_by_default(monkeypatch):
    monkeypatch.delenv("CDR_ALLOWED_ROOTS", raising=False)
    assert allowed_roots() == []


def test_allowed_roots_parses_pathsep(monkeypatch, tmp_path):
    a, b = tmp_path / "a", tmp_path / "b"
    a.mkdir()
    b.mkdir()
    monkeypatch.setenv("CDR_ALLOWED_ROOTS", f"{a}{os.pathsep}{b}")
    roots = allowed_roots()
    assert len(roots) == 2
    assert all(r.is_absolute() for r in roots)


def test_write_path_rejects_relative(monkeypatch):
    monkeypatch.delenv("CDR_ALLOWED_ROOTS", raising=False)
    with pytest.raises(CdrBadArgument):
        resolve_write_path("relative/out.png")


def test_write_path_rejects_missing_parent(monkeypatch, tmp_path):
    monkeypatch.delenv("CDR_ALLOWED_ROOTS", raising=False)
    with pytest.raises(CdrBadArgument):
        resolve_write_path(str(tmp_path / "nope" / "out.png"))


def test_write_path_respects_whitelist(monkeypatch, tmp_path):
    allowed = tmp_path / "allowed"
    other = tmp_path / "other"
    allowed.mkdir()
    other.mkdir()
    monkeypatch.setenv("CDR_ALLOWED_ROOTS", str(allowed))

    p, warn = resolve_write_path(str(allowed / "ok.png"))
    assert p == (allowed / "ok.png").resolve()
    assert warn is None

    with pytest.raises(CdrPathDenied):
        resolve_write_path(str(other / "bad.png"))


def test_write_path_warns_without_whitelist(monkeypatch, tmp_path):
    monkeypatch.delenv("CDR_ALLOWED_ROOTS", raising=False)
    _p, warn = resolve_write_path(str(tmp_path / "out.png"))
    assert warn and "CDR_ALLOWED_ROOTS" in warn


def test_write_path_overwrite_guard(monkeypatch, tmp_path):
    monkeypatch.delenv("CDR_ALLOWED_ROOTS", raising=False)
    target = tmp_path / "exists.png"
    target.write_text("x")
    with pytest.raises(CdrFileExists):
        resolve_write_path(str(target))
    p, _warn = resolve_write_path(str(target), overwrite=True)
    assert p == target.resolve()


def test_read_path_requires_existing_absolute(tmp_path):
    with pytest.raises(CdrBadArgument):
        resolve_read_path("relative.cdr")
    with pytest.raises(CdrBadArgument):
        resolve_read_path(str(tmp_path / "missing.cdr"))
    f = tmp_path / "a.cdr"
    f.write_text("x")
    assert resolve_read_path(str(f)) == f.resolve()


@pytest.mark.parametrize(
    "raw,expected",
    [(None, False), ("1", True), ("true", True), ("ON", True), ("yes", True),
     ("0", False), ("false", False), ("", False)],
)
def test_env_flag(monkeypatch, raw, expected):
    if raw is None:
        monkeypatch.delenv("CDR_FLAG_X", raising=False)
    else:
        monkeypatch.setenv("CDR_FLAG_X", raw)
    assert env_flag("CDR_FLAG_X") is expected
