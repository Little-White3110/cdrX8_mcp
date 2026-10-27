# -*- coding: utf-8 -*-
"""错误 hint 覆盖率测试：源码里出现的错误码都必须有对应 hint。"""
import pathlib
import re

from cdrx8_mcp import server

_SRC = pathlib.Path(server.__file__).parent


def _error_codes() -> set[str]:
    codes: set[str] = set()
    for py in _SRC.glob("*.py"):
        text = py.read_text(encoding="utf-8", errors="ignore")
        codes |= set(re.findall(r'"error":\s*"([a-z_]+)"', text))
        codes |= set(re.findall(r"'error':\s*'([a-z_]+)'", text))
        codes |= set(re.findall(r'error="([a-z_]+)"', text))
    return codes


def test_every_error_code_has_hint():
    """每个错误码都要在 _ERROR_HINTS 里有可操作建议。"""
    missing = sorted(_error_codes() - set(server._ERROR_HINTS))
    assert not missing, f"缺少 hint 的错误码：{missing}"


def test_guard_adds_hint_to_returned_errors():
    """直接 return 的错误字典也要被 _guard 补上 hint。"""
    out = server._guard(lambda: {"error": "bad_argument", "message": "x"})
    assert out["hint"] == server._ERROR_HINTS["bad_argument"]


def test_save_rejects_missing_dir():
    """cdr_save 目标目录不存在时给出可读错误。"""
    out = server.cdr_save("no-doc", "Z:/missing-dir/out.cdr")
    assert out["error"] == "bad_argument" and "不存在" in out["message"]
    assert out.get("hint")
