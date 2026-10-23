"""安全层：路径白名单、覆盖保护、环境配置读取。

设计原则：**默认保守，靠环境变量放宽**。
* 未设置 CDR_ALLOWED_ROOTS 时允许写任意路径，但每次写操作都会在返回里带 warning；
  设置了就严格限制在其子目录内。
* 覆盖已有文件必须显式 overwrite=True。
"""

from __future__ import annotations

import os
from pathlib import Path

from .errors import CdrBadArgument, CdrFileExists, CdrPathDenied


def allowed_roots() -> list[Path]:
    """CDR_ALLOWED_ROOTS：用 os.pathsep（Windows 上是 ';'）分隔的目录列表。"""
    raw = os.environ.get("CDR_ALLOWED_ROOTS", "").strip()
    if not raw:
        return []
    roots: list[Path] = []
    for part in raw.split(os.pathsep):
        part = part.strip().strip('"')
        if part:
            try:
                roots.append(Path(part).expanduser().resolve())
            except OSError:
                continue
    return roots


def resolve_read_path(path: str) -> Path:
    """校验读取路径：必须是绝对路径且文件存在。"""
    if not path or not str(path).strip():
        raise CdrBadArgument("path 不能为空")
    p = Path(str(path)).expanduser()
    if not p.is_absolute():
        raise CdrBadArgument("path 必须是绝对路径")
    try:
        p = p.resolve()
    except OSError:
        p = Path(os.path.abspath(str(p)))
    if not p.exists():
        raise CdrBadArgument(f"文件不存在：{p}")
    return p


def resolve_write_path(path: str, *, overwrite: bool = False) -> tuple[Path, str | None]:
    """校验并解析写入路径。返回 (绝对路径, warning 或 None)。"""
    if not path or not str(path).strip():
        raise CdrBadArgument("path 不能为空")

    p = Path(str(path)).expanduser()
    if not p.is_absolute():
        raise CdrBadArgument("path 必须是绝对路径（避免相对路径落到不可预期的目录）")
    try:
        p = p.resolve()
    except OSError:
        p = Path(os.path.abspath(str(p)))

    if not p.parent.exists():
        raise CdrBadArgument(f"目标目录不存在：{p.parent}")

    warning: str | None = None
    roots = allowed_roots()
    if roots:
        if not any(_is_within(p, r) for r in roots):
            raise CdrPathDenied(
                f"{p} 不在允许的目录内；允许的根目录：{', '.join(str(r) for r in roots)}"
            )
    else:
        warning = "未设置 CDR_ALLOWED_ROOTS，本次写入路径未受白名单限制"

    if p.exists() and not overwrite:
        raise CdrFileExists(f"{p} 已存在；确认覆盖请传 overwrite=true")
    return p, warning


def _is_within(child: Path, parent: Path) -> bool:
    try:
        child.relative_to(parent)
        return True
    except ValueError:
        return False


def env_flag(name: str, default: bool = False) -> bool:
    raw = os.environ.get(name)
    if raw is None or raw == "":
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on", "y"}


def env_float(name: str, default: float) -> float:
    raw = os.environ.get(name)
    if not raw:
        return default
    try:
        return float(raw)
    except ValueError:
        return default
