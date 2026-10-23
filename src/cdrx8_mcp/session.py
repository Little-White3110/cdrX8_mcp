"""文档会话：document_id 的分配、revision 计数、快照令牌。

为什么要这层：
* COM 对象不能跨线程持有，所以会话里**只存文档名**，每次操作时按名重新取文档；
* Agent 可能持有过期的对象视图（用户手动改了图、或上一轮工具调用已经改了），
  所以每次变更都要求带 `expected_revision` + `snapshot_token`，对不上就拒绝执行。
"""

from __future__ import annotations

import hashlib
import time
from dataclasses import dataclass, field

from .errors import CdrDocumentNotFound


def snapshot_token(name: str, revision: int, pages: int, shapes: int) -> str:
    """把「文档身份 + 修订号 + 结构规模」揉成一个短令牌。

    带上页数和图形数，是为了让**用户在 CorelDRAW 里手动增删对象**也能被检测到——
    仅仅靠 revision（只有本服务修改才自增）会漏掉这类外部改动。
    """
    raw = f"{name}|{revision}|{pages}|{shapes}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]


@dataclass
class DocSession:
    sid: str
    name: str
    revision: int = 0
    created_at: float = field(default_factory=time.time)

    def bump(self) -> int:
        self.revision += 1
        return self.revision


class SessionRegistry:
    """按文档名索引会话，进程内单例。"""

    def __init__(self) -> None:
        self._by_name: dict[str, DocSession] = {}
        self._seq = 0

    def get_or_create(self, name: str) -> DocSession:
        sess = self._by_name.get(name)
        if sess is None:
            self._seq += 1
            sess = DocSession(sid=f"doc{self._seq}", name=name)
            self._by_name[name] = sess
        return sess

    def by_sid(self, sid: str) -> DocSession:
        for sess in self._by_name.values():
            if sess.sid == sid:
                return sess
        raise CdrDocumentNotFound(f"未知的 document_id：{sid}")

    def rename(self, old: str, new: str) -> None:
        """另存为之后文档名会变，会话要跟着迁移（保留 revision）。"""
        if old == new:
            return
        sess = self._by_name.pop(old, None)
        if sess is not None:
            sess.name = new
            self._by_name[new] = sess

    def drop(self, name: str) -> None:
        self._by_name.pop(name, None)

    def all(self) -> list[DocSession]:
        return list(self._by_name.values())


_REGISTRY = SessionRegistry()


def registry() -> SessionRegistry:
    return _REGISTRY
