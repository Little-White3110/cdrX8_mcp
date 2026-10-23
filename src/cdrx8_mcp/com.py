"""COM 桥：所有 CorelDRAW 调用都在**同一个专用 STA 线程**上串行执行。

为什么必须这样（实测与业界实践一致）：
* CorelDRAW 的 COM 对象只能在单线程套间（STA）里用，跨线程共享 COM 对象是崩溃头号原因；
* 只要某个调用卡在模态弹窗（缺字体、色彩配置转换），整个通道就堵住 ——
  所以这里做「一次只跑一个任务 + 超时报警」，绝不并发；
* COM 调用**不可中途取消**：超时只是本服务不再等待，CorelDRAW 仍在执行。
  因此写操作超时后不要自动重试（可能重复执行）。
"""

from __future__ import annotations

import queue
import threading
import time
from typing import Any, Callable

from .constants import DEFAULT_TIMEOUT_S, PROGIDS
from .errors import CdrComFailure, CdrNotConnected, CdrTimeout

try:  # pragma: no cover - 仅在 Windows 上可导入
    import pythoncom
    import pywintypes
    import win32com.client as win32
except ImportError:  # pragma: no cover
    pythoncom = None  # type: ignore
    pywintypes = None  # type: ignore
    win32 = None  # type: ignore


def _wrap_com_error(exc: BaseException) -> BaseException:
    """把 pywin32 的 com_error 转成带 hresult 的可读异常。"""
    if pywintypes is not None and isinstance(exc, pywintypes.com_error):
        args = exc.args
        hresult = args[0] if args else None
        text = str(args[1]) if len(args) > 1 and args[1] else str(exc)
        return CdrComFailure(f"CorelDRAW COM 调用失败：{text}", hresult=hresult)
    return exc


class ComBridge:
    """单线程 STA 执行器。所有 COM 操作都必须经由 `call()`。"""

    def __init__(self, timeout: float = DEFAULT_TIMEOUT_S):
        self.default_timeout = timeout
        self._q: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None
        self._ready = threading.Event()
        self._app: Any = None
        self._start_error: BaseException | None = None

    # ---------------- 生命周期 ----------------
    def start(self) -> None:
        if self._thread is not None and self._thread.is_alive():
            return
        self._ready.clear()
        self._start_error = None
        self._thread = threading.Thread(target=self._loop, name="cdr-com-sta", daemon=True)
        self._thread.start()
        if not self._ready.wait(10):
            raise CdrComFailure("COM 线程未能启动")
        if self._start_error is not None:
            raise _wrap_com_error(self._start_error)

    def _loop(self) -> None:  # pragma: no cover - 线程体
        try:
            pythoncom.CoInitialize()
        except BaseException as exc:  # noqa: BLE001
            self._start_error = exc
            self._ready.set()
            return
        self._ready.set()
        try:
            while True:
                job = self._q.get()
                if job is None:
                    break
                fn, box, done = job
                try:
                    box["value"] = fn(self)
                except BaseException as exc:  # noqa: BLE001 - 原样带回调用线程
                    box["error"] = exc
                finally:
                    done.set()
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def shutdown(self) -> None:
        if self._thread is None:
            return
        self._q.put(None)
        self._thread.join(timeout=5)
        self._thread = None

    # ---------------- 调用 ----------------
    def call(self, fn: Callable[["ComBridge"], Any], *, timeout: float | None = None, label: str = "com"):
        """在 COM 线程上执行 fn(bridge)，返回其结果。"""
        if pythoncom is None:  # pragma: no cover
            raise CdrComFailure("当前环境不是 Windows，或未安装 pywin32")
        self.start()
        box: dict[str, Any] = {}
        done = threading.Event()
        self._q.put((fn, box, done))
        limit = self.default_timeout if timeout is None else timeout
        if not done.wait(limit):
            raise CdrTimeout(f"{label} 超过 {limit:.0f}s 未返回")
        if "error" in box:
            raise _wrap_com_error(box["error"])
        return box.get("value")

    # ---------------- 应用连接 ----------------
    @property
    def app(self) -> Any:
        return self._app

    def _find_running(self) -> Any:
        if win32 is None:  # pragma: no cover
            raise CdrNotConnected("当前非 Windows 环境")
        for progid in PROGIDS:
            try:
                return win32.GetActiveObject(progid)
            except Exception:
                continue
        return None

    def find_running(self) -> Any:
        """公开版本：查询当前是否已有 CorelDRAW 实例在运行（不启动）。"""
        return self._find_running()

    def connect(self, *, launch: bool = False, startup_timeout: float | None = None) -> dict:
        """附着已运行的 CorelDRAW；launch=True 时才允许启动新实例。"""

        def _do(bridge: "ComBridge"):
            app = self._find_running()
            launched = False
            if app is None:
                if not launch:
                    raise CdrNotConnected("没有找到运行中的 CorelDRAW")
                last: BaseException | None = None
                for progid in PROGIDS:
                    try:
                        app = win32.Dispatch(progid)
                        launched = True
                        break
                    except BaseException as exc:  # noqa: BLE001
                        last = exc
                        continue
                if app is None:
                    raise CdrNotConnected(f"无法启动 CorelDRAW：{last}")
                try:
                    app.Visible = True
                except Exception:
                    pass
            bridge._app = app
            return {
                "launched": launched,
                "version": _safe(lambda: str(app.Version)),
                "version_major": _safe(lambda: int(app.VersionMajor)),
                "version_build": _safe(lambda: int(app.VersionBuild)),
                "path": _safe(lambda: str(app.Path)),
                "documents": _safe(lambda: int(app.Documents.Count)),
            }

        return self.call(_do, timeout=startup_timeout or 300.0, label="connect")

    def ensure_app(self) -> Any:
        if self._app is None:
            raise CdrNotConnected("尚未连接 CorelDRAW")
        return self._app

    def reset(self) -> None:
        """连接失效时清空缓存的 Application 对象。"""
        self._app = None


def _safe(fn: Callable[[], Any], default: Any = None) -> Any:
    try:
        return fn()
    except Exception:
        return default


# ---- 进程级单例 ----
_BRIDGE: ComBridge | None = None
_LOCK = threading.Lock()


def get_bridge(timeout: float = DEFAULT_TIMEOUT_S) -> ComBridge:
    global _BRIDGE
    with _LOCK:
        if _BRIDGE is None:
            _BRIDGE = ComBridge(timeout=timeout)
        return _BRIDGE


def reset_bridge() -> None:
    global _BRIDGE
    with _LOCK:
        if _BRIDGE is not None:
            _BRIDGE.shutdown()
        _BRIDGE = None
