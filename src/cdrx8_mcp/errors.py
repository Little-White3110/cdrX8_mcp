"""异常类型：所有对外错误都收敛成 CdrError 的子类，便于 MCP 层转成结构化错误。"""

from __future__ import annotations


class CdrError(Exception):
    """所有本服务错误的基类。code 会原样返回给调用方 Agent。"""

    code = "cdr_error"
    hint = ""

    def as_dict(self) -> dict:
        d = {"error": self.code, "message": str(self)}
        if self.hint:
            d["hint"] = self.hint
        return d


class CdrTimeout(CdrError):
    code = "timeout"
    hint = (
        "CorelDRAW 可能卡在模态弹窗（缺字体/色彩配置）。COM 调用无法中途取消，"
        "超时只代表本服务不再等待；请人工检查 CorelDRAW 窗口，必要时结束进程后重连。"
    )


class CdrNotConnected(CdrError):
    code = "not_connected"
    hint = "先调用 cdr_status 查看状态，再用 cdr_connect(launch=true) 启动或附着 CorelDRAW。"


class CdrLaunchForbidden(CdrError):
    code = "launch_forbidden"
    hint = "启动 CorelDRAW 属于有副作用的操作，需显式传 launch=true（或设置 CDR_ALLOW_LAUNCH=1）。"


class CdrDocumentNotFound(CdrError):
    code = "document_not_found"
    hint = "调用 cdr_documents 获取当前打开的文档列表与 document_id。"


class CdrStaleView(CdrError):
    code = "stale_view"
    hint = "文档在你上次 cdr_inspect 之后发生了变化。重新调用 cdr_inspect 获取新的 revision 与 snapshot_token。"


class CdrRevisionMismatch(CdrStaleView):
    code = "revision_mismatch"


class CdrPathDenied(CdrError):
    code = "path_denied"
    hint = "导出/保存路径必须位于 CDR_ALLOWED_ROOTS 允许的目录内。"


class CdrFileExists(CdrError):
    code = "file_exists"
    hint = "目标文件已存在。确认要覆盖时传 overwrite=true。"


class CdrBadArgument(CdrError):
    code = "bad_argument"


class CdrComFailure(CdrError):
    code = "com_error"

    def __init__(self, message: str, hresult: int | None = None):
        super().__init__(message)
        self.hresult = hresult

    def as_dict(self) -> dict:
        d = super().as_dict()
        if self.hresult is not None:
            d["hresult"] = self.hresult
        return d
