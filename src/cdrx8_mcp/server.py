"""CorelDRAW X8 的 MCP 服务器（stdio）。

工具面刻意收敛成 12 个语义化操作，而不是把 800 多个 COM 接口直接暴露给模型——
模型在几千个方法里会迷路，安全审计也无从下手。
"""

from __future__ import annotations

import os
from typing import Any

from mcp.server.fastmcp import FastMCP

from . import ops
from .com import get_bridge
from .constants import (
    DEFAULT_EXPORT_TIMEOUT_S,
    DEFAULT_STARTUP_TIMEOUT_S,
    DEFAULT_TIMEOUT_S,
)
from .errors import (
    CdrComFailure,
    CdrError,
    CdrLaunchForbidden,
    CdrStaleView,
)
from .safety import env_flag, env_float, resolve_read_path, resolve_write_path
from .session import registry, snapshot_token

mcp = FastMCP("coreldraw-x8")


# --------------------------------------------------------------------------- #
# 基础封装
# --------------------------------------------------------------------------- #
def _bridge():
    return get_bridge(timeout=env_float("CDR_TIMEOUT", DEFAULT_TIMEOUT_S))


def _com(fn, *, timeout: float | None = None, label: str = "com"):
    return _bridge().call(fn, timeout=timeout, label=label)


def _guard(fn, *, extra: dict | None = None) -> dict:
    """统一错误出口：所有 CdrError 都转成结构化结果，不让异常穿透 MCP 层。"""
    try:
        result = fn()
        if extra:
            result = {**extra, **result}
        return result
    except CdrError as exc:
        return exc.as_dict()
    except Exception as exc:  # noqa: BLE001 - 兜底，避免把栈打给模型
        return {"error": "internal_error", "message": f"{type(exc).__name__}: {exc}"}


def _require_app(launch: bool = False):
    bridge = _bridge()
    if bridge.app is not None:
        return bridge.app
    if launch and not env_flag("CDR_ALLOW_LAUNCH", True):
        raise CdrLaunchForbidden()
    return None


def _doc_state(doc: Any, sess) -> dict:
    """给调用方的"当前视图"信息：revision + 快照令牌 + 结构规模。"""
    pages = ops._n(lambda: int(doc.Pages.Count), 0)
    shapes = ops.total_shapes(doc)
    token = snapshot_token(sess.name, sess.revision, pages, shapes)
    return {
        "document_id": sess.sid,
        "document_name": sess.name,
        "revision": sess.revision,
        "snapshot_token": token,
        "pages": pages,
        "shapes": shapes,
    }


def _check_view(sess, expected_revision: int | None, token: str | None, doc: Any) -> list[str]:
    """乐观锁校验。未传则放行但给出 warning；CDR_STRICT_REVISION=1 时强制。"""
    warnings: list[str] = []
    strict = env_flag("CDR_STRICT_REVISION", False)

    if expected_revision is None or token is None:
        msg = "未提供 expected_revision/snapshot_token，本次跳过了并发校验"
        if strict:
            raise CdrStaleView(msg + "；当前为严格模式（CDR_STRICT_REVISION=1），请先调用 cdr_inspect")
        warnings.append(msg)
        return warnings

    # 注意：这里必须回到 COM 线程读文档，不能在调用线程直接摸对象
    current = _com(lambda _b: _doc_state(doc, sess), label="view_check")
    if int(expected_revision) != current["revision"] or token != current["snapshot_token"]:
        raise CdrStaleView(
            "文档已变化：期望 revision={} / token={}，实际 revision={} / token={}。"
            "请重新调用 cdr_inspect。".format(
                expected_revision, token, current["revision"], current["snapshot_token"]
            )
        )
    return warnings


def _with_doc(document_id: str):
    """在 COM 线程上解析出文档对象，返回 (doc, sess)。"""
    sess = registry().by_sid(document_id)

    def _do(_bridge):
        app = _bridge.ensure_app()
        return ops.find_document(app, sess.name)

    return _com(_do, label="find_document"), sess


# --------------------------------------------------------------------------- #
# 工具
# --------------------------------------------------------------------------- #
@mcp.tool()
def cdr_status() -> dict:
    """查看 CorelDRAW 安装/运行状态与服务配置。不产生任何副作用，可随时调用。"""

    def _run():
        info: dict = {
            "platform": os.name,
            "allowed_roots": os.environ.get("CDR_ALLOWED_ROOTS", "") or None,
            "strict_revision": env_flag("CDR_STRICT_REVISION", False),
            "timeout_s": env_float("CDR_TIMEOUT", DEFAULT_TIMEOUT_S),
        }
        try:
            bridge = _bridge()

            def _probe(b):
                found = b.find_running()
                if found is None:
                    return {"running": False}
                return {
                    "running": True,
                    "version": str(found.Version),
                    "version_major": int(found.VersionMajor),
                    "version_build": int(found.VersionBuild),
                    "path": str(found.Path),
                    "documents": int(found.Documents.Count),
                }

            info.update(_com(_probe, timeout=30.0, label="status"))
        except CdrError as exc:
            info.update({"running": False, "probe_error": exc.message if hasattr(exc, "message") else str(exc)})
        except Exception as exc:  # noqa: BLE001
            info.update({"running": False, "probe_error": f"{type(exc).__name__}: {exc}"})
        return info

    return _guard(_run)


@mcp.tool()
def cdr_connect(launch: bool = False) -> dict:
    """连接 CorelDRAW。默认只**附着**已运行的实例；只有 launch=true 才会启动新进程。

    Args:
        launch: 允许启动 CorelDRAW（未运行时）。属于有副作用的操作。
    """
    def _run():
        result = _bridge().connect(launch=bool(launch), startup_timeout=DEFAULT_STARTUP_TIMEOUT_S)
        result["hint"] = "接下来调用 cdr_documents 取 document_id，或 cdr_open 新建/打开文档"
        return result

    return _guard(_run)


@mcp.tool()
def cdr_documents() -> dict:
    """列出当前打开的文档，并给出每个文档的 document_id / revision / snapshot_token。"""
    def _run():
        def _do(_b):
            app = _b.ensure_app()
            return ops.list_documents(app)

        docs = _com(_do, label="cdr_documents")
        out = []
        for d in docs:
            sess = registry().get_or_create(d["name"])
            out.append({**d, "document_id": sess.sid})
        return {"count": len(out), "documents": out,
                "hint": "用 cdr_inspect 取 revision 与 snapshot_token 后再修改文档"}

    return _guard(_run)


@mcp.tool()
def cdr_open(path: str | None = None, width_mm: float | None = None,
             height_mm: float | None = None, orientation: str | None = None) -> dict:
    """打开已有 .cdr（给 path），或新建空白文档（给 width_mm/height_mm）。

    Args:
        path: 要打开的 .cdr 绝对路径；为空则新建文档。
        width_mm: 新建文档宽度（毫米）。
        height_mm: 新建文档高度（毫米）。
        orientation: "portrait" 或 "landscape"。
    """
    def _run():
        target = None
        if path:
            target = str(resolve_read_path(path))

        def _do(_b):
            app = _b.ensure_app()
            if target:
                doc = ops.open_document(app, target)
            else:
                doc = ops.create_document(app, width_mm, height_mm, orientation)
            name = str(doc.Name)
            with ops.mm_units(doc):
                page = doc.ActivePage
                return {
                    "name": name,
                    "file_name": ops._s(lambda: str(doc.FileName)) or None,
                    "width_mm": ops._round(ops._s(lambda: float(page.SizeWidth))),
                    "height_mm": ops._round(ops._s(lambda: float(page.SizeHeight))),
                    "pages": ops._n(lambda: int(doc.Pages.Count), 1),
                }

        info = _com(_do, timeout=DEFAULT_STARTUP_TIMEOUT_S, label="cdr_open")
        sess = registry().get_or_create(info["name"])
        return {**info, "document_id": sess.sid, "revision": sess.revision,
                "hint": "调用 cdr_inspect 获取 snapshot_token"}

    return _guard(_run)


@mcp.tool()
def cdr_inspect(document_id: str, page_index: int | None = None,
                include: str = "shapes") -> dict:
    """读取文档当前状态：页面、revision、snapshot_token，以及页面上每个对象的 bbox 与文字。

    这是**修改前必须调用**的视图接口——它返回的 revision + snapshot_token 要原样传给
    cdr_create / cdr_modify / cdr_replace_text / cdr_pages / cdr_undo。

    Args:
        document_id: 来自 cdr_documents / cdr_open。
        page_index: 只看某一页（1 起）；不传则看活动页。
        include: "shapes"（默认，含对象列表）或 "pages"（只看页面概览）。
    """
    def _run():
        doc, sess = _with_doc(document_id)

        def _do(_b):
            state = _doc_state(doc, sess)
            result = {"state": state, "unit": "mm", "origin": "页面左下角，Y 轴向上"}
            result["pages"] = ops.page_summary(doc)
            if include != "pages":
                result["shapes"] = ops.list_shapes(doc, page_index)
                result["active_page"] = ops._n(lambda: int(doc.ActivePage.Index), 1)
            return result

        return _com(_do, label="cdr_inspect")

    return _guard(_run)


@mcp.tool()
def cdr_create(document_id: str, items: list[dict],
               expected_revision: int | None = None, snapshot_token: str | None = None,
               page_index: int | None = None) -> dict:
    """批量创建图元，整批合成**一个 undo 步**。

    支持的类型（全部在 CorelDRAW X8 上实测可用）：
      * {"type":"rectangle","x":10,"y":10,"width":50,"height":30,"fill":"#CC0000","outline":{"width":0.5}}
      * {"type":"ellipse","cx":80,"cy":40,"rx":15,"ry":10,"fill":{"c":0,"m":100,"y":100,"k":0}}
      * {"type":"text","x":10,"y":60,"text":"标题","font":"Arial","size":24,"fill":"#000000"}
      * {"type":"paragraph_text","x":10,"y":10,"x2":90,"y2":50,"text":"段落文字","size":10}

    坐标单位毫米、原点在页面**左下角**、Y 轴向上；矩形是「左下角 + 宽高」，
    椭圆是「圆心 + 半径」。
    """
    def _run():
        if not items:
            return {"error": "bad_argument", "message": "items 不能为空"}
        doc, sess = _with_doc(document_id)
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)

        created = _com(lambda _b: ops.create_items(doc, page_index, items), label="cdr_create")
        sess.bump()
        state = _com(lambda _b: _doc_state(doc, sess), label="state")
        return {"created": created, "count": len(created), "state": state, "warnings": warnings}

    return _guard(_run)


@mcp.tool()
def cdr_modify(document_id: str, target: dict, changes: dict,
               expected_revision: int | None = None, snapshot_token: str | None = None,
               page_index: int | None = None) -> dict:
    """修改单个对象。定位用 target，改动用 changes。

    target: {"index":1} 或 {"name":"cdr_ab12cd34"} 或 {"selection":true}
    changes 支持的键：
      dx/dy（相对移动）、x/y（绝对左下角坐标）、width/height（尺寸）、rotate（角度）、
      fill（"#RRGGBB" 或 {c,m,y,k} 或 {r,g,b}）、outline（数字宽度或 {width,color}）、
      text（改文字内容）、font/size/bold/italic、name（重命名）、
      to_front/to_back、duplicate、delete
    """
    def _run():
        if not isinstance(changes, dict) or not changes:
            return {"error": "bad_argument", "message": "changes 不能为空"}
        doc, sess = _with_doc(document_id)
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)

        result = _com(lambda _b: ops.modify_shape(doc, page_index, target, changes), label="cdr_modify")
        sess.bump()
        state = _com(lambda _b: _doc_state(doc, sess), label="state")
        result.setdefault("warnings", [])
        result["warnings"] = warnings + result["warnings"]
        result["state"] = state
        return result

    return _guard(_run)


@mcp.tool()
def cdr_replace_text(document_id: str, find: str, replace: str,
                     expected_revision: int | None = None, snapshot_token: str | None = None,
                     all_pages: bool = True, case_sensitive: bool = True) -> dict:
    """全文档文字替换（换版/改年份/换客户名的高频操作），整批合成一个 undo 步。

    Args:
        find: 要查找的文本。
        replace: 替换成什么（可以是空串，表示删除）。
        all_pages: true=所有页；false=仅活动页。
        case_sensitive: 是否区分大小写。
    """
    def _run():
        doc, sess = _with_doc(document_id)
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)

        result = _com(lambda _b: ops.replace_text(
            doc, find, replace, all_pages=all_pages, case_sensitive=case_sensitive
        ), label="cdr_replace_text")
        sess.bump()
        result["state"] = _com(lambda _b: _doc_state(doc, sess), label="state")
        result["warnings"] = warnings
        return result

    return _guard(_run)


@mcp.tool()
def cdr_pages(document_id: str, action: str = "list",
              page_index: int | None = None, count: int = 1,
              width_mm: float | None = None, height_mm: float | None = None,
              expected_revision: int | None = None,
              snapshot_token: str | None = None) -> dict:
    """页面管理。

    Args:
        action: list（列表）/ add（新增 count 页）/ activate（切到 page_index）/ delete（删 page_index）/ size（设 page_index 的尺寸）
        page_index: 目标页号（1 起）。
        count: action=add 时新增几页。
        width_mm / height_mm: action=size 时的新尺寸。
    """
    def _run():
        doc, sess = _with_doc(document_id)
        act = (action or "list").strip().lower()

        if act == "list":
            return _com(lambda _b: {"pages": ops.page_summary(doc),
                                    "active": ops._n(lambda: int(doc.ActivePage.Index), 1)},
                        label="cdr_pages")

        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        if act == "add":
            result = _com(lambda _b: ops.add_pages(doc, count), label="cdr_pages:add")
        elif act == "activate":
            if page_index is None:
                return {"error": "bad_argument", "message": "action=activate 需要 page_index"}
            result = _com(lambda _b: ops.activate_page(doc, page_index), label="cdr_pages:activate")
        elif act == "delete":
            if page_index is None:
                return {"error": "bad_argument", "message": "action=delete 需要 page_index"}
            result = _com(lambda _b: ops.delete_page(doc, page_index), label="cdr_pages:delete")
        elif act == "size":
            if width_mm is None or height_mm is None:
                return {"error": "bad_argument", "message": "action=size 需要 width_mm 与 height_mm"}
            result = _com(lambda _b: ops.set_page_size(doc, width_mm, height_mm, page_index),
                          label="cdr_pages:size")
        else:
            return {"error": "bad_argument",
                    "message": f"未知 action：{act}；可用 list/add/activate/delete/size"}

        sess.bump()
        result["state"] = _com(lambda _b: _doc_state(doc, sess), label="state")
        result["warnings"] = warnings
        return result

    return _guard(_run)


@mcp.tool()
def cdr_export(document_id: str, path: str, format: str = "png",
               overwrite: bool = False) -> dict:
    """导出当前页为 png / svg / pdf（绝对路径）。

    Args:
        path: 目标文件绝对路径。
        format: png / svg / pdf。
        overwrite: 目标已存在时是否覆盖，默认 false。
    """
    def _run():
        fmt = (format or "").lower().lstrip(".")
        if fmt not in {"png", "svg", "pdf"}:
            return {"error": "bad_argument", "message": "format 只支持 png / svg / pdf"}
        p, warn = resolve_write_path(path, overwrite=overwrite)
        doc, _sess = _with_doc(document_id)
        result = _com(lambda _b: ops.export(doc, fmt, str(p)),
                      timeout=DEFAULT_EXPORT_TIMEOUT_S, label="cdr_export")
        result["bytes"] = ops._s(lambda: os.path.getsize(p))
        result["exists"] = p.exists()
        if warn:
            result["warnings"] = [warn]
        return result

    return _guard(_run)


@mcp.tool()
def cdr_save(document_id: str, path: str, overwrite: bool = False) -> dict:
    """另存为 .cdr（**不会**覆盖源文件，除非你显式把 path 指向它并 overwrite=true）。

    Args:
        path: 目标 .cdr 绝对路径。
        overwrite: 目标已存在时是否覆盖，默认 false。
    """
    def _run():
        p, warn = resolve_write_path(path, overwrite=overwrite)
        if p.suffix.lower() != ".cdr":
            return {"error": "bad_argument", "message": "保存路径必须以 .cdr 结尾"}
        doc, sess = _with_doc(document_id)
        info = _com(lambda _b: ops.save_as(doc, str(p)), label="cdr_save")
        registry().rename(sess.name, info["name"])
        sess.bump()
        result = {"saved": str(p), "bytes": ops._s(lambda: os.path.getsize(p)),
                  "state": _com(lambda _b: _doc_state(doc, sess), label="state")}
        if warn:
            result["warnings"] = [warn]
        return result

    return _guard(_run)


@mcp.tool()
def cdr_undo(document_id: str, steps: int = 1) -> dict:
    """撤销最近 N 步（一步 = 一次工具调用，因为每次变更都做成了原子 undo 组）。"""
    def _run():
        doc, sess = _with_doc(document_id)
        result = _com(lambda _b: ops.undo(doc, steps), label="cdr_undo")
        sess.bump()
        result["state"] = _com(lambda _b: _doc_state(doc, sess), label="state")
        return result

    return _guard(_run)


@mcp.tool()
def cdr_close(document_id: str, save: bool = False) -> dict:
    """关闭文档。默认**丢弃**未保存改动（用于清理测试文档 / 结束会话）。

    Args:
        save: true 时先保存再关闭（仅当文档已有文件路径；新文档请先 cdr_save）。
    """
    def _run():
        doc, sess = _with_doc(document_id)
        result = _com(lambda _b: ops.close_document(doc, save=bool(save)), label="cdr_close")
        registry().drop(sess.name)
        return result

    return _guard(_run)


# --------------------------------------------------------------------------- #
# 入口
# --------------------------------------------------------------------------- #
def main() -> None:
    """console_scripts 入口。默认 stdio；设置 CDR_TRANSPORT=http 可起本地 HTTP。"""
    transport = os.environ.get("CDR_TRANSPORT", "stdio").lower()
    if transport in {"http", "streamable-http", "sse"}:
        mcp.settings.host = os.environ.get("CDR_HOST", "127.0.0.1")
        mcp.settings.port = int(os.environ.get("CDR_PORT", "8765"))
        mcp.run(transport="streamable-http")
    else:
        mcp.run()


if __name__ == "__main__":
    main()
