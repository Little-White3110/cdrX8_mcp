"""CorelDRAW X8 的 MCP 服务器（stdio / 可选 HTTP）。

工具面是 20 个**语义化**操作，而不是把 6800 多个 COM 方法直接暴露给模型——
模型在几千个方法里会迷路，安全审计也无从下手。
每个工具背后能做的事，都是先反解 vgcoreauto.tlb、再在 X8 真机上实测通过才写进来的。
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
    SHAPE_TYPES,
    TEXT_TYPE,
)
from .errors import CdrComFailure, CdrError, CdrLaunchForbidden, CdrStaleView
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
    try:
        result = fn()
        if extra:
            result = {**extra, **result}
        return result
    except CdrError as exc:
        return exc.as_dict()
    except Exception as exc:  # noqa: BLE001
        return {"error": "internal_error", "message": f"{type(exc).__name__}: {exc}"}


def _doc_state(doc: Any, sess) -> dict:
    pages = ops._n(lambda: int(doc.Pages.Count), 0)
    shapes = ops.total_shapes(doc)
    return {
        "document_id": sess.sid,
        "document_name": sess.name,
        "revision": sess.revision,
        "snapshot_token": snapshot_token(sess.name, sess.revision, pages, shapes),
        "pages": pages,
        "shapes": shapes,
    }


def _check_view(sess, expected_revision: int | None, token: str | None, doc: Any) -> list[str]:
    warnings: list[str] = []
    strict = env_flag("CDR_STRICT_REVISION", False)
    if expected_revision is None or token is None:
        msg = "未提供 expected_revision/snapshot_token，本次跳过了并发校验"
        if strict:
            raise CdrStaleView(msg + "；当前为严格模式（CDR_STRICT_REVISION=1），请先调用 cdr_inspect")
        warnings.append(msg)
        return warnings
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
    sess = registry().by_sid(document_id)

    def _do(bridge):
        app = bridge.ensure_app()
        return ops.find_document(app, sess.name)

    return _com(_do, label="find_document"), sess


def _commit(doc: Any, sess, result: dict, warnings: list[str] | None = None) -> dict:
    """统一收尾：推进 revision、回填最新状态与 warning（合并而非覆盖）。"""
    sess.bump()
    result["state"] = _com(lambda _b: _doc_state(doc, sess), label="state")
    merged = list(result.get("warnings") or [])
    for w in warnings or []:
        if w not in merged:
            merged.append(w)
    if merged:
        result["warnings"] = merged
    return result


# --------------------------------------------------------------------------- #
# 状态 / 连接 / 文档
# --------------------------------------------------------------------------- #
@mcp.tool()
def cdr_status() -> dict:
    """查看 CorelDRAW 安装/运行状态与服务配置。无副作用，可随时调用。"""

    def _run():
        info: dict = {
            "platform": os.name,
            "allowed_roots": os.environ.get("CDR_ALLOWED_ROOTS", "") or None,
            "strict_revision": env_flag("CDR_STRICT_REVISION", False),
            "timeout_s": env_float("CDR_TIMEOUT", DEFAULT_TIMEOUT_S),
        }
        try:
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
            info.update({"running": False, "probe_error": str(exc)})
        except Exception as exc:  # noqa: BLE001
            info.update({"running": False, "probe_error": f"{type(exc).__name__}: {exc}"})
        info["tools"] = 25
        return info

    return _guard(_run)


@mcp.tool()
def cdr_connect(launch: bool = False) -> dict:
    """连接 CorelDRAW。默认只**附着**已运行实例；launch=true 才允许启动新进程。"""
    def _run():
        result = _bridge().connect(launch=bool(launch),
                                   startup_timeout=DEFAULT_STARTUP_TIMEOUT_S)
        result["hint"] = "接着用 cdr_documents 取 document_id，或 cdr_open 新建/打开文档"
        return result

    return _guard(_run)


@mcp.tool()
def cdr_documents() -> dict:
    """列出当前打开的文档，给出 document_id / revision / snapshot_token。"""
    def _run():
        docs = _com(lambda _b: ops.list_documents(_b.ensure_app()), label="cdr_documents")
        out = []
        for d in docs:
            sess = registry().get_or_create(d["name"])
            out.append({**d, "document_id": sess.sid})
        return {"count": len(out), "documents": out,
                "hint": "用 cdr_inspect 取 revision 与 snapshot_token 后再修改"}

    return _guard(_run)


@mcp.tool()
def cdr_open(path: str | None = None, width_mm: float | None = None,
             height_mm: float | None = None, orientation: str | None = None,
             preset: str | None = None) -> dict:
    """打开已有 .cdr，或新建空白文档。

    Args:
        path: 要打开的 .cdr 绝对路径；为空则新建。
        width_mm / height_mm: 新建文档尺寸（毫米）。
        orientation: "portrait" / "landscape"。
        preset: 页面预设，如 "A4"/"A3"/"card"（名片）；给了就忽略 width/height。
    """
    def _run():
        target = str(resolve_read_path(path)) if path else None

        def _do(_b):
            app = _b.ensure_app()
            doc = (ops.open_document(app, target) if target
                   else ops.create_document(app, width_mm, height_mm, orientation, preset))
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
    """读取文档当前状态：页面、图层、revision、snapshot_token、每个对象的 bbox 与文字。

    修改前**必须**先调用它，把返回的 revision + snapshot_token 原样传给修改类工具。

    Args:
        document_id: 来自 cdr_documents / cdr_open。
        page_index: 只看某一页（1 起）；不传则看活动页。
        include: "shapes"（含对象列表）或 "pages"（只看页面/图层概览）。
    """
    def _run():
        doc, sess = _with_doc(document_id)

        def _do(_b):
            state = _doc_state(doc, sess)
            result = {"state": state, "unit": "mm", "origin": "页面左下角，Y 轴向上"}
            result["pages"] = ops.page_summary(doc)
            result["layers"] = ops.list_layers(doc, page_index)
            if include != "pages":
                result["shapes"] = ops.list_shapes(doc, page_index)
                result["active_page"] = ops._n(lambda: int(doc.ActivePage.Index), 1)
            return result

        return _com(_do, label="cdr_inspect")

    return _guard(_run)


# --------------------------------------------------------------------------- #
# 创建 / 修改
# --------------------------------------------------------------------------- #
@mcp.tool()
def cdr_create(document_id: str, items: list[dict],
               expected_revision: int | None = None, snapshot_token: str | None = None,
               page_index: int | None = None) -> dict:
    """批量创建图元，整批合成**一个 undo 步**。全部坐标毫米、原点页面左下角、Y 向上。

    可用类型与关键参数：
      * rectangle: x,y,width,height[,radius / radius_ul/ur/lr/ll]
      * ellipse:   cx,cy,rx,ry[,start_angle,end_angle,pie]（给角度就是圆弧/扇形）
      * polygon:   cx,cy,radius,sides[,angle]；star: 同上 + inner_radius[,sharpness -1..1]
      * spiral:    left,top,right,bottom[,revolutions,spiral_type,growth]
      * grid:      left,top,right,bottom,cols,rows
      * line:      x1,y1,x2,y2
      * bezier:    x1,y1,c1x,c1y,c2x,c2y,x2,y2
      * polyline:  points:[[x,y],…][,closed] —— 任意多边形/折线（逐点连线）
      * text:      x,y,text[,font,size,bold,italic,underline,align]
      * paragraph_text: x,y,x2,y2,text[,font,size,align]
      * guide:     x1,y1,x2,y2（参考线，不是形状）

    每个图元还可带样式：fill（"#RRGGBB" / {c,m,y,k} / {r,g,b} /
    {type:"fountain",start,end,gradient_type,angle}）、
    outline（数字宽度或 {width,color,caps,join}）、drop_shadow（{opacity,offset_x,offset_y,color}）。
    """
    def _run():
        if not items:
            return {"error": "bad_argument", "message": "items 不能为空"}
        doc, sess = _with_doc(document_id)
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        created = _com(lambda _b: ops.create_items(doc, page_index, items), label="cdr_create")
        return _commit(doc, sess, {"created": created, "count": len(created)}, warnings)

    return _guard(_run)


@mcp.tool()
def cdr_modify(document_id: str, target: dict, changes: dict,
               expected_revision: int | None = None, snapshot_token: str | None = None,
               page_index: int | None = None) -> dict:
    """修改对象。target 定位，changes 描述改动。

    **单个** target:{"index":1} / {"name":"cdr_ab12"} / {"selection":true}
    **批量** target:{"all":true} 或 {"type_id":1}（本页所有 / 按类型）—— 一批一个 undo 步

    changes 可用键：
      dx,dy（相对移动）｜x,y（绝对左下角）｜width,height[,keep_aspect]｜rotate
      affine [d11,d12,d21,d22,cx,cy]｜align_to {mode:"page_center"|"point"|"grid", x,y,type}
      fill（纯色/渐变/纹理/"none"）｜fill_from（复制另一对象的填充：{"name":"..."}）
      outline（宽度或 {width,color,caps,join,behind_fill,dash,start_arrow,end_arrow}）
      outline_from（复制另一对象的描边：{"name":"..."}）
      drop_shadow {opacity,offset_x,offset_y,color}｜contour {offset,steps}｜envelope 1..39
      transparency（0~100 均匀透明度 / {type:"fountain",start,end,angle} / "none"）
      gradient {angle, edge_pad, translate:[dx,dy], transformations:[d11,d12,d21,d22], opaque}
      effect {type:"bci"|"gamma"|"hsl"|"invert"|"posterize"|"color_balance", ...}
      flip（0 水平 / 1 垂直）｜skew {x,y}｜clone {dx,dy}｜locked true/false
      text｜font｜size｜bold｜italic｜underline｜alignment
      fit_to_path（{"name":"曲线对象"}）｜fit_frame（文本适应文本框）
      name｜to_front｜to_back｜duplicate[,dup_dx,dup_dy]｜delete

    颜色额外支持 {"gray":50} / {"lab":[l,a,b]} / {"hls":[h,l,s]} / "registration"。
    注：outline 的 dash / start_arrow / end_arrow 传的是 CorelDRAW 线条样式库与
    箭头库的**索引**（实测 Outline.Style=数字会类型不匹配，必须赋对象）。
    """
    def _run():
        if not isinstance(changes, dict) or not changes:
            return {"error": "bad_argument", "message": "changes 不能为空"}
        doc, sess = _with_doc(document_id)
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        # 批量选择器：{"all":true} / {"type_id":n} / {"name":...} / {"selection":true}
        if isinstance(target, dict) and (target.get("all") or "type_id" in target
                                         or target.get("selection")):
            result = _com(lambda _b: ops.batch_modify(doc, page_index, target, changes),
                          label="cdr_modify:batch")
            result.setdefault("warnings", [])
            return _commit(doc, sess, result, warnings + result["warnings"])
        result = _com(lambda _b: ops.modify_shape(doc, page_index, target, changes),
                      label="cdr_modify")
        result.setdefault("warnings", [])
        return _commit(doc, sess, result, warnings + result["warnings"])

    return _guard(_run)


@mcp.tool()
def cdr_text(document_id: str, target: dict, text: str | None = None,
             font: str | None = None, size: float | None = None, bold: bool | None = None,
             italic: bool | None = None, underline: bool | None = None,
             alignment: str | None = None, to_curves: bool = False,
             fit_to_path: dict | None = None, fit_frame: bool = False,
             change_case: str | None = None, bullet: str | None = None,
             drop_cap: int | None = None, no_effect: bool = False,
             line_spacing: float | None = None, char_spacing: float | None = None,
             word_spacing: float | None = None,
             insert_before: str | None = None, insert_after: str | None = None,
             to_paragraph: bool = False, to_artistic: bool = False,
             contents: str | None = None, find: str | None = None,
             replace_old: str | None = None, replace_new: str | None = None,
             select_all: bool = False, align_to_baseline: bool = False,
             copy_attributes_from: dict | None = None,
             expected_revision: int | None = None, snapshot_token: str | None = None,
             page_index: int | None = None) -> dict:
    """文字的排版级操作。

    内容与类型：text / contents（段落文本整体替换）/ find / replace_old+replace_new /
                to_paragraph / to_artistic（艺术字↔段落字互转）
    样式：font / size / bold / italic / underline / alignment /
          change_case（sentence·lower·upper·title·toggle）/ bullet / drop_cap /
          line_spacing / char_spacing / word_spacing
    文本区间：select_all / align_to_baseline / copy_attributes_from（复制另一对象的文字属性）
    版式：fit_to_path（沿曲线）/ fit_frame（自动适配文本框）/ insert_before / insert_after
    输出：to_curves（转曲，发排前必做）

    注意：contents / replace_old / find 走的是段落文本接口，对艺术字会返回 warning。
    """
    def _run():
        doc, sess = _with_doc(document_id)
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        result = _com(lambda _b: ops.set_text(
            doc, page_index, target, text=text, font=font, size=size, bold=bold,
            italic=italic, underline=underline, alignment=alignment, to_curves=to_curves,
            fit_to_path=fit_to_path, fit_frame=fit_frame, change_case=change_case,
            bullet=bullet, drop_cap=drop_cap, no_effect=no_effect,
            line_spacing=line_spacing, char_spacing=char_spacing,
            word_spacing=word_spacing, insert_before=insert_before,
            insert_after=insert_after, to_paragraph=to_paragraph, to_artistic=to_artistic,
            contents=contents, find=find, replace_old=replace_old,
            replace_new=replace_new, select_all=select_all,
            align_to_baseline=align_to_baseline,
            copy_attributes_from=copy_attributes_from,
        ), label="cdr_text")
        return _commit(doc, sess, result, warnings)

    return _guard(_run)


@mcp.tool()
def cdr_replace_text(document_id: str, find: str, replace: str,
                     expected_revision: int | None = None, snapshot_token: str | None = None,
                     all_pages: bool = True, case_sensitive: bool = True) -> dict:
    """全文档文字替换（换版/改年份/换客户名）。底层用 X8 原生 Page.TextReplace。"""
    def _run():
        doc, sess = _with_doc(document_id)
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        result = _com(lambda _b: ops.replace_text(
            doc, find, replace, all_pages=all_pages, case_sensitive=case_sensitive,
        ), label="cdr_replace_text")
        return _commit(doc, sess, result, warnings)

    return _guard(_run)


# --------------------------------------------------------------------------- #
# 组合 / 布尔 / 整形 / 对齐
# --------------------------------------------------------------------------- #
@mcp.tool()
def cdr_select(document_id: str, page_index: int | None = None, name: str | None = None,
               type_id: int | None = None, rect: dict | None = None,
               all_shapes: bool = False) -> dict:
    """把对象加入当前选择集，供 cdr_shape_ops(group/combine/ungroup) 使用。

    type_id: 1=矩形 2=椭圆 3=曲线 4=多边形 6=文字 7=网格
    rect: {"x1":0,"y1":0,"x2":100,"y2":50,"touch":true} 按矩形框选
    """
    def _run():
        doc, sess = _with_doc(document_id)

        def _do(_b):
            picked = ops.select_shapes(doc, page_index, name=name, type_id=type_id,
                                       rect=rect, all_shapes=all_shapes)
            return picked

        result = _com(_do, label="cdr_select")
        result["hint"] = "接着调用 cdr_shape_ops 做组合/布尔"
        return result

    return _guard(_run)


@mcp.tool()
def cdr_shape_ops(document_id: str, op: str, target: dict | None = None,
                  other: dict | None = None, radius: float = 0,
                  distance_a: float = 0, distance_b: float = 0,
                  leave_source: bool = True, leave_target: bool = True,
                  symbol_name: str | None = None,
                  expected_revision: int | None = None,
                  snapshot_token: str | None = None,
                  page_index: int | None = None) -> dict:
    """布尔运算与整形。

    op 取值：
      weld / trim / intersect  需要 target + other（两个对象）
      powerclip                把 target 放进 other 这个容器
      group / ungroup / combine 作用于**当前选择集**（ungroup 用 target 指定组合体）
      break_apart              打散 target
      to_curves                转曲
      outline_to_curves        描边转曲（对象原地从形状变成曲线轮廓，发排/切割常用）
      to_symbol                把 target 转成符号（需 symbol_name）
      fillet（radius）/ chamfer（distance_a,distance_b）/ scallop（radius）
    """
    def _run():
        doc, sess = _with_doc(document_id)
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        result = _com(lambda _b: ops.shape_op(
            doc, page_index, op, target=target, other=other, radius=radius,
            distance_a=distance_a, distance_b=distance_b,
            leave_source=leave_source, leave_target=leave_target,
            symbol_name=symbol_name,
        ), label="cdr_shape_ops")
        return _commit(doc, sess, result, warnings)

    return _guard(_run)


@mcp.tool()
def cdr_align(document_id: str, horizontal: str = "none", vertical: str = "none",
              to: str = "selection", name: str | None = None,
              type_id: int | None = None, rect: dict | None = None,
              expected_revision: int | None = None, snapshot_token: str | None = None,
              page_index: int | None = None) -> dict:
    """对齐与分布。

    horizontal: none / left / center / right / space（等间距）
    vertical:   none / bottom / middle / top / space
    to:         "selection"（相互对齐）或 "page"（对齐页面）
    先用 name / type_id / rect 选出对象；都不给就用当前选择集。
    """
    def _run():
        doc, sess = _with_doc(document_id)
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        result = _com(lambda _b: ops.align_shapes(
            doc, page_index, horizontal=horizontal, vertical=vertical, to=to,
            name=name, type_id=type_id, rect=rect,
        ), label="cdr_align")
        return _commit(doc, sess, result, warnings)

    return _guard(_run)


# --------------------------------------------------------------------------- #
# 图层 / 查找 / 页面
# --------------------------------------------------------------------------- #
@mcp.tool()
def cdr_layers(document_id: str, action: str = "list",
               name: str | None = None, layer_index: int | None = None,
               page_index: int | None = None, visible: bool | None = None,
               editable: bool | None = None, printable: bool | None = None,
               move_above: str | None = None,
               expected_revision: int | None = None,
               snapshot_token: str | None = None) -> dict:
    """图层管理。

    action: list / create（给 name）/ activate / set（visible/editable/printable）/ move（move_above=目标图层名）/ delete
    layer_index 或 name 用来定位图层；都不给则作用于活动图层。
    """
    def _run():
        doc, sess = _with_doc(document_id)
        act = (action or "list").lower()
        if act == "list":
            return _com(lambda _b: {"layers": ops.list_layers(doc, page_index)},
                        label="cdr_layers")
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        result = _com(lambda _b: ops.layer_action(
            doc, act, name=name, page_index=page_index, index=layer_index,
            visible=visible, editable=editable, printable=printable, move=move_above,
        ), label="cdr_layers")
        return _commit(doc, sess, result, warnings)

    return _guard(_run)


@mcp.tool()
def cdr_find(document_id: str, name: str | None = None, type_id: int | None = None,
             page_index: int | None = None) -> dict:
    """按名字或类型查找对象（type_id: 1矩形 2椭圆 3曲线 4多边形 6文字 7组合体 9参考线）。"""
    def _run():
        doc, _sess = _with_doc(document_id)
        found = _com(lambda _b: ops.find_shapes(doc, name=name, type_id=type_id,
                                                page_index=page_index), label="cdr_find")
        return {"count": len(found), "shapes": found}

    return _guard(_run)


@mcp.tool()
def cdr_pages(document_id: str, action: str = "list",
              page_index: int | None = None, count: int = 1,
              width_mm: float | None = None, height_mm: float | None = None,
              preset: str | None = None, landscape: bool = False,
              x: float | None = None, y: float | None = None,
              text: str | None = None, move_to_index: int | None = None,
              hot_area: float = 2.0,
              expected_revision: int | None = None,
              snapshot_token: str | None = None) -> dict:
    """页面管理。

    action:
      list                          列表（含图层与对象数）
      add（count）                  新增页
      activate（page_index）        切页
      delete（page_index）          删页
      size（width_mm,height_mm）    自定义尺寸
      preset（preset="A4"/"card"）  用内置纸张尺寸
      guides                        列出参考线
      unlock                        解锁本页所有对象
      bounds                        读取本页所有对象的总包围盒
      geometry                      读页面尺寸与中心
      shape_at（x,y）               按坐标点找对象
      select_at（x,y）              按坐标点选择对象
      text_find（text）             在页面上查找文本
      move_to（move_to_index）      把当前页移到指定位置
    """
    def _run():
        doc, sess = _with_doc(document_id)
        act = (action or "list").strip().lower()

        if act == "list":
            return _com(lambda _b: {
                "pages": ops.page_summary(doc),
                "active": ops._n(lambda: int(doc.ActivePage.Index), 1),
                "presets": sorted(ops.PAGE_PRESETS),
            }, label="cdr_pages")
        if act == "guides":
            return _com(lambda _b: ops.list_guides(doc, page_index), label="cdr_pages:guides")
        if act == "bounds":
            return _com(lambda _b: ops.objects_bounds(doc, page_index), label="cdr_pages:bounds")
        if act in {"geometry", "shape_at", "select_at", "text_find", "move_to"}:
            return _com(lambda _b: ops.page_query(
                doc, act, page_index=page_index, x=x, y=y, text=text,
                to_index=move_to_index, hot_area=hot_area), label="cdr_pages:query")

        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        if act == "unlock":
            result = _com(lambda _b: ops.unlock_all(doc, page_index),
                          label="cdr_pages:unlock")
            return _commit(doc, sess, result, warnings)
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
        elif act == "preset":
            if not preset:
                return {"error": "bad_argument", "message": "action=preset 需要 preset"}
            result = _com(lambda _b: ops.set_page_preset(doc, preset, landscape, page_index),
                          label="cdr_pages:preset")
        else:
            return {"error": "bad_argument",
                    "message": f"未知 action：{act}；可用 list/add/activate/delete/size/"
                               f"preset/guides/unlock/bounds"}

        return _commit(doc, sess, result, warnings)

    return _guard(_run)


@mcp.tool()
def cdr_path(document_id: str, target: dict, action: str = "info",
             offset: float = 1.0, direction: int = 0,
             x: float | None = None, y: float | None = None,
             radius: float | None = None, rx: float | None = None, ry: float | None = None,
             rect: list | None = None, arc: list | None = None, bezier: list | None = None,
             other: dict | None = None, amount: float = 30.0,
             subpath: int = 1, t: float = 0.5, dx: float = 0.0, dy: float = 0.0,
             expected_revision: int | None = None,
             snapshot_token: str | None = None,
             page_index: int | None = None) -> dict:
    """曲线 / 子路径级操作（对象需是曲线，可先 `cdr_shape_ops(op="to_curves")` 转曲）。

    只读：info（长度/面积/子路径）、polyline、point_inside、intersects、
          subpath_info、get_point、get_tangent

    变更：contour（偏移出刀模线）、add_subpath、
          append_circle / append_ellipse / append_rectangle / append_arc / append_bezier
          （往同一条曲线里拼更多子路径）、
          reduce_nodes（节点精简）、self_weld（自焊接）、
          equal_divide（等分）、add_node、break_apart、move_subpath
    """
    def _run():
        act = (action or "info").lower()
        readonly = act in {"info", "polyline", "point_inside", "intersects",
                           "subpath_info", "get_point", "get_tangent"}
        doc, sess = _with_doc(document_id)
        kwargs = dict(offset=offset, direction=direction, x=x, y=y, radius=radius,
                      rx=rx, ry=ry, rect=rect, arc=arc, bezier=bezier, other=other,
                      amount=amount, subpath=subpath, t=t, dx=dx, dy=dy)
        if readonly:
            return _com(lambda _b: ops.path_ops(doc, page_index, target, act, **kwargs),
                        label="cdr_path")
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        result = _com(lambda _b: ops.path_ops(doc, page_index, target, act, **kwargs),
                      label="cdr_path")
        return _commit(doc, sess, result, warnings)

    return _guard(_run)


# --------------------------------------------------------------------------- #
# 导入 / 导出 / 保存
# --------------------------------------------------------------------------- #
@mcp.tool()
def cdr_color(action: str = "info", color: Any = None, color2: Any = None,
              ratio: float = 0.5) -> dict:
    """颜色运算：解析 / 转换色彩空间 / 两色混合。

    Args:
        action: info（解析并返回各通道）/ convert（转空间，color2 给目标空间名）/
                blend（按 ratio 混合 color 与 color2，0=全 color，1=全 color2）
        color: 任意颜色写法：#RRGGBB / {c,m,y,k} / {r,g,b} / {gray} / {lab} / {hls}
        color2: blend 的第二个颜色；convert 时传目标空间名（cmyk/rgb/gray/lab/hls/hsb/cmy）

    注意：这是**颜色对象**运算，不修改文档，所以不需要 document_id。
    """
    def _run():
        return ops.color_ops(None, action, color=color, color2=color2, ratio=ratio)

    return _guard(_run)


@mcp.tool()
def cdr_doc(document_id: str, action: str = "info", target: dict | None = None,
            expected_revision: int | None = None, snapshot_token: str | None = None,
            page_index: int | None = None) -> dict:
    """文档级操作。

    action:
      info             文档概况（路径、脏标记、页数、单位、对象总数）
      clear_selection  清空选择集
      clear_undo       清空撤销栈（批处理前后释放内存）
      palette          把文档用到的颜色登记进调色板（印前交接）
      arrowhead        用一条曲线生成自定义箭头（target 指向曲线）
    """
    def _run():
        doc, sess = _with_doc(document_id)
        if action == "info":
            return _com(lambda _b: ops.doc_ops(doc, "info", page_index=page_index),
                        label="cdr_doc")
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        result = _com(lambda _b: ops.doc_ops(doc, action, target=target,
                                             page_index=page_index), label="cdr_doc")
        if action == "clear_selection":
            return result
        return _commit(doc, sess, result, warnings)

    return _guard(_run)


@mcp.tool()
def cdr_view(document_id: str, action: str = "refresh",
             x: float | None = None, y: float | None = None,
             distance: float = 10.0) -> dict:
    """视图/窗口操作。

    action:
      refresh              强制重绘（导出前刷新、或改完看效果）
      activate             激活窗口
      to_screen（x,y）      文档坐标 -> 屏幕坐标
      to_document（x,y）    屏幕坐标 -> 文档坐标
      distance_to_screen（distance）  文档距离 -> 屏幕距离
    """
    def _run():
        doc, _sess = _with_doc(document_id)
        return _com(lambda _b: ops.view_ops(doc, action, x=x, y=y, distance=distance),
                    label="cdr_view")

    return _guard(_run)


@mcp.tool()
def cdr_import(document_id: str, path: str, center: bool = True,
               page_index: int | None = None,
               expected_revision: int | None = None,
               snapshot_token: str | None = None) -> dict:
    """导入图片或矢量文件（png / jpg / svg / pdf / ai / eps / dxf 等，取决于 X8 装了什么过滤器）。

    Args:
        path: 源文件绝对路径。
        center: 是否把导入对象居中放到页面上。
    """
    def _run():
        src = resolve_read_path(path)
        doc, sess = _with_doc(document_id)
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        result = _com(lambda _b: ops.import_file(doc, str(src), page_index, center),
                      timeout=DEFAULT_EXPORT_TIMEOUT_S, label="cdr_import")
        return _commit(doc, sess, result, warnings)

    return _guard(_run)


@mcp.tool()
def cdr_export(document_id: str, path: str, format: str = "png",
               dpi: int | None = None, overwrite: bool = False,
               selection: bool = False) -> dict:
    """导出 png / jpg / svg / pdf。

    Args:
        path: 目标文件绝对路径。
        format: png / jpg / svg / pdf。
        dpi: 位图导出分辨率（png/jpg 有效）。实测：resolution≈dpi，大页面误差 <1%。
        overwrite: 目标已存在时是否覆盖。
        selection: 只导出当前选中对象（png/jpg/svg）。
    """
    def _run():
        fmt = (format or "").lower().lstrip(".")
        if fmt not in {"png", "jpg", "svg", "pdf"}:
            return {"error": "bad_argument", "message": "format 只支持 png / jpg / svg / pdf"}
        p, warn = resolve_write_path(path, overwrite=overwrite)
        doc, _sess = _with_doc(document_id)
        result = _com(lambda _b: ops.export(doc, fmt, str(p), dpi=dpi, selection=selection),
                      timeout=DEFAULT_EXPORT_TIMEOUT_S, label="cdr_export")
        if fmt in {"png", "jpg", "svg"}:
            result["scope"] = "selection" if selection else "current_page"
        if warn:
            result.setdefault("warnings", []).append(warn)
        return result

    return _guard(_run)


@mcp.tool()
def cdr_save(document_id: str, path: str, overwrite: bool = False,
             as_copy: bool = False) -> dict:
    """另存为 .cdr（默认不覆盖已有文件，也不会改动源文档）。

    Args:
        as_copy: true 时用 SaveAsCopy —— 写副本但**保持当前文档的关联路径不变**
                 （适合"先存个中间版本再继续改"）。
    """
    def _run():
        p, warn = resolve_write_path(path, overwrite=overwrite)
        if p.suffix.lower() != ".cdr":
            return {"error": "bad_argument", "message": "保存路径必须以 .cdr 结尾"}
        doc, sess = _with_doc(document_id)
        if as_copy:
            info = _com(lambda _b: ops.save_as_copy(doc, str(p)), label="cdr_save:copy")
            # 实测：SaveAsCopy 会把文档名改成副本名，必须同步会话
            registry().rename(sess.name, info["name"])
            result = {"saved": str(p), "bytes": ops._s(lambda: os.path.getsize(p)),
                      "as_copy": True, "document_name": info["name"]}
            if warn:
                result.setdefault("warnings", []).append(warn)
            return _commit(doc, sess, result)
        info = _com(lambda _b: ops.save_as(doc, str(p)), label="cdr_save")
        registry().rename(sess.name, info["name"])
        result = {"saved": str(p), "bytes": ops._s(lambda: os.path.getsize(p))}
        if warn:
            result.setdefault("warnings", []).append(warn)
        return _commit(doc, sess, result)

    return _guard(_run)


@mcp.tool()
def cdr_undo(document_id: str, steps: int = 1, redo: bool = False) -> dict:
    """撤销（或重做）最近 N 步。

    Args:
        steps: 步数；一次工具调用 = 一步（每次变更都做成了原子 undo 组）。
        redo: true 时改为**重做**。
    """
    def _run():
        doc, sess = _with_doc(document_id)
        result = _com(lambda _b: ops.undo(doc, steps, redo=bool(redo)), label="cdr_undo")
        return _commit(doc, sess, result)

    return _guard(_run)


@mcp.tool()
def cdr_bitmap(document_id: str, target: dict, action: str = "info",
               mode: str | None = None, width: float | None = None,
               height: float | None = None, dpi: float | None = None,
               anti_alias: bool = True, threshold: float = 128.0,
               inflate_x: float = 2.0, inflate_y: float = 2.0,
               expected_revision: int | None = None,
               snapshot_token: str | None = None,
               page_index: int | None = None) -> dict:
    """位图处理（只对 type_id=5 的位图对象有效，例如 cdr_import 导入的图片）。

    action:
      info          读取尺寸、像素、分辨率、颜色模式
      convert       改颜色模式，mode = rgb / cmyk / gray / bw
      bw            阈值二值化（threshold 0~255），适合做丝印/单色稿
      resample      重采样，width/height 目标像素、dpi 目标分辨率
      inflate       四周扩边，inflate_x / inflate_y 毫米
      crop          执行裁剪
      reset_crop    取消裁剪

    ⚠️ **不提供 PowerTRACE（位图自动转矢量）**：实测 `Bitmap.Trace` 会让 CorelDRAW
    卡成「无响应」（弹出模态对话框），而 COM 调用无法中断，会把调用方一起拖死。
    需要描摹请在 CorelDRAW 里手动做，或改用位图→SVG 的外部工具。
    """
    def _run():
        act = (action or "info").lower()
        doc, sess = _with_doc(document_id)
        if act == "info":
            return _com(lambda _b: ops.bitmap_action(doc, page_index, target, "info"),
                        label="cdr_bitmap")
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        result = _com(lambda _b: ops.bitmap_action(
            doc, page_index, target, act, mode=mode, width=width, height=height,
            dpi=dpi, anti_alias=anti_alias, threshold=threshold,
            inflate_x=inflate_x, inflate_y=inflate_y), label="cdr_bitmap")
        return _commit(doc, sess, result, warnings)

    return _guard(_run)


@mcp.tool()
def cdr_close(document_id: str, save: bool = False) -> dict:
    """关闭文档。默认丢弃未保存改动。"""
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
    transport = os.environ.get("CDR_TRANSPORT", "stdio").lower()
    if transport in {"http", "streamable-http", "sse"}:
        mcp.settings.host = os.environ.get("CDR_HOST", "127.0.0.1")
        mcp.settings.port = int(os.environ.get("CDR_PORT", "8765"))
        mcp.run(transport="streamable-http")
    else:
        mcp.run()


if __name__ == "__main__":
    main()
