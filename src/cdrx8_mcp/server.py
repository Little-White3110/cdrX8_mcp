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
        doc = ops.find_document(app, sess.name)
        try:
            bridge._install_doc_events(doc)
        except Exception:
            pass
        return doc

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
        info["tools"] = 27
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
             preset: str | None = None, template: str | None = None,
             include_graphics: bool = True) -> dict:
    """打开已有 .cdr，或新建空白文档，或**从模板创建**。

    Args:
        path: 要打开的 .cdr 绝对路径；为空则新建。
        width_mm / height_mm: 新建文档尺寸（毫米）。
        orientation: "portrait" / "landscape"。
        preset: 页面预设，如 "A4"/"A3"/"card"（名片）；给了就忽略 width/height。
        template: 模板文件路径（实测 .cdr 也可当模板）；给定时走
                  `Application.CreateDocumentFromTemplate`。
        include_graphics: 从模板创建时是否带模板里的图形（默认 true）。
    """
    def _run():
        target = str(resolve_read_path(path)) if path else None
        tpl = str(resolve_read_path(template)) if template else None
        if target and tpl:
            return {"error": "bad_argument", "message": "path 与 template 只能给一个"}

        def _do(_b):
            app = _b.ensure_app()
            if tpl:
                doc = app.CreateDocumentFromTemplate(tpl, bool(include_graphics))
            elif target:
                doc = ops.open_document(app, target)
            else:
                doc = ops.create_document(app, width_mm, height_mm, orientation, preset)
            name = str(doc.Name)
            with ops.mm_units(doc):
                page = doc.ActivePage
                return {
                    "name": name,
                    "file_name": ops._s(lambda: str(doc.FileName)) or None,
                    "width_mm": ops._round(ops._s(lambda: float(page.SizeWidth))),
                    "height_mm": ops._round(ops._s(lambda: float(page.SizeHeight))),
                    "pages": ops._n(lambda: int(doc.Pages.Count), 1),
                    "from_template": bool(tpl),
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
      * bitmap:    x,y,width,height[,color,pixels_w,pixels_h,image_type,fills]
                   —— 从内存像素缓冲区创建位图；
                   fills=[{x,y,width,height,color},…] 按像素坐标涂色块；
                   image_type: 4=RGB（默认）5=CMYK 2=灰度 0=黑白
      * text:      x,y,text[,font,size,bold,italic,underline,align]
      * paragraph_text: x,y,x2,y2,text[,font,size,align]
      * guide:     x1,y1,x2,y2（参考线，不是形状）
      * ole_file:  path（绝对路径）[,link,display_as_icon,caption] —— 嵌入 OLE 对象
                   （实测 .txt 可嵌入，返回 type_id=12 的 OLE 形状）

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
      fill（纯色/渐变/纹理/图案/"none"）｜fill_from（复制另一对象的填充：{"name":"..."}）
      图案填充：fill={"type":"pattern","canvas_index":1,"front":"#000","end":"#fff"}
      （canvas_index 来自 cdr_doc(action="pattern_canvas")）
      图案微调：pattern_fill={"rotation_angle":30,"origin_x":2,"tile_width":8,
                "mirror_fill":true,"front_color":"#0F0"}（作用于已有图案填充）
      专色填充：fill={"type":"spot","palette_index":1,"color_index":10,"tint":100}
      （先 cdr_doc(action="palette_open_fixed", palette_id=3) 打开 PANTONE 调色板）
      outline（宽度或 {width,color,caps,join,behind_fill,dash,start_arrow,end_arrow}）
      outline 的 color 也可用专色：{"type":"spot","palette_index":1,"color_index":1}
      自定义虚线：outline={"width":2,"dash_pattern":[2,2],"save_outline_styles":false}
                （2/4/6 个数字，实线/空白交替；实测会注册成新的线型样式）
      对象数据：changes={"data":{"SKU":"A-123"},"data_format":"0"}（数据合并字段，持久化在文档里）
      叠印：changes={"overprint_fill":true,"overprint_outline":true,"overprint_bitmap":true}
      outline_from（复制另一对象的描边：{"name":"..."}）
      drop_shadow {opacity,offset_x,offset_y,color}｜contour {offset,steps}｜envelope 1..39
      transparency（0~100 均匀透明度 / {type:"fountain",start,end,angle} /
                    {type:"pattern",canvas_index:1,front:0,back:100} / "none"）
      freeze_transparency（true=冻结透明度并补一个新名字，false=解冻；
                          实测冻结会把对象变成组合体并清空名字，已自动补名）
      gradient {angle, edge_pad, translate:[dx,dy], transformations:[d11,d12,d21,d22], opaque}
      effect {type:"bci"|"gamma"|"hsl"|"invert"|"posterize"|"color_balance", ...}
      flip（0 水平 / 1 垂直）｜skew {x,y}｜clone {dx,dy}｜locked true/false
      order（"front_of"|"back_of" 需 ref / "forward_one" / "back_one" / "reverse"）
      clear_transformations（清除对象变换）｜rotation_center {x,y}
      copy_clipboard（复制到系统剪贴板，可跨文档）/ cut_clipboard（剪切）
      move_to_layer（图层名）/ copy_to_layer（图层名）
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
             indent_left: float | None = None, indent_right: float | None = None,
             indent_first_line: float | None = None, clear_tabs: bool = False,
             tab_stops: list | None = None,
             insert_before: str | None = None, insert_after: str | None = None,
             to_paragraph: bool = False, to_artistic: bool = False,
             contents: str | None = None, find: str | None = None,
             replace_old: str | None = None, replace_new: str | None = None,
             select_all: bool = False, align_to_baseline: bool = False,
             copy_attributes_from: dict | None = None,
             read_contents: bool = False, range_start: int | None = None,
             range_end: int | None = None, export_path: str | None = None,
             export_start: int = 1, export_count: int = 0,
             export_overwrite: bool = False, opentype: dict | None = None,
             read_font: bool = False, read_hyphenation: bool = False,
             range_delete: bool = False, range_copy: bool = False,
             range_paste: bool = False, range_straighten: bool = False,
             find_query: str | None = None, runs: bool = False,
             expected_revision: int | None = None, snapshot_token: str | None = None,
             page_index: int | None = None) -> dict:
    """文字的排版级操作。

    内容与类型：text / contents（段落文本整体替换）/ find / replace_old+replace_new /
                to_paragraph / to_artistic（艺术字↔段落字互转）
    样式：font / size / bold / italic / underline / alignment /
          change_case（sentence·lower·upper·title·toggle）/ bullet / drop_cap /
          line_spacing / char_spacing / word_spacing /
          indent_left / indent_right / indent_first_line（段落缩进，mm）/
          tab_stops（制表位：["10mm" 数值 / {position,alignment,leadered}]）/
          clear_tabs（先清空制表位）
    文本区间：select_all / align_to_baseline / copy_attributes_from（复制另一对象的文字属性）
    版式：fit_to_path（沿曲线）/ fit_frame（自动适配文本框）/ insert_before / insert_after
    输出：to_curves（转曲，发排前必做）
    读取/导出：read_contents（读全文）/ range_start+range_end（读指定区间，0 基）/
              export_path（导出纯文本，export_start 从 1 起）
              read_font（读字体属性：名称/字号/样式/下划线…）/
              read_hyphenation（读断字设置；实测 X8 断字设置写回无效，只读）
              find_query（按格式查询文本区间，如 "@font = 'Arial'" / "@bold = true"）
              runs（列出格式段：每段的文字/字体/字号/粗斜体/起止位置）
    区间操作：range_delete / range_copy / range_paste / range_straighten
              （作用于 range_start..range_end，0 基；copy/paste 走系统剪贴板）
    OpenType：opentype={"liga":true,"kern":false,...}（实测 SetOpenTypeFeature 可用）

    注意：contents / replace_old / find 走的是段落文本接口，对艺术字会返回 warning。
    """
    def _run():
        doc, sess = _with_doc(document_id)
        mutating = any(v is not None for v in (
            text, font, size, bold, italic, underline, alignment,
            fit_to_path, change_case, bullet, drop_cap, line_spacing,
            char_spacing, word_spacing, indent_left, indent_right,
            indent_first_line, insert_before, insert_after,
            contents, find, replace_old, replace_new, copy_attributes_from,
            opentype,
        )) or any((to_curves, fit_frame, no_effect, to_paragraph, to_artistic,
                   select_all, align_to_baseline, range_delete, range_copy,
                   range_paste, range_straighten, clear_tabs)) or bool(tab_stops)
        warnings = _check_view(sess, expected_revision, snapshot_token, doc) \
            if mutating else []
        target_path = resolve_write_path(export_path, overwrite=export_overwrite)[0] \
            if export_path else None
        result = _com(lambda _b: ops.set_text(
            doc, page_index, target, text=text, font=font, size=size, bold=bold,
            italic=italic, underline=underline, alignment=alignment, to_curves=to_curves,
            fit_to_path=fit_to_path, fit_frame=fit_frame, change_case=change_case,
            bullet=bullet, drop_cap=drop_cap, no_effect=no_effect,
            line_spacing=line_spacing, char_spacing=char_spacing,
            word_spacing=word_spacing, indent_left=indent_left,
            indent_right=indent_right, indent_first_line=indent_first_line,
            clear_tabs=clear_tabs, tab_stops=tab_stops, insert_before=insert_before,
            insert_after=insert_after, to_paragraph=to_paragraph, to_artistic=to_artistic,
            contents=contents, find=find, replace_old=replace_old,
            replace_new=replace_new, select_all=select_all,
            align_to_baseline=align_to_baseline,
            copy_attributes_from=copy_attributes_from,
            read_contents=read_contents, range_start=range_start, range_end=range_end,
            export_path=str(target_path) if target_path else None,
            export_start=export_start, export_count=export_count,
            opentype=opentype,
            read_font=read_font, read_hyphenation=read_hyphenation,
            range_delete=range_delete, range_copy=range_copy,
            range_paste=range_paste, range_straighten=range_straighten,
            find_query=find_query, runs=runs,
        ), label="cdr_text")
        if not mutating:
            return result
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
               all_shapes: bool = False, remove: bool = False,
               clear: bool = False) -> dict:
    """把对象加入当前选择集，供 cdr_shape_ops(group/combine/ungroup) 使用。

    type_id: 1=矩形 2=椭圆 3=曲线 4=多边形 6=文字 7=网格
    rect: {"x1":0,"y1":0,"x2":100,"y2":50,"touch":true} 按矩形框选
    remove=true 从选择集移除；clear=true 直接清空选择集
    """
    def _run():
        doc, sess = _with_doc(document_id)

        def _do(_b):
            picked = ops.select_shapes(doc, page_index, name=name, type_id=type_id,
                                       rect=rect, all_shapes=all_shapes,
                                       remove=remove, clear=clear)
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
                  copies: int = 1, distance_x: float = 0, distance_y: float = 0,
                  mode_x: int = 1, direction_x: int = 1,
                  mode_y: int = 1, direction_y: int = 3,
                  distribute_type: int = 1, page_extent: bool = False,
                  place_on_top: bool = True, delete_source: bool = False,
                  x: float | None = None, y: float | None = None,
                  hx: float | None = None, hy: float | None = None,
                  vx: float | None = None, vy: float | None = None,
                  curve_count: int = 2, curve_distance: float = 5.0,
                  resolution: int = 72, grayscale: bool = False,
                  transparent: bool = True,
                  envelope_mode: int = 2, compare_type: int = 1,
                  condition: int = 0, pixel_aligned: bool = True,
                  fill_mode: int = 0, plane: int = 0,
                  reference_point: int = 9, apply_to_duplicate: bool = False,
                  extrude_type: int = 1, vp_type: int = 1,
                  vp_x: float | None = None, vp_y: float | None = None,
                  depth: float = 20.0, shading: int = 0,
                  base_color: Any = "#808080", shading_color: Any = "#404040",
                  bevel_depth: float = 0.0, bevel_angle: float = 45.0,
                  bevel_color: Any = "#000000", bevel_only: bool = False,
                  lens_type: int = 1, rate: float = 50.0,
                  color1: Any = "#FFFFFF", color2: Any = "#000000",
                  color_map_palette: int = 0, effect: dict | None = None,
                  steps: int = 20, color_blend_type: int = 0,
                  blend_mode: int = 0, spacing: float = 0.0,
                  angle: float = 0.0, loop: bool = False,
                  rotate_shapes: bool = False, spacing_accel: int = 0,
                  color_accel: int = 0, accel_size: bool = False,
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
      step_repeat（copies,distance_x,distance_y,mode_x,direction_x,mode_y,direction_y）
                 阵列复制（X8 原生 StepAndRepeat，实测 direction_y 必须 2 或 3）
      distribute（distribute_type,page_extent）  分布排列（X8 原生 Distribute）
      boundary（x,y,place_on_top,delete_source） 生成形状边界曲线
      perspective（可选 hx,hy,vx,vy）            透视效果（实测无参可用）
      separate                                   拆开 target 上的效果/组合
      parallel_curves（curve_count,curve_distance）范围级：并行曲线（只支持曲线对象，先转曲）
      rasterize（resolution,grayscale,transparent）范围级：把选择集栅格化成位图
      batch_outline_to_curves  范围级：选择集描边转曲
      fill_mode（fill_mode 0~2）  范围级：填充规则
      pixel_aligned（True/False） 范围级：像素对齐渲染（输出小尺寸位图更清晰）
      range_boundary（x,y）      范围级：选择集边界曲线
      envelope_from_shape（other,envelope_mode）用另一对象/曲线做封套
      compare（other,compare_type,condition）对象比较（返回 same 布尔）
      project / unproject（plane 0=front 1=right 2=top, reference_point 1~9,
                          apply_to_duplicate）把对象投影到 3D 平面/反投影
      to_document（target 或当前选择集）把对象/选区提取成**新文档**（返回 document_id）
      extrude（extrude_type,vp_type,vp_x/vp_y,depth,shading,base_color,shading_color,
               bevel_depth,bevel_angle,bevel_color,bevel_only）创建立体化效果；
               effect={depth,angle_x,angle_y,angle_z,use_bevel,...} 微调
      lens（lens_type,rate,color1,color2,color_map_palette）创建透镜效果；
               effect={rate,magnification,fill_color,...} 微调
      extrude_set / lens_set  只改已有效果属性（effect={...}）
      effect_info            列出对象上的效果
      clear_effects          清除对象上的所有效果
      blend（target+other，steps/spacing/angle/color_blend_type/blend_mode/
             loop/rotate_shapes/spacing_accel/color_accel/accel_size）
             创建混合效果；effect={steps,spacing,angle,...} 微调
      blend_set              只改已有混合效果属性
    """
    def _run():
        doc, sess = _with_doc(document_id)
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        result = _com(lambda _b: ops.shape_op(
            doc, page_index, op, target=target, other=other, radius=radius,
            distance_a=distance_a, distance_b=distance_b,
            leave_source=leave_source, leave_target=leave_target,
            symbol_name=symbol_name, copies=copies,
            distance_x=distance_x, distance_y=distance_y,
            mode_x=mode_x, direction_x=direction_x,
            mode_y=mode_y, direction_y=direction_y,
            distribute_type=distribute_type, page_extent=page_extent,
            place_on_top=place_on_top, delete_source=delete_source,
            x=x, y=y, hx=hx, hy=hy, vx=vx, vy=vy,
            curve_count=curve_count, curve_distance=curve_distance,
            resolution=resolution, grayscale=grayscale, transparent=transparent,
            envelope_mode=envelope_mode, compare_type=compare_type,
            condition=condition, pixel_aligned=pixel_aligned, fill_mode=fill_mode,
            plane=plane, reference_point=reference_point,
            apply_to_duplicate=apply_to_duplicate,
            extrude_type=extrude_type, vp_type=vp_type, vp_x=vp_x, vp_y=vp_y,
            depth=depth, shading=shading, base_color=base_color,
            shading_color=shading_color, bevel_depth=bevel_depth,
            bevel_angle=bevel_angle, bevel_color=bevel_color,
            bevel_only=bevel_only, lens_type=lens_type, rate=rate,
            color1=color1, color2=color2, color_map_palette=color_map_palette,
            effect=effect, steps=steps, color_blend_type=color_blend_type,
            blend_mode=blend_mode, spacing=spacing, angle=angle, loop=loop,
            rotate_shapes=rotate_shapes, spacing_accel=spacing_accel,
            color_accel=color_accel, accel_size=accel_size,
        ), label="cdr_shape_ops")
        if str(op).lower() == "to_document" and result.get("document_name"):
            new_sess = registry().get_or_create(result["document_name"])
            result["document_id"] = new_sess.sid
            result["hint"] = "新文档已建会话；调用 cdr_inspect(document_id) 获取 revision/token"
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
               move_below: str | None = None,
               expected_revision: int | None = None,
               snapshot_token: str | None = None) -> dict:
    """图层管理。

    action: list / create（给 name）/ activate / set（visible/editable/printable）/
            move（move_above=目标图层名）/ move_below（move_below=目标图层名）/
            above / below（读相邻图层）/ paste（把系统剪贴板内容粘到该图层）/ delete
    layer_index 或 name 用来定位图层；都不给则作用于活动图层。
    """
    def _run():
        doc, sess = _with_doc(document_id)
        act = (action or "list").lower()
        if act == "list":
            return _com(lambda _b: {"layers": ops.list_layers(doc, page_index)},
                        label="cdr_layers")
        warnings = _check_view(sess, expected_revision, snapshot_token, doc) if act not in {
            "above", "below"} else []
        result = _com(lambda _b: ops.layer_action(
            doc, act, name=name, page_index=page_index, index=layer_index,
            visible=visible, editable=editable, printable=printable, move=move_above,
            move_below=move_below,
        ), label="cdr_layers")
        if act in {"above", "below"}:
            return result
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
def cdr_node(document_id: str, target: dict, action: str = "list",
             node: int | None = None, start: int | None = None, end: int | None = None,
             all_nodes: bool = False, x: float | None = None, y: float | None = None,
             dx: float = 0.0, dy: float = 0.0, node_type: int | None = None,
             radius: float = 0.0, distance_a: float = 0.0, distance_b: float = 0.0,
             ratio_x: float = 1.0, ratio_y: float = 1.0, angle: float = 0.0,
             angle_x: float = 0.0, angle_y: float = 0.0, offset: float = 0.5,
             segment_index: int | None = None, control: str | None = None,
             hot_area: float = -1.0,
             expected_revision: int | None = None, snapshot_token: str | None = None,
             page_index: int | None = None) -> dict:
    """节点 / 线段级编辑（X8 实测可用：curve.Nodes / curve.Segments）。

    target 定位曲线对象（不是曲线的先 `cdr_shape_ops(op="to_curves")`）。

    节点选择器：node=n（第 n 个节点）｜start+end（节点范围）｜all_nodes=true。
    读操作默认列全部节点；**写操作必须显式指定** node / start+end / all_nodes。

    action（只读）：
      list           列出节点（索引/坐标/类型/所在子路径/是否端点）
      segments       列出线段（类型/长度/端点/贝塞尔控制点/中点/曲率）
      segment_info   单条线段详情（segment_index + offset）
      find           按坐标找最近的节点与线段（x,y,hot_area）

    action（变更）：
      move           dx,dy 相对移动；单节点也可用 x,y 绝对定位
      select         选中节点（供后续节点级操作/查看）
      delete         删除节点
      type           node_type（0~2；实测 X8 对多数节点会忽略，改平滑请用 smoothen）
      fillet（radius）/ chamfer（distance_a,distance_b）/ scallop（radius）
      smoothen（radius 当平滑度 0~100）/ stretch（ratio_x,ratio_y）/
      rotate（angle）/ skew（angle_x,angle_y）/ break_apart
      add_node（segment_index + offset）      在线段上加节点
      break_segment（segment_index + offset） 在线段处断开
      control_points（segment_index[,control="start"/"end",x,y]）读写贝塞尔控制点
    """
    def _run():
        act = (action or "list").lower()
        doc, sess = _with_doc(document_id)
        kwargs = dict(node=node, start=start, end=end, all_nodes=all_nodes,
                      x=x, y=y, dx=dx, dy=dy, node_type=node_type, radius=radius,
                      distance_a=distance_a, distance_b=distance_b,
                      ratio_x=ratio_x, ratio_y=ratio_y, angle=angle,
                      angle_x=angle_x, angle_y=angle_y, offset=offset,
                      segment_index=segment_index, control=control,
                      hot_area=hot_area)
        if act not in ops.NODE_MUTATIONS or (act == "control_points" and control is None):
            return _com(lambda _b: ops.node_ops(doc, page_index, target, act, **kwargs),
                        label="cdr_node")
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        result = _com(lambda _b: ops.node_ops(doc, page_index, target, act, **kwargs),
                      label="cdr_node")
        return _commit(doc, sess, result, warnings)

    return _guard(_run)


@mcp.tool()
def cdr_style(document_id: str, action: str = "list",
              target: dict | None = None, category: str = "fill",
              style_name: str | None = None, new_name: str | None = None,
              based_on: str | None = None, property: str | None = None,
              value: Any = None, sheet_path: str | None = None,
              use_fills: bool = True, use_outlines: bool = True,
              selected_only: bool = False, harmonies: int = 0,
              convert_colors_to: int = 99,
              expected_revision: int | None = None, snapshot_token: str | None = None,
              page_index: int | None = None) -> dict:
    """对象样式（Object Styles）：列出 / 创建 / 派生 / 套用 / 改属性 / 导入导出样式表 / 颜色样式。

    action:
      list                       列出文档样式（名称、类别）
      get（style_name）          读样式定义（可读属性、派生样式、ToString JSON）
      create（target, category, style_name）
                                 从对象创建样式；category 实测可用：
                                 fill / outline（图形），character / paragraph（文字）
      create（category, based_on, style_name）
                                 基于已有样式派生（派生样式在 base 的 DerivedStyles 里）
      apply（target, style_name）把样式套到对象
      rename（style_name, new_name）
      delete（style_name）
      set_property（style_name, property, value）
                                 改样式属性；标量走字符串，*color 结尾的用颜色对象
      derived（style_name）      列出派生样式
      export_sheet（sheet_path） 把样式表存成 .cdst
      import_sheet（sheet_path） 从 .cdst 载入样式
      color_styles               列出颜色样式（AllColorStyles）
      create_color_styles（harmonies, selected_only, convert_colors_to）
                                 从文档/选区生成颜色样式与颜色协调

    历史结论「样式类别常量 0~8 全失败」是误判：Category 是**字符串**不是数字。
    套用必须用 `Shape.ApplyStyle("样式名")`（传 Style 对象会报「未找到样式」）。
    """
    def _run():
        act = (action or "list").lower()
        doc, sess = _with_doc(document_id)
        resolved_sheet = None
        if sheet_path:
            if act == "import_sheet":
                resolved_sheet = str(resolve_read_path(sheet_path))
            elif act == "export_sheet":
                resolved_sheet = str(resolve_write_path(sheet_path, overwrite=True)[0])
        kwargs = dict(target=target, category=category, name=style_name,
                      new_name=new_name, based_on=based_on, property=property,
                      value=value, sheet_path=resolved_sheet,
                      use_fills=use_fills, use_outlines=use_outlines,
                      selected_only=selected_only, harmonies=harmonies,
                      convert_colors_to=convert_colors_to)
        if act not in ops.STYLE_MUTATIONS:
            return _com(lambda _b: ops.style_ops(doc, page_index, act, **kwargs),
                        label="cdr_style")
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        result = _com(lambda _b: ops.style_ops(doc, page_index, act, **kwargs),
                      label="cdr_style")
        return _commit(doc, sess, result, warnings)

    return _guard(_run)


@mcp.tool()
def cdr_pages(document_id: str, action: str = "list",
              page_index: int | None = None, count: int = 1,
              width_mm: float | None = None, height_mm: float | None = None,
              preset: str | None = None, landscape: bool = False,
              x: float | None = None, y: float | None = None,
              text: str | None = None, move_to_index: int | None = None,
              hot_area: float = 2.0, angle: float = 45.0,
              insert_before: bool = True, snap_types: int = 2,
              expected_revision: int | None = None,
              snapshot_token: str | None = None) -> dict:
    """页面管理。

    action:
      list                          列表（含图层与对象数）
      add（count）                  末尾新增页
      insert（count, page_index）   在第 page_index 页前/后插入空白页
                                    （insert_before=true 插到前面；实测 Document.InsertPages）
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
      guide_angle（x,y,angle）      按基点+角度加参考线（X8 CreateGuideAngle）
      snap_point（x,y,snap_types） 找最近吸附点（位掩码 2对象 4包围盒 8边；3/6/15/63 可用）
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
        if act == "snap_point":
            if x is None or y is None:
                return {"error": "bad_argument",
                        "message": "action=snap_point 需要 x 与 y"}
            return _com(lambda _b: ops.page_snap_point(
                doc, page_index, x, y, snap_types, hot_area),
                label="cdr_pages:snap_point")

        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        if act == "guide_angle":
            if x is None or y is None:
                return {"error": "bad_argument",
                        "message": "action=guide_angle 需要 x 与 y（基点）"}
            result = _com(lambda _b: ops.add_guide_angle(doc, page_index, x, y, angle),
                          label="cdr_pages:guide_angle")
            return _commit(doc, sess, result, warnings)
        if act == "unlock":
            result = _com(lambda _b: ops.unlock_all(doc, page_index),
                          label="cdr_pages:unlock")
            return _commit(doc, sess, result, warnings)
        if act == "add":
            result = _com(lambda _b: ops.add_pages(doc, count), label="cdr_pages:add")
        elif act == "insert":
            result = _com(lambda _b: ops.insert_pages(doc, count, insert_before, page_index),
                          label="cdr_pages:insert")
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
                    "message": f"未知 action：{act}；可用 list/add/insert/activate/delete/size/"
                               f"preset/guides/guide_angle/snap_point/unlock/bounds"}

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
             precision: int = 50,
             tolerance: float = 0.0, allow_reversals: bool = True, hot_area: float = -1.0,
             expected_revision: int | None = None,
             snapshot_token: str | None = None,
             page_index: int | None = None) -> dict:
    """曲线 / 子路径级操作（对象需是曲线，可先 `cdr_shape_ops(op="to_curves")` 转曲）。

    只读：info（长度/面积/子路径）、polyline（折线采样，已修复）、point_inside、intersects、
          subpath_info、get_point、get_tangent、
          curvature（曲率）、perpendicular（法线角度）、evenly_spaced（等距点）、
          closest_segment（最近线段）、is_on_path（点是否在路径上）、
          segment_at（某位置所在线段）、is_rect_on_edge（矩形是否贴边）、
          point_position（t 位置坐标）、tangent_vector / perpendicular_vector（向量+角度）、
          curve_speed（速度）、intersections（与另一条曲线的交点）、
          find_node / find_segment（按坐标找节点/线段）

    变更：contour（偏移出刀模线）、add_subpath、
          append_circle / append_ellipse / append_rectangle / append_arc / append_bezier
          （往同一条曲线里拼更多子路径）、
          reduce_nodes（节点精简）、self_weld（自焊接）、
          equal_divide（等分）、add_node、break_apart、move_subpath、
          reverse_direction（反转方向）、join_subpaths（焊接相邻子路径）、
          remove_overlaps（去重叠）
    """
    def _run():
        act = (action or "info").lower()
        readonly = act in {"info", "polyline", "point_inside", "intersects",
                           "subpath_info", "get_point", "get_tangent",
                           "curvature", "perpendicular", "evenly_spaced",
                           "closest_segment", "is_on_path", "segment_at",
                           "is_rect_on_edge", "point_position", "tangent_vector",
                           "perpendicular_vector", "curve_speed", "intersections",
                           "find_node", "find_segment"}
        doc, sess = _with_doc(document_id)
        kwargs = dict(offset=offset, direction=direction, x=x, y=y, radius=radius,
                      rx=rx, ry=ry, rect=rect, arc=arc, bezier=bezier, other=other,
                      amount=amount, subpath=subpath, t=t, dx=dx, dy=dy,
                      precision=precision, tolerance=tolerance,
                      allow_reversals=allow_reversals, hot_area=hot_area)
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
            value: float | None = None, from_unit: int | None = None,
            to_unit: int | None = None, x: float | None = None,
            y: float | None = None, x2: float | None = None, y2: float | None = None,
            x_samples: int = 4, y_samples: int = 4, layout: int | None = None,
            query: str | None = None, limit: int = 100,
            palette_index: int | None = None, color_index: int = 1,
            palette_id: int | None = None, spot_name: str | None = None,
            tint: float = 100.0,
            pattern_fills: list | None = None, pattern_pixels: list | None = None,
            pattern_lines: list | None = None, clear_canvas: bool = True,
            metadata: dict | None = None, palette_name: str | None = None,
            palette_path: str | None = None, palette_colors: list | None = None,
            overwrite: bool = False, field_name: str | None = None,
            field_format: str = "0", settings: dict | None = None,
            register: bool = False, copy_areas: list | None = None,
            rotate_areas: list | None = None, flip_areas: list | None = None,
            model: int | str | None = None, profile_name: str | None = None,
            rendering_intent: int | None = None,
            pref_group: str = "cdrX8 MCP", pref_key: str | None = None,
            pref_value: Any = None, pref_default: Any = None,
            view_name: str | None = None, view_index: int = 1,
            zoom: float = 100.0, origin_x: float = 0.0,
            origin_y: float = 0.0, grid_spacing_x: float | None = None,
            grid_spacing_y: float | None = None,
            grid_visible: bool | None = None, grid_snap: bool | None = None,
            grid_type: int | None = None,
            expected_revision: int | None = None, snapshot_token: str | None = None,
            page_index: int | None = None) -> dict:
    """文档级操作。

    action:
      info             文档概况（路径、脏标记、页数、单位、对象总数）
      clear_selection  清空选择集
      clear_undo       清空撤销栈（批处理前后释放内存）
      palette          把文档用到的颜色登记进调色板（印前交接）
      arrowhead        用一条曲线生成自定义箭头（target 指向曲线）
      units            单位换算：value + from_unit（换算到文档单位）或 to_unit（文档单位→目标）
                       单位枚举实测：1=in 3=mm 4=cm 5=px 7=m
      sample_color     取色：x,y（采样一个点）或再加 x2,y2（区域平均）
      show_all         显示所有隐藏对象
      layout           页面版式（layout 数字枚举，实测 1 可用）
      repeat           重复上一操作（等价 Ctrl+R，慎用）
      fonts            列出已安装字体（query 过滤、limit 截断，实测 2000+ 字体）
      palettes         列出调色板（名称/GUID/色数/是否默认）
      palette_color    读调色板颜色：palette_index + color_index（1 基）
      palette_open_fixed  打开内置固定调色板：palette_id（本机 3=PANTONE 匹配系统）
      palette_colors   列调色板颜色名（query 过滤、limit 截断）
      spot_color       创建专色：palette_index + color_index（或 spot_name）+ tint
      pattern_canvas   新建 32x32 图案画布并登记：pattern_fills=[[x1,y1,x2,y2],…] /
                       pattern_pixels=[[x,y],…] / pattern_lines=[[x1,y1,x2,y2],…] /
                       copy_areas=[[x1,y1,x2,y2,dx,dy],…] /
                       rotate_areas=[[x1,y1,x2,y2,angle],…] /
                       flip_areas=[[x1,y1,x2,y2,axes],…]；
                       返回 canvas_index，供 cdr_modify 图案填充/图案透明使用
      textures         列出 Corel 纹理库名称（供 fill={"type":"texture","name":...}）
      metadata         读文档元数据（Title/Author/Subject/Keywords/Notes/Copyright…）
      set_metadata     写文档元数据（metadata={...}，只写白名单字段）
      metadata_localized       读本地化元数据（各语言 Title/Subject/Keywords/Notes/Copyright）
      set_metadata_localized   写本地化元数据：metadata={"Title":{"default":"…","2052":"…"}}
      palette_create   新建自定义调色板并落盘：palette_name + palette_path +
                       palette_colors=[...]；返回 palette_index（可用于位图 9 号调色板模式）
      data_fields      列出文档数据字段（数据合并用）
      add_data_field   新增数据字段：field_name + field_format（实测 "0" 可用）
      object_data      读对象的数据字段值（target 指定对象）
      color_context    读文档色彩管理（渲染意图、混合模式、配置文件列表）
      outline_styles   列出线型样式（索引/虚线段数/enhanced）
      settings         读文档设置（跨图层编辑/曲线精度/分辨率/原点…）
      set_settings     写文档设置：settings={edit_across_layers,apply_to_duplicate,
                       curve_precision,resolution,preserve_selection,
                       shape_enum_direction,world_scale,facing_pages,
                       first_page_on_right_side}
      revert           丢弃未保存修改、回到上次保存状态（**破坏性**，会撤销当前改动）
      app_info         读应用信息（OpenType 特性列表、平台版本检查、UI 语言可用性）
      color_profiles   列出已安装色彩配置文件（可按 model 0=RGB/1=CMYK/2=Gray/3=Lab）
      convert_color    把文档颜色转换到指定配置文件：model 或 profile_name，
                       rendering_intent 可选（0~3）
      pref_register    注册用户级应用偏好（GroupName/KeyName/默认值），持久化在 CorelDRAW 工作区
      pref_get         读取用户偏好（未注册时可传 pref_default 自动注册）
      pref_set         写入用户偏好（未注册时自动注册再写）
      views            列出命名视图（名称/缩放/页面）
      create_view      创建命名视图：view_name + origin_x/y + zoom + page_index
      activate_view    激活指定视图（view_index 从 1 开始）
      delete_view      删除指定视图（view_index 从 1 开始）
      grid             读文档网格（类型/间距/显示/吸附）
      set_grid         写网格：grid_spacing_x/y + grid_visible + grid_snap + grid_type
      overprint        读对象的叠印设置（fill/outline/bitmap）
      events           读取应用事件队列（DocumentNew/Open/Close/Save/Export/
                       SelectionChange/WindowActivate…）
      clear_events     清空事件队列
    """
    def _run():
        act = (action or "info").lower()
        if act in {"events", "clear_events"}:
            bridge = _bridge()
            if act == "clear_events":
                return {"action": act, "cleared": bridge.clear_events()}
            items = bridge.events_snapshot(limit or 100)
            return {"action": "events", "count": len(items), "events": items,
                    "note": "事件在 cdr_connect 后开始记录；ApplicationEvent 高频事件已忽略"}
        doc, sess = _with_doc(document_id)
        resolved_palette_path = str(resolve_write_path(palette_path, overwrite=overwrite)[0]) \
            if palette_path else None
        kwargs = dict(target=target, page_index=page_index, value=value,
                      from_unit=from_unit, to_unit=to_unit, x=x, y=y, x2=x2, y2=y2,
                      x_samples=x_samples, y_samples=y_samples, layout=layout,
                      query=query, limit=limit, palette_index=palette_index,
                      color_index=color_index, palette_id=palette_id,
                      spot_name=spot_name, tint=tint,
                      pattern_fills=pattern_fills, pattern_pixels=pattern_pixels,
                      pattern_lines=pattern_lines, clear_canvas=clear_canvas,
                      metadata=metadata, palette_name=palette_name,
                      palette_path=resolved_palette_path,
                      palette_colors=palette_colors, overwrite=overwrite,
                      field_name=field_name, field_format=field_format,
                      settings=settings, register=register,
                      copy_areas=copy_areas, rotate_areas=rotate_areas,
                      flip_areas=flip_areas, model=model,
                      profile_name=profile_name,
                      rendering_intent=rendering_intent,
                      pref_group=pref_group, pref_key=pref_key,
                      pref_value=pref_value, pref_default=pref_default,
                      view_name=view_name, view_index=view_index, zoom=zoom,
                      origin_x=origin_x, origin_y=origin_y,
                      grid_spacing_x=grid_spacing_x,
                      grid_spacing_y=grid_spacing_y,
                      grid_visible=grid_visible, grid_snap=grid_snap,
                      grid_type=grid_type)
        if action == "info":
            return _com(lambda _b: ops.doc_ops(doc, "info", **kwargs), label="cdr_doc")
        if action in {"clear_selection", "units", "sample_color", "fonts",
                      "palettes", "palette_color", "palette_open_fixed",
                      "palette_colors", "spot_color", "pattern_canvas",
                      "textures", "metadata", "palette_create", "data_fields",
                      "object_data", "color_context", "outline_styles",
                      "settings", "app_info", "color_profiles",
                      "metadata_localized",
                      "pref_register", "pref_get", "pref_set",
                      "views", "create_view", "activate_view", "delete_view",
                      "grid", "set_grid", "overprint"}:
            return _com(lambda _b: ops.doc_ops(doc, action, **kwargs), label="cdr_doc")
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        result = _com(lambda _b: ops.doc_ops(doc, action, **kwargs), label="cdr_doc")
        return _commit(doc, sess, result, warnings)

    return _guard(_run)


@mcp.tool()
def cdr_view(document_id: str, action: str = "refresh",
             x: float | None = None, y: float | None = None,
             distance: float = 10.0, text: str | None = None,
             x2: float | None = None, y2: float | None = None,
             radius: float = 20.0, color: int = 0x0000FF,
             width: int = 2) -> dict:
    """视图/窗口操作。

    action:
      refresh              强制重绘（导出前刷新、或改完看效果）
      activate             激活窗口
      to_screen（x,y）      文档坐标 -> 屏幕坐标
      to_document（x,y）    屏幕坐标 -> 文档坐标
      distance_to_screen（distance）  文档距离 -> 屏幕距离
      distance_to_document（distance）  屏幕距离 -> 文档距离
      overlay_text（text,x,y,color）        屏幕叠加文字（xor，不写入文档）
      overlay_rect（x,y,x2,y2,color,width） 屏幕叠加矩形
      overlay_line（x,y,x2,y2,color,width） 屏幕叠加直线
      overlay_circle（x,y,radius,color,width） 屏幕叠加圆
      overlay_clear                         隐藏所有叠加层
    """
    def _run():
        doc, _sess = _with_doc(document_id)
        return _com(lambda _b: ops.view_ops(doc, action, x=x, y=y,
                                           distance=distance, text=text,
                                           x2=x2, y2=y2, radius=radius,
                                           color=color, width=width),
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
               selection: bool = False, export_options: dict | None = None) -> dict:
    """导出 png / jpg / svg / pdf。

    Args:
        path: 目标文件绝对路径。
        format: png / jpg / svg / pdf。
        dpi: 位图导出分辨率（png/jpg 有效）。实测：resolution≈dpi，大页面误差 <1%。
        overwrite: 目标已存在时是否覆盖。
        selection: 只导出当前选中对象（png/jpg/svg）。
        export_options: 额外导出选项（png/jpg/svg 有效），如
            {"anti_aliasing_type":2,"dithered":false,"transparent":true,
             "resolution_x":150,"resolution_y":150,"maintain_aspect":true}
    """
    def _run():
        fmt = (format or "").lower().lstrip(".")
        if fmt not in {"png", "jpg", "svg", "pdf"}:
            return {"error": "bad_argument", "message": "format 只支持 png / jpg / svg / pdf"}
        p, warn = resolve_write_path(path, overwrite=overwrite)
        doc, _sess = _with_doc(document_id)
        result = _com(lambda _b: ops.export(doc, fmt, str(p), dpi=dpi,
                                            selection=selection,
                                            options=export_options),
                      timeout=DEFAULT_EXPORT_TIMEOUT_S, label="cdr_export")
        if fmt in {"png", "jpg", "svg"}:
            result["scope"] = "selection" if selection else "current_page"
        if warn:
            result.setdefault("warnings", []).append(warn)
        return result

    return _guard(_run)


@mcp.tool()
def cdr_save(document_id: str, path: str, overwrite: bool = False,
             as_copy: bool = False, save_options: dict | None = None) -> dict:
    """另存为 .cdr（默认不覆盖已有文件，也不会改动源文档）。

    Args:
        as_copy: true 时用 SaveAsCopy —— 写副本但**保持当前文档的关联路径不变**
                 （适合"先存个中间版本再继续改"）。
        save_options: 额外保存选项，如 {"embed_icc_profile":true,
                      "keep_appearance":true,"overwrite":true,"thumbnail_size":1}
    """
    def _run():
        p, warn = resolve_write_path(path, overwrite=overwrite)
        if p.suffix.lower() != ".cdr":
            return {"error": "bad_argument", "message": "保存路径必须以 .cdr 结尾"}
        doc, sess = _with_doc(document_id)
        if as_copy:
            info = _com(lambda _b: ops.save_as_copy(doc, str(p), options=save_options),
                        label="cdr_save:copy")
            # 实测：SaveAsCopy 会把文档名改成副本名，必须同步会话
            registry().rename(sess.name, info["name"])
            result = {"saved": str(p), "bytes": ops._s(lambda: os.path.getsize(p)),
                      "as_copy": True, "document_name": info["name"]}
            if info.get("save_options"):
                result["save_options"] = info["save_options"]
            if warn:
                result.setdefault("warnings", []).append(warn)
            return _commit(doc, sess, result)
        info = _com(lambda _b: ops.save_as(doc, str(p), options=save_options),
                    label="cdr_save")
        registry().rename(sess.name, info["name"])
        result = {"saved": str(p), "bytes": ops._s(lambda: os.path.getsize(p))}
        if info.get("save_options"):
            result["save_options"] = info["save_options"]
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
               x_px: int | None = None, y_px: int | None = None,
               duotone_type: int | None = None, ink_colors: list | None = None,
               preset: str | None = None, palette_type: int = 3,
               num_colors: int = 256, palette_index: int | None = None,
               image_color: Any = "#FFFFFF", image_pixels_w: int = 16,
               image_pixels_h: int = 16, image_fills: list | None = None,
               image_pixels: list | None = None, offset_x: int = 0,
               offset_y: int = 0, effect_name: str | None = None,
               effect_command: str | None = None,
               expected_revision: int | None = None,
               snapshot_token: str | None = None,
               page_index: int | None = None) -> dict:
    """位图处理（只对 type_id=5 的位图对象有效，例如 cdr_import 导入的图片）。

    action:
      info          读取尺寸、像素、分辨率、颜色模式
      image_info    读图像缓冲区尺寸/类型/只读标记
      pixel         取像素颜色：x_px,y_px（像素坐标，从 0 开始），返回颜色串与色名
      convert       改颜色模式，mode = rgb / cmyk / gray / bw
      bw            阈值二值化（threshold 0~255），适合做丝印/单色稿
      paletted      转调色板模式，palette_type = 0~9
                    （0=uniform 1=stdVGA 2=adaptive 3=optimized 4=blackbody
                      5=grayscale 6=system 7=IE 8=Netscape；9=自定义，需 palette_index）
      duotone       转双色调：duotone_type 0~3（单/双/三/四色调），ink_colors 给各墨色；
                    也可 preset 指定 .cpd 预设文件路径
      replace_image 用新像素缓冲区替换位图内容：image_color + image_pixels_w/h +
                    image_fills=[{x,y,width,height,color}] +
                    image_pixels=[[x,y,color],…] + offset_x/y
                    （实测 SetImageData 必须把 Alpha 传 None）
      effect        位图特效：effect_name + effect_command（命令串语法见 X8 文档），
                    例：effect_name="Whirlpool",
                        effect_command="WhirlpoolEffect WhirpoolSpacing=20,"
                                       "WhirlpoolSmear=9,WhirlpoolTwist=70,"
                                       "WhirlpoolStreak=60,WhirpoolWarp=1,"
                                       "WhirlpoolRandomSeed=0"
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
        preset_path = str(resolve_read_path(preset)) if preset else None
        if act in {"info", "image_info", "pixel"}:
            return _com(lambda _b: ops.bitmap_action(
                doc, page_index, target, act, x_px=x_px, y_px=y_px),
                label="cdr_bitmap")
        warnings = _check_view(sess, expected_revision, snapshot_token, doc)
        result = _com(lambda _b: ops.bitmap_action(
            doc, page_index, target, act, mode=mode, width=width, height=height,
            dpi=dpi, anti_alias=anti_alias, threshold=threshold,
            inflate_x=inflate_x, inflate_y=inflate_y,
            duotone_type=duotone_type, ink_colors=ink_colors, preset=preset_path,
            palette_type=palette_type, num_colors=num_colors,
            palette_index=palette_index, image_color=image_color,
            image_pixels_w=image_pixels_w, image_pixels_h=image_pixels_h,
            image_fills=image_fills, image_pixels=image_pixels,
            offset_x=offset_x, offset_y=offset_y,
            effect_name=effect_name, effect_command=effect_command), label="cdr_bitmap")
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
