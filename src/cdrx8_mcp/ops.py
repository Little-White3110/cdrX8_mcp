"""领域层：把 CorelDRAW X8 的 COM 对象模型包成 Agent 好用的语义操作。

坐标一律 **毫米 / 页面左下角原点 / Y 轴向上**。
调用前后用 `mm_units()` 临时把文档单位切成毫米再切回（实测改 Unit 不会让文档变 Dirty）。

本文件里的每个调用形态都在 CorelDRAW X8 18.0.0.450 上实跑验证过，
并与 vgcoreauto.tlb 反解出的签名核对（见 tmp/ 里的探测脚本，已清理）。
"""

from __future__ import annotations

import hashlib
import os
import struct
import time
from contextlib import contextmanager
from typing import Any, Iterable, Iterator

from .constants import (
    EXPORT_FILTERS,
    EXPORT_RANGE_CURRENT_PAGE,
    EXPORT_RANGE_SELECTION,
    NAME_PREFIX,
    PAGE_PRESETS,
    SHAPE_TYPES,
    TEXT_TYPE,
    UNIT_MM,
)
from .errors import CdrBadArgument, CdrComFailure, CdrDocumentNotFound

ORIENTATION = {1: "portrait", 2: "landscape"}
TEXT_ALIGN = {"left": 0, "center": 1, "right": 2, "full": 3, "justify": 3}


# --------------------------------------------------------------------------- #
# 基础设施
# --------------------------------------------------------------------------- #
@contextmanager
def mm_units(doc: Any) -> Iterator[None]:
    """临时把文档单位设为毫米，退出时恢复。"""
    original = None
    changed = False
    try:
        original = int(doc.Unit)
    except Exception:
        original = None
    if original != UNIT_MM:
        doc.Unit = UNIT_MM
        changed = True
    try:
        yield
    finally:
        if changed and original is not None:
            try:
                doc.Unit = original
            except Exception:
                pass


@contextmanager
def atomic(doc: Any, label: str) -> Iterator[None]:
    """把一组操作合成 CorelDRAW 的**一个 undo 步**（实测有效：3→5→3）。"""
    doc.BeginCommandGroup(label)
    try:
        yield
    finally:
        try:
            doc.EndCommandGroup()
        except Exception:
            pass


def _s(fn, default=None):
    try:
        return fn()
    except Exception:
        return default


def _n(fn, default=None):
    return _s(fn, default)


def _round(value, digits: int = 3):
    if value is None:
        return None
    try:
        return round(float(value), digits)
    except (TypeError, ValueError):
        return None


_FONT_TABLE = {"refreshed": False}


def ensure_font_table(app: Any) -> bool:
    """保证 Corel 的字体表是新的（每个进程只刷一次）。

    实测坑：长时间批量建文本后，`CreateArtisticText` 传 "Arial" 会**静默退回默认字体**
    （中文机器上是宋体）并且丢掉 Bold/Italic；调一次 `Application.ForceUpdateFontTable()`
    之后同一个调用立刻恢复正常。所以建文本前先兜底刷一次。
    """
    if _FONT_TABLE["refreshed"]:
        return True
    ok = False
    try:
        app.ForceUpdateFontTable()
        ok = True
    except Exception:
        ok = False
    _FONT_TABLE["refreshed"] = True
    return ok


def restate_text_format(shape: Any, spec: dict) -> dict:
    """建完文本后再用 Story 属性重申一遍字体/字号/粗斜下。

    实测坑：`CreateArtisticText` 的 Font/Bold/Italic 形参在长时间运行后会**静默失效**
    （字体回退成默认宋体、Bold/Italic 变 False，但 Underline 还能生效）。
    创建立即通过 `Text.Story` 复写这几个属性则稳定得多，所以这里做一次兜底。
    """
    story = _s(lambda: shape.Text.Story)
    if story is None:
        return {}
    done = {}
    if spec.get("font"):
        try:
            story.Font = str(spec["font"])
            done["font"] = str(_s(lambda: str(story.Font)))
        except Exception:
            pass
    if spec.get("size"):
        try:
            story.Size = float(spec["size"])
            done["size"] = _round(_s(lambda: float(story.Size)))
        except Exception:
            pass
    for key, prop in (("bold", "Bold"), ("italic", "Italic")):
        if key in spec and spec[key] is not None:
            try:
                setattr(story, prop, _tristate(spec[key]))
                done[key] = bool(_s(lambda prop=prop: bool(getattr(story, prop)), False))
            except Exception:
                pass
    if "underline" in spec and spec["underline"] is not None:
        try:
            story.Underline = _underline(spec["underline"])
            done["underline"] = bool(_s(lambda: bool(story.Underline), False))
        except Exception:
            pass
    return done


def _underline(value: Any) -> int:
    """下划线形参是 `cdrTextUnderlineStyle` 之类的枚举，不是 cdrTriState。

    实测：0 = 无，1~5 = 有（不同样式），-1 也等于“无”。所以只能用 1/0。
    """
    return 1 if value else 0


def _tristate(value: Any) -> int:
    """X8 的布尔型形参用 `cdrTriState`：`cdrTrue = -1`、`cdrFalse = 0`、`cdrUndefined = -2`。

    实测踩坑：`Text.Story.Bold = True` 会把粗体**关掉**（Python 的 1 不在枚举里），
    必须传 **-1**；创建文本时的 Bold/Italic/Underline 形参同理。
    """
    if value is None:
        return 0
    return -1 if bool(value) else 0


def _app_of(doc: Any) -> Any:
    """取 Application 工厂对象。

    实测教训：`doc.Application` 有时能拿到，但它派生的 CreateStructExportOptions
    交给 ExportEx 会失败（"发生意外"）。所以**优先用桥接持有的那个 Application**，
    只有拿不到时才退回 doc.Application。
    """
    from .com import get_bridge

    app = get_bridge().app
    if app is not None:
        return app
    try:
        app = doc.Application
        if app is not None:
            return app
    except Exception:
        pass
    raise CdrComFailure("无法取得 CorelDRAW Application 对象")


# --------------------------------------------------------------------------- #
# 文档
# --------------------------------------------------------------------------- #
def find_document(app: Any, name: str) -> Any:
    try:
        count = int(app.Documents.Count)
    except Exception as exc:  # noqa: BLE001
        raise CdrComFailure(f"无法读取文档列表：{exc}") from exc
    for i in range(1, count + 1):
        try:
            doc = app.Documents.Item(i)
        except Exception:
            continue
        if str(doc.Name) == name:
            return doc
    raise CdrDocumentNotFound(f"文档已关闭或被重命名：{name}")


def list_documents(app: Any) -> list[dict]:
    out: list[dict] = []
    try:
        count = int(app.Documents.Count)
    except Exception:
        return out
    for i in range(1, count + 1):
        try:
            doc = app.Documents.Item(i)
        except Exception:
            continue
        out.append(
            {
                "name": _s(lambda: str(doc.Name)),
                "file_name": _s(lambda: str(doc.FileName)) or None,
                "dirty": bool(_s(lambda: bool(doc.Dirty), False)),
                "pages": _n(lambda: int(doc.Pages.Count)),
                "page_width_mm": _round(_s(lambda: float(doc.ActivePage.SizeWidth))),
                "page_height_mm": _round(_s(lambda: float(doc.ActivePage.SizeHeight))),
                "unit": _s(lambda: int(doc.Unit)),
            }
        )
    return out


def create_document(app: Any, width: float | None, height: float | None,
                    orientation: str | None, preset: str | None = None,
                    options: dict | None = None) -> Any:
    """新建文档。`options` 会走 `CreateStructCreateOptions + CreateDocumentEx`
    （实测可一次定好 名称 / 页宽 / 页高 / 分辨率，见第三十八波）。"""
    if options:
        opts = app.CreateStructCreateOptions()
        if opts is None:
            raise CdrComFailure("CreateStructCreateOptions 返回空")
        if options.get("name"):
            opts.Name = str(options["name"])
        if options.get("width_mm"):
            opts.PageWidth = float(options["width_mm"])
        if options.get("height_mm"):
            opts.PageHeight = float(options["height_mm"])
        if options.get("resolution"):
            opts.Resolution = int(options["resolution"])
        if options.get("units") is not None:
            try:
                opts.Units = int(options["units"])
            except Exception:  # noqa: BLE001
                pass
        doc = app.CreateDocumentEx(opts)
        if doc is None:
            raise CdrComFailure("CreateDocumentEx 返回空")
        return doc
    doc = app.CreateDocument()
    with mm_units(doc):
        if preset:
            name = PAGE_PRESETS.get(str(preset).lower(), str(preset))
            try:
                doc.ActivePage.SelectSize(name, False)
            except Exception as exc:  # noqa: BLE001
                raise CdrBadArgument(f"未知页面预设 {preset!r}：{exc}") from exc
        elif width and height:
            if width <= 0 or height <= 0:
                raise CdrBadArgument("width/height 必须为正数（毫米）")
            doc.ActivePage.SetSize(float(width), float(height))
        if orientation:
            doc.ActivePage.Orientation = 2 if orientation.lower().startswith("land") else 1
    return doc


def open_document(app: Any, path: str, options: dict | None = None,
                  as_copy: bool = False) -> Any:
    """打开文档。`options` 走 `CreateStructOpenOptions + OpenDocumentEx`
    （结构体成员实测只有 CodePage / ColorConversionOptions）。
    `as_copy=True` 时走 `Application.OpenDocumentAsCopy(path, 打开选项)` ——
    返回值是**未命名副本**（FileName 为空），内容与源文件一致。"""
    if as_copy:
        opts = app.CreateStructOpenOptions()
        return app.OpenDocumentAsCopy(path, opts)
    if options:
        opts = app.CreateStructOpenOptions()
        if opts is None:
            raise CdrComFailure("CreateStructOpenOptions 返回空")
        if options.get("code_page") is not None:
            try:
                opts.CodePage = int(options["code_page"])
            except Exception:  # noqa: BLE001
                pass
        return app.OpenDocumentEx(path, opts)
    return app.OpenDocument(path)


def _save_options(app: Any, spec: dict | None) -> tuple[Any, list]:
    if not spec:
        return None, []
    opts = app.CreateStructSaveAsOptions()
    if opts is None:
        return None, []
    mapping = {
        "embed_icc_profile": "EmbedICCProfile",
        "embed_vba_project": "EmbedVBAProject",
        "include_cmx_data": "IncludeCMXData",
        "keep_appearance": "KeepAppearance",
        "overwrite": "Overwrite",
        "thumbnail_size": "ThumbnailSize",
        "version": "Version", "filter": "Filter", "range": "Range",
    }
    applied = []
    for key, val in spec.items():
        attr = mapping.get(str(key).lower())
        if attr is None:
            continue
        try:
            setattr(opts, attr, val)
            applied.append(str(key))
        except Exception:
            continue
    return opts, applied


def save_as(doc: Any, path: str, *, options: dict | None = None) -> dict:
    """X8 的 SaveAs 必须传第二个参数（Options），只传路径会抛 TypeError。"""
    old_name = str(doc.Name)
    app = _app_of(doc)
    opts, applied = _save_options(app, options)
    doc.SaveAs(path, opts)
    return {
        "name": _s(lambda: str(doc.Name), old_name),
        "file_name": _s(lambda: str(doc.FileName)),
        "old_name": old_name,
        "dirty": bool(_s(lambda: bool(doc.Dirty), False)),
        "save_options": applied,
    }


def undo(doc: Any, steps: int = 1, *, redo: bool = False) -> dict:
    n = max(1, int(steps))
    done = 0
    for _ in range(n):
        try:
            doc.Redo(1) if redo else doc.Undo()
            done += 1
        except Exception:
            break
    return {"redone" if redo else "undone": done}


def close_document(doc: Any, save: bool = False) -> dict:
    name = _s(lambda: str(doc.Name), "")
    file_name = _s(lambda: str(doc.FileName), "") or ""
    saved = False
    if save and file_name:
        try:
            doc.Save()
            saved = True
        except Exception:
            saved = False
    doc.Close()
    return {"closed": name, "saved": saved, "file_name": file_name or None}


def save_as_copy(doc: Any, path: str, *, options: dict | None = None) -> dict:
    """另存副本。

    ⚠️ 实测：X8 的 `SaveAsCopy` 会把**当前文档名也改成副本名**（a.cdr -> b.cdr，
    FileName 同步变），所以调用方必须用返回的 `name` 更新会话，否则后续按名查找会失败。
    """
    old_name = _s(lambda: str(doc.Name), "")
    app = _app_of(doc)
    opts, applied = _save_options(app, options)
    doc.SaveAsCopy(path, opts)
    return {
        "copied_to": path,
        "name": _s(lambda: str(doc.Name), old_name),
        "old_name": old_name,
        "file_name": _s(lambda: str(doc.FileName)) or None,
        "save_options": applied,
    }


VIEW_ACTIONS = ("refresh", "activate", "new_window", "overlay_handle", "overlay_points",
                "to_screen", "to_document",
                "zoom_fit_page", "zoom_fit_width", "zoom_fit_height",
                "zoom_all_objects", "zoom_selection", "actual_size", "zoom_in",
                "zoom_out", "zoom_in_at_point", "set_view_point", "view_area",
                "set_view_area", "fit_area", "fit_shape", "fit_shape_range",
                "distance_to_screen", "distance_to_document",
                "overlay_text", "overlay_rect", "overlay_line",
                "overlay_circle", "overlay_clear")
_ONSCREEN = {"items": []}


def view_ops(doc: Any, action: str, *, x: float | None = None, y: float | None = None,
             distance: float = 10.0, text: str | None = None,
             x2: float | None = None, y2: float | None = None,
             radius: float = 20.0, color: int = 0x0000FF,
             width: int = 2, target: dict | None = None,
             page_index: int | None = None) -> dict:
    """视图/窗口操作（IVGWindow，实测可用）。"""
    act = (action or "refresh").lower()
    if act not in VIEW_ACTIONS:
        raise CdrBadArgument(f"未知 action {action!r}；可用：{', '.join(VIEW_ACTIONS)}")
    if act in {"zoom_fit_page", "zoom_fit_width", "zoom_fit_height",
               "zoom_all_objects", "zoom_selection", "actual_size", "zoom_in",
               "zoom_out", "zoom_in_at_point", "set_view_point", "view_area",
               "set_view_area", "fit_area", "fit_shape", "fit_shape_range"}:
        # 实测：Application.ActiveWindow.ActiveView 是 IVGActiveView，
        # 15 个缩放/取景方法全部可用（Zoom 会随操作变化）
        app = _app_of(doc)
        view = _s(lambda: app.ActiveWindow.ActiveView)
        if view is None:
            raise CdrComFailure("取不到 ActiveView")
        before = _round(_s(lambda: float(view.Zoom)))
        if act == "zoom_fit_page":
            view.ToFitPage()
        elif act == "zoom_fit_width":
            view.ToFitPageWidth()
        elif act == "zoom_fit_height":
            view.ToFitPageHeight()
        elif act == "zoom_all_objects":
            view.ToFitAllObjects()
        elif act == "zoom_selection":
            view.ToFitSelection()
        elif act == "actual_size":
            view.SetActualSize()
        elif act == "zoom_in":
            view.ZoomIn()
        elif act == "zoom_out":
            view.ZoomOut()
        elif act == "zoom_in_at_point":
            view.ZoomInAtPoint(float(x or 0.0), float(y or 0.0))
        elif act == "set_view_point":
            view.SetViewPoint(float(x or 0.0), float(y or 0.0), 0)
        elif act == "view_area":
            area = _s(lambda: view.GetViewArea(0.0, 0.0, 0.0, 0.0))
            return {"action": act, "view_type": _n(lambda: int(view.Type)),
                    "zoom": before,
                    "area": ([_round(float(v)) for v in area]
                             if isinstance(area, (tuple, list)) else None)}
        elif act == "set_view_area":
            view.SetViewArea(float(x or 0.0), float(y or 0.0),
                             float(x2 or 100.0), float(y2 or 100.0))
        elif act == "fit_area":
            view.ToFitArea(float(x or 0.0), float(y or 0.0),
                           float(x2 or 100.0), float(y2 or 100.0))
        elif act == "fit_shape":
            view.ToFitShape(find_shape(doc, _page(doc, None), target or {}))
        else:
            rng = selection_range(doc)
            if _n(lambda: int(rng.Count), 0) < 1:
                raise CdrBadArgument("fit_shape_range 需要先选中对象")
            view.ToFitShapeRange(rng)
        return {"action": act, "zoom_before": before,
                "zoom_after": _round(_s(lambda: float(view.Zoom))),
                "view_type": _n(lambda: int(view.Type))}
    if act == "overlay_handle":
        # 实测：Application.CreateOnScreenHandle() → IVGOnScreenHandle；
        # SetPosition / SetHandleColor / Show / Hide / UpdateHotTracking 可用，IsHotTracked 可读
        spec = dict(params or {}) if "params" in dir() else {}
        app = _app_of(doc)
        handle = _s(lambda: app.CreateOnScreenHandle())
        if handle is None:
            raise CdrComFailure("CreateOnScreenHandle 返回空")
        cx = float(x if x is not None else 10.0)
        cy = float(y if y is not None else 10.0)
        applied = []
        with atomic(doc, "cdr_view:overlay_handle"):
            try:
                handle.SetPosition(cx, cy)
                applied.append("set_position")
            except Exception:
                pass
            try:
                handle.SetHandleColor(int(color))
                applied.append("set_handle_color")
            except Exception:
                pass
            try:
                handle.UpdateHotTracking(cx, cy)
                applied.append("update_hot_tracking")
            except Exception:
                pass
            try:
                if spec.get("hide"):
                    handle.Hide()
                    applied.append("hide")
                else:
                    handle.Show()
                    applied.append("show")
            except Exception:
                pass
        return {"action": act, "position": [cx, cy],
                "color": int(color),
                "hot_tracked": bool(_s(lambda: bool(handle.IsHotTracked), False)),
                "applied": applied}
    if act == "new_window":
        # 实测：Application.ActiveWindow.NewWindow() 可开新窗口（Document 上没有这个方法）
        win = _s(lambda: _app_of(doc).ActiveWindow.NewWindow())
        if win is None:
            raise CdrComFailure("ActiveWindow.NewWindow 返回空")
        return {"action": act, "created": True,
                "windows": _n(lambda: int(_app_of(doc).Windows.Count), 0)}
    if act == "overlay_clear":
        hidden = 0
        for obj in list(_ONSCREEN["items"]):
            try:
                obj.Hide()
                hidden += 1
            except Exception:
                pass
        _ONSCREEN["items"] = []
        return {"action": act, "hidden": hidden}
    if act == "overlay_points":
        # 实测：IVGOnScreenCurve 的 SetNoPen / SetBrush / SetPoints 都可用
        app = _app_of(doc)
        obj = app.CreateOnScreenCurve()
        points = [[float(x if x is not None else 50), float(y if y is not None else 50)],
                  [float(x2 if x2 is not None else 120), float(y2 if y2 is not None else 90)]]
        applied = []
        try:
            obj.SetNoPen()
            applied.append("set_no_pen")
        except Exception:
            pass
        try:
            obj.SetBrush(int(color))
            applied.append("set_brush")
        except Exception:
            pass
        try:
            # 实测：SetPoints 要的是 IVGPointRange（不是 Python list）
            obj.SetPoints(_make_point_range(app, points))
            applied.append("set_points")
        except Exception:
            pass
        obj.Show()
        _ONSCREEN["items"].append(obj)
        return {"action": act, "points": len(points), "applied": applied,
                "shown": len(_ONSCREEN["items"])}
    if act.startswith("overlay_"):
        app = _app_of(doc)
        if act == "overlay_text":
            obj = app.CreateOnScreenText()
            tx = float(x if x is not None else 100)
            ty = float(y if y is not None else 100)
            obj.SetTextAndPosition(str(text or "overlay"), tx, ty)
            obj.SetTextColor(int(color))
            # 实测：IVGOnScreenText 还有 SetText / SetPixelOffset / UpdatePosition
            try:
                obj.SetText(str(text or "overlay"))
            except Exception:
                pass
            try:
                obj.SetPixelOffset(int(distance), 0)
            except Exception:
                pass
            try:
                obj.UpdatePosition(tx, ty)
            except Exception:
                pass
        else:
            obj = app.CreateOnScreenCurve()
            obj.SetPen(int(color), int(width), 0)
            obj.SetNoBrush()
            ox = float(x if x is not None else 50)
            oy = float(y if y is not None else 50)
            ex = float(x2 if x2 is not None else ox + 100)
            ey = float(y2 if y2 is not None else oy + 60)
            if act == "overlay_rect":
                obj.SetRectangle(ox, oy, ex, ey)
            elif act == "overlay_line":
                obj.SetLine(ox, oy, ex, ey)
            else:  # overlay_circle
                obj.SetCircle(ox, oy, float(radius))
        obj.Show()
        _ONSCREEN["items"].append(obj)
        return {"action": act, "shown": len(_ONSCREEN["items"]),
                "screen_coords": True}
    win = _s(lambda: doc.ActiveWindow)
    if win is None:
        # 实测：按名取回的 Document 偶尔取不到 ActiveWindow，退到 Application 层
        app = _app_of(doc)
        win = _s(lambda: app.ActiveWindow)
        if win is None:
            wins = _s(lambda: app.Windows)
            if wins is not None and _n(lambda: int(wins.Count), 0) > 0:
                win = _s(lambda: wins.Item(1))
    if win is None:
        raise CdrBadArgument(
            "当前没有活动窗口（文档可能未激活）；可先调用 cdr_pages(action='activate')")

    if act == "refresh":
        win.Refresh()
        return {"action": act, "refreshed": True}
    if act == "activate":
        win.Activate()
        return {"action": act, "activated": True}
    if act == "to_screen":
        if x is None or y is None:
            raise CdrBadArgument("to_screen 需要 x 与 y（文档坐标）")
        pt = _s(lambda: tuple(win.DocumentToScreen(float(x), float(y))))
        return {"action": act, "screen": [_round(v) for v in pt] if pt else None}
    if act == "to_document":
        if x is None or y is None:
            raise CdrBadArgument("to_document 需要 x 与 y（屏幕坐标）")
        pt = _s(lambda: tuple(win.ScreenToDocument(float(x), float(y))))
        return {"action": act, "document": [_round(v) for v in pt] if pt else None}
    if act == "distance_to_document":
        val = _s(lambda: float(win.ScreenDistanceToDocumentDistance(float(distance))))
        return {"action": act, "screen_distance": distance,
                "document_distance": _round(val)}
    val = _s(lambda: float(win.DocumentDistanceToScreenDistance(float(distance))))
    return {"action": act, "document_distance": distance, "screen_distance": _round(val)}


DOC_ACTIONS = ("info", "clear_selection", "clear_undo", "palette", "arrowhead",
               "units", "sample_color", "show_all", "layout", "repeat",
               "fonts", "refresh_fonts", "palettes", "palette_color",
               "palette_open_fixed", "palette_colors", "spot_color",
               "pattern_canvas", "textures", "metadata", "set_metadata",
               "palette_create", "data_fields", "add_data_field", "object_data",
               "color_context", "outline_styles", "settings", "set_settings",
               "revert", "app_info", "color_profiles", "convert_color",
               "pref_register", "pref_get", "pref_set",
               "views", "create_view", "activate_view", "delete_view",
               "grid", "set_grid", "overprint",
               "metadata_localized", "set_metadata_localized",
               # 第三十五波：单位换算（COM）/ 文档设置快照 / 解析位图链接
               "convert_units", "save_settings", "restore_settings",
               "reset_settings", "resolve_links", "arrowhead_named",
               # 第四十二波：颜色上下文指派 / 样式表设为默认 / 箭头变体
               "assign_color_context", "stylesheet_default", "arrowhead_ex",
               # 第四十四波：颜色上下文深操作 / 渐变变换矩阵
               "color_context_info", "color_context_copy", "color_context_merge",
               "color_context_create", "proof_color_settings",
               "app_status", "palette_manager",
               "outline_style_dash", "tool_shape_attributes", "tree_manager",
               "set_color_profile", "palette_from", "windows",
               "select_color_profile",
               "fountain_transform", "arrowhead_unnamed",
               # 第四十六波：CorelScriptTools（VBA 辅助对象）单位/数学/日期
               "script_tools")


def doc_ops(doc: Any, action: str, *, target: dict | None = None,
            page_index: int | None = None, value: float | None = None,
            from_unit: int | None = None, to_unit: int | None = None,
            x: float | None = None, y: float | None = None,
            x2: float | None = None, y2: float | None = None,
            x_samples: int = 4, y_samples: int = 4,
            layout: int | None = None, query: str | None = None,
            limit: int = 100, palette_index: int | None = None,
            color_index: int = 1, palette_id: int | None = None,
            spot_name: str | None = None, tint: float = 100.0,
            pattern_fills: list | None = None, pattern_pixels: list | None = None,
            pattern_lines: list | None = None, clear_canvas: bool = True,
            metadata: dict | None = None, palette_name: str | None = None,
            palette_path: str | None = None, palette_colors: list | None = None,
            overwrite: bool = False, code_page: int | None = None,
            field_name: str | None = None,
            field_format: str = "0", settings: dict | None = None,
            register: bool = False, copy_areas: list | None = None,
            rotate_areas: list | None = None, flip_areas: list | None = None,
            model: int | str | None = None, profile_name: str | None = None,
            rendering_intent: int | None = None,
            pref_group: str = "cdrX8 MCP", pref_key: str | None = None,
            pref_value: Any = None, pref_default: Any = None,
            view_name: str | None = None, view_index: int = 1,
            zoom: float = 100.0, origin_x: float = 0.0,
            origin_y: float = 0.0,
            grid_spacing_x: float | None = None,
            grid_spacing_y: float | None = None,
            grid_visible: bool | None = None,
            grid_snap: bool | None = None,
            grid_type: int | None = None) -> dict:
    """文档级操作。

    action:
      info             文档概况（路径、脏标记、页数、单位）
      clear_selection  清空选择集
      clear_undo       清空撤销栈（批处理前后释放内存）
      palette          把文档用到的颜色登记进调色板（AddColorsToDocPalette）
      arrowhead        用一条曲线生成自定义箭头（返回箭头对象可用信息）
      units            单位换算：doc.ToUnits / doc.FromUnits（value/from_unit/to_unit）
      sample_color     取色：SampleColorAtPoint（x,y）或 SampleColorInArea（x,y,x2,y2）
      show_all         显示所有隐藏对象（ShowAllHiddenObjects）
      layout           页面版式（SetLayout：layout 传数字枚举）
      repeat           重复上一操作（Document.Repeat，等同 Ctrl+R）
    """
    act = (action or "info").lower()
    if act not in DOC_ACTIONS:
        raise CdrBadArgument(f"未知 action {action!r}；可用：{', '.join(DOC_ACTIONS)}")

    if act == "info":
        return {
            "action": act,
            "name": _s(lambda: str(doc.Name)),
            "file_name": _s(lambda: str(doc.FileName)) or None,
            "dirty": bool(_s(lambda: bool(doc.Dirty), False)),
            "pages": _n(lambda: int(doc.Pages.Count)),
            "unit": _n(lambda: int(doc.Unit)),
            "shapes": total_shapes(doc),
        }
    if act == "clear_selection":
        doc.ClearSelection()
        return {"action": act, "selected": _n(lambda: int(doc.Selection().Shapes.Count), 0)}
    if act == "clear_undo":
        doc.ClearUndoList()
        return {"action": act, "cleared": True}
    if act == "palette":
        doc.AddColorsToDocPalette(False, 0)
        return {"action": act, "done": True}
    if act == "convert_units":
        # 实测 Application.ConvertUnits(Value, FromUnit, ToUnit) 用 cdrUnit 枚举：
        # inch=1, foot=2, mm=3, cm=4, pixel=5, meter=7, point=14 …
        if value is None or from_unit is None or to_unit is None:
            raise CdrBadArgument("convert_units 需要 value / from_unit / to_unit"
                                 "（cdrUnit：inch=1 mm=3 cm=4 point=14）")
        got = _s(lambda: float(_app_of(doc).ConvertUnits(float(value),
                                                         int(from_unit), int(to_unit))))
        if got is None:
            raise CdrComFailure("Application.ConvertUnits 返回空")
        return {"action": act, "value": float(value), "from_unit": int(from_unit),
                "to_unit": int(to_unit), "result": _round(got), "source": "com"}
    if act in {"save_settings", "restore_settings", "reset_settings"}:
        # 实测：SaveSettings(Tag) / RestoreSettings(Tag) / ResetSettings() 都返回 2 个值（无异常），
        # 用于把当前文档设置打快照，再恢复/重置。
        tag = str(query or "mcp_snapshot")
        if act == "save_settings":
            doc.SaveSettings(tag)
        elif act == "restore_settings":
            doc.RestoreSettings(tag)
        else:
            doc.ResetSettings()
        return {"action": act, "tag": tag if act != "reset_settings" else None,
                "unit": _n(lambda: int(doc.Unit)),
                "resolution": _n(lambda: int(doc.Resolution)),
                "edit_across_layers": bool(_s(
                    lambda: bool(doc.EditAcrossLayers), False))}
    if act == "resolve_links":
        doc.ResolveAllBitmapsLinks()
        return {"action": act, "resolved": True,
                "shapes": total_shapes(doc)}
    if act == "units":
        if value is None:
            raise CdrBadArgument("units 需要 value")
        out: dict = {"action": act, "value": float(value), "doc_unit": _n(lambda: int(doc.Unit))}
        if from_unit is not None:
            out["to_document"] = _round(_s(lambda: float(doc.ToUnits(float(value),
                                                                   int(from_unit)))))
        if to_unit is not None:
            out["from_document"] = _round(_s(lambda: float(doc.FromUnits(float(value),
                                                                       int(to_unit)))))
        if from_unit is None and to_unit is None:
            out["to_document"] = _round(_s(lambda: float(doc.ToUnits(float(value),
                                                                   _n(lambda: int(doc.Unit), 1)))))
        return out
    if act == "sample_color":
        if x is None or y is None:
            raise CdrBadArgument("sample_color 需要 x 与 y")
        with mm_units(doc):
            if x2 is not None and y2 is not None:
                col = _s(lambda: doc.SampleColorInArea(float(x), float(y), float(x2),
                                                       float(y2), int(x_samples),
                                                       int(y_samples), 99))
            else:
                col = _s(lambda: doc.SampleColorAtPoint(float(x), float(y), 99))
        if col is None:
            return {"action": act, "sampled": False, "color": None}
        return {"action": act, "sampled": True,
                "color": _s(lambda: str(col.ToString())),
                "name": _s(lambda: str(col.GetName())) or None}
    if act == "show_all":
        doc.ShowAllHiddenObjects()
        return {"action": act, "done": True}
    if act == "layout":
        if layout is None:
            raise CdrBadArgument("layout 需要 layout（数字枚举）")
        doc.SetLayout(int(layout), False, True)
        return {"action": act, "layout": int(layout),
                "pages": _n(lambda: int(doc.Pages.Count))}
    if act == "repeat":
        doc.Repeat()
        return {"action": act, "done": True,
                "shapes": total_shapes(doc)}
    if act == "refresh_fonts":
        app = _app_of(doc)
        _FONT_TABLE["refreshed"] = False
        ok = ensure_font_table(app)
        return {"action": act, "refreshed": ok,
                "fonts": _n(lambda: int(app.FontList.Count), 0)}
    if act == "fonts":
        app = _app_of(doc)
        total = _n(lambda: int(app.FontList.Count), 0)
        needle = str(query or "").strip().lower()
        cap = max(1, min(int(limit), 2000))
        # FontList.Item 每次都是一次 COM 调用（本机 2000+ 字体时全量扫描要 1~2 分钟），
        # 所以按需扫描：默认最多扫 max(300, limit*10) 条，返回 scanned/truncated 说明。
        scan_cap = min(total, max(300, cap * 10))
        names: list[str] = []
        scanned = 0
        for i in range(1, scan_cap + 1):
            try:
                nm = str(app.FontList.Item(i))
            except Exception:
                continue
            scanned += 1
            if not needle or needle in nm.lower():
                names.append(nm)
                if len(names) >= cap:
                    break
        return {"action": act, "total": total, "matched": len(names),
                "fonts": names, "scanned": scanned,
                "truncated": len(names) >= cap or scanned < total,
                "note": "按需扫描；要更多结果请增大 limit（全量扫描较慢）"}
    if act == "palettes":
        app = _app_of(doc)
        count = _n(lambda: int(app.Palettes.Count), 0)
        items = []
        for i in range(1, count + 1):
            palette = app.Palettes.Item(i)
            items.append({
                "index": i,
                "name": _s(lambda p=palette: str(p.Name)),
                "identifier": _s(lambda p=palette: str(p.Identifier)),
                "palette_id": _n(lambda p=palette: int(p.PaletteID)),
                "color_count": _n(lambda p=palette: int(p.ColorCount)),
                "default": bool(_s(lambda p=palette: bool(p.Default), False)),
                "type": _n(lambda p=palette: int(p.Type)),
            })
        return {"action": act, "count": count, "palettes": items}
    if act == "palette_color":
        app = _app_of(doc)
        pi = int(palette_index if palette_index is not None else 1)
        ci = int(color_index)
        count = _n(lambda: int(app.Palettes.Count), 0)
        if not 1 <= pi <= count:
            raise CdrBadArgument(f"palette_index 越界：{pi}（共 {count} 个调色板）")
        palette = app.Palettes.Item(pi)
        color_count = _n(lambda: int(palette.ColorCount), 0)
        if not 1 <= ci <= color_count:
            raise CdrBadArgument(f"color_index 越界：{ci}（该调色板共 {color_count} 色）")
        color = palette.Color(ci)
        return {"action": act, "palette": _s(lambda: str(palette.Name)),
                "palette_index": pi, "color_index": ci,
                "color": _s(lambda: str(color.ToString())),
                "name": _s(lambda: str(color.GetName())) or None}
    if act == "palette_open_fixed":
        if palette_id is None:
            raise CdrBadArgument("palette_open_fixed 需要 palette_id（本机实测 3=PANTONE 匹配系统）")
        app = _app_of(doc)
        pal = app.Palettes.OpenFixed(int(palette_id))
        if pal is None:
            raise CdrBadArgument(f"OpenFixed({palette_id}) 返回空（该固定调色板可能未安装）")
        count = _n(lambda: int(app.Palettes.Count), 0)
        index = None
        ident = _s(lambda: str(pal.Identifier), "")
        for i in range(1, count + 1):
            if str(_s(lambda i=i: str(app.Palettes.Item(i).Identifier), "")) == ident:
                index = i
                break
        return {"action": act, "palette_id": int(palette_id),
                "name": _s(lambda: str(pal.Name)),
                "identifier": ident or None,
                "color_count": _n(lambda: int(pal.ColorCount)),
                "palette_index": index, "opened": count,
                "note": "固定调色板编号因安装内容不同可能变化；打开后用 palettes/palette_colors 确认"}
    if act == "palette_colors":
        pal = _find_palette(doc, palette_index if palette_index is not None else 1)
        total = _n(lambda: int(pal.ColorCount), 0)
        needle = str(query or "").strip().lower()
        cap = max(1, min(int(limit), 2000))
        colors: list[dict] = []
        for i in range(1, total + 1):
            try:
                color = pal.Color(i)
                name = str(color.GetName())
            except Exception:
                continue
            if needle and needle not in name.lower():
                continue
            colors.append({"index": i, "name": name,
                           "color": _s(lambda c=color: str(c.ToString()))})
            if len(colors) >= cap:
                break
        return {"action": act, "palette": _s(lambda: str(pal.Name)),
                "palette_index": palette_index, "total": total,
                "matched": len(colors), "colors": colors,
                "truncated": len(colors) >= cap}
    if act == "spot_color":
        app = _app_of(doc)
        pi = int(palette_index if palette_index is not None else 1)
        pal = _find_palette(doc, pi)
        name = str(spot_name) if spot_name else None
        if not name:
            ci = int(color_index)
            color_count = _n(lambda: int(pal.ColorCount), 0)
            if not 1 <= ci <= color_count:
                raise CdrBadArgument(f"color_index 越界：{ci}（该调色板共 {color_count} 色）")
            name = _s(lambda: str(pal.Color(ci).GetName()))
        color = app.CreateSpotColorByName(
            str(_s(lambda: str(pal.Identifier), "")), str(name), float(tint))
        if color is None:
            raise CdrBadArgument(f"专色创建失败：{name!r}（名称需与调色板里的完全一致）")
        return {"action": act, "palette": _s(lambda: str(pal.Name)),
                "palette_index": pi, "spot_name": name, "tint": float(tint),
                "color": _s(lambda: str(color.ToString())),
                "name": _s(lambda: str(color.GetName())) or None}
    if act == "pattern_canvas":
        app = _app_of(doc)
        canvas = app.CreatePatternCanvas()
        if clear_canvas:
            canvas.Clear()
        drawn = 0
        for area in (pattern_fills or []):
            if isinstance(area, dict):
                x1, y1 = int(area.get("x1", area.get("x", 0))), int(area.get("y1", area.get("y", 0)))
                x2, y2 = int(area.get("x2", x1 + area.get("width", 1))), \
                    int(area.get("y2", y1 + area.get("height", 1)))
                state = bool(area.get("state", True))
            elif isinstance(area, (list, tuple)):
                vals = list(area) + [True]
                x1, y1, x2, y2 = (int(v) for v in vals[:4])
                state = bool(vals[4])
            else:
                continue
            canvas.FillArea(x1, y1, x2, y2, state)
            drawn += 1
        for line in (pattern_lines or []):
            if not isinstance(line, (list, tuple)) or len(line) < 4:
                continue
            vals = list(line)
            flags = int(vals[0]) if len(vals) >= 5 else 0
            if len(vals) >= 5:
                lx1, ly1, lx2, ly2 = (int(v) for v in vals[1:5])
            else:
                lx1, ly1, lx2, ly2 = (int(v) for v in vals[:4])
            canvas.Line(flags, lx1, ly1, lx2, ly2, True)
            drawn += 1
        for px in (pattern_pixels or []):
            if not isinstance(px, (list, tuple)) or len(px) < 2:
                continue
            state = bool(px[2]) if len(px) > 2 else True
            canvas.SetPixel(int(px[0]), int(px[1]), state)
            drawn += 1
        for area in (copy_areas or []):
            if not isinstance(area, (list, tuple)) or len(area) < 6:
                continue
            canvas.CopyArea(int(area[0]), int(area[1]), int(area[2]),
                            int(area[3]), int(area[4]), int(area[5]))
            drawn += 1
        for area in (rotate_areas or []):
            if not isinstance(area, (list, tuple)) or len(area) < 5:
                continue
            canvas.RotateArea(int(area[0]), int(area[1]), int(area[2]),
                              int(area[3]), float(area[4]))
            drawn += 1
        for area in (flip_areas or []):
            if not isinstance(area, (list, tuple)) or len(area) < 5:
                continue
            canvas.FlipArea(int(area[0]), int(area[1]), int(area[2]),
                            int(area[3]), int(area[4]))
            drawn += 1
        index = int(app.PatternCanvases.Add(canvas))
        return {"action": act, "canvas_index": index,
                "width": _n(lambda: int(canvas.Width)),
                "height": _n(lambda: int(canvas.Height)),
                "drawn": drawn,
                "canvases": _n(lambda: int(app.PatternCanvases.Count)),
                "note": "用 cdr_modify(fill={\"type\":\"pattern\",\"canvas_index\":N,...}) 应用"}
    if act == "textures":
        app = _app_of(doc)
        app_path = _s(lambda: str(app.Path), "") or ""
        root = os.path.abspath(os.path.join(app_path, os.pardir, "Custom Data", "Textures")) \
            if app_path else ""
        names: list[str] = []
        if root and os.path.isdir(root):
            for fn in sorted(os.listdir(root)):
                if fn.lower().endswith((".tif", ".tiff", ".bmp", ".jpg", ".jpeg", ".png")):
                    names.append(os.path.splitext(fn)[0])
        needle = str(query or "").strip().lower()
        matched = [n for n in names if not needle or needle in n.lower()]
        cap = max(1, min(int(limit), 2000))
        return {"action": act, "root": root or None, "total": len(names),
                "matched": len(matched), "textures": matched[:cap],
                "truncated": len(matched) > cap,
                "note": "用 fill={\"type\":\"texture\",\"name\":<名称>,\"library\":\"Textures\"}"}
    if act in {"metadata", "set_metadata"}:
        md = _s(lambda: doc.Metadata)
        if md is None:
            raise CdrBadArgument("该文档没有 Metadata 对象")
        read_fields = ("Title", "Author", "LastAuthor", "Subject", "Keywords",
                       "Notes", "Copyright", "Revision", "DocLanguage", "DocID")
        writable = {"Title", "Author", "Subject", "Keywords", "Notes", "Copyright"}
        applied: list[str] = []
        warnings: list[str] = []
        if act == "set_metadata":
            if not isinstance(metadata, dict) or not metadata:
                raise CdrBadArgument("set_metadata 需要 metadata 对象")
            for key, val in metadata.items():
                if key not in writable:
                    warnings.append(f"{key} 不可写（可用：{', '.join(sorted(writable))}）")
                    continue
                try:
                    setattr(md, key, str(val))
                    applied.append(key)
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"{key} 写入失败：{exc}")
        values = {f: _s(lambda f=f: str(getattr(md, f))) for f in read_fields}
        result = {"action": act, "metadata": values, "writable": sorted(writable)}
        if act == "set_metadata":
            result["applied"] = applied
            result["warnings"] = warnings
        return result
    if act in {"metadata_localized", "set_metadata_localized"}:
        md = _s(lambda: doc.Metadata)
        if md is None:
            raise CdrBadArgument("该文档没有 Metadata 对象")
        fields = {"Title": "LocalizableTitle", "Subject": "LocalizableSubject",
                  "Keywords": "LocalizableKeywords", "Notes": "LocalizableNotes",
                  "Copyright": "LocalizableCopyright"}
        applied: list[str] = []
        warnings: list[str] = []
        if act == "set_metadata_localized":
            if not isinstance(metadata, dict) or not metadata:
                raise CdrBadArgument("set_metadata_localized 需要 metadata 对象")
            for key, val in metadata.items():
                attr = fields.get(key)
                if attr is None:
                    warnings.append(f"{key} 不可写（可用：{', '.join(fields)}）")
                    continue
                ls = _s(lambda attr=attr: getattr(md, attr))
                if ls is None:
                    warnings.append(f"{key} 没有本地化字符串对象")
                    continue
                try:
                    if isinstance(val, dict):
                        for lang, txt in val.items():
                            if str(lang).lower() in {"default", "0", ""}:
                                ls.DefaultLangString = str(txt)
                            else:
                                ls.SetLangString(int(lang), str(txt))
                    else:
                        ls.DefaultLangString = str(val)
                    applied.append(key)
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"{key} 写入失败：{exc}")
        values = {}
        for key, attr in fields.items():
            ls = _s(lambda attr=attr: getattr(md, attr))
            if ls is None:
                values[key] = None
                continue
            langs = _s(lambda ls=ls: list(ls.GetLanguages()), []) or []
            lang_map = {}
            for lang in langs:
                name = str(lang)
                if name in {"", "0"}:
                    continue
                lang_map[name] = _s(lambda ls=ls, lang=lang: str(ls.GetLangString(int(lang))))
            values[key] = {
                "default": _s(lambda ls=ls: str(ls.DefaultLangString)),
                "has_default": bool(_s(lambda ls=ls: bool(ls.HasDefaultLangString), False)),
                "languages": lang_map,
            }
        result = {"action": act, "localizable": values}
        if act == "set_metadata_localized":
            result["applied"] = applied
            result["warnings"] = warnings
        return result
    if act == "palette_create":
        if not palette_name or not palette_path:
            raise CdrBadArgument("palette_create 需要 palette_name 与 palette_path")
        app = _app_of(doc)
        pal = app.Palettes.Create(str(palette_name), str(palette_path), bool(overwrite))
        if pal is None:
            raise CdrBadArgument("Palettes.Create 返回空（检查路径/名称）")
        added: list = []
        add_warnings: list[str] = []
        for spec in (palette_colors or []):
            try:
                pal.AddColor(_color(doc, spec))
                added.append(spec)
            except Exception as exc:  # noqa: BLE001
                add_warnings.append(f"{spec!r}: {exc}")
        pal.SaveAs(str(palette_path))
        ident = _s(lambda: str(pal.Identifier), "")
        count = _n(lambda: int(app.Palettes.Count), 0)
        index = None
        for i in range(1, count + 1):
            if str(_s(lambda i=i: str(app.Palettes.Item(i).Identifier), "")) == ident:
                index = i
                break
        return {"action": act, "name": _s(lambda: str(pal.Name)),
                "identifier": ident or None, "palette_index": index,
                "color_count": _n(lambda: int(pal.ColorCount)),
                "added": len(added), "path": str(palette_path),
                "exists": os.path.exists(str(palette_path)),
                "bytes": os.path.getsize(str(palette_path))
                if os.path.exists(str(palette_path)) else 0,
                "warnings": add_warnings}
    if act in {"data_fields", "add_data_field"}:
        fields = doc.DataFields
        if act == "add_data_field":
            if not field_name:
                raise CdrBadArgument("add_data_field 需要 field_name")
            field = fields.Add(str(field_name), str(field_format), False, False, False)
            if field is None:
                raise CdrBadArgument("数据字段创建失败（格式字符串无效？实测 '0' 可用）")
            return {"action": act,
                    "field": {"name": _s(lambda: str(field.Name)),
                              "data_type": _n(lambda: int(field.DataType)),
                              "default_value": _s(lambda: str(field.DefaultValue))},
                    "count": _n(lambda: int(fields.Count))}
        count = _n(lambda: int(fields.Count), 0)
        items = []
        for i in range(1, min(count, 300) + 1):
            try:
                f = fields.Item(i)
            except Exception:
                continue
            items.append({"name": _s(lambda f=f: str(f.Name)),
                          "data_type": _n(lambda f=f: int(f.DataType)),
                          "default_value": _s(lambda f=f: str(f.DefaultValue))})
        return {"action": act, "count": count, "fields": items}
    if act == "object_data":
        page = _page(doc, page_index)
        shape = find_shape(doc, page, target or {})
        od = shape.ObjectData
        count = _n(lambda: int(od.Count), 0)
        items = []
        for i in range(1, min(count, 200) + 1):
            try:
                it = od.Item(i)
            except Exception:
                continue
            items.append({"field": _s(lambda it=it: str(it.DataField.Name)),
                          "value": _s(lambda it=it: str(it.Value)),
                          "formatted": _s(lambda it=it: str(it.FormattedValue))})
        return {"action": act, "target": _s(lambda: str(shape.Name)),
                "count": count, "data": items}
    if act == "overprint":
        page = _page(doc, page_index)
        shape = find_shape(doc, page, target or {})
        return {"action": act, "target": _s(lambda: str(shape.Name)),
                "overprint_fill": bool(_s(lambda: bool(shape.OverprintFill), False)),
                "overprint_outline": bool(_s(lambda: bool(shape.OverprintOutline), False)),
                "overprint_bitmap": bool(_s(lambda: bool(shape.OverprintBitmap), False))}
    if act == "color_context":
        cc = _s(lambda: doc.ColorContext)
        if cc is None:
            raise CdrBadArgument("该文档没有 ColorContext")
        return {"action": act,
                "read_only": bool(_s(lambda: bool(cc.ReadOnly), False)),
                "rendering_intent": _n(lambda: int(cc.RenderingIntent)),
                "blending_color_model": _n(lambda: int(cc.BlendingColorModel)),
                "profile_count": _n(lambda: int(cc.ColorProfiles.Count), 0),
                "profile_names": _s(lambda: str(cc.ColorProfileNameList)) or ""}
    if act == "outline_styles":
        app = _app_of(doc)
        count = _n(lambda: int(app.OutlineStyles.Count), 0)
        items = []
        for i in range(1, min(count, 300) + 1):
            try:
                st = app.OutlineStyles.Item(i)
            except Exception:
                continue
            items.append({"index": _n(lambda st=st: int(st.Index)),
                          "dash_count": _n(lambda st=st: int(st.DashCount)),
                          "enhanced": bool(_s(lambda st=st: bool(st.Enhanced), False))})
        return {"action": act, "count": count, "styles": items}
    if act in {"settings", "set_settings"}:
        spec = {
            "edit_across_layers": "EditAcrossLayers",
            "apply_to_duplicate": "ApplyToDuplicate",
            "curve_precision": "CurvePrecision",
            "resolution": "Resolution",
            "preserve_selection": "PreserveSelection",
            "shape_enum_direction": "ShapeEnumDirection",
            "reference_point": "ReferencePoint",
            "text_formatter": "TextFormatter",
            "world_scale": "WorldScale",
            "facing_pages": "FacingPages",
            "first_page_on_right_side": "FirstPageOnRightSide",
            "layout": "Layout",
            "unit": "Unit",
            "drawing_origin_x": "DrawingOriginX",
            "drawing_origin_y": "DrawingOriginY",
            "dirty": "Dirty",
        }
        writable = {"edit_across_layers", "apply_to_duplicate", "curve_precision",
                    "resolution", "preserve_selection", "shape_enum_direction",
                    "world_scale", "text_formatter", "facing_pages",
                    "first_page_on_right_side"}
        applied: list[str] = []
        warnings: list[str] = []
        if act == "set_settings":
            if not isinstance(settings, dict) or not settings:
                raise CdrBadArgument("set_settings 需要 settings 对象")
            for key, val in settings.items():
                if key not in spec or key not in writable:
                    warnings.append(f"{key} 不可写（可用：{', '.join(sorted(writable))}）")
                    continue
                try:
                    setattr(doc, spec[key], val)
                    applied.append(key)
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"{key} 写入失败：{exc}")
        values = {}
        for key, attr in spec.items():
            try:
                values[key] = getattr(doc, attr)
            except Exception:
                values[key] = None
        result = {"action": act, "settings": values, "writable": sorted(writable)}
        if act == "set_settings":
            result["applied"] = applied
            result["warnings"] = warnings
        return result
    if act == "revert":
        doc.Revert()
        return {"action": act, "reverted": True,
                "name": _s(lambda: str(doc.Name)),
                "file_name": _s(lambda: str(doc.FileName)) or None,
                "dirty": bool(_s(lambda: bool(doc.Dirty), False)),
                "pages": _n(lambda: int(doc.Pages.Count)),
                "shapes": total_shapes(doc)}
    if act == "app_info":
        app = _app_of(doc)
        features = _s(lambda: list(app.GetSupportedOpenTypeFeatures()), []) or []
        needle = str(query or "").strip().lower()
        matched = [str(f) for f in features if not needle or needle in str(f).lower()]
        cap = max(1, min(int(limit), 400))
        return {"action": act,
                "open_type_features": matched[:cap],
                "feature_count": len(features),
                "feature_matched": len(matched),
                "active_tool": _s(lambda: str(app.ActiveTool)),
                "active_tool_state": _s(lambda: str(app.ActiveToolStateGuid)),
                "platform_checks": {str(v): bool(_s(
                    lambda v=v: bool(app.CheckPlatformVersion(int(v), 0)), False))
                    for v in (16, 17, 18)},
                "ui_language": {str(lang): bool(_s(
                    lambda lang=lang: bool(app.IsUILanguageAvailable(int(lang))), False))
                    for lang in (1033, 2052)}}
    if act in {"color_profiles", "convert_color"}:
        app = _app_of(doc)
        cm = _s(lambda: app.ColorManager)
        if cm is None:
            raise CdrBadArgument("Application.ColorManager 不可用")
        def _norm_model(value):
            if value is None:
                return None
            if isinstance(value, str):
                aliases = {"rgb": 0, "cmyk": 1, "gray": 2, "grayscale": 2,
                           "grey": 2, "lab": 3}
                key = value.strip().lower()
                if key in aliases:
                    return aliases[key]
                raise CdrBadArgument("model 可用：0=RGB, 1=CMYK, 2=Grayscale, 3=Lab")
            return int(value)
        model_id = _norm_model(model)
        if act == "color_profiles":
            if model_id is None:
                coll = _s(lambda: cm.ColorProfiles)
            else:
                coll = _s(lambda: cm.GetProfilesByColorModel(model_id))
            count = _n(lambda: int(coll.Count), 0) if coll is not None else 0
            cap = max(1, min(int(limit), 400))
            items = []
            for i in range(1, min(count, cap) + 1):
                p = _s(lambda i=i: coll.Item(i))
                if p is None:
                    continue
                items.append({
                    "name": _s(lambda p=p: str(p.Name)),
                    "model": _n(lambda p=p: int(p.ColorModel)),
                    "generic": bool(_s(lambda p=p: bool(p.Generic), False)),
                    "installed": bool(_s(lambda p=p: bool(p.Installed), False)),
                    "manufacturer": _s(lambda p=p: str(p.Manufacturer)),
                    "file_name": _s(lambda p=p: str(p.FileName)) or None,
                })
            default_ctx = _s(lambda: cm.DefaultColorContext)
            return {"action": act, "count": count, "model": model_id,
                    "profiles": items,
                    "default_context": _s(
                        lambda: str(default_ctx.ColorProfileNameList))
                    if default_ctx is not None else None}
        # convert_color
        target = None
        if profile_name:
            needle = str(profile_name).strip().casefold()
            all_profiles = _s(lambda: cm.ColorProfiles)
            count = _n(lambda: int(all_profiles.Count), 0) if all_profiles is not None else 0
            for i in range(1, count + 1):
                p = _s(lambda i=i: all_profiles.Item(i))
                if p is None:
                    continue
                name = str(_s(lambda p=p: str(p.Name), ""))
                if name.casefold() == needle:
                    target = p
                    break
                if target is None and needle in name.casefold():
                    target = p
        if target is None:
            if model_id is None:
                raise CdrBadArgument("convert_color 需要 model 或 profile_name")
            coll = _s(lambda: cm.GetProfilesByColorModel(model_id))
            count = _n(lambda: int(coll.Count), 0) if coll is not None else 0
            if count < 1:
                raise CdrBadArgument(f"没有 model={model_id} 的已安装配置文件")
            prefer = {0: ("srgb iec61966-2.1", "adobe rgb (1998)"),
                      1: ("japan color 2001 coated", "coated fogra39"),
                      2: ("dot gain 15%", "dot gain 20%")}.get(model_id, ())
            for needle in prefer:
                for i in range(1, count + 1):
                    p = _s(lambda i=i: coll.Item(i))
                    if p is None:
                        continue
                    if needle in str(_s(lambda p=p: str(p.Name), "")).casefold():
                        target = p
                        break
                if target is not None:
                    break
            if target is None:
                target = coll.Item(1)
        if target is None:
            raise CdrBadArgument("找不到指定配置文件")
        context = target.CreateColorContext()
        if context is None:
            raise CdrBadArgument("CreateColorContext 返回空")
        if rendering_intent is not None:
            context.RenderingIntent = int(rendering_intent)
        before = _s(lambda: str(doc.ColorContext.ColorProfileNameList))
        doc.ConvertToColorContext(context)
        after = _s(lambda: str(doc.ColorContext.ColorProfileNameList))
        return {"action": act,
                "profile": _s(lambda: str(target.Name)),
                "model": _n(lambda: int(target.ColorModel)),
                "rendering_intent": _n(lambda: int(context.RenderingIntent)),
                "before": before, "after": after}
    if act in {"pref_register", "pref_get", "pref_set"}:
        if not pref_key:
            raise CdrBadArgument("pref_* 需要 pref_key")
        app = _app_of(doc)
        group = str(pref_group or "cdrX8 MCP")
        key = str(pref_key)
        if act == "pref_register":
            default = "" if pref_default is None else pref_default
            app.RegisterUserApplicationPreference(group, key, default)
            return {"action": act, "group": group, "key": key,
                    "default": default,
                    "value": _s(lambda: app.GetApplicationPreferenceValue(group, key))}
        if act == "pref_get":
            try:
                value = app.GetApplicationPreferenceValue(group, key)
            except Exception:
                if pref_default is None:
                    raise CdrBadArgument(
                        f"偏好 {group}/{key} 未注册；先 pref_register 或传 pref_default")
                app.RegisterUserApplicationPreference(group, key, pref_default)
                value = app.GetApplicationPreferenceValue(group, key)
            return {"action": act, "group": group, "key": key,
                    "value": _s(lambda: value)}
        # pref_set
        try:
            app.SetApplicationPreferenceValue(group, key, pref_value)
        except Exception:
            default = "" if pref_default is None else pref_default
            app.RegisterUserApplicationPreference(group, key, default)
            app.SetApplicationPreferenceValue(group, key, pref_value)
        return {"action": act, "group": group, "key": key,
                "value": _s(lambda: app.GetApplicationPreferenceValue(group, key))}
    if act in {"views", "create_view", "activate_view", "delete_view"}:
        views = _s(lambda: doc.Views)
        if views is None:
            raise CdrBadArgument("该文档没有 Views")
        def _view_dict(view, index):
            return {"index": index,
                    "name": _s(lambda: str(view.Name)),
                    "origin_x": _s(lambda: float(view.OriginX)),
                    "origin_y": _s(lambda: float(view.OriginY)),
                    "zoom": _s(lambda: float(view.Zoom)),
                    "use_zoom": bool(_s(lambda: bool(view.UseZoom), False)),
                    "use_page": bool(_s(lambda: bool(view.UsePage), False)),
                    "page": _n(lambda: int(view.Page.Index))}
        if act == "views":
            count = _n(lambda: int(views.Count), 0)
            return {"action": act, "count": count,
                    "views": [_view_dict(views.Item(i), i)
                              for i in range(1, min(count, 200) + 1)]}
        if act == "create_view":
            if not view_name:
                raise CdrBadArgument("create_view 需要 view_name")
            page = _page(doc, page_index)
            view = doc.CreateView(str(view_name), float(origin_x),
                                  float(origin_y), float(zoom), page)
            if view is None:
                raise CdrBadArgument("CreateView 返回空")
            count = _n(lambda: int(views.Count), 0)
            return {"action": act, "created": str(view_name),
                    "count": count,
                    "views": [_view_dict(views.Item(i), i)
                              for i in range(1, min(count, 200) + 1)]}
        idx = max(1, int(view_index))
        count = _n(lambda: int(views.Count), 0)
        if idx > count:
            raise CdrBadArgument(f"视图索引越界：{idx}（共 {count} 个）")
        view = views.Item(idx)
        if act == "activate_view":
            view.Activate()
            return {"action": act, "activated": idx,
                    "name": _s(lambda: str(view.Name))}
        view.Delete()
        return {"action": act, "deleted": idx,
                "count": _n(lambda: int(views.Count), 0)}
    if act in {"grid", "set_grid"}:
        with mm_units(doc):
            grid = _s(lambda: doc.Grid)
            if grid is None:
                raise CdrBadArgument("该文档没有 Grid")
            if act == "set_grid":
                # 实测：Document.Grid 是 IVGGrid，SetFrequency(x, y) 按"每单位多少格"设置
                if isinstance(settings, dict) and settings.get("frequency"):
                    freq = settings["frequency"]
                    fx, fy = (freq if isinstance(freq, (list, tuple)) and len(freq) >= 2
                              else (freq, freq))
                    grid.SetFrequency(float(fx), float(fy))
                    return {"action": act, "frequency": [float(fx), float(fy)],
                            "spacing_x": _round(_s(lambda: float(grid.SpacingX))),
                            "spacing_y": _round(_s(lambda: float(grid.SpacingY))),
                            "type": _n(lambda: int(grid.Type)),
                            "visible": bool(_s(lambda: bool(grid.Visible), False)),
                            "snap": bool(_s(lambda: bool(grid.Snap), False))}
                if grid_spacing_x is not None:
                    grid.SpacingX = float(grid_spacing_x)
                if grid_spacing_y is not None:
                    grid.SpacingY = float(grid_spacing_y)
                if grid_visible is not None:
                    grid.Visible = bool(grid_visible)
                if grid_snap is not None:
                    grid.Snap = bool(grid_snap)
                if grid_type is not None:
                    grid.Type = int(grid_type)
            return {"action": act,
                    "type": _n(lambda: int(grid.Type)),
                    "spacing_x": _s(lambda: float(grid.SpacingX)),
                    "spacing_y": _s(lambda: float(grid.SpacingY)),
                    "visible": bool(_s(lambda: bool(grid.Visible), False)),
                    "snap": bool(_s(lambda: bool(grid.Snap), False))}
    # arrowhead_named（第三十八波：Document.CreateArrowHead2 可给箭头命名）
    if act == "arrowhead_unnamed":
        # 实测：Document.CreateArrowHeadEx(Curve, cx, cy, tint, ?) 可用，但生成箭头 Name 为空
        with mm_units(doc):
            page = _page(doc, page_index)
            shape = find_shape(doc, page, target or {})
            curve = _s(lambda: shape.Curve)
            if curve is None:
                raise CdrBadArgument("arrowhead_unnamed 需要一条曲线对象")
            cx = float(x if x is not None else 0.0)
            cy = float(y if y is not None else 0.0)
            head = _s(lambda: doc.CreateArrowHeadEx(curve, cx, cy,
                                                    float(tint or 1.0), 0.0))
            if head is None:
                raise CdrComFailure("Document.CreateArrowHeadEx 返回空")
            return {"action": act, "created": True,
                    "name": _s(lambda: str(head.Name)),
                    "center": [cx, cy],
                    "note": "CreateArrowHeadEx 生成箭头 Name 为空；要命名用 arrowhead_ex（CreateArrowHeadEx2）"}
    if act == "arrowhead_named":
        with mm_units(doc):
            page = _page(doc, page_index)
            shape = find_shape(doc, page, target or {})
            curve = _s(lambda: shape.Curve)
            if curve is None:
                raise CdrBadArgument("arrowhead_named 需要一条曲线对象")
            head_name = str(query or "MCP 箭头")
            head = _s(lambda: doc.CreateArrowHead2(curve, head_name))
            if head is None:
                raise CdrComFailure("Document.CreateArrowHead2 返回空")
            app = _app_of(doc)
            before = _n(lambda: int(app.ArrowHeads.Count), 0)
            registered = False
            if register:
                app.ArrowHeads.Add(head)
                registered = True
            after = _n(lambda: int(app.ArrowHeads.Count), before)
            return {"action": act, "created": True, "name": head_name,
                    "curve": _s(lambda: str(shape.Name)),
                    "registered": registered, "arrowheads": after,
                    "index": after if registered else None}
    if act == "script_tools":
        # 实测：Application.CorelScriptTools() 返回 ICorelScriptTools（VBA 辅助对象），
        # FromInches(1)=254000（1/10000 mm）、FromPoints(72)=254000 等可直接用
        tools = _s(lambda: _app_of(doc).CorelScriptTools())
        if tools is None:
            raise CdrComFailure("Application.CorelScriptTools 返回空")
        val = float(value if value is not None else 1.0)
        # FromX 把英寸/磅转成内部单位（1/10000 mm）；ToX 把内部单位转回去
        inches_internal = _s(lambda: tools.FromInches(val))
        points_internal = _s(lambda: tools.FromPoints(val))
        out = {"action": act, "value": val,
               "from_inches": _round(_s(lambda: float(inches_internal))),
               "from_points": _round(_s(lambda: float(points_internal))),
               "to_inches_roundtrip": (_round(_s(lambda: float(
                   tools.ToInches(inches_internal))))
                   if inches_internal is not None else None),
               "to_points_roundtrip": (_round(_s(lambda: float(
                   tools.ToPoints(points_internal))))
                   if points_internal is not None else None)}
        if -1.0 <= val <= 1.0:
            # 实测 ASin 返回**弧度**（ASin(1)=1.5708）
            out["asin_rad"] = _round(_s(lambda: float(tools.ASin(val))))
        stamp = _s(lambda: tools.BuildDate(2026, 10, 26, 12, 0, 0))
        if stamp is None:
            stamp = _s(lambda: tools.BuildDate(2026, 10, 26))
        out["build_date"] = (_round(_s(lambda: float(stamp)))
                             if isinstance(stamp, (int, float)) else
                             (str(stamp) if stamp is not None else None))
        if code_page is not None:
            app = _app_of(doc)
            out["code_page"] = int(code_page)
            if query:
                # ConvertToUnicode(字节数组) ↔ ConvertFromUnicode(字符串)
                try:
                    out["unicode_from_bytes"] = str(app.ConvertToUnicode(
                        bytes.fromhex(str(query)), int(code_page)))
                except Exception as exc:  # noqa: BLE001
                    out["unicode_from_bytes_error"] = str(exc)[:80]
                try:
                    out["unicode_to_bytes"] = str(app.ConvertFromUnicode(
                        str(query), int(code_page)))
                except Exception as exc:  # noqa: BLE001
                    out["unicode_to_bytes_error"] = str(exc)[:80]
        out["note"] = ("CorelScriptTools 的值以 1/10000 mm 为单位：FromInches(1)=254000；"
                       "ASin 返回弧度；query+code_page 可做 Unicode↔ANSI 码页转换")
        return out
    if act == "select_color_profile":
        # 实测：IVGColorProfiles.SelectByName(名称) 可调用（找不到/不匹配时返回 False）
        app = _app_of(doc)
        profiles = _s(lambda: app.ColorManager.ColorProfiles)
        if profiles is None:
            raise CdrComFailure("取不到 ColorProfiles")
        target_name = str(profile_name or "sRGB IEC61966-2.1")
        ok = bool(_s(lambda: bool(profiles.SelectByName(target_name)), False))
        return {"action": act, "requested": target_name, "ok": ok,
                "count": _n(lambda: int(profiles.Count), 0)}

    if act in {"palette_from", "windows"}:
        app = _app_of(doc)
        if act == "palette_from":
            # 实测：Application.Palettes → IVGPalettes；
            # CreateFromDocument(名) / CreateFromSelection(名) / Open(路径) 可用
            pals = _s(lambda: app.Palettes)
            if pals is None:
                raise CdrComFailure("取不到 Palettes")
            spec = dict(settings or {})
            name = str(spec.get("name") or "MCP 调色板")
            made = None
            mode = str(spec.get("mode") or "document").lower()
            if mode == "selection":
                made = _s(lambda: pals.CreateFromSelection(name))
                if made is None:
                    # 没有选中对象时 CreateFromSelection 会失败，退回整篇文档
                    made = _s(lambda: pals.CreateFromDocument(name))
                    mode = "document(fallback)"
            else:
                made = _s(lambda: pals.CreateFromDocument(name))
            if made is None:
                # 同名调色板已存在时创建会失败，换个唯一名字再试一次
                uniq = f"{name} {_short_id()}"
                if mode.startswith("selection"):
                    made = _s(lambda: pals.CreateFromSelection(uniq))
                if made is None:
                    made = _s(lambda: pals.CreateFromDocument(uniq))
                if made is not None:
                    name = uniq
            result = {"action": act, "mode": mode, "name": name,
                      "created": made is not None}
            if made is not None:
                result["color_count"] = _n(lambda: int(made.ColorCount))
                result["palette_name"] = _s(lambda: str(made.Name))[:60]
                result["index"] = _n(lambda: int(made.Index))
            if spec.get("open_path"):
                # 打开已有调色板文件（.cpl/.xml 等，路径由调用方给出）
                opened = _s(lambda: pals.Open(str(spec["open_path"])))
                result["opened"] = opened is not None
                if opened is not None:
                    result["opened_count"] = _n(lambda: int(opened.ColorCount))
            return result
        # windows
        wins = _s(lambda: app.Windows)
        if wins is None:
            raise CdrComFailure("取不到 Windows 集合")
        count = _n(lambda: int(wins.Count), 0) or 0
        spec = dict(settings or {})
        result = {"action": act, "count": count}
        if spec.get("arrange") is not None:
            # 实测：Windows.Arrange(mode) 可调用（0/1/2 都接受）
            wins.Arrange(int(spec["arrange"]))
            result["arranged"] = int(spec["arrange"])
        if spec.get("close_all"):
            wins.CloseAll()
            result["closed_all"] = True
        if spec.get("find_caption"):
            found = _s(lambda: wins.FindWindow(str(spec["find_caption"])))
            result["found"] = found is not None
            result["caption"] = (_s(lambda: str(found.Caption))
                                 if found is not None else None)
        return result

    if act == "set_color_profile":
        # 实测：ColorContext.SetColorProfile(模型号, 配置文件) 可用，
        # 但**配置文件的 ColorModel 必须与模型号一致**（否则报「颜色模型不匹配」）
        app = _app_of(doc)
        ctx = _s(lambda: doc.ColorContext)
        if ctx is None:
            raise CdrComFailure("Document.ColorContext 返回空")
        want_model = int(model) if model is not None else 1
        profiles = _s(lambda: app.ColorManager.ColorProfiles)
        count = _n(lambda: int(profiles.Count), 0) or 0
        chosen = None
        chosen_name = None
        for i in range(1, count + 1):
            prof = _s(lambda i=i: profiles.Item(i))
            if prof is None:
                continue
            if profile_name and _s(lambda prof=prof: str(prof.Name)).lower().startswith(
                    str(profile_name).lower()):
                chosen = prof
                chosen_name = _s(lambda prof=prof: str(prof.Name))
                break
            if int(_s(lambda prof=prof: int(prof.ColorModel), -1)) == want_model:
                chosen = prof
                chosen_name = _s(lambda prof=prof: str(prof.Name))
                if not profile_name:
                    break
        if chosen is None:
            raise CdrBadArgument(f"没有 ColorModel={want_model} 的配置文件")
        ctx.SetColorProfile(want_model, chosen)
        return {"action": act, "model": want_model,
                "profile": (chosen_name or "")[:60],
                "readback": _s(lambda: str(ctx.ColorProfile(want_model).Name))[:60],
                "rendering_intent": _n(lambda: int(ctx.RenderingIntent))}

    if act == "tree_manager":
        # 实测：Document.TreeManager → IVGTreeManager
        # （SelectedNodeCount / FirstSelectedNode / VirtualLayer 可读，
        #   CreateGroupNode() / CleanGroupNode(node) 可用）
        tm = _s(lambda: doc.TreeManager)
        if tm is None:
            raise CdrComFailure("取不到 TreeManager")
        spec = dict(settings or {})
        applied = []
        node = None
        if spec.get("create_group_node"):
            node = _s(lambda: tm.CreateGroupNode())
            if node is not None:
                applied.append("create_group_node")
        if node is not None and spec.get("clean"):
            tm.CleanGroupNode(node)
            applied.append("clean_group_node")
        vlayer = _s(lambda: tm.VirtualLayer)
        first = _s(lambda: tm.FirstSelectedNode)
        return {"action": act,
                "selected_node_count": _n(lambda: int(tm.SelectedNodeCount), 0),
                "first_selected": (([(_s(lambda: str(first.Name)) or "")[:60]]
                                    if first is not None else None)),
                "virtual_layer": ((_s(lambda: str(vlayer.Name)) or "")[:60]
                                  if vlayer is not None else None),
                "node": _s(lambda: str(node))[:60] if node is not None else None,
                "applied": applied}

    if act in {"outline_style_dash", "tool_shape_attributes"}:
        app = _app_of(doc)
        if act == "outline_style_dash":
            # 实测：Application.OutlineStyles.Item(n) 是 IVGOutlineStyle。
            # DashCount / DashLength(i) / GapLength(i) / Enhanced 都能**读**；
            # 写入（DashCount= / DashLength=(i, v)）调用不报错但不生效
            styles = _s(lambda: app.OutlineStyles)
            if styles is None:
                raise CdrComFailure("取不到 OutlineStyles")
            count = _n(lambda: int(styles.Count), 0) or 0
            idx = int(palette_index or 1)
            if not 1 <= idx <= count:
                raise CdrBadArgument(f"线型序号越界：{idx}（共 {count} 个）")
            style = styles.Item(idx)
            # 实测：SetDashLength(i, v) / SetGapLength(i, v) 可调（但读回可能仍是 0）
            applied = []
            spec = dict(settings or {})
            if spec.get("dash_length") is not None:
                dash_count = _n(lambda: int(style.DashCount), 1) or 1
                for k in range(1, dash_count + 1):
                    try:
                        style.SetDashLength(k, float(spec["dash_length"]))
                    except Exception:
                        continue
                applied.append("set_dash_length")
            if spec.get("gap_length") is not None:
                dash_count = _n(lambda: int(style.DashCount), 1) or 1
                for k in range(1, dash_count + 1):
                    try:
                        style.SetGapLength(k, float(spec["gap_length"]))
                    except Exception:
                        continue
                applied.append("set_gap_length")
            dashes = _n(lambda: int(style.DashCount), 0) or 0
            segments = []
            for k in range(1, min(dashes, 20) + 1):
                segments.append({
                    "dash": _round(_s(lambda k=k: float(style.DashLength(k)))),
                    "gap": _round(_s(lambda k=k: float(style.GapLength(k))),)})
            return {"action": act, "index": idx,
                    "style_count": count,
                    "index_of_style": _n(lambda: int(style.Index)),
                    "dash_count": dashes,
                    "enhanced": bool(_s(lambda: bool(style.Enhanced), False)),
                    "segments": segments,
                    "applied": applied if 'applied' in dir() else [],
                    "note": "DashLength/GapLength 用 (序号) 取；"
                            "SetDashLength/SetGapLength 可调但读回可能仍为 0"}
        # tool_shape_attributes：造一个自定义工具形状属性对象并逐项设置
        # （12 个 Set* 都是 IVGToolShapeAttributes 的方法，实测全部可用）
        attrs = _s(lambda: app.CreateToolShapeAttributes())
        if attrs is None:
            raise CdrComFailure("CreateToolShapeAttributes 返回空")
        spec = dict(settings or {})
        applied: list[str] = []
        guid_zero = "00000000-0000-0000-0000-000000000000"

        def _mark(name: str) -> None:
            applied.append(name)

        try:
            attrs.SetCanResize(bool(spec.get("can_resize", True)))
            _mark("CanResize")
        except Exception:
            pass
        try:
            attrs.SetCanRotate(bool(spec.get("can_rotate", True)))
            _mark("CanRotate")
        except Exception:
            pass
        try:
            attrs.SetCanSkew(bool(spec.get("can_skew", True)))
            _mark("CanSkew")
        except Exception:
            pass
        try:
            attrs.SetCanSizeDisproportionally(
                bool(spec.get("can_size_disproportionally", True)))
            _mark("CanSizeDisproportionally")
        except Exception:
            pass
        try:
            attrs.SetCanApplyNonlinearTransforms(
                bool(spec.get("can_apply_nonlinear", True)))
            _mark("CanApplyNonlinearTransforms")
        except Exception:
            pass
        try:
            attrs.SetRegenerateOnTransform(
                bool(spec.get("regenerate_on_transform", True)))
            _mark("RegenerateOnTransform")
        except Exception:
            pass
        try:
            attrs.SetRegenerateOnStyleChange(
                bool(spec.get("regenerate_on_style_change", True)))
            _mark("RegenerateOnStyleChange")
        except Exception:
            pass
        try:
            attrs.SetPropertyBarGuid(str(spec.get("property_bar_guid", guid_zero)))
            _mark("PropertyBarGuid")
        except Exception:
            pass
        try:
            attrs.SetContextMenuGuid(str(spec.get("context_menu_guid", guid_zero)))
            _mark("ContextMenuGuid")
        except Exception:
            pass
        try:
            attrs.SetObjectManagerBitmapGuid(
                str(spec.get("object_manager_guid", guid_zero)))
            _mark("ObjectManagerBitmapGuid")
        except Exception:
            pass
        try:
            attrs.SetEditStateGuid(str(spec.get("edit_state_guid", guid_zero)))
            _mark("EditStateGuid")
        except Exception:
            pass
        try:
            attrs.SetDefaultShapename(str(spec.get("default_shape_name", "MCP 形状")))
            _mark("DefaultShapename")
        except Exception:
            pass
        return {"action": act, "applied": applied, "count": len(applied),
                "note": "自定义工具形状的属性对象；X8 没有可用的 RegisterToolShape，"
                        "所以此处只做配置、不注册"}

    if act in {"app_status", "palette_manager"}:
        app = _app_of(doc)
        if act == "palette_manager":
            # 实测：Application.PaletteManager → IVGPaletteManager
            # （PaletteCount / GetPalette(index) / LoadAllPalettes / LoadPalette）
            pm = _s(lambda: app.PaletteManager)
            if pm is None:
                raise CdrComFailure("取不到 PaletteManager")
            applied = []
            if settings and settings.get("load_all"):
                pm.LoadAllPalettes()
                applied.append("load_all_palettes")
            if palette_index is not None:
                try:
                    pm.LoadPalette(int(palette_index))
                    applied.append("load_palette")
                except Exception:
                    pass
            count = _n(lambda: int(pm.PaletteCount), 0) or 0
            names = []
            for i in range(1, min(count, 10) + 1):
                pal = _s(lambda i=i: pm.GetPalette(i))
                if pal is not None:
                    names.append(_s(lambda pal=pal: str(pal.Name))[:40])
            return {"action": act, "palette_count": count,
                    "first_palettes": names, "applied": applied,
                    "open_palettes": _n(lambda: int(pm.OpenPalettes.Count), 0)}
        # app_status：进度条（实测 BeginProgress/UpdateProgress/SetProgressMessage/
        # EndProgress 四个都可用，Progress 会随 UpdateProgress 变化）
        st = _s(lambda: app.Status)
        if st is None:
            raise CdrComFailure("取不到 Application.Status")
        spec = dict(settings or {})
        applied = []
        if spec.get("begin") is not None:
            st.BeginProgress(int(spec.get("max", 10)), 0)
            applied.append("begin")
        if spec.get("update") is not None:
            st.UpdateProgress(int(spec["update"]))
            applied.append("update")
        if spec.get("message") is not None:
            st.SetProgressMessage(int(spec["message"]))
            applied.append("message")
        progress = _n(lambda: int(st.Progress))
        aborted = bool(_s(lambda: bool(st.Aborted), False))
        if spec.get("end"):
            st.EndProgress()
            applied.append("end")
        return {"action": act, "progress": progress, "aborted": aborted,
                "applied": applied,
                "note": "进度条 API；BeginProgress(上限, 0) 后 Progress 从 -1 变 0"}

    if act == "color_context_create":
        # 实测：Application.CreateColorContext2(配置名串, RenderingIntent, BlendingColorModel)
        # —— 配置名串直接用 ColorContext.ColorProfileNameList（逗号分隔）；**RenderingIntent 有效范围 1~4**
        app = _app_of(doc)
        names = _s(lambda: str(doc.ColorContext.ColorProfileNameList))
        if not names:
            raise CdrComFailure("取不到 ColorProfileNameList")
        intent = int(rendering_intent if rendering_intent is not None else 2)
        model_id = int(model if isinstance(model, int) else 1)
        if not 1 <= intent <= 4:
            raise CdrBadArgument("RenderingIntent 有效范围是 1~4（0 会报越界）")
        ctx = _s(lambda: app.CreateColorContext2(str(names), intent, model_id))
        if ctx is None:
            raise CdrComFailure("CreateColorContext2 返回空")
        return {"action": act, "profile_names": str(names)[:300],
                "rendering_intent": _n(lambda: int(ctx.RenderingIntent)),
                "blending_model": _n(lambda: int(ctx.BlendingColorModel)),
                "profiles": _n(lambda: int(ctx.ColorProfiles.Count), 0),
                "read_only": bool(_s(lambda: bool(ctx.ReadOnly), False))}
    if act == "proof_color_settings":
        # 实测：CreateProofColorSettings(配置名, RenderingIntent(1~4), 保留颜色值) 返回
        # IVGProofColorSettings（ColorContext / ShowOutOfGamutWarning / OutOfGamutColor …）
        app = _app_of(doc)
        profile = str(profile_name or "Japan Color 2001 Coated")
        intent = int(rendering_intent if rendering_intent is not None else 1)
        if not 1 <= intent <= 4:
            raise CdrBadArgument("RenderingIntent 有效范围是 1~4")
        proof = _s(lambda: app.CreateProofColorSettings(
            profile, intent, bool(settings.get("preserve_color_values", False))
            if isinstance(settings, dict) else False))
        if proof is None:
            raise CdrComFailure("CreateProofColorSettings 返回空")
        ctx = _s(lambda: proof.ColorContext)
        return {"action": act, "profile": profile, "rendering_intent": intent,
                "preserve_color_values": bool(_s(lambda: bool(proof.PreserveColorValues),
                                                 False)),
                "show_out_of_gamut": bool(_s(lambda: bool(proof.ShowOutOfGamutWarning),
                                             False)),
                "out_of_gamut_transparency": _n(
                    lambda: int(proof.OutOfGamutTransparency)),
                "context_rendering_intent": (_n(lambda: int(ctx.RenderingIntent))
                                             if ctx is not None else None),
                "context_profiles": (_n(lambda: int(ctx.ColorProfiles.Count), 0)
                                     if ctx is not None else 0)}
    if act == "color_context_info":
        ctx = _s(lambda: doc.ColorContext)
        if ctx is None:
            raise CdrComFailure("Document.ColorContext 返回空")
        return {"action": act,
                "rendering_intent": _n(lambda: int(ctx.RenderingIntent)),
                "blending_model": _n(lambda: int(ctx.BlendingColorModel)),
                "profiles": _n(lambda: int(ctx.ColorProfiles.Count), 0),
                "profile_names": _s(lambda: str(ctx.ColorProfileNameList))[:400],
                "read_only": bool(_s(lambda: bool(ctx.ReadOnly), False)),
                "rgb_profile": _s(lambda: str(ctx.RGBColorProfile.Name))[:120],
                "cmyk_profile": _s(lambda: str(ctx.CMYKColorProfile.Name))[:120],
                "gray_profile": _s(lambda: str(ctx.GrayscaleColorProfile.Name))[:120],
                # 实测：ColorProfile(模型号) 0=RGB / 1=CMYK / 2=灰度
                "profiles_by_model": {
                    str(m): (_s(lambda m=m: str(ctx.ColorProfile(m).Name)) or "")[:80]
                    for m in (0, 1, 2)}}
    if act in {"color_context_copy", "color_context_merge"}:
        ctx = _s(lambda: doc.ColorContext)
        if ctx is None:
            raise CdrComFailure("Document.ColorContext 返回空")
        if act == "color_context_copy":
            clone = _s(lambda: ctx.GetCopy())
            if clone is None:
                raise CdrComFailure("ColorContext.GetCopy 返回空")
            return {"action": act,
                    "same_as_source": bool(_s(lambda: bool(clone.IsSame(ctx)), False)),
                    "rendering_intent": _n(lambda: int(clone.RenderingIntent)),
                    "profiles": _n(lambda: int(clone.ColorProfiles.Count), 0)}
        # merge：取指定配置文件建上下文再 Merge 进文档上下文
        app = _app_of(doc)
        idx = int(palette_index or 1)
        count = _n(lambda: int(app.ColorManager.ColorProfiles.Count), 0) or 0
        if not 1 <= idx <= count:
            raise CdrBadArgument(f"profile_index 越界：{idx}（共 {count} 个）")
        prof = app.ColorManager.ColorProfiles.Item(idx)
        other = _s(lambda: prof.CreateColorContext())
        if other is None:
            raise CdrComFailure("CreateColorContext 返回空")
        merged = bool(_s(lambda: bool(ctx.Merge(other)), False))
        return {"action": act, "profile_index": idx,
                "profile": _s(lambda: str(prof.Name))[:120],
                "merged": merged,
                "same_after": bool(_s(lambda: bool(ctx.IsSame(other)), False)),
                "rendering_intent": _n(lambda: int(ctx.RenderingIntent)),
                "profile_names": _s(lambda: str(ctx.ColorProfileNameList))[:400]}
    if act == "fountain_transform":
        with mm_units(doc):
            page = _page(doc, page_index)
            shape = find_shape(doc, page, target or {})
            fountain = _s(lambda: shape.Fill.Fountain)
            if fountain is None:
                raise CdrBadArgument("目标不是渐变填充（先用 cdr_modify 应用渐变）")
            mat = _s(lambda: fountain.GetTransformations())
            return {"action": act, "fill_type": _n(lambda: int(shape.Fill.Type)),
                    "transformations": ([_round(float(v)) for v in mat]
                                        if isinstance(mat, (tuple, list)) else None),
                    "note": "GetTransformations 用 4 个出参返回 [a, b, c, d] 变换矩阵"}
    if act == "assign_color_context":
        # 实测：ColorManager.ColorProfiles.Item(i).CreateColorContext() →
        # Document.AssignColorContext(ctx) 可把文档色彩上下文换成该配置文件
        app = _app_of(doc)
        idx = int(palette_index or 1)
        count = _n(lambda: int(app.ColorManager.ColorProfiles.Count), 0) or 0
        if not 1 <= idx <= count:
            raise CdrBadArgument(f"profile_index 越界：{idx}（共 {count} 个配置文件）")
        prof = app.ColorManager.ColorProfiles.Item(idx)
        ctx = _s(lambda: prof.CreateColorContext())
        if ctx is None:
            raise CdrComFailure("ColorProfile.CreateColorContext 返回空")
        doc.AssignColorContext(ctx)
        return {"action": act, "profile_index": idx,
                "profile": _s(lambda: str(prof.Name))[:120],
                "color_model": _n(lambda: int(prof.ColorModel)),
                "rendering_intent": _n(
                    lambda: int(doc.ColorContext.RenderingIntent))}
    if act == "stylesheet_default":
        doc.SaveStyleSheetAsDefault()
        return {"action": act, "saved": True,
                "styles": _n(lambda: int(doc.StyleSheet.Styles.Count), 0)}
    if act == "arrowhead_ex":
        with mm_units(doc):
            page = _page(doc, page_index)
            shape = find_shape(doc, page, target or {})
            curve = _s(lambda: shape.Curve)
            if curve is None:
                raise CdrBadArgument("arrowhead_ex 需要一条曲线对象")
            head_name = str(query or "MCP 箭头")
            cx = float(x if x is not None else 0.0)
            cy = float(y if y is not None else 0.0)
            head = _s(lambda: doc.CreateArrowHeadEx2(curve, head_name, cx, cy,
                                                     float(tint or 1.0), 0.0))
            if head is None:
                raise CdrComFailure("Document.CreateArrowHeadEx2 返回空")
            app = _app_of(doc)
            before = _n(lambda: int(app.ArrowHeads.Count), 0)
            registered = False
            if register:
                app.ArrowHeads.Add(head)
                registered = True
            after = _n(lambda: int(app.ArrowHeads.Count), before)
            # 顺带读一下 options 对象（CreateArrowHeadOptions）
            opts = _s(lambda: doc.CreateArrowHeadOptions(float(cx), float(cy),
                                                         1.0, 0.0, 1.0, False, False))
            return {"action": act, "created": True, "name": head_name,
                    "center": [cx, cy], "registered": registered,
                    "arrowheads": after, "index": after if registered else None,
                    "options_length": _round(_s(lambda: float(opts.Length)))
                    if opts is not None else None,
                    "options_offset": ([_round(_s(lambda: float(opts.OffsetX))),
                                        _round(_s(lambda: float(opts.OffsetY)))]
                                       if opts is not None else None)}
    # arrowhead
    with mm_units(doc):
        page = _page(doc, page_index)
        shape = find_shape(doc, page, target or {})
        curve = _s(lambda: shape.Curve)
        if curve is None:
            raise CdrBadArgument("arrowhead 需要一条曲线对象")
        head = doc.CreateArrowHead(curve)
        app = _app_of(doc)
        before = _n(lambda: int(app.ArrowHeads.Count), 0)
        registered = False
        if register and head is not None:
            app.ArrowHeads.Add(head)
            registered = True
        after = _n(lambda: int(app.ArrowHeads.Count), before)
        return {"action": act, "created": head is not None,
                "registered": registered, "arrowheads": after,
                "index": after if registered else None}


# --------------------------------------------------------------------------- #
# 页面
# --------------------------------------------------------------------------- #
def page_summary(doc: Any) -> list[dict]:
    pages = []
    with mm_units(doc):
        count = _n(lambda: int(doc.Pages.Count), 0)
        active = _n(lambda: int(doc.ActivePage.Index), 0)
        for i in range(1, count + 1):
            page = doc.Pages.Item(i)
            pages.append(
                {
                    "index": i,
                    "active": i == active,
                    "width_mm": _round(_s(lambda: float(page.SizeWidth))),
                    "height_mm": _round(_s(lambda: float(page.SizeHeight))),
                    "orientation": ORIENTATION.get(_n(lambda: int(page.Orientation), 0), "unknown"),
                    "shapes": _n(lambda: int(page.Shapes.Count), 0),
                    "layers": [{"name": n} for n in _layer_names(page)],
                }
            )
    return pages


def _layer_names(page: Any) -> list[str]:
    names: list[str] = []
    try:
        for j in range(1, int(page.Layers.Count) + 1):
            try:
                names.append(str(page.Layers.Item(j).Name))
            except Exception:
                continue
    except Exception:
        pass
    return names


def _page(doc: Any, index: int | None) -> Any:
    return doc.ActivePage if index is None else doc.Pages.Item(int(index))


def set_page_size(doc: Any, width: float, height: float, page_index: int | None = None) -> dict:
    with mm_units(doc):
        page = _page(doc, page_index)
        page.SetSize(float(width), float(height))
        return {
            "index": _n(lambda: int(page.Index)),
            "width_mm": _round(_s(lambda: float(page.SizeWidth))),
            "height_mm": _round(_s(lambda: float(page.SizeHeight))),
        }


def set_page_preset(doc: Any, preset: str, landscape: bool = False,
                    page_index: int | None = None) -> dict:
    name = PAGE_PRESETS.get(str(preset).lower(), str(preset))
    with mm_units(doc):
        page = _page(doc, page_index)
        try:
            page.SelectSize(name, bool(landscape))
        except Exception as exc:  # noqa: BLE001
            raise CdrBadArgument(
                f"未知页面预设 {preset!r}；可用：{', '.join(sorted(PAGE_PRESETS))}"
            ) from exc
        return {
            "index": _n(lambda: int(page.Index)),
            "preset": name,
            "width_mm": _round(_s(lambda: float(page.SizeWidth))),
            "height_mm": _round(_s(lambda: float(page.SizeHeight))),
        }


def activate_page(doc: Any, index: int) -> dict:
    doc.Pages.Item(int(index)).Activate()
    return {"active": _n(lambda: int(doc.ActivePage.Index))}


def add_pages(doc: Any, count: int = 1, width: float | None = None,
              height: float | None = None) -> dict:
    """加页。给 width+height 时走 `AddPagesEx`（一次把新页尺寸定好）。"""
    if width is not None and height is not None:
        with mm_units(doc):
            made = doc.AddPagesEx(max(1, int(count)), float(width), float(height))
        return {"pages": _n(lambda: int(doc.Pages.Count)),
                "width": float(width), "height": float(height),
                "sized": True, "returned": made is not None}
    doc.AddPages(max(1, int(count)))
    return {"pages": _n(lambda: int(doc.Pages.Count)), "sized": False}


def insert_pages(doc: Any, count: int = 1, before: bool = True,
                 page_index: int | None = None, width: float | None = None,
                 height: float | None = None) -> dict:
    """插入空白页（`InsertPages`；给 width+height 时走 `InsertPagesEx` 定尺寸）。

    第三个参数是**页码数字**，不是 Page 对象（传对象会 TypeError）。
    """
    idx = int(page_index if page_index is not None else 1)
    sized = width is not None and height is not None
    with mm_units(doc):
        if sized:
            created = doc.InsertPagesEx(max(1, int(count)), bool(before), idx,
                                        float(width), float(height))
        else:
            created = doc.InsertPages(max(1, int(count)), bool(before), idx)
        active = _n(lambda: int(doc.ActivePage.Index), 0)
    out = {"action": "insert", "count": max(1, int(count)), "before": bool(before),
           "page": idx, "inserted": created is not None, "active_page": active,
           "note": "插入到第 1 页前面会把新页设为活动页；后续操作请显式传 page_index",
           "pages": _n(lambda: int(doc.Pages.Count)), "sized": sized}
    if sized:
        out.update({"width": float(width), "height": float(height)})
    return out


def delete_page(doc: Any, index: int) -> dict:
    doc.DeletePages(int(index), 1)
    return {"pages": _n(lambda: int(doc.Pages.Count))}


def list_guides(doc: Any, page_index: int | None = None) -> dict:
    """列出参考线（X8 的 GetGuides 返回对象，这里尽量解析成坐标）。"""
    with mm_units(doc):
        page = _page(doc, page_index)
        guides = page.GetGuides(0)
        out: list[dict] = []
        count = _n(lambda: int(guides.Count), 0)
        for i in range(1, count + 1):
            item = _s(lambda i=i: guides.Item(i))
            if item is None:
                continue
            out.append({
                "index": i,
                "horizontal": _s(lambda it=item: bool(it.Horizontal)),
                "position": _round(_s(lambda it=item: float(it.Position))),
                "type": _s(lambda it=item: int(it.Type)),
            })
        return {"count": len(out), "guides": out}


def unlock_all(doc: Any, page_index: int | None = None) -> dict:
    with mm_units(doc):
        page = _page(doc, page_index)
        page.UnlockAllShapes()
        return {"unlocked": True}


def page_query(doc: Any, action: str, *, page_index: int | None = None,
               x: float | None = None, y: float | None = None,
               text: str | None = None, to_index: int | None = None,
               hot_area: float = 2.0) -> dict:
    """页面查询类操作（实测可用）。

    action:
      geometry      读页面尺寸与中心（GetSize / GetCenterPosition）
      shape_at      按坐标点找对象（FindShapeAtPoint）
      select_at     按坐标点选择对象（SelectShapesAtPoint）
      text_find     在页面上查找文本（TextFind）
      move_to       把当前页移动到指定位置（MoveTo）
    """
    act = (action or "geometry").lower()
    with mm_units(doc):
        page = _page(doc, page_index)
        if act == "geometry":
            size = _s(lambda: tuple(page.GetSize()))
            center = _s(lambda: tuple(page.GetCenterPosition()))
            return {
                "action": act,
                "size": [_round(v) for v in size] if size else None,
                "center": [_round(v) for v in center] if center else None,
                "orientation": ORIENTATION.get(_n(lambda: int(page.Orientation), 0), "unknown"),
            }
        if act == "shape_at":
            if x is None or y is None:
                raise CdrBadArgument("shape_at 需要 x 与 y")
            found = _s(lambda: page.FindShapeAtPoint(float(x), float(y), True))
            if found is None:
                return {"action": act, "found": False, "shape": None}
            return {"action": act, "found": True,
                    "shape": {"name": _s(lambda: str(found.Name)),
                              "type": SHAPE_TYPES.get(_n(lambda: int(found.Type), -1), "unknown")}}
        if act == "select_at":
            if x is None or y is None:
                raise CdrBadArgument("select_at 需要 x 与 y")
            page.SelectShapesAtPoint(float(x), float(y), False, float(hot_area))
            return {"action": act,
                    "selected": _n(lambda: int(doc.Selection().Shapes.Count), 0)}
        if act == "text_find":
            if not text:
                raise CdrBadArgument("text_find 需要 text")
            return {"action": act, "found": page.TextFind(str(text), True) is not None}
        if act == "move_to":
            if to_index is None:
                raise CdrBadArgument("move_to 需要 to_index")
            page.MoveTo(int(to_index))
            return {"action": act, "pages": _n(lambda: int(doc.Pages.Count))}
    raise CdrBadArgument(f"未知 action {act}；可用 geometry/shape_at/select_at/text_find/move_to")


def add_guide_angle(doc: Any, page_index: int | None = None, x: float = 0.0,
                    y: float = 0.0, angle: float = 45.0) -> dict:
    """按「基点 + 角度」加参考线（实测 IVGLayer.CreateGuideAngle 可用）。"""
    with mm_units(doc):
        page = _page(doc, page_index)
        layer = _active_layer(page)
        shape = layer.CreateGuideAngle(float(x), float(y), float(angle))
        guides = 0
        try:
            for _i in range(1, int(page.Shapes.Count) + 1):
                try:
                    if int(page.Shapes.Item(_i).Type) == 9:
                        guides += 1
                except Exception:
                    continue
        except Exception:
            pass
        return {"action": "guide_angle", "created": shape is not None,
                "x": x, "y": y, "angle": angle, "guides": guides}


def page_snap_point(doc: Any, page_index: int | None = None, x: float = 0.0,
                    y: float = 0.0, snap_types: int = 2,
                    hot_area: float = 5.0) -> dict:
    """按坐标找最近的吸附点（实测 TypeSet 位掩码：1用户 2对象 4包围盒 8边 16自由 32自动）。

    注意：实测 TypeSet=1 / 16 / 32 会报「服务器出现意外情况」，
    2 / 4 / 8 以及它们的组合（3/6/15/63）可用。
    """
    with mm_units(doc):
        page = _page(doc, page_index)
        snap = _s(lambda: page.FindClosestSnapPoint(
            int(snap_types), float(x), float(y), float(hot_area)))
        if snap is None:
            return {"action": "snap_point", "found": False,
                    "snap_types": int(snap_types)}
        return {"action": "snap_point", "found": True,
                "snap_types": int(snap_types),
                "position": [_round(_s(lambda: float(snap.PositionX))),
                             _round(_s(lambda: float(snap.PositionY)))],
                "type": _n(lambda: int(snap.Type)),
                "direction": _round(_s(lambda: float(snap.Direction))),
                "uses_direction": bool(_s(lambda: bool(snap.UsesDirection), False)),
                "selectable": bool(_s(lambda: bool(snap.IsSelectable), False)),
                "movable": bool(_s(lambda: bool(snap.IsMovable), False)),
                "deletable": bool(_s(lambda: bool(snap.IsDeletable), False))}


# --------------------------------------------------------------------------- #
# 吸附点（IVGSnapPoint / IVGSnapPoints / IVGSnapPointRange）
# --------------------------------------------------------------------------- #
SNAP_ACTIONS = ("list", "info", "of_type", "add_user", "free", "create", "find",
                "move", "delete", "direction", "select", "clear_selection", "range",
                "auto", "bbox", "obj", "edge", "user", "near")
SNAP_READ_ACTIONS = ("list", "info", "of_type", "find", "auto", "bbox", "obj",
                     "edge", "user", "near")
SNAP_RANGE_OPS = ("move", "delete", "add", "remove", "remove_by_reference", "find",
                  "selection", "add_to_selection", "remove_from_selection",
                  "auto_snap", "change_direction")
# 实测 cdrObjectSnapPointType 是位掩码：1 用户 / 2 对象 / 4 包围盒 / 8 边 / 16 自由 / 32 自动
SNAP_TYPE_NAMES = {1: "user", 2: "object", 4: "bbox", 8: "edge", 16: "free", 32: "auto"}


def _snap_type_name(value: Any) -> Any:
    iv = _n(lambda: int(value))
    if iv is None:
        return None
    return SNAP_TYPE_NAMES.get(iv, f"type_{iv}")


def _snap_point_dict(snap: Any, index: int | None = None) -> dict:
    """把一个 IVGSnapPoint 压成字典，并按类型补 BBox/Edge/User/Object 子信息。"""
    if snap is None:
        raise CdrBadArgument("吸附点对象为空")
    info: dict = {
        "index": None if index is None else int(index),
        "type": _n(lambda: int(snap.Type)),
        "kind": None,
        "position": [_round(_s(lambda: float(snap.PositionX))),
                     _round(_s(lambda: float(snap.PositionY)))],
        "direction": _round(_s(lambda: float(snap.Direction), 0.0)),
        "uses_direction": bool(_s(lambda: bool(snap.UsesDirection), False)),
        "movable": bool(_s(lambda: bool(snap.IsMovable), False)),
        "deletable": bool(_s(lambda: bool(snap.IsDeletable), False)),
        "selectable": bool(_s(lambda: bool(snap.IsSelectable), False)),
        "selected": bool(_s(lambda: bool(snap.Selected), False)),
        "can_change_direction": bool(_s(lambda: bool(snap.CanChangeDirection), False)),
    }
    info["kind"] = _snap_type_name(info["type"])
    got = _s(lambda: snap.GetPosition())
    if isinstance(got, (tuple, list)) and len(got) == 2:
        info["position_get"] = [_round(float(got[0])), _round(float(got[1]))]
    ref = _s(lambda: snap.ReferenceData)
    if isinstance(ref, str) and ref:
        info["reference"] = ref
    node = _s(lambda: snap.Node)
    if node is not None:
        info["node_index"] = _n(lambda: int(node.Index))
    if info["type"] == 1:
        user = _s(lambda: snap.User)
        if user is not None:
            info["id"] = _s(lambda: str(user.ID))
            info["auto_snap"] = bool(_s(lambda: bool(user.AutoSnap), False))
    elif info["type"] == 2:
        obj = _s(lambda: snap.Object)
        if obj is not None:
            info["object_type"] = _n(lambda: int(obj.Type))
    elif info["type"] == 4:
        box = _s(lambda: snap.BBox)
        if box is not None:
            info["bbox_type"] = _n(lambda: int(box.Type))
    elif info["type"] == 8:
        edge = _s(lambda: snap.Edge)
        if edge is not None:
            info["segment_index"] = _n(lambda: int(edge.SegmentIndex))
            info["segment_offset"] = _round(_s(lambda: float(edge.SegmentOffset), 0.0))
    return info


def _snap_collection(shape: Any) -> Any:
    sps = _s(lambda: shape.SnapPoints)
    if sps is None:
        raise CdrBadArgument("该对象取不到 Shape.SnapPoints（仅矢量对象有吸附点）")
    return sps


def _snap_list(collection: Any, limit: int = 500) -> list[dict]:
    total = _n(lambda: int(collection.Count), 0) or 0
    out = []
    for i in range(1, min(total, limit) + 1):
        item = _s(lambda i=i: collection.Item(i))
        if item is None:
            continue
        out.append(_snap_point_dict(item, i))
    return out


def _snap_range(sps: Any, indices: Iterable[int] | None) -> Any:
    wanted = [int(i) for i in (indices or [])]
    if not wanted:
        raise CdrBadArgument("indices 不能为空（要操作哪些吸附点）")
    refs = []
    total = _n(lambda: int(sps.Count), 0) or 0
    for i in wanted:
        if i < 1 or i > total:
            raise CdrBadArgument(f"吸附点 index 越界：{i}（共 {total} 个）")
        ref = _s(lambda i=i: sps.Item(i).ReferenceData)
        if not isinstance(ref, str) or not ref:
            raise CdrBadArgument(f"第 {i} 个吸附点取不到 ReferenceData")
        refs.append(ref)
    rng = _s(lambda: sps.Range(refs))
    if rng is None:
        raise CdrBadArgument("SnapPoints.Range 未返回集合")
    return rng


def _snap_kind_counts(points: Iterable[dict]) -> dict:
    counts: dict = {}
    for p in points:
        key = str(p.get("kind") or p.get("type"))
        counts[key] = counts.get(key, 0) + 1
    return counts


def snap_ops(doc: Any, action: str, *, target: dict | None = None,
             page_index: int | None = None, index: int | None = None,
             indices: Iterable[int] | None = None, x: float | None = None,
             y: float | None = None, dx: float = 0.0, dy: float = 0.0,
             angle: float = 0.0, direction_type: int = 0,
             uses_direction: bool | None = None, use_direction: bool = False,
             typeset: int | None = None, hot_area: float = 5.0,
             auto_snap: bool | None = None, range_op: str | None = None,
             reference: str | None = None, position: int | None = None,
             offset: float = 0.5, mode: str = "create") -> dict:
    """吸附点操作（实测可用：枚举 / 过滤 / 新建用户点 / 自由点 / 移动 / 删除 / 方向）。

    `cdrObjectSnapPointType` 是位掩码：1 用户 2 对象 4 包围盒 8 边 16 自由 32 自动；
    `FindClosestSnapPoint` 未命中时返回 Type=16 的“自由吸附点”（坐标=查询点），需自行剔除。
    实测怪癖：**移动对象吸附点会把它变成用户吸附点**（Type 2 -> 1，并排到用户点段）。
    """
    if action not in SNAP_ACTIONS:
        raise CdrBadArgument(f"未知 action {action!r}；可用：{', '.join(SNAP_ACTIONS)}")

    with mm_units(doc):
        page = _page(doc, page_index)

        if action in ("free", "create"):
            if x is None or y is None:
                raise CdrBadArgument(f"action={action} 需要 x / y（文档坐标 mm）")
            if action == "free":
                snap = _s(lambda: doc.CreateFreeSnapPoint(float(x), float(y)))
            else:
                app = _app_of(doc)
                snap = _s(lambda: app.CreateSnapPoint(float(x), float(y)))
            if snap is None:
                raise CdrBadArgument(f"action={action} 没有返回吸附点")
            result = _snap_point_dict(snap)
            result["action"] = action
            result["source"] = ("document_free_snap_point" if action == "free"
                                else "application_snap_point")
            return result

        if action == "find":
            if x is None or y is None:
                raise CdrBadArgument("action=find 需要 x / y（文档坐标 mm）")
            want = 2 if typeset is None else int(typeset)
            snap = _s(lambda: page.FindClosestSnapPoint(
                want, float(x), float(y), float(hot_area)))
            if snap is None:
                return {"action": "find", "found": False, "typeset": want,
                        "position": [float(x), float(y)], "hot_area": float(hot_area)}
            result = _snap_point_dict(snap)
            # 实测怪癖：热区内没有匹配点时会返回 Type=16 的“自由吸附点”，
            # 坐标恰好等于查询点。这种不能算命中。
            px, py = result["position"]
            fallback = (result["type"] == 16
                        and abs(float(px) - float(x)) < 1e-6
                        and abs(float(py) - float(y)) < 1e-6)
            result.update({"action": "find", "found": not fallback, "typeset": want,
                           "query": [float(x), float(y)], "hot_area": float(hot_area),
                           "fallback": fallback})
            return result

        if action == "clear_selection":
            if target is None:
                raise CdrBadArgument("action=clear_selection 需要 target（哪个对象）")
            sps = _snap_collection(find_shape(doc, page, target))
            sps.ClearSelection()
            return {"action": "clear_selection",
                    "selected": _n(lambda: int(sps.Selection.Count), 0) or 0}

        if target is None:
            raise CdrBadArgument(f"action={action} 需要 target（如 {{\"name\": ...}}）")
        shape = find_shape(doc, page, target)
        sps = _snap_collection(shape)

        if action == "list":
            points = _snap_list(sps)
            return {
                "action": "list",
                "shape": _s(lambda: str(shape.Name)),
                "count": _n(lambda: int(sps.Count), 0) or 0,
                "by_kind": _snap_kind_counts(points),
                "all": _n(lambda: int(sps.All.Count), 0) or 0,
                "selected": _n(lambda: int(sps.Selection.Count), 0) or 0,
                "points": points,
            }

        if action == "info":
            if index is None:
                raise CdrBadArgument("action=info 需要 index（第几个吸附点，从 1 起）")
            total = _n(lambda: int(sps.Count), 0) or 0
            if int(index) < 1 or int(index) > total:
                raise CdrBadArgument(f"index 越界：{index}（共 {total} 个吸附点）")
            result = _snap_point_dict(_s(lambda: sps.Item(int(index))), int(index))
            result.update({"action": "info", "shape": _s(lambda: str(shape.Name)),
                           "count": total})
            return result

        if action == "of_type":
            want = 2 if typeset is None else int(typeset)
            collection = _s(lambda: shape.SnapPointsOfType(want))
            if collection is None:
                raise CdrBadArgument(f"SnapPointsOfType({want}) 没有返回集合")
            points = _snap_list(collection)
            return {"action": "of_type", "typeset": want, "count": len(points),
                    "by_kind": _snap_kind_counts(points), "points": points}

        if action in ("bbox", "obj", "edge"):
            # 集合级的定址取点：BBox(Type 1..9) / Object(Type 1..4) / Edge(SegmentIndex>=1, offset)
            if action == "bbox":
                pos = 1 if index is None else int(index)
                if not 1 <= pos <= 9:
                    raise CdrBadArgument(f"bbox 的 index 只能是 1..9，收到 {pos}")
                snap = _s(lambda: sps.BBox(pos))
            elif action == "obj":
                pos = 1 if index is None else int(index)
                if not 1 <= pos <= 4:
                    raise CdrBadArgument(f"obj 的 index 只能是 1..4，收到 {pos}")
                snap = _s(lambda: sps.Object(pos))
            else:
                seg = 1 if index is None else int(index)
                if seg < 1:
                    raise CdrBadArgument(f"edge 的 index（边序号）必须 >= 1，收到 {seg}")
                snap = _s(lambda: sps.Edge(seg, float(offset)))
            if snap is None:
                raise CdrBadArgument(f"{action} 没有返回吸附点")
            result = _snap_point_dict(snap, index)
            result["action"] = action
            if action == "edge":
                result["offset"] = float(offset)
            return result

        if action == "auto":
            # 实测 Auto 是方法（返回对象的自动吸附点），不是集合
            snap = _s(lambda: sps.Auto())
            if snap is None:
                raise CdrBadArgument("该对象没有自动吸附点（SnapPoints.Auto 返回空）")
            result = _snap_point_dict(snap)
            result["action"] = "auto"
            return result

        if action == "user":
            if not reference:
                raise CdrBadArgument(
                    "action=user 需要 reference（用户吸附点 GUID，可从 list 的 id 字段拿）")
            snap = _s(lambda: sps.User(reference))
            if snap is None:
                raise CdrBadArgument(f"没有 ID 为 {reference} 的用户吸附点")
            result = _snap_point_dict(snap)
            result["action"] = "user"
            result["query_id"] = reference
            return result

        if action == "near":
            if x is None or y is None:
                raise CdrBadArgument("action=near 需要 x / y（文档坐标 mm）")
            want = 2 if typeset is None else int(typeset)
            snap = _s(lambda: sps.FindClosest(want, float(x), float(y)))
            if snap is None:
                return {"action": "near", "typeset": want, "query": [float(x), float(y)],
                        "found": False}
            result = _snap_point_dict(snap)
            result.update({"action": "near", "typeset": want,
                           "query": [float(x), float(y)], "found": True})
            return result

        if action == "add_user":
            if x is None or y is None:
                raise CdrBadArgument("action=add_user 需要 x / y（文档坐标 mm，不是对象局部坐标）")
            new = _s(lambda: sps.AddUserSnapPoint(
                float(x), float(y), float(angle), bool(use_direction)))
            if new is None:
                raise CdrBadArgument("AddUserSnapPoint 没有返回吸附点")
            result = _snap_point_dict(new)
            result.update({"action": "add_user",
                           "count": _n(lambda: int(sps.Count), 0) or 0,
                           "shape": _s(lambda: str(shape.Name))})
            return result

        if action == "direction":
            if index is None:
                raise CdrBadArgument("action=direction 需要 index")
            item = _s(lambda: sps.Item(int(index)))
            if item is None:
                raise CdrBadArgument(f"index 越界：{index}")
            item.Direction = float(angle)
            if uses_direction is not None:
                item.UsesDirection = bool(uses_direction)
            result = _snap_point_dict(item, int(index))
            result["action"] = "direction"
            return result

        if action == "select":
            if index is None:
                raise CdrBadArgument("action=select 需要 index")
            item = _s(lambda: sps.Item(int(index)))
            if item is None:
                raise CdrBadArgument(f"index 越界：{index}")
            item.CreateSelection()
            return {"action": "select", "index": int(index),
                    "selected": _n(lambda: int(sps.Selection.Count), 0) or 0}

        if action in ("move", "delete"):
            wanted = list(indices or ([int(index)] if index is not None else []))
            rng = _snap_range(sps, wanted)
            if action == "move":
                if x is not None and y is not None:
                    # 绝对定位：对范围内每个点调用 IVGSnapPoint.SetPosition
                    for _i in range(1, int(rng.Count) + 1):
                        rng.Item(_i).SetPosition(float(x), float(y))
                elif not dx and not dy:
                    raise CdrBadArgument(
                        "action=move 需要 dx / dy（相对移动）或 x / y（绝对定位）")
                else:
                    rng.Move(float(dx), float(dy))
            else:
                rng.Delete()
            sps = _snap_collection(shape)
            return {"action": action, "indices": [int(i) for i in wanted],
                    "dx": float(dx), "dy": float(dy),
                    "absolute": [float(x), float(y)] if x is not None and y is not None else None,
                    "count": _n(lambda: int(sps.Count), 0) or 0,
                    "points": _snap_list(sps)}

        # action == "range"
        op = (range_op or "").strip().lower()
        if op not in SNAP_RANGE_OPS:
            raise CdrBadArgument(
                f"action=range 需要 range_op，可选：{', '.join(SNAP_RANGE_OPS)}")
        rng = _snap_range(sps, indices)
        payload: dict = {"action": "range", "range_op": op,
                         "indices": [int(i) for i in (indices or [])],
                         "count_before": _n(lambda: int(rng.Count), 0) or 0}
        if op == "move":
            rng.Move(float(dx), float(dy))
            payload.update({"dx": float(dx), "dy": float(dy)})
        elif op == "delete":
            rng.Delete()
        elif op == "add":
            if index is None:
                raise CdrBadArgument("range_op=add 需要 index（把第几个点加进范围）")
            point = _s(lambda: sps.Item(int(index)))
            if point is None:
                raise CdrBadArgument(f"index 越界：{index}")
            payload["added"] = bool(_s(lambda: bool(rng.Add(point)), False))
        elif op == "remove":
            pos = 1 if position is None else int(position)
            payload["removed"] = bool(_s(lambda: bool(rng.Remove(pos)), False))
            payload["position"] = pos
        elif op == "remove_by_reference":
            if not reference:
                raise CdrBadArgument("range_op=remove 需要 reference 或 position")
            payload["removed"] = bool(_s(lambda: bool(rng.RemoveByReference(reference)), False))
        elif op == "find":
            if not reference:
                raise CdrBadArgument("range_op=find 需要 reference（吸附点 ReferenceData）")
            found = _s(lambda: rng.Find(reference))
            payload["found"] = bool(found is not None)
        elif op == "selection":
            rng.CreateSelection()
        elif op == "add_to_selection":
            rng.AddToSelection()
        elif op == "remove_from_selection":
            rng.RemoveFromSelection()
        elif op == "auto_snap":
            if auto_snap is None:
                raise CdrBadArgument("range_op=auto_snap 需要 auto_snap=true/false")
            rng.SetAutoSnap(bool(auto_snap))
            payload["auto_snap"] = bool(auto_snap)
        elif op == "change_direction":
            rng.ChangeDirection(float(angle), int(direction_type))
            payload.update({"angle": float(angle), "direction_type": int(direction_type)})
        total = _n(lambda: int(rng.Count), 0) or 0
        payload["count_after"] = total
        payload["points"] = _snap_list(rng)
        payload["shape_count"] = _n(lambda: int(sps.Count), 0) or 0
        payload["shape_selected"] = _n(lambda: int(sps.Selection.Count), 0) or 0
        return payload


# --------------------------------------------------------------------------- #
# 描边对象（IVGOutline 深度操作）
# --------------------------------------------------------------------------- #
OUTLINE_ACTIONS = ("info", "set", "convert_to_object", "copy", "compare")
OUTLINE_PROPS = {
    "width": "Width",
    "nib_stretch": "NibStretch",
    "nib_angle": "NibAngle",
    "miter_limit": "MiterLimit",
    "justification": "Justification",
    "dashdot_length": "DashDotLength",
    "behind_fill": "BehindFill",
    "scale_with_shape": "ScaleWithShape",
}


def _outline_info(shape: Any) -> dict:
    ol = _s(lambda: shape.Outline)
    if ol is None:
        raise CdrBadArgument("该对象取不到 Outline")
    color = _s(lambda: ol.Color)
    info: dict = {
        "shape": _s(lambda: str(shape.Name)),
        "width": _round(_s(lambda: float(ol.Width), 0.0)),
        "type": _n(lambda: int(ol.Type)),
        "color": _s(lambda: str(color.ToString())) if color is not None else None,
        "style_index": _n(lambda: int(ol.Style.Index),),
        "line_caps": _n(lambda: int(ol.LineCaps)),
        "line_join": _n(lambda: int(ol.LineJoin)),
        "behind_fill": bool(_s(lambda: bool(ol.BehindFill), False)),
        "scale_with_shape": bool(_s(lambda: bool(ol.ScaleWithShape), False)),
        "nib_stretch": _n(lambda: int(ol.NibStretch)),
        "nib_angle": _round(_s(lambda: float(ol.NibAngle), 0.0)),
        "miter_limit": _round(_s(lambda: float(ol.MiterLimit), 0.0)),
        "justification": _n(lambda: int(ol.Justification)),
        "dashdot_length": _round(_s(lambda: float(ol.DashDotLength), 0.0)),
        "pen_width": _round(_s(lambda: float(ol.PenWidth), 0.0)),
        "start_arrow": _s(lambda: str(ol.StartArrow.ToString())),
        "end_arrow": _s(lambda: str(ol.EndArrow.ToString())),
    }
    return info


def outline_ops(doc: Any, page_index: int | None, target: dict, action: str, *,
                changes: dict | None = None, target2: dict | None = None) -> dict:
    """描边对象深度操作（实测可用）。

    * 属性：Width / Type / Color / Style.Index / LineCaps / LineJoin / BehindFill /
      ScaleWithShape / NibStretch / NibAngle / MiterLimit / Justification
    * `Outline.SetProperties(...)` 只要**14 个参数全是实参**（可选位传 0 / 0.0 / None 的对象位）
      就能用；早先「TypeError 不可用」是传了 None 到 int 形参导致的。
    * `Outline.ConvertToObject()` 把描边变成独立曲线对象（源对象被替换，对象总数不变）。
    * `GetCopy` / `CopyAssign` / `CompareWith` 可用；`PenWidth` 恒为 0（派生属性，没用）、
      `DashDotLength` 写入不生效、`ToString/StringAssign` 在 pywin32 包装类里不存在。
    """
    act = (action or "").lower()
    if act not in OUTLINE_ACTIONS:
        raise CdrBadArgument(f"未知 action {action!r}；可用：{', '.join(OUTLINE_ACTIONS)}")

    with mm_units(doc):
        page = _page(doc, page_index)
        shape = find_shape(doc, page, target)

        if act == "info":
            return {"action": act, **_outline_info(shape)}

        if act == "compare":
            if target2 is None:
                raise CdrBadArgument("action=compare 需要 target2（对比对象）")
            other = find_shape(doc, page, target2)
            same = bool(_s(lambda: bool(shape.Outline.CompareWith(other.Outline)), False))
            return {"action": act, "same": same,
                    "a": _outline_info(shape), "b": _outline_info(other)}

        if act == "copy":
            if target2 is None:
                raise CdrBadArgument("action=copy 需要 target2（源对象，复制它的描边）")
            src = find_shape(doc, page, target2)
            shape.Outline.CopyAssign(src.Outline)
            return {"action": act, "from": _s(lambda: str(src.Name)),
                    **_outline_info(shape)}

        if act == "convert_to_object":
            before = _n(lambda: int(page.Shapes.Count), 0) or 0
            made = _s(lambda: shape.Outline.ConvertToObject())
            after = _n(lambda: int(page.Shapes.Count), 0) or 0
            if made is None:
                raise CdrComFailure("Outline.ConvertToObject 没有返回对象")
            name = _ensure_name(made)
            return {"action": act, "created": {"name": name,
                                               "type_id": _n(lambda: int(made.Type)),
                                               "bbox": _bbox(made)},
                    "shapes_before": before, "shapes_after": after}

        # action == "set"
        spec = dict(changes or {})
        if not spec:
            raise CdrBadArgument("action=set 需要 changes")
        applied: dict = {}
        if spec.get("use_set_properties_ex"):
            # 实测：SetPropertiesEx 多一个 Justification（第 15 参），其余同 SetProperties
            shape.Outline.SetPropertiesEx(
                float(spec.get("width", shape.Outline.Width)),
                None, shape.Outline.Color, None, None,
                int(bool(spec.get("behind_fill", False))),
                int(bool(spec.get("scale_with_shape", False))),
                int(spec.get("line_caps", 0)),
                int(spec.get("line_join", 0)),
                float(spec.get("nib_angle", 0.0)),
                int(spec.get("nib_stretch", 100)),
                float(spec.get("dashdot_length", 0.0)),
                float(spec.get("pen_width", 1.0)),
                float(spec.get("miter_limit", 5.0)),
                int(spec.get("justification", 0)),
            )
            applied["use_set_properties_ex"] = True
        elif spec.get("use_set_properties"):
            shape.Outline.SetProperties(
                float(spec.get("width", shape.Outline.Width)),
                None, shape.Outline.Color, None, None,
                int(bool(spec.get("behind_fill", False))),
                int(bool(spec.get("scale_with_shape", False))),
                int(spec.get("line_caps", 0)),
                int(spec.get("line_join", 0)),
                float(spec.get("nib_angle", 0.0)),
                int(spec.get("nib_stretch", 100)),
                float(spec.get("dashdot_length", 0.0)),
                float(spec.get("pen_width", 1.0)),
                float(spec.get("miter_limit", 5.0)),
            )
            applied["use_set_properties"] = True
        else:
            generic = {k: v for k, v in spec.items() if k in (
                "width", "color", "caps", "join", "behind_fill", "scale_with_shape",
                "dash", "dash_pattern", "start_arrow", "end_arrow", "none",
                "type", "save_outline_styles")}
            if generic:
                apply_outline(doc, shape, generic)
                applied.update({k: (v if not isinstance(v, (list, tuple)) else "…")
                                for k, v in generic.items()})
            for key in ("nib_stretch", "nib_angle", "miter_limit", "justification",
                        "dashdot_length"):
                if key in spec:
                    setattr(shape.Outline, OUTLINE_PROPS[key], spec[key])
                    applied[key] = spec[key]
        return {"action": act, "applied": applied, **_outline_info(shape)}


# --------------------------------------------------------------------------- #
# 矩形对象（IVGRect）与矩形造形
# --------------------------------------------------------------------------- #
RECT_ACTIONS = ("info", "offset", "inflate", "union", "intersect", "contains",
                "copy", "create_rectangle", "create_ellipse")
RECT_READ_ACTIONS = ("info", "offset", "inflate", "union", "intersect", "contains",
                     "copy")


def _rect_from_spec(app: Any, spec: dict | None) -> Any:
    """用 `Application.CreateRect()` + `SetRect(x,y,w,h)` 造一个 IVGRect。

    实测：IVGRect 没有 Left/Top/Right/Bottom 的 setter（只有 getter），
    尺寸写入必须走小写 `x`/`y` + `Width`/`Height`，或 `SetRect`。
    """
    spec = spec or {}
    r = app.CreateRect()
    if r is None:
        raise CdrComFailure("Application.CreateRect 返回空")
    r.SetRect(float(spec.get("x", 0.0)), float(spec.get("y", 0.0)),
              float(spec.get("width", spec.get("w", 10.0))),
              float(spec.get("height", spec.get("h", 10.0))))
    return r


def _rect_dict(r: Any) -> dict:
    if r is None:
        raise CdrComFailure("矩形对象为空")
    return {
        "x": _round(_s(lambda: float(r.x))),
        "y": _round(_s(lambda: float(r.y))),
        "width": _round(_s(lambda: float(r.Width))),
        "height": _round(_s(lambda: float(r.Height))),
        "left": _round(_s(lambda: float(r.Left))),
        "right": _round(_s(lambda: float(r.Right))),
        "top": _round(_s(lambda: float(r.Top))),
        "bottom": _round(_s(lambda: float(r.Bottom))),
        "center": [_round(_s(lambda: float(r.CenterX))),
                   _round(_s(lambda: float(r.CenterY)))],
        "empty": bool(_s(lambda: bool(r.IsEmpty), False)),
    }


def rect_ops(doc: Any, action: str, *, rect: dict | None = None,
             rect2: dict | None = None, dx: float = 0.0, dy: float = 0.0,
             left: float = 0.0, top: float = 0.0, right: float = 0.0,
             bottom: float = 0.0, point: dict | None = None,
             kind: str = "rectangle", radii: list | None = None,
             start_angle: float = 0.0, end_angle: float = 360.0,
             pie: bool = False, layer: str | None = None,
             page_index: int | None = None) -> dict:
    """IVGRect 矩形数学 + 用矩形造形（实测全通）。

    * `SetRect(x,y,w,h)` / `x`·`y`·`Width`·`Height`（可写）/ `Left`·`Right`·`Top`·`Bottom`·
      `CenterX`·`CenterY`（只读）/ `IsEmpty` / `GetCopy` / `CopyAssign` / `Offset` /
      `Inflate(l,t,r,b)` / `IsPointInside(x,y)` / `Union` / `Intersect` / `Clear`
    * `Layer.CreateRectangleRect(rect, rUL, rUR, rLR, rLL)` 与
      `Layer.CreateEllipseRect(rect, start, end, pie)` 可直接按矩形造形
    """
    act = (action or "").lower()
    if act not in RECT_ACTIONS:
        raise CdrBadArgument(f"未知 action {action!r}；可用：{', '.join(RECT_ACTIONS)}")

    if act in RECT_READ_ACTIONS:
        # 纯 Application 侧，不依赖文档
        app = _app_of(doc)
        r = _rect_from_spec(app, rect)
        if act == "info":
            result = {"action": act, "rect": _rect_dict(r)}
            result["is_point_inside"] = None
            return result
        if act == "offset":
            r.Offset(float(dx), float(dy))
            return {"action": act, "dx": float(dx), "dy": float(dy),
                    "rect": _rect_dict(r)}
        if act == "inflate":
            changed = bool(_s(lambda: bool(r.Inflate(float(left), float(top),
                                                     float(right), float(bottom))), False))
            return {"action": act, "deltas": [float(left), float(top), float(right),
                                              float(bottom)],
                    "changed": changed, "rect": _rect_dict(r)}
        if act in ("union", "intersect"):
            other = _rect_from_spec(app, rect2)
            got = r.Union(other) if act == "union" else r.Intersect(other)
            return {"action": act, "a": _rect_dict(r), "b": _rect_dict(other),
                    "result": _rect_dict(got)}
        if act == "contains":
            pt = point or {}
            inside = bool(_s(lambda: bool(r.IsPointInside(float(pt.get("x", 0)),
                                                          float(pt.get("y", 0)))), False))
            return {"action": act, "point": [float(pt.get("x", 0)),
                                             float(pt.get("y", 0))],
                    "inside": inside, "rect": _rect_dict(r)}
        # copy
        clone = _s(lambda: r.GetCopy())
        return {"action": act, "source": _rect_dict(r), "copy": _rect_dict(clone)}

    # create_rectangle / create_ellipse —— 需要文档与图层
    with mm_units(doc):
        page = _page(doc, page_index)
        if layer:
            target_layer = _find_layer(page, layer, None)
        else:
            target_layer = _active_layer(page)
        r = _rect_from_spec(_app_of(doc), rect)
        if act == "create_rectangle":
            rs = list(radii or [0, 0, 0, 0]) + [0, 0, 0, 0]
            shape = target_layer.CreateRectangleRect(r, float(rs[0]), float(rs[1]),
                                                     float(rs[2]), float(rs[3]))
        else:
            shape = target_layer.CreateEllipseRect(r, float(start_angle),
                                                   float(end_angle), bool(pie))
        if shape is None:
            raise CdrComFailure(f"{act} 没有返回对象")
        name = _ensure_name(shape)
        return {"action": act, "rect": _rect_dict(r),
                "kind": ("rectangle" if act == "create_rectangle" else "ellipse"),
                "shape": {"name": name, "type_id": _n(lambda: int(shape.Type)),
                          "bbox": _bbox(shape)}}


def objects_bounds(doc: Any, page_index: int | None = None,
                   selected_only: bool = False) -> dict:
    """本页所有对象（或当前选中对象）的总包围盒。

    实测：`Page.GetObjectsBoundingBox()` 返回的是 `IVGRect` 对象，而 `IVGRect.GetRect()`
    的四个出参 pywin32 取不回来，所以这里直接遍历对象自己求并集，结果可预期。
    """
    with mm_units(doc):
        page = _page(doc, page_index)
        boxes: list[dict] = []
        if selected_only:
            sel = doc.Selection().Shapes
            for i in range(1, int(sel.Count) + 1):
                box = _bbox(sel.Item(i))
                if box:
                    boxes.append(box)
        else:
            for _idx, _li, _j, shape in _iter_page_shapes(page):
                box = _bbox(shape)
                if box:
                    boxes.append(box)
        if not boxes:
            return {"bbox": None, "count": 0}
        left = min(b["x"] for b in boxes)
        bottom = min(b["y"] for b in boxes)
        right = max(b["x"] + b["width"] for b in boxes)
        top = max(b["y"] + b["height"] for b in boxes)
        return {
            "count": len(boxes),
            "bbox": {
                "x": _round(left), "y": _round(bottom),
                "width": _round(right - left), "height": _round(top - bottom),
            },
        }


# --------------------------------------------------------------------------- #
# 位图
# --------------------------------------------------------------------------- #
BITMAP_MODES = {"rgb": 1, "cmyk": 2, "gray": 3, "grayscale": 3, "bw": 4, "blackwhite": 4}
BITMAP_ACTIONS = ("info", "convert", "resample", "crop", "reset_crop", "bw", "inflate",
                  "pixel", "image_info", "duotone", "paletted", "paletted_options",
                  "replace_image", "effect", "link_info", "update_link", "resolve_link")


def _bitmap_of(shape: Any) -> Any:
    try:
        bmp = shape.Bitmap
        if bmp is not None:
            return bmp
    except Exception:
        pass
    raise CdrBadArgument("目标不是位图对象（type_id 应为 5）")


def bitmap_action(doc: Any, page_index: int | None, target: dict, action: str, *,
                  mode: str | None = None, width: float | None = None,
                  height: float | None = None, dpi: float | None = None,
                  anti_alias: bool = True, threshold: float = 128.0,
                  inflate_x: float = 2.0, inflate_y: float = 2.0,
                  x_px: int | None = None, y_px: int | None = None,
                  duotone_type: int | None = None, ink_colors: list | None = None,
                  preset: str | None = None, palette_type: int | None = 3,
                  num_colors: int = 256, palette_index: int | None = None,
                  dither_type: int | None = None,
                  image_color: Any = "#FFFFFF", image_pixels_w: int = 16,
                  image_pixels_h: int = 16, image_fills: list | None = None,
                  image_pixels: list | None = None, offset_x: int = 0,
                  offset_y: int = 0, effect_name: str | None = None,
                  effect_command: str | None = None,
                  label: str = "cdr_bitmap") -> dict:
    """位图处理。

    实测可用：ConvertTo（RGB/CMYK/灰度/黑白）、ConvertToBW（阈值二值化）、
    Resample（尺寸+分辨率）、Inflate（四周扩边）、Crop、ResetCropEnvelope、
    **GetPixel 取像素色**、**ConvertToDuotone（双色调/三色调/四色调）**、
    **ConvertToPaletted（调色板模式 0~8）**。
    **PowerTRACE（Bitmap.Trace）不提供**：实测会让 CorelDRAW 卡成「无响应」（模态对话框），
    COM 无法中断，会把调用方一起拖死。
    """
    act = (action or "info").lower()
    if act not in BITMAP_ACTIONS:
        raise CdrBadArgument(f"未知 action {action!r}；可用：{', '.join(BITMAP_ACTIONS)}")

    with mm_units(doc):
        page = _page(doc, page_index)
        shape = find_shape(doc, page, target)
        bmp = _bitmap_of(shape)

        def info() -> dict:
            return {
                "bbox": _bbox(shape),
                "pixels": [_round(_s(lambda: float(bmp.SizeWidth))),
                           _round(_s(lambda: float(bmp.SizeHeight)))],
                "dpi": _round(_s(lambda: float(bmp.ResolutionX))),
                "mode": _n(lambda: int(bmp.Mode)),
            }

        if act == "info":
            return {"action": act, **info()}
        if act == "image_info":
            img = _s(lambda: bmp.Image)
            if img is None:
                raise CdrBadArgument("该位图没有图像缓冲区（可能是链接/嵌入异常）")
            return {"action": act,
                    "image_size": [_n(lambda: int(img.Width)), _n(lambda: int(img.Height))],
                    "image_type": _n(lambda: int(img.Type)),
                    "read_only": bool(_s(lambda: bool(img.ReadOnly), False)),
                    **info()}
        if act == "pixel":
            img = _s(lambda: bmp.Image)
            if img is None:
                raise CdrBadArgument("该位图没有图像缓冲区")
            if x_px is None or y_px is None:
                raise CdrBadArgument("pixel 需要 x_px 与 y_px（像素坐标，从 0 开始）")
            px, py = int(x_px), int(y_px)
            width_px = _n(lambda: int(img.Width), 0)
            height_px = _n(lambda: int(img.Height), 0)
            if not (0 <= px < width_px and 0 <= py < height_px):
                raise CdrBadArgument(
                    f"像素越界：({px},{py})；图像为 {width_px}x{height_px}")
            color = _s(lambda: img.Pixel(px, py))
            if color is None:
                raise CdrBadArgument("读取像素失败")
            return {"action": act, "pixel": [px, py],
                    "color": _s(lambda: str(color.ToString())),
                    "name": _s(lambda: str(color.GetName())) or None,
                    "image_size": [width_px, height_px], **info()}

        if act == "link_info":
            return {"action": act,
                    "externally_linked": bool(_s(lambda: bool(bmp.ExternallyLinked), False)),
                    "link_file_name": _s(lambda: str(bmp.LinkFileName)) or None,
                    "embedded": bool(_s(lambda: bool(bmp.Embedded), False)),
                    **info()}

        with atomic(doc, label):
            if act == "update_link":
                bmp.UpdateLink()
                return {"action": act, "updated": True, "link_file_name":
                        _s(lambda: str(bmp.LinkFileName)) or None}
            if act == "resolve_link":
                bmp.ResolveLink()
                return {"action": act, "resolved": True, "link_file_name":
                        _s(lambda: str(bmp.LinkFileName)) or None}
            if act == "convert":
                key = str(mode or "").lower().replace("-", "").replace("_", "")
                if key not in BITMAP_MODES:
                    raise CdrBadArgument(
                        f"未知颜色模式 {mode!r}；可用 rgb / cmyk / gray / bw")
                bmp.ConvertTo(BITMAP_MODES[key])
            elif act == "resample":
                bmp.Resample(float(width or 0), float(height or 0), bool(anti_alias),
                             float(dpi or 0), float(dpi or 0))
            elif act == "crop":
                bmp.Crop()
            elif act == "reset_crop":
                bmp.ResetCropEnvelope()
            elif act == "bw":
                # 实测：ConvertToBW(RenderType, Intensity, Threshold, Halftone, Angle, Size)
                bmp.ConvertToBW(0, 0, float(threshold), False, 0, 0)
            elif act == "inflate":
                bmp.Inflate(float(inflate_x), float(inflate_y))
            elif act == "duotone":
                app = _app_of(doc)
                duo = app.CreateDuotone()
                if preset:
                    duo.Load(str(preset))
                else:
                    dtype = int(duotone_type if duotone_type is not None else 1)
                    if not 0 <= dtype <= 3:
                        raise CdrBadArgument("duotone_type 可用 0~3（0=单色调 3=四色调）")
                    duo.Type = dtype
                inks = [c for c in (ink_colors or []) if c is not None]
                ink_count = _n(lambda: int(duo.InkCount), 1)
                for idx, spec in enumerate(inks[:ink_count], 1):
                    duo.Inks(idx).Color = _color(doc, spec)
                # 实测：IVGDuotone 还有 GetCopy / CopyAssign / ResetOverprints / Load / Save；
                # 转完之后 bmp.Duotone 就能拿到当前的 IVGDuotone，再取 Overprints(i)
                # 得到 IVGDuotoneOverprint（SetValues / Reset）
                duo_copy = _s(lambda: duo.GetCopy())
                if duo_copy is not None:
                    duo.CopyAssign(duo_copy)
                bmp.ConvertToDuotone(duo)
                live_duo = _s(lambda: bmp.Duotone)
                res: dict = {"action": act,
                             "mode": _n(lambda: int(bmp.Mode)),
                             "ink_count": ink_count}
                if live_duo is not None:
                    res["overprint_count"] = _n(
                        lambda: int(live_duo.OverprintCount), 0)
                    res["overprints"] = _n(lambda: int(live_duo.OverprintCount), 0)
                    over = _s(lambda: live_duo.Overprints(1))
                    if over is not None:
                        res["overprint_cyan"] = _n(lambda: int(over.Cyan))
                        try:
                            over.SetValues(10, 20, 30, 40)
                            res["overprint_set"] = True
                            over.Reset()
                            res["overprint_reset"] = True
                        except Exception:
                            res["overprint_set"] = False
                    live_duo.ResetOverprints()
                    duo_path = preset or ""
                    if isinstance(preset, str) and preset:
                        live_duo.Save(str(preset))
                        res["saved_to"] = str(preset)
                return res
            elif act == "paletted_options":
                # 实测：Application.CreateStructPaletteOptions() + Bitmap.ConvertToPaletted2(opts)
                # 可指定 NumColors / DitherType / PaletteType，并可挂已打开的调色板
                app = _app_of(doc)
                opts = _s(lambda: app.CreateStructPaletteOptions())
                if opts is None:
                    raise CdrComFailure("CreateStructPaletteOptions 返回空")
                if num_colors:
                    try:
                        opts.NumColors = int(num_colors)
                    except Exception:
                        pass
                if dither_type is not None:
                    try:
                        opts.DitherType = int(dither_type)
                    except Exception:
                        pass
                if palette_type is not None:
                    try:
                        opts.PaletteType = int(palette_type)
                    except Exception:
                        pass
                if palette_index is not None:
                    pal = _find_palette(doc, int(palette_index))
                    try:
                        opts.Palette = pal
                    except Exception:
                        pass
                bmp.ConvertToPaletted2(opts)
                return {"action": act, "mode": _n(lambda: int(bmp.Mode)),
                        "num_colors": _n(lambda: int(opts.NumColors)),
                        "palette_type": _n(lambda: int(opts.PaletteType)),
                        "dither_type": _n(lambda: int(opts.DitherType))}
            elif act == "paletted":
                pt = int(palette_type)
                palette = None
                if pt == 9:
                    if palette_index is None:
                        raise CdrBadArgument(
                            "palette_type=9（自定义）需要 palette_index"
                            "（先用 cdr_doc(action='palettes') 或 palette_open_fixed 打开）")
                    palette = _find_palette(doc, int(palette_index))
                elif not 0 <= pt <= 8:
                    raise CdrBadArgument(
                        "palette_type 可用 0~9（0=uniform 1=stdVGA 2=adaptive 3=optimized "
                        "4=blackbody 5=grayscale 6=system 7=IE 8=Netscape 9=自定义+palette_index）")
                if palette is not None:
                    bmp.ConvertToPaletted(pt, 0, 100, 0, int(num_colors), False,
                                          16777215, 500, 300, 300, 100, palette)
                else:
                    bmp.ConvertToPaletted(pt, 0, 100, 0, int(num_colors))
            elif act == "replace_image":
                px_w = max(1, min(int(image_pixels_w), 8192))
                px_h = max(1, min(int(image_pixels_h), 8192))
                image = doc.CreateImage(4, px_w, px_h, _color(doc, image_color))
                for area in (image_fills or []):
                    if not isinstance(area, dict):
                        continue
                    image.FillArea(int(area.get("x", 0)), int(area.get("y", 0)),
                                   int(area.get("width", 1)), int(area.get("height", 1)),
                                   _color(doc, area.get("color", "#000000")))
                for px in (image_pixels or []):
                    if not isinstance(px, (list, tuple)) or len(px) < 3:
                        continue
                    image.FillArea(int(px[0]), int(px[1]), int(px[0]), int(px[1]),
                                   _color(doc, px[2]))
                # ImageAlpha 必须传 None（传 0 会 TypeError）
                bmp.SetImageData(image, None, int(offset_x), int(offset_y))
            elif act == "effect":
                if not effect_name or not effect_command:
                    raise CdrBadArgument("effect 需要 effect_name 与 effect_command")
                bmp.ApplyBitmapEffect(str(effect_name), str(effect_command))
    return {"action": act, **info()}


# --------------------------------------------------------------------------- #
# 路径（曲线）
# --------------------------------------------------------------------------- #
PATH_ACTIONS = (
    "info", "contour", "add_subpath", "polyline",
    # 拼接（曲线级）
    "append_circle", "append_ellipse", "append_rectangle", "append_arc", "append_bezier",
    # 整形（曲线级）
    "reduce_nodes", "self_weld",
    # 判定（曲线级）
    "point_inside", "intersects",
    # 子路径级
    "subpath_info", "equal_divide", "add_node", "break_apart",
    "get_point", "get_tangent", "move_subpath",
    # 查询（第三批）
    "curvature", "perpendicular", "evenly_spaced", "closest_segment",
    "is_on_path", "segment_at", "is_rect_on_edge",
    # 第四批（第五波实测）
    "point_position", "tangent_vector", "perpendicular_vector", "curve_speed",
    "intersections", "find_node", "find_segment",
    "reverse_direction", "join_subpaths", "remove_overlaps",
    # 第三十六波：曲线合并 / 焊接 / 命中判定
    "merge_curve", "weld_curve", "is_on_curve", "curve_copy", "curve_info",
    # 第四十一波：任意线段集合（Application.CreateSegmentRange + Add）
    "segment_range",
    # 第五十波：子路径追加曲线段 / 抽出子路径 / 焊接结果落地
    "append_curve_segment", "extract_subpath", "weld_to_shape",
    # 第五十二波：按偏移取点 / 按偏移定位线段
    "point_at_offset", "segment_at_offset",
    # 第五十六波：把曲线映射到另一条"骨架线"上（艺术笔映射）
    "map_to_stroke",
    # 第五十九波：点集 → 拟合曲线 / 追加点集
    "fit_to_points", "append_points",
)


def _polyline_points(curve: Any, precision: int = 50, limit: int = 4000) -> list[list[float]]:
    """把曲线采样成折线点（实测 X8 可用）。

    关键发现：`Curve.GetPolyline()` 返回的是 **IVGCurve**（不是 pywin32 取不回的
    变体数组）。遍历返回曲线的 SubPaths → Nodes 即可拿到坐标；采样点里会有连续
    重复点（退化子路径），这里按"连续相同"去重。
    """
    poly = curve.GetPolyline(int(precision))
    points: list[list[float]] = []
    last: tuple[float, float] | None = None
    for si in range(1, int(poly.SubPaths.Count) + 1):
        sp = poly.SubPaths.Item(si)
        for ni in range(1, int(sp.Nodes.Count) + 1):
            node = sp.Nodes.Item(ni)
            pt = (_round(_s(lambda n=node: float(n.PositionX), 0.0)),
                  _round(_s(lambda n=node: float(n.PositionY), 0.0)))
            if pt != last:
                points.append([pt[0], pt[1]])
                last = pt
            if len(points) >= limit:
                return points
    return points


def path_ops(doc: Any, page_index: int | None, target: dict, action: str, *,
             offset: float = 1.0, direction: int = 0, x: float | None = None,
             y: float | None = None, precision: int = 50,
             radius: float | None = None, rx: float | None = None, ry: float | None = None,
             rect: list | None = None, arc: list | None = None, bezier: list | None = None,
             other: dict | None = None, amount: float = 30.0, closed: bool = True,
             subpath: int = 1, t: float = 0.5, dx: float = 0.0, dy: float = 0.0,
             tolerance: float = 0.0, allow_reversals: bool = True,
             hot_area: float = -1.0, params: dict | None = None,
             label: str = "cdr_path") -> dict:
    """曲线 / 子路径级操作。所有坐标毫米。"""
    act = (action or "info").lower()
    if act not in PATH_ACTIONS:
        raise CdrBadArgument(f"未知 action {action!r}；可用：{', '.join(PATH_ACTIONS)}")

    with mm_units(doc):
        page = _page(doc, page_index)
        shape = find_shape(doc, page, target)
        curve = _s(lambda: shape.Curve)
        if curve is None:
            raise CdrBadArgument("目标不是曲线对象（可先用 cdr_shape_ops(op='to_curves') 转曲）")

        def info() -> dict:
            return {
                "length": _round(_s(lambda: float(curve.Length))),
                "area": _round(_s(lambda: float(curve.Area))),
                "subpaths": _n(lambda: int(curve.SubPaths.Count)),
                "bbox": _bbox(shape),
            }

        def sub() -> Any:
            count = _n(lambda: int(curve.SubPaths.Count), 0)
            if not 1 <= int(subpath) <= count:
                raise CdrBadArgument(f"subpath 越界：{subpath}（共 {count} 条）")
            return curve.SubPaths.Item(int(subpath))

        # ---- 只读 ----
        if act == "info":
            return {"action": act, **info()}
        if act == "polyline":
            try:
                pts = _polyline_points(curve, int(precision))
                return {"action": act, "points": pts[:200], "count": len(pts)}
            except Exception as exc:  # noqa: BLE001
                return {"action": act, "points": [],
                        "note": f"该曲线无法采样（{type(exc).__name__}: {exc}）"}
        if act == "point_inside":
            if x is None or y is None:
                raise CdrBadArgument("point_inside 需要 x 与 y")
            return {"action": act, "inside": bool(curve.IsPointInside(float(x), float(y)))}
        if act == "intersects":
            other_shape = find_shape(doc, page, other or {})
            oc = _s(lambda: other_shape.Curve)
            if oc is None:
                raise CdrBadArgument("other 必须是曲线对象")
            return {"action": act, "intersects": bool(curve.IntersectsWith(oc))}
        if act == "subpath_info":
            sp = sub()
            pos = _s(lambda: tuple(sp.GetPosition()))
            return {
                "action": act,
                "position": [_round(v) for v in pos] if pos else None,
                "length": _round(_s(lambda: float(sp.Length))),
            }

        # ---- 查询类（第三批，实测均可）----
        if act == "curvature":
            sp = sub()
            return {"action": act, "t": t,
                    "curvature": _round(_s(lambda: float(sp.GetCurvatureAt(float(t), 0))))}
        if act == "perpendicular":
            sp = sub()
            return {"action": act, "t": t,
                    "angle": _round(_s(lambda: float(sp.GetPerpendicularAt(float(t), 0))))}
        if act == "evenly_spaced":
            sp = sub()
            try:
                pts = tuple(sp.GetEvenlySpacedPoints(float(offset or 10), True))
                return {"action": act, "count": len(pts),
                        "points": [_round(v) for v in pts[:100]]}
            except Exception:
                return {"action": act, "points": [],
                        "note": "该子路径取不到等距点数组（pywin32 限制）"}
        if act == "closest_segment":
            sp = sub()
            if x is None or y is None:
                raise CdrBadArgument("closest_segment 需要 x 与 y")
            res = _s(lambda: tuple(sp.FindClosestSegment(float(x), float(y), 0)))
            return {"action": act, "result": [_round(v) for v in res] if res else None}
        if act == "segment_at":
            sp = sub()
            res = _s(lambda: tuple(sp.GetSegmentAt(float(t), 0, 0)))
            return {"action": act, "t": t,
                    "segment": [_round(v) for v in res] if res else None}
        if act == "is_on_path":
            sp = sub()
            if x is None or y is None:
                raise CdrBadArgument("is_on_path 需要 x 与 y")
            return {"action": act,
                    "on_path": bool(_s(lambda: sp.IsOnSubPath(float(x), float(y), 2.0), 0))}
        if act == "is_rect_on_edge":
            r = rect or []
            if len(r) != 4:
                raise CdrBadArgument("is_rect_on_edge 需要 rect=[x1,y1,x2,y2]")
            return {"action": act,
                    "on_edge": bool(_s(lambda: curve.IsRectOnEdge(
                        float(r[0]), float(r[1]), float(r[2]), float(r[3])), 0))}
        if act in {"get_point", "get_tangent"}:
            sp = sub()
            if act == "get_tangent":
                return {"action": act, "t": t,
                        "tangent": _round(_s(lambda: float(sp.GetTangentAt(float(t), 0))))}
            pt = _s(lambda: sp.GetPointAt(float(t), 0))
            # 实测：GetPointAt 返回 IVGPoint，坐标是 .x / .y（小写）
            if pt is None:
                return {"action": act, "t": t, "point": None}
            return {"action": act, "t": t,
                    "point": [_round(_s(lambda: float(pt.x))),
                              _round(_s(lambda: float(pt.y)))]}

        # ---- 查询（第四批，2026-10-25 实测）----
        if act == "point_position":
            sp = sub()
            pt = _s(lambda: sp.GetPointAt(float(t), 1))
            if pt is None:
                return {"action": act, "t": t, "point": None}
            return {"action": act, "t": t,
                    "point": [_round(_s(lambda: float(pt.x))),
                              _round(_s(lambda: float(pt.y)))]}
        if act in {"tangent_vector", "perpendicular_vector"}:
            sp = sub()
            # 实测：Normalize=True 时 X8 返回零向量，必须取 False 再自己归一化
            if act == "tangent_vector":
                vec = _s(lambda: sp.GetTangentVectorAt(float(t), 1, False))
            else:
                vec = _s(lambda: sp.GetPerpendicularVectorAt(float(t), 1, False))
            if vec is None:
                return {"action": act, "t": t, "vector": None}
            vx = _s(lambda: float(vec.x))
            vy = _s(lambda: float(vec.y))
            length = _round(_s(lambda: float(vec.Length)))
            normalized = None
            if vx is not None and vy is not None and length:
                normalized = [_round(vx / length, 6), _round(vy / length, 6)]
            return {"action": act, "t": t,
                    "vector": [_round(vx), _round(vy)],
                    "normalized": normalized,
                    "angle": _round(_s(lambda: float(vec.Angle))),
                    "length": length}
        if act == "curve_speed":
            sp = sub()
            return {"action": act, "t": t,
                    "speed": _round(_s(lambda: float(sp.GetCurveSpeedAt(float(t), 1))))}
        if act == "intersections":
            sp = sub()
            other_shape = find_shape(doc, page, other or {})
            oc = _s(lambda: other_shape.Curve)
            if oc is None:
                raise CdrBadArgument("other 必须是曲线对象")
            count = _n(lambda: int(oc.SubPaths.Count), 0)
            if count < 1:
                raise CdrBadArgument("other 没有子路径")
            idx = max(1, min(int(subpath), count))
            target_sp = oc.SubPaths.Item(idx)
            cross = _s(lambda: sp.GetIntersections(target_sp, 1))
            points: list[list[float]] = []
            if cross is not None:
                for i in range(1, int(cross.Count) + 1):
                    pt = cross.Item(i)
                    points.append([_round(_s(lambda p=pt: float(p.x))),
                                   _round(_s(lambda p=pt: float(p.y)))])
            return {"action": act, "count": len(points), "points": points}
        if act == "find_node":
            if x is None or y is None:
                raise CdrBadArgument("find_node 需要 x 与 y")
            node = _s(lambda: curve.FindNodeAtPoint(float(x), float(y), float(hot_area)))
            if node is None:
                return {"action": act, "found": False, "node": None}
            return {"action": act, "found": True, "node": _node_dict(node)}
        if act == "find_segment":
            if x is None or y is None:
                raise CdrBadArgument("find_segment 需要 x 与 y")
            res = _s(lambda: curve.FindSegmentAtPoint(
                float(x), float(y), HotArea=float(hot_area)))
            if res is None:
                return {"action": act, "found": False, "segment": None}
            seg = res[0] if isinstance(res, (tuple, list)) else res
            offset_val = res[1] if isinstance(res, (tuple, list)) and len(res) > 1 else None
            return {"action": act, "found": True,
                    "segment": _segment_dict(seg),
                    "offset": _round(offset_val) if offset_val is not None else None}

        # ---- 变更 ----
        with atomic(doc, label):
            if act == "contour":
                new_curve = curve.Contour(float(offset))
                new_shape = _active_layer(page).CreateCurve(new_curve)
                return {"action": act, "offset": offset,
                        "created": _s(lambda: str(new_shape.Name)) if new_shape is not None else None,
                        "bbox": _bbox(new_shape) if new_shape is not None else None}

            if act == "add_subpath":
                if x is None or y is None:
                    raise CdrBadArgument("add_subpath 需要 x 与 y")
                curve.CreateSubPath(float(x), float(y))
                return {"action": act, **info()}

            if act == "append_circle":
                if x is None or y is None:
                    raise CdrBadArgument("append_circle 需要 x,y,radius")
                curve.AppendSubpathCircle(float(x), float(y), float(radius or 5))
                return {"action": act, **info()}
            if act == "append_ellipse":
                if x is None or y is None:
                    raise CdrBadArgument("append_ellipse 需要 x,y,rx,ry")
                curve.AppendSubpathEllipse(float(x), float(y),
                                           float(rx or radius or 5), float(ry or radius or 5))
                return {"action": act, **info()}
            if act == "append_rectangle":
                r = rect or [x or 0, y or 0, (x or 0) + 20, (y or 0) + 20]
                if len(r) != 4:
                    raise CdrBadArgument("append_rectangle 需要 rect=[left,bottom,right,top]")
                curve.AppendSubpathRectangle(float(r[0]), float(r[1]), float(r[2]), float(r[3]))
                return {"action": act, **info()}
            if act == "append_arc":
                a = arc or []
                if len(a) != 6:
                    raise CdrBadArgument("append_arc 需要 arc=[startX,startY,endX,endY,thirdX,thirdY]")
                curve.AppendSubpathThreePointArc(*[float(v) for v in a])
                return {"action": act, **info()}
            if act == "append_bezier":
                sp = sub()
                b = bezier or []
                if len(b) < 2:
                    raise CdrBadArgument(
                        "append_bezier 需要 bezier=[x,y,c1x,c1y,c2x,c2y]（追加一段贝塞尔）")
                pts = [float(v) for v in b]
                if len(pts) == 6:
                    sp.AppendCurveSegment2(pts[0], pts[1], pts[2], pts[3],
                                           pts[4], pts[5], False)
                elif len(pts) == 2:
                    sp.AppendLineSegment(pts[0], pts[1])
                else:
                    raise CdrBadArgument("bezier 需为 6 个数或 2 个数")
                return {"action": act, **info()}

            if act == "reduce_nodes":
                curve.AutoReduceNodes(float(amount), False)
                return {"action": act, **info()}
            if act == "segment_range":
                # 实测：CreateSegmentRange() + Add(segment) 可凑任意线段集合，
                # 再 SetType / AddNode / CreateSelection / RemoveAll
                spec = dict(params or {})
                segs = [int(i) for i in (spec.get("segments") or [])]
                if not segs:
                    raise CdrBadArgument("segment_range 需要 params.segments=[线段序号…]")
                total = _n(lambda: int(curve.Segments.Count), 0) or 0
                srange = _s(lambda: _app_of(doc).CreateSegmentRange())
                if srange is None:
                    raise CdrComFailure("Application.CreateSegmentRange 返回空")
                for idx in segs:
                    if not 1 <= idx <= total:
                        raise CdrBadArgument(f"线段 {idx} 越界（共 {total} 条）")
                    srange.Add(curve.Segments.Item(idx))
                count = _n(lambda: int(srange.Count), 0) or 0
                op = str(spec.get("op", "select")).lower()
                nodes_before = _node_count(curve)
                with atomic(doc, label):
                    if op == "set_type":
                        srange.SetType(int(spec.get("segment_type", 0)))
                    elif op == "add_node":
                        srange.AddNode()
                    elif op == "select":
                        srange.CreateSelection()
                    elif op == "remove_all":
                        srange.RemoveAll()
                    else:
                        raise CdrBadArgument(
                            "segment_range 的 op 可用 set_type / add_node / select / remove_all")
                return {"action": act, "op": op, "count": count,
                        "segments_total": total,
                        "nodes_before": nodes_before, "nodes_after": _node_count(curve),
                        "first_segment": _n(lambda: int(srange.FirstSegment.Index)),
                        "last_segment": _n(lambda: int(srange.LastSegment.Index)),
                        **info()}
            if act in {"fit_to_points", "append_points"}:
                # 实测：IVGPointRange 可从 SubPath.GetEvenlySpacedPoints(间距, 缩放) 拿到，
                # 也可用 AddPointXY 逐个补点；之后：
                #   Document.CreateCurveFitToPoints(range, False, 容差) → 拟合曲线
                #   Curve.AppendSubpathFromPoints(range, 闭合) / AppendSubpathFitToPoints(…)
                spec = dict(params or {})
                app = _app_of(doc)
                points = spec.get("points")
                rng = None
                source_name = None
                if other:
                    stroke = find_shape(doc, page, other)
                    sub = _s(lambda: stroke.Curve.SubPaths.Item(
                        int(spec.get("subpath", 1))))
                    if sub is None:
                        raise CdrBadArgument("other 不是曲线对象")
                    rng = sub.GetEvenlySpacedPoints(
                        float(spec.get("spacing", 5.0)),
                        bool(spec.get("scale_to_fit", True)))
                    source_name = _s(lambda: str(stroke.Name))
                if rng is None:
                    if not points or len(points) < 2:
                        raise CdrBadArgument(
                            "fit_to_points 需要 other（取点来源）或 points=[[x,y],…]")
                    rng = _make_point_range(app, points)
                if spec.get("extra_points"):
                    for pt in spec["extra_points"]:
                        rng.AddPointXY(float(pt[0]), float(pt[1]))
                if act == "fit_to_points":
                    with atomic(doc, label):
                        if spec.get("with_cusps"):
                            # 带尖点数组的版本：尖点索引传空 tuple 即可
                            made_curve = doc.CreateCurveFitToPointsAndCusps(
                                rng, (), False, float(spec.get("tolerance", 3.0)))
                        else:
                            made_curve = doc.CreateCurveFitToPoints(
                                rng, False, float(spec.get("tolerance", 3.0)))
                        if made_curve is None:
                            raise CdrComFailure("CreateCurveFitToPoints 返回空")
                        made = _active_layer(page).CreateCurve(made_curve)
                    if made is None:
                        raise CdrComFailure("Layer.CreateCurve(拟合结果) 返回空")
                    name = _ensure_name(made)
                    return {"action": act, "shape": name,
                            "points": _n(lambda: int(rng.Count), 0),
                            "type_id": _n(lambda: int(made.Type)),
                            "nodes": _n(lambda: int(made.Curve.Nodes.Count), 0),
                            "bbox": _bbox(made), "from": source_name}
                close_sub = bool(spec.get("close", False))
                mode = str(spec.get("mode", "plain")).lower()
                before = _node_count(curve)
                with atomic(doc, label):
                    if mode == "fit":
                        curve.AppendSubpathFitToPoints(
                            rng, False, float(spec.get("tolerance", 3.0)))
                    elif mode == "cusps":
                        curve.AppendSubpathFitToPointsAndCusps(
                            rng, (), False, float(spec.get("tolerance", 3.0)))
                    else:
                        curve.AppendSubpathFromPoints(rng, close_sub)
                return {"action": act, "mode": mode, "points": _n(
                    lambda: int(rng.Count), 0),
                    "nodes_before": before, "nodes_after": _node_count(curve),
                    **info()}
            if act == "map_to_stroke":
                # 实测：Curve.CreateCurveMappedToStroke(IVGSubPath, ScaleY, SelfWeld)
                # 与 ...AndReferenceLine(IVGSubPath, IVGPoint, IVGPoint, ScaleY, SelfWeld)
                # —— 第二个参数要的是**子路径**（不是整条 Curve），参考线版的两个点是 IVGPoint
                if not other:
                    raise CdrBadArgument("map_to_stroke 需要 other（骨架线对象）")
                spec = dict(params or {})
                stroke_shape = find_shape(doc, page, other)
                stroke_sub = _s(lambda: stroke_shape.Curve.SubPaths.Item(1))
                if stroke_sub is None:
                    raise CdrBadArgument("other 不是曲线对象（需要 IVGCurve）")
                scale_y = float(spec.get("scale_y", 1.0))
                self_weld = bool(spec.get("self_weld", False))
                with atomic(doc, label):
                    if spec.get("use_reference_line"):
                        nodes = stroke_shape.Curve.Nodes
                        p1 = nodes.Item(1).GetPoint()
                        p2 = nodes.Item(int(_n(lambda: int(nodes.Count), 2) or 2)).GetPoint()
                        made_curve = curve.CreateCurveMappedToStrokeAndReferenceLine(
                            stroke_sub, p1, p2, scale_y, self_weld)
                        used = "CreateCurveMappedToStrokeAndReferenceLine"
                    else:
                        made_curve = curve.CreateCurveMappedToStroke(
                            stroke_sub, scale_y, self_weld)
                        used = "CreateCurveMappedToStroke"
                    if made_curve is None:
                        raise CdrComFailure(f"{used} 返回空")
                    made = _active_layer(page).CreateCurve(made_curve)
                if made is None:
                    raise CdrComFailure("Layer.CreateCurve(映射结果) 返回空")
                name = _ensure_name(made)
                return {"action": act, "api": used, "shape": name,
                        "type_id": _n(lambda: int(made.Type)),
                        "nodes": _n(lambda: int(made.Curve.Nodes.Count), 0),
                        "bbox": _bbox(made),
                        "stroke_from": _s(lambda: str(stroke_shape.Name))}
            if act == "point_at_offset":
                # 实测：SubPath.GetPointPositionAt(x, y, 偏移, 偏移类型) 返回 (X, Y)。
                # 偏移类型 0 = 绝对长度(mm)，1 = 归一化 0~1（前两个入参被忽略）
                spec = dict(params or {})
                off = float(spec.get("offset", 0.5))
                normalized = bool(spec.get("normalized", True))
                sub_index = int(spec.get("subpath", 1))
                sub = _s(lambda: curve.SubPaths.Item(sub_index))
                if sub is None:
                    raise CdrBadArgument(f"子路径 {sub_index} 不存在")
                got = _s(lambda: sub.GetPointPositionAt(0.0, 0.0, off,
                                                        1 if normalized else 0))
                if not isinstance(got, (tuple, list)) or len(got) < 2:
                    raise CdrComFailure("GetPointPositionAt 没有返回 (X, Y)")
                return {"action": act, "subpath": sub_index, "offset": off,
                        "normalized": normalized,
                        "point": [_round(float(got[0])), _round(float(got[1]))],
                        **info()}
            if act == "segment_at_offset":
                # 实测：SubPath.FindSegmentOffset(绝对偏移, Segment, ParamOffset, Remainder)
                # 返回 (bool, IVGSegment, 参数偏移, 余数)
                spec = dict(params or {})
                off = float(spec.get("offset", 10.0))
                sub_index = int(spec.get("subpath", 1))
                sub = _s(lambda: curve.SubPaths.Item(sub_index))
                if sub is None:
                    raise CdrBadArgument(f"子路径 {sub_index} 不存在")
                got = _s(lambda: sub.FindSegmentOffset(off, None, 0.0, 0.0))
                if not isinstance(got, (tuple, list)) or len(got) < 3:
                    raise CdrComFailure("FindSegmentOffset 没有返回 (bool, segment, param, …)")
                seg = got[1]
                return {"action": act, "subpath": sub_index, "offset": off,
                        "found": bool(got[0]),
                        "segment_index": _n(lambda: int(seg.Index))
                        if seg is not None else None,
                        "param_offset": _round(float(got[2])),
                        "remainder": _round(float(got[3])) if len(got) > 3 else None,
                        **info()}
            if act == "append_curve_segment":
                # 实测：SubPath.AppendCurveSegment(x,y,起控长,起控角,终控长,终控角,atBegin)
                spec = dict(params or {})
                sub = _s(lambda: curve.SubPaths.Item(int(spec.get("subpath", 1))))
                if sub is None:
                    raise CdrBadArgument("append_curve_segment 取不到子路径")
                before = _node_count(curve)
                with atomic(doc, label):
                    seg = sub.AppendCurveSegment(
                        float(spec.get("x", 20.0)), float(spec.get("y", 20.0)),
                        float(spec.get("start_ctrl_length", 5.0)),
                        float(spec.get("start_ctrl_angle", 90.0)),
                        float(spec.get("end_ctrl_length", 5.0)),
                        float(spec.get("end_ctrl_angle", 45.0)),
                        bool(spec.get("at_beginning", False)))
                return {"action": act, "nodes_before": before,
                        "nodes_after": _node_count(curve),
                        "segment_type": _n(lambda: int(seg.Type)) if seg is not None else None,
                        **info()}
            if act == "extract_subpath":
                # 实测：SubPath.Extract(Shape) 返回 (新 Shape, …)，把子路径抽成独立对象
                spec = dict(params or {})
                idx = int(spec.get("subpath", 1))
                sub = _s(lambda: curve.SubPaths.Item(idx))
                if sub is None:
                    raise CdrBadArgument(f"子路径 {idx} 不存在")
                before = _n(lambda: int(page.Shapes.Count), 0) or 0
                with atomic(doc, label):
                    got = sub.Extract(shape)
                made = None
                if isinstance(got, (tuple, list)):
                    for cand in got:
                        if cand is not None and _n(lambda c=cand: int(c.Type)) is not None:
                            made = cand
                            break
                else:
                    made = got
                if made is None:
                    raise CdrComFailure("SubPath.Extract 没有返回形状")
                name = _ensure_name(made)
                return {"action": act, "subpath": idx, "shape": name,
                        "type_id": _n(lambda: int(made.Type)), "bbox": _bbox(made),
                        "page_shapes_before": before,
                        "page_shapes_after": _n(lambda: int(page.Shapes.Count), 0)}
            if act == "weld_to_shape":
                # 实测：Curve.WeldWith 返回"游离曲线"，再 BindToDocument + Layer.CreateCurve
                # 才能真正落到文档里（这一步以前一直缺）
                if not other:
                    raise CdrBadArgument("weld_to_shape 需要 other（要焊接的另一个曲线对象）")
                src = find_shape(doc, page, other)
                src_curve = _s(lambda: src.Curve)
                if src_curve is None:
                    raise CdrBadArgument("other 不是曲线对象")
                with atomic(doc, label):
                    welded = curve.WeldWith(src_curve)
                    if welded is None:
                        raise CdrComFailure("Curve.WeldWith 返回空")
                    welded.BindToDocument(doc)
                    made = layer.CreateCurve(welded) if False else _active_layer(page).CreateCurve(welded)
                if made is None:
                    raise CdrComFailure("Layer.CreateCurve(焊接结果) 返回空")
                name = _ensure_name(made)
                return {"action": act, "welded_from": _s(lambda: str(src.Name)),
                        "shape": name, "type_id": _n(lambda: int(made.Type)),
                        "nodes": _n(lambda: int(made.Curve.Nodes.Count), 0),
                        "bbox": _bbox(made),
                        "page_shapes": _n(lambda: int(page.Shapes.Count), 0)}
            if act == "merge_curve":
                # Curve.AppendCurve(Source)：把另一条曲线**追加成本曲线的子路径**
                # （实测 2 节点曲线 → 4 节点 / 2 子路径，同一个对象里）
                if not other:
                    raise CdrBadArgument("merge_curve 需要 other（要并入的另一个对象）")
                src = find_shape(doc, page, other)
                src_curve = _s(lambda: src.Curve)
                if src_curve is None:
                    raise CdrBadArgument("other 不是曲线对象")
                curve.AppendCurve(src_curve)
                return {"action": act, "merged_from": _s(lambda: str(src.Name)),
                        "nodes": _node_count(curve),
                        "subpaths": _n(lambda: int(curve.SubPaths.Count), 0), **info()}
            if act == "weld_curve":
                # Curve.WeldWith(Curve) 返回**新曲线对象**（不在文档里），这里只报告结果
                if not other:
                    raise CdrBadArgument("weld_curve 需要 other（要焊接的另一个对象）")
                src = find_shape(doc, page, other)
                src_curve = _s(lambda: src.Curve)
                if src_curve is None:
                    raise CdrBadArgument("other 不是曲线对象")
                welded = _s(lambda: curve.WeldWith(src_curve))
                if welded is None:
                    raise CdrComFailure("Curve.WeldWith 返回空")
                return {"action": act, "welded_from": _s(lambda: str(src.Name)),
                        "welded_nodes": _n(lambda: int(welded.Nodes.Count), 0),
                        "welded_length": _round(_s(lambda: float(welded.Length))),
                        "note": "WeldWith 返回的是新曲线对象，不会替换原对象；"
                                "需要替换请用 cdr_path(merge_curve) 或 CDR 的焊接工具",
                        "nodes": _node_count(curve), **info()}
            if act == "is_on_curve":
                if x is None or y is None:
                    raise CdrBadArgument("is_on_curve 需要 x / y")
                state = _n(lambda: int(curve.IsOnCurve(float(x), float(y),
                                                       float(hot_area))))
                return {"action": act, "x": float(x), "y": float(y),
                        "hot_area": float(hot_area), "hit_state": state,
                        "on_curve": state == 1, **info()}
            if act == "curve_copy":
                clone = _s(lambda: curve.GetCopy())
                if clone is None:
                    raise CdrComFailure("Curve.GetCopy 返回空")
                return {"action": act, "copy_nodes": _n(lambda: int(clone.Nodes.Count), 0),
                        "copy_subpaths": _n(lambda: int(clone.SubPaths.Count), 0),
                        "copy_length": _round(_s(lambda: float(clone.Length))),
                        "nodes": _node_count(curve), **info()}
            if act == "curve_info":
                arr = _s(lambda: curve.GetCurveInfo())
                count = 0
                try:
                    count = len(arr) if arr is not None else 0
                except TypeError:
                    count = int(arr or 0)
                return {"action": act, "info_elements": count,
                        "first": (str(arr[0])[:120] if count else None),
                        "nodes": _node_count(curve), **info()}
            if act == "self_weld":
                curve.SelfWeldClosedSubpaths()
                return {"action": act, **info()}
            if act == "reverse_direction":
                curve.ReverseDirection()
                return {"action": act,
                        "clockwise": bool(_s(lambda: bool(curve.IsClockwise), False)), **info()}
            if act == "join_subpaths":
                curve.JoinTouchingSubpaths(bool(allow_reversals), float(tolerance))
                return {"action": act, "subpaths_after": _n(lambda: int(curve.SubPaths.Count), 0),
                        **info()}
            if act == "remove_overlaps":
                produced = _s(lambda: curve.RemoveOverlaps())
                return {"action": act, "result_curve": produced is not None, **info()}

            if act == "equal_divide":
                sp = sub()
                sp.EqualDivide(int(radius or 4), float(offset or 0))
                return {"action": act, "divisions": int(radius or 4), **info()}
            if act == "add_node":
                if t is None:
                    raise CdrBadArgument("add_node 需要 t（0~1 之间的位置）")
                sp = sub()
                sp.AddNodeAt(float(t), 0)
                return {"action": act, "t": t, **info()}
            if act == "break_apart":
                sp = sub()
                sp.BreakApartAt(float(t), 0)
                return {"action": act, "t": t, **info()}
            if act == "move_subpath":
                sp = sub()
                sp.Move(float(dx), float(dy))
                return {"action": act, **info()}
    return {"action": act, **info()}


# --------------------------------------------------------------------------- #
# 节点级编辑（第五波实测打通）
#
# 历史结论「拿不到节点集合」是错的：`IVGCurve.Nodes` / `.Segments` 是**属性**
# （不是方法），反解脚本只扫了方法名所以漏掉了。实测 X8 上可以：
#   curve.Nodes.Count / Item(i) / Range(a,b) / All() / First / Last
#   node.PositionX/Y / Type / Index / SubPathIndex / IsEnding / Selected
#   node.SetPosition / Move / CreateSelection / Delete
#   nodeRange.Fillet / Chamfer / Scallop / Smoothen / SetType / Move /
#              Rotate / Skew / Stretch / BreakApart / GetBoundingBox / AddToSelection
#   curve.Segments.Count / Item(i) / All() / Range(a,b)
#   segment.GetPointAt / GetPointPositionAt / GetTangentAt / GetCurvatureAt /
#           GetCurveSpeedAt / GetStartingControlPointPosition /
#           SetStartingControlPointPosition / AddNodeAt / BreakApartAt
# 注意：doc.Unit 决定坐标单位；本函数内统一用 mm_units。
# --------------------------------------------------------------------------- #
NODE_ACTIONS = (
    "list", "segments", "segment_info", "find",
    "move", "select", "delete", "type",
    "fillet", "chamfer", "scallop", "smoothen",
    "stretch", "rotate", "skew", "break_apart",
    "add_node", "break_segment", "control_points",
    # 第三十六波：节点拓扑
    "neighbors", "distance", "average", "join_nodes",
    # 第四十一波：任意节点集合（Application.CreateNodeRange + Add）
    "arbitrary",
    # 第四十七波：节点坐标搬运（GetPoint / SetPoint）
    "copy_point",
)

NODE_MUTATIONS = {
    "move", "select", "delete", "type", "fillet", "chamfer", "scallop",
    "smoothen", "stretch", "rotate", "skew", "break_apart",
    "average", "join_nodes", "arbitrary", "copy_point",
    "add_node", "break_segment", "control_points",
}


def _node_dict(node: Any) -> dict:
    return {
        "index": _n(lambda: int(node.Index)),
        "absolute_index": _n(lambda: int(node.AbsoluteIndex)),
        "subpath": _n(lambda: int(node.SubPathIndex)),
        "type": _n(lambda: int(node.Type)),
        "x": _round(_s(lambda: float(node.PositionX))),
        "y": _round(_s(lambda: float(node.PositionY))),
        "ending": bool(_s(lambda: bool(node.IsEnding), False)),
        "selected": bool(_s(lambda: bool(node.Selected), False)),
    }


def _segment_control_points(seg: Any) -> dict:
    out: dict = {"control_start": None, "control_end": None}
    try:
        start = tuple(seg.GetStartingControlPointPosition())
        out["control_start"] = [_round(v) for v in start]
    except Exception:
        pass
    try:
        end = tuple(seg.GetEndingControlPointPosition())
        out["control_end"] = [_round(v) for v in end]
    except Exception:
        pass
    return out


def _segment_dict(seg: Any, t: float = 0.5) -> dict:
    """线段详情。

    实测 X8 的 Offset 语义：OffsetType=0 是"沿曲线绝对距离"，OffsetType=1 才是
    0~1 的参数比例。这里统一用 1；点坐标走 GetPointAt(offset, 1)。
    """
    start_pt = _s(lambda: seg.GetPointAt(0.0, 1))
    end_pt = _s(lambda: seg.GetPointAt(1.0, 1))
    mid_pt = _s(lambda: seg.GetPointAt(float(t), 1))
    info = {
        "index": _n(lambda: int(seg.Index)),
        "absolute_index": _n(lambda: int(seg.AbsoluteIndex)),
        "type": _n(lambda: int(seg.Type)),
        "length": _round(_s(lambda: float(seg.Length))),
        "start": ([_round(_s(lambda: float(start_pt.x))),
                   _round(_s(lambda: float(start_pt.y)))] if start_pt else None),
        "end": ([_round(_s(lambda: float(end_pt.x))),
                 _round(_s(lambda: float(end_pt.y)))] if end_pt else None),
        "point_at_t": ([_round(_s(lambda: float(mid_pt.x))),
                        _round(_s(lambda: float(mid_pt.y)))] if mid_pt else None),
        "curvature": _round(_s(lambda: float(seg.GetCurvatureAt(float(t), 1)))),
        "speed": _round(_s(lambda: float(seg.GetCurveSpeedAt(float(t), 1)))),
    }
    info.update(_segment_control_points(seg))
    return info


def _node_count(curve: Any) -> int:
    return _n(lambda: int(curve.Nodes.Count), 0)


def _node_range(curve: Any, node: int | None, start: int | None, end: int | None,
                all_nodes: bool) -> Any:
    count = _node_count(curve)
    if count < 1:
        raise CdrBadArgument("曲线没有节点")
    if node is not None:
        idx = int(node)
        if not 1 <= idx <= count:
            raise CdrBadArgument(f"node 越界：{idx}（共 {count} 个）")
        curve.Nodes.Item(idx).CreateSelection()
        return curve.Selection()
    if start is not None or end is not None:
        a = int(start if start is not None else 1)
        b = int(end if end is not None else count)
        if not 1 <= a <= b <= count:
            raise CdrBadArgument(f"节点范围越界：[{a},{b}]（共 {count} 个）")
        return curve.Nodes.Range(a, b)
    if all_nodes:
        return curve.Nodes.All()
    return None


def _node_ref(curve: Any, node: int | None) -> Any:
    count = _node_count(curve)
    idx = int(node) if node is not None else 1
    if not 1 <= idx <= count:
        raise CdrBadArgument(f"node 越界：{idx}（共 {count} 个）")
    return curve.Nodes.Item(idx)


def _segment_ref(curve: Any, segment_index: int | None) -> Any:
    count = _n(lambda: int(curve.Segments.Count), 0)
    idx = int(segment_index) if segment_index is not None else 1
    if not 1 <= idx <= count:
        raise CdrBadArgument(f"segment 越界：{idx}（共 {count} 条）")
    return curve.Segments.Item(idx)


def node_ops(doc: Any, page_index: int | None, target: dict, action: str, *,
             node: int | None = None, start: int | None = None, end: int | None = None,
             all_nodes: bool = False, x: float | None = None, y: float | None = None,
             dx: float = 0.0, dy: float = 0.0, node_type: int | None = None,
             radius: float = 0.0, distance_a: float = 0.0, distance_b: float = 0.0,
             ratio_x: float = 1.0, ratio_y: float = 1.0, angle: float = 0.0,
             angle_x: float = 0.0, angle_y: float = 0.0, offset: float = 0.5,
             segment_index: int | None = None, control: str | None = None,
             hot_area: float = -1.0, params: dict | None = None,
             label: str = "cdr_node") -> dict:
    """节点 / 线段级编辑。

    action（只读）：list / segments / segment_info / find
    action（变更）：move / select / delete / type / fillet / chamfer / scallop /
                    smoothen / stretch / rotate / skew / break_apart /
                    add_node / break_segment / control_points
    选择器：node=第几个节点 ｜ start/end=节点范围 ｜ all_nodes=true（读操作默认全部）
    """
    act = (action or "list").lower()
    if act not in NODE_ACTIONS:
        raise CdrBadArgument(f"未知 action {action!r}；可用：{', '.join(NODE_ACTIONS)}")

    with mm_units(doc):
        page = _page(doc, page_index)
        shape = find_shape(doc, page, target)
        curve = _s(lambda: shape.Curve)
        if curve is None:
            raise CdrBadArgument("目标不是曲线对象（可先用 cdr_shape_ops(op='to_curves') 转曲）")

        def base() -> dict:
            return {
                "name": _s(lambda: str(shape.Name)),
                "closed": bool(_s(lambda: bool(curve.Closed), False)),
                "clockwise": bool(_s(lambda: bool(curve.IsClockwise), False)),
                "nodes": _node_count(curve),
                "segments": _n(lambda: int(curve.Segments.Count), 0),
                "subpaths": _n(lambda: int(curve.SubPaths.Count), 0),
            }

        # ---------- 只读 ----------
        if act == "list":
            count = _node_count(curve)
            items = [_node_dict(curve.Nodes.Item(i)) for i in range(1, min(count, 500) + 1)]
            xs = [n["x"] for n in items if n["x"] is not None]
            ys = [n["y"] for n in items if n["y"] is not None]
            bbox = None
            if xs and ys:
                bbox = {"x": min(xs), "y": min(ys),
                        "width": _round(max(xs) - min(xs)),
                        "height": _round(max(ys) - min(ys))}
            return {"action": act, **base(), "node_list": items, "bbox": bbox,
                    "truncated": count > 500}
        if act == "arbitrary":
            # 实测：Application.CreateNodeRange() + Add(node) 可以凑出**不连续**的节点集合，
            # 再 Move/Rotate/Skew/Stretch/AutoReduce/Smoothen/SetType/Fillet/Chamfer/Scallop
            spec = dict(params or {})
            wanted = [int(i) for i in (spec.get("nodes") or [])]
            if not wanted:
                raise CdrBadArgument("arbitrary 需要 params.nodes=[节点序号…]")
            total = _node_count(curve)
            rng = _s(lambda: _app_of(doc).CreateNodeRange())
            if rng is None:
                raise CdrComFailure("Application.CreateNodeRange 返回空")
            for idx in wanted:
                if not 1 <= idx <= total:
                    raise CdrBadArgument(f"节点 {idx} 越界（共 {total} 个）")
                rng.Add(curve.Nodes.Item(idx))
            count = _n(lambda: int(rng.Count), 0) or 0
            op = str(spec.get("op", "move")).lower()
            with atomic(doc, label):
                if op == "move":
                    rng.Move(float(spec.get("dx", 0.0)), float(spec.get("dy", 0.0)),
                             1, False)
                elif op == "rotate":
                    rng.Rotate(float(spec.get("angle", 0.0)), False,
                               float(spec.get("cx", 0.0)), float(spec.get("cy", 0.0)))
                elif op == "skew":
                    rng.Skew(float(spec.get("angle_x", 0.0)),
                             float(spec.get("angle_y", 0.0)), False,
                             float(spec.get("cx", 0.0)), float(spec.get("cy", 0.0)))
                elif op == "stretch":
                    rng.Stretch(float(spec.get("ratio_x", 1.0)),
                                float(spec.get("ratio_y", 1.0)), False,
                                float(spec.get("cx", 0.0)), float(spec.get("cy", 0.0)))
                elif op == "autoreduce":
                    rng.AutoReduce(float(spec.get("tolerance", 0.5)))
                elif op == "smoothen":
                    rng.Smoothen(int(spec.get("smoothness", 0)))
                elif op == "set_type":
                    rng.SetType(int(spec.get("node_type", 0)))
                elif op == "delete":
                    rng.Delete()
                elif op == "fillet":
                    rng.Fillet(float(spec.get("radius", 2.0)), False)
                elif op == "chamfer":
                    rng.Chamfer(float(spec.get("radius", 2.0)),
                                float(spec.get("radius2", spec.get("radius", 2.0))), False)
                elif op == "scallop":
                    rng.Scallop(float(spec.get("radius", 2.0)), False)
                elif op == "select":
                    rng.CreateSelection()
                else:
                    raise CdrBadArgument(
                        "arbitrary 的 op 可用 move/rotate/skew/stretch/autoreduce/"
                        "smoothen/set_type/delete/fillet/chamfer/scallop/select")
            bbox = _s(lambda: rng.GetBoundingBox())
            return {"action": act, **base(), "op": op, "nodes_selected": count,
                    "nodes_after": _node_count(curve),
                    "range_bbox": ([_round(float(v)) for v in bbox[1:]]
                                   if isinstance(bbox, (tuple, list)) and len(bbox) >= 5
                                   else None),
                    "note": "CreateNodeRange 支持不连续节点；Move 的第三个参数是"
                            "范围内起始索引（实测 1 可用）"}
        if act == "copy_point":
            # 实测：Node.GetPoint() 返回 IVGPoint，Node.SetPoint(point) 可直接搬运坐标
            if node is None or end is None:
                raise CdrBadArgument("copy_point 需要 node（源）与 end（目标）节点序号")
            src_node = curve.Nodes.Item(int(node))
            dst_node = curve.Nodes.Item(int(end))
            point = _s(lambda: src_node.GetPoint())
            if point is None:
                raise CdrComFailure("Node.GetPoint 返回空")
            with atomic(doc, label):
                dst_node.SetPoint(point)
            return {"action": act, **base(), "from": int(node), "to": int(end),
                    "source_pos": [_round(_s(lambda: float(src_node.PositionX))),
                                   _round(_s(lambda: float(src_node.PositionY)))],
                    "target_pos": [_round(_s(lambda: float(dst_node.PositionX))),
                                   _round(_s(lambda: float(dst_node.PositionY)))]}
        if act == "neighbors":
            # Next/Previous 是**方法**（返回 IVGNode），不是属性
            if node is None:
                raise CdrBadArgument("neighbors 需要 node")
            ref = curve.Nodes.Item(int(node))
            nxt = _s(lambda: ref.Next())
            prv = _s(lambda: ref.Previous())
            return {"action": act, **base(),
                    "node": int(node),
                    "next": _n(lambda: int(nxt.Index)) if nxt is not None else None,
                    "previous": _n(lambda: int(prv.Index)) if prv is not None else None}
        if act == "distance":
            if node is None or end is None:
                raise CdrBadArgument("distance 需要 node 与 end（两个节点序号）")
            a1 = curve.Nodes.Item(int(node))
            b1 = curve.Nodes.Item(int(end))
            dist = _s(lambda: float(a1.GetDistanceFrom(b1)))
            return {"action": act, **base(), "node": int(node), "other": int(end),
                    "distance": _round(dist)}
        if act in ("average", "join_nodes"):
            if node is None or end is None:
                raise CdrBadArgument(f"{act} 需要 node 与 end（两个**开放曲线末端**节点）")
            a1 = curve.Nodes.Item(int(node))
            b1 = curve.Nodes.Item(int(end))
            if act == "average":
                a1.AveragePositionWith(b1, False)
            else:
                a1.JoinWith(b1)
            return {"action": act, **base(),
                    "node": int(node), "other": int(end),
                    "pos": [_round(_s(lambda: float(curve.Nodes.Item(int(node)).PositionX))),
                            _round(_s(lambda: float(curve.Nodes.Item(int(node)).PositionY)))],
                    "note": "X8 要求两个节点都是开放曲线的末端节点，否则报"
                            "「两个节点都必须是开放曲线的末端节点」"}
        if act == "segments":
            count = _n(lambda: int(curve.Segments.Count), 0)
            items = [_segment_dict(curve.Segments.Item(i), float(offset))
                     for i in range(1, min(count, 300) + 1)]
            return {"action": act, **base(), "segment_list": items,
                    "truncated": count > 300}
        if act == "segment_info":
            seg = _segment_ref(curve, segment_index)
            return {"action": act, **base(), "segment": _segment_dict(seg, float(offset))}
        if act == "find":
            if x is None or y is None:
                raise CdrBadArgument("find 需要 x 与 y")
            found_node = _s(lambda: curve.FindNodeAtPoint(float(x), float(y), float(hot_area)))
            found_seg = _s(lambda: curve.FindSegmentAtPoint(
                float(x), float(y), HotArea=float(hot_area)))
            seg = None
            off = None
            if found_seg is not None:
                if isinstance(found_seg, (tuple, list)):
                    seg = _segment_dict(found_seg[0], float(offset))
                    off = _round(found_seg[1]) if len(found_seg) > 1 else None
                else:
                    seg = _segment_dict(found_seg, float(offset))
            return {"action": act,
                    "node": _node_dict(found_node) if found_node is not None else None,
                    "segment": seg, "segment_offset": off}

        # ---------- 变更 ----------
        if act == "add_node":
            seg = _segment_ref(curve, segment_index)
            with atomic(doc, label):
                new_node = seg.AddNodeAt(float(offset))
            return {"action": act, "segment": segment_index or 1, "offset": offset,
                    "node": _node_dict(new_node) if new_node is not None else None,
                    **base()}
        if act == "break_segment":
            seg = _segment_ref(curve, segment_index)
            with atomic(doc, label):
                new_node = seg.BreakApartAt(float(offset))
            return {"action": act, "segment": segment_index or 1, "offset": offset,
                    "node": _node_dict(new_node) if new_node is not None else None,
                    **base()}
        if act == "control_points":
            seg = _segment_ref(curve, segment_index)
            if control and x is not None and y is not None:
                with atomic(doc, label):
                    if str(control).lower().startswith("s"):
                        seg.SetStartingControlPointPosition(float(x), float(y))
                        applied = "start"
                    else:
                        seg.SetEndingControlPointPosition(float(x), float(y))
                        applied = "end"
                return {"action": act, "segment": segment_index or 1,
                        "set": applied, **_segment_control_points(seg)}
            return {"action": act, "segment": segment_index or 1,
                    **_segment_control_points(seg)}

        rng = _node_range(curve, node, start, end, all_nodes)
        if rng is None:
            raise CdrBadArgument(
                "该 action 需要指定节点：node=n / start+end / all_nodes=true")

        with atomic(doc, label):
            if act == "move":
                if x is not None and y is not None and node is not None:
                    _node_ref(curve, node).SetPosition(float(x), float(y))
                else:
                    rng.Move(float(dx), float(dy))
            elif act == "select":
                rng.CreateSelection()
            elif act == "delete":
                rng.Delete()
            elif act == "type":
                if node_type is None:
                    raise CdrBadArgument("type 需要 node_type（实测有效范围 0~2）")
                rng.SetType(int(node_type))
            elif act == "fillet":
                rng.Fillet(float(radius), True)
            elif act == "chamfer":
                rng.Chamfer(float(distance_a), float(distance_b), True)
            elif act == "scallop":
                rng.Scallop(float(radius), True)
            elif act == "smoothen":
                rng.Smoothen(int(max(0, min(100, int(radius or 100)))))
            elif act == "stretch":
                rng.Stretch(float(ratio_x), float(ratio_y))
            elif act == "rotate":
                rng.Rotate(float(angle), False)
            elif act == "skew":
                rng.Skew(float(angle_x), float(angle_y))
            elif act == "break_apart":
                rng.BreakApart()

        result = {"action": act, **base()}
        if act == "select":
            result["selected"] = _n(lambda: int(curve.Selection().Count), 0)
        if act == "type":
            read_back = [_n(lambda i=i: int(curve.Nodes.Item(i).Type))
                         for i in range(1, min(_node_count(curve), 500) + 1)]
            result["types"] = sorted(set(t for t in read_back if t is not None))
            if node_type is not None and int(node_type) not in result["types"]:
                result.setdefault("warnings", []).append(
                    "Type 写入未被 X8 接受（实测：节点类型只能读，改类型请用 smoothen）")
        if act in {"delete", "break_apart", "fillet", "chamfer", "scallop", "smoothen"}:
            result["bbox"] = _bbox(shape)
        return result


# --------------------------------------------------------------------------- #
# 对象样式（第六波实测打通）
#
# 历史结论「样式类别常量 0~8 全失败」也是误判：`CreateStyleFromShape` 的
# Category 是**字符串**，不是数字常量。实测可用类别：
#   fill / outline（图形）｜character / paragraph（文字）
# 创建后可以用 `Shape.ApplyStyle("样式名")` 套用（传 Style 对象会报「未找到样式」），
# 也可以 Rename / Delete / ToString 读定义。
# --------------------------------------------------------------------------- #
STYLE_ACTIONS = ("list", "get", "create", "apply", "rename", "delete",
                 "set_property", "derived", "export_sheet", "import_sheet",
                 "color_styles", "create_color_styles", "override_props", "inherited",
                 "get_property", "clear_property", "set_based_on", "assign",
                 "fill_file")
STYLE_MUTATIONS = {"create", "apply", "rename", "delete", "set_property",
                   "import_sheet", "create_color_styles", "set_based_on", "assign"}
STYLE_CATEGORIES = ("fill", "outline", "character", "paragraph")


def _find_style(doc: Any, name: str | None) -> Any:
    """按名称找样式。实测 Styles.Find 对「刚用 CreateStyleFromShape 建的样式」可能返回空，
    因此再加一层「遍历 Styles 比对 Name」的兜底。"""
    if not name:
        raise CdrBadArgument("需要 style_name")
    style = _s(lambda: doc.StyleSheet.Styles.Find(str(name)))
    if style is not None:
        return style
    count = _n(lambda: int(doc.StyleSheet.Styles.Count), 0) or 0
    for i in range(1, count + 1):
        cand = _s(lambda i=i: doc.StyleSheet.Styles.Item(i))
        if cand is not None and _s(lambda cand=cand: str(cand.Name)) == str(name):
            return cand
    raise CdrBadArgument(f"未找到样式 {name!r}（可先用 action='list' 查看）")


def _style_dict(style: Any) -> dict:
    return {
        "style_name": _s(lambda: str(style.Name)),
        "category": _s(lambda: str(style.CategoryName)) or None,
        "display_name": _s(lambda: str(style.DisplayName)) or None,
        "display_category": _s(lambda: str(style.DisplayCategoryName)) or None,
        "based_on": _s(lambda: str(style.BasedOn.Name)) or None,
    }


def _style_readable_props(style: Any, limit: int = 120) -> dict:
    """把样式的可读属性逐项转成字符串（实测 69 个属性，绝大多数可读）。

    注意：`fill` / `outline` 这类"分组名"不能直接读，真正可读的是
    primaryColor / secondaryColor / angle / type 等叶子属性。
    """
    names = _s(lambda: list(style.GetAllPropertyNames()), []) or []
    out: dict = {}
    for prop in names[:limit]:
        try:
            out[str(prop)] = style.GetPropertyAsString(str(prop))
        except Exception:
            continue
    return out


def _style_derived(style: Any) -> list[dict]:
    count = _n(lambda: int(style.DerivedStyles.Count), 0)
    return [_style_dict(style.DerivedStyles.Item(i))
            for i in range(1, min(count, 200) + 1)]


def style_ops(doc: Any, page_index: int | None, action: str, *,
              target: dict | None = None, category: str = "fill",
              name: str | None = None, new_name: str | None = None,
              based_on: str | None = None, property: str | None = None,
              value: Any = None, sheet_path: str | None = None,
              use_fills: bool = True, use_outlines: bool = True,
              selected_only: bool = False, harmonies: int = 0,
              convert_colors_to: int = 99,
              label: str = "cdr_style") -> dict:
    """对象样式（Object Styles）的列出 / 创建 / 套用 / 改名 / 删除 / 改属性 /
    样式表导入导出 / 颜色样式。

    action:
      list                     列出文档里的样式（名称 + 类别）
      get（name）              读样式定义（可读属性、继承链、ToString JSON）
      create（target,category,name[,based_on]）
                               从形状创建样式，或基于已有样式派生（based_on）；
                               category 实测可用：fill / outline / character / paragraph
      apply（target,name）     把样式套到对象上
      rename（name,new_name）  样式改名
      delete（name）           删除样式
      set_property（name,property,value）  改样式属性（标量走字符串；*color 结尾走颜色对象）
      derived（name）          列派生样式
      export_sheet（sheet_path） 把当前样式表存成 .cdst
      import_sheet（sheet_path） 从 .cdst 载入样式（实测可恢复被删样式）
      color_styles             列出颜色样式（AllColorStyles）
      create_color_styles（harmonies, selected_only, convert_colors_to）
                               从文档/选区生成颜色样式与颜色协调
    """
    act = (action or "list").lower()
    if act not in STYLE_ACTIONS:
        raise CdrBadArgument(f"未知 action {action!r}；可用：{', '.join(STYLE_ACTIONS)}")

    with mm_units(doc):
        sheet = doc.StyleSheet
        if act == "list":
            count = _n(lambda: int(sheet.Styles.Count), 0)
            styles = [_style_dict(sheet.Styles.Item(i))
                      for i in range(1, min(count, 500) + 1)]
            return {"action": act, "count": count, "styles": styles}
        if act == "get":
            style = _find_style(doc, name)
            props = _s(lambda: list(style.GetAllPropertyNames()), []) or []
            return {"action": act, **_style_dict(style),
                    "properties": [str(p) for p in props[:120]],
                    "property_count": len(props),
                    "readable_properties": _style_readable_props(style),
                    "derived": _style_derived(style),
                    "definition": (_s(lambda: str(style.ToString()), "") or "")[:1200]}
        if act == "fill_file":
            # 实测：Style.Fill 是 IVGStyleFill，SaveFill(FileName, Metadata) /
            # LoadFill(FileName, Metadata) 都**接受路径 + IVGFillMetadata**
            # （metadata 由 Application.CreateFillMetadata() 造；返回 False 表示没存成功）
            style = _find_style(doc, name)
            sf = _s(lambda: style.Fill)
            if sf is None:
                raise CdrComFailure("取不到 Style.Fill")
            if not sheet_path:
                raise CdrBadArgument("fill_file 需要 sheet_path（.cpf 文件路径）")
            meta = _s(lambda: _app_of(doc).CreateFillMetadata())
            direction = str(property or "save").lower()
            if direction == "load":
                got = _s(lambda: sf.LoadFill(str(sheet_path), meta))
                ok = bool(got[0]) if isinstance(got, (tuple, list)) and got else False
                return {"action": act, "direction": "load", "file": str(sheet_path),
                        "ok": ok, "raw": str(got)[:80]}
            ok = bool(_s(lambda: sf.SaveFill(str(sheet_path), meta), False))
            return {"action": act, "direction": "save", "file": str(sheet_path),
                    "ok": ok, "fill_type": _n(lambda: int(sf.Type)),
                    "metadata_ok": meta is not None}
        if act == "set_based_on":
            # 实测：Style.SetBasedOn(父样式**名称字符串**) —— 传样式对象会「类型不匹配」
            style = _find_style(doc, name)
            if not based_on:
                raise CdrBadArgument("set_based_on 需要 based_on（父样式名）")
            ok = bool(_s(lambda: bool(style.SetBasedOn(str(based_on))), False))
            parent = _s(lambda: style.BasedOn)
            return {"action": act, **_style_dict(style), "parent_name": str(based_on),
                    "applied": ok,
                    "based_on_name": (_s(lambda: str(parent.Name))
                                      if parent is not None else None)}
        if act == "assign":
            # 实测：Style.Assign(另一个 IVGStyle) 可用（拷属性，不改父级链）
            style = _find_style(doc, name)
            if not based_on:
                raise CdrBadArgument("assign 需要 based_on（源样式名）")
            src = _find_style(doc, based_on)
            style.Assign(src)
            return {"action": act, **_style_dict(style),
                    "source": str(based_on),
                    "definition": (_s(lambda: str(style.ToString()), "") or "")[:200]}
        if act == "override_props":
            # 实测 Style.GetOverridePropertyNames() 返回 ('primaryColor','secondaryColor')
            style = _find_style(doc, name)
            names = _s(lambda: list(style.GetOverridePropertyNames()), []) or []
            return {"action": act, **_style_dict(style),
                    "override_properties": [str(n) for n in names],
                    "count": len(names)}
        if act == "inherited":
            style = _find_style(doc, name)
            prop_name = str(property or "primaryColor")
            flag = bool(_s(lambda: bool(style.IsPropertyInherited(prop_name)), False))
            return {"action": act, **_style_dict(style), "property": prop_name,
                    "inherited": flag}
        if act == "get_property":
            style = _find_style(doc, name)
            prop_name = str(property or "primaryColor")
            got = _s(lambda: style.GetProperty(prop_name))
            if got is None:
                avail = _s(lambda: [str(n) for n in style.GetOverridePropertyNames()], [])
                raise CdrBadArgument(
                    f"样式 {name!r} 取不到属性 {prop_name!r}；可用覆盖属性：{avail}")
            value: Any = got
            try:
                if hasattr(got, "ToString"):
                    value = str(got.ToString())
            except Exception:  # noqa: BLE001
                value = "COM"
            return {"action": act, **_style_dict(style), "property": prop_name,
                    "value": value,
                    "type": _n(lambda: int(got.Type)) if hasattr(got, "Type") else None}
        if act == "clear_property":
            style = _find_style(doc, name)
            prop_name = str(property or "primaryColor")
            cleared = bool(_s(lambda: bool(style.ClearProperty(prop_name)), False))
            return {"action": act, **_style_dict(style), "property": prop_name,
                    "cleared": cleared}
        if act == "export_sheet":
            if not sheet_path:
                raise CdrBadArgument("export_sheet 需要 sheet_path（.cdst）")
            doc.SaveStyleSheet(str(sheet_path))
            return {"action": act, "saved": str(sheet_path),
                    "exists": os.path.exists(str(sheet_path)),
                    "bytes": os.path.getsize(str(sheet_path))
                    if os.path.exists(str(sheet_path)) else 0,
                    "styles": _n(lambda: int(sheet.Styles.Count), 0)}
        if act == "import_sheet":
            if not sheet_path:
                raise CdrBadArgument("import_sheet 需要 sheet_path（.cdst）")
            if not os.path.exists(str(sheet_path)):
                raise CdrBadArgument(f"样式表不存在：{sheet_path}")
            with atomic(doc, label):
                doc.LoadStyleSheet(str(sheet_path))
            count = _n(lambda: int(sheet.Styles.Count), 0)
            return {"action": act, "loaded": str(sheet_path), "count": count,
                    "styles": [str(sheet.Styles.Item(i).Name)
                               for i in range(1, min(count, 200) + 1)]}
        if act == "color_styles":
            colors = sheet.AllColorStyles
            count = _n(lambda: int(colors.Count), 0)
            items = []
            for i in range(1, min(count, 500) + 1):
                col = colors.Item(i)
                items.append({"name": _s(lambda c=col: str(c.GetName())) or None,
                              "color": _s(lambda c=col: str(c.ToString()))})
            return {"action": act, "count": count, "color_styles": items}
        if act == "create_color_styles":
            before = _n(lambda: int(sheet.AllColorStyles.Count), 0)
            with atomic(doc, label):
                affected = _n(lambda: int(doc.CreateColorStyles(
                    bool(use_fills), bool(use_outlines), bool(selected_only),
                    int(harmonies), int(convert_colors_to))), 0)
            after = _n(lambda: int(sheet.AllColorStyles.Count), 0)
            return {"action": act, "affected": affected,
                    "before": before, "after": after,
                    "harmonies": int(harmonies),
                    "convert_colors_to": int(convert_colors_to)}
        if act == "create":
            cat = str(category or "fill").lower()
            if cat not in STYLE_CATEGORIES:
                raise CdrBadArgument(
                    f"category 可用：{', '.join(STYLE_CATEGORIES)}"
                    "（fill/outline 用于图形，character/paragraph 用于文字）")
            style_name = str(name or f"cdr_style_{_short_id()}")
            target_name = None
            if based_on:
                parent = _find_style(doc, based_on)
                parent_name = _s(lambda: str(parent.Name))
                with atomic(doc, label):
                    created_style = sheet.CreateStyle(cat, parent_name, style_name, False)
                if created_style is None:
                    raise CdrBadArgument("CreateStyle 返回空对象（检查 category/based_on）")
                return {"action": act, **_style_dict(created_style),
                        "category": cat, "based_on": parent_name,
                        "created": [style_name]}
            page = _page(doc, page_index)
            shape = find_shape(doc, page, target or {})
            target_name = _s(lambda: str(shape.Name))
            with atomic(doc, label):
                created = sheet.CreateStyleFromShape(shape, cat, style_name, False)
            names: list[str] = []
            try:
                for i in range(1, int(created.Count) + 1):
                    names.append(str(created.Item(i).Name))
            except Exception:
                pass
            return {"action": act, "created": names or [style_name],
                    "style_name": names[0] if names else style_name,
                    "category": cat, "target": target_name}
        if act == "apply":
            page = _page(doc, page_index)
            shape = find_shape(doc, page, target or {})
            style = _find_style(doc, name)
            style_name = _s(lambda: str(style.Name))
            with atomic(doc, label):
                shape.ApplyStyle(str(style_name))
            return {"action": act, "applied": style_name,
                    "target": _s(lambda: str(shape.Name)), "bbox": _bbox(shape)}
        if act == "rename":
            style = _find_style(doc, name)
            if not new_name:
                raise CdrBadArgument("rename 需要 new_name")
            old = _s(lambda: str(style.Name))
            with atomic(doc, label):
                renamed = bool(_s(lambda: style.Rename(str(new_name)), False))
            return {"action": act, "renamed": renamed, "old": old, "new": new_name}
        if act == "set_property":
            style = _find_style(doc, name)
            if not property:
                raise CdrBadArgument("set_property 需要 property")
            prop = str(property)
            readback = None
            with atomic(doc, label):
                if prop.lower().endswith("color"):
                    # 实测：颜色属性用 SetProperty(名字, IVGColor) 最稳；
                    # SetPropertyAsString 需要给带引号的颜色串。
                    style.SetProperty(prop, _color(doc, value))
                else:
                    ok = bool(_s(lambda: style.SetPropertyAsString(prop, str(value)), False))
                    if not ok:
                        raise CdrBadArgument(f"SetPropertyAsString({prop!r}) 未成功")
            readback = _s(lambda: style.GetPropertyAsString(prop))
            return {"action": act, **_style_dict(style), "property": prop,
                    "value": value, "readback": readback}
        if act == "derived":
            style = _find_style(doc, name)
            return {"action": act, **_style_dict(style),
                    "derived": _style_derived(style)}
        style = _find_style(doc, name)
        old = _s(lambda: str(style.Name))
        with atomic(doc, label):
            deleted = bool(_s(lambda: style.Delete(), False))
        return {"action": "delete", "deleted": deleted, "style_name": old}


# --------------------------------------------------------------------------- #
# 图层
# --------------------------------------------------------------------------- #
def list_layers(doc: Any, page_index: int | None = None) -> list[dict]:
    page = _page(doc, page_index)
    active = _s(lambda: str(doc.ActiveLayer.Name), "")
    out = []
    for j in range(1, int(page.Layers.Count) + 1):
        layer = page.Layers.Item(j)
        out.append({
            "index": j,
            "name": _s(lambda: str(layer.Name)),
            "active": _s(lambda: str(layer.Name)) == active,
            "visible": bool(_s(lambda: bool(layer.Visible), True)),
            "editable": bool(_s(lambda: bool(layer.Editable), True)),
            "printable": bool(_s(lambda: bool(layer.Printable), True)),
            "shapes": _n(lambda: int(layer.Shapes.Count), 0),
        })
    return out


def layer_action(doc: Any, action: str, name: str | None = None,
                 page_index: int | None = None, index: int | None = None,
                 visible: bool | None = None, editable: bool | None = None,
                 printable: bool | None = None, move: str | None = None,
                 move_below: str | None = None, shape_name: str | None = None,
                 shape_type: int | None = None, static_id: int | None = None,
                 recursive: bool = True) -> dict:
    page = _page(doc, page_index)
    act = (action or "").lower()

    if act == "create":
        if not name:
            raise CdrBadArgument("action=create 需要 name")
        layer = page.CreateLayer(name)
        return {"created": _s(lambda: str(layer.Name), name), "layers": list_layers(doc, page_index)}

    layer = _find_layer(page, name=name, index=index)

    if act == "info":
        # 实测 IsGuidesLayer 是**属性**（不是方法），FindShape 需要 Name/Type/StaticID/Recursive
        return {"name": _s(lambda: str(layer.Name)),
                "index": _n(lambda: int(layer.Index)),
                "visible": bool(_s(lambda: bool(layer.Visible), True)),
                "editable": bool(_s(lambda: bool(layer.Editable), True)),
                "printable": bool(_s(lambda: bool(layer.Printable), True)),
                "master": bool(_s(lambda: bool(layer.Master), False)),
                "is_guides_layer": bool(_s(lambda: bool(layer.IsGuidesLayer), False))}
    if act == "paste_ex":
        # 实测：Application.CreateStructPasteOptions() → Layer.PasteEx(opts) 返回 ShapeRange
        app = _app_of(doc)
        opts = _s(lambda: app.CreateStructPasteOptions())
        if opts is None:
            raise CdrComFailure("CreateStructPasteOptions 返回空")
        rng = _s(lambda: layer.PasteEx(opts))
        if rng is None:
            raise CdrComFailure("Layer.PasteEx 返回空（剪贴板可能为空）")
        count = _n(lambda: int(rng.Count), 0) or 0
        return {"pasted": count,
                "names": [_s(lambda i=i: str(rng.Item(i).Name))
                          for i in range(1, min(count, 20) + 1)],
                "layer": _s(lambda: str(layer.Name))}
    if act == "find":
        if not shape_name:
            raise CdrBadArgument("action=find 需要 shape_name（要查找的对象名）")
        found = _s(lambda: layer.FindShape(str(shape_name),
                                           int(shape_type or 0),
                                           int(static_id or 0),
                                           bool(recursive)))
        if found is None:
            return {"found": False, "shape_name": str(shape_name)}
        return {"found": True, "shape_name": str(shape_name),
                "name": _s(lambda: str(found.Name)),
                "type_id": _n(lambda: int(found.Type)),
                "bbox": _bbox(found)}

    if act == "activate":
        layer.Activate()
        return {"active": _s(lambda: str(doc.ActiveLayer.Name))}
    if act in {"set", "update"}:
        if visible is not None:
            layer.Visible = bool(visible)
        if editable is not None:
            layer.Editable = bool(editable)
        if printable is not None:
            layer.Printable = bool(printable)
        return {"layers": list_layers(doc, page_index)}
    if act == "move":
        target = _find_layer(page, name=move, index=None)
        layer.MoveAbove(target)
        return {"layers": list_layers(doc, page_index)}
    if act == "move_below":
        target = _find_layer(page, name=move_below, index=None)
        layer.MoveBelow(target)
        return {"layers": list_layers(doc, page_index)}
    if act in {"above", "below"}:
        neighbor = layer.GetAbove(True) if act == "above" else layer.GetBelow(True)
        return {"action": act, "layer": _s(lambda: str(layer.Name)),
                "neighbor": _s(lambda: str(neighbor.Name)) if neighbor is not None else None}
    if act == "paste":
        pasted = _s(lambda: layer.Paste())
        return {"action": act, "pasted": pasted is not None,
                "name": _s(lambda: str(pasted.Name)) if pasted is not None else None,
                "type": SHAPE_TYPES.get(_n(lambda: int(pasted.Type), -1), "unknown")
                        if pasted is not None else None,
                "note": "使用系统剪贴板；剪贴板为空时 pasted=false"}
    if act == "delete":
        layer.Delete()
        return {"layers": list_layers(doc, page_index)}
    raise CdrBadArgument(
        f"未知 action {action!r}；可用 list/create/activate/set/move/move_below/"
        f"above/below/paste/delete")


def _find_layer(page: Any, name: str | None, index: int | None) -> Any:
    if index is not None:
        return page.Layers.Item(int(index))
    if name:
        for j in range(1, int(page.Layers.Count) + 1):
            layer = page.Layers.Item(j)
            if str(layer.Name) == name:
                return layer
        raise CdrBadArgument(f"找不到图层 {name!r}")
    return page.ActiveLayer


# --------------------------------------------------------------------------- #
# 形状：读取
# --------------------------------------------------------------------------- #
def shape_to_dict(shape: Any, index: int) -> dict:
    bbox = _bbox(shape)
    d: dict[str, Any] = {
        "index": index,
        "name": _s(lambda: str(shape.Name)) or "",
        "type": SHAPE_TYPES.get(_n(lambda: int(shape.Type), -1), f"type_{_n(lambda: int(shape.Type), -1)}"),
        "type_id": _n(lambda: int(shape.Type), -1),
        "bbox": bbox,
        "center": {
            "x": _round(_s(lambda: float(shape.CenterX))),
            "y": _round(_s(lambda: float(shape.CenterY))),
        },
    }
    if d["type_id"] == TEXT_TYPE:
        d["text"] = _s(lambda: str(shape.Text.Story.Text), "")
    return d


def _bbox(shape: Any) -> dict | None:
    box = _s(lambda: tuple(shape.GetBoundingBox()))
    if not box or len(box) < 4:
        return None
    return {
        "x": _round(float(box[0])),
        "y": _round(float(box[1])),
        "width": _round(float(box[2])),
        "height": _round(float(box[3])),
    }


def _iter_page_shapes(page: Any):
    """遍历整页所有图层上的对象，产出 (页内序号, 图层序号, 层内序号, shape)。

    实测坑：`page.Shapes` 在**新建图层后只反映活动图层**，会导致"对象明明在却找不到"。
    所以这里按图层顺序自己编号，页内序号从 1 开始、跨图层连续。
    """
    n = 0
    lay_count = _n(lambda: int(page.Layers.Count), 0)
    for li in range(1, lay_count + 1):
        try:
            layer = page.Layers.Item(li)
            cnt = int(layer.Shapes.Count)
        except Exception:
            continue
        for j in range(1, cnt + 1):
            try:
                shape = layer.Shapes.Item(j)
            except Exception:
                continue
            n += 1
            yield n, li, j, shape


def list_shapes(doc: Any, page_index: int | None = None) -> list[dict]:
    with mm_units(doc):
        page = _page(doc, page_index)
        out = []
        for idx, li, _j, shape in _iter_page_shapes(page):
            item = shape_to_dict(shape, idx)
            item["layer"] = li
            out.append(item)
        return out


def total_shapes(doc: Any) -> int:
    with mm_units(doc):
        total = 0
        for i in range(1, _n(lambda: int(doc.Pages.Count), 1) + 1):
            try:
                page = doc.Pages.Item(i)
                total += sum(1 for _ in _iter_page_shapes(page))
            except Exception:
                pass
        return total


def find_shapes(doc: Any, name: str | None = None, type_id: int | None = None,
                page_index: int | None = None) -> list[dict]:
    """按名字/类型查找（跨图层）。

    实测：X8 的 `Page.FindShapes(name, type, recursive)` 传类型 ID 时不生效（返回 0），
    所以统一用自己的遍历，行为可预期。
    """
    with mm_units(doc):
        page = _page(doc, page_index)
        out: list[dict] = []
        for idx, li, _j, shape in _iter_page_shapes(page):
            if name and str(_s(lambda: shape.Name, "")) != name:
                continue
            if type_id and int(_s(lambda: int(shape.Type), -1)) != int(type_id):
                continue
            item = shape_to_dict(shape, idx)
            item["layer"] = li
            out.append(item)
        return out


def select_shapes(doc: Any, page_index: int | None = None, name: str | None = None,
                  type_id: int | None = None, rect: dict | None = None,
                  all_shapes: bool = False, remove: bool = False,
                  clear: bool = False) -> dict:
    """把对象加入选择集（供对齐/布尔/组合使用）。

    remove=true 改为从选择集里移除；clear=true 直接清空选择集。
    """
    with mm_units(doc):
        page = _page(doc, page_index)
        if clear:
            try:
                doc.Selection().RemoveAll()
            except Exception:
                pass
            return {"selected": 0, "cleared": True}
        if not remove:
            try:
                doc.Selection().RemoveAll()
            except Exception:
                pass
        picked = 0
        if rect:
            if remove:
                raise CdrBadArgument("remove 不支持 rect 框选；请按 name/type_id 移除")
            page.SelectShapesFromRectangle(
                float(rect["x1"]), float(rect["y1"]),
                float(rect["x2"]), float(rect["y2"]), bool(rect.get("touch", True)),
            )
            picked = _n(lambda: int(doc.Selection().Shapes.Count), 0)
        else:
            for _n_, _li, _j, shape in _iter_page_shapes(page):
                if name and str(_s(lambda: shape.Name, "")) != name:
                    continue
                if type_id and int(_s(lambda: int(shape.Type), -1)) != int(type_id):
                    continue
                if remove:
                    shape.RemoveFromSelection()
                else:
                    shape.AddToSelection()
                picked += 1
        result = {"selected": _n(lambda: int(doc.Selection().Shapes.Count), 0)}
        if remove:
            result["removed"] = picked
        return result


def find_shape(doc: Any, page: Any, target: dict) -> Any:
    if not isinstance(target, dict):
        raise CdrBadArgument('target 必须是对象，如 {"index": 1} 或 {"name": "cdr_ab12"}')

    if target.get("selection"):
        shapes = doc.Selection().Shapes
        if int(shapes.Count) == 0:
            raise CdrBadArgument("当前没有选中任何对象")
        return shapes.Item(1)

    if "name" in target:
        want = str(target["name"])
        for _idx, _li, _j, shape in _iter_page_shapes(page):
            if str(_s(lambda: shape.Name, "")) == want:
                return shape
        raise CdrBadArgument(f"当前页找不到名为 {want} 的对象")

    if "index" in target:
        idx = int(target["index"])
        for n, _li, _j, shape in _iter_page_shapes(page):
            if n == idx:
                return shape
        total = sum(1 for _ in _iter_page_shapes(page))
        raise CdrBadArgument(f"index 越界：{idx}（当前页共 {total} 个对象）")

    raise CdrBadArgument("target 需要 index / name / selection 之一")


def selection_range(doc: Any) -> Any:
    """把当前选择集转成 ShapeRange（布尔/对齐/组合都要用它）。"""
    app = _app_of(doc)
    sr = app.CreateShapeRange()
    shapes = doc.Selection().Shapes
    count = int(shapes.Count)
    for i in range(1, count + 1):
        sr.Add(shapes.Item(i))
    return sr


# --------------------------------------------------------------------------- #
# 形状：创建
# --------------------------------------------------------------------------- #
CREATABLE = (
    "rectangle", "ellipse", "polygon", "spiral", "grid", "line", "bezier",
    "polyline", "bitmap", "text", "paragraph_text", "curve_segment", "guide",
    "ole_file",
    # 第三十八波：显式 (left, top, right, bottom) 版本，走 Layer.CreateRectangle/
    # CreateEllipse/CreatePolygon 的完整参数形态
    "rect_ltrb", "ellipse_ltrb", "polygon_full",
    # 第四十一波：CreateBitmap2 / CreateBitmapRect（用 IVGImage 直接摆位图）
    "bitmap_at", "bitmap_rect",
    # 第四十七波：宽文本（带语言/字符集）构造
    "text_wide", "paragraph_wide",
    # 第五十三波：B 样条（控制点 → IVGBSpline → Layer.CreateBSpline）
    "bspline",
    # 第五十五波：连接器 / 线性标注 / 角度标注（Application.CreateSnapPoint 造端点）
    "connector", "connector_ex", "right_angle_connector",
    "linear_dimension", "angular_dimension",
)


def create_items(doc: Any, page_index: int | None, items: Iterable[dict],
                 label: str = "cdr_create") -> list[dict]:
    created: list[dict] = []
    with mm_units(doc):
        page = _page(doc, page_index)
        layer = _active_layer(page)
        with atomic(doc, label):
            for item in items:
                shape = _create_one(doc, layer, item)
                if shape is None:      # guide 不是形状
                    created.append({"name": None, "type": "guide", "bbox": None})
                    continue
                created.append({
                    "name": _s(lambda: str(shape.Name)),
                    "type": SHAPE_TYPES.get(_n(lambda: int(shape.Type), -1),
                                            f"type_{_n(lambda: int(shape.Type), -1)}"),
                    "bbox": _bbox(shape),
                })
    return created


def _active_layer(page: Any) -> Any:
    try:
        return page.ActiveLayer
    except Exception:
        return page.Layers.Item(1)


def _create_one(doc: Any, layer: Any, item: dict) -> Any:
    kind = str(item.get("type", "")).lower()
    g = lambda k, d=None: item.get(k, d)  # noqa: E731

    if kind in {"rectangle", "rect"}:
        shape = layer.CreateRectangle2(
            float(g("x", 0)), float(g("y", 0)), float(g("width", 10)), float(g("height", 10)),
            float(g("radius_ul", g("radius", 0))), float(g("radius_ur", g("radius", 0))),
            float(g("radius_lr", g("radius", 0))), float(g("radius_ll", g("radius", 0))),
        )
    elif kind in {"rect_ltrb"}:
        shape = layer.CreateRectangle(
            float(g("left", 0)), float(g("top", 10)), float(g("right", 20)),
            float(g("bottom", 0)),
            float(g("radius_ul", g("radius", 0))), float(g("radius_ur", g("radius", 0))),
            float(g("radius_lr", g("radius", 0))), float(g("radius_ll", g("radius", 0))),
        )
    elif kind in {"ellipse_ltrb"}:
        shape = layer.CreateEllipse(
            float(g("left", 0)), float(g("top", 10)), float(g("right", 20)),
            float(g("bottom", 0)),
            float(g("start_angle", 0)), float(g("end_angle", 360)),
            bool(g("pie", False)),
        )
    elif kind in {"polygon_full"}:
        shape = layer.CreatePolygon(
            float(g("left", 0)), float(g("top", 10)), float(g("right", 20)),
            float(g("bottom", 0)),
            int(g("sides", 5)), int(g("subpaths", 1)), int(g("complexity", 1)),
            bool(g("star", False)), int(g("star_complexity", 0)),
            int(g("max_complexity", 1)),
        )
    elif kind in {"ellipse", "circle", "arc", "pie"}:
        pie = kind in {"arc", "pie"} or bool(g("pie", False))
        shape = layer.CreateEllipse2(
            float(g("cx", 0)), float(g("cy", 0)),
            float(g("rx", g("r", 5))), float(g("ry", g("rx", g("r", 5)))),
            float(g("start_angle", 0)), float(g("end_angle", 360)), pie,
        )
    elif kind in {"polygon", "star"}:
        shape = layer.CreatePolygon2(
            float(g("cx", 0)), float(g("cy", 0)), float(g("radius", g("r", 10))),
            int(g("sides", 5 if kind == "polygon" else 5)),
            float(g("angle", 0)), float(g("inner_radius", 0)),
            bool(g("star", kind == "star")) or kind == "star",
            max(-1.0, min(1.0, float(g("sharpness", 0)))),
        )
    elif kind == "spiral":
        shape = layer.CreateSpiral(
            float(g("left", 0)), float(g("top", 10)), float(g("right", 10)), float(g("bottom", 0)),
            int(g("revolutions", 3)), int(g("spiral_type", 1)), float(g("growth", 1.2)),
        )
    elif kind in {"grid", "grid_boxes"}:
        shape = layer.CreateGridBoxes(
            float(g("left", 0)), float(g("top", 10)), float(g("right", 10)), float(g("bottom", 0)),
            int(g("cols", 3)), int(g("rows", 3)),
        )
    elif kind in {"line", "line_segment"}:
        shape = layer.CreateLineSegment(
            float(g("x1", 0)), float(g("y1", 0)), float(g("x2", 10)), float(g("y2", 10)),
        )
    elif kind in {"bezier", "curve"}:
        shape = layer.CreateCurveSegment2(
            float(g("x1", 0)), float(g("y1", 0)),
            float(g("c1x", 3)), float(g("c1y", 3)),
            float(g("c2x", 7)), float(g("c2y", 7)),
            float(g("x2", 10)), float(g("y2", 10)),
        )
    elif kind in {"connector", "connector_ex", "right_angle_connector",
                  "linear_dimension", "angular_dimension"}:
        # 实测：Application.CreateSnapPoint(x, y) 返回 IVGSnapPoint，可直接当
        # CreateConnector/CreateLinearDimension 的端点参数用
        app = _app_of(doc)

        def _snap(key_x, key_y, dx=0.0, dy=0.0):
            return app.CreateSnapPoint(float(g(key_x, dx)), float(g(key_y, dy)))

        p1 = _snap("x1", "y1", 0.0, 0.0)
        p2 = _snap("x2", "y2", 40.0, 20.0)
        if kind == "connector":
            shape = layer.CreateConnector(p1, p2)
        elif kind == "connector_ex":
            shape = layer.CreateConnectorEx(int(g("connector_type", 0)), p1, p2)
        elif kind == "right_angle_connector":
            shape = layer.CreateRightAngleConnector(p1, p2, float(g("corner_radius", 5.0)))
        elif kind == "linear_dimension":
            shape = layer.CreateLinearDimension(
                int(g("dimension_type", 0)), p1, p2,
                True, float(g("text_x", (g("x1", 0.0) + g("x2", 40.0)) / 2)),
                float(g("text_y", g("y1", 0.0) - 10.0)),
                int(g("style", 0)), int(g("precision", 2)), True,
                int(g("units", 0)), int(g("placement", 0)), True, False, True,
                str(g("prefix", "")), str(g("suffix", "")),
                float(g("outline_width", 0.2)), None, None,
                str(g("text_font", "Arial")), float(g("text_size", 8.0)), None)
        else:
            pc = _snap("cx", "cy", 0.0, 0.0)
            p3 = _snap("x3", "y3", 40.0, 0.0)
            shape = layer.CreateAngularDimension(
                pc, p1, p3, float(g("text_x", 20.0)), float(g("text_y", -10.0)),
                int(g("precision", 2)), True, int(g("units", 0)), False, True,
                str(g("prefix", "")), str(g("suffix", "")),
                float(g("outline_width", 0.2)), None, None,
                str(g("text_font", "Arial")), float(g("text_size", 8.0)), None)
        if shape is None:
            raise CdrComFailure(f"创建 {kind} 失败")
    elif kind in {"bspline"}:
        # 实测：Document.CreateBSpline(控制点数, 闭合) → IVGBSpline；
        # 用 InsertControlPoint(序号, x, y, Clamped) 填点（**两端必须 Clamped=True**），
        # 再 Layer.CreateBSpline(bs) 落成 Type 21 的样条对象（转曲后 Type 3）
        pts = item.get("points") or []
        if len(pts) < 3:
            raise CdrBadArgument("bspline 需要 points：[[x,y], …] 至少 3 个控制点")
        closed = bool(item.get("closed", False))
        bs = doc.CreateBSpline(len(pts), closed)
        for idx, point in enumerate(pts, start=1):
            clamped = idx in (1, len(pts))
            bs.InsertControlPoint(idx, float(point[0]), float(point[1]), clamped)
        shape = layer.CreateBSpline(bs)
        if shape is None:
            raise CdrComFailure("Layer.CreateBSpline 返回空")
        if bool(item.get("to_curves", False)):
            shape.ConvertToCurves()
    elif kind in {"polyline", "points", "path"}:
        pts = item.get("points") or []
        if len(pts) < 2:
            raise CdrBadArgument("polyline 需要 points：[[x,y], [x,y], ...] 至少 2 个点")
        x0, y0 = float(pts[0][0]), float(pts[0][1])
        x1, y1 = float(pts[1][0]), float(pts[1][1])
        # X8 没有"创建空曲线"的入口（CreateCurve 必须给 Source）。
        # 用第一段线段当载体，再往同一个子路径追加剩余点——
        # 这样不会像旧的"退化线段 + CreateSubPath"那样留下重复节点。
        shape = layer.CreateLineSegment(x0, y0, x1, y1)
        sub = shape.Curve.SubPaths.Item(1)
        for point in pts[2:]:
            sub.AppendLineSegment(float(point[0]), float(point[1]))
        closed = bool(item.get("closed", True))
        if closed:
            sub.AppendLineSegment(x0, y0)
        sub.Closed = closed
    elif kind in {"bitmap_at", "bitmap_rect"}:
        # 实测：CreateBitmap2(x,y,w,h,Image,Alpha) 与 CreateBitmapRect(IVGRect,Image,Alpha)
        # 都要 IVGImage（用 Document.CreateImage 造）与 Alpha=None
        px_w = int(g("pixels_w", g("px_width", 32)))
        px_h = int(g("pixels_h", g("px_height", 32)))
        image_type = int(g("image_type", 4))
        image = doc.CreateImage(image_type, max(1, px_w), max(1, px_h),
                                _color(doc, g("color", "#FFFFFF")))
        if kind == "bitmap_at":
            shape = layer.CreateBitmap2(
                float(g("x", 0)), float(g("y", 0)),
                float(g("width", 20)), float(g("height", 20)), image, None)
        else:
            rect = _app_of(doc).CreateRect()
            rect.SetRect(float(g("x", 0)), float(g("y", 0)),
                         float(g("width", 20)), float(g("height", 20)))
            shape = layer.CreateBitmapRect(rect, image, None)
    elif kind in {"bitmap", "image"}:
        # 从内存像素缓冲区创建位图（实测：Document.CreateImage → IVGImage，
        # 再 Layer.CreateBitmap 落地；ImageAlpha 必须传对象或 None，传 0 会 TypeError）
        x = float(g("x", 0))
        y = float(g("y", 0))
        w = float(g("width", 50))
        h = float(g("height", 50))
        px_w = int(g("pixels_w", g("px_width", 64)))
        px_h = int(g("pixels_h", g("px_height", 64)))
        if px_w < 1 or px_h < 1 or px_w > 8192 or px_h > 8192:
            raise CdrBadArgument("bitmap 的 pixels_w/pixels_h 需在 1~8192 之间")
        image_type = int(g("image_type", 4))
        base = _color(doc, g("color", "#FFFFFF"))
        image = doc.CreateImage(image_type, px_w, px_h, base)
        for area in (g("fills", []) or []):
            if not isinstance(area, dict):
                continue
            image.FillArea(int(area.get("x", 0)), int(area.get("y", 0)),
                           int(area.get("width", 1)), int(area.get("height", 1)),
                           _color(doc, area.get("color", "#000000")))
        shape = layer.CreateBitmap(x, y, x + w, y + h, image, None)
    elif kind in {"text", "artistic_text"}:
        if g("font"):
            ensure_font_table(_app_of(doc))
        shape = layer.CreateArtisticText(
            float(g("x", 0)), float(g("y", 0)), str(g("text", "")),
            int(g("language_id", 0)), int(g("charset", 0)),
            str(g("font", "")), float(g("size", 0)),
            _tristate(g("bold", False)), _tristate(g("italic", False)),
            _underline(g("underline", False)),
            int(_align(g("align", "left"))),
        )
        restate_text_format(shape, item)
    elif kind in {"text_wide"}:
        # 实测：CreateArtisticTextWide 与窄版参数一致，但语言/字符集是显式枚举
        if g("font"):
            ensure_font_table(_app_of(doc))
        shape = layer.CreateArtisticTextWide(
            float(g("x", 0)), float(g("y", 0)), str(g("text", "")),
            int(g("language_id", 0)), int(g("charset", 0)),
            str(g("font", "")), float(g("size", 0)),
            _tristate(g("bold", False)), _tristate(g("italic", False)),
            _underline(g("underline", False)),
            int(_align(g("align", "left"))),
        )
        restate_text_format(shape, item)
    elif kind in {"paragraph_wide"}:
        if g("font"):
            ensure_font_table(_app_of(doc))
        shape = layer.CreateParagraphTextWide(
            float(g("x", 0)), float(g("y", 0)), float(g("x2", 60)), float(g("y2", 20)),
            str(g("text", "")),
            int(g("language_id", 0)), int(g("charset", 0)),
            str(g("font", "")), float(g("size", 0)),
            _tristate(g("bold", False)), _tristate(g("italic", False)),
            _underline(g("underline", False)),
            int(_align(g("align", "left"))),
        )
        restate_text_format(shape, item)
    elif kind in {"paragraph_text", "paragraph"}:
        if g("font"):
            ensure_font_table(_app_of(doc))
        shape = layer.CreateParagraphText(
            float(g("x", 0)), float(g("y", 0)), float(g("x2", 60)), float(g("y2", 20)),
            str(g("text", "")),
            int(g("language_id", 0)), int(g("charset", 0)),
            str(g("font", "")), float(g("size", 0)),
            _tristate(g("bold", False)), _tristate(g("italic", False)),
            _underline(g("underline", False)),
            int(_align(g("align", "left"))),
        )
        restate_text_format(shape, item)
    elif kind == "curve_segment":
        # 实测：Layer.CreateCurveSegment 直接给端点 + 两端控制柄长度/角度就能造曲线
        shape = layer.CreateCurveSegment(
            float(g("x1", 0)), float(g("y1", 0)), float(g("x2", 20)), float(g("y2", 10)),
            float(g("start_ctrl_length", 5)), float(g("start_ctrl_angle", 90)),
            float(g("end_ctrl_length", 5)), float(g("end_ctrl_angle", 45)),
        )
    elif kind == "guide":
        layer.CreateGuide(float(g("x1", 0)), float(g("y1", 0)),
                          float(g("x2", 100)), float(g("y2", 0)))
        return None
    elif kind in {"ole_file", "ole"}:
        path = str(g("path", ""))
        if not path:
            raise CdrBadArgument("ole_file 需要 path（绝对路径）")
        if not os.path.exists(path):
            raise CdrBadArgument(f"ole_file 路径不存在：{path}")
        shape = layer.CreateOLEObjectFromFile(
            path, bool(g("link", False)), bool(g("display_as_icon", False)),
            str(g("caption", "")), None)
    else:
        raise CdrBadArgument(f"不支持的类型 {kind}；可用：{', '.join(CREATABLE)}")

    _apply_style(doc, shape, item)
    name = str(item.get("name") or "").strip()
    shape.Name = name if name else f"{NAME_PREFIX}{_short_id()}"
    return shape


def _align(value: Any) -> int:
    if isinstance(value, int):
        return value
    return TEXT_ALIGN.get(str(value).lower(), 0)


_CREATED = {"n": 0}
_OUTLINE_STYLE_CACHE: dict = {}


def _short_id() -> str:
    _CREATED["n"] += 1
    return hashlib.sha1(f"{_CREATED['n']}-{time.time()}".encode()).hexdigest()[:8]


def _ensure_name(shape: Any) -> str:
    """对象名字为空时补一个可寻址的名字（CreateBoundary / Freeze 会产出无名对象）。"""
    name = _s(lambda: str(shape.Name), "") or ""
    if not name:
        name = f"{NAME_PREFIX}{_short_id()}"
        shape.Name = name
    return name


def _apply_style(doc: Any, shape: Any, spec: dict) -> None:
    """填充 / 描边 / 效果 / 文字样式。单项失败不阻断创建。"""
    fill = spec.get("fill")
    if fill is not None:
        try:
            apply_fill(doc, shape, fill)
        except Exception:
            pass

    outline = spec.get("outline")
    if outline is not None:
        try:
            apply_outline(doc, shape, outline)
        except Exception:
            pass

    shadow = spec.get("drop_shadow")
    if shadow:
        try:
            s = shadow if isinstance(shadow, dict) else {}
            shape.CreateDropShadow(
                int(s.get("type", 1)), float(s.get("opacity", 50)), float(s.get("feather", 0)),
                float(s.get("offset_x", 2)), float(s.get("offset_y", -2)),
                _color(doc, s.get("color", "#000000")),
            )
        except Exception:
            pass

    if int(_s(lambda: int(shape.Type), -1)) == TEXT_TYPE:
        tri = {"bold", "italic", "underline"}
        for key, prop in (("font", "Font"), ("size", "Size"),
                          ("bold", "Bold"), ("italic", "Italic"), ("underline", "Underline")):
            if spec.get(key) not in (None, ""):
                try:
                    setattr(shape.Text.Story, prop,
                            _underline(spec[key]) if key == "underline"
                            else (_tristate(spec[key]) if key in tri else spec[key]))
                except Exception:
                    pass


def _palette_at(doc: Any, index: int) -> Any:
    app = _app_of(doc)
    count = _n(lambda: int(app.Palettes.Count), 0)
    if not 1 <= int(index) <= count:
        raise CdrBadArgument(f"palette_index 越界：{index}（当前已打开 {count} 个调色板）")
    return app.Palettes.Item(int(index))


def _find_palette(doc: Any, token: Any) -> Any:
    """按序号 / 名称 / GUID 找已打开的调色板。"""
    app = _app_of(doc)
    if isinstance(token, int) or (isinstance(token, str) and token.isdigit()):
        return _palette_at(doc, int(token))
    if not token:
        raise CdrBadArgument("需要 palette_index 或 palette（调色板名/GUID）")
    want = str(token)
    count = _n(lambda: int(app.Palettes.Count), 0)
    for i in range(1, count + 1):
        pal = app.Palettes.Item(i)
        if str(_s(lambda p=pal: str(p.Name), "")) == want or \
                str(_s(lambda p=pal: str(p.Identifier), "")) == want:
            return pal
    raise CdrBadArgument(
        f"未找到调色板 {want!r}；可先 cdr_doc(action='palette_open_fixed', palette_id=...) 打开，"
        "或 cdr_doc(action='palettes') 查看已打开的调色板")


def _spot_color(doc: Any, spec: dict) -> Any:
    """按调色板 + 序号/名称创建专色（实测 PANTONE 调色板可用）。

    spec 支持：{"palette_index": 1, "color_index": 10}；
               {"palette_index": 1, "name": "PANTONE Process Yellow CV"}；
               {"palette": "PANTONE MATCHING SYSTEM - Corel 8", "name": ...}
    """
    if not isinstance(spec, dict):
        raise CdrBadArgument("spot 需要对象：{palette_index|palette, color_index|name, tint}")
    app = _app_of(doc)
    pal = _find_palette(doc, spec.get("palette_index", spec.get("palette")))
    spot_name = spec.get("name")
    if not spot_name:
        ci = int(spec.get("color_index", 1))
        color_count = _n(lambda: int(pal.ColorCount), 0)
        if not 1 <= ci <= color_count:
            raise CdrBadArgument(f"color_index 越界：{ci}（该调色板共 {color_count} 色）")
        spot_name = _s(lambda: str(pal.Color(ci).GetName()))
    if not spot_name:
        raise CdrBadArgument("无法确定专色名（给 name 或有效的 color_index）")
    tint = float(spec.get("tint", 100))
    color = app.CreateSpotColorByName(
        str(_s(lambda: str(pal.Identifier), "")), str(spot_name), tint)
    if color is None:
        raise CdrBadArgument(f"创建专色失败：{spot_name!r}")
    return color


def _outline_style(app: Any, pattern: tuple) -> Any:
    """创建/复用自定义虚线样式（虚拟样式必须 AddStyle 进集合才能赋值给 Outline）。"""
    cache = _OUTLINE_STYLE_CACHE.setdefault("styles", {})
    style = cache.get(pattern)
    if style is not None:
        try:
            _ = int(style.Index)
            return style
        except Exception:
            cache.pop(pattern, None)
    virtual = app.CreateOutlineStyle(*[float(v) for v in pattern])
    style = app.OutlineStyles.AddStyle(virtual)
    if style is not None:
        cache[pattern] = style
    return style


def apply_fill(doc: Any, shape: Any, spec: Any) -> str:
    """填充：纯色 / 渐变 / 纹理 / 专色 / 无色。返回实际应用的种类。"""
    if spec in (None, False, "none", {"type": "none"}):
        shape.Fill.ApplyNoFill()
        return "none"

    if isinstance(spec, dict) and str(spec.get("type", "")).lower() == "spot":
        shape.Fill.ApplyUniformFill(_spot_color(doc, spec))
        return "spot"

    if isinstance(spec, dict) and spec.get("type") in {"fountain", "gradient"}:
        start = _color(doc, spec.get("start", "#FFFFFF"))
        end = _color(doc, spec.get("end", "#000000"))
        shape.Fill.ApplyFountainFill(
            start, end,
            int(spec.get("gradient_type", 1)),      # 1=线性 2=辐射 3=锥形 4=方形
            float(spec.get("angle", 0)),
            int(spec.get("steps", 256)),
            float(spec.get("edge_pad", 0)),
            float(spec.get("mid_point", 50)),
            int(spec.get("blend", 0)),
            float(spec.get("center_x", 0)),
            float(spec.get("center_y", 0)),
        )
        return "fountain"

    if isinstance(spec, dict) and str(spec.get("type", "")).lower() == "pattern":
        # 图案填充：优先用内存 PatternCanvas（不需要 .pat 文件）
        idx = int(spec.get("canvas_index", 1))
        front = _color(doc, spec.get("front", "#000000"))
        end = _color(doc, spec.get("end", "#FFFFFF"))
        shape.Fill.ApplyPatternFill(
            int(spec.get("pattern_type", 0)), str(spec.get("file", "") or ""), idx,
            front, end, bool(spec.get("transform_with_shape", False)))
        return "pattern"

    if isinstance(spec, dict) and spec.get("type") == "texture":
        shape.Fill.ApplyTextureFill(str(spec.get("name", "")), str(spec.get("library", "")))
        return "texture"

    if isinstance(spec, dict) and spec.get("type") == "hatch":
        # 实测：Fill.ApplyHatchFill(库名, HatchNameOrIndex, 背景色, 随形变换, 随形缩放,
        # 世界坐标, FillScale, LineScale, FillAngle, FillSkew)。
        # **HatchNameOrIndex 必须是 int**（传字符串报「发生意外」）
        shape.Fill.ApplyHatchFill(
            str(spec.get("library", "") or ""), int(spec.get("index", 1)),
            _color(doc, spec.get("back", "#FFFFFF")),
            bool(spec.get("transform_with_shape", True)),
            bool(spec.get("scale_lines_with_shape", True)),
            bool(spec.get("use_world_coordinates", False)),
            float(spec.get("fill_scale", 1.0)), float(spec.get("line_scale", 1.0)),
            float(spec.get("fill_angle", 0.0)), float(spec.get("fill_skew", 0.0)))
        return "hatch"

    if isinstance(spec, dict) and spec.get("type") == "hatch_custom":
        # 实测：Fill.ApplyCustomHatchFill(18 参) 可用，Style 可传 None
        shape.Fill.ApplyCustomHatchFill(
            float(spec.get("angle", 45.0)), float(spec.get("spacing", 2.0)),
            float(spec.get("shift", 0.0)), float(spec.get("origin_x", 0.0)),
            float(spec.get("origin_y", 0.0)), float(spec.get("width", 0.5)),
            _color(doc, spec.get("color", "#000000")), None,
            float(spec.get("dashdot_length", 0.0)), float(spec.get("pen_width", 0.3)),
            _color(doc, spec.get("back", "#FFFFFF")),
            bool(spec.get("transform_with_shape", True)),
            bool(spec.get("scale_lines_with_shape", True)),
            bool(spec.get("use_world_coordinates", False)),
            float(spec.get("fill_scale", 1.0)), float(spec.get("line_scale", 1.0)),
            float(spec.get("fill_angle", 45.0)), float(spec.get("fill_skew", 0.0)))
        return "hatch_custom"

    if isinstance(spec, dict) and spec.get("type") == "postscript":
        # 实测可用：按 PostScript 底纹索引填充
        shape.Fill.ApplyPostscriptFill(int(spec.get("index", 1)))
        return "postscript"

    shape.Fill.ApplyUniformFill(_color(doc, spec))
    return "uniform"


def apply_outline(doc: Any, shape: Any, spec: Any) -> None:
    """描边：数字=宽度；或者 {width, color, caps, join, behind_fill, dash, start_arrow, end_arrow}。

    `dash` / `start_arrow` / `end_arrow` 传**索引**（CorelDRAW 线条样式库 / 箭头库的下标）。
    实测：`Outline.Style = 1` 会「类型不匹配」，必须赋 `app.OutlineStyles.Item(n)` 对象；
    箭头同理要赋 `app.ArrowHeads.Item(n)` 对象。
    """
    if spec in (None, "none"):
        shape.Outline.SetNoOutline()
        return
    if isinstance(spec, (int, float)):
        shape.Outline.Width = float(spec)
        return
    if not isinstance(spec, dict):
        raise CdrBadArgument("outline 需为数字或对象")

    if spec.get("none"):
        shape.Outline.SetNoOutline()
        return
    if str(spec.get("type", "")).lower() == "spot":
        _outline_color(doc, shape, spec)

    if spec.get("dash_pattern"):
        pattern = tuple(float(v) for v in spec["dash_pattern"])
        if len(pattern) not in (2, 4, 6):
            raise CdrBadArgument("dash_pattern 需 2/4/6 个数字（实线段/空白段交替长度）")
        app = _app_of(doc)
        style = _outline_style(app, pattern)
        if style is None:
            raise CdrBadArgument("自定义虚线样式创建失败")
        shape.Outline.Style = style
        if spec.get("save_outline_styles"):
            app.OutlineStyles.Save()

    app = None
    if "width" in spec:
        shape.Outline.Width = float(spec["width"])
    if "color" in spec:
        _outline_color(doc, shape, spec["color"])
    for key, attr in (("caps", "LineCaps"), ("join", "LineJoin"),
                      ("behind_fill", "BehindFill"), ("scale_with_shape", "ScaleWithShape")):
        if key in spec and spec[key] is not None:
            try:
                setattr(shape.Outline, attr, spec[key])
            except Exception:
                pass

    if spec.get("dash") is not None:
        app = app or _app_of(doc)
        shape.Outline.Style = app.OutlineStyles.Item(int(spec["dash"]))
    if spec.get("start_arrow") is not None:
        app = app or _app_of(doc)
        shape.Outline.StartArrow = app.ArrowHeads.Item(int(spec["start_arrow"]))
    if spec.get("end_arrow") is not None:
        app = app or _app_of(doc)
        shape.Outline.EndArrow = app.ArrowHeads.Item(int(spec["end_arrow"]))


def _outline_color(doc: Any, shape: Any, color: Any) -> None:
    if isinstance(color, dict) and str(color.get("type", "")).lower() == "spot":
        spot = _spot_color(doc, color)
        try:
            shape.Outline.Color.CopyAppearance(spot)
            return
        except Exception:
            shape.Outline.Color.StringAssign(str(spot.ToString()))
            return
    if isinstance(color, str):
        r, g, b = _hex_to_rgb(color)
        try:
            shape.Outline.Color.RGBAssign(r, g, b)
            return
        except Exception:
            pass
        shape.Outline.Color.CMYKAssign(*_rgb_to_cmyk(r, g, b))
        return
    if isinstance(color, dict):
        if all(k in color for k in ("c", "m", "y", "k")):
            shape.Outline.Color.CMYKAssign(
                int(color["c"]), int(color["m"]), int(color["y"]), int(color["k"]))
        elif all(k in color for k in ("r", "g", "b")):
            shape.Outline.Color.RGBAssign(int(color["r"]), int(color["g"]), int(color["b"]))
        else:
            raise CdrBadArgument("描边颜色需为 #RRGGBB 或 {c,m,y,k} 或 {r,g,b}")


def _color(doc: Any, spec: Any) -> Any:
    """支持 #RRGGBB / {c,m,y,k} / {r,g,b} / {gray} / {lab} / {hls} / "registration" /
    {"type":"spot", "palette_index":…, "color_index"/"name":…, "tint":100}。

    X8 实测可用：CreateRGBColor / CreateCMYKColor / CreateGrayColor / CreateLabColor /
    CreateHLSColor / CreateRegistrationColor；专色需要先用
    `cdr_doc(action="palette_open_fixed")` 打开 PANTONE 等调色板，再走
    `CreateSpotColorByName`（见 `_spot_color`）。
    """
    app = _app_of(doc)
    if isinstance(spec, dict) and str(spec.get("type", "")).lower() == "spot":
        return _spot_color(doc, spec)
    if isinstance(spec, str):
        s = spec.strip().lower()
        if s in {"registration", "套版色", "#registration"}:
            return app.CreateRegistrationColor()
        r, g, b = _hex_to_rgb(spec)
        return app.CreateRGBColor(r, g, b)
    if isinstance(spec, dict):
        if all(k in spec for k in ("c", "m", "y", "k")):
            return app.CreateCMYKColor(
                int(spec["c"]), int(spec["m"]), int(spec["y"]), int(spec["k"]))
        if all(k in spec for k in ("r", "g", "b")):
            return app.CreateRGBColor(int(spec["r"]), int(spec["g"]), int(spec["b"]))
        if "gray" in spec:
            return app.CreateGrayColor(float(spec["gray"]))
        lab = spec.get("lab")
        if isinstance(lab, (list, tuple)) and len(lab) == 3:
            return app.CreateLabColor(float(lab[0]), float(lab[1]), float(lab[2]))
        if all(k in spec for k in ("l", "a", "b")):
            return app.CreateLabColor(float(spec["l"]), float(spec["a"]), float(spec["b"]))
        hls = spec.get("hls")
        if isinstance(hls, (list, tuple)) and len(hls) == 3:
            return app.CreateHLSColor(float(hls[0]), float(hls[1]), float(hls[2]))
        if all(k in spec for k in ("h", "l", "s")):
            return app.CreateHLSColor(float(spec["h"]), float(spec["l"]), float(spec["s"]))
    raise CdrBadArgument(
        '颜色需为 "#RRGGBB" / {c,m,y,k} / {r,g,b} / {"gray":50} / {"lab":[l,a,b]} / '
        '{"hls":[h,l,s]} / "registration"')


def _apply_transparency(doc: Any, shape: Any, spec: Any) -> str:
    """透明：数字=均匀透明度；{type:"fountain",start,end,...} 或 "none"。"""
    if spec in (None, "none", 0):
        shape.Transparency.ApplyNoTransparency()
        return "none"
    if isinstance(spec, (int, float)):
        shape.Transparency.ApplyUniformTransparency(int(spec))
        return "uniform"
    if isinstance(spec, dict):
        if spec.get("type") == "texture":
            # 实测：Transparency.ApplyTextureTransparency(名称, 库, ?, ?) 返回 IVGTextureFill，
            # Transparency.Type 变 4
            shape.Transparency.ApplyTextureTransparency(
                str(spec.get("texture", "BARK")),
                str(spec.get("library", "Textures")),
                int(spec.get("texture_type", 0)), 0)
            return "texture"
        if spec.get("type") == "pattern":
            shape.Transparency.ApplyPatternTransparency(
                int(spec.get("pattern_type", 0)), str(spec.get("file", "") or ""),
                int(spec.get("canvas_index", 1)),
                int(spec.get("front", 0)), int(spec.get("back", 100)),
                bool(spec.get("transform_with_shape", False)))
            return "pattern"
        if spec.get("type") in {"fountain", "gradient"}:
            shape.Transparency.ApplyFountainTransparency(
                int(spec.get("start", 0)), int(spec.get("end", 100)),
                int(spec.get("gradient_type", 1)), float(spec.get("angle", 0)),
                int(spec.get("steps", 256)), float(spec.get("edge_pad", 0)),
                float(spec.get("mid_point", 50)),
                float(spec.get("center_x", 0)), float(spec.get("center_y", 0)),
            )
            return "fountain"
        if "value" in spec:
            shape.Transparency.ApplyUniformTransparency(int(spec["value"]))
            return "uniform"
    raise CdrBadArgument(
        'transparency 需为数字（0~100）、"none" 或 {type:"fountain",start,end,angle}')


def _apply_gradient_tweak(shape: Any, spec: Any) -> None:
    """微调已有渐变（实测可用：SetAngle / SetEdgePad / Translate / SetTransformations / MakeOpaque）。"""
    if not isinstance(spec, dict):
        raise CdrBadArgument("gradient 需为对象，如 {angle:90, edge_pad:10}")
    fountain = shape.Fill.Fountain
    if "angle" in spec:
        fountain.SetAngle(float(spec["angle"]))
    if "edge_pad" in spec:
        fountain.SetEdgePad(float(spec["edge_pad"]))
    if "mid_point" in spec:
        # MidPoint 没有 setter，只能通过重设渐变；这里保持兼容性忽略
        pass
    if "translate" in spec:
        dx, dy = spec["translate"]
        fountain.Translate(float(dx), float(dy))
    if "transformations" in spec:
        t = list(spec["transformations"])
        if len(t) != 4:
            raise CdrBadArgument("transformations 需要 4 个数字 [d11,d12,d21,d22]")
        fountain.SetTransformations(float(t[0]), float(t[1]), float(t[2]), float(t[3]))
    if spec.get("opaque"):
        fountain.MakeOpaque()


EFFECT_KINDS = ("bci", "gamma", "hsl", "invert", "posterize", "color_balance")


def _apply_effect(shape: Any, spec: Any) -> str:
    """对象/位图色彩效果（实测可用）。"""
    if not isinstance(spec, dict) or "type" not in spec:
        raise CdrBadArgument(f"effect 需为 {{type:...}}；可用 {', '.join(EFFECT_KINDS)}")
    kind = str(spec["type"]).lower().replace("-", "_")
    if kind == "bci":
        shape.ApplyEffectBCI(float(spec.get("brightness", 0)),
                             float(spec.get("contrast", 0)),
                             float(spec.get("intensity", 0)))
    elif kind == "gamma":
        shape.ApplyEffectGamma(float(spec.get("gamma", 1.0)))
    elif kind == "hsl":
        shape.ApplyEffectHSL(float(spec.get("hue", 0)),
                             float(spec.get("saturation", 0)),
                             float(spec.get("lightness", 0)))
    elif kind == "invert":
        shape.ApplyEffectInvert()
    elif kind == "posterize":
        shape.ApplyEffectPosterize(int(spec.get("level", 4)))
    elif kind == "color_balance":
        shape.ApplyEffectColorBalance(
            float(spec.get("cyan_red", 0)), float(spec.get("magenta_green", 0)),
            float(spec.get("yellow_blue", 0)),
            bool(spec.get("shadows", True)), bool(spec.get("midtones", True)),
            bool(spec.get("highlights", True)), bool(spec.get("preserve_luminance", True)),
        )
    else:
        raise CdrBadArgument(f"未知 effect {kind}；可用 {', '.join(EFFECT_KINDS)}")
    return kind


COLOR_OBJECT_ACTIONS = ("object_info", "object_name", "object_compare", "object_blend",
                        "object_copy", "object_string_assign", "object_convert",
                        "object_fixed", "object_assign",
                        # 第三十五波：其它色彩模型工厂 + 颜色串工厂
                        "object_model", "object_from_string",
                        # 第四十八波：固定调色板/脚本/调色板取色/失效
                        "object_palette", "object_invalidate",
                        # 第四十九波：分量读取（CorelScriptGetComponent）/ 转指定调色板
                        "object_components", "object_convert_palette",
                        # 第五十一波：颜色风格对象改写
                        "object_modify_style_color",
                        # 第五十八波：调色板工厂 / 专色工厂
                        "object_palette_factory", "object_from_spot")
# 显式写法（而不是 getattr 动态调用）：既方便 grep，也让覆盖率审计能统计到
COLOR_CONVERT_CALLS = {
    "rgb": lambda c: c.ConvertToRGB(),
    "cmyk": lambda c: c.ConvertToCMYK(),
    "gray": lambda c: c.ConvertToGray(),
    "hls": lambda c: c.ConvertToHLS(),
    "hsb": lambda c: c.ConvertToHSB(),
    "lab": lambda c: c.ConvertToLab(),
    "cmy": lambda c: c.ConvertToCMY(),
    "yiq": lambda c: c.ConvertToYIQ(),
}
COLOR_COMPONENT_ARITY = {"rgb": 3, "cmyk": 4, "gray": 1, "hls": 3, "hsb": 3,
                         "lab": 3, "cmy": 3, "yiq": 3, "bw": 1, "registration": 0,
                         "fixed": 3, "script": 5}
COLOR_ASSIGN_CALLS = {
    "rgb": lambda c, v: c.RGBAssign(int(v[0]), int(v[1]), int(v[2])),
    "cmyk": lambda c, v: c.CMYKAssign(int(v[0]), int(v[1]), int(v[2]), int(v[3])),
    "gray": lambda c, v: c.GrayAssign(int(v[0])),
    "hls": lambda c, v: c.HLSAssign(int(v[0]), int(v[1]), int(v[2])),
    "hsb": lambda c, v: c.HSBAssign(int(v[0]), int(v[1]), int(v[2])),
    "lab": lambda c, v: c.LabAssign(float(v[0]), float(v[1]), float(v[2])),
    "cmy": lambda c, v: c.CMYAssign(int(v[0]), int(v[1]), int(v[2])),
    "yiq": lambda c, v: c.YIQAssign(int(v[0]), int(v[1]), int(v[2])),
    "bw": lambda c, v: c.BWAssign(int(v[0])),
    "registration": lambda c, v: c.RegistrationAssign(),
    # 第四十八波：固定调色板（palette_id, index, tint）与脚本式分量
    "fixed": lambda c, v: c.FixedAssign(int(v[0]), int(v[1]), int(v[2])),
    "script": lambda c, v: c.CorelScriptAssign(int(v[0]), int(v[1]), int(v[2]),
                                               int(v[3]), int(v[4]), 0, 0, 0),
}


def _color_object(spec: Any) -> Any:
    """按 spec 造一个 IVGColor（走桥接持有的 Application，与文档无关）。"""
    if spec is None:
        raise CdrBadArgument("需要 color")
    return _color(None, spec)


def _color_object_info(spec: Any) -> dict:
    obj = _color_object(spec)
    return {
        "input": spec if isinstance(spec, (str, int, float)) else "dict",
        "type": _n(lambda: int(obj.Type)),
        "to_string": _s(lambda: str(obj.ToString())),
        "name": _s(lambda: str(obj.GetName())),
    }


def color_object_ops(action: str, *, color: Any = None, color2: Any = None,
                     ratio: float = 50.0, name: str | None = None,
                     model: str | None = None, mode: int = 0,
                     values: list | None = None) -> dict:
    """X8 `IVGColor` 对象级操作（实测：分量属性不可读，但下面这些都能用）。

    * `IsSame` / `IsSimilar` / `GetColorDistanceFrom` 可比较两种颜色
    * `BlendWith(other, 0..100)` **原地**改色，返回 None
    * `ConvertToRGB/CMYK/Gray/HLS/HSB/Lab/CMY/YIQ` **原地**转换（返回 None，类型会变）
    * `ConvertToFixed(0/1/2)` 转固定调色板（实测 0=USER、1=TRU、2=PAN_PROC）
    * `GetName/SetName` 读写颜色名；`StringAssign("CMYK,USER,…")` 从颜色串赋值
    * `ConvertToPalette` / `SpotAssignByName` / `CopyAppearance` 实测报错，未暴露
    """
    act = (action or "").lower()
    if act not in COLOR_OBJECT_ACTIONS:
        raise CdrBadArgument(f"未知 action {action!r}；可用：{', '.join(COLOR_OBJECT_ACTIONS)}")

    if act == "object_info":
        return {"action": act, **_color_object_info(color)}

    if act == "object_name":
        obj = _color_object(color)
        before = _s(lambda: str(obj.GetName()))
        if name:
            obj.SetName(str(name))
        return {"action": act, "name_before": before,
                "name_after": _s(lambda: str(obj.GetName())),
                "to_string": _s(lambda: str(obj.ToString()))}

    if act == "object_compare":
        if color2 is None:
            raise CdrBadArgument("object_compare 需要 color2")
        a = _color_object(color)
        b = _color_object(color2)
        distance = _s(lambda: float(a.GetColorDistanceFrom(b)))
        return {
            "action": act,
            "a": _s(lambda: str(a.ToString())),
            "b": _s(lambda: str(b.ToString())),
            "is_same": bool(_s(lambda: bool(a.IsSame(b)), False)),
            "is_similar": bool(_s(lambda: bool(a.IsSimilar(b)), False)),
            "distance": _round(distance) if distance is not None else None,
        }

    if act == "object_blend":
        if color2 is None:
            raise CdrBadArgument("object_blend 需要 color2")
        a = _color_object(color)
        b = _color_object(color2)
        before = _s(lambda: str(a.ToString()))
        a.BlendWith(b, float(ratio))
        return {"action": act, "ratio": float(ratio), "before": before,
                "after": _s(lambda: str(a.ToString())),
                "type_after": _n(lambda: int(a.Type))}

    if act == "object_copy":
        a = _color_object(color)
        clone = _s(lambda: a.GetCopy())
        if clone is None:
            raise CdrComFailure("Color.GetCopy 返回空")
        return {"action": act, "original": _s(lambda: str(a.ToString())),
                "copy": _s(lambda: str(clone.ToString())),
                "is_same": bool(_s(lambda: bool(clone.IsSame(a)), False))}

    if act == "object_string_assign":
        a = _color_object(color)
        if not isinstance(color2, str) or not color2:
            raise CdrBadArgument('object_string_assign 需要 color2 = 颜色串，'
                                 '如 "CMYK,USER,10,20,30,40,100,<GUID>"')
        before = _s(lambda: str(a.ToString()))
        a.StringAssign(color2)
        return {"action": act, "before": before,
                "after": _s(lambda: str(a.ToString())),
                "type_after": _n(lambda: int(a.Type))}

    if act == "object_convert":
        want = str(model or color2 or "cmyk").lower()
        if want not in COLOR_CONVERT_CALLS:
            raise CdrBadArgument(
                f"不支持的目标模型 {want!r}；可用 {sorted(COLOR_CONVERT_CALLS)}")
        a = _color_object(color)
        before = _s(lambda: str(a.ToString()))
        COLOR_CONVERT_CALLS[want](a)
        return {"action": act, "model": want, "before": before,
                "after": _s(lambda: str(a.ToString())),
                "type_after": _n(lambda: int(a.Type))}

    if act == "object_model":
        # 实测各模型工厂：CreateHSBColor(H,S,B) / CreateCMYColor(C,M,Y) /
        # CreateBWColor(White) / CreateYIQColor(Y,I,Q) / CreateFixedColor(palette, index, tint)
        want = str(model or "hsb").lower()
        if want not in ("hsb", "cmy", "bw", "yiq", "fixed", "ex"):
            raise CdrBadArgument("object_model 的 model 可用 hsb/cmy/bw/yiq/fixed/ex")
        vals = [float(v) for v in (values or [])]
        app = _app_of(None)
        if want == "hsb":
            obj = app.CreateHSBColor(int(vals[0]), int(vals[1]), int(vals[2]))
        elif want == "cmy":
            obj = app.CreateCMYColor(int(vals[0]), int(vals[1]), int(vals[2]))
        elif want == "bw":
            obj = app.CreateBWColor(bool(vals[0]) if vals else True)
        elif want == "yiq":
            obj = app.CreateYIQColor(int(vals[0]), int(vals[1]), int(vals[2]))
        elif want == "fixed":
            obj = app.CreateFixedColor(int(vals[0]), int(vals[1]), int(vals[2]))
        elif want == "ex":
            # CreateColorEx(ColorModel, V1..V7)：模型号同 ToString 前缀（2=CMYK, 5=RGB…）
            model_id = int(vals[0])
            rest = [int(v) for v in vals[1:8]] + [0] * (7 - len(vals[1:8]))
            obj = app.CreateColorEx(model_id, *rest)
        else:
            raise CdrBadArgument("object_model 的 model 可用 hsb/cmy/bw/yiq/fixed/ex")
        if obj is None:
            raise CdrComFailure(f"Create{want}Color 返回空")
        return {"action": act, "model": want, "values": vals,
                "type": _n(lambda: int(obj.Type)),
                "to_string": _s(lambda: str(obj.ToString()))}

    if act == "object_from_string":
        if not isinstance(color2, str) or not color2:
            raise CdrBadArgument('object_from_string 需要 color2 = 颜色串')
        obj = _s(lambda: _app_of(None).CreateColor(color2))
        if obj is None:
            raise CdrComFailure("Application.CreateColor 返回空")
        return {"action": act, "input": color2,
                "type": _n(lambda: int(obj.Type)),
                "to_string": _s(lambda: str(obj.ToString())),
                "name": _s(lambda: str(obj.GetName()))}

    if act == "object_convert_palette":
        # 实测：Color.ConvertToPalette(调色板 **GUID**)：RGB(120,60,30) + TRUMATCH
        # → CMYK255,TRU,166,230,255,15,<guid>；传数字序号会报「找不到调色板 N」
        if not isinstance(color2, str) or not color2:
            raise CdrBadArgument("object_convert_palette 需要 color2 = 调色板 Identifier（GUID）")
        obj = _color_object(color if color is not None else "#000000")
        before = _s(lambda: str(obj.ToString()))
        obj.ConvertToPalette(color2)
        return {"action": act, "palette": color2, "before": before,
                "after": _s(lambda: str(obj.ToString())),
                "type_after": _n(lambda: int(obj.Type))}

    if act in {"object_palette_factory", "object_from_spot"}:
        # 实测：Application.CreatePaletteColor(调色板 GUID, 序号) 可用（TRUMATCH → CMYK255,TRU,…）；
        # CreateSpotColor(调色板 GUID, 专色ID, Tint) 需要**含专色**的调色板，否则报「未找到指定的对象」
        app = _app_of(None)
        if not isinstance(color2, str) or not color2:
            raise CdrBadArgument(f"{act} 需要 color2 = 调色板 Identifier（GUID）")
        vals = list(values or [])
        if act == "object_palette_factory":
            idx = int(vals[0]) if vals else 1
            obj = app.CreatePaletteColor(color2, idx)
            if obj is None:
                raise CdrComFailure("CreatePaletteColor 返回空")
            return {"action": act, "palette": color2, "color_index": idx,
                    "to_string": _s(lambda: str(obj.ToString())),
                    "type": _n(lambda: int(obj.Type))}
        spot_id = int(vals[0]) if vals else 1
        tint = int(vals[1]) if len(vals) > 1 else 100
        obj = app.CreateSpotColor(color2, spot_id, tint)
        if obj is None:
            raise CdrComFailure("CreateSpotColor 返回空")
        return {"action": act, "palette": color2, "spot_id": spot_id, "tint": tint,
                "to_string": _s(lambda: str(obj.ToString())),
                "type": _n(lambda: int(obj.Type))}

    if act == "object_modify_style_color":
        # 实测：Color.ModifyColorStyleColor(目标颜色, ChangeWholeHarmony) 可调用
        target = color2 if color2 is not None else "#FFFFFF"
        src = _color_object(color if color is not None else "#000000")
        dst = _color_object(target)
        src.ModifyColorStyleColor(dst, False)
        return {"action": act, "source": _s(lambda: str(src.ToString())),
                "target": _s(lambda: str(dst.ToString())),
                "note": "把颜色改为目标颜色的风格色表示（ChangeWholeHarmony=False）"}

    if act == "object_components":
        # 实测：Color.CorelScriptGetComponent(模型号, 0×7) 返回
        # (状态码, v1..v7)：CMYK(2) → (5002,10,20,30,40,0,0,0)；
        # RGB(5) → (5005,11,22,33,0,0,0,0)。这是**唯一能读到颜色分量**的路径。
        want = int(model) if model is not None else 2
        obj = _color_object(color if color is not None else "#000000")
        raw = _s(lambda: obj.CorelScriptGetComponent(want, 0, 0, 0, 0, 0, 0, 0))
        if raw is None:
            raise CdrComFailure("CorelScriptGetComponent 返回空")
        vals = list(raw) if isinstance(raw, (tuple, list)) else [raw]
        return {"action": act, "model": want,
                "raw": [(_round(float(v)) if isinstance(v, (int, float)) else str(v))
                        for v in vals],
                "values": [(_round(float(v)) if isinstance(v, (int, float)) else str(v))
                           for v in vals[1:]],
                "status": int(vals[0]) if vals else None,
                "to_string": _s(lambda: str(obj.ToString())),
                "note": "返回 (状态码, v1..v7)；状态码 = 5000 + 模型号"}

    if act == "object_palette":
        # 实测：PaletteAssign(调色板 Identifier, 颜色序号) /
        # SpotAssignByName(调色板 Identifier, 专色名, Tint) —— 第一个参数是**调色板 GUID 字符串**
        token = color2
        if token is None:
            raise CdrBadArgument("object_palette 需要 color2 = 调色板 Identifier 或已打开调色板的序号")
        app = _app_of(None)
        pal_index = None
        if isinstance(token, int) or (isinstance(token, str) and token.isdigit()):
            idx = int(token)
            count = _n(lambda: int(app.Palettes.Count), 0) or 0
            if not 1 <= idx <= count:
                raise CdrBadArgument(f"调色板序号越界：{idx}（已打开 {count} 个）")
            pal = app.Palettes.Item(idx)
            ident = _s(lambda: str(pal.Identifier))
        else:
            ident = str(token)
        if not ident:
            raise CdrComFailure("取不到调色板 Identifier")
        obj = _color_object(color if color is not None else "#000000")
        vals = list(values or [])
        if name:
            tint = int(vals[0]) if vals else 100
            obj.SpotAssignByName(ident, str(name), tint)
            return {"action": act, "mode": "spot", "palette": ident,
                    "spot_name": str(name), "tint": tint,
                    "to_string": _s(lambda: str(obj.ToString()))}
        color_index = int(vals[0]) if vals else 1
        obj.PaletteAssign(ident, color_index)
        return {"action": act, "mode": "palette", "palette": ident,
                "color_index": color_index,
                "to_string": _s(lambda: str(obj.ToString()))}

    if act == "object_invalidate":
        obj = _color_object(color if color is not None else "#000000")
        obj.Invalidate()
        return {"action": act, "invalidated": True,
                "to_string": _s(lambda: str(obj.ToString()))}

    if act == "object_assign":
        want = str(model or "rgb").lower()
        if want not in COLOR_ASSIGN_CALLS:
            raise CdrBadArgument(
                f"不支持的分量模型 {want!r}；可用 {sorted(COLOR_ASSIGN_CALLS)}")
        need = COLOR_COMPONENT_ARITY[want]
        vals = list(values or [])
        if need and len(vals) != need:
            raise CdrBadArgument(
                f"{want} 需要 {need} 个分量，收到 {len(vals)} 个（用 values=[…] 传）")
        a = _color_object(color if color is not None else "#000000")
        before = _s(lambda: str(a.ToString()))
        COLOR_ASSIGN_CALLS[want](a, vals)
        return {"action": act, "model": want, "values": vals, "before": before,
                "after": _s(lambda: str(a.ToString())),
                "type_after": _n(lambda: int(a.Type))}

    # object_fixed
    a = _color_object(color)
    before = _s(lambda: str(a.ToString()))
    a.ConvertToFixed(int(mode))
    return {"action": act, "mode": int(mode), "before": before,
            "after": _s(lambda: str(a.ToString())),
            "type_after": _n(lambda: int(a.Type))}


def color_ops(doc: Any, action: str, *, color: Any = None, color2: Any = None,
              ratio: float = 0.5, index: int = 1,
              model: str | None = None, mode: int = 0, name: str | None = None,
              values: list | None = None, label: str = "cdr_color") -> dict:
    """颜色运算（**纯 Python 计算，不依赖 CorelDRAW**）。

    实测原因：X8 的 IVGColor 不暴露任何分量属性（Red/Cyan/Gray 全部 AttributeError），
    ConvertToCMYK 只改类型不返回值，所以运算在本地做。

    action:
      info     解析颜色，返回 HEX / RGB / CMYK / Gray / HLS / Lab
      convert  同上，但把目标空间单独拎出来（color2 = cmyk/rgb/gray/hls/lab）
      blend    两个颜色按 ratio 混合（0=全 color，1=全 color2）
    """
    from . import colors as colorlib

    act = (action or "info").lower()
    if act in COLOR_OBJECT_ACTIONS:
        # ratio 对 COM 侧是 0..100；调用方习惯 0..1，这里统一归一
        pct = float(ratio) * 100.0 if float(ratio) <= 1.0 else float(ratio)
        return color_object_ops(act, color=color, color2=color2, ratio=pct,
                                model=model, mode=mode, name=name, values=values)
    if act not in {"info", "convert", "blend"}:
        raise CdrBadArgument(
            "action 可用 info / convert / blend / " + " / ".join(COLOR_OBJECT_ACTIONS))
    if color is None:
        raise CdrBadArgument("需要 color")

    try:
        if act == "blend":
            if color2 is None:
                raise CdrBadArgument("blend 需要 color2")
            a = colorlib.describe(color)["rgb"]
            b = colorlib.describe(color2)["rgb"]
            mixed = colorlib.blend(a, b, ratio)
            out = colorlib.describe({"r": mixed[0], "g": mixed[1], "b": mixed[2]})
            return {"action": act, "ratio": ratio, **out}
        base = colorlib.describe(color)
        if act == "convert":
            target = str(color2 or "cmyk").lower()
            if target not in base:
                raise CdrBadArgument(
                    f"不支持的目标空间 {target}；可用 {sorted(base)}")
            return {"action": act, "target": target, "value": base[target], **base}
        return {"action": act, **base}
    except ValueError as exc:
        raise CdrBadArgument(str(exc)) from exc


def _hex_to_rgb(text: str) -> tuple[int, int, int]:
    s = str(text).strip().lstrip("#")
    if len(s) != 6:
        raise CdrBadArgument(f"颜色格式不对：{text}")
    return int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)


def _rgb_to_cmyk(r: int, g: int, b: int) -> tuple[int, int, int, int]:
    if (r, g, b) == (0, 0, 0):
        return 0, 0, 0, 100
    k = 1 - max(r, g, b) / 255
    c = (1 - r / 255 - k) / (1 - k)
    m = (1 - g / 255 - k) / (1 - k)
    y = (1 - b / 255 - k) / (1 - k)
    return round(c * 100), round(m * 100), round(y * 100), round(k * 100)


# --------------------------------------------------------------------------- #
# 形状：修改
# --------------------------------------------------------------------------- #
def modify_shape(doc: Any, page_index: int | None, target: dict, changes: dict,
                 label: str = "cdr_modify") -> dict:
    with mm_units(doc):
        page = _page(doc, page_index)
        shape = find_shape(doc, page, target)
        with atomic(doc, label):
            return _apply_changes(doc, shape, changes)


def _apply_changes(doc: Any, shape: Any, changes: dict) -> dict:
    applied: list[str] = []
    warnings: list[str] = []
    forced_name: str | None = None
    layer_out: str | None = None

    if "dx" in changes or "dy" in changes:
        shape.Move(float(changes.get("dx", 0)), float(changes.get("dy", 0)))
        applied.append("move")

    if "x" in changes or "y" in changes:
        box = shape.GetBoundingBox()
        shape.SetPosition(float(changes.get("x", box[0])), float(changes.get("y", box[1])))
        applied.append("position")

    if "width" in changes or "height" in changes:
        box = shape.GetBoundingBox()
        shape.SetBoundingBox(float(box[0]), float(box[1]),
                             float(changes.get("width", box[2])),
                             float(changes.get("height", box[3])),
                             bool(changes.get("keep_aspect", False)), 0)
        applied.append("size")

    if "rotate" in changes:
        shape.Rotate(float(changes["rotate"]))
        applied.append("rotate")

    if "fill" in changes:
        try:
            applied.append("fill:" + apply_fill(doc, shape, changes["fill"]))
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"填充失败：{exc}")

    if isinstance(changes.get("pattern_fill"), dict) and changes["pattern_fill"]:
        try:
            pf = _s(lambda: shape.Fill.Pattern)
            if pf is None:
                raise CdrBadArgument("对象当前不是图案填充")
            spec = changes["pattern_fill"]
            mapping = {
                "rotation_angle": "RotationAngle", "skew_angle": "SkewAngle",
                "origin_x": "OriginX", "origin_y": "OriginY",
                "tile_width": "TileWidth", "tile_height": "TileHeight",
                "tile_offset": "TileOffset", "tile_offset_type": "TileOffsetType",
                "mirror_fill": "MirrorFill", "mirror_fill_x": "MirrorFillX",
                "mirror_fill_y": "MirrorFillY",
                "transform_with_shape": "TransformWithShape",
                "canvas_index": "Canvas", "pattern_type": "Type",
            }
            for key, attr in mapping.items():
                if key in spec and spec[key] is not None:
                    setattr(pf, attr, spec[key])
            for key, attr in (("front_color", "FrontColor"),
                              ("back_color", "BackColor")):
                if spec.get(key) is not None:
                    setattr(pf, attr, _color(doc, spec[key]))
            applied.append("pattern_fill")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"图案填充微调失败：{exc}")

    if "outline" in changes:
        try:
            apply_outline(doc, shape, changes["outline"])
            applied.append("outline")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"描边失败：{exc}")

    if "transparency" in changes:
        try:
            applied.append("transparency:" + _apply_transparency(
                doc, shape, changes["transparency"]))
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"透明度失败：{exc}")

    if changes.get("freeze_transparency") is not None:
        try:
            if bool(changes["freeze_transparency"]):
                page = doc.ActivePage
                before = set()
                for _i, _l, _j, s in _iter_page_shapes(page):
                    before.add(str(_s(lambda s=s: s.Name, "")))
                shape.Transparency.Freeze()
                # 实测：Freeze 会把对象替换成**无名组合体**（type 7，Name=''），
                # 原引用失效且改名无效；这里找到新对象并补名，保证后续能按名定位。
                frozen = None
                for _i, _l, _j, s in _iter_page_shapes(page):
                    nm = str(_s(lambda s=s: s.Name, ""))
                    if not nm or nm not in before:
                        frozen = s
                if frozen is not None:
                    forced_name = f"{NAME_PREFIX}{_short_id()}"
                    frozen.Name = forced_name
                applied.append("freeze_transparency")
            else:
                shape.Transparency.Unfreeze()
                applied.append("unfreeze_transparency")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"冻结/解冻透明度失败：{exc}")

    if "gradient" in changes:
        try:
            _apply_gradient_tweak(shape, changes["gradient"])
            applied.append("gradient")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"渐变调整失败（需要对象已有渐变填充）：{exc}")

    if "effect" in changes:
        try:
            applied.append("effect:" + _apply_effect(shape, changes["effect"]))
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"效果失败：{exc}")

    if "outline_from" in changes:
        try:
            src = find_shape(doc, doc.ActivePage, changes["outline_from"])
            shape.Outline.CopyAssign(src.Outline)
            applied.append("outline_from")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"复制描边失败：{exc}")

    if "fill_from" in changes:
        try:
            src = find_shape(doc, doc.ActivePage, changes["fill_from"])
            shape.Fill.CopyAssign(src.Fill)
            applied.append("fill_from")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"复制填充失败：{exc}")

    if changes.get("affine"):
        try:
            m = list(changes["affine"])
            if len(m) != 6:
                raise CdrBadArgument("affine 需要 6 个数字 [d11,d12,d21,d22,centerX,centerY]")
            shape.AffineTransform(*[float(v) for v in m])
            applied.append("affine")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"仿射变换失败：{exc}")

    if changes.get("align_to"):
        try:
            spec = changes["align_to"] if isinstance(changes["align_to"], dict) else {}
            mode = str(spec.get("mode", "page_center")).lower()
            typ = int(spec.get("type", 1))
            if mode == "page_center":
                shape.AlignToPageCenter(typ, 0)
            elif mode == "point":
                shape.AlignToPoint(typ, float(spec.get("x", 0)), float(spec.get("y", 0)), 0)
            elif mode == "grid":
                shape.AlignToGrid(typ, 0)
            else:
                raise CdrBadArgument("align_to.mode 可用 page_center / point / grid")
            applied.append("align_to")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"对齐失败：{exc}")

    if changes.get("flip") is not None:
        try:
            shape.Flip(int(changes["flip"]))
            applied.append("flip")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"翻转失败：{exc}")

    if changes.get("skew"):
        try:
            s = changes["skew"]
            if isinstance(s, dict):
                shape.Skew(float(s.get("x", 0)), float(s.get("y", 0)))
            else:
                shape.Skew(float(s), 0.0)
            applied.append("skew")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"斜切失败：{exc}")

    if "locked" in changes:
        try:
            shape.Locked = bool(changes["locked"])
            applied.append("locked")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"锁定设置失败：{exc}")

    if "matrix" in changes:
        try:
            m = list(changes["matrix"])
            if len(m) == 6:
                shape.TransformMatrix(float(m[0]), float(m[1]), float(m[2]),
                                      float(m[3]), float(m[4]), float(m[5]))
                applied.append("matrix")
            else:
                warnings.append("matrix 需要 6 个数字 [d11,d12,d21,d22,tx,ty]")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"变换矩阵失败：{exc}")

    if changes.get("clone"):
        try:
            c = changes["clone"] if isinstance(changes["clone"], dict) else {}
            new_shape = shape.Clone(float(c.get("dx", 5)), float(c.get("dy", -5)))
            applied.append("clone")
            if new_shape is not None:
                # 克隆体的 Name 与源对象**完全相同**，必须强制改名，否则按名字找不到
                new_name = f"{NAME_PREFIX}{_short_id()}"
                try:
                    new_shape.Name = new_name
                except Exception:
                    new_name = _ensure_name(new_shape)
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"克隆失败：{exc}")

    if changes.get("align_to_page"):
        try:
            shape.AlignToPage(int(changes["align_to_page"]), 0)
            applied.append("align_to_page")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"对齐到页面失败：{exc}")

    if changes.get("align_to_shape"):
        try:
            spec = changes["align_to_shape"]
            other = find_shape(doc, doc.ActivePage, spec.get("target", spec))
            shape.AlignToShape(int(spec.get("type", 1)), other, 0)
            applied.append("align_to_shape")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"对齐到对象失败：{exc}")

    if "drop_shadow" in changes and changes["drop_shadow"]:
        s = changes["drop_shadow"] if isinstance(changes["drop_shadow"], dict) else {}
        try:
            shape.CreateDropShadow(
                int(s.get("type", 1)), float(s.get("opacity", 50)), float(s.get("feather", 0)),
                float(s.get("offset_x", 2)), float(s.get("offset_y", -2)),
                _color(doc, s.get("color", "#000000")),
            )
            applied.append("drop_shadow")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"阴影失败：{exc}")

    if "contour" in changes and changes["contour"]:
        c = changes["contour"] if isinstance(changes["contour"], dict) else {}
        try:
            shape.CreateContour(int(c.get("direction", 0)), float(c.get("offset", 2)),
                                int(c.get("steps", 3)), int(c.get("blend_type", 1)),
                                None, None, None)
            applied.append("contour")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"轮廓图失败：{exc}")

    if "envelope" in changes:
        try:
            shape.CreateEnvelope(int(changes["envelope"]), 0, True)
            applied.append("envelope")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"封套失败：{exc}")

    if "text" in changes:
        try:
            shape.Text.Story.Text = str(changes["text"])
            applied.append("text")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"改文字失败（该对象可能不是文字）：{exc}")

    if changes.get("fit_to_path"):
        try:
            path_shape = find_shape(doc, doc.ActivePage, changes["fit_to_path"])
            shape.Text.FitToPath(path_shape)
            applied.append("fit_to_path")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"文字沿路径失败：{exc}")

    if changes.get("fit_frame"):
        try:
            shape.Text.FitTextToFrame()
            applied.append("fit_frame")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"文本适应文本框失败：{exc}")

    tri = {"bold", "italic", "underline"}
    for key, prop in (("font", "Font"), ("size", "Size"), ("bold", "Bold"),
                      ("italic", "Italic"), ("underline", "Underline"),
                      ("alignment", None)):
        if key in changes and changes[key] is not None:
            try:
                if key == "alignment":
                    _set_text_align(shape, changes[key])
                else:
                    setattr(shape.Text.Story, prop,
                            _underline(changes[key]) if key == "underline"
                            else (_tristate(changes[key]) if key in tri else changes[key]))
                applied.append(key)
            except Exception as exc:  # noqa: BLE001
                warnings.append(f"{key} 设置失败：{exc}")

    if "name" in changes:
        shape.Name = str(changes["name"])
        applied.append("name")

    if changes.get("to_front"):
        shape.OrderToFront()
        applied.append("to_front")
    if changes.get("to_back"):
        shape.OrderToBack()
        applied.append("to_back")

    if changes.get("order"):
        try:
            spec = changes["order"]
            mode = str(spec.get("mode", spec) if isinstance(spec, dict) else spec).lower()
            ref_spec = spec.get("ref") if isinstance(spec, dict) else changes.get("order_ref")
            if mode in {"front_of", "back_of"}:
                if not ref_spec:
                    raise CdrBadArgument("order=front_of/back_of 需要 ref 指定参照对象")
                ref = find_shape(doc, doc.ActivePage, ref_spec)
                if mode == "front_of":
                    shape.OrderFrontOf(ref)
                else:
                    shape.OrderBackOf(ref)
            elif mode in {"forward_one", "forward"}:
                shape.OrderForwardOne()
            elif mode in {"back_one", "backward"}:
                shape.OrderBackOne()
            elif mode == "reverse":
                shape.OrderReverse()
            else:
                raise CdrBadArgument(
                    "order 可用 front_of / back_of / forward_one / back_one / reverse")
            applied.append(f"order:{mode}")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"层级调整失败：{exc}")

    if changes.get("clear_transformations"):
        try:
            shape.ClearTransformations()
            applied.append("clear_transformations")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"清除变换失败：{exc}")

    if isinstance(changes.get("rotation_center"), dict):
        try:
            rc = changes["rotation_center"]
            shape.SetRotationCenter(float(rc.get("x", 0)), float(rc.get("y", 0)))
            applied.append("rotation_center")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"设置旋转中心失败：{exc}")

    if any(k in changes for k in ("overprint_fill", "overprint_outline",
                                  "overprint_bitmap")):
        try:
            if "overprint_fill" in changes:
                shape.OverprintFill = bool(changes["overprint_fill"])
            if "overprint_outline" in changes:
                shape.OverprintOutline = bool(changes["overprint_outline"])
            if "overprint_bitmap" in changes:
                shape.OverprintBitmap = bool(changes["overprint_bitmap"])
            applied.append("overprint")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"叠印设置失败：{exc}")

    if changes.get("copy_clipboard"):
        try:
            shape.Copy()
            applied.append("copy_clipboard")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"复制到剪贴板失败：{exc}")

    if isinstance(changes.get("data"), dict) and changes["data"]:
        try:
            data_format = str(changes.get("data_format", "0"))
            fields = doc.DataFields
            obj_data = shape.ObjectData
            for field_name, val in changes["data"].items():
                name = str(field_name)
                field = None
                if _s(lambda: bool(fields.IsPresent(name)), False):
                    field = _s(lambda: fields.Item(name))
                if field is None:
                    field = fields.Add(name, data_format, False, False, False)
                existing = _s(lambda: obj_data.Item(name))
                if existing is not None:
                    existing.Value = str(val)
                else:
                    obj_data.Add(field, str(val))
                applied.append(f"data:{name}")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"对象数据字段设置失败：{exc}")

    if changes.get("cut_clipboard"):
        try:
            shape.Cut()
            applied.append("cut_clipboard")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"剪切到剪贴板失败：{exc}")

    if changes.get("move_to_layer") is not None or changes.get("copy_to_layer") is not None:
        try:
            page = doc.ActivePage
            target_name = (changes.get("move_to_layer")
                           if changes.get("move_to_layer") is not None
                           else changes.get("copy_to_layer"))
            layer = _find_layer(page, name=str(target_name), index=None)
            if changes.get("move_to_layer") is not None:
                shape.MoveToLayer(layer)
                applied.append("move_to_layer")
            else:
                new_shape = shape.CopyToLayer(layer)
                if new_shape is not None:
                    forced_name = _ensure_name(new_shape)
                applied.append("copy_to_layer")
            layer_out = _s(lambda: str(layer.Name))
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"移动到/复制到图层失败：{exc}")

    new_name = None
    if changes.get("duplicate"):
        dup = shape.Duplicate(float(changes.get("dup_dx", 5)), float(changes.get("dup_dy", -5)))
        new_name = _s(lambda: str(dup.Name))
        applied.append("duplicate")

    if changes.get("delete"):
        shape.Delete()
        applied.append("delete")

    result = {
        "applied": applied,
        "name": forced_name or new_name or _s(lambda: str(shape.Name)),
        "bbox": None if changes.get("delete") or changes.get("cut_clipboard") else _bbox(shape),
        "warnings": warnings,
    }
    if layer_out is not None:
        result["layer"] = layer_out
    return result


def _set_text_align(shape: Any, value: Any) -> None:
    """对齐：X8 上走 Text.GetAlignProperties / SetAlignProperties。

    实测这条链在部分文字对象上会抛 COM 错误，所以调用方会用 try/except 兜住，
    失败时降级成 warning，不让整个 cdr_text 挂掉。
    """
    props = shape.Text.GetAlignProperties(0)
    props.Alignment = _align(value)
    shape.Text.SetAlignProperties(0, props)


# --------------------------------------------------------------------------- #
# 形状：组合 / 布尔 / 整形
# --------------------------------------------------------------------------- #
SHAPE_OPS = (
    "group", "ungroup", "weld", "trim", "intersect", "combine", "break_apart",
    "fillet", "chamfer", "scallop", "to_curves", "powerclip", "to_symbol",
    "outline_to_curves",
    # 第五波：排列 / 阵列 / 边界 / 立体透视
    "step_repeat", "distribute", "boundary", "perspective", "separate",
    # 第六波：范围级（选择集）
    "parallel_curves", "rasterize",
    # 第七波：包络 / 比较 / 范围批量
    "envelope_from_shape", "compare",
    "batch_outline_to_curves", "fill_mode", "pixel_aligned", "range_boundary",
    # 第九波：投影平面
    "project", "unproject",
    # 第十波：把对象/选区提取成新文档
    "to_document",
    # 第十五波：立体化 / 透镜 / 效果管理
    "extrude", "lens", "extrude_set", "lens_set", "effect_info", "clear_effects",
    # 第十七波：混合
    "blend", "blend_set",
    # 第三十三波：范围级批量（解锁/全解组/炸开/清效果/范围查询/范围合并）
    "unlock_all", "ungroup_all", "ungroup_all_ex", "ungroup_ex", "break_apart_ex",
    "clear_effect", "range_info",
    "range_add", "range_remove",
    # 第三十八波：用 Document.CreateUniformFill + ShapeRange.ApplyFill 上色
    "apply_fill",
    # 第三十九波：形状矩阵 / 扭曲 / 文字内嵌 / 范围复制 / 命中与层序
    "matrix", "is_type", "in_front_of", "is_on_shape", "distort", "place_text",
    "duplicate_range", "clone_range",
    # 第四十波：参考点定位 / 吸附点反查 / 链接对象 / 位图化 / 叠印状态
    "position_ex", "find_snap", "linked", "rasterize_ex", "overprint_state",
    # 第四十二波：用对象列表直接建 ShapeRange 并整批操作
    "range_from_list", "range_bitmap",
    # 第四十七波：整批描边设置（SetOutlineProperties / Ex）
    "range_outline",
    # 第四十八波：复制到其它图层（AsRange）/ 对齐到形状范围
    "copy_to_layer_range", "align_to_range",
    # 第五十三波：把某个对象的描边整批套到范围
    "range_outline_apply",
    # 第五十四波：同层相邻对象遍历（GetNext / GetPrevious）
    "sibling",
    # 第六十一波：按中心定尺寸/拉伸（SetSizeEx / StretchEx）与范围排除
    "set_size_ex", "stretch_ex", "all_excluding",
    # 第六十四波：标注对象 / 图框内容 / 辅助线几何
    "dimension_info", "dimension_set", "powerclip_extract", "guide_info", "guide_set",
    # 第六十五波：混合效果对象（IVGEffectBlend）
    "blend_effect",
    # 第六十七波：矩形圆角 / 克隆链 / 投影偏移 / 效果子对象
    "corners", "clone_links", "drop_shadow",
    # 第七十一波：剖面线参数 / 渐变控制点 / 椭圆几何
    "hatch_fill", "fountain_colors", "ellipse_geometry",
    # 第七十二波：包络效果对象（IVGEffectEnvelope）
    "envelope_object",
    # 第七十四波：拉伸消失点共享
    "extrude_vp_share",
    # 第四十九波：按掩码复制属性
    "copy_props",
    # 第五十波：形状级斜切（SkewEx）
    "skew_ex",
)


_EXTRUDE_PROPS = {
    "type": "Type", "shading": "Shading", "depth": "Depth",
    "angle_x": "AngleX", "angle_y": "AngleY", "angle_z": "AngleZ",
    "use_bevel": "UseBevel", "bevel_depth": "BevelDepth",
    "bevel_angle": "BevelAngle", "show_bevel_only": "ShowBevelOnly",
    "use_extrude_color_for_bevel": "UseExtrudeColorForBevel",
    "use_full_color_range": "UseFullColorRange",
}
_LENS_PROPS = {
    "type": "Type", "rate": "Rate", "magnification": "Magnification",
    "remove_face": "RemoveFace", "use_view_point": "UseViewPoint",
    "view_point_x": "ViewPointX", "view_point_y": "ViewPointY",
    "use_fill_color": "UseFillColor", "use_outline_color": "UseOutlineColor",
    "palette_rotation": "PaletteRotation", "color_map_palette": "ColorMapPalette",
}
_BLEND_PROPS = {
    "steps": "Steps", "spacing": "Spacing", "angle": "Angle",
    "color_blend_type": "ColorBlendType", "mode": "Mode", "loop": "Loop",
    "rotate_shapes": "RotateShapes", "map_nodes": "MapNodes",
    "fuse_start": "FuseStart", "fuse_end": "FuseEnd",
    "color_acceleration": "ColorAcceleration",
    "spacing_acceleration": "SpacingAcceleration",
    "link_acceleration": "LinkAcceleration",
    "accelerate_size": "AccelerateSize",
}
_EFFECT_COLORS = ("BaseColor", "ShadingColor", "BevelColor", "Color", "FillColor",
                  "OutlineColor", "FromColor", "ToColor")


def _apply_effect_props(doc: Any, effect: Any, spec: dict, mapping: dict) -> list:
    applied = []
    for key, attr in mapping.items():
        if key in spec and spec[key] is not None:
            setattr(effect, attr, spec[key])
            applied.append(key)
    for key, attr in (("base_color", "BaseColor"), ("shading_color", "ShadingColor"),
                      ("bevel_color", "BevelColor"), ("color", "Color"),
                      ("fill_color", "FillColor"), ("outline_color", "OutlineColor"),
                      ("from_color", "FromColor"), ("to_color", "ToColor")):
        if spec.get(key) is not None:
            setattr(effect, attr, _color(doc, spec[key]))
            applied.append(key)
    return applied


def _find_effect(shape: Any, attr: str) -> tuple:
    effects = _s(lambda: shape.Effects)
    if effects is None:
        return None, None
    count = _n(lambda: int(effects.Count), 0)
    for i in range(1, count + 1):
        eff = _s(lambda i=i: effects.Item(i))
        if eff is None:
            continue
        obj = _s(lambda eff=eff: getattr(eff, attr))
        if obj is not None:
            return eff, obj
    return None, None


def _effect_info(shape: Any) -> dict:
    effects = _s(lambda: shape.Effects)
    if effects is None:
        return {"count": 0, "effects": []}
    count = _n(lambda: int(effects.Count), 0)
    items = []
    for i in range(1, count + 1):
        eff = _s(lambda i=i: effects.Item(i))
        if eff is None:
            continue
        info = {"type": _n(lambda eff=eff: int(eff.Type))}
        for name in ("Extrude", "Lens", "Contour", "Blend", "DropShadow",
                     "Envelope", "Perspective", "Distortion"):
            obj = _s(lambda name=name, eff=eff: getattr(eff, name))
            if obj is None:
                continue
            info["kind"] = name.lower()
            info["props"] = {k: _s(lambda obj=obj, a=a: getattr(obj, a))
                             for a, k in (("Type", "type"), ("Depth", "depth"),
                                          ("Rate", "rate"),
                                          ("Magnification", "magnification"),
                                          ("Steps", "steps"),
                                          ("Spacing", "spacing"),
                                          ("Angle", "angle"),
                                          ("ColorBlendType", "color_blend_type"),
                                          ("Mode", "mode"))}
            break
        items.append(info)
    return {"count": count, "effects": items}


def shape_op(doc: Any, page_index: int | None, op: str, target: dict | None = None,
             other: dict | None = None, *, radius: float = 0, distance_a: float = 0,
             distance_b: float = 0, leave_source: bool = True, leave_target: bool = True,
             center_in_container: bool = True, symbol_name: str | None = None,
             copies: int = 1, distance_x: float = 0.0, distance_y: float = 0.0,
             mode_x: int = 1, direction_x: int = 1, mode_y: int = 1, direction_y: int = 3,
             distribute_type: int = 1, page_extent: bool = False,
             place_on_top: bool = True, delete_source: bool = False,
             x: float | None = None, y: float | None = None,
             hx: float | None = None, hy: float | None = None,
             vx: float | None = None, vy: float | None = None,
             curve_count: int = 2, curve_distance: float = 5.0,
             resolution: int = 72, grayscale: bool = False, transparent: bool = True,
             envelope_mode: int = 2, compare_type: int = 1, condition: int = 0,
             pixel_aligned: bool = True, fill_mode: int = 0,
             plane: int = 0, reference_point: int = 9,
             apply_to_duplicate: bool = False,
             extrude_type: int = 1, vp_type: int = 1,
             vp_x: float | None = None, vp_y: float | None = None,
             depth: float = 20.0, shading: int = 0,
             base_color: Any = "#808080", shading_color: Any = "#404040",
             bevel_depth: float = 0.0, bevel_angle: float = 45.0,
             bevel_color: Any = "#000000", bevel_only: bool = False,
             lens_type: int = 1, rate: float = 50.0,
             color1: Any = "#FFFFFF", color2: Any = "#000000",
             color_map_palette: int = 0, effect: dict | None = None,
             steps: int = 20, color_blend_type: int = 0, blend_mode: int = 0,
             spacing: float = 0.0, angle: float = 0.0, loop: bool = False,
             rotate_shapes: bool = False, spacing_accel: int = 0,
             color_accel: int = 0, accel_size: bool = False,
             params: dict | None = None,
             label: str = "cdr_shape_ops") -> dict:
    op = (op or "").lower().replace("-", "_")
    if op not in SHAPE_OPS:
        raise CdrBadArgument(f"未知操作 {op!r}；可用：{', '.join(SHAPE_OPS)}")

    with mm_units(doc):
        page = _page(doc, page_index)

        if op == "group" or op == "combine":
            # 以"当前选择集"为输入
            with atomic(doc, label):
                sr = selection_range(doc)
                if int(sr.Count) < 2:
                    raise CdrBadArgument(
                        f"{op} 至少需要 2 个被选中的对象；先调用 cdr_select")
                result_shape = sr.Group() if op == "group" else sr.Combine()
                # 组合本身会清空选择集，这里把结果重新选上，方便接着 ungroup / 对齐
                try:
                    if result_shape is not None:
                        doc.Selection().RemoveAll()
                        result_shape.AddToSelection()
                except Exception:
                    pass
            info = {"op": op, "done": True}
            if result_shape is not None:
                info["result_name"] = _s(lambda: str(result_shape.Name))
                info["result_type"] = SHAPE_TYPES.get(
                    _n(lambda: int(result_shape.Type), -1), "unknown")
                info["result_bbox"] = _bbox(result_shape)
            return info

        if op == "ungroup":
            # 实测：ShapeRange.Ungroup() 是空操作；必须用 Shape.Ungroup()
            with atomic(doc, label):
                shape = None
                if isinstance(target, dict) and target:
                    shape = find_shape(doc, page, target)
                else:
                    selected = doc.Selection().Shapes
                    for i in range(1, int(selected.Count) + 1):
                        cand = selected.Item(i)
                        if int(_s(lambda: int(cand.Type), -1)) == 7:
                            shape = cand
                            break
                    if shape is None and int(selected.Count) == 1:
                        shape = selected.Item(1)
                if shape is None:
                    raise CdrBadArgument(
                        "ungroup 需要 target（组合体）或在选择集中选中一个组合体；"
                        "组合体 type_id = 7")
                shape.Ungroup()
                still_grouped = int(_s(lambda: int(shape.Type), -1)) == 7
            return {"op": op, "done": True,
                    "warning": "该对象未完全解开（可能含子组合）" if still_grouped else None}

        if op == "to_document":
            with atomic(doc, label):
                if target:
                    src_shape = find_shape(doc, page, target)
                    new_doc = src_shape.CreateDocumentFrom(False)
                else:
                    sr = selection_range(doc)
                    if _n(lambda: int(sr.Count), 0) < 1:
                        raise CdrBadArgument("to_document 需要 target 或先选中对象")
                    new_doc = sr.CreateDocumentFrom(False)
            return {"op": op,
                    "document_name": _s(lambda: str(new_doc.Name)),
                    "pages": _n(lambda: int(new_doc.Pages.Count), 0),
                    "shapes": total_shapes(new_doc),
                    "note": "新文档已创建；可用 cdr_documents 获取 document_id"}

        if op == "apply_fill":
            # 实测：Document.CreateUniformFill(IVGColor) → IVGFill，
            # 再 ShapeRange.ApplyFill(fill) 可一次给多个对象上均匀填充
            with atomic(doc, label):
                sr = selection_range(doc)
                if target:
                    try:
                        doc.Selection().RemoveAll()
                    except Exception:
                        pass
                    find_shape(doc, page, target).AddToSelection()
                    sr = selection_range(doc)
                if _n(lambda: int(sr.Count), 0) < 1:
                    raise CdrBadArgument("apply_fill 需要先选中对象，或用 target 指定")
                color = _color(doc, base_color)
                fill = doc.CreateUniformFill(color)
                if fill is None:
                    raise CdrComFailure("Document.CreateUniformFill 返回空")
                sr.ApplyFill(fill)
                first = _s(lambda: sr.Item(1))
                return {"op": op, "count": _n(lambda: int(sr.Count), 0),
                        "fill_type": _n(lambda: int(first.Fill.Type), 0)
                        if first is not None else None,
                        "color": _s(lambda: str(color.ToString()))}
        if op == "range_outline":
            spec = dict(params or {})
            with atomic(doc, label):
                sr = selection_range(doc)
                if target:
                    try:
                        doc.Selection().RemoveAll()
                    except Exception:
                        pass
                    find_shape(doc, page, target).AddToSelection()
                    sr = selection_range(doc)
                if _n(lambda: int(sr.Count), 0) < 1:
                    raise CdrBadArgument("range_outline 需要先选中对象，或用 target 指定")
                width = float(spec.get("width", 1.0))
                caps = int(spec.get("line_caps", 0))
                join = int(spec.get("line_join", 0))
                behind = int(bool(spec.get("behind_fill", False)))
                scale = int(bool(spec.get("scale_with_shape", True)))
                nib_angle = float(spec.get("nib_angle", 0.0))
                nib_stretch = int(spec.get("nib_stretch", 100))
                dashdot = float(spec.get("dashdot_length", 0.0))
                pen_width = float(spec.get("pen_width", 1.0))
                miter = float(spec.get("miter_limit", 5.0))
                if spec.get("with_justification"):
                    sr.SetOutlinePropertiesEx(width, None, None, None, None, behind,
                                              scale, caps, join, nib_angle, nib_stretch,
                                              dashdot, pen_width, miter,
                                              int(spec.get("justification", 0)))
                    used = "SetOutlinePropertiesEx"
                else:
                    sr.SetOutlineProperties(width, None, None, None, None, behind,
                                            scale, caps, join, nib_angle, nib_stretch,
                                            dashdot, pen_width, miter)
                    used = "SetOutlineProperties"
                sr2 = selection_range(doc)
                first = _s(lambda: sr2.Item(1))
            return {"op": op, "count": _n(lambda: int(sr2.Count), 0),
                    "api": used,
                    "first_width": _round(_s(lambda: float(first.Outline.Width)))
                    if first is not None else None}
        if op in {"set_size_ex", "stretch_ex"}:
            # 实测：SetSizeEx(CenterX, CenterY, Width, Height) 以中心定尺寸；
            # StretchEx(CenterX, CenterY, StretchX, StretchY, StretchCharactersSize) 以中心按倍数拉伸
            spec = dict(params or {})
            use_range = bool(spec.get("use_range", False))
            target_obj = None
            if use_range:
                target_obj = selection_range(doc)
                if _n(lambda: int(target_obj.Count), 0) < 1:
                    raise CdrBadArgument("use_range=true 时需要有选中对象")
            else:
                target_obj = find_shape(doc, page, target or {})
            before = _bbox(target_obj)
            cx = float(spec.get("cx", 0.0))
            cy = float(spec.get("cy", 0.0))
            with atomic(doc, label):
                if op == "set_size_ex":
                    target_obj.SetSizeEx(cx, cy, float(spec.get("width", 10.0)),
                                         float(spec.get("height", 10.0)))
                else:
                    target_obj.StretchEx(cx, cy, float(spec.get("stretch_x", 1.0)),
                                         float(spec.get("stretch_y", 1.0)),
                                         bool(spec.get("stretch_characters", False)))
            after = _bbox(target_obj)
            return {"op": op, "scope": "range" if use_range else "shape",
                    "center": [cx, cy], "bbox_before": before, "bbox_after": after}
        if op == "all_excluding":
            # 实测：ShapeRange.AllExcluding(索引数组) 用 **1 基**索引排除成员，
            # 返回新的 ShapeRange（Python list 直接可用）
            spec = dict(params or {})
            exclude = spec.get("exclude") or []
            sr_src = selection_range(doc)
            if _n(lambda: int(sr_src.Count), 0) < 1:
                raise CdrBadArgument("all_excluding 需要先选中对象")
            got = _s(lambda: sr_src.AllExcluding(list(exclude)))
            if got is None:
                raise CdrComFailure("ShapeRange.AllExcluding 返回空")
            count = _n(lambda: int(got.Count), 0) or 0
            return {"op": op, "source_count": _n(lambda: int(sr_src.Count), 0),
                    "excluded": list(exclude), "result_count": count,
                    "names": [_s(lambda i=i: str(got.Item(i).Name))
                              for i in range(1, min(count, 20) + 1)]}
        if op == "envelope_object":
            # 实测：Shape.CreateEnvelope(PresetIndex(1 基), Mode, KeepLines) 返回 IVGEffect；
            # effect.Envelope 是 IVGEffectEnvelope，Select / CreateFrom / CopyFrom /
            # CopyFromCurve 都可用（CopyFromShape 参数形态未通）
            spec = dict(params or {})
            source = find_shape(doc, page, target or {})
            preset = int(spec.get("preset", 1))
            applied = []
            with atomic(doc, label):
                eff = _s(lambda: source.CreateEnvelope(preset, int(spec.get("mode", 0)),
                                                       bool(spec.get("keep_lines", True))))
                env = None
                if eff is not None:
                    env = _s(lambda: eff.Envelope)
                if env is None:
                    effects = _s(lambda: source.Effects)
                    if effects is not None and (_n(lambda: int(effects.Count), 0) or 0) > 0:
                        env = _s(lambda: effects.Item(1).Envelope)
                if env is None:
                    raise CdrComFailure("取不到 Envelope 效果对象")
                if spec.get("select") is not None:
                    env.Select(int(spec["select"]))
                    applied.append("select")
                if other:
                    src_shape = find_shape(doc, page, other)
                    if spec.get("copy_from"):
                        # 从另一个已带包络的对象复制效果（IVGEffectEnvelope.CopyFrom）
                        src_eff = _s(lambda: src_shape.Effects.Item(1))
                        src_env = _s(lambda: src_eff.Envelope) if src_eff is not None else None
                        if src_env is not None:
                            env.CopyFrom(src_env)
                            applied.append("copy_from")
                    elif spec.get("from_curve"):
                        src_curve = _s(lambda: src_shape.Curve)
                        if src_curve is not None:
                            env.CopyFromCurve(src_curve, int(spec.get("mode", 0)),
                                              bool(spec.get("keep_lines", True)), 0, ())
                            applied.append("copy_from_curve")
                    else:
                        env.CreateFrom(src_shape)
                        applied.append("create_from")
            return {"op": op, "preset": preset,
                    "mode": _n(lambda: int(env.Mode)),
                    "keep_lines": bool(_s(lambda: bool(env.KeepLines), False)),
                    "corner_indices": _s(lambda: str(env.CornerIndices))[:60],
                    "applied": applied}
        if op in {"hatch_fill", "fountain_colors", "ellipse_geometry"}:
            source = find_shape(doc, page, target or {})
            spec = dict(params or {})
            applied = []
            if op == "hatch_fill":
                # 实测：Fill.Hatch → IVGHatchFill；GetFillScale/GetLineScale 零参，
                # SetFillScale/SetLineScale/SetNoBackColor 可用
                hf = _s(lambda: source.Fill.Hatch)
                if hf is None:
                    raise CdrBadArgument("目标不是剖面线填充（先用 fill=hatch）")
                with atomic(doc, label):
                    if spec.get("fill_scale") is not None:
                        hf.SetFillScale(float(spec["fill_scale"]))
                        applied.append("fill_scale")
                    if spec.get("line_scale") is not None:
                        hf.SetLineScale(float(spec["line_scale"]))
                        applied.append("line_scale")
                    if spec.get("no_back_color"):
                        hf.SetNoBackColor()
                        applied.append("no_back_color")
                return {"op": op, "fill_scale": _round(_s(lambda: float(hf.GetFillScale()))),
                        "line_scale": _round(_s(lambda: float(hf.GetLineScale()))),
                        "applied": applied}
            if op == "fountain_colors":
                # 实测：Fill.Fountain.Colors → IVGFountainColors；
                # GrayLevel(i) 读、SetGrayLevel(i, v) 写（Count 可读）
                fountain = _s(lambda: source.Fill.Fountain)
                colors = _s(lambda: fountain.Colors) if fountain is not None else None
                if colors is None:
                    raise CdrBadArgument("目标不是渐变填充")
                count = _n(lambda: int(colors.Count), 0) or 0
                with atomic(doc, label):
                    if spec.get("index") is not None and spec.get("gray_level") is not None:
                        colors.SetGrayLevel(int(spec["index"]),
                                            float(spec["gray_level"]))
                        applied.append("set_gray_level")
                idx = int(spec.get("index", 1))
                return {"op": op, "count": count, "index": idx,
                        "gray_level": _round(_s(lambda: float(colors.GrayLevel(idx)))),
                        "levels": [_round(_s(lambda i=i: float(colors.GrayLevel(i))))
                                   for i in range(1, min(count, 10) + 1)],
                        "applied": applied}
            # ellipse_geometry
            ell = _s(lambda: source.Ellipse)
            if ell is None:
                raise CdrBadArgument("目标不是椭圆（没有 Ellipse 子对象）")
            radius = _s(lambda: ell.GetRadius())
            with atomic(doc, label):
                if spec.get("center_x") is not None and spec.get("center_y") is not None:
                    ell.SetCenterPosition(float(spec["center_x"]),
                                          float(spec["center_y"]))
                    applied.append("set_center")
            return {"op": op,
                    "radius": ([_round(float(v)) for v in radius]
                               if isinstance(radius, (tuple, list)) else None),
                    "bbox": _bbox(source), "applied": applied}
        if op in {"corners", "clone_links", "drop_shadow"}:
            # 实测入口：Shape.Rectangle（IVGRectangle）/ Shape.CloneLink（IVGCloneLink）/
            # Shape.Effects.Item(1).DropShadow（IVGEffectDropShadow）
            source = find_shape(doc, page, target or {})
            spec = dict(params or {})
            applied = []
            if op == "corners":
                rgn = _s(lambda: source.Rectangle)
                if rgn is None:
                    raise CdrBadArgument("目标不是矩形（没有 Rectangle 子对象）")
                with atomic(doc, label):
                    if spec.get("radius") is not None:
                        rgn.SetRadius(float(spec["radius"]))
                        applied.append("set_radius")
                    if spec.get("roundness") is not None:
                        rgn.SetRoundness(float(spec["roundness"]))
                        applied.append("set_roundness")
                return {"op": op, "applied": applied,
                        "equal_corners": bool(_s(lambda: bool(rgn.EqualCorners), False)),
                        "max_radius": _round(_s(lambda: float(rgn.MaxRadius))),
                        "radius_upper_left": _round(
                            _s(lambda: float(rgn.RadiusUpperLeft))),
                        "corner_type": _n(lambda: int(rgn.CornerType))}
            if op == "clone_links":
                cl = _s(lambda: source.CloneLink)
                if cl is None:
                    raise CdrBadArgument("目标不是克隆对象（先用 clone 生成）")
                with atomic(doc, label):
                    if spec.get("restore"):
                        cl.RestoreAllLinks()
                        applied.append("restore_all_links")
                return {"op": op, "applied": applied,
                        "fill_linked": bool(_s(lambda: bool(cl.FillLinked), False)),
                        "outline_linked": bool(_s(lambda: bool(cl.OutlineLinked), False)),
                        "shape_linked": bool(_s(lambda: bool(cl.ShapeLinked), False)),
                        "transform_linked": bool(
                            _s(lambda: bool(cl.TransformLinked), False)),
                        "parent": _s(lambda: str(cl.CloneParent.Name))}
            # drop_shadow
            effects = _s(lambda: source.Effects)
            if effects is None or (_n(lambda: int(effects.Count), 0) or 0) < 1:
                raise CdrBadArgument("目标没有效果（先用 cdr_modify 加投影）")
            eff = _s(lambda: effects.Item(1))
            sh = _s(lambda: eff.DropShadow) if eff is not None else None
            if sh is None:
                raise CdrBadArgument("该效果不是投影")
            with atomic(doc, label):
                if spec.get("offset_x") is not None or spec.get("offset_y") is not None:
                    sh.SetOffset(float(spec.get("offset_x", 0.0)),
                                 float(spec.get("offset_y", 0.0)))
                    applied.append("set_offset")
                for key, attr in (("opacity", "Opacity"), ("feather", "Feather"),
                                  ("fade", "Fade"), ("type", "Type")):
                    if key in spec:
                        try:
                            setattr(sh, attr, spec[key])
                            applied.append(key)
                        except Exception:
                            continue
            return {"op": op, "applied": applied,
                    "offset": [_round(_s(lambda: float(sh.OffsetX))),
                               _round(_s(lambda: float(sh.OffsetY)))],
                    "opacity": _n(lambda: int(sh.Opacity)),
                    "feather": _round(_s(lambda: float(sh.Feather))),
                    "fade": _n(lambda: int(sh.Fade)),
                    "type": _n(lambda: int(sh.Type))}
        if op == "blend_effect":
            # 实测：形状混合后 shape.Effects.Item(1).Blend 就是 IVGEffectBlend，
            # FuseStart / FuseEnd / Split 可用（CopyFrom 需要另一条有效混合）
            source = find_shape(doc, page, target or {})
            effects = _s(lambda: source.Effects)
            if effects is None or (_n(lambda: int(effects.Count), 0) or 0) < 1:
                raise CdrBadArgument("目标没有效果（先用 blend 创建混合）")
            eff = _s(lambda: effects.Item(1))
            blend = _s(lambda: eff.Blend) if eff is not None else None
            if blend is None:
                raise CdrBadArgument("该效果不是混合（Type 不是 blend）")
            spec = dict(params or {})
            applied = []
            with atomic(doc, label):
                if spec.get("fuse_start"):
                    blend.FuseStart()
                    applied.append("fuse_start")
                if spec.get("fuse_end"):
                    blend.FuseEnd()
                    applied.append("fuse_end")
                if spec.get("split") is not None:
                    made = _s(lambda: blend.Split(float(spec["split"])))
                    applied.append("split")
            return {"op": op, "effect_type": _n(lambda: int(eff.Type)),
                    "steps": _n(lambda: int(blend.Steps)),
                    "mode": _n(lambda: int(blend.Mode)),
                    "start_shape": _s(lambda: str(blend.StartShape.Name)),
                    "end_shape": _s(lambda: str(blend.EndShape.Name)),
                    "applied": applied,
                    "split_shape": (_s(lambda: str(made.Name))
                                    if spec.get("split") is not None else None)}
        if op in {"dimension_info", "dimension_set"}:
            # 实测：标注形状的 `Shape.Dimension` 就是 IVGDimension
            source = find_shape(doc, page, target or {})
            dim = _s(lambda: source.Dimension)
            if dim is None:
                raise CdrBadArgument("目标不是标注对象（先用 linear_dimension 创建）")
            spec = dict(params or {})
            applied = []
            if op == "dimension_set":
                with atomic(doc, label):
                    for key, fn in (
                            ("text_label_gap", lambda v: dim.SetTextLabelGap(float(v))),
                            ("overhang", lambda v: dim.SetOverhang(float(v))),
                            ("distance_from_object",
                             lambda v: dim.SetDistanceFromObject(float(v)))):
                        if key in spec:
                            try:
                                fn(spec[key])
                                applied.append(key)
                            except Exception:
                                continue
                    linear_obj = _s(lambda: dim.Linear)
                    if linear_obj is not None and spec.get(
                            "reverse_terminator_length") is not None:
                        try:
                            linear_obj.SetReverseTerminatorLength(
                                float(spec["reverse_terminator_length"]))
                            applied.append("reverse_terminator_length")
                        except Exception:
                            pass
                    for key, attr in (("precision", "Precision"),
                                      ("prefix", "Prefix"), ("suffix", "Suffix"),
                                      ("boxed_text", "BoxedText"),
                                      ("show_units", "ShowUnits"),
                                      ("leading_zero", "LeadingZero"),
                                      ("text_centered", "TextCentered"),
                                      ("horizontal_text", "HorizontalText"),
                                      ("dynamic_text", "DynamicText"),
                                      ("placement", "Placement")):
                        if key in spec:
                            try:
                                setattr(dim, attr, spec[key])
                                applied.append(key)
                            except Exception:
                                continue
            linear_info = _s(lambda: dim.Linear)
            return {"op": op,
                    "linear_type": (_n(lambda: int(linear_info.Type))
                                    if linear_info is not None else None),
                    "reverse_terminators": (
                        bool(_s(lambda: bool(linear_info.ReverseTerminators), False))
                        if linear_info is not None else None),
                    "precision": _n(lambda: int(dim.Precision)),
                    "prefix": _s(lambda: str(dim.Prefix)),
                    "suffix": _s(lambda: str(dim.Suffix)),
                    "linear": bool(_s(lambda: bool(dim.Linear), False)),
                    "boxed_text": bool(_s(lambda: bool(dim.BoxedText), False)),
                    "show_units": bool(_s(lambda: bool(dim.ShowUnits), False)),
                    "text_label_gap": _round(_s(lambda: float(dim.TextLabelGap))),
                    "overhang": _round(_s(lambda: float(dim.Overhang))),
                    "distance_from_object": _round(
                        _s(lambda: float(dim.DistanceFromObject))),
                    "applied": applied,
                    "note": "标注用 Shape.Dimension 取 IVGDimension；"
                            "Set* 三个方法调用不报错，但读回可能仍为 0"}
        if op == "powerclip_extract":
            # 实测：Shape.PowerClip → IVGPowerClip；EnterEditMode / LeaveEditMode /
            # ExtractShapes 都可用，ExtractShapes 会把内容取回页面（对象数 +1）
            source = find_shape(doc, page, target or {})
            pc = _s(lambda: source.PowerClip)
            if pc is None:
                raise CdrBadArgument("目标不是图框（先用 powerclip 装入内容）")
            contents = _n(lambda: int(pc.Shapes.Count), 0)
            before = _n(lambda: int(page.Shapes.Count), 0) or 0
            with atomic(doc, label):
                try:
                    pc.EnterEditMode()
                    entered = True
                except Exception:
                    entered = False
                made = None
                try:
                    made = _s(lambda: pc.ExtractShapes())
                finally:
                    try:
                        pc.LeaveEditMode()
                    except Exception:
                        pass
            after = _n(lambda: int(page.Shapes.Count), 0) or 0
            return {"op": op, "contents": contents,
                    "entered_edit_mode": entered,
                    "extracted": (_n(lambda: int(made.Count), 0)
                                  if made is not None else None),
                    "page_shapes_before": before, "page_shapes_after": after,
                    "contents_locked": bool(_s(lambda: bool(pc.ContentsLocked), False))}
        if op in {"guide_info", "guide_set"}:
            # 实测：辅助线形状的 `Shape.Guide` 是 IVGGuide：
            # MakeEditable / GetPoints / GetPointAndAngle / SetPointAndAngle 可用
            source = find_shape(doc, page, target or {})
            gd = _s(lambda: source.Guide)
            if gd is None:
                raise CdrBadArgument("目标不是辅助线（先用 create_guide 创建）")
            spec = dict(params or {})
            applied = []
            with atomic(doc, label):
                try:
                    gd.MakeEditable()
                    applied.append("make_editable")
                except Exception:
                    pass
                if op == "guide_set":
                    if all(k in spec for k in ("x", "y", "angle")):
                        gd.SetPointAndAngle(float(spec["x"]), float(spec["y"]),
                                            float(spec["angle"]))
                        applied.append("set_point_and_angle")
                    elif all(k in spec for k in ("x1", "y1", "x2", "y2")):
                        gd.SetPoints(float(spec["x1"]), float(spec["y1"]),
                                     float(spec["x2"]), float(spec["y2"]))
                        applied.append("set_points")
            pts = _s(lambda: gd.GetPoints(0.0, 0.0, 0.0, 0.0))
            pa = _s(lambda: gd.GetPointAndAngle(0.0, 0.0, 0.0))
            return {"op": op, "type": _n(lambda: int(gd.Type)),
                    "angle": _round(_s(lambda: float(gd.Angle))),
                    "point1": [_round(_s(lambda: float(gd.Point1X))),
                               _round(_s(lambda: float(gd.Point1Y)))],
                    "point2": [_round(_s(lambda: float(gd.Point2X))),
                               _round(_s(lambda: float(gd.Point2Y)))],
                    "get_points": ([_round(float(v)) for v in pts]
                                   if isinstance(pts, (tuple, list)) else None),
                    "get_point_and_angle": ([_round(float(v)) for v in pa]
                                            if isinstance(pa, (tuple, list)) else None),
                    "applied": applied}
        if op == "sibling":
            # 实测：Shape.GetNext(方向) / GetPrevious(方向) 返回 IVGShape（同级对象）
            spec = dict(params or {})
            direction = str(spec.get("direction", "next")).lower()
            source = find_shape(doc, page, target or {})
            arg = 0
            got = None
            chain: list = []
            steps = max(1, min(int(spec.get("steps", 1)), 50))
            cur = source
            for _ in range(steps):
                try:
                    got = (cur.GetNext(arg) if direction == "next"
                           else cur.GetPrevious(arg))
                except Exception:
                    got = None
                if got is None:
                    break
                chain.append(_s(lambda g=got: str(g.Name)))
                cur = got
            return {"op": op, "direction": direction, "steps": steps,
                    "found": got is not None,
                    "names": chain,
                    "last_type": (_n(lambda: int(got.Type))
                                  if got is not None else None),
                    "last_bbox": _bbox(got) if got is not None else None}
        if op == "range_outline_apply":
            # 实测：ShapeRange.ApplyOutline(Outline) —— 用一个描边对象整批覆盖范围内所有对象
            if not other:
                raise CdrBadArgument("range_outline_apply 需要 other（描边来源对象）")
            source_shape = find_shape(doc, page, other)
            outline = _s(lambda: source_shape.Outline.GetCopy())
            if outline is None:
                raise CdrComFailure("Outline.GetCopy 返回空")
            spec = dict(params or {})
            if "width" in spec:
                try:
                    outline.Width = float(spec["width"])
                except Exception:
                    pass
            with atomic(doc, label):
                sr = selection_range(doc)
                if target:
                    try:
                        doc.Selection().RemoveAll()
                    except Exception:
                        pass
                    find_shape(doc, page, target).AddToSelection()
                    sr = selection_range(doc)
                sr.ApplyOutline(outline)
                sr2 = selection_range(doc)
                first = _s(lambda: sr2.Item(1))
            return {"op": op, "count": _n(lambda: int(sr2.Count), 0),
                    "source": _s(lambda: str(source_shape.Name)),
                    "first_width": _round(_s(lambda: float(first.Outline.Width)))
                    if first is not None else None}
        if op == "skew_ex":
            spec = dict(params or {})
            source = find_shape(doc, page, target or {})
            before = _bbox(source)
            with atomic(doc, label):
                source.SkewEx(float(spec.get("angle_x", 0.0)),
                              float(spec.get("angle_y", 0.0)),
                              float(spec.get("cx", 0.0)),
                              float(spec.get("cy", 0.0)))
            return {"op": op, "bbox_before": before, "bbox_after": _bbox(source),
                    "angles": [float(spec.get("angle_x", 0.0)),
                               float(spec.get("angle_y", 0.0))],
                    "center": [float(spec.get("cx", 0.0)), float(spec.get("cy", 0.0))]}
        if op == "copy_props":
            # 实测：Shape.CopyPropertiesFrom(source, 掩码) 可用；
            # 掩码 4/8/15 会把填充复制过来，1/2 不影响填充
            spec = dict(params or {})
            source = find_shape(doc, page, target or {})
            if not other:
                raise CdrBadArgument("copy_props 需要 other（源对象）")
            mask = int(spec.get("mask", 15))
            src = find_shape(doc, page, other)
            with atomic(doc, label):
                ok = bool(_s(lambda: bool(source.CopyPropertiesFrom(src, mask)), False))
            return {"op": op, "from": _s(lambda: str(src.Name)), "mask": mask,
                    "applied": ok, "bbox": _bbox(source),
                    "note": "实测掩码 4/8/15 会复制填充；1/2 不影响填充"}
        if op in {"copy_to_layer_range", "align_to_range"}:
            spec = dict(params or {})
            source = find_shape(doc, page, target or {})
            if op == "copy_to_layer_range":
                # 实测：Shape.CopyToLayerAsRange(Layer) 返回新图层上的 ShapeRange
                layer_name = str(spec.get("layer", ""))
                if not layer_name:
                    raise CdrBadArgument("copy_to_layer_range 需要 params.layer（目标图层名）")
                dest = _find_layer(page, layer_name, None)
                with atomic(doc, label):
                    rng = source.CopyToLayerAsRange(dest)
                if rng is None:
                    raise CdrComFailure("Shape.CopyToLayerAsRange 返回空")
                count = _n(lambda: int(rng.Count), 0) or 0
                first = _s(lambda: rng.Item(1))
                return {"op": op, "count": count,
                        "layer": _s(lambda: str(dest.Name)),
                        "new_name": _s(lambda: str(first.Name)) if first is not None else None,
                        "new_bbox": _bbox(first) if first is not None else None}
            # align_to_range：把对象按 cdrAlignType 对齐到另一批对象组成的范围
            if not other:
                raise CdrBadArgument("align_to_range 需要 other（参照对象）")
            align_type = int(spec.get("align_type", 1))
            ref = find_shape(doc, page, other)
            rng = _app_of(doc).CreateShapeRange()
            rng.Add(ref)
            with atomic(doc, label):
                source.AlignToShapeRange(align_type, rng, int(spec.get("text_origin", 0)))
            return {"op": op, "align_type": align_type,
                    "reference": _s(lambda: str(ref.Name)),
                    "bbox": _bbox(source)}
        if op in {"range_from_list", "range_bitmap"}:
            spec = dict(params or {})
            targets = list(spec.get("targets") or [])
            if op == "range_from_list" and not targets:
                raise CdrBadArgument(
                    'range_from_list 需要 params.targets=[{"name":…}, …]（至少 1 个对象）')
            if op == "range_from_list":
                # 实测 Document.CreateShapeRangeFromArray 接受 Python list[Shape]
                shapes = [find_shape(doc, page, t) for t in targets]
                rng = _s(lambda: doc.CreateShapeRangeFromArray(shapes))
                if rng is None:
                    raise CdrComFailure("Document.CreateShapeRangeFromArray 返回空")
                count = _n(lambda: int(rng.Count), 0) or 0
                sub = str(spec.get("then", "")).lower()
                applied = None
                if sub == "rotate_ex":
                    with atomic(doc, label):
                        rng.RotateEx(float(spec.get("angle", 30.0)),
                                     float(spec.get("cx", 0.0)),
                                     float(spec.get("cy", 0.0)))
                    applied = "rotate_ex"
                return {"op": op, "count": count,
                        "names": [_s(lambda i=i: str(rng.Item(i).Name))
                                  for i in range(1, min(count, 30) + 1)],
                        "exists_first": bool(_s(
                            lambda: bool(rng.Exists(shapes[0])), False)),
                        "applied": applied,
                        "note": "用 Document.CreateShapeRangeFromArray 按对象列表建范围，"
                                "可再 then=rotate_ex 整批旋转"}
            # range_bitmap：把当前选择集整体位图化
            with atomic(doc, label):
                sr = selection_range(doc)
                if target:
                    try:
                        doc.Selection().RemoveAll()
                    except Exception:
                        pass
                    find_shape(doc, page, target).AddToSelection()
                    sr = selection_range(doc)
                if _n(lambda: int(sr.Count), 0) < 1:
                    raise CdrBadArgument("range_bitmap 需要先选中对象，或用 target 指定")
                made = sr.ConvertToBitmapEx(
                    int(spec.get("mode", 4)), bool(spec.get("dithered", False)),
                    bool(spec.get("transparent", True)),
                    int(spec.get("resolution", 150)),
                    int(spec.get("anti_aliasing", 2)),
                    bool(spec.get("use_color_profile", False)),
                    bool(spec.get("overprint_black", False)),
                    int(spec.get("overprint_black_limit", 0)))
            if made is None:
                raise CdrComFailure("ShapeRange.ConvertToBitmapEx 返回空")
            return {"op": op, "type_id": _n(lambda: int(made.Type)),
                    "bbox": _bbox(made),
                    "has_bitmap": bool(_s(lambda: made.Bitmap is not None, False))}
        if op in {"position_ex", "find_snap", "linked", "rasterize_ex",
                  "overprint_state"}:
            source = find_shape(doc, page, target or {})
            spec = dict(params or {})
            if op == "position_ex":
                ref_point = int(spec.get("ref", 9))
                if spec.get("x") is not None and spec.get("y") is not None:
                    with atomic(doc, label):
                        source.SetPositionEx(ref_point, float(spec["x"]),
                                             float(spec["y"]))
                    got = _s(lambda: source.GetPositionEx(ref_point))
                    return {"op": op, "ref": ref_point, "applied": True,
                            "position": [round(float(v), 3) for v in got]
                            if isinstance(got, (tuple, list)) else None,
                            "bbox": _bbox(source)}
                got = _s(lambda: source.GetPositionEx(ref_point))
                return {"op": op, "ref": ref_point, "applied": False,
                        "position": [round(float(v), 3) for v in got]
                        if isinstance(got, (tuple, list)) else None,
                        "note": "cdrReferencePoint：1 右上 / 2 上中 / 3 左上 / 4 左中 / "
                                "5 左下 / 6 下中 / 7 右下 / 8 右中 / 9 中心",
                        "bbox": _bbox(source)}
            if op == "find_snap":
                refdata = str(spec.get("reference", ""))
                if not refdata:
                    raise CdrBadArgument(
                        "find_snap 需要 params.reference（吸附点的 ReferenceData，"
                        "可从 cdr_snap action=list 的 reference 字段取）")
                snap = _s(lambda: source.FindSnapPoint(refdata))
                if snap is None:
                    return {"op": op, "found": False, "reference": refdata}
                return {"op": op, "found": True, "reference": refdata,
                        "type": _n(lambda: int(snap.Type)),
                        "position": [_round(_s(lambda: float(snap.PositionX))),
                                     _round(_s(lambda: float(snap.PositionY)))]}
            if op == "linked":
                link_type = int(spec.get("link_type", 0))
                rng = _s(lambda: source.GetLinkedShapes(link_type))
                count = _n(lambda: int(rng.Count), 0) if rng is not None else 0
                return {"op": op, "link_type": link_type, "count": count or 0,
                        "names": [_s(lambda i=i: str(rng.Item(i).Name))
                                  for i in range(1, min(count or 0, 20) + 1)],
                        "note": "cdrShapeLinkType：0 直接依赖 / 3 全部依赖 / 6 直接控制 / 8 全部连接"}
            if op == "rasterize_ex":
                before = _bbox(source)
                with atomic(doc, label):
                    made = source.ConvertToBitmapEx(
                        int(spec.get("mode", 4)), bool(spec.get("dithered", False)),
                        bool(spec.get("transparent", True)),
                        int(spec.get("resolution", 150)),
                        int(spec.get("anti_aliasing", 2)),
                        bool(spec.get("use_color_profile", False)),
                        bool(spec.get("overprint_black", False)),
                        int(spec.get("overprint_black_limit", 0)))
                if made is None:
                    raise CdrComFailure("Shape.ConvertToBitmapEx 返回空")
                return {"op": op, "type_id": _n(lambda: int(made.Type)),
                        "bbox_before": before, "bbox_after": _bbox(made),
                        "has_bitmap": bool(_s(lambda: made.Bitmap is not None, False)),
                        "page_shapes": _n(lambda: int(page.Shapes.Count), 0)}
            # overprint_state
            return {"op": op,
                    "fill_state": _n(lambda: int(source.GetOverprintFillState())),
                    "outline_state": _n(lambda: int(source.GetOverprintOutlineState())),
                    "overprint_fill": bool(_s(lambda: bool(source.OverprintFill), False)),
                    "overprint_outline": bool(_s(
                        lambda: bool(source.OverprintOutline), False)),
                    "note": "实测 outline 状态 2→4 随叠印开关变化；fill 状态恒为 1，"
                            "判断填充是否叠印请读 OverprintFill 布尔属性"}
        if op in {"matrix", "is_type", "in_front_of", "is_on_shape", "distort",
                  "place_text", "duplicate_range", "clone_range"}:
            source = find_shape(doc, page, target or {})
            spec = dict(params or {})
            if op == "matrix":
                mat = spec.get("matrix")
                if mat:
                    vals = [float(v) for v in list(mat)[:6]]
                    if len(vals) != 6:
                        raise CdrBadArgument(
                            "matrix 需要 6 个数：[d11,d12,d21,d22,tx,ty]")
                    with atomic(doc, label):
                        source.SetMatrix(*vals)
                    read = _s(lambda: source.GetMatrix())
                    return {"op": op, "applied": vals,
                            "matrix": [round(float(v), 4) for v in read]
                            if isinstance(read, (tuple, list)) else None,
                            "bbox": _bbox(source)}
                read = _s(lambda: source.GetMatrix())
                return {"op": op,
                        "matrix": [round(float(v), 4) for v in read]
                        if isinstance(read, (tuple, list)) else None,
                        "note": "矩阵为 [d11,d12,d21,d22,tx,ty]；tx/ty 用 X8 内部单位"
                                "（实测 5 → 约 5/100000 mm 的表达），读回是精确的",
                        "bbox": _bbox(source)}
            if op == "is_type":
                types = [int(t) for t in (spec.get("types") or [])]
                if not types:
                    raise CdrBadArgument("is_type 需要 params.types=[类型号…]")
                return {"op": op, "types": types,
                        "type_id": _n(lambda: int(source.Type)),
                        "matches": bool(_s(lambda: bool(source.IsTypeAnyOf(types)), False))}
            if op == "in_front_of":
                if not other:
                    raise CdrBadArgument("in_front_of 需要 other（另一个对象）")
                other_shape = find_shape(doc, page, other)
                return {"op": op, "a": _s(lambda: str(source.Name)),
                        "b": _s(lambda: str(other_source.Name)),
                        "in_front": bool(_s(
                            lambda: bool(source.OrderIsInFrontOf(other_shape)), False))}
            if op == "is_on_shape":
                if x is None or y is None:
                    raise CdrBadArgument("is_on_shape 需要 x / y")
                hot = float(spec.get("hot_area", 2.0))
                state = _n(lambda: int(source.IsOnShape(float(x), float(y), hot)))
                return {"op": op, "x": float(x), "y": float(y), "hit_state": state,
                        "inside": state == 2, "on_edge": state == 1,
                        "outside": state == 0,
                        "note": "实测 IsOnShape 返回 2=内部、1=边界、0=外部"}
            if op == "distort":
                kind = str(spec.get("kind", "pushpull")).lower()
                ox = float(spec.get("origin_x", x if x is not None else 0.0))
                oy = float(spec.get("origin_y", y if y is not None else 0.0))
                before = _bbox(source)
                with atomic(doc, label):
                    if kind in {"pushpull", "push_pull"}:
                        source.CreatePushPullDistortion(
                            ox, oy, int(spec.get("amplitude", 20)))
                    elif kind == "twister":
                        source.CreateTwisterDistortion(
                            ox, oy, float(spec.get("angle", 30)))
                    elif kind == "zipper":
                        source.CreateZipperDistortion(
                            ox, oy, int(spec.get("amplitude", 10)),
                            int(spec.get("frequency", 10)),
                            bool(spec.get("random", False)),
                            bool(spec.get("smooth", True)),
                            bool(spec.get("local", False)))
                    else:
                        raise CdrBadArgument(
                            "distort 的 params.kind 可用 pushpull / twister / zipper")
                return {"op": op, "kind": kind, "origin": [ox, oy],
                        "bbox_before": before, "bbox_after": _bbox(source),
                        "effects": _n(lambda: int(source.Effects.Count), 0)}
            if op == "place_text":
                if not other:
                    raise CdrBadArgument("place_text 需要 other（要内嵌的文字对象）")
                text_shape = find_shape(doc, page, other)
                if int(_s(lambda: int(text_shape.Type), -1)) != TEXT_TYPE:
                    raise CdrBadArgument("other 必须是文字对象（type_id=6）")
                with atomic(doc, label):
                    made = source.PlaceTextInside(text_shape)
                if made is None:
                    raise CdrComFailure("Shape.PlaceTextInside 返回空")
                return {"op": op, "container": _s(lambda: str(source.Name)),
                        "type_id": _n(lambda: int(made.Type)),
                        "text": _s(lambda: str(made.Text.Story.Text)),
                        "bbox": _bbox(made),
                        "page_shapes": _n(lambda: int(page.Shapes.Count), 0)}
            # duplicate_range / clone_range
            offx = float(spec.get("offset_x", 5.0))
            offy = float(spec.get("offset_y", 5.0))
            with atomic(doc, label):
                rng = (source.CloneAsRange(offx, offy) if op == "clone_range"
                       else source.DuplicateAsRange(offx, offy))
            if rng is None:
                raise CdrComFailure(f"Shape.{op} 返回空")
            count = _n(lambda: int(rng.Count), 0) or 0
            return {"op": op, "count": count, "offset": [offx, offy],
                    "names": [_s(lambda i=i: str(rng.Item(i).Name))
                              for i in range(1, min(count, 20) + 1)],
                    "page_shapes": _n(lambda: int(page.Shapes.Count), 0)}
        if op == "effect_info":
            src = find_shape(doc, page, target or {})
            return {"op": op, "target": _s(lambda: str(src.Name)),
                    **_effect_info(src)}

        if op in {"unlock_all", "ungroup_all", "ungroup_all_ex", "ungroup_ex",
                  "break_apart_ex", "clear_effect",
                  "range_info", "range_add", "range_remove"}:
            with atomic(doc, label):
                if target:
                    try:
                        doc.Selection().RemoveAll()
                    except Exception:
                        pass
                    find_shape(doc, page, target).AddToSelection()
                sr = selection_range(doc)
                count = _n(lambda: int(sr.Count), 0) or 0
                if count < 1:
                    raise CdrBadArgument(
                        f"{op} 需要先选中对象（可用 cdr_select），或用 target 指定一个对象")
                if op == "range_info":
                    ids = [_n(lambda i=i: int(sr.Item(i).Type), -1)
                           for i in range(1, count + 1)]
                    names = [_s(lambda i=i: str(sr.Item(i).Name))
                             for i in range(1, count + 1)]
                    probe = [1, 2, 3, 4, 5, 6, 7, 8, 9]
                    found = _s(lambda: sr.FindAnyOfType(probe))
                    return {"op": op, "count": count, "names": names, "type_ids": ids,
                            "count_any_of_type": _n(
                                lambda: int(sr.CountAnyOfType(probe)), 0),
                            "exists_any_of_type": bool(_s(
                                lambda: bool(sr.ExistsAnyOfType(probe)), False)),
                            "find_any_of_type": (
                                _s(lambda: str(found.Name)) if found is not None else None),
                            "index_of_target": (
                                _n(lambda: int(sr.IndexOf(
                                    find_shape(doc, page, target))), 0)
                                if target else None)}
                if op == "clear_effect":
                    sr.ClearEffect(int(distance_a or 5))
                    return {"op": op, "cleared": True, "effect_type": int(distance_a or 5),
                            "count": count}
                if op == "range_add" or op == "range_remove":
                    if not other:
                        raise CdrBadArgument(f"{op} 需要 other（另一个对象）")
                    src = find_shape(doc, page, other)
                    extra = _app_of(doc).CreateShapeRange()
                    extra.Add(src)
                    if op == "range_add":
                        sr.AddRange(extra)
                    else:
                        sr.RemoveRange(extra)
                    return {"op": op, "other": _s(lambda: str(src.Name)),
                            "count_before": count,
                            "count_after": _n(lambda: int(sr.Count), 0) or 0}
                if op == "unlock_all":
                    sr.Unlock()
                    return {"op": op, "count": count,
                            "locked_after": [bool(_s(lambda i=i: bool(sr.Item(i).Locked),
                                                     False))
                                             for i in range(1, count + 1)]}
                if op in {"ungroup_all", "ungroup_all_ex", "ungroup_ex"}:
                    # 第五十四波纠正：三个方法**都真实生效**（之前"空操作"是测量方式错了）。
                    # UngroupAll → 无返回；UngroupAllEx / UngroupEx → 返回 ShapeRange。
                    # 选中的若是组合体内部的对象（Type≠7），改对整个页面上的组合体动手
                    sel_has_group = False
                    for _k in range(1, (_n(lambda: int(sr.Count), 0) or 0) + 1):
                        if int(_s(lambda _k=_k: int(sr.Item(_k).Type), -1)) == 7:
                            sel_has_group = True
                            break
                    if not sel_has_group:
                        groups = [sh for _i, _l, _j, sh in _iter_page_shapes(page)
                                  if int(_s(lambda sh=sh: int(sh.Type), -1)) == 7]
                        if groups:
                            sr = _app_of(doc).CreateShapeRange()
                            for sh in groups:
                                try:
                                    sr.Add(sh)
                                except Exception:
                                    continue
                    with atomic(doc, label):
                        before = _n(lambda: int(page.Shapes.Count), 0) or 0
                        made = None
                        if op == "ungroup_all":
                            sr.UngroupAll()
                        elif op == "ungroup_all_ex":
                            made = _s(lambda: sr.UngroupAllEx())
                        else:
                            made = _s(lambda: sr.UngroupEx())
                        after = _n(lambda: int(page.Shapes.Count), 0) or 0
                    return {"op": op, "page_shapes_before": before,
                            "page_shapes_after": after,
                            "unwrapped": max(after - before, 0),
                            "result_count": (_n(lambda: int(made.Count), 0)
                                             if made is not None else None)}
                if False:
                    depth = 0
                    seen: list = []
                    for _idx, _li, _j, sh in _iter_page_shapes(page):
                        guard = 0
                        while int(_s(lambda sh=sh: int(sh.Type), -1)) == 7 and guard < 20:
                            try:
                                sh.Ungroup()
                            except Exception:
                                break
                            guard += 1
                            depth += 1
                        seen.append(_s(lambda sh=sh: str(sh.Name)))
                    return {"op": op, "groups_opened": depth,
                            "page_shapes": _n(lambda: int(page.Shapes.Count), 0) or 0,
                            "names": seen[:40],
                            "note": "ShapeRange.UngroupAll 在 X8 是空操作，这里对整页"
                                    "组合体逐个 Shape.Ungroup() 递归解开"}
                # break_apart_ex
                made = _s(lambda: sr.BreakApartEx())
                return {"op": op, "count_before": count,
                        "page_shapes": _n(lambda: int(page.Shapes.Count), 0) or 0,
                        "result_count": (_n(lambda: int(made.Count), 0)
                                         if made is not None else None)}

        shape = find_shape(doc, page, target or {})

        if op == "extrude_vp_share":
            # 实测：IVGExtrudeVanishingPoint.Share(源) 传**拉伸效果对象**才会返回布尔
            # （传另一个消失点会「类型不匹配」）
            source = find_shape(doc, page, target or {})
            spec = dict(params or {})
            if not other:
                raise CdrBadArgument("extrude_vp_share 需要 other（提供源消失点的拉伸对象）")
            src_shape = find_shape(doc, page, other)
            with atomic(doc, label):
                eff = _s(lambda: source.Effects.Item(1))
                ext = _s(lambda: eff.Extrude) if eff is not None else None
                vp = _s(lambda: ext.VanishingPoint) if ext is not None else None
                src_eff = _s(lambda: src_shape.Effects.Item(1))
                src_ext = _s(lambda: src_eff.Extrude) if src_eff is not None else None
                if vp is None or src_ext is None:
                    raise CdrComFailure("源或目标没有拉伸效果")
                ok = bool(_s(lambda: bool(vp.Share(src_ext)), False))
            return {"op": op, "shared": ok,
                    "position": [_round(_s(lambda: float(vp.PositionX))),
                                 _round(_s(lambda: float(vp.PositionY)))],
                    "source": _s(lambda: str(src_shape.Name))}
        if op in {"extrude", "lens", "extrude_set", "lens_set", "clear_effects",
                  "blend", "blend_set"}:
            with atomic(doc, label):
                if op == "clear_effects":
                    effects = shape.Effects
                    cleared = 0
                    while _n(lambda: int(effects.Count), 0) > 0:
                        effects.Item(1).Clear()
                        cleared += 1
                    return {"op": op, "cleared": cleared,
                            "effects": _effect_info(shape)["count"]}
                if op == "blend":
                    if not other:
                        raise CdrBadArgument("blend 需要 target + other 两个对象")
                    other_shape = find_shape(doc, page, other)
                    shape.CreateBlend(
                        other_shape, int(steps), int(color_blend_type),
                        int(blend_mode), float(spacing), float(angle),
                        bool(loop), None, bool(rotate_shapes),
                        int(spacing_accel), int(color_accel), bool(accel_size))
                    eff = shape.Effects.Item(int(shape.Effects.Count))
                    typed = eff.Blend
                    mapping = _BLEND_PROPS
                elif op == "blend_set":
                    eff, typed = _find_effect(shape, "Blend")
                    if eff is None:
                        raise CdrBadArgument("对象没有混合效果")
                    mapping = _BLEND_PROPS
                elif op == "extrude":
                    if vp_x is None or vp_y is None:
                        bb = _bbox(shape) or {}
                        vp_x = float(bb.get("x", 0)) + float(bb.get("width", 0)) / 2
                        vp_y = float(bb.get("y", 0)) + float(bb.get("height", 0)) / 2
                    shape.CreateExtrude(
                        int(extrude_type), int(vp_type), float(vp_x), float(vp_y),
                        float(depth), int(shading), _color(doc, base_color),
                        _color(doc, shading_color), float(bevel_depth),
                        float(bevel_angle), _color(doc, bevel_color), bool(bevel_only))
                    eff = shape.Effects.Item(int(shape.Effects.Count))
                    typed = eff.Extrude
                    mapping = _EXTRUDE_PROPS
                elif op == "lens":
                    shape.CreateLens(int(lens_type), float(rate),
                                     _color(doc, color1), _color(doc, color2),
                                     int(color_map_palette))
                    eff = shape.Effects.Item(int(shape.Effects.Count))
                    typed = eff.Lens
                    mapping = _LENS_PROPS
                elif op == "extrude_set":
                    eff = shape.Effects.ExtrudeEffect
                    if eff is None:
                        raise CdrBadArgument("对象没有立体化效果")
                    typed = eff.Extrude
                    mapping = _EXTRUDE_PROPS
                else:  # lens_set
                    eff = shape.Effects.LensEffect
                    if eff is None:
                        raise CdrBadArgument("对象没有透镜效果")
                    typed = eff.Lens
                    mapping = _LENS_PROPS
                applied = _apply_effect_props(doc, typed, effect or {}, mapping)
            return {"op": op, "target": _s(lambda: str(shape.Name)),
                    "effect_type": _n(lambda: int(eff.Type)),
                    "applied": applied,
                    "effects": _effect_info(shape)}

        if op == "break_apart":
            shape.BreakApart()
            return {"op": op, "done": True}
        if op == "to_curves":
            shape.ConvertToCurves()
            return {"op": op, "name": _s(lambda: str(shape.Name)), "bbox": _bbox(shape)}
        if op == "to_symbol":
            shape.ConvertToSymbol(str(symbol_name or "Symbol 1"))
            return {"op": op, "symbol": str(symbol_name or "Symbol 1"),
                    "name": _s(lambda: str(shape.Name))}
        if op == "outline_to_curves":
            # 实测：ConvertToObject 是**原地转换**——对象本身从矩形变成曲线轮廓
            before_type = _n(lambda: int(shape.Type), -1)
            shape.Outline.ConvertToObject()
            return {"op": op, "type_before": SHAPE_TYPES.get(before_type, before_type),
                    "type_after": SHAPE_TYPES.get(_n(lambda: int(shape.Type), -1),
                                                  _n(lambda: int(shape.Type), -1)),
                    "bbox": _bbox(shape)}
        if op == "fillet":
            shape.Fillet(float(radius), True)
            return {"op": op, "radius": radius, "bbox": _bbox(shape)}
        if op == "chamfer":
            shape.Chamfer(float(distance_a), float(distance_b), True)
            return {"op": op, "bbox": _bbox(shape)}
        if op == "scallop":
            shape.Scallop(float(radius), True)
            return {"op": op, "radius": radius, "bbox": _bbox(shape)}

        # ---- 第五波：排列 / 阵列 / 边界 / 透视（2026-10-25 实测）----
        if op == "step_repeat":
            before = {str(_s(lambda s=s: s.Name, ""))
                      for _i, _l, _j, s in _iter_page_shapes(page)}
            before_count = _n(lambda: int(page.Shapes.Count), 0)
            with atomic(doc, label):
                shape.StepAndRepeat(
                    int(max(1, copies)), float(distance_x), float(distance_y),
                    int(mode_x), int(direction_x), int(mode_y), int(direction_y))
            after = [str(_s(lambda s=s: s.Name, "")) for _idx, _li, _j, s
                     in _iter_page_shapes(page)]
            created = [n for n in after if n and n not in before]
            after_count = _n(lambda: int(page.Shapes.Count), 0)
            return {"op": op, "copies": int(max(1, copies)),
                    "distance": [distance_x, distance_y],
                    "created": created, "copies_created": after_count - before_count,
                    "shapes_on_page": after_count,
                    "bbox": _bbox(shape)}
        if op == "distribute":
            with atomic(doc, label):
                shape.Distribute(int(distribute_type), bool(page_extent))
            return {"op": op, "distribute_type": int(distribute_type),
                    "page_extent": bool(page_extent), "bbox": _bbox(shape)}
        if op == "boundary":
            if x is None or y is None:
                raise CdrBadArgument("boundary 需要 x 与 y（形状边界上的参考点）")
            with atomic(doc, label):
                created = shape.CreateBoundary(float(x), float(y),
                                               bool(place_on_top), bool(delete_source))
            return {"op": op, "created": created is not None,
                    "name": _ensure_name(created) if created is not None else None,
                    "bbox": _bbox(created) if created is not None else None,
                    "shapes_on_page": _n(lambda: int(page.Shapes.Count), 0)}
        if op == "perspective":
            with atomic(doc, label):
                if None not in (hx, hy, vx, vy):
                    effect = shape.CreatePerspective(float(hx), float(hy),
                                                     float(vx), float(vy))
                else:
                    effect = shape.CreatePerspective()
            return {"op": op, "effect": effect is not None,
                    "vanish": ([hx, hy, vx, vy] if None not in (hx, hy, vx, vy) else None),
                    "bbox": _bbox(shape)}
        if op == "separate":
            with atomic(doc, label):
                shape.Separate()
            return {"op": op, "done": True,
                    "shapes_on_page": _n(lambda: int(page.Shapes.Count), 0)}

        if op == "envelope_from_shape":
            other_shape = find_shape(doc, page, other or {})
            with atomic(doc, label):
                if int(_s(lambda: int(other_shape.Type), -1)) == 3:
                    effect = shape.CreateEnvelopeFromCurve(
                        other_shape.Curve, int(envelope_mode), False, 1)
                else:
                    effect = shape.CreateEnvelopeFromShape(
                        other_shape, int(envelope_mode), False, 1)
            return {"op": op, "done": True,
                    "effect_object": type(effect).__name__ if effect is not None else None,
                    "source": _s(lambda: str(other_shape.Name)),
                    "mode": int(envelope_mode), "bbox": _bbox(shape)}
        if op == "compare":
            other_shape = find_shape(doc, page, other or {})
            same = bool(_s(lambda: shape.CompareTo(other_shape, int(compare_type),
                                                   int(condition)), False))
            return {"op": op, "same": same, "compare_type": int(compare_type),
                    "condition": int(condition),
                    "other": _s(lambda: str(other_shape.Name))}

        if op in {"project", "unproject"}:
            pl = int(plane)
            if pl not in (0, 1, 2):
                raise CdrBadArgument("plane 可用 0=front / 1=right / 2=top")
            with atomic(doc, label):
                if op == "project":
                    shape.Project(pl, int(reference_point), bool(apply_to_duplicate))
                else:
                    shape.Unproject(pl, int(reference_point), bool(apply_to_duplicate))
            return {"op": op, "plane": pl, "reference_point": int(reference_point),
                    "apply_to_duplicate": bool(apply_to_duplicate),
                    "bbox": _bbox(shape)}

        if op in {"parallel_curves", "rasterize", "batch_outline_to_curves",
                  "fill_mode", "pixel_aligned", "range_boundary"}:
            # 范围级操作：显式给了 target 就只处理它（先清空选择集，避免混入
            # cdr_create 留下的"最后创建对象"选中态）；没给就用当前选择集。
            if target:
                doc.ClearSelection()
                shape.AddToSelection()
            sr = _s(lambda: selection_range(doc))
            if sr is None or _n(lambda: int(sr.Count), 0) < 1:
                shape.AddToSelection()
                sr = selection_range(doc)
            if sr is None or _n(lambda: int(sr.Count), 0) < 1:
                raise CdrBadArgument(f"{op} 需要至少 1 个对象（先 cdr_select 或给 target）")
            if op == "batch_outline_to_curves":
                with atomic(doc, label):
                    made = sr.ConvertOutlineToObject()
                return {"op": op, "done": made is not None,
                        "shapes_on_page": _n(lambda: int(page.Shapes.Count), 0)}
            if op == "fill_mode":
                with atomic(doc, label):
                    sr.SetFillMode(int(fill_mode))
                return {"op": op, "fill_mode": int(fill_mode)}
            if op == "pixel_aligned":
                with atomic(doc, label):
                    sr.SetPixelAlignedRendering(bool(pixel_aligned))
                return {"op": op, "pixel_aligned": bool(pixel_aligned)}
            if op == "range_boundary":
                if x is None or y is None:
                    raise CdrBadArgument("range_boundary 需要 x 与 y（边界参考点）")
                with atomic(doc, label):
                    made = sr.CreateBoundary(float(x), float(y), True, False)
                return {"op": op, "created": made is not None,
                        "name": _ensure_name(made) if made is not None else None,
                        "bbox": _bbox(made) if made is not None else None,
                        "shapes_on_page": _n(lambda: int(page.Shapes.Count), 0)}
            if op == "parallel_curves":
                try:
                    with atomic(doc, label):
                        made = sr.CreateParallelCurves(int(max(1, curve_count)),
                                                       float(curve_distance))
                except CdrComFailure as exc:
                    raise CdrBadArgument(
                        "parallel_curves 只支持曲线对象；请先 "
                        f"cdr_shape_ops(op='to_curves') 转曲（原始错误：{exc}）")
                return {"op": op, "count": int(max(1, curve_count)),
                        "distance": float(curve_distance),
                        "created": _n(lambda: int(made.Count), 0) if made is not None else 0,
                        "shapes_on_page": _n(lambda: int(page.Shapes.Count), 0)}
            with atomic(doc, label):
                made = sr.ConvertToBitmap(24, bool(grayscale), True, bool(transparent),
                                          int(resolution))
            return {"op": op, "resolution": int(resolution),
                    "grayscale": bool(grayscale), "transparent": bool(transparent),
                    "created": made is not None,
                    "name": _s(lambda: str(made.Name)) if made is not None else None,
                    "type_id": _n(lambda: int(made.Type), -1) if made is not None else None,
                    "bbox": _bbox(made) if made is not None else None,
                    "shapes_on_page": _n(lambda: int(page.Shapes.Count), 0)}

        # 需要第二个形状
        target_shape = find_shape(doc, page, other or {})
        if op == "weld":
            shape.Weld(target_shape, bool(leave_source), bool(leave_target))
        elif op == "trim":
            shape.Trim(target_shape, bool(leave_source), bool(leave_target))
        elif op == "intersect":
            shape.Intersect(target_shape, bool(leave_source), bool(leave_target))
        elif op == "powerclip":
            shape.AddToPowerClip(target_shape, bool(center_in_container))
        return {"op": op, "done": True}


def batch_modify(doc: Any, page_index: int | None, selector: dict, changes: dict,
                 label: str = "cdr_modify_all") -> dict:
    """对一批对象套用同样的改动。

    selector 支持：
      {"all": true}                 本页所有对象
      {"type_id": 1}                按类型（1矩形 2椭圆 3曲线 4多边形 6文字 7组合体）
      {"name": "cdr_xxx"}           按名字
      {"selection": true}           当前选择集
    整批合成**一个 undo 步**；单个对象失败不会中断整批，失败原因会汇总返回。
    """
    if not isinstance(selector, dict) or not selector:
        raise CdrBadArgument("selector 不能为空")
    matched: list[Any] = []
    errors: list[str] = []
    with mm_units(doc):
        page = _page(doc, page_index)
        if selector.get("selection"):
            sel = doc.Selection().Shapes
            for i in range(1, int(sel.Count) + 1):
                matched.append(sel.Item(i))
        else:
            want_type = selector.get("type_id")
            want_name = selector.get("name")
            for _idx, _li, _j, shape in _iter_page_shapes(page):
                if want_type and int(_s(lambda s=shape: int(s.Type), -1)) != int(want_type):
                    continue
                if want_name and str(_s(lambda s=shape: str(s.Name), "")) != str(want_name):
                    continue
                matched.append(shape)
        if not matched:
            raise CdrBadArgument("没有匹配到任何对象")

        done = 0
        with atomic(doc, label):
            for i, shape in enumerate(matched, 1):
                try:
                    _apply_changes(doc, shape, changes)
                    done += 1
                except Exception as exc:  # noqa: BLE001
                    if len(errors) < 5:
                        errors.append(f"#{i}: {exc}")
        return {"matched": len(matched), "modified": done, "errors": errors,
                "changes": sorted(changes)}


def align_shapes(doc: Any, page_index: int | None = None, horizontal: str = "none",
                 vertical: str = "none", to: str = "selection", name: str | None = None,
                 type_id: int | None = None, rect: dict | None = None,
                 spacing: bool = False, x: float | None = None, y: float | None = None,
                 target2: dict | None = None) -> dict:
    """对齐/分布。horizontal: none/left/center/right；vertical: none/bottom/middle/top。

    `to` 除了 `selection` / `page`，还支持 X8 原生的**范围对齐**（第三十三波实测）：
      * `page_center`  AlignRangeToPageCenter
      * `grid`         AlignRangeToGrid
      * `point`        AlignRangeToPoint（要 x / y）
      * `shape`        AlignRangeToShape（要 target2）
      * `range`        AlignRangeToShapeRange（要 target2，用第二个对象建单元素范围）
    这些原生方法一次只吃一个轴（`cdrAlignLeft=1 / Right=2 / HCenter=3 / Top=4 /
    Bottom=8 / VCenter=12`），所以 horizontal/vertical 会被拆成两次调用。
    """
    h_map = {"none": 0, "left": 1, "center": 2, "middle": 2, "right": 3,
             "space": 4, "distribute": 4}
    v_map = {"none": 0, "bottom": 1, "center": 2, "middle": 2, "top": 3,
             "space": 4, "distribute": 4}
    h = h_map.get(str(horizontal).lower(), -1)
    v = v_map.get(str(vertical).lower(), -1)
    if h < 0 or v < 0:
        raise CdrBadArgument("horizontal 取值 none/left/center/right/space；"
                             "vertical 取值 none/bottom/middle/top/space")
    if h == 0 and v == 0:
        raise CdrBadArgument("horizontal 与 vertical 不能同时为 none")
    native = str(to).lower()
    if native not in {"selection", "page", "page_edge", "page_center", "grid", "point",
                      "shape", "range"}:
        raise CdrBadArgument(
            "to 取值 selection/page/page_center/grid/point/shape/range")
    if native == "point" and (x is None or y is None):
        raise CdrBadArgument("to=point 需要 x / y")
    if native in {"shape", "range"} and not target2:
        raise CdrBadArgument(f"to={native} 需要 target2（参照对象）")

    # 原生范围对齐的轴枚举
    native_axis = []
    if h:
        native_axis.append({1: 1, 2: 3, 3: 2, 4: 1}[h])  # left/HCenter/right
    if v:
        native_axis.append({1: 8, 2: 12, 3: 4}[v])       # bottom/VCenter/top

    with mm_units(doc):
        select_shapes(doc, page_index, name=name, type_id=type_id, rect=rect)
        sr = selection_range(doc)
        if int(sr.Count) < 1:
            raise CdrBadArgument("没有选中任何对象")
        if native in {"selection", "page"}:
            with atomic(doc, "cdr_align"):
                sr.AlignAndDistribute(
                    h, v,
                    1 if native == "page" else 0,
                    0, False, 0, 0.0, 0.0, None,
                )
            return {"aligned": int(sr.Count), "horizontal": h, "vertical": v, "to": native,
                    "native": False}

        page = _page(doc, page_index)
        ref_shape = find_shape(doc, page, target2) if target2 else None
        with atomic(doc, "cdr_align"):
            for axis in native_axis:
                if native == "page_edge":
                    # 实测：ShapeRange.AlignRangeToPage(cdrAlignType) 直接对齐到页边
                    sr.AlignRangeToPage(axis)
                elif native == "page_center":
                    sr.AlignRangeToPageCenter(axis)
                elif native == "grid":
                    sr.AlignRangeToGrid(axis)
                elif native == "point":
                    sr.AlignRangeToPoint(axis, float(x), float(y))
                elif native == "shape":
                    sr.AlignRangeToShape(axis, ref_shape)
                else:  # range
                    ref_range = _app_of(doc).CreateShapeRange()
                    ref_range.Add(ref_shape)
                    sr.AlignRangeToShapeRange(axis, ref_range)
        return {"aligned": int(sr.Count), "horizontal": h, "vertical": v, "to": native,
                "native": True, "axes": native_axis,
                "point": [float(x), float(y)] if native == "point" else None,
                "reference": _s(lambda: str(ref_shape.Name)) if ref_shape is not None else None}


# --------------------------------------------------------------------------- #
# 文字
# --------------------------------------------------------------------------- #
CASE_MODES = {"sentence": 0, "lower": 1, "lowercase": 1, "upper": 2, "uppercase": 2,
              "title": 3, "toggle": 4}


def _case_index(value: Any) -> int:
    if isinstance(value, int):
        return value
    key = str(value).lower()
    if key not in CASE_MODES:
        raise CdrBadArgument(f"未知大小写模式 {value!r}；可用 sentence/lower/upper/title/toggle")
    return CASE_MODES[key]


def _text_structs(shape: Any) -> dict:
    """读 `Text.FontProperties / SpaceProperties / AlignProperties / HyphenationSettings`。

    实测：这四个在 pywin32 里是**属性**（返回 IVGStruct* 对象），带 frames 的 propget
    调不成方法（`TypeError: object is not callable`）；直接赋值 setter
    （`text.FontProperties = struct`）会把字号写成 1.17e-38 之类的垃圾值 —— 所以**只读不写**。
    """
    text = _s(lambda: shape.Text)
    if text is None:
        return {}
    out: dict = {}
    fp = _s(lambda: text.FontProperties)
    if fp is not None:
        out["font"] = {
            "name": _s(lambda: str(fp.Name)),
            "size": _round(_s(lambda: float(fp.Size))),
            "style": _n(lambda: int(fp.Style)),
            "underline": _n(lambda: int(fp.Underline)),
            "overscore": _n(lambda: int(fp.Overscore)),
            "strikethru": _n(lambda: int(fp.Strikethru)),
            "uppercase": _n(lambda: int(fp.Uppercase)),
            "position": _n(lambda: int(fp.Position)),
            "range_kerning": _round(_s(lambda: float(fp.RangeKerning))),
        }
    sp = _s(lambda: text.SpaceProperties)
    if sp is not None:
        out["space"] = {
            "character_spacing": _round(_s(lambda: float(sp.CharacterSpacing))),
            "word_spacing": _round(_s(lambda: float(sp.WordSpacing))),
            "line_spacing": _round(_s(lambda: float(sp.LineSpacing))),
            "line_spacing_type": _n(lambda: int(sp.LineSpacingType)),
            "before_paragraph": _round(_s(lambda: float(sp.BeforeParagraphSpacing))),
            "after_paragraph": _round(_s(lambda: float(sp.AfterParagraphSpacing))),
        }
    ap = _s(lambda: text.AlignProperties)
    if ap is not None:
        out["align"] = {
            "alignment": _n(lambda: int(ap.Alignment)),
            "left_indent": _round(_s(lambda: float(ap.LeftIndent))),
            "right_indent": _round(_s(lambda: float(ap.RightIndent))),
            "first_line_indent": _round(_s(lambda: float(ap.FirstLineIndent))),
            "max_word_spacing": _round(_s(lambda: float(ap.MaxWordSpacing))),
            "min_word_spacing": _round(_s(lambda: float(ap.MinWordSpacing))),
            "max_char_spacing": _round(_s(lambda: float(ap.MaxCharacterSpacing))),
            "char_rotation": _round(_s(lambda: float(ap.CharacterRotation))),
        }
    hp = _s(lambda: text.HyphenationSettings)
    if hp is not None:
        out["hyphenation"] = {
            "use_automatic": bool(_s(lambda: bool(hp.UseAutomaticHyphenation), False)),
            "break_capitalized": bool(_s(lambda: bool(hp.BreakCapitalized), False)),
            "hot_zone": _round(_s(lambda: float(hp.HotZone))),
            "min_word_length": _n(lambda: int(hp.MinWordLength)),
            "min_before": _n(lambda: int(hp.MinCharactersBefore)),
            "min_after": _n(lambda: int(hp.MinCharactersAfter)),
            "break_all_cap_words": bool(_s(lambda: bool(hp.BreakAllCapWords), False)),
        }
    return out


def _make_point_range(app: Any, points: list) -> Any:
    """把 [[x,y],…] 变成 IVGPointRange：先借用一条临时线的子路径取空范围再接 AddPointXY。"""
    pts = [[float(p[0]), float(p[1])] for p in points]
    if len(pts) < 2:
        raise CdrBadArgument("需要至少 2 个点")
    probe = app.ActiveDocument.ActiveLayer.CreateLineSegment(
        pts[0][0], pts[0][1], pts[1][0], pts[1][1])
    sub = probe.Curve.SubPaths.Item(1)
    rng = sub.GetEvenlySpacedPoints(1.0, True)
    # 清空后逐个补点
    try:
        rng.RemoveAll()
    except Exception:
        pass
    for point in pts:
        rng.AddPointXY(point[0], point[1])
    try:
        probe.Delete()
    except Exception:
        pass
    return rng


def _text_struct_dict(struct: Any, attrs: tuple) -> dict | None:
    """把 IVGStruct* 快照转成只保留白名单字段的 dict（字段读失败就跳过）。"""
    if struct is None:
        return None
    out: dict = {}
    for attr in attrs:
        try:
            val = getattr(struct, attr)
        except Exception:
            continue
        if isinstance(val, (int, float)):
            out[attr] = _round(float(val))
        else:
            out[attr] = str(val)[:80]
    return out


def _text_range_action(story: Any, spec: dict) -> dict:
    """TextRange 的宽接口 / 游标操作（第三十七波实测）。

    op: collapse / set_range / in_range / is_same / insert_before / insert_after /
        replace / set_line_spacing
    """
    op = str(spec.get("op", "")).lower()
    start = int(spec.get("start", 0))
    count = int(spec.get("count", 0))
    end = int(spec.get("end", start + max(count, 1)))
    rng = story.Range(start, end)
    if op == "collapse":
        rng.Collapse(bool(spec.get("to_end", True)))
        return {"range": [int(_s(lambda: int(rng.Start), 0)),
                          int(_s(lambda: int(rng.End), 0))]}
    if op == "set_range":
        rng.SetRange(int(spec.get("new_start", start)), int(spec.get("new_end", end)))
        return {"range": [int(_s(lambda: int(rng.Start), 0)),
                          int(_s(lambda: int(rng.End), 0))]}
    if op in ("in_range", "is_same"):
        other = story.Range(int(spec.get("other_start", start)),
                            int(spec.get("other_end", end)))
        if op == "in_range":
            inside = bool(_s(lambda: bool(rng.InRange(other)), False))
            return {"in_range": inside}
        same = bool(_s(lambda: bool(rng.IsSame(other)), False))
        return {"same": same}
    if op in ("insert_before", "insert_after", "replace"):
        text = str(spec.get("text", ""))
        lang = int(spec.get("language_id", 0))
        charset = int(spec.get("charset", 0))
        fnt = str(spec.get("font", ""))
        if op == "insert_before":
            made = rng.InsertBeforeWide(text, lang, charset, fnt)
        elif op == "insert_after":
            made = rng.InsertAfterWide(text, lang, charset, fnt)
        else:
            made = rng.ReplaceWide(text, lang, charset, fnt)
        return {"op": op, "text": text,
                "range": [_n(lambda: int(made.Start)) if made is not None else None,
                          _n(lambda: int(made.End)) if made is not None else None],
                "contents": _s(lambda: str(story.Text))}
    if op == "set_line_spacing":
        rng.SetLineSpacing(int(spec.get("type", 1)),
                           float(spec.get("line_spacing", 100.0)),
                           float(spec.get("before", 0.0)),
                           float(spec.get("after", 0.0)))
        return {"type": int(spec.get("type", 1)),
                "line_spacing": float(spec.get("line_spacing", 100.0))}
    raise CdrBadArgument(
        "range_action.op 可用：collapse / set_range / in_range / is_same / "
        "insert_before / insert_after / replace / set_line_spacing")


def _text_font_info(shape: Any) -> dict | None:
    """读字体属性。

    实测坑：`Text.GetFontProperties(Frames)` 在这些对象上会返回**默认结构体**
    （Name=''、Size=1.17e-38），不可信；`Text.Story` 上的 Font/Size/Bold/Italic/
    Underline 才是真实值。
    """
    story = _s(lambda: shape.Text.Story)
    if story is None:
        return None
    return {
        "name": _s(lambda: str(story.Font)) or None,
        "size": _round(_s(lambda: float(story.Size))),
        "bold": bool(_s(lambda: bool(story.Bold), False)),
        "italic": bool(_s(lambda: bool(story.Italic), False)),
        "underline": bool(_s(lambda: bool(story.Underline), False)),
    }


def _text_hyphenation_info(shape: Any) -> dict | None:
    """读断字设置。实测 `SetHyphenationSettings` 写回不生效，这里只读。"""
    hs = _s(lambda: shape.Text.GetHyphenationSettings(2))
    if hs is None:
        return None
    return {
        "automatic": bool(_s(lambda: bool(hs.UseAutomaticHyphenation), False)),
        "min_word_length": _n(lambda: int(hs.MinWordLength)),
        "min_before": _n(lambda: int(hs.MinCharactersBefore)),
        "min_after": _n(lambda: int(hs.MinCharactersAfter)),
        "hot_zone": _round(_s(lambda: float(hs.HotZone))),
        "break_capitalized": bool(_s(lambda: bool(hs.BreakCapitalized), False)),
        "break_all_cap_words": bool(_s(lambda: bool(hs.BreakAllCapWords), False)),
    }


def _text_runs(shape: Any) -> list[dict]:
    """列出文本的格式段（EnumRanges(-1)）：文字/字体/字号/粗斜体/区间。"""
    ranges = _s(lambda: shape.Text.Story.EnumRanges(-1))
    if ranges is None:
        return []
    count = _n(lambda: int(ranges.Count), 0)
    out: list[dict] = []
    for i in range(1, min(count, 300) + 1):
        r = _s(lambda i=i: ranges.Item(i))
        if r is None:
            continue
        out.append({
            "text": _s(lambda r=r: str(r.Text)),
            "font": _s(lambda r=r: str(r.Font)) or None,
            "size": _round(_s(lambda r=r: float(r.Size))),
            "bold": bool(_s(lambda r=r: bool(r.Bold), False)),
            "italic": bool(_s(lambda r=r: bool(r.Italic), False)),
            "underline": bool(_s(lambda r=r: bool(r.Underline), False)),
            "start": _n(lambda r=r: int(r.Start)),
            "end": _n(lambda r=r: int(r.End)),
        })
    return out


def set_text(doc: Any, page_index: int | None, target: dict, *, text: str | None = None,
             font: str | None = None, size: float | None = None, bold: bool | None = None,
             italic: bool | None = None, underline: bool | None = None,
             alignment: str | None = None, to_curves: bool = False,
             fit_to_path: dict | None = None, fit_frame: bool = False,
             change_case: Any = None, bullet: Any = None, drop_cap: Any = None,
             no_effect: bool = False, line_spacing: float | None = None,
             char_spacing: float | None = None, word_spacing: float | None = None,
             indent_left: float | None = None, indent_right: float | None = None,
             indent_first_line: float | None = None,
             clear_tabs: bool = False, tab_stops: list | None = None,
             insert_before: str | None = None, insert_after: str | None = None,
             to_paragraph: bool = False, to_artistic: bool = False,
             contents: str | None = None, find: str | None = None,
             replace_old: str | None = None, replace_new: str | None = None,
             select_all: bool = False, align_to_baseline: bool = False,
             copy_attributes_from: dict | None = None,
             read_contents: bool = False, range_start: int | None = None,
             range_end: int | None = None, export_path: str | None = None,
             export_start: int = 1, export_count: int = 0,
             opentype: dict | None = None,
             read_font: bool = False, read_hyphenation: bool = False,
             range_delete: bool = False, range_copy: bool = False,
             range_paste: bool = False, range_straighten: bool = False,
             find_query: str | None = None, runs: bool = False,
             begin_edit: bool = False, paragraph_format: dict | None = None,
             range_struct: dict | None = None,
             link_to: dict | None = None, unlink: bool = False,
             make_html_compatible: bool | None = None,
             frame_info: bool = False,
             struct_info: bool = False,
             range_action: dict | None = None,
             label: str = "cdr_text") -> dict:
    with mm_units(doc):
        page = _page(doc, page_index)
        shape = find_shape(doc, page, target)
        if int(_s(lambda: int(shape.Type), -1)) != TEXT_TYPE:
            raise CdrBadArgument("目标不是文字对象")
        mutating = any(v is not None for v in (
            text, font, size, bold, italic, underline, alignment,
            fit_to_path, change_case, bullet, drop_cap, line_spacing,
            char_spacing, word_spacing, indent_left, indent_right,
            indent_first_line, insert_before, insert_after,
            contents, find, replace_old, replace_new, copy_attributes_from,
            opentype,
        )) or any((to_curves, fit_frame, no_effect, to_paragraph, to_artistic,
                   select_all, align_to_baseline, range_delete, range_copy,
                   range_paste, range_straighten, clear_tabs)) or bool(tab_stops) \
            or begin_edit or paragraph_format is not None \
            or link_to is not None or unlink or make_html_compatible is not None \
            or frame_info
        if not mutating:
            # 纯读取 / 导出：不建 undo 组，也不改动文档
            result: dict = {"applied": [], "warnings": [],
                            "text": _s(lambda: str(shape.Text.Story.Text), ""),
                            "bbox": _bbox(shape)}
            if range_struct is not None:
                # 实测：Text.FontPropertiesInRange / SpacePropertiesInRange /
                # AlignPropertiesInRange / HyphenationSettingsInRange(StartIndex, Count,
                # IndexingType) 都能返回结构体（注意方法名**没有** Get 前缀）
                spec = dict(range_struct or {})
                # 注意：StartIndex 是 1 基；传 0 会返回一堆垃圾值（Size=0/Style=16）
                start = int(spec.get("start", 1))
                count = int(spec.get("count", 5))
                indexing = int(spec.get("indexing", 0))
                text_obj = shape.Text
                result["range_struct"] = {
                    "start": start, "count": count, "indexing": indexing,
                    "font": _text_struct_dict(
                        _s(lambda: text_obj.FontPropertiesInRange(start, count, indexing)),
                        ("Name", "Size", "Style", "Underline")),
                    "space": _text_struct_dict(
                        _s(lambda: text_obj.SpacePropertiesInRange(start, count, indexing)),
                        ("CharacterSpacing", "WordSpacing", "LineSpacing")),
                    "align": _text_struct_dict(
                        _s(lambda: text_obj.AlignPropertiesInRange(start, count, indexing)),
                        ("Alignment", "LeftIndent", "RightIndent", "FirstIndent")),
                    "hyphenation": _text_struct_dict(
                        _s(lambda: text_obj.HyphenationSettingsInRange(start, count,
                                                                      indexing)),
                        ("UseAutomaticHyphenation", "HotZone")),
                }
            if read_contents:
                result["contents"] = _s(lambda: shape.Text.GetContents(2), "")
            if range_start is not None or range_end is not None:
                a = int(range_start or 0)
                b = int(range_end if range_end is not None else a + 1)
                if not 0 <= a < b:
                    result["warnings"].append(
                        "range_start/range_end 需满足 0 <= start < end（实测 0 基）")
                else:
                    result["range"] = [a, b]
                    result["range_text"] = _s(
                        lambda: shape.Text.Range(a, b).Text, "")
            if export_path:
                shape.Text.ExportToFile(str(export_path), int(export_start),
                                        int(export_count), 0)
                result["applied"].append("export_text")
                result["exported"] = str(export_path)
            if find_query:
                try:
                    ranges = shape.Text.Story.FindRanges(str(find_query))
                    count = _n(lambda: int(ranges.Count), 0)
                    result["find_ranges"] = {
                        "query": str(find_query), "count": count,
                        "texts": [_s(lambda i=i: str(ranges.Item(i).Text))
                                  for i in range(1, min(count, 20) + 1)]}
                except Exception as exc:  # noqa: BLE001
                    result["find_ranges"] = {"query": str(find_query),
                                             "error": str(exc)}
            if runs:
                result["runs"] = _text_runs(shape)
            if read_font:
                info = _text_font_info(shape)
                if info is not None:
                    result["font"] = info
            if read_hyphenation:
                info = _text_hyphenation_info(shape)
                if info is not None:
                    result["hyphenation"] = info
            if struct_info:
                result["structs"] = _text_structs(shape)
            if range_action:
                try:
                    result["range_action"] = _text_range_action(
                        shape.Text.Story, dict(range_action))
                    result["applied"].append("range_action")
                except CdrBadArgument:
                    raise
                except Exception as exc:  # noqa: BLE001
                    result["warnings"].append(f"range_action 失败：{exc}")
            story = _s(lambda: shape.Text.Story)
            if story is not None:
                result["indent"] = {
                    "left": _s(lambda: float(story.LeftIndent)),
                    "right": _s(lambda: float(story.RightIndent)),
                    "first_line": _s(lambda: float(story.FirstLineIndent)),
                }
                try:
                    tabs = story.Tabs
                    count = _n(lambda: int(tabs.Count), 0)
                    result["tabs"] = [
                        {"position": _s(lambda i=i: float(tabs.Item(i).Position)),
                         "alignment": _n(lambda i=i: int(tabs.Item(i).Alignment)),
                         "leadered": bool(_s(lambda i=i: bool(tabs.Item(i).Leadered), False))}
                        for i in range(1, min(count, 100) + 1)]
                except Exception:
                    result["tabs"] = []
            return result
        with atomic(doc, label):
            story = shape.Text.Story
            applied: list[str] = []
            warnings: list[str] = []

            # ---- 文本类型转换（实测：艺术字↔段落字互转可用）----
            if to_paragraph:
                try:
                    shape.Text.ConvertToParagraph()
                    applied.append("to_paragraph")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"转段落文本失败：{exc}")
            if to_artistic:
                try:
                    shape.Text.ConvertToArtistic()
                    applied.append("to_artistic")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"转艺术字失败：{exc}")

            # ---- 段落文本内容（实测：SetContents/GetContents/Replace 仅段落文本可用）----
            if contents is not None:
                try:
                    shape.Text.SetContents(0, str(contents))
                    applied.append("contents")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"设置文本内容失败（需段落文本）：{exc}")
            if replace_old is not None:
                try:
                    shape.Text.Replace(str(replace_old), str(replace_new or ""),
                                       True, 1, True, True, 0)
                    applied.append("replace")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"文本替换失败（需段落文本）：{exc}")
            if find is not None:
                try:
                    warnings.append(f"find:{find} -> 位置 {shape.Text.Find(str(find), True, 1, True, 0)}")
                    applied.append("find")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"查找失败：{exc}")

            # ---- 选区/基线与属性复制 ----
            if select_all:
                try:
                    story.Select()
                    applied.append("select_all")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"全选文本失败：{exc}")
            if align_to_baseline:
                try:
                    story.AlignToBaseline()
                    applied.append("align_to_baseline")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"对齐基线失败：{exc}")
            if copy_attributes_from:
                try:
                    src = find_shape(doc, _page(doc, page_index), copy_attributes_from)
                    story.CopyAttributes(src.Text.Story)
                    applied.append("copy_attributes")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"复制文字属性失败：{exc}")

            if frame_info:
                # 实测：Text.Frames.Item(i) 是 IVGTextFrame，
                # ColumnWidth(i) / ColumnGutter(i) 按列号读（SetColumns 需要 SAFEARRAY，暂不可用）
                frames = _s(lambda: shape.Text.Frames)
                if frames is None:
                    raise CdrComFailure("取不到 Text.Frames")
                count = _n(lambda: int(frames.Count), 0) or 0
                info = []
                for i in range(1, min(count, 10) + 1):
                    tf = _s(lambda i=i: frames.Item(i))
                    if tf is None:
                        continue
                    cols = _n(lambda tf=tf: int(tf.ColumnCount), 0) or 0
                    info.append({
                        "index": _n(lambda tf=tf: int(tf.Index)),
                        "is_first": bool(_s(lambda tf=tf: bool(tf.IsFirst), False)),
                        "is_last": bool(_s(lambda tf=tf: bool(tf.IsLast), False)),
                        "empty": bool(_s(lambda tf=tf: bool(tf.Empty), False)),
                        "columns": cols,
                        "vertical_alignment": _n(
                            lambda tf=tf: int(tf.VerticalAlignment)),
                        "column_widths": [
                            _round(_s(lambda tf=tf, k=k: float(tf.ColumnWidth(k))))
                            for k in range(1, min(cols, 8) + 1)],
                        "column_gutters": [
                            _round(_s(lambda tf=tf, k=k: float(tf.ColumnGutter(k))))
                            for k in range(1, min(cols, 8) + 1)],
                    })
                return {"applied": ["frame_info"], "frames": info,
                        "frames_in_link": _n(lambda: int(shape.Text.FramesInLink), 0),
                        "unused_frames": _n(lambda: int(shape.Text.UnusedFramesInLink), 0)}
            if make_html_compatible is not None:
                # 实测：Text.MakeHTMLCompatible(bool) 返回 True（对艺术字/段落字都可调）
                ok = bool(_s(lambda: bool(shape.Text.MakeHTMLCompatible(
                    bool(make_html_compatible))), False))
                applied.append("make_html_compatible")
                return {"applied": applied, "warnings": warnings,
                        "html_compatible": ok,
                        "requested": bool(make_html_compatible)}
            if link_to is not None or unlink:
                # 实测：Text.Frames.Item(1).LinkTo(另一个文本形状) 可把段落框串起来
                frames = _s(lambda: shape.Text.Frames)
                tf = _s(lambda: frames.Item(1)) if frames is not None else None
                if tf is None:
                    raise CdrComFailure("取不到文本框架（Text.Frames）")
                with atomic(doc, label):
                    if unlink:
                        tf.UnLink()
                        applied.append("unlink")
                    else:
                        other_shape = find_shape(doc, _page(doc, page_index), link_to)
                        tf.LinkTo(other_shape)
                        applied.append("link_to")
                return {"applied": applied, "warnings": warnings,
                        "frames_in_link": _n(lambda: int(shape.Text.FramesInLink), 0),
                        "frame_index": _n(lambda: int(tf.Index)),
                        "column_count": _n(lambda: int(tf.ColumnCount))}
            if paragraph_format:
                # 实测：Application.CreateStructAlignProperties / SpaceProperties /
                # HyphenationSettings 造出的结构体可直接喂给 Set*Properties(0, 结构体)，
                # **但只对段落文本生效**（艺术字会报「该参数需要段落文本」）
                spec = dict(paragraph_format)
                app = _app_of(doc)
                range_start = spec.get("range_start")
                range_count = int(spec.get("range_count", 5))
                indexing = int(spec.get("indexing", 0))
                if any(k in spec for k in ("alignment", "left_indent", "right_indent",
                                           "first_indent", "vertical_alignment")):
                    ap = app.CreateStructAlignProperties()
                    for key, attr in (("alignment", "Alignment"),
                                      ("left_indent", "LeftIndent"),
                                      ("right_indent", "RightIndent"),
                                      ("first_indent", "FirstIndent"),
                                      ("vertical_alignment", "VerticalAlignment")):
                        if key in spec:
                            try:
                                setattr(ap, attr, spec[key])
                            except Exception:
                                pass
                    if range_start is not None:
                        shape.Text.SetAlignPropertiesInRange(int(range_start),
                                                             range_count, indexing, ap)
                        applied.append("align_struct_in_range")
                    else:
                        shape.Text.SetAlignProperties(0, ap)
                        applied.append("align_struct")
                if any(k in spec for k in ("character_spacing", "word_spacing",
                                           "line_spacing")):
                    sp = app.CreateStructSpaceProperties()
                    for key, attr in (("character_spacing", "CharacterSpacing"),
                                      ("word_spacing", "WordSpacing"),
                                      ("line_spacing", "LineSpacing")):
                        if key in spec:
                            try:
                                setattr(sp, attr, float(spec[key]))
                            except Exception:
                                pass
                    if range_start is not None:
                        # 实测：*InRange 的 StartIndex 是 **1 基**（0 会报越界）
                        shape.Text.SetSpacePropertiesInRange(int(range_start),
                                                             range_count, indexing, sp)
                        applied.append("space_struct_in_range")
                    else:
                        shape.Text.SetSpaceProperties(0, sp)
                        applied.append("space_struct")
                if spec.get("report_font_struct"):
                    # 实测：Application.CreateStructFontProperties() 可用；读回的是默认值
                    fs = app.CreateStructFontProperties()
                    applied.append("font_struct_factory")
                    warnings.append(
                        "CreateStructFontProperties 造出的结构体可读（Name/Size/Style/"
                        "Underline），但 SetFontProperties(frames, 结构体) 对可见文本无效")
                    try:
                        warnings.append(f"默认字体结构体：Name={str(fs.Name)[:40]} "
                                        f"Size={float(fs.Size)}")
                    except Exception:
                        pass
                if any(k in spec for k in ("hyphenation", "hot_zone")):
                    hp = app.CreateStructHyphenationSettings()
                    if "hyphenation" in spec:
                        try:
                            hp.UseAutomaticHyphenation = int(bool(spec["hyphenation"]))
                        except Exception:
                            pass
                    if "hot_zone" in spec:
                        try:
                            hp.HotZone = float(spec["hot_zone"])
                        except Exception:
                            pass
                    if range_start is not None:
                        shape.Text.SetHyphenationSettingsInRange(
                            int(range_start), range_count, indexing, hp)
                        applied.append("hyphenation_struct_in_range")
                    else:
                        shape.Text.SetHyphenationSettings(0, hp)
                        applied.append("hyphenation_struct")
            if begin_edit:
                try:
                    shape.Text.BeginEdit()
                    applied.append("begin_edit")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"BeginEdit 失败：{exc}")
            if text is not None:
                story.Text = str(text)
                applied.append("text")
            tri = {"bold", "italic", "underline"}
            for key, prop, val in (("font", "Font", font), ("size", "Size", size),
                                   ("bold", "Bold", bold), ("italic", "Italic", italic),
                                   ("underline", "Underline", underline)):
                if val is not None:
                    try:
                        setattr(story, prop,
                                _underline(val) if key == "underline"
                                else (_tristate(val) if key in tri else val))
                        applied.append(key)
                    except Exception as exc:  # noqa: BLE001
                        warnings.append(f"{key} 设置失败：{exc}")
            if alignment is not None:
                try:
                    _set_text_align(shape, alignment)
                    applied.append("alignment")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"对齐设置失败（X8 该文字对象不支持）：{exc}")
            if change_case is not None:
                try:
                    story.ChangeCase(_case_index(change_case))
                    applied.append("change_case")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"大小写转换失败：{exc}")
            if bullet is not None:
                try:
                    sym = bullet if isinstance(bullet, str) else "\u2022"
                    story.ApplyBulletEffect(sym, "", 0, 0, 0, 0)
                    applied.append("bullet")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"项目符号失败（需段落文字）：{exc}")
            if drop_cap:
                try:
                    lines = int(drop_cap) if not isinstance(drop_cap, bool) else 2
                    story.ApplyDropCapEffect(lines, 0, 0)
                    applied.append("drop_cap")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"首字下沉失败（需段落文字）：{exc}")
            if no_effect:
                try:
                    story.ApplyNoEffect()
                    applied.append("no_effect")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"清除文字效果失败：{exc}")
            if insert_before:
                try:
                    story.InsertBefore(str(insert_before), 0, 0, "")
                    applied.append("insert_before")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"前置插入失败：{exc}")
            if insert_after:
                try:
                    story.InsertAfter(str(insert_after), 0, 0, "")
                    applied.append("insert_after")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"后置插入失败：{exc}")
            if any(v is not None for v in (line_spacing, char_spacing, word_spacing)):
                try:
                    props = shape.Text.GetSpaceProperties(0)
                    if line_spacing is not None:
                        props.LineSpacing = float(line_spacing)
                    if char_spacing is not None:
                        props.CharacterSpacing = float(char_spacing)
                    if word_spacing is not None:
                        props.WordSpacing = float(word_spacing)
                    shape.Text.SetSpaceProperties(0, props)
                    applied.append("spacing")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"间距设置失败：{exc}")
            if any(v is not None for v in (indent_left, indent_right,
                                           indent_first_line)):
                try:
                    para = shape.Text.Story
                    if indent_left is not None:
                        para.LeftIndent = float(indent_left)
                    if indent_right is not None:
                        para.RightIndent = float(indent_right)
                    if indent_first_line is not None:
                        para.FirstLineIndent = float(indent_first_line)
                    applied.append("indent")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"段落缩进设置失败：{exc}")
            if clear_tabs or tab_stops:
                try:
                    tabs = shape.Text.Story.Tabs
                    if clear_tabs:
                        tabs.Clear()
                    for spec in (tab_stops or []):
                        if isinstance(spec, dict):
                            tabs.Add(float(spec.get("position", 0)),
                                     int(spec.get("alignment", 0)),
                                     bool(spec.get("leadered", False)))
                        else:
                            tabs.Add(float(spec), 0, False)
                    applied.append("tabs")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"制表位设置失败：{exc}")
            if fit_to_path:
                try:
                    path_shape = find_shape(doc, _page(doc, page_index), fit_to_path)
                    shape.Text.FitToPath(path_shape)
                    applied.append("fit_to_path")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"文字沿路径失败：{exc}")
            if fit_frame:
                try:
                    shape.Text.FitTextToFrame()
                    applied.append("fit_frame")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"文本适应文本框失败：{exc}")
            if to_curves:
                shape.ConvertToCurves()
                applied.append("to_curves")

            # ---- 文本区间操作（Delete/Copy/Paste/Straighten，实测可用）----
            if any((range_delete, range_copy, range_paste, range_straighten)):
                try:
                    total = len(_s(lambda: str(story.Text), ""))
                    a = int(range_start or 0)
                    b = int(range_end if range_end is not None else total)
                    if not (0 <= a <= b <= max(total, 0)):
                        warnings.append(
                            f"区间越界：[{a},{b}]（当前文本长度 {total}）")
                    else:
                        if range_delete:
                            story.Range(a, b).Delete()
                            applied.append("range_delete")
                        if range_copy:
                            story.Range(a, b).Copy()
                            applied.append("range_copy")
                        if range_paste:
                            story.Range(a, b).Paste()
                            applied.append("range_paste")
                        if range_straighten:
                            story.Range(a, b).Straighten()
                            applied.append("range_straighten")
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"文本区间操作失败：{exc}")

            # ---- OpenType 特性（第六波实测：SetOpenTypeFeature/GetOpenTypeFeature）----
            reading: dict = {}
            if opentype and isinstance(opentype, dict):
                try:
                    total = len(_s(lambda: str(story.Text), ""))
                    a = int(range_start or 0)
                    b = int(range_end if range_end is not None else total)
                    if not (0 <= a < b <= max(total, 1)):
                        b = max(total, 1)
                        a = 0
                    rng = story.Range(a, b)
                    readback: dict = {}
                    for feat, on in opentype.items():
                        rng.SetOpenTypeFeature(str(feat), bool(on))
                        readback[str(feat)] = _n(lambda f=feat: int(rng.GetOpenTypeFeature(str(f))))
                    applied.append("opentype")
                    reading["opentype"] = readback
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"OpenType 特性设置失败：{exc}")

            # ---- 读取 / 导出（第五波实测：GetContents / Range / ExportToFile）----
            if read_contents:
                reading["contents"] = _s(lambda: shape.Text.GetContents(2), "")
            if range_start is not None or range_end is not None:
                a = int(range_start or 0)
                b = int(range_end if range_end is not None else a + 1)
                if not 0 <= a < b:
                    warnings.append("range_start/range_end 需满足 0 <= start < end（实测 0 基）")
                else:
                    reading["range"] = [a, b]
                    reading["range_text"] = _s(lambda: shape.Text.Range(a, b).Text, "")
            if export_path:
                try:
                    shape.Text.ExportToFile(str(export_path), int(export_start),
                                            int(export_count), 0)
                    applied.append("export_text")
                    reading["exported"] = str(export_path)
                except Exception as exc:  # noqa: BLE001
                    warnings.append(f"导出文本失败（StartIndex 需 >=1）：{exc}")
            if find_query:
                try:
                    ranges = shape.Text.Story.FindRanges(str(find_query))
                    count = _n(lambda: int(ranges.Count), 0)
                    reading["find_ranges"] = {
                        "query": str(find_query), "count": count,
                        "texts": [_s(lambda i=i: str(ranges.Item(i).Text))
                                  for i in range(1, min(count, 20) + 1)]}
                except Exception as exc:  # noqa: BLE001
                    reading["find_ranges"] = {"query": str(find_query),
                                              "error": str(exc)}
            if runs:
                reading["runs"] = _text_runs(shape)
            if read_font:
                info = _text_font_info(shape)
                if info is not None:
                    reading["font"] = info
            if read_hyphenation:
                info = _text_hyphenation_info(shape)
                if info is not None:
                    reading["hyphenation"] = info
            story = _s(lambda: shape.Text.Story)
            if story is not None:
                reading["indent"] = {
                    "left": _s(lambda: float(story.LeftIndent)),
                    "right": _s(lambda: float(story.RightIndent)),
                    "first_line": _s(lambda: float(story.FirstLineIndent)),
                }
                try:
                    tabs = story.Tabs
                    count = _n(lambda: int(tabs.Count), 0)
                    reading["tabs"] = [
                        {"position": _s(lambda i=i: float(tabs.Item(i).Position)),
                         "alignment": _n(lambda i=i: int(tabs.Item(i).Alignment)),
                         "leadered": bool(_s(lambda i=i: bool(tabs.Item(i).Leadered), False))}
                        for i in range(1, min(count, 100) + 1)]
                except Exception:
                    reading["tabs"] = []
            return {
                "applied": applied,
                "warnings": warnings,
                "text": _s(lambda: str(story.Text), ""),
                "bbox": _bbox(shape),
                **reading,
            }


def replace_text(doc: Any, find: str, replace: str, *, all_pages: bool = True,
                 case_sensitive: bool = True) -> dict:
    """优先用 X8 原生 Page.TextReplace；失败时回退到逐对象替换。"""
    if not find:
        raise CdrBadArgument("find 不能为空")
    replaced = 0
    details: list[dict] = []
    with mm_units(doc):
        start = 1 if all_pages else int(doc.ActivePage.Index)
        stop = int(doc.Pages.Count) if all_pages else start
        with atomic(doc, "cdr_replace_text"):
            for i in range(start, stop + 1):
                page = doc.Pages.Item(i)
                before = _page_text(page)
                try:
                    page.TextReplace(find, replace, bool(case_sensitive), False)
                    after = _page_text(page)
                    if before != after:
                        replaced += after.count(replace) - before.count(replace)
                    continue
                except Exception:
                    pass
                # 回退：逐对象替换
                for j in range(1, int(page.Shapes.Count) + 1):
                    shape = page.Shapes.Item(j)
                    try:
                        if int(shape.Type) != TEXT_TYPE:
                            continue
                        story = shape.Text.Story
                        original = str(story.Text)
                    except Exception:
                        continue
                    hay = original if case_sensitive else original.lower()
                    needle = find if case_sensitive else find.lower()
                    if needle not in hay:
                        continue
                    story.Text = original.replace(find, replace)
                    replaced += 1
                    details.append({"page": i, "name": str(shape.Name),
                                    "before": original[:60]})
    return {"replaced": max(replaced, 0), "matches": details[:20]}


def _page_text(page: Any) -> str:
    parts: list[str] = []
    try:
        count = int(page.Shapes.Count)
    except Exception:
        return ""
    for i in range(1, count + 1):
        try:
            shape = page.Shapes.Item(i)
            if int(shape.Type) == TEXT_TYPE:
                parts.append(str(shape.Text.Story.Text))
        except Exception:
            continue
    return "\n".join(parts)


# --------------------------------------------------------------------------- #
# 导入
# --------------------------------------------------------------------------- #
def import_file(doc: Any, path: str, page_index: int | None = None,
                center: bool = True, options: dict | None = None,
                filter_id: int | None = None) -> dict:
    """导入位图/矢量文件（X8 必须传 3 个参数：文件名、过滤器、导入选项对象）。"""
    app = _app_of(doc)
    applied_opts: list[str] = []
    with mm_units(doc):
        page = _page(doc, page_index)
        layer = _active_layer(page)
        opts = app.CreateStructImportOptions()
        opt_map = {
            "link_bitmap_externally": "LinkBitmapExternally",
            "maintain_layers": "MaintainLayers",
            "resample_dpi_x": "ResampleDpiX", "resample_dpi_y": "ResampleDpiY",
            "resample_dpi": "ResampleDpiX",
            "resample_width": "ResampleWidth", "resample_height": "ResampleHeight",
            "crop_left": "CropLeft", "crop_top": "CropTop",
            "crop_width": "CropWidth", "crop_height": "CropHeight",
            "detect_watermark": "DetectWatermark",
            "extract_icc_profile": "ExtractICCProfile",
            "icc_file_name": "ICCFileName", "use_opi_links": "UseOPILinks",
            "convert_tables_to_text": "ConvertTablesToText",
            "force_cmyk_black_text": "ForceCMYKBlackText",
            "text_formatting": "TextFormatting", "image_index": "ImageIndex",
            "mode": "Mode", "code_page": "CodePage",
            "combine_multilayer_bitmaps": "CombineMultilayerBitmaps",
        }
        for key, val in (options or {}).items():
            attr = opt_map.get(str(key).lower())
            if attr is None:
                continue
            try:
                setattr(opts, attr, val)
                applied_opts.append(str(key))
                if str(key).lower() == "resample_dpi":
                    opts.ResampleDpiY = val
            except Exception:
                continue
        before = _n(lambda: int(page.Shapes.Count), 0)
        if filter_id is not None:
            # 实测：Layer.ImportEx(FileName, Filter, Options) 可显式指定过滤器
            layer.ImportEx(path, int(filter_id), opts)
        else:
            layer.Import(path, 0, opts)
        after = _n(lambda: int(page.Shapes.Count), 0)
        shape = None
        if after > before:
            shape = _s(lambda: page.Shapes.Item(after))
            if shape is not None:
                try:
                    _ensure_name(shape)
                except Exception:
                    pass
            if center and shape is not None:
                try:
                    pw = float(page.SizeWidth)
                    ph = float(page.SizeHeight)
                    shape.SetPosition((pw - float(shape.SizeWidth)) / 2,
                                      (ph - float(shape.SizeHeight)) / 2)
                except Exception:
                    pass
        result = {
            "imported": True,
            "shapes_added": max(after - before, 0),
            "shape": shape_to_dict(shape, after) if shape is not None else None,
        }
        if applied_opts:
            result["import_options"] = applied_opts
        return result


# --------------------------------------------------------------------------- #
# 导出
# --------------------------------------------------------------------------- #
def export(doc: Any, fmt: str, path: str, *, dpi: int | None = None,
           selection: bool = False, options: dict | None = None,
           api: str = "ex") -> dict:
    """导出。`api` 可选 ex（ExportEx，默认）/ document（Document.Export）/
    bitmap（Document.ExportBitmap，只对位图格式有意义）。"""
    fmt = (fmt or "").lower().strip().lstrip(".")
    api = (api or "ex").lower()
    if api in {"document", "bitmap"}:
        # 实测：Document.Export(FileName, Filter, Range, Options, PaletteOptions) 与
        # Document.ExportBitmap(16 参) 都能落盘；bitmap 版返回 ICorelExportFilter
        if fmt == "pdf":
            doc.PublishToPDF(path)
            return _export_result(fmt, path, dpi=None)
        if fmt not in EXPORT_FILTERS:
            raise CdrBadArgument(f"不支持的格式 {fmt}；可用：png / jpg / svg / pdf")
        rng = EXPORT_RANGE_SELECTION if selection else EXPORT_RANGE_CURRENT_PAGE
        app = _app_of(doc)
        if api == "document":
            doc.Export(path, EXPORT_FILTERS[fmt], rng, None, None)
        else:
            # ExportBitmap 返回 ICorelExportFilter，必须调 Finish() 才会真正写盘
            res = int(dpi or 96)
            handle = doc.ExportBitmap(path, EXPORT_FILTERS[fmt], rng, 0, res, res,
                                      96, 96, 0, False, False, False, False, 0,
                                      None, None)
            if handle is not None:
                handle.Finish()
        for _ in range(40):
            if os.path.exists(path) and os.path.getsize(path) > 0:
                break
            time.sleep(0.05)
        result = _export_result(fmt, path, dpi=dpi)
        result["export_api"] = api
        return result
    if fmt == "pdf":
        doc.PublishToPDF(path)
        return _export_result(fmt, path, dpi=None)

    if fmt not in EXPORT_FILTERS:
        raise CdrBadArgument(f"不支持的格式 {fmt}；可用：png / jpg / svg / pdf")

    app = _app_of(doc)
    applied_opts: list[str] = []
    opt_map = {
        "anti_aliasing_type": "AntiAliasingType", "compression": "Compression",
        "dithered": "Dithered", "transparent": "Transparent",
        "maintain_aspect": "MaintainAspect", "maintain_layers": "MaintainLayers",
        "overwrite": "Overwrite", "matted": "Matted",
        "matte_masked_only": "MatteMaskedOnly",
        "always_overprint_black": "AlwaysOverprintBlack",
        "use_color_profile": "UseColorProfile",
        "resolution_x": "ResolutionX", "resolution_y": "ResolutionY",
        "size_x": "SizeX", "size_y": "SizeY",
    }
    opts = None
    if dpi or options:
        opts = app.CreateStructExportOptions()
        if dpi:
            opts.ResolutionX = int(dpi)
            opts.ResolutionY = int(dpi)
        for key, val in (options or {}).items():
            attr = opt_map.get(str(key).lower())
            if attr is None:
                continue
            try:
                setattr(opts, attr, val)
                applied_opts.append(str(key))
            except Exception:
                continue

    rng = EXPORT_RANGE_SELECTION if selection else EXPORT_RANGE_CURRENT_PAGE
    try:
        handle = doc.ExportEx(path, EXPORT_FILTERS[fmt], rng, opts, None)
        handle.Finish()
        for _ in range(40):
            if os.path.exists(path) and os.path.getsize(path) > 0:
                break
            time.sleep(0.05)
        result = _export_result(fmt, path, dpi=dpi)
        if applied_opts:
            result["export_options"] = applied_opts
        return result
    except Exception:
        # 实测：**空页面上带 StructExportOptions 导出会失败**（com_error「发生意外」）。
        # 此时退回默认分辨率再试一次，并把情况如实告诉调用方。
        if opts is None:
            raise
        handle = doc.ExportEx(path, EXPORT_FILTERS[fmt], rng, None, None)
        handle.Finish()
        for _ in range(40):
            if os.path.exists(path) and os.path.getsize(path) > 0:
                break
            time.sleep(0.05)
        result = _export_result(fmt, path, dpi=None)
        result.setdefault("warnings", []).append(
            f"指定 DPI({dpi})/选项导出失败（页面可能为空）；已按 CorelDRAW 默认分辨率导出"
        )
        return result


def _export_result(fmt: str, path: str, dpi: int | None) -> dict:
    pixels = pixel_size(path)
    result = {
        "format": fmt,
        "path": path,
        "bytes": _s(lambda: os.path.getsize(path)),
        "pixels": list(pixels) if pixels else None,
    }
    if dpi:
        result["requested_dpi"] = dpi
        result["note"] = "X8 实测：resolution 参数≈实际 dpi，大页面误差 <1%，小页面可能有 5~8% 舍入"
    if fmt == "pdf":
        result["note"] = "PDF 使用文档当前发布设置；首次调用较慢（实测 11~20s）"
    return result


def pixel_size(path: str) -> tuple[int, int] | None:
    """读 PNG / JPEG 的像素尺寸，用于回报真实导出结果。"""
    try:
        with open(path, "rb") as fh:
            head = fh.read(32)
            if head[:8] == b"\x89PNG\r\n\x1a\n":
                w, h = struct.unpack(">II", head[16:24])
                return int(w), int(h)
            if head[:2] == b"\xff\xd8":
                fh.seek(2)
                while True:
                    marker = fh.read(2)
                    if len(marker) < 2 or marker[0] != 0xFF:
                        return None
                    code = marker[1]
                    if code in (0xC0, 0xC1, 0xC2, 0xC3):
                        fh.read(3)
                        h, w = struct.unpack(">HH", fh.read(4))
                        return int(w), int(h)
                    seg_len = struct.unpack(">H", fh.read(2))[0]
                    fh.seek(seg_len - 2, 1)
    except Exception:
        return None
    return None
