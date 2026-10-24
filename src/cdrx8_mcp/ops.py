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
                    orientation: str | None, preset: str | None = None) -> Any:
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


def open_document(app: Any, path: str) -> Any:
    return app.OpenDocument(path)


def save_as(doc: Any, path: str) -> dict:
    """X8 的 SaveAs 必须传第二个参数（Options），只传路径会抛 TypeError。"""
    old_name = str(doc.Name)
    doc.SaveAs(path, None)
    return {
        "name": _s(lambda: str(doc.Name), old_name),
        "file_name": _s(lambda: str(doc.FileName)),
        "old_name": old_name,
        "dirty": bool(_s(lambda: bool(doc.Dirty), False)),
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


def save_as_copy(doc: Any, path: str) -> dict:
    """另存副本。

    ⚠️ 实测：X8 的 `SaveAsCopy` 会把**当前文档名也改成副本名**（a.cdr -> b.cdr，
    FileName 同步变），所以调用方必须用返回的 `name` 更新会话，否则后续按名查找会失败。
    """
    old_name = _s(lambda: str(doc.Name), "")
    doc.SaveAsCopy(path, None)
    return {
        "copied_to": path,
        "name": _s(lambda: str(doc.Name), old_name),
        "old_name": old_name,
        "file_name": _s(lambda: str(doc.FileName)) or None,
    }


VIEW_ACTIONS = ("refresh", "activate", "to_screen", "to_document",
                "distance_to_screen", "distance_to_document")


def view_ops(doc: Any, action: str, *, x: float | None = None, y: float | None = None,
             distance: float = 10.0) -> dict:
    """视图/窗口操作（IVGWindow，实测可用）。"""
    act = (action or "refresh").lower()
    if act not in VIEW_ACTIONS:
        raise CdrBadArgument(f"未知 action {action!r}；可用：{', '.join(VIEW_ACTIONS)}")
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
               "fonts", "palettes", "palette_color",
               "palette_open_fixed", "palette_colors", "spot_color",
               "pattern_canvas", "textures", "metadata", "set_metadata",
               "palette_create", "data_fields", "add_data_field", "object_data",
               "color_context", "outline_styles", "settings", "set_settings")


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
            overwrite: bool = False, field_name: str | None = None,
            field_format: str = "0", settings: dict | None = None,
            register: bool = False) -> dict:
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


def add_pages(doc: Any, count: int = 1) -> dict:
    doc.AddPages(max(1, int(count)))
    return {"pages": _n(lambda: int(doc.Pages.Count))}


def insert_pages(doc: Any, count: int = 1, before: bool = True,
                 page_index: int | None = None) -> dict:
    """插入空白页（实测 `Document.InsertPages(count, before, pageIndex)`）。

    第三个参数是**页码数字**，不是 Page 对象（传对象会 TypeError）。
    """
    idx = int(page_index if page_index is not None else 1)
    with mm_units(doc):
        created = doc.InsertPages(max(1, int(count)), bool(before), idx)
        active = _n(lambda: int(doc.ActivePage.Index), 0)
    return {"action": "insert", "count": max(1, int(count)), "before": bool(before),
            "page": idx, "inserted": created is not None, "active_page": active,
            "note": "插入到第 1 页前面会把新页设为活动页；后续操作请显式传 page_index",
            "pages": _n(lambda: int(doc.Pages.Count))}


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
                  "pixel", "image_info", "duotone", "paletted")


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
                  preset: str | None = None, palette_type: int = 3,
                  num_colors: int = 256, palette_index: int | None = None,
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

        with atomic(doc, label):
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
                bmp.ConvertToDuotone(duo)
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
             hot_area: float = -1.0,
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
)

NODE_MUTATIONS = {
    "move", "select", "delete", "type", "fillet", "chamfer", "scallop",
    "smoothen", "stretch", "rotate", "skew", "break_apart",
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
             hot_area: float = -1.0, label: str = "cdr_node") -> dict:
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
                 "color_styles", "create_color_styles")
STYLE_MUTATIONS = {"create", "apply", "rename", "delete", "set_property",
                   "import_sheet", "create_color_styles"}
STYLE_CATEGORIES = ("fill", "outline", "character", "paragraph")


def _find_style(doc: Any, name: str | None) -> Any:
    if not name:
        raise CdrBadArgument("需要 style_name")
    style = _s(lambda: doc.StyleSheet.Styles.Find(str(name)))
    if style is None:
        raise CdrBadArgument(f"未找到样式 {name!r}（可先用 action='list' 查看）")
    return style


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
                 move_below: str | None = None) -> dict:
    page = _page(doc, page_index)
    act = (action or "").lower()

    if act == "create":
        if not name:
            raise CdrBadArgument("action=create 需要 name")
        layer = page.CreateLayer(name)
        return {"created": _s(lambda: str(layer.Name), name), "layers": list_layers(doc, page_index)}

    layer = _find_layer(page, name=name, index=index)

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
    "polyline", "bitmap", "text", "paragraph_text", "guide",
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
        shape = layer.CreateArtisticText(
            float(g("x", 0)), float(g("y", 0)), str(g("text", "")),
            int(g("language_id", 0)), int(g("charset", 0)),
            str(g("font", "")), float(g("size", 0)),
            bool(g("bold", False)), bool(g("italic", False)), bool(g("underline", False)),
            int(_align(g("align", "left"))),
        )
    elif kind in {"paragraph_text", "paragraph"}:
        shape = layer.CreateParagraphText(
            float(g("x", 0)), float(g("y", 0)), float(g("x2", 60)), float(g("y2", 20)),
            str(g("text", "")),
            int(g("language_id", 0)), int(g("charset", 0)),
            str(g("font", "")), float(g("size", 0)),
            bool(g("bold", False)), bool(g("italic", False)), bool(g("underline", False)),
            int(_align(g("align", "left"))),
        )
    elif kind == "guide":
        layer.CreateGuide(float(g("x1", 0)), float(g("y1", 0)),
                          float(g("x2", 100)), float(g("y2", 0)))
        return None
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
        for key, prop in (("font", "Font"), ("size", "Size"),
                          ("bold", "Bold"), ("italic", "Italic"), ("underline", "Underline")):
            if spec.get(key) not in (None, ""):
                try:
                    setattr(shape.Text.Story, prop, spec[key])
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


def color_ops(doc: Any, action: str, *, color: Any = None, color2: Any = None,
              ratio: float = 0.5, index: int = 1,
              label: str = "cdr_color") -> dict:
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
    if act not in {"info", "convert", "blend"}:
        raise CdrBadArgument("action 可用 info / convert / blend")
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
                new_name = _s(lambda: str(new_shape.Name))
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

    for key, prop in (("font", "Font"), ("size", "Size"), ("bold", "Bold"),
                      ("italic", "Italic"), ("underline", "Underline"),
                      ("alignment", None)):
        if key in changes and changes[key] is not None:
            try:
                if key == "alignment":
                    _set_text_align(shape, changes[key])
                else:
                    setattr(shape.Text.Story, prop, changes[key])
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
)


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

        shape = find_shape(doc, page, target or {})

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
                 spacing: bool = False) -> dict:
    """对齐/分布。horizontal: none/left/center/right；vertical: none/bottom/middle/top。"""
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

    with mm_units(doc):
        select_shapes(doc, page_index, name=name, type_id=type_id, rect=rect)
        sr = selection_range(doc)
        if int(sr.Count) < 1:
            raise CdrBadArgument("没有选中任何对象")
        with atomic(doc, "cdr_align"):
            sr.AlignAndDistribute(
                h, v,
                1 if str(to).lower() == "page" else 0,
                0, False, 0, 0.0, 0.0, None,
            )
        return {"aligned": int(sr.Count), "horizontal": h, "vertical": v, "to": to}


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


def set_text(doc: Any, page_index: int | None, target: dict, *, text: str | None = None,
             font: str | None = None, size: float | None = None, bold: bool | None = None,
             italic: bool | None = None, underline: bool | None = None,
             alignment: str | None = None, to_curves: bool = False,
             fit_to_path: dict | None = None, fit_frame: bool = False,
             change_case: Any = None, bullet: Any = None, drop_cap: Any = None,
             no_effect: bool = False, line_spacing: float | None = None,
             char_spacing: float | None = None, word_spacing: float | None = None,
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
             label: str = "cdr_text") -> dict:
    with mm_units(doc):
        page = _page(doc, page_index)
        shape = find_shape(doc, page, target)
        if int(_s(lambda: int(shape.Type), -1)) != TEXT_TYPE:
            raise CdrBadArgument("目标不是文字对象")
        mutating = any(v is not None for v in (
            text, font, size, bold, italic, underline, alignment,
            fit_to_path, change_case, bullet, drop_cap, line_spacing,
            char_spacing, word_spacing, insert_before, insert_after,
            contents, find, replace_old, replace_new, copy_attributes_from,
            opentype,
        )) or any((to_curves, fit_frame, no_effect, to_paragraph, to_artistic,
                   select_all, align_to_baseline, range_delete, range_copy,
                   range_paste, range_straighten))
        if not mutating:
            # 纯读取 / 导出：不建 undo 组，也不改动文档
            result: dict = {"applied": [], "warnings": [],
                            "text": _s(lambda: str(shape.Text.Story.Text), ""),
                            "bbox": _bbox(shape)}
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
            if read_font:
                info = _text_font_info(shape)
                if info is not None:
                    result["font"] = info
            if read_hyphenation:
                info = _text_hyphenation_info(shape)
                if info is not None:
                    result["hyphenation"] = info
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

            if text is not None:
                story.Text = str(text)
                applied.append("text")
            for key, prop, val in (("font", "Font", font), ("size", "Size", size),
                                   ("bold", "Bold", bold), ("italic", "Italic", italic),
                                   ("underline", "Underline", underline)):
                if val is not None:
                    try:
                        setattr(story, prop, val)
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
            if read_font:
                info = _text_font_info(shape)
                if info is not None:
                    reading["font"] = info
            if read_hyphenation:
                info = _text_hyphenation_info(shape)
                if info is not None:
                    reading["hyphenation"] = info
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
                center: bool = True) -> dict:
    """导入位图/矢量文件（X8 必须传 3 个参数：文件名、过滤器、导入选项对象）。"""
    app = _app_of(doc)
    with mm_units(doc):
        page = _page(doc, page_index)
        layer = _active_layer(page)
        opts = app.CreateStructImportOptions()
        before = _n(lambda: int(page.Shapes.Count), 0)
        layer.Import(path, 0, opts)
        after = _n(lambda: int(page.Shapes.Count), 0)
        shape = None
        if after > before:
            shape = _s(lambda: page.Shapes.Item(after))
            if center and shape is not None:
                try:
                    pw = float(page.SizeWidth)
                    ph = float(page.SizeHeight)
                    shape.SetPosition((pw - float(shape.SizeWidth)) / 2,
                                      (ph - float(shape.SizeHeight)) / 2)
                except Exception:
                    pass
        return {
            "imported": True,
            "shapes_added": max(after - before, 0),
            "shape": shape_to_dict(shape, after) if shape is not None else None,
        }


# --------------------------------------------------------------------------- #
# 导出
# --------------------------------------------------------------------------- #
def export(doc: Any, fmt: str, path: str, *, dpi: int | None = None,
           selection: bool = False) -> dict:
    fmt = (fmt or "").lower().strip().lstrip(".")
    if fmt == "pdf":
        doc.PublishToPDF(path)
        return _export_result(fmt, path, dpi=None)

    if fmt not in EXPORT_FILTERS:
        raise CdrBadArgument(f"不支持的格式 {fmt}；可用：png / jpg / svg / pdf")

    app = _app_of(doc)
    opts = app.CreateStructExportOptions() if dpi else None
    if opts is not None:
        opts.ResolutionX = int(dpi)
        opts.ResolutionY = int(dpi)

    rng = EXPORT_RANGE_SELECTION if selection else EXPORT_RANGE_CURRENT_PAGE
    try:
        handle = doc.ExportEx(path, EXPORT_FILTERS[fmt], rng, opts, None)
        handle.Finish()
        return _export_result(fmt, path, dpi=dpi)
    except Exception:
        # 实测：**空页面上带 StructExportOptions 导出会失败**（com_error「发生意外」）。
        # 此时退回默认分辨率再试一次，并把情况如实告诉调用方。
        if opts is None:
            raise
        handle = doc.ExportEx(path, EXPORT_FILTERS[fmt], rng, None, None)
        handle.Finish()
        result = _export_result(fmt, path, dpi=None)
        result.setdefault("warnings", []).append(
            f"指定 DPI({dpi}) 导出失败（页面可能为空）；已按 CorelDRAW 默认分辨率导出"
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
