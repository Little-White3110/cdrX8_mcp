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


def undo(doc: Any, steps: int = 1) -> dict:
    n = max(1, int(steps))
    done = 0
    for _ in range(n):
        try:
            doc.Undo()
            done += 1
        except Exception:
            break
    return {"undone": done}


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


def delete_page(doc: Any, index: int) -> dict:
    doc.DeletePages(int(index), 1)
    return {"pages": _n(lambda: int(doc.Pages.Count))}


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
                 printable: bool | None = None, move: str | None = None) -> dict:
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
    if act == "delete":
        layer.Delete()
        return {"layers": list_layers(doc, page_index)}
    raise CdrBadArgument(f"未知 action {action!r}；可用 list/create/activate/set/move/delete")


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
                  all_shapes: bool = False) -> dict:
    """把对象加入选择集（供对齐/布尔/组合使用）。"""
    with mm_units(doc):
        page = _page(doc, page_index)
        try:
            doc.Selection().RemoveAll()
        except Exception:
            pass
        picked = 0
        if rect:
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
                shape.AddToSelection()
                picked += 1
        return {"selected": picked}


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
    "text", "paragraph_text", "guide",
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


def _short_id() -> str:
    _CREATED["n"] += 1
    return hashlib.sha1(f"{_CREATED['n']}-{time.time()}".encode()).hexdigest()[:8]


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


def apply_fill(doc: Any, shape: Any, spec: Any) -> str:
    """填充：纯色 / 渐变 / 纹理 / 无色。返回实际应用的种类。"""
    if spec in (None, False, "none", {"type": "none"}):
        shape.Fill.ApplyNoFill()
        return "none"

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

    if isinstance(spec, dict) and spec.get("type") == "texture":
        shape.Fill.ApplyTextureFill(str(spec.get("name", "")), str(spec.get("library", "")))
        return "texture"

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

    app = None
    if "width" in spec:
        shape.Outline.Width = float(spec["width"])
    if "color" in spec:
        _outline_color(shape, spec["color"])
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


def _outline_color(shape: Any, color: Any) -> None:
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
    app = _app_of(doc)
    if isinstance(spec, str):
        r, g, b = _hex_to_rgb(spec)
        return app.CreateRGBColor(r, g, b)
    if isinstance(spec, dict):
        if all(k in spec for k in ("c", "m", "y", "k")):
            return app.CreateCMYKColor(
                int(spec["c"]), int(spec["m"]), int(spec["y"]), int(spec["k"]))
        if all(k in spec for k in ("r", "g", "b")):
            return app.CreateRGBColor(int(spec["r"]), int(spec["g"]), int(spec["b"]))
    raise CdrBadArgument('颜色需为 "#RRGGBB" 或 {c,m,y,k} 或 {r,g,b}')


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

    new_name = None
    if changes.get("duplicate"):
        dup = shape.Duplicate(float(changes.get("dup_dx", 5)), float(changes.get("dup_dy", -5)))
        new_name = _s(lambda: str(dup.Name))
        applied.append("duplicate")

    if changes.get("delete"):
        shape.Delete()
        applied.append("delete")

    return {
        "applied": applied,
        "name": new_name or _s(lambda: str(shape.Name)),
        "bbox": None if changes.get("delete") else _bbox(shape),
        "warnings": warnings,
    }


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
    "fillet", "chamfer", "scallop", "to_curves", "powerclip",
)


def shape_op(doc: Any, page_index: int | None, op: str, target: dict | None = None,
             other: dict | None = None, *, radius: float = 0, distance_a: float = 0,
             distance_b: float = 0, leave_source: bool = True, leave_target: bool = True,
             center_in_container: bool = True, label: str = "cdr_shape_ops") -> dict:
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

        shape = find_shape(doc, page, target or {})

        if op == "break_apart":
            shape.BreakApart()
            return {"op": op, "done": True}
        if op == "to_curves":
            shape.ConvertToCurves()
            return {"op": op, "name": _s(lambda: str(shape.Name)), "bbox": _bbox(shape)}
        if op == "fillet":
            shape.Fillet(float(radius), True)
            return {"op": op, "radius": radius, "bbox": _bbox(shape)}
        if op == "chamfer":
            shape.Chamfer(float(distance_a), float(distance_b), True)
            return {"op": op, "bbox": _bbox(shape)}
        if op == "scallop":
            shape.Scallop(float(radius), True)
            return {"op": op, "radius": radius, "bbox": _bbox(shape)}

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
def set_text(doc: Any, page_index: int | None, target: dict, *, text: str | None = None,
             font: str | None = None, size: float | None = None, bold: bool | None = None,
             italic: bool | None = None, underline: bool | None = None,
             alignment: str | None = None, to_curves: bool = False,
             fit_to_path: dict | None = None, fit_frame: bool = False,
             label: str = "cdr_text") -> dict:
    with mm_units(doc):
        page = _page(doc, page_index)
        shape = find_shape(doc, page, target)
        if int(_s(lambda: int(shape.Type), -1)) != TEXT_TYPE:
            raise CdrBadArgument("目标不是文字对象")
        with atomic(doc, label):
            story = shape.Text.Story
            applied: list[str] = []
            warnings: list[str] = []
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
            return {
                "applied": applied,
                "warnings": warnings,
                "text": _s(lambda: str(story.Text), ""),
                "bbox": _bbox(shape),
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
