"""领域层：把 CorelDRAW X8 的 COM 对象模型包成 Agent 好用的语义操作。

所有坐标一律是**毫米**、**页面左下角为原点、Y 轴向上**；
调用前后会用 `mm_units()` 临时把文档单位切成毫米再切回去
（实测：改 Unit 不会让文档变 Dirty，见 README「X8 差异」）。
"""

from __future__ import annotations

from contextlib import contextmanager
from typing import Any, Iterable, Iterator

from .constants import (
    EXPORT_RANGE_CURRENT_PAGE,
    FILTER_PNG,
    FILTER_SVG,
    NAME_PREFIX,
    SHAPE_TYPES,
    TEXT_TYPE,
    UNIT_MM,
)
from .errors import CdrBadArgument, CdrComFailure, CdrDocumentNotFound

ORIENTATION = {1: "portrait", 2: "landscape"}


# --------------------------------------------------------------------------- #
# 单位
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


# --------------------------------------------------------------------------- #
# 文档
# --------------------------------------------------------------------------- #
def find_document(app: Any, name: str) -> Any:
    """按文档名在当前 Application 里取文档对象（COM 对象不跨调用持有）。"""
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
                    orientation: str | None) -> Any:
    doc = app.CreateDocument()
    with mm_units(doc):
        if width and height:
            if width <= 0 or height <= 0:
                raise CdrBadArgument("width/height 必须为正数（毫米）")
            doc.ActivePage.SetSize(float(width), float(height))
        if orientation:
            want = 2 if orientation.lower().startswith("land") else 1
            doc.ActivePage.Orientation = want
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
    """关闭文档。save=True 时先尝试保存（仅当文档已有路径）。"""
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
                    "layers": [{"name": name} for name in _layer_names(page)],
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


def set_page_size(doc: Any, width: float, height: float, page_index: int | None = None) -> dict:
    with mm_units(doc):
        page = _page(doc, page_index)
        page.SetSize(float(width), float(height))
        return {
            "index": _n(lambda: int(page.Index)),
            "width_mm": _round(_s(lambda: float(page.SizeWidth))),
            "height_mm": _round(_s(lambda: float(page.SizeHeight))),
        }


def activate_page(doc: Any, index: int) -> dict:
    page = doc.Pages.Item(int(index))
    page.Activate()
    return {"active": _n(lambda: int(doc.ActivePage.Index))}


def add_pages(doc: Any, count: int = 1) -> dict:
    doc.AddPages(max(1, int(count)))
    return {"pages": _n(lambda: int(doc.Pages.Count))}


def delete_page(doc: Any, index: int) -> dict:
    doc.DeletePages(int(index))
    return {"pages": _n(lambda: int(doc.Pages.Count))}


def _page(doc: Any, index: int | None) -> Any:
    if index is None:
        return doc.ActivePage
    return doc.Pages.Item(int(index))


# --------------------------------------------------------------------------- #
# 形状
# --------------------------------------------------------------------------- #
def shape_to_dict(shape: Any, index: int) -> dict:
    bbox = _bbox(shape)
    d: dict[str, Any] = {
        "index": index,
        "name": _s(lambda: str(shape.Name)) or "",
        "type": SHAPE_TYPES.get(_n(lambda: int(shape.Type), -1), "unknown"),
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
    left, bottom, width, height = (float(box[0]), float(box[1]), float(box[2]), float(box[3]))
    return {
        "x": _round(left),
        "y": _round(bottom),
        "width": _round(width),
        "height": _round(height),
    }


def list_shapes(doc: Any, page_index: int | None = None) -> list[dict]:
    with mm_units(doc):
        page = _page(doc, page_index)
        count = _n(lambda: int(page.Shapes.Count), 0)
        return [shape_to_dict(page.Shapes.Item(i), i) for i in range(1, count + 1)]


def total_shapes(doc: Any) -> int:
    with mm_units(doc):
        total = 0
        for i in range(1, _n(lambda: int(doc.Pages.Count), 1) + 1):
            try:
                total += int(doc.Pages.Item(i).Shapes.Count)
            except Exception:
                pass
        return total


def find_shape(doc: Any, page: Any, target: dict) -> Any:
    """按 index / name / selection 定位一个形状。"""
    if not isinstance(target, dict):
        raise CdrBadArgument("target 必须是对象，如 {\"index\": 1} 或 {\"name\": \"cdr_ab12\"}")

    if target.get("selection"):
        shapes = doc.Selection().Shapes
        if int(shapes.Count) == 0:
            raise CdrBadArgument("当前没有选中任何对象")
        return shapes.Item(1)

    if "name" in target:
        want = str(target["name"])
        for i in range(1, int(page.Shapes.Count) + 1):
            shape = page.Shapes.Item(i)
            if str(shape.Name) == want:
                return shape
        raise CdrBadArgument(f"当前页找不到名为 {want} 的对象")

    if "index" in target:
        idx = int(target["index"])
        count = int(page.Shapes.Count)
        if not 1 <= idx <= count:
            raise CdrBadArgument(f"index 越界：{idx}（当前页共 {count} 个对象）")
        return page.Shapes.Item(idx)

    raise CdrBadArgument("target 需要 index / name / selection 之一")


def create_items(doc: Any, page_index: int | None, items: Iterable[dict],
                 label: str = "cdr_create") -> list[dict]:
    """批量创建图元；整批合成一个 undo 步。"""
    created: list[dict] = []
    with mm_units(doc):
        page = _page(doc, page_index)
        layer = page.ActiveLayer
        with atomic(doc, label):
            for item in items:
                shape = _create_one(doc, layer, item)
                created.append(
                    {
                        "name": _s(lambda: str(shape.Name)),
                        "type": SHAPE_TYPES.get(_n(lambda: int(shape.Type), -1), "unknown"),
                        "bbox": _bbox(shape),
                    }
                )
    return created


def _create_one(doc: Any, layer: Any, item: dict) -> Any:
    kind = str(item.get("type", "")).lower()

    if kind in {"rectangle", "rect"}:
        shape = layer.CreateRectangle2(
            float(item.get("x", 0)), float(item.get("y", 0)),
            float(item["width"]), float(item["height"]), 0, 0, 0, 0,
        )
    elif kind in {"ellipse", "circle"}:
        shape = layer.CreateEllipse2(
            float(item.get("cx", 0)), float(item.get("cy", 0)),
            float(item.get("rx", 0)), float(item.get("ry", item.get("rx", 0))),
        )
    elif kind in {"text", "artistic_text"}:
        shape = layer.CreateArtisticText(
            float(item.get("x", 0)), float(item.get("y", 0)), str(item.get("text", "")),
        )
    elif kind in {"paragraph_text", "paragraph"}:
        shape = layer.CreateParagraphText(
            float(item.get("x", 0)), float(item.get("y", 0)),
            float(item.get("x2", 60)), float(item.get("y2", 20)),
            str(item.get("text", "")),
        )
    else:
        raise CdrBadArgument(
            f"不支持的类型 {kind}；可用：rectangle / ellipse / text / paragraph_text"
        )

    _apply_style(doc, shape, item)
    name = str(item.get("name") or "").strip()
    shape.Name = name if name else f"{NAME_PREFIX}{_short_id(created_counter())}"
    return shape


_CREATED = {"n": 0}


def created_counter() -> int:
    _CREATED["n"] += 1
    return _CREATED["n"]


def _short_id(n: int) -> str:
    import hashlib
    import time

    return hashlib.sha1(f"{n}-{time.time()}".encode()).hexdigest()[:8]


def _apply_style(doc: Any, shape: Any, spec: dict) -> None:
    """填充 / 描边 / 文字样式。任何一项失败都不阻断创建，只在结果里说明。"""
    fill = spec.get("fill")
    if fill:
        try:
            shape.Fill.ApplyUniformFill(_color(doc, fill))
        except Exception:
            pass

    outline = spec.get("outline")
    if outline:
        try:
            if isinstance(outline, (int, float)):
                shape.Outline.Width = float(outline)
            elif isinstance(outline, dict):
                if "width" in outline:
                    shape.Outline.Width = float(outline["width"])
                if "color" in outline:
                    _outline_color(shape, outline["color"])
        except Exception:
            pass

    for key, prop in (("font", "Font"), ("size", "Size"), ("bold", "Bold"), ("italic", "Italic")):
        if key in spec and spec[key] is not None and int(shape.Type) == TEXT_TYPE:
            try:
                setattr(shape.Text.Story, prop, spec[key])
            except Exception:
                pass


def _outline_color(shape: Any, color: Any) -> None:
    """描边颜色在 X8 上的写法与填充不同，两种都试一次。"""
    if isinstance(color, str):
        rgb = _hex_to_rgb(color)
        try:
            shape.Outline.Color.RGBAssign(*rgb)
            return
        except Exception:
            pass
        try:
            shape.Outline.Color.CMYKAssign(*_rgb_to_cmyk(*rgb))
            return
        except Exception:
            pass
        return
    if isinstance(color, dict):
        shape.Outline.Color.CMYKAssign(
            int(color.get("c", 0)), int(color.get("m", 0)),
            int(color.get("y", 0)), int(color.get("k", 0)),
        )


def _color(doc: Any, spec: Any) -> Any:
    """支持 "#RRGGBB" / {"c","m","y","k"} / {"r","g","b"}。"""
    app = _app_of(doc)
    if isinstance(spec, str):
        r, g, b = _hex_to_rgb(spec)
        return app.CreateRGBColor(r, g, b)
    if isinstance(spec, dict):
        if all(k in spec for k in ("c", "m", "y", "k")):
            return app.CreateCMYKColor(
                int(spec["c"]), int(spec["m"]), int(spec["y"]), int(spec["k"])
            )
        if all(k in spec for k in ("r", "g", "b")):
            return app.CreateRGBColor(
                int(spec["r"]), int(spec["g"]), int(spec["b"])
            )
    raise CdrBadArgument("颜色需为 \"#RRGGBB\" 或 {c,m,y,k} 或 {r,g,b}")


def _app_of(doc: Any) -> Any:
    """取文档所属的 Application 对象（部分 X8 对象不暴露 Application 属性时回退到桥接实例）。"""
    try:
        app = doc.Application
        if app is not None:
            return app
    except Exception:
        pass
    from .com import get_bridge

    app = get_bridge().app
    if app is None:
        raise CdrComFailure("无法取得 CorelDRAW Application 对象")
    return app


def _hex_to_rgb(text: str) -> tuple[int, int, int]:
    s = text.strip().lstrip("#")
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
# 修改
# --------------------------------------------------------------------------- #
MOVE_KEYS = ("dx", "dy", "x", "y", "width", "height", "rotate")
STYLE_KEYS = ("fill", "outline", "text", "font", "size", "bold", "italic", "name")
ORDER_KEYS = ("to_front", "to_back", "duplicate", "delete")


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
        shape.SetPosition(
            float(changes.get("x", box[0])), float(changes.get("y", box[1])),
        )
        applied.append("position")

    if "width" in changes or "height" in changes:
        box = shape.GetBoundingBox()
        shape.SetBoundingBox(
            float(box[0]), float(box[1]),
            float(changes.get("width", box[2])), float(changes.get("height", box[3])),
        )
        applied.append("size")

    if "rotate" in changes:
        shape.Rotate(float(changes["rotate"]))
        applied.append("rotate")

    if "fill" in changes:
        try:
            shape.Fill.ApplyUniformFill(_color(doc, changes["fill"]))
            applied.append("fill")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"填充失败：{exc}")

    if "outline" in changes:
        outline = changes["outline"]
        try:
            if isinstance(outline, (int, float)):
                shape.Outline.Width = float(outline)
            elif isinstance(outline, dict):
                if "width" in outline:
                    shape.Outline.Width = float(outline["width"])
                if "color" in outline:
                    _outline_color(shape, outline["color"])
            applied.append("outline")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"描边失败：{exc}")

    if "text" in changes:
        try:
            shape.Text.Story.Text = str(changes["text"])
            applied.append("text")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"改文字失败（该对象可能不是文字）：{exc}")

    for key, prop in (("font", "Font"), ("size", "Size"), ("bold", "Bold"), ("italic", "Italic")):
        if key in changes and changes[key] is not None:
            try:
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
        dup = shape.Duplicate()
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


# --------------------------------------------------------------------------- #
# 文字
# --------------------------------------------------------------------------- #
def iter_text_shapes(doc: Any) -> Iterator[tuple[Any, Any]]:
    """遍历全文档文字对象，产出 (page, shape)。"""
    pages = _n(lambda: int(doc.Pages.Count), 0)
    for i in range(1, pages + 1):
        page = doc.Pages.Item(i)
        try:
            count = int(page.Shapes.Count)
        except Exception:
            continue
        for j in range(1, count + 1):
            try:
                shape = page.Shapes.Item(j)
                if int(shape.Type) == TEXT_TYPE:
                    yield page, shape
            except Exception:
                continue


def replace_text(doc: Any, find: str, replace: str, *, all_pages: bool = True,
                 case_sensitive: bool = True) -> dict:
    if not find:
        raise CdrBadArgument("find 不能为空")
    hits: list[dict] = []
    with mm_units(doc):
        start = 1 if all_pages else int(doc.ActivePage.Index)
        stop = int(doc.Pages.Count) if all_pages else start
        with atomic(doc, "cdr_replace_text"):
            for i in range(start, stop + 1):
                page = doc.Pages.Item(i)
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
                    updated = (
                        original.replace(find, replace)
                        if case_sensitive
                        else _replace_ci(original, find, replace)
                    )
                    try:
                        story.Text = updated
                        hits.append({"page": i, "name": str(shape.Name), "before": original[:80]})
                    except Exception:
                        continue
    return {"replaced": len(hits), "matches": hits}


def _replace_ci(text: str, find: str, replace: str) -> str:
    import re

    return re.sub(re.escape(find), lambda _m: replace, text, flags=re.IGNORECASE)


# --------------------------------------------------------------------------- #
# 导出
# --------------------------------------------------------------------------- #
def export(doc: Any, fmt: str, path: str) -> dict:
    fmt = fmt.lower().strip()
    if fmt in {"png"}:
        handle = doc.ExportEx(path, FILTER_PNG, EXPORT_RANGE_CURRENT_PAGE, None, None)
        handle.Finish()
    elif fmt in {"svg"}:
        handle = doc.ExportEx(path, FILTER_SVG, EXPORT_RANGE_CURRENT_PAGE, None, None)
        handle.Finish()
    elif fmt in {"pdf"}:
        doc.PublishToPDF(path)
    else:
        raise CdrBadArgument(f"不支持的格式 {fmt}；可用 png / svg / pdf")
    return {
        "format": fmt,
        "path": path,
        "hint": (
            "PNG/SVG 只导出**当前页**；X8 上无法通过 COM 指定导出 DPI（默认 300dpi）。"
            if fmt in {"png", "svg"}
            else "PDF 使用文档当前的发布设置；首次调用较慢（实测 11~20s）。"
        ),
    }


# --------------------------------------------------------------------------- #
# 小工具
# --------------------------------------------------------------------------- #
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
