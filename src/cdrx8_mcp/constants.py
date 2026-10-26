"""CorelDRAW **X8 (v18)** 实测常量。

本文件所有数字都来自 18.0.0.450 真机实测（并已与 vgcoreauto.tlb 反解结果核对），
**不要照 CorelDRAW 新版文档改**。详见 README「X8 差异」。

实测来源摘要：
* 单位枚举：以 A4(210x297mm) 默认文档标定 —— 1=in, 2=ft, 3=mm, 4=cm, 5=px, 7=m
* 形状类型：CreateRectangle2→1, CreateEllipse2→2,
  CreateCurveSegment2/CreateLineSegment/CreateSpiral→3,
  CreatePolygon2→4, CreateArtisticText→6, CreateGridBoxes→7
* 导出过滤器：802=PNG, 774=JPG, 1345=SVG 可用；EPS/AI/CDR 无可用 ID
"""

from __future__ import annotations

PROGIDS: tuple[str, ...] = ("CorelDRAW.Application",) + tuple(
    f"CorelDRAW.Application.{v}" for v in range(27, 15, -1)
)

# ---- 单位 ----
UNIT_MM = 3
UNIT_NAMES: dict[int, str] = {1: "in", 2: "ft", 3: "mm", 4: "cm", 5: "px", 6: "unknown", 7: "m"}

# ---- 导出过滤器（实测可用）----
FILTER_PNG = 802
FILTER_JPG = 774
FILTER_SVG = 1345
EXPORT_FILTERS: dict[str, int] = {"png": FILTER_PNG, "jpg": FILTER_JPG, "svg": FILTER_SVG}
EXPORT_RANGE_CURRENT_PAGE = 1
EXPORT_RANGE_SELECTION = 2

# ---- 形状类型 ----
SHAPE_TYPES: dict[int, str] = {
    1: "rectangle",
    2: "ellipse",
    3: "curve",
    4: "polygon",
    6: "text",
    7: "group",     # 实测：组合体与「图纸网格」都是 7（网格本质就是一组曲线）
    9: "guide",
    21: "bspline",   # 第五十三波实测：Layer.CreateBSpline 得到 21，ConvertToCurves 后变 3
}
TEXT_TYPE = 6

# ---- 常用页面尺寸（CorelDRAW 内置名称；"名片" 实测可用）----
PAGE_PRESETS: dict[str, str] = {
    "a0": "A0", "a1": "A1", "a2": "A2", "a3": "A3", "a4": "A4", "a5": "A5",
    "letter": "Letter", "legal": "Legal", "tabloid": "Tabloid",
    "card": "名片", "business_card": "名片",
}

# ---- 对齐方法 ID（ShapeRange.AlignAndDistribute 的 MethodH / MethodV）----
ALIGN_NONE = 0
ALIGN_MIN = 1      # 左 / 下
ALIGN_CENTER = 2   # 居中
ALIGN_MAX = 3      # 右 / 上
ALIGN_SPACE = 4    # 等间距分布
ALIGN_TO_SELECTION = 0
ALIGN_TO_PAGE = 1

# ---- 效果 ----
DROP_SHADOW_FLAT = 1
ENVELOPE_PRESET_MIN = 1
ENVELOPE_PRESET_MAX = 39

# ---- 默认值 ----
DEFAULT_TIMEOUT_S = 60.0
DEFAULT_STARTUP_TIMEOUT_S = 180.0
DEFAULT_EXPORT_TIMEOUT_S = 180.0
NAME_PREFIX = "cdr_"
