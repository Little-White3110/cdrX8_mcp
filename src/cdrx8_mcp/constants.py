"""CorelDRAW X8（v18）实测常量。

本文件里的每个数字都是在 CorelDRAW Graphics Suite X8 18.0.0.450 上实跑得出的，
**不要凭 CorelDRAW 新版文档或网络示例修改**——X8 与 2024+ 的差异见 README「X8 差异」。
"""

from __future__ import annotations

# ProgID 探测顺序：先无版本号（跟随最后注册者），再逐个主版本回退。
# X8 = 18。范围覆盖 X6(16) ~ 2026(27)，便于同一份代码跨机使用。
PROGIDS: tuple[str, ...] = ("CorelDRAW.Application",) + tuple(
    f"CorelDRAW.Application.{v}" for v in range(27, 15, -1)
)

# ---- 单位 ----
# 实测标定（以 A4 = 210×297mm 的默认文档为基准）：
#   1 -> 8.27×11.69  英寸
#   2 -> 0.689×0.974 英尺      ← 不是毫米！误用会让尺寸放大 304.8 倍
#   3 -> 210×297     毫米      ← 本项目统一使用
#   4 -> 21×29.7     厘米      ← 误用会让尺寸放大 10 倍
#   5 -> 2480×3507   像素(300dpi)
#   7 -> 0.21×0.297  米
UNIT_MM = 3
UNIT_NAMES: dict[int, str] = {1: "in", 2: "ft", 3: "mm", 4: "cm", 5: "px", 6: "unknown", 7: "m"}

# ---- 导出过滤器 ID ----
# 实测可用：802=PNG、1345=SVG。CDR 本身没有可用的 ExportEx 过滤器
# （500/501/399 均返回「指定的过滤器 ID 无效」），存盘必须走 Document.SaveAs。
FILTER_PNG = 802
FILTER_SVG = 1345
FILTER_UNSUPPORTED = {500, 501, 399}

# ---- 形状类型 ----
# 实测：矩形=1、椭圆=2、艺术字与段落文字都是 6。
SHAPE_TYPES: dict[int, str] = {
    1: "rectangle",
    2: "ellipse",
    3: "curve",
    4: "line",
    5: "polygon",
    6: "text",
}
TEXT_TYPE = 6

# ---- 导出范围 ----
EXPORT_RANGE_CURRENT_PAGE = 1

# ---- 默认值 ----
DEFAULT_TIMEOUT_S = 60.0       # 单次 COM 调用超时
DEFAULT_STARTUP_TIMEOUT_S = 180.0  # 启动 CorelDRAW + 建文档可能较慢
DEFAULT_EXPORT_TIMEOUT_S = 180.0   # PDF 首次发布实测 11~20s，留足余量
NAME_PREFIX = "cdr_"
