"""纯 Python 色彩运算。

为什么不用 CorelDRAW 的 IVGColor 来算：实测 X8 的 `IVGColor` **不暴露任何分量属性**
（`Red`/`Cyan`/`Gray` 全部 AttributeError，只有 `Type` 和 `Name`），
`ConvertToCMYK()` 只改类型不给值。所以色彩运算放在本地做，既确定又可单测；
COM 只负责把颜色**应用**到对象上。
"""

from __future__ import annotations


def clamp(v: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, v))


def hex_to_rgb(text: str) -> tuple[int, int, int]:
    s = str(text).strip().lstrip("#")
    if len(s) == 3:
        s = "".join(c * 2 for c in s)
    if len(s) != 6:
        raise ValueError(f"颜色格式不对：{text}")
    return int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)


def rgb_to_hex(rgb: tuple[int, int, int]) -> str:
    return "#{:02X}{:02X}{:02X}".format(*[int(clamp(c, 0, 255)) for c in rgb])


def rgb_to_cmyk(r: float, g: float, b: float) -> tuple[float, float, float, float]:
    """0~255 的 RGB -> 0~100 的 CMYK（百分比）。"""
    r, g, b = r / 255, g / 255, b / 255
    k = 1 - max(r, g, b)
    if k >= 1:
        return 0.0, 0.0, 0.0, 100.0
    c = (1 - r - k) / (1 - k)
    m = (1 - g - k) / (1 - k)
    y = (1 - b - k) / (1 - k)
    return round(c * 100, 1), round(m * 100, 1), round(y * 100, 1), round(k * 100, 1)


def cmyk_to_rgb(c: float, m: float, y: float, k: float) -> tuple[int, int, int]:
    """0~100 的 CMYK -> 0~255 的 RGB。"""
    c, m, y, k = (v / 100 for v in (c, m, y, k))
    r = 255 * (1 - c) * (1 - k)
    g = 255 * (1 - m) * (1 - k)
    b = 255 * (1 - y) * (1 - k)
    return round(r), round(g), round(b)


def rgb_to_gray(r: float, g: float, b: float) -> float:
    """按 ITU-R BT.601 亮度权重，返回 0~100 的灰度百分比。"""
    return round((0.299 * r + 0.587 * g + 0.114 * b) / 255 * 100, 1)


def rgb_to_hls(r: float, g: float, b: float) -> tuple[float, float, float]:
    """-> (H 0~360, L 0~100, S 0~100)。"""
    r, g, b = r / 255, g / 255, b / 255
    mx, mn = max(r, g, b), min(r, g, b)
    lightness = (mx + mn) / 2
    if mx == mn:
        return 0.0, round(lightness * 100, 1), 0.0
    d = mx - mn
    s = d / (2 - mx - mn) if lightness > 0.5 else d / (mx + mn)
    if mx == r:
        h = ((g - b) / d + (6 if g < b else 0)) * 60
    elif mx == g:
        h = ((b - r) / d + 2) * 60
    else:
        h = ((r - g) / d + 4) * 60
    return round(h, 1), round(lightness * 100, 1), round(s * 100, 1)


def rgb_to_lab(r: float, g: float, b: float) -> tuple[float, float, float]:
    """sRGB -> CIE L*a*b*（D65）。"""
    def lin(c: float) -> float:
        c = c / 255
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4

    rl, gl, bl = lin(r), lin(g), lin(b)
    x = (rl * 0.4124 + gl * 0.3576 + bl * 0.1805) / 0.95047
    y = (rl * 0.2126 + gl * 0.7152 + bl * 0.0722) / 1.00000
    z = (rl * 0.0193 + gl * 0.1192 + bl * 0.9505) / 1.08883

    def f(t: float) -> float:
        return t ** (1 / 3) if t > 0.008856 else (7.787 * t + 16 / 116)

    fx, fy, fz = f(x), f(y), f(z)
    return round(116 * fy - 16, 1), round(500 * (fx - fy), 1), round(200 * (fy - fz), 1)


def blend(rgb1: tuple[float, float, float], rgb2: tuple[float, float, float],
          ratio: float) -> tuple[int, int, int]:
    """按 ratio 线性混合：0 = 全 rgb1，1 = 全 rgb2。"""
    t = clamp(float(ratio), 0.0, 1.0)
    return tuple(round(a + (b - a) * t) for a, b in zip(rgb1, rgb2))  # type: ignore[return-value]


def describe(spec: dict | str) -> dict:
    """把任意颜色写法解析成统一表示。

    支持 "#RRGGBB" / {r,g,b} / {c,m,y,k} / {gray} / {lab:[l,a,b]} / {hls:[h,l,s]}
    """
    if isinstance(spec, str):
        rgb = hex_to_rgb(spec)
    elif isinstance(spec, dict):
        if "gray" in spec:
            v = float(spec["gray"]) / 100 * 255
            rgb = (round(v), round(v), round(v))
        elif all(k in spec for k in ("r", "g", "b")):
            rgb = (int(spec["r"]), int(spec["g"]), int(spec["b"]))
        elif all(k in spec for k in ("c", "m", "y", "k")):
            rgb = cmyk_to_rgb(float(spec["c"]), float(spec["m"]),
                              float(spec["y"]), float(spec["k"]))
        elif isinstance(spec.get("hls"), (list, tuple)):
            h, l, s = spec["hls"]
            rgb = _hls_to_rgb(float(h), float(l), float(s))
        elif isinstance(spec.get("lab"), (list, tuple)):
            raise ValueError("Lab 输入需要先转 RGB；请直接给 #RRGGBB 或 {r,g,b}")
        else:
            raise ValueError(f"无法解析的颜色：{spec}")
    else:
        raise ValueError(f"无法解析的颜色：{spec!r}")

    r, g, b = (clamp(float(v), 0, 255) for v in rgb)
    return {
        "hex": rgb_to_hex((round(r), round(g), round(b))),
        "rgb": [round(r), round(g), round(b)],
        "cmyk": list(rgb_to_cmyk(r, g, b)),
        "gray": rgb_to_gray(r, g, b),
        "hls": list(rgb_to_hls(r, g, b)),
        "lab": list(rgb_to_lab(r, g, b)),
    }


def _hls_to_rgb(h: float, l: float, s: float) -> tuple[int, int, int]:
    h, l, s = h % 360 / 360, clamp(l / 100, 0, 1), clamp(s / 100, 0, 1)
    if s == 0:
        v = round(l * 255)
        return v, v, v
    q = l * (1 + s) if l < 0.5 else l + s - l * s
    p = 2 * l - q

    def hue(t: float) -> float:
        t = t % 1
        if t < 1 / 6:
            return p + (q - p) * 6 * t
        if t < 1 / 2:
            return q
        if t < 2 / 3:
            return p + (q - p) * (2 / 3 - t) * 6
        return p

    return (round(hue(h + 1 / 3) * 255), round(hue(h) * 255), round(hue(h - 1 / 3) * 255))
