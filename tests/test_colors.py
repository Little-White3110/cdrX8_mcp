"""纯 Python 色彩运算的单元测试（不依赖 CorelDRAW）。"""

from __future__ import annotations

import pytest

from cdrx8_mcp import colors


def test_hex_roundtrip():
    assert colors.hex_to_rgb("#3366CC") == (51, 102, 204)
    assert colors.rgb_to_hex((51, 102, 204)) == "#3366CC"
    assert colors.hex_to_rgb("#F00") == (255, 0, 0)      # 三位简写
    assert colors.hex_to_rgb("3366cc") == (51, 102, 204)  # 无 # 且小写


def test_hex_rejects_bad_input():
    with pytest.raises(ValueError):
        colors.hex_to_rgb("#FF")


def test_rgb_cmyk_known_values():
    assert colors.rgb_to_cmyk(0, 0, 0) == (0.0, 0.0, 0.0, 100.0)
    assert colors.rgb_to_cmyk(255, 255, 255) == (0.0, 0.0, 0.0, 0.0)
    assert colors.rgb_to_cmyk(255, 0, 0) == (0.0, 100.0, 100.0, 0.0)
    assert colors.cmyk_to_rgb(0, 100, 100, 0) == (255, 0, 0)


def test_cmyk_roundtrip_is_close():
    for rgb in [(51, 102, 204), (200, 100, 50), (17, 200, 90)]:
        back = colors.cmyk_to_rgb(*colors.rgb_to_cmyk(*rgb))
        assert all(abs(a - b) <= 1 for a, b in zip(rgb, back)), (rgb, back)


def test_gray_uses_luma_weights():
    assert colors.rgb_to_gray(255, 255, 255) == 100.0
    assert colors.rgb_to_gray(0, 0, 0) == 0.0
    # 绿色权重最高
    assert colors.rgb_to_gray(0, 255, 0) > colors.rgb_to_gray(255, 0, 0)


def test_hls_roundtrip():
    for rgb in [(255, 0, 0), (0, 128, 255), (128, 128, 128)]:
        h, l, s = colors.rgb_to_hls(*rgb)
        back = colors._hls_to_rgb(h, l, s)
        assert all(abs(a - b) <= 1 for a, b in zip(rgb, back)), (rgb, back)


def test_red_hls_is_zero_hue():
    h, l, s = colors.rgb_to_hls(255, 0, 0)
    assert h == 0.0 and l == 50.0 and s == 100.0


def test_lab_white_and_black():
    assert colors.rgb_to_lab(255, 255, 255)[0] == pytest.approx(100, abs=0.2)
    assert colors.rgb_to_lab(0, 0, 0)[0] == pytest.approx(0, abs=0.2)


def test_blend_endpoints_and_midpoint():
    a, b = (255, 0, 0), (0, 0, 255)
    assert colors.blend(a, b, 0.0) == a
    assert colors.blend(a, b, 1.0) == b
    assert colors.blend(a, b, 0.5) == (128, 0, 128)


def test_blend_clamps_ratio():
    a, b = (0, 0, 0), (255, 255, 255)
    assert colors.blend(a, b, -5) == a
    assert colors.blend(a, b, 5) == b


def test_describe_accepts_all_spec_forms():
    by_hex = colors.describe("#3366CC")
    by_rgb = colors.describe({"r": 51, "g": 102, "b": 204})
    by_cmyk = colors.describe({"c": 75, "m": 50, "y": 0, "k": 20})
    by_gray = colors.describe({"gray": 50})
    by_hls = colors.describe({"hls": [220, 50, 60]})

    assert by_hex["rgb"] == by_rgb["rgb"] == [51, 102, 204]
    assert by_hex["hex"] == "#3366CC"
    # CMYK 输入转回 RGB 后应接近原色
    assert all(abs(a - b) <= 3 for a, b in zip(by_cmyk["rgb"], [51, 102, 204]))
    assert by_gray["rgb"] == [128, 128, 128]
    assert all(isinstance(v, float) for v in by_hls["cmyk"])


def test_describe_rejects_unknown_spec():
    with pytest.raises(ValueError):
        colors.describe({"wat": 1})
    with pytest.raises(ValueError):
        colors.describe(12345)  # type: ignore[arg-type]
