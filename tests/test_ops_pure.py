"""纯逻辑测试：不依赖 CorelDRAW，覆盖几何/颜色/导出解析等内部工具函数。"""

from __future__ import annotations

import struct

import pytest

from cdrx8_mcp import ops
from cdrx8_mcp.constants import EXPORT_FILTERS, PAGE_PRESETS, SHAPE_TYPES, TEXT_TYPE


def test_shape_type_map_matches_measurements():
    # 这几个 ID 是实测结论，改错会导致 find/select 全部失准
    assert SHAPE_TYPES[1] == "rectangle"
    assert SHAPE_TYPES[2] == "ellipse"
    assert SHAPE_TYPES[3] == "curve"
    assert SHAPE_TYPES[4] == "polygon"
    assert SHAPE_TYPES[6] == "text"
    assert SHAPE_TYPES[7] == "group", "实测：组合体与图纸网格都是 7"
    assert SHAPE_TYPES[9] == "guide"
    assert TEXT_TYPE == 6


def test_export_filters():
    assert EXPORT_FILTERS == {"png": 802, "jpg": 774, "svg": 1345}


def test_page_presets_include_common_sizes():
    for key in ("a4", "a3", "card"):
        assert key in PAGE_PRESETS


@pytest.mark.parametrize("text,rgb", [
    ("#FF0000", (255, 0, 0)),
    ("00FF00", (0, 255, 0)),
    ("#0000ff", (0, 0, 255)),
])
def test_hex_to_rgb(text, rgb):
    assert ops._hex_to_rgb(text) == rgb


def test_hex_to_rgb_rejects_bad_input():
    with pytest.raises(Exception):
        ops._hex_to_rgb("#FFF")


def test_rgb_to_cmyk_black_and_white():
    assert ops._rgb_to_cmyk(0, 0, 0) == (0, 0, 0, 100)
    assert ops._rgb_to_cmyk(255, 255, 255) == (0, 0, 0, 0)
    c, m, y, k = ops._rgb_to_cmyk(255, 0, 0)
    assert (c, m, y, k) == (0, 100, 100, 0)


@pytest.mark.parametrize("value,expected", [
    ("left", 0), ("center", 1), ("right", 2), ("justify", 3), (2, 2), ("unknown", 0),
])
def test_text_align_mapping(value, expected):
    assert ops._align(value) == expected


def test_pixel_size_reads_png(tmp_path):
    p = tmp_path / "a.png"
    ihdr = struct.pack(">II", 640, 480)
    p.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * 8 + ihdr + b"\x00" * 9)
    assert ops.pixel_size(str(p)) == (640, 480)


def test_pixel_size_returns_none_for_garbage(tmp_path):
    p = tmp_path / "x.bin"
    p.write_bytes(b"not an image at all")
    assert ops.pixel_size(str(p)) is None


def test_pixel_size_missing_file():
    assert ops.pixel_size(r"C:\definitely\not\here.png") is None
