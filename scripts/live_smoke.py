#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""端到端联调：直接调用 MCP 工具函数，走一遍真实 CorelDRAW。

这是**唯一**需要 CorelDRAW 的测试。它会：
  启动 CorelDRAW → 新建文档 → 画图 → 改图 → 替换文字 → 导出三格式 → 存 CDR → 撤销 → 改页面 → 关闭

用法：
    python scripts/live_smoke.py
    python scripts/live_smoke.py --keep     # 结束后不关闭文档，方便肉眼看效果
"""

from __future__ import annotations

import argparse
import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src"))

from cdrx8_mcp import server  # noqa: E402

CHECKS: list[tuple[str, bool, str]] = []


def check(name: str, ok: bool, detail: str = "") -> bool:
    CHECKS.append((name, bool(ok), detail))
    print(f"  {'PASS' if ok else 'FAIL'}  {name}{('  ' + detail) if detail else ''}")
    return bool(ok)


def no_error(label: str, result: dict) -> bool:
    if "error" in result:
        check(label, False, f"{result.get('error')}: {result.get('message')}")
        return False
    return True


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--keep", action="store_true", help="结束后不关闭文档")
    ap.add_argument("-o", "--out", default=None)
    args = ap.parse_args()
    out = args.out or tempfile.mkdtemp(prefix="cdrx8_smoke_")
    os.makedirs(out, exist_ok=True)

    print(f"\n输出目录：{out}\n")

    print("[1] 状态与连接")
    st = server.cdr_status()
    check("cdr_status 返回 running 字段", "running" in st, f"running={st.get('running')}")
    conn = server.cdr_connect(launch=True)
    if not no_error("cdr_connect", conn):
        return 1
    check("连接到的版本是 X8(18)", conn.get("version_major") == 18, str(conn.get("version")))

    print("\n[2] 新建文档")
    op = server.cdr_open(width_mm=100, height_mm=60)
    if not no_error("cdr_open", op):
        return 1
    doc_id = op["document_id"]
    check("页面尺寸 100x60mm", op.get("width_mm") == 100.0 and op.get("height_mm") == 60.0,
          f"{op.get('width_mm')}x{op.get('height_mm')}")

    print("\n[3] 读取视图（拿 revision / snapshot_token）")
    insp = server.cdr_inspect(doc_id)
    if not no_error("cdr_inspect", insp):
        return 1
    state = insp["state"]
    rev, token = state["revision"], state["snapshot_token"]
    check("空文档 0 个对象", state["shapes"] == 0, f"shapes={state['shapes']}")
    check("拿到 snapshot_token", bool(token), token)

    print("\n[4] 批量创建（矩形 + 椭圆 + 文字）")
    created = server.cdr_create(
        doc_id,
        items=[
            {"type": "rectangle", "x": 10, "y": 10, "width": 50, "height": 30,
             "fill": {"c": 0, "m": 100, "y": 100, "k": 0}, "outline": {"width": 0.5}},
            {"type": "ellipse", "cx": 78, "cy": 25, "rx": 12, "ry": 12,
             "fill": "#0066CC"},
            {"type": "text", "x": 10, "y": 48, "text": "CorelDRAW X8 + Agent",
             "font": "Arial", "size": 12, "fill": "#000000"},
        ],
        expected_revision=rev, snapshot_token=token,
    )
    if not no_error("cdr_create", created):
        return 1
    check("创建了 3 个对象", created.get("count") == 3, str(created.get("count")))
    state = created["state"]

    print("\n[5] 复核对象与几何")
    insp = server.cdr_inspect(doc_id)
    shapes = insp.get("shapes", [])
    check("页面现在有 3 个对象", len(shapes) == 3, f"len={len(shapes)}")
    rect = next((s for s in shapes if s["type"] == "rectangle"), None)
    text = next((s for s in shapes if s["type"] == "text"), None)
    check("矩形 bbox = 50x30mm", rect and rect["bbox"]["width"] == 50.0
          and rect["bbox"]["height"] == 30.0, str(rect and rect["bbox"]))
    check("文字内容正确", text and "Agent" in (text.get("text") or ""), str(text and text.get("text")))

    print("\n[6] 修改对象（移动 + 换色 + 改字）")
    if rect is None:
        return 1
    mod = server.cdr_modify(
        doc_id, target={"name": rect["name"]},
        changes={"dx": 5, "dy": 5, "fill": "#00AA00"},
        expected_revision=state["revision"], snapshot_token=state["snapshot_token"],
    )
    if not no_error("cdr_modify", mod):
        return 1
    check("移动已生效", mod["bbox"]["x"] == 15.0 and mod["bbox"]["y"] == 15.0, str(mod["bbox"]))
    state = mod["state"]

    print("\n[7] 全文档文字替换")
    rep = server.cdr_replace_text(
        doc_id, "X8", "2026", expected_revision=state["revision"],
        snapshot_token=state["snapshot_token"],
    )
    if not no_error("cdr_replace_text", rep):
        return 1
    check("替换命中 1 处", rep.get("replaced") == 1, str(rep.get("replaced")))
    state = rep["state"]

    print("\n[8] 导出三种格式")
    for fmt in ("png", "svg", "pdf"):
        path = os.path.join(out, f"smoke.{fmt}")
        r = server.cdr_export(doc_id, path=path, format=fmt, overwrite=True)
        if no_error(f"cdr_export({fmt})", r):
            check(f"{fmt} 文件已生成", bool(r.get("exists")) and (r.get("bytes") or 0) > 500,
                  f"{r.get('bytes')} bytes")

    print("\n[9] 另存 .cdr")
    cdr = os.path.join(out, "smoke.cdr")
    sv = server.cdr_save(doc_id, path=cdr, overwrite=True)
    if no_error("cdr_save", sv):
        check("cdr 文件已生成", bool(sv.get("bytes")), f"{sv.get('bytes')} bytes")

    print("\n[10] 撤销一步")
    un = server.cdr_undo(doc_id, steps=1)
    if no_error("cdr_undo", un):
        check("撤销返回 undone>=1", un.get("undone", 0) >= 1, str(un.get("undone")))

    print("\n[11] 页面尺寸与多页")
    ps = server.cdr_pages(doc_id, action="size", width_mm=148, height_mm=105, page_index=1)
    if no_error("cdr_pages:size", ps):
        check("页面改为 148x105", ps.get("width_mm") == 148.0 and ps.get("height_mm") == 105.0,
              f"{ps.get('width_mm')}x{ps.get('height_mm')}")
    add = server.cdr_pages(doc_id, action="add", count=1)
    if no_error("cdr_pages:add", add):
        check("新增页后共 2 页", add.get("state", {}).get("pages") == 2,
              str(add.get("state", {}).get("pages")))

    print("\n[12] 并发校验（故意用过期的 token）")
    stale = server.cdr_create(
        doc_id, items=[{"type": "rectangle", "x": 1, "y": 1, "width": 5, "height": 5}],
        expected_revision=0, snapshot_token="deadbeefdeadbeef",
    )
    check("过期 token 被拒绝", stale.get("error") in {"stale_view", "revision_mismatch"},
          str(stale.get("error")))

    if not args.keep:
        print("\n[13] 关闭文档")
        cl = server.cdr_close(doc_id)
        no_error("cdr_close", cl)

    passed = sum(1 for _, ok, _ in CHECKS if ok)
    print("\n" + "=" * 60)
    print(f"结果：{passed}/{len(CHECKS)} 通过")
    for name, ok, detail in CHECKS:
        if not ok:
            print(f"  未通过：{name} → {detail}")
    print(f"产物：{out}")
    print("=" * 60)
    return 0 if passed == len(CHECKS) else 1


if __name__ == "__main__":
    sys.exit(main())
