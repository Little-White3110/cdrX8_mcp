#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""端到端联调：直接调用 MCP 工具函数，走一遍真实 CorelDRAW X8。

覆盖 20 个工具里的绝大多数路径。这是**唯一**需要 CorelDRAW 的测试。

用法：
    python scripts/live_smoke.py
    python scripts/live_smoke.py --keep     # 结束后保留文档，方便肉眼看效果
"""

from __future__ import annotations

import argparse
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))

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
    ap.add_argument("--keep", action="store_true")
    ap.add_argument("-o", "--out", default=None,
                    help="产物目录，默认 <repo>/tmp/live_smoke（已在 gitignore 内）")
    args = ap.parse_args()
    out = args.out or os.path.join(ROOT, "tmp", "live_smoke")
    os.makedirs(out, exist_ok=True)
    print(f"\n输出目录：{out}\n")

    # ---------------- 1. 连接 ----------------
    print("[1] 状态与连接")
    st = server.cdr_status()
    check("cdr_status 有 running 字段", "running" in st, f"tools={st.get('tools')}")
    conn = server.cdr_connect(launch=True)
    if not no_error("cdr_connect", conn):
        return 1
    check("版本为 X8(18)", conn.get("version_major") == 18, str(conn.get("version")))

    # ---------------- 2. 新建（页面预设） ----------------
    print("\n[2] 新建文档（预设 A4）")
    op = server.cdr_open(preset="A4")
    if not no_error("cdr_open(preset=A4)", op):
        return 1
    doc_id = op["document_id"]
    check("A4 = 210x297mm", op.get("width_mm") == 210.0 and op.get("height_mm") == 297.0,
          f"{op.get('width_mm')}x{op.get('height_mm')}")

    insp = server.cdr_inspect(doc_id)
    if not no_error("cdr_inspect", insp):
        return 1
    state = insp["state"]
    check("空文档 0 对象", state["shapes"] == 0, str(state["shapes"]))
    check("能列出图层", isinstance(insp.get("layers"), list) and len(insp["layers"]) >= 1,
          str([ly["name"] for ly in insp.get("layers", [])]))

    # ---------------- 3. 创建多种图元 ----------------
    print("\n[3] 创建图元（圆角矩形/圆弧/星形/线/贝塞尔/文字/参考线）")
    items = [
        {"type": "rectangle", "x": 10, "y": 10, "width": 60, "height": 40, "radius": 5,
         "fill": {"type": "fountain", "start": "#FF0000", "end": "#0066FF",
                  "gradient_type": 1, "angle": 45},
         "outline": {"width": 0.8, "color": "#000000"}},
        {"type": "ellipse", "cx": 110, "cy": 30, "rx": 20, "ry": 20,
         "start_angle": 0, "end_angle": 120, "pie": True, "fill": "#FFAA00"},
        {"type": "star", "cx": 160, "cy": 30, "radius": 18, "sides": 5,
         "inner_radius": 8, "fill": {"c": 0, "m": 100, "y": 100, "k": 0}},
        {"type": "line", "x1": 10, "y1": 60, "x2": 90, "y2": 60},
        {"type": "bezier", "x1": 100, "y1": 60, "c1x": 120, "c1y": 80,
         "c2x": 140, "c2y": 40, "x2": 160, "y2": 60},
        {"type": "polyline", "points": [[10, 70], [50, 70], [50, 85], [10, 85]],
         "closed": True, "fill": "#E0B000"},
        {"type": "text", "x": 10, "y": 90, "text": "CorelDRAW X8", "font": "Arial",
         "size": 24, "bold": True, "fill": "#000000"},
        {"type": "guide", "x1": 0, "y1": 120, "x2": 210, "y2": 120},
    ]
    cre = server.cdr_create(doc_id, items, expected_revision=state["revision"],
                            snapshot_token=state["snapshot_token"])
    if not no_error("cdr_create(8 项)", cre):
        return 1
    check("创建 8 项", cre.get("count") == 8, str(cre.get("count")))
    state = cre["state"]

    insp = server.cdr_inspect(doc_id)
    shapes = insp.get("shapes", [])
    types = [s["type"] for s in shapes]
    check("形状含 rectangle/ellipse/polygon/curve/text",
          all(t in types for t in ("rectangle", "ellipse", "polygon", "curve", "text")),
          str(types))

    # ---------------- 4. 文字工具 ----------------
    print("\n[4] 文字工具（字体/字号/对齐/转曲）")
    text_shape = next((s for s in shapes if s["type"] == "text"), None)
    if text_shape:
        txt = server.cdr_text(doc_id, target={"name": text_shape["name"]},
                              font="Times New Roman", size=30, bold=False,
                              alignment="center",
                              expected_revision=state["revision"],
                              snapshot_token=state["snapshot_token"])
        if no_error("cdr_text", txt):
            check("文字样式已应用", "font" in txt.get("applied", []) and "size" in txt.get("applied", []),
                  str(txt.get("applied")))
            state = txt["state"]

    # ---------------- 5. 修改：阴影 + 描边 ----------------
    print("\n[5] 修改（投影 + 描边）")
    rect = next((s for s in shapes if s["type"] == "rectangle"), None)
    if rect:
        mod = server.cdr_modify(doc_id, target={"name": rect["name"]},
                                changes={"drop_shadow": {"opacity": 60, "offset_x": 3,
                                                         "offset_y": -3, "color": "#444444"},
                                         "outline": {"width": 1.0, "color": "#006600"}},
                                expected_revision=state["revision"],
                                snapshot_token=state["snapshot_token"])
        if no_error("cdr_modify(阴影+描边)", mod):
            check("投影/描边已应用",
                  "drop_shadow" in mod.get("applied", []) and "outline" in mod.get("applied", []),
                  str(mod.get("applied")))
            state = mod["state"]

    # ---------------- 5b. 虚线 / 箭头 / 文字沿路径 ----------------
    print("\n[5b] 虚线描边 + 箭头 + 文字沿路径")
    curve_s = next((s for s in shapes if s["type"] == "curve"), None)
    if curve_s:
        dash = server.cdr_modify(doc_id, target={"name": curve_s["name"]},
                                 changes={"outline": {"width": 1.0, "dash": 5,
                                                      "start_arrow": 2, "end_arrow": 3}},
                                 expected_revision=state["revision"],
                                 snapshot_token=state["snapshot_token"])
        if no_error("cdr_modify(虚线+箭头)", dash):
            check("虚线/箭头已应用", "outline" in dash.get("applied", []),
                  f"applied={dash.get('applied')} warns={dash.get('warnings')}")
            state = dash["state"]
        if text_shape and curve_s:
            onpath = server.cdr_text(doc_id, target={"name": text_shape["name"]},
                                     fit_to_path={"name": curve_s["name"]},
                                     expected_revision=state["revision"],
                                     snapshot_token=state["snapshot_token"])
            if no_error("cdr_text(fit_to_path)", onpath):
                check("文字沿路径已应用", "fit_to_path" in onpath.get("applied", []),
                      f"applied={onpath.get('applied')} warns={onpath.get('warnings')}")
                state = onpath["state"]

    # ---------------- 5c. 透明度 / 翻转 / 斜切 / 灰度 ----------------
    print("\n[5c] 透明度 / 翻转 / 斜切 / 克隆 / 灰度色")
    if rect:
        extra = server.cdr_modify(doc_id, target={"name": rect["name"]},
                                  changes={"transparency": 40, "flip": 0, "skew": {"x": 3},
                                           "fill": {"gray": 30}},
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
        if no_error("cdr_modify(透明/翻转/斜切/灰度)", extra):
            check("四项变换均生效",
                  all(k in extra.get("applied", [])
                      for k in ("transparency:uniform", "flip", "skew", "fill:uniform")),
                  f"applied={extra.get('applied')}")
            state = extra["state"]

    # ---------------- 5d. 段落文字进阶 ----------------
    print("\n[5d] 段落文字：大小写 / 行距 / 项目符号")
    para = server.cdr_create(doc_id, items=[
        {"type": "paragraph_text", "x": 120, "y": 20, "x2": 200, "y2": 90,
         "text": "line one\nline two", "font": "Arial", "size": 10}],
        expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建段落文字", para):
        state = para["state"]
        pname = para["created"][0]["name"]
        t1 = server.cdr_text(doc_id, target={"name": pname}, change_case="upper",
                             line_spacing=150, expected_revision=state["revision"],
                             snapshot_token=state["snapshot_token"])
        if no_error("cdr_text(大写+行距)", t1):
            check("大小写/行距已应用",
                  "change_case" in t1.get("applied", []) and "spacing" in t1.get("applied", []),
                  f"applied={t1.get('applied')}")
            check("文本确实变成大写", "LINE" in (t1.get("text") or "").upper(),
                  (t1.get("text") or "")[:24])
            state = t1["state"]
        t2 = server.cdr_text(doc_id, target={"name": pname}, bullet="•",
                             expected_revision=state["revision"],
                             snapshot_token=state["snapshot_token"])
        if no_error("cdr_text(项目符号)", t2):
            check("项目符号已应用", "bullet" in t2.get("applied", []),
                  f"applied={t2.get('applied')} warns={t2.get('warnings')}")
            state = t2["state"]

    # ---------------- 5e. 路径偏移 + 对象锁定 ----------------
    print("\n[5e] 路径偏移 / 锁定")
    if curve_s:
        pi = server.cdr_path(doc_id, target={"name": curve_s["name"]}, action="info")
        if no_error("cdr_path(info)", pi):
            check("能读曲线长度与子路径", (pi.get("length") or 0) > 0,
                  f"length={pi.get('length')} subpaths={pi.get('subpaths')}")
        pc = server.cdr_path(doc_id, target={"name": curve_s["name"]}, action="contour",
                             offset=2, expected_revision=state["revision"],
                             snapshot_token=state["snapshot_token"])
        if no_error("cdr_path(contour)", pc):
            check("偏移出新曲线", pc.get("created") is not None, str(pc.get("created")))
            state = pc["state"]
    if rect:
        lk = server.cdr_modify(doc_id, target={"name": rect["name"]},
                               changes={"locked": True},
                               expected_revision=state["revision"],
                               snapshot_token=state["snapshot_token"])
        if no_error("cdr_modify(locked)", lk):
            check("锁定已应用", "locked" in lk.get("applied", []), str(lk.get("applied")))
            state = lk["state"]
            ul = server.cdr_modify(doc_id, target={"name": rect["name"]},
                                   changes={"locked": False},
                                   expected_revision=state["revision"],
                                   snapshot_token=state["snapshot_token"])
            if no_error("cdr_modify(unlock)", ul):
                state = ul["state"]

    # ---------------- 6. 布尔运算与整形 ----------------
    print("\n[6] 布尔运算（weld / trim / intersect / fillet / to_curves）")
    seed = server.cdr_create(doc_id, items=[
        {"type": "rectangle", "x": 10, "y": 150, "width": 40, "height": 30},
        {"type": "ellipse", "cx": 45, "cy": 165, "rx": 15, "ry": 15},
    ], expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("准备布尔对象", seed):
        state = seed["state"]
        pair = server.cdr_find(doc_id, type_id=1)
        pair2 = server.cdr_find(doc_id, type_id=2)
        if pair["count"] and pair2["count"]:
            weld = server.cdr_shape_ops(
                doc_id, "weld", target={"name": pair["shapes"][-1]["name"]},
                other={"name": pair2["shapes"][-1]["name"]},
                expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
            if no_error("weld", weld):
                check("weld 执行成功", weld.get("op") == "weld")
                state = weld["state"]

    f = server.cdr_create(doc_id, items=[
        {"type": "rectangle", "x": 120, "y": 150, "width": 40, "height": 30}],
        expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("准备整形对象", f):
        state = f["state"]
        fil = server.cdr_shape_ops(doc_id, "fillet", target={"name": f["created"][0]["name"]},
                                   radius=4, expected_revision=state["revision"],
                                   snapshot_token=state["snapshot_token"])
        if no_error("fillet", fil):
            check("fillet 生效", fil.get("op") == "fillet")
            state = fil["state"]

    # ---------------- 7. 选择 + 组合 ----------------
    print("\n[7] 选择 / 组合 / 取消组合")
    sel = server.cdr_select(doc_id, rect={"x1": 0, "y1": 0, "x2": 210, "y2": 120})
    if no_error("cdr_select(框选)", sel):
        check("框选到对象", sel.get("selected", 0) > 0, str(sel.get("selected")))
    if sel.get("selected", 0) >= 2:
        grp = server.cdr_shape_ops(doc_id, "group", expected_revision=state["revision"],
                                   snapshot_token=state["snapshot_token"])
        if no_error("group", grp):
            check("group 成功", grp.get("done") is True,
                  f"组合体={grp.get('result_name')} type={grp.get('result_type')}")
            state = grp["state"]
            ung = server.cdr_shape_ops(doc_id, "ungroup",
                                       target={"name": grp.get("result_name")},
                                       expected_revision=state["revision"],
                                       snapshot_token=state["snapshot_token"])
            if no_error("ungroup", ung):
                check("ungroup 成功", ung.get("done") is True, str(ung.get("warning")))
                state = ung["state"]

    # ---------------- 8. 对齐 ----------------
    print("\n[8] 对齐（顶部对齐页面）")
    al = server.cdr_align(doc_id, vertical="top", to="page", type_id=1,
                          expected_revision=state["revision"],
                          snapshot_token=state["snapshot_token"])
    if no_error("cdr_align", al):
        check("对齐到 >=1 个对象", al.get("aligned", 0) >= 1, str(al.get("aligned")))
        state = al["state"]

    # ---------------- 9. 图层 ----------------
    print("\n[9] 图层（新建/激活/隐藏/移动）")
    ly = server.cdr_layers(doc_id, action="create", name="辅助层",
                           expected_revision=state["revision"],
                           snapshot_token=state["snapshot_token"])
    if no_error("cdr_layers:create", ly):
        check("新图层存在", any(x["name"] == "辅助层" for x in ly.get("layers", [])),
              str([x["name"] for x in ly.get("layers", [])]))
        state = ly["state"]
        hid = server.cdr_layers(doc_id, action="set", name="辅助层", visible=False,
                                expected_revision=state["revision"],
                                snapshot_token=state["snapshot_token"])
        if no_error("cdr_layers:set", hid):
            check("图层已隐藏",
                  any(x["name"] == "辅助层" and x["visible"] is False for x in hid.get("layers", [])))
            state = hid["state"]

    # ---------------- 10. 查找 ----------------
    print("\n[10] 查找")
    found = server.cdr_find(doc_id, type_id=6)
    if no_error("cdr_find(type=text)", found):
        check("找到文字对象", found.get("count", 0) >= 1, str(found.get("count")))

    # ---------------- 11. 页面 ----------------
    print("\n[11] 页面（新增/改尺寸/预设/切页/删除）")
    pg = server.cdr_pages(doc_id, action="add", count=1,
                          expected_revision=state["revision"],
                          snapshot_token=state["snapshot_token"])
    if no_error("cdr_pages:add", pg):
        check("变成 2 页", pg["state"]["pages"] == 2, str(pg["state"]["pages"]))
        state = pg["state"]
    psz = server.cdr_pages(doc_id, action="size", width_mm=148, height_mm=105,
                           expected_revision=state["revision"],
                           snapshot_token=state["snapshot_token"])
    if no_error("cdr_pages:size", psz):
        check("改成 148x105", psz.get("width_mm") == 148.0 and psz.get("height_mm") == 105.0,
              f"{psz.get('width_mm')}x{psz.get('height_mm')}")
        state = psz["state"]
    pres = server.cdr_pages(doc_id, action="preset", preset="A5",
                            expected_revision=state["revision"],
                            snapshot_token=state["snapshot_token"])
    if no_error("cdr_pages:preset(A5)", pres):
        check("A5 = 148x210mm", pres.get("width_mm") == 148.0 and pres.get("height_mm") == 210.0,
              f"{pres.get('width_mm')}x{pres.get('height_mm')}")
        state = pres["state"]
    guides = server.cdr_pages(doc_id, action="guides", page_index=1)
    if no_error("cdr_pages:guides", guides):
        check("能列出参考线", "count" in guides, f"count={guides.get('count')}")
    bounds = server.cdr_pages(doc_id, action="bounds", page_index=1)
    if no_error("cdr_pages:bounds", bounds):
        check("能读对象总包围盒", bool(bounds.get("bbox")), str(bounds.get("bbox")))

    # ---------------- 12. 导出（含 DPI） ----------------
    print("\n[12] 导出（切回第 1 页，png@150dpi / jpg / svg / pdf）")
    back = server.cdr_pages(doc_id, action="activate", page_index=1,
                            expected_revision=state["revision"],
                            snapshot_token=state["snapshot_token"])
    if no_error("cdr_pages:activate(1)", back):
        state = back["state"]
    pix: dict[str, list] = {}
    for fmt, dpi in (("png", 150), ("jpg", 96), ("svg", None), ("pdf", None)):
        p = os.path.join(out, f"smoke.{fmt}")
        r = server.cdr_export(doc_id, path=p, format=fmt, dpi=dpi, overwrite=True)
        if no_error(f"cdr_export({fmt})", r):
            ok = bool(r.get("exists", True)) and (r.get("bytes") or 0) > 500
            extra = f"{r.get('bytes')}B"
            if r.get("pixels"):
                extra += f" {r['pixels']}"
                pix[fmt] = r["pixels"]
            check(f"{fmt} 已生成", ok, extra)

    # DPI 只能保证"倍数关系"：X8 的 resolution 与最终 dpi 之间有个不稳定的比例，
    # 所以断言 300dpi 的像素宽约为 150dpi 的 2 倍（实测 991 → 1982）。
    r300 = server.cdr_export(doc_id, path=os.path.join(out, "dpi300.png"), format="png",
                             dpi=300, overwrite=True)
    if no_error("cdr_export(png@300dpi)", r300) and r300.get("pixels") and pix.get("png"):
        ratio = r300["pixels"][0] / max(pix["png"][0], 1)
        check("300dpi 像素宽约为 150dpi 的 2 倍", 1.85 <= ratio <= 2.15,
              f"{pix['png']} -> {r300['pixels']}  ratio={ratio:.2f}")

    # ---------------- 13. 导入 ----------------
    print("\n[13] 导入（把刚导出的 png 导回文档）")
    src = os.path.join(out, "smoke.png")
    if os.path.exists(src):
        imp = server.cdr_import(doc_id, path=src, center=True,
                                expected_revision=state["revision"],
                                snapshot_token=state["snapshot_token"])
        if no_error("cdr_import", imp):
            check("导入增加了对象", imp.get("shapes_added", 0) >= 1,
                  str(imp.get("shapes_added")))
            state = imp["state"]
            # 位图：先找出 type_id=5 的对象
            bmp = server.cdr_find(doc_id, type_id=5)
            if bmp.get("count"):
                bname = bmp["shapes"][0]["name"]
                bi = server.cdr_bitmap(doc_id, target={"name": bname}, action="info")
                if no_error("cdr_bitmap(info)", bi):
                    check("读到位图信息", bool(bi.get("dpi") or bi.get("pixels")),
                          f"pixels={bi.get('pixels')} dpi={bi.get('dpi')} mode={bi.get('mode')}")
                bc = server.cdr_bitmap(doc_id, target={"name": bname}, action="convert",
                                       mode="gray", expected_revision=state["revision"],
                                       snapshot_token=state["snapshot_token"])
                if no_error("cdr_bitmap(convert gray)", bc):
                    check("转灰度成功", bc.get("mode") == 3, f"mode={bc.get('mode')}")
                    state = bc["state"]
                br = server.cdr_bitmap(doc_id, target={"name": bname}, action="resample",
                                       width=300, height=300, dpi=150,
                                       expected_revision=state["revision"],
                                       snapshot_token=state["snapshot_token"])
                if no_error("cdr_bitmap(resample)", br):
                    check("重采样成功", bool(br.get("pixels")), f"pixels={br.get('pixels')}")
                    state = br["state"]

    # ---------------- 14. 替换文字 ----------------
    print("\n[14] 全文档文字替换")
    rep = server.cdr_replace_text(doc_id, "X8", "2026",
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
    if no_error("cdr_replace_text", rep):
        check("替换执行", rep.get("replaced", 0) >= 0, str(rep.get("replaced")))
        state = rep["state"]

    # ---------------- 15. 保存 + 撤销 ----------------
    print("\n[15] 另存 .cdr + 撤销")
    cdr = os.path.join(out, "smoke.cdr")
    sv = server.cdr_save(doc_id, path=cdr, overwrite=True)
    if no_error("cdr_save", sv):
        check("cdr 已生成", (sv.get("bytes") or 0) > 1000, f"{sv.get('bytes')}B")
        state = sv["state"]                      # 保存后 revision 变了，必须刷新
        fresh = server.cdr_inspect(doc_id)       # 再取一次保证 token 同步
        if no_error("cdr_inspect(保存后)", fresh):
            state = fresh["state"]
    un = server.cdr_undo(doc_id, steps=1)
    if no_error("cdr_undo", un):
        check("撤销成功", un.get("undone", 0) >= 1, str(un.get("undone")))
    rd = server.cdr_undo(doc_id, steps=1, redo=True)
    if no_error("cdr_undo(redo)", rd):
        check("重做成功", rd.get("redone", 0) >= 1, str(rd.get("redone")))
        state = rd["state"]

    # ---------------- 15b. 新能力：渐变微调 / 描边转曲 / 曲线拼接 / 文字内容 / 页面查询 / 文档级 ----------------
    print("\n[15b] 渐变微调 + 描边转曲")
    grad = server.cdr_create(doc_id, items=[
        {"type": "rectangle", "x": 150, "y": 60, "width": 40, "height": 30,
         "fill": {"type": "fountain", "start": "#FF0000", "end": "#0000FF",
                  "gradient_type": 1, "angle": 0}}],
        expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建渐变矩形", grad):
        state = grad["state"]
        gt = server.cdr_modify(doc_id, target={"name": grad["created"][0]["name"]},
                               changes={"gradient": {"angle": 90, "edge_pad": 5}},
                               expected_revision=state["revision"],
                               snapshot_token=state["snapshot_token"])
        if no_error("cdr_modify(gradient)", gt):
            check("渐变微调已应用", "gradient" in gt.get("applied", []),
                  f"applied={gt.get('applied')} warns={gt.get('warnings')}")
            state = gt["state"]
    otc = server.cdr_create(doc_id, items=[
        {"type": "rectangle", "x": 160, "y": 20, "width": 30, "height": 20,
         "outline": {"width": 2, "color": "#000000"}}],
        expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建描边矩形", otc):
        state = otc["state"]
        conv = server.cdr_shape_ops(doc_id, "outline_to_curves",
                                    target={"name": otc["created"][0]["name"]},
                                    expected_revision=state["revision"],
                                    snapshot_token=state["snapshot_token"])
        if no_error("描边转曲", conv):
            check("描边转曲后变成曲线", conv.get("type_after") == "curve",
                  f"{conv.get('type_before')} -> {conv.get('type_after')}")
            state = conv["state"]

    print("\n[15c] 曲线拼接 / 子路径 / 几何判定")
    pc = server.cdr_create(doc_id, items=[
        {"type": "bezier", "x1": 10, "y1": 120, "c1x": 30, "c1y": 140,
         "c2x": 50, "c2y": 100, "x2": 70, "y2": 120}],
        expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建测试曲线", pc):
        state = pc["state"]
        cname = pc["created"][0]["name"]
        for act, kw, label in (
            ("append_circle", {"x": 150, "y": 40, "radius": 8}, "追加圆"),
            ("append_rectangle", {"rect": [10, 20, 40, 40]}, "追加矩形"),
        ):
            r = server.cdr_path(doc_id, target={"name": cname}, action=act,
                                expected_revision=state["revision"],
                                snapshot_token=state["snapshot_token"], **kw)
            if no_error(f"cdr_path({act})", r):
                check(f"{label}成功", (r.get("subpaths") or 0) >= 2, f"subpaths={r.get('subpaths')}")
                state = r["state"]
        rn = server.cdr_path(doc_id, target={"name": cname}, action="reduce_nodes",
                             amount=20, expected_revision=state["revision"],
                             snapshot_token=state["snapshot_token"])
        if no_error("cdr_path(reduce_nodes)", rn):
            check("节点精简执行", "length" in rn, f"length={rn.get('length')}")
            state = rn["state"]
        sw = server.cdr_path(doc_id, target={"name": cname}, action="self_weld",
                             expected_revision=state["revision"],
                             snapshot_token=state["snapshot_token"])
        if no_error("cdr_path(self_weld)", sw):
            state = sw["state"]
        pi = server.cdr_path(doc_id, target={"name": cname}, action="point_inside",
                             x=150, y=40)
        if no_error("cdr_path(point_inside)", pi):
            check("点在形状内判定可用", "inside" in pi, f"inside={pi.get('inside')}")
        ed = server.cdr_path(doc_id, target={"name": cname}, action="equal_divide",
                             subpath=1, radius=4, offset=1,
                             expected_revision=state["revision"],
                             snapshot_token=state["snapshot_token"])
        if no_error("cdr_path(equal_divide)", ed):
            check("路径等分执行", ed.get("divisions") == 4, str(ed.get("divisions")))
            state = ed["state"]
        gp = server.cdr_path(doc_id, target={"name": cname}, action="get_point",
                             subpath=1, t=0.5)
        if no_error("cdr_path(get_point)", gp):
            check("取路径上的点", gp.get("point") is not None, str(gp.get("point")))
        si = server.cdr_path(doc_id, target={"name": cname}, action="subpath_info", subpath=1)
        if no_error("cdr_path(subpath_info)", si):
            check("子路径信息可读", "position" in si, str(si.get("position")))
        for qact, kw in (("curvature", {"t": 0.5}), ("perpendicular", {"t": 0.5}),
                         ("closest_segment", {"x": 40, "y": 130}),
                         ("is_on_path", {"x": 40, "y": 130}),
                         ("segment_at", {"t": 0.5})):
            qr = server.cdr_path(doc_id, target={"name": cname}, action=qact,
                                 subpath=1, **kw)
            if no_error(f"cdr_path({qact})", qr):
                check(f"路径查询 {qact} 可用",
                      any(k in qr for k in ("curvature", "angle", "result", "on_path", "segment")),
                      str({k: v for k, v in qr.items() if k != "action"})[:60])

    print("\n[15d] 文字内容 / 类型互转 / 文本区间")
    tpara = server.cdr_create(doc_id, items=[
        {"type": "paragraph_text", "x": 10, "y": 220, "x2": 120, "y2": 260,
         "text": "alpha beta", "font": "Arial", "size": 10}],
        expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建段落文字(内容测试)", tpara):
        state = tpara["state"]
        pname2 = tpara["created"][0]["name"]
        cs = server.cdr_text(doc_id, target={"name": pname2}, contents="gamma delta",
                             expected_revision=state["revision"],
                             snapshot_token=state["snapshot_token"])
        if no_error("cdr_text(contents)", cs):
            check("整段替换成功", "contents" in cs.get("applied", []),
                  f"applied={cs.get('applied')} text={cs.get('text')!r}")
            state = cs["state"]
        rp = server.cdr_text(doc_id, target={"name": pname2},
                             replace_old="gamma", replace_new="GAMMA",
                             expected_revision=state["revision"],
                             snapshot_token=state["snapshot_token"])
        if no_error("cdr_text(replace)", rp):
            check("文本内替换成功", "replace" in rp.get("applied", []), str(rp.get("applied")))
            state = rp["state"]
        fd = server.cdr_text(doc_id, target={"name": pname2}, find="GAMMA",
                             expected_revision=state["revision"],
                             snapshot_token=state["snapshot_token"])
        if no_error("cdr_text(find)", fd):
            check("文本查找返回位置", any("find:" in w for w in fd.get("warnings", [])),
                  str(fd.get("warnings")))
            state = fd["state"]
        to_art = server.cdr_text(doc_id, target={"name": pname2}, to_artistic=True,
                                 expected_revision=state["revision"],
                                 snapshot_token=state["snapshot_token"])
        if no_error("cdr_text(to_artistic)", to_art):
            check("段落字转艺术字", "to_artistic" in to_art.get("applied", []),
                  str(to_art.get("applied")))
            state = to_art["state"]
        to_par = server.cdr_text(doc_id, target={"name": pname2}, to_paragraph=True,
                                 expected_revision=state["revision"],
                                 snapshot_token=state["snapshot_token"])
        if no_error("cdr_text(to_paragraph)", to_par):
            check("艺术字转段落字", "to_paragraph" in to_par.get("applied", []),
                  str(to_par.get("applied")))
            state = to_par["state"]

    print("\n[15e] 页面查询")
    geo = server.cdr_pages(doc_id, action="geometry", page_index=1)
    if no_error("cdr_pages(geometry)", geo):
        check("读到页面尺寸与中心", bool(geo.get("size")), str(geo.get("size")))
    p1 = server.cdr_inspect(doc_id, page_index=1)
    probe = None
    if no_error("cdr_inspect(page1)", p1):
        for s in p1.get("shapes", []):
            b = s.get("bbox")
            if b and b.get("width") and b.get("height"):
                probe = (b["x"] + b["width"] / 2, b["y"] + b["height"] / 2, s["name"])
                break
    if probe:
        px, py, pname = probe
        sat = server.cdr_pages(doc_id, action="shape_at", page_index=1, x=px, y=py)
        if no_error("cdr_pages(shape_at)", sat):
            check("按点查到对象", sat.get("found") is True,
                  f"({px:.1f},{py:.1f}) -> {sat.get('shape')}")
        selat = server.cdr_pages(doc_id, action="select_at", page_index=1, x=px, y=py)
        if no_error("cdr_pages(select_at)", selat):
            check("按点选择对象", (selat.get("selected") or 0) >= 1, str(selat.get("selected")))
    else:
        check("找到可用于点查询的对象", False, "页面无可定位对象")
    tf = server.cdr_pages(doc_id, action="text_find", page_index=1, text="GAMMA")
    if no_error("cdr_pages(text_find)", tf):
        check("页面文本查找返回布尔", "found" in tf, str(tf.get("found")))

    print("\n[15f] 颜色运算")
    ci = server.cdr_color(action="info", color="#3366CC")
    if no_error("cdr_color(info)", ci):
        check("解析颜色得到 RGB/CMYK", bool(ci.get("rgb")) and bool(ci.get("cmyk")),
              f"rgb={ci.get('rgb')} cmyk={ci.get('cmyk')}")
    cv = server.cdr_color(action="convert", color={"r": 200, "g": 100, "b": 50},
                          color2="cmyk")
    if no_error("cdr_color(convert)", cv):
        check("转换到 CMYK", bool(cv.get("value")), str(cv.get("value")))
    bl = server.cdr_color(action="blend", color="#FF0000", color2="#0000FF", ratio=0.5)
    if no_error("cdr_color(blend)", bl):
        check("两色混合正确（红+蓝各半）", bl.get("hex") == "#800080", str(bl.get("hex")))

    print("\n[15g] 文档级操作")
    di = server.cdr_doc(doc_id, action="info")
    if no_error("cdr_doc(info)", di):
        check("读到文档概况", di.get("pages") is not None, f"pages={di.get('pages')}")
    cs2 = server.cdr_doc(doc_id, action="clear_selection")
    if no_error("cdr_doc(clear_selection)", cs2):
        check("清空选择集", cs2.get("selected") == 0, str(cs2.get("selected")))
    pa = server.cdr_doc(doc_id, action="palette", expected_revision=state["revision"],
                        snapshot_token=state["snapshot_token"])
    if no_error("cdr_doc(palette)", pa):
        check("颜色登记进调色板", pa.get("done") is True)
        state = pa["state"]
    if curve_s:
        ah = server.cdr_doc(doc_id, action="arrowhead", target={"name": curve_s["name"]},
                            expected_revision=state["revision"],
                            snapshot_token=state["snapshot_token"])
        if no_error("cdr_doc(arrowhead)", ah):
            check("用曲线生成箭头", ah.get("created") is True, str(ah.get("created")))
            state = ah["state"]
    cp = server.cdr_save(doc_id, path=os.path.join(out, "copy.cdr"), overwrite=True,
                         as_copy=True)
    if no_error("cdr_save(as_copy)", cp):
        check("另存副本成功", (cp.get("bytes") or 0) > 1000, f"{cp.get('bytes')}B")
        state = cp["state"]                       # SaveAsCopy 也推进 revision

    print("\n[15h] 位图色彩效果")
    fresh3 = server.cdr_inspect(doc_id)
    if no_error("cdr_inspect(副本后)", fresh3):
        old = state
        new = fresh3["state"]
        if old["snapshot_token"] != new["snapshot_token"]:
            print(f"    [diag] token 变化: rev {old['revision']}->{new['revision']}, "
                  f"pages {old['pages']}->{new['pages']}, shapes {old['shapes']}->{new['shapes']}")
        state = new
    bmp2 = server.cdr_find(doc_id, type_id=5)
    if bmp2.get("count"):
        bname2 = bmp2["shapes"][0]["name"]
        for eff in ({"type": "bci", "brightness": 5, "contrast": 5, "intensity": 5},
                    {"type": "invert"},
                    {"type": "posterize", "level": 4}):
            ef = server.cdr_modify(doc_id, target={"name": bname2}, changes={"effect": eff},
                                   expected_revision=state["revision"],
                                   snapshot_token=state["snapshot_token"])
            if no_error(f"cdr_modify(effect:{eff['type']})", ef):
                check(f"滤镜 {eff['type']} 生效",
                      any(a.startswith("effect:") for a in ef.get("applied", [])),
                      str(ef.get("applied")))
                state = ef["state"]

    # ---------------- 15i. 批量样式 / 填充复制 / 对齐增强 / 视图 ----------------
    print("\n[15i] 批量样式 + 填充复制 + 对齐增强 + 视图")
    pr = server.cdr_create(doc_id, items=[
        {"type": "rectangle", "x": 20, "y": 120, "width": 30, "height": 20,
         "fill": "#22AA22"},
        {"type": "rectangle", "x": 60, "y": 120, "width": 30, "height": 20}],
        expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建批量测试对象", pr):
        state = pr["state"]
        src_name, dst_name = pr["created"][0]["name"], pr["created"][1]["name"]
        ff = server.cdr_modify(doc_id, target={"name": dst_name},
                               changes={"fill_from": {"name": src_name}},
                               expected_revision=state["revision"],
                               snapshot_token=state["snapshot_token"])
        if no_error("cdr_modify(fill_from)", ff):
            check("复制填充已应用", "fill_from" in ff.get("applied", []),
                  str(ff.get("applied")))
            state = ff["state"]
        al = server.cdr_modify(doc_id, target={"name": dst_name},
                               changes={"align_to": {"mode": "page_center"}},
                               expected_revision=state["revision"],
                               snapshot_token=state["snapshot_token"])
        if no_error("cdr_modify(align_to page_center)", al):
            check("对齐页面中心已应用", "align_to" in al.get("applied", []),
                  str(al.get("applied")))
            state = al["state"]
        # 批量：按类型给所有矩形换描边
        bm = server.cdr_modify(doc_id, target={"type_id": 1},
                               changes={"outline": {"width": 0.5, "color": "#999999"}},
                               expected_revision=state["revision"],
                               snapshot_token=state["snapshot_token"])
        if no_error("cdr_modify(type_id=1 批量)", bm):
            check("批量修改命中多个对象",
                  (bm.get("matched") or 0) >= 2 and bm.get("modified") == bm.get("matched"),
                  f"matched={bm.get('matched')} modified={bm.get('modified')}")
            state = bm["state"]
        bm2 = server.cdr_modify(doc_id, target={"all": True}, changes={"locked": False},
                                expected_revision=state["revision"],
                                snapshot_token=state["snapshot_token"])
        if no_error("cdr_modify(all 批量)", bm2):
            check("全页批量修改可用", (bm2.get("matched") or 0) >= 1,
                  f"matched={bm2.get('matched')} errors={len(bm2.get('errors') or [])}")
            state = bm2["state"]
    vw = server.cdr_view(doc_id, action="refresh")
    if no_error("cdr_view(refresh)", vw):
        check("视图刷新可用", vw.get("refreshed") is True)
    vs = server.cdr_view(doc_id, action="to_screen", x=100, y=100)
    if no_error("cdr_view(to_screen)", vs):
        check("文档->屏幕坐标可用", vs.get("screen") is not None, str(vs.get("screen")))

    print("\n[15j] 位图二值化 / 扩边 + PostScript 底纹填充")
    bmp3 = server.cdr_find(doc_id, type_id=5)
    if bmp3.get("count"):
        b3 = bmp3["shapes"][0]["name"]
        bw = server.cdr_bitmap(doc_id, target={"name": b3}, action="bw", threshold=120,
                               expected_revision=state["revision"],
                               snapshot_token=state["snapshot_token"])
        if no_error("cdr_bitmap(bw)", bw):
            check("二值化成功（Mode->0）", bw.get("mode") == 0, f"mode={bw.get('mode')}")
            state = bw["state"]
        inf = server.cdr_bitmap(doc_id, target={"name": b3}, action="inflate",
                                inflate_x=1, inflate_y=1,
                                expected_revision=state["revision"],
                                snapshot_token=state["snapshot_token"])
        if no_error("cdr_bitmap(inflate)", inf):
            check("位图扩边成功", inf.get("action") == "inflate", str(inf.get("action")))
            state = inf["state"]
    ps = server.cdr_create(doc_id, items=[
        {"type": "rectangle", "x": 130, "y": 240, "width": 40, "height": 30,
         "fill": {"type": "postscript", "index": 1}}],
        expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建 PostScript 填充矩形", ps):
        check("PostScript 底纹填充可用", ps.get("count") == 1, str(ps.get("count")))
        state = ps["state"]

    # ---------------- 16. 并发校验 ----------------
    print("\n[16] 乐观锁（过期 token）")
    stale = server.cdr_create(doc_id, items=[{"type": "line", "x1": 0, "y1": 0,
                                              "x2": 10, "y2": 10}],
                              expected_revision=0, snapshot_token="deadbeefdeadbeef")
    check("过期 token 被拒绝", stale.get("error") in {"stale_view", "revision_mismatch"},
          str(stale.get("error")))

    # ---------------- 17. 收尾 ----------------
    if not args.keep:
        print("\n[17] 关闭文档")
        no_error("cdr_close", server.cdr_close(doc_id))

    passed = sum(1 for _, ok, _ in CHECKS if ok)
    print("\n" + "=" * 62)
    print(f"结果：{passed}/{len(CHECKS)} 通过")
    for name, ok, detail in CHECKS:
        if not ok:
            print(f"  未通过：{name} → {detail}")
    print(f"产物：{out}")
    print("=" * 62)
    return 0 if passed == len(CHECKS) else 1


if __name__ == "__main__":
    sys.exit(main())
