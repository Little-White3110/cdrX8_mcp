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
        # 画布上对象类型/叠放顺序会影响 IsSelectable，逐个候选点试，直到能选中
        selected_ok = False
        detail = ""
        for cand in p1.get("shapes", [])[:8]:
            b = cand.get("bbox")
            if not b or not b.get("width") or not b.get("height"):
                continue
            cx = b["x"] + b["width"] / 2
            cy = b["y"] + b["height"] / 2
            selat = server.cdr_pages(doc_id, action="select_at", page_index=1, x=cx, y=cy)
            if "error" in selat:
                detail = f"{cand['name']}: {selat.get('message')}"
                continue
            got = selat.get("selected") or 0
            detail = f"{cand['name']}@{cx:.1f},{cy:.1f} -> {got}"
            if got >= 1:
                selected_ok = True
                break
        check("按点选择对象", selected_ok, detail)
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
    fresh_g = server.cdr_inspect(doc_id)
    if "state" in fresh_g:
        state = fresh_g["state"]
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
    bname2 = None
    fresh_png = os.path.join(out, "smoke.png")
    if os.path.exists(fresh_png):
        freshimp = server.cdr_import(doc_id, path=fresh_png, page_index=1, center=True,
                                     expected_revision=state["revision"],
                                     snapshot_token=state["snapshot_token"])
        if no_error("为滤镜导入干净 RGB 位图", freshimp) and freshimp.get("shape"):
            state = freshimp["state"]
            bname2 = freshimp["shape"]["name"]
    if not bname2:
        bmp2 = server.cdr_find(doc_id, type_id=5)
        if bmp2.get("count"):
            bname2 = bmp2["shapes"][0]["name"]
    if bname2:
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

    # ---------------- 15k. 第五波：节点级编辑 / 路径查询 / 排列 / 文本读取 ----------------
    print("\n[15k] 第五波：节点级编辑 + 路径查询 + 排列 + 文本读取")
    w5 = server.cdr_create(doc_id, items=[
        {"type": "polyline", "points": [[10, 160], [50, 160], [50, 190], [10, 190]],
         "closed": True, "fill": "#DDDDDD"},
        {"type": "bezier", "x1": 80, "y1": 170, "c1x": 100, "c1y": 190,
         "c2x": 120, "c2y": 150, "x2": 140, "y2": 170}],
        expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建节点测试对象", w5):
        state = w5["state"]
        poly_name = w5["created"][0]["name"]
        bez_name = w5["created"][1]["name"]

        nl = server.cdr_node(doc_id, target={"name": poly_name}, action="list")
        if no_error("cdr_node(list)", nl):
            check("节点列表读到 4 个节点（polyline 无重复节点）",
                  nl.get("nodes") == 4 and len(nl.get("node_list") or []) == 4,
                  f"nodes={nl.get('nodes')}")
        nseg = server.cdr_node(doc_id, target={"name": poly_name}, action="segments")
        if no_error("cdr_node(segments)", nseg):
            seg0 = (nseg.get("segment_list") or [{}])[0]
            check("线段详情含端点坐标",
                  seg0.get("start") is not None and seg0.get("end") is not None,
                  f"start={seg0.get('start')} end={seg0.get('end')}")
        nmv = server.cdr_node(doc_id, target={"name": poly_name}, action="move",
                              node=2, x=50, y=162,
                              expected_revision=state["revision"],
                              snapshot_token=state["snapshot_token"])
        if no_error("cdr_node(move)", nmv):
            check("节点坐标已改动", nmv.get("nodes") == 4, str(nmv.get("nodes")))
            state = nmv["state"]
        nfl = server.cdr_node(doc_id, target={"name": poly_name}, action="fillet",
                              start=1, end=2, radius=2,
                              expected_revision=state["revision"],
                              snapshot_token=state["snapshot_token"])
        if no_error("cdr_node(fillet)", nfl):
            check("节点圆角新增节点", (nfl.get("nodes") or 0) >= 5,
                  f"nodes={nfl.get('nodes')}")
            state = nfl["state"]
        ncp = server.cdr_node(doc_id, target={"name": bez_name}, action="control_points",
                              segment_index=1, control="start", x=100, y=195,
                              expected_revision=state["revision"],
                              snapshot_token=state["snapshot_token"])
        if no_error("cdr_node(control_points set)", ncp):
            check("贝塞尔控制点已写入",
                  (ncp.get("control_start") or [None, None])[1] == 195.0,
                  str(ncp.get("control_start")))
            state = ncp["state"]
        nfound = server.cdr_node(doc_id, target={"name": poly_name}, action="find",
                                 x=10, y=160, hot_area=5)
        if no_error("cdr_node(find)", nfound):
            check("按坐标找到节点", (nfound.get("node") or {}).get("index") is not None,
                  str(nfound.get("node")))

        pl = server.cdr_path(doc_id, target={"name": poly_name}, action="polyline",
                             precision=10)
        if no_error("cdr_path(polyline)", pl):
            check("折线采样取回角点", (pl.get("count") or 0) >= 4,
                  f"count={pl.get('count')} points={pl.get('points')}")
        pp = server.cdr_path(doc_id, target={"name": bez_name}, action="point_position", t=0.5)
        if no_error("cdr_path(point_position)", pp):
            check("曲线上取点可用", pp.get("point") is not None, str(pp.get("point")))
        tv = server.cdr_path(doc_id, target={"name": bez_name}, action="tangent_vector", t=0.5)
        if no_error("cdr_path(tangent_vector)", tv):
            check("切向量非零", bool(tv.get("length")), str(tv.get("vector")))
        fn = server.cdr_path(doc_id, target={"name": poly_name}, action="find_node",
                             x=50, y=162, hot_area=5)
        if no_error("cdr_path(find_node)", fn):
            check("路径级按坐标找节点", fn.get("found") is True, str(fn.get("node")))
        rv = server.cdr_path(doc_id, target={"name": poly_name}, action="reverse_direction",
                             expected_revision=state["revision"],
                             snapshot_token=state["snapshot_token"])
        if no_error("cdr_path(reverse_direction)", rv):
            check("路径方向已反转", rv.get("clockwise") is True, str(rv.get("clockwise")))
            state = rv["state"]

        md = server.cdr_modify(doc_id, target={"name": poly_name},
                               changes={"order": {"mode": "front_of",
                                                  "ref": {"name": bez_name}}},
                               expected_revision=state["revision"],
                               snapshot_token=state["snapshot_token"])
        if no_error("cdr_modify(order front_of)", md):
            check("层级相对调整已应用", "order:front_of" in md.get("applied", []),
                  str(md.get("applied")))
            state = md["state"]
        sr = server.cdr_shape_ops(doc_id, "step_repeat", target={"name": poly_name},
                                  copies=2, distance_x=10, distance_y=0,
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(step_repeat)", sr):
            check("阵列复制新建 2 个对象", (sr.get("copies_created") or 0) == 2,
                  f"copies_created={sr.get('copies_created')}")
            state = sr["state"]
        ds = server.cdr_shape_ops(doc_id, "distribute", target={"name": bez_name},
                                  distribute_type=1,
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(distribute)", ds):
            check("原生分布排列可用", ds.get("op") == "distribute", str(ds.get("op")))
            state = ds["state"]
        ps2 = server.cdr_shape_ops(doc_id, "perspective", target={"name": bez_name},
                                   expected_revision=state["revision"],
                                   snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(perspective)", ps2):
            check("透视效果已创建", ps2.get("effect") is True)
            state = ps2["state"]
        sp2 = server.cdr_shape_ops(doc_id, "separate", target={"name": bez_name},
                                   expected_revision=state["revision"],
                                   snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(separate)", sp2):
            check("效果拆分可用", sp2.get("done") is True)
            state = sp2["state"]

        # 文本读取 / 导出（复用前面创建的文字对象）
        tf = server.cdr_find(doc_id, type_id=6)
        if tf.get("count"):
            tname = tf["shapes"][0]["name"]
            tr = server.cdr_text(doc_id, target={"name": tname}, read_contents=True,
                                 range_start=0, range_end=5)
            if no_error("cdr_text(read_contents)", tr):
                check("文本读取可用", isinstance(tr.get("contents"), str)
                      and tr.get("range_text") is not None,
                      f"contents={tr.get('contents')!r} range={tr.get('range_text')!r}")
            tout = os.path.join(out, "wave5_text.txt")
            if os.path.exists(tout):
                os.remove(tout)
            te = server.cdr_text(doc_id, target={"name": tname}, export_path=tout,
                                 export_start=1, export_count=5,
                                 expected_revision=state["revision"],
                                 snapshot_token=state["snapshot_token"])
            if no_error("cdr_text(export_path)", te):
                check("文本导出成文件", os.path.exists(tout),
                      f"exists={os.path.exists(tout)}")
                state = te.get("state", state)

        un = server.cdr_doc(doc_id, "units", value=1, from_unit=1, to_unit=3)
        if no_error("cdr_doc(units)", un):
            check("单位换算可用", un.get("from_document") == 25.4, str(un))
        sc = server.cdr_doc(doc_id, "sample_color", x=20, y=170)
        if no_error("cdr_doc(sample_color)", sc):
            check("取色返回颜色串", bool(sc.get("color")), str(sc.get("color")))
        vd = server.cdr_view(doc_id, action="distance_to_document", distance=100)
        if no_error("cdr_view(distance_to_document)", vd):
            check("屏幕距离->文档距离可用", vd.get("document_distance") is not None,
                  str(vd.get("document_distance")))
        ga = server.cdr_pages(doc_id, action="guide_angle", x=0, y=0, angle=30,
                              expected_revision=state["revision"],
                              snapshot_token=state["snapshot_token"])
        if no_error("cdr_pages(guide_angle)", ga):
            check("角度参考线已创建", ga.get("created") is True)
            state = ga["state"]

    # ---------------- 15l. 第六波：对象样式 / 内存位图 / 页面插入 / 范围栅格 ----------------
    print("\n[15l] 第六波：对象样式 + 内存位图 + 页面插入 + 范围操作")
    fresh6 = server.cdr_inspect(doc_id)
    if "state" in fresh6:
        state = fresh6["state"]
    w6 = server.cdr_create(doc_id, items=[
        {"type": "rectangle", "x": 20, "y": 200, "width": 40, "height": 25,
         "fill": "#CC0000", "outline": {"width": 1, "color": "#000000"}},
        {"type": "rectangle", "x": 70, "y": 200, "width": 40, "height": 25,
         "fill": "#00CC00"},
        {"type": "polyline", "points": [[130, 200], [170, 200], [170, 230]],
         "closed": False, "outline": {"width": 1, "color": "#000000"}},
        {"type": "paragraph_text", "x": 20, "y": 240, "x2": 120, "y2": 270,
         "text": "Wave Six Style", "font": "Arial", "size": 12},
        {"type": "bitmap", "x": 180, "y": 240, "width": 30, "height": 30,
         "color": "#FF0000", "pixels_w": 16, "pixels_h": 16,
         "fills": [{"x": 0, "y": 0, "width": 8, "height": 8, "color": "#0000FF"}]},
    ], expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建第六波测试对象", w6):
        state = w6["state"]
        r1, r2, cv1, tx1 = (w6["created"][i]["name"] for i in range(4))
        bmp_name = w6["created"][4]["name"]

        # --- 对象样式 ---
        sl = server.cdr_style(doc_id, action="list")
        if no_error("cdr_style(list)", sl):
            check("初始样式列表可读", isinstance(sl.get("styles"), list),
                  f"count={sl.get('count')}")
        sc1 = server.cdr_style(doc_id, action="create", target={"name": r1},
                               category="fill", style_name="W6Fill",
                               expected_revision=state["revision"],
                               snapshot_token=state["snapshot_token"])
        if no_error("cdr_style(create fill)", sc1):
            check("从对象创建填充样式", "W6Fill" in (sc1.get("created") or []),
                  str(sc1.get("created")))
            state = sc1["state"]
        sc2 = server.cdr_style(doc_id, action="create", target={"name": r1},
                               category="outline", style_name="W6Outline",
                               expected_revision=state["revision"],
                               snapshot_token=state["snapshot_token"])
        if no_error("cdr_style(create outline)", sc2):
            check("从对象创建轮廓样式", "W6Outline" in (sc2.get("created") or []),
                  str(sc2.get("created")))
            state = sc2["state"]
        sa = server.cdr_style(doc_id, action="apply", target={"name": r2},
                              style_name="W6Fill",
                              expected_revision=state["revision"],
                              snapshot_token=state["snapshot_token"])
        if no_error("cdr_style(apply)", sa):
            check("样式可按名套用", sa.get("applied") == "W6Fill", str(sa.get("applied")))
            state = sa["state"]
        sg = server.cdr_style(doc_id, action="get", style_name="W6Fill")
        if no_error("cdr_style(get)", sg):
            check("样式定义含属性清单", (sg.get("property_count") or 0) > 10,
                  f"props={sg.get('property_count')}")
        sren = server.cdr_style(doc_id, action="rename", style_name="W6Fill",
                                new_name="W6FillRenamed",
                                expected_revision=state["revision"],
                                snapshot_token=state["snapshot_token"])
        if no_error("cdr_style(rename)", sren):
            check("样式改名成功", sren.get("renamed") is True, str(sren.get("renamed")))
            state = sren["state"]
        sdel = server.cdr_style(doc_id, action="delete", style_name="W6Outline",
                                expected_revision=state["revision"],
                                snapshot_token=state["snapshot_token"])
        if no_error("cdr_style(delete)", sdel):
            check("样式删除成功", sdel.get("deleted") is True, str(sdel.get("deleted")))
            state = sdel["state"]
        sct = server.cdr_style(doc_id, action="create", target={"name": tx1},
                               category="character", style_name="W6Char",
                               expected_revision=state["revision"],
                               snapshot_token=state["snapshot_token"])
        if no_error("cdr_style(create character)", sct):
            check("文字对象可创建字符样式", "W6Char" in (sct.get("created") or []),
                  str(sct.get("created")))
            state = sct["state"]

        # --- 内存位图：像素读取 / 调色板 / 双色调 ---
        bi6 = server.cdr_bitmap(doc_id, target={"name": bmp_name}, action="image_info")
        if no_error("cdr_bitmap(image_info)", bi6):
            check("内存位图缓冲为 16x16", bi6.get("image_size") == [16, 16],
                  str(bi6.get("image_size")))
        bp1 = server.cdr_bitmap(doc_id, target={"name": bmp_name}, action="pixel",
                                x_px=2, y_px=2)
        if no_error("cdr_bitmap(pixel fill)", bp1):
            check("像素取色读到蓝块", bp1.get("name") == "蓝"
                  or "0,0,255" in str(bp1.get("color")), str(bp1.get("color")))
        bp2 = server.cdr_bitmap(doc_id, target={"name": bmp_name}, action="pixel",
                                x_px=12, y_px=12)
        if no_error("cdr_bitmap(pixel bg)", bp2):
            check("像素取色读到红底", bp2.get("name") == "红"
                  or "255,0,0" in str(bp2.get("color")), str(bp2.get("color")))
        bp3 = server.cdr_bitmap(doc_id, target={"name": bmp_name}, action="pixel",
                                x_px=99, y_px=99)
        check("像素越界被拒绝", bp3.get("error") == "bad_argument", str(bp3.get("error")))
        bpal = server.cdr_bitmap(doc_id, target={"name": bmp_name}, action="paletted",
                                 palette_type=3, expected_revision=state["revision"],
                                 snapshot_token=state["snapshot_token"])
        if no_error("cdr_bitmap(paletted)", bpal):
            check("调色板模式转换成功（Mode=3）", bpal.get("mode") == 3,
                  f"mode={bpal.get('mode')}")
            state = bpal["state"]
        bduo = server.cdr_bitmap(doc_id, target={"name": bmp_name}, action="duotone",
                                 duotone_type=1, ink_colors=["#FF0000", "#0000FF"],
                                 expected_revision=state["revision"],
                                 snapshot_token=state["snapshot_token"])
        if no_error("cdr_bitmap(duotone)", bduo):
            check("双色调转换成功（Mode=6）", bduo.get("mode") == 6,
                  f"mode={bduo.get('mode')}")
            state = bduo["state"]

        # --- 范围操作：并行曲线 / 栅格化 ---
        pc = server.cdr_shape_ops(doc_id, "parallel_curves", target={"name": cv1},
                                  curve_count=2, curve_distance=3,
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(parallel_curves)", pc):
            check("并行曲线生成 2 条", pc.get("created") == 2, str(pc.get("created")))
            state = pc["state"]
        rz = server.cdr_shape_ops(doc_id, "rasterize", target={"name": r1},
                                  resolution=150, expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(rasterize)", rz):
            check("选择集栅格化为位图", rz.get("created") is True
                  and rz.get("type_id") == 5, f"type_id={rz.get('type_id')}")
            state = rz["state"]

        # --- 图层相邻/下移 ---
        lc = server.cdr_layers(doc_id, action="create", name="W6Layer",
                               expected_revision=state["revision"],
                               snapshot_token=state["snapshot_token"])
        if no_error("cdr_layers(create)", lc):
            state = lc["state"]
            la = server.cdr_layers(doc_id, action="above", name="W6Layer")
            if no_error("cdr_layers(above)", la):
                check("读取上方相邻图层", la.get("neighbor") is not None,
                      str(la.get("neighbor")))
            lb = server.cdr_layers(doc_id, action="below", name="W6Layer")
            if no_error("cdr_layers(below)", lb):
                check("读取下方相邻图层", lb.get("neighbor") is not None,
                      str(lb.get("neighbor")))
            lm = server.cdr_layers(doc_id, action="move_below", name="W6Layer",
                                   move_below="图层 1",
                                   expected_revision=state["revision"],
                                   snapshot_token=state["snapshot_token"])
            if no_error("cdr_layers(move_below)", lm):
                check("图层下移生效", isinstance(lm.get("layers"), list),
                      str([x["name"] for x in lm.get("layers") or []]))
                state = lm["state"]

        # --- 冻结/解冻透明度（冻结后对象变组合体并自动补名） ---
        fz = server.cdr_modify(doc_id, {"name": r2}, {
            "transparency": 40, "freeze_transparency": True},
            expected_revision=state["revision"],
            snapshot_token=state["snapshot_token"])
        if no_error("cdr_modify(freeze_transparency)", fz):
            check("冻结透明度产出可定位对象", bool(fz.get("name")),
                  f"name={fz.get('name')}")
            state = fz["state"]
            fz_name = fz.get("name")
            uz = server.cdr_modify(doc_id, {"name": fz_name},
                                   {"freeze_transparency": False},
                                   expected_revision=state["revision"],
                                   snapshot_token=state["snapshot_token"])
            if no_error("cdr_modify(unfreeze_transparency)", uz):
                check("解冻透明度可用",
                      "unfreeze_transparency" in (uz.get("applied") or []),
                      str(uz.get("applied")))
                state = uz["state"]

        # --- 字体 / 调色板 ---
        fo = server.cdr_doc(doc_id, "fonts", query="Americana", limit=5)
        if no_error("cdr_doc(fonts)", fo):
            check("字体列表可按需扫描", (fo.get("total") or 0) > 100
                  and len(fo.get("fonts") or []) >= 1,
                  f"total={fo.get('total')} scanned={fo.get('scanned')}")
        pl = server.cdr_doc(doc_id, "palettes")
        if no_error("cdr_doc(palettes)", pl):
            check("调色板列表可读", (pl.get("count") or 0) >= 1
                  and bool((pl.get("palettes") or [{}])[0].get("identifier")),
                  f"count={pl.get('count')}")
        pc2 = server.cdr_doc(doc_id, "palette_color", palette_index=1, color_index=1)
        if no_error("cdr_doc(palette_color)", pc2):
            check("调色板取色可用", bool(pc2.get("color")), str(pc2.get("color")))

        # --- OpenType 特性 ---
        ot = server.cdr_text(doc_id, {"name": tx1},
                             opentype={"liga": True, "kern": False},
                             range_start=0, range_end=5,
                             expected_revision=state["revision"],
                             snapshot_token=state["snapshot_token"])
        if no_error("cdr_text(opentype)", ot):
            check("OpenType 特性写入并回读",
                  (ot.get("opentype") or {}).get("liga") == 1
                  and (ot.get("opentype") or {}).get("kern") == 0,
                  str(ot.get("opentype")))
            state = ot["state"]

        # --- 页面插入（放最后：插到第 1 页前面会切换活动页，后续按名查对象会"找不到"） ---
        pg_before = (server.cdr_pages(doc_id, action="list").get("pages") or [])
        pgi = server.cdr_pages(doc_id, action="insert", count=1, page_index=1,
                               insert_before=True, expected_revision=state["revision"],
                               snapshot_token=state["snapshot_token"])
        if no_error("cdr_pages(insert)", pgi):
            check("插入空白页成功", (pgi.get("pages") or 0) == len(pg_before) + 1,
                  f"{len(pg_before)} -> {pgi.get('pages')}")
            state = pgi["state"]

    # ---------------- 15m. 第七波：专色/固定调色板 / 样式属性与派生 / 范围批量 / 包络 / 比较 ----------------
    print("\n[15m] 第七波：专色 + 样式属性/派生 + 范围批量 + 包络 + 吸附")
    fresh7 = server.cdr_inspect(doc_id)
    if "state" in fresh7:
        state = fresh7["state"]
    w7 = server.cdr_create(doc_id, items=[
        {"type": "rectangle", "x": 20, "y": 40, "width": 40, "height": 30,
         "fill": "#CC0000", "outline": {"width": 2, "color": "#000000"}},
        {"type": "rectangle", "x": 70, "y": 40, "width": 40, "height": 30,
         "fill": "#00CC00"},
        {"type": "bezier", "x1": 20, "y1": 90, "c1x": 40, "c1y": 110,
         "c2x": 60, "c2y": 70, "x2": 80, "y2": 90},
        {"type": "bitmap", "x": 130, "y": 40, "width": 30, "height": 30,
         "color": "#FF0000", "pixels_w": 16, "pixels_h": 16},
    ], expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建第七波测试对象", w7):
        state = w7["state"]
        r7a, r7b, c7, b7 = (w7["created"][i]["name"] for i in range(4))

        # --- 固定调色板 / 专色 ---
        opf = server.cdr_doc(doc_id, "palette_open_fixed", palette_id=3)
        palette_index = None
        if no_error("cdr_doc(palette_open_fixed=3)", opf):
            palette_index = opf.get("palette_index")
            check("固定调色板可打开（PANTONE 等）", (opf.get("color_count") or 0) > 0
                  and bool(opf.get("identifier")),
                  f"name={opf.get('name')} colors={opf.get('color_count')} index={palette_index}")
        if palette_index:
            pcs = server.cdr_doc(doc_id, "palette_colors", palette_index=palette_index,
                                 query="PANTONE", limit=5)
            if no_error("cdr_doc(palette_colors)", pcs):
                check("调色板颜色名可枚举", (pcs.get("matched") or 0) >= 1,
                      f"matched={pcs.get('matched')}")
            spc = server.cdr_doc(doc_id, "spot_color", palette_index=palette_index,
                                 color_index=1, tint=100)
            spot_ok = False
            if no_error("cdr_doc(spot_color)", spc):
                spot_ok = bool(spc.get("color")) and str(spc.get("color")).startswith("PANTONE")
                check("专色对象创建成功", spot_ok, str(spc.get("color"))[:80])
            if spot_ok:
                fsp = server.cdr_modify(doc_id, {"name": r7a}, {
                    "fill": {"type": "spot", "palette_index": palette_index,
                             "color_index": 1, "tint": 80}},
                    expected_revision=state["revision"],
                    snapshot_token=state["snapshot_token"])
                if no_error("cdr_modify(fill spot)", fsp):
                    check("专色填充已应用", any(a == "fill:spot" for a in fsp.get("applied", [])),
                          str(fsp.get("applied")))
                    state = fsp["state"]
                osp = server.cdr_modify(doc_id, {"name": r7b}, {
                    "outline": {"width": 1, "color": {"type": "spot",
                                                      "palette_index": palette_index,
                                                      "color_index": 1}}},
                    expected_revision=state["revision"],
                    snapshot_token=state["snapshot_token"])
                if no_error("cdr_modify(outline spot)", osp):
                    check("专色描边已应用", "outline" in (osp.get("applied") or []),
                          str(osp.get("applied")))
                    state = osp["state"]
            # 自定义调色板 + 位图 9 号模式
            bpal = server.cdr_bitmap(doc_id, {"name": b7}, action="paletted",
                                     palette_type=9, palette_index=palette_index,
                                     expected_revision=state["revision"],
                                     snapshot_token=state["snapshot_token"])
            if no_error("cdr_bitmap(paletted=9 自定义)", bpal):
                check("自定义调色板转换成功（Mode=3）", bpal.get("mode") == 3,
                      f"mode={bpal.get('mode')}")
                state = bpal["state"]

        # --- 样式属性 / 派生 ---
        sb = server.cdr_style(doc_id, action="create", target={"name": r7a},
                              category="fill", style_name="W7Base",
                              expected_revision=state["revision"],
                              snapshot_token=state["snapshot_token"])
        if no_error("cdr_style(create base)", sb):
            state = sb["state"]
            sg = server.cdr_style(doc_id, action="get", style_name="W7Base")
            if no_error("cdr_style(get 属性)", sg):
                check("样式可读属性可枚举",
                      len(sg.get("readable_properties") or {}) >= 10,
                      f"readable={len(sg.get('readable_properties') or {})}")
            sp = server.cdr_style(doc_id, action="set_property", style_name="W7Base",
                                  property="angle", value="15",
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
            if no_error("cdr_style(set_property)", sp):
                check("样式标量属性可写", str(sp.get("readback")).strip('"') == "15",
                      str(sp.get("readback")))
                state = sp["state"]
            sc = server.cdr_style(doc_id, action="create", category="fill",
                                  based_on="W7Base", style_name="W7Child",
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
            if no_error("cdr_style(create based_on)", sc):
                check("样式可基于已有样式派生", sc.get("based_on") == "W7Base",
                      str(sc.get("based_on")))
                state = sc["state"]
            sd = server.cdr_style(doc_id, action="derived", style_name="W7Base")
            if no_error("cdr_style(derived)", sd):
                check("派生样式可列出",
                      any(x.get("style_name") == "W7Child" for x in sd.get("derived") or []),
                      str([x.get("style_name") for x in sd.get("derived") or []]))
            sap = server.cdr_style(doc_id, action="apply", target={"name": r7b},
                                   style_name="W7Base",
                                   expected_revision=state["revision"],
                                   snapshot_token=state["snapshot_token"])
            if no_error("cdr_style(apply base)", sap):
                check("样式套用成功", sap.get("applied") == "W7Base",
                      str(sap.get("applied")))
                state = sap["state"]

        # --- 范围批量 / 包络 / 比较 ---
        cmp1 = server.cdr_shape_ops(doc_id, "compare", target={"name": r7a},
                                    other={"name": r7b}, compare_type=1,
                                    expected_revision=state["revision"],
                                    snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(compare type=1)", cmp1):
            check("对象比较（类型）返回布尔", cmp1.get("same") is True,
                  str(cmp1.get("same")))
            state = cmp1["state"]
        env = server.cdr_shape_ops(doc_id, "envelope_from_shape",
                                   target={"name": r7b}, other={"name": c7},
                                   envelope_mode=2,
                                   expected_revision=state["revision"],
                                   snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(envelope_from_shape)", env):
            check("用另一曲线做封套可用", env.get("done") is True,
                  f"effect={env.get('effect_object')}")
            state = env["state"]
        boc = server.cdr_shape_ops(doc_id, "batch_outline_to_curves",
                                   target={"name": r7a},
                                   expected_revision=state["revision"],
                                   snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(batch_outline_to_curves)", boc):
            check("选择集描边转曲可用", boc.get("done") is True)
            state = boc["state"]
        fm = server.cdr_shape_ops(doc_id, "fill_mode", target={"name": r7b},
                                  fill_mode=1,
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(fill_mode)", fm):
            check("填充规则设置可用", fm.get("fill_mode") == 1)
            state = fm["state"]
        pa = server.cdr_shape_ops(doc_id, "pixel_aligned", target={"name": r7b},
                                  pixel_aligned=True,
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(pixel_aligned)", pa):
            check("像素对齐渲染设置可用", pa.get("pixel_aligned") is True)
            state = pa["state"]
        rb = server.cdr_shape_ops(doc_id, "range_boundary", target={"name": r7b},
                                  x=80, y=50,
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(range_boundary)", rb):
            check("范围边界曲线已创建并命名",
                  rb.get("created") is True and bool(rb.get("name")),
                  f"name={rb.get('name')}")
            state = rb["state"]

        # --- 选择集移除 / 清空 ---
        sa7 = server.cdr_select(doc_id, name=r7b)
        if no_error("cdr_select(add)", sa7):
            check("选择集加入成功", (sa7.get("selected") or 0) >= 1, str(sa7.get("selected")))
        sr7 = server.cdr_select(doc_id, name=r7b, remove=True)
        if no_error("cdr_select(remove)", sr7):
            check("选择集移除成功", (sr7.get("removed") or 0) >= 1, str(sr7.get("removed")))
        scl = server.cdr_select(doc_id, clear=True)
        if no_error("cdr_select(clear)", scl):
            check("选择集清空成功", scl.get("selected") == 0, str(scl.get("selected")))

        # --- 吸附点 ---
        snap = server.cdr_pages(doc_id, action="snap_point", x=21, y=41,
                                snap_types=2, hot_area=5)
        if no_error("cdr_pages(snap_point)", snap):
            check("最近吸附点可查询", snap.get("found") is True
                  and bool(snap.get("position")), str(snap.get("position")))

    # ---------------- 15n. 第八波：图案画布填充 / 跨文档剪贴板 / 对象到图层 ----------------
    print("\n[15n] 第八波：图案画布 + 跨文档剪贴板 + 对象到图层")
    fresh8 = server.cdr_inspect(doc_id)
    if "state" in fresh8:
        state = fresh8["state"]
    w8 = server.cdr_create(doc_id, items=[
        {"type": "rectangle", "x": 20, "y": 60, "width": 40, "height": 30,
         "fill": "#CC0000", "outline": {"width": 1, "color": "#000000"}},
        {"type": "rectangle", "x": 70, "y": 60, "width": 40, "height": 30,
         "fill": "#00CC00"},
        {"type": "rectangle", "x": 120, "y": 60, "width": 30, "height": 30,
         "fill": "#0000CC"},
    ], expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建第八波测试对象", w8):
        state = w8["state"]
        p1, p2, p3 = (w8["created"][i]["name"] for i in range(3))

        # --- 图案画布 / 图案填充 / 图案透明 ---
        pcv = server.cdr_doc(doc_id, "pattern_canvas",
                             pattern_fills=[[0, 0, 15, 15], [20, 20, 31, 31]],
                             pattern_pixels=[[5, 25]])
        canvas_index = None
        if no_error("cdr_doc(pattern_canvas)", pcv):
            canvas_index = pcv.get("canvas_index")
            check("图案画布创建并登记", bool(canvas_index)
                  and pcv.get("width") == 32 and (pcv.get("drawn") or 0) == 3,
                  f"index={canvas_index} drawn={pcv.get('drawn')}")
        if canvas_index:
            pf = server.cdr_modify(doc_id, {"name": p1}, {
                "fill": {"type": "pattern", "canvas_index": canvas_index,
                         "front": "#000000", "end": "#FFFFFF"}},
                expected_revision=state["revision"],
                snapshot_token=state["snapshot_token"])
            if no_error("cdr_modify(fill pattern)", pf):
                check("图案填充已应用", "fill:pattern" in (pf.get("applied") or []),
                      str(pf.get("applied")))
                state = pf["state"]
            pt = server.cdr_modify(doc_id, {"name": p2}, {
                "transparency": {"type": "pattern", "canvas_index": canvas_index,
                                 "front": 0, "back": 100}},
                expected_revision=state["revision"],
                snapshot_token=state["snapshot_token"])
            if no_error("cdr_modify(transparency pattern)", pt):
                check("图案透明已应用", "transparency:pattern" in (pt.get("applied") or []),
                      str(pt.get("applied")))
                state = pt["state"]

        # --- 剪贴板：同文档复制粘贴 ---
        cp = server.cdr_modify(doc_id, {"name": p1}, {"copy_clipboard": True},
                               expected_revision=state["revision"],
                               snapshot_token=state["snapshot_token"])
        if no_error("cdr_modify(copy_clipboard)", cp):
            check("对象已复制到剪贴板", "copy_clipboard" in (cp.get("applied") or []),
                  str(cp.get("applied")))
            state = cp["state"]
            paste1 = server.cdr_layers(doc_id, action="paste",
                                       expected_revision=state["revision"],
                                       snapshot_token=state["snapshot_token"])
            if no_error("cdr_layers(paste)", paste1):
                check("剪贴板粘贴出对象", paste1.get("pasted") is True,
                      f"name={paste1.get('name')} type={paste1.get('type')}")
                state = paste1["state"]

        # --- 剪贴板：跨文档 ---
        doc2 = server.cdr_open(preset="A5")
        if no_error("cdr_open(跨文档目标)", doc2):
            doc2_id = doc2["document_id"]
            cross = server.cdr_layers(doc2_id, action="paste")
            if no_error("cdr_layers(paste 跨文档)", cross):
                check("跨文档粘贴成功", cross.get("pasted") is True,
                      f"name={cross.get('name')} type={cross.get('type')}")
            else:
                check("跨文档粘贴成功", False, str(cross.get("message")))
            no_error("关闭跨文档目标", server.cdr_close(doc2_id))

        # --- 剪切（用一个一次性对象验证） ---
        cut = server.cdr_modify(doc_id, {"name": p3}, {"cut_clipboard": True},
                                expected_revision=state["revision"],
                                snapshot_token=state["snapshot_token"])
        if no_error("cdr_modify(cut_clipboard)", cut):
            check("剪切到剪贴板成功", "cut_clipboard" in (cut.get("applied") or []),
                  str(cut.get("applied")))
            state = cut["state"]
            paste2 = server.cdr_layers(doc_id, action="paste",
                                       expected_revision=state["revision"],
                                       snapshot_token=state["snapshot_token"])
            if no_error("cdr_layers(paste 剪切后)", paste2):
                check("剪切内容可再粘贴", paste2.get("pasted") is True,
                      str(paste2.get("name")))
                state = paste2["state"]

        # --- 对象移动到 / 复制到图层 ---
        lnew = server.cdr_layers(doc_id, action="create", name="W8Layer",
                                 expected_revision=state["revision"],
                                 snapshot_token=state["snapshot_token"])
        if no_error("cdr_layers(create W8Layer)", lnew):
            state = lnew["state"]
            mv = server.cdr_modify(doc_id, {"name": p2}, {"move_to_layer": "W8Layer"},
                                   expected_revision=state["revision"],
                                   snapshot_token=state["snapshot_token"])
            if no_error("cdr_modify(move_to_layer)", mv):
                check("对象移动到图层", mv.get("layer") == "W8Layer",
                      f"layer={mv.get('layer')}")
                state = mv["state"]
            cpy = server.cdr_modify(doc_id, {"name": p1}, {"copy_to_layer": "W8Layer"},
                                    expected_revision=state["revision"],
                                    snapshot_token=state["snapshot_token"])
            if no_error("cdr_modify(copy_to_layer)", cpy):
                check("对象复制到图层", cpy.get("layer") == "W8Layer",
                      f"layer={cpy.get('layer')} name={cpy.get('name')}")
                state = cpy["state"]
            layers8 = server.cdr_layers(doc_id, action="list")
            if no_error("cdr_layers(list after moves)", layers8):
                target = next((x for x in layers8.get("layers", [])
                               if x.get("name") == "W8Layer"), {})
                check("目标图层对象数 >= 2", (target.get("shapes") or 0) >= 2,
                      f"shapes={target.get('shapes')}")

    # ---------------- 15o. 第九波：模板创建 / 样式表导入导出 / 颜色样式 / 纹理 / 投影 / 字体读取 ----------------
    print("\n[15o] 第九波：模板 + 样式表 + 颜色样式 + 纹理 + 3D 投影 + 字体读取")
    fresh9 = server.cdr_inspect(doc_id)
    if "state" in fresh9:
        state = fresh9["state"]
    w9 = server.cdr_create(doc_id, items=[
        {"type": "rectangle", "x": 20, "y": 140, "width": 40, "height": 30,
         "fill": "#CC0000", "outline": {"width": 2, "color": "#000000"}},
        {"type": "paragraph_text", "x": 20, "y": 200, "x2": 150, "y2": 240,
         "text": "Hyphenation sample text", "font": "Arial", "size": 12},
    ], expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建第九波测试对象", w9):
        state = w9["state"]
        q1 = w9["created"][0]["name"]
        qt = w9["created"][1]["name"]

        # --- 模板创建 ---
        tpl = os.path.join(out, "smoke.cdr")
        if os.path.exists(tpl):
            td = server.cdr_open(template=tpl, include_graphics=True)
            if no_error("cdr_open(template)", td):
                check("从模板创建文档", td.get("from_template") is True
                      and (td.get("pages") or 0) >= 1,
                      f"pages={td.get('pages')} name={td.get('name')}")
                tdoc = td.get("document_id")
                if tdoc:
                    no_error("关闭模板文档", server.cdr_close(tdoc))
        else:
            print("    [skip] smoke.cdr 不存在，跳过模板测试", flush=True)

        # --- 样式表导出 / 导入 ---
        s9 = server.cdr_style(doc_id, action="create", target={"name": q1},
                              category="fill", style_name="W9Style",
                              expected_revision=state["revision"],
                              snapshot_token=state["snapshot_token"])
        if no_error("cdr_style(create W9Style)", s9):
            state = s9["state"]
            sheet = os.path.join(out, "w9_styles.cdst")
            if os.path.exists(sheet):
                os.remove(sheet)
            se = server.cdr_style(doc_id, action="export_sheet", sheet_path=sheet)
            if no_error("cdr_style(export_sheet)", se):
                check("样式表导出成文件", se.get("exists") is True
                      and (se.get("bytes") or 0) > 1000, f"bytes={se.get('bytes')}")
            sd = server.cdr_style(doc_id, action="delete", style_name="W9Style",
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
            if no_error("cdr_style(delete W9Style)", sd):
                state = sd["state"]
                si = server.cdr_style(doc_id, action="import_sheet", sheet_path=sheet,
                                      expected_revision=state["revision"],
                                      snapshot_token=state["snapshot_token"])
                if no_error("cdr_style(import_sheet)", si):
                    check("样式表导入恢复样式", "W9Style" in (si.get("styles") or []),
                          str(si.get("styles")))
                    state = si["state"]

        # --- 颜色样式 ---
        cs = server.cdr_style(doc_id, action="create_color_styles", harmonies=1,
                              expected_revision=state["revision"],
                              snapshot_token=state["snapshot_token"])
        if no_error("cdr_style(create_color_styles)", cs):
            check("颜色样式生成成功", (cs.get("after") or 0) > (cs.get("before") or 0),
                  f"{cs.get('before')} -> {cs.get('after')}")
            state = cs["state"]
            csl = server.cdr_style(doc_id, action="color_styles")
            if no_error("cdr_style(color_styles)", csl):
                check("颜色样式可枚举", (csl.get("count") or 0) >= 1
                      and bool((csl.get("color_styles") or [{}])[0].get("color")),
                      f"count={csl.get('count')}")

        # --- 纹理列表 / 纹理填充 ---
        tex = server.cdr_doc(doc_id, "textures")
        if no_error("cdr_doc(textures)", tex):
            check("纹理库名称可枚举", (tex.get("total") or 0) >= 1,
                  f"total={tex.get('total')} first={(tex.get('textures') or [None])[0]}")
            if tex.get("textures"):
                tf = server.cdr_modify(doc_id, {"name": q1}, {
                    "fill": {"type": "texture", "name": tex["textures"][0],
                             "library": "Textures"}},
                    expected_revision=state["revision"],
                    snapshot_token=state["snapshot_token"])
                if no_error("cdr_modify(fill texture)", tf):
                    check("纹理填充已应用", "fill:texture" in (tf.get("applied") or []),
                          str(tf.get("applied")))
                    state = tf["state"]

        # --- 3D 投影 / 反投影 ---
        bb_before = None
        ins9 = server.cdr_inspect(doc_id)
        if "shapes" in ins9:
            bb_before = next((s.get("bbox") for s in ins9["shapes"]
                              if s.get("name") == q1), None)
        pj = server.cdr_shape_ops(doc_id, "project", target={"name": q1}, plane=2,
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(project)", pj):
            check("对象可投影到 3D 平面", pj.get("bbox") is not None,
                  f"bbox={pj.get('bbox')}")
            state = pj["state"]
            up = server.cdr_shape_ops(doc_id, "unproject", target={"name": q1}, plane=2,
                                      expected_revision=state["revision"],
                                      snapshot_token=state["snapshot_token"])
            if no_error("cdr_shape_ops(unproject)", up):
                bb_after = up.get("bbox") or {}
                same = (bb_before and bb_after
                        and abs((bb_before.get("x") or 0) - (bb_after.get("x") or 0)) < 0.5
                        and abs((bb_before.get("width") or 0) - (bb_after.get("width") or 0)) < 0.5)
                check("反投影恢复原包围盒", bool(same),
                      f"before={bb_before} after={bb_after}")
                state = up["state"]

        # --- 字体属性 / 断字设置读取 ---
        ft = server.cdr_text(doc_id, {"name": qt}, read_font=True, read_hyphenation=True)
        if no_error("cdr_text(read_font/read_hyphenation)", ft):
            font = ft.get("font") or {}
            hyph = ft.get("hyphenation") or {}
            check("字体属性读取正确", font.get("name") == "Arial"
                  and (font.get("size") or 0) > 0,
                  f"font={font}")
            check("断字设置可读", (hyph.get("min_word_length") or 0) > 0, f"hyph={hyph}")

    # ---------------- 15p. 第十波：文档元数据 / 对象提取成新文档 / 自定义调色板 ----------------
    print("\n[15p] 第十波：文档元数据 + 提取新文档 + 自定义调色板")
    fresh10 = server.cdr_inspect(doc_id)
    if "state" in fresh10:
        state = fresh10["state"]
    w10 = server.cdr_create(doc_id, items=[
        {"type": "rectangle", "x": 20, "y": 260, "width": 40, "height": 25,
         "fill": "#CC0000", "outline": {"width": 1, "color": "#000000"}},
        {"type": "bitmap", "x": 90, "y": 260, "width": 24, "height": 24,
         "color": "#00FF00", "pixels_w": 8, "pixels_h": 8},
    ], expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建第十波测试对象", w10):
        state = w10["state"]
        e1 = w10["created"][0]["name"]
        eb = w10["created"][1]["name"]

        # --- 文档元数据 ---
        md0 = server.cdr_doc(doc_id, "metadata")
        if no_error("cdr_doc(metadata get)", md0):
            check("文档元数据可读", isinstance(md0.get("metadata"), dict)
                  and "DocID" in (md0.get("metadata") or {}),
                  f"keys={list((md0.get('metadata') or {}).keys())[:6]}")
        md1 = server.cdr_doc(doc_id, "set_metadata", metadata={
            "Title": "W10 Smoke", "Author": "cdrX8_mcp",
            "Keywords": "mcp;coreldraw", "Notes": "wave10"},
            expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
        if no_error("cdr_doc(set_metadata)", md1):
            vals = md1.get("metadata") or {}
            check("文档元数据可写", vals.get("Title") == "W10 Smoke"
                  and vals.get("Author") == "cdrX8_mcp",
                  f"Title={vals.get('Title')} Author={vals.get('Author')}")
            state = md1["state"]

        # --- 对象提取成新文档 ---
        td = server.cdr_shape_ops(doc_id, "to_document", target={"name": e1},
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(to_document)", td):
            check("对象提取为新文档", bool(td.get("document_name"))
                  and (td.get("pages") or 0) >= 1 and (td.get("shapes") or 0) >= 1,
                  f"name={td.get('document_name')} shapes={td.get('shapes')}")
            state = td["state"]
            ndoc = td.get("document_id")
            if ndoc:
                nin = server.cdr_inspect(ndoc)
                if no_error("cdr_inspect(提取出的文档)", nin):
                    check("新文档可独立查看", (nin.get("state") or {}).get("shapes") >= 1,
                          str((nin.get("state") or {}).get("shapes")))
                # 关闭失败也不影响主流程（可能因无其他文档而报服务器错误）
                server.cdr_close(ndoc)

        # --- 自定义调色板 ---
        pal_file = os.path.join(out, "w10_palette.xml")
        if os.path.exists(pal_file):
            os.remove(pal_file)
        pal = server.cdr_doc(doc_id, "palette_create", palette_name="W10Palette",
                             palette_path=pal_file,
                             palette_colors=["#FF0000", "#00FF00", "#0000FF"],
                             overwrite=True)
        if no_error("cdr_doc(palette_create)", pal):
            check("自定义调色板已创建并落盘", bool(pal.get("palette_index"))
                  and pal.get("color_count") == 3 and pal.get("exists") is True,
                  f"index={pal.get('palette_index')} bytes={pal.get('bytes')}")
            pcol = server.cdr_doc(doc_id, "palette_colors",
                                  palette_index=pal.get("palette_index"), limit=5)
            if no_error("cdr_doc(palette_colors 自定义)", pcol):
                check("自定义调色板颜色可枚举", (pcol.get("matched") or 0) == 3,
                      f"matched={pcol.get('matched')}")
            pb = server.cdr_bitmap(doc_id, {"name": eb}, action="paletted",
                                   palette_type=9,
                                   palette_index=pal.get("palette_index"),
                                   expected_revision=state["revision"],
                                   snapshot_token=state["snapshot_token"])
            if no_error("cdr_bitmap(paletted 自定义调色板)", pb):
                check("自定义调色板可用于位图", pb.get("mode") == 3,
                      f"mode={pb.get('mode')}")
                state = pb["state"]

    # ---------------- 15q. 第十一波：对象数据字段 / 自定义虚线 / 色彩管理读取 ----------------
    print("\n[15q] 第十一波：对象数据字段 + 自定义虚线 + 色彩管理")
    fresh11 = server.cdr_inspect(doc_id)
    if "state" in fresh11:
        state = fresh11["state"]
    w11 = server.cdr_create(doc_id, items=[
        {"type": "rectangle", "x": 20, "y": 280, "width": 50, "height": 10,
         "fill": "none", "outline": {"width": 2, "color": "#000000"}},
    ], expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建第十一波测试对象", w11):
        state = w11["state"]
        q11 = w11["created"][0]["name"]

        # --- 自定义虚线 ---
        os_before = server.cdr_doc(doc_id, "outline_styles")
        os_count_before = (os_before.get("count") or 0) if "error" not in os_before else 0
        dash = server.cdr_modify(doc_id, {"name": q11}, {
            "outline": {"width": 2, "color": "#000000",
                        "dash_pattern": [2, 2, 0.5, 2]}},
            expected_revision=state["revision"],
            snapshot_token=state["snapshot_token"])
        if no_error("cdr_modify(outline dash_pattern)", dash):
            check("自定义虚线已应用", "outline" in (dash.get("applied") or []),
                  str(dash.get("applied")) + " " + str(dash.get("warnings")))
            state = dash["state"]
        os_after = server.cdr_doc(doc_id, "outline_styles")
        if no_error("cdr_doc(outline_styles)", os_after):
            check("线型样式已注册", (os_after.get("count") or 0) >= os_count_before,
                  f"{os_count_before} -> {os_after.get('count')}")

        # --- 对象数据字段 ---
        adf = server.cdr_doc(doc_id, "add_data_field", field_name="W11SKU",
                             field_format="0",
                             expected_revision=state["revision"],
                             snapshot_token=state["snapshot_token"])
        if no_error("cdr_doc(add_data_field)", adf):
            check("数据字段创建成功", (adf.get("field") or {}).get("name") == "W11SKU",
                  str(adf.get("field")))
            state = adf["state"]
        df = server.cdr_doc(doc_id, "data_fields")
        if no_error("cdr_doc(data_fields)", df):
            check("数据字段列表可读", isinstance(df.get("fields"), list),
                  f"count={df.get('count')}")
        md = server.cdr_modify(doc_id, {"name": q11},
                               {"data": {"W11SKU": "A-123"}},
                               expected_revision=state["revision"],
                               snapshot_token=state["snapshot_token"])
        if no_error("cdr_modify(data)", md):
            check("对象数据字段已写入",
                  any(a == "data:W11SKU" for a in md.get("applied", [])),
                  str(md.get("applied")))
            state = md["state"]
        od = server.cdr_doc(doc_id, "object_data", target={"name": q11})
        if no_error("cdr_doc(object_data)", od):
            item = next((x for x in od.get("data", []) if x.get("field") == "W11SKU"),
                        {})
            check("对象数据字段可读", item.get("value") == "A-123", str(item))
        upd = server.cdr_modify(doc_id, {"name": q11},
                                {"data": {"W11SKU": "B-456"}},
                                expected_revision=state["revision"],
                                snapshot_token=state["snapshot_token"])
        if no_error("cdr_modify(data update)", upd):
            state = upd["state"]
            od2 = server.cdr_doc(doc_id, "object_data", target={"name": q11})
            item2 = next((x for x in od2.get("data", []) if x.get("field") == "W11SKU"),
                         {})
            check("对象数据字段可更新", item2.get("value") == "B-456", str(item2))

        # --- 色彩管理读取 ---
        cc = server.cdr_doc(doc_id, "color_context")
        if no_error("cdr_doc(color_context)", cc):
            check("色彩管理信息可读", isinstance(cc.get("rendering_intent"), int)
                  and bool(cc.get("profile_names")),
                  f"intent={cc.get('rendering_intent')} profiles={str(cc.get('profile_names'))[:50]}")

    # ---------------- 15r. 第十二波：文档设置 / 文本区间编辑 / 自定义箭头登记 ----------------
    print("\n[15r] 第十二波：文档设置 + 文本区间编辑 + 箭头登记")
    fresh12 = server.cdr_inspect(doc_id)
    if "state" in fresh12:
        state = fresh12["state"]
    w12 = server.cdr_create(doc_id, items=[
        {"type": "paragraph_text", "x": 20, "y": 265, "x2": 150, "y2": 285,
         "text": "Alpha Beta Gamma", "font": "Arial", "size": 12},
        {"type": "bezier", "x1": 20, "y1": 285, "c1x": 40, "c1y": 295,
         "c2x": 60, "c2y": 275, "x2": 80, "y2": 285},
    ], expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建第十二波测试对象", w12):
        state = w12["state"]
        tr12 = w12["created"][0]["name"]
        bz12 = w12["created"][1]["name"]

        # --- 文档设置读写 ---
        st0 = server.cdr_doc(doc_id, "settings")
        if no_error("cdr_doc(settings)", st0):
            vals = st0.get("settings") or {}
            check("文档设置可读", isinstance(vals, dict) and "curve_precision" in vals,
                  f"keys={list(vals.keys())[:6]}")
            orig_eal = vals.get("edit_across_layers")
            orig_cp = vals.get("curve_precision")
            orig_res = vals.get("resolution")
            sset = server.cdr_doc(doc_id, "set_settings", settings={
                "edit_across_layers": not bool(orig_eal),
                "curve_precision": 40,
                "resolution": 200,
            }, expected_revision=state["revision"],
                snapshot_token=state["snapshot_token"])
            if no_error("cdr_doc(set_settings)", sset):
                sv = sset.get("settings") or {}
                check("文档设置可写", sv.get("curve_precision") == 40
                      and sv.get("resolution") == 200,
                      f"cp={sv.get('curve_precision')} res={sv.get('resolution')}")
                state = sset["state"]
                # 恢复原值
                restored = server.cdr_doc(doc_id, "set_settings", settings={
                    "edit_across_layers": bool(orig_eal),
                    "curve_precision": orig_cp if orig_cp is not None else 50,
                    "resolution": orig_res if orig_res is not None else 300,
                }, expected_revision=state["revision"],
                    snapshot_token=state["snapshot_token"])
                if "state" in restored:
                    state = restored["state"]

        # --- 文本区间编辑 ---
        trd = server.cdr_text(doc_id, {"name": tr12},
                              range_start=6, range_end=11, range_delete=True,
                              expected_revision=state["revision"],
                              snapshot_token=state["snapshot_token"])
        if no_error("cdr_text(range_delete)", trd):
            check("文本区间删除生效", "range_delete" in (trd.get("applied") or [])
                  and trd.get("text") == "Alpha Gamma",
                  f"text={trd.get('text')!r}")
            state = trd["state"]
        trc = server.cdr_text(doc_id, {"name": tr12},
                              range_start=0, range_end=5, range_copy=True,
                              expected_revision=state["revision"],
                              snapshot_token=state["snapshot_token"])
        if no_error("cdr_text(range_copy)", trc):
            check("文本区间复制生效", "range_copy" in (trc.get("applied") or []))
            state = trc["state"]
        trp = server.cdr_text(doc_id, {"name": tr12},
                              range_start=0, range_end=0, range_paste=True,
                              expected_revision=state["revision"],
                              snapshot_token=state["snapshot_token"])
        if no_error("cdr_text(range_paste)", trp):
            check("文本区间粘贴生效", "range_paste" in (trp.get("applied") or [])
                  and (trp.get("text") or "").startswith("AlphaAlpha"),
                  f"text={trp.get('text')!r}")
            state = trp["state"]

        # --- 自定义箭头登记到全局箭头库 ---
        ah = server.cdr_doc(doc_id, "arrowhead", target={"name": bz12},
                            register=True, expected_revision=state["revision"],
                            snapshot_token=state["snapshot_token"])
        if no_error("cdr_doc(arrowhead register)", ah):
            check("自定义箭头已登记", ah.get("created") is True
                  and ah.get("registered") is True and (ah.get("arrowheads") or 0) > 0,
                  f"index={ah.get('index')} count={ah.get('arrowheads')}")
            state = ah["state"]

    # ---------------- 15s. 第十三波：图案画布变换/图案微调/文本格式查询 ----------------
    print("\n[15s] 第十三波：图案画布变换 + 图案微调 + 文本格式查询")
    fresh13 = server.cdr_inspect(doc_id)
    if "state" in fresh13:
        state = fresh13["state"]
    w13 = server.cdr_create(doc_id, items=[
        {"type": "rectangle", "x": 20, "y": 150, "width": 50, "height": 20,
         "fill": "#CC0000"},
        {"type": "paragraph_text", "x": 20, "y": 180, "x2": 170, "y2": 210,
         "text": "Alpha Beta Gamma Delta", "font": "Arial", "size": 12},
    ], expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建第十三波测试对象", w13):
        state = w13["state"]
        r13 = w13["created"][0]["name"]
        t13 = w13["created"][1]["name"]

        # --- 图案画布 + 区域变换 ---
        pcv13 = server.cdr_doc(doc_id, "pattern_canvas",
                               pattern_fills=[[0, 0, 15, 15]],
                               pattern_pixels=[[20, 20]],
                               copy_areas=[[0, 0, 7, 7, 16, 16]],
                               rotate_areas=[[0, 0, 15, 15, 45]],
                               flip_areas=[[0, 0, 15, 15, 1]])
        cidx = None
        if no_error("cdr_doc(pattern_canvas 变换)", pcv13):
            cidx = pcv13.get("canvas_index")
            check("图案画布区域变换可用", bool(cidx) and (pcv13.get("drawn") or 0) >= 5,
                  f"index={cidx} drawn={pcv13.get('drawn')}")
        if cidx:
            pf13 = server.cdr_modify(doc_id, {"name": r13}, {
                "fill": {"type": "pattern", "canvas_index": cidx,
                         "front": "#000000", "end": "#FFFFFF"}},
                expected_revision=state["revision"],
                snapshot_token=state["snapshot_token"])
            if no_error("cdr_modify(fill pattern)", pf13):
                check("图案填充应用成功", "fill:pattern" in (pf13.get("applied") or []),
                      str(pf13.get("applied")))
                state = pf13["state"]
            pt13 = server.cdr_modify(doc_id, {"name": r13}, {
                "pattern_fill": {"rotation_angle": 30, "origin_x": 2,
                                 "tile_width": 8, "mirror_fill": True,
                                 "front_color": "#00FF00", "skew_angle": 5}},
                expected_revision=state["revision"],
                snapshot_token=state["snapshot_token"])
            if no_error("cdr_modify(pattern_fill 微调)", pt13):
                check("图案填充微调生效", "pattern_fill" in (pt13.get("applied") or []),
                      str(pt13.get("applied")))
                state = pt13["state"]

        # --- 文本格式查询 ---
        fq = server.cdr_text(doc_id, {"name": t13}, find_query="@bold = false")
        if no_error("cdr_text(find_query bold=false)", fq):
            fr = fq.get("find_ranges") or {}
            check("按格式查询文本区间", (fr.get("count") or 0) >= 1,
                  f"count={fr.get('count')} texts={fr.get('texts')}")
        fq2 = server.cdr_text(doc_id, {"name": t13}, find_query="@bold = true")
        if no_error("cdr_text(find_query bold=true)", fq2):
            fr2 = fq2.get("find_ranges") or {}
            check("格式查询无匹配返回 0", fr2.get("count") == 0,
                  f"count={fr2.get('count')}")

    # ---------------- 15t. 第十四波：位图内容替换 / 文本格式段 / 应用信息 / 文档回滚 ----------------
    print("\n[15t] 第十四波：位图替换 + 文本格式段 + 应用信息 + 文档回滚")
    fresh14 = server.cdr_inspect(doc_id)
    if "state" in fresh14:
        state = fresh14["state"]
    w14 = server.cdr_create(doc_id, items=[
        {"type": "bitmap", "x": 20, "y": 120, "width": 24, "height": 24,
         "color": "#FF0000", "pixels_w": 16, "pixels_h": 16},
        {"type": "paragraph_text", "x": 20, "y": 150, "x2": 170, "y2": 180,
         "text": "Alpha Beta Gamma", "font": "Arial", "size": 12},
    ], expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建第十四波测试对象", w14):
        state = w14["state"]
        b14 = w14["created"][0]["name"]
        t14 = w14["created"][1]["name"]

        # --- 位图内容替换 ---
        ri = server.cdr_bitmap(doc_id, {"name": b14}, action="replace_image",
                               image_color="#0000FF", image_pixels_w=16,
                               image_pixels_h=16,
                               image_fills=[{"x": 0, "y": 0, "width": 7,
                                             "height": 7, "color": "#00FF00"}],
                               expected_revision=state["revision"],
                               snapshot_token=state["snapshot_token"])
        if no_error("cdr_bitmap(replace_image)", ri):
            check("位图内容替换成功", (ri.get("pixels") or [0])[0] == 16
                  and ri.get("mode") == 4,
                  f"pixels={ri.get('pixels')} mode={ri.get('mode')}")
            state = ri["state"]
        pg = server.cdr_bitmap(doc_id, {"name": b14}, action="pixel",
                               x_px=1, y_px=1)
        if no_error("cdr_bitmap(pixel 替换后绿块)", pg):
            check("替换后像素为绿", pg.get("name") == "绿", str(pg.get("color")))
        pb = server.cdr_bitmap(doc_id, {"name": b14}, action="pixel",
                               x_px=10, y_px=10)
        if no_error("cdr_bitmap(pixel 替换后蓝底)", pb):
            check("替换后像素为蓝", pb.get("name") == "蓝", str(pb.get("color")))

        # --- 文本格式段 ---
        rn = server.cdr_text(doc_id, {"name": t14}, runs=True)
        if no_error("cdr_text(runs)", rn):
            runs = rn.get("runs") or []
            check("文本格式段可列出", (len(runs) >= 1
                  and (runs[0].get("font") == "Arial")
                  and (runs[0].get("size") or 0) > 0),
                  f"runs={runs[:1]}")

        # --- 应用信息 ---
        ai = server.cdr_doc(doc_id, "app_info", limit=5)
        if no_error("cdr_doc(app_info)", ai):
            check("应用信息可读", (ai.get("feature_count") or 0) > 0
                  and len(ai.get("platform_checks") or {}) == 3,
                  f"features={ai.get('feature_count')} platform={ai.get('platform_checks')}")

    # --- 文档回滚（在已保存副本上验证：修改 -> Revert -> 恢复） ---
    saved = os.path.join(out, "smoke.cdr")
    if os.path.exists(saved):
        d14 = server.cdr_open(path=saved)
        if no_error("cdr_open(回滚测试文档)", d14):
            did = d14["document_id"]
            before = len((server.cdr_inspect(did).get("shapes") or []))
            server.cdr_create(did, [{"type": "line", "x1": 0, "y1": 0,
                                     "x2": 10, "y2": 10}])
            mid = len((server.cdr_inspect(did).get("shapes") or []))
            rv = server.cdr_doc(did, "revert")
            after = len((server.cdr_inspect(did).get("shapes") or []))
            if no_error("cdr_doc(revert)", rv):
                check("文档回滚恢复到上次保存", mid == before + 1 and after == before,
                      f"{before} -> {mid} -> {after}")
            no_error("关闭回滚测试文档", server.cdr_close(did))

    # ---------------- 15u. 第十五波：立体化 / 透镜 / 效果管理 ----------------
    print("\n[15u] 第十五波：立体化 + 透镜 + 效果管理")
    fresh15 = server.cdr_inspect(doc_id)
    if "state" in fresh15:
        state = fresh15["state"]
    w15 = server.cdr_create(doc_id, items=[
        {"type": "ellipse", "x": 20, "y": 190, "width": 40, "height": 30,
         "fill": "#FFCC00", "outline": {"width": 1, "color": "#000000"}},
        {"type": "rectangle", "x": 80, "y": 190, "width": 40, "height": 30,
         "fill": "#00CCFF"},
    ], expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建第十五波测试对象", w15):
        state = w15["state"]
        e15 = w15["created"][0]["name"]
        r15 = w15["created"][1]["name"]

        # --- 立体化 ---
        ex = server.cdr_shape_ops(doc_id, "extrude", target={"name": e15},
                                  extrude_type=1, vp_type=1, vp_x=40, vp_y=205,
                                  depth=25, effect={"angle_x": 30, "shading": 1},
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(extrude)", ex):
            eff = (ex.get("effects") or {}).get("effects") or [{}]
            props = eff[0].get("props") or {}
            check("立体化效果已创建", (ex.get("effects") or {}).get("count") == 1
                  and eff[0].get("kind") == "extrude" and props.get("depth") == 25,
                  f"effects={ex.get('effects')}")
            state = ex["state"]
        exs = server.cdr_shape_ops(doc_id, "extrude_set", target={"name": e15},
                                   effect={"depth": 35},
                                   expected_revision=state["revision"],
                                   snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(extrude_set)", exs):
            eff = (exs.get("effects") or {}).get("effects") or [{}]
            check("立体化深度可微调", ((eff[0].get("props") or {}).get("depth")) == 35,
                  str(eff[0].get("props")))
            state = exs["state"]

        # --- 透镜 ---
        ln = server.cdr_shape_ops(doc_id, "lens", target={"name": r15},
                                  lens_type=1, rate=60,
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(lens)", ln):
            eff = (ln.get("effects") or {}).get("effects") or [{}]
            props = eff[0].get("props") or {}
            check("透镜效果已创建", (ln.get("effects") or {}).get("count") == 1
                  and eff[0].get("kind") == "lens" and props.get("rate") == 60,
                  f"effects={ln.get('effects')}")
            state = ln["state"]

        # --- 效果信息 / 清除 ---
        ei = server.cdr_shape_ops(doc_id, "effect_info", target={"name": r15})
        if no_error("cdr_shape_ops(effect_info)", ei):
            check("效果可枚举", (ei.get("count") or 0) == 1, f"count={ei.get('count')}")
            if "state" in ei:
                state = ei["state"]
        ce = server.cdr_shape_ops(doc_id, "clear_effects", target={"name": r15},
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(clear_effects)", ce):
            check("效果可清除", (ce.get("cleared") or 0) == 1
                  and (ce.get("effects") or 0) == 0, str(ce))
            state = ce["state"]

    # ---------------- 15v. 第十六波：位图特效（ApplyBitmapEffect 命令串） ----------------
    print("\n[15v] 第十六波：位图特效命令串")
    fresh16 = server.cdr_inspect(doc_id)
    if "state" in fresh16:
        state = fresh16["state"]
    w16 = server.cdr_create(doc_id, items=[
        {"type": "bitmap", "x": 60, "y": 260, "width": 30, "height": 30,
         "color": "#FF0000", "pixels_w": 16, "pixels_h": 16},
    ], expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建第十六波位图", w16):
        state = w16["state"]
        b16 = w16["created"][0]["name"]
        ri16 = server.cdr_bitmap(doc_id, {"name": b16}, action="replace_image",
                                 image_color="#0000FF",
                                 image_fills=[{"x": 0, "y": 0, "width": 7,
                                               "height": 7, "color": "#00FF00"}],
                                 expected_revision=state["revision"],
                                 snapshot_token=state["snapshot_token"])
        if no_error("cdr_bitmap(replace_image 前置)", ri16):
            state = ri16["state"]
        p_before = [server.cdr_bitmap(doc_id, {"name": b16}, action="pixel",
                                      x_px=px, y_px=px).get("color")
                    for px in (1, 4, 8)]
        eff16 = server.cdr_bitmap(doc_id, {"name": b16}, action="effect",
                                  effect_name="Whirlpool",
                                  effect_command="WhirlpoolEffect WhirpoolSpacing=20,"
                                                 "WhirlpoolSmear=9,WhirlpoolTwist=70,"
                                                 "WhirlpoolStreak=60,WhirpoolWarp=1,"
                                                 "WhirlpoolRandomSeed=0",
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
        if no_error("cdr_bitmap(effect Whirlpool)", eff16):
            check("位图特效命令串可执行", eff16.get("action") == "effect",
                  f"action={eff16.get('action')}")
            state = eff16["state"]
            p_after = [server.cdr_bitmap(doc_id, {"name": b16}, action="pixel",
                                         x_px=px, y_px=px).get("color")
                       for px in (1, 4, 8)]
            check("位图特效改变像素", p_before != p_after,
                  f"before={p_before} after={p_after}")

    # ---------------- 15w. 第十七波：混合（Blend） ----------------
    print("\n[15w] 第十七波：混合效果")
    fresh17 = server.cdr_inspect(doc_id)
    if "state" in fresh17:
        state = fresh17["state"]
    w17 = server.cdr_create(doc_id, items=[
        {"type": "ellipse", "x": 20, "y": 220, "width": 24, "height": 24,
         "fill": "#FF0000"},
        {"type": "ellipse", "x": 90, "y": 220, "width": 24, "height": 24,
         "fill": "#0000FF"},
    ], expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建第十七波混合对象", w17):
        state = w17["state"]
        a17 = w17["created"][0]["name"]
        b17 = w17["created"][1]["name"]
        bl = server.cdr_shape_ops(doc_id, "blend",
                                  target={"name": a17}, other={"name": b17},
                                  steps=5,
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(blend)", bl):
            eff = (bl.get("effects") or {}).get("effects") or [{}]
            props = eff[0].get("props") or {}
            check("混合效果已创建", eff[0].get("kind") == "blend"
                  and props.get("steps") == 5,
                  f"effects={bl.get('effects')}")
            state = bl["state"]
        bs = server.cdr_shape_ops(doc_id, "blend_set", target={"name": a17},
                                  effect={"steps": 8, "spacing": 2.0},
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(blend_set)", bs):
            eff = (bs.get("effects") or {}).get("effects") or [{}]
            props = eff[0].get("props") or {}
            check("混合步数可微调", props.get("steps") == 8
                  and props.get("spacing") == 2.0, str(props))
            state = bs["state"]
        bi = server.cdr_shape_ops(doc_id, "effect_info", target={"name": a17})
        if no_error("cdr_shape_ops(effect_info blend)", bi):
            check("混合效果可枚举", (bi.get("count") or 0) == 1,
                  f"count={bi.get('count')}")
            if "state" in bi:
                state = bi["state"]
        bc = server.cdr_shape_ops(doc_id, "clear_effects", target={"name": a17},
                                  expected_revision=state["revision"],
                                  snapshot_token=state["snapshot_token"])
        if no_error("cdr_shape_ops(clear_effects blend)", bc):
            check("混合效果可清除", (bc.get("cleared") or 0) == 1
                  and (bc.get("effects") or 0) == 0, str(bc))
            state = bc["state"]

    # ---------------- 15x. 第十八波：色彩配置文件列表 / 文档颜色转换 ----------------
    print("\n[15x] 第十八波：色彩配置文件 + 文档颜色转换")
    fresh18 = server.cdr_inspect(doc_id)
    if "state" in fresh18:
        state = fresh18["state"]
    w18 = server.cdr_create(doc_id, items=[
        {"type": "rectangle", "x": 20, "y": 20, "width": 30, "height": 20,
         "fill": "#FF0000"},
    ], expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建第十八波对象", w18):
        state = w18["state"]
        cp = server.cdr_doc(doc_id, "color_profiles", model=1, limit=3)
        if no_error("cdr_doc(color_profiles)", cp):
            check("色彩配置文件可列出", (cp.get("count") or 0) > 0,
                  f"count={cp.get('count')} first={(cp.get('profiles') or [{}])[0].get('name')}")
        cv = server.cdr_doc(doc_id, "convert_color", model=1,
                            expected_revision=state["revision"],
                            snapshot_token=state["snapshot_token"])
        if no_error("cdr_doc(convert_color cmyk)", cv):
            check("文档转 CMYK 成功", bool(cv.get("after"))
                  and cv.get("model") == 1, str(cv.get("after")))
            state = cv["state"]
            cc1 = server.cdr_doc(doc_id, "color_context")
            check("转换后色彩上下文为 CMYK",
                  (cc1.get("blending_color_model") == 1), str(cc1.get("profile_names")))
        cr = server.cdr_doc(doc_id, "convert_color", model=0,
                            expected_revision=state["revision"],
                            snapshot_token=state["snapshot_token"])
        if no_error("cdr_doc(convert_color rgb)", cr):
            cc2 = server.cdr_doc(doc_id, "color_context")
            check("可转回 RGB", (cc2.get("blending_color_model") == 0),
                  str(cc2.get("profile_names")))
            state = cr["state"]
        cb = server.cdr_doc(doc_id, "convert_color", profile_name="Dot Gain 15%",
                            expected_revision=state["revision"],
                            snapshot_token=state["snapshot_token"])
        if no_error("cdr_doc(convert_color gray)", cb):
            check("按配置文件转灰度成功", cb.get("model") == 2
                  and "Dot Gain" in (cb.get("after") or ""), str(cb.get("after")))
            state = cb["state"]

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
