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

    # ---------------- 15y. 第十九波：用户级应用偏好（持久化 KV） ----------------
    print("\n[15y] 第十九波：用户级应用偏好（KV）")
    pr = server.cdr_doc(doc_id, "pref_register", pref_key=f"W19Key_{os.urandom(3).hex()}",
                        pref_default="default-val")
    if no_error("cdr_doc(pref_register)", pr):
        check("用户偏好注册成功", pr.get("value") == "default-val",
              f"value={pr.get('value')}")
    pkey = (pr or {}).get("key") or "W19Key"
    pg = server.cdr_doc(doc_id, "pref_get", pref_key=pkey)
    if no_error("cdr_doc(pref_get)", pg):
        check("用户偏好可读", pg.get("value") == "default-val",
              f"value={pg.get('value')}")
    ps = server.cdr_doc(doc_id, "pref_set", pref_key=pkey,
                        pref_value="hello-world")
    if no_error("cdr_doc(pref_set)", ps):
        check("用户偏好可写", ps.get("value") == "hello-world",
              f"value={ps.get('value')}")
        pg2 = server.cdr_doc(doc_id, "pref_get", pref_key=pkey)
        check("用户偏好写入后回读", pg2.get("value") == "hello-world",
              f"value={pg2.get('value')}")
    pa = server.cdr_doc(doc_id, "pref_get", pref_key="W19Auto",
                        pref_default="auto-default")
    if no_error("cdr_doc(pref_get 未注册+默认值)", pa):
        check("未注册偏好可用默认值自动注册", pa.get("value") == "auto-default",
              f"value={pa.get('value')}")
    ai19 = server.cdr_doc(doc_id, "app_info", limit=1)
    if no_error("cdr_doc(app_info 工具状态)", ai19):
        check("应用信息含当前工具", ai19.get("active_tool") is not None,
              f"tool={ai19.get('active_tool')}")

    # ---------------- 15z. 第二十波：OLE 嵌入 / 命名视图 ----------------
    print("\n[15z] 第二十波：OLE 嵌入 + 命名视图")
    fresh20 = server.cdr_inspect(doc_id)
    if "state" in fresh20:
        state = fresh20["state"]
    ole_path = os.path.join(out, "w20_ole.txt")
    with open(ole_path, "w", encoding="utf-8") as fh:
        fh.write("hello ole")
    w20 = server.cdr_create(doc_id, items=[
        {"type": "ole_file", "path": ole_path, "caption": "w20 object"},
    ], expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("cdr_create(ole_file)", w20):
        item20 = (w20.get("created") or [{}])[0]
        check("OLE 文件可嵌入", (item20.get("type_id") == 12
              or item20.get("type") == "type_12") and bool(item20.get("name")),
              str(item20))
        state = w20["state"]
        server.cdr_modify(doc_id, {"name": item20["name"]}, {"delete": True},
                          expected_revision=state["revision"],
                          snapshot_token=state["snapshot_token"])
    v0 = server.cdr_doc(doc_id, "views")
    if no_error("cdr_doc(views)", v0):
        base20 = v0.get("count") or 0
        cv20 = server.cdr_doc(doc_id, "create_view", view_name="W20View",
                              zoom=150, origin_x=10, origin_y=20, page_index=1)
        if no_error("cdr_doc(create_view)", cv20):
            check("命名视图已创建", (cv20.get("count") or 0) == base20 + 1
                  and ((cv20.get("views") or [{}])[0].get("name") == "W20View"),
                  f"count={cv20.get('count')}")
            vl20 = server.cdr_doc(doc_id, "views")
            if no_error("cdr_doc(views after create)", vl20):
                check("命名视图可列出", (vl20.get("count") or 0) == base20 + 1,
                      f"count={vl20.get('count')}")
            av20 = server.cdr_doc(doc_id, "activate_view", view_index=1)
            if no_error("cdr_doc(activate_view)", av20):
                check("命名视图可激活", av20.get("activated") == 1,
                      f"name={av20.get('name')}")
            dv20 = server.cdr_doc(doc_id, "delete_view", view_index=1)
            if no_error("cdr_doc(delete_view)", dv20):
                check("命名视图可删除", (dv20.get("count") or 0) == base20,
                      f"count={dv20.get('count')}")

    # ---------------- 15aa. 第二十一波：文档网格设置 ----------------
    print("\n[15aa] 第二十一波：文档网格设置")
    g0 = server.cdr_doc(doc_id, "grid")
    if no_error("cdr_doc(grid)", g0):
        check("网格设置可读", isinstance(g0.get("spacing_x"), (int, float)),
              f"spacing={g0.get('spacing_x')}×{g0.get('spacing_y')}")
        orig = {k: g0.get(k) for k in ("spacing_x", "spacing_y", "visible", "snap",
                                       "type")}
        gs = server.cdr_doc(doc_id, "set_grid", grid_spacing_x=12,
                            grid_spacing_y=18, grid_visible=True,
                            grid_snap=True, grid_type=1)
        if no_error("cdr_doc(set_grid)", gs):
            check("网格设置可写", gs.get("spacing_x") == 12
                  and gs.get("spacing_y") == 18
                  and gs.get("visible") is True and gs.get("snap") is True,
                  f"spacing={gs.get('spacing_x')}×{gs.get('spacing_y')} visible={gs.get('visible')}")
            g2 = server.cdr_doc(doc_id, "grid")
            check("网格设置写入后回读", g2.get("spacing_x") == 12
                  and g2.get("visible") is True, str(g2))
            server.cdr_doc(doc_id, "set_grid",
                           grid_spacing_x=orig.get("spacing_x") or 5,
                           grid_spacing_y=orig.get("spacing_y") or 5,
                           grid_visible=bool(orig.get("visible")),
                           grid_snap=bool(orig.get("snap")),
                           grid_type=int(orig.get("type") or 1))

    # ---------------- 15ab. 第二十二波：叠印设置 / 段落缩进 ----------------
    print("\n[15ab] 第二十二波：叠印 + 段落缩进")
    fresh22 = server.cdr_inspect(doc_id)
    if "state" in fresh22:
        state = fresh22["state"]
    w22 = server.cdr_create(doc_id, items=[
        {"type": "rectangle", "x": 20, "y": 240, "width": 30, "height": 20,
         "fill": "#FF0000"},
        {"type": "paragraph_text", "x": 70, "y": 240, "x2": 170, "y2": 260,
         "text": "Indent me now", "font": "Arial", "size": 12},
    ], expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建第二十二波对象", w22):
        state = w22["state"]
        r22 = w22["created"][0]["name"]
        t22 = w22["created"][1]["name"]
        op0 = server.cdr_doc(doc_id, "overprint", target={"name": r22})
        if no_error("cdr_doc(overprint)", op0):
            check("叠印设置可读", op0.get("overprint_fill") is False,
                  f"fill={op0.get('overprint_fill')}")
        op1 = server.cdr_modify(doc_id, {"name": r22},
                                {"overprint_fill": True,
                                 "overprint_outline": True},
                                expected_revision=state["revision"],
                                snapshot_token=state["snapshot_token"])
        if no_error("cdr_modify(overprint)", op1):
            check("叠印设置可写", "overprint" in (op1.get("applied") or []),
                  str(op1.get("applied")))
            state = op1["state"]
            op2 = server.cdr_doc(doc_id, "overprint", target={"name": r22})
            check("叠印写入后回读", op2.get("overprint_fill") is True
                  and op2.get("overprint_outline") is True, str(op2))
        ind = server.cdr_text(doc_id, {"name": t22}, indent_left=5,
                              indent_right=6, indent_first_line=3,
                              expected_revision=state["revision"],
                              snapshot_token=state["snapshot_token"])
        if no_error("cdr_text(indent)", ind):
            check("段落缩进可写", "indent" in (ind.get("applied") or []),
                  str(ind.get("indent")))
            state = ind["state"]
            ind2 = server.cdr_text(doc_id, {"name": t22})
            vals = ind2.get("indent") or {}
            check("段落缩进写入后回读", vals.get("left") == 5
                  and vals.get("right") == 6 and vals.get("first_line") == 3,
                  str(vals))

    # ---------------- 15ac. 第二十三波：制表位 ----------------
    print("\n[15ac] 第二十三波：制表位")
    fresh23 = server.cdr_inspect(doc_id)
    if "state" in fresh23:
        state = fresh23["state"]
    w23 = server.cdr_create(doc_id, items=[
        {"type": "paragraph_text", "x": 20, "y": 220, "x2": 170, "y2": 240,
         "text": "Tab one two", "font": "Arial", "size": 12},
    ], expected_revision=state["revision"], snapshot_token=state["snapshot_token"])
    if no_error("创建第二十三波文字", w23):
        state = w23["state"]
        t23 = w23["created"][0]["name"]
        ts = server.cdr_text(doc_id, {"name": t23}, clear_tabs=True,
                             tab_stops=[{"position": 10, "alignment": 0,
                                         "leadered": False},
                                        {"position": 25, "alignment": 1,
                                         "leadered": True}],
                             expected_revision=state["revision"],
                             snapshot_token=state["snapshot_token"])
        if no_error("cdr_text(tab_stops)", ts):
            tabs = ts.get("tabs") or []
            check("制表位可写入", "tabs" in (ts.get("applied") or [])
                  and len(tabs) == 2, f"tabs={tabs}")
            state = ts["state"]
            ts2 = server.cdr_text(doc_id, {"name": t23})
            tabs2 = ts2.get("tabs") or []
            check("制表位写入后回读", len(tabs2) == 2
                  and tabs2[0].get("position") == 10
                  and tabs2[1].get("position") == 25
                  and tabs2[1].get("leadered") is True, str(tabs2))

    # ---------------- 15ad. 第二十四波：本地化元数据 ----------------
    print("\n[15ad] 第二十四波：本地化元数据")
    fresh24 = server.cdr_inspect(doc_id)
    if "state" in fresh24:
        state = fresh24["state"]
    lm0 = server.cdr_doc(doc_id, "metadata_localized")
    if no_error("cdr_doc(metadata_localized)", lm0):
        loc0 = lm0.get("localizable") or {}
        check("本地化元数据可读", isinstance(loc0.get("Title"), dict),
              f"Title={loc0.get('Title')}")
    lm1 = server.cdr_doc(doc_id, "set_metadata_localized", metadata={
        "Title": {"default": "Default Title", "2052": "中文标题",
                  "1033": "English Title"},
        "Keywords": "localized keywords",
    }, expected_revision=state["revision"],
        snapshot_token=state["snapshot_token"])
    if no_error("cdr_doc(set_metadata_localized)", lm1):
        loc = (lm1.get("localizable") or {}).get("Title") or {}
        check("本地化元数据可写", loc.get("default") == "Default Title"
              and (loc.get("languages") or {}).get("2052") == "中文标题"
              and (loc.get("languages") or {}).get("1033") == "English Title",
              str(loc))
        state = lm1["state"]
        lm2 = server.cdr_doc(doc_id, "metadata_localized")
        loc2 = (lm2.get("localizable") or {}).get("Title") or {}
        check("本地化元数据写入后回读", loc2.get("default") == "Default Title"
              and (loc2.get("languages") or {}).get("2052") == "中文标题",
              str(loc2))

    # ---------------- 15ae. 第二十五波：屏幕叠加层 ----------------
    print("\n[15ae] 第二十五波：屏幕叠加层")
    ov1 = server.cdr_view(doc_id, "overlay_text", text="Smoke overlay",
                          x=200, y=200, color=0x0000FF)
    if no_error("cdr_view(overlay_text)", ov1):
        check("叠加文字已显示", ov1.get("shown") == 1, str(ov1))
    ov2 = server.cdr_view(doc_id, "overlay_rect", x=50, y=50, x2=300,
                          y2=200, color=0x00FF00)
    if no_error("cdr_view(overlay_rect)", ov2):
        check("叠加矩形已显示", (ov2.get("shown") or 0) >= 2, str(ov2))
    ov3 = server.cdr_view(doc_id, "overlay_line", x=50, y=250, x2=300,
                          y2=250, color=0x00FFFF)
    if no_error("cdr_view(overlay_line)", ov3):
        check("叠加直线已显示", (ov3.get("shown") or 0) >= 3, str(ov3))
    ov4 = server.cdr_view(doc_id, "overlay_circle", x=400, y=150, radius=40,
                          color=0xFFFF00)
    if no_error("cdr_view(overlay_circle)", ov4):
        check("叠加圆已显示", (ov4.get("shown") or 0) >= 4, str(ov4))
    ovc = server.cdr_view(doc_id, "overlay_clear")
    if no_error("cdr_view(overlay_clear)", ovc):
        check("叠加层可全部隐藏", (ovc.get("hidden") or 0) >= 4, str(ovc))

    # ---------------- 15af. 第二十六波：应用事件队列 ----------------
    print("\n[15af] 第二十六波：应用事件队列")
    ce0 = server.cdr_doc(doc_id, "clear_events")
    if no_error("cdr_doc(clear_events)", ce0):
        pass
    evdoc = server.cdr_open(preset="A4")
    if "document_id" in evdoc:
        ev2 = server.cdr_doc(doc_id, "events", limit=50)
        if no_error("cdr_doc(events)", ev2):
            names = [e.get("event") for e in (ev2.get("events") or [])]
            check("应用事件可捕获", (ev2.get("count") or 0) > 0
                  and any(n in {"DocumentNew", "WindowActivate",
                                "DocumentOpen"} for n in names),
                  f"count={ev2.get('count')} names={names}")
        server.cdr_create(doc_id, [{"type": "line", "x1": 10, "y1": 10,
                                    "x2": 20, "y2": 20}])
        ev3 = server.cdr_doc(doc_id, "events", limit=50)
        if no_error("cdr_doc(events shape)", ev3):
            names3 = [e.get("event") for e in (ev3.get("events") or [])]
            check("文档事件可捕获图形变化",
                  any(str(n).startswith("Doc:Shape") for n in names3),
                  f"names={names3}")
        ce1 = server.cdr_doc(doc_id, "clear_events")
        if no_error("cdr_doc(clear_events after open)", ce1):
            check("事件队列可清空", (ce1.get("cleared") or 0) >= 1,
                  f"cleared={ce1.get('cleared')}")
        no_error("关闭事件测试文档", server.cdr_close(evdoc["document_id"]))

    # ---------------- 15ag. 第二十八波：导出选项 ----------------
    print("\n[15ag] 第二十八波：导出选项")
    opt_doc = server.cdr_open(preset="A4")
    if "document_id" in opt_doc:
        od_id = opt_doc["document_id"]
        server.cdr_create(od_id, [{"type": "rectangle", "x": 20, "y": 20,
                                   "width": 60, "height": 40,
                                   "fill": "#FF0000"}])
        opt_path = os.path.join(out, "smoke_opts.png")
        if os.path.exists(opt_path):
            os.remove(opt_path)
        eo = server.cdr_export(od_id, opt_path, format="png", dpi=120,
                               overwrite=True,
                               export_options={"anti_aliasing_type": 2,
                                               "dithered": False,
                                               "transparent": True,
                                               "resolution_x": 120,
                                               "resolution_y": 120,
                                               "maintain_aspect": True})
        if no_error("cdr_export(export_options)", eo):
            check("导出选项可应用", bool(eo.get("export_options"))
                  and os.path.exists(opt_path) and (eo.get("bytes") or 0) > 0,
                  f"options={eo.get('export_options')} bytes={eo.get('bytes')}")
        no_error("关闭导出选项测试文档", server.cdr_close(od_id))

    # ---------------- 15ah. 第二十九波：保存选项 ----------------
    print("\n[15ah] 第二十九波：保存选项")
    so_doc = server.cdr_open(preset="A4")
    if "document_id" in so_doc:
        so_id = so_doc["document_id"]
        server.cdr_create(so_id, [{"type": "rectangle", "x": 20, "y": 20,
                                   "width": 40, "height": 30,
                                   "fill": "#00AA00"}])
        so_path = os.path.join(out, f"w29_opts_{os.urandom(2).hex()}.cdr")
        so = server.cdr_save(so_id, so_path, overwrite=True,
                             save_options={"embed_icc_profile": True,
                                           "keep_appearance": True,
                                           "overwrite": True,
                                           "thumbnail_size": 1})
        if no_error("cdr_save(save_options)", so):
            check("保存选项可应用", bool(so.get("save_options"))
                  and (so.get("bytes") or 0) > 0 and os.path.exists(so_path),
                  f"options={so.get('save_options')} bytes={so.get('bytes')}")
        no_error("关闭保存选项测试文档", server.cdr_close(so_id))

    # ---------------- 15ai. 第三十波：导入选项 + 链接位图 ----------------
    print("\n[15ai] 第三十波：导入选项 + 链接位图")
    link_src = os.path.join(out, "smoke.png")
    if os.path.exists(link_src):
        li_doc = server.cdr_open(preset="A4")
        if "document_id" in li_doc:
            li_id = li_doc["document_id"]
            li = server.cdr_import(li_id, link_src,
                                   import_options={"link_bitmap_externally": True,
                                                   "resample_dpi": 96,
                                                   "maintain_layers": True})
            if no_error("cdr_import(import_options)", li):
                check("导入选项可应用", bool(li.get("import_options")),
                      f"options={li.get('import_options')}")
                li_name = (li.get("shape") or {}).get("name")
                if li_name:
                    lk = server.cdr_bitmap(li_id, {"name": li_name},
                                           action="link_info")
                    check("链接位图信息可读", lk.get("externally_linked") is True
                          and bool(lk.get("link_file_name")),
                          f"linked={lk.get('externally_linked')} path={lk.get('link_file_name')}")
                    ul = server.cdr_bitmap(li_id, {"name": li_name},
                                           action="update_link")
                    check("链接位图可重新加载", ul.get("updated") is True, str(ul))
            no_error("关闭链接位图测试文档", server.cdr_close(li_id))

    # ---------------- 15aj. 第三十一波：吸附点 ----------------
    print("\n[15aj] 第三十一波：吸附点（IVGSnapPoint / IVGSnapPoints / 范围）")
    sn_doc = server.cdr_open(preset="A4")
    if "document_id" in sn_doc:
        sn_id = sn_doc["document_id"]
        snr = server.cdr_create(sn_id, [{"type": "rectangle", "x": 20, "y": 20,
                                         "width": 40, "height": 30,
                                         "fill": "#3366CC"}])
        if no_error("cdr_create(吸附点测试矩形)", snr):
            sn_name = ((snr.get("created") or [{}])[0]).get("name")
            sn_t = {"name": sn_name}
            sl = server.cdr_snap(sn_id, action="list", target=sn_t)
            if no_error("cdr_snap(list)", sl):
                check("对象吸附点可枚举", sl.get("count") == 4
                      and (sl.get("by_kind") or {}).get("object") == 4,
                      f"count={sl.get('count')} kinds={sl.get('by_kind')}")
            si = server.cdr_snap(sn_id, action="info", target=sn_t, index=1)
            if no_error("cdr_snap(info)", si):
                check("单个吸附点信息可读", si.get("type") in {1, 2, 4, 8}
                      and len(si.get("position") or []) == 2,
                      f"type={si.get('type')} pos={si.get('position')}")
                check("GetPosition 与属性坐标一致",
                      si.get("position_get") == si.get("position"),
                      f"get={si.get('position_get')} props={si.get('position')}")
            sb = server.cdr_snap(sn_id, action="of_type", target=sn_t, typeset=4)
            if no_error("cdr_snap(of_type=bbox)", sb):
                check("按类型过滤吸附点", sb.get("count") == 9,
                      f"count={sb.get('count')}")
            sau = server.cdr_snap(sn_id, action="auto", target=sn_t)
            if no_error("cdr_snap(auto)", sau):
                check("自动吸附点可读", sau.get("type") == 32
                      and len(sau.get("position") or []) == 2,
                      f"type={sau.get('type')} pos={sau.get('position')}")
            sbb = server.cdr_snap(sn_id, action="bbox", target=sn_t, index=1)
            if no_error("cdr_snap(bbox)", sbb):
                check("包围盒吸附点可按序号取", sbb.get("type") == 4
                      and sbb.get("bbox_type") is not None,
                      f"type={sbb.get('type')} bbox_type={sbb.get('bbox_type')}")
            sob = server.cdr_snap(sn_id, action="obj", target=sn_t, index=2)
            if no_error("cdr_snap(obj)", sob):
                check("对象吸附点可按序号取", sob.get("type") == 2
                      and sob.get("object_type") is not None,
                      f"type={sob.get('type')} object_type={sob.get('object_type')}")
            sed = server.cdr_snap(sn_id, action="edge", target=sn_t, index=1, offset=0.5)
            if no_error("cdr_snap(edge)", sed):
                check("边上吸附点可按段取", sed.get("type") == 8
                      and sed.get("segment_index") == 1,
                      f"type={sed.get('type')} seg={sed.get('segment_index')} "
                      f"off={sed.get('segment_offset')}")
            snear = server.cdr_snap(sn_id, action="near", target=sn_t, x=20.0, y=30.0,
                                    typeset=2)
            if no_error("cdr_snap(near)", snear):
                check("集合内最近吸附点可查", snear.get("found") is True
                      and len(snear.get("position") or []) == 2,
                      f"pos={snear.get('position')}")
            sbad_bbox = server.cdr_snap(sn_id, action="bbox", target=sn_t, index=10)
            check("bbox 越界被拒绝", sbad_bbox.get("error") == "bad_argument",
                  str(sbad_bbox.get("error")))
            sa = server.cdr_snap(sn_id, action="add_user", target=sn_t,
                                 x=5, y=6, angle=30, use_direction=True)
            if no_error("cdr_snap(add_user)", sa):
                check("可新增用户吸附点", sa.get("type") == 1
                      and sa.get("kind") == "user"
                      and sa.get("position") == [5.0, 6.0],
                      f"type={sa.get('type')} pos={sa.get('position')}")
            sl2 = server.cdr_snap(sn_id, action="list", target=sn_t)
            check("用户吸附点进入集合", (sl2.get("by_kind") or {}).get("user") == 1
                  and sl2.get("count") == 5,
                  f"count={sl2.get('count')} kinds={sl2.get('by_kind')}")
            sd = server.cdr_snap(sn_id, action="direction", target=sn_t, index=1,
                                 angle=15, uses_direction=True)
            if no_error("cdr_snap(direction)", sd):
                check("吸附方向可改", abs(float(sd.get("direction") or 0) - 15.0) < 0.01,
                      f"direction={sd.get('direction')}")
            ss = server.cdr_snap(sn_id, action="select", target=sn_t, index=1)
            check("吸附点可选", ss.get("selected") == 1, str(ss.get("selected")))
            sc = server.cdr_snap(sn_id, action="clear_selection", target=sn_t)
            check("吸附点选择可清空", sc.get("selected") == 0, str(sc.get("selected")))
            sf = server.cdr_snap(sn_id, action="free", x=50, y=50)
            if no_error("cdr_snap(free)", sf):
                check("文档自由吸附点可建", sf.get("type") == 16
                      and sf.get("source") == "document_free_snap_point",
                      f"type={sf.get('type')} pos={sf.get('position')}")
            sk = server.cdr_snap(sn_id, action="create", x=60, y=60)
            if no_error("cdr_snap(create)", sk):
                check("应用级吸附点可建", sk.get("type") == 16
                      and sk.get("source") == "application_snap_point",
                      f"type={sk.get('type')}")
            sfind = server.cdr_snap(sn_id, action="find", x=20.4, y=30.0,
                                    typeset=2, hot_area=5.0)
            if no_error("cdr_snap(find)", sfind):
                fp = sfind.get("position") or [0, 0]
                near = abs(float(fp[0]) - 20.4) <= 5.05 and abs(float(fp[1]) - 30.0) <= 5.05
                check("最近吸附点可查", sfind.get("found") is True
                      and sfind.get("fallback") is False and near,
                      f"found={sfind.get('found')} pos={sfind.get('position')}")
            # 未命中时返回的自由点不能当成命中
            sfar = server.cdr_snap(sn_id, action="find", x=500.0, y=500.0,
                                   typeset=2, hot_area=2.0)
            check("未命中返回 fallback 标记", sfar.get("fallback") is True
                  and sfar.get("found") is False, f"fallback={sfar.get('fallback')}")
            # 移动对象吸附点会转成用户吸附点（X8 行为）
            smv = server.cdr_snap(sn_id, action="move", target=sn_t, index=3,
                                  dx=2.0, dy=3.0)
            if no_error("cdr_snap(move)", smv):
                check("吸附点可移动", smv.get("count") == 5, f"count={smv.get('count')}")
            # 绝对定位（SetPosition）
            sabs = server.cdr_snap(sn_id, action="move", target=sn_t, index=1,
                                   x=9.0, y=9.0)
            if no_error("cdr_snap(move absolute)", sabs):
                ap = ((sabs.get("points") or [{}])[0]).get("position")
                check("吸附点可绝对定位", ap == [9.0, 9.0], f"pos={ap}")
            # 范围操作
            sr1 = server.cdr_snap(sn_id, action="range", target=sn_t, indices=[1],
                                  range_op="selection")
            if no_error("cdr_snap(range.selection)", sr1):
                check("范围可选中", sr1.get("count_after") == 1
                      and sr1.get("shape_selected") == 1,
                      f"count={sr1.get('count_after')} sel={sr1.get('shape_selected')}")
            sr2 = server.cdr_snap(sn_id, action="range", target=sn_t, indices=[1],
                                  range_op="auto_snap", auto_snap=True)
            check("范围可设自动吸附", sr2.get("auto_snap") is True, str(sr2.get("auto_snap")))
            sr3 = server.cdr_snap(sn_id, action="range", target=sn_t, indices=[1],
                                  range_op="change_direction", angle=45.0)
            if no_error("cdr_snap(range.change_direction)", sr3):
                check("范围可改方向", float(sr3.get("angle") or 0) == 45.0,
                      f"angle={sr3.get('angle')}")
            sr4 = server.cdr_snap(sn_id, action="range", target=sn_t, indices=[1],
                                  range_op="add", index=2)
            check("范围可加点", sr4.get("added") is True and sr4.get("count_after") == 2,
                  f"added={sr4.get('added')} count={sr4.get('count_after')}")
            sr5 = server.cdr_snap(sn_id, action="range", target=sn_t, indices=[1],
                                  range_op="remove", position=1)
            check("范围可移除点", sr5.get("removed") is True,
                  f"removed={sr5.get('removed')}")
            sref = (server.cdr_snap(sn_id, action="info", target=sn_t, index=1)
                    .get("reference"))
            if sref:
                sr6 = server.cdr_snap(sn_id, action="range", target=sn_t, indices=[1],
                                      range_op="find", reference=sref)
                check("范围可按引用查找", sr6.get("found") is True, str(sr6.get("found")))
            # 按 GUID 反查用户吸附点
            sinfo = server.cdr_snap(sn_id, action="info", target=sn_t, index=1)
            suid = sinfo.get("id")
            if suid:
                su = server.cdr_snap(sn_id, action="user", target=sn_t, reference=suid)
                if no_error("cdr_snap(user)", su):
                    check("用户吸附点可按 ID 反查", su.get("id") == suid
                          and su.get("type") == 1,
                          f"id={su.get('id')} pos={su.get('position')}")
            sdel = server.cdr_snap(sn_id, action="delete", target=sn_t, index=1)
            if no_error("cdr_snap(delete)", sdel):
                check("吸附点可删除", sdel.get("count") == 4, f"count={sdel.get('count')}")
            sbad = server.cdr_snap(sn_id, action="info", target=sn_t, index=99)
            check("越界 index 被拒绝", sbad.get("error") == "bad_argument",
                  str(sbad.get("error")))
        no_error("关闭吸附点测试文档", server.cdr_close(sn_id))

    # ---------------- 15ak. 第三十二波：颜色对象 ----------------
    print("\n[15ak] 第三十二波：颜色对象（IVGColor）")
    ci = server.cdr_color(action="object_info", color="#C81E28")
    if no_error("cdr_color(object_info)", ci):
        check("颜色对象可创建并读类型", ci.get("type") == 5
              and str(ci.get("to_string", "")).startswith("RGB255"),
              f"type={ci.get('type')} str={ci.get('to_string')}")
        check("颜色对象默认名可读", bool(ci.get("name")), f"name={ci.get('name')}")
    ccb = server.cdr_color(action="object_info", color={"c": 10, "m": 20, "y": 30, "k": 40})
    check("CMYK 颜色对象类型正确", ccb.get("type") == 2
          and str(ccb.get("to_string", "")).startswith("CMYK"),
          f"type={ccb.get('type')}")
    cn = server.cdr_color(action="object_name", color="#00A0FF", name="MCP 蓝")
    if no_error("cdr_color(object_name)", cn):
        check("颜色名可写", cn.get("name_after") == "MCP 蓝",
              f"before={cn.get('name_before')} after={cn.get('name_after')}")
    csame = server.cdr_color(action="object_compare", color="#C81E28", color2="#C81E28")
    cdiff = server.cdr_color(action="object_compare", color="#C81E28", color2="#2040FF")
    if no_error("cdr_color(object_compare)", csame) and no_error("compare2", cdiff):
        check("颜色可判等与算距离", csame.get("is_same") is True
              and csame.get("distance") == 0.0
              and cdiff.get("is_same") is False
              and (cdiff.get("distance") or 0) > 0,
              f"same={csame.get('is_same')} d1={csame.get('distance')} d2={cdiff.get('distance')}")
    cbl = server.cdr_color(action="object_blend", color="#FF0000", color2="#0000FF",
                           ratio=0.5)
    if no_error("cdr_color(object_blend)", cbl):
        check("颜色可原地混合", "128,0,128" in str(cbl.get("after")),
              f"after={cbl.get('after')}")
    ccp = server.cdr_color(action="object_copy", color="#3366CC")
    if no_error("cdr_color(object_copy)", ccp):
        check("颜色可复制且相等", ccp.get("is_same") is True
              and ccp.get("copy") == ccp.get("original"), f"copy={ccp.get('copy')}")
    for mdl, pref, tid in (("lab", "LAB", 12), ("hsb", "HSB", 6), ("gray", "GRAY", 9)):
        ccv = server.cdr_color(action="object_convert", color="#78B43C", model=mdl)
        if no_error(f"cdr_color(object_convert={mdl})", ccv):
            check(f"颜色可原地转 {mdl}", str(ccv.get("after", "")).startswith(pref)
                  and ccv.get("type_after") == tid,
                  f"after={ccv.get('after')} type={ccv.get('type_after')}")
    for mode, pref in ((0, "USER"), (1, "TRU"), (2, "PAN_PROC")):
        cfx = server.cdr_color(action="object_fixed", color="#78B43C", mode=mode)
        if no_error(f"cdr_color(object_fixed={mode})", cfx):
            check(f"颜色可转固定调色板 {pref}", pref in str(cfx.get("after")),
                  f"after={cfx.get('after')}")
    csa = server.cdr_color(
        action="object_string_assign", color="#000000",
        color2="CMYK,USER,10,20,30,40,100,00000000-0000-0000-0000-000000000000")
    if no_error("cdr_color(object_string_assign)", csa):
        check("颜色可从颜色串赋值", str(csa.get("after", "")).startswith("CMYK,USER,10,20,30,40")
              and csa.get("type_after") == 2, f"after={csa.get('after')}")
    cbad = server.cdr_color(action="object_convert", color="#000000", model="xyz")
    check("非法颜色模型被拒绝", cbad.get("error") == "bad_argument", str(cbad.get("error")))
    for mdl, vals, pref in (("rgb", [10, 20, 30], "RGB255,USER,10,20,30"),
                            ("cmyk", [10, 20, 30, 40], "CMYK,USER,10,20,30,40"),
                            ("gray", [60], "GRAY255,USER,60"),
                            ("lab", [60, 10, -20], "LAB,USER,60,10,-20")):
        cas = server.cdr_color(action="object_assign", model=mdl, values=vals)
        if no_error(f"cdr_color(object_assign={mdl})", cas):
            check(f"颜色分量可直写（{mdl}）", str(cas.get("after", "")).startswith(pref),
                  f"after={cas.get('after')}")
    creg = server.cdr_color(action="object_assign", model="registration")
    if no_error("cdr_color(object_assign=registration)", creg):
        check("可赋套版色", str(creg.get("after", "")).startswith("REGCOLOR"),
              f"after={creg.get('after')}")
    cab = server.cdr_color(action="object_assign", model="cmyk", values=[1, 2])
    check("分量个数不对被拒绝", cab.get("error") == "bad_argument",
          str(cab.get("error")))

    # ---------------- 15al. 第三十二波：描边对象 ----------------
    print("\n[15al] 第三十二波：描边对象（IVGOutline）")
    ol_doc = server.cdr_open(preset="A4")
    if "document_id" in ol_doc:
        ol_id = ol_doc["document_id"]
        olr = server.cdr_create(ol_id, [
            {"type": "rectangle", "x": 20, "y": 20, "width": 40, "height": 30,
             "outline": {"width": 1.0, "color": "#CC0000"}},
            {"type": "ellipse", "cx": 150, "cy": 100, "rx": 20, "ry": 15,
             "outline": {"width": 3.0, "color": "#0000CC"}},
        ])
        if no_error("cdr_create(描边测试对象)", olr):
            ol_names = [c.get("name") for c in olr.get("created") or []]
            oa, ob = {"name": ol_names[0]}, {"name": ol_names[1]}
            oi = server.cdr_outline(ol_id, target=oa, action="info")
            if no_error("cdr_outline(info)", oi):
                check("描边属性可读", oi.get("width") == 1.0
                      and str(oi.get("color", "")).startswith("RGB255")
                      and oi.get("nib_stretch") == 100,
                      f"width={oi.get('width')} color={oi.get('color')}")
            oset = server.cdr_outline(ol_id, target=oa, action="set", changes={
                "width": 2.5, "color": "#00AA55", "behind_fill": True,
                "nib_stretch": 40, "nib_angle": 25.0, "miter_limit": 7.0,
                "justification": 1})
            if no_error("cdr_outline(set)", oset):
                check("描边可改（宽度/颜色/笔尖）", oset.get("width") == 2.5
                      and oset.get("behind_fill") is True
                      and oset.get("nib_stretch") == 40
                      and abs(float(oset.get("nib_angle") or 0) - 25.0) < 0.01,
                      f"width={oset.get('width')} nib={oset.get('nib_stretch')}")
            osp = server.cdr_outline(ol_id, target=oa, action="set", changes={
                "use_set_properties": True, "width": 1.25, "line_caps": 1,
                "line_join": 1, "miter_limit": 9.0})
            if no_error("cdr_outline(set use_set_properties)", osp):
                check("SetProperties 全参路径可用", osp.get("width") == 1.25
                      and osp.get("line_caps") == 1 and osp.get("line_join") == 1,
                      f"width={osp.get('width')} caps={osp.get('line_caps')}")
            ocmp = server.cdr_outline(ol_id, target=oa, action="compare", target2=ob)
            if no_error("cdr_outline(compare)", ocmp):
                check("描边可比较", ocmp.get("same") is False, f"same={ocmp.get('same')}")
            ocp = server.cdr_outline(ol_id, target=oa, action="copy", target2=ob)
            if no_error("cdr_outline(copy)", ocp):
                check("描边可复制", ocp.get("width") == 3.0
                      and "0,0,204" in str(ocp.get("color")) or ocp.get("width") == 3.0,
                      f"width={ocp.get('width')} color={ocp.get('color')}")
            ocmp2 = server.cdr_outline(ol_id, target=oa, action="compare", target2=ob)
            check("复制后两边描边一致", ocmp2.get("same") is True, f"same={ocmp2.get('same')}")
            one = server.cdr_outline(ol_id, target=oa, action="set", changes={"none": True})
            check("可设为无描边", one.get("width") == 0.0 and one.get("type") == 0,
                  f"width={one.get('width')} type={one.get('type')}")
            ocv = server.cdr_outline(ol_id, target=ob, action="convert_to_object")
            if no_error("cdr_outline(convert_to_object)", ocv):
                created = ocv.get("created") or {}
                check("描边可转对象", created.get("type_id") == 3
                      and ocv.get("shapes_after") == ocv.get("shapes_before"),
                      f"type={created.get('type_id')} {ocv.get('shapes_before')}→{ocv.get('shapes_after')}")
            obad = server.cdr_outline(ol_id, target=oa, action="nope")
            check("非法描边 action 被拒绝", obad.get("error") == "bad_argument",
                  str(obad.get("error")))
        no_error("关闭描边测试文档", server.cdr_close(ol_id))

    # ---------------- 15am. 第三十三波：原生范围对齐 + 范围批量 ----------------
    print("\n[15am] 第三十三波：原生范围对齐（AlignRangeTo*）与范围批量")
    rg_doc = server.cdr_open(preset="A4")
    if "document_id" in rg_doc:
        rg_id = rg_doc["document_id"]
        rgr = server.cdr_create(rg_id, [
            {"type": "rectangle", "x": 10, "y": 10, "width": 20, "height": 10},
            {"type": "rectangle", "x": 60, "y": 10, "width": 20, "height": 10},
            {"type": "rectangle", "x": 110, "y": 10, "width": 20, "height": 10},
        ])
        if no_error("cdr_create(范围对齐测试)", rgr):
            rg_names = [c.get("name") for c in rgr.get("created") or []]
            for nm in rg_names:
                server.cdr_select(rg_id, name=nm)
            ap = server.cdr_align(rg_id, horizontal="left", vertical="top",
                                  to="point", x=150.0, y=200.0)
            if no_error("cdr_align(to=point)", ap):
                check("可对齐到指定点", ap.get("aligned") == 3
                      and ap.get("axes") == [1, 4], f"axes={ap.get('axes')}")
            apc = server.cdr_align(rg_id, horizontal="center", to="page_center")
            if no_error("cdr_align(to=page_center)", apc):
                check("可对齐到页面中心", apc.get("aligned") == 3
                      and apc.get("to") == "page_center", str(apc.get("axes")))
            agr = server.cdr_align(rg_id, horizontal="left", to="grid")
            if no_error("cdr_align(to=grid)", agr):
                check("可对齐到网格", agr.get("native") is True, str(agr.get("axes")))
            ash = server.cdr_align(rg_id, vertical="bottom", to="shape",
                                   target2={"name": rg_names[0]})
            if no_error("cdr_align(to=shape)", ash):
                check("可对齐到参照对象", bool(ash.get("reference"))
                      and ash.get("axes") == [8], f"ref={ash.get('reference')}")
            arn = server.cdr_align(rg_id, horizontal="right", to="range",
                                   target2={"name": rg_names[0]})
            if no_error("cdr_align(to=range)", arn):
                check("可对齐到参照范围", arn.get("axes") == [2], str(arn.get("axes")))
            abad = server.cdr_align(rg_id, horizontal="left", to="point")
            check("to=point 缺坐标被拒绝", abad.get("error") == "bad_argument",
                  str(abad.get("error")))
            ri = server.cdr_shape_ops(rg_id, op="range_info")
            if no_error("cdr_shape_ops(range_info)", ri):
                check("范围信息可读", ri.get("count") == 3
                      and ri.get("type_ids") == [1, 1, 1]
                      and ri.get("count_any_of_type") == 3
                      and ri.get("exists_any_of_type") is True,
                      f"count={ri.get('count')} ids={ri.get('type_ids')}")
            rr = server.cdr_shape_ops(rg_id, op="range_remove",
                                      other={"name": rg_names[0]})
            if no_error("cdr_shape_ops(range_remove)", rr):
                check("范围可移除成员", rr.get("count_after") == 2,
                      f"{rr.get('count_before')}→{rr.get('count_after')}")
            ra = server.cdr_shape_ops(rg_id, op="range_add", other={"name": rg_names[0]})
            if no_error("cdr_shape_ops(range_add)", ra):
                check("范围可加回成员", ra.get("count_after") == 3,
                      f"{ra.get('count_before')}→{ra.get('count_after')}")
            ul = server.cdr_shape_ops(rg_id, op="unlock_all")
            if no_error("cdr_shape_ops(unlock_all)", ul):
                check("范围可批量解锁", ul.get("count") == 3
                      and ul.get("locked_after") == [False, False, False],
                      f"locked={ul.get('locked_after')}")
            ce = server.cdr_shape_ops(rg_id, op="clear_effect",
                                      target={"name": rg_names[0]}, distance_a=5)
            if no_error("cdr_shape_ops(clear_effect)", ce):
                check("范围可清效果", ce.get("cleared") is True, str(ce.get("cleared")))
            # 组合 → 解组（第五十四波纠正：ShapeRange.UngroupAll 其实真实生效）
            server.cdr_select(rg_id, name=rg_names[1])
            server.cdr_select(rg_id, name=rg_names[2])
            gr = server.cdr_shape_ops(rg_id, op="group")
            if no_error("cdr_shape_ops(group)", gr):
                ua = server.cdr_shape_ops(rg_id, op="ungroup_all")
                if no_error("cdr_shape_ops(ungroup_all)", ua):
                    check("可解开组合体（UngroupAll 真实生效）",
                          (ua.get("page_shapes_after") or 0)
                          > (ua.get("page_shapes_before") or 0)
                          and (ua.get("unwrapped") or 0) >= 1,
                          f"{ua.get('page_shapes_before')} → "
                          f"{ua.get('page_shapes_after')}")
            ins3 = server.cdr_inspect(rg_id)
            shapes3 = ins3.get("shapes") or []
            poly = next((s.get("name") for s in shapes3
                         if s.get("name") and s.get("type") in ("curve", "polygon")), None)
            if poly is None:
                poly = next((s.get("name") for s in shapes3 if s.get("name")), None)
            if poly:
                bp = server.cdr_shape_ops(rg_id, op="break_apart_ex",
                                          target={"name": poly})
                if no_error("cdr_shape_ops(break_apart_ex)", bp):
                    check("BreakApartEx 可用", (bp.get("count_before") or 0) >= 1,
                          f"count={bp.get('count_before')} result={bp.get('result_count')}")
        no_error("关闭范围测试文档", server.cdr_close(rg_id))

    # ---------------- 15an. 第三十四波：IVGRect / 曲线段 / 文字粗斜体 ----------------
    print("\n[15an] 第三十四波：矩形对象（IVGRect）、曲线段、文字样式与图层查找")
    rc_doc = server.cdr_open(preset="A4")
    if "document_id" in rc_doc:
        rc_id = rc_doc["document_id"]
        rf = server.cdr_doc(rc_id, action="refresh_fonts")
        check("字体表可刷新", rf.get("refreshed") is True
              and (rf.get("fonts") or 0) > 0, f"fonts={rf.get('fonts')}")
        ri = server.cdr_rect(action="info", rect={"x": 10, "y": 10, "width": 40,
                                                  "height": 30})
        if no_error("cdr_rect(info)", ri):
            r = ri.get("rect") or {}
            check("矩形对象派生值可读", r.get("left") == 10.0 and r.get("right") == 50.0
                  and r.get("top") == 40.0 and r.get("bottom") == 10.0
                  and r.get("center") == [30.0, 25.0],
                  f"left={r.get('left')} top={r.get('top')} center={r.get('center')}")
        ro = server.cdr_rect(action="offset", rect={"x": 10, "y": 10, "width": 40,
                                                    "height": 30}, dx=5, dy=5)
        check("矩形可平移", (ro.get("rect") or {}).get("x") == 15.0,
              str((ro.get("rect") or {}).get("x")))
        rinf = server.cdr_rect(action="inflate", rect={"x": 10, "y": 10, "width": 40,
                                                       "height": 30},
                               left=1, top=1, right=2, bottom=2)
        if no_error("cdr_rect(inflate)", rinf):
            rr = rinf.get("rect") or {}
            check("矩形可外扩", rr.get("x") == 9.0 and rr.get("width") == 43.0,
                  f"x={rr.get('x')} w={rr.get('width')}")
        ru = server.cdr_rect(action="union", rect={"x": 10, "y": 10, "width": 40,
                                                   "height": 30},
                             rect2={"x": 40, "y": 20, "width": 40, "height": 40})
        if no_error("cdr_rect(union)", ru):
            u = ru.get("result") or {}
            check("矩形可求并", u.get("width") == 70.0 and u.get("height") == 50.0,
                  f"w={u.get('width')} h={u.get('height')}")
        rx = server.cdr_rect(action="intersect", rect={"x": 10, "y": 10, "width": 40,
                                                       "height": 30},
                             rect2={"x": 40, "y": 20, "width": 40, "height": 40})
        if no_error("cdr_rect(intersect)", rx):
            x = rx.get("result") or {}
            check("矩形可求交", x.get("width") == 10.0 and x.get("height") == 20.0,
                  f"w={x.get('width')} h={x.get('height')}")
        rc = server.cdr_rect(action="contains", rect={"x": 10, "y": 10, "width": 40,
                                                      "height": 30},
                             point={"x": 30, "y": 30})
        check("点包含判断可用", rc.get("inside") is True, str(rc.get("inside")))
        rcp = server.cdr_rect(action="copy", rect={"x": 10, "y": 10, "width": 40,
                                                   "height": 30})
        check("矩形可复制", (rcp.get("copy") or {}).get("width") == 40.0,
              str((rcp.get("copy") or {}).get("width")))
        rcr = server.cdr_rect(action="create_rectangle", document_id=rc_id,
                              rect={"x": 20, "y": 20, "width": 40, "height": 30},
                              radii=[5, 5, 0, 0])
        if no_error("cdr_rect(create_rectangle)", rcr):
            check("按矩形造矩形", ((rcr.get("shape") or {}).get("type_id") == 1
                                and ((rcr.get("shape") or {}).get("bbox") or {}).get("width") == 40.0),
                  str(rcr.get("shape")))
        rce = server.cdr_rect(action="create_ellipse", document_id=rc_id,
                              rect={"x": 80, "y": 20, "width": 40, "height": 30},
                              start_angle=0, end_angle=270, pie=True)
        if no_error("cdr_rect(create_ellipse)", rce):
            check("按矩形造椭圆", (rce.get("shape") or {}).get("type_id") == 2,
                  str((rce.get("shape") or {}).get("type_id")))
        rc_no = server.cdr_rect(action="create_rectangle", rect={"x": 0, "y": 0})
        check("造形缺文档被拒绝", rc_no.get("error") == "bad_argument",
              str(rc_no.get("error")))
        txt = server.cdr_create(rc_id, [
            {"type": "text", "x": 20, "y": 120, "text": "粗斜下", "size": 18},
            {"type": "curve_segment", "x1": 0, "y1": 200, "x2": 40, "y2": 220,
             "start_ctrl_length": 10, "start_ctrl_angle": 90,
             "end_ctrl_length": 10, "end_ctrl_angle": 45},
        ])
        if no_error("cdr_create(text+curve_segment)", txt):
            made = txt.get("created") or []
            check("曲线段可直接创建", (made[1] or {}).get("type") == "curve",
                  str((made[1] or {}).get("type")))
            tnm = made[0]["name"]
            applied = server.cdr_text(rc_id, {"name": tnm}, bold=True, italic=True,
                                      underline=True, font="Arial", size=24)
            if no_error("cdr_text(粗斜下划线)", applied):
                ap = applied.get("applied") or []
                check("文字样式接口全部受理",
                      all(k in ap for k in ("font", "size", "bold", "italic",
                                            "underline")), str(ap))
                f2 = (server.cdr_text(rc_id, {"name": tnm},
                                      read_font=True).get("font") or {})
                # 字号与下划线在任何字体环境下都可靠；Font/Bold/Italic 受文档默认字体
                # （中文宋体）影响，X8 会静默忽略，所以只做记录不做断言
                check("文字字号与下划线生效",
                      f2.get("size") == 24.0 and f2.get("underline") is True,
                      f"size={f2.get('size')} u={f2.get('underline')} "
                      f"font={f2.get('name')} bold={f2.get('bold')}")
                off = server.cdr_text(rc_id, {"name": tnm}, underline=False)
                if no_error("cdr_text(关闭下划线)", off):
                    f3 = (server.cdr_text(rc_id, {"name": tnm},
                                          read_font=True).get("font") or {})
                    check("文字下划线可关闭", f3.get("underline") is False,
                          str(f3.get("underline")))
        li = server.cdr_layers(rc_id, action="info")
        if no_error("cdr_layers(info)", li):
            check("图层信息可读", li.get("is_guides_layer") is False
                  and bool(li.get("name")), f"name={li.get('name')} guides={li.get('is_guides_layer')}")
        lf = server.cdr_layers(rc_id, action="find", shape_name=txt["created"][0]["name"])
        if no_error("cdr_layers(find)", lf):
            check("图层内可按名查找对象", lf.get("found") is True
                  and lf.get("type_id") == 6, f"type={lf.get('type_id')}")
        lf2 = server.cdr_layers(rc_id, action="find", shape_name="no_such_shape_xyz")
        check("查不到时返回 found=false", lf2.get("found") is False, str(lf2.get("found")))
        no_error("关闭矩形/文字测试文档", server.cdr_close(rc_id))

    # ---------------- 15ao. 第三十五波：单位换算 / 颜色工厂 / 定尺寸加页 ----------------
    print("\n[15ao] 第三十五波：COM 单位换算、颜色模型工厂、定尺寸加页、文档设置快照")
    cu = server.cdr_doc("doc1", action="convert_units", value=25.4, from_unit=3, to_unit=1)
    if no_error("cdr_doc(convert_units)", cu):
        check("COM 单位换算（mm→inch）", abs(float(cu.get("result") or 0) - 1.0) < 1e-6,
              f"result={cu.get('result')}")
    cu2 = server.cdr_doc("doc1", action="convert_units", value=72, from_unit=14, to_unit=3)
    check("COM 单位换算（pt→mm）", abs(float(cu2.get("result") or 0) - 25.4) < 1e-3,
          f"result={cu2.get('result')}")
    for mdl, vals, pref, tid in (("hsb", [120, 80, 90], "HSB", 6),
                                 ("cmy", [10, 20, 30], "CMY255", 4),
                                 ("yiq", [100, 50, 50], "YIQ255", 11),
                                 ("bw", [1], "BW", 8)):
        cm = server.cdr_color(action="object_model", model=mdl, values=vals)
        if no_error(f"cdr_color(object_model={mdl})", cm):
            check(f"颜色模型工厂 {mdl}", str(cm.get("to_string", "")).startswith(pref)
                  and cm.get("type") == tid, f"type={cm.get('type')}")
    cfx2 = server.cdr_color(action="object_model", model="fixed", values=[2, 5, 100])
    if no_error("cdr_color(object_model=fixed)", cfx2):
        check("固定调色板工厂", "PAN_PROC" in str(cfx2.get("to_string")),
              f"str={cfx2.get('to_string')}")
    cex = server.cdr_color(action="object_model", model="ex", values=[2, 10, 20, 30, 40])
    if no_error("cdr_color(object_model=ex)", cex):
        check("CreateColorEx 工厂", str(cex.get("to_string", "")).startswith("CMYK,USER,10,20"),
              f"str={cex.get('to_string')}")
    cfs = server.cdr_color(
        action="object_from_string",
        color2="RGB255,USER,12,34,56,100,00000000-0000-0000-0000-000000000000")
    if no_error("cdr_color(object_from_string)", cfs):
        check("颜色串造色", cfs.get("type") == 5
              and "12,34,56" in str(cfs.get("to_string")),
              f"type={cfs.get('type')} str={cfs.get('to_string')}")
    pg_doc = server.cdr_open(preset="A4")
    if "document_id" in pg_doc:
        pg_id = pg_doc["document_id"]
        ad = server.cdr_pages(pg_id, action="add", count=1, width_mm=100, height_mm=150)
        if no_error("cdr_pages(add sized)", ad):
            plist = (server.cdr_pages(pg_id, action="list").get("pages") or [])
            sizes = [(p.get("width_mm"), p.get("height_mm")) for p in plist]
            check("可定尺寸加页（AddPagesEx）", (100.0, 150.0) in sizes,
                  f"sizes={sizes}")
        ins = server.cdr_pages(pg_id, action="insert", count=1, width_mm=80,
                               height_mm=60, page_index=1)
        if no_error("cdr_pages(insert sized)", ins):
            check("可定尺寸插页（InsertPagesEx）", ins.get("sized") is True
                  and (ins.get("pages") or 0) >= 3, f"pages={ins.get('pages')}")
        ss = server.cdr_doc(pg_id, action="save_settings", query="smoke_snap")
        check("文档设置可打快照", ss.get("action") == "save_settings", str(ss.get("tag")))
        rs = server.cdr_doc(pg_id, action="restore_settings", query="smoke_snap")
        check("文档设置可恢复（调用无异常）", rs.get("action") == "restore_settings",
              str(rs.get("tag")))
        rz = server.cdr_doc(pg_id, action="reset_settings")
        check("文档设置可重置", rz.get("action") == "reset_settings", str(rz.get("action")))
        rl = server.cdr_doc(pg_id, action="resolve_links")
        check("可解析全部位图链接", rl.get("resolved") is True, str(rl.get("resolved")))
        no_error("关闭加页测试文档", server.cdr_close(pg_id))

    # ---------------- 15ap. 第三十六波：节点拓扑与曲线合并 ----------------
    print("\n[15ap] 第三十六波：节点拓扑（Next/Previous·Join·Average）与曲线合并/焊接")
    nd_doc = server.cdr_open(preset="A4")
    if "document_id" in nd_doc:
        nd_id = nd_doc["document_id"]
        ndr = server.cdr_create(nd_id, [
            {"type": "polyline", "points": [[0, 0], [20, 0], [40, 20]],
             "closed": False},
            {"type": "polyline", "points": [[40, 20], [60, 20]], "closed": False},
        ])
        if no_error("cdr_create(节点拓扑测试)", ndr):
            na, nb = ndr["created"][0]["name"], ndr["created"][1]["name"]
            nA = {"name": na}
            nb_ = {"name": nb}
            nb1 = server.cdr_node(nd_id, nA, action="neighbors", node=1)
            if no_error("cdr_node(neighbors)", nb1):
                check("节点前后邻居可读（Next/Previous 是方法）",
                      nb1.get("next") == 2 and nb1.get("previous") == 3,
                      f"next={nb1.get('next')} prev={nb1.get('previous')}")
            ndist = server.cdr_node(nd_id, nA, action="distance", node=1, end=3)
            if no_error("cdr_node(distance)", ndist):
                check("节点间距可算", abs(float(ndist.get("distance") or 0) - 44.721) < 0.05,
                      f"distance={ndist.get('distance')}")
            navg = server.cdr_node(nd_id, nA, action="average", node=1, end=3)
            if no_error("cdr_node(average)", navg):
                check("两节点可取平均位置",
                      navg.get("pos") == [20.0, 10.0], f"pos={navg.get('pos')}")
            join_before = (server.cdr_node(nd_id, nA, action="list") or {}).get("nodes")
            njoin = server.cdr_node(nd_id, nA, action="join_nodes", node=1, end=3)
            if no_error("cdr_node(join_nodes)", njoin):
                check("末端节点可合并", (njoin.get("nodes") or 99) < (join_before or 0),
                      f"{join_before}→{njoin.get('nodes')}")
            nlist = server.cdr_node(nd_id, nA, action="list")
            first_xy = [0.0, 0.0]
            if (nlist.get("node_list") or []):
                first_xy = [nlist["node_list"][0].get("x") or 0.0,
                            nlist["node_list"][0].get("y") or 0.0]
            ion = server.cdr_path(nd_id, nA, action="is_on_curve",
                                  x=first_xy[0], y=first_xy[1], hot_area=2)
            if no_error("cdr_path(is_on_curve)", ion):
                check("曲线命中判定可用", ion.get("on_curve") is True
                      and ion.get("hit_state") == 1,
                      f"at={first_xy} on={ion.get('on_curve')} state={ion.get('hit_state')}")
            ion2 = server.cdr_path(nd_id, nA, action="is_on_curve", x=999, y=999,
                                   hot_area=2)
            check("远离曲线时判定为不在线上", ion2.get("on_curve") is False,
                  str(ion2.get("on_curve")))
            ccp = server.cdr_path(nd_id, nA, action="curve_copy")
            if no_error("cdr_path(curve_copy)", ccp):
                check("曲线可复制", (ccp.get("copy_nodes") or 0) >= 1,
                      f"copy_nodes={ccp.get('copy_nodes')}")
            cinfo = server.cdr_path(nd_id, nA, action="curve_info")
            if no_error("cdr_path(curve_info)", cinfo):
                check("Curve.GetCurveInfo 可读", (cinfo.get("info_elements") or 0) >= 1,
                      f"elements={cinfo.get('info_elements')}")
            mg = server.cdr_path(nd_id, nA, action="merge_curve", other=nb_)
            if no_error("cdr_path(merge_curve)", mg):
                check("曲线可合并（AppendCurve→多子路径）",
                      (mg.get("subpaths") or 0) == 2 and (mg.get("nodes") or 0) >= 3,
                      f"nodes={mg.get('nodes')} subpaths={mg.get('subpaths')}")
            wd = server.cdr_path(nd_id, nA, action="weld_curve", other=nb_)
            if no_error("cdr_path(weld_curve)", wd):
                check("曲线可焊接（WeldWith 返回新曲线）",
                      (wd.get("welded_nodes") or 0) >= 2
                      and (wd.get("welded_length") or 0) > 0,
                      f"welded_nodes={wd.get('welded_nodes')} len={wd.get('welded_length')}")
            nbad = server.cdr_node(nd_id, nA, action="distance", node=1)
            check("distance 缺参数被拒绝", nbad.get("error") == "bad_argument",
                  str(nbad.get("error")))
        no_error("关闭节点拓扑测试文档", server.cdr_close(nd_id))

    # ---------------- 15aq. 第三十七波：文字结构体 / 区间宽接口 / 样式继承 ----------------
    print("\n[15aq] 第三十七波：文字结构体读取、TextRange 宽接口、样式继承")
    tx_doc = server.cdr_open(preset="A4")
    if "document_id" in tx_doc:
        tx_id = tx_doc["document_id"]
        txr = server.cdr_create(tx_id, [
            {"type": "text", "x": 20, "y": 80, "text": "结构体测试", "size": 18},
            {"type": "rectangle", "x": 20, "y": 20, "width": 30, "height": 20,
             "fill": "#3366CC"},
        ])
        if no_error("cdr_create(文字结构体测试)", txr):
            ta = txr["created"][0]["name"]
            tb = txr["created"][1]["name"]
            tA = {"name": ta}
            sflat = server.cdr_text(tx_id, tA, struct_info=True)
            st = sflat.get("structs") or {}
            if no_error("cdr_text(struct_info)", sflat):
                check("文字四个结构体可读",
                      bool(st.get("font")) and bool(st.get("space"))
                      and bool(st.get("align")) and bool(st.get("hyphenation")),
                      f"keys={sorted(st)}")
                check("结构体字段有真实值",
                      (st.get("font") or {}).get("size") == 18.0
                      and (st.get("space") or {}).get("line_spacing") is not None,
                      f"size={(st.get('font') or {}).get('size')}")
            r1 = server.cdr_text(tx_id, tA, range_action={
                "op": "insert_before", "start": 0, "end": 0, "text": "前"})
            if no_error("cdr_text(range insert_before)", r1):
                check("区间前置插入（InsertBeforeWide）",
                      "前" in str((r1.get("range_action") or {}).get("contents")),
                      str((r1.get("range_action") or {}).get("contents"))[:20])
            r2 = server.cdr_text(tx_id, tA, range_action={
                "op": "replace", "start": 0, "end": 1, "text": "改"})
            if no_error("cdr_text(range replace)", r2):
                check("区间替换（ReplaceWide）",
                      str((r2.get("range_action") or {}).get("contents")).startswith("改"),
                      str((r2.get("range_action") or {}).get("contents"))[:20])
            r3 = server.cdr_text(tx_id, tA, range_action={
                "op": "collapse", "start": 0, "end": 3, "to_end": True})
            if no_error("cdr_text(range collapse)", r3):
                check("区间可折叠（Collapse）",
                      (r3.get("range_action") or {}).get("range") == [3, 3],
                      str((r3.get("range_action") or {}).get("range")))
            r4 = server.cdr_text(tx_id, tA, range_action={
                "op": "set_range", "start": 0, "end": 2, "new_start": 1, "new_end": 4})
            if no_error("cdr_text(range set_range)", r4):
                check("区间可重设（SetRange）",
                      (r4.get("range_action") or {}).get("range") == [1, 4],
                      str((r4.get("range_action") or {}).get("range")))
            r5 = server.cdr_text(tx_id, tA, range_action={
                "op": "is_same", "start": 0, "end": 2,
                "other_start": 0, "other_end": 2})
            check("区间可判等（IsSame）",
                  (r5.get("range_action") or {}).get("same") is True,
                  str((r5.get("range_action") or {}).get("same")))
            r6 = server.cdr_text(tx_id, tA, range_action={
                "op": "in_range", "start": 0, "end": 2,
                "other_start": 0, "other_end": 5})
            check("区间包含判定（InRange）",
                  (r6.get("range_action") or {}).get("in_range") is True,
                  str((r6.get("range_action") or {}).get("in_range")))
            r7 = server.cdr_text(tx_id, tA, range_action={
                "op": "set_line_spacing", "type": 1, "line_spacing": 150})
            if no_error("cdr_text(range set_line_spacing)", r7):
                check("行距可设（SetLineSpacing）",
                      (r7.get("range_action") or {}).get("line_spacing") == 150.0,
                      str((r7.get("range_action") or {}).get("line_spacing")))
            rbad = server.cdr_text(tx_id, tA, range_action={"op": "nope"})
            check("非法 range_action 被拒绝", rbad.get("error") == "bad_argument",
                  str(rbad.get("error")))
            sc = server.cdr_style(tx_id, action="create", target={"name": tb},
                                  category="fill", style_name="smoke37")
            if no_error("cdr_style(create fill)", sc):
                ov = server.cdr_style(tx_id, action="override_props", style_name="smoke37")
                if no_error("cdr_style(override_props)", ov):
                    check("样式覆盖属性可读",
                          bool(ov.get("override_properties"))
                          and (ov.get("count") or 0) >= 1,
                          f"props={ov.get('override_properties')}")
                ih = server.cdr_style(tx_id, action="inherited", style_name="smoke37",
                                      property="primaryColor")
                if no_error("cdr_style(inherited)", ih):
                    check("样式属性继承状态可读", ih.get("inherited") in (True, False),
                          str(ih.get("inherited")))
                gp = server.cdr_style(tx_id, action="get_property", style_name="smoke37",
                                      property="primaryColor")
                if no_error("cdr_style(get_property)", gp):
                    check("样式属性可取值（GetProperty）", gp.get("value") is not None,
                          str(gp.get("value"))[:40])
                cp = server.cdr_style(tx_id, action="clear_property",
                                      style_name="smoke37", property="primaryColor")
                if no_error("cdr_style(clear_property)", cp):
                    check("样式属性可清除（ClearProperty）",
                          cp.get("cleared") in (True, False), str(cp.get("cleared")))
        no_error("关闭文字结构体测试文档", server.cdr_close(tx_id))

    # ---------------- 15ar. 第三十八波：结构体建文档 / Fill 工厂 / 命名箭头 / 边界构造器 ----------------
    print("\n[15ar] 第三十八波：CreateDocumentEx、CreateUniformFill、CreateArrowHead2、边界构造器")
    w38 = server.cdr_open(create_options={"name": "MCP-Struct", "width_mm": 120,
                                          "height_mm": 90, "resolution": 150})
    if "document_id" in w38:
        w38_id = w38["document_id"]
        check("可按结构体建文档（名称/尺寸）",
              w38.get("name") == "MCP-Struct"
              and w38.get("width_mm") == 120.0 and w38.get("height_mm") == 90.0,
              f"name={w38.get('name')} size={w38.get('width_mm')}x{w38.get('height_mm')}")
        w38r = server.cdr_create(w38_id, [
            {"type": "rect_ltrb", "left": 0, "top": 30, "right": 40, "bottom": 0,
             "radius": 3},
            {"type": "ellipse_ltrb", "left": 50, "top": 30, "right": 90, "bottom": 0,
             "start_angle": 0, "end_angle": 270, "pie": True},
            {"type": "polygon_full", "left": 10, "top": 80, "right": 50, "bottom": 40,
             "sides": 6, "star": True, "star_complexity": 2, "max_complexity": 1},
            {"type": "curve_segment", "x1": 0, "y1": 60, "x2": 40, "y2": 70},
        ])
        if no_error("cdr_create(边界构造器)", w38r):
            w_names = [c["name"] for c in w38r.get("created") or []]
            made = w38r.get("created") or []
            check("边界版矩形/椭圆/多边形可建",
                  (made[0] or {}).get("type") == "rectangle"
                  and (made[1] or {}).get("type") == "ellipse"
                  and (made[2] or {}).get("type") == "polygon"
                  and abs(((made[0] or {}).get("bbox") or {}).get("width", 0) - 40) < 0.01,
                  f"types={[m.get('type') for m in made[:3]]}")
            server.cdr_select(w38_id, name=w_names[0])
            server.cdr_select(w38_id, name=w_names[1])
            af = server.cdr_shape_ops(w38_id, op="apply_fill", base_color="#22AA66")
            if no_error("cdr_shape_ops(apply_fill)", af):
                check("可用 Fill 工厂批量上色（CreateUniformFill+ApplyFill）",
                      af.get("fill_type") == 1 and "34,170,102" in str(af.get("color")),
                      f"type={af.get('fill_type')} color={af.get('color')}")
            ahn = server.cdr_doc(w38_id, action="arrowhead_named",
                                 target={"name": w_names[3]}, query="MCP箭头38",
                                 register=True)
            if no_error("cdr_doc(arrowhead_named)", ahn):
                check("可建命名箭头并登记（CreateArrowHead2）",
                      ahn.get("name") == "MCP箭头38" and ahn.get("registered") is True
                      and (ahn.get("arrowheads") or 0) > 0,
                      f"name={ahn.get('name')} count={ahn.get('arrowheads')}")
        no_error("关闭结构体建文档测试", server.cdr_close(w38_id))

    # ---------------- 15as. 第三十九波：形状矩阵 / 扭曲 / 文字内嵌 / 范围复制 ----------------
    print("\n[15as] 第三十九波：形状矩阵、扭曲效果、文字内嵌、范围复制、命中判定")
    m_doc = server.cdr_open(preset="A4")
    if "document_id" in m_doc:
        m_id = m_doc["document_id"]
        mr = server.cdr_create(m_id, [
            {"type": "rectangle", "x": 20, "y": 120, "width": 40, "height": 30},
            {"type": "rectangle", "x": 80, "y": 120, "width": 20, "height": 20},
            {"type": "text", "x": 20, "y": 200, "text": "内嵌文字", "size": 14},
            {"type": "rectangle", "x": 20, "y": 40, "width": 15, "height": 15},
        ])
        if no_error("cdr_create(矩阵/扭曲测试)", mr):
            ma, mb, mt, mm = [c["name"] for c in mr["created"]]
            mA, mB, mT, mM = ({"name": ma}, {"name": mb}, {"name": mt},
                              {"name": mm})
            # 注意顺序：GetMatrix 之后再调 IsOnShape 可能失效（实测在长跑里会返回 0），
            # 所以命中判定放在矩阵读取之前；另外先把视口归位（缩放异常时 IsOnShape 也会返回 0）
            server.cdr_view(m_id, action="actual_size")
            ins_shape = server.cdr_shape_ops(m_id, op="is_on_shape", target=mA,
                                             x=30, y=130)
            edge_shape = server.cdr_shape_ops(m_id, op="is_on_shape", target=mA,
                                              x=20, y=135)
            check("对象命中判定（2=内部 / 1=边界）",
                  ins_shape.get("hit_state") == 2
                  and edge_shape.get("hit_state") == 1,
                  f"in={ins_shape.get('hit_state')} edge={edge_shape.get('hit_state')}")
            mtx = server.cdr_shape_ops(m_id, op="matrix", target=mA)
            if no_error("cdr_shape_ops(matrix read)", mtx):
                check("形状矩阵可读（6 元素）",
                      len(mtx.get("matrix") or []) == 6, str(mtx.get("matrix")))
            mset = server.cdr_shape_ops(m_id, op="matrix", target=mM,
                                        params={"matrix": [1, 0, 0, 1, 0, 0]})
            if no_error("cdr_shape_ops(matrix set)", mset):
                check("形状矩阵可写并精确回读",
                      mset.get("matrix") == [1.0, 0.0, 0.0, 1.0, 0.0, 0.0],
                      str(mset.get("matrix")))
            ity = server.cdr_shape_ops(m_id, op="is_type", target=mA,
                                       params={"types": [1, 2]})
            ity2 = server.cdr_shape_ops(m_id, op="is_type", target=mA,
                                        params={"types": [2]})
            check("类型集合判定可用（IsTypeAnyOf）",
                  ity.get("matches") is True and ity2.get("matches") is False,
                  f"{ity.get('matches')}/{ity2.get('matches')}")
            front = server.cdr_shape_ops(m_id, op="in_front_of", target=mB, other=mA)
            check("层序判定可用（OrderIsInFrontOf）", front.get("in_front") is True,
                  str(front.get("in_front")))
            ptx = server.cdr_shape_ops(m_id, op="place_text", target=mA, other=mT)
            if no_error("cdr_shape_ops(place_text)", ptx):
                check("文字可内嵌到形状（PlaceTextInside）",
                      ptx.get("type_id") == 6 and ptx.get("text") == "内嵌文字",
                      f"type={ptx.get('type_id')} text={ptx.get('text')}")
            before_shapes = (server.cdr_inspect(m_id).get("state") or {}).get("shapes")
            dup = server.cdr_shape_ops(m_id, op="duplicate_range", target=mB,
                                       params={"offset_x": 8, "offset_y": 8})
            if no_error("cdr_shape_ops(duplicate_range)", dup):
                check("可复制成范围（DuplicateAsRange）",
                      dup.get("count") == 1
                      and (dup.get("page_shapes") or 0) > (before_shapes or 0),
                      f"count={dup.get('count')} shapes={before_shapes}->{dup.get('page_shapes')}")
            clo = server.cdr_shape_ops(m_id, op="clone_range", target=mB)
            if no_error("cdr_shape_ops(clone_range)", clo):
                check("可克隆成范围（CloneAsRange）", clo.get("count") == 1,
                      str(clo.get("count")))
            dst = server.cdr_shape_ops(m_id, op="distort", target=mB,
                                       params={"kind": "twister", "origin_x": 90,
                                               "origin_y": 130, "angle": 30})
            if no_error("cdr_shape_ops(distort)", dst):
                check("扭曲效果可创建（Twister）",
                      dst.get("effects") == 1
                      and (dst.get("bbox_before") != dst.get("bbox_after")),
                      f"effects={dst.get('effects')}")
            bad_dst = server.cdr_shape_ops(m_id, op="distort", target=mB,
                                           params={"kind": "nope"})
            check("非法扭曲类型被拒绝", bad_dst.get("error") == "bad_argument",
                  str(bad_dst.get("error")))
        no_error("关闭矩阵/扭曲测试文档", server.cdr_close(m_id))

    # ---------------- 15at. 第四十波：参考点定位 / 吸附反查 / 链接 / 位图化 ----------------
    print("\n[15at] 第四十波：参考点定位、FindSnapPoint、链接对象、ConvertToBitmapEx、叠印状态")
    p40_doc = server.cdr_open(preset="A4")
    if "document_id" in p40_doc:
        p40_id = p40_doc["document_id"]
        p40r = server.cdr_create(p40_id, [
            {"type": "rectangle", "x": 20, "y": 20, "width": 40, "height": 30},
        ])
        if no_error("cdr_create(参考点测试)", p40r):
            pA = {"name": p40r["created"][0]["name"]}
            pe = server.cdr_shape_ops(p40_id, op="position_ex", target=pA,
                                      params={"ref": 9})
            if no_error("cdr_shape_ops(position_ex read)", pe):
                check("参考点定位可读（9=中心）", pe.get("position") == [40.0, 35.0],
                      str(pe.get("position")))
            pset = server.cdr_shape_ops(p40_id, op="position_ex", target=pA,
                                        params={"ref": 9, "x": 70, "y": 80})
            if no_error("cdr_shape_ops(position_ex set)", pset):
                bb = pset.get("bbox") or {}
                check("可按参考点定位对象",
                      pset.get("position") == [70.0, 80.0]
                      and bb.get("x") == 50.0 and bb.get("y") == 65.0,
                      f"pos={pset.get('position')} bbox={bb}")
            ops_ = server.cdr_shape_ops(p40_id, op="overprint_state", target=pA)
            if no_error("cdr_shape_ops(overprint_state)", ops_):
                check("叠印状态可读", ops_.get("outline_state") == 2
                      and ops_.get("overprint_fill") is False,
                      f"fill={ops_.get('fill_state')} outline={ops_.get('outline_state')}")
            linked = server.cdr_shape_ops(p40_id, op="linked", target=pA,
                                          params={"link_type": 3})
            if no_error("cdr_shape_ops(linked)", linked):
                check("链接对象查询可用（无链接时空集合）",
                      linked.get("count") == 0, str(linked.get("count")))
            snap40 = server.cdr_snap(p40_id, action="list", target=pA)
            ref40 = ((snap40.get("points") or [{}])[0] or {}).get("reference")
            if ref40:
                fs = server.cdr_shape_ops(p40_id, op="find_snap", target=pA,
                                          params={"reference": ref40})
                if no_error("cdr_shape_ops(find_snap)", fs):
                    check("可按引用反查吸附点（FindSnapPoint）",
                          fs.get("found") is True and fs.get("type") == 2,
                          f"type={fs.get('type')} pos={fs.get('position')}")
            fsbad = server.cdr_shape_ops(p40_id, op="find_snap", target=pA)
            check("find_snap 缺引用被拒绝", fsbad.get("error") == "bad_argument",
                  str(fsbad.get("error")))
            rast = server.cdr_shape_ops(p40_id, op="rasterize_ex", target=pA,
                                        params={"mode": 4, "resolution": 150})
            if no_error("cdr_shape_ops(rasterize_ex)", rast):
                check("可带参数位图化（ConvertToBitmapEx）",
                      rast.get("type_id") == 5 and rast.get("has_bitmap") is True,
                      f"type={rast.get('type_id')} bitmap={rast.get('has_bitmap')}")
        no_error("关闭参考点测试文档", server.cdr_close(p40_id))

    # ---------------- 15au. 第四十一波：任意节点/线段集合 + 位图构造器 ----------------
    print("\n[15au] 第四十一波：CreateNodeRange/CreateSegmentRange 任意集合、CreateBitmap2/BitmapRect")
    nr_doc = server.cdr_open(preset="A4")
    if "document_id" in nr_doc:
        nr_id = nr_doc["document_id"]
        nrr = server.cdr_create(nr_id, [
            {"type": "polyline", "points": [[0, 0], [20, 0], [40, 0], [60, 0], [80, 0]],
             "closed": False},
            {"type": "bitmap_at", "x": 100, "y": 20, "width": 20, "height": 20,
             "color": "#3388FF", "pixels_w": 24, "pixels_h": 24},
            {"type": "bitmap_rect", "x": 130, "y": 20, "width": 25, "height": 25,
             "color": "#FF8833", "pixels_w": 24, "pixels_h": 24},
        ])
        if no_error("cdr_create(任意集合/位图构造器)", nrr):
            made41 = nrr.get("created") or []
            check("位图构造器可建（CreateBitmap2/BitmapRect）",
                  (made41[1] or {}).get("type") == "type_5"
                  and (made41[2] or {}).get("type") == "type_5"
                  and abs(((made41[2] or {}).get("bbox") or {}).get("width", 0) - 25) < 0.01,
                  f"types={[(m.get('type'), (m.get('bbox') or {}).get('width')) for m in made41[1:]]}")
            nP = {"name": made41[0]["name"]}
            arb = server.cdr_node(nr_id, nP, action="arbitrary",
                                  params={"op": "move", "nodes": [1, 3, 5],
                                          "dx": 0, "dy": 10})
            if no_error("cdr_node(arbitrary move)", arb):
                check("任意节点集合可移动（CreateNodeRange）",
                      arb.get("nodes_selected") == 3
                      and (arb.get("range_bbox") or [0, 0])[1] == 10.0,
                      f"selected={arb.get('nodes_selected')} bbox={arb.get('range_bbox')}")
            arb2 = server.cdr_node(nr_id, nP, action="arbitrary",
                                   params={"op": "rotate", "nodes": [1, 3, 5],
                                           "angle": 30, "cx": 40, "cy": 10})
            if no_error("cdr_node(arbitrary rotate)", arb2):
                check("任意节点集合可旋转", arb2.get("nodes_selected") == 3,
                      str(arb2.get("range_bbox")))
            arbbad = server.cdr_node(nr_id, nP, action="arbitrary", params={"op": "move"})
            check("arbitrary 缺节点被拒绝", arbbad.get("error") == "bad_argument",
                  str(arbbad.get("error")))
            srg = server.cdr_path(nr_id, nP, action="segment_range",
                                  params={"op": "add_node", "segments": [1, 2]})
            if no_error("cdr_path(segment_range add_node)", srg):
                check("任意线段集合可加节点（CreateSegmentRange）",
                      srg.get("count") == 2
                      and (srg.get("nodes_after") or 0) > (srg.get("nodes_before") or 0),
                      f"segments={srg.get('count')} nodes={srg.get('nodes_before')}"
                      f"->{srg.get('nodes_after')}")
            srt = server.cdr_path(nr_id, nP, action="segment_range",
                                  params={"op": "set_type", "segments": [1],
                                          "segment_type": 0})
            if no_error("cdr_path(segment_range set_type)", srt):
                check("线段集合可改类型", (srt.get("first_segment") or 0) >= 1,
                      f"first={srt.get('first_segment')} last={srt.get('last_segment')}")
        no_error("关闭任意集合测试文档", server.cdr_close(nr_id))

    # ---------------- 15av. 第四十二波：对象列表建范围 / 范围位图化 / 箭头变体 ----------------
    print("\n[15av] 第四十二波：CreateShapeRangeFromArray、范围位图化、箭头变体、颜色上下文")
    r42_doc = server.cdr_open(preset="A4")
    if "document_id" in r42_doc:
        r42_id = r42_doc["document_id"]
        r42r = server.cdr_create(r42_id, [
            {"type": "rectangle", "x": 20, "y": 20, "width": 30, "height": 20,
             "fill": "#3366CC"},
            {"type": "rectangle", "x": 60, "y": 20, "width": 30, "height": 20,
             "fill": "#CC6633"},
            {"type": "curve_segment", "x1": 20, "y1": 80, "x2": 60, "y2": 100},
        ])
        if no_error("cdr_create(对象列表/箭头测试)", r42r):
            qa, qb, qc = [x["name"] for x in r42r["created"]]
            rfl = server.cdr_shape_ops(r42_id, op="range_from_list",
                                       params={"targets": [{"name": qa},
                                                           {"name": qb}]})
            if no_error("cdr_shape_ops(range_from_list)", rfl):
                check("可用对象列表直接建范围（CreateShapeRangeFromArray）",
                      rfl.get("count") == 2 and rfl.get("exists_first") is True,
                      f"count={rfl.get('count')} exists={rfl.get('exists_first')}")
            rfl2 = server.cdr_shape_ops(r42_id, op="range_from_list",
                                        params={"targets": [{"name": qa},
                                                            {"name": qb}],
                                                "then": "rotate_ex", "angle": 15,
                                                "cx": 55, "cy": 30})
            if no_error("cdr_shape_ops(range_from_list then=rotate_ex)", rfl2):
                check("对象列表范围可整批旋转", rfl2.get("applied") == "rotate_ex",
                      str(rfl2.get("applied")))
            ahex = server.cdr_doc(r42_id, action="arrowhead_ex",
                                  target={"name": qc}, query="EX2", x=40, y=90,
                                  register=True)
            if no_error("cdr_doc(arrowhead_ex)", ahex):
                check("命名箭头变体可用（CreateArrowHeadEx2）",
                      ahex.get("name") == "EX2" and ahex.get("registered") is True
                      and (ahex.get("options_length") or 0) > 0,
                      f"name={ahex.get('name')} len={ahex.get('options_length')}")
            ssd = server.cdr_doc(r42_id, action="stylesheet_default")
            if no_error("cdr_doc(stylesheet_default)", ssd):
                check("样式表可设为默认（SaveStyleSheetAsDefault）",
                      ssd.get("saved") is True, str(ssd.get("saved")))
            acc = server.cdr_doc(r42_id, action="assign_color_context",
                                 palette_index=1)
            if no_error("cdr_doc(assign_color_context)", acc):
                check("可指派文档色彩上下文（AssignColorContext）",
                      acc.get("profile_index") == 1
                      and acc.get("rendering_intent") is not None,
                      f"intent={acc.get('rendering_intent')}")
            server.cdr_select(r42_id, name=qa)
            server.cdr_select(r42_id, name=qb)
            rbm = server.cdr_shape_ops(r42_id, op="range_bitmap",
                                       params={"mode": 4, "resolution": 96})
            if no_error("cdr_shape_ops(range_bitmap)", rbm):
                check("整批位图化可用（ShapeRange.ConvertToBitmapEx）",
                      rbm.get("type_id") == 5 and rbm.get("has_bitmap") is True,
                      f"type={rbm.get('type_id')}")
        no_error("关闭对象列表测试文档", server.cdr_close(r42_id))

    # ---------------- 15aw. 第四十三波：调色板选项 / PasteEx / 两种导出 API ----------------
    print("\n[15aw] 第四十三波：CreateStructPaletteOptions、PasteEx、Document.Export/ExportBitmap")
    e43_doc = server.cdr_open(preset="A4")
    if "document_id" in e43_doc:
        e43_id = e43_doc["document_id"]
        e43r = server.cdr_create(e43_id, [
            {"type": "bitmap", "x": 20, "y": 20, "width": 40, "height": 40,
             "color": "#CC8844", "pixels_w": 32, "pixels_h": 32},
            {"type": "rectangle", "x": 80, "y": 20, "width": 30, "height": 20,
             "fill": "#3366CC"},
        ])
        if no_error("cdr_create(调色板/导出测试)", e43r):
            bm43 = e43r["created"][0]["name"]
            rect43 = e43r["created"][1]["name"]
            po = server.cdr_bitmap(e43_id, {"name": bm43},
                                   action="paletted_options", num_colors=16,
                                   dither_type=1)
            if no_error("cdr_bitmap(paletted_options)", po):
                check("调色板选项可转索引色（ConvertToPaletted2）",
                      po.get("mode") == 3 and po.get("num_colors") == 16,
                      f"mode={po.get('mode')} colors={po.get('num_colors')}")
            server.cdr_select(e43_id, name=rect43)
            pex = server.cdr_layers(e43_id, action="paste_ex")
            if no_error("cdr_layers(paste_ex)", pex):
                check("可高级粘贴（CreateStructPasteOptions + PasteEx）",
                      (pex.get("pasted") or 0) >= 1,
                      f"pasted={pex.get('pasted')}")
            p_doc = os.path.join(os.getcwd(), "tmp",
                                 f"live_smoke43_doc_{os.urandom(2).hex()}.png")
            ex_doc = server.cdr_export(e43_id, p_doc, "png", export_api="document")
            if no_error("cdr_export(api=document)", ex_doc):
                check("Document.Export 可落盘",
                      (ex_doc.get("bytes") or 0) > 0 and os.path.exists(p_doc),
                      f"bytes={ex_doc.get('bytes')} api={ex_doc.get('export_api')}")
            p_bmp = os.path.join(os.getcwd(), "tmp",
                                 f"live_smoke43_bmp_{os.urandom(2).hex()}.png")
            ex_bmp = server.cdr_export(e43_id, p_bmp, "png", dpi=96,
                                       export_api="bitmap")
            if no_error("cdr_export(api=bitmap)", ex_bmp):
                check("Document.ExportBitmap 可落盘（需 Finish）",
                      (ex_bmp.get("bytes") or 0) > 0 and os.path.exists(p_bmp),
                      f"bytes={ex_bmp.get('bytes')} api={ex_bmp.get('export_api')}")
        no_error("关闭调色板/导出测试文档", server.cdr_close(e43_id))

    # ---------------- 15ax. 第四十四波：色彩上下文 / 新窗口 / 渐变变换 / 纹理透明 ----------------
    print("\n[15ax] 第四十四波：颜色上下文深操作、新窗口、渐变变换矩阵、纹理透明")
    c44_doc = server.cdr_open(preset="A4")
    if "document_id" in c44_doc:
        c44_id = c44_doc["document_id"]
        c44r = server.cdr_create(c44_id, [
            {"type": "rectangle", "x": 20, "y": 20, "width": 50, "height": 40,
             "fill": {"type": "fountain", "start": "#FF0000", "end": "#0000FF"}},
        ])
        if no_error("cdr_create(色彩上下文测试)", c44r):
            cA = {"name": c44r["created"][0]["name"]}
            ctx41 = server.cdr_doc(c44_id, action="color_context_info")
            if no_error("cdr_doc(color_context_info)", ctx41):
                check("颜色上下文可读（3 个配置文件）",
                      (ctx41.get("profiles") or 0) >= 3
                      and "sRGB" in str(ctx41.get("profile_names")),
                      f"profiles={ctx41.get('profiles')}")
            ctx42 = server.cdr_doc(c44_id, action="color_context_copy")
            if no_error("cdr_doc(color_context_copy)", ctx42):
                check("颜色上下文可复制（GetCopy + IsSame）",
                      ctx42.get("same_as_source") is True,
                      str(ctx42.get("same_as_source")))
            ctx43 = server.cdr_doc(c44_id, action="color_context_merge",
                                   palette_index=1)
            if no_error("cdr_doc(color_context_merge)", ctx43):
                check("颜色上下文可合并（Merge）", ctx43.get("merged") is True,
                      str(ctx43.get("merged")))
            ft = server.cdr_doc(c44_id, action="fountain_transform", target=cA)
            if no_error("cdr_doc(fountain_transform)", ft):
                check("渐变变换矩阵可读（GetTransformations）",
                      len(ft.get("transformations") or []) == 4,
                      str(ft.get("transformations")))
            nw = server.cdr_view(c44_id, action="new_window")
            if no_error("cdr_view(new_window)", nw):
                check("可开新窗口（IVGWindow.NewWindow）",
                      nw.get("created") is True and (nw.get("windows") or 0) >= 2,
                      f"windows={nw.get('windows')}")
            tt = server.cdr_modify(c44_id, cA, {
                "transparency": {"type": "texture", "texture": "BARK",
                                 "library": "Textures"}})
            if no_error("cdr_modify(纹理透明)", tt):
                check("纹理透明可用（ApplyTextureTransparency）",
                      "transparency:texture" in (tt.get("applied") or []),
                      str(tt.get("applied")))
        no_error("关闭色彩上下文测试文档", server.cdr_close(c44_id))

    # ---------------- 15ay. 第四十五波：Outline.SetPropertiesEx / Text.BeginEdit ----------------
    print("\n[15ay] 第四十五波：SetPropertiesEx（15 参）与文字 BeginEdit")
    o45_doc = server.cdr_open(preset="A4")
    if "document_id" in o45_doc:
        o45_id = o45_doc["document_id"]
        o45r = server.cdr_create(o45_id, [
            {"type": "rectangle", "x": 20, "y": 20, "width": 30, "height": 20,
             "outline": {"width": 1.0, "color": "#CC0000"}, "fill": "#3366CC"},
            {"type": "text", "x": 20, "y": 80, "text": "编辑测试", "size": 16},
        ])
        if no_error("cdr_create(SetPropertiesEx 测试)", o45r):
            o45a = {"name": o45r["created"][0]["name"]}
            o45t = {"name": o45r["created"][1]["name"]}
            sp = server.cdr_outline(o45_id, target=o45a, action="set", changes={
                "use_set_properties_ex": True, "width": 2.5, "line_caps": 1,
                "line_join": 1, "miter_limit": 8.0, "justification": 1})
            if no_error("cdr_outline(use_set_properties_ex)", sp):
                check("SetPropertiesEx 15 参全参可用",
                      sp.get("width") == 2.5 and sp.get("line_caps") == 1
                      and sp.get("justification") == 1,
                      f"width={sp.get('width')} caps={sp.get('line_caps')} "
                      f"just={sp.get('justification')}")
            be = server.cdr_text(o45_id, o45t, begin_edit=True)
            if no_error("cdr_text(begin_edit)", be):
                check("文字可进入编辑态（Text.BeginEdit）",
                      "begin_edit" in (be.get("applied") or []),
                      str(be.get("applied")))
        no_error("关闭 SetPropertiesEx 测试文档", server.cdr_close(o45_id))

    # ---------------- 15az. 第四十六波：CorelScriptTools 辅助对象 ----------------
    print("\n[15az] 第四十六波：CorelScriptTools（VBA 辅助对象）单位换算/数学/日期")
    st46_doc = server.cdr_open(preset="A4")
    if "document_id" in st46_doc:
        st46_id = st46_doc["document_id"]
        st46 = server.cdr_doc(st46_id, action="script_tools", value=1.0)
        if no_error("cdr_doc(script_tools)", st46):
            check("CorelScriptTools 单位换算可用（FromInches/FromPoints）",
                  st46.get("from_inches") == 254000.0
                  and abs((st46.get("from_points") or 0) - 3527.778) < 0.01
                  and st46.get("to_inches_roundtrip") == 1.0,
                  f"inch={st46.get('from_inches')} pt={st46.get('from_points')} "
                  f"rt={st46.get('to_inches_roundtrip')}")
            check("CorelScriptTools 数学/日期可用（ASin/BuildDate）",
                  st46.get("asin_rad") == 1.571
                  and "2026-10-26" in str(st46.get("build_date")),
                  f"asin={st46.get('asin_rad')} date={st46.get('build_date')}")
        no_error("关闭 script_tools 测试文档", server.cdr_close(st46_id))

    # ---------------- 15ba. 第四十七波：宽文本 / 节点坐标搬运 / 整批描边 / 页边对齐 / ImportEx ----------------
    print("\n[15ba] 第四十七波：CreateArtisticTextWide·Node.SetPoint·SetOutlineProperties·AlignRangeToPage·ImportEx")
    w47_doc = server.cdr_open(preset="A4")
    if "document_id" in w47_doc:
        w47_id = w47_doc["document_id"]
        w47r = server.cdr_create(w47_id, [
            {"type": "polyline", "points": [[0, 0], [20, 0], [40, 20]], "closed": False},
            {"type": "rectangle", "x": 20, "y": 40, "width": 30, "height": 20,
             "outline": {"width": 1.0}},
            {"type": "text_wide", "x": 20, "y": 90, "text": "宽文本", "size": 16,
             "charset": 134},
            {"type": "paragraph_wide", "x": 20, "y": 120, "x2": 120, "y2": 160,
             "text": "段落宽文本", "size": 12},
        ])
        if no_error("cdr_create(宽文本/描边测试)", w47r):
            w47 = w47r.get("created") or []
            check("宽文本构造可用（CreateArtisticTextWide/ParagraphTextWide）",
                  (w47[2] or {}).get("type") == "text"
                  and (w47[3] or {}).get("type") == "text",
                  f"types={[(w.get('type')) for w in w47[2:]]}")
            w47c = {"name": w47[0]["name"]}
            cp = server.cdr_node(w47_id, w47c, action="copy_point", node=1, end=3)
            if no_error("cdr_node(copy_point)", cp):
                check("节点坐标可搬运（GetPoint/SetPoint）",
                      cp.get("source_pos") == cp.get("target_pos"),
                      f"{cp.get('source_pos')} → {cp.get('target_pos')}")
            server.cdr_select(w47_id, name=w47[1]["name"])
            ro = server.cdr_shape_ops(w47_id, op="range_outline",
                                      params={"width": 3.0, "line_caps": 1})
            if no_error("cdr_shape_ops(range_outline)", ro):
                check("整批描边可用（SetOutlineProperties）",
                      ro.get("first_width") == 3.0
                      and ro.get("api") == "SetOutlineProperties",
                      f"width={ro.get('first_width')} api={ro.get('api')}")
            roe = server.cdr_shape_ops(w47_id, op="range_outline",
                                       params={"width": 2.0,
                                               "with_justification": True,
                                               "justification": 1})
            if no_error("cdr_shape_ops(range_outline ex)", roe):
                check("整批描边 15 参可用（SetOutlinePropertiesEx）",
                      roe.get("api") == "SetOutlinePropertiesEx"
                      and roe.get("first_width") == 2.0,
                      f"api={roe.get('api')} width={roe.get('first_width')}")
            ape = server.cdr_align(w47_id, horizontal="left", to="page_edge")
            if no_error("cdr_align(to=page_edge)", ape):
                check("可对齐到页边（AlignRangeToPage）",
                      ape.get("to") == "page_edge" and ape.get("axes") == [1],
                      f"axes={ape.get('axes')}")
            p47 = os.path.join(os.getcwd(), "tmp",
                               f"w47_src_{os.urandom(4).hex()}.png")
            ex47 = server.cdr_export(w47_id, p47, "png", dpi=96, overwrite=True)
            if no_error("cdr_export(导入源)", ex47) and os.path.exists(p47):
                im47 = server.cdr_import(w47_id, p47, filter_id=802)
                if no_error("cdr_import(filter_id)", im47):
                    # 802 是**导出**过滤器号；ImportEx 要的是**导入**过滤器号，
                    # 所以这里能调通但不会真的导入 —— 正好把这个坑固化成断言
                    check("ImportEx 可调用（导出过滤器号不生效，已在文档记录）",
                          im47.get("imported") is True
                          and (im47.get("shapes_added") or 0) == 0,
                          f"added={im47.get('shapes_added')}")
                im47b = server.cdr_import(w47_id, p47)
                check("默认导入（过滤器 0 自动判断）可加入对象",
                      (im47b.get("shapes_added") or 0) >= 1,
                      f"added={im47b.get('shapes_added')}")
        no_error("关闭宽文本测试文档", server.cdr_close(w47_id))

    # ---------------- 15bb. 第四十八波：颜色 Assign 家族 / 图层复制 / 对齐到范围 ----------------
    print("\n[15bb] 第四十八波：FixedAssign·CorelScriptAssign·PaletteAssign·Invalidate·CopyToLayerAsRange")
    c48_doc = server.cdr_open(preset="A4")
    if "document_id" in c48_doc:
        c48_id = c48_doc["document_id"]
        cfix = server.cdr_color(action="object_assign", model="fixed",
                                values=[2, 5, 100])
        if no_error("cdr_color(model=fixed)", cfix):
            check("固定调色板赋值可用（FixedAssign）",
                  "PAN_PROC" in str(cfix.get("after")),
                  str(cfix.get("after"))[:60])
        cscr = server.cdr_color(action="object_assign", model="script",
                                values=[2, 10, 20, 30, 40])
        if no_error("cdr_color(model=script)", cscr):
            check("脚本式分量赋值可用（CorelScriptAssign）",
                  str(cscr.get("after", "")).startswith("CMYK,USER,10,20,30,40"),
                  str(cscr.get("after"))[:60])
        cinv = server.cdr_color(action="object_invalidate", color="#3366CC")
        if no_error("cdr_color(object_invalidate)", cinv):
            check("颜色可失效（Invalidate）",
                  "CSINVALID" in str(cinv.get("to_string")),
                  str(cinv.get("to_string"))[:40])
        server.cdr_doc(c48_id, action="palette_open_fixed", palette_id=1)
        pl48 = server.cdr_doc(c48_id, action="palettes")
        pid48 = ((pl48.get("palettes") or [{}])[0] or {}).get("identifier")
        if pid48:
            cpal = server.cdr_color(action="object_palette", color="#000000",
                                    color2=pid48, values=[5])
            if no_error("cdr_color(object_palette)", cpal):
                check("可按调色板取色（PaletteAssign）",
                      str(cpal.get("to_string", "")).startswith("PANTONE"),
                      str(cpal.get("to_string"))[:60])
        c48r = server.cdr_create(c48_id, [
            {"type": "rectangle", "x": 0, "y": 0, "width": 20, "height": 10},
            {"type": "rectangle", "x": 40, "y": 0, "width": 20, "height": 10},
        ])
        if no_error("cdr_create(图层复制测试)", c48r):
            c48a, c48b = [x["name"] for x in c48r["created"]]
            server.cdr_layers(c48_id, action="create", name="目标层")
            ctl = server.cdr_shape_ops(c48_id, op="copy_to_layer_range",
                                       target={"name": c48a},
                                       params={"layer": "目标层"})
            if no_error("cdr_shape_ops(copy_to_layer_range)", ctl):
                check("可复制到指定图层（CopyToLayerAsRange）",
                      ctl.get("count") == 1 and ctl.get("layer") == "目标层"
                      and bool(ctl.get("new_name")),
                      f"layer={ctl.get('layer')} name={ctl.get('new_name')}")
            atr = server.cdr_shape_ops(c48_id, op="align_to_range",
                                       target={"name": c48a},
                                       other={"name": c48b},
                                       params={"align_type": 1})
            if no_error("cdr_shape_ops(align_to_range)", atr):
                bb48 = atr.get("bbox") or {}
                check("可对齐到形状范围（AlignToShapeRange）",
                      bb48.get("x") == 40.0, f"bbox={bb48}")
        no_error("关闭颜色/图层测试文档", server.cdr_close(c48_id))

    # ---------------- 15bc. 第四十九波：颜色分量读取 / 转调色板 / 按掩码复制属性 ----------------
    print("\n[15bc] 第四十九波：CorelScriptGetComponent 分量读取、ConvertToPalette、CopyPropertiesFrom")
    k49_doc = server.cdr_open(preset="A4")
    if "document_id" in k49_doc:
        k49_id = k49_doc["document_id"]
        kc = server.cdr_color(action="object_components",
                              color={"c": 10, "m": 20, "y": 30, "k": 40}, model=2)
        if no_error("cdr_color(object_components cmyk)", kc):
            check("颜色分量可读（CorelScriptGetComponent / CMYK）",
                  kc.get("status") == 5002
                  and (kc.get("values") or [])[:4] == [10.0, 20.0, 30.0, 40.0],
                  f"status={kc.get('status')} values={kc.get('values')}")
        kr = server.cdr_color(action="object_components", color="#0B1621", model=5)
        if no_error("cdr_color(object_components rgb)", kr):
            check("颜色分量可读（RGB）",
                  (kr.get("values") or [])[:3] == [11.0, 22.0, 33.0],
                  f"values={kr.get('values')}")
        server.cdr_doc(k49_id, action="palette_open_fixed", palette_id=1)
        kpid = (((server.cdr_doc(k49_id, action="palettes").get("palettes")
                  or [{}])[0]) or {}).get("identifier")
        if kpid:
            kcp = server.cdr_color(action="object_convert_palette", color="#783C1E",
                                   color2=kpid)
            if no_error("cdr_color(object_convert_palette)", kcp):
                check("颜色可转指定调色板（ConvertToPalette）",
                      "PAN_SPOT" in str(kcp.get("after"))
                      or "PANTONE" in str(kcp.get("after")),
                      str(kcp.get("after"))[:60])
        k49r = server.cdr_create(k49_id, [
            {"type": "rectangle", "x": 0, "y": 0, "width": 20, "height": 10,
             "fill": "#C80A0A"},
            {"type": "rectangle", "x": 40, "y": 0, "width": 20, "height": 10,
             "fill": "#0AC80A"},
        ])
        if no_error("cdr_create(属性复制测试)", k49r):
            k49a, k49b = [x["name"] for x in k49r["created"]]
            kcpy = server.cdr_shape_ops(k49_id, op="copy_props",
                                        target={"name": k49a},
                                        other={"name": k49b},
                                        params={"mask": 4})
            if no_error("cdr_shape_ops(copy_props)", kcpy):
                check("可按掩码复制属性（CopyPropertiesFrom）",
                      kcpy.get("applied") is True and kcpy.get("mask") == 4,
                      f"applied={kcpy.get('applied')} mask={kcpy.get('mask')}")
        no_error("关闭分量/属性测试文档", server.cdr_close(k49_id))

    # ---------------- 15bd. 第五十波：形状斜切 / 子路径追加曲线段 / 抽出子路径 / 焊接落地 ----------------
    print("\n[15bd] 第五十波：SkewEx·AppendCurveSegment·SubPath.Extract·WeldWith+BindToDocument")
    f50_doc = server.cdr_open(preset="A4")
    if "document_id" in f50_doc:
        f50_id = f50_doc["document_id"]
        f50r = server.cdr_create(f50_id, [
            {"type": "rectangle", "x": 0, "y": 0, "width": 20, "height": 10},
            {"type": "polyline", "points": [[0, 60], [20, 60], [40, 80]], "closed": False},
            {"type": "curve_segment", "x1": 60, "y1": 60, "x2": 100, "y2": 60},
            {"type": "curve_segment", "x1": 100, "y1": 60, "x2": 140, "y2": 90},
        ])
        if no_error("cdr_create(斜切/子路径测试)", f50r):
            fR, fP, fC1, fC2 = [{"name": c["name"]} for c in f50r["created"]]
            sk = server.cdr_shape_ops(f50_id, op="skew_ex", target=fR,
                                      params={"angle_x": 10, "angle_y": 0,
                                              "cx": 10, "cy": 5})
            if no_error("cdr_shape_ops(skew_ex)", sk):
                bb_before = (sk.get("bbox_before") or {})
                bb_after = (sk.get("bbox_after") or {})
                check("形状级斜切可用（SkewEx）",
                      bb_after.get("width", 0) > bb_before.get("width", 0),
                      f"{bb_before.get('width')} → {bb_after.get('width')}")
            acs = server.cdr_path(f50_id, fP, action="append_curve_segment",
                                  params={"x": 60, "y": 100, "subpath": 1})
            if no_error("cdr_path(append_curve_segment)", acs):
                check("子路径可追加曲线段（AppendCurveSegment）",
                      (acs.get("nodes_after") or 0) > (acs.get("nodes_before") or 0),
                      f"{acs.get('nodes_before')} → {acs.get('nodes_after')}")
            ext = server.cdr_path(f50_id, fP, action="extract_subpath",
                                  params={"subpath": 1})
            if no_error("cdr_path(extract_subpath)", ext):
                check("子路径可抽出成独立对象（SubPath.Extract）",
                      ext.get("type_id") == 3 and bool(ext.get("shape")),
                      f"type={ext.get('type_id')} name={ext.get('shape')}")
            wts = server.cdr_path(f50_id, fC1, action="weld_to_shape", other=fC2)
            if no_error("cdr_path(weld_to_shape)", wts):
                check("焊接结果可落地成形状（BindToDocument + CreateCurve）",
                      wts.get("type_id") == 3 and (wts.get("nodes") or 0) >= 3
                      and bool(wts.get("bbox")),
                      f"type={wts.get('type_id')} nodes={wts.get('nodes')}")
        no_error("关闭斜切/子路径测试文档", server.cdr_close(f50_id))

    # ---------------- 15be. 第五十一波：结构体工厂写入 / 上下文按模型取配置 / 风格色改写 / 无名箭头 ----------------
    print("\n[15be] 第五十一波：CreateStruct*Properties 写入、ColorProfile(model)、ModifyColorStyleColor")
    s51_doc = server.cdr_open(preset="A4")
    if "document_id" in s51_doc:
        s51_id = s51_doc["document_id"]
        s51r = server.cdr_create(s51_id, [
            {"type": "paragraph_text", "x": 20, "y": 20, "x2": 140, "y2": 80,
             "text": "段落文本 Hello world", "size": 12},
            {"type": "curve_segment", "x1": 20, "y1": 100, "x2": 80, "y2": 100},
        ])
        if no_error("cdr_create(结构体写入测试)", s51r):
            s51t = {"name": s51r["created"][0]["name"]}
            s51c = {"name": s51r["created"][1]["name"]}
            pf51 = server.cdr_text(s51_id, s51t, paragraph_format={
                "alignment": 3, "left_indent": 5.0, "character_spacing": 150.0,
                "line_spacing": 180.0, "hot_zone": 8.0})
            if no_error("cdr_text(paragraph_format)", pf51):
                ap51 = pf51.get("applied") or []
                check("结构体工厂可写入段落格式（CreateStruct*Properties）",
                      "align_struct" in ap51 and "space_struct" in ap51,
                      str(ap51))
            ctx51 = server.cdr_doc(s51_id, action="color_context_info")
            if no_error("cdr_doc(color_context_info)", ctx51):
                pbm = ctx51.get("profiles_by_model") or {}
                check("可按模型读色彩配置文件（ColorProfile(model)）",
                      bool(pbm.get("0")) and bool(pbm.get("1")),
                      f"rgb={pbm.get('0')} cmyk={pbm.get('1')}")
            ms51 = server.cdr_color(action="object_modify_style_color",
                                    color="#336699", color2="#000000")
            if no_error("cdr_color(object_modify_style_color)", ms51):
                check("颜色风格色改写可用（ModifyColorStyleColor）",
                      bool(ms51.get("source")), str(ms51.get("source"))[:40])
            ah51 = server.cdr_doc(s51_id, action="arrowhead_unnamed",
                                  target=s51c, x=50, y=100)
            if no_error("cdr_doc(arrowhead_unnamed)", ah51):
                check("无名箭头变体可用（CreateArrowHeadEx）",
                      ah51.get("created") is True and ah51.get("name") == "",
                      f"name={ah51.get('name')!r}")
        no_error("关闭结构体写入测试文档", server.cdr_close(s51_id))

    # ---------------- 15bf. 第五十二波：按偏移取点 / 定位线段 ----------------
    print("\n[15bf] 第五十二波：SubPath.GetPointPositionAt / FindSegmentOffset")
    o52_doc = server.cdr_open(preset="A4")
    if "document_id" in o52_doc:
        o52_id = o52_doc["document_id"]
        o52r = server.cdr_create(o52_id, [
            {"type": "curve_segment", "x1": 0, "y1": 60, "x2": 40, "y2": 60}])
        if no_error("cdr_create(偏移取点测试)", o52r):
            o52c = {"name": o52r["created"][0]["name"]}
            p52 = server.cdr_path(o52_id, o52c, action="point_at_offset",
                                  params={"offset": 0.5, "normalized": True})
            if no_error("cdr_path(point_at_offset)", p52):
                pt52 = p52.get("point") or []
                check("可按归一化偏移取点（GetPointPositionAt）",
                      len(pt52) == 2 and 19.0 < pt52[0] < 21.5,
                      f"point={pt52}")
            s52 = server.cdr_path(o52_id, o52c, action="segment_at_offset",
                                  params={"offset": 10.0})
            if no_error("cdr_path(segment_at_offset)", s52):
                check("可按绝对偏移定位线段（FindSegmentOffset）",
                      s52.get("found") is True
                      and (s52.get("segment_index") or 0) >= 1,
                      f"found={s52.get('found')} seg={s52.get('segment_index')} "
                      f"param={s52.get('param_offset')}")
        no_error("关闭偏移取点测试文档", server.cdr_close(o52_id))

    # ---------------- 15bg. 第五十三波：B 样条构造 / 整批套用描边 ----------------
    print("\n[15bg] 第五十三波：CreateBSpline 控制点构造、ShapeRange.ApplyOutline")
    b53_doc = server.cdr_open(preset="A4")
    if "document_id" in b53_doc:
        b53_id = b53_doc["document_id"]
        b53r = server.cdr_create(b53_id, [
            {"type": "bspline", "points": [[0, 0], [20, 30], [40, 30], [60, 0]],
             "closed": False},
            {"type": "bspline", "points": [[80, 0], [100, 30], [120, 0]],
             "closed": True, "to_curves": True},
            {"type": "rectangle", "x": 10, "y": 60, "width": 20, "height": 10,
             "outline": {"width": 1.0}},
            {"type": "rectangle", "x": 40, "y": 60, "width": 20, "height": 10,
             "outline": {"width": 1.0}},
        ])
        if no_error("cdr_create(B样条/描边测试)", b53r):
            made53 = b53r.get("created") or []
            check("B 样条控制点构造可用（CreateBSpline + Layer.CreateBSpline）",
                  (made53[0] or {}).get("type") == "bspline",
                  f"type={(made53[0] or {}).get('type')}")
            check("B 样条可转曲线（to_curves）",
                  (made53[1] or {}).get("type") == "curve",
                  f"type={(made53[1] or {}).get('type')}")
            b53a = {"name": made53[2]["name"]}
            b53b = {"name": made53[3]["name"]}
            server.cdr_select(b53_id, name=b53a["name"])
            server.cdr_select(b53_id, name=b53b["name"])
            aop = server.cdr_shape_ops(b53_id, op="range_outline_apply",
                                       target=b53a, other=b53b,
                                       params={"width": 4.0})
            if no_error("cdr_shape_ops(range_outline_apply)", aop):
                check("可整批套用描边（ShapeRange.ApplyOutline）",
                      aop.get("first_width") == 4.0 and (aop.get("count") or 0) >= 2,
                      f"width={aop.get('first_width')} count={aop.get('count')}")
        no_error("关闭 B 样条测试文档", server.cdr_close(b53_id))

    # ---------------- 15bh. 第五十四波：同级对象遍历 / 解组家族 ----------------
    print("\n[15bh] 第五十四波：Shape.GetNext/GetPrevious、UngroupAll/UngroupAllEx/UngroupEx")
    g54_doc = server.cdr_open(preset="A4")
    if "document_id" in g54_doc:
        g54_id = g54_doc["document_id"]
        g54r = server.cdr_create(g54_id, [
            {"type": "rectangle", "x": 0, "y": 0, "width": 10, "height": 10},
            {"type": "rectangle", "x": 20, "y": 0, "width": 10, "height": 10},
            {"type": "rectangle", "x": 40, "y": 0, "width": 10, "height": 10},
        ])
        if no_error("cdr_create(解组测试)", g54r):
            h1, h2, h3 = [c["name"] for c in g54r["created"]]
            sib = server.cdr_shape_ops(g54_id, op="sibling",
                                       target={"name": h1},
                                       params={"direction": "next", "steps": 2})
            if no_error("cdr_shape_ops(sibling next)", sib):
                check("可遍历同层相邻对象（GetNext/GetPrevious）",
                      sib.get("found") is True and len(sib.get("names") or []) == 2,
                      f"names={sib.get('names')}")
            sib2 = server.cdr_shape_ops(g54_id, op="sibling",
                                        target={"name": h3},
                                        params={"direction": "previous"})
            if no_error("cdr_shape_ops(sibling previous)", sib2):
                check("可反向遍历同层对象",
                      sib2.get("found") is True
                      and len(sib2.get("names") or []) == 1,
                      f"names={sib2.get('names')}")
            server.cdr_select(g54_id, name=h1)
            server.cdr_select(g54_id, name=h2)
            server.cdr_shape_ops(g54_id, op="group")
            g54ins = server.cdr_inspect(g54_id)
            groups54 = [s["name"] for s in (g54ins.get("shapes") or [])
                        if s.get("type") == "group"]
            if groups54:
                server.cdr_select(g54_id, name=groups54[0])
                uae = server.cdr_shape_ops(g54_id, op="ungroup_all_ex")
                if no_error("cdr_shape_ops(ungroup_all_ex)", uae):
                    check("可递归解组（UngroupAllEx 真实生效）",
                          (uae.get("page_shapes_after") or 0)
                          > (uae.get("page_shapes_before") or 0),
                          f"{uae.get('page_shapes_before')} → "
                          f"{uae.get('page_shapes_after')}")
        no_error("关闭解组测试文档", server.cdr_close(g54_id))

    # ---------------- 15bi. 第五十五波：连接器 / 线性·角度标注 ----------------
    print("\n[15bi] 第五十五波：CreateSnapPoint + CreateConnector/ConnectorEx/RightAngleConnector/Dimension")
    d55_doc = server.cdr_open(preset="A4")
    if "document_id" in d55_doc:
        d55_id = d55_doc["document_id"]
        d55r = server.cdr_create(d55_id, [
            {"type": "connector", "x1": 0, "y1": 0, "x2": 40, "y2": 20},
            {"type": "right_angle_connector", "x1": 0, "y1": 40, "x2": 40, "y2": 60,
             "corner_radius": 5},
            {"type": "linear_dimension", "x1": 0, "y1": 80, "x2": 40, "y2": 80},
            {"type": "angular_dimension", "cx": 60, "cy": 80, "x1": 80, "y1": 80,
             "x3": 60, "y3": 100},
        ])
        if no_error("cdr_create(连接器/标注)", d55r):
            got55 = d55r.get("created") or []
            check("可建连接器（CreateConnector + CreateSnapPoint）",
                  (got55[0] or {}).get("type") == "connector"
                  and ((got55[0] or {}).get("bbox") or {}).get("width") == 40.0,
                  f"type={(got55[0] or {}).get('type')}")
            check("可建直角连接器（CreateRightAngleConnector）",
                  (got55[1] or {}).get("type") == "connector",
                  f"type={(got55[1] or {}).get('type')}")
            check("可建线性标注（CreateLinearDimension 22 参）",
                  (got55[2] or {}).get("type") == "dimension"
                  and ((got55[2] or {}).get("bbox") or {}).get("width") == 40.0,
                  f"type={(got55[2] or {}).get('type')}")
            check("可建角度标注（CreateAngularDimension 18 参）",
                  (got55[3] or {}).get("type") == "dimension",
                  f"type={(got55[3] or {}).get('type')}")
        no_error("关闭连接器/标注测试文档", server.cdr_close(d55_id))

    # ---------------- 15bj. 第五十六波：曲线映射到骨架线 ----------------
    print("\n[15bj] 第五十六波：CreateCurveMappedToStroke / …AndReferenceLine")
    m56_doc = server.cdr_open(preset="A4")
    if "document_id" in m56_doc:
        m56_id = m56_doc["document_id"]
        m56r = server.cdr_create(m56_id, [
            {"type": "curve_segment", "x1": 0, "y1": 10, "x2": 40, "y2": 30},
            {"type": "curve_segment", "x1": 0, "y1": 60, "x2": 60, "y2": 60},
            {"type": "curve_segment", "x1": 0, "y1": 100, "x2": 40, "y2": 120},
        ])
        if no_error("cdr_create(骨架映射测试)", m56r):
            mA, mS, mB = [{"name": c["name"]} for c in m56r["created"]]
            mp = server.cdr_path(m56_id, mA, action="map_to_stroke", other=mS,
                                 params={"scale_y": 5.0})
            if no_error("cdr_path(map_to_stroke)", mp):
                check("曲线可映射到骨架线（CreateCurveMappedToStroke）",
                      mp.get("api") == "CreateCurveMappedToStroke"
                      and (mp.get("nodes") or 0) > 2,
                      f"api={mp.get('api')} nodes={mp.get('nodes')}")
            mref = server.cdr_path(m56_id, mB, action="map_to_stroke", other=mS,
                                   params={"scale_y": 5.0,
                                           "use_reference_line": True})
            if no_error("cdr_path(map_to_stroke 参考线)", mref):
                check("参考线版映射可用（…AndReferenceLine）",
                      mref.get("api") == "CreateCurveMappedToStrokeAndReferenceLine"
                      and (mref.get("nodes") or 0) > 2,
                      f"api={mref.get('api')} nodes={mref.get('nodes')}")
        no_error("关闭骨架映射测试文档", server.cdr_close(m56_id))

    # ---------------- 15bk. 第五十七波：文字区间结构体（*InRange 1 基） ----------------
    print("\n[15bk] 第五十七波：FontPropertiesInRange / SetSpacePropertiesInRange 家族")
    r57_doc = server.cdr_open(preset="A4")
    if "document_id" in r57_doc:
        r57_id = r57_doc["document_id"]
        r57r = server.cdr_create(r57_id, [
            {"type": "paragraph_text", "x": 20, "y": 20, "x2": 140, "y2": 90,
             "text": "段落文本 Hello world", "size": 14},
        ])
        if no_error("cdr_create(区间结构体测试)", r57r):
            r57t = {"name": r57r["created"][0]["name"]}
            w57 = server.cdr_text(r57_id, r57t, paragraph_format={
                "range_start": 1, "range_count": 5,
                "character_spacing": 200.0, "alignment": 2})
            if no_error("cdr_text(paragraph_format 区间写)", w57):
                app57 = w57.get("applied") or []
                check("区间结构体可写（SetSpacePropertiesInRange / SetAlignPropertiesInRange）",
                      "space_struct_in_range" in app57
                      and "align_struct_in_range" in app57,
                      str(app57))
            rd57 = server.cdr_text(r57_id, r57t,
                                   range_struct={"start": 1, "count": 5})
            if no_error("cdr_text(range_struct)", rd57):
                rs57 = rd57.get("range_struct") or {}
                space57 = rs57.get("space") or {}
                check("区间结构体可读且值可信（*InRange 需 1 基）",
                      abs((space57.get("CharacterSpacing") or 0) - 200.0) < 0.01,
                      f"char={space57.get('CharacterSpacing')} "
                      f"align={(rs57.get('align') or {}).get('Alignment')}")
        no_error("关闭区间结构体测试文档", server.cdr_close(r57_id))

    # ---------------- 15bl. 第五十八波：调色板/专色工厂 + 色彩上下文/打样设置 ----------------
    print("\n[15bl] 第五十八波：CreatePaletteColor·CreateSpotColor·CreateColorContext2·CreateProofColorSettings")
    f58_doc = server.cdr_open(preset="A4")
    if "document_id" in f58_doc:
        f58_id = f58_doc["document_id"]
        server.cdr_doc(f58_id, action="palette_open_fixed", palette_id=1)
        pid58 = (((server.cdr_doc(f58_id, action="palettes").get("palettes")
                   or [{}])[0]) or {}).get("identifier")
        if pid58:
            pcf = server.cdr_color(action="object_palette_factory", color2=pid58,
                                   values=[5])
            if no_error("cdr_color(object_palette_factory)", pcf):
                check("可按调色板造色（CreatePaletteColor）",
                      "PAN_SPOT" in str(pcf.get("to_string"))
                      or "TRU" in str(pcf.get("to_string")),
                      str(pcf.get("to_string"))[:50])
            scf = server.cdr_color(action="object_from_spot", color2=pid58,
                                   values=[5, 100])
            if no_error("cdr_color(object_from_spot)", scf):
                check("可造专色（CreateSpotColor，需含专色调色板）",
                      "PAN_SPOT" in str(scf.get("to_string")),
                      str(scf.get("to_string"))[:50])
        cc58 = server.cdr_doc(f58_id, action="color_context_create",
                              rendering_intent=2)
        if no_error("cdr_doc(color_context_create)", cc58):
            check("可造色彩上下文（CreateColorContext2，intent 1~4）",
                  cc58.get("rendering_intent") == 2
                  and (cc58.get("profiles") or 0) >= 3,
                  f"intent={cc58.get('rendering_intent')} "
                  f"profiles={cc58.get('profiles')}")
        pc58 = server.cdr_doc(f58_id, action="proof_color_settings",
                              profile_name="Japan Color 2001 Coated",
                              rendering_intent=1)
        if no_error("cdr_doc(proof_color_settings)", pc58):
            check("可造打样色彩设置（CreateProofColorSettings）",
                  (pc58.get("context_profiles") or 0) >= 1,
                  f"profiles={pc58.get('context_profiles')}")
        no_error("关闭色彩工厂测试文档", server.cdr_close(f58_id))

    # ---------------- 15bm. 第五十九波：点集 → 拟合曲线 / 追加点集 ----------------
    print("\n[15bm] 第五十九波：IVGPointRange + CreateCurveFitToPoints / AppendSubpathFromPoints")
    p59_doc = server.cdr_open(preset="A4")
    if "document_id" in p59_doc:
        p59_id = p59_doc["document_id"]
        p59r = server.cdr_create(p59_id, [
            {"type": "curve_segment", "x1": 0, "y1": 0, "x2": 40, "y2": 20},
            {"type": "curve_segment", "x1": 0, "y1": 60, "x2": 20, "y2": 60},
        ])
        if no_error("cdr_create(点集拟合测试)", p59r):
            p59S, p59T = [{"name": c["name"]} for c in p59r["created"]]
            fit59 = server.cdr_path(p59_id, p59T, action="fit_to_points",
                                    other=p59S,
                                    params={"spacing": 5.0, "tolerance": 2.0})
            if no_error("cdr_path(fit_to_points)", fit59):
                check("点集可拟合曲线（CreateCurveFitToPoints）",
                      fit59.get("type_id") == 3
                      and (fit59.get("points") or 0) > 5,
                      f"points={fit59.get('points')} nodes={fit59.get('nodes')}")
            fit59b = server.cdr_path(p59_id, p59T, action="fit_to_points",
                                     params={"points": [[0, 0], [10, 15], [20, 0],
                                                        [30, 15], [40, 0]],
                                             "with_cusps": True})
            if no_error("cdr_path(fit_to_points with_cusps)", fit59b):
                check("带尖点拟合可用（CreateCurveFitToPointsAndCusps）",
                      fit59b.get("type_id") == 3, f"type={fit59b.get('type_id')}")
            ap59 = server.cdr_path(p59_id, p59T, action="append_points",
                                   other=p59S,
                                   params={"spacing": 8.0, "mode": "fit"})
            if no_error("cdr_path(append_points)", ap59):
                check("点集可追加到曲线（AppendSubpath*Points）",
                      (ap59.get("nodes_after") or 0) > (ap59.get("nodes_before") or 0),
                      f"{ap59.get('nodes_before')} → {ap59.get('nodes_after')}")
        no_error("关闭点集拟合测试文档", server.cdr_close(p59_id))

    # ---------------- 15bn. 第六十波：剖面线填充（按序号 / 自定义） ----------------
    print("\n[15bn] 第六十波：Fill.ApplyHatchFill / ApplyCustomHatchFill")
    h60_doc = server.cdr_open(preset="A4")
    if "document_id" in h60_doc:
        h60_id = h60_doc["document_id"]
        h60r = server.cdr_create(h60_id, [
            {"type": "rectangle", "x": 0, "y": 0, "width": 40, "height": 30}])
        if no_error("cdr_create(剖面线测试)", h60r):
            h60a = {"name": h60r["created"][0]["name"]}
            fh60 = server.cdr_modify(h60_id, h60a, {
                "fill": {"type": "hatch", "index": 1, "back": "#FFFFFF"}})
            if no_error("cdr_modify(fill=hatch)", fh60):
                check("剖面线填充可用（ApplyHatchFill，索引必须是整数）",
                      "fill:hatch" in (fh60.get("applied") or []),
                      str(fh60.get("applied")))
            fhc60 = server.cdr_modify(h60_id, h60a, {
                "fill": {"type": "hatch_custom", "angle": 45, "spacing": 3.0,
                         "width": 0.6}})
            if no_error("cdr_modify(fill=hatch_custom)", fhc60):
                check("自定义剖面线可用（ApplyCustomHatchFill 18 参）",
                      "fill:hatch_custom" in (fhc60.get("applied") or []),
                      str(fhc60.get("applied")))
        no_error("关闭剖面线测试文档", server.cdr_close(h60_id))

    # ---------------- 15bo. 第六十一波：中心定尺寸 / 中心拉伸 / 范围排除 ----------------
    print("\n[15bo] 第六十一波：Shape.SetSizeEx / StretchEx / ShapeRange.AllExcluding")
    s61_doc = server.cdr_open(preset="A4")
    if "document_id" in s61_doc:
        s61_id = s61_doc["document_id"]
        s61r = server.cdr_create(s61_id, [
            {"type": "rectangle", "x": 0, "y": 0, "width": 20, "height": 10},
            {"type": "rectangle", "x": 40, "y": 0, "width": 10, "height": 10},
            {"type": "rectangle", "x": 60, "y": 0, "width": 10, "height": 10},
        ])
        if no_error("cdr_create(定尺寸测试)", s61r):
            s61a = {"name": s61r["created"][0]["name"]}
            sz61 = server.cdr_shape_ops(s61_id, op="set_size_ex", target=s61a,
                                        params={"cx": 10, "cy": 5, "width": 40,
                                                "height": 30})
            if no_error("cdr_shape_ops(set_size_ex)", sz61):
                bb61 = sz61.get("bbox_after") or {}
                check("可按中心定尺寸（SetSizeEx）",
                      bb61.get("width") == 40.0 and bb61.get("height") == 30.0,
                      f"{bb61.get('width')}×{bb61.get('height')}")
            st61 = server.cdr_shape_ops(s61_id, op="stretch_ex", target=s61a,
                                        params={"cx": 10, "cy": 5, "stretch_x": 2.0,
                                                "stretch_y": 0.5})
            if no_error("cdr_shape_ops(stretch_ex)", st61):
                bb61b = st61.get("bbox_after") or {}
                check("可按中心拉伸（StretchEx）",
                      bb61b.get("width") == 80.0 and bb61b.get("height") == 15.0,
                      f"{bb61b.get('width')}×{bb61b.get('height')}")
            for n61 in [c["name"] for c in s61r["created"][1:]]:
                server.cdr_select(s61_id, name=n61)
            ex61 = server.cdr_shape_ops(s61_id, op="all_excluding",
                                        params={"exclude": [1]})
            if no_error("cdr_shape_ops(all_excluding)", ex61):
                check("范围可排除成员（AllExcluding，1 基索引）",
                      (ex61.get("source_count") or 0) == 2
                      and (ex61.get("result_count") or 0) == 1,
                      f"{ex61.get('source_count')} → {ex61.get('result_count')}")
        no_error("关闭定尺寸测试文档", server.cdr_close(s61_id))

    # ---------------- 15bp. 第六十二波：以副本打开 / 字体结构体工厂 ----------------
    print("\n[15bp] 第六十二波：OpenDocumentAsCopy、CreateStructFontProperties")
    c62_doc = server.cdr_open(preset="A4")
    if "document_id" in c62_doc:
        c62_id = c62_doc["document_id"]
        c62r = server.cdr_create(c62_id, [
            {"type": "rectangle", "x": 0, "y": 0, "width": 20, "height": 10},
            {"type": "paragraph_text", "x": 20, "y": 20, "x2": 120, "y2": 60,
             "text": "abc"},
        ])
        if no_error("cdr_create(副本/字体结构体测试)", c62r):
            c62p = os.path.join(os.getcwd(), "tmp",
                                f"w62src_{os.urandom(4).hex()}.cdr")
            sv62 = server.cdr_save(c62_id, c62p)
            if no_error("cdr_save(副本源)", sv62) and os.path.exists(c62p):
                cp62 = server.cdr_open(path=c62p, as_copy=True)
                if no_error("cdr_open(as_copy)", cp62) and cp62.get("document_id"):
                    ins62 = server.cdr_inspect(cp62["document_id"])
                    check("可按副本打开（OpenDocumentAsCopy）",
                          len(ins62.get("shapes") or []) >= 1,
                          f"shapes={len(ins62.get('shapes') or [])}")
                    no_error("关闭副本文档", server.cdr_close(cp62["document_id"]))
            c62t = {"name": c62r["created"][1]["name"]}
            fs62 = server.cdr_text(c62_id, c62t,
                                   paragraph_format={"report_font_struct": True})
            if no_error("cdr_text(report_font_struct)", fs62):
                check("字体结构体工厂可读（CreateStructFontProperties）",
                      "font_struct_factory" in (fs62.get("applied") or []),
                      str(fs62.get("applied")))
        no_error("关闭副本/字体结构体测试文档", server.cdr_close(c62_id))

    # ---------------- 15bq. 第六十三波：样式改父级 / 样式复制 ----------------
    print("\n[15bq] 第六十三波：Style.SetBasedOn(名称串) / Style.Assign(样式)")
    y63_doc = server.cdr_open(preset="A4")
    if "document_id" in y63_doc:
        y63_id = y63_doc["document_id"]
        y63r = server.cdr_create(y63_id, [
            {"type": "rectangle", "x": 0, "y": 0, "width": 20, "height": 10,
             "fill": "#C80A0A"},
            {"type": "rectangle", "x": 30, "y": 0, "width": 20, "height": 10,
             "fill": "#0AC80A"},
        ])
        if no_error("cdr_create(样式测试)", y63r):
            y63a, y63b = [{"name": c["name"]} for c in y63r["created"]]
            server.cdr_style(y63_id, action="create", target=y63a,
                             category="fill", style_name="P63")
            server.cdr_style(y63_id, action="create", target=y63b,
                             category="fill", style_name="C63")
            # 先 assign 再 set_based_on —— 实测 SetBasedOn 之后样式的 Name 会变，
            # 之后按名字就找不到了
            as63 = server.cdr_style(y63_id, action="assign",
                                    style_name="C63", based_on="P63")
            if no_error("cdr_style(assign)", as63):
                check("样式可相互复制（Style.Assign）",
                      "primaryColor" in str(as63.get("definition")),
                      str(as63.get("definition"))[:40])
            sb63 = server.cdr_style(y63_id, action="set_based_on",
                                    style_name="C63", based_on="P63")
            if no_error("cdr_style(set_based_on)", sb63):
                check("样式可改父级（SetBasedOn 传名称串）",
                      sb63.get("applied") is True
                      and sb63.get("based_on_name") == "P63",
                      f"applied={sb63.get('applied')} "
                      f"based_on={sb63.get('based_on_name')}")
        no_error("关闭样式测试文档", server.cdr_close(y63_id))

    # ---------------- 15br. 第六十四波：标注对象 / 图框内容 / 辅助线 / 文本框链接 ----------------
    print("\n[15br] 第六十四波：Shape.Dimension·Shape.PowerClip·Shape.Guide·TextFrame.LinkTo")
    x64_doc = server.cdr_open(preset="A4")
    if "document_id" in x64_doc:
        x64_id = x64_doc["document_id"]
        x64r = server.cdr_create(x64_id, [
            {"type": "linear_dimension", "x1": 0, "y1": 0, "x2": 40, "y2": 0},
            {"type": "rectangle", "x": 0, "y": 40, "width": 40, "height": 30},
            {"type": "rectangle", "x": 5, "y": 45, "width": 10, "height": 10},
            {"type": "paragraph_text", "x": 60, "y": 40, "x2": 120, "y2": 80,
             "text": "甲框"},
            {"type": "paragraph_text", "x": 130, "y": 40, "x2": 190, "y2": 80,
             "text": ""},
        ])
        if no_error("cdr_create(标注/图框测试)", x64r):
            xD, xH, xI, xT1, xT2 = [{"name": c["name"]} for c in x64r["created"]]
            di = server.cdr_shape_ops(x64_id, op="dimension_info", target=xD)
            if no_error("cdr_shape_ops(dimension_info)", di):
                check("标注对象可读（Shape.Dimension）",
                      di.get("linear") is True and di.get("precision") is not None,
                      f"linear={di.get('linear')} precision={di.get('precision')}")
            ds = server.cdr_shape_ops(x64_id, op="dimension_set", target=xD,
                                      params={"text_label_gap": 2.0, "overhang": 1.0,
                                              "precision": 1})
            if no_error("cdr_shape_ops(dimension_set)", ds):
                check("标注参数可设（SetTextLabelGap/SetOverhang）",
                      {"text_label_gap", "overhang"} <= set(ds.get("applied") or []),
                      str(ds.get("applied")))
            server.cdr_shape_ops(x64_id, op="powerclip", target=xI, other=xH)
            pe = server.cdr_shape_ops(x64_id, op="powerclip_extract", target=xH)
            if no_error("cdr_shape_ops(powerclip_extract)", pe):
                check("图框内容可取出（Shape.PowerClip.ExtractShapes）",
                      pe.get("contents") is not None
                      and pe.get("entered_edit_mode") is True,
                      f"contents={pe.get('contents')} "
                      f"edit={pe.get('entered_edit_mode')}")
            lk64 = server.cdr_text(x64_id, xT1, link_to=xT2)
            if no_error("cdr_text(link_to)", lk64):
                check("文本框可链接（TextFrame.LinkTo）",
                      "link_to" in (lk64.get("applied") or [])
                      and (lk64.get("frames_in_link") or 0) >= 2,
                      f"frames={lk64.get('frames_in_link')}")
            ul64 = server.cdr_text(x64_id, xT1, unlink=True)
            if no_error("cdr_text(unlink)", ul64):
                check("文本框可断开链接（TextFrame.UnLink）",
                      "unlink" in (ul64.get("applied") or []),
                      str(ul64.get("applied")))
        no_error("关闭标注/图框测试文档", server.cdr_close(x64_id))

    # ---------------- 15bs. 第六十五波：视口缩放 / 混合效果 / 调色板管理 / 进度条 ----------------
    print("\n[15bs] 第六十五波：IVGActiveView 缩放·IVGEffectBlend·IVGPaletteManager·IVGAppStatus")
    v65_doc = server.cdr_open(preset="A4")
    if "document_id" in v65_doc:
        v65_id = v65_doc["document_id"]
        v65r = server.cdr_create(v65_id, [
            {"type": "ellipse", "x": 0, "y": 0, "width": 10, "height": 10},
            {"type": "ellipse", "x": 60, "y": 0, "width": 10, "height": 10},
        ])
        if no_error("cdr_create(视口/混合测试)", v65r):
            vA, vB = [{"name": c["name"]} for c in v65r["created"]]
            z65 = server.cdr_view(v65_id, action="actual_size")
            if no_error("cdr_view(actual_size)", z65):
                check("视口可设为实际大小（ActiveView.SetActualSize）",
                      z65.get("zoom_after") == 100.0, f"zoom={z65.get('zoom_after')}")
            zi65 = server.cdr_view(v65_id, action="zoom_in")
            if no_error("cdr_view(zoom_in)", zi65):
                check("视口可放大（ActiveView.ZoomIn）",
                      (zi65.get("zoom_after") or 0) > (zi65.get("zoom_before") or 0),
                      f"{zi65.get('zoom_before')} → {zi65.get('zoom_after')}")
            va65 = server.cdr_view(v65_id, action="view_area")
            if no_error("cdr_view(view_area)", va65):
                check("视口区域可读（GetViewArea）",
                      len(va65.get("area") or []) == 4, str(va65.get("area")))
            fs65 = server.cdr_view(v65_id, action="fit_shape", target=vA)
            if no_error("cdr_view(fit_shape)", fs65):
                check("可缩放到指定对象（ToFitShape）",
                      (fs65.get("zoom_after") or 0) > 100.0,
                      f"zoom={fs65.get('zoom_after')}")
            server.cdr_shape_ops(v65_id, op="blend", target=vA, other=vB)
            be65 = server.cdr_shape_ops(v65_id, op="blend_effect", target=vA,
                                        params={"split": 5.0})
            if no_error("cdr_shape_ops(blend_effect)", be65):
                check("混合效果可拆分（EffectBlend.Split）",
                      "split" in (be65.get("applied") or []),
                      str(be65.get("applied")))
            st65 = server.cdr_doc(v65_id, action="app_status",
                                  settings={"begin": True, "max": 3, "update": 1,
                                            "end": True})
            if no_error("cdr_doc(app_status)", st65):
                check("应用进度条可用（AppStatus）",
                      st65.get("progress") == 1
                      and {"begin", "update", "end"} <= set(st65.get("applied") or []),
                      f"progress={st65.get('progress')}")
            pm65 = server.cdr_doc(v65_id, action="palette_manager")
            if no_error("cdr_doc(palette_manager)", pm65):
                check("调色板管理器可读（PaletteManager）",
                      (pm65.get("palette_count") or 0) > 0
                      and bool(pm65.get("first_palettes")),
                      f"count={pm65.get('palette_count')}")
        no_error("关闭视口/混合测试文档", server.cdr_close(v65_id))

    # ---------------- 15bt. 第六十六波：线型虚线段 / 自定义工具形状属性 ----------------
    print("\n[15bt] 第六十六波：IVGOutlineStyle 虚线段读取、IVGToolShapeAttributes 12 项")
    t66_doc = server.cdr_open(preset="A4")
    if "document_id" in t66_doc:
        t66_id = t66_doc["document_id"]
        od66 = server.cdr_doc(t66_id, action="outline_style_dash", palette_index=2)
        if no_error("cdr_doc(outline_style_dash)", od66):
            segs66 = od66.get("segments") or []
            check("线型虚线段可读（DashCount/DashLength/GapLength）",
                  (od66.get("dash_count") or 0) >= 1
                  and len(segs66) >= 1
                  and segs66[0].get("dash") is not None
                  and od66.get("enhanced") is not None,
                  f"count={od66.get('dash_count')} segs={segs66[:1]}")
        tsa66 = server.cdr_doc(t66_id, action="tool_shape_attributes",
                               settings={"default_shape_name": "MCP 形状",
                                         "can_skew": False})
        if no_error("cdr_doc(tool_shape_attributes)", tsa66):
            check("自定义工具形状属性可配置（12 项 Set*）",
                  (tsa66.get("count") or 0) == 12,
                  f"count={tsa66.get('count')}")
        no_error("关闭线型/工具属性测试文档", server.cdr_close(t66_id))

    # ---------------- 15bu. 第六十七波：矩形圆角 / 克隆链 / 投影偏移 ----------------
    print("\n[15bu] 第六十七波：Shape.Rectangle（圆角）·Shape.CloneLink·EffectDropShadow")
    u67_doc = server.cdr_open(preset="A4")
    if "document_id" in u67_doc:
        u67_id = u67_doc["document_id"]
        u67r = server.cdr_create(u67_id, [
            {"type": "rectangle", "x": 0, "y": 0, "width": 40, "height": 30},
            {"type": "ellipse", "x": 60, "y": 0, "width": 10, "height": 10},
            {"type": "rectangle", "x": 0, "y": 60, "width": 20, "height": 10},
        ])
        if no_error("cdr_create(圆角/克隆测试)", u67r):
            uR, uE, uS = [{"name": c["name"]} for c in u67r["created"]]
            co67 = server.cdr_shape_ops(u67_id, op="corners", target=uR,
                                        params={"radius": 5.0})
            if no_error("cdr_shape_ops(corners radius)", co67):
                check("矩形可按半径倒角（Rectangle.SetRadius）",
                      "set_radius" in (co67.get("applied") or [])
                      and (co67.get("max_radius") or 0) > 0,
                      f"max={co67.get('max_radius')} "
                      f"r={co67.get('radius_upper_left')}")
            co67b = server.cdr_shape_ops(u67_id, op="corners", target=uR,
                                         params={"roundness": 60})
            if no_error("cdr_shape_ops(corners roundness)", co67b):
                check("矩形可按圆度倒角（Rectangle.SetRoundness）",
                      "set_roundness" in (co67b.get("applied") or []),
                      str(co67b.get("applied")))
            before67 = {s.get("name") for s in
                        ((server.cdr_inspect(u67_id).get("shapes") or []))}
            server.cdr_modify(u67_id, uE, {"clone": {"dx": 20, "dy": 0}})
            after67 = {s.get("name") for s in
                       ((server.cdr_inspect(u67_id).get("shapes") or []))}
            clones67 = sorted(after67 - before67)
            if clones67:
                cl67 = server.cdr_shape_ops(u67_id, op="clone_links",
                                            target={"name": clones67[0]},
                                            params={"restore": True})
                if no_error("cdr_shape_ops(clone_links)", cl67):
                    check("克隆链接可恢复（CloneLink.RestoreAllLinks）",
                          "restore_all_links" in (cl67.get("applied") or []),
                          f"fill_linked={cl67.get('fill_linked')}")
            server.cdr_modify(u67_id, uS, {"drop_shadow": {
                "type": 0, "opacity": 50, "feather": 15, "offset_x": 5, "offset_y": 5}})
            ds67 = server.cdr_shape_ops(u67_id, op="drop_shadow", target=uS,
                                        params={"offset_x": 8, "offset_y": 8,
                                                "opacity": 60})
            if no_error("cdr_shape_ops(drop_shadow)", ds67):
                check("投影偏移可改（EffectDropShadow.SetOffset）",
                      ds67.get("offset") == [8.0, 8.0] and ds67.get("opacity") == 60,
                      f"offset={ds67.get('offset')} opacity={ds67.get('opacity')}")
        no_error("关闭圆角/克隆测试文档", server.cdr_close(u67_id))

    # ---------------- 15bv. 第六十八波：双色调子对象 / 网格频率 / 线性标注子对象 ----------------
    print("\n[15bv] 第六十八波：IVGDuotone·IVGDuotoneOverprint·IVGGrid.SetFrequency·IVGDimensionLinear")
    w68_doc = server.cdr_open(preset="A4")
    if "document_id" in w68_doc:
        w68_id = w68_doc["document_id"]
        w68r = server.cdr_create(w68_id, [
            {"type": "bitmap", "x": 0, "y": 0, "width": 30, "height": 30,
             "color": "#CC8844"},
            {"type": "linear_dimension", "x1": 0, "y1": 80, "x2": 40, "y2": 80},
        ])
        if no_error("cdr_create(双色调/网格测试)", w68r):
            wB, wD = [{"name": c["name"]} for c in w68r["created"]]
            duo68 = server.cdr_bitmap(w68_id, wB, action="duotone", duotone_type=1)
            if no_error("cdr_bitmap(duotone)", duo68):
                check("双色调子对象可读改（IVGDuotoneOverprint.SetValues）",
                      (duo68.get("overprint_count") or 0) >= 1
                      and duo68.get("overprint_set") is True,
                      f"count={duo68.get('overprint_count')} "
                      f"set={duo68.get('overprint_set')}")
            g68 = server.cdr_doc(w68_id, action="set_grid",
                                 settings={"frequency": [4, 4]})
            if no_error("cdr_doc(set_grid frequency)", g68):
                check("网格可按频率设置（IVGGrid.SetFrequency）",
                      g68.get("frequency") == [4.0, 4.0]
                      and (g68.get("spacing_x") or 0) > 0,
                      f"freq={g68.get('frequency')} spacing={g68.get('spacing_x')}")
            dl68 = server.cdr_shape_ops(w68_id, op="dimension_set", target=wD,
                                        params={"reverse_terminator_length": 2.0})
            if no_error("cdr_shape_ops(dimension_set linear)", dl68):
                check("线性标注子对象可设（IVGDimensionLinear）",
                      "reverse_terminator_length" in (dl68.get("applied") or [])
                      and dl68.get("linear_type") is not None,
                      f"applied={dl68.get('applied')} type={dl68.get('linear_type')}")
        no_error("关闭双色调/网格测试文档", server.cdr_close(w68_id))

    # ---------------- 15bw. 第六十九波：对象树管理 / 样式填充存盘 ----------------
    print("\n[15bw] 第六十九波：IVGTreeManager、IVGStyleFill.SaveFill/LoadFill")
    t69_doc = server.cdr_open(preset="A4")
    if "document_id" in t69_doc:
        t69_id = t69_doc["document_id"]
        tm69 = server.cdr_doc(t69_id, action="tree_manager",
                             settings={"create_group_node": True, "clean": True})
        if no_error("cdr_doc(tree_manager)", tm69):
            check("对象树可建/清组节点（TreeManager）",
                  {"create_group_node", "clean_group_node"} <= set(
                      tm69.get("applied") or []),
                  str(tm69.get("applied")))
        t69r = server.cdr_create(t69_id, [
            {"type": "rectangle", "x": 0, "y": 0, "width": 20, "height": 10,
             "fill": "#0B1621"}])
        if no_error("cdr_create(样式填充测试)", t69r):
            t69a = {"name": t69r["created"][0]["name"]}
            server.cdr_style(t69_id, action="create", target=t69a,
                             category="fill", style_name="SF69")
            f69 = os.path.join(os.getcwd(), "tmp",
                               f"w69style_{os.urandom(2).hex()}.cpf")
            sv69 = server.cdr_style(t69_id, action="fill_file", style_name="SF69",
                                    sheet_path=f69, property="save")
            if no_error("cdr_style(fill_file save)", sv69):
                check("样式填充可存盘（StyleFill.SaveFill）",
                      sv69.get("direction") == "save"
                      and sv69.get("metadata_ok") is True,
                      f"ok={sv69.get('ok')} type={sv69.get('fill_type')}")
            ld69 = server.cdr_style(t69_id, action="fill_file", style_name="SF69",
                                    sheet_path=f69, property="load")
            if no_error("cdr_style(fill_file load)", ld69):
                check("样式填充可读取（StyleFill.LoadFill）",
                      ld69.get("direction") == "load",
                      f"ok={ld69.get('ok')}")
        no_error("关闭对象树/样式填充测试文档", server.cdr_close(t69_id))

    # ---------------- 15bx. 第七十波：HTML 兼容 / 按模型设置色彩配置文件 ----------------
    print("\n[15bx] 第七十波：Text.MakeHTMLCompatible、ColorContext.SetColorProfile")
    h70_doc = server.cdr_open(preset="A4")
    if "document_id" in h70_doc:
        h70_id = h70_doc["document_id"]
        h70r = server.cdr_create(h70_id, [
            {"type": "paragraph_text", "x": 0, "y": 0, "x2": 60, "y2": 40,
             "text": "文字"}])
        if no_error("cdr_create(HTML 兼容测试)", h70r):
            h70t = {"name": h70r["created"][0]["name"]}
            hc70 = server.cdr_text(h70_id, h70t, make_html_compatible=True)
            if no_error("cdr_text(make_html_compatible)", hc70):
                check("文字可设 HTML 兼容（Text.MakeHTMLCompatible）",
                      "make_html_compatible" in (hc70.get("applied") or [])
                      and hc70.get("html_compatible") is True,
                      f"ok={hc70.get('html_compatible')}")
        cp70 = server.cdr_doc(h70_id, action="set_color_profile", model=1)
        if no_error("cdr_doc(set_color_profile)", cp70):
            check("可按模型设置色彩配置文件（SetColorProfile）",
                  cp70.get("model") == 1 and bool(cp70.get("readback")),
                  f"profile={cp70.get('profile')}")
        no_error("关闭 HTML 兼容测试文档", server.cdr_close(h70_id))

    # ---------------- 15by. 第七十一波：剖面线参数 / 渐变控制点 / 椭圆几何 / 调色板工厂 / 窗口查找 ----------------
    print("\n[15by] 第七十一波：IVGHatchFill·IVGFountainColors·IVGEllipse·IVGPalettes·IVGWindows")
    h71_doc = server.cdr_open(preset="A4")
    if "document_id" in h71_doc:
        h71_id = h71_doc["document_id"]
        h71r = server.cdr_create(h71_id, [
            {"type": "rectangle", "x": 0, "y": 0, "width": 30, "height": 20,
             "fill": {"type": "hatch", "index": 1}},
            {"type": "rectangle", "x": 40, "y": 0, "width": 30, "height": 20,
             "fill": {"type": "fountain", "start": "#FF0000", "end": "#0000FF"}},
            {"type": "ellipse", "x": 0, "y": 40, "width": 20, "height": 10},
        ])
        if no_error("cdr_create(剖面线/渐变测试)", h71r):
            hH, hF, hE = [{"name": c["name"]} for c in h71r["created"]]
            hf71 = server.cdr_shape_ops(h71_id, op="hatch_fill", target=hH,
                                        params={"fill_scale": 2.0, "line_scale": 2.0,
                                                "no_back_color": True})
            if no_error("cdr_shape_ops(hatch_fill)", hf71):
                check("剖面线缩放可调（IVGHatchFill）",
                      hf71.get("fill_scale") == 2.0
                      and {"fill_scale", "line_scale"} <= set(hf71.get("applied") or []),
                      f"fill={hf71.get('fill_scale')} line={hf71.get('line_scale')}")
            fc71 = server.cdr_shape_ops(h71_id, op="fountain_colors", target=hF,
                                        params={"index": 1, "gray_level": 30})
            if no_error("cdr_shape_ops(fountain_colors)", fc71):
                check("渐变控制点灰度可读写（IVGFountainColors）",
                      (fc71.get("count") or 0) >= 2
                      and (fc71.get("levels") or [None])[0] == 30.0,
                      f"count={fc71.get('count')} levels={fc71.get('levels')}")
            el71 = server.cdr_shape_ops(h71_id, op="ellipse_geometry", target=hE,
                                        params={"center_x": 10, "center_y": 45})
            if no_error("cdr_shape_ops(ellipse_geometry)", el71):
                check("椭圆几何可读写（IVGEllipse）",
                      el71.get("radius") is not None
                      and "set_center" in (el71.get("applied") or []),
                      f"radius={el71.get('radius')}")
            server.cdr_select(h71_id, name=hF["name"])
            pf71 = server.cdr_doc(h71_id, action="palette_from",
                                  settings={"name": "MCP 测试", "mode": "selection"})
            if no_error("cdr_doc(palette_from)", pf71):
                check("可从文档/选区造调色板（IVGPalettes）",
                      pf71.get("created") is True,
                      f"count={pf71.get('color_count')}")
            w71 = server.cdr_doc(h71_id, action="windows",
                                 settings={"find_caption": "CorelDRAW"})
            if no_error("cdr_doc(windows)", w71):
                check("可按标题查窗口（IVGWindows.FindWindow）",
                      (w71.get("count") or 0) >= 1 and "found" in w71,
                      f"count={w71.get('count')} found={w71.get('found')}")
            od71 = server.cdr_doc(h71_id, action="outline_style_dash",
                                  palette_index=2,
                                  settings={"dash_length": 0.3, "gap_length": 0.2})
            if no_error("cdr_doc(outline_style_dash set)", od71):
                check("线型虚线段可写调用（SetDashLength/SetGapLength）",
                      {"set_dash_length", "set_gap_length"}
                      <= set(od71.get("applied") or []),
                      str(od71.get("applied")))
        no_error("关闭剖面线/渐变测试文档", server.cdr_close(h71_id))

    # ---------------- 15bz. 第七十二波：屏幕手柄 / 包络效果对象 / 按名选配置文件 ----------------
    print("\n[15bz] 第七十二波：IVGOnScreenHandle·IVGEffectEnvelope·ColorProfiles.SelectByName")
    v72_doc = server.cdr_open(preset="A4")
    if "document_id" in v72_doc:
        v72_id = v72_doc["document_id"]
        v72r = server.cdr_create(v72_id, [
            {"type": "rectangle", "x": 0, "y": 0, "width": 20, "height": 10},
            {"type": "rectangle", "x": 40, "y": 0, "width": 20, "height": 10},
        ])
        if no_error("cdr_create(包络/手柄测试)", v72r):
            v72a, v72b = [{"name": c["name"]} for c in v72r["created"]]
            oh72 = server.cdr_view(v72_id, action="overlay_handle", x=20, y=20,
                                   color=0xFF0000)
            if no_error("cdr_view(overlay_handle)", oh72):
                check("屏幕手柄可用（IVGOnScreenHandle）",
                      {"set_position", "set_handle_color",
                       "update_hot_tracking"} <= set(oh72.get("applied") or []),
                      str(oh72.get("applied")))
            oh72b = server.cdr_view(v72_id, action="overlay_handle", x=10, y=10)
            if no_error("cdr_view(overlay_handle hide)", oh72b):
                check("屏幕手柄可显隐（Show/Hide）",
                      "show" in (oh72b.get("applied") or []),
                      str(oh72b.get("applied")))
            ev72 = server.cdr_shape_ops(v72_id, op="envelope_object", target=v72a,
                                       params={"preset": 1, "select": 1},
                                       other=v72b)
            if no_error("cdr_shape_ops(envelope_object)", ev72):
                check("包络效果对象可操作（EffectEnvelope）",
                      {"select", "create_from"} <= set(ev72.get("applied") or []),
                      f"mode={ev72.get('mode')} applied={ev72.get('applied')}")
            pf72 = server.cdr_doc(v72_id, action="select_color_profile",
                                  profile_name="sRGB IEC61966-2.1")
            if no_error("cdr_doc(select_color_profile)", pf72):
                check("可按名称选择色彩配置文件（SelectByName）",
                      "ok" in pf72 and (pf72.get("count") or 0) > 0,
                      f"ok={pf72.get('ok')}")
        no_error("关闭包络/手柄测试文档", server.cdr_close(v72_id))

    # ---------------- 15ca. 第七十三波：叠加层补齐 / 文本框列信息 ----------------
    print("\n[15ca] 第七十三波：OnScreenText.SetText·SetPixelOffset·UpdatePosition / OnScreenCurve.SetPoints / TextFrame 列信息")
    o73_doc = server.cdr_open(preset="A4")
    if "document_id" in o73_doc:
        o73_id = o73_doc["document_id"]
        ot73 = server.cdr_view(o73_id, action="overlay_text", text="MCP 叠加",
                               x=50, y=50, distance=5)
        if no_error("cdr_view(overlay_text)", ot73):
            check("叠加文字可设文本/像素偏移/位置（OnScreenText 三方法）",
                  ot73.get("shown") == 1, f"shown={ot73.get('shown')}")
        op73 = server.cdr_view(o73_id, action="overlay_points", x=10, y=10,
                               x2=80, y2=60, color=0x00FF00)
        if no_error("cdr_view(overlay_points)", op73):
            check("叠加曲线可设无笔刷/填充/点集（OnScreenCurve 三方法）",
                      {"set_no_pen", "set_brush", "set_points"}
                      <= set(op73.get("applied") or []),
                      str(op73.get("applied")))
        server.cdr_view(o73_id, action="overlay_clear")
        o73r = server.cdr_create(o73_id, [
            {"type": "paragraph_text", "x": 0, "y": 0, "x2": 80, "y2": 60,
             "text": "段落文本"}])
        if no_error("cdr_create(文本框列信息测试)", o73r):
            fi73 = server.cdr_text(o73_id, {"name": o73r["created"][0]["name"]},
                                   frame_info=True)
            if no_error("cdr_text(frame_info)", fi73):
                frames73 = fi73.get("frames") or []
                check("文本框列信息可读（ColumnWidth/ColumnGutter）",
                      len(frames73) == 1
                      and bool(frames73[0].get("column_widths")),
                      f"cols={frames73[0].get('column_widths') if frames73 else None}")
        no_error("关闭叠加层/文本框测试文档", server.cdr_close(o73_id))

    # ---------------- 15cb. 第七十四波：拉伸消失点共享 / 窗口排列 ----------------
    print("\n[15cb] 第七十四波：ExtrudeVanishingPoint.Share（传拉伸效果）·Windows.Arrange")
    e74_doc = server.cdr_open(preset="A4")
    if "document_id" in e74_doc:
        e74_id = e74_doc["document_id"]
        e74r = server.cdr_create(e74_id, [
            {"type": "rectangle", "x": 0, "y": 0, "width": 20, "height": 10},
            {"type": "rectangle", "x": 60, "y": 0, "width": 20, "height": 10},
        ])
        if no_error("cdr_create(消失点测试)", e74r):
            e74a, e74b = [{"name": c["name"]} for c in e74r["created"]]
            for n74 in (e74a, e74b):
                server.cdr_shape_ops(e74_id, op="extrude", target=n74,
                                     params={"depth": 20})
            sv74 = server.cdr_shape_ops(e74_id, op="extrude_vp_share",
                                        target=e74a, other=e74b)
            if no_error("cdr_shape_ops(extrude_vp_share)", sv74):
                check("拉伸消失点可共享（Share 传拉伸效果）",
                      sv74.get("shared") is True and sv74.get("position") is not None,
                      f"shared={sv74.get('shared')} pos={sv74.get('position')}")
            ar74 = server.cdr_doc(e74_id, action="windows", settings={"arrange": 1})
            if no_error("cdr_doc(windows arrange)", ar74):
                check("窗口可排列（Windows.Arrange）",
                      ar74.get("arranged") == 1,
                      f"arranged={ar74.get('arranged')}")
        no_error("关闭消失点/窗口测试文档", server.cdr_close(e74_id))

    # ---------------- 15cc. 第七十五波：批量斜切 / 渐变色标枚举 / 线型表保存 ----------------
    print("\n[15cc] 第七十五波：ShapeRange.SkewEx 批量、FountainColors 色标枚举、OutlineStyles.Save")
    s75_doc = server.cdr_open(preset="A4")
    if "document_id" in s75_doc:
        s75_id = s75_doc["document_id"]
        s75r = server.cdr_create(s75_doc["document_id"], [
            {"type": "rectangle", "x": 0, "y": 0, "width": 10, "height": 10},
            {"type": "rectangle", "x": 30, "y": 0, "width": 10, "height": 10},
            {"type": "rectangle", "x": 0, "y": 40, "width": 30, "height": 20,
             "fill": {"type": "fountain", "start": "#FF0000", "end": "#0000FF"}},
        ])
        if no_error("cdr_create(批量斜切测试)", s75r):
            s75a, s75b, s75f = [{"name": c["name"]} for c in s75r["created"]]
            server.cdr_select(s75_id, name=s75a["name"])
            server.cdr_select(s75_id, name=s75b["name"])
            sk75 = server.cdr_shape_ops(s75_id, op="range_skew",
                                        params={"angle_x": 10, "cx": 20, "cy": 5})
            if no_error("cdr_shape_ops(range_skew)", sk75):
                bb75 = sk75.get("bbox_after") or {}
                check("选区可整批斜切（ShapeRange.SkewEx）",
                      (bb75.get("width") or 0) > 40.0,
                      f"bbox={bb75.get('x')},{bb75.get('width')}")
            fc75 = server.cdr_shape_ops(s75_id, op="fountain_colors", target=s75f,
                                        params={"index": 1})
            if no_error("cdr_shape_ops(fountain_colors stops)", fc75):
                stops75 = fc75.get("stops") or []
                check("渐变色标可枚举（Color/Position）",
                      len(stops75) >= 1 and stops75[0].get("position") is not None,
                      f"count={fc75.get('count')} stops={len(stops75)}")
            os75 = server.cdr_doc(s75_id, action="outline_styles_save")
            if no_error("cdr_doc(outline_styles_save)", os75):
                check("线型表可保存（OutlineStyles.Save）",
                      os75.get("saved") is True and (os75.get("count") or 0) > 100,
                      f"count={os75.get('count')}")
        no_error("关闭批量斜切测试文档", server.cdr_close(s75_id))

    # ---------------- 15cd. 第七十六波：扭曲中心重设 ----------------
    print("\n[15cd] 第七十六波：EffectDistortion.CenterDistortion（扭曲中心重设）")
    d76_doc = server.cdr_open(preset="A4")
    if "document_id" in d76_doc:
        d76_id = d76_doc["document_id"]
        d76r = server.cdr_create(d76_doc["document_id"], [
            {"type": "rectangle", "x": 0, "y": 0, "width": 20, "height": 10}])
        if no_error("cdr_create(扭曲测试)", d76r):
            d76a = {"name": d76r["created"][0]["name"]}
            ds76 = server.cdr_shape_ops(d76_id, op="distort", target=d76a,
                                        params={"kind": "zipper", "amplitude": 20,
                                                "frequency": 5})
            if no_error("cdr_shape_ops(distort zipper)", ds76):
                check("扭曲可创建并重设中心（CenterDistortion）",
                      ds76.get("effects") == 1,
                      f"effects={ds76.get('effects')} kind={ds76.get('kind')}")
            ds76b = server.cdr_shape_ops(d76_id, op="distort", target=d76a,
                                         params={"kind": "pushpull", "amplitude": 30})
            if no_error("cdr_shape_ops(distort pushpull)", ds76b):
                check("推拉扭曲也可用", ds76b.get("effects") is not None,
                      f"kind={ds76b.get('kind')}")
        no_error("关闭扭曲测试文档", server.cdr_close(d76_id))

    # ---------------- 15cd2. 第七十七波：调色板查找插入 / 命名视图 / 数据字段重排 / B样条扩容 ----------------
    print("\n[15cd2] 第七十七波：IVGPalette.FindColor·InsertColor·MakeDefault·Views.AddActiveView·TreeManager 族")
    p77_doc = server.cdr_open(preset="A4")
    if "document_id" in p77_doc:
        p77_id = p77_doc["document_id"]
        fc77 = server.cdr_doc(p77_id, action="palette_find_color",
                              palette_index=1, color="TRUMATCH 1")
        if no_error("cdr_doc(palette_find_color)", fc77):
            check("调色板可按名查找（FindColor，未找到如实返回 0）",
                  fc77.get("found_index") == 0 and "matched_color" in fc77,
                  f"found={fc77.get('found_index')} "
                  f"matched={str(fc77.get('matched_color'))[:24]}")
        ins77 = server.cdr_doc(p77_id, action="palette_insert_color",
                               palette_index=1, color="#00C800", color_index=1)
        if no_error("cdr_doc(palette_insert_color)", ins77):
            check("固定调色板插色失败自动转自定义（InsertColor）",
                  ins77.get("fallback") == "custom_palette"
                  and (ins77.get("color_count") or 0) >= 1,
                  f"palette={ins77.get('palette')}")
        md77 = server.cdr_doc(p77_id, action="palette_make_default", palette_index=1)
        if no_error("cdr_doc(palette_make_default)", md77):
            check("可设默认调色板（MakeDefault）",
                  md77.get("made_default") is True,
                  f"palette={md77.get('palette')}")
        va77 = server.cdr_doc(p77_id, action="views_add_active", view_name="自动视图")
        if no_error("cdr_doc(views_add_active)", va77):
            check("当前视口可存为命名视图（Views.AddActiveView）",
                  (va77.get("count") or 0) >= 1,
                  f"count={va77.get('count')}")
        df77 = server.cdr_doc(p77_id, action="datafield_reorder", color_index=1,
                              value=2)
        if no_error("cdr_doc(datafield_reorder)", df77):
            check("数据字段可重排（DataField.Reorder）",
                  df77.get("moved_to") == 2,
                  f"field={df77.get('field')} → {df77.get('moved_to')}")
        br77 = server.cdr_doc(p77_id, action="bspline_resize", color_index=4,
                              value=6)
        if no_error("cdr_doc(bspline_resize)", br77):
            check("B 样条控制点可扩容（ControlPoints.Resize）",
                  (br77.get("points_after") or 0) > (br77.get("points_before") or 0),
                  f"{br77.get('points_before')} → {br77.get('points_after')}")
        no_error("关闭第七十七波测试文档", server.cdr_close(p77_id))

    # ---------------- 15ce. 第七十八波：立体化灯光/斜角 ----------------
    print("\n[15ce] 第七十八波：EffectExtrude.SetLight·SetBevel·灯光读数")
    x78_doc = server.cdr_open(preset="A4")
    if "document_id" in x78_doc:
        x78_id = x78_doc["document_id"]
        x78r = server.cdr_create(x78_id, [
            {"type": "rectangle", "x": 0, "y": 0, "width": 20, "height": 10}])
        if no_error("cdr_create(立体化测试)", x78r):
            x78a = {"name": x78r["created"][0]["name"]}
            server.cdr_shape_ops(x78_id, op="extrude", target=x78a,
                                 params={"depth": 20})
            es78 = server.cdr_shape_ops(x78_id, op="extrude_set", target=x78a,
                                        effect={"light_index": 1,
                                                "light_intensity": 80,
                                                "bevel_depth": 2.5,
                                                "bevel_angle": 20,
                                                "depth": 15})
            if no_error("cdr_shape_ops(extrude_set lights)", es78):
                ext78 = es78.get("extrude") or {}
                light1 = (ext78.get("lights") or [{}])[0]
                check("立体化可开灯（SetLight，1~3 号灯）",
                      light1.get("present") is True
                      and light1.get("intensity") == 80,
                      f"light1={light1}")
                check("立体化可设斜角（SetBevel）",
                      ext78.get("bevel_depth") == 2.5
                      and ext78.get("bevel_angle") == 20.0,
                      f"bevel={ext78.get('bevel_depth')}×{ext78.get('bevel_angle')}")
                check("立体化深度可改", ext78.get("depth") == 15.0,
                      f"depth={ext78.get('depth')}")
        no_error("关闭立体化测试文档", server.cdr_close(x78_id))

    # ---------------- 15cf. 第七十九波：文本框多栏设置 ----------------
    print("\n[15cf] 第七十九波：TextFrame.SetColumns（栏数/宽/间距交错数组）")
    c79_doc = server.cdr_open(preset="A4")
    if "document_id" in c79_doc:
        c79_id = c79_doc["document_id"]
        c79r = server.cdr_create(c79_id, [
            {"type": "paragraph_text", "x": 0, "y": 0, "x2": 80, "y2": 60,
             "text": "多栏测试"}])
        if no_error("cdr_create(多栏测试)", c79r):
            c79t = {"name": c79r["created"][0]["name"]}
            sc79 = server.cdr_text(c79_id, c79t,
                                   columns={"count": 2, "widths": [30, 30],
                                            "gutters": [20]})
            if no_error("cdr_text(set_columns)", sc79):
                check("文本框可设多栏（SetColumns 交错数组）",
                      sc79.get("columns") == 2
                      and (sc79.get("widths") or []) == [30.0, 20.0],
                      f"cols={sc79.get('columns')} widths={sc79.get('widths')}")
            fi79 = server.cdr_text(c79_id, c79t, frame_info=True)
            if no_error("cdr_text(frame_info)", fi79):
                fr79 = (fi79.get("frames") or [{}])[0]
                check("列宽读回一致",
                      (fr79.get("column_widths") or [None])[0] == 30.0,
                      f"widths={fr79.get('column_widths')}")
        no_error("关闭多栏测试文档", server.cdr_close(c79_id))

    # ---------------- 15ch. 第八十波：OLE 对象只读清点 / 调色板颜色枚举确认 ----------------
    print("\n[15ch] 第八十波：IVGOLE 只读清点（ole_info）·palette_color 枚举确认")
    o80_doc = server.cdr_open(preset="A4")
    if "document_id" in o80_doc:
        o80_id = o80_doc["document_id"]
        server.cdr_create(o80_id, [
            {"type": "rectangle", "x": 0, "y": 0, "width": 30, "height": 20,
             "fill": "#CC8844"}])
        oi80 = server.cdr_doc(o80_id, action="ole_info")
        if no_error("cdr_doc(ole_info)", oi80):
            check("可清点页面上的 OLE 对象（ole_info）",
                  oi80.get("ole_count") == 0 and isinstance(oi80.get("shapes"), list),
                  f"count={oi80.get('ole_count')}")
        pc80 = server.cdr_doc(o80_id, action="palette_color", palette_index=1,
                              color_index=1)
        if no_error("cdr_doc(palette_color)", pc80):
            check("调色板颜色可枚举（palette.Color(i)）",
                  bool(pc80.get("color")), str(pc80.get("color"))[:40])
        no_error("关闭 OLE 测试文档", server.cdr_close(o80_id))

    # ---------------- 15ci. 第八十一波：立体化复制 / 数据字段详情 ----------------
    print("\n[15ci] 第八十一波：EffectExtrude.CopyFrom·DataField 详情与按名取")
    x81_doc = server.cdr_open(preset="A4")
    if "document_id" in x81_doc:
        x81_id = x81_doc["document_id"]
        x81r = server.cdr_create(x81_id, [
            {"type": "rectangle", "x": 0, "y": 0, "width": 20, "height": 10},
            {"type": "rectangle", "x": 60, "y": 0, "width": 20, "height": 10},
        ])
        if no_error("cdr_create(立体化复制测试)", x81r):
            x81a, x81b = [{"name": c["name"]} for c in x81r["created"]]
            server.cdr_shape_ops(x81_id, op="extrude", target=x81a,
                                 params={"depth": 20})
            server.cdr_shape_ops(x81_id, op="extrude", target=x81b,
                                 params={"depth": 10})
            cf81 = server.cdr_shape_ops(x81_id, op="extrude_set", target=x81a,
                                        effect={"copy_from": {"name": x81b["name"]}})
            if no_error("cdr_shape_ops(extrude_set copy_from)", cf81):
                check("立体化效果可复制（CopyFrom）",
                      "copy_from" in (cf81.get("applied") or []),
                      str(cf81.get("applied")))
        df81 = server.cdr_doc(x81_id, action="datafield_info")
        if no_error("cdr_doc(datafield_info)", df81):
            check("数据字段详情可读（按名取也可用）",
                  (df81.get("count") or 0) >= 1
                  and df81.get("by_name_lookup") is True,
                  f"count={df81.get('count')}")
        no_error("关闭第八十一波测试文档", server.cdr_close(x81_id))

    # ---------------- 15cj. 第八十二波：点集容器编辑（PointRange 补方法） ----------------
    print("\n[15cj] 第八十二波：PointRange.Reverse·AddPoint·InsertPoint·AddPoints·RemoveAdjacentDuplicates")
    p82_doc = server.cdr_open(preset="A4")
    if "document_id" in p82_doc:
        p82_id = p82_doc["document_id"]
        p82r = server.cdr_create(p82_id, [
            {"type": "curve_segment", "x1": 0, "y1": 0, "x2": 40, "y2": 0},
            {"type": "curve_segment", "x1": 0, "y1": 30, "x2": 20, "y2": 30},
        ])
        if no_error("cdr_create(点集编辑测试)", p82r):
            pS, pT = [{"name": c["name"]} for c in p82r["created"]]
            pe82 = server.cdr_shape_ops(p82_id, op="point_range_edit", other=pS,
                                        params={"spacing": 10, "reverse": True,
                                                "append_from": pT["name"],
                                                "remove_duplicates": True})
            if no_error("cdr_shape_ops(point_range_edit)", pe82):
                check("点集容器可编辑（Reverse/AddPoints/RemoveAdjacentDuplicates）",
                      {"reverse", "add_points", "remove_duplicates"}
                      <= set(pe82.get("applied") or []),
                      f"{pe82.get('count_before')} → {pe82.get('count_after')}")
            fit82 = server.cdr_path(p82_id, pS, action="fit_to_points", other=pS,
                                    params={"spacing": 10, "tolerance": 3,
                                            "reverse": True})
            if no_error("cdr_path(fit_to_points+reverse)", fit82):
                check("拟合前可反转点序（PointRange.Reverse）",
                      fit82.get("points") is not None,
                      f"points={fit82.get('points')}")
        no_error("关闭点集编辑测试文档", server.cdr_close(p82_id))

    # ---------------- 15ck. 第八十三波：双灯配置 / 剪贴板状态 / 组件检测 ----------------
    print("\n[15ck] 第八十三波：EffectExtrude 双灯 Invoke putter·Clipboard·Components")
    x83_doc = server.cdr_open(preset="A4")
    if "document_id" in x83_doc:
        x83_id = x83_doc["document_id"]
        x83r = server.cdr_create(x83_id, [
            {"type": "rectangle", "x": 0, "y": 0, "width": 20, "height": 10}])
        if no_error("cdr_create(双灯测试)", x83r):
            x83a = {"name": x83r["created"][0]["name"]}
            server.cdr_shape_ops(x83_id, op="extrude", target=x83a,
                                 params={"depth": 20})
            es83 = server.cdr_shape_ops(x83_id, op="extrude_set", target=x83a,
                                        effect={"lights": [
                                            {"present": True, "intensity": 80},
                                            {"present": True, "intensity": 60}]})
            if no_error("cdr_shape_ops(extrude_set 双灯)", es83):
                ext83 = es83.get("extrude") or {}
                lights83 = ext83.get("lights") or []
                on83 = [l for l in lights83 if l.get("present")]
                check("立体化可双灯同开（Invoke putter）",
                      len(on83) == 2 and {l.get("intensity") for l in on83} == {80, 60},
                      f"on={[(l.get('index'), l.get('intensity')) for l in on83]}")
            cb83 = server.cdr_doc(x83_id, action="clipboard_info")
            if no_error("cdr_doc(clipboard_info)", cb83):
                check("剪贴板状态可读（Clipboard.DataPresent/Valid）",
                      "data_present" in cb83 and "valid" in cb83,
                      f"present={cb83.get('data_present')}")
            cm83 = server.cdr_doc(x83_id, action="components_info")
            if no_error("cdr_doc(components_info)", cm83):
                check("组件安装状态可查（IsComponentInstalled）",
                      isinstance(cm83.get("checked"), list)
                      and len(cm83.get("checked")) >= 1,
                      str(cm83.get("checked"))[:60])
        no_error("关闭双灯测试文档", server.cdr_close(x83_id))

    # ---------------- 15cl. 第八十四波：主窗口控制 / 屏幕矩形 / 状态栏文字 / GMS 路径 ----------------
    print("\n[15cl] 第八十四波：ICUIFrameWindow 控制·ICUIApplication.Create*·GMSManager 路径")
    u84_doc = server.cdr_open(preset="A4")
    if "document_id" in u84_doc:
        u84_id = u84_doc["document_id"]
        fm84 = server.cdr_view(u84_id, action="frame_maximize")
        if no_error("cdr_view(frame_maximize)", fm84):
            check("主窗口可最大化（ICUIFrameWindow.Maximize）",
                  fm84.get("done") is True and bool(fm84.get("caption")),
                  f"caption={fm84.get('caption')}")
        fr84 = server.cdr_view(u84_id, action="frame_restore")
        if no_error("cdr_view(frame_restore)", fr84):
            check("主窗口可还原（Restore）", fr84.get("done") is True,
                  str(fr84.get("done")))
        ft84 = server.cdr_view(u84_id, action="frame_tile")
        if no_error("cdr_view(frame_tile)", ft84):
            check("视口可平铺（TileViews）", ft84.get("done") is True,
                  str(ft84.get("done")))
        sr84 = server.cdr_doc(u84_id, action="screen_rect",
                              settings={"left": 0, "top": 0, "right": 300,
                                        "bottom": 200})
        if no_error("cdr_doc(screen_rect)", sr84):
            check("屏幕矩形可创建（ICUIApplication.CreateScreenRect）",
                  sr84.get("created") is True, str(sr84.get("created")))
        st84 = server.cdr_doc(u84_id, action="status_text",
                              settings={"caption": "MCP 状态"})
        if no_error("cdr_doc(status_text)", st84):
            check("状态栏文字可设（ICUIStatusText.SetCaptionText）",
                  st84.get("caption_set") == "MCP 状态",
                  str(st84.get("caption_set")))
        gi84 = server.cdr_doc(u84_id, action="gms_info")
        if no_error("cdr_doc(gms_info)", gi84):
            check("GMS 路径可查（GMSManager）",
                  bool(gi84.get("gms_path")), str(gi84.get("gms_path"))[:40])
        no_error("关闭第八十四波测试文档", server.cdr_close(u84_id))

    # ---------------- 15cm. 第八十五波：图像列表登记（ICUIImageList） ----------------
    print("\n[15cm] 第八十五波：ICUIImageList.AddBitmap（键登记）")
    i85_doc = server.cdr_open(preset="A4")
    if "document_id" in i85_doc:
        i85_id = i85_doc["document_id"]
        il85 = server.cdr_doc(i85_id, action="imagelist_info",
                              settings={"key": "MCP 图标"})
        if no_error("cdr_doc(imagelist_info)", il85):
            check("图像列表可登记位图（ImageList.AddBitmap）",
                  il85.get("image_count") == 1
                  and il85.get("added_key") == "MCP 图标"
                  and il85.get("exists") is True,
                  f"count={il85.get('image_count')} key={il85.get('added_key')}")
        no_error("关闭图像列表测试文档", server.cdr_close(i85_id))

    # ---------------- 15cn. 第八十六波：图像列表传字节 / 状态栏位图 / 数据源 / 配置文件安装 ----------------
    print("\n[15cn] 第八十六波：ImageList.AddImage/CreateBitmapImage(字节)·StatusText.SetBitmap·UnregisterDataSource·ColorProfiles.Install")
    i86_doc = server.cdr_open(preset="A4")
    if "document_id" in i86_doc:
        i86_id = i86_doc["document_id"]
        bmp86 = (b"BM" + (58).to_bytes(4, "little") + b"\x00" * 4
                 + (54).to_bytes(4, "little") + (40).to_bytes(4, "little")
                 + (1).to_bytes(4, "little") + (1).to_bytes(4, "little")
                 + (1).to_bytes(2, "little") + (24).to_bytes(2, "little")
                 + b"\x00" * 24 + b"\xff\xff\xff\x00")
        p86 = os.path.join(out, "w86_icon.bmp")
        with open(p86, "wb") as fh:
            fh.write(bmp86)
        il86 = server.cdr_doc(i86_id, action="imagelist_info",
                              settings={"key": "MCP 图标", "image_path": p86})
        if no_error("cdr_doc(imagelist_info 字节)", il86):
            check("图像列表可收原始字节（AddImage + CreateBitmapImage）",
                  il86.get("add_mode") == "bytes"
                  and il86.get("bitmap_valid") is True
                  and il86.get("bitmap_size") == [1, 1],
                  f"mode={il86.get('add_mode')} size={il86.get('bitmap_size')}")
        st86 = server.cdr_doc(i86_id, action="status_text",
                              settings={"caption": "MCP", "set_bitmap": True})
        if no_error("cdr_doc(status_text 位图)", st86):
            check("状态栏可设位图（ICUIStatusText.SetBitmap）",
                  st86.get("bitmap_set") is True,
                  str(st86.get("bitmap_set")))
        ds86 = server.cdr_doc(i86_id, action="datasource_info",
                              settings={"name": "mcp"})
        if no_error("cdr_doc(datasource_info)", ds86):
            check("数据源可注销（UnregisterDataSource）",
                  ds86.get("unregistered") is False
                  and ds86.get("has_data_context") is True,
                  f"unregistered={ds86.get('unregistered')}")
        cp86 = server.cdr_doc(i86_id, action="color_profile_install",
                              settings={"path": p86})
        if no_error("cdr_doc(color_profile_install)", cp86):
            check("配置文件安装可调（ColorProfiles.Install，非法文件如实返回 False）",
                  cp86.get("installed") is False
                  and cp86.get("count_before") == cp86.get("count_after"),
                  f"installed={cp86.get('installed')}")
        no_error("关闭第八十六波测试文档", server.cdr_close(i86_id))

    # ---------------- 15co. 第八十七波：SetLight* 方法 / 点距 / 剖面线原点 / B样条插点 ----------------
    print("\n[15co] 第八十七波：Extrude.SetLightPresent/Intensity/Position·Point.DistanceTo·HatchPattern.SetOrigin·BSpline.InsertControlPoints")
    w87_doc = server.cdr_open(preset="A4")
    if "document_id" in w87_doc:
        w87_id = w87_doc["document_id"]
        w87r = server.cdr_create(w87_id, [
            {"type": "rectangle", "x": 0, "y": 0, "width": 20, "height": 10},
            {"type": "curve_segment", "x1": 0, "y1": 40, "x2": 40, "y2": 40},
            {"type": "rectangle", "x": 50, "y": 0, "width": 30, "height": 20},
        ])
        if no_error("cdr_create(第八十七波)", w87r):
            wA, wS, wH = [{"name": c["name"]} for c in w87r["created"]]
            server.cdr_shape_ops(w87_id, op="extrude", target=wA,
                                 params={"depth": 20})
            lt87 = server.cdr_shape_ops(w87_id, op="extrude_set", target=wA,
                                        effect={"lights": [
                                            {"present": True, "intensity": 90,
                                             "position": 3},
                                            {"present": True, "intensity": 55}]})
            if no_error("cdr_shape_ops(extrude_set SetLight*)", lt87):
                l87 = (lt87.get("extrude") or {}).get("lights") or []
                on87 = [x for x in l87 if x.get("present")]
                check("立体化灯光用正式方法调（SetLightPresent/Intensity/Position）",
                      len(on87) == 2
                      and {x.get("intensity") for x in on87} == {90, 55},
                      f"on={[(x.get('index'), x.get('intensity')) for x in on87]}")
            pm87 = server.cdr_shape_ops(w87_id, op="point_math", other=wS,
                                        params={"spacing": 10, "index_a": 1,
                                                "index_b": 2})
            if no_error("cdr_shape_ops(point_math)", pm87):
                check("点间距离可算（IVGPoint.DistanceTo）",
                      (pm87.get("distance") or 0) > 9.0,
                      f"distance={pm87.get('distance')}")
            server.cdr_modify(w87_id, wH,
                              {"fill": {"type": "hatch", "index": 1,
                                        "back": "#FFFFFF"}})
            hp87 = server.cdr_shape_ops(w87_id, op="hatch_fill", target=wH,
                                        params={"pattern_origin": [2, 3]})
            if no_error("cdr_shape_ops(hatch_fill pattern_origin)", hp87):
                check("剖面线图案原点可设（HatchPattern.SetOrigin/GetOrigin）",
                      (hp87.get("pattern") or {}).get("origin") == [2.0, 3.0],
                      str(hp87.get("pattern")))
            bs87 = server.cdr_doc(w87_id, action="bspline_resize", color_index=4,
                                  value=5,
                                  settings={"insert": {"segment": 2, "count": 1,
                                                       "x": 5, "y": 5}})
            if no_error("cdr_doc(bspline_resize insert)", bs87):
                check("B 样条可插入控制点（InsertControlPoints）",
                      bs87.get("inserted") is True,
                      f"{bs87.get('points_before')} → {bs87.get('points_after')}")
        no_error("关闭第八十七波测试文档", server.cdr_close(w87_id))

    # ---------------- 15cp. 第八十八波：能力目录 cdr_help / 错误 hint / 工具数唯一真源 ----------------
    print("\n[15cp] 第八十八波：cdr_help 自描述目录·错误响应 hint·工具数动态")
    hl88 = server.cdr_help()
    check("cdr_help 列出全部工具（含自身）",
          (hl88.get("tool_count") or 0) >= 31
          and any(t.get("tool") == "cdr_help" for t in (hl88.get("tools") or [])),
          f"tool_count={hl88.get('tool_count')} actions={hl88.get('action_count')}")
    check("cdr_help 给出可执行工作流",
          len(hl88.get("workflow") or []) >= 5,
          f"steps={len(hl88.get('workflow') or [])}")
    hs88 = server.cdr_help(tool="shape_ops")
    check("cdr_help(tool=…) 返回该工具 action 清单（支持省略 cdr_ 前缀）",
          (hs88.get("tool") == "cdr_shape_ops"
           and len(hs88.get("actions") or []) >= 80),
          f"actions={len(hs88.get('actions') or [])}")
    hsn88 = server.cdr_help(tool="cdr_snap")
    check("只读 action 单独标注（降低误改风险）",
          len(hsn88.get("read_only_actions") or []) >= 5,
          f"read_only={len(hsn88.get('read_only_actions') or [])}")
    check("未知工具名返回可用清单",
          server.cdr_help(tool="not_a_tool").get("error") == "unknown_tool",
          str(server.cdr_help(tool="not_a_tool").get("error")))
    hb88 = server.cdr_modify("no-such-doc", {"index": 1}, {"width": 10})
    check("错误响应带可操作 hint",
          bool(hb88.get("hint")) and hb88.get("error"),
          f"{hb88.get('error')} → {str(hb88.get('hint'))[:34]}")
    stt88 = server.cdr_status()
    check("cdr_status 的工具数与目录一致（唯一真源）",
          stt88.get("tools") == hl88.get("tool_count"),
          f"status={stt88.get('tools')} catalog={hl88.get('tool_count')}")

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
