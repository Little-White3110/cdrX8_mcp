#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""覆盖率审计：类型库方法全集 vs 本仓库实际调用的 API 方法。

做法：
  1. 用 pywin32 gencache 从 vgcoreauto.tlb 展开出全部自动化类与方法
     （临时目录重定向到 <repo>/tmp/gen_py，不会污染系统临时目录，跑完自动删除）
  2. AST 扫描 src/ 与 scripts/，收集 `.方法名(` 与 setattr/getattr 的属性名
  3. 求交集，按接口输出覆盖率

用法：
    python scripts/coverage_audit.py
    python scripts/coverage_audit.py --show IVGCurve     # 看某接口的未用方法
"""

from __future__ import annotations

import argparse
import ast
import os
import re
import shutil
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
GEN_ROOT = os.path.join(ROOT, "tmp", "gen_py")

TLB = r"C:\Program Files\Corel\CorelDRAW Graphics Suite X8\Programs64\vgcoreauto.tlb"

MAIN_IFACES = [
    "IVGApplication", "IVGDocument", "IVGPage", "IVGLayer", "IVGShape",
    "IVGShapeRange", "IVGFill", "IVGOutline", "IVGCurve", "IVGSubPath",
    "IVGNode", "IVGText", "IVGTextRange", "IVGColor", "IVGBitmap",
    "IVGFountainFill", "IVGTransparency", "IVGStyle", "IVGWindow", "IVGColorContext",
]

INFRA = {"QueryInterface", "AddRef", "Release", "GetTypeInfoCount", "GetTypeInfo",
         "GetIDsOfNames", "Invoke", "__iter__", "__str__", "__int__", "__len__",
         "__bool__", "__call__"}

ATTR_CALL = re.compile(r"\.([A-Za-z_][A-Za-z0-9_]*)\s*\(")
SETATTR = re.compile(r'setattr\(\s*[A-Za-z_][A-Za-z0-9_.]*\s*,\s*["\']([A-Za-z_][A-Za-z0-9_]*)["\']')
GETATTR = re.compile(r'getattr\(\s*[A-Za-z_][A-Za-z0-9_.]*\s*,\s*["\']([A-Za-z_][A-Za-z0-9_]*)["\']')


def load_api() -> dict[str, set[str]]:
    import pythoncom  # noqa: F401  (仅 Windows)
    import win32com.client.gencache as gencache

    os.makedirs(GEN_ROOT, exist_ok=True)
    os.environ["TEMP"] = GEN_ROOT
    os.environ["TMP"] = GEN_ROOT
    tl = pythoncom.LoadTypeLib(TLB)
    attr = tl.GetLibAttr()
    gencache.is_readonly = False
    mod = gencache.EnsureModule(str(attr[0]), attr[1], attr[3], attr[4])
    if mod is None:
        raise RuntimeError("EnsureModule 失败")
    src = open(mod.__file__, encoding="utf-8", errors="replace").read()
    api: dict[str, set[str]] = {}
    for node in ast.parse(src).body:
        if isinstance(node, ast.ClassDef):
            names = {it.name for it in node.body if isinstance(it, ast.FunctionDef)} - INFRA
            if names:
                api[node.name] = names
    return api


def used_names() -> set[str]:
    used: set[str] = set()
    for base in ("src", "scripts"):
        for root, _dirs, files in os.walk(os.path.join(ROOT, base)):
            for fn in files:
                if not fn.endswith(".py"):
                    continue
                text = open(os.path.join(root, fn), encoding="utf-8").read()
                used |= set(ATTR_CALL.findall(text))
                used |= set(SETATTR.findall(text))
                used |= set(GETATTR.findall(text))
    return used


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--show", help="列出某个接口未使用的方法")
    args = ap.parse_args()

    try:
        api = load_api()
    except Exception as exc:  # noqa: BLE001
        print(f"读取类型库失败（需要 Windows + 已安装 CorelDRAW X8 + pywin32）：{exc}")
        return 2

    used = used_names()
    total = sum(len(v) for v in api.values())
    hits = sum(len(v & used) for v in api.values())

    print(f"自动化类 {len(api)}  方法 {total}  已调用 {hits}  总覆盖 {hits / total * 100:.1f}%\n")
    print(f"{'接口':<20}{'方法数':>7}{'已调用':>7}{'覆盖':>8}")
    print("-" * 44)
    main_total = main_hit = 0
    rows = []
    for iface in MAIN_IFACES:
        methods = api.get(iface, set())
        hit = methods & used
        rows.append((iface, len(methods), len(hit), sorted(methods - used)))
        main_total += len(methods)
        main_hit += len(hit)
    rows.sort(key=lambda r: -r[2])
    for iface, t, h, _u in rows:
        pct = f"{h / t * 100:.0f}%" if t else "-"
        print(f"{iface:<20}{t:>7}{h:>7}{pct:>8}")
    print("-" * 44)
    print(f"{'主干合计':<20}{main_total:>7}{main_hit:>7}"
          f"{main_hit / main_total * 100:>7.1f}%")

    if args.show:
        for iface, _t, _h, unused in rows:
            if iface.lower() == args.show.lower():
                print(f"\n{iface} 未使用（{len(unused)}）：")
                print("  " + ", ".join(unused))
                break

    shutil.rmtree(GEN_ROOT, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
