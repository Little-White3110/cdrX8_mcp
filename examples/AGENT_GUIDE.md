# 让 Agent 稳定操作 CorelDRAW 的建议

把这段放进 system prompt（或项目的 AGENTS.md），能显著减少无效调用。

---

## 工具使用规则

1. **先看再改。** 任何修改前必须先 `cdr_inspect`，把返回的 `revision` 和 `snapshot_token`
   原样传给修改类工具。收到 `stale_view` 错误时，重新 `cdr_inspect` 再试，不要重试旧参数。
2. **一次一个语义单元。** 画一个版式别拆成 20 次调用——`cdr_create` 支持批量，
   一批 = 一个 undo 步。
3. **坐标是毫米、页面左下角、Y 向上。** 矩形给「左下角 + 宽高」，椭圆给「圆心 + 半径」。
4. **印前用 CMYK。** 屏幕稿用 `#RRGGBB`，印刷件用 `{"c":..,"m":..,"y":..,"k":..}`。
5. **对象用名字定位。** 创建时给 `name`；按 `{"index": n}` 会因增删而漂移。
6. **发排前转曲。** 用 `cdr_text(to_curves=true)` 或 `cdr_shape_ops(op="to_curves")`。
7. **导出不要覆盖。** 默认 `overwrite=false`；只有用户明确同意才传 `true`。
8. **做不到就说做不到。** 立体化/透镜/虚线描边/箭头/表格/字形/刀刻/橡皮没有可用接口，
   直接告诉用户要手动做，不要反复换参数试。

## 按意图选工具

| 我想… | 用 |
|---|---|
| 画矩形/椭圆/星形/螺旋/直线/贝塞尔 | `cdr_create` |
| 加标题文字、改字体字号 | `cdr_create(type="text")` / `cdr_text` |
| 换配色、上渐变 | `cdr_modify(changes={"fill":{...}})` |
| 加投影、轮廓图、封套 | `cdr_modify(changes={"drop_shadow"/"contour"/"envelope"})` |
| 虚线、加箭头 | `cdr_modify(changes={"outline":{"dash":5,"start_arrow":2,"end_arrow":3}})` |
| 文字沿曲线 | `cdr_text(target=文字, fit_to_path={"name":"曲线对象"})` |
| 段落文字溢出 | `cdr_text(target=文字, fit_frame=true)` |
| 两个形状焊接/修剪/相交 | `cdr_shape_ops(op="weld"/"trim"/"intersect", target, other)` |
| 组合/解组 | `cdr_select` → `cdr_shape_ops(op="group"/"combine")`；解组用 `op="ungroup"` + `target` |
| 圆角/倒角/扇贝 | `cdr_shape_ops(op="fillet"/"chamfer"/"scallop", radius=...)` |
| 把图片放进形状 | `cdr_shape_ops(op="powerclip", target=内容, other=容器)` |
| 排整齐 | `cdr_align(horizontal/vertical, to="page"/"selection")` |
| 换客户名/改年份 | `cdr_replace_text(find, replace)` |
| 分图层管理 | `cdr_layers` |
| 批量出图 | 循环：`cdr_text` + `cdr_replace_text` + `cdr_export` |
| 导出印刷 PDF + 预览图 | `cdr_export(format="pdf")` + `cdr_export(format="png", dpi=150)` |
| 把素材图放进来 | `cdr_import(path)` |

## 推荐工作流

```
1. cdr_status / cdr_connect            → 确认环境
2. cdr_open(width_mm=..., height_mm=...)  或  cdr_open(preset="A4")  或  cdr_open(path=...)
3. cdr_inspect                          → 记录 revision + snapshot_token
4. cdr_create(items=[...])              → 批量搭建
5. cdr_inspect → cdr_modify / cdr_text / cdr_shape_ops / cdr_align  → 精修
6. cdr_export(format="pdf") + cdr_export(format="png", dpi=150)     → 交付 + 预览
7. cdr_save(path="...")                 → 存档
```

## 出问题时的排查顺序

| 现象 | 先查什么 |
|---|---|
| `not_connected` | `cdr_status` 看 `running`；用 `cdr_connect(launch=true)` |
| `stale_view` / `revision_mismatch` | 重新 `cdr_inspect`，用新 token |
| `path_denied` | 目标目录不在 `CDR_ALLOWED_ROOTS` 内 |
| `file_exists` | 换输出路径，或确认后传 `overwrite=true` |
| `bad_argument: 找不到名为 … 的对象` | 对象可能被组合了（组合体 type_id=7），先用 `cdr_find` 看现状 |
| `timeout` | CorelDRAW 可能弹了模态框（缺字体/色彩配置）；**不要自动重试写操作** |
| 导出返回 `warnings` 提到 DPI | 空页面不支持指定 DPI，已按默认分辨率导出 |

## 一个完整的例子

> 帮我出一张 90×54mm 的名片：白底、左上角一个 20mm 的红黑色块、
> 右侧放名字「张三」24pt 和电话「138-0000-0000」10pt，导出 PDF 和 PNG 到
> `D:\design\out\`，再存一份 cdr。

Agent 应当：
```
cdr_open(90, 54) → cdr_inspect
→ cdr_create([矩形(带圆角), 文字, 文字])        # 一个 undo 步
→ cdr_align(vertical="middle", to="page")       # 对齐
→ cdr_text(target=名字, size=24)                # 微调
→ cdr_export(pdf) → cdr_export(png, dpi=150) → cdr_save
```
用户不满意可以一句话 `cdr_undo` 回滚。
