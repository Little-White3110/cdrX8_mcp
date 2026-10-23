# 让 Agent 稳定操作 CorelDRAW 的建议

把这段放进 system prompt（或项目的 AGENTS.md），能显著减少无效调用。

---

## 工具使用规则

1. **先看再改。** 任何修改前必须先 `cdr_inspect`，把返回的 `revision` 和 `snapshot_token`
   原样传给修改类工具。收到 `stale_view` 错误时，重新 `cdr_inspect` 再试一次，不要重试旧参数。
2. **一次一个语义单元。** 画一个版式就别拆成 20 次工具调用——`cdr_create` 支持批量，
   一批 = 一个 undo 步。
3. **坐标是毫米、页面左下角、Y 向上。** 矩形给「左下角 + 宽高」，椭圆给「圆心 + 半径」。
   不要假设 Y 向下。
4. **印前用 CMYK。** 屏幕稿可以用 `#RRGGBB`，但印刷件请用 `{"c":..,"m":..,"y":..,"k":..}`。
5. **对象用名字定位。** 创建时给 `name`，之后按 `{"name": "..."}` 定位；
   按 `{"index": n}` 会因增删而漂移。
6. **导出不要覆盖。** 默认 `overwrite=false`；只有用户明确同意时才传 `true`。
7. **做不到就说做不到。** 刀刻、橡皮、网格填充、智能绘图、表格、字形没有编程接口，
   直接告诉用户需要手动操作，不要反复尝试其他工具。

## 推荐工作流

```
1. cdr_status / cdr_connect          → 确认环境
2. cdr_open(width_mm=..., height_mm=...)  或  cdr_open(path="...")
3. cdr_inspect                        → 记录 revision + snapshot_token
4. cdr_create(items=[...])            → 批量搭建
5. cdr_inspect → cdr_modify / cdr_replace_text → 精修
6. cdr_export(format="pdf") + cdr_export(format="png")  → 交付 + 预览
7. cdr_save(path="...")               → 存档
```

## 出问题时的排查顺序

| 现象 | 先查什么 |
|---|---|
| `not_connected` | `cdr_status` 看 `running`；用 `cdr_connect(launch=true)` |
| `stale_view` / `revision_mismatch` | 重新 `cdr_inspect`，用新 token |
| `path_denied` | 目标目录不在 `CDR_ALLOWED_ROOTS` 内 |
| `file_exists` | 换输出路径，或确认后传 `overwrite=true` |
| `timeout` | CorelDRAW 可能弹了模态框（缺字体/色彩配置）；**不要自动重试写操作** |
| `com_error` | 看 `hresult`；常见是对象已被删除或文档被关闭 |

## 一个完整的例子

> 帮我出一张 90×54mm 的名片：白色底、左上角一个 20mm 的红黑双色色块、
> 右侧放名字「张三」24pt 和电话「138-0000-0000」10pt，导出 PDF 和 PNG 到
> `D:\design\out\`，再存一份 cdr。

Agent 应当：`cdr_open(90, 54)` → `cdr_inspect` → `cdr_create([矩形, 文字, 文字])`
→ `cdr_export(pdf)` → `cdr_export(png)` → `cdr_save`。**五次创建合成一个 undo 步**，
用户不满意可以一句话回滚。
