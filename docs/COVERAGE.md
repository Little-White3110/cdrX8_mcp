# API 覆盖清单

> 目的：证明"还有多少没探"。每一行要么有对应的工具，要么有实测失败证据。
> 环境：CorelDRAW Graphics Suite X8 `18.0.0.450`（64 位）/ Python 3.11.9 64 位

## 怎么得出这份清单

1. **反解类型库**（只读 `vgcoreauto.tlb`）→ 692 个类型 / 6789 个方法声明；
   用 pywin32 gencache 展开成可调用签名 → 256 个自动化类 / 2248 个 dispatch 方法。
2. **按接口族逐项实测**：能跑通的写成工具，跑不通的记下异常原文。
3. **每条结论都有出处**：要么是 `scripts/live_smoke.py` 里的断言，要么是本文件"不可用"小节里的错误信息。

---

## 一、主干接口覆盖

| 接口 | 方法数 | 覆盖情况 |
|---|---:|---|
| `IVGApplication` | 82 | 已用：CreateDocument / OpenDocument / 全套 Create*Color / CreateShapeRange / CreateStructExportOptions / CreateStructImportOptions / OutlineStyles / ArrowHeads / Version / Path / Documents |
| `IVGDocument` | 78 | 已用：Unit / Pages / AddPages / DeletePages / ActivePage / ActiveLayer / Selection / Save·SaveAs / ExportEx / ExportBitmap(不适用) / PublishToPDF / Undo / Redo / Begin·EndCommandGroup / FileName / Dirty / Name / StyleSheet(只读) |
| `IVGPage` | 22 | **全覆盖**：SetSize / SelectSize / GetSize / CreateLayer / Activate / Delete / MoveTo / FindShape(s) / TextFind / TextReplace / GetGuides / UnlockAllShapes / GetObjectsBoundingBox(改用自算) / SelectShapesFromRectangle / Shapes / Layers / Orientation |
| `IVGLayer` | 50 | 已用：Create*（矩形/椭圆/多边形/螺旋/网格/线/贝塞尔/艺术字/段落字/参考线）/ Import / ImportEx / CreateCurve / Activate / Delete / MoveAbove / Visible / Editable / Printable / Master / FindShape |
| `IVGShape` | 124 | 已用：几何（Move/SetPosition/SetBoundingBox/Rotate/Flip/Skew/TransformMatrix）/ Fill / Outline / Transparency / Locked / Name / Type / GetBoundingBox / Duplicate / Clone / Delete / OrderToFront·Back / ConvertToCurves / ConvertToSymbol / Weld·Trim·Intersect / AddToPowerClip / CreateDropShadow·Contour·Envelope / AlignToPage·AlignToShape / Fillet·Chamfer·Scallop / Group / Ungroup / BreakApart / Combine |
| `IVGShapeRange` | 131 | 已用：Group / Combine / AlignAndDistribute / ApplyFountainFill / SetOutlineProperties(不适用，改属性直写) / ConvertToCurves / Add / Count |
| `IVGFill` | 13 | 已用：ApplyUniformFill / ApplyFountainFill / ApplyTextureFill / ApplyNoFill；Pattern·PostScript·Hatch 需外部文件或库 |
| `IVGOutline` | 9 | 已用：Width / Color / Style(赋对象) / StartArrow·EndArrow(赋对象) / LineCaps / LineJoin / BehindFill / ScaleWithShape / SetNoOutline |
| `IVGCurve` | 37 | 已用：Length / Area / SubPaths / Contour / CreateSubPath / GetCopy |
| `IVGSubPath` | 39 | 已用：CreateSubPath / AppendLineSegment / Closed（用于 `cdr_create(type="polyline")` 逐点构建任意多边形）；求交/等分等未用 |
| `IVGNode` | 19 | **未覆盖**：类型库中没有任何 `.Nodes` 访问器，拿不到节点集合 |
| `IVGText` | 30 | 已用：Story / GetSpaceProperties / SetSpaceProperties / GetAlignProperties / SetAlignProperties / FitToPath / FitTextToFrame |
| `IVGTextRange` | 35 | 已用：Text / Font / Size / Bold / Italic / Underline / ChangeCase / ApplyBulletEffect / ApplyDropCapEffect / ApplyNoEffect / InsertBefore / InsertAfter |
| `IVGColor` | 44 | 通过 Create*Color 间接使用（RGB/CMYK/Gray/Lab/HLS/Registration） |
| `IVGBitmap` | 16 | 已用：ConvertTo / Resample / Crop / ResetCropEnvelope / Mode / ResolutionX / Size；**Trace 禁用**（见下） |
| `IVGFountainFill` | 7 | 通过 ApplyFountainFill 设置，未做二次编辑 |
| `IVGTransparency` | 9 | 已用：ApplyUniformTransparency / ApplyFountainTransparency / ApplyNoTransparency |
| `IVGStyle` / `IVGStyleSheet` | 16 / 16 | **未覆盖**：CreateStyleFromShape 需要有效类别常量，未打通 |
| `IVGWindow` | 9 | 未封装（缩放/坐标换算，Agent 场景用不到） |
| `IVGColorContext` | 7 | 未封装（色彩管理配置，属交付环节而非创作环节） |

---

## 二、能力域 → 工具（22 个）

| 能力域 | 工具 | 状态 |
|---|---|---|
| 会话/文档/状态 | `cdr_status` `cdr_connect` `cdr_documents` `cdr_open` `cdr_inspect` `cdr_save` `cdr_close` | ✅ 46 项联调覆盖 |
| 绘图 | `cdr_create`（矩形含圆角 / 椭圆含圆弧扇形 / 多边形 / 星形 / 螺旋 / 网格 / 直线 / 贝塞尔 / 文字 / 段落文字 / 参考线） | ✅ |
| 样式 | `cdr_modify`（纯色/渐变/纹理/无色、描边含虚线箭头、投影、轮廓图、封套、透明） | ✅ |
| 变换 | `cdr_modify`（移动/尺寸/旋转/翻转/斜切/矩阵/克隆/层级/复制/删除/锁定） | ✅ |
| 文字排版 | `cdr_text`（字体字号粗斜体下划线 / 对齐 / 大小写 / 项目符号 / 首字下沉 / 行距字距 / 前后插入 / 沿路径 / 自适应 / 转曲） | ✅ |
| 布尔与整形 | `cdr_shape_ops`（weld/trim/intersect/combine/group/ungroup/break_apart/to_curves/to_symbol/powerclip/fillet/chamfer/scallop） | ✅ |
| 对齐 | `cdr_align`（左中右/上下中/等间距，对齐对象或页面）+ `cdr_modify(align_to_shape/align_to_page)` | ✅ |
| 选择与查找 | `cdr_select` `cdr_find` | ✅ |
| 图层 | `cdr_layers`（增删改/显隐/锁定/打印/移动） | ✅ |
| 页面 | `cdr_pages`（列表/增删/切页/自定义尺寸/纸张预设/参考线/解锁/包围盒） | ✅ |
| 位图 | `cdr_bitmap`（信息/转色彩模式/重采样/裁剪/取消裁剪） | ✅（Trace 除外） |
| 路径 | `cdr_path`（长度面积/偏移轮廓/追加子路径/折线点） | ✅ |
| 导入导出 | `cdr_import` `cdr_export`（png/jpg/svg/pdf，可指定 DPI，可只导选区） | ✅ |
| 历史 | `cdr_undo`（撤销/重做） | ✅ |
| 全局替换 | `cdr_replace_text`（原生 Page.TextReplace） | ✅ |

---

## 三、明确不可用（附实测证据）

| 能力 | API | 实测结果 | 影响 |
|---|---|---|---|
| **位图自动转矢量** | `Bitmap.Trace` | CorelDRAW 进程变为**无响应**（`Responding=False`），必须强制结束进程 | **严禁暴露**：COM 无法中断，会把 Agent 一起拖死 |
| **节点级编辑** | `IVGNode` | 类型库全库搜索 `.Nodes` **无任何命中**，没有访问器 | 无法用代码改锚点/控制柄 |
| 立体化 / 透镜 / 尺寸标注 / 连接线 / 符号插入 | `CreateExtrude` `CreateLens` `CreateLinearDimension` `CreateConnector` `CreateSymbol` | `TypeError: Python instance can not be converted to a COM object`（5 组参数组合均失败） | 需要 pywin32 无法构造的结构体/枚举参数 |
| 描边整体设置 | `Outline.SetProperties(Ex)` | 同样 TypeError | 已改用逐属性赋值绕过 |
| 专色 | `CreateSpotColor` | `未找到调色板 0` | 需要本机调色板 ID |
| 图案填充 | `Fill.ApplyPatternFill` | `无法装载到图案文件` | 需要 .pat 文件 |
| 对象样式 | `StyleSheet.CreateStyleFromShape` | `未找到类别 0~8`（9 个取值全试过） | 类别常量不是普通整数，未找到可用取值 |
| 按名套用样式 | `Shape.ApplyStyle` | `未找到样式 xxx` | 依赖上一条，样式建不出来就无从应用 |
| 符号库 | `app.SymbolLibraries` | 对象存在，但 `Count = 0`（本机未装符号库） | `CreateSymbol` 需要库内符号名，无库可用（`ConvertToSymbol` 可用，已封装为 `to_symbol`） |
| 剪贴板 | `Document.Copy/Paste` | `AttributeError` | Document 上没有这两个方法 |
| 吸附设置 | `Document.SnapToGrid/SnapToObjects` | `AttributeError` | 同上 |
| PowerClip 提取 | `Shape.PowerClip` | 对象存在但无 `Extract` | 只能"放入"不能"取出" |
| 页面背景 | `Page.Background` | 不存在 | — |
| 文档元数据 | `Document.Keywords/Notes` | 不存在 | — |
| 折线取样 | `Curve.GetPolyline` | 返回的变体数组无 `Count`，pywin32 取不回来 | `cdr_path(action="polyline")` 会返回空并说明 |

---

## 四、已知但未封装（低频）

这些**能用**，只是 Agent 日常用不上，暂不占工具位：

* `IVGWindow`：缩放、文档↔屏幕坐标换算（做截图/预览才需要）
* `IVGColorContext`：色彩配置、配置文件合并（交付环节）
* `IVGFountainFill`：对已有渐变做二次编辑（目前已能创建渐变，改参数需重设）
* `IVGTraceSettings`：PowerTRACE 参数对象（Trace 本身不安全，不封装）
* `IVGSubPath` 的求交/等分/切向：需要先解决"点集取不回"的问题
* 符号库插入：需要符号库名称，`CreateSymbol` 本身也未打通

---

## 五、复现方式

```powershell
python -m pytest -q              # 43 项纯逻辑
python scripts/live_smoke.py     # 46 项端到端（真实 CorelDRAW）
```

联调断言覆盖了上表所有 ✅ 项。反解脚本属于过程产物，已按要求放在 `tmp/` 并在完成时清理；
如需重建，按本文件"怎么得出这份清单"三步即可复现。

---

## 六、量化覆盖率（2026-10-24 第二轮实测）

方法：反解类型库拿到 196 个自动化类 / **1619** 个方法（已剔除 `QueryInterface` 等 COM 基础设施）；
用 AST 扫描 `src/cdrx8_mcp/*.py`，收集所有 `.方法名(` 调用与 `setattr/getattr` 的属性名，
与主干接口的方法名求交集。

**全局：480 / 1619 = 29.6%**（第一轮 21.2% → 第二轮 27.6% → 第三轮 28.9% → 本轮 29.6%）
**主干 20 接口：275 / 743 = 37.0%**

复现命令：`python scripts/coverage_audit.py`（会临时展开类型库到 `tmp/gen_py`，跑完自动清理）

| 接口 | 方法数 | 已调用 | 覆盖 |
|---|---:|---:|---:|
| `IVGShape` | 123 | 49 | 40% |
| `IVGShapeRange` | 125 | 49 | 39% |
| `IVGSubPath` | 38 | 22 | **58%** |
| `IVGDocument` | 74 | 22 | 30% |
| `IVGPage` | 21 | 18 | **86%** |
| `IVGCurve` | 36 | 17 | 47% |
| `IVGLayer` | 49 | 16 | 33% |
| `IVGTextRange` | 31 | 13 | 42% |
| `IVGApplication` | 78 | 13 | 17% |
| `IVGText` | 29 | 11 | 38% |
| `IVGNode` ⚠️ | 18 | 8 | 44%（同名误计，实际无节点入口） |
| `IVGBitmap` | 15 | 7 | 47% |
| `IVGFill` | 12 | 6 | **50%** |
| `IVGWindow` | 8 | 6 | **75%** |
| `IVGFountainFill` | 6 | 5 | 83% |
| `IVGColor` | 43 | 5 | 12%（色彩运算已改纯 Python，见下） |
| `IVGOutline` | 8 | 3 | 38% |
| `IVGTransparency` | 8 | 3 | 38% |
| `IVGStyle` | 15 | 1 | 7% |
| `IVGColorContext` | 6 | 1 | 17% |
| **合计** | **743** | **275** | **37.0%** |

> **怎么读这两个数**：
> 1. 全局 29.6% 覆盖 1619 个方法（含变体与重复）；主干 20 接口 37.0%。
> 2. 它是**方法名匹配**的近似值，同名方法会误计（`IVGNode` 44% 是假的——我们拿不到节点集合）。
> 3. `IVGColor` 的 12% 是**故意的**：实测 `IVGColor` 不暴露任何分量属性（`Red`/`Cyan`/`Gray`
>    全部 AttributeError，`ConvertToCMYK` 只改类型不返回值），所以色彩运算迁到纯 Python
>    （`src/cdrx8_mcp/colors.py`），COM 只负责把颜色**应用**到对象。
> 4. 未调用的方法里很大一部分是变体：`CreateArrowHead`×5、`CreateBitmap`×3、
>    `ConvertTo*`×10、大量 `Wide`/`Ex`/`2`/`3` 后缀。

---

## 七、第二轮已接入的能力（原先列在"待做"）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| 描边转曲 | `cdr_shape_ops(op="outline_to_curves")` | ✅ 对象原地 矩形→曲线 |
| 渐变二次编辑 | `cdr_modify(changes={"gradient":{angle,edge_pad,translate,transformations,opaque}})` | ✅ |
| 曲线级画圆/椭圆/矩形/三点弧 | `cdr_path(action="append_circle/append_ellipse/append_rectangle/append_arc")` | ✅ |
| 追加贝塞尔段 | `cdr_path(action="append_bezier")` | ✅ |
| 路径等分 | `cdr_path(action="equal_divide")` | ✅ |
| 节点精简 / 自焊接 | `cdr_path(action="reduce_nodes"/"self_weld")` | ✅ |
| 几何判定 | `cdr_path(action="point_inside"/"intersects"/"get_point"/"get_tangent"/"subpath_info")` | ✅ |
| 颜色解析/转换/混合 | `cdr_color`（纯 Python，不依赖 COM） | ✅ |
| 位图色彩滤镜 | `cdr_modify(changes={"effect":{"type":"bci"/"gamma"/"hsl"/"invert"/"posterize"/"color_balance"}})` | ✅ |
| 艺术字 ↔ 段落字互转 | `cdr_text(to_paragraph=true / to_artistic=true)` | ✅ |
| 段落文本内容读写与替换 | `cdr_text(contents=… / replace_old+replace_new=… / find=…)` | ✅ 仅段落文本 |
| 文本区间操作 | `cdr_text(select_all / align_to_baseline / copy_attributes_from)` | ✅ |
| 按坐标点查询/选择对象 | `cdr_pages(action="shape_at"/"select_at")` | ✅ |
| 页面几何与文本查找 | `cdr_pages(action="geometry"/"text_find"/"move_to")` | ✅ |
| 文档调色板登记 | `cdr_doc(action="palette")` | ✅ |
| 自定义箭头 | `cdr_doc(action="arrowhead")` | ✅ |
| 清撤销栈 / 清选区 | `cdr_doc(action="clear_undo"/"clear_selection")` | ✅ |
| 另存副本（不换当前路径） | `cdr_save(as_copy=true)` | ✅（注意：X8 仍会改文档名） |
| **批量样式** | `cdr_modify(target={"all":true}/{"type_id":n})` | ✅ 21 个对象一批改完，0 失败 |
| 复制填充 / 描边 | `cdr_modify(changes={"fill_from"/"outline_from"})` | ✅ |
| 仿射变换 / 对齐到页面中心·点·网格 | `cdr_modify(changes={"affine":[...], "align_to":{mode}})` | ✅ |
| 视图刷新与坐标换算 | `cdr_view(action="refresh"/"to_screen"/"to_document"/"distance_to_screen")` | ✅ |
| 路径查询 | `cdr_path(action="curvature"/"perpendicular"/"evenly_spaced"/"closest_segment"/"is_on_path"/"segment_at"/"is_rect_on_edge")` | ✅ |

## 七之二、仍未接入（第三轮候选）

| 能力 | 接口 | 为什么还没做 |
|---|---|---|
| 从内存图建位图 | `Layer.CreateBitmap/CreateBitmap2` | 需要 ImageData 对象，未找到构造入口 |
| 位图转黑白/双色调/调色板 | `Bitmap.ConvertToBW/Duotone/Paletted` | 需要 RenderType / Duotone / Palette 选项对象 |
| 自定义描边图案 | `Fill.ApplyPostscriptFill/ApplyHatchFill` | 需要 PostScript 索引 / Hatch 库 |
| 色彩管理 | `IVGColorContext`（6 个） | 交付环节，非创作环节 |
| 对象样式 | `IVGStyle`（15 个） | 类别常量未找到有效取值（0~8 全失败） |
| 打印 | `Document.PrintOut()` | **故意不接**：会真的发出打印任务 |
| 整页插入 / 跨文档合并 | `Document.InsertPages` | 传 Page 对象报「找不到成员」，未找到正确形态 |
| 事件订阅 | `IVGApplicationEvents` / `IVGDocumentEvents` | 需要 COM 连接点，pywin32 支持有限 |
| CorelScript 工具函数 | `ICorelScriptTools`（45 个：单位换算、三角、日期、文件夹遍历） | 大部分与 CorelDRAW 无关；单位换算已在 `mm_units` 内部处理 |
| 数学工具 | `IVGMathUtils`（19 个） | 同上 |

> 说明：`ICorelScriptTools` 名字像"脚本工具"，实测是 VBA 的工具函数集合（`FromInches` /
> `FromPoints` / `ASin` / `BuildDate` / `FindFirstFolder`…），**不是宏执行器**。
> 想通过它执行 VBA 宏这条路走不通。

---

## 八、一句话总结

* **能力域覆盖**：15 个能力域 / 22 个工具，日常出图闭环已完整。
* **方法级覆盖**：24.1%（179/743），剩余多为变体与重复，也有上面 18 项真正有价值的能力。
* **明确不可用**：位图描摹、节点编辑、立体化/透镜/尺寸/连接线、对象样式、符号库、专色。


