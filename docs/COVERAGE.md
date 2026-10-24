# API 覆盖清单

> 目的：证明"还有多少没探"。每一行要么有对应的工具，要么有实测失败证据。
> 环境：CorelDRAW Graphics Suite X8 `18.0.0.450`（64 位）/ Python 3.11.9 64 位
> 最近更新：2026-10-25 第五波（节点级编辑打通 + 路径/排列/文本读取扩展）

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
| `IVGDocument` | 78 | 已用：Unit / Pages / AddPages / DeletePages / ActivePage / ActiveLayer / Selection / Save·SaveAs / ExportEx / ExportBitmap(不适用) / PublishToPDF / Undo / Redo / Begin·EndCommandGroup / FileName / Dirty / Name / StyleSheet(只读) / **ToUnits / FromUnits / SampleColorAtPoint / SampleColorInArea / ShowAllHiddenObjects / SetLayout / Repeat** |
| `IVGPage` | 22 | **全覆盖**：SetSize / SelectSize / GetSize / CreateLayer / Activate / Delete / MoveTo / FindShape(s) / TextFind / TextReplace / GetGuides / UnlockAllShapes / GetObjectsBoundingBox(改用自算) / SelectShapesFromRectangle / Shapes / Layers / Orientation |
| `IVGLayer` | 50 | 已用：Create*（矩形/椭圆/多边形/螺旋/网格/线/贝塞尔/艺术字/段落字/参考线/**角度参考线**）/ Import / ImportEx / CreateCurve / Activate / Delete / MoveAbove / Visible / Editable / Printable / Master / FindShape |
| `IVGShape` | 124 | 已用：几何（Move/SetPosition/SetBoundingBox/Rotate/Flip/Skew/TransformMatrix/**ClearTransformations**/**SetRotationCenter**）/ Fill / Outline / Transparency / Locked / Name / Type / GetBoundingBox / Duplicate / Clone / Delete / **层级（OrderToFront·Back / OrderFrontOf / OrderBackOf / OrderForwardOne / OrderBackOne / OrderReverse）** / **StepAndRepeat / Distribute / CreateBoundary / CreatePerspective / Separate** / ConvertToCurves / ConvertToSymbol / Weld·Trim·Intersect / AddToPowerClip / CreateDropShadow·Contour·Envelope / AlignToPage·AlignToShape / Fillet·Chamfer·Scallop / Group / Ungroup / BreakApart / Combine / **GetOverprintFillState / GetOverprintOutlineState（已实测，未封装）** |
| `IVGShapeRange` | 131 | 已用：Group / Combine / AlignAndDistribute / ApplyFountainFill / SetOutlineProperties(不适用，改属性直写) / ConvertToCurves / Add / Count |
| `IVGFill` | 13 | 已用：ApplyUniformFill / ApplyFountainFill / ApplyTextureFill / ApplyNoFill；Pattern·PostScript·Hatch 需外部文件或库 |
| `IVGOutline` | 9 | 已用：Width / Color / Style(赋对象) / StartArrow·EndArrow(赋对象) / LineCaps / LineJoin / BehindFill / ScaleWithShape / SetNoOutline |
| `IVGCurve` | 37 | 已用：Length / Area / SubPaths / Contour / CreateSubPath / GetCopy / **Nodes / Segments / Selection / FindNodeAtPoint / FindSegmentAtPoint / ReverseDirection / JoinTouchingSubpaths / RemoveOverlaps / GetPolyline / IsOnCurve / IsClockwise / Closed / GetCurveInfo**；数组入参（AppendSubpathFromPoints / CreateSubPathFromArray）不可用 |
| `IVGSubPath` | 39 | 已用：CreateSubPath / AppendLineSegment / AppendCurveSegment2 / Closed / **Nodes / Segments / StartNode / EndNode / GetPointAt / GetPointPositionAt / GetTangentVectorAt / GetPerpendicularVectorAt / GetCurveSpeedAt / GetIntersections / ReverseDirection / Next / Previous / Extract / GetCopy**；求交/等分/切向已打通，点集数组仍取不回 |
| `IVGNode` | 19 | **已覆盖（第五波纠正）**：`curve.Nodes` / `curve.SubPaths.Item(i).Nodes` 是**属性**；`GetPosition/SetPosition/Move/Type/Index/SubPathIndex/IsEnding/Selected/Delete/CreateSelection` 均可用；范围操作走 `curve.Nodes.Range(a,b)` → `IVGNodeRange`（Fillet/Chamfer/Scallop/Smoothen/SetType/Move/Rotate/Skew/Stretch/BreakApart/GetBoundingBox）；`app.CreateNodeRange()` 返回空集合 |
| `IVGText` | 30 | 已用：Story / GetSpaceProperties / SetSpaceProperties / GetAlignProperties / SetAlignProperties / FitToPath / FitTextToFrame / **GetContents / Range / ExportToFile** |
| `IVGTextRange` | 35 | 已用：Text / Font / Size / Bold / Italic / Underline / ChangeCase / ApplyBulletEffect / ApplyDropCapEffect / ApplyNoEffect / InsertBefore / InsertAfter |
| `IVGColor` | 44 | 通过 Create*Color 间接使用（RGB/CMYK/Gray/Lab/HLS/Registration） |
| `IVGBitmap` | 16 | 已用：ConvertTo / Resample / Crop / ResetCropEnvelope / Mode / ResolutionX / Size；**Trace 禁用**（见下） |
| `IVGFountainFill` | 7 | 通过 ApplyFountainFill 设置，未做二次编辑 |
| `IVGTransparency` | 9 | 已用：ApplyUniformTransparency / ApplyFountainTransparency / ApplyNoTransparency |
| `IVGStyle` / `IVGStyleSheet` | 16 / 16 | **未覆盖**：CreateStyleFromShape 需要有效类别常量，未打通 |
| `IVGWindow` | 9 | 已用：Refresh / Activate / DocumentToScreen / ScreenToDocument / DocumentDistanceToScreenDistance / **ScreenDistanceToDocumentDistance**；缩放/新建窗口未封装 |
| `IVGColorContext` | 7 | 未封装（色彩管理配置，属交付环节而非创作环节） |

---

## 二、能力域 → 工具（26 个）

| 能力域 | 工具 | 状态 |
|---|---|---|
| 会话/文档/状态 | `cdr_status` `cdr_connect` `cdr_documents` `cdr_open` `cdr_inspect` `cdr_save` `cdr_close` | ✅ 46 项联调覆盖 |
| 绘图 | `cdr_create`（矩形含圆角 / 椭圆含圆弧扇形 / 多边形 / 星形 / 螺旋 / 网格 / 直线 / 贝塞尔 / 文字 / 段落文字 / 参考线 / **任意折线（第五波修正重复节点）**） | ✅ |
| 样式 | `cdr_modify`（纯色/渐变/纹理/无色、描边含虚线箭头、投影、轮廓图、封套、透明） | ✅ |
| 变换 | `cdr_modify`（移动/尺寸/旋转/翻转/斜切/矩阵/克隆/层级/**层级相对调整**/**清除变换**/**旋转中心**/复制/删除/锁定） | ✅ |
| **节点/线段** | `cdr_node`（列节点/列线段/移动/选中/删除/圆角/倒角/scallop/平滑/拉伸/旋转/斜切/断开/加节点/断线/控制点/按坐标查找） | ✅ 第五波 |
| 文字排版 | `cdr_text`（字体字号粗斜体下划线 / 对齐 / 大小写 / 项目符号 / 首字下沉 / 行距字距 / 前后插入 / 沿路径 / 自适应 / 转曲 / **读取全文·区间·导出纯文本**） | ✅ |
| 布尔与整形 | `cdr_shape_ops`（weld/trim/intersect/combine/group/ungroup/break_apart/to_curves/to_symbol/powerclip/fillet/chamfer/scallop/**step_repeat/distribute/boundary/perspective/separate**） | ✅ |
| 对齐 | `cdr_align`（左中右/上下中/等间距，对齐对象或页面）+ `cdr_modify(align_to_shape/align_to_page)` | ✅ |
| 选择与查找 | `cdr_select` `cdr_find` | ✅ |
| 图层 | `cdr_layers`（增删改/显隐/锁定/打印/移动） | ✅ |
| 页面 | `cdr_pages`（列表/增删/切页/自定义尺寸/纸张预设/参考线/**角度参考线**/解锁/包围盒） | ✅ |
| 位图 | `cdr_bitmap`（信息/转色彩模式/重采样/裁剪/取消裁剪） | ✅（Trace 除外） |
| 路径 | `cdr_path`（长度面积/偏移轮廓/追加子路径/折线点/**反转方向·焊子路径·去重叠·参数取点·切/法向量·速度·交点·找节点/线段**） | ✅ |
| 导入导出 | `cdr_import` `cdr_export`（png/jpg/svg/pdf，可指定 DPI，可只导选区） | ✅ |
| 文档 | `cdr_doc`（概况/清选区/清撤销栈/调色板/箭头/**单位换算·取色·显示隐藏·版式·重复操作**） | ✅ 第五波 |
| 视图 | `cdr_view`（刷新/激活/文档↔屏幕坐标/**屏幕↔文档距离**） | ✅ 第五波 |
| 历史 | `cdr_undo`（撤销/重做） | ✅ |
| 全局替换 | `cdr_replace_text`（原生 Page.TextReplace） | ✅ |

---

## 三、明确不可用（附实测证据）

| 能力 | API | 实测结果 | 影响 |
|---|---|---|---|
| **位图自动转矢量** | `Bitmap.Trace` | CorelDRAW 进程变为**无响应**（`Responding=False`），必须强制结束进程 | **严禁暴露**：COM 无法中断，会把 Agent 一起拖死 |
| **文本导入** | `Text.ImportFromFile` | 调用后 COM 通道被模态框卡住（实测 60 s 超时，进程仍在跑） | 只做读取与导出；反向导入不提供 |
| **点集直接建曲线** | `Curve.AppendSubpathFromPoints/FitToPoints`、`Curve.CreateSubPathFromArray`、`Document.CreateCurveFitToPoints/CreateBSpline` | 传 list / `VARIANT(VT_ARRAY\|VT_R8)` 均报 `TypeError: Python instance can not be converted to a COM object`；`CreateSubPathFromArray` 甚至让客户端进程崩溃 | 用逐点 `AppendLineSegment`（`cdr_create(type="polyline")`）代替 |
| **混合** | `Shape.CreateBlend` | 第二个形状参数同样 TypeError | 暂不支持混合效果 |
| **文字区间查询** | `TextRange.FindRanges(Query)` | `Identifier 'Wave' is not a method of 'global'`（Query 是查询表达式，不是普通字符串） | 只提供 `Text.Range(start,end)` 直接取区间 |
| 立体化 / 透镜 / 尺寸标注 / 连接线 / 符号插入 | `CreateExtrude` `CreateLens` `CreateLinearDimension` `CreateConnector` `CreateSymbol` | `TypeError: Python instance can not be converted to a COM object`（5 组参数组合均失败） | 需要 pywin32 无法构造的结构体/枚举参数 |
| 描边整体设置 | `Outline.SetProperties(Ex)` | 同样 TypeError | 已改用逐属性赋值绕过 |
| 专色 | `CreateSpotColor` / `CreatePaletteColor(0,n)` | `未找到调色板 0` | 需要本机调色板 ID |
| 图案填充 | `Fill.ApplyPatternFill` | `无法装载到图案文件` | 需要 .pat 文件 |
| 对象样式 | `StyleSheet.CreateStyleFromShape` | `未找到类别 0~8`（9 个取值全试过） | 类别常量不是普通整数，未找到可用取值 |
| 按名套用样式 | `Shape.ApplyStyle` | `未找到样式 xxx` | 依赖上一条，样式建不出来就无从应用 |
| 符号库 | `app.SymbolLibraries` | 对象存在，但 `Count = 0`（本机未装符号库） | `CreateSymbol` 需要库内符号名，无库可用（`ConvertToSymbol` 可用，已封装为 `to_symbol`） |
| 剪贴板 | `Document.Copy/Paste` | `AttributeError` | Document 上没有这两个方法（但 `Layer.Paste` 可用，见第四） |
| 吸附设置 | `Document.SnapToGrid/SnapToObjects` | `AttributeError` | 同上 |
| PowerClip 提取 | `Shape.PowerClip` | 对象存在但无 `Extract` | 只能"放入"不能"取出" |
| 页面背景 | `Page.Background` | 不存在 | — |
| 文档元数据 | `Document.Keywords/Notes` | 不存在 | — |

> 已修正的历史结论：
> * 折线取样 `Curve.GetPolyline` **可用**（返回 `IVGCurve`，遍历 `SubPaths/Nodes` 即可），
>   `cdr_path(action="polyline")` 已能返回采样点。
> * 节点级编辑 **可用**（`curve.Nodes` / `curve.Segments` 是属性），已封装为 `cdr_node`。

---

## 四、已知但未封装（低频）

这些**能用**，只是 Agent 日常用不上，暂不占工具位：

* `IVGWindow` 的缩放/新建窗口（`Zoom`、`NewWindow`）：截图/预览才需要
* `IVGColorContext`：色彩配置、配置文件合并（交付环节）
* `IVGFountainFill.GetTransformations`：读渐变变换矩阵（渐变微调已能改角度/平移/矩阵）
* `IVGTraceSettings`：PowerTRACE 参数对象（Trace 本身不安全，不封装）
* `IVGStyle` 的 `SetProperty/GetProperty/GetAllPropertyNames`：能拿到属性名，但属性 ID 语义未解，暂不封装
* `Shape.CopyPropertiesFrom(Source, Properties)`：实测 `Properties` 是 1~15 的位掩码，但每位含义未逆向出来
* `Curve.GetCurveInfo`：返回 `com_struct(PositionX, PositionY, ElementType, NodeType, Flags)` 数组，语义未完全确认
* `Layer.Paste` / `PasteEx` / `PasteSpecial`：实测 `Paste` 可用（依赖系统剪贴板内容），暂不暴露
* `Application.OpenDocumentAsCopy` / `CreateDocumentFromTemplate`：需要模板/文件路径，未纳入工具面
* `Application.CreateHSBColor/CreateYIQColor/CreateCMYColor/CreateBWColor`：实测可用，但颜色运算已走纯 Python
* 符号库插入：需要符号库名称，`CreateSymbol` 本身也未打通；`Page.FindClosestSnapPoint` 实测报
  「服务器出现意外情况」，不封装
* `Bitmap.ResolveLink/UpdateLink`：需要外部链接位图才有意义，未做联调

---

## 五、复现方式

```powershell
python -m pytest -q              # 55 项纯逻辑
python scripts/live_smoke.py     # 111 项端到端（真实 CorelDRAW）
```

联调断言覆盖了上表所有 ✅ 项。反解脚本属于过程产物，已按要求放在 `tmp/` 并在完成时清理；
如需重建，按本文件"怎么得出这份清单"三步即可复现。

---

## 六、量化覆盖率（2026-10-25 第五波实测）

方法：反解类型库拿到 196 个自动化类 / **1619** 个方法（已剔除 `QueryInterface` 等 COM 基础设施）；
用 AST 扫描 `src/cdrx8_mcp/*.py`，收集所有 `.方法名(` 调用与 `setattr/getattr` 的属性名，
与主干接口的方法名求交集。**注意：属性名（如 `curve.Nodes`）也会被计入**，所以本波数字含真实属性访问。

**全局：565 / 1619 = 34.9%**（第一轮 21.2% → 第二轮 27.6% → 第三轮 28.9% → 第四轮 29.6% → 本轮 34.9%）
**主干 20 接口：333 / 743 = 44.8%**

复现命令：`python scripts/coverage_audit.py`（会临时展开类型库到 `tmp/gen_py`，跑完自动清理）

| 接口 | 方法数 | 已调用 | 覆盖 |
|---|---:|---:|---:|
| `IVGShape` | 123 | 63 | 51% |
| `IVGShapeRange` | 125 | 63 | 50% |
| `IVGSubPath` | 38 | 29 | **76%** |
| `IVGDocument` | 74 | 30 | 41% |
| `IVGPage` | 21 | 18 | **86%** |
| `IVGCurve` | 36 | 22 | 61% |
| `IVGLayer` | 49 | 17 | 35% |
| `IVGTextRange` | 31 | 14 | 45% |
| `IVGApplication` | 78 | 13 | 17% |
| `IVGText` | 29 | 14 | 48% |
| `IVGNode` | 18 | 9 | 50%（本波起为真实调用） |
| `IVGBitmap` | 15 | 7 | 47% |
| `IVGFill` | 12 | 6 | 50% |
| `IVGWindow` | 8 | 7 | **88%** |
| `IVGFountainFill` | 6 | 5 | 83% |
| `IVGColor` | 43 | 7 | 16%（色彩运算已改纯 Python，见下） |
| `IVGOutline` | 8 | 3 | 38% |
| `IVGTransparency` | 8 | 3 | 38% |
| `IVGStyle` | 15 | 2 | 13% |
| `IVGColorContext` | 6 | 1 | 17% |
| **合计** | **743** | **333** | **44.8%** |

> **怎么读这两个数**：
> 1. 全局 34.9% 覆盖 1619 个方法（含变体与重复）；主干 20 接口 44.8%。
> 2. 它是**方法/属性名匹配**的近似值，同名方法会跨接口误计（例如 `Fillet`/`Chamfer`/`Scallop`
>    同时存在于 Shape、ShapeRange、Node、NodeRange，数字会一起涨）。
> 3. `IVGColor` 的 16% 是**故意的**：实测 `IVGColor` 不暴露任何分量属性（`Red`/`Cyan`/`Gray`
>    全部 AttributeError，`ConvertToCMYK` 只改类型不返回值），所以色彩运算迁到纯 Python
>    （`src/cdrx8_mcp/colors.py`），COM 只负责把颜色**应用**到对象；本波新增的
>    `ToString` / `GetName` 用于取色采样输出。
> 4. 未调用的方法里很大一部分是变体：`CreateArrowHead`×5、`CreateBitmap`×3、
>    `ConvertTo*`×10、大量 `Wide`/`Ex`/`2`/`3` 后缀；另有事件订阅（需要 COM 连接点）
>    与 CorelScriptTools/MathUtils（与绘图无关）。

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

## 七之三、第五波已接入（2026-10-25，节点/路径/排列/读取）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **节点级编辑**（此前误判为不可用） | `cdr_node`：list/segments/segment_info/find/move/select/delete/type/fillet/chamfer/scallop/smoothen/stretch/rotate/skew/break_apart/add_node/break_segment/control_points | ✅ 4 节点折线列点/移动/圆角（4→6 节点）/平滑/加节点/断线全部生效；节点坐标随 `doc.Unit` 变化 |
| 节点范围操作 | `curve.Nodes.Range(a,b)` → `IVGNodeRange`（Fillet/Chamfer/Scallop/Smoothen/Move/Rotate/Skew/Stretch/BreakApart/GetBoundingBox） | ✅ |
| 线段级几何 | `curve.Segments.Item(i)`：type/length/start/end/GetPointAt/GetCurvatureAt/GetCurveSpeedAt | ✅ 起点/终点/中点/曲率/速度均可读 |
| 贝塞尔控制点读写 | `Segment.GetStartingControlPointPosition` / `SetStartingControlPointPosition`（同理 Ending） | ✅ 40,140 → 40,150 已写回 |
| 按坐标查找节点/线段 | `Curve.FindNodeAtPoint` / `FindSegmentAtPoint`（返回 segment + paramOffset） | ✅ |
| 折线采样（修正旧结论） | `Curve.GetPolyline` 返回 `IVGCurve`，遍历 `SubPaths/Nodes` | ✅ 4 点折线取回 4 个角点 |
| 路径方向与清理 | `cdr_path(action="reverse_direction"/"join_subpaths"/"remove_overlaps")` | ✅ `IsClockwise` 翻转，子路径焊回 1 条 |
| 路径参数查询 | `cdr_path(action="point_position"/"tangent_vector"/"perpendicular_vector"/"curve_speed"/"intersections"/"find_node"/"find_segment")` | ✅ 切向量 (0,30) / 法向量 (-30,0) 与角度一致 |
| 层级相对调整 | `cdr_modify(changes={"order":{"mode":"front_of","ref":{...}}})`，另有 back_of/forward_one/back_one/reverse | ✅ |
| 清除变换 / 旋转中心 | `cdr_modify(changes={"clear_transformations":true,"rotation_center":{x,y}})` | ✅ |
| 阵列复制 | `cdr_shape_ops(op="step_repeat", copies,distance_x,distance_y,mode_x,direction_x,mode_y,direction_y)` | ✅ 5→7 对象；`direction_y` 只能 2/3 |
| 原生分布 | `cdr_shape_ops(op="distribute", distribute_type=1..6)` | ✅ 六种类型均可调用 |
| 形状边界 | `cdr_shape_ops(op="boundary", x,y)` → `Shape.CreateBoundary` | ✅ 新建 1 个边界对象 |
| 透视 / 拆效果 | `cdr_shape_ops(op="perspective"/"separate")` | ✅ `CreatePerspective` 返回 `IVGEffect` |
| 文本读取 / 导出 | `cdr_text(read_contents=true / range_start+range_end / export_path)` | ✅ 读回全文与区间；导出文件内容精确为 5 字符 |
| 单位换算 | `cdr_doc(action="units", value, from_unit, to_unit)` → `Document.ToUnits/FromUnits` | ✅ 1 in → 25.4 mm |
| 取色 | `cdr_doc(action="sample_color", x,y[,x2,y2])` → `SampleColorAtPoint/InArea` + `Color.ToString/GetName` | ✅ 返回 `RGB255,...` 与中文色名 |
| 显示隐藏对象 / 版式 / 重复操作 | `cdr_doc(action="show_all"/"layout"/"repeat")` | ✅ |
| 屏幕距离 → 文档距离 | `cdr_view(action="distance_to_document")` | ✅ |
| 角度参考线 | `cdr_pages(action="guide_angle", x, y, angle)` → `Layer.CreateGuideAngle` | ✅ |
| 折线创建修正 | `cdr_create(type="polyline")` 改用首段线段作载体 | ✅ 4 点折线只有 4 个节点（旧实现会留 3 个重复节点 + 退化子路径） |

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

* **能力域覆盖**：18 个能力域 / 26 个工具，日常出图闭环已完整，并补上节点级编辑。
* **方法级覆盖**：全局 34.9%（565/1619），主干 20 接口 44.8%（333/743）。
* **明确不可用**：位图描摹、文本导入、点集数组建曲线、混合（CreateBlend）、
  立体化/透镜/尺寸/连接线、对象样式、符号库、专色、图案填充。
* **本轮修正的历史结论**：节点级编辑可用（`curve.Nodes/.Segments` 是属性）；
  `GetPolyline` 可用（返回 `IVGCurve`）；折线创建不再产生重复节点。


