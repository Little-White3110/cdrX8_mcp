# API 覆盖清单

> 目的：证明"还有多少没探"。每一行要么有对应的工具，要么有实测失败证据。
> 环境：CorelDRAW Graphics Suite X8 `18.0.0.450`（64 位）/ Python 3.11.9 64 位
> 最近更新：2026-10-25 第十九波（用户级应用偏好 KV / 当前工具信息）

## 怎么得出这份清单

1. **反解类型库**（只读 `vgcoreauto.tlb`）→ 692 个类型 / 6789 个方法声明；
   用 pywin32 gencache 展开成可调用签名 → 256 个自动化类 / 2248 个 dispatch 方法。
2. **按接口族逐项实测**：能跑通的写成工具，跑不通的记下异常原文。
3. **每条结论都有出处**：要么是 `scripts/live_smoke.py` 里的断言，要么是本文件"不可用"小节里的错误信息。

---

## 一、主干接口覆盖

| 接口 | 方法数 | 覆盖情况 |
|---|---:|---|
| `IVGApplication` | 82 | 已用：CreateDocument / **CreateDocumentFromTemplate** / OpenDocument / 全套 Create*Color / CreateShapeRange / CreateStructExportOptions / CreateStructImportOptions / OutlineStyles（含 **AddStyle/Save**）/ ArrowHeads / Version / Path / Documents / **ColorManager / ActiveTool / ActiveToolStateGuid / RegisterUserApplicationPreference / GetApplicationPreferenceValue / SetApplicationPreferenceValue** / **FontList / Palettes（含 OpenFixed / Create / Item） / CreateDuotone / CreateSpotColorByName / CreatePatternCanvas / PatternCanvases / GetSupportedOpenTypeFeatures / CheckPlatformVersion / IsUILanguageAvailable**；OpenDocumentAsCopy 参数 TypeError |
| `IVGDocument` | 78 | 已用：Unit / Pages / AddPages / DeletePages / ActivePage / ActiveLayer / Selection / Save·SaveAs / ExportEx / ExportBitmap(不适用) / PublishToPDF / Undo / Redo / Begin·EndCommandGroup / FileName / Dirty / Name / StyleSheet / **Metadata / DataFields / ColorContext / ConvertToColorContext / Revert** / **设置：EditAcrossLayers / ApplyToDuplicate / CurvePrecision / Resolution / ReferencePoint / ShapeEnumDirection / TextFormatter / WorldScale / FacingPages / FirstPageOnRightSide / Layout / Unit / DrawingOriginX·Y / PreserveSelection** / **Keywords·Notes（可写，读需走 Metadata）** / **ToUnits / FromUnits / SampleColorAtPoint / SampleColorInArea / ShowAllHiddenObjects / SetLayout / Repeat / CreateImage / CreateColorStyles / InsertPages / HatchLibraries / SaveStyleSheet / LoadStyleSheet / CreateFill / CreateOutline（字符串格式无效）** |
| `IVGPage` | 22 | **全覆盖**：SetSize / SelectSize / GetSize / CreateLayer / Activate / Delete / MoveTo / FindShape(s) / TextFind / TextReplace / GetGuides / UnlockAllShapes / GetObjectsBoundingBox(改用自算) / SelectShapesFromRectangle / Shapes / Layers / Orientation / **FindClosestSnapPoint** |
| `IVGLayer` | 50 | 已用：Create*（矩形/椭圆/多边形/螺旋/网格/线/贝塞尔/艺术字/段落字/参考线/**角度参考线**/**位图 CreateBitmap·CreateBitmap2**）/ Import / ImportEx / CreateCurve / Activate / Delete / MoveAbove / **MoveBelow / GetAbove / GetBelow / Paste** / Visible / Editable / Printable / Master / FindShape |
| `IVGShape` | 124 | 已用：几何 / Fill / Outline / Transparency / Locked / Name / Type / GetBoundingBox / Duplicate / Clone / Delete / **层级 / StepAndRepeat / Distribute / CreateBoundary / CreatePerspective / Separate / CreateEnvelopeFromCurve / CreateEnvelopeFromShape / CompareTo / CreateSelection / RemoveFromSelection / IsOnShape / Copy / Cut / CopyToLayer / MoveToLayer / Project / Unproject / CreateDocumentFrom / CreateExtrude / CreateLens / CreateBlend** / ConvertToCurves / ConvertToSymbol / Weld·Trim·Intersect / AddToPowerClip / CreateDropShadow·Contour·Envelope / AlignToPage·AlignToShape / Fillet·Chamfer·Scallop / Group / Ungroup / BreakApart / Combine |
| `IVGShapeRange` | 131 | 已用：Group / Combine / AlignAndDistribute / ApplyFountainFill / SetOutlineProperties(不适用，改属性直写) / ConvertToCurves / Add / Count / **CreateParallelCurves / ConvertToBitmap / ConvertOutlineToObject / CreateBoundary / SetFillMode / SetPixelAlignedRendering / ApplyFill / ApplyOutline / CreateDocumentFrom** |
| `IVGFill` | 13 | 已用：ApplyUniformFill / ApplyFountainFill / ApplyTextureFill / **ApplyPatternFill（内存 PatternCanvas 可用；文件型 Pattern 仍需图像文件）** / ApplyPostscriptFill / ApplyNoFill / **GetCopy / CompareWith**；Hatch 需本机 Hatch 库（本机为空） |
| `IVGOutline` | 9 | 已用：Width / Color / **Style（可赋自定义线型）** / StartArrow·EndArrow(赋对象) / LineCaps / LineJoin / BehindFill / ScaleWithShape / SetNoOutline / **GetCopy / CompareWith** |
| `IVGOutlineStyle` / `IVGOutlineStyles` | 6 / 6 | **新打通（第十一波）**：`CreateOutlineStyle(dash,gap[,dash,gap,…])` 生成虚拟线型（Index=-1），`OutlineStyles.AddStyle(virtual)` 注册后得到可用 Index，赋给 `Outline.Style` 即生效；`Save()` 可持久化 |
| `IVGDataFields` / `IVGDataField` / `IVGDataItems` | 8 / 18 / 4 | **新打通（第十一波）**：`DataFields.Add(Name,"0")`、`IsPresent/Item`、`Shape.ObjectData.Add(DataField, Value)`、`DataItem.Value/FormattedValue` 可读写（对象级持久化数据字段） |
| `IVGCurve` | 37 | 已用：Length / Area / SubPaths / Contour / CreateSubPath / GetCopy / **Nodes / Segments / Selection / FindNodeAtPoint / FindSegmentAtPoint / ReverseDirection / JoinTouchingSubpaths / RemoveOverlaps / GetPolyline / IsOnCurve / IsClockwise / Closed / GetCurveInfo**；数组入参（AppendSubpathFromPoints / CreateSubPathFromArray）不可用；CreateCurveMappedToStroke 类型不匹配 |
| `IVGSubPath` | 39 | 已用：CreateSubPath / AppendLineSegment / AppendCurveSegment2 / Closed / **Nodes / Segments / StartNode / EndNode / GetPointAt / GetPointPositionAt / GetTangentVectorAt / GetPerpendicularVectorAt / GetCurveSpeedAt / GetIntersections / ReverseDirection / Next / Previous / Extract / GetCopy**；求交/等分/切向已打通，点集数组仍取不回 |
| `IVGNode` | 19 | **已覆盖（第五波纠正）**：`curve.Nodes` / `curve.SubPaths.Item(i).Nodes` 是**属性**；`GetPosition/SetPosition/Move/Type/Index/SubPathIndex/IsEnding/Selected/Delete/CreateSelection` 均可用；范围操作走 `curve.Nodes.Range(a,b)` → `IVGNodeRange`（Fillet/Chamfer/Scallop/Smoothen/SetType/Move/Rotate/Skew/Stretch/BreakApart/GetBoundingBox）；`app.CreateNodeRange()` 返回空集合 |
| `IVGText` | 30 | 已用：Story（**Font / Size / Bold / Italic / Underline 读取**）/ GetSpaceProperties / SetSpaceProperties / GetAlignProperties / SetAlignProperties / FitToPath / FitTextToFrame / **GetContents / Range / ExportToFile / GetHyphenationSettings**；GetFontProperties 结构体返回默认值不可信 |
| `IVGTextRange` | 35 | 已用：Text / Font / Size / Bold / Italic / Underline / ChangeCase / ApplyBulletEffect / ApplyDropCapEffect / ApplyNoEffect / InsertBefore / InsertAfter / **SetOpenTypeFeature / GetOpenTypeFeature / Delete / Copy / Paste / Straighten / FindRanges / EnumRanges**；ApplyStyle 调用不报错但未观测到效果 |
| `IVGColor` | 44 | 通过 Create*Color 间接使用（RGB/CMYK/Gray/Lab/HLS/Registration）；**ToString / GetName / CopyAppearance / StringAssign / IsSame / BlendWith / GetColorDistanceFrom** 均已实测 |
| `IVGBitmap` | 16 | 已用：ConvertTo / ConvertToBW / **ConvertToPaletted（0~9） / ConvertToDuotone** / Resample / Crop / ResetCropEnvelope / Inflate / Mode / ResolutionX / Size / Image / Duotone / **SetImageData（Alpha 必须传 None） / ApplyBitmapEffect（命令串）**；Trace 禁用（见下） |
| `IVGFountainFill` | 7 | 通过 ApplyFountainFill 设置，未做二次编辑 |
| `IVGTransparency` | 9 | 已用：ApplyUniformTransparency / ApplyFountainTransparency / **ApplyPatternTransparency（PatternCanvas 可用）** / ApplyNoTransparency / **Freeze / Unfreeze / Frozen** |
| `IVGStyle` / `IVGStyleSheet` | 16 / 16 | **已覆盖（第六/七/九波）**：`CreateStyleFromShape`、`CreateStyle(category, based_on, name)`、`Styles.Find/Item/Count`、`Style.Name/CategoryName/DisplayName/BasedOn`、`Rename / Delete / GetAllPropertyNames / GetPropertyAsString / SetPropertyAsString / SetProperty / DerivedStyles / ToString`、`Shape.ApplyStyle("样式名")`；**`Document.SaveStyleSheet/LoadStyleSheet` 可导出/恢复 .cdst**；**`Document.CreateColorStyles(harmony>0)` 会生成颜色样式（AllColorStyles）**；`fill`/`outline` 分组名不可直接读写 |
| `IVGWindow` | 9 | 已用：Refresh / Activate / DocumentToScreen / ScreenToDocument / DocumentDistanceToScreenDistance / **ScreenDistanceToDocumentDistance**；缩放/新建窗口未封装 |
| `IVGColors`（颜色样式集合） | 8 | **新打通（第九波）**：`StyleSheet.AllColorStyles` 的 `Count / Item(i) / GetName / ToString`；配合 `Document.CreateColorStyles(harmonies=1..3)` 生成颜色协调 |
| `IVGMetadata` | 21 | **新打通（第十波）**：`Document.Metadata` 的 Title / Author / LastAuthor / Subject / Keywords / Notes / Copyright / Revision / DocLanguage / DocID；前六项可写；Localizable* 是 `IVGLocalizableString` 对象，未纳入读取 |
| `IVGEffects` / `IVGEffect` | 12 / 12 | **新打通（第十五波）**：`Shape.Effects.Count/Item(i)`、`ExtrudeEffect/LensEffect/DropShadowEffect/BlendEffects` 等；`Effect.Type/Extrude/Lens/Separate/Clear` 可用 |
| `IVGEffectExtrude` | 18 | **新打通（第十五波）**：Type/Depth/AngleX·Y·Z/Shading/BaseColor/ShadingColor/UseBevel/BevelDepth/BevelAngle/ShowBevelOnly/UseExtrudeColorForBevel/UseFullColorRange/VanishingPoint 可读写；`Rotate/SetBevel/SetLight` 未封装 |
| `IVGEffectLens` | 18 | **新打通（第十五波）**：Type/Rate/Magnification/Color/FillColor/OutlineColor/FromColor/ToColor/UseFillColor/UseOutlineColor/PaletteRotation/ColorMapPalette/RemoveFace/ViewPointX·Y/UseViewPoint/Frozen 可读写；`Freeze/Unfreeze/Ungroup` 未封装 |
| `IVGEffectBlend` | 31 | **新打通（第十七波）**：Steps / Spacing / Angle / ColorBlendType / Mode / Loop / RotateShapes / MapNodes / FuseStart / FuseEnd / ColorAcceleration / SpacingAcceleration / LinkAcceleration / AccelerateSize / StartPoint / EndPoint / StartShape / EndShape / Path 可读写；`Split()` 可用 |
| `IVGArrowHeads` | 7 | **新打通（第十二波）**：`Application.ArrowHeads.Add(自定义箭头)` 把 `CreateArrowHead(curve)` 的结果登记进全局箭头库（89→90），可赋给 `Outline.StartArrow/EndArrow` |
| `IVGColorContext` | 7 | **已读（第十一波）**：`Document.ColorContext` 的 RenderingIntent / BlendingColorModel / ColorProfiles.Count / ColorProfileNameList；**第十八波**：`profile.CreateColorContext()` + `Document.ConvertToColorContext()` 可做文档颜色转换；直接 `Application.CreateColorContext()` 参数不可传 |
| `IVGColorManager` | 31 | **新打通（第十八波）**：`Application.ColorManager` 的 ColorProfiles（本机 64 个）/ GetProfilesByColorModel(0=RGB,1=CMYK,2=Gray,3=Lab) / DefaultColorContext；用于选配置文件并转文档颜色 |
| `IVGColorProfile` | 10 | **新打通（第十八波）**：Name / ColorModel / FileName / Installed / Generic / Manufacturer 可读，`CreateColorContext()` 得到该配置文件的上下文 |
| `IVGImage` | 5 | **新打通**：`Document.CreateImage` 返回像素缓冲区；`Width / Height / Type / ReadOnly / Pixel / FillArea / Blit` 可用（`cdr_create(type="bitmap")`、`cdr_bitmap(action="pixel")`） |
| `IVGDuotone` / `IVGDuotoneInk` | 8 / 8 | **新打通**：`app.CreateDuotone()` → `Type 0~3`、`Inks(i).Color`、`Load(.cpd)`；配合 `Bitmap.ConvertToDuotone` 转双色调 |
| `IVGPalettes` / `IVGPalette` | 8 / 15 | **新打通**：`Palettes.OpenFixed(id)` 可打开 PANTONE 匹配系统（946 色）等内置调色板；`Create/AddColor/SaveAs/Color/ColorCount/Identifier/FindColor` 可用；`CreatePaletteColor(Identifier GUID, index)`、`CreateSpotColorByName(Identifier, 名称, tint)` 可用 |
| `IVGFontList` | 3 | **新打通**：`app.FontList.Count / Item(i)` 返回字体名字符串（本机 2068 个） |
| `IVGSnapPoint` | 9 | **新打通**：`Page.FindClosestSnapPoint(TypeSet,x,y,hot)` 返回吸附点；`PositionX/PositionY/Type/Direction/UsesDirection/IsSelectable/IsMovable/IsDeletable` 可读（TypeSet 1/16/32 会报错，2/4/8 及组合可用） |
| `IVGPatternCanvas` / `IVGPatternCanvases` | 13 / 8 | **新打通（第八/十三波）**：`CreatePatternCanvas()` 生成 32×32 单色画布；`Clear / FillArea / SetPixel / Line / CopyArea / RotateArea / FlipArea / PutCopy / Select / Data` 可用；`PatternCanvases.Add(canvas)` 登记后 `Fill.ApplyPatternFill(0,"",index,front,end,transform)` 即可图案填充，`Transparency.ApplyPatternTransparency(...)` 图案透明 |
| `IVGPatternFill` | 18 | **新打通（第十三波）**：已有图案填充可微调 `FrontColor/BackColor/TileWidth/TileHeight/OriginX·Y/RotationAngle/SkewAngle/MirrorFill/MirrorFillX·Y/TransformWithShape/TileOffset/Canvas/Type`（全部可写且回读正确） |

---

## 二、能力域 → 工具（27 个）

| 能力域 | 工具 | 状态 |
|---|---|---|
| 会话/文档/状态 | `cdr_status` `cdr_connect` `cdr_documents` `cdr_open`（含**从模板创建**） `cdr_inspect` `cdr_save` `cdr_close` | ✅ |
| 绘图 | `cdr_create`（矩形含圆角 / 椭圆含圆弧扇形 / 多边形 / 星形 / 螺旋 / 网格 / 直线 / 贝塞尔 / 文字 / 段落文字 / 参考线 / 任意折线 / **内存位图（纯色 + 像素色块）**） | ✅ |
| 样式 | `cdr_modify`（纯色/渐变/纹理/图案（画布）/专色/无色、描边含虚线箭头/自定义虚线线型、投影、轮廓图、封套、透明/图案透明、冻结/解冻透明度、**图案填充微调**） | ✅ |
| 变换 | `cdr_modify`（移动/尺寸/旋转/翻转/斜切/矩阵/克隆/层级/**层级相对调整**/**清除变换**/**旋转中心**/复制/删除/锁定） | ✅ |
| **对象样式** | `cdr_style`（列出/从对象创建/派生/套用/改名/删除/读定义/改属性/**样式表 .cdst 导入导出**/**颜色样式与颜色协调**） | ✅ 第六/七/九波 |
| **节点/线段** | `cdr_node`（列节点/列线段/移动/选中/删除/圆角/倒角/scallop/平滑/拉伸/旋转/斜切/断开/加节点/断线/控制点/按坐标查找） | ✅ 第五波 |
| 文字排版 | `cdr_text`（字体字号粗斜体下划线 / 对齐 / 大小写 / 项目符号 / 首字下沉 / 行距字距 / 前后插入 / 沿路径 / 自适应 / 转曲 / 读取全文·区间·导出纯文本 / 字体属性读取·断字设置读取·**格式段列表** / 区间删除·复制·粘贴·拉直 / 按格式查询区间（@font/@bold…） / OpenType 特性） | ✅ |
| 布尔与整形 | `cdr_shape_ops`（weld/trim/intersect/combine/group/ungroup/break_apart/to_curves/to_symbol/powerclip/fillet/chamfer/scallop/step_repeat/distribute/boundary/perspective/separate/parallel_curves/rasterize/batch_outline_to_curves/fill_mode/pixel_aligned/range_boundary/envelope_from_shape/compare/project/unproject/to_document/extrude/lens/extrude_set/lens_set/effect_info/clear_effects/**blend/blend_set**） | ✅ |
| 对齐 | `cdr_align`（左中右/上下中/等间距，对齐对象或页面）+ `cdr_modify(align_to_shape/align_to_page)` | ✅ |
| 选择与查找 | `cdr_select`（加入/**移除/清空**）`cdr_find` | ✅ |
| 图层 | `cdr_layers`（增删改/显隐/锁定/打印/移动/**上下相邻查询·下移·粘贴剪贴板**） | ✅ |
| **跨文档 / 对象到图层** | `cdr_modify(copy_clipboard/cut_clipboard)` + `cdr_layers(action="paste")`；`cdr_modify(move_to_layer/copy_to_layer)` | ✅ 第八波 |
| 页面 | `cdr_pages`（列表/增删/插入空白页/切页/自定义尺寸/纸张预设/参考线/角度参考线/**吸附点**/解锁/包围盒） | ✅ |
| 位图 | `cdr_bitmap`（信息/转色彩模式/重采样/裁剪/取消裁剪/取像素色·图像缓冲信息·调色板模式 0~9·双色调/内容替换 SetImageData/**位图特效命令串 ApplyBitmapEffect**） | ✅（Trace 除外） |
| 路径 | `cdr_path`（长度面积/偏移轮廓/追加子路径/折线点/反转方向·焊子路径·去重叠·参数取点·切/法向量·速度·交点·找节点/线段） | ✅ |
| 导入导出 | `cdr_import` `cdr_export`（png/jpg/svg/pdf，可指定 DPI，可只导选区） | ✅ |
| 文档 | `cdr_doc`（概况/清选区/清撤销栈/调色板/箭头/箭头登记/单位换算·取色·显示隐藏·版式·重复操作/字体列表·调色板列表·调色板取色/固定调色板打开·颜色枚举·专色创建/图案画布创建（含区域变换）/纹理库列表/文档元数据读写/自定义调色板创建/数据字段·对象数据·色彩管理·线型样式/文档设置读写/应用信息·文档回滚/色彩配置文件·文档颜色转换/**用户偏好 KV**） | ✅ |
| **对象数据** | `cdr_modify(changes={"data":{...}})` + `cdr_doc(action="object_data"/"data_fields"/"add_data_field")` | ✅ 第十一波 |
| 视图 | `cdr_view`（刷新/激活/文档↔屏幕坐标/屏幕↔文档距离） | ✅ 第五波 |
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
| 样式分组属性 | `Style.GetProperty("fill"/"outline")` | 报「为参数 Name 指定的值无效」 | 分组名不可读写，只能读写叶子属性（primaryColor / type / angle…） |
| 按拷贝打开 | `Application.OpenDocumentAsCopy(FileName, Options)` | 无论 Options 传 0 还是 None 都是 `TypeError: Python instance can not be converted to a COM object` | 只能 OpenDocument / 从模板创建；需要"副本"请 SaveAs 新路径 |
| HTML 兼容文本 | `Text.MakeHTMLCompatible(HTML)` | 返回 False，未见效果 | 不封装 |
| 断字设置写入 | `Text.SetHyphenationSettings(Frames, struct)` | 调用不报错但值不生效（读回仍是默认） | 只读暴露，写入不提供 |
| 字体结构体读取 | `Text.GetFontProperties(Frames)` | 返回**默认结构体**（Name=''、Size≈1e-38）；UI 字体实际值要用 `Text.Story.Font/Size` | 已改用 Story 读取 |
| 文本区间套样式 | `TextRange.ApplyStyle(样式名)` | 调用不报错，但实测 `Bold/Size` 未变化（字符样式未生效或需其他条件） | 不推荐；整体样式用 `Shape.ApplyStyle` |
| 文本格式查询的尺寸比较 | `TextRange.FindRanges("@size > 10")` | 报 `Incompatible unit powers of the operands: 'pt' and 'pt^0'`；`@size > 10pt` 解析失败 | 尺寸条件暂不可用；`@font = 'Arial'`、`@bold = true` 可用 |
| 位图换内容 | `Bitmap.SetImageData(Image, Alpha, X, Y)` | `Alpha=0` 报 `TypeError: Python instance can not be converted to a COM object`；**Alpha=None 可用**，像素内容被替换 | 替换像素必须传 None |
| 位图特效 | `Bitmap.ApplyBitmapEffect(UndoString, Command)` | 两个都是字符串；`Command` 是**效果命令串**，如 `"WhirlpoolEffect WhirpoolSpacing=20,WhirlpoolSmear=9,WhirlpoolTwist=70,WhirlpoolStreak=60,WhirpoolWarp=1,WhirlpoolRandomSeed=0"`，实测像素被改变 | 效果名/参数名需按 Corel 文档；已封装 `cdr_bitmap(action="effect")` |
| 尺寸/连接线 | `Layer.CreateLinearDimension/CreateConnector` | 需要 IVGPoint；`MathUtils` CLSID 未注册，SubPath 取的点仍报 TypeError/类型不匹配 | 仍不可用；`CreateFreeConnector` 报「未实现」 |
| 符号库 | `Application.SymbolLibraries` | 本机 Count=0，无符号库可 Add/插入 | 需先安装符号库文件 |
| 文档颜色转换 | `Document.ConvertToColorContext(profile.CreateColorContext())` | RGB→CMYK→RGB 实测成功（填充色由 RGB 变 CMYK 再回 RGB）；**灰度→RGB 会报「服务器出现意外情况」** | 已封装 `cdr_doc(action="convert_color")`；转换前建议先以 RGB→CMYK→RGB 顺序，避免灰→RGB |
| 色彩配置文件 | `Application.ColorManager.ColorProfiles` | 本机 64 个；`GetProfilesByColorModel(0/1/2)` 分别可取 RGB/CMYK/Gray；`DefaultColorContext` 含 3 个 | `CreateColorContext()` 直接调用不可用，须用 `profile.CreateColorContext()` |
| 应用偏好 | `RegisterUserApplicationPreference` / `Get·SetApplicationPreferenceValue` | 用户偏好可注册、读写（存 CorelDRAW 工作区）；读取真实应用偏好组（如 `VGDoc Pref Settings`）报「发生意外错误」 | 只支持用户自定义偏好（Group/Key） |
| 文档回滚 | `Document.Revert()` | 可用（修改后 15→16→15 对象）；**破坏性**，丢弃未保存修改 | 已封装为 `cdr_doc(action="revert")`，调用前确认已保存 |
| 立体化/透镜创建 | `Shape.CreateExtrude(...)` / `Shape.CreateLens(...)` | 只传前几个 int 参数报 `TypeError: Python instance can not be converted to a COM object`；**传全参数（含颜色对象）即可** | 已封装并补齐默认颜色参数 |
| 点对象创建 | `Application.MathUtils.CreatePoint` | `MathUtils` CLSID `{E98800C5-…}` 报「没有注册类」，Application 上也没有 MathUtils 属性 | 尺寸标注/连接线仍不可用 |
| 混合创建 | `Shape.CreateBlend(shape, …)` | 省略可选参数会 `TypeError`（Path 参数默认 0 不是 IDispatch）；**传全参数且 Path=None 可用** | 已封装 `cdr_shape_ops(op="blend")` |
| 图案填充（文件型） | `Fill.ApplyPatternFill(Type=1/2, FileName=…)` | 提示「无法装入以下图样文件」 | 需要图案图像文件；**内存画布型（Type=0 + PatternCanvas）已可用** |
| 影线填充 | `Fill.ApplyHatchFill` / `Document.HatchLibraries` | `HatchLibraries.Count = 0`、`ActiveLibrary = None`（本机未装 Hatch 库） | 需要先安装 Corel 的 Hatch 填充库 |
| 曲线贴到笔触 | `Curve.CreateCurveMappedToStroke` | `com_error(-2147352571, 类型不匹配)`（Stroke 参数不是普通 Curve） | 需要 Artistic Media 的笔触对象，暂不封装 |
| 符号库 | `app.SymbolLibraries` | 对象存在，但 `Count = 0`（本机未装符号库） | `CreateSymbol` 需要库内符号名，无库可用（`ConvertToSymbol` 可用，已封装为 `to_symbol`） |
| 剪贴板 | `Document.Copy/Paste` | `AttributeError` | Document 上没有这两个方法（但 `Layer.Paste` 可用，已封装为 `cdr_layers(action="paste")`） |
| 吸附设置 | `Document.SnapToGrid/SnapToObjects` | `AttributeError` | 同上 |
| PowerClip 提取 | `Shape.PowerClip` | 对象存在但无 `Extract` | 只能"放入"不能"取出" |
| 页面背景 | `Page.Background` | 不存在 | — |
| 文档元数据（直接属性） | `Document.Keywords/Notes` | 直接**读**会 AttributeError，但**可以赋值**（write-only，写入后 Metadata 能读到） | 读元数据请走 `Document.Metadata` |

> 已修正的历史结论：
> * 折线取样 `Curve.GetPolyline` **可用**（返回 `IVGCurve`，遍历 `SubPaths/Nodes` 即可），
>   `cdr_path(action="polyline")` 已能返回采样点。
> * 节点级编辑 **可用**（`curve.Nodes` / `curve.Segments` 是属性），已封装为 `cdr_node`。
> * 对象样式 **可用**（`CreateStyleFromShape` 的 Category 是字符串：fill/outline/character/paragraph），
>   已封装为 `cdr_style`；`Shape.ApplyStyle("样式名")` 可套用。
> * 位图调色板/双色调 **可用**（`ConvertToPaletted(0~8)`、`app.CreateDuotone()` + `ConvertToDuotone`）。
> * 内存位图 **可用**（`Document.CreateImage` → `IVGImage` → `Layer.CreateBitmap`），
>   像素级 `FillArea/Blit/Pixel` 均可调用。
> * 专色 **可用**（第七波）：先用 `Palettes.OpenFixed(3)` 打开 PANTONE 匹配系统，
>   再用 `CreateSpotColorByName(调色板 GUID, 调色板里的精确颜色名, tint)`；
>   `cdr_doc(action="spot_color")` / `cdr_modify(fill={"type":"spot",...})` 已封装。
> * 对象样式属性/派生 **可用**：`Style.GetPropertyAsString / SetPropertyAsString / SetProperty`
>   与 `StyleSheet.CreateStyle(category, based_on, name)`（派生样式出现在 base 的 `DerivedStyles` 里）。
> * 范围批量 **可用**：`ShapeRange.ConvertOutlineToObject / SetFillMode / SetPixelAlignedRendering / CreateBoundary`。
> * 吸附点 **可用**：`Page.FindClosestSnapPoint(TypeSet=2/4/8 及组合)`。
> * 图案填充/图案透明 **可用（第八波）**：`app.CreatePatternCanvas()` 画 1-bit 图案，
>   `PatternCanvases.Add(canvas)` 后 `ApplyPatternFill(0,"",index,front,end,…)` 直接填充。
> * 跨文档复制 **可用（第八波）**：`Shape.Copy/Cut` 进系统剪贴板，目标文档 `Layer.Paste` 取出。
> * 对象换图层 **可用**：`Shape.MoveToLayer / CopyToLayer`。
> * `Document.CreateFill/CreateOutline` 传字符串会返回 `IVGFill/IVGOutline`，
>   但应用后填充类型为 0（无效），暂无可用字符串格式。
> * 模板创建 **可用（第九波）**：`Application.CreateDocumentFromTemplate(路径, 带图形)`，
>   实测 **.cdr 文件也能当模板**；已封装为 `cdr_open(template=...)`。
> * 样式表导入导出 **可用**：`Document.SaveStyleSheet/LoadStyleSheet`（.cdst），
>   删除样式后 Load 可恢复。
> * 颜色样式 **可用**：`Document.CreateColorStyles(harmonies=1..3)` 生成 5 个左右颜色样式，
>   `StyleSheet.AllColorStyles` 可枚举名称与颜色串。
> * 纹理填充 **可用**：`Fill.ApplyTextureFill("BARK","Textures")` → Fill.Type=8；
>   `cdr_doc(action="textures")` 可列出本机 31 个纹理名。
> * 3D 平面投影 **可用**：`Shape.Project/Unproject(plane 0/1/2, referencePoint)`，
>   project 后 bbox 改变，unproject 可精确恢复。
> * 文档元数据 **可用（第十波）**：`Document.Metadata` 可读写 Title/Author/Subject/
>   Keywords/Notes/Copyright（直接属性 `Document.Keywords` 不存在）。
> * 对象提取成新文档 **可用**：`Shape.CreateDocumentFrom` / `ShapeRange.CreateDocumentFrom`，
>   返回新 `IVGDocument`（1 页、副本）。
> * 自定义调色板 **可用**：`Palettes.Create(名称, 路径)` + `AddColor` + `SaveAs`，
>   落盘约 260 B；返回的 palette_index 可用于位图 9 号调色板模式。
> * 文档设置 **可读写（第十二波）**：`EditAcrossLayers / ApplyToDuplicate / CurvePrecision /
>   Resolution / PreserveSelection / ShapeEnumDirection / TextFormatter / WorldScale /
>   FacingPages / FirstPageOnRightSide / Layout / DrawingOriginX·Y` 等均为可写属性。
> * 文本区间 **可编辑**：`TextRange.Delete/Copy/Paste/Straighten` 实测可用（copy/paste 走系统剪贴板）。
> * 自定义箭头 **可登记**：`CreateArrowHead(curve)` → `Application.ArrowHeads.Add()` 进全局箭头库。
> * 文本格式查询 **可用（第十三波）**：`TextRange.FindRanges("@font = 'Arial'")`、
>   `"@bold = true"` 返回 `IVGTextRanges`；尺寸比较因单位语法暂不可用。
> * 图案填充微调 **可用**：`Fill.Pattern` 的 RotationAngle / TileWidth / Origin / MirrorFill /
>   FrontColor 等 15 个属性可写且回读正确。
> * 图案画布区域变换 **可用**：`CopyArea / RotateArea / FlipArea / PutCopy`。
> * 位图内容替换 **可用（第十四波）**：`Bitmap.SetImageData(image, None, x, y)`，
>   Alpha 传 None；可直接把新建像素缓冲写进已有位图。
> * 文本格式段 **可读**：`TextRange.EnumRanges(-1)` 逐段返回文字/字体/字号/粗斜体/起止。
> * 文档回滚 **可用**：`Document.Revert()` 丢弃未保存修改、回到上次保存状态（破坏性）。
> * 立体化/透镜 **可用（第十五波）**：`CreateExtrude/CreateLens` 传**全参数（含颜色对象）**可创建；
>   效果经 `Shape.Effects.Item(i).Extrude/.Lens` 读写（IVGEffectExtrude/IVGEffectLens）。
> * 效果管理 **可用**：`Shape.Effects.Count/Item`、`Effect.Separate()/Clear()`；`clear_effects` 可一键清除。
> * 位图特效 **可用（第十六波）**：`ApplyBitmapEffect(UndoString, Command)` 命令串，
>   例 Whirlpool 命令串实测改变像素。
> * 尺寸标注/连接线 **仍不可用**：`MathUtils` 未注册、点对象无法构造；
>   `CreateFreeConnector` 报「未实现」；本机无符号库。
> * 混合 **可用（第十七波）**：`CreateBlend` 传全参数且 `Path=None` 时成功；
>   效果经 `Effects.Item(i).Blend`（IVGEffectBlend）读写 Steps/Spacing/Angle 等。
> * 文档颜色转换 **可用（第十八波）**：`Application.ColorManager` 提供 64 个配置文件，
>   `profile.CreateColorContext()` → `Document.ConvertToColorContext()` 可 RGB↔CMYK↔Gray；
>   灰度→RGB 会报服务器异常，其余方向正常。

---

## 四、已知但未封装（低频）

这些**能用**，只是 Agent 日常用不上，暂不占工具位：

* `IVGWindow` 的缩放/新建窗口（`Zoom`、`NewWindow`）：截图/预览才需要
* `IVGColorContext`：色彩配置、配置文件合并（交付环节）
* `IVGFountainFill.GetTransformations`：读渐变变换矩阵（渐变微调已能改角度/平移/矩阵）
* `IVGTraceSettings`：PowerTRACE 参数对象（Trace 本身不安全，不封装）
* `IVGStyle` 的 `SetBasedOn` / `GetCopy` / `Assign`：派生已用 `CreateStyle(based_on)`，改父级暂未暴露
* `Shape.CopyPropertiesFrom(Source, Properties)`：实测 1~15 都返回 True，但位掩码语义未逆向出来
* `Curve.GetCurveInfo`：返回 `com_struct(PositionX, PositionY, ElementType, NodeType, Flags)` 数组，语义未完全确认
* `Layer.PasteEx` / `PasteSpecial`：`PasteEx` 传 0 会 TypeError；`Paste` 已封装
* `IVGImage.Blit` / `FlipArea`：像素缓冲区合成可用（已实测 Blit），暂未做成工具参数
* `Application.OpenDocumentAsCopy` / `CreateDocumentFromTemplate`：需要模板/文件路径，未纳入工具面
* `Application.CreateHSBColor/CreateYIQColor/CreateCMYColor/CreateBWColor`：实测可用，但颜色运算已走纯 Python
* `Palettes.Create / CreateFromDocument / CreateFromSelection`：可新建调色板并 `AddColor/SaveAs`，
  但会落盘点色板文件，暂未做成工具（`OpenFixed` 已封装）
* 符号库插入：需要符号库名称，`CreateSymbol` 本身也未打通；`Page.FindClosestSnapPoint` 的
  TypeSet 1/16/32 会报「服务器出现意外情况」（2/4/8 及组合可用）
* `Bitmap.ResolveLink/UpdateLink`：需要外部链接位图才有意义，未做联调
* `Bitmap.ConvertToPaletted(9=自定义)`：需要 `palette_index`（已支持，配合 `palette_open_fixed`）

---

## 五、复现方式

```powershell
python -m pytest -q              # 55 项纯逻辑
python scripts/live_smoke.py     # 234 项端到端（真实 CorelDRAW）
```

联调断言覆盖了上表所有 ✅ 项。反解脚本属于过程产物，已按要求放在 `tmp/` 并在完成时清理；
如需重建，按本文件"怎么得出这份清单"三步即可复现。

---

## 六、量化覆盖率（2026-10-25 第十九波实测）

方法：反解类型库拿到 196 个自动化类 / **1619** 个方法（已剔除 `QueryInterface` 等 COM 基础设施）；
用 AST 扫描 `src/cdrx8_mcp/*.py`，收集所有 `.方法名(` 调用与 `setattr/getattr` 的属性名，
与主干接口的方法名求交集。**注意：属性名（如 `curve.Nodes`）也会被计入**，所以本波数字含真实属性访问。

**全局：696 / 1619 = 43.0%**（21.2% → … → 42.8% → 本轮 43.0%）
**主干 20 接口：416 / 743 = 56.0%**

复现命令：`python scripts/coverage_audit.py`（会临时展开类型库到 `tmp/gen_py`，跑完自动清理）

| 接口 | 方法数 | 已调用 | 覆盖 |
|---|---:|---:|---:|
| `IVGShape` | 123 | 80 | **65%** |
| `IVGShapeRange` | 125 | 77 | 62% |
| `IVGSubPath` | 38 | 29 | **76%** |
| `IVGDocument` | 74 | 38 | 51% |
| `IVGPage` | 21 | 19 | **90%** |
| `IVGCurve` | 36 | 22 | 61% |
| `IVGLayer` | 49 | 22 | 45% |
| `IVGTextRange` | 31 | 22 | **71%** |
| `IVGApplication` | 78 | 25 | **32%** |
| `IVGText` | 29 | 16 | **55%** |
| `IVGNode` | 18 | 9 | 50% |
| `IVGBitmap` | 15 | 11 | **73%** |
| `IVGFill` | 12 | 7 | **58%** |
| `IVGWindow` | 8 | 7 | **88%** |
| `IVGFountainFill` | 6 | 5 | 83% |
| `IVGColor` | 43 | 9 | 21%（色彩运算已改纯 Python，见下） |
| `IVGOutline` | 8 | 3 | 38% |
| `IVGTransparency` | 8 | 6 | **75%** |
| `IVGStyle` | 15 | 8 | **53%** |
| `IVGColorContext` | 6 | 1 | 17% |
| **合计** | **743** | **416** | **56.0%** |

> **怎么读这两个数**：
> 1. 全局 43.0% 覆盖 1619 个方法（含变体与重复）；主干 20 接口 56.0%。
> 2. 它是**方法/属性名匹配**的近似值，同名方法会跨接口误计。
> 3. `IVGColor` 的 21% 是**故意的**：实测不暴露分量属性，色彩运算在纯 Python 完成。
> 4. 未调用方法多为变体（Wide/Ex/2/3 后缀）与事件订阅、CorelScriptTools/MathUtils。

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

## 七之四、第六波已接入（2026-10-25，对象样式/内存位图/字体/调色板/页面插入）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **对象样式**（此前误判为不可用） | `cdr_style`：list / get / create / apply / rename / delete | ✅ 创建 fill/outline/character 样式；`ApplyStyle("名")` 套用；可改名/删除；读到 69 个属性名与 ToString JSON |
| 修正「Category 常量」误判 | `StyleSheet.CreateStyleFromShape(shape, category: str, name)` | ✅ Category 是字符串：fill / outline / character / paragraph；数字 0~8 会报「未找到类别」 |
| **内存位图** | `cdr_create(type="bitmap", color, pixels_w/h, fills=[...])` | ✅ `Document.CreateImage → IVGImage → Layer.CreateBitmap` 落地为位图对象（type 5） |
| 像素级读取 | `cdr_bitmap(action="pixel", x_px, y_px)` | ✅ 读到色块「蓝」/底色「红」；越界明确报错 |
| 图像缓冲信息 | `cdr_bitmap(action="image_info")` | ✅ 返回缓冲 16x16、类型 4(RGB)、只读标记 |
| **调色板模式** | `cdr_bitmap(action="paletted", palette_type=3)` | ✅ Mode 4 → 3；0~8 调色板类型均可用 |
| **双色调** | `cdr_bitmap(action="duotone", duotone_type=1, ink_colors=[...])` 或 `preset=.cpd` | ✅ Mode 4 → 6；`app.CreateDuotone()` + `Inks(i).Color`；本机自带 30+ 个 .cpd 预设 |
| **插入页面**（此前未打通） | `cdr_pages(action="insert", count, page_index, insert_before)` | ✅ 1→2→3 页；第三参是**页码数字**不是 Page 对象；插到首页前会切换活动页（已返回 active_page） |
| **并行曲线**（范围级） | `cdr_shape_ops(op="parallel_curves", curve_count, curve_distance)` | ✅ 曲线对象生成 2 条偏移曲线；矩形需先 to_curves |
| **选择集栅格化** | `cdr_shape_ops(op="rasterize", resolution, grayscale, transparent)` | ✅ 选中的矩形换成位图（type 5） |
| 图层相邻/下移 | `cdr_layers(action="above"/"below"/"move_below")` | ✅ 读到上下邻居名；下移后排到目标图层下方 |
| 图层粘贴 | `cdr_layers(action="paste")` | ✅ 返回粘贴出的对象（依赖系统剪贴板） |
| 冻结/解冻透明度 | `cdr_modify(changes={"freeze_transparency":true/false})` | ✅ 解冻后恢复；**冻结会把对象换成无名组合体（type 7）**，工具会自动补名并可继续定位 |
| 字体列表 | `cdr_doc(action="fonts", query, limit)` | ✅ 本机 2068 个字体；按需扫描（默认扫 max(300, limit×10) 条，返回 scanned） |
| 调色板列表/取色 | `cdr_doc(action="palettes"/"palette_color")` | ✅ 读到文档调色板（3 色）与默认调色板（100 色）、GUID 与颜色串 |
| **OpenType 特性** | `cdr_text(opentype={"liga":true,"kern":false})` | ✅ 写入并回读 liga=1 / kern=0 |

## 七之五、第七波已接入（2026-10-25，专色/无关调色板、样式属性与派生、范围批量、包络、吸附）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **固定调色板** | `cdr_doc(action="palette_open_fixed", palette_id)` | ✅ 本机 `OpenFixed(3)` = PANTONE MATCHING SYSTEM（946 色）；同时可打开 TRUMATCH、FOCOLTONE、TOYO、SpectraMaster、标准色等（编号随安装变化） |
| 调色板颜色枚举 | `cdr_doc(action="palette_colors", palette_index, query, limit)` | ✅ 按名称过滤（如 "Yellow" 命中 2 个），返回颜色索引/名称/颜色串 |
| **专色创建** | `cdr_doc(action="spot_color", palette_index, color_index/spot_name, tint)` | ✅ `PANTONE Process Yellow CV` → `PANTONE,PAN_SPOT,1,100,<guid>` |
| **专色填充 / 描边** | `cdr_modify(fill={"type":"spot",...})` / `outline={"color":{"type":"spot",...}}` | ✅ 填充 applied=`fill:spot`；描边 `Outline.Color.CopyAppearance(专色)` 已生效 |
| 自定义调色板位图 | `cdr_bitmap(action="paletted", palette_type=9, palette_index=…)` | ✅ Mode 4 → 3（自定义调色板） |
| 样式属性读写 | `cdr_style(action="get")` 的 `readable_properties`；`set_property` | ✅ 69 个可读属性；`angle=15` 写入并回读 `"15"`；颜色属性走 `SetProperty(名字, IVGColor)` |
| **样式派生** | `cdr_style(action="create", category, based_on, style_name)` + `derived` | ✅ 子样式出现在父样式的 `DerivedStyles` 中，`based_on` 正确 |
| 范围描边转曲 | `cdr_shape_ops(op="batch_outline_to_curves")` | ✅ 选择集一次转曲（页对象 +1） |
| 填充规则 / 像素对齐 | `cdr_shape_ops(op="fill_mode"/"pixel_aligned")` | ✅ `SetFillMode(0~2)`、`SetPixelAlignedRendering(True/False)` |
| 范围边界 | `cdr_shape_ops(op="range_boundary", x, y)` | ✅ 生成边界对象并自动补名 |
| **包络（另一对象为形）** | `cdr_shape_ops(op="envelope_from_shape", other, envelope_mode)` | ✅ 曲线来源走 `CreateEnvelopeFromCurve`，形状来源走 `CreateEnvelopeFromShape`（返回 `IVGEffect`/None 均视为已调用） |
| **对象比较** | `cdr_shape_ops(op="compare", other, compare_type, condition)` | ✅ 类型比较 1/2/64/128 可用，返回布尔；位义见 constants（Fill=4、OutlineWidth=8…） |
| 选择集移除/清空 | `cdr_select(remove=true / clear=true)` | ✅ 移除单个对象、清空选择集 |
| **吸附点查询** | `cdr_pages(action="snap_point", x, y, snap_types, hot_area)` | ✅ TypeSet=2/4/8/3/6/15/63 返回 `IVGSnapPoint` 位置/类型；1/16/32 报服务器错误 |

## 七之六、第八波已接入（2026-10-25，图案画布、跨文档剪贴板、对象到图层）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **图案画布** | `cdr_doc(action="pattern_canvas", pattern_fills/pattern_pixels/pattern_lines)` | ✅ `CreatePatternCanvas()` 生成 32×32 单色画布；`FillArea/SetPixel/Line` 按 bool 状态绘制；`PatternCanvases.Add` 返回 canvas_index |
| **图案填充（无需 .pat 文件）** | `cdr_modify(fill={"type":"pattern","canvas_index":N,"front","end"})` | ✅ `ApplyPatternFill(0,"",index,front,end,False)` → `Fill.Type=9`；类型 1/2 仍需图像文件 |
| **图案透明** | `cdr_modify(transparency={"type":"pattern","canvas_index":N,"front","back"})` | ✅ `ApplyPatternTransparency(0,"",index,0,100,False)` |
| **跨文档复制** | `cdr_modify(copy_clipboard/cut_clipboard)` + `cdr_layers(action="paste")` | ✅ 对象进系统剪贴板；新建 A5 文档后 `Layer.Paste` 取回同类型同尺寸对象；剪切会从源文档移除 |
| **对象到图层** | `cdr_modify(move_to_layer/copy_to_layer)` | ✅ `MoveToLayer` 后对象计数从原图层消失、目标图层 +1；`CopyToLayer` 返回副本并自动补名 |
| 填充/描边对象复制与比较 | `Fill.GetCopy / CompareWith`、`Outline.GetCopy / CompareWith`、`ShapeRange.ApplyFill/ApplyOutline` | ✅ 返回填充/描边对象、比较布尔；范围应用可用 |
| 虚拟填充/描边字符串 | `Document.CreateFill/CreateOutline(FillString)` | ⚠️ 任何字符串都会返回对象，但应用后填充类型为 0（无效）；字符串格式未解 |

## 七之七、第九波已接入（2026-10-25，模板、样式表、颜色样式、纹理、3D 投影、字体读取）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **从模板创建文档** | `cdr_open(template=路径, include_graphics=True)` | ✅ `CreateDocumentFromTemplate` 返回文档；实测 **.cdr 也能当模板**（页数/图形随模板） |
| **样式表导出/导入** | `cdr_style(export_sheet/import_sheet, sheet_path=…)` | ✅ 导出 .cdst（≈395 KB）；删除样式后 Load 可恢复，且可继续 ApplyStyle |
| **颜色样式 / 颜色协调** | `cdr_style(create_color_styles, harmonies=1)` + `color_styles` | ✅ `CreateColorStyles` 后 AllColorStyles 0→3/5；名称（颜色样式N）与颜色串可读 |
| **纹理填充** | `cdr_doc(action="textures")` + `fill={"type":"texture","name":"BARK","library":"Textures"}` | ✅ 本机 31 个纹理名；填充后 Fill.Type=8 |
| **3D 平面投影** | `cdr_shape_ops(op="project"/"unproject", plane=0/1/2, reference_point)` | ✅ 投影后 bbox 改变；反投影 bbox 精确恢复 |
| 字体属性读取 | `cdr_text(read_font=true)` | ✅ Story 读到 Arial/12；`GetFontProperties` 结构体不可信（默认值），已弃用 |
| 断字设置读取 | `cdr_text(read_hyphenation=true)` | ✅ 读到自动断字/最小词长/前后字符/HotZone；**写回不生效** |
| 拷贝打开 | `Application.OpenDocumentAsCopy` | ❌ 任何 Options 都 `TypeError`，不封装 |
| HTML 兼容文本 | `Text.MakeHTMLCompatible` | ❌ 返回 False、无效果，不封装 |

## 七之八、第十波已接入（2026-10-25，文档元数据、对象提取新文档、自定义调色板）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **文档元数据读写** | `cdr_doc(action="metadata"/"set_metadata", metadata={...})` | ✅ `Document.Metadata` 读 Title/Author/LastAuthor/Subject/Keywords/Notes/Copyright/Revision/DocLanguage/DocID；写 Title/Author/Subject/Keywords/Notes/Copyright 后回读正确 |
| **对象/选区提取成新文档** | `cdr_shape_ops(op="to_document")` | ✅ `Shape.CreateDocumentFrom(False)` → 新 `IVGDocument`（1 页、1 对象）；选区走 `ShapeRange.CreateDocumentFrom`（2 对象）；返回 `document_id` 可直接 `cdr_inspect` |
| **自定义调色板创建** | `cdr_doc(action="palette_create", palette_name, palette_path, palette_colors)` | ✅ `Palettes.Create` + `AddColor`×3 + `SaveAs` → 落盘 258 B；返回 palette_index，`palette_colors` 可列出 3 色 |
| 自定义调色板用于位图 | `cdr_bitmap(paletted, palette_type=9, palette_index=…)` | ✅ Mode 4→3 |
| 混合 | `Shape.CreateBlend` | ✅ 已接入（第十七波，须全参数 + Path=None） |
| 曲线贴笔触 | `Curve.CreateCurveMappedToStroke` | ❌ 传 shape / curve / `_oleobj_` 均「类型不匹配」，需要 Artistic Media 专用笔触对象 |
| 填充字符串 | `Document.CreateFill/CreateOutline` | ❌ 任意字符串返回 Type=0 的空填充；无 ToString 可反推格式 |
| 事件订阅 | `Application.AdviseEvents` | ⚠️ 未接入（需要连接点 sink + STA 消息泵，暂不冒险） |

## 七之九、第十一波已接入（2026-10-25，对象数据字段、自定义虚线、色彩管理读取）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **对象数据字段** | `cdr_doc(action="add_data_field"/"data_fields"/"object_data")` + `cdr_modify(changes={"data":{...}})` | ✅ `DataFields.Add("W11SKU","0")`；`ObjectData.Add(field,"A-123")` → `DataItem.Value` 可读；更新为 B-456 后回读正确；数据持久化在 .cdr 里 |
| **自定义虚线线型** | `cdr_modify(outline={"dash_pattern":[2,2,0.5,2]})` | ✅ `CreateOutlineStyle(2,2,0.5,2)` 生成虚拟线型（Index=-1）→ `OutlineStyles.AddStyle()` 注册为 Index 28 → 赋给 `Outline.Style` 生效；可 `Save()` 持久化 |
| 线型样式列表 | `cdr_doc(action="outline_styles")` | ✅ 27 个内置线型 + 自定义线型，可读 Index/DashCount/Enhanced |
| **色彩管理读取** | `cdr_doc(action="color_context")` | ✅ RenderingIntent=3、BlendingColorModel=1、profiles="sRGB IEC61966-2.1,Japan Color 2001 Coated,Dot Gain 15%" |
| 文档组件 | `Document.Components` | ⚠️ 本机 Count=0（无组件），暂不封装 |
| 数据合并格式 | `DataFields.Add(Name, Format)` | Format 必须是数字格式串（"0" 可用；"General"/""/"@" 均报「格式字符串无效」） |

## 七之十、第十二波已接入（2026-10-25，文档设置、文本区间编辑、自定义箭头登记）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **文档设置读写** | `cdr_doc(action="settings"/"set_settings")` | ✅ 可读 EditAcrossLayers/ApplyToDuplicate/CurvePrecision/Resolution/PreserveSelection/ShapeEnumDirection/TextFormatter/WorldScale/FacingPages/FirstPageOnRightSide/Layout/Unit/DrawingOriginX·Y；可写前 10 项（写后回读正确） |
| **文本区间编辑** | `cdr_text(range_delete/range_copy/range_paste/range_straighten)` | ✅ 删除 `[6,11]` → "Alpha Gamma"；复制 `[0,5]` 后粘贴到 0 → "AlphaAlpha Gamma" |
| **自定义箭头登记** | `cdr_doc(action="arrowhead", register=true)` | ✅ `CreateArrowHead(curve)` → `ArrowHeads.Add()`，箭头库 89→90，可赋给 `Outline.StartArrow` |
| 文档 Keywords/Notes | `Document.Keywords/Notes` | ✅ **可写**（write-only）；写入后 `Document.Metadata.Keywords/Notes` 可读 |
| 保存/恢复设置 | `Document.SaveSettings/RestoreSettings(Tag)` | ⚠️ 调用不报错，但实测未恢复 EditAcrossLayers（是否覆盖全部设置未确认），暂不封装 |
| 位图链接 | `Bitmap.ResolveLink/UpdateLink` | ⚠️ 非外部链接位图报「未外部链接到文件」，本机无链接位图可测 |

## 七之十一、第十三波已接入（2026-10-25，图案画布变换、图案填充微调、文本格式查询）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **图案画布区域变换** | `cdr_doc(pattern_canvas, copy_areas/rotate_areas/flip_areas)` | ✅ `CopyArea/RotateArea/FlipArea` 各调用一次 + 填充/像素 → drawn=5；`PutCopy/SetItem/Select` 也可用 |
| **图案填充微调** | `cdr_modify(pattern_fill={rotation_angle, origin_x, tile_width, mirror_fill, front_color, skew_angle})` | ✅ 15 个 PatternFill 属性可写且回读正确（RotationAngle=45、TileWidth=10、FrontColor=HSB 等） |
| **文本格式查询** | `cdr_text(find_query="@font = 'Arial'")` / `"@bold = true"` | ✅ 返回 `IVGTextRanges`；font 查询命中整段、bold=true 命中 0；单引号字符串语法可用 |
| 文本尺寸查询 | `FindRanges("@size > 10")` | ❌ 报 `Incompatible unit powers: 'pt' and 'pt^0'`；`10pt` 又解析失败 |
| 文本区间套样式 | `TextRange.ApplyStyle("样式名")` | ⚠️ 调用不报错，但 Bold/Size 未变化，效果不可观测 |
| 自定义箭头/文档设置 | 见第十二波 | ✅ |

## 七之十二、第十四波已接入（2026-10-25，位图内容替换、文本格式段、应用信息、文档回滚）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **位图内容替换** | `cdr_bitmap(action="replace_image", image_color, image_pixels_w/h, image_fills, image_pixels, offset_x/y)` | ✅ `Document.CreateImage` 建像素缓冲 → `SetImageData(image, None, x, y)`；替换后 (1,1)=绿、(10,10)=蓝 |
| **文本格式段** | `cdr_text(runs=true)` | ✅ `EnumRanges(-1)` 返回 1 段：Text="Alpha Beta Gamma"、Font=Arial、Size=12、Bold=false、Start=0、End=16 |
| **应用信息** | `cdr_doc(action="app_info", limit=N)` | ✅ OpenType 特性 230 个（可过滤）；平台检查 16/17=true、18=false；UI 语言 1033/2052=false |
| **文档回滚** | `cdr_doc(action="revert")` | ✅ 已保存副本 15→建线 16→Revert 15；**破坏性** |
| 位图 SetImageData Alpha | `Bitmap.SetImageData(img, None, x, y)` | ✅ Alpha=None；传 0 会 TypeError |
| 文本区间套样式 | `TextRange.ApplyStyle` | ⚠️ 仍未观测到效果（沿用第十三波结论） |

## 七之十三、第十五波已接入（2026-10-25，立体化、透镜、效果管理）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **立体化（Extrude）** | `cdr_shape_ops(op="extrude", extrude_type, vp_type, vp_x/y, depth, shading, base_color, shading_color, bevel_depth, bevel_angle, bevel_color, bevel_only, effect={...})` | ✅ 创建后 `Effects.Count=1`、`Item(1).Extrude` 为 `IVGEffectExtrude`；depth=25；`extrude_set` 改到 35 |
| **透镜（Lens）** | `cdr_shape_ops(op="lens", lens_type, rate, color1, color2, color_map_palette, effect={...})` | ✅ 创建后 `Effects.Count=1`、kind=lens、rate=60；`Effect.Type=9` |
| **效果枚举/清除** | `cdr_shape_ops(op="effect_info"/"clear_effects")` | ✅ effect_info 返回 count/kind/props；clear_effects 清 1 个后 count=0 |
| 效果属性 | `IVGEffectExtrude` / `IVGEffectLens` | ✅ Type/Depth/AngleX·Y·Z/Shading/BaseColor/ShadingColor/UseBevel/BevelDepth/BevelAngle；Lens 的 Type/Rate/Color/FillColor/OutlineColor/FromColor/ToColor/ViewPoint 等可读写 |
| 关键坑 | `CreateExtrude/CreateLens` | 只传前几个 int 参数报 TypeError；**必须传全参数（含颜色对象）** |
| MathUtils | `Application.MathUtils` | ❌ CLSID `{E98800C5-…}` 未注册、Application 无属性；尺寸/连接线点对象不可得 |
| 混合 | `CreateBlend` | ✅ 全参数 + Path=None 可用（第十七波） |

## 七之十四、第十六波已接入（2026-10-25，位图特效命令串）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **位图特效** | `cdr_bitmap(action="effect", effect_name=..., effect_command=...)` | ✅ `ApplyBitmapEffect("Whirlpool", "WhirlpoolEffect WhirpoolSpacing=20,WhirlpoolSmear=9,WhirlpoolTwist=70,WhirlpoolStreak=60,WhirpoolWarp=1,WhirlpoolRandomSeed=0")`；执行后像素 (1,1) 由绿(0,255,0) → (0,243,0) |
| 命令串格式 | 第一参数 Undo 名，第二参数 `效果名Effect 参数=值,…` | 参考 Corel VBA 文档；Whirlpool 已实测，其他效果需按文档给参数 |
| 尺寸/连接线 | `Layer.CreateLinearDimension/CreateConnector` | ❌ 点对象无法构造（MathUtils 未注册）；`CreateFreeConnector` 报「未实现」 |
| 符号库 | `Application.SymbolLibraries` | ❌ Count=0，本机无库 |

## 七之十五、第十七波已接入（2026-10-25，混合 Blend）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **混合（Blend）** | `cdr_shape_ops(op="blend", target, other, steps, spacing, angle, color_blend_type, blend_mode, loop, rotate_shapes, spacing_accel, color_accel, accel_size, effect={...})` | ✅ 两个椭圆混合 steps=5；`Effects.Item(1).Blend` 为 IVGEffectBlend，props steps=5/spacing=0/angle=0 |
| **混合微调** | `cdr_shape_ops(op="blend_set", effect={steps,spacing,angle,...})` | ✅ steps 5→8、spacing 0→2.0，回读正确 |
| 混合属性 | `IVGEffectBlend` | ✅ Steps/Spacing/Angle/ColorBlendType/Mode/Loop/RotateShapes/MapNodes/FuseStart/FuseEnd/ColorAcceleration/SpacingAcceleration/LinkAcceleration/AccelerateSize 可读写 |
| 关键坑 | `CreateBlend` | 省略可选参数会 TypeError（Path 默认 0 不是 IDispatch）；**必须传全参数且 Path=None** |
| 清除 | `clear_effects` | ✅ 混合效果清除后 Effects.Count=0 |

## 七之十六、第十八波已接入（2026-10-25，色彩配置文件与文档颜色转换）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **色彩配置文件列表** | `cdr_doc(action="color_profiles", model=0/1/2/3, limit=N)` | ✅ 本机 64 个配置；model=1 前 3 个含 Coated FOGRA27/FOGRA39/GRACoL；返回 name/model/generic/installed/manufacturer/file_name |
| **文档颜色转换** | `cdr_doc(action="convert_color", model=1)` / `profile_name="..."` | ✅ RGB 红 → CMYK(0,251,255,0) → 回 RGB(226,35,43)；上下文 profile 名随转换变化 |
| 按名字选配置 | `profile_name="Dot Gain 15%"` | ✅ 转灰度后 fill 变 GRAY，上下文为 Dot Gain 15% |
| 模型常量 | model 0=RGB 1=CMYK 2=Gray 3=Lab | 来自 ColorProfile.ColorModel |
| 灰度→RGB | `ConvertToColorContext` | ❌ 从灰度转 RGB 报「服务器出现意外情况」；其余方向正常 |
| 关键坑 | `Application.CreateColorContext()` | 直接调用 TypeError；必须 `profile.CreateColorContext()` |

## 七之十七、第十九波已接入（2026-10-25，用户级应用偏好 KV、当前工具信息）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **用户偏好注册** | `cdr_doc(action="pref_register", pref_group, pref_key, pref_default)` | ✅ `RegisterUserApplicationPreference("cdrX8 MCP", key, "default-val")`，随后 get 返回默认值 |
| **用户偏好读写** | `cdr_doc(action="pref_get"/"pref_set", pref_key, pref_value, pref_default)` | ✅ set 为 "hello-world" 后回读一致；未注册时可用默认值自动注册 |
| **当前工具信息** | `cdr_doc(action="app_info")` → `active_tool` / `active_tool_state` | ✅ 返回当前工具名与 GUID |
| 真实应用偏好 | `GetApplicationPreferenceValue("VGDoc Pref Settings", "KeepDesktopObjectsOnLayer")` | ❌ 报「发生意外错误」；只能读写用户注册的偏好 |
| 持久化范围 | CorelDRAW 工作区 | 同会话内已验证；跨重启持久化未单独验证 |

## 七之二、仍未接入（第三轮候选）

| 能力 | 接口 | 为什么还没做 |
|---|---|---|
| 位图调色板(自定义) | `Bitmap.ConvertToPaletted(9, Palette=...)` | 0~8 已支持；9 需要构造 Palette 对象，暂未接 |
| 自定义影线图案 | `Fill.ApplyHatchFill` / `Document.HatchLibraries` | 本机 Hatch 库为空（`Count=0`），没有可用图案 |
| 色彩管理 | `IVGColorContext`（6 个） | 交付环节，非创作环节 |
| 打印 | `Document.PrintOut()` | **故意不接**：会真的发出打印任务 |
| 跨文档合并 | `Document.InsertPages` 的跨文档形态 | 插入空白页已可用（`cdr_pages(action="insert")`）；跨文档内容合并未找到可用参数 |
| 事件订阅 | `IVGApplicationEvents` / `IVGDocumentEvents` | 需要 COM 连接点，pywin32 支持有限 |
| CorelScript 工具函数 | `ICorelScriptTools`（45 个：单位换算、三角、日期、文件夹遍历） | 大部分与 CorelDRAW 无关；单位换算已在 `mm_units` 内部处理 |
| 数学工具 | `IVGMathUtils`（19 个） | 同上 |

> 说明：`ICorelScriptTools` 名字像"脚本工具"，实测是 VBA 的工具函数集合（`FromInches` /
> `FromPoints` / `ASin` / `BuildDate` / `FindFirstFolder`…），**不是宏执行器**。
> 想通过它执行 VBA 宏这条路走不通。

---

## 八、一句话总结

* **能力域覆盖**：42 个能力域 / 27 个工具，日常出图闭环完整；节点编辑、对象样式、专色、
  图案/纹理填充、跨文档复制、模板/样式表/颜色样式、文档元数据、对象数据字段、自定义虚线、
  文档设置、文本区间编辑、自定义箭头、图案微调、文本格式查询、位图内容替换、文档回滚、
  立体化/透镜/效果管理、位图特效命令串、混合、文档颜色转换、用户偏好 KV 都已补上。
* **方法级覆盖**：全局 43.0%（696/1619），主干 20 接口 56.0%（416/743）。
* **明确不可用**：位图描摹、文本导入、点集数组建曲线、尺寸标注/连接线（MathUtils 点对象
  不可创建）、符号库、文件型图案/影线填充（本机无库）、曲线贴笔触、OpenDocumentAsCopy、
  HTML 兼容文本、断字写入、填充字符串、文本尺寸查询、灰度→RGB 转换、真实应用偏好读取。
* **第十九波新增**：`RegisterUserApplicationPreference` + `Get/SetApplicationPreferenceValue`
  提供用户偏好 KV；`ActiveTool/ActiveToolStateGuid` 可读。


