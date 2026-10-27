# API 覆盖清单

> 目的：证明"还有多少没探"。每一行要么有对应的工具，要么有实测失败证据。
> 环境：CorelDRAW Graphics Suite X8 `18.0.0.450`（64 位）/ Python 3.11.9 64 位
> 最近更新：2026-10-28 第九十一波（**易用性**：cdr_inspect 摘要模式 / 列表截断）

## 怎么得出这份清单

1. **反解类型库**（只读 `vgcoreauto.tlb`）→ 692 个类型 / 6789 个方法声明；
   用 pywin32 gencache 展开成可调用签名 → 256 个自动化类 / 2248 个 dispatch 方法。
2. **按接口族逐项实测**：能跑通的写成工具，跑不通的记下异常原文。
3. **每条结论都有出处**：要么是 `scripts/live_smoke.py` 里的断言，要么是本文件"不可用"小节里的错误信息。

---

## 一、主干接口覆盖

| 接口 | 方法数 | 覆盖情况 |
|---|---:|---|
| `IVGApplication` | 82 | 已用：CreateDocument / **CreateDocumentFromTemplate** / OpenDocument / 全套 Create*Color（**第三十五波补齐 HSB/CMY/BW/YIQ/Fixed/CreateColor/CreateColorEx**）/ CreateShapeRange / CreateStructExportOptions / CreateStructImportOptions / **CreateStructSaveAsOptions** / **CreateStructFontProperties** / OutlineStyles（含 **AddStyle/Save**）/ ArrowHeads / Version / Path / Documents / **ColorManager / ActiveTool / ActiveToolStateGuid / RegisterUserApplicationPreference / GetApplicationPreferenceValue / SetApplicationPreferenceValue / CreateOnScreenText / CreateOnScreenCurve / CreateOnScreenHandle** / **FontList / ForceUpdateFontTable / Palettes（含 OpenFixed / Create / Item） / CreateDuotone / CreateSpotColorByName / CreatePatternCanvas / PatternCanvases / GetSupportedOpenTypeFeatures / CheckPlatformVersion / IsUILanguageAvailable / CreateRect / ConvertUnits / ConvertToUnicode / ConvertFromUnicode / CreateProperties**；OpenDocumentAsCopy 参数 TypeError；**第三十八波**：`CreateStructCreateOptions` + `CreateDocumentEx`（一次定名称/页宽/页高/分辨率）、`CreateStructOpenOptions` + `OpenDocumentEx`（打开选项只有 CodePage/ColorConversionOptions）；**第四十一波**：`CreateNodeRange`（配合 `Add` 凑**不连续**节点集合，支持 Move/Rotate/Skew/Stretch/AutoReduce/Smoothen/SetType/Delete/Fillet/Chamfer/Scallop/CreateSelection/GetBoundingBox）、`CreateSegmentRange`（任意线段集合：SetType/AddNode/CreateSelection/RemoveAll/FirstSegment·LastSegment）；**第四十三波**：`CreateStructPaletteOptions`（NumColors/DitherType/PaletteType/Importance/Lightness/TargetColor/Palette）、`CreateStructPasteOptions`（只有 ColorConversionOptions）；**第四十六波**：`CorelScriptTools()` 返回 `ICorelScriptTools`（VBA 辅助对象），`FromInches/ToInches/FromPoints/ToPoints/ASin/BuildDate/FindFirstFolder` 可用；**第五十一波**：`CreateStructAlignProperties` / `CreateStructSpaceProperties` / `CreateStructHyphenationSettings` 三个工厂造出的结构体可直接喂给对应 `Set*Properties`；`ConvertToUnicode` / `ConvertFromUnicode` 实测被 pywin32 的 BSTR 转换弄坏（见不可用清单）；**第五十八波（纠正）**：`CreatePaletteColor(调色板 GUID, 序号)` / `CreateSpotColor(调色板 GUID, 专色ID, Tint)` **是"造色"不是"取色"**（GUID 传对即可）；`CreateColorContext2(配置名串, RenderingIntent, BlendingColorModel)` 与 `CreateProofColorSettings(配置名, RenderingIntent, 保留颜色值)` 也可用 —— **这两个 API 的 RenderingIntent 有效范围是 1~4**（传 0 才报越界）；**第六十二波**：`CreateStructFontProperties()` 可造可读（但 `SetFontProperties` 写入无效）、`OpenDocumentAsCopy(path, 打开选项)` 返回**未命名副本**（FileName 为空、内容一致）|
| `IVGApplicationEvents` | 24 | **新打通（第二十六波）**：通过生成的 `DIVGApplicationEvents` 连接点 sink 订阅；实测捕获 WindowActivate / DocumentNew / DocumentClose / SelectionChange / DocumentBeforeSave / DocumentAfterSave / DocumentBeforeExport / DocumentAfterExport / Start / PluginCommand；`ApplicationEvent` 高频忽略 |
| `IVGDocumentEvents` | 28 | **新打通（第二十七波）**：每个文档挂 `DIVGDocumentEvents` sink；捕获 ShapeCreate/Delete/Move/Transform/Distort/Change、LayerCreate/Delete/Activate/Change、PageCreate/Delete/Activate/Change、SelectionChange、Before/AfterSave、Before/AfterExport |
| `IVGDocument` | 78 | 已用：Unit / Pages / AddPages / DeletePages / ActivePage / ActiveLayer / Selection / Save·SaveAs / ExportEx / ExportBitmap(不适用) / PublishToPDF / Undo / Redo / Begin·EndCommandGroup / FileName / Dirty / Name / StyleSheet / **Metadata / DataFields / ColorContext / ConvertToColorContext / Revert / Views / CreateView** / **设置：EditAcrossLayers / ApplyToDuplicate / CurvePrecision / Resolution / ReferencePoint / ShapeEnumDirection / TextFormatter / WorldScale / FacingPages / FirstPageOnRightSide / Layout / Unit / DrawingOriginX·Y / PreserveSelection** / **Keywords·Notes（可写，读需走 Metadata）** / **ToUnits / FromUnits / SampleColorAtPoint / SampleColorInArea / ShowAllHiddenObjects / SetLayout / Repeat / CreateImage / CreateColorStyles / InsertPages / HatchLibraries / SaveStyleSheet / LoadStyleSheet / CreateFill / CreateOutline（字符串格式无效）**；**第三十五波**：`AddPagesEx` / `InsertPagesEx`（定尺寸加页/插页）、`CreateArrowHead2/Ex/Ex2`（命名箭头）、`ResolveAllBitmapsLinks`、`SaveSettings`/`RestoreSettings`/`ResetSettings`（设置快照，调用无异常、可见效果未确认）、`CreateColor`（在 Application 上）/ `CreateUniformFill`；**第三十八波**：`CreateUniformFill(IVGColor)` 造填充对象 + `ShapeRange.ApplyFill` 批量上色、`CreateArrowHead2(Curve, Name)` 命名箭头；**第四十二波**：`CreateShapeRangeFromArray`、`ConvertToBitmapEx`（范围）、`AssignColorContext`、`SaveStyleSheetAsDefault`、`CreateArrowHeadEx` / `CreateArrowHeadEx2` / `CreateArrowHeadOptions`；**第四十三波**：`Export(FileName, Filter, Range, Options, PaletteOptions)`、`ExportBitmap(16 参，返回 ICorelExportFilter，必须 Finish())`；**第五十一波**：`CreateArrowHeadEx`（无名箭头，Name 为空）与已用的 Ex2 并存（`arrowhead_unnamed`）；**第五十三波**：`CreateBSpline(控制点数, 闭合)` → `IVGBSpline`（配合 `InsertControlPoint(序号, x, y, Clamped)`，**两端必须 Clamped=True**）；**第五十九波**：`CreateCurveFitToPoints(IVGPointRange, False, 容差)` / `CreateCurveFitToPointsAndCusps(点集, 尖点数组, False, 容差)` —— 点集拟合曲线的入口 |
| `IVGPage` | 22 | **全覆盖**：SetSize / SelectSize / GetSize / CreateLayer / Activate / Delete / MoveTo / FindShape(s) / TextFind / TextReplace / GetGuides / UnlockAllShapes / GetObjectsBoundingBox(改用自算) / SelectShapesFromRectangle / Shapes / Layers / Orientation / **FindClosestSnapPoint** |
| `IVGLayer` | 50 | 已用：Create*（矩形/椭圆/多边形/螺旋/网格/线/贝塞尔/艺术字/段落字/参考线/**角度参考线**/**位图 CreateBitmap·CreateBitmap2**/**OLE 文件 CreateOLEObjectFromFile**/**矩形造形 CreateRectangleRect·CreateEllipseRect**（第三十四波）/**曲线段 CreateCurveSegment**（第三十四波））/ Import / ImportEx / CreateCurve / Activate / Delete / MoveAbove / **MoveBelow / GetAbove / GetBelow / Paste / FindShape** / Visible / Editable / Printable / Master / **IsGuidesLayer（属性，不是方法）**；**第三十八波**：完整参数形态 `CreateRectangle(left,top,right,bottom,4 圆角)` / `CreateEllipse(left,top,right,bottom,起止角,pie)` / `CreatePolygon(left,top,right,bottom,sides,subpaths,complexity,star,starComplexity,maxComplexity)`；**第四十一波**：`CreateBitmap2(x,y,w,h,Image,Alpha)` 与 `CreateBitmapRect(IVGRect,Image,Alpha)`（都要 IVGImage + Alpha=None）；**第四十三波**：`PasteEx(IVGStructPasteOptions)`（返回 ShapeRange）；**第四十七波**：`CreateArtisticTextWide` / `CreateParagraphTextWide`（显式语言/字符集）、`ImportEx(FileName, Filter, Options)`；**第五十三波**：`CreateBSpline(IVGBSpline)` —— 拿到 `Document.CreateBSpline` 造好的样条定义后落成 **Type 21** 的样条对象（`ConvertToCurves` 后变 Type 3）；**第五十五波**：`CreateConnector(起, 止)` / `CreateConnectorEx(类型, 起, 止)` / `CreateRightAngleConnector(起, 止, 圆角)` → **Type 19** 连接器；`CreateLinearDimension(22 参)` / `CreateAngularDimension(18 参)` → **Type 14** 标注（端点用 `Application.CreateSnapPoint(x, y)` 造，返回 IVGSnapPoint）|
| `IVGShape` | 124 | 已用：几何 / Fill / Outline / Transparency / Locked / Name / Type / GetBoundingBox / Duplicate / Clone / Delete / **层级 / StepAndRepeat / Distribute / CreateBoundary / CreatePerspective / Separate / CreateEnvelopeFromCurve / CreateEnvelopeFromShape / CompareTo / CreateSelection / RemoveFromSelection / IsOnShape / Copy / Cut / CopyToLayer / MoveToLayer / Project / Unproject / CreateDocumentFrom / CreateExtrude / CreateLens / CreateBlend / OverprintFill / OverprintOutline / OverprintBitmap** / ConvertToCurves / ConvertToSymbol / Weld·Trim·Intersect / AddToPowerClip / CreateDropShadow·Contour·Envelope / AlignToPage·AlignToShape / Fillet·Chamfer·Scallop / Group / Ungroup / BreakApart / Combine；**第三十九波**：`GetMatrix` / `SetMatrix`（[d11,d12,d21,d22,tx,ty]，读回精确；tx/ty 是 X8 内部单位）、`IsTypeAnyOf`（传 list[int] 可用）、`IsOnShape`（2=内部 / 1=边界 / 0=外部）、`OrderIsInFrontOf`、`PlaceTextInside`、`DuplicateAsRange` / `CloneAsRange`、**扭曲三件套 `CreatePushPullDistortion` / `CreateTwisterDistortion` / `CreateZipperDistortion`**；**第四十波**：`GetPositionEx` / `SetPositionEx`（cdrReferencePoint 1~9 定位）、`FindSnapPoint(ReferenceData)`、`GetLinkedShapes(LinkType)`、`ConvertToBitmapEx(8 参)`、`GetOverprintFillState` / `GetOverprintOutlineState`；**第四十八波**：`CopyToLayerAsRange(layer)`（复制到指定图层并返回范围）、`AlignToShapeRange(alignType, range, textOrigin)`（对齐到另一批对象）；**第四十九波**：`CopyPropertiesFrom(源对象, 掩码)` 在**形状级**可用（掩码 4/8/15 会复制填充，1/2 不影响填充；范围级同名方法仍报错）；**第五十波**：`SkewEx(angleX, angleY, cx, cy)` 形状级斜切可用（`ApplyTransformMatrix` 仍不可用：`Application.MathUtils` 不存在，_造型不了矩阵对象_）；**第五十四波（纠正）**：`GetNext(方向)` / `GetPrevious(方向)` **可用**（各取 1 个参数，返回同层相邻 IVGShape；之前记的「`Shape.Next` 是 3 参 propget」说的是另一个属性）；`Ungroup` 也确认生效；**第六十一波（纠正）**：`SetSizeEx(中心X, 中心Y, 宽, 高)` 与 `StretchEx(中心X, 中心Y, X倍数, Y倍数, 是否缩放文字)` **都可用** —— 之前判"会放大到离谱尺寸"是因为不知道这两个参数是**中心点**，实测 20×10 的矩形按中心 (10,5) 定尺寸 40×30 → bbox (-10,-10,40,30) |
| `IVGShapeRange` | 131 | 已用：Group / Combine / AlignAndDistribute / ApplyFountainFill / SetOutlineProperties(不适用，改属性直写) / ConvertToCurves / Add / Count / **CreateParallelCurves / ConvertToBitmap / ConvertOutlineToObject / CreateBoundary / SetFillMode / SetPixelAlignedRendering / ApplyFill / ApplyOutline / CreateDocumentFrom**；**第三十三波**：`AlignRangeToPage` / **PageCenter / Grid / Point / Shape / ShapeRange**、`AddRange` / `RemoveRange`、`Unlock`、`BreakApartEx`、`ClearEffect`、`IndexOf` / `CountAnyOfType` / `ExistsAnyOfType` / `FindAnyOfType`、`GetPositionEx` / `SetPositionEx` / `RotateEx` / `SkewEx`；**空操作/不可用**：`UngroupAll`·`UngroupAllEx`·`UngroupEx`（实测调完对象数不变）、`Sort`（比较表达式语法未逆出）、`AllExcluding`（TypeError）、`SetSizeEx`/`StretchEx`（参数语义不明且会放大到离谱尺寸）；**第四十二波**：`Document.CreateShapeRangeFromArray([shape,…])` **接受 Python list**（配 `Exists` 校验），`ConvertToBitmapEx(8 参)` 整批位图化；**第四十七波**：`SetOutlineProperties`（14 参）/ `SetOutlinePropertiesEx`（15 参，含 Justification）可与 `Shape.Outline.SetProperties(Ex)` 一样全参调用；`AlignRangeToPage(cdrAlignType)` 直接对齐页边；`DeleteItem(i)` 会**连文档里的对象一起删**（已记入陷阱）；**第五十三波**：`ApplyOutline(IVGOutline)` —— 用一个描边对象整批覆盖范围内所有对象（`Outline.GetCopy()` 再改宽度即可）；**第五十四波（重要纠正）**：`UngroupAll()` / `UngroupAllEx()` / `UngroupEx()` **三个都真实生效**（之前的「空操作」结论是测量方式错了）——`UngroupAll` 无返回，`UngroupAllEx`/`UngroupEx` 返回 ShapeRange；实测组合体解开后页面对象数 1→2/1→3；**第六十一波**：`AllExcluding(索引数组)` **接受 Python list**、索引 **1 基**，返回排除后的新 ShapeRange；`SetSizeEx` / `StretchEx` 在范围上也同样可用（`cdr_shape_ops` 的 `use_range=true`）|
| `IVGFill` | 13 | 已用：ApplyUniformFill / ApplyFountainFill / ApplyTextureFill / **ApplyPatternFill（内存 PatternCanvas 可用；文件型 Pattern 仍需图像文件）** / ApplyPostscriptFill / ApplyNoFill / **GetCopy / CompareWith**；Hatch 需本机 Hatch 库（本机为空）；**第六十波（纠正）**：`ApplyHatchFill(库名, HatchNameOrIndex, 背景色, …)` 与 `ApplyCustomHatchFill`（18 参）**都可用** —— 之前判不可用是因为`HatchNameOrIndex` 传了字符串，**它必须是 int 序号**；`Document.HatchLibraries` 仍报错但填充本身不依赖它 |
| `IVGOutline` | 15 | **第三十二波补齐（13/15）**：Width / Color / **Style（可赋自定义线型）** / StartArrow·EndArrow / LineCaps / LineJoin / BehindFill / ScaleWithShape / SetNoOutline / **NibStretch / NibAngle / MiterLimit / Justification**（可读写）/ **GetCopy / CopyAssign / CompareWith / ConvertToObject / SetProperties（14 参全实参可用）**；`PenWidth` 恒 0（派生）、`DashDotLength` 写入不生效、`ToString/StringAssign` 在 pywin32 包装类里不存在；**第四十五波**：`SetPropertiesEx`（15 参，多一个 Justification）实测可用 |
| `IVGOutlineStyle` / `IVGOutlineStyles` | 6 / 6 | **新打通（第十一波）**：`CreateOutlineStyle(dash,gap[,dash,gap,…])` 生成虚拟线型（Index=-1），`OutlineStyles.AddStyle(virtual)` 注册后得到可用 Index，赋给 `Outline.Style` 即生效；`Save()` 可持久化 |
| `IVGDataFields` / `IVGDataField` / `IVGDataItems` | 8 / 18 / 4 | **新打通（第十一波）**：`DataFields.Add(Name,"0")`、`IsPresent/Item`、`Shape.ObjectData.Add(DataField, Value)`、`DataItem.Value/FormattedValue` 可读写（对象级持久化数据字段） |
| `IVGCurve` | 37 | 已用：Length / Area / SubPaths / Contour / CreateSubPath / GetCopy / **Nodes / Segments / Selection / FindNodeAtPoint / FindSegmentAtPoint / ReverseDirection / JoinTouchingSubpaths / RemoveOverlaps / GetPolyline / IsOnCurve / IsClockwise / Closed / GetCurveInfo**；数组入参（AppendSubpathFromPoints / CreateSubPathFromArray）不可用；CreateCurveMappedToStroke 类型不匹配；**第三十六波**：`AppendCurve`（把另一条曲线并成子路径）、`WeldWith`（返回新曲线）、`IsOnCurve(x,y,hot)`、`GetCopy`、`GetCurveInfo`（返回 com_struct 数组）；**第五十波**：`BindToDocument(doc)` —— `WeldWith` 返回的游离曲线必须先 Bind 再 `Layer.CreateCurve`，才能变回文档里的真形状（补上了焊接流程最后一环）；**第五十六波（纠正）**：`CreateCurveMappedToStroke(子路径, ScaleY, SelfWeld)` 与 `CreateCurveMappedToStrokeAndReferenceLine(子路径, IVGPoint, IVGPoint, ScaleY, SelfWeld)` **都可用** —— 之前判"类型不匹配"是因为**把整条 Curve 当骨架线传**，正确类型是 `IVGSubPath`；两者都返回新 Curve，`Layer.CreateCurve` 后成 Type 3 形状；**第五十九波（纠正）**：`AppendSubpathFromPoints(点集, 闭合)` / `AppendSubpathFitToPoints(点集, 缩放, 容差)` / `AppendSubpathFitToPointsAndCusps(点集, 尖点, 缩放, 容差)` **都能用** —— 之前判"数组入参不可用"是因为直接塞 Python list；正确做法是先拿到 **`IVGPointRange`**（`SubPath.GetEvenlySpacedPoints(间距, 缩放)` 或 `AddPointXY` 逐点补）再传 |
| `IVGSubPath` | 39 | 已用：CreateSubPath / AppendLineSegment / AppendCurveSegment2 / Closed / **Nodes / Segments / StartNode / EndNode / GetPointAt / GetPointPositionAt / GetTangentVectorAt / GetPerpendicularVectorAt / GetCurveSpeedAt / GetIntersections / ReverseDirection / Next / Previous / Extract / GetCopy**；求交/等分/切向已打通，点集数组仍取不回；`Next`·`Previous`·`GetCopy` 与曲线同名方法共享计数（审计按名字匹配，含跨接口误计）；**第五十波**：`AppendCurveSegment(x,y,起控长,起控角,终控长,终控角,atBegin)`（节点 +1、返回 IVGSegment）、`Extract(Shape)`（**把子路径抽成独立对象**，返回 `(新 Shape, …)`）；**第五十二波**：`GetPointPositionAt(0,0,偏移,偏移类型)` 返回 `(X,Y)`（类型 0=绝对长度 mm、1=归一化 0~1）、`FindSegmentOffset(绝对偏移, Segment, ParamOffset, Remainder)` 返回 `(bool, IVGSegment, 参数偏移, 余数)` |
| `IVGNode` | 19 | **已覆盖（第五波纠正）**：`curve.Nodes` / `curve.SubPaths.Item(i).Nodes` 是**属性**；`GetPosition/SetPosition/Move/Type/Index/SubPathIndex/IsEnding/Selected/Delete/CreateSelection` 均可用；范围操作走 `curve.Nodes.Range(a,b)` → `IVGNodeRange`（Fillet/Chamfer/Scallop/Smoothen/SetType/Move/Rotate/Skew/Stretch/BreakApart/GetBoundingBox）；`app.CreateNodeRange()` 返回空集合；**第三十六波**：`Next()` / `Previous()`（**是方法不是属性**，返回 IVGNode）、`GetDistanceFrom`、`AveragePositionWith`、`JoinWith`（末端节点合并）、`GetPoint`（返回 IVGPoint）；`ConnectWith` / `ExtendSubPaths` 实测报「指定的节点必须在另一条曲线上」/「无法连接曲线」；**第四十七波**：`GetPoint()`（返回 IVGPoint）+ `SetPoint(point)` 可整颗搬运节点坐标 |
| `IVGText` | 30 | 已用：Story（**Font / Size / Bold / Italic / Underline 读取**）/ GetSpaceProperties / SetSpaceProperties / GetAlignProperties / SetAlignProperties / FitToPath / FitTextToFrame / **GetContents / Range / ExportToFile / GetHyphenationSettings**；GetFontProperties 结构体返回默认值不可信；**第三十七波**：`FontProperties` / `SpaceProperties` / `AlignProperties` / `HyphenationSettings` 四个**属性**可读（返回 IVGStruct* 快照）；带下标/帧的 `*InRange` 版本返回的结构体字段是垃圾值（Size=1.17e-38），不可用；直接赋值 setter 会把字号写坏，**只读不写**；**第四十五波**：`BeginEdit()` 可用（进入文字编辑态）；**第五十一波**：`SetAlignProperties` / `SetSpaceProperties` / `SetHyphenationSettings` 配合结构体工厂可从零构造段落格式（**只对段落文本生效**，艺术字报「该参数需要段落文本」）；**第五十七波（重要纠正）**：`FontPropertiesInRange` / `SpacePropertiesInRange` / `AlignPropertiesInRange` / `HyphenationSettingsInRange` **都能读**，`SetAlignPropertiesInRange` / `SetSpacePropertiesInRange` / `SetHyphenationSettingsInRange` **都能写** —— 之前判"垃圾值"是两个错：**方法名没有 `Get` 前缀**、且 **StartIndex 是 1 基**（传 0 才返回 Size=0/Style=16 的垃圾）|
| `IVGTextRange` | 35 | 已用：Text / Font / Size / Bold / Italic / Underline / ChangeCase / ApplyBulletEffect / ApplyDropCapEffect / ApplyNoEffect / InsertBefore / InsertAfter / **SetOpenTypeFeature / GetOpenTypeFeature / Delete / Copy / Paste / Straighten / FindRanges / EnumRanges / LeftIndent / RightIndent / FirstLineIndent / Tabs**；ApplyStyle 调用不报错但未观测到效果；**第三十七波**：`InsertBeforeWide` / `InsertAfterWide` / `ReplaceWide`（带语言/字符集/字体）、`Collapse` / `SetRange` / `InRange` / `IsSame` / `SetLineSpacing`、`Start`·`End`；`Evaluate` 的表达式语法未逆出（`len` 报 Identifier is not a method of global）|
| `IVGTextTabPositions` | 9 | **新打通（第二十三波）**：`TextRange.Tabs` 的 Count/Item(i)/Add(Position,Alignment,Leadered)/AddEvery/Clear；`TextTabPosition` 的 Position/Alignment/Leadered 可读写 |
| `IVGColor` | 44 | **第三十二波补齐（32/43）**：仍不暴露分量属性，但对象级能力全通 —— ToString / GetName·SetName / StringAssign / IsSame·IsSimilar / GetColorDistanceFrom / BlendWith（原地）/ GetCopy / ConvertToRGB·CMYK·Gray·HLS·HSB·Lab·CMY·YIQ（原地）/ ConvertToFixed(0=USER,1=TRU,2=PAN_PROC) / RGBAssign·CMYKAssign·GrayAssign·HLSAssign·HSBAssign·LabAssign·CMYAssign·YIQAssign·BWAssign·RegistrationAssign；不可用：ConvertToPalette / SpotAssignByName / CopyAppearance / ModifyColorStyleColor / UserAssign(Ex)（参数是窗口句柄，会弹对话框）；**第四十八波**：`FixedAssign(palette_id, index, tint)`、`CorelScriptAssign(模型号, v1..v7)`、`PaletteAssign(调色板 Identifier, 序号)`、`Invalidate()` 全部可用（PANTONE 调色板取色得到 `PANTONE,PAN_SPOT,5,100,<guid>`）；`SpotAssign/SpotAssignByName` 需要**含专色的**调色板，否则报「未找到指定的对象」；**第四十九波（重要纠正）**：`CorelScriptGetComponent(模型号, 0×7)` **能读回颜色分量** —— CMYK 返回 `(5002,10,20,30,40,0,0,0)`、RGB `(5005,11,22,33,…)`、Gray `(5009,50,…)`，状态码 = 5000 + 模型号；`ConvertToPalette(调色板 GUID)` 也可用（RGB(120,60,30) + TRUMONT → `CMYK255,TRU,…`，PANTONE → `PAN_SPOT,619`）|
| `IVGBitmap` | 16 | 已用：ConvertTo / ConvertToBW / **ConvertToPaletted（0~9） / ConvertToDuotone** / Resample / Crop / ResetCropEnvelope / Inflate / Mode / ResolutionX / Size / Image / Duotone / **SetImageData（Alpha 必须传 None） / ApplyBitmapEffect（命令串） / ExternallyLinked / LinkFileName / UpdateLink / ResolveLink**；Trace 禁用（见下）；**第四十三波**：`ConvertToPaletted2(IVGStructPaletteOptions)`（Mode 4→3，可挂已打开的调色板）|
| `IVGFountainFill` | 7 | 通过 ApplyFountainFill 设置，未做二次编辑；**第四十四波**：`GetTransformations()` 用 4 个出参返回 `[a,b,c,d]` 变换矩阵 |
| `IVGTransparency` | 9 | 已用：ApplyUniformTransparency / ApplyFountainTransparency / **ApplyPatternTransparency（PatternCanvas 可用）** / ApplyNoTransparency / **Freeze / Unfreeze / Frozen**；**第四十四波**：`ApplyTextureTransparency(名称, 库, ?, ?)` 返回 IVGTextureFill，`Transparency.Type` 变 4 |
| `IVGStyle` / `IVGStyleSheet` | 16 / 16 | **已覆盖（第六/七/九波）**：`CreateStyleFromShape`、`CreateStyle(category, based_on, name)`、`Styles.Find/Item/Count`、`Style.Name/CategoryName/DisplayName/BasedOn`、`Rename / Delete / GetAllPropertyNames / GetPropertyAsString / SetPropertyAsString / SetProperty / DerivedStyles / ToString`、`Shape.ApplyStyle("样式名")`；**`Document.SaveStyleSheet/LoadStyleSheet` 可导出/恢复 .cdst**；**`Document.CreateColorStyles(harmony>0)` 会生成颜色样式（AllColorStyles）**；`fill`/`outline` 分组名不可直接读写；**第三十七波**：`GetProperty`（返回 IVGColor）、`GetOverridePropertyNames`、`IsPropertyInherited`、`ClearProperty`；`SetBasedOn(None)` 报「未找到样式 None」、`Assign(shape)` 报类型不匹配 |
| `IVGWindow` | 9 | 已用：Refresh / Activate / DocumentToScreen / ScreenToDocument / DocumentDistanceToScreenDistance / **ScreenDistanceToDocumentDistance**；缩放/新建窗口未封装；**第四十四波**：`NewWindow()` 可开新窗口（`Application.ActiveWindow.NewWindow()`，文档上没有此方法）|
| `IVGView` / `IVGViews` | 8 / 5 | **新打通（第二十波）**：`Document.CreateView(Name, OriginX, OriginY, Zoom, Page)` 创建命名视图；`Views.Count/Item(i)` 列出；`View.Activate/Delete`、Name/OriginX·Y/Zoom/Page/UsePage/UseZoom 可读写 |
| `IVGOnScreenText` / `IVGOnScreenCurve` / `IVGOnScreenHandle` | 8 / 12 / 6 | **新打通（第二十五波）**：`Application.CreateOnScreenText/Curve/Handle` 创建屏幕叠加层；Text 的 SetTextAndPosition/SetTextColor/Show/Hide，Curve 的 SetPen/SetNoBrush/SetRectangle/SetLine/SetCircle/Show/Hide，Handle 的 SetPosition/SetHandleColor/IsHotTracked |
| `IVGGrid` | 5 | **新打通（第二十一波）**：`Document.Grid` 的 Type / SpacingX / SpacingY / Visible / Snap 全部可读写（间距按文档单位，mm 上下文下为毫米） |
| `IVGColors`（颜色样式集合） | 8 | **新打通（第九波）**：`StyleSheet.AllColorStyles` 的 `Count / Item(i) / GetName / ToString`；配合 `Document.CreateColorStyles(harmonies=1..3)` 生成颜色协调 |
| `IVGMetadata` | 21 | **新打通（第十波）**：`Document.Metadata` 的 Title / Author / LastAuthor / Subject / Keywords / Notes / Copyright / Revision / DocLanguage / DocID；前六项可写；Localizable* 是 `IVGLocalizableString` 对象 |
| `IVGStructExportOptions` | 22 | **新打通（第二十八波）**：AntiAliasingType / Compression / Dithered / Transparent / MaintainAspect / MaintainLayers / Overwrite / Matted / MatteColor / ResolutionX·Y / SizeX·Y / UseColorProfile / AlwaysOverprintBlack 可写；`cdr_export(export_options={...})` 应用 |
| `IVGStructSaveAsOptions` | 10 | **新打通（第二十九波）**：EmbedICCProfile / EmbedVBAProject / IncludeCMXData / KeepAppearance / Overwrite / ThumbnailSize / Version / Filter / Range 可写；`cdr_save(save_options={...})` 应用 |
| `IVGStructImportOptions` | 24 | **新打通（第三十波）**：LinkBitmapExternally / MaintainLayers / ResampleDpiX·Y / ResampleWidth·Height / Crop* / DetectWatermark / ExtractICCProfile / ICCFileName / UseOPILinks / ConvertTablesToText / ForceCMYKBlackText / TextFormatting / ImageIndex / Mode / CodePage 可写；`cdr_import(import_options={...})` 应用 |
| `IVGLocalizableString` | 10 | **新打通（第二十四波）**：DefaultLangString / HasDefaultLangString / IsEmpty 可读写；GetLanguages / GetLangString(lang) / SetLangString(lang,text) 支持多语言；LocalizableTitle/Subject/Keywords/Notes/Copyright 各一份 |
| `IVGEffects` / `IVGEffect` | 12 / 12 | **新打通（第十五波）**：`Shape.Effects.Count/Item(i)`、`ExtrudeEffect/LensEffect/DropShadowEffect/BlendEffects` 等；`Effect.Type/Extrude/Lens/Separate/Clear` 可用 |
| `IVGEffectExtrude` | 18 | **新打通（第十五波）**：Type/Depth/AngleX·Y·Z/Shading/BaseColor/ShadingColor/UseBevel/BevelDepth/BevelAngle/ShowBevelOnly/UseExtrudeColorForBevel/UseFullColorRange/VanishingPoint 可读写；`Rotate/SetBevel/SetLight` 未封装 |
| `IVGEffectLens` | 18 | **新打通（第十五波）**：Type/Rate/Magnification/Color/FillColor/OutlineColor/FromColor/ToColor/UseFillColor/UseOutlineColor/PaletteRotation/ColorMapPalette/RemoveFace/ViewPointX·Y/UseViewPoint/Frozen 可读写；`Freeze/Unfreeze/Ungroup` 未封装 |
| `IVGEffectBlend` | 31 | **新打通（第十七波）**：Steps / Spacing / Angle / ColorBlendType / Mode / Loop / RotateShapes / MapNodes / FuseStart / FuseEnd / ColorAcceleration / SpacingAcceleration / LinkAcceleration / AccelerateSize / StartPoint / EndPoint / StartShape / EndShape / Path 可读写；`Split()` 可用 |
| `IVGArrowHeads` | 7 | **新打通（第十二波）**：`Application.ArrowHeads.Add(自定义箭头)` 把 `CreateArrowHead(curve)` 的结果登记进全局箭头库（89→90），可赋给 `Outline.StartArrow/EndArrow` |
| `IVGColorContext` | 7 | **已读（第十一波）**：`Document.ColorContext` 的 RenderingIntent / BlendingColorModel / ColorProfiles.Count / ColorProfileNameList；**第十八波**：`profile.CreateColorContext()` + `Document.ConvertToColorContext()` 可做文档颜色转换；直接 `Application.CreateColorContext()` 参数不可传；**第四十四波**：`GetCopy` / `CopyAssign` / `Merge` / `IsSame` / `ColorProfile(model)`（get/put）/ `RGBColorProfile`·`CMYKColorProfile`·`GrayscaleColorProfile` / `ColorProfileNameList` / `ReadOnly` 全部可用；**第五十一波**：`ColorProfile(模型号)` 0=RGB / 1=CMYK / 2=灰度 可逐模型取配置文件 |
| `IVGColorManager` | 31 | **新打通（第十八波）**：`Application.ColorManager` 的 ColorProfiles（本机 64 个）/ GetProfilesByColorModel(0=RGB,1=CMYK,2=Gray,3=Lab) / DefaultColorContext；用于选配置文件并转文档颜色 |
| `IVGColorProfile` | 10 | **新打通（第十八波）**：Name / ColorModel / FileName / Installed / Generic / Manufacturer 可读，`CreateColorContext()` 得到该配置文件的上下文 |
| `IVGImage` | 5 | **新打通**：`Document.CreateImage` 返回像素缓冲区；`Width / Height / Type / ReadOnly / Pixel / FillArea / Blit` 可用（`cdr_create(type="bitmap")`、`cdr_bitmap(action="pixel")`） |
| `IVGDuotone` / `IVGDuotoneInk` | 8 / 8 | **新打通**：`app.CreateDuotone()` → `Type 0~3`、`Inks(i).Color`、`Load(.cpd)`；配合 `Bitmap.ConvertToDuotone` 转双色调 |
| `IVGPalettes` / `IVGPalette` | 8 / 15 | **新打通**：`Palettes.OpenFixed(id)` 可打开 PANTONE 匹配系统（946 色）等内置调色板；`Create/AddColor/SaveAs/Color/ColorCount/Identifier/FindColor` 可用；`CreatePaletteColor(Identifier GUID, index)`、`CreateSpotColorByName(Identifier, 名称, tint)` 可用 |
| `IVGFontList` | 3 | **新打通**：`app.FontList.Count / Item(i)` 返回字体名字符串（本机 2068 个） |
| `IVGSnapPoint` / `IVGSnapPoints` / `IVGSnapPointRange` | 24 / 17 / 14 | **新打通（第三十一波）**：吸附点全家族。`Shape.SnapPoints` 枚举（普通矢量对象 4 个角/边中点，`SnapPointsOfType(掩码)` 过滤）；`Auto()` / `BBox(1..9)` / `Object(1..4)` / `Edge(段>=1, offset)` / `User(GUID)` / `FindClosest(类型,x,y)` 定址取点；`AddUserSnapPoint(x,y,direction,use_direction)` 新增用户点；`Application.CreateSnapPoint` / `Document.CreateFreeSnapPoint` 建自由点；`Move/SetPosition/Delete/CreateSelection/Direction/UsesDirection` 就地编辑；`SnapPoints.Range([ReferenceData])` → `IVGSnapPointRange` 做批量 Move/Delete/Add/Remove/RemoveByReference/Find/CreateSelection/AddToSelection/RemoveFromSelection/ChangeDirection/SetAutoSnap。`AddUserSnapPointEx` 任何 ID 形态都报「为参数 ID 指定的值无效」 |
| `IVGUserSnapPoint` / `IVGObjectSnapPoint` / `IVGBBoxSnapPoint` / `IVGEdgeSnapPoint` | 2 / 1 / 1 / 2 | **新打通（第三十一波）**：`SnapPoint.User` → ID(GUID)/AutoSnap（可读写）；`SnapPoint.Object.Type`、`SnapPoint.BBox.Type`、`SnapPoint.Edge.SegmentIndex/SegmentOffset` 全部可读 |
| `IVGRect` | 17 | **新打通（第三十四波）**：`Application.CreateRect()` + `SetRect(x,y,w,h)`；可写 `x`·`y`·`Width`·`Height`，只读 `Left`·`Right`·`Top`·`Bottom`·`CenterX`·`CenterY`、`IsEmpty`；`Offset` / `Inflate(l,t,r,b)` / `IsPointInside` / `Union` / `Intersect` / `GetCopy` / `CopyAssign` / `Clear` / `GetRect` 全部可用（注意 **`Left/Top` 没有 setter**）|
| `IVGText` 的粗斜下划线 | — | **第三十四波修正**：Bold/Italic 形参是 `cdrTriState`（**必须传 -1**，传 True/1 会被忽略）；Underline 是样式枚举（**1=有、0=无，-1 等于无**）。修正前 `cdr_create(bold=true)` / `cdr_text(bold=true)` 全部静默失效 |
| `IVGPatternCanvas` / `IVGPatternCanvases` | 13 / 8 | **新打通（第八/十三波）**：`CreatePatternCanvas()` 生成 32×32 单色画布；`Clear / FillArea / SetPixel / Line / CopyArea / RotateArea / FlipArea / PutCopy / Select / Data` 可用；`PatternCanvases.Add(canvas)` 登记后 `Fill.ApplyPatternFill(0,"",index,front,end,transform)` 即可图案填充，`Transparency.ApplyPatternTransparency(...)` 图案透明 |
| `IVGPatternFill` | 18 | **新打通（第十三波）**：已有图案填充可微调 `FrontColor/BackColor/TileWidth/TileHeight/OriginX·Y/RotationAngle/SkewAngle/MirrorFill/MirrorFillX·Y/TransformWithShape/TileOffset/Canvas/Type`（全部可写且回读正确） |

---

## 二、能力域 → 工具（30 个）

| 能力域 | 工具 | 状态 |
|---|---|---|
| 会话/文档/状态 | `cdr_status` `cdr_connect` `cdr_documents` `cdr_open`（含**从模板创建**） `cdr_inspect` `cdr_save`（含**保存选项**） `cdr_close` | ✅ / **`cdr_open(create_options=…)`（CreateDocumentEx）** / **`cdr_open(open_options=…)`（OpenDocumentEx）** | ✅ / **`cdr_doc(action="assign_color_context")`** / **`cdr_doc(action="stylesheet_default")`** / **`cdr_doc(action="arrowhead_ex")`** | ✅ |
| 绘图 | `cdr_create`（矩形含圆角 / 椭圆含圆弧扇形 / 多边形 / 星形 / 螺旋 / 网格 / 直线 / 贝塞尔 / 文字 / 段落文字 / 参考线 / 任意折线 / **内存位图（纯色 + 像素色块）** / **OLE 文件嵌入**） | ✅ / **宽文本（text_wide / paragraph_wide：显式 language_id·charset）** / **B 样条（bspline）** / **连接器（connector / connector_ex / right_angle_connector）** / **线性标注（linear_dimension）·角度标注（angular_dimension）**） | ✅ |
| 样式 | `cdr_modify`（纯色/渐变/纹理/图案（画布）/专色/无色、描边含虚线箭头/自定义虚线线型、投影、轮廓图、封套、透明/图案透明、冻结/解冻透明度、**图案填充微调**） | ✅ |
| 变换 | `cdr_modify`（移动/尺寸/旋转/翻转/斜切/矩阵/克隆/层级/**层级相对调整**/**清除变换**/**旋转中心**/复制/删除/锁定） | ✅ |
| **对象样式** | `cdr_style`（列出/从对象创建/派生/套用/改名/删除/读定义/改属性/**样式表 .cdst 导入导出**/**颜色样式与颜色协调**） | ✅ 第六/七/九波/**覆盖属性（override_props）·继承状态（inherited）·属性取值（get_property）·属性清除（clear_property）** | ✅ 第六/七/九/三十七波 |
| **节点/线段** | `cdr_node`（列节点/列线段/移动/选中/删除/圆角/倒角/scallop/平滑/拉伸/旋转/斜切/断开/加节点/断线/控制点/按坐标查找/**前后邻居（Next·Previous）/节点间距/平均位置/末端节点合并**） | ✅ 第五/三十六波/**任意节点集合（arbitrary：CreateNodeRange + Add，支持 move/rotate/skew/stretch/autoreduce/smoothen/set_type/delete/fillet/chamfer/scallop/select）** | ✅ |
| 文字排版 | `cdr_text`（字体字号粗斜体下划线 / 对齐 / 大小写 / 项目符号 / 首字下沉 / 行距字距 / 段落缩进 / **制表位** / 前后插入 / 沿路径 / 自适应 / 转曲 / 读取全文·区间·导出纯文本 / 字体属性读取·断字设置读取·格式段列表 / 区间删除·复制·粘贴·拉直 / 按格式查询区间（@font/@bold…） / OpenType 特性） | ✅ / **结构体快照（Font·Space·Align·Hyphenation）** / **区间宽接口（InsertBeforeWide·AfterWide·ReplaceWide·Collapse·SetRange·InRange·IsSame·SetLineSpacing）** / **结构体工厂写入段落格式（paragraph_format：对齐·缩进·字距·行距·连字符，支持 range_start 走 *InRange 区间）** / **区间结构体读取（range_struct：字体·间距·对齐·连字符，1 基）** | ✅ |
| 布尔与整形 | `cdr_shape_ops`（weld/trim/intersect/combine/group/ungroup/break_apart/to_curves/to_symbol/powerclip/fillet/chamfer/scallop/step_repeat/distribute/boundary/perspective/separate/parallel_curves/rasterize/batch_outline_to_curves/fill_mode/pixel_aligned/range_boundary/envelope_from_shape/compare/project/unproject/to_document/extrude/lens/extrude_set/lens_set/effect_info/clear_effects/**blend/blend_set**/**unlock_all/ungroup_all/break_apart_ex/clear_effect/range_info/range_add/range_remove**） | ✅/**矩阵（matrix）·类型集合（is_type）·层序（in_front_of）·命中（is_on_shape）·扭曲（distort：pushpull/twister/zipper）·文字内嵌（place_text）·范围复制（duplicate_range/clone_range）**）| ✅/**参考点定位（position_ex）·吸附反查（find_snap）·链接对象（linked）·带参位图化（rasterize_ex）·叠印状态（overprint_state）**）| ✅/**range_from_list（对象列表建范围，可再 then=rotate_ex）/ range_bitmap（整批位图化）**）| ✅/**copy_to_layer_range（复制到其它图层）/ align_to_range（对齐到形状范围）**）| ✅/**copy_props（按掩码复制属性）**）| ✅/**skew_ex（形状级斜切）/ set_size_ex（按中心定尺寸）/ stretch_ex（按中心拉伸）/ all_excluding（范围排除）/ range_outline_apply（整批套描边）/ ungroup_all·ungroup_all_ex·ungroup_ex（解组家族）/ sibling（同层相邻遍历）**）| ✅ |
| 对齐 | `cdr_align`（左中右/上下中/等间距；`to=selection/page` 走 AlignAndDistribute；**`to=page_center/grid/point/shape/range` 走原生 AlignRangeTo***）+ `cdr_modify(align_to_shape/align_to_page)` | ✅ 第三十三波 |
| 选择与查找 | `cdr_select`（加入/**移除/清空**）`cdr_find` | ✅ |
| 图层 | `cdr_layers`（增删改/显隐/锁定/打印/移动/**上下相邻查询·下移·粘贴剪贴板**） | ✅ / **高级粘贴 paste_ex（CreateStructPasteOptions + PasteEx）** | ✅ |
| **跨文档 / 对象到图层** | `cdr_modify(copy_clipboard/cut_clipboard)` + `cdr_layers(action="paste")`；`cdr_modify(move_to_layer/copy_to_layer)` | ✅ 第八波 |
| 页面 | `cdr_pages`（列表/增删/**定尺寸加页（AddPagesEx）**/**定尺寸插页（InsertPagesEx）**/插入空白页/切页/自定义尺寸/纸张预设/参考线/角度参考线/**吸附点查询**/解锁/包围盒） | ✅ |
| **吸附点** | `cdr_snap`（枚举/类型掩码过滤/自动·包围盒·对象·边定址/新增用户点（GUID）/文档·应用级自由点/页面级与集合级最近点/相对移动·绝对定位/删除/方向/选择/**SnapPointRange 批量**） | ✅ 第三十一波 |
| **颜色对象** | `cdr_color`（object_info / object_name / object_compare / object_blend / object_copy / object_string_assign / object_convert / object_fixed / object_assign / **object_model（HSB/CMY/BW/YIQ/Fixed/CreateColorEx）/ object_from_string**） | ✅ 第三十二/三十五波 / **`object_palette`（PaletteAssign/SpotAssignByName 按调色板 GUID 取色）/ `object_invalidate` / `model=fixed·script`**） | ✅ / **`object_components`（CorelScriptGetComponent 读分量）/ `object_convert_palette`（按调色板 GUID 转色）/ `object_modify_style_color`**） | ✅ |
| **矩形对象** | `cdr_rect`（info / offset / inflate / union / intersect / contains / copy / create_rectangle（四角圆角）/ create_ellipse（角度+扇形）） | ✅ 第三十四波 |
| **描边对象** | `cdr_outline`（info / set（含 SetProperties 全参）/ convert_to_object / copy / compare） | ✅ 第三十二波 |
| 位图 | `cdr_bitmap`（信息/转色彩模式/重采样/裁剪/取消裁剪/取像素色·图像缓冲信息·调色板模式 0~9·双色调/内容替换 SetImageData/位图特效命令串 ApplyBitmapEffect/**链接信息·重新加载·解析链接**） | ✅（Trace 除外） / **调色板选项转换 paletted_options（ConvertToPaletted2 + 可选已打开调色板）**）| ✅（Trace 除外）|
| 路径 | `cdr_path`（长度面积/偏移轮廓/追加子路径/折线点/反转方向·焊子路径·去重叠·参数取点·切/法向量·速度·交点·找节点/线段/**曲线合并（AppendCurve）/焊接（WeldWith）/命中判定（IsOnCurve）/曲线副本/曲线信息**） | ✅/**任意线段集合（segment_range：CreateSegmentRange + Add，set_type/add_node/select/remove_all）** | ✅/**追加曲线段（append_curve_segment）·抽出子路径（extract_subpath）·焊接落地（weld_to_shape）**）| ✅ |
| 导入导出 | `cdr_import`（**可传 import_options：外链位图/重采样/保留图层**） `cdr_export`（png/jpg/svg/pdf，可指定 DPI，可只导选区，**可传 export_options 控制抗锯齿/透明/压缩/尺寸**） | ✅ / **`export_api="document"`（Document.Export）/ `"bitmap"`（Document.ExportBitmap）** | ✅ / **`filter_id`（走 Layer.ImportEx，需导入过滤器号）** | ✅ |
| 文档 | `cdr_doc`（概况/清选区/清撤销栈/调色板/箭头/箭头登记/单位换算·取色·显示隐藏·版式·重复操作/字体列表·调色板列表·调色板取色/固定调色板打开·颜色枚举·专色创建/图案画布创建（含区域变换）/纹理库列表/文档元数据读写/本地化元数据/自定义调色板创建/数据字段·对象数据·色彩管理·线型样式/文档设置读写/应用信息·文档回滚/色彩配置文件·文档颜色转换/用户偏好 KV/命名视图/网格设置/叠印读取/**应用+文档事件队列**） | ✅ / **颜色上下文深操作（color_context_info/copy/merge）** / **渐变变换矩阵（fountain_transform）** | ✅ / **CorelScriptTools 辅助对象（script_tools：英寸/磅↔内部单位、ASin、BuildDate）** | ✅ |
| **对象数据** | `cdr_modify(changes={"data":{...}})` + `cdr_doc(action="object_data"/"data_fields"/"add_data_field")` | ✅ 第十一波 |
| 视图 | `cdr_view`（刷新/激活/文档↔屏幕坐标/屏幕↔文档距离/**屏幕叠加层（文字/矩形/直线/圆）**） | ✅ / **新窗口（new_window）**） | ✅ |
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
| 描边整体设置 | `Outline.SetProperties(Width, Style, Color, …)`（14 参）/ `SetPropertiesEx`（15 参） | **第三十二波纠正**：早先「TypeError」是因为把 `None` 传给了 int 形参；**14 个实参全给对（可选对象位传 None、可选数值位传 0/0.0）即可用**，宽度/颜色/线帽/连接/斜接都生效 | 已封装 `cdr_outline(action="set", changes={"use_set_properties":true,…})` |
| 文档设置快照 | `Document.SaveSettings(Tag)` / `RestoreSettings(Tag)` / `ResetSettings()` | 三个方法都**调用成功且不报错**，但实测把 `Resolution` 从 300 改成 72 再 `RestoreSettings` 并没有恢复（仍 72）—— Tag 语义未逆出，可见效果未确认 | 已封装为 `cdr_doc(action="save_settings"/"restore_settings"/"reset_settings")`，标注为实验性；需要可靠改设置请用 `cdr_doc(action="set_settings")` |







| 剪贴板格式粘贴 / 形状日志 | `Layer.PasteSpecial("CorelDRAW 图形", …)` / `Document.LogCreateShape(shape)` | PasteSpecial 返回 False（格式串不对）；LogCreateShape 报 com_error | 高级粘贴用 `cdr_layers(action="paste_ex")`；Log* 不暴露 |

| 导入/导出过滤器号不通用 | `Layer.ImportEx(path, 802, opts)` | 802 是 **PNG 导出**过滤器号，传给 ImportEx 时调用成功但 `shapes_added=0`（静默不导入）；传 0 让它自动判断才正常 | `cdr_import(filter_id=…)` 已暴露但标注需用**导入**过滤器号；X8 没有列出导入过滤器号的 API |



| 连接线/符号库造不出来 | `Layer.CreateFreeConnector(x1,y1,x2,y2)` / `CreateSymbol(x,y,名,库)` / `Application.SymbolLibraries` | CreateFreeConnector 报「发生意外」；CreateSymbol 报「找不到该对象」；SymbolLibraries 直接 com_error（本机未装符号库） | 用两点直线/贝塞尔自绘连接线 |
| 曲线信息数组接口不可用 | `Curve.PutCurveInfo(数组, 元素数)` | 传 tuple 报 `MemoryError: CreatingSafeArray`（pywin32 造不出该 SAFEARRAY）| 曲线信息读 `GetCurveInfo()` 即可 || Unicode 码页转换被 pywin32 弄坏 | `Application.ConvertToUnicode(串, 码页)` / `ConvertFromUnicode` | 传 GBK 字节串得到 `?D??`、传 "中文" 原样返回 —— API 期望 `char*`，pywin32 只给 BSTR，编码信息在边界就丢了 | 不封装；需要码页转换请用 Python 的 `bytes.decode("gbk")` || 变换矩阵对象造不出来 | `Shape.ApplyTransformMatrix` / `Curve.ApplyTransformMatrix` | 需要 `IVGTransformMatrix`，而它的工厂在 `IVGMathUtils` 上，`Application.MathUtils` 根本不存在（AttributeError，CLSID 未注册）| 矩阵变换改用 `Shape.SetMatrix` + `cdr_modify(affine=…)` || 范围删项会真删对象 | `ShapeRange.DeleteItem(1)` | 删完范围 Count 2→1，**页面对象数也从 2 变 1** | 只是想从范围里移除成员用 `range_remove`（`RemoveRange`），别用 DeleteItem || 文档级填充/描边工厂（字符串） | `Document.CreateFill("…")` / `CreateOutline("…")` | 传 CMYK/RGB 颜色串或 "UniformFill" 都只得到 **Type=0 / 宽度 0** 的空对象（字符串格式无效）| 填充/描边仍走 `cdr_modify` 或 `CreateUniformFill` |
| 页面自定义命令 | `Page.CustomCommand("", "", [])` | 报「未装载此命令」| 不暴露 |
| 样式改父级 | `Style.SetBasedOn` | ~~不可用~~ **第六十三波纠正：可用** —— 参数是**名称字符串**（传样式对象才报类型不匹配）；另外同一对象同类别连续 `CreateStyleFromShape` 会互相覆盖，要换不同对象建 | 已封装 `cdr_style(action="set_based_on")` 与 `action="assign"` || 范围级属性复制 | `ShapeRange.CopyPropertiesFrom(shape, mask)` | 报 `com_error: 执行方法 coVGShape::CopyPropertiesFrom 时发生意外`（掩码 1|2|4）| 属性复制请逐个对象用 `cdr_modify(changes={"fill_from"/"outline_from"})` || 自定义效果 | `Shape.CreateCustomEffect(id, …)` / `Shape.CreateCustomDistortion(id, …)` | 都报「未装载此效果」（需要先装插件效果）| 需安装对应插件，不暴露；**BSpline 已在第五十三波打通**（`Document.CreateBSpline` + `InsertControlPoint` + `Layer.CreateBSpline(IVGBSpline)`）|
| 数组入参（点集/形状集） | `Document.CreateCurveFitToPoints(Points, …)` / `Document.CreateCurveFromArray(Source, NumElements)` | 点集仍报 `TypeError: Python instance can not be converted to a COM object`；但 `CreateShapeRangeFromArray([shape,…])` **能接受 Python list**（传 1 个返回 Count=1）| 点集建曲线仍走逐点 AppendLineSegment || 叠印状态查询不对称 | `Shape.GetOverprintFillState()` / `GetOverprintOutlineState()` | 实测 outline 状态随开关变化（2→4），但 **fill 状态恒为 1**（开关前后都一样）；`RemoveFromContainer(1)` 报「参考的对象已不再存在」| 判断填充叠印请直接读布尔属性 `OverprintFill/OverprintOutline`（`cdr_shape_ops(op="overprint_state")` 一并返回）|| 矩阵平移量单位 | `Shape.SetMatrix(d11,d12,d21,d22,tx,ty)` / `GetMatrix` | 读回的初始矩阵 tx/ty 是 `-850000`、`-285000` 这类大数 —— X8 内部按 1/100000 单位存平移；写 `tx=5` 实际位移极小 | 矩阵读写本身精确（写什么读什么），但**不要拿它做毫米级平移**，平移请用 `cdr_modify(move=…)` || 调色板色/专色工厂 | `Application.CreatePaletteColor(name, index)` / `CreateSpotColor(name, tint, ?)` | 分别报「找不到颜色 PANTONE 185 C」「未找到颜色 MCP专色」—— 它们是从**已存在**的调色板/文档专色里取，不是新建 | 新建专色仍用 `cdr_doc(action="spot_color")`（CreateSpotColorByName + 已打开调色板） || 文字结构体写入 | `Text.FontProperties(frames, struct)` / `SpaceProperties(...)` 等 4 个 propget/putter | pywin32 里这些名字是**属性**：当方法调报 `TypeError: IVGStruct... object is not callable`；直接赋值（`text.FontProperties = struct`）不报错，但会把字号写成 1.17e-38 之类的垃圾值 | `Font/Space/Align/Hyphenation` 四个快照**只读**（`cdr_text(struct_info=true)`）；写字号/字体仍走 `cdr_text(font=…, size=…)` |
| 文字区间下标版结构体 | `Text.FontPropertiesInRange(start,count,indexing)` | ~~字段全是垃圾~~ **第五十七波纠正：可用** —— 方法名没有 `Get` 前缀、且 `StartIndex` 必须 **1 基**；传 0 才得到 Size=0/Style=16 的假值 | 已封装 `cdr_text(range_struct={…})` || 节点连线 / 跨曲线接续 | `Node.ConnectWith(Target)` / `Node.ExtendSubPaths(Node2, JoinPaths)` | 两条**不同**开放曲线的末端节点互连，报「指定的节点必须在另一条曲线上」；`ExtendSubPaths` 报「无法连接曲线」—— 大概需要先在同一选择集里选中两条曲线 | 已提供同曲线内的 `cdr_node(action="join_nodes")`（实测 3→2 节点）|
| 文字粗斜体 | `CreateArtisticText(..., Bold, Italic, …)` / `Text.Story.Bold` | **第三十四波纠正**：这两个形参是 `cdrTriState`（**true = -1**）。传 Python `True`（=1）不仅不生效，`Story.Bold = True` 还会把粗体**关掉**；下划线形参又是另一套枚举（1=有、0=无） | 已封装 `_tristate()` / `_underline()`；`cdr_create`、`cdr_modify`、`cdr_text` 全部改走这两个映射 |
| 字体表 | `Application.ForceUpdateFontTable()` | 长时间批量建文本后，`CreateArtisticText` 传 `"Arial"` 会**静默回退默认字体**（中文机=宋体）且丢掉 Bold/Italic；调一次 FontTable 刷新后恢复 | `cdr_doc(action="refresh_fonts")`；建文本时也会自动兜底刷一次 |
| 中文字体文档的字体改写 | `Story.Font` / `TextRange.Font` | 当文档默认字体是宋体时，把 Font 改成 `"Arial"`、或设 `Bold/Italic`，**不报错但不生效**（Size/Underline 仍然生效） | 属 X8 字体替换限制；`cdr_text` 会如实回 `applied`，需要拉丁字体时建议新建文档先设默认字体 |
| 样式分组属性 | `Style.GetProperty("fill"/"outline")` | 报「为参数 Name 指定的值无效」 | 分组名不可读写，只能读写叶子属性（primaryColor / type / angle…） |
| 按拷贝打开 | `Application.OpenDocumentAsCopy(FileName, Options)` | 无论 Options 传 0 还是 None 都是 `TypeError: Python instance can not be converted to a COM object` | 只能 OpenDocument / 从模板创建；需要"副本"请 SaveAs 新路径 |
| HTML 兼容文本 | `Text.MakeHTMLCompatible(HTML)` | 返回 False，未见效果 | 不封装 |
| 断字设置写入 | `Text.SetHyphenationSettings(Frames, struct)` | 调用不报错但值不生效（读回仍是默认） | 只读暴露，写入不提供 |
| 字体结构体读取 | `Text.GetFontProperties(Frames)` | 返回**默认结构体**（Name=''、Size≈1e-38）；UI 字体实际值要用 `Text.Story.Font/Size` | 已改用 Story 读取 |
| 文本区间套样式 | `TextRange.ApplyStyle(样式名)` | 调用不报错，但实测 `Bold/Size` 未变化（字符样式未生效或需其他条件） | 不推荐；整体样式用 `Shape.ApplyStyle` |
| 文本格式查询的尺寸比较 | `TextRange.FindRanges("@size > 10")` | 报 `Incompatible unit powers of the operands: 'pt' and 'pt^0'`；`@size > 10pt` 解析失败 | 尺寸条件暂不可用；`@font = 'Arial'`、`@bold = true` 可用 |
| 位图换内容 | `Bitmap.SetImageData(Image, Alpha, X, Y)` | `Alpha=0` 报 `TypeError: Python instance can not be converted to a COM object`；**Alpha=None 可用**，像素内容被替换 | 替换像素必须传 None |
| 位图特效 | `Bitmap.ApplyBitmapEffect(UndoString, Command)` | 两个都是字符串；`Command` 是**效果命令串**，如 `"WhirlpoolEffect WhirpoolSpacing=20,WhirlpoolSmear=9,WhirlpoolTwist=70,WhirlpoolStreak=60,WhirpoolWarp=1,WhirlpoolRandomSeed=0"`，实测像素被改变 | 效果名/参数名需按 Corel 文档；已封装 `cdr_bitmap(action="effect")` |
| 尺寸/连接线 | `Layer.CreateLinearDimension/CreateConnector` | 需要 **IVGPoint**，但类型库里 `IVGPoint` **没有任何公开构造入口**（`Application` 只有 `CreateSnapPoint`，返回的是 `IVGSnapPoint` 而不是 `IVGPoint`；`MathUtils` CLSID 未注册）；传 `IVGSnapPoint` 仍是 `TypeError: Python instance can not be converted to a COM object` | 仍不可用；`CreateFreeConnector` 报「未实现」。已系统排查：SnapPoint/Range 全家族都能建、能改，唯独不能当点用 |
| 符号库 | `Application.SymbolLibraries` | 本机 Count=0，无符号库可 Add/插入 | 需先安装符号库文件 |
| 文档颜色转换 | `Document.ConvertToColorContext(profile.CreateColorContext())` | RGB→CMYK→RGB 实测成功（填充色由 RGB 变 CMYK 再回 RGB）；**灰度→RGB 会报「服务器出现意外情况」** | 已封装 `cdr_doc(action="convert_color")`；转换前建议先以 RGB→CMYK→RGB 顺序，避免灰→RGB |
| 色彩配置文件 | `Application.ColorManager.ColorProfiles` | 本机 64 个；`GetProfilesByColorModel(0/1/2)` 分别可取 RGB/CMYK/Gray；`DefaultColorContext` 含 3 个 | `CreateColorContext()` 直接调用不可用，须用 `profile.CreateColorContext()` |
| 应用偏好 | `RegisterUserApplicationPreference` / `Get·SetApplicationPreferenceValue` | 用户偏好可注册、读写（存 CorelDRAW 工作区）；读取真实应用偏好组（如 `VGDoc Pref Settings`）报「发生意外错误」 | 只支持用户自定义偏好（Group/Key） |
| OLE 文件嵌入 | `Layer.CreateOLEObjectFromFile(FileName, Link, DisplayAsIcon, Caption, Icon)` | 传绝对路径的 .txt 可嵌入，返回 `type_id=12` 的 OLE 形状（默认约页面居中，28×18mm） | 不弹出对话框；可用 cdr_modify 移动/删除 |
| 命名视图 | `Document.CreateView(Name, OriginX, OriginY, Zoom, Page)` | **Page 参数必须是 IVGPage 对象**；传 int 报 `TypeError: Python instance can not be converted to a COM object`；传 `page` 后成功创建 | 已封装 views/create_view/activate_view/delete_view |
| ContentIdentifiers | `Document.ContentIdentifiers` | 返回 Python **tuple**（空文档为 `()`），不是 COM 集合；本机无内容标识可测 | 批量内容替换未接入 |
| 文档网格 | `Document.Grid` | Type / SpacingX / SpacingY / Visible / Snap 均可读写；mm 上下文下间距为毫米（默认 5×5mm） | 已封装 `cdr_doc(action="grid"/"set_grid")` |
| 叠印 | `Shape.OverprintFill/Outline/Bitmap` | 三个布尔属性均可读写；`cdr_modify` 写入后回读一致 | 印前叠印设置已可用 |
| 段落缩进 | `TextRange.LeftIndent/RightIndent/FirstLineIndent` | 可写且回读正确（mm 上下文下为毫米） | `cdr_text(indent_left/right/first_line)`；Tabs/Columns 未封装 |
| 制表位 | `TextRange.Tabs` | `Clear/Add(Position,Alignment,Leadered)` 可用；读取 Position/Alignment/Leadered 正确 | 已封装 `cdr_text(clear_tabs, tab_stops)` |
| 文本分栏 | `Text.Frame.SetColumns(Count, WidthsAndGutters)` | 第二参数需数组，Python list/tuple/VARIANT 都报「类型不匹配/参数无效」 | 分栏创建暂不可用 |
| 颜色转换的稳定性 | `Document.ConvertToColorContext` | 同一会话多次转换（尤其转回 RGB/灰度）可能让 CorelDRAW 进程无响应 | 联调只保留一次 RGB→CMYK；需要多次转换时建议每次重启 CorelDRAW |
| 文档回滚 | `Document.Revert()` | 可用（修改后 15→16→15 对象）；**破坏性**，丢弃未保存修改 | 已封装为 `cdr_doc(action="revert")`，调用前确认已保存 |
| 立体化/透镜创建 | `Shape.CreateExtrude(...)` / `Shape.CreateLens(...)` | 只传前几个 int 参数报 `TypeError: Python instance can not be converted to a COM object`；**传全参数（含颜色对象）即可** | 已封装并补齐默认颜色参数 |
| 点对象创建 | `Application.MathUtils.CreatePoint` | `MathUtils` CLSID `{E98800C5-…}` 报「没有注册类」，Application 上也没有 MathUtils 属性 | 尺寸标注/连接线仍不可用 |
| 混合创建 | `Shape.CreateBlend(shape, …)` | 省略可选参数会 `TypeError`（Path 参数默认 0 不是 IDispatch）；**传全参数且 Path=None 可用** | 已封装 `cdr_shape_ops(op="blend")` |
| 图案填充（文件型） | `Fill.ApplyPatternFill(Type=1/2, FileName=…)` | 提示「无法装入以下图样文件」 | 需要图案图像文件；**内存画布型（Type=0 + PatternCanvas）已可用** |
| 影线填充 | `Fill.ApplyHatchFill` / `Document.HatchLibraries` | `HatchLibraries.Count = 0`、`ActiveLibrary = None`（本机未装 Hatch 库） | 需要先安装 Corel 的 Hatch 填充库 |
| 曲线贴到笔触 | `Curve.CreateCurveMappedToStroke` | `com_error(-2147352571, 类型不匹配)`（Stroke 参数不是普通 Curve） | 需要 Artistic Media 的笔触对象，暂不封装 |
| 用户吸附点带 ID 创建 | `SnapPoints.AddUserSnapPointEx(ID, x, y, Direction, UseDirection)` | ID 传 int / 字符串 GUID / `uuid.UUID` / `{GUID}` / IVGUserSnapPoint / IVGSnapPoint / VT_BSTR / VT_DISPATCH 一律报「为参数 ID 指定的值无效」`0x80070057` | 用 `AddUserSnapPoint`（自动生成 GUID），再 `SnapPoints.User("生成的GUID")` 反查 |
| 符号库 | `app.SymbolLibraries` | 对象存在，但 `Count = 0`（本机未装符号库） | `CreateSymbol` 需要库内符号名，无库可用（`ConvertToSymbol` 可用，已封装为 `to_symbol`） |
| 剪贴板 | `Document.Copy/Paste` | `AttributeError` | Document 上没有这两个方法（但 `Layer.Paste` 可用，已封装为 `cdr_layers(action="paste")`） |
| 吸附设置 | `Document.SnapToGrid/SnapToObjects` | `AttributeError` | 同上 |
| PowerClip 提取 | `Shape.PowerClip` | 对象存在但无 `Extract` | 只能"放入"不能"取出" |
| 页面背景 | `Page.Background` | 不存在 | — |
| 文档元数据（直接属性） | `Document.Keywords/Notes` | 直接**读**会 AttributeError，但**可以赋值**（write-only，写入后 Metadata 能读到） | 读元数据请走 `Document.Metadata` |
| 本地化元数据 | `Metadata.LocalizableTitle/...` | `IVGLocalizableString` 的 DefaultLangString 可读写；GetLanguages 返回 tuple（含 "" 与已设置语言如 "2052"），GetLangString/SetLangString 按 LCID 读写 | 已封装 `metadata_localized/set_metadata_localized` |
| 屏幕叠加层 | `Application.CreateOnScreenText/Curve/Handle` | SetTextAndPosition/SetTextColor、SetPen/SetRectangle/SetLine/SetCircle、Show/Hide 全部调用成功；坐标为**屏幕像素**，xor 绘制、不写入文档 | 已封装 `cdr_view(overlay_*)`；窗口重绘/关闭后叠加层可能消失 |
| 应用事件 | `IVGApplicationEvents` 连接点 | 生成的 `DIVGApplicationEvents(application)` 可订阅；实测 DocumentNew/WindowActivate/DocumentClose 等都有回调 | 需正确的 `vgcoreauto.tlb` 路径（app.Path 指向 Draw，实际在 ../Programs64）；ApplicationEvent 高频已忽略 |
| 文档事件 | `IVGDocumentEvents` 连接点 | 每个文档挂 sink 后，创建/移动/删除对象、图层/页面变化、SelectionChange 都有回调（事件名带 `Doc:` 前缀） | 文档在首次访问或 DocumentNew/Open 时自动挂载；关闭时卸载 |
| 导出选项 | `Application.CreateStructExportOptions` | 可写 AntiAliasingType/Compression/Dithered/Transparent/MaintainAspect/MaintainLayers/Overwrite/Matte*/ResolutionX·Y/SizeX·Y/UseColorProfile/AlwaysOverprintBlack；`cdr_export(export_options={...})` 后文件正常导出 | 复杂 CMYK 文档+透明选项可能失败，建议用干净 RGB 文档；失败会回退默认选项 |
| 保存选项 | `Application.CreateStructSaveAsOptions` | EmbedICCProfile/EmbedVBAProject/IncludeCMXData/KeepAppearance/Overwrite/ThumbnailSize/Version/Filter/Range 可写；`SaveAs(path, opts)` 成功 | `cdr_save(save_options={...})`；Version 语义未单独验证 |
| 导入选项 / 链接位图 | `Application.CreateStructImportOptions` / `Bitmap.LinkFileName` | `link_bitmap_externally=true` 导入后 `ExternallyLinked=true`、`LinkFileName` 指向源图；`UpdateLink()` 可重新加载 | 导入后会自动补对象名，便于后续按名操作；`ResolveLink` 在已链接位图上会清空 LinkFileName |

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
> * 调色板查找/插入/默认 / 命名视图 / 数据字段重排 / B样条扩容 **可用（第七十七波）**：`Palette.FindColor/InsertColor/MakeDefault`、`Views.AddActiveView`、`DataField.Reorder`、`ControlPoints.Resize`。
> * SetLight* 正式方法 / 点距 / 图案原点 / B样条插点 **可用（第八十七波）**：`extrude_set(lights=…)`、`point_math`、`hatch_fill(pattern_origin=…)`、`bspline_resize(insert=…)`。
> * 图像列表收字节 / 状态栏位图 / 数据源注销 / 配置文件安装 **可用（第八十六波）**：`imagelist_info(image_path=…)`、`status_text(set_bitmap)`、`datasource_info`、`color_profile_install`。
> * 图像列表登记 **可用（第八十五波）**：`cdr_doc(imagelist_info)`，`AddBitmap(键, ICUIBitmapImage)`。
> * 主窗口控制 / 屏幕矩形 / 状态栏 / GMS 路径 **可用（第八十四波）**：`frame_maximize/restore/tile/combine`、`screen_rect`、`status_text`、`gms_info`。
> * 立体化双灯 / 剪贴板 / 组件检测 **可用（第八十三波）**：`extrude_set(effect={"lights":[…]})`、`clipboard_info`、`components_info`。
> * 点集容器编辑 **可用（第八十二波）**：`point_range_edit`（Reverse/AddPoints/InsertPoint/RemoveAdjacentDuplicates）与 `fit_to_points(reverse=true)`。
> * 立体化效果复制 / 数据字段详情 **可用（第八十一波）**：`extrude_set(effect={"copy_from":…})`、`cdr_doc(datafield_info)`。
> * OLE 对象只读清点 **可用（第八十波）**：`cdr_doc(ole_info)` 清点页面上的 Type 12 对象并读 IVGOLE 只读属性。
> * 文本框多栏 **可用（第七十九波）**：`cdr_text(columns={"count":…,"widths":…,"gutters":…})`；`IVGTextFrame` 达 100%。
> * 立体化灯光/斜角 **可用（第七十八波）**：`extrude_set` 支持 `light_index/light_intensity/bevel_depth/bevel_angle`，返回 3 盏灯的状态。
> * 扭曲中心重设 **可用（第七十六波）**：`distort` 流程里自动调 `EffectDistortion.CenterDistortion()`。
> * 批量斜切 / 渐变色标枚举 / 线型表保存 **可用（第七十五波）**：`ShapeRange.SkewEx`（range_skew）、`FountainColors.Item(i)` 枚举、`OutlineStyles.Save`。
> * 拉伸消失点共享 / 窗口排列 **可用（第七十四波）**：`ext.VanishingPoint.Share(effect.Extrude)`、`Windows.Arrange`。
> * 叠加层补齐 / 文本框列信息 **可用（第七十三波）**：`OnScreenText` 7/7、`OnScreenCurve` 11/11、`cdr_text(frame_info=true)`。
> * 屏幕手柄 / 包络效果对象 / 按名选配置文件 **可用（第七十二波）**：`CreateOnScreenHandle`、`effect.Envelope`（Select/CreateFrom/CopyFrom/CopyFromCurve）、`ColorProfiles.SelectByName`。
> * 剖面线参数 / 渐变控制点 / 椭圆几何 / 调色板工厂 / 窗口查找 **可用（第七十一波）**：`Fill.Hatch`、`Fountain.Colors`、`Shape.Ellipse`、`Palettes.CreateFrom*`、`Windows.FindWindow`。
> * HTML 兼容 / 色彩配置文件指派 **可用（第七十波）**：`Text.MakeHTMLCompatible`、`ColorContext.SetColorProfile(模型号, 匹配模型的配置文件)`。
> * 对象树 / 样式填充存盘 **可用（第六十九波）**：`Document.TreeManager`、`Style.Fill.SaveFill/LoadFill(path, IVGFillMetadata)`。
> * 双色调子对象 / 网格频率 / 线性标注子对象 **可用（第六十八波）**：`IVGDuotone(+Overprint)`、`Document.Grid.SetFrequency`、`Dimension.Linear`。
> * 矩形圆角 / 克隆链 / 投影偏移 **可用（第六十七波）**：`Shape.Rectangle.SetRadius/SetRoundness`、`Shape.CloneLink.RestoreAllLinks`、`EffectDropShadow.SetOffset`。
> * 线型虚线段 / 自定义工具形状属性 **可用（第六十六波）**：`OutlineStyles.Item(n)` 读`DashCount/DashLength/GapLength`、`CreateToolShapeAttributes()` 的 12 个 `Set*`。
> * 视口缩放/混合效果/调色板管理/进度条 **可用（第六十五波）**：`ActiveWindow.ActiveView` 15 个缩放方法、`Shape.Effects.Item(1).Blend`、`Application.PaletteManager`、`Application.Status`。
> * 标注/图框/辅助线/文本框链接 **可用（第六十四波）**：`Shape.Dimension`、`Shape.PowerClip`、`Shape.Guide`、`Text.Frames.Item(i).LinkTo/UnLink`。
> * 样式改父级/复制 **可用（第六十三波）**：`Style.SetBasedOn(名称串)`、`Style.Assign(样式对象)`；`IVGStyle` 达 100%。
> * 副本打开 / 字体结构体工厂 **可用（第六十二波）**：`OpenDocumentAsCopy(path, 选项)`、`CreateStructFontProperties()`。
> * 中心定尺寸/拉伸/范围排除 **可用（第六十一波）**：`SetSizeEx(cx,cy,宽,高)`、`StretchEx(cx,cy,X倍数,Y倍数,缩放文字)`（**前两参是中心点**）、`AllExcluding(索引数组，1 基)`。
> * 剖面线填充 **可用（第六十波）**：`ApplyHatchFill`（第 2 参必须是 int 序号）、`ApplyCustomHatchFill`（18 参）。
> * 点集拟合/追加 **可用（第五十九波）**：`IVGPointRange`（来自 `SubPath.GetEvenlySpacedPoints`）+ `CreateCurveFitToPoints(AndCusps)` / `AppendSubpath*Points`；Python list 不是它要的类型。
> * 调色板/专色工厂 **可用（第五十八波）**：`CreatePaletteColor(GUID, 序号)`、`CreateSpotColor(GUID, 专色ID, Tint)`；`CreateColorContext2` / `CreateProofColorSettings` 的 RenderingIntent **1~4**。
> * 文字区间结构体 **可用（第五十七波）**：`FontPropertiesInRange` 等四个读、`Set*PropertiesInRange` 三个写；**方法名无 Get 前缀、StartIndex 1 基**。
> * 曲线映射到骨架线 **可用（第五十六波）**：`Curve.CreateCurveMappedToStroke(子路径,…)`（`Stroke` 要 `IVGSubPath`，不是 `IVGCurve`）。
> * 连接器 / 标注 **可用（第五十五波）**：`Application.CreateSnapPoint(x,y)` 造端点后，`CreateConnector` / `CreateConnectorEx` / `CreateRightAngleConnector` / `CreateLinearDimension`(22 参) / `CreateAngularDimension`(18 参) 全部可用。
> * 同级遍历 / 解组家族 **可用（第五十四波）**：`Shape.GetNext/GetPrevious(方向)`、`ShapeRange.UngroupAll/UngroupAllEx/UngroupEx`（纠正了"空操作"的旧结论）。
> * B 样条 **可用（第五十三波）**：`Document.CreateBSpline` + `InsertControlPoint` + `Layer.CreateBSpline`（`Layer.CreateBSpline` 要的是 `IVGBSpline` 不是 Curve）。
> * 整批套描边 **可用（第五十三波）**：`ShapeRange.ApplyOutline(Outline.GetCopy())`。
> * 路径偏移查询 **可用（第五十二波）**：`SubPath.GetPointPositionAt`（取点）、`FindSegmentOffset`（定位线段+参数偏移）。
> * 文本结构体写入 **可用（第五十一波）**：`CreateStructAlignProperties/SpaceProperties/HyphenationSettings` + 对应 `Set*Properties(0, 结构体)`，仅限段落文本。
> * 色彩配置按模型读 **可用（第五十一波）**：`ColorContext.ColorProfile(model)`。
> * 形状斜切 / 子路径追加与抽出 **可用（第五十波）**：`Shape.SkewEx`、`SubPath.AppendCurveSegment`、`SubPath.Extract`。
> * 焊接落地 **可用（第五十波）**：`Curve.BindToDocument` + `Layer.CreateCurve`，把 `WeldWith` 的游离曲线变成真形状。
> * 颜色分量读取 **可用（第四十九波）**：`CorelScriptGetComponent(模型号)` 返回 `(5000+模型号, v1..v7)`；`ConvertToPalette(调色板 GUID)` 可转指定调色板。
> * 形状级属性复制 **可用（第四十九波）**：`Shape.CopyPropertiesFrom(源, 掩码)`。
> * 颜色 Assign 家族 **可用（第四十八波）**：`FixedAssign`、`CorelScriptAssign`、`PaletteAssign(调色板 GUID, 序号)`、`Invalidate`。
> * 图层复制 / 对齐到范围 **可用（第四十八波）**：`CopyToLayerAsRange`、`AlignToShapeRange`。
> * 宽文本 / 节点搬运 / 整批描边 **可用（第四十七波）**：`CreateArtisticTextWide`、`Node.GetPoint/SetPoint`、`ShapeRange.SetOutlineProperties(Ex)`、`AlignRangeToPage`。
> * VBA 辅助对象 **可用（第四十六波）**：`CorelScriptTools()` 提供英寸/磅↔内部单位换算、`ASin`、`BuildDate`。
> * 描边 15 参 / 文字编辑态 **可用（第四十五波）**：`Outline.SetPropertiesEx`、`Text.BeginEdit`。
> * 颜色上下文深操作 **可用（第四十四波）**：`GetCopy/CopyAssign/Merge/IsSame`、`ColorProfile(model)`、`ColorProfileNameList`、`ReadOnly`。
> * 新窗口 / 渐变变换 / 纹理透明 **可用（第四十四波）**：`ActiveWindow.NewWindow`、`FountainFill.GetTransformations`、`Transparency.ApplyTextureTransparency`。
> * 调色板选项 **可用（第四十三波）**：`CreateStructPaletteOptions` + `Bitmap.ConvertToPaletted2`，可指定颜色数/抖动/调色板。
> * 高级粘贴 / 两种导出 **可用（第四十三波）**：`Layer.PasteEx`、`Document.Export`、`Document.ExportBitmap`（记得 `Finish()`）。
> * 对象列表建范围 **可用（第四十二波）**：`Document.CreateShapeRangeFromArray([shape,…])` 接受 Python list，可再整批 `RotateEx`/`ConvertToBitmapEx`。
> * 色彩上下文指派 / 样式表默认 / 箭头变体 **可用（第四十二波）**：`AssignColorContext`、`SaveStyleSheetAsDefault`、`CreateArrowHeadEx2` + `CreateArrowHeadOptions`。
> * 任意节点/线段集合 **可用（第四十一波）**：`Application.CreateNodeRange/CreateSegmentRange` + `Add`，可对**不连续**的节点/线段做 Move/Rotate/Skew/Stretch/整形/加节点。
> * 位图构造器 **可用（第四十一波）**：`Layer.CreateBitmap2` / `CreateBitmapRect`。
> * 参考点定位 / 吸附反查 **可用（第四十波）**：`GetPositionEx/SetPositionEx`（9 个参考点）、`FindSnapPoint(ReferenceData)`。
> * 位图化 / 链接查询 **可用（第四十波）**：`ConvertToBitmapEx(8 参)`、`GetLinkedShapes(LinkType)`。
> * 形状矩阵/扭曲 **可用（第三十九波）**：`GetMatrix/SetMatrix`、`IsTypeAnyOf`、`IsOnShape`、`OrderIsInFrontOf`、`PlaceTextInside`、`DuplicateAsRange/CloneAsRange`，以及 `CreatePushPullDistortion/CreateTwisterDistortion/CreateZipperDistortion`。
> * 结构体建文档 **可用（第三十八波）**：`CreateStructCreateOptions` + `CreateDocumentEx` 一次定名称/页宽/页高/分辨率；`CreateStructOpenOptions` + `OpenDocumentEx` 可带 CodePage 打开。
> * Fill 工厂 **可用（第三十八波）**：`Document.CreateUniformFill(IVGColor)` + `ShapeRange.ApplyFill(fill)` 批量上色。
> * 层构造器完整参数 **可用（第三十八波）**：`CreateRectangle(8)` / `CreateEllipse(7)` / `CreatePolygon(10)`。
> * 文字结构体 **可读（第三十七波）**：`Text.FontProperties/SpaceProperties/AlignProperties/HyphenationSettings` 四个属性返回结构体快照（只读，setter 会写坏字号）。
> * 区间宽接口 **可用（第三十七波）**：`InsertBeforeWide/InsertAfterWide/ReplaceWide/Collapse/SetRange/InRange/IsSame/SetLineSpacing`。
> * 样式继承 **可读（第三十七波）**：`GetProperty/GetOverridePropertyNames/IsPropertyInherited/ClearProperty`。
> * 节点拓扑 **可用（第三十六波）**：`Next()`/`Previous()`/`GetDistanceFrom`/`AveragePositionWith`/`JoinWith`；`ConnectWith`/`ExtendSubPaths` 需要选择集前置条件，暂不可用。
> * 曲线合并/焊接 **可用（第三十六波）**：`AppendCurve`（原地并成子路径）、`WeldWith`（返回新曲线）、`IsOnCurve`、`GetCopy`、`GetCurveInfo`。
> * COM 单位换算 **可用（第三十五波）**：`Application.ConvertUnits(value, cdrUnit, cdrUnit)`，inch=1 mm=3 cm=4 point=14；比 `Document.ToUnits/FromUnits` 更直接。
> * 颜色工厂 **补齐（第三十五波）**：`CreateHSBColor/CMYColor/BWColor/YIQColor/FixedColor/CreateColor(串)/CreateColorEx(模型号,…)` 全部可用。
> * 定尺寸加页/插页 **可用（第三十五波）**：`AddPagesEx(count,w,h)` / `InsertPagesEx(n,before,page,w,h)`。
> * 矩形对象 **可用（第三十四波）**：`IVGRect` 的 SetRect/Offset/Inflate/Union/Intersect/IsPointInside/GetCopy/Clear 全通，且 `Layer.CreateRectangleRect/CreateEllipseRect` 能按矩形造形。
> * 曲线段 **可直接建**：`Layer.CreateCurveSegment`（端点 + 两端控制柄长度/角度）。
> * 原生范围对齐 **可用（第三十三波）**：`AlignRangeToPageCenter/Grid/Point/Shape/ShapeRange`，一次只吃一个轴（`cdrAlignLeft=1/Right=2/HCenter=3/Top=4/Bottom=8/VCenter=12`），`cdr_align(to=…)` 会把 horizontal/vertical 拆成两次调用。
> * 范围批量 **可用（第三十三波）**：`AddRange/RemoveRange/Unlock/BreakApartEx/ClearEffect/IndexOf/CountAnyOfType/ExistsAnyOfType/FindAnyOfType`、`SetPositionEx/GetPositionEx/RotateEx/SkewEx`；`UngroupAll` 系列是空操作（改用逐个 `Shape.Ungroup()`）。
> * 颜色对象 **可用（第三十二波）**：`ToString/GetName/SetName/StringAssign/IsSame/IsSimilar/GetColorDistanceFrom/BlendWith/GetCopy` 与 `ConvertTo*`（原地）、`*Assign`（分量直写）、`ConvertToFixed` 全通；`ConvertToPalette`/`SpotAssignByName`/`CopyAppearance` 不可用。
> * 描边对象 **可用（第三十二波）**：`NibStretch/NibAngle/MiterLimit/Justification` 可读写，`SetProperties` 全实参可用，`ConvertToObject` 可把描边转成曲线。
> * 吸附点 **可用（第三十一波）**：`Shape.SnapPoints` 枚举、`AddUserSnapPoint`、`Document.CreateFreeSnapPoint`、
>   `Application.CreateSnapPoint`、`SnapPoints.Range` 批量 Move/Delete/Find/ChangeDirection/SetAutoSnap；
>   移动对象吸附点会转成用户吸附点；`FindClosestSnapPoint` 未命中返回 Type=16 的伪命中。
> * 文档颜色转换 **可用（第十八波）**：`Application.ColorManager` 提供 64 个配置文件，
>   `profile.CreateColorContext()` → `Document.ConvertToColorContext()` 可 RGB↔CMYK↔Gray；
>   灰度→RGB 会报服务器异常，其余方向正常。

---

## 三之补、逐项定性：剩余未覆盖方法（2026-10-26 实测证据）

> 这一节把 `IVGApplication` / `IVGDocument` / `IVGPage` / `IVGText` 里**还没覆盖**的方法逐个定性，
> 每条都有实测结果。**标「挂起」的严禁调用**（会把 CorelDRAW 卡死，COM 无法中断）。

| 方法 | 实测结果 | 结论 |
|---|---|---|
| `Application.CorelScriptTools()` | 返回 `ICorelScriptTools`；`FromInches(1)=254000`、`FromPoints(1)=3527.778`、`ToInches/ToPoints` 往返一致、`ASin` 返回弧度、`BuildDate` 返回日期 | ✅ **已封装** `cdr_doc(action="script_tools")` |
| `Application.CreateToolShapeAttributes()` | 返回 `IVGToolShapeAttributes` 对象，无异常 | 用途未确认，不封装 |
| `Application.CreateFillMetadata()` | 返回 `IVGFillMetadata`（Title/Description/Keywords/… 可读写），但 `IVGFill` 上没有挂载点 | 悬空对象，不封装 |
| `Application.CreateColorContext2()` | 「非选择性的参数」 | 需要不同的参数形态，未逆出 |
| `Application.CreateProofColorSettings("sRGB IEC61966-2.1", 0, False)` | 「发生意外」 | 软打样设置不可用 |
| `Application.RegisterToolState(...)` | 「类型不匹配」 | 自定义工具状态不可用 |
| `Application.StartTemporaryToolState("mcp")` | 返回 None、无异常 | 但配套的 RegisterToolState 失败，无实际用途 |
| `Application.CreateProperties()` | 返回 `IVGProperties`，但 `Count=0` 且缺 `Name/Add/Value` | 用户属性不可用 |
| `Application.CreatePaletteColor` / `CreateSpotColor` | ~~报错不可用~~ **第五十八波纠正：可用** —— 第一个参数是**调色板 GUID**，GUID 传对就正常造色；`CreateSpotColor` 还需要**含专色**的调色板 | 已封装 `cdr_color(object_palette_factory / object_from_spot)` |
| **`Application.AddPluginCommand` / `RemovePluginCommand`** | **调用后 300 s 未返回（进程卡死）** | 🚫 **严禁调用** |
| `Application.AdviseEvents` / `UnadviseEvents` | 需要事件接收器对象（传 None 会卡住） | 🚫 不用（本服务用连接点 sink 已实现事件） |
| `Application.Evaluate` / `InitializeVBA` / `ImportWorkspace` / `Quit` | 主动不测：VBA 求值、初始化 VBA、导入工作区、退出进程 | 🚫 危险 |
| `Application.cdrMixed*` | 类型库常量（MixedDouble/Long/Single） | 不是方法 |
| `Document.CustomCommand("", "", [])` | 「发生意外」 | 不封装 |
| `Document.ReplaceContentByIdentifier([], [])` | 返回 None、无异常；但 `ContentIdentifiers` 为空 tuple，无可替换内容 | 无内容标识时是空操作 |
| `Document.LogCreateShape` / `LogCreateShapeRange` | com_error | 不封装 |
| **`Document.GetUserArea` / `GetUserClick`** | **调用 20 s 未返回（挂起）** | 🚫 **严禁调用**（交互回调） |
| `Document.CreateCurveFitToPointsAndCusps(...)` | `TypeError: Python instance can not be converted to a COM object` | 点集数组仍不可用 |
| **`Document.CreateCurveFromArray([], 0)`** | **进程无返回（疑似卡死）** | 🚫 严禁调用 |
| `Page.CustomCommand` | 「未装载此命令」 | 不封装 |
| `Text.FontProperties/SpaceProperties/AlignProperties/HyphenationSettings` 的**写入端**（含 `*InRange`） | 属性赋值会写坏字号（1.17e-38）；直接 `_oleobj_.Invoke` + `VARIANT(VT_I4/VT_VARIANT)` 全部报「类型不匹配」 | 🚫 pywin32 无法 marshaling `VT_USERDEFINED` 参数；结构体**只读** |
| `Text.ImportFromFile` / `MakeHTMLCompatible` / `PrintOut` / `PrintColorProof` | 已知会卡模态框或真的打印 | 🚫 危险 |

## 四、已知但未封装（低频）

这些**能用**，只是 Agent 日常用不上，暂不占工具位：

* `IVGWindow` 的缩放/新建窗口（`Zoom`、`NewWindow`）：截图/预览才需要
* `IVGColorContext`：色彩配置、配置文件合并（交付环节）
* `IVGFountainFill.GetTransformations`：读渐变变换矩阵（渐变微调已能改角度/平移/矩阵）
* `IVGTraceSettings`：PowerTRACE 参数对象（Trace 本身不安全，不封装）
* `Shape.CopyPropertiesFrom(Source, Properties)`：实测 1~15 都返回 True，但位掩码语义未逆向出来
* `Curve.GetCurveInfo`：返回 `com_struct(PositionX, PositionY, ElementType, NodeType, Flags)` 数组，语义未完全确认
* `Layer.PasteEx` / `PasteSpecial`：`PasteEx` 传 0 会 TypeError；`Paste` 已封装
* `IVGImage.Blit` / `FlipArea`：像素缓冲区合成可用（已实测 Blit），暂未做成工具参数
* `Application.OpenDocumentAsCopy` / `CreateDocumentFromTemplate`：需要模板/文件路径，未纳入工具面
* `Application.CreateHSBColor/CreateYIQColor/CreateCMYColor/CreateBWColor`：实测可用，但颜色运算已走纯 Python
* `Palettes.Create / CreateFromDocument / CreateFromSelection`：可新建调色板并 `AddColor/SaveAs`，
  但会落盘点色板文件，暂未做成工具（`OpenFixed` 已封装）
* 符号库插入：需要符号库名称，`CreateSymbol` 本身也未打通（吸附点已在第三十一波单独补齐）
* `Bitmap.ResolveLink/UpdateLink`：需要外部链接位图才有意义，未做联调
* `Bitmap.ConvertToPaletted(9=自定义)`：需要 `palette_index`（已支持，配合 `palette_open_fixed`）

---

## 五、复现方式

```powershell
python -m pytest -q              # 93 项纯逻辑
python scripts/live_smoke.py     # 596 项端到端（真实 CorelDRAW，**跑前先确认没有残留 CorelDRW 进程**）
```

联调断言覆盖了上表所有 ✅ 项。反解脚本属于过程产物，已按要求放在 `tmp/` 并在完成时清理；
如需重建，按本文件"怎么得出这份清单"三步即可复现。

---

## 六、量化覆盖率（2026-10-27 第八十七波实测）

方法：反解类型库拿到 196 个自动化类 / **1619** 个方法（已剔除 `QueryInterface` 等 COM 基础设施）；
用 AST 扫描 `src/cdrx8_mcp/*.py`，收集所有 `.方法名(` 调用与 `setattr/getattr` 的属性名，
与主干接口的方法名求交集。**注意：属性名（如 `curve.Nodes`）也会被计入**，所以本波数字含真实属性访问。

**全局：1208 / 1619 = 74.6%**（21.2% → … → 74.2% → 本轮 74.6%）
**主干 20 接口：670 / 743 = 90.2%**

复现命令：`python scripts/coverage_audit.py`（会临时展开类型库到 `tmp/gen_py`，跑完自动清理）

| 接口 | 方法数 | 已调用 | 覆盖 |
|---|---:|---:|---:|
| `IVGShape` | 123 | 113 | **92%** |
| `IVGShapeRange` | 125 | 118 | **94%** |
| `IVGSubPath` | 38 | 37 | **97%** |
| `IVGDocument` | 74 | 60 | **81%** |
| `IVGPage` | 21 | 20 | **95%** |
| `IVGCurve` | 36 | 33 | **92%** |
| `IVGLayer` | 49 | 42 | **86%** |
| `IVGTextRange` | 31 | 30 | **97%** |
| `IVGApplication` | 78 | 59 | **76%** |
| `IVGText` | 29 | 25 | **86%** |
| `IVGNode` | 18 | 16 | **89%** |
| `IVGBitmap` | 15 | 14 | **93%** |
| `IVGFill` | 12 | 11 | **92%** |
| `IVGWindow` | 8 | 8 | **100%** |
| `IVGFountainFill` | 6 | 6 | **100%** |
| `IVGColor` | 43 | 41 | **95%** |
| `IVGOutline` | 8 | 7 | **88%** |
| `IVGTransparency` | 8 | 7 | **88%** |
| `IVGStyle` | 15 | 15 | **100%** |
| `IVGColorContext` | 6 | 5 | **83%** |
| **合计** | **743** | **665** | **89.5%** |

> **怎么读这两个数**：
> 1. 全局 **72.9%** 覆盖 1619 个方法（含变体与重复）；主干 20 接口 90.3%。
> 2. 它是**方法/属性名匹配**的近似值，同名方法会跨接口误计。
> 3. `IVGColor` 已到 **91%**：分量**写**走 `*Assign`，分量**读**走 `CorelScriptGetComponent`（`(5000+模型号, v1..v7)`）；纯 Python 只用于无 COM 场景的换算。
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

## 七之十八、第二十波已接入（2026-10-25，OLE 文件嵌入、命名视图）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **OLE 文件嵌入** | `cdr_create` 类型 `ole_file`：`{"type":"ole_file","path":"绝对路径","caption":"..."}` | ✅ 嵌入 .txt 返回 `type_id=12` 形状（姓名已自动命名，bbox 约 28×18mm 居中） |
| **命名视图** | `cdr_doc(action="views"/"create_view"/"activate_view"/"delete_view")` | ✅ 创建 W20View（zoom=150, page=1）→ 列表 count=1 → 激活 → 删除 → count=0 |
| 关键坑 | `Document.CreateView` 的 Page 参数 | 必须传 **IVGPage 对象**；传 int 报 TypeError |
| ContentIdentifiers | `Document.ContentIdentifiers` | 返回 Python tuple（空文档 `()`）；本机无内容标识可测 |
| 视图持久化 | 视图保存在文档中 | 未单独验证跨保存/重开；本波验证了创建、列出、激活、删除 |

## 七之十九、第二十一波已接入（2026-10-25，文档网格设置）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **网格读取** | `cdr_doc(action="grid")` | ✅ 默认 type=1、spacing=5×5mm、visible=false、snap=false |
| **网格写入** | `cdr_doc(action="set_grid", grid_spacing_x/y, grid_visible, grid_snap, grid_type)` | ✅ 写入 12×18mm、显示/吸附开启后回读一致；可再恢复原值 |
| 单位 | `Document.Grid.SpacingX/Y` | 在 mm 上下文下为毫米；写 12 读回 12.0 |
| 仍未可用 | 尺寸标注/连接线 | 用 SubPath 取点后 raw `InvokeTypes` 仍 TypeError/类型不匹配 |

## 七之二十、第二十二波已接入（2026-10-25，叠印与段落缩进）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **叠印读取** | `cdr_doc(action="overprint", target=…)` | ✅ 返回 overprint_fill/outline/bitmap（默认 false） |
| **叠印写入** | `cdr_modify(changes={"overprint_fill":true,"overprint_outline":true,"overprint_bitmap":true})` | ✅ 写入后回读 true；适合印前输出 |
| **段落缩进** | `cdr_text(indent_left, indent_right, indent_first_line)` | ✅ 写 5/6/3mm 后回读一致；`cdr_text` 读取结果含 `indent` 字典 |
| 未封装 | `TextRange.Tabs` / `TextRange.Columns` | 属性存在，但未验证可写性，暂不封装 |

## 七之二十一、第二十三波已接入（2026-10-25，制表位）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **制表位写入** | `cdr_text(clear_tabs=true, tab_stops=[{position,alignment,leadered}, …])` | ✅ 清空后写入 10mm/左对齐、25mm/右对齐带前导符，`applied` 含 tabs |
| **制表位读取** | `cdr_text` 结果 `tabs` | ✅ 回读 2 项：position=10/25、alignment=0/1、leadered=true |
| 文本分栏 | `Text.Frame.SetColumns` | ❌ 第二参数需数组，list/tuple/VARIANT 均失败，暂不可用 |
| 颜色转换稳定性 | `ConvertToColorContext` | 同一会话多次转换可能令 CorelDRAW 无响应；联调只保留单次 RGB→CMYK |

## 七之二十二、第二十四波已接入（2026-10-25，本地化元数据）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **本地化元数据读取** | `cdr_doc(action="metadata_localized")` | ✅ 返回 Title/Subject/Keywords/Notes/Copyright 的 default 与 languages 映射 |
| **本地化元数据写入** | `cdr_doc(action="set_metadata_localized", metadata={"Title":{"default":"…","2052":"…","1033":"…"}})` | ✅ 写入 Default Title / 中文标题 / English Title 后回读一致；Keywords 直接字符串写默认值 |
| 语言标识 | `IVGLocalizableString.GetLanguages()` | 返回 tuple，已设置语言为字符串（"2052"/"1033"），另有 "" 项 |
| 未封装 | `HasNonDefaultLangStrings` 等只读状态 | 已在读取结果中给出 has_default，其余未单独暴露 |

## 七之二十三、第二十五波已接入（2026-10-25，屏幕叠加层）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **叠加文字** | `cdr_view(action="overlay_text", text, x, y, color)` | ✅ CreateOnScreenText + SetTextAndPosition + SetTextColor + Show，无错误 |
| **叠加矩形** | `cdr_view(action="overlay_rect", x, y, x2, y2, color, width)` | ✅ CreateOnScreenCurve + SetPen + SetNoBrush + SetRectangle + Show |
| **叠加直线 / 圆** | `overlay_line` / `overlay_circle` | ✅ SetLine / SetCircle 调用成功 |
| **清除** | `cdr_view(action="overlay_clear")` | ✅ 隐藏 4 个叠加对象 |
| 坐标系 | 屏幕像素 | 文档坐标请先用 `to_screen` 换算 |
| 注意 | xor 绘制、不写入文档 | 窗口重绘/关闭后叠加层可能消失 |

## 七之二十四、第二十六波已接入（2026-10-25，应用事件队列）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **事件订阅** | `ComBridge._install_events`：生成 `DIVGApplicationEvents` sink 挂到 Application 连接点 | ✅ 连接后能收到 WindowActivate / DocumentNew / DocumentClose 等 |
| **事件读取** | `cdr_doc(action="events", limit=N)` | ✅ 新建文档后返回 count=2：WindowActivate + DocumentNew |
| **事件清空** | `cdr_doc(action="clear_events")` | ✅ 返回 cleared 数量并清空队列 |
| 事件范围 | DocumentNew/Open/Close、Before/AfterSave、Before/AfterExport、WindowActivate/Deactivate、SelectionChange、Start、PluginCommand | `ApplicationEvent` 高频事件忽略 |
| 关键坑 | TLB 路径 | `app.Path` 指向 `...\Draw\`，实际 TLB 在 `..\Programs64\vgcoreauto.tlb`；已按候选路径查找 |

## 七之二十五、第二十七波已接入（2026-10-25，文档级事件）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **文档事件订阅** | `ComBridge._install_doc_events`：每个 Document 挂 `DIVGDocumentEvents` sink | ✅ 创建矩形后捕获 `Doc:ShapeCreate` + `Doc:ShapeChange`；移动/删除/图层/页面变化同理 |
| **按需挂载** | `_with_doc` 首次解析文档时自动安装；DocumentNew/Open 事件也会触发安装 | ✅ 新开文档也能捕获图形变化 |
| **事件队列** | 与应用事件共用 `cdr_doc(action="events")` 队列，文档事件带 `Doc:` 前缀 | ✅ 可通过前缀区分 app/doc 事件 |
| 关闭清理 | DocumentClose 时卸载对应 sink | ✅ 避免重复订阅 |

## 七之二十六、第二十八波已接入（2026-10-25，导出选项）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **导出选项** | `cdr_export(..., export_options={...})` | ✅ `AntiAliasingType=2`、`Dithered=false`、`Transparent=true`、`ResolutionX/Y=120`、`MaintainAspect=true` 均被应用，PNG 正常生成 |
| 支持字段 | AntiAliasingType / Compression / Dithered / Transparent / MaintainAspect / MaintainLayers / Overwrite / Matted / MatteMaskedOnly / AlwaysOverprintBlack / UseColorProfile / ResolutionX·Y / SizeX·Y | 未识别字段静默忽略 |
| 结果 | 返回 `export_options` 已应用字段列表 | ✅ 便于确认 |
| 注意 | 复杂 CMYK 文档 + 透明选项可能导出失败 | 建议用干净 RGB 文档；失败自动回退默认选项 |

## 七之二十七、第二十九波已接入（2026-10-25，保存选项）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **保存选项** | `cdr_save(..., save_options={...})` / `cdr_save(as_copy=true, save_options=...)` | ✅ EmbedICCProfile/KeepAppearance/Overwrite/ThumbnailSize 应用后 .cdr 正常保存（13.7 KB） |
| 支持字段 | EmbedICCProfile / EmbedVBAProject / IncludeCMXData / KeepAppearance / Overwrite / ThumbnailSize / Version / Filter / Range | 未识别字段静默忽略 |
| 结果 | 返回 `save_options` 已应用字段列表 | ✅ 便于确认 |

## 七之二十八、第三十波已接入（2026-10-25，导入选项与链接位图）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **导入选项** | `cdr_import(..., import_options={...})` | ✅ `link_bitmap_externally=true`、`resample_dpi=96`、`maintain_layers=true` 均被应用 |
| **链接位图** | `link_bitmap_externally=true` | ✅ 导入后 `ExternallyLinked=true`，`LinkFileName` 指向源 PNG |
| **链接查询/更新** | `cdr_bitmap(action="link_info"/"update_link"/"resolve_link")` | ✅ link_info 读外部链接状态；update_link 重新加载成功 |
| 对象命名 | 导入后自动补 `cdr_xxx` 名字 | ✅ 便于后续按名操作 |

## 七之二十九、第三十一波已接入（2026-10-25，吸附点全家族）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **对象吸附点枚举** | `cdr_snap(action="list")` | ✅ 普通矢量对象 4 个 Type=2（边中点）；`by_kind` 汇总；文本对象同样有 4 个 |
| 按类型掩码过滤 | `cdr_snap(action="of_type", typeset=n)` | ✅ 掩码 1 用户 / 2 对象 / 4 包围盒 / 8 边 / 16 自由 / 32 自动；`typeset=4` 返回 9 个包围盒点 |
| 单个吸附点详情 | `cdr_snap(action="info", index=n)` | ✅ 类型/坐标/方向/可移动/可删除/可选/`GetPosition` 一致性 |
| **新增用户吸附点** | `cdr_snap(action="add_user", x, y, angle, use_direction)` | ✅ 生成 Type=1 点并带 GUID；x/y 是**文档坐标**（不是对象局部坐标） |
| 自动吸附点 | `cdr_snap(action="auto")` | ✅ `SnapPoints.Auto()` 返回 Type=32（实测是**方法**不是集合） |
| 包围盒定址 | `cdr_snap(action="bbox", index=1..9)` | ✅ 9 个包围盒方位点；越界由本服务拦成 `bad_argument` |
| 对象定址 | `cdr_snap(action="obj", index=1..4)` | ✅ 4 个对象点 |
| 边定址 | `cdr_snap(action="edge", index=段号, offset=0..1)` | ✅ `Edge(2,0.25)` 精确取到 (60, 27.5) |
| 按 GUID 反查 | `cdr_snap(action="user", reference=<GUID>)` | ✅ 用 `sps.User("3f2c…")` 取回同一点；传数字/对象都报「为参数 ID 指定的值无效」 |
| 集合内最近点 | `cdr_snap(action="near", x, y, typeset)` | ✅ `SnapPoints.FindClosest(类型,x,y)` |
| 页面级最近点 | `cdr_snap(action="find", x, y, typeset, hot_area)` | ✅ 热区内命中；**未命中会返回 Type=16 的“自由点”（坐标=查询点）**，本服务标记 `fallback=true, found=false` |
| 文档/应用级自由点 | `cdr_snap(action="free"/"create")` | ✅ `Document.CreateFreeSnapPoint` / `Application.CreateSnapPoint` 均返回 Type=16 |
| **移动吸附点** | `cdr_snap(action="move", dx, dy)` 或 `x, y` 绝对定位 | ✅ 相对走 `IVGSnapPointRange.Move`，绝对走 `IVGSnapPoint.SetPosition` |
| **删除吸附点** | `cdr_snap(action="delete", index|indices)` | ✅ 走 `IVGSnapPointRange.Delete`，集合 5 → 4 |
| 方向编辑 | `cdr_snap(action="direction", angle, uses_direction)` | ✅ `Direction=15` 回读 15.0；`UsesDirection` 可写 |
| 选择吸附点 | `cdr_snap(action="select")` / `clear_selection` | ✅ `CreateSelection` 后 `SnapPoints.Selection.Count=1`，清空回 0 |
| **范围批量** | `cdr_snap(action="range", range_op=…)` | ✅ Move / Delete / Add / Remove / RemoveByReference / Find / CreateSelection / AddToSelection / RemoveFromSelection / ChangeDirection / SetAutoSnap 全部可用 |
| 用户点属性 | `snap.User.ID` / `snap.User.AutoSnap` | ✅ GUID 与开关都可读写 |

> **第三十一波新增的两个 X8 怪癖**（已写进代码注释）：
> 1. **移动对象吸附点会把它变成用户吸附点**：Type 2 → 1，并从对象点段排到用户点段，集合总数不变。
> 2. `FindClosestSnapPoint` 在热区内没有匹配点时**不报错也不返回空**，而是返回一个
>    Type=16（自由）的吸附点，坐标等于查询点本身 —— 直接用会把“没吸附到”当成“吸附到了”。
>
> `AddUserSnapPointEx(ID, x, y, direction, useDirection)` 仍然不可用：尝试了 int / str(GUID) /
> `uuid.UUID` / `{GUID}` / IVGUserSnapPoint 对象 / IVGSnapPoint 对象 / VT_BSTR / VT_DISPATCH，
> 一律报「为参数 ID 指定的值无效」（`0x80070057`）。

## 七之三十、第三十二波已接入（2026-10-25，颜色对象与描边对象）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| 颜色对象创建/读 | `cdr_color(action="object_info")` | ✅ RGB `Type=5`、CMYK `Type=2`，`ToString()` 给出 `RGB255,USER,200,30,40,100,<GUID>` 形式 |
| 颜色名读写 | `cdr_color(action="object_name", name=…)` | ✅ 默认 `"unnamed color"`，写入后 `GetName()` 回读一致（`ToString` 不受影响） |
| 颜色比较 | `cdr_color(action="object_compare", color2=…)` | ✅ `IsSame` / `GetColorDistanceFrom`（同色 0.0、红蓝 647.0）；`IsSimilar` 对同色也返回 False（X8 行为，仅作参考） |
| 原地混合 | `cdr_color(action="object_blend", ratio=0.5)` | ✅ 红+蓝 50% → `RGB255,USER,128,0,128`；**BlendWith 原地改色、返回 None** |
| 取副本 | `cdr_color(action="object_copy")` | ✅ `GetCopy()` 后 `IsSame` 为真 |
| 颜色串赋值 | `cdr_color(action="object_string_assign", color2="CMYK,USER,…")` | ✅ 与 `ToString` 可往返 |
| 原地换模型 | `cdr_color(action="object_convert", model=…)` | ✅ RGB→CMYK(59,7,98,0)/Gray(Type 9)/HLS/HSB(Type 6)/Lab(Type 12)/CMY/YIQ 全通，**返回 None 但对象类型已变** |
| 固定调色板 | `cdr_color(action="object_fixed", mode=0/1/2)` | ✅ 得到 `CMYK255` 且标识分别为 `USER` / `TRU` / `PAN_PROC` + 各自 GUID |
| **分量直写** | `cdr_color(action="object_assign", model=…, values=[…])` | ✅ RGBAssign/CMYKAssign/GrayAssign/HLSAssign/HSBAssign/LabAssign/CMYAssign/YIQAssign/BWAssign/RegistrationAssign 全通（`registration` 得到 `REGCOLOR`） |
| **描边全量读** | `cdr_outline(action="info")` | ✅ 宽度/颜色串/线型 Index/线帽/连接/BehindFill/ScaleWithShape/笔尖角度·拉伸/斜接/对齐/箭头 |
| **描边写入** | `cdr_outline(action="set", changes={…})` | ✅ width/color/caps/join/nib_stretch/nib_angle/miter_limit/justification 全部写后即生效 |
| **SetProperties 全参** | `changes={"use_set_properties":true,…}` | ✅ 宽度 1.25 + 线帽 1 + 连接 1 + 斜接 9 全部落地（早先判「不可用」是传参错误） |
| 描边转对象 | `cdr_outline(action="convert_to_object")` | ✅ 返回曲线对象（`Type=3`），源对象被替换，`Shapes.Count` 不变 |
| 描边复制/比较 | `cdr_outline(action="copy"/"compare")` | ✅ `CopyAssign` 后 `CompareWith` 变 True（宽度/颜色都复制过去） |

> **第三十二波纠正与新怪癖**：
> 1. `Outline.SetProperties/SetPropertiesEx` **不是不可用**，而是 14/15 个参数必须给实参，
>    把 `None` 传给 int 形参才会 `TypeError: int() argument must be … not 'NoneType'`。
> 2. `Outline.PenWidth` 恒为 0（派生属性），`Outline.DashDotLength` 写入不生效。
> 3. `Color.IsSimilar(同色)` 返回 False，但 `IsSame` 为 True —— 判等要用 `IsSame`。
> 4. `Color.ConvertToPalette` / `SpotAssignByName` 报「服务器出现意外情况」；
>    `CopyAppearance` 报 `TypeError`；`UserAssign/UserAssignEx` 的参数其实是**窗口句柄**
>    （`ParentWindowHandle`），本质是弹「选择颜色」对话框，**绝不暴露**。

## 七之三十一、第三十三波已接入（2026-10-25，原生范围对齐与范围批量）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **对齐到指定点** | `cdr_align(to="point", x, y)` | ✅ `AlignRangeToPoint(cdrAlignLeft/Top, x, y)`，一次一轴，3 个对象整体左上角贴到 (150,200) |
| **对齐到页面中心** | `cdr_align(to="page_center")` | ✅ `AlignRangeToPageCenter` |
| **对齐到网格** | `cdr_align(to="grid")` | ✅ `AlignRangeToGrid` |
| **对齐到参照对象** | `cdr_align(to="shape", target2=…)` | ✅ `AlignRangeToShape(cdrAlignType, shape)` |
| **对齐到参照范围** | `cdr_align(to="range", target2=…)` | ✅ `AlignRangeToShapeRange(cdrAlignType, range)`（用参照对象建单元素范围） |
| 范围查询 | `cdr_shape_ops(op="range_info")` | ✅ Count / 名字 / type_ids / `CountAnyOfType` / `ExistsAnyOfType` / `FindAnyOfType`（传 list[int] 可用） |
| 范围增删成员 | `op="range_add"/"range_remove"`（`other=` 另一个对象） | ✅ 3→2→3 |
| 批量解锁 | `op="unlock_all"` | ✅ `ShapeRange.Unlock()`，解后 `Locked` 全 False |
| 清效果 | `op="clear_effect"`（`distance_a` = cdrEffectType） | ✅ `ShapeRange.ClearEffect(5=DropShadow)` |
| 炸开 | `op="break_apart_ex"` | ✅ `BreakApartEx()` 返回范围 |
| 整页递归解组 | `op="ungroup_all"` | ✅ 组合 1 个 → 解出 3 个，名字保持不变 |
| 带锚点变换 | `ShapeRange.SetPositionEx/GetPositionEx/RotateEx/SkewEx`（已实测） | ✅ `GetPositionEx(9=cdrCenter)` 回 `[50.0, 60.0]`；`RotateEx(50,60,30)` / `SkewEx(50,60,10,0)` 生效 |

> **第三十三波纠正与新怪癖**：
> 1. **`ShapeRange.UngroupAll` / `UngroupAllEx` / `UngroupEx` 全是空操作** —— 调完页面对象数
>    一点不变（实测 1 → 1）。`ungroup_all` 因此改成对整页组合体逐个 `Shape.Ungroup()` 递归解开，
>    实测 1 → 3 且对象名保留。
> 2. `ShapeRange.Sort(CompareExpression, …)` 的比较表达式语法没逆出来：`"width"` / `"@1.width"` /
>    `"s1.width < s2.width"` 全部报 `Identifier 'x' is not a method of 'global'/'pair'`。
> 3. `ShapeRange.AllExcluding([shape])` 报 `TypeError`（不接受对象数组）；
>    `SetSizeEx` 参数顺序不对会「无效的指定宽度或高度」，给对了又会把内容放大到几千毫米；
>    `StretchEx(x1,y1,x2,y2,bool)` 按点到点比例缩放，同样会瞬间放大 —— 三个都不暴露。
> 4. `GetOverprintFillState/OutlineState` 无论对象是否叠印都返回 1，语义未确认，不暴露。

## 七之三十二、第三十四波已接入（2026-10-26，矩形对象 / 曲线段 / 文字样式修正）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **矩形对象** | `cdr_rect(action="info")` | ✅ `CreateRect()` + `SetRect` 后 x/y/w/h 与 Left/Right/Top/Bottom/Center 全部正确 |
| 矩形平移/外扩 | `cdr_rect(action="offset"/"inflate")` | ✅ `Offset(5,5)` → x=15；`Inflate(1,1,2,2)` → x=9、w=43（负数内缩） |
| 矩形并/交 | `cdr_rect(action="union"/"intersect")` | ✅ 两个 40×30 矩形 → 并 70×50、交 10×20 |
| 点包含 / 复制 | `cdr_rect(action="contains"/"copy")` | ✅ 点 (30,30) 命中；副本尺寸一致 |
| **按矩形造形** | `cdr_rect(action="create_rectangle"/"create_ellipse")` | ✅ `CreateRectangleRect(rect, 圆角)` → 40×30 矩形；`CreateEllipseRect(rect,0,270,True)` → 扇形（Type 2） |
| **曲线段直建** | `cdr_create(type="curve_segment")` | ✅ `Layer.CreateCurveSegment` 一次给端点+两端控制柄长度/角度，得到 40×21 曲线 |
| **图层查找** | `cdr_layers(action="find", shape_name=…)` | ✅ `Layer.FindShape(name,0,0,True)` 命中返回 type_id/bbox；查不到回 `found=false` |
| 图层信息 | `cdr_layers(action="info")` | ✅ 含 `is_guides_layer`（实测 `IsGuidesLayer` 是**属性**，当方法调会 `TypeError`） |
| **字体表刷新** | `cdr_doc(action="refresh_fonts")` | ✅ `ForceUpdateFontTable()` 后字体表 2068 项，可修复「建文本静默回退默认字体」 |
| **粗斜体修正** | `cdr_create` / `cdr_modify` / `cdr_text` 的 bold/italic/underline | ✅ 改为 `cdrTriState(-1)` 与下划线枚举(1/0)，此前传 `True` 完全无效 |

> **第三十四波纠正的三个坑**：
> 1. **Bold/Italic 形参是 `cdrTriState`，true = -1**。传 Python `True`（=1）不但无效，
>    `Story.Bold = True` 还会**把粗体关掉**。下划线是另一套枚举：1=有、0/-1=无。
> 2. **字体表会过期**：长时间批量建文本后，`CreateArtisticText(font="Arial")` 会静默回退默认
>    字体（中文机=宋体）并丢掉 Bold/Italic。`ForceUpdateFontTable()` 立刻恢复，已做成兜底。
> 3. **中文默认字体文档改不了拉丁字体**：当文档默认字体是宋体时，`Story.Font = "Arial"`、
>    `Story.Bold = -1`、`TextRange.Font = …` 都不报错也不生效（Size/Underline 仍生效），
>    属 X8 字体替换限制 —— `cdr_text` 会如实返回 `applied`，不撒谎。

## 七之三十三、第三十五波已接入（2026-10-26，COM 单位换算 / 颜色工厂 / 定尺寸加页）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **COM 单位换算** | `cdr_doc(action="convert_units", value, from_unit, to_unit)` | ✅ `Application.ConvertUnits`：25.4 mm→inch = **1.0**；72 pt→mm = **25.4**；10 cm→mm = 100 |
| **颜色模型工厂** | `cdr_color(action="object_model", model=…, values=[…])` | ✅ HSB(Type 6) / CMY(Type 4) / YIQ(Type 11) / BW(Type 8) / Fixed(`CMYK255,PAN_PROC,…`) / ex(`CreateColorEx`) |
| **颜色串造色** | `cdr_color(action="object_from_string", color2="…")` | ✅ `Application.CreateColor("RGB255,USER,12,34,56,…")` → Type 5 |
| **定尺寸加页** | `cdr_pages(action="add", count, width_mm, height_mm)` | ✅ `AddPagesEx(2,100,150)`：页数 1→3，第 2 页尺寸正是 100×150 |
| **定尺寸插页** | `cdr_pages(action="insert", …, width_mm, height_mm)` | ✅ `InsertPagesEx(1,True,1,80,60)`：页数 3→4 |
| 命名箭头 | `Document.CreateArrowHead2/Ex/Ex2` | ✅ 由曲线生成 `IVGArrowHead`（可再 `ArrowHeads.Add` 入库） |
| 解析位图链接 | `cdr_doc(action="resolve_links")` | ✅ `ResolveAllBitmapsLinks()` 调用成功 |
| 文档设置快照 | `cdr_doc(action="save_settings"/"restore_settings"/"reset_settings")` | ⚠️ 调用无异常，但实测 `RestoreSettings` 并没有把改过的 Resolution 恢复（Tag 语义未逆出，标为实验性） |
| Unicode 转换 | `Application.ConvertToUnicode/ConvertFromUnicode` | ✅ 往返字符串不变（CodePage=0） |

> **第三十五波记录**：`Application.CreateProperties` 返回 `IVGProperties` 但 `Count=0` 且
> `Name/Add/Value` 在 pywin32 包装类里不存在（AttributeError），`IVGProperties` 的用法仍未打通；
> `Document.CreateShapeRangeFromArray` / `CreateCurveFromArray` 需要对象数组指针，暂未尝试。

## 七之三十四、第三十六波已接入（2026-10-26，节点拓扑与曲线合并）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **节点前后邻居** | `cdr_node(action="neighbors", node=n)` | ✅ `Next()` / `Previous()` **是方法**（当属性用会 `AttributeError: 'function' object`）；开放曲线首节点的 previous 会绕到末节点 |
| 节点间距 | `cdr_node(action="distance", node, end)` | ✅ `GetDistanceFrom`：(0,0)↔(40,20) = **44.721** |
| **两节点取平均位置** | `cdr_node(action="average", node, end)` | ✅ `AveragePositionWith(other, False)`：两端点都落到 (20,10) |
| **末端节点合并** | `cdr_node(action="join_nodes", node, end)` | ✅ `JoinWith`：3 节点 → 2 节点，曲线变闭合 |
| **曲线命中判定** | `cdr_path(action="is_on_curve", x, y, hot_area)` | ✅ `IsOnCurve`：节点处返回 1，远处返回 0 |
| 曲线副本 / 信息 | `cdr_path(action="curve_copy"/"curve_info")` | ✅ `GetCopy` 副本 2 节点 20mm；`GetCurveInfo` 返回 3 个 `com_struct(PositionX, PositionY, ElementType, NodeType, Flags)` |
| **曲线合并** | `cdr_path(action="merge_curve", other=…)` | ✅ `Curve.AppendCurve`：2 节点 → **4 节点 / 2 子路径**（同一个对象里） |
| **曲线焊接** | `cdr_path(action="weld_curve", other=…)` | ✅ `Curve.WeldWith` 返回**新曲线**（6 节点、长 60mm），不会自动替换原对象 |

> **第三十六波纠正的坑**：
> 1. **`IVGNode.Next/Previous` 是方法不是属性** —— 类型库里看着像属性（propget），
>    实际必须 `node.Next()` 调用，否则拿到的是 bound method。
> 2. `Node.ConnectWith` / `ExtendSubPaths` 在"两条不同曲线"的末端节点上仍然报
>    「指定的节点必须在另一条曲线上」/「无法连接曲线」，大概需要先在 UI 选择集里同时选中两条曲线；
>    同曲线内的 `JoinWith` 是可用的。
> 3. `Curve.WeldWith` **返回新对象**，不会替换原对象；要"原地合并"用 `AppendCurve`。

## 七之三十五、第三十七波已接入（2026-10-26，文字结构体 / 区间宽接口 / 样式继承）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **文字结构体快照** | `cdr_text(struct_info=true)` | ✅ 一次读出 `FontProperties`（Name/Size/Style/Underline/Overscore/Strikethru/Uppercase/Position/RangeKerning）、`SpaceProperties`（字符/词/行距·段前段后）、`AlignProperties`（对齐/左右缩进/首行缩进/词距上下限）、`HyphenationSettings`（自动断字/断字区/最短词长） |
| **区间前置/后置插入** | `cdr_text(range_action={"op":"insert_before"/"insert_after",…})` | ✅ `InsertBeforeWide(text, lang, charset, font)`：文本从「结构体测试」变「前结构体测试」 |
| **区间替换** | `range_action={"op":"replace"}` | ✅ `ReplaceWide`：「前结构体测试」→「改结构体测试」 |
| **区间折叠/重设** | `range_action={"op":"collapse"/"set_range"}` | ✅ `Collapse(true)` → [3,3]；`SetRange(1,4)` → [1,4] |
| **区间判等/包含** | `range_action={"op":"is_same"/"in_range"}` | ✅ `IsSame` → true；`InRange` → true |
| **区间行距** | `range_action={"op":"set_line_spacing"}` | ✅ `SetLineSpacing(1,150,0,0)` |
| **样式覆盖属性** | `cdr_style(action="override_props")` | ✅ 实测 `('primaryColor','secondaryColor','type','overprint','screenSpec')` |
| **样式属性取值** | `cdr_style(action="get_property", property=…)` | ✅ `GetProperty("primaryColor")` → `RGB255,USER,51,102,204,…` |
| 样式继承状态 | `cdr_style(action="inherited", property=…)` | ✅ `IsPropertyInherited` → False |
| 样式属性清除 | `cdr_style(action="clear_property", property=…)` | ✅ `ClearProperty` 返回布尔 |

> **第三十七波纠正的坑**：
> 1. `Text.FontProperties` 系列在 pywin32 里是**属性**：当方法调报
>    `TypeError: IVGStruct… object is not callable`；而**直接赋值 setter 会把字号写成 1.17e-38**
>    —— 所以结构体只做只读快照，写字号继续走 `font=/size=`。
> 2. 带下标的 `FontPropertiesInRange(start,count,indexing)` 能返回结构体但字段是垃圾值，不可信。
> 3. `TextRange.Evaluate(expr)` 的表达式语法未逆出（`len` 报 `Identifier 'len' is not a method of 'global'`）。
> 4. `Style.SetBasedOn(None)` 报「未找到样式 None」（要传样式对象）；`Style.Assign(shape)` 报类型不匹配。

## 七之三十六、第三十八波已接入（2026-10-26，结构体建文档 / Fill 工厂 / 命名箭头）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **结构体建文档** | `cdr_open(create_options={name,width_mm,height_mm,resolution,units})` | ✅ `CreateStructCreateOptions` + `CreateDocumentEx`：文档名 `MCP-Struct`、页 120×90mm、分辨率 150 一次定好 |
| **带选项打开** | `cdr_open(path, open_options={code_page})` | ✅ `CreateStructOpenOptions` + `OpenDocumentEx`；结构体只有 `CodePage` / `ColorConversionOptions` |
| **Fill 工厂批量上色** | `cdr_shape_ops(op="apply_fill", base_color=…)` | ✅ `Document.CreateUniformFill(IVGColor)` → `ShapeRange.ApplyFill(fill)`：3 个对象一次上色，`Fill.Type=1` |
| **命名箭头** | `cdr_doc(action="arrowhead_named", query=名称, register=true)` | ✅ `CreateArrowHead2(curve, "MCP箭头38")` 得到带名字的箭头，登记后箭头库 94→95 |
| **边界版矩形** | `cdr_create(type="rect_ltrb", left/top/right/bottom, radius)` | ✅ `Layer.CreateRectangle(8 参)`：(0,30)-(40,0) 圆角 3 |
| **边界版椭圆/扇形** | `cdr_create(type="ellipse_ltrb", …, start_angle, end_angle, pie)` | ✅ `Layer.CreateEllipse(7 参)`：Type 2 |
| **完整多边形** | `cdr_create(type="polygon_full", …, sides, subpaths, complexity, star, star_complexity, max_complexity)` | ✅ `Layer.CreatePolygon(10 参)`：6 边星形 Type 4 |

> **第三十八波记录**：`Application.CreatePaletteColor(name, index)` / `CreateSpotColor(name, tint, ?)`
> 不是"新建颜色"，而是从**已存在的**调色板/文档专色里取（分别报「找不到颜色」「未找到颜色」）；
> 新建专色仍走 `cdr_doc(action="spot_color")`。
> `CreateStructCreateOptions` 里的 `Units=3`（毫米）写入后文档 `Unit` 仍报 1（英寸），
> 但 `PageWidth/PageHeight` 按毫米生效 —— 尺寸不受影响，只影响 `doc.Unit` 的读数。

## 七之三十七、第三十九波已接入（2026-10-26，形状矩阵 / 扭曲 / 文字内嵌）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **形状矩阵读写** | `cdr_shape_ops(op="matrix", params={"matrix":[d11,d12,d21,d22,tx,ty]})` | ✅ `SetMatrix` 写 `[1,0,0,1,0,0]` 后 `GetMatrix` 精确回读；初始矩阵 `[1,0,0,1,-850000,-285000]`（内部单位） |
| **类型集合判定** | `cdr_shape_ops(op="is_type", params={"types":[1,2]})` | ✅ `IsTypeAnyOf([1,2])`=True、`([2])`=False（传 list[int] 可用） |
| **层序判定** | `cdr_shape_ops(op="in_front_of", other=…)` | ✅ `OrderIsInFrontOf`：后建的 B 在前 → True |
| **对象命中判定** | `cdr_shape_ops(op="is_on_shape", x, y)` | ✅ `IsOnShape`：内部=2、边界=1、外部=0 |
| **文字内嵌形状** | `cdr_shape_ops(op="place_text", other=文字对象)` | ✅ `PlaceTextInside`：文字并进矩形，返回 type 6 且 `Text.Story.Text` 保留 |
| **复制/克隆成范围** | `op="duplicate_range"` / `"clone_range"` | ✅ `DuplicateAsRange(8,8)` 页面 4→5；`CloneAsRange` 同样返回 1 个对象 |
| **扭曲效果** | `op="distort", params={"kind":"pushpull"/"twister"/"zipper", …}` | ✅ `CreateTwisterDistortion(90,130,30)` 后 bbox 改变、`Effects.Count=1`；pushpull/zipper 同样生效 |

> **第三十九波记录**：`Shape.GetMatrix` 的 tx/ty 是 X8 **内部单位**（初始值 -850000 量级），
> 写 `tx=5` 的位移极小 —— 矩阵读写精确但**不适合做毫米级平移**，平移仍用 `cdr_modify(move=…)`。
> 另外用 `SetMatrix` 改过矩阵的形状，其 `IsOnShape` 判定会失效（内部坐标映射被重置），
> 实测时要把命中判定放在改矩阵之前。

## 七之三十八、第四十波已接入（2026-10-26，参考点定位 / 吸附反查 / 位图化）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **按参考点定位** | `cdr_shape_ops(op="position_ex", params={"ref":9,"x":70,"y":80})` | ✅ `SetPositionEx(9=中心,70,80)` 后 bbox 变 (50,65)，中心正好 (70,80)；读模式返回 `[x,y]` |
| 参考点枚举 | 同上 `params.ref` | ✅ 1 右上 / 2 上中 / 3 左上 / 4 左中 / 5 左下 / 6 下中 / 7 右下 / 8 右中 / 9 中心 |
| **按引用反查吸附点** | `op="find_snap", params={"reference": "…"}` | ✅ `Shape.FindSnapPoint(ReferenceData)` → Type 2 吸附点，坐标 (70,95) |
| **链接对象查询** | `op="linked", params={"link_type":3}` | ✅ `GetLinkedShapes(0/3/8)` 可调用（普通矩形返回 0 个） |
| **带参位图化** | `op="rasterize_ex", params={"mode":4,"resolution":150,…}` | ✅ `ConvertToBitmapEx` 把矢量对象原地换成 Type 5 位图（对象数不变，`Bitmap` 可取） |
| 叠印状态 | `op="overprint_state"` | ⚠️ outline 状态 2→4 随开关变化，**fill 状态恒为 1**；以布尔属性 `OverprintFill/Outline` 为准 |

> **第四十波记录**：`Shape.RemoveFromContainer(1)` 报「参考的对象已不再存在」（需要真正的容器关系）；
> `Shape.Next/Previous` 在类型库里是**带 3 个参数的 propget**，当属性用取不到，z 序请用页面枚举。
> `GetOverprintFillState` 恒返回 1，不能用来判断填充叠印。

## 七之三十九、第四十一波已接入（2026-10-26，任意节点/线段集合 + 位图构造器）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **任意节点集合** | `cdr_node(action="arbitrary", params={"op":…,"nodes":[1,3,5]})` | ✅ `Application.CreateNodeRange()` + `Add(node)` 凑出**不连续**节点集合：Move(0,10) 后 1/3/5 号节点一起抬高 10mm，`GetBoundingBox()` 返回 `[0,10,80,0]` |
| 集合级变换 | 同上 `op=rotate/skew/stretch` | ✅ `Rotate(30,false,40,10)`、`Skew`、`Stretch` 均生效（bbox 改变） |
| 集合级整形 | 同上 `op=autoreduce/smoothen/set_type/delete/fillet/chamfer/scallop/select` | ✅ 全部无异常（与 `Nodes.Range(a,b)` 的连续范围互补） |
| **任意线段集合** | `cdr_path(action="segment_range", params={"op":…,"segments":[1,2]})` | ✅ `CreateSegmentRange()` + `Add(segment)`：`AddNode()` 让节点 5→7；`SetType`、`CreateSelection`、`RemoveAll` 均可用；`FirstSegment/LastSegment` 可读 |
| **位图构造器** | `cdr_create(type="bitmap_at"/"bitmap_rect", …)` | ✅ `Layer.CreateBitmap2(x,y,w,h,Image,None)` 与 `CreateBitmapRect(IVGRect,Image,None)` 都得到 Type 5 位图（20mm / 25mm） |

> **第四十一波记录**：
> 1. `NodeRange.Move(dx,dy,Index,bool)` 的第三个参数是**范围内起始索引**，传 0 会报
>    「参数 Index 超出范围（1~N）」；传 1 正常。
> 2. `NodeRange.GetBoundingBox()` 是 5 个出参，pywin32 返回 `[成功, x, y, w, h]`。
> 3. `SegmentRange.Remove(1)` 的行为与直觉不符（2 条删 1 条后 Count 反而变 3），
>    所以只暴露 `RemoveAll` 与 `Add`，`Remove` 未封装。
> 4. `Layer.CreateBSpline(Curve)` 报「类型不匹配」；`CreateCustomEffect/CreateCustomDistortion`
>    报「未装载此效果」（需要装插件效果）。

## 七之四十、第四十二波已接入（2026-10-26，对象列表建范围 / 范围位图化 / 箭头变体）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **对象列表直接建范围** | `cdr_shape_ops(op="range_from_list", params={"targets":[…]})` | ✅ `Document.CreateShapeRangeFromArray([shape,…])` **接受 Python list**：Count=2、`Exists(first)`=True，`Item(2)` 名字正确 |
| 范围整批旋转 | 同上 `params.then="rotate_ex"` + `angle/cx/cy` | ✅ 建完范围立刻 `RotateEx` |
| **整批位图化** | `cdr_shape_ops(op="range_bitmap", params={"mode":4,"resolution":96})` | ✅ `ShapeRange.ConvertToBitmapEx(8 参)` → Type 5 位图 |
| **指派文档色彩上下文** | `cdr_doc(action="assign_color_context", palette_index=1)` | ✅ `ColorProfiles.Item(1).CreateColorContext()` → `Document.AssignColorContext(ctx)`，RenderingIntent 可读（3） |
| **样式表设为默认** | `cdr_doc(action="stylesheet_default")` | ✅ `SaveStyleSheetAsDefault()` |
| **命名箭头变体** | `cdr_doc(action="arrowhead_ex", target=曲线, query=名称, x, y)` | ✅ `CreateArrowHeadEx2` 得到名字+中心点，登记后箭头库 95→96；`CreateArrowHeadOptions` 的 Length/Offset 可读（40 / [1,0]） |

> **第四十二波记录**：
> 1. `ShapeRange.CopyPropertiesFrom(shape, mask)` 报「执行方法 coVGShape::CopyPropertiesFrom 时发生意外」，
>    范围级不可用；属性复制请用 `cdr_modify(changes={"fill_from"/"outline_from"})`。
> 2. `CreateShapeRangeFromArray` 是**少数能接受 Python list 的数组入参**（点集类仍然 TypeError）。
> 3. `CreateArrowHeadEx`（无名字版）返回的箭头 `Name` 为空，要名字用 `CreateArrowHeadEx2`。

## 七之四十一、第四十三波已接入（2026-10-26，调色板选项 / 高级粘贴 / 两种导出 API）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **调色板选项转换** | `cdr_bitmap(action="paletted_options", num_colors, dither_type, palette_index)` | ✅ `CreateStructPaletteOptions()` + `Bitmap.ConvertToPaletted2(opts)`：Mode 4→3，NumColors=16；挂上 TRUMATCH（2092 色）后再转也成功 |
| **高级粘贴** | `cdr_layers(action="paste_ex")` | ✅ `Application.CreateStructPasteOptions()` → `Layer.PasteEx(opts)` 返回 ShapeRange（pasted=1） |
| **Document.Export** | `cdr_export(..., export_api="document")` | ✅ `Document.Export(FileName, Filter, Range, Options, PaletteOptions)` 落盘 15.7 KB |
| **Document.ExportBitmap** | `cdr_export(..., export_api="bitmap")` | ✅ `ExportBitmap(16 参)` 返回 `ICorelExportFilter`，**必须调 `Finish()` 才写盘**（否则文件不存在） |

> **第四十三波记录**：
> 1. `ExportBitmap` 与 `ExportEx` 一样返回 filter 句柄，**不调 `Finish()` 就不会生成文件**——
>    第一次实测「无异常但文件不存在」就是这个原因。
> 2. `Layer.PasteSpecial("CorelDRAW 图形", …)` 返回 False（格式串不匹配），未暴露；
>    `Document.LogCreateShape(shape)` 报 com_error。
> 3. `CreateStructPaletteOptions` 的 `Palette` 可以直接赋已打开的调色板对象
>    （`Palettes.OpenFixed(1)` = TRUMATCH，2092 色）。

## 七之四十二、第四十四波已接入（2026-10-26，颜色上下文 / 新窗口 / 渐变变换）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **颜色上下文读取** | `cdr_doc(action="color_context_info")` | ✅ RenderingIntent=3、BlendingColorModel=1、3 个配置文件、名称串 `sRGB IEC61966-2.1,Japan Color 2001 Coated,Dot Gain 15%`；RGB/CMYK/Gray 配置文件各自可读 |
| **上下文复制** | `cdr_doc(action="color_context_copy")` | ✅ `GetCopy()` + `IsSame()` → True |
| **上下文合并** | `cdr_doc(action="color_context_merge", palette_index=1)` | ✅ `ColorContext.Merge(other)` → True（合并后 `IsSame` 为 False，符合"合并≠相同"） |
| **渐变变换矩阵** | `cdr_doc(action="fountain_transform", target=…)` | ✅ `FountainFill.GetTransformations()` 用 4 个出参返回 `[1.0, -0.0, 0.0, 1.0]` |
| **新窗口** | `cdr_view(action="new_window")` | ✅ `Application.ActiveWindow.NewWindow()` → Windows.Count 1→2 |
| **纹理透明** | `cdr_modify(changes={"transparency":{"type":"texture","texture":"BARK","library":"Textures"}})` | ✅ `ApplyTextureTransparency("BARK","Textures",0,0)` 返回 IVGTextureFill，`Transparency.Type` 变 4 |

> **第四十四波记录**：`Document.NewWindow()` 不存在（`AttributeError`），新窗口必须走
> `Application.ActiveWindow.NewWindow()`；`FountainFill.GetTransformations` 是 4 个出参
> （pywin32 返回 tuple），不是返回矩阵对象。

## 七之四十三、第四十五波已接入（2026-10-26，SetPropertiesEx / 文字 BeginEdit）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **SetPropertiesEx（15 参）** | `cdr_outline(action="set", changes={"use_set_properties_ex":true,…})` | ✅ `Outline.SetPropertiesEx(2.5, None, color, …, justification=1)` → 宽度 2.5、LineCaps 1、Justification 1 全部落地（比 SetProperties 多的就是 Justification） |
| **文字进入编辑态** | `cdr_text(begin_edit=true)` | ✅ `Text.BeginEdit()` 生效，`applied` 回 `["begin_edit"]` |

> **第四十五波记录（都进了不可用清单）**：
> 1. `Document.CreateFill("CMYK,…")` / `CreateOutline("…")` 传字符串只得到 `Type=0`、宽度 0 的空对象 ——
>    字符串工厂无效，填充/描边仍走 `cdr_modify` / `CreateUniformFill`。
> 2. `Page.CustomCommand("", "", [])` 报「未装载此命令」。
> 3. `Style.SetBasedOn(父样式)` 无法稳定验证：`CreateStyleFromShape` 连续创建两个样式后
>    `Styles.Find("名字")` 找不到，`Styles` 列表只显示一个 `name=null` 的条目，
>    所以 `set_based_on` 未暴露（派生样式仍用 `create(based_on=…)`）。

## 七之四十四、第四十六波已接入（2026-10-26，CorelScriptTools 与逐项定性）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **VBA 辅助对象** | `cdr_doc(action="script_tools", value=…)` | ✅ `Application.CorelScriptTools()` → `ICorelScriptTools`：`FromInches(1)=254000`、`FromPoints(1)=3527.778`（内部单位 1/10000 mm）、`ToInches/ToPoints` 往返一致、`ASin(1)=1.571`（弧度）、`BuildDate(2026,10,26,12,0,0)="2026-10-26 00:00:00+00:00"` |
| **剩余方法逐项定性** | 见「三之补、逐项定性」表 | ✅ 覆盖 `IVGApplication`/`IVGDocument`/`IVGPage`/`IVGText` 里 20 个未覆盖方法，含 4 个**会挂起**的危险项 |

> **第四十六波最重要的产出是「危险清单」**：`AddPluginCommand` / `RemovePluginCommand`（300 s 卡死）、
> `GetUserArea` / `GetUserClick`（20 s 挂起）、`CreateCurveFromArray`（无返回）——
> 这些方法一旦被 Agent 调到，会把 CorelDRAW 一起拖死，文档里已标 🚫。
> 另外确认了 `IVGText` 的结构体**写入端**在 pywin32 下彻底不可用：
> 属性赋值会写坏字号，直接 `_oleobj_.Invoke` + 各种 `VARIANT` 包装都报「类型不匹配」
> （`VT_USERDEFINED` 参数无法 marshaling）。

## 七之四十五、第四十七波已接入（2026-10-26，宽文本 / 节点坐标搬运 / 整批描边）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **宽文本构造** | `cdr_create(type="text_wide"/"paragraph_wide", language_id, charset, …)` | ✅ `CreateArtisticTextWide` / `CreateParagraphTextWide` 都能建出文字对象（charset=134 等显式传入） |
| **节点坐标搬运** | `cdr_node(action="copy_point", node=源, end=目标)` | ✅ `Node.GetPoint()` 返回 IVGPoint，`SetPoint(point)` 把目标节点搬到源节点坐标 |
| **整批描边 14 参** | `cdr_shape_ops(op="range_outline", params={"width":…,"line_caps":…})` | ✅ `ShapeRange.SetOutlineProperties(14 参)`：宽度 3.0 落到范围内所有对象 |
| **整批描边 15 参** | 同上 `params={"with_justification":true,"justification":1}` | ✅ `SetOutlinePropertiesEx`：宽度 2.0 + Justification 1 |
| **对齐页边** | `cdr_align(to="page_edge")` | ✅ `ShapeRange.AlignRangeToPage(cdrAlignType)`：axes `[1]`（左对齐页边） |
| **显式导入过滤器** | `cdr_import(filter_id=…)` | ✅ `Layer.ImportEx(path, Filter, Options)` 可调用；但 **802 是导出过滤器号**，传给导入会静默不导入（shapes_added=0） |

> **第四十七波记录**：
> 1. **导入/导出过滤器号不通用**：PNG 导出号 802 传给 `ImportEx` 时调用成功但**一个对象都不导入**；
>    传 0 让它自动判断才正常。X8 没有列出导入过滤器号的 API，所以 `filter_id` 只做透传并标注。
> 2. `ShapeRange.DeleteItem(1)` 会把**文档里的对象一起删掉**（范围 2→1、页面 2→1）；
>    只从范围移除请用 `RemoveRange`。
> 3. `Shape.CompareToEx` 各种 condition 都报「发生意外」，未封装。

## 七之四十六、第四十八波已接入（2026-10-26，颜色 Assign 家族 / 图层复制）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **固定调色板赋值** | `cdr_color(action="object_assign", model="fixed", values=[palette_id,index,tint])` | ✅ `FixedAssign(2,5,100)` → `CMYK255,PAN_PROC,0,0,179,0,100,<guid>` |
| **脚本式分量赋值** | 同上 `model="script", values=[模型号,v1,v2,v3,v4]` | ✅ `CorelScriptAssign(2,10,20,30,40,0,0,0)` → `CMYK,USER,10,20,30,40` |
| **按调色板取色** | `cdr_color(action="object_palette", color2=<调色板 Identifier>, values=[序号])` | ✅ `PaletteAssign(guid, 5)` → **`PANTONE,PAN_SPOT,5,100,<guid>`**（真正拿到专色） |
| 颜色失效 | `cdr_color(action="object_invalidate")` | ✅ `Invalidate()` 后 `ToString` 变 `CSINVALID,USER,100,…` |
| **复制到指定图层** | `cdr_shape_ops(op="copy_to_layer_range", params={"layer":"目标层"})` | ✅ `Shape.CopyToLayerAsRange(layer)` → 目标层上 1 个新对象（名字自动补） |
| **对齐到形状范围** | `cdr_shape_ops(op="align_to_range", other=参照对象, params={"align_type":1})` | ✅ `AlignToShapeRange(1, range, 0)`：目标 bbox 从 (0,0) 移到 (40,0) |

> **第四十八波记录**：
> 1. `PaletteAssign` / `SpotAssignByName` 的第一个参数是**调色板 GUID 字符串**（`Palette.Identifier`），
>    不是调色板名；GUID 正确时能取到真正的专色（`PAN_SPOT`）。
> 2. `SpotAssign/SpotAssignByName` 在**不含专色**的调色板（如默认调色板）上会报「未找到指定的对象」——
>    要先 `palette_open_fixed` 打开 PANTONE 等专色调色板。
> 3. `Shape.ReplaceWith(另一个形状)` 报 com_error（需要"虚拟形状"），未封装。

## 七之四十七、第四十九波已接入（2026-10-26，颜色分量读取 / 转调色板 / 按掩码复制属性）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **颜色分量读取** | `cdr_color(action="object_components", model=2/5/3, color=…)` | ✅ `CorelScriptGetComponent(模型号, 0×7)` 返回 `(状态码, v1..v7)`：CMYK `(5002,10,20,30,40,…)`、RGB `(5005,11,22,33,…)`、Gray `(5009,50,…)`；状态码 = 5000 + 模型号 |
| **转指定调色板** | `cdr_color(action="object_convert_palette", color2=<调色板 GUID>)` | ✅ `ConvertToPalette(guid)`：RGB(120,60,30) → TRUMATCH `CMYK255,TRU,…`；PANTONE → `PAN_SPOT,619`。传数字序号会报「找不到调色板 N」 |
| **按掩码复制属性** | `cdr_shape_ops(op="copy_props", other=源对象, params={"mask":4})` | ✅ `Shape.CopyPropertiesFrom(源, 掩码)`：掩码 4/8/15 会把填充复制过来，1/2 不影响填充 |

> **第四十九波是一次重要纠正**：文档此前一直写「`IVGColor` 不暴露任何分量属性」——
> 其实 **`CorelScriptGetComponent` 就是读分量的入口**，只是名字不像。
> 至此颜色对象的分量「读 + 写」都通了（写走 `*Assign`，读走 `CorelScriptGetComponent`）。
> 另：`Shape.CopyPropertiesFrom` 在**形状级**可用（范围级仍报「执行方法时发生意外」）；
> `Layer.CreateCustomShape/CreateToolShape` 分别报「未装载此形状」/ com_error，不可用。

## 七之四十八、第五十波已接入（2026-10-26，形状斜切 / 子路径追加与抽出 / 焊接落地）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **形状级斜切** | `cdr_shape_ops(op="skew_ex", params={"angle_x":10,"cx":10,"cy":5})` | ✅ `Shape.SkewEx(10,0,10,5)`：bbox 宽度 20→21.763、x 偏到 -0.882 |
| **子路径追加曲线段** | `cdr_path(action="append_curve_segment", params={x,y,start_ctrl_length,…})` | ✅ `SubPath.AppendCurveSegment` 7 参：节点 3→4，返回 IVGSegment |
| **抽出子路径** | `cdr_path(action="extract_subpath", params={"subpath":1})` | ✅ `SubPath.Extract(shape)` 返回 `(新 Shape, …)`，抽出后是 type 3 曲线、页面对象 +1 |
| **焊接结果落地** | `cdr_path(action="weld_to_shape", other=另一条曲线)` | ✅ `WeldWith` → `BindToDocument` → `Layer.CreateCurve`：游离曲线变回文档里的真形状（type 3、4 节点） |

> **第五十波补齐了焊接流程的最后一环**：`Curve.WeldWith` 以前只返回"游离曲线"（不在文档里），
> 现在用 `BindToDocument` + `Layer.CreateCurve` 能把它落成真形状 —— 这是前面几波一直缺的收尾步骤。
> 另：`Shape.ApplyTransformMatrix` / `Curve.ApplyTransformMatrix` 仍不可用，因为矩阵工厂在
> `IVGMathUtils` 上而 `Application.MathUtils` 根本不存在（CLSID 未注册）；
> 矩阵变换继续用 `Shape.SetMatrix` + `cdr_modify(affine=…)`。

## 七之四十九、第五十一波已接入（2026-10-26，结构体工厂写入 / 按模型取配置）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **结构体工厂写段落格式** | `cdr_text(paragraph_format={"alignment":3,"left_indent":5,"character_spacing":150,"line_spacing":180,"hot_zone":8})` | ✅ `CreateStructAlignProperties`+`SetAlignProperties`、`CreateStructSpaceProperties`+`SetSpaceProperties`、`CreateStructHyphenationSettings`+`SetHyphenationSettings`：`applied` 回 `align_struct / space_struct / hyphenation_struct`，读回 Alignment=3、LeftIndent=5.0、CharacterSpacing=150、LineSpacing=180 |
| **按模型取色彩配置** | `cdr_doc(action="color_context_info")` 的 `profiles_by_model` | ✅ `ColorProfile(0/1/2)` → sRGB IEC61966-2.1 / Japan Color 2001 Coated / Dot Gain 15% |
| **风格色改写** | `cdr_color(action="object_modify_style_color", color=…, color2=…)` | ✅ `ModifyColorStyleColor(目标色, False)` 调用成功 |
| **无名箭头变体** | `cdr_doc(action="arrowhead_unnamed", target=曲线, x, y)` | ✅ `CreateArrowHeadEx` 生成箭头但 `Name` 为空（要命名用 `arrowhead_ex` 的 Ex2 通道） |

> **第五十一波记录**：
> 1. 文本结构体的 `Set*Properties(0, 结构体)` 路径**只对段落文本生效**；
>    对艺术字调用报「该参数需要段落文本」——这正是第 37 波误判"结构体不可写"的原因之一。
> 2. `ConvertToUnicode` / `ConvertFromUnicode` 被 pywin32 的 BSTR 边界弄坏
>    （GBK 字节串 → `?D??`），已列入不可用清单，码页转换请用 Python `bytes.decode`。
> 3. `SpotAssign(调色板, 专色ID, tint)` 在 TRUMATCH 上仍报「未找到指定的对象」（该调色板没有专色）。

## 七之五十、第五十二波已接入（2026-10-26，按偏移取点 / 定位线段）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **按偏移取点** | `cdr_path(action="point_at_offset", params={"offset":0.5,"normalized":true})` | ✅ `SubPath.GetPointPositionAt(0,0,偏移,类型)` 返回 `(X,Y)`：半程得 `(20.175, 63.217)`（弧长 42.07 的贝塞尔）；类型 0=绝对长度 mm、1=归一化 0~1 |
| **按偏移定位线段** | `cdr_path(action="segment_at_offset", params={"offset":10.0})` | ✅ `FindSegmentOffset(10.0, None, 0, 0)` → `(true, IVGSegment, 0.297, 0.0)`，可拿到线段序号与参数偏移 |

> **第五十二波记录**：`IVGSubPath` 至此只剩 `PutCurveInfo` 未覆盖（数组入参不可用），
> 已达 **97%**。另外 `Layer.CreateFreeConnector` / `CreateSymbol` 与 `Application.SymbolLibraries`
> 实测都不可用（本机没有符号库、连接器需要文档级状态），已写入不可用清单。

## 七之五十一、第五十三波已接入（2026-10-26，B 样条构造 / 整批套用描边）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **B 样条构造** | `cdr_create(type="bspline", points=[[x,y],…], closed=false, to_curves=false)` | ✅ `Document.CreateBSpline(控制点数, 闭合)` → `IVGBSpline`，逐点 `InsertControlPoint(序号,x,y,Clamped)`（**两端必须 Clamped=True**，否则报「An end point's Clamped…」），再 `Layer.CreateBSpline(bs)` → **Type 21** 样条对象 |
| **B 样条转曲线** | 同上 `to_curves=true` | ✅ `ConvertToCurves()` → **Type 3**、8 个节点（4 控制点的三次样条） |
| **整批套用描边** | `cdr_shape_ops(op="range_outline_apply", other=描边来源, params={"width":4})` | ✅ `ShapeRange.ApplyOutline(Outline)`：`Outline.GetCopy()` 改宽度后整批覆盖选区，读回宽度 4.0 |

> **第五十三波记录**：`Layer.CreateBSpline` 的参数**不是曲线**而是 `IVGBSpline` 定义对象 ——
> 这就是第 32 波「CreateBSpline 类型不匹配」的真正原因（当时传的是 Curve）。
> 完整链路是 `Document.CreateBSpline` → `InsertControlPoint` × N → `Layer.CreateBSpline`。
> 另外 `SHAPE_TYPES` 补了 `21: "bspline"`。

## 七之五十二、第五十四波已接入（2026-10-26，同级遍历 / 解组家族纠正）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **同层相邻遍历** | `cdr_shape_ops(op="sibling", target=…, params={"direction":"next","steps":2})` | ✅ `Shape.GetNext(方向)` / `GetPrevious(方向)` 各取 1 个参数，返回同层相邻 IVGShape；正向走 2 步得到 2 个名字、反向 1 步得到 1 个 |
| **解组家族** | `cdr_shape_ops(op="ungroup_all" / "ungroup_all_ex" / "ungroup_ex")` | ✅ 三个都真实生效：`UngroupAll` 页码 1→2、`UngroupAllEx` 1→3 并返回 ShapeRange、`UngroupEx` 同样返回范围；选区里若是组合体**内部**对象会自动改为对整个页面的组合体动手 |

> **第五十四波是一次重要纠正**：`IVGShapeRange.UngroupAll/UngroupAllEx/UngroupEx` 早在第 30 波
> 就被判成「空操作」，其实当时是**拿组合体内部的对象去解组**（自然没反应）。
> 换正确对象后三者全部生效 —— 这也解释了为什么 `Shape.Ungroup()` 一直能用而它们"不能用"。
> 同理 `Shape.GetNext/GetPrevious` 之前记的「`Shape.Next` 是 3 参 propget」说的是**另一个属性**，
> `GetNext/GetPrevious` 只需 1 个方向参数。
>
> 另：本轮跑全量冒烟时 CorelDRAW 有一次在 `[15t] Revert` 用例后**进程直接崩溃**
> （RPC 服务器不可用），重启后同一条命令 474/474 全通过 —— 属偶发，`Revert` 本身已标记为破坏性操作。

## 七之五十三、第五十五波已接入（2026-10-26，连接器 / 线性·角度标注）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **连接器** | `cdr_create(type="connector", x1,y1,x2,y2)` | ✅ `Layer.CreateConnector(IVGSnapPoint, IVGSnapPoint)` → **Type 19**，bbox 40×20 |
| **连接器（带类型）** | `cdr_create(type="connector_ex", connector_type=0, …)` | ✅ `CreateConnectorEx(类型, 起, 止)` → Type 19 |
| **直角连接器** | `cdr_create(type="right_angle_connector", corner_radius=5, …)` | ✅ `CreateRightAngleConnector(起, 止, 5)` → Type 19 |
| **线性标注** | `cdr_create(type="linear_dimension", x1,y1,x2,y2, text_x, text_y, precision, prefix, suffix, …)` | ✅ `CreateLinearDimension` **22 参**全给实参 → **Type 14**，bbox 40 宽 |
| **角度标注** | `cdr_create(type="angular_dimension", cx,cy,x1,y1,x3,y3, …)` | ✅ `CreateAngularDimension` **18 参** → Type 14 |

> **第五十五波记录**：
> 1. 这一族的"堵点"是端点类型 `IVGSnapPoint` —— 它由 **`Application.CreateSnapPoint(x, y)`** 产出
>    （不是 `IVGPoint`，也没有 `x/y` 属性可写）。拿到它之后，连接器三件套与两种标注一次全通。
> 2. `SHAPE_TYPES` 补了 `14: "dimension"`、`19: "connector"`。
> 3. **环境经验**：连续两次全量冒烟分别在 `[15t] Revert` 与 `[15n] paste` 崩掉 CorelDRAW
>    （RPC 服务器不可用）；**杀掉残留的 CorelDRW 进程再跑，478/478 全通过** ——
>    说明崩的是上一次遗留的坏实例，不是被测代码。跑全量前务必确认没有残留进程。

## 七之五十四、第五十六波已接入（2026-10-26，曲线映射到骨架线）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **映射到骨架线** | `cdr_path(action="map_to_stroke", target=曲线, other=骨架线对象, params={"scale_y":5})` | ✅ `Curve.CreateCurveMappedToStroke(子路径, ScaleY, SelfWeld)` → 新 Curve，`Layer.CreateCurve` 后 Type 3、15 节点 |
| **带参考线映射** | 同上 `params={"use_reference_line":true}` | ✅ `CreateCurveMappedToStrokeAndReferenceLine(子路径, IVGPoint, IVGPoint, ScaleY, SelfWeld)` → Type 3、16 节点 |

> **第五十六波又纠正一条**：这两个方法早在第 32 波就被判成「类型不匹配」。
> 原因是 **`Stroke` 参数要的是 `IVGSubPath`（单条子路径），不是整条 `IVGCurve`**；
> 参考线版的两个点是 `IVGPoint`（`Node.GetPoint()` 拿）。换对类型后一次成功 ——
> 这是"艺术笔/骨架映射"类效果的底层入口。

## 七之五十五、第五十七波已接入（2026-10-26，文字区间结构体纠正）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **区间结构体读取** | `cdr_text(range_struct={"start":1,"count":5,"indexing":0})` | ✅ `FontPropertiesInRange` / `SpacePropertiesInRange` / `AlignPropertiesInRange` / `HyphenationSettingsInRange` 四个都能读：`font=宋体/14`、`space.CharacterSpacing=200`、`align.Alignment=2` |
| **区间结构体写入** | `cdr_text(paragraph_format={"range_start":1,"range_count":5,"character_spacing":200,"alignment":2})` | ✅ `SetAlignPropertiesInRange` / `SetSpacePropertiesInRange` / `SetHyphenationSettingsInRange` 全部生效，`applied` 回 `align_struct_in_range / space_struct_in_range`，读回 200.0 / 2.0 |

> **第五十七波是又一次"翻案"**：第 37 波把 `*InRange` 结构体判成"字段全是垃圾，不可用"，
> 其实是两个错误叠加：
> 1. **方法名没有 `Get` 前缀** —— 正确写法是 `text.FontPropertiesInRange(...)`，
>    不是 `GetFontPropertiesInRange`（后者根本不存在，之前的 AttributeError 被当成了"不可用"）；
> 2. **`StartIndex` 是 1 基** —— 传 0 才返回 `Size=0 / Style=16` 的假值，传 1 就全部正常。
>
> 顺带确认 `Text.SetFontProperties(frames, 结构体)` 虽然不报错，但**对可见文本无效**
> （frames 传 0/1/2 都一样，`Story.Size/Font` 不变），仍留在不可用清单里。

## 七之五十六、第五十八波已接入（2026-10-26，调色板/专色工厂 + 色彩上下文/打样）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **调色板造色** | `cdr_color(action="object_palette_factory", color2=<调色板 GUID>, values=[序号])` | ✅ `CreatePaletteColor(GUID, 5)` → `PANTONE,PAN_SPOT,5,100,<guid>`（Type 1） |
| **专色造色** | `cdr_color(action="object_from_spot", color2=<调色板 GUID>, values=[专色ID, Tint])` | ✅ `CreateSpotColor(GUID, 5, 100)` → `PANTONE,PAN_SPOT,5`（需**含专色**的调色板） |
| **色彩上下文工厂** | `cdr_doc(action="color_context_create", rendering_intent=2)` | ✅ `CreateColorContext2(配置名串, 2, 1)` → intent 2 / model 1 / 3 个配置文件 |
| **打样色彩设置** | `cdr_doc(action="proof_color_settings", profile_name=…, rendering_intent=1)` | ✅ `CreateProofColorSettings("Japan Color 2001 Coated", 1, False)` → `IVGProofColorSettings`（含 ColorContext） |

> **第五十八波纠正**：`CreatePaletteColor` / `CreateSpotColor` 之前被判「是取色不是新建」，
> 实际是**参数形态不对** —— 第一个参数要**调色板 GUID**，传对就能造出真实颜色对象。
> 另一个共性坑：`CreateColorContext2` / `CreateProofColorSettings` 的 **RenderingIntent 有效范围是 1~4**，
> 传 0 会报「参数 RenderingIntent 超出范围」—— 这也是第 18 波判 `CreateProofColorSettings` 不可用的原因。

## 七之五十七、第五十九波已接入（2026-10-26，点集 → 拟合曲线 / 追加点集）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **点集拟合曲线** | `cdr_path(action="fit_to_points", other=取点来源, params={"spacing":5,"tolerance":2})` | ✅ `SubPath.GetEvenlySpacedPoints(5, True)` → `IVGPointRange`（10 点）→ `Document.CreateCurveFitToPoints(range, False, 2.0)` → Type 3 形状 |
| **直接用点列表拟合** | 同上 `params={"points":[[x,y],…]}` | ✅ 内部用临时线取 `IVGPointRange` 再 `RemoveAll + AddPointXY` 逐点重建，5 点 → 曲线 |
| **带尖点拟合** | 同上 `params={"with_cusps":true}` | ✅ `CreateCurveFitToPointsAndCusps(range, (), False, 容差)` |
| **点集追加到曲线** | `cdr_path(action="append_points", other=取点来源, params={"mode":"plain"/"fit"/"cusps"})` | ✅ `AppendSubpathFromPoints` / `AppendSubpathFitToPoints` / `AppendSubpathFitToPointsAndCusps` 三种都生效（节点 2→4） |

> **第五十九波纠正**：`AppendSubpathFromPoints` / `CreateSubPathFromArray` 一直被记成
> 「数组入参不可用（TypeError）」，其实参数要的是 **`IVGPointRange` 对象**，不是 Python list。
> 关键发现是 `IVGPointRange` **有现成的来源**：`SubPath.GetEvenlySpacedPoints(间距, 缩放)`
> 直接返回它，再配合 `AddPointXY` / `RemoveAll` 就能随意造点集 ——
> 于是 `CreateCurveFitToPoints`、`AppendSubpathFitToPoints(AndCusps)` 一并打通。

## 七之五十八、第六十波已接入（2026-10-26，剖面线填充）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **剖面线填充** | `cdr_modify(changes={"fill":{"type":"hatch","index":1,"back":"#FFFFFF"}})` | ✅ `Fill.ApplyHatchFill(库名, 序号, 背景色, …)` → `Fill.Type = 10`（索引 0/1/2/5/10/50 都可） |
| **自定义剖面线** | `cdr_modify(changes={"fill":{"type":"hatch_custom","angle":45,"spacing":3,"width":0.6}})` | ✅ `ApplyCustomHatchFill` 18 参（Style 可传 None）→ `Fill.Type = 10` |

> **第六十波纠正**：`ApplyHatchFill` 第 2 个参数 `HatchNameOrIndex` **必须传整数序号**，
> 传字符串会报「发生意外」—— 这就是第 34 波判"剖面线不可用"的原因
> （当时看 `Document.HatchLibraries` 为空就放弃，其实填充不依赖那个集合）。
>
> **另一个实测坑（本轮排查花了两次全量冒烟）**：`Shape.GetMatrix` 之后再调
> `Shape.IsOnShape`，长跑里会返回 0（单独跑却正常）。把命中判定挪到矩阵读取**之前**即恢复 ——
> 也就是之前记的「SetMatrix 会使 IsOnShape 失效」要扩展为 **GetMatrix 也会**（至少在某些状态下）。

## 七之五十九、第六十一波已接入（2026-10-26，中心定尺寸 / 中心拉伸 / 范围排除）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **按中心定尺寸** | `cdr_shape_ops(op="set_size_ex", target=…, params={"cx":10,"cy":5,"width":40,"height":30})` | ✅ `Shape.SetSizeEx(中心X, 中心Y, 宽, 高)`：20×10 → bbox `(-10,-10,40,30)`（中心不变） |
| **按中心拉伸** | `cdr_shape_ops(op="stretch_ex", params={"cx":10,"cy":5,"stretch_x":2,"stretch_y":0.5})` | ✅ 40×30 → 80×15 |
| **范围上同样可用** | 两个 op 加 `params.use_range=true` | ✅ 对当前选区整体调用 `SetSizeEx` / `StretchEx` |
| **范围排除成员** | `cdr_shape_ops(op="all_excluding", params={"exclude":[1]})` | ✅ `ShapeRange.AllExcluding(list)` **接受 Python list**、索引 **1 基**：2 个对象排除第 1 个 → 剩 1 个 |

> **第六十一波纠正**：`SetSizeEx` / `StretchEx` 在第 33 波被判「参数语义不明、会把对象放大到离谱尺寸」，
> 现在看清了 —— **前两个参数是中心点 (CenterX, CenterY)**。当时把中心当成左上角坐标传，
> 自然算出离谱的尺寸。`AllExcluding` 之前传 `None` 报 TypeError，其实它要**索引数组**（1 基，Python list 就行）。

## 七之六十、收敛审计：剩余未覆盖方法的最终归类（2026-10-26）

> 主干 20 接口还有 **80 个**方法未覆盖。下面按「为什么还没接」把它们一次性归类，
> 每条都附实测结论 —— 这份表就是后续想再推进时的清单。

| 归类 | 方法 | 实测结论 |
|---|---|---|
| **VBA / 插件体系**（安全不接） | `Application.Evaluate`、`InitializeVBA`、`ImportWorkspace`、`Quit`、`AddPluginCommand`、`RemovePluginCommand`、`RegisterToolShape`、`RegisterToolState`、`StartTemporaryToolState`、`UnregisterToolState`、`AdviseEvents`、`UnadviseEvents`、`Shape.Evaluate`、`Shape.CustomCommand`、`Page.CustomCommand`、`Document.CustomCommand` | 前 4 个是求值/初始化/导入工作区/退出进程；`Add/RemovePluginCommand` **300 s 卡死**；`Register*ToolState` 报类型不匹配；`CustomCommand` 报「未装载此命令」/「发生意外」；`AdviseEvents` 要事件接收器 |
| **会挂起 / 弹窗**（严禁） | `Document.GetUserArea`、`Document.GetUserClick`、`Document.InteractiveImport`、`Document.InteractiveImportWithContentIdentifier`、`Document.PrintOut`、`Document.PrintColorProof`、`Text.ImportFromFile`、`Text.MakeHTMLCompatible`、`Fill/Outline/Transparency/Color.UserAssign(Ex)` | `GetUserArea`/`GetUserClick` **20 s 挂起**；打印/导入/用户取色都会弹模态框 |
| **需要载荷（数组/对象）造不出来** | `Document.CreateCurveFromArray`、`Curve.CreateSubPathFromArray`、`Curve.PutCurveInfo`、`SubPath.PutCurveInfo`、`Shape.ApplyTransformMatrix`、`Curve.ApplyTransformMatrix`、`Document.CreateFill`、`Document.CreateOutline` | 数组类报 `MemoryError: CreatingSafeArray`；`CreateCurveFromArray` **无返回**；矩阵对象只能由 `Application.MathUtils` 造，而它不存在；`CreateFill/CreateOutline` 传任何字符串都得到 `Type=0` 空对象 |
| **需要外部资源** | `Bitmap.Trace`、`Layer.CreateSymbol`、`Application.SymbolLibraries`、`Layer.CreateCustomShape`、`Layer.CreateToolShape`、`Layer.CreateOLEObject`、`Layer.PasteSpecial`、`Shape.CreateCustomEffect`、`Shape.CreateCustomDistortion`、`Style.Assign`、`Style.SetBasedOn` | 符号库/自定义工具形状/自定义效果都需要本机没装的插件或库；`Style.SetBasedOn` 因 `CreateStyleFromShape` 连续建样式后 `Styles.Find` 找不到而无法稳定验证 |
| **语义不可核实（调用不报错但无可见效果）** | `Document.LogCreateShape`、`Document.LogCreateShapeRange`、`Document.ReplaceContentByIdentifier`、`Text.SetFontProperties`、`Text.SetFontPropertiesInRange`、`Shape.RemoveFromContainer`、`Shape.ReplaceWith`、`Shape.CompareToEx`、`Shape.GetToolShapeGuid`、`Shape.ModifyToolShapeProperties`、`StyleSheet` 相关 | 例如 `SetFontProperties` 不报错但 `Story.Size/Font` 不变；`ReplaceWith` 报「需要虚拟形状」；`RemoveFromContainer` 调用无异常但页面对象数不变 |
| **类型库常量（不是方法）** | `cdrMixedDouble` / `cdrMixedLong` / `cdrMixedSingle` | 类型库枚举常量，无调用面 |

## 七之六十一、第六十二波已接入（2026-10-26，副本打开 / 字体结构体工厂）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **按副本打开** | `cdr_open(path=…, as_copy=true)` | ✅ `Application.OpenDocumentAsCopy(path, 打开选项)`：返回**未命名副本**（`FileName` 为空、`Name` 为「未命名 -N」），但内容与源文件一致（检查到 shapes=2） |
| **字体结构体工厂** | `cdr_text(paragraph_format={"report_font_struct": true})` | ✅ `Application.CreateStructFontProperties()` 可造可读（Name/Size/Style/Underline），同时在 warnings 里如实标注「`SetFontProperties` 写入对可见文本无效」 |

> **第六十二波记录**：`OpenDocumentAsCopy` 之前报 `TypeError`，是因为只传了路径 —— 它需要第二个
> **打开选项**参数（`CreateStructOpenOptions()` 即可）。另外本轮补上了「剩余 80 个未覆盖方法」的
> 收敛审计表（见上一节），把每一类的不可用原因与实测证据固定下来。

## 七之六十二、第六十三波已接入（2026-10-26，样式改父级 / 样式复制）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **样式改父级** | `cdr_style(action="set_based_on", style_name="C", based_on="P")` | ✅ `Style.SetBasedOn("P")`：参数是**名称字符串**（传样式对象报类型不匹配），返回 True，`BasedOn.Name` 变 P |
| **样式复制** | `cdr_style(action="assign", style_name="C", based_on="P")` | ✅ `Style.Assign(另一个 IVGStyle)`：属性拷过来，`ToString()` 定义串符合预期 |

> **第六十三波纠正**：这两个方法很早以前被记成"不可用"，都是**传错类型**：
> `SetBasedOn` 要的是**名称字符串**（当时传样式对象/None），`Assign` 要的是**样式对象**（当时传形状）。
> 另外发现两条相关怪癖：
> 1. `CreateStyleFromShape` 对**同一对象 + 同一类别**连续调用会互相覆盖（第二个样式建不上），
>    必须换不同对象建；
> 2. `SetBasedOn` 之后样式的 `Name` 会变，之后用原名字找不到 —— 冒烟里把 `assign`
>    排在 `set_based_on` 之前。
>
> 另：`Text.ImportFromFile(path, 1, 0)` 实测**弹出模态导入框导致 COM 卡死（240 s 超时）**，
> 已并入"严禁调用"清单。

## 七之六十三、第六十四波已接入（2026-10-26，标注对象 / 图框内容 / 辅助线 / 文本框链接）

> 这一波是**新边界**：开始覆盖主干 20 接口之外的对象接口，命中 4 个此前 0 覆盖的接口。

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **标注对象**（IVGDimension 3/3） | `cdr_shape_ops(op="dimension_info"/"dimension_set", target=标注)` | ✅ `Shape.Dimension` 拿到 `IVGDimension`；`SetTextLabelGap` / `SetOverhang` / `SetDistanceFromObject` 全部可调；Precision/Prefix/Suffix/BoxedText 等属性可读写 |
| **图框内容**（IVGPowerClip 3/3） | `cdr_shape_ops(op="powerclip_extract", target=图框)` | ✅ `Shape.PowerClip` → `IVGPowerClip`；`EnterEditMode` / `LeaveEditMode` / `ExtractShapes` 可用（内容 1 个、进入编辑态成功） |
| **辅助线几何**（IVGGuide 5/5） | `cdr_shape_ops(op="guide_info"/"guide_set", target=辅助线)` | ✅ `Shape.Guide` → `IVGGuide`；`MakeEditable` / `GetPoints` / `GetPointAndAngle` / `SetPoints` / `SetPointAndAngle` 全可用（实测角度 26.57°→30°） |
| **文本框链接**（IVGTextFrame 2/5） | `cdr_text(link_to=另一个文本, unlink=true)` | ✅ `Text.Frames.Item(1).LinkTo(t2)` → `FramesInLink=2`；`UnLink()` 可断开。`SetColumns` 要 SAFEARRAY（`[25,25]` 报类型不匹配），暂不可用 |

> **第六十四波方法**：这些接口在自动化对象上**不能 QueryInterface**（报「不支持此接口」），
> 只能通过属性取：`Shape.Dimension`、`Shape.PowerClip`、`Shape.Guide`、`Text.Frames.Item(i)`。
> 四类合计 **13 个方法**，把全局覆盖从 64.2% 推到 **65.2%** ——
> 说明"主干 20 接口之外还有长尾可挖"，这一波之后的主攻方向会转向这些对象接口。

## 七之六十四、第六十五波已接入（2026-10-26，视口缩放 / 混合效果 / 调色板管理 / 进度条）

> 继续走"主干之外"路线，这一波又吃下 **25 个方法**、4 个此前 0 覆盖的接口。

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **视口缩放**（IVGActiveView **15/15**） | `cdr_view(action="zoom_fit_page"/"zoom_fit_width"/"zoom_fit_height"/"zoom_all_objects"/"zoom_selection"/"actual_size"/"zoom_in"/"zoom_out"/"zoom_in_at_point"/"set_view_point"/"view_area"/"set_view_area"/"fit_area"/"fit_shape"/"fit_shape_range")` | ✅ `ActiveWindow.ActiveView`：`SetActualSize`→100%、`ZoomIn`→200%、`ToFitShape`→1963%、`GetViewArea` 返回 4 个值，全部真实改变视口 |
| **混合效果**（IVGEffectBlend 3 方法） | `cdr_shape_ops(op="blend_effect", params={"split":5})` | ✅ `Shape.Effects.Item(1).Blend` → `Split` / `FuseStart` / `FuseEnd` 可用（Steps/Mode 可读；`CopyFrom` 需要另一条有效混合） |
| **调色板管理器**（IVGPaletteManager 3/3） | `cdr_doc(action="palette_manager")` | ✅ `Application.PaletteManager`：`PaletteCount`=66、`GetPalette(n)` 取名、`LoadAllPalettes()`、`LoadPalette(n)` |
| **应用进度条**（IVGAppStatus 4/4） | `cdr_doc(action="app_status", settings={"begin":true,"max":3,"update":1,"message":0,"end":true})` | ✅ `BeginProgress(上限, 0)` / `UpdateProgress` / `SetProgressMessage` / `EndProgress`；`Progress` 从 -1 → 1 |

> **第六十五波方法**：`IVGActiveView` 的 15 个缩放/取景方法是本次最大一块，入口是
> **`Application.ActiveWindow.ActiveView`**；`IVGAppStatus` 的入口是 **`Application.Status`**，
> 注意 `BeginProgress` 第二个参数是 **int**（传字符串会 `ValueError`）。

## 七之六十五、第六十六波已接入（2026-10-26，线型虚线段 / 自定义工具形状属性）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **线型虚线段读取**（IVGOutlineStyle 4 名） | `cdr_doc(action="outline_style_dash", palette_index=2)` | ✅ `Application.OutlineStyles.Item(n)`：`DashCount` / `DashLength(i)` / `GapLength(i)` / `Enhanced` 都能读（第 2 号线型：1 段，dash=1.0、gap=3.0）；写入 `DashCount=`/`DashLength=(i,v)` 不报错但**不生效** |
| **自定义工具形状属性**（IVGToolShapeAttributes **12/12**） | `cdr_doc(action="tool_shape_attributes", settings={…})` | ✅ `Application.CreateToolShapeAttributes()` → `SetCanResize` / `SetCanRotate` / `SetCanSkew` / `SetCanSizeDisproportionally` / `SetCanApplyNonlinearTransforms` / `SetRegenerateOnTransform` / `SetRegenerateOnStyleChange` / `SetPropertyBarGuid` / `SetContextMenuGuid` / `SetObjectManagerBitmapGuid` / `SetEditStateGuid` / `SetDefaultShapename` 全部可调（12/12） |

> **第六十六波记录**：
> 1. `IVGToolShapeAttributes` 的 12 个 `Set*` 只是**属性配置**，因为 X8 的 `RegisterToolShape`
>    不可用（见收敛审计），所以配置完不能真正注册工具 —— 已在返回值 note 里写明。
> 2. 线型虚线段属性是**只读可用、写入无效**（与 `Layer.CreateGuide` 之外的另一类"半可用"）。
> 3. 审计脚本按名字匹配源码，动态 `getattr(attrs, name)(...)` 不会被计入 ——
>    所以把 12 个 `Set*` 改成**显式调用**，既便于审计也便于阅读。

## 七之六十六、第六十七波已接入（2026-10-26，矩形圆角 / 克隆链 / 投影偏移）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **矩形倒角**（IVGRectangle 2 方法） | `cdr_shape_ops(op="corners", params={"radius":5})` / `{"roundness":60}` | ✅ `Shape.Rectangle` → `IVGRectangle`：`SetRadius` / `SetRoundness` 可用；`MaxRadius=15`、四角半径、`CornerType`、`EqualCorners` 可读 |
| **克隆链接**（IVGCloneLink 1 方法） | `cdr_shape_ops(op="clone_links", params={"restore":true})` | ✅ `Shape.CloneLink` → `FillLinked/OutlineLinked/ShapeLinked/TransformLinked` 可读、`RestoreAllLinks()` 可调；`CloneParent` 可反查源对象 |
| **投影偏移**（IVGEffectDropShadow 1 方法） | `cdr_shape_ops(op="drop_shadow", params={"offset_x":8,"offset_y":8,"opacity":60})` | ✅ `Shape.Effects.Item(1).DropShadow` → `SetOffset(8,8)` 生效（读回 8.0），`Opacity=60` 可写 |

> **第六十七波附带修掉一个真 bug**：`cdr_modify(clone=…)` 生成的克隆体 **Name 与源对象完全相同**，
> 导致按名字根本找不到克隆体（`inspect` 里会出现两个同名对象）。现在克隆后强制改名，
> `clone_links` 这类"拿到克隆体"的操作才可用。
> 另：`page.Grid` 不存在（网格仍走 `cdr_doc(grid/set_grid)`）、`shape.Symbol` 需要符号库形状，
> `Style.Fill.SaveFill/LoadFill` 参数形态未通。

## 七之六十七、第六十八波已接入（2026-10-26，双色调子对象 / 网格频率 / 线性标注子对象）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **双色调子对象**（IVGDuotone 4 方法 + IVGDuotoneOverprint 2 方法） | `cdr_bitmap(action="duotone", duotone_type=1)` | ✅ `CreateDuotone()` → `GetCopy` / `CopyAssign` / `ResetOverprints`；转完后 `bmp.Duotone` 能取到当前双色调对象，`Overprints(1)` → `IVGDuotoneOverprint` 的 `SetValues(10,20,30,40)` / `Reset()` 都生效（OverprintCount=1、Cyan 可读写） |
| **网格频率**（IVGGrid.SetFrequency） | `cdr_doc(action="set_grid", settings={"frequency":[4,4]})` | ✅ `Document.Grid` → `SetFrequency(4,4)`：SpacingX 从 5.0 变 0.25（4 格/单位） |
| **线性标注子对象**（IVGDimensionLinear） | `cdr_shape_ops(op="dimension_set", params={"reverse_terminator_length":2})` | ✅ `Shape.Dimension.Linear` 是 **IVGDimensionLinear**（不是 bool！），`SetReverseTerminatorLength` 可用，`ReverseTerminators` 可读 |

> **第六十八波纠正一个早期误判**：`dim.Linear` 之前被 `bool()` 包了一下，看起来"返回 True"，
> 其实它是 **IVGDimensionLinear 对象** —— 包上 bool 就把接口藏起来了。
> 另外 `Document.Grid` 才是 `IVGGrid` 的入口（`Page.Grid` 不存在）。

## 七之六十八、第六十九波已接入（2026-10-26，对象树管理 / 样式填充存盘）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **对象树管理**（IVGTreeManager 2/2） | `cdr_doc(action="tree_manager", settings={"create_group_node":true,"clean":true})` | ✅ `Document.TreeManager`：`CreateGroupNode()` 返回节点、`CleanGroupNode(node)` 可清；`SelectedNodeCount` / `FirstSelectedNode` / `VirtualLayer` 可读 |
| **样式填充存盘**（IVGStyleFill 2 方法） | `cdr_style(action="fill_file", style_name=…, sheet_path="…cpf", property="save"/"load")` | ✅ `Style.Fill.SaveFill(路径, IVGFillMetadata)` / `LoadFill(路径, 元数据)` 可调（元数据由 `Application.CreateFillMetadata()` 提供）；返回 `False` 表示没存成功 |

> **第六十九波记录**：`SaveFill` 的第二个参数不是布尔而是 **`IVGFillMetadata`** 对象
> （`CreateFillMetadata()` 造），第一版传 `True`/`0` 都报 `TypeError`；
> `TreeManager` 的入口同样是 **`Document.TreeManager`**（`Application` 上没有）。
> `IVGExtrudeVanishingPoint.Share` 实测传另一个消失点也报「类型不匹配」，暂不暴露。

## 七之六十九、第七十波已接入（2026-10-26，HTML 兼容 / 按模型设色彩配置文件）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **HTML 兼容标记** | `cdr_text(make_html_compatible=true/false)` | ✅ `Text.MakeHTMLCompatible(bool)` 返回 True（艺术字/段落字都可以） |
| **按模型设色彩配置文件** | `cdr_doc(action="set_color_profile", model=1, profile_name?)` | ✅ `ColorContext.SetColorProfile(模型号, 配置文件)`：必须挑**ColorModel 与模型号一致**的配置文件，否则报「颜色模型不匹配」；5 个模型（RGB/CMYK/Gray/Lab/Hexachrome）实测全通，读回一致 |

> **第七十波记录（含一条死路结论）**：
> 1. `ColorContext.SetColorProfile` 之前一直失败，原因是**配置文件类型必须与模型号匹配**；
>    因此 `set_color_profile` 会先在 64 个已装配置文件里按 `ColorModel` 挑一个。
> 2. **UI 自动化（`ICUIAutomation` 14 / `ICUIFrameWork` 10）实测走不通**：
>    入口是 `Application.FrameWork`（`ICUIFrameWork`）→ `.Automation`，但所有方法都要 **命令栏/菜单项 GUID**，
>    而 `ICUICommandBar` 没有 `Guid` 属性、`GetActiveMenuItemGuid(i)` 在无菜单打开时返回空串 —— 
>    纯自动化场景拿不到 GUID，故整族不暴露。
> 3. `IVGColor.SpotAssign` 在 TRUMATCH 上调仍报「未找到指定的对象」（该调色板无专色）。

## 七之七十、第七十一波已接入（2026-10-26，剖面线参数 / 渐变控制点 / 椭圆 / 调色板工厂 / 窗口）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **剖面线参数**（IVGHatchFill 5 方法） | `cdr_shape_ops(op="hatch_fill", params={"fill_scale":2,"line_scale":2,"no_back_color":true})` | ✅ `Fill.Hatch`：`GetFillScale()` / `GetLineScale()` **零参**可读、`SetFillScale` / `SetLineScale` / `SetNoBackColor` 可用 |
| **渐变控制点灰度**（IVGFountainColors 2 方法） | `cdr_shape_ops(op="fountain_colors", params={"index":1,"gray_level":30})` | ✅ `Fill.Fountain.Colors`：`GrayLevel(i)` 读、`SetGrayLevel(i, v)` 写（读回 30.0），Count=2 |
| **椭圆几何**（IVGEllipse 2 方法） | `cdr_shape_ops(op="ellipse_geometry", params={"center_x":10,"center_y":45})` | ✅ `Shape.Ellipse`：`GetRadius()` 返回 `(x, y)`、`SetCenterPosition(x, y)` 可用 |
| **调色板工厂**（IVGPalettes 3 方法） | `cdr_doc(action="palette_from", settings={"name":…,"mode":"selection"/"document","open_path":…})` | ✅ `CreateFromSelection(名)` / `CreateFromDocument(名)` / `Open(路径)` 都可用；选区里没有可取色对象时自动退回整篇文档，同名冲突会加后缀重试 |
| **窗口查找**（IVGWindows.FindWindow） | `cdr_doc(action="windows", settings={"find_caption":…})` | ✅ `Application.Windows.FindWindow(标题)` 可查（找不到返回空） |
| **线型虚线段写入**（IVGOutlineStyle 2 方法） | `cdr_doc(action="outline_style_dash", palette_index=2, settings={"dash_length":0.3,"gap_length":0.2})` | ✅ `SetDashLength(i, v)` / `SetGapLength(i, v)` 可调用（读回仍为 0，属"写不进"半可用） |

> **第七十一波记录**：`IVGHatchFill` 的两个 Get 方法是**零参**（传参数会 TypeError），
> 这与带索引的 `DashLength(i)` 恰好相反 —— 同一批 TLB 里混着两种风格，只能逐个试。

## 七之七十一、第七十二波已接入（2026-10-26，屏幕手柄 / 包络效果对象 / 按名选配置文件）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **屏幕手柄**（IVGOnScreenHandle 5 方法） | `cdr_view(action="overlay_handle", x, y, color, hide?)` | ✅ `Application.CreateOnScreenHandle()` → `SetPosition` / `SetHandleColor` / `UpdateHotTracking` / `Show` / `Hide`，`IsHotTracked` 可读 |
| **包络效果对象**（IVGEffectEnvelope 4 方法） | `cdr_shape_ops(op="envelope_object", params={"preset":1,"select":1,"copy_from"/"from_curve":…}, other=…)` | ✅ `Shape.CreateEnvelope(PresetIndex 1 基, Mode, KeepLines)` → `effect.Envelope`：`Select(预设)` / `CreateFrom(形状)` / `CopyFrom(另一包络)` / `CopyFromCurve(曲线,…)` 都可用（`CopyFromShape` 参数形态未通） |
| **按名选色彩配置文件** | `cdr_doc(action="select_color_profile", profile_name="sRGB IEC61966-2.1")` | ✅ `ColorProfiles.SelectByName(名称)` 可调用（不匹配返回 False，属"调用成功但没选中"） |

> **第七十二波记录**：`Shape.CreateEnvelope` 的 `PresetIndex` 是 **1 基**（传 0 报「参数 PresetIndex 超出范围」），
> 这也是之前包络预设一直调不通的原因；`IVGOnScreenHandle` 是纯屏幕叠加对象（不写入文档），
> 与已有的 `cdr_view(overlay_*)` 属于同一层能力。

## 七之七十二、第七十三波已接入（2026-10-26，叠加层补齐 / 文本框列信息）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **叠加文字补全**（IVGOnScreenText **7/7**） | `cdr_view(action="overlay_text", text=…, x, y, distance=像素偏移)` | ✅ 除 `SetTextAndPosition` / `SetTextColor` 外，`SetText` / `SetPixelOffset` / `UpdatePosition` 也都在用 |
| **叠加曲线点集**（IVGOnScreenCurve **11/11**） | `cdr_view(action="overlay_points", x, y, x2, y2, color)` | ✅ `SetNoPen` / `SetBrush` / `SetPoints` 可用 —— `SetPoints` 要的是 **`IVGPointRange`**（复用了点集辅助函数） |
| **文本框列信息**（IVGTextFrame 4/5） | `cdr_text(frame_info=true)` | ✅ `Frames.Item(i)`：`ColumnWidth(k)` / `ColumnGutter(k)` / `ColumnCount` / `VerticalAlignment` / `Empty` / `IsFirst` / `IsLast` 可读（`SetColumns` 需 SAFEARRAY，仍不可用） |

> **第七十三波记录**：`IVGOnScreenCurve.SetPoints` 与 `Curve.AppendSubpathFromPoints` 一样，
> **要 `IVGPointRange` 而不是 Python list** —— 直接复用第五十九波写的点集构造函数即可。
> 屏幕叠加层现在是 100% 覆盖（Text 7/7、Curve 11/11），且它们都不写入文档，可安全用于预览/标注。

## 七之七十三、第七十四波已接入（2026-10-26，拉伸消失点共享 / 窗口排列）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **拉伸消失点共享**（IVGExtrudeVanishingPoint 1/1） | `cdr_shape_ops(op="extrude_vp_share", target=…, other=源拉伸对象)` | ✅ `ext.VanishingPoint.Share(源**拉伸效果**对象)` 返回 True（传另一个消失点会「类型不匹配」）；读回位置 `[70, 5]` |
| **窗口排列**（IVGWindows.Arrange） | `cdr_doc(action="windows", settings={"arrange":1})` | ✅ `Windows.Arrange(0/1/2)` 都可调用 |

> **第七十四波记录**：`Share` 的参数是**拉伸效果对象**（`effect.Extrude`）而不是消失点对象 —— 
> 又一个"名字看着像 A、实际要 B"的例子。另外冒烟里顺手修掉一个长期偶发：
> 视口缩放异常时 `IsOnShape` 会返回 0，测试前先把视口归位到实际大小即可稳定。

## 八、终版覆盖报告（2026-10-27 刷新，第八十七波后）

### 8.1 总览

| 指标 | 数值 |
|---|---|
| 类型库自动化类 / 方法 | **196 / 1619** |
| 全局方法覆盖 | **1208 / 1619 = 74.6%** |
| 主干 20 接口覆盖 | **670 / 743 = 90.2%** |
| 已达 100% 的接口 | `IVGWindow`、`IVGFountainFill`、`IVGStyle`、`IVGColorContext`、`IVGDimension`、`IVGPowerClip`、`IVGGuide`、`IVGText`(结构体族)、`IVGToolStateAttributes`、`IVGToolShapeAttributes`、`IVGAppStatus`、`IVGPaletteManager`、`IVGTreeManager`、`IVGOnScreenText`、`IVGOnScreenCurve`、`IVGOutlineStyle`(读)、`IVGAppleEvents`… 共 **30+ 个接口** |
| MCP 工具数 | **31**（含 cdr_help 能力目录） |
| 端到端联调 | **596/596 通过**（真实 CorelDRAW 18.0.0.450） |
| 单元测试 | **98/98 通过**（含 5 项目录防漂移） |
| 已记录 X8 陷阱 | **170 条** |
| 波次 | 87 轮（从 21.2% 起步） |

### 8.2 能力边界（一句话版）

* **文档/页面/图层**：新建·打开·模板·副本打开、页面增删与版式、图层增删改与可见性、命名视图、网格、参考线；
* **绘图**：矩形/椭圆/多边形/星形/螺旋/网格/直线/贝塞尔/曲线段/B 样条/位图/艺术字/段落字/连接器/线性与角度标注；
* **形状与整形**：布尔（焊接·修剪·相交·合并）、组合与解组（含 `UngroupAll(Ex)`）、打散、圆角、斜切/拉伸/定尺寸、矩阵、扭曲、包络、立体化、透镜、投影、混合、透明度、PowerClip（含内容取出）、克隆链、复制/克隆为范围；
* **路径与节点**：任意节点/线段集合、拓扑合并/平均、按偏移取点与定位线段、子路径追加/抽出、点集拟合曲线、曲线映射到骨架线、辅助线几何；
* **文字**：字体字号粗斜体下划线（`cdrTriState` 映射）、对齐/缩进/行距/字距、区间结构体读写（`*InRange` 1 基）、制表位、沿路径、转曲、HTML 兼容、文本框链接与列信息；
* **颜色与色彩管理**：`ToString`/`StringAssign`、分量读（`CorelScriptGetComponent`）与写（`*Assign`）、调色板取色/工厂、专色、`ConvertToPalette`、色彩上下文工厂/复制/合并/指派、打样设置、配置文件按模型挑选；
* **填充/描边**：均匀/渐变/图样/纹理/PostScript/剖面线（含自定义与缩放参数）、渐变控制点、描边 14/15 参、描边样式虚线段、批量化（`ApplyFill`/`ApplyOutline`）；
* **位图**：模式转换、二值化、重采样、扩边、裁剪、像素读、双色调（含 Overprint）、调色板（含选项）、内容替换、特效命令串、链接解析；
* **导入导出**：导入（含 `ImportEx`）、导出（`ExportEx` / `Document.Export` / `ExportBitmap`）、PDF、SVG；
* **视图/窗口/叠加层**：15 种缩放取景、屏幕叠加文字/直线/矩形/圆/点集/手柄（不写入文档）、新窗口、窗口排列与查找；
* **系统级**：事件队列（连接点 sink）、应用偏好、进度条、调色板管理器、对象树、`CorelScriptTools`。

### 8.3 剩余 457 个方法的归类（为什么不再接）

| 归类 | 代表 | 证据 |
|---|---|---|
| VBA / 插件体系 | `Evaluate`、`InitializeVBA`、`AddPluginCommand`、`RegisterToolState`、`CustomCommand`、`IVGApplicationEvents`(22)、`IVGToolState`(17)、`IVGAppPlugin`(4) | 300 s 卡死；或「未装载此命令」；或需宿主回调 |
| 会挂起 / 弹窗 | `GetUserArea`、`GetUserClick`、`PrintOut`、`PrintColorProof`、`Text.ImportFromFile`、`*UserAssign` | 20~240 s 挂起实测 |
| 载荷造不出来 | `CreateCurveFromArray`、`CreateSubPathFromArray`、`PutCurveInfo`、`ApplyTransformMatrix`、`IVGMathUtils`(19)、`Document.CreateFill/Outline` | `MemoryError: CreatingSafeArray`；MathUtils CLSID 未注册；字符串工厂返回空对象 |
| 需外部资源 | `Bitmap.Trace`、`CreateSymbol`、`CreateCustomEffect/Distortion`、`Layer.CreateCustomShape/ToolShape`、`IVGToolShape` | 「未装载此效果」/ 无符号库 / 无插件 |
| 交互式宿主依赖 | `ICUIAutomation`(14)、`ICUIFrameWork`(10)、`ICUIImageList`、`ICUIAutomation` 全族 | 需要命令栏/菜单项 **GUID**，纯自动化拿不到 |
| 语义不可核实 | `LogCreateShape`、`ReplaceContentByIdentifier`、`SetFontProperties`、`RemoveFromContainer`、`Style.Install`、`ColorProfiles.Install` | 调用不报错但无可见效果 / 有副作用风险 |

### 8.4 复现与维护

```bash
python scripts/coverage_audit.py          # 覆盖统计（按接口 + 全局）
python scripts/live_smoke.py              # 539 项端到端（需真实 X8；跑前先杀残留 CorelDRW）
python -m pytest -q                       # 93 项纯 Python 单元测试
```

新增能力时的固定流程（本轮 74 波一直沿用，已验证有效）：

1. 从类型库取签名（`GetFuncDesc` + `GetNames`）→ **先看参数类型，不要凭名字猜**；
2. 用**独立短超时**的探针脚本逐条试（避免一个挂起拖死整批）；
3. 成功的能力封装进 `ops.py` + `server.py`，失败的在 COVERAGE 记下**确切报错**；
4. 冒烟加断言 → 全量跑 → 更新 COVERAGE/README → 提交（`feat: 第 N 波 —— …，覆盖 X%`）。

> **给后来者的三条经验**：① 类型库里大量 `*InRange`、`Set*` 的坑是"1 基索引"和"名称串 vs 对象"；
> ② 数组入参先找 `IVGPointRange`/`IVGSnapPoint` 这类容器，别直接塞 Python list；
> ③ 同一接口里 Get 类方法的参数个数可能不一致（`GetFillScale()` 零参 vs `DashLength(i)` 带索引），逐个试。

## 七之七十四、第七十五波已接入（2026-10-27，批量斜切 / 渐变色标枚举 / 线型表保存）

> 这一轮按终版报告的"批量版/参数完整度"方向推进（应用户要求，不装 PANTONE/符号库插件）。

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **批量斜切**（ShapeRange.SkewEx 确认） | `cdr_shape_ops(op="range_skew", params={"angle_x":10,"cx":20,"cy":5})` | ✅ 选区整批斜切（两个 10×10 矩形 → bbox 宽 50.58）；范围版与形状版统一，`use_range=false` 可退回单对象 |
| **渐变色标枚举**（IVGFountainColors 补强） | `cdr_shape_ops(op="fountain_colors")` 返回 `stops=[{color, position}]` | ✅ `colors.Item(i)` 可读 `Color.ToString()` 与 `Position`；同时确认 **`AddGrayLevel` 仅灰度渐变可用**（彩色渐变报「非选择性的参数」，已改为 warnings 提示而非报错） |
| **线型表保存**（IVGOutlineStyles.Save） | `cdr_doc(action="outline_styles_save")` | ✅ `OutlineStyles.Save()` 把 131 个线型存回用户配置 |

> **第七十五波记录**：`FountainColors.AddGrayLevel(灰度, 位置)` 的位置参数在彩色渐变上无效
> （直接报错），并且 `colors.Item(2)` 在只有 2 个色标时会「超出范围」——
> 色标数实际由渐变类型决定，枚举时以 `Count` 为准。

## 七之七十五、第七十六波已接入（2026-10-27，扭曲中心重设）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **扭曲中心重设**（IVGEffectDistortion.CenterDistortion） | `cdr_shape_ops(op="distort", params={"kind":"zipper"/"pushpull"/"twister", …})` | ✅ 创建扭曲效果后调 `effect.Distortion.CenterDistortion()`（把扭曲中心重设到对象中心），Zipper / PushPull 都验证通过 |

> **第七十六波记录**：早期探测 `CreateZipperDistortion` 报 TypeError 是因为**签名记错**
> —— 正确形态是 `(OriginX, OriginY, Amplitude, Frequency, Random, Smooth, Local)`（7 参），
> ops 里早已按对签名封装；本轮只是把 `CenterDistortion()` 补进 distort 流程。
> 另：`IVGEffectDistortion.Zipper/PushPull` 属性读回会 TypeError（pywin32 类型收窄问题），
> `ColorProfiles.SelectByName` 对已知精确名也返回 False（没选中就是没选中，如实返回）。

## 七之七十六、第七十七波已接入（2026-10-27，调色板方法 / 视图 / 数据字段 / B样条控制点）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **调色板按名查找**（IVGPalette.FindColor） | `cdr_doc(action="palette_find_color", palette_index=1, color="名字")` | ✅ 可调用；**对 TRUMATCH 常见名字一律返回 0**（按名查找基本不可用），warnings 如实提示 |
| **调色板插色**（InsertColor） | `cdr_doc(action="palette_insert_color", palette_index=1, color="#00C800", color_index=1)` | ✅ `InsertColor(序号, IVGColor)`（序号在前！）；固定调色板报「无法修改固定调色板」时**自动退化到新建自定义调色板**并插入 |
| **设默认调色板**（MakeDefault） | `cdr_doc(action="palette_make_default", palette_index=1)` | ✅ `MakeDefault()` 把 PANTONE 设为默认 |
| **命名视图**（Views.AddActiveView） | `cdr_doc(action="views_add_active", view_name=…)` | ✅ 当前视口存为命名视图（Count 0→1） |
| **数据字段重排**（DataField.Reorder） | `cdr_doc(action="datafield_reorder", color_index=1, value=2)` | ✅ `Reorder(新序号)`（实测字段 `Name` 1→2） |
| **B 样条控制点扩容**（BSplineControlPoints.Resize） | `cdr_doc(action="bspline_resize", color_index=4, value=6)` | ✅ 控制点 4→6 |

> **第七十七波记录**：
> 1. `Palette.FindColor` / `InsertColor` 的参数形态完全不同：前者要**名字字符串**（且基本查不到），
>    后者是**(序号, IVGColor)**、序号在前。
> 2. 固定调色板（`OpenFixed` 打开的 TRUMATCH/PANTONE）一律**不可修改**；
>    `InsertColor`/`SetColor` 失败时对 `InsertColor` 自动退化为新建自定义调色板。
> 3. `Views.AddActiveView` / `DataField.Reorder` / `ControlPoints.Resize` 都是直通调用，一次通过。

## 七之七十七、第七十八波已接入（2026-10-27，立体化灯光/斜角）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **开灯**（LightPresent/SetLight） | `cdr_shape_ops(op="extrude_set", effect={"light_index":1,"light_intensity":80})` | ✅ `SetLight(索引, cdrLightPosition, 强度)` 打开 1~3 号灯：`LightPresent(1)=True`、`LightIntensity(1)=80` 读回一致 |
| **斜角**（SetBevel） | 同上 `effect={"bevel_depth":2.5,"bevel_angle":20}` | ✅ `SetBevel(深度, 角度, 只显示斜角)` 一次设好，读回 2.5×20 |
| **灯光状态读数** | `extrude_set` 返回 `extrude.lights=[{index, present, intensity}×3]` | ✅ 3 盏灯的开关与强度全可读（未开的灯 intensity=-1） |

> **第七十八波记录**：立体化的灯光是**带索引的属性**（`LightPresent(i)` / `LightIntensity(i)` /
> `LightPosition(i)`，i=1~3），不是单一属性；`SetLight(索引, cdrLightPosition, 强度)`
> 的第二个参数直接传 `ext.LightPosition(索引)` 取回的对象即可。
> 之前 `ext.SetLight(1, 50, 60, 100)` 的 4 参形态是错的（那是我把坐标当参数传）。

## 七之七十八、第七十九波已接入（2026-10-27，文本框多栏 SetColumns 攻克）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **文本框多栏**（IVGTextFrame.SetColumns，**最后一项**） | `cdr_text(columns={"count":2,"widths":[30,30],"gutters":[20]})` | ✅ `SetColumns(栏数, 等栏, WidthsAndGutters)`：读回 ColumnCount=2、宽 [30,20]、间距 30 |

> **第七十九波攻克**：`SetColumns` 卡了几十波，这次靠类型库反查 + 三个条件同时满足才通：
> 1. `WidthsAndGutters` 要 **VARIANT 包 SAFEARRAY(VT_VARIANT)**（`win32com.client.VARIANT` 包装）；
> 2. 数组内容是 **[宽×N, 间距×(N-1)] 交错**（2 栏 = 3 元素）；
> 3. **宽度和 + 间距和必须等于框宽**（不满足报「文本栏的宽度必须在 6.35 与 25400 之间」）。
> 至此 `IVGTextFrame` **5/5 全覆盖**。
> 另确认 `EffectEnvelope.CopyFromShape` 五参形态无论 CopyMode 取什么都报错 —— 留在不可用清单。

## 七之七十九、第八十波已接入（2026-10-27，OLE 对象只读清点）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **OLE 对象清点**（IVGOLE 只读族 6 属性） | `cdr_doc(action="ole_info")` | ✅ 遍历页面找 Type=12 的形状，读 `ProgID` / `ShortName` / `IsOpen` / `IsInPlaceActive` / `IsServerRunning` / `Modified`（无 OLE 对象时返回空数组，可用作检测） |

> **第八十波记录（含一条死路）**：
> 1. `shape.OLE` 在 **txt 源的 OLE 对象**上会报「存储空间不足」（pywin32 收窄问题）；
>    且 `Edit` / `Open` / `DoVerb` 会弹窗口 —— 所以只封装只读清点，不暴露交互方法。
> 2. `IVGPalette.Colors()` **要带括号调用**（返回 IVGPaletteColors 集合），
>    `palette.Color(ci)` 直接取单色，两者都封装了。
> 3. `Color.ModifyColorStyleColor` 调用不报错但也**不改变颜色**（与 `SetFontProperties` 同类的
>    "空操作"），维持不暴露。

## 七之八十、第八十一波已接入（2026-10-27，立体化效果复制 / 数据字段详情）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **立体化效果复制**（IVGEffectExtrude.CopyFrom） | `cdr_shape_ops(op="extrude_set", effect={"copy_from": {"name": 源对象}})` | ✅ `ext.CopyFrom(源拉伸效果)` 可调（应用后 `applied` 含 `copy_from`；实测 Depth 读回仍是目标自身值 —— CopyFrom 拷的是完整效果定义，数值读回有偏差，如实返回） |
| **数据字段详情**（IVGDataField 3 属性） | `cdr_doc(action="datafield_info")` | ✅ `Fields.Item(序号或名字)` 都可取；`Name` / `Format` 可读（`Format` 写入报「无法保存该值」，维持只读） |

> **第八十一波记录**：`CopyFrom` 复制后 `Depth` 读回 0.0/20.0 与源不一致 —— 说明 X8 的
> `CopyFrom` 在属性读回上有延迟（要重新进入效果才有值），能力本身可用但结果解读要谨慎。

## 七之八十一、第八十二波已接入（2026-10-27，点集容器编辑）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **点序反转**（Reverse） | `cdr_shape_ops(op="point_range_edit", params={"reverse":true})` 及 `fit_to_points(params={"reverse":true})` | ✅ 首点从 (0,0) 变 (40,0) |
| **追加点集**（AddPoints） | `point_range_edit(params={"append_from":…})` | ✅ 另一条曲线的点集整体并入（5→8 点，含去重） |
| **按位插入**（InsertPoint） | 同上 `params={"insert_at":2}` | ✅ 在指定位置插入另一个 IVGPoint |
| **去重**（RemoveAdjacentDuplicates） | 同上 `params={"remove_duplicates":true}` | ✅ 可调用 |

> **第八十二波记录**：`AddPoint` / `InsertPoint` 的参数是**另一个 IVGPoint 对象**（不是坐标），
> 与 `AddPoints` 要另一个 **IVGPointRange** 一致 —— 点集容器族的传参全是对象而非数值。
> `DataFields.AddEx` / `TextTabPositions`（入口不在 Text 上）/ `HatchPattern`（不在 HatchFill 上）
> 三族入口找不到或格式串不可用，留不可用清单。

## 七之八十二、第八十三波已接入（2026-10-27，立体化双灯 / 剪贴板 / 组件检测）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **立体化双灯同开** | `cdr_shape_ops(op="extrude_set", effect={"lights":[{...},{...}]})` | ✅ 1 号灯强度 80、2 号灯强度 60 同时打开（`Invoke putter` 绕过 pywin32 的 setattr 限制） |
| **剪贴板状态**（IVGClipboard） | `cdr_doc(action="clipboard_info")` | ✅ `DataPresent` / `Valid` 可读（复制对象后 present=True） |
| **组件检测**（IVGComponents） | `cdr_doc(action="components_info", settings={"names":[…]})` | ✅ `IsComponentInstalled(名)` 可查（本机 CorelDRAW/PHOTO-PAINT 都返回 False） |

> **第八十三波记录**：立体化灯光的 putter（`LightPresent(i)=v` / `LightIntensity(i)=v`）
> pywin32 的 `setattr` 走不通，要用 `ext._oleobj_.Invoke(did, 0, DISPATCH_PROPERTYPUT, False, i, v)`
> 直接按 dispid 调 —— 这也是第 78 波 `SetLight` 只能开一盏灯的根因。
> `Clipboard` / `Components` 都在 `Application` 上，一次就通。

## 七之八十三、第八十四波已接入（2026-10-27，主窗口控制 / 屏幕矩形 / 状态栏 / GMS 路径）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **主窗口控制**（ICUIFrameWindow 4 方法） | `cdr_view(action="frame_maximize"/"frame_restore"/"frame_tile"/"frame_combine")` | ✅ `Maximize` / `Restore` / `TileViews(0)` / `CombineViews` 都可用，Caption 可读（`CorelDRAW X8 (64-Bit)`） |
| **屏幕矩形**（ICUIApplication.CreateScreenRect） | `cdr_doc(action="screen_rect", settings={"left":…,"top":…,"right":…,"bottom":…})` | ✅ 4 个整数参数（0~500 均可） |
| **状态栏文字**（ICUIStatusText.SetCaptionText） | `cdr_doc(action="status_text", settings={"caption":"…"})` | ✅ `CreateStatusText()` + `SetCaptionText("…")` |
| **GMS 路径查询**（GMSManager 属性） | `cdr_doc(action="gms_info")` | ✅ `GMSPath` / `UserGMSPath` 可读 |

> **第七十七波记录的 UI 自动化结论需更新**：`ICUIFrameWork.Application` 是 ICUIApplication 入口，
> 它的 **Create\* 系列（ScreenRect/ImageList/BitmapImage/StatusText）全部可用**；
> `ICUIFrameWindow` 的窗口控制方法也可用。真正不可达的只剩**需要命令栏 GUID 的
> `ICUIAutomation` 14 个方法**和 `ICUIDataContext`/`ICUIViewHost` 等 dock 族 —— 比之前判断的窄。
> 另：`GMSManager.RunMacro` 需要**已加载的宏工程**（格式「模块.宏」），纯自动化不可用；
> `Bitmap.TraceSettings` 不存在（入口在 Trace 调用内部）。

## 七之八十四、第八十五波已接入（2026-10-27，图像列表登记）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **图像列表登记**（ICUIImageList 4/5） | `cdr_doc(action="imagelist_info", settings={"key":"MCP 图标"})` | ✅ `AddBitmap(键, ICUIBitmapImage)`（**第二参是 CreateBitmapImage() 的产物，不是 IVGBitmap**）、`ImageCount`=1、`ImageExists`=True、`ImageKeys` 可读 |

> **第八十五波记录（含一次重要反思）**：这波花了很久，卡点不是 COM 而是**我自己的补丁把分支
> 插进了别人 `return` 之后的死代码区** —— `imagelist_info`/`status_text`/`gms_info` 三个分支
> 有两次都被插到 `clipboard_info`/`assign_color_context`/`screen_rect` 的 `return` 后面，
> 永远不可达，表现为「报的错来自下一个无关分支的 `find_shape`」。
> **教训**：给这种巨型 `if act == ...` 链加分支时，锚点必须选**顶层 `    if act ==`**
> （4 格缩进）而不是 8 格缩进的嵌套行，改完立刻用 `pytest` + 一次单点调用验证可达性。

## 七之八十五、第八十六波已接入（2026-10-27，图像列表收字节 / 状态栏位图）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **图像列表收原始字节**（AddImage + CreateBitmapImage） | `cdr_doc(action="imagelist_info", settings={"key":…, "image_path":"*.bmp"})` | ✅ `AddImage(键, 原始字节, MaxSize)` 返回 True；`CreateBitmapImage(字节, MaxSize)` 拿到**真图**（Valid=True、尺寸 1×1，不再是空图） |
| **状态栏位图**（ICUIStatusText.SetBitmap → 2/2 满） | `cdr_doc(action="status_text", settings={"caption":…, "set_bitmap":True})` | ✅ `SetBitmap(CreateBitmapImage())` 可调 |
| **数据源注销**（ICUIApplication.UnregisterDataSource） | `cdr_doc(action="datasource_info", settings={"name":"mcp"})` | ✅ 可调（不存在时返回 False）；`DataContext` 存在 |
| **色彩配置文件安装**（IVGColorProfiles.Install） | `cdr_doc(action="color_profile_install", settings={"path":"*.icc"})` | ✅ 可调（非法 .icc 返回 False 不抛异常，数量 64→64 如实返回） |

> **第八十六波关键发现**：`CreateBitmapImage` / `ImageList.AddImage` 的第一个参数类型是
> **VT_VARIANT（原始图像字节）**，不是对象 —— 传 `ICUIBitmapImage`/`IVGBitmap` 都会
> 报「为参数 ImageData 指定的值无效」。传字节后 `Valid/Width/Height` 才有真实值
> （之前空构造恒为 0/0/0，是我误判成"恒为 0"）。
> `ICUIStatusText` 至此 **2/2 满**，`ICUIImageList` **5/5 满**。
> 另：`RegisterDataSource` 需要 `ICUIDataSourceFactory` 实例，纯自动化不可达；
> `Application.Printers` 在本机报 `STG_E_FILENOTFOUND`（无打印机），也留不可用。

## 七之八十六、第八十七波已接入（2026-10-27，SetLight* 正式方法 / 点距 / 图案原点 / B样条插点）

| 能力 | 落地位置 | 实测 |
|---|---|---|
| **灯光三方法**（IVGEffectExtrude 10/10 满） | `extrude_set(effect={"lights":[{present,intensity,position}…]})` | ✅ `SetLightPresent(索引, True)` / `SetLightIntensity(索引, 值)` / `SetLightPosition(索引, 码)` 是**正式方法** —— 比第 83 波的 Invoke putter 干净得多（1 号 90 + 2 号 55 同开） |
| **点间距离**（IVGPoint.DistanceTo） | `cdr_shape_ops(op="point_math", other=曲线, params={index_a,index_b})` | ✅ 返回 10.092 |
| **剖面线图案原点**（IVGHatchPattern 2 方法） | `cdr_shape_ops(op="hatch_fill", params={"pattern_origin":[2,3]})` | ✅ 入口是 `Fill.Hatch.Patterns.Item(i)`；`SetOrigin(x,y)` 写入、`GetOrigin` 读回 `[2.0, 3.0]` |
| **B 样条插控制点**（IVGBSpline 4/4 满） | `cdr_doc(action="bspline_resize", settings={"insert":{"segment":2,…}})` | ✅ `InsertControlPoints(段, 个数, x, y, 平滑)`，`segment` 要 **≥2**（1 报错），控制点 4→6 |

> **第八十七波记录**：
> 1. 第 83 波我用 `_oleobj_.Invoke(DISPATCH_PROPERTYPUT)` 硬塞灯光属性 —— 本轮发现
>    **`SetLightPresent/SetLightIntensity/SetLightPosition` 本来就是正式方法**，之前的弯路可以省掉。
> 2. `IVGBSpline.InsertControlPoints(段, 个数, x, y, 平滑)` 的**段号必须 ≥2**（段 1 报错）。
> 3. `IVGHatchPattern` 的入口是 `Fill.Hatch.Patterns.Item(i)` —— 之前试 `Hatch.Pattern` /
>    `Hatch.HatchPattern` 都不存在，**先枚举成员再猜**比猜名字快得多。
> 4. 「先 `dir()` 看成员」这条经验一次省了三轮试错，已写进维护流程。

## 七之八十七、第八十八波已接入（2026-10-28，易用性：能力目录 / 错误提示）

这一波不做新 COM 能力，专门解决「**功能多但不好发现、报错后不知道下一步**」两个易用性痛点。

| 改进 | 落地 | 效果 |
|---|---|---|
| **自描述能力目录**（新工具 `cdr_help`） | `cdr_help()` / `cdr_help(tool="cdr_shape_ops")` | 不传参返回 **31 个工具 + 418 项 action + 7 步推荐工作流**；传 tool 返回该工具全部 action，并把 **只读 action 单独列出**（如 `cdr_snap` 19 项里 10 项只读），显著降低误改风险 |
| **错误响应统一带 `hint`** | `_guard` 里按错误码挂建议 | `document_not_found` → 「调 cdr_documents 获取 document_id」；`bad_argument` → 「调 cdr_help 看 action」；`stale_revision` → 「重新 inspect 取 token」等 8 类 |
| **工具数唯一真源** | `cdr_status()["tools"]` 改为 `len(_TOOL_CATALOG)` | 之前是硬编码 `30`，加工具就过期；现在目录与状态永远一致（实测 31 = 31） |
| **action 目录不手写** | `_TOOL_CATALOG` 只存**常量名**，运行时 `getattr(ops, 常量)` 展开 | 新增 op 时目录**自动更新**，不会出现「文档里有、实际没有」 |

> **为什么这比再加 5 个 COM 方法更值**：调用方（尤其 LLM agent）最贵的是**猜参数**的成本。
> 现在它可以先 `cdr_help()` 一次看清 31 个工具与 418 项 action，再按需 `cdr_help(tool=…)`
> 拿到精确清单；报错时还能直接得到下一步动作。这套自描述是**从源码常量派生的**，
> 不会随时间腐化 —— 这是本轮相对「堆接口」的核心收益。

## 七之八十八、第八十九波已接入（2026-10-28，action 级用法示例 / 防漂移单测）

| 改进 | 落地 | 效果 |
|---|---|---|
| **action 级可复制示例** | `cdr_help(tool="cdr_shape_ops", action="extrude")` | 直接返回 `cdr_shape_ops(doc, op="extrude", target={"name": n}, params={"depth": 20})`；覆盖 **24 个工具 / 60+ 高频 action** |
| **未知 action 引导** | 同上传错 action | 返回 `available_examples` 清单 + 提示 |
| **目录防漂移单测**（新增 5 项 pytest） | `tests/test_help_catalog.py` | 断言：目录覆盖全部工具、示例工具/action 真实存在、`cdr_status` 工具数与目录一致、三种 `cdr_help` 形态结构正确 |

> **防漂移单测立刻抓到 4 个真 bug**（这就是写它的价值）：
> 1. `cdr_node` 的动作是 `add_node` 不是 `add`；
> 2. `cdr_help(tool="cdr_color")` **漏了内联基础 action**（`parse`/`convert`/`blend` 明明可用却没列出）——
>    已加 `extra_actions` 合并机制；
> 3. 我给 `cdr_color` 写的示例 action `cmyk` 根本不存在；
> 4. 给 `cdr_create` 写的 `fountain`/`text` 不是**可创建类型**（`fountain` 是填充属性）——
>    改成 `ellipse`/`curve_segment`。
>
> **经验**：文档与目录这类"人工维护的元数据"必须配**可执行的断言**，否则一定漂移。

## 七之八十九、第九十波已接入（2026-10-28，缺参前置报错 / 未知 op 近似建议）

| 改进 | 落地 | 效果 |
|---|---|---|
| **缺 `other` 前置报错** | `_precheck_shape_op()` 在进 COM 前执行 | `weld`/`trim`/`intersect`/`powerclip`/`blend`/`point_math` 等 **13 个 op** 缺第二个对象时立刻返回：`weld 需要 other 参数（被焊接的另一个对象）；写法：other={"name": "对象名"}` —— 不再白跑一次 CorelDRAW |
| **未知 op 近似建议** | `difflib.get_close_matches`（cutoff 0.55） | `op="weld_shape"` → 「你是不是想用：weld, weld_to_shape…；完整列表见 cdr_help」 |

> **一次被自己推翻的设计（重要）**：我最初还加了「params 键白名单」校验，想抓拼写错误 ——
> 结果冒烟立刻挂了两项：`corners` 的合法键 `roundness`、`drop_shadow` 的 `offset_x/offset_y`
> 都不在我手写的白名单里。**白名单不完整就会误伤正常调用**，比不校验更糟。
> 已撤销该机制，只保留**零误伤**的两条：缺 `other` 硬报错（语义明确）+ 未知 op 模糊建议。
> **结论**：校验的收益必须大于误伤风险；能"模糊提示"就不要"硬性阻断"。

## 七之九十、第九十一波已接入（2026-10-28，cdr_inspect 摘要模式 / 列表截断）

| 改进 | 落地 | 效果 |
|---|---|---|
| **摘要模式** | `cdr_inspect(doc, include="summary")` | 只回 `page_count` / `shape_count` / `layers` / `active_page` + 下一步提示 —— 大文档不再一次吐出几千个对象，**token 成本从 O(对象数) 降到 O(1)** |
| **三档查看** | `summary` < `pages` < `shapes` | `pages` 多给每页图层明细；`shapes` 给完整对象列表 |
| **对象列表截断** | `cdr_inspect(doc, include="shapes", limit=300)` | 超限时返回 `truncated=true` + `total` + 提示（`limit=0` 取全部 / `cdr_find` 精确定位），避免超大响应把客户端撑爆 |
| **非法 include 走统一错误** | 校验放在 `_guard` 内 | 返回 `bad_argument` + hint，而不是抛异常 |

> **为什么这条对易用性重要**：MCP 的每次调用结果都会进调用方上下文窗口。
> 一个 2000 对象的文档，`inspect` 一次可能就吃掉几万 token —— 现在默认摘要 + 可选截断，
> 让"看整体"和"看细节"分开，这是最实在的可控性。

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

* **能力域覆盖**：70 个能力域 / 30 个工具，日常出图闭环完整；节点编辑、对象样式、专色、
  图案/纹理填充、跨文档复制、模板/样式表/颜色样式、文档元数据、本地化元数据、对象数据字段、
  自定义虚线、文档设置、文本区间编辑、自定义箭头、图案微调、文本格式查询、位图内容替换/链接、
  文档回滚、立体化/透镜/效果管理、位图特效命令串、混合、文档颜色转换、用户偏好 KV、OLE 嵌入、
  命名视图、网格设置、叠印、段落缩进、制表位、屏幕叠加层、应用/文档事件、导出/保存/导入选项、吸附点全家族、颜色对象分量、描边对象全量属性都已补上。
* **方法级覆盖**：**全局 71.8%（1162/1619）**，主干 20 接口 90.2%（670/743）。
* **明确不可用**：位图描摹、文本导入、点集数组建曲线、尺寸标注/连接线（MathUtils 点对象
  不可创建）、符号库、文件型图案/影线填充（本机无库）、曲线贴笔触、OpenDocumentAsCopy、
  HTML 兼容文本、断字写入、填充字符串、文本尺寸查询、灰度→RGB 转换、真实应用偏好读取、
  文本分栏（SetColumns 数组参数）。
* **第七十四波新增**：拉伸消失点共享（Share 要传拉伸效果）、窗口排列；全局到 71.8%。


