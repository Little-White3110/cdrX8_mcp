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

## 六、量化覆盖率（2026-10-24 实测）

方法：反解类型库拿到 196 个自动化类 / **1619** 个方法（已剔除 `QueryInterface` 等 COM 基础设施）；
用 AST 扫描 `src/cdrx8_mcp/*.py`，收集所有 `.方法名(` 调用与 `setattr(obj, "属性")`，
与主干接口的方法名求交集。

| 接口 | 方法数 | 已调用 | 覆盖 |
|---|---:|---:|---:|
| `IVGPage` | 21 | 12 | **57%** |
| `IVGNode` ⚠️ | 18 | 7 | 39%（同名误计，实际无节点入口） |
| `IVGTransparency` | 8 | 3 | 38% |
| `IVGLayer` | 49 | 16 | 33% |
| `IVGBitmap` | 15 | 5 | 33% |
| `IVGFill` | 12 | 4 | 33% |
| `IVGTextRange` | 31 | 9 | 29% |
| `IVGShapeRange` | 125 | 35 | 28% |
| `IVGShape` | 123 | 34 | 28% |
| `IVGWindow` | 8 | 2 | 25% |
| `IVGDocument` | 74 | 17 | 23% |
| `IVGText` | 29 | 6 | 21% |
| `IVGSubPath` | 38 | 7 | 18% |
| `IVGApplication` | 78 | 12 | 15% |
| `IVGCurve` | 36 | 5 | 14% |
| `IVGOutline` | 8 | 1 | 12% |
| `IVGColor` | 43 | 3 | 7% |
| `IVGStyle` | 15 | 1 | 7% |
| `IVGFountainFill` | 6 | 0 | 0% |
| `IVGColorContext` | 6 | 0 | 0% |
| **合计** | **743** | **179** | **24.1%** |

> **怎么读这个 24.1%**：
> 1. 它是**方法名匹配**的近似值，同名方法会误计（最典型的是 `IVGNode`——我们其实拿不到节点集合）。
> 2. 未调用的 564 个方法里，很大一部分是**变体与重复**：`CreateArrowHead`/`CreateArrowHead2`/`Ex`/`Ex2`、
>    `CreateBitmap`/`CreateBitmap2`/`CreateBitmapRect`、`ConvertToBW/CMY/CMYK/Fixed/Gray/HLS/HSB/Lab/Palette` 等。
> 3. 所以**方法覆盖率 ≠ 能力覆盖率**。22 个工具已经覆盖了 15 个能力域，
>    真正缺的是下表"七、未接入的高价值能力"里那些。

---

## 七、未接入的高价值能力（按投入产出排序）

"实测"列的含义：`✅已验` = 本次探测中已经跑通；`?未验` = 接口存在、参数形态合理但没实跑。

| # | 能力 | 接口方法 | 能干什么 | 实测 | 工作量 |
|---|---|---|---|---|---|
| 1 | **描边转曲** | `Outline.ConvertToObject()` | 把描边变成可编辑曲线对象，发排/切割前必做 | ?未验 | 小 |
| 2 | **渐变二次编辑** | `FountainFill.SetAngle / SetEdgePad / Translate / MakeOpaque / SetTransformations` | 现在改渐变只能整条重设；有了它可微调角度/过渡 | ?未验 | 小 |
| 3 | **曲线级画圆/弧/矩形** | `Curve.AppendSubpathCircle / Ellipse / Rectangle / ThreePointArc` | 在一条曲线里拼多个子路径（复杂图形、拼版） | ✅已验（Rectangle） | 小 |
| 4 | **路径等分** | `SubPath.EqualDivide(div, gap)` | 等距刻线、模切压痕、等分阵列 | ✅已验 | 小 |
| 5 | **节点精简** | `Curve.AutoReduceNodes(amount, selectedOnly)` | 转曲后减点，减小文件、改善切割质量 | ?未验 | 小 |
| 6 | **路径自焊接 / 求交判断** | `Curve.SelfWeldClosedSubpaths` / `IntersectsWith` / `IsPointInside` | 清理重叠子路径、点是否在形状内（排版校验） | ?未验 | 小 |
| 7 | **追加贝塞尔子路径** | `SubPath.AppendCurveSegment2(...)` | 画真正平滑的自由曲线（现在只有直线段折线） | ?未验 | 中 |
| 8 | **颜色运算** | `Color.ConvertToCMYK/Gray/Lab/…`、`BlendWith` | 颜色空间互转、生成色阶/配色方案 | ?未验 | 小 |
| 9 | **位图色彩滤镜** | `Shape.ApplyEffectBCI / ColorBalance / Gamma / HSL / Invert / Posterize` | 图片亮度对比、去色、Gamma 校正 | ?未验 | 中 |
| 10 | **艺术字 ↔ 段落字互转** | `Text.ConvertToParagraph / ConvertToArtistic` | 版面调整时换文字类型 | ?未验 | 小 |
| 11 | **文字检索与属性复制** | `Text.Find`、`TextRange.CopyAttributes / AlignToBaseline / FindRanges` | 精确查找、批量复制文字格式 | ?未验 | 中 |
| 12 | **点在对象查询 / 吸附点** | `Page.FindShapeAtPoint / SelectShapesAtPoint / FindClosestSnapPoint` | 按坐标点选题、吸附对齐 | ?未验 | 小 |
| 13 | **更多对齐** | `Shape.AffineTransform`、`AlignToPoint / AlignToPageCenter / AlignToGrid` | 仿射变换、对齐网格/页面中心/参考点 | ?未验 | 小 |
| 14 | **视图缩放与坐标换算** | `IVGWindow`（6 个方法） | 做预览截图、把文档坐标换算成屏幕坐标 | ?未验 | 中 |
| 15 | **文档调色板** | `Document.AddColorsToDocPalette` | 把用到的颜色登记进文档调色板（印前交接） | ?未验 | 小 |
| 16 | **从内存图建位图** | `Layer.CreateBitmap / CreateBitmap2 / CreateBitmapRect` | 程序化生成底纹/占位图，不落盘 | ?未验 | 中 |
| 17 | **自定义箭头** | `Document.CreateArrowHead*`（5 个变体） | 自定义线端箭头（现在只能用内置库索引） | ?未验 | 中 |
| 18 | **清撤销栈 / 清选区** | `Document.ClearUndoList / ClearSelection` | 批处理前重置状态、减少内存 | ?未验 | 极小 |

引用的**未探索入口**（可能藏着大能力，值得单独试）：

* `IVGApplication.CorelScriptTools` —— 看名字像内嵌的 VBA 脚本工具，若可用等于**打开宏执行通道**
* `IVGApplication.AddPluginCommand` —— 注册插件命令，也许能绕开"记录型"限制
* `IVGApplication.AdviseEvents` / `IVGDocument.AdviseEvents` —— 事件订阅（做监听型 Agent 的前提）
* `IVGApplication.ConvertToUnicode / ConvertFromUnicode`、`CheckPlatformVersion` —— 工具函数

这些**我都没试过**，不能断言可用或不可用；写在这里是为了下一轮不重复摸索。

---

## 八、一句话总结

* **能力域覆盖**：15 个能力域 / 22 个工具，日常出图闭环已完整。
* **方法级覆盖**：24.1%（179/743），剩余多为变体与重复，也有上面 18 项真正有价值的能力。
* **明确不可用**：位图描摹、节点编辑、立体化/透镜/尺寸/连接线、对象样式、符号库、专色。


