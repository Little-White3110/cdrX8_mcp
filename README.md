# cdrX8_mcp

让 AI Agent 直接操作 **CorelDRAW X8 (v18)** 的 MCP 服务器。

通过 CorelDRAW 的 COM 对象模型驱动真实文档：画图、排版、上色、布尔运算、对齐、图层、导入导出、存盘、撤销。
所有坐标统一 **毫米 / 页面左下角原点 / Y 轴向上**，每次变更都收成**一个原子 undo 步**。

> **已在真机验证**：CorelDRAW Graphics Suite X8 `18.0.0.450`（64 位）+ Python 3.11.9 64 位
> - 单元测试 **55 项通过**
> - 端到端联调（真实 CorelDRAW）**158/158 通过**
> - 能力边界来自对 `vgcoreauto.tlb` 的反解（196 个自动化类 / 1619 个方法）+ 逐项实跑
> - 方法级覆盖 **38.9%**（主干 20 接口 50.2%），可用 `python scripts/coverage_audit.py` 复现
> - 完整清单见 [docs/COVERAGE.md](docs/COVERAGE.md)

---

## 快速开始

```powershell
cd "C:\path\to\cdrX8_mcp"
python -m pip install -e .

python -m pytest -q              # 单元测试（不需要 CorelDRAW）
python scripts/live_smoke.py     # 端到端联调（会启动 CorelDRAW，跑完自动关闭）
python -m cdrx8_mcp              # 启动 MCP 服务器（stdio）
```

### 客户端配置

```json
{
  "mcpServers": {
    "coreldraw-x8": {
      "command": "C:\\Python311\\python.exe",
      "args": ["-m", "cdrx8_mcp"],
      "env": {
        "CDR_ALLOWED_ROOTS": "D:\\design\\out",
        "CDR_STRICT_REVISION": "1"
      }
    }
  }
}
```

完整示例见 [`examples/`](examples/)。

---

## 工具一览（27 个）

### 会话与文档

| 工具 | 作用 |
|---|---|
| `cdr_status` | 运行状态、版本、服务配置（无副作用） |
| `cdr_connect` | 附着已运行实例；`launch=true` 才允许启动 |
| `cdr_documents` | 列出打开的文档，给出 `document_id` |
| `cdr_open` | 打开 `.cdr`，或按毫米尺寸 / 页面预设新建 |
| `cdr_inspect` | 读视图：页面、图层、`revision`、`snapshot_token`、每个对象 |
| `cdr_save` | 另存为 `.cdr`（默认不覆盖） |
| `cdr_undo` | 撤销 / 重做最近 N 步 |
| `cdr_close` | 关闭文档（默认丢弃未保存改动） |

### 绘制与修改

| 工具 | 作用 |
|---|---|
| `cdr_create` | 批量创建：矩形（含圆角）/ 椭圆（含圆弧、扇形）/ 多边形 / 星形 / 螺旋 / 网格 / 直线 / 贝塞尔 / 文字 / 段落文字 / 参考线 / **内存位图（纯色 + 像素色块）** |
| `cdr_modify` | 单个或**批量**修改：移动、缩放、旋转、翻转、斜切、矩阵/仿射变换、**清除变换**、**旋转中心**、克隆、锁定；填充（纯色/渐变/纹理/**专色**/无色/**从别的对象复制**）；描边（含**虚线**与**箭头**、**专色**）；**透明度**、**冻结/解冻透明度**；投影、轮廓图、封套；**渐变微调**；**色彩滤镜**；**层级相对调整（前置/后置/上移/下移/反转）**、复制、删除。批量用 `target={"all":true}` 或 `{"type_id":n}` |
| `cdr_text` | 排版级：内容 / 字体 / 字号 / 粗斜体 / 下划线 / 对齐 / **大小写** / **项目符号** / **首字下沉** / **行距字距** / 前后插入 / **沿路径** / 文本框自适应 / **转曲** / **艺术字↔段落字互转** / **段落内容读写与替换** / **文本区间选择·基线对齐·属性复制** / **读取全文·指定区间·导出纯文本** / **OpenType 特性** |
| `cdr_replace_text` | 全文档文字替换（底层用 X8 原生 `Page.TextReplace`） |
| `cdr_select` | 按名字 / 类型 / 矩形框选，加入选择集；**可移除或清空选择集** |
| `cdr_shape_ops` | **布尔与整形**：weld / trim / intersect / combine / group / ungroup / break_apart / to_curves / **outline_to_curves（描边转曲）** / to_symbol / powerclip / fillet / chamfer / scallop / **step_repeat（阵列）** / **distribute（原生分布）** / **boundary（边界曲线）** / **perspective（透视）** / **separate（拆效果）** / **parallel_curves（并行曲线）** / **rasterize（选择集栅格化）** / **batch_outline_to_curves / fill_mode / pixel_aligned / range_boundary / envelope_from_shape / compare** |
| `cdr_align` | 对齐与分布（左中右 / 上下中 / 等间距，对齐对象或页面） |
| `cdr_find` | 按名字或类型查找（跨图层） |
| `cdr_style` | **对象样式**：列出 / 从对象创建（fill·outline·character·paragraph）/ **基于已有样式派生** / 按名套用 / 改名 / 删除 / **读可读属性·改属性** / 列派生样式 |
| `cdr_path` | 曲线/子路径级：长度面积 / 偏移轮廓（刀模线）/ **折线采样** / **追加圆·椭圆·矩形·三点弧·贝塞尔** / **节点精简** / **自焊接** / **路径等分** / **加节点·断开·移动子路径** / **点在形状内·曲线求交·取点取切向** / **参数取点·切向量·法向量·速度·交点·按坐标找节点/线段·反转方向·焊接子路径·去重叠** |
| `cdr_node` | **节点/线段级编辑（第五波打通）**：列节点·列线段 / 移动·选中·删除 / 圆角·倒角· scallop·平滑 / 拉伸·旋转·斜切·断开 / 加节点·断线 / 读写贝塞尔控制点 / 按坐标找节点与线段 |
| `cdr_bitmap` | 位图：信息 / 转色彩模式（rgb·cmyk·gray·bw）/ **阈值二值化** / **调色板模式 0~9（9=自定义调色板）** / **双色调 0~3（可给 .cpd 预设）** / **取像素颜色** / 图像缓冲信息 / 重采样 / **扩边** / 裁剪 |
| `cdr_color` | **颜色运算（纯 Python）**：解析成 HEX/RGB/CMYK/Gray/HLS/Lab、转换、两色混合 |
| `cdr_doc` | **文档级**：概况 / 清空选择集 / 清撤销栈 / **登记调色板** / **用曲线生成自定义箭头** / **单位换算** / **取色（点/区域）** / **字体列表** / **调色板列表·取色·固定调色板打开·颜色枚举·专色创建** / **显示隐藏对象** / **页面版式** / **重复上一操作** |
| `cdr_view` | **视图**：刷新重绘 / 激活窗口 / 文档↔屏幕坐标换算 / **屏幕距离↔文档距离** |

### 图层 / 页面 / 文件

| 工具 | 作用 |
|---|---|
| `cdr_layers` | 图层列表 / 新建 / 激活 / 显隐 / 锁定 / 打印 / **上移·下移** / **读上下相邻图层** / **粘贴剪贴板** / 删除 |
| `cdr_pages` | 页面列表 / 新增 / **插入空白页** / 切页 / 删除 / 自定义尺寸 / **纸张预设** / **参考线列表** / **角度参考线** / **吸附点查询** / **解锁全部** / **对象总包围盒** |
| `cdr_import` | 导入 png / jpg / svg / pdf / ai / eps / dxf 等 |
| `cdr_export` | 导出 **png / jpg / svg / pdf**，位图可指定 DPI，可只导选区 |

### 标准调用顺序

```
cdr_status → cdr_connect → cdr_open → cdr_inspect
   → cdr_create / cdr_modify / cdr_text / cdr_shape_ops / cdr_align / cdr_layers / cdr_pages
   → cdr_export / cdr_save
```

`revision` 和 `snapshot_token` **每次变更都会变**，必须用上一步返回的新值。

---

## 能力速查

### 图元参数

| 类型 | 关键参数 |
|---|---|
| `rectangle` | `x,y,width,height` + `radius` 或 `radius_ul/ur/lr/ll` |
| `ellipse` | `cx,cy,rx,ry` + `start_angle,end_angle,pie`（给角度就是圆弧/扇形） |
| `polygon` / `star` | `cx,cy,radius,sides[,angle]`；`star` 加 `inner_radius`（`sharpness` -1~1） |
| `spiral` | `left,top,right,bottom,revolutions,spiral_type,growth` |
| `grid` | `left,top,right,bottom,cols,rows`（生成的是曲线组合体） |
| `line` | `x1,y1,x2,y2` |
| `bezier` | `x1,y1,c1x,c1y,c2x,c2y,x2,y2` |
| `polyline` | `points:[[x,y],…]` + `closed`（默认 true）—— 任意多边形/折线，逐点连线 |
| `bitmap` | `x,y,width,height` + `color` + `pixels_w,pixels_h`（默认 64×64）+ `fills:[{x,y,width,height,color}]`（像素坐标色块）+ `image_type`（4=RGB 默认，5=CMYK，2=灰度，0=黑白） |
| `text` | `x,y,text[,font,size,bold,italic,underline,align]` |
| `paragraph_text` | `x,y,x2,y2,text[,…]` |
| `guide` | `x1,y1,x2,y2`（参考线，不算图形） |

### 样式写法

```jsonc
"fill": "#CC0000"                                  // RGB
"fill": {"c":0,"m":100,"y":100,"k":0}              // CMYK（印前用这个）
"fill": {"type":"fountain","start":"#FF0000","end":"#0066FF",
         "gradient_type":1,"angle":45}             // 渐变：1线性 2辐射 3锥形 4方形
"fill": {"type":"texture","name":"...","library":"..."}
"fill": "none"                                     // 无色
"outline": 0.8                                     // 宽度
"outline": {"width":1,"color":"#000000","caps":1,"join":0,"behind_fill":true}
"outline": {"width":1,"dash":5,"start_arrow":2,"end_arrow":3}   // 虚线 + 箭头（传库索引）
"drop_shadow": {"opacity":60,"offset_x":3,"offset_y":-3,"color":"#444444"}
```

> `dash` / `start_arrow` / `end_arrow` 传的是 CorelDRAW **线条样式库 / 箭头库的索引**
> （实测必须赋 `app.OutlineStyles.Item(n)` / `app.ArrowHeads.Item(n)` 对象，直接给数字会「类型不匹配」）。
> 索引与图案的对应关系取决于本机 CorelDRAW 安装，建议先试 2~8 找到想要的样式。

### 对象类型 ID

`1`=矩形　`2`=椭圆　`3`=曲线　`4`=多边形　`6`=文字　`7`=组合体　`9`=参考线

---

## 坐标与单位约定

| 项目 | 约定 |
|---|---|
| 单位 | 毫米（内部自动把文档单位切成 mm 再切回，**不会**改动你的文档设置） |
| 原点 | 页面**左下角**，Y 轴**向上** |
| 矩形 | `{"x":10,"y":10,"width":50,"height":30}` = 左下角 + 宽高 |
| 椭圆 | `{"cx":78,"cy":25,"rx":12,"ry":12}` = 圆心 + 半径 |
| 颜色 | `#RRGGBB` / `{c,m,y,k}` / `{r,g,b}` |

---

## 安全模型

| 机制 | 说明 |
|---|---|
| **乐观锁** | 修改类工具接受 `expected_revision` + `snapshot_token`；令牌含页数与对象总数，**用户手动增删对象也会被检测到** |
| **原子 undo** | 一次工具调用 = 一个 undo 步（实测 3→5→3），`cdr_undo` 整批回滚 |
| **路径白名单** | `CDR_ALLOWED_ROOTS` 设定后只能写这些目录 |
| **覆盖保护** | 目标文件存在时默认报错，需显式 `overwrite=true` |
| **不自动启动** | 默认只附着；启动需 `launch=true`（`CDR_ALLOW_LAUNCH=0` 可全局禁止） |
| **不碰源文件** | `cdr_save` 只"另存为"，源 `.cdr` 全程只读 |

---

## 环境变量

| 变量 | 默认 | 说明 |
|---|---|---|
| `CDR_ALLOWED_ROOTS` | 空 | 允许写入的目录，`;` 分隔 |
| `CDR_STRICT_REVISION` | `0` | `1` = 强制要求 revision/token |
| `CDR_ALLOW_LAUNCH` | `1` | `0` = 永不启动 CorelDRAW |
| `CDR_TIMEOUT` | `60` | 单次 COM 调用超时（秒） |
| `CDR_TRANSPORT` | `stdio` | `http` 起本地 HTTP（`CDR_HOST`/`CDR_PORT`，默认 `127.0.0.1:8765`） |

---

## X8 特有问题（45 条，全部实测）

这些是 **X8 与 2024+ 的差异**，也是本项目存在的主要理由。

| # | 现象 | 正确写法 | 踩错的后果 |
|---|---|---|---|
| 1 | 单位枚举反直觉 | `doc.Unit = 3` 才是毫米（`2`=英尺、`4`=厘米） | 尺寸放大 10 / 304.8 倍 |
| 2 | `Document` 上没有页面属性 | `doc.ActivePage.SizeWidth / SetSize()` | AttributeError |
| 3 | `Page` 上没有 `Document` 属性 | undo 分组用外部传入的 `doc` | AttributeError |
| 4 | 没有 `CreateArtisticText2` | `layer.CreateArtisticText(...)` | AttributeError |
| 5 | `SaveAs` 必须两参 | `doc.SaveAs(path, None)` | TypeError |
| 6 | 字体不在 `shape.Text` 上 | `shape.Text.Story.Font/.Size/.Text` | 属性不可设置 |
| 7 | **`page.Shapes` 只反映活动图层** | 按 `page.Layers` 逐层遍历 | 新建图层后"对象明明在却找不到" |
| 8 | **`ShapeRange.Ungroup()` 是空操作** | `shape.Ungroup()`（`Shape` 上的方法） | 组合体永远解不开 |
| 9 | **类型 7 是组合体**（不是网格） | `SHAPE_TYPES[7] = "group"` | 判断对象类型全错 |
| 10 | **空页面 + DPI 选项导出会失败** | 失败时退回默认分辨率 | com_error「发生意外」 |
| 11 | `doc.Application` 派生的 ExportOptions 不可靠 | 优先用桥接持有的 Application | 导出失败 |
| 12 | `Layer.Import` 必须三参 | `layer.Import(path, 0, app.CreateStructImportOptions())` | TypeError |
| 13 | CDR 没有可用的 `ExportEx` 过滤器 | 存盘只能 `SaveAs` | "指定的过滤器 ID 无效" |
| 14 | 没有 `SetActivePage` | `doc.Pages.Item(i).Activate()` | AttributeError |
| 15 | **`Outline.Style` / 箭头不能直接给数字** | 赋 `app.OutlineStyles.Item(n)` / `app.ArrowHeads.Item(n)` 对象 | 类型不匹配 |
| 16 | **`Curve.Contour` 返回的是未绑定曲线** | 还要 `Layer.CreateCurve(该曲线)` 才会变成页面对象 | 偏移后页面上什么都没多 |
| 17 | **`Page.GetObjectsBoundingBox` 返回 `IVGRect`** | `IVGRect.GetRect()` 的 4 个出参 pywin32 取不回，自己遍历求和 | 拿不到包围盒 |
| 18 | **`Bitmap.Trace` 会把 CorelDRAW 卡死** | 不要调用（弹模态框、进程无响应、COM 无法中断） | 调用方一起挂起 |
| 19 | **`curve.Nodes` / `curve.Segments` 是属性，不是方法**（第五波修正） | `curve.Nodes.Item(i)` / `curve.Nodes.Range(a,b)` / `curve.Segments.Item(i)`；节点编辑通过 `cdr_node` 暴露 | 只扫方法名会误判「节点级编辑不可用」 |
| 20 | `Document` 上没有 `Copy/Paste/SnapTo*`；`Page` 没有 `Background` | 这些能力不存在 | AttributeError |
| 21 | **`SaveAsCopy` 会把当前文档名改掉** | 调用后必须用返回的 `name` 更新会话 | 之后按旧名查文档全部失败 |
| 22 | **`IVGColor` 不暴露任何分量属性** | `Red`/`Cyan`/`Gray` 全 AttributeError；`ConvertToCMYK` 只改 Type | 色彩运算改到纯 Python（`colors.py`） |
| 23 | **`SubPath.GetPointAt` 返回 `IVGPoint`** | 坐标取 `.x` / `.y`（小写） | 取 `PositionX`/元组会失败 |
| 24 | **文本 `GetContents/SetContents/Replace` 只支持段落文本** | 艺术字上调用会报「该操作需要段落文本对象」 | 先 `to_paragraph` |
| 25 | **按名取回的 Document 有时取不到 `ActiveWindow`** | 退到 `app.ActiveWindow`，再退到 `app.Windows.Item(1)`（已内置） | 视图类操作报「没有活动窗口」 |
| 26 | **`Curve.GetPolyline` 返回的是 `IVGCurve`**，不是数组 | 遍历返回曲线的 `SubPaths → Nodes` 取坐标（本项目已这么实现 `cdr_path(action="polyline")`） | 直接 `tuple()` 会失败或得到空 |
| 27 | **偏移量有两套语义** | `OffsetType=0` = 沿曲线绝对距离；`OffsetType=1` = 0~1 参数比例（取点/曲率/速度都用 1） | 同一个 `t` 取到完全不同的点 |
| 28 | **`GetTangentVectorAt(Normalize=True)` 返回零向量** | 取 `Normalize=False` 拿真实 x/y，再自己归一化 | 向量恒为 (0,0) |
| 29 | **`Text.ImportFromFile` 会卡住 COM 通道** | 不要调用；需要读文本请用 `GetContents` / `Text.Range` | 模态框阻塞，调用超时后进程仍在跑 |
| 30 | **点集参数（`AppendSubpathFromPoints` 等）pywin32 无法构造** | `TypeError: Python instance can not be converted to a COM object`；用逐点 `AppendLineSegment` 代替 | 传 list / VARIANT 都不行 |
| 31 | **`Shape.CreateBlend` 的第二个形状参数传不进去** | 同样 `TypeError`（需要 pywin32 无法构造的引用） | 混合效果暂不支持 |
| 32 | **样式 Category 是字符串，不是数字常量**（第六波修正） | `CreateStyleFromShape(shape, "fill"/"outline"/"character"/"paragraph", name)`；套用用 `Shape.ApplyStyle("样式名")` | 传 0~8 会报「未找到类别」；传 Style 对象会报「未找到样式」 |
| 33 | **`Document.CreateImage` 的 ImageType 是位图模式枚举** | 4=RGB、5=CMYK、2=灰度、0=黑白、1=Grayscale16、6=Duotone、7=LAB、8=RGB48、9=MultiChannel（3 调色板需 Palette 对象） | 传 1/3 会报「为参数 ImageType 指定的值无效」 |
| 34 | **`Layer.CreateBitmap` 的 ImageAlpha 不能传 0** | 必须传 `None` 或另一个 `IVGImage` 对象 | `TypeError: Python instance can not be converted to a COM object` |
| 35 | **`Transparency.Freeze()` 会把对象换成无名组合体** | 冻结后原引用失效、Name 变空；需要重新从页面找对象并补名（`cdr_modify` 已自动补名） | 按旧名再也找不到对象 |
| 36 | **`FontList.Item(i)` 是逐项 COM 调用，很慢** | 本机 2068 个字体；按需扫描并限制条数（`cdr_doc(action="fonts")` 默认只扫 300 条） | 全量遍历会超过 60 s 默认超时 |
| 37 | **`ShapeRange.CreateParallelCurves` 只接受曲线** | 矩形/椭圆等先 `cdr_shape_ops(op="to_curves")` | 报「对象不属于支持该操作的正确类型」 |
| 38 | **`Document.InsertPages` 第三个参数是页码数字** | `InsertPages(count, before, pageIndex)`；传 Page 对象会 TypeError | 报 `int() argument must be ... not 'IVGPage'` |
| 39 | **`CreatePaletteColor` 需要调色板 GUID 字符串** | 先读 `app.Palettes.Item(i).Identifier`（字符串 GUID），再 `CreatePaletteColor(identifier, index)` | 传 PaletteID(5) 或 0 都报「找不到调色板」 |
| 40 | **固定调色板编号随安装内容变化** | 本机 `Palettes.OpenFixed(3)` = PANTONE MATCHING SYSTEM - Corel 8（946 色）；TRUMATCH/FOCOLTONE/TOYO/标准色等也可打开 | 硬编码编号换机器可能指到别的调色板 |
| 41 | **专色名必须与调色板里的名字完全一致** | `CreateSpotColorByName(guid, "PANTONE Process Yellow CV", 100)` 可用；`"PANTONE 186 C"` 在 Corel 8 库里不存在 | 报「未找到指定的对象」 |
| 42 | **样式属性 `fill` / `outline` 是分组名，不可直接读写** | 只能读写叶子属性：`type` / `primaryColor` / `secondaryColor` / `angle` / `overprint` / `screenSpec`…（颜色属性用 `SetProperty(名字, IVGColor)`） | `GetProperty("fill")` 报「为参数 Name 指定的值无效」 |
| 43 | **`Shape.CreateBlend` 的参数 pywin32 传不进去** | 传 Shape / ShapeRange / 曲线 / 关键字参数均 `TypeError` | 混合效果暂不支持 |
| 44 | **吸附点 `FindClosestSnapPoint` 的 TypeSet 有坑** | 2/4/8 及组合（3/6/15/63）可用；1/16/32 报「服务器出现意外情况」 | 传错位掩码会直接失败 |
| 45 | **`Palettes.OpenFixed` 会往 `app.Palettes` 里插入调色板** | 打开后索引会变；用返回的 `palette_index` 或 GUID 定位 | 之前记的索引可能失效 |

改代码前请先读 [`src/cdrx8_mcp/constants.py`](src/cdrx8_mcp/constants.py) 里的注释。

---

## 已知限制（诚实清单）

**做不到：**
* **位图自动转矢量（PowerTRACE）**：`Bitmap.Trace` 实测会让 CorelDRAW 进程变成
  **无响应**（弹模态框、`Responding=False`、必须强杀进程），而 COM 调用无法中断——
  暴露它等于给 Agent 埋一个必挂的雷，所以**故意不实现**
* **文本导入**：`Text.ImportFromFile` 会把 COM 通道卡住（模态框），因此只提供
  `GetContents` / `Text.Range` 读取与 `ExportToFile` 导出，不提供反向导入
* **点集直接建曲线**：`AppendSubpathFromPoints` / `AppendSubpathFitToPoints` /
  `CreateSubPathFromArray` / `Document.CreateCurveFitToPoints` 的数组参数
  pywin32 构造不出来（`TypeError: Python instance can not be converted to a COM object`），
  逐点 `AppendLineSegment` 是目前唯一稳定路径
* **符号库插入**：`app.SymbolLibraries` 存在但本机 `Count=0`（未装符号库），
  所以只能把形状转成符号（`to_symbol`），不能插入库内符号
* **图案填充 / 影线填充**：本机没有图案文件、`Document.HatchLibraries.Count = 0`。
  专色已可用：先用 `cdr_doc(action="palette_open_fixed", palette_id=3)` 打开 PANTONE 匹配系统，
  再按调色板里的精确色名创建（`cdr_doc(action="spot_color")` / `fill={"type":"spot",...}`）
* **立体化 / 透镜 / 尺寸标注 / 连接线 / 符号插入 / 混合 / 曲线贴笔触**：
  `CreateExtrude` `CreateLens` `CreateLinearDimension` `CreateConnector` `CreateSymbol`
  `Shape.CreateBlend` 全部卡在 `TypeError: Python instance can not be converted to a COM object`
  （需要 pywin32 无法构造的结构体/枚举参数）；`Curve.CreateCurveMappedToStroke` 报「类型不匹配」

> 每一条的实测证据都记在 [docs/COVERAGE.md](docs/COVERAGE.md) 的「明确不可用」表里。
* **图案填充需要图案文件**：`ApplyPatternFill` 无文件时报错
* **尺寸标注 / 表格 / 字形 / 刀刻 / 橡皮 / 网格填充 / 智能绘图**：对象模型层面就没有接口，只能用鼠标

**有限制：**
* **导出只到当前页**：多页请逐页切换后导出
* **DPI 是近似值**：`dpi` 参数与最终像素的比例在 0.8~1.0 之间浮动（与页面/文档设置有关），**以返回的 `pixels` 为准**；同一文档内 150→300 dpi 的像素宽严格 2 倍
* **对象寻址靠 index / name / 选区**：X8 没有稳定 GUID，服务创建时会自动起 `cdr_xxxxxxxx` 名字，建议按 `name` 定位（按 index 会因增删漂移）
* **COM 不可中断**：卡在模态弹窗时调用会超时但进程仍在跑，**写操作超时后不要自动重试**
* **单实例串行**：所有调用走同一条 STA 线程，天然不支持并发（安全特性）

---

## 测试

```powershell
python -m pytest -q                # 55 项纯逻辑测试，任何平台可跑
python scripts/live_smoke.py       # 158 项端到端，需要 CorelDRAW X8
python scripts/live_smoke.py --keep   # 保留文档方便肉眼看效果
python scripts/coverage_audit.py   # 覆盖率审计（复现 README/COVERAGE 里的数字）
```

联调产物默认写到 `tmp/live_smoke/`（已在 `.gitignore` 内，可随时整个删掉）。

联调覆盖：页面预设 → 8 类图元 → 文字排版 → 投影/描边 → 虚线+箭头 → 文字沿路径 →
透明/翻转/斜切/灰度 → 段落文字大小写/行距/项目符号 → 路径偏移 → 锁定 →
**渐变微调 → 描边转曲 → 曲线拼接/等分/节点精简/几何判定 → 文字内容与类型互转 →
页面按点查询 → 颜色运算 → 文档调色板/箭头/清撤销栈 → 位图滤镜** →
weld/fillet → 框选/组合/解组 → 对齐 → 图层 → 查找 → 分页/预设/参考线/包围盒 →
四格式导出（含 DPI 倍增校验）→ 导入 → 位图转模式/重采样 → 文字替换 → 存盘/另存副本 →
撤销/重做 → **节点级编辑/线段几何/贝塞尔控制点 → 折线采样/参数取点/切法向量/按坐标查找 →
层级相对调整/阵列复制/原生分布/透视/拆效果 → 文本读取与导出 → 单位换算/取色/角度参考线** →
**对象样式创建·套用·改名·删除 → 内存位图像素取色/调色板/双色调 → 页面插入 →
并行曲线/选择集栅格化 → 图层相邻/下移 → 冻结/解冻透明度 → 字体/调色板列表 → OpenType 特性** →
**PANTONE 固定调色板/专色填充描边 → 样式属性读写/派生 → 范围描边转曲/填充规则/像素对齐/边界 →
封套/对象比较 → 选择集移除/清空 → 吸附点查询** →
过期 token 拒绝 → 关闭。

完整能力矩阵见 [docs/COVERAGE.md](docs/COVERAGE.md)。

---

## 致谢

设计思路参考（本仓库为独立实现，未复制其代码）：

* [CubicDev1/corel-mcp](https://github.com/CubicDev1/corel-mcp) — 工具面收敛、原子 undo、快照令牌
* [hmsusbusas-png/mcp-corel](https://github.com/hmsusbusas-png/mcp-corel) — 单 COM 线程 + 超时
* [pefedotov/coreldraw_mcp](https://github.com/pefedotov/coreldraw_mcp) — 能力矩阵（哪些能做、哪些只能鼠标）
* [logicbyroshan/foldermate](https://github.com/logicbyroshan/foldermate) — 进程外隔离与看门狗

## 许可

MIT。本项目只通过 COM 驱动本机已授权的 CorelDRAW，不包含、不分发、不修改任何 Corel 软件。
CorelDRAW 与 Corel 是 Corel Corporation 的商标。
