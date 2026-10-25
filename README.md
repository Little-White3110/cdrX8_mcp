# cdrX8_mcp

让 AI Agent 直接操作 **CorelDRAW X8 (v18)** 的 MCP 服务器。

通过 CorelDRAW 的 COM 对象模型驱动真实文档：画图、排版、上色、布尔运算、对齐、图层、导入导出、存盘、撤销。
所有坐标统一 **毫米 / 页面左下角原点 / Y 轴向上**，每次变更都收成**一个原子 undo 步**。

> **已在真机验证**：CorelDRAW Graphics Suite X8 `18.0.0.450`（64 位）+ Python 3.11.9 64 位
> - 单元测试 **68 项通过**
> - 端到端联调（真实 CorelDRAW）**323/323 通过**
> - 能力边界来自对 `vgcoreauto.tlb` 的反解（196 个自动化类 / 1619 个方法）+ 逐项实跑
> - 方法级覆盖 **49.4%**（主干 20 接口 62.3%），可用 `python scripts/coverage_audit.py` 复现
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

## 工具一览（29 个）

### 会话与文档

| 工具 | 作用 |
|---|---|
| `cdr_status` | 运行状态、版本、服务配置（无副作用） |
| `cdr_connect` | 附着已运行实例；`launch=true` 才允许启动 |
| `cdr_documents` | 列出打开的文档，给出 `document_id` |
| `cdr_open` | 打开 `.cdr`，按毫米尺寸 / 页面预设新建，或**从模板创建**（`template=路径`） |
| `cdr_inspect` | 读视图：页面、图层、`revision`、`snapshot_token`、每个对象 |
| `cdr_save` | 另存为 `.cdr`（默认不覆盖；**可传 save_options**：嵌入 ICC/VBA、缩略图、覆盖、版本等） |
| `cdr_undo` | 撤销 / 重做最近 N 步 |
| `cdr_close` | 关闭文档（默认丢弃未保存改动） |

### 绘制与修改

| 工具 | 作用 |
|---|---|
| `cdr_create` | 批量创建：矩形（含圆角）/ 椭圆（含圆弧、扇形）/ 多边形 / 星形 / 螺旋 / 网格 / 直线 / 贝塞尔 / 文字 / 段落文字 / 参考线 / **内存位图（纯色 + 像素色块）** / **OLE 文件嵌入（ole_file）** |
| `cdr_modify` | 单个或**批量**修改：移动、缩放、旋转、翻转、斜切、矩阵/仿射变换、清除变换、旋转中心、克隆、锁定；填充（纯色/渐变/纹理/图案/专色/无色/复制）；**图案填充微调（旋转/平铺/镜像/前景色）**；描边（虚线、箭头、专色、自定义虚线线型）；透明度/图案透明、冻结/解冻；投影、轮廓图、封套；渐变微调；色彩滤镜；层级调整、复制/剪切到剪贴板、移动到/复制到图层、对象数据字段、**叠印（填充/描边/位图）**、复制、删除 |
| `cdr_text` | 排版级：内容 / 字体 / 字号 / 粗斜体 / 下划线 / 对齐 / 大小写 / 项目符号 / 首字下沉 / 行距字距 / 段落缩进 / **制表位** / 前后插入 / 沿路径 / 自适应 / 转曲 / 艺术字↔段落字互转 / 段落内容读写与替换 / 文本区间 / 区间删除·复制·粘贴·拉直 / 读取全文·指定区间·导出纯文本 / **字体属性读取·断字设置读取·格式段列表·按格式查询区间（@font/@bold）** / OpenType 特性 |
| `cdr_replace_text` | 全文档文字替换（底层用 X8 原生 `Page.TextReplace`） |
| `cdr_select` | 按名字 / 类型 / 矩形框选，加入选择集；**可移除或清空选择集** |
| `cdr_shape_ops` | **布尔与整形**：weld / trim / intersect / combine / group / ungroup / break_apart / to_curves / **outline_to_curves** / to_symbol / powerclip / fillet / chamfer / scallop / **step_repeat / distribute / boundary / perspective / separate / parallel_curves / rasterize / batch_outline_to_curves / fill_mode / pixel_aligned / range_boundary / envelope_from_shape / compare / project / unproject / to_document / extrude / lens / extrude_set / lens_set / effect_info / clear_effects / blend / blend_set** |
| `cdr_align` | 对齐与分布（左中右 / 上下中 / 等间距，对齐对象或页面） |
| `cdr_find` | 按名字或类型查找（跨图层） |
| `cdr_style` | **对象样式**：列出 / 从对象创建（fill·outline·character·paragraph）/ **基于已有样式派生** / 按名套用 / 改名 / 删除 / **读可读属性·改属性** / 列派生样式 / **样式表 .cdst 导入导出** / **颜色样式与颜色协调** |
| `cdr_path` | 曲线/子路径级：长度面积 / 偏移轮廓（刀模线）/ **折线采样** / **追加圆·椭圆·矩形·三点弧·贝塞尔** / **节点精简** / **自焊接** / **路径等分** / **加节点·断开·移动子路径** / **点在形状内·曲线求交·取点取切向** / **参数取点·切向量·法向量·速度·交点·按坐标找节点/线段·反转方向·焊接子路径·去重叠** |
| `cdr_node` | **节点/线段级编辑（第五波打通）**：列节点·列线段 / 移动·选中·删除 / 圆角·倒角· scallop·平滑 / 拉伸·旋转·斜切·断开 / 加节点·断线 / 读写贝塞尔控制点 / 按坐标找节点与线段 |
| `cdr_bitmap` | 位图：信息 / 转色彩模式（rgb·cmyk·gray·bw）/ **阈值二值化** / **调色板模式 0~9（9=自定义调色板）** / **双色调 0~3（可给 .cpd 预设）** / **取像素颜色** / 图像缓冲信息 / 重采样 / **扩边** / 裁剪 / **内容替换（SetImageData，支持色块/像素绘制）** / **位图特效命令串（ApplyBitmapEffect）** / **链接信息·重新加载·解析链接** |
| `cdr_color` | **颜色运算**：解析成 HEX/RGB/CMYK/Gray/HLS/Lab、转换、两色混合（纯 Python）；**X8 颜色对象**：`object_info` / `object_name` / `object_compare`（IsSame·IsSimilar·距离）/ `object_blend`（原地混合）/ `object_copy` / `object_string_assign` / `object_convert`（8 种模型原地转换）/ `object_fixed`（USER·TRU·PAN_PROC）/ `object_assign`（分量直写，含套版色）|
| `cdr_doc` | **文档级**：概况 / 清空选择集 / 清撤销栈 / **登记调色板** / **用曲线生成自定义箭头（可登记全局箭头库）** / **单位换算** / **取色（点/区域）** / **字体列表** / **调色板列表·取色·固定调色板打开·颜色枚举·专色创建·自定义调色板创建** / **图案画布创建（含区域变换）** / **纹理库列表** / **文档元数据读写·本地化元数据** / **数据字段·对象数据·色彩管理·线型样式·叠印读取** / **文档设置读写** / **应用信息·文档回滚** / **色彩配置文件·文档颜色转换** / **用户偏好 KV** / **命名视图** / **网格设置** / **应用/文档事件队列** / **显示隐藏对象** / **页面版式** / **重复上一操作** |
| `cdr_view` | **视图**：刷新重绘 / 激活窗口 / 文档↔屏幕坐标换算 / **屏幕距离↔文档距离** / **屏幕叠加层（文字/矩形/直线/圆，不写入文档）** |
| `cdr_snap` | **吸附点**：对象吸附点枚举（类型掩码过滤）/ **自动点·包围盒点·对象点·边上点定址** / **新增用户吸附点（GUID）** / **文档·应用级自由点** / 最近吸附点查询（页面级与集合级）/ 移动（相对+绝对）/ 删除 / 方向 / 选择 / **SnapPointRange 批量（Move·Delete·Add·Remove·Find·ChangeDirection·SetAutoSnap）** |
| `cdr_outline` | **描边对象**：全量属性读取（宽度/颜色/线型/线帽/连接/笔尖角度与拉伸/斜接/对齐/箭头）/ 写入（含 `SetProperties` 全参路径）/ **描边转曲线对象** / 复制描边 / 描边比较 |

### 图层 / 页面 / 文件

| 工具 | 作用 |
|---|---|
| `cdr_layers` | 图层列表 / 新建 / 激活 / 显隐 / 锁定 / 打印 / **上移·下移** / **读上下相邻图层** / **粘贴剪贴板（可跨文档）** / 删除 |
| `cdr_pages` | 页面列表 / 新增 / **插入空白页** / 切页 / 删除 / 自定义尺寸 / **纸张预设** / **参考线列表** / **角度参考线** / **吸附点查询** / **解锁全部** / **对象总包围盒** |
| `cdr_import` | 导入 png / jpg / svg / pdf / ai / eps / dxf 等，**可传 import_options（外链位图/重采样/保留图层/水印检测/ICC）** |
| `cdr_export` | 导出 **png / jpg / svg / pdf**，位图可指定 DPI，可只导选区，**可传 export_options（抗锯齿/透明/压缩/尺寸/分辨率等）** |

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

## X8 特有问题（111 条，全部实测）

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
| 46 | **`IVGPatternCanvas` 的绘制参数是布尔掩码，不是颜色** | `FillArea(x1,y1,x2,y2, State=True)`、`SetPixel(x,y,True)`、`Line(Flags,x1,y1,x2,y2,True)`、`Clear()`；颜色在应用填充时用 front/end 给 | 传 `IVGColor` 会 `TypeError: int() argument...` |
| 47 | **图案填充 Type=0 才走内存画布** | `ApplyPatternFill(0,"",canvasIndex,front,end,False)`；Type=1/2 要图像文件 | 传 1/2 报「无法装入以下图样文件」 |
| 48 | **`Document.CreateFill/CreateOutline` 传字符串不生效** | 任意字符串都返回 `IVGFill/IVGOutline`，但应用后填充类型为 0（无效） | 不要依赖字符串填充/描边 |
| 49 | **跨文档复制靠系统剪贴板** | `Shape.Copy()`/`Shape.Cut()` → 目标文档 `Layer.Paste()`；粘贴对象保留原名字 | 同名对象可能造成定位混淆，建议粘贴后改名 |
| 50 | **对象换图层是 `Shape.MoveToLayer/CopyToLayer`** | 传 `page.Layers.Item(i)` 对象；`CopyToLayer` 返回副本 | `Layer` 上没有"移动对象"的方法 |
| 51 | **`.cdr` 也能当模板用** | `Application.CreateDocumentFromTemplate("任意.cdr", True)` 可创建新文档；本机无 .cdt 也能测 | 模板路径不存在时报错 |
| 52 | **`SaveStyleSheet/LoadStyleSheet` 是恢复样式表的可靠办法** | `doc.SaveStyleSheet(path.cdst)` 后删除样式，再 `LoadStyleSheet(path)` 可恢复 | 样式表文件约 400 KB |
| 53 | **`CreateColorStyles` 的返回值恒为 0，但确实创建了颜色样式** | 用 `StyleSheet.AllColorStyles.Count` 前后对比判断；`harmonies=1..3` 会生成一组协调色 | 按返回值判断会误以为失败 |
| 54 | **`Text.GetFontProperties` 返回默认结构体** | Name=''、Size≈1e-38；真实字体用 `Text.Story.Font/Size/Bold/Italic/Underline` | 用结构体会读到空字体 |
| 55 | **断字设置与 HTML 兼容文本写不进去** | `SetHyphenationSettings` 调用不报错但值不变；`MakeHTMLCompatible` 返回 False | 只提供断字读取 |
| 56 | **`OpenDocumentAsCopy` 参数无法传递** | Options 传 0 / None 都是 `TypeError: Python instance can not be converted to a COM object` | 需要副本请用 SaveAs 新路径 |
| 57 | **文档元数据在 `Document.Metadata` 下，不是 Document 直接属性** | `Metadata.Title/Author/Subject/Keywords/Notes/Copyright` 可读写；`Revision/DocLanguage/DocID/LastAuthor` 只读 | 直接 `doc.Keywords` 会 AttributeError |
| 58 | **对象/选区可以直接变成新文档** | `Shape.CreateDocumentFrom(False)` / `ShapeRange.CreateDocumentFrom(False)` 返回 `IVGDocument` | 新文档关闭时偶发「服务器出现意外情况」，需容错 |
| 59 | **自定义调色板可创建并落盘** | `app.Palettes.Create(名称, 路径, True)` + `AddColor(颜色)` + `SaveAs(路径)`；文件约 260 B | 路径必须是绝对路径 |
| 60 | **`Metadata.Localizable*` 是 `IVGLocalizableString` 对象** | 不能直接 `str()` 当文本用；读取普通 Title/Keywords 即可 | 读 Localizable 字段会得到对象 repr |
| 61 | **混合/曲线贴笔触确定不可用** | `CreateBlend` 传 Shape/ShapeRange/`_oleobj_` 全部 TypeError；`CreateCurveMappedToStroke` 传 shape/curve/`_oleobj_` 全部「类型不匹配」 | 不要再尝试绕过 |
| 62 | **`CreateOutlineStyle` 返回虚拟线型（Index=-1）** | 必须 `app.OutlineStyles.AddStyle(virtual)` 注册后再赋给 `Outline.Style`；`OutlineStyles.Save()` 才持久化 | 直接赋虚拟样式无效 |
| 63 | **`DataFields.Add` 的 Format 必须是数字格式串** | `"0"` 可用；`"General"`/`""`/`"@"` 报「指定的格式字符串无效」 | 数据字段建不出来 |
| 64 | **对象数据字段可持久化在 .cdr 里** | `Shape.ObjectData.Add(Field, Value)`；`DataItem.Value` 可更新；`cdr_doc(object_data)` 可读 | 适合给对象打 SKU/备注标签 |
| 65 | **色彩管理可读、不可建** | `Document.ColorContext` 读 RenderingIntent/BlendingColorModel/ColorProfileNameList；`CreateColorContext()` 参数 TypeError | 只做读取，不改文档色彩配置 |
| 66 | **`Document.Keywords/Notes` 是 write-only** | 直接赋值可用，直接读 AttributeError；读请走 `Document.Metadata.Keywords/Notes` | 读直接属性会报错 |
| 67 | **文档设置可读写** | `EditAcrossLayers / ApplyToDuplicate / CurvePrecision / Resolution / PreserveSelection / ShapeEnumDirection / TextFormatter / WorldScale / FacingPages / FirstPageOnRightSide / Layout / DrawingOriginX·Y` | `SaveSettings/RestoreSettings` 实测未恢复 EditAcrossLayers，不可靠 |
| 68 | **文本区间可删除/复制/粘贴/拉直** | `TextRange.Delete/Copy/Paste/Straighten`；copy/paste 走系统剪贴板 | 删除 `[6,11]` 后文本变短、粘贴会插到 0 位置 |
| 69 | **自定义箭头可登记全局箭头库** | `CreateArrowHead(curve)` → `Application.ArrowHeads.Add(head)` → 箭头库 89→90；可赋给 `Outline.StartArrow` | 箭头名为空，用索引定位 |
| 70 | **`TextRange.FindRanges` 用单引号字符串** | `"@font = 'Arial'"` 命中整段、`"@bold = true"` 命中 0；双引号会解析失败 | 尺寸比较（`@size > 10`）报单位不匹配，暂不可用 |
| 71 | **`TextRange.ApplyStyle` 调用不报错但未观测到效果** | 字符样式套到区间后 Bold/Size 未变化 | 整体样式请用 `Shape.ApplyStyle` |
| 72 | **PatternFill 15 个属性可微调** | `RotationAngle/TileWidth/TileHeight/OriginX·Y/MirrorFill/MirrorFillX·Y/TransformWithShape/FrontColor/BackColor/SkewAngle…` 可写且回读正确 | 只能在已有图案填充上微调 |
| 73 | **PatternCanvas 支持区域变换** | `CopyArea/RotateArea/FlipArea/PutCopy` 均可用；`PatternCanvases.SetItem` 可替换画布 | FillArea 的 x2/y2 是终点不是宽高 |
| 74 | **`Bitmap.SetImageData` 的 Alpha 必须传 None** | 传 0 报 `TypeError: Python instance can not be converted to a COM object`；传 None 后像素被替换 | 可用 `Document.CreateImage` + `FillArea` 造新像素再替换 |
| 75 | **`TextRange.EnumRanges(-1)` 可列格式段** | 每段返回 Text/Font/Size/Bold/Italic/Start/End | 混合格式文本可逐段读取 |
| 76 | **`Document.Revert()` 可回滚到上次保存** | 实测 15→16→15 对象；**破坏性**，丢弃未保存修改 | 调用前确认已保存 |
| 77 | **应用信息可读** | `GetSupportedOpenTypeFeatures()` 230 项；`CheckPlatformVersion(16/17)=true、18=false`；`IsUILanguageAvailable(1033/2052)=false` | 平台/语言判断语义未完全确认，按返回值展示 |
| 78 | **`CreateExtrude`/`CreateLens` 必须传全参数（含颜色对象）** | 只传前几个 int 参数报 `TypeError: Python instance can not be converted to a COM object`；补齐 Color 对象后成功 | 已封装默认颜色参数 |
| 79 | **效果通过 `Shape.Effects` 访问** | `Effects.Count/Item(i)`；`Effect.Extrude` → IVGEffectExtrude、`Effect.Lens` → IVGEffectLens；`Separate()/Clear()` 可用 | 属性可读写（Depth/Angle/Shading/Rate/Color…） |
| 80 | **`MathUtils` 无法创建点对象** | CLSID `{E98800C5-8C42-45FD-81EA-2E751150E40F}` 报「没有注册类」，Application 无 MathUtils 属性 | 尺寸标注/连接线因此仍不可用 |
| 81 | **`ApplyBitmapEffect` 用命令串控制效果** | `ApplyBitmapEffect("Whirlpool", "WhirlpoolEffect WhirpoolSpacing=20,WhirlpoolSmear=9,WhirlpoolTwist=70,WhirlpoolStreak=60,WhirpoolWarp=1,WhirlpoolRandomSeed=0")`；执行后像素由绿(0,255,0)变(0,243,0) | 命令串格式：`效果名Effect 参数=值,…` |
| 82 | **`CreateFreeConnector` 报「未实现」** | 即使只传 4 个数字也返回 E_NOTIMPL；`CreateLinearDimension/CreateConnector` 需要点对象 | 连接线/尺寸标注仍不可用 |
| 83 | **本机 `SymbolLibraries.Count=0`** | 没有符号库，`CreateSymbol` 无法取到符号名 | 需先安装符号库文件 |
| 84 | **`CreateBlend` 必须传全参数且 `Path=None`** | 只传 2 个参数报 `TypeError: Python instance can not be converted to a COM object`（Path 默认 0 不是 IDispatch）；补全 12 个参数、Path=None 后成功 | 已封装 `cdr_shape_ops(op="blend")` |
| 85 | **混合通过 `Effects.Item(i).Blend` 访问** | `IVGEffectBlend` 的 Steps/Spacing/Angle/ColorBlendType/Mode/Loop/RotateShapes/MapNodes/FuseStart/FuseEnd/ColorAcceleration/SpacingAcceleration/LinkAcceleration/AccelerateSize 可读写 | `blend_set` 可微调步数/间距等 |
| 86 | **`Application.ColorManager` 提供 64 个色彩配置文件** | `ColorProfiles`、`GetProfilesByColorModel(0=RGB,1=CMYK,2=Gray,3=Lab)`、`DefaultColorContext` 均可用 | 本机配置文件来自 Windows/Corel 安装目录 |
| 87 | **文档颜色转换可行，但灰度→RGB 会报错** | `profile.CreateColorContext()` → `Document.ConvertToColorContext()`；RGB→CMYK→RGB 成功，灰度→RGB 报「服务器出现意外情况」 | 转换前建议先 RGB→CMYK→RGB 路径 |
| 88 | **`Application.CreateColorContext()` 直接调用不可用** | 无参/全参都 TypeError；必须用 `profile.CreateColorContext()` | 已封装 `cdr_doc(action="convert_color")` |
| 89 | **用户偏好可当持久化 KV 用** | `RegisterUserApplicationPreference(Group,Key,Default)` + `Get/SetApplicationPreferenceValue`；同会话读写正常 | 真实应用偏好组（`VGDoc Pref Settings`）读取报「发生意外错误」 |
| 90 | **当前工具信息可读** | `ActiveTool` / `ActiveToolStateGuid` 可读，已并入 `cdr_doc(action="app_info")` | 工具 GUID 可用于后续自定义工具开发 |
| 91 | **OLE 文件可直接嵌入** | `Layer.CreateOLEObjectFromFile(绝对路径, False, False, caption, None)` 返回 `type_id=12` 形状（约 28×18mm 居中）；.txt 实测成功 | 不弹对话框；可用 cdr_modify 移动/删除 |
| 92 | **`Document.CreateView` 的 Page 参数必须是 IVGPage 对象** | 传 int 报 `TypeError: Python instance can not be converted to a COM object`；传 `doc.ActivePage` 才成功 | 命名视图可创建/列出/激活/删除 |
| 93 | **`Document.Grid` 网格设置可读写** | Type / SpacingX / SpacingY / Visible / Snap 全部可写；mm 上下文下 spacing 单位是毫米（默认 5×5） | 已封装 `cdr_doc(action="grid"/"set_grid")` |
| 94 | **叠印设置可读写** | `Shape.OverprintFill/OverprintOutline/OverprintBitmap` 三个布尔属性；cdr_modify 写入后回读一致 | 适合印前输出 |
| 95 | **段落缩进可读写** | `TextRange.LeftIndent/RightIndent/FirstLineIndent`；写 5/6/3mm 后回读一致 | `cdr_text` 读取结果含 `indent`；Tabs/Columns 未封装 |
| 96 | **制表位可增删查** | `TextRange.Tabs.Clear()` / `Add(Position,Alignment,Leadered)`；TabPosition 的 Position/Alignment/Leadered 可读 | 已封装 `cdr_text(clear_tabs, tab_stops)` |
| 97 | **文本分栏 `SetColumns` 的数组参数传不进去** | 第二参数 WidthsAndGutters 需要数组；list/tuple/VARIANT 都报类型不匹配/参数无效 | 分栏创建暂不可用 |
| 98 | **同一会话多次颜色转换可能让 CorelDRAW 无响应** | 连续 RGB→CMYK→RGB/gray 后进程可能卡死；单次 RGB→CMYK 稳定 | 需要多次转换时重启 CorelDRAW |
| 99 | **本地化元数据按 LCID 读写** | `Metadata.LocalizableTitle` 等返回 `IVGLocalizableString`；`GetLanguages()` 是 tuple（含 "" 与 "2052"）；`SetLangString(2052,"中文")` 有效 | 默认值用 `DefaultLangString` |
| 100 | **屏幕叠加层可即时显示** | `CreateOnScreenText/Curve/Handle` + SetTextAndPosition/SetPen/SetRectangle/SetLine/SetCircle/Show；坐标为屏幕像素，xor 绘制 | 不写入文档，窗口重绘/关闭后可能消失；用 cdr_view(overlay_*) |
| 101 | **Application 事件连接点可用** | 生成的 `DIVGApplicationEvents(application)` 能收到 DocumentNew/WindowActivate/DocumentClose/SelectionChange 等；队列用 `cdr_doc(action="events")` 读 | `ApplicationEvent` 高频且无信息量，已忽略 |
| 102 | **事件 sink 的 TLB 路径要猜** | `app.Path` 是 `...\Draw\`，TLB 实际在 `..\Programs64\vgcoreauto.tlb`；按候选路径查找 | 找不到 TLB 时静默降级，不影响主功能 |
| 103 | **文档级事件可用** | 每个文档挂 `DIVGDocumentEvents` sink，可捕获 `Doc:ShapeCreate/Delete/Move/Transform/Change`、图层/页面变化、SelectionChange；首次访问或新建时自动挂载 | 事件名带 `Doc:` 前缀，与 app 事件共用队列 |
| 104 | **导出选项可配置** | `cdr_export(export_options={anti_aliasing_type,transparent,compression,...})` 会写入 `StructExportOptions`；实测 PNG 正常生成并返回已应用字段 | 复杂 CMYK 文档 + 透明选项可能失败，建议用干净 RGB 文档 |
| 105 | **保存选项可配置** | `cdr_save(save_options={embed_icc_profile,keep_appearance,thumbnail_size,...})` 会写入 `StructSaveAsOptions`；实测 .cdr 正常保存 | Version 字段语义未单独验证 |
| 106 | **导入选项可外链位图** | `cdr_import(import_options={link_bitmap_externally:true,resample_dpi:96})` 导入后 `Bitmap.ExternallyLinked=true`、`LinkFileName` 指向源图；`UpdateLink()` 可重新加载 | 导入后自动补对象名；`ResolveLink` 会清空 LinkFileName |
| 107 | **移动对象吸附点会变成用户吸附点** | `SnapPoints.Item(i).Move(dx,dy)` 或 `Range.Move` 作用于 Type=2（对象点）后，该点变 Type=1（用户点）并排到用户点段，`Count` 不变 | 想保留对象点就别移动；用 `cdr_snap(action="move")` 并检查返回的 `kind` |
| 108 | **`FindClosestSnapPoint` 未命中不报错** | 热区内没有匹配点时返回 Type=16 的“自由吸附点”，坐标恰好等于查询点 | 本服务的 `cdr_snap(action="find")` 会识别并回 `fallback=true, found=false`；别把 Type=16 当命中 |
| 109 | **`Outline.SetProperties` 不是不能用，而是参数必须给全** | 14 个形参里 int 位传 `None` 会 `TypeError: int() argument must be … not NoneType`；把实参全给对（对象可选位传 None、数值可选位传 0/0.0）后，宽度/颜色/线帽/连接/斜接全部生效 | `cdr_outline(action="set", changes={"use_set_properties":true,…})` |
| 110 | **`Color.IsSimilar` 对同色返回 False** | 同一个 RGB 值分别建两个颜色对象，`IsSame`=True 但 `IsSimilar`=False | 判色用 `IsSame`；要数值差异用 `GetColorDistanceFrom` |
| 111 | **`Color.UserAssign/UserAssignEx` 参数是窗口句柄** | 形参名 `ParentWindowHandle`，传颜色对象报 `int() argument must be … not IVGColor` —— 本质是弹「选择颜色」对话框 | **不暴露**（会卡住无人值守会话）；用 `*Assign` 分量直写代替 |

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
* **文件型图案 / 影线填充**：`ApplyPatternFill(Type=1/2)` 需要图案图像文件，
  `Document.HatchLibraries.Count = 0`（本机无影线库）。**内存图案画布填充已可用**：
  `cdr_doc(action="pattern_canvas")` 画掩码 → `fill={"type":"pattern","canvas_index":N,...}`。
  专色也已可用：`palette_open_fixed(3)` 打开 PANTONE 后按精确色名创建。
* **尺寸标注 / 连接线 / 符号插入 / 曲线贴笔触**：
  `CreateLinearDimension` `CreateConnector` 需要 `MathUtils.CreatePoint` 创建的 IVGPoint，
  但 MathUtils CLSID 未注册、Application 无该属性；`CreateSymbol` 需要符号库（本机 Count=0）；
  `Curve.CreateCurveMappedToStroke` 报「类型不匹配」
* **混合**：**可用**（必须传全参数且 Path=None），见 `cdr_shape_ops(op="blend")`
* **立体化/透镜**：**可用**（须传全参数含颜色对象），见 `cdr_shape_ops(op="extrude"/"lens")`
* **按副本打开 / HTML 兼容文本 / 断字写入**：`OpenDocumentAsCopy` 参数 TypeError；
  `MakeHTMLCompatible` 返回 False；`SetHyphenationSettings` 写回不生效（只提供断字读取）

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
python -m pytest -q                # 68 项纯逻辑测试，任何平台可跑
python scripts/live_smoke.py       # 323 项端到端，需要 CorelDRAW X8
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
**图案画布填充/图案透明 → 剪贴板复制/剪切（同文档+跨文档）→ 对象移入/复制到图层** →
**模板创建文档 → 样式表 .cdst 导出/导入 → 颜色样式/颜色协调 → 纹理库列表与纹理填充 →
3D 平面投影/反投影 → 字体属性与断字设置读取** →
**文档元数据读写 → 对象/选区提取为新文档 → 自定义调色板创建与位图应用** →
**对象数据字段读写 → 自定义虚线线型 → 文档色彩管理读取** →
**文档设置读写 → 文本区间删除/复制/粘贴 → 自定义箭头登记** →
**图案画布区域变换 → 图案填充微调 → 文本格式查询（@font/@bold）** →
**位图内容替换 → 文本格式段读取 → 应用信息 → 文档回滚（已保存副本）** →
**立体化 → 立体化微调 → 透镜 → 效果枚举/清除** →
**位图内容替换 → 位图 Whirlpool 特效（命令串）** →
**混合（blend）→ 混合微调（blend_set）→ 清除** →
**色彩配置文件列表 → 文档转 CMYK → 转回 RGB → 按名转灰度** →
**用户偏好注册/读写 → 当前工具信息** →
**OLE 文件嵌入 → 命名视图创建/激活/删除** →
**网格设置读写（间距/显示/吸附）** →
**叠印 Fill/Outline 读写 → 段落左右/首行缩进 → 制表位** →
**本地化元数据（默认 + 中文/英文）** →
**屏幕叠加层（文字/矩形/直线/圆）→ 清除** →
**应用事件队列（DocumentNew/WindowActivate）→ 清空** →
**文档事件（Doc:ShapeCreate/ShapeChange）** →
**导出选项（抗锯齿/透明/分辨率）** →
**保存选项（嵌入 ICC/缩略图）** →
**外链位图导入（import_options）→ 链接信息/更新** →
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
