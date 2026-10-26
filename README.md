# cdrX8_mcp

让 AI Agent 直接操作 **CorelDRAW X8 (v18)** 的 MCP 服务器。

通过 CorelDRAW 的 COM 对象模型驱动真实文档：画图、排版、上色、布尔运算、对齐、图层、导入导出、存盘、撤销。
所有坐标统一 **毫米 / 页面左下角原点 / Y 轴向上**，每次变更都收成**一个原子 undo 步**。

> **已在真机验证**：CorelDRAW Graphics Suite X8 `18.0.0.450`（64 位）+ Python 3.11.9 64 位
> - 单元测试 **93 项通过**
> - 端到端联调（真实 CorelDRAW）**553/553 通过**
> - 能力边界来自对 `vgcoreauto.tlb` 的反解（196 个自动化类 / 1619 个方法）+ 逐项实跑
> - 方法级覆盖 **71.8%**（主干 20 接口 90.2%），可用 `python scripts/coverage_audit.py` 复现
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

## 工具一览（30 个）

### 会话与文档

| 工具 | 作用 |
|---|---|
| `cdr_status` | 运行状态、版本、服务配置（无副作用） |
| `cdr_connect` | 附着已运行实例；`launch=true` 才允许启动 |
| `cdr_documents` | 列出打开的文档，给出 `document_id` |
| `cdr_open` | 打开 `.cdr`，按毫米尺寸 / 页面预设新建，或**从模板创建**（`template=路径`）；**`create_options`（CreateDocumentEx：名称/页宽/页高/分辨率）**、**`open_options`（OpenDocumentEx：CodePage）** |
| `cdr_inspect` | 读视图：页面、图层、`revision`、`snapshot_token`、每个对象 |
| `cdr_save` | 另存为 `.cdr`（默认不覆盖；**可传 save_options**：嵌入 ICC/VBA、缩略图、覆盖、版本等） |
| `cdr_undo` | 撤销 / 重做最近 N 步 |
| `cdr_close` | 关闭文档（默认丢弃未保存改动） |

### 绘制与修改

| 工具 | 作用 |
|---|---|
| `cdr_create` | 批量创建：矩形（含圆角）/ 椭圆（含圆弧、扇形）/ 多边形 / 星形 / 螺旋 / 网格 / 直线 / 贝塞尔 / **曲线段（curve_segment：端点+两端控制柄）** / **边界版矩形·椭圆·完整多边形（rect_ltrb / ellipse_ltrb / polygon_full）** / **位图构造器（bitmap_at / bitmap_rect）** / **宽文本（text_wide / paragraph_wide：显式语言/字符集）** / **B 样条（bspline：控制点构造，可 to_curves）** / **连接器（connector / connector_ex / right_angle_connector）** / **线性标注（linear_dimension）·角度标注（angular_dimension）** / 文字 / 段落文字 / 参考线 / **内存位图（纯色 + 像素色块）** / **OLE 文件嵌入（ole_file）** |
| `cdr_modify` | 单个或**批量**修改：移动、缩放、旋转、翻转、斜切、矩阵/仿射变换、清除变换、旋转中心、克隆、锁定；填充（纯色/渐变/纹理/图案/专色/无色/复制）；**图案填充微调（旋转/平铺/镜像/前景色）**；描边（虚线、箭头、专色、自定义虚线线型）；透明度/图案透明、冻结/解冻；投影、轮廓图、封套；渐变微调；色彩滤镜；层级调整、复制/剪切到剪贴板、移动到/复制到图层、对象数据字段、**叠印（填充/描边/位图）**、复制、删除 |
| `cdr_text` | 排版级：内容 / 字体 / 字号 / 粗斜体 / 下划线 / 对齐 / 大小写 / 项目符号 / 首字下沉 / 行距字距 / 段落缩进 / **制表位** / 前后插入 / 沿路径 / 自适应 / 转曲 / 艺术字↔段落字互转 / 段落内容读写与替换 / 文本区间 / 区间删除·复制·粘贴·拉直 / 读取全文·指定区间·导出纯文本 / **字体属性读取·断字设置读取·格式段列表·按格式查询区间（@font/@bold）** / OpenType 特性 / **四个结构体快照（struct_info：字体·间距·对齐·断字）** / **区间宽接口（range_action：插入·替换·折叠·重设·判等·包含·行距）** / **进入编辑态（begin_edit）** / **结构体工厂写段落格式（paragraph_format：对齐·缩进·字距·行距·连字符，支持 range_start 走 *InRange 区间）** / **区间结构体读取（range_struct：字体·间距·对齐·连字符，1 基）** / **文本框链接（link_to / unlink）** |
| `cdr_replace_text` | 全文档文字替换（底层用 X8 原生 `Page.TextReplace`） |
| `cdr_select` | 按名字 / 类型 / 矩形框选，加入选择集；**可移除或清空选择集** |
| `cdr_shape_ops` | **布尔与整形**：weld / trim / intersect / combine / group / ungroup / break_apart / to_curves / **outline_to_curves** / to_symbol / powerclip / fillet / chamfer / scallop / **step_repeat / distribute / boundary / perspective / separate / parallel_curves / rasterize / batch_outline_to_curves / fill_mode / pixel_aligned / range_boundary / envelope_from_shape / compare / project / unproject / to_document / extrude / lens / extrude_set / lens_set / effect_info / clear_effects / blend / blend_set / **unlock_all / ungroup_all / break_apart_ex / clear_effect / range_info / range_add / range_remove / apply_fill（Fill 工厂批量上色）/ **matrix（矩阵读写）/ is_type / in_front_of / is_on_shape / distort（pushpull·twister·zipper）/ place_text（文字内嵌）/ duplicate_range · clone_range / **position_ex（9 个参考点）/ find_snap（按引用反查吸附点）/ linked（链接对象）/ rasterize_ex（带参位图化）/ overprint_state / **range_from_list（对象列表建范围）/ range_bitmap（整批位图化）/ range_outline（整批描边 14/15 参）/ copy_to_layer_range（复制到其它图层）/ align_to_range（对齐到形状范围）/ copy_props（按掩码复制属性） / corners（矩形圆角）/ clone_links（克隆链）/ drop_shadow（投影偏移） / dimension_info·dimension_set（标注对象）/ powerclip_extract（图框内容）/ guide_info·guide_set（辅助线）/ skew_ex（形状级斜切）/ range_skew（选区整批斜切）/ set_size_ex（按中心定尺寸）/ stretch_ex（按中心拉伸）/ all_excluding（范围排除）/ range_outline_apply（整批套描边）/ ungroup_all·ungroup_all_ex·ungroup_ex（解组家族）/ sibling（同层相邻遍历）** |
| `cdr_align` | 对齐与分布（左中右 / 上下中 / 等间距；`to=selection/page`，**以及原生范围对齐 `to=page_edge/page_center/grid/point/shape/range`**） |
| `cdr_find` | 按名字或类型查找（跨图层） |
| `cdr_style` | **对象样式**：列出 / 从对象创建（fill·outline·character·paragraph）/ **基于已有样式派生** / 按名套用 / 改名 / 删除 / **读可读属性·改属性** / 列派生样式 / **样式表 .cdst 导入导出** / **颜色样式与颜色协调** / **覆盖属性（override_props）·继承状态（inherited）·属性取值（get_property）·属性清除** / **改父级（set_based_on：传名称串）·样式复制（assign）** |
| `cdr_path` | 曲线/子路径级：长度面积 / 偏移轮廓（刀模线）/ **折线采样** / **追加圆·椭圆·矩形·三点弧·贝塞尔** / **节点精简** / **自焊接** / **路径等分** / **加节点·断开·移动子路径** / **点在形状内·曲线求交·取点取切向** / **参数取点·切向量·法向量·速度·交点·按坐标找节点/线段·反转方向·焊接子路径·去重叠** / **曲线合并（AppendCurve）·焊接（WeldWith）·命中判定·曲线副本·曲线结构信息·任意线段集合（segment_range）·追加曲线段（append_curve_segment）·抽出子路径（extract_subpath）·焊接落地（weld_to_shape）·按偏移取点（point_at_offset）·按偏移定位线段（segment_at_offset）·曲线映射到骨架线（map_to_stroke）·点集拟合曲线（fit_to_points）·追加点集（append_points：plain/fit/cusps）** |
| `cdr_node` | **节点/线段级编辑（第五波打通）**：列节点·列线段 / 移动·选中·删除 / 圆角·倒角· scallop·平滑 / 拉伸·旋转·斜切·断开 / 加节点·断线 / 读写贝塞尔控制点 / 按坐标找节点与线段 / **前后邻居（Next·Previous）·节点间距·平均位置·末端节点合并·任意不连续节点集合（arbitrary：move/rotate/skew/stretch/autoreduce/smoothen/set_type/delete/fillet/chamfer/scallop）·**节点坐标搬运（copy_point：GetPoint/SetPoint）** |
| `cdr_bitmap` | 位图：信息 / 转色彩模式（rgb·cmyk·gray·bw）/ **阈值二值化** / **调色板模式 0~9（9=自定义调色板）** / **双色调 0~3（可给 .cpd 预设）** / **取像素颜色** / 图像缓冲信息 / 重采样 / **扩边** / 裁剪 / **内容替换（SetImageData，支持色块/像素绘制）** / **位图特效命令串（ApplyBitmapEffect）** / **链接信息·重新加载·解析链接** / **调色板选项转换（paletted_options：颜色数·抖动·指定调色板）** / **剖面线填充（hatch，索引必须是整数）/ 自定义剖面线（hatch_custom 18 参）** |
| `cdr_rect` | **矩形对象（IVGRect）**：读派生值 / 平移 / 四边外扩内缩 / 并 / 交 / 点包含 / 复制；**按矩形造形**（CreateRectangleRect 带四角圆角、CreateEllipseRect 带角度/扇形） |
| `cdr_color` | **颜色运算**：解析成 HEX/RGB/CMYK/Gray/HLS/Lab、转换、两色混合（纯 Python）；**X8 颜色对象**：`object_info` / `object_name` / `object_compare`（IsSame·IsSimilar·距离）/ `object_blend`（原地混合）/ `object_copy` / `object_string_assign` / `object_convert`（8 种模型原地转换）/ `object_fixed`（USER·TRU·PAN_PROC）/ `object_assign`（分量直写，含套版色·固定调色板 fixed·脚本式 script）/ **`object_model`（HSB·CMY·BW·YIQ·Fixed·CreateColorEx 工厂）/ `object_from_string`（颜色串造色）/ `object_palette`（按调色板 GUID 取专色）/ `object_invalidate` / **`object_components`（CorelScriptGetComponent 读分量）/ `object_convert_palette`（转指定调色板）/ `object_modify_style_color`（ModifyColorStyleColor）/ `object_palette_factory`（CreatePaletteColor）/ `object_from_spot`（CreateSpotColor）** |
| `cdr_doc` | **文档级**：概况 / 清空选择集 / 清撤销栈 / **登记调色板** / **用曲线生成自定义箭头（可登记全局箭头库）/ 命名箭头（CreateArrowHead2）** / **单位换算** / **取色（点/区域）** / **字体列表** / **调色板列表·取色·固定调色板打开·颜色枚举·专色创建·自定义调色板创建** / **图案画布创建（含区域变换）** / **纹理库列表** / **文档元数据读写·本地化元数据** / **数据字段·对象数据·色彩管理·线型样式·叠印读取** / **文档设置读写** / **应用信息·文档回滚** / **色彩配置文件·文档颜色转换** / **用户偏好 KV** / **命名视图** / **网格设置** / **应用/文档事件队列** / **显示隐藏对象** / **页面版式** / **重复上一操作**  / **COM 单位换算（convert_units）** / **文档设置快照（save_settings·restore_settings·reset_settings，实验性）** / **解析全部位图链接（resolve_links）** / **指派色彩上下文（assign_color_context）** / **样式表设为默认（stylesheet_default）** / **命名箭头变体（arrowhead_ex）** / **颜色上下文深操作（color_context_info/copy/merge）** / **渐变变换矩阵（fountain_transform）** / **CorelScriptTools 辅助对象（script_tools）** / **色彩上下文工厂（color_context_create）** / **打样设置（proof_color_settings）** / **调色板管理器（palette_manager）** / **进度条（app_status）** / **线型虚线段（outline_style_dash）** / **自定义工具形状属性（tool_shape_attributes）** |
| `cdr_view` | **视图**：刷新重绘 / 激活窗口 / **新窗口（new_window）** / 文档↔屏幕坐标换算 / **屏幕距离↔文档距离** / **屏幕叠加层（文字/矩形/直线/圆，不写入文档）** |
| `cdr_snap` | **吸附点**：对象吸附点枚举（类型掩码过滤）/ **自动点·包围盒点·对象点·边上点定址** / **新增用户吸附点（GUID）** / **文档·应用级自由点** / 最近吸附点查询（页面级与集合级）/ 移动（相对+绝对）/ 删除 / 方向 / 选择 / **SnapPointRange 批量（Move·Delete·Add·Remove·Find·ChangeDirection·SetAutoSnap）** |
| `cdr_outline` | **描边对象**：全量属性读取（宽度/颜色/线型/线帽/连接/笔尖角度与拉伸/斜接/对齐/箭头）/ 写入（含 `SetProperties` 14 参与 `SetPropertiesEx` 15 参两条全参路径）/ **描边转曲线对象** / 复制描边 / 描边比较 |

### 图层 / 页面 / 文件

| 工具 | 作用 |
|---|---|
| `cdr_layers` | 图层列表 / 新建 / 激活 / 显隐 / 锁定 / 打印 / **上移·下移** / **读上下相邻图层** / **粘贴剪贴板（可跨文档）** / **图层信息（含 is_guides_layer）** / **图层内按名找对象（FindShape）** / **高级粘贴（paste_ex）** / 删除 |
| `cdr_pages` | 页面列表 / 新增（**可定尺寸 AddPagesEx**）/ **插入空白页（可定尺寸 InsertPagesEx）** / 切页 / 删除 / 自定义尺寸 / **纸张预设** / **参考线列表** / **角度参考线** / **吸附点查询** / **解锁全部** / **对象总包围盒** |
| `cdr_import` | 导入 png / jpg / svg / pdf / ai / eps / dxf 等，**可传 import_options（外链位图/重采样/保留图层/水印检测/ICC）** |
| `cdr_export` | 导出 **png / jpg / svg / pdf**，位图可指定 DPI，可只导选区，**可传 export_options（抗锯齿/透明/压缩/尺寸/分辨率等）**；**`export_api` 可选 ex（ExportEx）/ document（Document.Export）/ bitmap（Document.ExportBitmap）** |

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

## X8 特有问题（163 条，全部实测）

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
| 112 | **`ShapeRange.UngroupAll/UngroupAllEx/UngroupEx` 是空操作** | 选一个组合体、页面 1 个对象，调完这三个方法页面对象数仍然是 1；只有 `Shape.Ungroup()` 能真解开 | `cdr_shape_ops(op="ungroup_all")` 已改成对整页组合体逐个 `Shape.Ungroup()` 递归解（实测 1→3，名字保留） |
| 113 | **`ShapeRange.Sort` 的比较表达式没逆出来** | `"width"` / `"@1.width"` / `"s1.width < s2.width"` 全部报 `Identifier x is not a method of global/pair`；空串报 `参数 CompareExpression 不能为空字符串` | 不暴露排序；要 z 序用图层/层级操作 |
| 114 | **`SetSizeEx`/`StretchEx` 会瞬间把内容放大几千毫米** | `SetSizeEx(0,0,60,40)` 后对象跑到 (3540,1560) 且尺寸比例失控；`StretchEx(0,0,100,50,1)` 把 20mm 的对象变成 2000mm+；`AllExcluding([shape])` 直接 TypeError | 三个都不暴露（参数语义未确认，风险高） |
| 115 | **粗体/斜体的形参是 `cdrTriState`，true = -1** | `CreateArtisticText(...,Bold=True)` 与 `Text.Story.Bold = True` 都**不生效**（后者还会把粗体关掉）；传 `-1` 才生效。下划线又是另一套枚举：1=有、0/-1=无 | `cdr_create`/`cdr_modify`/`cdr_text` 已改走 `_tristate()` 与 `_underline()`；此 bug 存在了 30 多波才被发现 |
| 116 | **字体表会过期，建文本静默回退默认字体** | 长时间批量建文本后，`CreateArtisticText(font="Arial")` 得到的是宋体且 Bold/Italic 丢失；`Application.ForceUpdateFontTable()` 后恢复 | `cdr_doc(action="refresh_fonts")`；建文本时自动兜底刷一次 |
| 117 | **中文默认字体文档改不动拉丁字体** | 文档默认字体=宋体时，`Story.Font="Arial"`、`Story.Bold=-1`、`TextRange.Font=…` 全部「不报错也不生效」，Size/Underline 仍生效 | X8 字体替换限制；`cdr_text` 如实返回 `applied`，需要 Arial 等拉丁字体时建议新文档先设默认字体 |
| 118 | **`SaveSettings/RestoreSettings` 调用成功但看不到效果** | 把 `Resolution` 300→72 后`RestoreSettings(Tag)` 仍是 72（三方法都不报错）；`CreateProperties` 返回的 `IVGProperties` 也缺 `Name/Add/Value` 方法 | 设置快照标为实验性；可靠改设置走 `cdr_doc(action="set_settings")` |
| 119 | **`Node.Next/Previous` 是方法不是属性** | 类型库里看着是 propget，实际必须 `node.Next()`；当属性用只拿到 bound method（`AttributeError: function object has no attribute Index`）。另外 `ConnectWith`/`ExtendSubPaths` 在两条不同曲线上仍报「指定的节点必须在另一条曲线上」 | 已封装 `cdr_node(action="neighbors")`；跨曲线接续暂不可用，同曲线合并用 `join_nodes` |
| 120 | **文字结构体只能读不能写** | `Text.FontProperties(2, struct)` 当方法调报 `TypeError: IVGStruct... object is not callable`；改写成属性赋值 `text.FontProperties = struct` 不报错但把字号写成 1.17e-38 之类的垃圾值；带下标的 `FontPropertiesInRange` 返回的结构体字段同样是垃圾 | 四个结构体只做**只读快照**（`cdr_text(struct_info=true)`）；写字号/字体继续走 `font=/size=` |
| 121 | **`CreatePaletteColor` / `CreateSpotColor` 不是"新建颜色"** | `CreatePaletteColor("PANTONE 185 C", 0)` 报「找不到颜色」、`CreateSpotColor("MCP专色", 100, 0)` 报「未找到颜色」—— 它们是从**已存在**的调色板/文档专色里取色 | 新建专色仍用 `cdr_doc(action="spot_color")`（CreateSpotColorByName + 已打开调色板） |
| 122 | **`Shape.GetMatrix` 的 tx/ty 是内部单位** | 读回的初始矩阵是 `[1,0,0,1,-850000,-285000]`；写 `tx=5` 实际位移极小（内部按 1/100000 存）。另外用 `SetMatrix` 改过矩阵的形状，`IsOnShape` 判定会失效（坐标映射被重置） | 矩阵读写精确、但别拿它做毫米级平移（用 `cdr_modify(move=…)`）；做命中判定要放在改矩阵之前 |
| 123 | **`GetOverprintFillState` 恒为 1** | `Shape.GetOverprintFillState()` 在`OverprintFill` 开/关前后都返回 1（outline 版才随开关 2→4 变化）；`RemoveFromContainer(1)` 报「参考的对象已不再存在」；`Shape.Next/Previous` 是带 3 个参数的 propget，取不到 | 判断填充叠印读布尔属性 `OverprintFill/OverprintOutline`；z 序用页面枚举 |
| 124 | **`NodeRange.Move` 的第三参是"范围内起始索引"** | `Move(0,10,0,False)` 报「参数 Index 超出范围（1~N）」；传 1 才正常。另外 `SegmentRange.Remove(1)` 把 2 条删成 3 条（行为反直觉），`Layer.CreateBSpline(Curve)` 报类型不匹配 | 已封装 `arbitrary`/`segment_range` 并规避这些参数坑；`SegmentRange.Remove` 不暴露 |
| 125 | **`CreateShapeRangeFromArray` 是少数吃 Python list 的数组入参** | `doc.CreateShapeRangeFromArray([shape,…])` 直接可用（Count/Exists/Item 都正常），而点集类 `CreateCurveFitToPoints([[...]])` 仍报 TypeError；`ShapeRange.CopyPropertiesFrom` 报「执行方法 coVGShape::CopyPropertiesFrom 时发生意外」 | 批量操作用对象列表建范围；属性复制走 `cdr_modify(fill_from/outline_from)` |
| 126 | **`ExportBitmap` 不调 `Finish()` 就不生成文件** | `doc.ExportBitmap(...)` 返回 `ICorelExportFilter`，调用本身不报错、也不写盘；必须再调 `handle.Finish()`。`ExportEx` 同理 | 已在 `ops.export` 里统一 `Finish()` + 等待文件出现；`PasteSpecial` 格式串不匹配返回 False，`LogCreateShape` 直接报 com_error |
| 127 | **新窗口在 Document 上没有** | `doc.NewWindow()` 报 `AttributeError`；必须走 `Application.ActiveWindow.NewWindow()`（返回 IVGWindow，Windows.Count +1）。`FountainFill.GetTransformations()` 是 **4 个出参**（返回 tuple），不是矩阵对象 | 已封装 `cdr_view(action="new_window")` 与 `cdr_doc(action="fountain_transform")` |
| 128 | **文档级字符串填充工厂是空对象** | `Document.CreateFill("CMYK,…")` 返回 `Type=0` 的空填充、`CreateOutline("…")` 宽度 0 —— 字符串格式无效；`Page.CustomCommand` 报「未装载此命令」；`CreateStyleFromShape` 连续建两个样式后 `Styles.Find` 找不到（列表只剩一个 name=null） | 填充/描边用 `cdr_modify` 或 `CreateUniformFill`；派生样式用 `create(based_on=…)` |
| 129 | **四个 API 会把 CorelDRAW 卡死** | `AddPluginCommand`/`RemovePluginCommand` 300 s 未返回；`Document.GetUserArea`/`GetUserClick` 20 s 挂起；`Document.CreateCurveFromArray` 无返回。COM 调用无法中断，一旦被调到整个进程就废了 | **全部列入禁用清单**，工具面永不暴露 |
| 130 | **文字结构体写入端在 pywin32 下彻底不可用** | `Text.FontProperties` 等属性赋值把字号写成 1.17e-38；改用 `_oleobj_.Invoke(memid, DISPATCH_PROPERTYPUT, …)` 配 `VARIANT(VT_I4/VT_VARIANT)` 仍报「类型不匹配」（`VT_USERDEFINED` 参数无法 marshaling） | 结构体只做只读快照；写字号/字体走 `font=/size=` 与 `_tristate()` 映射 |
| 131 | **导入/导出过滤器号不通用** | PNG 导出过滤器号 802 传给 `Layer.ImportEx(path, 802, opts)` 时调用成功但 `shapes_added=0`（静默不导入）；传 0 自动判断才正常。另外 `ShapeRange.DeleteItem(1)` 会把**文档里的对象一起删掉**（范围与页面都 2→1） | `cdr_import(filter_id=…)` 只做透传并标注；想只移除范围成员用 `range_remove` |
| 132 | **`PaletteAssign`/`SpotAssignByName` 的第一参是调色板 GUID** | `PaletteAssign(调色板名, 序号)` 会失败；正确姿势是 `Palette.Identifier`（GUID 字符串）。在**不含专色**的调色板（如默认调色板）上取专色报「未找到指定的对象」 | `cdr_color(action="object_palette", color2=<GUID>, values=[序号])`；专色要先 `palette_open_fixed` 打开 PANTONE 等专色调色板 |
| 133 | **分量其实读得到：入口是 `CorelScriptGetComponent`** | 以前一直写「`IVGColor` 不暴露分量」，其实 `Color.CorelScriptGetComponent(模型号, 0,0,0,0,0,0,0)` 返回 `(5000+模型号, v1..v7)`：CMYK `(5002,10,20,30,40,…)`、RGB `(5005,11,22,33,…)`、Gray `(5009,50,…)`。`ConvertToPalette` 的参数也是**调色板 GUID**，传序号报「找不到调色板 N」 | 已封装 `cdr_color(object_components/object_convert_palette)`；至此颜色分量读写都通了 |
| 134 | **`ApplyTransformMatrix` 永远拿不到矩阵对象** | `Shape/Curve.ApplyTransformMatrix` 需要 `IVGTransformMatrix`，而工厂全在 `IVGMathUtils` 上，`Application.MathUtils` 直接 AttributeError（CLSID `{E98800C5-…}` 未注册）。另外 `Curve.WeldWith` 返回的是**游离曲线**，必须 `BindToDocument(doc)` 再 `Layer.CreateCurve` 才能变成文档里的真形状 | 矩阵变换用 `Shape.SetMatrix`/`cdr_modify(affine=…)`；焊接用 `cdr_path(weld_to_shape)`（已封装收尾） |
| 135 | **文本结构体写入只认段落文本** | `Text.SetAlignProperties(0, 结构体)` 等三个 setter 对**艺术字**报「该参数需要段落文本」；对段落文本则完全正常（用 `CreateStructAlignProperties/SpaceProperties/HyphenationSettings` 从零造结构体即可）。另外 `ConvertToUnicode/ConvertFromUnicode` 被 BSTR 边界弄坏（GBK → `?D??`） | 已封装 `cdr_text(paragraph_format=…)`；码页转换请用 Python `bytes.decode` |
| 136 | **连接器 / 符号库 / 曲线信息数组都造不出来** | `Layer.CreateFreeConnector(x1,y1,x2,y2)` 报「发生意外」；`CreateSymbol(x,y,名,库)` 报「找不到该对象」；`Application.SymbolLibraries` 直接 com_error（本机未装符号库）；`Curve.PutCurveInfo(tuple, n)` 报 `MemoryError: CreatingSafeArray` | 连接线用两点直线/贝塞尔自绘；符号用 `cdr_shapes` 植物件；曲线信息用 `GetCurveInfo()` |
| 137 | **`GetPointPositionAt` 的偏移类型是 0/1 而不是枚举名** | `SubPath.GetPointPositionAt(0, 0, 偏移, 类型)`：类型 **0 = 绝对长度（mm）**、**1 = 归一化 0~1**；前两个入参被忽略，返回 `(X, Y)` | 已封装 `cdr_path(point_at_offset, params={"offset":…,"normalized":true})` |
| 138 | **`Layer.CreateBSpline` 要的是 `IVGBSpline` 不是 Curve** | 第 32 波把它判成「类型不匹配」是因为传了 `Curve`。正确链路：`Document.CreateBSpline(点数, 闭合)` → `bs.InsertControlPoint(序号, x, y, Clamped)`（**首尾必须 Clamped=True**，否则报 「An end point's Clamped must be true」）→ `Layer.CreateBSpline(bs)` 得到 Type **21**，`ConvertToCurves()` 后变 Type 3 | 已封装 `cdr_create(type="bspline", points=…, to_curves=…)`；`SHAPE_TYPES` 补了 `21: "bspline"` |
| 139 | **解组家族不是空操作，`GetNext/GetPrevious` 也不是那个 3 参属性** | 第 30 波把 `ShapeRange.UngroupAll/UngroupAllEx/UngroupEx` 判成空操作，其实是**拿组合体内部对象**去解组（自然没反应）；对组合体本身调用三者都生效（页码 1→2 / 1→3）。`Shape.GetNext(方向)`、`GetPrevious(方向)` 各只要 1 个参数就返回同层相邻对象，与之前记的「`Shape.Next` 是 3 参 propget」是两个不同成员 | 已封装 `cdr_shape_ops(op="ungroup_all"/"ungroup_all_ex"/"ungroup_ex"/"sibling")` |
| 140 | **连接器/标注的端点是 `IVGSnapPoint`，由 `CreateSnapPoint(x,y)` 产出** | `Layer.CreateConnector/CreateLinearDimension` 要的端点不是 `IVGPoint`（它没有可写的 x/y），而是 `Application.CreateSnapPoint(x, y)` 返回的 `IVGSnapPoint`。拿到它之后连接器三件套 + 线性/角度标注（22 参 / 18 参）一次全通 | 已封装 `cdr_create(type="connector"/"right_angle_connector"/"linear_dimension"/"angular_dimension")`；`SHAPE_TYPES` 补 14 与 19 |
| 141 | **`CreateCurveMappedToStroke` 的 Stroke 要 `IVGSubPath` 不是 `IVGCurve`** | 第 32 波判「类型不匹配」是因为传了整条 Curve。正确：`Curve.CreateCurveMappedToStroke(子路径, ScaleY, SelfWeld)`；参考线版再要两个 `IVGPoint`（`Node.GetPoint()`）。返回新 Curve，`Layer.CreateCurve` 后成 Type 3 | 已封装 `cdr_path(action="map_to_stroke", other=骨架线, params={"scale_y":…,"use_reference_line":…})` |
| 142 | **文字 `*InRange` 结构体：方法名无 `Get` 前缀 + StartIndex 1 基** | 第 37 波判它们"字段全是垃圾"其实是两个错叠加：正确写法是 `Text.FontPropertiesInRange(start, count, indexing)`（没有 `Get` 前缀，`GetFontPropertiesInRange` 根本不存在）；且 `StartIndex` **从 1 开始**（传 0 才返回 `Size=0 / Style=16` 的假值）。写入端 `Set*PropertiesInRange` 同样 1 基。另外 `SetFontProperties(frames, 结构体)` 不报错但对可见文本**无效** | 已封装 `cdr_text(range_struct={…})` 与 `paragraph_format={"range_start":1,…}` |
| 143 | **色彩上下文/打样的 `RenderingIntent` 是 1~4；调色板工厂第一参是 GUID** | `CreateColorContext2(名称串, intent, model)` 与 `CreateProofColorSettings(名称, intent, bool)` 的 `RenderingIntent` **有效范围 1~4**（传 0 报「超出范围」，第 18 波就是这样判死的）；`CreatePaletteColor`/`CreateSpotColor` 第一参要**调色板 GUID**（之前传名字被判「取色不是新建」）。`CreateSpotColor` 还需要含专色的调色板 | 已封装 `color_context_create` / `proof_color_settings` / `object_palette_factory` / `object_from_spot` |
| 144 | **点集参数要的是 `IVGPointRange` 对象，不是 Python list** | `Curve.AppendSubpathFromPoints(list)` 报 `TypeError: Python instance can not be converted`，一直被判"数组入参不可用"。其实参数类型是 `IVGPointRange`，而它**有现成来源**：`SubPath.GetEvenlySpacedPoints(间距, 缩放)` 直接返回，`AddPointXY`/`RemoveAll` 可随意造点集。拿到后 `CreateCurveFitToPoints(AndCusps)`、`AppendSubpath*Points` 全部可用 | 已封装 `cdr_path(action="fit_to_points"/"append_points")` |
| 145 | **`ApplyHatchFill` 的第 2 参必须是整数序号** | `Fill.ApplyHatchFill(库名, "ANGLE", …)` 报「发生意外」被判不可用；传 **int**（0/1/2/5/10/50…均可）、库名传空串即可 → `Fill.Type = 10`。`ApplyCustomHatchFill` 18 参的 Style 可传 None。`Document.HatchLibraries` 报错不影响填充 | 已封装 `cdr_modify(fill={"type":"hatch"/"hatch_custom"})` |
| 146 | **`GetMatrix` 也会让 `IsOnShape` 失效** | 之前只记了 SetMatrix；长跑里「先 GetMatrix 再 IsOnShape」会稳定返回 0（单独跑正常）。把命中判定挪到矩阵读取之前即恢复 | 冒烟里已把 `is_on_shape` 放到 `matrix` 之前；调用方若两者都用，注意这个顺序 |
| 147 | **`SetSizeEx`/`StretchEx` 前两参是中心点；`AllExcluding` 要索引数组（1 基）** | 第 33 波把 `SetSizeEx`/`StretchEx` 判成「参数语义不明、会放大到离谱尺寸」，其实是把**中心点**当左上角传：`SetSizeEx(cx, cy, 宽, 高)`、`StretchEx(cx, cy, X倍数, Y倍数, 缩放文字)`。`AllExcluding([1])` 接受 Python list、索引 **1 基**（传 None 才报 TypeError） | 已封装 `cdr_shape_ops(op="set_size_ex"/"stretch_ex"，可加 use_range) / "all_excluding"` |
| 148 | **`OpenDocumentAsCopy` 需要第二个「打开选项」参数** | 只传路径报 `TypeError`；加 `CreateStructOpenOptions()` 即可，返回的是**未命名副本**（`Name` = 未命名 -N、`FileName` 为空），内容与源文件一致 | 已封装 `cdr_open(path=…, as_copy=true)`；`CreateStructFontProperties()` 也能造可读结构体（但 `SetFontProperties` 写入无效） |
| 149 | **`Style.SetBasedOn` 要名称串、`Style.Assign` 要样式对象** | `SetBasedOn(样式对象)` 报类型不匹配（传 None 报「未找到样式 None」）—— 参数其实是**名称字符串**；`Assign(形状)` 同样错，它要**另一个 IVGStyle**。另：同一对象同类别连续 `CreateStyleFromShape` 会互相覆盖；`SetBasedOn` 之后样式的 Name 会变 | 已封装 `cdr_style(action="set_based_on"/"assign")`；建样式请用不同对象 |
| 150 | **`Text.ImportFromFile` 会弹模态框卡死** | `Text.ImportFromFile(path, 1, 0)` 调用后 240 s 未返回（模态导入对话框，COM 无法中断） | 🚫 严禁调用；要导入文本请读文件后用 `cdr_text(text=…)` |
| 151 | **对象接口只能通过属性取，QueryInterface 一律失败** | 标注/图框/辅助线/文本框这些接口（`IVGDimension`/`IVGPowerClip`/`IVGGuide`/`IVGTextFrame`）对自动化对象做 `QueryInterface` 都报「不支持此接口」。正确入口是属性：`Shape.Dimension`、`Shape.PowerClip`、`Shape.Guide`、`Text.Frames.Item(i)` | 已封装 `dimension_info/dimension_set`、`powerclip_extract`、`guide_info/guide_set`、`cdr_text(link_to/unlink)`；`TextFrame.SetColumns` 要 SAFEARRAY（`[25,25]` 报类型不匹配）暂不可用 |
| 152 | **`ActiveView` / `Status` / `PaletteManager` 的入口是属性** | `Application.ActiveWindow.ActiveView` 是 `IVGActiveView`（15 个缩放/取景方法，`SetActualSize`→100%、`ZoomIn`→200%、`ToFitShape`→1963% 实测有效）；`Application.Status` 是 `IVGAppStatus`（进度条四件套，`BeginProgress` 第二参是 **int**，传字符串报 `ValueError`）；`Application.PaletteManager` 是 `IVGPaletteManager`（PaletteCount=66、GetPalette、LoadAllPalettes） | 已封装 `cdr_view` 缩放系列、`cdr_doc(app_status/palette_manager)` |
| 153 | **`IVGToolShapeAttributes` 的 12 个 `Set*` 只是配置，注册不了工具** | `CreateToolShapeAttributes()` 造出的对象上 `SetCanResize/SetCanRotate/SetCanSkew/…/SetDefaultShapename` 12 个方法都能调（12/12），但 X8 的 `RegisterToolShape` 不可用，所以配置完无法真正注册成工具。另 `IVGOutlineStyle` 的 `DashCount`/`DashLength(i)`/`GapLength(i)` **只读可用、写入不生效** | 已封装 `cdr_doc(tool_shape_attributes)` 与 `cdr_doc(outline_style_dash)`，并在返回值里注明"只配置不注册" |
| 154 | **克隆体的 Name 与源对象完全相同** | `Shape.Clone(dx, dy)` 生成的克隆体`Name` 与源对象一模一样（`inspect` 里会出现两个同名对象），按名字找不到克隆体，`CloneLink` 也就无从下手。另 `page.Grid` 不存在（网格走 `cdr_doc(grid/set_grid)`）、`shape.Symbol` 需要符号库形状、`Style.Fill.SaveFill/LoadFill` 参数形态未通 | `cdr_modify(clone=…)` 已强制给克隆体改名；圆角/克隆链/投影偏移分别由 `corners` / `clone_links` / `drop_shadow` 三个 op 提供 |
| 155 | **`Dimension.Linear` 是对象不是布尔值** | 早期用 `bool(dim.Linear)` 探测，看起来"返回 True"，实际它是 **IVGDimensionLinear 对象**（`SetReverseTerminatorLength`、`ReverseTerminators` 都在里面）。另外 `Page.Grid` 不存在，`Document.Grid` 才是 `IVGGrid` | 已封装 `cdr_doc(set_grid, settings={"frequency":…})` 与 `cdr_shape_ops(dimension_set, params={"reverse_terminator_length":…})` |
| 156 | **`StyleFill.SaveFill` 第二参是 `IVGFillMetadata` 对象；`TreeManager` 在 Document 上** | `SaveFill(path, True)` / `(path, 0)` 都报 `TypeError`，正确第二参是 `Application.CreateFillMetadata()` 造出来的元数据对象；`LoadFill` 同理。对象树入口是 `Document.TreeManager`（`Application` 上没有 `TreeManager`） | 已封装 `cdr_style(fill_file)` 与 `cdr_doc(tree_manager)`；`IVGExtrudeVanishingPoint.Share` 传另一消失点仍「类型不匹配」，不暴露 |
| 157 | **UI 自动化整族因缺 GUID 不可达** | `Application.FrameWork`（ICUIFrameWork）→ `.Automation`（ICUIAutomation）能拿到对象，但所有方法都要**命令栏/菜单项 GUID**：`ICUICommandBar` 没有 `Guid` 属性，`GetActiveMenuItemGuid(i)` 在无菜单打开时返回空串。另 `ColorContext.SetColorProfile(模型号, 配置文件)` 要求**配置文件的 ColorModel 与模型号一致** | UI 自动化不暴露；色彩配置文件已封装 `cdr_doc(set_color_profile)`（自动按模型挑配置） |
| 158 | **同一批 TLB 里 Get 类方法的参数风格不一致** | `IVGHatchFill.GetFillScale()` / `GetLineScale()` 是**零参**（传 `(0.0)` 会 TypeError），而 `IVGOutlineStyle.DashLength(i)` 要索引；`IVGEllipse.GetRadius()` 零参返回 `(x, y)` tuple。另 `Palettes.CreateFromSelection` 在选区里没有可取色对象时返回空，需退回整篇文档 | 已按"先试零参、再试带参"的写法封装 `hatch_fill` / `ellipse_geometry` / `palette_from` |
| 159 | **`Shape.CreateEnvelope` 的 PresetIndex 是 1 基** | 传 0 报「参数 PresetIndex 超出范围」，传 1 才能建出包络；`effect.Envelope` 的 `Select(预设)` / `CreateFrom(形状)` / `CopyFrom(另一包络)` / `CopyFromCurve(曲线,…)` 都可用，但 `CopyFromShape` 的五参形态未通 | 已封装 `cdr_shape_ops(op="envelope_object")`；屏幕手柄走 `cdr_view(action="overlay_handle")` |
| 160 | **`IVGExtrudeVanishingPoint.Share` 要传拉伸效果而不是消失点** | `Share(另一个消失点)` 报「类型不匹配」，`Share(effect.Extrude)` 才返回布尔；另 `IsOnShape` 在视口缩放异常时会返回 0（长跑里偶发，测试前先 `actual_size` 归位） | 已封装 `cdr_shape_ops(op="extrude_vp_share")`、`cdr_doc(action="windows", settings={"arrange":1})` |
| 161 | **`FountainColors.AddGrayLevel` 仅灰度渐变可用** | 彩色渐变调用`AddGrayLevel(灰度[, 位置])` 报「非选择性的参数」；色标数由渐变类型决定（双色渐变恒为 2），`Item(2)` 之外会「超出范围」。另 `SkewEx` 在范围上也真实生效 | 已封装 `cdr_shape_ops(range_skew)` 与 `fountain_colors` 的 `stops` 枚举；`AddGrayLevel` 失败转为 warnings |
| 162 | **调色板方法参数形态各异 + 固定调色板只读** | `FindColor(名字串)` 对 TRUMATCH 常见名一律返回 0；`InsertColor(序号, IVGColor)` **序号在前**；`OpenFixed` 打开的固定调色板 `SetColor`/`InsertColor` 都报「无法修改固定调色板」 | 已封装 `palette_find_color`/`palette_insert_color`/`palette_make_default`；`insert_color` 失败自动退化为新建自定义调色板 |
| 163 | **立体化灯光是带索引的属性，`SetLight` 三参** | `LightPresent(i)` / `LightIntensity(i)` / `LightPosition(i)`（i=1~3）各自独立；`SetLight(索引, cdrLightPosition, 强度)` 第二参直接传 `LightPosition(索引)` 取回的对象；未开的灯 `LightIntensity(i)=-1`。`SetBevel(深度, 角度, 只显示斜角)` 一次设斜角 | 已封装 `cdr_shape_ops(extrude_set, effect={"light_index":…, "bevel_depth":…})`，返回 3 盏灯状态 |

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
python -m pytest -q                # 93 项逻辑测试，任何平台可跑
python scripts/live_smoke.py       # 553 项端到端（跑前先确认没有残留 CorelDRW 进程）
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
