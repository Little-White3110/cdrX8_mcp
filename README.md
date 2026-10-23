# cdrX8_mcp

让 AI Agent 直接操作 **CorelDRAW X8 (v18)** 的 MCP 服务器。

通过 CorelDRAW 的 COM 对象模型驱动真实文档：画图、排版、上色、布尔运算、对齐、图层、导入导出、存盘、撤销。
所有坐标统一 **毫米 / 页面左下角原点 / Y 轴向上**，每次变更都收成**一个原子 undo 步**。

> **已在真机验证**：CorelDRAW Graphics Suite X8 `18.0.0.450`（64 位）+ Python 3.11.9 64 位
> - 单元测试 **43 项通过**
> - 端到端联调（真实 CorelDRAW）**31/31 通过**
> - 能力边界来自对 `vgcoreauto.tlb` 的反解（692 个接口 / 6789 个方法）+ 逐项实跑

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

## 工具一览（20 个）

### 会话与文档

| 工具 | 作用 |
|---|---|
| `cdr_status` | 运行状态、版本、服务配置（无副作用） |
| `cdr_connect` | 附着已运行实例；`launch=true` 才允许启动 |
| `cdr_documents` | 列出打开的文桇，给出 `document_id` |
| `cdr_open` | 打开 `.cdr`，或按毫米尺寸 / 页面预设新建 |
| `cdr_inspect` | 读视图：页面、图层、`revision`、`snapshot_token`、每个对象 |
| `cdr_save` | 另存为 `.cdr`（默认不覆盖） |
| `cdr_undo` | 撤销最近 N 步 |
| `cdr_close` | 关闭文档（默认丢弃未保存改动） |

### 绘制与修改

| 工具 | 作用 |
|---|---|
| `cdr_create` | 批量创建：矩形（含圆角）/ 椭圆（含圆弧、扇形）/ 多边形 / 星形 / 螺旋 / 网格 / 直线 / 贝塞尔 / 文字 / 段落文字 / 参考线 |
| `cdr_modify` | 移动、缩放、旋转、填充（纯色/渐变/纹理/无色）、描边、投影、轮廓图、封套、改字、层级、复制、删除 |
| `cdr_text` | 排版级文字操作：内容 / 字体 / 字号 / 粗斜体 / 下划线 / 对齐 / **转曲** |
| `cdr_replace_text` | 全文档文字替换（底层用 X8 原生 `Page.TextReplace`） |
| `cdr_select` | 按名字 / 类型 / 矩形框选，加入选择集 |
| `cdr_shape_ops` | **布尔与整形**：weld / trim / intersect / combine / group / ungroup / break_apart / to_curves / powerclip / fillet / chamfer / scallop |
| `cdr_align` | 对齐与分布（左中右 / 上下中 / 等间距，对齐对象或页面） |
| `cdr_find` | 按名字或类型查找（跨图层） |

### 图层 / 页面 / 文件

| 工具 | 作用 |
|---|---|
| `cdr_layers` | 图层列表 / 新建 / 激活 / 显隐 / 锁定 / 打印 / 移动 / 删除 |
| `cdr_pages` | 页面列表 / 新增 / 切页 / 删除 / 自定义尺寸 / **内置纸张预设**（A4、A3、名片…） |
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
"drop_shadow": {"opacity":60,"offset_x":3,"offset_y":-3,"color":"#444444"}
```

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

## X8 特有问题（14 条，全部实测）

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

改代码前请先读 [`src/cdrx8_mcp/constants.py`](src/cdrx8_mcp/constants.py) 里的注释。

---

## 已知限制（诚实清单）

**做不到：**
* **描边样式（虚线）与箭头没法设**：`Outline.Style = 1` / `StartArrow = 1` 都返回「类型不匹配」——X8 需要传对象而不是枚举，暂未打通
* **立体化 / 透镜 不可用**：`CreateExtrude` / `CreateLens` 参数组合没试通；轮廓图、封套、投影已验证可用
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
python -m pytest -q                # 43 项纯逻辑测试，任何平台可跑
python scripts/live_smoke.py       # 31 项端到端，需要 CorelDRAW X8
python scripts/live_smoke.py --keep   # 保留文档方便肉眼看效果
```

联调覆盖：页面预设 → 多图元创建 → 文字排版 → 投影/描边 → weld/fillet → 框选/组合/解组 → 对齐 → 图层 → 查找 → 分页/预设 → 四格式导出（含 DPI 倍增校验）→ 导入 → 文字替换 → 存盘 → 撤销 → 过期 token 拒绝 → 关闭。

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
