# cdrX8_mcp

让 AI Agent 直接操作 **CorelDRAW X8 (v18)** 的 MCP 服务器。

通过 CorelDRAW 的 COM 对象模型驱动真实文档：新建、画图、上色、改字、全文档替换、导出、存盘、撤销。
所有坐标统一 **毫米 / 页面左下角原点 / Y 轴向上**，每次变更都收成**一个原子 undo 步**。

> **已在真机验证**：CorelDRAW Graphics Suite X8 `18.0.0.450`（64 位）+ Python 3.11.9 64 位。
> 端到端联调 19/19 通过，单元测试 26 项通过。
> 本项目里的每个常量与 API 写法都是在这台机器上实跑出来的——**不是照抄新版文档**（X8 与 2024+ 差异不小，见[下文](#x8-特有问题)。

---

## 快速开始

```powershell
# 1. 安装（需要 64 位 Python，与 CorelDRAW 位数一致）
cd "C:\path\to\cdrX8_mcp"
python -m pip install -e .

# 2. 自检：单元测试（不需要 CorelDRAW）
python -m pytest -q

# 3. 端到端联调（会启动 CorelDRAW，跑完自动关闭文档）
python scripts/live_smoke.py

# 4. 启动 MCP 服务器（stdio）
python -m cdrx8_mcp
```

### 客户端配置

Claude Desktop / Cursor 等通用 stdio 客户端：

```json
{
  "mcpServers": {
    "coreldraw-x8": {
      "command": "C:\\Python311\\python.exe",
      "args": ["-m", "cdrx8_mcp"],
      "env": {
        "CDR_ALLOWED_ROOTS": "D:\\design\\out;C:\\Users\\me\\Desktop\\out",
        "CDR_STRICT_REVISION": "1"
      }
    }
  }
}
```

完整示例见 [`examples/`](examples/)。

---

## 工具一览（13 个）

| 工具 | 作用 | 会改文档 |
|---|---|---|
| `cdr_status` | 查看安装/运行状态与服务配置（无副作用） | — |
| `cdr_connect` | 附着已运行的 CorelDRAW；`launch=true` 才允许启动 | — |
| `cdr_documents` | 列出打开的文档，给出 `document_id` | — |
| `cdr_open` | 打开 `.cdr`，或按毫米尺寸新建空白文档 | — |
| `cdr_inspect` | **读视图**：页面、`revision`、`snapshot_token`、每个对象的 bbox 与文字 | — |
| `cdr_create` | 批量创建 矩形 / 椭圆 / 艺术字 / 段落文字 | ✅ 一个 undo 步 |
| `cdr_modify` | 移动、改尺寸、旋转、填充、描边、改字、改名、层级、复制、删除 | ✅ 一个 undo 步 |
| `cdr_replace_text` | 全文档文字替换（可跨页、可选大小写） | ✅ 一个 undo 步 |
| `cdr_pages` | 页面 列表 / 新增 / 切换 / 删除 / 改尺寸 | 部分 |
| `cdr_export` | 导出 PNG / SVG / PDF（当前页） | — |
| `cdr_save` | 另存为 `.cdr`（默认不覆盖） | — |
| `cdr_undo` | 撤销最近 N 步 | ✅ |
| `cdr_close` | 关闭文档（默认丢弃未保存改动） | — |

### 标准调用顺序

```
cdr_status            → 看有没有在跑
cdr_connect           → 附着（或 launch=true 启动）
cdr_open              → 新建 100×60mm 文档，拿 document_id
cdr_inspect           → 拿 revision + snapshot_token + 现有对象
cdr_create            → 画东西（带上 revision/token）
cdr_inspect           → 复核（拿到新的 revision/token）
cdr_modify            → 改（带新的 revision/token）
cdr_export / cdr_save → 出图 / 存档
```

**`revision` 和 `snapshot_token` 每次变更都会变**，必须用上一步返回的新值。这是防"Agent 基于过期视图瞎改"的乐观锁（详见[安全模型](#安全模型)）。

---

## 坐标与单位约定

| 项目 | 约定 |
|---|---|
| 单位 | 毫米（内部自动把文档单位切到 mm 再切回，**不会**改动你的文档设置） |
| 原点 | 页面**左下角**，Y 轴**向上** |
| 矩形 | `{"x": 10, "y": 10, "width": 50, "height": 30}` = **左下角 + 宽高** |
| 椭圆 | `{"cx": 78, "cy": 25, "rx": 12, "ry": 12}` = **圆心 + 半径** |
| 文字 | `{"x": 10, "y": 48, "text": "标题"}` = 文字锚点 |
| 颜色 | `"#RRGGBB"` 或 `{"c":0,"m":100,"y":100,"k":0}`（印前请用 CMYK）或 `{"r","g","b"}` |
| 描边 | `"outline": 0.5`（宽度）或 `{"width": 0.5, "color": "#000000"}` |
| 字号 | 点（pt） |

---

## 安全模型

设计目标：**Agent 可以放心改图，用户随时能反悔。**

| 机制 | 说明 |
|---|---|
| **乐观锁** | 修改类工具接受 `expected_revision` + `snapshot_token`；对不上直接拒绝并提示重新 `cdr_inspect`。令牌里带了页数与对象总数，所以**用户手动增删对象**也会被检测到 |
| **原子 undo** | 一次工具调用 = 一个 CorelDRAW undo 步。`cdr_create` 画 20 个对象，一次 `cdr_undo` 全部回滚（实测 3→5→3） |
| **路径白名单** | `CDR_ALLOWED_ROOTS` 设定后，导出/保存只能写到这些目录下 |
| **覆盖保护** | 目标文件已存在时默认报错，必须显式 `overwrite=true` |
| **不自动启动** | `cdr_connect` 默认只**附着**已有实例；启动进程要显式 `launch=true`（可用 `CDR_ALLOW_LAUNCH=0` 全局禁止） |
| **不碰源文件** | `cdr_save` 只做"另存为"，源 `.cdr` 全程只读 |
| **只监听本机** | HTTP 模式默认 `127.0.0.1` |

严格模式：`CDR_STRICT_REVISION=1` 时，**不带** `revision`/`token` 的写操作直接拒绝（默认只给 warning 放行）。

---

## 环境变量

| 变量 | 默认 | 说明 |
|---|---|---|
| `CDR_ALLOWED_ROOTS` | 空 | 允许写入的目录，`;` 分隔。空 = 不限制（每次写操作会返回 warning） |
| `CDR_STRICT_REVISION` | `0` | `1` = 强制要求 revision/token |
| `CDR_ALLOW_LAUNCH` | `1` | `0` = 永远不启动 CorelDRAW，只允许附着 |
| `CDR_TIMEOUT` | `60` | 单次 COM 调用超时（秒） |
| `CDR_TRANSPORT` | `stdio` | `http` 可起本地 HTTP（配合 `CDR_HOST` / `CDR_PORT`，默认 `127.0.0.1:8765`） |

---

## X8 特有问题

这些是 **X8 与 2024+ 的差异**，也是本项目存在的主要理由。全部实测确认：

| # | 现象 | 正确写法 | 踩错的后果 |
|---|---|---|---|
| 1 | 单位枚举不是常规认知 | `doc.Unit = 3` 才是毫米（`2` 是**英尺**、`4` 是**厘米**） | 尺寸放大 10 倍 / 304.8 倍 |
| 2 | `Document` 上没有页面属性 | `doc.ActivePage.SizeWidth / SetSize()` | `AttributeError` |
| 3 | `Page` 上没有 `Document` 属性 | 用外部传入的 `doc` 做 undo 分组 | `AttributeError: <unknown>.Document` |
| 4 | 没有 `CreateArtisticText2` | `layer.CreateArtisticText(x, y, text)` | `AttributeError` |
| 5 | `SaveAs` 必须两个参数 | `doc.SaveAs(path, None)` | `TypeError: Python instance can not be converted to a COM object` |
| 6 | 字体不在 `shape.Text` 上 | `shape.Text.Story.Font / .Size / .Text` | 属性不可设置 |
| 7 | 没有 `SetActivePage` | `doc.Pages.Item(i).Activate()` | `AttributeError` |
| 8 | 没有 `app.CreateExportOptions` | 无法通过 COM 指定导出 DPI | 导出固定走 Corel 默认 300dpi |
| 9 | CDR 没有可用的 `ExportEx` 过滤器 | 存盘只能 `SaveAs` | "指定的过滤器 ID 无效" |

改代码前请先读 [`src/cdrx8_mcp/constants.py`](src/cdrx8_mcp/constants.py) 里的注释。

---

## 已知限制

* **导出只到当前页**：PNG / SVG 用 `ExportEx(range=1)`；多页请逐页切换后导出。
* **无法指定导出 DPI**：X8 没有 `CreateExportOptions`，默认 300dpi。
* **对象寻址靠 index / name / 选区**：X8 的 COM 对象没有稳定 GUID，服务在创建时给对象起 `cdr_xxxxxxxx` 的名字，建议后续按 `name` 定位（按 `index` 会因插入/删除而漂移）。
* **无编程接口的能力**（版本无关，别浪费时间）：刀刻、橡皮、网格填充、智能绘图、表格、字形——这些只能用鼠标。
* **COM 不可中断**：CorelDRAW 卡在模态弹窗（缺字体/色彩配置）时，调用会超时但进程仍在跑。**写操作超时后不要自动重试**。
* **单实例串行**：所有调用走同一条 STA 线程，天然不支持并发（这是安全特性，不是缺陷）。

---

## 测试

```powershell
python -m pytest -q              # 26 项纯逻辑测试，任何平台可跑
python scripts/live_smoke.py     # 19 项端到端，需要 CorelDRAW X8
python scripts/live_smoke.py --keep   # 保留文档方便肉眼看效果
```

---

## 致谢

设计思路参考了这些开源项目（本仓库为独立实现，未复制其代码）：

* [CubicDev1/corel-mcp](https://github.com/CubicDev1/corel-mcp) — 14 工具面、原子 undo、快照令牌的设计取舍
* [hmsusbusas-png/mcp-corel](https://github.com/hmsusbusas-png/mcp-corel) — 单 COM 线程 + 三档超时
* [pefedotov/coreldraw_mcp](https://github.com/pefedotov/coreldraw_mcp) — 工具箱能力矩阵（哪些能做、哪些只能鼠标）
* [logicbyroshan/foldermate](https://github.com/logicbyroshan/foldermate) — 进程外隔离与看门狗思路

## 许可

MIT。本项目只通过 COM 驱动本机已授权的 CorelDRAW，不包含、不分发、不修改任何 Corel 软件。
CorelDRAW 与 Corel 是 Corel Corporation 的商标。
