# 物料盘点 PWA

> 月度物料盘点工具 · 本地 SQLite 存储 · 完全离线可用 · 可「添加到主屏幕」当原生 App 用

一个**纯前端、零后端**的物料盘点应用。所有数据只存在你自己的设备里，不上传任何服务器，适合一人或一个小团队做月度/周期盘点。

---

## 一、主要功能

- **物料档案**：物料的增 / 改 / 删除，支持回收站与「批量删除」
- **盘点单**：新建盘点单、录入物料数量、本单内查找、多选批量加/删
- **数据导入导出**：
  - 从**金蝶盘点单**按单号导入物料（静态同步库，见第六节）
  - xlsx 批量导入 / 导出（从 Excel 盘点表迁移）
  - 两种导入方式可选：**更新添加**（按编码合并，默认安全）/ **清空后导入**（先删全部再导入，二次确认 + 失败自动回滚）
- **离线优先**：首次打开后缓存到本地，无网络也能用
- **自动保存**：每次操作防抖落盘，关页面也不丢数据
- **版本与更新**：右上角常驻显示当前版本号，点开可手动「检测更新」并一键刷新到最新（自动更新条未弹出时的可靠入口）

---

## 二、在手机上使用（iPhone）

1. 用 **Safari** 打开：**https://likuan250-hash.github.io/stocktake-pwa/**
2. 点底部 **分享** 按钮 → **「添加到主屏幕」**
3. 桌面会出现「物料盘点」图标，点开即是全屏 App，可离线使用

> Android / 桌面 Chrome 同理：菜单 →「安装应用 / Add to Home Screen」。

---

## 三、数据与安全

- 所有数据存在**手机浏览器本地**（`IndexedDB` + `sql.js` / WebAssembly 版 SQLite），**不上传任何服务器**。
- 换手机、清浏览器缓存会丢失本地数据；重要数据请用应用内「备份」导出 JSON 留存。
- 多台设备之间数据**各自独立**，不会自动同步（这是刻意的隐私设计）。
- **首次安装 / 首次更新需联网**：应用启动时需联网下载 WASM 运行时与静态资源，并注册 Service Worker。下载完成后即进入离线可用状态；之后断网也能正常使用。若离线缓存注册失败，应用会提示「离线缓存不可用」，此时需保持联网使用。

> 「清空后导入」与「备份恢复」操作前都会二次确认；清空导入全程有整库快照，若中途失败会自动回滚到导入前状态，不会丢失既有数据。

---

## 四、技术架构

| 方面 | 方案 |
|------|------|
| 前端 | 原生 JS 单页应用（SPA），无框架，轻量 |
| 数据库 | `sql.js`（WebAssembly 版 SQLite）+ `IndexedDB` 持久化 |
| 离线 | Service Worker 缓存（当前版本 `stocktake-pwa-v34`；数据文件 `kingdee-sheets.js` 走 network-first 确保拿到最新金蝶库） |
| 安装 | PWA `manifest.webmanifest` + 图标 |
| 托管 | GitHub Pages |
| 部署 | GitHub Actions 自动部署（push 即上线，先跑数据层自测再发布） |

> 项目代码全部使用**相对路径**，因此部署在 `用户名.github.io/仓库名/` 这类子路径下也不会 404。

---

## 五、目录结构

```
物料盘点PWA/
├── index.html              # 入口页面
├── manifest.webmanifest    # PWA 配置（名称/图标/启动方式）
├── sw.js                   # Service Worker（离线缓存，含版本号）
├── .nojekyll               # 阻止 GitHub 用 Jekyll 处理 .wasm 等文件
├── .github/
│   └── workflows/
│       └── deploy.yml      # GitHub Actions：先跑 tests/selftest.js，再部署
├── css/
│   └── styles.css          # 样式
├── js/
│   ├── app.js              # 主逻辑（SPA 路由、各视图交互、导入编排）
│   ├── db-core.js          # 数据库核心（建表/增删改查/迁移/clearAllMaterials）
│   ├── db.js               # 本地持久化（IndexedDB 读写、整库备份恢复）
│   ├── import.js           # xlsx 导入解析
│   ├── export.js           # xlsx 导出
│   └── kingdee-sheets.js   # 金蝶盘点单静态同步库（由 scripts/gen_kingdee_sheets.py 生成）
├── scripts/
│   └── gen_kingdee_sheets.py  # 金蝶同步库生成脚本（见第六节 Runbook）
├── tests/
│   └── selftest.js         # 数据层自测（合并/清空/回收站/幂等）
├── vendor/
│   ├── sql-wasm.js / .wasm # SQLite 的 WASM 运行时
│   └── xlsx.full.min.js    # Excel 解析库
└── assets/                 # 图标（192/512/png、apple-touch-icon）
```

---

## 六、金蝶同步管线（维护者阅读）

物料档案页「从金蝶盘点单导入」依赖一个**静态同步库** `js/kingdee-sheets.js`（全局 `window.KINGDEE_SHEETS`）。该库由开发者经 **KingdeeMCP** 按单号拉取金蝶云星空盘点单（表单 `STK_StockCountInput`）后，用 `scripts/gen_kingdee_sheets.py` 生成为仓库内的只读 JSON。App 本身不直连金蝶——它只读取这个已生成的静态库，按单号选单导入。

### 重新生成同步库（Runbook）

> 适用场景：金蝶侧新增了已审核盘点单，需要纳入同步库。

1. **查单**：用 KingdeeMCP 查询盘点单
   - `form_id = STK_StockCountInput`
   - 过滤条件：`FStockOrgId.FName IN ('测试89','中央厨房') AND FDocumentStatus = 'C'`（仅已审核）
   - 记录每张单的 `单据编号(billNo)` / 组织 / 日期 / `FID`。
2. **拉分录**：用 `kingdee_view_bill` 拉每张单的物料分录，保存为「干净 entries JSON」，每行一个物料，字段为：
   ```
   { "code": "...", "name": "...", "spec": "...", "unit": "...", "warehouse": "..." }
   ```
   ⚠️ **关键坑（曾污染 1356 条规格）**：中央厨房单的 `Specification` 是
   `list-of-dicts` `[{Key:2052, Value:'中文规格'}]`，**务必取 `Key==2052` 的 `Value`**，
   直接 `str(list)` 会把整段 Python repr 写进规格字段。
3. **放数据**：把各单 entries JSON 放到 `scripts/_data/`，命名如 `pdzy016370_entries.json`。
   也可通过环境变量指定目录：`KD_DATA_DIR=/path/to/entries python scripts/gen_kingdee_sheets.py`。
4. **登记**：在 `scripts/gen_kingdee_sheets.py` 的 `CENTRAL` 列表里登记单号/组织/日期/fid/文件名；
   测试89 单（0010012）已写死在 `TEST89_RAW`。
5. **生成**：
   ```bash
   python scripts/gen_kingdee_sheets.py
   ```
   脚本会写出 `js/kingdee-sheets.js`（含生成日期、单号、去重后物料总数）。
6. **发布**：`git commit` + `git push` 到 `main` → Actions 跑自测 + 部署；用户端「发现新版本 → 立即刷新」即拿到新库
   （`kingdee-sheets.js` 走 network-first，无需等缓存过期）。

> 新增单号只需重复上述步骤，无需改 App 代码。同步范围受金蝶授权组织限制（当前：测试89 + 中央厨房全量）。

---

## 七、部署与更新（维护者阅读）

- **仓库**：`github.com/likuan250-hash/stocktake-pwa`
- **部署方式**：GitHub Actions（`.github/workflows/deploy.yml`）。向 `main` 分支推送即自动：① 跑 `node tests/selftest.js` → ② 构建并发布到 GitHub Pages。**自测不过则不会部署。**
- **关键文件 `.nojekyll`**：必须存在，否则 GitHub 的 Jekyll 会错误处理 `.wasm` 等二进制文件导致白屏。
- **改完代码**：`commit` + `push` 到 `main` → Actions 跑完 → 刷新页面即是新版（Service Worker 会后台更新缓存，并提示「发现新版本」）。
- **旧地址**：之前用的 CloudStudio 托管地址已废弃，以本 GitHub Pages 地址为准。

### 部署凭证安全（重要）

- **撤销旧 full-access token**：早期部署用过含 `admin:org` / `delete_repo` 的高权限 PAT，请在 GitHub **Developer settings → Personal access tokens** 中删除，不再使用。
- **推荐凭证（二选一，权限最小化）**：
  1. **Fine-grained PAT**：仅授权本仓库，权限勾选 `Contents(读写)` + `Pages(读写)` + `Workflows(读写)`；或
  2. **GitHub MCP 推送**：通过已连接的 GitHub 连接器 `push_files` 部署，无需在本地/Shell 保存任何 token。
- 切勿把 token 写入仓库文件、CI 日志或记忆中。

---

## 八、自测

- `tests/selftest.js`（Node + 仓库自带 `vendor/sql-wasm.js` + `js/db-core.js`）覆盖数据层核心逻辑：
  **合并模式新增/更新、清空模式整库替换、回收站保留、导入幂等、回收站同编码恢复语义、默认列表编码升序**等共 21 项断言。
- 每次推送 `main` 由 GitHub Actions 自动运行；本地亦可直接 `node tests/selftest.js` 复验。
- 导入编排层（`app.js` 的 `importMaterials`：二次确认、整库快照回滚、回收站同编码提示、模式选择弹窗）
  依赖浏览器 DOM/弹窗，由手动与真机验证覆盖。

---

## 版本

当前版本：**v34**（对应 `sw.js` 中缓存名 `stocktake-pwa-v34`，每次发布递增）。
