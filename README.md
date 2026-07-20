# 物料盘点 PWA

> 月度物料盘点工具 · 本地 SQLite 存储 · 完全离线可用 · 可「添加到主屏幕」当原生 App 用

一个**纯前端、零后端**的物料盘点应用。所有数据只存在你自己的设备里，不上传任何服务器，适合一人或一个小团队做月度/周期盘点。

---

## 一、主要功能

- **物料档案**：物料的增 / 改 / 删除，支持回收站与「批量删除」
- **盘点单**：新建盘点单、录入物料数量、本单内查找、多选批量加/删
- **数据导入导出**：支持 xlsx 批量导入与导出（方便从 Excel 盘点表迁移）
- **离线优先**：首次打开后缓存到本地，无网络也能用
- **自动保存**：每次操作防抖落盘，关页面也不丢数据

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

---

## 四、技术架构

| 方面 | 方案 |
|------|------|
| 前端 | 原生 JS 单页应用（SPA），无框架，轻量 |
| 数据库 | `sql.js`（WebAssembly 版 SQLite）+ `IndexedDB` 持久化 |
| 离线 | Service Worker 缓存（当前版本 `stocktake-pwa-v25`） |
| 安装 | PWA `manifest.webmanifest` + 图标 |
| 托管 | GitHub Pages |
| 部署 | GitHub Actions 自动部署（push 即上线） |

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
│       └── deploy.yml      # GitHub Actions 自动部署配置
├── css/
│   └── styles.css          # 样式
├── js/
│   ├── app.js              # 主逻辑（SPA 路由、各视图交互）
│   ├── db-core.js          # 数据库核心（建表/增删改查/迁移）
│   ├── db.js               # 本地持久化（IndexedDB 读写、备份恢复）
│   ├── import.js           # xlsx 导入
│   └── export.js           # xlsx 导出
├── vendor/
│   ├── sql-wasm.js / .wasm # SQLite 的 WASM 运行时
│   └── xlsx.full.min.js    # Excel 解析库
└── assets/                 # 图标（192/512/png、apple-touch-icon）
```

---

## 六、部署与更新（维护者阅读）

- **仓库**：`github.com/likuan250-hash/stocktake-pwa`
- **部署方式**：GitHub Actions（`.github/workflows/deploy.yml`）。向 `main` 分支推送即自动构建并发布到 GitHub Pages。
- **关键文件 `.nojekyll`**：必须存在，否则 GitHub 的 Jekyll 会错误处理 `.wasm` 等二进制文件导致白屏。
- **改完代码**：`commit` + `push` 到 `main` → Actions 跑完 → 刷新页面即是新版（Service Worker 会后台更新缓存）。
- **旧地址**：之前用的 CloudStudio 托管地址已废弃，以本 GitHub Pages 地址为准。

---

## 七、自测

仓库外的自测脚手架 `_qa_selftest.js`（基于 jsdom + sql.js）覆盖搜索、仓库筛选、勾选批量加、数量±原地更新、本单查找、物料档案/盘点单批量删除、明细行选择模式等 **31 项**端到端校验，每次改动后会先自测再发布。

---

## 版本

当前版本：**v25**（对应 `sw.js` 中缓存名 `stocktake-pwa-v25`，每次发布递增）。
