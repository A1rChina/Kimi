# 项目结构规划

## 目标

项目最终承担三件事：

1. 从 CMMS 页面抓取原始数据。
2. 对抓取结果做标准化、增量去重和水位线维护。
3. 将稳定后的记录写入 Cloudflare D1。

因此项目结构应按“抓取、整理、入库”分层，而不是让主脚本继续堆积所有数据源逻辑。

## 目标目录

```text
Kimi/
├── config/
│   ├── defaults.json
│   └── sources.json
├── scripts/
│   └── scrape_cmms_exports.py
├── src/
│   └── kimi/
│       ├── auth.py
│       ├── browser.py
│       ├── config.py
│       ├── extractors/
│       │   ├── excel_export.py
│       │   └── html_table.py
│       ├── normalize.py
│       ├── incremental.py
│       ├── outputs.py
│       └── d1.py
├── db/
│   └── migrations/
├── data/
│   ├── raw/
│   ├── normalized/
│   ├── latest/
│   ├── state/
│   └── debug/
└── docs/
```

## 配置职责

`config/sources.json` 是数据源契约，保存：

- 数据源名称和抓取顺序。
- URL 对应的环境变量名。
- 页面类型和选择器。
- 增量模式、水位线字段、回看天数、去重字段。
- 本地输出目录和 D1 目标表。

配置中不保存真实 URL、IP、账号或密码。所有敏感值从 GitHub Secrets 或本地环境变量读取。

## 模块职责

`scripts/scrape_cmms_exports.py`

命令行入口，只负责解析参数、读取配置、调度流程。

`src/kimi/config.py`

读取并校验 `config/defaults.json` 和 `config/sources.json`，把配置转换为运行对象。

`src/kimi/auth.py`

处理登录流程。

`src/kimi/browser.py`

封装 Playwright 公共能力，例如 frame 查找、截图、等待下载。

`src/kimi/extractors/excel_export.py`

处理有“开始查询”和“导出Excel”的页面。

`src/kimi/extractors/html_table.py`

预留给没有 Excel 导出按钮、只能解析页面表格的数据源。

`src/kimi/normalize.py`

执行字段名清理、空行空列处理、日期和数值标准化。

`src/kimi/incremental.py`

维护 `data/state/<source>.json`，生成查询窗口，计算 `record_key`，过滤重复记录。

`src/kimi/outputs.py`

保存 raw、normalized、latest、debug 文件。

`src/kimi/d1.py`

负责 Cloudflare D1 写入。建议以 `record_key` 为主键执行 upsert。

## 迁移步骤

### 第一步：配置驱动抓取

把当前 `scripts/scrape_purchase_inbound.py` 中硬编码的数据源列表改为读取 `config/sources.json`。

### 第二步：拆分输出目录

将当前 `data/excel_export/<source>/latest.*` 逐步迁移为：

```text
data/raw/<source>/
data/normalized/<source>/
data/latest/<source>.json
data/state/<source>.json
data/debug/<source>/
```

### 第三步：增量整理

对每个 source 明确：

- `watermark_field`
- `lookback_days`
- `key_fields`

如果字段还不确定，先保留空数组，后续根据实际导出表头补齐。

### 第四步：写入 D1

新增 D1 迁移 SQL，并用 `record_key` 做唯一键。写入流程放在增量整理之后，只写 normalized 记录。

### 第五步：保留重建能力

GitHub Actions 保留手动参数：

```text
mode = full | incremental
source = all | <source_name>
```
