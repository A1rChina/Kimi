# Kimi

这是一个面向公司 CMMS / ERP Web 系统的数据抓取项目。

项目目标不是改造原系统，而是通过 GitHub Actions + Playwright 定时登录系统，自动查询并导出关键业务明细，将网页中的 Excel 数据沉淀为可复用的 CSV / JSON / JSONL 数据文件，供后续做库存、入库、出库、发货、项目进度和异常风险分析。

## 1. 项目定位

```text
CMMS / ERP 网页系统
→ GitHub Actions 定时运行
→ Playwright 自动登录
→ 打开指定业务页面
→ 输入业务时间范围
→ 点击“开始查询”
→ 点击“导出Excel”
→ 保存原始 Excel
→ 转换为 CSV / JSON / JSONL
→ 提交到仓库 data/ 目录
```

一句话：

```text
先把系统里真实存在的数据自动拿出来，形成自己的数据底座。
```

## 2. 当前已支持的数据源

当前脚本已接入 14 类业务页面：

| 数据源 | 脚本内部名称 | 说明 |
|---|---|---|
| 采购入库明细 | `purchase_inbound` | 按业务时间范围查询并导出 |
| 仓库明细 | `warehouse_detail` | 按业务时间范围查询并导出 |
| 仓库台账 / 库存台账 | `warehouse_ledger` | 当前不设置日期范围，直接查询导出 |
| 销售出库明细 | `sales_outbound_detail` | 按业务时间范围查询并导出 |
| 不良品跟踪 | `defective_tracking` | 当前不设置日期范围，直接查询导出 |
| 后工序报工跟踪 | `post_process_report_tracking` | 按 `text0` 时间范围查询并导出 |
| 生产节拍维护 | `production_takt_maintenance` | 当前不设置日期范围，直接查询导出 |
| 生产设备维护 | `production_equipment_maintenance` | 当前不设置日期范围，直接查询导出 |
| 设备停机查询 | `equipment_downtime_query` | 按 `text0` 时间范围查询并导出 |
| 工装更换查询 | `tooling_replacement_query` | 按 `text0` 时间范围查询并导出 |
| 机加转序查询 | `machining_transfer_query` | 按 `text0` 时间范围查询并导出 |
| 项目信息维护 | `project_info_maintenance` | 当前不设置日期范围，直接查询导出 |
| 产品材料匹配维护 | `product_material_match_maintenance` | 当前不设置日期范围，直接查询导出 |
| 周计划执行跟踪 | `weekly_plan_execution_tracking` | 当前不设置日期范围，直接查询导出 |

## 3. 项目结构

```text
Kimi/
├── .github/
│   └── workflows/
│       └── scrape-purchase-inbound.yml   # GitHub Actions 自动抓取流程
├── data/
│   └── excel_export/
│       ├── purchase_inbound/             # 采购入库明细输出目录
│       │   ├── latest.xls                # 最近一次导出的原始 Excel
│       │   ├── latest.csv                # 最近一次转换后的 CSV
│       │   ├── latest.json               # 最近一次转换后的 JSON，含元信息与 records
│       │   ├── latest.jsonl              # 最近一次转换后的 JSONL
│       │   ├── raw/                      # 按时间戳归档的历史 Excel
│       │   └── debug/                    # 页面 HTML / 截图调试文件
│       ├── warehouse_detail/             # 仓库明细输出目录
│       ├── warehouse_ledger/             # 仓库台账输出目录
│       ├── sales_outbound_detail/        # 销售出库明细输出目录
│       ├── defective_tracking/           # 不良品跟踪输出目录
│       ├── post_process_report_tracking/ # 后工序报工跟踪输出目录
│       ├── production_takt_maintenance/  # 生产节拍维护输出目录
│       ├── production_equipment_maintenance/ # 生产设备维护输出目录
│       ├── equipment_downtime_query/     # 设备停机查询输出目录
│       ├── tooling_replacement_query/    # 工装更换查询输出目录
│       ├── machining_transfer_query/     # 机加转序查询输出目录
│       ├── project_info_maintenance/     # 项目信息维护输出目录
│       ├── product_material_match_maintenance/ # 产品材料匹配维护输出目录
│       └── weekly_plan_execution_tracking/ # 周计划执行跟踪输出目录
├── scripts/
│   └── scrape_purchase_inbound.py        # 当前主抓取脚本
├── requirements.txt                      # Python 依赖
└── README.md                             # 项目说明
```

## 4. 核心脚本说明

主脚本：

```text
scripts/scrape_purchase_inbound.py
```

虽然文件名仍叫 `scrape_purchase_inbound.py`，但当前已经不是只抓采购入库，而是统一抓取多个 CMMS 业务页面。

### 4.1 环境变量读取

脚本启动后，从环境变量读取登录地址、各业务页面地址、账号、密码和业务时间范围：

```python
CMMS_LOGIN_URL
CMMS_PURCHASE_INBOUND_URL
CMMS_WAREHOUSE_DETAIL_URL
CMMS_WAREHOUSE_LEDGER_URL
CMMS_SALES_OUTBOUND_DETAIL_URL
CMMS_DEFECTIVE_TRACKING_URL
CMMS_POST_PROCESS_REPORT_TRACKING_URL
CMMS_PRODUCTION_TAKT_MAINTENANCE_URL
CMMS_PRODUCTION_EQUIPMENT_MAINTENANCE_URL
CMMS_EQUIPMENT_DOWNTIME_QUERY_URL
CMMS_TOOLING_REPLACEMENT_QUERY_URL
CMMS_MACHINING_TRANSFER_QUERY_URL
CMMS_PROJECT_INFO_MAINTENANCE_URL
CMMS_PRODUCT_MATERIAL_MATCH_MAINTENANCE_URL
CMMS_WEEKLY_PLAN_EXECUTION_TRACKING_URL
CMMS_USERNAME
CMMS_PASSWORD
DATE_RANGE
```

这些敏感信息不应写入代码，应放在 GitHub Secrets 中。

### 4.2 输出目录初始化

脚本会为每个数据源自动创建输出目录：

```text
data/excel_export/purchase_inbound/
data/excel_export/warehouse_detail/
data/excel_export/warehouse_ledger/
data/excel_export/sales_outbound_detail/
data/excel_export/defective_tracking/
data/excel_export/post_process_report_tracking/
data/excel_export/production_takt_maintenance/
data/excel_export/production_equipment_maintenance/
data/excel_export/equipment_downtime_query/
data/excel_export/tooling_replacement_query/
data/excel_export/machining_transfer_query/
data/excel_export/project_info_maintenance/
data/excel_export/product_material_match_maintenance/
data/excel_export/weekly_plan_execution_tracking/
```

每个目录下会自动创建：

```text
raw/      # 原始 Excel 历史归档
debug/    # 页面 HTML 和截图，用于排查抓取失败
```

### 4.3 默认日期范围

如果没有传入 `DATE_RANGE`，脚本默认抓取最近 3 天：

```text
今天 - 3 天 / 今天
```

格式为：

```text
YYYY-MM-DD/YYYY-MM-DD
```

例如：

```text
2026-05-21/2026-05-24
```

### 4.4 登录系统

脚本使用 Playwright 启动 Chromium 浏览器：

```text
headless=True
accept_downloads=True
viewport=1440x1000
ignore_https_errors=True
```

执行步骤：

```text
打开登录页
→ 填写用户名
→ 填写密码
→ 尝试点击登录按钮
→ 等待页面加载
```

登录按钮会按以下选择器依次尝试：

```text
input[type="submit"]
button
text=登录
text=登 录
text=Login
```

### 4.5 地址修正

脚本包含 `normalize_url_keep_login_host()` 方法。

作用是：

```text
当业务页面 URL 与登录页 hostname 相同，但端口或 netloc 不一致时，自动使用登录页的 netloc。
```

这是为了避免系统地址中 IP、端口、Host 不一致导致登录态失效。

### 4.6 页面查询与导出

统一抓取方法：

```python
scrape_export(page, name, url, output_dir, payload, date_selector=None, date_value=None)
```

每个业务页面的执行逻辑：

```text
打开业务页面
→ 保存页面调试信息
→ 在所有 iframe / frame 中寻找日期输入框或查询按钮
→ 如有日期输入框，则填入 DATE_RANGE
→ 点击 #Button2 开始查询
→ 等待查询结果
→ 再次保存调试信息
→ 点击“导出Excel”
→ 保存 latest.xls
→ 复制一份到 raw/历史归档
→ 解析 Excel 并生成 CSV / JSON / JSONL
```

当前各页面的日期控件配置：

| 数据源 | 日期选择器 |
|---|---|
| 采购入库明细 | `#text3` |
| 仓库明细 | `#text0` |
| 仓库台账 | 不设置日期 |
| 销售出库明细 | `#text0` |
| 不良品跟踪 | 不设置日期 |
| 后工序报工跟踪 | `#text0` |
| 生产节拍维护 | 不设置日期 |
| 生产设备维护 | 不设置日期 |
| 设备停机查询 | `#text0` |
| 工装更换查询 | `#text0` |
| 机加转序查询 | `#text0` |
| 项目信息维护 | 不设置日期 |
| 产品材料匹配维护 | 不设置日期 |
| 周计划执行跟踪 | 不设置日期 |

查询按钮统一使用：

```text
#Button2
```

导出按钮文本统一匹配：

```text
导出Excel
```

### 4.7 iframe / frame 处理

系统页面可能把真实表单放在 iframe 或 frame 中。

脚本通过 `find_frame_with_selector()` 遍历所有 frame，只要某个 frame 中存在目标选择器，就在该 frame 内执行填表、查询、导出操作。

这比直接在主页面查找元素更稳。

### 4.8 调试文件保存

每次关键页面都会保存：

```text
*.html
*.png
```

位置：

```text
data/excel_export/<source>/debug/
```

用途：

- 确认是否登录成功；
- 确认是否进入正确页面；
- 确认按钮、输入框、iframe 是否变化；
- 抓取失败时给 Codex / AI 分析页面结构。

### 4.9 Excel 解析与标准化

下载 Excel 后，脚本先尝试：

```python
pd.read_excel(xls_path)
```

如果失败，则回退到：

```python
pd.read_html(xls_path)
```

这是因为部分老系统导出的 `.xls` 实际上可能是 HTML 表格伪装成 Excel。

解析后会执行基础清洗：

```text
删除全空行
删除全空列
字段名去除前后空格
```

### 4.10 输出格式

每个数据源都会输出 4 类文件：

```text
latest.xls      # 最近一次原始导出文件
latest.csv      # 表格化结果，方便 Excel / pandas 读取
latest.json     # 带 source、date_range、synced_at、count、records 的完整 JSON
latest.jsonl    # 一行一条记录，适合后续增量处理或导入数据库
```

同时会在 `raw/` 下保存按时间戳命名的历史原始 Excel：

```text
raw/20260524_214728.xls
```

## 5. GitHub Actions 工作流

工作流文件：

```text
.github/workflows/scrape-purchase-inbound.yml
```

工作流名称：

```text
Scrape Purchase Inbound
```

名称还保留早期叫法，但当前实际已经抓取多个数据源。

### 5.1 触发方式

支持手动运行：

```text
Actions → Scrape Purchase Inbound → Run workflow
```

手动运行时可填写：

```text
date_range = YYYY-MM-DD/YYYY-MM-DD
```

也支持定时运行：

```yaml
schedule:
  - cron: "0 */2 * * *"
```

即每 2 小时运行一次。

### 5.2 Actions 执行步骤

```text
Checkout repository
→ Setup Python 3.12
→ pip install -r requirements.txt
→ python -m playwright install chromium
→ python scripts/scrape_purchase_inbound.py
→ git add data/
→ 如 data/ 有变化，则自动 commit + push
```

### 5.3 自动提交说明

运行成功后，工作流会检查 `data/` 目录是否有变化：

```text
有变化：提交 Update purchase inbound data 并 push
无变化：输出 No data changes.
```

## 6. GitHub Secrets 配置

进入仓库：

```text
Settings → Secrets and variables → Actions → New repository secret
```

需要配置：

| Secret 名称 | 用途 |
|---|---|
| `CMMS_LOGIN_URL` | CMMS 登录页地址 |
| `CMMS_PURCHASE_INBOUND_URL` | 采购入库明细页面地址 |
| `CMMS_WAREHOUSE_DETAIL_URL` | 仓库明细页面地址 |
| `CMMS_WAREHOUSE_LEDGER_URL` | 仓库台账页面地址 |
| `CMMS_SALES_OUTBOUND_DETAIL_URL` | 销售出库明细页面地址 |
| `CMMS_DEFECTIVE_TRACKING_URL` | 不良品跟踪页面地址 |
| `CMMS_POST_PROCESS_REPORT_TRACKING_URL` | 后工序报工跟踪页面地址 |
| `CMMS_PRODUCTION_TAKT_MAINTENANCE_URL` | 生产节拍维护页面地址 |
| `CMMS_PRODUCTION_EQUIPMENT_MAINTENANCE_URL` | 生产设备维护页面地址 |
| `CMMS_EQUIPMENT_DOWNTIME_QUERY_URL` | 设备停机查询页面地址 |
| `CMMS_TOOLING_REPLACEMENT_QUERY_URL` | 工装更换查询页面地址 |
| `CMMS_MACHINING_TRANSFER_QUERY_URL` | 机加转序查询页面地址 |
| `CMMS_PROJECT_INFO_MAINTENANCE_URL` | 项目信息维护页面地址 |
| `CMMS_PRODUCT_MATERIAL_MATCH_MAINTENANCE_URL` | 产品材料匹配维护页面地址 |
| `CMMS_WEEKLY_PLAN_EXECUTION_TRACKING_URL` | 周计划执行跟踪页面地址 |
| `CMMS_USERNAME` | 登录账号 |
| `CMMS_PASSWORD` | 登录密码 |

不要把真实账号、密码、内部系统地址直接写进代码或 README。

## 7. 本地运行方法

### 7.1 安装依赖

```bash
pip install -r requirements.txt
python -m playwright install chromium
```

当前依赖：

```text
playwright
pandas
xlrd
openpyxl
lxml
```

### 7.2 设置环境变量

macOS / Linux：

```bash
export CMMS_LOGIN_URL="登录页地址"
export CMMS_PURCHASE_INBOUND_URL="采购入库明细地址"
export CMMS_WAREHOUSE_DETAIL_URL="仓库明细地址"
export CMMS_WAREHOUSE_LEDGER_URL="仓库台账地址"
export CMMS_SALES_OUTBOUND_DETAIL_URL="销售出库明细地址"
export CMMS_DEFECTIVE_TRACKING_URL="不良品跟踪地址"
export CMMS_POST_PROCESS_REPORT_TRACKING_URL="后工序报工跟踪地址"
export CMMS_PRODUCTION_TAKT_MAINTENANCE_URL="生产节拍维护地址"
export CMMS_PRODUCTION_EQUIPMENT_MAINTENANCE_URL="生产设备维护地址"
export CMMS_EQUIPMENT_DOWNTIME_QUERY_URL="设备停机查询地址"
export CMMS_TOOLING_REPLACEMENT_QUERY_URL="工装更换查询地址"
export CMMS_MACHINING_TRANSFER_QUERY_URL="机加转序查询地址"
export CMMS_PROJECT_INFO_MAINTENANCE_URL="项目信息维护地址"
export CMMS_PRODUCT_MATERIAL_MATCH_MAINTENANCE_URL="产品材料匹配维护地址"
export CMMS_WEEKLY_PLAN_EXECUTION_TRACKING_URL="周计划执行跟踪地址"
export CMMS_USERNAME="账号"
export CMMS_PASSWORD="密码"
export DATE_RANGE="2026-05-21/2026-05-24"
```

Windows PowerShell：

```powershell
$env:CMMS_LOGIN_URL="登录页地址"
$env:CMMS_PURCHASE_INBOUND_URL="采购入库明细地址"
$env:CMMS_WAREHOUSE_DETAIL_URL="仓库明细地址"
$env:CMMS_WAREHOUSE_LEDGER_URL="仓库台账地址"
$env:CMMS_SALES_OUTBOUND_DETAIL_URL="销售出库明细地址"
$env:CMMS_DEFECTIVE_TRACKING_URL="不良品跟踪地址"
$env:CMMS_POST_PROCESS_REPORT_TRACKING_URL="后工序报工跟踪地址"
$env:CMMS_PRODUCTION_TAKT_MAINTENANCE_URL="生产节拍维护地址"
$env:CMMS_PRODUCTION_EQUIPMENT_MAINTENANCE_URL="生产设备维护地址"
$env:CMMS_EQUIPMENT_DOWNTIME_QUERY_URL="设备停机查询地址"
$env:CMMS_TOOLING_REPLACEMENT_QUERY_URL="工装更换查询地址"
$env:CMMS_MACHINING_TRANSFER_QUERY_URL="机加转序查询地址"
$env:CMMS_PROJECT_INFO_MAINTENANCE_URL="项目信息维护地址"
$env:CMMS_PRODUCT_MATERIAL_MATCH_MAINTENANCE_URL="产品材料匹配维护地址"
$env:CMMS_WEEKLY_PLAN_EXECUTION_TRACKING_URL="周计划执行跟踪地址"
$env:CMMS_USERNAME="账号"
$env:CMMS_PASSWORD="密码"
$env:DATE_RANGE="2026-05-21/2026-05-24"
```

### 7.3 运行脚本

```bash
python scripts/scrape_purchase_inbound.py
```

## 8. 当前能力边界

当前项目能做：

- 自动登录网页系统；
- 自动打开多个业务页面；
- 自动填写日期范围；
- 自动点击查询；
- 自动导出 Excel；
- 自动转换 CSV / JSON / JSONL；
- 自动保存调试页面和截图；
- 自动通过 GitHub Actions 定时执行；
- 自动把抓取结果提交回仓库。

当前项目暂不处理：

- 验证码登录；
- 多因素认证；
- 页面结构频繁变更后的自动适配；
- 复杂字段语义清洗；
- 多表关联分析；
- 增量去重入库；
- 可视化看板；
- 数据库持久化。

## 9. 后续优化方向

建议按以下顺序推进：

1. 将脚本文件名从 `scrape_purchase_inbound.py` 改为更准确的 `scrape_cmms_exports.py`；
2. 将数据源配置抽成配置表，减少重复代码；
3. 为每个数据源建立字段映射表；
4. 增加数据校验：空表、字段缺失、数量异常、日期异常；
5. 将 JSONL 增量写入 SQLite / PostgreSQL / Baserow；
6. 建立项目主控表：入库、库存、出库、发货、异常联动；
7. 对接企业微信 / 邮件，推送缺料、延期、库存异常；
8. 将 `latest` 与历史归档分离，避免仓库数据持续膨胀。

## 10. 安全原则

- 账号密码只放 GitHub Secrets；
- 内部系统真实地址不要写入公开文档；
- 如仓库是公开仓库，谨慎提交导出的业务数据；
- debug HTML / 截图可能包含敏感业务信息，必要时应加入 `.gitignore` 或改为 artifact 保存；
- 自动化只读取和导出数据，不应修改原系统业务数据。
