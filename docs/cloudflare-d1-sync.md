# Cloudflare D1 同步说明

本项目的 GitHub Actions 会在抓取完成后运行：

```bash
python scripts/clean_database_inputs.py
```

清洗结果输出到：

```text
data/normalized/
```

其中最关键的两个文件是：

```text
data/normalized/schema.sql   # 建表和索引
data/normalized/upsert.sql   # 当前清洗数据的幂等写入
```

`upsert.sql` 是 Actions 运行时产物，内容会重复包含一份清洗数据，因此默认不提交到 Git；它会随 workflow artifact 上传，并在 D1 写入步骤中直接使用。

默认定时任务使用 `UPSERT_MODE=incremental_window`：只把带日期查询条件的滚动窗口表写入 D1，包括采购入库、仓库明细、销售出库、后工序报工、设备停机、工装更换和机加转序。没有日期条件的快照/维表不在默认增量 upsert 中重复写入。

手动运行 workflow 时可以把 `d1_upload_scope` 切到 `all_latest`，用于初始化或需要刷新快照/维表的场景。

## GitHub 侧配置

在 GitHub 仓库中配置以下 Secrets：

| 名称 | 用途 |
| --- | --- |
| `CLOUDFLARE_API_TOKEN` | Wrangler 访问 Cloudflare API |
| `CLOUDFLARE_ACCOUNT_ID` | Cloudflare Account ID |
| `CLOUDFLARE_D1_DATABASE_NAME` | D1 数据库名称，也可以放在 Variables |

建议再配置一个仓库 Variable：

| 名称 | 建议值 | 用途 |
| --- | --- | --- |
| `UPLOAD_TO_D1` | `false` | 是否让定时任务自动写入 D1 |

如果只想手动测试，可以不启用 `UPLOAD_TO_D1`，在手动运行 workflow 时把 `upload_to_d1` 选为 `true`。

## GitHub Actions 写入 D1 的命令

workflow 使用 Wrangler 执行官方支持的 D1 SQL 文件导入方式：

```bash
npx wrangler d1 execute "$D1_DATABASE_NAME" --remote --yes --file=data/normalized/schema.sql
npx wrangler d1 execute "$D1_DATABASE_NAME" --remote --yes --file=data/normalized/upsert.sql
```

`schema.sql` 使用 `CREATE TABLE IF NOT EXISTS`，可以重复执行。

`upsert.sql` 使用：

```sql
INSERT INTO table_name (...) VALUES (...)
ON CONFLICT(record_key) DO UPDATE SET ...
```

因此重复同步同一批数据不会插入重复行。若源系统回写或修正了同一业务记录，`record_hash` 会随内容变化而更新。

注意：`upsert.sql` 不包含显式事务包装。Wrangler 远端执行 D1 SQL 文件时不接受显式事务语句，失败时 D1 会按导入机制回滚，可直接重试。

## 推荐启用顺序

1. 先手动运行 workflow，`upload_to_d1=false`，确认抓取和清洗产物正常。
2. 在 Cloudflare 创建 D1 数据库，并把数据库名称写入 `CLOUDFLARE_D1_DATABASE_NAME`。
3. 配置 `CLOUDFLARE_API_TOKEN` 和 `CLOUDFLARE_ACCOUNT_ID`。
4. 首次全量同步时，手动运行 workflow，`upload_to_d1=true` 且 `d1_upload_scope=all_latest`。
5. 验证无误后，把仓库 Variable `UPLOAD_TO_D1` 改为 `true`，让定时任务自动同步。
6. 定时任务默认使用 `d1_upload_scope=incremental_window`，每次只写入最近几天窗口数据，并由 D1 根据 `record_key` 去重/upsert。

## 验证 D1

可在本地或 Actions 中执行：

```bash
npx wrangler d1 execute "$D1_DATABASE_NAME" --remote --command="SELECT COUNT(*) AS c FROM purchase_inbound"
```

也可以查询清洗元数据：

```bash
npx wrangler d1 execute "$D1_DATABASE_NAME" --remote --command="SELECT source_table, COUNT(*) AS c FROM purchase_inbound GROUP BY source_table"
```
