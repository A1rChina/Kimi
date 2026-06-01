# Cloudflare D1 同步已停用

本项目当前不再把抓取后的数据同步到 Cloudflare D1。

GitHub Actions 仍会在抓取完成后运行：

```bash
python scripts/clean_database_inputs.py
```

清洗结果输出到：

```text
data/normalized/
```

主要产物包括：

```text
data/normalized/schema.sql   # SQLite 建表和索引
data/normalized/upsert.sql   # 当前清洗数据的幂等写入 SQL
```

这些文件会作为 workflow artifact 上传，供人工下载、检查或导入本地 SQLite 使用。workflow 中已经移除 Wrangler、Cloudflare Secrets、`UPLOAD_TO_D1`、`upload_to_d1` 和 `d1_upload_scope` 相关配置。

定时抓取频率当前为每 6 小时一次：

```yaml
schedule:
  - cron: "7 */6 * * *"
```
