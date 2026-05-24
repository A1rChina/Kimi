# 增量抓取策略

## 目标

旧 CMMS/ERP 系统存在行数限制：超过一定数量的数据不会在页面显示或下载。因此程序稳定后不能长期依赖全量抓取，必须改为增量抓取。

## 核心原则

先全量初始化，再增量维护。

```text
第一次运行：尽可能按时间窗口分段抓全量历史数据
后续运行：只抓上次水位线之后的新数据
```

## 两类采集源

### 1. HTML 表格类

无 Excel 导出按钮，只能解析页面表格。

增量依据优先级：

1. 业务时间
2. 业务单号
3. 行级唯一 ID，如 checkbox value
4. 多字段 hash

### 2. Excel 导出类

有 Excel 导出按钮，优先下载 Excel。

增量依据优先级：

1. 查询条件中的起止日期
2. Excel 内业务时间
3. 单据号 + 物料编码 + 数量 + 时间 hash

## 水位线文件

每个数据源维护一个 state 文件：

```text
data/state/<source_name>.json
```

示例：

```json
{
  "source": "inventory_flow",
  "last_success_at": "2026-05-24T19:30:00+08:00",
  "watermark_field": "业务时间",
  "watermark_value": "2026/5/22 14:38:45",
  "last_doc_no": "R260522005",
  "seen_keys_file": "data/state/inventory_flow_seen_keys.txt"
}
```

## 增量抓取窗口

为了避免漏数，不应从水位线精确时间开始，而应向前回退一段时间。

建议：

```text
每次抓取开始时间 = 上次水位线 - 2天
```

然后通过唯一键去重。

原因：

- 老系统可能补录历史单据
- 用户可能修改旧单据
- 业务时间和录入时间可能不一致
- 定时任务可能中断

## 去重键设计

如果系统有行级 ID，优先使用行级 ID。

例如 HTML 中 checkbox 的 value：

```html
<input value="5780738" />
```

可以作为 source_row_id。

如果没有行级 ID，则使用组合 hash：

```text
业务单号 + 业务时间 + 业务类型 + 物资ID + 数量 + 库房
```

生成：

```text
record_key = sha256(组合字段)
```

## 文件落地策略

不要每次覆盖全量历史文件。

建议采用：

```text
data/raw/<source>/<YYYY-MM-DD>.jsonl
data/latest/<source>.json
data/state/<source>.json
```

其中：

- raw：追加保存增量明细
- latest：保存最近一次抓取结果
- state：保存水位线

## 为什么用 jsonl

每行一条 JSON，适合追加写入和后续导入数据库。

示例：

```jsonl
{"record_key":"abc","业务单号":"R260522005","业务时间":"2026/5/22 14:38:45"}
{"record_key":"def","业务单号":"R260522006","业务时间":"2026/5/22 15:01:12"}
```

## V1 版本拆分

### V1.0

只抓第一页，验证登录和解析。

### V1.1

抓取结果同步仓库。

### V1.2

加入 state 文件和去重键。

### V1.3

支持按日期窗口抓取。

### V1.4

稳定后改为增量抓取。

### V1.5

支持 Excel 导出类数据源。

## 注意

即使进入增量模式，也建议保留手动全量重建能力。

```text
workflow_dispatch 参数：mode=full 或 mode=incremental
```
