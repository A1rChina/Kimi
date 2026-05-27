from __future__ import annotations

import hashlib
import json
import os
import re
import warnings
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DATA_ROOT = ROOT / "data" / "excel_export"
NORMALIZED_ROOT = Path(os.environ.get("NORMALIZED_ROOT", ROOT / "data" / "normalized"))
REPORT_ROOT = ROOT / "reports"

READ_ENCODINGS = ("utf-8-sig", "utf-8", "gb18030", "gbk")

SOURCE_TITLES = {
    "purchase_inbound": "采购入库明细",
    "warehouse_detail": "仓库明细",
    "warehouse_ledger": "仓库台账",
    "sales_outbound_detail": "销售出库明细",
    "defective_tracking": "不良品跟踪",
    "post_process_report_tracking": "后工序报工跟踪",
    "production_takt_maintenance": "生产节拍维护",
    "production_equipment_maintenance": "生产设备维护",
    "equipment_downtime_query": "设备停机查询",
    "tooling_replacement_query": "工装更换查询",
    "machining_transfer_query": "机加转序查询",
    "project_info_maintenance": "项目信息维护",
    "product_material_match_maintenance": "产品材料匹配维护",
    "weekly_plan_execution_tracking": "周计划执行跟踪",
}

COLUMN_MAP = {
    "质量跟踪码": "quality_tracking_code",
    "公司名称": "company_name",
    "公司ID": "company_id",
    "物料编码": "material_code",
    "物料名称": "material_name",
    "物资编码": "material_code",
    "物资名称": "material_name",
    "物资ID": "material_id",
    "单位": "unit",
    "计量单位": "unit",
    "申报时间": "reported_at",
    "申报人员": "reporter",
    "申报人": "reporter",
    "部门名称": "department_name",
    "部门": "department_name",
    "反馈来源": "feedback_source",
    "缺陷描述": "defect_description",
    "申报说明": "report_description",
    "后序处置": "followup_disposition",
    "供应商": "supplier",
    "入库性质": "inbound_nature",
    "总量": "total_qty",
    "待入库量": "pending_inbound_qty",
    "入库数量": "inbound_qty",
    "出库量": "outbound_qty",
    "在库量": "on_hand_qty",
    "入出库明细": "inventory_flow_detail",
    "异常单号": "exception_no",
    "设备号": "equipment_id",
    "设备ID": "equipment_id",
    "设备名称": "equipment_name",
    "加工设备": "equipment_name",
    "设备型号": "equipment_model",
    "报停时间": "downtime_started_at",
    "重启时间": "downtime_restarted_at",
    "班次": "shift",
    "停机原因": "downtime_reason",
    "夹具码": "fixture_code",
    "业务状态": "business_status",
    "原因描述": "reason_description",
    "重启人": "restart_user",
    "报停时长": "downtime_hours",
    "审核说明": "audit_note",
    "审核时长": "audit_hours",
    "审核人姓名": "auditor_name",
    "审核人": "auditor",
    "审核时间": "audited_at",
    "制程码": "process_code",
    "工序": "operation",
    "机台": "machine_name",
    "生产车间": "production_workshop",
    "规格型号": "spec_model",
    "规格图号": "drawing_no",
    "合格品": "qualified_qty",
    "合格数": "qualified_qty",
    "不良品": "defective_qty",
    "报工人员": "work_reporter",
    "报工时间": "work_reported_at",
    "标签来源": "label_source",
    "产品/项目": "project_code",
    "项目/车型": "project_code",
    "产品编码": "product_code",
    "产品名称": "product_name",
    "物料ID": "material_id",
    "上线人": "start_user",
    "操作时间": "operated_at",
    "上线数": "started_qty",
    "在线时间": "online_hours",
    "领料物资ID": "issue_material_id",
    "材料规格": "material_spec",
    "领料物资名称": "issue_material_name",
    "单台用量": "usage_per_unit",
    "毛坯重【KG】": "blank_weight_kg",
    "产品重【KG】": "product_weight_kg",
    "铝块重【KG】": "aluminum_block_weight_kg",
    "铝屑重【KG】": "aluminum_chip_weight_kg",
    "领料单位": "issue_department",
    "默认客户": "default_customer",
    "排程分类": "scheduling_category",
    "资产编码": "asset_code",
    "当前用户": "current_user",
    "当前时间": "current_at",
    "记录行号": "source_row_no",
    "客户ID": "customer_id",
    "总节拍【分】": "total_takt_minutes",
    "生产时长【分】": "production_minutes",
    "辅助时长【分】": "auxiliary_minutes",
    "维护人员": "maintainer",
    "维护时间": "maintained_at",
    "历史维护信息": "maintenance_history",
    "图片附件": "image_attachments",
    "客户名称": "customer_name",
    "客户全称": "customer_full_name",
    "核算期": "accounting_period",
    "业务人员": "business_user",
    "业务时间": "business_at",
    "业务单号": "business_doc_no",
    "业务类型": "business_type",
    "库房": "warehouse",
    "计划价": "planned_price",
    "数量": "quantity",
    "二级": "secondary_supplier",
    "毛坯重": "blank_weight",
    "报价重": "quoted_weight",
    "操作人员": "operator",
    "调入库房编码": "to_warehouse_code",
    "调入库房": "to_warehouse",
    "客户": "customer",
    "捆包号": "bundle_no",
    "捆包识别码": "bundle_identifier",
    "上传人员": "uploader",
    "上传时间": "uploaded_at",
    "集成标志": "integration_flag",
    "调出货主编码": "from_owner_code",
    "调出货主": "from_owner",
    "调入货主编码": "to_owner_code",
    "调入货主": "to_owner",
    "调入组织编码": "to_org_code",
    "调入库存组织": "to_inventory_org",
    "调出组织编码": "from_org_code",
    "调出库存组织": "from_inventory_org",
    "单据类型编码": "document_type_code",
    "单据类型名称": "document_type_name",
    "结算组织编码": "settlement_org_code",
    "结算组织": "settlement_org",
    "销售组织编码": "sales_org_code",
    "销售组织": "sales_org",
    "数据来源": "data_source",
    "数量1": "secondary_quantity",
    "车牌号": "vehicle_plate",
    "器具描述": "container_description",
    "单据号": "document_no",
    "入出库编码": "inventory_flow_code",
    "入出库类型": "inventory_flow_type",
    "入出库标识": "inventory_flow_direction",
    "用途": "usage",
    "合作伙伴": "partner",
    "单据描述": "document_description",
    "行描述": "line_description",
    "要货日期": "requested_date",
    "月度": "accounting_month",
    "库房组": "warehouse_group",
    "期初库存": "opening_inventory_qty",
    "入库": "inbound_qty",
    "出库": "outbound_qty",
    "期末库存": "closing_inventory_qty",
    "当前库存": "current_inventory_qty",
    "虚拟库存": "virtual_inventory_qty",
    "在途": "in_transit_qty",
    "起始时间": "started_at",
    "截止时间": "ended_at",
    "更换项目": "replacement_project",
    "零件": "part_name",
    "用时H": "duration_hours",
    "调试人": "debugger",
    "试料数(报废)": "trial_scrap_qty",
    "检验人": "inspector",
    "夹具维护": "fixture_maintenance",
    "备注": "remark",
    "年周": "year_week",
    "开始日期": "start_date",
    "单车用量": "usage_per_vehicle",
    "星期日": "sunday_qty",
    "星期一": "monday_qty",
    "星期二": "tuesday_qty",
    "星期三": "wednesday_qty",
    "星期四": "thursday_qty",
    "星期五": "friday_qty",
    "星期六": "saturday_qty",
    "计划量": "planned_qty",
    "完成量": "completed_qty",
    "完成率": "completion_rate",
    "属性": "attribute",
}

TYPE_OVERRIDES = {
    "accounting_period": "text",
    "accounting_month": "text",
    "year_week": "text",
    "source_row_no": "integer",
    "quality_tracking_code": "text",
    "start_date": "date",
    "requested_date": "date",
    "inventory_flow_detail": "text",
    "inventory_flow_code": "text",
    "inventory_flow_type": "text",
    "inventory_flow_direction": "text",
    "to_inventory_org": "text",
    "from_inventory_org": "text",
}

KEY_FIELDS = {
    "defective_tracking": ["quality_tracking_code"],
    "equipment_downtime_query": ["exception_no"],
    "machining_transfer_query": ["process_code"],
    "post_process_report_tracking": [
        "process_code",
        "operation",
        "operated_at",
        "start_user",
        "started_qty",
    ],
    "product_material_match_maintenance": ["product_code"],
    "production_equipment_maintenance": ["equipment_id"],
    "production_takt_maintenance": ["source_row_no"],
    "project_info_maintenance": ["project_code"],
    "purchase_inbound": ["business_doc_no", "material_id", "quantity"],
    "sales_outbound_detail": ["business_doc_no", "bundle_no", "material_code", "quantity"],
    "tooling_replacement_query": [
        "started_at",
        "ended_at",
        "equipment_id",
        "replacement_project",
        "part_name",
    ],
    "warehouse_detail": ["document_no", "material_code", "business_at", "quantity"],
    "warehouse_ledger": ["accounting_month", "warehouse", "material_code"],
    "weekly_plan_execution_tracking": ["year_week", "project_code", "product_code", "attribute"],
}

DATE_WINDOW_SOURCES = {
    "purchase_inbound",
    "warehouse_detail",
    "sales_outbound_detail",
    "post_process_report_tracking",
    "equipment_downtime_query",
    "tooling_replacement_query",
    "machining_transfer_query",
}

UPSERT_MODE = os.environ.get("UPSERT_MODE", "all_latest").strip().lower()

NUMERIC_KEYWORDS = (
    "qty",
    "quantity",
    "price",
    "weight",
    "hours",
    "minutes",
    "rate",
    "usage_per",
    "inventory",
)
DATETIME_SUFFIXES = ("_at",)
DATE_SUFFIXES = ("_date",)


def read_latest_csv(path: Path) -> tuple[pd.DataFrame, str]:
    last_error: Exception | None = None
    for encoding in READ_ENCODINGS:
        try:
            df = pd.read_csv(path, dtype=str, encoding=encoding, keep_default_na=False)
            return df.fillna(""), encoding
        except Exception as exc:  # noqa: BLE001 - keep trying declared encodings.
            last_error = exc
    raise RuntimeError(f"Unable to read {path}: {last_error}")


def snake_case(value: str) -> str:
    value = re.sub(r"[^0-9A-Za-z]+", "_", value).strip("_").lower()
    return value or "column"


def dedupe_names(names: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    result = []
    for name in names:
        count = seen.get(name, 0) + 1
        seen[name] = count
        result.append(name if count == 1 else f"{name}_{count}")
    return result


def normalize_text(value: Any) -> str | None:
    text = "" if value is None else str(value).strip()
    if text.lower() in {"", "nan", "none", "null"}:
        return None
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_number(value: Any) -> float | None:
    text = normalize_text(value)
    if text is None:
        return None
    text = text.replace(",", "").replace("%", "")
    try:
        return float(text)
    except ValueError:
        return None


def normalize_integer(value: Any) -> int | None:
    number = normalize_number(value)
    if number is None:
        return None
    return int(number)


def normalize_datetime(value: Any) -> str | None:
    text = normalize_text(value)
    if text is None:
        return None
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.strftime("%Y-%m-%d %H:%M:%S")


def normalize_date(value: Any) -> str | None:
    text = normalize_text(value)
    if text is None:
        return None
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.strftime("%Y-%m-%d")


def infer_type(column_name: str, series: pd.Series) -> str:
    if column_name in TYPE_OVERRIDES:
        return TYPE_OVERRIDES[column_name]
    if column_name.endswith(DATETIME_SUFFIXES):
        return "datetime"
    if column_name.endswith(DATE_SUFFIXES):
        return "date"
    if any(keyword in column_name for keyword in NUMERIC_KEYWORDS):
        return "real"

    nonempty = series.map(normalize_text).dropna()
    if nonempty.empty:
        return "text"
    numeric = nonempty.map(normalize_number).notna().mean()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        parsed_dates = pd.to_datetime(nonempty, errors="coerce").notna().mean()
    if numeric >= 0.90:
        return "real"
    if parsed_dates >= 0.90:
        return "datetime"
    return "text"


def sqlite_type(kind: str) -> str:
    return {
        "integer": "INTEGER",
        "real": "REAL",
        "date": "TEXT",
        "datetime": "TEXT",
        "text": "TEXT",
    }[kind]


def normalize_series(series: pd.Series, kind: str) -> pd.Series:
    if kind == "integer":
        return series.map(normalize_integer).astype("Int64")
    if kind == "real":
        return series.map(normalize_number)
    if kind == "date":
        return series.map(normalize_date)
    if kind == "datetime":
        return series.map(normalize_datetime)
    return series.map(normalize_text)


def record_key(row: pd.Series, key_fields: list[str]) -> str:
    parts = []
    for field in key_fields:
        value = row.get(field)
        parts.append("" if pd.isna(value) else str(value))
    text = "|".join(parts)
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def row_hash(row: pd.Series, fields: list[str]) -> str:
    text = json.dumps(
        {field: (None if pd.isna(row.get(field)) else row.get(field)) for field in fields},
        ensure_ascii=False,
        sort_keys=True,
        default=str,
    )
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def build_table(source: str, path: Path) -> tuple[pd.DataFrame, dict[str, Any]]:
    raw_df, encoding = read_latest_csv(path)
    raw_columns = list(raw_df.columns)
    mapped_columns = dedupe_names([COLUMN_MAP.get(column, snake_case(column)) for column in raw_columns])
    df = raw_df.copy()
    df.columns = mapped_columns

    column_profiles = []
    normalized = pd.DataFrame()
    for raw_name, clean_name in zip(raw_columns, mapped_columns):
        kind = infer_type(clean_name, df[clean_name])
        before_nonempty = df[clean_name].map(normalize_text).notna()
        normalized[clean_name] = normalize_series(df[clean_name], kind)
        after_valid = normalized[clean_name].notna()
        parse_failure_rate = 0.0
        if kind in {"integer", "real", "date", "datetime"} and before_nonempty.any():
            parse_failure_rate = float((before_nonempty & ~after_valid).sum() / before_nonempty.sum())
        column_profiles.append(
            {
                "raw_name": raw_name,
                "name": clean_name,
                "type": kind,
                "sqlite_type": sqlite_type(kind),
                "null_rate": round(float(normalized[clean_name].isna().mean()), 4)
                if len(normalized)
                else 0.0,
                "unique_count": int(normalized[clean_name].nunique(dropna=True)),
                "parse_failure_rate": round(parse_failure_rate, 4),
            }
        )

    data_fields = list(normalized.columns)
    normalized.insert(0, "source_table", source)
    normalized.insert(1, "source_file", str(path.as_posix()))
    normalized.insert(2, "source_row_index", range(1, len(normalized) + 1))
    normalized.insert(3, "extracted_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    configured_keys = [field for field in KEY_FIELDS.get(source, []) if field in normalized.columns]
    if not configured_keys or normalized[configured_keys].isna().any(axis=None):
        configured_keys = data_fields
    keys_used = configured_keys
    key_values = normalized.apply(lambda row: record_key(row, keys_used), axis=1)
    if key_values.duplicated().any():
        keys_used = data_fields
        key_values = normalized.apply(lambda row: record_key(row, keys_used), axis=1)
    if key_values.duplicated().any():
        keys_used = data_fields + ["source_row_index"]
        key_values = normalized.apply(lambda row: record_key(row, keys_used), axis=1)
    normalized.insert(0, "record_key", key_values)
    normalized.insert(
        1,
        "record_hash",
        normalized.apply(lambda row: row_hash(row, data_fields), axis=1),
    )

    duplicate_record_keys = int(normalized["record_key"].duplicated().sum())
    duplicate_rows = int(normalized[data_fields].duplicated().sum()) if data_fields else 0
    profile = {
        "source": source,
        "title": SOURCE_TITLES.get(source, source),
        "input_file": str(path.as_posix()),
        "input_encoding": encoding,
        "output_file": str((NORMALIZED_ROOT / f"{source}.csv").as_posix()),
        "rows": int(len(normalized)),
        "input_columns": int(len(raw_columns)),
        "output_columns": int(len(normalized.columns)),
        "data_columns": int(len(data_fields)),
        "key_fields": keys_used,
        "duplicate_record_keys": duplicate_record_keys,
        "duplicate_data_rows": duplicate_rows,
        "columns": column_profiles,
        "index_fields": [field for field in KEY_FIELDS.get(source, []) if field in normalized.columns],
    }
    return normalized, profile


def write_schema_sql(profiles: list[dict[str, Any]]) -> None:
    lines = [
        "-- SQLite / Cloudflare D1 compatible schema generated from data/normalized.",
        "-- record_key is deterministic for idempotent imports.",
        "",
    ]
    metadata_columns = [
        ("record_key", "TEXT PRIMARY KEY"),
        ("record_hash", "TEXT NOT NULL"),
        ("source_table", "TEXT NOT NULL"),
        ("source_file", "TEXT NOT NULL"),
        ("source_row_index", "INTEGER NOT NULL"),
        ("extracted_at", "TEXT NOT NULL"),
    ]
    for profile in profiles:
        lines.append(f"CREATE TABLE IF NOT EXISTS {profile['source']} (")
        column_lines = [f"  {name} {kind}" for name, kind in metadata_columns]
        for column in profile["columns"]:
            column_lines.append(f"  {column['name']} {column['sqlite_type']}")
        lines.append(",\n".join(column_lines))
        lines.append(");")
        for field in profile["index_fields"]:
            if field != "record_key":
                lines.append(
                    f"CREATE INDEX IF NOT EXISTS idx_{profile['source']}_{field} "
                    f"ON {profile['source']} ({field});"
                )
        lines.append("")
    (NORMALIZED_ROOT / "schema.sql").write_text("\n".join(lines), encoding="utf-8")


def sql_literal(value: Any) -> str:
    if value is None:
        return "NULL"
    try:
        if pd.isna(value):
            return "NULL"
    except TypeError:
        pass
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return str(value)
    text = str(value).replace("'", "''")
    return f"'{text}'"


def upsert_tables_for_mode(tables: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    if UPSERT_MODE in {"all", "all_latest", "full"}:
        return tables
    if UPSERT_MODE in {"incremental", "incremental_window", "date_window"}:
        return {name: df for name, df in tables.items() if name in DATE_WINDOW_SOURCES}
    raise ValueError(f"Unsupported UPSERT_MODE: {UPSERT_MODE}")


def write_upsert_sql(tables: dict[str, pd.DataFrame]) -> None:
    tables = upsert_tables_for_mode(tables)
    lines = [
        "-- SQLite / Cloudflare D1 compatible upsert data generated from data/normalized.",
        "-- Run schema.sql before this file.",
        "-- Remote D1 imports reject explicit transaction wrappers.",
        f"-- UPSERT_MODE={UPSERT_MODE}; tables={','.join(sorted(tables))}",
        "",
    ]
    for source, df in sorted(tables.items()):
        columns = list(df.columns)
        column_sql = ", ".join(columns)
        update_columns = [column for column in columns if column != "record_key"]
        update_sql = ", ".join(f"{column}=excluded.{column}" for column in update_columns)
        for _, row in df.iterrows():
            values_sql = ", ".join(sql_literal(row[column]) for column in columns)
            lines.append(
                f"INSERT INTO {source} ({column_sql}) VALUES ({values_sql}) "
                f"ON CONFLICT(record_key) DO UPDATE SET {update_sql};"
            )
        lines.append("")
    (NORMALIZED_ROOT / "upsert.sql").write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_markdown_report(profiles: list[dict[str, Any]]) -> None:
    REPORT_ROOT.mkdir(parents=True, exist_ok=True)
    total_rows = sum(profile["rows"] for profile in profiles)
    high_null_fields = []
    parse_issues = []
    duplicate_key_tables = []
    for profile in profiles:
        if profile["duplicate_record_keys"]:
            duplicate_key_tables.append(
                f"- {profile['title']} (`{profile['source']}`): {profile['duplicate_record_keys']} 个重复 record_key"
            )
        for column in profile["columns"]:
            if column["null_rate"] >= 0.5:
                high_null_fields.append(
                    f"- {profile['title']}.`{column['raw_name']}` -> `{column['name']}`: 空值率 {column['null_rate']:.1%}"
                )
            if column["parse_failure_rate"] > 0:
                parse_issues.append(
                    f"- {profile['title']}.`{column['raw_name']}` -> `{column['name']}`: 转换失败率 {column['parse_failure_rate']:.1%}"
                )

    lines = [
        "# 08 数据库入库清洗评估",
        "",
        "## 总体结论",
        "",
        f"- 已清洗 {len(profiles)} 张 `latest.csv`，共 {total_rows} 行。",
        "- 输出目录：`data/normalized/`。",
        "- 每张表都补充 `record_key`、`record_hash`、`source_table`、`source_file`、`source_row_index`、`extracted_at`。",
        "- 字段名已从中文表头映射为稳定的英文 snake_case，日期统一为 `YYYY-MM-DD` 或 `YYYY-MM-DD HH:MM:SS`，数值列统一为数据库可读数字。",
        "- 生成 `data/normalized/schema.sql`，可用于 SQLite / Cloudflare D1 建表。",
        f"- 生成 `data/normalized/upsert.sql`，可用于向 SQLite / Cloudflare D1 幂等写入当前清洗数据；当前 `UPSERT_MODE={UPSERT_MODE}`。",
        "",
        "## 表级评估",
        "",
        "| 表 | 行数 | 数据列 | 输入编码 | 主键/去重字段 | 重复 record_key | 输出 |",
        "| --- | ---: | ---: | --- | --- | ---: | --- |",
    ]
    for profile in profiles:
        if len(profile["key_fields"]) > 8:
            keys = "`all_data_columns` + `source_row_index`"
        else:
            keys = ", ".join(f"`{field}`" for field in profile["key_fields"])
        lines.append(
            f"| {profile['title']} (`{profile['source']}`) | {profile['rows']} | "
            f"{profile['data_columns']} | {profile['input_encoding']} | {keys} | "
            f"{profile['duplicate_record_keys']} | `{profile['output_file']}` |"
        )

    lines.extend(
        [
            "",
            "## 仍需关注的问题",
            "",
            "### 高空值字段",
            "",
        ]
    )
    lines.extend(high_null_fields or ["- 未发现空值率超过 50% 的字段。"])
    lines.extend(["", "### 类型转换问题", ""])
    lines.extend(parse_issues or ["- 未发现日期/数值字段转换失败。"])
    lines.extend(["", "### 去重键问题", ""])
    lines.extend(duplicate_key_tables or ["- 当前清洗结果中未发现重复 `record_key`。"])
    lines.extend(
        [
            "",
            "## 入库建议",
            "",
            "1. 先执行 `data/normalized/schema.sql` 建表。",
            "2. 再执行 `data/normalized/upsert.sql` 写入数据；如果 `record_hash` 变化，说明同一业务键对应的内容被系统回写或修正。",
            "3. 原始导出仍保留在 `data/excel_export`，清洗层只作为数据库导入输入，不反向覆盖原始数据。",
            "4. 周计划表原始编码为 GB18030，清洗层已统一输出 UTF-8，后续导入应只使用 `data/normalized/weekly_plan_execution_tracking.csv`。",
        ]
    )
    (REPORT_ROOT / "08_database_readiness.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    NORMALIZED_ROOT.mkdir(parents=True, exist_ok=True)
    profiles = []
    tables = {}
    for source_dir in sorted(DATA_ROOT.iterdir()):
        if not source_dir.is_dir():
            continue
        latest_csv = source_dir / "latest.csv"
        if not latest_csv.exists():
            continue
        df, profile = build_table(source_dir.name, latest_csv)
        output_file = NORMALIZED_ROOT / f"{source_dir.name}.csv"
        df.to_csv(output_file, index=False, encoding="utf-8", lineterminator="\n")
        tables[source_dir.name] = df
        profiles.append(profile)

    profiles.sort(key=lambda item: item["source"])
    (NORMALIZED_ROOT / "manifest.json").write_text(
        json.dumps({"generated_at": datetime.now().isoformat(timespec="seconds"), "tables": profiles}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    write_schema_sql(profiles)
    write_upsert_sql(tables)
    write_markdown_report(profiles)
    print(f"Cleaned {len(profiles)} tables into {NORMALIZED_ROOT}")


if __name__ == "__main__":
    main()
