from __future__ import annotations

import csv
import json
import math
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DATA_ROOT = ROOT / "data" / "excel_export"
REPORT_ROOT = ROOT / "reports"

SOURCE_LABELS = {
    "purchase_inbound": "采购入库明细",
    "warehouse_detail": "仓库明细",
    "warehouse_ledger": "仓库台账 / 库存台账",
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

BUSINESS_FIELDS = [
    "项目编号 / 项目名称",
    "产品编码 / 零件号 / 图号",
    "产品名称",
    "型材编码 / 物料编码",
    "客户",
    "供应商",
    "日期 / 业务日期",
    "数量",
    "入库数量",
    "出库数量",
    "库存数量",
    "报工数量",
    "转序数量",
    "不良数量",
    "报废数量",
    "返修数量",
    "工序",
    "设备",
    "工装",
    "停机时间",
    "节拍",
    "状态",
    "备注",
]

ROLE_KEYWORDS = {
    "date": ["日期", "时间", "date", "time", "业务时间", "创建时间", "操作时间", "出库时间", "入库时间"],
    "quantity": ["数量", "qty", "支数", "件数", "库存", "入库", "出库", "报工", "转序", "不良", "报废", "返修"],
    "amount": ["金额", "单价", "总价", "price", "amount", "cost", "费用"],
    "project": ["项目", "project"],
    "product_code": ["产品编码", "产品编号", "成品编码", "零件号", "图号", "型号", "料号", "物料编码", "编码"],
    "product_name": ["产品名称", "品名", "名称", "物料名称", "零件名称"],
    "process": ["工序", "工艺", "工段", "工位", "制程"],
    "equipment": ["设备", "机台", "机床", "产线", "线体"],
    "status": ["状态", "结果", "是否", "标识", "审核"],
    "remark": ["备注", "说明", "原因", "描述", "意见", "remark", "memo"],
    "customer": ["客户"],
    "supplier": ["供应商", "供方", "厂商"],
    "tooling": ["工装", "模具", "夹具", "刀具"],
    "takt": ["节拍", "节奏", "周期", "标准工时", "工时"],
    "downtime": ["停机", "停线", "维修", "故障"],
}

BUSINESS_KEYWORDS = {
    "项目编号 / 项目名称": ["项目编号", "项目名称", "项目", "project"],
    "产品编码 / 零件号 / 图号": ["产品编码", "产品编号", "零件号", "图号", "成品编码", "型号", "产品图号"],
    "产品名称": ["产品名称", "品名", "零件名称", "物料名称"],
    "型材编码 / 物料编码": ["型材编码", "物料编码", "材料编码", "原材料编码", "物料编号", "材料编号"],
    "客户": ["客户", "客户名称"],
    "供应商": ["供应商", "供方", "厂商"],
    "日期 / 业务日期": ["业务日期", "业务时间", "日期", "时间", "date"],
    "数量": ["数量", "支数", "件数", "qty"],
    "入库数量": ["入库数量", "入库数", "收货数量"],
    "出库数量": ["出库数量", "出库数", "发货数量"],
    "库存数量": ["库存数量", "当前库存", "库存", "结存", "余额"],
    "报工数量": ["报工数量", "报工数", "完工数量"],
    "转序数量": ["转序数量", "转序数", "转出数量", "转入数量"],
    "不良数量": ["不良数量", "不良数"],
    "报废数量": ["报废数量", "报废数"],
    "返修数量": ["返修数量", "返修数"],
    "工序": ["工序", "工艺", "制程"],
    "设备": ["设备", "机台", "机床"],
    "工装": ["工装", "模具", "夹具", "刀具"],
    "停机时间": ["停机时间", "停机时长", "停线时间", "维修时间"],
    "节拍": ["节拍", "标准工时", "周期"],
    "状态": ["状态", "是否", "结果", "审核"],
    "备注": ["备注", "原因", "说明", "描述"],
}


@dataclass
class ColumnProfile:
    name: str
    dtype: str
    null_rate: float
    unique_count: int
    samples: list[str]
    flags: dict[str, bool]
    inferred_business_meaning: str
    business_meanings: list[str]
    confidence: str


@dataclass
class SourceProfile:
    source: str
    label: str
    file: str
    rows: int
    cols: int
    columns: list[ColumnProfile]


def read_table(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        for enc in ("utf-8-sig", "utf-8", "gb18030"):
            try:
                return pd.read_csv(path, dtype=str, encoding=enc, keep_default_na=False)
            except UnicodeDecodeError:
                continue
        return pd.read_csv(path, dtype=str, keep_default_na=False)
    if path.suffix.lower() in {".json", ".jsonl"}:
        text = path.read_text(encoding="utf-8-sig")
        if path.suffix.lower() == ".jsonl":
            rows = [json.loads(line) for line in text.splitlines() if line.strip()]
            return pd.DataFrame(rows, dtype=str)
        data = json.loads(text) if text.strip() else []
        if isinstance(data, dict):
            if isinstance(data.get("rows"), list):
                data = data["rows"]
            elif isinstance(data.get("data"), list):
                data = data["data"]
        return pd.DataFrame(data, dtype=str)
    return pd.read_excel(path, dtype=str).fillna("")


def source_file(source_dir: Path) -> Path | None:
    for name in ("latest.csv", "latest.json", "latest.jsonl", "latest.xls", "latest.xlsx"):
        p = source_dir / name
        if p.exists():
            return p
    return None


def norm_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return str(value).strip()


def nonempty(series: pd.Series) -> pd.Series:
    return series.map(norm_text).replace({"": pd.NA}).dropna()


def looks_date(series: pd.Series, name: str) -> bool:
    if any(k.lower() in name.lower() for k in ROLE_KEYWORDS["date"]):
        return True
    vals = nonempty(series).head(200)
    if len(vals) == 0:
        return False
    parsed = pd.to_datetime(vals, errors="coerce", format="mixed")
    return parsed.notna().mean() >= 0.65


def looks_number(series: pd.Series, name: str, keywords: list[str]) -> bool:
    lower = name.lower()
    if any(x in lower for x in ["编码", "编号", "单号", "id", "人员", "用户", "日期", "时间", "率", "车牌"]):
        return False
    if any(k.lower() in lower for k in keywords):
        return True
    vals = nonempty(series).head(200)
    if len(vals) == 0:
        return False
    cleaned = vals.astype(str).str.replace(",", "", regex=False).str.replace("%", "", regex=False)
    numeric = pd.to_numeric(cleaned, errors="coerce")
    return numeric.notna().mean() >= 0.75


def sample_values(series: pd.Series, limit: int = 10) -> list[str]:
    seen = []
    for value in series.map(norm_text):
        if not value:
            continue
        if value not in seen:
            seen.append(value)
        if len(seen) >= limit:
            break
    return seen


def flag_column(series: pd.Series, name: str) -> dict[str, bool]:
    lower = name.lower()
    vals = sample_values(series, 30)
    joined = " ".join(vals).lower()
    text = lower + " " + joined
    code_context = any(k in name for k in ["产品", "物资", "物料", "材料", "规格", "图号", "零件", "型材", "领料"])
    code_exclude = any(k in name for k in ["业务单号", "异常单号", "单据号", "公司", "设备", "资产", "当前用户", "人员", "组织", "库房", "捆包", "车牌", "制程码", "质量跟踪码"])
    flags = {
        "疑似日期字段": looks_date(series, name),
        "疑似数量字段": looks_number(series, name, ROLE_KEYWORDS["quantity"]),
        "疑似金额字段": looks_number(series, name, ROLE_KEYWORDS["amount"]) and any(k.lower() in lower for k in ROLE_KEYWORDS["amount"]),
        "疑似项目字段": any(k.lower() in text for k in ROLE_KEYWORDS["project"]),
        "疑似产品编码字段": code_context and not code_exclude and (any(k.lower() in lower for k in ["产品编码", "产品编号", "成品编码", "零件号", "图号", "规格图号", "物料编码", "物资编码", "物资id", "物料id", "领料物资id"]) or code_like_ratio(vals) >= 0.65),
        "疑似产品名称字段": any(k.lower() in lower for k in ROLE_KEYWORDS["product_name"]) and not any(k in name for k in ["设备名称", "公司名称", "单据类型名称", "部门名称"]),
        "疑似工序字段": any(k.lower() in text for k in ROLE_KEYWORDS["process"]),
        "疑似设备字段": any(k.lower() in text for k in ROLE_KEYWORDS["equipment"]),
        "疑似状态字段": any(k.lower() in text for k in ROLE_KEYWORDS["status"]) or small_category(vals),
        "疑似备注字段": any(k.lower() in text for k in ROLE_KEYWORDS["remark"]) or long_text_ratio(vals) >= 0.4,
    }
    return {k: bool(v) for k, v in flags.items()}


def code_like_ratio(vals: list[str]) -> float:
    if not vals:
        return 0.0
    hits = 0
    for v in vals:
        s = v.strip()
        if re.search(r"[A-Za-z]", s) and re.search(r"\d", s) and len(s) >= 4:
            hits += 1
        elif re.fullmatch(r"[A-Z0-9][A-Z0-9\-_/\.]{4,}", s, flags=re.I):
            hits += 1
    return hits / len(vals)


def long_text_ratio(vals: list[str]) -> float:
    if not vals:
        return 0.0
    return sum(1 for v in vals if len(v) >= 18) / len(vals)


def small_category(vals: list[str]) -> bool:
    if not vals:
        return False
    status_words = ["已", "未", "完成", "审核", "关闭", "正常", "异常", "合格", "不合格", "是", "否"]
    return any(any(w in v for w in status_words) for v in vals)


def infer_meaning(name: str, series: pd.Series) -> tuple[str, str]:
    meanings = infer_business_fields(name, series)
    if meanings:
        return meanings[0][0], meanings[0][1]
    return "待人工确认", "低"


def infer_business_fields(name: str, series: pd.Series) -> list[tuple[str, str]]:
    samples = sample_values(series, 20)
    nonempty_count = len(samples)
    out: list[tuple[str, str]] = []

    def add(field: str, confidence: str = "高") -> None:
        if field not in [x[0] for x in out]:
            out.append((field, confidence))

    if name in {"产品/项目", "项目/车型"}:
        add("项目编号 / 项目名称", "高")
    if name in {"产品编码"}:
        add("产品编码 / 零件号 / 图号", "高")
    if name in {"规格图号", "规格型号"}:
        add("产品编码 / 零件号 / 图号", "中")
    if name in {"物资编码", "物料编码", "物资ID", "领料物资ID"}:
        add("型材编码 / 物料编码", "高")
        if name in {"物资编码", "物料编码"}:
            add("产品编码 / 零件号 / 图号", "中")
    if name in {"产品名称"}:
        add("产品名称", "高")
    if name in {"物资名称", "物料名称", "领料物资名称"}:
        add("产品名称", "中")
    if name in {"客户", "客户名称", "客户全称", "客户ID", "默认客户"}:
        add("客户", "高")
    if name in {"供应商"}:
        add("供应商", "高")
    if name in {"合作伙伴", "二级"}:
        add("供应商", "低")
    if name in {"业务时间", "申报时间", "报停时间", "重启时间", "审核时间", "报工时间", "操作时间", "在线时间", "起始时间", "截止时间", "维护时间", "当前时间", "上传时间", "开始日期", "要货日期"}:
        add("日期 / 业务日期", "高")
    if name in {"数量", "计划量", "完成量", "单车用量", "单台用量", "毛坯重【KG】", "产品重【KG】", "铝块重【KG】", "铝屑重【KG】", "总量", "待入库量", "出库量", "在库量", "入库", "出库", "在途"}:
        add("数量", "中")
    if name in {"入库数量", "入库"}:
        add("入库数量", "高")
    if name in {"出库量", "出库"}:
        add("出库数量", "高")
    if name in {"期初库存", "期末库存", "当前库存", "虚拟库存", "在库量"}:
        add("库存数量", "高")
    if name in {"合格品", "合格数", "上线数", "完成量"}:
        add("报工数量", "中")
    if name in {"合格品"}:
        add("转序数量", "中")
    if name in {"不良品"}:
        add("不良数量", "高")
    if name in {"试料数(报废)"}:
        add("报废数量", "高")
    if name in {"工序"}:
        add("工序", "高")
    if name in {"设备号", "设备名称", "设备型号", "设备ID", "加工设备", "机台"}:
        add("设备", "高")
    if name in {"夹具码", "夹具维护"}:
        add("工装", "高")
    if name in {"报停时间", "重启时间", "报停时长", "审核时长", "用时H"}:
        add("停机时间", "高" if "停" in name else "中")
    if name in {"总节拍【分】", "生产时长【分】", "辅助时长【分】"}:
        add("节拍", "高")
    if name in {"业务状态", "集成标志", "完成率", "属性", "后序处置", "入库性质", "入出库标识", "排程分类", "标签来源"}:
        add("状态", "高")
    if name in {"缺陷描述", "申报说明", "停机原因", "原因描述", "审核说明", "单据描述", "行描述", "器具描述", "历史维护信息", "备注"}:
        add("备注", "高")
    if name == "后序处置" and any("返修" in s for s in samples):
        add("返修数量", "低")
    if name == "后序处置" and any("报废" in s for s in samples):
        add("报废数量", "低")

    if not out:
        flags = flag_column(series, name)
        if flags["疑似日期字段"]:
            add("日期 / 业务日期", "中")
        elif flags["疑似数量字段"] and nonempty_count:
            add("数量", "低")
        elif flags["疑似备注字段"]:
            add("备注", "低")
    return out


def legacy_infer_meaning(name: str, series: pd.Series) -> tuple[str, str]:
    flags = flag_column(series, name)
    candidates = []
    lower = name.lower()
    vals = sample_values(series, 20)
    value_text = " ".join(vals).lower()
    combined = lower + " " + value_text
    for field, keys in BUSINESS_KEYWORDS.items():
        score = 0
        for key in keys:
            k = key.lower()
            if k in lower:
                score += 3
            elif k in combined:
                score += 1
        candidates.append((score, field))
    candidates.sort(reverse=True)
    if candidates and candidates[0][0] >= 3:
        return candidates[0][1], "高"
    if flags["疑似日期字段"]:
        return "日期 / 业务日期", "中"
    if flags["疑似数量字段"]:
        return "数量", "中"
    if flags["疑似产品编码字段"]:
        return "产品编码 / 零件号 / 图号", "中"
    if flags["疑似状态字段"]:
        return "状态", "低"
    if flags["疑似备注字段"]:
        return "备注", "低"
    return "待人工确认", "低"


def profile_sources() -> tuple[list[SourceProfile], dict[str, pd.DataFrame]]:
    profiles = []
    frames = {}
    for source_dir in sorted(DATA_ROOT.iterdir()):
        if not source_dir.is_dir():
            continue
        path = source_file(source_dir)
        if path is None:
            continue
        df = read_table(path).fillna("")
        df.columns = [str(c).strip() for c in df.columns]
        frames[source_dir.name] = df
        columns = []
        for col in df.columns:
            ser = df[col]
            values = ser.map(norm_text)
            empty = values.eq("").sum()
            samples = sample_values(ser)
            meaning, confidence = infer_meaning(col, ser)
            meanings = [x[0] for x in infer_business_fields(col, ser)]
            columns.append(
                ColumnProfile(
                    name=col,
                    dtype=infer_dtype(ser),
                    null_rate=round(empty / len(df), 4) if len(df) else 1.0,
                    unique_count=int(nonempty(ser).nunique()),
                    samples=samples,
                    flags=flag_column(ser, col),
                    inferred_business_meaning=meaning,
                    business_meanings=meanings,
                    confidence=confidence,
                )
            )
        profiles.append(
            SourceProfile(
                source=source_dir.name,
                label=SOURCE_LABELS.get(source_dir.name, source_dir.name),
                file=str(path.relative_to(ROOT)),
                rows=int(len(df)),
                cols=int(len(df.columns)),
                columns=columns,
            )
        )
    return profiles, frames


def infer_dtype(series: pd.Series) -> str:
    vals = nonempty(series)
    if len(vals) == 0:
        return "empty"
    parsed_date = pd.to_datetime(vals.head(200), errors="coerce", format="mixed")
    if parsed_date.notna().mean() >= 0.8:
        return "datetime-like"
    cleaned = vals.astype(str).str.replace(",", "", regex=False).str.replace("%", "", regex=False)
    numeric = pd.to_numeric(cleaned.head(500), errors="coerce")
    if numeric.notna().mean() >= 0.9:
        return "numeric-like"
    return "text"


def md_table(headers: list[str], rows: list[list[Any]]) -> str:
    out = ["| " + " | ".join(headers) + " |", "| " + " | ".join(["---"] * len(headers)) + " |"]
    for row in rows:
        out.append("| " + " | ".join(md_cell(x) for x in row) + " |")
    return "\n".join(out)


def md_cell(value: Any) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\n", "<br>").replace("|", "\\|")
    if len(text) > 180:
        text = text[:177] + "..."
    return text


def yesno(v: bool) -> str:
    return "是" if v else "否"


def write_inventory(profiles: list[SourceProfile]) -> None:
    lines = ["# 01 数据文件与字段盘点", ""]
    for p in profiles:
        lines += [
            f"## {p.label}（`{p.source}`）",
            "",
            f"- 文件名 / 表名：`{p.file}` / `{p.source}`",
            f"- 行数：{p.rows}",
            f"- 列数：{p.cols}",
            "",
        ]
        rows = []
        for c in p.columns:
            rows.append(
                [
                    c.name,
                    c.dtype,
                    f"{c.null_rate:.1%}",
                    c.unique_count,
                    "；".join(c.samples[:10]),
                    yesno(c.flags["疑似日期字段"]),
                    yesno(c.flags["疑似数量字段"]),
                    yesno(c.flags["疑似金额字段"]),
                    yesno(c.flags["疑似项目字段"]),
                    yesno(c.flags["疑似产品编码字段"]),
                    yesno(c.flags["疑似产品名称字段"]),
                    yesno(c.flags["疑似工序字段"]),
                    yesno(c.flags["疑似设备字段"]),
                    yesno(c.flags["疑似状态字段"]),
                    yesno(c.flags["疑似备注字段"]),
                    c.inferred_business_meaning,
                    c.confidence,
                ]
            )
        lines.append(
            md_table(
                [
                    "字段",
                    "数据类型",
                    "空值率",
                    "唯一值数",
                    "前10个样本值",
                    "日期",
                    "数量",
                    "金额",
                    "项目",
                    "产品编码",
                    "产品名称",
                    "工序",
                    "设备",
                    "状态",
                    "备注",
                    "推断业务含义",
                    "置信度",
                ],
                rows,
            )
        )
        lines.append("")
    (REPORT_ROOT / "01_data_inventory.md").write_text("\n".join(lines), encoding="utf-8")


def build_business_mapping(profiles: list[SourceProfile]) -> dict[str, list[dict[str, Any]]]:
    mapping: dict[str, list[dict[str, Any]]] = {field: [] for field in BUSINESS_FIELDS}
    needs_confirm = []
    for p in profiles:
        for c in p.columns:
            meanings = c.business_meanings or ([c.inferred_business_meaning] if c.inferred_business_meaning in mapping else [])
            for meaning in meanings:
                evidence = []
                if c.name:
                    evidence.append(f"字段名 `{c.name}`")
                if c.samples:
                    evidence.append("样本：" + "、".join(c.samples[:5]))
                mapping[meaning].append(
                    {
                        "source": p.source,
                        "label": p.label,
                        "field": c.name,
                        "confidence": c.confidence,
                        "evidence": "；".join(evidence),
                    }
                )
            if not meanings:
                needs_confirm.append(
                    {
                        "source": p.source,
                        "label": p.label,
                        "field": c.name,
                        "samples": "；".join(c.samples[:8]),
                    }
                )
    lines = ["# 02 核心业务字段识别", "", "说明：字段归类同时参考字段名、样本值形态、数据类型、空值率和唯一值分布；低置信度项需要人工确认。", ""]
    rows = []
    for field in BUSINESS_FIELDS:
        items = mapping[field]
        if not items:
            rows.append([field, "未识别到明确字段", "", "", "缺失或待人工确认"])
            continue
        fields = []
        evidence = []
        confidence = []
        for item in items:
            fields.append(f"{item['label']}：`{item['field']}`")
            evidence.append(f"{item['label']} / `{item['field']}`：{item['evidence']}")
            confidence.append(item["confidence"])
        status = "存在" if any(c in {"高", "中"} for c in confidence) else "待人工确认"
        rows.append([field, status, "<br>".join(fields), "；".join(sorted(set(confidence))), "<br>".join(evidence[:12])])
    lines.append(md_table(["业务字段", "识别结果", "候选数据源字段", "置信度", "判断依据"], rows))
    lines += ["", "## 需要人工确认的字段", ""]
    lines.append(
        md_table(
            ["数据源", "字段", "样本值"],
            [[f"{x['label']}（`{x['source']}`）", x["field"], x["samples"]] for x in needs_confirm],
        )
    )
    (REPORT_ROOT / "02_business_field_mapping.md").write_text("\n".join(lines), encoding="utf-8")
    return mapping


def field_candidates(profiles: list[SourceProfile]) -> dict[str, dict[str, list[str]]]:
    out: dict[str, dict[str, list[str]]] = defaultdict(lambda: defaultdict(list))
    for p in profiles:
        for c in p.columns:
            for meaning in c.business_meanings or [c.inferred_business_meaning]:
                if meaning != "待人工确认":
                    out[p.source][meaning].append(c.name)
            for flag, active in c.flags.items():
                if active:
                    out[p.source][flag].append(c.name)
    return out


def overlap_score(frames: dict[str, pd.DataFrame], left: str, left_cols: list[str], right: str, right_cols: list[str]) -> str:
    if left not in frames or right not in frames or not left_cols or not right_cols:
        return "无可比字段"
    details = []
    for lc in left_cols[:3]:
        if lc not in frames[left].columns:
            continue
        lv = set(nonempty(frames[left][lc]).astype(str).str.strip())
        if not lv:
            continue
        for rc in right_cols[:3]:
            if rc not in frames[right].columns:
                continue
            rv = set(nonempty(frames[right][rc]).astype(str).str.strip())
            if not rv:
                continue
            inter = lv & rv
            denom = min(len(lv), len(rv))
            ratio = len(inter) / denom if denom else 0
            if len(inter):
                examples = "、".join(list(inter)[:5])
                details.append(f"`{lc}`↔`{rc}` 交集 {len(inter)} / min({len(lv)},{len(rv)}) = {ratio:.1%}，例：{examples}")
    return "<br>".join(details) if details else "字段存在但样本值暂无交集"


def relationship_rows(profiles: list[SourceProfile], frames: dict[str, pd.DataFrame]) -> list[list[str]]:
    fc = field_candidates(profiles)

    def cols(source: str, meaning: str) -> list[str]:
        return fc.get(source, {}).get(meaning, [])

    def product_cols(source: str) -> list[str]:
        return cols(source, "产品编码 / 零件号 / 图号") + cols(source, "型材编码 / 物料编码")

    cases = [
        ("采购入库明细 ↔ 仓库台账 / 仓库明细", "purchase_inbound", ["warehouse_ledger", "warehouse_detail"], ["物料/产品编码", "产品名称", "项目", "日期"]),
        ("仓库台账 / 仓库明细 ↔ 机加转序", "warehouse_ledger", ["warehouse_detail", "machining_transfer_query"], ["物料/产品编码", "产品名称", "项目", "日期"]),
        ("机加转序 ↔ 后工序报工", "machining_transfer_query", ["post_process_report_tracking"], ["产品编码", "产品名称", "工序", "日期", "数量"]),
        ("后工序报工 ↔ 不良品跟踪", "post_process_report_tracking", ["defective_tracking"], ["产品编码", "产品名称", "工序", "数量"]),
        ("仓库台账 / 仓库明细 ↔ 销售出库明细", "warehouse_ledger", ["warehouse_detail", "sales_outbound_detail"], ["产品编码", "产品名称", "项目", "日期", "数量"]),
        ("生产节拍维护 ↔ 设备维护 ↔ 设备停机 ↔ 工装更换", "production_takt_maintenance", ["production_equipment_maintenance", "equipment_downtime_query", "tooling_replacement_query"], ["产品编码", "产品名称", "设备", "工装", "节拍", "日期"]),
    ]
    rows = []
    for title, left, rights, keys in cases:
        evidence_parts = []
        for right in rights:
            if right not in frames:
                continue
            comparisons = []
            comparisons.append(overlap_score(frames, left, product_cols(left), right, product_cols(right)))
            comparisons.append(overlap_score(frames, left, cols(left, "产品名称"), right, cols(right, "产品名称")))
            comparisons.append(overlap_score(frames, left, cols(left, "项目编号 / 项目名称"), right, cols(right, "项目编号 / 项目名称")))
            comparisons.append(overlap_score(frames, left, cols(left, "日期 / 业务日期"), right, cols(right, "日期 / 业务日期")))
            comparisons.append(overlap_score(frames, left, cols(left, "设备"), right, cols(right, "设备")))
            evidence_parts.append(f"{SOURCE_LABELS.get(left,left)} ↔ {SOURCE_LABELS.get(right,right)}：<br>" + "<br>".join(x for x in comparisons if x and x != "无可比字段"))
        available_fields = []
        for src in [left] + rights:
            src_fields = []
            for k in ["产品编码 / 零件号 / 图号", "型材编码 / 物料编码", "产品名称", "项目编号 / 项目名称", "日期 / 业务日期", "数量", "工序", "设备", "工装", "节拍"]:
                if cols(src, k):
                    src_fields.append(f"{k}={','.join(cols(src,k)[:3])}")
            available_fields.append(f"{SOURCE_LABELS.get(src,src)}：{'；'.join(src_fields) if src_fields else '未识别明确关联字段'}")
        joined_evidence = "<br>".join(evidence_parts)
        can = "可尝试关联（需确认字段口径）" if re.search(r"交集 \d+", joined_evidence) else "关联能力弱 / 需人工确认键"
        rows.append([title, "、".join(keys), "<br>".join(available_fields), joined_evidence or "没有足够字段形成样本交集", can])
    return rows


def write_relationships(profiles: list[SourceProfile], frames: dict[str, pd.DataFrame]) -> None:
    lines = ["# 03 数据源可关联性分析", "", "说明：下表基于候选字段是否存在、样本值是否有交集来判断。字段含义不明确或值无交集时，不强行认定可关联。", ""]
    lines.append(md_table(["重点关系", "尝试关联键", "各表可用字段", "样本值交集证据", "结论"], relationship_rows(profiles, frames)))
    lines += [
        "",
        "## 关系判断摘要",
        "",
        "- 采购入库、仓库明细/台账、销售出库如果存在物料/产品编码和数量字段，可支持进销存的事后对账；能否做到按项目或日期精确追溯，取决于这些字段在各表中的一致性。",
        "- 机加转序、后工序报工、不良品跟踪之间需要共同的产品/零件标识、工序标识和数量字段；如果缺少统一流转单号，只能做产品维度或日期窗口维度的近似关联。",
        "- 节拍、设备、停机、工装更换之间如果只有设备或产品字段交集，可以分析产能影响方向；要精确量化，还需要班次日历、实际生产时段、工装适配关系和计划产量。",
    ]
    (REPORT_ROOT / "03_relationship_analysis.md").write_text("\n".join(lines), encoding="utf-8")


def has_any(mapping: dict[str, list[dict[str, Any]]], fields: list[str]) -> bool:
    return any(mapping.get(f) for f in fields)


def field_list(mapping: dict[str, list[dict[str, Any]]], fields: list[str]) -> str:
    items = []
    for f in fields:
        for item in mapping.get(f, []):
            items.append(f"{item['label']}.`{item['field']}`")
    return "；".join(items[:16]) if items else "未识别到明确字段"


def write_available_analysis(mapping: dict[str, list[dict[str, Any]]]) -> None:
    questions = [
        ("物料类", "当前哪些产品有库存？", ["warehouse_ledger", "warehouse_detail"], ["产品编码 / 零件号 / 图号", "产品名称", "库存数量"], "需要库存台账/明细中有产品或物料标识及库存数量。"),
        ("物料类", "哪些产品近期有采购入库？", ["purchase_inbound"], ["产品编码 / 零件号 / 图号", "型材编码 / 物料编码", "日期 / 业务日期", "入库数量"], "日期窗口导出可做近期入库统计。"),
        ("物料类", "哪些产品库存为 0？", ["warehouse_ledger"], ["产品编码 / 零件号 / 图号", "库存数量"], "若库存字段为当前快照，可直接筛选；若仅明细流水，则需计算结存口径。"),
        ("物料类", "哪些产品库存异常波动？", ["warehouse_detail", "warehouse_ledger"], ["产品编码 / 零件号 / 图号", "日期 / 业务日期", "入库数量", "出库数量", "库存数量"], "需要连续时间快照或完整流水，目前 latest 只代表最新导出。"),
        ("物料类", "是否能判断缺料？", ["warehouse_ledger", "weekly_plan_execution_tracking", "product_material_match_maintenance"], ["产品编码 / 零件号 / 图号", "库存数量", "数量"], "周计划执行跟踪可作为需求/交付计划来源，产品材料匹配可作为用料关系来源；需要把周计划按日展开，并确认单车用量、单台用量和库存口径。"),
        ("物料类", "是否能判断应该催料？", ["purchase_inbound", "weekly_plan_execution_tracking", "product_material_match_maintenance", "warehouse_ledger"], ["日期 / 业务日期", "数量", "库存数量"], "可用周计划需求、材料匹配、库存和采购入库做催料线索；精确催料仍缺采购订单交期、未到货量和供应商承诺日期。"),
        ("生产流转类", "哪些产品已经机加转序？", ["machining_transfer_query"], ["产品编码 / 零件号 / 图号", "产品名称", "转序数量", "日期 / 业务日期"], "机加转序表按时间范围导出，可事后统计。"),
        ("生产流转类", "哪些产品进入后工序？", ["post_process_report_tracking"], ["产品编码 / 零件号 / 图号", "产品名称", "报工数量", "工序"], "后工序报工表可显示已报工记录。"),
        ("生产流转类", "哪些产品机加后没有后工序记录？", ["machining_transfer_query", "post_process_report_tracking"], ["产品编码 / 零件号 / 图号", "转序数量", "报工数量", "日期 / 业务日期"], "缺少统一流转单号时只能用产品+日期窗口近似比对。"),
        ("生产流转类", "哪些产品后工序报工异常？", ["post_process_report_tracking"], ["产品编码 / 零件号 / 图号", "报工数量", "工序", "状态"], "需要异常状态或标准节拍/计划量作为判定口径。"),
        ("生产流转类", "是否能判断生产卡在哪个环节？", ["machining_transfer_query", "post_process_report_tracking", "weekly_plan_execution_tracking"], ["工序", "数量", "日期 / 业务日期", "状态"], "后工序报工可提炼事实工序路径和顺序，周计划可提供需求目标；仍需确认工序顺序是否等同标准工艺路线，以及WIP/卡点判定口径。"),
        ("质量类", "哪些产品存在不良？", ["defective_tracking"], ["产品编码 / 零件号 / 图号", "产品名称", "不良数量"], "不良跟踪表可做产品维度统计。"),
        ("质量类", "哪些产品存在报废？", ["defective_tracking"], ["产品编码 / 零件号 / 图号", "报废数量"], "若表中有报废字段可直接筛选，否则需人工确认不良状态字段。"),
        ("质量类", "是否能计算不良率？", ["defective_tracking", "post_process_report_tracking"], ["不良数量", "报工数量", "产品编码 / 零件号 / 图号"], "需要同口径产出数量与不良数量可关联。"),
        ("质量类", "是否能判断不良是否影响交付？", ["defective_tracking", "sales_outbound_detail", "weekly_plan_execution_tracking"], ["不良数量", "出库数量", "日期 / 业务日期"], "需要交付计划/需求日期和不良处置状态。"),
        ("质量类", "是否能判断报废是否已退料或扣账？", ["defective_tracking", "warehouse_detail"], ["报废数量", "库存数量", "日期 / 业务日期"], "需要报废退料/扣账单号或库存扣减原因。"),
        ("产能类", "是否能计算理论产能？", ["production_takt_maintenance", "production_equipment_maintenance"], ["节拍", "设备", "产品编码 / 零件号 / 图号"], "产品-设备-节拍已经存在，可计算单件理论节拍产能；折算到班/日还需要可用工时口径，当前最关键缺口是设备对应工装/夹具适配信息。"),
        ("产能类", "是否能判断设备停机影响？", ["equipment_downtime_query", "production_equipment_maintenance"], ["设备", "停机时间", "日期 / 业务日期"], "可统计停机，但影响产能还需计划生产时段和设备-产品任务。"),
        ("产能类", "是否能判断工装更换影响？", ["tooling_replacement_query", "production_takt_maintenance"], ["工装", "设备", "日期 / 业务日期"], "需要工装更换耗时、换型前后产品和排产计划。"),
        ("产能类", "是否能判断某产品当前产能是否够？", ["production_takt_maintenance", "weekly_plan_execution_tracking"], ["节拍", "产品编码 / 零件号 / 图号", "数量"], "周计划可作为需求计划，节拍可作为理论产能基础；需要补充或提炼设备-工装适配，并确认可用工时和换型影响。"),
        ("交付类", "哪些产品已经销售出库？", ["sales_outbound_detail"], ["产品编码 / 零件号 / 图号", "产品名称", "出库数量", "日期 / 业务日期"], "销售出库明细可事后统计。"),
        ("交付类", "哪些产品有库存但未出库？", ["warehouse_ledger", "sales_outbound_detail"], ["库存数量", "出库数量", "产品编码 / 零件号 / 图号"], "可按产品做近似比对；需一致编码和时间口径。"),
        ("交付类", "哪些产品出库异常？", ["sales_outbound_detail", "warehouse_detail"], ["出库数量", "库存数量", "状态"], "异常需定义：负库存、超计划、无库存出库等业务口径。"),
        ("交付类", "是否能判断交付风险？", ["weekly_plan_execution_tracking", "sales_outbound_detail", "warehouse_ledger"], ["日期 / 业务日期", "数量", "库存数量"], "周计划执行跟踪可承担客户需求/交付计划清洗来源；需要按日展开计划量，并与库存、销售出库、生产完成量统一编码。"),
        ("交付类", "是否能判断未来几天是否来得及交付？", ["weekly_plan_execution_tracking", "production_takt_maintenance", "post_process_report_tracking"], ["日期 / 业务日期", "节拍", "数量", "工序"], "可用周计划日需求、节拍和后工序事实进度做部分判断；仍需设备-工装适配、可用工时和剩余工序口径。"),
    ]
    rows = []
    for category, q, sources, fields, note in questions:
        present = [f for f in fields if mapping.get(f)]
        missing = [f for f in fields if not mapping.get(f)]
        if not missing and q not in {"哪些产品库存异常波动？", "是否能判断缺料？", "是否能判断应该催料？", "哪些产品机加后没有后工序记录？", "是否能判断生产卡在哪个环节？", "是否能判断不良是否影响交付？", "是否能判断报废是否已退料或扣账？", "是否能判断某产品当前产能是否够？", "是否能判断交付风险？", "是否能判断未来几天是否来得及交付？"}:
            level = "A"
            can = "可直接判断"
            issue = "字段基本存在；仍需确认导出范围和字段口径。"
        elif present and q in {"哪些产品库存异常波动？", "是否能判断缺料？", "是否能判断应该催料？", "哪些产品机加后没有后工序记录？", "是否能判断生产卡在哪个环节？", "是否能计算不良率？", "是否能判断设备停机影响？", "是否能判断工装更换影响？", "是否能判断某产品当前产能是否够？", "哪些产品有库存但未出库？", "哪些产品出库异常？", "是否能判断交付风险？", "是否能判断未来几天是否来得及交付？"}:
            level = "B"
            can = "可部分判断"
            issue = note + (" 缺失/不明确：" + "、".join(missing) if missing else "")
        elif present:
            level = "C" if "已经" in q or "存在" in q or "统计" in note else "D"
            can = "可部分判断" if level in {"B", "C"} else "暂时不能判断"
            issue = note + (" 缺失/不明确：" + "、".join(missing) if missing else "")
        else:
            level = "D"
            can = "暂时不能判断"
            issue = "未识别到关键字段：" + "、".join(fields)
        rows.append([f"{category}：{q}", can, "；".join(SOURCE_LABELS.get(s, s) for s in sources), field_list(mapping, fields), issue, level])
    lines = ["# 04 当前数据可支持的业务判断", "", md_table(["分析问题", "当前能否判断", "需要哪些数据源", "依赖字段", "当前问题", "判断等级"], rows)]
    (REPORT_ROOT / "04_available_analysis.md").write_text("\n".join(lines), encoding="utf-8")


def write_missing_data(mapping: dict[str, list[dict[str, Any]]]) -> None:
    missing_specs = [
        ("判断缺料、催料、未来交付风险", "周计划执行跟踪的清洗结果：按日需求、产品编码、计划量、完成量、属性；另需采购订单交期、未到货量", "周计划字段存在但需清洗；采购交期/未到货量缺失", "周计划执行跟踪已经可承担客户需求/交付计划来源，但要先把年周、开始日期、星期日到星期六展开成日计划；催料还需要供应商承诺和未到货。", "P0"),
        ("判断生产是否来得及", "周计划日需求、后工序事实进度、节拍产能、设备-工装适配", "核心字段存在但需整理；设备-工装适配缺失", "现有数据能形成需求、进度、节拍三块基础，但缺少设备对应工装/夹具能力后，无法判断某设备能否生产某产品。", "P0"),
        ("判断生产卡在哪个环节", "从后工序报工提炼出的工序顺序、标准路线确认、每道工序WIP/完成口径", "字段存在但需要提炼和业务确认", "后工序报工有工序、制程码、操作时间、报工时间，可提炼事实路线；但事实出现顺序不一定等于标准工艺路线，需要人工确认例外和返工路径。", "P0"),
        ("按产品判断缺料", "产品与型材/物料对应关系、用量系数、库存口径、周计划日需求", "字段存在但需校验字段质量", "产品材料匹配维护有产品、领料物资和单台用量，可作为用料关系基础；需确认覆盖率、替代料和有效期。", "P0"),
        ("计算理论产能与排产能力", "设备-工装/夹具适配关系、可用工时口径、换型影响", "节拍字段存在；工装适配缺失", "生产节拍维护已提供产品、设备和总节拍，可计算理论节拍产能；当前短板是设备可加工产品与工装/夹具限制。", "P1"),
        ("量化停机影响", "设备任务计划或从周计划/节拍推导的设备负荷、停机影响产品/工序", "停机字段存在但缺少任务关联", "停机查询有设备、报停/重启时间和时长；若不能关联到当时设备任务，只能统计停机，不能精确量化影响哪一批产品。", "P1"),
        ("量化工装更换影响", "工装与产品/设备适配关系、换型前后产品、换型计划", "工装更换记录存在但关联键不足", "工装更换查询有设备、零件、用时、夹具维护，但缺少稳定的产品编码或工装主数据，难以连到周计划和节拍。", "P1"),
        ("判断不良是否影响交付", "异常处理状态、返修闭环记录、返修完成日期、可交付判定", "完全缺失或业务口径缺失", "不良数量不等于不可交付数量，需要处置结果。", "P1"),
        ("判断报废是否已退料或扣账", "报废退料记录、库存扣减单号、扣账原因", "完全缺失或缺少关联键", "不良/报废与仓库扣减之间没有统一单号时无法闭环。", "P1"),
        ("解释交付计划波动", "客户计划变更记录、版本号、变更原因", "完全缺失", "没有计划版本就无法区分生产延误和客户变更。", "P2"),
    ]
    lines = ["# 05 缺失数据分析", ""]
    lines.append(md_table(["想要判断的问题", "当前缺什么数据", "是完全缺失还是字段质量差", "为什么影响判断", "优先级"], missing_specs))
    lines += [
        "",
        "## 对现有字段的补充说明",
        "",
        f"- 已识别到的产品/物料相关字段：{field_list(mapping, ['产品编码 / 零件号 / 图号', '型材编码 / 物料编码', '产品名称'])}",
        f"- 已识别到的数量相关字段：{field_list(mapping, ['数量', '入库数量', '出库数量', '库存数量', '报工数量', '转序数量', '不良数量', '报废数量', '返修数量'])}",
        f"- 已识别到的时间相关字段：{field_list(mapping, ['日期 / 业务日期', '停机时间'])}",
        "- 修正判断：周计划执行跟踪已经可作为客户需求计划/交付计划的数据来源，问题不再是完全缺失，而是需要清洗成年周-日期-产品-日计划量-完成量的标准结构。",
        "- 修正判断：后工序报工可以提炼事实工艺路线和工序顺序，但需要业务确认标准路线、返工路径和例外工序。",
        "- 修正判断：生产节拍维护已经能提供产品-设备-节拍，理论产能基础存在；当前更关键的缺口是设备与工装/夹具、产品适配关系。",
    ]
    (REPORT_ROOT / "05_missing_data_analysis.md").write_text("\n".join(lines), encoding="utf-8")


def write_summary(profiles: list[SourceProfile], mapping: dict[str, list[dict[str, Any]]]) -> None:
    total_rows = sum(p.rows for p in profiles)
    total_cols = sum(p.cols for p in profiles)
    emptyish = []
    needs_confirm = []
    for p in profiles:
        for c in p.columns:
            if c.null_rate >= 0.8:
                emptyish.append(f"{p.label}.`{c.name}` 空值率 {c.null_rate:.0%}")
            if c.inferred_business_meaning == "待人工确认":
                needs_confirm.append(f"{p.label}.`{c.name}`")
    value_uses = [
        "用仓库台账/仓库明细识别当前库存、有库存产品、零库存产品和基础库存结构。",
        "用采购入库明细统计近期到料情况，并和库存字段做事后对账。",
        "用机加转序与后工序报工做产品维度的生产流转追踪，前提是编码/名称可统一。",
        "用不良品跟踪识别不良、报废或返修相关产品，并尝试与报工量计算不良率。",
        "用节拍、设备、停机、工装数据建立产能影响的基础事实库。"
    ]
    unsupported = [
        "无法直接给出稳定的未来交付风险结论：周计划可作为需求/交付计划，但需要按日清洗，并与库存、报工、销售出库统一编码。",
        "无法可靠判断是否应该催料到具体供应商/订单：周计划、材料匹配、库存可提供缺料线索，但缺采购订单交期和未到货量。",
        "无法把事实工序顺序直接等同标准工艺路线：后工序报工可提炼路线，但返工、跳序、例外工序需要业务确认。",
        "无法闭环判断报废是否已退料或扣账：缺少报废退料/库存扣减关联单据。",
        "无法量化工装对具体产品排产的限制：当前缺少稳定的设备-工装/夹具-产品适配关系。"
    ]
    priority_tables = [
        "周计划清洗宽表：年周、日期、项目/车型、产品编码/名称、日计划量、周计划量、完成量、完成率、属性。",
        "后工序路线提炼表：产品编码、产品名称、制程码、工序、工序序号、首次/末次报工时间、上线数、合格数、不良数。",
        "设备工装适配表：设备ID/设备名称、工装/夹具编码、适配产品/零件、换型耗时、维护状态、有效期。",
    ]
    master_fields = [
        "项目编号/名称", "客户", "产品编码/零件号/图号", "产品名称", "物料/型材编码", "计划交付日期",
        "客户需求数量", "当前库存数量", "已采购入库数量", "机加转序数量", "后工序报工数量", "不良/报废/返修数量",
        "销售出库数量", "当前工序", "下一工序", "责任设备", "节拍", "可用产能", "停机影响时长", "工装状态",
        "风险等级", "异常状态", "最后更新时间", "数据来源字段"
    ]
    lines = [
        "# 00 数据可用性分析总结",
        "",
        "## 1. 当前数据总体质量评价",
        "",
        f"- 本次读取 `data/excel_export` 下 {len(profiles)} 个数据源的 `latest` 导出，共 {total_rows} 行、{total_cols} 个字段。",
        "- 数据覆盖采购、仓库、销售、机加、后工序、不良、设备、停机、工装、节拍、项目和周计划等环节，适合先做“事实记录盘点”和“事后统计”。",
        "- 主要风险是字段命名与业务口径不统一，部分字段需要人工确认；跨表如果缺少统一编码或流转单号，结论只能停留在产品/日期窗口的近似关联。",
        f"- 高空值字段示例：{'；'.join(emptyish[:10]) if emptyish else '未发现空值率超过 80% 的字段'}。",
        f"- 需要人工确认字段数量：{len(needs_confirm)}。详见 `02_business_field_mapping.md`。",
        "",
        "## 2. 当前数据最有价值的 5 个用途",
        "",
        *[f"{i}. {x}" for i, x in enumerate(value_uses, 1)],
        "",
        "## 3. 当前数据暂时不能支撑的 5 个判断",
        "",
        *[f"{i}. {x}" for i, x in enumerate(unsupported, 1)],
        "",
        "## 4. 最关键的缺失数据",
        "",
        "- 周计划清洗结果：周计划执行跟踪已有计划量、完成量和日分布字段，但需要整理为可关联的日需求/交付计划。",
        "- 后工序路线提炼结果：后工序报工已有工序事实记录，但需要沉淀成产品-工序顺序，并标出需人工确认的异常路径。",
        "- 设备-工装/夹具-产品适配：这是当前理论产能转排产约束时最关键的缺口。",
        "- 产品-物料/BOM/型材对应与需求系数：已有产品材料匹配维护，但要确认覆盖率、替代料和有效期。",
        "- 异常闭环数据：不良处置、返修、报废退料、库存扣账和客户计划变更。",
        "",
        "## 5. 最建议优先补充的 3 张表",
        "",
        *[f"{i}. {x}" for i, x in enumerate(priority_tables, 1)],
        "",
        "## 6. 是否值得建立生产节奏主控表",
        "",
        "值得，但建议先作为“清洗汇总视图/主控宽表”逐步生成，而不是现在直接新增复杂业务主表。依据是：周计划可提供需求与交付目标，后工序报工可提炼工序进度，节拍可提供理论产能，库存/采购/销售/质量/停机可提供事实约束；当前真正需要补强的是设备-工装/夹具-产品适配和若干业务口径确认。",
        "",
        "## 7. 如果建立主控表，建议字段",
        "",
        "、".join(master_fields),
        "",
        "## 8. 下一步分析方向",
        "",
        "下一步不应停留在字段盘点，而应进入“信息提炼层”：从周计划提炼日需求，从后工序报工提炼事实工艺路线，从节拍提炼理论产能，从库存/采购/销售/质量/停机/工装组合推导库存覆盖、生产进度、质量风险、产能约束和交付风险。详见 `06_derived_information_analysis.md`。",
    ]
    (REPORT_ROOT / "00_summary.md").write_text("\n".join(lines), encoding="utf-8")


def safe_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series(dtype=str)
    return nonempty(df[col]).astype(str).str.strip()


def unique_overlap(frames: dict[str, pd.DataFrame], left: str, left_col: str, right: str, right_col: str) -> tuple[int, int, int, list[str]]:
    if left not in frames or right not in frames:
        return 0, 0, 0, []
    lv = set(safe_series(frames[left], left_col))
    rv = set(safe_series(frames[right], right_col))
    inter = sorted(lv & rv)
    return len(lv), len(rv), len(inter), inter[:8]


def overlap_evidence(frames: dict[str, pd.DataFrame], pairs: list[tuple[str, str, str, str]]) -> str:
    parts = []
    for left, left_col, right, right_col in pairs:
        lcnt, rcnt, icnt, examples = unique_overlap(frames, left, left_col, right, right_col)
        if not lcnt and not rcnt:
            continue
        ratio = icnt / min(lcnt, rcnt) if min(lcnt, rcnt) else 0
        parts.append(
            f"{SOURCE_LABELS.get(left,left)}.`{left_col}` ↔ {SOURCE_LABELS.get(right,right)}.`{right_col}`：交集 {icnt}/min({lcnt},{rcnt})={ratio:.1%}"
            + (f"，例：{'、'.join(examples)}" if examples else "")
        )
    return "<br>".join(parts) if parts else "暂无可量化交集证据"


def write_derived_information_analysis(frames: dict[str, pd.DataFrame]) -> None:
    rows = [
        [
            "周计划执行跟踪",
            "客户需求/交付计划",
            "`年周` + `开始日期` + `星期日`~`星期六` 可展开为日需求；`计划量`、`完成量`、`完成率` 表示周计划执行。",
            "可得到产品-日期-计划量、周计划完成率、短缺数量、未来几天需求。",
            "高",
            "需要确认 `星期日` 是否从 `开始日期` 当天开始；需要处理空日计划、属性、本厂/外协等口径。",
        ],
        [
            "后工序报工跟踪",
            "事实工艺路线 / 后工序进度",
            "`物资编码`、`工序`、`操作时间`、`报工时间`、`上线数`、`合格数`、`不良品` 可按产品和时间排序。",
            "可提炼每个产品实际经过哪些工序、工序先后、各工序报工数量、后工序在制/积压线索。",
            "中-高",
            "事实路线不等于标准工艺路线；返工、跳序、补录会影响顺序，需要人工确认异常路径。",
        ],
        [
            "生产节拍维护 + 生产设备维护",
            "产品-设备理论产能",
            "`产品编码`、`设备ID`、`设备名称`、`总节拍【分】` 可换算单位产能；设备维护可校验设备是否排程类。",
            "可得到每台设备加工某产品的理论分钟/件、小时产能、可选设备清单。",
            "高",
            "缺少设备-工装/夹具适配、班次可用工时和换型规则时，理论产能不能直接等于可排产产能。",
        ],
        [
            "产品材料匹配维护",
            "产品用料关系 / 材料需求展开",
            "`产品编码`、`领料物资ID`、`单台用量`、重量字段可作为产品到材料的映射。",
            "结合周计划可计算材料需求量；结合仓库台账可判断材料覆盖天数和潜在缺料。",
            "中-高",
            "需确认覆盖率、替代料、有效期、单位和材料损耗口径。",
        ],
        [
            "仓库台账 + 仓库明细",
            "库存快照与库存流水",
            "台账含 `期初库存`、`入库`、`出库`、`当前库存`；明细含 `入出库类型`、`用途`、`业务时间`、`数量`。",
            "可判断当前库存、库存变化原因、近期入出库结构、库存异常波动线索。",
            "高",
            "台账是快照，明细是流水；需统一期间、库房、物资编码和正负方向。",
        ],
        [
            "采购入库明细 + 仓库台账/明细",
            "到料与入库闭环",
            "采购入库 `物资ID/规格图号/数量/供应商/业务时间` 可与仓库物资编码、入库流水比对。",
            "可判断近期哪些料已到、是否进入库存、供应商到料结构。",
            "中",
            "缺采购订单、应到日期、未到货量；只能判断已发生到料，不能完整判断催料对象。",
        ],
        [
            "机加转序 + 后工序报工",
            "机加后是否进入后工序",
            "机加 `物料编码/规格型号/合格品/不良品/报工时间` 与后工序 `物资编码/工序/上线数/合格数` 可按编码和时间窗口关联。",
            "可判断机加完成后是否有后工序记录、转序数量与后工序上线/合格数量差异、后工序积压线索。",
            "中",
            "缺统一流转单号时只能按产品编码和时间窗口近似；需要防止同产品多批次混淆。",
        ],
        [
            "后工序报工 + 不良品跟踪",
            "质量损失与不良率",
            "两表均有物资编码/名称和数量类字段；报工提供分母，不良跟踪提供缺陷、处置和数量。",
            "可计算产品或工序层面的不良率、主要缺陷原因、后序处置结构。",
            "中",
            "不良跟踪未必与某次报工一一对应；报废/返修/退料闭环需要处置结果和库存扣减记录。",
        ],
        [
            "销售出库 + 周计划 + 库存",
            "交付执行与库存保障",
            "销售出库 `客户/物料编码/数量/业务时间`，周计划 `产品编码/日计划`，库存 `当前库存`。",
            "可判断计划产品是否已出库、有库存未出库、出库与计划差异、库存对近期计划的覆盖。",
            "中",
            "周计划产品编码与销售/库存物料编码需统一；销售出库是否等同客户交付需确认。",
        ],
        [
            "设备停机 + 节拍 + 周计划",
            "停机对理论产能的影响",
            "停机有设备号、报停/重启时间、时长；节拍有设备产品能力；周计划有需求。",
            "可估算设备停机损失产能，并识别受影响的计划产品范围。",
            "中",
            "缺设备当天具体生产任务时，只能按该设备可加工产品集合分摊影响。",
        ],
        [
            "工装更换 + 设备 + 节拍",
            "换型/工装对节奏的影响",
            "工装更换有设备、零件、更换项目、用时、试料报废；节拍表有设备产品能力。",
            "可评估换型耗时、试料损失、某设备当天可用时间减少。",
            "中-低",
            "缺稳定工装编码和产品适配关系；`零件` 与产品编码需建立映射。",
        ],
        [
            "项目信息维护 + 周计划/节拍/材料匹配",
            "项目/客户维度归集",
            "项目信息有项目、客户；周计划、节拍、材料匹配均有项目/车型或产品/项目。",
            "可把产品需求、产能、用料、库存、质量按客户/项目汇总。",
            "中",
            "项目命名需要标准化，如项目/车型、产品/项目、客户ID/客户名称之间需映射。",
        ],
    ]

    evidence_rows = [
        [
            "计划 ↔ 库存/销售/节拍",
            overlap_evidence(
                frames,
                [
                    ("weekly_plan_execution_tracking", "产品编码", "warehouse_ledger", "物资编码"),
                    ("weekly_plan_execution_tracking", "产品编码", "sales_outbound_detail", "物料编码"),
                    ("weekly_plan_execution_tracking", "产品编码", "production_takt_maintenance", "产品编码"),
                    ("weekly_plan_execution_tracking", "产品编码", "product_material_match_maintenance", "产品编码"),
                ],
            ),
            "周计划产品编码可作为主线，连接库存、销售、节拍、材料匹配。",
        ],
        [
            "后工序 ↔ 机加/不良",
            overlap_evidence(
                frames,
                [
                    ("post_process_report_tracking", "物资编码", "machining_transfer_query", "物料编码"),
                    ("post_process_report_tracking", "物资编码", "defective_tracking", "物料编码"),
                ],
            ),
            "可以做产品维度的生产流转与质量关联；批次级需要制程码/时间窗口。",
        ],
        [
            "节拍 ↔ 设备/停机/工装",
            overlap_evidence(
                frames,
                [
                    ("production_takt_maintenance", "设备ID", "production_equipment_maintenance", "设备ID"),
                    ("production_takt_maintenance", "设备ID", "equipment_downtime_query", "设备号"),
                    ("production_takt_maintenance", "设备ID", "tooling_replacement_query", "设备号"),
                ],
            ),
            "设备ID/设备号交集较强，可把节拍、停机、工装更换放到同一设备维度。",
        ],
        [
            "采购/仓库",
            overlap_evidence(
                frames,
                [
                    ("purchase_inbound", "物资ID", "warehouse_ledger", "物资编码"),
                    ("purchase_inbound", "物资ID", "warehouse_detail", "物资编码"),
                ],
            ),
            "采购入库可与仓库库存和流水做已到料闭环。",
        ],
    ]

    derived_metrics = [
        ["日需求量", "周计划执行跟踪", "`开始日期` + 星期字段展开", "产品-日期-计划量"],
        ["计划缺口", "周计划执行跟踪", "`计划量 - 完成量`", "周计划未完成数量"],
        ["库存覆盖", "周计划 + 仓库台账", "`当前库存 / 未来日均需求`", "可覆盖天数"],
        ["材料需求", "周计划 + 产品材料匹配", "`日计划量 * 单台用量`", "产品需求折算到材料需求"],
        ["理论小时产能", "生产节拍维护", "`60 / 总节拍【分】`", "产品-设备小时产能"],
        ["停机损失产能", "设备停机 + 节拍", "`报停时长(小时) * 小时产能`", "设备停机造成的理论少产数量"],
        ["工装换型损失", "工装更换 + 节拍", "`用时H * 小时产能 + 试料数(报废)`", "换型造成的时间和试料损失"],
        ["后工序良率", "后工序报工", "`合格数 / 上线数`", "后工序产出质量"],
        ["机加良率", "机加转序", "`合格品 / (合格品 + 不良品)`", "机加产出质量"],
        ["交付执行", "销售出库 + 周计划", "`销售出库数量 / 日或周计划量`", "交付完成情况，需确认销售出库等同交付"],
        ["库存流水方向", "仓库明细", "`入出库标识/入出库类型 + 数量`", "库存变动来源"],
        ["异常闭环状态", "不良品跟踪", "`后序处置 + 入库数量/出库量/在库量`", "不良处置进度线索"],
    ]

    lines = [
        "# 06 可提炼信息与组合分析",
        "",
        "本报告用于补足“字段盘点”之外的第二层判断：很多业务信息不是单表直接给出，而是需要从现有字段清洗、提炼、组合后得到。",
        "",
        "## 1. 单表可提炼信息",
        "",
        md_table(["数据源", "可提炼信息", "依据字段/逻辑", "可得到的业务信息", "置信度", "限制与需确认"], rows),
        "",
        "## 2. 跨表组合证据",
        "",
        md_table(["组合关系", "样本交集证据", "业务含义"], evidence_rows),
        "",
        "## 3. 可派生指标清单",
        "",
        md_table(["派生指标", "数据来源", "计算/提炼逻辑", "输出信息"], derived_metrics),
        "",
        "## 4. 建议的提炼顺序",
        "",
        "1. 先清洗周计划：展开为 `日期-项目-产品编码-产品名称-日计划量-周计划量-完成量-属性`。",
        "2. 再提炼事实工艺路线：按 `物资编码 + 工序 + 操作时间/报工时间` 排序，生成产品工序顺序，并标出多路径/跳序/返工。",
        "3. 建立产品-设备-节拍能力表：从节拍维护抽取产品、设备、总节拍，计算小时产能。",
        "4. 建立设备-工装/夹具适配表：这是当前组合分析最需要补强的主数据。",
        "5. 把库存、采购入库、销售出库、不良、停机、工装更换都接到产品/设备/日期维度，形成生产节奏主控视图。",
        "",
        "## 5. 需要人工确认的关键口径",
        "",
        "- 周计划中的 `星期日` 是否对应 `开始日期`，以及空值是否代表 0。",
        "- `销售出库明细` 是否可以等同客户交付，还是只是调拨/出库动作。",
        "- 后工序报工的事实工序顺序是否可作为标准工艺路线，哪些工序属于返工或例外。",
        "- `物料编码`、`物资编码`、`产品编码` 在不同表中是否属于同一编码体系，哪些是成品、半成品、原材料。",
        "- 工装更换表中的 `更换项目`、`零件`、`夹具维护` 如何映射到产品编码和夹具/工装编码。",
        "- 节拍表中的 `总节拍【分】` 是单件节拍、单批节拍还是包含辅助时间的标准工时。",
    ]
    (REPORT_ROOT / "06_derived_information_analysis.md").write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    REPORT_ROOT.mkdir(parents=True, exist_ok=True)
    profiles, frames = profile_sources()
    write_inventory(profiles)
    mapping = build_business_mapping(profiles)
    write_relationships(profiles, frames)
    write_available_analysis(mapping)
    write_missing_data(mapping)
    write_derived_information_analysis(frames)
    write_summary(profiles, mapping)


if __name__ == "__main__":
    main()
