from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path
from typing import Iterable

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DATA_ROOT = ROOT / "data" / "excel_export"
OUTPUT_ROOT = Path(os.environ.get("PRODUCTION_RHYTHM_OUTPUT_DIR", ROOT / "output"))
REPORT_ROOT = ROOT / "reports"
CURRENT_DATE = pd.Timestamp("2026-05-27")


def read_csv(source: str) -> pd.DataFrame:
    path = DATA_ROOT / source / "latest.csv"
    return pd.read_csv(path, dtype=str, encoding="utf-8-sig", keep_default_na=False).fillna("")


def to_num(value) -> float:
    if value is None:
        return 0.0
    s = str(value).strip().replace(",", "").replace("%", "")
    if not s:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def to_date(value) -> pd.Timestamp | pd.NaT:
    if value is None or str(value).strip() == "":
        return pd.NaT
    return pd.to_datetime(value, errors="coerce")


def date_str(value) -> str:
    dt = to_date(value)
    if pd.isna(dt):
        return ""
    return dt.strftime("%Y-%m-%d")


def write_output(df: pd.DataFrame, name: str) -> Path:
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    path = OUTPUT_ROOT / name
    df.to_csv(path, index=False, encoding="utf-8-sig")
    return path


def risk_rank(level: str) -> int:
    return {"高风险": 4, "中风险": 3, "低风险": 2, "无法判断": 1, "无计划需求": 0, "无明显风险": 0, "": 0}.get(str(level), 0)


def max_risk(values: Iterable[str]) -> str:
    values = [v for v in values if str(v).strip()]
    if not values:
        return ""
    return sorted(values, key=risk_rank, reverse=True)[0]


def first_nonempty(values: Iterable[str]) -> str:
    for v in values:
        if pd.isna(v):
            continue
        s = str(v).strip()
        if s and s.lower() not in {"nan", "none"}:
            return s
    return ""


def aggregate_reason(values: Iterable[str], limit: int = 3) -> str:
    seen = []
    for value in values:
        if pd.isna(value):
            continue
        for part in str(value).split("；"):
            part = part.strip()
            if part and part.lower() not in {"nan", "none"} and part not in seen:
                seen.append(part)
            if len(seen) >= limit:
                return "；".join(seen)
    return "；".join(seen)


def load_customer_maps(project_info: pd.DataFrame, product_material: pd.DataFrame, takt: pd.DataFrame) -> tuple[dict[str, str], dict[str, str]]:
    project_customer = {}
    if {"产品/项目", "客户名称"}.issubset(project_info.columns):
        for _, row in project_info.iterrows():
            project = str(row["产品/项目"]).strip()
            customer = str(row["客户名称"]).strip() or str(row.get("客户全称", "")).strip()
            if project and customer:
                project_customer.setdefault(project, customer)
    product_customer = {}
    if {"产品编码", "默认客户"}.issubset(product_material.columns):
        for _, row in product_material.iterrows():
            code = str(row["产品编码"]).strip()
            customer = str(row["默认客户"]).strip()
            if code and customer:
                product_customer.setdefault(code, customer)
    if {"产品编码", "客户ID"}.issubset(takt.columns):
        for _, row in takt.iterrows():
            code = str(row["产品编码"]).strip()
            customer = str(row["客户ID"]).strip()
            if code and customer:
                product_customer.setdefault(code, customer)
    return project_customer, product_customer


def build_demand_daily(weekly: pd.DataFrame, project_info: pd.DataFrame, product_material: pd.DataFrame, takt: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str]]:
    project_customer, product_customer = load_customer_maps(project_info, product_material, takt)
    day_cols = [
        ("星期日", 0),
        ("星期一", 1),
        ("星期二", 2),
        ("星期三", 3),
        ("星期四", 4),
        ("星期五", 5),
        ("星期六", 6),
    ]
    start_dates = pd.to_datetime(weekly["开始日期"], errors="coerce")
    sunday_matches_start = bool((start_dates.dropna().dt.weekday == 6).all()) if start_dates.notna().any() else False
    validation = {
        "sunday_mapping": "confirmed" if sunday_matches_start else "unable_to_confirm",
        "sunday_mapping_note": "所有非空开始日期均为星期日，按 星期日=开始日期 展开。"
        if sunday_matches_start
        else "无法确认所有开始日期均为星期日，本版仍按 星期日=开始日期 近似展开。",
        "blank_plan_note": "星期字段空值按 0 处理。",
    }

    rows = []
    for _, row in weekly.iterrows():
        start = to_date(row.get("开始日期", ""))
        if pd.isna(start):
            continue
        project = str(row.get("项目/车型", "")).strip()
        product_code = str(row.get("产品编码", "")).strip()
        customer = project_customer.get(project) or product_customer.get(product_code, "")
        for col, offset in day_cols:
            plan_date = start + pd.Timedelta(days=offset)
            if plan_date < CURRENT_DATE:
                continue
            rows.append(
                {
                    "date": plan_date.strftime("%Y-%m-%d"),
                    "week": str(row.get("年周", "")).strip(),
                    "project_code": project,
                    "product_code": product_code,
                    "product_name": str(row.get("产品名称", "")).strip(),
                    "customer": customer,
                    "daily_plan_qty": to_num(row.get(col, "")),
                    "weekly_plan_qty": to_num(row.get("计划量", "")),
                    "completed_qty": to_num(row.get("完成量", "")),
                    "completion_rate": to_num(row.get("完成率", "")),
                    "attribute": str(row.get("属性", "")).strip(),
                }
            )
    df = pd.DataFrame(rows)
    if df.empty:
        return df, validation
    group_cols = ["date", "week", "project_code", "product_code", "product_name", "customer", "attribute"]
    df = (
        df.groupby(group_cols, dropna=False, as_index=False)
        .agg(
            daily_plan_qty=("daily_plan_qty", "sum"),
            weekly_plan_qty=("weekly_plan_qty", "sum"),
            completed_qty=("completed_qty", "sum"),
            completion_rate=("completion_rate", "max"),
        )
        .sort_values(["date", "project_code", "product_code"])
    )
    df = df[
        [
            "date",
            "week",
            "project_code",
            "product_code",
            "product_name",
            "customer",
            "daily_plan_qty",
            "weekly_plan_qty",
            "completed_qty",
            "completion_rate",
            "attribute",
        ]
    ]
    return df, validation


def inventory_by_code(ledger: pd.DataFrame) -> pd.DataFrame:
    ledger = latest_ledger_snapshot(ledger)
    rows = []
    for _, row in ledger.iterrows():
        code = str(row.get("物资编码", "")).strip()
        if not code:
            continue
        rows.append(
            {
                "product_code": code,
                "product_name_inventory": str(row.get("物资名称", "")).strip(),
                "current_inventory": to_num(row.get("当前库存", "")),
            }
        )
    inv = pd.DataFrame(rows)
    if inv.empty:
        return pd.DataFrame(columns=["product_code", "current_inventory", "product_name_inventory"])
    result = inv.groupby("product_code", as_index=False).agg(
        current_inventory=("current_inventory", "sum"),
        product_name_inventory=("product_name_inventory", first_nonempty),
    )
    return result


def latest_ledger_snapshot(ledger: pd.DataFrame) -> pd.DataFrame:
    if "月度" not in ledger.columns or ledger.empty:
        return ledger.copy()
    month_num = pd.to_numeric(ledger["月度"], errors="coerce")
    if month_num.notna().any():
        latest_month = month_num.max()
        return ledger[month_num == latest_month].copy()
    latest_month_text = ledger["月度"].astype(str).str.strip().max()
    return ledger[ledger["月度"].astype(str).str.strip() == latest_month_text].copy()


def latest_ledger_month(ledger: pd.DataFrame) -> str:
    if "月度" not in ledger.columns or ledger.empty:
        return "未识别"
    month_num = pd.to_numeric(ledger["月度"], errors="coerce")
    if month_num.notna().any():
        return str(int(month_num.max()))
    return str(ledger["月度"].astype(str).str.strip().max())


def future_demand(demand_daily: pd.DataFrame, product_code: str, date: str, days: int) -> float:
    start = pd.Timestamp(date)
    end = start + pd.Timedelta(days=days - 1)
    mask = (
        (demand_daily["product_code"] == product_code)
        & (pd.to_datetime(demand_daily["date"]) >= start)
        & (pd.to_datetime(demand_daily["date"]) <= end)
    )
    return float(demand_daily.loc[mask, "daily_plan_qty"].sum())


def build_inventory_coverage(demand_daily: pd.DataFrame, ledger: pd.DataFrame) -> pd.DataFrame:
    inv = inventory_by_code(ledger)
    base = demand_daily.groupby(["date", "project_code", "product_code", "product_name"], as_index=False).agg(
        daily_plan_qty=("daily_plan_qty", "sum")
    )
    rows = []
    inv_map = inv.set_index("product_code").to_dict("index") if not inv.empty else {}
    for _, row in base.iterrows():
        code = row["product_code"]
        inv_rec = inv_map.get(code)
        future3 = future_demand(demand_daily, code, row["date"], 3)
        future7 = future_demand(demand_daily, code, row["date"], 7)
        if not inv_rec:
            current = ""
            risk = "无法判断"
            shortage = ""
            coverage_days = ""
            reason = "产品编码无法关联仓库台账物资编码"
        else:
            current = float(inv_rec["current_inventory"])
            avg_daily = future7 / 7 if future7 else 0
            coverage_days = round(current / avg_daily, 2) if avg_daily > 0 else ""
            shortage = max(0.0, future3 - current)
            if current <= 0:
                risk = "高风险"
                reason = "当前库存 <= 0"
            elif current < future3:
                risk = "高风险"
                reason = "当前库存 < 未来3天需求"
            elif current < future7:
                risk = "中风险"
                reason = "当前库存 < 未来7天需求"
            else:
                risk = "低风险"
                reason = "当前库存 >= 未来7天需求"
        rows.append(
            {
                "date": row["date"],
                "project_code": row["project_code"],
                "product_code": code,
                "product_name": row["product_name"],
                "current_inventory": current,
                "future_3_days_demand": future3,
                "future_7_days_demand": future7,
                "inventory_coverage_days": coverage_days,
                "shortage_qty": shortage,
                "inventory_risk_level": risk,
                "risk_reason": reason,
            }
        )
    return pd.DataFrame(rows).sort_values(["date", "inventory_risk_level", "product_code"])


def build_material_coverage(demand_daily: pd.DataFrame, product_material: pd.DataFrame, ledger: pd.DataFrame) -> pd.DataFrame:
    inv = inventory_by_code(ledger).rename(columns={"product_code": "material_code", "current_inventory": "current_material_inventory"})
    inv_map = inv.set_index("material_code").to_dict("index") if not inv.empty else {}
    material_rows = []
    material_map = defaultdict_list()
    for _, row in product_material.iterrows():
        product_code = str(row.get("产品编码", "")).strip()
        if not product_code:
            continue
        material_map[product_code].append(
            {
                "material_code": str(row.get("领料物资ID", "")).strip(),
                "material_name": str(row.get("领料物资名称", "")).strip(),
                "unit_usage": to_num(row.get("单台用量", "")),
            }
        )
    base = demand_daily[demand_daily["daily_plan_qty"] > 0].copy()
    for _, row in base.iterrows():
        product_code = row["product_code"]
        materials = material_map.get(product_code, [])
        if not materials:
            material_rows.append(
                {
                    "date": row["date"],
                    "product_code": product_code,
                    "product_name": row["product_name"],
                    "material_code": "",
                    "material_name": "",
                    "daily_plan_qty": row["daily_plan_qty"],
                    "unit_usage": "",
                    "required_material_qty": "",
                    "current_material_inventory": "",
                    "material_shortage_qty": "",
                    "material_coverage_days": "",
                    "material_risk_level": "无法判断",
                    "risk_reason": "缺产品-材料映射",
                }
            )
            continue
        for mat in materials:
            material_code = mat["material_code"]
            required = float(row["daily_plan_qty"]) * float(mat["unit_usage"])
            inv_rec = inv_map.get(material_code)
            if not material_code:
                current = ""
                shortage = ""
                coverage_days = ""
                risk = "无法判断"
                reason = "材料编码为空，不强行猜测材料编码"
            elif not inv_rec:
                current = ""
                shortage = ""
                coverage_days = ""
                risk = "无法判断"
                reason = "缺材料库存"
            else:
                current = float(inv_rec["current_material_inventory"])
                future7_req = future_material_required(demand_daily, product_code, row["date"], 7, mat["unit_usage"])
                avg_daily = future7_req / 7 if future7_req else 0
                coverage_days = round(current / avg_daily, 2) if avg_daily else ""
                shortage = max(0.0, required - current)
                if current <= 0:
                    risk = "高风险"
                    reason = "材料当前库存 <= 0"
                elif current < required:
                    risk = "高风险"
                    reason = "材料当前库存 < 当日材料需求"
                elif current < future7_req:
                    risk = "中风险"
                    reason = "材料当前库存 < 未来7天材料需求"
                else:
                    risk = "低风险"
                    reason = "材料库存可覆盖未来7天需求"
            material_rows.append(
                {
                    "date": row["date"],
                    "product_code": product_code,
                    "product_name": row["product_name"],
                    "material_code": material_code,
                    "material_name": mat["material_name"],
                    "daily_plan_qty": row["daily_plan_qty"],
                    "unit_usage": mat["unit_usage"],
                    "required_material_qty": required,
                    "current_material_inventory": current,
                    "material_shortage_qty": shortage,
                    "material_coverage_days": coverage_days,
                    "material_risk_level": risk,
                    "risk_reason": reason,
                }
            )
    return pd.DataFrame(material_rows)


def defaultdict_list():
    from collections import defaultdict

    return defaultdict(list)


def future_material_required(demand_daily: pd.DataFrame, product_code: str, date: str, days: int, unit_usage: float) -> float:
    return future_demand(demand_daily, product_code, date, days) * float(unit_usage)


def build_flow_gap(machining: pd.DataFrame, post: pd.DataFrame, product_material: pd.DataFrame) -> pd.DataFrame:
    project_by_code = {}
    for _, row in product_material.iterrows():
        code = str(row.get("产品编码", "")).strip()
        project = str(row.get("产品/项目", "")).strip()
        if code and project:
            project_by_code.setdefault(code, project)
    m = machining.copy()
    m["date"] = m["报工时间"].map(date_str)
    m["product_code"] = m["物料编码"].astype(str).str.strip()
    m["product_name"] = m["物料名称"].astype(str).str.strip()
    m["machining_good_qty"] = m["合格品"].map(to_num)
    m["machining_defective_qty"] = m["不良品"].map(to_num)
    mg = m.groupby(["date", "product_code", "product_name"], as_index=False).agg(
        machining_good_qty=("machining_good_qty", "sum"),
        machining_defective_qty=("machining_defective_qty", "sum"),
    )
    p = post.copy()
    p["date"] = p["报工时间"].map(date_str)
    p.loc[p["date"] == "", "date"] = p.loc[p["date"] == "", "操作时间"].map(date_str)
    p["product_code"] = p["物资编码"].astype(str).str.strip()
    p["post_process_online_qty"] = p["上线数"].map(to_num)
    p["post_process_good_qty"] = p["合格数"].map(to_num)

    rows = []
    for _, row in mg.iterrows():
        date = row["date"]
        if not date:
            continue
        start = pd.Timestamp(date)
        end = start + pd.Timedelta(days=1)
        mask = (
            (p["product_code"] == row["product_code"])
            & (pd.to_datetime(p["date"], errors="coerce") >= start)
            & (pd.to_datetime(p["date"], errors="coerce") <= end)
        )
        online = float(p.loc[mask, "post_process_online_qty"].sum())
        good = float(p.loc[mask, "post_process_good_qty"].sum())
        flow_gap = float(row["machining_good_qty"]) - online
        if row["machining_good_qty"] > 0 and online == 0:
            status = "高风险"
            reason = "疑似后工序未承接；无统一流转单号，按产品+时间窗口近似判断"
        elif row["machining_good_qty"] > online:
            status = "中风险"
            reason = "疑似后工序积压；无统一流转单号，按产品+时间窗口近似判断"
        elif online > row["machining_good_qty"]:
            status = "口径异常或跨期流转"
            reason = "后工序上线数 > 机加合格品，可能存在跨期流转、前期在制或编码口径差异"
        else:
            status = "低风险"
            reason = "机加合格品与后工序上线数基本匹配；仍为按产品+时间窗口近似判断"
        rows.append(
            {
                "date": date,
                "project_code": project_by_code.get(row["product_code"], ""),
                "product_code": row["product_code"],
                "product_name": row["product_name"],
                "machining_good_qty": row["machining_good_qty"],
                "machining_defective_qty": row["machining_defective_qty"],
                "post_process_online_qty": online,
                "post_process_good_qty": good,
                "flow_gap_qty": flow_gap,
                "possible_wip_qty": max(0.0, flow_gap),
                "flow_status": status,
                "risk_reason": reason,
            }
        )
    return pd.DataFrame(rows)


def build_quality_loss(post: pd.DataFrame, defective: pd.DataFrame, machining: pd.DataFrame) -> pd.DataFrame:
    post_rows = post.copy()
    post_rows["date"] = post_rows["报工时间"].map(date_str)
    post_rows.loc[post_rows["date"] == "", "date"] = post_rows.loc[post_rows["date"] == "", "操作时间"].map(date_str)
    post_rows["project_code"] = post_rows["产品/项目"].astype(str).str.strip()
    post_rows["product_code"] = post_rows["物资编码"].astype(str).str.strip()
    post_rows["product_name"] = post_rows["物资名称"].astype(str).str.strip()
    post_rows["process_name"] = post_rows["工序"].astype(str).str.strip()
    post_rows["online_qty"] = post_rows["上线数"].map(to_num)
    post_rows["good_qty"] = post_rows["合格数"].map(to_num)
    post_rows["post_defective_qty"] = post_rows["不良品"].map(to_num)
    grouped = post_rows.groupby(["date", "project_code", "product_code", "product_name", "process_name"], as_index=False).agg(
        online_qty=("online_qty", "sum"),
        good_qty=("good_qty", "sum"),
        post_defective_qty=("post_defective_qty", "sum"),
    )

    m = machining.copy()
    m["date"] = m["报工时间"].map(date_str)
    m["product_code"] = m["物料编码"].astype(str).str.strip()
    m["machining_defective_qty"] = m["不良品"].map(to_num)
    mg = m.groupby(["date", "product_code"], as_index=False).agg(machining_defective_qty=("machining_defective_qty", "sum"))

    d = defective.copy()
    d["date"] = d["申报时间"].map(date_str)
    d["product_code"] = d["物料编码"].astype(str).str.strip()
    d["tracking_defective_qty"] = d["总量"].map(to_num)
    d["scrap_qty"] = d.apply(lambda r: to_num(r["总量"]) if "报废" in str(r.get("后序处置", "")) else 0.0, axis=1)
    d["repair_qty"] = d.apply(lambda r: to_num(r["总量"]) if "返修" in str(r.get("后序处置", "")) else 0.0, axis=1)
    d["disposal_status"] = d["后序处置"].astype(str).str.strip()
    dg = d.groupby(["date", "product_code"], as_index=False).agg(
        tracking_defective_qty=("tracking_defective_qty", "sum"),
        scrap_qty=("scrap_qty", "sum"),
        repair_qty=("repair_qty", "sum"),
        disposal_status=("disposal_status", aggregate_reason),
    )

    out = grouped.merge(mg, on=["date", "product_code"], how="left").merge(dg, on=["date", "product_code"], how="left")
    for col in ["machining_defective_qty", "tracking_defective_qty", "scrap_qty", "repair_qty"]:
        out[col] = out[col].fillna(0)
    out["disposal_status"] = out["disposal_status"].fillna("")
    out["defective_qty"] = out["post_defective_qty"] + out["machining_defective_qty"] + out["tracking_defective_qty"]
    out["defective_rate"] = out.apply(lambda r: round(r["defective_qty"] / r["online_qty"], 4) if r["online_qty"] else "", axis=1)
    out["scrap_rate"] = out.apply(lambda r: round(r["scrap_qty"] / r["online_qty"], 4) if r["online_qty"] else "", axis=1)

    risks = []
    reasons = []
    for _, row in out.iterrows():
        defective_rate = row["defective_rate"] if row["defective_rate"] != "" else 0
        if defective_rate > 0.10:
            risk = "高风险"
            reason = "不良率 > 10%"
        elif defective_rate > 0.05:
            risk = "中风险"
            reason = "不良率 > 5%"
        elif row["scrap_qty"] > 0:
            risk = "中风险"
            reason = "报废数量 > 0"
        elif row["defective_qty"] > 0:
            risk = "低风险"
            reason = "存在不良记录"
        else:
            risk = "无明显风险"
            reason = "未见不良数量"
        if row["scrap_qty"] > 0 or row["repair_qty"] > 0 or row["tracking_defective_qty"] > 0:
            reason += "；缺闭环证据，无法确认是否已退料或扣账"
        reasons.append(reason)
        risks.append(risk)
    out["quality_risk_level"] = risks
    out["risk_reason"] = reasons
    return out[
        [
            "date",
            "project_code",
            "product_code",
            "product_name",
            "process_name",
            "online_qty",
            "good_qty",
            "defective_qty",
            "scrap_qty",
            "repair_qty",
            "defective_rate",
            "scrap_rate",
            "quality_risk_level",
            "disposal_status",
            "risk_reason",
        ]
    ]


def downtime_by_equipment(downtime: pd.DataFrame) -> pd.DataFrame:
    d = downtime.copy()
    d["date"] = d["报停时间"].map(date_str)
    d["equipment_id"] = d["设备号"].astype(str).str.strip()
    d["equipment_name_dt"] = d["设备名称"].astype(str).str.strip()
    d["downtime_minutes"] = d.apply(parse_downtime_minutes, axis=1)
    return d.groupby(["date", "equipment_id"], as_index=False).agg(
        equipment_name_dt=("equipment_name_dt", first_nonempty),
        downtime_minutes=("downtime_minutes", "sum"),
    )


def parse_downtime_minutes(row) -> float:
    value = to_num(row.get("报停时长", ""))
    if value:
        return value * 60
    start = to_date(row.get("报停时间", ""))
    end = to_date(row.get("重启时间", ""))
    if pd.notna(start) and pd.notna(end) and end >= start:
        return (end - start).total_seconds() / 60
    return 0.0


def tooling_by_equipment(tooling: pd.DataFrame) -> pd.DataFrame:
    t = tooling.copy()
    t["date"] = t["起始时间"].map(date_str)
    t["equipment_id"] = t["设备号"].astype(str).str.strip()
    t["equipment_name_tool"] = t["设备名称"].astype(str).str.strip()
    t["tooling_change_minutes"] = t["用时H"].map(to_num) * 60
    t["trial_scrap_qty"] = t["试料数(报废)"].map(to_num)
    return t.groupby(["date", "equipment_id"], as_index=False).agg(
        equipment_name_tool=("equipment_name_tool", first_nonempty),
        tooling_change_minutes=("tooling_change_minutes", "sum"),
        trial_scrap_qty=("trial_scrap_qty", "sum"),
    )


def build_capacity_loss(takt: pd.DataFrame, equipment: pd.DataFrame, downtime: pd.DataFrame, tooling: pd.DataFrame) -> pd.DataFrame:
    cap = takt.copy()
    cap["equipment_id"] = cap["设备ID"].astype(str).str.strip()
    cap["equipment_name"] = cap["设备名称"].astype(str).str.strip()
    cap["product_code"] = cap["产品编码"].astype(str).str.strip()
    cap["product_name"] = cap["产品名称"].astype(str).str.strip()
    cap["takt_minutes"] = cap["总节拍【分】"].map(to_num)
    cap = cap[(cap["equipment_id"] != "") & (cap["product_code"] != "") & (cap["takt_minutes"] > 0)].copy()
    cap["hourly_capacity"] = 60 / cap["takt_minutes"]
    cap = cap[["equipment_id", "equipment_name", "product_code", "product_name", "takt_minutes", "hourly_capacity"]].drop_duplicates()

    eq = equipment.copy()
    if {"设备ID", "排程分类"}.issubset(eq.columns):
        schedulable = set(eq.loc[eq["排程分类"].astype(str).str.contains("排程", na=False), "设备ID"].astype(str).str.strip())
        cap = cap[cap["equipment_id"].isin(schedulable) | (len(schedulable) == 0)]

    dt = downtime_by_equipment(downtime)
    tg = tooling_by_equipment(tooling)
    events = dt.merge(tg, on=["date", "equipment_id"], how="outer")
    for col in ["downtime_minutes", "tooling_change_minutes", "trial_scrap_qty"]:
        events[col] = events[col].fillna(0)
    events["equipment_name_event"] = events.apply(
        lambda r: first_nonempty([r.get("equipment_name_dt", ""), r.get("equipment_name_tool", "")]), axis=1
    )
    rows = []
    for _, event in events.iterrows():
        candidates = cap[cap["equipment_id"] == event["equipment_id"]]
        if candidates.empty:
            rows.append(
                {
                    "date": event["date"],
                    "equipment_id": event["equipment_id"],
                    "equipment_name": event["equipment_name_event"],
                    "product_code": "",
                    "product_name": "",
                    "takt_minutes": "",
                    "hourly_capacity": "",
                    "downtime_minutes": event["downtime_minutes"],
                    "tooling_change_minutes": event["tooling_change_minutes"],
                    "trial_scrap_qty": event["trial_scrap_qty"],
                    "lost_capacity_by_downtime": "",
                    "lost_capacity_by_tooling": "",
                    "available_capacity_estimate": "",
                    "capacity_risk_level": "无法判断",
                    "risk_reason": "设备无节拍产品集合，无法估算；理论估算不可当成真实产量",
                }
            )
            continue
        for _, product in candidates.iterrows():
            hourly = float(product["hourly_capacity"])
            lost_dt = event["downtime_minutes"] / 60 * hourly
            lost_tool = event["tooling_change_minutes"] / 60 * hourly + event["trial_scrap_qty"]
            available = max(0.0, (8 - event["downtime_minutes"] / 60 - event["tooling_change_minutes"] / 60) * hourly)
            total_loss_minutes = event["downtime_minutes"] + event["tooling_change_minutes"]
            if total_loss_minutes >= 240:
                risk = "高风险"
            elif total_loss_minutes > 0:
                risk = "中风险"
            else:
                risk = "低风险"
            rows.append(
                {
                    "date": event["date"],
                    "equipment_id": event["equipment_id"],
                    "equipment_name": product["equipment_name"] or event["equipment_name_event"],
                    "product_code": product["product_code"],
                    "product_name": product["product_name"],
                    "takt_minutes": round(product["takt_minutes"], 4),
                    "hourly_capacity": round(hourly, 4),
                    "downtime_minutes": round(event["downtime_minutes"], 2),
                    "tooling_change_minutes": round(event["tooling_change_minutes"], 2),
                    "trial_scrap_qty": event["trial_scrap_qty"],
                    "lost_capacity_by_downtime": round(lost_dt, 2),
                    "lost_capacity_by_tooling": round(lost_tool, 2),
                    "available_capacity_estimate": round(available, 2),
                    "capacity_risk_level": risk,
                    "risk_reason": "理论估算；无设备当天具体生产任务，按设备可加工产品集合估算；默认8小时班次，缺真实可用工时口径",
                }
            )
    return pd.DataFrame(rows)


def build_summary(
    demand: pd.DataFrame,
    inventory: pd.DataFrame,
    material: pd.DataFrame,
    flow: pd.DataFrame,
    quality: pd.DataFrame,
    capacity: pd.DataFrame,
) -> pd.DataFrame:
    base = demand.groupby(["date", "project_code", "product_code", "product_name"], as_index=False).agg(
        daily_plan_qty=("daily_plan_qty", "sum"),
        weekly_plan_qty=("weekly_plan_qty", "sum"),
        completed_qty=("completed_qty", "sum"),
        completion_rate=("completion_rate", "max"),
    )
    base = base[base["daily_plan_qty"] > 0].copy()
    inv = inventory[["date", "product_code", "current_inventory", "future_3_days_demand", "inventory_risk_level", "risk_reason"]].rename(
        columns={"inventory_risk_level": "inventory_risk", "risk_reason": "inventory_reason"}
    )
    out = base.merge(inv, on=["date", "product_code"], how="left")

    mat = material.groupby(["date", "product_code"], as_index=False).agg(
        material_risk=("material_risk_level", max_risk),
        material_reason=("risk_reason", aggregate_reason),
    )
    out = out.merge(mat, on=["date", "product_code"], how="left")

    fl = flow.groupby(["date", "product_code"], as_index=False).agg(
        flow_risk=("flow_status", max_risk),
        flow_reason=("risk_reason", aggregate_reason),
    )
    out = out.merge(fl, on=["date", "product_code"], how="left")

    q = quality.groupby(["date", "product_code"], as_index=False).agg(
        quality_risk=("quality_risk_level", max_risk),
        quality_reason=("risk_reason", aggregate_reason),
    )
    out = out.merge(q, on=["date", "product_code"], how="left")

    c = capacity.groupby(["date", "product_code"], as_index=False).agg(
        capacity_risk=("capacity_risk_level", max_risk),
        capacity_reason=("risk_reason", aggregate_reason),
    )
    out = out.merge(c, on=["date", "product_code"], how="left")

    rows = []
    for _, row in out.iterrows():
        risks = {
            "inventory": row.get("inventory_risk", ""),
            "material": row.get("material_risk", ""),
            "flow": row.get("flow_risk", ""),
            "quality": row.get("quality_risk", ""),
            "capacity": row.get("capacity_risk", ""),
        }
        reasons = [
            row.get("inventory_reason", ""),
            row.get("material_reason", ""),
            row.get("flow_reason", ""),
            row.get("quality_reason", ""),
            row.get("capacity_reason", ""),
        ]
        main_risk = max_risk(risks.values())
        main_reason = aggregate_reason(reasons, limit=4)
        completion = to_num(row.get("completion_rate", ""))
        if completion >= 100:
            delivery_status = "周计划已完成"
        elif row["completed_qty"] > 0:
            delivery_status = f"周计划执行中，完成率 {completion:g}%"
        else:
            delivery_status = "周计划未见完成量"
        suggested = suggest_action(risks, main_risk)
        missing = missing_reason(row, risks)
        confidence = "low" if "无法判断" in risks.values() or row.get("capacity_risk", "") in {"中风险", "高风险"} else "medium"
        if row.get("inventory_risk", "") in {"低风险", "中风险", "高风险"} and row.get("material_risk", "") in {"低风险", "中风险", "高风险"}:
            confidence = "medium"
        rows.append(
            {
                "date": row["date"],
                "project_code": row["project_code"],
                "product_code": row["product_code"],
                "product_name": row["product_name"],
                "daily_plan_qty": row["daily_plan_qty"],
                "current_inventory": row.get("current_inventory", ""),
                "future_3_days_demand": row.get("future_3_days_demand", ""),
                "inventory_risk": row.get("inventory_risk", "无法判断"),
                "material_risk": row.get("material_risk", "无法判断"),
                "flow_risk": row.get("flow_risk", "无法判断"),
                "quality_risk": row.get("quality_risk", "无法判断"),
                "capacity_risk": row.get("capacity_risk", "无法判断"),
                "delivery_execution_status": delivery_status,
                "main_risk_reason": main_reason,
                "suggested_action": suggested,
                "data_confidence": confidence,
                "missing_data_reason": missing,
            }
        )
    return pd.DataFrame(rows).sort_values(["date", "data_confidence", "product_code"])


def suggest_action(risks: dict[str, str], main_risk: str) -> str:
    if risks.get("inventory") == "高风险":
        return "优先核对成品库存与近期入出库，确认是否需补产或调整交付"
    if risks.get("material") == "高风险":
        return "优先核对材料库存、产品材料映射和采购到料"
    if risks.get("flow") == "高风险":
        return "追踪机加后工序承接，确认是否存在在制积压"
    if risks.get("quality") in {"高风险", "中风险"}:
        return "确认不良/报废处置闭环，评估可交付数量"
    if risks.get("capacity") in {"高风险", "中风险"}:
        return "核对设备停机、工装更换和当天实际生产任务"
    if main_risk == "无法判断":
        return "补齐关联键或口径后重新判断"
    return "持续跟踪计划完成、库存和生产流转"


def missing_reason(row, risks: dict[str, str]) -> str:
    missing = []
    if risks.get("inventory") == "无法判断":
        missing.append("产品编码无法关联仓库台账")
    if risks.get("material") == "无法判断":
        missing.append("缺产品-材料映射或材料库存")
    if risks.get("flow") == "无法判断":
        missing.append("缺同日/次日机加与后工序可比记录或统一流转单号")
    if risks.get("quality") == "无法判断":
        missing.append("缺质量闭环或可比不良记录")
    if risks.get("capacity") == "无法判断":
        missing.append("缺设备当天任务或设备-工装适配")
    return "；".join(missing)


def write_report(paths: dict[str, Path], counts: dict[str, int], validation: dict[str, str], ledger_month: str) -> None:
    REPORT_ROOT.mkdir(parents=True, exist_ok=True)
    lines = [
        "# 07 第一版生产节奏分析输出总结",
        "",
        "## 1. 本次成功生成的分析表",
        "",
        md_table(
            ["输出文件", "行数", "说明"],
            [
                ["output/01_demand_daily.csv", counts["demand"], "周计划按日展开后的未来计划需求"],
                ["output/02_inventory_coverage.csv", counts["inventory"], "产品库存对未来3天/7天需求的覆盖判断"],
                ["output/03_material_coverage.csv", counts["material"], "产品计划展开到材料需求后的材料覆盖判断"],
                ["output/04_flow_gap_analysis.csv", counts["flow"], "机加转序到后工序承接的近似差异分析"],
                ["output/05_quality_loss_analysis.csv", counts["quality"], "后工序、机加和不良品跟踪合并后的质量损失分析"],
                ["output/06_capacity_loss_analysis.csv", counts["capacity"], "设备停机/工装更换对理论产能的估算影响"],
                ["output/00_production_rhythm_summary.csv", counts["summary"], "按产品-日期汇总的生产节奏总览"],
            ],
        ),
        "",
        "## 2. 每张分析表使用的数据源",
        "",
        md_table(
            ["分析表", "使用数据源"],
            [
                ["01_demand_daily", "weekly_plan_execution_tracking；project_info_maintenance；product_material_match_maintenance；production_takt_maintenance"],
                ["02_inventory_coverage", "01_demand_daily；warehouse_ledger"],
                ["03_material_coverage", "01_demand_daily；product_material_match_maintenance；warehouse_ledger"],
                ["04_flow_gap_analysis", "machining_transfer_query；post_process_report_tracking；product_material_match_maintenance"],
                ["05_quality_loss_analysis", "post_process_report_tracking；defective_tracking；machining_transfer_query"],
                ["06_capacity_loss_analysis", "production_takt_maintenance；production_equipment_maintenance；equipment_downtime_query；tooling_replacement_query"],
                ["00_production_rhythm_summary", "以上所有分析结果汇总"],
            ],
        ),
        "",
        "## 3. 字段来源",
        "",
        md_table(
            ["输出主题", "核心字段来源"],
            [
                ["日计划", "`年周`、`开始日期`、`星期日`~`星期六`、`计划量`、`完成量`、`完成率` 来自周计划执行跟踪"],
                ["库存覆盖", "`当前库存` 来自仓库台账；需求来自日计划展开结果"],
                ["材料需求", "`产品编码`、`领料物资ID`、`领料物资名称`、`单台用量` 来自产品材料匹配维护"],
                ["流转差异", "`合格品`、`不良品` 来自机加转序；`上线数`、`合格数` 来自后工序报工"],
                ["质量损失", "`不良品` 来自后工序和机加；`总量`、`后序处置` 来自不良品跟踪"],
                ["产能损失", "`总节拍【分】` 来自节拍维护；`报停时长` 来自设备停机；`用时H`、`试料数(报废)` 来自工装更换"],
            ],
        ),
        "",
        "## 4. 可信度较高的结果",
        "",
        "- 日计划展开可信度较高：`开始日期` 非空且校验结果为：" + validation["sunday_mapping_note"],
        f"- 库存覆盖在产品编码能匹配仓库台账时可信度较高，依据是仓库台账最新 `月度={ledger_month}` 的 `当前库存` 与周计划 `产品编码`。",
        "- 节拍的单位理论产能计算可信度较高，公式为 `60 / 总节拍【分】`；但它仍然只是理论产能。",
        "",
        "## 5. 近似判断",
        "",
        "- 机加到后工序流转使用“产品编码 + 当天至次日”窗口近似判断，因为没有统一流转单号。",
        "- 停机/工装更换的产能影响按设备可加工产品集合估算，因为没有设备当天具体生产任务。",
        "- 材料覆盖依赖产品材料匹配维护的 `单台用量`，未处理替代料、损耗和有效期。",
        "- 质量损失合并了多个来源的不良线索，存在重复统计可能，需后续用质量单号或流转单号闭环。",
        "",
        "## 6. 仍然做不了的判断",
        "",
        "- 不能把销售出库直接等同客户签收；本次未使用销售出库判断签收。",
        "- 不能确认不良、报废是否已经退料或扣账；输出中标记为“缺闭环证据”。",
        "- 不能得到真实可用产能；产能表只输出理论估算。",
        "- 不能精确判断批次级流转；缺统一流转单号。",
        "- 不能精确催料到采购订单；缺采购订单交期、未到货量和供应商承诺日期。",
        "",
        "## 7. 最优先补充的 5 类数据",
        "",
        "1. 设备-工装/夹具-产品适配关系。",
        "2. 采购订单、应到日期、未到货量、供应商承诺日期。",
        "3. 统一流转单号或批次号，用于连接机加、后工序、质量和库存。",
        "4. 不良/报废/返修闭环记录，包括退料、扣账和处置完成时间。",
        "5. 班次/工作日历和设备当天生产任务。",
        "",
        "## 8. 是否可以制作生产节奏主控看板",
        "",
        "可以开始做第一版主控看板，但建议看板直接读取本次 CSV 输出，不要先重构数据库。第一版看板应明确展示数据置信度和无法判断原因，尤其产能模块必须标注“理论估算”。",
        "",
        "## 9. 处理口径说明",
        "",
        f"- {validation['blank_plan_note']}",
        f"- 库存覆盖只使用仓库台账最新月份 `月度={ledger_month}`，不累加历史月份，也不代表未来入库。",
        "- 产能估算默认 8 小时班次，仅用于相对影响评估。",
    ]
    (REPORT_ROOT / "07_analysis_output_summary.md").write_text("\n".join(lines), encoding="utf-8")


def md_table(headers: list[str], rows: list[list[object]]) -> str:
    out = ["| " + " | ".join(headers) + " |", "| " + " | ".join(["---"] * len(headers)) + " |"]
    for row in rows:
        out.append("| " + " | ".join(str(x).replace("|", "\\|").replace("\n", "<br>") for x in row) + " |")
    return "\n".join(out)


def main() -> None:
    weekly = read_csv("weekly_plan_execution_tracking")
    project_info = read_csv("project_info_maintenance")
    product_material = read_csv("product_material_match_maintenance")
    takt = read_csv("production_takt_maintenance")
    ledger = read_csv("warehouse_ledger")
    machining = read_csv("machining_transfer_query")
    post = read_csv("post_process_report_tracking")
    defective = read_csv("defective_tracking")
    equipment = read_csv("production_equipment_maintenance")
    downtime = read_csv("equipment_downtime_query")
    tooling = read_csv("tooling_replacement_query")

    demand, validation = build_demand_daily(weekly, project_info, product_material, takt)
    inventory = build_inventory_coverage(demand, ledger)
    material = build_material_coverage(demand, product_material, ledger)
    flow = build_flow_gap(machining, post, product_material)
    quality = build_quality_loss(post, defective, machining)
    capacity = build_capacity_loss(takt, equipment, downtime, tooling)
    summary = build_summary(demand, inventory, material, flow, quality, capacity)

    paths = {
        "demand": write_output(demand, "01_demand_daily.csv"),
        "inventory": write_output(inventory, "02_inventory_coverage.csv"),
        "material": write_output(material, "03_material_coverage.csv"),
        "flow": write_output(flow, "04_flow_gap_analysis.csv"),
        "quality": write_output(quality, "05_quality_loss_analysis.csv"),
        "capacity": write_output(capacity, "06_capacity_loss_analysis.csv"),
        "summary": write_output(summary, "00_production_rhythm_summary.csv"),
    }
    counts = {
        "demand": len(demand),
        "inventory": len(inventory),
        "material": len(material),
        "flow": len(flow),
        "quality": len(quality),
        "capacity": len(capacity),
        "summary": len(summary),
    }
    write_report(paths, counts, validation, latest_ledger_month(ledger))


if __name__ == "__main__":
    main()
