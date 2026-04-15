#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import statistics
from collections import defaultdict
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from xlsx_fallback_reader import maybe_number, maybe_percent, read_result_rows


def key_from_filename(path: Path) -> str:
    stem = path.stem
    return stem.split("_", 1)[0] if "_" in stem else stem


def load_all_records(input_dir: Path) -> dict[str, list[dict[str, str]]]:
    out: dict[str, list[dict[str, str]]] = {}
    for path in sorted(input_dir.glob("*.xlsx")):
        out[key_from_filename(path)] = read_result_rows(path)
    return out


def parse_date(value: str) -> str | None:
    if not value:
        return None
    s = value.strip()
    if "T" in s:
        s = s.split("T", 1)[0]
    for fmt in ("%Y-%m-%d", "%Y-%m", "%Y/%m/%d", "%Y/%m"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d" if fmt in ("%Y-%m-%d", "%Y/%m/%d") else "%Y-%m")
        except ValueError:
            continue
    return s


def fmt_int(value: float | int | None) -> str:
    if value is None:
        return "-"
    return f"{int(round(value)):,}"


def fmt_pct(value: float | None) -> str:
    if value is None:
        return "-"
    return f"{value * 100:.1f}%"


def normalize_period(row: dict[str, str]) -> str | None:
    for key in ("日期", "周起始日", "周起", "周开始(周一)", "week_start", "月份", "月"):
        if key in row and row[key]:
            return parse_date(row[key])
    return None


def pick_value_field(row: dict[str, str]) -> str | None:
    candidates = [
        "访问人数",
        "任务布置数",
        "任务数",
        "访问次数",
        "数量",
        "过评项创建数",
        "过评项发布数",
        "过评分录入数",
        "行政班考勤班次",
        "已考勤课节数",
        "德育分数录入数",
        "调代课次数",
        "学期评语数",
        "校历事件数",
        "评教提交次数",
        "选课志愿数",
        "考试创建数",
        "场地预约发起数",
        "all_count",
        "人数",
        "学生人数",
    ]
    for name in candidates:
        if name in row and maybe_number(row[name]) is not None:
            return name
    for key, value in row.items():
        if key in {"角色", "类型", "name", "项目", "状态", "学校", "学部", "年级名称", "届别", "评价项名称"}:
            continue
        if maybe_number(value) is not None:
            return key
    return None


def pick_category_field(row: dict[str, str]) -> str | None:
    for name in ("角色", "类型", "name", "项目", "状态", "学部", "年级名称", "学期名称", "评价项名称"):
        if name in row and row[name]:
            return name
    return None


def aggregate_series(records: list[dict[str, str]]) -> dict[str, Any]:
    period_totals: defaultdict[str, float] = defaultdict(float)
    category_totals: defaultdict[str, float] = defaultdict(float)

    for row in records:
        period = normalize_period(row)
        value_field = pick_value_field(row)
        if not period or not value_field:
            continue
        value = maybe_number(row.get(value_field))
        if value is None:
            continue
        period_totals[period] += value
        category_field = pick_category_field(row)
        if category_field:
            category = row.get(category_field, "").strip() or "未分类"
            category_totals[category] += value

    periods = sorted(period_totals)
    series = [{"period": p, "value": period_totals[p]} for p in periods]
    values = [item["value"] for item in series]
    total = sum(values)
    mean = statistics.mean(values) if values else None
    stdev = statistics.pstdev(values) if len(values) >= 2 else None
    volatility = (stdev / mean) if stdev is not None and mean not in (None, 0) else None
    peak = max(series, key=lambda item: item["value"]) if series else None
    latest = series[-1] if series else None
    top_category = None
    top_category_share = None
    if category_totals:
        name, val = max(category_totals.items(), key=lambda item: item[1])
        top_category = {"name": name, "value": val}
        if total:
            top_category_share = val / total

    return {
        "series": series,
        "total_value": total,
        "mean_value": mean,
        "volatility_ratio": volatility,
        "peak_period": peak["period"] if peak else None,
        "peak_value": peak["value"] if peak else None,
        "latest_period": latest["period"] if latest else None,
        "latest_value": latest["value"] if latest else None,
        "top_category": top_category,
        "top_category_share": top_category_share,
        "category_totals": dict(sorted(category_totals.items(), key=lambda item: item[1], reverse=True)),
    }


def build_overview(records_map: dict[str, list[dict[str, str]]]) -> dict[str, Any]:
    user_rows = records_map.get("用户信息", [])
    grade_rows = records_map.get("每年级学生人数", [])
    week_rows = records_map.get("周活（近", [])
    daily_rows = records_map.get("日活（近", [])

    role_summary = []
    total_users = 0.0
    for row in user_rows:
        count = maybe_number(row.get("人数")) or 0
        total_users += count
        role_summary.append(
            {
                "role": row.get("角色", ""),
                "count": count,
                "wechat_rate": maybe_percent(row.get("微信绑定率")),
                "mobile_rate": maybe_percent(row.get("手机绑定率")),
                "dingtalk_rate": maybe_percent(row.get("钉钉绑定率")),
            }
        )

    division_counts: defaultdict[str, float] = defaultdict(float)
    grade_table = []
    for row in grade_rows:
        count = maybe_number(row.get("学生人数")) or 0
        division = row.get("学部", "").strip() or "未标注学部"
        division_counts[division] += count
        grade_table.append({"division": division, "grade": row.get("年级名称", ""), "students": count})

    schools = sorted({row.get("学校", "").strip() for row in week_rows if row.get("学校", "").strip()})
    school_name = schools[0] if len(schools) == 1 else None
    school_name_confidence = "high" if school_name else ("mixed" if schools else "unknown")

    dates = []
    for row in week_rows + daily_rows:
        period = normalize_period(row)
        if period:
            dates.append(period)

    return {
        "school_name": school_name,
        "school_name_candidates": schools,
        "school_name_confidence": school_name_confidence,
        "coverage_start": min(dates) if dates else None,
        "coverage_end": max(dates) if dates else None,
        "total_users": total_users,
        "roles": role_summary,
        "divisions": dict(sorted(division_counts.items(), key=lambda item: item[1], reverse=True)),
        "grades": grade_table,
    }


def create_chart_payload(module_title: str, summary: dict[str, Any]) -> dict[str, Any] | None:
    if len(summary.get("series", [])) >= 2:
        return {
            "kind": "line",
            "title": module_title,
            "labels": [item["period"] for item in summary["series"]],
            "values": [float(item["value"]) for item in summary["series"]],
        }
    if len(summary.get("category_totals", {})) >= 1:
        items = list(summary["category_totals"].items())[:8]
        return {
            "kind": "bar",
            "title": module_title,
            "labels": [str(k) for k, _ in items],
            "values": [float(v) for _, v in items],
        }
    return None


def build_module(key: str, title: str, records: list[dict[str, str]]) -> dict[str, Any]:
    summary = aggregate_series(records)
    chart_payload = create_chart_payload(title, summary)
    return {
        "key": key,
        "title": title,
        "record_count": len(records),
        "chart_payload": chart_payload,
        "chart_status": "available" if chart_payload else "data_only",
        **summary,
    }


def module_priority(module: dict[str, Any]) -> float:
    return float(module.get("total_value") or 0) + float(module.get("peak_value") or 0) + float(module.get("latest_value") or 0)


def build_dynamic_sections(overview: dict[str, Any], modules: list[dict[str, Any]]) -> list[dict[str, Any]]:
    sections: list[dict[str, Any]] = []
    if overview.get("roles") or overview.get("grades"):
        sections.append(
            {"id": "overview", "title": "学校概览与用户基础", "intro": "展示平台用户基础、绑定情况与学生结构，用于界定后续业务分析的起点。", "module_keys": [], "kind": "overview", "priority": float(overview.get("total_users") or 0)}
        )
    groups = [
        ("活跃度与使用黏性", "观察平台总体活跃水平、关键周期峰值和近期使用温度。", {"周活（近", "日活（近"}),
        ("课程与教学执行", "观察教学流程是否稳定嵌入平台，包括课程任务、排课、课表与选课相关场景。", {"每周课程任务布置数（近", "每周课程任务提交与批阅数（近", "每周课表查询次数（近一年）", "每周排课操作数（近一年）", "每月选课志愿数（近", "每月学期评语采集数（近一年）"}),
        ("成绩、评价与反馈", "观察评价项、过评分录入、评教与考试等反馈类业务的使用表现。", {"每周过评项创建数（近", "每周过评项发布数（近", "每周过评分录入数（近", "近半年评价项名称抽样", "每月评教提交人次（近", "考试创建数"}),
        ("考勤、请假与日常管理", "观察考勤、请假和学生日常管理是否形成连续使用。", {"每周行政班考勤班次（近一年）", "每周已考勤课节数（近", "每周学生", "每周调代课次数", "每周德育分数录入数（近"}),
        ("通知、校历与家校协同", "观察通知、校历等协同业务是否承担统一沟通入口作用。", {"每周通知发送量（近一年）", "每月校历事件数（近一年）"}),
        ("其他高频业务", "承接仍有分析价值但不属于主业务簇的模块。", {"场馆预约次数"}),
    ]
    assigned = set()
    for title, intro, keys in groups:
        matched = [m for m in modules if m["key"] in keys]
        if not matched:
            continue
        matched.sort(key=module_priority, reverse=True)
        sections.append({"id": title, "title": title, "intro": intro, "module_keys": [m["key"] for m in matched], "kind": "modules", "priority": sum(module_priority(m) for m in matched)})
        assigned.update(m["key"] for m in matched)
    remainder = [m for m in modules if m["key"] not in assigned]
    if remainder:
        remainder.sort(key=module_priority, reverse=True)
        sections.append({"id": "custom", "title": "补充业务观察", "intro": "以下模块虽不属于主业务簇，但对理解学校系统使用特征仍有补充价值。", "module_keys": [m["key"] for m in remainder], "kind": "modules", "priority": sum(module_priority(m) for m in remainder)})
    return sections[:1] + sorted(sections[1:], key=lambda item: item["priority"], reverse=True)


def detect_issues(overview: dict[str, Any], modules: list[dict[str, Any]]) -> list[dict[str, str]]:
    issues: list[dict[str, str]] = []
    for role in overview.get("roles", []):
        mobile = role.get("mobile_rate")
        if mobile is not None and mobile < 0.5 and role.get("role") in {"学生", "家长"}:
            issues.append({"title": f"{role['role']}端基础触达能力偏弱", "phenomenon": f"{role['role']}手机绑定率仅为 {fmt_pct(mobile)}。", "evidence": "用户信息数据显示关键终端角色的移动端绑定率偏低。", "cause": "可能说明账号激活和家校触达入口尚未形成稳定覆盖。", "implication": "通知、成绩查看和家校沟通等高频场景的实际触达效率可能受限。"})
    for module in modules:
        vol = module.get("volatility_ratio")
        latest = module.get("latest_value")
        peak = module.get("peak_value")
        share = module.get("top_category_share")
        if vol is not None and vol > 0.60:
            issues.append({"title": f"{module['title']}使用节奏波动较大", "phenomenon": f"{module['title']}的离散程度较高，波动系数约为 {vol:.2f}。", "evidence": f"峰值 {fmt_int(peak)}，最近周期 {fmt_int(latest)}。", "cause": "结合校内业务节奏判断，相关使用可能高度依赖开学、考试或集中处理窗口。", "implication": "若缺少常态化机制，该模块会呈现节点型冲高、平时偏弱的断层特征。"})
        if peak and latest is not None and latest < peak * 0.5 and len(module.get("series", [])) >= 6:
            issues.append({"title": f"{module['title']}近期活跃度较峰值明显回落", "phenomenon": f"最近周期仅为峰值的 {latest / peak:.0%}。", "evidence": f"峰值 {fmt_int(peak)}，最近周期 {fmt_int(latest)}。", "cause": "可能反映该业务更依赖阶段性场景，而没有稳定嵌入日常管理。", "implication": "需要判断其属于正常季节性特征，还是流程执行不足。"})
        if share is not None and share >= 0.70 and module.get("top_category"):
            issues.append({"title": f"{module['title']}结构集中度偏高", "phenomenon": f"主要类型“{module['top_category']['name']}”占比约 {fmt_pct(share)}。", "evidence": "类别分布高度集中，其他场景贡献有限。", "cause": "可能说明系统主要承接单一高频场景，多元业务覆盖尚未形成。", "implication": "业务覆盖面的区分度不足，平台价值易被局限在局部流程。"})
    deduped = []
    seen = set()
    for issue in issues:
        if issue["title"] in seen:
            continue
        seen.add(issue["title"])
        deduped.append(issue)
    return deduped[:8]


def build_actions(issues: list[dict[str, str]]) -> list[dict[str, str]]:
    actions = []
    for issue in issues:
        title = issue["title"]
        if "触达" in title or "绑定率" in issue["phenomenon"]:
            actions.append({"owner": "校级管理 / 班主任", "issue": title, "action": "将学生和家长端绑定完成率纳入开学前两周基础任务，并按班级追踪未绑定名单。"})
        elif "波动" in title or "回落" in title:
            actions.append({"owner": "业务负责人", "issue": title, "action": "围绕该模块建立月度或双周固定使用节点，减少仅在峰值场景集中使用的节奏断层。"})
        elif "集中度偏高" in title:
            actions.append({"owner": "教务 / 信息化", "issue": title, "action": "识别当前单一高频场景外可迁移的两到三个流程，扩展系统的横向应用覆盖。"})
        else:
            actions.append({"owner": "信息化支持", "issue": title, "action": "补充周期复盘和使用台账，核查问题来自业务季节性，还是推广和流程设计存在缺口。"})
    return actions[:6]


def build_summary_bullets(overview: dict[str, Any], modules: list[dict[str, Any]], issues: list[dict[str, str]]) -> list[str]:
    bullets = []
    if overview.get("total_users"):
        bullets.append(f"平台当前覆盖约 {fmt_int(overview['total_users'])} 个角色账号，已形成教师、学生、家长三端基础用户池。")
    week = next((m for m in modules if m["key"] == "周活（近"), None)
    if week and week.get("peak_value") is not None:
        bullets.append(f"周活峰值出现在 {week['peak_period']}，达到 {fmt_int(week['peak_value'])}，说明平台在关键学期节点具备较强集中使用能力。")
    dominant = max((m for m in modules if m.get("total_value")), key=lambda item: item["total_value"], default=None)
    if dominant:
        bullets.append(f"{dominant['title']}是当前最强使用场景之一，累计量达到 {fmt_int(dominant['total_value'])}，可视为平台嵌入业务流程的重要抓手。")
    if issues:
        bullets.append(f"当前最主要的管理挑战集中在“{issues[0]['title']}”，需要通过组织推进与流程固化同步处理。")
    return bullets[:6]


def build_metric_cards(overview: dict[str, Any], modules: list[dict[str, Any]]) -> list[dict[str, str]]:
    cards = [{"label": "覆盖账号数", "value": fmt_int(overview.get("total_users"))}]
    if overview.get("coverage_end"):
        cards.append({"label": "数据截止", "value": overview["coverage_end"]})
    week = next((m for m in modules if m["key"] == "周活（近"), None)
    if week and week.get("peak_value") is not None:
        cards.append({"label": "周活峰值", "value": fmt_int(week["peak_value"])})
    if week and week.get("latest_value") is not None:
        cards.append({"label": "最近周活", "value": fmt_int(week["latest_value"])})
    return cards[:4]


def build_modules(records_map: dict[str, list[dict[str, str]]]) -> list[dict[str, Any]]:
    module_defs = [
        ("周活（近", "周活趋势"),
        ("日活（近", "日活趋势"),
        ("每周课程任务布置数（近", "课程任务布置"),
        ("每周课程任务提交与批阅数（近", "课程任务提交与批阅"),
        ("每周课表查询次数（近一年）", "课表查询"),
        ("每周排课操作数（近一年）", "排课操作"),
        ("每周过评项创建数（近", "过评项创建"),
        ("每周过评项发布数（近", "过评项发布"),
        ("每周过评分录入数（近", "过评分录入"),
        ("近半年评价项名称抽样", "评价项名称抽样"),
        ("每周行政班考勤班次（近一年）", "行政班考勤"),
        ("每周已考勤课节数（近", "已考勤课节"),
        ("每周学生", "学生与教师请假"),
        ("每周通知发送量（近一年）", "通知发送"),
        ("每月校历事件数（近一年）", "校历事件"),
        ("每月选课志愿数（近", "选课志愿"),
        ("每月学期评语采集数（近一年）", "学期评语采集"),
        ("每月评教提交人次（近", "评教提交"),
        ("场馆预约次数", "场馆预约"),
        ("考试创建数", "考试创建"),
    ]
    modules = []
    for key, title in module_defs:
        records = records_map.get(key)
        if not records:
            continue
        modules.append(build_module(key, title, records))
    return modules


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--workspace-root", required=True)
    parser.add_argument("--input-dir", required=True)
    parser.add_argument("--artifacts-dir", required=True)
    parser.add_argument("--output-json", required=True)
    parser.add_argument("--school-name")
    args = parser.parse_args()

    workspace_root = Path(args.workspace_root).resolve()
    input_dir = Path(args.input_dir).resolve()
    artifacts_dir = Path(args.artifacts_dir).resolve()
    artifacts_dir.mkdir(parents=True, exist_ok=True)

    records_map = load_all_records(input_dir)
    overview = build_overview(records_map)
    if args.school_name:
        overview["school_name"] = args.school_name
        overview["school_name_confidence"] = "manual"

    modules = build_modules(records_map)
    issues = detect_issues(overview, modules)
    actions = build_actions(issues)
    summary_bullets = build_summary_bullets(overview, modules, issues)
    metric_cards = build_metric_cards(overview, modules)
    dynamic_sections = build_dynamic_sections(overview, modules)

    result = {
        "workspace_root": str(workspace_root),
        "generated_at": datetime.now(UTC).isoformat(),
        "title": f"{overview.get('school_name') or 'XX学校'}希悦系统使用情况说明",
        "overview": overview,
        "metric_cards": metric_cards,
        "summary_bullets": summary_bullets,
        "modules": modules,
        "dynamic_sections": dynamic_sections,
        "issues": issues,
        "actions": actions,
        "data_note": "图表由 bundle 内置的 HTML canvas 图表流程根据 Excel 直接生成；若个别模块不适合图表展示，则以表格和文字分析为主。",
    }
    Path(args.output_json).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"output_json": str(Path(args.output_json).resolve()), "module_count": len(modules), "issue_count": len(issues)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
