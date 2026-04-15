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
            normalized = datetime.strptime(s, fmt)
            if fmt in ("%Y-%m-%d", "%Y/%m/%d"):
                return normalized.strftime("%Y-%m-%d")
            return normalized.strftime("%Y-%m")
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
        raw = row.get(key)
        if raw:
            return parse_date(raw)
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
        if row.get(name):
            return name
    return None


def aggregate_records(records: list[dict[str, str]]) -> dict[str, Any]:
    period_totals: defaultdict[str, float] = defaultdict(float)
    category_totals: defaultdict[str, float] = defaultdict(float)
    overall_total = 0.0

    for row in records:
        value_field = pick_value_field(row)
        if not value_field:
            continue
        value = maybe_number(row.get(value_field))
        if value is None:
            continue
        overall_total += value
        period = normalize_period(row)
        if period:
            period_totals[period] += value
        category_field = pick_category_field(row)
        if category_field:
            category = row.get(category_field, "").strip() or "未分类"
            category_totals[category] += value

    periods = sorted(period_totals)
    series = [{"period": p, "value": period_totals[p]} for p in periods]
    series_values = [item["value"] for item in series]
    mean = statistics.mean(series_values) if series_values else None
    stdev = statistics.pstdev(series_values) if len(series_values) >= 2 else None
    volatility = (stdev / mean) if stdev is not None and mean not in (None, 0) else None
    peak = max(series, key=lambda item: item["value"]) if series else None
    latest = series[-1] if series else None

    top_category = None
    top_category_share = None
    if category_totals:
        name, val = max(category_totals.items(), key=lambda item: item[1])
        top_category = {"name": name, "value": val}
        if overall_total:
            top_category_share = val / overall_total

    return {
        "series": series,
        "total_value": overall_total,
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


def chart_payload_from_summary(module_title: str, summary: dict[str, Any]) -> dict[str, Any] | None:
    series = summary.get("series", [])
    categories = summary.get("category_totals", {})
    if len(series) >= 2:
        return {
            "kind": "line",
            "title": module_title,
            "series_name": "趋势值",
            "labels": [item["period"] for item in series],
            "values": [float(item["value"]) for item in series],
            "legend": [{"label": "趋势值", "color": "#2f80ed"}],
        }
    if categories:
        items = list(categories.items())[:8]
        colors = ["#2f80ed", "#00b8d9", "#5b8ff9", "#36cfc9", "#fa8c16", "#f6bd16", "#7256ff", "#eb2f96"]
        return {
            "kind": "bar",
            "title": module_title,
            "series_name": "分类汇总",
            "labels": [str(k) for k, _ in items],
            "values": [float(v) for _, v in items],
            "legend": [{"label": label, "color": colors[idx % len(colors)]} for idx, (label, _) in enumerate(items)],
        }
    return None


def module_intro(module: dict[str, Any]) -> str:
    if module.get("series"):
        latest_text = f"最近周期为 {module['latest_period']}，当前值 {fmt_int(module.get('latest_value'))}" if module.get("latest_period") else "具备最近周期观测"
        return f"本模块用于观察“{module['title']}”在校内业务中的时间演变与近期状态，重点识别使用节奏、峰值表现与连续性特征；{latest_text}。"
    if module.get("top_category"):
        return f"本模块用于观察“{module['title']}”的结构分布，重点识别当前最主要的应用类型、集中度水平以及业务覆盖的均衡性。"
    return f"本模块用于补充观察“{module['title']}”的总体使用表现，用于判断该业务是否已被学校稳定纳入系统流程。"


def build_module_issue(module: dict[str, Any]) -> dict[str, str]:
    vol = module.get("volatility_ratio")
    latest = module.get("latest_value")
    peak = module.get("peak_value")
    share = module.get("top_category_share")
    top_category = module.get("top_category")

    if vol is not None and vol > 0.60:
        return {
            "title": "节奏波动较大",
            "evidence": f"时间序列离散程度约为 {vol:.2f}，峰值出现在 {module.get('peak_period') or '历史观测期'}，峰值为 {fmt_int(peak)}。",
            "cause": "结合学校业务节奏判断，该模块更接近节点驱动型使用，常态化运行机制可能仍偏弱。",
            "implication": "若没有固定使用要求或稳定复盘机制，业务执行容易在关键节点后迅速回落。",
        }
    if peak and latest is not None and peak > 0 and latest < peak * 0.5 and len(module.get("series", [])) >= 6:
        return {
            "title": "近期水平较历史峰值明显回落",
            "evidence": f"最近周期为 {module.get('latest_period') or '-'}，当前值 {fmt_int(latest)}，仅为历史峰值 {fmt_int(peak)} 的 {latest / peak:.0%}。",
            "cause": "可能说明该业务在高峰节点后缺乏持续推动，尚未形成连续、稳定的流程嵌入。",
            "implication": "管理层需要区分这是正常季节性变化，还是流程执行断层导致的活跃回落。",
        }
    if share is not None and share >= 0.70 and top_category:
        return {
            "title": "结构集中度偏高",
            "evidence": f"当前以“{top_category['name']}”为主，占比约 {fmt_pct(share)}，其余类型贡献相对有限。",
            "cause": "系统价值目前更集中在单一高频场景，横向扩展到更多业务环节的区分度仍不充分。",
            "implication": "若长期维持单点使用，平台对学校综合治理的支撑深度会受到限制。",
        }
    if module.get("chart_status") == "data_only":
        return {
            "title": "可视化证据相对有限",
            "evidence": f"当前模块可用记录 {module.get('record_count', 0)} 条，主要支持结构汇总和聚合判断。",
            "cause": "数据表更偏向静态结构或单期汇总，不足以形成稳定趋势观测。",
            "implication": "对该模块的判断应更多基于结构分布和累计表现，不宜过度推断趋势性结论。",
        }
    return {
        "title": "当前未见强烈异常",
        "evidence": "峰值、近期值和结构占比之间未出现明显失衡。",
        "cause": "该模块目前更接近常态化运行状态，业务节奏与平台使用之间具有一定一致性。",
        "implication": "后续应继续观察关键学期节点，以验证其稳定性是否具有持续信度。",
    }


def module_findings(module: dict[str, Any]) -> list[str]:
    findings: list[str] = []
    if module.get("peak_value") is not None:
        findings.append(f"峰值周期为 {module.get('peak_period')}，峰值达到 {fmt_int(module.get('peak_value'))}。")
    if module.get("latest_value") is not None:
        findings.append(f"最近周期为 {module.get('latest_period')}，当前水平为 {fmt_int(module.get('latest_value'))}。")
    if module.get("top_category"):
        share = module.get("top_category_share")
        share_text = f"，占比约 {fmt_pct(share)}" if isinstance(share, (int, float)) else ""
        findings.append(f"结构上以“{module['top_category']['name']}”为主{share_text}。")
    if module.get("volatility_ratio") is not None:
        findings.append(f"时间序列离散程度约为 {module['volatility_ratio']:.2f}，可作为稳定性判断的辅助证据。")
    if module.get("total_value"):
        findings.append(f"观测期内累计业务量约为 {fmt_int(module.get('total_value'))}。")
    return findings[:4]


def module_interpretation(module: dict[str, Any], issue: dict[str, str]) -> str:
    if issue["title"] == "节奏波动较大":
        return "从趋势表现看，该模块并非持续、平缓地运行，而是在少数时点集中抬升。这种波动通常意味着业务更多依赖阶段性工作安排，而不是被嵌入日常操作链路。"
    if issue["title"] == "近期水平较历史峰值明显回落":
        return "从常模参照角度看，当前水平相较校内历史高点明显偏低，说明近期使用热度不足。若这一回落并非由自然学期节奏解释，则可能对应执行频率下降或组织要求减弱。"
    if issue["title"] == "结构集中度偏高":
        return "从结构表现看，业务使用高度集中于少数类型，说明平台的应用覆盖仍偏窄。当前结构对单点业务支撑较强，但跨场景扩展的区分度尚未完全显现。"
    if issue["title"] == "可视化证据相对有限":
        return "由于该模块更偏向静态结构或单期汇总，其分析重点应放在覆盖面和分布格局，而不应过度引申为长期趋势判断。"
    return "该模块整体运行相对平稳，说明其已经形成一定的日常使用惯性。后续更应关注关键节点前后的波动是否持续可复现，以提高判断的稳定性。"


def build_module(key: str, title: str, records: list[dict[str, str]]) -> dict[str, Any]:
    summary = aggregate_records(records)
    chart_payload = chart_payload_from_summary(title, summary)
    module = {
        "key": key,
        "title": title,
        "record_count": len(records),
        "chart_payload": chart_payload,
        "chart_status": "available" if chart_payload else "data_only",
        **summary,
    }
    issue = build_module_issue(module)
    module["intro"] = module_intro(module)
    module["findings"] = module_findings(module)
    module["issue_analysis"] = issue
    module["interpretation"] = module_interpretation(module, issue)
    return module


def module_priority(module: dict[str, Any]) -> float:
    return float(module.get("total_value") or 0) + float(module.get("peak_value") or 0) + float(module.get("latest_value") or 0)


def build_dynamic_sections(overview: dict[str, Any], modules: list[dict[str, Any]]) -> list[dict[str, Any]]:
    sections: list[dict[str, Any]] = []
    if overview.get("roles") or overview.get("grades"):
        sections.append(
            {
                "id": "overview",
                "title": "学校概览与用户基础",
                "intro": "展示平台用户基础、终端绑定情况与学生结构，用于界定后续业务分析的参照范围。",
                "module_keys": [],
                "kind": "overview",
                "priority": float(overview.get("total_users") or 0),
            }
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
        sections.append(
            {
                "id": title,
                "title": title,
                "intro": intro,
                "module_keys": [m["key"] for m in matched],
                "kind": "modules",
                "priority": sum(module_priority(m) for m in matched),
            }
        )
        assigned.update(m["key"] for m in matched)
    remainder = [m for m in modules if m["key"] not in assigned]
    if remainder:
        remainder.sort(key=module_priority, reverse=True)
        sections.append(
            {
                "id": "custom",
                "title": "补充业务观察",
                "intro": "以下模块虽不属于主业务簇，但对理解学校系统使用特征仍有补充价值。",
                "module_keys": [m["key"] for m in remainder],
                "kind": "modules",
                "priority": sum(module_priority(m) for m in remainder),
            }
        )
    return sections[:1] + sorted(sections[1:], key=lambda item: item["priority"], reverse=True)


def build_summary_bullets(overview: dict[str, Any], modules: list[dict[str, Any]]) -> list[str]:
    bullets: list[str] = []
    if overview.get("total_users"):
        bullets.append(f"平台当前覆盖约 {fmt_int(overview['total_users'])} 个角色账号，已形成教师、学生、家长三端基础用户池。")
    week = next((m for m in modules if m["key"] == "周活（近"), None)
    if week and week.get("peak_value") is not None:
        bullets.append(f"周活峰值出现在 {week['peak_period']}，达到 {fmt_int(week['peak_value'])}，说明平台在关键学期节点具备较强集中使用能力。")
    dominant = max((m for m in modules if m.get("total_value")), key=lambda item: item["total_value"], default=None)
    if dominant:
        bullets.append(f"{dominant['title']}是当前最强使用场景之一，累计量达到 {fmt_int(dominant['total_value'])}，可视为平台嵌入业务流程的重要抓手。")
    module_with_issue = next((m for m in modules if m.get("issue_analysis", {}).get("title") not in {"当前未见强烈异常", "可视化证据相对有限"}), None)
    if module_with_issue:
        bullets.append(f"当前最值得关注的问题集中在“{module_with_issue['title']}”模块，其表现为{module_with_issue['issue_analysis']['title']}。")
    if overview.get("school_name_confidence") != "high":
        bullets.append("学校名称识别存在不确定性，正式交付前应结合原始导出文件再次校核。")
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
        ("每周调代课次数", "调代课"),
        ("每周德育分数录入数（近", "德育分数录入"),
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
    summary_bullets = build_summary_bullets(overview, modules)
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
    }
    Path(args.output_json).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"output_json": str(Path(args.output_json).resolve()), "module_count": len(modules)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
