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


TIME_KEYS = ("日期", "周起始日", "周起", "周开始(周一)", "week_start", "月份", "月")
CATEGORY_KEYS = ("角色", "类型", "name", "项目", "状态", "学部", "年级名称", "学期名称", "评价项名称", "评语任务", "考场状态", "type")
VALUE_CANDIDATES = (
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
    "is_important_count",
    "wide_range_count",
    "is_var_count",
    "人数",
    "学生人数",
    "snapshot_count",
)


def key_from_filename(path: Path) -> str:
    stem = path.stem
    return stem.split("_", 1)[0] if "_" in stem else stem


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


def load_all_records(input_dir: Path) -> dict[str, dict[str, Any]]:
    out: dict[str, dict[str, Any]] = {}
    for path in sorted(input_dir.glob("*.xlsx")):
        key = key_from_filename(path)
        out[key] = {
            "path": str(path),
            "file_name": path.name,
            "source_name": key,
            "records": read_result_rows(path),
        }
    return out


def normalize_period(row: dict[str, str]) -> str | None:
    for key in TIME_KEYS:
        raw = row.get(key)
        if raw:
            return parse_date(raw)
    return None


def pick_value_field(row: dict[str, str]) -> str | None:
    for name in VALUE_CANDIDATES:
        if name in row and maybe_number(row[name]) is not None:
            return name
    for key, value in row.items():
        if key in CATEGORY_KEYS or key in TIME_KEYS or key in {"学校", "周止", "周结束(周日)", "开始日期", "结束日期", "届别", "当月最早访问", "当月最近访问", "year", "week", "年周"}:
            continue
        if maybe_number(value) is not None:
            return key
    return None


def pick_category_field(row: dict[str, str]) -> str | None:
    for name in CATEGORY_KEYS:
        if row.get(name):
            return name
    return None


def aggregate_records(records: list[dict[str, str]]) -> dict[str, Any]:
    period_totals: defaultdict[str, float] = defaultdict(float)
    category_totals: defaultdict[str, float] = defaultdict(float)
    overall_total = 0.0
    value_field: str | None = None
    category_field: str | None = None
    has_time_series = False

    for row in records:
        row_value_field = pick_value_field(row)
        if not row_value_field:
            continue
        value = maybe_number(row.get(row_value_field))
        if value is None:
            continue
        value_field = value_field or row_value_field
        overall_total += value
        period = normalize_period(row)
        if period:
            has_time_series = True
            period_totals[period] += value
        row_category_field = pick_category_field(row)
        if row_category_field:
            category_field = category_field or row_category_field
            category = row.get(row_category_field, "").strip() or "未分类"
            category_totals[category] += value

    periods = sorted(period_totals)
    series = [{"period": p, "value": period_totals[p]} for p in periods]
    series_values = [item["value"] for item in series]
    mean = statistics.mean(series_values) if series_values else None
    stdev = statistics.pstdev(series_values) if len(series_values) >= 2 else None
    volatility = (stdev / mean) if stdev is not None and mean not in (None, 0) and len(series_values) >= 4 else None
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
        "value_field": value_field,
        "category_field": category_field,
        "has_time_series": has_time_series,
    }


def infer_school_candidates(dataset: dict[str, dict[str, Any]]) -> list[str]:
    schools = set()
    for info in dataset.values():
        for row in info["records"]:
            school = (row.get("学校") or "").strip()
            if school:
                schools.add(school)
    return sorted(schools)


def build_overview(dataset: dict[str, dict[str, Any]]) -> dict[str, Any]:
    user_rows = dataset.get("用户信息", {}).get("records", [])
    grade_rows = dataset.get("每年级学生人数", {}).get("records", [])

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

    dates: list[str] = []
    for info in dataset.values():
        for row in info["records"]:
            period = normalize_period(row)
            if period:
                dates.append(period)

    schools = infer_school_candidates(dataset)
    school_name = schools[0] if len(schools) == 1 else None
    confidence = "high" if school_name else ("mixed" if schools else "unknown")

    semester_rows = dataset.get("最近四个正式学期", {}).get("records", [])
    semesters = [
        {
            "name": row.get("学期名称", ""),
            "start": parse_date(row.get("开始日期", "")),
            "end": parse_date(row.get("结束日期", "")),
        }
        for row in semester_rows
        if row.get("学期名称")
    ]

    return {
        "school_name": school_name,
        "school_name_candidates": schools,
        "school_name_confidence": confidence,
        "coverage_start": min(dates) if dates else None,
        "coverage_end": max(dates) if dates else None,
        "total_users": total_users,
        "roles": role_summary,
        "divisions": dict(sorted(division_counts.items(), key=lambda item: item[1], reverse=True)),
        "grades": grade_table,
        "semesters": semesters,
    }


def normalize_title(source_name: str) -> str:
    title = source_name
    title = title.replace("（近一年）", "")
    title = title.replace("（近", "（")
    title = title.replace("_", "")
    title = title.replace("每周", "")
    title = title.replace("每月", "")
    return title.strip("（） ")


def infer_domain(source_name: str, headers: list[str]) -> tuple[str, str]:
    text = f"{source_name} {' '.join(headers)}"
    rules = [
        ("基础画像", "学校概览与用户基础", ["用户信息", "每年级学生人数", "学期名称"]),
        ("活跃使用", "活跃度与使用黏性", ["周活", "日活", "访问人数"]),
        ("课程教学", "课程与教学执行", ["课程任务", "排课", "课表", "选课", "课时", "评语"]),
        ("评价考试", "成绩、评价与反馈", ["过评", "评价项", "评教", "考试"]),
        ("学生管理", "考勤、请假与日常管理", ["考勤", "请假", "德育", "调代课"]),
        ("协同沟通", "通知、校历与家校协同", ["通知", "校历"]),
        ("资源服务", "资源与空间服务", ["场馆", "预约"]),
    ]
    for domain_key, domain_title, keywords in rules:
        if any(keyword in text for keyword in keywords):
            return domain_key, domain_title
    return "补充观察", "补充业务观察"


def domain_intro(domain_key: str) -> str:
    mapping = {
        "活跃使用": "重点观察平台在学期推进中的活跃表现、关键时段峰值以及近期黏性变化，用于判断系统是否形成持续使用基础。",
        "课程教学": "重点观察课务组织、任务布置、选课与教学执行是否已经稳定嵌入教务流程。",
        "评价考试": "重点观察评价、考试与反馈业务的数据沉淀情况，用于判断教学评价链条的闭环程度。",
        "学生管理": "重点观察考勤、请假、德育与调代课等日常管理流程是否形成制度化线上运行。",
        "协同沟通": "重点观察通知、校历等协同模块是否承担统一信息发布与过程协调职能。",
        "资源服务": "重点观察场地与资源调配是否实现可见、可追踪和相对稳定的线上支撑。",
        "补充观察": "以下模块虽不属于主要教务主链，但仍可补充反映学校系统使用的结构特征。",
    }
    return mapping.get(domain_key, "根据输入数据自动组织的补充业务章节。")


def metric_label(value_field: str | None) -> str:
    return value_field or "业务量"


def build_chart_payload(module: dict[str, Any]) -> dict[str, Any] | None:
    colors = ["#2f80ed", "#00b8d9", "#5b8ff9", "#36cfc9", "#fa8c16", "#f6bd16", "#7256ff", "#eb2f96"]
    series = module.get("series", [])
    categories = module.get("category_totals", {})
    label = metric_label(module.get("value_field"))
    if len(series) >= 2:
        return {
            "kind": "line",
            "title": module["title"],
            "series_name": label,
            "labels": [item["period"] for item in series],
            "values": [float(item["value"]) for item in series],
            "legend": [{"label": label, "color": colors[0]}],
        }
    if categories:
        items = list(categories.items())[:8]
        return {
            "kind": "bar",
            "title": module["title"],
            "series_name": label,
            "labels": [str(k) for k, _ in items],
            "values": [float(v) for _, v in items],
            "legend": [{"label": label, "color": colors[0]}] if len(items) == 1 else [{"label": label_name, "color": colors[idx % len(colors)]} for idx, (label_name, _) in enumerate(items)],
        }
    return None


def module_intro(module: dict[str, Any]) -> str:
    metric = metric_label(module.get("value_field"))
    if module.get("has_time_series"):
        latest_text = f"最近周期为 {module.get('latest_period') or '-'}，当前 {metric}为 {fmt_int(module.get('latest_value'))}"
        return f"本模块用于从教务管理视角观察“{module['title']}”的运行节奏，重点识别该业务在学期过程中的执行连续性、关键节点强度以及近期状态；{latest_text}。"
    if module.get("category_field"):
        return f"本模块用于从教务管理视角观察“{module['title']}”的结构分布，重点识别不同角色、类型或项目之间的资源投向与业务集中度。"
    return f"本模块用于从教务管理视角补充观察“{module['title']}”的总体使用情况，判断该业务是否已被学校稳定纳入日常运行。"


def build_module_issue(module: dict[str, Any]) -> dict[str, str]:
    vol = module.get("volatility_ratio")
    latest = module.get("latest_value")
    peak = module.get("peak_value")
    share = module.get("top_category_share")
    top_category = module.get("top_category")
    domain_key = module.get("domain_key")

    if vol is not None and vol > 0.60:
        return {
            "title": "业务执行波动较大",
            "evidence": f"离散程度约为 {vol:.2f}，峰值出现在 {module.get('peak_period') or '历史观测期'}，峰值为 {fmt_int(peak)}。",
            "cause": "结合校内学期节奏判断，该业务更依赖集中节点推动，尚未表现出稳定、均衡的过程性执行。",
            "implication": "对教务管理而言，这意味着流程执行更像阶段性冲刺，而不是被持续嵌入日常工作机制。",
        }
    if peak and latest is not None and peak > 0 and latest < peak * 0.5 and len(module.get("series", [])) >= 6:
        return {
            "title": "近期执行水平明显回落",
            "evidence": f"最近周期为 {module.get('latest_period') or '-'}，当前值 {fmt_int(latest)}，仅为历史峰值 {fmt_int(peak)} 的 {latest / peak:.0%}。",
            "cause": "可能说明该模块在高峰节点后缺乏持续督导，业务使用尚未转化为稳定的过程要求。",
            "implication": "若这一回落不完全由学期节点解释，则反映出教务推进节奏和系统使用之间存在断层。",
        }
    if share is not None and share >= 0.70 and top_category:
        return {
            "title": "业务覆盖集中于少数场景",
            "evidence": f"当前以“{top_category['name']}”为主，占比约 {fmt_pct(share)}。",
            "cause": "平台目前主要承接单点高频场景，相关业务尚未形成更广的流程覆盖或分层应用。",
            "implication": "从教务管理角度看，这意味着平台支撑能力更多停留在局部环节，横向协同区分度仍不足。",
        }
    if domain_key == "课程教学" and module.get("total_value") is not None and (module.get("latest_value") or 0) == 0 and module.get("series"):
        return {
            "title": "教学过程留痕存在空档风险",
            "evidence": f"观测期内累计业务量为 {fmt_int(module.get('total_value'))}，但最近周期未观察到有效记录。",
            "cause": "可能存在阶段性停用、线下转移处理或录入要求不稳定等情况。",
            "implication": "这会削弱教务部门对教学执行过程的连续追踪能力。",
        }
    if module.get("chart_status") == "data_only":
        return {
            "title": "当前更适合做结构性判断",
            "evidence": f"当前模块共有 {module.get('record_count', 0)} 条记录，主要体现为分类或静态汇总。",
            "cause": "原始数据更偏结构型而非过程型，不足以支持高信度的趋势判断。",
            "implication": "该模块宜用于识别覆盖面和分布格局，而不宜过度延展为长期变化趋势。",
        }
    return {
        "title": "当前运行相对平稳",
        "evidence": "峰值、近期值和结构分布之间未出现明显失衡。",
        "cause": "该模块与校内流程之间已形成一定程度的日常耦合，使用节奏较为自然。",
        "implication": "后续更适合持续观察其稳定性，而非将其判定为显著风险点。",
    }


def module_findings(module: dict[str, Any]) -> list[str]:
    findings: list[str] = []
    metric = metric_label(module.get("value_field"))
    if module.get("peak_value") is not None:
        findings.append(f"峰值周期为 {module.get('peak_period')}，峰值 {metric}达到 {fmt_int(module.get('peak_value'))}。")
    if module.get("latest_value") is not None:
        findings.append(f"最近周期为 {module.get('latest_period')}，当前 {metric}为 {fmt_int(module.get('latest_value'))}。")
    if module.get("top_category"):
        share = module.get("top_category_share")
        share_text = f"，占比约 {fmt_pct(share)}" if isinstance(share, (int, float)) else ""
        findings.append(f"结构上以“{module['top_category']['name']}”为主{share_text}。")
    if module.get("volatility_ratio") is not None:
        findings.append(f"时间序列离散程度约为 {module['volatility_ratio']:.2f}，可作为判断业务稳定性的辅助证据。")
    if module.get("total_value"):
        findings.append(f"观测期内累计 {metric}约为 {fmt_int(module.get('total_value'))}。")
    return findings[:4]


def module_interpretation(module: dict[str, Any], issue: dict[str, str]) -> str:
    title = issue["title"]
    if title == "业务执行波动较大":
        return "从教务管理视角看，该模块运行节奏并不均匀，说明相关流程更多依赖节点性推动，而非被稳定纳入周常或月常工作。若长期维持这种形态，管理者将难以获得连续、可复盘的过程证据。"
    if title == "近期执行水平明显回落":
        return "以校内历史周期为常模参照，当前表现明显低于既往高位，说明近期执行强度不足。若不是学期自然收尾所致，则更可能反映出管理要求和业务跟进有所减弱。"
    if title == "业务覆盖集中于少数场景":
        return "从结构分布看，该模块的应用仍高度集中于少数类型。这样虽然能在局部业务上形成效率优势，但对学校整体教务治理的支撑面仍偏窄。"
    if title == "当前更适合做结构性判断":
        return "由于该模块原始数据更偏结构型而非连续过程型，因此更适合用来识别覆盖面、参与对象和业务配置格局，而不宜直接延伸为趋势性结论。"
    if title == "教学过程留痕存在空档风险":
        return "对教务部门而言，过程留痕的连续性比单一高峰更重要。若最近周期缺乏记录，即使历史上曾有使用，也会削弱过程监管与质量复盘的完整性。"
    return "该模块目前未呈现出明显异常，说明其与学校既有教务运行节奏之间已有一定适配。后续仍应结合学期节点持续观察其稳定性和覆盖深度。"


def build_module(source_name: str, info: dict[str, Any]) -> dict[str, Any] | None:
    records = info["records"]
    if not records:
        return None
    headers = list(records[0].keys()) if records else []
    summary = aggregate_records(records)
    if summary.get("value_field") is None and not summary.get("category_totals"):
        return None
    domain_key, domain_title = infer_domain(source_name, headers)
    module = {
        "key": source_name,
        "title": normalize_title(source_name),
        "source_name": source_name,
        "source_file": info["file_name"],
        "headers": headers,
        "record_count": len(records),
        "domain_key": domain_key,
        "domain_title": domain_title,
        **summary,
    }
    module["chart_payload"] = build_chart_payload(module)
    module["chart_status"] = "available" if module["chart_payload"] else "data_only"
    issue = build_module_issue(module)
    module["intro"] = module_intro(module)
    module["findings"] = module_findings(module)
    module["issue_analysis"] = issue
    module["interpretation"] = module_interpretation(module, issue)
    return module


def module_priority(module: dict[str, Any]) -> float:
    return float(module.get("total_value") or 0) + float(module.get("peak_value") or 0) + float(module.get("latest_value") or 0) + float(module.get("record_count") or 0)


def build_modules(dataset: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    modules = []
    for source_name, info in dataset.items():
        if source_name in {"用户信息", "每年级学生人数", "最近四个正式学期"}:
            continue
        module = build_module(source_name, info)
        if module:
            modules.append(module)
    modules.sort(key=module_priority, reverse=True)
    return modules


def build_dynamic_sections(overview: dict[str, Any], modules: list[dict[str, Any]]) -> list[dict[str, Any]]:
    sections: list[dict[str, Any]] = []
    if overview.get("roles") or overview.get("grades") or overview.get("semesters"):
        sections.append(
            {
                "id": "overview",
                "title": "学校概览与用户基础",
                "intro": "展示平台用户基础、终端绑定、学段结构与学期范围，用于界定后续教务分析的参照边界。",
                "module_keys": [],
                "kind": "overview",
                "priority": float(overview.get("total_users") or 0),
            }
        )

    grouped: defaultdict[str, list[dict[str, Any]]] = defaultdict(list)
    for module in modules:
        grouped[module["domain_title"]].append(module)

    for domain_title, items in sorted(grouped.items(), key=lambda item: sum(module_priority(module) for module in item[1]), reverse=True):
        domain_key = items[0]["domain_key"]
        items.sort(key=module_priority, reverse=True)
        sections.append(
            {
                "id": domain_key,
                "title": domain_title,
                "intro": domain_intro(domain_key),
                "module_keys": [module["key"] for module in items],
                "kind": "modules",
                "priority": sum(module_priority(module) for module in items),
            }
        )
    return sections


def build_summary_bullets(overview: dict[str, Any], modules: list[dict[str, Any]]) -> list[str]:
    bullets: list[str] = []
    if overview.get("total_users"):
        bullets.append(f"平台当前覆盖约 {fmt_int(overview['total_users'])} 个角色账号，已具备支撑校内多角色协同使用的基础盘。")

    active_module = next((m for m in modules if m["domain_key"] == "活跃使用" and m.get("peak_value") is not None), None)
    if active_module:
        bullets.append(f"{active_module['title']}在 {active_module['peak_period']} 达到峰值 {fmt_int(active_module['peak_value'])}，说明平台在关键学期节点具备较强集中使用能力。")

    teaching_module = next((m for m in modules if m["domain_key"] == "课程教学"), None)
    if teaching_module and teaching_module.get("total_value"):
        bullets.append(f"{teaching_module['title']}累计业务量达到 {fmt_int(teaching_module['total_value'])}，反映该系统已在部分教学执行环节形成明显嵌入。")

    issue_module = next((m for m in modules if m["issue_analysis"]["title"] not in {"当前运行相对平稳", "当前更适合做结构性判断"}), None)
    if issue_module:
        bullets.append(f"当前最值得关注的风险来自“{issue_module['title']}”，其主要表现为{issue_module['issue_analysis']['title']}。")

    coverage_domains = sorted({m["domain_title"] for m in modules})
    if coverage_domains:
        bullets.append(f"本次数据已覆盖 {len(coverage_domains)} 个业务域，报告结构将随输入数据范围自动裁剪，不对未出现模块作推断。")

    if overview.get("school_name_confidence") != "high":
        bullets.append("学校名称识别存在不确定性，正式交付前应结合原始导出文件再次校核。")
    return bullets[:6]


def build_metric_cards(overview: dict[str, Any], modules: list[dict[str, Any]]) -> list[dict[str, str]]:
    cards = [{"label": "覆盖账号数", "value": fmt_int(overview.get("total_users"))}]
    if overview.get("coverage_end"):
        cards.append({"label": "数据截止", "value": overview["coverage_end"]})
    if overview.get("semesters"):
        cards.append({"label": "识别学期数", "value": str(len(overview["semesters"]))})
    strongest = max((m for m in modules if m.get("total_value")), key=lambda item: item["total_value"], default=None)
    if strongest:
        cards.append({"label": "高频业务", "value": strongest["title"]})
    return cards[:4]


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

    dataset = load_all_records(input_dir)
    overview = build_overview(dataset)
    if args.school_name:
        overview["school_name"] = args.school_name
        overview["school_name_confidence"] = "manual"

    modules = build_modules(dataset)
    dynamic_sections = build_dynamic_sections(overview, modules)
    summary_bullets = build_summary_bullets(overview, modules)
    metric_cards = build_metric_cards(overview, modules)

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
