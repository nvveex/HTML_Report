#!/usr/bin/env python3
from __future__ import annotations

import argparse
import html
import json
from pathlib import Path
from typing import Any


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
ASSETS_DIR = SKILL_DIR / "assets"


def esc(value: Any) -> str:
    return html.escape("" if value is None else str(value))


def load_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def top_metrics(cards: list[dict[str, str]]) -> str:
    return "".join(
        f'<div class="hero-metric"><span class="hero-metric-label">{esc(card.get("label"))}</span><span class="hero-metric-value">{esc(card.get("value"))}</span></div>'
        for card in cards
    )


def summary_section(summary_bullets: list[str]) -> str:
    items = "".join(f"<li>{esc(item)}</li>" for item in summary_bullets)
    return f"""
    <section class="section">
      <div class="section-head">
        <p class="eyebrow">Executive Summary</p>
        <h2>执行摘要</h2>
      </div>
      <ul class="summary-list">{items}</ul>
    </section>
    """


def overview_block(overview: dict[str, Any]) -> str:
    role_rows = "".join(
        f"<tr><td>{esc(row.get('role'))}</td><td>{esc(int(round(row.get('count', 0))))}</td><td>{esc(_pct(row.get('wechat_rate')))}</td><td>{esc(_pct(row.get('mobile_rate')))}</td><td>{esc(_pct(row.get('dingtalk_rate')))}</td></tr>"
        for row in overview.get("roles", [])
    )
    grade_rows = "".join(
        f"<tr><td>{esc(row.get('division'))}</td><td>{esc(row.get('grade'))}</td><td>{esc(int(round(row.get('students', 0))))}</td></tr>"
        for row in overview.get("grades", [])
    )
    inline = "".join(
        f'<span class="inline-metric"><strong>{esc(name)}</strong> {esc(int(round(value)))}</span>'
        for name, value in overview.get("divisions", {}).items()
    )
    return f"""
    <div class="subsection">
      <h3>学校概览</h3>
      <p class="section-intro">学校概览部分用于界定平台的基础覆盖面，包括角色规模、终端绑定和学生结构，为后续业务表现分析提供参照。</p>
      <div class="subsection">
        <h3>用户规模与绑定情况</h3>
        <table class="data-table">
          <thead><tr><th>角色</th><th>人数</th><th>微信绑定率</th><th>手机绑定率</th><th>钉钉绑定率</th></tr></thead>
          <tbody>{role_rows}</tbody>
        </table>
      </div>
      <div class="subsection">
        <h3>学段与年级结构</h3>
        <table class="data-table">
          <thead><tr><th>学部</th><th>年级</th><th>学生人数</th></tr></thead>
          <tbody>{grade_rows}</tbody>
        </table>
        <div class="inline-metrics">{inline}</div>
      </div>
    </div>
    """


def _pct(value: Any) -> str:
    if value is None:
        return "-"
    return f"{float(value) * 100:.1f}%"


def _fmt_num(value: Any) -> str:
    if value is None:
        return "-"
    return f"{int(round(float(value))):,}"


def module_findings(module: dict[str, Any]) -> list[str]:
    findings = []
    if module.get("peak_value") is not None:
        findings.append(f"峰值周期为 {module.get('peak_period')}，峰值为 {_fmt_num(module.get('peak_value'))}。")
    if module.get("latest_value") is not None:
        findings.append(f"最近周期为 {module.get('latest_period')}，当前值为 {_fmt_num(module.get('latest_value'))}。")
    if module.get("top_category"):
        share = module.get("top_category_share")
        share_text = f"，占比约 {share * 100:.1f}%" if isinstance(share, (float, int)) else ""
        findings.append(f"结构上以“{module['top_category']['name']}”为主{share_text}。")
    if module.get("volatility_ratio") is not None:
        findings.append(f"时间序列离散程度约为 {module['volatility_ratio']:.2f}，可用于判断节奏稳定性。")
    return findings[:4]


def module_interpretation(module: dict[str, Any]) -> str:
    vol = module.get("volatility_ratio")
    share = module.get("top_category_share")
    if vol is not None and vol > 0.60:
        return "该模块表现出较强的阶段性波动，结合学校业务节奏判断，更接近节点驱动型使用，而非常态连续使用。"
    if share is not None and share > 0.70:
        return "该模块的结构集中度偏高，说明系统已经深度嵌入某个高频场景，但横向扩展到其他场景的区分度仍有限。"
    return "该模块整体较为平稳，说明相关业务流程与平台之间已建立一定的日常协同关系。"


def module_implication(module: dict[str, Any]) -> str:
    if module.get("chart_status") != "available":
        return "当前模块未形成图表，本段结论主要依据明细数据和聚合统计，解释信心相对保守。"
    if module.get("latest_value") is not None and module.get("peak_value") is not None:
        latest = float(module["latest_value"])
        peak = float(module["peak_value"])
        if peak and latest < peak * 0.5:
            return "近期水平明显低于峰值，需要进一步区分其属于正常学期节奏变化，还是流程执行未形成常态化。"
    return "当前模块未发现强烈异常，但仍应结合学期节点持续观察其稳定性和覆盖面。"


def canvas_chart(module: dict[str, Any]) -> str:
    payload = module.get("chart_payload")
    if not payload:
        return '<div class="module-chart placeholder"><p>当前模块未生成图表，以下结论以表格和聚合统计为主。</p></div>'
    payload_json = esc(json.dumps(payload, ensure_ascii=False))
    canvas_id = f"chart-{abs(hash(module.get('key')))}"
    return f"""
    <div class="module-chart">
      <canvas id="{canvas_id}" class="report-chart" data-chart='{payload_json}' width="640" height="340"></canvas>
      <p class="chart-note">图表由 bundle 内置 canvas 图表流程根据 Excel 自动生成，样式在整份报告中保持统一。</p>
    </div>
    """


def module_article(module: dict[str, Any]) -> str:
    findings = "".join(f"<li>{esc(item)}</li>" for item in module_findings(module))
    return f"""
    <article class="module-article">
      <div class="module-header">
        <span class="module-label">{esc(module.get("title"))}</span>
        <h3>{esc(module.get("title"))}</h3>
      </div>
      <div class="module-layout">
        <div class="module-text">
          <ul class="bullet-list">{findings}</ul>
          <p>{esc(module_interpretation(module))}</p>
          <p class="small-note">{esc(module_implication(module))}</p>
        </div>
        {canvas_chart(module)}
      </div>
    </article>
    """


def data_sections(metrics: dict[str, Any]) -> str:
    overview = metrics.get("overview", {})
    module_map = {m["key"]: m for m in metrics.get("modules", [])}
    parts = []
    for section in metrics.get("dynamic_sections", []):
        if section.get("kind") == "overview":
            content = overview_block(overview)
        else:
            modules = [module_map[key] for key in section.get("module_keys", []) if key in module_map]
            if not modules:
                continue
            content = '<div class="module-section">' + "".join(module_article(module) for module in modules) + "</div>"
        parts.append(
            f"""
            <section class="section">
              <div class="section-head">
                <p class="eyebrow">Data Analysis</p>
                <h2>{esc(section.get("title"))}</h2>
                <p class="section-intro">{esc(section.get("intro"))}</p>
              </div>
              {content}
            </section>
            """
        )
    return "".join(parts)


def issues_section(issues: list[dict[str, str]]) -> str:
    items = "".join(
        f"""
        <div class="issue-item">
          <h3>{esc(issue.get("title"))}</h3>
          <p><strong>现象：</strong>{esc(issue.get("phenomenon"))}</p>
          <p><strong>证据：</strong>{esc(issue.get("evidence"))}</p>
          <p><strong>可能原因：</strong>{esc(issue.get("cause"))}</p>
          <p class="small-note"><strong>管理含义：</strong>{esc(issue.get("implication"))}</p>
        </div>
        """
        for issue in issues
    )
    return f"""
    <section class="section">
      <div class="section-head">
        <p class="eyebrow">Challenges</p>
        <h2>问题与挑战分析</h2>
        <p class="section-intro">问题分析重点关注结构失衡、使用断层和节奏波动，不重复铺陈原始数据。</p>
      </div>
      {items}
    </section>
    """


def actions_section(actions: list[dict[str, str]]) -> str:
    items = "".join(
        f"""
        <div class="action-item">
          <p class="owner">{esc(action.get("owner"))}</p>
          <h3>对应问题：{esc(action.get("issue"))}</h3>
          <p>{esc(action.get("action"))}</p>
        </div>
        """
        for action in actions
    )
    return f"""
    <section class="section">
      <div class="section-head">
        <p class="eyebrow">Action Plan</p>
        <h2>下一步管理行动建议</h2>
        <p class="section-intro">行动建议与前述问题逐一对应，强调责任主体、目标动作和短周期可执行性。</p>
      </div>
      {items}
    </section>
    """


def chart_script() -> str:
    return """
(function () {
  const ink = '#24343b';
  const axis = '#cfc7b8';
  const accent = '#2d5a56';
  const fill = 'rgba(45, 90, 86, 0.12)';
  document.querySelectorAll('canvas.report-chart').forEach((canvas) => {
    const payload = JSON.parse(canvas.dataset.chart || '{}');
    const ctx = canvas.getContext('2d');
    const dpr = window.devicePixelRatio || 1;
    const cssWidth = canvas.width;
    const cssHeight = canvas.height;
    canvas.style.width = cssWidth + 'px';
    canvas.style.height = cssHeight + 'px';
    canvas.width = cssWidth * dpr;
    canvas.height = cssHeight * dpr;
    ctx.scale(dpr, dpr);
    ctx.clearRect(0, 0, cssWidth, cssHeight);
    const left = 54, right = 18, top = 24, bottom = 42;
    const plotW = cssWidth - left - right;
    const plotH = cssHeight - top - bottom;
    const values = (payload.values || []).map(Number);
    const labels = payload.labels || [];
    if (!values.length) return;
    const maxV = Math.max(...values, 1);
    const minV = payload.kind === 'line' ? Math.min(...values) : 0;
    const span = Math.max(maxV - minV, 1);
    ctx.strokeStyle = axis;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(left, top);
    ctx.lineTo(left, top + plotH);
    ctx.lineTo(left + plotW, top + plotH);
    ctx.stroke();
    ctx.fillStyle = '#6a7279';
    ctx.font = '12px sans-serif';
    ctx.textAlign = 'right';
    ctx.fillText(String(Math.round(maxV)), left - 8, top + 10);
    ctx.fillText(String(Math.round(minV)), left - 8, top + plotH);
    if (payload.kind === 'line') {
      const points = values.map((value, index) => {
        const x = left + (plotW * index / Math.max(values.length - 1, 1));
        const y = top + plotH - ((value - minV) / span) * plotH;
        return {x, y, value};
      });
      ctx.beginPath();
      points.forEach((p, idx) => idx ? ctx.lineTo(p.x, p.y) : ctx.moveTo(p.x, p.y));
      ctx.strokeStyle = accent;
      ctx.lineWidth = 3;
      ctx.stroke();
      ctx.lineTo(points[points.length - 1].x, top + plotH);
      ctx.lineTo(points[0].x, top + plotH);
      ctx.closePath();
      ctx.fillStyle = fill;
      ctx.fill();
      ctx.fillStyle = accent;
      points.forEach((p) => {
        ctx.beginPath();
        ctx.arc(p.x, p.y, 3.5, 0, Math.PI * 2);
        ctx.fill();
      });
      ctx.fillStyle = '#6a7279';
      ctx.textAlign = 'left';
      ctx.fillText(labels[0] || '', left, cssHeight - 12);
      ctx.textAlign = 'right';
      ctx.fillText(labels[labels.length - 1] || '', left + plotW, cssHeight - 12);
    } else {
      const rowH = Math.max(26, Math.min(42, Math.floor(plotH / Math.max(values.length, 1))));
      ctx.font = '12px sans-serif';
      values.forEach((value, index) => {
        const y = top + index * rowH;
        const barW = (value / maxV) * plotW;
        ctx.fillStyle = accent;
        ctx.fillRect(left, y + 6, barW, 16);
        ctx.fillStyle = ink;
        ctx.textAlign = 'right';
        ctx.fillText((labels[index] || '').slice(0, 18), left - 10, y + 18);
        ctx.textAlign = 'left';
        ctx.fillText(String(Math.round(value)), left + barW + 8, y + 18);
      });
    }
  });
})();
"""


def render_html(metrics: dict[str, Any]) -> str:
    template = load_text(ASSETS_DIR / "report_template.html")
    css = load_text(ASSETS_DIR / "report.css")
    overview = metrics.get("overview", {})
    date_range = f"{overview.get('coverage_start') or '-'} 至 {overview.get('coverage_end') or '-'}"
    body = "".join(
        [
            summary_section(metrics.get("summary_bullets", [])),
            data_sections(metrics),
            issues_section(metrics.get("issues", [])),
            actions_section(metrics.get("actions", [])),
        ]
    )
    return (
        template.replace("{{TITLE}}", esc(metrics.get("title")))
        .replace("{{INLINE_CSS}}", css)
        .replace("{{DATE_RANGE}}", esc(date_range))
        .replace("{{SCHOOL_NAME}}", esc(overview.get("school_name") or "未可靠识别"))
        .replace("{{DATA_NOTE}}", esc(metrics.get("data_note", "")))
        .replace("{{TOP_CARDS}}", top_metrics(metrics.get("metric_cards", [])))
        .replace("{{BODY_HTML}}", body)
        .replace("{{CHART_SCRIPT}}", chart_script())
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--metrics-json", required=True)
    parser.add_argument("--output-html", required=True)
    args = parser.parse_args()

    metrics = json.loads(Path(args.metrics_json).read_text(encoding="utf-8"))
    Path(args.output_html).write_text(render_html(metrics), encoding="utf-8")
    print(json.dumps({"output_html": str(Path(args.output_html).resolve())}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
