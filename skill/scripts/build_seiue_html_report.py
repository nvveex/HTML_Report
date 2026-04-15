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


def _pct(value: Any) -> str:
    if value is None:
        return "-"
    return f"{float(value) * 100:.1f}%"


def _fmt_num(value: Any) -> str:
    if value is None:
        return "-"
    return f"{int(round(float(value))):,}"


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


def canvas_chart(module: dict[str, Any]) -> str:
    payload = module.get("chart_payload")
    if not payload:
        categories = module.get("category_totals", {})
        if not categories:
            return ""
        rows = "".join(
            f"<tr><td>{esc(name)}</td><td>{esc(_fmt_num(value))}</td></tr>"
            for name, value in list(categories.items())[:8]
        )
        return f"""
        <div class="module-chart">
          <table class="data-table compact-table">
            <thead><tr><th>分类</th><th>数值</th></tr></thead>
            <tbody>{rows}</tbody>
          </table>
        </div>
        """
    payload_json = esc(json.dumps(payload, ensure_ascii=False))
    canvas_id = f"chart-{abs(hash(module.get('key')))}"
    legend_html = "".join(
        f'<span class="chart-legend-item"><i style="background:{esc(item.get("color"))}"></i>{esc(item.get("label"))}</span>'
        for item in payload.get("legend", [])
    )
    return f"""
    <div class="module-chart">
      <div class="chart-frame">
        <canvas id="{canvas_id}" class="report-chart" data-chart='{payload_json}' width="780" height="340"></canvas>
      </div>
      <div class="chart-legend">{legend_html}</div>
    </div>
    """


def module_findings(module: dict[str, Any]) -> str:
    items = module.get("findings") or []
    if not items:
        return ""
    return "<ul class=\"bullet-list\">" + "".join(f"<li>{esc(item)}</li>" for item in items) + "</ul>"


def module_analysis(module: dict[str, Any]) -> str:
    issue = module.get("issue_analysis") or {}
    issue_block = ""
    if issue:
        issue_block = f"""
        <div class="module-issue">
          <p><strong>主要问题：</strong>{esc(issue.get("title"))}</p>
          <p><strong>证据：</strong>{esc(issue.get("evidence"))}</p>
          <p><strong>可能原因：</strong>{esc(issue.get("cause"))}</p>
          <p><strong>管理含义：</strong>{esc(issue.get("implication"))}</p>
        </div>
        """
    return f"""
    <div class="module-analysis">
      {module_findings(module)}
      <p>{esc(module.get("interpretation"))}</p>
      {issue_block}
    </div>
    """


def module_article(module: dict[str, Any]) -> str:
    chart_html = canvas_chart(module)
    if not chart_html:
        chart_html = """
        <div class="module-chart">
          <table class="data-table compact-table">
            <thead><tr><th>指标</th><th>数值</th></tr></thead>
            <tbody><tr><td>观测记录数</td><td>-</td></tr></tbody>
          </table>
        </div>
        """
    return f"""
    <article class="module-article">
      <div class="module-block module-intro">
        <div class="module-header">
          <span class="module-label">{esc(module.get("title"))}</span>
          <h3>{esc(module.get("title"))}</h3>
        </div>
        <p>{esc(module.get("intro"))}</p>
      </div>
      <div class="module-block">
        {chart_html}
      </div>
      <div class="module-block">
        {module_analysis(module)}
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


def chart_script() -> str:
    return """
(function () {
  const ink = '#15314f';
  const grid = '#d7e4f2';
  const axis = '#90a8c2';
  const palette = ['#2f80ed', '#00b8d9', '#5b8ff9', '#36cfc9', '#fa8c16', '#f6bd16', '#7256ff', '#eb2f96'];
  const fill = 'rgba(47, 128, 237, 0.12)';
  const font = "'PingFang SC','Hiragino Sans GB','Microsoft YaHei',sans-serif";

  function niceNumber(value) {
    if (!Number.isFinite(value)) return '0';
    if (Math.abs(value) >= 10000) return Math.round(value).toLocaleString('zh-CN');
    return String(Math.round(value));
  }

  document.querySelectorAll('canvas.report-chart').forEach((canvas) => {
    const payload = JSON.parse(canvas.dataset.chart || '{}');
    const ctx = canvas.getContext('2d');
    const dpr = window.devicePixelRatio || 1;
    const cssWidth = canvas.width;
    const cssHeight = canvas.height;
    canvas.style.width = '100%';
    canvas.style.maxWidth = cssWidth + 'px';
    canvas.style.height = 'auto';
    canvas.width = cssWidth * dpr;
    canvas.height = cssHeight * dpr;
    ctx.scale(dpr, dpr);
    ctx.clearRect(0, 0, cssWidth, cssHeight);

    const left = 64;
    const right = 24;
    const top = 28;
    const bottom = 54;
    const plotW = cssWidth - left - right;
    const plotH = cssHeight - top - bottom;
    const values = (payload.values || []).map(Number);
    const labels = payload.labels || [];
    if (!values.length) return;

    const maxV = Math.max(...values, 1);
    const minV = payload.kind === 'line' ? Math.min(...values) : 0;
    const span = Math.max(maxV - minV, 1);

    ctx.font = `12px ${font}`;
    ctx.strokeStyle = grid;
    ctx.lineWidth = 1;
    for (let i = 0; i <= 4; i += 1) {
      const y = top + (plotH / 4) * i;
      ctx.beginPath();
      ctx.moveTo(left, y);
      ctx.lineTo(left + plotW, y);
      ctx.stroke();
      const tickValue = maxV - (span / 4) * i;
      ctx.fillStyle = axis;
      ctx.textAlign = 'right';
      ctx.fillText(niceNumber(tickValue), left - 10, y + 4);
    }

    ctx.strokeStyle = axis;
    ctx.beginPath();
    ctx.moveTo(left, top);
    ctx.lineTo(left, top + plotH);
    ctx.lineTo(left + plotW, top + plotH);
    ctx.stroke();

    if (payload.kind === 'line') {
      const points = values.map((value, index) => {
        const x = left + (plotW * index / Math.max(values.length - 1, 1));
        const y = top + plotH - ((value - minV) / span) * plotH;
        return { x, y, value };
      });

      ctx.beginPath();
      points.forEach((point, index) => {
        if (index === 0) ctx.moveTo(point.x, point.y);
        else ctx.lineTo(point.x, point.y);
      });
      ctx.strokeStyle = palette[0];
      ctx.lineWidth = 3;
      ctx.stroke();

      ctx.lineTo(points[points.length - 1].x, top + plotH);
      ctx.lineTo(points[0].x, top + plotH);
      ctx.closePath();
      ctx.fillStyle = fill;
      ctx.fill();

      points.forEach((point) => {
        ctx.fillStyle = '#ffffff';
        ctx.beginPath();
        ctx.arc(point.x, point.y, 5, 0, Math.PI * 2);
        ctx.fill();
        ctx.strokeStyle = palette[0];
        ctx.lineWidth = 2;
        ctx.stroke();
      });

      ctx.fillStyle = axis;
      ctx.textAlign = 'center';
      labels.forEach((label, index) => {
        const x = left + (plotW * index / Math.max(labels.length - 1, 1));
        if (index === 0 || index === labels.length - 1 || index % Math.ceil(labels.length / 5) === 0) {
          ctx.fillText(String(label), x, cssHeight - 18);
        }
      });
    } else {
      const rowGap = 16;
      const barHeight = Math.max(16, Math.min(28, Math.floor((plotH - rowGap * (values.length - 1)) / Math.max(values.length, 1))));
      values.forEach((value, index) => {
        const y = top + index * (barHeight + rowGap);
        const barW = (value / maxV) * plotW;
        const color = palette[index % palette.length];
        ctx.fillStyle = color;
        ctx.fillRect(left, y, barW, barHeight);
        ctx.fillStyle = ink;
        ctx.textAlign = 'right';
        ctx.fillText(String(labels[index] || '').slice(0, 18), left - 10, y + barHeight * 0.72);
        ctx.textAlign = 'left';
        ctx.fillText(niceNumber(value), left + barW + 8, y + barHeight * 0.72);
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
        ]
    )
    return (
        template.replace("{{TITLE}}", esc(metrics.get("title")))
        .replace("{{INLINE_CSS}}", css)
        .replace("{{DATE_RANGE}}", esc(date_range))
        .replace("{{SCHOOL_NAME}}", esc(overview.get("school_name") or "未可靠识别"))
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
