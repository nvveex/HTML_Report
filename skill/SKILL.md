---
name: seiue-usage-report
description: 基于工作区 `workspace/input/` 中的 Seiue/希悦 Excel 文件生成专业型 HTML 学校使用报告。适用于用户从 GitHub 拉取一个 bundle 文件夹后，把 Excel 放入固定输入目录，再通过 Codex 调用 skill 自动生成报告的场景。输出应包含执行摘要和个性化业务分析，不应采用磁贴式图表汇编；遇到任何不明确之处时，必须先向用户确认，再执行生成。
---

# Seiue 使用情况报告

这个 skill 用于处理一个固定结构的分发包。用户从 GitHub 拉取 bundle 后，只需要把 Excel 放进 `workspace/input/`，再调用 `$seiue-usage-report`，就应生成一份专业型 HTML 报告。

这个 skill 不依赖外部项目仓库结构，也不再识别 `originaldata/`、`charts/` 或其他外部脚本目录。数据来源只允许是 `workspace/input/` 中的 Excel 文件。

## 目录约定

bundle 应包含以下目录：

- `skill/`
- `workspace/input/`
- `workspace/output/`
- `workspace/artifacts/`

工作时默认以仓库根目录为基准。对于本仓库，默认输入目录必须视为 `HTML_Report/workspace/input/`：
- 输入目录：`workspace/input/`
- 输出目录：`workspace/output/`
- 中间产物目录：`workspace/artifacts/`

除非用户明确要求改目录，否则每次都默认根据 `HTML_Report/workspace/input/` 内的数据生成报告。

## 工作流程

1. 检查 `workspace/input/` 是否存在 Excel。
2. 读取 Excel，抽取学校名称、时间范围、角色规模、年级结构和业务趋势。
3. 在需要时直接生成统一样式的图表，不依赖外部 charts 目录或本地图形库。
4. 按学校实际数据覆盖情况动态组织报告章节。
5. 输出专业型 HTML 报告和中间指标 JSON。

入口约束：
- 只扫描 `workspace/input/` 下的 `*.xlsx`
- 不读取 `originaldata/`
- 不复用 `charts/`
- 不依赖项目外部 `scripts/`

如果在学校名称、输出范围、模块取舍、字段解释、结论口径或用户偏好上存在任何不明确之处，必须先向用户确认，完全确认后再执行报告生成。

## 使用方式

如果用户的当前目录是 bundle 根目录，可直接运行：

```bash
python3 skill/scripts/run_seiue_report.py --bundle-root "$PWD"
```

如果通过 Codex 调用 skill，则应默认：
- 扫描当前 bundle 的 `workspace/input/`
- 将结果输出到 `workspace/output/`
- 把中间文件写入 `workspace/artifacts/`

不要要求用户理解内部脚本链路。

## 输入输出规则

输入：
- 默认扫描 `HTML_Report/workspace/input/*.xlsx`

输出：
- HTML 报告输出到 `workspace/output/`
- 指标 JSON 输出到 `workspace/artifacts/`

默认文件命名：
- `<学校名>-希悦系统使用情况说明-<YYYY-MM-DD>.html`
- 如果学校名无法可靠识别，则使用 `XX学校-希悦系统使用情况说明-<YYYY-MM-DD>.html`

同一 bundle 多次运行时，应保留历史 HTML，不默认覆盖旧文件。

## 报告规则

报告必须是专业分析报告，不是磁贴页。以下规则属于强制执行项，不允许在生成时随意替换或放宽。

强制执行项：
- 结构动态生成，不强制套用统一章节模板
- 有标题、执行摘要、正文分析、图表或表格
- 每条结论应尽量追溯到具体指标、表格或图表
- 每个业务章节应体现“判断 + 证据 + 解释 + 管理含义”
- 不得输出单独的“问题与挑战分析”总章节，模块问题必须直接写在对应模块下
- 不得输出“下一步管理行动建议”章节
- 不得使用磁贴式、卡片墙式、仪表盘式布局
- 主题色固定为蓝白体系
- 图表固定为统一样式的 canvas 图表
- 图表配色必须丰富，但整体风格保持专业、现代、克制
- 图表尺寸必须受正文版心约束，不得超出文档宽度
- 每个模块必须严格按“标题说明 → 图表 → 分析”三段纵向排列
- 不得采用图表与分析横向并排布局
- 图表必须带图例
- 不得输出“图表由 bundle 内置流程生成”“当前模块未生成图表”等实现性提示文本
- 允许用户在报告中继续增加个性化内容，渲染结构不得阻碍后续人工增补

## 图表策略

打包版本默认不依赖预先存在的 charts。

优先策略：
- 在报告生成流程中直接产出统一风格的图表
- 图表默认由 HTML 内置 canvas 绘制，不使用 SVG
- 如果某模块不适合图表展示，则使用表格和文字分析完成表达
- 不允许因为缺少 charts 目录就把报告降级成半成品

## 分析标准

详细规则见：
- [references/report-structure.md](references/report-structure.md)
- [references/metric-rules.md](references/metric-rules.md)
- [references/analysis-writing.md](references/analysis-writing.md)
- [references/terminology-guide.md](references/terminology-guide.md)
- [references/style-guide.md](references/style-guide.md)

硬性要求：
- 不复述原始数据
- 不做无证据因果断言
- 原因分析使用审慎表达
- 教育测量/评估术语只有在证据满足时使用

## 版式要求

页面应采用正式报告版式，且以下要求强制执行：
- 标题、摘要、正文段落、图文分析、表格、结论
- 不使用仪表盘式、磁贴式、卡片墙式布局
- 图表服务于论证，不喧宾夺主
- 页面应适合学校管理层阅读和交付
- 蓝白主题为唯一允许的页面主色方案
- 模块之间必须通过留白、浅底分节或分隔线形成清晰边界

## 产物要求

最终至少输出：
- 1 份 HTML 报告
- 1 份中间指标 JSON
- 必要时输出图表资源到 `workspace/artifacts/figures/`

如果用户要求直接生成报告，就直接运行入口脚本，不要只解释步骤。
