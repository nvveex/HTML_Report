# seiue-report-bundle

安装与使用：

1. 拉取仓库后，先在 bundle 根目录执行：

```bash
bash install_skill.sh
```

2. 把希悦导出的 Excel 文件放进 `workspace/input/`
3. 在 Codex 中调用 `$seiue-usage-report`

说明：
- skill 只会读取 `workspace/input/` 中的 Excel 文件
- 不再识别 `originaldata/`
- 不再依赖 `charts/` 或外部脚本目录
- `install_skill.sh` 会把当前仓库里的 `skill/` 同步到 `~/.codex/skills/seiue-usage-report/`
- 如果仓库里的 skill 有更新，重新执行一次 `bash install_skill.sh` 即可覆盖更新本机版本

如果需要在本地直接验证，也可以运行：

```bash
python3 skill/scripts/run_seiue_report.py --bundle-root .
```

生成结果：
- 报告 HTML：`workspace/output/`
- 中间指标：`workspace/artifacts/`

默认输出文件名为：
- `<学校名>-希悦系统使用情况说明-<日期>.html`
