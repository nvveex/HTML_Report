# seiue-report-bundle

使用方式很简单：

1. 把希悦导出的 Excel 文件放进 `workspace/input/`
2. 在 bundle 根目录调用 `$seiue-usage-report`

如果需要在本地直接验证，也可以运行：

```bash
python3 skill/scripts/run_seiue_report.py --bundle-root .
```

生成结果：
- 报告 HTML：`workspace/output/`
- 中间指标：`workspace/artifacts/`

默认输出文件名为：
- `<学校名>-希悦系统使用情况说明-<日期>.html`
