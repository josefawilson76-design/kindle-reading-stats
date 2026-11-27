---
name: kindle-reading-stats
description: "从通过 USB 挂载的 Kindle KRDS sidecar 文件中提取每本书的累计阅读时长、WPM、阅读进度和词数，并导出 CSV、JSON 及可视化 HTML 表格。适用于 KPW4 5.16.3.1 等使用 KRDS v1 sidecar 的设备；不处理 Amazon 云端阅读统计或加密电子书正文。"
---

# Kindle Reading Stats

## Workflow

1. 确认用户提供的是 Kindle 挂载根目录或 `documents` 目录，例如 `E:\`、`/media/user/Kindle`。
2. 以只读方式运行解析器：

   ```bash
   python scripts/extract_kindle_stats.py \
     --kindle-path /path/to/kindle \
     --output-dir ./kindle-reading-stats-output \
     --format all
   ```
   如果当前环境没有名为 `python` 的命令，可使用 Windows Python Launcher：

   ```powershell
   py -3 scripts/extract_kindle_stats.py --help
   ```

3. 检查命令输出的书籍数、解析失败数和警告。默认生成：
   - `kindle-reading-stats.csv`：UTF-8 BOM，便于 Excel 直接打开。
   - `kindle-reading-stats.json`：包含书籍统计与逐文件诊断。
   - `kindle-reading-stats.html`：可搜索、排序并显示进度条的可视化表格。
4. 汇报主要发现、输出文件位置，以及书名/作者或统计字段来自 fallback 的情况。

## Operating Constraints

- 始终只读扫描 Kindle。不要写回设备，除非用户明确要求并单独确认。
- WPM 计算为 `totalWords / (totalTimeMs / 60000)`。正数才输出，分母为零时保留为空。
- 阅读进度优先使用 `timer.model.totalPercent`；该字段无效时使用 `totalWords / book.info.store.numberOfWords`。
- `book.info.store.percentOfBook` 表示词数覆盖率，不是阅读进度，单独保留用于诊断。
- KRDS 通常不保存书名和作者。优先匹配 `.sdr` 旁对应的电子书文件并读取 MOBI/AZW3 元数据；KFX 或元数据不可读时退回文件名解析。
- 单文件损坏不应终止整次扫描。保留已解码记录，并在 JSON/CSV 中标记 `partial`、`error` 或 `empty`。

## Reference

修改二进制解析器或调查新固件兼容性时，读取 [references/krds-format.md](references/krds-format.md)。普通导出任务不需要加载该参考。

## Validation

从技能目录运行：

```bash
python -m unittest discover -s tests -v
```
