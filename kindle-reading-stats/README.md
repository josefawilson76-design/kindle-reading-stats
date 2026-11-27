# Kindle Reading Stats

这是一个 Codex Skill，也是一个可直接运行的 Python 工具。它通过 USB
读取 Kindle 的 KRDS sidecar 文件，提取每本书的累计阅读时间、阅读速度
（WPM）、阅读进度和词数，并生成 Excel 友好的 CSV、JSON 和 HTML 报告。

目标设备包括 Kindle Paperwhite 4（KPW4）固件 5.16.3.1。解析器只读访问
设备，不需要越狱，也不会修改 Kindle 上的数据。

## 数据来源

Kindle 把阅读状态放在电子书旁边的 `.sdr` 目录中。常见 sidecar 扩展名：

| 电子书格式 | 可变 sidecar | 常驻 sidecar |
| --- | --- | --- |
| KF8/AZW3 | `.azw3f` | `.azw3r` |
| KFX | `.yjf` | `.yjr` |
| MOBI | `.mbs` | `.mbp1` |
| PDF | `.pdt` | `.pds` |
| Topaz | `.tas` | `.tal` |

本工具递归扫描这些扩展名。核心字段：

- `timer.model.totalTime`：累计阅读毫秒数。
- `timer.model.totalWords`：累计已读词数。
- `timer.model.totalPercent`：阅读器记录的进度。
- `book.info.store.numberOfWords`：书籍估算总词数。
- `book.info.store.percentOfBook`：词数统计对书籍的覆盖率。

WPM 使用以下公式计算：

```text
WPM = totalWords / (totalTime / 60000)
```

如果 `totalPercent` 无效，阅读进度回退为：

```text
progress = min(totalWords / numberOfWords * 100, 100)
```

## 安装 Skill

把整个 `kindle-reading-stats` 目录放入 Codex 的个人技能目录：

```text
Windows: %USERPROFILE%\.codex\skills\kindle-reading-stats
macOS/Linux: ~/.codex/skills/kindle-reading-stats
```

重新打开 Codex 后，可以用 `$kindle-reading-stats` 调用。

## 直接运行脚本

要求 Python 3.9 或更高版本，无第三方依赖。

在 Codex 桌面环境中，`python` 可能不在 `PATH` 中，但 Codex 自带运行时通常
位于 `%USERPROFILE%\.cache\codex-runtimes`。可以直接把
`python.exe` 的绝对路径替换到下面的示例中。

### Windows

```powershell
python .\kindle-reading-stats\scripts\extract_kindle_stats.py `
  --kindle-path E:\ `
  --output-dir .\kindle-reading-stats-output `
  --format all
```

### macOS / Linux

```bash
python3 kindle-reading-stats/scripts/extract_kindle_stats.py \
  --kindle-path /media/user/Kindle \
  --output-dir ./kindle-reading-stats-output \
  --format all
```

只扫描 `documents` 目录通常更快：

```bash
python3 kindle-reading-stats/scripts/extract_kindle_stats.py \
  --kindle-path /media/user/Kindle/documents \
  --format csv
```

## 输出文件

`--format all` 会在输出目录生成：

- `kindle-reading-stats.csv`：适合 Excel、Numbers、LibreOffice 和
  pandas，文件使用 UTF-8 BOM。
- `kindle-reading-stats.json`：适合程序处理，包含书籍统计、来源文件和
  逐文件解析诊断。
- `kindle-reading-stats.html`：无需联网的本地可视化表格，支持搜索、
  点击表头排序和进度条。

CSV 主要列：

| 列名 | 含义 |
| --- | --- |
| `title` | 书名 |
| `author` | 作者 |
| `total_time_ms` | 累计阅读毫秒数 |
| `total_time_hours` | 累计阅读小时数 |
| `total_words` | 累计已读词数 |
| `book_words` | 书籍估算总词数 |
| `wpm` | 阅读速度 |
| `progress_percent` | 阅读进度百分比 |
| `progress_source` | 进度来源 |
| `source_files` | 使用的 KRDS 文件 |
| `status` | `ok`、`partial` 或 `empty` |
| `warnings` | 容错解析产生的警告 |

## 常用参数

```text
--format all|csv|json|html       输出格式，可逗号组合
--output-dir PATH               输出目录，默认当前目录
--prefix NAME                   输出文件前缀
--include-empty                 保留没有有效计时字段的 sidecar 分组
--quiet                         不打印控制台汇总表
--verbose                       打印每个文件的解析细节
```

## 兼容性与限制

- 只统计本机 sidecar 中仍然存在的数据；已从设备删除的书籍历史无法恢复。
- KRDS 是 Amazon 未公开格式，固件升级后可能出现新的字段。解析器会尽量
  保留已识别字段，并把未知或截断部分写入诊断。
- 书名和作者通常不直接存在于计时 KRDS 记录中。AZW3/MOBI 优先从文件
  元数据读取；KFX 或元数据损坏时使用文件名和 `.sdr` 目录名推断。
- `totalPercent` 与 `totalWords / numberOfWords` 可能因重读、样章或统计
  口径不同而不一致。CSV 会保留两个值，HTML 显示选定的主要进度。
- 不要对正在写入的 Kindle 拔线。导出前先在系统中安全弹出设备。
