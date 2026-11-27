# Kindle Reading Dashboard

在 Windows 上以只读方式读取通过 USB 挂载的 Kindle，汇总阅读时长、阅读进度、
WPM、词数、漫画页数与历史趋势，并生成本地 CSV、HTML 看板和 XLSX 报表。

[English README](README.en.md) | [发布页](https://josefawilson76-design.github.io/kindle-reading-stats/) | [GitHub Releases](https://github.com/josefawilson76-design/kindle-reading-stats/releases)

**版本：** v1.13.0

**发布日期：** 2026-11-27

**平台：** Windows 10 / 11 x64

**许可证：** GPL v3

## 下载

| 版本 | 适用场景 | 下载 |
| --- | --- | --- |
| 纯文字版 | 文字书统计、看板、历史、成就；不加载 OCR，不分类漫画 | [KindleReader-Text-v1.13.0.exe](https://github.com/josefawilson76-design/kindle-reading-stats/releases/download/v1.13.0/KindleReader-Text-v1.13.0.exe) |
| 文字 + 漫画版 | 额外支持漫画分类、页数/进度、KFX 位置映射和可选 OCR | [KindleReader-OCR-v1.13.0.exe](https://github.com/josefawilson76-design/kindle-reading-stats/releases/download/v1.13.0/KindleReader-OCR-v1.13.0.exe) |

两个程序使用同一套核心统计逻辑。区别是：**纯文字版不支持漫画，也不加载 OCR
模块；漫画版支持漫画和可选 OCR。**

Release 同时提供：

- [`SHA256SUMS.txt`](https://github.com/josefawilson76-design/kindle-reading-stats/releases/download/v1.13.0/SHA256SUMS.txt)
- `settings.example.json`
- `book_modes.example.json`

## 快速开始

1. 用支持数据传输的 USB 线连接 Kindle，解锁设备并允许文件访问。
2. 在文件资源管理器中确认 Kindle 盘符，例如 `E:\`。
3. 从 Release 下载所需 exe，并放到一个可写的本地目录。
4. 可选：把 `settings.example.json` 复制为 `settings.json`，修改
   `kindle_path` 和 `output_dir`。
5. 运行 exe。程序扫描期间不会写入、删除、移动或修改 Kindle 文件。
6. 完成后打开配置输出目录中的 `kindle_reading_dashboard.html`。

默认输出包括：

- `kindle_reading_live.csv`
- `kindle_reading_dashboard.html`
- `kindle_reading_dashboard.xlsx`
- `reading_history.json`
- `achievement_profile.json`
- `_covers/`、`_wordcount/`、`_ocr/` 等缓存目录（按实际功能启用）

## 校验下载

在 PowerShell 中进入下载目录，将 Release 的 SHA-256 与本地文件对比：

```powershell
Get-FileHash .\KindleReader-Text-v1.13.0.exe -Algorithm SHA256
Get-FileHash .\KindleReader-OCR-v1.13.0.exe -Algorithm SHA256
```

把结果与 `SHA256SUMS.txt` 中对应文件名比较。不要在哈希不一致时运行文件。

当前 exe 未做商业代码签名，Windows SmartScreen 可能显示未知发布者提示。请只从
本项目 GitHub Releases 下载，并使用 `SHA256SUMS.txt` 校验。

## 功能

- 从 Kindle KRDS v1 sidecar 读取阅读时长、进度、最后阅读时间和计时模型数据。
- 生成可离线打开的响应式 HTML 看板，支持中英文、深色模式、筛选、排序和打印。
- 输出 UTF-8 BOM CSV，便于 Excel 直接打开；同时生成真实 XLSX 工作簿。
- 使用历史快照展示趋势、连续阅读、周期统计和成就进度。
- 可区分 Kindle 原始值、Kindle 外推值、正文精确词数和 OCR 估算值。
- 文字版支持本地图书封面缩略图；封面以内嵌 data URL 写入 HTML。
- 漫画版支持漫画/文字书分类、源文件页数、已读页数、进度、阅读时长和 KFX
  位置映射。
- OCR 仅作为显式可选的估算路径，结果不会冒充 Kindle 原始值或正文精确值。
- 在 Kindle 未连接或缺少某项数据时明确显示“未知”或“无时长数据”，不伪造 0。

## 设置

可复制的模板是 [`settings.example.json`](settings.example.json)。

常用字段：

| 字段 | 说明 |
| --- | --- |
| `kindle_path` | Kindle 根目录，例如 `E:\\` |
| `output_dir` | 报告、历史和缓存输出目录 |
| `language` | `auto`、`zh`、`ja` 或 `en` |
| `manga_classify` | 是否启用漫画/文字书分类 |
| `manga_ocr` | 是否显式启用 OCR；默认关闭 |
| `ocr_mode` | `classify`、`sample` 或 `full` |
| `exact_word_count` | 是否用 Calibre 转换正文并统计精确词数 |
| `open_html` | 完成后是否打开 HTML 看板 |

书籍模式模板是 [`book_modes.example.json`](book_modes.example.json)。可将示例复制
为 `book_modes.json`，通过合成 ID 或标题登记 `book` / `manga`。真实映射文件可能
包含个人信息，已由 `.gitignore` 排除。

## 漫画版依赖

漫画分类的显式映射和部分轻量判断不要求 OCR。以下组件仅在对应功能实际启用时
使用，**不会随 exe 一起分发**：

- **Calibre**：KFX 正文转换、元数据和 KFX 位置映射的本地支持。
- **Poppler**：`pdfinfo` 读取页数，`pdftoppm` 渲染页面。
- **PaddleOCR + PaddlePaddle**：中文/英文，以及可选的日文 OCR。
- **manga-ocr**：日文漫画 OCR 优先方案，通常还需要 PyTorch 与 Transformers。
- **Tesseract + `eng` 训练数据**：可选的英文 OCR 后端。

OCR 模型和大型运行时需要用户自行安装和准备。程序不会自动联网安装依赖、下载模型
或绕过 DRM。可通过以下环境变量指向明确的本地可执行文件：

```powershell
$env:KINDLE_OCR_EBOOK_CONVERT = "C:\Tools\Calibre\ebook-convert.exe"
$env:KINDLE_OCR_PDFTOPPM = "C:\Tools\Poppler\pdftoppm.exe"
$env:KINDLE_OCR_PDFINFO = "C:\Tools\Poppler\pdfinfo.exe"
$env:KINDLE_OCR_TESSERACT = "C:\Tools\Tesseract\tesseract.exe"
$env:KINDLE_OCR_PADDLEOCR_MODEL_DIR = "D:\Models\paddleocr"
```

推荐把 `.ocr-pip/` 和 `_ocr_models/` 放在 exe 同目录，然后保持离线运行。缺少
依赖时，相关书籍会记录明确的缺失状态，其他书籍仍会继续处理。

## 从源码运行

要求 Python 3.12：

```powershell
python -m pip install -r requirements.txt
python -X utf8 .\kindle_report.py --path E:\ --output-csv .\dist\kindle_reading_live.csv
```

也可以复制示例设置后运行：

```powershell
Copy-Item .\settings.example.json .\settings.json
python -X utf8 .\kindle_report.py
```

构建两个单文件 exe：

```powershell
python -m pip install -r requirements-build.txt
python -X utf8 .\kindle_report.py --build-exe
```

## 测试

```powershell
python -X utf8 -m py_compile kindle_reader.py kindle_report.py kindle_manga_ocr.py
python -X utf8 -m unittest test_kindle_reader -v
```

v1.13.0 的回归基线为 77 项测试。新增改动不得减少测试数量或弱化断言。

## 隐私与只读原则

- 应用不登录 Amazon，不读取云端账号数据，不要求令牌或 Cookie。
- 应用只读取挂载 Kindle 的 sidecar、电子书和本地 `fmcache.db`。
- 应用不写入、删除、移动或修改 Kindle 上的任何文件。
- 报告、历史和缓存只写入本机输出目录。
- 输出可能包含书名、作者、阅读记录、本地路径和封面。发布或分享前应自行脱敏。
- 仓库中的五册 CSV/HTML 是合成夹具，不是真实阅读记录。

项目不会提交 `reading_history.json`、`achievement_profile.json`、
`word_counts_cache.json`、`settings.json`、`book_modes.json`、封面、OCR 页面、
OCR 结果、模型、设备标识或 Calibre 元数据。清理范围见
[`docs/PRIVACY_CLEANUP.md`](docs/PRIVACY_CLEANUP.md)。

## 数据口径

| 数据 | 来源 | 说明 |
| --- | --- | --- |
| 阅读时长、进度、最后阅读时间 | Kindle sidecar / `fmcache.db` | 设备记录值；缺失时保持未知 |
| Kindle 词数 | sidecar `totalWords` / `numberOfWords` | 反映 Kindle 的统计口径 |
| 正文精确词数 | Calibre + jieba / Janome / 英文正则 | 取决于转换和分词是否成功 |
| OCR 词数 | PaddleOCR / manga-ocr / Tesseract | 估算值，明确标记，不替代精确值 |
| 漫画页数 | `pdfinfo` 读取源文件 | 可能与 Kindle 内部分页不同 |
| 阅读时长缺失 | 不计算 | 显示“无时长数据”，不伪造 0 |

## 已知限制

- 当前主要针对 Kindle Paperwhite 4、固件 5.16.3.1 和 KRDS v1 sidecar 验证，
  不对所有 Kindle 型号或固件作保证。
- 不支持 Amazon 云端阅读统计，也不处理加密电子书正文。
- 当前发布平台为 Windows 10 / 11 x64，未提供 macOS 或 Linux 安装包。
- KFX 位置映射依赖本地 Calibre 与 KFX Input 支持。
- 漫画 OCR 依赖额外工具和模型；离线模型缺失时相关 OCR 不会运行。
- OCR 是估算，可能受页面裁剪、气泡文字、语言判断、模型版本和低置信度影响。
- 源文件页数与 Kindle 内部分页可能不同。
- 单文件 exe 启动时需要解压到临时目录，首次启动可能较慢。
- 应用不会绕过 DRM，也不会判断用户是否有权访问某本电子书。

## 贡献、安全与许可证

- 贡献说明：[`CONTRIBUTING.md`](CONTRIBUTING.md)
- 安全报告：[`SECURITY.md`](SECURITY.md)
- 变更记录：[`CHANGELOG.md`](CHANGELOG.md)
- 第三方组件与许可证：[`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md)

本项目以 GNU General Public License version 3 发布，见 [`LICENSE`](LICENSE)。
Kindle、Amazon、Calibre 及其他第三方名称归各自权利人所有；本项目与这些权利人
没有隶属或背书关系。
