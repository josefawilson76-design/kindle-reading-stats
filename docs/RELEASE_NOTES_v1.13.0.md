# Kindle Reading Dashboard v1.13.0 | Text + Manga

Released 2026-11-27 for Windows 10 / 11 x64.

## Downloads

| Edition | File | Scope |
| --- | --- | --- |
| Text-only | `KindleReader-Text-v1.13.0.exe` | Text books, dashboard, history and achievements; no manga or OCR module |
| Text + Manga | `KindleReader-OCR-v1.13.0.exe` | Adds manga classification, page/progress reporting, KFX position mapping and optional OCR |

Also attached: `SHA256SUMS.txt`, `settings.example.json` and
`book_modes.example.json`.

## 中文说明

本版本提供两个 Windows 单文件程序：

- **纯文字版**：只处理文字书统计，不加载 OCR 模块，也不提供漫画分类界面。
- **文字 + 漫画版**：在文字书功能之外，支持漫画分类、源文件页数、已读页数、
  进度、阅读时长和 KFX 位置映射。OCR 默认关闭，只有显式启用时才作为估算值运行。

漫画 OCR 不包含在 exe 中。需要时请自行准备 Calibre、Poppler、PaddleOCR、
PaddlePaddle 或 manga-ocr，以及对应的离线模型。程序不会自动下载依赖或模型。

阅读时长缺失时保持“无时长数据”，不会伪造 0 或写入估算时长。OCR 词数不会冒充
Kindle 原始值或正文精确值。

下载后请使用 `SHA256SUMS.txt` 核对文件。exe 未做商业代码签名，SmartScreen 可能
显示未知发布者提示。

## English notes

This release provides two Windows single-file executables:

- **Text-only** processes text books and does not load the OCR module or expose
  manga classification UI.
- **Text + Manga** adds manga classification, source-file page count, pages read,
  progress, duration and KFX position mapping. OCR is off by default and runs
  only when explicitly enabled as an estimate.

Manga OCR is not embedded in the executable. When needed, install Calibre,
Poppler, PaddleOCR, PaddlePaddle or manga-ocr and prepare offline model files
separately. The application does not download dependencies or models.

Missing reading duration remains unavailable; the app does not fabricate zero or
estimated duration. OCR word counts are never presented as Kindle values or exact
text-derived counts.

Verify every download against `SHA256SUMS.txt`. The executables are not
commercially code-signed, so Windows SmartScreen may show an unknown publisher
warning.

## Checksums

See `SHA256SUMS.txt` in this Release.

## Privacy

No Amazon login, cloud account, token or cookie is required. The Kindle remains
read-only. Reports and caches are written only to the selected local directory.
The public repository and test fixtures use synthetic data.
