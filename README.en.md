# Kindle Reading Dashboard

A Windows desktop tool that reads a USB-mounted Kindle without writing to the
device, then builds local CSV, HTML and XLSX reports for reading time, progress,
WPM, word counts, manga pages and historical trends.

[中文说明](README.md) | [Project site](https://josefawilson76-design.github.io/kindle-reading-stats/) | [GitHub Releases](https://github.com/josefawilson76-design/kindle-reading-stats/releases)

**Version:** v1.13.0

**Release date:** 2026-11-27

**Platform:** Windows 10 / 11 x64

**License:** GPL v3

## Download

| Edition | Best for | Download |
| --- | --- | --- |
| Text-only | Text books, dashboards, history and achievements; does not load OCR or classify manga | [KindleReader-Text-v1.13.0.exe](https://github.com/josefawilson76-design/kindle-reading-stats/releases/download/v1.13.0/KindleReader-Text-v1.13.0.exe) |
| Text + Manga | Adds manga classification, page/progress reporting, KFX position mapping and optional OCR | [KindleReader-OCR-v1.13.0.exe](https://github.com/josefawilson76-design/kindle-reading-stats/releases/download/v1.13.0/KindleReader-OCR-v1.13.0.exe) |

Both editions share the same reporting core. The text-only edition **does not
support manga** and does not load the OCR module. The text + manga edition adds
manga workflows and optional OCR.

The Release also includes:

- [`SHA256SUMS.txt`](https://github.com/josefawilson76-design/kindle-reading-stats/releases/download/v1.13.0/SHA256SUMS.txt)
- `settings.example.json`
- `book_modes.example.json`

## Quick start

1. Connect the Kindle over a USB data cable, unlock it and allow file access.
2. Confirm the mounted drive letter, for example `E:\`.
3. Download the required executable from the Release and place it in a writable
   local folder.
4. Optional: copy `settings.example.json` to `settings.json` and set
   `kindle_path` and `output_dir`.
5. Run the executable. The app never writes to, deletes from, moves or modifies
   files on the Kindle.
6. Open `kindle_reading_dashboard.html` in the configured output directory.

Typical output includes CSV, HTML, a real XLSX workbook, history snapshots,
achievement data and caches for enabled features.

## Download and verify

After downloading, compare each local SHA-256 with the Release checksum:

```powershell
Get-FileHash .\KindleReader-Text-v1.13.0.exe -Algorithm SHA256
Get-FileHash .\KindleReader-OCR-v1.13.0.exe -Algorithm SHA256
```

Do not run a file whose hash differs from `SHA256SUMS.txt`.

The executables are not commercially code-signed, so Windows SmartScreen may
show an unknown publisher warning. Download only from this repository's GitHub
Releases and verify the checksum.

## What it does

- Reads reading time, progress, last-read time and timer data from Kindle KRDS
  v1 sidecars.
- Builds an offline responsive dashboard with Chinese/English UI, dark mode,
  filtering, sorting and print support.
- Exports an Excel-friendly UTF-8 BOM CSV and a real XLSX workbook.
- Tracks history, streaks, period summaries and achievement progress.
- Separates Kindle values, exact text-derived counts and OCR estimates.
- Embeds local cover thumbnails as data URLs without network resources.
- In the text + manga edition, reports source-file page count, pages read,
  progress, duration and KFX position mapping.
- Keeps missing duration as unknown instead of inventing zero or an estimate.

## Settings

Start from [`settings.example.json`](settings.example.json). The most useful
fields are:

| Field | Meaning |
| --- | --- |
| `kindle_path` | Kindle root, for example `E:\\` |
| `output_dir` | Folder for reports, history and caches |
| `language` | `auto`, `zh`, `ja` or `en` |
| `manga_classify` | Enable manga/text classification |
| `manga_ocr` | Explicit OCR opt-in; off by default |
| `ocr_mode` | `classify`, `sample` or `full` |
| `exact_word_count` | Convert text locally for exact word counts |
| `open_html` | Open the dashboard after a successful run |

[`book_modes.example.json`](book_modes.example.json) shows the mapping format.
Copy it to `book_modes.json` and map synthetic or real books by ID or title as
`book` or `manga`. A real mapping file may contain personal data and is ignored
by Git.

## Requirements

The text-only executable needs no Python installation.

Optional text and manga conversion can use:

- **[Calibre](https://calibre-ebook.com/download_windows64)** for KFX conversion,
  metadata and local KFX position support.
- **[Poppler Windows builds](https://github.com/oschwartz10612/poppler-windows/releases)**
  (`pdfinfo`, `pdftoppm`) for PDF page counts and rendering.
- **[PaddleOCR](https://github.com/PaddlePaddle/PaddleOCR/releases)** +
  **[PaddlePaddle](https://www.paddlepaddle.org.cn/install/quick)** for Chinese,
  English and optional Japanese OCR.
- **[manga-ocr](https://github.com/kha-white/manga-ocr)** for Japanese manga OCR.
  The offline base model is available from
  **[manga-ocr-base](https://huggingface.co/kha-white/manga-ocr-base)**.
- **[Tesseract Windows builds](https://github.com/UB-Mannheim/tesseract/wiki)** with
  **[tessdata](https://github.com/tesseract-ocr/tessdata)** as an optional
  English OCR backend.

These tools and OCR model files are not bundled. The app does not download
models, install dependencies, use network services or bypass DRM. You can point
to local executables with environment variables:

```powershell
$env:KINDLE_OCR_EBOOK_CONVERT = "C:\Tools\Calibre\ebook-convert.exe"
$env:KINDLE_OCR_PDFTOPPM = "C:\Tools\Poppler\pdftoppm.exe"
$env:KINDLE_OCR_PDFINFO = "C:\Tools\Poppler\pdfinfo.exe"
$env:KINDLE_OCR_TESSERACT = "C:\Tools\Tesseract\tesseract.exe"
$env:KINDLE_OCR_PADDLEOCR_MODEL_DIR = "D:\Models\paddleocr"
```

Project-side `.ocr-pip/` and `_ocr_models/` folders beside the executable are
supported for offline use. Missing dependencies are recorded per book; other
books continue to process.

## Run from source

Python 3.12 is required:

```powershell
python -m pip install -r requirements.txt
python -X utf8 .\kindle_report.py --path E:\ --output-csv .\dist\kindle_reading_live.csv
```

Build the Windows executables:

```powershell
python -m pip install -r requirements-build.txt
python -X utf8 .\kindle_report.py --build-exe
```

## Test

```powershell
python -X utf8 -m py_compile kindle_reader.py kindle_report.py kindle_manga_ocr.py
python -X utf8 -m unittest test_kindle_reader -v
```

The v1.13.0 regression baseline is 77 tests. Pull requests must not reduce the
test count or weaken assertions to pass.

## Privacy and read-only behavior

- No Amazon login, cloud account access, token or cookie is required.
- The app reads Kindle sidecars, ebooks and the local `fmcache.db` only.
- It does not write to, delete from, move or modify the Kindle.
- Reports and caches are written to a local folder on the computer.
- Reports can contain titles, authors, reading history, local paths and covers;
  review them before sharing.
- The five-book CSV/HTML files in this repository are synthetic fixtures.

The public repository excludes reading history, achievement profiles, word-count
caches, local settings, book-mode mappings, covers, OCR pages/results, models,
device identifiers and Calibre metadata. See
[`docs/PRIVACY_CLEANUP.md`](docs/PRIVACY_CLEANUP.md).

## Data semantics

| Data | Source | Notes |
| --- | --- | --- |
| Reading time, progress, last-read time | Kindle sidecar / `fmcache.db` | Device values; missing stays unknown |
| Kindle word count | Sidecar `totalWords` / `numberOfWords` | Kindle's own counting semantics |
| Exact text word count | Calibre plus jieba / Janome / English regex | Depends on successful conversion |
| OCR word count | PaddleOCR / manga-ocr / Tesseract | Explicit estimate; never substituted for exact data |
| Manga pages | `pdfinfo` on the source file | May differ from Kindle pagination |
| Missing duration | Not calculated | Displayed as unavailable, not as fabricated zero |

## Known limitations

- The main verified baseline is Kindle Paperwhite 4 with firmware 5.16.3.1 and
  KRDS v1 sidecars. Other models and firmware versions are not guaranteed.
- Amazon cloud reading statistics and encrypted ebook text are outside scope.
- Current release artifacts target Windows 10 / 11 x64. No macOS or Linux
  installer is provided.
- KFX position mapping depends on a local Calibre installation and KFX Input
  support.
- Manga OCR needs external tools and offline models.
- OCR is an estimate and can be affected by page rendering, speech bubbles,
  language detection, model version and low-confidence output.
- Source-file page count can differ from Kindle internal pagination.
- Single-file executables extract to a temporary folder and may take longer on
  first launch.
- The project does not bypass DRM or determine whether a user may access an
  ebook.

## License and third-party notices

- Contributing: [`CONTRIBUTING.md`](CONTRIBUTING.md)
- Security: [`SECURITY.md`](SECURITY.md)
- Changelog: [`CHANGELOG.md`](CHANGELOG.md)
- Third-party notices: [`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md)

This project is released under GPL v3. See [`LICENSE`](LICENSE). Kindle, Amazon,
Calibre and other product names belong to their respective owners. This project
is not affiliated with or endorsed by those owners.
