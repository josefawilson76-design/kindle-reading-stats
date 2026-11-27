# Third-Party Notices

This project is licensed under GPL v3. The notices below cover bundled source
and optional external components. Optional components are not downloaded,
installed or redistributed by this repository.

## Bundled source

### Calibre KFX Input plugin

- Path: `third_party/calibre-kfx-input/`
- Upstream: John Howell's KFX Input plugin, mirrored from
  `https://github.com/kluyg/calibre-kfx-input`
- Source revision: `44db6b6ee8c0094a98c33770575a9070ddb90fda`
- License: GNU GPL v3; full text in
  `third_party/calibre-kfx-input/LICENSE`

The plugin is bundled as source to support local KFX position mapping. It is
not a Calibre executable distribution and does not include user library data.

## Libraries included in the Windows executables

| Component | Purpose | License |
| --- | --- | --- |
| Python 3.12 | Runtime | PSF License |
| PyInstaller 6.22.3 | Single-file executable builder | GPL-2.0-or-later with bootloader exception |
| Pillow | Cover thumbnails | MIT-CMU / HPND |
| jieba 0.42.1 | Chinese tokenization | MIT |
| Janome 0.5.0 | Japanese tokenization | Apache-2.0 |
| openpyxl | XLSX export | MIT |

The exact dependency metadata is reproduced in the installed distributions.

## Optional external tools

These tools are user-installed and are not shipped in the executable files:

| Component | Purpose | Upstream license |
| --- | --- | --- |
| Calibre | KFX conversion and metadata | GPL-3.0 |
| Poppler (`pdfinfo`, `pdftoppm`) | PDF page count and rendering | GPL-2.0-or-later |
| Tesseract OCR | English OCR | Apache-2.0 |
| PaddleOCR | Chinese, English and Japanese OCR | Apache-2.0 |
| PaddlePaddle | PaddleOCR runtime | Apache-2.0 |
| manga-ocr | Japanese manga OCR | Apache-2.0 |
| PyTorch | manga-ocr runtime | BSD-3-Clause |
| Transformers | manga-ocr model runtime | Apache-2.0 |
| NumPy | Numeric runtime | BSD-3-Clause |

No OCR model weights, Calibre or Poppler binaries, Tesseract data, fonts or
ebooks are included in this repository or the release executables. Users are
responsible for installing optional tools and model files under their upstream
licenses and keeping them offline if desired.

## Trademarks

Kindle and Amazon are trademarks of Amazon.com, Inc. or its affiliates. Calibre
and the names of other third-party products are trademarks of their respective
owners. This project is not affiliated with or endorsed by those owners.
