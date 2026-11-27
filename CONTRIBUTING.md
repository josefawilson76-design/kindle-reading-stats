# Contributing

## Before opening a pull request

1. Use Windows 10 or 11 x64 and Python 3.12 for the release-compatible test
   environment.
2. Install source dependencies with `python -m pip install -r requirements.txt`.
3. Run the full test suite and syntax check:

   ```powershell
   python -X utf8 -m unittest test_kindle_reader -v
   python -X utf8 -m py_compile kindle_reader.py kindle_report.py kindle_manga_ocr.py
   ```

4. Keep the 77-test baseline or increase it. Do not weaken assertions to make a
   change pass.
5. Describe user-visible behavior, privacy impact and any new dependency in the
   pull request.

## Test data

- Use synthetic titles, authors, identifiers, paths and reading values.
- Never commit a real Kindle ASIN or storage ID, device serial, account ID,
  Calibre metadata, book cover, OCR page, reading history or personal path.
- The files `kindle_reading_live.csv` and `kindle_reading_dashboard.html` are
  five-book regression fixtures. They must remain synthetic and semantically
  compatible.

## Third-party code and models

- Do not commit ebooks, OCR model weights, PaddleOCR/manga-ocr wheels, PyTorch,
  Calibre binaries, Poppler binaries, fonts or other large dependency bundles.
- New source dependencies must have a compatible license and be added to
  `THIRD_PARTY_NOTICES.md`.
- Preserve license and attribution files for bundled source.

## Style and scope

- Keep changes focused and preserve CSV append-only compatibility.
- Do not describe estimates as exact values.
- Do not claim support for a Kindle model or firmware version that has not been
  tested.
- Update `CHANGELOG.md` for user-facing changes.

By contributing, you agree that your contribution may be distributed under the
GNU General Public License version 3.
