# Privacy cleanup record

Release: `v1.13.0`

Audit date: 2026-11-27

## Excluded from the public repository

- `reading_history.json`
- `achievement_profile.json`
- `settings.json`
- `book_modes.json`
- `word_counts_cache.json`
- `metadata.calibre`
- `_ocr/`, `_covers/`, `_wordcount/` and `_ocr_models/`
- `.ocr-pip/`, `.tmp-pip/`, local pip caches and virtual environments
- `.codex-backup-*`, Calibre test homes and build/test caches
- `build/`, `dist/` and generated XLSX reports
- nested Git metadata under `third_party/calibre-kfx-input/.git`

## Fixture migration

The five-book regression CSV originally contained real titles, authors, Kindle
asset identifiers, storage identifiers and device-side paths. It was replaced by
a five-row all-synthetic fixture with:

- synthetic titles and author names;
- synthetic 32-character asset/storage identifiers;
- generic `E:\documents` example paths;
- the original column order and numeric semantics;
- preserved English, Chinese, Japanese and manga test coverage.

The root HTML dashboard was rebuilt from the synthetic CSV. Existing personal
covers, OCR pages, OCR results, history snapshots and Calibre metadata were not
copied into the release workspace.

## Pre-publication checks

- Added ignore rules for local data, caches, models and build output.
- Removed machine-specific Python paths from user-facing scripts and
  documentation.
- Scanned publishable text and source for tokens, cookies, API keys, personal
  absolute paths, real fixture titles and the original Kindle identifiers.
- Verified no executable, model, dependency bundle, ebook or cover is committed.
- Verified release checksums after the final Windows build.

This record contains no personal reading data or credentials.
