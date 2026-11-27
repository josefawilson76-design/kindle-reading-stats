# Changelog

All notable changes are documented here. This project follows semantic
versioning for user-facing releases.

## [1.13.0] - 2026-11-27

### Added

- Public Windows x64 builds for the text-only and text + manga editions.
- Offline HTML dashboard, CSV, real XLSX workbook and optional history output.
- Manga classification, page/progress reporting, KFX position mapping and
  explicit optional OCR estimates.
- Bilingual README, bilingual GitHub Pages download page and release notes.
- Reproducible example settings and book-mode mappings.
- SHA-256 checksums for every released executable and configuration asset.

### Changed

- Cleaned the public fixture to five clearly synthetic books while preserving
  the original regression semantics.
- Removed machine-specific configuration paths and excluded local Kindle data,
  covers, OCR results, models, caches and build artifacts from source control.
- Documented the distinction between Kindle values, exact word counts and OCR
  estimates.

### Fixed

- The text-only build filters manga rows, uses four text-book achievement
  tracks and hides every manga/OCR shelf, filter and achievement control.
- The text-only CLI help no longer exposes manga or OCR options.
- Missing reading duration remains unknown instead of being rendered as a
  fabricated zero or estimated value.

### Release assets

- `KindleReader-Text-v1.13.0.exe`
- `KindleReader-OCR-v1.13.0.exe`
- `SHA256SUMS.txt`
- `settings.example.json`
- `book_modes.example.json`

## [1.12.0] - 2026-10-04

- Internal development baseline for manga page reporting, KFX position mapping,
  dashboard preferences and achievement history.
