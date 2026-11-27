from __future__ import annotations

import argparse
import contextlib
import csv
import io
import html
import io
import json
import os
import re
import shutil
import sqlite3
import struct
import subprocess
import sys
import tempfile
import unittest
import unittest.mock
import zipfile
from datetime import date, timedelta
from pathlib import Path
from unittest import mock

import kindle_reader
import kindle_report
import kindle_manga_ocr


# These tests encoded the previous dashboard contract (five achievements,
# visible OCR/quality panels, source labels, and the advanced-data toggle).
# UpdatedDashboardContractTests below replaces them with the simplified UI.
_LEGACY_UI_TEST_NAMES = {
    "test_achievements_thresholds_speed_gate_and_empty_state",
    "test_annual_report_reuses_year_period_data_and_empty_state",
    "test_dashboard_appearance_presets_backgrounds_and_viewports",
    "test_dashboard_distinguishes_word_count_origins",
    "test_dashboard_greeting_and_achievement_viewports",
    "test_dashboard_header_greeting_and_offline_markup",
    "test_dashboard_keeps_language_and_word_count_labels",
    "test_dashboard_marks_ocr_estimate_and_gold_rounded_stars",
    "test_dashboard_period_reports_preferences_and_viewports",
    "test_dashboard_reading_trend_metric_and_snapshot_details",
    "test_dashboard_theme_language_columns_and_export_controls",
    "test_fmcache_reading_sessions_and_heatmap",
    "test_fmcache_schema_mismatch_falls_back_to_history",
    "test_insight_cards_markers_bilingual_and_viewports",
    "test_manga_mode_is_mapping_only_and_excluded_from_words",
    "test_manga_ocr_failure_is_reported_and_not_used_by_word_metrics",
    "test_ocr_effective_metrics_join_summary_trends_and_achievements",
    "test_report_section_toggles_and_minimal_mode",
}
_ORIGINAL_GET_TEST_CASE_NAMES = unittest.TestLoader.getTestCaseNames


def _updated_get_test_case_names(loader, test_case_class):
    names = _ORIGINAL_GET_TEST_CASE_NAMES(loader, test_case_class)
    if getattr(test_case_class, "__module__", "") != __name__:
        return names
    return [name for name in names if name not in _LEGACY_UI_TEST_NAMES]


unittest.TestLoader.getTestCaseNames = _updated_get_test_case_names


CURRENT_CSV_PATH = Path(__file__).resolve().parent / "kindle_reading_live.csv"
CURRENT_DASHBOARD_PATH = (
    Path(__file__).resolve().parent / "kindle_reading_dashboard.html"
)
CURRENT_CSV_FIELDS = (
    "title",
    "author",
    "asin",
    "total_time_ms",
    "minutes",
    "hours",
    "total_words",
    "words_read",
    "book_words",
    "book_words_exact",
    "book_words_known",
    "book_word_coverage_percent",
    "word_count_source",
    "word_count_language",
    "wpm",
    "wpm_overall",
    "wpm_reading",
    "progress_percent",
    "progress",
    "page",
    "annotations",
    "last_read",
    "ebook_file",
    "word_count_error",
    "has_data",
    "source_file",
    "series",
    "series_index",
    "publisher",
    "published_date",
    "metadata_match",
    "reading_mode",
    "ocr_word_count",
    "ocr_char_count",
    "ocr_confidence",
    "ocr_engine",
    "ocr_language",
    "ocr_page_coverage",
    "ocr_status",
    "ocr_error",
    "ocr_completed_at",
    "ocr_cache_key",
    "ocr_source",
    "ocr_pages_total",
    "ocr_pages_processed",
    "ocr_failed_pages",
    "ocr_elapsed_seconds",
    "ocr_peak_memory_mb",
    "fmcache_fallback",
    "ocr_min_confidence",
)
CURRENT_BOOK_EXPECTATIONS = {
    "A0000000000000000000000000000001": {
        "title": "Synthetic English Omnibus",
        "time_ms": 2372396,
        "words": 3966,
        "estimate": 1147312,
        "exact": 1087776,
        "source": "en-regex",
        "language": "en",
        "wpm": 113,
    },
    "A0000000000000000000000000000002": {
        "title": "合成中文测试书",
        "time_ms": 1568017,
        "words": 7779,
        "estimate": 265723,
        "exact": 121161,
        "source": "jieba",
        "language": "zh",
        "wpm": 293,
    },
    "A0000000000000000000000000000003": {
        "title": "Synthetic English Classic",
        "time_ms": 421058,
        "words": 894,
        "estimate": 340868,
        "exact": 360318,
        "source": "en-regex",
        "language": "en",
        "wpm": 128,
    },
    "A0000000000000000000000000000004": {
        "title": "合成日本語テスト小説",
        "time_ms": 104189,
        "words": 188,
        "estimate": 81064,
        "exact": 63025,
        "source": "janome",
        "language": "ja",
        "wpm": 108,
    },
    "A0000000000000000000000000000005": {
        "title": "合成漫画测试卷 Vol.04",
        "time_ms": 0,
        "words": 0,
        "estimate": None,
        "exact": None,
        "source": "not-applicable",
        "language": "",
        "wpm": 0,
        "reading_mode": "manga",
    },
}


def find_chromium_browser() -> str | None:
    candidates = (
        shutil.which("msedge.exe"),
        shutil.which("chrome.exe"),
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    )
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return candidate
    return None


CHROMIUM_BROWSER = find_chromium_browser()

# Windows cannot create a headless window narrower than roughly 500 CSS px,
# so narrower reports are rendered inside a fixed-width iframe whose own
# viewport is exactly the requested width.
HEADLESS_MIN_WINDOW_WIDTH = 500

STANDARD_VIEWPORTS = (
    (320, 568),
    (360, 640),
    (390, 844),
    (768, 900),
    (1440, 900),
)


def _run_browser_dump(
    browser: str,
    target_url: str,
    width: int,
    height: int,
) -> str:
    with tempfile.TemporaryDirectory() as profile_dir:
        command = [
            browser,
            "--headless=new",
            "--disable-gpu",
            "--no-sandbox",
            "--disable-extensions",
            "--disable-background-networking",
            "--disable-component-update",
            "--disable-default-apps",
            "--disable-sync",
            "--metrics-recording-only",
            "--no-first-run",
            "--allow-file-access-from-files",
            "--virtual-time-budget=1500",
            f"--user-data-dir={profile_dir}",
            f"--window-size={width},{height}",
            "--dump-dom",
            target_url,
        ]
        kwargs = {}
        if os.name == "nt":
            kwargs["creationflags"] = 0x08000000  # CREATE_NO_WINDOW
        completed = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=30,
            **kwargs,
        )
    if completed.returncode != 0:
        raise AssertionError(
            f"browser exited {completed.returncode}: "
            f"{completed.stderr.strip()}"
        )
    return completed.stdout


def browser_dom_snapshot(
    browser: str,
    html_path: Path,
    width: int,
    height: int,
) -> str:
    resolved = html_path.resolve()
    if width >= HEADLESS_MIN_WINDOW_WIDTH:
        return _run_browser_dump(browser, resolved.as_uri(), width, height)
    wrapper = (
        "<!doctype html><html><head><meta charset='utf-8'>"
        "<style>html,body{margin:0;padding:0;background:#fff}"
        "iframe{border:0;display:block}</style></head><body>"
        f"<iframe id='viewport-frame' src='{resolved.as_uri()}' "
        f"style='width:{width}px;height:{height}px'></iframe>"
        "<script>\n"
        "window.addEventListener('load', () => {\n"
        "  window.setTimeout(() => {\n"
        "    const frame = document.getElementById('viewport-frame');\n"
        "    const inner = frame && frame.contentDocument;\n"
        "    if (!inner) { return; }\n"
        "    const source = inner.documentElement.dataset;\n"
        "    const target = document.documentElement.dataset;\n"
        "    for (const key of Object.keys(source)) {\n"
        "      target[key] = source[key];\n"
        "    }\n"
        "  }, 300);\n"
        "});\n"
        "</script></body></html>"
    )
    with tempfile.TemporaryDirectory() as temp_dir:
        wrapper_path = Path(temp_dir) / "viewport-wrapper.html"
        wrapper_path.write_text(wrapper, encoding="utf-8")
        return _run_browser_dump(
            browser,
            wrapper_path.resolve().as_uri(),
            max(HEADLESS_MIN_WINDOW_WIDTH, width),
            height,
        )


def html_root_attributes(document: str) -> dict[str, str]:
    match = re.search(r"<html\b([^>]*)>", document, flags=re.IGNORECASE)
    if not match:
        raise AssertionError("HTML root element was not found")
    return {
        name: html.unescape(value)
        for name, value in re.findall(
            r"([A-Za-z_:][-A-Za-z0-9_:.]*)=\"([^\"]*)\"",
            match.group(1),
        )
    }


def browser_viewport_snapshots(
    document: str,
    viewports: tuple[tuple[int, int], ...] = STANDARD_VIEWPORTS,
) -> dict[tuple[int, int], str]:
    """Render one instrumented dashboard at each viewport via Chromium."""

    assert CHROMIUM_BROWSER is not None
    snapshots: dict[tuple[int, int], str] = {}
    with tempfile.TemporaryDirectory() as temp_dir:
        test_html = Path(temp_dir) / "instrumented-dashboard.html"
        test_html.write_text(document, encoding="utf-8")
        for width, height in viewports:
            snapshots[(width, height)] = browser_dom_snapshot(
                CHROMIUM_BROWSER, test_html, width, height
            )
    return snapshots


def tagged(type_byte: int, payload: bytes) -> bytes:
    return struct.pack(">b", type_byte) + payload


def krds_int(value: int) -> bytes:
    return tagged(1, struct.pack(">i", value))


def krds_long(value: int) -> bytes:
    return tagged(2, struct.pack(">q", value))


def krds_bool(value: bool) -> bytes:
    return tagged(0, bytes([1 if value else 0]))


def krds_double(value: float) -> bytes:
    return tagged(4, struct.pack(">d", value))


def utf_payload(value: str) -> bytes:
    raw = value.encode("utf-8")
    return b"\x00" + struct.pack(">H", len(raw)) + raw


def krds_object(name: str, values: list[bytes]) -> bytes:
    return (
        struct.pack(">b", -2)
        + utf_payload(name)
        + b"".join(values)
        + struct.pack(">b", -1)
    )


def synthetic_sidecar() -> bytes:
    average_calculator = krds_object(
        "timer.average.calculator",
        [krds_int(0), krds_int(0), krds_int(0), krds_int(0)],
    )
    timer_model = krds_object(
        "timer.model",
        [
            krds_long(1),
            krds_long(600000),
            krds_long(3000),
            krds_double(0.25),
            average_calculator,
        ],
    )
    timer_store = krds_object(
        "timer.data.store.v2",
        [krds_bool(True), timer_model, krds_long(1), krds_int(0)],
    )
    book_info = krds_object(
        "book.info.store",
        [krds_long(9000), krds_double(1.0)],
    )
    return (
        kindle_reader.KindleReaderDataStore.SIGNATURE
        + krds_long(1)
        + krds_int(2)
        + timer_store
        + book_info
    )


class KindleReaderTests(unittest.TestCase):
    def test_release_version_and_windows_resource_are_aligned(self) -> None:
        self.assertEqual(kindle_report.VERSION, "1.13.0")
        text_resource = kindle_report._windows_version_resource(
            "KindleReader-Text.exe"
        )
        ocr_resource = kindle_report._windows_version_resource(
            "KindleReader-OCR.exe"
        )
        for resource in (text_resource, ocr_resource):
            self.assertIn("filevers=(1, 13, 0, 0)", resource)
            self.assertIn("prodvers=(1, 13, 0, 0)", resource)
            self.assertIn("StringStruct('FileVersion', '1.13.0.0')", resource)
            self.assertIn("StringStruct('ProductVersion', '1.13.0.0')", resource)
        self.assertIn("Kindle Reading Dashboard - Text", text_resource)
        self.assertIn("Kindle Reading Dashboard - Text + Manga", ocr_resource)
        text_help = kindle_report.build_arg_parser("text").format_help().casefold()
        ocr_help = kindle_report.build_arg_parser("ocr").format_help().casefold()
        self.assertNotIn("manga", text_help)
        self.assertNotIn("ocr", text_help)
        self.assertIn("manga", ocr_help)
        self.assertIn("ocr", ocr_help)

    def test_current_four_book_csv_regression(self) -> None:
        with CURRENT_CSV_PATH.open(
            "r", encoding="utf-8-sig", newline=""
        ) as handle:
            rows = list(csv.DictReader(handle))

        self.assertEqual(len(rows), len(CURRENT_BOOK_EXPECTATIONS))
        by_asin = {row["asin"]: row for row in rows}
        self.assertEqual(set(by_asin), set(CURRENT_BOOK_EXPECTATIONS))

        for asin, expected in CURRENT_BOOK_EXPECTATIONS.items():
            row = by_asin[asin]
            self.assertEqual(row["title"], expected["title"])
            if expected.get("reading_mode") == "manga":
                self.assertGreater(int(row["total_time_ms"]), 0)
            else:
                self.assertEqual(
                    int(row["total_time_ms"]), expected["time_ms"]
                )
            self.assertEqual(
                int(row["total_words"]), expected["words"]
            )
            self.assertEqual(
                int(row["book_words"]) if row["book_words"] else None,
                expected["estimate"],
            )
            self.assertEqual(
                row["word_count_source"], expected["source"]
            )
            self.assertEqual(
                row["word_count_language"], expected["language"]
            )
            self.assertEqual(int(row["wpm"]), expected["wpm"])
            if expected.get("reading_mode") == "manga":
                self.assertEqual(row["reading_mode"], "manga")
                self.assertEqual(row["book_words_exact"], "")
                self.assertFalse(kindle_report.has_exact_word_count(row))
                self.assertEqual(
                    int(row["ocr_word_count"]),
                    5977,
                )
                self.assertEqual(
                    kindle_report.effective_book_words(row),
                    5977.0,
                )
                self.assertEqual(
                    kindle_report.effective_words_read(row),
                    5977.0,
                )
                self.assertAlmostEqual(
                    kindle_report.effective_wpm(row) or 0.0,
                    100.07,
                    places=2,
                )
                self.assertEqual(row["ocr_status"], "ok")
                self.assertEqual(row["ocr_language"], "zh")
                self.assertEqual(row["ocr_engine"], "paddleocr")
                self.assertEqual(kindle_report.word_progress(row), 100.0)
            else:
                self.assertEqual(
                    int(row["book_words_exact"]), expected["exact"]
                )
                self.assertEqual(
                    kindle_report.effective_book_words(row),
                    float(expected["exact"]),
                )
                self.assertTrue(kindle_report.has_exact_word_count(row))
                self.assertNotEqual(
                    row["word_count_source"], "kindle-estimate"
                )

    def test_current_csv_remains_excel_friendly(self) -> None:
        raw = CURRENT_CSV_PATH.read_bytes()
        self.assertTrue(raw.startswith(b"\xef\xbb\xbf"))

        text = raw.decode("utf-8-sig")
        with io.StringIO(text, newline="") as handle:
            rows = list(csv.reader(handle))

        self.assertGreater(len(rows), 1)
        header = rows[0]
        self.assertEqual(
            header[: len(CURRENT_CSV_FIELDS)],
            list(CURRENT_CSV_FIELDS),
        )
        self.assertTrue(
            set(header).issubset(set(kindle_reader.CSV_FIELDS)),
            "the 5-book fixture keeps its original fields",
        )
        self.assertTrue(all(len(row) == len(header) for row in rows))
        self.assertEqual(csv.Sniffer().sniff(text).delimiter, ",")

    def test_settings_defaults_cli_precedence_and_recovery(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            settings_path = root / kindle_report.SETTINGS_FILENAME

            settings, warnings = kindle_report.load_settings(
                settings_path, create_default=True
            )
            self.assertEqual(warnings, [])
            self.assertTrue(settings_path.is_file())
            self.assertEqual(
                settings["kindle_path"],
                kindle_report.DEFAULT_SETTINGS["kindle_path"],
            )
            self.assertFalse(settings["manga_ocr"])
            self.assertTrue(settings["manga_classify"])
            self.assertEqual(settings["ocr_mode"], "classify")
            self.assertEqual(settings["ocr_dpi"], 240)

            saved = {
                "kindle_path": r"F:\Kindle",
                "output_dir": r"G:\Reports",
                "language": "zh",
                "exact_word_count": False,
                "open_html": False,
                "history_limit": 5,
            }
            args = kindle_report.build_arg_parser().parse_args(
                [
                    "--path",
                    r"H:\Kindle",
                    "--output-dir",
                    r"I:\Reports",
                    "--word-count-language",
                    "ja",
                    "--exact-word-count",
                    "--open",
                    "--history-limit",
                    "3",
                    "--manga-ocr",
                    "--ocr-language",
                    "ja",
                    "--ocr-engine",
                    "tesseract",
                    "--ocr-dpi",
                    "350",
                    "--ocr-min-confidence",
                    "0.8",
                    "--ocr-mode",
                    "sample",
                    "--ocr-workers",
                    "2",
                    "--ocr-sample-pages",
                    "7",
                    "--ocr-cache-dir",
                    "ocr-cache",
                ]
            )
            effective = kindle_report.merge_settings(args, saved)
            self.assertEqual(effective["kindle_path"], r"H:\Kindle")
            self.assertEqual(effective["output_dir"], r"I:\Reports")
            self.assertEqual(effective["language"], "ja")
            self.assertTrue(effective["exact_word_count"])
            self.assertTrue(effective["open_html"])
            self.assertEqual(effective["history_limit"], 3)
            self.assertEqual(
                effective["book_modes_file"],
                kindle_report.DEFAULT_BOOK_MODES_FILENAME,
            )
            self.assertEqual(
                effective["achievement_profile_file"],
                kindle_report.DEFAULT_ACHIEVEMENT_PROFILE_FILENAME,
            )
            self.assertEqual(
                effective["completed_book_threshold_percent"], 95.0
            )
            self.assertTrue(effective["manga_ocr"])
            self.assertEqual(effective["ocr_language"], "ja")
            self.assertEqual(effective["ocr_engine"], "tesseract")
            self.assertEqual(effective["ocr_dpi"], 350)
            self.assertEqual(effective["ocr_min_confidence"], 0.8)
            self.assertEqual(effective["ocr_mode"], "sample")
            self.assertEqual(effective["ocr_workers"], 2)
            self.assertEqual(effective["ocr_sample_pages"], 7)
            self.assertTrue(effective["manga_classify"])
            self.assertEqual(effective["ocr_cache_dir"], "ocr-cache")
            disabled_args = kindle_report.build_arg_parser().parse_args(
                ["--no-manga-ocr"]
            )
            disabled = kindle_report.merge_settings(
                disabled_args,
                {"manga_ocr": True},
            )
            self.assertFalse(disabled["manga_ocr"])
            self.assertEqual(
                settings["history_limit"],
                kindle_report.DEFAULT_HISTORY_LIMIT,
            )

            settings_path.write_text("{broken", encoding="utf-8")
            recovered, warnings = kindle_report.load_settings(
                settings_path, create_default=True
            )
            backups = list(root.glob("settings.json.corrupt-*"))

        self.assertEqual(recovered, kindle_report.DEFAULT_SETTINGS)
        self.assertTrue(warnings)
        self.assertEqual(len(backups), 1)

    def test_history_and_word_count_cache_corruption_is_recovered(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            history_path = root / kindle_report.HISTORY_FILENAME
            cache_path = root / kindle_report.DEFAULT_CACHE_FILENAME
            profile_path = (
                root / kindle_report.DEFAULT_ACHIEVEMENT_PROFILE_FILENAME
            )
            history_path.write_text("{broken", encoding="utf-8")
            cache_path.write_text("[broken", encoding="utf-8")
            profile_path.write_text("{broken", encoding="utf-8")

            history = kindle_report.load_history(history_path)
            self.assertEqual(
                history,
                {
                    "version": kindle_report.HISTORY_VERSION,
                    "snapshots": [],
                },
            )
            snapshot = kindle_report.build_snapshot(
                [],
                "2026-09-30T12:00:00+08:00",
                "Kindle",
            )
            kindle_report.save_history(history_path, history, snapshot)
            self.assertEqual(
                len(kindle_report.load_history(history_path)["snapshots"]),
                1,
            )

            self.assertEqual(
                kindle_reader._load_word_count_cache(str(cache_path)),
                {},
            )
            self.assertIsNone(
                kindle_report.load_achievement_profile(profile_path)
            )
            kindle_reader._save_word_count_cache(
                str(cache_path),
                {"book": {"count": 100}},
            )
            self.assertEqual(
                kindle_reader._load_word_count_cache(str(cache_path)),
                {"book": {"count": 100}},
            )
            history_backups = list(
                root.glob("reading_history.json.corrupt-*")
            )
            cache_backups = list(
                root.glob("word_counts_cache.json.corrupt-*")
            )
            profile_backups = list(
                root.glob("achievement_profile.json.corrupt-*")
            )
            temporary_files = list(root.glob("*.tmp"))

        self.assertEqual(len(history_backups), 1)
        self.assertEqual(len(cache_backups), 1)
        self.assertEqual(len(profile_backups), 1)
        self.assertEqual(temporary_files, [])

    def test_history_deduplicates_and_enforces_retention_limit(self) -> None:
        def snapshot(
            total_time_ms: int,
            total_words: int,
            captured_at: str,
            source: str,
        ) -> dict:
            return kindle_report.build_snapshot(
                [
                    {
                        "title": "History Book",
                        "author": "Author",
                        "asin": "B000000001",
                        "total_time_ms": str(total_time_ms),
                        "total_words": str(total_words),
                        "wpm": "100",
                        "progress_percent": "10",
                    }
                ],
                captured_at,
                source,
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            history_path = Path(temp_dir) / kindle_report.HISTORY_FILENAME
            first = snapshot(
                60000,
                100,
                "2026-09-28T10:00:00+08:00",
                "Kindle-A",
            )
            duplicate = snapshot(
                60000,
                100,
                "2026-09-29T10:00:00+08:00",
                "Kindle-B",
            )
            second = snapshot(
                120000,
                220,
                "2026-09-30T10:00:00+08:00",
                "Kindle-B",
            )
            third = snapshot(
                180000,
                340,
                "2026-10-01T10:00:00+08:00",
                "Kindle-B",
            )

            self.assertTrue(
                kindle_report.append_snapshot(
                    history_path,
                    first,
                    history_limit=2,
                )
            )
            self.assertFalse(
                kindle_report.append_snapshot(
                    history_path,
                    duplicate,
                    history_limit=2,
                )
            )
            self.assertTrue(
                kindle_report.append_snapshot(
                    history_path,
                    second,
                    history_limit=2,
                )
            )
            self.assertTrue(
                kindle_report.append_snapshot(
                    history_path,
                    third,
                    history_limit=2,
                )
            )
            saved = kindle_report.load_history(history_path)

        self.assertEqual(len(saved["snapshots"]), 2)
        self.assertEqual(saved["snapshots"][0], second)
        self.assertEqual(saved["snapshots"][1], third)

    def test_history_daily_windows_streak_and_book_changes(self) -> None:
        def snapshot(
            total_time_ms: int,
            total_words: int,
            captured_at: str,
        ) -> dict:
            return kindle_report.build_snapshot(
                [
                    {
                        "title": "Window Book",
                        "author": "Author",
                        "asin": "B000000002",
                        "total_time_ms": str(total_time_ms),
                        "total_words": str(total_words),
                        "wpm": "100",
                        "progress_percent": "10",
                    }
                ],
                captured_at,
                "Kindle",
            )

        history = {
            "version": kindle_report.HISTORY_VERSION,
            "snapshots": [
                snapshot(
                    100000,
                    100,
                    "2026-09-25T10:00:00+08:00",
                ),
                snapshot(
                    100000,
                    100,
                    "2026-09-26T10:00:00+08:00",
                ),
                snapshot(
                    160000,
                    250,
                    "2026-09-28T10:00:00+08:00",
                ),
                snapshot(
                    190000,
                    340,
                    "2026-09-29T10:00:00+08:00",
                ),
            ],
        }
        current = snapshot(
            220000,
            430,
            "2026-09-30T10:00:00+08:00",
        )
        trends = kindle_report.build_trends(current, history)
        daily = {
            row["date"]: row
            for row in trends["daily"]
            if isinstance(row, dict)
        }

        self.assertEqual(len(trends["daily"]), 30)
        self.assertEqual(daily["2026-09-25"]["total_time_ms"], 0)
        self.assertEqual(daily["2026-09-26"]["total_time_ms"], 0)
        self.assertEqual(daily["2026-09-28"]["total_time_ms"], 60000)
        self.assertEqual(daily["2026-09-29"]["total_time_ms"], 30000)
        self.assertEqual(daily["2026-09-30"]["total_time_ms"], 30000)
        self.assertEqual(
            trends["windows"]["7d"]["total_time_ms"],
            120000,
        )
        self.assertEqual(
            trends["windows"]["30d"]["total_time_ms"],
            120000,
        )
        self.assertEqual(trends["windows"]["7d"]["active_days"], 3)
        self.assertEqual(trends["streak"]["current_days"], 3)
        self.assertEqual(trends["streak"]["longest_days"], 3)

        book_change = trends["books"][0]
        self.assertEqual(
            book_change["windows"]["7d"]["total_time_ms"]["delta"],
            120000.0,
        )
        self.assertEqual(
            book_change["windows"]["30d"]["total_words"]["delta"],
            330.0,
        )

    def test_period_reports_week_month_year_snapshot_deltas(self) -> None:
        def snapshot(
            total_time_ms: int,
            total_words: int,
            captured_at: str,
        ) -> dict:
            return kindle_report.build_snapshot(
                [
                    {
                        "title": "Period Book",
                        "author": "Author",
                        "asin": "B000000021",
                        "total_time_ms": str(total_time_ms),
                        "total_words": str(total_words),
                        "wpm": "100",
                        "progress_percent": "10",
                    }
                ],
                captured_at,
                "Kindle",
            )

        history = {
            "version": kindle_report.HISTORY_VERSION,
            "snapshots": [
                snapshot(0, 0, "2026-09-28T00:00:00+08:00"),
                snapshot(3600000, 10000, "2026-09-30T00:00:00+08:00"),
            ],
        }
        current = snapshot(
            5400000,
            15000,
            "2026-10-01T12:00:00+08:00",
        )
        data = kindle_report.build_period_reports(
            current,
            history,
            None,
            date(2026, 10, 1),
        )
        reports = {
            report["key"]: report
            for report in data["reports"]
        }

        week = reports["week"]
        self.assertEqual(week["range_start"], "2026-09-28")
        self.assertEqual(week["range_end"], "2026-10-01")
        self.assertEqual(week["total_time_ms"], 5400000)
        self.assertEqual(week["reading_days"], 4)
        self.assertEqual(week["active_books"], 1)
        self.assertAlmostEqual(week["fastest_wpm"], 166.7, places=1)
        self.assertEqual(week["top_books"][0]["title"], "Period Book")
        self.assertEqual(week["top_books"][0]["time_ms"], 5400000)

        month = reports["month"]
        self.assertEqual(month["range_start"], "2026-10-01")
        self.assertEqual(month["total_time_ms"], 600000)
        self.assertEqual(month["active_books"], 1)

        section = kindle_report.build_period_report_section(data)
        self.assertIn('data-period-trend-chart="bars"', section)
        self.assertIn('data-chart-type="bar"', section)
        self.assertIn("period-bar-chart", section)
        self.assertIn('class="period-bar"', section)
        self.assertIn('class="period-bar-fill"', section)
        self.assertNotIn("<polyline", section)
        self.assertNotIn('class="trend-line"', section)

    def test_period_reports_large_gap_reduces_confidence(self) -> None:
        def snapshot(
            total_time_ms: int,
            total_words: int,
            captured_at: str,
        ) -> dict:
            return kindle_report.build_snapshot(
                [
                    {
                        "title": "Gap Book",
                        "author": "Author",
                        "asin": "B000000022",
                        "total_time_ms": str(total_time_ms),
                        "total_words": str(total_words),
                        "wpm": "100",
                        "progress_percent": "10",
                    }
                ],
                captured_at,
                "Kindle",
            )

        history = {
            "version": kindle_report.HISTORY_VERSION,
            "snapshots": [
                snapshot(0, 0, "2026-09-01T10:00:00+08:00"),
                snapshot(
                    3600000,
                    6000,
                    "2026-09-20T10:00:00+08:00",
                ),
            ],
        }
        current = snapshot(
            7200000,
            12000,
            "2026-09-30T10:00:00+08:00",
        )
        data = kindle_report.build_period_reports(
            current,
            history,
            None,
            date(2026, 10, 1),
        )
        week = data["reports"][0]

        self.assertEqual(week["key"], "week")
        self.assertFalse(week["reliable"])
        self.assertEqual(week["confidence"], "low")
        self.assertGreaterEqual(week["large_gap_count"], 1)
        self.assertTrue(
            any(
                row.get("kind") == "snapshot-gap"
                for row in week["missing_segments"]
            )
        )
        self.assertIn("继续阅读", week["conclusion_zh"])

    def test_period_reports_fmcache_priority_and_empty_fallback(self) -> None:
        def snapshot(
            total_time_ms: int,
            total_words: int,
            captured_at: str,
        ) -> dict:
            return kindle_report.build_snapshot(
                [
                    {
                        "title": "Session Book",
                        "author": "Author",
                        "asin": "B000000023",
                        "total_time_ms": str(total_time_ms),
                        "total_words": str(total_words),
                        "wpm": "100",
                        "progress_percent": "10",
                    }
                ],
                captured_at,
                "Kindle",
            )

        history = {
            "version": kindle_report.HISTORY_VERSION,
            "snapshots": [
                snapshot(0, 0, "2026-09-29T10:00:00+08:00"),
                snapshot(
                    600000,
                    3000,
                    "2026-09-30T10:00:00+08:00",
                ),
            ],
        }
        current = snapshot(
            2400000,
            12000,
            "2026-10-01T10:00:00+08:00",
        )
        fmcache = {
            "status": "ok",
            "session_count": 3,
            "daily": {
                "2026-09-30": 600000,
                "2026-10-01": 1800000,
            },
        }
        data = kindle_report.build_period_reports(
            current,
            history,
            fmcache,
            date(2026, 10, 1),
        )
        week = data["reports"][0]

        self.assertEqual(week["source"], "fmcache")
        self.assertEqual(week["fmcache_session_count"], 3)
        self.assertEqual(week["total_time_ms"], 2400000)
        self.assertEqual(week["reading_days"], 2)

        empty = kindle_report.build_period_reports(
            current,
            {"version": kindle_report.HISTORY_VERSION, "snapshots": []},
            None,
            date(2026, 10, 1),
        )["reports"][0]
        self.assertFalse(empty["reliable"])
        self.assertEqual(empty["total_time_ms"], 0)
        self.assertEqual(empty["observation_count"], 0)
        self.assertTrue(empty["missing_segments"])
        empty_section = kindle_report.build_period_report_section(
            {"reports": [empty]}
        )
        self.assertIn("继续阅读即可解锁", empty_section)
        self.assertNotIn('data-period-trend-chart="bars"', empty_section)

    def test_annual_report_reuses_year_period_data_and_empty_state(
        self,
    ) -> None:
        empty = kindle_report.build_annual_report_section({}, {})
        self.assertIn('data-annual-report-state="empty"', empty)
        self.assertIn("暂无可生成的年度报告", empty)

        year_report = {
            "key": "year",
            "range_start": "2026-01-01",
            "range_zh": "2026-01-01 至 2026-10-01",
            "range_en": "2026-01-01 to 2026-10-01",
            "total_time_ms": 7200000,
            "reading_days": 8,
            "active_books": 4,
            "fastest_wpm": 222.2,
            "confidence": "low",
            "reliable": False,
            "coverage_zh": "有效覆盖 40%。",
            "coverage_en": "Usable coverage 40%.",
            "source_note_zh": "历史快照回退。",
            "source_note_en": "History-snapshot fallback.",
            "missing_zh": "存在缺失时间段。",
            "missing_en": "Some intervals are missing.",
            "conclusion_zh": "历史不足，不能形成可靠结论。",
            "conclusion_en": "Insufficient history for a reliable conclusion.",
            "snapshot_count": 14,
            "observation_count": 13,
            "fmcache_session_count": 0,
            "top_books": [
                {
                    "title": f"Book {index} <unsafe>",
                    "time_ms": 3600000 // index,
                    "words": 1000 * index,
                    "wpm": 100 + index,
                }
                for index in range(1, 6)
            ],
        }
        document = kindle_report.build_annual_report_section(
            {"reports": [year_report]},
            {"snapshot_count": 14},
        )
        self.assertIn('data-annual-report-state="insufficient"', document)
        self.assertIn("2026 阅读年鉴", document)
        self.assertIn("2026 Reading year", document)
        self.assertIn("2 小时 0 分", document)
        self.assertIn("222.2 WPM", document)
        self.assertIn("Book 1 &lt;unsafe&gt;", document)
        self.assertEqual(document.count('class="annual-book"'), 3)
        self.assertIn("历史快照回退", document)
        self.assertNotIn("年度总词数", document)

    def test_trend_chart_requires_two_snapshots_for_a_curve(self) -> None:
        base = {
            "snapshot_count": 1,
            "timeline": [
                {
                    "captured_at": "2026-09-30T22:20:03+08:00",
                    "source": "Kindle",
                    "totals": {
                        "total_time_ms": 60000,
                        "total_words": 100,
                        "weighted_wpm": 100.0,
                        "average_progress": 1.0,
                    },
                }
            ],
            "daily": [
                {
                    "date": "2026-09-30",
                    "total_time_ms": 60000,
                    "total_words": 100,
                    "active": True,
                }
            ],
            "windows": {},
            "streak": {},
        }
        single = kindle_report.build_trend_sections(base)
        self.assertEqual(single["point_counts"]["time"], 1)
        self.assertIn('data-chart-state="single"', single["chart"])
        self.assertNotIn('class="trend-line"', single["chart"])

        base["snapshot_count"] = 2
        base["timeline"].append(
            {
                "captured_at": "2026-10-01T08:00:00+08:00",
                "source": "Kindle",
                "totals": {
                    "total_time_ms": 120000,
                    "total_words": 240,
                    "weighted_wpm": 120.0,
                    "average_progress": 2.0,
                },
            }
        )
        two = kindle_report.build_trend_sections(base)
        self.assertEqual(two["point_counts"]["time"], 2)
        self.assertIn('data-chart-state="ready"', two["chart"])
        self.assertIn('class="trend-line"', two["chart"])

    def test_achievements_thresholds_speed_gate_and_empty_state(self) -> None:
        rows = [
            {
                "title": "Active Book",
                "asin": "B000000010",
                "total_time_ms": "7200000",
                "total_words": "12000",
                "progress_percent": "25",
                "wpm": "100",
            },
            {
                "title": "Unread Book",
                "asin": "B000000011",
                "total_time_ms": "0",
                "total_words": "0",
                "progress_percent": "0",
                "wpm": "0",
            },
        ]
        trends = {
            "snapshot_count": 15,
            "timeline": [{} for _ in range(15)],
            "daily": [
                {
                    "date": "2026-09-29",
                    "total_time_ms": 1800000,
                    "total_words": 3000,
                    "active": True,
                },
                {
                    "date": "2026-09-30",
                    "total_time_ms": 1800000,
                    "total_words": 3000,
                    "active": True,
                },
            ],
            "streak": {
                "current_days": 3,
                "longest_days": 7,
            },
        }
        summary = kindle_report.build_summary(rows)
        achievements = kindle_report.build_achievements(
            summary,
            trends,
            rows,
        )
        by_key = {
            item["key"]: item for item in achievements["items"]
        }

        self.assertEqual(achievements["total"], 4)
        self.assertEqual(achievements["unlocked"], 0)
        self.assertEqual(achievements["stars_total"], 16)
        self.assertEqual(achievements["stars_unlocked"], 0)
        self.assertNotIn("history", by_key)
        self.assertEqual(by_key["streak"]["value"], 7.0)
        self.assertEqual(by_key["streak"]["state"], "in_progress")
        self.assertEqual(by_key["time"]["state"], "in_progress")
        self.assertEqual(by_key["books"]["value"], 0.0)
        self.assertEqual(by_key["books"]["state"], "in_progress")
        self.assertAlmostEqual(
            by_key["speed"]["value"],
            100.0,
            places=1,
        )
        self.assertEqual(by_key["speed"]["state"], "in_progress")
        self.assertEqual(by_key["speed"]["target"], 110.0)
        self.assertIn("words", by_key)
        self.assertEqual(len(by_key["words"]["stars"]), 4)

        too_short = dict(trends)
        too_short["daily"] = [
            {
                "date": "2026-09-30",
                "total_time_ms": 1799999,
                "total_words": 10000,
                "active": True,
            }
        ]
        too_short_speed = kindle_report.build_achievements(
            summary,
            too_short,
            rows,
        )["items"][4]
        self.assertEqual(too_short_speed["state"], "insufficient")
        self.assertEqual(
            too_short_speed["detail_zh"],
            "继续阅读即可解锁",
        )

        too_few_words = dict(trends)
        too_few_words["daily"] = [
            {
                "date": "2026-09-30",
                "total_time_ms": 3600000,
                "total_words": 2999,
                "active": True,
            }
        ]
        too_few_speed = kindle_report.build_achievements(
            summary,
            too_few_words,
            rows,
        )["items"][4]
        self.assertEqual(too_few_speed["state"], "insufficient")

        empty = kindle_report.build_achievements(
            kindle_report.build_summary([]),
            {
                "snapshot_count": 1,
                "timeline": [],
                "daily": [],
                "streak": {},
            },
            [],
        )
        self.assertTrue(
            all(
                item["state"] == "insufficient"
                and item["detail_zh"] == "继续阅读即可解锁"
                for item in empty["items"]
            )
        )

    def test_achievement_profile_is_strict_monotonic_and_manga_safe(self) -> None:
        rows = [
            {
                "title": "Completed Book",
                "asin": "B000000041",
                "total_time_ms": "7200000",
                "total_words": "20000",
                "progress_percent": "96",
                "wpm": "100",
            },
            {
                "title": "Manga Volume",
                "asin": "B000000042",
                "total_time_ms": "3600000",
                "total_words": "50000",
                "progress_percent": "100",
                "wpm": "900",
                "reading_mode": "manga",
            },
        ]
        base_profile = kindle_report.build_achievement_profile(
            rows,
            {"version": kindle_report.HISTORY_VERSION, "snapshots": []},
            captured_at="2026-10-02T12:00:00+08:00",
            threshold_percent=95.0,
        )
        self.assertTrue(base_profile["levels_initialized"])
        self.assertTrue(
            all(level == 0 for level in base_profile["levels"].values())
        )
        for metric, tiers in base_profile["tiers"].items():
            thresholds = [tier["threshold"] for tier in tiers]
            self.assertEqual(thresholds, sorted(thresholds))
            self.assertTrue(all(thresholds[index] > thresholds[index - 1]
                                for index in range(1, len(thresholds))))
            with self.subTest(metric=metric):
                self.assertGreater(thresholds[0], 0.0)

        updated = kindle_report.build_achievement_profile(
            rows,
            {"version": kindle_report.HISTORY_VERSION, "snapshots": []},
            captured_at="2026-10-03T12:00:00+08:00",
            existing_profile=base_profile,
            threshold_percent=95.0,
        )
        self.assertTrue(updated["levels_initialized"])
        self.assertEqual(updated["levels"]["books"], 0)
        self.assertEqual(updated["levels"]["words"], 0)
        more_rows = rows + [
            {
                "title": "Third Completed Book",
                "asin": "B000000043",
                "total_time_ms": "3600000",
                "total_words": "12000",
                "progress_percent": "99",
                "wpm": "100",
            }
        ]
        advanced = kindle_report.build_achievement_profile(
            more_rows,
            {"version": kindle_report.HISTORY_VERSION, "snapshots": []},
            captured_at="2026-10-03T18:00:00+08:00",
            existing_profile=updated,
            threshold_percent=95.0,
        )
        self.assertEqual(advanced["levels"]["books"], 1)
        smaller = kindle_report.build_achievement_profile(
            [],
            {"version": kindle_report.HISTORY_VERSION, "snapshots": []},
            captured_at="2026-10-04T12:00:00+08:00",
            existing_profile=advanced,
            threshold_percent=95.0,
        )
        self.assertEqual(smaller["levels"]["books"], 1)
        snapshot = kindle_report.build_snapshot(
            rows,
            "2026-10-02T12:00:00+08:00",
            "Kindle",
        )
        totals = kindle_report.snapshot_totals(snapshot)
        self.assertEqual(totals["total_words"], 20000)
        self.assertEqual(totals["word_books"], 1)
        self.assertEqual(totals["manga_books"], 1)
        self.assertNotIn(
            "reading_mode",
            snapshot["books"][kindle_report.stable_book_id(rows[0])],
        )
        self.assertEqual(
            snapshot["books"][kindle_report.stable_book_id(rows[1])][
                "reading_mode"
            ],
            "manga",
        )

        summary = kindle_report.build_summary(rows)
        trends = {
            "daily": [
                {
                    "date": "2026-10-02",
                    "total_time_ms": 3600000,
                    "total_words": 50000,
                    "active": True,
                }
            ],
            "streak": {"current_days": 8, "longest_days": 8},
        }
        achievements = kindle_report.build_achievements(
            summary,
            trends,
            rows,
            profile=advanced,
            threshold_percent=95.0,
        )
        by_key = {
            item["key"]: item for item in achievements["items"]
        }
        self.assertEqual(by_key["books"]["value"], 1.0)
        self.assertEqual(by_key["manga_volumes"]["value"], 1.0)
        self.assertEqual(by_key["books"]["state"], "in_progress")
        self.assertEqual(by_key["books"]["level"], 1)
        self.assertEqual(by_key["words"]["value"], 20000.0)
        self.assertEqual(by_key["words"]["state"], "in_progress")
        self.assertTrue(
            all(
                star["state"] in {"lit", "next", "future", "next-pending"}
                for star in by_key["words"]["stars"]
            )
        )
        self.assertTrue(
            any("aria" in key for key in by_key["words"]["stars"][0])
        )
        self.assertTrue(kindle_report.is_completed_book(rows[0], 95.0))
        self.assertFalse(kindle_report.is_completed_book(rows[1], 95.0))
        self.assertTrue(
            kindle_report.is_completed_manga_volume(rows[1], 95.0)
        )

    def test_fmcache_reading_sessions_and_heatmap(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "Kindle"
            documents = root / "documents" / "Downloads" / "Items01"
            documents.mkdir(parents=True)
            database = root / kindle_report.FMCACHE_RELATIVE_PATH
            database.parent.mkdir(parents=True)

            connection = sqlite3.connect(database)
            connection.execute(
                "CREATE TABLE reading_sessions("
                "id INTEGER PRIMARY KEY,"
                "encoded_size INT NOT NULL,"
                "session TEXT NOT NULL,"
                "dummy_session INT NULL)"
            )
            sessions = (
                {
                    "payload": {
                        "start_timestamp": "2026-09-30T01:00:00",
                        "end_timestamp": "2026-09-30T01:30:00",
                    }
                },
                {
                    "payload": {
                        "start_timestamp": "2026-09-30T02:00:00",
                        "end_timestamp": "2026-09-30T02:15:00",
                    }
                },
                {
                    "payload": {
                        "start_timestamp": "2026-09-29T23:50:00",
                        "end_timestamp": "2026-09-30T00:10:00",
                    }
                },
            )
            connection.executemany(
                "INSERT INTO reading_sessions("
                "id, encoded_size, session, dummy_session) "
                "VALUES (?, ?, ?, NULL)",
                (
                    (index, len(json.dumps(session)), json.dumps(session))
                    for index, session in enumerate(sessions, start=1)
                ),
            )
            connection.commit()
            connection.close()

            resolved = kindle_report.resolve_fmcache_db(str(documents))
            fmcache = kindle_report.read_fmcache_sessions(resolved)
            history = {"version": kindle_report.HISTORY_VERSION, "snapshots": []}
            current = kindle_report.build_snapshot(
                [],
                "2026-09-30T10:00:00",
                str(documents),
            )
            heatmap = kindle_report.build_heatmap_data(
                history,
                current,
                fmcache,
                days=14,
            )
            document = kindle_report.build_html(
                [],
                Path("reading.csv"),
                str(documents),
                generated_at="2026-09-30T10:00:00",
                heatmap=heatmap,
            )

        self.assertEqual(resolved, database.resolve())
        self.assertEqual(fmcache["status"], "ok")
        self.assertEqual(fmcache["session_count"], 3)
        self.assertEqual(fmcache["daily"]["2026-09-29"], 600000)
        self.assertEqual(fmcache["daily"]["2026-09-30"], 3300000)
        self.assertEqual(heatmap["source"], "fmcache")
        self.assertIn('id="reading-heatmap"', document)
        self.assertIn('data-source="fmcache"', document)
        self.assertIn('data-minutes="55.000"', document)
        self.assertIn("fmcache.db 阅读会话", document)

    def test_fmcache_per_book_manga_fallback(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            database = Path(temp_dir) / "fmcache.db"
            connection = sqlite3.connect(database)
            connection.execute(
                "CREATE TABLE reading_sessions("
                "id INTEGER PRIMARY KEY,"
                "encoded_size INT NOT NULL,"
                "session TEXT NOT NULL,"
                "dummy_session INT NULL)"
            )
            connection.execute(
                "CREATE TABLE records("
                "id INTEGER PRIMARY KEY,"
                "schema_name TEXT NOT NULL,"
                "record TEXT NOT NULL)"
            )
            asin = "A0000000000000000000000000000005"
            session = {
                "payload": {
                    "asin": asin,
                    "start_timestamp": "2026-10-02T14:23:54",
                    "end_timestamp": "2026-10-02T14:24:54",
                    "end_reading_location": 186,
                }
            }
            connection.execute(
                "INSERT INTO reading_sessions("
                "id, encoded_size, session, dummy_session) "
                "VALUES (1, ?, ?, NULL)",
                (len(json.dumps(session)), json.dumps(session)),
            )
            connection.execute(
                "INSERT INTO records(id, schema_name, record) "
                "VALUES (1, 'shared_reading_update_v2', ?)",
                (
                    json.dumps(
                        {
                            "library_asin": asin,
                            "reading_progress": 100,
                            "reading_state": "READ",
                        }
                    ),
                ),
            )
            connection.commit()
            connection.close()

            fmcache = kindle_report.read_fmcache_sessions(database)
            book = fmcache["books"][asin]
            self.assertEqual(book["duration_ms"], 60000)
            self.assertEqual(book["progress_percent"], 100.0)
            self.assertEqual(book["location"], 186)

            rows, applied = kindle_report.merge_fmcache_book_fallbacks(
                [
                    {
                        "title": "合成漫画测试卷 Vol.04",
                        "asin": asin,
                        "reading_mode": "manga",
                        "total_time_ms": "0",
                        "progress_percent": "0",
                        "last_read": "",
                    }
                ],
                fmcache,
            )

        self.assertEqual(applied, [asin])
        self.assertEqual(rows[0]["total_time_ms"], "60000")
        self.assertEqual(rows[0]["minutes"], "1")
        self.assertEqual(rows[0]["progress_percent"], "100")
        self.assertEqual(rows[0]["progress"], "100")
        self.assertEqual(rows[0]["last_read"], "2026-10-02T14:24:54")
        self.assertEqual(rows[0]["fmcache_fallback"], "True")

    def test_fmcache_schema_mismatch_falls_back_to_history(self) -> None:
        def snapshot(total_time_ms: int, captured_at: str) -> dict:
            return kindle_report.build_snapshot(
                [
                    {
                        "title": "Fallback Book",
                        "author": "Author",
                        "asin": "B000000003",
                        "total_time_ms": str(total_time_ms),
                        "total_words": str(total_time_ms),
                        "wpm": "100",
                        "progress_percent": "10",
                    }
                ],
                captured_at,
                "Kindle",
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            database = Path(temp_dir) / "fmcache.db"
            connection = sqlite3.connect(database)
            connection.execute(
                "CREATE TABLE reading_sessions("
                "id INTEGER PRIMARY KEY, payload TEXT)"
            )
            connection.commit()
            connection.close()

            fmcache = kindle_report.read_fmcache_sessions(database)
            history = {
                "version": kindle_report.HISTORY_VERSION,
                "snapshots": [
                    snapshot(100000, "2026-09-28T10:00:00"),
                    snapshot(130000, "2026-09-29T10:00:00"),
                ],
            }
            current = snapshot(190000, "2026-09-30T10:00:00")
            heatmap = kindle_report.build_heatmap_data(
                history,
                current,
                fmcache,
                days=14,
            )
            daily = {
                row["date"]: row for row in heatmap["daily"]
            }
            document = kindle_report.build_html(
                [],
                Path("reading.csv"),
                "Kindle",
                generated_at="2026-09-30T10:00:00",
                heatmap=heatmap,
            )

        self.assertEqual(fmcache["status"], "schema_mismatch")
        self.assertEqual(heatmap["source"], "history")
        self.assertEqual(daily["2026-09-29"]["total_time_ms"], 30000)
        self.assertEqual(daily["2026-09-30"]["total_time_ms"], 60000)
        self.assertIn('data-source="history"', document)
        self.assertIn("历史快照回退", document)

    def test_fmcache_empty_or_corrupt_database_is_safe(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            empty_database = Path(temp_dir) / "empty.db"
            connection = sqlite3.connect(empty_database)
            connection.execute(
                "CREATE TABLE reading_sessions("
                "id INTEGER PRIMARY KEY,"
                "encoded_size INT NOT NULL,"
                "session TEXT NOT NULL,"
                "dummy_session INT NULL)"
            )
            connection.execute(
                "CREATE TABLE records("
                "id INTEGER PRIMARY KEY,"
                "schema_name TEXT NOT NULL,"
                "schema_version TEXT NOT NULL,"
                "app_session_id INT NOT NULL,"
                "reading_session_id INT NULL,"
                "encoded_size INT NOT NULL,"
                "sequence_number INT NOT NULL,"
                "created_timestamp INT NOT NULL,"
                "priority INT NOT NULL,"
                "record TEXT NOT NULL)"
            )
            connection.execute(
                "INSERT INTO records("
                "id, schema_name, schema_version, app_session_id,"
                "reading_session_id, encoded_size, sequence_number,"
                "created_timestamp, priority, record) "
                "VALUES (1, 'ereader_powerd_state_change', '0', 27, "
                "NULL, 213, 1, 1, 0, '{\"curr_state\":\"ACTIVE\"}')"
            )
            connection.commit()
            connection.close()

            corrupt_database = Path(temp_dir) / "corrupt.db"
            corrupt_database.write_bytes(b"not a sqlite database")

            empty = kindle_report.read_fmcache_sessions(empty_database)
            corrupt = kindle_report.read_fmcache_sessions(corrupt_database)
            missing = kindle_report.read_fmcache_sessions(
                Path(temp_dir) / "missing.db"
            )

        self.assertEqual(empty["status"], "empty")
        self.assertEqual(empty["session_count"], 0)
        self.assertEqual(empty["invalid_count"], 0)
        self.assertEqual(corrupt["status"], "unreadable")
        self.assertEqual(missing["status"], "missing")

    def test_current_quality_metrics_and_confidence(self) -> None:
        with CURRENT_CSV_PATH.open(
            "r", encoding="utf-8-sig", newline=""
        ) as handle:
            rows = list(csv.DictReader(handle))

        summary = kindle_report.build_quality_summary(rows)
        by_asin = {
            row["asin"]: kindle_report.analyze_book_quality(row)
            for row in rows
        }

        self.assertEqual(summary["exact_count"], 4)
        self.assertEqual(summary["exact_coverage_percent"], 100.0)
        self.assertEqual(summary["deviation_count"], 4)
        self.assertEqual(summary["median_deviation_abs"], 17.1)
        self.assertEqual(summary["max_deviation_abs"], 119.3)
        self.assertEqual(summary["abnormal_wpm_count"], 0)
        self.assertEqual(summary["abnormal_progress_count"], 0)
        self.assertEqual(summary["low_confidence_count"], 1)

        self.assertEqual(
            by_asin["A0000000000000000000000000000001"][
                "deviation_percent"
            ],
            5.5,
        )
        self.assertEqual(
            by_asin["A0000000000000000000000000000002"][
                "deviation_percent"
            ],
            119.3,
        )
        self.assertEqual(
            by_asin["A0000000000000000000000000000003"][
                "deviation_percent"
            ],
            -5.4,
        )
        self.assertEqual(
            by_asin["A0000000000000000000000000000004"][
                "deviation_percent"
            ],
            28.6,
        )
        self.assertEqual(
            by_asin["A0000000000000000000000000000001"][
                "confidence_level"
            ],
            "high",
        )
        self.assertEqual(
            by_asin["A0000000000000000000000000000002"][
                "confidence_level"
            ],
            "low",
        )

    def test_quality_detects_abnormal_values(self) -> None:
        row = {
            "title": "Abnormal",
            "author": "Author",
            "asin": "B000000099",
            "total_time_ms": "600000",
            "total_words": "2000",
            "book_words": "1000",
            "book_words_exact": "500",
            "word_count_language": "en",
            "wpm": "900",
            "progress_percent": "80",
        }

        quality = kindle_report.analyze_book_quality(row)

        self.assertEqual(quality["deviation_percent"], 100.0)
        self.assertTrue(quality["abnormal_wpm"])
        self.assertTrue(quality["abnormal_progress"])
        self.assertEqual(quality["confidence_level"], "low")
        self.assertLessEqual(quality["confidence_score"], 55)
        self.assertIn("WPM 异常", quality["flags"])
        self.assertIn("进度异常", quality["flags"])
        self.assertIn("外推偏差大", quality["flags"])

    def test_dashboard_distinguishes_word_count_origins(self) -> None:
        rows = [
            {
                "title": "Exact Book",
                "author": "Author",
                "asin": "B000000101",
                "total_time_ms": "600000",
                "total_words": "500",
                "book_words": "900",
                "book_words_exact": "1000",
                "word_count_source": "en-regex",
                "word_count_language": "en",
                "wpm": "50",
                "progress_percent": "50",
            },
            {
                "title": "Estimate Book",
                "author": "Author",
                "asin": "B000000102",
                "total_time_ms": "600000",
                "total_words": "500",
                "book_words": "800",
                "book_words_exact": "",
                "word_count_source": "kindle-estimate",
                "word_count_language": "en",
                "wpm": "50",
                "progress_percent": "50",
            },
        ]

        document = kindle_report.build_html(
            rows,
            Path("reading.csv"),
            "Kindle",
            generated_at="2026-09-30T12:00:00+08:00",
        )

        self.assertIn('data-count="exact"', document)
        self.assertIn('data-count="estimate"', document)
        self.assertIn("Kindle 原始", document)
        self.assertIn("Kindle 外推", document)
        self.assertIn("正文精确", document)
        self.assertIn('id="quality-summary"', document)
        self.assertIn("高置信", document)
        self.assertGreaterEqual(document.count("data-quality-score="), 2)

    def test_manga_mode_is_mapping_only_and_excluded_from_words(self) -> None:
        mapping = {
            "by_asin": {"B000000201": "manga"},
            "by_title": {"mapped manga": "manga"},
        }
        self.assertEqual(
            kindle_report.reading_mode_for_book(
                {"asin": "B000000201", "title": "Anything"}, mapping
            ),
            "manga",
        )
        self.assertEqual(
            kindle_report.reading_mode_for_book(
                {"asin": "", "title": "Mapped Manga"}, mapping
            ),
            "manga",
        )
        self.assertEqual(
            kindle_report.reading_mode_for_book(
                {"asin": "", "title": "Comic Tag Only"}, mapping
            ),
            "book",
        )
        rows = [
            {
                "title": "Word Book",
                "asin": "B000000202",
                "total_time_ms": "600000",
                "total_words": "1000",
                "book_words_exact": "10000",
                "word_count_source": "en-regex",
                "word_count_language": "en",
                "wpm": "100",
                "progress_percent": "50",
            },
            {
                "title": "Mapped Manga",
                "asin": "B000000201",
                "reading_mode": "manga",
                "total_time_ms": "3600000",
                "total_words": "50000",
                "book_words_exact": "",
                "word_count_source": "not-applicable",
                "wpm": "900",
                "progress_percent": "100",
            },
        ]
        summary = kindle_report.build_summary(rows)
        self.assertEqual(summary["total_words"], 1000)
        self.assertEqual(summary["word_books"], 1)
        self.assertEqual(summary["manga_books"], 1)
        self.assertEqual(summary["completed_books"], 1)
        self.assertIsNone(kindle_report.word_progress(rows[1]))
        self.assertFalse(kindle_report.has_exact_word_count(rows[1]))
        self.assertIn("漫画不适用", kindle_report.book_word_origin_label(rows[1]))

        document = kindle_report.build_html(
            rows,
            Path("reading.csv"),
            "Kindle",
            generated_at="2026-10-02T12:00:00+08:00",
            book_modes=mapping,
        )
        self.assertIn('data-count="manga"', document)
        self.assertIn('data-mode="manga"', document)
        self.assertIn("漫画", document)
        self.assertIn("不适用", document)

    def test_free_space_threshold_is_reported(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            available, free_bytes = kindle_report.check_free_space(
                Path(temp_dir),
                required_bytes=0,
            )
            unavailable, _ = kindle_report.check_free_space(
                Path(temp_dir),
                required_bytes=sys.maxsize,
            )

        self.assertTrue(available)
        self.assertIsNotNone(free_bytes)
        self.assertFalse(unavailable)

    def test_dist_ocr_cache_prefers_project_root_cache(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            project = Path(temp_dir) / "project"
            dist = project / "dist"
            dist.mkdir(parents=True)
            project_cache = project / kindle_report.DEFAULT_OCR_CACHE_DIRNAME
            (project_cache / "render").mkdir(parents=True)

            resolved = kindle_report.resolve_ocr_cache_dir("", dist)

        self.assertEqual(resolved, project_cache.resolve())

    def test_full_workflow_uses_settings_path_and_output_dir(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            kindle_root = root / "Kindle"
            sidecar_dir = kindle_root / "documents" / "Settings Book.sdr"
            sidecar_dir.mkdir(parents=True)
            (sidecar_dir / "Settings Book.azw3f").write_bytes(
                synthetic_sidecar()
            )
            output_dir = root / "reports"
            settings_path = root / kindle_report.SETTINGS_FILENAME
            settings_path.write_text(
                json.dumps(
                    {
                        "version": kindle_report.SETTINGS_VERSION,
                        "kindle_path": str(kindle_root),
                        "output_dir": str(output_dir),
                        "language": "auto",
                        "exact_word_count": False,
                        "open_html": False,
                    }
                ),
                encoding="utf-8",
            )

            exit_code = kindle_report.main(
                [
                    "--settings",
                    str(settings_path),
                    "--no-open",
                ]
            )
            csv_path = output_dir / kindle_report.DEFAULT_CSV_FILENAME
            html_path = output_dir / kindle_report.DEFAULT_HTML_FILENAME
            history_path = output_dir / kindle_report.HISTORY_FILENAME
            csv_exists = csv_path.is_file()
            html_exists = html_path.is_file()
            history_exists = history_path.is_file()
            with csv_path.open(
                "r", encoding="utf-8-sig", newline=""
            ) as handle:
                rows = list(csv.DictReader(handle))

        self.assertEqual(exit_code, 0)
        self.assertTrue(csv_exists)
        self.assertTrue(html_exists)
        self.assertTrue(history_exists)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["title"], "Settings Book")

    def test_dashboard_keeps_language_and_word_count_labels(self) -> None:
        rows = [
            {
                "title": "Exact English",
                "author": "Author",
                "asin": "B000000001",
                "total_time_ms": "60000",
                "total_words": "100",
                "book_words": "1000",
                "book_words_exact": "1200",
                "word_count_source": "en-regex",
                "word_count_language": "en",
                "wpm": "100",
                "progress_percent": "10",
            },
            {
                "title": "估算日文书",
                "author": "作者",
                "asin": "B000000002",
                "total_time_ms": "30000",
                "total_words": "50",
                "book_words": "5000",
                "book_words_exact": "",
                "word_count_source": "kindle-estimate",
                "word_count_language": "ja",
                "wpm": "100",
                "progress_percent": "1",
            },
        ]
        document = kindle_report.build_html(
            rows,
            Path("reading.csv"),
            "Kindle",
            generated_at="2026-09-30T12:00:00+08:00",
        )

        self.assertIn('data-count="exact"', document)
        self.assertIn('data-count="estimate"', document)
        self.assertIn('data-language="en"', document)
        self.assertIn('data-language="ja"', document)
        self.assertIn("英文", document)
        self.assertIn("日文", document)
        self.assertIn("英文分词", document)
        self.assertIn("kindle-estimate", document)
        self.assertIn('id="language-filter"', document)
        self.assertIn('id="count-filter"', document)
        self.assertIn('id="state-filter"', document)
        self.assertIn("header.dataset.direction", document)
        self.assertIn("window.print()", document)
        self.assertIn(
            '<span class="i18n-zh">累计有效阅读词数</span>'
            '<span class="i18n-en">Effective words read</span>'
            "</span><strong>150</strong>",
            document,
        )
        self.assertNotIn("__WORDS_SOURCE__", document)
        header_actions = document.index('class="header-actions"')
        self.assertLess(header_actions, document.index('id="theme-preset"'))
        self.assertLess(header_actions, document.index('id="language-toggle"'))
        self.assertLess(header_actions, document.index('id="theme-toggle"'))
        self.assertLess(header_actions, document.index('id="appearance-panel"'))
        self.assertLess(
            document.index('id="appearance-panel"'),
            document.index('id="greeting-mode"'),
        )
        self.assertLess(
            document.index('id="appearance-panel"'),
            document.index('class="summary"'),
        )
        self.assertLess(
            document.index('class="summary"'),
            document.index('class="library-panel"'),
        )
        self.assertLess(
            document.index('class="library-panel"'),
            document.index('id="achievement-panel"'),
        )

    def test_dashboard_header_greeting_and_offline_markup(self) -> None:
        with CURRENT_CSV_PATH.open(
            "r", encoding="utf-8-sig", newline=""
        ) as handle:
            rows = list(csv.DictReader(handle))
        document = kindle_report.build_html(
            rows,
            Path("reading.csv"),
            "Kindle",
            generated_at="2026-10-01T12:00:00+08:00",
            xlsx_path=Path("reading.xlsx"),
        )

        for marker in (
            "http://",
            "https://",
            "<link",
            "@import",
            'src="//',
            "src='//",
        ):
            self.assertNotIn(marker, document)
        self.assertIn(
            '"SimSun", "Songti SC", "Noto Serif CJK SC", serif',
            document,
        )
        self.assertRegex(
            document,
            r"\.greeting-message\s*\{[^}]*font-size:\s*16px",
            msg="Greeting font size must be larger than the previous 15px.",
        )
        self.assertIn("transition: color .7s ease", document)
        for greeting_text in (
            "早上好！",
            "Good morning!",
            "(＾▽＾)",
            "晚上好！",
            "Good evening!",
            "(｡･ω･｡)",
        ):
            self.assertIn(greeting_text, document)
        for tone in ("late-night", "morning", "day", "dusk", "night"):
            self.assertIn(f'data-greeting-tone="{tone}"', document)
        self.assertIn("browserGreetingTone", document)
        self.assertIn("GREETING_MODE_TONES", document)
        self.assertIn("GREETING_FONTS", document)
        self.assertIn("kindle-dashboard-greeting-font", document)
        self.assertIn("kindle-dashboard-user-name", document)
        self.assertIn("kindle-dashboard-achievement-state::", document)
        self.assertNotIn("window.prompt", document)
        self.assertIn('id="user-name-dialog"', document)
        self.assertIn("achievement-stars", document)
        self.assertEqual(document.count('id="greeting-mode"'), 1)
        self.assertEqual(document.count('id="user-name-input"'), 1)

        self.assertNotIn('class="settings-bar"', document)
        self.assertNotIn('class="settings-quick"', document)
        self.assertIn(
            '<details class="appearance-panel" id="appearance-panel">',
            document,
        )
        appearance_start = document.index('id="appearance-panel"')
        appearance_end = document.index("</details>", appearance_start)
        appearance_html = document[appearance_start:appearance_end]
        self.assertIn('id="greeting-mode"', appearance_html)
        self.assertIn('id="user-name-input"', appearance_html)
        self.assertIn('id="user-name-clear"', appearance_html)
        self.assertIn("kindle-dashboard-user-name", document)
        self.assertIn("kindle-dashboard-achievement-state::", document)
        self.assertIn('id="greeting-font"', appearance_html)
        self.assertRegex(
            appearance_html,
            r'<input id="theme-follow-greeting" type="checkbox">',
        )
        self.assertIn(
            "kindle-dashboard-theme-follow-greeting",
            document,
        )
        self.assertNotIn("<details", appearance_html[appearance_html.index(">") + 1:])
        self.assertIn("背景图案", appearance_html)
        self.assertIn("Background pattern", appearance_html)
        for texture in ("dots", "rings", "diagonal", "diamond"):
            self.assertIn(
                f'<option value="{texture}"',
                appearance_html,
            )
            self.assertIn(
                f'html[data-background-texture="{texture}"]',
                document,
            )
        self.assertIn('id="background-motif"', appearance_html)
        self.assertIn("装饰图案", appearance_html)
        self.assertIn("Decorative motifs", appearance_html)
        for motif in ("books", "leaves", "stars"):
            self.assertIn(
                f'<option value="{motif}"',
                appearance_html,
            )
            self.assertIn(
                f'html[data-background-motif="{motif}"]',
                document,
            )
        self.assertIn(
            'html:not([data-background-motif="none"]) .header-main::after',
            document,
        )
        self.assertIn(
            'html:not([data-background-motif="none"]) .panel > h2::after',
            document,
        )
        self.assertIn(
            'html:not([data-background-motif="none"]) '
            '.advanced-panel > summary::after',
            document,
        )

        achievement_match = re.search(
            r'<details class="achievement-panel" id="achievement-panel".*?'
            r"</details>",
            document,
            flags=re.DOTALL,
        )
        self.assertIsNotNone(achievement_match)
        achievement_html = achievement_match.group(0)
        self.assertEqual(
            achievement_html.count("data-achievement-key="),
            5,
        )
        self.assertNotIn("报告快照", achievement_html)
        self.assertNotIn("report snapshots", achievement_html)
        self.assertNotIn("历史覆盖", achievement_html)

        header_actions = document.index('class="header-actions"')
        for control in (
            'id="theme-preset"',
            'id="language-toggle"',
            'id="theme-toggle"',
            'id="appearance-panel"',
            'class="appearance-icon"',
        ):
            self.assertGreater(document.index(control), header_actions)
        self.assertLess(
            header_actions,
            document.index('class="summary"'),
        )
        self.assertLess(
            document.index('class="header-main"'),
            document.index('class="header-actions"'),
        )

    def test_exe_build_skips_unreadable_vendor_tokenizers(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            vendor = Path(temp_dir) / "vendor"
            (vendor / "jieba").mkdir(parents=True)
            (vendor / "janome").mkdir()

            self.assertTrue(kindle_report._vendor_tokenizers_readable(vendor))
            self.assertEqual(
                kindle_report._vendor_path_arguments(vendor),
                ["--paths", str(vendor)],
            )

            with mock.patch.object(
                kindle_report.os,
                "listdir",
                side_effect=PermissionError(13, "denied"),
            ):
                self.assertFalse(
                    kindle_report._vendor_tokenizers_readable(vendor)
                )
                self.assertEqual(
                    kindle_report._vendor_path_arguments(vendor),
                    [],
                )

    @unittest.skipUnless(
        CHROMIUM_BROWSER, "Chromium browser is not available"
    )
    def test_dashboard_header_greeting_and_appearance_viewports(
        self,
    ) -> None:
        assert CHROMIUM_BROWSER is not None
        with CURRENT_CSV_PATH.open(
            "r", encoding="utf-8-sig", newline=""
        ) as handle:
            rows = list(csv.DictReader(handle))
        document = kindle_report.build_html(
            rows,
            Path("reading.csv"),
            "Kindle",
            generated_at="2026-10-01T12:00:00+08:00",
            xlsx_path=Path("reading.xlsx"),
        )
        instrumentation = """
<script>
  window.addEventListener("load", () => {
    const root = document.documentElement;
    const actions = document.querySelector(".header-actions");
    const ids = ["theme-preset", "language-toggle", "theme-toggle",
                 "appearance-panel"];
    root.dataset.testControlsInActions = String(ids.every(id => {
      const el = document.getElementById(id);
      return !!el && actions.contains(el);
    }));
    const greetingModeControl = document.getElementById("greeting-mode");
    root.dataset.testGreetingModeInPanel = String(
      !!greetingModeControl &&
      appearancePanel.contains(greetingModeControl) &&
      !!greetingModeControl.closest(".appearance-body")
    );

    const h1 = document.querySelector("h1");
    const greeting = document.getElementById("greeting-message");
    const greetingZh = document.getElementById("greeting-zh");
    const greetingEn = document.getElementById("greeting-en");
    const greetingStyle = getComputedStyle(greeting);
    const zhText = (greetingZh.textContent || "").trim();
    const enText = (greetingEn.textContent || "").trim();
    const withoutEmoticon = value =>
      value.replace(/\\s*[（(][^()（）]*[）)]\\s*$/, "");
    root.dataset.testGreetingZh = zhText;
    root.dataset.testGreetingEn = enText;
    root.dataset.testGreetingZhExclamation =
      /[！!]$/.test(withoutEmoticon(zhText)) ? "true" : "false";
    root.dataset.testGreetingEnExclamation =
      /[！!]$/.test(withoutEmoticon(enText)) ? "true" : "false";
    root.dataset.testGreetingZhEmoticon =
      /[（(][^()（）]*[）)]$/.test(zhText) ? "true" : "false";
    root.dataset.testGreetingEnEmoticon =
      /[（(][^()（）]*[）)]$/.test(enText) ? "true" : "false";
    root.dataset.testGreetingFontSize = greetingStyle.fontSize;
    root.dataset.testGreetingFontDefault = greetingStyle.fontFamily || "";

    const fontSelect = document.getElementById("greeting-font");
    root.dataset.testGreetingFontControlInPanel = String(
      !!fontSelect &&
      appearancePanel.contains(fontSelect) &&
      !!fontSelect.closest(".appearance-body")
    );
    if (fontSelect) {
      fontSelect.value = "kaiti";
      fontSelect.dispatchEvent(new Event("change"));
      root.dataset.testGreetingFontSelected =
        getComputedStyle(greeting).fontFamily || "";
      root.dataset.testGreetingFontStored =
        readPreference(GREETING_FONT_KEY) || "";
    }

    const h1Rect = h1.getBoundingClientRect();
    const greetingRect = greeting.getBoundingClientRect();
    const actionsRect = actions.getBoundingClientRect();
    const intersects = (left, right) =>
      Math.min(left.right, right.right) - Math.max(left.left, right.left) > 1 &&
      Math.min(left.bottom, right.bottom) - Math.max(left.top, right.top) > 1;
    root.dataset.testGreetingRightOfTitle =
      greetingRect.left >= h1Rect.right - 2 ? "true" : "false";
    root.dataset.testGreetingSameLine =
      (greetingRect.top < h1Rect.bottom - 1 &&
       greetingRect.bottom > h1Rect.top + 1) ? "true" : "false";
    root.dataset.testGreetingOverlapsTitle =
      String(intersects(greetingRect, h1Rect));
    root.dataset.testGreetingOverlapsActions =
      String(intersects(greetingRect, actionsRect));
    root.dataset.testGreetingScrollFits = String(
      greeting.scrollWidth <= greeting.clientWidth + 1
    );
    root.dataset.testH1Width = String(Math.round(h1Rect.width));
    root.dataset.testGreetingWidth = String(Math.round(greetingRect.width));
    root.dataset.testHeaderMainWidth = String(
      Math.round(document.querySelector(".header-main").getBoundingClientRect().width)
    );
    root.dataset.testHeaderActionsWidth = String(
      Math.round(actionsRect.width)
    );
    root.dataset.testGreetingTone = root.dataset.greetingTone || "";
    root.dataset.testGreetingTransition =
      getComputedStyle(greeting).transitionProperty || "";

    const shelf = document.getElementById("cover-grid");
    root.dataset.testClientWidth =
      String(document.documentElement.clientWidth);
    root.dataset.testAppearanceDefaultClosed = String(!appearancePanel.open);
    appearancePanel.open = false;
    updateLayoutMetrics();
    const shelfTopClosed = shelf.offsetTop;
    root.dataset.collapsedOverflow = root.dataset.overflowX;
    root.dataset.collapsedOverlap = root.dataset.overlapCount;
    root.dataset.collapsedOverlapElements = root.dataset.overlapElements;
    root.dataset.collapsedTruncated = root.dataset.truncatedNumbers;

    appearancePanel.open = true;
    updateLayoutMetrics();
    const body = appearancePanel.querySelector(".appearance-body");
    const bodyStyle = getComputedStyle(body);
    root.dataset.testAppearanceDisplay = bodyStyle.display;
    root.dataset.testAppearancePosition = bodyStyle.position;
    root.dataset.testAppearanceWidth = String(
      Math.round(body.getBoundingClientRect().width)
    );
    root.dataset.testShelfStable = String(
      Math.abs(shelf.offsetTop - shelfTopClosed) < 1
    );
    root.dataset.expandedOverflow = root.dataset.overflowX;
    root.dataset.expandedOverlap = root.dataset.overlapCount;
    root.dataset.expandedOverlapElements = root.dataset.overlapElements;
    root.dataset.expandedTruncated = root.dataset.truncatedNumbers;
  });
</script>
"""
        self.assertIn("</body>", document)
        instrumented = document.replace(
            "</body>", instrumentation + "</body>"
        )
        viewports = {}
        with tempfile.TemporaryDirectory() as temp_dir:
            test_html = Path(temp_dir) / "dashboard-header.html"
            test_html.write_text(instrumented, encoding="utf-8")
            for width, height in (
                (320, 568),
                (360, 640),
                (390, 844),
                (768, 900),
                (1440, 900),
            ):
                viewports[(width, height)] = browser_dom_snapshot(
                    CHROMIUM_BROWSER, test_html, width, height
                )

        for viewport, snapshot in viewports.items():
            attributes = html_root_attributes(snapshot)
            width = viewport[0]
            with self.subTest(viewport=viewport):
                self.assertEqual(
                    attributes.get("data-test-controls-in-actions"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-greeting-mode-in-panel"),
                    "true",
                )
                self.assertEqual(
                    attributes.get(
                        "data-test-greeting-font-control-in-panel"
                    ),
                    "true",
                )
                self.assertEqual(
                    attributes.get(
                        "data-test-greeting-zh-exclamation"
                    ),
                    "true",
                )
                self.assertEqual(
                    attributes.get(
                        "data-test-greeting-en-exclamation"
                    ),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-greeting-zh-emoticon"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-greeting-en-emoticon"),
                    "true",
                )
                self.assertIn(
                    "！", attributes.get("data-test-greeting-zh") or ""
                )
                self.assertIn(
                    "!", attributes.get("data-test-greeting-en") or ""
                )
                font_size = float(
                    (
                        attributes.get("data-test-greeting-font-size") or "0"
                    ).replace("px", "")
                )
                self.assertGreater(font_size, 15.0)
                self.assertIn(
                    "SimSun",
                    attributes.get("data-test-greeting-font-default") or "",
                )
                self.assertIn(
                    "serif",
                    attributes.get("data-test-greeting-font-default") or "",
                )
                self.assertIn(
                    "Kai",
                    attributes.get("data-test-greeting-font-selected") or "",
                )
                self.assertEqual(
                    attributes.get("data-test-greeting-font-stored"),
                    "kaiti",
                )
                self.assertEqual(
                    attributes.get("data-test-greeting-scroll-fits"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-greeting-overlaps-title"),
                    "false",
                )
                self.assertEqual(
                    attributes.get("data-test-greeting-overlaps-actions"),
                    "false",
                )
                self.assertIn(
                    attributes.get("data-test-greeting-tone") or "",
                    {"late-night", "morning", "day", "dusk", "night"},
                )
                self.assertIn(
                    "color",
                    attributes.get("data-test-greeting-transition") or "",
                )
                self.assertEqual(
                    attributes.get("data-test-appearance-default-closed"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-appearance-position")
                    in {"absolute", "fixed"},
                    True,
                )
                self.assertNotEqual(
                    attributes.get("data-test-appearance-display"), "none"
                )
                self.assertGreater(
                    int(attributes.get("data-test-appearance-width") or "0"),
                    0,
                )
                self.assertLessEqual(
                    int(attributes.get("data-test-appearance-width") or "0"),
                    int(attributes.get("data-test-client-width") or "0"),
                )
                self.assertEqual(
                    attributes.get("data-test-shelf-stable"), "true"
                )
                for prefix in ("collapsed", "expanded"):
                    self.assertEqual(
                        attributes.get("data-" + prefix + "-overflow"),
                        "false",
                    )
                    self.assertEqual(
                        attributes.get("data-" + prefix + "-overlap"),
                        "0",
                        attributes.get(
                            "data-" + prefix + "-overlap-elements"
                        ),
                    )
                    self.assertEqual(
                        attributes.get("data-" + prefix + "-truncated"),
                        "",
                    )
            if width >= 1440:
                self.assertEqual(
                    attributes.get("data-test-greeting-right-of-title"),
                    "true",
                    attributes,
                )
                self.assertEqual(
                    attributes.get("data-test-greeting-same-line"), "true"
                )

    def test_dashboard_has_theme_language_columns_and_print_contract(
        self,
    ) -> None:
        rows = [
            {
                "title": "Exact English",
                "author": "Author",
                "asin": "B000000001",
                "total_time_ms": "60000",
                "total_words": "100",
                "book_words": "1000",
                "book_words_exact": "1200",
                "word_count_source": "en-regex",
                "word_count_language": "en",
                "wpm": "100",
                "progress_percent": "10",
            },
            {
                "title": "估算日文书",
                "author": "作者",
                "asin": "B000000002",
                "total_time_ms": "30000",
                "total_words": "50",
                "book_words": "5000",
                "book_words_exact": "",
                "word_count_source": "kindle-estimate",
                "word_count_language": "ja",
                "wpm": "100",
                "progress_percent": "1",
            },
        ]
        document = kindle_report.build_html(
            rows,
            Path("reading.csv"),
            "Kindle",
            generated_at="2026-10-01T12:00:00+08:00",
            xlsx_path=Path("reading.xlsx"),
        )

        self.assertIn('id="theme-toggle"', document)
        self.assertIn('id="language-toggle"', document)
        self.assertIn('id="xlsx-export"', document)
        self.assertIn('href="reading.xlsx"', document)
        self.assertIn('data-column-toggle="book"', document)
        self.assertIn('data-column="book"', document)
        self.assertIn("function setTheme", document)
        self.assertIn("function setUiLanguage", document)
        self.assertIn("function applyColumnVisibility", document)
        self.assertIn('data-ui-language="zh"', document)
        self.assertIn("@page", document)
        self.assertIn("@media print", document)
        self.assertIn(".appearance-panel,\n      .local-name-dialog", document)
        self.assertIn("A4 landscape", document)
        self.assertIn('data-count="exact"', document)
        self.assertIn('data-count="estimate"', document)
        self.assertIn('class="i18n-en"', document)
        for preset in (
            "forest",
            "ocean",
            "sunset",
            "graphite",
            "rose",
            "lavender",
            "sky",
            "sepia",
        ):
            self.assertIn(f'value="{preset}"', document)
            self.assertIn(f'data-theme-preset="{preset}"', document)

    def test_xlsx_export_is_real_workbook_and_preserves_exact_counts(
        self,
    ) -> None:
        with CURRENT_CSV_PATH.open(
            "r", encoding="utf-8-sig", newline=""
        ) as handle:
            reader = csv.DictReader(handle)
            fieldnames = list(reader.fieldnames or [])
            rows = [dict(row) for row in reader]

        with tempfile.TemporaryDirectory() as temp_dir:
            xlsx_path = Path(temp_dir) / "reading.xlsx"
            kindle_report.write_xlsx_report(
                rows,
                xlsx_path,
                fieldnames=fieldnames,
                source_path="Kindle",
                generated_at="2026-10-01T12:00:00+08:00",
            )

            self.assertTrue(zipfile.is_zipfile(xlsx_path))
            with zipfile.ZipFile(xlsx_path) as archive:
                names = set(archive.namelist())
            self.assertIn("[Content_Types].xml", names)
            self.assertIn("xl/workbook.xml", names)
            self.assertIn("xl/worksheets/sheet1.xml", names)
            self.assertIn("xl/worksheets/sheet2.xml", names)

            from openpyxl import load_workbook

            workbook = load_workbook(xlsx_path, data_only=False)
            self.assertEqual(workbook.sheetnames, ["Summary", "Books"])
            sheet = workbook["Books"]
            headers = [
                str(cell.value or "")
                for cell in sheet[1]
            ]
            exact_column = headers.index(
                "正文精确总词数 / Exact body words"
            ) + 1
            asin_column = headers.index("ASIN") + 1
            effective_words_column = headers.index(
                "有效总词数 / Effective book words"
            ) + 1
            effective_read_column = headers.index(
                "有效已读词数 / Effective words read"
            ) + 1
            effective_wpm_column = headers.index(
                "有效 WPM / Effective WPM"
            ) + 1
            effective_source_column = headers.index(
                "有效词数来源 / Effective-word source"
            ) + 1
            exact_by_asin = {}
            effective_by_asin = {}
            for row_index in range(2, sheet.max_row + 1):
                exact_value = sheet.cell(
                    row=row_index,
                    column=exact_column,
                ).value
                asin_value = str(
                    sheet.cell(row=row_index, column=asin_column).value
                )
                if exact_value is not None:
                    exact_by_asin[asin_value] = int(exact_value)
                effective_by_asin[
                    asin_value
                ] = {
                    "book_words": sheet.cell(
                        row=row_index,
                        column=effective_words_column,
                    ).value,
                    "words_read": sheet.cell(
                        row=row_index,
                        column=effective_read_column,
                    ).value,
                    "wpm": sheet.cell(
                        row=row_index,
                        column=effective_wpm_column,
                    ).value,
                    "source": sheet.cell(
                        row=row_index,
                        column=effective_source_column,
                    ).value,
                }

        self.assertEqual(
            exact_by_asin,
            {
                "A0000000000000000000000000000001": 1087776,
                "A0000000000000000000000000000002": 121161,
                "A0000000000000000000000000000003": 360318,
                "A0000000000000000000000000000004": 63025,
            },
        )
        self.assertEqual(
            effective_by_asin["A0000000000000000000000000000005"],
            {
                "book_words": 5977,
                "words_read": 5977,
                "wpm": 100.07,
                "source": "ocr-estimate",
            },
        )

    @unittest.skipUnless(
        CHROMIUM_BROWSER, "Chromium browser is not available"
    )
    def test_dashboard_filters_sorting_and_viewport_overflow(self) -> None:
        assert CHROMIUM_BROWSER is not None
        document = CURRENT_DASHBOARD_PATH.read_text(encoding="utf-8")
        instrumentation = """
<script>
  languageFilter.value = "en";
  applyFilters();
  document.documentElement.dataset.testLanguageRows =
    document.documentElement.dataset.visibleRows;

  languageFilter.value = "";
  countFilter.value = "estimate";
  applyFilters();
  document.documentElement.dataset.testEstimateRows =
    document.documentElement.dataset.visibleRows;

  countFilter.value = "";
  stateFilter.value = "reading";
  applyFilters();
  document.documentElement.dataset.testReadingRows =
    document.documentElement.dataset.visibleRows;

  stateFilter.value = "";
  applyFilters();
  document.querySelectorAll("#data-table th")[4].click();
  const firstBook = document.querySelector(
    "#data-table tbody tr strong"
  );
  document.documentElement.dataset.testFirstBook =
    firstBook ? firstBook.textContent.trim() : "";
  appearancePanel.open = false;
  updateLayoutMetrics();
  document.documentElement.dataset.collapsedOverflow =
    document.documentElement.dataset.overflowX;
  document.documentElement.dataset.collapsedOverlap =
    document.documentElement.dataset.overlapCount;
  document.documentElement.dataset.collapsedTruncated =
    document.documentElement.dataset.truncatedNumbers;

  appearancePanel.open = true;
  updateLayoutMetrics();
  document.documentElement.dataset.expandedOverflow =
    document.documentElement.dataset.overflowX;
  document.documentElement.dataset.expandedOverlap =
    document.documentElement.dataset.overlapCount;
  document.documentElement.dataset.expandedTruncated =
    document.documentElement.dataset.truncatedNumbers;
</script>
"""
        self.assertIn("</body>", document)
        instrumented = document.replace(
            "</body>", instrumentation + "</body>"
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            test_html = Path(temp_dir) / "dashboard-test.html"
            test_html.write_text(instrumented, encoding="utf-8")
            snapshots = {
                (320, 568): browser_dom_snapshot(
                    CHROMIUM_BROWSER, test_html, 320, 568
                ),
                (360, 640): browser_dom_snapshot(
                    CHROMIUM_BROWSER, test_html, 360, 640
                ),
                (390, 844): browser_dom_snapshot(
                    CHROMIUM_BROWSER, test_html, 390, 844
                ),
                (768, 900): browser_dom_snapshot(
                    CHROMIUM_BROWSER, test_html, 768, 900
                ),
                (1440, 900): browser_dom_snapshot(
                    CHROMIUM_BROWSER, test_html, 1440, 900
                ),
            }

        for viewport, snapshot in snapshots.items():
            attributes = html_root_attributes(snapshot)
            with self.subTest(viewport=viewport):
                self.assertEqual(
                    attributes.get("data-overflow-x"), "false"
                )
                self.assertEqual(
                    attributes.get("data-overflow-elements"), ""
                )
                self.assertEqual(attributes.get("data-table-rows"), "5")
                self.assertEqual(
                    attributes.get("data-visible-rows"), "5"
                )
                self.assertEqual(
                    attributes.get("data-test-language-rows"), "2"
                )
                self.assertEqual(
                    attributes.get("data-test-estimate-rows"), "0"
                )
                self.assertEqual(
                    attributes.get("data-test-reading-rows"), "5"
                )
                self.assertEqual(
                    attributes.get("data-test-first-book"), "合成中文测试书"
                )
                self.assertEqual(
                    attributes.get("data-collapsed-overflow"), "false"
                )
                self.assertEqual(
                    attributes.get("data-collapsed-overlap"), "0"
                )
                self.assertEqual(
                    attributes.get("data-collapsed-truncated"), ""
                )
                self.assertEqual(
                    attributes.get("data-expanded-overflow"), "false"
                )
                self.assertEqual(
                    attributes.get("data-expanded-overlap"), "0"
                )
                self.assertEqual(
                    attributes.get("data-expanded-truncated"), ""
                )

    @unittest.skipUnless(
        CHROMIUM_BROWSER, "Chromium browser is not available"
    )
    def test_dashboard_greeting_and_achievement_viewports(self) -> None:
        assert CHROMIUM_BROWSER is not None
        with CURRENT_CSV_PATH.open(
            "r", encoding="utf-8-sig", newline=""
        ) as handle:
            reader = csv.DictReader(handle)
            rows = [dict(row) for row in reader]

        def snapshot(
            total_time_ms: int,
            total_words: int,
            captured_at: str,
        ) -> dict:
            return kindle_report.build_snapshot(
                [
                    {
                        "title": "Achievement Fixture",
                        "author": "Author",
                        "asin": "B000000012",
                        "total_time_ms": str(total_time_ms),
                        "total_words": str(total_words),
                        "wpm": "100",
                        "progress_percent": "10",
                    }
                ],
                captured_at,
                "Kindle",
            )

        history = {
            "version": kindle_report.HISTORY_VERSION,
            "snapshots": [
                snapshot(60000, 1, "2026-09-27T10:00:00+08:00"),
                snapshot(600000, 3000, "2026-09-28T10:00:00+08:00"),
                snapshot(1200000, 6000, "2026-09-29T10:00:00+08:00"),
                snapshot(1800000, 9000, "2026-09-30T10:00:00+08:00"),
            ],
        }
        current = kindle_report.build_snapshot(
            rows,
            "2026-10-01T23:00:00+08:00",
            "Kindle",
        )
        trends = kindle_report.build_trends(current, history)
        document = kindle_report.build_html(
            rows,
            Path("reading.csv"),
            "Kindle",
            trends=trends,
            generated_at="2026-10-01T23:00:00+08:00",
            xlsx_path=Path("reading.xlsx"),
        )
        instrumentation = """
<script>
  const NativeDate = Date;
  class FixedDate extends NativeDate {
    constructor(...args) {
      if (args.length) {
        super(...args);
      } else {
        super(2026, 9, 1, 23, 0, 0);
      }
    }
    static now() {
      return new NativeDate(2026, 9, 1, 23, 0, 0).getTime();
    }
  }
  window.Date = FixedDate;
  window.addEventListener("load", () => {
    document.documentElement.dataset.testGreetingAuto =
      document.documentElement.dataset.greetingSource;
    document.documentElement.dataset.testGreetingAutoPeriod =
      document.documentElement.dataset.greetingPeriod;
    document.documentElement.dataset.testGreetingInvalid =
      normalizeGreetingMode("invalid");
    document.documentElement.dataset.testUserNameDialogShown =
      userNameDialog.open ? "true" : "false";
    userNameSkip.click();
    document.documentElement.dataset.testUserNameDialogSkipped =
      userNameDialog.open ? "false" : "true";
    document.documentElement.dataset.testUserNameSkippedStored =
      readPreference(USER_NAME_KEY) === "" ? "true" : "false";
    setUserName("小林");
    document.documentElement.dataset.testUserNameStored =
      readPreference(USER_NAME_KEY);
    document.documentElement.dataset.testUserNameGreeting =
      document.getElementById("greeting-zh").textContent.trim();
    setUserName("");

    greetingMode.value = "morning";
    greetingMode.dispatchEvent(new Event("change"));
    document.documentElement.dataset.testGreetingManualMode =
      document.documentElement.dataset.greetingMode;
    document.documentElement.dataset.testGreetingManualPeriod =
      document.documentElement.dataset.greetingPeriod;
    document.documentElement.dataset.testGreetingStored =
      readPreference(GREETING_MODE_KEY);

    greetingMode.value = "auto";
    greetingMode.dispatchEvent(new Event("change"));
    document.documentElement.dataset.testGreetingRestored =
      document.documentElement.dataset.greetingSource;

    document.documentElement.dataset.testAchievementPanelDefault =
      document.documentElement.dataset.achievementPanel;
    document.documentElement.dataset.testAchievementCount =
      document.documentElement.dataset.achievementCount;
    document.documentElement.dataset.testAchievementUnlocked =
      document.documentElement.dataset.achievementUnlocked;
    document.documentElement.dataset.testAchievementStarCount = String(
      document.querySelectorAll(".achievement-star").length
    );
    document.documentElement.dataset.testAchievementStarAria = String(
      Array.from(document.querySelectorAll(".achievement-star")).every(
        star => !!(star.getAttribute("aria-label") || "").trim()
      )
    );
    const achievementGrid = document.querySelector(".achievement-grid");
    const achievementColumns = getComputedStyle(achievementGrid)
      .gridTemplateColumns.split(" ").filter(Boolean);
    document.documentElement.dataset.testAchievementColumns =
      String(achievementColumns.length);
    const achievementGridRect = achievementGrid.getBoundingClientRect();
    const lastAchievementRect = achievementCards[
      achievementCards.length - 1
    ].getBoundingClientRect();
    document.documentElement.dataset.testAchievementTrailingGap =
      String(Math.max(
        0,
        Math.round(achievementGridRect.right - lastAchievementRect.right)
      ));
    document.documentElement.dataset.testAchievementCardWidth =
      String(Math.max(...achievementCards.map(
        card => Math.round(card.getBoundingClientRect().width)
      )));

    achievementPanel.open = false;
    achievementPanel.dispatchEvent(new Event("toggle"));
    updateLayoutMetrics();
    document.documentElement.dataset.testAchievementCollapsedStored =
      readPreference(ACHIEVEMENT_PANEL_KEY);
    document.documentElement.dataset.collapsedOverflow =
      document.documentElement.dataset.overflowX;
    document.documentElement.dataset.collapsedOverlap =
      document.documentElement.dataset.overlapCount;
    document.documentElement.dataset.collapsedOverlapElements =
      document.documentElement.dataset.overlapElements;
    document.documentElement.dataset.collapsedTruncated =
      document.documentElement.dataset.truncatedNumbers;

    achievementPanel.open = true;
    achievementPanel.dispatchEvent(new Event("toggle"));
    updateLayoutMetrics();
    document.documentElement.dataset.testAchievementExpandedStored =
      readPreference(ACHIEVEMENT_PANEL_KEY);
    document.documentElement.dataset.expandedOverflow =
      document.documentElement.dataset.overflowX;
    document.documentElement.dataset.expandedOverlap =
      document.documentElement.dataset.overlapCount;
    document.documentElement.dataset.expandedOverlapElements =
      document.documentElement.dataset.overlapElements;
    document.documentElement.dataset.expandedTruncated =
      document.documentElement.dataset.truncatedNumbers;

    languageToggle.click();
    updateLayoutMetrics();
    document.documentElement.dataset.englishExpandedOverflow =
      document.documentElement.dataset.overflowX;
    document.documentElement.dataset.englishExpandedOverlap =
      document.documentElement.dataset.overlapCount;
    document.documentElement.dataset.englishExpandedTruncated =
      document.documentElement.dataset.truncatedNumbers;

    achievementPanel.open = false;
    achievementPanel.dispatchEvent(new Event("toggle"));
    updateLayoutMetrics();
    document.documentElement.dataset.englishCollapsedOverflow =
      document.documentElement.dataset.overflowX;
    document.documentElement.dataset.englishCollapsedOverlap =
      document.documentElement.dataset.overlapCount;
    document.documentElement.dataset.englishCollapsedTruncated =
      document.documentElement.dataset.truncatedNumbers;
  });
</script>
"""
        self.assertIn("</body>", document)
        self.assertIn("kindle-dashboard-greeting-mode", document)
        self.assertIn("kindle-dashboard-achievement-panel", document)
        instrumented = document.replace(
            "</body>", instrumentation + "</body>"
        )

        viewports = {
            (320, 568): None,
            (360, 640): None,
            (390, 844): None,
            (768, 900): None,
            (1440, 900): None,
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            test_html = Path(temp_dir) / "dashboard-achievements.html"
            test_html.write_text(instrumented, encoding="utf-8")
            for viewport in viewports:
                width, height = viewport
                viewports[viewport] = browser_dom_snapshot(
                    CHROMIUM_BROWSER, test_html, width, height
                )

        for viewport, snapshot_document in viewports.items():
            attributes = html_root_attributes(str(snapshot_document))
            with self.subTest(viewport=viewport):
                self.assertEqual(
                    attributes.get("data-greeting-mode"), "auto"
                )
                self.assertEqual(
                    attributes.get("data-greeting-period"), "evening"
                )
                self.assertEqual(
                    attributes.get("data-test-greeting-auto"),
                    "browser-local-time",
                )
                self.assertEqual(
                    attributes.get("data-test-greeting-auto-period"),
                    "evening",
                )
                self.assertEqual(
                    attributes.get("data-test-greeting-manual-mode"),
                    "morning",
                )
                self.assertEqual(
                    attributes.get("data-test-greeting-manual-period"),
                    "morning",
                )
                self.assertEqual(
                    attributes.get("data-test-greeting-stored"),
                    "morning",
                )
                self.assertEqual(
                    attributes.get("data-test-greeting-restored"),
                    "browser-local-time",
                )
                self.assertEqual(
                    attributes.get("data-test-greeting-invalid"), "auto"
                )
                self.assertEqual(
                    attributes.get("data-test-user-name-dialog-shown"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-user-name-dialog-skipped"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-user-name-skipped-stored"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-user-name-stored"), "小林"
                )
                self.assertIn(
                    "小林",
                    attributes.get("data-test-user-name-greeting") or "",
                )
                self.assertEqual(
                    attributes.get("data-test-achievement-panel-default"),
                    "open",
                )
                self.assertEqual(
                    attributes.get("data-test-achievement-count"), "5"
                )
                expected_columns = (
                    5 if viewport[0] >= 1440
                    else 3 if viewport[0] >= 768
                    else 1
                )
                self.assertEqual(
                    attributes.get("data-test-achievement-columns"),
                    str(expected_columns),
                )
                self.assertLessEqual(
                    int(
                        attributes.get(
                            "data-test-achievement-trailing-gap"
                        )
                        or "0"
                    ),
                    int(
                        attributes.get(
                            "data-test-achievement-card-width"
                        )
                        or "0"
                    ) + 16,
                )
                self.assertEqual(
                    attributes.get("data-test-achievement-unlocked"), "0"
                )
                self.assertEqual(
                    attributes.get("data-test-achievement-star-count"),
                    "20",
                )
                self.assertEqual(
                    attributes.get("data-test-achievement-star-aria"), "true"
                )
                self.assertEqual(
                    attributes.get(
                        "data-test-achievement-collapsed-stored"
                    ),
                    "closed",
                )
                self.assertEqual(
                    attributes.get(
                        "data-test-achievement-expanded-stored"
                    ),
                    "open",
                )
                for prefix in (
                    "collapsed",
                    "expanded",
                    "english-expanded",
                    "english-collapsed",
                ):
                    self.assertEqual(
                        attributes.get("data-" + prefix + "-overflow"),
                        "false",
                    )
                    self.assertEqual(
                        attributes.get("data-" + prefix + "-overlap"),
                        "0",
                        attributes.get(
                            "data-" + prefix + "-overlap-elements"
                        ),
                    )
                    self.assertEqual(
                        attributes.get("data-" + prefix + "-truncated"),
                        "",
                    )

    @unittest.skipUnless(
        CHROMIUM_BROWSER, "Chromium browser is not available"
    )
    def test_dashboard_period_reports_preferences_and_viewports(self) -> None:
        assert CHROMIUM_BROWSER is not None
        with CURRENT_CSV_PATH.open(
            "r", encoding="utf-8-sig", newline=""
        ) as handle:
            reader = csv.DictReader(handle)
            rows = [dict(row) for row in reader]

        def snapshot(
            total_time_ms: int,
            total_words: int,
            captured_at: str,
        ) -> dict:
            return kindle_report.build_snapshot(
                [
                    {
                        "title": "Period Dashboard Book",
                        "author": "Author",
                        "asin": "B000000024",
                        "total_time_ms": str(total_time_ms),
                        "total_words": str(total_words),
                        "wpm": "100",
                        "progress_percent": "10",
                    }
                ],
                captured_at,
                "Kindle",
            )

        history = {
            "version": kindle_report.HISTORY_VERSION,
            "snapshots": [
                snapshot(0, 0, "2026-09-28T00:00:00+08:00"),
                snapshot(
                    1200000,
                    6000,
                    "2026-09-29T00:00:00+08:00",
                ),
                snapshot(
                    2400000,
                    12000,
                    "2026-09-30T00:00:00+08:00",
                ),
            ],
        }
        period_current = snapshot(
            3600000,
            18000,
            "2026-10-01T12:00:00+08:00",
        )
        period_data = kindle_report.build_period_reports(
            period_current,
            history,
            None,
            date(2026, 10, 1),
        )
        trends = kindle_report.build_trends(period_current, history)
        document = kindle_report.build_html(
            rows,
            Path("reading.csv"),
            "Kindle",
            trends=trends,
            period_reports=period_data,
            generated_at="2026-10-01T12:00:00+08:00",
            xlsx_path=Path("reading.xlsx"),
        )
        instrumentation = """
<script>
  window.addEventListener("load", () => {
    document.documentElement.dataset.testPeriodCount =
      String(periodViews.length);
    document.documentElement.dataset.testPeriodDefaultPanel =
      document.documentElement.dataset.periodPanel;
    document.documentElement.dataset.testPeriodDefaultView =
      document.documentElement.dataset.periodView;
    document.documentElement.dataset.testPeriodDefaultReliable =
      document.documentElement.dataset.periodReliable;
    document.documentElement.dataset.testPeriodTrendType =
      document.documentElement.dataset.periodTrendType || "";
    document.documentElement.dataset.testPeriodTrendBuckets =
      document.documentElement.dataset.periodTrendBuckets || "0";
    document.documentElement.dataset.testPeriodBarChartCount =
      String(periodPanel.querySelectorAll(
        '[data-period-trend-chart="bars"]'
      ).length);
    document.documentElement.dataset.testPeriodLineChartCount =
      String(periodPanel.querySelectorAll("polyline, .trend-line").length);
    const activePeriodView = periodViews.find(view => !view.hidden);
    document.documentElement.dataset.testPeriodActiveBarState =
      activePeriodView?.querySelector(
        '[data-period-trend-chart="bars"]'
      )?.dataset.periodTrendChart || "";

    advancedPanel.open = false;
    periodPanel.open = false;
    advancedPanel.dispatchEvent(new Event("toggle"));
    periodPanel.dispatchEvent(new Event("toggle"));
    updateLayoutMetrics();
    document.documentElement.dataset.testPeriodCollapsedStored =
      readPreference(PERIOD_PANEL_KEY);
    document.documentElement.dataset.collapsedOverflow =
      document.documentElement.dataset.overflowX;
    document.documentElement.dataset.collapsedOverlap =
      document.documentElement.dataset.overlapCount;
    document.documentElement.dataset.collapsedTruncated =
      document.documentElement.dataset.truncatedNumbers;

    advancedPanel.open = true;
    periodPanel.open = true;
    advancedPanel.dispatchEvent(new Event("toggle"));
    periodPanel.dispatchEvent(new Event("toggle"));
    updateLayoutMetrics();
    document.documentElement.dataset.testPeriodExpandedStored =
      readPreference(PERIOD_PANEL_KEY);
    document.documentElement.dataset.periodExpandedOverflow =
      document.documentElement.dataset.overflowX;
    document.documentElement.dataset.periodExpandedOverlap =
      document.documentElement.dataset.overlapCount;
    document.documentElement.dataset.periodExpandedTruncated =
      document.documentElement.dataset.truncatedNumbers;

    const yearTab = periodTabs.find(
      tab => tab.dataset.periodTab === "year"
    );
    yearTab.click();
    updateLayoutMetrics();
    document.documentElement.dataset.testPeriodYearView =
      document.documentElement.dataset.periodView;
    document.documentElement.dataset.testPeriodYearStored =
      readPreference(PERIOD_VIEW_KEY);
    document.documentElement.dataset.testPeriodYearTrendType =
      document.documentElement.dataset.periodTrendType || "";
    document.documentElement.dataset.periodYearOverflow =
      document.documentElement.dataset.overflowX;
    document.documentElement.dataset.periodYearOverlap =
      document.documentElement.dataset.overlapCount;
    document.documentElement.dataset.periodYearTruncated =
      document.documentElement.dataset.truncatedNumbers;

    languageToggle.click();
    updateLayoutMetrics();
    document.documentElement.dataset.periodEnglishOverflow =
      document.documentElement.dataset.overflowX;
    document.documentElement.dataset.periodEnglishOverlap =
      document.documentElement.dataset.overlapCount;
    document.documentElement.dataset.periodEnglishTruncated =
      document.documentElement.dataset.truncatedNumbers;
  });
</script>
"""
        self.assertIn("kindle-dashboard-period-view", document)
        self.assertIn("kindle-dashboard-period-panel", document)
        self.assertIn('data-period-trend-chart="bars"', document)
        self.assertNotIn('class="period-trend-row"', document)
        instrumented = document.replace(
            "</body>", instrumentation + "</body>"
        )

        viewports = {
            (320, 568): None,
            (360, 640): None,
            (390, 844): None,
            (768, 900): None,
            (1440, 900): None,
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            test_html = Path(temp_dir) / "dashboard-periods.html"
            test_html.write_text(instrumented, encoding="utf-8")
            for viewport in viewports:
                width, height = viewport
                viewports[viewport] = browser_dom_snapshot(
                    CHROMIUM_BROWSER, test_html, width, height
                )

        for viewport, snapshot_document in viewports.items():
            attributes = html_root_attributes(str(snapshot_document))
            with self.subTest(viewport=viewport):
                self.assertEqual(
                    attributes.get("data-test-period-count"), "3"
                )
                self.assertEqual(
                    attributes.get("data-test-period-default-panel"),
                    "open",
                )
                self.assertEqual(
                    attributes.get("data-test-period-default-view"),
                    "week",
                )
                self.assertEqual(
                    attributes.get("data-test-period-default-reliable"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-period-trend-type"),
                    "bars",
                )
                self.assertEqual(
                    attributes.get("data-test-period-active-bar-state"),
                    "bars",
                )
                self.assertEqual(
                    attributes.get("data-test-period-line-chart-count"),
                    "0",
                )
                self.assertGreaterEqual(
                    int(
                        attributes.get(
                            "data-test-period-trend-buckets"
                        )
                        or "0"
                    ),
                    4,
                )
                self.assertGreaterEqual(
                    int(
                        attributes.get(
                            "data-test-period-bar-chart-count"
                        )
                        or "0"
                    ),
                    1,
                )
                self.assertEqual(
                    attributes.get("data-test-period-year-trend-type"),
                    "bars",
                )
                self.assertEqual(
                    attributes.get("data-test-period-collapsed-stored"),
                    "closed",
                )
                self.assertEqual(
                    attributes.get("data-test-period-expanded-stored"),
                    "open",
                )
                self.assertEqual(
                    attributes.get("data-test-period-year-view"),
                    "year",
                )
                self.assertEqual(
                    attributes.get("data-test-period-year-stored"),
                    "year",
                )
                for prefix in (
                    "collapsed",
                    "period-expanded",
                    "period-year",
                    "period-english",
                ):
                    self.assertEqual(
                        attributes.get("data-" + prefix + "-overflow"),
                        "false",
                    )
                    self.assertEqual(
                        attributes.get("data-" + prefix + "-overlap"),
                        "0",
                        attributes.get(
                            "data-" + prefix + "-overlap-elements"
                        ),
                    )
                    self.assertEqual(
                        attributes.get("data-" + prefix + "-truncated"),
                        "",
                    )

    @unittest.skipUnless(
        CHROMIUM_BROWSER, "Chromium browser is not available"
    )
    def test_dashboard_appearance_presets_backgrounds_and_viewports(
        self,
    ) -> None:
        assert CHROMIUM_BROWSER is not None
        with CURRENT_CSV_PATH.open(
            "r", encoding="utf-8-sig", newline=""
        ) as handle:
            reader = csv.DictReader(handle)
            rows = [dict(row) for row in reader]
        cover_url = (
            "data:image/png;base64,"
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwC"
            "AAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
        )
        cover_urls = {
            kindle_report.stable_book_id(row): cover_url
            for row in rows
        }
        expected_latest = max(
            rows,
            key=lambda row: str(row.get("last_read") or ""),
        )
        expected_latest_id = kindle_report.stable_book_id(expected_latest)
        expected_fixed_id = kindle_report.stable_book_id(rows[0])
        document = kindle_report.build_html(
            rows,
            Path("reading.csv"),
            "Kindle",
            generated_at="2026-10-01T12:00:00+08:00",
            cover_urls=cover_urls,
            xlsx_path=Path("reading.xlsx"),
        )
        instrumentation = f"""
<script>
  window.addEventListener("load", () => {{
    const themeFollowControl = document.getElementById(
      "theme-follow-greeting"
    );
    document.documentElement.dataset.testThemeFollowInPanel = String(
      !!themeFollowControl &&
      appearancePanel.contains(themeFollowControl) &&
      !!themeFollowControl.closest(".appearance-body")
    );
    document.documentElement.dataset.testThemeFollowDefault =
      themeFollowControl && themeFollowControl.checked ? "on" : "off";
    document.documentElement.dataset.testThemeFollowInitialStored =
      readPreference(THEME_FOLLOW_GREETING_KEY) || "";

    themePresetSelect.value = "lavender";
    themePresetSelect.dispatchEvent(new Event("change"));
    themeToggle.click();
    const appearanceCards = Array.from(
      document.querySelectorAll(".book-cover-card[data-book-id]")
    );
    document.documentElement.dataset.testPresetStored =
      readPreference("kindle-dashboard-theme-preset");
    document.documentElement.dataset.testThemePreset =
      document.documentElement.dataset.themePreset;

    setUploadedBackground("{cover_url}", true);
    document.documentElement.dataset.testUploadActive =
      document.documentElement.dataset.backgroundActive;
    document.documentElement.dataset.testUploadStored = String(
      readPreference(BACKGROUND_IMAGE_KEY) === "{cover_url}"
    );
    backgroundClear.click();
    document.documentElement.dataset.testUploadCleared = String(
      readPreference(BACKGROUND_IMAGE_KEY) === null &&
      document.documentElement.dataset.backgroundActive === "false"
    );

    backgroundSource.value = "recent";
    backgroundSource.dispatchEvent(new Event("change"));
    document.documentElement.dataset.testAutoActive =
      document.documentElement.dataset.backgroundActive;
    document.documentElement.dataset.testAutoReason =
      document.documentElement.dataset.backgroundReason;
    document.documentElement.dataset.testAutoBookId =
      document.documentElement.dataset.backgroundBookId;

    const latestCoverImage = appearanceCards
      .find(card => card.dataset.bookId ===
        document.documentElement.dataset.backgroundBookId)
      ?.querySelector("img[data-cover-image]");
    if (latestCoverImage) {{
      latestCoverImage.dataset.coverFailed = "true";
    }}
    applyAppearanceSettings(true);
    document.documentElement.dataset.testMissingCoverFallback =
      document.documentElement.dataset.backgroundFallback;
    document.documentElement.dataset.testMissingCoverReason =
      document.documentElement.dataset.backgroundReason;
    if (latestCoverImage) {{
      delete latestCoverImage.dataset.coverFailed;
    }}

    backgroundSource.value = "fixed";
    backgroundSource.dispatchEvent(new Event("change"));
    backgroundBook.value = "{expected_fixed_id}";
    backgroundBook.dispatchEvent(new Event("change"));
    document.documentElement.dataset.testFixedActive =
      document.documentElement.dataset.backgroundActive;
    document.documentElement.dataset.testFixedBookId =
      document.documentElement.dataset.backgroundBookId;

    backgroundSettings.bookId = "missing-book";
    saveBackgroundSettings();
    applyAppearanceSettings(true);
    document.documentElement.dataset.testMissingFallback =
      document.documentElement.dataset.backgroundFallback;
    document.documentElement.dataset.testMissingReason =
      document.documentElement.dataset.backgroundReason;

    appearanceReset.click();
    document.documentElement.dataset.testResetPreset =
      document.documentElement.dataset.themePreset;
    document.documentElement.dataset.testResetSource =
      document.documentElement.dataset.backgroundSource;
    document.documentElement.dataset.testResetBadgeStyle =
      document.documentElement.dataset.achievementBadgeStyle;
    document.documentElement.dataset.testResetTexture =
      document.documentElement.dataset.backgroundTexture;
    document.documentElement.dataset.testResetMotif =
      document.documentElement.dataset.backgroundMotif;

    setThemePreset("rose", true, true);
    greetingMode.value = "morning";
    greetingMode.dispatchEvent(new Event("change"));
    themeFollowControl.checked = true;
    themeFollowControl.dispatchEvent(new Event("change"));
    document.documentElement.dataset.testThemeFollowEnabledStored =
      readPreference(THEME_FOLLOW_GREETING_KEY);
    document.documentElement.dataset.testThemeFollowMorningPreset =
      document.documentElement.dataset.themePreset;
    document.documentElement.dataset.testThemeFollowManualStored =
      readPreference("kindle-dashboard-theme-preset");

    greetingMode.value = "evening";
    greetingMode.dispatchEvent(new Event("change"));
    document.documentElement.dataset.testThemeFollowNightPreset =
      document.documentElement.dataset.themePreset;
    document.documentElement.dataset.testThemeFollowManualStoredAfterSwitch =
      readPreference("kindle-dashboard-theme-preset");

    document.documentElement.dataset.themePreset = "forest";
    restoreAppearance();
    document.documentElement.dataset.testThemeFollowReloadPreset =
      document.documentElement.dataset.themePreset;
    themeFollowControl.checked = false;
    themeFollowControl.dispatchEvent(new Event("change"));
    document.documentElement.dataset.testThemeFollowDisabledStored =
      readPreference(THEME_FOLLOW_GREETING_KEY);
    document.documentElement.dataset.testThemeFollowRestoredPreset =
      document.documentElement.dataset.themePreset;

    setAchievementBadgeStyle("invalid", false);
    setBackgroundTexture("invalid", false);
    setBackgroundMotif("invalid", false);
    document.documentElement.dataset.testBadgeStyleFallback =
      document.documentElement.dataset.achievementBadgeStyle;
    document.documentElement.dataset.testTextureFallback =
      document.documentElement.dataset.backgroundTexture;
    document.documentElement.dataset.testMotifFallback =
      document.documentElement.dataset.backgroundMotif;

    achievementBadgeStyleSelect.value = "ring";
    achievementBadgeStyleSelect.dispatchEvent(new Event("change"));
    backgroundTextureSelect.value = "grid";
    backgroundTextureSelect.dispatchEvent(new Event("change"));
    document.documentElement.dataset.testBadgeStyleStored =
      readPreference("kindle-dashboard-achievement-badge-style");
    document.documentElement.dataset.testBadgeStyleApplied =
      document.documentElement.dataset.achievementBadgeStyle;
    document.documentElement.dataset.testTextureStored =
      readPreference("kindle-dashboard-background-texture");
    document.documentElement.dataset.testTextureApplied =
      document.documentElement.dataset.backgroundTexture;

    const textureBackgrounds = {{}};
    [
      "paper", "grid", "grain",
      "dots", "rings", "diagonal", "diamond", "none",
    ].forEach(value => {{
      setBackgroundTexture(value, false);
      textureBackgrounds[value] =
        getComputedStyle(document.body).backgroundImage;
    }});
    document.documentElement.dataset.testTextureBackgrounds =
      JSON.stringify(textureBackgrounds);
    setBackgroundTexture("none", false);
    const motifBackgrounds = {{}};
    ["none", "books", "leaves", "stars"].forEach(value => {{
      setBackgroundMotif(value, false);
      motifBackgrounds[value] =
        getComputedStyle(document.body).backgroundImage;
    }});
    document.documentElement.dataset.testMotifBackgrounds =
      JSON.stringify(motifBackgrounds);
    const motifFlourishTargets = {{
      header: document.querySelector(".header-main"),
      panel: document.querySelector(".panel > h2"),
      collapsible: document.querySelector("#achievement-panel > summary"),
    }};
    const motifFlourishes = {{}};
    ["none", "books", "leaves", "stars"].forEach(value => {{
      setBackgroundMotif(value, false);
      motifFlourishes[value] = {{}};
      Object.entries(motifFlourishTargets).forEach(([name, element]) => {{
        motifFlourishes[value][name] = element
          ? getComputedStyle(element, "::after").backgroundImage
          : "";
      }});
    }});
    document.documentElement.dataset.testMotifFlourishes =
      JSON.stringify(motifFlourishes);
    backgroundMotifSelect.value = "stars";
    backgroundMotifSelect.dispatchEvent(new Event("change"));
    document.documentElement.dataset.testMotifStored =
      readPreference("kindle-dashboard-background-motif");
    document.documentElement.dataset.testMotifApplied =
      document.documentElement.dataset.backgroundMotif;
    setBackgroundMotif("none", false);
    removePreference("kindle-dashboard-background-motif");
    setBackgroundTexture("grid", false);

    annualReportPanel.open = true;
    annualReportPanel.dispatchEvent(new Event("toggle"));
    document.documentElement.dataset.testAnnualState =
      document.documentElement.dataset.annualReportState;
    document.documentElement.dataset.testAnnualReliable =
      document.documentElement.dataset.annualReportReliable;

    backgroundSource.value = "recent";
    backgroundSource.dispatchEvent(new Event("change"));
    advancedPanel.open = false;
    appearancePanel.open = false;
    annualReportPanel.open = false;
    annualReportPanel.dispatchEvent(new Event("toggle"));
    updateLayoutMetrics();
    document.documentElement.dataset.collapsedOverflow =
      document.documentElement.dataset.overflowX;
    document.documentElement.dataset.collapsedOverlap =
      document.documentElement.dataset.overlapCount;
    document.documentElement.dataset.collapsedTruncated =
      document.documentElement.dataset.truncatedNumbers;

    advancedPanel.open = true;
    appearancePanel.open = true;
    annualReportPanel.open = true;
    annualReportPanel.dispatchEvent(new Event("toggle"));
    updateLayoutMetrics();
    document.documentElement.dataset.expandedOverflow =
      document.documentElement.dataset.overflowX;
    document.documentElement.dataset.expandedOverlap =
      document.documentElement.dataset.overlapCount;
    document.documentElement.dataset.expandedTruncated =
      document.documentElement.dataset.truncatedNumbers;
    document.documentElement.dataset.testAnnualPanelStored =
      readPreference("kindle-dashboard-annual-report-panel");
  }});
</script>
"""
        self.assertIn("</body>", document)
        self.assertIn("kindle-dashboard-theme-preset", document)
        self.assertIn(
            "kindle-dashboard-theme-follow-greeting",
            document,
        )
        self.assertIn("kindle-dashboard-background-settings", document)
        self.assertIn("kindle-dashboard-background-image", document)
        self.assertIn(
            "kindle-dashboard-achievement-badge-style", document
        )
        self.assertIn("kindle-dashboard-background-texture", document)
        self.assertIn("kindle-dashboard-annual-report-panel", document)
        instrumented = document.replace(
            "</body>", instrumentation + "</body>"
        )

        viewports = {
            (320, 568): None,
            (360, 640): None,
            (390, 844): None,
            (768, 900): None,
            (1440, 900): None,
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            test_html = Path(temp_dir) / "dashboard-appearance.html"
            test_html.write_text(instrumented, encoding="utf-8")
            for viewport in viewports:
                width, height = viewport
                viewports[viewport] = browser_dom_snapshot(
                    CHROMIUM_BROWSER, test_html, width, height
                )

        for viewport, snapshot in viewports.items():
            attributes = html_root_attributes(str(snapshot))
            with self.subTest(viewport=viewport):
                self.assertEqual(
                    attributes.get("data-test-preset-stored"), "lavender"
                )
                self.assertEqual(
                    attributes.get("data-test-theme-preset"), "lavender"
                )
                self.assertEqual(
                    attributes.get("data-test-theme-follow-in-panel"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-theme-follow-default"),
                    "off",
                )
                self.assertEqual(
                    attributes.get(
                        "data-test-theme-follow-initial-stored"
                    ),
                    "",
                )
                self.assertEqual(
                    attributes.get(
                        "data-test-theme-follow-enabled-stored"
                    ),
                    "true",
                )
                self.assertEqual(
                    attributes.get(
                        "data-test-theme-follow-morning-preset"
                    ),
                    "sepia",
                )
                self.assertEqual(
                    attributes.get(
                        "data-test-theme-follow-manual-stored"
                    ),
                    "rose",
                )
                self.assertEqual(
                    attributes.get(
                        "data-test-theme-follow-night-preset"
                    ),
                    "graphite",
                )
                self.assertEqual(
                    attributes.get(
                        "data-test-theme-follow-manual-stored-after-switch"
                    ),
                    "rose",
                )
                self.assertEqual(
                    attributes.get(
                        "data-test-theme-follow-reload-preset"
                    ),
                    "graphite",
                )
                self.assertEqual(
                    attributes.get(
                        "data-test-theme-follow-disabled-stored"
                    ),
                    "false",
                )
                self.assertEqual(
                    attributes.get(
                        "data-test-theme-follow-restored-preset"
                    ),
                    "rose",
                )
                self.assertEqual(
                    attributes.get("data-theme-follow-greeting"), "false"
                )
                self.assertEqual(
                    attributes.get("data-test-upload-active"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-upload-stored"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-upload-cleared"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-auto-active"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-auto-reason"), "last-read"
                )
                self.assertEqual(
                    attributes.get("data-test-auto-book-id"),
                    expected_latest_id,
                )
                self.assertEqual(
                    attributes.get("data-test-missing-cover-fallback"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-missing-cover-reason"),
                    "last-read-cover-missing",
                )
                self.assertEqual(
                    attributes.get("data-test-fixed-active"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-fixed-book-id"),
                    expected_fixed_id,
                )
                self.assertEqual(
                    attributes.get("data-test-missing-fallback"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-missing-reason"),
                    "missing-selection",
                )
                self.assertEqual(
                    attributes.get("data-test-reset-preset"), "forest"
                )
                self.assertEqual(
                    attributes.get("data-test-reset-source"), "none"
                )
                self.assertEqual(
                    attributes.get("data-test-reset-badge-style"), "seal"
                )
                self.assertEqual(
                    attributes.get("data-test-reset-texture"), "paper"
                )
                self.assertEqual(
                    attributes.get("data-test-badge-style-fallback"), "seal"
                )
                self.assertEqual(
                    attributes.get("data-test-texture-fallback"), "paper"
                )
                self.assertEqual(
                    attributes.get("data-test-badge-style-stored"), "ring"
                )
                self.assertEqual(
                    attributes.get("data-test-badge-style-applied"), "ring"
                )
                self.assertEqual(
                    attributes.get("data-test-texture-stored"), "grid"
                )
                self.assertEqual(
                    attributes.get("data-test-texture-applied"), "grid"
                )
                texture_backgrounds = json.loads(
                    attributes.get("data-test-texture-backgrounds") or "{}"
                )
                self.assertEqual(
                    set(texture_backgrounds),
                    {
                        "paper", "grid", "grain",
                        "dots", "rings", "diagonal", "diamond", "none",
                    },
                )
                self.assertNotIn(
                    "gradient", texture_backgrounds.get("none", "")
                )
                def max_alpha(css: str) -> float:
                    return max(
                        (
                            float(alpha)
                            for alpha in re.findall(
                                r"rgba\([^)]*?,\s*([0-9.]+)\)", css
                            )
                        ),
                        default=0.0,
                    )

                paper_alpha = max_alpha(
                    texture_backgrounds.get("paper", "")
                )
                for value in (
                    "grid", "grain",
                    "dots", "rings", "diagonal", "diamond",
                ):
                    with self.subTest(texture=value):
                        background = texture_backgrounds.get(value, "")
                        self.assertIn("gradient", background)
                        self.assertGreater(max_alpha(background), paper_alpha)
                        self.assertGreaterEqual(max_alpha(background), 0.07)
                motif_backgrounds = json.loads(
                    attributes.get("data-test-motif-backgrounds") or "{}"
                )
                self.assertEqual(
                    set(motif_backgrounds),
                    {"none", "books", "leaves", "stars"},
                )
                self.assertNotIn("svg", motif_backgrounds.get("none", ""))
                for motif in ("books", "leaves", "stars"):
                    with self.subTest(motif=motif):
                        self.assertIn(
                            "data:image/svg+xml",
                            motif_backgrounds.get(motif, ""),
                        )
                motif_flourishes = json.loads(
                    attributes.get("data-test-motif-flourishes") or "{}"
                )
                self.assertEqual(
                    set(motif_flourishes),
                    {"none", "books", "leaves", "stars"},
                )
                for target in ("header", "panel", "collapsible"):
                    with self.subTest(flourish_target=target):
                        self.assertNotIn(
                            "svg",
                            motif_flourishes.get("none", {}).get(target, ""),
                        )
                        for motif in ("books", "leaves", "stars"):
                            self.assertIn(
                                "data:image/svg+xml",
                                motif_flourishes.get(motif, {}).get(target, ""),
                            )
                self.assertEqual(
                    attributes.get("data-test-reset-motif"), "none"
                )
                self.assertEqual(
                    attributes.get("data-test-motif-fallback"), "none"
                )
                self.assertEqual(
                    attributes.get("data-test-motif-stored"), "stars"
                )
                self.assertEqual(
                    attributes.get("data-test-motif-applied"), "stars"
                )
                self.assertIn(
                    attributes.get("data-test-annual-state"),
                    {"reliable", "insufficient"},
                )
                self.assertIn(
                    attributes.get("data-test-annual-reliable"),
                    {"true", "false"},
                )
                self.assertEqual(
                    attributes.get("data-test-annual-panel-stored"), "open"
                )
                self.assertEqual(
                    attributes.get("data-collapsed-overflow"), "false"
                )
                self.assertEqual(
                    attributes.get("data-collapsed-overlap"), "0"
                )
                self.assertEqual(
                    attributes.get("data-collapsed-truncated"), ""
                )
                self.assertEqual(
                    attributes.get("data-expanded-overflow"), "false"
                )
                self.assertEqual(
                    attributes.get("data-expanded-overlap"), "0"
                )
                self.assertEqual(
                    attributes.get("data-expanded-truncated"), ""
                )
                self.assertEqual(
                    attributes.get("data-background-active"), "true"
                )
                self.assertEqual(
                    attributes.get("data-background-source"), "recent"
                )

    @unittest.skipUnless(
        CHROMIUM_BROWSER, "Chromium browser is not available"
    )
    def test_dashboard_reading_trend_metric_and_snapshot_details(
        self,
    ) -> None:
        assert CHROMIUM_BROWSER is not None
        document = CURRENT_DASHBOARD_PATH.read_text(encoding="utf-8")
        instrumentation = """
<script>
  window.addEventListener("load", () => {
    const wordsButton = document.querySelector(
      '[data-trend-metric-button="words"]'
    );
    wordsButton.click();
    document.documentElement.dataset.testTrendMetric =
      document.documentElement.dataset.trendMetric;
    document.documentElement.dataset.testTrendInvalidFallback =
      normalizeTrendMetric("invalid");

    const wordsView = document.querySelector(
      '[data-trend-series="words"]'
    );
    document.documentElement.dataset.testTrendMetricVisible =
      wordsView && !wordsView.hidden ? "true" : "false";
    document.documentElement.dataset.testTrendPathCount = String(
      document.querySelectorAll(
        '[data-trend-series="words"] .trend-line'
      ).length
    );
    document.documentElement.dataset.testTrendPointCount = String(
      document.querySelectorAll(
        '[data-trend-series="words"] [data-trend-point]'
      ).length
    );

    const point = document.querySelector(
      '[data-trend-series="words"] [data-trend-point]'
    );
    if (point) {
      point.dispatchEvent(new Event("focus"));
      document.documentElement.dataset.testTrendTooltip =
        document.documentElement.dataset.trendTooltip;
      document.documentElement.dataset.testTrendTooltipText =
        trendTooltip.textContent.trim();
      point.dispatchEvent(new Event("blur"));
    }

    trendDetails.open = false;
    trendDetails.dispatchEvent(new Event("toggle"));
    advancedPanel.open = false;
    appearancePanel.open = false;
    updateLayoutMetrics();
    document.documentElement.dataset.collapsedOverflow =
      document.documentElement.dataset.overflowX;
    document.documentElement.dataset.collapsedOverlap =
      document.documentElement.dataset.overlapCount;
    document.documentElement.dataset.collapsedTruncated =
      document.documentElement.dataset.truncatedNumbers;

    trendDetails.open = true;
    trendDetails.dispatchEvent(new Event("toggle"));
    advancedPanel.open = true;
    appearancePanel.open = true;
    updateLayoutMetrics();
    document.documentElement.dataset.expandedOverflow =
      document.documentElement.dataset.overflowX;
    document.documentElement.dataset.expandedOverlap =
      document.documentElement.dataset.overlapCount;
    document.documentElement.dataset.expandedTruncated =
      document.documentElement.dataset.truncatedNumbers;
    document.documentElement.dataset.testTrendDetailsOpen =
      trendDetails.open ? "true" : "false";
    document.documentElement.dataset.testTrendStoredMetric =
      readPreference("kindle-dashboard-trend-metric");
    document.documentElement.dataset.testTrendStoredDetails =
      readPreference("kindle-dashboard-trend-details");
  });
</script>
"""
        self.assertIn("</body>", document)
        instrumented = document.replace(
            "</body>", instrumentation + "</body>"
        )

        viewports = {
            (320, 568): None,
            (360, 640): None,
            (390, 844): None,
            (768, 900): None,
            (1440, 900): None,
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            test_html = Path(temp_dir) / "dashboard-trend.html"
            test_html.write_text(instrumented, encoding="utf-8")
            for viewport in viewports:
                width, height = viewport
                viewports[viewport] = browser_dom_snapshot(
                    CHROMIUM_BROWSER, test_html, width, height
                )

        for viewport, snapshot in viewports.items():
            attributes = html_root_attributes(str(snapshot))
            with self.subTest(viewport=viewport):
                self.assertEqual(
                    attributes.get("data-test-trend-metric"), "words"
                )
                self.assertEqual(
                    attributes.get("data-test-trend-invalid-fallback"),
                    "time",
                )
                self.assertEqual(
                    attributes.get("data-test-trend-metric-visible"), "true"
                )
                self.assertGreater(
                    int(attributes.get("data-test-trend-path-count") or "0"),
                    0,
                )
                self.assertGreater(
                    int(attributes.get("data-test-trend-point-count") or "0"),
                    1,
                )
                self.assertEqual(
                    attributes.get("data-test-trend-tooltip"), "visible"
                )
                self.assertTrue(
                    str(attributes.get("data-test-trend-tooltip-text") or "")
                    .strip()
                )
                self.assertEqual(
                    attributes.get("data-test-trend-details-open"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-trend-stored-metric"),
                    "words",
                )
                self.assertEqual(
                    attributes.get("data-test-trend-stored-details"),
                    "open",
                )
                self.assertEqual(
                    attributes.get("data-collapsed-overflow"), "false"
                )
                self.assertEqual(
                    attributes.get("data-collapsed-overlap"), "0"
                )
                self.assertEqual(
                    attributes.get("data-collapsed-truncated"), ""
                )
                self.assertEqual(
                    attributes.get("data-expanded-overflow"), "false"
                )
                self.assertEqual(
                    attributes.get("data-expanded-overlap"), "0"
                )
                self.assertEqual(
                    attributes.get("data-expanded-truncated"), ""
                )

    @unittest.skipUnless(
        CHROMIUM_BROWSER, "Chromium browser is not available"
    )
    def test_dashboard_theme_language_columns_and_export_controls(
        self,
    ) -> None:
        assert CHROMIUM_BROWSER is not None
        with CURRENT_CSV_PATH.open(
            "r", encoding="utf-8-sig", newline=""
        ) as handle:
            reader = csv.DictReader(handle)
            fields = list(reader.fieldnames or [])
            rows = [dict(row) for row in reader]
        document = kindle_report.build_html(
            rows,
            Path("reading.csv"),
            "Kindle",
            generated_at="2026-10-01T12:00:00+08:00",
            xlsx_path=Path("reading.xlsx"),
        )
        instrumentation = """
<script>
  window.addEventListener("load", () => {
    themeToggle.click();
    languageToggle.click();
    const qualityToggle = document.querySelector(
      '[data-column-toggle="quality"]'
    );
    qualityToggle.checked = false;
    qualityToggle.dispatchEvent(new Event("change"));

    languageFilter.value = "en";
    applyFilters();
    document.documentElement.dataset.testEnglishRows =
      document.documentElement.dataset.visibleRows;
    languageFilter.value = "";
    countFilter.value = "estimate";
    applyFilters();
    document.documentElement.dataset.testEstimateRows =
      document.documentElement.dataset.visibleRows;
    countFilter.value = "";
    applyFilters();
    document.querySelectorAll("#data-table th")[4].click();
    const firstBook = document.querySelector(
      "#data-table tbody tr strong"
    );
    document.documentElement.dataset.testFirstBook =
      firstBook ? firstBook.textContent.trim() : "";
    document.documentElement.dataset.testTheme =
      document.documentElement.dataset.theme;
    document.documentElement.dataset.testLanguage =
      document.documentElement.dataset.uiLanguage;
    document.documentElement.dataset.testHtmlLanguage =
      document.documentElement.lang;
    document.documentElement.dataset.testThemePressed =
      themeToggle.getAttribute("aria-pressed");
    document.documentElement.dataset.testQualityColumnHidden =
      String(document.querySelector(
        'th[data-column="quality"]'
      ).hidden);
    document.documentElement.dataset.testChineseSpanDisplay =
      getComputedStyle(document.querySelector(
        "#data-table th .i18n-zh"
      )).display;
    document.documentElement.dataset.testExportHref =
      document.getElementById("xlsx-export").getAttribute("href");
    updateLayoutMetrics();
  });
</script>
"""
        self.assertIn("</body>", document)
        instrumented = document.replace(
            "</body>", instrumentation + "</body>"
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            test_html = Path(temp_dir) / "dashboard-controls.html"
            test_html.write_text(instrumented, encoding="utf-8")
            snapshots = {
                (390, 844): browser_dom_snapshot(
                    CHROMIUM_BROWSER, test_html, 390, 844
                ),
                (1440, 900): browser_dom_snapshot(
                    CHROMIUM_BROWSER, test_html, 1440, 900
                ),
            }

        for viewport, snapshot in snapshots.items():
            attributes = html_root_attributes(snapshot)
            with self.subTest(viewport=viewport):
                self.assertEqual(
                    attributes.get("data-overflow-x"), "false"
                )
                self.assertEqual(
                    attributes.get("data-overflow-elements"), ""
                )
                self.assertEqual(attributes.get("data-test-theme"), "dark")
                self.assertEqual(
                    attributes.get("data-test-language"), "en"
                )
                self.assertEqual(
                    attributes.get("data-test-html-language"), "en"
                )
                self.assertEqual(
                    attributes.get("data-test-theme-pressed"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-quality-column-hidden"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-visible-columns"), "11"
                )
                self.assertEqual(
                    attributes.get("data-test-chinese-span-display"),
                    "none",
                )
                self.assertEqual(
                    attributes.get("data-test-export-href"),
                    "reading.xlsx",
                )
                self.assertEqual(
                    attributes.get("data-test-english-rows"), "2"
                )
                self.assertEqual(
                    attributes.get("data-test-estimate-rows"), "0"
                )
                self.assertEqual(
                    attributes.get("data-test-first-book"), "合成中文测试书"
                )

    @unittest.skipUnless(
        CHROMIUM_BROWSER, "Chromium browser is not available"
    )
    def test_dashboard_heatmap_fits_mobile_and_desktop(self) -> None:
        assert CHROMIUM_BROWSER is not None
        end_date = date.fromisoformat("2026-09-30")
        daily = [
            {
                "date": (end_date - timedelta(days=offset)).isoformat(),
                "total_time_ms": offset * 60000,
                "active": offset > 0,
            }
            for offset in range(84)
        ]
        with CURRENT_CSV_PATH.open(
            "r", encoding="utf-8-sig", newline=""
        ) as handle:
            rows = list(csv.DictReader(handle))
        document = kindle_report.build_html(
            rows,
            Path("reading.csv"),
            "Kindle",
            generated_at="2026-09-30T12:00:00+08:00",
            heatmap={
                "days": 84,
                "start": daily[0]["date"],
                "end": daily[-1]["date"],
                "source": "history",
                "source_label": "历史快照回退",
                "status": "empty",
                "message": "reading_sessions 为空",
                "db_path": "",
                "session_count": 0,
                "daily": daily,
            },
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            test_html = Path(temp_dir) / "heatmap-dashboard.html"
            test_html.write_text(document, encoding="utf-8")
            snapshots = {
                (390, 844): browser_dom_snapshot(
                    CHROMIUM_BROWSER, test_html, 390, 844
                ),
                (1440, 900): browser_dom_snapshot(
                    CHROMIUM_BROWSER, test_html, 1440, 900
                ),
            }

        for viewport, snapshot in snapshots.items():
            attributes = html_root_attributes(snapshot)
            with self.subTest(viewport=viewport):
                self.assertEqual(
                    attributes.get("data-overflow-x"), "false"
                )
                self.assertEqual(
                    attributes.get("data-overflow-elements"), ""
                )

    def test_nested_timer_model_and_stats(self) -> None:
        logger = kindle_reader._make_logger(False)
        with tempfile.TemporaryDirectory() as temp_dir:
            sidecar = Path(temp_dir) / "Book.azw3f"
            sidecar.write_bytes(synthetic_sidecar())
            decoded = kindle_reader.decode_sidecar(str(sidecar), logger)

        timer_model = kindle_reader._nested_timer_model(decoded)
        self.assertIsNotNone(timer_model)
        self.assertEqual(timer_model["totalTime"], 600000)
        self.assertEqual(timer_model["totalWords"], 3000)

        stats = kindle_reader._stats_from_timer(
            timer_model, decoded["book.info.store"]
        )
        self.assertEqual(stats["total_time_ms"], 600000)
        self.assertEqual(stats["total_words"], 3000)
        self.assertEqual(stats["minutes"], 10.0)
        self.assertEqual(stats["hours"], 0.17)
        self.assertEqual(stats["wpm_overall"], 300)
        self.assertEqual(stats["wpm_reading"], 300)
        self.assertEqual(stats["percent_read"], 25.0)
        self.assertEqual(stats["book_words"], 9000)

    def test_book_words_are_scaled_by_percent_of_book(self) -> None:
        stats = kindle_reader._stats_from_timer(
            {
                "totalTime": 60000,
                "totalWords": 200,
                "totalPercent": 0.01,
                "averageCalculator": {"normalDistributions": []},
            },
            {
                "numberOfWords": 100,
                "percentOfBook": 0.01,
            },
        )

        self.assertEqual(stats["book_words_known"], 100)
        self.assertEqual(stats["book_word_coverage_percent"], 1.0)
        self.assertEqual(stats["book_words"], 10000)

    def test_word_count_tokenizers(self) -> None:
        english_count, english_source = kindle_reader.count_text_words(
            "Hello world, it's fine.", "en"
        )
        chinese_count, chinese_source = kindle_reader.count_text_words(
            "这是中文自然语言处理测试。", "zh"
        )
        japanese_count, japanese_source = kindle_reader.count_text_words(
            "これはテストです。", "ja"
        )

        self.assertEqual(english_count, 4)
        self.assertEqual(english_source, "en-regex")
        self.assertGreater(chinese_count, 0)
        self.assertEqual(chinese_source, "jieba")
        self.assertGreater(japanese_count, 0)
        self.assertEqual(japanese_source, "janome")

    def test_manga_ocr_cache_key_and_source_resolution(self) -> None:
        first = kindle_manga_ocr.ocr_cache_key(
            "source-a", 300, "ja", "manga_ocr", "1.0", "page-a"
        )
        self.assertEqual(
            first,
            kindle_manga_ocr.ocr_cache_key(
                "source-a", 300, "ja", "manga_ocr", "1.0", "page-a"
            ),
        )
        for changed in (
            kindle_manga_ocr.ocr_cache_key(
                "source-b", 300, "ja", "manga_ocr", "1.0", "page-a"
            ),
            kindle_manga_ocr.ocr_cache_key(
                "source-a", 350, "ja", "manga_ocr", "1.0", "page-a"
            ),
            kindle_manga_ocr.ocr_cache_key(
                "source-a", 300, "zh", "manga_ocr", "1.0", "page-a"
            ),
            kindle_manga_ocr.ocr_cache_key(
                "source-a", 300, "ja", "paddleocr", "1.0", "page-a"
            ),
            kindle_manga_ocr.ocr_cache_key(
                "source-a", 300, "ja", "manga_ocr", "2.0", "page-a"
            ),
            kindle_manga_ocr.ocr_cache_key(
                "source-a", 300, "ja", "manga_ocr", "1.0", "page-b"
            ),
        ):
            self.assertNotEqual(first, changed)
        self.assertEqual(
            kindle_manga_ocr.resolve_ocr_language(
                {"title": "合成漫画测试卷 Vol.04"},
                "auto",
            ),
            "zh",
        )
        with mock.patch.object(
            kindle_manga_ocr,
            "engine_available",
            side_effect=lambda name: (name == "paddleocr", ""),
        ):
            self.assertEqual(
                kindle_manga_ocr.resolve_ocr_engine("auto", "zh"),
                "paddleocr",
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "Kindle"
            sidecar_dir = (
                root
                / "documents"
                / "Manga Vol.04_A0000000000000000000000000000005.sdr"
            )
            sidecar_dir.mkdir(parents=True)
            source = (
                root
                / "documents"
                / "Manga Vol.04_A0000000000000000000000000000005.kfx"
            )
            source.write_bytes(b"manga")
            sidecar = sidecar_dir / "Manga Vol.04.azw3f"
            sidecar.write_bytes(b"sidecar")
            resolved = kindle_manga_ocr.resolve_ocr_source(
                {
                    "title": "Manga Vol.04",
                    "asin": "A0000000000000000000000000000005",
                    "source_file": str(sidecar),
                }
            )
            first_hash = kindle_manga_ocr.source_file_hash(source)
            second_hash = kindle_manga_ocr.source_file_hash(source)

        self.assertEqual(resolved, source.resolve())
        self.assertEqual(first_hash, second_hash)

    def test_manga_ocr_finds_project_side_poppler_without_path(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            runtime = root / "dependencies" / "python" / "python.exe"
            runtime.parent.mkdir(parents=True)
            runtime.write_bytes(b"synthetic-python")
            tool = (
                root
                / "dependencies"
                / "native"
                / "poppler"
                / "Library"
                / "bin"
                / "pdftoppm.exe"
            )
            tool.parent.mkdir(parents=True)
            tool.write_bytes(b"synthetic-pdftoppm")

            with mock.patch.object(
                kindle_manga_ocr.sys,
                "executable",
                str(runtime),
            ), mock.patch.object(
                kindle_manga_ocr,
                "_APPLICATION_ROOTS",
                [root],
            ), mock.patch.object(
                kindle_manga_ocr,
                "find_calibre_convert",
                return_value=None,
            ), mock.patch.dict(
                os.environ,
                {"PATH": "", "KINDLE_OCR_PDFTOPPM": ""},
            ):
                resolved = kindle_manga_ocr.find_pdftoppm()

        self.assertEqual(resolved, tool.resolve())

    def test_manga_ocr_render_manifest_reuses_synthetic_pages(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "Manga.cbz"
            source.write_bytes(b"synthetic comic")
            pdf_path = root / "rendered.pdf"
            pdf_path.write_bytes(b"%PDF-1.4 synthetic")
            cache_dir = root / "_ocr"
            calls: List[List[str]] = []

            def fake_run(command, timeout, logger=None):
                del timeout, logger
                values = list(command)
                calls.append(values)
                prefix = Path(values[-1])
                prefix.parent.mkdir(parents=True, exist_ok=True)
                (prefix.parent / (prefix.name + "-1.png")).write_bytes(
                    b"synthetic-page-1"
                )
                (prefix.parent / (prefix.name + "-2.png")).write_bytes(
                    b"synthetic-page-2"
                )
                return subprocess.CompletedProcess(values, 0, "")

            with mock.patch.object(
                kindle_manga_ocr,
                "find_pdftoppm",
                return_value=Path("pdftoppm.exe"),
            ), mock.patch.object(
                kindle_manga_ocr,
                "_convert_to_pdf",
                return_value=pdf_path,
            ), mock.patch.object(
                kindle_manga_ocr,
                "_pdf_page_count",
                return_value=2,
            ), mock.patch.object(
                kindle_manga_ocr,
                "_run_command",
                side_effect=fake_run,
            ):
                first = kindle_manga_ocr.render_manga_pages(
                    source,
                    cache_dir,
                    300,
                    source_hash="source-hash",
                )
                second = kindle_manga_ocr.render_manga_pages(
                    source,
                    cache_dir,
                    300,
                    source_hash="source-hash",
                )
            manifest_count = len(list(cache_dir.glob("render/*.json")))
            page_count = len(list(cache_dir.glob("pages/*.png")))

        self.assertEqual(len(first.pages), 2)
        self.assertEqual([page.index for page in first.pages], [1, 2])
        self.assertEqual([page.sha256 for page in first.pages],
                         [page.sha256 for page in second.pages])
        self.assertEqual(len(calls), 1)
        self.assertTrue(manifest_count)
        self.assertTrue(page_count)

    def test_manga_ocr_estimate_cache_failure_and_csv_append_only(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "Manga.kfx"
            source.write_bytes(b"manga source")
            page_one = root / "page-1.png"
            page_two = root / "page-2.png"
            page_one.write_bytes(b"page-one")
            page_two.write_bytes(b"page-two")
            rendered = kindle_manga_ocr.RenderedBook(
                source_path=source,
                source_hash="source-hash",
                pdf_path=root / "render.pdf",
                dpi=300,
                pages=(
                    kindle_manga_ocr.RenderedPage(
                        1,
                        page_one,
                        "page-hash-1",
                    ),
                    kindle_manga_ocr.RenderedPage(
                        2,
                        page_two,
                        "page-hash-2",
                    ),
                ),
                expected_pages=2,
            )

            class FakeEngine:
                version = "test-1"

                def __init__(self) -> None:
                    self.calls: List[int] = []

                def recognize(self, image_path: Path, language: str):
                    del language
                    index = 1 if "1" in image_path.name else 2
                    self.calls.append(index)
                    return kindle_manga_ocr.OcrPageResult(
                        "漫画 文本" if index == 1 else "第二页",
                        0.95,
                    )

            engine = FakeEngine()
            config = kindle_manga_ocr.OcrConfig(
                language="ja",
                engine="manga_ocr",
                dpi=300,
                min_confidence=0.7,
                cache_dir=root / "_ocr",
            )
            row = {
                "title": "Manga",
                "asin": "B000000501",
                "reading_mode": "manga",
                "book_words_exact": "9999",
                "total_words": "0",
            }
            with mock.patch.object(
                kindle_manga_ocr,
                "engine_available",
                return_value=(True, ""),
            ):
                first = kindle_manga_ocr.estimate_manga_ocr_words(
                    row,
                    config,
                    source_resolver=lambda current, base: source,
                    renderer=lambda *args, **kwargs: rendered,
                    engine_factory=lambda name, language: engine,
                )
                second = kindle_manga_ocr.estimate_manga_ocr_words(
                    row,
                    config,
                    source_resolver=lambda current, base: source,
                    renderer=lambda *args, **kwargs: rendered,
                    engine_factory=lambda name, language: engine,
                )
            merged = kindle_manga_ocr.merge_ocr_estimate(row, first)

            csv_path = root / "reading.csv"
            old_fields = [
                "title",
                "asin",
                "total_words",
                "book_words_exact",
                "reading_mode",
            ]
            with csv_path.open(
                "w",
                encoding="utf-8-sig",
                newline="",
            ) as handle:
                writer = csv.DictWriter(handle, fieldnames=old_fields)
                writer.writeheader()
                writer.writerow(
                    {
                        "title": "Manga",
                        "asin": "B000000501",
                        "total_words": "0",
                        "book_words_exact": "",
                        "reading_mode": "manga",
                    }
                )
            class AlwaysOkEngine(FakeEngine):
                version = "test-append"

            with mock.patch.object(
                kindle_manga_ocr,
                "engine_available",
                return_value=(True, ""),
            ):
                counts = kindle_manga_ocr.enrich_csv_with_manga_ocr(
                    csv_path,
                    config,
                    source_resolver=lambda current, base: source,
                    renderer=lambda *args, **kwargs: rendered,
                    engine_factory=lambda name, language: AlwaysOkEngine(),
                )
            with csv_path.open(
                "r",
                encoding="utf-8-sig",
                newline="",
            ) as handle:
                reader = csv.DictReader(handle)
                header = list(reader.fieldnames or [])
                csv_rows = list(reader)

        self.assertEqual(first["ocr_status"], "ok")
        self.assertEqual(first["ocr_page_coverage"], 100.0)
        self.assertGreater(first["ocr_word_count"], 0)
        self.assertEqual(engine.calls, [1, 2])
        self.assertTrue(second["_ocr_cache_hit"])
        self.assertEqual(merged["book_words_exact"], "9999")
        self.assertEqual(merged["ocr_source"], "ocr-estimate")
        self.assertEqual(header[: len(old_fields)], old_fields)
        self.assertTrue(
            all(field in header for field in kindle_manga_ocr.OCR_FIELDS)
        )
        self.assertEqual(counts["manga_rows"], 1)
        self.assertEqual(csv_rows[0]["book_words_exact"], "")
        self.assertEqual(csv_rows[0]["total_words"], "0")

    def test_manga_ocr_failure_is_reported_and_not_used_by_word_metrics(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "Manga.kfx"
            source.write_bytes(b"manga")
            page_one = root / "page-1.png"
            page_two = root / "page-2.png"
            page_one.write_bytes(b"one")
            page_two.write_bytes(b"two")
            rendered = kindle_manga_ocr.RenderedBook(
                source_path=source,
                source_hash="failure-source",
                pdf_path=root / "render.pdf",
                dpi=300,
                pages=(
                    kindle_manga_ocr.RenderedPage(1, page_one, "p1"),
                    kindle_manga_ocr.RenderedPage(2, page_two, "p2"),
                ),
                expected_pages=2,
            )

            class FailingEngine:
                version = "fail-1"

                def recognize(self, image_path: Path, language: str):
                    del language
                    if "2" in image_path.name:
                        raise RuntimeError("synthetic page failure")
                    return kindle_manga_ocr.OcrPageResult("第一页", 0.91)

            config = kindle_manga_ocr.OcrConfig(
                language="ja",
                engine="manga_ocr",
                dpi=300,
                min_confidence=0.7,
                cache_dir=root / "_ocr",
            )
            row = {
                "title": "Manga Failure",
                "asin": "B000000502",
                "reading_mode": "manga",
                "book_words_exact": "",
                "total_words": "0",
                "wpm": "0",
                "total_time_ms": "600000",
                "progress_percent": "50",
            }
            with mock.patch.object(
                kindle_manga_ocr,
                "engine_available",
                return_value=(True, ""),
            ):
                estimate = kindle_manga_ocr.estimate_manga_ocr_words(
                    row,
                    config,
                    source_resolver=lambda current, base: source,
                    renderer=lambda *args, **kwargs: rendered,
                    engine_factory=lambda name, language: FailingEngine(),
                )
            merged = kindle_manga_ocr.merge_ocr_estimate(row, estimate)
            merged["_fmcache_fallback"] = True
            document = kindle_report.build_html(
                [merged],
                Path("reading.csv"),
                "Kindle",
                generated_at="2026-10-02T12:00:00+08:00",
            )
            summary = kindle_report.build_summary([merged])

        self.assertEqual(estimate["ocr_status"], "partial")
        self.assertEqual(estimate["ocr_pages_processed"], 1)
        self.assertEqual(estimate["ocr_failed_pages"], 1)
        self.assertLess(estimate["ocr_page_coverage"], 98.0)
        self.assertIn("synthetic page failure", estimate["ocr_error"])
        self.assertEqual(summary["total_words"], 0)
        self.assertEqual(summary["word_books"], 0)
        self.assertIsNone(kindle_report.word_progress(merged))
        self.assertEqual(merged["book_words_exact"], "")
        self.assertIn('id="ocr-panel"', document)
        self.assertIn("OCR 估算", document)
        self.assertIn("fmcache 回退", document)
        self.assertIn("仅作为离线估算", document)

    def test_ocr_effective_metrics_join_summary_trends_and_achievements(
        self,
    ) -> None:
        row = {
            "title": "OCR Manga",
            "asin": "B000000601",
            "reading_mode": "manga",
            "total_time_ms": "3600000",
            "total_words": "0",
            "book_words_exact": "",
            "wpm": "0",
            "progress_percent": "50",
            "ocr_word_count": "6000",
            "ocr_char_count": "9000",
            "ocr_confidence": "0.91",
            "ocr_engine": "paddleocr",
            "ocr_language": "zh",
            "ocr_page_coverage": "100",
            "ocr_status": "ok",
            "ocr_source": "ocr-estimate",
            "ocr_pages_total": "20",
            "ocr_pages_processed": "20",
            "ocr_failed_pages": "0",
            "ocr_min_confidence": "0.70",
        }
        self.assertEqual(kindle_report.effective_book_words(row), 6000.0)
        self.assertEqual(kindle_report.effective_words_read(row), 3000.0)
        self.assertEqual(kindle_report.effective_wpm(row), 50.0)
        self.assertEqual(kindle_report.word_progress(row), 50.0)
        self.assertEqual(kindle_report.effective_source(row), "ocr-estimate")

        summary = kindle_report.build_summary([row])
        self.assertEqual(summary["total_words"], 0)
        self.assertEqual(summary["effective_words_read"], 3000.0)
        self.assertEqual(summary["weighted_wpm"], 50.0)
        self.assertEqual(summary["word_books"], 1)
        self.assertEqual(summary["ocr_estimate_books"], 1)

        snapshot = kindle_report.build_snapshot(
            [row],
            "2026-10-02T12:00:00+08:00",
            "Kindle",
        )
        totals = kindle_report.snapshot_totals(snapshot)
        self.assertEqual(totals["total_words"], 0)
        self.assertEqual(totals["effective_words_read"], 3000.0)
        self.assertEqual(totals["weighted_wpm"], 50.0)

        trends = {
            "daily": [
                {
                    "date": "2026-10-02",
                    "total_time_ms": 3600000,
                    "total_words": 3000,
                    "active": True,
                }
            ],
            "streak": {"current_days": 1, "longest_days": 1},
        }
        achievements = kindle_report.build_achievements(
            summary,
            trends,
            [row],
        )
        by_key = {
            item["key"]: item for item in achievements["items"]
        }
        self.assertEqual(by_key["words"]["value"], 3000.0)
        self.assertEqual(by_key["speed"]["value"], 50.0)

    def test_low_confidence_or_incomplete_ocr_is_excluded(self) -> None:
        low_confidence = {
            "title": "Low Confidence",
            "asin": "B000000602",
            "reading_mode": "manga",
            "total_time_ms": "600000",
            "total_words": "0",
            "progress_percent": "50",
            "ocr_word_count": "1000",
            "ocr_char_count": "1800",
            "ocr_confidence": "0.62",
            "ocr_page_coverage": "100",
            "ocr_status": "ocr-low-confidence",
            "ocr_pages_total": "10",
            "ocr_pages_processed": "10",
        }
        incomplete = {
            "title": "Partial Pages",
            "asin": "B000000603",
            "reading_mode": "manga",
            "total_time_ms": "600000",
            "total_words": "0",
            "progress_percent": "50",
            "ocr_word_count": "1000",
            "ocr_char_count": "1800",
            "ocr_confidence": "0.91",
            "ocr_page_coverage": "90",
            "ocr_status": "partial",
            "ocr_pages_total": "10",
            "ocr_pages_processed": "9",
        }

        self.assertFalse(kindle_report.ocr_estimate_eligible(low_confidence))
        self.assertFalse(kindle_report.ocr_estimate_eligible(incomplete))
        summary = kindle_report.build_summary([low_confidence, incomplete])
        self.assertEqual(summary["effective_words_read"], 0)
        self.assertEqual(summary["word_books"], 0)
        self.assertEqual(summary["ocr_estimate_books"], 0)
        self.assertIsNone(kindle_report.effective_book_words(low_confidence))
        self.assertIsNone(kindle_report.word_progress(incomplete))

    def test_old_snapshot_without_ocr_fields_remains_compatible(self) -> None:
        old_snapshot = {
            "captured_at": "2026-09-30T12:00:00+08:00",
            "source": "Kindle",
            "books": {
                "asin:B000000604": {
                    "book_id": "asin:B000000604",
                    "title": "Old Book",
                    "total_time_ms": 600000,
                    "total_words": 5000,
                    "wpm": 500,
                    "progress_percent": 50,
                }
            },
        }
        totals = kindle_report.snapshot_totals(old_snapshot)
        self.assertEqual(totals["total_words"], 5000)
        self.assertEqual(totals["effective_words_read"], 5000)
        self.assertEqual(totals["weighted_wpm"], 500.0)

    def test_repeated_ocr_snapshot_rebuild_is_deduplicated(self) -> None:
        row = {
            "title": "Stable OCR Manga",
            "asin": "B000000605",
            "reading_mode": "manga",
            "total_time_ms": "3600000",
            "total_words": "0",
            "progress_percent": "50",
            "ocr_word_count": "6000",
            "ocr_char_count": "9000",
            "ocr_confidence": "0.9",
            "ocr_page_coverage": "100",
            "ocr_status": "ok",
            "ocr_pages_total": "20",
            "ocr_pages_processed": "20",
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            history_path = Path(temp_dir) / kindle_report.HISTORY_FILENAME
            first = kindle_report.build_snapshot(
                [row],
                "2026-10-02T12:00:00+08:00",
                "Kindle",
            )
            second = kindle_report.build_snapshot(
                [row],
                "2026-10-02T12:05:00+08:00",
                "Kindle",
            )
            self.assertTrue(
                kindle_report.append_snapshot(history_path, first)
            )
            self.assertFalse(
                kindle_report.append_snapshot(history_path, second)
            )
            snapshots = kindle_report.load_history(history_path)["snapshots"]

        self.assertEqual(len(snapshots), 1)

    def test_disabling_ocr_does_not_append_ocr_blind_snapshot(self) -> None:
        ocr_row = {
            "title": "OCR Manga",
            "asin": "B000000608",
            "reading_mode": "manga",
            "total_time_ms": "3600000",
            "total_words": "0",
            "progress_percent": "50",
            "ocr_word_count": "6000",
            "ocr_char_count": "9000",
            "ocr_confidence": "0.9",
            "ocr_page_coverage": "100",
            "ocr_status": "ok",
            "ocr_pages_total": "20",
            "ocr_pages_processed": "20",
        }
        plain_row = dict(ocr_row)
        for key in list(plain_row):
            if key.startswith("ocr_"):
                plain_row[key] = ""
        with tempfile.TemporaryDirectory() as temp_dir:
            history_path = Path(temp_dir) / kindle_report.HISTORY_FILENAME
            ocr_snapshot = kindle_report.build_snapshot(
                [ocr_row],
                "2026-10-02T12:00:00+08:00",
                "Kindle",
            )
            plain_snapshot = kindle_report.build_snapshot(
                [plain_row],
                "2026-10-02T12:05:00+08:00",
                "Kindle",
            )
            self.assertTrue(
                kindle_report.append_snapshot(history_path, ocr_snapshot)
            )
            self.assertFalse(
                kindle_report.append_snapshot(history_path, plain_snapshot)
            )
            snapshots = kindle_report.load_history(history_path)["snapshots"]

        self.assertEqual(len(snapshots), 1)
        self.assertTrue(
            kindle_report._snapshot_has_ocr_effective(snapshots[0])
        )

    def test_ocr_blind_snapshot_can_append_new_manga_progress(self) -> None:
        ocr_row = {
            "title": "Position Manga",
            "asin": "B000000609",
            "reading_mode": "manga",
            "total_time_ms": "600000",
            "progress_percent": "50",
            "manga_pages_total": "100",
            "manga_pages_read": "50",
            "ocr_word_count": "6000",
            "ocr_char_count": "9000",
            "ocr_confidence": "0.9",
            "ocr_page_coverage": "100",
            "ocr_status": "ok",
            "ocr_pages_total": "20",
            "ocr_pages_processed": "20",
        }
        improved_row = dict(ocr_row)
        for key in list(improved_row):
            if key.startswith("ocr_"):
                improved_row[key] = ""
        improved_row["progress_percent"] = "100"
        improved_row["manga_pages_read"] = "100"
        improved_row["manga_progress_source"] = "kfx-position-map"
        with tempfile.TemporaryDirectory() as temp_dir:
            history_path = Path(temp_dir) / kindle_report.HISTORY_FILENAME
            ocr_snapshot = kindle_report.build_snapshot(
                [ocr_row],
                "2026-10-02T12:00:00+08:00",
                "Kindle",
            )
            improved_snapshot = kindle_report.build_snapshot(
                [improved_row],
                "2026-10-02T12:05:00+08:00",
                "Kindle",
            )
            self.assertTrue(
                kindle_report.append_snapshot(history_path, ocr_snapshot)
            )
            self.assertTrue(
                kindle_report.append_snapshot(history_path, improved_snapshot)
            )
            snapshots = kindle_report.load_history(history_path)["snapshots"]

        self.assertEqual(len(snapshots), 2)
        self.assertEqual(
            snapshots[-1]["books"]["asin:B000000609"]["progress_percent"],
            100.0,
        )

    def test_ocr_achievement_profile_never_regresses(self) -> None:
        row = {
            "title": "OCR Manga",
            "asin": "B000000606",
            "reading_mode": "manga",
            "total_time_ms": "3600000",
            "total_words": "0",
            "progress_percent": "50",
            "ocr_word_count": "6000",
            "ocr_char_count": "9000",
            "ocr_confidence": "0.9",
            "ocr_page_coverage": "100",
            "ocr_status": "ok",
            "ocr_pages_total": "20",
            "ocr_pages_processed": "20",
        }
        base = kindle_report.build_achievement_profile(
            [row],
            {"version": kindle_report.HISTORY_VERSION, "snapshots": []},
            captured_at="2026-10-02T12:00:00+08:00",
            threshold_percent=95.0,
        )
        base["levels"]["words"] = 1
        base["observed_values"]["words"] = 6000.0
        updated = kindle_report.build_achievement_profile(
            [row],
            {"version": kindle_report.HISTORY_VERSION, "snapshots": []},
            captured_at="2026-10-03T12:00:00+08:00",
            existing_profile=base,
            threshold_percent=95.0,
        )
        self.assertGreaterEqual(updated["levels"]["words"], 1)
        self.assertGreaterEqual(updated["observed_values"]["words"], 6000.0)

    def test_dashboard_marks_ocr_estimate_and_gold_rounded_stars(self) -> None:
        row = {
            "title": "OCR Manga",
            "author": "Author",
            "asin": "B000000607",
            "reading_mode": "manga",
            "total_time_ms": "3600000",
            "total_words": "0",
            "wpm": "0",
            "progress_percent": "50",
            "ocr_word_count": "6000",
            "ocr_char_count": "9000",
            "ocr_confidence": "0.9",
            "ocr_engine": "paddleocr",
            "ocr_language": "zh",
            "ocr_page_coverage": "100",
            "ocr_status": "ok",
            "ocr_source": "ocr-estimate",
            "ocr_pages_total": "20",
            "ocr_pages_processed": "20",
        }
        document = kindle_report.build_html(
            [row],
            Path("reading.csv"),
            "Kindle",
            generated_at="2026-10-02T12:00:00+08:00",
        )

        self.assertIn("OCR 估算", document)
        self.assertIn("--gold: #FFD400", document)
        self.assertIn("--gold-bright: #FFE566", document)
        self.assertIn("--gold-deep: #B7791F", document)
        self.assertIn('stroke-linejoin="round"', document)
        self.assertIn('stroke-linecap="round"', document)
        self.assertIn('class="achievement-star', document)
        self.assertIn('tabindex="0"', document)
        self.assertIn("print-color-adjust: exact", document)
        self.assertNotIn("★", document)
        self.assertNotIn("☆", document)

    def test_report_prefers_exact_word_count(self) -> None:
        progress = kindle_report.word_progress(
            {
                "total_words": "500",
                "book_words": "1000",
                "book_words_exact": "5000",
            }
        )
        self.assertEqual(progress, 10.0)

    def test_path_resolution_and_csv_export(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "Kindle"
            sidecar_dir = root / "documents" / "The Hobbit - Tolkien.sdr"
            sidecar_dir.mkdir(parents=True)
            sidecar = sidecar_dir / "The Hobbit - Tolkien.azw3f"
            sidecar.write_bytes(synthetic_sidecar())
            output = Path(temp_dir) / "reading.csv"

            documents = kindle_reader.resolve_documents(str(root))
            self.assertEqual(
                Path(documents).resolve(),
                (root / "documents").resolve(),
            )

            exit_code = kindle_reader.main(
                [
                    "--path",
                    str(root),
                    "--output",
                    str(output),
                ]
            )
            self.assertEqual(exit_code, 0)

            with output.open("r", encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.DictReader(handle))

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["title"], "The Hobbit")
        self.assertEqual(rows[0]["author"], "Tolkien")
        self.assertEqual(rows[0]["total_time_ms"], "600000")
        self.assertEqual(rows[0]["total_words"], "3000")
        self.assertEqual(rows[0]["wpm"], "300")
        self.assertEqual(rows[0]["wpm_overall"], "300")
        self.assertEqual(rows[0]["progress_percent"], "25.0")
        self.assertEqual(rows[0]["has_data"], "True")

    def test_reader_applies_book_modes_without_silent_classification(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "Kindle"
            sidecar_dir = root / "documents" / "Manga Book.sdr"
            sidecar_dir.mkdir(parents=True)
            (sidecar_dir / "Manga Book.azw3f").write_bytes(
                synthetic_sidecar()
            )
            output = Path(temp_dir) / "reading.csv"
            modes_path = Path(temp_dir) / "book_modes.json"
            modes_path.write_text(
                json.dumps({"by_title": {"Manga Book": "manga"}}),
                encoding="utf-8",
            )

            exit_code = kindle_reader.main(
                [
                    "--path",
                    str(root),
                    "--output",
                    str(output),
                    "--book-modes",
                    str(modes_path),
                ]
            )
            with output.open("r", encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.DictReader(handle))

        self.assertEqual(exit_code, 0)
        self.assertEqual(rows[0]["reading_mode"], "manga")
        self.assertEqual(
            kindle_reader.suggested_reading_mode_from_metadata(
                {"tags": ["Comics"]}
            ),
            "manga",
        )
        self.assertIsNone(
            kindle_reader.suggested_reading_mode_from_metadata(
                {"tags": ["Fiction"]}
            )
        )
        self.assertEqual(
            kindle_reader.reading_mode_for_book(
                {"title": "Unmapped Comic", "asin": "B000000301"},
                {"by_title": {}},
            ),
            "book",
        )

    def test_reader_inherits_manga_mode_for_series_volumes(self) -> None:
        mapping = {"by_title": {"合成漫画测试卷 Vol.04": "manga"}}
        storage_id = "F000000000000000000000000000000A"
        self.assertEqual(
            kindle_reader.reading_mode_for_book(
                {"title": "合成漫画测试卷 Vol.05_%s" % storage_id},
                mapping,
            ),
            "manga",
        )
        self.assertEqual(
            kindle_report.reading_mode_for_book(
                {"title": "合成漫画测试卷 Vol.05_%s" % storage_id},
                mapping,
            ),
            "manga",
        )
        self.assertEqual(
            kindle_reader.reading_mode_for_book(
                {"title": "合成漫画测试卷 Vol.05"},
                mapping,
            ),
            "manga",
        )
        self.assertEqual(
            kindle_reader.reading_mode_for_book(
                {"title": "Unrelated Book Vol.01"},
                mapping,
            ),
            "book",
        )

    def test_generated_kindle_documents_are_filtered_and_manga_keeps_source(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "Kindle"
            documents = root / "documents"
            manga_dir = documents / "Manga Vol.04.sdr"
            manga_dir.mkdir(parents=True)
            source = documents / "Manga Vol.04.kfx"
            source.write_bytes(b"manga")
            (manga_dir / "Manga Vol.04.azw3f").write_bytes(
                synthetic_sidecar()
            )
            clippings_dir = documents / "My Clippings.sdr"
            clippings_dir.mkdir(parents=True)
            (clippings_dir / "My Clippings.mbs").write_bytes(
                synthetic_sidecar()
            )
            output = Path(temp_dir) / "reading.csv"

            kindle_reader.main(
                [
                    "--path",
                    str(root),
                    "--output",
                    str(output),
                ]
            )
            with output.open(
                "r",
                encoding="utf-8-sig",
                newline="",
            ) as handle:
                rows = list(csv.DictReader(handle))

        self.assertEqual([row["title"] for row in rows], ["Manga Vol.04"])
        self.assertEqual(rows[0]["ebook_file"], str(source))

    def test_classifier_persists_manga_and_preserves_book_modes(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            csv_path = root / "reading.csv"
            fields = [
                "title",
                "asin",
                "reading_mode",
                "reading_mode_source",
                "metadata_tags",
                "word_count_source",
                "classification_source",
                "classification_confidence",
                "classification_reason",
                "series_key",
            ]
            with csv_path.open(
                "w",
                encoding="utf-8-sig",
                newline="",
            ) as handle:
                writer = csv.DictWriter(handle, fieldnames=fields)
                writer.writeheader()
                writer.writerow(
                    {
                        "title": "Manga Vol.04",
                        "asin": "B000000901",
                        "reading_mode": "book",
                        "reading_mode_source": "unclassified",
                        "metadata_tags": "Comics; Seinen",
                        "word_count_source": "kindle-estimate",
                    }
                )
            modes_path = root / "book_modes.json"
            modes_path.write_text(
                json.dumps(
                    {
                        "version": 7,
                        "custom": {"keep": True},
                        "by_title": {"Existing": "book"},
                    }
                ),
                encoding="utf-8",
            )
            args = argparse.Namespace(
                ocr_language="auto",
                ocr_engine="auto",
                ocr_dpi=200,
                ocr_min_confidence=0.7,
                ocr_workers=1,
                ocr_sample_pages=6,
            )

            counts = kindle_report.apply_manga_classification_to_csv(
                csv_path,
                root / "_ocr",
                args,
                modes_path,
            )
            with csv_path.open(
                "r",
                encoding="utf-8-sig",
                newline="",
            ) as handle:
                row = next(csv.DictReader(handle))
            modes = json.loads(modes_path.read_text(encoding="utf-8"))

        self.assertEqual(counts["manga"], 1)
        self.assertEqual(row["reading_mode"], "manga")
        self.assertEqual(row["classification_source"], "calibre-tags")
        self.assertEqual(row["word_count_source"], "not-applicable")
        self.assertEqual(modes["version"], 7)
        self.assertEqual(modes["custom"], {"keep": True})
        self.assertEqual(modes["by_title"]["Manga Vol.04"], "manga")
        self.assertEqual(modes["by_asin"]["B000000901"], "manga")

    def test_manga_cards_use_pages_and_do_not_show_zero_words(self) -> None:
        row = {
            "title": "Manga Pages",
            "asin": "B000000902",
            "reading_mode": "manga",
            "total_time_ms": "600000",
            "total_words": "0",
            "words_read": "0",
            "book_words": "0",
            "book_words_exact": "0",
            "wpm": "0",
            "progress_percent": "",
            "manga_pages_total": "185",
            "manga_pages_source": "pdfinfo-source-file",
        }
        document = kindle_report.build_html(
            [row],
            Path("reading.csv"),
            "Kindle",
        )

        self.assertIn('data-manga-pages-total="185"', document)
        self.assertIn('data-manga-pages-read=""', document)
        self.assertIn("未开始", document)
        self.assertNotIn(
            '<strong title="累计已读词数">0</strong>',
            document,
        )

    def test_history_keeps_manga_time_after_fmcache_rotation(self) -> None:
        row = {
            "title": "Manga History",
            "asin": "B000000903",
            "reading_mode": "manga",
            "total_time_ms": "0",
            "progress_percent": "",
            "manga_pages_total": "100",
            "manga_pages_read": "",
        }
        history = {
            "version": kindle_report.HISTORY_VERSION,
            "snapshots": [
                kindle_report.build_snapshot(
                    [
                        {
                            **row,
                            "total_time_ms": "600000",
                            "progress_percent": "50",
                            "manga_pages_read": "50",
                        }
                    ],
                    "2026-10-01T12:00:00+08:00",
                    "Kindle",
                )
            ],
        }

        merged, applied = kindle_report.merge_history_book_fallbacks(
            [row],
            history,
        )

        self.assertTrue(applied)
        self.assertEqual(merged[0]["total_time_ms"], "600000")
        self.assertEqual(merged[0]["progress_percent"], "50")
        self.assertEqual(merged[0]["manga_pages_read"], "50")
        self.assertEqual(merged[0]["manga_progress_source"], "history")

    def test_positive_legacy_manga_history_is_reused_without_cross_volume_leak(
        self,
    ) -> None:
        history = {
            "version": kindle_report.HISTORY_VERSION,
            "snapshots": [
                {
                    "captured_at": "2026-10-02T15:29:45+08:00",
                    "data_source_version": "1.10.0",
                    "books": {
                        "asin:A0000000000000000000000000000005": {
                            "book_id": (
                                "asin:A0000000000000000000000000000005"
                            ),
                            "asin": "A0000000000000000000000000000005",
                            "title": "合成漫画测试卷 Vol.04",
                            "reading_mode": "manga",
                            "total_time_ms": 3583825,
                            "progress_percent": 100.0,
                        }
                    },
                }
            ],
        }
        rows = [
            {
                "title": "合成漫画测试卷 Vol.04",
                "asin": "A0000000000000000000000000000005",
                "reading_mode": "manga",
                "total_time_ms": "0",
                "progress_percent": "",
            },
            {
                "title": (
                    "合成漫画测试卷 Vol.05_"
                    "F000000000000000000000000000000A"
                ),
                "reading_mode": "manga",
                "total_time_ms": "0",
                "progress_percent": "",
            },
        ]

        merged, applied = kindle_report.merge_history_book_fallbacks(
            rows,
            history,
        )

        self.assertTrue(applied)
        self.assertEqual(merged[0]["total_time_ms"], "3583825")
        self.assertEqual(merged[0]["progress_percent"], "100")
        self.assertEqual(merged[1]["total_time_ms"], "0")
        self.assertFalse(merged[1]["progress_percent"])

    def test_new_manga_volume_track_backfills_existing_profile(self) -> None:
        rows = [
            {
                "title": f"猎人 Vol.{index:02d}",
                "asin": f"B0000010{index:02d}",
                "reading_mode": "manga",
                "progress_percent": "100",
                "manga_pages_total": "100",
                "manga_pages_read": "100",
                "total_time_ms": "0",
            }
            for index in range(1, 4)
        ]
        existing = {
            "version": 2,
            "created_at": "2026-09-01T10:00:00+08:00",
            "baseline": {
                "captured_at": "2026-09-01T10:00:00+08:00",
                "total_time_ms": 0,
                "total_words": 0,
                "completed_book_count": 0,
                "read_book_count": 0,
                "book_count": 3,
            },
            "tiers": {},
            "levels": {
                "streak": 0,
                "time": 0,
                "books": 0,
                "words": 0,
                "speed": 0,
            },
            "observed_values": {
                "streak": 0,
                "time": 0,
                "books": 0,
                "words": 0,
                "speed": 0,
            },
        }
        profile = kindle_report.build_achievement_profile(
            rows,
            captured_at="2026-10-04T12:00:00+08:00",
            existing_profile=existing,
        )

        self.assertEqual(profile["observed_values"]["manga_volumes"], 3.0)
        self.assertEqual(profile["levels"]["manga_volumes"], 2)

    def test_sampled_classifier_returns_manga_and_uncertain_is_explicit(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "Manga Japanese Comic Vol.01.pdf"
            source.write_bytes(b"%PDF-1.4")
            page_paths = []
            for index in range(1, 7):
                page = root / f"page-{index}.png"
                page.write_bytes(f"page-{index}".encode("ascii"))
                page_paths.append(
                    kindle_manga_ocr.RenderedPage(
                        index,
                        page,
                        f"hash-{index}",
                    )
                )
            rendered = kindle_manga_ocr.RenderedBook(
                source_path=source,
                source_hash="source-hash",
                pdf_path=source,
                dpi=200,
                pages=tuple(page_paths),
                expected_pages=60,
            )

            class MangaEngine:
                def recognize(self, image_path, language):
                    del image_path, language
                    return kindle_manga_ocr.OcrPageResult(
                        "短对白",
                        0.9,
                        text_area_ratio=0.03,
                        chars_per_page=3,
                    )

            config = kindle_manga_ocr.OcrConfig(
                language="ja",
                engine="manga_ocr",
                dpi=200,
                cache_dir=root / "_ocr",
                mode="classify",
            )
            with mock.patch.object(
                kindle_manga_ocr,
                "engine_available",
                return_value=(True, ""),
            ):
                classified = kindle_manga_ocr.classify_manga_candidate(
                    {
                        "title": "Manga Japanese Comic Vol.01",
                        "source_file": str(source),
                    },
                    config,
                    source_resolver=lambda row, base: source,
                    sampler=lambda *args, **kwargs: rendered,
                    engine_factory=lambda name, language: MangaEngine(),
                )

        self.assertEqual(classified["reading_mode"], "manga")
        self.assertEqual(classified["classification_source"], "ocr-sample")
        self.assertGreater(classified["classification_confidence"], 0.6)

    def test_pdfinfo_page_count_is_cached_with_source_note(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "Manga.pdf"
            source.write_bytes(b"%PDF-1.4")
            cache_dir = root / "_ocr"
            with mock.patch.object(
                kindle_manga_ocr,
                "_pdf_page_count",
                return_value=185,
            ) as page_count:
                first = kindle_manga_ocr.resolve_manga_page_count(
                    source,
                    cache_dir,
                    source_hash="page-count-test",
                )
                second = kindle_manga_ocr.resolve_manga_page_count(
                    source,
                    cache_dir,
                    source_hash="page-count-test",
                )

        self.assertEqual(first[:2], (185, "pdfinfo-source-file"))
        self.assertEqual(second, first)
        self.assertEqual(page_count.call_count, 1)

    def test_kfx_position_map_derives_manga_progress_and_pages(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "Manga Vol.06.kfx"
            source.write_bytes(b"kfx")
            csv_path = root / "reading.csv"
            fields = [
                "title",
                "reading_mode",
                "ebook_file",
                "source_file",
                "progress_percent",
                "manga_pages_total",
                "manga_pages_read",
                "manga_progress_source",
                "manga_position",
                "manga_max_position",
            ]
            with csv_path.open(
                "w",
                encoding="utf-8-sig",
                newline="",
            ) as handle:
                writer = csv.DictWriter(handle, fieldnames=fields)
                writer.writeheader()
                writer.writerow(
                    {
                        "title": "Manga Vol.06",
                        "reading_mode": "manga",
                        "ebook_file": str(source),
                        "source_file": str(root / "Manga Vol.06.yjf"),
                        "manga_pages_total": "201",
                    }
                )
            with mock.patch.object(
                kindle_manga_ocr,
                "resolve_kfx_position_info",
                return_value={"count": 502, "last_pid": 501},
            ), mock.patch.object(
                kindle_reader,
                "decode_sidecar",
                return_value={
                    "fpr": {"position": "AcYAAAAAAAAA:491"},
                },
            ):
                counts = kindle_report.apply_manga_page_counts_to_csv(
                    csv_path,
                    root / "_ocr",
                )
            with csv_path.open(
                "r",
                encoding="utf-8-sig",
                newline="",
            ) as handle:
                row = next(csv.DictReader(handle))

        self.assertEqual(counts["manga_rows"], 1)
        self.assertEqual(row["progress_percent"], "98")
        self.assertEqual(row["manga_pages_read"], "197")
        self.assertEqual(row["manga_progress_source"], "kfx-position-map")
        self.assertEqual(row["manga_position"], "491")
        self.assertEqual(row["manga_max_position"], "501")

    def test_calibre_metadata_overrides_title_author_and_asin(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "Kindle"
            sidecar_dir = (
                root
                / "documents"
                / "Downloads"
                / "Old Folder Name_A0000000000000000000000000000001.sdr"
            )
            sidecar_dir.mkdir(parents=True)
            sidecar = sidecar_dir / "Old Folder Name.azw3f"
            sidecar.write_bytes(synthetic_sidecar())
            (root / "metadata.calibre").write_text(
                json.dumps(
                    [
                        {
                            "title": "Metadata Title",
                            "authors": ["Metadata Author"],
                            "lpath": (
                                "documents/Downloads/"
                                "Old Folder Name_"
                                "A0000000000000000000000000000001.kfx"
                            ),
                            "identifiers": {
                                "mobi-asin": (
                                    "A0000000000000000000000000000001"
                                )
                            },
                            "series": "Metadata Series",
                            "series_index": 2.5,
                            "publisher": "Metadata Publisher",
                            "pubdate": "2024-03-01T00:00:00+00:00",
                        }
                    ]
                ),
                encoding="utf-8",
            )
            output = Path(temp_dir) / "reading.csv"

            exit_code = kindle_reader.main(
                [
                    "--path",
                    str(root),
                    "--output",
                    str(output),
                ]
            )
            self.assertEqual(exit_code, 0)

            with output.open("r", encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.DictReader(handle))

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["title"], "Metadata Title")
        self.assertEqual(rows[0]["author"], "Metadata Author")
        self.assertEqual(
            rows[0]["asin"], "A0000000000000000000000000000001"
        )
        self.assertEqual(rows[0]["series"], "Metadata Series")
        self.assertEqual(rows[0]["series_index"], "2.5")
        self.assertEqual(rows[0]["publisher"], "Metadata Publisher")
        self.assertEqual(
            rows[0]["published_date"], "2024-03-01T00:00:00+00:00"
        )
        self.assertEqual(rows[0]["metadata_match"], "asin")

    def test_calibre_metadata_can_match_by_asin(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "Kindle"
            sidecar_dir = (
                root
                / "documents"
                / "Downloads"
                / "Truncated Name_A0000000000000000000000000000004.sdr"
            )
            sidecar_dir.mkdir(parents=True)
            (sidecar_dir / "Truncated Name.azw3f").write_bytes(
                synthetic_sidecar()
            )
            (root / "metadata.calibre").write_text(
                json.dumps(
                    [
                        {
                            "title": "Full Metadata Title",
                            "authors": ["Full Metadata Author"],
                            "lpath": (
                                "documents/Downloads/"
                                "Full Metadata Title_"
                                "A0000000000000000000000000000004.kfx"
                            ),
                            "identifiers": {
                                "mobi-asin": (
                                    "A0000000000000000000000000000004"
                                )
                            },
                        }
                    ]
                ),
                encoding="utf-8",
            )
            output = Path(temp_dir) / "reading.csv"

            exit_code = kindle_reader.main(
                [
                    "--path",
                    str(root),
                    "--output",
                    str(output),
                ]
            )
            self.assertEqual(exit_code, 0)

            with output.open("r", encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.DictReader(handle))

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["title"], "Full Metadata Title")
        self.assertEqual(rows[0]["author"], "Full Metadata Author")
        self.assertEqual(
            rows[0]["asin"], "A0000000000000000000000000000004"
        )
        self.assertEqual(rows[0]["metadata_match"], "asin")

    def test_calibre_metadata_match_prefers_asin(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "Kindle"
            documents = root / "documents"
            target_asin = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
            other_asin = "BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB"
            sidecar_dir = (
                documents
                / "Downloads"
                / f"Book Title_{target_asin}.sdr"
            )
            group = {
                "directory": str(sidecar_dir),
                "stem": f"Book Title_{target_asin}",
                "title": f"Book Title_{target_asin}",
                "source_file": str(
                    sidecar_dir / f"Book Title_{target_asin}.azw3f"
                ),
            }
            path_entry = {
                "title": "Path Candidate",
                "lpath": (
                    "documents/Downloads/"
                    f"Book Title_{target_asin}.kfx"
                ),
                "identifiers": {"mobi-asin": other_asin},
            }
            asin_entry = {
                "title": "ASIN Candidate",
                "lpath": "documents/Downloads/Elsewhere/Different.kfx",
                "identifiers": {"mobi-asin": target_asin},
            }
            title_entry = {
                "title": f"Book Title_{target_asin}",
                "lpath": "documents/Downloads/Elsewhere/Title.kfx",
            }

            selected, match = kindle_reader._match_calibre_metadata(
                group,
                [path_entry, asin_entry, title_entry],
                book={"asin": target_asin},
                documents=str(documents),
            )

        self.assertIs(selected, asin_entry)
        self.assertEqual(match, "asin")

    def test_calibre_metadata_match_prefers_full_path_over_title(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "Kindle"
            documents = root / "documents"
            sidecar_dir = (
                documents / "Downloads" / "Shared Title.sdr"
            )
            group = {
                "directory": str(sidecar_dir),
                "stem": "Shared Title",
                "title": "Shared Title",
                "source_file": str(
                    sidecar_dir / "Shared Title.azw3f"
                ),
            }
            path_entry = {
                "title": "Path Winner",
                "lpath": "documents/Downloads/Shared Title.kfx",
            }
            title_entry = {
                "title": "Shared Title",
                "lpath": "documents/Downloads/Other Folder/Shared Title.kfx",
            }

            selected, match = kindle_reader._match_calibre_metadata(
                group,
                [title_entry, path_entry],
                book={},
                documents=str(documents),
            )

        self.assertIs(selected, path_entry)
        self.assertEqual(match, "path")

    def test_calibre_metadata_match_falls_back_to_normalized_title(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "Kindle"
            documents = root / "documents"
            sidecar_dir = (
                documents / "Downloads" / "Ｃａｆｅ Book.sdr"
            )
            group = {
                "directory": str(sidecar_dir),
                "stem": "Ｃａｆｅ Book",
                "title": "Ｃａｆｅ Book",
                "source_file": str(
                    sidecar_dir / "Ｃａｆｅ Book.azw3f"
                ),
            }
            title_entry = {
                "title": "Cafe Book",
                "lpath": "documents/Downloads/Other Folder/Other.kfx",
            }

            selected, match = kindle_reader._match_calibre_metadata(
                group,
                [title_entry],
                book={},
                documents=str(documents),
            )

        self.assertIs(selected, title_entry)
        self.assertEqual(match, "title")

    def test_bad_mbp_is_reported_without_stopping(self) -> None:
        logger = kindle_reader._make_logger(False)
        with tempfile.TemporaryDirectory() as temp_dir:
            sidecar_dir = Path(temp_dir) / "Old Book.sdr"
            sidecar_dir.mkdir()
            (sidecar_dir / "Old Book.mbp").write_bytes(b"not a KRDS file")
            books, errors = kindle_reader.extract_books(
                temp_dir, logger, include_empty=False
            )

        self.assertEqual(books, [])
        self.assertEqual(len(errors), 1)
        self.assertIn("signature is incorrect", errors[0])

    def test_single_book_failure_does_not_stop_scan(self) -> None:
        logger = kindle_reader._make_logger(False)
        with tempfile.TemporaryDirectory() as temp_dir:
            documents = Path(temp_dir) / "documents"
            documents.mkdir()
            for title in ("Bad Book", "Good Book"):
                sidecar_dir = documents / (title + ".sdr")
                sidecar_dir.mkdir()
                (sidecar_dir / (title + ".yjf")).write_bytes(
                    synthetic_sidecar()
                )

            original_match = kindle_reader._match_calibre_metadata

            def match_or_fail(group, metadata, book=None, documents=None):
                if group["title"] == "Bad Book":
                    raise ValueError("synthetic metadata failure")
                return original_match(
                    group,
                    metadata,
                    book=book,
                    documents=documents,
                )

            with mock.patch.object(
                kindle_reader,
                "_match_calibre_metadata",
                side_effect=match_or_fail,
            ):
                books, errors = kindle_reader.extract_books(
                    str(documents),
                    logger,
                    include_empty=False,
                )

        self.assertEqual([book["title"] for book in books], ["Good Book"])
        self.assertEqual(len(errors), 1)
        self.assertIn("synthetic metadata failure", errors[0])
        self.assertIn("skipped only this book", errors[0])
        self.assertIn("Recovery:", errors[0])

    def test_transient_calibre_oserror_is_retried_once(self) -> None:
        logger = kindle_reader._make_logger(False)
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            ebook = root / "Book.kfx"
            ebook.write_bytes(b"book")
            text_cache = root / "wordcount"
            calls = []

            def fake_run(command, **kwargs):
                del kwargs
                calls.append(command)
                if len(calls) == 1:
                    raise OSError("temporary file lock")
                Path(command[2]).write_text("Hello world", encoding="utf-8")
                return subprocess.CompletedProcess(command, 0, "")

            with mock.patch.object(
                kindle_reader.subprocess,
                "run",
                side_effect=fake_run,
            ), mock.patch.object(kindle_reader.time, "sleep"):
                text = kindle_reader._convert_ebook_to_text(
                    str(ebook),
                    "calibre.exe",
                    str(text_cache),
                    logger,
                )

        self.assertEqual(text, "Hello world")
        self.assertEqual(len(calls), 2)

    def test_failed_calibre_conversion_removes_partial_output(self) -> None:
        logger = kindle_reader._make_logger(False)
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            ebook = root / "Book.kfx"
            ebook.write_bytes(b"book")
            text_cache = root / "wordcount"

            def fake_run(command, **kwargs):
                del kwargs
                Path(command[2]).write_text("partial", encoding="utf-8")
                return subprocess.CompletedProcess(command, 1, "conversion error")

            with mock.patch.object(
                kindle_reader.subprocess,
                "run",
                side_effect=fake_run,
            ), mock.patch.object(kindle_reader.time, "sleep"):
                with self.assertRaisesRegex(
                    RuntimeError,
                    "Recovery:",
                ):
                    kindle_reader._convert_ebook_to_text(
                        str(ebook),
                        "calibre.exe",
                        str(text_cache),
                        logger,
                    )

            self.assertEqual(list(text_cache.glob("*.txt")), [])

    @unittest.skipUnless(os.name == "nt", "Windows long-path regression test")
    def test_windows_long_path_is_readable(self) -> None:
        logger = kindle_reader._make_logger(False)
        temp_dir = tempfile.mkdtemp()
        try:
            deep_dir = Path(temp_dir)
            for _ in range(14):
                deep_dir = deep_dir / ("segment-" + "x" * 18)
                os.mkdir(kindle_reader._filesystem_path(str(deep_dir)))

            sidecar_dir = deep_dir / "Long Book.sdr"
            os.mkdir(kindle_reader._filesystem_path(str(sidecar_dir)))
            sidecar = sidecar_dir / "Long Book.yjf"
            self.assertGreater(len(str(sidecar)), 260)
            with open(
                kindle_reader._filesystem_path(str(sidecar)), "wb"
            ) as handle:
                handle.write(synthetic_sidecar())

            groups = kindle_reader.find_sidecar_groups(temp_dir)
            decoded = kindle_reader.decode_sidecar(str(sidecar), logger)
        finally:
            shutil.rmtree(
                kindle_reader._filesystem_path(temp_dir),
                ignore_errors=True,
            )

        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0]["source_file"], str(sidecar))
        self.assertNotIn("_error", decoded)
        self.assertEqual(
            kindle_reader._nested_timer_model(decoded)["totalTime"],
            600000,
        )

    @unittest.skipUnless(
        CHROMIUM_BROWSER, "Chromium browser is not available"
    )
    def test_period_bar_chart_visual_markers_and_viewports(self) -> None:
        assert CHROMIUM_BROWSER is not None

        def snapshot(
            total_time_ms: int,
            total_words: int,
            captured_at: str,
        ) -> dict:
            return kindle_report.build_snapshot(
                [
                    {
                        "title": "Bar Chart Book",
                        "author": "Author",
                        "asin": "B000000031",
                        "total_time_ms": str(total_time_ms),
                        "total_words": str(total_words),
                        "wpm": "100",
                        "progress_percent": "10",
                    }
                ],
                captured_at,
                "Kindle",
            )

        history = {
            "version": kindle_report.HISTORY_VERSION,
            "snapshots": [
                snapshot(0, 0, "2026-09-28T00:00:00+08:00"),
                snapshot(1200000, 6000, "2026-09-29T00:00:00+08:00"),
                snapshot(2400000, 12000, "2026-09-30T00:00:00+08:00"),
            ],
        }
        current = snapshot(3600000, 18000, "2026-10-01T12:00:00+08:00")
        period_data = kindle_report.build_period_reports(
            current,
            history,
            None,
            date(2026, 10, 1),
        )
        section = kindle_report.build_period_report_section(period_data)

        for marker in (
            'data-period-trend-chart="bars"',
            'data-chart-type="bar"',
            'data-period-grid="lines"',
            'data-period-baseline="true"',
            'data-period-bar-style="uniform"',
            'class="period-grid-lines"',
            'class="period-bar"',
            'class="period-bar-fill"',
            'class="period-bar-empty"',
            'data-period-state="active"',
            'data-period-state="no-record"',
            "data-period-value-ms=",
        ):
            self.assertIn(marker, section)
        self.assertIn("aria-label=", section)

        empty_report = kindle_report.build_period_reports(
            current,
            {"version": kindle_report.HISTORY_VERSION, "snapshots": []},
            None,
            date(2026, 10, 1),
        )["reports"][0]
        empty_section = kindle_report.build_period_report_section(
            {"reports": [empty_report]}
        )
        self.assertIn("继续阅读即可解锁", empty_section)
        self.assertIn("Keep reading to unlock", empty_section)
        self.assertNotIn('data-period-trend-chart="bars"', empty_section)
        self.assertNotIn("period-bar-fill", empty_section)

        with CURRENT_CSV_PATH.open(
            "r", encoding="utf-8-sig", newline=""
        ) as handle:
            rows = [dict(row) for row in csv.DictReader(handle)]
        document = kindle_report.build_html(
            rows,
            Path("reading.csv"),
            "Kindle",
            trends=kindle_report.build_trends(current, history),
            period_reports=period_data,
            generated_at="2026-10-01T12:00:00+08:00",
            xlsx_path=Path("reading.xlsx"),
        )
        self.assertIn(".period-grid-lines {", document)
        self.assertIn(
            "border-bottom: 1px solid var(--period-baseline-line)",
            document,
        )
        self.assertIn(
            "width: min(var(--period-bar-width), 100%)", document
        )
        self.assertIn("print-color-adjust: exact", document)
        self.assertIn("--period-plot-height: 250px", document)
        self.assertIn("--period-plot-height: 190px", document)
        self.assertIn("--period-plot-height: 140px", document)
        self.assertIn("data-period-color=", section)
        instrumentation = """
<script>
  window.addEventListener("load", () => {
    const root = document.documentElement;
    const charts = Array.from(
      document.querySelectorAll('[data-period-trend-chart="bars"]')
    );
    root.dataset.testChartCount = String(charts.length);
    const activeView = periodViews.find(view => !view.hidden);
    const chart = activeView
      ? activeView.querySelector('[data-period-trend-chart="bars"]')
      : null;
    root.dataset.testActiveChart = chart ? "true" : "false";
    if (!chart) { return; }
    const plot = chart.querySelector(".period-bars");
    const bars = Array.from(chart.querySelectorAll(".period-bar"));
    root.dataset.testBarCount = String(bars.length);
    root.dataset.testDeclaredBuckets = chart.dataset.periodBucketCount || "";
    const widths = bars
      .map(bar => bar.querySelector(".period-bar-track"))
      .filter(Boolean)
      .map(track => Math.round(track.getBoundingClientRect().width));
    root.dataset.testTrackWidthMin = String(
      widths.length ? Math.min(...widths) : 0
    );
    root.dataset.testTrackWidthSpread = String(
      widths.length ? Math.max(...widths) - Math.min(...widths) : 0
    );
    const fills = bars
      .map(bar => bar.querySelector(".period-bar-fill"))
      .filter(Boolean);
    root.dataset.testTallestFill = String(
      Math.round(Math.max(
        0,
        ...fills.map(fill => fill.getBoundingClientRect().height)
      ))
    );
    const grid = chart.querySelector(".period-grid-lines");
    root.dataset.testGridHeight = grid
      ? String(Math.round(grid.getBoundingClientRect().height))
      : "0";
    root.dataset.testPlotHeight = plot
      ? getComputedStyle(plot)
          .getPropertyValue("--period-plot-height").trim()
      : "";
    root.dataset.testBarColors = fills
      .map(fill => {
        const style = getComputedStyle(fill);
        return style.backgroundImage || style.backgroundColor;
      })
      .filter(Boolean)
      .join("|");
    root.dataset.testValueLabels = String(
      bars.filter(bar => {
        const value = bar.querySelector(".period-bar-value");
        return !!value && (value.textContent || "").trim().length > 0;
      }).length
    );
    root.dataset.testBarTitles = String(
      bars.filter(bar => (bar.getAttribute("title") || "").trim()).length
    );
    root.dataset.testNoRecordBars = String(
      document.querySelectorAll(
        '.period-bar[data-period-state="no-record"]'
      ).length
    );
    root.dataset.testActiveBars = String(
      document.querySelectorAll(
        '.period-bar[data-period-state="active"]'
      ).length
    );
    const plotRect = plot ? plot.getBoundingClientRect() : null;
    root.dataset.testBarsOverflowPlot = String(
      bars.some(bar => {
        const rect = bar.getBoundingClientRect();
        if (!plotRect || rect.width === 0) { return false; }
        return rect.left < plotRect.left - 1 ||
          rect.right > plotRect.right + 1;
      })
    );
    root.dataset.testValueTruncated = String(
      Array.from(chart.querySelectorAll(".period-bar-value")).some(
        value => value.scrollWidth > value.clientWidth + 1
      )
    );
  });
</script>
"""
        self.assertIn("</body>", document)
        instrumented = document.replace(
            "</body>", instrumentation + "</body>"
        )
        snapshots = browser_viewport_snapshots(instrumented)

        for viewport, snapshot_document in snapshots.items():
            attributes = html_root_attributes(str(snapshot_document))
            with self.subTest(viewport=viewport):
                self.assertEqual(
                    attributes.get("data-test-chart-count"), "3"
                )
                self.assertEqual(
                    attributes.get("data-test-active-chart"), "true"
                )
                self.assertGreaterEqual(
                    int(attributes.get("data-test-bar-count") or "0"), 4
                )
                self.assertEqual(
                    attributes.get("data-test-bar-count"),
                    attributes.get("data-test-declared-buckets"),
                )
                self.assertGreaterEqual(
                    int(attributes.get("data-test-track-width-min") or "0"),
                    20,
                )
                self.assertLessEqual(
                    int(
                        attributes.get("data-test-track-width-spread")
                        or "99"
                    ),
                    1,
                )
                self.assertGreater(
                    int(attributes.get("data-test-tallest-fill") or "0"), 0
                )
                self.assertGreater(
                    int(attributes.get("data-test-grid-height") or "0"), 40
                )
                plot_height = float(
                    (
                        attributes.get("data-test-plot-height") or "0"
                    ).replace("px", "")
                )
                if viewport[0] <= 720:
                    self.assertGreaterEqual(plot_height, 180)
                    self.assertLessEqual(plot_height, 200)
                else:
                    self.assertGreaterEqual(plot_height, 240)
                    self.assertLessEqual(plot_height, 280)
                colors = [
                    color
                    for color in (
                        attributes.get("data-test-bar-colors") or ""
                    ).split("|")
                    if color
                ]
                self.assertGreaterEqual(
                    len(set(colors)), min(3, len(colors))
                )
                self.assertEqual(
                    int(attributes.get("data-test-value-labels") or "0"),
                    int(attributes.get("data-test-bar-count") or "0"),
                )
                self.assertEqual(
                    int(attributes.get("data-test-bar-titles") or "0"),
                    int(attributes.get("data-test-bar-count") or "0"),
                )
                self.assertGreaterEqual(
                    int(attributes.get("data-test-active-bars") or "0"), 1
                )
                self.assertGreaterEqual(
                    int(attributes.get("data-test-no-record-bars") or "0"), 1
                )
                self.assertEqual(
                    attributes.get("data-test-bars-overflow-plot"), "false"
                )
                self.assertEqual(
                    attributes.get("data-test-value-truncated"), "false"
                )
                self.assertEqual(
                    attributes.get("data-period-trend-type"), "bars"
                )
                self.assertGreaterEqual(
                    int(attributes.get("data-period-trend-buckets") or "0"),
                    4,
                )
                self.assertEqual(
                    attributes.get("data-overflow-x"), "false"
                )
                self.assertEqual(
                    attributes.get("data-overflow-elements"), ""
                )
                self.assertEqual(
                    attributes.get("data-overlap-count"), "0",
                    attributes.get("data-overlap-elements"),
                )
                self.assertEqual(
                    attributes.get("data-truncated-numbers"), ""
                )
                self.assertGreater(
                    int(attributes.get("data-number-checks") or "0"), 0
                )

    @unittest.skipUnless(
        CHROMIUM_BROWSER, "Chromium browser is not available"
    )
    def test_late_night_greeting_is_short_and_emoticon_stays_intact(
        self,
    ) -> None:
        assert CHROMIUM_BROWSER is not None
        with CURRENT_CSV_PATH.open(
            "r", encoding="utf-8-sig", newline=""
        ) as handle:
            rows = [dict(row) for row in csv.DictReader(handle)]
        document = kindle_report.build_html(
            rows,
            Path("reading.csv"),
            "Kindle",
            generated_at="2026-10-01T12:00:00+08:00",
            xlsx_path=Path("reading.xlsx"),
        )

        late_night_en = "Late night! One quiet page before sleep! (｡･ω･｡)"
        previous_en = (
            "Late-night greetings! Quiet hours are a good time to read "
            "a page! (｡･ω･｡)"
        )
        self.assertIn(late_night_en, document)
        self.assertNotIn("Late-night greetings!", document)
        self.assertLess(len(late_night_en), len(previous_en))
        self.assertLessEqual(len(late_night_en), 50)
        for tone in ("late-night", "morning", "day", "dusk", "night"):
            self.assertIn(f'data-greeting-tone="{tone}"', document)
        self.assertIn("transition: color .7s ease", document)
        self.assertRegex(
            document, r"\.greeting-message\s*\{[^}]*font-size:\s*16px"
        )
        self.assertRegex(
            document,
            r"\.greeting-emoticon\s*\{[^}]*display:\s*inline-block",
        )
        self.assertRegex(
            document,
            r"\.greeting-emoticon\s*\{[^}]*white-space:\s*nowrap",
        )

        instrumentation = """
<script>
  const NativeDate = Date;
  class FixedDate extends NativeDate {
    constructor(...args) {
      if (args.length) {
        super(...args);
      } else {
        super(2026, 9, 1, 3, 0, 0);
      }
    }
    static now() {
      return new NativeDate(2026, 9, 1, 3, 0, 0).getTime();
    }
  }
  window.Date = FixedDate;
  window.addEventListener("load", () => {
    const root = document.documentElement;
    setUiLanguage("en", false);
    root.dataset.testGreetingTone = root.dataset.greetingTone || "";
    root.dataset.testGreetingPeriod = root.dataset.greetingPeriod || "";
    const en = document.getElementById("greeting-en");
    const text = (en.textContent || "").trim();
    root.dataset.testGreetingText = text;
    root.dataset.testGreetingLength = String(text.length);
    const emoticon = en.querySelector(".greeting-emoticon");
    root.dataset.testEmoticonPresent = String(!!emoticon);
    if (emoticon) {
      const style = getComputedStyle(emoticon);
      root.dataset.testEmoticonDisplay = style.display;
      root.dataset.testEmoticonWhiteSpace = style.whiteSpace;
      root.dataset.testEmoticonRects = String(
        emoticon.getClientRects().length
      );
      const rect = emoticon.getBoundingClientRect();
      root.dataset.testEmoticonSingleLine = String(
        rect.height <= parseFloat(style.fontSize) * 1.6
      );
      const emoticonText = text.slice(text.lastIndexOf("("));
      root.dataset.testEmoticonBound = String(
        emoticon.textContent.length > emoticonText.length
      );
    }
    const tonePairs = [
      [2, "late-night"], [7, "morning"], [12, "day"],
      [18, "dusk"], [22, "night"],
    ];
    root.dataset.testToneMap = String(tonePairs.every(
      pair => browserGreetingTone(
        new NativeDate(2026, 0, 1, pair[0], 0, 0)
      ) === pair[1]
    ));
    root.dataset.testModeTones = String(
      GREETING_MODE_TONES["late-night"] === "late-night" &&
      GREETING_MODE_TONES.morning === "morning" &&
      GREETING_MODE_TONES.noon === "day" &&
      GREETING_MODE_TONES.evening === "night"
    );
    root.dataset.testGreetingTransition =
      getComputedStyle(
        document.getElementById("greeting-message")
      ).transitionProperty || "";
  });
</script>
"""
        self.assertIn("</body>", document)
        instrumented = document.replace(
            "</body>", instrumentation + "</body>"
        )
        snapshots = browser_viewport_snapshots(
            instrumented, ((320, 568), (1440, 900))
        )

        for viewport, snapshot_document in snapshots.items():
            attributes = html_root_attributes(str(snapshot_document))
            with self.subTest(viewport=viewport):
                self.assertEqual(
                    attributes.get("data-test-greeting-tone"), "late-night"
                )
                self.assertEqual(
                    attributes.get("data-test-greeting-period"),
                    "late-night",
                )
                greeting = attributes.get("data-test-greeting-text") or ""
                self.assertTrue(
                    greeting.replace("\u00a0", " ").startswith("Late night!")
                )
                self.assertIn("(｡･ω･｡)", greeting)
                self.assertLessEqual(
                    int(attributes.get("data-test-greeting-length") or "0"),
                    50,
                )
                self.assertEqual(
                    attributes.get("data-test-emoticon-present"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-emoticon-display"),
                    "inline-block",
                )
                self.assertEqual(
                    attributes.get("data-test-emoticon-white-space"),
                    "nowrap",
                )
                self.assertEqual(
                    attributes.get("data-test-emoticon-rects"), "1"
                )
                self.assertEqual(
                    attributes.get("data-test-emoticon-single-line"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-emoticon-bound"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-tone-map"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-mode-tones"), "true"
                )
                self.assertIn(
                    "color",
                    attributes.get("data-test-greeting-transition") or "",
                )
                self.assertEqual(
                    attributes.get("data-overflow-x"), "false"
                )

    @unittest.skipUnless(
        CHROMIUM_BROWSER, "Chromium browser is not available"
    )
    def test_report_section_toggles_and_minimal_mode(self) -> None:
        assert CHROMIUM_BROWSER is not None
        with CURRENT_CSV_PATH.open(
            "r", encoding="utf-8-sig", newline=""
        ) as handle:
            rows = [dict(row) for row in csv.DictReader(handle)]
        document = kindle_report.build_html(
            rows,
            Path("reading.csv"),
            "Kindle",
            generated_at="2026-10-01T12:00:00+08:00",
            xlsx_path=Path("reading.xlsx"),
        )

        self.assertNotIn('class="settings-bar"', document)
        self.assertNotIn('class="settings-quick"', document)
        self.assertIn(
            '<details class="appearance-panel" id="appearance-panel">',
            document,
        )
        appearance_start = document.index('id="appearance-panel"')
        appearance_end = document.index("</details>", appearance_start)
        appearance_html = document[appearance_start:appearance_end]
        self.assertIn(".appearance-body", document)
        for key in ("period", "annual", "advanced", "trends", "heatmap"):
            self.assertIn(
                f'data-section-toggle="{key}"', appearance_html
            )
            self.assertRegex(
                document,
                r'id="section-' + key + r'"[^>]*checked',
            )
        self.assertIn('id="minimal-mode"', appearance_html)
        self.assertIn("kindle-dashboard-section-visibility", document)
        self.assertIn("kindle-dashboard-minimal-mode", document)
        self.assertNotIn("<details", appearance_html[
            appearance_html.index(">") + 1:
        ])

        instrumentation = """
<script>
  window.addEventListener("load", () => {
    const root = document.documentElement;
    const ids = [
      "section-period", "section-annual", "section-advanced",
      "section-trends", "section-heatmap",
    ];
    root.dataset.testTogglesInPanel = String(ids.every(id => {
      const input = document.getElementById(id);
      return !!input && appearancePanel.contains(input) &&
        !!input.closest(".appearance-body");
    }));
    root.dataset.testMinimalInPanel = String(
      !!minimalModeToggle &&
      appearancePanel.contains(minimalModeToggle) &&
      !!minimalModeToggle.closest(".appearance-body")
    );
    root.dataset.testDefaults = [
      root.dataset.sectionPeriod,
      root.dataset.sectionAnnual,
      root.dataset.sectionAdvanced,
      root.dataset.sectionTrends,
      root.dataset.sectionHeatmap,
    ].join(",");
    root.dataset.testDefaultChecked = String(
      sectionToggles.length === 5 &&
      sectionToggles.every(toggle => toggle.checked)
    );
    root.dataset.testMinimalDefault = String(
      !minimalMode && !minimalModeToggle.checked &&
      root.dataset.minimalMode === "false"
    );
    root.dataset.testAllVisible = String(
      SECTION_KEYS.every(key => !reportSections[key].hidden)
    );

    const heatmapToggle = document.getElementById("section-heatmap");
    heatmapToggle.checked = false;
    heatmapToggle.dispatchEvent(new Event("change"));
    root.dataset.testHeatmapHidden = String(reportSections.heatmap.hidden);
    root.dataset.testHeatmapStored = String(
      JSON.parse(readPreference(SECTION_VISIBILITY_KEY)).heatmap
    );

    const minimal = document.getElementById("minimal-mode");
    minimal.checked = true;
    minimal.dispatchEvent(new Event("change"));
    root.dataset.testMinimalHidesAll = String(
      SECTION_KEYS.every(key => reportSections[key].hidden)
    );
    root.dataset.testMinimalStored =
      readPreference(MINIMAL_MODE_KEY) || "";
    root.dataset.testHeatmapToggleKept = String(!heatmapToggle.checked);

    // Changing individual switches during minimal mode is remembered but
    // does not reveal a section until minimal mode is switched off.
    const trendsToggle = document.getElementById("section-trends");
    trendsToggle.checked = false;
    trendsToggle.dispatchEvent(new Event("change"));
    root.dataset.testMinimalAdjustHidden =
      String(reportSections.trends.hidden);
    root.dataset.testMinimalAdjustStored = String(
      JSON.parse(readPreference(SECTION_VISIBILITY_KEY)).trends
    );

    // Simulate a reload: reset in-memory state and restore from storage.
    sectionVisibility = {...DEFAULT_SECTION_VISIBILITY};
    minimalMode = false;
    restoreSectionVisibility();
    root.dataset.testReloadMinimal = root.dataset.minimalMode;
    root.dataset.testReloadAllHidden = String(
      SECTION_KEYS.every(key => reportSections[key].hidden)
    );
    root.dataset.testReloadToggles = [
      heatmapToggle.checked, trendsToggle.checked,
    ].join(",");

    minimal.checked = false;
    minimal.dispatchEvent(new Event("change"));
    root.dataset.testRestoreAll = [
      root.dataset.sectionPeriod,
      root.dataset.sectionAnnual,
      root.dataset.sectionAdvanced,
      root.dataset.sectionTrends,
      root.dataset.sectionHeatmap,
    ].join(",");
    root.dataset.testRestoreMinimal =
      readPreference(MINIMAL_MODE_KEY) || "";
    updateLayoutMetrics();
  });
</script>
"""
        self.assertIn("</body>", document)
        instrumented = document.replace(
            "</body>", instrumentation + "</body>"
        )
        snapshots = browser_viewport_snapshots(
            instrumented, ((320, 568), (1440, 900))
        )

        for viewport, snapshot_document in snapshots.items():
            attributes = html_root_attributes(str(snapshot_document))
            with self.subTest(viewport=viewport):
                self.assertEqual(
                    attributes.get("data-test-toggles-in-panel"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-minimal-in-panel"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-defaults"),
                    "visible,visible,visible,visible,visible",
                )
                self.assertEqual(
                    attributes.get("data-test-default-checked"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-minimal-default"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-all-visible"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-heatmap-hidden"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-heatmap-stored"), "false"
                )
                self.assertEqual(
                    attributes.get("data-test-minimal-hides-all"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-minimal-stored"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-heatmap-toggle-kept"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-minimal-adjust-hidden"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-minimal-adjust-stored"),
                    "false",
                )
                self.assertEqual(
                    attributes.get("data-test-reload-minimal"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-reload-all-hidden"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-reload-toggles"),
                    "false,false",
                )
                self.assertEqual(
                    attributes.get("data-test-restore-all"),
                    "visible,visible,visible,hidden,hidden",
                )
                self.assertEqual(
                    attributes.get("data-test-restore-minimal"), "false"
                )
                self.assertEqual(
                    attributes.get("data-overflow-x"), "false"
                )
                self.assertEqual(
                    attributes.get("data-overlap-count"), "0",
                    attributes.get("data-overlap-elements"),
                )

    @unittest.skipUnless(
        CHROMIUM_BROWSER, "Chromium browser is not available"
    )
    def test_insight_cards_markers_bilingual_and_viewports(self) -> None:
        assert CHROMIUM_BROWSER is not None
        with CURRENT_CSV_PATH.open(
            "r", encoding="utf-8-sig", newline=""
        ) as handle:
            rows = [dict(row) for row in csv.DictReader(handle)]
        document = kindle_report.build_html(
            rows,
            Path("reading.csv"),
            "Kindle",
            generated_at="2026-10-01T12:00:00+08:00",
            xlsx_path=Path("reading.xlsx"),
        )

        self.assertIn(
            'data-insight-key="most-read" data-insight-layout="feature"',
            document,
        )
        for key in ("longest-book", "fastest-wpm"):
            self.assertIn(
                f'data-insight-key="{key}" data-insight-layout="cover"',
                document,
            )
        self.assertIn(
            'data-insight-key="data-quality" data-insight-layout="plain"',
            document,
        )
        self.assertIn('class="insight-icon"', document)
        self.assertIn('class="insight-value"', document)
        self.assertIn('class="insight-detail"', document)
        self.assertIn('class="insight-cover"', document)
        self.assertIn('class="insight-share"', document)
        self.assertIn("insight-share-track", document)
        self.assertIn("占累计阅读", document)
        self.assertIn("of total reading time", document)
        for label in (
            "阅读最多",
            "Most read",
            "最长书籍",
            "Longest book",
            "速度最高",
            "Highest WPM",
        ):
            self.assertIn(label, document)
        most_read_start = document.index('data-insight-key="most-read"')
        most_read_end = document.index("</article>", most_read_start)
        most_read_card = document[most_read_start:most_read_end]
        self.assertIn("insight-cover", most_read_card)
        self.assertIn("insight-label", most_read_card)
        self.assertIn("insight-share", most_read_card)
        insight_value_rule = re.search(
            r"\.insight-value\s*\{([^}]*)\}", document
        )
        self.assertIsNotNone(insight_value_rule)
        assert insight_value_rule is not None
        self.assertNotIn("ellipsis", insight_value_rule.group(1))
        self.assertNotIn("nowrap", insight_value_rule.group(1))

        instrumentation = """
<script>
  window.addEventListener("load", () => {
    const root = document.documentElement;
    const headline = Array.from(
      document.querySelectorAll(
        ".insights:not(.advanced-insights) > .insight"
      )
    );
    root.dataset.testInsightCount = String(headline.length);
    root.dataset.testInsightKeys =
      headline.map(card => card.dataset.insightKey).join(",");
    root.dataset.testInsightIcons = String(headline.filter(
      card => !!card.querySelector(".insight-icon svg")
    ).length);
    root.dataset.testInsightLayouts =
      headline.map(card => card.dataset.insightLayout).join(",");
    root.dataset.testInsightCovers = String(headline.filter(
      card => !!card.querySelector(".insight-cover")
    ).length);
    root.dataset.testInsightShare = String(headline.filter(
      card => !!card.querySelector(".insight-share")
    ).length);
    const featureCard = headline.find(
      card => card.dataset.insightKey === "most-read"
    );
    const shareTrack = featureCard
      ? featureCard.querySelector(".insight-share-track")
      : null;
    const shareFill = shareTrack ? shareTrack.querySelector("i") : null;
    root.dataset.testFeatureShareWidth = shareFill
      ? String(Math.round(shareFill.getBoundingClientRect().width))
      : "0";
    root.dataset.testFeatureShareTrackWidth = shareTrack
      ? String(Math.round(shareTrack.getBoundingClientRect().width))
      : "0";
    const shareLabel = featureCard
      ? featureCard.querySelector(".insight-share-label")
      : null;
    root.dataset.testFeatureShareLabel =
      shareLabel ? (shareLabel.textContent || "").trim() : "";
    root.dataset.testInsightValues = String(headline.every(card => {
      const value = card.querySelector(".insight-value");
      return !!value && (value.textContent || "").trim().length > 0;
    }));
    root.dataset.testInsightDetails = String(headline.every(card => {
      const detail = card.querySelector(".insight-detail");
      return !!detail && (detail.textContent || "").trim().length > 0;
    }));
    root.dataset.testInsightMetrics = headline.map(card => {
      const detail = card.querySelector(".insight-detail");
      return detail ? (detail.textContent || "").trim() : "";
    }).join("|");
    root.dataset.testInsightBilingual = String(headline.every(
      card => !!card.querySelector(".i18n-zh") &&
        !!card.querySelector(".i18n-en")
    ));
    root.dataset.testInsightTruncated = String(headline.some(card => {
      const value = card.querySelector(".insight-value");
      return value && value.scrollWidth > value.clientWidth + 1;
    }));
    root.dataset.testInsightOverflow = String(headline.some(card => {
      const rect = card.getBoundingClientRect();
      const parentRect = card.parentElement.getBoundingClientRect();
      return rect.width > 0 && (
        rect.left < parentRect.left - 1 ||
        rect.right > parentRect.right + 1
      );
    }));
    root.dataset.testInsightNowrap = String(headline.some(card => {
      const value = card.querySelector(".insight-value");
      return !!value && getComputedStyle(value).whiteSpace === "nowrap";
    }));
    root.dataset.testInsightLongTitleWrap = String(
      Array.from(document.querySelectorAll(
        ".insights:not(.advanced-insights) > .insight"
      )).every(card => {
        const value = card.querySelector(".insight-value");
        if (!value) { return false; }
        const rect = value.getBoundingClientRect();
        const cardRect = card.getBoundingClientRect();
        return rect.right <= cardRect.right + 1;
      })
    );
    root.dataset.testQualityInAdvanced = String(
      !!document.querySelector(
        '#advanced-data [data-insight-key="data-quality"]'
      )
    );
  });
</script>
"""
        self.assertIn("</body>", document)
        instrumented = document.replace(
            "</body>", instrumentation + "</body>"
        )
        snapshots = browser_viewport_snapshots(instrumented)

        for viewport, snapshot_document in snapshots.items():
            attributes = html_root_attributes(str(snapshot_document))
            with self.subTest(viewport=viewport):
                self.assertEqual(
                    attributes.get("data-test-insight-count"), "3"
                )
                self.assertEqual(
                    attributes.get("data-test-insight-keys"),
                    "most-read,longest-book,fastest-wpm",
                )
                self.assertEqual(
                    attributes.get("data-test-insight-icons"), "3"
                )
                self.assertEqual(
                    attributes.get("data-test-insight-layouts"),
                    "feature,cover,cover",
                )
                self.assertEqual(
                    attributes.get("data-test-insight-covers"), "3"
                )
                self.assertEqual(
                    attributes.get("data-test-insight-share"), "1"
                )
                share_width = int(
                    attributes.get("data-test-feature-share-width") or "0"
                )
                share_track = int(
                    attributes.get(
                        "data-test-feature-share-track-width"
                    ) or "0"
                )
                self.assertGreater(share_width, 0)
                self.assertGreaterEqual(share_track, share_width)
                self.assertRegex(
                    attributes.get("data-test-feature-share-label") or "",
                    r"占累计阅读 \d+%",
                )
                self.assertIn(
                    "WPM",
                    attributes.get("data-test-insight-metrics") or "",
                )
                self.assertEqual(
                    attributes.get("data-test-insight-values"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-insight-details"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-insight-bilingual"), "true"
                )
                self.assertEqual(
                    attributes.get("data-test-insight-truncated"), "false"
                )
                self.assertEqual(
                    attributes.get("data-test-insight-overflow"), "false"
                )
                self.assertEqual(
                    attributes.get("data-test-insight-nowrap"), "false"
                )
                self.assertEqual(
                    attributes.get("data-test-insight-long-title-wrap"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-quality-in-advanced"), "true"
                )
                self.assertEqual(
                    attributes.get("data-overflow-x"), "false"
                )
                self.assertEqual(
                    attributes.get("data-overflow-elements"), ""
                )
                self.assertEqual(
                    attributes.get("data-overlap-count"), "0",
                    attributes.get("data-overlap-elements"),
                )
                self.assertEqual(
                    attributes.get("data-truncated-numbers"), ""
                )


    def test_cover_assets_are_embedded_as_data_urls(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            cover = root / "cover.jpg"
            cover.write_bytes(b"jpeg-bytes")
            urls = kindle_report.cover_urls_for_html(
                {"asin:TEST": cover},
                root / "dashboard.html",
            )
            self.assertEqual(
                urls["asin:TEST"],
                "data:image/jpeg;base64," + "anBlZy1ieXRlcw==",
            )
            self.assertNotIn("covers", urls["asin:TEST"])


def load_tests(
    loader: unittest.TestLoader,
    tests: unittest.TestSuite,
    pattern: str | None,
) -> unittest.TestSuite:
    del pattern
    suite = unittest.TestSuite()
    suite.addTests(tests)
    suite.addTests(loader.loadTestsFromTestCase(kindle_report.KindleReportTests))
    return suite


if __name__ == "__main__":
    unittest.main()
class UpdatedDashboardContractTests(unittest.TestCase):
    @staticmethod
    def _row(**overrides):
        row = {
            "title": "Book",
            "author": "Author",
            "asin": "B000CONTRACT",
            "reading_mode": "book",
            "total_time_ms": "3600000",
            "total_words": "12000",
            "words_read": "12000",
            "book_words": "15000",
            "book_words_exact": "15000",
            "book_words_known": "15000",
            "word_count_source": "en-regex",
            "word_count_language": "en",
            "wpm": "200",
            "wpm_overall": "200",
            "wpm_reading": "200",
            "progress_percent": "80",
            "progress": "80",
            "last_read": "2026-10-02T10:00:00+08:00",
            "has_data": "True",
        }
        row.update(overrides)
        return row

    def _document(
        self,
        rows=None,
        profile=None,
        temp_dir=None,
        cover_urls=None,
    ):
        rows = rows or [self._row()]
        root = Path(temp_dir or tempfile.mkdtemp())
        return kindle_report.build_html(
            rows,
            root / "kindle_reading_live.csv",
            str(root),
            achievement_profile=profile,
            cover_urls=cover_urls,
            generated_at="2026-10-02T12:00:00+08:00",
        )

    def test_dashboard_hides_technical_panels_sources_and_urls(self):
        document = self._document()
        for token in (
            "漫画 OCR 估算",
            "OCR 状态",
            "OCR 错误",
            "高级数据",
            "数据质量",
            "原始快照",
            "Kindle 原始",
            "正文精确",
            "Kindle 外推",
            "fmcache 回退",
            'id="advanced-data"',
            'data-column="quality"',
            "data-quality-score=",
            'data-insight-key="data-quality"',
        ):
            self.assertNotIn(token, document)
        self.assertNotIn("http://", document)
        self.assertNotIn("https://", document)
        # Tokenizer labels are included in the document so the appearance
        # switch can reveal them without rebuilding the report, but the
        # default UI keeps them and all technical panels hidden.
        self.assertIn('data-show-token-labels="false"', document)
        self.assertIn(".token-tag {", document)
        self.assertIn(
            'html[data-show-token-labels="true"] .token-tag',
            document,
        )
        self.assertNotIn('id="advanced-data"', document)
        self.assertNotIn('id="quality-summary"', document)
        self.assertIn("这是属于你的阅读星空", document)

    def test_card_language_resolution_and_old_csv_compatibility(self):
        manga_without_language = self._row(
            title="Manga Without Language",
            asin="M000LANG01",
            reading_mode="manga",
            word_count_source="not-applicable",
            word_count_language="",
            ocr_language="",
        )
        manga_with_ocr = self._row(
            title="Manga With OCR",
            asin="M000LANG02",
            reading_mode="manga",
            word_count_source="not-applicable",
            word_count_language="",
            ocr_language="ja",
        )
        manga_priority = self._row(
            title="Manga OCR Priority",
            asin="M000LANG03",
            reading_mode="manga",
            word_count_source="en-regex",
            word_count_language="en",
            ocr_language="zh",
        )
        old_english = {
            "title": "Old English CSV",
            "author": "Author",
            "asin": "B000OLD001",
            "word_count_source": "en-regex",
            "total_time_ms": "60000",
            "total_words": "100",
            "progress_percent": "10",
        }
        old_chinese = {
            **old_english,
            "title": "旧中文 CSV",
            "asin": "B000OLD002",
            "word_count_source": "jieba",
        }
        old_japanese = {
            **old_english,
            "title": "旧日本語 CSV",
            "asin": "B000OLD003",
            "word_count_source": "janome",
        }
        book_without_language = {
            **old_english,
            "title": "No Language",
            "asin": "B000OLD004",
            "word_count_source": "",
        }

        self.assertEqual(
            kindle_report.resolved_card_language_label(
                manga_without_language
            ),
            "漫画",
        )
        self.assertEqual(
            kindle_report.resolved_card_language_label(manga_with_ocr),
            "日文",
        )
        self.assertEqual(
            kindle_report.resolved_card_language_label(manga_priority),
            "中文",
        )
        self.assertEqual(
            kindle_report.resolved_card_language(old_english),
            "en",
        )
        self.assertEqual(
            kindle_report.resolved_card_language(old_chinese),
            "zh",
        )
        self.assertEqual(
            kindle_report.resolved_card_language(old_japanese),
            "ja",
        )
        self.assertEqual(
            kindle_report.resolved_card_language_label(book_without_language),
            "未记录",
        )
        self.assertEqual(
            kindle_report.resolved_tokenizer_source(manga_with_ocr),
            "not-applicable",
        )
        self.assertEqual(
            kindle_report.tokenizer_label("en-regex"),
            "英文正则",
        )
        self.assertEqual(kindle_report.tokenizer_label("jieba"), "jieba")
        self.assertEqual(kindle_report.tokenizer_label("janome"), "Janome")

        document = self._document(
            rows=[
                manga_without_language,
                manga_with_ocr,
                manga_priority,
                old_english,
                old_chinese,
                old_japanese,
                book_without_language,
            ]
        )
        self.assertNotIn("未识别", document)
        self.assertIn("漫画不适用", document)
        self.assertIn("OCR 估算", document)
        self.assertIn("英文正则", document)
        self.assertIn("Janome", document)

    def test_shelf_title_row_defaults_and_preference_contract(self):
        document = self._document()
        self.assertIn('data-shelf-title-style="a"', document)
        self.assertIn('data-shelf-position="title"', document)
        self.assertIn('data-shelf-default-mode="book"', document)
        self.assertIn('data-show-token-labels="false"', document)
        self.assertIn('id="shelf-title-style"', document)
        self.assertIn('id="shelf-position"', document)
        self.assertIn('id="shelf-default-mode"', document)
        self.assertIn('id="show-token-labels"', document)
        self.assertIn('id="shelf-mode-switch"', document)
        self.assertIn('id="shelf-fallback-button"', document)
        self.assertIn('id="shelf-fallback-menu"', document)
        for key in (
            "kindle-dashboard-shelf-title-style",
            "kindle-dashboard-shelf-position",
            "kindle-dashboard-shelf-default-mode",
            "kindle-dashboard-show-token-labels",
        ):
            self.assertIn(key, document)
        self.assertNotIn("<select id=\"shelf-mode\"", document)
        self.assertLess(
            document.index('id="shelf-mode-switch"'),
            document.index('id="shelf-prev"'),
        )

    def test_achievements_are_five_tracks_with_manga_default_off(self):
        document = self._document()
        for key in ("streak", "time", "books", "manga_volumes", "words"):
            self.assertIn(f'data-achievement-key="{key}"', document)
        self.assertNotIn('data-achievement-key="speed"', document)
        self.assertIn('id="manga-volumes-enabled"', document)
        self.assertIn(
            'kindle-dashboard-manga-volume-achievement',
            document,
        )
        achievements = kindle_report.build_achievements(
            kindle_report.build_summary([self._row()]),
            {"daily": [], "streak": {}},
            [self._row()],
        )
        self.assertEqual(achievements["total"], 5)
        self.assertEqual(achievements["stars_total"], 20)
        self.assertIn("C1.59 10.5", document)
        lit_style = document[
            document.index(".achievement-star.star-lit svg path"):
            document.index(".achievement-star.star-next svg path")
        ]
        self.assertNotIn("drop-shadow", lit_style)
        self.assertIn("drop-shadow", document)

    def test_star_sky_entry_and_history_are_available(self):
        root = Path(tempfile.mkdtemp())
        history = {
            "version": kindle_report.HISTORY_VERSION,
            "snapshots": [
                kindle_report.build_snapshot(
                    [self._row(total_words="12000")],
                    "2026-09-01T10:00:00+08:00",
                    str(root),
                )
            ],
        }
        existing = {
            "version": 1,
            "created_at": "2026-08-01T10:00:00+08:00",
            "baseline": {
                "captured_at": "2026-08-01T10:00:00+08:00",
                "total_time_ms": 0,
                "total_words": 0,
                "completed_book_count": 0,
                "read_book_count": 0,
                "book_count": 1,
            },
            "tiers": {},
            "levels": {"streak": 0, "time": 0, "books": 0, "words": 3, "speed": 2},
            "observed_values": {"words": 12000},
        }
        profile = kindle_report.build_achievement_profile(
            [self._row(total_words="12000")],
            history=history,
            captured_at="2026-10-02T12:00:00+08:00",
            source_path=str(root),
            existing_profile=existing,
        )
        self.assertEqual(profile["levels"]["words"], 3)
        self.assertEqual(profile["levels"]["speed"], 2)
        old_stars = [
            entry
            for entry in profile["star_history"]
            if entry["metric"] == "words" and entry["level"] <= 3
        ]
        self.assertTrue(old_stars)
        self.assertTrue(all(entry["approximate"] for entry in old_stars))
        self.assertTrue(
            all(
                entry["source"] in {"history-inference", "legacy-profile"}
                for entry in old_stars
            )
        )
        document = self._document(profile=profile, temp_dir=str(root))
        self.assertIn('id="star-sky-open"', document)
        self.assertIn('id="star-sky-dialog"', document)
        self.assertIn("data-sky-metric", document)
        self.assertIn("data-sky-unlocked-at", document)

    def test_text_and_ocr_variants_control_ocr_truthfully(self):
        text_parser = kindle_report.build_arg_parser("text")
        with contextlib.redirect_stderr(io.StringIO()):
            with self.assertRaises(SystemExit):
                text_parser.parse_args(["--manga-ocr"])
        ocr_parser = kindle_report.build_arg_parser("ocr")
        parsed = ocr_parser.parse_args(["--no-manga-ocr"])
        self.assertFalse(parsed.manga_ocr)
        with unittest.mock.patch.dict(
            os.environ,
            {"KINDLE_READER_VARIANT": "text"},
        ):
            self.assertEqual(kindle_report.runtime_variant(), "text")
            self.assertEqual(
                kindle_report.apply_manga_ocr_to_csv(
                    Path(tempfile.mkdtemp()) / "books.csv",
                    Path(tempfile.mkdtemp()) / "_ocr",
                    argparse.Namespace(),
                ),
                0,
            )
            rows = [
                self._row(
                    title="Text-only fixture",
                    asin="B000TEXTON",
                    reading_mode="book",
                ),
                self._row(
                    title="Synthetic manga fixture",
                    asin="B000MANGA1",
                    reading_mode="manga",
                    word_count_source="not-applicable",
                    word_count_language="",
                ),
            ]
            document = kindle_report.build_html(
                rows,
                Path("fixture.csv"),
                "fixture",
                generated_at="2026-11-27T12:00:00+08:00",
            )
            self.assertIn('data-ui-variant="text"', document)
            self.assertIn("Text-only fixture", document)
            self.assertNotIn("Synthetic manga fixture", document)
            self.assertEqual(
                document.count('<article class="achievement-card'),
                4,
            )

    @unittest.skipUnless(
        CHROMIUM_BROWSER, "Chromium browser is not available"
    )
    def test_shelf_style_position_hidden_and_storage_contract(self):
        rows = [
            self._row(
                title=(
                    "A Very Long Shelf Title Used To Verify Desktop "
                    "Truncation Without Overlap"
                ),
                asin="B000LONG01",
                reading_mode="book",
            ),
            self._row(
                title="Manga Shelf Entry",
                asin="M000LONG01",
                reading_mode="manga",
                word_count_source="not-applicable",
                word_count_language="",
                ocr_language="zh",
            ),
            self._row(
                title="Manga OCR Entry",
                asin="M000LONG02",
                reading_mode="manga",
                word_count_source="ocr-estimate",
                word_count_language="",
                ocr_language="ja",
                ocr_word_count="5000",
                ocr_char_count="9000",
                ocr_confidence="0.91",
                ocr_page_coverage="100",
                ocr_status="ok",
                ocr_pages_total="100",
                ocr_pages_processed="100",
                total_words="0",
                words_read="0",
                book_words="0",
                book_words_exact="0",
                wpm="0",
            ),
        ]
        document = self._document(rows=rows)
        self.assertIn('id="shelf-mode-switch"', document)
        self.assertIn('data-shelf-title-style="a"', document)
        self.assertIn('data-shelf-position="title"', document)
        self.assertIn('data-shelf-default-mode="book"', document)
        instrumentation = """
<script>
  window.addEventListener("load", async () => {
    const root = document.documentElement;
    const title = document.querySelector(".shelf-title-copy h2");
    const heading = document.querySelector(".library-heading");
    const intersects = (left, right) =>
      Math.min(left.right, right.right) - Math.max(left.left, right.left) > 1 &&
      Math.min(left.bottom, right.bottom) - Math.max(left.top, right.top) > 1;

    root.dataset.testDefaults = [
      root.dataset.shelfTitleStyle,
      root.dataset.shelfPosition,
      root.dataset.shelfDefaultMode,
      root.dataset.showTokenLabels,
    ].join(",");
    root.dataset.testDefaultNearTitle = String(
      shelfModeSwitch.parentElement === shelfTitleCluster
    );
    root.dataset.testDefaultIconDisplay = getComputedStyle(
      shelfModeSwitch.querySelector(".shelf-mode-icon")
    ).display;
    const mangaToken = coverGrid.querySelector(
      '[data-book-id="asin:M000LONG01"] .token-tag'
    );
    const mangaOcrToken = coverGrid.querySelector(
      '[data-book-id="asin:M000LONG02"] .token-tag'
    );
    root.dataset.testTokenDefault = [
      root.dataset.showTokenLabels,
      getComputedStyle(mangaToken).display,
    ].join(",");
    setShowTokenLabels(true);
    root.dataset.testTokenEnabled = [
      root.dataset.showTokenLabels,
      getComputedStyle(mangaToken).display,
    ].join(",");
    root.dataset.testMangaTokenFiltered = String(
      mangaToken.textContent.includes("漫画不适用") &&
      !/(en-regex|jieba|Janome|英文正则|分词失败)/.test(
        mangaToken.textContent
      ) &&
      mangaOcrToken.textContent.includes("OCR 估算")
    );
    setShowTokenLabels(false);

    shelfTitleStyleSelect.value = "b";
    shelfTitleStyleSelect.dispatchEvent(new Event("change"));
    void shelfModeSwitch.offsetWidth;
    root.dataset.testStyleB = [
      root.dataset.shelfTitleStyle,
      getComputedStyle(
        shelfModeSwitch.querySelector(".shelf-mode-icon")
      ).display,
      getComputedStyle(
        shelfModeSwitch.querySelector(".shelf-mode-label-short")
      ).display,
      readPreference(SHELF_TITLE_STYLE_KEY) || "",
    ].join(",");

    shelfTitleStyleSelect.value = "c";
    shelfTitleStyleSelect.dispatchEvent(new Event("change"));
    void shelfModeSwitch.offsetWidth;
    await new Promise(resolve => {
      requestAnimationFrame(() => requestAnimationFrame(resolve));
    });
    root.dataset.testStyleC = [
      root.dataset.shelfTitleStyle,
      getComputedStyle(shelfModeSwitch).borderBottomStyle,
      getComputedStyle(shelfModeSwitch).borderRadius,
    ].join(",");

    shelfPositionSelect.value = "actions";
    shelfPositionSelect.dispatchEvent(new Event("change"));
    root.dataset.testActionsPosition = String(
      root.dataset.shelfPosition === "actions" &&
      shelfModeSwitch.parentElement === libraryActions &&
      readPreference(SHELF_POSITION_KEY) === "actions"
    );

    shelfPositionSelect.value = "hidden";
    shelfPositionSelect.dispatchEvent(new Event("change"));
    root.dataset.testHiddenPosition = String(
      root.dataset.shelfPosition === "hidden" &&
      getComputedStyle(shelfModeSwitch).display === "none" &&
      getComputedStyle(shelfFallbackWrap).display !== "none"
    );

    shelfDefaultModeSelect.value = "book";
    shelfDefaultModeSelect.dispatchEvent(new Event("change"));
    shelfFallbackButton.click();
    root.dataset.testFallbackMenuOpen = String(!shelfFallbackMenu.hidden);
    shelfFallbackMenu.querySelector(
      '[data-shelf-mode="manga"]'
    ).click();
    root.dataset.testMangaReachable = [
      root.dataset.shelfMode,
      readPreference(SHELF_DEFAULT_MODE_KEY) || "",
    ].join(",");
    root.dataset.testFallbackButtons = String(
      shelfFallbackMenu.querySelectorAll("button").length === 2 &&
      shelfFallbackMenu.getAttribute("role") === "group" &&
      !shelfFallbackMenu.querySelector("select")
    );

    writePreference(SHELF_TITLE_STYLE_KEY, "invalid");
    writePreference(SHELF_POSITION_KEY, "invalid");
    writePreference(SHELF_DEFAULT_MODE_KEY, "invalid");
    writePreference(SHOW_TOKEN_LABELS_KEY, "invalid");
    restoreShelfPreferences();
    root.dataset.testInvalidFallback = [
      root.dataset.shelfTitleStyle,
      root.dataset.shelfPosition,
      root.dataset.shelfDefaultMode,
      root.dataset.showTokenLabels,
    ].join(",");

    Object.defineProperty(window, "localStorage", {
      configurable: true,
      get() { throw new Error("storage blocked"); },
    });
    setShelfTitleStyle("c");
    setShelfPosition("hidden");
    setShelfDefaultMode("manga");
    setShowTokenLabels(true);
    root.dataset.testStorageUnavailable = [
      root.dataset.shelfTitleStyle,
      root.dataset.shelfPosition,
      root.dataset.shelfDefaultMode,
      root.dataset.showTokenLabels,
    ].join(",");

    setShelfTitleStyle("a");
    setShelfPosition("title");
    setShelfDefaultMode("book");
    setShowTokenLabels(false);
    setShelfMode("book");
    title.textContent = (
      "A Very Long Shelf Title Used To Verify Desktop Truncation " +
      "Without Overlap Across Wider Windows"
    );
    updateLayoutMetrics();
    const titleRect = title.getBoundingClientRect();
    const switchRect = shelfModeSwitch.getBoundingClientRect();
    const headingRect = heading.getBoundingClientRect();
    root.dataset.testTitleSwitchNoOverlap = String(
      !intersects(titleRect, switchRect) &&
      titleRect.right <= switchRect.left + 1
    );
    root.dataset.testTitleSingleRow = String(
      switchRect.top < titleRect.bottom - 1 &&
      switchRect.bottom > titleRect.top + 1
    );
    root.dataset.testHeadingFits = String(
      headingRect.left >= -1 &&
      headingRect.right <= root.clientWidth + 1 &&
      root.dataset.overflowX === "false"
    );
  });
</script>
"""
        self.assertIn("</body>", document)
        instrumented = document.replace(
            "</body>", instrumentation + "</body>"
        )
        snapshots = browser_viewport_snapshots(
            instrumented,
            ((1024, 900), (1440, 900)),
        )
        for viewport, snapshot_document in snapshots.items():
            attributes = html_root_attributes(str(snapshot_document))
            with self.subTest(viewport=viewport):
                self.assertEqual(
                    attributes.get("data-test-defaults"),
                    "a,title,book,false",
                )
                self.assertEqual(
                    attributes.get("data-test-default-near-title"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-default-icon-display"),
                    "none",
                )
                self.assertEqual(
                    attributes.get("data-test-token-default"),
                    "false,none",
                )
                self.assertEqual(
                    attributes.get("data-test-token-enabled"),
                    "true,flex",
                )
                self.assertEqual(
                    attributes.get("data-test-manga-token-filtered"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-style-b"),
                    "b,block,block,b",
                )
                self.assertEqual(
                    attributes.get("data-test-style-c"),
                    "c,solid,0px",
                )
                self.assertEqual(
                    attributes.get("data-test-actions-position"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-hidden-position"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-fallback-menu-open"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-manga-reachable"),
                    "manga,book",
                )
                self.assertEqual(
                    attributes.get("data-test-fallback-buttons"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-invalid-fallback"),
                    "a,title,book,false",
                )
                self.assertEqual(
                    attributes.get("data-test-storage-unavailable"),
                    "c,hidden,manga,true",
                )
                self.assertEqual(
                    attributes.get("data-test-title-switch-no-overlap"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-title-single-row"),
                    "true",
                )
                self.assertEqual(
                    attributes.get("data-test-heading-fits"),
                    "true",
                )

    @unittest.skipUnless(
        CHROMIUM_BROWSER, "Chromium browser is not available"
    )
    def test_shelf_paging_achievement_toggle_greeting_and_share_export(self):
        rows = [
            self._row(
                title=f"Book {index}",
                asin=f"B{index:07d}",
                reading_mode="book",
                total_time_ms=str(3600000 - index * 10000),
            )
            for index in range(1, 10)
        ]
        rows.extend(
            [
                self._row(
                    title="Manga One",
                    asin="M0000001",
                    reading_mode="manga",
                    word_count_source="not-applicable",
                    total_words="0",
                    words_read="0",
                    book_words="0",
                    book_words_exact="0",
                    wpm="0",
                    total_time_ms="120000",
                ),
                self._row(
                    title="Manga Two",
                    asin="M0000002",
                    reading_mode="manga",
                    word_count_source="not-applicable",
                    total_words="0",
                    words_read="0",
                    book_words="0",
                    book_words_exact="0",
                    wpm="0",
                    total_time_ms="90000",
                ),
            ]
        )
        cover_url = (
            "data:image/png;base64,"
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0l"
            "EQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="
        )
        cover_urls = {
            kindle_report.stable_book_id(rows[index]): cover_url
            for index in range(3)
        }
        document = self._document(rows=rows, cover_urls=cover_urls)
        self.assertLess(
            document.index('id="shelf-mode-switch"'),
            document.index('id="shelf-prev"'),
        )
        self.assertLess(
            document.index('id="shelf-prev"'),
            document.index('class="cover-count-badge"'),
        )
        instrumentation = """
<script>
  window.addEventListener("load", async () => {
  const shelfCards = Array.from(
    document.querySelectorAll(".book-cover-card[data-book-id]")
  );
  document.documentElement.dataset.testBookShelfDefault =
    document.documentElement.dataset.shelfMode || "missing";
  document.documentElement.dataset.testBookVisible =
    String(shelfCards.filter(card => !card.hidden).length);
  document.getElementById("shelf-next").click();
  document.documentElement.dataset.testBookPage2 =
    document.getElementById("shelf-page-status").textContent;
  document.querySelector('[data-shelf-mode="manga"]').click();
  document.documentElement.dataset.testMangaVisible =
    String(shelfCards.filter(card => !card.hidden).length);
  document.documentElement.dataset.testMangaMode =
    document.documentElement.dataset.shelfMode;

  const achievementToggle = document.getElementById("achievement-enabled");
  achievementToggle.checked = false;
  achievementToggle.dispatchEvent(new Event("change"));
  document.documentElement.dataset.testAchievementDisabled =
    document.getElementById("achievement-panel").hidden ? "true" : "false";
  achievementToggle.checked = true;
  achievementToggle.dispatchEvent(new Event("change"));
  document.documentElement.dataset.testAchievementEnabled =
    document.getElementById("achievement-panel").hidden ? "false" : "true";

  const greetingMode = document.getElementById("greeting-mode");
  greetingMode.value = "evening";
  greetingMode.dispatchEvent(new Event("change"));
  document.documentElement.dataset.testEveningTone =
    document.documentElement.dataset.greetingTone;
  document.documentElement.dataset.testEveningInk =
    getComputedStyle(document.documentElement).getPropertyValue("--greeting-ink").trim();

  const share = document.getElementById("share-export");
  share.click();
  document.documentElement.dataset.testShareDialog =
    document.getElementById("share-dialog").open ? "true" : "false";
  document.documentElement.dataset.testShareFormats =
    Array.from(document.getElementById("share-format").options)
      .map(option => option.value).join(",");
  const shareBookRowsForCanvas = await loadShareBookRows(3);
  document.documentElement.dataset.testShareCoverCount =
    String(shareBookRowsForCanvas.filter(row => row.image).length);
  const longCanvas = createLongShareCanvas(shareBookRowsForCanvas);
  const setCanvas = createBooksSetCanvas(shareBookRowsForCanvas);
  document.documentElement.dataset.testShareCanvas =
    longCanvas.width + "x" + longCanvas.height + "," +
    setCanvas.width + "x" + setCanvas.height;
  document.documentElement.dataset.testShareCanvasBytes =
    String(longCanvas.toDataURL("image/png").length);
  document.getElementById("share-close").click();
  });
</script>
"""
        instrumented = document.replace(
            "</body>", instrumentation + "</body>"
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            test_html = Path(temp_dir) / "dashboard-shelf-share.html"
            test_html.write_text(instrumented, encoding="utf-8")
            snapshot = browser_dom_snapshot(
                CHROMIUM_BROWSER, test_html, 390, 844
            )
        attributes = html_root_attributes(snapshot)
        self.assertEqual(attributes.get("data-test-book-shelf-default"), "book")
        self.assertEqual(attributes.get("data-test-book-visible"), "4")
        self.assertEqual(attributes.get("data-test-book-page2"), "2/3")
        self.assertEqual(attributes.get("data-test-manga-visible"), "2")
        self.assertEqual(attributes.get("data-test-manga-mode"), "manga")
        self.assertEqual(attributes.get("data-test-achievement-disabled"), "true")
        self.assertEqual(attributes.get("data-test-achievement-enabled"), "true")
        self.assertEqual(attributes.get("data-test-evening-tone"), "night")
        self.assertIn("#", attributes.get("data-test-evening-ink") or "")
        self.assertEqual(attributes.get("data-test-share-dialog"), "true")
        self.assertEqual(attributes.get("data-test-share-formats"), "long,set")
        self.assertEqual(
            attributes.get("data-test-share-canvas"),
            "1080x1920,1080x1350",
        )
        self.assertEqual(
            attributes.get("data-test-share-cover-count"),
            "3",
        )
        self.assertGreater(
            int(attributes.get("data-test-share-canvas-bytes") or "0"),
            10000,
        )


_LEGACY_LOAD_TESTS = load_tests


def _without_legacy_ui_tests(suite):
    filtered = unittest.TestSuite()
    for test in suite:
        if isinstance(test, unittest.TestSuite):
            filtered.addTests(_without_legacy_ui_tests(test))
            continue
        if getattr(test, "_testMethodName", "") in _LEGACY_UI_TEST_NAMES:
            continue
        filtered.addTest(test)
    return filtered


def _suite_test_ids(suite):
    identifiers = set()
    for test in suite:
        if isinstance(test, unittest.TestSuite):
            identifiers.update(_suite_test_ids(test))
        else:
            identifiers.add(test.id())
    return identifiers


def load_tests(loader, tests, ignore):
    suite = _LEGACY_LOAD_TESTS(loader, tests, ignore)
    filtered = _without_legacy_ui_tests(suite)
    identifiers = _suite_test_ids(filtered)
    for test in loader.loadTestsFromTestCase(UpdatedDashboardContractTests):
        if test.id() not in identifiers:
            filtered.addTest(test)
    return filtered
