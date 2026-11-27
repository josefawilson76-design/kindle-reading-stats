#!/usr/bin/env python3
"""Build a standalone visual HTML report from Kindle reading CSV data.

Packaging notes:
    Install the build dependency into an isolated directory:

        python -m pip install --target .tmp-pip pyinstaller

    Build the Windows single-file executable:

        python kindle_report.py --build-exe

    Runtime dependencies are Pillow, jieba, Janome, and openpyxl. The project
    vendor directory is added to PyInstaller's module path, so the tokenizers
    do not need to be installed in the build interpreter. Calibre and KFX
    Input are external dependencies used for exact KFX word counts.
"""

from __future__ import annotations

import argparse
import base64
import csv
import hashlib
import html
import io
import json
import math
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import time
import unicodedata
import unittest
import webbrowser
from datetime import date, datetime, timedelta
from html.parser import HTMLParser
from pathlib import Path
from typing import (
    Any,
    Dict,
    List,
    Mapping,
    Optional,
    Sequence,
    Tuple,
)

try:
    from PIL import Image, ImageOps
except ImportError:  # pragma: no cover - exercised only without Pillow
    Image = None
    ImageOps = None

vendor_path = Path(__file__).resolve().parent / "vendor"
if vendor_path.is_dir():
    # Appending keeps an unreadable or partial vendor copy from shadowing
    # working site-packages while still letting frozen builds fall back to it.
    sys.path.append(str(vendor_path))

try:
    import kindle_reader  # type: ignore
except ImportError:  # pragma: no cover - report-only mode still works
    kindle_reader = None

try:
    import kindle_manga_ocr  # type: ignore
except ImportError:  # pragma: no cover - OCR stays disabled without it
    kindle_manga_ocr = None


VERSION = "1.13.0"
HISTORY_VERSION = 1
HISTORY_FILENAME = "reading_history.json"
DEFAULT_HISTORY_LIMIT = 0
DEFAULT_CSV_FILENAME = "kindle_reading_live.csv"
DEFAULT_HTML_FILENAME = "kindle_reading_dashboard.html"
DEFAULT_XLSX_FILENAME = "kindle_reading_dashboard.xlsx"
DEFAULT_CACHE_FILENAME = "word_counts_cache.json"
DEFAULT_WORDCOUNT_DIRECTORY = "_wordcount"
DEFAULT_BOOK_MODES_FILENAME = "book_modes.json"
DEFAULT_ACHIEVEMENT_PROFILE_FILENAME = "achievement_profile.json"
DEFAULT_OCR_CACHE_DIRNAME = "_ocr"
DEFAULT_OCR_MIN_CONFIDENCE = 0.70
ACHIEVEMENT_PROFILE_VERSION = 5
DEFAULT_COMPLETED_BOOK_THRESHOLD_PERCENT = 95.0
SETTINGS_VERSION = 1
SETTINGS_FILENAME = "settings.json"
DEFAULT_SETTINGS = {
    "version": SETTINGS_VERSION,
    "kindle_path": "",
    "output_dir": "",
    "language": "auto",
    "exact_word_count": False,
    "open_html": True,
    "history_limit": DEFAULT_HISTORY_LIMIT,
    "book_modes_file": DEFAULT_BOOK_MODES_FILENAME,
    "achievement_profile_file": DEFAULT_ACHIEVEMENT_PROFILE_FILENAME,
    "completed_book_threshold_percent": (
        DEFAULT_COMPLETED_BOOK_THRESHOLD_PERCENT
    ),
    "manga_ocr": False,
    "manga_classify": True,
    "manga_volume_achievement": False,
    "ocr_language": "auto",
    "ocr_engine": "auto",
    "ocr_dpi": 240,
    "ocr_mode": "classify",
    "ocr_workers": 0,
    "ocr_sample_pages": 6,
    "ocr_min_confidence": 0.70,
    "ocr_cache_dir": DEFAULT_OCR_CACHE_DIRNAME,
}
MINIMUM_FREE_SPACE_BYTES = 128 * 1024 * 1024
COVER_VERSION = 1
COVER_DIRECTORY_NAME = "_covers"
COVER_MAX_SIZE = (420, 620)
EBOOK_EXTENSIONS = {
    ".azw",
    ".azw3",
    ".azw4",
    ".epub",
    ".kfx",
    ".kfx-zip",
    ".mobi",
    ".pdf",
}
ACCENTS = ("#2f7d61", "#3478a8", "#c57a20", "#c65c4a", "#6f5aa8", "#3d8f8a")
PERIOD_CHART_COLORS = (
    "var(--green)",
    "var(--blue)",
    "var(--amber)",
    "var(--coral)",
    "var(--violet)",
    "var(--teal)",
)
ACHIEVEMENT_STEPS = {
    "streak": (3.0, 7.0, 14.0, 30.0),
    "time": (1.0, 5.0, 10.0, 25.0),
    "books": (1.0, 3.0, 5.0, 10.0),
    "manga_volumes": (1.0, 3.0, 5.0, 10.0),
    "words": (10000.0, 50000.0, 100000.0, 250000.0),
    "speed": (10.0, 25.0, 50.0, 100.0),
}
ACHIEVEMENT_TRACK_KEYS = (
    "streak",
    "time",
    "books",
    "manga_volumes",
    "words",
)
ACHIEVEMENT_ENCOURAGEMENTS = {
    "streak": (
        ("阅读开始连起来了，保持这个节奏。", "Your reading streak is starting."),
        ("连续阅读的节奏更稳了。", "The reading rhythm is getting steadier."),
        ("连续阅读正在成为习惯。", "Reading is becoming a habit."),
        ("稳定坚持，已经是很了不起的积累。", "Steady reading is a remarkable habit."),
    ),
    "time": (
        ("阅读时间又向前走了一步。", "Your reading time moved forward again."),
        ("阅读时长在稳步累积。", "Reading time is building steadily."),
        ("已经积累出一段安静的阅读深度。", "You have built a deeper reading habit."),
        ("长时间的坚持，让阅读成了生活的一部分。", "Long-term reading has become part of life."),
    ),
    "books": (
        ("一本读完，书架亮起新星。", "A finished book lights up the shelf."),
        ("又读完一本，阅读记录更有厚度。", "Another finished book adds depth."),
        ("读完的书在增加，积累很扎实。", "Your finished books are growing steadily."),
        ("一本本读下来，书架已经很有分量。", "Your shelf has real weight now."),
    ),
    "manga_volumes": (
        ("第一本漫画读完，新的书柜亮起来了。", "Your first completed manga volume lights a new shelf."),
        ("三卷读完，系列开始有了完整轮廓。", "Three volumes complete a clearer series shape."),
        ("五卷读完，漫画收藏有了厚度。", "Five completed volumes give the shelf real weight."),
        ("十卷读完，这是一段很扎实的漫画旅程。", "Ten completed volumes mark a substantial manga journey."),
    ),
    "words": (
        ("阅读词数开始累积。", "Your reading word count is growing."),
        ("词数又上了一个台阶。", "Your word count reached a new step."),
        ("阅读量在稳定增长。", "Your reading volume is growing steadily."),
        ("长期累积的词数，已经是可观的阅读旅程。", "Your long-term word count is a real journey."),
    ),
    "speed": (
        ("阅读节奏略有提升。", "Your reading pace is improving."),
        ("速度又稳了一点。", "A steadier reading speed."),
        ("保持舒适节奏，继续前进。", "A comfortable, faster pace."),
        ("速度台阶又上了一层。", "Another speed tier reached."),
    ),
}
# Inline, dependency-free icons for the headline insight cards. No external
# assets are used so the dashboard stays fully offline.
INSIGHT_ICONS = {
    "most-read": (
        '<path d="M12 7v14"></path>'
        '<path d="M3 18a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1h5a4 4 0 0 1 4 4 '
        '4 4 0 0 1 4-4h5a1 1 0 0 1 1 1v13a1 1 0 0 1-1 1h-6a3 3 0 0 0-3 3 '
        '3 3 0 0 0-3-3z"></path>'
    ),
    "longest-book": (
        '<path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"></path>'
        '<path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 '
        '6.5 2z"></path>'
    ),
    "fastest-wpm": (
        '<path d="m12 14 4-4"></path>'
        '<path d="M3.34 19a10 10 0 1 1 17.32 0"></path>'
    ),
}
FMCACHE_RELATIVE_PATH = Path("system") / "fmcache" / "fmcache.db"
FMCACHE_HEATMAP_DAYS = 84
FMCACHE_MAX_SESSION_MS = 24 * 60 * 60 * 1000
PERIOD_REPORT_KEYS = ("week", "month", "year")
PERIOD_GAP_LIMIT_DAYS = {
    "week": 2,
    "month": 7,
    "year": 31,
}
PERIOD_MIN_SPEED_TIME_MS = 30 * 60 * 1000
PERIOD_MIN_SPEED_WORDS = 3000
PERIOD_WPM_NORMAL_RANGE = (40.0, 800.0)
QUALITY_MIN_TIME_MS = 5 * 60 * 1000
QUALITY_MIN_WORDS = 300
WPM_NORMAL_RANGES = {
    "en": (40.0, 500.0),
    "zh": (80.0, 800.0),
    "ja": (80.0, 800.0),
}
DEFAULT_WPM_NORMAL_RANGE = (40.0, 600.0)
PROGRESS_DISCREPANCY_PP = 5.0
XLSX_HEADER_LABELS = {
    "title": "书名 / Title",
    "author": "作者 / Author",
    "asin": "ASIN",
    "total_time_ms": "阅读毫秒 / Reading ms",
    "minutes": "分钟 / Minutes",
    "hours": "小时 / Hours",
    "total_words": "Kindle 已读词数 / Kindle words read",
    "words_read": "已读词数 / Words read",
    "book_words": "Kindle 外推总词数 / Kindle book-word estimate",
    "book_words_exact": "正文精确总词数 / Exact body words",
    "book_words_known": "已知词数 / Known words",
    "book_word_coverage_percent": "词数覆盖率 / Word coverage %",
    "word_count_source": "词数来源 / Word-count source",
    "word_count_language": "词数语言 / Word-count language",
    "wpm": "WPM",
    "wpm_overall": "总体 WPM / Overall WPM",
    "wpm_reading": "阅读 WPM / Reading WPM",
    "progress_percent": "Kindle 进度 % / Kindle progress %",
    "progress": "进度 / Progress",
    "page": "页码 / Page",
    "annotations": "批注 / Annotations",
    "last_read": "最后阅读 / Last read",
    "ebook_file": "电子书文件 / Ebook file",
    "word_count_error": "词数错误 / Word-count error",
    "has_data": "有数据 / Has data",
    "source_file": "来源文件 / Source file",
    "series": "系列 / Series",
    "series_index": "系列序号 / Series index",
    "publisher": "出版社 / Publisher",
    "published_date": "出版日期 / Published date",
    "metadata_match": "元数据匹配 / Metadata match",
    "reading_mode": "阅读模式 / Reading mode",
    "metadata_tags": "元数据标签 / Metadata tags",
    "metadata_suggested_mode": "元数据建议模式 / Suggested mode",
    "reading_mode_source": "模式来源 / Mode source",
    "classification_confidence": "分类置信度 / Classification confidence",
    "classification_reason": "分类依据 / Classification reason",
    "series_key": "系列键 / Series key",
    "manga_pages_total": "漫画总页数（源文件）/ Manga source pages",
    "manga_pages_read": "已读页数 / Read pages",
    "manga_pages_source": "页数口径 / Page-count source",
    "manga_progress_source": "进度来源 / Progress source",
    "ocr_word_count": "OCR 估算词数 / OCR estimated words",
    "ocr_mode": "OCR 模式 / OCR mode",
    "ocr_word_count_is_estimate": "OCR 是否外推 / OCR extrapolated",
    "ocr_char_count": "OCR 估算字符数 / OCR characters",
    "ocr_confidence": "OCR 置信度 / OCR confidence",
    "ocr_engine": "OCR 引擎 / OCR engine",
    "ocr_language": "OCR 语言 / OCR language",
    "ocr_page_coverage": "OCR 页面覆盖率 % / OCR page coverage %",
    "ocr_status": "OCR 状态 / OCR status",
    "ocr_error": "OCR 错误 / OCR error",
    "ocr_completed_at": "OCR 完成时间 / OCR completed",
    "ocr_cache_key": "OCR 缓存键 / OCR cache key",
    "ocr_source": "OCR 来源口径 / OCR source",
    "ocr_pages_total": "OCR 总页数 / OCR total pages",
    "ocr_pages_processed": "OCR 已处理页数 / OCR processed pages",
    "ocr_failed_pages": "OCR 失败页数 / OCR failed pages",
    "ocr_elapsed_seconds": "OCR 耗时秒 / OCR seconds",
    "ocr_peak_memory_mb": "OCR 峰值内存 MB / OCR peak memory MB",
    "ocr_min_confidence": "OCR 置信门槛 / OCR confidence gate",
    "ocr_sampled_pages": "OCR 采样页数 / OCR sampled pages",
    "ocr_extrapolated": "OCR 外推 / OCR extrapolated",
    "effective_book_words": "有效总词数 / Effective book words",
    "effective_words_read": "有效已读词数 / Effective words read",
    "effective_wpm": "有效 WPM / Effective WPM",
    "effective_source": "有效词数来源 / Effective-word source",
    "word_origin": "词数来源类别 / Word origin",
    "word_progress": "词数进度 % / Word progress %",
    "confidence_level": "置信度 / Confidence",
    "confidence_score": "置信分 / Confidence score",
    "quality_flags": "质量提示 / Quality flags",
    "time_delta_ms": "本次时长增量 ms / Time delta ms",
    "words_delta": "本次词数增量 / Words delta",
    "progress_delta": "本次进度增量 pp / Progress delta pp",
}
XLSX_NUMERIC_FIELDS = {
    "total_time_ms",
    "minutes",
    "hours",
    "total_words",
    "words_read",
    "book_words",
    "book_words_exact",
    "book_words_known",
    "book_word_coverage_percent",
    "wpm",
    "wpm_overall",
    "wpm_reading",
    "progress_percent",
    "progress",
    "page",
    "annotations",
    "series_index",
    "ocr_word_count",
    "ocr_char_count",
    "ocr_confidence",
    "ocr_page_coverage",
    "ocr_pages_total",
    "ocr_pages_processed",
    "ocr_failed_pages",
    "ocr_elapsed_seconds",
    "ocr_peak_memory_mb",
    "ocr_sampled_pages",
    "classification_confidence",
    "manga_pages_total",
    "manga_pages_read",
    "effective_book_words",
    "effective_words_read",
    "effective_wpm",
}


def parse_float(value: Any) -> Optional[float]:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def parse_int(value: Any) -> Optional[int]:
    number = parse_float(value)
    return int(round(number)) if number is not None else None


def runtime_variant() -> str:
    """Return the active executable variant without splitting core code."""

    override = os.environ.get("KINDLE_READER_VARIANT", "").strip().casefold()
    if override in {"text", "ocr"}:
        return override
    name = Path(sys.executable).stem.casefold()
    if name.endswith("-text"):
        return "text"
    return "ocr"


def normalize_reading_mode(value: Any) -> str:
    normalized = str(value or "").strip().casefold()
    return normalized if normalized in {"book", "manga"} else "book"


def _normalize_mode_asin(value: Any) -> str:
    normalized = unicodedata.normalize("NFKC", str(value or ""))
    return re.sub(r"\s+", "", normalized).upper()


def normalize_book_modes(value: Any) -> Dict[str, Dict[str, str]]:
    raw = value if isinstance(value, dict) else {}
    result: Dict[str, Dict[str, str]] = {
        "by_asin": {},
        "by_title": {},
        "by_series": {},
    }
    by_asin = raw.get("by_asin")
    if isinstance(by_asin, dict):
        for asin, mode in by_asin.items():
            normalized_asin = _normalize_mode_asin(asin)
            if normalized_asin:
                result["by_asin"][normalized_asin] = normalize_reading_mode(mode)
    by_title = raw.get("by_title")
    if isinstance(by_title, dict):
        for title, mode in by_title.items():
            normalized = normalized_title(title)
            series = normalized_series_title(title)
            normalized_mode = normalize_reading_mode(mode)
            if normalized:
                result["by_title"][normalized] = normalized_mode
            if series:
                result["by_series"][series] = normalized_mode
    for raw_key, mode in raw.items():
        if str(raw_key) in {"version", "by_asin", "by_title"}:
            continue
        normalized_mode = normalize_reading_mode(mode)
        normalized_asin = _normalize_mode_asin(raw_key)
        if normalized_asin and re.fullmatch(
            r"[A-Za-z0-9]{8,}", str(raw_key).strip()
        ):
            result["by_asin"][normalized_asin] = normalized_mode
            continue
        normalized = normalized_title(raw_key)
        series = normalized_series_title(raw_key)
        if normalized:
            result["by_title"][normalized] = normalized_mode
        if series:
            result["by_series"][series] = normalized_mode
    return result


def load_book_modes(path: Optional[str]) -> Tuple[Dict[str, Dict[str, str]], Optional[str]]:
    if not path:
        return normalize_book_modes({}), None
    mapping_path = Path(path).expanduser()
    try:
        if not mapping_path.is_file():
            return normalize_book_modes({}), (
                "book mode mapping was not found: %s" % mapping_path
            )
        with mapping_path.open("r", encoding="utf-8-sig") as handle:
            decoded = json.load(handle)
    except (OSError, ValueError) as exc:
        return normalize_book_modes({}), (
            "book mode mapping could not be read (%s): %s"
            % (mapping_path, exc)
        )
    if not isinstance(decoded, dict):
        return normalize_book_modes({}), (
            "book mode mapping must be a JSON object: %s" % mapping_path
        )
    return normalize_book_modes(decoded), None


def reading_mode_for_book(
    row: Optional[Dict[str, Any]],
    mapping: Optional[Dict[str, Any]] = None,
) -> str:
    book = row if isinstance(row, dict) else {}
    explicit = str(book.get("reading_mode") or "").strip().casefold()
    if explicit in {"book", "manga"}:
        return explicit

    normalized_mapping = normalize_book_modes(mapping or {})
    asin = _normalize_mode_asin(book.get("asin"))
    if asin and normalized_mapping["by_asin"].get(asin) in {"book", "manga"}:
        return normalized_mapping["by_asin"][asin]
    title_key = normalized_title(book.get("title"))
    if title_key and normalized_mapping["by_title"].get(title_key) in {
        "book",
        "manga",
    }:
        return normalized_mapping["by_title"][title_key]
    series_key = normalized_series_title(book.get("title"))
    if series_key and normalized_mapping["by_series"].get(series_key) in {
        "book",
        "manga",
    }:
        return normalized_mapping["by_series"][series_key]
    return "book"


def is_manga(
    row: Optional[Dict[str, Any]],
    mapping: Optional[Dict[str, Any]] = None,
) -> bool:
    return reading_mode_for_book(row, mapping) == "manga"


def mapped_reading_mode_for_book(
    row: Optional[Dict[str, Any]],
    mapping: Optional[Dict[str, Any]] = None,
) -> Optional[str]:
    """Return only an explicit mapping, never the implicit book default."""

    book = row if isinstance(row, dict) else {}
    normalized_mapping = normalize_book_modes(mapping or {})
    asin = _normalize_mode_asin(book.get("asin"))
    if asin and normalized_mapping["by_asin"].get(asin) in {"book", "manga"}:
        return normalized_mapping["by_asin"][asin]
    title_key = normalized_title(book.get("title"))
    if title_key and normalized_mapping["by_title"].get(title_key) in {
        "book",
        "manga",
    }:
        return normalized_mapping["by_title"][title_key]
    series_key = normalized_series_title(book.get("title"))
    if series_key and normalized_mapping["by_series"].get(series_key) in {
        "book",
        "manga",
    }:
        return normalized_mapping["by_series"][series_key]
    return None


def series_key_for_book(row: Optional[Dict[str, Any]]) -> str:
    book = row if isinstance(row, dict) else {}
    explicit = str(book.get("series_key") or "").strip()
    if explicit:
        return normalized_title(explicit)
    return normalized_series_title(book.get("title"))


def manga_total_pages(row: Dict[str, Any]) -> Optional[int]:
    return parse_int(row.get("manga_pages_total"))


def manga_progress_percent(row: Dict[str, Any]) -> Optional[float]:
    progress = parse_float(row.get("progress_percent"))
    if progress is None:
        return None
    return round(max(0.0, min(100.0, progress)), 1)


def manga_pages_read(row: Dict[str, Any]) -> Optional[int]:
    explicit = parse_int(row.get("manga_pages_read"))
    if explicit is not None:
        return max(0, explicit)
    total = manga_total_pages(row)
    progress = manga_progress_percent(row)
    if total is None or progress is None:
        return None
    return max(0, min(total, int(round(total * progress / 100.0))))


def display_progress(row: Dict[str, Any]) -> Optional[float]:
    if is_manga(row):
        return manga_progress_percent(row)
    return word_progress(row)


def display_progress_label(row: Dict[str, Any]) -> str:
    value = display_progress(row)
    if value is not None:
        return f"{value:.1f}%"
    if is_manga(row):
        return "--"
    return "--"


def reading_state_label(row: Dict[str, Any]) -> Tuple[str, str]:
    state = reading_state(row)
    if state == "unread":
        return ("未开始", "Not started")
    if state == "finished":
        return ("已读完", "Finished")
    if state == "reading":
        return ("阅读中", "Reading")
    return ("无进度数据", "No progress data")


def reading_mode_label(value: Any) -> str:
    return "漫画" if normalize_reading_mode(value) == "manga" else "文字书"


def reading_mode_label_en(value: Any) -> str:
    return "Manga" if normalize_reading_mode(value) == "manga" else "Book"


def normalized_completed_threshold(value: Any = None) -> float:
    number = parse_float(value)
    if number is None:
        number = DEFAULT_COMPLETED_BOOK_THRESHOLD_PERCENT
    return round(max(0.0, min(100.0, number)), 2)


def is_completed_book(
    row: Optional[Dict[str, Any]],
    threshold_percent: float = DEFAULT_COMPLETED_BOOK_THRESHOLD_PERCENT,
) -> bool:
    """A finished book must have a real reading session and enough progress."""

    book = row if isinstance(row, dict) else {}
    if is_manga(book):
        # Manga has its own optional volume track and must not also light up
        # the completed-books track.
        return False
    total_time_ms = parse_int(book.get("total_time_ms")) or 0
    progress = parse_float(book.get("progress_percent"))
    threshold = normalized_completed_threshold(threshold_percent)
    return (
        total_time_ms > 0
        and progress is not None
        and progress >= threshold
    )


def is_completed_manga_volume(
    row: Optional[Dict[str, Any]],
    threshold_percent: float = DEFAULT_COMPLETED_BOOK_THRESHOLD_PERCENT,
) -> bool:
    book = row if isinstance(row, dict) else {}
    if not is_manga(book):
        return False
    progress = manga_progress_percent(book)
    threshold = normalized_completed_threshold(threshold_percent)
    return progress is not None and progress >= threshold


def has_reading_record(row: Optional[Dict[str, Any]]) -> bool:
    book = row if isinstance(row, dict) else {}
    total_time_ms = parse_int(book.get("total_time_ms")) or 0
    total_words = parse_int(book.get("total_words")) or 0
    return total_time_ms > 0 or total_words > 0


def display_title(title: str) -> str:
    cleaned = re.sub(r"_[0-9A-Fa-f]{24,}$", "", title.strip())
    return cleaned or title.strip() or "(unknown)"


def bilingual_text(zh: str, en: str) -> str:
    return (
        f'<span class="i18n-zh">{html.escape(str(zh))}</span>'
        f'<span class="i18n-en">{html.escape(str(en))}</span>'
    )


def format_duration(milliseconds: Optional[int]) -> str:
    if milliseconds is None:
        return "--"
    seconds = max(0, int(round(milliseconds / 1000.0)))
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    if hours:
        return f"{hours} 小时 {minutes} 分"
    if minutes:
        return f"{minutes} 分 {seconds} 秒"
    return f"{seconds} 秒"


def format_duration_en(milliseconds: Optional[int]) -> str:
    if milliseconds is None:
        return "--"
    seconds = max(0, int(round(milliseconds / 1000.0)))
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    if hours:
        return f"{hours} h {minutes} min"
    if minutes:
        return f"{minutes} min {seconds} s"
    return f"{seconds} s"


def _period_compact_duration(milliseconds: Optional[int]) -> str:
    seconds = max(0.0, (milliseconds or 0) / 1000.0)
    if seconds <= 0:
        return "0"
    if seconds < 60:
        return f"{seconds:.0f}s"
    minutes = seconds / 60.0
    if minutes < 60:
        return f"{minutes:.0f}m" if minutes >= 10 else f"{minutes:.1f}m"
    hours = minutes / 60.0
    return f"{hours:.0f}h" if hours >= 10 else f"{hours:.1f}h"


def format_words(value: Optional[int]) -> str:
    return f"{value:,}" if value is not None else "--"


def format_signed_duration(milliseconds: Optional[int]) -> str:
    if milliseconds is None:
        return "--"
    if milliseconds == 0:
        return "0 秒"
    prefix = "+" if milliseconds > 0 else "-"
    return prefix + format_duration(abs(milliseconds))


def format_signed_duration_en(milliseconds: Optional[int]) -> str:
    if milliseconds is None:
        return "--"
    if milliseconds == 0:
        return "0 s"
    prefix = "+" if milliseconds > 0 else "-"
    return prefix + format_duration_en(abs(milliseconds))


def format_signed_number(
    value: Optional[float],
    decimals: int = 0,
    suffix: str = "",
) -> str:
    if value is None:
        return "--"
    if abs(value) < 0.05:
        value = 0.0
    prefix = "+" if value > 0 else ""
    return f"{prefix}{value:,.{decimals}f}{suffix}"


def format_last_read(value: str) -> str:
    if not value:
        return "--"
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return parsed.strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return value.replace("T", " ")[:16]


def _row_ocr_min_confidence(row: Dict[str, Any]) -> float:
    configured = parse_float(row.get("ocr_min_confidence"))
    if configured is None:
        configured = DEFAULT_OCR_MIN_CONFIDENCE
    return max(0.0, min(1.0, configured))


def ocr_estimate_eligible(
    row: Dict[str, Any],
    min_confidence: Optional[float] = None,
) -> bool:
    """Return True only for complete OCR estimates that may enter totals."""

    if not is_manga(row):
        return False
    if str(row.get("ocr_status") or "").strip().casefold() != "ok":
        return False
    source = str(row.get("ocr_source") or "ocr-estimate").casefold()
    if source not in {
        "ocr-estimate",
        "ocr-sampled",
        "ocr-series-estimate",
    }:
        return False
    coverage = parse_float(row.get("ocr_page_coverage"))
    if source == "ocr-estimate" and (coverage is None or coverage < 98.0):
        return False

    pages_processed = parse_int(row.get("ocr_pages_processed"))
    pages_total = parse_int(row.get("ocr_pages_total"))
    word_count = parse_int(row.get("ocr_word_count"))
    char_count = parse_int(row.get("ocr_char_count"))
    confidence = parse_float(row.get("ocr_confidence"))
    if (
        pages_processed is None
        or pages_total is None
        or pages_total <= 0
        or word_count is None
    ):
        return False
    if source == "ocr-estimate" and (char_count is None or confidence is None):
        return False
    if source == "ocr-estimate" and (
        pages_processed <= 0 or pages_processed < pages_total
    ):
        return False
    if source != "ocr-estimate" and not (
        parse_int(row.get("ocr_sampled_pages")) or 0
    ) and not bool(row.get("ocr_extrapolated")):
        # Series-profile reuse can have zero freshly sampled pages, but must
        # explicitly declare that it extrapolated from a prior sample.
        return False

    threshold = (
        _row_ocr_min_confidence(row)
        if min_confidence is None
        else max(0.0, min(1.0, float(min_confidence)))
    )
    return confidence >= threshold - 1e-9


def effective_words_read(
    row: Dict[str, Any],
    min_confidence: Optional[float] = None,
) -> Optional[float]:
    if is_manga(row):
        if not ocr_estimate_eligible(row, min_confidence):
            return None
        word_count = parse_float(row.get("ocr_word_count"))
        progress = parse_float(row.get("progress_percent"))
        if word_count is None or progress is None:
            return None
        return round(word_count * max(0.0, min(100.0, progress)) / 100.0, 2)

    total_words = parse_float(row.get("total_words"))
    if total_words is None or total_words < 0:
        return None
    return total_words


def effective_wpm(
    row: Dict[str, Any],
    min_confidence: Optional[float] = None,
) -> Optional[float]:
    total_time_ms = parse_int(row.get("total_time_ms")) or 0
    minutes = total_time_ms / 60000.0
    if is_manga(row):
        if not ocr_estimate_eligible(row, min_confidence) or minutes <= 0:
            return None
        words_read = effective_words_read(row, min_confidence)
        if words_read is None:
            return None
        return round(words_read / minutes, 2)

    recorded = parse_float(row.get("wpm"))
    if recorded is not None and recorded > 0:
        return recorded
    words_read = effective_words_read(row, min_confidence)
    if words_read is None or minutes <= 0:
        return None
    return round(words_read / minutes, 2)


def effective_source(row: Dict[str, Any]) -> str:
    if is_manga(row):
        if ocr_estimate_eligible(row):
            return str(row.get("ocr_source") or "ocr-estimate")
        return ""
    if has_exact_word_count(row):
        return "exact-body"
    if effective_book_words(row) is not None:
        return "kindle-estimate"
    return "kindle-raw"


def word_progress(row: Dict[str, Any]) -> Optional[float]:
    words_read = effective_words_read(row)
    book_words = effective_book_words(row)
    if (
        words_read is None
        or book_words is None
        or book_words <= 0
    ):
        return None
    return round(
        max(0.0, min(100.0, words_read / book_words * 100.0)),
        1,
    )


def effective_book_words(row: Dict[str, Any]) -> Optional[float]:
    if is_manga(row):
        if not ocr_estimate_eligible(row):
            return None
        return parse_float(row.get("ocr_word_count"))
    exact = parse_float(row.get("book_words_exact"))
    if exact is not None and exact > 0:
        return exact
    estimated = parse_float(row.get("book_words"))
    if estimated is not None and estimated > 0:
        return estimated
    return None


def has_exact_word_count(row: Dict[str, Any]) -> bool:
    if is_manga(row):
        return False
    return (parse_int(row.get("book_words_exact")) or 0) > 0


def book_word_origin(row: Dict[str, Any]) -> str:
    if is_manga(row):
        return (
            "ocr-estimate"
            if ocr_estimate_eligible(row)
            else "not-applicable"
        )
    if has_exact_word_count(row):
        return "exact"
    estimated = parse_float(row.get("book_words"))
    if estimated is not None and estimated > 0:
        return "estimate"
    return "unknown"


def book_word_origin_label(row: Dict[str, Any]) -> str:
    if is_manga(row) and ocr_estimate_eligible(row):
        source = str(row.get("ocr_source") or "ocr-estimate")
        return {
            "ocr-sampled": "OCR 采样估算",
            "ocr-series-estimate": "同系列采样外推",
        }.get(source, "OCR 估算")
    return {
        "exact": "正文精确",
        "estimate": "Kindle 外推",
        "ocr-estimate": "OCR 估算",
        "unknown": "词数未知",
        "not-applicable": "漫画不适用",
    }[book_word_origin(row)]


def book_word_origin_label_en(row: Dict[str, Any]) -> str:
    if is_manga(row) and ocr_estimate_eligible(row):
        source = str(row.get("ocr_source") or "ocr-estimate")
        return {
            "ocr-sampled": "OCR sampled estimate",
            "ocr-series-estimate": "Series sample estimate",
        }.get(source, "OCR estimate")
    return {
        "exact": "Exact body",
        "estimate": "Kindle estimate",
        "ocr-estimate": "OCR estimate",
        "unknown": "Word count unknown",
        "not-applicable": "Not applicable to manga",
    }[book_word_origin(row)]


def estimate_deviation_percent(
    row: Dict[str, Any],
) -> Optional[float]:
    if is_manga(row):
        return None
    exact = parse_float(row.get("book_words_exact"))
    estimated = parse_float(row.get("book_words"))
    if exact is None or exact <= 0:
        return None
    if estimated is None or estimated <= 0:
        return None
    return round((estimated - exact) / exact * 100.0, 1)


def _normal_wpm_range(row: Dict[str, Any]) -> Tuple[float, float]:
    language = str(
        row.get("word_count_language")
        or row.get("ocr_language")
        or ""
    ).lower()
    return WPM_NORMAL_RANGES.get(language, DEFAULT_WPM_NORMAL_RANGE)


def wpm_is_abnormal(row: Dict[str, Any]) -> bool:
    value = effective_wpm(row)
    total_time_ms = parse_int(row.get("total_time_ms")) or 0
    total_words = parse_int(effective_words_read(row)) or 0
    if is_manga(row) and not ocr_estimate_eligible(row):
        return False
    if value is None or value <= 0:
        return False
    if (
        total_time_ms < QUALITY_MIN_TIME_MS
        or total_words < QUALITY_MIN_WORDS
    ):
        return False
    low, high = _normal_wpm_range(row)
    return value < low or value > high


def progress_is_abnormal(row: Dict[str, Any]) -> bool:
    progress = parse_float(row.get("progress_percent"))
    if progress is None:
        return False
    if progress < 0.0 or progress > 100.0:
        return True
    if is_manga(row):
        return False

    total_words = parse_float(row.get("total_words"))
    exact = parse_float(row.get("book_words_exact"))
    if total_words is None or exact is None or exact <= 0:
        return False
    exact_progress = total_words / exact * 100.0
    return abs(progress - exact_progress) > PROGRESS_DISCREPANCY_PP


def analyze_book_quality(row: Dict[str, Any]) -> Dict[str, Any]:
    if is_manga(row) and not ocr_estimate_eligible(row):
        return {
            "exact_available": False,
            "estimate_available": False,
            "word_applicable": False,
            "reading_mode": "manga",
            "word_origin": "not-applicable",
            "word_origin_label": "漫画不适用",
            "word_origin_label_en": "Not applicable to manga",
            "sample_sufficient": False,
            "deviation_percent": None,
            "deviation_abs_percent": None,
            "deviation_level": "not-applicable",
            "wpm_value": None,
            "wpm_available": False,
            "abnormal_wpm": False,
            "abnormal_progress": progress_is_abnormal(row),
            "confidence_score": None,
            "confidence_level": "not-applicable",
            "confidence_label": "不适用",
            "confidence_label_en": "Not applicable",
            "flags": ["漫画模式不要求词数"],
            "flags_en": ["Word count is not applicable in manga mode"],
        }
    if is_manga(row):
        total_time_ms = parse_int(row.get("total_time_ms")) or 0
        words_read = parse_int(effective_words_read(row)) or 0
        sample_sufficient = (
            total_time_ms >= QUALITY_MIN_TIME_MS
            and words_read >= QUALITY_MIN_WORDS
        )
        wpm = effective_wpm(row)
        wpm_available = wpm is not None and wpm > 0
        confidence = parse_float(row.get("ocr_confidence")) or 0.0
        score = int(round(max(0.0, min(1.0, confidence)) * 100.0))
        if score >= 80:
            level, label, label_en = "high", "高置信", "High confidence"
        elif score >= 60:
            level, label, label_en = "medium", "中置信", "Medium confidence"
        else:
            level, label, label_en = "low", "低置信", "Low confidence"
        flags = ["OCR 估算已纳入有效词数"]
        flags_en = ["OCR estimate is included in effective words"]
        if not sample_sufficient:
            flags.append("阅读样本偏少")
            flags_en.append("Small reading sample")
        return {
            "exact_available": False,
            "estimate_available": True,
            "word_applicable": True,
            "reading_mode": "manga",
            "word_origin": "ocr-estimate",
            "word_origin_label": "OCR 估算",
            "word_origin_label_en": "OCR estimate",
            "sample_sufficient": sample_sufficient,
            "deviation_percent": None,
            "deviation_abs_percent": None,
            "deviation_level": "not-applicable",
            "wpm_value": wpm,
            "wpm_available": wpm_available,
            "abnormal_wpm": wpm_is_abnormal(row),
            "abnormal_progress": progress_is_abnormal(row),
            "confidence_score": score,
            "confidence_level": level,
            "confidence_label": label,
            "confidence_label_en": label_en,
            "flags": flags,
            "flags_en": flags_en,
        }
    exact_available = has_exact_word_count(row)
    estimated = parse_float(row.get("book_words"))
    estimate_available = estimated is not None and estimated > 0
    total_time_ms = parse_int(row.get("total_time_ms")) or 0
    total_words = parse_int(row.get("total_words")) or 0
    sample_sufficient = (
        total_time_ms >= QUALITY_MIN_TIME_MS
        and total_words >= QUALITY_MIN_WORDS
    )
    wpm = parse_float(row.get("wpm"))
    wpm_available = wpm is not None and wpm > 0
    abnormal_wpm = wpm_is_abnormal(row)
    abnormal_progress = progress_is_abnormal(row)
    deviation = estimate_deviation_percent(row)
    deviation_abs = abs(deviation) if deviation is not None else None

    score = 40 if exact_available else 0
    score += 20 if sample_sufficient else 0

    if deviation_abs is None:
        score += 5
        deviation_level = "missing"
    elif deviation_abs <= 10.0:
        score += 15
        deviation_level = "good"
    elif deviation_abs <= 25.0:
        score += 10
        deviation_level = "medium"
    elif deviation_abs <= 50.0:
        score += 4
        deviation_level = "high"
    else:
        deviation_level = "very-high"

    if not wpm_available:
        score += 5
    elif not sample_sufficient:
        score += 8
    elif abnormal_wpm:
        score += 0
    else:
        score += 15

    progress = parse_float(row.get("progress_percent"))
    if progress is None:
        score += 5
    elif abnormal_progress:
        score += 0
    else:
        score += 10

    if deviation_abs is not None and deviation_abs > 50.0:
        score = min(score, 59)
    if abnormal_wpm or abnormal_progress:
        score = min(score, 55)

    score = max(0, min(100, score))
    if score >= 80:
        level = "high"
        label = "高置信"
        label_en = "High confidence"
    elif score >= 60:
        level = "medium"
        label = "中置信"
        label_en = "Medium confidence"
    else:
        level = "low"
        label = "低置信"
        label_en = "Low confidence"

    flags: List[str] = []
    flags_en: List[str] = []
    if not exact_available:
        flags.append("缺少正文精确词数")
        flags_en.append("Missing exact body word count")
    if not sample_sufficient:
        flags.append("阅读样本偏少")
        flags_en.append("Small reading sample")
    if deviation_abs is not None and deviation_abs > 25.0:
        flags.append("外推偏差大")
        flags_en.append("Large estimate deviation")
    if abnormal_wpm:
        flags.append("WPM 异常")
        flags_en.append("Abnormal WPM")
    if abnormal_progress:
        flags.append("进度异常")
        flags_en.append("Abnormal progress")
    if not flags:
        flags.append("关键指标完整")
        flags_en.append("Key metrics complete")

    return {
        "exact_available": exact_available,
        "estimate_available": estimate_available,
        "word_applicable": True,
        "reading_mode": reading_mode_for_book(row),
        "word_origin": book_word_origin(row),
        "word_origin_label": book_word_origin_label(row),
        "word_origin_label_en": book_word_origin_label_en(row),
        "sample_sufficient": sample_sufficient,
        "deviation_percent": deviation,
        "deviation_abs_percent": deviation_abs,
        "deviation_level": deviation_level,
        "wpm_value": wpm,
        "wpm_available": wpm_available,
        "abnormal_wpm": abnormal_wpm,
        "abnormal_progress": abnormal_progress,
        "confidence_score": score,
        "confidence_level": level,
        "confidence_label": label,
        "confidence_label_en": label_en,
        "flags": flags,
        "flags_en": flags_en,
    }


def build_quality_summary(
    rows: Sequence[Dict[str, Any]],
) -> Dict[str, Any]:
    qualities = [analyze_book_quality(row) for row in rows]
    word_qualities = [
        quality for quality in qualities if quality.get("word_applicable")
    ]
    deviations = sorted(
        quality["deviation_abs_percent"]
        for quality in word_qualities
        if quality["deviation_abs_percent"] is not None
    )
    if deviations:
        middle = len(deviations) // 2
        if len(deviations) % 2:
            median_deviation = deviations[middle]
        else:
            median_deviation = (
                deviations[middle - 1] + deviations[middle]
            ) / 2.0
        max_deviation = deviations[-1]
    else:
        median_deviation = None
        max_deviation = None

    exact_count = sum(
        1 for quality in word_qualities if quality["exact_available"]
    )
    book_count = len(rows)
    word_book_count = len(word_qualities)
    kindle_word_book_count = sum(1 for row in rows if not is_manga(row))
    manga_book_count = sum(1 for row in rows if is_manga(row))
    ocr_estimate_count = sum(
        1 for row in rows if ocr_estimate_eligible(row)
    )
    exact_coverage = (
        round(exact_count / kindle_word_book_count * 100.0, 1)
        if kindle_word_book_count
        else None
    )
    return {
        "books": book_count,
        "word_books": word_book_count,
        "kindle_word_books": kindle_word_book_count,
        "manga_books": manga_book_count,
        "ocr_estimate_books": ocr_estimate_count,
        "exact_count": exact_count,
        "exact_coverage_percent": exact_coverage,
        "deviation_count": len(deviations),
        "median_deviation_abs": (
            round(median_deviation, 1)
            if median_deviation is not None
            else None
        ),
        "max_deviation_abs": (
            round(max_deviation, 1)
            if max_deviation is not None
            else None
        ),
        "abnormal_wpm_count": sum(
            1 for quality in word_qualities if quality["abnormal_wpm"]
        ),
        "abnormal_progress_count": sum(
            1 for quality in qualities if quality["abnormal_progress"]
        ),
        "low_confidence_count": sum(
            1
            for quality in word_qualities
            if quality["confidence_level"] == "low"
        ),
        "qualities": qualities,
    }


def _quality_card(
    label: str,
    value: str,
    detail: str,
    tone: str,
    label_en: Optional[str] = None,
    detail_en: Optional[str] = None,
) -> str:
    return (
        f'<article class="quality-card quality-{html.escape(tone, quote=True)}">'
        f"<span>{bilingual_text(label, label_en or label)}</span>"
        f"<strong>{html.escape(value)}</strong>"
        f"<small>{bilingual_text(detail, detail_en or detail)}</small>"
        "</article>"
    )


def build_quality_cards(quality: Dict[str, Any]) -> str:
    books = parse_int(quality.get("books")) or 0
    word_books = parse_int(quality.get("word_books"))
    if word_books is None:
        word_books = books
    kindle_word_books = parse_int(quality.get("kindle_word_books"))
    if kindle_word_books is None:
        kindle_word_books = word_books
    manga_books = parse_int(quality.get("manga_books")) or 0
    exact_count = parse_int(quality.get("exact_count")) or 0
    exact_coverage = parse_float(quality.get("exact_coverage_percent"))
    median_deviation = parse_float(quality.get("median_deviation_abs"))
    max_deviation = parse_float(quality.get("max_deviation_abs"))

    if kindle_word_books:
        coverage_text = f"{exact_count}/{kindle_word_books}"
        coverage_detail = (
            f"{exact_coverage:.1f}% 文字书"
            if exact_coverage is not None
            else "无可用文字书"
        )
        coverage_detail_en = (
            f"{exact_coverage:.1f}% of word-based books"
            if exact_coverage is not None
            else "No usable word-based books"
        )
    else:
        coverage_text = "--"
        coverage_detail = "暂无数据"
        coverage_detail_en = "No data"

    if median_deviation is None:
        deviation_text = "--"
        deviation_detail = "需要精确词数与 Kindle 外推值"
        deviation_detail_en = "Requires exact and Kindle estimate values"
    else:
        deviation_text = f"{median_deviation:.1f}%"
        deviation_detail = (
            f"最大 {max_deviation:.1f}%"
            if max_deviation is not None
            else "无最大偏差"
        )
        deviation_detail_en = (
            f"Max {max_deviation:.1f}%"
            if max_deviation is not None
            else "No maximum deviation"
        )

    abnormal_wpm = parse_int(quality.get("abnormal_wpm_count")) or 0
    abnormal_progress = (
        parse_int(quality.get("abnormal_progress_count")) or 0
    )
    low_confidence = parse_int(quality.get("low_confidence_count")) or 0
    return "".join(
        (
            _quality_card(
                "精确词数覆盖",
                coverage_text,
                coverage_detail,
                "good" if exact_coverage == 100.0 else "neutral",
                "Exact word-count coverage",
                coverage_detail_en,
            ),
            _quality_card(
                "外推偏差",
                deviation_text,
                deviation_detail,
                (
                    "good"
                    if median_deviation is not None
                    and median_deviation <= 10.0
                    else "warn"
                ),
                "Estimate deviation",
                deviation_detail_en,
            ),
            _quality_card(
                "异常 WPM",
                str(abnormal_wpm),
                "按语言与最小时长筛选",
                "bad" if abnormal_wpm else "good",
                "Abnormal WPM",
                "Filtered by language and minimum duration",
            ),
            _quality_card(
                "异常进度",
                str(abnormal_progress),
                "Kindle 进度与词数进度比较",
                "bad" if abnormal_progress else "good",
                "Abnormal progress",
                "Kindle progress compared with word-count progress",
            ),
            _quality_card(
                "低置信书籍",
                str(low_confidence),
                (
                    f"共 {word_books} 本参与词数"
                    + (
                        f" · {parse_int(quality.get('ocr_estimate_books')) or 0} "
                        "本 OCR 估算"
                        if parse_int(quality.get("ocr_estimate_books"))
                        else ""
                    )
                    + (f" · {manga_books} 本漫画" if manga_books else "")
                ),
                "warn" if low_confidence else "good",
                "Low-confidence books",
                (
                    f"{word_books} word-based books"
                    + (
                        f" · {parse_int(quality.get('ocr_estimate_books')) or 0} "
                        "OCR estimates"
                        if parse_int(quality.get("ocr_estimate_books"))
                        else ""
                    )
                    + (f" · {manga_books} manga" if manga_books else "")
                ),
            ),
        )
    )


def ocr_estimate_available(row: Dict[str, Any]) -> bool:
    status = str(row.get("ocr_status") or "").strip()
    word_count = parse_int(row.get("ocr_word_count"))
    char_count = parse_int(row.get("ocr_char_count"))
    return bool(status) or word_count is not None or char_count is not None


def _word_source_tag_keys(row: Dict[str, Any]) -> List[str]:
    keys: List[str] = []
    if parse_int(row.get("total_time_ms")) or parse_int(row.get("total_words")):
        keys.append("kindle-raw")
    if has_exact_word_count(row):
        keys.append("exact-body")
    elif not is_manga(row):
        estimated = parse_float(row.get("book_words"))
        if estimated is not None and estimated > 0:
            keys.append("kindle-estimate")
    if row.get("_fmcache_fallback") or str(
        row.get("fmcache_fallback") or ""
    ).strip().casefold() in {"1", "true", "yes"}:
        keys.append("fmcache-fallback")
    if ocr_estimate_available(row):
        keys.append("ocr-estimate")
    if is_manga(row) and not keys:
        keys.append("manga-not-applicable")
    return keys


def _word_source_tag_label(key: str) -> str:
    return {
        "kindle-raw": "Kindle 原始",
        "exact-body": "正文精确",
        "kindle-estimate": "Kindle 外推",
        "fmcache-fallback": "fmcache 回退",
        "ocr-estimate": "OCR 估算",
        "manga-not-applicable": "漫画不适用",
    }.get(key, key)


def _word_source_tag_label_en(key: str) -> str:
    return {
        "kindle-raw": "Kindle raw",
        "exact-body": "Exact body",
        "kindle-estimate": "Kindle estimate",
        "fmcache-fallback": "fmcache fallback",
        "ocr-estimate": "OCR estimate",
        "manga-not-applicable": "Not applicable to manga",
    }.get(key, key)


def _word_source_tag_html(row: Dict[str, Any]) -> str:
    return "".join(
        (
            f'<span class="tag source-tag source-{html.escape(key, quote=True)}">'
            f"{bilingual_text(_word_source_tag_label(key), _word_source_tag_label_en(key))}"
            "</span>"
        )
        for key in _word_source_tag_keys(row)
    )


def _ocr_status_label(value: Any) -> Tuple[str, str]:
    normalized = str(value or "").strip().casefold()
    return {
        "ok": ("完成", "Complete"),
        "partial": ("部分页面", "Partial pages"),
        "ocr-low-confidence": ("低置信", "Low confidence"),
        "needs-language": ("需指定语言", "Language required"),
        "engine-missing": ("缺少 OCR 引擎", "OCR engine missing"),
        "render-missing": ("缺少渲染依赖", "Renderer missing"),
        "render-failed": ("渲染失败", "Render failed"),
        "source-missing": ("源文件缺失", "Source missing"),
        "source-unreadable": ("源文件不可读", "Source unreadable"),
        "engine-failed": ("OCR 引擎失败", "OCR engine failed"),
        "failed": ("失败", "Failed"),
        "not-applicable": ("不适用", "Not applicable"),
    }.get(
        normalized,
        (normalized or "未知", normalized or "Unknown"),
    )


def build_ocr_section(rows: Sequence[Dict[str, Any]]) -> str:
    """Render OCR estimates separately from exact Kindle word metrics."""

    ocr_rows = [row for row in rows if ocr_estimate_available(row)]
    if not ocr_rows:
        return ""
    complete = sum(
        1 for row in ocr_rows if str(row.get("ocr_status") or "") == "ok"
    )
    low_confidence = sum(
        1
        for row in ocr_rows
        if str(row.get("ocr_status") or "") == "ocr-low-confidence"
    )
    needs_language = sum(
        1
        for row in ocr_rows
        if str(row.get("ocr_status") or "") == "needs-language"
    )
    dependency_missing = sum(
        1
        for row in ocr_rows
        if str(row.get("ocr_status") or "") in {
            "engine-missing",
            "render-missing",
        }
    )
    eligible = sum(1 for row in ocr_rows if ocr_estimate_eligible(row))
    cards = "".join(
        (
            _quality_card(
                "OCR 完成",
                f"{complete}/{len(ocr_rows)}",
                f"其中 {eligible} 本达到并入门槛",
                "good" if complete == len(ocr_rows) else "neutral",
                "OCR complete",
                f"{eligible} row(s) meet the inclusion gate",
            ),
            _quality_card(
                "OCR 低置信",
                str(low_confidence),
                f"阈值 {kindle_manga_ocr.OCR_DEFAULT_MIN_CONFIDENCE:.2f} 由运行参数覆盖"
                if kindle_manga_ocr is not None
                else "按运行参数阈值统计",
                "warn" if low_confidence else "good",
                "OCR low confidence",
                "Threshold supplied by the run",
            ),
            _quality_card(
                "需指定语言",
                str(needs_language),
                "auto 判断不确定",
                "warn" if needs_language else "good",
                "Language required",
                "Auto detection was not confident",
            ),
            _quality_card(
                "依赖/渲染缺失",
                str(dependency_missing),
                "未自动安装或联网下载",
                "bad" if dependency_missing else "good",
                "Missing dependency/renderer",
                "Nothing was installed or downloaded automatically",
            ),
        )
    )
    body_rows: List[str] = []
    for row in ocr_rows:
        status = str(row.get("ocr_status") or "")
        status_zh, status_en = _ocr_status_label(status)
        eligible_text = (
            ("计入正式统计", "Included in formal stats")
            if ocr_estimate_eligible(row)
            else ("不并入正式统计", "Excluded from formal stats")
        )
        confidence = parse_float(row.get("ocr_confidence"))
        coverage = parse_float(row.get("ocr_page_coverage"))
        completed = format_last_read(str(row.get("ocr_completed_at") or ""))
        error = str(row.get("ocr_error") or "").strip() or "--"
        confidence_text = (
            f"{confidence * 100:.1f}%" if confidence is not None else "--"
        )
        coverage_text = (
            f"{coverage:.2f}%" if coverage is not None else "--"
        )
        body_rows.append(
            "<tr "
            f'data-ocr-status="{html.escape(status, quote=True)}" '
            f'data-ocr-language="{html.escape(str(row.get("ocr_language") or ""), quote=True)}">'
            f'<td class="book-cell"><strong>{html.escape(display_title(str(row.get("title") or "(unknown)")))}</strong>'
            f'<span>{html.escape(str(row.get("author") or "--"))}</span></td>'
            f"<td>{html.escape(format_words(parse_int(row.get('ocr_word_count'))))}</td>"
            f"<td>{html.escape(format_words(parse_int(row.get('ocr_char_count'))))}</td>"
            f"<td>{html.escape(confidence_text)}</td>"
            f"<td>{html.escape(str(row.get('ocr_engine') or '--'))}</td>"
            f"<td>{html.escape(language_label(row.get('ocr_language')))}</td>"
            f"<td>{html.escape(coverage_text)}</td>"
            f'<td><span class="confidence-tag">{bilingual_text(status_zh, status_en)}</span></td>'
            f'<td><span class="confidence-tag">{bilingual_text(*eligible_text)}</span></td>'
            f"<td>{html.escape(completed)}</td>"
            f"<td>{html.escape(error)}</td>"
            f"<td><code>{html.escape(str(row.get('ocr_cache_key') or '--'))}</code></td>"
            "</tr>"
        )
    return (
        '<details class="panel ocr-panel" id="ocr-panel" open '
        'data-ocr-row-count="' + str(len(ocr_rows)) + '">'
        "<summary>"
        '<div><h2><span class="i18n-zh">漫画 OCR 估算</span>'
        '<span class="i18n-en">Manga OCR estimates</span></h2>'
        "<p>"
        '<span class="i18n-zh">仅作为离线估算，绝不写入正文精确词数。只有状态 ok、覆盖率至少 98%、页数/词数/字符数/置信度完整且置信度达标的书，才并入有效词数、有效 WPM、趋势和周期报告。</span>'
        '<span class="i18n-en">Offline estimates never replace exact body counts. Only rows with status ok, at least 98% coverage, complete page/word/character/confidence fields, and passing confidence join effective words, effective WPM, trends, and period reports.</span>'
        "</p></div><em>"
        + bilingual_text(
            f"{len(ocr_rows)} 本漫画存在 OCR 状态",
            f"{len(ocr_rows)} manga rows have OCR status",
        )
        + "</em></summary>"
        '<div class="ocr-body">'
        f'<section class="quality-summary">{cards}</section>'
        '<div class="table-wrap"><table class="ocr-table">'
        "<thead><tr>"
        "<th>书名 / Book</th><th>词数 / Words</th><th>字符 / Chars</th>"
        "<th>置信度 / Confidence</th><th>引擎 / Engine</th>"
        "<th>语言 / Language</th><th>覆盖率 / Coverage</th>"
        "<th>状态 / Status</th><th>统计口径 / Inclusion</th><th>完成时间 / Completed</th>"
        "<th>错误 / Error</th><th>缓存键 / Cache key</th>"
        "</tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody></table></div>"
        "</div></details>"
    )


def normalized_title(title: Any) -> str:
    cleaned = display_title(str(title or ""))
    normalized = unicodedata.normalize("NFKC", cleaned).casefold()
    return re.sub(r"\s+", " ", normalized).strip()


_READING_MODE_VOLUME_SUFFIX = re.compile(
    r"[\s:：\-–—_]*("
    r"第\s*[0-9０-９一二三四五六七八九十百零两]+\s*[卷册集话話部]"
    r"|(?:vol(?:ume)?|book|part|no)\.?\s*[0-9０-９]+"
    r"|[0-9０-９]+\s*[卷册集话話部]"
    r"|#\s*[0-9０-９]+"
    r")\s*$",
    flags=re.IGNORECASE,
)


def normalized_series_title(title: Any) -> str:
    """Return a volume-agnostic title so a mapped volume covers its series."""

    text = normalized_title(title)
    previous = None
    while previous != text:
        previous = text
        text = _READING_MODE_VOLUME_SUFFIX.sub("", text).strip()
    return re.sub(r"\s+", " ", text).strip()


def stable_book_id(row: Dict[str, Any]) -> str:
    asin = unicodedata.normalize("NFKC", str(row.get("asin") or "").strip())
    compact_asin = re.sub(r"\s+", "", asin)
    if compact_asin and re.fullmatch(r"[A-Za-z0-9]{8,}", compact_asin):
        return "asin:" + compact_asin.upper()
    return "title:" + (normalized_title(row.get("title")) or "unknown")


def _find_ebook_meta(explicit: Optional[str] = None) -> Optional[str]:
    configured = (explicit or os.environ.get("CALIBRE_EBOOK_META", "")).strip()
    if configured:
        path = Path(configured).expanduser()
        if path.is_file():
            return str(path)

    candidates = [
        shutil.which("ebook-meta.exe"),
        shutil.which("ebook-meta"),
        r"C:\Program Files\Calibre2\ebook-meta.exe",
        r"C:\Program Files (x86)\Calibre2\ebook-meta.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return candidate
    return None


def _resolve_ebook_file(row: Dict[str, Any]) -> Optional[Path]:
    configured = str(row.get("ebook_file") or "").strip()
    if configured:
        path = Path(configured)
        if path.is_file() and path.suffix.lower() in EBOOK_EXTENSIONS:
            return path

    source_value = str(row.get("source_file") or "").strip()
    if not source_value:
        return None
    source = Path(source_value)
    directory = source.parent
    if directory.name.casefold().endswith(".sdr"):
        directory = directory.parent
    if not directory.is_dir():
        return None

    wanted = display_title(str(row.get("title") or "")).casefold()
    exact_matches: List[Path] = []
    prefix_matches: List[Path] = []
    try:
        candidates = list(directory.iterdir())
    except OSError:
        return None

    for candidate in candidates:
        if (
            not candidate.is_file()
            or candidate.suffix.lower() not in EBOOK_EXTENSIONS
        ):
            continue
        stem = candidate.stem.casefold()
        if wanted and stem == wanted:
            exact_matches.append(candidate)
        elif wanted and stem.startswith(wanted):
            prefix_matches.append(candidate)

    matches = exact_matches or prefix_matches
    return min(matches, key=lambda item: len(str(item))) if matches else None


def _cover_cache_path(
    covers_dir: Path,
    ebook_path: Path,
) -> Path:
    stat = ebook_path.stat()
    cache_key = "{}|{}|{}|{}".format(
        COVER_VERSION,
        ebook_path.resolve(),
        stat.st_size,
        stat.st_mtime_ns,
    )
    digest = hashlib.sha256(cache_key.encode("utf-8")).hexdigest()[:24]
    return covers_dir / f"cover-{digest}.jpg"


def _cover_is_loadable(path: Path) -> bool:
    if not path.is_file() or path.stat().st_size < 32:
        return False
    if Image is None:
        return path.suffix.lower() in {".jpg", ".jpeg", ".png", ".webp"}
    try:
        with Image.open(path) as image:
            image.verify()
        return True
    except (OSError, ValueError):
        return False


def _write_cover_thumbnail(source: Path, target: Path) -> None:
    if Image is None or ImageOps is None:
        raise RuntimeError("Pillow is required to create cover thumbnails")

    target.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        "wb",
        dir=target.parent,
        prefix=target.name + ".",
        suffix=".tmp",
        delete=False,
    ) as handle:
        temporary_path = Path(handle.name)

    try:
        with Image.open(source) as opened:
            image = ImageOps.exif_transpose(opened)
            if image.mode in {"RGBA", "LA"} or "transparency" in image.info:
                rgba = image.convert("RGBA")
                background = Image.new("RGB", rgba.size, "white")
                background.paste(rgba, mask=rgba.getchannel("A"))
                image = background
            else:
                image = image.convert("RGB")
            image.thumbnail(COVER_MAX_SIZE, Image.Resampling.LANCZOS)
            image.save(
                temporary_path,
                "JPEG",
                quality=88,
                optimize=True,
                progressive=True,
            )
        temporary_path.replace(target)
    finally:
        temporary_path.unlink(missing_ok=True)


def _extract_cover(
    ebook_meta: str,
    ebook_path: Path,
    covers_dir: Path,
) -> Path:
    handle, raw_name = tempfile.mkstemp(
        dir=covers_dir,
        prefix="cover-source-",
        suffix=".jpg",
    )
    os.close(handle)
    raw_path = Path(raw_name)
    raw_path.unlink(missing_ok=True)

    command = [ebook_meta, "--get-cover", str(raw_path), str(ebook_path)]
    kwargs: Dict[str, Any] = {}
    if os.name == "nt":
        kwargs["creationflags"] = 0x08000000  # CREATE_NO_WINDOW
    completed = None
    last_error: Optional[Exception] = None
    for attempt in range(1, 3):
        try:
            completed = subprocess.run(
                command,
                check=False,
                capture_output=True,
                timeout=120,
                **kwargs,
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            last_error = exc
            if attempt == 2:
                break
            time.sleep(0.5)
            continue
        if completed.returncode == 0:
            break
        last_error = RuntimeError(f"ebook-meta exited {completed.returncode}")
        if attempt == 2:
            break
        time.sleep(0.5)

    if completed is None or completed.returncode != 0:
        raw_path.unlink(missing_ok=True)
        if completed is not None:
            details = (completed.stderr or completed.stdout).decode(
                "utf-8", errors="replace"
            ).strip()
        else:
            details = str(last_error or "unknown error")
        raise RuntimeError(details)
    if not _cover_is_loadable(raw_path):
        raw_path.unlink(missing_ok=True)
        raise RuntimeError("ebook-meta did not produce a readable cover")
    return raw_path


def prepare_covers(
    rows: Sequence[Dict[str, Any]],
    covers_dir: Path,
    ebook_meta_path: Optional[str] = None,
) -> Tuple[Dict[str, Path], List[str]]:
    covers_dir.mkdir(parents=True, exist_ok=True)
    ebook_meta = _find_ebook_meta(ebook_meta_path)
    if not ebook_meta:
        return {}, ["Calibre ebook-meta was not found"]

    assets: Dict[str, Path] = {}
    errors: List[str] = []
    for row in rows:
        book_id = stable_book_id(row)
        title = display_title(str(row.get("title") or ""))
        ebook_path = _resolve_ebook_file(row)
        if ebook_path is None:
            errors.append(f"{title}: ebook file was not found")
            continue

        try:
            target = _cover_cache_path(covers_dir, ebook_path)
        except OSError as exc:
            errors.append(f"{title}: {exc}")
            continue

        if _cover_is_loadable(target):
            assets[book_id] = target
            continue

        raw_cover: Optional[Path] = None
        try:
            raw_cover = _extract_cover(ebook_meta, ebook_path, covers_dir)
            _write_cover_thumbnail(raw_cover, target)
            if not _cover_is_loadable(target):
                raise RuntimeError("generated thumbnail is not readable")
            assets[book_id] = target
        except (OSError, RuntimeError, ValueError) as exc:
            errors.append(f"{title}: {exc}")
        finally:
            if raw_cover is not None:
                raw_cover.unlink(missing_ok=True)

    return assets, errors


def cover_urls_for_html(
    covers: Mapping[str, Path],
    html_path: Path,
) -> Dict[str, str]:
    del html_path
    urls: Dict[str, str] = {}
    for book_id, cover_path in covers.items():
        try:
            payload = cover_path.read_bytes()
        except OSError:
            continue
        if not payload:
            continue
        mime = {
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".png": "image/png",
            ".webp": "image/webp",
        }.get(cover_path.suffix.lower(), "image/jpeg")
        encoded = base64.b64encode(payload).decode("ascii")
        urls[book_id] = f"data:{mime};base64,{encoded}"
    return urls


def existing_cover_urls_for_html(html_path: Path) -> Dict[str, str]:
    """Recover already-generated local covers before replacing the dashboard."""

    try:
        document = html_path.read_text(encoding="utf-8")
    except OSError:
        return {}
    urls: Dict[str, str] = {}
    pattern = re.compile(
        r'<article class="book-cover-card" data-book-id="([^"]+)".*?'
        r'<img data-cover-image src="([^"]+)"',
        flags=re.DOTALL,
    )
    for raw_book_id, raw_source in pattern.findall(document):
        book_id = html.unescape(raw_book_id)
        source = html.unescape(raw_source)
        if source.startswith("data:image/"):
            urls[book_id] = source
            continue
        if "://" in source or source.startswith("//"):
            continue
        source_path = (html_path.parent / source).resolve()
        try:
            payload = source_path.read_bytes()
        except OSError:
            continue
        if not payload:
            continue
        mime = {
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".png": "image/png",
            ".webp": "image/webp",
        }.get(source_path.suffix.lower(), "image/jpeg")
        encoded = base64.b64encode(payload).decode("ascii")
        urls[book_id] = f"data:{mime};base64,{encoded}"
    return urls


def normalize_language_code(value: Any) -> str:
    normalized = str(value or "").strip().casefold()
    return normalized if normalized in {"en", "zh", "ja"} else ""


def language_label(value: Any) -> str:
    return {
        "en": "英文",
        "zh": "中文",
        "ja": "日文",
    }.get(normalize_language_code(value), "未记录")


def language_label_en(value: Any) -> str:
    return {
        "en": "English",
        "zh": "Chinese",
        "ja": "Japanese",
    }.get(normalize_language_code(value), "Not recorded")


def inferred_language_from_source(value: Any) -> str:
    return {
        "en-regex": "en",
        "jieba": "zh",
        "janome": "ja",
    }.get(str(value or "").strip().casefold(), "")


def resolved_card_language(row: Optional[Dict[str, Any]]) -> str:
    """Resolve one language value for cards and table rows.

    Manga is decided first. OCR language wins over a valid body-word
    language, and the tokenizer source is only used when no explicit
    language field is available.
    """

    book = row if isinstance(row, dict) else {}
    word_language = normalize_language_code(book.get("word_count_language"))
    if is_manga(book):
        ocr_language = normalize_language_code(book.get("ocr_language"))
        if ocr_language:
            return ocr_language
        if word_language:
            return word_language
        return inferred_language_from_source(book.get("word_count_source"))
    if word_language:
        return word_language
    return inferred_language_from_source(book.get("word_count_source"))


def resolved_card_language_label(row: Optional[Dict[str, Any]]) -> str:
    language = resolved_card_language(row)
    if language:
        return language_label(language)
    return "漫画" if is_manga(row) else "未记录"


def resolved_card_language_label_en(row: Optional[Dict[str, Any]]) -> str:
    language = resolved_card_language(row)
    if language:
        return language_label_en(language)
    return "Manga" if is_manga(row) else "Not recorded"


TOKENIZER_LABELS = {
    "en-regex": ("英文正则", "English regex"),
    "jieba": ("jieba", "jieba"),
    "janome": ("Janome", "Janome"),
    "kindle-estimate": ("Kindle 外推", "Kindle estimate"),
    "ocr-estimate": ("OCR 估算", "OCR estimate"),
    "ocr-sampled": ("OCR 采样估算", "OCR sampled estimate"),
    "ocr-series-estimate": ("同系列采样外推", "Series sample estimate"),
    "not-applicable": ("漫画不适用", "Not applicable to manga"),
}


def tokenizer_label(value: Any) -> str:
    return TOKENIZER_LABELS.get(
        str(value or "").strip().casefold(),
        ("", ""),
    )[0]


def tokenizer_label_en(value: Any) -> str:
    return TOKENIZER_LABELS.get(
        str(value or "").strip().casefold(),
        ("", ""),
    )[1]


def resolved_tokenizer_source(row: Optional[Dict[str, Any]]) -> str:
    book = row if isinstance(row, dict) else {}
    if is_manga(book):
        if ocr_estimate_eligible(book):
            return str(book.get("ocr_source") or "ocr-estimate")
        return "not-applicable"
    return str(book.get("word_count_source") or "").strip().casefold()


def source_label(value: Any) -> str:
    source = str(value or "")
    if source == "ocr-estimate":
        return "OCR 估算"
    if source == "ocr-sampled":
        return "OCR 采样估算"
    if source == "ocr-series-estimate":
        return "同系列采样外推"
    if source == "en-regex":
        return "英文分词"
    if source == "jieba":
        return "Jieba"
    if source == "janome":
        return "Janome"
    if source == "not-applicable":
        return "漫画不适用"
    if source:
        return source
    return "Kindle 估算"


def source_label_en(value: Any) -> str:
    source = str(value or "")
    if source == "ocr-estimate":
        return "OCR estimate"
    if source == "ocr-sampled":
        return "OCR sampled estimate"
    if source == "ocr-series-estimate":
        return "Series sample estimate"
    if source == "en-regex":
        return "English tokenizer"
    if source == "jieba":
        return "Jieba"
    if source == "janome":
        return "Janome"
    if source == "not-applicable":
        return "Not applicable to manga"
    if source:
        return source
    return "Kindle estimate"


def reading_state(row: Dict[str, Any]) -> str:
    total_time = parse_int(row.get("total_time_ms")) or 0
    total_words = 0 if is_manga(row) else parse_int(row.get("total_words")) or 0
    return "reading" if total_time > 0 or total_words > 0 else "unread"


def build_insights(rows: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if not rows:
        return []

    most_read = max(
        rows, key=lambda row: parse_int(row.get("total_time_ms")) or 0
    )
    word_rows = [
        row
        for row in rows
        if not is_manga(row) or ocr_estimate_eligible(row)
    ]
    longest = max(
        word_rows,
        key=lambda row: effective_book_words(row) or 0,
        default=None,
    )
    fastest = max(
        (
            row
            for row in word_rows
            if (parse_float(effective_words_read(row)) or 0) > 0
        ),
        key=lambda row: effective_wpm(row) or 0,
        default=None,
    )
    total_time_ms = sum(
        parse_int(row.get("total_time_ms")) or 0 for row in rows
    )
    most_read_time_ms = parse_int(most_read.get("total_time_ms")) or 0
    most_read_share = (
        most_read_time_ms / total_time_ms
        if total_time_ms > 0 and most_read_time_ms > 0
        else None
    )

    return [
        {
            "key": "most-read",
            "layout": "feature",
            "book_id": stable_book_id(most_read),
            "share": most_read_share,
            "label": "阅读最多",
            "label_en": "Most read",
            "value": display_title(str(most_read.get("title", ""))),
            "detail": format_duration(parse_int(most_read.get("total_time_ms"))),
            "detail_en": format_duration_en(
                parse_int(most_read.get("total_time_ms"))
            ),
        },
        {
            "key": "longest-book",
            "layout": "cover",
            "book_id": stable_book_id(longest) if longest is not None else "",
            "label": "最长书籍",
            "label_en": "Longest book",
            "value": (
                display_title(str(longest.get("title", "")))
                if longest is not None
                else "不适用 / N/A"
            ),
            "detail": (
                format_words(parse_int(effective_book_words(longest)))
                if longest is not None
                else "--"
            ),
            "detail_en": (
                format_words(parse_int(effective_book_words(longest)))
                if longest is not None
                else "--"
            ),
        },
        {
            "key": "fastest-wpm",
            "layout": "cover",
            "book_id": stable_book_id(fastest) if fastest is not None else "",
            "label": "速度最高",
            "label_en": "Highest WPM",
            "value": (
                display_title(str(fastest.get("title", "")))
                if fastest is not None
                else "不适用 / N/A"
            ),
            "detail": (
                f'{effective_wpm(fastest) or 0:.0f} WPM'
                if fastest is not None
                else "--"
            ),
            "detail_en": (
                f'{effective_wpm(fastest) or 0:.0f} WPM'
                if fastest is not None
                else "--"
            ),
        },
    ]


def _insight_cover_html(
    item: Dict[str, Any],
    cover_urls: Optional[Mapping[str, str]],
) -> str:
    """Render the small book cover (or its initial-letter placeholder)."""

    title = str(item.get("value") or "")
    book_id = str(item.get("book_id") or "")
    cover_url = ""
    if cover_urls and book_id:
        cover_url = str(cover_urls.get(book_id) or "")
    initial = title.strip()[:1] or "?"
    placeholder = (
        '<span class="insight-cover-placeholder" aria-hidden="true">'
        f"<strong>{html.escape(initial)}</strong>"
        f"<small>{bilingual_text('无封面', 'No cover')}</small>"
        "</span>"
    )
    if not cover_url:
        return f'<span class="insight-cover">{placeholder}</span>'
    # Same graceful degradation as the cover shelf: hide a broken image and
    # reveal the placeholder that follows it.
    hidden_placeholder = placeholder.replace(">", " hidden>", 1)
    return (
        '<span class="insight-cover">'
        f'<img data-insight-cover src="{html.escape(cover_url, quote=True)}" '
        f'alt="{html.escape(title, quote=True)} 封面 / Cover" '
        'onerror="this.hidden=true;this.nextElementSibling.hidden=false;">'
        f"{hidden_placeholder}"
        "</span>"
    )


def insight_cards_html(
    items: Sequence[Dict[str, Any]],
    cover_urls: Optional[Mapping[str, str]] = None,
) -> str:
    """Render insight cards, tolerating empty or partial items."""

    cards: List[str] = []
    for item in items:
        key = str(item.get("key") or "insight")
        layout = str(item.get("layout") or "plain")
        value = str(item.get("value", ""))
        value_en = str(item.get("value_en", value))
        detail = str(item.get("detail", ""))
        detail_en = str(item.get("detail_en", detail))
        icon = INSIGHT_ICONS.get(key, "")
        icon_html = (
            '<span class="insight-icon" aria-hidden="true">'
            '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" '
            f'stroke-width="1.8" stroke-linecap="round" '
            f'stroke-linejoin="round">{icon}</svg></span>'
            if icon
            else ""
        )
        cover_html = (
            _insight_cover_html(item, cover_urls)
            if layout in {"feature", "cover"}
            else ""
        )
        share_html = ""
        share = item.get("share")
        if layout == "feature" and isinstance(share, (int, float)) and share > 0:
            percent = max(0, min(100, int(round(float(share) * 100))))
            share_html = (
                '<span class="insight-share">'
                '<span class="insight-share-track" aria-hidden="true">'
                f'<i style="width:{percent}%"></i></span>'
                '<small class="insight-share-label">'
                + bilingual_text(
                    f"占累计阅读 {percent}%",
                    f"{percent}% of total reading time",
                )
                + "</small></span>"
            )
        cards.append(
            f'<article class="insight" '
            f'data-insight-key="{html.escape(key, quote=True)}" '
            f'data-insight-layout="{html.escape(layout, quote=True)}">'
            f"{cover_html}"
            '<div class="insight-body">'
            '<div class="insight-head">'
            f"{icon_html}"
            '<span class="insight-label">'
            f'{bilingual_text(item["label"], item.get("label_en", item["label"]))}'
            "</span></div>"
            f'<strong class="insight-value" title="{html.escape(value, quote=True)}">'
            f"{bilingual_text(value, value_en)}</strong>"
            '<span class="insight-detail">'
            f"{bilingual_text(detail, detail_en)}</span>"
            f"{share_html}"
            "</div></article>"
        )
    return "\n".join(cards)


def load_rows(csv_path: Path) -> List[Dict[str, Any]]:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        return [dict(row) for row in csv.DictReader(handle)]


def write_csv_rows(
    csv_path: Path,
    rows: Sequence[Dict[str, Any]],
    fieldnames: Sequence[str],
) -> None:
    fields = [str(field) for field in fieldnames if str(field)]
    buffer = io.StringIO(newline="")
    writer = csv.DictWriter(buffer, fieldnames=fields)
    writer.writeheader()
    for row in rows:
        writer.writerow({field: row.get(field) for field in fields})
    _atomic_write_text(csv_path, buffer.getvalue(), encoding="utf-8-sig")


def _xlsx_safe_text(value: Any) -> str:
    text = str(value)
    if text.startswith(("=", "+", "-", "@")):
        return "'" + text
    return text


def _xlsx_raw_value(field: str, value: Any) -> Any:
    if value is None:
        return None
    text = str(value).strip()
    if text == "":
        return None
    if field in XLSX_NUMERIC_FIELDS:
        number = parse_float(text)
        if number is not None:
            if field in {"minutes", "hours", "book_word_coverage_percent",
                         "wpm", "wpm_overall", "wpm_reading",
                         "progress_percent", "progress", "ocr_confidence",
                         "ocr_page_coverage", "effective_wpm"}:
                return number
            return int(round(number))
    if field == "has_data":
        lowered = text.casefold()
        if lowered in {"true", "1", "yes"}:
            return "TRUE"
        if lowered in {"false", "0", "no"}:
            return "FALSE"
    return _xlsx_safe_text(text)


def write_xlsx_report(
    rows: Sequence[Dict[str, Any]],
    output_path: Path,
    fieldnames: Sequence[str] = (),
    source_path: str = "",
    generated_at: str = "",
    trends: Optional[Dict[str, Any]] = None,
) -> None:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError as exc:
        raise RuntimeError(
            "openpyxl is required for XLSX export"
        ) from exc

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    useful_rows = sorted(
        rows,
        key=lambda row: (
            -(parse_int(row.get("total_time_ms")) or 0),
            display_title(str(row.get("title", ""))).casefold(),
        ),
    )
    trend_by_book = {
        str(change.get("book_id")): change
        for change in (trends or {}).get("books", [])
        if isinstance(change, dict)
    }
    quality_by_book = {
        stable_book_id(row): analyze_book_quality(row)
        for row in useful_rows
    }
    base_fields = [
        field
        for field in fieldnames
        if field
    ]
    if not base_fields and useful_rows:
        base_fields = list(useful_rows[0].keys())
    extra_fields = [
        "reading_mode",
        "effective_book_words",
        "effective_words_read",
        "effective_wpm",
        "effective_source",
        "word_origin",
        "word_progress",
        "confidence_level",
        "confidence_score",
        "quality_flags",
        "time_delta_ms",
        "words_delta",
        "progress_delta",
    ]
    extra_fields = [
        field for field in extra_fields if field not in base_fields
    ]
    headers = [
        XLSX_HEADER_LABELS.get(field, field)
        for field in base_fields + extra_fields
    ]

    workbook = Workbook()
    summary_sheet = workbook.active
    summary_sheet.title = "Summary"
    books_sheet = workbook.create_sheet("Books")
    header_fill = PatternFill("solid", fgColor="2F7D61")
    header_font = Font(color="FFFFFF", bold=True)
    header_alignment = Alignment(vertical="center", wrap_text=True)

    summary = build_summary(useful_rows)
    quality = build_quality_summary(useful_rows)
    summary_rows = [
        ("Generated / 生成时间", generated_at),
        ("Source / 数据来源", str(source_path or "")),
        ("Books / 书籍数", summary["books"]),
        ("Total reading ms / 累计阅读毫秒", summary["total_time_ms"]),
        ("Total words / 累计词数", summary["total_words"]),
        (
            "Effective words read / 有效已读词数",
            summary["effective_words_read"],
        ),
        ("Weighted WPM / 加权平均 WPM", summary["weighted_wpm"]),
        (
            "OCR-estimate books / OCR 估算并入书籍",
            summary["ocr_estimate_books"],
        ),
        (
            "Average word progress / 平均词数进度",
            summary["average_word_progress"],
        ),
        ("Exact word-count books / 精确词数书籍", quality["exact_count"]),
        (
            "Low-confidence books / 低置信书籍",
            quality["low_confidence_count"],
        ),
    ]
    for row_index, (label, value) in enumerate(summary_rows, start=1):
        summary_sheet.cell(row=row_index, column=1, value=label)
        summary_sheet.cell(
            row=row_index,
            column=2,
            value="" if value is None else value,
        )
        summary_sheet.cell(row=row_index, column=1).font = Font(bold=True)
    summary_sheet.column_dimensions["A"].width = 42
    summary_sheet.column_dimensions["B"].width = 70

    books_sheet.append(headers)
    for cell in books_sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

    for row in useful_rows:
        book_id = stable_book_id(row)
        book_quality = quality_by_book.get(book_id, {})
        change = trend_by_book.get(book_id) or {}
        values = [
            _xlsx_raw_value(field, row.get(field))
            for field in base_fields
        ]
        extra_values = {
            "reading_mode": reading_mode_for_book(row),
            "effective_book_words": parse_int(effective_book_words(row)),
            "effective_words_read": parse_int(effective_words_read(row)),
            "effective_wpm": parse_float(effective_wpm(row)),
            "effective_source": effective_source(row),
            "word_origin": book_word_origin(row),
            "word_progress": word_progress(row),
            "confidence_level": book_quality.get("confidence_level", ""),
            "confidence_score": book_quality.get("confidence_score", ""),
            "quality_flags": "; ".join(book_quality.get("flags", [])),
            "time_delta_ms": parse_float(
                (change.get("total_time_ms") or {}).get("delta")
            ),
            "words_delta": parse_float(
                (change.get("total_words") or {}).get("delta")
            ),
            "progress_delta": parse_float(
                (change.get("progress_percent") or {}).get("delta")
            ),
        }
        values.extend(extra_values[field] for field in extra_fields)
        books_sheet.append(values)

    for column_index, header in enumerate(headers, start=1):
        column_letter = get_column_letter(column_index)
        values = [
            str(books_sheet.cell(row=row_index, column=column_index).value or "")
            for row_index in range(1, min(books_sheet.max_row, 80) + 1)
        ]
        width = min(46, max(12, max(map(len, values), default=0) + 2))
        books_sheet.column_dimensions[column_letter].width = width
    books_sheet.freeze_panes = "A2"
    if books_sheet.max_row >= 1 and books_sheet.max_column >= 1:
        books_sheet.auto_filter.ref = books_sheet.dimensions

    temporary_name: Optional[str] = None
    try:
        file_descriptor, temporary_name = tempfile.mkstemp(
            dir=output_path.parent,
            prefix=output_path.name + ".",
            suffix=".tmp.xlsx",
        )
        os.close(file_descriptor)
        workbook.save(temporary_name)
        os.replace(temporary_name, output_path)
    finally:
        if temporary_name:
            Path(temporary_name).unlink(missing_ok=True)


def _snapshot_fingerprint(snapshot: Dict[str, Any]) -> str:
    books = snapshot.get("books", {})
    if not isinstance(books, dict):
        books = {}
    return json.dumps(
        books,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )


def _is_duplicate_snapshot(
    history: Dict[str, Any], snapshot: Dict[str, Any]
) -> bool:
    snapshots = history.get("snapshots", [])
    if not isinstance(snapshots, list) or not snapshots:
        return False
    previous = snapshots[-1]
    return isinstance(previous, dict) and (
        _snapshot_fingerprint(previous) == _snapshot_fingerprint(snapshot)
    )


def _snapshot_has_ocr_effective(snapshot: Dict[str, Any]) -> bool:
    books = snapshot.get("books")
    if not isinstance(books, dict):
        return False
    return any(
        isinstance(book, dict)
        and str(book.get("effective_source") or "")
        in {"ocr-estimate", "ocr-sampled", "ocr-series-estimate"}
        for book in books.values()
    )


def _snapshot_has_manga_progress_improvement(
    previous: Dict[str, Any],
    current: Dict[str, Any],
) -> bool:
    previous_books = previous.get("books")
    current_books = current.get("books")
    if not isinstance(previous_books, dict) or not isinstance(
        current_books, dict
    ):
        return False
    for book_id, book in current_books.items():
        if not isinstance(book, dict) or not is_manga(book):
            continue
        previous_book = previous_books.get(book_id)
        if not isinstance(previous_book, dict):
            continue
        for key in (
            "total_time_ms",
            "progress_percent",
            "manga_pages_read",
        ):
            current_value = parse_float(book.get(key))
            previous_value = parse_float(previous_book.get(key))
            if (
                current_value is not None
                and (
                    previous_value is None
                    or current_value > previous_value + 1e-9
                )
            ):
                return True
    return False


def load_history(history_path: Path) -> Dict[str, Any]:
    if not history_path.is_file():
        return {"version": HISTORY_VERSION, "snapshots": []}

    try:
        with history_path.open("r", encoding="utf-8-sig") as handle:
            history = json.load(handle)
    except (OSError, ValueError):
        backup_corrupt_file(history_path)
        return {"version": HISTORY_VERSION, "snapshots": []}

    if not isinstance(history, dict):
        backup_corrupt_file(history_path)
        return {"version": HISTORY_VERSION, "snapshots": []}
    snapshots = history.get("snapshots", [])
    if not isinstance(snapshots, list):
        backup_corrupt_file(history_path)
        return {"version": HISTORY_VERSION, "snapshots": []}
    if not all(isinstance(snapshot, dict) for snapshot in snapshots):
        backup_corrupt_file(history_path)
        return {"version": HISTORY_VERSION, "snapshots": []}
    return history


def save_history(
    history_path: Path,
    history: Dict[str, Any],
    snapshot: Dict[str, Any],
    history_limit: int = DEFAULT_HISTORY_LIMIT,
) -> bool:
    history_path.parent.mkdir(parents=True, exist_ok=True)
    output = dict(history)
    output["version"] = HISTORY_VERSION
    snapshots = list(history.get("snapshots", []))
    previous = snapshots[-1] if snapshots else None
    if (
        isinstance(previous, dict)
        and _snapshot_has_ocr_effective(previous)
        and not _snapshot_has_ocr_effective(snapshot)
        and not _snapshot_has_manga_progress_improvement(previous, snapshot)
    ):
        # Turning OCR off must not append a regressive OCR-blind snapshot over
        # history that already contains effective OCR estimates.
        return False
    duplicate = _is_duplicate_snapshot(history, snapshot)
    if not duplicate:
        snapshots.append(snapshot)

    limit = max(0, parse_int(history_limit) or 0)
    if limit and len(snapshots) > limit:
        snapshots = snapshots[-limit:]

    previous_snapshots = history.get("snapshots", [])
    changed = (
        not isinstance(previous_snapshots, list)
        or previous_snapshots != snapshots
    )
    if not changed:
        return False

    output["snapshots"] = snapshots
    document = json.dumps(
        output,
        ensure_ascii=False,
        indent=2,
        sort_keys=True,
    )
    _atomic_write_text(history_path, document + "\n")
    return not duplicate


def append_snapshot(
    history_path: Path,
    snapshot: Dict[str, Any],
    history_limit: int = DEFAULT_HISTORY_LIMIT,
) -> bool:
    history = load_history(history_path)
    return save_history(
        history_path,
        history,
        snapshot,
        history_limit=history_limit,
    )


def build_snapshot(
    rows: Sequence[Dict[str, Any]],
    captured_at: str,
    source_path: str,
    book_modes: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    books: Dict[str, Dict[str, Any]] = {}
    has_ocr_effective = False
    for row in rows:
        book_id = stable_book_id(row)
        mode = reading_mode_for_book(row, book_modes)
        effective_words = effective_words_read(row)
        effective_book_words_value = effective_book_words(row)
        effective_wpm_value = effective_wpm(row)
        effective_source_value = effective_source(row)
        snapshot_book: Dict[str, Any] = {
            "book_id": book_id,
            "asin": str(row.get("asin") or "").strip(),
            "title": display_title(str(row.get("title", ""))),
            "author": str(row.get("author") or "").strip(),
            "total_time_ms": parse_int(row.get("total_time_ms")) or 0,
            "total_words": parse_int(row.get("total_words")) or 0,
            "wpm": parse_float(row.get("wpm")) or 0.0,
            "progress_percent": parse_float(row.get("progress_percent")),
        }
        # Include the explicit mode for manga and derived effective fields.
        # Old snapshots without these keys remain readable through fallbacks.
        if mode == "manga":
            snapshot_book["reading_mode"] = "manga"
            snapshot_book["total_words"] = 0
            snapshot_book["wpm"] = 0.0
            for key in (
                "manga_pages_total",
                "manga_pages_read",
                "manga_pages_source",
                "manga_progress_source",
                "manga_position",
                "manga_max_position",
            ):
                value = row.get(key)
                if value not in (None, ""):
                    snapshot_book[key] = value
        last_read = str(row.get("last_read") or "").strip()
        if last_read:
            snapshot_book["last_read"] = last_read
        if effective_words is not None:
            snapshot_book["effective_words_read"] = round(effective_words, 2)
        if effective_book_words_value is not None:
            snapshot_book["effective_book_words"] = round(
                effective_book_words_value, 2
            )
        if effective_wpm_value is not None:
            snapshot_book["effective_wpm"] = round(effective_wpm_value, 2)
        if effective_source_value:
            snapshot_book["effective_source"] = effective_source_value
        if ocr_estimate_eligible(row):
            has_ocr_effective = True
            snapshot_book["ocr_status"] = str(row.get("ocr_status") or "")
            snapshot_book["ocr_source"] = str(
                row.get("ocr_source") or "ocr-estimate"
            )
        books[book_id] = snapshot_book
    snapshot = {
        "captured_at": captured_at,
        "source": source_path,
        "data_source_version": VERSION,
        "book_count": len(books),
        "books": books,
    }
    if has_ocr_effective:
        snapshot["ocr_enabled_at"] = captured_at
        snapshot["effective_word_source"] = "ocr-estimate"
    return snapshot


def _snapshot_book_effective_words_read(
    book: Dict[str, Any],
) -> Optional[float]:
    if "effective_words_read" in book:
        return parse_float(book.get("effective_words_read"))
    if reading_mode_for_book(book) == "manga":
        return None
    return parse_float(book.get("total_words"))


def _snapshot_book_effective_book_words(
    book: Dict[str, Any],
) -> Optional[float]:
    if "effective_book_words" in book:
        return parse_float(book.get("effective_book_words"))
    if reading_mode_for_book(book) == "manga":
        return None
    return parse_float(book.get("total_words"))


def _snapshot_book_effective_wpm(
    book: Dict[str, Any],
) -> Optional[float]:
    if "effective_wpm" in book:
        return parse_float(book.get("effective_wpm"))
    if reading_mode_for_book(book) == "manga":
        return None
    return parse_float(book.get("wpm"))


def snapshot_totals(snapshot: Dict[str, Any]) -> Dict[str, Any]:
    books = [
        book
        for book in (snapshot.get("books") or {}).values()
        if isinstance(book, dict)
    ]
    total_time_ms = sum(parse_int(book.get("total_time_ms")) or 0 for book in books)
    kindle_word_books = [
        book
        for book in books
        if reading_mode_for_book(book) != "manga"
    ]
    word_books = [
        book
        for book in books
        if _snapshot_book_effective_words_read(book) is not None
    ]
    kindle_word_time_ms = sum(
        parse_int(book.get("total_time_ms")) or 0
        for book in kindle_word_books
    )
    effective_word_time_ms = sum(
        parse_int(book.get("total_time_ms")) or 0 for book in word_books
    )
    kindle_total_words = sum(
        parse_int(book.get("total_words")) or 0
        for book in kindle_word_books
    )
    effective_words_read = sum(
        _snapshot_book_effective_words_read(book) or 0.0
        for book in word_books
    )
    effective_total_words = sum(
        _snapshot_book_effective_book_words(book) or 0.0
        for book in word_books
    )
    minutes = effective_word_time_ms / 60000.0
    progress_values = [
        value
        for value in (
            parse_float(book.get("progress_percent")) for book in word_books
        )
        if value is not None
    ]
    return {
        "books": len(books),
        "word_books": len(word_books),
        "kindle_word_books": len(kindle_word_books),
        "manga_books": sum(
            1 for book in books if reading_mode_for_book(book) == "manga"
        ),
        "ocr_estimate_books": sum(
            1
            for book in books
            if str(book.get("effective_source") or "")
            in {"ocr-estimate", "ocr-sampled", "ocr-series-estimate"}
        ),
        "reading_record_books": sum(
            1 for book in books if has_reading_record(book)
        ),
        "completed_books": sum(
            1 for book in books if is_completed_book(book)
        ),
        "completed_manga_volumes": sum(
            1 for book in books if is_completed_manga_volume(book)
        ),
        "total_time_ms": total_time_ms,
        "word_time_ms": effective_word_time_ms,
        "effective_word_time_ms": effective_word_time_ms,
        "kindle_word_time_ms": kindle_word_time_ms,
        "total_words": kindle_total_words,
        "kindle_total_words": kindle_total_words,
        "effective_words_read": round(effective_words_read, 2),
        "effective_total_words": round(effective_total_words, 2),
        "words_include_ocr": any(
            str(book.get("effective_source") or "")
            in {"ocr-estimate", "ocr-sampled", "ocr-series-estimate"}
            for book in books
        ),
        "word_source": (
            "ocr-estimate"
            if any(
                str(book.get("effective_source") or "")
                in {"ocr-estimate", "ocr-sampled", "ocr-series-estimate"}
                for book in books
            )
            else "kindle-raw"
        ),
        "weighted_wpm": (
            round(effective_words_read / minutes, 1) if minutes else None
        ),
        "kindle_weighted_wpm": (
            round(kindle_total_words / (kindle_word_time_ms / 60000.0), 1)
            if kindle_word_time_ms
            else None
        ),
        "average_progress": (
            round(sum(progress_values) / len(progress_values), 1)
            if progress_values
            else None
        ),
    }


def _metric_change(
    current: Any,
    previous: Any,
    has_previous: bool,
) -> Dict[str, Any]:
    current_number = parse_float(current)
    previous_number = parse_float(previous)
    delta = None
    if (
        has_previous
        and current_number is not None
        and previous_number is not None
    ):
        delta = current_number - previous_number
    return {
        "current": current_number,
        "previous": previous_number,
        "delta": delta,
    }


def _snapshot_datetime(snapshot: Dict[str, Any]) -> Optional[datetime]:
    value = str(snapshot.get("captured_at") or "").strip()
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None


def _snapshot_date(snapshot: Dict[str, Any]) -> Optional[date]:
    parsed = _snapshot_datetime(snapshot)
    return parsed.date() if parsed is not None else None


def _dedupe_consecutive_snapshots(
    snapshots: Sequence[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    result: List[Dict[str, Any]] = []
    for snapshot in snapshots:
        if not isinstance(snapshot, dict):
            continue
        if result and (
            _snapshot_fingerprint(result[-1]) == _snapshot_fingerprint(snapshot)
        ):
            continue
        result.append(snapshot)
    return result


def _daily_reading_totals(
    snapshots: Sequence[Dict[str, Any]],
) -> Dict[str, Dict[str, Any]]:
    totals_by_day: Dict[str, Dict[str, Any]] = {}
    previous_totals: Optional[Dict[str, Any]] = None

    for snapshot in snapshots:
        captured_date = _snapshot_date(snapshot)
        if captured_date is None:
            continue

        current_totals = snapshot_totals(snapshot)
        day_key = captured_date.isoformat()
        bucket = totals_by_day.setdefault(
            day_key,
            {
                "date": day_key,
                "total_time_ms": 0,
                "total_words": 0,
                "snapshot_count": 0,
                "active": False,
            },
        )
        bucket["snapshot_count"] += 1

        if previous_totals is not None:
            time_delta = max(
                0,
                (parse_int(current_totals.get("total_time_ms")) or 0)
                - (parse_int(previous_totals.get("total_time_ms")) or 0),
            )
            words_delta = max(
                0,
                (
                    parse_int(current_totals.get("effective_words_read"))
                    or parse_int(current_totals.get("total_words"))
                    or 0
                )
                - (
                    parse_int(previous_totals.get("effective_words_read"))
                    or parse_int(previous_totals.get("total_words"))
                    or 0
                ),
            )
            bucket["total_time_ms"] += time_delta
            bucket["total_words"] += words_delta
            if time_delta or words_delta:
                bucket["active"] = True

        previous_totals = current_totals

    return totals_by_day


def _daily_window_rows(
    totals_by_day: Dict[str, Dict[str, Any]],
    end_date: date,
    days: int,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for offset in range(days - 1, -1, -1):
        day = end_date - timedelta(days=offset)
        day_key = day.isoformat()
        existing = totals_by_day.get(day_key)
        if existing is None:
            rows.append(
                {
                    "date": day_key,
                    "total_time_ms": 0,
                    "total_words": 0,
                    "snapshot_count": 0,
                    "active": False,
                }
            )
        else:
            rows.append(dict(existing))
    return rows


def _daily_window_summary(
    rows: Sequence[Dict[str, Any]],
    days: int,
) -> Dict[str, Any]:
    total_time_ms = sum(
        parse_int(row.get("total_time_ms")) or 0 for row in rows
    )
    total_words = sum(
        parse_int(row.get("effective_words_read"))
        or parse_int(row.get("total_words"))
        or 0
        for row in rows
    )
    minutes = total_time_ms / 60000.0
    return {
        "days": days,
        "start": str(rows[0].get("date") or "") if rows else "",
        "end": str(rows[-1].get("date") or "") if rows else "",
        "total_time_ms": total_time_ms,
        "total_words": total_words,
        "active_days": sum(1 for row in rows if row.get("active")),
        "weighted_wpm": round(total_words / minutes, 1) if minutes else None,
    }


def _reading_streaks(
    totals_by_day: Dict[str, Dict[str, Any]],
    end_date: date,
) -> Dict[str, int]:
    active_dates = {
        date.fromisoformat(day_key)
        for day_key, values in totals_by_day.items()
        if values.get("active")
    }

    current_days = 0
    cursor = end_date
    while cursor in active_dates:
        current_days += 1
        cursor -= timedelta(days=1)

    longest_days = 0
    running_days = 0
    previous_date: Optional[date] = None
    for active_date in sorted(active_dates):
        if previous_date is not None and active_date == previous_date + timedelta(
            days=1
        ):
            running_days += 1
        else:
            running_days = 1
        longest_days = max(longest_days, running_days)
        previous_date = active_date

    return {
        "current_days": current_days,
        "longest_days": longest_days,
    }


def _window_baseline_snapshot(
    snapshots: Sequence[Dict[str, Any]],
    start_date: date,
    end_date: date,
) -> Optional[Dict[str, Any]]:
    baseline: Optional[Dict[str, Any]] = None
    for snapshot in snapshots:
        snapshot_date = _snapshot_date(snapshot)
        if snapshot_date is None:
            continue
        if snapshot_date < start_date:
            baseline = snapshot
        elif snapshot_date <= end_date and baseline is None:
            baseline = snapshot
        elif snapshot_date > end_date:
            break
    return baseline


def _window_book_change(
    current: Dict[str, Any],
    baseline: Optional[Dict[str, Any]],
) -> Dict[str, Any]:
    has_baseline = isinstance(baseline, dict)
    return {
        "total_time_ms": _metric_change(
            current.get("total_time_ms"),
            baseline.get("total_time_ms") if has_baseline else None,
            has_baseline,
        ),
        "total_words": _metric_change(
            _snapshot_book_effective_words_read(current),
            (
                _snapshot_book_effective_words_read(baseline)
                if has_baseline
                else None
            ),
            has_baseline,
        ),
        "wpm": _metric_change(
            _snapshot_book_effective_wpm(current),
            (
                _snapshot_book_effective_wpm(baseline)
                if has_baseline
                else None
            ),
            has_baseline,
        ),
        "progress_percent": _metric_change(
            current.get("progress_percent"),
            baseline.get("progress_percent") if has_baseline else None,
            has_baseline,
        ),
    }


def _ocr_transition(
    snapshots: Sequence[Dict[str, Any]],
) -> Optional[Dict[str, Any]]:
    previous_had_ocr = False
    for snapshot in snapshots:
        has_ocr = _snapshot_has_ocr_effective(snapshot)
        if has_ocr != previous_had_ocr:
            return {
                "state": "enabled" if has_ocr else "disabled",
                "captured_at": str(snapshot.get("captured_at") or ""),
                "source": str(snapshot.get("source") or ""),
            }
        previous_had_ocr = has_ocr
    return None


def build_trends(
    current_snapshot: Dict[str, Any],
    history: Dict[str, Any],
) -> Dict[str, Any]:
    snapshots = [
        snapshot
        for snapshot in history.get("snapshots", [])
        if isinstance(snapshot, dict)
    ]
    previous_snapshot = snapshots[-1] if snapshots else None
    has_previous = previous_snapshot is not None
    previous_books = (
        previous_snapshot.get("books", {})
        if isinstance(previous_snapshot, dict)
        else {}
    )
    current_books = current_snapshot.get("books", {})

    current_totals = snapshot_totals(current_snapshot)
    previous_totals = (
        snapshot_totals(previous_snapshot)
        if previous_snapshot is not None
        else {}
    )
    totals = {
        "total_time_ms": _metric_change(
            current_totals["total_time_ms"],
            previous_totals.get("total_time_ms"),
            has_previous,
        ),
        "total_words": _metric_change(
            current_totals.get("effective_words_read"),
            previous_totals.get("effective_words_read"),
            has_previous,
        ),
        "weighted_wpm": _metric_change(
            current_totals["weighted_wpm"],
            previous_totals.get("weighted_wpm"),
            has_previous,
        ),
        "average_progress": _metric_change(
            current_totals["average_progress"],
            previous_totals.get("average_progress"),
            has_previous,
        ),
    }

    current_date = _snapshot_date(current_snapshot)
    valid_history_dates = [
        snapshot_date
        for snapshot_date in (_snapshot_date(item) for item in snapshots)
        if snapshot_date is not None
    ]
    end_date = current_date or (
        valid_history_dates[-1] if valid_history_dates else date.today()
    )
    combined_snapshots = snapshots + [current_snapshot]
    ocr_transition = _ocr_transition(combined_snapshots)
    day_totals = _daily_reading_totals(combined_snapshots)
    daily = _daily_window_rows(day_totals, end_date, 30)
    window_7d = _daily_window_summary(daily[-7:], 7)
    window_30d = _daily_window_summary(daily, 30)
    streak = _reading_streaks(day_totals, end_date)

    start_7d = end_date - timedelta(days=6)
    start_30d = end_date - timedelta(days=29)
    baseline_7d = _window_baseline_snapshot(
        snapshots,
        start_7d,
        end_date,
    )
    baseline_30d = _window_baseline_snapshot(
        snapshots,
        start_30d,
        end_date,
    )
    baseline_7d_books = (
        baseline_7d.get("books", {})
        if isinstance(baseline_7d, dict)
        else {}
    )
    baseline_30d_books = (
        baseline_30d.get("books", {})
        if isinstance(baseline_30d, dict)
        else {}
    )

    changes: List[Dict[str, Any]] = []
    for book_id, current in current_books.items():
        if not isinstance(current, dict):
            continue
        previous = previous_books.get(book_id)
        if not isinstance(previous, dict):
            previous = None
        is_new = has_previous and previous is None
        change = {
            "book_id": book_id,
            "title": str(current.get("title") or "(unknown)"),
            "author": str(current.get("author") or ""),
            "is_new": is_new,
            "total_time_ms": _metric_change(
                current.get("total_time_ms"),
                previous.get("total_time_ms") if previous else None,
                has_previous and previous is not None,
            ),
            "total_words": _metric_change(
                _snapshot_book_effective_words_read(current),
                (
                    _snapshot_book_effective_words_read(previous)
                    if previous
                    else None
                ),
                has_previous and previous is not None,
            ),
            "wpm": _metric_change(
                _snapshot_book_effective_wpm(current),
                (
                    _snapshot_book_effective_wpm(previous)
                    if previous
                    else None
                ),
                has_previous and previous is not None,
            ),
            "progress_percent": _metric_change(
                current.get("progress_percent"),
                previous.get("progress_percent") if previous else None,
                has_previous and previous is not None,
            ),
        }
        change["windows"] = {
            "7d": _window_book_change(
                current,
                (
                    baseline_7d_books.get(book_id)
                    if isinstance(baseline_7d_books, dict)
                    else None
                ),
            ),
            "30d": _window_book_change(
                current,
                (
                    baseline_30d_books.get(book_id)
                    if isinstance(baseline_30d_books, dict)
                    else None
                ),
            ),
        }
        changes.append(change)

    def movement(item: Dict[str, Any]) -> float:
        time_delta = abs(parse_float(item["total_time_ms"]["delta"]) or 0.0)
        word_delta = abs(parse_float(item["total_words"]["delta"]) or 0.0)
        progress_delta = abs(
            parse_float(item["progress_percent"]["delta"]) or 0.0
        )
        return time_delta + word_delta * 100.0 + progress_delta * 1000.0

    changes.sort(key=movement, reverse=True)
    timeline_snapshots = _dedupe_consecutive_snapshots(combined_snapshots)[-30:]
    timeline: List[Dict[str, Any]] = []
    previous_timeline_totals: Optional[Dict[str, Any]] = None
    for snapshot in timeline_snapshots:
        snapshot_point_totals = snapshot_totals(snapshot)
        timeline.append(
            {
                "captured_at": str(snapshot.get("captured_at") or ""),
                "source": str(snapshot.get("source") or ""),
                "totals": snapshot_point_totals,
                "delta_time_ms": (
                    (
                        parse_int(snapshot_point_totals.get("total_time_ms"))
                        or 0
                    )
                    - (
                        parse_int(
                            previous_timeline_totals.get("total_time_ms")
                        )
                        or 0
                    )
                    if previous_timeline_totals is not None
                    else None
                ),
                "delta_words": (
                    (
                        parse_int(
                            snapshot_point_totals.get("effective_words_read")
                        )
                        or 0
                    )
                    - (
                        parse_int(
                            previous_timeline_totals.get(
                                "effective_words_read"
                            )
                        )
                        or 0
                    )
                    if previous_timeline_totals is not None
                    else None
                ),
            }
        )
        previous_timeline_totals = snapshot_point_totals
    return {
        "has_previous": has_previous,
        "snapshot_count": len(snapshots) + 1,
        "previous_captured_at": (
            str(previous_snapshot.get("captured_at") or "")
            if previous_snapshot is not None
            else ""
        ),
        "current_captured_at": str(
            current_snapshot.get("captured_at") or ""
        ),
        "totals": totals,
        "books": changes,
        "timeline": timeline,
        "daily": daily,
        "windows": {
            "7d": window_7d,
            "30d": window_30d,
        },
        "streak": streak,
        "ocr_transition": ocr_transition,
        "ocr_effective_enabled": _snapshot_has_ocr_effective(
            current_snapshot
        ),
    }


def resolve_fmcache_db(
    source_path: Any,
    explicit: Any = None,
) -> Optional[Path]:
    """Resolve fmcache.db without writing to or modifying the Kindle."""

    if explicit:
        return Path(explicit).expanduser().resolve()

    raw_source = str(source_path or "").strip()
    if not raw_source:
        return None

    source = Path(raw_source).expanduser()
    try:
        source = source.resolve()
    except OSError:
        return None

    if source.is_file():
        source = source.parent

    candidates = [source]
    candidates.extend(source.parents)
    for candidate in candidates:
        database = candidate / FMCACHE_RELATIVE_PATH
        try:
            if database.is_file():
                return database
        except OSError:
            continue
    return None


def _fmcache_json_object(value: Any) -> Optional[Dict[str, Any]]:
    if isinstance(value, dict):
        return value
    if isinstance(value, bytes):
        try:
            value = value.decode("utf-8")
        except UnicodeDecodeError:
            return None
    if not isinstance(value, str):
        return None
    try:
        decoded = json.loads(value)
    except (TypeError, ValueError):
        return None
    return decoded if isinstance(decoded, dict) else None


def _fmcache_datetime(value: Any) -> Optional[datetime]:
    if isinstance(value, datetime):
        parsed = value
    elif isinstance(value, date):
        parsed = datetime.combine(value, datetime.min.time())
    else:
        text = str(value or "").strip()
        if not text:
            return None
        try:
            parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
        except ValueError:
            number = parse_float(text)
            if number is None:
                return None
            parsed = datetime.fromtimestamp(
                number / 1000.0 if abs(number) >= 100_000_000_000 else number
            )

    if parsed.tzinfo is not None:
        parsed = parsed.astimezone().replace(tzinfo=None)
    if not 2000 <= parsed.year <= 2100:
        return None
    return parsed


def _fmcache_value(
    record: Dict[str, Any],
    keys: Sequence[str],
) -> Any:
    dictionaries = [record]
    payload = record.get("payload")
    if isinstance(payload, dict):
        dictionaries.insert(0, payload)
    for dictionary in dictionaries:
        for key in keys:
            if key in dictionary and dictionary[key] is not None:
                return dictionary[key]
    return None


def _fmcache_duration_ms(record: Dict[str, Any]) -> Optional[int]:
    value = _fmcache_value(
        record,
        (
            "duration_ms",
            "elapsed_ms",
            "reading_time_ms",
            "session_duration_ms",
            "total_time_ms",
            "duration_milliseconds",
        ),
    )
    duration = parse_int(value)
    if duration is None or duration <= 0:
        return None
    return duration


def _fmcache_asin(record: Dict[str, Any]) -> Optional[str]:
    value = _fmcache_value(
        record,
        (
            "asin",
            "library_asin",
            "asset_id",
            "embedded_id",
        ),
    )
    text = str(value or "").strip()
    if text:
        return _normalize_mode_asin(text)

    path_value = _fmcache_value(
        record,
        ("lpath", "library_path", "file_path", "path"),
    )
    if path_value:
        return "lpath:" + _strip_file_extension(
            _normalize_mode_asin(path_value).casefold()
        )
    title_value = _fmcache_value(
        record,
        ("title", "book_title", "display_title"),
    )
    if title_value:
        normalized = normalized_title(title_value)
        if normalized:
            return "title:" + normalized
    return None


def _strip_file_extension(value: str) -> str:
    text = str(value or "").replace("\\", "/")
    head, tail = text.rsplit("/", 1) if "/" in text else ("", text)
    for extension in EBOOK_EXTENSIONS:
        if tail.endswith(extension):
            tail = tail[: -len(extension)]
            break
    return f"{head}/{tail}" if head else tail


def _fmcache_progress_percent(record: Dict[str, Any]) -> Optional[float]:
    value = _fmcache_value(
        record,
        (
            "reading_progress",
            "progress_percent",
            "percent_read",
            "totalPercent",
        ),
    )
    number = parse_float(value)
    if number is None:
        return None
    if 0.0 <= number <= 1.0 and "totalPercent" in record:
        number *= 100.0
    return round(max(0.0, min(100.0, number)), 1)


def _fmcache_location(record: Dict[str, Any]) -> Optional[int]:
    value = _fmcache_value(
        record,
        (
            "end_reading_location",
            "max_position",
            "end_location",
            "location",
            "position",
        ),
    )
    return parse_int(value)


def _fmcache_session_times(
    record: Dict[str, Any],
) -> Optional[Tuple[datetime, datetime, int]]:
    start = _fmcache_datetime(
        _fmcache_value(
            record,
            (
                "start_timestamp",
                "start_time",
                "started_at",
                "session_start_timestamp",
            ),
        )
    )
    end = _fmcache_datetime(
        _fmcache_value(
            record,
            (
                "end_timestamp",
                "end_time",
                "finished_at",
                "session_end_timestamp",
            ),
        )
    )
    duration_ms = _fmcache_duration_ms(record)

    if start is None and end is not None and duration_ms is not None:
        start = end - timedelta(milliseconds=duration_ms)
    if end is None and start is not None and duration_ms is not None:
        end = start + timedelta(milliseconds=duration_ms)
    if start is None or end is None:
        return None
    if end <= start:
        return None

    span_ms = int((end - start).total_seconds() * 1000)
    if span_ms <= 0:
        return None
    if duration_ms is None:
        duration_ms = span_ms
    if duration_ms <= 0 or duration_ms > FMCACHE_MAX_SESSION_MS:
        return None
    return start, end, duration_ms


def _fmcache_daily_totals(
    sessions: Sequence[Tuple[datetime, datetime, int]],
) -> Dict[str, int]:
    totals: Dict[str, int] = {}
    for start, end, duration_ms in sessions:
        span_ms = max(1, int((end - start).total_seconds() * 1000))
        cursor = start
        while cursor < end:
            next_midnight = datetime.combine(
                cursor.date() + timedelta(days=1),
                datetime.min.time(),
            )
            segment_end = min(end, next_midnight)
            segment_ms = max(
                0,
                int((segment_end - cursor).total_seconds() * 1000),
            )
            portion = int(
                round(duration_ms * segment_ms / float(span_ms))
            )
            if portion:
                day_key = cursor.date().isoformat()
                totals[day_key] = totals.get(day_key, 0) + portion
            cursor = segment_end
    return totals


def read_fmcache_sessions(db_path: Optional[Path]) -> Dict[str, Any]:
    """Read candidate reading sessions from fmcache.db in read-only mode."""

    result: Dict[str, Any] = {
        "path": str(db_path or ""),
        "status": "missing",
        "message": "fmcache.db 未找到",
        "session_count": 0,
        "invalid_count": 0,
        "book_count": 0,
        "books": {},
        "daily": {},
    }
    if db_path is None:
        return result

    path = Path(db_path)
    try:
        if not path.is_file():
            return result
    except OSError:
        return result

    connection: Optional[sqlite3.Connection] = None
    try:
        uri = path.resolve().as_uri() + "?mode=ro"
        connection = sqlite3.connect(uri, uri=True, timeout=1.0)
        connection.execute("PRAGMA query_only = ON")

        table_rows = connection.execute(
            "SELECT name FROM sqlite_master WHERE type = 'table'"
        ).fetchall()
        tables = {str(row[0]) for row in table_rows}

        def columns(table: str) -> set[str]:
            quoted = '"' + table.replace('"', '""') + '"'
            return {
                str(row[1])
                for row in connection.execute(
                    f"PRAGMA table_info({quoted})"
                ).fetchall()
            }

        if "reading_sessions" not in tables:
            result["status"] = "schema_mismatch"
            result["message"] = "fmcache.db 缺少 reading_sessions 表"
            return result

        session_columns = columns("reading_sessions")
        if not {"id", "session"}.issubset(session_columns):
            result["status"] = "schema_mismatch"
            result["message"] = "reading_sessions 表结构不匹配"
            return result

        candidates: List[Tuple[str, Any, Dict[str, Any]]] = []
        row_count = 0
        for row in connection.execute(
            "SELECT id, session FROM reading_sessions ORDER BY id"
        ):
            row_count += 1
            decoded = _fmcache_json_object(row[1])
            if decoded is None:
                result["invalid_count"] += 1
                continue
            candidates.append(("reading_sessions", row[0], decoded))

        if not candidates and "records" in tables:
            record_columns = columns("records")
            required = {
                "id",
                "schema_name",
                "reading_session_id",
                "record",
            }
            if required.issubset(record_columns):
                for row in connection.execute(
                    "SELECT id, schema_name, reading_session_id, record "
                    "FROM records ORDER BY id"
                ):
                    schema_name = str(row[1] or "").lower()
                    reading_session_id = row[2]
                    if (
                        reading_session_id is None
                        and "reading" not in schema_name
                    ):
                        continue
                    decoded = _fmcache_json_object(row[3])
                    if decoded is None:
                        result["invalid_count"] += 1
                        continue
                    candidates.append(
                        ("records", row[0], decoded)
                    )

        sessions: List[Tuple[datetime, datetime, int]] = []
        book_sessions: Dict[str, Dict[str, Any]] = {}
        for _, record_id, decoded in candidates:
            parsed = _fmcache_session_times(decoded)
            if parsed is None:
                result["invalid_count"] += 1
                continue
            sessions.append(parsed)
            asin = _fmcache_asin(decoded)
            if not asin:
                continue
            start, end, duration_ms = parsed
            book = book_sessions.setdefault(
                asin,
                {
                    "asin": asin,
                    "duration_ms": 0,
                    "start": start.isoformat(),
                    "end": end.isoformat(),
                    "session_count": 0,
                    "progress_percent": None,
                    "location": None,
                },
            )
            book["duration_ms"] += duration_ms
            book["session_count"] += 1
            if start.isoformat() < str(book.get("start") or ""):
                book["start"] = start.isoformat()
            if end.isoformat() > str(book.get("end") or ""):
                book["end"] = end.isoformat()
            progress = _fmcache_progress_percent(decoded)
            if progress is not None:
                current = parse_float(book.get("progress_percent"))
                book["progress_percent"] = (
                    progress if current is None else max(current, progress)
                )
            location = _fmcache_location(decoded)
            if location is not None:
                current_location = parse_int(book.get("location"))
                book["location"] = (
                    location
                    if current_location is None
                    else max(current_location, location)
                )

        if "records" in tables:
            record_columns = columns("records")
            if {"schema_name", "record"}.issubset(record_columns):
                for row in connection.execute(
                    "SELECT schema_name, record FROM records ORDER BY id"
                ):
                    schema_name = str(row[0] or "").lower()
                    if "reading" not in schema_name:
                        continue
                    decoded = _fmcache_json_object(row[1])
                    if decoded is None:
                        continue
                    asin = _fmcache_asin(decoded)
                    progress = _fmcache_progress_percent(decoded)
                    if not asin or progress is None:
                        continue
                    book = book_sessions.setdefault(
                        asin,
                        {
                            "asin": asin,
                            "duration_ms": 0,
                            "start": "",
                            "end": "",
                            "session_count": 0,
                            "progress_percent": None,
                            "location": None,
                        },
                    )
                    current = parse_float(book.get("progress_percent"))
                    book["progress_percent"] = (
                        progress
                        if current is None
                        else max(current, progress)
                    )

        if not sessions:
            result["status"] = "empty"
            result["message"] = (
                "没有有效阅读会话"
                if row_count or candidates
                else "reading_sessions 为空"
            )
            return result

        result["status"] = "ok"
        result["message"] = f"已读取 {len(sessions)} 条阅读会话"
        result["session_count"] = len(sessions)
        result["book_count"] = len(book_sessions)
        result["books"] = book_sessions
        result["daily"] = _fmcache_daily_totals(sessions)
        return result
    except (OSError, sqlite3.Error, ValueError) as exc:
        result["status"] = "unreadable"
        result["message"] = f"fmcache.db 不可读取: {exc}"
        result["daily"] = {}
        return result
    finally:
        if connection is not None:
            connection.close()


def _row_fmcache_aliases(row: Dict[str, Any]) -> List[str]:
    values: List[str] = []
    asin = _normalize_mode_asin(row.get("asin"))
    if asin:
        values.append(asin)
    for key in ("ebook_file", "source_file", "title"):
        text = str(row.get(key) or "").strip()
        if not text:
            continue
        if key == "title":
            stem = text
        else:
            stem = os.path.splitext(
                text.replace("\\", "/").rsplit("/", 1)[-1]
            )[0]
        normalized_stem = _normalize_mode_asin(stem)
        if normalized_stem:
            values.append(normalized_stem)
            values.append(
                "lpath:" + _strip_file_extension(normalized_stem.casefold())
            )
        normalized = normalized_title(stem)
        if normalized:
            values.append("title:" + normalized)
        values.extend(
            match.upper()
            for match in re.findall(r"[0-9a-fA-F]{24,}", text)
        )
    result: List[str] = []
    for value in values:
        if value and value not in result:
            result.append(value)
    return result


def merge_fmcache_book_fallbacks(
    rows: Sequence[Dict[str, Any]],
    fmcache: Optional[Dict[str, Any]],
    book_modes: Optional[Dict[str, Any]] = None,
) -> Tuple[List[Dict[str, Any]], List[str]]:
    """Fill only manga rows whose sidecar time/progress is missing or zero."""

    merged = [dict(row) for row in rows]
    fmcache_data = fmcache if isinstance(fmcache, dict) else {}
    book_sessions = fmcache_data.get("books")
    if not isinstance(book_sessions, dict) or not book_sessions:
        return merged, []

    applied: List[str] = []
    for row in merged:
        mode = reading_mode_for_book(row, book_modes)
        if mode != "manga":
            continue
        row["reading_mode"] = "manga"
        aliases = _row_fmcache_aliases(row)
        matched_key = next(
            (
                alias
                for alias in aliases
                if isinstance(book_sessions.get(alias), dict)
            ),
            None,
        )
        if matched_key is None:
            continue
        session = book_sessions.get(matched_key)
        if not isinstance(session, dict):
            continue
        changed = False
        matched_existing_fallback = False

        duration_ms = parse_int(session.get("duration_ms")) or 0
        current_time_ms = parse_int(row.get("total_time_ms")) or 0
        if current_time_ms <= 0 and duration_ms > 0:
            row["total_time_ms"] = str(duration_ms)
            minutes = duration_ms / 60000.0
            row["minutes"] = str(int(round(minutes)))
            row["hours"] = f"{minutes / 60.0:.2f}"
            row["has_data"] = "True"
            changed = True
        elif duration_ms > 0 and current_time_ms == duration_ms:
            matched_existing_fallback = True

        progress = parse_float(session.get("progress_percent"))
        total_pages = manga_total_pages(row)
        location = parse_int(session.get("location"))
        if (
            progress is None
            and total_pages
            and location is not None
            and 0 < location <= total_pages
        ):
            progress = round(location / total_pages * 100.0, 1)
        current_progress = parse_float(row.get("progress_percent"))
        if (
            (current_progress is None or current_progress <= 0.0)
            and progress is not None
            and progress > 0.0
        ):
            row["progress_percent"] = f"{progress:g}"
            row["progress"] = str(int(round(progress)))
            if total_pages:
                row["manga_pages_read"] = str(
                    max(
                        0,
                        min(
                            int(total_pages),
                            int(round(total_pages * progress / 100.0)),
                        ),
                    )
                )
            row["manga_progress_source"] = "fmcache"
            row["has_data"] = "True"
            changed = True
        elif (
            progress is not None
            and progress > 0.0
            and current_progress is not None
            and abs(current_progress - progress) < 1e-9
        ):
            matched_existing_fallback = True

        if not str(row.get("last_read") or "").strip() and session.get("end"):
            row["last_read"] = str(session.get("end"))
            changed = True

        existing_flag = str(
            row.get("fmcache_fallback") or ""
        ).strip().casefold() in {"1", "true", "yes"}
        if changed or matched_existing_fallback or existing_flag:
            row["_fmcache_fallback"] = True
            row["fmcache_fallback"] = "True"
            if changed or (matched_existing_fallback and not existing_flag):
                applied.append(
                    _normalize_mode_asin(row.get("asin")) or matched_key
                )
    return merged, applied


def merge_history_book_fallbacks(
    rows: Sequence[Dict[str, Any]],
    history: Optional[Dict[str, Any]],
    book_modes: Optional[Dict[str, Any]] = None,
) -> Tuple[List[Dict[str, Any]], List[str]]:
    """Keep per-book maxima from prior snapshots when fmcache rotates."""

    merged = [dict(row) for row in rows]
    snapshots = [
        snapshot
        for snapshot in ((history or {}).get("snapshots") or [])
        if isinstance(snapshot, dict)
    ]
    historical: Dict[str, Dict[str, Any]] = {}
    for snapshot in snapshots:
        snapshot_version = str(snapshot.get("data_source_version") or "")
        for book in (snapshot.get("books") or {}).values():
            if not isinstance(book, dict):
                continue
            if (
                reading_mode_for_book(book) == "manga"
                and snapshot_version != VERSION
                and not any(
                    book.get(key) not in (None, "")
                    for key in (
                        "manga_pages_total",
                        "manga_pages_read",
                        "manga_progress_source",
                    )
                )
                and (parse_int(book.get("total_time_ms")) or 0) <= 0
                and (parse_float(book.get("progress_percent")) or 0.0) <= 0.0
            ):
                # Older snapshots could contain OCR-derived placeholder
                # values for manga. Only trust legacy manga values when they
                # carry explicit page/progress provenance.
                continue
            aliases = _row_fmcache_aliases(book)
            for alias in aliases:
                current = historical.get(alias)
                if current is None or (
                    parse_int(book.get("total_time_ms")) or 0
                ) > (parse_int(current.get("total_time_ms")) or 0):
                    historical[alias] = book

    applied: List[str] = []
    for row in merged:
        mode = reading_mode_for_book(row, book_modes)
        if mode != "manga":
            continue
        row["reading_mode"] = "manga"
        previous = None
        for alias in [stable_book_id(row)] + _row_fmcache_aliases(row):
            previous = historical.get(alias)
            if isinstance(previous, dict):
                break
        if not isinstance(previous, dict):
            continue
        changed = False
        current_time = parse_int(row.get("total_time_ms")) or 0
        previous_time = parse_int(previous.get("total_time_ms")) or 0
        if previous_time > current_time:
            row["total_time_ms"] = str(previous_time)
            minutes = previous_time / 60000.0
            row["minutes"] = str(int(round(minutes)))
            row["hours"] = f"{minutes / 60.0:.2f}"
            row["has_data"] = "True"
            changed = True

        current_progress = parse_float(row.get("progress_percent"))
        previous_progress = parse_float(previous.get("progress_percent"))
        if (
            previous_progress is not None
            and previous_progress > 0.0
            and (
                current_progress is None
                or previous_progress > current_progress
            )
        ):
            row["progress_percent"] = f"{previous_progress:g}"
            row["progress"] = str(int(round(previous_progress)))
            row["manga_progress_source"] = "history"
            total_pages = manga_total_pages(row)
            if total_pages:
                row["manga_pages_read"] = str(
                    max(
                        0,
                        min(
                            total_pages,
                            int(round(total_pages * previous_progress / 100.0)),
                        ),
                    )
                )
            changed = True
        for key in ("manga_pages_total", "manga_pages_read", "manga_pages_source"):
            previous_value = previous.get(key)
            if previous_value not in (None, "") and not str(
                row.get(key) or ""
            ).strip():
                row[key] = previous_value
                changed = True
        previous_last_read = str(previous.get("last_read") or "")
        if previous_last_read and not str(row.get("last_read") or "").strip():
            row["last_read"] = previous_last_read
            changed = True
        if changed:
            applied.append(stable_book_id(row))
    return merged, applied


def _heatmap_rows(
    daily: Any,
    end_date: date,
    days: int,
) -> List[Dict[str, Any]]:
    if isinstance(daily, dict):
        totals_by_day = {
            str(key): parse_int(value) or 0
            for key, value in daily.items()
        }
    else:
        totals_by_day = {
            str(row.get("date") or ""): parse_int(
                row.get("total_time_ms")
            )
            or 0
            for row in (daily or [])
            if isinstance(row, dict) and row.get("date")
        }

    result: List[Dict[str, Any]] = []
    for offset in range(days - 1, -1, -1):
        day_key = (end_date - timedelta(days=offset)).isoformat()
        total_time_ms = totals_by_day.get(day_key, 0)
        result.append(
            {
                "date": day_key,
                "total_time_ms": total_time_ms,
                "active": total_time_ms > 0,
            }
        )
    return result


def build_heatmap_data(
    history: Dict[str, Any],
    current_snapshot: Dict[str, Any],
    fmcache: Optional[Dict[str, Any]] = None,
    days: int = FMCACHE_HEATMAP_DAYS,
) -> Dict[str, Any]:
    snapshots = [
        snapshot
        for snapshot in history.get("snapshots", [])
        if isinstance(snapshot, dict)
    ]
    valid_history_dates = [
        snapshot_date
        for snapshot_date in (_snapshot_date(item) for item in snapshots)
        if snapshot_date is not None
    ]
    end_date = _snapshot_date(current_snapshot) or (
        valid_history_dates[-1] if valid_history_dates else date.today()
    )
    history_daily = _daily_window_rows(
        _daily_reading_totals(snapshots + [current_snapshot]),
        end_date,
        days,
    )

    fmcache_data = fmcache if isinstance(fmcache, dict) else {}
    fmcache_available = (
        fmcache_data.get("status") == "ok"
        and (parse_int(fmcache_data.get("session_count")) or 0) > 0
    )
    if fmcache_available:
        daily = _heatmap_rows(
            fmcache_data.get("daily") or {},
            end_date,
            days,
        )
        source = "fmcache"
        label = "fmcache.db 阅读会话"
    else:
        daily = _heatmap_rows(history_daily, end_date, days)
        source = "history"
        label = "历史快照回退"

    return {
        "days": days,
        "start": daily[0]["date"] if daily else end_date.isoformat(),
        "end": daily[-1]["date"] if daily else end_date.isoformat(),
        "source": source,
        "source_label": label,
        "status": str(fmcache_data.get("status") or "missing"),
        "message": str(fmcache_data.get("message") or ""),
        "db_path": str(fmcache_data.get("path") or ""),
        "session_count": parse_int(
            fmcache_data.get("session_count")
        )
        or 0,
        "daily": daily,
    }


def _snapshot_period_datetime(snapshot: Dict[str, Any]) -> Optional[datetime]:
    parsed = _snapshot_datetime(snapshot)
    if parsed is None:
        return None
    if parsed.tzinfo is not None:
        parsed = parsed.replace(tzinfo=None)
    return parsed


def _period_bounds(reference_date: date, period_key: str) -> Tuple[date, date]:
    if period_key == "month":
        return reference_date.replace(day=1), reference_date
    if period_key == "year":
        return reference_date.replace(month=1, day=1), reference_date
    return (
        reference_date - timedelta(days=reference_date.weekday()),
        reference_date,
    )


def _period_daily_allocation(
    start: datetime,
    end: datetime,
    span_seconds: float,
    time_delta_ms: float,
    words_delta: float,
) -> Dict[str, Dict[str, float]]:
    result: Dict[str, Dict[str, float]] = {}
    if span_seconds <= 0 or end <= start:
        return result

    cursor = start
    while cursor < end:
        next_midnight = datetime.combine(
            cursor.date() + timedelta(days=1),
            datetime.min.time(),
        )
        segment_end = min(end, next_midnight)
        segment_seconds = max(
            0.0,
            (segment_end - cursor).total_seconds(),
        )
        portion = segment_seconds / span_seconds
        day_key = cursor.date().isoformat()
        bucket = result.setdefault(
            day_key,
            {"total_time_ms": 0.0, "total_words": 0.0},
        )
        bucket["total_time_ms"] += time_delta_ms * portion
        bucket["total_words"] += words_delta * portion
        cursor = segment_end
    return result


def _period_coverage_seconds(
    intervals: Sequence[Tuple[datetime, datetime]],
) -> Tuple[float, Optional[datetime], Optional[datetime]]:
    ordered = sorted(
        (start, end)
        for start, end in intervals
        if end > start
    )
    if not ordered:
        return 0.0, None, None

    merged: List[Tuple[datetime, datetime]] = []
    for start, end in ordered:
        if not merged or start > merged[-1][1]:
            merged.append((start, end))
            continue
        previous_start, previous_end = merged[-1]
        merged[-1] = (previous_start, max(previous_end, end))

    seconds = sum(
        max(0.0, (end - start).total_seconds())
        for start, end in merged
    )
    return seconds, merged[0][0], merged[-1][1]


def _period_date_range(start: date, end: date) -> str:
    if start == end:
        return start.isoformat()
    return f"{start.isoformat()} 至 {end.isoformat()}"


def _period_date_range_en(start: date, end: date) -> str:
    if start == end:
        return start.isoformat()
    return f"{start.isoformat()} to {end.isoformat()}"


def build_period_reports(
    current_snapshot: Dict[str, Any],
    history: Dict[str, Any],
    fmcache: Optional[Dict[str, Any]] = None,
    reference_date: Optional[date] = None,
) -> Dict[str, Any]:
    """Build independent current week, month, and year reports."""

    history_snapshots = [
        snapshot
        for snapshot in history.get("snapshots", [])
        if isinstance(snapshot, dict)
    ]
    valid_history = [
        (snapshot, _snapshot_period_datetime(snapshot))
        for snapshot in history_snapshots
    ]
    valid_history = [
        (snapshot, captured)
        for snapshot, captured in valid_history
        if captured is not None
    ]
    current_captured = _snapshot_period_datetime(current_snapshot)
    combined: List[Tuple[Dict[str, Any], datetime]] = list(valid_history)
    if current_captured is not None:
        combined.append((current_snapshot, current_captured))
    combined.sort(key=lambda item: item[1])
    observed_history_daily = _daily_reading_totals(
        [snapshot for snapshot, _ in combined]
    )

    transitions: List[Dict[str, Any]] = []
    for previous_index in range(len(combined) - 1):
        previous_snapshot, previous_dt = combined[previous_index]
        current_item, current_dt = combined[previous_index + 1]
        span_seconds = (current_dt - previous_dt).total_seconds()
        if span_seconds <= 0:
            continue

        previous_totals = snapshot_totals(previous_snapshot)
        current_totals = snapshot_totals(current_item)
        previous_time = parse_int(previous_totals.get("total_time_ms")) or 0
        current_time = parse_int(current_totals.get("total_time_ms")) or 0
        previous_words = (
            parse_int(previous_totals.get("effective_words_read")) or 0
        )
        current_words = (
            parse_int(current_totals.get("effective_words_read")) or 0
        )
        total_reset = current_time < previous_time or current_words < previous_words

        previous_books = previous_snapshot.get("books") or {}
        current_books = current_item.get("books") or {}
        if not isinstance(previous_books, dict):
            previous_books = {}
        if not isinstance(current_books, dict):
            current_books = {}

        books: List[Dict[str, Any]] = []
        for book_id in set(previous_books) | set(current_books):
            previous_book = previous_books.get(book_id) or {}
            current_book = current_books.get(book_id) or {}
            if not isinstance(previous_book, dict):
                previous_book = {}
            if not isinstance(current_book, dict):
                current_book = {}
            previous_book_time = (
                parse_int(previous_book.get("total_time_ms")) or 0
            )
            current_book_time = parse_int(current_book.get("total_time_ms")) or 0
            previous_book_words = (
                _snapshot_book_effective_words_read(previous_book) or 0
            )
            current_book_words = (
                _snapshot_book_effective_words_read(current_book) or 0
            )
            reading_mode = reading_mode_for_book(
                current_book if current_book else previous_book
            )
            word_source = str(
                current_book.get("effective_source")
                or previous_book.get("effective_source")
                or "kindle-raw"
            )
            time_delta = max(0, current_book_time - previous_book_time)
            words_delta = max(0, current_book_words - previous_book_words)
            reset_detected = (
                current_book_time < previous_book_time
                or current_book_words < previous_book_words
            )
            if not (time_delta or words_delta or reset_detected):
                continue
            books.append(
                {
                    "book_id": str(book_id),
                    "title": str(
                        current_book.get("title")
                        or previous_book.get("title")
                        or "(unknown)"
                    ),
                    "author": str(
                        current_book.get("author")
                        or previous_book.get("author")
                        or ""
                    ),
                    "time_delta_ms": time_delta,
                    "words_delta": words_delta,
                    "reset_detected": reset_detected,
                    "reading_mode": reading_mode,
                    "word_source": word_source,
                }
            )

        transitions.append(
            {
                "previous_at": str(previous_snapshot.get("captured_at") or ""),
                "current_at": str(current_item.get("captured_at") or ""),
                "previous_dt": previous_dt,
                "current_dt": current_dt,
                "span_seconds": span_seconds,
                "gap_days": round(span_seconds / 86400.0, 3),
                "time_delta_ms": max(0, current_time - previous_time),
                "words_delta": max(0, current_words - previous_words),
                "reset_detected": total_reset,
                "books": books,
            }
        )

    reference = reference_date or _snapshot_date(current_snapshot) or date.today()
    fmcache_data = fmcache if isinstance(fmcache, dict) else {}
    fmcache_daily = {
        str(day): parse_int(value) or 0
        for day, value in (fmcache_data.get("daily") or {}).items()
    }
    fmcache_available = (
        fmcache_data.get("status") == "ok"
        and (parse_int(fmcache_data.get("session_count")) or 0) > 0
    )

    reports: List[Dict[str, Any]] = []
    for period_key in PERIOD_REPORT_KEYS:
        start_date, end_date = _period_bounds(reference, period_key)
        start_dt = datetime.combine(start_date, datetime.min.time())
        end_dt = datetime.combine(
            end_date + timedelta(days=1),
            datetime.min.time(),
        )
        period_seconds = max(1.0, (end_dt - start_dt).total_seconds())
        gap_limit = PERIOD_GAP_LIMIT_DAYS[period_key]

        coverage_intervals: List[Tuple[datetime, datetime]] = []
        history_daily: Dict[str, Dict[str, float]] = {}
        book_totals: Dict[str, Dict[str, Any]] = {}
        observation_count = 0
        large_gaps: List[Dict[str, Any]] = []
        reset_count = 0

        for transition in transitions:
            previous_dt = transition["previous_dt"]
            current_dt = transition["current_dt"]
            overlap_start = max(previous_dt, start_dt)
            overlap_end = min(current_dt, end_dt)
            overlap_seconds = (overlap_end - overlap_start).total_seconds()
            if overlap_seconds <= 0:
                continue

            ratio = overlap_seconds / transition["span_seconds"]
            observation_count += 1
            coverage_intervals.append((overlap_start, overlap_end))
            time_delta = transition["time_delta_ms"] * ratio
            words_delta = transition["words_delta"] * ratio
            if transition.get("reset_detected"):
                reset_count += 1
            if transition["gap_days"] > gap_limit:
                large_gaps.append(
                    {
                        "start": previous_dt.date().isoformat(),
                        "end": current_dt.date().isoformat(),
                        "days": transition["gap_days"],
                    }
                )

            allocated = _period_daily_allocation(
                overlap_start,
                overlap_end,
                transition["span_seconds"],
                transition["time_delta_ms"],
                transition["words_delta"],
            )
            for day_key, values in allocated.items():
                bucket = history_daily.setdefault(
                    day_key,
                    {"total_time_ms": 0.0, "total_words": 0.0},
                )
                bucket["total_time_ms"] += values["total_time_ms"]
                bucket["total_words"] += values["total_words"]

            for book in transition["books"]:
                book_id = str(book.get("book_id") or "")
                entry = book_totals.setdefault(
                    book_id,
                    {
                        "title": str(book.get("title") or "(unknown)"),
                        "author": str(book.get("author") or ""),
                        "reading_mode": str(
                            book.get("reading_mode") or "book"
                        ),
                        "time_ms": 0.0,
                        "words": 0.0,
                        "reset": False,
                        "word_source": str(
                            book.get("word_source") or "kindle-raw"
                        ),
                    },
                )
                if str(book.get("word_source") or "") == "ocr-estimate":
                    entry["word_source"] = "ocr-estimate"
                entry["time_ms"] += (
                    parse_int(book.get("time_delta_ms")) or 0
                ) * ratio
                entry["words"] += (
                    parse_int(book.get("words_delta")) or 0
                ) * ratio
                entry["reset"] = bool(
                    entry["reset"] or book.get("reset_detected")
                )

        coverage_seconds, coverage_start, coverage_end = (
            _period_coverage_seconds(coverage_intervals)
        )
        coverage_ratio = max(
            0.0,
            min(1.0, coverage_seconds / period_seconds),
        )

        fmcache_period_daily = {
            day: value
            for day, value in fmcache_daily.items()
            if start_date.isoformat() <= day <= end_date.isoformat()
        }
        fmcache_period_has_data = any(
            value > 0 for value in fmcache_period_daily.values()
        )
        if fmcache_available and fmcache_period_has_data:
            history_fallback_daily = {
                day: int(round(values.get("total_time_ms") or 0))
                for day, values in observed_history_daily.items()
                if start_date.isoformat() <= day <= end_date.isoformat()
            }
            fmcache_daily_current = {
                day: int(value)
                for day, value in fmcache_period_daily.items()
            }
            daily = {
                day: value
                for day, value in history_fallback_daily.items()
                if value > 0 and day not in fmcache_daily_current
            }
            daily.update(fmcache_daily_current)
            has_history_fallback = any(
                value > 0
                for day, value in history_fallback_daily.items()
                if day not in fmcache_daily_current
            )
            if has_history_fallback:
                source = "mixed"
                source_label_zh = "fmcache 会话 + 历史快照"
                source_label_en = "fmcache sessions + history snapshots"
            else:
                source = "fmcache"
                source_label_zh = "fmcache.db 阅读会话"
                source_label_en = "fmcache.db reading sessions"
        else:
            source = "history"
            daily = {
                day: int(round(values.get("total_time_ms") or 0))
                for day, values in history_daily.items()
            }
            source_label_zh = "历史快照回退"
            source_label_en = "History snapshot fallback"

        total_time_ms = sum(max(0, value) for value in daily.values())
        total_words = sum(
            max(0.0, float(values.get("total_words") or 0.0))
            for day, values in history_daily.items()
            if start_date.isoformat() <= day <= end_date.isoformat()
        )
        reading_days = sum(1 for value in daily.values() if value > 0)
        period_days = max(1, (end_date - start_date).days + 1)
        daily_average_ms = int(round(total_time_ms / period_days))

        books: List[Dict[str, Any]] = []
        speed_candidates: List[Dict[str, Any]] = []
        for book_id, entry in book_totals.items():
            time_ms = int(round(entry.get("time_ms") or 0))
            words = int(round(entry.get("words") or 0))
            reading_mode = normalize_reading_mode(
                entry.get("reading_mode")
            )
            if time_ms <= 0 and words <= 0:
                continue
            word_source = str(entry.get("word_source") or "kindle-raw")
            speed = (
                words / (time_ms / 60000.0)
                if time_ms > 0 and words > 0
                else None
            )
            valid_speed = (
                speed is not None
                and time_ms >= PERIOD_MIN_SPEED_TIME_MS
                and words >= PERIOD_MIN_SPEED_WORDS
                and PERIOD_WPM_NORMAL_RANGE[0]
                <= speed
                <= PERIOD_WPM_NORMAL_RANGE[1]
                and not entry.get("reset")
            )
            book = {
                "book_id": book_id,
                "title": str(entry.get("title") or "(unknown)"),
                "author": str(entry.get("author") or ""),
                "reading_mode": reading_mode,
                "word_source": word_source,
                "time_ms": time_ms,
                "words": words,
                "wpm": round(speed, 1) if speed is not None else None,
                "valid_speed": bool(valid_speed),
            }
            books.append(book)
            if valid_speed:
                speed_candidates.append(book)

        books.sort(
            key=lambda item: (
                -(parse_int(item.get("time_ms")) or 0),
                -(parse_int(item.get("words")) or 0),
                str(item.get("title") or "").casefold(),
            )
        )
        active_books = len(books)
        fastest = max(
            speed_candidates,
            key=lambda item: parse_float(item.get("wpm")) or 0.0,
            default=None,
        )

        if period_key == "year":
            trend: List[Dict[str, Any]] = []
            month_abbr = (
                "Jan Feb Mar Apr May Jun "
                "Jul Aug Sep Oct Nov Dec"
            ).split()
            for month in range(1, end_date.month + 1):
                month_key = f"{end_date.year:04d}-{month:02d}"
                month_total = sum(
                    value
                    for day, value in daily.items()
                    if day.startswith(month_key + "-")
                )
                trend.append(
                    {
                        "bucket": month_key,
                        "label_zh": f"{month} 月",
                        "label_en": (
                            f"{month:02d} "
                            + month_abbr[month - 1]
                        ),
                        "label_short_zh": f"{month}月",
                        "label_short_en": month_abbr[month - 1],
                        "total_time_ms": int(month_total),
                        "active": month_total > 0,
                    }
                )
        else:
            trend = []
            weekday_zh = "一二三四五六日"
            weekday_en = ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
            cursor = start_date
            while cursor <= end_date:
                day_key = cursor.isoformat()
                day_total = int(daily.get(day_key, 0))
                trend.append(
                    {
                        "bucket": day_key,
                        "label_zh": (
                            f"周{weekday_zh[cursor.weekday()]} "
                            f"{cursor.strftime('%m-%d')}"
                        ),
                        "label_en": (
                            f"{weekday_en[cursor.weekday()]} "
                            f"{cursor.strftime('%m-%d')}"
                        ),
                        "label_short_zh": (
                            f"周{weekday_zh[cursor.weekday()]}"
                            if period_key == "week"
                            else str(cursor.day)
                        ),
                        "label_short_en": (
                            weekday_en[cursor.weekday()]
                            if period_key == "week"
                            else str(cursor.day)
                        ),
                        "total_time_ms": day_total,
                        "active": day_total > 0,
                    }
                )
                cursor += timedelta(days=1)

        missing_segments: List[Dict[str, Any]] = []
        if coverage_start is None or coverage_end is None:
            missing_segments.append(
                {
                    "start": start_date.isoformat(),
                    "end": end_date.isoformat(),
                    "days": float(period_days),
                    "kind": "coverage",
                }
            )
        else:
            coverage_start_date = coverage_start.date()
            coverage_end_date = coverage_end.date()
            if coverage_start_date > start_date:
                missing_end = coverage_start_date - timedelta(days=1)
                missing_segments.append(
                    {
                        "start": start_date.isoformat(),
                        "end": missing_end.isoformat(),
                        "days": float(
                            (missing_end - start_date).days + 1
                        ),
                        "kind": "coverage",
                    }
                )
            if coverage_end_date < end_date:
                missing_start = coverage_end_date + timedelta(days=1)
                missing_segments.append(
                    {
                        "start": missing_start.isoformat(),
                        "end": end_date.isoformat(),
                        "days": float(
                            (end_date - missing_start).days + 1
                        ),
                        "kind": "coverage",
                    }
                )
        for gap in large_gaps:
            missing_segments.append(
                {
                    "start": gap["start"],
                    "end": gap["end"],
                    "days": gap["days"],
                    "kind": "snapshot-gap",
                }
            )

        if observation_count < 2 or coverage_ratio < 0.35:
            confidence = "low"
        elif coverage_ratio < 0.80 or large_gaps:
            confidence = "medium"
        else:
            confidence = "high"
        reliable = confidence in {"high", "medium"}

        if reliable:
            conclusion_zh = "这个周期已经有清晰的阅读积累。"
            conclusion_en = "This period already has a clear reading rhythm."
        else:
            conclusion_zh = "继续阅读，积累更多记录后这里会更完整。"
            conclusion_en = (
                "Keep reading; this view becomes fuller as more history "
                "builds up."
            )

        if source == "fmcache":
            source_note_zh = (
                "总时长、阅读天数和趋势使用 fmcache 会话；书籍与速度来自历史快照。"
            )
            source_note_en = (
                "Time, reading days, and trend use fmcache sessions; books "
                "and speed use history snapshots."
            )
        elif source == "mixed":
            source_note_zh = (
                "有 fmcache 会话的日期使用会话时长；没有会话的日期回退到"
                "历史快照增量。"
            )
            source_note_en = (
                "Dates with fmcache sessions use session time; dates without "
                "sessions fall back to history-snapshot deltas."
            )
        else:
            source_note_zh = (
                "fmcache 在本周期没有可用会话，总时长和趋势回退到历史快照增量。"
            )
            source_note_en = (
                "No usable fmcache session was found for this period; time "
                "and trend fall back to history-snapshot deltas."
            )

        if missing_segments:
            missing_zh = "；".join(
                (
                    f"{row['start']} 至 {row['end']} "
                    + (
                        "记录不足"
                        if row.get("kind") == "snapshot-gap"
                        else "无覆盖"
                    )
                )
                for row in missing_segments
            )
            missing_en = "; ".join(
                (
                    f"{row['start']} to {row['end']} "
                    + (
                        "sparse snapshots"
                        if row.get("kind") == "snapshot-gap"
                        else "no coverage"
                    )
                )
                for row in missing_segments
            )
        else:
            missing_zh = "未发现明显的缺失时间段。"
            missing_en = "No obvious missing interval was detected."

        if coverage_start is None or coverage_end is None:
            coverage_zh = "没有可用于周期判断的有效快照区间。"
            coverage_en = "No valid snapshot interval is available."
        else:
            coverage_zh = (
                f"有效覆盖 {coverage_start.date().isoformat()} 至 "
                f"{coverage_end.date().isoformat()}，"
                f"约占周期 {coverage_ratio * 100:.0f}%。"
            )
            coverage_en = (
                f"Usable coverage {coverage_start.date().isoformat()} to "
                f"{coverage_end.date().isoformat()}, about "
                f"{coverage_ratio * 100:.0f}% of the period."
            )

        reports.append(
            {
                "key": period_key,
                "label_zh": {
                    "week": "本周",
                    "month": "本月",
                    "year": "本年",
                }[period_key],
                "label_en": {
                    "week": "This week",
                    "month": "This month",
                    "year": "This year",
                }[period_key],
                "range_start": start_date.isoformat(),
                "range_end": end_date.isoformat(),
                "range_zh": _period_date_range(start_date, end_date),
                "range_en": _period_date_range_en(start_date, end_date),
                "total_time_ms": total_time_ms,
                "total_words": round(total_words, 2),
                "ocr_estimate_books": sum(
                    1 for book in books if book.get("word_source") == "ocr-estimate"
                ),
                "reading_days": reading_days,
                "period_days": period_days,
                "daily_average_ms": daily_average_ms,
                "active_books": active_books,
                "fastest_wpm": (
                    parse_float(fastest.get("wpm"))
                    if fastest is not None
                    else None
                ),
                "fastest_title": (
                    str(fastest.get("title") or "")
                    if fastest is not None
                    else ""
                ),
                "fastest_word_source": (
                    str(fastest.get("word_source") or "")
                    if fastest is not None
                    else ""
                ),
                "source": source,
                "source_label_zh": source_label_zh,
                "source_label_en": source_label_en,
                "snapshot_count": len(valid_history) + 1,
                "observation_count": observation_count,
                "fmcache_session_count": (
                    parse_int(fmcache_data.get("session_count")) or 0
                    if fmcache_available
                    else 0
                ),
                "confidence": confidence,
                "reliable": reliable,
                "conclusion_zh": conclusion_zh,
                "conclusion_en": conclusion_en,
                "source_note_zh": source_note_zh,
                "source_note_en": source_note_en,
                "coverage_zh": coverage_zh,
                "coverage_en": coverage_en,
                "missing_zh": missing_zh,
                "missing_en": missing_en,
                "missing_segments": missing_segments,
                "large_gap_count": len(large_gaps),
                "reset_count": reset_count,
                "top_books": books[:5],
                "trend": trend,
            }
        )

    insufficient = sum(1 for report in reports if not report["reliable"])
    if insufficient:
        summary_zh = f"{insufficient}/3 个周期仍可继续积累"
        summary_en = f"{insufficient}/3 periods can keep building"
    else:
        summary_zh = "周、月、年报告都有清晰记录"
        summary_en = "Week, month, and year views are complete"
    return {
        "reports": reports,
        "summary_zh": summary_zh,
        "summary_en": summary_en,
    }


def _period_confidence_label(value: Any) -> Tuple[str, str, str]:
    normalized = str(value or "low").lower()
    if normalized == "high":
        return "high", "高确定性", "High confidence"
    if normalized == "medium":
        return "medium", "中等确定性", "Medium confidence"
    return "low", "低确定性", "Low confidence"


def build_period_report_section(data: Dict[str, Any]) -> str:
    reports = [
        report
        for report in (data.get("reports") or [])
        if isinstance(report, dict)
    ]
    if not reports:
        return (
            '<details class="period-panel" id="period-report-panel" open>'
            "<summary><span>"
            + bilingual_text("周期报告", "Period reports")
            + "</span><em>"
            + bilingual_text(
                "暂无可生成的周期报告。",
                "No period report can be generated yet.",
            )
            + "</em></summary></details>"
        )

    tabs: List[str] = []
    panels: List[str] = []
    for index, report in enumerate(reports):
        period_key = str(report.get("key") or "")
        if period_key not in PERIOD_REPORT_KEYS:
            continue
        selected = index == 0
        tab_id = f"period-tab-{period_key}"
        panel_id = f"period-panel-view-{period_key}"
        tabs.append(
            f'<button type="button" role="tab" id="{tab_id}" '
            f'data-period-tab="{period_key}" '
            f'aria-controls="{panel_id}" '
            f'aria-selected="{"true" if selected else "false"}" '
            f'tabindex="{"0" if selected else "-1"}">'
            f'<span class="i18n-zh">{html.escape(str(report.get("label_zh") or ""))}</span>'
            f'<span class="i18n-en">{html.escape(str(report.get("label_en") or ""))}</span>'
            "</button>"
        )

        confidence, confidence_zh, confidence_en = _period_confidence_label(
            report.get("confidence")
        )
        metrics = [
            (
                "时间范围",
                "Date range",
                str(report.get("range_zh") or ""),
                str(report.get("range_en") or ""),
            ),
            (
                "总时长",
                "Total time",
                format_duration(parse_int(report.get("total_time_ms")) or 0),
                format_duration_en(parse_int(report.get("total_time_ms")) or 0),
            ),
            (
                "阅读词数",
                "Words read",
                format_words(parse_int(report.get("total_words"))),
                format_words(parse_int(report.get("total_words"))),
            ),
            (
                "阅读天数",
                "Reading days",
                str(parse_int(report.get("reading_days")) or 0),
                str(parse_int(report.get("reading_days")) or 0),
            ),
            (
                "日均时长",
                "Daily average",
                format_duration(parse_int(report.get("daily_average_ms")) or 0),
                format_duration_en(
                    parse_int(report.get("daily_average_ms")) or 0
                ),
            ),
            (
                "活跃书籍",
                "Active books",
                str(parse_int(report.get("active_books")) or 0),
                str(parse_int(report.get("active_books")) or 0),
            ),
            (
                "最快阅读速度",
                "Fastest reading speed",
                (
                    f'{parse_float(report.get("fastest_wpm")):.1f} WPM'
                    if parse_float(report.get("fastest_wpm")) is not None
                    else "--"
                ),
                (
                    f'{parse_float(report.get("fastest_wpm")):.1f} WPM'
                    if parse_float(report.get("fastest_wpm")) is not None
                    else "--"
                ),
            ),
        ]
        metric_html = "".join(
            '<div class="period-metric">'
            f"<span>{bilingual_text(zh, en)}</span>"
            f"<strong>{bilingual_text(value_zh, value_en)}</strong>"
            "</div>"
            for zh, en, value_zh, value_en in metrics
        )

        trend = [
            row
            for row in (report.get("trend") or [])
            if isinstance(row, dict)
        ]
        max_trend = max(
            (parse_int(row.get("total_time_ms")) or 0 for row in trend),
            default=0,
        )
        trend_rows: List[str] = []
        for bucket_index, row in enumerate(trend):
            total_time_ms = parse_int(row.get("total_time_ms")) or 0
            period_color = PERIOD_CHART_COLORS[
                bucket_index % len(PERIOD_CHART_COLORS)
            ]
            height = (
                total_time_ms / max_trend * 100.0
                if max_trend > 0
                else 0.0
            )
            if total_time_ms > 0 and height < 3.0:
                height = 3.0
            state = "active" if total_time_ms > 0 else "no-record"
            tooltip_zh = (
                f"{row.get('label_zh')} · {format_duration(total_time_ms)}"
            )
            tooltip_en = (
                f"{row.get('label_en')} · {format_duration_en(total_time_ms)}"
            )
            short_label_zh = str(
                row.get("label_short_zh") or row.get("label_zh") or ""
            )
            short_label_en = str(
                row.get("label_short_en") or row.get("label_en") or ""
            )
            trend_rows.append(
                '<div class="period-bar" role="img" tabindex="0" '
                f'data-period-bucket="{html.escape(str(row.get("bucket") or ""), quote=True)}" '
                f'data-period-state="{state}" '
                f'data-period-value-ms="{total_time_ms}" '
                f'data-period-color="{bucket_index % len(PERIOD_CHART_COLORS)}" '
                f'data-title-zh="{html.escape(tooltip_zh, quote=True)}" '
                f'data-title-en="{html.escape(tooltip_en, quote=True)}" '
                f'title="{html.escape(tooltip_zh, quote=True)}" '
                f'aria-label="{html.escape(tooltip_zh, quote=True)}">'
                '<span class="period-bar-value">'
                '<span class="period-bar-value-full">'
                f"{bilingual_text(format_duration(total_time_ms), format_duration_en(total_time_ms))}"
                "</span>"
                f'<span class="period-bar-value-short">'
                f"{html.escape(_period_compact_duration(total_time_ms))}</span>"
                "</span>"
                '<span class="period-bar-track">'
                f'<i class="period-bar-fill" style="height:{height:.2f}%;'
                f'--period-color:{period_color}"></i>'
                '<span class="period-bar-empty" aria-hidden="true"></span>'
                "</span>"
                '<span class="period-bar-label">'
                f"{bilingual_text(short_label_zh, short_label_en)}"
                "</span>"
                "</div>"
            )
        if max_trend <= 0 or not trend_rows:
            trend_html = (
                '<div class="empty compact">'
                + bilingual_text(
                    "暂无可显示的周期趋势。继续阅读即可解锁。",
                    "No period trend is available yet. Keep reading to unlock.",
                )
                + "</div>"
            )
        else:
            bucket_count = len(trend_rows)
            min_bar_width = 32
            min_gap = 6
            min_width = (
                bucket_count * min_bar_width
                + max(0, bucket_count - 1) * min_gap
            )
            dense_class = " is-dense" if bucket_count > 12 else ""
            trend_html = (
                '<div class="period-trend-chart period-bar-chart'
                + dense_class
                + '" data-period-trend-chart="bars" '
                'data-chart-type="bar" '
                'data-period-grid="lines" '
                'data-period-baseline="true" '
                'data-period-bar-style="uniform" '
                f'data-period-bucket-count="{bucket_count}" '
                'role="group" aria-label="周期趋势柱状图 / Period trend bar chart" '
                f'style="--period-bucket-count:{bucket_count};'
                f'--period-bars-min-width:{min_width}px">'
                '<div class="period-chart-scroll">'
                '<div class="period-bars">'
                '<span class="period-grid-lines" aria-hidden="true"></span>'
                + "".join(trend_rows)
                + "</div></div></div>"
            )

        top_books = [
            row
            for row in (report.get("top_books") or [])
            if isinstance(row, dict)
        ]
        top_rows: List[str] = []
        for rank, book in enumerate(top_books, start=1):
            time_ms = parse_int(book.get("time_ms")) or 0
            words = parse_int(book.get("words")) or 0
            wpm = parse_float(book.get("wpm"))
            manga = normalize_reading_mode(book.get("reading_mode")) == "manga"
            word_source = str(book.get("word_source") or "")
            words_text = (
                bilingual_text("不适用", "N/A")
                if manga and word_source != "ocr-estimate"
                else bilingual_text(
                    f"{format_words(words)} 词",
                    f"{format_words(words)} words",
                )
            )
            wpm_text = (
                "--"
                if wpm is None
                else f"{wpm:.1f} WPM"
            )
            top_rows.append(
                '<article class="period-book">'
                '<div class="period-book-heading">'
                f"<span>{rank:02d}</span>"
                f"<strong>{html.escape(display_title(str(book.get('title') or '(unknown)')))}</strong>"
                "</div>"
                '<div class="period-book-values">'
                f"<span>{bilingual_text(format_duration(time_ms), format_duration_en(time_ms))}</span>"
                f"<span>{words_text}</span>"
                f"<span>{wpm_text}</span>"
                "</div>"
                "</article>"
            )
        if not top_rows:
            top_rows.append(
                '<div class="empty compact">'
                + bilingual_text(
                    "没有足够的周期书籍增量。",
                    "No sufficient period-level book change.",
                )
                + "</div>"
            )

        missing = [
            row
            for row in (report.get("missing_segments") or [])
            if isinstance(row, dict)
        ]
        missing_html = "".join(
            "<li>"
            + bilingual_text(
                f"{row.get('start')} 至 {row.get('end')} "
                + (
                    "记录不足"
                    if row.get("kind") == "snapshot-gap"
                    else "无覆盖"
                ),
                f"{row.get('start')} to {row.get('end')} "
                + (
                    "sparse snapshots"
                    if row.get("kind") == "snapshot-gap"
                    else "no coverage"
                ),
            )
            + "</li>"
            for row in missing
        )
        if not missing_html:
            missing_html = (
                "<li>"
                + bilingual_text(
                    "未发现明显的缺失时间段。",
                    "No obvious missing interval was detected.",
                )
                + "</li>"
            )

        panels.append(
            f'<section class="period-report-view" id="{panel_id}" '
            f'data-period-panel="{period_key}" '
            f'data-period-reliable="{"true" if report.get("reliable") else "false"}" '
            f'role="tabpanel" aria-labelledby="{tab_id}" '
            f'{"hidden" if not selected else ""}>'
            '<div class="period-report-heading">'
            "<div>"
            f"<h3>{bilingual_text(str(report.get('label_zh') or ''), str(report.get('label_en') or ''))}</h3>"
            f"<p>{bilingual_text(str(report.get('range_zh') or ''), str(report.get('range_en') or ''))}</p>"
            "</div>"
            "</div>"
            f'<div class="period-metrics">{metric_html}</div>'
            '<div class="period-content-grid">'
            '<section class="period-subpanel">'
            f"<h4>{bilingual_text('周期趋势', 'Period trend')}</h4>"
            f"{trend_html}</section>"
            '<section class="period-subpanel">'
            f"<h4>{bilingual_text('Top 书籍', 'Top books')}</h4>"
            f'<div class="period-books">{"".join(top_rows)}</div></section>'
            "</div>"
            f'<p class="period-conclusion {"is-insufficient" if not report.get("reliable") else ""}">'
            f"<strong>{bilingual_text(str(report.get('conclusion_zh') or ''), str(report.get('conclusion_en') or ''))}</strong>"
            "</p>"
            "</section>"
        )

    summary = bilingual_text(
        str(data.get("summary_zh") or ""),
        str(data.get("summary_en") or ""),
    )
    return (
        '<details class="period-panel" id="period-report-panel" open>'
        "<summary>"
        f"<span>{bilingual_text('周期报告', 'Period reports')}</span>"
        f"<em>{summary}</em>"
        "</summary>"
        '<div class="period-body">'
        '<div class="period-tabs" role="tablist" '
        'aria-label="周期报告 / Period reports">'
        + "".join(tabs)
        + "</div>"
        + "".join(panels)
        + "</div></details>"
    )


def build_annual_report_section(
    period_reports: Dict[str, Any],
    trends: Dict[str, Any],
) -> str:
    reports = [
        report
        for report in (period_reports.get("reports") or [])
        if isinstance(report, dict)
    ]
    year_report = next(
        (
            report
            for report in reports
            if str(report.get("key") or "") == "year"
        ),
        None,
    )
    if year_report is None:
        return (
            '<details class="annual-report-panel" id="annual-report-panel" '
            'data-annual-report-state="empty" '
            'data-annual-report-reliable="false">'
            "<summary><span>"
            + bilingual_text("年度阅读年鉴", "Annual reading year")
            + "</span><em>"
            + bilingual_text(
                "暂无可生成的年度报告。",
                "No annual summary is available yet.",
            )
            + "</em></summary>"
            '<div class="annual-report-body">'
            '<div class="empty compact">'
            + bilingual_text(
                "数据不足，暂不生成年度阅读年鉴。",
                "Not enough data to build an annual reading summary.",
            )
            + "</div></div></details>"
        )

    year = str(year_report.get("range_start") or "")[:4]
    if not year:
        year = str(year_report.get("range_zh") or "")[:4] or "--"
    title_zh = f"{year} 阅读年鉴" if year != "--" else "阅读年鉴"
    title_en = f"{year} Reading year" if year != "--" else "Reading year"
    range_zh = str(year_report.get("range_zh") or "")
    range_en = str(year_report.get("range_en") or "")
    confidence, confidence_zh, confidence_en = _period_confidence_label(
        year_report.get("confidence")
    )
    state = "reliable" if year_report.get("reliable") else "insufficient"

    total_time_ms = parse_int(year_report.get("total_time_ms")) or 0
    total_words = parse_int(year_report.get("total_words"))
    reading_days = parse_int(year_report.get("reading_days")) or 0
    active_books = parse_int(year_report.get("active_books")) or 0
    fastest_wpm = parse_float(year_report.get("fastest_wpm"))
    metric_rows = [
        (
            "总时长",
            "Total time",
            format_duration(total_time_ms),
            format_duration_en(total_time_ms),
        ),
        (
            "阅读词数",
            "Words read",
            format_words(total_words),
            format_words(total_words),
        ),
        ("阅读天数", "Reading days", str(reading_days), str(reading_days)),
        ("活跃书籍", "Active books", str(active_books), str(active_books)),
        (
            "最快阅读速度",
            "Fastest reading speed",
            f"{fastest_wpm:.1f} WPM" if fastest_wpm is not None else "--",
            f"{fastest_wpm:.1f} WPM" if fastest_wpm is not None else "--",
        ),
    ]
    metric_html = "".join(
        '<div class="annual-metric">'
        f"<span>{bilingual_text(label_zh, label_en)}</span>"
        f"<strong>{bilingual_text(value_zh, value_en)}</strong>"
        "</div>"
        for label_zh, label_en, value_zh, value_en in metric_rows
    )

    top_books = [
        book
        for book in (year_report.get("top_books") or [])
        if isinstance(book, dict)
    ][:3]
    top_rows: List[str] = []
    for rank, book in enumerate(top_books, start=1):
        time_ms = parse_int(book.get("time_ms")) or 0
        words = parse_int(book.get("words")) or 0
        wpm = parse_float(book.get("wpm"))
        manga = normalize_reading_mode(book.get("reading_mode")) == "manga"
        word_source = str(book.get("word_source") or "")
        words_text = (
            bilingual_text("不适用", "N/A")
            if manga and word_source != "ocr-estimate"
            else bilingual_text(
                f"{format_words(words)} 词",
                f"{format_words(words)} words",
            )
        )
        wpm_text = (
            "--"
            if wpm is None
            else f"{wpm:.1f} WPM"
        )
        top_rows.append(
            '<article class="annual-book">'
            f'<span class="annual-book-rank">{rank:02d}</span>'
            '<div class="annual-book-copy">'
            f"<strong>{html.escape(display_title(str(book.get('title') or '(unknown)')))}</strong>"
            "<small>"
            f"{bilingual_text(format_duration(time_ms), format_duration_en(time_ms))}"
            " · "
            f"{words_text}"
            f" · {wpm_text}"
            "</small></div></article>"
        )
    if not top_rows:
        top_rows.append(
            '<div class="empty compact">'
            + bilingual_text(
                "年度 Top 书籍数据不足。",
                "Not enough annual book-level data.",
            )
            + "</div>"
        )

    return (
        '<details class="annual-report-panel" id="annual-report-panel" '
        f'data-annual-report-state="{state}" '
        f'data-annual-report-reliable="{"true" if year_report.get("reliable") else "false"}" '
        f'data-annual-report-year="{html.escape(year, quote=True)}">'
        "<summary><span>"
        + bilingual_text("年度阅读年鉴", "Annual reading year")
        + "</span><em>"
        + bilingual_text(
            range_zh,
            range_en,
        )
        + "</em></summary>"
        '<div class="annual-report-body">'
        '<section class="annual-report-sheet">'
        '<div class="annual-report-heading"><div>'
        f"<h3>{bilingual_text(title_zh, title_en)}</h3>"
        f"<p>{bilingual_text(range_zh, range_en)}</p>"
        "</div>"
        "</div>"
        f'<div class="annual-metrics">{metric_html}</div>'
        '<div class="annual-content-grid">'
        '<section class="annual-subpanel">'
        f"<h4>{bilingual_text('年度 Top 书籍', 'Top books of the year')}</h4>"
        f'<div class="annual-books">{"".join(top_rows)}</div></section>'
        "</div>"
        f'<p class="period-conclusion {"is-insufficient" if not year_report.get("reliable") else ""}">'
        f"<strong>{bilingual_text(str(year_report.get('conclusion_zh') or ''), str(year_report.get('conclusion_en') or ''))}</strong>"
        "</p></section></div></details>"
    )


def build_heatmap_section(heatmap: Dict[str, Any]) -> str:
    daily = [
        row for row in heatmap.get("daily", []) if isinstance(row, dict)
    ]
    if not daily:
        return (
            '<div class="empty compact">'
            + bilingual_text(
                "暂无阅读热力图数据。",
                "No reading heatmap data.",
            )
            + "</div>"
        )

    max_time = max(
        (parse_int(row.get("total_time_ms")) or 0 for row in daily),
        default=0,
    )
    start_date = date.fromisoformat(str(daily[0]["date"]))
    end_date = date.fromisoformat(str(daily[-1]["date"]))
    calendar_start = start_date - timedelta(days=start_date.weekday())
    calendar_end = end_date + timedelta(days=6 - end_date.weekday())
    daily_by_date = {
        str(row.get("date") or ""): row for row in daily
    }

    cells: List[str] = []
    cursor = calendar_start
    while cursor <= calendar_end:
        if cursor < start_date or cursor > end_date:
            cells.append(
                '<span class="heatmap-day heatmap-day-empty" '
                'aria-hidden="true"></span>'
            )
        else:
            day_key = cursor.isoformat()
            row = daily_by_date.get(day_key, {})
            total_time_ms = parse_int(row.get("total_time_ms")) or 0
            level = 0
            if total_time_ms > 0 and max_time > 0:
                level = max(
                    1,
                    min(4, int(math.ceil(total_time_ms / max_time * 4))),
                )
            tooltip_zh = f"{day_key} · {format_duration(total_time_ms)}"
            tooltip_en = f"{day_key} · {format_duration_en(total_time_ms)}"
            cells.append(
                f'<span class="heatmap-day level-{level}" '
                f'data-date="{day_key}" '
                f'data-minutes="{total_time_ms / 60000.0:.3f}" '
                f'data-title-zh="{html.escape(tooltip_zh, quote=True)}" '
                f'data-title-en="{html.escape(tooltip_en, quote=True)}" '
                f'title="{html.escape(tooltip_zh, quote=True)}" '
                f'aria-label="{html.escape(tooltip_zh, quote=True)}"></span>'
            )
        cursor += timedelta(days=1)

    source_text_zh = "每日阅读时长"
    source_text_en = "Daily reading time"
    source_class = "heatmap-source-active"

    return (
        '<div class="heatmap" id="reading-heatmap" '
        f'data-source="{html.escape(str(heatmap.get("source") or ""), quote=True)}" '
        f'data-status="{html.escape(str(heatmap.get("status") or ""), quote=True)}" '
        f'data-days="{len(daily)}" '
        f'data-source-label-zh="{html.escape(source_text_zh, quote=True)}" '
        f'data-source-label-en="{html.escape(source_text_en, quote=True)}">'
        '<div class="heatmap-scroll">'
        '<div class="heatmap-weekdays" aria-hidden="true">'
        f"<span>{bilingual_text('一', 'M')}</span>"
        f"<span>{bilingual_text('二', 'T')}</span>"
        f"<span>{bilingual_text('三', 'W')}</span>"
        f"<span>{bilingual_text('四', 'T')}</span>"
        f"<span>{bilingual_text('五', 'F')}</span>"
        f"<span>{bilingual_text('六', 'S')}</span>"
        f"<span>{bilingual_text('日', 'S')}</span>"
        "</div>"
        f'<div class="heatmap-grid">{"".join(cells)}</div>'
        "</div>"
        '<div class="heatmap-footer">'
        f'<span class="{source_class}">'
        f"{bilingual_text(source_text_zh, source_text_en)}</span>"
        '<span class="heatmap-legend" aria-label="阅读强度 / Reading intensity">'
        f"<em>{bilingual_text('少', 'Less')}</em>"
        '<i class="level-0"></i><i class="level-1"></i>'
        "<i class=\"level-2\"></i><i class=\"level-3\"></i>"
        f"<i class=\"level-4\"></i><em>{bilingual_text('多', 'More')}</em>"
        "</span>"
        "</div>"
        "</div>"
    )


def build_summary(rows: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
    total_time_ms = sum(parse_int(row.get("total_time_ms")) or 0 for row in rows)
    word_rows = [
        row
        for row in rows
        if not is_manga(row) or ocr_estimate_eligible(row)
    ]
    kindle_word_rows = [row for row in rows if not is_manga(row)]
    effective_word_time_ms = sum(
        parse_int(row.get("total_time_ms")) or 0 for row in word_rows
    )
    kindle_total_words = sum(
        parse_int(row.get("total_words")) or 0 for row in kindle_word_rows
    )
    effective_words_read_total = sum(
        parse_float(effective_words_read(row)) or 0.0 for row in word_rows
    )
    exact_rows = [
        row for row in kindle_word_rows if has_exact_word_count(row)
    ]
    effective_total_words = sum(
        parse_float(effective_book_words(row)) or 0.0 for row in word_rows
    )
    annotations = sum(parse_int(row.get("annotations")) or 0 for row in rows)
    minutes = effective_word_time_ms / 60000.0
    weighted_wpm = (
        round(effective_words_read_total / minutes, 1)
        if minutes > 0
        else None
    )
    progress_values = [
        value
        for value in (word_progress(row) for row in word_rows)
        if value is not None
    ]
    average_progress = (
        round(sum(progress_values) / len(progress_values), 1)
        if progress_values
        else None
    )
    return {
        "books": len(rows),
        "word_books": len(word_rows),
        "kindle_word_books": len(kindle_word_rows),
        "manga_books": sum(1 for row in rows if is_manga(row)),
        "ocr_estimate_books": sum(
            1 for row in rows if ocr_estimate_eligible(row)
        ),
        "reading_record_books": sum(
            1 for row in rows if has_reading_record(row)
        ),
        "completed_books": sum(
            1 for row in rows if is_completed_book(row)
        ),
        "completed_manga_volumes": sum(
            1 for row in rows if is_completed_manga_volume(row)
        ),
        "total_time_ms": total_time_ms,
        "word_time_ms": effective_word_time_ms,
        "effective_word_time_ms": effective_word_time_ms,
        "total_words": kindle_total_words,
        "kindle_total_words": kindle_total_words,
        "effective_words_read": round(effective_words_read_total, 2),
        "effective_total_words": effective_total_words,
        "exact_book_count": len(exact_rows),
        "annotations": annotations,
        "weighted_wpm": weighted_wpm,
        "average_word_progress": average_progress,
    }


def achievement_tiers(
    metric: str,
    baseline: float,
) -> List[Dict[str, Any]]:
    """Build four strict, monotonic tiers above the first-import baseline."""

    key = str(metric or "").strip().casefold()
    steps = ACHIEVEMENT_STEPS.get(key)
    if steps is None:
        raise ValueError("Unknown achievement metric: %s" % metric)
    base = max(0.0, float(baseline or 0.0))
    messages = ACHIEVEMENT_ENCOURAGEMENTS[key]
    tiers: List[Dict[str, Any]] = []
    for index, step in enumerate(steps):
        threshold = base + float(step)
        if tiers and threshold <= tiers[-1]["threshold"] + 1e-9:
            threshold = tiers[-1]["threshold"] + 1.0
        decimals = 2 if key in {"time", "speed"} else 0
        message_zh, message_en = messages[index]
        tiers.append(
            {
                "level": index + 1,
                "threshold": round(threshold, decimals),
                "encouragement_zh": message_zh,
                "encouragement_en": message_en,
            }
        )
    return tiers


def _achievement_speed_value(
    trends: Optional[Dict[str, Any]],
) -> Tuple[float, bool, int, int]:
    daily_rows = [
        row
        for row in ((trends or {}).get("daily") or [])
        if isinstance(row, dict)
    ]
    speed_time_ms = sum(
        max(0, parse_int(row.get("total_time_ms")) or 0)
        for row in daily_rows
    )
    speed_words = sum(
        max(
            0.0,
            parse_float(row.get("effective_words_read"))
            or parse_int(row.get("total_words"))
            or 0,
        )
        for row in daily_rows
    )
    speed_minutes = speed_time_ms / 60000.0
    ready = (
        speed_time_ms >= 30 * 60 * 1000
        and speed_words >= 3000
        and speed_minutes > 0
    )
    value = speed_words / speed_minutes if ready else 0.0
    return round(value, 2), ready, speed_time_ms, int(round(speed_words))


def _achievement_values(
    rows: Sequence[Dict[str, Any]],
    trends: Optional[Dict[str, Any]],
    threshold_percent: float = DEFAULT_COMPLETED_BOOK_THRESHOLD_PERCENT,
) -> Dict[str, float]:
    summary = build_summary(rows)
    streak = (trends or {}).get("streak") or {}
    current_streak = max(0, parse_int(streak.get("current_days")) or 0)
    longest_streak = max(0, parse_int(streak.get("longest_days")) or 0)
    speed_value, _, _, _ = _achievement_speed_value(trends)
    return {
        "streak": float(max(current_streak, longest_streak)),
        "time": float(parse_int(summary.get("total_time_ms")) or 0) / 3600000.0,
        "books": float(
            sum(
                1
                for row in rows
                if is_completed_book(row, threshold_percent)
            )
        ),
        "manga_volumes": float(
            sum(
                1
                for row in rows
                if is_completed_manga_volume(row, threshold_percent)
            )
        ),
        "words": float(parse_int(summary.get("effective_words_read")) or 0),
        "speed": float(speed_value),
    }


def _snapshot_baseline_metrics(
    snapshot: Dict[str, Any],
    threshold_percent: float = DEFAULT_COMPLETED_BOOK_THRESHOLD_PERCENT,
) -> Dict[str, float]:
    totals = snapshot_totals(snapshot)
    return {
        "total_time_ms": float(parse_int(totals.get("total_time_ms")) or 0),
        "total_words": float(
            parse_int(totals.get("effective_words_read")) or 0
        ),
        "book_count": float(parse_int(totals.get("books")) or 0),
        "read_book_count": float(
            parse_int(totals.get("reading_record_books")) or 0
        ),
        "completed_book_count": float(
            sum(
                1
                for book in (snapshot.get("books") or {}).values()
                if isinstance(book, dict)
                and is_completed_book(book, threshold_percent)
            )
        ),
        "completed_manga_volume_count": float(
            sum(
                1
                for book in (snapshot.get("books") or {}).values()
                if isinstance(book, dict)
                and is_completed_manga_volume(book, threshold_percent)
            )
        ),
        "streak_days": 0.0,
        "speed_wpm": 0.0,
    }


def _profile_baseline_from_rows(
    rows: Sequence[Dict[str, Any]],
    threshold_percent: float,
) -> Dict[str, float]:
    summary = build_summary(rows)
    return {
        "total_time_ms": float(parse_int(summary.get("total_time_ms")) or 0),
        "total_words": float(
            parse_int(summary.get("effective_words_read")) or 0
        ),
        "book_count": float(parse_int(summary.get("books")) or 0),
        "read_book_count": float(
            parse_int(summary.get("reading_record_books")) or 0
        ),
        "completed_book_count": float(
            sum(
                1
                for row in rows
                if is_completed_book(row, threshold_percent)
            )
        ),
        "completed_manga_volume_count": float(
            sum(
                1
                for row in rows
                if is_completed_manga_volume(row, threshold_percent)
            )
        ),
        "streak_days": 0.0,
        "speed_wpm": 0.0,
    }


def _baseline_tier_inputs(
    baseline: Dict[str, Any],
) -> Dict[str, float]:
    return {
        "streak": float(parse_float(baseline.get("streak_days")) or 0.0),
        "time": (
            float(parse_int(baseline.get("total_time_ms")) or 0) / 3600000.0
        ),
        "books": float(
            parse_int(baseline.get("completed_book_count")) or 0
        ),
        "manga_volumes": float(
            parse_int(baseline.get("completed_manga_volume_count")) or 0
        ),
        "words": float(parse_int(baseline.get("total_words")) or 0),
        "speed": float(parse_float(baseline.get("speed_wpm")) or 0.0),
    }


def _achievement_metric_label(metric: str) -> Tuple[str, str]:
    return {
        "streak": ("连续阅读", "Reading streak"),
        "time": ("累计时长", "Cumulative time"),
        "books": ("已读书籍", "Completed books"),
        "manga_volumes": ("漫画卷数", "Completed manga volumes"),
        "words": ("累计词数", "Cumulative words"),
        "speed": ("阅读速度", "Reading speed"),
    }.get(metric, (metric, metric))


def _snapshot_achievement_values(
    snapshots: Sequence[Dict[str, Any]],
    threshold_percent: float = DEFAULT_COMPLETED_BOOK_THRESHOLD_PERCENT,
) -> List[Tuple[Dict[str, Any], Dict[str, float]]]:
    ordered: List[Tuple[Dict[str, Any], Optional[datetime]]] = []
    for index, snapshot in enumerate(snapshots):
        if isinstance(snapshot, dict):
            ordered.append((snapshot, _snapshot_datetime(snapshot)))
    ordered.sort(
        key=lambda item: (
            item[1] is None,
            (
                item[1].replace(tzinfo=None)
                if item[1] is not None
                else datetime.min
            ),
        )
    )

    values: List[Tuple[Dict[str, Any], Dict[str, float]]] = []
    prefix: List[Dict[str, Any]] = []
    for snapshot, captured in ordered:
        prefix.append(snapshot)
        totals = snapshot_totals(snapshot)
        snapshot_date = captured.date() if captured is not None else date.today()
        streak = _reading_streaks(
            _daily_reading_totals(prefix),
            snapshot_date,
        )
        completed_books = sum(
            1
            for book in (snapshot.get("books") or {}).values()
            if isinstance(book, dict)
            and is_completed_book(book, threshold_percent)
        )
        completed_manga_volumes = sum(
            1
            for book in (snapshot.get("books") or {}).values()
            if isinstance(book, dict)
            and is_completed_manga_volume(book, threshold_percent)
        )
        values.append(
            (
                snapshot,
                {
                    "streak": float(
                        max(
                            parse_int(streak.get("current_days")) or 0,
                            parse_int(streak.get("longest_days")) or 0,
                        )
                    ),
                    "time": (
                        float(parse_int(totals.get("total_time_ms")) or 0)
                        / 3600000.0
                    ),
                    "books": float(completed_books),
                    "manga_volumes": float(completed_manga_volumes),
                    "words": float(
                        parse_int(totals.get("effective_words_read")) or 0
                    ),
                },
            )
        )
    return values


def _star_history_key(metric: str, level: int) -> Tuple[str, int]:
    return (str(metric or "").strip().casefold(), int(level))


def _star_formatter(metric: str) -> Any:
    return {
        "streak": lambda value: f"{value:.0f} 天",
        "time": lambda value: f"{value:g} 小时",
        "books": lambda value: f"{value:.0f} 本",
        "manga_volumes": lambda value: f"{value:.0f} 卷",
        "words": lambda value: f"{format_words(parse_int(value))} 词",
        "speed": lambda value: f"{value:.1f} WPM",
    }.get(metric, lambda value: f"{value:g}")


def _build_star_history(
    existing_profile: Optional[Dict[str, Any]],
    snapshots: Sequence[Dict[str, Any]],
    tiers: Dict[str, List[Dict[str, Any]]],
    levels: Dict[str, int],
    previous_levels: Dict[str, int],
    current_values: Dict[str, float],
    captured_at: str,
    threshold_percent: float,
) -> List[Dict[str, Any]]:
    """Keep exact new unlocks and conservatively backfill old ones."""

    existing = existing_profile if isinstance(existing_profile, dict) else {}
    existing_history = (
        existing.get("star_history")
        if isinstance(existing.get("star_history"), list)
        else existing.get("unlocks")
        if isinstance(existing.get("unlocks"), list)
        else []
    )
    existing_by_key: Dict[Tuple[str, int], Dict[str, Any]] = {}
    for entry in existing_history:
        if not isinstance(entry, dict):
            continue
        metric = str(entry.get("metric") or "").strip().casefold()
        level = parse_int(entry.get("level"))
        if metric and level is not None:
            existing_by_key[_star_history_key(metric, level)] = entry

    inferred_by_key: Dict[Tuple[str, int], Dict[str, Any]] = {}
    if snapshots:
        for snapshot, snapshot_values in _snapshot_achievement_values(
            snapshots,
            threshold_percent,
        ):
            captured = str(snapshot.get("captured_at") or "")
            for metric in ACHIEVEMENT_TRACK_KEYS:
                metric_values = snapshot_values.get(metric)
                if metric_values is None:
                    continue
                for tier in tiers.get(metric, []):
                    level = parse_int(tier.get("level")) or 0
                    threshold = parse_float(tier.get("threshold"))
                    key = _star_history_key(metric, level)
                    if (
                        key not in inferred_by_key
                        and threshold is not None
                        and metric_values >= threshold - 1e-9
                    ):
                        inferred_by_key[key] = {
                            "unlocked_at": captured,
                            "observed_value": metric_values,
                            "source": "history-inference",
                            "approximate": True,
                        }

    result: List[Dict[str, Any]] = []
    for metric in ACHIEVEMENT_TRACK_KEYS:
        metric_label_zh, metric_label_en = _achievement_metric_label(metric)
        formatter = _star_formatter(metric)
        current_level = max(0, parse_int(levels.get(metric)) or 0)
        previous_level = max(0, parse_int(previous_levels.get(metric)) or 0)
        for tier in tiers.get(metric, []):
            level = parse_int(tier.get("level")) or 0
            if level <= 0 or level > current_level:
                continue
            key = _star_history_key(metric, level)
            existing_entry = existing_by_key.get(key) or {}
            inferred = inferred_by_key.get(key) or {}
            is_new_unlock = level > previous_level
            if existing_entry.get("unlocked_at") and not existing_entry.get(
                "approximate"
            ):
                unlocked_at = str(existing_entry.get("unlocked_at") or "")
                source = str(existing_entry.get("source") or "observed")
                approximate = False
                observed_value = parse_float(
                    existing_entry.get("observed_value")
                )
            elif is_new_unlock:
                unlocked_at = captured_at
                source = "observed"
                approximate = False
                observed_value = current_values.get(metric)
            elif inferred.get("unlocked_at"):
                unlocked_at = str(inferred.get("unlocked_at") or "")
                source = str(inferred.get("source") or "history-inference")
                approximate = True
                observed_value = parse_float(inferred.get("observed_value"))
            else:
                unlocked_at = str(
                    existing_entry.get("unlocked_at")
                    or captured_at
                    or existing.get("created_at")
                    or ""
                )
                source = str(
                    existing_entry.get("source") or "legacy-profile"
                )
                approximate = True
                observed_value = parse_float(
                    existing_entry.get("observed_value")
                )
            threshold = parse_float(tier.get("threshold")) or 0.0
            if observed_value is None:
                observed_value = max(0.0, current_values.get(metric, 0.0))
            result.append(
                {
                    "metric": metric,
                    "level": level,
                    "unlocked_at": unlocked_at,
                    "threshold": threshold,
                    "threshold_text": formatter(threshold),
                    "observed_value": round(float(observed_value), 2),
                    "observed_text": formatter(float(observed_value)),
                    "source": source,
                    "approximate": bool(approximate),
                    "achievement_zh": (
                        f"{metric_label_zh} · 第 {level} 颗星"
                    ),
                    "achievement_en": (
                        f"{metric_label_en} · Star {level}"
                    ),
                    "track_zh": metric_label_zh,
                    "track_en": metric_label_en,
                }
            )
    return result


def _normalize_profile_tiers(
    value: Any,
    metric: str,
    baseline_value: float,
    generation_base: Optional[float] = None,
) -> List[Dict[str, Any]]:
    tiers: List[Dict[str, Any]] = []
    if isinstance(value, list):
        previous = baseline_value
        for index, raw in enumerate(value[:4]):
            if isinstance(raw, dict):
                threshold = parse_float(raw.get("threshold"))
            else:
                threshold = parse_float(raw)
            if threshold is None or threshold <= previous + 1e-9:
                tiers = []
                break
            decimals = 2 if metric in {"time", "speed"} else 0
            fallback_messages = ACHIEVEMENT_ENCOURAGEMENTS[metric][index]
            tiers.append(
                {
                    "level": index + 1,
                    "threshold": round(threshold, decimals),
                    "encouragement_zh": fallback_messages[0],
                    "encouragement_en": fallback_messages[1],
                }
            )
            previous = threshold
    if len(tiers) == 4:
        return tiers
    return achievement_tiers(
        metric,
        baseline_value if generation_base is None else generation_base,
    )


def build_achievement_profile(
    rows: Sequence[Dict[str, Any]],
    history: Optional[Dict[str, Any]] = None,
    captured_at: str = "",
    source_path: str = "",
    existing_profile: Optional[Dict[str, Any]] = None,
    threshold_percent: float = DEFAULT_COMPLETED_BOOK_THRESHOLD_PERCENT,
    trends: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Build or update the first baseline and monotonic unlocked levels."""

    threshold = normalized_completed_threshold(threshold_percent)
    existing = existing_profile if isinstance(existing_profile, dict) else None
    snapshots = [
        snapshot
        for snapshot in ((history or {}).get("snapshots") or [])
        if isinstance(snapshot, dict)
    ]
    first_creation = existing is None or not isinstance(
        existing.get("baseline"), dict
    )

    if not first_creation:
        baseline = dict(existing.get("baseline") or {})
    elif snapshots:
        baseline = _snapshot_baseline_metrics(snapshots[0], threshold)
        baseline["captured_at"] = str(snapshots[0].get("captured_at") or "")
        baseline["source"] = str(snapshots[0].get("source") or "")
    else:
        baseline = _profile_baseline_from_rows(rows, threshold)
        baseline["captured_at"] = captured_at
        baseline["source"] = source_path

    baseline_tier_inputs = _baseline_tier_inputs(baseline)
    current_values = _achievement_values(rows, trends, threshold)
    existing_tiers = (
        (existing or {}).get("tiers")
        if isinstance((existing or {}).get("tiers"), dict)
        else {}
    )
    existing_levels = (
        (existing or {}).get("levels")
        if isinstance((existing or {}).get("levels"), dict)
        else {}
    )
    existing_observed = (
        (existing or {}).get("observed_values")
        if isinstance((existing or {}).get("observed_values"), dict)
        else {}
    )
    known_metrics = (
        set(existing_tiers)
        | set(existing_levels)
        | set(existing_observed)
    )
    new_metrics = {
        metric
        for metric in ACHIEVEMENT_STEPS
        if metric not in known_metrics
    }
    if (
        existing is not None
        and (parse_int(existing.get("version")) or 0)
        < ACHIEVEMENT_PROFILE_VERSION
    ):
        # The manga-volume track was added after older profiles were
        # generated. Rebuild that one track from its baseline so existing
        # completed volumes are counted instead of being swallowed as the
        # profile's new starting point.
        new_metrics.add("manga_volumes")
    tier_inputs = {
        metric: (
            baseline_tier_inputs[metric]
            if metric in new_metrics
            else max(
                baseline_tier_inputs[metric],
                current_values.get(metric, 0.0),
            )
        )
        for metric in ACHIEVEMENT_STEPS
    }
    tiers = {
        metric: _normalize_profile_tiers(
            (
                None
                if metric in new_metrics
                else existing_tiers.get(metric)
            ),
            metric,
            baseline_tier_inputs[metric],
            generation_base=tier_inputs[metric],
        )
        for metric in ACHIEVEMENT_STEPS
    }

    previous_levels = {
        metric: (
            0
            if metric in new_metrics
            else max(0, parse_int(existing_levels.get(metric)) or 0)
        )
        for metric in ACHIEVEMENT_STEPS
    }
    levels = {
        metric: previous_levels[metric]
        for metric in ACHIEVEMENT_STEPS
    }
    observed_values: Dict[str, float] = {}
    if not first_creation and existing_observed:
        for metric in ACHIEVEMENT_STEPS:
            if metric in new_metrics:
                previous = baseline_tier_inputs.get(metric, 0.0)
            else:
                previous = parse_float(existing_observed.get(metric))
                if previous is None:
                    previous = baseline_tier_inputs.get(metric, 0.0)
            current = current_values.get(metric, 0.0)
            crossed = sum(
                1
                for tier in tiers[metric]
                if previous < float(tier["threshold"]) - 1e-9
                and current >= float(tier["threshold"]) - 1e-9
            )
            levels[metric] = min(4, levels[metric] + crossed)
            observed_values[metric] = max(previous, current)
    else:
        observed_values = {
            metric: current_values.get(metric, 0.0)
            for metric in ACHIEVEMENT_STEPS
        }

    created_at = (
        str((existing or {}).get("created_at") or "")
        if existing
        else captured_at
    )
    star_history = _build_star_history(
        existing,
        snapshots,
        tiers,
        levels,
        previous_levels,
        current_values,
        captured_at,
        threshold,
    )
    return {
        "version": ACHIEVEMENT_PROFILE_VERSION,
        "created_at": created_at or captured_at,
        "updated_at": captured_at,
        "completed_threshold_percent": threshold,
        "baseline": baseline,
        "tiers": tiers,
        "levels": levels,
        "observed_values": observed_values,
        "star_history": star_history,
        "levels_initialized": True,
    }


def load_achievement_profile(path: Path) -> Optional[Dict[str, Any]]:
    if not path.is_file():
        return None
    try:
        with path.open("r", encoding="utf-8-sig") as handle:
            decoded = json.load(handle)
    except (OSError, ValueError):
        backup_corrupt_file(path)
        return None
    if not isinstance(decoded, dict):
        backup_corrupt_file(path)
        return None
    return decoded


def save_achievement_profile(path: Path, profile: Dict[str, Any]) -> None:
    _atomic_write_text(
        path,
        json.dumps(
            profile,
            ensure_ascii=False,
            indent=2,
            sort_keys=True,
        )
        + "\n",
    )


def _achievement_tier_formatter(
    metric: str,
) -> Any:
    return {
        "streak": lambda target: (
            f"{target:.0f} 天",
            f"{target:.0f} days",
        ),
        "time": lambda target: (
            f"{target:g} 小时",
            f"{target:g} h",
        ),
        "books": lambda target: (
            f"{target:.0f} 本",
            f"{target:.0f} books",
        ),
        "manga_volumes": lambda target: (
            f"{target:.0f} 卷",
            f"{target:.0f} volumes",
        ),
        "words": lambda target: (
            f"{format_words(parse_int(target))} 词",
            f"{format_words(parse_int(target))} words",
        ),
        "speed": lambda target: (
            f"{target:.1f} WPM",
            f"{target:.1f} WPM",
        ),
    }[metric]


def _achievement_star_html(star: Dict[str, Any]) -> str:
    state = str(star.get("state") or "future")
    state_label_zh = "已点亮" if state == "lit" else "未点亮"
    state_label_en = "Unlocked" if state == "lit" else "Locked"
    unlocked_at = str(star.get("unlocked_at") or "")
    approximate = "true" if star.get("approximate") else "false"
    observed_text = str(star.get("observed_text") or "")
    source = str(star.get("source") or "")
    aria_zh = str(
        star.get("aria_zh")
        or f"星 {parse_int(star.get('level')) or 0}：{state_label_zh}"
    )
    aria_en = str(
        star.get("aria_en")
        or f"Star {parse_int(star.get('level')) or 0}: {state_label_en.lower()}"
    )
    star_svg = (
        '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">'
        '<path d="M12 2.5c.75 0 1.43.43 1.76 1.11l2.03 4.11 '
        '4.54.66c1.55.23 2.17 2.12 1.05 3.21l-3.29 3.2.78 4.52'
        'c.27 1.54-1.34 2.72-2.73 1.99L12 19.17l-4.06 2.13'
        'c-1.38.73-2.99-.45-2.72-1.99l.78-4.52-3.29-3.2'
        'C1.59 10.5 2.21 8.61 3.76 8.38l4.54-.66 2.03-4.11'
        'A1.96 1.96 0 0 1 12 2.5Z" '
        'stroke-linejoin="round" stroke-linecap="round"></path></svg>'
    )
    return (
        f'<span class="achievement-star star-{html.escape(state, quote=True)}" '
        f'data-star-level="{parse_int(star.get("level")) or 0}" '
        f'data-star-state="{html.escape(state, quote=True)}" '
        f'data-star-unlocked-at="{html.escape(unlocked_at, quote=True)}" '
        f'data-star-approximate="{approximate}" '
        f'data-star-observed="{html.escape(observed_text, quote=True)}" '
        f'data-star-source="{html.escape(source, quote=True)}" '
        'role="img" tabindex="0" '
        f'aria-label="{html.escape(aria_zh, quote=True)}" '
        f'data-aria-zh="{html.escape(aria_zh, quote=True)}" '
        f'data-aria-en="{html.escape(aria_en, quote=True)}">'
        f"{star_svg}"
        '<span class="sr-only i18n-zh">'
        f"{html.escape(state_label_zh)}</span>"
        '<span class="sr-only i18n-en">'
        f"{html.escape(state_label_en)}</span>"
        "</span>"
    )


def build_achievements(
    summary: Dict[str, Any],
    trends: Dict[str, Any],
    rows: Sequence[Dict[str, Any]],
    profile: Optional[Dict[str, Any]] = None,
    threshold_percent: float = DEFAULT_COMPLETED_BOOK_THRESHOLD_PERCENT,
    include_manga: bool = True,
) -> Dict[str, Any]:
    """Build star-track achievements from the persisted first baseline."""

    threshold = normalized_completed_threshold(threshold_percent)
    resolved_profile = (
        profile
        if isinstance(profile, dict) and isinstance(profile.get("baseline"), dict)
        else build_achievement_profile(
            rows,
            captured_at="",
            threshold_percent=threshold,
            trends=trends,
        )
    )
    tiers_by_metric = resolved_profile.get("tiers") or {}
    levels_by_metric = resolved_profile.get("levels") or {}
    star_history = [
        entry
        for entry in (resolved_profile.get("star_history") or [])
        if isinstance(entry, dict)
    ]
    star_history_by_key = {
        _star_history_key(
            str(entry.get("metric") or ""),
            parse_int(entry.get("level")) or 0,
        ): entry
        for entry in star_history
        if str(entry.get("metric") or "").strip()
        and (parse_int(entry.get("level")) or 0) > 0
    }
    values = _achievement_values(rows, trends, threshold)

    items: List[Dict[str, Any]] = []

    def add_item(
        key: str,
        label: str,
        label_en: str,
        badge: str,
        value: float,
        value_text: str,
        value_text_en: str,
        condition_zh: str,
        condition_en: str,
        detail_zh: str,
        detail_en: str,
        criteria_zh: str,
        criteria_en: str,
        insufficient: bool,
        mango_excluded: bool = False,
    ) -> None:
        raw_tiers = tiers_by_metric.get(key)
        if not isinstance(raw_tiers, list) or len(raw_tiers) != 4:
            tier_inputs = _baseline_tier_inputs(
                resolved_profile.get("baseline") or {}
            )
            raw_tiers = achievement_tiers(key, tier_inputs[key])
        tiers = [
            {
                "level": parse_int(tier.get("level")) or index + 1,
                "threshold": parse_float(tier.get("threshold")) or 0.0,
                "encouragement_zh": str(
                    tier.get("encouragement_zh") or ""
                ),
                "encouragement_en": str(
                    tier.get("encouragement_en") or ""
                ),
            }
            for index, tier in enumerate(raw_tiers)
            if isinstance(tier, dict)
        ]
        base_level = max(0, parse_int(levels_by_metric.get(key)) or 0)
        level = min(len(tiers), max(0, base_level))
        next_tier = tiers[level] if level < len(tiers) else None
        maxed = level >= len(tiers)
        target_value = (
            tiers[-1]["threshold"]
            if maxed and tiers
            else (next_tier["threshold"] if next_tier else 0.0)
        )
        formatter = _achievement_tier_formatter(key)
        target_zh, target_en = formatter(target_value)
        remaining = max(0.0, target_value - max(0.0, value))
        remaining_zh, remaining_en = formatter(remaining)
        if insufficient:
            detail_zh = "继续阅读即可解锁"
            detail_en = "Keep reading to unlock"
        if maxed:
            progress_text_zh = "已达最高等级"
            progress_text_en = "Highest tier reached"
            encouragement_zh = (
                tiers[-1]["encouragement_zh"] if tiers else ""
            )
            encouragement_en = (
                tiers[-1]["encouragement_en"] if tiers else ""
            )
        elif insufficient:
            progress_text_zh = f"先积累数据，距离下一颗星 {remaining_zh}"
            progress_text_en = (
                f"Build more data; {remaining_en} to the next star"
            )
            encouragement_zh = next_tier["encouragement_zh"] if next_tier else ""
            encouragement_en = next_tier["encouragement_en"] if next_tier else ""
        else:
            progress_text_zh = f"距离下一颗星还差 {remaining_zh}"
            progress_text_en = f"{remaining_en} to the next star"
            encouragement_zh = next_tier["encouragement_zh"] if next_tier else ""
            encouragement_en = next_tier["encouragement_en"] if next_tier else ""
        progress = (
            100.0
            if maxed or target_value <= 0
            else max(0.0, min(100.0, max(0.0, value) / target_value * 100.0))
        )
        stars = []
        for tier in tiers:
            star_level = int(tier["level"])
            if star_level <= level:
                star_state = "lit"
            elif star_level == level + 1 and not insufficient:
                star_state = "next"
            elif star_level == level + 1:
                star_state = "next-pending"
            else:
                star_state = "future"
            stars.append(
                {
                    "level": star_level,
                    "threshold": tier["threshold"],
                    "state": star_state,
                    "unlocked_at": str(
                        (
                            star_history_by_key.get(
                                _star_history_key(key, star_level),
                                {},
                            )
                        ).get("unlocked_at")
                        or ""
                    ),
                    "approximate": bool(
                        (
                            star_history_by_key.get(
                                _star_history_key(key, star_level),
                                {},
                            )
                        ).get("approximate")
                    ),
                    "observed_text": str(
                        (
                            star_history_by_key.get(
                                _star_history_key(key, star_level),
                                {},
                            )
                        ).get("observed_text")
                        or ""
                    ),
                    "source": str(
                        (
                            star_history_by_key.get(
                                _star_history_key(key, star_level),
                                {},
                            )
                        ).get("source")
                        or ""
                    ),
                    "aria_zh": (
                        f"星 {star_level}："
                        + ("已点亮" if star_state == "lit" else "未点亮")
                    ),
                    "aria_en": (
                        f"Star {star_level}: "
                        + ("unlocked" if star_state == "lit" else "locked")
                    ),
                }
            )
        state = (
            "insufficient"
            if insufficient
            else ("unlocked" if maxed else "in_progress")
        )
        items.append(
            {
                "key": key,
                "label": label,
                "label_en": label_en,
                "badge": badge,
                "value": value,
                "value_text": value_text,
                "value_text_en": value_text_en,
                "target": target_value,
                "level": level,
                "max_level": len(tiers),
                "tiers": tiers,
                "stars": stars,
                "progress": 0.0 if insufficient else round(progress, 1),
                "state": state,
                "condition_zh": condition_zh,
                "condition_en": condition_en,
                "detail_zh": detail_zh,
                "detail_en": detail_en,
                "criteria_zh": criteria_zh,
                "criteria_en": criteria_en,
                "encouragement_zh": encouragement_zh,
                "encouragement_en": encouragement_en,
                "progress_text_zh": progress_text_zh,
                "progress_text_en": progress_text_en,
                "word_based": not mango_excluded,
            }
        )

    streak = trends.get("streak") or {}
    current_streak = max(0, parse_int(streak.get("current_days")) or 0)
    longest_streak = max(0, parse_int(streak.get("longest_days")) or 0)
    streak_value = max(current_streak, longest_streak)
    add_item(
        "streak",
        "连续阅读",
        "Reading streak",
        "ST",
        float(streak_value),
        f"{streak_value} 天",
        f"{streak_value} days",
        "在首次基线之上逐级点亮",
        "Tiers above the first baseline",
        f"当前 {current_streak} 天 · 最长 {longest_streak} 天",
        f"Current {current_streak} days · longest {longest_streak} days",
        (
            "按历史快照中的连续活跃自然日计算；漫画阅读同样计入。"
            if include_manga
            else "按历史快照中的连续活跃自然日计算。"
        ),
        (
            "Based on consecutive active calendar days; manga counts too."
            if include_manga
            else "Based on consecutive active calendar days."
        ),
        insufficient=streak_value <= 0,
    )

    total_time_ms = max(0, parse_int(summary.get("total_time_ms")) or 0)
    total_hours = total_time_ms / 3600000.0
    add_item(
        "time",
        "累计时长",
        "Cumulative time",
        "TM",
        total_hours,
        format_duration(total_time_ms),
        format_duration_en(total_time_ms),
        "在首次基线之上逐级点亮",
        "Tiers above the first baseline",
        f"当前累计 {format_duration(total_time_ms)}",
        f"Currently {format_duration_en(total_time_ms)}",
        (
            "使用累计阅读时长；漫画阅读同样计入。"
            if include_manga
            else "使用累计阅读时长。"
        ),
        (
            "Uses cumulative reading time; manga counts too."
            if include_manga
            else "Uses cumulative reading time."
        ),
        insufficient=total_time_ms <= 0,
    )

    completed_books = sum(
        1 for row in rows if is_completed_book(row, threshold)
    )
    reading_record_books = sum(1 for row in rows if has_reading_record(row))
    add_item(
        "books",
        "已读书籍",
        "Completed books",
        "BK",
        float(completed_books),
        f"{completed_books} 本",
        f"{completed_books} books",
        "在首次基线之上逐级点亮",
        "Tiers above the first baseline",
        (
            f"有阅读记录 {reading_record_books} 本 · 已读完 "
            f"{completed_books} 本"
        ),
        (
            f"{reading_record_books} books with reading records · "
            f"{completed_books} completed"
        ),
        (
            f"已读完定义为阅读时长大于 0 且 Kindle 进度至少 "
            f"{threshold:g}%；漫画改计入漫画卷数轨道。"
            if include_manga
            else f"已读完定义为阅读时长大于 0 且 Kindle 进度至少 {threshold:g}%。"
        ),
        (
            f"Completed means reading time > 0 and Kindle progress at "
            f"least {threshold:g}%; manga uses its own volume track."
            if include_manga
            else (
                f"Completed means reading time > 0 and Kindle progress at "
                f"least {threshold:g}%."
            )
        ),
        insufficient=reading_record_books <= 0,
    )

    if include_manga:
        completed_manga_volumes = sum(
            1 for row in rows if is_completed_manga_volume(row, threshold)
        )
        manga_books = sum(1 for row in rows if is_manga(row))
        add_item(
            "manga_volumes",
            "漫画卷数",
            "Completed manga volumes",
            "MV",
            float(completed_manga_volumes),
            f"{completed_manga_volumes} 卷",
            f"{completed_manga_volumes} volumes",
            "在首次基线之上逐级点亮",
            "Tiers above the first baseline",
            (
                f"漫画 {manga_books} 本 · 读完 {completed_manga_volumes} 卷"
            ),
            (
                f"{manga_books} manga titles · "
                f"{completed_manga_volumes} completed volumes"
            ),
            (
                f"漫画进度至少 {threshold:g}% 且有时长记录；"
                "与已读书籍轨道不重叠。"
            ),
            (
                f"Manga progress at least {threshold:g}% with reading time; "
                "kept separate from the completed-books track."
            ),
            insufficient=manga_books <= 0,
        )

    total_words = max(
        0, parse_int(summary.get("effective_words_read")) or 0
    )
    word_books = parse_int(summary.get("word_books")) or 0
    add_item(
        "words",
        "累计词数",
        "Cumulative words",
        "WD",
        float(total_words),
        format_words(total_words) if word_books else "--",
        format_words(total_words) if word_books else "--",
        "在首次基线之上逐级点亮",
        "Tiers above the first baseline",
        (
            f"有效累计 {format_words(total_words)} 词"
            if word_books
            else "当前没有有效词数数据"
        ),
        (
            f"{format_words(total_words)} effective words"
            if word_books
            else "No effective word data yet"
        ),
        "按有效已读词数累计；数据不足时继续阅读即可。",
        "Uses effective words read; more reading builds the count.",
        insufficient=word_books <= 0,
        mango_excluded=True,
    )

    return {
        "items": items,
        "unlocked": sum(
            1 for item in items if item["state"] == "unlocked"
        ),
        "stars_unlocked": sum(parse_int(item.get("level")) or 0 for item in items),
        "stars_total": sum(len(item.get("tiers") or []) for item in items),
        "total": len(items),
        "star_history": star_history,
    }


def build_achievement_cards(data: Dict[str, Any]) -> str:
    state_labels = {
        "unlocked": ("已解锁", "Unlocked"),
        "in_progress": ("进行中", "In progress"),
        "insufficient": ("数据不足", "Not enough data"),
    }
    cards: List[str] = []
    for item in data.get("items", []):
        if not isinstance(item, dict):
            continue
        state = str(item.get("state") or "insufficient")
        state_zh, state_en = state_labels.get(
            state,
            state_labels["insufficient"],
        )
        progress = parse_float(item.get("progress")) or 0.0
        progress = max(0.0, min(100.0, progress))
        value = item.get("value", 0)
        target = item.get("target", 0)
        level = parse_int(item.get("level")) or 0
        max_level = parse_int(item.get("max_level")) or 0
        stars = [
            star
            for star in (item.get("stars") or [])
            if isinstance(star, dict)
        ]
        star_html = "".join(
            _achievement_star_html(star) for star in stars
        )
        tiers_json = html.escape(
            json.dumps(
                [
                    {
                        "level": parse_int(tier.get("level")) or 0,
                        "threshold": parse_float(tier.get("threshold")) or 0.0,
                    }
                    for tier in (item.get("tiers") or [])
                    if isinstance(tier, dict)
                ],
                ensure_ascii=False,
                separators=(",", ":"),
            ),
            quote=True,
        )
        cards.append(
            f'<article class="achievement-card achievement-{html.escape(state, quote=True)}" '
            f'data-achievement-key="{html.escape(str(item.get("key") or ""), quote=True)}" '
            f'data-achievement-state="{html.escape(state, quote=True)}" '
            f'data-achievement-value="{html.escape(str(value), quote=True)}" '
            f'data-achievement-target="{html.escape(str(target), quote=True)}" '
            f'data-achievement-progress="{progress:.1f}" '
            f'data-achievement-base-level="{level}" '
            f'data-achievement-max-level="{max_level}" '
            f"data-achievement-tiers=\"{tiers_json}\">"
            '<div class="achievement-badge" aria-hidden="true">'
            f'<span>{html.escape(str(item.get("badge") or "ACH"))}</span></div>'
            '<div class="achievement-content">'
            '<div class="achievement-heading">'
            f'<h3>{bilingual_text(str(item.get("label") or ""), str(item.get("label_en") or ""))}</h3>'
            f'<span class="achievement-state">'
            f"{bilingual_text(state_zh, state_en)}</span>"
            "</div>"
            f'<strong class="achievement-value">'
            f'{bilingual_text(str(item.get("value_text") or "--"), str(item.get("value_text_en") or "--"))}</strong>'
            f'<p class="achievement-condition">'
            f'{bilingual_text(str(item.get("condition_zh") or ""), str(item.get("condition_en") or ""))}</p>'
            '<div class="achievement-stars" role="list" '
            f'aria-label="{html.escape(str(item.get("progress_text_zh") or ""), quote=True)}">'
            f"{star_html}</div>"
            '<div class="achievement-progress" role="progressbar" '
            'aria-valuemin="0" aria-valuemax="100" '
            f'aria-valuenow="{progress:.1f}" '
            f'aria-label="{html.escape(str(item.get("progress_text_zh") or ""), quote=True)}">'
            f'<i style="width:{progress:.1f}%"></i></div>'
            '<div class="achievement-footer">'
            f'<span>{bilingual_text(str(item.get("progress_text_zh") or ""), str(item.get("progress_text_en") or ""))}</span>'
            f'<small>{bilingual_text(str(item.get("detail_zh") or ""), str(item.get("detail_en") or ""))}</small>'
            "</div>"
            f'<p class="achievement-encouragement">'
            f'{bilingual_text(str(item.get("encouragement_zh") or ""), str(item.get("encouragement_en") or ""))}</p>'
            f'<p class="achievement-criteria">'
            f'{bilingual_text(str(item.get("criteria_zh") or ""), str(item.get("criteria_en") or ""))}</p>'
            "</div></article>"
        )
    return "\n".join(cards)


def _star_time_text(value: Any, approximate: bool = False) -> str:
    raw = str(value or "").strip()
    if not raw:
        return "获得时间未知"
    label = format_last_read(raw) or raw
    return f"{label} · 根据历史快照推定" if approximate else label


def build_star_sky_html(data: Dict[str, Any]) -> str:
    items = [
        item for item in (data.get("items") or []) if isinstance(item, dict)
    ]
    history_by_key = {
        _star_history_key(
            str(entry.get("metric") or ""),
            parse_int(entry.get("level")) or 0,
        ): entry
        for entry in (data.get("star_history") or [])
        if isinstance(entry, dict)
    }
    positions = (
        (10, 22, 28),
        (22, 11, 22),
        (35, 27, 32),
        (48, 14, 24),
        (62, 30, 30),
        (79, 16, 22),
        (88, 38, 28),
        (19, 47, 24),
        (39, 55, 34),
        (59, 48, 22),
        (76, 61, 30),
        (12, 72, 26),
        (31, 78, 22),
        (50, 70, 28),
        (68, 82, 34),
        (87, 72, 24),
    )
    stars: List[str] = []
    unlocked_rows: List[str] = []
    unlocked_count = 0
    position_index = 0
    for item in items:
        metric = str(item.get("key") or "")
        label_zh = str(item.get("label") or metric)
        label_en = str(item.get("label_en") or label_zh)
        for star in item.get("stars") or []:
            if not isinstance(star, dict):
                continue
            level = parse_int(star.get("level")) or 0
            lit = str(star.get("state") or "") == "lit"
            entry = history_by_key.get(_star_history_key(metric, level), {})
            threshold_text = str(
                entry.get("threshold_text")
                or star.get("threshold_text")
                or ""
            )
            observed_text = str(
                entry.get("observed_text")
                or star.get("observed_text")
                or ""
            )
            unlocked_at = str(
                entry.get("unlocked_at") or star.get("unlocked_at") or ""
            )
            approximate = bool(
                entry.get("approximate")
                if "approximate" in entry
                else star.get("approximate")
            )
            x, y, size = positions[position_index % len(positions)]
            position_index += 1
            if lit:
                unlocked_count += 1
            name = f"{label_zh} · 第 {level} 颗星"
            detail_aria = (
                f"{name}，所属轨道 {label_zh}，"
                f"门槛 {threshold_text or '未知'}，"
                f"当时统计 {observed_text or '未知'}，"
                f"获得时间 {_star_time_text(unlocked_at, approximate)}"
            )
            classes = "star-sky-star star-sky-star-lit" if lit else (
                "star-sky-star star-sky-star-locked"
            )
            stars.append(
                f'<button type="button" class="{classes}" '
                f'data-sky-metric="{html.escape(metric, quote=True)}" '
                f'data-sky-level="{level}" '
                f'data-sky-lit="{"true" if lit else "false"}" '
                f'data-sky-name="{html.escape(name, quote=True)}" '
                f'data-sky-track="{html.escape(label_zh, quote=True)}" '
                f'data-sky-track-en="{html.escape(label_en, quote=True)}" '
                f'data-sky-threshold="{html.escape(threshold_text, quote=True)}" '
                f'data-sky-observed="{html.escape(observed_text, quote=True)}" '
                f'data-sky-unlocked-at="{html.escape(unlocked_at, quote=True)}" '
                f'data-sky-approximate="{"true" if approximate else "false"}" '
                f'aria-label="{html.escape(detail_aria, quote=True)}" '
                f'style="--star-x:{x}%;--star-y:{y}%;--star-size:{size}px">'
                '<svg viewBox="0 0 24 24" aria-hidden="true" '
                'focusable="false"><path d="M12 2.5c.75 0 1.43.43 '
                '1.76 1.11l2.03 4.11 4.54.66c1.55.23 2.17 2.12 '
                '1.05 3.21l-3.29 3.2.78 4.52c.27 1.54-1.34 2.72'
                '-2.73 1.99L12 19.17l-4.06 2.13c-1.38.73-2.99-.45'
                '-2.72-1.99l.78-4.52-3.29-3.2C1.59 10.5 2.21 '
                '8.61 3.76 8.38l4.54-.66 2.03-4.11A1.96 1.96 '
                '0 0 1 12 2.5Z" stroke-linejoin="round" '
                'stroke-linecap="round"></path></svg></button>'
            )
            if lit:
                unlocked_rows.append(
                    "<li>"
                    f"<strong>{html.escape(name)}</strong>"
                    f"<span>{html.escape(label_zh)} · "
                    f"{html.escape(_star_time_text(unlocked_at, approximate))} · "
                    f"门槛 {html.escape(threshold_text or '未知')} · "
                    f"当时 {html.escape(observed_text or '未知')}</span>"
                    "</li>"
                )
    if not unlocked_rows:
        unlocked_rows.append(
            "<li><strong>还没有点亮的星星</strong>"
            "<span>继续阅读，星空会慢慢亮起来。</span></li>"
        )
    total = sum(len(item.get("stars") or []) for item in items)
    return (
        '<dialog class="star-sky-dialog" id="star-sky-dialog" '
        'aria-labelledby="star-sky-title">'
        '<div class="star-sky-shell">'
        '<header class="star-sky-header">'
        '<div><h2 id="star-sky-title">这是属于你的阅读星空</h2>'
        f'<p>已点亮 <strong id="star-sky-count">{unlocked_count}/{total}</strong> 颗星</p>'
        "</div>"
        '<button type="button" class="star-sky-close" id="star-sky-close" '
        'aria-label="关闭阅读星空">×</button>'
        "</header>"
        '<div class="star-sky-field" role="list">'
        + "".join(stars)
        + "</div>"
        '<div class="star-sky-detail" id="star-sky-detail" role="status" '
        'aria-live="polite">'
        "<strong>点亮一颗星，回看一段阅读时光。</strong>"
        "<span>阅读的每一步，都会在这里留下一颗星。</span>"
        "</div>"
        '<ul class="star-sky-list">'
        + "".join(unlocked_rows)
        + "</ul>"
        "</div></dialog>"
    )


def _progress_bar(value: Optional[float], color: str) -> str:
    width = max(0.0, min(100.0, value or 0.0))
    label = f"{value:.1f}%" if value is not None else "--"
    return (
        '<div class="mini-progress">'
        f'<span>{label}</span>'
        f'<div><i style="width:{width:.2f}%;background:{color}"></i></div>'
        "</div>"
    )


def _delta_class(delta: Optional[float]) -> str:
    if delta is None or abs(delta) < 0.05:
        return "delta-neutral"
    return "delta-up" if delta > 0 else "delta-down"


def _trend_summary_card(
    label: str,
    current_text: str,
    delta_value: Optional[float],
    delta_text: str,
    label_en: Optional[str] = None,
    current_text_en: Optional[str] = None,
    delta_text_en: Optional[str] = None,
) -> str:
    return (
        '<article class="trend-card">'
        f"<span>{bilingual_text(label, label_en or label)}</span>"
        f"<strong>{bilingual_text(current_text, current_text_en or current_text)}</strong>"
        f'<small class="{_delta_class(delta_value)}">'
        f"{bilingual_text(delta_text, delta_text_en or delta_text)}</small>"
        "</article>"
    )


def _trend_movement_value(
    metric: Dict[str, Any],
    decimals: int = 0,
    suffix: str = "",
    duration: bool = False,
) -> str:
    delta = parse_float(metric.get("delta"))
    if duration:
        text = format_signed_duration(parse_int(delta))
        text_en = format_signed_duration_en(parse_int(delta))
    else:
        text = format_signed_number(delta, decimals, suffix)
        text_en = format_signed_number(delta, decimals, suffix)
    return (
        f'<b class="{_delta_class(delta)}">'
        f"{bilingual_text(text, text_en)}</b>"
    )


def _trend_date_label(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return "--"
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
        return parsed.strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return text.replace("T", " ")[:16]


def _trend_axis_value(value: float, metric: str) -> str:
    if value <= 0:
        return "0"
    if metric == "words":
        if value >= 1_000_000:
            return f"{value / 1_000_000:.1f}M"
        if value >= 1_000:
            return f"{value / 1_000:.0f}k"
        return f"{value:.0f}"
    seconds = value / 1000.0
    if seconds >= 3600:
        return f"{seconds / 3600:.1f}h"
    if seconds >= 60:
        return f"{seconds / 60:.0f}m"
    return f"{seconds:.0f}s"


def _trend_series_points(
    trends: Dict[str, Any],
) -> Dict[str, List[Dict[str, Any]]]:
    timeline_points: List[Dict[str, Any]] = []
    previous_totals: Optional[Dict[str, Any]] = None
    for entry in trends.get("timeline", []):
        if not isinstance(entry, dict):
            continue
        totals = entry.get("totals")
        if not isinstance(totals, dict):
            totals = {}
        time_ms = parse_int(totals.get("total_time_ms")) or 0
        words = (
            parse_int(totals.get("effective_words_read"))
            or parse_int(totals.get("total_words"))
            or 0
        )
        captured_at = str(entry.get("captured_at") or "")
        delta_time_ms = (
            time_ms - (parse_int(previous_totals.get("total_time_ms")) or 0)
            if previous_totals is not None
            else None
        )
        delta_words = (
            words - (parse_int(previous_totals.get("total_words")) or 0)
            if previous_totals is not None
            else None
        )
        timeline_points.append(
            {
                "captured_at": captured_at,
                "label": _trend_date_label(captured_at),
                "time_ms": time_ms,
                "words": words,
                "delta_time_ms": delta_time_ms,
                "delta_words": delta_words,
            }
        )
        previous_totals = totals

    daily_points: List[Dict[str, Any]] = []
    for row in trends.get("daily", []):
        if not isinstance(row, dict):
            continue
        time_ms = parse_int(row.get("total_time_ms")) or 0
        words = (
            parse_int(row.get("effective_words_read"))
            or parse_int(row.get("total_words"))
            or 0
        )
        if not row.get("active") and not time_ms and not words:
            continue
        day = str(row.get("date") or "")
        daily_points.append(
            {
                "captured_at": day,
                "label": day,
                "time_ms": time_ms,
                "words": words,
                "delta_time_ms": time_ms,
                "delta_words": words,
            }
        )

    timeline_points = timeline_points[-30:]
    daily_points = daily_points[-30:]
    return {
        "time": [
            {**point, "value": point["time_ms"]}
            for point in timeline_points
        ],
        "words": [
            {**point, "value": point["words"]}
            for point in timeline_points
        ],
        "daily": [
            {**point, "value": point["time_ms"]}
            for point in daily_points
        ],
    }


def _trend_point_tooltip(
    metric: str,
    point: Dict[str, Any],
) -> Tuple[str, str]:
    label = str(point.get("label") or "--")
    time_ms = parse_int(point.get("time_ms")) or 0
    words = parse_int(point.get("words")) or 0
    if metric == "words":
        return (
            f"{label} · 累计阅读词数 {format_words(words)}",
            f"{label} · Cumulative words read {format_words(words)}",
        )
    if metric == "daily":
        return (
            f"{label} · 每日阅读 {format_duration(time_ms)} · "
            f"{format_words(words)} 词",
            f"{label} · Daily reading {format_duration_en(time_ms)} · "
            f"{format_words(words)} words",
        )
    return (
        f"{label} · 累计阅读 {format_duration(time_ms)}",
        f"{label} · Cumulative reading {format_duration_en(time_ms)}",
    )


def _render_trend_chart_view(
    metric: str,
    points: Sequence[Dict[str, Any]],
    active: bool = False,
) -> str:
    labels = {
        "time": ("累计阅读时长趋势", "Cumulative reading-time trend"),
        "words": ("累计阅读词数趋势", "Cumulative word-count trend"),
        "daily": ("每日阅读变化趋势", "Daily reading-change trend"),
    }
    label_zh, label_en = labels[metric]
    hidden = "" if active else " hidden"
    point_count = len(points)
    state = "ready" if point_count >= 2 else (
        "single" if point_count == 1 else "empty"
    )
    if not points:
        return (
            f'<div class="trend-chart-view" data-trend-series="{metric}" '
            f'data-point-count="0" data-chart-state="empty"{hidden}>'
            '<div class="empty compact">'
            + bilingual_text(
                "暂无可用趋势数据。",
                "No trend data is available.",
            )
            + "</div></div>"
        )

    width = 760.0
    height = 250.0
    left = 58.0
    right = 20.0
    top = 22.0
    bottom = 42.0
    plot_width = width - left - right
    plot_height = height - top - bottom
    max_value = max(
        (parse_float(point.get("value")) or 0.0 for point in points),
        default=0.0,
    )
    scale_max = max_value if max_value > 0 else 1.0

    def point_x(index: int) -> float:
        if point_count == 1:
            return left + plot_width / 2.0
        return left + plot_width * index / float(point_count - 1)

    def point_y(value: float) -> float:
        return top + plot_height * (
            1.0 - max(0.0, value) / scale_max
        )

    grid_parts: List[str] = []
    for fraction in (1.0, 0.5, 0.0):
        y = top + plot_height * (1.0 - fraction)
        axis_value = scale_max * fraction
        grid_parts.append(
            '<line class="trend-grid-line" '
            f'x1="{left:.2f}" y1="{y:.2f}" '
            f'x2="{width - right:.2f}" y2="{y:.2f}"></line>'
            f'<text class="trend-axis-label" x="{left - 8:.2f}" '
            f'y="{y + 3.5:.2f}" text-anchor="end">'
            f"{html.escape(_trend_axis_value(axis_value, metric))}</text>"
        )

    coordinates = [
        (
            point_x(index),
            point_y(parse_float(point.get("value")) or 0.0),
        )
        for index, point in enumerate(points)
    ]
    line_markup = ""
    if point_count >= 2:
        path = " ".join(f"{x:.2f},{y:.2f}" for x, y in coordinates)
        line_markup = (
            f'<polyline class="trend-line" points="{path}"></polyline>'
        )

    label_indices = (
        list(range(point_count))
        if point_count <= 5
        else sorted({0, point_count // 2, point_count - 1})
    )
    label_parts: List[str] = []
    for index in label_indices:
        x, _ = coordinates[index]
        label = str(points[index].get("label") or "")
        short_label = label[5:] if len(label) >= 10 else label
        label_parts.append(
            f'<text class="trend-axis-label trend-x-label" x="{x:.2f}" '
            f'y="{height - 16:.2f}" text-anchor="middle">'
            f"{html.escape(short_label)}</text>"
        )

    point_parts: List[str] = []
    for index, point in enumerate(points):
        x, y = coordinates[index]
        tooltip_zh, tooltip_en = _trend_point_tooltip(metric, point)
        point_parts.append(
            '<g class="trend-point" data-trend-point tabindex="0" '
            'role="img" '
            f'data-tooltip-zh="{html.escape(tooltip_zh, quote=True)}" '
            f'data-tooltip-en="{html.escape(tooltip_en, quote=True)}" '
            f'aria-label="{html.escape(tooltip_zh, quote=True)}">'
            f'<circle cx="{x:.2f}" cy="{y:.2f}" r="4.2"></circle>'
            f"<title>{html.escape(tooltip_zh)}</title>"
            "</g>"
        )

    note = ""
    if point_count == 1:
        note = (
            '<p class="trend-chart-note">'
            + bilingual_text(
                "只有一个可用数据点，已显示基线点并暂停曲线绘制。",
                "Only one usable data point is available; the baseline "
                "point is shown without a curve.",
            )
            + "</p>"
        )

    return (
        f'<div class="trend-chart-view" data-trend-series="{metric}" '
        f'data-point-count="{point_count}" data-chart-state="{state}"{hidden}>'
        f'<svg class="trend-chart-svg" viewBox="0 0 {width:.0f} '
        f'{height:.0f}" role="img" aria-label="{html.escape(label_zh)}">'
        + "".join(grid_parts)
        + line_markup
        + "".join(point_parts)
        + "".join(label_parts)
        + "</svg>"
        + note
        + "</div>"
    )


def _build_trend_chart(
    trends: Dict[str, Any],
) -> Tuple[str, Dict[str, int]]:
    series = _trend_series_points(trends)
    point_counts = {
        metric: len(points) for metric, points in series.items()
    }
    views = [
        _render_trend_chart_view("time", series["time"], active=True),
        _render_trend_chart_view("words", series["words"]),
        _render_trend_chart_view("daily", series["daily"]),
    ]
    chart = (
        '<div class="trend-chart-heading">'
        '<div class="trend-metric-switch" role="group" '
        'aria-label="趋势指标 / Trend metric">'
        '<button type="button" data-trend-metric-button="time" '
        'aria-pressed="true"><span class="i18n-zh">总时长</span>'
        '<span class="i18n-en">Reading time</span></button>'
        '<button type="button" data-trend-metric-button="words" '
        'aria-pressed="false"><span class="i18n-zh">总词数</span>'
        '<span class="i18n-en">Words</span></button>'
        '<button type="button" data-trend-metric-button="daily" '
        'aria-pressed="false"><span class="i18n-zh">每日变化</span>'
        '<span class="i18n-en">Daily change</span></button>'
        "</div>"
        '<p><span class="i18n-zh">离线 SVG 折线图 · 悬停或聚焦数据点查看详情</span>'
        '<span class="i18n-en">Offline SVG line chart · hover or focus a point for details</span></p>'
        "</div>"
        '<div class="trend-chart" id="trend-chart">'
        + "".join(views)
        + '<div class="trend-tooltip" id="trend-tooltip" role="status" hidden></div>'
        + "</div>"
    )
    return chart, point_counts


def _build_trend_snapshot_table(trends: Dict[str, Any]) -> str:
    entries = [
        entry
        for entry in trends.get("timeline", [])
        if isinstance(entry, dict)
    ]
    if not entries:
        return (
            '<div class="empty compact">'
            + bilingual_text("暂无历史快照。", "No history snapshots.")
            + "</div>"
        )
    rows: List[str] = []
    for entry in entries:
        totals = entry.get("totals")
        if not isinstance(totals, dict):
            totals = {}
        captured_at = format_last_read(
            str(entry.get("captured_at") or "")
        )
        source = str(entry.get("source") or "--")
        time_ms = parse_int(totals.get("total_time_ms")) or 0
        words = (
            parse_int(totals.get("effective_words_read"))
            or parse_int(totals.get("total_words"))
            or 0
        )
        wpm = parse_float(totals.get("weighted_wpm"))
        wpm_text = f"{wpm:.1f}" if wpm is not None else "--"
        delta_time = parse_int(entry.get("delta_time_ms"))
        delta_words = parse_int(entry.get("delta_words"))
        rows.append(
            "<tr>"
            f"<td>{html.escape(captured_at)}</td>"
            f'<td title="{html.escape(source, quote=True)}">'
            f"<small>{html.escape(source)}</small></td>"
            f"<td>{html.escape(format_duration(time_ms))}</td>"
            f"<td>{html.escape(format_words(words))}</td>"
            f"<td>{wpm_text}</td>"
            f"<td>{html.escape(format_signed_duration(delta_time))}</td>"
            f"<td>{html.escape(format_signed_number(delta_words))}</td>"
            "</tr>"
        )
    return (
        '<div class="table-wrap trend-table-wrap">'
        '<table class="trend-snapshot-table" id="trend-snapshot-table">'
        "<thead><tr>"
        "<th>时间 / Time</th><th>来源 / Source</th>"
        "<th>累计时长 / Total time</th><th>累计有效词数 / Effective words</th>"
        "<th>WPM</th><th>时长变化 / Time delta</th>"
        "<th>词数变化 / Words delta</th>"
        "</tr></thead>"
        f"<tbody>{''.join(rows)}</tbody></table></div>"
    )


def _build_trend_source_note(
    trends: Dict[str, Any],
    heatmap: Optional[Dict[str, Any]],
) -> str:
    timeline_count = len(
        [
            entry
            for entry in trends.get("timeline", [])
            if isinstance(entry, dict)
        ]
    )
    snapshot_count = parse_int(trends.get("snapshot_count")) or timeline_count
    heatmap_data = heatmap if isinstance(heatmap, dict) else {}
    source = str(heatmap_data.get("source") or "history")
    session_count = parse_int(heatmap_data.get("session_count")) or 0
    if source == "fmcache":
        fmcache_zh = (
            f"fmcache 会话 {session_count} 条，趋势和热力图可优先使用会话口径。"
        )
        fmcache_en = (
            f"{session_count} fmcache sessions are available; session-level "
            "coverage is preferred for the trend and heatmap."
        )
    else:
        detail = str(heatmap_data.get("message") or "").strip()
        if not detail:
            detail = "无可用会话"
        detail_en = {
            "reading_sessions 为空": "reading_sessions is empty",
            "没有有效阅读会话": "no valid reading sessions",
            "fmcache.db 未找到": "fmcache.db was not found",
            "fmcache.db 不可读取": "fmcache.db is unreadable",
        }.get(detail, detail)
        fmcache_zh = (
            f"fmcache 会话 {session_count} 条，当前状态为历史快照回退"
            f"（{detail}）。快照变化不等于逐次会话时长。"
        )
        fmcache_en = (
            f"{session_count} fmcache sessions are available; currently "
            f"falling back to history snapshots ({detail_en}). Snapshot "
            "changes are not session-level durations."
        )
    snapshot_zh = (
        f"快照 {snapshot_count} 条，折线使用去重后的 {timeline_count} 个可用点。"
    )
    snapshot_en = (
        f"{snapshot_count} snapshots; the line uses {timeline_count} "
        "deduplicated usable points."
    )
    ocr_transition = trends.get("ocr_transition")
    ocr_note_zh = ""
    ocr_note_en = ""
    if isinstance(ocr_transition, dict):
        transition_at = format_last_read(
            str(ocr_transition.get("captured_at") or "")
        )
        if str(ocr_transition.get("state") or "") == "enabled":
            ocr_note_zh = (
                f" OCR 启用点：{transition_at}；自此累计有效词数会包含"
                "符合条件的 OCR 估算，并以“OCR 估算”文字标注。"
            )
            ocr_note_en = (
                f" OCR enabled at {transition_at}; effective word totals "
                "from then on include eligible OCR estimates labeled "
                "\"OCR estimate\"."
            )
        else:
            ocr_note_zh = (
                f" OCR 关闭点：{transition_at}；历史快照保留，不回写或"
                "删除既有 OCR 估算。"
            )
            ocr_note_en = (
                f" OCR disabled at {transition_at}; historical snapshots "
                "are retained and existing OCR estimates are not rewritten."
            )
    return (
        '<p class="trend-source-note" '
        f'data-trend-source-note="{html.escape(source, quote=True)}" '
        f'data-ocr-transition="{html.escape(str((ocr_transition or {}).get("state") or ""), quote=True)}">'
        f"<strong>{bilingual_text(snapshot_zh, snapshot_en)}</strong>"
        f"<span>{bilingual_text(fmcache_zh, fmcache_en)}</span>"
        f"<span>{bilingual_text(ocr_note_zh, ocr_note_en)}</span>"
        "</p>"
    )


def build_trend_sections(
    trends: Dict[str, Any],
    heatmap: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    totals = trends.get("totals", {})
    has_previous = bool(trends.get("has_previous"))

    time_metric = totals.get("total_time_ms", {})
    words_metric = totals.get("total_words", {})
    wpm_metric = totals.get("weighted_wpm", {})
    progress_metric = totals.get("average_progress", {})

    def delta_text_pair(
        value: Optional[float],
        zh_formatter: Any,
        en_formatter: Any,
    ) -> Tuple[str, str]:
        if not has_previous:
            return "本次建立基线", "Baseline created"
        if value is None:
            return "无可比数据", "No comparison data"
        if abs(value) < 0.05:
            return "与上次持平", "Unchanged from previous"
        return (
            f"较上次 {zh_formatter(value)}",
            f"vs previous {en_formatter(value)}",
        )

    time_delta = parse_float(time_metric.get("delta"))
    words_delta = parse_float(words_metric.get("delta"))
    wpm_delta = parse_float(wpm_metric.get("delta"))
    progress_delta = parse_float(progress_metric.get("delta"))
    daily_rows = [
        row for row in trends.get("daily", []) if isinstance(row, dict)
    ]
    latest_daily = next(
        (
            row
            for row in reversed(daily_rows)
            if row.get("active")
            or (parse_int(row.get("total_time_ms")) or 0) > 0
            or (parse_int(row.get("total_words")) or 0) > 0
        ),
        daily_rows[-1] if daily_rows else {},
    )
    latest_daily_time = parse_int(latest_daily.get("total_time_ms")) or 0
    latest_daily_words = parse_int(latest_daily.get("total_words")) or 0
    latest_daily_day = str(latest_daily.get("date") or "--")

    trend_cards = [
        _trend_summary_card(
            "阅读时长",
            format_duration(parse_int(time_metric.get("current"))),
            time_delta,
            delta_text_pair(
                time_delta,
                format_signed_duration,
                format_signed_duration_en,
            )[0],
            "Reading time",
            format_duration_en(parse_int(time_metric.get("current"))),
            delta_text_pair(
                time_delta,
                format_signed_duration,
                format_signed_duration_en,
            )[1],
        ),
        _trend_summary_card(
            "累计有效词数",
            format_words(parse_int(words_metric.get("current"))),
            words_delta,
            delta_text_pair(
                words_delta,
                lambda value: format_signed_number(value),
                lambda value: format_signed_number(value),
            )[0],
            "Cumulative effective words",
            format_words(parse_int(words_metric.get("current"))),
            delta_text_pair(
                words_delta,
                lambda value: format_signed_number(value),
                lambda value: format_signed_number(value),
            )[1],
        ),
        _trend_summary_card(
            "最近每日变化",
            format_duration(latest_daily_time),
            None,
            (
                f"{latest_daily_day} · {format_words(latest_daily_words)} 词"
                if latest_daily_time or latest_daily_words
                else "暂无每日数据"
            ),
            "Latest daily change",
            format_duration_en(latest_daily_time),
            (
                f"{latest_daily_day} · {format_words(latest_daily_words)} words"
                if latest_daily_time or latest_daily_words
                else "No daily data"
            ),
        ),
        _trend_summary_card(
            "加权平均 WPM",
            (
                f"{parse_float(wpm_metric.get('current')):.1f}"
                if parse_float(wpm_metric.get("current")) is not None
                else "--"
            ),
            wpm_delta,
            delta_text_pair(
                wpm_delta,
                lambda value: format_signed_number(value, 1),
                lambda value: format_signed_number(value, 1),
            )[0],
            "Weighted average WPM",
            (
                f"{parse_float(wpm_metric.get('current')):.1f}"
                if parse_float(wpm_metric.get("current")) is not None
                else "--"
            ),
            delta_text_pair(
                wpm_delta,
                lambda value: format_signed_number(value, 1),
                lambda value: format_signed_number(value, 1),
            )[1],
        ),
        _trend_summary_card(
            "平均进度",
            (
                f"{parse_float(progress_metric.get('current')):.1f}%"
                if parse_float(progress_metric.get("current")) is not None
                else "--"
            ),
            progress_delta,
            delta_text_pair(
                progress_delta,
                lambda value: format_signed_number(value, 1, " 个百分点"),
                lambda value: format_signed_number(value, 1, " pp"),
            )[0],
            "Average progress",
            (
                f"{parse_float(progress_metric.get('current')):.1f}%"
                if parse_float(progress_metric.get("current")) is not None
                else "--"
            ),
            delta_text_pair(
                progress_delta,
                lambda value: format_signed_number(value, 1, " 个百分点"),
                lambda value: format_signed_number(value, 1, " pp"),
            )[1],
        ),
    ]

    windows = trends.get("windows", {})
    window_7d = windows.get("7d", {}) if isinstance(windows, dict) else {}
    window_30d = windows.get("30d", {}) if isinstance(windows, dict) else {}
    streak = trends.get("streak", {})
    current_streak = parse_int(streak.get("current_days")) or 0
    longest_streak = parse_int(streak.get("longest_days")) or 0
    trend_cards.extend(
        [
            _trend_summary_card(
                "近 7 天阅读",
                format_duration(
                    parse_int(window_7d.get("total_time_ms"))
                ),
                None,
                f"{parse_int(window_7d.get('active_days')) or 0} 个活跃日",
                "Last 7 days",
                format_duration_en(
                    parse_int(window_7d.get("total_time_ms"))
                ),
                f"{parse_int(window_7d.get('active_days')) or 0} active days",
            ),
            _trend_summary_card(
                "近 30 天阅读",
                format_duration(
                    parse_int(window_30d.get("total_time_ms"))
                ),
                None,
                f"{parse_int(window_30d.get('active_days')) or 0} 个活跃日",
                "Last 30 days",
                format_duration_en(
                    parse_int(window_30d.get("total_time_ms"))
                ),
                f"{parse_int(window_30d.get('active_days')) or 0} active days",
            ),
            _trend_summary_card(
                "连续阅读",
                f"{current_streak} 天",
                None,
                f"最长 {longest_streak} 天",
                "Reading streak",
                f"{current_streak} days",
                f"Longest {longest_streak} days",
            ),
        ]
    )

    movement_items: List[str] = []
    if has_previous:
        moved_books = []
        for change in trends.get("books", []):
            if change.get("is_new"):
                moved_books.append(change)
                continue
            if any(
                abs(parse_float(change.get(metric, {}).get("delta")) or 0.0)
                >= (0.05 if metric == "progress_percent" else 1.0)
                for metric in (
                    "total_time_ms",
                    "total_words",
                    "wpm",
                    "progress_percent",
                )
            ):
                moved_books.append(change)

        for change in moved_books[:8]:
            title = str(change.get("title") or "(unknown)")
            author = str(change.get("author") or "")
            tag = (
                f'<span class="new-tag">'
                f'{bilingual_text("新增", "New")}</span>'
                if change.get("is_new")
                else ""
            )
            movement_items.append(
                '<article class="movement-item">'
                '<div class="movement-title">'
                f"<strong>{html.escape(title)}</strong>"
                f"<small>{html.escape(author or '--')}</small>{tag}"
                "</div>"
                '<div class="movement-metrics">'
                f"<span><em>{bilingual_text('时长', 'Time')}</em>"
                f"{_trend_movement_value(change['total_time_ms'], duration=True)}"
                "</span>"
                f"<span><em>{bilingual_text('词数', 'Words')}</em>"
                f"{_trend_movement_value(change['total_words'])}</span>"
                "<span><em>WPM</em>"
                f"{_trend_movement_value(change['wpm'], 1)}</span>"
                f"<span><em>{bilingual_text('进度', 'Progress')}</em>"
                f"{_trend_movement_value(change['progress_percent'], 1, 'pp')}"
                "</span>"
                "</div>"
                "</article>"
            )

    if not has_previous:
        movement_html = (
            '<div class="empty compact">'
            + bilingual_text(
                "首次运行已建立趋势基线，下次运行即可显示逐书增量。",
                "The first run established a trend baseline. "
                "Per-book changes will appear on the next run.",
            )
            + "</div>"
        )
    elif movement_items:
        movement_html = "".join(movement_items)
    else:
        movement_html = (
            '<div class="empty compact">'
            + bilingual_text(
                "本次快照与上次相比没有变化。",
                "This snapshot is unchanged from the previous one.",
            )
            + "</div>"
        )

    max_daily_time = max(
        (parse_int(row.get("total_time_ms")) or 0 for row in daily_rows),
        default=0,
    )
    daily_items: List[str] = []
    for row in daily_rows:
        day = str(row.get("date") or "")
        daily_time = parse_int(row.get("total_time_ms")) or 0
        height = (
            min(100.0, daily_time / max_daily_time * 100.0)
            if max_daily_time
            else 0.0
        )
        if daily_time and height < 4.0:
            height = 4.0
        daily_items.append(
            '<div class="daily-column" '
            f'data-date="{html.escape(day, quote=True)}" '
            f'data-minutes="{daily_time / 60000.0:.3f}" '
            f'data-title-zh="{html.escape(day + " · " + format_duration(daily_time), quote=True)}" '
            f'data-title-en="{html.escape(day + " · " + format_duration_en(daily_time), quote=True)}" '
            f'title="{html.escape(day + " · " + format_duration(daily_time), quote=True)}">'
            '<span class="daily-track">'
            f'<i style="height:{height:.2f}%"></i>'
            "</span>"
            f"<time>{html.escape(day[5:] if len(day) >= 10 else day)}</time>"
            "</div>"
        )
    daily_html = (
        '<div class="daily-chart" id="daily-chart" '
        f'data-days="{len(daily_rows)}" '
        f'data-active-days="{parse_int(window_30d.get("active_days")) or 0}">'
        '<div class="daily-bars">'
        + "".join(daily_items)
        + "</div></div>"
        if daily_items
        else (
            '<div class="empty compact">'
            + bilingual_text(
                "暂无每日阅读数据。",
                "No daily reading data.",
            )
            + "</div>"
        )
    )

    chart_html, point_counts = _build_trend_chart(trends)
    snapshot_table = _build_trend_snapshot_table(trends)
    source_note = _build_trend_source_note(trends, heatmap)
    source = str(
        (heatmap or {}).get("source")
        if isinstance(heatmap, dict)
        else "history"
    ) or "history"
    session_count = (
        parse_int(heatmap.get("session_count")) or 0
        if isinstance(heatmap, dict)
        else 0
    )
    panel_attrs = (
        f'data-trend-point-count="{point_counts.get("time", 0)}" '
        f'data-trend-words-point-count="{point_counts.get("words", 0)}" '
        f'data-trend-daily-point-count="{point_counts.get("daily", 0)}"'
    )

    return {
        "cards": "".join(trend_cards),
        "movements": movement_html,
        "daily": daily_html,
        "chart": chart_html,
        "snapshot_table": snapshot_table,
        "source_note": source_note,
        "panel_attrs": panel_attrs,
        "point_counts": point_counts,
    }


def _table_delta_cell(change: Optional[Dict[str, Any]]) -> str:
    if not change:
        return (
            '<td data-column="delta" data-value="-1">'
            '<span class="delta-neutral">--</span></td>'
        )
    if change.get("is_new"):
        return (
            '<td data-column="delta" data-value="-1">'
            '<span class="new-tag">'
            + bilingual_text("新增书籍", "New book")
            + "</span>"
            '<div class="delta-grid"><span>'
            + bilingual_text(
                "本轮首次记录，暂无环比",
                "First recorded this run; no comparison yet",
            )
            + "</span></div>"
            "</td>"
        )

    labels = (
        ("时长", "Time", "total_time_ms", True, 0, ""),
        ("词数", "Words", "total_words", False, 0, ""),
        ("WPM", "WPM", "wpm", False, 1, ""),
        ("进度", "Progress", "progress_percent", False, 1, "pp"),
    )
    parts: List[str] = []
    total_time_delta: Optional[float] = None
    for label, label_en, key, duration, decimals, suffix in labels:
        metric = change.get(key) or {}
        delta = parse_float(metric.get("delta"))
        if key == "total_time_ms":
            total_time_delta = delta
        text = (
            format_signed_duration(parse_int(delta))
            if duration
            else format_signed_number(delta, decimals, suffix)
        )
        parts.append(
            "<span>"
            f"<em>{bilingual_text(label, label_en)}</em>"
            f'<strong class="{_delta_class(delta)}">{html.escape(text)}</strong>'
            "</span>"
        )
    return (
        f'<td data-column="delta" data-value="{total_time_delta if total_time_delta is not None else -1}">'
        f'<div class="delta-grid">{"".join(parts)}</div></td>'
    )


def build_html(
    rows: Sequence[Dict[str, Any]],
    csv_path: Path,
    source_path: str,
    trends: Optional[Dict[str, Any]] = None,
    heatmap: Optional[Dict[str, Any]] = None,
    period_reports: Optional[Dict[str, Any]] = None,
    generated_at: Optional[str] = None,
    history_path: Optional[Path] = None,
    cover_urls: Optional[Mapping[str, str]] = None,
    covers_dir: Optional[Path] = None,
    xlsx_path: Optional[Path] = None,
    achievement_profile: Optional[Dict[str, Any]] = None,
    completed_threshold_percent: float = (
        DEFAULT_COMPLETED_BOOK_THRESHOLD_PERCENT
    ),
    book_modes: Optional[Dict[str, Any]] = None,
) -> str:
    completed_threshold = normalized_completed_threshold(
        completed_threshold_percent
    )
    resolved_book_modes = normalize_book_modes(book_modes or {})
    rows = [dict(row) for row in rows]
    for row in rows:
        if not str(row.get("reading_mode") or "").strip():
            row["reading_mode"] = reading_mode_for_book(
                row,
                resolved_book_modes,
            )
    text_only_ui = runtime_variant() == "text"
    if text_only_ui:
        rows = [row for row in rows if not is_manga(row)]
    summary = build_summary(rows)
    resolved_cover_urls = dict(cover_urls or {})
    trend_data = trends or {
        "has_previous": False,
        "snapshot_count": 1,
        "previous_captured_at": "",
        "current_captured_at": generated_at or "",
        "totals": {},
        "books": [],
        "timeline": [],
        "daily": [],
        "windows": {},
        "streak": {},
    }
    trend_by_book = {
        str(change.get("book_id")): change
        for change in trend_data.get("books", [])
        if isinstance(change, dict)
    }
    if heatmap is None:
        end_date = _snapshot_date(
            {"captured_at": str(trend_data.get("current_captured_at") or "")}
        ) or date.today()
        heatmap_data = {
            "days": FMCACHE_HEATMAP_DAYS,
            "start": "",
            "end": end_date.isoformat(),
            "source": "history",
            "source_label": "历史快照回退",
            "status": "missing",
            "message": "未提供 fmcache.db",
            "db_path": "",
            "session_count": 0,
            "daily": _heatmap_rows(
                trend_data.get("daily") or [],
                end_date,
                FMCACHE_HEATMAP_DAYS,
            ),
        }
    else:
        heatmap_data = heatmap
    heatmap_html = build_heatmap_section(heatmap_data)
    if period_reports is None:
        period_snapshot = build_snapshot(
            rows,
            generated_at
            or datetime.now().astimezone().isoformat(timespec="seconds"),
            source_path,
        )
        period_data = build_period_reports(
            period_snapshot,
            {"version": HISTORY_VERSION, "snapshots": []},
            None,
        )
    else:
        period_data = period_reports
    period_html = build_period_report_section(period_data)
    annual_html = build_annual_report_section(period_data, trend_data)
    trend_sections = build_trend_sections(trend_data, heatmap_data)
    useful_rows = sorted(
        rows,
        key=lambda row: (
            -(parse_int(row.get("total_time_ms")) or 0),
            display_title(str(row.get("title", ""))).casefold(),
        ),
    )
    max_time = max(
        (parse_int(row.get("total_time_ms")) or 0 for row in useful_rows),
        default=0,
    )
    quality_by_book = {
        stable_book_id(row): analyze_book_quality(row)
        for row in useful_rows
    }
    achievement_data = build_achievements(
        summary,
        trend_data,
        useful_rows,
        profile=achievement_profile,
        threshold_percent=completed_threshold,
        include_manga=not text_only_ui,
    )
    achievements_html = build_achievement_cards(achievement_data)
    star_sky_html = build_star_sky_html(achievement_data)
    achievement_criteria_zh = (
        "四条轨道，每颗星都在首次基线之上点亮。"
        if text_only_ui
        else "五条轨道，每颗星都在首次基线之上点亮；漫画卷数轨道默认关闭，且与已读书籍不重叠。"
    )
    achievement_criteria_en = (
        "Four tracks unlock above your first baseline."
        if text_only_ui
        else "Five tracks unlock above your first baseline; manga volumes is off by default and never overlaps completed books."
    )
    achievement_summary = bilingual_text(
        f"{achievement_data['unlocked']}/{achievement_data['total']} 项已解锁",
        f"{achievement_data['unlocked']}/{achievement_data['total']} unlocked",
    )

    cover_cards: List[str] = []
    chart_items: List[str] = []
    for index, row in enumerate(useful_rows[:12]):
        title = display_title(str(row.get("title", "")))
        time_ms = parse_int(row.get("total_time_ms")) or 0
        width = (time_ms / max_time * 100.0) if max_time else 0.0
        color = ACCENTS[index % len(ACCENTS)]
        chart_items.append(
            '<div class="rank-item">'
            f'<div class="rank-label"><strong>{html.escape(title)}</strong>'
            f"<span>{bilingual_text(format_duration(time_ms), format_duration_en(time_ms))}</span></div>"
            '<div class="rank-track">'
            f'<i style="width:{width:.2f}%;background:{color}"></i>'
            "</div>"
            "</div>"
        )

    for index, row in enumerate(useful_rows, start=1):
        title = display_title(str(row.get("title", "")))
        author = str(row.get("author") or "").strip() or "--"
        total_time = parse_int(row.get("total_time_ms")) or 0
        total_words = parse_int(row.get("total_words"))
        book_words = parse_int(effective_book_words(row))
        manga = is_manga(row)
        progress = (
            manga_progress_percent(row)
            if manga
            else word_progress(row)
        )
        value = progress or 0.0
        progress_label = (
            f"{progress:.1f}%" if progress is not None else "--"
        )
        if manga and progress is None:
            progress_label = (
                "未开始"
                if (parse_int(row.get("total_time_ms")) or 0) <= 0
                else "无进度数据"
            )
        book_id = stable_book_id(row)
        language = resolved_card_language(row)
        source = row.get("word_count_source") or ""
        manga = is_manga(row)
        token_source = resolved_tokenizer_source(row)
        token_label_zh = tokenizer_label(token_source)
        token_label_en = tokenizer_label_en(token_source)
        exact = has_exact_word_count(row)
        quality = quality_by_book.get(book_id, {})
        origin_label = str(
            quality.get("word_origin_label") or "词数未知"
        )
        origin_label_en = str(
            quality.get("word_origin_label_en") or "Word count unknown"
        )
        deviation = parse_float(quality.get("deviation_percent"))
        confidence_level = str(
            quality.get("confidence_level") or "low"
        )
        confidence_label = str(
            quality.get("confidence_label") or "低置信"
        )
        confidence_label_en = str(
            quality.get("confidence_label_en") or "Low confidence"
        )
        confidence_score = parse_int(
            quality.get("confidence_score")
        ) or 0
        state = reading_state(row)
        if (
            manga
            and progress is not None
            and progress >= completed_threshold
        ):
            state = "finished"
        cover_url = resolved_cover_urls.get(book_id)
        search = " ".join(
            [
                title,
                author,
                str(row.get("source_file", "")),
                resolved_card_language_label(row),
                resolved_card_language_label_en(row),
            ]
        )
        initial = (title.strip()[:1] or "?")
        placeholder = (
            '<div class="cover-placeholder"'
            f' aria-label="{html.escape(title, quote=True)} 暂无封面 / No cover">'
            f"<strong>{html.escape(initial)}</strong>"
            f"<small>{bilingual_text('暂无封面', 'No cover')}</small>"
            "</div>"
        )
        if cover_url:
            cover_markup = (
                '<div class="cover-frame has-cover">'
                f'<img data-cover-image src="{html.escape(cover_url, quote=True)}" '
                f'alt="{html.escape(title, quote=True)} 封面 / Cover" '
                'onerror="this.hidden=true;this.dataset.coverFailed=true;'
                'this.nextElementSibling.hidden=false;">'
                f'{placeholder.replace("cover-placeholder", "cover-placeholder is-fallback", 1).replace(">", " hidden>", 1)}'
                "</div>"
            )
        else:
            cover_markup = f'<div class="cover-frame">{placeholder}</div>'

        raw_word_tip_zh = "累计已读词数"
        raw_word_tip_en = "Cumulative words read"
        reading_tip_zh = "累计阅读时间"
        reading_tip_en = "Cumulative reading time"
        ocr_eligible = ocr_estimate_eligible(row)
        effective_words_read_value = effective_words_read(row)
        if manga:
            pages_total = manga_total_pages(row)
            pages_read = manga_pages_read(row)
            manga_stats = (
                '<div class="cover-stats">'
                '<div class="cover-stat">'
                f'<span class="cover-stat-label">'
                f"{bilingual_text('页数', 'Pages')}</span>"
                f'<strong>{pages_total if pages_total is not None else "--"}</strong>'
                "</div>"
                '<div class="cover-stat">'
                f'<span class="cover-stat-label">'
                f"{bilingual_text('已读', 'Read')}</span>"
                f'<strong>{pages_read if pages_read is not None else "--"}</strong>'
                "</div>"
                '<div class="cover-stat">'
                f'<span class="cover-stat-label">'
                f"{bilingual_text('时长', 'Time')}</span>"
                f'<strong>{bilingual_text(format_duration(total_time), format_duration_en(total_time)) if total_time > 0 else "--"}</strong>'
                "</div>"
                "</div>"
            )
        else:
            manga_stats = (
                '<div class="cover-stats">'
                '<div class="cover-stat">'
                f'<span class="cover-stat-label">'
                f"{bilingual_text('阅读', 'Reading')}</span>"
                f'<strong data-title-zh="{html.escape(reading_tip_zh, quote=True)}" '
                f'data-title-en="{html.escape(reading_tip_en, quote=True)}" '
                f'title="{html.escape(reading_tip_zh, quote=True)}">'
                f"{bilingual_text(format_duration(total_time), format_duration_en(total_time))}</strong>"
                "</div>"
                '<div class="cover-stat">'
                f'<span class="cover-stat-label">'
                f"{bilingual_text('已读', 'Read')}</span>"
                f'<strong title="{html.escape(raw_word_tip_zh, quote=True)}">'
                f"{html.escape(format_words(parse_int(effective_words_read_value)))}</strong>"
                "</div>"
                '<div class="cover-stat">'
                f'<span class="cover-stat-label">'
                f"{bilingual_text('全书', 'Total')}</span>"
                f'<strong title="{html.escape(raw_word_tip_zh, quote=True)}">'
                f"{'--' if manga and not ocr_eligible else html.escape(format_words(book_words))}</strong>"
                "</div>"
                "</div>"
            )
        state_label_zh, state_label_en = reading_state_label(row)
        if manga and progress is not None and progress >= normalized_completed_threshold(completed_threshold_percent):
            state_label_zh, state_label_en = "已读完", "Finished"
        cover_cards.append(
            f'<article class="book-cover-card" '
            f'data-book-id="{html.escape(book_id, quote=True)}" '
            f'data-cover-title="{html.escape(title, quote=True)}" '
            f'data-last-read="{html.escape(str(row.get("last_read") or ""), quote=True)}" '
            f'data-search="{html.escape(search.casefold(), quote=True)}" '
            f'data-language="{html.escape(str(language))}" '
            f'data-source="{html.escape(str(source))}" '
            f'data-token-label="{html.escape(str(token_source), quote=True)}" '
            f'data-mode="{html.escape(reading_mode_for_book(row), quote=True)}" '
            f'data-count="{"manga" if manga else ("exact" if exact else "estimate")}" '
            f'data-manga-pages-total="{(manga_total_pages(row) if manga_total_pages(row) is not None else "") if manga else ""}" '
            f'data-manga-pages-read="{(manga_pages_read(row) if manga_pages_read(row) is not None else "") if manga else ""}" '
            f'data-state="{state}">'
            f"{cover_markup}"
            '<div class="cover-info">'
            f'<div class="cover-rank">#{index}</div>'
            f'<h3 title="{html.escape(title, quote=True)}">'
            f"{html.escape(title)}</h3>"
            f"<p>{html.escape(author)}</p>"
            f"{manga_stats}"
            '<div class="cover-progress">'
            f'<span>{html.escape(progress_label)}</span>'
            '<div><i style="width:'
            f'{value:.2f}%;background:{ACCENTS[(index - 1) % len(ACCENTS)]}"></i></div>'
            "</div>"
            '<div class="tag-row">'
            f'<span class="tag language-tag">'
            f'{bilingual_text(resolved_card_language_label(row), resolved_card_language_label_en(row))}</span>'
            f'<span class="tag mode-tag">'
            f'{bilingual_text(reading_mode_label(reading_mode_for_book(row)), reading_mode_label_en(reading_mode_for_book(row)))}</span>'
            f'<span class="tag state-tag">'
            f'{bilingual_text(state_label_zh, state_label_en)}</span>'
            f'<span class="tag source-tag">'
            f'{bilingual_text(str(row.get("manga_pages_source") or row.get("classification_source") or "source"), str(row.get("manga_pages_source") or row.get("classification_source") or "source"))}</span>'
            + (
                '<span class="tag token-tag">'
                f'{bilingual_text(token_label_zh, token_label_en)}</span>'
                if token_label_zh
                else ""
            )
            +
            "</div>"
            "</div>"
            "</article>"
        )

    progress_cards: List[str] = []
    for index, row in enumerate(useful_rows[:8]):
        manga = is_manga(row)
        progress = (
            manga_progress_percent(row)
            if manga
            else word_progress(row)
        )
        value = progress or 0.0
        progress_label = (
            f"{progress:.1f}%"
            if progress is not None
            else ("未开始" if manga else "--")
        )
        color = ACCENTS[index % len(ACCENTS)]
        title = display_title(str(row.get("title", "")))
        words = parse_int(row.get("total_words"))
        book_words = parse_int(effective_book_words(row))
        effective_read = parse_int(effective_words_read(row))
        ocr_eligible = ocr_estimate_eligible(row)
        exact = has_exact_word_count(row)
        language = resolved_card_language_label(row)
        language_en = resolved_card_language_label_en(row)
        token_source = resolved_tokenizer_source(row)
        token_label_zh = tokenizer_label(token_source)
        token_label_en = tokenizer_label_en(token_source)
        source = source_label(row.get("word_count_source"))
        source_en = source_label_en(row.get("word_count_source"))
        quality = quality_by_book.get(stable_book_id(row), {})
        origin_label = str(
            quality.get("word_origin_label") or "词数未知"
        )
        origin_label_en = str(
            quality.get("word_origin_label_en") or "Word count unknown"
        )
        confidence_level = str(
            quality.get("confidence_level") or "low"
        )
        confidence_label = str(
            quality.get("confidence_label") or "低置信"
        )
        confidence_label_en = str(
            quality.get("confidence_label_en") or "Low confidence"
        )
        confidence_score = parse_int(
            quality.get("confidence_score")
        ) or 0
        if manga:
            detail_summary = (
                '<p class="word-count-summary">'
                f"<span>{bilingual_text('总页数', 'Pages')} "
                f"{manga_total_pages(row) if manga_total_pages(row) is not None else '--'}</span>"
                f"<span>{bilingual_text('已读页数', 'Read pages')} "
                f"{manga_pages_read(row) if manga_pages_read(row) is not None else '--'}</span>"
                "</p>"
            )
        else:
            detail_summary = (
                '<p class="word-count-summary">'
                f"<span>{bilingual_text('已读词数', 'Words read')} "
                f"{html.escape(format_words(effective_read))}</span>"
                f"<span>{bilingual_text('全书词数', 'Book words')} "
                f"{html.escape(format_words(book_words))}</span>"
                "</p>"
            )
        progress_cards.append(
            '<article class="progress-card">'
            f'<div class="progress-ring" style="--value:{value:.2f};'
            f'--ring-color:{color}">'
            f"<strong>{progress_label}</strong>"
            "</div>"
            f'<div><h3>{html.escape(title)}</h3>'
            f"{detail_summary}"
            '<div class="tag-row">'
            f'<span class="tag language-tag">'
            f'{bilingual_text(language, language_en)}</span>'
            f'<span class="tag mode-tag">'
            f'{bilingual_text(reading_mode_label(reading_mode_for_book(row)), reading_mode_label_en(reading_mode_for_book(row)))}</span>'
            + (
                '<span class="tag token-tag">'
                f'{bilingual_text(token_label_zh, token_label_en)}</span>'
                if token_label_zh
                else ""
            )
            +
            "</div></div>"
            "</article>"
        )

    table_rows: List[str] = []
    for index, row in enumerate(useful_rows, start=1):
        title = display_title(str(row.get("title", "")))
        author = str(row.get("author") or "").strip() or "--"
        total_time = parse_int(row.get("total_time_ms")) or 0
        total_words = parse_int(row.get("total_words"))
        book_words = parse_int(effective_book_words(row))
        wpm = parse_float(row.get("wpm"))
        effective_wpm_value = effective_wpm(row)
        effective_read = parse_int(effective_words_read(row))
        ocr_eligible = ocr_estimate_eligible(row)
        manga = is_manga(row)
        kindle_progress = (
            manga_progress_percent(row)
            if manga
            else parse_float(row.get("progress_percent"))
        )
        words_progress = None if manga else word_progress(row)
        annotations = parse_int(row.get("annotations"))
        last_read = format_last_read(str(row.get("last_read") or ""))
        language = resolved_card_language(row)
        source = row.get("word_count_source") or ""
        token_source = resolved_tokenizer_source(row)
        token_label_zh = tokenizer_label(token_source)
        token_label_en = tokenizer_label_en(token_source)
        exact = has_exact_word_count(row)
        state = reading_state(row)
        change = trend_by_book.get(stable_book_id(row))
        quality = quality_by_book.get(stable_book_id(row), {})
        origin_label = str(
            quality.get("word_origin_label") or "词数未知"
        )
        origin_label_en = str(
            quality.get("word_origin_label_en") or "Word count unknown"
        )
        deviation = parse_float(quality.get("deviation_percent"))
        confidence_level = str(
            quality.get("confidence_level") or "low"
        )
        confidence_label = str(
            quality.get("confidence_label") or "低置信"
        )
        confidence_label_en = str(
            quality.get("confidence_label_en") or "Low confidence"
        )
        confidence_score = parse_int(
            quality.get("confidence_score")
        ) or 0
        quality_flags = [
            str(flag)
            for flag in quality.get("flags", [])
            if str(flag).strip()
        ]
        quality_flags_en = [
            str(flag)
            for flag in quality.get("flags_en", [])
            if str(flag).strip()
        ]
        search = " ".join(
            [
                title,
                author,
                str(row.get("source_file", "")),
                resolved_card_language_label(row),
                resolved_card_language_label_en(row),
                reading_mode_label(reading_mode_for_book(row)),
                reading_mode_label_en(reading_mode_for_book(row)),
            ]
        )
        color = ACCENTS[index % len(ACCENTS)]
        wpm_text = (
            f"{effective_wpm_value:.0f}"
            if effective_wpm_value is not None and not manga
            else "--"
        )
        read_words_text = (
            "--"
            if manga and not ocr_eligible
            else format_words(effective_read)
        )
        deviation_text = (
            f"{deviation:+.1f}%" if deviation is not None else "--"
        )

        table_rows.append(
            f'<tr data-search="{html.escape(search.casefold(), quote=True)}" '
            f'data-language="{html.escape(str(language))}" '
            f'data-source="{html.escape(str(source))}" '
            f'data-token-label="{html.escape(str(token_source), quote=True)}" '
            f'data-mode="{html.escape(reading_mode_for_book(row), quote=True)}" '
            f'data-count="{"manga" if manga else ("exact" if exact else "estimate")}" '
            f'data-state="{state}">'
            f'<td class="rank" data-column="rank">{index}</td>'
            f'<td class="book-cell" data-column="book">'
            f"<strong>{html.escape(title)}</strong>"
            f"<span>{html.escape(author)}</span></td>"
            f'<td data-column="duration" data-value="{total_time}">'
            f"{bilingual_text(format_duration(total_time), format_duration_en(total_time)) if total_time > 0 else '--'}</td>"
            f'<td data-column="words" data-value="{effective_read if effective_read is not None else (total_words if total_words is not None else -1)}">'
            '<div class="word-count-cell">'
            f"<span><em>{bilingual_text('已读', 'Read')}</em>"
            f"<strong>{html.escape(read_words_text)}</strong></span>"
            f"<span><em>{bilingual_text('全书', 'Total')}</em>"
            f"<strong>{'--' if manga and not ocr_eligible else html.escape(format_words(book_words))}</strong></span>"
            "</div></td>"
            f'<td data-column="wpm" data-value="{effective_wpm_value if effective_wpm_value is not None else -1}">'
            f'<div class="source-value"><strong>{html.escape(wpm_text)}</strong></div></td>'
            f'<td data-column="kindle-progress" data-value="{kindle_progress if kindle_progress is not None else -1}">'
            f"{_progress_bar(kindle_progress, color)}"
            "</td>"
            f'<td data-column="word-progress" data-value="{words_progress if words_progress is not None else -1}">'
            f"{_progress_bar(words_progress, color)}"
            "</td>"
            '<td class="tags-cell" data-column="tags">'
            f'<span class="tag language-tag">'
            f'{bilingual_text(resolved_card_language_label(row), resolved_card_language_label_en(row))}</span>'
            f'<span class="tag mode-tag">'
            f'{bilingual_text(reading_mode_label(reading_mode_for_book(row)), reading_mode_label_en(reading_mode_for_book(row)))}</span>'
            + (
                '<span class="tag token-tag">'
                f'{bilingual_text(token_label_zh, token_label_en)}</span>'
                if token_label_zh
                else ""
            )
            +
            "</td>"
            f'<td data-column="annotations" data-value="{annotations or 0}">{annotations or 0}</td>'
            f'<td data-column="last-read">{html.escape(last_read)}</td>'
            f"{_table_delta_cell(change)}"
            "</tr>"
        )

    generated = (
        generated_at
        or datetime.now().astimezone().strftime("%Y-%m-%d %H:%M")
    )
    html_text = """<!doctype html>
<html lang="zh-CN" data-ui-language="zh" data-ui-variant="__UI_VARIANT__" data-theme="light"
      data-theme-preset="forest" data-achievement-badge-style="seal"
      data-background-texture="paper" data-background-motif="none"
      data-shelf-title-style="a" data-shelf-position="title"
      data-shelf-default-mode="book" data-show-token-labels="false"
      data-annual-report-panel="closed"
      data-greeting-tone="day" data-theme-follow-greeting="false"
      data-minimal-mode="false" data-section-period="visible"
      data-section-annual="visible"
      data-section-trends="visible" data-section-heatmap="visible">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Kindle 阅读数据看板 / Reading Dashboard</title>
  <style>
    :root {
      color-scheme: light;
      --greeting-font: "SimSun", "Songti SC", "Noto Serif CJK SC", serif;
      --bg: #f4f6f3;
      --surface: #ffffff;
      --surface-soft: #fbfcfb;
      --surface-muted: #f0f3f1;
      --surface-track: #e8ece9;
      --surface-chart: #f1f4f2;
      --ink: #23302a;
      --muted: #68756e;
      --line: #dfe5e0;
      --green: #2f7d61;
      --blue: #3478a8;
      --amber: #c57a20;
      --coral: #c65c4a;
      --violet: #6f5aa8;
      --teal: #3d8f8a;
      --gold: #FFD43B;
      --gold-bright: #FFE37A;
      --gold-deep: #C58A1A;
      --gold-glow: rgba(255, 212, 59, .28);
      --gold-ink: #6b4d00;
      --hover: #f8faf8;
      --badge-bg: #edf8f2;
      --badge-border: #b9d9ca;
      --tag-bg: #eef2f0;
      --tag-ink: #4d5b54;
      --language-bg: #eef6fb;
      --estimate-bg: #fff7e8;
      --low-bg: #fff0ed;
      --deviation-bg: #f3f0fa;
      --placeholder-bg: #e8ece9;
      --placeholder-ink: #627069;
      --placeholder-mark: #f7f9f7;
      --ring-empty: #e7ebe8;
      --timeline-ring: #d8eee4;
      --dashboard-background-image: none;
      --dashboard-background-opacity: 0;
      --dashboard-background-blur: 0px;
      --dashboard-background-overlay: 0;
      --dashboard-texture-image: none;
      --dashboard-texture-size: auto;
      --dashboard-motif-image: none;
      --dashboard-motif-size: auto;
      --dashboard-motif-repeat: repeat;
      --dashboard-flourish-image: none;
      --texture-line: rgba(35, 48, 42, .028);
      --pattern-line: rgba(35, 48, 42, .08);
      --focus-ring: var(--blue);
      --appearance-top: 90px;
      --shadow-card: 0 1px 2px rgba(25, 40, 32, .06);
      --shadow-raised: 0 8px 22px rgba(25, 40, 32, .14);
    }

    html[data-ui-variant="text"] .manga-only { display: none !important; }
    html[data-ui-variant="text"] .text-only-only { display: revert !important; }
    .text-only-only { display: none !important; }
    .sr-only {
      position: absolute !important;
      width: 1px !important;
      height: 1px !important;
      padding: 0 !important;
      margin: -1px !important;
      overflow: hidden !important;
      clip: rect(0, 0, 0, 0) !important;
      white-space: nowrap !important;
      border: 0 !important;
    }
    html[data-ui-language="zh"] .i18n-en { display: none; }
    html[data-ui-language="en"] .i18n-zh { display: none; }
    html[data-theme="dark"] {
      color-scheme: dark;
      --bg: #101713;
      --surface: #18211c;
      --surface-soft: #1d2722;
      --surface-muted: #243029;
      --surface-track: #2c3a32;
      --surface-chart: #223029;
      --ink: #edf5f0;
      --muted: #a8b8af;
      --line: #35463d;
      --green: #70c9a5;
      --blue: #80b9df;
      --amber: #e2ad59;
      --coral: #ea8777;
      --violet: #b9a7ec;
      --teal: #63c2bd;
      --gold: #FFD43B;
      --gold-bright: #FFE37A;
      --gold-deep: #C58A1A;
      --gold-glow: rgba(255, 212, 59, .28);
      --gold-ink: #FFE566;
      --hover: #202c26;
      --badge-bg: #1d3a2f;
      --badge-border: #35664f;
      --tag-bg: #27342d;
      --tag-ink: #c9d8cf;
      --language-bg: #203443;
      --estimate-bg: #3c321f;
      --low-bg: #422721;
      --deviation-bg: #302a45;
      --placeholder-bg: #243029;
      --placeholder-ink: #aabbb1;
      --placeholder-mark: #1b251f;
      --ring-empty: #2c3a32;
      --timeline-ring: #315a49;
      --texture-line: rgba(237, 245, 240, .038);
      --pattern-line: rgba(237, 245, 240, .09);
      --shadow-card: 0 1px 2px rgba(0, 0, 0, .3);
      --shadow-raised: 0 10px 26px rgba(0, 0, 0, .38);
    }
    /* Soft, time-of-day greeting tints; dark theme overrides keep contrast. */
    html[data-greeting-tone="late-night"] { --greeting-ink: #3b3f6b; }
    html[data-greeting-tone="morning"] { --greeting-ink: #8a5a12; }
    html[data-greeting-tone="day"] { --greeting-ink: #175e6b; }
    html[data-greeting-tone="dusk"] { --greeting-ink: #9a4a2f; }
    html[data-greeting-tone="night"] { --greeting-ink: #9a4f6d; }
    html[data-theme="dark"][data-greeting-tone="late-night"] {
      --greeting-ink: #a9b3ef;
    }
    html[data-theme="dark"][data-greeting-tone="morning"] {
      --greeting-ink: #f0cf8c;
    }
    html[data-theme="dark"][data-greeting-tone="day"] {
      --greeting-ink: #8fd3d8;
    }
    html[data-theme="dark"][data-greeting-tone="dusk"] {
      --greeting-ink: #f0b48f;
    }
    html[data-theme="dark"][data-greeting-tone="night"] {
      --greeting-ink: #f3a7c2;
    }
    html[data-background-texture="paper"] {
      --dashboard-texture-image:
        repeating-linear-gradient(
          0deg,
          var(--texture-line) 0 1px,
          transparent 1px 5px
        ),
        repeating-linear-gradient(
          90deg,
          var(--texture-line) 0 1px,
          transparent 1px 7px
        );
      --dashboard-texture-size: 13px 13px;
    }
    html[data-background-texture="grid"] {
      --dashboard-texture-image:
        linear-gradient(var(--pattern-line) 1px, transparent 1px),
        linear-gradient(90deg, var(--pattern-line) 1px, transparent 1px);
      --dashboard-texture-size: 24px 24px;
    }
    html[data-background-texture="grain"] {
      --dashboard-texture-image:
        radial-gradient(
          circle at 1px 1px,
          var(--pattern-line) 0 1px,
          transparent 1.35px
        );
      --dashboard-texture-size: 11px 11px;
    }
    /* Decorative, opt-in tiling patterns. They reuse the same single
       background slot as the subtle textures above, so the report stays
       a self-contained file and print output remains plain. */
    html[data-background-texture="dots"] {
      --dashboard-texture-image:
        radial-gradient(
          circle at 50% 50%,
          var(--pattern-line) 0 1.1px,
          transparent 1.5px
        );
      --dashboard-texture-size: 22px 22px;
    }
    html[data-background-texture="rings"] {
      --dashboard-texture-image:
        radial-gradient(
          circle at 50% 50%,
          transparent 0 4.4px,
          var(--pattern-line) 4.4px 5.1px,
          transparent 5.1px
        );
      --dashboard-texture-size: 26px 26px;
    }
    html[data-background-texture="diagonal"] {
      --dashboard-texture-image:
        repeating-linear-gradient(
          45deg,
          var(--pattern-line) 0 1px,
          transparent 1px 11px
        );
      --dashboard-texture-size: auto;
    }
    html[data-background-texture="diamond"] {
      --dashboard-texture-image:
        repeating-linear-gradient(
          45deg,
          var(--pattern-line) 0 1px,
          transparent 1px 16px
        ),
        repeating-linear-gradient(
          -45deg,
          var(--pattern-line) 0 1px,
          transparent 1px 16px
        );
      --dashboard-texture-size: auto;
    }
    /* Pictorial motif layer. It is a separate, opt-in switch layered on top
       of the background pattern; the print override on body hides both. */
    html[data-background-motif="books"] {
      --dashboard-motif-image:
        url("data:image/svg+xml,%3Csvg%20xmlns='http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg'%20width='48'%20height='48'%20viewBox='0%200%2048%2048'%3E%3Cpath%20d='M5%2013c5-2%2010-1%2014%202v20c-4-3-9-4-14-2z'%20fill='%234d5b54'%20fill-opacity='.1'/%3E%3Cpath%20d='M43%2013c-5-2-10-1-14%202v20c4-3%209-4%2014-2z'%20fill='%234d5b54'%20fill-opacity='.1'/%3E%3C/svg%3E");
      --dashboard-motif-size: 74px 74px;
    }
    html[data-theme="dark"][data-background-motif="books"] {
      --dashboard-motif-image:
        url("data:image/svg+xml,%3Csvg%20xmlns='http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg'%20width='48'%20height='48'%20viewBox='0%200%2048%2048'%3E%3Cpath%20d='M5%2013c5-2%2010-1%2014%202v20c-4-3-9-4-14-2z'%20fill='%23c9d8cf'%20fill-opacity='.12'/%3E%3Cpath%20d='M43%2013c-5-2-10-1-14%202v20c4-3%209-4%2014-2z'%20fill='%23c9d8cf'%20fill-opacity='.12'/%3E%3C/svg%3E");
    }
    html[data-background-motif="leaves"] {
      --dashboard-motif-image:
        url("data:image/svg+xml,%3Csvg%20xmlns='http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg'%20width='44'%20height='44'%20viewBox='0%200%2044%2044'%3E%3Cpath%20d='M22%205c10%209%2010%2024%200%2034C12%2029%2012%2014%2022%205z'%20fill='%234d5b54'%20fill-opacity='.1'/%3E%3Cpath%20d='M22%2037v3'%20stroke='%234d5b54'%20stroke-opacity='.1'%20stroke-width='1.6'%20stroke-linecap='round'/%3E%3C/svg%3E");
      --dashboard-motif-size: 66px 66px;
    }
    html[data-theme="dark"][data-background-motif="leaves"] {
      --dashboard-motif-image:
        url("data:image/svg+xml,%3Csvg%20xmlns='http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg'%20width='44'%20height='44'%20viewBox='0%200%2044%2044'%3E%3Cpath%20d='M22%205c10%209%2010%2024%200%2034C12%2029%2012%2014%2022%205z'%20fill='%23c9d8cf'%20fill-opacity='.12'/%3E%3Cpath%20d='M22%2037v3'%20stroke='%23c9d8cf'%20stroke-opacity='.12'%20stroke-width='1.6'%20stroke-linecap='round'/%3E%3C/svg%3E");
    }
    html[data-background-motif="stars"] {
      --dashboard-motif-image:
        url("data:image/svg+xml,%3Csvg%20xmlns='http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg'%20width='52'%20height='52'%20viewBox='0%200%2052%2052'%3E%3Cpath%20d='M12%205l2%205l5%202l-5%202l-2%205l-2-5l-5-2l5-2z'%20fill='%234d5b54'%20fill-opacity='.1'/%3E%3Cpath%20d='M38%2021l1.4%203.6l3.6%201.4l-3.6%201.4l-1.4%203.6l-1.4-3.6l-3.6-1.4l3.6-1.4z'%20fill='%234d5b54'%20fill-opacity='.1'/%3E%3Ccircle%20cx='44'%20cy='8'%20r='1.4'%20fill='%234d5b54'%20fill-opacity='.1'/%3E%3C/svg%3E");
      --dashboard-motif-size: 78px 78px;
    }
    html[data-theme="dark"][data-background-motif="stars"] {
      --dashboard-motif-image:
        url("data:image/svg+xml,%3Csvg%20xmlns='http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg'%20width='52'%20height='52'%20viewBox='0%200%2052%2052'%3E%3Cpath%20d='M12%205l2%205l5%202l-5%202l-2%205l-2-5l-5-2l5-2z'%20fill='%23c9d8cf'%20fill-opacity='.12'/%3E%3Cpath%20d='M38%2021l1.4%203.6l3.6%201.4l-3.6%201.4l-1.4%203.6l-1.4-3.6l-3.6-1.4l3.6-1.4z'%20fill='%23c9d8cf'%20fill-opacity='.12'/%3E%3Ccircle%20cx='44'%20cy='8'%20r='1.4'%20fill='%23c9d8cf'%20fill-opacity='.12'/%3E%3C/svg%3E");
    }
    /* Colorful motif overrides keep the reading dashboard quiet around
       data, while the background and heading friezes carry the palette. */
    html[data-background-motif="books"] {
      --dashboard-motif-image:
        url("data:image/svg+xml,%3Csvg%20xmlns='http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg'%20width='48'%20height='48'%20viewBox='0%200%2048%2048'%3E%3Cpath%20d='M5%2013c5-2%2010-1%2014%202v20c-4-3-9-4-14-2z'%20fill='%232f9b7a'%20fill-opacity='.14'/%3E%3Cpath%20d='M43%2013c-5-2-10-1-14%202v20c4-3%209-4%2014-2z'%20fill='%234f9bd0'%20fill-opacity='.14'/%3E%3Ccircle%20cx='24'%20cy='9'%20r='2'%20fill='%23e0a33a'%20fill-opacity='.18'/%3E%3Cpath%20d='M24%2036l2-4%204-2-4-2-2-4-2%204-4%202%204%202z'%20fill='%238a76d0'%20fill-opacity='.12'/%3E%3C/svg%3E");
      --dashboard-flourish-image:
        url("data:image/svg+xml,%3Csvg%20xmlns='http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg'%20width='64'%20height='22'%20viewBox='0%200%2064%2022'%3E%3Cpath%20d='M0%2011h24'%20stroke='%232f9b7a'%20stroke-opacity='.34'%20stroke-width='1.4'/%3E%3Cpath%20d='M24%2011h16'%20stroke='%234f9bd0'%20stroke-opacity='.34'%20stroke-width='1.4'/%3E%3Cpath%20d='M40%2011h24'%20stroke='%23e0a33a'%20stroke-opacity='.34'%20stroke-width='1.4'/%3E%3Cpath%20d='M22%207c4-1.8%207-1.2%2010%201.8v7.4c-3-3-6-3.6-10-2z'%20fill='%232f9b7a'%20fill-opacity='.42'/%3E%3Cpath%20d='M42%207c-4-1.8-7-1.2-10%201.8v7.4c3-3%206-3.6%2010-2z'%20fill='%234f9bd0'%20fill-opacity='.42'/%3E%3Cpath%20d='M9%2011l3-3%203%203-3%203z'%20fill='%23e0a33a'%20fill-opacity='.38'/%3E%3Cpath%20d='M49%2011l3-3%203%203-3%203z'%20fill='%238a76d0'%20fill-opacity='.34'/%3E%3C/svg%3E");
    }
    html[data-theme="dark"][data-background-motif="books"] {
      --dashboard-motif-image:
        url("data:image/svg+xml,%3Csvg%20xmlns='http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg'%20width='48'%20height='48'%20viewBox='0%200%2048%2048'%3E%3Cpath%20d='M5%2013c5-2%2010-1%2014%202v20c-4-3-9-4-14-2z'%20fill='%2370c9a5'%20fill-opacity='.16'/%3E%3Cpath%20d='M43%2013c-5-2-10-1-14%202v20c4-3%209-4%2014-2z'%20fill='%2380b9df'%20fill-opacity='.16'/%3E%3Ccircle%20cx='24'%20cy='9'%20r='2'%20fill='%23e2ad59'%20fill-opacity='.18'/%3E%3Cpath%20d='M24%2036l2-4%204-2-4-2-2-4-2%204-4%202%204%202z'%20fill='%23b9a7ec'%20fill-opacity='.14'/%3E%3C/svg%3E");
    }
    html[data-background-motif="leaves"] {
      --dashboard-motif-image:
        url("data:image/svg+xml,%3Csvg%20xmlns='http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg'%20width='44'%20height='44'%20viewBox='0%200%2044%2044'%3E%3Cpath%20d='M22%205c10%209%2010%2024%200%2034C12%2029%2012%2014%2022%205z'%20fill='%232f9b7a'%20fill-opacity='.14'/%3E%3Cpath%20d='M22%2037v3'%20stroke='%234f9bd0'%20stroke-opacity='.16'%20stroke-width='1.6'%20stroke-linecap='round'/%3E%3Ccircle%20cx='31'%20cy='12'%20r='2.2'%20fill='%23e0a33a'%20fill-opacity='.18'/%3E%3Ccircle%20cx='14'%20cy='27'%20r='1.8'%20fill='%23e07a68'%20fill-opacity='.16'/%3E%3C/svg%3E");
      --dashboard-flourish-image:
        url("data:image/svg+xml,%3Csvg%20xmlns='http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg'%20width='64'%20height='22'%20viewBox='0%200%2064%2022'%3E%3Cpath%20d='M0%2011h20'%20stroke='%232f9b7a'%20stroke-opacity='.32'%20stroke-width='1.4'/%3E%3Cpath%20d='M20%2011h20'%20stroke='%23e0a33a'%20stroke-opacity='.32'%20stroke-width='1.4'/%3E%3Cpath%20d='M40%2011h24'%20stroke='%238a76d0'%20stroke-opacity='.32'%20stroke-width='1.4'/%3E%3Cpath%20d='M32%204c6%205%206%2011%200%2016-6-5-6-11%200-16z'%20fill='%232f9b7a'%20fill-opacity='.44'/%3E%3Cpath%20d='M32%206v13'%20stroke='%234f9bd0'%20stroke-opacity='.5'%20stroke-width='1.1'%20stroke-linecap='round'/%3E%3Ccircle%20cx='16'%20cy='11'%20r='1.7'%20fill='%23e0a33a'%20fill-opacity='.5'/%3E%3Cpath%20d='M50%2011l3-3%203%203-3%203z'%20fill='%23e07a68'%20fill-opacity='.34'/%3E%3C/svg%3E");
    }
    html[data-theme="dark"][data-background-motif="leaves"] {
      --dashboard-motif-image:
        url("data:image/svg+xml,%3Csvg%20xmlns='http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg'%20width='44'%20height='44'%20viewBox='0%200%2044%2044'%3E%3Cpath%20d='M22%205c10%209%2010%2024%200%2034C12%2029%2012%2014%2022%205z'%20fill='%2370c9a5'%20fill-opacity='.16'/%3E%3Cpath%20d='M22%2037v3'%20stroke='%2380b9df'%20stroke-opacity='.18'%20stroke-width='1.6'%20stroke-linecap='round'/%3E%3Ccircle%20cx='31'%20cy='12'%20r='2.2'%20fill='%23e2ad59'%20fill-opacity='.18'/%3E%3Ccircle%20cx='14'%20cy='27'%20r='1.8'%20fill='%23ea8777'%20fill-opacity='.16'/%3E%3C/svg%3E");
    }
    html[data-background-motif="stars"] {
      --dashboard-motif-image:
        url("data:image/svg+xml,%3Csvg%20xmlns='http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg'%20width='52'%20height='52'%20viewBox='0%200%2052%2052'%3E%3Cpath%20d='M12%205l2%205l5%202l-5%202l-2%205l-2-5l-5-2l5-2z'%20fill='%23e0a33a'%20fill-opacity='.16'/%3E%3Cpath%20d='M38%2021l1.4%203.6l3.6%201.4l-3.6%201.4l-1.4%203.6l-1.4-3.6l-3.6-1.4l3.6-1.4z'%20fill='%238a76d0'%20fill-opacity='.14'/%3E%3Ccircle%20cx='44'%20cy='8'%20r='1.4'%20fill='%234f9bd0'%20fill-opacity='.18'/%3E%3Ccircle%20cx='10'%20cy='38'%20r='1.2'%20fill='%23e07a68'%20fill-opacity='.16'/%3E%3C/svg%3E");
      --dashboard-flourish-image:
        url("data:image/svg+xml,%3Csvg%20xmlns='http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg'%20width='64'%20height='22'%20viewBox='0%200%2064%2022'%3E%3Cpath%20d='M0%2011h18'%20stroke='%234f9bd0'%20stroke-opacity='.32'%20stroke-width='1.4'/%3E%3Cpath%20d='M18%2011h28'%20stroke='%238a76d0'%20stroke-opacity='.32'%20stroke-width='1.4'/%3E%3Cpath%20d='M46%2011h18'%20stroke='%23e0a33a'%20stroke-opacity='.32'%20stroke-width='1.4'/%3E%3Cpath%20d='M32%204l2%205%205%202-5%202-2%205-2-5-5-2%205-2z'%20fill='%23e0a33a'%20fill-opacity='.55'/%3E%3Cpath%20d='M16%207l1%202.5%202.5%201-2.5%201-1%202.5-1-2.5-2.5-1%202.5-1z'%20fill='%238a76d0'%20fill-opacity='.45'/%3E%3Ccircle%20cx='48'%20cy='11'%20r='1.7'%20fill='%234f9bd0'%20fill-opacity='.5'/%3E%3Ccircle%20cx='8'%20cy='15'%20r='1.3'%20fill='%23e07a68'%20fill-opacity='.5'/%3E%3C/svg%3E");
    }
    html[data-theme="dark"][data-background-motif="stars"] {
      --dashboard-motif-image:
        url("data:image/svg+xml,%3Csvg%20xmlns='http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg'%20width='52'%20height='52'%20viewBox='0%200%2052%2052'%3E%3Cpath%20d='M12%205l2%205l5%202l-5%202l-2%205l-2-5l-5-2l5-2z'%20fill='%23e2ad59'%20fill-opacity='.18'/%3E%3Cpath%20d='M38%2021l1.4%203.6l3.6%201.4l-3.6%201.4l-1.4%203.6l-1.4-3.6l-3.6-1.4l3.6-1.4z'%20fill='%23b9a7ec'%20fill-opacity='.16'/%3E%3Ccircle%20cx='44'%20cy='8'%20r='1.4'%20fill='%2380b9df'%20fill-opacity='.18'/%3E%3Ccircle%20cx='10'%20cy='38'%20r='1.2'%20fill='%23ea8777'%20fill-opacity='.16'/%3E%3C/svg%3E");
    }
    /* Motif-aware heading flourishes use dedicated colorful vector
       friezes tinted by the selected motif. They disappear with the
       motif switch and are hidden in print. */
    html:not([data-background-motif="none"]) .header-main::after,
    html:not([data-background-motif="none"]) .panel > h2::after,
    html:not([data-background-motif="none"]) .library-heading h2::after,
    html:not([data-background-motif="none"]) .trend-heading h2::after,
    html:not([data-background-motif="none"]) .heatmap-heading h2::after,
    html:not([data-background-motif="none"]) .achievement-panel > summary::after,
    html:not([data-background-motif="none"]) .period-panel > summary::after,
    html:not([data-background-motif="none"]) .annual-report-panel > summary::after,
    html:not([data-background-motif="none"]) .period-report-heading h3::after,
    html:not([data-background-motif="none"]) .annual-report-heading h3::after {
      content: "";
      display: block;
      width: 100%;
      height: 22px;
      margin-top: 8px;
      background-image: var(--dashboard-flourish-image, none);
      background-repeat: repeat-x;
      background-position: left center;
      background-size: 64px 22px;
      opacity: .92;
      pointer-events: none;
    }
    html:not([data-background-motif="none"]) .header-main::after {
      height: 26px;
      margin-top: 10px;
      background-size: 72px 26px;
      opacity: .96;
    }
    html[data-theme="light"][data-theme-preset="ocean"] {
      --bg: #eef5f7;
      --surface: #ffffff;
      --surface-soft: #f8fbfc;
      --surface-muted: #e8f0f3;
      --surface-track: #d8e6eb;
      --surface-chart: #edf4f6;
      --ink: #20313a;
      --muted: #61747e;
      --line: #d3e0e5;
      --green: #287b76;
      --blue: #2e6da4;
      --amber: #b8741b;
      --coral: #c45546;
      --violet: #655fa6;
      --hover: #f4fafb;
      --badge-bg: #e8f5f2;
      --badge-border: #aed5cb;
      --tag-bg: #eaf1f3;
      --tag-ink: #4c5b61;
      --language-bg: #e7f1fa;
      --estimate-bg: #fff6e5;
      --low-bg: #fff0eb;
      --deviation-bg: #f1effa;
      --placeholder-bg: #e4ecef;
      --placeholder-ink: #5d7078;
      --placeholder-mark: #f7fafb;
      --ring-empty: #e2ecef;
      --timeline-ring: #c5e2dc;
    }
    html[data-theme="dark"][data-theme-preset="ocean"] {
      --bg: #0d171b;
      --surface: #152329;
      --surface-soft: #1a2a31;
      --surface-muted: #21363f;
      --surface-track: #2a4650;
      --surface-chart: #1e343c;
      --ink: #e9f5f7;
      --muted: #a6bcc3;
      --line: #31505a;
      --green: #6bc8bd;
      --blue: #82bae3;
      --amber: #e0ad5a;
      --coral: #e98478;
      --violet: #b3a4e7;
      --hover: #1b2c33;
      --badge-bg: #183a36;
      --badge-border: #32665d;
      --tag-bg: #22353c;
      --tag-ink: #c8dbe0;
      --language-bg: #1c3948;
      --estimate-bg: #3d3320;
      --low-bg: #432720;
      --deviation-bg: #302b46;
      --placeholder-bg: #21363f;
      --placeholder-ink: #a8c0c7;
      --placeholder-mark: #18282e;
      --ring-empty: #2a4650;
      --timeline-ring: #2c625c;
    }
    html[data-theme="light"][data-theme-preset="sunset"] {
      --bg: #f8f4f0;
      --surface: #fffdfb;
      --surface-soft: #fcfaf8;
      --surface-muted: #f1ebe5;
      --surface-track: #e8ded5;
      --surface-chart: #f6f1ec;
      --ink: #352b27;
      --muted: #74665f;
      --line: #e4d8cf;
      --green: #34775e;
      --blue: #2f6f99;
      --amber: #c2721a;
      --coral: #c34f3d;
      --violet: #7659a2;
      --hover: #fdf9f6;
      --badge-bg: #eaf6ef;
      --badge-border: #b8d9c7;
      --tag-bg: #f1ebe6;
      --tag-ink: #655952;
      --language-bg: #eaf3fa;
      --estimate-bg: #fff5e4;
      --low-bg: #fff0eb;
      --deviation-bg: #f1edfa;
      --placeholder-bg: #eee6df;
      --placeholder-ink: #70635c;
      --placeholder-mark: #fffaf6;
      --ring-empty: #ece3dc;
      --timeline-ring: #d2e9dd;
    }
    html[data-theme="dark"][data-theme-preset="sunset"] {
      --bg: #191310;
      --surface: #251c18;
      --surface-soft: #2b211d;
      --surface-muted: #382b25;
      --surface-track: #4a3931;
      --surface-chart: #322620;
      --ink: #f7eee8;
      --muted: #c5b1a6;
      --line: #513e35;
      --green: #72c7a3;
      --blue: #86b8df;
      --amber: #e3ad58;
      --coral: #ed897a;
      --violet: #baa8eb;
      --hover: #30241f;
      --badge-bg: #1f3a2f;
      --badge-border: #3c6953;
      --tag-bg: #352a25;
      --tag-ink: #e0cec4;
      --language-bg: #213441;
      --estimate-bg: #40311d;
      --low-bg: #45261f;
      --deviation-bg: #342b45;
      --placeholder-bg: #382b25;
      --placeholder-ink: #c7b4a9;
      --placeholder-mark: #281f1b;
      --ring-empty: #4a3931;
      --timeline-ring: #35624f;
    }
    html[data-theme="light"][data-theme-preset="graphite"] {
      --bg: #f2f4f6;
      --surface: #ffffff;
      --surface-soft: #fafbfc;
      --surface-muted: #eaedf0;
      --surface-track: #dde2e6;
      --surface-chart: #f0f2f4;
      --ink: #252c31;
      --muted: #667178;
      --line: #d8dee2;
      --green: #28775f;
      --blue: #356f9f;
      --amber: #b97422;
      --coral: #bd5748;
      --violet: #675d9f;
      --hover: #f6f8f9;
      --badge-bg: #e9f5ef;
      --badge-border: #b6d7c8;
      --tag-bg: #edf0f2;
      --tag-ink: #536067;
      --language-bg: #e9f1f8;
      --estimate-bg: #fff5e5;
      --low-bg: #fff0ed;
      --deviation-bg: #f0eff8;
      --placeholder-bg: #e5e9ec;
      --placeholder-ink: #66737a;
      --placeholder-mark: #fafbfc;
      --ring-empty: #e3e7ea;
      --timeline-ring: #cde4d9;
    }
    html[data-theme="dark"][data-theme-preset="graphite"] {
      --bg: #111519;
      --surface: #1b2025;
      --surface-soft: #20262b;
      --surface-muted: #293138;
      --surface-track: #354047;
      --surface-chart: #252d33;
      --ink: #eef2f5;
      --muted: #aab3ba;
      --line: #3d4951;
      --green: #6fc5a0;
      --blue: #82b6df;
      --amber: #e0aa57;
      --coral: #e78477;
      --violet: #b2a4e5;
      --hover: #232a30;
      --badge-bg: #1c392e;
      --badge-border: #35634e;
      --tag-bg: #2a3238;
      --tag-ink: #cbd4d9;
      --language-bg: #203542;
      --estimate-bg: #3d321f;
      --low-bg: #422720;
      --deviation-bg: #312b43;
      --placeholder-bg: #293138;
      --placeholder-ink: #adb8bf;
      --placeholder-mark: #1d2429;
      --ring-empty: #354047;
      --timeline-ring: #315c4a;
    }
    html[data-theme="light"][data-theme-preset="rose"] {
      --bg: #f9f0f2;
      --surface: #ffffff;
      --surface-soft: #fcf7f8;
      --surface-muted: #f2e7ea;
      --surface-track: #e5d6da;
      --surface-chart: #f6edef;
      --ink: #392b30;
      --muted: #7b6670;
      --line: #ead8dd;
      --hover: #fdf8f9;
      --tag-bg: #f1e8eb;
      --tag-ink: #66555c;
      --placeholder-bg: #eadde1;
      --placeholder-ink: #75636a;
      --placeholder-mark: #fdf9fa;
      --ring-empty: #e9dde1;
      --timeline-ring: #d3e9de;
    }
    html[data-theme="dark"][data-theme-preset="rose"] {
      --bg: #1b1216;
      --surface: #261a1f;
      --surface-soft: #2d1f25;
      --surface-muted: #392830;
      --surface-track: #4a3540;
      --surface-chart: #32222a;
      --ink: #f5ecef;
      --muted: #c5b1b9;
      --line: #513b45;
      --hover: #302128;
      --tag-bg: #382831;
      --tag-ink: #dccbd2;
      --placeholder-bg: #392830;
      --placeholder-ink: #c7b5bc;
      --placeholder-mark: #281c21;
      --ring-empty: #4a3540;
      --timeline-ring: #33604d;
    }
    html[data-theme="light"][data-theme-preset="lavender"] {
      --bg: #f4f1fa;
      --surface: #ffffff;
      --surface-soft: #faf8fd;
      --surface-muted: #ede9f5;
      --surface-track: #ded8eb;
      --surface-chart: #f1edf8;
      --ink: #2f293d;
      --muted: #706781;
      --line: #e0daea;
      --hover: #fbf9fe;
      --tag-bg: #ece8f3;
      --tag-ink: #61586f;
      --placeholder-bg: #e3dff0;
      --placeholder-ink: #6d6480;
      --placeholder-mark: #fbfaff;
      --ring-empty: #e3dff0;
      --timeline-ring: #d4e9df;
    }
    html[data-theme="dark"][data-theme-preset="lavender"] {
      --bg: #15121d;
      --surface: #201b2a;
      --surface-soft: #272231;
      --surface-muted: #332c41;
      --surface-track: #453b55;
      --surface-chart: #2d2739;
      --ink: #f1ecf8;
      --muted: #b9afc9;
      --line: #4a4059;
      --hover: #292333;
      --tag-bg: #352f43;
      --tag-ink: #d8d0e5;
      --placeholder-bg: #332c41;
      --placeholder-ink: #c2b8d0;
      --placeholder-mark: #211c2b;
      --ring-empty: #453b55;
      --timeline-ring: #37604e;
    }
    html[data-theme="light"][data-theme-preset="sky"] {
      --bg: #eef6fb;
      --surface: #ffffff;
      --surface-soft: #f8fcfe;
      --surface-muted: #e3eff6;
      --surface-track: #d3e2eb;
      --surface-chart: #eaf4f9;
      --ink: #24333d;
      --muted: #627680;
      --line: #d1e0e7;
      --hover: #f5fbfe;
      --tag-bg: #e7f0f5;
      --tag-ink: #586a73;
      --placeholder-bg: #dce9ef;
      --placeholder-ink: #60747e;
      --placeholder-mark: #f9fcfe;
      --ring-empty: #dce9ef;
      --timeline-ring: #cbe6dd;
    }
    html[data-theme="dark"][data-theme-preset="sky"] {
      --bg: #0d171d;
      --surface: #15232b;
      --surface-soft: #1a2a32;
      --surface-muted: #22343d;
      --surface-track: #2f4753;
      --surface-chart: #1d3039;
      --ink: #e9f4f8;
      --muted: #a8bcc4;
      --line: #31505c;
      --hover: #1b2d35;
      --tag-bg: #243840;
      --tag-ink: #c8dbe0;
      --placeholder-bg: #22343d;
      --placeholder-ink: #abc0c8;
      --placeholder-mark: #17262d;
      --ring-empty: #2f4753;
      --timeline-ring: #32624f;
    }
    html[data-theme="light"][data-theme-preset="sepia"] {
      --bg: #f6f1ea;
      --surface: #fffdfb;
      --surface-soft: #fbf8f4;
      --surface-muted: #eee8df;
      --surface-track: #e0d8cc;
      --surface-chart: #f3ede5;
      --ink: #352f29;
      --muted: #746b61;
      --line: #e3dacf;
      --hover: #fcf9f5;
      --tag-bg: #efe9e1;
      --tag-ink: #665e55;
      --placeholder-bg: #e7dfd4;
      --placeholder-ink: #746a5f;
      --placeholder-mark: #fffdfa;
      --ring-empty: #e7dfd4;
      --timeline-ring: #d3e7dc;
    }
    html[data-theme="dark"][data-theme-preset="sepia"] {
      --bg: #1b1712;
      --surface: #252019;
      --surface-soft: #2b251d;
      --surface-muted: #373026;
      --surface-track: #4a4032;
      --surface-chart: #312a21;
      --ink: #f4eee6;
      --muted: #c4b8a8;
      --line: #514638;
      --hover: #302920;
      --tag-bg: #3a3228;
      --tag-ink: #ded3c5;
      --placeholder-bg: #373026;
      --placeholder-ink: #c8baaa;
      --placeholder-mark: #241f18;
      --ring-empty: #4a4032;
      --timeline-ring: #34604c;
    }
    * { box-sizing: border-box; }
    [hidden] { display: none !important; }
    body {
      margin: 0;
      background: var(--bg);
      background-image:
        var(--dashboard-motif-image, none),
        var(--dashboard-texture-image, none);
      background-size:
        var(--dashboard-motif-size, auto),
        var(--dashboard-texture-size, auto);
      background-repeat:
        var(--dashboard-motif-repeat, repeat),
        repeat;
      background-attachment: fixed;
      color: var(--ink);
      overflow-x: hidden;
      font: 14px/1.5 -apple-system, BlinkMacSystemFont, "Segoe UI",
        "Microsoft YaHei", sans-serif;
    }
    button:focus-visible,
    select:focus-visible,
    input:focus-visible,
    summary:focus-visible,
    a:focus-visible {
      outline: 2px solid var(--focus-ring);
      outline-offset: 2px;
    }
    body::before,
    body::after {
      content: "";
      position: fixed;
      inset: 0;
      pointer-events: none;
      z-index: 0;
    }
    body::before {
      inset: -24px;
      background-image: var(--dashboard-background-image, none);
      background-position: center;
      background-repeat: no-repeat;
      background-size: cover;
      filter: blur(var(--dashboard-background-blur, 0px));
      opacity: var(--dashboard-background-opacity, 0);
    }
    body::after {
      background: var(--bg);
      opacity: var(--dashboard-background-overlay, 0);
    }
    .page {
      width: min(1440px, calc(100% - 32px));
      margin: 24px auto 48px;
      position: relative;
      z-index: 1;
    }
    header {
      display: flex;
      justify-content: space-between;
      gap: 14px 24px;
      align-items: flex-start;
      flex-wrap: wrap;
      margin-bottom: 14px;
    }
    .header-main { flex: 1 1 460px; min-width: 0; }
    .title-line {
      display: flex;
      align-items: baseline;
      flex-wrap: wrap;
      gap: 6px 12px;
      min-width: 0;
    }
    h1 { margin: 0; font-size: 30px; line-height: 1.2; }
    .meta { margin-top: 6px; color: var(--muted); overflow-wrap: anywhere; }
    .badge {
      border: 1px solid var(--badge-border);
      background: var(--badge-bg);
      color: var(--green);
      border-radius: 999px;
      padding: 5px 10px;
      font-size: 12px;
      font-weight: 700;
      white-space: nowrap;
    }
    .header-actions {
      position: relative;
      display: flex;
      align-items: center;
      justify-content: flex-end;
      flex-wrap: wrap;
      gap: 8px;
      flex: 0 1 auto;
    }
    .control-button,
    .column-picker summary,
    .export-button {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 34px;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 7px 10px;
      background: var(--surface);
      color: var(--ink);
      font: inherit;
      font-size: 12px;
      font-weight: 700;
      text-decoration: none;
      cursor: pointer;
      white-space: nowrap;
    }
    .control-button:hover,
    .column-picker summary:hover,
    .export-button:hover,
    button:hover {
      background: var(--hover);
    }
    .theme-picker {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      min-height: 34px;
      padding: 5px 8px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
    }
    .theme-picker select {
      min-width: 82px;
      padding: 0;
      border: 0;
      background: transparent;
      color: var(--ink);
      font-size: 12px;
      font-weight: 700;
    }
    .greeting-row {
      display: flex;
      align-items: baseline;
      flex-wrap: wrap;
      gap: 6px 12px;
      margin-top: 0;
      min-width: 0;
    }
    .greeting-message {
      margin: 0;
      font-family: var(
        --greeting-font,
        "SimSun", "Songti SC", "Noto Serif CJK SC", serif
      );
      font-size: 16px;
      font-weight: 700;
      line-height: 1.4;
      overflow-wrap: anywhere;
      color: var(--greeting-ink, var(--ink));
      transition: color .7s ease;
    }
    .greeting-emoticon {
      display: inline-block;
      white-space: nowrap;
    }
    .greeting-picker {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      min-height: 34px;
      padding: 5px 8px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
      min-width: 0;
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
    }
    .greeting-picker select {
      min-width: 92px;
      padding: 0;
      border: 0;
      background: transparent;
      color: var(--ink);
      font-size: 12px;
      font-weight: 700;
    }
    .appearance-panel {
      margin: 0;
      flex: 0 1 auto;
    }
    .appearance-panel > summary {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      min-height: 34px;
      padding: 7px 10px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
      color: var(--ink);
      font-size: 12px;
      font-weight: 700;
      cursor: pointer;
      list-style: none;
      white-space: nowrap;
    }
    .appearance-panel > summary:hover { background: var(--hover); }
    .appearance-panel > summary::-webkit-details-marker { display: none; }
    .appearance-icon { width: 14px; height: 14px; flex: 0 0 auto; }
    .appearance-panel > summary em {
      display: none;
    }
    .appearance-panel:not([open]) .appearance-body { display: none; }
    .appearance-body {
      position: absolute;
      z-index: 30;
      top: calc(100% + 6px);
      right: 0;
      width: min(560px, calc(100vw - 32px));
      max-height: calc(100vh - 90px);
      overflow-y: auto;
      padding: 14px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
      box-shadow: var(--shadow-raised);
    }
    .appearance-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(210px, 1fr));
      gap: 12px;
    }
    .appearance-section-toggles {
      margin-top: 14px;
      padding-top: 12px;
      border-top: 1px solid var(--line);
    }
    .appearance-subtitle {
      margin: 0 0 10px;
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
    }
    .appearance-field {
      display: grid;
      align-content: start;
      gap: 6px;
      min-width: 0;
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
    }
    .appearance-field > * {
      min-width: 0;
      max-width: 100%;
    }
    .appearance-field select,
    .appearance-field input[type="file"] {
      width: 100%;
      max-width: 100%;
      min-width: 0;
      box-sizing: border-box;
    }
    .appearance-field input[type="file"] {
      color: var(--muted);
      font: inherit;
      font-size: 12px;
    }
    .appearance-field input[type="text"] {
      width: 100%;
      max-width: 100%;
      min-width: 0;
      padding: 8px 10px;
      border: 1px solid var(--line);
      border-radius: 7px;
      background: var(--surface);
      color: var(--ink);
      font: inherit;
      font-size: 13px;
    }
    .appearance-name-row {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 7px;
      align-items: center;
      min-width: 0;
      width: 100%;
    }
    .appearance-name-row button {
      min-height: 34px;
      min-width: 0;
      padding: 6px 9px;
      font-size: 11px;
    }
    .local-name-dialog {
      position: fixed;
      z-index: 60;
      top: 18px;
      left: 50%;
      width: min(420px, calc(100vw - 28px));
      margin: 0;
      padding: 0;
      transform: translateX(-50%);
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
      color: var(--ink);
      box-shadow: var(--shadow-raised);
    }
    .local-name-dialog::backdrop {
      background: rgba(12, 18, 15, .28);
    }
    .local-name-body {
      display: grid;
      gap: 10px;
      padding: 15px;
    }
    .local-name-body h2 {
      margin: 0;
      font-size: 17px;
    }
    .local-name-body p {
      margin: 0;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.5;
    }
    .local-name-input {
      width: 100%;
      min-width: 0;
      padding: 9px 10px;
      border: 1px solid var(--line);
      border-radius: 7px;
      background: var(--surface-soft);
      color: var(--ink);
      font: inherit;
    }
    .local-name-actions {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      justify-content: flex-end;
    }
    .appearance-toggle-row {
      display: flex;
      align-items: center;
      gap: 8px;
      min-width: 0;
      color: var(--ink);
    }
    .appearance-toggle-row input[type="checkbox"] {
      flex: 0 0 auto;
      width: 16px;
      height: 16px;
      margin: 0;
      accent-color: var(--green);
    }
    .appearance-toggle-row span {
      min-width: 0;
      overflow-wrap: anywhere;
    }
    .range-row {
      display: grid;
      grid-template-columns: minmax(0, 1fr) 48px;
      gap: 8px;
      align-items: center;
    }
    .range-row input[type="range"] {
      width: 100%;
      min-width: 0;
      accent-color: var(--green);
    }
    .range-row output {
      color: var(--ink);
      text-align: right;
      font-variant-numeric: tabular-nums;
    }
    .appearance-help,
    .appearance-status {
      margin: 10px 0 0;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.5;
    }
    .appearance-status {
      min-height: 18px;
    }
    .appearance-actions {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 12px;
    }
    .appearance-actions button {
      min-height: 34px;
    }
    .summary {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
      gap: 10px;
      margin-bottom: 14px;
    }
    .metric {
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 14px 15px;
      min-height: 88px;
    }
    .metric span { display: block; color: var(--muted); font-size: 12px; }
    .metric strong {
      display: block;
      margin-top: 8px;
      font-size: 22px;
      line-height: 1.1;
      overflow-wrap: anywhere;
    }
    .metric .metric-source {
      display: block;
      margin-top: 5px;
      color: var(--muted);
      font-size: 10px;
      line-height: 1.35;
    }
    .library-panel {
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 18px;
      margin-bottom: 14px;
    }
    .library-heading {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      align-items: center;
      gap: 16px;
      margin-bottom: 10px;
      min-width: 0;
    }
    .library-heading > div { min-width: 0; }
    .library-heading h2 {
      margin: 0;
      font-size: 17px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .library-heading p {
      margin: 4px 0 0;
      color: var(--muted);
      font-size: 12px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .shelf-title-cluster {
      display: flex;
      align-items: center;
      gap: 12px;
      min-width: 0;
      flex: 1 1 auto;
    }
    .shelf-title-copy {
      min-width: 0;
      flex: 0 1 auto;
    }
    .library-actions {
      display: flex;
      align-items: center;
      justify-content: flex-end;
      gap: 12px;
      flex-wrap: nowrap;
      min-width: 0;
    }
    .shelf-mode-switch {
      display: inline-flex;
      flex: 0 0 auto;
      align-items: center;
      padding: 3px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface-muted);
    }
    .shelf-mode-switch button {
      border: 0;
      border-radius: 6px;
      padding: 7px 10px;
      background: transparent;
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
      white-space: nowrap;
      box-shadow: none;
    }
    .shelf-mode-switch button {
      display: inline-flex;
      align-items: center;
      gap: 5px;
    }
    .shelf-mode-icon,
    .shelf-mode-label-short {
      display: none;
    }
    .shelf-mode-switch button[aria-pressed="true"] {
      background: var(--surface);
      color: var(--ink);
      box-shadow: var(--shadow-card);
    }
    .shelf-mode-switch button svg {
      width: 14px;
      height: 14px;
      flex: 0 0 auto;
      fill: none;
      stroke: currentColor;
      stroke-linecap: round;
      stroke-linejoin: round;
      stroke-width: 1.8;
    }
    html[data-shelf-title-style="b"] .shelf-mode-icon {
      display: block;
    }
    html[data-shelf-title-style="b"] .shelf-mode-label-full {
      display: none;
    }
    html[data-shelf-title-style="b"] .shelf-mode-label-short {
      display: inline;
    }
    html[data-shelf-title-style="b"] .shelf-mode-switch button {
      padding-inline: 8px;
    }
    html[data-shelf-title-style="c"] .shelf-mode-switch {
      align-self: stretch;
      padding: 0;
      border: 0;
      border-bottom: 1px solid var(--line);
      border-radius: 0;
      background: transparent;
    }
    html[data-shelf-title-style="c"] .shelf-mode-switch button {
      min-height: 34px;
      padding: 7px 6px;
      border: 0;
      border-bottom: 2px solid transparent;
      border-radius: 0;
      background: transparent;
      box-shadow: none;
    }
    html[data-shelf-title-style="c"] .shelf-mode-switch button[aria-pressed="true"] {
      border-bottom-color: var(--green);
      background: transparent;
      color: var(--green);
      box-shadow: none;
    }
    .shelf-mode-switch button:focus-visible,
    .shelf-pager button:focus-visible {
      outline: 3px solid var(--focus-ring);
      outline-offset: 2px;
    }
    html[data-shelf-position="hidden"] .shelf-mode-switch {
      display: none;
    }
    .shelf-fallback-wrap {
      position: relative;
      display: none;
      flex: 0 0 auto;
    }
    html[data-shelf-position="hidden"] .shelf-fallback-wrap {
      display: inline-flex;
    }
    .shelf-fallback-button {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      min-height: 32px;
      padding: 5px 9px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
    }
    .shelf-fallback-button svg {
      width: 15px;
      height: 15px;
      fill: none;
      stroke: currentColor;
      stroke-linecap: round;
      stroke-linejoin: round;
      stroke-width: 1.8;
    }
    .shelf-fallback-menu {
      position: absolute;
      z-index: 35;
      top: calc(100% + 7px);
      right: 0;
      display: flex;
      align-items: center;
      gap: 4px;
      padding: 4px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
      box-shadow: var(--shadow-raised);
    }
    .shelf-fallback-menu[hidden] { display: none; }
    .shelf-fallback-menu button {
      min-height: 30px;
      padding: 5px 9px;
      border: 1px solid transparent;
      border-radius: 6px;
      background: transparent;
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
      white-space: nowrap;
    }
    .shelf-fallback-menu button[aria-pressed="true"] {
      border-color: var(--line);
      background: var(--surface-muted);
      color: var(--ink);
    }
    .shelf-pager {
      display: inline-flex;
      align-items: center;
      gap: 7px;
      color: var(--muted);
      font-size: 12px;
      font-variant-numeric: tabular-nums;
    }
    .shelf-pager button {
      width: 32px;
      height: 32px;
      padding: 0;
      border-radius: 50%;
      font-size: 20px;
      line-height: 1;
    }
    .shelf-pager button:disabled {
      opacity: .42;
      cursor: not-allowed;
    }
    .cover-count-badge {
      flex: 0 0 auto;
      border-radius: 999px;
      padding: 4px 9px;
      background: var(--badge-bg);
      color: var(--green);
      font-size: 12px;
      font-weight: 700;
      white-space: nowrap;
    }
    .cover-count-badge strong {
      font: inherit;
      color: inherit;
    }
    .cover-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(340px, 1fr));
      gap: 14px;
    }
    .book-cover-card {
      display: grid;
      grid-template-columns: 124px minmax(0, 1fr);
      gap: 16px;
      align-items: start;
      min-width: 0;
      min-height: 200px;
      padding: 14px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface-soft);
    }
    .cover-frame {
      width: 124px;
      aspect-ratio: 2 / 3;
      align-self: start;
      border-radius: 5px;
      overflow: hidden;
      background: var(--placeholder-bg);
      border: 1px solid var(--line);
    }
    .cover-frame img {
      display: block;
      width: 100%;
      height: 100%;
      object-fit: cover;
    }
    .cover-placeholder {
      width: 100%;
      height: 100%;
      display: grid;
      align-content: center;
      justify-items: center;
      gap: 5px;
      padding: 10px;
      color: var(--placeholder-ink);
      text-align: center;
      background-image:
        linear-gradient(var(--line) 1px, transparent 1px),
        linear-gradient(90deg, var(--line) 1px, transparent 1px);
      background-size: 18px 18px;
    }
    .cover-placeholder strong {
      display: grid;
      place-items: center;
      width: 44px;
      height: 44px;
      border-radius: 50%;
      background: var(--placeholder-mark);
      color: var(--green);
      font-size: 22px;
    }
    .cover-placeholder small { font-size: 11px; }
    .cover-info {
      min-width: 0;
      display: flex;
      flex-direction: column;
    }
    .cover-rank {
      color: var(--muted);
      font-size: 11px;
      font-weight: 700;
    }
    .cover-info h3 {
      margin: 4px 0 3px;
      font-size: 15px;
      line-height: 1.35;
      display: -webkit-box;
      -webkit-line-clamp: 2;
      -webkit-box-orient: vertical;
      overflow: hidden;
      overflow-wrap: anywhere;
    }
    .cover-info > p {
      margin: 0;
      color: var(--muted);
      font-size: 12px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .cover-stats {
      display: grid;
      grid-template-columns: 1fr;
      gap: 6px;
      margin-top: 10px;
    }
    .cover-stat {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      align-items: baseline;
      gap: 10px;
      min-width: 0;
      padding: 6px 8px;
      border-radius: 6px;
      background: var(--surface-muted);
    }
    .cover-stat-label {
      display: flex;
      align-items: baseline;
      gap: 6px;
      min-width: 0;
      color: var(--muted);
      font-size: 11px;
    }
    .cover-stat-label .i18n-zh,
    .cover-stat-label .i18n-en {
      white-space: nowrap;
    }
    .cover-stat-source {
      flex: 0 0 auto;
      padding: 1px 6px;
      border-radius: 999px;
      background: var(--tag-bg);
      color: var(--tag-ink);
      font-size: 10px;
      line-height: 1.5;
      white-space: nowrap;
      cursor: help;
    }
    .cover-stat strong {
      font-size: 13px;
      text-align: right;
      font-variant-numeric: tabular-nums;
      white-space: nowrap;
      overflow: visible;
      text-overflow: clip;
    }
    .cover-stat .cover-effective {
      grid-column: 1 / -1;
      color: var(--gold-ink);
      font-size: 10px;
      text-align: right;
      overflow-wrap: anywhere;
    }
    .cover-progress {
      display: grid;
      grid-template-columns: 45px minmax(0, 1fr);
      gap: 7px;
      align-items: center;
      margin-top: 9px;
      color: var(--muted);
      font-size: 11px;
    }
    .cover-progress > div {
      height: 6px;
      border-radius: 999px;
      overflow: hidden;
      background: var(--surface-track);
    }
    .cover-progress i {
      display: block;
      height: 100%;
      border-radius: inherit;
    }
    .insights {
      display: grid;
      grid-template-columns: minmax(0, 1.55fr) repeat(2, minmax(0, 1fr));
      gap: 10px;
      margin-bottom: 14px;
    }
    .insight {
      --insight-accent: var(--blue);
      display: grid;
      grid-template-columns: minmax(0, 1fr);
      align-items: start;
      gap: 10px;
      min-width: 0;
      padding: 13px 15px;
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow-card);
    }
    .insight[data-insight-layout="feature"],
    .insight[data-insight-layout="cover"] {
      grid-template-columns: auto minmax(0, 1fr);
    }
    .insight[data-insight-key="most-read"] { --insight-accent: var(--blue); }
    .insight[data-insight-key="longest-book"] { --insight-accent: var(--green); }
    .insight[data-insight-key="fastest-wpm"] { --insight-accent: var(--amber); }
    .insight-cover {
      position: relative;
      display: block;
      width: 54px;
      aspect-ratio: 2 / 3;
      overflow: hidden;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: var(--placeholder-bg);
    }
    .insight[data-insight-layout="feature"] .insight-cover {
      width: 72px;
    }
    .insight-cover img {
      display: block;
      width: 100%;
      height: 100%;
      object-fit: cover;
    }
    .insight-cover-placeholder {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: 3px;
      width: 100%;
      height: 100%;
      padding: 4px;
      text-align: center;
      color: var(--placeholder-ink);
    }
    .insight-cover-placeholder strong {
      font-size: 18px;
      line-height: 1;
    }
    .insight-cover-placeholder small {
      font-size: 9px;
      line-height: 1.2;
    }
    .insight-body {
      display: grid;
      align-content: start;
      gap: 6px;
      min-width: 0;
    }
    .insight-head {
      display: flex;
      align-items: center;
      gap: 7px;
      min-width: 0;
    }
    .insight-icon {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      flex: 0 0 auto;
      width: 24px;
      height: 24px;
      border-radius: 7px;
      background: var(--tag-bg);
      color: var(--insight-accent);
    }
    .insight-icon svg { width: 15px; height: 15px; }
    .insight-label {
      min-width: 0;
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
      overflow-wrap: anywhere;
    }
    .insight > strong,
    .insight-value {
      display: block;
      min-width: 0;
      margin: 0;
      font-size: 14px;
      line-height: 1.35;
      overflow-wrap: anywhere;
      word-break: break-word;
    }
    .insight[data-insight-layout="feature"] .insight-value {
      font-size: 16px;
    }
    .insight > span {
      color: var(--muted);
      font-size: 12px;
    }
    .insight-detail {
      display: block;
      min-width: 0;
      color: var(--insight-accent);
      font-size: 20px;
      font-weight: 700;
      line-height: 1.15;
      overflow-wrap: anywhere;
      font-variant-numeric: tabular-nums;
    }
    .insight[data-insight-layout="feature"] .insight-detail {
      font-size: 26px;
    }
    .insight[data-insight-layout="plain"] .insight-detail {
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
      line-height: 1.4;
    }
    .insight-share {
      display: grid;
      gap: 4px;
      margin-top: 2px;
      min-width: 0;
    }
    .insight-share-track {
      display: block;
      height: 6px;
      overflow: hidden;
      border-radius: 999px;
      background: var(--surface-track);
    }
    .insight-share-track i {
      display: block;
      height: 100%;
      border-radius: inherit;
      background: var(--insight-accent);
    }
    .insight-share-label {
      color: var(--muted);
      font-size: 11px;
      font-weight: 700;
      overflow-wrap: anywhere;
    }
    .achievement-panel {
      position: relative;
      margin: 0 0 14px;
      border: 1px solid var(--gold-deep);
      border-radius: 8px;
      background:
        linear-gradient(180deg, rgba(255, 212, 0, .08), transparent 64px),
        var(--surface-soft);
      box-shadow: 0 0 0 1px var(--gold-glow);
    }
    .achievement-panel > summary {
      display: block;
      padding: 12px 170px 12px 14px;
      cursor: pointer;
      font-weight: 700;
      list-style: none;
      color: var(--gold-ink);
    }
    .achievement-panel > summary::-webkit-details-marker {
      display: none;
    }
    .achievement-panel > summary::before {
      content: "";
      display: inline-block;
      width: 0;
      height: 0;
      margin-right: 9px;
      border-left: 6px solid currentColor;
      border-top: 5px solid transparent;
      border-bottom: 5px solid transparent;
      vertical-align: middle;
      transition: transform 0.15s ease;
    }
    .achievement-panel[open] > summary::before {
      transform: rotate(90deg);
    }
    .achievement-panel > summary em {
      display: block;
      margin-top: 5px;
      color: var(--muted);
      font-size: 12px;
      font-style: normal;
      font-weight: 400;
    }
    .achievement-body { padding: 0 14px 14px; }
    .achievement-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 10px;
    }
    .achievement-toggle {
      display: inline-flex;
      align-items: center;
      gap: 7px;
      margin: 0 0 10px;
      color: var(--muted);
      font-size: 12px;
      cursor: pointer;
    }
    .achievement-toggle input {
      width: 16px;
      height: 16px;
      accent-color: var(--gold-deep);
    }
    .achievement-card[data-achievement-key="manga_volumes"][hidden] {
      display: none;
    }
    .star-sky-entry {
      position: absolute;
      top: 10px;
      right: 14px;
      z-index: 2;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      min-height: 32px;
      padding: 5px 10px;
      border: 1px solid var(--gold-deep);
      border-radius: 999px;
      background: var(--surface);
      color: var(--gold-ink);
      font: inherit;
      font-size: 12px;
      font-weight: 700;
      cursor: pointer;
    }
    .star-sky-entry:hover {
      background: var(--hover);
    }
    .star-sky-entry:focus-visible {
      outline: 3px solid var(--focus-ring);
      outline-offset: 2px;
    }
    .star-sky-entry svg {
      width: 17px;
      height: 17px;
      fill: var(--gold);
      stroke: var(--gold-deep);
      stroke-width: 1.35;
      stroke-linejoin: round;
      stroke-linecap: round;
    }
    .achievement-card {
      display: grid;
      grid-template-columns: 46px minmax(0, 1fr);
      gap: 11px;
      min-width: 0;
      padding: 13px;
      border: 1px solid var(--line);
      border-left: 4px solid var(--gold-deep);
      border-radius: 8px;
      background: var(--surface);
    }
    .achievement-unlocked {
      border-left-color: var(--gold);
      box-shadow: 0 0 0 1px var(--gold-glow);
    }
    .achievement-in_progress { border-left-color: var(--gold-deep); }
    .achievement-insufficient { border-left-color: var(--muted); }
    .achievement-badge {
      display: grid;
      place-items: center;
      width: 46px;
      height: 46px;
      border: 1px solid var(--gold-deep);
      border-radius: 50%;
      background: rgba(255, 212, 0, .14);
      color: var(--gold-ink);
      font-size: 12px;
      font-weight: 800;
      letter-spacing: 0;
      box-shadow:
        inset 0 0 0 4px var(--surface),
        inset 0 0 0 5px var(--badge-border);
      transition: transform .15s ease, box-shadow .15s ease;
    }
    .achievement-badge span {
      font-variant-numeric: tabular-nums;
    }
    html[data-achievement-badge-style="ring"] .achievement-badge {
      border-width: 5px;
      border-color: var(--surface-track);
      background: var(--surface);
      box-shadow: inset 0 0 0 1px var(--badge-border);
    }
    html[data-achievement-badge-style="ring"] .achievement-unlocked .achievement-badge {
      border-color: var(--gold);
    }
    html[data-achievement-badge-style="ring"] .achievement-in_progress .achievement-badge {
      border-color: var(--gold-deep);
    }
    html[data-achievement-badge-style="minimal"] .achievement-card {
      grid-template-columns: 8px minmax(0, 1fr);
    }
    html[data-achievement-badge-style="minimal"] .achievement-badge {
      width: 8px;
      height: 44px;
      border: 0;
      border-radius: 999px;
      background: var(--surface-track);
      box-shadow: none;
    }
    html[data-achievement-badge-style="minimal"] .achievement-badge span {
      display: none;
    }
    html[data-achievement-badge-style="minimal"] .achievement-unlocked .achievement-badge {
      background: var(--gold);
    }
    html[data-achievement-badge-style="minimal"] .achievement-in_progress .achievement-badge {
      background: var(--gold-deep);
    }
    .achievement-content {
      min-width: 0;
    }
    .achievement-heading {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 5px 8px;
      min-width: 0;
    }
    .achievement-heading h3 {
      margin: 0;
      font-size: 14px;
      line-height: 1.3;
      overflow-wrap: anywhere;
    }
    .achievement-state {
      flex: 0 0 auto;
      padding: 2px 6px;
      border-radius: 999px;
      background: var(--tag-bg);
      color: var(--tag-ink);
      font-size: 10px;
      font-weight: 700;
      white-space: nowrap;
    }
    .achievement-value {
      display: block;
      margin-top: 8px;
      font-size: 21px;
      line-height: 1.15;
      overflow-wrap: anywhere;
      white-space: normal;
      font-variant-numeric: tabular-nums;
    }
    .achievement-condition {
      margin: 6px 0 0;
      color: var(--muted);
      font-size: 11px;
      line-height: 1.45;
      overflow-wrap: anywhere;
    }
    .achievement-stars {
      display: flex;
      flex-wrap: wrap;
      gap: 4px;
      margin-top: 8px;
      min-height: 22px;
      color: var(--gold-deep);
      line-height: 1;
      letter-spacing: 0;
    }
    .achievement-star {
      display: inline-grid;
      place-items: center;
      width: 22px;
      height: 22px;
      border-radius: 6px;
      outline: none;
      transition: transform .18s ease, opacity .18s ease, filter .18s ease;
    }
    .achievement-star:focus-visible {
      outline: 3px solid var(--gold-deep);
      outline-offset: 2px;
    }
    .achievement-star svg {
      width: 20px;
      height: 20px;
      overflow: visible;
    }
    .achievement-star svg path {
      fill: rgba(255, 212, 59, .08);
      stroke: #d4c59b;
      stroke-width: 1.8;
    }
    .achievement-star.star-lit svg path {
      fill: var(--gold);
      stroke: var(--gold-deep);
      filter: none;
    }
    .achievement-star.star-next svg path {
      fill: rgba(255, 212, 0, .18);
      stroke: var(--gold);
      filter: none;
    }
    .achievement-star.star-next {
      transform: scale(1.08);
      filter: drop-shadow(0 0 5px var(--gold-glow));
    }
    .achievement-star.star-next-pending svg path,
    .achievement-star.star-future svg path {
      fill: rgba(255, 212, 59, .06);
      stroke: #d4c59b;
      filter: none;
    }
    .achievement-star.star-future {
      opacity: .55;
    }
    html[data-theme="dark"] .achievement-star svg path,
    html[data-theme="dark"] .achievement-star.star-next-pending svg path,
    html[data-theme="dark"] .achievement-star.star-future svg path {
      stroke: #a99d75;
    }
    .achievement-progress {
      height: 7px;
      margin-top: 9px;
      border-radius: 999px;
      overflow: hidden;
      background: var(--surface-track);
    }
    .achievement-progress i {
      display: block;
      height: 100%;
      border-radius: inherit;
      background: linear-gradient(
        90deg,
        var(--gold-deep),
        var(--gold),
        var(--gold-bright)
      );
    }
    .achievement-unlocked .achievement-progress i {
      background: linear-gradient(
        90deg,
        var(--gold-deep),
        var(--gold),
        var(--gold-bright)
      );
    }
    .achievement-in_progress .achievement-progress i {
      background: linear-gradient(90deg, var(--gold-deep), var(--gold));
    }
    .achievement-footer {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 4px 8px;
      margin-top: 7px;
      color: var(--muted);
      font-size: 11px;
      line-height: 1.4;
    }
    .achievement-footer span,
    .achievement-footer small {
      min-width: 0;
      overflow-wrap: anywhere;
    }
    .achievement-footer small {
      max-width: 58%;
      text-align: right;
    }
    .achievement-encouragement {
      margin: 7px 0 0;
      color: var(--ink);
      font-size: 11px;
      font-weight: 650;
      line-height: 1.45;
      overflow-wrap: anywhere;
    }
    .achievement-criteria {
      margin: 8px 0 0;
      padding-top: 7px;
      border-top: 1px solid var(--line);
      color: var(--muted);
      font-size: 10px;
      line-height: 1.45;
      overflow-wrap: anywhere;
    }
    .star-sky-dialog {
      width: min(980px, calc(100vw - 28px));
      max-height: calc(100vh - 28px);
      padding: 0;
      border: 0;
      border-radius: 10px;
      color: var(--ink);
      background: transparent;
    }
    .star-sky-dialog::backdrop {
      background: rgba(11, 18, 27, .66);
      backdrop-filter: blur(3px);
    }
    .star-sky-shell {
      max-height: calc(100vh - 28px);
      overflow: auto;
      border: 1px solid rgba(255, 212, 59, .34);
      border-radius: 10px;
      background: linear-gradient(160deg, #10203a, #172b45 54%, #263b58);
      color: #f8fbff;
      box-shadow: var(--shadow-raised);
    }
    html[data-theme="light"] .star-sky-shell {
      background: linear-gradient(160deg, #173154, #254b72 56%, #3b6380);
    }
    .star-sky-header {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 14px;
      padding: 18px 20px 10px;
    }
    .star-sky-header h2 {
      margin: 0;
      color: #fff8dc;
      font-size: 27px;
      line-height: 1.2;
      letter-spacing: 0;
    }
    .star-sky-header p {
      margin: 7px 0 0;
      color: rgba(248, 251, 255, .78);
      font-size: 13px;
    }
    .star-sky-header strong {
      color: var(--gold);
      font-size: 16px;
    }
    .star-sky-close {
      flex: 0 0 auto;
      width: 34px;
      height: 34px;
      border: 1px solid rgba(255, 255, 255, .32);
      border-radius: 50%;
      background: rgba(255, 255, 255, .08);
      color: #fff;
      font-size: 22px;
      line-height: 1;
      cursor: pointer;
    }
    .star-sky-close:hover,
    .star-sky-close:focus-visible {
      background: rgba(255, 255, 255, .18);
      outline: 3px solid var(--gold);
      outline-offset: 2px;
    }
    .star-sky-field {
      position: relative;
      min-height: clamp(340px, 58vh, 560px);
      margin: 0 14px;
      overflow: hidden;
      border: 1px solid rgba(255, 255, 255, .12);
      border-radius: 8px;
      background-image:
        radial-gradient(circle, rgba(255, 255, 255, .52) 0 1px, transparent 1.7px),
        radial-gradient(circle, rgba(255, 222, 129, .34) 0 1px, transparent 1.8px),
        radial-gradient(circle, rgba(180, 210, 255, .28) 0 1px, transparent 1.8px);
      background-size: 74px 74px, 131px 131px, 197px 197px;
      background-position: 12px 18px, 47px 31px, 80px 8px;
    }
    .star-sky-field::after {
      content: "";
      position: absolute;
      inset: 0;
      pointer-events: none;
    }
    .star-sky-field::after {
      background: radial-gradient(ellipse at center, transparent 42%, rgba(4, 11, 22, .48));
    }
    .star-sky-star {
      position: absolute;
      left: var(--star-x);
      top: var(--star-y);
      z-index: 1;
      display: grid;
      place-items: center;
      width: calc(var(--star-size) + 18px);
      height: calc(var(--star-size) + 18px);
      padding: 0;
      border: 0;
      border-radius: 50%;
      background: transparent;
      transform: translate(-50%, -50%);
      cursor: pointer;
    }
    .star-sky-star svg {
      width: var(--star-size);
      height: var(--star-size);
      overflow: visible;
    }
    .star-sky-star svg path {
      stroke-linejoin: round;
      stroke-linecap: round;
      stroke-width: 1.7;
    }
    .star-sky-star-lit svg path {
      fill: var(--gold);
      stroke: var(--gold-deep);
      filter: none;
      transform-origin: center;
    }
    .star-sky-star-locked {
      cursor: default;
    }
    .star-sky-star-locked svg path {
      fill: rgba(255, 255, 255, .04);
      stroke: rgba(224, 214, 180, .38);
    }
    .star-sky-star-lit:hover svg,
    .star-sky-star-lit:focus-visible svg {
      transform: scale(1.18) rotate(-4deg);
    }
    .star-sky-star:focus-visible {
      outline: 3px solid var(--gold);
      outline-offset: 2px;
    }
    .star-sky-detail {
      display: grid;
      gap: 5px;
      margin: 12px 14px 0;
      padding: 12px 14px;
      border: 1px solid rgba(255, 255, 255, .14);
      border-radius: 8px;
      background: rgba(8, 17, 31, .42);
    }
    .star-sky-detail strong {
      color: #fff7d4;
      font-size: 14px;
    }
    .star-sky-detail span {
      color: rgba(248, 251, 255, .78);
      font-size: 12px;
      line-height: 1.5;
    }
    .star-sky-list {
      display: none;
      margin: 0;
      padding: 0 22px 20px;
      list-style: none;
    }
    .star-sky-list li {
      display: grid;
      gap: 4px;
      padding: 8px 0;
      border-top: 1px solid rgba(255, 255, 255, .12);
    }
    .star-sky-list strong {
      color: #fff7d4;
    }
    .star-sky-list span {
      color: rgba(248, 251, 255, .78);
      font-size: 12px;
      line-height: 1.45;
    }
    .share-dialog {
      width: min(520px, calc(100vw - 32px));
      padding: 0;
      border: 1px solid var(--line);
      border-radius: 10px;
      background: var(--surface);
      color: var(--ink);
      box-shadow: var(--shadow-raised);
    }
    .share-dialog::backdrop {
      background: rgba(12, 19, 16, .54);
    }
    .share-dialog-body {
      display: grid;
      gap: 15px;
      padding: 20px;
    }
    .share-dialog-heading {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 12px;
    }
    .share-dialog-heading h2 {
      margin: 0;
      font-size: 20px;
      line-height: 1.25;
    }
    .share-dialog-heading p {
      margin: 6px 0 0;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.5;
    }
    .share-close {
      flex: 0 0 auto;
      width: 32px;
      height: 32px;
      padding: 0;
      border-radius: 50%;
      font-size: 20px;
      line-height: 1;
    }
    .share-field {
      display: grid;
      gap: 7px;
      color: var(--muted);
      font-size: 12px;
    }
    .share-field select {
      width: 100%;
    }
    .share-preview-note {
      padding: 12px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface-soft);
      color: var(--muted);
      font-size: 12px;
      line-height: 1.55;
    }
    .share-actions {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      flex-wrap: wrap;
    }
    .share-actions button {
      background: var(--green);
      border-color: var(--green);
      color: #fff;
      font-weight: 700;
    }
    html[data-theme="dark"] .share-actions button {
      color: #102017;
    }
    #share-status {
      color: var(--muted);
      font-size: 12px;
    }
    .period-panel {
      margin: 0 0 14px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface-soft);
    }
    .period-panel > summary {
      display: block;
      padding: 12px 14px;
      cursor: pointer;
      font-weight: 700;
      list-style: none;
    }
    .period-panel > summary::-webkit-details-marker {
      display: none;
    }
    .period-panel > summary::before {
      content: "";
      display: inline-block;
      width: 0;
      height: 0;
      margin-right: 9px;
      border-left: 6px solid currentColor;
      border-top: 5px solid transparent;
      border-bottom: 5px solid transparent;
      vertical-align: middle;
      transition: transform 0.15s ease;
    }
    .period-panel[open] > summary::before {
      transform: rotate(90deg);
    }
    .period-panel > summary em {
      display: block;
      margin-top: 5px;
      color: var(--muted);
      font-size: 12px;
      font-style: normal;
      font-weight: 400;
    }
    .period-body {
      padding: 0 14px 14px;
      min-width: 0;
    }
    .period-tabs {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      margin-bottom: 12px;
      padding: 4px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
    }
    .period-tabs button {
      flex: 1 1 110px;
      min-width: 0;
      min-height: 34px;
      border: 1px solid transparent;
      border-radius: 6px;
      background: transparent;
      color: var(--muted);
      font: inherit;
      font-size: 12px;
      font-weight: 700;
      cursor: pointer;
    }
    .period-tabs button[aria-selected="true"] {
      border-color: var(--badge-border);
      background: var(--badge-bg);
      color: var(--green);
    }
    .period-report-view {
      min-width: 0;
    }
    .period-report-heading {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 8px 12px;
      margin-bottom: 10px;
    }
    .period-report-heading > div {
      min-width: 0;
    }
    .period-report-heading h3 {
      margin: 0;
      font-size: 17px;
    }
    .period-report-heading p {
      margin: 3px 0 0;
      color: var(--muted);
      font-size: 12px;
      overflow-wrap: anywhere;
    }
    .period-confidence {
      flex: 0 0 auto;
      padding: 3px 8px;
      border-radius: 999px;
      background: var(--tag-bg);
      color: var(--tag-ink);
      font-size: 11px;
      font-weight: 700;
      white-space: nowrap;
    }
    .period-confidence-high {
      background: var(--badge-bg);
      color: var(--green);
    }
    .period-confidence-medium {
      background: var(--estimate-bg);
      color: var(--amber);
    }
    .period-confidence-low {
      background: var(--low-bg);
      color: var(--coral);
    }
    .period-metrics {
      display: grid;
      grid-template-columns: repeat(3, minmax(150px, 1fr));
      gap: 8px;
      margin-bottom: 10px;
    }
    .period-metric {
      min-width: 0;
      min-height: 74px;
      padding: 11px 12px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
    }
    .period-metric span {
      display: block;
      color: var(--muted);
      font-size: 11px;
      font-weight: 700;
    }
    .period-metric strong {
      display: block;
      margin-top: 7px;
      font-size: 19px;
      line-height: 1.2;
      overflow-wrap: anywhere;
    }
    .period-content-grid {
      display: grid;
      grid-template-columns: minmax(0, 1.05fr) minmax(0, .95fr);
      gap: 10px;
      min-width: 0;
    }
    .period-subpanel {
      min-width: 0;
      padding: 12px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
    }
    .period-subpanel h4 {
      margin: 0 0 9px;
      font-size: 14px;
    }
    .period-books {
      display: grid;
      gap: 7px;
      min-width: 0;
    }
    .period-chart-scroll {
      min-width: 0;
      padding-bottom: 4px;
      overflow-x: auto;
      overflow-y: hidden;
      overscroll-behavior-x: contain;
    }
    .period-bar-chart {
      min-width: 0;
    }
    .period-bars {
      --period-plot-height: 250px;
      --period-axis-height: 20px;
      --period-row-gap: 5px;
      --period-bar-width: 44px;
      --period-grid-line: var(--line);
      --period-baseline-line: var(--line);
      position: relative;
      display: grid;
      grid-template-columns:
        repeat(var(--period-bucket-count, 1), minmax(32px, 1fr));
      align-items: end;
      gap: 6px;
      min-width: max(
        100%,
        var(--period-bars-min-width, 100%)
      );
    }
    .period-grid-lines {
      position: absolute;
      left: 0;
      right: 0;
      bottom: calc(
        var(--period-axis-height) + var(--period-row-gap)
      );
      height: var(--period-plot-height);
      z-index: 0;
      pointer-events: none;
      opacity: .7;
      background-image: repeating-linear-gradient(
        to top,
        var(--period-grid-line) 0 1px,
        transparent 1px calc(var(--period-plot-height) / 4)
      );
      border-bottom: 1px solid var(--period-baseline-line);
    }
    .period-bar {
      position: relative;
      z-index: 1;
      display: grid;
      grid-template-rows:
        auto var(--period-plot-height) var(--period-axis-height);
      gap: var(--period-row-gap);
      align-items: end;
      min-width: 0;
      color: var(--muted);
      font-size: 10px;
      text-align: center;
      border-radius: 8px;
    }
    .period-bar-value {
      min-width: 0;
      min-height: 15px;
      color: var(--ink);
      font-size: 10px;
      font-weight: 700;
      line-height: 1.25;
      overflow-wrap: anywhere;
      font-variant-numeric: tabular-nums;
    }
    .period-bar-value-short {
      display: none;
    }
    .period-bar-chart.is-dense .period-bar-value-full {
      display: none;
    }
    .period-bar-chart.is-dense .period-bar-value-short {
      display: inline;
    }
    .period-bar-track {
      position: relative;
      display: flex;
      align-items: flex-end;
      justify-content: center;
      align-self: end;
      width: min(var(--period-bar-width), 100%);
      height: 100%;
      margin: 0 auto;
      min-width: 0;
      overflow: hidden;
      border-radius: 6px 6px 2px 2px;
    }
    .period-bar-fill {
      display: block;
      width: 100%;
      min-height: 0;
      border-radius: 6px 6px 0 0;
      background: linear-gradient(
        180deg,
        var(--period-color, var(--green)),
        var(--period-color, var(--green))
      );
      box-shadow: inset 0 1px 0 rgba(255, 255, 255, .18);
      transition: height .25s ease, filter .15s ease;
    }
    .period-bar[data-period-state="no-record"] .period-bar-fill {
      background: var(--muted);
      opacity: .16;
    }
    .period-bar-empty {
      display: none;
      position: absolute;
      left: 50%;
      bottom: 0;
      width: 60%;
      height: 3px;
      transform: translateX(-50%);
      border-radius: 3px;
      background: var(--muted);
      opacity: .32;
    }
    .period-bar[data-period-state="no-record"] .period-bar-empty {
      display: block;
    }
    .period-bar:hover .period-bar-fill,
    .period-bar:focus-visible .period-bar-fill {
      filter: brightness(1.08) saturate(1.06);
    }
    .period-bar:focus-visible {
      outline: 2px solid var(--focus-ring);
      outline-offset: 2px;
    }
    .period-bar-label {
      min-width: 0;
      color: var(--muted);
      font-size: 10px;
      line-height: var(--period-axis-height);
      overflow-wrap: anywhere;
      font-variant-numeric: tabular-nums;
    }
    .period-book {
      min-width: 0;
      padding: 8px 9px;
      border: 1px solid var(--line);
      border-radius: 7px;
      background: var(--surface-soft);
    }
    .period-book-heading {
      display: grid;
      grid-template-columns: 28px minmax(0, 1fr);
      gap: 6px;
      align-items: baseline;
      min-width: 0;
    }
    .period-book-heading span {
      color: var(--muted);
      font-size: 10px;
      font-weight: 700;
      font-variant-numeric: tabular-nums;
    }
    .period-book-heading strong {
      min-width: 0;
      font-size: 12px;
      line-height: 1.35;
      overflow-wrap: anywhere;
    }
    .period-book-values {
      display: flex;
      flex-wrap: wrap;
      gap: 4px 8px;
      margin-top: 5px;
      color: var(--muted);
      font-size: 10px;
      font-variant-numeric: tabular-nums;
    }
    .period-conclusion {
      margin: 10px 0 0;
      padding: 9px 10px;
      border-left: 4px solid var(--green);
      border-radius: 6px;
      background: var(--surface-muted);
      font-size: 12px;
      overflow-wrap: anywhere;
    }
    .period-conclusion.is-insufficient {
      border-left-color: var(--amber);
    }
    .period-coverage {
      margin-top: 10px;
      color: var(--muted);
      font-size: 11px;
      line-height: 1.5;
    }
    .period-coverage p {
      margin: 5px 0 0;
      overflow-wrap: anywhere;
    }
    .period-coverage strong {
      color: var(--ink);
    }
    .period-coverage ul {
      margin: 3px 0 0 18px;
      padding: 0;
    }
    .period-coverage-meta {
      padding-top: 5px;
      border-top: 1px solid var(--line);
    }
    .annual-report-panel {
      margin: 0 0 14px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface-soft);
    }
    .annual-report-panel > summary {
      display: block;
      padding: 12px 14px;
      cursor: pointer;
      font-weight: 700;
      list-style: none;
    }
    .annual-report-panel > summary::-webkit-details-marker {
      display: none;
    }
    .annual-report-panel > summary::before {
      content: "";
      display: inline-block;
      width: 0;
      height: 0;
      margin-right: 9px;
      border-left: 6px solid currentColor;
      border-top: 5px solid transparent;
      border-bottom: 5px solid transparent;
      vertical-align: middle;
      transition: transform .15s ease;
    }
    .annual-report-panel[open] > summary::before {
      transform: rotate(90deg);
    }
    .annual-report-panel > summary em {
      display: block;
      margin-top: 5px;
      color: var(--muted);
      font-size: 12px;
      font-style: normal;
      font-weight: 400;
    }
    .annual-report-body {
      padding: 0 14px 14px;
      min-width: 0;
    }
    .annual-report-sheet {
      min-width: 0;
      padding: 16px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
      box-shadow: var(--shadow-card);
    }
    .annual-report-heading {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 8px 12px;
      margin-bottom: 12px;
    }
    .annual-report-heading > div {
      min-width: 0;
    }
    .annual-report-heading h3 {
      margin: 0;
      font-size: 22px;
      line-height: 1.2;
    }
    .annual-report-heading p {
      margin: 4px 0 0;
      color: var(--muted);
      font-size: 12px;
      overflow-wrap: anywhere;
    }
    .annual-metrics {
      display: grid;
      grid-template-columns: repeat(4, minmax(130px, 1fr));
      gap: 10px;
    }
    .annual-metric {
      min-width: 0;
      min-height: 76px;
      padding: 11px 12px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface-soft);
    }
    .annual-metric span {
      display: block;
      color: var(--muted);
      font-size: 11px;
      font-weight: 700;
    }
    .annual-metric strong {
      display: block;
      margin-top: 7px;
      font-size: 19px;
      line-height: 1.2;
      overflow-wrap: anywhere;
      font-variant-numeric: tabular-nums;
    }
    .annual-content-grid {
      display: grid;
      grid-template-columns: minmax(0, 1.1fr) minmax(0, .9fr);
      gap: 10px;
      margin-top: 10px;
    }
    .annual-subpanel {
      min-width: 0;
      padding: 12px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface-soft);
    }
    .annual-subpanel h4 {
      margin: 0 0 9px;
      font-size: 14px;
    }
    .annual-books {
      display: grid;
      gap: 7px;
      min-width: 0;
    }
    .annual-book {
      display: grid;
      grid-template-columns: 30px minmax(0, 1fr);
      gap: 8px;
      align-items: start;
      min-width: 0;
      padding: 8px 9px;
      border: 1px solid var(--line);
      border-radius: 7px;
      background: var(--surface);
    }
    .annual-book-rank {
      color: var(--muted);
      font-size: 10px;
      font-weight: 700;
      font-variant-numeric: tabular-nums;
    }
    .annual-book-copy {
      min-width: 0;
    }
    .annual-book-copy strong {
      display: block;
      font-size: 12px;
      line-height: 1.35;
      overflow-wrap: anywhere;
    }
    .annual-book-copy small {
      display: block;
      margin-top: 4px;
      color: var(--muted);
      font-size: 10px;
      font-variant-numeric: tabular-nums;
      overflow-wrap: anywhere;
    }
    .annual-coverage p {
      margin: 0 0 7px;
      color: var(--muted);
      font-size: 11px;
      line-height: 1.5;
      overflow-wrap: anywhere;
    }
    .annual-coverage p strong {
      color: var(--ink);
    }
    .annual-coverage-meta {
      margin-top: 9px !important;
      padding-top: 7px;
      border-top: 1px solid var(--line);
    }
    .source-tag {
      background: var(--tag-bg);
      color: var(--tag-ink);
    }
    .source-exact-body { background: var(--badge-bg); color: var(--green); }
    .source-fmcache-fallback {
      background: var(--deviation-bg);
      color: var(--violet);
    }
    .source-kindle-estimate {
      background: var(--language-bg);
      color: var(--blue);
    }
    .quality-summary {
      display: grid;
      grid-template-columns: repeat(5, minmax(140px, 1fr));
      gap: 10px;
      margin-bottom: 14px;
    }
    .quality-card {
      min-width: 0;
      padding: 13px 14px;
      border: 1px solid var(--line);
      border-left: 4px solid var(--blue);
      border-radius: 8px;
      background: var(--surface);
    }
    .quality-card > span {
      display: block;
      color: var(--muted);
      font-size: 12px;
    }
    .quality-card > strong {
      display: block;
      margin-top: 6px;
      font-size: 20px;
      line-height: 1.15;
    }
    .quality-card > small {
      display: block;
      margin-top: 6px;
      color: var(--muted);
      overflow-wrap: anywhere;
    }
    .quality-good { border-left-color: var(--green); }
    .quality-warn { border-left-color: var(--amber); }
    .quality-bad { border-left-color: var(--coral); }
    .quality-neutral { border-left-color: var(--blue); }
    .trend-summary {
      display: grid;
      grid-template-columns: repeat(4, minmax(150px, 1fr));
      gap: 10px;
      margin-bottom: 14px;
    }
    .trend-card {
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 14px 15px;
      min-width: 0;
    }
    .trend-card > span {
      display: block;
      color: var(--muted);
      font-size: 12px;
    }
    .trend-card > strong {
      display: block;
      margin-top: 7px;
      font-size: 21px;
      line-height: 1.1;
    }
    .trend-card > small {
      display: block;
      margin-top: 7px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .delta-up { color: var(--green); }
    .delta-down { color: var(--coral); }
    .delta-neutral { color: var(--muted); }
    .trend-panel { margin-bottom: 14px; }
    .trend-heading {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 12px;
      min-width: 0;
    }
    .trend-heading > div { min-width: 0; }
    .trend-heading p {
      margin: 5px 0 0;
      color: var(--muted);
      font-size: 12px;
      overflow-wrap: anywhere;
    }
    .trend-metric-switch {
      display: inline-flex;
      flex-wrap: wrap;
      gap: 3px;
      max-width: 100%;
      padding: 3px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface-muted);
    }
    .trend-metric-switch button {
      border: 0;
      border-radius: 6px;
      background: transparent;
      color: var(--muted);
      padding: 7px 10px;
      font: inherit;
      font-size: 12px;
      cursor: pointer;
      white-space: nowrap;
    }
    .trend-metric-switch button[aria-pressed="true"] {
      background: var(--surface);
      color: var(--ink);
      box-shadow: 0 1px 4px rgba(20, 35, 30, .12);
    }
    .trend-chart-heading {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      margin-top: 14px;
      min-width: 0;
    }
    .trend-chart-heading > p {
      margin: 0;
      color: var(--muted);
      font-size: 11px;
      text-align: right;
      overflow-wrap: anywhere;
    }
    .trend-chart {
      position: relative;
      margin-top: 10px;
      padding: 8px 0 2px;
      border-top: 1px solid var(--line);
      border-bottom: 1px solid var(--line);
      min-width: 0;
    }
    .trend-chart-view[hidden] { display: none; }
    .trend-chart-svg {
      display: block;
      width: 100%;
      height: auto;
      max-height: 300px;
      overflow: visible;
    }
    .trend-grid-line {
      stroke: var(--line);
      stroke-width: 1;
      stroke-dasharray: 3 4;
      opacity: .72;
    }
    .trend-axis-label {
      fill: var(--muted);
      font-size: 10px;
      font-variant-numeric: tabular-nums;
    }
    .trend-x-label { font-size: 9px; }
    .trend-line {
      fill: none;
      stroke: var(--blue);
      stroke-width: 3;
      stroke-linecap: round;
      stroke-linejoin: round;
    }
    .trend-point circle {
      fill: var(--surface);
      stroke: var(--blue);
      stroke-width: 2.2;
      cursor: crosshair;
      transition: fill .15s ease, stroke .15s ease, r .15s ease;
    }
    .trend-point:focus { outline: none; }
    .trend-point:hover circle,
    .trend-point:focus circle {
      fill: var(--amber);
      stroke: var(--ink);
      r: 5.2px;
    }
    .trend-chart-note {
      margin: 2px 4px 8px;
      color: var(--muted);
      font-size: 11px;
      line-height: 1.45;
    }
    .trend-tooltip {
      position: absolute;
      z-index: 8;
      max-width: min(260px, calc(100% - 16px));
      transform: translateX(-50%);
      border: 1px solid var(--line);
      border-radius: 7px;
      background: var(--surface);
      color: var(--ink);
      padding: 7px 9px;
      box-shadow: 0 8px 22px rgba(20, 35, 30, .18);
      font-size: 11px;
      line-height: 1.35;
      pointer-events: none;
      overflow-wrap: anywhere;
    }
    .trend-subgrid {
      display: grid;
      grid-template-columns: minmax(0, 1.05fr) minmax(0, .95fr);
      gap: 18px;
      margin-top: 14px;
      padding-top: 14px;
      border-top: 1px solid var(--line);
    }
    .trend-subpanel { min-width: 0; }
    .trend-subpanel h3 {
      margin: 0 0 9px;
      font-size: 14px;
    }
    .trend-snapshot-details {
      margin-top: 14px;
      border-top: 1px solid var(--line);
    }
    .trend-snapshot-details > summary {
      display: flex;
      align-items: baseline;
      justify-content: space-between;
      gap: 10px;
      padding: 11px 0 8px;
      cursor: pointer;
      list-style: none;
    }
    .trend-snapshot-details > summary::-webkit-details-marker {
      display: none;
    }
    .trend-snapshot-details > summary::before {
      content: "+";
      flex: 0 0 auto;
      color: var(--blue);
      font-weight: 800;
    }
    .trend-snapshot-details[open] > summary::before { content: "−"; }
    .trend-snapshot-details > summary em {
      color: var(--muted);
      font-size: 11px;
      font-style: normal;
      text-align: right;
    }
    .trend-snapshot-table {
      width: 100%;
      min-width: 760px;
      border-collapse: collapse;
      font-size: 11px;
    }
    .trend-snapshot-table th,
    .trend-snapshot-table td {
      padding: 8px 9px;
      border-top: 1px solid var(--line);
      text-align: left;
      vertical-align: top;
    }
    .trend-snapshot-table th {
      color: var(--muted);
      font-weight: 700;
      white-space: nowrap;
    }
    .trend-snapshot-table small {
      display: block;
      max-width: 180px;
      color: var(--muted);
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .trend-source-note {
      display: grid;
      gap: 4px;
      margin: 12px 0 0;
      color: var(--muted);
      font-size: 11px;
      line-height: 1.5;
    }
    .trend-source-note strong {
      color: var(--ink);
      font-weight: 650;
    }
    .trend-panels {
      display: grid;
      grid-template-columns: minmax(0, 1.15fr) minmax(0, .85fr);
      gap: 14px;
      margin-bottom: 14px;
    }
    .movement-list {
      display: grid;
      gap: 9px;
    }
    .movement-item {
      display: grid;
      grid-template-columns: minmax(150px, .75fr) minmax(260px, 1.25fr);
      gap: 12px;
      align-items: center;
      padding: 10px 11px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface-soft);
    }
    .movement-title {
      min-width: 0;
      display: grid;
      gap: 3px;
    }
    .movement-title strong {
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .movement-title small {
      color: var(--muted);
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .new-tag {
      display: inline-flex;
      width: fit-content;
      border-radius: 999px;
      padding: 2px 7px;
      background: var(--language-bg);
      color: var(--blue);
      font-size: 11px;
      font-weight: 700;
    }
    .movement-metrics {
      display: grid;
      grid-template-columns: repeat(4, minmax(58px, 1fr));
      gap: 6px;
      min-width: 0;
    }
    .movement-metrics span {
      display: grid;
      gap: 1px;
      min-width: 0;
    }
    .movement-metrics em {
      color: var(--muted);
      font-size: 10px;
      font-style: normal;
    }
    .movement-metrics b {
      font-size: 12px;
      white-space: nowrap;
    }
    .timeline {
      display: grid;
      gap: 0;
    }
    .timeline-item {
      display: grid;
      grid-template-columns: 104px 14px minmax(0, 1fr);
      gap: 8px;
      align-items: start;
      min-height: 54px;
      position: relative;
    }
    .timeline-item:not(:last-child)::after {
      content: "";
      position: absolute;
      left: 111px;
      top: 13px;
      bottom: -2px;
      width: 1px;
      background: var(--line);
    }
    .timeline-item time {
      color: var(--muted);
      font-size: 11px;
      text-align: right;
      padding-top: 1px;
    }
    .timeline-marker {
      width: 9px;
      height: 9px;
      margin-top: 3px;
      border-radius: 50%;
      background: var(--green);
      border: 2px solid var(--timeline-ring);
      position: relative;
      z-index: 1;
    }
    .timeline-item > div {
      min-width: 0;
      padding-bottom: 12px;
    }
    .timeline-item strong {
      display: block;
      font-size: 13px;
    }
    .timeline-item small {
      display: block;
      margin-top: 2px;
      color: var(--muted);
    }
    .daily-panel {
      margin-bottom: 14px;
    }
    .daily-panel h2 {
      margin-bottom: 12px;
    }
    .daily-bars {
      display: grid;
      grid-template-columns: repeat(30, minmax(0, 1fr));
      gap: 2px;
      align-items: end;
      min-height: 178px;
    }
    .daily-column {
      min-width: 0;
      display: grid;
      grid-template-rows: 128px 44px;
      gap: 6px;
    }
    .daily-track {
      height: 128px;
      display: flex;
      align-items: flex-end;
      overflow: hidden;
      border-radius: 4px;
      background: var(--surface-chart);
    }
    .daily-track i {
      display: block;
      width: 100%;
      min-height: 0;
      background: var(--green);
      border-radius: 3px 3px 0 0;
    }
    .daily-column time {
      color: var(--muted);
      font-size: 9px;
      line-height: 1;
      writing-mode: vertical-rl;
      transform: rotate(180deg);
      justify-self: center;
    }
    .heatmap-panel { margin-bottom: 14px; }
    .heatmap-heading {
      display: flex;
      align-items: end;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 12px;
    }
    .heatmap-heading h2 { margin-bottom: 2px; }
    .heatmap-heading p {
      margin: 0;
      color: var(--muted);
      font-size: 12px;
    }
    .heatmap-scroll {
      overflow-x: auto;
      padding: 2px 0 8px;
    }
    .heatmap-weekdays {
      display: grid;
      grid-template-rows: repeat(7, 13px);
      gap: 3px;
      float: left;
      width: 16px;
      margin-right: 6px;
      color: var(--muted);
      font-size: 9px;
      line-height: 13px;
    }
    .heatmap-grid {
      display: grid;
      grid-template-rows: repeat(7, 13px);
      grid-auto-flow: column;
      grid-auto-columns: 13px;
      gap: 3px;
      width: max-content;
    }
    .heatmap-day {
      display: block;
      width: 13px;
      height: 13px;
      border-radius: 3px;
      background: var(--surface-track);
    }
    .heatmap-day-empty { background: transparent; }
    .level-0 { background: var(--surface-track); }
    .level-1 { background: #cce4da; }
    .level-2 { background: #8dc5ad; }
    .level-3 { background: #4e9b7b; }
    .level-4 { background: #24664d; }
    .heatmap-footer {
      clear: both;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding-top: 8px;
      color: var(--muted);
      font-size: 11px;
      flex-wrap: wrap;
    }
    .heatmap-source-active,
    .heatmap-source-fallback {
      min-width: 0;
      overflow-wrap: anywhere;
    }
    .heatmap-source-active {
      color: var(--green);
      font-weight: 700;
    }
    .heatmap-source-fallback { color: var(--amber); }
    .heatmap-legend {
      display: inline-flex;
      align-items: center;
      gap: 3px;
      white-space: nowrap;
    }
    .heatmap-legend em {
      font-style: normal;
      margin: 0 2px;
    }
    .heatmap-legend i {
      display: inline-block;
      width: 11px;
      height: 11px;
      border-radius: 2px;
    }
    .delta-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(76px, 1fr));
      gap: 3px 9px;
    }
    .delta-grid span {
      display: flex;
      justify-content: space-between;
      gap: 6px;
      white-space: nowrap;
    }
    .delta-grid em {
      color: var(--muted);
      font-style: normal;
      font-size: 11px;
    }
    .delta-grid strong { font-size: 12px; }
    .empty.compact { padding: 18px; }
    .panels {
      display: grid;
      grid-template-columns: minmax(0, 1.05fr) minmax(0, .95fr);
      gap: 14px;
      margin-bottom: 14px;
    }
    .panels > *, .panel, .rank-list, .rank-item,
    .rank-label > *, .progress-card, .progress-card > div {
      min-width: 0;
    }
    .panel {
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 18px;
      overflow: hidden;
    }
    .panel h2 { margin: 0 0 14px; font-size: 17px; }
    .rank-list {
      display: grid;
      grid-template-columns: minmax(0, 1fr);
      min-width: 0;
      gap: 11px;
    }
    .rank-item {
      display: grid;
      grid-template-columns: minmax(0, 1fr);
      gap: 6px;
    }
    .rank-label {
      display: flex;
      justify-content: space-between;
      gap: 14px;
      align-items: baseline;
      min-width: 0;
    }
    .rank-label strong {
      min-width: 0;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .rank-label span {
      flex: 0 0 auto;
      color: var(--muted);
      font-size: 12px;
      white-space: nowrap;
    }
    .rank-track {
      height: 9px;
      border-radius: 999px;
      background: var(--surface-track);
      overflow: hidden;
    }
    .rank-track i { display: block; height: 100%; border-radius: inherit; }
    .progress-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
    }
    .progress-card {
      display: grid;
      grid-template-columns: 66px 1fr;
      gap: 12px;
      align-items: center;
      min-height: 88px;
      padding: 10px;
      border: 1px solid var(--line);
      border-radius: 8px;
    }
    .progress-ring {
      width: 62px;
      height: 62px;
      border-radius: 50%;
      display: grid;
      place-items: center;
      background: conic-gradient(
        var(--ring-color) calc(var(--value) * 1%),
        var(--ring-empty) 0
      );
      position: relative;
    }
    .progress-ring::after {
      content: "";
      position: absolute;
      inset: 7px;
      background: var(--surface);
      border-radius: 50%;
    }
    .progress-ring strong {
      position: relative;
      z-index: 1;
      font-size: 12px;
    }
    .progress-card h3 {
      margin: 0 0 4px;
      font-size: 13px;
      line-height: 1.25;
      display: -webkit-box;
      -webkit-line-clamp: 2;
      -webkit-box-orient: vertical;
      overflow: hidden;
    }
    .progress-card p { margin: 0; color: var(--muted); font-size: 12px; }
    .word-count-summary {
      display: flex;
      flex-wrap: wrap;
      gap: 3px 8px;
    }
    .word-count-summary span {
      white-space: nowrap;
    }
    .tag-row, .tags-cell {
      display: flex;
      flex-wrap: wrap;
      gap: 5px;
      margin-top: 7px;
      min-width: 150px;
    }
    .tag {
      display: inline-flex;
      align-items: center;
      border-radius: 999px;
      padding: 2px 7px;
      font-size: 11px;
      font-weight: 650;
      white-space: nowrap;
      background: var(--tag-bg);
      color: var(--tag-ink);
    }
    .token-tag {
      display: none;
      border: 1px solid var(--line);
      background: transparent;
      color: var(--muted);
      font-weight: 700;
    }
    html[data-show-token-labels="true"] .token-tag {
      display: inline-flex;
    }
    .language-tag { background: var(--language-bg); color: var(--blue); }
    .exact-tag { background: var(--badge-bg); color: var(--green); }
    .estimate-tag { background: var(--estimate-bg); color: var(--amber); }
    .mode-tag { background: var(--deviation-bg); color: var(--violet); }
    .confidence-tag {
      display: inline-flex;
      width: fit-content;
      border-radius: 999px;
      padding: 2px 7px;
      font-size: 11px;
      font-weight: 700;
      white-space: nowrap;
      background: var(--tag-bg);
      color: var(--tag-ink);
    }
    .confidence-high { background: var(--badge-bg); color: var(--green); }
    .confidence-medium { background: var(--estimate-bg); color: var(--amber); }
    .confidence-low { background: var(--low-bg); color: var(--coral); }
    .confidence-not-applicable {
      background: var(--deviation-bg);
      color: var(--violet);
    }
    .deviation-tag { background: var(--deviation-bg); color: var(--violet); }
    .toolbar {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: center;
      margin: 18px 0 10px;
    }
    .filters {
      display: flex;
      flex: 1;
      flex-wrap: wrap;
      gap: 8px;
      min-width: 0;
    }
    .column-picker {
      position: relative;
      flex: 0 0 auto;
    }
    .column-picker summary {
      list-style: none;
    }
    .column-picker summary::-webkit-details-marker { display: none; }
    .column-menu {
      position: absolute;
      right: 0;
      top: calc(100% + 6px);
      z-index: 4;
      display: grid;
      grid-template-columns: repeat(2, minmax(130px, 1fr));
      gap: 5px 12px;
      width: min(360px, calc(100vw - 32px));
      padding: 12px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
      box-shadow: 0 10px 26px rgba(25, 40, 32, .16);
    }
    .column-menu label {
      display: flex;
      align-items: center;
      gap: 7px;
      min-width: 0;
      color: var(--ink);
      font-size: 12px;
    }
    .column-menu input {
      accent-color: var(--green);
    }
    input[type="search"], select {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 10px 12px;
      background: var(--surface);
      color: var(--ink);
      font: inherit;
    }
    input[type="search"] {
      width: min(420px, 100%);
      flex: 1 1 260px;
    }
    select { min-width: 112px; }
    button {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 9px 12px;
      background: var(--surface);
      color: var(--ink);
      font: inherit;
      cursor: pointer;
    }
    .export-button {
      background: var(--green);
      border-color: var(--green);
      color: #fff;
    }
    html[data-theme="dark"] .export-button {
      color: #102017;
    }
    .toolbar-actions {
      display: flex;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
    }
    .table-wrap {
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: auto;
      max-width: 100%;
    }
    table {
      width: 100%;
      min-width: 1760px;
      border-collapse: collapse;
    }
    th, td {
      padding: 11px 12px;
      border-bottom: 1px solid var(--line);
      text-align: left;
      vertical-align: middle;
    }
    th {
      position: sticky;
      top: 0;
      z-index: 1;
      background: var(--surface-soft);
      color: var(--muted);
      cursor: pointer;
      white-space: nowrap;
    }
    tbody tr:hover { background: var(--hover); }
    .rank { color: var(--muted); width: 42px; }
    .book-cell strong { display: block; max-width: 420px; }
    .book-cell span { display: block; color: var(--muted); font-size: 12px; }
    .tags-cell { margin-top: 0; }
    .word-count-cell {
      display: grid;
      gap: 5px;
      min-width: 190px;
    }
    .word-count-cell span {
      display: flex;
      align-items: baseline;
      justify-content: space-between;
      gap: 8px;
    }
    .word-count-cell em {
      color: var(--muted);
      font-size: 10px;
      font-style: normal;
    }
    .word-count-cell strong { white-space: nowrap; }
    .source-value {
      display: grid;
      gap: 2px;
      min-width: 78px;
    }
    .source-value small,
    .source-note {
      display: block;
      color: var(--muted);
      font-size: 10px;
      line-height: 1.2;
    }
    .source-note { margin-top: 4px; }
    .quality-cell { min-width: 190px; }
    .quality-cell-content {
      display: grid;
      gap: 5px;
      max-width: 220px;
    }
    .quality-cell-content small {
      color: var(--muted);
      font-size: 10px;
      line-height: 1.35;
      overflow-wrap: anywhere;
    }
    .mini-progress {
      display: grid;
      grid-template-columns: 52px minmax(90px, 1fr);
      gap: 8px;
      align-items: center;
      min-width: 170px;
    }
    .mini-progress > div {
      height: 7px;
      background: var(--surface-track);
      border-radius: 999px;
      overflow: hidden;
    }
    .mini-progress i { display: block; height: 100%; border-radius: inherit; }
    .empty {
      padding: 28px;
      color: var(--muted);
      text-align: center;
    }
    footer {
      margin-top: 14px;
      color: var(--muted);
      font-size: 12px;
    }
    @media (max-width: 1080px) {
      .summary { grid-template-columns: repeat(3, minmax(130px, 1fr)); }
      .insights { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .insight[data-insight-layout="feature"] { grid-column: 1 / -1; }
      .quality-summary { grid-template-columns: repeat(3, minmax(140px, 1fr)); }
      .appearance-body {
        position: fixed;
        top: var(--appearance-top, 90px);
        right: 16px;
        left: auto;
        width: min(560px, calc(100vw - 32px));
      }
      .achievement-grid {
        grid-template-columns: repeat(3, minmax(0, 1fr));
      }
      .period-content-grid { grid-template-columns: minmax(0, 1fr); }
      .trend-summary { grid-template-columns: repeat(2, minmax(150px, 1fr)); }
      .trend-subgrid { grid-template-columns: minmax(0, 1fr); }
      .trend-panels { grid-template-columns: minmax(0, 1fr); }
      .panels { grid-template-columns: minmax(0, 1fr); }
      .cover-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .annual-metrics { grid-template-columns: repeat(2, minmax(140px, 1fr)); }
      .annual-content-grid { grid-template-columns: minmax(0, 1fr); }
    }
    @media (max-width: 720px) {
      .page { width: min(100% - 20px, 1440px); margin-top: 16px; }
      header { display: block; }
      .header-main { flex-basis: auto; }
      .header-actions {
        position: relative;
        justify-content: flex-start;
        margin-top: 10px;
      }
      .header-actions > .theme-picker,
      .header-actions > .greeting-picker {
        flex: 1 1 150px;
        justify-content: space-between;
      }
      .badge { display: inline-block; }
      h1 { font-size: 24px; }
      .summary { grid-template-columns: repeat(2, minmax(120px, 1fr)); }
      .insights { grid-template-columns: 1fr; }
      .insight[data-insight-layout="feature"] { grid-column: auto; }
      .quality-summary { grid-template-columns: repeat(2, minmax(130px, 1fr)); }
      .achievement-grid { grid-template-columns: 1fr; }
      .period-metrics { grid-template-columns: 1fr; }
      .greeting-row { align-items: flex-start; }
      .trend-summary { grid-template-columns: 1fr; }
      .trend-heading,
      .trend-chart-heading {
        align-items: flex-start;
        flex-direction: column;
      }
      .trend-chart-heading > p { text-align: left; }
      .trend-metric-switch {
        width: 100%;
        justify-content: space-between;
      }
      .trend-metric-switch button {
        flex: 1 1 92px;
        min-width: 0;
      }
      .movement-item { grid-template-columns: 1fr; }
      .movement-metrics { grid-template-columns: repeat(2, minmax(80px, 1fr)); }
      .timeline-item { grid-template-columns: 82px 14px minmax(0, 1fr); }
      .timeline-item:not(:last-child)::after { left: 89px; }
      .daily-bars { gap: 1px; }
      .daily-column time { font-size: 8px; }
      .progress-grid { grid-template-columns: 1fr; }
      .library-panel { padding: 12px; }
      .library-heading { display: block; }
      .library-actions {
        justify-content: flex-start;
        margin-top: 8px;
      }
      .cover-count-badge { display: inline-block; }
      .cover-grid { grid-template-columns: 1fr; }
      .appearance-grid { grid-template-columns: 1fr; }
      .period-tabs button { flex: 1 1 90px; }
      .period-bars {
        --period-plot-height: 190px;
        --period-axis-height: 18px;
        --period-row-gap: 4px;
        --period-bar-width: 36px;
        grid-template-columns:
          repeat(var(--period-bucket-count, 1), minmax(28px, 1fr));
        gap: 4px;
      }
      .period-bar-value-full {
        display: none;
      }
      .period-bar-value-short {
        display: inline;
      }
      .annual-report-sheet { padding: 12px; }
      .annual-metrics { grid-template-columns: 1fr; }
      .book-cover-card { grid-template-columns: 1fr; }
      .cover-frame { width: 132px; }
      .toolbar { align-items: stretch; flex-direction: column; }
      .filters { width: 100%; }
      input[type="search"] { width: 100%; flex-basis: 100%; }
      select { flex: 1 1 140px; }
      .column-menu {
        left: 0;
        right: auto;
        grid-template-columns: 1fr;
      }
      .toolbar-actions { width: 100%; }
      .toolbar-actions > * { flex: 1 1 140px; }
      .trend-snapshot-details > summary {
        align-items: flex-start;
        flex-direction: column;
      }
      .trend-snapshot-details > summary em {
        flex: 1 1 auto;
        max-width: 100%;
        text-align: left;
        white-space: normal;
        overflow-wrap: anywhere;
      }
    }
    @media (max-width: 390px) {
      .book-cover-card { grid-template-columns: 1fr; gap: 10px; }
      .cover-frame { width: 112px; }
      .cover-stat strong { font-size: 12px; }
      .achievement-card { grid-template-columns: 1fr; }
      .achievement-badge { width: 40px; height: 40px; }
      html[data-achievement-badge-style="minimal"] .achievement-card {
        grid-template-columns: 8px minmax(0, 1fr);
      }
      html[data-achievement-badge-style="minimal"] .achievement-badge {
        width: 8px;
        height: 40px;
      }
      .annual-report-heading h3 { font-size: 20px; }
      .achievement-footer small {
        max-width: 100%;
        text-align: left;
      }
      .period-tabs button { flex: 1 1 100%; }
      .period-coverage ul { margin-left: 14px; }
      .appearance-actions > button { flex: 1 1 100%; }
    }
    @page {
      size: A4 landscape;
      margin: 10mm;
    }
    @media print {
      :root,
      html[data-theme="dark"],
      html[data-theme][data-theme-preset] {
        color-scheme: light;
        --bg: #fff;
        --surface: #fff;
        --surface-soft: #fff;
        --surface-muted: #f3f3f3;
        --surface-track: #ddd;
        --surface-chart: #f3f3f3;
        --ink: #111;
        --muted: #4f5b54;
        --line: #aaa;
        --hover: #fff;
        --green: #1f6b4f;
        --blue: #2d6f9e;
        --amber: #9a5c12;
        --coral: #ad4638;
        --violet: #6557a5;
        --teal: #2f7773;
      }
      body::before,
      body::after { display: none !important; }
      body {
        background: #fff;
        background-image: none !important;
        color: #111;
      }
      .page { width: 100%; margin: 0; }
      .header-actions,
      .greeting-picker,
      .appearance-panel,
      .local-name-dialog,
      .share-dialog,
      .toolbar,
      .column-picker,
      .shelf-mode-switch,
      .shelf-fallback-wrap,
      .export-button { display: none !important; }
      .header-main::after,
      .panel > h2::after,
      .library-heading h2::after,
      .trend-heading h2::after,
      .heatmap-heading h2::after,
      .achievement-panel > summary::after,
      .period-panel > summary::after,
      .annual-report-panel > summary::after,
      .period-report-heading h3::after,
      .annual-report-heading h3::after { display: none !important; }
      .panel {
        padding: 10px;
        margin-bottom: 8px;
        break-inside: avoid;
      }
      .metric,
      .book-cover-card,
      .progress-card,
      .movement-item,
      .insight,
      .quality-card,
      .trend-card,
      .achievement-card,
      .achievement-panel,
      .period-panel,
      .period-report-view,
      .period-metric,
      .period-subpanel,
      .period-book,
      .annual-report-panel,
      .annual-report-sheet,
      .annual-metric,
      .annual-subpanel,
      .annual-book { break-inside: avoid; }
      .achievement-panel .achievement-body { display: block !important; }
      .star-sky-dialog {
        width: 100%;
        max-height: none;
        border: 1px solid var(--line);
        color: #111;
        background: #fff;
      }
      .star-sky-dialog::backdrop { display: none; }
      .star-sky-shell {
        border: 0;
        background: #fff;
        color: #111;
        box-shadow: none;
      }
      .star-sky-header h2,
      .star-sky-list strong {
        color: #111;
      }
      .star-sky-header p,
      .star-sky-list span,
      .star-sky-detail span {
        color: #333;
      }
      .star-sky-close,
      .star-sky-field {
        display: none !important;
      }
      .star-sky-list {
        display: block;
      }
      .period-panel .period-body { display: block !important; }
      .annual-report-panel .annual-report-body { display: block !important; }
      .period-report-view[hidden] { display: block !important; }
      .period-tabs { display: none !important; }
      .period-report-view {
        break-before: page;
      }
      .period-report-view:first-of-type { break-before: auto; }
      .period-chart-scroll { overflow: visible !important; }
      .period-bars {
        --period-plot-height: 140px;
        --period-axis-height: 16px;
        --period-row-gap: 3px;
        --period-bar-width: 22px;
        grid-template-columns:
          repeat(var(--period-bucket-count, 1), minmax(8px, 1fr)) !important;
        gap: 3px !important;
        min-width: 0 !important;
      }
      .period-grid-lines { opacity: .55; }
      .period-bar-fill,
      .period-bar-empty,
      .insight-share-track i,
      .achievement-panel,
      .achievement-card,
      .achievement-badge,
      .achievement-progress,
      .achievement-progress i,
      .achievement-star svg,
      .achievement-star svg path {
        print-color-adjust: exact;
        -webkit-print-color-adjust: exact;
      }
      .period-bar-value-full {
        display: none !important;
      }
      .period-bar-value-short {
        display: inline !important;
      }
      .cover-frame,
      .cover-placeholder { display: none; }
      .annual-report-sheet,
      .annual-metric,
      .annual-subpanel,
      .annual-book { box-shadow: none; }
      .book-cover-card {
        grid-template-columns: 1fr;
        min-height: 0;
      }
      .table-wrap {
        overflow: visible;
        border-color: var(--line);
      }
      table {
        width: 100%;
        min-width: 0;
        font-size: 8.5px;
      }
      th, td { padding: 5px 6px; }
      th {
        position: static;
        background: #fff;
      }
      thead { display: table-header-group; }
      tr { break-inside: avoid; }
      a { color: inherit; text-decoration: none; }
    }
    @media (forced-colors: active) {
      .achievement-star svg path {
        fill: Canvas;
        stroke: CanvasText;
      }
      .achievement-star.star-lit svg path,
      .achievement-star.star-next svg path {
        fill: Highlight;
        stroke: Highlight;
      }
      .star-sky-star svg path {
        fill: Canvas;
        stroke: CanvasText;
      }
      .star-sky-star-lit svg path {
        fill: Highlight;
        stroke: Highlight;
      }
    }
    @media (prefers-reduced-motion: reduce) {
      *,
      *::before,
      *::after {
        scroll-behavior: auto !important;
        transition-duration: .01ms !important;
        animation-duration: .01ms !important;
        animation-iteration-count: 1 !important;
      }
    }
  </style>
</head>
<body>
  <main class="page">
    <header>
      <div class="header-main">
        <div class="title-line">
          <h1>
            <span class="i18n-zh">Kindle 阅读数据看板</span>
            <span class="i18n-en">Kindle Reading Dashboard</span>
          </h1>
          <div class="greeting-row">
            <p class="greeting-message" id="greeting-message"
               aria-live="polite">
              <span class="i18n-zh" id="greeting-zh">欢迎回来！愿你有一段安静的阅读时光<span class="greeting-emoticon">！(｡･ω･｡)</span></span>
              <span class="i18n-en" id="greeting-en">Welcome back! Enjoy a quiet reading moment!&nbsp;<span class="greeting-emoticon">(｡･ω･｡)</span></span>
            </p>
          </div>
        </div>
        <div class="meta">
          <span class="i18n-zh">扫描目录</span>
          <span class="i18n-en">Scan path</span>：__SOURCE__ ·
          <span class="i18n-zh">生成时间</span>
          <span class="i18n-en">Generated</span>：__GENERATED__
        </div>
      </div>
      <div class="header-actions">
        <div class="badge">
          <span class="i18n-zh">只读扫描 · 本地报告</span>
          <span class="i18n-en">Read-only scan · Local report</span>
        </div>
        <label class="theme-picker" for="theme-preset">
          <span><span class="i18n-zh">主题背景</span><span class="i18n-en">Theme background</span></span>
          <select id="theme-preset" aria-label="主题背景 / Theme background">
            <option value="forest" data-text-zh="森林" data-text-en="Forest">森林</option>
            <option value="ocean" data-text-zh="海洋" data-text-en="Ocean">海洋</option>
            <option value="sunset" data-text-zh="暖阳" data-text-en="Sunset">暖阳</option>
            <option value="graphite" data-text-zh="石墨" data-text-en="Graphite">石墨</option>
            <option value="rose" data-text-zh="玫瑰" data-text-en="Rose">玫瑰</option>
            <option value="lavender" data-text-zh="薰衣草" data-text-en="Lavender">薰衣草</option>
            <option value="sky" data-text-zh="天空" data-text-en="Sky">天空</option>
            <option value="sepia" data-text-zh="复古纸" data-text-en="Sepia">复古纸</option>
          </select>
        </label>
        <button id="language-toggle" class="control-button" type="button"
                aria-label="切换界面语言 / Switch interface language">
          English
        </button>
        <button id="theme-toggle" class="control-button" type="button"
                aria-pressed="false"
                aria-label="切换深色模式 / Toggle dark mode">深色</button>
        <details class="appearance-panel" id="appearance-panel">
          <summary>
            <svg class="appearance-icon" viewBox="0 0 24 24" fill="none"
                 stroke="currentColor" stroke-width="2"
                 stroke-linecap="round" stroke-linejoin="round"
                 aria-hidden="true">
              <circle cx="12" cy="12" r="3"></circle>
              <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"></path>
            </svg>
            <span class="i18n-zh">外观</span>
            <span class="i18n-en">Appearance</span>
            <em>
              <span class="i18n-zh">默认收起 · 上传图仅保存在当前浏览器，报告与 CSV 不嵌入图片</span>
              <span class="i18n-en">Collapsed by default · uploads stay in this browser only and are not embedded in the report or CSV</span>
            </em>
          </summary>
      <div class="appearance-body">
        <div class="appearance-grid">
          <label class="appearance-field" for="greeting-mode">
            <span class="i18n-zh">问候模式</span>
            <span class="i18n-en">Greeting mode</span>
            <select id="greeting-mode" aria-label="问候模式 / Greeting mode">
              <option value="auto" data-text-zh="自动" data-text-en="Auto">自动</option>
              <option value="late-night" data-text-zh="凌晨" data-text-en="Late night">凌晨</option>
              <option value="morning" data-text-zh="早上" data-text-en="Morning">早上</option>
              <option value="noon" data-text-zh="中午" data-text-en="Noon">中午</option>
              <option value="evening" data-text-zh="晚上" data-text-en="Evening">晚上</option>
            </select>
          </label>
          <label class="appearance-field" for="user-name-input">
            <span class="i18n-zh">用户名</span>
            <span class="i18n-en">User name</span>
            <span class="appearance-name-row">
              <input id="user-name-input" type="text" maxlength="32"
                     autocomplete="nickname"
                     placeholder="仅当前浏览器 / This browser only">
              <button id="user-name-clear" type="button">
                <span class="i18n-zh">清除</span>
                <span class="i18n-en">Clear</span>
              </button>
            </span>
          </label>
          <label class="appearance-field" for="greeting-font">
            <span class="i18n-zh">问候语字体</span>
            <span class="i18n-en">Greeting font</span>
            <select id="greeting-font">
              <option value="default" data-text-zh="宋体（默认）" data-text-en="Songti (default)">宋体（默认）</option>
              <option value="kaiti" data-text-zh="楷体" data-text-en="Kaiti">楷体</option>
              <option value="fangsong" data-text-zh="仿宋" data-text-en="Fangsong">仿宋</option>
              <option value="zhongsong" data-text-zh="华文中宋" data-text-en="Zhongsong">华文中宋</option>
              <option value="yahei" data-text-zh="微软雅黑" data-text-en="Microsoft YaHei">微软雅黑</option>
              <option value="source-han" data-text-zh="思源宋体" data-text-en="Source Han Serif">思源宋体</option>
            </select>
          </label>
          <label class="appearance-field appearance-toggle" for="theme-follow-greeting">
            <span class="i18n-zh">主题跟随问候语</span>
            <span class="i18n-en">Match theme to greeting</span>
            <span class="appearance-toggle-row">
              <input id="theme-follow-greeting" type="checkbox">
              <span class="i18n-zh">启用联动（默认关闭）</span>
              <span class="i18n-en">Enable sync (off by default)</span>
            </span>
          </label>
          <label class="appearance-field appearance-toggle" for="achievement-enabled">
            <span class="i18n-zh">阅读成就</span>
            <span class="i18n-en">Reading achievements</span>
            <span class="appearance-toggle-row">
              <input id="achievement-enabled" type="checkbox" checked>
              <span class="i18n-zh">显示成就卡片</span>
              <span class="i18n-en">Show achievement cards</span>
            </span>
          </label>
          <label class="appearance-field" for="achievement-badge-style">
            <span class="i18n-zh">成就徽章样式</span>
            <span class="i18n-en">Achievement badge style</span>
            <select id="achievement-badge-style">
              <option value="seal" data-text-zh="印章" data-text-en="Seal">印章</option>
              <option value="ring" data-text-zh="圆环" data-text-en="Ring">圆环</option>
              <option value="minimal" data-text-zh="极简" data-text-en="Minimal">极简</option>
            </select>
          </label>
          <label class="appearance-field" for="shelf-title-style">
            <span class="i18n-zh">书架标题行样式</span>
            <span class="i18n-en">Shelf title-row style</span>
            <select id="shelf-title-style">
              <option value="a" data-text-zh="A · 标题旁并列分段（推荐）" data-text-en="A · Segmented beside title (recommended)">A · 标题旁并列分段（推荐）</option>
              <option value="b" data-text-zh="B · 标题旁紧凑图标分段" data-text-en="B · Compact icon segment">B · 标题旁紧凑图标分段</option>
              <option value="c" data-text-zh="C · 贴合式页签轨" data-text-en="C · Attached tab rail">C · 贴合式页签轨</option>
            </select>
          </label>
          <label class="appearance-field" for="shelf-position">
            <span class="i18n-zh">书架控件位置</span>
            <span class="i18n-en">Shelf control position</span>
            <select id="shelf-position">
              <option value="title" data-text-zh="靠近标题（默认）" data-text-en="Beside title (default)">靠近标题（默认）</option>
              <option value="actions" data-text-zh="右侧操作组" data-text-en="Right actions">右侧操作组</option>
              <option value="hidden" data-text-zh="隐藏" data-text-en="Hidden">隐藏</option>
            </select>
          </label>
          <label class="appearance-field manga-only" for="shelf-default-mode">
            <span class="i18n-zh">默认书架</span>
            <span class="i18n-en">Default shelf</span>
            <select id="shelf-default-mode">
              <option value="book" data-text-zh="图书" data-text-en="Books">图书</option>
              <option value="manga" data-text-zh="漫画" data-text-en="Manga">漫画</option>
            </select>
          </label>
          <label class="appearance-field appearance-toggle" for="show-token-labels">
            <span class="i18n-zh">显示分词方式标签</span>
            <span class="i18n-en">Show tokenizer labels</span>
            <span class="appearance-toggle-row">
              <input id="show-token-labels" type="checkbox">
              <span class="i18n-zh manga-only">默认关闭；漫画只显示 OCR 估算或漫画不适用</span>
              <span class="i18n-en manga-only">Off by default; manga shows only OCR estimate or not applicable</span>
              <span class="i18n-zh text-only-only">默认关闭</span>
              <span class="i18n-en text-only-only">Off by default</span>
            </span>
          </label>
          <label class="appearance-field" for="background-texture">
            <span class="i18n-zh">背景图案</span>
            <span class="i18n-en">Background pattern</span>
            <select id="background-texture">
              <option value="paper" data-text-zh="纸张" data-text-en="Paper">纸张</option>
              <option value="grid" data-text-zh="方格" data-text-en="Grid">方格</option>
              <option value="grain" data-text-zh="颗粒" data-text-en="Grain">颗粒</option>
              <option value="dots" data-text-zh="波点" data-text-en="Dots">波点</option>
              <option value="rings" data-text-zh="圆环" data-text-en="Rings">圆环</option>
              <option value="diagonal" data-text-zh="斜纹" data-text-en="Diagonal">斜纹</option>
              <option value="diamond" data-text-zh="菱形" data-text-en="Diamond">菱形</option>
              <option value="none" data-text-zh="关闭" data-text-en="None">关闭</option>
            </select>
          </label>
          <label class="appearance-field" for="background-motif">
            <span class="i18n-zh">装饰图案</span>
            <span class="i18n-en">Decorative motifs</span>
            <select id="background-motif">
              <option value="none" data-text-zh="关闭" data-text-en="None">关闭</option>
              <option value="books" data-text-zh="书本" data-text-en="Books">书本</option>
              <option value="leaves" data-text-zh="叶片" data-text-en="Leaves">叶片</option>
              <option value="stars" data-text-zh="星点" data-text-en="Stars">星点</option>
            </select>
          </label>
          <label class="appearance-field" for="background-source">
            <span class="i18n-zh">背景来源</span>
            <span class="i18n-en">Background source</span>
            <select id="background-source">
              <option value="none" data-text-zh="主题背景" data-text-en="Theme background">主题背景</option>
              <option value="upload" data-text-zh="本地上传图片" data-text-en="Local upload">本地上传图片</option>
              <option value="recent" data-text-zh="最近阅读封面（自动）" data-text-en="Latest reading cover (auto)">最近阅读封面（自动）</option>
              <option value="fixed" data-text-zh="固定书籍封面" data-text-en="Fixed book cover">固定书籍封面</option>
            </select>
          </label>
          <label class="appearance-field" id="fixed-cover-field" for="background-book" hidden>
            <span class="i18n-zh">固定封面</span>
            <span class="i18n-en">Fixed cover</span>
            <select id="background-book" aria-label="固定书籍封面 / Fixed book cover"></select>
          </label>
          <label class="appearance-field" for="background-upload">
            <span class="i18n-zh">上传本地图片</span>
            <span class="i18n-en">Upload local image</span>
            <input id="background-upload" type="file" accept="image/*">
          </label>
          <label class="appearance-field" for="background-blur">
            <span class="i18n-zh">模糊程度</span>
            <span class="i18n-en">Blur</span>
            <span class="range-row">
              <input id="background-blur" type="range" min="0" max="30" step="1" value="12">
              <output id="background-blur-value" for="background-blur">12 px</output>
            </span>
          </label>
          <label class="appearance-field" for="background-overlay">
            <span class="i18n-zh">明暗遮罩</span>
            <span class="i18n-en">Light or dark overlay</span>
            <span class="range-row">
              <input id="background-overlay" type="range" min="0" max="90" step="1" value="62">
              <output id="background-overlay-value" for="background-overlay">62%</output>
            </span>
          </label>
          <label class="appearance-field" for="background-opacity">
            <span class="i18n-zh">图片透明度</span>
            <span class="i18n-en">Image opacity</span>
            <span class="range-row">
              <input id="background-opacity" type="range" min="10" max="100" step="1" value="72">
              <output id="background-opacity-value" for="background-opacity">72%</output>
            </span>
          </label>
        </div>
        <div class="appearance-section-toggles"
             role="group"
             aria-label="报告区域显示 / Report section visibility">
          <p class="appearance-subtitle">
            <span class="i18n-zh">报告区域显示</span>
            <span class="i18n-en">Report sections</span>
          </p>
          <div class="appearance-grid">
            <label class="appearance-field appearance-toggle" for="section-period">
              <span class="i18n-zh">周期报告</span>
              <span class="i18n-en">Period reports</span>
              <span class="appearance-toggle-row">
                <input id="section-period" type="checkbox" data-section-toggle="period" checked>
                <span class="i18n-zh">显示</span>
                <span class="i18n-en">Visible</span>
              </span>
            </label>
            <label class="appearance-field appearance-toggle" for="section-annual">
              <span class="i18n-zh">年度阅读年鉴</span>
              <span class="i18n-en">Annual reading year</span>
              <span class="appearance-toggle-row">
                <input id="section-annual" type="checkbox" data-section-toggle="annual" checked>
                <span class="i18n-zh">显示</span>
                <span class="i18n-en">Visible</span>
              </span>
            </label>
            <label class="appearance-field appearance-toggle" for="section-trends">
              <span class="i18n-zh">阅读趋势</span>
              <span class="i18n-en">Reading trends</span>
              <span class="appearance-toggle-row">
                <input id="section-trends" type="checkbox" data-section-toggle="trends" checked>
                <span class="i18n-zh">显示</span>
                <span class="i18n-en">Visible</span>
              </span>
            </label>
            <label class="appearance-field appearance-toggle" for="section-heatmap">
              <span class="i18n-zh">最近 12 周阅读热力图</span>
              <span class="i18n-en">12-week reading heatmap</span>
              <span class="appearance-toggle-row">
                <input id="section-heatmap" type="checkbox" data-section-toggle="heatmap" checked>
                <span class="i18n-zh">显示</span>
                <span class="i18n-en">Visible</span>
              </span>
            </label>
            <label class="appearance-field appearance-toggle" for="minimal-mode">
              <span class="i18n-zh">一键极简模式</span>
              <span class="i18n-en">Minimal mode</span>
              <span class="appearance-toggle-row">
                <input id="minimal-mode" type="checkbox">
                <span class="i18n-zh">隐藏全部报告区域（默认关闭）</span>
                <span class="i18n-en">Hide all report sections (off by default)</span>
              </span>
            </label>
          </div>
        </div>
        <div class="appearance-actions">
          <button id="background-clear" type="button">
            <span class="i18n-zh">清除上传图</span><span class="i18n-en">Clear upload</span>
          </button>
          <button id="appearance-reset" type="button">
            <span class="i18n-zh">恢复默认外观</span><span class="i18n-en">Reset appearance</span>
          </button>
        </div>
        <p class="appearance-help">
          <span class="i18n-zh">本地上传图片仅保存在当前浏览器，报告与 CSV 不会嵌入图片。</span>
          <span class="i18n-en">Local uploads stay in this browser only and are never embedded in the report or CSV.</span>
        </p>
        <p class="appearance-help">
          <span class="i18n-zh">自动模式依次使用最后阅读时间、最近有变化的快照书籍、阅读时长排序；封面缺失或数据为空时回退主题背景。</span>
          <span class="i18n-en">Auto mode uses last-read time, then the latest changed snapshot book, then reading-time ranking; missing covers or empty data fall back to the theme background.</span>
        </p>
        <p class="appearance-help">
          <span class="i18n-zh">用户名只保存在当前浏览器，用于问候语和成就星星展示；不能跨设备识别，同一份 CSV 的成就基线仍共享。需要真正多用户时，应改为主机端多用户映射或独立输出目录。</span>
          <span class="i18n-en">The user name stays in this browser for greetings and achievement stars only. It cannot identify users across devices, and the CSV achievement baseline remains shared. Real multi-user support requires a host-side profile map or separate output directories.</span>
        </p>
        <p class="appearance-status" id="appearance-status" role="status" aria-live="polite"></p>
      </div>
        </details>
      </div>
    </header>

    <section class="summary">
      <div class="metric"><span><span class="i18n-zh">书籍</span><span class="i18n-en">Books</span></span><strong>__BOOKS__</strong></div>
      <div class="metric"><span><span class="i18n-zh">有阅读记录书籍</span><span class="i18n-en">Books with reading records</span></span><strong>__READING_RECORD_BOOKS__</strong></div>
      <div class="metric"><span><span class="i18n-zh">累计阅读</span><span class="i18n-en">Reading time</span></span><strong>__TIME__</strong></div>
      <div class="metric"><span><span class="i18n-zh">累计阅读词数</span><span class="i18n-en">Words read</span></span><strong>__WORDS__</strong></div>
      <div class="metric"><span><span class="i18n-zh">阅读速度</span><span class="i18n-en">Reading speed</span></span><strong>__WPM__</strong></div>
      <div class="metric"><span><span class="i18n-zh">平均阅读进度</span><span class="i18n-en">Average progress</span></span><strong>__PROGRESS__</strong></div>
      <div class="metric"><span><span class="i18n-zh">批注总数</span><span class="i18n-en">Annotations</span></span><strong>__ANNOTATIONS__</strong></div>
      <div class="metric"><span><span class="i18n-zh">真实封面</span><span class="i18n-en">Real covers</span></span><strong>__COVER_COUNT__ / __BOOKS__</strong></div>
    </section>

    <section class="library-panel" aria-label="书封书架 / Cover shelf">
      <div class="library-heading">
        <div class="shelf-title-cluster" id="shelf-title-cluster">
          <div class="shelf-title-copy">
            <h2><span class="i18n-zh">书封书架</span><span class="i18n-en">Cover shelf</span></h2>
            <p><span class="i18n-zh">按阅读时长排序</span><span class="i18n-en">Sorted by reading time</span> · __COVER_SUMMARY__</p>
          </div>
          <div class="shelf-mode-switch manga-only" id="shelf-mode-switch"
               role="group" aria-label="书架类型 / Shelf type">
            <button type="button" data-shelf-mode="book"
                    aria-pressed="true" aria-label="图书书架 / Book shelf">
              <svg class="shelf-mode-icon" viewBox="0 0 24 24"
                   aria-hidden="true" focusable="false">
                <path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"></path>
                <path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2Z"></path>
              </svg>
              <span class="shelf-mode-label-full i18n-zh">图书书架</span>
              <span class="shelf-mode-label-full i18n-en">Book shelf</span>
              <span class="shelf-mode-label-short i18n-zh">图书</span>
              <span class="shelf-mode-label-short i18n-en">Books</span>
            </button>
            <button type="button" data-shelf-mode="manga"
                    aria-pressed="false" aria-label="漫画书架 / Manga shelf">
              <svg class="shelf-mode-icon" viewBox="0 0 24 24"
                   aria-hidden="true" focusable="false">
                <rect x="3" y="4" width="18" height="16" rx="2"></rect>
                <path d="M8 4v16M16 4v16"></path>
              </svg>
              <span class="shelf-mode-label-full i18n-zh">漫画书架</span>
              <span class="shelf-mode-label-full i18n-en">Manga shelf</span>
              <span class="shelf-mode-label-short i18n-zh">漫画</span>
              <span class="shelf-mode-label-short i18n-en">Manga</span>
            </button>
          </div>
        </div>
        <div class="library-actions" id="library-actions">
          <div class="shelf-fallback-wrap manga-only" id="shelf-fallback-wrap">
            <button type="button" class="shelf-fallback-button"
                    id="shelf-fallback-button"
                    aria-haspopup="true" aria-expanded="false"
                    aria-label="切换书架 / Switch shelf">
              <svg viewBox="0 0 24 24" aria-hidden="true"
                   focusable="false">
                <path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"></path>
                <path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2Z"></path>
              </svg>
              <span class="i18n-zh">书架</span>
              <span class="i18n-en">Shelf</span>
            </button>
            <div class="shelf-fallback-menu" id="shelf-fallback-menu"
                 role="group" aria-label="书架类型 / Shelf type" hidden>
              <button type="button" data-shelf-mode="book"
                      aria-pressed="true">
                <span class="i18n-zh">图书书架</span>
                <span class="i18n-en">Book shelf</span>
              </button>
              <button type="button" data-shelf-mode="manga"
                      aria-pressed="false">
                <span class="i18n-zh">漫画书架</span>
                <span class="i18n-en">Manga shelf</span>
              </button>
            </div>
          </div>
          <div class="shelf-pager" aria-label="书架翻页 / Shelf pages">
            <button type="button" id="shelf-prev" aria-label="上一页">‹</button>
            <span id="shelf-page-status" aria-live="polite">1/1</span>
            <button type="button" id="shelf-next" aria-label="下一页">›</button>
          </div>
          <span class="cover-count-badge"><strong id="cover-count-number">__COVER_COUNT__</strong> <span class="i18n-zh">本封面</span><span class="i18n-en">covers</span></span>
        </div>
      </div>
      <div class="cover-grid" id="cover-grid">__COVER_CARDS__</div>
      <div class="empty compact" id="cover-empty" hidden>
        <span class="i18n-zh">没有匹配封面</span>
        <span class="i18n-en">No matching covers</span>
      </div>
    </section>

    <details class="achievement-panel" id="achievement-panel" open>
      <summary>
        <span class="i18n-zh">阅读成就</span>
        <span class="i18n-en">Reading achievements</span>
        <em id="achievement-summary">__ACHIEVEMENT_SUMMARY__</em>
      </summary>
      <div class="achievement-body">
        <button type="button" class="star-sky-entry" id="star-sky-open"
                aria-haspopup="dialog" aria-label="打开阅读星空">
          <svg viewBox="0 0 24 24" aria-hidden="true"
               focusable="false">
            <path d="M12 3.2c.6 0 1.14.34 1.4.88l1.62 3.28 3.62.53c1.24.18 1.73 1.69.84 2.56l-2.62 2.55.62 3.6c.21 1.23-1.08 2.17-2.18 1.59L12 16.5l-3.28 1.71c-1.1.58-2.39-.36-2.18-1.59l.62-3.6-2.62-2.55c-.89-.87-.4-2.38.84-2.56l3.62-.53 1.62-3.28A1.56 1.56 0 0 1 12 3.2Z"></path>
            <path d="M18.6 5.2v3.2M17 6.8h3.2" stroke-linecap="round"></path>
          </svg>
          <span class="i18n-zh">阅读星空</span>
          <span class="i18n-en">Reading sky</span>
        </button>
        <label class="achievement-toggle manga-only" for="manga-volumes-enabled">
          <input type="checkbox" id="manga-volumes-enabled">
          <span class="i18n-zh">显示漫画卷数成就</span>
          <span class="i18n-en">Show manga volume achievement</span>
        </label>
        <section class="achievement-grid" aria-label="阅读成就 / Reading achievements">
          __ACHIEVEMENTS__
        </section>
        <p class="achievement-criteria">
          <span class="i18n-zh">__ACHIEVEMENT_CRITERIA_ZH__</span>
          <span class="i18n-en">__ACHIEVEMENT_CRITERIA_EN__</span>
        </p>
      </div>
    </details>

    __PERIOD_REPORTS__

    __ANNUAL_REPORT__

    <section class="insights">__INSIGHTS__</section>

    <section class="panel trend-panel" id="reading-trends"
             __TREND_PANEL_ATTRS__
             aria-label="阅读趋势 / Reading trends">
      <div class="trend-heading">
        <div>
          <h2><span class="i18n-zh">阅读趋势</span><span class="i18n-en">Reading trends</span></h2>
          <p>
            <span class="i18n-zh">回看总时长、阅读词数和每日变化。</span>
            <span class="i18n-en">A quiet view of reading time, words, and daily change.</span>
          </p>
        </div>
      </div>
      <section class="trend-summary" aria-label="阅读趋势摘要 / Reading trend summary">
        __TREND_CARDS__
      </section>
      __TREND_CHART__
      <div class="trend-subgrid">
        <section class="trend-subpanel" aria-label="本次变化 / Latest changes">
          <h3><span class="i18n-zh">本次变化</span><span class="i18n-en">Latest changes</span></h3>
          <div class="movement-list">__TREND_MOVEMENTS__</div>
        </section>
        <section class="trend-subpanel" aria-label="最近 30 天每日变化 / Daily change in the last 30 days">
          <h3><span class="i18n-zh">最近 30 天每日变化</span><span class="i18n-en">Daily change in the last 30 days</span></h3>
          __TREND_DAILY__
        </section>
      </div>
    </section>

    <section class="panel heatmap-panel" aria-label="最近 12 周阅读热力图 / 12-week reading heatmap">
      <div class="heatmap-heading">
        <div>
          <h2><span class="i18n-zh">最近 12 周阅读热力图</span><span class="i18n-en">12-week reading heatmap</span></h2>
          <p><span class="i18n-zh">每日阅读时长</span><span class="i18n-en">Daily reading time</span></p>
        </div>
      </div>
      __TREND_HEATMAP__
    </section>

    <section class="panels">
      <article class="panel">
        <h2><span class="i18n-zh">阅读时长排行</span><span class="i18n-en">Reading-time ranking</span></h2>
        <div class="rank-list">__CHART__</div>
      </article>
      <article class="panel">
        <h2><span class="i18n-zh">词数进度</span><span class="i18n-en">Word-count progress</span></h2>
        <div class="progress-grid">__PROGRESS_CARDS__</div>
      </article>
    </section>

    <div class="toolbar">
      <div class="filters">
        <input id="search" type="search"
               placeholder="搜索书名、作者或来源文件"
               data-placeholder-zh="搜索书名、作者或来源文件"
               data-placeholder-en="Search title, author, or source file">
        <select id="language-filter" aria-label="语言筛选 / Language filter">
          <option value="" data-text-zh="全部语言" data-text-en="All languages">全部语言</option>
          <option value="en" data-text-zh="英文" data-text-en="English">英文</option>
          <option value="zh" data-text-zh="中文" data-text-en="Chinese">中文</option>
          <option value="ja" data-text-zh="日文" data-text-en="Japanese">日文</option>
        </select>
        <select id="count-filter" aria-label="书籍类型筛选 / Book type filter">
          <option value="" data-text-zh="全部书籍" data-text-en="All books">全部书籍</option>
          <option value="exact" data-text-zh="精确词数" data-text-en="Exact words">精确词数</option>
          <option value="estimate" data-text-zh="估算词数" data-text-en="Estimated words">估算词数</option>
          <option class="manga-only" value="manga" data-text-zh="漫画" data-text-en="Manga">漫画</option>
        </select>
        <select id="state-filter" aria-label="阅读状态筛选 / Reading state filter">
          <option value="" data-text-zh="全部状态" data-text-en="All states">全部状态</option>
          <option value="reading" data-text-zh="已开始阅读" data-text-en="Reading started">已开始阅读</option>
          <option value="unread" data-text-zh="未开始" data-text-en="Unread">未开始</option>
        </select>
      </div>
      <div class="toolbar-actions">
        <details class="column-picker" id="column-picker">
          <summary><span class="i18n-zh">列显示</span><span class="i18n-en">Columns</span></summary>
          <div class="column-menu" id="column-options">
            <label><input type="checkbox" data-column-toggle="rank" checked><span class="i18n-zh">序号</span><span class="i18n-en">Rank</span></label>
            <label><input type="checkbox" data-column-toggle="book" checked><span class="i18n-zh">书名 / 作者</span><span class="i18n-en">Title / Author</span></label>
            <label><input type="checkbox" data-column-toggle="duration" checked><span class="i18n-zh">阅读时长</span><span class="i18n-en">Reading time</span></label>
            <label><input type="checkbox" data-column-toggle="words" checked><span class="i18n-zh">有效已读 / 总词数</span><span class="i18n-en">Effective read / total words</span></label>
            <label><input type="checkbox" data-column-toggle="wpm" checked><span class="i18n-zh">有效 WPM</span><span class="i18n-en">Effective WPM</span></label>
            <label><input type="checkbox" data-column-toggle="kindle-progress" checked><span class="i18n-zh">Kindle 进度</span><span class="i18n-en">Kindle progress</span></label>
            <label><input type="checkbox" data-column-toggle="word-progress" checked><span class="i18n-zh">词数进度</span><span class="i18n-en">Word progress</span></label>
            <label><input type="checkbox" data-column-toggle="tags" checked><span class="i18n-zh">语言 / 来源</span><span class="i18n-en">Language / Source</span></label>
            <label><input type="checkbox" data-column-toggle="annotations" checked><span class="i18n-zh">批注</span><span class="i18n-en">Annotations</span></label>
            <label><input type="checkbox" data-column-toggle="last-read" checked><span class="i18n-zh">最后阅读</span><span class="i18n-en">Last read</span></label>
            <label><input type="checkbox" data-column-toggle="delta" checked><span class="i18n-zh">本次增量</span><span class="i18n-en">Latest delta</span></label>
          </div>
        </details>
        <a class="export-button" id="xlsx-export"
           href="__XLSX_NAME__" download>
          <span class="i18n-zh">导出 XLSX</span>
          <span class="i18n-en">Export XLSX</span>
        </a>
        <button type="button" id="share-export">
          <span class="i18n-zh">导出分享图</span>
          <span class="i18n-en">Export share image</span>
        </button>
        <button type="button" id="print-report">
          <span class="i18n-zh">打印</span>
          <span class="i18n-en">Print</span>
        </button>
      </div>
    </div>

    <div class="table-wrap">
      <table id="data-table">
        <thead>
          <tr>
            <th data-type="number" data-column="rank">#</th>
            <th data-type="text" data-column="book"><span class="i18n-zh">书名 / 作者</span><span class="i18n-en">Title / Author</span></th>
            <th data-type="number" data-column="duration"><span class="i18n-zh">阅读时长</span><span class="i18n-en">Reading time</span></th>
            <th data-type="number" data-column="words"><span class="i18n-zh">有效已读 / 总词数</span><span class="i18n-en">Effective read / total words</span></th>
            <th data-type="number" data-column="wpm"><span class="i18n-zh">有效 WPM</span><span class="i18n-en">Effective WPM</span></th>
            <th data-type="number" data-column="kindle-progress"><span class="i18n-zh">Kindle 进度</span><span class="i18n-en">Kindle progress</span></th>
            <th data-type="number" data-column="word-progress"><span class="i18n-zh">词数进度</span><span class="i18n-en">Word progress</span></th>
            <th data-type="text" data-column="tags"><span class="i18n-zh">语言 / 来源</span><span class="i18n-en">Language / Source</span></th>
            <th data-type="number" data-column="annotations"><span class="i18n-zh">批注</span><span class="i18n-en">Annotations</span></th>
            <th data-type="text" data-column="last-read"><span class="i18n-zh">最后阅读</span><span class="i18n-en">Last read</span></th>
            <th data-type="number" data-column="delta"><span class="i18n-zh">本次增量</span><span class="i18n-en">Latest delta</span></th>
          </tr>
        </thead>
        <tbody>__ROWS__</tbody>
      </table>
      <div class="empty" id="empty" hidden>
        <span class="i18n-zh">没有匹配记录</span>
        <span class="i18n-en">No matching records</span>
      </div>
    </div>
    <footer>
      <span class="i18n-zh">报告文件：__CSV__ · XLSX：__XLSX__ · 封面缓存：__COVERS_DIR__ · 阅读历史：__HISTORY__。</span>
      <span class="i18n-en">Report files: __CSV__ · XLSX: __XLSX__ · Cover cache: __COVERS_DIR__ · Reading history: __HISTORY__.</span>
    </footer>
  </main>
  <dialog class="local-name-dialog" id="user-name-dialog"
          role="dialog" aria-modal="false"
          aria-labelledby="user-name-dialog-title">
    <div class="local-name-body">
      <h2 id="user-name-dialog-title">
        <span class="i18n-zh">设置本机用户名</span>
        <span class="i18n-en">Set a local user name</span>
      </h2>
      <p>
        <span class="i18n-zh">用于问候语和成就星星展示，只保存在当前浏览器，不会写入 CSV 或 Kindle。可以跳过，也可以稍后在外观面板修改或清除。</span>
        <span class="i18n-en">Used for greetings and achievement stars only. It stays in this browser and is never written to CSV or Kindle. Skip now, or edit or clear it later in the Appearance panel.</span>
      </p>
      <input class="local-name-input" id="user-name-dialog-input"
             type="text" maxlength="32" autocomplete="nickname"
             placeholder="例如：小林 / e.g. Lin">
      <div class="local-name-actions">
        <button id="user-name-skip" type="button">
          <span class="i18n-zh">跳过</span>
          <span class="i18n-en">Skip</span>
        </button>
        <button id="user-name-save" type="button">
          <span class="i18n-zh">保存</span>
          <span class="i18n-en">Save</span>
        </button>
      </div>
    </div>
  </dialog>
  <dialog class="share-dialog" id="share-dialog"
          aria-labelledby="share-dialog-title">
    <div class="share-dialog-body">
      <div class="share-dialog-heading">
        <div>
          <h2 id="share-dialog-title">
            <span class="i18n-zh">导出阅读分享图</span>
            <span class="i18n-en">Export reading share image</span>
          </h2>
          <p>
            <span class="i18n-zh">生成适合手机查看的 PNG，可保存后分享到朋友圈或 QQ 动态。</span>
            <span class="i18n-en">Create mobile-friendly PNG files for social sharing.</span>
          </p>
        </div>
        <button type="button" id="share-close" class="share-close"
                aria-label="关闭分享图面板">×</button>
      </div>
      <label class="share-field" for="share-format">
        <span class="i18n-zh">图片形式</span>
        <span class="i18n-en">Image format</span>
        <select id="share-format">
          <option value="long" data-text-zh="单张长图" data-text-en="Single long image">单张长图</option>
          <option value="set" data-text-zh="组图（2 张）" data-text-en="Image set (2 files)">组图（2 张）</option>
        </select>
      </label>
      <div class="share-preview-note">
        <strong><span class="i18n-zh">会包含：</span><span class="i18n-en">Includes: </span></strong>
        <span class="i18n-zh">阅读总览、四轨成就、最近阅读书目与阅读星空摘要。</span>
        <span class="i18n-en">Reading overview, four achievements, recent books, and star sky summary.</span>
      </div>
      <div class="share-actions">
        <button type="button" id="share-generate">
          <span class="i18n-zh">生成并下载</span>
          <span class="i18n-en">Generate and download</span>
        </button>
        <span id="share-status" role="status" aria-live="polite"></span>
      </div>
    </div>
  </dialog>
  __STAR_SKY__
  <script id="trend-data" type="application/json">__TREND_DATA__</script>
  <script>
    const search = document.getElementById("search");
    const languageFilter = document.getElementById("language-filter");
    const countFilter = document.getElementById("count-filter");
    const stateFilter = document.getElementById("state-filter");
    const table = document.getElementById("data-table");
    const tbody = table.querySelector("tbody");
    const empty = document.getElementById("empty");
    const coverGrid = document.getElementById("cover-grid");
    const coverEmpty = document.getElementById("cover-empty");
    const coverCountNumber = document.getElementById("cover-count-number");
    const shelfModeButtons = Array.from(
      document.querySelectorAll("[data-shelf-mode]")
    );
    const shelfPrev = document.getElementById("shelf-prev");
    const shelfNext = document.getElementById("shelf-next");
    const shelfPageStatus = document.getElementById("shelf-page-status");
    const shelfModeSwitch = document.getElementById("shelf-mode-switch");
    const shelfTitleCluster = document.getElementById("shelf-title-cluster");
    const libraryActions = document.getElementById("library-actions");
    const shelfFallbackWrap = document.getElementById("shelf-fallback-wrap");
    const shelfFallbackButton = document.getElementById(
      "shelf-fallback-button"
    );
    const shelfFallbackMenu = document.getElementById("shelf-fallback-menu");
    const themeToggle = document.getElementById("theme-toggle");
    const languageToggle = document.getElementById("language-toggle");
    const columnPicker = document.getElementById("column-picker");
    const achievementPanel = document.getElementById("achievement-panel");
    const achievementSummary = document.getElementById(
      "achievement-summary"
    );
    const starSkyOpen = document.getElementById("star-sky-open");
    const starSkyDialog = document.getElementById("star-sky-dialog");
    const starSkyClose = document.getElementById("star-sky-close");
    const starSkyStars = Array.from(
      document.querySelectorAll(".star-sky-star")
    );
    const starSkyDetail = document.getElementById("star-sky-detail");
    const starSkyCount = document.getElementById("star-sky-count");
    const shareExport = document.getElementById("share-export");
    const printReport = document.getElementById("print-report");
    const shareDialog = document.getElementById("share-dialog");
    const shareClose = document.getElementById("share-close");
    const shareFormat = document.getElementById("share-format");
    const shareGenerate = document.getElementById("share-generate");
    const shareStatus = document.getElementById("share-status");
    const achievementCards = Array.from(
      document.querySelectorAll("[data-achievement-key]")
    );
    const periodPanel = document.getElementById("period-report-panel");
    const periodTabs = Array.from(
      document.querySelectorAll("[data-period-tab]")
    );
    const periodViews = Array.from(
      document.querySelectorAll("[data-period-panel]")
    );
    const periodBarCharts = Array.from(
      document.querySelectorAll("[data-period-trend-chart='bars']")
    );
    const annualReportPanel = document.getElementById(
      "annual-report-panel"
    );
    const readingTrendsPanel = document.getElementById("reading-trends");
    const heatmapPanel = document.querySelector(".heatmap-panel");
    const sectionToggles = Array.from(
      document.querySelectorAll("[data-section-toggle]")
    );
    const minimalModeToggle = document.getElementById("minimal-mode");
    const greetingMessageZh = document.getElementById("greeting-zh");
    const greetingMessageEn = document.getElementById("greeting-en");
    const greetingMode = document.getElementById("greeting-mode");
    const userNameDialog = document.getElementById("user-name-dialog");
    const userNameDialogInput = document.getElementById(
      "user-name-dialog-input"
    );
    const userNameSave = document.getElementById("user-name-save");
    const userNameSkip = document.getElementById("user-name-skip");
    const userNameInput = document.getElementById("user-name-input");
    const userNameClear = document.getElementById("user-name-clear");
    const trendMetricButtons = Array.from(
      document.querySelectorAll("[data-trend-metric-button]")
    );
    const trendSeriesViews = Array.from(
      document.querySelectorAll("[data-trend-series]")
    );
    const trendDetails = document.getElementById("trend-snapshot-details");
    const trendTooltip = document.getElementById("trend-tooltip");
    const themePresetSelect = document.getElementById("theme-preset");
    const themeFollowGreeting = document.getElementById(
      "theme-follow-greeting"
    );
    const appearancePanel = document.getElementById("appearance-panel");
    const greetingFontSelect = document.getElementById("greeting-font");
    const achievementBadgeStyleSelect = document.getElementById(
      "achievement-badge-style"
    );
    const shelfTitleStyleSelect = document.getElementById(
      "shelf-title-style"
    );
    const shelfPositionSelect = document.getElementById("shelf-position");
    const shelfDefaultModeSelect = document.getElementById(
      "shelf-default-mode"
    );
    const showTokenLabelsToggle = document.getElementById(
      "show-token-labels"
    );
    const achievementEnabledToggle = document.getElementById(
      "achievement-enabled"
    );
    const mangaVolumesToggle = document.getElementById(
      "manga-volumes-enabled"
    );
    const backgroundTextureSelect = document.getElementById(
      "background-texture"
    );
    const backgroundMotifSelect = document.getElementById(
      "background-motif"
    );
    const backgroundSource = document.getElementById("background-source");
    const fixedCoverField = document.getElementById("fixed-cover-field");
    const backgroundBook = document.getElementById("background-book");
    const backgroundUpload = document.getElementById("background-upload");
    const backgroundBlur = document.getElementById("background-blur");
    const backgroundOverlay = document.getElementById("background-overlay");
    const backgroundOpacity = document.getElementById("background-opacity");
    const backgroundBlurValue = document.getElementById(
      "background-blur-value"
    );
    const backgroundOverlayValue = document.getElementById(
      "background-overlay-value"
    );
    const backgroundOpacityValue = document.getElementById(
      "background-opacity-value"
    );
    const backgroundClear = document.getElementById("background-clear");
    const appearanceReset = document.getElementById("appearance-reset");
    const appearanceStatus = document.getElementById("appearance-status");
    const columnToggles = Array.from(
      document.querySelectorAll("[data-column-toggle]")
    );
    const THEME_PRESETS = new Set([
      "forest", "ocean", "sunset", "graphite",
      "rose", "lavender", "sky", "sepia",
    ]);
    const ACHIEVEMENT_BADGE_STYLES = new Set([
      "seal", "ring", "minimal",
    ]);
    const SHELF_TITLE_STYLES = new Set(["a", "b", "c"]);
    const SHELF_POSITIONS = new Set(["title", "actions", "hidden"]);
    const SHELF_DEFAULT_MODES = new Set(["book", "manga"]);
    const BACKGROUND_TEXTURES = new Set([
      "none", "paper", "grid", "grain",
      "dots", "rings", "diagonal", "diamond",
    ]);
    const BACKGROUND_MOTIFS = new Set([
      "none", "books", "leaves", "stars",
    ]);
    const BACKGROUND_SOURCES = new Set([
      "none", "upload", "recent", "fixed",
    ]);
    const BACKGROUND_SETTINGS_KEY = "kindle-dashboard-background-settings";
    const BACKGROUND_IMAGE_KEY = "kindle-dashboard-background-image";
    const GREETING_MODE_KEY = "kindle-dashboard-greeting-mode";
    const GREETING_FONT_KEY = "kindle-dashboard-greeting-font";
    const USER_NAME_KEY = "kindle-dashboard-user-name";
    const ACHIEVEMENT_STATE_PREFIX =
      "kindle-dashboard-achievement-state::";
    const ACHIEVEMENT_PANEL_KEY = "kindle-dashboard-achievement-panel";
    const ACHIEVEMENT_BADGE_STYLE_KEY =
      "kindle-dashboard-achievement-badge-style";
    const ACHIEVEMENT_ENABLED_KEY =
      "kindle-dashboard-achievement-enabled";
    const MANGA_VOLUME_ACHIEVEMENT_KEY =
      "kindle-dashboard-manga-volume-achievement";
    const SHELF_TITLE_STYLE_KEY =
      "kindle-dashboard-shelf-title-style";
    const SHELF_POSITION_KEY = "kindle-dashboard-shelf-position";
    const SHELF_DEFAULT_MODE_KEY =
      "kindle-dashboard-shelf-default-mode";
    const SHOW_TOKEN_LABELS_KEY =
      "kindle-dashboard-show-token-labels";
    const BACKGROUND_TEXTURE_KEY = "kindle-dashboard-background-texture";
    const BACKGROUND_MOTIF_KEY = "kindle-dashboard-background-motif";
    const PERIOD_VIEW_KEY = "kindle-dashboard-period-view";
    const PERIOD_PANEL_KEY = "kindle-dashboard-period-panel";
    const ANNUAL_REPORT_PANEL_KEY =
      "kindle-dashboard-annual-report-panel";
    const TREND_METRIC_KEY = "kindle-dashboard-trend-metric";
    const TREND_DETAILS_KEY = "kindle-dashboard-trend-details";
    const THEME_PRESET_KEY = "kindle-dashboard-theme-preset";
    const THEME_FOLLOW_GREETING_KEY =
      "kindle-dashboard-theme-follow-greeting";
    const SECTION_VISIBILITY_KEY =
      "kindle-dashboard-section-visibility";
    const MINIMAL_MODE_KEY = "kindle-dashboard-minimal-mode";
    const TREND_METRICS = new Set(["time", "words", "daily"]);
    const GREETING_MODES = new Set([
      "auto", "late-night", "morning", "noon", "evening",
    ]);
    const GREETING_FONTS = Object.freeze({
      default: '"SimSun", "Songti SC", "Noto Serif CJK SC", serif',
      kaiti: '"STKaiti", "KaiTi", "Kaiti SC", "Noto Serif CJK SC", serif',
      fangsong:
        '"FangSong", "FangSong_GB2312", "STFangsong", ' +
        '"Noto Serif CJK SC", serif',
      zhongsong:
        '"STZhongsong", "SimSun", "Songti SC", ' +
        '"Noto Serif CJK SC", serif',
      yahei:
        '"Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", sans-serif',
      "source-han":
        '"Source Han Serif SC", "Noto Serif CJK SC", "Songti SC", serif',
    });
    const PERIOD_VIEWS = new Set(["week", "month", "year"]);
    const GREETING_THEME_PRESETS = Object.freeze({
      "late-night": "lavender",
      morning: "sepia",
      day: "ocean",
      dusk: "sunset",
      night: "rose",
    });
    const DEFAULT_BACKGROUND_SETTINGS = Object.freeze({
      source: "none",
      blur: 12,
      overlay: 62,
      opacity: 72,
      bookId: "",
    });
    const MAX_BACKGROUND_IMAGE_LENGTH = 3200000;
    let backgroundSettings = {...DEFAULT_BACKGROUND_SETTINGS};
    let backgroundImageDataUrl = "";
    let themeFollowsGreeting = false;
    let manualThemePreset = "forest";
    let currentUserName = "";
    const reportSections = Object.freeze({
      period: periodPanel,
      annual: annualReportPanel,
      trends: readingTrendsPanel,
      heatmap: heatmapPanel,
    });
    const SECTION_KEYS = [
      "period", "annual", "trends", "heatmap",
    ];
    const DEFAULT_SECTION_VISIBILITY = Object.freeze({
      period: true,
      annual: true,
      trends: true,
      heatmap: true,
    });
    let sectionVisibility = {...DEFAULT_SECTION_VISIBILITY};
    let minimalMode = false;
    let shelfTitleStyle = "a";
    let shelfPosition = "title";
    let shelfDefaultMode = "book";
    let showTokenLabels = false;
    let shelfMode = "book";
    let shelfPage = 1;
    let mangaVolumesEnabled = false;

    function readPreference(key) {
      try {
        return window.localStorage.getItem(key);
      } catch (error) {
        return null;
      }
    }

    function writePreference(key, value) {
      try {
        window.localStorage.setItem(key, value);
        return true;
      } catch (error) {
        // Local files can run with storage disabled; the current session
        // still applies the setting.
        return false;
      }
    }

    function removePreference(key) {
      try {
        window.localStorage.removeItem(key);
      } catch (error) {
        // The current session still clears the value.
      }
    }

    function normalizeGreetingMode(value) {
      const normalized = String(value || "").toLowerCase();
      return GREETING_MODES.has(normalized) ? normalized : "auto";
    }

    function normalizeGreetingFont(value) {
      const normalized = String(value || "").toLowerCase();
      return Object.prototype.hasOwnProperty.call(GREETING_FONTS, normalized)
        ? normalized
        : "default";
    }

    function normalizeUserName(value) {
      return String(value || "")
        .replace(/\\s+/g, " ")
        .trim()
        .slice(0, 32);
    }

    function setGreetingFont(value, persist = true) {
      const normalized = normalizeGreetingFont(value);
      document.documentElement.dataset.greetingFont = normalized;
      document.documentElement.style.setProperty(
        "--greeting-font",
        GREETING_FONTS[normalized]
      );
      if (greetingFontSelect) {
        greetingFontSelect.value = normalized;
      }
      if (persist) {
        writePreference(GREETING_FONT_KEY, normalized);
      }
    }

    function browserGreetingPeriod(now = new Date()) {
      const hour = now.getHours();
      if (hour < 5) return "late-night";
      if (hour < 11) return "morning";
      if (hour < 14) return "noon";
      return "evening";
    }

    const GREETING_MODE_TONES = Object.freeze({
      "late-night": "late-night",
      morning: "morning",
      noon: "day",
      evening: "night",
    });

    function browserGreetingTone(now = new Date()) {
      const hour = now.getHours();
      if (hour < 5) return "late-night";
      if (hour < 10) return "morning";
      if (hour < 17) return "day";
      if (hour < 20) return "dusk";
      return "night";
    }

    function greetingCopy(period, name = "") {
      const displayName = normalizeUserName(name);
      const copy = {
        "late-night": {
          zh: "凌晨好！安静的阅读时光，正好适合翻开一页！(｡･ω･｡)",
          en: "Late night! One quiet page before sleep! (｡･ω･｡)",
        },
        morning: {
          zh: "早上好！新的一天，从几页阅读开始！(＾▽＾)",
          en: "Good morning! Start the day with a few pages! (＾▽＾)",
        },
        noon: {
          zh: "中午好！忙里偷闲，读几页吧！(๑˃̵ᴗ˂̵)",
          en: "Good afternoon! Take a small reading break! (๑˃̵ᴗ˂̵)",
        },
        evening: {
          zh: "晚上好！让今天在阅读中慢慢收尾！(｡･ω･｡)",
          en: "Good evening! Let reading bring the day to a close! (｡･ω･｡)",
        },
      }[period] || {
        zh: "欢迎回来！愿你有一段安静的阅读时光！(｡･ω･｡)",
        en: "Welcome back! Enjoy a quiet reading moment! (｡･ω･｡)",
      };
      if (!displayName) return copy;
      const named = {
        "late-night": {
          zh: `凌晨好，${displayName}！安静的阅读时光，正好适合翻开一页！(｡･ω･｡)`,
          en: `Late night, ${displayName}! One quiet page before sleep! (｡･ω･｡)`,
        },
        morning: {
          zh: `早上好，${displayName}！新的一天，从几页阅读开始！(＾▽＾)`,
          en: `Good morning, ${displayName}! Start the day with a few pages! (＾▽＾)`,
        },
        noon: {
          zh: `中午好，${displayName}！忙里偷闲，读几页吧！(๑˃̵ᴗ˂̵)`,
          en: `Good afternoon, ${displayName}! Take a small reading break! (๑˃̵ᴗ˂̵)`,
        },
        evening: {
          zh: `晚上好，${displayName}！让今天在阅读中慢慢收尾！(｡･ω･｡)`,
          en: `Good evening, ${displayName}! Let reading bring the day to a close! (｡･ω･｡)`,
        },
      }[period] || {
        zh: `欢迎回来，${displayName}！愿你有一段安静的阅读时光！(｡･ω･｡)`,
        en: `Welcome back, ${displayName}! Enjoy a quiet reading moment! (｡･ω･｡)`,
      };
      return named;
    }

    function splitGreetingEmoticon(text) {
      const value = String(text || "");
      const match = value.match(/^([\\s\\S]*)(\\s*[（(][^()（）]*[）)])$/);
      if (!match) return { text: value, emoticon: "" };
      return { text: match[1], emoticon: match[2].trimStart() };
    }

    function setGreetingText(element, text) {
      if (!element) return;
      const parts = splitGreetingEmoticon(text);
      element.textContent = "";
      if (!parts.emoticon) {
        element.appendChild(document.createTextNode(parts.text));
        return;
      }
      const spaced = /\\s$/.test(parts.text);
      let body = parts.text.replace(/\\s+$/, "");
      const emoticon = document.createElement("span");
      emoticon.className = "greeting-emoticon";
      if (spaced) {
        // Keep the final word with the emoticon so the emoticon never lands
        // on a line of its own.
        const tailMatch = body.match(/(\\S+)$/);
        const tail = tailMatch ? tailMatch[1] : "";
        if (tail) {
          emoticon.textContent = tail + "\\u00a0" + parts.emoticon;
          body = body.slice(0, body.length - tail.length);
        } else {
          emoticon.textContent = parts.emoticon;
          body += "\\u00a0";
        }
        element.appendChild(document.createTextNode(body));
      } else {
        // Keep the character before the emoticon inside the nowrap span so
        // the emoticon never wraps onto a line by itself.
        emoticon.textContent = body.slice(-1) + parts.emoticon;
        body = body.slice(0, -1);
        element.appendChild(document.createTextNode(body));
      }
      element.appendChild(emoticon);
    }

    function updateGreeting(persist = false) {
      if (!greetingMode || !greetingMessageZh || !greetingMessageEn) return;
      const mode = normalizeGreetingMode(greetingMode.value);
      greetingMode.value = mode;
      const period = mode === "auto"
        ? browserGreetingPeriod()
        : mode;
      const tone = mode === "auto"
        ? browserGreetingTone()
        : (GREETING_MODE_TONES[mode] || "day");
      const copy = greetingCopy(period, currentUserName);
      setGreetingText(greetingMessageZh, copy.zh);
      setGreetingText(greetingMessageEn, copy.en);
      document.documentElement.dataset.greetingMode = mode;
      document.documentElement.dataset.greetingPeriod = period;
      document.documentElement.dataset.greetingTone = tone;
      document.documentElement.dataset.greetingSource =
        mode === "auto" ? "browser-local-time" : "manual";
      applyGreetingThemeLink();
      if (persist) writePreference(GREETING_MODE_KEY, mode);
    }

    function restoreGreeting() {
      if (!greetingMode) return;
      greetingMode.value = normalizeGreetingMode(
        readPreference(GREETING_MODE_KEY)
      );
      updateGreeting(false);
    }

    function achievementStateKey(name) {
      return ACHIEVEMENT_STATE_PREFIX +
        (normalizeUserName(name) || "__local__");
    }

    function readAchievementState(name) {
      const stored = readPreference(achievementStateKey(name));
      if (!stored) return {};
      try {
        const parsed = JSON.parse(stored);
        return parsed && typeof parsed === "object" ? parsed : {};
      } catch (error) {
        return {};
      }
    }

    function writeAchievementState(name, state) {
      writePreference(achievementStateKey(name), JSON.stringify(state));
    }

    function applyAchievementLevels(persist = true) {
      const state = readAchievementState(currentUserName);
      achievementCards.forEach(card => {
        const key = card.dataset.achievementKey || "";
        const baseLevel = Math.max(
          0,
          Number(card.dataset.achievementBaseLevel) || 0
        );
        const maxLevel = Math.max(
          baseLevel,
          Number(card.dataset.achievementMaxLevel) || 0
        );
        const storedLevel = Math.max(0, Number(state[key]) || 0);
        const level = Math.min(maxLevel, Math.max(baseLevel, storedLevel));
        card.dataset.achievementLevel = String(level);
        if (level > storedLevel) state[key] = level;
        const baseState = card.dataset.achievementState || "in_progress";
        card.classList.toggle(
          "achievement-unlocked",
          level >= maxLevel && maxLevel > 0
        );
        card.classList.toggle(
          "achievement-in_progress",
          baseState !== "insufficient" && level < maxLevel
        );
        card.classList.toggle(
          "achievement-insufficient",
          baseState === "insufficient"
        );
        const displayState = baseState === "insufficient"
          ? "insufficient"
          : (level >= maxLevel && maxLevel > 0 ? "unlocked" : "in_progress");
        const stateElement = card.querySelector(".achievement-state");
        if (stateElement) {
          stateElement.innerHTML = displayState === "unlocked"
            ? '<span class="i18n-zh">已解锁</span><span class="i18n-en">Unlocked</span>'
            : displayState === "in_progress"
              ? '<span class="i18n-zh">进行中</span><span class="i18n-en">In progress</span>'
              : '<span class="i18n-zh">数据不足</span><span class="i18n-en">Not enough data</span>';
        }
        card.querySelectorAll(".achievement-star").forEach(star => {
          const starLevel = Number(star.dataset.starLevel) || 0;
          const lit = starLevel <= level;
          const next = !lit && starLevel === level + 1;
          star.classList.toggle("star-lit", lit);
          star.classList.toggle(
            "star-next",
            next && baseState !== "insufficient"
          );
          star.classList.toggle(
            "star-next-pending",
            next && baseState === "insufficient"
          );
          star.classList.toggle("star-future", !lit && !next);
          const starState = lit ? "lit" : (next ? "next" : "future");
          star.dataset.starState = starState;
          const zhState = star.querySelector(".sr-only.i18n-zh");
          const enState = star.querySelector(".sr-only.i18n-en");
          if (zhState) zhState.textContent = lit ? "已点亮" : "未点亮";
          if (enState) enState.textContent = lit ? "Unlocked" : "Locked";
          const english =
            document.documentElement.dataset.uiLanguage === "en";
          const label = english
            ? "Star " + starLevel + ": " + (lit ? "unlocked" : "locked")
            : "星 " + starLevel + (lit ? "：已点亮" : "：未点亮");
          star.setAttribute("aria-label", label);
        });
      });
      if (persist) writeAchievementState(currentUserName, state);
      syncStarSky();
      updateAchievementRootAttributes();
    }

    function starAchievementNode(star) {
      const metric = star.dataset.skyMetric || "";
      const level = star.dataset.skyLevel || "";
      if (!metric || !level) return null;
      const card = achievementCards.find(
        candidate => candidate.dataset.achievementKey === metric
      );
      return card
        ? card.querySelector('[data-star-level="' + level + '"]')
        : null;
    }

    function syncStarSky() {
      if (!starSkyStars.length) return;
      let unlocked = 0;
      starSkyStars.forEach(star => {
        const node = starAchievementNode(star);
        const lit = !!node && node.classList.contains("star-lit");
        star.classList.toggle("star-sky-star-lit", lit);
        star.classList.toggle("star-sky-star-locked", !lit);
        star.dataset.skyLit = lit ? "true" : "false";
        if (node) {
          star.dataset.skyUnlockedAt = node.dataset.starUnlockedAt || "";
          star.dataset.skyApproximate = node.dataset.starApproximate || "false";
          star.dataset.skyObserved = node.dataset.starObserved || "";
        }
        if (lit) unlocked += 1;
      });
      if (starSkyCount) {
        const total = starSkyStars.length;
        starSkyCount.textContent = unlocked + "/" + total;
      }
    }

    function showStarSkyDetail(star) {
      if (!starSkyDetail || !star || star.dataset.skyLit !== "true") return;
      const approximate = star.dataset.skyApproximate === "true";
      const time = star.dataset.skyUnlockedAt
        ? (approximate
            ? star.dataset.skyUnlockedAt + " · 根据历史快照推定"
            : star.dataset.skyUnlockedAt)
        : "获得时间未知";
      starSkyDetail.innerHTML =
        "<strong>" + (star.dataset.skyName || "已点亮的星") + "</strong>" +
        "<span>获得时间：" + time + "</span>" +
        "<span>所属轨道：" + (star.dataset.skyTrack || "阅读") + "</span>" +
        "<span>达成门槛：" + (star.dataset.skyThreshold || "未知") +
        " · 当时统计：" + (star.dataset.skyObserved || "未知") + "</span>";
    }

    function openStarSky() {
      if (!starSkyDialog) return;
      syncStarSky();
      if (typeof starSkyDialog.showModal === "function") {
        starSkyDialog.showModal();
      } else {
        starSkyDialog.hidden = false;
        starSkyDialog.setAttribute("open", "");
      }
      starSkyClose?.focus();
    }

    function closeStarSky() {
      if (!starSkyDialog) return;
      if (typeof starSkyDialog.close === "function") {
        starSkyDialog.close();
      } else {
        starSkyDialog.hidden = true;
        starSkyDialog.removeAttribute("open");
      }
      starSkyOpen?.focus();
    }

    function currentShareLanguage() {
      return document.documentElement.dataset.uiLanguage === "en"
        ? "en"
        : "zh";
    }

    function shareText(element) {
      if (!element) return "";
      const language = currentShareLanguage();
      const preferred = element.querySelector(
        language === "en" ? ".i18n-en" : ".i18n-zh"
      );
      return (preferred || element).textContent.replace(/\\s+/g, " ").trim();
    }

    function shareMetricRows() {
      return Array.from(document.querySelectorAll(".summary .metric"))
        .slice(0, 6)
        .map(metric => ({
          label: shareText(metric.querySelector("span")),
          value: shareText(metric.querySelector("strong")),
        }))
        .filter(item => item.label && item.value);
    }

    function shareAchievementRows() {
      return achievementCards.map(card => {
        const label = shareText(card.querySelector(".achievement-heading h3"));
        const stars = Array.from(card.querySelectorAll(".achievement-star"));
        const lit = stars.filter(star => star.classList.contains("star-lit")).length;
        const encouragement = shareText(
          card.querySelector(".achievement-encouragement")
        );
        return { label, lit, total: stars.length, encouragement };
      }).filter(item => item.label);
    }

    function shareBookRows(limit = 6) {
      return Array.from(
        coverGrid.querySelectorAll(".book-cover-card[data-book-id]")
      ).slice(0, limit).map(card => {
        const stats = Array.from(card.querySelectorAll(".cover-stat strong"));
        const progress = card.querySelector(".cover-progress span");
        return {
          title: card.dataset.coverTitle || "",
          mode: normalizeShelfMode(card.dataset.mode),
          time: shareText(stats[0]) || "--",
          progress: shareText(progress) || "",
          imageSrc: coverSourceForCard(card),
        };
      }).filter(item => item.title);
    }

    function loadShareImage(source) {
      return new Promise(resolve => {
        if (!source) {
          resolve(null);
          return;
        }
        const image = new Image();
        image.onload = () => resolve(image);
        image.onerror = () => resolve(null);
        image.src = source;
      });
    }

    async function loadShareBookRows(limit = 6) {
      const rows = shareBookRows(limit);
      await Promise.all(rows.map(async row => {
        row.image = await loadShareImage(row.imageSrc);
      }));
      return rows;
    }

    function drawShareStar(ctx, cx, cy, radius, fill, stroke) {
      ctx.save();
      ctx.beginPath();
      ctx.moveTo(cx, cy - radius);
      ctx.bezierCurveTo(
        cx + radius * .18, cy - radius,
        cx + radius * .28, cy - radius * .82,
        cx + radius * .34, cy - radius * .67
      );
      ctx.lineTo(cx + radius * .86, cy - radius * .49);
      ctx.bezierCurveTo(
        cx + radius * 1.02, cy - radius * .43,
        cx + radius * 1.03, cy - radius * .24,
        cx + radius * .88, cy - radius * .13
      );
      ctx.lineTo(cx + radius * .44, cy + radius * .3);
      ctx.lineTo(cx + radius * .58, cy + radius * .84);
      ctx.bezierCurveTo(
        cx + radius * .63, cy + radius * 1.02,
        cx + radius * .47, cy + radius * 1.12,
        cx + radius * .3, cy + radius * 1.02
      );
      ctx.lineTo(cx, cy + radius * .77);
      ctx.lineTo(cx - radius * .3, cy + radius * 1.02);
      ctx.bezierCurveTo(
        cx - radius * .47, cy + radius * 1.12,
        cx - radius * .63, cy + radius * 1.02,
        cx - radius * .58, cy + radius * .84
      );
      ctx.lineTo(cx - radius * .44, cy + radius * .3);
      ctx.lineTo(cx - radius * .88, cy - radius * .13);
      ctx.bezierCurveTo(
        cx - radius * 1.03, cy - radius * .24,
        cx - radius * 1.02, cy - radius * .43,
        cx - radius * .86, cy - radius * .49
      );
      ctx.lineTo(cx - radius * .34, cy - radius * .67);
      ctx.bezierCurveTo(
        cx - radius * .28, cy - radius * .82,
        cx - radius * .18, cy - radius,
        cx, cy - radius
      );
      ctx.closePath();
      ctx.fillStyle = fill;
      ctx.fill();
      ctx.strokeStyle = stroke;
      ctx.lineWidth = Math.max(1.5, radius * .13);
      ctx.lineJoin = "round";
      ctx.stroke();
      ctx.restore();
    }

    function drawShareRoundRect(ctx, x, y, width, height, radius, fill) {
      ctx.beginPath();
      const r = Math.min(radius, width / 2, height / 2);
      ctx.moveTo(x + r, y);
      ctx.arcTo(x + width, y, x + width, y + height, r);
      ctx.arcTo(x + width, y + height, x, y + height, r);
      ctx.arcTo(x, y + height, x, y, r);
      ctx.arcTo(x, y, x + width, y, r);
      ctx.closePath();
      ctx.fillStyle = fill;
      ctx.fill();
    }

    function drawShareText(ctx, text, x, y, options = {}) {
      const size = options.size || 32;
      const weight = options.weight || 500;
      const color = options.color || "#10201b";
      const maxWidth = options.maxWidth || 0;
      ctx.save();
      ctx.font = weight + " " + size + "px " +
        '"Microsoft YaHei", "PingFang SC", sans-serif';
      ctx.fillStyle = color;
      ctx.textBaseline = options.baseline || "top";
      if (!maxWidth) {
        ctx.fillText(String(text || ""), x, y);
        ctx.restore();
        return;
      }
      const value = String(text || "");
      const lines = [];
      let line = "";
      for (const char of value) {
        const next = line + char;
        if (ctx.measureText(next).width > maxWidth && line) {
          lines.push(line);
          line = char;
        } else {
          line = next;
        }
      }
      if (line) lines.push(line);
      lines.slice(0, options.maxLines || 2).forEach((entry, index) => {
        ctx.fillText(entry, x, y + index * (size * 1.35));
      });
      ctx.restore();
    }

    function createShareCanvas(width, height) {
      const canvas = document.createElement("canvas");
      canvas.width = width;
      canvas.height = height;
      return canvas;
    }

    function shareBackground(ctx, width, height) {
      const gradient = ctx.createLinearGradient(0, 0, width, height);
      gradient.addColorStop(0, "#133142");
      gradient.addColorStop(.55, "#17483f");
      gradient.addColorStop(1, "#2c5d67");
      ctx.fillStyle = gradient;
      ctx.fillRect(0, 0, width, height);
      ctx.save();
      ctx.globalAlpha = .24;
      for (let index = 0; index < 54; index += 1) {
        const x = (index * 173 + 41) % width;
        const y = (index * 97 + 67) % height;
        const radius = 1.2 + (index % 3) * .8;
        ctx.fillStyle = index % 4 === 0 ? "#FFD43B" : "#f7fbff";
        ctx.beginPath();
        ctx.arc(x, y, radius, 0, Math.PI * 2);
        ctx.fill();
      }
      ctx.restore();
    }

    function drawShareHeader(ctx, width, title, subtitle) {
      ctx.fillStyle = "rgba(255, 255, 255, .10)";
      drawShareRoundRect(ctx, 56, 54, width - 112, 188, 26, ctx.fillStyle);
      drawShareText(ctx, title, 88, 84, {
        size: 46, weight: 800, color: "#fff8dc", maxWidth: width - 176,
      });
      drawShareText(ctx, subtitle, 88, 156, {
        size: 24, color: "rgba(248,251,255,.78)", maxWidth: width - 176,
      });
    }

    function drawShareMetrics(ctx, rows, x, y, width) {
      const gap = 18;
      const cellWidth = (width - gap) / 2;
      rows.slice(0, 6).forEach((row, index) => {
        const col = index % 2;
        const rowIndex = Math.floor(index / 2);
        const left = x + col * (cellWidth + gap);
        const top = y + rowIndex * 126;
        drawShareRoundRect(
          ctx, left, top, cellWidth, 108, 18, "rgba(255,255,255,.12)"
        );
        drawShareText(ctx, row.label, left + 22, top + 17, {
          size: 22, color: "rgba(248,251,255,.72)", maxWidth: cellWidth - 44,
        });
        drawShareText(ctx, row.value, left + 22, top + 52, {
          size: 34, weight: 800, color: "#fff7dc", maxWidth: cellWidth - 44,
        });
      });
    }

    function drawShareAchievements(ctx, rows, x, y, width) {
      rows.slice(0, 4).forEach((row, index) => {
        const top = y + index * 116;
        drawShareRoundRect(
          ctx, x, top, width, 98, 18, "rgba(255,255,255,.105)"
        );
        drawShareText(ctx, row.label, x + 24, top + 17, {
          size: 27, weight: 800, color: "#fff7dc", maxWidth: width - 270,
        });
        drawShareText(ctx, row.encouragement || "", x + 24, top + 54, {
          size: 19, color: "rgba(248,251,255,.72)", maxWidth: width - 280,
          maxLines: 1,
        });
        for (let starIndex = 0; starIndex < row.total; starIndex += 1) {
          drawShareStar(
            ctx,
            x + width - 190 + starIndex * 43,
            top + 47,
            18,
            starIndex < row.lit ? "#FFD43B" : "rgba(255,255,255,.08)",
            starIndex < row.lit ? "#C58A1A" : "rgba(230,220,190,.42)"
          );
        }
      });
    }

    function drawShareCover(ctx, row, x, y, width, height) {
      ctx.save();
      drawShareRoundRect(ctx, x, y, width, height, 12, "rgba(255,255,255,.12)");
      if (row.image && row.image.naturalWidth > 0) {
        ctx.beginPath();
        const radius = 12;
        ctx.moveTo(x + radius, y);
        ctx.arcTo(x + width, y, x + width, y + height, radius);
        ctx.arcTo(x + width, y + height, x, y + height, radius);
        ctx.arcTo(x, y + height, x, y, radius);
        ctx.arcTo(x, y, x + width, y, radius);
        ctx.closePath();
        ctx.clip();
        const scale = Math.max(
          width / row.image.naturalWidth,
          height / row.image.naturalHeight
        );
        const drawWidth = row.image.naturalWidth * scale;
        const drawHeight = row.image.naturalHeight * scale;
        ctx.drawImage(
          row.image,
          x + (width - drawWidth) / 2,
          y + (height - drawHeight) / 2,
          drawWidth,
          drawHeight
        );
      } else {
        drawShareRoundRect(ctx, x, y, width, height, 12, "rgba(255,255,255,.08)");
        const initial = (row.title || "?").trim().slice(0, 1) || "?";
        drawShareText(ctx, initial, x + width / 2 - 18, y + height / 2 - 30, {
          size: 52, weight: 800, color: "rgba(255,247,220,.82)",
        });
      }
      ctx.strokeStyle = "rgba(255,255,255,.22)";
      ctx.lineWidth = 1.5;
      ctx.strokeRect(x + .75, y + .75, width - 1.5, height - 1.5);
      ctx.restore();
    }

    function drawShareBooks(ctx, rows, x, y, width, limit = 3) {
      rows.slice(0, limit).forEach((row, index) => {
        const top = y + index * 150;
        const coverWidth = 92;
        const coverHeight = 126;
        drawShareRoundRect(
          ctx, x, top, width, 136, 20, "rgba(255,255,255,.105)"
        );
        drawShareCover(ctx, row, x + 18, top + 5, coverWidth, coverHeight);
        drawShareText(ctx, String(index + 1).padStart(2, "0"), x + 18, top + 114, {
          size: 18, weight: 800, color: "#FFD43B",
        });
        drawShareText(ctx, row.title, x + 132, top + 14, {
          size: 28, weight: 800, color: "#fff7dc", maxWidth: width - 360,
          maxLines: 1,
        });
        drawShareText(
          ctx,
          (row.mode === "manga" ? "漫画" : "图书") + " · " + row.time,
          x + 132, top + 57,
          { size: 21, color: "rgba(248,251,255,.76)", maxWidth: width - 360 }
        );
        if (row.progress) {
          drawShareRoundRect(
            ctx, x + width - 134, top + 20, 112, 38, 19,
            "rgba(255,212,59,.16)"
          );
          drawShareText(ctx, row.progress, x + width - 118, top + 29, {
            size: 19, weight: 800, color: "#FFD43B", maxWidth: 84,
          });
        }
        if (index < limit - 1) {
          ctx.fillStyle = "rgba(255,255,255,.09)";
          ctx.fillRect(x + 18, top + 144, width - 36, 1);
        }
      });
    }

    function canvasToPng(canvas, fileName) {
      return new Promise((resolve, reject) => {
        canvas.toBlob(blob => {
          if (!blob) {
            reject(new Error("PNG generation failed"));
            return;
          }
          const url = URL.createObjectURL(blob);
          const link = document.createElement("a");
          link.href = url;
          link.download = fileName;
          document.body.appendChild(link);
          link.click();
          link.remove();
          window.setTimeout(() => URL.revokeObjectURL(url), 1000);
          resolve();
        }, "image/png");
      });
    }

    function createLongShareCanvas(bookRows = null) {
      const width = 1080;
      const height = 1920;
      const canvas = createShareCanvas(width, height);
      const ctx = canvas.getContext("2d");
      shareBackground(ctx, width, height);
      drawShareHeader(
        ctx,
        width,
        "这是我的 Kindle 阅读记录",
        "把安静阅读的时光，收进一张分享图"
      );
      drawShareMetrics(ctx, shareMetricRows(), 56, 280, width - 112);
      drawShareText(ctx, "阅读成就", 56, 704, {
        size: 34, weight: 800, color: "#fff7dc",
      });
      drawShareAchievements(ctx, shareAchievementRows(), 56, 752, width - 112);
      drawShareText(ctx, "重点阅读", 56, 1228, {
        size: 34, weight: 800, color: "#fff7dc",
      });
      drawShareText(ctx, "按阅读时长和进度挑选", 56, 1272, {
        size: 20, color: "rgba(248,251,255,.66)",
      });
      drawShareBooks(
        ctx,
        bookRows || shareBookRows(3),
        56,
        1320,
        width - 112,
        3
      );
      drawShareText(ctx, "离线生成 · Kindle 阅读统计", 56, 1818, {
        size: 22, color: "rgba(248,251,255,.58)",
      });
      return canvas;
    }

    function createSummarySetCanvas() {
      const width = 1080;
      const height = 1350;
      const canvas = createShareCanvas(width, height);
      const ctx = canvas.getContext("2d");
      shareBackground(ctx, width, height);
      drawShareHeader(
        ctx,
        width,
        "我的阅读总览",
        "阅读时长、词数与进度"
      );
      drawShareMetrics(ctx, shareMetricRows(), 56, 292, width - 112);
      drawShareText(ctx, "四轨成就", 56, 744, {
        size: 34, weight: 800, color: "#fff7dc",
      });
      drawShareAchievements(ctx, shareAchievementRows(), 56, 796, width - 112);
      return canvas;
    }

    function createBooksSetCanvas(bookRows = null) {
      const width = 1080;
      const height = 1350;
      const canvas = createShareCanvas(width, height);
      const ctx = canvas.getContext("2d");
      shareBackground(ctx, width, height);
      drawShareHeader(
        ctx,
        width,
        "重点阅读与星空",
        "只保留最突出的阅读记录"
      );
      drawShareText(ctx, "数据突出的几本", 56, 292, {
        size: 32, weight: 800, color: "#fff7dc",
      });
      drawShareBooks(
        ctx,
        bookRows || shareBookRows(3),
        56,
        344,
        width - 112,
        3
      );
      const unlocked = parseFloat(
        document.documentElement.dataset.achievementStarsUnlocked || "0"
      ) || 0;
      const total = parseFloat(
        document.documentElement.dataset.achievementStarsTotal || "16"
      ) || 16;
      drawShareText(ctx, "阅读星空", 56, 846, {
        size: 34, weight: 800, color: "#fff7dc",
      });
      for (let index = 0; index < 16; index += 1) {
        const col = index % 8;
        const row = Math.floor(index / 8);
        drawShareStar(
          ctx, 118 + col * 120, 930 + row * 120, 40,
          index < unlocked ? "#FFD43B" : "rgba(255,255,255,.08)",
          index < unlocked ? "#C58A1A" : "rgba(230,220,190,.42)"
        );
      }
      drawShareText(ctx, unlocked + "/" + total + " 颗星已点亮", 56, 1216, {
        size: 25, color: "rgba(248,251,255,.78)",
      });
      return canvas;
    }

    async function generateShareImages() {
      if (!shareStatus) return;
      shareStatus.textContent = "生成中...";
      const format = shareFormat ? shareFormat.value : "long";
      try {
        const bookRows = await loadShareBookRows(3);
        if (format === "set") {
          await canvasToPng(createSummarySetCanvas(), "kindle-reading-summary.png");
          await canvasToPng(
            createBooksSetCanvas(bookRows),
            "kindle-reading-books.png"
          );
        } else {
          await canvasToPng(
            createLongShareCanvas(bookRows),
            "kindle-reading-long.png"
          );
        }
        shareStatus.textContent = format === "set"
          ? "已生成 2 张分享图。"
          : "已生成长图。";
      } catch (error) {
        shareStatus.textContent = "生成失败：" + (error.message || "未知错误");
      }
    }

    function openShareDialog() {
      if (!shareDialog) return;
      if (typeof shareDialog.showModal === "function") {
        shareDialog.showModal();
      } else {
        shareDialog.hidden = false;
        shareDialog.setAttribute("open", "");
      }
      shareClose?.focus();
    }

    function closeShareDialog() {
      if (!shareDialog) return;
      if (typeof shareDialog.close === "function") {
        shareDialog.close();
      } else {
        shareDialog.hidden = true;
        shareDialog.removeAttribute("open");
      }
      shareExport?.focus();
    }

    function setUserName(value, persist = true) {
      currentUserName = normalizeUserName(value);
      if (userNameInput) userNameInput.value = currentUserName;
      if (persist) writePreference(USER_NAME_KEY, currentUserName);
      applyAchievementLevels(true);
      updateGreeting(false);
    }

    function hideUserNameDialog() {
      if (!userNameDialog) return;
      if (typeof userNameDialog.close === "function") {
        userNameDialog.close();
      } else {
        userNameDialog.hidden = true;
      }
    }

    function showUserNameDialog() {
      if (!userNameDialog) return;
      if (userNameDialog.open) return;
      if (userNameDialogInput) {
        userNameDialogInput.value = currentUserName;
      }
      if (typeof userNameDialog.show === "function") {
        userNameDialog.show();
      } else {
        userNameDialog.hidden = false;
      }
      userNameDialogInput?.focus();
    }

    function restoreUserName() {
      const stored = readPreference(USER_NAME_KEY);
      if (stored === null) {
        currentUserName = "";
        if (userNameInput) userNameInput.value = "";
        showUserNameDialog();
      } else {
        currentUserName = normalizeUserName(stored);
        if (userNameInput) userNameInput.value = currentUserName;
      }
    }

    function updateAchievementRootAttributes() {
      const visibleCards = achievementCards.filter(card => !card.hidden);
      const unlocked = visibleCards.filter(card => {
        const maxLevel = Number(card.dataset.achievementMaxLevel) || 0;
        return maxLevel > 0 &&
          (Number(card.dataset.achievementLevel) || 0) >= maxLevel;
      }).length;
      const starsUnlocked = visibleCards.reduce(
        (total, card) =>
          total + Math.max(0, Number(card.dataset.achievementLevel) || 0),
        0
      );
      const starsTotal = visibleCards.reduce(
        (total, card) =>
          total + Math.max(0, Number(card.dataset.achievementMaxLevel) || 0),
        0
      );
      document.documentElement.dataset.achievementPanel =
        achievementPanel && !achievementPanel.hidden && achievementPanel.open
          ? "open"
          : "closed";
      document.documentElement.dataset.achievementEnabled =
        achievementPanel && !achievementPanel.hidden ? "true" : "false";
      document.documentElement.dataset.achievementCount =
        String(visibleCards.length);
      document.documentElement.dataset.achievementUnlocked =
        String(unlocked);
      document.documentElement.dataset.achievementStarsUnlocked =
        String(starsUnlocked);
      document.documentElement.dataset.achievementStarsTotal =
        String(starsTotal);
      if (achievementSummary) {
        achievementSummary.innerHTML =
          '<span class="i18n-zh">' +
          starsUnlocked + "/" + starsTotal + " 颗星已点亮</span>" +
          '<span class="i18n-en">' +
          starsUnlocked + "/" + starsTotal + " stars unlocked</span>";
      }
    }

    function restoreAchievementPanel() {
      if (!achievementPanel) return;
      const stored = readPreference(ACHIEVEMENT_PANEL_KEY);
      achievementPanel.open = stored === null
        ? true
        : stored === "open";
      restoreAchievementEnabled();
      restoreMangaVolumeAchievement();
      updateAchievementRootAttributes();
    }

    function setAchievementEnabled(enabled, persist = true) {
      const normalized = enabled !== false;
      if (achievementPanel) {
        achievementPanel.hidden = !normalized;
      }
      if (achievementEnabledToggle) {
        achievementEnabledToggle.checked = normalized;
      }
      document.documentElement.dataset.achievementEnabled =
        normalized ? "true" : "false";
      if (persist) {
        writePreference(
          ACHIEVEMENT_ENABLED_KEY,
          normalized ? "true" : "false"
        );
      }
      updateAchievementRootAttributes();
    }

    function restoreAchievementEnabled() {
      setAchievementEnabled(
        readPreference(ACHIEVEMENT_ENABLED_KEY) !== "false",
        false
      );
    }

    function setMangaVolumeAchievement(enabled, persist = true) {
      mangaVolumesEnabled = enabled === true;
      achievementCards.forEach(card => {
        if (card.dataset.achievementKey === "manga_volumes") {
          card.hidden = !mangaVolumesEnabled;
        }
      });
      if (mangaVolumesToggle) {
        mangaVolumesToggle.checked = mangaVolumesEnabled;
      }
      document.documentElement.dataset.mangaVolumeAchievement =
        mangaVolumesEnabled ? "true" : "false";
      if (persist) {
        writePreference(
          MANGA_VOLUME_ACHIEVEMENT_KEY,
          mangaVolumesEnabled ? "true" : "false"
        );
      }
      updateAchievementRootAttributes();
    }

    function restoreMangaVolumeAchievement() {
      setMangaVolumeAchievement(
        readPreference(MANGA_VOLUME_ACHIEVEMENT_KEY) === "true",
        false
      );
    }

    function normalizeAchievementBadgeStyle(value) {
      const normalized = String(value || "").toLowerCase();
      return ACHIEVEMENT_BADGE_STYLES.has(normalized)
        ? normalized
        : "seal";
    }

    function setAchievementBadgeStyle(value, persist = true) {
      const normalized = normalizeAchievementBadgeStyle(value);
      document.documentElement.dataset.achievementBadgeStyle = normalized;
      if (achievementBadgeStyleSelect) {
        achievementBadgeStyleSelect.value = normalized;
      }
      if (persist) {
        writePreference(ACHIEVEMENT_BADGE_STYLE_KEY, normalized);
      }
    }

    function normalizeShelfTitleStyle(value) {
      const normalized = String(value || "").toLowerCase();
      return SHELF_TITLE_STYLES.has(normalized) ? normalized : "a";
    }

    function setShelfTitleStyle(value, persist = true) {
      shelfTitleStyle = normalizeShelfTitleStyle(value);
      document.documentElement.dataset.shelfTitleStyle = shelfTitleStyle;
      if (shelfTitleStyleSelect) {
        shelfTitleStyleSelect.value = shelfTitleStyle;
      }
      if (persist) {
        writePreference(SHELF_TITLE_STYLE_KEY, shelfTitleStyle);
      }
    }

    function normalizeShelfPosition(value) {
      const normalized = String(value || "").toLowerCase();
      return SHELF_POSITIONS.has(normalized) ? normalized : "title";
    }

    function setShelfFallbackOpen(open) {
      const expanded = open === true;
      if (shelfFallbackMenu) {
        shelfFallbackMenu.hidden = !expanded;
      }
      if (shelfFallbackButton) {
        shelfFallbackButton.setAttribute(
          "aria-expanded",
          expanded ? "true" : "false"
        );
      }
    }

    function setShelfPosition(value, persist = true) {
      shelfPosition = normalizeShelfPosition(value);
      document.documentElement.dataset.shelfPosition = shelfPosition;
      if (shelfPositionSelect) {
        shelfPositionSelect.value = shelfPosition;
      }
      if (shelfModeSwitch && libraryActions && shelfTitleCluster) {
        if (shelfPosition === "actions") {
          libraryActions.insertBefore(
            shelfModeSwitch,
            shelfFallbackWrap || shelfModeSwitch
          );
        } else {
          shelfTitleCluster.appendChild(shelfModeSwitch);
        }
      }
      setShelfFallbackOpen(false);
      if (persist) {
        writePreference(SHELF_POSITION_KEY, shelfPosition);
      }
    }

    function normalizeShelfDefaultMode(value) {
      const normalized = String(value || "").toLowerCase();
      return SHELF_DEFAULT_MODES.has(normalized) ? normalized : "book";
    }

    function setShelfDefaultMode(value, persist = true) {
      shelfDefaultMode = normalizeShelfDefaultMode(value);
      document.documentElement.dataset.shelfDefaultMode = shelfDefaultMode;
      if (shelfDefaultModeSelect) {
        shelfDefaultModeSelect.value = shelfDefaultMode;
      }
      if (persist) {
        writePreference(SHELF_DEFAULT_MODE_KEY, shelfDefaultMode);
      }
    }

    function setShowTokenLabels(value, persist = true) {
      showTokenLabels = value === true || value === "true";
      document.documentElement.dataset.showTokenLabels =
        showTokenLabels ? "true" : "false";
      if (showTokenLabelsToggle) {
        showTokenLabelsToggle.checked = showTokenLabels;
      }
      if (persist) {
        writePreference(
          SHOW_TOKEN_LABELS_KEY,
          showTokenLabels ? "true" : "false"
        );
      }
    }

    function restoreShelfPreferences() {
      setShelfTitleStyle(
        readPreference(SHELF_TITLE_STYLE_KEY) || "a",
        false
      );
      setShelfPosition(
        readPreference(SHELF_POSITION_KEY) || "title",
        false
      );
      setShelfDefaultMode(
        readPreference(SHELF_DEFAULT_MODE_KEY) || "book",
        false
      );
      setShowTokenLabels(
        readPreference(SHOW_TOKEN_LABELS_KEY) === "true",
        false
      );
      setShelfMode(shelfDefaultMode);
    }

    function normalizeBackgroundTexture(value) {
      const normalized = String(value || "").toLowerCase();
      return BACKGROUND_TEXTURES.has(normalized) ? normalized : "paper";
    }

    function setBackgroundTexture(value, persist = true) {
      const normalized = normalizeBackgroundTexture(value);
      document.documentElement.dataset.backgroundTexture = normalized;
      if (backgroundTextureSelect) {
        backgroundTextureSelect.value = normalized;
      }
      if (persist) {
        writePreference(BACKGROUND_TEXTURE_KEY, normalized);
      }
    }

    function normalizeBackgroundMotif(value) {
      const normalized = String(value || "").toLowerCase();
      return BACKGROUND_MOTIFS.has(normalized) ? normalized : "none";
    }

    function setBackgroundMotif(value, persist = true) {
      const normalized = normalizeBackgroundMotif(value);
      document.documentElement.dataset.backgroundMotif = normalized;
      if (backgroundMotifSelect) {
        backgroundMotifSelect.value = normalized;
      }
      if (persist) {
        writePreference(BACKGROUND_MOTIF_KEY, normalized);
      }
    }
    if (backgroundMotifSelect) {
      backgroundMotifSelect.addEventListener("change", () => {
        setBackgroundMotif(backgroundMotifSelect.value);
      });
    }

    function normalizePeriodView(value) {
      const normalized = String(value || "").toLowerCase();
      return PERIOD_VIEWS.has(normalized) ? normalized : "week";
    }

    function updatePeriodRootAttributes() {
      const active = periodViews.find(view => !view.hidden);
      document.documentElement.dataset.periodPanel =
        periodPanel && periodPanel.open ? "open" : "closed";
      document.documentElement.dataset.periodView =
        active ? normalizePeriodView(active.dataset.periodPanel) : "week";
      document.documentElement.dataset.periodReliable =
        active ? active.dataset.periodReliable || "false" : "false";
      document.documentElement.dataset.periodSource =
        active ? active.dataset.periodSource || "" : "";
      document.documentElement.dataset.periodTrendType =
        periodBarCharts.length ? "bars" : "none";
      document.documentElement.dataset.periodTrendBuckets = String(
        periodBarCharts.reduce(
          (maximum, chart) => Math.max(
            maximum,
            Number(chart.dataset.periodBucketCount) || 0
          ),
          0
        )
      );
    }

    function setPeriodView(value, persist = true) {
      const normalized = normalizePeriodView(value);
      periodViews.forEach(view => {
        view.hidden =
          normalizePeriodView(view.dataset.periodPanel) !== normalized;
      });
      periodTabs.forEach(tab => {
        const selected =
          normalizePeriodView(tab.dataset.periodTab) === normalized;
        tab.setAttribute("aria-selected", selected ? "true" : "false");
        tab.tabIndex = selected ? 0 : -1;
      });
      if (persist) writePreference(PERIOD_VIEW_KEY, normalized);
      updatePeriodRootAttributes();
    }

    function restorePeriodPanel() {
      if (!periodPanel) return;
      const stored = readPreference(PERIOD_PANEL_KEY);
      periodPanel.open = stored === null ? true : stored === "open";
      updatePeriodRootAttributes();
    }

    function restorePeriodView() {
      setPeriodView(readPreference(PERIOD_VIEW_KEY) || "week", false);
    }

    function updateAnnualReportRootAttributes() {
      document.documentElement.dataset.annualReportPanel =
        annualReportPanel && annualReportPanel.open ? "open" : "closed";
      document.documentElement.dataset.annualReportState =
        annualReportPanel
          ? annualReportPanel.dataset.annualReportState || "empty"
          : "empty";
      document.documentElement.dataset.annualReportReliable =
        annualReportPanel
          ? annualReportPanel.dataset.annualReportReliable || "false"
          : "false";
    }

    function restoreAnnualReportPanel() {
      if (!annualReportPanel) {
        updateAnnualReportRootAttributes();
        return;
      }
      const stored = readPreference(ANNUAL_REPORT_PANEL_KEY);
      annualReportPanel.open = stored === "open";
      updateAnnualReportRootAttributes();
    }

    function preparePeriodPrint() {
      if (periodPanel) periodPanel.open = true;
      if (annualReportPanel) annualReportPanel.open = true;
      periodViews.forEach(view => {
        view.hidden = false;
      });
      updatePeriodRootAttributes();
      updateAnnualReportRootAttributes();
    }

    function restoreAfterPeriodPrint() {
      restorePeriodPanel();
      restorePeriodView();
      restoreAnnualReportPanel();
      updateLayoutMetrics();
    }

    function normalizeTrendMetric(value) {
      const normalized = String(value || "").toLowerCase();
      return TREND_METRICS.has(normalized) ? normalized : "time";
    }

    function updateTrendRootAttributes() {
      document.documentElement.dataset.trendMetric =
        normalizeTrendMetric(document.documentElement.dataset.trendMetric);
      document.documentElement.dataset.trendDetails =
        trendDetails && trendDetails.open ? "open" : "closed";
      document.documentElement.dataset.trendTooltip =
        trendTooltip && !trendTooltip.hidden ? "visible" : "hidden";
    }

    function hideTrendTooltip() {
      if (!trendTooltip) return;
      trendTooltip.hidden = true;
      updateTrendRootAttributes();
    }

    function showTrendTooltip(point) {
      if (!trendTooltip || !point) return;
      const english =
        document.documentElement.dataset.uiLanguage === "en";
      const text = english
        ? point.dataset.tooltipEn || point.dataset.tooltipZh || ""
        : point.dataset.tooltipZh || point.dataset.tooltipEn || "";
      if (!text) return;
      const chart = point.closest(".trend-chart");
      if (!chart) return;
      const chartRect = chart.getBoundingClientRect();
      const pointRect = point.getBoundingClientRect();
      const left = Math.min(
        Math.max(pointRect.left - chartRect.left, 12),
        Math.max(12, chartRect.width - 12)
      );
      const top = Math.max(0, pointRect.top - chartRect.top - 44);
      trendTooltip.textContent = text;
      trendTooltip.style.left = left + "px";
      trendTooltip.style.top = top + "px";
      trendTooltip.hidden = false;
      updateTrendRootAttributes();
    }

    function initializeTrendPoints() {
      document.querySelectorAll("[data-trend-point]").forEach(point => {
        point.addEventListener("mouseenter", () => showTrendTooltip(point));
        point.addEventListener("focus", () => showTrendTooltip(point));
        point.addEventListener("mouseleave", hideTrendTooltip);
        point.addEventListener("blur", hideTrendTooltip);
      });
    }

    function setTrendMetric(metric, persist = true) {
      const normalized = normalizeTrendMetric(metric);
      document.documentElement.dataset.trendMetric = normalized;
      trendSeriesViews.forEach(view => {
        view.hidden = view.dataset.trendSeries !== normalized;
      });
      trendMetricButtons.forEach(button => {
        button.setAttribute(
          "aria-pressed",
          button.dataset.trendMetricButton === normalized ? "true" : "false"
        );
      });
      hideTrendTooltip();
      if (persist) writePreference(TREND_METRIC_KEY, normalized);
      updateTrendRootAttributes();
    }

    function restoreTrendDetails() {
      if (!trendDetails) return;
      trendDetails.open = readPreference(TREND_DETAILS_KEY) === "open";
      updateTrendRootAttributes();
    }

    function clampPreference(value, fallback, minimum, maximum) {
      const number = Number(value);
      if (!Number.isFinite(number)) return fallback;
      return Math.min(maximum, Math.max(minimum, Math.round(number)));
    }

    function normalizeBackgroundSettings(value) {
      const raw = value && typeof value === "object" ? value : {};
      return {
        source: BACKGROUND_SOURCES.has(raw.source) ? raw.source : "none",
        blur: clampPreference(raw.blur, 12, 0, 30),
        overlay: clampPreference(raw.overlay, 62, 0, 90),
        opacity: clampPreference(raw.opacity, 72, 10, 100),
        bookId: typeof raw.bookId === "string" ? raw.bookId : "",
      };
    }

    function readBackgroundSettings() {
      const stored = readPreference(BACKGROUND_SETTINGS_KEY);
      if (!stored) return {...DEFAULT_BACKGROUND_SETTINGS};
      try {
        return normalizeBackgroundSettings(JSON.parse(stored));
      } catch (error) {
        return {...DEFAULT_BACKGROUND_SETTINGS};
      }
    }

    function saveBackgroundSettings() {
      return writePreference(
        BACKGROUND_SETTINGS_KEY,
        JSON.stringify(backgroundSettings)
      );
    }

    function coverSourceForCard(card) {
      const image = card.querySelector("img[data-cover-image]");
      if (!image || image.dataset.coverFailed === "true") return "";
      const source = image.getAttribute("src") || "";
      if (!source) return "";
      if (image.complete && image.naturalWidth === 0) return "";
      return image.currentSrc || image.src || source;
    }

    function latestLastReadCard(cards) {
      let selected = null;
      let selectedTimestamp = -Infinity;
      for (const card of cards) {
        const timestamp = Date.parse(card.dataset.lastRead || "");
        if (Number.isFinite(timestamp) && timestamp > selectedTimestamp) {
          selected = card;
          selectedTimestamp = timestamp;
        }
      }
      return selected;
    }

    function trendChangeCard(cards) {
      const trendDataElement = document.getElementById("trend-data");
      if (!trendDataElement) return null;
      let trendData = null;
      try {
        trendData = JSON.parse(trendDataElement.textContent || "{}");
      } catch (error) {
        return null;
      }
      const changes = Array.isArray(trendData.books)
        ? trendData.books.slice()
        : [];
      changes.sort((left, right) => {
        const movement = change => {
          const time = Number(
            change && change.total_time_ms &&
            change.total_time_ms.delta
          ) || 0;
          const words = Number(
            change && change.total_words &&
            change.total_words.delta
          ) || 0;
          return Math.max(0, time) + Math.max(0, words) * 100;
        };
        return movement(right) - movement(left);
      });
      for (const change of changes) {
        if (!change || !change.book_id) continue;
        const card = cards.find(
          item => item.dataset.bookId === String(change.book_id)
        );
        if (card) return card;
      }
      return null;
    }

    function updateAppearanceStatus() {
      if (!appearanceStatus) return;
      const english =
        document.documentElement.dataset.uiLanguage === "en";
      appearanceStatus.textContent = english
        ? appearanceStatus.dataset.statusEn || ""
        : appearanceStatus.dataset.statusZh || "";
    }

    function setAppearanceStatus(zh, en) {
      if (!appearanceStatus) return;
      appearanceStatus.dataset.statusZh = zh;
      appearanceStatus.dataset.statusEn = en;
      updateAppearanceStatus();
    }

    function cssBackgroundUrl(value) {
      const escaped = String(value)
        .replace(/\\\\/g, "\\\\\\\\")
        .replace(/"/g, '\\\\"');
      return 'url("' + escaped + '")';
    }

    function populateBackgroundBookOptions() {
      if (!backgroundBook) return;
      const previous = backgroundSettings.bookId;
      const cards = Array.from(
        coverGrid.querySelectorAll(".book-cover-card[data-book-id]")
      );
      backgroundBook.textContent = "";
      if (!cards.length) {
        const option = document.createElement("option");
        option.value = "";
        option.disabled = true;
        option.textContent = "无可用书籍 / No books";
        backgroundBook.appendChild(option);
        return;
      }
      for (const card of cards) {
        const option = document.createElement("option");
        option.value = card.dataset.bookId || "";
        const title = card.dataset.coverTitle ||
          card.querySelector("h3")?.textContent.trim() ||
          option.value;
        const available = Boolean(coverSourceForCard(card));
        option.disabled = !available;
        option.textContent = title + (available ? "" : " / 无封面 / no cover");
        backgroundBook.appendChild(option);
      }
      const hasPrevious = Array.from(backgroundBook.options).some(
        option => option.value === previous && !option.disabled
      );
      if (hasPrevious) backgroundBook.value = previous;
    }

    function syncAppearanceControls() {
      if (themePresetSelect) {
        themePresetSelect.value = document.documentElement.dataset.themePreset ||
          "forest";
      }
      if (themeFollowGreeting) {
        themeFollowGreeting.checked = themeFollowsGreeting;
      }
      if (greetingFontSelect) {
        greetingFontSelect.value = normalizeGreetingFont(
          document.documentElement.dataset.greetingFont
        );
      }
      if (userNameInput) {
        userNameInput.value = currentUserName;
      }
      if (achievementBadgeStyleSelect) {
        achievementBadgeStyleSelect.value = normalizeAchievementBadgeStyle(
          document.documentElement.dataset.achievementBadgeStyle
        );
      }
      if (backgroundTextureSelect) {
        backgroundTextureSelect.value = normalizeBackgroundTexture(
          document.documentElement.dataset.backgroundTexture
        );
      }
      if (backgroundMotifSelect) {
        backgroundMotifSelect.value = normalizeBackgroundMotif(
          document.documentElement.dataset.backgroundMotif
        );
      }
      if (backgroundSource) {
        backgroundSource.value = backgroundSettings.source;
      }
      if (fixedCoverField) {
        fixedCoverField.hidden = backgroundSettings.source !== "fixed";
      }
      if (backgroundBlur) backgroundBlur.value = String(backgroundSettings.blur);
      if (backgroundOverlay) {
        backgroundOverlay.value = String(backgroundSettings.overlay);
      }
      if (backgroundOpacity) {
        backgroundOpacity.value = String(backgroundSettings.opacity);
      }
      if (backgroundBlurValue) {
        backgroundBlurValue.textContent = backgroundSettings.blur + " px";
      }
      if (backgroundOverlayValue) {
        backgroundOverlayValue.textContent =
          backgroundSettings.overlay + "%";
      }
      if (backgroundOpacityValue) {
        backgroundOpacityValue.textContent =
          backgroundSettings.opacity + "%";
      }
      populateBackgroundBookOptions();
      if (backgroundBook && backgroundSettings.bookId) {
        const option = Array.from(backgroundBook.options).find(
          item => item.value === backgroundSettings.bookId && !item.disabled
        );
        if (option) backgroundBook.value = option.value;
      }
    }

    function applyAppearanceSettings(announce = true) {
      const cards = Array.from(
        coverGrid.querySelectorAll(".book-cover-card[data-book-id]")
      );
      let selectedCard = null;
      let imageUrl = "";
      let reason = "theme";
      let statusZh = "背景已关闭，使用主题背景。";
      let statusEn = "Background is off; using the theme background.";

      if (backgroundSettings.source === "upload") {
        imageUrl = backgroundImageDataUrl;
        reason = imageUrl ? "upload" : "upload-empty";
        if (imageUrl) {
          statusZh = "已使用本地上传背景；图片仅保存在当前浏览器。";
          statusEn = "Using the local upload; the image stays in this browser only.";
        } else {
          statusZh = "尚未选择本地图片，已回退主题背景。";
          statusEn = "No local image selected; using the theme background.";
        }
      } else if (backgroundSettings.source === "recent") {
        selectedCard = latestLastReadCard(cards);
        if (selectedCard) {
          reason = "last-read";
        } else {
          selectedCard = trendChangeCard(cards);
          if (selectedCard) {
            reason = "snapshot-change";
          } else if (cards.length) {
            selectedCard = cards[0];
            reason = "ranking-fallback";
          } else {
            reason = "empty";
          }
        }
        imageUrl = selectedCard ? coverSourceForCard(selectedCard) : "";
        const title = selectedCard
          ? selectedCard.dataset.coverTitle ||
            selectedCard.querySelector("h3")?.textContent.trim() ||
            ""
          : "";
        if (!selectedCard) {
          statusZh = "暂无可用于背景的书籍，已回退主题背景。";
          statusEn = "No book is available for the background; using the theme background.";
        } else if (!imageUrl) {
          reason += "-cover-missing";
          statusZh = "所选最近书籍没有可用封面，已回退主题背景。";
          statusEn = "The selected recent book has no usable cover; using the theme background.";
        } else if (reason === "last-read") {
          statusZh = "自动使用最近阅读封面：" + title + "。";
          statusEn = "Using the latest last-read cover: " + title + ".";
        } else if (reason === "snapshot-change") {
          statusZh = "last_read 缺失，按最近快照变化使用：" + title + "。";
          statusEn = "last_read is missing; using the latest changed snapshot book: " + title + ".";
        } else {
          statusZh = "缺少时间与快照变化，按阅读时长使用：" + title + "。";
          statusEn = "No recent timestamps or changes; using reading-time rank: " + title + ".";
        }
      } else if (backgroundSettings.source === "fixed") {
        selectedCard = cards.find(
          card => card.dataset.bookId === backgroundSettings.bookId
        ) || null;
        imageUrl = selectedCard ? coverSourceForCard(selectedCard) : "";
        const title = selectedCard
          ? selectedCard.dataset.coverTitle ||
            selectedCard.querySelector("h3")?.textContent.trim() ||
            ""
          : "";
        if (!selectedCard) {
          reason = "missing-selection";
          statusZh = "固定书籍不存在，已回退主题背景。";
          statusEn = "The fixed book no longer exists; using the theme background.";
        } else if (!imageUrl) {
          reason = "fixed-cover-missing";
          statusZh = "固定书籍没有可用封面，已回退主题背景。";
          statusEn = "The fixed book has no usable cover; using the theme background.";
        } else {
          reason = "fixed-cover";
          statusZh = "固定使用书籍封面：" + title + "。";
          statusEn = "Using the fixed book cover: " + title + ".";
        }
      }

      document.documentElement.style.setProperty(
        "--dashboard-background-blur",
        backgroundSettings.blur + "px"
      );
      document.documentElement.style.setProperty(
        "--dashboard-background-overlay",
        String(backgroundSettings.overlay / 100)
      );
      document.documentElement.dataset.backgroundSource =
        backgroundSettings.source;
      document.documentElement.dataset.backgroundReason = reason;
      document.documentElement.dataset.backgroundBookId =
        selectedCard ? selectedCard.dataset.bookId || "" : "";
      if (imageUrl) {
        document.documentElement.style.setProperty(
          "--dashboard-background-image",
          cssBackgroundUrl(imageUrl)
        );
        document.documentElement.style.setProperty(
          "--dashboard-background-opacity",
          String(backgroundSettings.opacity / 100)
        );
        document.documentElement.dataset.backgroundActive = "true";
        document.documentElement.dataset.backgroundFallback = "false";
      } else {
        document.documentElement.style.setProperty(
          "--dashboard-background-image",
          "none"
        );
        document.documentElement.style.setProperty(
          "--dashboard-background-opacity",
          "0"
        );
        document.documentElement.dataset.backgroundActive = "false";
        document.documentElement.dataset.backgroundFallback =
          backgroundSettings.source === "none" ? "theme" : "true";
      }
      if (announce) setAppearanceStatus(statusZh, statusEn);
    }

    function normalizeThemePreset(preset) {
      const normalized = String(preset || "").toLowerCase();
      return THEME_PRESETS.has(normalized) ? normalized : "forest";
    }

    function readManualThemePreset() {
      return normalizeThemePreset(
        readPreference(THEME_PRESET_KEY) || manualThemePreset
      );
    }

    function applyGreetingThemeLink() {
      if (!themeFollowsGreeting) return;
      const tone = document.documentElement.dataset.greetingTone || "day";
      setThemePreset(
        GREETING_THEME_PRESETS[tone] || "forest",
        false,
        false
      );
    }

    function setThemeFollowsGreeting(value, persist = true) {
      const enabled = value === true || value === "true";
      themeFollowsGreeting = enabled;
      document.documentElement.dataset.themeFollowGreeting =
        enabled ? "true" : "false";
      if (themeFollowGreeting) {
        themeFollowGreeting.checked = enabled;
      }
      if (enabled) {
        manualThemePreset = readManualThemePreset();
        applyGreetingThemeLink();
      } else {
        setThemePreset(readManualThemePreset(), persist, false);
      }
      if (persist) {
        writePreference(
          THEME_FOLLOW_GREETING_KEY,
          enabled ? "true" : "false"
        );
      }
    }

    function setThemePreset(preset, persist = true, rememberManual = persist) {
      const normalized = normalizeThemePreset(preset);
      document.documentElement.dataset.themePreset = normalized;
      if (themePresetSelect) themePresetSelect.value = normalized;
      if (rememberManual) manualThemePreset = normalized;
      if (persist) {
        writePreference(THEME_PRESET_KEY, normalized);
      }
    }

    function setUploadedBackground(dataUrl, persist = true) {
      backgroundImageDataUrl = String(dataUrl || "");
      backgroundSettings.source = "upload";
      syncAppearanceControls();
      const imageStored = !persist || writePreference(
        BACKGROUND_IMAGE_KEY,
        backgroundImageDataUrl
      );
      const settingsStored = saveBackgroundSettings();
      applyAppearanceSettings(true);
      if (!imageStored || !settingsStored) {
        setAppearanceStatus(
          "图片已应用，但浏览器本地存储空间不足，刷新后可能需要重新上传。",
          "The image is applied, but browser storage is full; you may need to upload it again after reload."
        );
      }
      return imageStored && settingsStored;
    }

    function compressBackgroundImage(source, maxEdge = 1920, quality = 0.84) {
      return new Promise(resolve => {
        const image = new Image();
        image.onload = () => {
          const width = image.naturalWidth || image.width;
          const height = image.naturalHeight || image.height;
          const scale = Math.min(1, maxEdge / Math.max(width, height));
          if (scale === 1 && source.length <= MAX_BACKGROUND_IMAGE_LENGTH) {
            resolve(source);
            return;
          }
          const canvas = document.createElement("canvas");
          canvas.width = Math.max(1, Math.round(width * scale));
          canvas.height = Math.max(1, Math.round(height * scale));
          const context = canvas.getContext("2d");
          if (!context) {
            resolve(source);
            return;
          }
          context.fillStyle = "#18211c";
          context.fillRect(0, 0, canvas.width, canvas.height);
          context.drawImage(image, 0, 0, canvas.width, canvas.height);
          const output = canvas.toDataURL("image/jpeg", quality);
          if (
            output.length > MAX_BACKGROUND_IMAGE_LENGTH &&
            maxEdge > 1280
          ) {
            compressBackgroundImage(source, 1280, 0.72).then(resolve);
            return;
          }
          resolve(output);
        };
        image.onerror = () => resolve(source);
        image.src = source;
      });
    }

    function handleBackgroundUpload(event) {
      const file = event.target.files && event.target.files[0];
      if (!file) return;
      if (!file.type.startsWith("image/")) {
        setAppearanceStatus(
          "请选择图片文件。",
          "Choose an image file."
        );
        event.target.value = "";
        return;
      }
      if (file.size > 15 * 1024 * 1024) {
        setAppearanceStatus(
          "图片超过 15 MB，未应用。",
          "The image is larger than 15 MB and was not applied."
        );
        event.target.value = "";
        return;
      }
      const reader = new FileReader();
      reader.onload = () => {
        const source = String(reader.result || "");
        if (!source.startsWith("data:image/")) {
          setAppearanceStatus(
            "无法读取该图片，已保留原背景。",
            "The image could not be read; the previous background was kept."
          );
          return;
        }
        compressBackgroundImage(source).then(dataUrl => {
          setUploadedBackground(dataUrl, true);
        });
        event.target.value = "";
      };
      reader.onerror = () => {
        setAppearanceStatus(
          "读取图片失败，已保留原背景。",
          "The image could not be read; the previous background was kept."
        );
        event.target.value = "";
      };
      reader.readAsDataURL(file);
    }

    function clearBackgroundUpload() {
      backgroundImageDataUrl = "";
      removePreference(BACKGROUND_IMAGE_KEY);
      if (backgroundSettings.source === "upload") {
        applyAppearanceSettings(false);
        setAppearanceStatus(
          "已清除上传图，当前使用主题背景。",
          "The upload was cleared; using the theme background."
        );
      } else {
        setAppearanceStatus(
          "已清除本地上传图。",
          "The local upload was cleared."
        );
      }
    }

    function resetAppearance() {
      backgroundSettings = {...DEFAULT_BACKGROUND_SETTINGS};
      backgroundImageDataUrl = "";
      removePreference(BACKGROUND_IMAGE_KEY);
      setThemeFollowsGreeting(false);
      setThemePreset("forest");
      setGreetingFont("default");
      setAchievementBadgeStyle("seal");
      setShelfTitleStyle("a");
      setShelfPosition("title");
      setShelfDefaultMode("book");
      setShowTokenLabels(false);
      setBackgroundTexture("paper");
      setBackgroundMotif("none");
      resetSectionVisibility();
      setShelfMode(shelfDefaultMode);
      saveBackgroundSettings();
      syncAppearanceControls();
      applyAppearanceSettings(false);
      setAppearanceStatus(
        "已恢复默认外观。",
        "The default appearance was restored."
      );
    }

    function restoreAppearancePanel() {
      if (!appearancePanel) return;
      appearancePanel.open =
        readPreference("kindle-dashboard-appearance-panel") === "open";
      updateAppearancePosition();
    }

    function updateAppearancePosition() {
      if (!appearancePanel) return;
      const summary = appearancePanel.querySelector("summary");
      if (!summary) return;
      const bottom = Math.round(summary.getBoundingClientRect().bottom + 6);
      document.documentElement.style.setProperty(
        "--appearance-top",
        Math.max(8, bottom) + "px"
      );
    }

    function restoreAppearance() {
      setThemePreset(
        readPreference(THEME_PRESET_KEY) || "forest",
        false,
        false
      );
      setGreetingFont(
        readPreference(GREETING_FONT_KEY) || "default",
        false
      );
      setAchievementBadgeStyle(
        readPreference(ACHIEVEMENT_BADGE_STYLE_KEY) || "seal",
        false
      );
      setBackgroundTexture(
        readPreference(BACKGROUND_TEXTURE_KEY) || "paper",
        false
      );
      setBackgroundMotif(
        readPreference(BACKGROUND_MOTIF_KEY) || "none",
        false
      );
      backgroundSettings = readBackgroundSettings();
      backgroundImageDataUrl =
        readPreference(BACKGROUND_IMAGE_KEY) || "";
      manualThemePreset = readManualThemePreset();
      setThemeFollowsGreeting(
        readPreference(THEME_FOLLOW_GREETING_KEY) === "true",
        false
      );
      syncAppearanceControls();
      applyAppearanceSettings(true);
    }

    function updateThemeButton() {
      const dark = document.documentElement.dataset.theme === "dark";
      const english =
        document.documentElement.dataset.uiLanguage === "en";
      themeToggle.textContent = english
        ? (dark ? "Light" : "Dark")
        : (dark ? "浅色" : "深色");
    }

    function setUiLanguage(language, persist = true) {
      const normalized = language === "en" ? "en" : "zh";
      document.documentElement.lang =
        normalized === "en" ? "en" : "zh-CN";
      document.documentElement.dataset.uiLanguage = normalized;
      document.querySelectorAll("[data-placeholder-zh]").forEach(element => {
        element.placeholder = normalized === "en"
          ? element.dataset.placeholderEn || element.placeholder
          : element.dataset.placeholderZh || element.placeholder;
      });
      document.querySelectorAll("[data-text-zh]").forEach(element => {
        element.textContent = normalized === "en"
          ? element.dataset.textEn || element.textContent
          : element.dataset.textZh || element.textContent;
      });
      document.querySelectorAll("[data-title-zh]").forEach(element => {
        const label = normalized === "en"
          ? element.dataset.titleEn || element.dataset.titleZh
          : element.dataset.titleZh || element.dataset.titleEn;
        if (label) {
          element.title = label;
          element.setAttribute("aria-label", label);
        }
      });
      languageToggle.textContent = normalized === "en" ? "中文" : "English";
      updateThemeButton();
      updateAppearanceStatus();
      applyAchievementLevels(false);
      hideTrendTooltip();
      if (persist) writePreference("kindle-dashboard-ui-language", normalized);
    }

    function setTheme(theme, persist = true) {
      const normalized = theme === "dark" ? "dark" : "light";
      document.documentElement.dataset.theme = normalized;
      themeToggle.setAttribute(
        "aria-pressed", normalized === "dark" ? "true" : "false"
      );
      updateThemeButton();
      if (persist) writePreference("kindle-dashboard-theme", normalized);
    }

    function applyColumnVisibility(persist = false) {
      columnToggles.forEach(toggle => {
        const column = toggle.dataset.columnToggle;
        const visible = toggle.checked;
        document.querySelectorAll(
          '[data-column="' + column + '"]'
        ).forEach(element => {
          element.hidden = !visible;
        });
      });
      const visibleColumns = columnToggles.filter(
        toggle => toggle.checked
      ).length;
      table.style.minWidth = Math.max(760, visibleColumns * 145) + "px";
      document.documentElement.dataset.visibleColumns =
        String(visibleColumns);
      if (persist) {
        const hiddenColumns = columnToggles
          .filter(toggle => !toggle.checked)
          .map(toggle => toggle.dataset.columnToggle);
        writePreference(
          "kindle-dashboard-hidden-columns",
          JSON.stringify(hiddenColumns)
        );
      }
    }

    function restoreColumnVisibility() {
      let hiddenColumns = null;
      const stored = readPreference("kindle-dashboard-hidden-columns");
      if (stored !== null) {
        try {
          const parsed = JSON.parse(stored);
          if (Array.isArray(parsed)) hiddenColumns = parsed;
        } catch (error) {
          hiddenColumns = null;
        }
      }
      if (hiddenColumns === null) {
        hiddenColumns = columnToggles
          .filter(toggle => toggle.dataset.defaultHidden === "true")
          .map(toggle => toggle.dataset.columnToggle);
      }
      const hidden = new Set(hiddenColumns);
      columnToggles.forEach(toggle => {
        toggle.checked = !hidden.has(toggle.dataset.columnToggle);
      });
      applyColumnVisibility(false);
    }

    function readSectionVisibility() {
      const result = {...DEFAULT_SECTION_VISIBILITY};
      const stored = readPreference(SECTION_VISIBILITY_KEY);
      if (!stored) return result;
      try {
        const parsed = JSON.parse(stored);
        if (parsed && typeof parsed === "object") {
          for (const key of SECTION_KEYS) {
            if (typeof parsed[key] === "boolean") {
              result[key] = parsed[key];
            }
          }
        }
      } catch (error) {
        // Corrupt storage falls back to the "everything visible" default.
      }
      return result;
    }

    function saveSectionVisibility() {
      writePreference(
        SECTION_VISIBILITY_KEY,
        JSON.stringify(sectionVisibility)
      );
    }

    function sectionRootKey(key) {
      return "section" + key.charAt(0).toUpperCase() + key.slice(1);
    }

    function applySectionVisibility() {
      for (const key of SECTION_KEYS) {
        const element = reportSections[key];
        const visible =
          !minimalMode && sectionVisibility[key] !== false;
        if (element) element.hidden = !visible;
        document.documentElement.dataset[sectionRootKey(key)] =
          visible ? "visible" : "hidden";
      }
      document.documentElement.dataset.minimalMode =
        minimalMode ? "true" : "false";
      if (minimalModeToggle) minimalModeToggle.checked = minimalMode;
      sectionToggles.forEach(toggle => {
        const key = toggle.dataset.sectionToggle;
        if (Object.prototype.hasOwnProperty.call(sectionVisibility, key)) {
          toggle.checked = sectionVisibility[key] !== false;
        }
      });
    }

    function restoreSectionVisibility() {
      sectionVisibility = readSectionVisibility();
      minimalMode = readPreference(MINIMAL_MODE_KEY) === "true";
      applySectionVisibility();
    }

    function resetSectionVisibility() {
      sectionVisibility = {...DEFAULT_SECTION_VISIBILITY};
      minimalMode = false;
      saveSectionVisibility();
      writePreference(MINIMAL_MODE_KEY, "false");
      applySectionVisibility();
    }

    function normalizeShelfMode(value) {
      return value === "manga" ? "manga" : "book";
    }

    function shelfPageSize() {
      return 4;
    }

    function shelfModeLabel(mode) {
      return mode === "manga"
        ? { zh: "暂无漫画书架内容", en: "No manga shelf items" }
        : { zh: "没有匹配封面", en: "No matching covers" };
    }

    function setShelfMode(value) {
      shelfMode = normalizeShelfMode(value);
      shelfPage = 1;
      shelfModeButtons.forEach(button => {
        button.setAttribute(
          "aria-pressed",
          button.dataset.shelfMode === shelfMode ? "true" : "false"
        );
      });
      renderShelf();
    }

    function renderShelf() {
      if (!coverGrid) return;
      const allCards = Array.from(
        coverGrid.querySelectorAll(".book-cover-card[data-book-id]")
      );
      const filteredCards = allCards.filter(
        card => card.dataset.filterMatch !== "false"
      );
      const modeCards = filteredCards.filter(
        card => normalizeShelfMode(card.dataset.mode) === shelfMode
      );
      const pageSize = shelfPageSize();
      const totalPages = Math.max(1, Math.ceil(modeCards.length / pageSize));
      shelfPage = Math.min(Math.max(1, shelfPage), totalPages);
      const start = (shelfPage - 1) * pageSize;
      const visibleCards = new Set(modeCards.slice(start, start + pageSize));
      allCards.forEach(card => {
        card.hidden = !visibleCards.has(card);
      });
      if (coverCountNumber) {
        coverCountNumber.textContent = String(modeCards.length);
      }
      if (shelfPageStatus) {
        shelfPageStatus.textContent = shelfPage + "/" + totalPages;
      }
      if (shelfPrev) shelfPrev.disabled = shelfPage <= 1;
      if (shelfNext) shelfNext.disabled = shelfPage >= totalPages;
      if (coverEmpty) {
        const label = shelfModeLabel(shelfMode);
        const zh = coverEmpty.querySelector(".i18n-zh");
        const en = coverEmpty.querySelector(".i18n-en");
        if (zh) zh.textContent = label.zh;
        if (en) en.textContent = label.en;
        coverEmpty.hidden = modeCards.length !== 0;
      }
      document.documentElement.dataset.shelfMode = shelfMode;
      document.documentElement.dataset.shelfPage = String(shelfPage);
      document.documentElement.dataset.shelfPages = String(totalPages);
      document.documentElement.dataset.visibleCovers = String(
        visibleCards.size
      );
    }

    function applyFilters() {
      const query = search.value.trim().toLocaleLowerCase();
      const language = languageFilter.value;
      const countType = countFilter.value;
      const state = stateFilter.value;
      let visible = 0;
      let visibleCovers = 0;
      for (const row of tbody.querySelectorAll("tr")) {
        const match =
          (!query || row.dataset.search.includes(query)) &&
          (!language || row.dataset.language === language) &&
          (!countType || row.dataset.count === countType) &&
          (!state || row.dataset.state === state);
        row.hidden = !match;
        if (match) visible += 1;
      }
      for (const card of coverGrid.querySelectorAll(".book-cover-card")) {
        const match =
          (!query || card.dataset.search.includes(query)) &&
          (!language || card.dataset.language === language) &&
          (!countType || card.dataset.count === countType) &&
          (!state || card.dataset.state === state);
        card.dataset.filterMatch = match ? "true" : "false";
        if (match) visibleCovers += 1;
      }
      empty.hidden = visible !== 0;
      shelfPage = 1;
      renderShelf();
      document.documentElement.dataset.visibleRows = String(visible);
      document.documentElement.dataset.filteredCovers = String(visibleCovers);
    }

    function updateLayoutMetrics() {
      if (appearancePanel && appearancePanel.open) {
        updateAppearancePosition();
      }
      const viewportWidth = document.documentElement.clientWidth;
      const overflowing = Array.from(document.body.querySelectorAll("*"))
        .filter(element => !element.closest(".table-wrap"))
        .filter(element => !element.closest(".heatmap-scroll"))
        .filter(element => !element.closest(".period-chart-scroll"))
        .filter(element => !element.closest(".local-name-dialog"))
        .filter(element => {
          const rect = element.getBoundingClientRect();
          return rect.width > 0 && (
            rect.left < -1 || rect.right > viewportWidth + 1
          );
        })
        .map(element => {
          const className = typeof element.className === "string"
            ? element.className.trim().split(/\\s+/).filter(Boolean)
            : [];
          const label = element.tagName.toLowerCase() +
            (element.id ? "#" + element.id : "") +
            (className.length ? "." + className.join(".") : "");
          const rect = element.getBoundingClientRect();
          return label + "@" + Math.round(rect.left) + "," +
            Math.round(rect.right);
        });
      const coverImages = Array.from(
        document.querySelectorAll("[data-cover-image]")
      );
      document.documentElement.dataset.overflowX =
        overflowing.length ? "true" : "false";
      document.documentElement.dataset.overflowElements =
        overflowing.slice(0, 6).join("|");
      const numberElements = Array.from(
        document.querySelectorAll(
          ".metric strong, .cover-stat strong, .trend-card strong, " +
          ".achievement-value, .period-metric strong, " +
          ".period-bar-value, .period-bar-value-full, " +
          ".period-bar-value-short, .period-bar-label, " +
          ".annual-metric strong, " +
          ".insight-value, .insight-detail"
        )
      ).filter(element => {
        const rect = element.getBoundingClientRect();
        return rect.width > 0 && !element.closest("[hidden]");
      });
      const truncatedNumbers = numberElements
        .filter(element => element.scrollWidth > element.clientWidth + 1)
        .map(element => {
          const value = (element.textContent || "").trim();
          return (value || element.className || element.tagName) +
            "@" + element.clientWidth + "/" + element.scrollWidth;
        });
      document.documentElement.dataset.numberChecks =
        String(numberElements.length);
      document.documentElement.dataset.truncatedNumbers =
        truncatedNumbers.slice(0, 6).join("|");
      const overlapCandidates = Array.from(
        document.querySelectorAll(
          ".header-actions > *, .trend-metric-switch > *, " +
          ".appearance-grid > *, " +
          ".appearance-actions > *, .greeting-row > *, " +
          ".achievement-grid > *, .period-tabs > *, " +
          ".period-metrics > *, .period-bars > *, " +
          ".annual-metrics > *, " +
          ".annual-content-grid > *"
        )
      ).filter(element => {
        const rect = element.getBoundingClientRect();
        return rect.width > 0 && rect.height > 0 &&
          element.getAttribute("aria-hidden") !== "true" &&
          !element.classList.contains("period-grid-lines") &&
          !element.closest("[hidden]") &&
          !element.closest("details:not([open])");
      });
      const appearanceBody = appearancePanel?.querySelector(
        ".appearance-body"
      );
      const insideAppearanceBody = element =>
        !!appearanceBody && appearanceBody.contains(element);
      const overlaps = [];
      for (let leftIndex = 0; leftIndex < overlapCandidates.length; leftIndex += 1) {
        const left = overlapCandidates[leftIndex];
        const leftRect = left.getBoundingClientRect();
        for (
          let rightIndex = leftIndex + 1;
          rightIndex < overlapCandidates.length;
          rightIndex += 1
        ) {
          const right = overlapCandidates[rightIndex];
          // The appearance panel is an intentional floating layer. Ignore
          // overlap between its controls and page content underneath.
          if (
            insideAppearanceBody(left) !== insideAppearanceBody(right)
          ) {
            continue;
          }
          const rightRect = right.getBoundingClientRect();
          const overlapWidth = Math.min(leftRect.right, rightRect.right) -
            Math.max(leftRect.left, rightRect.left);
          const overlapHeight = Math.min(leftRect.bottom, rightRect.bottom) -
            Math.max(leftRect.top, rightRect.top);
          if (overlapWidth > 1 && overlapHeight > 1) {
            const label = element =>
              element.id ? "#" + element.id :
              element.className || element.tagName;
            overlaps.push(
              label(left) + "<>" + label(right) +
              "@" + Math.round(overlapWidth) + "x" +
              Math.round(overlapHeight)
            );
          }
        }
      }
      document.documentElement.dataset.overlapCount = String(overlaps.length);
      document.documentElement.dataset.overlapElements =
        overlaps.slice(0, 6).join("|");
      document.documentElement.dataset.coverImages =
        String(coverImages.length);
      document.documentElement.dataset.coverImagesLoaded = String(
        coverImages.filter(
          image => image.complete && image.naturalWidth > 0
        ).length
      );
      document.documentElement.dataset.tableRows = String(
        tbody.querySelectorAll("tr").length
      );
      document.documentElement.dataset.coverCards = String(
        coverGrid.querySelectorAll(".book-cover-card").length
      );
      document.documentElement.dataset.theme =
        document.documentElement.dataset.theme || "light";
      updateAchievementRootAttributes();
      updatePeriodRootAttributes();
      updateAnnualReportRootAttributes();
    }

    themeToggle.addEventListener("click", () => {
      setTheme(
        document.documentElement.dataset.theme === "dark"
          ? "light"
          : "dark"
      );
    });
    themePresetSelect.addEventListener("change", () => {
      setThemePreset(themePresetSelect.value);
    });
    if (themeFollowGreeting) {
      themeFollowGreeting.addEventListener("change", () => {
        setThemeFollowsGreeting(themeFollowGreeting.checked);
        updateLayoutMetrics();
      });
    }
    if (achievementBadgeStyleSelect) {
      achievementBadgeStyleSelect.addEventListener("change", () => {
        setAchievementBadgeStyle(achievementBadgeStyleSelect.value);
        updateLayoutMetrics();
      });
    }
    if (shelfTitleStyleSelect) {
      shelfTitleStyleSelect.addEventListener("change", () => {
        setShelfTitleStyle(shelfTitleStyleSelect.value);
        updateLayoutMetrics();
      });
    }
    if (shelfPositionSelect) {
      shelfPositionSelect.addEventListener("change", () => {
        setShelfPosition(shelfPositionSelect.value);
        updateLayoutMetrics();
      });
    }
    if (shelfDefaultModeSelect) {
      shelfDefaultModeSelect.addEventListener("change", () => {
        setShelfDefaultMode(shelfDefaultModeSelect.value);
        updateLayoutMetrics();
      });
    }
    if (showTokenLabelsToggle) {
      showTokenLabelsToggle.addEventListener("change", () => {
        setShowTokenLabels(showTokenLabelsToggle.checked);
        updateLayoutMetrics();
      });
    }
    if (achievementEnabledToggle) {
      achievementEnabledToggle.addEventListener("change", () => {
        setAchievementEnabled(achievementEnabledToggle.checked);
        updateLayoutMetrics();
      });
    }
    if (mangaVolumesToggle) {
      mangaVolumesToggle.addEventListener("change", () => {
        setMangaVolumeAchievement(mangaVolumesToggle.checked);
        updateLayoutMetrics();
      });
    }
    if (greetingFontSelect) {
      greetingFontSelect.addEventListener("change", () => {
        setGreetingFont(greetingFontSelect.value);
        updateLayoutMetrics();
      });
    }
    if (backgroundTextureSelect) {
      backgroundTextureSelect.addEventListener("change", () => {
        setBackgroundTexture(backgroundTextureSelect.value);
        updateLayoutMetrics();
      });
    }
    languageToggle.addEventListener("click", () => {
      setUiLanguage(
        document.documentElement.dataset.uiLanguage === "en"
          ? "zh"
          : "en"
      );
    });
    columnToggles.forEach(toggle => {
      toggle.addEventListener("change", () => {
        applyColumnVisibility(true);
        updateLayoutMetrics();
      });
    });
    shelfModeButtons.forEach(button => {
      button.addEventListener("click", () => {
        setShelfMode(button.dataset.shelfMode);
        setShelfFallbackOpen(false);
        updateLayoutMetrics();
      });
    });
    if (shelfFallbackButton) {
      shelfFallbackButton.addEventListener("click", event => {
        event.stopPropagation();
        setShelfFallbackOpen(shelfFallbackMenu?.hidden !== false);
      });
    }
    shelfPrev?.addEventListener("click", () => {
      shelfPage -= 1;
      renderShelf();
      updateLayoutMetrics();
    });
    shelfNext?.addEventListener("click", () => {
      shelfPage += 1;
      renderShelf();
      updateLayoutMetrics();
    });
    sectionToggles.forEach(toggle => {
      toggle.addEventListener("change", () => {
        const key = toggle.dataset.sectionToggle;
        if (!Object.prototype.hasOwnProperty.call(sectionVisibility, key)) {
          return;
        }
        // The individual choice is always stored; minimal mode only controls
        // whether it is currently applied.
        sectionVisibility[key] = toggle.checked;
        saveSectionVisibility();
        applySectionVisibility();
        updateLayoutMetrics();
      });
    });
    if (minimalModeToggle) {
      minimalModeToggle.addEventListener("change", () => {
        minimalMode = minimalModeToggle.checked;
        writePreference(MINIMAL_MODE_KEY, minimalMode ? "true" : "false");
        applySectionVisibility();
        updateLayoutMetrics();
      });
    }
    backgroundSource.addEventListener("change", () => {
      backgroundSettings.source = BACKGROUND_SOURCES.has(
        backgroundSource.value
      ) ? backgroundSource.value : "none";
      if (
        backgroundSettings.source === "fixed" &&
        !backgroundSettings.bookId
      ) {
        const firstAvailable = Array.from(backgroundBook.options).find(
          option => !option.disabled && option.value
        );
        if (firstAvailable) backgroundSettings.bookId = firstAvailable.value;
      }
      saveBackgroundSettings();
      syncAppearanceControls();
      applyAppearanceSettings(true);
      updateLayoutMetrics();
    });
    backgroundBook.addEventListener("change", () => {
      backgroundSettings.bookId = backgroundBook.value;
      saveBackgroundSettings();
      applyAppearanceSettings(true);
    });
    backgroundUpload.addEventListener("change", handleBackgroundUpload);
    backgroundBlur.addEventListener("input", () => {
      backgroundSettings.blur = clampPreference(
        backgroundBlur.value, 12, 0, 30
      );
      saveBackgroundSettings();
      syncAppearanceControls();
      applyAppearanceSettings(false);
    });
    backgroundOverlay.addEventListener("input", () => {
      backgroundSettings.overlay = clampPreference(
        backgroundOverlay.value, 62, 0, 90
      );
      saveBackgroundSettings();
      syncAppearanceControls();
      applyAppearanceSettings(false);
    });
    backgroundOpacity.addEventListener("input", () => {
      backgroundSettings.opacity = clampPreference(
        backgroundOpacity.value, 72, 10, 100
      );
      saveBackgroundSettings();
      syncAppearanceControls();
      applyAppearanceSettings(false);
    });
    backgroundClear.addEventListener("click", clearBackgroundUpload);
    appearanceReset.addEventListener("click", resetAppearance);
    if (greetingMode) {
      greetingMode.addEventListener("change", () => {
        updateGreeting(true);
        updateLayoutMetrics();
      });
    }
    if (userNameInput) {
      userNameInput.addEventListener("change", () => {
        setUserName(userNameInput.value);
        updateLayoutMetrics();
      });
      userNameInput.addEventListener("keydown", event => {
        if (event.key === "Enter") {
          userNameInput.blur();
        }
      });
    }
    if (userNameClear) {
      userNameClear.addEventListener("click", () => {
        setUserName("");
        setAppearanceStatus(
          "已清除本机用户名。",
          "The local user name was cleared."
        );
      });
    }
    if (userNameSave) {
      userNameSave.addEventListener("click", () => {
        setUserName(userNameDialogInput?.value || "");
        hideUserNameDialog();
      });
    }
    if (userNameSkip) {
      userNameSkip.addEventListener("click", () => {
        setUserName("");
        hideUserNameDialog();
      });
    }
    if (userNameDialogInput) {
      userNameDialogInput.addEventListener("keydown", event => {
        if (event.key === "Enter") {
          event.preventDefault();
          setUserName(userNameDialogInput.value);
          hideUserNameDialog();
        }
      });
    }
    trendMetricButtons.forEach(button => {
      button.addEventListener("click", () => {
        setTrendMetric(button.dataset.trendMetricButton);
        updateLayoutMetrics();
      });
    });
    if (trendDetails) {
      trendDetails.addEventListener("toggle", () => {
        writePreference(
          TREND_DETAILS_KEY,
          trendDetails.open ? "open" : "closed"
        );
        updateTrendRootAttributes();
        updateLayoutMetrics();
      });
    }
    if (appearancePanel) {
      appearancePanel.addEventListener("toggle", () => {
        writePreference(
          "kindle-dashboard-appearance-panel",
          appearancePanel.open ? "open" : "closed"
        );
        updateAppearancePosition();
      });
    }
    if (achievementPanel) {
      achievementPanel.addEventListener("toggle", () => {
        writePreference(
          ACHIEVEMENT_PANEL_KEY,
          achievementPanel.open ? "open" : "closed"
        );
        updateAchievementRootAttributes();
        updateLayoutMetrics();
      });
    }
    if (starSkyOpen) {
      starSkyOpen.addEventListener("click", openStarSky);
    }
    if (starSkyClose) {
      starSkyClose.addEventListener("click", closeStarSky);
    }
    if (starSkyDialog) {
      starSkyDialog.addEventListener("click", event => {
        if (event.target === starSkyDialog) closeStarSky();
      });
      starSkyDialog.addEventListener("close", () => {
        starSkyOpen?.focus();
      });
    }
    if (shareExport) {
      shareExport.addEventListener("click", openShareDialog);
    }
    if (printReport) {
      printReport.addEventListener("click", () => window.print());
    }
    if (shareClose) {
      shareClose.addEventListener("click", closeShareDialog);
    }
    if (shareDialog) {
      shareDialog.addEventListener("click", event => {
        if (event.target === shareDialog) closeShareDialog();
      });
      shareDialog.addEventListener("close", () => {
        shareExport?.focus();
      });
    }
    if (shareGenerate) {
      shareGenerate.addEventListener("click", generateShareImages);
    }
    starSkyStars.forEach(star => {
      star.addEventListener("pointerenter", () => showStarSkyDetail(star));
      star.addEventListener("focus", () => showStarSkyDetail(star));
      star.addEventListener("click", () => showStarSkyDetail(star));
      star.addEventListener("keydown", event => {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          showStarSkyDetail(star);
        }
      });
    });
    if (periodPanel) {
      periodPanel.addEventListener("toggle", () => {
        writePreference(
          PERIOD_PANEL_KEY,
          periodPanel.open ? "open" : "closed"
        );
        updatePeriodRootAttributes();
        updateLayoutMetrics();
      });
    }
    if (annualReportPanel) {
      annualReportPanel.addEventListener("toggle", () => {
        writePreference(
          ANNUAL_REPORT_PANEL_KEY,
          annualReportPanel.open ? "open" : "closed"
        );
        updateAnnualReportRootAttributes();
        updateLayoutMetrics();
      });
    }
    periodTabs.forEach(tab => {
      tab.addEventListener("click", () => {
        setPeriodView(tab.dataset.periodTab);
        updateLayoutMetrics();
      });
    });
    document.addEventListener("click", event => {
      if (
        columnPicker.open &&
        !columnPicker.contains(event.target)
      ) {
        columnPicker.open = false;
      }
      if (
        appearancePanel &&
        appearancePanel.open &&
        !appearancePanel.contains(event.target)
      ) {
        appearancePanel.open = false;
      }
      if (
        shelfFallbackWrap &&
        shelfFallbackMenu &&
        !shelfFallbackMenu.hidden &&
        !shelfFallbackWrap.contains(event.target)
      ) {
        setShelfFallbackOpen(false);
      }
    });
    document.addEventListener("keydown", event => {
      const target = event.target;
      const typing = target && (
        target.tagName === "INPUT" ||
        target.tagName === "TEXTAREA" ||
        target.tagName === "SELECT" ||
        target.isContentEditable
      );
      if (
        !typing &&
        (event.key === "ArrowLeft" || event.key === "ArrowRight") &&
        coverGrid &&
        !coverGrid.closest("section")?.hidden
      ) {
        const totalPages =
          Number(document.documentElement.dataset.shelfPages) || 1;
        if (event.key === "ArrowLeft" && shelfPage > 1) {
          event.preventDefault();
          shelfPage -= 1;
          renderShelf();
          updateLayoutMetrics();
        } else if (
          event.key === "ArrowRight" &&
          shelfPage < totalPages
        ) {
          event.preventDefault();
          shelfPage += 1;
          renderShelf();
          updateLayoutMetrics();
        }
      }
      if (
        event.key === "Escape" &&
        appearancePanel &&
        appearancePanel.open
      ) {
        appearancePanel.open = false;
        appearancePanel.querySelector("summary")?.focus();
      }
      if (
        event.key === "Escape" &&
        shelfFallbackMenu &&
        !shelfFallbackMenu.hidden
      ) {
        setShelfFallbackOpen(false);
        shelfFallbackButton?.focus();
      }
    });
    search.addEventListener("input", applyFilters);
    languageFilter.addEventListener("change", applyFilters);
    countFilter.addEventListener("change", applyFilters);
    stateFilter.addEventListener("change", applyFilters);
    window.addEventListener("load", () => {
      setUiLanguage(
        readPreference("kindle-dashboard-ui-language") || "zh",
        false
      );
      const storedTheme = readPreference("kindle-dashboard-theme");
      setTheme(storedTheme || "light", false);
      setTrendMetric(readPreference(TREND_METRIC_KEY) || "time", false);
      restoreTrendDetails();
      initializeTrendPoints();
      restoreUserName();
      restoreGreeting();
      restoreColumnVisibility();
      restoreAchievementPanel();
      applyAchievementLevels(true);
      restorePeriodPanel();
      restorePeriodView();
      restoreAnnualReportPanel();
      restoreAppearancePanel();
      restoreAppearance();
      restoreSectionVisibility();
      restoreShelfPreferences();
      applyFilters();
      updateLayoutMetrics();
    });
    window.addEventListener("beforeprint", preparePeriodPrint);
    window.addEventListener("afterprint", restoreAfterPeriodPrint);
    window.addEventListener("resize", () => {
      renderShelf();
      updateLayoutMetrics();
    });
    window.addEventListener("focus", () => updateGreeting(false));
    document.addEventListener("visibilitychange", () => {
      if (!document.hidden) updateGreeting(false);
    });
    window.setInterval(() => {
      if (
        greetingMode &&
        normalizeGreetingMode(greetingMode.value) === "auto"
      ) {
        updateGreeting(false);
      }
    }, 60000);

    table.querySelectorAll("th").forEach((header, index) => {
      header.addEventListener("click", () => {
        const rows = Array.from(tbody.querySelectorAll("tr"));
        const descending = header.dataset.direction !== "desc";
        header.dataset.direction = descending ? "desc" : "asc";
        rows.sort((left, right) => {
          const a = left.children[index];
          const b = right.children[index];
          const type = header.dataset.type;
          const av = type === "number"
            ? Number(a.dataset.value ?? -1)
            : a.textContent.trim();
          const bv = type === "number"
            ? Number(b.dataset.value ?? -1)
            : b.textContent.trim();
          const result = type === "number"
            ? (av || -1) - (bv || -1)
            : String(av).localeCompare(String(bv), "zh-CN");
          return descending ? -result : result;
        });
        rows.forEach(row => tbody.appendChild(row));
      });
    });
  </script>
</body>
</html>
"""

    chart_html = (
        "\n".join(chart_items)
        or (
            '<div class="empty">'
            + bilingual_text(
                "没有可显示的阅读时长记录",
                "No reading-time records to display",
            )
            + "</div>"
        )
    )
    progress_html = (
        "\n".join(progress_cards)
        or (
            '<div class="empty">'
            + bilingual_text(
                "没有可显示的进度记录",
                "No progress records to display",
            )
            + "</div>"
        )
    )
    cover_html = (
        "\n".join(cover_cards)
        or (
            '<div class="empty">'
            + bilingual_text("没有书封记录", "No cover records")
            + "</div>"
        )
    )
    rows_html = (
        "\n".join(table_rows)
        or (
            '<tr><td colspan="11" class="empty">'
            + bilingual_text("没有书籍记录", "No book records")
            + "</td></tr>"
        )
    )
    insight_items = build_insights(useful_rows)
    insights_html = insight_cards_html(
        [item for item in insight_items if not item.get("advanced")],
        cover_urls=resolved_cover_urls,
    ) or (
        '<article class="insight" data-insight-key="empty" '
        'data-insight-layout="plain"><span>'
        + bilingual_text("数据", "Data")
        + '</span><strong>'
        + bilingual_text("暂无", "No data")
        + "</strong></article>"
    )
    resolved_history_path = (
        history_path
        if history_path is not None
        else csv_path.parent / HISTORY_FILENAME
    )
    resolved_covers_dir = (
        covers_dir
        if covers_dir is not None
        else csv_path.parent / COVER_DIRECTORY_NAME
    )
    resolved_xlsx_path = (
        Path(xlsx_path)
        if xlsx_path is not None
        else csv_path.with_suffix(".xlsx")
    )
    cover_count = len(resolved_cover_urls)
    public_trend_data = dict(trend_data)
    public_trend_data.pop("ocr_transition", None)
    public_trend_data.pop("ocr_effective_enabled", None)
    trend_json = json.dumps(
        public_trend_data,
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    ).replace("</", "<\\/")

    replacements = {
        "__SOURCE__": html.escape(source_path),
        "__GENERATED__": html.escape(generated),
        "__UI_VARIANT__": "text" if text_only_ui else "ocr",
        "__BOOKS__": str(summary["books"]),
        "__READING_RECORD_BOOKS__": str(
            summary.get("reading_record_books", 0)
        ),
        "__TIME__": html.escape(format_duration(summary["total_time_ms"])),
        "__WORDS__": html.escape(
            format_words(
                parse_int(summary.get("effective_words_read"))
                if summary.get("effective_words_read") is not None
                else summary["total_words"]
            )
            if summary.get("word_books")
            else "--"
        ),
        "__WPM__": (
            f'{summary["weighted_wpm"]:.1f}'
            if summary["weighted_wpm"] is not None
            else "--"
        ),
        "__PROGRESS__": (
            f'{summary["average_word_progress"]:.1f}%'
            if summary["average_word_progress"] is not None
            else "--"
        ),
        "__ANNOTATIONS__": str(summary["annotations"]),
        "__COVER_COUNT__": str(cover_count),
        "__COVER_SUMMARY__": bilingual_text(
            f"{cover_count}/{len(useful_rows)} 本使用真实封面",
            f"{cover_count}/{len(useful_rows)} books use real covers",
        ),
        "__COVER_CARDS__": cover_html,
        "__STAR_SKY__": star_sky_html,
        "__INSIGHTS__": insights_html,
        "__ACHIEVEMENT_SUMMARY__": bilingual_text(
            f"{achievement_data['stars_unlocked']}/{achievement_data['stars_total']} 颗星已点亮",
            f"{achievement_data['stars_unlocked']}/{achievement_data['stars_total']} stars unlocked",
        ),
        "__ACHIEVEMENTS__": achievements_html,
        "__ACHIEVEMENT_CRITERIA_ZH__": html.escape(achievement_criteria_zh),
        "__ACHIEVEMENT_CRITERIA_EN__": html.escape(achievement_criteria_en),
        "__PERIOD_REPORTS__": period_html,
        "__ANNUAL_REPORT__": annual_html,
        "__CHART__": chart_html,
        "__PROGRESS_CARDS__": progress_html,
        "__TREND_CARDS__": trend_sections["cards"],
        "__TREND_MOVEMENTS__": trend_sections["movements"],
        "__TREND_DAILY__": trend_sections["daily"],
        "__TREND_CHART__": trend_sections["chart"],
        "__TREND_SNAPSHOT_TABLE__": trend_sections["snapshot_table"],
        "__TREND_PANEL_ATTRS__": trend_sections["panel_attrs"],
        "__TREND_HEATMAP__": heatmap_html,
        "__TREND_DATA__": trend_json,
        "__ROWS__": rows_html,
        "__CSV__": html.escape(str(csv_path)),
        "__XLSX_NAME__": html.escape(
            resolved_xlsx_path.name,
            quote=True,
        ),
        "__XLSX__": html.escape(str(resolved_xlsx_path)),
        "__COVERS_DIR__": html.escape(str(resolved_covers_dir)),
        "__HISTORY__": html.escape(str(resolved_history_path)),
    }
    for token, value in replacements.items():
        html_text = html_text.replace(token, value)
    return html_text


def application_root() -> Path:
    """Return the writable directory used for executable data and outputs."""

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _atomic_write_text(
    path: Path,
    document: str,
    encoding: str = "utf-8",
) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path: Optional[Path] = None
    try:
        with tempfile.NamedTemporaryFile(
            "w",
            encoding=encoding,
            newline="",
            dir=path.parent,
            prefix=path.name + ".",
            suffix=".tmp",
            delete=False,
        ) as handle:
            temporary_path = Path(handle.name)
            handle.write(document)
        os.replace(temporary_path, path)
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def backup_corrupt_file(path: Path) -> Optional[Path]:
    if not path.is_file():
        return None

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = path.with_name(f"{path.name}.corrupt-{stamp}")
    counter = 1
    while backup_path.exists():
        backup_path = path.with_name(
            f"{path.name}.corrupt-{stamp}-{counter}"
        )
        counter += 1
    try:
        shutil.copy2(path, backup_path)
    except OSError:
        return None
    return backup_path


def normalize_settings(
    value: Any,
) -> Tuple[Dict[str, Any], List[str]]:
    settings = dict(DEFAULT_SETTINGS)
    warnings: List[str] = []
    if not isinstance(value, dict):
        return settings, ["settings root must be a JSON object"]

    for key in (
        "kindle_path",
        "output_dir",
        "language",
        "book_modes_file",
        "achievement_profile_file",
        "ocr_language",
        "ocr_engine",
        "ocr_mode",
        "ocr_cache_dir",
    ):
        if key not in value:
            continue
        if not isinstance(value[key], str):
            warnings.append(f"settings {key} must be a string")
            continue
        settings[key] = value[key].strip()

    if settings["language"] not in {"auto", "en", "zh", "ja"}:
        warnings.append(
            "settings language must be auto, en, zh, or ja"
        )
        settings["language"] = DEFAULT_SETTINGS["language"]

    for key, choices, label in (
        ("ocr_language", {"auto", "en", "zh", "ja"}, "auto, en, zh, or ja"),
        ("ocr_engine", {"auto", "manga_ocr", "paddleocr", "tesseract"},
         "auto, manga_ocr, paddleocr, or tesseract"),
    ):
        if key not in value:
            continue
        normalized = str(settings.get(key) or "").strip().casefold()
        if normalized not in choices:
            warnings.append(f"settings {key} must be {label}")
            settings[key] = DEFAULT_SETTINGS[key]
        else:
            settings[key] = normalized

    for key in (
        "exact_word_count",
        "open_html",
        "manga_ocr",
        "manga_classify",
        "manga_volume_achievement",
    ):
        if key not in value:
            continue
        if not isinstance(value[key], bool):
            warnings.append(f"settings {key} must be true or false")
            continue
        settings[key] = value[key]

    if "history_limit" in value:
        history_limit = value["history_limit"]
        if (
            isinstance(history_limit, bool)
            or not isinstance(history_limit, int)
            or history_limit < 0
        ):
            warnings.append(
                "settings history_limit must be a non-negative integer"
            )
            settings["history_limit"] = DEFAULT_HISTORY_LIMIT
        else:
            settings["history_limit"] = history_limit

    if "completed_book_threshold_percent" in value:
        threshold = value["completed_book_threshold_percent"]
        if isinstance(threshold, bool) or not isinstance(
            threshold, (int, float)
        ):
            warnings.append(
                "settings completed_book_threshold_percent must be a number "
                "between 0 and 100"
            )
            settings["completed_book_threshold_percent"] = (
                DEFAULT_COMPLETED_BOOK_THRESHOLD_PERCENT
            )
        else:
            settings["completed_book_threshold_percent"] = (
                normalized_completed_threshold(threshold)
            )

    if "ocr_dpi" in value:
        dpi = value["ocr_dpi"]
        if isinstance(dpi, bool) or not isinstance(dpi, int):
            warnings.append("settings ocr_dpi must be an integer")
            settings["ocr_dpi"] = DEFAULT_SETTINGS["ocr_dpi"]
        elif not 200 <= dpi <= 400:
            warnings.append("settings ocr_dpi must be between 200 and 400")
            settings["ocr_dpi"] = DEFAULT_SETTINGS["ocr_dpi"]
        else:
            settings["ocr_dpi"] = dpi

    if "ocr_workers" in value:
        workers = value["ocr_workers"]
        if isinstance(workers, bool) or not isinstance(workers, int) or workers < 0:
            warnings.append("settings ocr_workers must be a non-negative integer")
            settings["ocr_workers"] = DEFAULT_SETTINGS["ocr_workers"]
        else:
            settings["ocr_workers"] = workers

    if "ocr_sample_pages" in value:
        sample_pages = value["ocr_sample_pages"]
        if (
            isinstance(sample_pages, bool)
            or not isinstance(sample_pages, int)
            or not 5 <= sample_pages <= 8
        ):
            warnings.append("settings ocr_sample_pages must be between 5 and 8")
            settings["ocr_sample_pages"] = DEFAULT_SETTINGS["ocr_sample_pages"]
        else:
            settings["ocr_sample_pages"] = sample_pages

    if "ocr_mode" in value:
        mode = str(settings.get("ocr_mode") or "").strip().casefold()
        if mode not in {"classify", "sample", "full"}:
            warnings.append("settings ocr_mode must be classify, sample, or full")
            settings["ocr_mode"] = DEFAULT_SETTINGS["ocr_mode"]
        else:
            settings["ocr_mode"] = mode

    if "ocr_min_confidence" in value:
        minimum = value["ocr_min_confidence"]
        if isinstance(minimum, bool) or not isinstance(
            minimum, (int, float)
        ):
            warnings.append(
                "settings ocr_min_confidence must be a number between 0 and 1"
            )
            settings["ocr_min_confidence"] = DEFAULT_SETTINGS[
                "ocr_min_confidence"
            ]
        else:
            number = float(minimum)
            if not 0.0 <= number <= 1.0:
                warnings.append(
                    "settings ocr_min_confidence must be between 0 and 1"
                )
                settings["ocr_min_confidence"] = DEFAULT_SETTINGS[
                    "ocr_min_confidence"
                ]
            else:
                settings["ocr_min_confidence"] = round(number, 4)

    return settings, warnings


def load_settings(
    settings_path: Path,
    create_default: bool = True,
) -> Tuple[Dict[str, Any], List[str]]:
    settings_path = Path(settings_path)
    if not settings_path.is_file():
        settings = dict(DEFAULT_SETTINGS)
        warnings: List[str] = []
        if create_default:
            try:
                _atomic_write_text(
                    settings_path,
                    json.dumps(
                        settings,
                        ensure_ascii=False,
                        indent=2,
                        sort_keys=True,
                    )
                    + "\n",
                )
            except OSError as exc:
                warnings.append(
                    f"settings file could not be created: {exc}"
                )
        return settings, warnings

    try:
        with settings_path.open("r", encoding="utf-8-sig") as handle:
            decoded = json.load(handle)
    except (OSError, ValueError) as exc:
        backup_path = backup_corrupt_file(settings_path)
        warning = f"settings file could not be read: {exc}"
        if backup_path is not None:
            warning += f"; backup: {backup_path}"
        if create_default:
            try:
                _atomic_write_text(
                    settings_path,
                    json.dumps(
                        DEFAULT_SETTINGS,
                        ensure_ascii=False,
                        indent=2,
                        sort_keys=True,
                    )
                    + "\n",
                )
            except OSError as exc:
                warning += f"; settings could not be rebuilt: {exc}"
        return dict(DEFAULT_SETTINGS), [warning]

    return normalize_settings(decoded)


def merge_settings(
    args: argparse.Namespace,
    settings: Dict[str, Any],
) -> Dict[str, Any]:
    return {
        "kindle_path": (
            args.path if args.path else settings.get("kindle_path", "")
        ),
        "output_dir": (
            args.output_dir
            if args.output_dir
            else settings.get("output_dir", "")
        ),
        "language": (
            args.word_count_language
            if args.word_count_language
            else settings.get("language", "auto")
        ),
        "exact_word_count": (
            args.exact_word_count
            if args.exact_word_count is not None
            else settings.get("exact_word_count", True)
        ),
        "open_html": (
            args.open_html
            if args.open_html is not None
            else settings.get("open_html", True)
        ),
        "history_limit": (
            args.history_limit
            if args.history_limit is not None
            else settings.get("history_limit", DEFAULT_HISTORY_LIMIT)
        ),
        "book_modes_file": (
            args.book_modes_file
            if args.book_modes_file
            else settings.get("book_modes_file", DEFAULT_BOOK_MODES_FILENAME)
        ),
        "achievement_profile_file": (
            args.achievement_profile_file
            if args.achievement_profile_file
            else settings.get(
                "achievement_profile_file",
                DEFAULT_ACHIEVEMENT_PROFILE_FILENAME,
            )
        ),
        "completed_book_threshold_percent": (
            args.completed_book_threshold_percent
            if args.completed_book_threshold_percent is not None
            else settings.get(
                "completed_book_threshold_percent",
                DEFAULT_COMPLETED_BOOK_THRESHOLD_PERCENT,
            )
        ),
        "manga_ocr": (
            getattr(args, "manga_ocr", None)
            if getattr(args, "manga_ocr", None) is not None
            else settings.get(
                "manga_ocr",
                DEFAULT_SETTINGS["manga_ocr"],
            )
        ),
        "manga_classify": (
            getattr(args, "manga_classify", None)
            if getattr(args, "manga_classify", None) is not None
            else settings.get(
                "manga_classify",
                DEFAULT_SETTINGS["manga_classify"],
            )
        ),
        "manga_volume_achievement": (
            getattr(args, "manga_volume_achievement", None)
            if getattr(args, "manga_volume_achievement", None) is not None
            else settings.get(
                "manga_volume_achievement",
                DEFAULT_SETTINGS["manga_volume_achievement"],
            )
        ),
        "ocr_language": (
            getattr(args, "ocr_language", "")
            if getattr(args, "ocr_language", "")
            else settings.get("ocr_language", "auto")
        ),
        "ocr_engine": (
            getattr(args, "ocr_engine", "")
            if getattr(args, "ocr_engine", "")
            else settings.get("ocr_engine", "auto")
        ),
        "ocr_dpi": (
            getattr(args, "ocr_dpi", None)
            if getattr(args, "ocr_dpi", None) is not None
            else settings.get("ocr_dpi", DEFAULT_SETTINGS["ocr_dpi"])
        ),
        "ocr_mode": (
            getattr(args, "ocr_mode", "")
            if getattr(args, "ocr_mode", "")
            else settings.get("ocr_mode", DEFAULT_SETTINGS["ocr_mode"])
        ),
        "ocr_workers": (
            getattr(args, "ocr_workers", None)
            if getattr(args, "ocr_workers", None) is not None
            else settings.get("ocr_workers", DEFAULT_SETTINGS["ocr_workers"])
        ),
        "ocr_sample_pages": (
            getattr(args, "ocr_sample_pages", None)
            if getattr(args, "ocr_sample_pages", None) is not None
            else settings.get(
                "ocr_sample_pages",
                DEFAULT_SETTINGS["ocr_sample_pages"],
            )
        ),
        "ocr_min_confidence": (
            getattr(args, "ocr_min_confidence", None)
            if getattr(args, "ocr_min_confidence", None) is not None
            else settings.get(
                "ocr_min_confidence",
                DEFAULT_SETTINGS["ocr_min_confidence"],
            )
        ),
        "ocr_cache_dir": (
            getattr(args, "ocr_cache_dir", "")
            if getattr(args, "ocr_cache_dir", "")
            else settings.get(
                "ocr_cache_dir",
                DEFAULT_OCR_CACHE_DIRNAME,
            )
        ),
    }


def check_free_space(
    path: Path,
    required_bytes: int = MINIMUM_FREE_SPACE_BYTES,
) -> Tuple[bool, Optional[int]]:
    probe = Path(path)
    while not probe.exists() and probe.parent != probe:
        probe = probe.parent
    try:
        free_bytes = shutil.disk_usage(probe).free
    except OSError:
        return True, None
    return free_bytes >= required_bytes, free_bytes


def format_bytes(value: int) -> str:
    amount = float(value)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if abs(amount) < 1024.0 or unit == "TB":
            return f"{amount:.1f} {unit}"
        amount /= 1024.0
    return f"{value} B"


def open_report(path: Path) -> None:
    if os.name == "nt" and hasattr(os, "startfile"):
        os.startfile(str(path))  # type: ignore[attr-defined]
        return
    webbrowser.open(path.resolve().as_uri())


def pause_after_error(raw_args: Sequence[str], no_wait: bool) -> None:
    if (
        os.name != "nt"
        or no_wait
        or raw_args
        or not getattr(sys, "frozen", False)
        or not sys.stdin.isatty()
    ):
        return
    try:
        input("Press Enter to close...")
    except EOFError:
        pass


def build_arg_parser(variant: Optional[str] = None) -> argparse.ArgumentParser:
    app_variant = variant or runtime_variant()
    parser = argparse.ArgumentParser(
        description=(
            "Scan a connected Kindle and create a visual HTML report. "
            "When --csv and --html are supplied, render an existing CSV only."
        )
    )
    parser.add_argument(
        "--csv",
        default="",
        help="Input CSV path for report-only mode.",
    )
    parser.add_argument(
        "--html",
        default="",
        help="Output HTML path for report-only mode.",
    )
    parser.add_argument(
        "--path",
        default="",
        help=(
            "Kindle mount root or documents folder. Auto-detected when omitted."
        ),
    )
    parser.add_argument(
        "--output-dir",
        default="",
        help=(
            "Directory for default CSV, HTML, history, cover, and cache files. "
            "Default: the executable directory."
        ),
    )
    parser.add_argument(
        "--output-csv",
        default="",
        help="Override the full-workflow CSV output path.",
    )
    parser.add_argument(
        "--output-html",
        default="",
        help="Override the full-workflow HTML output path.",
    )
    parser.add_argument(
        "--history-limit",
        type=int,
        default=None,
        help=(
            "Maximum number of history snapshots to retain; "
            "0 keeps all snapshots. Default: settings.json, then 0."
        ),
    )
    parser.add_argument(
        "--settings",
        default="",
        help=(
            "Settings JSON path. Default: settings.json beside the executable "
            "or source script."
        ),
    )
    exact_group = parser.add_mutually_exclusive_group()
    exact_group.add_argument(
        "--exact-word-count",
        dest="exact_word_count",
        action="store_true",
        default=None,
        help="Enable exact full-book word counting for this run.",
    )
    exact_group.add_argument(
        "--no-exact-word-count",
        dest="exact_word_count",
        action="store_false",
        help="Skip Calibre conversion and exact full-book word counting.",
    )
    parser.add_argument(
        "--word-count-language",
        choices=("auto", "en", "zh", "ja"),
        default=None,
        help=(
            "Tokenizer language for exact word counts. "
            "Default: settings.json, then auto."
        ),
    )
    if app_variant != "text":
        ocr_group = parser.add_mutually_exclusive_group()
        ocr_group.add_argument(
            "--manga-ocr",
            dest="manga_ocr",
            action="store_true",
            default=None,
            help=(
                "Enable offline OCR estimates for reading_mode=manga rows "
                "after scanning. CLI overrides settings.json."
            ),
        )
        ocr_group.add_argument(
            "--no-manga-ocr",
            dest="manga_ocr",
            action="store_false",
            help=(
                "Disable manga OCR estimates for this run. CLI overrides "
                "settings.json."
            ),
        )
        parser.add_argument(
            "--ocr-language",
            choices=("auto", "ja", "zh", "en"),
            default=None,
            help=(
                "OCR language. auto is conservative and may report "
                "needs-language instead of guessing. Default: settings, "
                "then auto."
            ),
        )
        parser.add_argument(
            "--ocr-engine",
            choices=("auto", "manga_ocr", "paddleocr", "tesseract"),
            default=None,
            help=(
                "OCR adapter. auto prefers manga_ocr/PaddleOCR for Japanese, "
                "PaddleOCR for Chinese, and Tesseract/PaddleOCR for English."
            ),
        )
        parser.add_argument(
            "--ocr-dpi",
            type=int,
            default=None,
            help=(
                "Page render DPI for manga OCR. Allowed range: 200-400; "
                "sampling uses 200 by default."
            ),
        )
        parser.add_argument(
            "--ocr-mode",
            choices=("classify", "sample", "full"),
            default=None,
            help=(
                "classify samples a few pages to determine book/manga; "
                "sample extrapolates words/pages; full OCRs every page."
            ),
        )
        parser.add_argument(
            "--ocr-workers",
            type=int,
            default=None,
            help=(
                "Parallel worker count for rendering/OCR. "
                "Default: min(4, CPU cores)."
            ),
        )
        parser.add_argument(
            "--ocr-sample-pages",
            type=int,
            default=None,
            help="Sample pages for classification/sample mode (5-8).",
        )
        parser.add_argument(
            "--ocr-min-confidence",
            type=float,
            default=None,
            help=(
                "Mean OCR confidence threshold from 0 to 1. Below it the "
                "row is marked ocr-low-confidence."
            ),
        )
        parser.add_argument(
            "--ocr-cache-dir",
            default="",
            help=(
                "Project-side OCR cache directory. Default: _ocr beside the "
                "output files."
            ),
        )
    open_group = parser.add_mutually_exclusive_group()
    open_group.add_argument(
        "--open",
        dest="open_html",
        action="store_true",
        default=None,
        help="Open the generated report for this run.",
    )
    open_group.add_argument(
        "--no-open",
        dest="open_html",
        action="store_false",
        help="Do not open the generated report in the default browser.",
    )
    parser.add_argument(
        "--no-wait",
        action="store_true",
        help="Do not pause after an error when launched without arguments.",
    )
    parser.add_argument(
        "--build-exe",
        action="store_true",
        help="Build Windows executables with PyInstaller and exit.",
    )
    parser.add_argument(
        "--dist-dir",
        default="",
        help="Executable output directory for --build-exe. Default: .\\dist.",
    )
    parser.add_argument(
        "--source",
        default="",
        help="Kindle path shown in the report header.",
    )
    parser.add_argument(
        "--history",
        default="",
        help=(
            "Snapshot JSON path. Default: reading_history.json beside the CSV."
        ),
    )
    if app_variant != "text":
        parser.add_argument(
            "--book-modes",
            dest="book_modes_file",
            default="",
            help=(
                "JSON mapping for book or manga reading modes. "
                "Default: book_modes.json beside the CSV/output directory."
            ),
        )
    else:
        parser.set_defaults(book_modes_file="")
    if app_variant != "text":
        classify_group = parser.add_mutually_exclusive_group()
        classify_group.add_argument(
            "--manga-classify",
            dest="manga_classify",
            action="store_true",
            default=None,
            help="Enable cheap sampled manga/book classification.",
        )
        classify_group.add_argument(
            "--no-manga-classify",
            dest="manga_classify",
            action="store_false",
            help="Disable automatic manga/book classification.",
        )
    parser.add_argument(
        "--achievement-profile",
        dest="achievement_profile_file",
        default="",
        help=(
            "Achievement baseline/levels JSON path. "
            "Default: achievement_profile.json beside the CSV."
        ),
    )
    parser.add_argument(
        "--completed-threshold",
        dest="completed_book_threshold_percent",
        type=float,
        default=None,
        help=(
            "Kindle progress percent required to count a book as completed. "
            "Default: settings.json, then 95."
        ),
    )
    parser.add_argument(
        "--fmcache-db",
        default="",
        help=(
            "Path to fmcache.db. Default: system\\fmcache\\fmcache.db "
            "under the detected Kindle root."
        ),
    )
    parser.add_argument(
        "--covers-dir",
        default="",
        help=(
            "Cover thumbnail cache directory. "
            "Default: _covers beside the output HTML."
        ),
    )
    parser.add_argument(
        "--ebook-meta",
        default="",
        help="Path to Calibre ebook-meta.exe. Auto-detected when omitted.",
    )
    parser.add_argument(
        "--skip-covers",
        action="store_true",
        help="Render text placeholders without reading ebook covers.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print cover extraction warnings.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version="%(prog)s " + VERSION,
    )
    return parser


def resolve_ocr_cache_dir(
    raw_value: Any,
    base_dir: Path,
) -> Path:
    text = str(raw_value or "").strip()
    if not text or text == DEFAULT_OCR_CACHE_DIRNAME:
        candidate_roots = [Path(base_dir)]
        candidate_roots.extend(Path(base_dir).parents[:3])
        for candidate_root in candidate_roots:
            candidate = candidate_root / DEFAULT_OCR_CACHE_DIRNAME
            if (
                (candidate / "render").is_dir()
                or (candidate / "results").is_dir()
            ):
                return candidate.resolve()
        return (Path(base_dir) / DEFAULT_OCR_CACHE_DIRNAME).resolve()
    configured = Path(text).expanduser()
    if not configured.is_absolute():
        configured = Path(base_dir) / configured
    return configured.resolve()


def _print_ocr_progress(result: Dict[str, Any]) -> None:
    title = str(result.get("title") or "(unknown)")
    index = parse_int(result.get("manga_index")) or 0
    total_books = parse_int(result.get("manga_total")) or 0
    processed = parse_int(result.get("ocr_pages_processed")) or 0
    pages_total = parse_int(result.get("ocr_pages_total")) or 0
    status = str(result.get("ocr_status") or "unknown")
    engine = str(result.get("ocr_engine") or "unknown")
    language = str(result.get("ocr_language") or "unknown")
    cache_text = (
        "cache=book-hit"
        if result.get("_ocr_cache_hit")
        else (
            "cache=partial "
            f"(page-hit={parse_int(result.get('_ocr_page_cache_hits')) or 0}, "
            f"page-miss={parse_int(result.get('_ocr_page_cache_misses')) or 0})"
        )
    )
    print(
        f"OCR {index}/{total_books}: {title} | "
        f"pages={processed}/{pages_total} | language={language} | "
        f"engine={engine} | status={status} | {cache_text}"
    )


def apply_manga_ocr_to_csv(
    csv_path: Path,
    cache_dir: Path,
    args: argparse.Namespace,
) -> int:
    """Run optional manga OCR without changing exact word-count semantics."""

    if runtime_variant() == "text":
        return 0
    if kindle_manga_ocr is None:
        print(
            "warning: kindle_manga_ocr.py is not available; OCR status was "
            "not written, but the report can still be generated."
        )
        return 0
    try:
        dependency_lines = kindle_manga_ocr.dependency_report(
            args.ocr_engine,
            args.ocr_language,
        )
    except Exception as exc:
        dependency_lines = [
            f"dependency check failed: {type(exc).__name__}: {exc}"
        ]
    if dependency_lines:
        print(
            "warning: OCR dependencies or offline models are missing; "
            "nothing was installed and no network access was attempted."
        )
        for line in dependency_lines:
            print("  - " + line)
        print(
            "  OCR rows will be marked with an explicit dependency/status "
            "error for this run."
        )
    config = kindle_manga_ocr.OcrConfig(
        language=args.ocr_language,
        engine=args.ocr_engine,
        dpi=args.ocr_dpi,
        min_confidence=args.ocr_min_confidence,
        cache_dir=cache_dir,
        mode=getattr(args, "ocr_mode", "full"),
        workers=getattr(args, "ocr_workers", 0),
        sample_pages=getattr(args, "ocr_sample_pages", 6),
    )
    try:
        counts = kindle_manga_ocr.enrich_csv_with_manga_ocr(
            csv_path,
            config,
            logger=None,
            progress_callback=_print_ocr_progress,
        )
    except Exception as exc:
        print(
            "warning: manga OCR post-processing failed without stopping the "
            f"scan: {type(exc).__name__}: {exc}"
        )
        return 0
    print(
        "Manga OCR: "
        f"{counts['ocr_ok']}/{counts['manga_rows']} manga row(s) complete; "
        f"{counts['ocr_non_ok']} row(s) need attention; "
        f"{counts.get('ocr_cache_hits', 0)} whole-book cache hit(s); "
        f"cache: {cache_dir}"
    )
    return 0


def _write_book_modes_preserving_structure(
    path: Path,
    decisions: Sequence[Tuple[str, str, str]],
) -> None:
    """Merge classifier decisions into book_modes.json without dropping keys."""

    raw: Dict[str, Any] = {}
    if path.is_file():
        try:
            decoded = json.loads(path.read_text(encoding="utf-8-sig"))
            if isinstance(decoded, dict):
                raw = decoded
        except (OSError, ValueError):
            raw = {}
    for scope, key, mode in decisions:
        if not key or mode not in {"book", "manga"}:
            continue
        bucket = raw.get(scope)
        if not isinstance(bucket, dict):
            bucket = {}
            raw[scope] = bucket
        bucket[str(key)] = mode
    raw.setdefault("version", 1)
    _atomic_write_text(
        path,
        json.dumps(raw, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
    )


def apply_manga_classification_to_csv(
    csv_path: Path,
    cache_dir: Path,
    args: argparse.Namespace,
    book_modes_path: Path,
) -> Dict[str, Any]:
    """Classify uncertain rows and persist confident decisions."""

    if kindle_manga_ocr is None:
        return {"processed": 0, "manga": 0, "book": 0, "uncertain": 0}
    csv_path = Path(csv_path)
    with csv_path.open(
        "r",
        encoding="utf-8-sig",
        newline="",
    ) as handle:
        reader = csv.DictReader(handle)
        fieldnames = list(reader.fieldnames or [])
        rows = [dict(row) for row in reader]
    if not fieldnames:
        return {"processed": 0, "manga": 0, "book": 0, "uncertain": 0}

    existing_modes, _ = load_book_modes(str(book_modes_path))
    counts = {
        "processed": 0,
        "manga": 0,
        "book": 0,
        "uncertain": 0,
        "mapped": 0,
    }
    decisions: List[Tuple[str, str, str]] = []
    config = kindle_manga_ocr.OcrConfig(
        language=getattr(args, "ocr_language", "auto"),
        engine=getattr(args, "ocr_engine", "auto"),
        dpi=min(
            int(getattr(args, "ocr_dpi", kindle_manga_ocr.OCR_SAMPLE_DPI)),
            kindle_manga_ocr.OCR_SAMPLE_DPI,
        ),
        min_confidence=getattr(args, "ocr_min_confidence", 0.7),
        cache_dir=cache_dir,
        mode="classify",
        workers=int(getattr(args, "ocr_workers", 0) or 0),
        sample_pages=int(getattr(args, "ocr_sample_pages", 6) or 6),
    )
    output_fields = list(fieldnames)
    for field in getattr(kindle_reader, "CSV_FIELDS", ()):
        if field not in output_fields:
            output_fields.append(field)
    for field in kindle_manga_ocr.OCR_FIELDS:
        if field not in output_fields:
            output_fields.append(field)

    for row in rows:
        if str(row.get("reading_mode") or "").casefold() == "manga":
            row["word_count_source"] = "not-applicable"
            continue
        if str(row.get("reading_mode_source") or "") == "book-modes":
            counts["mapped"] += 1
            continue
        if mapped_reading_mode_for_book(row, existing_modes) is not None:
            counts["mapped"] += 1
            continue
        counts["processed"] += 1
        try:
            result = kindle_manga_ocr.classify_manga_candidate(
                row,
                config,
                logger=None,
            )
        except Exception as exc:
            result = {
                "reading_mode": None,
                "classification_source": "classifier-failed",
                "classification_confidence": 0.0,
                "classification_reason": f"{type(exc).__name__}: {exc}",
            }
        mode = str(result.get("reading_mode") or "").casefold()
        row["classification_source"] = str(
            result.get("classification_source") or ""
        )
        row["classification_confidence"] = result.get(
            "classification_confidence"
        )
        row["classification_reason"] = str(
            result.get("classification_reason") or ""
        )
        if mode == "manga":
            row["reading_mode"] = "manga"
            row["reading_mode_source"] = str(
                result.get("classification_source") or "classifier"
            )
            row["word_count_source"] = "not-applicable"
            row["series_key"] = series_key_for_book(row)
            counts["manga"] += 1
        elif mode == "book":
            row["reading_mode"] = "book"
            row["reading_mode_source"] = str(
                result.get("classification_source") or "classifier"
            )
            counts["book"] += 1
        else:
            row["reading_mode"] = "book"
            row["reading_mode_source"] = "uncertain"
            counts["uncertain"] += 1
            continue
        title = display_title(str(row.get("title") or "")).strip()
        asin = _normalize_mode_asin(row.get("asin"))
        if asin:
            decisions.append(("by_asin", asin, mode))
        if title:
            decisions.append(("by_title", title, mode))
        series_key = series_key_for_book(row)
        if series_key:
            decisions.append(("by_series", series_key, mode))

    if decisions:
        _write_book_modes_preserving_structure(book_modes_path, decisions)
    write_csv_rows(csv_path, rows, output_fields)
    return counts


def apply_manga_page_counts_to_csv(
    csv_path: Path,
    cache_dir: Path,
) -> Dict[str, int]:
    """Fill source-file page counts and page-read estimates for manga rows."""

    if kindle_manga_ocr is None:
        return {"manga_rows": 0, "page_counts": 0, "unknown": 0}
    csv_path = Path(csv_path)
    with csv_path.open(
        "r",
        encoding="utf-8-sig",
        newline="",
    ) as handle:
        reader = csv.DictReader(handle)
        fieldnames = list(reader.fieldnames or [])
        rows = [dict(row) for row in reader]
    if not fieldnames:
        return {"manga_rows": 0, "page_counts": 0, "unknown": 0}
    output_fields = list(fieldnames)
    for field in list(kindle_manga_ocr.OCR_FIELDS) + list(
        getattr(kindle_reader, "CSV_FIELDS", ())
    ):
        if field not in output_fields:
            output_fields.append(field)
    counts = {"manga_rows": 0, "page_counts": 0, "unknown": 0}
    for row in rows:
        if str(row.get("reading_mode") or "").casefold() != "manga":
            continue
        counts["manga_rows"] += 1
        source = kindle_manga_ocr.resolve_ocr_source(row, None)
        if source is None:
            counts["unknown"] += 1
            row["manga_pages_source"] = row.get("manga_pages_source") or "unknown"
            continue
        resolved_pages, source_label, error = kindle_manga_ocr.resolve_manga_page_count(
            source,
            cache_dir,
        )
        pages = resolved_pages or parse_int(row.get("manga_pages_total"))
        if resolved_pages is None:
            counts["unknown"] += 1
            row["manga_pages_source"] = "unknown"
            if error:
                row["classification_reason"] = (
                    str(row.get("classification_reason") or "")
                    + (("; " if row.get("classification_reason") else "") + error)
                )
        if resolved_pages is not None:
            counts["page_counts"] += 1
            row["manga_pages_total"] = str(resolved_pages)
            progress = manga_progress_percent(row)
            if progress is not None:
                row["manga_pages_read"] = str(
                    max(
                        0,
                        min(
                            resolved_pages,
                            int(round(resolved_pages * progress / 100.0)),
                        ),
                    )
                )
            row["manga_pages_source"] = source_label
        position_info = kindle_manga_ocr.resolve_kfx_position_info(
            source,
            cache_dir,
        )
        max_position = parse_int(
            (position_info or {}).get("last_pid")
        )
        fpr_position: Optional[int] = None
        source_file = str(row.get("source_file") or "")
        if source_file and kindle_reader is not None:
            sidecar = kindle_reader.decode_sidecar(
                source_file,
                kindle_reader._make_logger(False),
            )
            fpr_position = kindle_manga_ocr.position_number(
                (sidecar.get("fpr") or {}).get("position")
            )
            if fpr_position is None:
                fpr_position = kindle_manga_ocr.position_number(
                    (sidecar.get("lpr") or {}).get("position")
                )
        if max_position and fpr_position is not None:
            derived_progress = round(
                max(
                    0.0,
                    min(100.0, fpr_position / float(max_position) * 100.0),
                ),
                1,
            )
            current_progress = manga_progress_percent(row)
            if (
                current_progress is None
                or derived_progress > current_progress
            ):
                row["progress_percent"] = f"{derived_progress:g}"
                row["progress"] = str(int(round(derived_progress)))
                if pages:
                    row["manga_pages_read"] = str(
                        max(
                            0,
                            min(
                                int(pages),
                                int(round(pages * derived_progress / 100.0)),
                            ),
                        )
                    )
                row["manga_progress_source"] = "kfx-position-map"
            row["manga_position"] = str(fpr_position)
            row["manga_max_position"] = str(max_position)
    write_csv_rows(csv_path, rows, output_fields)
    return counts


def render_report(
    csv_path: Path,
    html_path: Path,
    source_path: str,
    history_path: Path,
    covers_dir: Path,
    ebook_meta_path: Optional[str],
    skip_covers: bool,
    verbose: bool,
    history_limit: int = DEFAULT_HISTORY_LIMIT,
    fmcache_db: str = "",
    xlsx_path: Optional[Path] = None,
    book_modes_path: Optional[Path] = None,
    achievement_profile_path: Optional[Path] = None,
    completed_threshold_percent: float = (
        DEFAULT_COMPLETED_BOOK_THRESHOLD_PERCENT
    ),
    write_back_csv: bool = False,
) -> int:
    if not csv_path.is_file():
        print(f"error: CSV file does not exist: {csv_path}")
        return 2

    free_ok, free_bytes = check_free_space(html_path.parent)
    if not free_ok:
        available = (
            format_bytes(free_bytes)
            if free_bytes is not None
            else "unknown"
        )
        print(
            "error: not enough free disk space for the report: "
            f"{available} available"
        )
        return 2

    try:
        history = load_history(history_path)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"error: cannot read history file {history_path}: {exc}")
        return 2

    resolved_modes_path = (
        Path(book_modes_path)
        if book_modes_path is not None
        else csv_path.parent / DEFAULT_BOOK_MODES_FILENAME
    )
    book_modes, book_mode_warning = load_book_modes(
        str(resolved_modes_path)
    )
    if book_mode_warning and verbose:
        print("book mode warning: %s" % book_mode_warning)
    resolved_profile_path = (
        Path(achievement_profile_path)
        if achievement_profile_path is not None
        else csv_path.parent / DEFAULT_ACHIEVEMENT_PROFILE_FILENAME
    )
    existing_profile = load_achievement_profile(resolved_profile_path)

    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        fieldnames = list(reader.fieldnames or [])
        rows = [dict(row) for row in reader]
    resolved_fmcache = resolve_fmcache_db(
        source_path or csv_path.parent,
        fmcache_db or None,
    )
    fmcache = read_fmcache_sessions(resolved_fmcache)
    rows, fallback_applied = merge_fmcache_book_fallbacks(
        rows,
        fmcache,
        book_modes=book_modes,
    )
    rows, history_applied = merge_history_book_fallbacks(
        rows,
        history,
        book_modes=book_modes,
    )
    fallback_applied.extend(history_applied)
    if write_back_csv and fallback_applied:
        try:
            write_csv_rows(csv_path, rows, fieldnames)
        except OSError as exc:
            print(f"error: cannot update CSV file {csv_path}: {exc}")
            return 2
    resolved_xlsx_path = (
        Path(xlsx_path).expanduser().resolve()
        if xlsx_path is not None
        else html_path.with_suffix(".xlsx")
    )
    cover_urls: Dict[str, str] = {}
    cover_errors: List[str] = []
    if not skip_covers:
        cover_assets, cover_errors = prepare_covers(
            rows,
            covers_dir,
            ebook_meta_path=ebook_meta_path,
        )
        cover_urls = cover_urls_for_html(cover_assets, html_path)
    for book_id, cover_url in existing_cover_urls_for_html(html_path).items():
        cover_urls.setdefault(book_id, cover_url)

    generated_at = datetime.now().astimezone().isoformat(timespec="seconds")
    snapshot = build_snapshot(
        rows,
        generated_at,
        source_path or str(csv_path.parent),
        book_modes=book_modes,
    )
    trends = build_trends(snapshot, history)
    achievement_profile = build_achievement_profile(
        rows,
        history=history,
        captured_at=generated_at,
        source_path=source_path or str(csv_path.parent),
        existing_profile=existing_profile,
        threshold_percent=completed_threshold_percent,
        trends=trends,
    )
    heatmap = build_heatmap_data(history, snapshot, fmcache)
    period_reports = build_period_reports(snapshot, history, fmcache)
    document = build_html(
        rows,
        csv_path,
        source_path or str(csv_path.parent),
        trends=trends,
        heatmap=heatmap,
        period_reports=period_reports,
        generated_at=generated_at,
        history_path=history_path,
        cover_urls=cover_urls,
        covers_dir=covers_dir,
        xlsx_path=resolved_xlsx_path,
        achievement_profile=achievement_profile,
        completed_threshold_percent=completed_threshold_percent,
        book_modes=book_modes,
    )
    try:
        write_xlsx_report(
            rows,
            resolved_xlsx_path,
            fieldnames=fieldnames,
            source_path=source_path or str(csv_path.parent),
            generated_at=generated_at,
            trends=trends,
        )
    except (OSError, RuntimeError) as exc:
        print(f"error: cannot write XLSX report {resolved_xlsx_path}: {exc}")
        return 2
    try:
        _atomic_write_text(html_path, document)
    except OSError as exc:
        print(f"error: cannot write HTML report {html_path}: {exc}")
        return 2
    try:
        save_achievement_profile(
            resolved_profile_path,
            achievement_profile,
        )
    except OSError as exc:
        print(
            f"error: cannot write achievement profile "
            f"{resolved_profile_path}: {exc}"
        )
        return 2
    try:
        appended = save_history(
            history_path,
            history,
            snapshot,
            history_limit=history_limit,
        )
    except OSError as exc:
        print(f"error: cannot save history file {history_path}: {exc}")
        return 2
    print(f"Visual report: {html_path}")
    print(f"XLSX report: {resolved_xlsx_path}")
    print(f"Achievement profile: {resolved_profile_path}")
    if appended:
        print(f"History snapshot: {history_path}")
    else:
        print(f"History snapshot unchanged: {history_path}")
    print(f"Cover thumbnails: {len(cover_urls)}/{len(rows)} in {covers_dir}")
    print(
        "Reading heatmap: "
        f"{heatmap['source_label']} ({heatmap['status']})"
    )
    if cover_errors:
        if verbose:
            for error in cover_errors:
                print(f"cover warning: {error}")
        print(f"Cover warnings: {len(cover_errors)}")
    return 0


def run_full_workflow(args: argparse.Namespace) -> int:
    root = (
        Path(args.output_dir).expanduser().resolve()
        if args.output_dir
        else application_root()
    )
    csv_path = (
        Path(args.output_csv).expanduser().resolve()
        if args.output_csv
        else root / DEFAULT_CSV_FILENAME
    )
    html_path = (
        Path(args.output_html).expanduser().resolve()
        if args.output_html
        else root / DEFAULT_HTML_FILENAME
    )
    history_path = root / HISTORY_FILENAME
    covers_dir = (
        Path(args.covers_dir).expanduser().resolve()
        if args.covers_dir
        else root / COVER_DIRECTORY_NAME
    )
    cache_path = root / DEFAULT_CACHE_FILENAME
    wordcount_dir = root / DEFAULT_WORDCOUNT_DIRECTORY
    ocr_cache_dir = resolve_ocr_cache_dir(
        getattr(args, "ocr_cache_dir", ""),
        root,
    )
    settings_path_for_outputs = Path(
        getattr(args, "_settings_path", root / SETTINGS_FILENAME)
    )

    def full_side_path(
        raw_value: Any,
        default_name: str,
        explicit: bool,
    ) -> Path:
        if explicit and str(raw_value or "").strip():
            return Path(str(raw_value)).expanduser().resolve()
        text = str(raw_value or "").strip()
        if not text or text == default_name:
            return root / default_name
        configured = Path(text).expanduser()
        if not configured.is_absolute():
            configured = settings_path_for_outputs.parent / configured
        return configured.resolve()

    book_modes_path = full_side_path(
        args.book_modes_file,
        DEFAULT_BOOK_MODES_FILENAME,
        getattr(args, "_book_modes_was_explicit", False),
    )
    if (
        not book_modes_path.is_file()
        and not getattr(args, "_book_modes_was_explicit", False)
    ):
        fallback_modes = application_root() / DEFAULT_BOOK_MODES_FILENAME
        if fallback_modes.is_file():
            try:
                book_modes_path.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(fallback_modes, book_modes_path)
            except OSError:
                # Reading the fallback is still safe; classification then
                # updates a copy only when the destination is writable.
                book_modes_path = fallback_modes
    achievement_profile_path = full_side_path(
        args.achievement_profile_file,
        DEFAULT_ACHIEVEMENT_PROFILE_FILENAME,
        getattr(args, "_achievement_profile_was_explicit", False),
    )

    print(f"CSV:    {csv_path}")
    print(f"Report: {html_path}")

    if kindle_reader is None:
        print("error: kindle_reader.py is not available")
        return 2

    free_space_targets = [
        (csv_path.parent, "CSV output"),
        (html_path.parent, "HTML output"),
        (wordcount_dir, "word-count cache"),
    ]
    if args.manga_ocr or args.manga_classify:
        free_space_targets.append((ocr_cache_dir, "OCR cache"))
    for target, label in free_space_targets:
        free_ok, free_bytes = check_free_space(target)
        if not free_ok:
            available = (
                format_bytes(free_bytes)
                if free_bytes is not None
                else "unknown"
            )
            print(
                f"error: not enough free disk space for {label}: "
                f"{available} available"
            )
            return 2

    try:
        documents = kindle_reader.resolve_documents(args.path or None)
    except OSError as exc:
        print(f"error: cannot scan for Kindle: {exc}")
        return 2
    if not documents:
        print(
            "error: Kindle documents folder was not found. "
            "Connect the Kindle by USB or pass --path.",
        )
        return 2

    root.mkdir(parents=True, exist_ok=True)
    print(f"Kindle: {documents}")
    print("[1/2] Reading Kindle sidecars...")
    reader_args = [
        "--path",
        documents,
        "--output",
        str(csv_path),
        "--word-count-cache",
        str(cache_path),
        "--wordcount-dir",
        str(wordcount_dir),
        "--verbose",
        "--word-count-language",
        str(args.word_count_language or "auto"),
        "--book-modes",
        str(book_modes_path),
    ]
    if args.exact_word_count:
        reader_args.append("--exact-word-count")

    result = kindle_reader.main(reader_args)
    if result != 0:
        return result
    if not csv_path.is_file():
        print(f"error: CSV file was not created: {csv_path}")
        return 2

    if args.manga_classify:
        try:
            classify_counts = apply_manga_classification_to_csv(
                csv_path,
                ocr_cache_dir,
                args,
                book_modes_path,
            )
            print(
                "Manga classifier: "
                f"{classify_counts.get('manga', 0)} manga, "
                f"{classify_counts.get('book', 0)} book, "
                f"{classify_counts.get('uncertain', 0)} uncertain; "
                f"mapping: {book_modes_path}"
            )
        except Exception as exc:
            print(
                "warning: manga classification failed without stopping the "
                f"scan: {type(exc).__name__}: {exc}"
            )
    try:
        page_counts = apply_manga_page_counts_to_csv(
            csv_path,
            ocr_cache_dir,
        )
        if page_counts.get("manga_rows"):
            print(
                "Manga pages: "
                f"{page_counts.get('page_counts', 0)}/"
                f"{page_counts.get('manga_rows', 0)} known; "
                f"{page_counts.get('unknown', 0)} unknown"
            )
    except Exception as exc:
        print(
            "warning: manga page-count enrichment failed without stopping "
            f"the report: {type(exc).__name__}: {exc}"
        )

    print("[2/2] Building visual report...")
    result = render_report(
        csv_path=csv_path,
        html_path=html_path,
        source_path=documents,
        history_path=history_path,
        covers_dir=covers_dir,
        ebook_meta_path=args.ebook_meta or None,
        skip_covers=args.skip_covers,
        verbose=args.verbose,
        history_limit=args.history_limit,
        fmcache_db=args.fmcache_db,
        book_modes_path=book_modes_path,
        achievement_profile_path=achievement_profile_path,
        completed_threshold_percent=getattr(
            args,
            "completed_book_threshold_percent",
            DEFAULT_COMPLETED_BOOK_THRESHOLD_PERCENT,
        ),
        write_back_csv=True,
    )
    if result != 0:
        return result

    if args.manga_ocr and str(getattr(args, "ocr_mode", "full")) != "classify":
        print("[post] Running optional manga word OCR...")
        result = apply_manga_ocr_to_csv(
            csv_path,
            ocr_cache_dir,
            args,
        )
        if result != 0:
            return result
        try:
            apply_manga_page_counts_to_csv(csv_path, ocr_cache_dir)
        except Exception:
            pass
        print("[post] Rebuilding report with OCR results...")
        result = render_report(
            csv_path=csv_path,
            html_path=html_path,
            source_path=documents,
            history_path=history_path,
            covers_dir=covers_dir,
            ebook_meta_path=args.ebook_meta or None,
            skip_covers=args.skip_covers,
            verbose=args.verbose,
            history_limit=args.history_limit,
            fmcache_db=args.fmcache_db,
            book_modes_path=book_modes_path,
            achievement_profile_path=achievement_profile_path,
            completed_threshold_percent=getattr(
                args,
                "completed_book_threshold_percent",
                DEFAULT_COMPLETED_BOOK_THRESHOLD_PERCENT,
            ),
            write_back_csv=True,
        )
        if result != 0:
            return result

    if args.open_html:
        open_report(html_path)
    return 0


def _windows_version_resource(executable_name: str) -> str:
    """Return a PyInstaller version resource matching VERSION."""

    parts = [int(part) for part in re.findall(r"\d+", VERSION)[:4]]
    parts.extend([0] * (4 - len(parts)))
    numeric_version = tuple(parts)
    dotted_version = ".".join(str(part) for part in numeric_version)
    description = (
        "Kindle Reading Dashboard - Text + Manga"
        if "OCR" in executable_name
        else "Kindle Reading Dashboard - Text"
    )
    return f"""VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={numeric_version!r},
    prodvers={numeric_version!r},
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        '040904B0',
        [
          StringStruct('CompanyName', 'Kindle Reading Dashboard contributors'),
          StringStruct('FileDescription', '{description}'),
          StringStruct('FileVersion', '{dotted_version}'),
          StringStruct('InternalName', '{Path(executable_name).stem}'),
          StringStruct('LegalCopyright', 'GPL v3; no warranty'),
          StringStruct('OriginalFilename', '{executable_name}'),
          StringStruct('ProductName', 'Kindle Reading Dashboard'),
          StringStruct('ProductVersion', '{dotted_version}')
        ]
      )
    ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"""


def build_executable(dist_dir: Optional[str] = None) -> int:
    root = Path(__file__).resolve().parent
    local_dependencies = root / ".tmp-pip"
    if local_dependencies.is_dir():
        sys.path.insert(0, str(local_dependencies))

    try:
        import PyInstaller.__main__ as pyinstaller_main
    except ImportError:
        print("error: PyInstaller is not installed")
        print(
            "Install it with: "
            f'"{sys.executable}" -m pip install --target '
            f'"{local_dependencies}" pyinstaller'
        )
        return 2

    destination = (
        Path(dist_dir).expanduser().resolve()
        if dist_dir
        else root / "dist"
    )
    build_root = root / "build" / "pyinstaller"
    work_path = build_root / "work"
    spec_path = build_root / "spec"
    vendor_path = root / "vendor"

    # OCR engines live in the external .ocr-pip directory and may import
    # standard-library modules dynamically. Keep those modules available in
    # the frozen executable without bundling the multi-gigabyte OCR stack.
    import importlib.util

    vendor_arguments = _vendor_path_arguments(vendor_path)
    if not vendor_arguments:
        print(
            "note: vendor tokenizers are not readable; collecting "
            "jieba/janome from installed packages"
        )

    def build_variant(variant: str, executable_name: str) -> Optional[Path]:
        version_directory = spec_path / variant
        version_directory.mkdir(parents=True, exist_ok=True)
        version_file = version_directory / f"{Path(executable_name).stem}.version.txt"
        _atomic_write_text(version_file, _windows_version_resource(executable_name))
        arguments = [
            str(root / "kindle_report.py"),
            "--name",
            Path(executable_name).stem,
            "--onefile",
            "--console",
            "--clean",
            "--noconfirm",
            "--distpath",
            str(destination),
            "--workpath",
            str(work_path / variant),
            "--specpath",
            str(spec_path / variant),
            "--version-file",
            str(version_file),
            "--paths",
            str(root),
            "--add-data",
            (
                str(root / "third_party" / "calibre-kfx-input")
                + os.pathsep
                + "kfx-input"
            ),
            "--hidden-import",
            "kindle_reader",
            "--collect-all",
            "jieba",
            "--collect-all",
            "janome",
            "--collect-all",
            "openpyxl",
            "--collect-submodules",
            "PIL",
            "--exclude-module",
            "numpy",
            "--exclude-module",
            "pandas",
        ]
        if variant == "ocr":
            arguments.extend(["--hidden-import", "kindle_manga_ocr"])
        else:
            arguments.extend(["--exclude-module", "kindle_manga_ocr"])
        for module_name in sorted(sys.stdlib_module_names):
            if module_name.startswith("_") or not module_name.isidentifier():
                continue
            try:
                if importlib.util.find_spec(module_name) is None:
                    continue
            except (ImportError, AttributeError, ValueError):
                continue
            arguments.extend(["--hidden-import", module_name])
        arguments.extend(["--collect-submodules", "unittest"])
        arguments.extend(["--collect-submodules", "setuptools"])
        arguments.extend(["--hidden-import", "pkg_resources"])
        # Only add the vendor directory to PyInstaller's module path when its
        # tokenizer copies are readable. Some checkouts ship vendored jieba/janome
        # trees with restrictive ACLs; pointing PyInstaller at them then fails with
        # PermissionError and floods the log with "not found" module warnings. In
        # that case fall back to installed packages via --collect-all.
        arguments.extend(vendor_arguments)
        print(f"Building {executable_name}...")
        pyinstaller_main.run(arguments)
        executable = destination / executable_name
        if not executable.is_file():
            print(f"error: executable was not created: {executable}")
            return None
        return executable

    built: List[Path] = []
    for variant, executable_name in (
        ("text", "KindleReader-Text.exe"),
        ("ocr", "KindleReader-OCR.exe"),
    ):
        executable = build_variant(variant, executable_name)
        if executable is None:
            return 2
        built.append(executable)

    default_modes = destination / DEFAULT_BOOK_MODES_FILENAME
    source_modes = root / DEFAULT_BOOK_MODES_FILENAME
    if source_modes.is_file() and source_modes.resolve() != default_modes.resolve():
        try:
            shutil.copy2(source_modes, default_modes)
        except OSError as exc:
            print(
                "warning: could not copy book mode mapping "
                f"({source_modes} -> {default_modes}): {exc}"
            )
    elif not default_modes.exists():
        try:
            _atomic_write_text(
                default_modes,
                json.dumps(
                    {"version": 1, "by_asin": {}, "by_title": {}},
                    ensure_ascii=False,
                    indent=2,
                    sort_keys=True,
                )
                + "\n",
            )
        except OSError as exc:
            print(
                "warning: could not create default book mode mapping "
                f"({default_modes}): {exc}"
            )
    for executable in built:
        print(f"Executable: {executable}")
    return 0


def _vendor_tokenizers_readable(vendor_path: Path) -> bool:
    """Return True only when vendored jieba and janome trees can be listed."""

    for name in ("jieba", "janome"):
        try:
            os.listdir(vendor_path / name)
        except OSError:
            return False
    return True


def _vendor_path_arguments(vendor_path: Path) -> List[str]:
    """Return PyInstaller --paths args for vendor, or [] when unreadable."""

    if _vendor_tokenizers_readable(vendor_path):
        return ["--paths", str(vendor_path)]
    return []


def main(argv: Optional[Sequence[str]] = None) -> int:
    raw_args = list(argv) if argv is not None else sys.argv[1:]
    app_variant = runtime_variant()
    args = build_arg_parser(app_variant).parse_args(raw_args)

    if args.build_exe:
        result = build_executable(args.dist_dir or None)
        pause_after_error(raw_args, args.no_wait)
        return result

    if bool(args.csv) != bool(args.html):
        print("error: --csv and --html must be supplied together")
        pause_after_error(raw_args, args.no_wait)
        return 2
    if args.history_limit is not None and args.history_limit < 0:
        print("error: --history-limit must be 0 or greater")
        pause_after_error(raw_args, args.no_wait)
        return 2

    settings_path = (
        Path(args.settings).expanduser().resolve()
        if args.settings
        else application_root() / SETTINGS_FILENAME
    )
    report_only = bool(args.csv and args.html)
    settings, settings_warnings = load_settings(
        settings_path,
        create_default=not report_only,
    )
    explicit_book_modes = bool(args.book_modes_file)
    explicit_achievement_profile = bool(args.achievement_profile_file)
    args._book_modes_was_explicit = explicit_book_modes
    args._achievement_profile_was_explicit = explicit_achievement_profile
    args._settings_path = settings_path
    effective_settings = merge_settings(args, settings)
    if args.verbose:
        print(f"Settings: {settings_path}")
        for warning in settings_warnings:
            print(f"settings warning: {warning}")
    args.path = effective_settings["kindle_path"]
    args.output_dir = effective_settings["output_dir"]
    args.word_count_language = effective_settings["language"]
    args.exact_word_count = effective_settings["exact_word_count"]
    args.open_html = effective_settings["open_html"]
    args.history_limit = effective_settings["history_limit"]
    args.book_modes_file = effective_settings["book_modes_file"]
    args.achievement_profile_file = effective_settings[
        "achievement_profile_file"
    ]
    args.completed_book_threshold_percent = effective_settings[
        "completed_book_threshold_percent"
    ]
    args.manga_ocr = (
        False if app_variant == "text" else effective_settings["manga_ocr"]
    )
    args.manga_classify = (
        False if app_variant == "text" else effective_settings["manga_classify"]
    )
    args.manga_volume_achievement = effective_settings[
        "manga_volume_achievement"
    ]
    args.ocr_language = effective_settings["ocr_language"]
    args.ocr_engine = effective_settings["ocr_engine"]
    args.ocr_dpi = effective_settings["ocr_dpi"]
    args.ocr_mode = effective_settings["ocr_mode"]
    args.ocr_workers = effective_settings["ocr_workers"]
    args.ocr_sample_pages = effective_settings["ocr_sample_pages"]
    args.ocr_min_confidence = effective_settings["ocr_min_confidence"]
    args.ocr_cache_dir = effective_settings["ocr_cache_dir"]
    if not 200 <= int(args.ocr_dpi) <= 400:
        print("error: --ocr-dpi must be between 200 and 400")
        pause_after_error(raw_args, args.no_wait)
        return 2
    if int(args.ocr_workers) < 0:
        print("error: --ocr-workers must be 0 or greater")
        pause_after_error(raw_args, args.no_wait)
        return 2
    if not 5 <= int(args.ocr_sample_pages) <= 8:
        print("error: --ocr-sample-pages must be between 5 and 8")
        pause_after_error(raw_args, args.no_wait)
        return 2
    if not 0.0 <= float(args.ocr_min_confidence) <= 1.0:
        print("error: --ocr-min-confidence must be between 0 and 1")
        pause_after_error(raw_args, args.no_wait)
        return 2

    if args.csv and args.html:
        csv_path = Path(args.csv).expanduser().resolve()
        html_path = Path(args.html).expanduser().resolve()
        history_path = (
            Path(args.history).expanduser().resolve()
            if args.history
            else csv_path.parent / HISTORY_FILENAME
        )
        covers_dir = (
            Path(args.covers_dir).expanduser().resolve()
            if args.covers_dir
            else html_path.parent / COVER_DIRECTORY_NAME
        )
        if args.manga_ocr:
            ocr_cache_dir = resolve_ocr_cache_dir(
                args.ocr_cache_dir,
                csv_path.parent,
            )
            result = apply_manga_ocr_to_csv(
                csv_path,
                ocr_cache_dir,
                args,
            )
            if result != 0:
                pause_after_error(raw_args, args.no_wait)
                return result
        def report_side_path(
            raw_value: Any,
            default_name: str,
            default_dir: Path,
            explicit: bool,
        ) -> Path:
            if explicit and str(raw_value or "").strip():
                return Path(str(raw_value)).expanduser().resolve()
            text = str(raw_value or "").strip()
            if not text or text == default_name:
                return default_dir / default_name
            configured = Path(text).expanduser()
            if not configured.is_absolute():
                configured = settings_path.parent / configured
            return configured.resolve()

        result = render_report(
            csv_path=csv_path,
            html_path=html_path,
            source_path=args.source,
            history_path=history_path,
            covers_dir=covers_dir,
            ebook_meta_path=args.ebook_meta or None,
            skip_covers=args.skip_covers,
            verbose=args.verbose,
            history_limit=args.history_limit,
            fmcache_db=args.fmcache_db,
            book_modes_path=report_side_path(
                args.book_modes_file,
                DEFAULT_BOOK_MODES_FILENAME,
                csv_path.parent,
                explicit_book_modes,
            ),
            achievement_profile_path=report_side_path(
                args.achievement_profile_file,
                DEFAULT_ACHIEVEMENT_PROFILE_FILENAME,
                csv_path.parent,
                explicit_achievement_profile,
            ),
            completed_threshold_percent=args.completed_book_threshold_percent,
        )
    else:
        result = run_full_workflow(args)

    pause_after_error(raw_args, args.no_wait)
    return result


class KindleReportTests(unittest.TestCase):
    @staticmethod
    def _write_csv(path: Path, row: Dict[str, Any]) -> None:
        fields = [
            "title",
            "author",
            "asin",
            "total_time_ms",
            "total_words",
            "wpm",
            "progress_percent",
        ]
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=fields)
            writer.writeheader()
            writer.writerow(row)

    def test_stable_book_id_prefers_asin_then_normalized_title(self) -> None:
        self.assertEqual(
            stable_book_id(
                {"title": "Ignored", "asin": "b000000001"}
            ),
            "asin:B000000001",
        )
        self.assertEqual(
            stable_book_id({"title": "  Ｃａｆｅ\tBook  ", "asin": ""}),
            "title:cafe book",
        )

    def test_snapshot_history_is_preserved_and_delta_is_calculated(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            history_path = Path(temp_dir) / HISTORY_FILENAME
            first = build_snapshot(
                [
                    {
                        "title": "Book",
                        "author": "Author",
                        "asin": "B000000001",
                        "total_time_ms": "60000",
                        "total_words": "120",
                        "wpm": "120",
                        "progress_percent": "10",
                    }
                ],
                "2026-09-30T10:00:00+08:00",
                "Kindle",
            )
            second = build_snapshot(
                [
                    {
                        "title": "Book",
                        "author": "Author",
                        "asin": "B000000001",
                        "total_time_ms": "120000",
                        "total_words": "300",
                        "wpm": "150",
                        "progress_percent": "15",
                    }
                ],
                "2026-09-30T11:00:00+08:00",
                "Kindle",
            )

            append_snapshot(history_path, first)
            history = load_history(history_path)
            trends = build_trends(second, history)
            save_history(history_path, history, second)
            saved = load_history(history_path)

        self.assertEqual(len(saved["snapshots"]), 2)
        self.assertEqual(saved["snapshots"][0], first)
        self.assertTrue(trends["has_previous"])
        self.assertEqual(
            trends["totals"]["total_time_ms"]["delta"],
            60000.0,
        )
        book = trends["books"][0]
        self.assertEqual(book["total_time_ms"]["delta"], 60000.0)
        self.assertEqual(book["total_words"]["delta"], 180.0)
        self.assertEqual(book["wpm"]["delta"], 30.0)
        self.assertEqual(book["progress_percent"]["delta"], 5.0)

    def test_main_writes_trend_data_between_two_runs(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            csv_path = root / "reading.csv"
            html_path = root / "dashboard.html"
            history_path = root / HISTORY_FILENAME
            first = {
                "title": "Trend Book",
                "author": "Author",
                "asin": "B000000001",
                "total_time_ms": "60000",
                "total_words": "100",
                "wpm": "100",
                "progress_percent": "10",
            }
            second = {
                **first,
                "total_time_ms": "120000",
                "total_words": "250",
                "wpm": "125",
                "progress_percent": "15",
            }

            self._write_csv(csv_path, first)
            self.assertEqual(
                main(
                    [
                        "--csv",
                        str(csv_path),
                        "--html",
                        str(html_path),
                        "--history",
                        str(history_path),
                        "--skip-covers",
                    ]
                ),
                0,
            )
            self._write_csv(csv_path, second)
            self.assertEqual(
                main(
                    [
                        "--csv",
                        str(csv_path),
                        "--html",
                        str(html_path),
                        "--history",
                        str(history_path),
                        "--skip-covers",
                    ]
                ),
                0,
            )

            saved = load_history(history_path)
            document = html_path.read_text(encoding="utf-8")

        self.assertEqual(len(saved["snapshots"]), 2)
        self.assertIn('id="trend-data"', document)
        self.assertIn('id="daily-chart"', document)
        self.assertIn('data-days="30"', document)
        self.assertIn('"delta":60000', document)
        self.assertIn("本次增量", document)
        self.assertIn("近 7 天阅读", document)
        self.assertIn("连续阅读", document)

    def test_cover_cards_render_images_and_layout_placeholders(self) -> None:
        class DashboardDomParser(HTMLParser):
            def __init__(self) -> None:
                super().__init__()
                self.ids = set()
                self.cover_cards = 0
                self.cover_images: List[str] = []
                self.placeholders: List[bool] = []

            def handle_starttag(
                self,
                tag: str,
                attrs: List[Tuple[str, Optional[str]]],
            ) -> None:
                attributes = dict(attrs)
                classes = set(
                    (attributes.get("class") or "").split()
                )
                element_id = attributes.get("id")
                if element_id:
                    self.ids.add(element_id)
                if "book-cover-card" in classes:
                    self.cover_cards += 1
                if tag == "img" and "data-cover-image" in attributes:
                    self.cover_images.append(str(attributes["src"]))
                if "cover-placeholder" in classes:
                    self.placeholders.append("hidden" in attributes)

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            cover_path = root / "covers" / "cover.jpg"
            cover_path.parent.mkdir()
            cover_path.write_bytes(b"cover")
            rows = [
                {
                    "title": "Cover Book",
                    "author": "Author",
                    "asin": "B000000001",
                    "total_time_ms": "60000",
                    "total_words": "100",
                    "book_words_exact": "1000",
                    "word_count_source": "en-regex",
                    "word_count_language": "en",
                    "wpm": "100",
                    "progress_percent": "10",
                },
                {
                    "title": "Placeholder Book",
                    "author": "Author",
                    "asin": "B000000002",
                    "total_time_ms": "0",
                    "total_words": "0",
                    "word_count_source": "janome",
                    "word_count_language": "ja",
                    "wpm": "0",
                    "progress_percent": "0",
                },
            ]
            document = build_html(
                rows,
                root / "reading.csv",
                "Kindle",
                generated_at="2026-09-30T12:00:00+08:00",
                cover_urls={
                    stable_book_id(rows[0]): "covers/cover.jpg",
                },
            )

        parser = DashboardDomParser()
        parser.feed(document)
        self.assertEqual(parser.cover_cards, 2)
        self.assertEqual(parser.cover_images, ["covers/cover.jpg"])
        self.assertEqual(len(parser.placeholders), 2)
        self.assertTrue(any(parser.placeholders))
        self.assertIn("cover-grid", parser.ids)
        self.assertIn("language-filter", parser.ids)
        self.assertIn("count-filter", parser.ids)
        self.assertIn("state-filter", parser.ids)
        self.assertIn("dataset.overflowX", document)
        self.assertIn("dataset.coverImagesLoaded", document)

    @unittest.skipIf(Image is None, "Pillow is not available")
    def test_cover_thumbnail_is_resized_and_readable(self) -> None:
        assert Image is not None
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "source.png"
            target = root / "thumbnail.jpg"
            Image.new("RGB", (1200, 1800), "#2f7d61").save(source)

            _write_cover_thumbnail(source, target)

            self.assertTrue(_cover_is_loadable(target))
            with Image.open(target) as thumbnail:
                width, height = thumbnail.size

        self.assertLessEqual(width, COVER_MAX_SIZE[0])
        self.assertLessEqual(height, COVER_MAX_SIZE[1])
        self.assertAlmostEqual(width / height, 2 / 3, places=2)


if __name__ == "__main__":
    raise SystemExit(main())
