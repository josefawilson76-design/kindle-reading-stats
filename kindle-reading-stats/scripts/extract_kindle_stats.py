#!/usr/bin/env python3
"""Extract Kindle reading statistics from KRDS sidecar files.

The script is intentionally read-only and uses only the Python standard
library. It scans a mounted Kindle for KRDS sidecars, decodes the documented
v1 binary framing, and exports summary data as CSV, JSON, and an offline HTML
table.

Typical usage:

    python extract_kindle_stats.py --kindle-path /media/user/Kindle
    python extract_kindle_stats.py --kindle-path E:\\ --format csv
"""

from __future__ import annotations

import argparse
import csv
import html
import json
import math
import os
import struct
import sys
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, List, Optional, Sequence, Tuple, Union


VERSION = "1.0.0"
KRDS_SIGNATURE = b"\x00\x00\x00\x00\x00\x1A\xB1\x26"

KRDS_EXTENSIONS = {
    ".azw3f",
    ".azw3r",
    ".yjf",
    ".yjr",
    ".mbs",
    ".mbp1",
    ".pdt",
    ".pds",
    ".tas",
    ".tal",
}

EBOOK_EXTENSIONS = {
    ".azw3",
    ".azw",
    ".azw1",
    ".mobi",
    ".kfx",
    ".pdf",
    ".prc",
    ".epub",
}


class KrdsError(Exception):
    """Base error for KRDS decoding."""


class KrdsFormatError(KrdsError):
    """The file is not a supported KRDS container."""


class KrdsTruncatedError(KrdsError):
    """The KRDS stream ended before a value was complete."""


@dataclass
class KrdsObject:
    """A name and its ordered value list inside a KRDS stream."""

    name: str
    values: List[Any] = field(default_factory=list)


@dataclass
class KrdsDocument:
    """Decoded top-level KRDS objects plus diagnostic information."""

    root: Dict[str, Any]
    warnings: List[str] = field(default_factory=list)
    trailing_bytes: int = 0
    partial: bool = False


class KrdsDecoder:
    """Decode the public/reverse-engineered KRDS v1 binary framing."""

    BOOLEAN = 0
    INT = 1
    LONG = 2
    UTF = 3
    DOUBLE = 4
    SHORT = 5
    FLOAT = 6
    BYTE = 7
    CHAR = 9
    OBJECT_BEGIN = -2
    OBJECT_END = -1

    def __init__(self, data: bytes, max_depth: int = 64) -> None:
        self.data = data
        self.offset = 0
        self.max_depth = max_depth
        self.warnings: List[str] = []

    @property
    def remaining(self) -> int:
        return len(self.data) - self.offset

    def _ensure(self, size: int, label: str = "value") -> None:
        if size < 0 or self.remaining < size:
            raise KrdsTruncatedError(
                "Insufficient data for {} at offset {}: need {}, have {}".format(
                    label, self.offset, size, self.remaining
                )
            )

    def _read_exact(self, size: int, label: str = "value") -> bytes:
        self._ensure(size, label)
        result = self.data[self.offset : self.offset + size]
        self.offset += size
        return result

    def _unpack(self, fmt: str, label: str) -> Any:
        size = struct.calcsize(fmt)
        self._ensure(size, label)
        value = struct.unpack_from(fmt, self.data, self.offset)[0]
        self.offset += size
        return value

    def _read_signed_byte(self, advance: bool = True) -> int:
        self._ensure(1, "signed byte")
        value = struct.unpack_from(">b", self.data, self.offset)[0]
        if advance:
            self.offset += 1
        return value

    def _read_utf_payload(self) -> str:
        # UTF-8 values have no type tag when nested directly in an object name.
        empty_flag = self._read_exact(1, "UTF-8 empty flag")[0]
        if empty_flag == 1:
            return ""
        if empty_flag not in (0, 1):
            self.warnings.append(
                "Unexpected UTF-8 empty flag {} at offset {}".format(
                    empty_flag, self.offset - 1
                )
            )

        length = self._unpack(">H", "UTF-8 length")
        raw = self._read_exact(length, "UTF-8 data")
        try:
            return raw.decode("utf-8")
        except UnicodeDecodeError:
            self.warnings.append(
                "Invalid UTF-8 data at offset {}; decoded with replacement".format(
                    self.offset - length
                )
            )
            return raw.decode("utf-8", errors="replace")

    def decode_value(self, depth: int = 0) -> Any:
        if depth > self.max_depth:
            raise KrdsFormatError(
                "KRDS nesting exceeds {} levels at offset {}".format(
                    self.max_depth, self.offset
                )
            )

        tag_offset = self.offset
        tag = self._read_signed_byte()

        if tag == self.BOOLEAN:
            value = self._read_exact(1, "boolean")[0]
            if value not in (0, 1):
                self.warnings.append(
                    "Unexpected boolean value {} at offset {}".format(
                        value, tag_offset
                    )
                )
            return bool(value)

        if tag == self.INT:
            return self._unpack(">i", "int32")

        if tag == self.LONG:
            return self._unpack(">q", "int64")

        if tag == self.UTF:
            return self._read_utf_payload()

        if tag == self.DOUBLE:
            return self._unpack(">d", "double")

        if tag == self.SHORT:
            return self._unpack(">h", "int16")

        if tag == self.FLOAT:
            return self._unpack(">f", "float")

        if tag == self.BYTE:
            return self._read_signed_byte()

        if tag == self.CHAR:
            return chr(self._read_exact(1, "char")[0])

        if tag == self.OBJECT_BEGIN:
            name = self._read_utf_payload()
            values: List[Any] = []

            # Object boundaries are explicitly delimited, which lets the
            # decoder preserve completed values if a later child is damaged.
            while True:
                self._ensure(1, "object end marker")
                if self._read_signed_byte(advance=False) == self.OBJECT_END:
                    self.offset += 1
                    break
                values.append(self.decode_value(depth + 1))

            return KrdsObject(name=name, values=values)

        if tag == self.OBJECT_END:
            raise KrdsFormatError(
                "Unexpected object end marker at offset {}".format(tag_offset)
            )

        raise KrdsFormatError(
            "Unknown KRDS type tag {} at offset {}".format(tag, tag_offset)
        )

    def decode_document(self) -> KrdsDocument:
        if len(self.data) < len(KRDS_SIGNATURE):
            raise KrdsTruncatedError("File is shorter than the KRDS signature")

        signature = self._read_exact(len(KRDS_SIGNATURE), "KRDS signature")
        if signature != KRDS_SIGNATURE:
            raise KrdsFormatError("KRDS signature is incorrect")

        first_value = self.decode_value()
        if first_value != 1:
            raise KrdsFormatError(
                "Unexpected KRDS header value: {!r}".format(first_value)
            )

        object_count = self.decode_value()
        if not isinstance(object_count, int) or isinstance(object_count, bool):
            raise KrdsFormatError(
                "KRDS object count is not an integer: {!r}".format(object_count)
            )
        if object_count < 0 or object_count > self.remaining:
            raise KrdsFormatError(
                "Implausible KRDS object count {} at offset {}".format(
                    object_count, self.offset
                )
            )

        root: Dict[str, Any] = {}
        partial = False
        for _ in range(object_count):
            try:
                item = self.decode_value()
            except KrdsError as exc:
                partial = True
                self.warnings.append(
                    "Stopped after a damaged top-level record at offset {}: {}".format(
                        self.offset, exc
                    )
                )
                break

            if not isinstance(item, KrdsObject):
                partial = True
                self.warnings.append(
                    "Top-level value is not an object at offset {}; skipped".format(
                        self.offset
                    )
                )
                continue

            existing = root.get(item.name)
            if existing is None:
                root[item.name] = item
            elif isinstance(existing, list):
                existing.append(item)
            else:
                root[item.name] = [existing, item]

        trailing_bytes = self.remaining
        if trailing_bytes:
            self.warnings.append(
                "{} trailing byte(s) after the declared object count".format(
                    trailing_bytes
                )
            )

        return KrdsDocument(
            root=root,
            warnings=list(self.warnings),
            trailing_bytes=trailing_bytes,
            partial=partial,
        )


def iter_krds_objects(
    value: Any, path: Tuple[Union[str, int], ...] = ()
) -> Iterator[Tuple[Tuple[Union[str, int], ...], KrdsObject]]:
    """Yield every KrdsObject in a decoded tree with a path for diagnostics."""

    if isinstance(value, KrdsObject):
        current_path = path + (value.name,)
        yield current_path, value
        for index, child in enumerate(value.values):
            yield from iter_krds_objects(child, current_path + (index,))
        return

    if isinstance(value, dict):
        for key, child in value.items():
            yield from iter_krds_objects(child, path + (str(key),))
        return

    if isinstance(value, (list, tuple)):
        for index, child in enumerate(value):
            yield from iter_krds_objects(child, path + (index,))


def _as_int(value: Any) -> Optional[int]:
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    return None


def _as_float(value: Any) -> Optional[float]:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        result = float(value)
        return result if math.isfinite(result) else None
    return None


@dataclass
class ReadingStats:
    total_time_ms: Optional[int] = None
    total_words: Optional[int] = None
    timer_percent: Optional[float] = None
    book_words: Optional[int] = None
    book_word_coverage_percent: Optional[float] = None
    timer_path: Optional[str] = None
    book_info_path: Optional[str] = None


def _path_string(path: Sequence[Union[str, int]]) -> str:
    pieces: List[str] = []
    for item in path:
        if isinstance(item, int):
            pieces.append("[{}]".format(item))
        else:
            pieces.append(str(item))
    return ".".join(pieces)


def _parse_timer_model(
    obj: KrdsObject, path: Sequence[Union[str, int]]
) -> Optional[ReadingStats]:
    # timer.model uses positional fields: version, time, words, percent,
    # followed by a nested average-calculator object.
    if len(obj.values) < 4:
        return None

    total_time = _as_int(obj.values[1])
    total_words = _as_int(obj.values[2])
    total_percent = _as_float(obj.values[3])
    if (
        total_time is None
        or total_words is None
        or total_time < 0
        or total_words < 0
    ):
        return None

    return ReadingStats(
        total_time_ms=total_time,
        total_words=total_words,
        timer_percent=total_percent,
        timer_path=_path_string(path),
    )


def _parse_book_info(
    obj: KrdsObject, path: Sequence[Union[str, int]]
) -> Optional[ReadingStats]:
    # book.info.store is: numberOfWords (Int64), percentOfBook (Double).
    if not obj.values:
        return None

    number_of_words = _as_int(obj.values[0])
    percent_of_book = (
        _as_float(obj.values[1]) if len(obj.values) >= 2 else None
    )
    if number_of_words is not None and number_of_words < 0:
        number_of_words = None
    return ReadingStats(
        book_words=number_of_words,
        book_word_coverage_percent=normalize_percent(percent_of_book),
        book_info_path=_path_string(path),
    )


def extract_reading_stats(document: KrdsDocument) -> ReadingStats:
    """Find timer and book word counters without requiring a full schema."""

    timer_candidates: List[Tuple[int, int, ReadingStats]] = []
    book_candidates: List[Tuple[int, ReadingStats]] = []

    for path, obj in iter_krds_objects(document.root):
        if obj.name == "timer.model":
            parsed = _parse_timer_model(obj, path)
            if parsed is not None:
                joined_path = _path_string(path)
                if "timer.data.store.v2" in joined_path:
                    priority = 3
                elif "timer.data.store" in joined_path:
                    priority = 2
                else:
                    priority = 1
                timer_candidates.append(
                    (
                        priority,
                        parsed.total_time_ms or -1,
                        parsed,
                    )
                )
        elif obj.name == "book.info.store":
            parsed = _parse_book_info(obj, path)
            if parsed is not None:
                book_candidates.append((parsed.book_words or -1, parsed))

    result = ReadingStats()
    if timer_candidates:
        # Prefer the larger accumulated time; the path priority is a tie
        # breaker for older files that store multiple copies of timer.model.
        best_timer = max(timer_candidates, key=lambda item: (item[1], item[0]))
        result = best_timer[2]

    if book_candidates:
        best_book = max(book_candidates, key=lambda item: item[0])
        result.book_words = best_book[1].book_words
        result.book_word_coverage_percent = (
            best_book[1].book_word_coverage_percent
        )
        result.book_info_path = best_book[1].book_info_path

    return result


def normalize_percent(value: Optional[float]) -> Optional[float]:
    if value is None or not math.isfinite(value) or value < 0:
        return None
    if value <= 1.0:
        return round(value * 100.0, 6)
    if value <= 100.0:
        return round(value, 6)
    return None


def _clamp_percent(value: Optional[float]) -> Optional[float]:
    if value is None or not math.isfinite(value):
        return None
    return round(max(0.0, min(100.0, value)), 6)


def calculate_wpm(stats: ReadingStats) -> Optional[float]:
    if stats.total_time_ms is None or stats.total_words is None:
        return None
    if stats.total_time_ms <= 0 or stats.total_words < 0:
        return None
    minutes = stats.total_time_ms / 60000.0
    if minutes <= 0:
        return None
    wpm = stats.total_words / minutes
    return round(wpm, 2) if math.isfinite(wpm) else None


def calculate_progress(stats: ReadingStats) -> Tuple[
    Optional[float], Optional[str], Optional[float], Optional[float]
]:
    """Return progress, source, timer progress, and word-based progress."""

    timer_percent = normalize_percent(stats.timer_percent)
    word_progress = None
    if (
        stats.total_words is not None
        and stats.total_words >= 0
        and stats.book_words is not None
        and stats.book_words > 0
    ):
        word_progress = _clamp_percent(
            stats.total_words / stats.book_words * 100.0
        )

    if timer_percent is not None and timer_percent > 0:
        return (
            timer_percent,
            "timer.model.totalPercent",
            timer_percent,
            word_progress,
        )
    if word_progress is not None and word_progress > 0:
        return (
            word_progress,
            "totalWords / book.info.store.numberOfWords",
            timer_percent,
            word_progress,
        )
    if timer_percent is not None:
        return (
            timer_percent,
            "timer.model.totalPercent",
            timer_percent,
            word_progress,
        )
    if word_progress is not None:
        return (
            word_progress,
            "totalWords / book.info.store.numberOfWords",
            timer_percent,
            word_progress,
        )
    return None, None, timer_percent, word_progress


@dataclass
class BookIdentity:
    title: Optional[str]
    author: Optional[str]
    ebook_path: Optional[Path]
    metadata_source: str


def _decode_text(raw: bytes, encoding: int) -> str:
    if not raw:
        return ""
    codec = "utf-8" if encoding == 65001 else "cp1252"
    try:
        return raw.decode(codec).replace("\x00", "").strip()
    except UnicodeDecodeError:
        return raw.decode(codec, errors="replace").replace("\x00", "").strip()


def parse_mobi_metadata(path: Path) -> Tuple[Optional[str], Optional[str]]:
    """Read title/author from a MOBI or AZW3 PalmDB record 0."""

    try:
        file_size = path.stat().st_size
        if file_size < 86:
            return None, None

        with path.open("rb") as handle:
            header = handle.read(86)
            if len(header) < 86:
                return None, None

            record_count = struct.unpack_from(">H", header, 76)[0]
            if record_count <= 0:
                return None, None

            handle.seek(78)
            offsets: List[int] = []
            for index in range(record_count):
                raw_offset = handle.read(8)
                if len(raw_offset) != 8:
                    break
                offsets.append(struct.unpack(">I", raw_offset[:4])[0])
            if not offsets:
                return None, None

            first_offset = offsets[0]
            if first_offset >= file_size:
                return None, None
            record_end = (
                offsets[1]
                if len(offsets) > 1
                else min(file_size, first_offset + 2_000_000)
            )
            if record_end <= first_offset:
                return None, None

            handle.seek(first_offset)
            record0 = handle.read(min(record_end - first_offset, 2_000_000))
    except OSError:
        return None, None

    if len(record0) < 32 or record0[16:20] != b"MOBI":
        return None, None

    try:
        mobi_header_length = struct.unpack_from(">I", record0, 20)[0]
        text_encoding = (
            struct.unpack_from(">I", record0, 28)[0]
            if len(record0) >= 32
            else 65001
        )
    except struct.error:
        return None, None

    if mobi_header_length < 24 or mobi_header_length > len(record0) - 16:
        return None, None

    pdb_title = _decode_text(header[0:32], text_encoding)
    exth_title: Optional[str] = None
    author: Optional[str] = None

    if len(record0) >= 132:
        exth_flag = struct.unpack_from(">I", record0, 128)[0]
        exth_offset = 16 + mobi_header_length
        if (
            exth_flag
            and exth_offset + 12 <= len(record0)
            and record0[exth_offset : exth_offset + 4] == b"EXTH"
        ):
            exth_length = struct.unpack_from(">I", record0, exth_offset + 4)[0]
            exth_count = struct.unpack_from(">I", record0, exth_offset + 8)[0]
            exth_end = min(len(record0), exth_offset + exth_length)
            cursor = exth_offset + 12

            for _ in range(exth_count):
                if cursor + 8 > exth_end:
                    break
                record_type, record_length = struct.unpack_from(
                    ">II", record0, cursor
                )
                cursor += 8
                if record_length < 8 or cursor + record_length - 8 > exth_end:
                    break
                payload = record0[cursor : cursor + record_length - 8]
                cursor += record_length - 8

                if record_type == 100:
                    author = _decode_text(payload, text_encoding)
                elif record_type == 503:
                    exth_title = _decode_text(payload, text_encoding)

    title = exth_title or pdb_title or None
    return title or None, author or None


def parse_filename_identity(stem: str) -> Tuple[Optional[str], Optional[str]]:
    cleaned = stem.strip()
    if not cleaned:
        return None, None

    # Amazon and Calibre commonly use "Title - Author" in the documents tree.
    parts = [part.strip() for part in cleaned.split(" - ", 1)]
    if len(parts) == 2 and parts[0] and parts[1]:
        return parts[0], parts[1]

    # Some exports use "Title (Author)".
    if cleaned.endswith(")") and "(" in cleaned:
        prefix, suffix = cleaned.rsplit("(", 1)
        author = suffix[:-1].strip()
        if prefix.strip() and author and len(author) <= 80:
            return prefix.strip(), author

    return cleaned, None


def _find_matching_ebook(
    search_dirs: Sequence[Path], stem: str
) -> Optional[Path]:
    wanted = stem.casefold()
    for directory in search_dirs:
        try:
            children = list(directory.iterdir())
        except OSError:
            continue
        for child in children:
            try:
                if (
                    child.is_file()
                    and child.suffix.lower() in EBOOK_EXTENSIONS
                    and child.stem.casefold() == wanted
                ):
                    return child
            except OSError:
                continue
    return None


def infer_book_identity(krds_path: Path, kindle_root: Path) -> BookIdentity:
    parent = krds_path.parent
    if parent.name.lower().endswith(".sdr"):
        stem = parent.name[:-4]
        search_dirs = [parent, parent.parent, kindle_root]
    else:
        stem = krds_path.stem
        search_dirs = [parent, parent.parent, kindle_root]

    ebook_path = _find_matching_ebook(search_dirs, stem)
    filename_title, filename_author = parse_filename_identity(stem)

    if ebook_path is not None:
        metadata_title, metadata_author = parse_mobi_metadata(ebook_path)
        if metadata_title or metadata_author:
            title = metadata_title or filename_title
            author = metadata_author or filename_author
            return BookIdentity(
                title=title,
                author=author,
                ebook_path=ebook_path,
                metadata_source="embedded MOBI/AZW3 metadata",
            )
        return BookIdentity(
            title=filename_title,
            author=filename_author,
            ebook_path=ebook_path,
            metadata_source="matching ebook filename",
        )

    return BookIdentity(
        title=filename_title,
        author=filename_author,
        ebook_path=None,
        metadata_source="sidecar filename",
    )


@dataclass
class FileResult:
    path: Path
    relative_path: str
    group_key: str
    status: str
    stats: Optional[ReadingStats] = None
    identity: Optional[BookIdentity] = None
    warnings: List[str] = field(default_factory=list)
    error: Optional[str] = None
    mtime: Optional[float] = None


@dataclass
class BookRow:
    group_key: str
    title: str
    author: Optional[str]
    ebook_path: Optional[str]
    metadata_source: str
    total_time_ms: Optional[int]
    total_words: Optional[int]
    book_words: Optional[int]
    wpm: Optional[float]
    progress_percent: Optional[float]
    progress_source: Optional[str]
    timer_percent: Optional[float]
    word_progress_percent: Optional[float]
    book_word_coverage_percent: Optional[float]
    timer_path: Optional[str]
    book_info_path: Optional[str]
    source_files: List[str]
    status: str
    warnings: List[str]

    def to_dict(self) -> Dict[str, Any]:
        total_hours = (
            round(self.total_time_ms / 3_600_000.0, 4)
            if self.total_time_ms is not None
            else None
        )
        total_minutes = (
            round(self.total_time_ms / 60_000.0, 4)
            if self.total_time_ms is not None
            else None
        )
        return {
            "title": self.title,
            "author": self.author,
            "metadata_source": self.metadata_source,
            "ebook_file": self.ebook_path,
            "total_time_ms": self.total_time_ms,
            "total_time_seconds": (
                round(self.total_time_ms / 1000.0, 3)
                if self.total_time_ms is not None
                else None
            ),
            "total_time_minutes": total_minutes,
            "total_time_hours": total_hours,
            "total_words": self.total_words,
            "book_words": self.book_words,
            "wpm": self.wpm,
            "progress_percent": self.progress_percent,
            "progress_source": self.progress_source,
            "timer_percent": self.timer_percent,
            "word_progress_percent": self.word_progress_percent,
            "book_word_coverage_percent": self.book_word_coverage_percent,
            "timer_path": self.timer_path,
            "book_info_path": self.book_info_path,
            "source_files": list(self.source_files),
            "status": self.status,
            "warnings": list(self.warnings),
        }


def _merge_reading_stats(results: Sequence[FileResult]) -> ReadingStats:
    candidates = [
        result
        for result in results
        if result.stats is not None
        and result.stats.total_time_ms is not None
    ]
    if not candidates:
        return ReadingStats()

    primary = max(
        candidates,
        key=lambda result: (
            result.stats.total_time_ms if result.stats else -1,
            result.mtime or 0.0,
        ),
    )
    merged = ReadingStats(**vars(primary.stats))

    for result in results:
        if result.stats is None:
            continue
        if merged.book_words is None and result.stats.book_words is not None:
            merged.book_words = result.stats.book_words
            merged.book_info_path = result.stats.book_info_path
        if (
            merged.book_word_coverage_percent is None
            and result.stats.book_word_coverage_percent is not None
        ):
            merged.book_word_coverage_percent = (
                result.stats.book_word_coverage_percent
            )
        if merged.timer_percent is None and result.stats.timer_percent is not None:
            merged.timer_percent = result.stats.timer_percent
            merged.timer_path = result.stats.timer_path

    return merged


def build_book_rows(
    groups: Dict[str, List[FileResult]],
    include_empty: bool,
) -> List[BookRow]:
    rows: List[BookRow] = []

    for group_key, results in groups.items():
        valid = [result for result in results if result.status == "ok"]
        partial = [result for result in results if result.status == "partial"]
        errors = [result for result in results if result.status == "error"]
        has_timer = any(
            result.stats is not None
            and result.stats.total_time_ms is not None
            for result in results
        )
        if not has_timer and not include_empty:
            continue

        stats = _merge_reading_stats(results)
        identity = next(
            (result.identity for result in results if result.identity),
            None,
        )
        identity = identity or BookIdentity(
            title=Path(group_key).stem or None,
            author=None,
            ebook_path=None,
            metadata_source="group fallback",
        )

        wpm = calculate_wpm(stats)
        progress, progress_source, timer_percent, word_progress = (
            calculate_progress(stats)
        )
        warnings: List[str] = []
        for result in results:
            warnings.extend(result.warnings)
        errors_for_row = [result.error for result in errors if result.error]

        if errors and not (valid or partial):
            status = "error"
        elif errors or partial or any(result.warnings for result in results):
            status = "partial"
        elif not has_timer:
            status = "empty"
        else:
            status = "ok"

        source_files = [result.relative_path for result in results]
        ebook_path = (
            str(identity.ebook_path) if identity.ebook_path is not None else None
        )

        rows.append(
            BookRow(
                group_key=group_key,
                title=identity.title or "(unknown)",
                author=identity.author,
                ebook_path=ebook_path,
                metadata_source=identity.metadata_source,
                total_time_ms=stats.total_time_ms,
                total_words=stats.total_words,
                book_words=stats.book_words,
                wpm=wpm,
                progress_percent=progress,
                progress_source=progress_source,
                timer_percent=timer_percent,
                word_progress_percent=word_progress,
                book_word_coverage_percent=stats.book_word_coverage_percent,
                timer_path=stats.timer_path,
                book_info_path=stats.book_info_path,
                source_files=source_files,
                status=status,
                warnings=warnings + errors_for_row,
            )
        )

    rows.sort(
        key=lambda row: (
            -(row.total_time_ms or 0),
            (row.title or "").casefold(),
        )
    )
    return rows


def _file_group_key(path: Path) -> str:
    parent = path.parent
    if parent.name.lower().endswith(".sdr"):
        return str(parent.resolve()).casefold()
    return str(path.with_suffix("").resolve()).casefold()


def scan_krds_files(
    root: Path,
    extensions: Iterable[str],
    max_depth: Optional[int],
    max_file_size: int,
    verbose: bool,
) -> Tuple[List[Path], List[str]]:
    normalized_extensions = {ext.lower() for ext in extensions}
    files: List[Path] = []
    issues: List[str] = []

    if root.is_file():
        if root.suffix.lower() in normalized_extensions:
            return [root], issues
        return [], ["Input path is a file but is not a supported KRDS sidecar"]

    def on_walk_error(error: OSError) -> None:
        issues.append("Cannot walk {}: {}".format(error.filename, error))

    for dirpath, dirnames, filenames in os.walk(
        str(root), topdown=True, onerror=on_walk_error, followlinks=False
    ):
        current = Path(dirpath)
        try:
            relative = current.relative_to(root)
            depth = len(relative.parts)
        except ValueError:
            depth = 0

        if max_depth is not None and depth >= max_depth:
            dirnames[:] = []

        dirnames.sort(key=str.casefold)
        filenames.sort(key=str.casefold)

        for filename in filenames:
            path = current / filename
            if path.suffix.lower() not in normalized_extensions:
                continue
            try:
                size = path.stat().st_size
            except OSError as exc:
                issues.append("Cannot stat {}: {}".format(path, exc))
                continue
            if size > max_file_size:
                issues.append(
                    "Skipped {} because it is {} bytes (limit {})".format(
                        path, size, max_file_size
                    )
                )
                continue
            files.append(path)
            if verbose:
                print("Found KRDS: {}".format(path), file=sys.stderr)

    return files, issues


def analyze_file(
    path: Path,
    root: Path,
    identity: BookIdentity,
    verbose: bool,
) -> FileResult:
    try:
        relative_path = str(path.resolve().relative_to(root.resolve()))
    except ValueError:
        relative_path = str(path)

    try:
        mtime = path.stat().st_mtime
    except OSError:
        mtime = None

    try:
        data = path.read_bytes()
    except OSError as exc:
        return FileResult(
            path=path,
            relative_path=relative_path,
            group_key=_file_group_key(path),
            status="error",
            identity=identity,
            error="Cannot read file: {}".format(exc),
            mtime=mtime,
        )

    try:
        document = KrdsDecoder(data).decode_document()
    except KrdsError as exc:
        return FileResult(
            path=path,
            relative_path=relative_path,
            group_key=_file_group_key(path),
            status="error",
            identity=identity,
            error=str(exc),
            mtime=mtime,
        )

    stats = extract_reading_stats(document)
    warnings = list(document.warnings)
    if stats.total_time_ms is None:
        warnings.append("timer.model with totalTime was not found")
        status = "empty"
    else:
        status = "partial" if document.partial else "ok"

    if verbose:
        print(
            "Parsed {}: status={}, warnings={}".format(
                path, status, len(warnings)
            ),
            file=sys.stderr,
        )

    return FileResult(
        path=path,
        relative_path=relative_path,
        group_key=_file_group_key(path),
        status=status,
        stats=stats,
        identity=identity,
        warnings=warnings,
        mtime=mtime,
    )


def build_report(
    root: Path,
    rows: Sequence[BookRow],
    file_results: Sequence[FileResult],
    scan_issues: Sequence[str],
) -> Dict[str, Any]:
    total_time = sum(row.total_time_ms or 0 for row in rows)
    total_words = sum(row.total_words or 0 for row in rows)
    total_minutes = total_time / 60000.0
    weighted_wpm = (
        round(total_words / total_minutes, 2) if total_minutes > 0 else None
    )

    status_counts: Dict[str, int] = {}
    for row in rows:
        status_counts[row.status] = status_counts.get(row.status, 0) + 1

    file_status_counts: Dict[str, int] = {}
    for result in file_results:
        file_status_counts[result.status] = (
            file_status_counts.get(result.status, 0) + 1
        )

    return {
        "schema_version": "1.0",
        "tool_version": VERSION,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "kindle_path": str(root),
        "summary": {
            "books": len(rows),
            "total_time_ms": total_time,
            "total_time_hours": round(total_time / 3_600_000.0, 4),
            "total_words": total_words,
            "weighted_average_wpm": weighted_wpm,
            "row_status_counts": status_counts,
            "krds_files": len(file_results),
            "file_status_counts": file_status_counts,
            "scan_issue_count": len(scan_issues),
        },
        "books": [row.to_dict() for row in rows],
        "files": [
            {
                "path": str(result.path),
                "relative_path": result.relative_path,
                "status": result.status,
                "error": result.error,
                "warnings": list(result.warnings),
                "mtime": result.mtime,
            }
            for result in file_results
        ],
        "scan_issues": list(scan_issues),
        "semantics": {
            "wpm": "totalWords / (totalTimeMs / 60000)",
            "progress": (
                "timer.model.totalPercent when usable; otherwise "
                "totalWords / book.info.store.numberOfWords"
            ),
            "book_word_coverage_percent": (
                "book.info.store.percentOfBook, not reader progress"
            ),
        },
    }


CSV_FIELDS = [
    "title",
    "author",
    "total_time_ms",
    "total_time_seconds",
    "total_time_minutes",
    "total_time_hours",
    "total_words",
    "book_words",
    "wpm",
    "progress_percent",
    "progress_source",
    "timer_percent",
    "word_progress_percent",
    "book_word_coverage_percent",
    "timer_path",
    "book_info_path",
    "metadata_source",
    "ebook_file",
    "source_files",
    "status",
    "warnings",
]


def write_csv(path: Path, rows: Sequence[BookRow]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_FIELDS)
        writer.writeheader()
        for row in rows:
            data = row.to_dict()
            record = {
                key: (
                    "; ".join(data[key])
                    if key in {"source_files", "warnings"}
                    else data.get(key)
                )
                for key in CSV_FIELDS
            }
            writer.writerow(record)


def write_json(path: Path, report: Dict[str, Any]) -> None:
    with path.open("w", encoding="utf-8", newline="\n") as handle:
        json.dump(report, handle, ensure_ascii=False, indent=2)
        handle.write("\n")


def _format_number(value: Optional[float], digits: int = 2) -> str:
    if value is None:
        return "--"
    return ("{:." + str(digits) + "f}").format(value)


def _format_duration(milliseconds: Optional[int]) -> str:
    if milliseconds is None:
        return "--"
    total_seconds = max(0, int(round(milliseconds / 1000.0)))
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    if hours:
        return "{}h {}m".format(hours, minutes)
    if minutes:
        return "{}m {}s".format(minutes, seconds)
    return "{}s".format(seconds)


def write_html(path: Path, report: Dict[str, Any], rows: Sequence[BookRow]) -> None:
    summary = report["summary"]
    generated = html.escape(report["generated_at"])
    root = html.escape(report["kindle_path"])
    table_rows: List[str] = []

    for index, row in enumerate(rows, start=1):
        progress_value = row.progress_percent
        progress_width = (
            max(0.0, min(100.0, progress_value))
            if progress_value is not None
            else 0.0
        )
        progress_text = (
            "{:.1f}%".format(progress_value)
            if progress_value is not None
            else "--"
        )
        search_text = " ".join(
            filter(
                None,
                [
                    row.title,
                    row.author or "",
                    row.ebook_path or "",
                    row.progress_source or "",
                    row.status,
                ],
            )
        ).casefold()
        status_class = "status-{}".format(
            row.status if row.status in {"ok", "partial", "empty", "error"} else "empty"
        )
        warnings_text = " | ".join(row.warnings)
        source_text = " | ".join(row.source_files)

        table_rows.append(
            """
        <tr data-search="{search}">
          <td class="rank">{rank}</td>
          <td class="book">
            <strong>{title}</strong>
            <span>{ebook}</span>
          </td>
          <td>{author}</td>
          <td data-sort-value="{time_value}">{time_text}</td>
          <td data-sort-value="{wpm_value}">{wpm}</td>
          <td data-sort-value="{progress_value}">
            <div class="progress-row">
              <span>{progress_text}</span>
              <div class="bar"><i style="width:{progress_width:.2f}%"></i></div>
            </div>
            <small>{progress_source}</small>
          </td>
          <td data-sort-value="{words_value}">{words}</td>
          <td data-sort-value="{book_words_value}">{book_words}</td>
          <td><span class="status {status_class}">{status}</span></td>
          <td class="details" title="{details_title}">{source}</td>
        </tr>
        """.format(
                search=html.escape(search_text, quote=True),
                rank=index,
                title=html.escape(row.title or ""),
                ebook=html.escape(row.ebook_path or ""),
                author=html.escape(row.author or "--"),
                time_value=row.total_time_ms if row.total_time_ms is not None else -1,
                time_text=_format_duration(row.total_time_ms),
                wpm_value=row.wpm if row.wpm is not None else -1,
                wpm=_format_number(row.wpm),
                progress_value=(
                    progress_value if progress_value is not None else -1
                ),
                progress_text=progress_text,
                progress_width=progress_width,
                progress_source=html.escape(row.progress_source or ""),
                status_class=status_class,
                words_value=row.total_words if row.total_words is not None else -1,
                words=(
                    "{:,}".format(row.total_words)
                    if row.total_words is not None
                    else "--"
                ),
                book_words_value=(
                    row.book_words if row.book_words is not None else -1
                ),
                book_words=(
                    "{:,}".format(row.book_words)
                    if row.book_words is not None
                    else "--"
                ),
                status=html.escape(row.status),
                details_title=html.escape(
                    (warnings_text + " " + source_text).strip(), quote=True
                ),
                source=html.escape(source_text),
            )
        )

    document = """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Kindle 阅读统计</title>
  <style>
    :root {
      color-scheme: light;
      --ink: #1f2933;
      --muted: #667085;
      --line: #e4e7ec;
      --paper: #f7f8fa;
      --blue: #1769aa;
      --green: #147d64;
      --amber: #b54708;
      --red: #b42318;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: var(--paper);
      color: var(--ink);
      font: 14px/1.45 -apple-system, BlinkMacSystemFont, "Segoe UI",
        "Microsoft YaHei", sans-serif;
    }
    .page {
      width: min(1500px, calc(100% - 32px));
      margin: 24px auto 48px;
    }
    h1 { margin: 0 0 6px; font-size: 28px; line-height: 1.2; }
    .meta { color: var(--muted); overflow-wrap: anywhere; }
    .summary {
      display: grid;
      grid-template-columns: repeat(4, minmax(150px, 1fr));
      gap: 12px;
      margin: 20px 0;
    }
    .metric {
      background: #fff;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 14px 16px;
    }
    .metric span { color: var(--muted); display: block; margin-bottom: 4px; }
    .metric strong { font-size: 22px; }
    .toolbar {
      display: flex;
      gap: 12px;
      align-items: center;
      margin: 18px 0 10px;
    }
    input[type="search"] {
      width: min(420px, 100%);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 10px 12px;
      font: inherit;
      background: #fff;
    }
    .table-wrap {
      background: #fff;
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: auto;
    }
    table { width: 100%; border-collapse: collapse; min-width: 1080px; }
    th, td {
      border-bottom: 1px solid var(--line);
      padding: 11px 12px;
      text-align: left;
      vertical-align: middle;
    }
    th {
      position: sticky;
      top: 0;
      z-index: 1;
      background: #fbfcfd;
      color: #344054;
      cursor: pointer;
      white-space: nowrap;
    }
    tbody tr:hover { background: #f9fafb; }
    td.book strong { display: block; max-width: 420px; }
    td.book span, td small {
      display: block;
      color: var(--muted);
      font-size: 12px;
      overflow-wrap: anywhere;
    }
    .progress-row {
      display: grid;
      grid-template-columns: 52px 1fr;
      gap: 8px;
      align-items: center;
      min-width: 180px;
    }
    .bar {
      height: 8px;
      border-radius: 999px;
      background: #eaecf0;
      overflow: hidden;
    }
    .bar i {
      display: block;
      height: 100%;
      background: linear-gradient(90deg, #1769aa, #2e90fa);
    }
    .status {
      display: inline-block;
      border-radius: 999px;
      padding: 3px 8px;
      font-size: 12px;
      font-weight: 600;
    }
    .status-ok { color: var(--green); background: #ecfdf3; }
    .status-partial { color: var(--amber); background: #fffaeb; }
    .status-empty { color: var(--muted); background: #f2f4f7; }
    .status-error { color: var(--red); background: #fef3f2; }
    .details {
      max-width: 320px;
      color: var(--muted);
      font-size: 12px;
      overflow-wrap: anywhere;
    }
    .empty-state { padding: 24px; color: var(--muted); }
    @media (max-width: 900px) {
      .summary { grid-template-columns: repeat(2, minmax(140px, 1fr)); }
      .page { width: min(100% - 20px, 1500px); margin-top: 16px; }
      h1 { font-size: 24px; }
    }
  </style>
</head>
<body>
  <main class="page">
    <h1>Kindle 阅读统计</h1>
    <div class="meta">扫描目录：__ROOT__ · 生成时间：__GENERATED__</div>
    <section class="summary">
      <div class="metric"><span>书籍</span><strong>__BOOK_COUNT__</strong></div>
      <div class="metric"><span>累计阅读</span><strong>__HOURS__ 小时</strong></div>
      <div class="metric"><span>累计词数</span><strong>__WORDS__</strong></div>
      <div class="metric"><span>加权平均 WPM</span><strong>__WPM__</strong></div>
    </section>
    <div class="toolbar">
      <input id="search" type="search" placeholder="搜索书名、作者或文件">
    </div>
    <div class="table-wrap">
      <table id="stats-table">
        <thead>
          <tr>
            <th data-type="number">#</th>
            <th data-type="text">书名</th>
            <th data-type="text">作者</th>
            <th data-type="number">阅读时长</th>
            <th data-type="number">WPM</th>
            <th data-type="number">进度</th>
            <th data-type="number">已读词数</th>
            <th data-type="number">书籍词数</th>
            <th data-type="text">状态</th>
            <th data-type="text">来源</th>
          </tr>
        </thead>
        <tbody>
          __ROWS__
        </tbody>
      </table>
      <div class="empty-state" id="empty-state" hidden>没有匹配记录</div>
    </div>
  </main>
  <script>
    const search = document.getElementById("search");
    const table = document.getElementById("stats-table");
    const tbody = table.querySelector("tbody");
    const emptyState = document.getElementById("empty-state");

    search.addEventListener("input", () => {
      const query = search.value.trim().toLocaleLowerCase();
      let visible = 0;
      for (const row of tbody.querySelectorAll("tr")) {
        const match = !query || row.dataset.search.includes(query);
        row.hidden = !match;
        if (match) visible += 1;
      }
      emptyState.hidden = visible !== 0;
    });

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
            ? Number(a.dataset.sortValue ?? a.textContent.replace(/[^\\d.-]/g, ""))
            : a.textContent.trim();
          const bv = type === "number"
            ? Number(b.dataset.sortValue ?? b.textContent.replace(/[^\\d.-]/g, ""))
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
    document = (
        document.replace("__ROOT__", root)
        .replace("__GENERATED__", generated)
        .replace("__BOOK_COUNT__", "{:,}".format(summary["books"]))
        .replace("__HOURS__", "{:,.1f}".format(summary["total_time_hours"]))
        .replace("__WORDS__", "{:,}".format(summary["total_words"]))
        .replace("__WPM__", _format_number(summary["weighted_average_wpm"]))
        .replace(
            "__ROWS__",
            "\n".join(table_rows)
            or '<tr><td colspan="10" class="empty-state">没有书籍记录</td></tr>',
        )
    )

    with path.open("w", encoding="utf-8", newline="\n") as handle:
        handle.write(document)


def _truncate(value: str, width: int) -> str:
    if len(value) <= width:
        return value
    if width <= 1:
        return value[:width]
    return value[: width - 1] + "~"


def print_console_summary(report: Dict[str, Any], rows: Sequence[BookRow]) -> None:
    summary = report["summary"]
    print()
    print(
        "Kindle: {} | books: {} | KRDS files: {} | warnings: {}".format(
            report["kindle_path"],
            summary["books"],
            summary["krds_files"],
            summary["scan_issue_count"],
        )
    )
    print(
        "{:<3} {:<34} {:<22} {:>12} {:>8} {:>9}".format(
            "#", "Title", "Author", "Time", "WPM", "Progress"
        )
    )
    print("-" * 96)
    for index, row in enumerate(rows[:50], start=1):
        progress = (
            "{:.1f}%".format(row.progress_percent)
            if row.progress_percent is not None
            else "--"
        )
        print(
            "{:<3} {:<34} {:<22} {:>12} {:>8} {:>9}".format(
                index,
                _truncate(row.title or "", 34),
                _truncate(row.author or "", 22),
                _format_duration(row.total_time_ms),
                _format_number(row.wpm),
                progress,
            )
        )
    if len(rows) > 50:
        print("... {} more rows are available in the exported files".format(len(rows) - 50))
    print()


def parse_formats(raw: str) -> List[str]:
    if raw.strip().lower() == "all":
        return ["csv", "json", "html"]
    allowed = {"csv", "json", "html"}
    formats: List[str] = []
    for item in raw.split(","):
        value = item.strip().lower()
        if not value:
            continue
        if value not in allowed:
            raise ValueError(
                "Unknown format '{}'. Use all, csv, json, or html.".format(value)
            )
        if value not in formats:
            formats.append(value)
    if not formats:
        raise ValueError("At least one output format is required")
    return formats


def parse_extensions(raw: str) -> List[str]:
    result: List[str] = []
    for item in raw.split(","):
        value = item.strip().lower()
        if not value:
            continue
        if not value.startswith("."):
            value = "." + value
        if value not in result:
            result.append(value)
    return result or sorted(KRDS_EXTENSIONS)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Extract per-book reading time, WPM, and progress from Kindle "
            "KRDS sidecar files."
        )
    )
    parser.add_argument(
        "--kindle-path",
        required=True,
        help="Mounted Kindle root or its documents directory.",
    )
    parser.add_argument(
        "--output-dir",
        default=".",
        help="Directory for exported files. Defaults to the current directory.",
    )
    parser.add_argument(
        "--prefix",
        default="kindle-reading-stats",
        help="Output filename prefix.",
    )
    parser.add_argument(
        "--format",
        default="all",
        help="Output format: all, csv, json, html, or a comma-separated combination.",
    )
    parser.add_argument(
        "--extensions",
        default=",".join(sorted(KRDS_EXTENSIONS)),
        help="Comma-separated KRDS sidecar extensions to scan.",
    )
    parser.add_argument(
        "--max-depth",
        type=int,
        default=None,
        help="Maximum recursive depth below the input path. Default is unlimited.",
    )
    parser.add_argument(
        "--max-file-size-mb",
        type=float,
        default=32.0,
        help="Skip sidecars larger than this size. Default: 32 MB.",
    )
    parser.add_argument(
        "--include-empty",
        action="store_true",
        help="Include sidecar groups that have no usable timer values.",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Do not print the console summary table.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print per-file scanning and parsing details to stderr.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version="%(prog)s {}".format(VERSION),
    )
    return parser


def run(args: argparse.Namespace) -> int:
    try:
        formats = parse_formats(args.format)
        extensions = parse_extensions(args.extensions)
    except ValueError as exc:
        print("error: {}".format(exc), file=sys.stderr)
        return 2

    if args.max_depth is not None and args.max_depth < 0:
        print("error: --max-depth cannot be negative", file=sys.stderr)
        return 2
    if args.max_file_size_mb <= 0:
        print("error: --max-file-size-mb must be positive", file=sys.stderr)
        return 2

    root = Path(args.kindle_path).expanduser()
    if not root.exists():
        print("error: Kindle path does not exist: {}".format(root), file=sys.stderr)
        return 2
    root = root.resolve()

    output_dir = Path(args.output_dir).expanduser()
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        print(
            "error: cannot create output directory {}: {}".format(output_dir, exc),
            file=sys.stderr,
        )
        return 2

    files, scan_issues = scan_krds_files(
        root=root,
        extensions=extensions,
        max_depth=args.max_depth,
        max_file_size=int(args.max_file_size_mb * 1024 * 1024),
        verbose=args.verbose,
    )

    groups: Dict[str, List[FileResult]] = {}
    identity_cache: Dict[str, BookIdentity] = {}
    for path in files:
        group_key = _file_group_key(path)
        identity = identity_cache.get(group_key)
        if identity is None:
            identity = infer_book_identity(path, root)
            identity_cache[group_key] = identity
        result = analyze_file(
            path=path,
            root=root,
            identity=identity,
            verbose=args.verbose,
        )
        groups.setdefault(result.group_key, []).append(result)

    rows = build_book_rows(groups, include_empty=args.include_empty)
    file_results = [
        result
        for group_results in groups.values()
        for result in group_results
    ]
    report = build_report(root, rows, file_results, scan_issues)

    output_paths: List[Path] = []
    prefix = args.prefix.strip() or "kindle-reading-stats"
    if "csv" in formats:
        path = output_dir / "{}.csv".format(prefix)
        write_csv(path, rows)
        output_paths.append(path)
    if "json" in formats:
        path = output_dir / "{}.json".format(prefix)
        write_json(path, report)
        output_paths.append(path)
    if "html" in formats:
        path = output_dir / "{}.html".format(prefix)
        write_html(path, report, rows)
        output_paths.append(path)

    if not args.quiet:
        print_console_summary(report, rows)
        print("Exported:")
        for path in output_paths:
            print("  {}".format(path))

    if not files:
        print(
            "warning: no recognized KRDS sidecar files were found under {}".format(
                root
            ),
            file=sys.stderr,
        )
        return 2
    return 0


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)
    return run(args)


if __name__ == "__main__":
    raise SystemExit(main())
