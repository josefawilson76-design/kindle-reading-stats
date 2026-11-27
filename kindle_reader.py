#!/usr/bin/env python3
"""
Standalone Kindle on-device reading-statistics extractor.

This script ports the USB/KRDS portion of the Calibre "Kindle Reading
Dashboard" plugin to a standard-library-only command-line tool. It scans a
mounted Kindle for reader-data sidecars, decodes their KRDS payloads, and
exports one CSV row per book.

The KRDS parser is derived from John Howell's GPL v3 implementation. That
portion retains its original copyright and is distributed under GPL v3.

Example:
    python kindle_reader.py --path /media/user/Kindle --output readings.csv
    python kindle_reader.py --path E:\\ --output readings.csv
"""

from __future__ import annotations

import argparse
import collections
import contextlib
import csv
import datetime
import hashlib
import io
import json
import logging
import os
import re
import shutil
import site
import struct
import subprocess
import sys
import tempfile
import time
import unicodedata
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple


USER_SITE = site.getusersitepackages()
if USER_SITE and os.path.isdir(USER_SITE) and USER_SITE not in sys.path:
    sys.path.append(USER_SITE)

OPTIONAL_VENDOR_DIR = os.environ.get("KINDLE_READER_VENDOR", "").strip()
if OPTIONAL_VENDOR_DIR and os.path.isdir(OPTIONAL_VENDOR_DIR):
    sys.path.insert(0, OPTIONAL_VENDOR_DIR)


__license__ = "GPL v3"
__copyright__ = (
    "KRDS parser: 2019, John Howell <jhowell@acm.org>; "
    "standalone adaptation of Kindle Reading Dashboard"
)


# ---------------------------------------------------------------------------
# KRDS binary parser
# ---------------------------------------------------------------------------

class KindleReaderDataStore(object):
    """Decode a Kindle reader-data-store sidecar into Python objects."""

    SIGNATURE = b"\x00\x00\x00\x00\x00\x1A\xB1\x26"

    DATATYPE_BOOLEAN = 0
    DATATYPE_INT = 1
    DATATYPE_LONG = 2
    DATATYPE_UTF = 3
    DATATYPE_DOUBLE = 4
    DATATYPE_SHORT = 5
    DATATYPE_FLOAT = 6
    DATATYPE_BYTE = 7
    DATATYPE_CHAR = 9
    DATATYPE_OBJECT_BEGIN = -2
    DATATYPE_OBJECT_END = -1

    ANNOT_CLASS_NAMES = {
        0: "annotation.personal.bookmark",
        1: "annotation.personal.highlight",
        2: "annotation.personal.note",
        3: "annotation.personal.clip_article",
        10: "annotation.personal.handwritten_note",
        11: "annotation.personal.sticky_note",
        13: "annotation.personal.underline",
    }

    def __init__(self, log: logging.Logger, data: bytes) -> None:
        self.krds: Optional[Deserializer] = None
        self.log = log
        self.data = data

    def deserialize(self) -> Dict[str, Any]:
        self.krds = Deserializer(self.data)

        signature = self.krds.extract(len(self.SIGNATURE))
        if signature != self.SIGNATURE:
            raise ValueError("KindleReaderDataStore signature is incorrect")

        first_value = self.decode_next()
        if first_value != 1:
            raise ValueError("first_value = %s" % repr(first_value))

        value_count = self.decode_next()
        value: Dict[str, Any] = collections.OrderedDict()

        for _ in range(value_count):
            try:
                decoded = self.decode_next()
            except Exception:
                self.log.info(
                    "KindleReaderDataStore decode failed at offset %d",
                    len(self.data) - len(self.krds),
                )
                self.log.info("partial value = %s", repr(value))
                raise

            for key, item in decoded.items():
                if key in value:
                    raise ValueError(
                        "KindleReaderDataStore has duplicate item %s" % key
                    )
                value[key] = item

        if len(self.krds) != 0:
            self.log.error(
                "KindleReaderDataStore has %d bytes of extra data",
                len(self.krds),
            )

        return value

    def decode_next(self, datatype: Optional[int] = None) -> Any:
        if datatype is None:
            datatype = self.krds.unpack("b")

        if datatype == self.DATATYPE_BOOLEAN:
            raw = self.krds.unpack("b")
            if raw == 0:
                return False
            if raw == 1:
                return True
            raise ValueError("Unknown boolean value %d" % raw)

        if datatype == self.DATATYPE_INT:
            return self.krds.unpack(">l")

        if datatype == self.DATATYPE_LONG:
            return self.krds.unpack(">q")

        if datatype == self.DATATYPE_UTF:
            if self.decode_next(self.DATATYPE_BOOLEAN):
                return ""
            return self.krds.extract(self.krds.unpack(">H")).decode("utf-8")

        if datatype == self.DATATYPE_DOUBLE:
            return self.krds.unpack(">d")

        if datatype == self.DATATYPE_SHORT:
            return self.krds.unpack(">h")

        if datatype == self.DATATYPE_FLOAT:
            return self.krds.unpack(">f")

        if datatype == self.DATATYPE_BYTE:
            return self.krds.unpack("b")

        if datatype == self.DATATYPE_CHAR:
            return self.krds.unpack("c").decode("utf-8")

        if datatype == self.DATATYPE_OBJECT_BEGIN:
            name = self.decode_next(self.DATATYPE_UTF)
            values: List[Any] = []

            while (
                self.krds.unpack("b", advance=False)
                != self.DATATYPE_OBJECT_END
            ):
                values.append(self.decode_next())

            self.krds.unpack("b")
            return self.decode_object(name, values)

        raise ValueError("Unknown datatype %d" % datatype)

    def decode_object(self, name: str, values: List[Any]) -> Dict[str, Any]:
        raw = list(values)
        try:
            return {name: self._decode_known(name, values)}
        except Exception as exc:
            self.log.warning(
                "raw fallback for structure %s: %s", name, repr(exc)
            )
            return {name: raw}

    def _decode_known(self, name: str, values: List[Any]) -> Any:
        obj: Any = collections.OrderedDict()

        if name in {
            "clock.data.store",
            "dictionary",
            "lpu",
            "pdf.contrast",
            "sync_lpr",
            "tpz.line.spacing",
            "XRAY_OTA_UPDATE_STATE",
            "XRAY_SHOWING_SPOILERS",
            "XRAY_SORTING_STATE",
            "XRAY_TAB_STATE",
        }:
            obj = values.pop(0)

        elif name in {
            "dict.prefs.v2",
            "EndActions",
            "ReaderMetrics",
            "StartActions",
            "Translator",
            "Wikipedia",
        }:
            for _ in range(values.pop(0)):
                key = values.pop(0)
                obj[key] = values.pop(0)

        elif name in {
            "buy.asin.response.data",
            "next.in.series.info.data",
            "price.info.data",
        }:
            obj = values.pop(0)

        elif name == "erl":
            obj = self.decode_position(values.pop(0))

        elif name == "lpr":
            version = values.pop(0)
            if isinstance(version, str):
                obj["position"] = self.decode_position(version)
            elif version <= 2:
                obj["position"] = self.decode_position(values.pop(0))
                timestamp = values.pop(0)
                obj["time"] = (
                    datetime.datetime.fromtimestamp(
                        timestamp / 1000.0
                    ).isoformat()
                    if timestamp != -1
                    else None
                )
            else:
                raise ValueError("Unknown lpr version %d" % version)

        elif name in {"fpr", "updated_lpr"}:
            obj["position"] = self.decode_position(values.pop(0))
            timestamp = values.pop(0)
            obj["time"] = (
                datetime.datetime.fromtimestamp(
                    timestamp / 1000.0
                ).isoformat()
                if timestamp != -1
                else None
            )
            timezone = values.pop(0)
            obj["timeZoneOffset"] = timezone if timezone != -1 else None
            obj["country"] = values.pop(0)
            obj["device"] = values.pop(0)

        elif name == "annotation.cache.object":
            for _ in range(values.pop(0)):
                annotation_type = values.pop(0)
                annot_class_name = self.ANNOT_CLASS_NAMES.get(annotation_type)
                if annot_class_name is None:
                    raise ValueError(
                        "Unknown annotation type %d" % annotation_type
                    )

                annotations = []
                for annotation in values.pop(0)["saved.avl.interval.tree"]:
                    if (
                        len(annotation) != 1
                        or annot_class_name not in annotation
                    ):
                        raise ValueError(
                            "Unknown annotation format: %s" % repr(annotation)
                        )
                    annotations.append(annotation[annot_class_name])

                obj[annot_class_name] = annotations

        elif name == "saved.avl.interval.tree":
            obj = [values.pop(0) for _ in range(values.pop(0))]

        elif name in self.ANNOT_CLASS_NAMES.values():
            obj["startPosition"] = self.decode_position(values.pop(0))
            obj["endPosition"] = self.decode_position(values.pop(0))
            obj["creationTime"] = datetime.datetime.fromtimestamp(
                values.pop(0) / 1000.0
            ).isoformat()
            obj["lastModificationTime"] = datetime.datetime.fromtimestamp(
                values.pop(0) / 1000.0
            ).isoformat()
            obj["template"] = values.pop(0)

            if name == "annotation.personal.note":
                obj["note"] = values.pop(0)
            elif name == "annotation.personal.handwritten_note":
                obj["handwritten_note_nbk_ref"] = values.pop(0)
            elif name == "annotation.personal.sticky_note":
                obj["sticky_note_nbk_ref"] = values.pop(0)

        elif name == "apnx.key":
            obj["asin"] = values.pop(0)
            obj["cdeType"] = values.pop(0)
            obj["sidecarAvailable"] = values.pop(0)
            obj["oPNToPosition"] = [
                values.pop(0) for _ in range(values.pop(0))
            ]
            obj["first"] = values.pop(0)
            obj["unknown1"] = values.pop(0)
            obj["unknown2"] = values.pop(0)
            obj["pageMap"] = values.pop(0)

        elif name == "fixed.layout.data":
            obj["unknown1"] = values.pop(0)
            obj["unknown2"] = values.pop(0)
            obj["unknown3"] = values.pop(0)

        elif name == "sharing.limits":
            obj["accumulated"] = values.pop(0)

        elif name == "language.store":
            obj["language"] = values.pop(0)
            obj["unknown1"] = values.pop(0)

        elif name == "periodicals.view.state":
            obj["unknown1"] = values.pop(0)
            obj["unknown2"] = values.pop(0)

        elif name == "font.prefs":
            obj["typeface"] = values.pop(0)
            obj["lineSp"] = values.pop(0)
            obj["size"] = values.pop(0)
            obj["align"] = values.pop(0)
            obj["insetTop"] = values.pop(0)
            obj["insetLeft"] = values.pop(0)
            obj["insetBottom"] = values.pop(0)
            obj["insetRight"] = values.pop(0)
            obj["unknown1"] = values.pop(0)

            if values:
                obj["bold"] = values.pop(0)
            if values:
                obj["userSideloadableFont"] = values.pop(0)
            if values:
                obj["customFontIndex"] = values.pop(0)
            if values:
                obj["mobi7SystemFont"] = values.pop(0)
            if values:
                obj["mobi7RestoreFont"] = values.pop(0)
            if values:
                obj["readingPresetSelected"] = values.pop(0)
            if values:
                obj["unknown2"] = values.pop(0)

        elif name == "purchase.state.data":
            obj["state"] = values.pop(0)
            obj["time"] = datetime.datetime.fromtimestamp(
                values.pop(0) / 1000.0
            ).isoformat()

        elif name == "timer.data.store":
            obj["on"] = values.pop(0)
            obj["readingTimerModel"] = values.pop(0)
            obj["version"] = values.pop(0)

        elif name == "timer.data.store.v2":
            obj["on"] = values.pop(0)
            obj["readingTimerModel"] = values.pop(0)
            obj["version"] = values.pop(0)
            obj["lastOption"] = values.pop(0)

        elif name == "timer.model":
            obj["version"] = values.pop(0)
            obj["totalTime"] = values.pop(0)
            obj["totalWords"] = values.pop(0)
            obj["totalPercent"] = values.pop(0)
            obj["averageCalculator"] = values.pop(0)[
                "timer.average.calculator"
            ]

        elif name == "timer.average.calculator":
            obj["samples1"] = [
                values.pop(0) for _ in range(values.pop(0))
            ]
            obj["samples2"] = [
                values.pop(0) for _ in range(values.pop(0))
            ]
            obj["normalDistributions"] = [
                values.pop(0)[
                    "timer.average.calculator.distribution.normal"
                ]
                for _ in range(values.pop(0))
            ]
            obj["outliers"] = [
                values.pop(0)["timer.average.calculator.outliers"]
                for _ in range(values.pop(0))
            ]

        elif name == "timer.average.calculator.distribution.normal":
            obj["count"] = values.pop(0)
            obj["sum"] = values.pop(0)
            obj["sumOfSquares"] = values.pop(0)

        elif name == "timer.average.calculator.outliers":
            obj = [values.pop(0) for _ in range(values.pop(0))]

        elif name == "book.info.store":
            obj["numberOfWords"] = values.pop(0)
            obj["percentOfBook"] = values.pop(0)

        elif name == "page.history.store":
            obj = [
                values.pop(0)["page.history.record"]
                for _ in range(values.pop(0))
            ]

        elif name == "page.history.record":
            obj["position"] = self.decode_position(values.pop(0))
            obj["time"] = datetime.datetime.fromtimestamp(
                values.pop(0) / 1000.0
            ).isoformat()

        elif name == "reader.state.preferences":
            obj["fontPreferences"] = values.pop(0)
            obj["leftMargin"] = values.pop(0)
            obj["rightMargin"] = values.pop(0)
            obj["topMargin"] = values.pop(0)
            obj["bottomMargin"] = values.pop(0)
            obj["unknown1"] = values.pop(0)

        else:
            self.log.error("Unknown data structure %s", name)
            obj = values
            values = []

        if values:
            self.log.warning(
                "Excess values for structure %s (newer firmware?): %s",
                name,
                repr(values),
            )
            if isinstance(obj, dict):
                obj["_unparsed_trailing"] = list(values)
            values[:] = []

        return obj

    @staticmethod
    def decode_position(position: Any) -> Any:
        return position


class Deserializer(object):
    """Bounds-checked big-endian reader used by the KRDS parser."""

    def __init__(self, data: bytes) -> None:
        self.buffer = data
        self.offset = 0

    def unpack(self, fmt: str, advance: bool = True) -> Any:
        result = struct.unpack_from(fmt, self.buffer, self.offset)[0]
        if advance:
            self.offset += struct.calcsize(fmt)
        return result

    def extract(
        self,
        size: Optional[int] = None,
        upto: Optional[int] = None,
        advance: bool = True,
    ) -> bytes:
        if size is None:
            size = len(self) if upto is None else (upto - self.offset)

        data = self.buffer[self.offset : self.offset + size]
        if len(data) < size or size < 0:
            raise ValueError(
                "Deserializer: Insufficient data "
                "(need %d bytes, have %d bytes)" % (size, len(data))
            )

        if advance:
            self.offset += size
        return data

    def __len__(self) -> int:
        return len(self.buffer) - self.offset


# ---------------------------------------------------------------------------
# Kindle data discovery and extraction
# ---------------------------------------------------------------------------

# The "f/r" naming reflects KF8, but the pairs below cover every format
# documented by the bundled KRDS parser.
VARIABLE_SIDECAR_EXTENSIONS = (
    ".azw3f",
    ".yjf",
    ".mbs",
    ".pdt",
    ".tas",
    ".mbp",
)
SUPPORT_SIDECAR_EXTENSIONS = (
    ".azw3r",
    ".yjr",
    ".mbp1",
    ".pds",
    ".tal",
)

EBOOK_EXTENSIONS = (
    ".kfx",
    ".kfx-zip",
    ".azw3",
    ".azw",
    ".azw4",
    ".mobi",
    ".epub",
    ".prc",
    ".pdf",
    ".cbz",
    ".cbr",
    ".zip",
)

# These are generated Kindle documents rather than books in the user's
# library. The check below is intentionally paired with "no real ebook file
# exists", so a real dictionary/periodical ebook is never hidden merely
# because its title contains one of these words.
SYSTEM_DOCUMENT_NAME_PATTERNS = (
    r"\bmy\s+clippings\b",
    r"\bvocabulary\s+builder\b",
    r"\bdictionary\b",
    r"\bperiodicals?\b",
    r"\bnewspapers?\b",
    r"\bmagazines?\b",
    r"我的剪贴",
    r"生词本",
    r"词典",
    r"字典",
    r"期刊",
    r"杂志",
    r"报纸",
)
SYSTEM_DOCUMENT_NAME_RE = re.compile(
    "|".join(SYSTEM_DOCUMENT_NAME_PATTERNS),
    flags=re.IGNORECASE,
)

CALIBRE_METADATA_FILENAME = "metadata.calibre"
WORD_COUNT_CACHE_VERSION = 2
MINIMUM_FREE_SPACE_BYTES = 128 * 1024 * 1024
CONVERSION_ATTEMPTS = 2
CONVERSION_RETRY_DELAY_SECONDS = 0.5
CONVERSION_TIMEOUT_SECONDS = 900
DEFAULT_BOOK_MODES_FILENAME = "book_modes.json"
READING_MODES = ("book", "manga")
SUGGESTED_MANGA_TAGS = {
    "comic",
    "comics",
    "manga",
    "graphic",
    "graphic novel",
    "graphic novels",
}


def _make_logger(
    verbose: bool, name: str = "kindle_reader"
) -> logging.Logger:
    logger = logging.getLogger(name)
    logger.handlers[:] = [logging.NullHandler()]
    logger.propagate = False
    logger.setLevel(logging.DEBUG if verbose else logging.CRITICAL)
    if verbose:
        handler = logging.StreamHandler(sys.stderr)
        handler.setFormatter(logging.Formatter("%(message)s"))
        logger.handlers[:] = [handler]
    return logger


def _filesystem_path(path: str) -> str:
    """Enable long Windows paths without changing the user-visible path."""

    if os.name != "nt":
        return path

    absolute = os.path.abspath(path)
    if absolute.startswith("\\\\?\\"):
        return absolute
    if absolute.startswith("\\\\"):
        return "\\\\?\\UNC\\" + absolute[2:]
    return "\\\\?\\" + absolute


def _display_path(path: Any) -> str:
    """Return a user-facing path without the Windows extended prefix."""

    text = str(path or "")
    if text.startswith("\\\\?\\UNC\\"):
        return "\\\\" + text[8:]
    if text.startswith("\\\\?\\"):
        return text[4:]
    return text


def _remove_file_quietly(path: str) -> None:
    try:
        if os.path.isfile(_filesystem_path(path)):
            os.unlink(_filesystem_path(path))
    except OSError:
        pass


def _atomic_write_text(
    path: str,
    document: str,
    encoding: str = "utf-8",
) -> None:
    output_path = os.path.abspath(path)
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(_filesystem_path(output_dir), exist_ok=True)

    temporary_path: Optional[str] = None
    try:
        with tempfile.NamedTemporaryFile(
            "w",
            encoding=encoding,
            newline="",
            dir=output_dir or None,
            prefix=os.path.basename(output_path) + ".",
            suffix=".tmp",
            delete=False,
        ) as handle:
            temporary_path = handle.name
            handle.write(document)
        os.replace(
            _filesystem_path(temporary_path),
            _filesystem_path(output_path),
        )
    finally:
        if temporary_path and os.path.exists(
            _filesystem_path(temporary_path)
        ):
            os.unlink(_filesystem_path(temporary_path))


def _backup_corrupt_file(path: str) -> Optional[str]:
    if not os.path.isfile(_filesystem_path(path)):
        return None

    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = f"{path}.corrupt-{stamp}"
    counter = 1
    while os.path.exists(_filesystem_path(backup_path)):
        backup_path = f"{path}.corrupt-{stamp}-{counter}"
        counter += 1
    try:
        shutil.copy2(path, backup_path)
    except OSError:
        return None
    return backup_path


def _format_bytes(value: int) -> str:
    amount = float(value)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if abs(amount) < 1024.0 or unit == "TB":
            return f"{amount:.1f} {unit}"
        amount /= 1024.0
    return f"{value} B"


def _has_free_space(
    path: str,
    required_bytes: int = MINIMUM_FREE_SPACE_BYTES,
) -> Tuple[bool, Optional[int]]:
    probe = os.path.abspath(path)
    while not os.path.exists(_filesystem_path(probe)):
        parent = os.path.dirname(probe)
        if parent == probe:
            break
        probe = parent
    try:
        free_bytes = shutil.disk_usage(probe).free
    except OSError:
        return True, None
    return free_bytes >= required_bytes, free_bytes


def decode_sidecar(path: str, logger: logging.Logger) -> Dict[str, Any]:
    """Decode one sidecar, returning an error record instead of raising."""

    try:
        with open(_filesystem_path(path), "rb") as handle:
            data = handle.read()

        # The original parser logs unknown fields to stdout. Suppress that
        # behavior so a large library does not flood the terminal.
        with contextlib.redirect_stdout(io.StringIO()):
            return KindleReaderDataStore(logger, data).deserialize()
    except Exception as exc:
        return {"_error": "%s: %s" % (type(exc).__name__, exc)}


def _stats_from_timer(
    timer_model: Optional[Dict[str, Any]],
    book_info: Optional[Dict[str, Any]],
) -> Dict[str, Any]:
    """Preserve the original plugin's numeric statistics formulas."""

    out: Dict[str, Any] = {}
    if not isinstance(timer_model, dict):
        return out

    total_time = timer_model.get("totalTime") or 0
    total_words = timer_model.get("totalWords") or 0
    total_percent = timer_model.get("totalPercent") or 0.0

    minutes = total_time / 60000.0
    out["total_time_ms"] = total_time
    out["total_words"] = total_words
    out["minutes"] = round(minutes, 1)
    out["hours"] = round(minutes / 60.0, 2)
    out["words_read"] = total_words
    out["percent_read"] = round(total_percent * 100, 1)
    out["wpm_overall"] = round(total_words / minutes) if minutes else 0

    average_calculator = timer_model.get("averageCalculator") or {}
    distributions = average_calculator.get("normalDistributions") or []
    total_sum = sum(
        item.get("sum", 0)
        for item in distributions
        if isinstance(item, dict)
    )
    total_count = sum(
        item.get("count", 0)
        for item in distributions
        if isinstance(item, dict)
    )
    out["wpm_reading"] = (
        round(total_sum / total_count)
        if total_count
        else out["wpm_overall"]
    )

    if isinstance(book_info, dict):
        known_words = book_info.get("numberOfWords", 0)
        coverage = book_info.get("percentOfBook", 0)
        try:
            known_words = int(known_words)
        except (TypeError, ValueError):
            known_words = 0
        try:
            coverage = float(coverage)
        except (TypeError, ValueError):
            coverage = 0.0

        # numberOfWords only covers percentOfBook of the title. The plugin
        # treated numberOfWords itself as the total, which makes total words
        # and word-based progress far too small.
        if coverage > 1.0:
            coverage /= 100.0
        if known_words > 0 and 0.0 < coverage <= 1.0:
            out["book_words_known"] = known_words
            out["book_word_coverage_percent"] = round(
                coverage * 100.0, 6
            )
            out["book_words"] = int(round(known_words / coverage))
        elif known_words > 0:
            out["book_words_known"] = known_words
            out["book_words"] = known_words

    return out


def _nested_timer_model(decoded: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Read top-level or nested readingTimerModel structures."""

    timer_model = decoded.get("timer.model")
    if isinstance(timer_model, dict):
        return timer_model

    for store_name in ("timer.data.store.v2", "timer.data.store"):
        store = decoded.get(store_name)
        if not isinstance(store, dict):
            continue
        timer_model = store.get("readingTimerModel")
        if isinstance(timer_model, dict):
            if isinstance(timer_model.get("timer.model"), dict):
                return timer_model["timer.model"]
            return timer_model

    return None


def _apnx(support_sidecar: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    value = (support_sidecar or {}).get("apnx.key")
    return value if isinstance(value, dict) else None


def _total_pages(support_sidecar: Dict[str, Any]) -> Optional[int]:
    apnx = _apnx(support_sidecar)
    if apnx and isinstance(apnx.get("oPNToPosition"), list):
        return len(apnx["oPNToPosition"]) - 1
    return None


def _page_of(
    support_sidecar: Dict[str, Any], position: Optional[int]
) -> Optional[int]:
    apnx = _apnx(support_sidecar)
    if (
        not apnx
        or not isinstance(apnx.get("oPNToPosition"), list)
        or position is None
    ):
        return None

    positions = apnx["oPNToPosition"]
    low, high, answer = 0, len(positions) - 1, 0
    while low <= high:
        middle = (low + high) // 2
        if positions[middle] <= position:
            answer = middle
            low = middle + 1
        else:
            high = middle - 1
    return answer


def _furthest_position(decoded: Dict[str, Any]) -> Optional[int]:
    for key in ("fpr", "lpr"):
        value = decoded.get(key)
        if isinstance(value, dict) and value.get("position") is not None:
            try:
                return int(value["position"])
            except (TypeError, ValueError):
                continue
    return None


def _asin_from_info(decoded: Dict[str, Any]) -> Optional[str]:
    book_info = decoded.get("book.info.store")
    if isinstance(book_info, dict):
        for key in ("asin", "ASIN", "cdeKey", "cde_key"):
            if book_info.get(key):
                return str(book_info[key])
    return None


def _annotations_count(support_sidecar: Dict[str, Any]) -> Optional[int]:
    annotations = (support_sidecar or {}).get("annotation.cache.object")
    if not isinstance(annotations, dict):
        return None

    count = 0
    for items in annotations.values():
        if isinstance(items, list):
            count += len(items)
    return count or None


def _to_datetime(value: Any) -> Optional[datetime.datetime]:
    """Accept epoch seconds, epoch milliseconds, or an ISO timestamp."""

    if not value:
        return None

    if isinstance(value, (int, float)):
        timestamp = float(value)
        if timestamp > 1e12:
            timestamp /= 1000.0
        if timestamp <= 0:
            return None
        try:
            return datetime.datetime.fromtimestamp(
                timestamp, datetime.timezone.utc
            )
        except (OverflowError, OSError, ValueError):
            return None

    try:
        text = str(value).strip()
        try:
            timestamp = float(text)
        except ValueError:
            parsed = datetime.datetime.fromisoformat(
                text.replace("Z", "+00:00")
            )
            return parsed
        return _to_datetime(timestamp)
    except (TypeError, ValueError):
        return None


def _last_read(decoded: Dict[str, Any]) -> Optional[str]:
    for key in ("lpr", "fpr"):
        value = decoded.get(key)
        if isinstance(value, dict) and value.get("time"):
            parsed = _to_datetime(value["time"])
            if parsed:
                return parsed.isoformat()
    return None


def _book_from_sidecars(
    variable_sidecar: Dict[str, Any],
    support_sidecar: Dict[str, Any],
    title: str,
    author: str,
) -> Dict[str, Any]:
    timer_model = _nested_timer_model(variable_sidecar)
    stats = _stats_from_timer(
        timer_model, variable_sidecar.get("book.info.store")
    )
    position = _furthest_position(variable_sidecar)
    page = _page_of(support_sidecar, position)
    total_pages = _total_pages(support_sidecar)
    page_string = (
        "%d/%d" % (page, total_pages)
        if page is not None and total_pages
        else (str(page) if page is not None else None)
    )
    progress_percent = stats.get("percent_read")
    if (
        (stats.get("total_time_ms") or 0) <= 0
        and (progress_percent or 0.0) <= 0.0
        and (stats.get("total_words") or 0) <= 0
    ):
        # A zero timer on an unopened fixed-layout book is "unknown", not a
        # meaningful 0% reading position.
        progress_percent = None
    support_valid = bool(support_sidecar) and not support_sidecar.get(
        "_error"
    )

    return {
        "title": title,
        "author": author,
        "asin": _asin_from_info(variable_sidecar),
        "series": None,
        "series_index": None,
        "publisher": None,
        "published_date": None,
        "metadata_match": None,
        "metadata_tags": "",
        "metadata_suggested_mode": None,
        "total_time_ms": stats.get("total_time_ms"),
        "minutes": (
            int(round(stats["minutes"])) if stats.get("minutes") else None
        ),
        "hours": stats.get("hours"),
        "total_words": stats.get("total_words"),
        "words_read": stats.get("words_read"),
        "book_words": stats.get("book_words") or None,
        "book_words_known": stats.get("book_words_known"),
        "book_word_coverage_percent": stats.get(
            "book_word_coverage_percent"
        ),
        "book_words_exact": None,
        "word_count_source": "kindle-estimate",
        "word_count_language": None,
        "ebook_file": None,
        "word_count_error": None,
        "wpm": stats.get("wpm_reading"),
        "wpm_overall": stats.get("wpm_overall"),
        "wpm_reading": stats.get("wpm_reading"),
        "progress_percent": progress_percent,
        "progress": (
            int(round(progress_percent)) if progress_percent else None
        ),
        "page": page_string,
        "annotations": _annotations_count(support_sidecar),
        "last_read": _last_read(variable_sidecar),
        "has_data": bool(timer_model or support_valid),
        "fmcache_fallback": None,
    }


def _scan_for_kindle_documents() -> Optional[str]:
    """Best-effort USB Kindle documents folder discovery."""

    candidates: List[str] = []

    if os.name == "nt":
        try:
            import ctypes
            import string

            kernel32 = ctypes.windll.kernel32
            mask = kernel32.GetLogicalDrives()
            for index, letter in enumerate(string.ascii_uppercase):
                if not (mask >> index) & 1:
                    continue
                root = letter + ":\\"
                if kernel32.GetDriveTypeW(root) not in (2, 3):
                    continue
                candidates.append(root)
        except Exception:
            candidates = []
    else:
        for base in ("/media", "/run/media", "/mnt", "/Volumes"):
            if not os.path.isdir(base):
                continue
            try:
                names = os.listdir(base)
            except OSError:
                continue
            for name in names:
                path = os.path.join(base, name)
                candidates.append(path)
                if os.path.isdir(path):
                    try:
                        for child in os.listdir(path):
                            candidates.append(os.path.join(path, child))
                    except OSError:
                        pass

    best = ""
    for root in candidates:
        try:
            documents = os.path.join(root, "documents")
            if not os.path.isdir(_filesystem_path(documents)):
                continue
            if (
                os.path.isfile(_filesystem_path(os.path.join(root, "cc.db")))
                or os.path.isdir(_filesystem_path(os.path.join(root, "system")))
            ):
                return documents
            best = best or documents
        except OSError:
            continue
    return best or None


def resolve_documents(path: Optional[str]) -> Optional[str]:
    """Accept a Kindle root, documents folder, or auto-detect if omitted."""

    if path:
        root = os.path.abspath(os.path.expanduser(path))
        if not os.path.isdir(_filesystem_path(root)):
            return None

        if os.path.basename(os.path.normpath(root)).lower() == "documents":
            return root

        documents = os.path.join(root, "documents")
        if os.path.isdir(_filesystem_path(documents)):
            return documents

        # Also allow a user to point directly at a sidecar collection.
        return root

    return _scan_for_kindle_documents()


def _text_value(value: Any) -> Optional[str]:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def normalize_reading_mode(value: Any) -> str:
    """Normalize a configured reading mode without silently classifying books."""

    normalized = str(value or "").strip().casefold()
    return normalized if normalized in READING_MODES else "book"


def _reading_mode_title_key(value: Any) -> str:
    # Kindle titles and sidecar folder names can carry a trailing storage
    # identifier (for example "合成漫画测试卷 Vol.05_A000..."); strip it so a
    # mapping for the visible title still matches.
    return _normalize_match_text(_strip_metadata_identifier(value, None))


_READING_MODE_VOLUME_SUFFIX = re.compile(
    r"[\s:：\-–—_]*("
    r"第\s*[0-9０-９一二三四五六七八九十百零两]+\s*[卷册集话話部]"
    r"|(?:vol(?:ume)?|book|part|no)\.?\s*[0-9０-９]+"
    r"|[0-9０-９]+\s*[卷册集话話部]"
    r"|#\s*[0-9０-９]+"
    r")\s*$",
    flags=re.IGNORECASE,
)


def _reading_mode_series_key(value: Any) -> str:
    """Return a volume-agnostic key so a mapped volume covers its series."""

    text = unicodedata.normalize(
        "NFKC", _strip_metadata_identifier(value, None)
    ).strip()
    if not text:
        return ""
    previous = None
    while previous != text:
        previous = text
        text = _READING_MODE_VOLUME_SUFFIX.sub("", text).strip()
    return _normalize_match_text(text)


def normalize_book_modes(value: Any) -> Dict[str, Dict[str, str]]:
    """Normalize the ASIN/title/series mapping used by book_modes.json."""

    raw = value if isinstance(value, dict) else {}
    result: Dict[str, Dict[str, str]] = {
        "by_asin": {},
        "by_title": {},
        "by_series": {},
    }

    by_asin = raw.get("by_asin")
    if isinstance(by_asin, dict):
        for asin, mode in by_asin.items():
            normalized_asin = _normalize_asin(asin)
            if normalized_asin:
                result["by_asin"][normalized_asin] = normalize_reading_mode(mode)

    by_title = raw.get("by_title")
    if isinstance(by_title, dict):
        for title, mode in by_title.items():
            normalized_title = _reading_mode_title_key(title)
            normalized_series = _reading_mode_series_key(title)
            normalized_mode = normalize_reading_mode(mode)
            if normalized_title:
                result["by_title"][normalized_title] = normalized_mode
            if normalized_series:
                result["by_series"][normalized_series] = normalized_mode

    for raw_key, mode in raw.items():
        if str(raw_key) in {"version", "by_asin", "by_title"}:
            continue
        normalized_mode = normalize_reading_mode(mode)
        normalized_asin = _normalize_asin(raw_key)
        if normalized_asin and re.fullmatch(
            r"[A-Za-z0-9]{8,}", str(raw_key).strip()
        ):
            result["by_asin"][normalized_asin] = normalized_mode
            continue
        normalized_title = _reading_mode_title_key(raw_key)
        normalized_series = _reading_mode_series_key(raw_key)
        if normalized_title:
            result["by_title"][normalized_title] = normalized_mode
        if normalized_series:
            result["by_series"][normalized_series] = normalized_mode

    return result


def load_book_modes(
    path: Optional[str],
    logger: Optional[logging.Logger] = None,
) -> Tuple[Dict[str, Dict[str, str]], Optional[str]]:
    """Load a local mode mapping; a missing default file is valid and empty."""

    if not path:
        return normalize_book_modes({}), None

    mapping_path = Path(path).expanduser()
    try:
        if not mapping_path.is_file():
            message = "book mode mapping was not found: %s" % mapping_path
            if logger is not None:
                logger.warning(message)
            return normalize_book_modes({}), message
        with mapping_path.open("r", encoding="utf-8-sig") as handle:
            decoded = json.load(handle)
    except (OSError, ValueError) as exc:
        message = "book mode mapping could not be read (%s): %s" % (
            mapping_path,
            exc,
        )
        if logger is not None:
            logger.warning(message)
        return normalize_book_modes({}), message

    if not isinstance(decoded, dict):
        message = "book mode mapping must be a JSON object: %s" % mapping_path
        if logger is not None:
            logger.warning(message)
        return normalize_book_modes({}), message
    return normalize_book_modes(decoded), None


def reading_mode_for_book(
    row: Optional[Dict[str, Any]],
    mapping: Optional[Dict[str, Any]] = None,
) -> str:
    """Return the configured mode, preferring an explicit row value.

    The row's `reading_mode` is authoritative when present. Otherwise the
    local mapping is checked by ASIN first, by stable normalized title
    second, and by the volume-agnostic series title last so mapping one
    volume covers its siblings. This function never infers a mode from
    Calibre tags or suffixes.
    """

    book = row if isinstance(row, dict) else {}
    explicit = str(book.get("reading_mode") or "").strip().casefold()
    if explicit in READING_MODES:
        return explicit
    configured = mapped_reading_mode_for_book(book, mapping)
    return configured or "book"


def mapped_reading_mode_for_book(
    row: Optional[Dict[str, Any]],
    mapping: Optional[Dict[str, Any]] = None,
) -> Optional[str]:
    """Return only a mode explicitly present in book_modes.json."""

    book = row if isinstance(row, dict) else {}
    normalized_mapping = normalize_book_modes(mapping or {})
    asin = _normalize_asin(book.get("asin"))
    if asin:
        configured = normalized_mapping["by_asin"].get(asin)
        if configured in READING_MODES:
            return configured

    title_key = _reading_mode_title_key(book.get("title"))
    if title_key:
        configured = normalized_mapping["by_title"].get(title_key)
        if configured in READING_MODES:
            return configured

    series_key = _reading_mode_series_key(book.get("title"))
    if series_key:
        configured = normalized_mapping["by_series"].get(series_key)
        if configured in READING_MODES:
            return configured
    return None


def suggested_reading_mode_from_metadata(
    entry: Optional[Dict[str, Any]],
    ebook_path: Optional[str] = None,
) -> Optional[str]:
    """Return a non-binding manga suggestion from tags or a comic suffix."""

    suffix = os.path.splitext(str(ebook_path or ""))[1].casefold()
    if suffix in {".cbz", ".cbr"}:
        return "manga"

    if not isinstance(entry, dict):
        return None
    values: List[str] = []
    for key in ("tags", "categories", "subjects"):
        raw = entry.get(key)
        if isinstance(raw, list):
            values.extend(str(item) for item in raw)
        elif raw is not None:
            values.append(str(raw))
    for value in values:
        normalized = re.sub(r"[\s_\-]+", " ", value.strip().casefold())
        if normalized in SUGGESTED_MANGA_TAGS:
            return "manga"
    return None


def _metadata_author(entry: Dict[str, Any]) -> Optional[str]:
    authors = entry.get("authors")
    if isinstance(authors, list):
        names = [
            name
            for name in (_text_value(author) for author in authors)
            if name
        ]
        return ", ".join(names) or None
    return _text_value(authors)


def _metadata_asin(entry: Dict[str, Any]) -> Optional[str]:
    identifiers = entry.get("identifiers")
    if not isinstance(identifiers, dict):
        return None

    for key, value in identifiers.items():
        if str(key).casefold() in {"mobi-asin", "asin", "amazon"}:
            if isinstance(value, list):
                for item in value:
                    asin = _text_value(item)
                    if asin:
                        return asin
            else:
                asin = _text_value(value)
                if asin:
                    return asin

    for key, value in identifiers.items():
        if "asin" not in str(key).casefold():
            continue
        if isinstance(value, list):
            for item in value:
                asin = _text_value(item)
                if asin:
                    return asin
        else:
            asin = _text_value(value)
            if asin:
                return asin
    return None


def _normalize_match_text(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).casefold()
    return "".join(char for char in text if char.isalnum())


def _normalize_match_path(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).replace("\\", "/")
    text = re.sub(r"/+", "/", text).strip()
    while text.startswith("./"):
        text = text[2:]
    return text.casefold()


def _path_without_ebook_extension(value: Any) -> str:
    path = _normalize_match_path(value)
    if not path:
        return ""
    head, tail = path.rsplit("/", 1) if "/" in path else ("", path)
    if os.path.splitext(tail)[1] not in EBOOK_EXTENSIONS:
        return path
    stem = os.path.splitext(tail)[0]
    return f"{head}/{stem}" if head else stem


def _normalize_asin(value: Any) -> Optional[str]:
    text = unicodedata.normalize("NFKC", str(value or ""))
    compact = re.sub(r"\s+", "", text)
    return compact.upper() or None


def _strip_metadata_identifier(value: Any, asin: Optional[str]) -> str:
    text = str(value or "")
    if asin:
        text = re.sub(re.escape(asin), "", text, flags=re.IGNORECASE)
    return re.sub(
        r"(?:[_\-\s]*[0-9a-f]{24,})+$",
        "",
        text,
        flags=re.IGNORECASE,
    )


def _basename_stem(path: Any) -> str:
    text = str(path or "").replace("\\", "/")
    return os.path.splitext(text.rsplit("/", 1)[-1])[0]


def _document_relative_path(path: str, documents: Optional[str]) -> Optional[str]:
    if not path or not documents:
        return None

    try:
        relative = os.path.relpath(path, documents)
    except ValueError:
        return None
    if relative == os.pardir or relative.startswith(os.pardir + os.sep):
        return None

    relative = relative.replace("\\", "/")
    return "documents" if relative in {"", "."} else f"documents/{relative}"


def _group_path_stems(
    group: Dict[str, Any], documents: Optional[str]
) -> set[str]:
    group_directory = _text_value(group.get("directory"))
    group_stem = _text_value(group.get("stem")) or ""
    source_path = _text_value(group.get("source_file"))

    sidecar_directory = group_directory
    if not sidecar_directory and source_path:
        sidecar_directory = os.path.dirname(source_path)

    ebook_directory = sidecar_directory
    if (
        sidecar_directory
        and os.path.basename(os.path.normpath(sidecar_directory))
        .casefold()
        .endswith(".sdr")
    ):
        ebook_directory = os.path.dirname(os.path.normpath(sidecar_directory))
        group_stem = (
            os.path.basename(os.path.normpath(sidecar_directory))[:-4]
            or group_stem
        )

    candidates: List[str] = []
    if ebook_directory and group_stem:
        candidates.append(os.path.join(ebook_directory, group_stem))

    actual_ebook = _find_ebook_file(group)
    if actual_ebook:
        candidates.append(actual_ebook)

    stems: set[str] = set()
    for candidate in candidates:
        for value in (candidate, _document_relative_path(candidate, documents)):
            normalized = _path_without_ebook_extension(value)
            if normalized:
                stems.add(normalized)
    return stems


def _metadata_match_score(
    group: Dict[str, Any],
    entry: Dict[str, Any],
    book: Optional[Dict[str, Any]] = None,
    documents: Optional[str] = None,
) -> Tuple[int, Optional[str]]:
    book = book or {}
    group_stem = _text_value(group.get("stem")) or ""
    source_stem = _basename_stem(group.get("source_file"))
    fallback_title = _text_value(group.get("title")) or ""
    asin = _metadata_asin(entry)
    normalized_asin = _normalize_asin(asin)
    candidates = (
        group_stem,
        source_stem,
        fallback_title,
        _text_value(group.get("directory")),
    )
    if normalized_asin:
        book_asin = _normalize_asin(book.get("asin"))
        if book_asin and book_asin == normalized_asin:
            return 3, "asin"
        if any(
            normalized_asin.casefold() in candidate.casefold()
            for candidate in candidates
            if candidate
        ):
            return 3, "asin"

    metadata_path = _path_without_ebook_extension(
        _text_value(entry.get("lpath"))
    )
    if metadata_path and metadata_path in _group_path_stems(group, documents):
        return 2, "path"

    metadata_title = _text_value(entry.get("title"))
    normalized_title = _normalize_match_text(metadata_title)
    if not normalized_title:
        return 0, None

    metadata_asin = _text_value(asin)
    for candidate in (group_stem, source_stem, fallback_title):
        if not candidate:
            continue
        stripped = _strip_metadata_identifier(candidate, metadata_asin)
        if _normalize_match_text(stripped) == normalized_title:
            return 1, "title"
    return 0, None


def _match_calibre_metadata(
    group: Dict[str, Any],
    metadata: Sequence[Dict[str, Any]],
    book: Optional[Dict[str, Any]] = None,
    documents: Optional[str] = None,
) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    best_entry: Optional[Dict[str, Any]] = None
    best_score = 0
    best_match: Optional[str] = None
    for entry in metadata:
        score, match = _metadata_match_score(
            group,
            entry,
            book=book,
            documents=documents,
        )
        if score > best_score:
            best_entry = entry
            best_score = score
            best_match = match
    return best_entry, best_match


def _find_calibre_metadata(
    documents: str, requested: Optional[str] = None
) -> Optional[str]:
    if requested:
        path = os.path.abspath(os.path.expanduser(requested))
        return path if os.path.isfile(_filesystem_path(path)) else None

    documents_path = os.path.abspath(os.path.expanduser(documents))
    candidates = [
        os.path.join(
            os.path.dirname(os.path.normpath(documents_path)),
            CALIBRE_METADATA_FILENAME,
        ),
        os.path.join(
            os.path.normpath(documents_path),
            CALIBRE_METADATA_FILENAME,
        ),
    ]
    for path in candidates:
        if os.path.isfile(_filesystem_path(path)):
            return path
    return None


def _load_calibre_metadata(
    path: Optional[str], logger: logging.Logger
) -> Tuple[List[Dict[str, Any]], List[str]]:
    if not path:
        return [], []

    try:
        with open(
            _filesystem_path(path), "r", encoding="utf-8-sig"
        ) as handle:
            decoded = json.load(handle)
    except (OSError, ValueError) as exc:
        message = "Calibre metadata failed (%s): %s" % (path, exc)
        logger.warning(message)
        return [], [message]

    if not isinstance(decoded, list):
        message = "Calibre metadata is not a list: %s" % path
        logger.warning(message)
        return [], [message]

    records = [entry for entry in decoded if isinstance(entry, dict)]
    if len(records) != len(decoded):
        logger.warning(
            "Ignored %d non-object Calibre metadata record(s)",
            len(decoded) - len(records),
        )
    return records, []


def _apply_calibre_metadata(
    book: Dict[str, Any],
    entry: Optional[Dict[str, Any]],
    match: Optional[str] = None,
) -> None:
    if not entry:
        return

    title = _text_value(entry.get("title"))
    author = _metadata_author(entry)
    asin = _metadata_asin(entry)
    series = _text_value(entry.get("series"))
    series_index = _text_value(entry.get("series_index"))
    publisher = _text_value(entry.get("publisher"))
    published_date = next(
        (
            _text_value(entry.get(key))
            for key in ("pubdate", "published", "publication_date")
            if _text_value(entry.get(key))
        ),
        None,
    )
    if title:
        book["title"] = title
    if author:
        book["author"] = author
    if asin:
        book["asin"] = asin
    if series:
        book["series"] = series
    if series_index:
        book["series_index"] = series_index
    if publisher:
        book["publisher"] = publisher
    if published_date:
        book["published_date"] = published_date
    if match:
        book["metadata_match"] = match
    tags = entry.get("tags")
    if isinstance(tags, list):
        book["metadata_tags"] = "; ".join(
            text for text in (_text_value(tag) for tag in tags) if text
        )
    elif tags is not None:
        book["metadata_tags"] = _text_value(tags) or ""


def _first_matching(
    filenames: Iterable[str], extensions: Sequence[str]
) -> Optional[str]:
    by_suffix: Dict[str, List[str]] = {}
    for filename in filenames:
        suffix = os.path.splitext(filename)[1].lower()
        if suffix in extensions:
            by_suffix.setdefault(suffix, []).append(filename)

    for suffix in extensions:
        matches = sorted(by_suffix.get(suffix, []), key=str.lower)
        if matches:
            return matches[0]
    return None


def find_sidecar_groups(
    documents: str,
    errors: Optional[List[str]] = None,
) -> List[Dict[str, Any]]:
    """Walk documents and pair variable/support sidecars per directory."""

    groups: List[Dict[str, Any]] = []

    def on_walk_error(exc: OSError) -> None:
        if errors is None:
            return
        path = _display_path(getattr(exc, "filename", "") or documents)
        errors.append(
            "directory scan failed (%s): %s: %s; skipped the affected "
            "subtree and continued. Reconnect the Kindle, check the path "
            "permissions, then rerun." % (path, type(exc).__name__, exc)
        )

    for extended_dirpath, _dirnames, filenames in os.walk(
        _filesystem_path(documents),
        onerror=on_walk_error,
    ):
        dirpath = _display_path(extended_dirpath)
        variable_name = _first_matching(
            filenames, VARIABLE_SIDECAR_EXTENSIONS
        )
        support_name = _first_matching(
            filenames, SUPPORT_SIDECAR_EXTENSIONS
        )
        if not variable_name and not support_name:
            continue

        folder_name = os.path.basename(os.path.normpath(dirpath))
        stem = folder_name
        if folder_name.lower().endswith(".sdr"):
            stem = folder_name[:-4]

        title, author = stem, ""
        if " - " in stem:
            title, author = stem.rsplit(" - ", 1)

        variable_path = (
            os.path.join(dirpath, variable_name) if variable_name else None
        )
        support_path = (
            os.path.join(dirpath, support_name) if support_name else None
        )
        source_path = variable_path or support_path
        groups.append(
            {
                "directory": dirpath,
                "stem": stem,
                "title": title,
                "author": author,
                "variable_path": variable_path,
                "support_path": support_path,
                "source_file": source_path,
            }
        )

    groups.sort(key=lambda item: item["source_file"].lower())
    return groups


def _find_calibre_convert() -> Optional[str]:
    configured = os.environ.get("CALIBRE_EBOOK_CONVERT", "").strip()
    if configured and os.path.isfile(configured):
        return configured

    candidates = [
        r"C:\Program Files\Calibre2\ebook-convert.exe",
        r"C:\Program Files (x86)\Calibre2\ebook-convert.exe",
        shutil.which("ebook-convert.exe"),
        shutil.which("ebook-convert"),
    ]
    for candidate in candidates:
        if candidate and os.path.isfile(candidate):
            return candidate
    return None


def _find_ebook_file(group: Dict[str, Any]) -> Optional[str]:
    source_path = group.get("source_file")
    if not source_path:
        return None

    directory = os.path.dirname(source_path)
    if os.path.basename(directory).lower().endswith(".sdr"):
        directory = os.path.dirname(directory)

    wanted_raw = str(
        group.get("title") or group.get("stem") or ""
    ).strip()
    wanted_plain = wanted_raw.casefold()
    wanted = _normalize_match_text(
        _strip_metadata_identifier(wanted_raw, None)
    )
    exact_matches: List[str] = []
    prefix_matches: List[str] = []
    normalized_matches: List[str] = []

    try:
        names = os.listdir(_filesystem_path(directory))
    except (OSError, ValueError):
        return None

    for name in names:
        path = os.path.join(directory, name)
        extension = os.path.splitext(name)[1].lower()
        if extension not in EBOOK_EXTENSIONS or not os.path.isfile(
            _filesystem_path(path)
        ):
            continue
        stem = os.path.splitext(name)[0]
        folded = stem.casefold()
        if folded == wanted_plain:
            exact_matches.append(path)
        elif wanted_plain and folded.startswith(wanted_plain):
            prefix_matches.append(path)
        normalized = _normalize_match_text(
            _strip_metadata_identifier(stem, None)
        )
        if wanted and normalized == wanted:
            normalized_matches.append(path)

    matches = sorted(
        exact_matches or normalized_matches or prefix_matches,
        key=len,
    )
    return matches[0] if matches else None


def _system_document_reason(
    group: Dict[str, Any],
    ebook_path: Optional[str],
) -> Optional[str]:
    """Identify generated Kindle documents without hiding real ebooks."""

    if ebook_path:
        return None
    labels = [
        str(group.get("stem") or ""),
        str(group.get("title") or ""),
        os.path.basename(
            os.path.dirname(str(group.get("source_file") or ""))
        ),
    ]
    for label in labels:
        normalized = unicodedata.normalize("NFKC", label).strip()
        if normalized and SYSTEM_DOCUMENT_NAME_RE.search(normalized):
            return "system-document"
    return None


def _word_count_cache_key(ebook_path: str) -> str:
    stat = os.stat(ebook_path)
    raw = "{}|{}|{}|{}".format(
        os.path.abspath(ebook_path),
        stat.st_size,
        stat.st_mtime_ns,
        WORD_COUNT_CACHE_VERSION,
    )
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def _load_word_count_cache(path: str) -> Dict[str, Any]:
    try:
        with open(_filesystem_path(path), "r", encoding="utf-8") as handle:
            data = json.load(handle)
        if isinstance(data, dict):
            return data
        _backup_corrupt_file(path)
        return {}
    except (OSError, ValueError):
        _backup_corrupt_file(path)
        return {}


def _save_word_count_cache(path: str, cache: Dict[str, Any]) -> None:
    _atomic_write_text(
        path,
        json.dumps(
            cache,
            ensure_ascii=False,
            indent=2,
            sort_keys=True,
        )
        + "\n",
    )


def _convert_ebook_to_text(
    ebook_path: str,
    calibre_path: str,
    text_cache_dir: str,
    logger: logging.Logger,
) -> str:
    stat = os.stat(_filesystem_path(ebook_path))
    digest = hashlib.sha256(
        "{}|{}|{}".format(
            os.path.abspath(ebook_path), stat.st_size, stat.st_mtime_ns
        ).encode("utf-8")
    ).hexdigest()
    text_path = os.path.join(text_cache_dir, digest + ".txt")
    os.makedirs(_filesystem_path(text_cache_dir), exist_ok=True)

    try:
        if (
            os.path.isfile(_filesystem_path(text_path))
            and os.path.getmtime(_filesystem_path(text_path))
            >= os.path.getmtime(_filesystem_path(ebook_path))
        ):
            with open(
                _filesystem_path(text_path), "r", encoding="utf-8"
            ) as handle:
                return handle.read()
    except OSError:
        pass

    kwargs: Dict[str, Any] = {}
    if os.name == "nt":
        kwargs["creationflags"] = 0x08000000  # CREATE_NO_WINDOW

    logger.info("Converting ebook for exact word count: %s", ebook_path)
    free_ok, free_bytes = _has_free_space(text_cache_dir)
    if not free_ok:
        available = (
            _format_bytes(free_bytes)
            if free_bytes is not None
            else "unknown"
        )
        raise RuntimeError(
            "not enough free disk space for exact word count: "
            f"{available} available"
        )

    command = [
        calibre_path,
        ebook_path,
        text_path,
        "--txt-output-formatting=plain",
    ]
    last_error = ""
    last_exception: Optional[BaseException] = None
    for attempt in range(1, CONVERSION_ATTEMPTS + 1):
        last_error = ""
        last_exception = None
        _remove_file_quietly(text_path)
        try:
            result = subprocess.run(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=CONVERSION_TIMEOUT_SECONDS,
                **kwargs,
            )
        except subprocess.TimeoutExpired as exc:
            last_exception = exc
            last_error = (
                "Calibre conversion timed out after %d seconds"
                % CONVERSION_TIMEOUT_SECONDS
            )
        except OSError as exc:
            last_exception = exc
            last_error = "%s: %s" % (type(exc).__name__, exc)
        else:
            if result.returncode == 0:
                if os.path.isfile(_filesystem_path(text_path)):
                    break
                last_error = (
                    "Calibre exited successfully but did not create the "
                    "expected text file"
                )
            else:
                details = (result.stdout or "").strip()
                last_error = "Calibre exited with code %d%s" % (
                    result.returncode,
                    (": " + details) if details else "",
                )

        if attempt < CONVERSION_ATTEMPTS:
            logger.warning(
                "Calibre conversion attempt %d/%d failed for %s: %s; "
                "retrying",
                attempt,
                CONVERSION_ATTEMPTS,
                _display_path(ebook_path),
                last_error,
            )
            time.sleep(CONVERSION_RETRY_DELAY_SECONDS)

    if last_error:
        _remove_file_quietly(text_path)
        raise RuntimeError(
            "Calibre conversion failed after %d attempt(s) for ebook %s "
            "and output %s: %s. Recovery: verify Calibre and KFX Input, "
            "close applications that may lock the file, free disk space, "
            "then rerun."
            % (
                CONVERSION_ATTEMPTS,
                _display_path(ebook_path),
                _display_path(text_path),
                last_error,
            )
        ) from last_exception

    if not os.path.isfile(_filesystem_path(text_path)):
        raise RuntimeError(
            "Calibre did not create the expected text file (%s). Recovery: "
            "check Calibre and KFX Input, then rerun."
            % _display_path(text_path)
        )

    with open(_filesystem_path(text_path), "r", encoding="utf-8") as handle:
        return handle.read()


def _detect_word_count_language(text: str) -> str:
    sample = text[:2_000_000]
    cjk_count = len(re.findall(r"[\u3400-\u9fff]", sample))
    kana_count = len(re.findall(r"[\u3040-\u30ff]", sample))
    latin_count = len(
        re.findall(r"[A-Za-z0-9]+(?:['’-][A-Za-z0-9]+)*", sample)
    )
    if kana_count >= 20 and kana_count * 4 >= max(1, cjk_count):
        return "ja"
    if cjk_count > latin_count * 2:
        return "zh"
    return "en"


def _is_countable_token(token: str) -> bool:
    return any(char.isalnum() for char in token)


def count_text_words(text: str, language: str = "auto") -> Tuple[int, str]:
    """Count words with the best tokenizer available for the language."""

    selected = language if language != "auto" else _detect_word_count_language(
        text
    )

    if selected == "en":
        tokens = re.findall(
            r"[A-Za-z0-9]+(?:['’-][A-Za-z0-9]+)*", text
        )
        return len(tokens), "en-regex"

    if selected == "zh":
        try:
            import jieba  # type: ignore

            count = sum(
                1
                for token in jieba.cut(text, cut_all=False)
                if _is_countable_token(token)
            )
            return count, "jieba"
        except Exception:
            latin = re.findall(
                r"[A-Za-z0-9]+(?:['’-][A-Za-z0-9]+)*", text
            )
            cjk = re.findall(r"[\u3400-\u9fff]", text)
            return len(latin) + len(cjk), "zh-character-fallback"

    if selected == "ja":
        try:
            from janome.tokenizer import Tokenizer  # type: ignore

            tokenizer = Tokenizer()
            count = 0
            for token in tokenizer.tokenize(text):
                part_of_speech = token.part_of_speech.split(",")[0]
                if part_of_speech in {"記号", "空白"}:
                    continue
                if _is_countable_token(token.surface):
                    count += 1
            return count, "janome"
        except Exception:
            latin = re.findall(
                r"[A-Za-z0-9]+(?:['’-][A-Za-z0-9]+)*", text
            )
            cjk = re.findall(r"[\u3400-\u9fff\u3040-\u30ff]", text)
            return len(latin) + len(cjk), "ja-character-fallback"

    raise ValueError("Unknown word-count language: %s" % selected)


def enrich_exact_word_counts(
    books: List[Dict[str, Any]],
    calibre_path: Optional[str],
    cache_path: str,
    text_cache_dir: str,
    language: str,
    logger: logging.Logger,
) -> List[str]:
    """Add exact word-count fields when Calibre can convert the ebook."""

    if not calibre_path:
        return [
            "Calibre ebook-convert.exe was not found. Recovery: install "
            "Calibre or pass --calibre-path, then rerun the exact word-count "
            "step."
        ]

    cache = _load_word_count_cache(cache_path)
    errors: List[str] = []

    for index, book in enumerate(books, start=1):
        group = book.get("_group")
        title = str(book.get("title") or "<unknown book>")
        source_context = _display_path(
            book.get("source_file")
            or (group.get("source_file") if isinstance(group, dict) else "")
        )
        logger.info(
            "Exact word count %d/%d: %s",
            index,
            len(books),
            title,
        )
        if reading_mode_for_book(book) == "manga":
            book["book_words_exact"] = None
            book["word_count_source"] = "not-applicable"
            book["word_count_error"] = None
            logger.info("  skipped: manga mode does not use word counts")
            continue
        try:
            ebook_path = (
                _find_ebook_file(group) if isinstance(group, dict) else None
            )
            if not ebook_path:
                message = (
                    "%s (%s): ebook file was not found. Recovery: keep the "
                    "ebook and matching .sdr folder together, then rerun."
                    % (title, source_context or "<unknown source>")
                )
                errors.append(message)
                book["word_count_source"] = "ebook-not-found"
                book["word_count_error"] = message
                continue

            book["ebook_file"] = ebook_path
            key = _word_count_cache_key(ebook_path)
            cached = cache.get(key)
            if isinstance(cached, dict) and cached.get("count") is not None:
                book["book_words_exact"] = cached["count"]
                book["word_count_source"] = cached.get(
                    "source", "calibre-cache"
                )
                book["word_count_language"] = cached.get("language")
                logger.info(
                    "  cached: %s words (%s)",
                    "{:,}".format(book["book_words_exact"]),
                    book["word_count_source"],
                )
                continue

            text = _convert_ebook_to_text(
                ebook_path,
                calibre_path,
                text_cache_dir,
                logger,
            )
            count, source = count_text_words(text, language)
            selected_language = (
                language
                if language != "auto"
                else _detect_word_count_language(text)
            )
            cache[key] = {
                "count": count,
                "source": source,
                "language": selected_language,
                "ebook_path": os.path.abspath(ebook_path),
            }
            book["book_words_exact"] = count
            book["word_count_source"] = source
            book["word_count_language"] = selected_language
            logger.info(
                "  counted: %s words (%s, %s)",
                "{:,}".format(count),
                source,
                selected_language,
            )
        except Exception as exc:
            message = (
                "%s (%s): %s: %s. Recovery: exact value remains unavailable "
                "for this book only; verify Calibre and KFX Input, close "
                "file-locking applications, then rerun."
                % (
                    title,
                    source_context or "<unknown source>",
                    type(exc).__name__,
                    exc,
                )
            )
            errors.append(message)
            book["word_count_source"] = "exact-count-failed"
            book["word_count_error"] = message
            logger.warning(
                "Exact word count failed for %s: %s",
                source_context or title,
                exc,
            )

    try:
        _save_word_count_cache(cache_path, cache)
    except OSError as exc:
        errors.append(
            "word-count cache could not be saved (%s): %s: %s. Recovery: "
            "free disk space or fix permissions; exact counts from this run "
            "are still available in the CSV."
            % (
                _display_path(cache_path),
                type(exc).__name__,
                exc,
            )
        )
    return errors


CSV_FIELDS = [
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
    "metadata_tags",
    "metadata_suggested_mode",
    "reading_mode",
    "reading_mode_source",
    "classification_confidence",
    "classification_reason",
    "series_key",
    "manga_pages_total",
    "manga_pages_read",
    "manga_pages_source",
    "manga_progress_source",
    "manga_position",
    "manga_max_position",
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
]


def extract_books(
    documents: str,
    logger: logging.Logger,
    include_empty: bool = False,
    metadata_path: Optional[str] = None,
    book_modes: Optional[Dict[str, Any]] = None,
) -> Tuple[List[Dict[str, Any]], List[str]]:
    books: List[Dict[str, Any]] = []
    errors: List[str] = []
    resolved_metadata_path = _find_calibre_metadata(
        documents, requested=metadata_path
    )
    if metadata_path and not resolved_metadata_path:
        errors.append("Calibre metadata file not found: %s" % metadata_path)
    metadata, metadata_errors = _load_calibre_metadata(
        resolved_metadata_path, logger
    )
    errors.extend(metadata_errors)

    groups = find_sidecar_groups(documents, errors=errors)
    for index, group in enumerate(groups, start=1):
        source_path = _display_path(
            group.get("source_file") or group.get("directory") or ""
        )
        try:
            ebook_path = _find_ebook_file(group)
            group["ebook_file"] = ebook_path
            system_reason = _system_document_reason(group, ebook_path)
            if system_reason:
                logger.info(
                    "Skipped generated Kindle document %s (%s)",
                    group.get("title") or group.get("stem") or "<unknown>",
                    system_reason,
                )
                continue
            variable = (
                decode_sidecar(group["variable_path"], logger)
                if group["variable_path"]
                else {}
            )
            support = (
                decode_sidecar(group["support_path"], logger)
                if group["support_path"]
                else {}
            )

            for label, decoded, path in (
                ("variable", variable, group["variable_path"]),
                ("support", support, group["support_path"]),
            ):
                if decoded.get("_error"):
                    errors.append(
                        "%s sidecar failed (%s): %s; skipped only this "
                        "sidecar and continued. Recovery: reconnect the "
                        "Kindle or verify the listed file, then rerun."
                        % (
                            label,
                            _display_path(path),
                            decoded["_error"],
                        )
                    )

            book = _book_from_sidecars(
                variable,
                support,
                group["title"],
                group["author"],
            )
            metadata_entry, metadata_match = _match_calibre_metadata(
                group,
                metadata,
                book=book,
                documents=documents,
            )
            _apply_calibre_metadata(
                book,
                metadata_entry,
                match=metadata_match,
            )
            book["ebook_file"] = ebook_path
            suggestion = suggested_reading_mode_from_metadata(
                metadata_entry,
                ebook_path,
            )
            book["metadata_suggested_mode"] = suggestion
            configured_mode = mapped_reading_mode_for_book(
                book,
                book_modes,
            )
            if configured_mode:
                book["reading_mode"] = configured_mode
                book["reading_mode_source"] = "book-modes"
            elif suggestion:
                book["reading_mode"] = suggestion
                suffix = os.path.splitext(str(ebook_path or ""))[1].casefold()
                book["reading_mode_source"] = (
                    "comic-suffix"
                    if suffix in {".cbz", ".cbr", ".zip"}
                    else "calibre-tags"
                )
                book["classification_confidence"] = 0.95
            else:
                book["reading_mode"] = "book"
                book["reading_mode_source"] = "unclassified"
            if book["reading_mode"] == "manga":
                book["word_count_source"] = "not-applicable"
                book["word_count_error"] = None
            book["series_key"] = _reading_mode_series_key(
                book.get("title")
            )
            if suggestion and book.get("reading_mode") != suggestion:
                logger.info(
                    "Book mode suggestion for %s: %s (mapping was not "
                    "changed)",
                    book.get("title") or group.get("title") or "<unknown>",
                    suggestion,
                )
            book["source_file"] = group["source_file"]
            book["_group"] = group

            if not book["has_data"] and not include_empty:
                continue
            books.append(book)
        except Exception as exc:
            errors.append(
                "book %d failed (%s): %s: %s; skipped only this book and "
                "continued. Recovery: reconnect the Kindle or check the "
                "listed sidecar path, then rerun."
                % (
                    index,
                    source_path or "<unknown source>",
                    type(exc).__name__,
                    exc,
                )
            )
            logger.warning(
                "Book failed and was skipped: %s: %s",
                source_path or "<unknown source>",
                exc,
            )

    books.sort(
        key=lambda item: (
            -(item.get("total_time_ms") or 0),
            (item.get("title") or "").lower(),
        )
    )
    return books, errors


def write_csv(path: str, books: Sequence[Dict[str, Any]]) -> None:
    """Write Excel-friendly CSV using a UTF-8 BOM."""

    free_ok, free_bytes = _has_free_space(path, required_bytes=16 * 1024 * 1024)
    if not free_ok:
        available = (
            _format_bytes(free_bytes)
            if free_bytes is not None
            else "unknown"
        )
        raise OSError(
            "not enough free disk space for CSV output: "
            f"{available} available"
        )

    with io.StringIO(newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_FIELDS)
        writer.writeheader()
        for book in books:
            writer.writerow({field: book.get(field) for field in CSV_FIELDS})
        document = handle.getvalue()
    _atomic_write_text(path, document, encoding="utf-8-sig")


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Extract Kindle reading time, WPM, progress, and page data from "
            "USB sidecars and export CSV."
        )
    )
    parser.add_argument(
        "--path",
        default=None,
        help=(
            "Kindle mount root or its documents folder. If omitted, common "
            "mount locations are scanned automatically."
        ),
    )
    parser.add_argument(
        "--output",
        default="kindle_reading.csv",
        help="CSV output path. Default: kindle_reading.csv",
    )
    parser.add_argument(
        "--include-empty",
        action="store_true",
        help="Include sidecars that contain no readable timer/annotation data.",
    )
    parser.add_argument(
        "--exact-word-count",
        action="store_true",
        help=(
            "Convert ebook content with Calibre and count complete-book words. "
            "Requires the KFX Input plugin for .kfx files."
        ),
    )
    parser.add_argument(
        "--calibre-path",
        default=None,
        help="Path to Calibre ebook-convert.exe.",
    )
    parser.add_argument(
        "--calibre-metadata",
        "--metadata",
        default=None,
        help=(
            "Path to the Kindle root metadata.calibre file. "
            "Auto-detected beside the documents folder when omitted."
        ),
    )
    parser.add_argument(
        "--word-count-language",
        choices=("auto", "en", "zh", "ja"),
        default="auto",
        help="Tokenizer language for exact word counts. Default: auto.",
    )
    parser.add_argument(
        "--word-count-cache",
        default=str(Path(__file__).resolve().parent / "word_counts_cache.json"),
        help="JSON cache for exact word counts.",
    )
    parser.add_argument(
        "--book-modes",
        default=str(Path(__file__).resolve().parent / DEFAULT_BOOK_MODES_FILENAME),
        help=(
            "JSON mapping for reading modes. Default: book_modes.json beside "
            "the script. Values must be book or manga."
        ),
    )
    parser.add_argument(
        "--wordcount-dir",
        default=str(Path(__file__).resolve().parent / "_wordcount"),
        help="Directory for temporary/plain-text conversions.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print sidecar decode errors and the resolved documents path.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version="%(prog)s 1.5.0",
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)
    logger = _make_logger(args.verbose)
    parser_logger = _make_logger(False, "kindle_reader.parser")

    documents = resolve_documents(args.path)
    if not documents:
        print(
            "error: Kindle documents folder was not found. "
            "Pass it explicitly with --path.",
            file=sys.stderr,
        )
        return 2

    if args.verbose:
        print("Scanning: %s" % documents, file=sys.stderr)
        metadata_path = _find_calibre_metadata(
            documents, requested=args.calibre_metadata
        )
        print(
            "Calibre metadata: %s" % (metadata_path or "not found"),
            file=sys.stderr,
        )

    book_modes, mode_error = load_book_modes(args.book_modes, logger)
    if mode_error and args.verbose:
        print("Book modes: %s" % mode_error, file=sys.stderr)

    books, errors = extract_books(
        documents,
        parser_logger,
        include_empty=args.include_empty,
        metadata_path=args.calibre_metadata,
        book_modes=book_modes,
    )
    if args.exact_word_count:
        calibre_path = args.calibre_path or _find_calibre_convert()
        try:
            exact_errors = enrich_exact_word_counts(
                books=books,
                calibre_path=calibre_path,
                cache_path=args.word_count_cache,
                text_cache_dir=args.wordcount_dir,
                language=args.word_count_language,
                logger=logger,
            )
            errors.extend(exact_errors)
        except Exception as exc:
            errors.append(
                "exact word-count enrichment failed: %s: %s. Recovery: "
                "other books were retained; check Calibre, KFX Input, and "
                "the word-count cache path, then rerun."
                % (type(exc).__name__, exc)
            )

    output = os.path.abspath(os.path.expanduser(args.output))
    output_dir = os.path.dirname(output)
    if output_dir:
        try:
            os.makedirs(_filesystem_path(output_dir), exist_ok=True)
        except OSError as exc:
            print(
                "error: cannot create CSV output directory (%s): %s. "
                "Recovery: check the path and permissions, then rerun."
                % (_display_path(output_dir), exc),
                file=sys.stderr,
            )
            return 2
    try:
        write_csv(output, books)
    except OSError as exc:
        print("error: %s" % exc, file=sys.stderr)
        return 2

    if errors:
        print("warning: %d extraction error(s)" % len(errors), file=sys.stderr)
        if args.verbose:
            for error in errors:
                print("  %s" % error, file=sys.stderr)

    print(
        "Exported %d book(s) to %s" % (len(books), output),
        file=sys.stderr,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
