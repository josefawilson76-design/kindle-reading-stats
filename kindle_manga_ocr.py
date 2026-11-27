#!/usr/bin/env python3
"""Optional offline OCR estimates for locally stored manga ebooks.

This module is deliberately outside the KRDS sidecar scanner. It only reads
ebook files, writes cache files below the project OCR directory, and reports
missing OCR/render dependencies without installing anything or using network
access.
"""

from __future__ import annotations

import csv
import concurrent.futures
import datetime
import hashlib
import importlib.util
import io
import json
import logging
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import tracemalloc
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple


def _application_roots() -> List[Path]:
    roots = [Path(__file__).resolve().parent]
    if getattr(sys, "frozen", False):
        executable_dir = Path(sys.executable).resolve().parent
        roots.insert(0, executable_dir)
        roots.insert(1, executable_dir.parent)
        for parent in list(executable_dir.parents)[:3]:
            if parent not in roots:
                roots.append(parent)
    result: List[Path] = []
    for root in roots:
        if root not in result:
            result.append(root)
    return result


_APPLICATION_ROOTS = _application_roots()
_BUILDING_EXE = "--build-exe" in sys.argv
_DEPENDENCY_ROOT = next(
    (
        root / ".ocr-pip"
        for root in _APPLICATION_ROOTS
        if (root / ".ocr-pip").is_dir()
    ),
    _APPLICATION_ROOTS[0] / ".ocr-pip",
)
if _DEPENDENCY_ROOT.is_dir() and not _BUILDING_EXE:
    if str(_DEPENDENCY_ROOT) not in sys.path:
        sys.path.insert(0, str(_DEPENDENCY_ROOT))
    if os.name == "nt":
        _DLL_DIRECTORY_HANDLES: List[Any] = []
        for relative in (
            "paddle/libs",
            "torch/lib",
            "torchvision",
            "cv2",
        ):
            candidate = _DEPENDENCY_ROOT / relative
            if candidate.is_dir():
                try:
                    _DLL_DIRECTORY_HANDLES.append(
                        os.add_dll_directory(str(candidate))
                    )
                except (AttributeError, OSError):
                    pass

_MODEL_ROOT = next(
    (
        root / "_ocr_models"
        for root in _APPLICATION_ROOTS
        if (root / "_ocr_models").is_dir()
    ),
    _APPLICATION_ROOTS[0] / "_ocr_models",
)
os.environ.setdefault(
    "HF_HOME",
    str(_MODEL_ROOT / "huggingface"),
)
os.environ.setdefault(
    "PADDLE_OCR_BASE_DIR",
    str(_MODEL_ROOT / "paddle_cache"),
)

try:
    import kindle_reader  # type: ignore
except ImportError:  # pragma: no cover - standalone import fallback
    kindle_reader = None


OCR_CACHE_VERSION = 1
OCR_RENDER_VERSION = 1
OCR_SOURCE = "ocr-estimate"
OCR_SAMPLED_SOURCE = "ocr-sampled"
OCR_SERIES_SOURCE = "ocr-series-estimate"
OCR_COVERAGE_TARGET_PERCENT = 98.0
OCR_DEFAULT_DPI = 240
OCR_QUALITY_DPI = 300
OCR_SAMPLE_DPI = 200
OCR_CLASSIFY_SAMPLE_PAGES = 6
OCR_DEFAULT_MIN_CONFIDENCE = 0.70
OCR_DEFAULT_CACHE_DIRNAME = "_ocr"
OCR_CONVERSION_TIMEOUT_SECONDS = 1800
OCR_PDFTOPPM_TIMEOUT_SECONDS = 1800
SOURCE_EBOOK_EXTENSIONS = {
    ".azw",
    ".azw3",
    ".azw4",
    ".cbr",
    ".cbz",
    ".epub",
    ".kfx",
    ".kfx-zip",
    ".mobi",
    ".pdf",
    ".zip",
}

OCR_FIELDS = (
    "ocr_mode",
    "ocr_word_count",
    "ocr_word_count_is_estimate",
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
    "ocr_min_confidence",
    "ocr_sampled_pages",
    "ocr_extrapolated",
    "classification_source",
    "classification_confidence",
    "classification_reason",
    "manga_pages_total",
    "manga_pages_read",
    "manga_pages_source",
)

DEPENDENCY_HINTS = {
    "manga_ocr": (
        "manga_ocr is unavailable. Install it in an isolated environment "
        "(for example: python -m pip install manga-ocr) and prefetch/cache "
        "its model before running offline."
    ),
    "paddleocr": (
        "PaddleOCR is unavailable or has no local model directory. Install "
        "paddleocr and paddlepaddle, prefetch the required language models, "
        "then set KINDLE_OCR_PADDLEOCR_MODEL_DIR to that local directory."
    ),
    "tesseract": (
        "Tesseract is unavailable. Install Tesseract OCR plus the needed "
        "language data, then put tesseract.exe on PATH or set "
        "KINDLE_OCR_TESSERACT to its full path."
    ),
}


class OcrDependencyError(RuntimeError):
    """Raised when an OCR engine or renderer dependency cannot be used."""


class OcrRenderError(RuntimeError):
    """Raised when a manga ebook cannot be rendered into page images."""


@dataclass(frozen=True)
class OcrConfig:
    language: str = "auto"
    engine: str = "auto"
    dpi: int = OCR_DEFAULT_DPI
    min_confidence: float = OCR_DEFAULT_MIN_CONFIDENCE
    cache_dir: Path = Path(OCR_DEFAULT_CACHE_DIRNAME)
    max_pages: int = 0
    mode: str = "full"
    workers: int = 0
    sample_pages: int = OCR_CLASSIFY_SAMPLE_PAGES

    def normalized(self) -> "OcrConfig":
        language = str(self.language or "auto").strip().casefold()
        if language not in {"auto", "ja", "zh", "en"}:
            language = "auto"
        engine = str(self.engine or "auto").strip().casefold()
        if engine not in {"auto", "manga_ocr", "paddleocr", "tesseract"}:
            engine = "auto"
        dpi = max(200, min(400, int(self.dpi or OCR_DEFAULT_DPI)))
        if dpi not in {200, 240, 300, 400}:
            dpi = min((200, 240, 300, 400), key=lambda item: abs(item - dpi))
        confidence = float(self.min_confidence)
        confidence = max(0.0, min(1.0, confidence))
        max_pages = max(0, int(self.max_pages or 0))
        mode = str(self.mode or "full").strip().casefold()
        if mode not in {"classify", "sample", "full"}:
            mode = "full"
        workers = max(0, int(self.workers or 0))
        workers = workers or default_worker_count()
        sample_pages = max(
            5,
            min(8, int(self.sample_pages or OCR_CLASSIFY_SAMPLE_PAGES)),
        )
        return OcrConfig(
            language=language,
            engine=engine,
            dpi=dpi,
            min_confidence=confidence,
            cache_dir=Path(self.cache_dir).expanduser(),
            max_pages=max_pages,
            mode=mode,
            workers=workers,
            sample_pages=sample_pages,
        )


@dataclass(frozen=True)
class RenderedPage:
    index: int
    path: Path
    sha256: str


@dataclass(frozen=True)
class RenderedBook:
    source_path: Path
    source_hash: str
    pdf_path: Path
    dpi: int
    pages: Tuple[RenderedPage, ...]
    expected_pages: int
    errors: Tuple[str, ...] = ()


@dataclass(frozen=True)
class OcrPageResult:
    text: str
    confidence: Optional[float] = None
    text_area_ratio: Optional[float] = None
    chars_per_page: Optional[int] = None


def default_worker_count() -> int:
    """Keep default parallelism useful without loading four model copies."""

    return max(1, min(4, os.cpu_count() or 1))


def _filesystem_path(path: Any) -> str:
    if kindle_reader is not None:
        return kindle_reader._filesystem_path(str(path))
    return str(path)


def _display_path(path: Any) -> str:
    if kindle_reader is not None:
        return kindle_reader._display_path(str(path))
    return str(path)


def _normalized_text(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).casefold()
    return "".join(char for char in text if char.isalnum())


def ocr_cache_key(
    source_hash: str,
    dpi: int,
    language: str,
    engine: str,
    engine_version: str,
    page_hash: Any = "",
) -> str:
    """Return a stable cache key for one page or one whole-book result."""

    payload: Dict[str, Any] = {
        "cache_version": OCR_CACHE_VERSION,
        "source_hash": str(source_hash or ""),
        "dpi": int(dpi or 0),
        "language": str(language or ""),
        "engine": str(engine or ""),
        "engine_version": str(engine_version or ""),
    }
    if isinstance(page_hash, (list, tuple)):
        payload["page_hashes"] = [str(value or "") for value in page_hash]
    else:
        payload["page_hash"] = str(page_hash or "")
    encoded = json.dumps(
        payload,
        ensure_ascii=True,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


_SOURCE_HASH_CACHE: Dict[Tuple[str, int, int], str] = {}


def source_file_hash(path: Path) -> str:
    """Hash a source file without modifying or copying it."""

    resolved = Path(path).expanduser().resolve()
    stat = resolved.stat()
    cache_key = (str(resolved), int(stat.st_size), int(stat.st_mtime_ns))
    cached = _SOURCE_HASH_CACHE.get(cache_key)
    if cached:
        return cached

    digest = hashlib.sha256()
    with open(_filesystem_path(resolved), "rb") as handle:
        while True:
            block = handle.read(1024 * 1024)
            if not block:
                break
            digest.update(block)
    value = digest.hexdigest()
    _SOURCE_HASH_CACHE[cache_key] = value
    return value


def _sha256_bytes(payload: bytes) -> str:
    return hashlib.sha256(payload).hexdigest()


def _read_json(path: Path) -> Optional[Any]:
    try:
        with open(_filesystem_path(path), "r", encoding="utf-8") as handle:
            return json.load(handle)
    except (OSError, ValueError):
        return None


def _atomic_write_json(path: Path, value: Any) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path: Optional[Path] = None
    try:
        with tempfile.NamedTemporaryFile(
            "w",
            encoding="utf-8",
            newline="",
            dir=path.parent,
            prefix=path.name + ".",
            suffix=".tmp",
            delete=False,
        ) as handle:
            temporary_path = Path(handle.name)
            json.dump(
                value,
                handle,
                ensure_ascii=False,
                indent=2,
                sort_keys=True,
            )
            handle.write("\n")
        os.replace(
            _filesystem_path(temporary_path),
            _filesystem_path(path),
        )
    finally:
        if temporary_path is not None and os.path.exists(
            _filesystem_path(temporary_path)
        ):
            os.unlink(_filesystem_path(temporary_path))


def _atomic_copy_file(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary_path: Optional[Path] = None
    try:
        with tempfile.NamedTemporaryFile(
            "wb",
            dir=target.parent,
            prefix=target.name + ".",
            suffix=".tmp",
            delete=False,
        ) as handle:
            temporary_path = Path(handle.name)
            with open(_filesystem_path(source), "rb") as source_handle:
                shutil.copyfileobj(source_handle, handle, 1024 * 1024)
        os.replace(
            _filesystem_path(temporary_path),
            _filesystem_path(target),
        )
    finally:
        if temporary_path is not None and os.path.exists(
            _filesystem_path(temporary_path)
        ):
            os.unlink(_filesystem_path(temporary_path))


def _source_directory(row: Dict[str, Any], base_dir: Optional[Path]) -> Optional[Path]:
    for value in (row.get("ebook_file"), row.get("source_file")):
        text = str(value or "").strip()
        if not text:
            continue
        path = Path(text).expanduser()
        try:
            if path.is_file():
                directory = path.parent
                if directory.name.casefold().endswith(".sdr"):
                    directory = directory.parent
                return directory
            if path.is_dir():
                return path
        except OSError:
            continue
    return Path(base_dir).expanduser() if base_dir else None


def resolve_ocr_source(
    row: Dict[str, Any],
    base_dir: Optional[Path] = None,
) -> Optional[Path]:
    """Resolve the local ebook behind one manga CSV row."""

    configured = str(row.get("ebook_file") or "").strip()
    if configured:
        path = Path(configured).expanduser()
        try:
            if path.is_file() and path.suffix.casefold() in SOURCE_EBOOK_EXTENSIONS:
                return path.resolve()
        except OSError:
            pass

    directory = _source_directory(row, base_dir)
    if directory is None:
        return None
    try:
        if not directory.is_dir():
            return None
        candidates = [
            path
            for path in directory.iterdir()
            if path.is_file()
            and path.suffix.casefold() in SOURCE_EBOOK_EXTENSIONS
        ]
    except OSError:
        return None
    if not candidates:
        return None

    title_key = _normalized_text(row.get("title"))
    asin = re.sub(r"\s+", "", str(row.get("asin") or "")).upper()
    source_stem = Path(str(row.get("source_file") or "")).stem
    source_key = _normalized_text(source_stem)

    ranked: List[Tuple[int, int, str, Path]] = []
    for candidate in candidates:
        name_key = _normalized_text(candidate.stem)
        score = 0
        if asin and asin.casefold() in candidate.name.casefold():
            score += 60
        if title_key and name_key == title_key:
            score += 50
        elif title_key and name_key.startswith(title_key):
            score += 35
        elif title_key and title_key in name_key:
            score += 20
        if source_key and source_key in name_key:
            score += 10
        if score:
            ranked.append((score, -len(candidate.name), candidate.name.casefold(), candidate))

    if ranked:
        return max(ranked)[3].resolve()
    if len(candidates) == 1:
        return candidates[0].resolve()
    return None


def _find_tool(
    configured: str,
    candidates: Sequence[Optional[str]],
) -> Optional[Path]:
    values: List[Optional[str]] = [configured]
    values.extend(candidates)
    for value in values:
        text = str(value or "").strip()
        if not text:
            continue
        path = Path(text).expanduser()
        try:
            if path.is_file():
                return path.resolve()
        except OSError:
            continue
    return None


def find_calibre_convert() -> Optional[Path]:
    configured = os.environ.get("KINDLE_OCR_EBOOK_CONVERT", "")
    candidates = (
        None
        if kindle_reader is None
        else kindle_reader._find_calibre_convert(),
        shutil.which("ebook-convert.exe"),
        shutil.which("ebook-convert"),
        r"C:\Program Files\Calibre2\ebook-convert.exe",
        r"C:\Program Files (x86)\Calibre2\ebook-convert.exe",
    )
    return _find_tool(configured, candidates)


def find_calibre_debug() -> Optional[Path]:
    configured = os.environ.get("KINDLE_OCR_CALIBRE_DEBUG", "").strip()
    candidates = (
        shutil.which("calibre-debug.exe"),
        shutil.which("calibre-debug"),
        r"C:\Program Files\Calibre2\calibre-debug.exe",
        r"C:\Program Files (x86)\Calibre2\calibre-debug.exe",
    )
    calibre = find_calibre_convert()
    if calibre is not None:
        candidates = (str(calibre.with_name("calibre-debug.exe")),) + candidates
    return _find_tool(configured, candidates)


def _kfx_input_root() -> Optional[Path]:
    configured = os.environ.get("KINDLE_KFX_INPUT_DIR", "").strip()
    candidates = [
        Path(configured) if configured else None,
        Path(__file__).resolve().parent
        / "third_party"
        / "calibre-kfx-input",
    ]
    frozen_root = getattr(sys, "_MEIPASS", "")
    if frozen_root:
        candidates.insert(0, Path(frozen_root) / "kfx-input")
    for candidate in candidates:
        if candidate is not None and (candidate / "kfxlib").is_dir():
            return candidate
    return None


def _kfx_position_cache_path(cache_dir: Path, source_hash: str) -> Path:
    key = ocr_cache_key(
        source_hash,
        0,
        "kfx-position-map",
        "kfxlib",
        "1",
        page_hash="position-map",
    )
    return cache_dir / "render" / f"{key}.positions.json"


def resolve_kfx_position_info(
    source_path: Path,
    cache_dir: Path,
    source_hash: Optional[str] = None,
    logger: Optional[logging.Logger] = None,
) -> Optional[Dict[str, Any]]:
    """Return the exact KFX position-map size and final PID."""

    source_path = Path(source_path).expanduser().resolve()
    if source_path.suffix.casefold() not in {".kfx", ".kfx-zip", ".azw8"}:
        return None
    cache_dir = Path(cache_dir).expanduser()
    source_hash = source_hash or source_file_hash(source_path)
    cache_path = _kfx_position_cache_path(cache_dir, source_hash)
    cached = _read_json(cache_path)
    if isinstance(cached, dict):
        last_pid = _parse_float(cached.get("last_pid"))
        last_pid = int(last_pid) if last_pid is not None else None
        if last_pid and last_pid > 0:
            cached["_cache_hit"] = True
            return cached

    debug = find_calibre_debug()
    kfx_root = _kfx_input_root()
    if debug is None or kfx_root is None:
        return None

    script = (
        "import json,os,sys\n"
        f"sys.path.insert(0, {str(kfx_root)!r})\n"
        f"sys.path.insert(0, {str(kfx_root / 'kfxlib' / 'calibre-plugin-modules')!r})\n"
        "from kfxlib.yj_book import YJ_Book\n"
        f"book=YJ_Book({str(source_path)!r})\n"
        "book.decode_book()\n"
        "info=book.collect_position_map_info()\n"
        "print(json.dumps({'count':len(info),'last_pid':"
        "(info[-1].pid if info else 0)}))\n"
    )
    config_dir = cache_dir / "calibre-config"
    config_dir.mkdir(parents=True, exist_ok=True)
    env = dict(os.environ)
    env["CALIBRE_CONFIG_DIRECTORY"] = str(config_dir)
    kwargs: Dict[str, Any] = {}
    if os.name == "nt":
        kwargs["creationflags"] = 0x08000000
    try:
        result = subprocess.run(
            [str(debug), "-c", script],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=180,
            env=env,
            **kwargs,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        if logger is not None:
            logger.info("KFX position-map read failed: %s", exc)
        return None
    payload: Optional[Dict[str, Any]] = None
    for line in reversed((result.stdout or "").splitlines()):
        try:
            decoded = json.loads(line)
        except ValueError:
            continue
        if isinstance(decoded, dict) and decoded.get("last_pid"):
            payload = decoded
            break
    if payload is None:
        if logger is not None:
            logger.info(
                "KFX position-map read returned no data: %s",
                (result.stderr or result.stdout or "").strip()[-1000:],
            )
        return None
    payload["_cache_hit"] = False
    try:
        _atomic_write_json(cache_path, payload)
    except OSError:
        pass
    return payload


def position_number(value: Any) -> Optional[int]:
    text = str(value or "").strip()
    if not text:
        return None
    match = re.search(r":(\d+)$", text)
    if match:
        return int(match.group(1))
    number = _parse_float(text)
    return int(number) if number is not None else None


def _poppler_candidates(tool_name: str) -> List[Optional[str]]:
    candidates: List[Optional[str]] = [
        shutil.which(tool_name + ".exe"),
        shutil.which(tool_name),
    ]
    roots: List[Path] = []
    for root in _APPLICATION_ROOTS:
        roots.extend(
            (
                root / "poppler" / "Library" / "bin",
                root / "native" / "poppler" / "Library" / "bin",
                root / ".ocr-pip" / "poppler" / "Library" / "bin",
            )
        )
        roots.extend(parent for parent in root.parents[:3])

    runtime = Path(sys.executable).resolve()
    roots.extend(
        (
            runtime.parent,
            runtime.parent.parent,
            runtime.parent.parent / "native" / "poppler" / "Library" / "bin",
        )
    )
    for env_name in ("CONDA_PREFIX", "VIRTUAL_ENV", "ProgramFiles", "ProgramFiles(x86)"):
        configured = os.environ.get(env_name, "").strip()
        if configured:
            roots.append(Path(configured))
    local_app_data = os.environ.get("LOCALAPPDATA", "").strip()
    if local_app_data:
        roots.append(Path(local_app_data) / "Programs" / "Calibre2")

    calibre = find_calibre_convert()
    if calibre is not None:
        roots.extend((calibre.parent, calibre.parent / "app" / "bin"))

    for root in roots:
        for candidate in (
            root / tool_name,
            root / f"{tool_name}.exe",
            root / "Library" / "bin" / tool_name,
            root / "Library" / "bin" / f"{tool_name}.exe",
            root / "app" / "bin" / tool_name,
            root / "app" / "bin" / f"{tool_name}.exe",
        ):
            candidates.append(str(candidate))
    return candidates


def find_pdftoppm() -> Optional[Path]:
    configured = os.environ.get("KINDLE_OCR_PDFTOPPM", "")
    return _find_tool(configured, _poppler_candidates("pdftoppm"))


def find_pdfinfo() -> Optional[Path]:
    configured = os.environ.get("KINDLE_OCR_PDFINFO", "")
    return _find_tool(configured, _poppler_candidates("pdfinfo"))


def _run_command(
    command: Sequence[str],
    timeout: int,
    logger: Optional[logging.Logger] = None,
) -> subprocess.CompletedProcess[str]:
    kwargs: Dict[str, Any] = {}
    if os.name == "nt":
        kwargs["creationflags"] = 0x08000000
    if logger is not None:
        logger.info("OCR command: %s", " ".join(command))
    return subprocess.run(
        list(command),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        **kwargs,
    )


def _tool_version(tool: Path) -> str:
    try:
        result = subprocess.run(
            [str(tool), "--version"],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=15,
            creationflags=0x08000000 if os.name == "nt" else 0,
        )
        line = (result.stdout or "").strip().splitlines()
        return line[0].strip() if line else "unknown"
    except (OSError, subprocess.SubprocessError):
        return "unknown"


def _pdf_page_count(pdf_path: Path) -> Optional[int]:
    pdfinfo = find_pdfinfo()
    if pdfinfo is None:
        return None
    try:
        result = subprocess.run(
            [str(pdfinfo), str(pdf_path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=60,
            creationflags=0x08000000 if os.name == "nt" else 0,
        )
    except (OSError, subprocess.SubprocessError):
        return None
    if result.returncode != 0:
        return None
    match = re.search(r"^Pages:\s+(\d+)\s*$", result.stdout, re.MULTILINE)
    return int(match.group(1)) if match else None


def _page_count_cache_path(cache_dir: Path, source_hash: str) -> Path:
    key = ocr_cache_key(
        source_hash,
        0,
        "pdfinfo",
        "pdfinfo",
        OCR_RENDER_VERSION,
        page_hash="page-count",
    )
    return cache_dir / "render" / f"{key}.pagecount.json"


def resolve_manga_page_count(
    source_path: Path,
    cache_dir: Path,
    source_hash: Optional[str] = None,
    logger: Optional[logging.Logger] = None,
) -> Tuple[Optional[int], str, Optional[str]]:
    """Read page count from the local source PDF, with a small JSON cache."""

    source_path = Path(source_path).expanduser().resolve()
    cache_dir = Path(cache_dir).expanduser()
    source_hash = source_hash or source_file_hash(source_path)
    cache_path = _page_count_cache_path(cache_dir, source_hash)
    cached = _read_json(cache_path)
    if isinstance(cached, dict):
        pages = _parse_float(cached.get("pages"))
        if pages is not None and pages > 0:
            return int(pages), str(cached.get("source") or "pdfinfo"), None

    try:
        pdf_path = _convert_to_pdf(
            source_path,
            cache_dir,
            source_hash,
            logger=logger,
        )
        pages = _pdf_page_count(pdf_path)
    except OcrDependencyError as exc:
        return None, "unknown", str(exc)
    except OcrRenderError as exc:
        return None, "unknown", str(exc)
    except Exception as exc:
        return None, "unknown", f"{type(exc).__name__}: {exc}"

    if pages is None or pages <= 0:
        return None, "unknown", "pdfinfo did not report a positive page count"
    try:
        _atomic_write_json(
            cache_path,
            {
                "pages": pages,
                "source": "pdfinfo-source-file",
                "note": (
                    "Page count comes from the converted/source file, which "
                    "may differ from Kindle's internal pagination."
                ),
            },
        )
    except OSError:
        pass
    return pages, "pdfinfo-source-file", None


def sample_page_numbers(total_pages: int, count: int) -> Tuple[int, ...]:
    """Pick evenly spaced middle pages, avoiding covers and front matter."""

    if total_pages <= 0:
        return ()
    count = max(1, min(8, int(count or OCR_CLASSIFY_SAMPLE_PAGES)))
    if total_pages <= count:
        return tuple(range(1, total_pages + 1))
    start = max(1, int(round(total_pages * 0.15)))
    end = min(total_pages, int(round(total_pages * 0.85)))
    if end <= start:
        start = max(1, total_pages // 2)
        end = min(total_pages, start + count - 1)
    span = max(1, end - start)
    values = {
        min(
            end,
            max(start, int(round(start + span * index / max(1, count - 1)))),
        )
        for index in range(count)
    }
    return tuple(sorted(values))


def _render_page_range(
    pdftoppm: Path,
    pdf_path: Path,
    output_dir: Path,
    first: int,
    last: int,
    dpi: int,
    logger: Optional[logging.Logger] = None,
) -> List[Path]:
    prefix = f"page-{first:05d}"
    command = [
        str(pdftoppm),
        "-png",
        "-r",
        str(dpi),
        "-f",
        str(first),
        "-l",
        str(last),
        str(pdf_path),
        str(output_dir / prefix),
    ]
    result = _run_command(
        command,
        OCR_PDFTOPPM_TIMEOUT_SECONDS,
        logger=logger,
    )
    if result.returncode != 0:
        raise OcrRenderError(
            f"pdftoppm exited {result.returncode}: "
            f"{(result.stdout or '').strip()}"
        )
    return sorted(output_dir.glob(f"{prefix}-*.png"))


def _render_cache_key(source_hash: str, dpi: int) -> str:
    return ocr_cache_key(
        source_hash,
        dpi,
        "render",
        "pdftoppm",
        OCR_RENDER_VERSION,
        page_hash="manifest",
    )


def _load_render_manifest(
    cache_dir: Path,
    source_hash: str,
    dpi: int,
    source_path: Path,
    pdf_path: Path,
) -> Optional[RenderedBook]:
    key = _render_cache_key(source_hash, dpi)
    manifest = _read_json(cache_dir / "render" / f"{key}.json")
    if not isinstance(manifest, dict):
        return None
    raw_pages = manifest.get("pages")
    if not isinstance(raw_pages, list):
        return None
    pages: List[RenderedPage] = []
    for item in raw_pages:
        if not isinstance(item, dict):
            return None
        page_hash = str(item.get("sha256") or "")
        index = int(item.get("index") or 0)
        page_path = cache_dir / "pages" / f"{page_hash}.png"
        if not page_hash or not page_path.is_file():
            return None
        pages.append(RenderedPage(index=index, path=page_path, sha256=page_hash))
    return RenderedBook(
        source_path=source_path,
        source_hash=source_hash,
        pdf_path=pdf_path,
        dpi=dpi,
        pages=tuple(sorted(pages, key=lambda page: page.index)),
        expected_pages=max(
            int(manifest.get("expected_pages") or 0),
            len(pages),
        ),
        errors=tuple(str(value) for value in manifest.get("errors") or []),
    )


def _save_render_manifest(cache_dir: Path, book: RenderedBook) -> None:
    key = _render_cache_key(book.source_hash, book.dpi)
    _atomic_write_json(
        cache_dir / "render" / f"{key}.json",
        {
            "source_path": str(book.source_path),
            "source_hash": book.source_hash,
            "pdf_path": str(book.pdf_path),
            "dpi": book.dpi,
            "expected_pages": book.expected_pages,
            "errors": list(book.errors),
            "pages": [
                {"index": page.index, "sha256": page.sha256}
                for page in book.pages
            ],
        },
    )


def _convert_to_pdf(
    source_path: Path,
    cache_dir: Path,
    source_hash: str,
    logger: Optional[logging.Logger] = None,
) -> Path:
    if source_path.suffix.casefold() == ".pdf":
        return source_path
    calibre = find_calibre_convert()
    if calibre is None:
        raise OcrDependencyError(
            "Calibre ebook-convert is required to render non-PDF manga. "
            "Install Calibre or set KINDLE_OCR_EBOOK_CONVERT."
        )

    key = ocr_cache_key(
        source_hash,
        0,
        "render",
        "calibre",
        _tool_version(calibre),
        page_hash="pdf",
    )
    target = cache_dir / "render" / f"{key}.pdf"
    if target.is_file() and target.stat().st_size > 0:
        return target
    target.parent.mkdir(parents=True, exist_ok=True)

    temporary_path = target.with_name(target.name + ".tmp.pdf")
    temporary_path.unlink(missing_ok=True)
    command = [str(calibre), str(source_path), str(temporary_path)]
    last_error = ""
    for attempt in range(1, 3):
        temporary_path.unlink(missing_ok=True)
        try:
            result = _run_command(
                command,
                OCR_CONVERSION_TIMEOUT_SECONDS,
                logger=logger,
            )
        except subprocess.TimeoutExpired:
            last_error = "Calibre conversion timed out"
        except OSError as exc:
            last_error = f"{type(exc).__name__}: {exc}"
        else:
            if result.returncode == 0 and temporary_path.is_file():
                os.replace(
                    _filesystem_path(temporary_path),
                    _filesystem_path(target),
                )
                return target
            detail = (result.stdout or "").strip()
            last_error = (
                f"Calibre exited {result.returncode}"
                + (f": {detail}" if detail else "")
            )
        if attempt == 1:
            time.sleep(0.5)
    temporary_path.unlink(missing_ok=True)
    raise OcrRenderError(
        "Calibre could not render the manga ebook to PDF: " + last_error
    )


def _natural_page_index(path: Path, prefix: str) -> int:
    match = re.search(r"(\d+)$", path.stem[len(prefix) :])
    return int(match.group(1)) if match else 0


def render_manga_pages(
    source_path: Path,
    cache_dir: Path,
    dpi: int,
    source_hash: Optional[str] = None,
    max_pages: int = 0,
    workers: int = 1,
    logger: Optional[logging.Logger] = None,
) -> RenderedBook:
    """Render local manga pages into project-side PNG/PDF caches."""

    source_path = Path(source_path).expanduser().resolve()
    cache_dir = Path(cache_dir).expanduser()
    source_hash = source_hash or source_file_hash(source_path)
    cached = _load_render_manifest(
        cache_dir,
        source_hash,
        dpi,
        source_path,
        source_path,
    )
    if cached is not None and (not max_pages or len(cached.pages) <= max_pages):
        return cached

    pdf_path = _convert_to_pdf(source_path, cache_dir, source_hash, logger)
    cached = _load_render_manifest(
        cache_dir,
        source_hash,
        dpi,
        source_path,
        pdf_path,
    )
    if cached is not None and (not max_pages or len(cached.pages) <= max_pages):
        return cached

    pdftoppm = find_pdftoppm()
    if pdftoppm is None:
        raise OcrDependencyError(
            "pdftoppm is required to render manga pages. Install Poppler "
            "or set KINDLE_OCR_PDFTOPPM."
        )

    render_key = _render_cache_key(source_hash, dpi)
    temporary_root = cache_dir / "tmp" / render_key
    temporary_root.mkdir(parents=True, exist_ok=True)
    prefix = "page"
    rendered_files: List[Path] = []
    errors: List[str] = []
    try:
        expected_for_split = _pdf_page_count(pdf_path) or 0
        workers = max(1, int(workers or 1))
        if max_pages:
            expected_for_split = min(
                max_pages,
                expected_for_split or max_pages,
            )
        if expected_for_split > 1 and workers > 1:
            workers = min(workers, expected_for_split)
            chunk_size = max(
                1,
                math.ceil(expected_for_split / float(workers)),
            )
            ranges = []
            first = 1
            while first <= expected_for_split:
                last = min(expected_for_split, first + chunk_size - 1)
                ranges.append((first, last))
                first = last + 1

            def render_range(span: Tuple[int, int]) -> List[Path]:
                first, last = span
                chunk_dir = temporary_root / f"{first:05d}-{last:05d}"
                chunk_dir.mkdir(parents=True, exist_ok=True)
                return _render_page_range(
                    pdftoppm,
                    pdf_path,
                    chunk_dir,
                    first,
                    last,
                    dpi,
                    logger=logger,
                )

            with concurrent.futures.ThreadPoolExecutor(
                max_workers=workers
            ) as executor:
                for chunk_files in executor.map(render_range, ranges):
                    rendered_files.extend(chunk_files)
        else:
            command = [
                str(pdftoppm),
                "-png",
                "-r",
                str(dpi),
            ]
            if max_pages:
                command.extend(["-f", "1", "-l", str(max_pages)])
            command.extend([str(pdf_path), str(temporary_root / prefix)])
            result = _run_command(
                command,
                OCR_PDFTOPPM_TIMEOUT_SECONDS,
                logger=logger,
            )
            rendered_files = sorted(
                temporary_root.glob(f"{prefix}-*.png"),
                key=lambda path: _natural_page_index(path, prefix),
            )
            if result.returncode != 0:
                errors.append(
                    f"pdftoppm exited {result.returncode}: "
                    f"{(result.stdout or '').strip()}"
                )
        rendered_files = sorted(
            rendered_files,
            key=lambda path: _natural_page_index(path, prefix),
        )
        if not rendered_files:
            raise OcrRenderError(
                "pdftoppm did not create any page images: "
                + "; ".join(errors)
            )

        pages: List[RenderedPage] = []
        for path in rendered_files:
            payload = path.read_bytes()
            page_hash = _sha256_bytes(payload)
            target = cache_dir / "pages" / f"{page_hash}.png"
            if not target.is_file() or target.stat().st_size != len(payload):
                _atomic_copy_file(path, target)
            pages.append(
                RenderedPage(
                    index=_natural_page_index(path, prefix),
                    path=target,
                    sha256=page_hash,
                )
            )

        expected_pages = expected_for_split
        if max_pages:
            expected_pages = min(
                max_pages,
                expected_pages or len(pages),
            )
        expected_pages = max(expected_pages, len(pages))
        if expected_pages > len(pages):
            errors.append(
                f"rendered {len(pages)} of {expected_pages} expected pages"
            )
        book = RenderedBook(
            source_path=source_path,
            source_hash=source_hash,
            pdf_path=pdf_path,
            dpi=dpi,
            pages=tuple(sorted(pages, key=lambda page: page.index)),
            expected_pages=expected_pages,
            errors=tuple(errors),
        )
        _save_render_manifest(cache_dir, book)
        return book
    finally:
        shutil.rmtree(
            _filesystem_path(temporary_root),
            ignore_errors=True,
        )


def render_manga_sample_pages(
    source_path: Path,
    cache_dir: Path,
    dpi: int,
    source_hash: Optional[str] = None,
    sample_pages: int = OCR_CLASSIFY_SAMPLE_PAGES,
    workers: int = 0,
    logger: Optional[logging.Logger] = None,
) -> RenderedBook:
    """Render only a small middle slice of a PDF for cheap classification."""

    source_path = Path(source_path).expanduser().resolve()
    cache_dir = Path(cache_dir).expanduser()
    source_hash = source_hash or source_file_hash(source_path)
    total_pages, _, count_error = resolve_manga_page_count(
        source_path,
        cache_dir,
        source_hash=source_hash,
        logger=logger,
    )
    if total_pages is None:
        raise OcrRenderError(count_error or "Manga page count is unknown.")

    pdf_path = _convert_to_pdf(
        source_path,
        cache_dir,
        source_hash,
        logger=logger,
    )
    pdftoppm = find_pdftoppm()
    if pdftoppm is None:
        raise OcrDependencyError(
            "pdftoppm is required to render manga pages. Install Poppler "
            "or set KINDLE_OCR_PDFTOPPM."
        )

    indexes = sample_page_numbers(total_pages, sample_pages)
    workers = max(
        1,
        min(int(workers or default_worker_count()), len(indexes) or 1),
    )
    sample_root = (
        cache_dir
        / "tmp"
        / _render_cache_key(source_hash, dpi)
        / "sample"
    )
    sample_root.mkdir(parents=True, exist_ok=True)
    errors: List[str] = []

    def render_one(page_number: int) -> List[Path]:
        page_dir = sample_root / f"{page_number:05d}"
        page_dir.mkdir(parents=True, exist_ok=True)
        return _render_page_range(
            pdftoppm,
            pdf_path,
            page_dir,
            page_number,
            page_number,
            dpi,
            logger=logger,
        )

    rendered_pairs: List[Tuple[int, Path]] = []
    try:
        if workers <= 1:
            for page_number in indexes:
                for path in render_one(page_number):
                    rendered_pairs.append((page_number, path))
        else:
            with concurrent.futures.ThreadPoolExecutor(
                max_workers=workers
            ) as executor:
                future_map = {
                    executor.submit(render_one, page_number): page_number
                    for page_number in indexes
                }
                for future in concurrent.futures.as_completed(future_map):
                    page_number = future_map[future]
                    try:
                        for path in future.result():
                            rendered_pairs.append((page_number, path))
                    except Exception as exc:
                        errors.append(
                            f"page {page_number}: {type(exc).__name__}: {exc}"
                        )

        pages: List[RenderedPage] = []
        for page_number, path in sorted(rendered_pairs, key=lambda item: item[0]):
            payload = path.read_bytes()
            page_hash = _sha256_bytes(payload)
            target = cache_dir / "pages" / f"{page_hash}.png"
            if not target.is_file() or target.stat().st_size != len(payload):
                _atomic_copy_file(path, target)
            pages.append(
                RenderedPage(
                    index=page_number,
                    path=target,
                    sha256=page_hash,
                )
            )
        if not pages:
            raise OcrRenderError(
                "pdftoppm did not create any sampled page images: "
                + "; ".join(errors)
            )
        return RenderedBook(
            source_path=source_path,
            source_hash=source_hash,
            pdf_path=pdf_path,
            dpi=dpi,
            pages=tuple(pages),
            expected_pages=total_pages,
            errors=tuple(errors),
        )
    finally:
        shutil.rmtree(_filesystem_path(sample_root), ignore_errors=True)


def _package_version(name: str) -> str:
    try:
        from importlib.metadata import version

        return version(name)
    except Exception:
        return "unknown"


def _find_tesseract() -> Optional[Path]:
    configured = os.environ.get("KINDLE_OCR_TESSERACT", "")
    candidates = (
        shutil.which("tesseract.exe"),
        shutil.which("tesseract"),
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    )
    return _find_tool(configured, candidates)


def ocr_engine_version(name: str) -> str:
    """Return an engine version before loading an expensive local model."""

    normalized = str(name or "").strip().casefold()
    if normalized == "manga_ocr":
        return _package_version("manga-ocr")
    if normalized == "paddleocr":
        return _package_version("paddleocr")
    if normalized == "tesseract":
        binary = _find_tesseract()
        return _tool_version(binary) if binary is not None else "unknown"
    return "unknown"


def paddle_model_root() -> Optional[Path]:
    configured = os.environ.get("KINDLE_OCR_PADDLEOCR_MODEL_DIR", "").strip()
    if configured:
        path = Path(configured).expanduser()
        return path if path.is_dir() else None
    candidate = _MODEL_ROOT / "paddleocr"
    return candidate if candidate.is_dir() else None


def _paddle_language_key(language: str) -> str:
    normalized = str(language or "").casefold()
    return {
        "zh": "ch",
        "en": "en",
        "ja": "japan",
    }.get(normalized, normalized or "ch")


def _paddle_model_directory(
    model_root: Path,
    kind: str,
    language: str,
) -> Optional[Path]:
    language_key = _paddle_language_key(language)
    candidates = (
        model_root / f"{kind}_{language_key}",
        model_root / language_key / kind,
        model_root / kind / language_key,
        model_root / kind,
    )
    for candidate in candidates:
        if candidate.is_dir():
            current = candidate
            for _ in range(3):
                if (current / "inference.pdmodel").is_file():
                    return current
                children = [
                    child
                    for child in current.iterdir()
                    if child.is_dir()
                ]
                if len(children) != 1:
                    break
                current = children[0]
    return None


def engine_available(name: str) -> Tuple[bool, str]:
    normalized = str(name or "").casefold()
    if normalized == "manga_ocr":
        if importlib.util.find_spec("manga_ocr") is None:
            return False, DEPENDENCY_HINTS["manga_ocr"]
        return True, ""
    if normalized == "paddleocr":
        if importlib.util.find_spec("paddleocr") is None:
            return False, DEPENDENCY_HINTS["paddleocr"]
        if paddle_model_root() is None:
            return False, DEPENDENCY_HINTS["paddleocr"]
        return True, ""
    if normalized == "tesseract":
        if _find_tesseract() is None:
            return False, DEPENDENCY_HINTS["tesseract"]
        return True, ""
    return False, f"Unknown OCR engine: {name}"


def resolve_ocr_engine(requested: str, language: str) -> str:
    """Choose an engine, preferring offline-capable adapters per language."""

    normalized = str(requested or "auto").casefold()
    if normalized != "auto":
        return normalized
    priority = {
        "ja": ("manga_ocr", "paddleocr"),
        "zh": ("paddleocr",),
        "en": ("tesseract", "paddleocr"),
    }.get(str(language or "").casefold(), ())
    for candidate in priority:
        available, _ = engine_available(candidate)
        if available:
            return candidate
    return priority[0] if priority else "tesseract"


def _language_from_text_sample(value: Any) -> Optional[str]:
    text = str(value or "")
    kana = len(re.findall(r"[\u3040-\u30ff]", text))
    cjk = len(re.findall(r"[\u3400-\u9fff]", text))
    latin = len(re.findall(r"[A-Za-z]", text))
    if kana >= 2:
        return "ja"
    if cjk >= 4 and kana == 0:
        return "zh"
    if latin >= 10 and cjk == 0:
        return "en"
    return None


def resolve_ocr_language(
    row: Dict[str, Any],
    requested: str,
) -> Optional[str]:
    """Resolve auto language conservatively; uncertainty returns None."""

    normalized = str(requested or "auto").casefold()
    if normalized in {"ja", "zh", "en"}:
        return normalized
    metadata_language = str(row.get("word_count_language") or "").casefold()
    if metadata_language in {"ja", "zh", "en"}:
        return metadata_language
    sample = " ".join(
        (
            str(row.get("title") or ""),
            str(row.get("author") or ""),
            Path(str(row.get("source_file") or "")).stem,
        )
    )
    return _language_from_text_sample(sample)


class _TesseractEngine:
    name = "tesseract"

    def __init__(self) -> None:
        binary = _find_tesseract()
        if binary is None:
            raise OcrDependencyError(DEPENDENCY_HINTS["tesseract"])
        self.binary = binary
        self.version = _tool_version(binary)

    def recognize(self, image_path: Path, language: str) -> OcrPageResult:
        command = [
            str(self.binary),
            str(image_path),
            "stdout",
            "-l",
            language,
            "--psm",
            "6",
            "tsv",
        ]
        result = _run_command(command, 300)
        if result.returncode != 0:
            raise RuntimeError(
                f"tesseract exited {result.returncode}: "
                f"{(result.stdout or '').strip()}"
            )
        texts: List[str] = []
        confidences: List[float] = []
        box_area = 0.0
        rows = csv.DictReader(io.StringIO(result.stdout or ""), delimiter="\t")
        for row in rows:
            text = str(row.get("text") or "").strip()
            if not text:
                continue
            texts.append(text)
            width = _parse_float(row.get("width")) or 0.0
            height = _parse_float(row.get("height")) or 0.0
            box_area += max(0.0, width * height)
            confidence = _parse_float(row.get("conf"))
            if confidence is not None and confidence >= 0:
                confidences.append(confidence / 100.0)
        confidence_value = (
            sum(confidences) / len(confidences) if confidences else None
        )
        image_area = _image_pixel_area(image_path)
        return OcrPageResult(
            " ".join(texts).strip(),
            confidence_value,
            text_area_ratio=(
                min(1.0, box_area / image_area) if image_area else None
            ),
            chars_per_page=_count_chars(" ".join(texts)),
        )


class _MangaOcrEngine:
    name = "manga_ocr"

    def __init__(self) -> None:
        available, message = engine_available("manga_ocr")
        if not available:
            raise OcrDependencyError(message)
        os.environ.setdefault("HF_HUB_OFFLINE", "1")
        os.environ.setdefault("TRANSFORMERS_OFFLINE", "1")
        from manga_ocr import MangaOcr  # type: ignore

        model_dir = os.environ.get("KINDLE_OCR_MANGA_MODEL_DIR", "").strip()
        kwargs: Dict[str, Any] = {}
        if model_dir:
            kwargs["pretrained_model_name_or_path"] = model_dir
        self._engine = MangaOcr(**kwargs)
        self.version = ocr_engine_version("manga_ocr")

    def recognize(self, image_path: Path, language: str) -> OcrPageResult:
        del language
        text = str(self._engine(str(image_path)) or "")
        chars = _count_chars(text)
        return OcrPageResult(
            text,
            None,
            text_area_ratio=min(1.0, chars / 2500.0),
            chars_per_page=chars,
        )


class _PaddleOcrEngine:
    name = "paddleocr"

    def __init__(self, language: str) -> None:
        available, message = engine_available("paddleocr")
        if not available:
            raise OcrDependencyError(message)
        # PaddleOCR 2.x imports Albumentations, which imports Torch when it is
        # present. Preload Torch after the project DLL directories are set so
        # Windows resolves its bundled runtime libraries consistently.
        try:
            import torch  # type: ignore  # noqa: F401
        except Exception:
            pass
        from paddleocr import PaddleOCR  # type: ignore

        model_dir = paddle_model_root()
        if model_dir is None:
            raise OcrDependencyError(DEPENDENCY_HINTS["paddleocr"])
        det_dir = _paddle_model_directory(model_dir, "det", language)
        rec_dir = _paddle_model_directory(model_dir, "rec", language)
        cls_dir = _paddle_model_directory(model_dir, "cls", language)
        if rec_dir is None:
            raise OcrDependencyError(
                "Missing local PaddleOCR recognition model for "
                f"{language}: {model_dir}"
            )
        paddle_language = _paddle_language_key(language)
        kwargs: Dict[str, Any] = {
            "lang": paddle_language,
            "use_angle_cls": True,
            "show_log": False,
            "use_gpu": False,
            "enable_mkldnn": False,
        }
        for key, candidate in (
            ("det_model_dir", det_dir),
            ("rec_model_dir", rec_dir),
            ("cls_model_dir", cls_dir),
        ):
            if candidate is not None:
                kwargs[key] = str(candidate)
        try:
            self._engine = PaddleOCR(**kwargs)
        except TypeError:
            kwargs.pop("use_angle_cls", None)
            self._engine = PaddleOCR(**kwargs)
        self.version = ocr_engine_version("paddleocr")

    def recognize(self, image_path: Path, language: str) -> OcrPageResult:
        del language
        result = self._engine.ocr(str(image_path), cls=True)
        texts: List[str] = []
        confidences: List[float] = []
        box_area = 0.0
        for item in _walk_paddle_items(result):
            if not isinstance(item, (list, tuple)) or len(item) < 2:
                continue
            polygon = item[0]
            if isinstance(polygon, (list, tuple)) and len(polygon) >= 4:
                try:
                    xs = [float(point[0]) for point in polygon]
                    ys = [float(point[1]) for point in polygon]
                    box_area += max(0.0, max(xs) - min(xs)) * max(
                        0.0, max(ys) - min(ys)
                    )
                except (TypeError, ValueError, IndexError):
                    pass
            payload = item[1]
            if isinstance(payload, (list, tuple)) and len(payload) >= 2:
                text = str(payload[0] or "").strip()
                confidence = _parse_float(payload[1])
            else:
                text = str(payload or "").strip()
                confidence = None
            if text:
                texts.append(text)
            if confidence is not None:
                confidences.append(confidence)
        value = sum(confidences) / len(confidences) if confidences else None
        image_area = _image_pixel_area(image_path)
        text = "\n".join(texts).strip()
        return OcrPageResult(
            text,
            value,
            text_area_ratio=(
                min(1.0, box_area / image_area) if image_area else None
            ),
            chars_per_page=_count_chars(text),
        )


def _walk_paddle_items(value: Any) -> Iterable[Any]:
    if isinstance(value, (list, tuple)):
        if (
            len(value) >= 2
            and isinstance(value[1], (list, tuple))
            and value[1]
            and isinstance(value[1][0], str)
        ):
            yield value
            return
        for item in value:
            yield from _walk_paddle_items(item)


def create_ocr_engine(
    name: str,
    language: str,
) -> Any:
    normalized = str(name or "").casefold()
    if normalized == "tesseract":
        return _TesseractEngine()
    if normalized == "manga_ocr":
        return _MangaOcrEngine()
    if normalized == "paddleocr":
        return _PaddleOcrEngine(language)
    raise OcrDependencyError(f"Unknown OCR engine: {name}")


def _parse_float(value: Any) -> Optional[float]:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def _count_chars(text: str) -> int:
    return sum(1 for char in text if not char.isspace())


def _image_pixel_area(image_path: Path) -> float:
    try:
        from PIL import Image  # type: ignore

        with Image.open(image_path) as image:
            return float(max(1, image.width * image.height))
    except Exception:
        return 0.0


def _count_words(text: str, language: str) -> int:
    if not text.strip():
        return 0
    if kindle_reader is not None:
        try:
            count, _ = kindle_reader.count_text_words(text, language)
            return max(0, int(count))
        except Exception:
            pass
    if language == "en":
        return len(
            re.findall(r"[A-Za-z0-9]+(?:['-][A-Za-z0-9]+)*", text)
        )
    return len(
        re.findall(r"[A-Za-z0-9]+|[\u3040-\u30ff\u3400-\u9fff]", text)
    )


def _empty_result(
    status: str,
    message: str,
    *,
    language: str = "",
    engine: str = "",
    cache_key: str = "",
    source_hash: str = "",
    pages_total: int = 0,
    pages_processed: int = 0,
    page_errors: Sequence[str] = (),
    min_confidence: Optional[float] = None,
) -> Dict[str, Any]:
    confidence = None
    coverage = (
        round(pages_processed / pages_total * 100.0, 3)
        if pages_total
        else None
    )
    return {
        "ocr_mode": "",
        "ocr_word_count": None,
        "ocr_word_count_is_estimate": False,
        "ocr_char_count": None,
        "ocr_confidence": confidence,
        "ocr_engine": engine,
        "ocr_language": language,
        "ocr_page_coverage": coverage,
        "ocr_status": status,
        "ocr_error": message,
        "ocr_completed_at": datetime.datetime.now(
            datetime.timezone.utc
        ).isoformat(timespec="seconds"),
        "ocr_cache_key": cache_key,
        "ocr_source": OCR_SOURCE,
        "ocr_pages_total": pages_total,
        "ocr_pages_processed": pages_processed,
        "ocr_failed_pages": len(page_errors),
        "ocr_elapsed_seconds": None,
        "ocr_peak_memory_mb": None,
        "ocr_min_confidence": (
            round(max(0.0, min(1.0, float(min_confidence))), 4)
            if min_confidence is not None
            else OCR_DEFAULT_MIN_CONFIDENCE
        ),
        "ocr_sampled_pages": 0,
        "ocr_extrapolated": False,
        "classification_source": "",
        "classification_confidence": None,
        "classification_reason": "",
        "manga_pages_total": pages_total or None,
        "manga_pages_read": None,
        "manga_pages_source": "ocr" if pages_total else None,
        "_ocr_source_hash": source_hash,
    }


def _page_cache_path(cache_dir: Path, key: str) -> Path:
    return cache_dir / "results" / f"page-{key}.json"


def _book_cache_path(cache_dir: Path, key: str) -> Path:
    return cache_dir / "results" / f"book-{key}.json"


def _read_page_cache(
    cache_dir: Path,
    key: str,
) -> Optional[OcrPageResult]:
    decoded = _read_json(_page_cache_path(cache_dir, key))
    if not isinstance(decoded, dict) or decoded.get("status") != "ok":
        return None
    return OcrPageResult(
        text=str(decoded.get("text") or ""),
        confidence=_parse_float(decoded.get("confidence")),
    )


def _write_page_cache(
    cache_dir: Path,
    key: str,
    result: OcrPageResult,
) -> None:
    _atomic_write_json(
        _page_cache_path(cache_dir, key),
        {
            "status": "ok",
            "text": result.text,
            "confidence": result.confidence,
        },
    )


def _result_cache_payload(result: Dict[str, Any]) -> Dict[str, Any]:
    return {
        key: value
        for key, value in result.items()
        if not str(key).startswith("_")
    }


def _book_page_key(
    source_hash: str,
    dpi: int,
    language: str,
    engine: str,
    engine_version: str,
    pages: Sequence[RenderedPage],
) -> str:
    return ocr_cache_key(
        source_hash,
        dpi,
        language,
        engine,
        engine_version,
        page_hash=[page.sha256 for page in pages],
    )


def _row_number(row: Dict[str, Any], *keys: str) -> Optional[float]:
    for key in keys:
        value = _parse_float(row.get(key))
        if value is not None:
            return value
    return None


def _cheap_classification(
    row: Dict[str, Any],
    source_path: Optional[Path],
) -> Optional[Dict[str, Any]]:
    suffix = Path(source_path).suffix.casefold() if source_path else ""
    if suffix in {".cbz", ".cbr", ".zip"}:
        return {
            "reading_mode": "manga",
            "classification_source": "comic-suffix",
            "classification_confidence": 0.99,
            "classification_reason": f"源文件扩展名为 {suffix}",
        }

    tags = str(
        row.get("metadata_tags")
        or row.get("tags")
        or row.get("categories")
        or ""
    ).casefold()
    if tags and re.search(
        r"(^|[;,\s])(comic|comics|manga|graphic novel|graphic novels)([;,\s]|$)",
        tags,
    ):
        return {
            "reading_mode": "manga",
            "classification_source": "calibre-tags",
            "classification_confidence": 0.95,
            "classification_reason": "Calibre 标签包含漫画/Comics",
        }

    known_word_values = [
        _row_number(row, key)
        for key in ("book_words_known", "book_words", "total_words")
    ]
    known_words = max(
        (value for value in known_word_values if value is not None),
        default=0.0,
    )
    if known_words >= 100:
        return {
            "reading_mode": "book",
            "classification_source": "sidecar-words",
            "classification_confidence": 0.9,
            "classification_reason": "sidecar 已报告足量词数",
        }
    return None


def _classification_from_pages(
    page_results: Sequence[OcrPageResult],
) -> Dict[str, Any]:
    char_values = [
        result.chars_per_page
        if result.chars_per_page is not None
        else _count_chars(result.text)
        for result in page_results
    ]
    area_values = [
        result.text_area_ratio
        for result in page_results
        if result.text_area_ratio is not None
    ]
    if not char_values:
        return {
            "reading_mode": None,
            "classification_source": "ocr-sample",
            "classification_confidence": 0.0,
            "classification_reason": "抽样页没有 OCR 文本",
        }
    ordered = sorted(char_values)
    median_chars = ordered[len(ordered) // 2]
    median_area = (
        sorted(area_values)[len(area_values) // 2]
        if area_values
        else None
    )
    if median_chars >= 350:
        confidence = min(0.99, 0.6 + median_chars / 2500.0)
        return {
            "reading_mode": "book",
            "classification_source": "ocr-sample",
            "classification_confidence": round(confidence, 3),
            "classification_reason": (
                f"抽样中位字符数 {median_chars}/页，判定为文字书"
            ),
        }
    if median_chars <= 180 and (
        median_area is None or median_area <= 0.18
    ):
        confidence = min(
            0.98,
            0.62 + max(0.0, (180 - median_chars) / 500.0),
        )
        return {
            "reading_mode": "manga",
            "classification_source": "ocr-sample",
            "classification_confidence": round(confidence, 3),
            "classification_reason": (
                f"抽样中位字符数 {median_chars}/页"
                + (
                    f"，文本块面积 {median_area:.3f}"
                    if median_area is not None
                    else ""
                )
                + "，判定为漫画"
            ),
        }
    if median_chars >= 200 and median_area is not None and median_area >= 0.12:
        return {
            "reading_mode": "book",
            "classification_source": "ocr-sample",
            "classification_confidence": 0.75,
            "classification_reason": (
                f"抽样字符 {median_chars}/页，文本块面积 {median_area:.3f}"
            ),
        }
    return {
        "reading_mode": None,
        "classification_source": "ocr-sample",
        "classification_confidence": 0.45,
        "classification_reason": (
            f"抽样字符 {median_chars}/页"
            + (
                f"，文本块面积 {median_area:.3f}"
                if median_area is not None
                else ""
            )
            + "，无法可靠判定"
        ),
    }


def classify_manga_candidate(
    row: Dict[str, Any],
    config: OcrConfig,
    *,
    source_resolver: Callable[
        [Dict[str, Any], Optional[Path]], Optional[Path]
    ] = resolve_ocr_source,
    sampler: Callable[..., RenderedBook] = render_manga_sample_pages,
    engine_factory: Callable[[str, str], Any] = create_ocr_engine,
    logger: Optional[logging.Logger] = None,
) -> Dict[str, Any]:
    """Classify one uncertain row, falling back to an explicit unknown."""

    config = config.normalized()
    source_path = source_resolver(row, None)
    cheap = _cheap_classification(row, source_path)
    if cheap is not None:
        return cheap
    if source_path is None:
        return {
            "reading_mode": None,
            "classification_source": "missing-source",
            "classification_confidence": 0.0,
            "classification_reason": "找不到可用于抽样分类的源文件",
        }

    language = resolve_ocr_language(row, config.language)
    if language is None:
        return {
            "reading_mode": None,
            "classification_source": "needs-language",
            "classification_confidence": 0.0,
            "classification_reason": "无法安全推断抽样 OCR 语言",
        }
    engine_name = resolve_ocr_engine(config.engine, language)
    available, dependency_message = engine_available(engine_name)
    if not available:
        return {
            "reading_mode": None,
            "classification_source": "engine-missing",
            "classification_confidence": 0.0,
            "classification_reason": dependency_message,
        }

    source_hash = source_file_hash(source_path)
    try:
        rendered = sampler(
            source_path,
            config.cache_dir,
            min(config.dpi, OCR_SAMPLE_DPI),
            source_hash=source_hash,
            sample_pages=config.sample_pages,
            workers=config.workers,
            logger=logger,
        )
    except Exception as exc:
        return {
            "reading_mode": None,
            "classification_source": "render-failed",
            "classification_confidence": 0.0,
            "classification_reason": f"{type(exc).__name__}: {exc}",
        }
    if not rendered.pages:
        return {
            "reading_mode": None,
            "classification_source": "render-empty",
            "classification_confidence": 0.0,
            "classification_reason": "抽样页未渲染出图像",
        }

    try:
        engine = engine_factory(engine_name, language)
        page_results: List[OcrPageResult] = []
        for page in rendered.pages:
            page_results.append(engine.recognize(page.path, language))
    except Exception as exc:
        return {
            "reading_mode": None,
            "classification_source": "ocr-failed",
            "classification_confidence": 0.0,
            "classification_reason": f"{type(exc).__name__}: {exc}",
        }

    classified = _classification_from_pages(page_results)
    classified["classification_mode"] = "sample"
    classified["classification_sample_pages"] = len(page_results)
    return classified


def series_profile_key(row: Dict[str, Any]) -> str:
    value = str(row.get("series_key") or "").strip()
    if value:
        return _normalized_text(value)
    title = str(row.get("title") or "")
    if kindle_reader is not None:
        return kindle_reader._reading_mode_series_key(title)
    return _normalized_text(title)


def _series_profile_path(cache_dir: Path) -> Path:
    return Path(cache_dir) / "series-profiles.json"


_OCR_WORKER_ENGINE: Any = None


def _init_ocr_worker(engine_name: str, language: str) -> None:
    global _OCR_WORKER_ENGINE
    _OCR_WORKER_ENGINE = create_ocr_engine(engine_name, language)


def _ocr_page_worker(
    image_path: str,
    language: str,
) -> OcrPageResult:
    if _OCR_WORKER_ENGINE is None:
        raise RuntimeError("OCR worker was not initialized")
    return _OCR_WORKER_ENGINE.recognize(Path(image_path), language)


def load_series_profile(
    cache_dir: Path,
    row: Dict[str, Any],
) -> Optional[Dict[str, Any]]:
    key = series_profile_key(row)
    if not key:
        return None
    decoded = _read_json(_series_profile_path(cache_dir))
    if not isinstance(decoded, dict):
        return None
    profile = decoded.get(key)
    return profile if isinstance(profile, dict) else None


def save_series_profile(
    cache_dir: Path,
    row: Dict[str, Any],
    *,
    words_per_page: float,
    language: str,
    engine: str,
    confidence: Optional[float],
) -> None:
    key = series_profile_key(row)
    if not key or words_per_page <= 0:
        return
    path = _series_profile_path(cache_dir)
    decoded = _read_json(path)
    profiles = decoded if isinstance(decoded, dict) else {}
    profiles[key] = {
        "words_per_page": round(float(words_per_page), 3),
        "language": language,
        "engine": engine,
        "confidence": (
            round(float(confidence), 4)
            if confidence is not None
            else None
        ),
        "updated_at": datetime.datetime.now(
            datetime.timezone.utc
        ).isoformat(timespec="seconds"),
    }
    _atomic_write_json(path, profiles)


def estimate_manga_ocr_words(
    row: Dict[str, Any],
    config: OcrConfig,
    *,
    source_resolver: Callable[
        [Dict[str, Any], Optional[Path]], Optional[Path]
    ] = resolve_ocr_source,
    renderer: Callable[..., RenderedBook] = render_manga_pages,
    engine_factory: Callable[[str, str], Any] = create_ocr_engine,
    logger: Optional[logging.Logger] = None,
) -> Dict[str, Any]:
    """Estimate manga word/character counts with an injected OCR adapter."""

    started = time.perf_counter()
    tracing_started = not tracemalloc.is_tracing()
    if tracing_started:
        tracemalloc.start()
    try:
        result = _estimate_manga_ocr_words_impl(
            row,
            config.normalized(),
            source_resolver=source_resolver,
            renderer=renderer,
            engine_factory=engine_factory,
            logger=logger,
        )
        current, peak = tracemalloc.get_traced_memory()
        del current
        result["ocr_elapsed_seconds"] = round(
            max(0.0, time.perf_counter() - started),
            3,
        )
        result["ocr_peak_memory_mb"] = round(peak / (1024.0 * 1024.0), 3)
        return result
    finally:
        if tracing_started:
            tracemalloc.stop()


def _estimate_manga_ocr_words_impl(
    row: Dict[str, Any],
    config: OcrConfig,
    *,
    source_resolver: Callable[
        [Dict[str, Any], Optional[Path]], Optional[Path]
    ],
    renderer: Callable[..., RenderedBook],
    engine_factory: Callable[[str, str], Any],
    logger: Optional[logging.Logger],
) -> Dict[str, Any]:
    mode = str(row.get("reading_mode") or "").strip().casefold()
    if mode != "manga":
        return _empty_result(
            "not-applicable",
            "OCR estimates are only generated for reading_mode=manga.",
            min_confidence=config.min_confidence,
        )

    source_path = source_resolver(row, None)
    source_hash = ""
    if source_path is not None:
        try:
            source_hash = source_file_hash(source_path)
        except OSError as exc:
            return _empty_result(
                "source-unreadable",
                f"Could not read the manga source file: {exc}",
                min_confidence=config.min_confidence,
            )
    if source_path is None:
        return _empty_result(
            "source-missing",
            "The local manga ebook file was not found.",
            source_hash=source_hash,
            min_confidence=config.min_confidence,
        )

    language = resolve_ocr_language(row, config.language)
    if language is None:
        return _empty_result(
            "needs-language",
            "OCR language could not be inferred safely. Pass "
            "--ocr-language ja, zh, or en.",
            engine=str(config.engine or ""),
            cache_key=ocr_cache_key(
                source_hash,
                config.dpi,
                "auto",
                str(config.engine or ""),
                "unresolved",
                page_hash="language",
            ),
            source_hash=source_hash,
            min_confidence=config.min_confidence,
        )

    engine_name = resolve_ocr_engine(config.engine, language)
    available, dependency_message = engine_available(engine_name)
    if not available:
        return _empty_result(
            "engine-missing",
            dependency_message,
            language=language,
            engine=engine_name,
            cache_key=ocr_cache_key(
                source_hash,
                config.dpi,
                language,
                engine_name,
                "missing",
                page_hash="dependency",
            ),
            source_hash=source_hash,
            min_confidence=config.min_confidence,
        )

    cache_dir = Path(config.cache_dir).expanduser()
    if config.mode == "sample":
        total_pages, _, page_error = resolve_manga_page_count(
            source_path,
            cache_dir,
            source_hash=source_hash,
            logger=logger,
        )
        profile = load_series_profile(cache_dir, row)
        if total_pages and isinstance(profile, dict):
            words_per_page = _parse_float(profile.get("words_per_page"))
            if words_per_page is not None and words_per_page > 0:
                word_count = int(round(words_per_page * total_pages))
                confidence = _parse_float(profile.get("confidence"))
                return {
                    "ocr_mode": "sample",
                    "ocr_word_count": word_count,
                    "ocr_word_count_is_estimate": True,
                    "ocr_char_count": None,
                    "ocr_confidence": confidence,
                    "ocr_engine": str(profile.get("engine") or engine_name),
                    "ocr_language": str(profile.get("language") or language),
                    "ocr_page_coverage": round(
                        min(100.0, config.sample_pages / total_pages * 100.0),
                        3,
                    ),
                    "ocr_status": "ok",
                    "ocr_error": None,
                    "ocr_completed_at": datetime.datetime.now(
                        datetime.timezone.utc
                    ).isoformat(timespec="seconds"),
                    "ocr_cache_key": "",
                    "ocr_source": OCR_SERIES_SOURCE,
                    "ocr_pages_total": total_pages,
                    "ocr_pages_processed": 0,
                    "ocr_failed_pages": 0,
                    "ocr_elapsed_seconds": None,
                    "ocr_peak_memory_mb": None,
                    "ocr_min_confidence": round(config.min_confidence, 4),
                    "ocr_sampled_pages": 0,
                    "ocr_extrapolated": True,
                    "classification_source": "",
                    "classification_confidence": None,
                    "classification_reason": "复用同系列词/页档案",
                    "manga_pages_total": total_pages,
                    "manga_pages_read": None,
                    "manga_pages_source": "pdfinfo-source-file",
                    "_ocr_source_hash": source_hash,
                }
        if total_pages is None and page_error:
            return _empty_result(
                "source-page-count-unavailable",
                page_error,
                language=language,
                engine=engine_name,
                source_hash=source_hash,
                min_confidence=config.min_confidence,
            )
    try:
        if config.mode == "sample" and renderer is render_manga_pages:
            rendered = render_manga_sample_pages(
                source_path,
                cache_dir,
                min(config.dpi, OCR_SAMPLE_DPI),
                source_hash=source_hash,
                sample_pages=config.sample_pages,
                workers=config.workers,
                logger=logger,
            )
        elif config.mode == "sample":
            rendered = renderer(
                source_path,
                cache_dir,
                min(config.dpi, OCR_SAMPLE_DPI),
                source_hash=source_hash,
                sample_pages=config.sample_pages,
                workers=config.workers,
                logger=logger,
            )
        else:
            rendered = renderer(
                source_path,
                cache_dir,
                config.dpi,
                source_hash=source_hash,
                max_pages=config.max_pages,
                workers=config.workers,
                logger=logger,
            )
    except OcrDependencyError as exc:
        return _empty_result(
            "render-missing",
            str(exc),
            language=language,
            engine=engine_name,
            source_hash=source_hash,
            min_confidence=config.min_confidence,
        )
    except Exception as exc:
        return _empty_result(
            "render-failed",
            f"{type(exc).__name__}: {exc}",
            language=language,
            engine=engine_name,
            source_hash=source_hash,
            min_confidence=config.min_confidence,
        )

    pages = list(rendered.pages)
    pages_total = max(int(rendered.expected_pages or 0), len(pages))
    errors: List[str] = list(rendered.errors)
    if not pages:
        return _empty_result(
            "render-failed",
            "No manga page image could be rendered.",
            language=language,
            engine=engine_name,
            source_hash=source_hash,
            pages_total=pages_total,
            page_errors=errors,
            min_confidence=config.min_confidence,
        )

    engine_version = ocr_engine_version(engine_name)
    book_key = _book_page_key(
        source_hash,
        config.dpi,
        language,
        engine_name,
        engine_version,
        pages,
    )
    cached_result = _read_json(_book_cache_path(cache_dir, book_key))
    if isinstance(cached_result, dict) and cached_result.get("ocr_status"):
        cached_result.setdefault(
            "ocr_min_confidence",
            round(config.min_confidence, 4),
        )
        cached_result["_ocr_cache_hit"] = True
        cached_result["_ocr_source_hash"] = source_hash
        cached_result["_ocr_page_cache_hits"] = len(pages)
        cached_result["_ocr_page_cache_misses"] = 0
        return cached_result

    use_process_pool = bool(
        config.workers > 1
        and len(pages) > 1
        and engine_factory is create_ocr_engine
    )
    engine: Any = None
    if use_process_pool:
        actual_version = engine_version
    else:
        try:
            engine = engine_factory(engine_name, language)
            actual_version = str(
                getattr(engine, "version", "") or engine_version or "unknown"
            )
        except OcrDependencyError as exc:
            return _empty_result(
                "engine-missing",
                str(exc),
                language=language,
                engine=engine_name,
                source_hash=source_hash,
                pages_total=pages_total,
                page_errors=errors,
                min_confidence=config.min_confidence,
            )
        except Exception as exc:
            return _empty_result(
                "engine-failed",
                f"{type(exc).__name__}: {exc}",
                language=language,
                engine=engine_name,
                source_hash=source_hash,
                pages_total=pages_total,
                page_errors=errors,
                min_confidence=config.min_confidence,
            )

    if actual_version and actual_version != engine_version:
        engine_version = actual_version
        book_key = _book_page_key(
            source_hash,
            config.dpi,
            language,
            engine_name,
            engine_version,
            pages,
        )
        cached_result = _read_json(_book_cache_path(cache_dir, book_key))
        if isinstance(cached_result, dict) and cached_result.get("ocr_status"):
            cached_result.setdefault(
                "ocr_min_confidence",
                round(config.min_confidence, 4),
            )
            cached_result["_ocr_cache_hit"] = True
            cached_result["_ocr_source_hash"] = source_hash
            cached_result["_ocr_page_cache_hits"] = len(pages)
            cached_result["_ocr_page_cache_misses"] = 0
            return cached_result

    texts: List[str] = []
    confidences: List[Tuple[float, int]] = []
    processed = 0
    page_cache_hits = 0
    page_cache_misses = 0
    fresh_results: Dict[int, OcrPageResult] = {}
    pages_to_recognize = []
    for page in pages:
        page_key = ocr_cache_key(
            source_hash,
            config.dpi,
            language,
            engine_name,
            engine_version,
            page_hash=page.sha256,
        )
        page_result = _read_page_cache(cache_dir, page_key)
        if page_result is not None:
            fresh_results[page.index] = page_result
            page_cache_hits += 1
        else:
            page_cache_misses += 1
            pages_to_recognize.append(page)

    if pages_to_recognize:
        if use_process_pool:
            workers = max(
                1,
                min(config.workers, len(pages_to_recognize)),
            )
            try:
                with concurrent.futures.ProcessPoolExecutor(
                    max_workers=workers,
                    initializer=_init_ocr_worker,
                    initargs=(engine_name, language),
                ) as executor:
                    future_map = {
                        executor.submit(
                            _ocr_page_worker,
                            str(page.path),
                            language,
                        ): page
                        for page in pages_to_recognize
                    }
                    for future in concurrent.futures.as_completed(future_map):
                        page = future_map[future]
                        try:
                            fresh_results[page.index] = future.result()
                        except Exception as exc:
                            errors.append(
                                f"page {page.index}: "
                                f"{type(exc).__name__}: {exc}"
                            )
            except Exception as exc:
                errors.append(
                    f"OCR worker pool failed: {type(exc).__name__}: {exc}"
                )
        else:
            for page in pages_to_recognize:
                try:
                    fresh_results[page.index] = engine.recognize(
                        page.path,
                        language,
                    )
                except Exception as exc:
                    errors.append(
                        f"page {page.index}: {type(exc).__name__}: {exc}"
                    )

    for page in pages:
        page_result = fresh_results.get(page.index)
        if page_result is None:
            continue
        if page in pages_to_recognize:
            page_key = ocr_cache_key(
                source_hash,
                config.dpi,
                language,
                engine_name,
                engine_version,
                page_hash=page.sha256,
            )
            try:
                _write_page_cache(cache_dir, page_key, page_result)
            except OSError as exc:
                errors.append(
                    f"page {page.index}: page cache write failed: {exc}"
                )
        text = str(page_result.text or "")
        if text:
            texts.append(text)
        confidence = _parse_float(page_result.confidence)
        if confidence is not None:
            confidences.append((confidence, max(1, _count_chars(text))))
        processed += 1

    combined_text = "\n".join(texts)
    raw_word_count = _count_words(combined_text, language)
    raw_char_count = _count_chars(combined_text)
    extrapolated = bool(
        config.mode == "sample"
        and processed > 0
        and pages_total > processed
    )
    if extrapolated:
        word_count = int(round(raw_word_count * pages_total / processed))
        char_count = int(round(raw_char_count * pages_total / processed))
    else:
        word_count = raw_word_count
        char_count = raw_char_count
    confidence_value: Optional[float] = None
    if confidences:
        total_weight = sum(weight for _, weight in confidences)
        if total_weight:
            confidence_value = sum(
                value * weight for value, weight in confidences
            ) / total_weight
    coverage = (
        round(processed / pages_total * 100.0, 3) if pages_total else None
    )

    if processed <= 0:
        status = "failed"
    elif errors or (
        config.mode == "full"
        and coverage is not None
        and coverage < OCR_COVERAGE_TARGET_PERCENT
    ):
        status = "partial"
    elif (
        confidence_value is not None
        and confidence_value < config.min_confidence
    ):
        status = "ocr-low-confidence"
    else:
        status = "ok"

    if errors:
        error_text = "; ".join(errors)
    elif status == "ocr-low-confidence":
        error_text = (
            f"Mean confidence {confidence_value:.3f} is below "
            f"{config.min_confidence:.3f}"
        )
    else:
        error_text = None

    result_source = (
        OCR_SAMPLED_SOURCE if config.mode == "sample" else OCR_SOURCE
    )
    result: Dict[str, Any] = {
        "ocr_mode": config.mode,
        "ocr_word_count": word_count,
        "ocr_word_count_is_estimate": bool(extrapolated),
        "ocr_char_count": char_count,
        "ocr_confidence": (
            round(confidence_value, 4)
            if confidence_value is not None
            else None
        ),
        "ocr_engine": engine_name,
        "ocr_language": language,
        "ocr_page_coverage": coverage,
        "ocr_status": status,
        "ocr_error": error_text,
        "ocr_completed_at": datetime.datetime.now(
            datetime.timezone.utc
        ).isoformat(timespec="seconds"),
        "ocr_cache_key": book_key,
        "ocr_source": result_source,
        "ocr_pages_total": pages_total,
        "ocr_pages_processed": processed,
        "ocr_failed_pages": len([error for error in errors if error.startswith("page ")]),
        "ocr_elapsed_seconds": None,
        "ocr_peak_memory_mb": None,
        "ocr_min_confidence": round(config.min_confidence, 4),
        "ocr_sampled_pages": processed if config.mode == "sample" else 0,
        "ocr_extrapolated": extrapolated,
        "classification_source": "",
        "classification_confidence": None,
        "classification_reason": (
            f"按 {processed} 个采样页外推 {pages_total} 页"
            if extrapolated
            else ""
        ),
        "manga_pages_total": pages_total or None,
        "manga_pages_read": None,
        "manga_pages_source": "ocr",
        "_ocr_page_cache_hits": page_cache_hits,
        "_ocr_page_cache_misses": page_cache_misses,
        "_ocr_source_hash": source_hash,
    }
    if (
        config.mode == "sample"
        and processed > 0
        and raw_word_count > 0
        and extrapolated
    ):
        try:
            save_series_profile(
                cache_dir,
                row,
                words_per_page=raw_word_count / float(processed),
                language=language,
                engine=engine_name,
                confidence=confidence_value,
            )
        except OSError:
            pass
    if status in {"ok", "ocr-low-confidence", "partial"}:
        try:
            _atomic_write_json(
                _book_cache_path(cache_dir, book_key),
                _result_cache_payload(result),
            )
        except OSError:
            pass
    return result


def merge_ocr_estimate(
    row: Dict[str, Any],
    estimate: Optional[Dict[str, Any]],
) -> Dict[str, Any]:
    """Return a copy with OCR fields appended; never touch exact word counts."""

    merged = dict(row)
    if not isinstance(estimate, dict):
        return merged
    for field in OCR_FIELDS:
        merged[field] = estimate.get(field)
    return merged


def _write_csv_rows(
    path: Path,
    rows: Sequence[Dict[str, Any]],
    fieldnames: Sequence[str],
) -> None:
    buffer = io.StringIO(newline="")
    writer = csv.DictWriter(buffer, fieldnames=list(fieldnames))
    writer.writeheader()
    for row in rows:
        writer.writerow({field: row.get(field) for field in fieldnames})
    temporary_path: Optional[Path] = None
    try:
        with tempfile.NamedTemporaryFile(
            "w",
            encoding="utf-8-sig",
            newline="",
            dir=path.parent,
            prefix=path.name + ".",
            suffix=".tmp",
            delete=False,
        ) as handle:
            temporary_path = Path(handle.name)
            handle.write(buffer.getvalue())
        os.replace(
            _filesystem_path(temporary_path),
            _filesystem_path(path),
        )
    finally:
        if temporary_path is not None and os.path.exists(
            _filesystem_path(temporary_path)
        ):
            os.unlink(_filesystem_path(temporary_path))


def enrich_csv_with_manga_ocr(
    csv_path: Path,
    config: OcrConfig,
    *,
    source_resolver: Callable[
        [Dict[str, Any], Optional[Path]], Optional[Path]
    ] = resolve_ocr_source,
    renderer: Callable[..., RenderedBook] = render_manga_pages,
    engine_factory: Callable[[str, str], Any] = create_ocr_engine,
    logger: Optional[logging.Logger] = None,
    progress_callback: Optional[Callable[[Dict[str, Any]], None]] = None,
) -> Dict[str, Any]:
    """Append OCR fields to a CSV while preserving every existing field."""

    csv_path = Path(csv_path)
    with open(_filesystem_path(csv_path), "r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        fieldnames = list(reader.fieldnames or [])
        rows = [dict(row) for row in reader]
    if not fieldnames:
        raise ValueError("CSV has no header")

    output_fields = list(fieldnames)
    output_fields.extend(
        field for field in OCR_FIELDS if field not in output_fields
    )
    counts: Dict[str, int] = {
        "rows": len(rows),
        "manga_rows": 0,
        "ocr_ok": 0,
        "ocr_non_ok": 0,
        "ocr_cache_hits": 0,
    }
    manga_total = sum(
        1
        for row in rows
        if str(row.get("reading_mode") or "").casefold() == "manga"
    )
    manga_index = 0
    for index, row in enumerate(rows, start=1):
        if str(row.get("reading_mode") or "").casefold() != "manga":
            continue
        manga_index += 1
        counts["manga_rows"] += 1
        try:
            estimate = estimate_manga_ocr_words(
                row,
                config,
                source_resolver=source_resolver,
                renderer=renderer,
                engine_factory=engine_factory,
                logger=logger,
            )
        except Exception as exc:
            estimate = _empty_result(
                "ocr-unhandled",
                f"{type(exc).__name__}: {exc}",
                language=str(config.language or "auto"),
                engine=str(config.engine or "auto"),
                min_confidence=config.min_confidence,
            )
        rows[index - 1] = merge_ocr_estimate(row, estimate)
        if estimate.get("ocr_status") == "ok":
            counts["ocr_ok"] += 1
        else:
            counts["ocr_non_ok"] += 1
        if estimate.get("_ocr_cache_hit"):
            counts["ocr_cache_hits"] += 1
        if progress_callback is not None:
            progress_callback(
                {
                    **estimate,
                    "title": str(row.get("title") or "(unknown)"),
                    "manga_index": manga_index,
                    "manga_total": manga_total,
                }
            )
        if logger is not None:
            logger.info(
                "Manga OCR %d/%d: %s | pages %s/%s | cache hit=%s",
                manga_index,
                manga_total,
                str(row.get("title") or "(unknown)"),
                estimate.get("ocr_pages_processed") or 0,
                estimate.get("ocr_pages_total") or 0,
                bool(estimate.get("_ocr_cache_hit")),
            )
    _write_csv_rows(csv_path, rows, output_fields)
    counts["output_fields"] = len(output_fields)
    return counts


def dependency_report(
    engine: str,
    language: str,
) -> List[str]:
    """Return explicit dependency guidance without installing anything."""

    requested_engine = str(engine or "auto").casefold()
    requested_language = str(language or "auto").casefold()
    if requested_engine == "auto" and requested_language == "auto":
        if any(
            engine_available(name)[0]
            for name in ("manga_ocr", "paddleocr", "tesseract")
        ):
            lines = []
        else:
            lines = [
                DEPENDENCY_HINTS["manga_ocr"],
                DEPENDENCY_HINTS["paddleocr"],
                DEPENDENCY_HINTS["tesseract"],
            ]
        if find_calibre_convert() is None:
            lines.append(
                "Calibre ebook-convert is required for non-PDF manga sources."
            )
        if find_pdftoppm() is None:
            lines.append(
                "Poppler pdftoppm is required to render PDF pages."
            )
        return lines

    resolved_engine = (
        resolve_ocr_engine(engine, language)
        if requested_engine == "auto"
        else str(engine or "").casefold()
    )
    lines: List[str] = []
    available, message = engine_available(resolved_engine)
    if not available:
        lines.append(message)
    if find_calibre_convert() is None:
        lines.append(
            "Calibre ebook-convert is required for non-PDF manga sources."
        )
    if find_pdftoppm() is None:
        lines.append("Poppler pdftoppm is required to render PDF pages.")
    return lines
