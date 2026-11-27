from __future__ import annotations

import json
import struct
import sys
import tempfile
import unittest
from pathlib import Path


SKILL_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_DIR = SKILL_ROOT / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

import extract_kindle_stats as ek  # noqa: E402


def tagged(type_byte: int, payload: bytes) -> bytes:
    return struct.pack(">b", type_byte) + payload


def krds_long(value: int) -> bytes:
    return tagged(2, struct.pack(">q", value))


def krds_int(value: int) -> bytes:
    return tagged(1, struct.pack(">i", value))


def krds_bool(value: bool) -> bytes:
    return tagged(0, bytes([1 if value else 0]))


def krds_double(value: float) -> bytes:
    return tagged(4, struct.pack(">d", value))


def utf_payload(value: str) -> bytes:
    encoded = value.encode("utf-8")
    return b"\x00" + struct.pack(">H", len(encoded)) + encoded


def krds_utf(value: str) -> bytes:
    return tagged(3, utf_payload(value))


def krds_object(name: str, values: list[bytes]) -> bytes:
    return (
        struct.pack(">b", -2)
        + utf_payload(name)
        + b"".join(values)
        + struct.pack(">b", -1)
    )


def synthetic_krds(
    total_time_ms: int = 600000,
    total_words: int = 3000,
    timer_percent: float = 0.25,
    book_words: int = 9000,
) -> bytes:
    average_calculator = krds_object(
        "timer.average.calculator",
        [
            krds_int(0),
            krds_int(0),
            krds_int(0),
            krds_int(0),
        ],
    )
    timer_model = krds_object(
        "timer.model",
        [
            krds_long(1),
            krds_long(total_time_ms),
            krds_long(total_words),
            krds_double(timer_percent),
            average_calculator,
        ],
    )
    timer_store = krds_object(
        "timer.data.store.v2",
        [
            krds_bool(True),
            timer_model,
            krds_long(1),
            krds_int(0),
        ],
    )
    book_info = krds_object(
        "book.info.store",
        [
            krds_long(book_words),
            krds_double(1.0),
        ],
    )
    return (
        ek.KRDS_SIGNATURE
        + krds_long(1)
        + krds_int(2)
        + timer_store
        + book_info
    )


def synthetic_mobi(title: str = "Synthetic Book", author: str = "Test Author") -> bytes:
    exth_records = []
    for record_type, value in ((100, author), (503, title)):
        payload = value.encode("utf-8")
        exth_records.append(struct.pack(">II", record_type, len(payload) + 8) + payload)
    exth_payload = b"".join(exth_records)
    exth = b"EXTH" + struct.pack(">II", 12 + len(exth_payload), 2) + exth_payload

    mobi_header_length = 232
    palm_doc = bytearray(16)
    mobi_header = bytearray(mobi_header_length)
    mobi_header[0:4] = b"MOBI"
    struct.pack_into(">I", mobi_header, 4, mobi_header_length)
    struct.pack_into(">I", mobi_header, 12, 65001)
    struct.pack_into(">I", mobi_header, 112, 0x40)
    record0 = bytes(palm_doc + mobi_header) + exth

    first_offset = 78 + 16
    second_offset = first_offset + len(record0)
    header = bytearray(78)
    header[0:32] = title.encode("utf-8")[:32].ljust(32, b"\x00")
    struct.pack_into(">H", header, 76, 2)
    offsets = (
        struct.pack(">I", first_offset)
        + b"\x00\x00\x00\x00"
        + struct.pack(">I", second_offset)
        + b"\x00\x00\x00\x00"
    )
    return bytes(header) + offsets + record0


class KrdsParserTests(unittest.TestCase):
    def test_nested_timer_and_book_info_are_extracted(self) -> None:
        document = ek.KrdsDecoder(synthetic_krds()).decode_document()
        stats = ek.extract_reading_stats(document)

        self.assertEqual(stats.total_time_ms, 600000)
        self.assertEqual(stats.total_words, 3000)
        self.assertEqual(stats.book_words, 9000)
        self.assertEqual(ek.calculate_wpm(stats), 300.0)

        progress, source, timer_percent, word_progress = ek.calculate_progress(
            stats
        )
        self.assertEqual(progress, 25.0)
        self.assertEqual(timer_percent, 25.0)
        self.assertEqual(word_progress, 33.333333)
        self.assertEqual(source, "timer.model.totalPercent")

    def test_truncated_file_keeps_completed_records(self) -> None:
        data = synthetic_krds()
        document = ek.KrdsDecoder(data[:-5]).decode_document()

        self.assertTrue(document.partial)
        stats = ek.extract_reading_stats(document)
        self.assertEqual(stats.total_time_ms, 600000)
        self.assertTrue(document.warnings)

    def test_invalid_signature_is_rejected(self) -> None:
        with self.assertRaises(ek.KrdsFormatError):
            ek.KrdsDecoder(b"not a krds file").decode_document()

    def test_mobi_metadata(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "book.azw3"
            path.write_bytes(synthetic_mobi())
            title, author = ek.parse_mobi_metadata(path)

        self.assertEqual(title, "Synthetic Book")
        self.assertEqual(author, "Test Author")

    def test_end_to_end_exports(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "Kindle"
            documents = root / "documents"
            sidecar_dir = documents / "The Hobbit - J. R. R. Tolkien.sdr"
            sidecar_dir.mkdir(parents=True)
            (documents / "The Hobbit - J. R. R. Tolkien.azw3").write_bytes(
                b"not real mobi metadata"
            )
            (sidecar_dir / "The Hobbit - J. R. R. Tolkien.yjf").write_bytes(
                synthetic_krds()
            )
            output_dir = Path(temp_dir) / "output"

            exit_code = ek.main(
                [
                    "--kindle-path",
                    str(root),
                    "--output-dir",
                    str(output_dir),
                    "--format",
                    "all",
                    "--quiet",
                ]
            )

            self.assertEqual(exit_code, 0)
            csv_text = (output_dir / "kindle-reading-stats.csv").read_text(
                encoding="utf-8-sig"
            )
            json_data = json.loads(
                (output_dir / "kindle-reading-stats.json").read_text(
                    encoding="utf-8"
                )
            )
            html_text = (output_dir / "kindle-reading-stats.html").read_text(
                encoding="utf-8"
            )

        self.assertIn("The Hobbit", csv_text)
        self.assertIn("J. R. R. Tolkien", csv_text)
        self.assertIn("300.0", csv_text)
        self.assertEqual(json_data["summary"]["books"], 1)
        self.assertEqual(json_data["books"][0]["total_time_ms"], 600000)
        self.assertEqual(json_data["books"][0]["progress_percent"], 25.0)
        self.assertIn("Kindle 阅读统计", html_text)
        self.assertIn("The Hobbit", html_text)


if __name__ == "__main__":
    unittest.main()
