# KRDS v1 Format Notes

本文档只记录维护 `extract_kindle_stats.py` 所需的格式信息。KRDS 是 Amazon
未公开的 Kindle Reader Data Store 格式，字段含义可能随固件变化。

## Byte Order And Type Tags

All multi-byte numbers are big-endian. Every value starts with a signed type
tag:

| Tag | Type | Payload |
| --- | --- | --- |
| `0` | Boolean | 1 byte, `0` or `1` |
| `1` | Int32 | 4-byte signed integer |
| `2` | Int64 | 8-byte signed integer |
| `3` | UTF-8 string | empty flag byte, optional uint16 length, bytes |
| `4` | Double | 8-byte IEEE-754 |
| `5` | Int16 | 2-byte signed integer |
| `6` | Float | 4-byte IEEE-754 |
| `7` | Byte | 1-byte signed integer |
| `9` | Char | 1-byte character value |
| `-2` | Object begin | UTF-8 name payload, values, object end |
| `-1` | Object end | no payload |

UTF-8 strings encode a leading byte. `1` means an empty string. `0` means a
uint16 byte length follows, then the UTF-8 bytes. A zero-length string can be
represented with the leading byte set to `0`.

## Store Layout

The file begins with:

```text
00 00 00 00 00 1A B1 26
```

After the signature:

1. An encoded Int64 value equal to `1`.
2. An encoded Int32 object count.
3. That many object-begin frames.

Each top-level frame has an object name and a value list. Object framing is
delimiter-based, so a decoder can retain earlier records when a later record
is truncated.

Records are position-based rather than key-based. The relevant schemas are:

```text
timer.model:
  version: Int64
  total_time: Int64
  total_words: Int64
  total_percent: Double
  average_calculator: Object("timer.average.calculator")

book.info.store:
  number_of_words: Int64
  percent_of_book: Double
```

Older sidecars usually put `timer.model` at the top level. Newer files can
place it inside `timer.data.store` or `timer.data.store.v2` as the
`readingTimerModel` field. The extractor accepts all three locations and
selects the usable candidate with the greatest accumulated time.

## Field Semantics

- `total_time` is milliseconds.
- `total_words` is the reader's accumulated word counter.
- `total_percent` is the reader-reported progress. Some samples use a
  fraction (`0.42`), other code paths expose a percentage (`42.0`). The
  extractor normalizes values at or below `1.0` to a 0-100 percentage.
- `book.info.store.number_of_words` is an estimate of the book length.
- `book.info.store.percent_of_book` is the fraction of the book covered by
  that word estimate. It is not reader progress and must not be used as the
  primary progress value.

## Book Identity

The timer records do not contain reliable title and author fields. Sidecars
live in `Book Name.sdr`. The extractor:

1. Uses the `.sdr` directory stem.
2. Searches for a matching AZW3, MOBI, AWZ, KFX, Topaz, PDF, PRC, or EPUB
   file in the surrounding `documents` directory.
3. Reads MOBI/AZW3 PDB and EXTH metadata when possible.
4. Falls back to `Title - Author` filename parsing and finally to the stem
   as the title.

KFX metadata is a separate container format and is intentionally not parsed
by this skill; filename fallback is used instead.

## References

- John Howell's public parser and format discussion:
  https://www.mobileread.com/forums/showthread.php?t=322172
- Community-maintained Python implementation:
  https://github.com/K-R-D-S/KRDS
- Schema-oriented reader/writer implementation:
  https://github.com/midrare/krdsrw

These links are references for format research. Route new firmware quirks
through the JSON diagnostics before broadening the parser.
