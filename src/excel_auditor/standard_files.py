from __future__ import annotations

import csv
from pathlib import Path
from typing import Any, BinaryIO, Sequence

import ijson
from ijson.common import ObjectBuilder

from .models import RuleSet, SheetRule
from .snapshots import SpilledRecords


class _RecordAccumulator:
    """Canonicalize records and migrate payloads to disk past a threshold."""

    def __init__(self, sheet: SheetRule, spill_after_records: int) -> None:
        if spill_after_records < 1:
            raise ValueError("spill_after_records must be positive")
        self.sheet = sheet
        self.spill_after_records = spill_after_records
        self.records: list[dict[str, Any]] | SpilledRecords = []
        self.detached = False

    def append(self, row: Any) -> None:
        if not isinstance(row, dict):
            raise ValueError(f"STANDARD_DATA_INVALID: {self.sheet.id} must be an array of objects")
        canonical: dict[str, Any] = {}
        for column in self.sheet.columns:
            matched = next(
                (candidate for candidate in [column.name, column.title, *column.aliases] if candidate in row),
                None,
            )
            if matched is not None:
                canonical[column.name] = row[matched]
        if isinstance(self.records, list) and len(self.records) >= self.spill_after_records:
            spilled = SpilledRecords()
            try:
                for existing in self.records:
                    spilled.append(existing)
            except Exception:
                spilled.close()
                raise
            self.records = spilled
        self.records.append(canonical)

    def detach(self) -> Sequence[dict[str, Any]]:
        self.detached = True
        return self.records

    def close(self) -> None:
        if self.detached:
            return
        close = getattr(self.records, "close", None)
        if close is not None:
            close()


def load_standard_file(
    path: Path,
    rules: RuleSet,
    *,
    spill_after_records: int,
) -> dict[str, Sequence[dict[str, Any]]]:
    """Load uploaded JSON/CSV one record at a time with bounded payload memory."""
    if path.suffix.lower() == ".csv":
        return _load_csv(path, rules, spill_after_records)
    if path.suffix.lower() != ".json":
        raise ValueError("STANDARD_DATA_INVALID: only JSON and CSV inputs are supported")
    return _load_json(path, rules, spill_after_records)


def _load_csv(path: Path, rules: RuleSet, spill_after_records: int) -> dict[str, Sequence[dict[str, Any]]]:
    if len(rules.sheets) != 1:
        raise ValueError("STANDARD_DATA_INVALID: CSV input requires exactly one configured sheet")
    sheet = rules.sheets[0]
    accumulator = _RecordAccumulator(sheet, spill_after_records)
    count = 0
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            for row in csv.DictReader(handle):
                count = _append(accumulator, row, count, rules)
        return {sheet.id: accumulator.detach()}
    except (UnicodeDecodeError, csv.Error) as exc:
        raise ValueError("STANDARD_DATA_INVALID: uploaded CSV is malformed") from exc
    finally:
        accumulator.close()


def _load_json(path: Path, rules: RuleSet, spill_after_records: int) -> dict[str, Sequence[dict[str, Any]]]:
    accumulators: list[_RecordAccumulator] = []
    result: dict[str, Sequence[dict[str, Any]]] = {}
    try:
        with path.open("rb") as handle:
            marker = _prepare_json_stream(handle)
            if marker == b"[":
                if len(rules.sheets) != 1:
                    raise ValueError("STANDARD_DATA_INVALID: list input requires exactly one configured sheet")
                sheet = rules.sheets[0]
                accumulator = _RecordAccumulator(sheet, spill_after_records)
                accumulators.append(accumulator)
                count = 0
                for row in ijson.items(handle, "item", use_float=True):
                    count = _append(accumulator, row, count, rules)
                result[sheet.id] = accumulator.detach()
                return result
            if marker != b"{":
                raise ValueError("STANDARD_DATA_INVALID: root must be an object or array")
            parser = iter(ijson.parse(handle, use_float=True))
            _expect(next(parser), "start_map", "STANDARD_DATA_INVALID: root must be an object or array")
            count = 0
            seen_input_keys: set[str] = set()
            seen_sheet_ids: set[str] = set()
            by_key = {
                key: sheet
                for sheet in rules.sheets
                for key in (sheet.id, sheet.name)
            }
            while True:
                _prefix, event, value = next(parser)
                if event == "end_map":
                    break
                if event != "map_key":
                    raise ValueError("STANDARD_DATA_INVALID: root object values must be record arrays")
                key = str(value)
                if key in seen_input_keys:
                    raise ValueError(f"STANDARD_DATA_INVALID: duplicate sheet key: {key}")
                seen_input_keys.add(key)
                sheet = by_key.get(key)
                if sheet is None:
                    raise ValueError(f"STANDARD_DATA_INVALID: unknown sheet key: {key}")
                if sheet.id in seen_sheet_ids:
                    raise ValueError(f"STANDARD_DATA_INVALID: duplicate sheet mapping: {sheet.id}")
                seen_sheet_ids.add(sheet.id)
                _expect(next(parser), "start_array", f"STANDARD_DATA_INVALID: {key} must be an array of objects")
                accumulator = _RecordAccumulator(sheet, spill_after_records)
                accumulators.append(accumulator)
                while True:
                    current = next(parser)
                    if current[1] == "end_array":
                        break
                    if current[1] != "start_map":
                        raise ValueError(f"STANDARD_DATA_INVALID: {key} must be an array of objects")
                    row = _build_value(parser, current)
                    count = _append(accumulator, row, count, rules)
                result[sheet.id] = accumulator.detach()
            try:
                next(parser)
            except StopIteration:
                return result
            raise ValueError("STANDARD_DATA_INVALID: trailing JSON content")
    except (ijson.JSONError, UnicodeDecodeError, StopIteration) as exc:
        _close_sequences(result)
        raise ValueError("STANDARD_DATA_INVALID: uploaded JSON is malformed") from exc
    except Exception:
        _close_sequences(result)
        raise
    finally:
        for accumulator in accumulators:
            accumulator.close()


def _prepare_json_stream(handle: BinaryIO) -> bytes:
    if handle.read(3) != b"\xef\xbb\xbf":
        handle.seek(0)
    start = handle.tell()
    while byte := handle.read(1):
        if byte not in b" \t\r\n":
            handle.seek(start)
            return byte
    handle.seek(start)
    return b""


def _expect(item: tuple[str, str, Any], event: str, message: str) -> None:
    if item[1] != event:
        raise ValueError(message)


def _build_value(parser: Any, first: tuple[str, str, Any]) -> Any:
    builder = ObjectBuilder()
    depth = 0
    current = first
    while True:
        _prefix, event, value = current
        builder.event(event, value)
        if event in {"start_map", "start_array"}:
            depth += 1
        elif event in {"end_map", "end_array"}:
            depth -= 1
            if depth == 0:
                return builder.value
        current = next(parser)


def _append(
    accumulator: _RecordAccumulator,
    row: Any,
    count: int,
    rules: RuleSet,
) -> int:
    count += 1
    if count > rules.workbook.max_standard_records:
        raise ValueError(
            f"STANDARD_TOO_LARGE: {count} records exceed configured limit "
            f"{rules.workbook.max_standard_records}"
        )
    accumulator.append(row)
    return count


def _close_sequences(standard: dict[str, Sequence[dict[str, Any]]]) -> None:
    for rows in standard.values():
        close = getattr(rows, "close", None)
        if close is not None:
            close()
