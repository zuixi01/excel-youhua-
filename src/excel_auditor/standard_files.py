from __future__ import annotations

import csv
from pathlib import Path
from typing import Any, BinaryIO, Sequence

import ijson
from ijson.common import ObjectBuilder

from .models import RuleSet, SheetRule, normalize_header
from .normalization import parse_value
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
        canonical = canonicalize_standard_row(row, self.sheet, record_index=len(self.records) + 1)
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


def canonicalize_standard_row(
    row: Any,
    sheet: SheetRule,
    *,
    record_index: int | None = None,
) -> dict[str, Any]:
    """Map source field names without silently choosing conflicting aliases."""
    if not isinstance(row, dict):
        raise ValueError(f"STANDARD_DATA_INVALID: {sheet.id} must be an array of objects")
    if any(not isinstance(key, str) for key in row):
        location = f" at record {record_index}" if record_index is not None else ""
        raise ValueError(f"STANDARD_DATA_INVALID: {sheet.id} has a non-string field name{location}")

    normalized_keys: dict[str, list[str]] = {}
    for raw_key in row:
        normalized_keys.setdefault(normalize_header(raw_key), []).append(raw_key)

    canonical: dict[str, Any] = {}
    for column in sheet.columns:
        candidates = [column.name, column.title, *column.aliases]
        matched_keys: list[str] = []
        for candidate in candidates:
            for raw_key in normalized_keys.get(normalize_header(candidate), []):
                if raw_key not in matched_keys:
                    matched_keys.append(raw_key)
        if not matched_keys:
            continue

        if len(matched_keys) > 1:
            parsed = [parse_value(row[key], column) for key in matched_keys]
            if any(not _same_canonical_value(parsed[0], value) for value in parsed[1:]):
                location = f" at record {record_index}" if record_index is not None else ""
                raise ValueError(
                    f"STANDARD_DATA_INVALID: {sheet.id}.{column.name} has conflicting field representations{location}"
                )
        canonical[column.name] = row[matched_keys[0]]
    if sheet.primary_key_mode == "row_number" and sheet.row_number_field not in canonical:
        matched_row_fields = normalized_keys.get(normalize_header(sheet.row_number_field), [])
        if len(matched_row_fields) > 1:
            location = f" at record {record_index}" if record_index is not None else ""
            raise ValueError(
                f"STANDARD_DATA_INVALID: {sheet.id}.{sheet.row_number_field} has conflicting field representations{location}"
            )
        if matched_row_fields:
            canonical[sheet.row_number_field] = row[matched_row_fields[0]]
    return canonical


def _same_canonical_value(left: Any, right: Any) -> bool:
    if left.valid and right.valid:
        return left.normalized == right.normalized
    if left.valid != right.valid:
        return False
    return type(left.raw) is type(right.raw) and left.raw == right.raw


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
            reader = csv.DictReader(handle)
            fieldnames = reader.fieldnames or []
            if len(fieldnames) != len(set(fieldnames)):
                raise ValueError("STANDARD_DATA_INVALID: uploaded CSV has duplicate field names")
            for row in reader:
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
                for row in ijson.items(handle, "item", use_float=False):
                    count = _append(accumulator, row, count, rules)
                result[sheet.id] = accumulator.detach()
                return result
            if marker != b"{":
                raise ValueError("STANDARD_DATA_INVALID: root must be an object or array")
            parser = iter(ijson.parse(handle, use_float=False))
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
    containers: list[set[str] | None] = []
    current = first
    while True:
        _prefix, event, value = current
        if event == "map_key":
            seen = containers[-1] if containers else None
            if seen is not None:
                key = str(value)
                if key in seen:
                    raise ValueError("STANDARD_DATA_INVALID: uploaded JSON has a duplicate object key")
                seen.add(key)
        builder.event(event, value)
        if event == "start_map":
            containers.append(set())
            depth += 1
        elif event == "start_array":
            containers.append(None)
            depth += 1
        elif event in {"end_map", "end_array"}:
            containers.pop()
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
