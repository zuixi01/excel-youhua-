from __future__ import annotations

import hashlib
import json
import pickle
import tempfile
from array import array
from dataclasses import dataclass
from datetime import date, datetime, timezone
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterable, Sequence, overload

from .ids import new_ulid
from .strict_serialization import dump_json_exact


@dataclass(frozen=True)
class StandardSnapshot:
    snapshot_id: str
    path: Path
    sha256: str
    record_count: int
    fetched_at: datetime
    metadata: dict[str, Any]


class SpilledRecords(Sequence[dict[str, Any]]):
    """Disk-backed standard records with bounded payload memory."""

    def __init__(self) -> None:
        self._file = tempfile.TemporaryFile(prefix="excel-auditor-standard-", suffix=".bin")
        self._offsets = array("Q")

    def append(self, record: dict[str, Any]) -> None:
        self._offsets.append(self._file.tell())
        pickle.dump(record, self._file, protocol=5)

    def __len__(self) -> int:
        return len(self._offsets)

    @overload
    def __getitem__(self, index: int) -> dict[str, Any]: ...

    @overload
    def __getitem__(self, index: slice) -> Iterable[dict[str, Any]]: ...

    def __getitem__(self, index: int | slice) -> dict[str, Any] | Iterable[dict[str, Any]]:
        if isinstance(index, slice):
            return (self[position] for position in range(*index.indices(len(self))))
        if index < 0:
            index += len(self)
        if not 0 <= index < len(self):
            raise IndexError(index)
        self._file.seek(self._offsets[index])
        return pickle.load(self._file)

    def close(self) -> None:
        self._file.close()


def _json_default(value: Any) -> Any:
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, Decimal):
        return str(value)
    raise TypeError(f"not JSON serializable: {type(value).__name__}")


class _ExactDecimalEncodingRequired(TypeError):
    pass


def _fast_snapshot_default(value: Any) -> Any:
    if isinstance(value, Decimal):
        raise _ExactDecimalEncodingRequired
    return _json_default(value)


def _snapshot_json(value: Any) -> str:
    try:
        # Preserve the C encoder's throughput for the common string/integer
        # workload. Decimal-bearing records take the exact numeric path.
        return json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
            allow_nan=False,
            default=_fast_snapshot_default,
        )
    except _ExactDecimalEncodingRequired:
        return dump_json_exact(value, ensure_ascii=False, sort_keys=True, default=_json_default)


def create_snapshot(standard: dict[str, Sequence[dict[str, Any]]], directory: Path, metadata: dict[str, Any] | None = None) -> StandardSnapshot:
    directory.mkdir(parents=True, exist_ok=True)
    snapshot_id = new_ulid("std_")
    count = 0
    path = directory / f"{snapshot_id}.jsonl"
    digest = hashlib.sha256()
    with path.open("wb") as handle:
        for sheet_id in sorted(standard):
            rows = standard[sheet_id]
            # Source record order is part of the snapshot content.  Object keys are
            # canonicalized below, so no full in-memory sort/copy is required.
            for record in rows:
                line = (_snapshot_json({"sheet_id": sheet_id, "record": record}) + "\n").encode("utf-8")
                handle.write(line)
                digest.update(line)
                count += 1
    return StandardSnapshot(snapshot_id, path, digest.hexdigest(), count, datetime.now(timezone.utc), metadata or {})


def load_snapshot(snapshot: StandardSnapshot, spill_after_records: int = 100_000) -> dict[str, Sequence[dict[str, Any]]]:
    if spill_after_records < 1:
        raise ValueError("spill_after_records must be positive")
    spilled = snapshot.record_count > spill_after_records
    result: dict[str, list[dict[str, Any]] | SpilledRecords] = {}
    digest = hashlib.sha256()
    count = 0
    try:
        with snapshot.path.open("rb") as handle:
            for line_number, line in enumerate(handle, start=1):
                digest.update(line)
                try:
                    item = json.loads(line, parse_float=Decimal)
                    sheet_id = str(item["sheet_id"])
                    record = item["record"]
                except (json.JSONDecodeError, KeyError, TypeError, UnicodeDecodeError) as exc:
                    raise ValueError(f"STANDARD_DATA_INVALID: corrupt snapshot line {line_number}") from exc
                if not isinstance(record, dict):
                    raise ValueError(f"STANDARD_DATA_INVALID: corrupt snapshot record at line {line_number}")
                rows = result.get(sheet_id)
                if rows is None:
                    rows = SpilledRecords() if spilled else []
                    result[sheet_id] = rows
                rows.append(record)
                count += 1
        if digest.hexdigest() != snapshot.sha256:
            raise ValueError("STANDARD_DATA_INVALID: snapshot hash mismatch")
        if count != snapshot.record_count:
            raise ValueError("STANDARD_DATA_INVALID: snapshot record count mismatch")
        return result
    except Exception:
        for rows in result.values():
            close = getattr(rows, "close", None)
            if close is not None:
                close()
        raise
