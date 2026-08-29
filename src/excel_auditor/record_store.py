from __future__ import annotations

import os
import pickle
import sqlite3
import tempfile
from array import array
from collections.abc import Iterator, MutableMapping
from pathlib import Path
from typing import Any

from .partitioned_join import stable_key


class DiskBackedRecordMap(MutableMapping[tuple[Any, ...], dict[str, Any]]):
    """SQLite-backed normalized-key map used by large standard datasets."""

    def __init__(self) -> None:
        descriptor, raw_path = tempfile.mkstemp(prefix="excel-auditor-record-map-", suffix=".sqlite3")
        os.close(descriptor)
        self._path = Path(raw_path)
        self._connection = sqlite3.connect(self._path)
        self._connection.execute("PRAGMA journal_mode=OFF")
        self._connection.execute("PRAGMA synchronous=OFF")
        self._connection.execute(
            "CREATE TABLE records (stable_key TEXT PRIMARY KEY, ordinal INTEGER NOT NULL, key_blob BLOB NOT NULL, record_blob BLOB NOT NULL)"
        )
        self._connection.execute("CREATE INDEX records_ordinal ON records (ordinal)")
        self._next_ordinal = 0
        self._join_ordinals = array("Q")
        self._closed = False

    @property
    def path(self) -> Path:
        return self._path

    def __getitem__(self, key: tuple[Any, ...]) -> dict[str, Any]:
        row = self._connection.execute(
            "SELECT record_blob FROM records WHERE stable_key = ?", (stable_key(key),)
        ).fetchone()
        if row is None:
            raise KeyError(key)
        return pickle.loads(row[0])

    def __setitem__(self, key: tuple[Any, ...], value: dict[str, Any]) -> None:
        encoded = stable_key(key)
        try:
            self._connection.execute(
                "INSERT INTO records (stable_key, ordinal, key_blob, record_blob) VALUES (?, ?, ?, ?)",
                (encoded, self._next_ordinal, pickle.dumps(key, protocol=5), pickle.dumps(value, protocol=5)),
            )
        except sqlite3.IntegrityError as exc:
            raise KeyError(key) from exc
        self._next_ordinal += 1
        if self._next_ordinal % 10_000 == 0:
            self._connection.commit()

    def __delitem__(self, key: tuple[Any, ...]) -> None:
        cursor = self._connection.execute("DELETE FROM records WHERE stable_key = ?", (stable_key(key),))
        if cursor.rowcount == 0:
            raise KeyError(key)

    def __iter__(self) -> Iterator[tuple[Any, ...]]:
        self._connection.commit()
        cursor = self._connection.execute("SELECT key_blob FROM records ORDER BY ordinal")
        for (blob,) in cursor:
            yield pickle.loads(blob)

    def iter_join_keys(self) -> Iterator[tuple[Any, ...]]:
        """Yield keys while retaining only a packed join-index lookup."""
        self._connection.commit()
        self._join_ordinals = array("Q")
        cursor = self._connection.execute("SELECT ordinal, key_blob FROM records ORDER BY ordinal")
        for ordinal, blob in cursor:
            self._join_ordinals.append(int(ordinal))
            yield pickle.loads(blob)

    def item_at_join_index(self, index: int) -> tuple[tuple[Any, ...], dict[str, Any]]:
        if index < 0:
            index += len(self._join_ordinals)
        if not 0 <= index < len(self._join_ordinals):
            raise IndexError(index)
        row = self._connection.execute(
            "SELECT key_blob, record_blob FROM records WHERE ordinal = ?", (int(self._join_ordinals[index]),)
        ).fetchone()
        if row is None:
            raise IndexError(index)
        return pickle.loads(row[0]), pickle.loads(row[1])

    def __len__(self) -> int:
        row = self._connection.execute("SELECT COUNT(*) FROM records").fetchone()
        return int(row[0]) if row is not None else 0

    def __contains__(self, key: object) -> bool:
        if not isinstance(key, tuple):
            return False
        row = self._connection.execute(
            "SELECT 1 FROM records WHERE stable_key = ? LIMIT 1", (stable_key(key),)
        ).fetchone()
        return row is not None

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        self._connection.close()
        self._path.unlink(missing_ok=True)

    def __enter__(self) -> DiskBackedRecordMap:
        return self

    def __exit__(self, _exc_type: object, _exc: object, _traceback: object) -> None:
        self.close()

    def __del__(self) -> None:
        try:
            self.close()
        except Exception:
            pass
