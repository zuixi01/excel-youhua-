from __future__ import annotations

import json
import tempfile
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterable

import polars as pl


@dataclass(frozen=True)
class PartitionedKeyJoin:
    excel_only: list[int]
    standard_only: list[int]
    matched: list[tuple[int, int]]


class PolarsPartitionedKeyConnector:
    """Classify normalized keys through bounded Parquet partitions."""

    def __init__(self, batch_size: int = 50_000) -> None:
        if batch_size < 1:
            raise ValueError("batch_size must be positive")
        self.batch_size = batch_size

    def classify(self, excel_keys: Iterable[tuple[Any, ...]], standard_keys: Iterable[tuple[Any, ...]]) -> PartitionedKeyJoin:
        with tempfile.TemporaryDirectory(prefix="excel-auditor-key-join-") as temporary:
            root = Path(temporary)
            self._write_partitions(root, "excel", excel_keys)
            self._write_partitions(root, "standard", standard_keys)
            excel = pl.scan_parquet(str(root / "excel-*.parquet")).rename({"row_id": "excel_id"})
            standard = pl.scan_parquet(str(root / "standard-*.parquet")).rename({"row_id": "standard_id"})
            excel_only = (
                excel.join(standard.select("key"), on="key", how="anti")
                .sort("key")
                .select("excel_id")
                .collect(engine="streaming")["excel_id"]
                .to_list()
            )
            standard_only = (
                standard.join(excel.select("key"), on="key", how="anti")
                .sort("key")
                .select("standard_id")
                .collect(engine="streaming")["standard_id"]
                .to_list()
            )
            matched_frame = (
                excel.join(standard, on="key", how="inner")
                .sort("key")
                .select("excel_id", "standard_id")
                .collect(engine="streaming")
            )
            matched = list(zip(matched_frame["excel_id"].to_list(), matched_frame["standard_id"].to_list()))
            return PartitionedKeyJoin(excel_only, standard_only, matched)

    def _write_partitions(self, root: Path, prefix: str, keys: Iterable[tuple[Any, ...]]) -> None:
        batch_keys: list[str] = []
        batch_ids: list[int] = []
        part = 0
        for row_id, key in enumerate(keys):
            batch_keys.append(stable_key(key))
            batch_ids.append(row_id)
            if len(batch_keys) >= self.batch_size:
                self._write_part(root, prefix, part, batch_keys, batch_ids)
                batch_keys, batch_ids, part = [], [], part + 1
        if batch_keys or part == 0:
            self._write_part(root, prefix, part, batch_keys, batch_ids)

    @staticmethod
    def _write_part(root: Path, prefix: str, part: int, keys: list[str], row_ids: list[int]) -> None:
        frame = pl.DataFrame(
            {"key": keys, "row_id": row_ids},
            schema={"key": pl.String, "row_id": pl.UInt64},
        )
        frame.write_parquet(root / f"{prefix}-{part:06d}.parquet", compression="zstd")


def stable_key(key: tuple[Any, ...]) -> str:
    return json.dumps(key, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=_json_key_value)


def _json_key_value(value: Any) -> Any:
    if isinstance(value, Decimal):
        return {"$decimal": str(value)}
    if isinstance(value, datetime):
        return {"$datetime": value.isoformat()}
    if isinstance(value, date):
        return {"$date": value.isoformat()}
    raise TypeError(f"unsupported normalized key type: {type(value).__name__}")
