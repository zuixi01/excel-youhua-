from __future__ import annotations

from typing import Any, Iterable

import pandas as pd
from datacompy import PandasCompare


class DataComPyAdapter:
    """Optional independent join-set cross-check; typed comparison remains in the core engine."""

    def all_rows_overlap(self, excel_keys: Iterable[tuple[Any, ...]], standard_keys: Iterable[tuple[Any, ...]]) -> bool:
        left = pd.DataFrame({"__key": sorted((_stable_key(key) for key in excel_keys))})
        right = pd.DataFrame({"__key": sorted((_stable_key(key) for key in standard_keys))})
        comparison = PandasCompare(left, right, join_columns="__key", df1_name="excel", df2_name="standard")
        return bool(comparison.all_rows_overlap())


def _stable_key(key: tuple[Any, ...]) -> str:
    return "\x1f".join(f"{kind}:{value!s}" for kind, value in key)
