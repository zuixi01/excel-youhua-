from __future__ import annotations

import pickle
import tempfile
from array import array
from collections.abc import Callable, Iterable, Sequence
from typing import Generic, TypeVar, overload


T = TypeVar("T")


class _SpilledObjects(Sequence[T], Generic[T]):
    """Random-access temporary object sequence with a compact offset index."""

    def __init__(self) -> None:
        self._file = tempfile.TemporaryFile(prefix="excel-auditor-items-", suffix=".bin")
        self._offsets = array("Q")

    def append(self, item: T) -> None:
        self._offsets.append(self._file.tell())
        pickle.dump(item, self._file, protocol=5)

    def __len__(self) -> int:
        return len(self._offsets)

    @overload
    def __getitem__(self, index: int) -> T: ...

    @overload
    def __getitem__(self, index: slice) -> Iterable[T]: ...

    def __getitem__(self, index: int | slice) -> T | Iterable[T]:
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


class SpillableSequence(Sequence[T], Generic[T]):
    """Append-only sequence that migrates payloads to disk past a threshold."""

    def __init__(self, spill_after_items: int, on_append: Callable[[T], None] | None = None) -> None:
        if spill_after_items < 1:
            raise ValueError("spill_after_items must be positive")
        self._spill_after_items = spill_after_items
        self._on_append = on_append
        self._items: list[T] = []
        self._spilled: _SpilledObjects[T] | None = None

    @property
    def spilled(self) -> bool:
        return self._spilled is not None

    def append(self, item: T) -> None:
        if self._on_append is not None:
            self._on_append(item)
        if self._spilled is None and len(self._items) >= self._spill_after_items:
            spilled = _SpilledObjects[T]()
            try:
                for existing in self._items:
                    spilled.append(existing)
            except Exception:
                spilled.close()
                raise
            self._items.clear()
            self._spilled = spilled
        if self._spilled is None:
            self._items.append(item)
        else:
            self._spilled.append(item)

    def extend(self, items: Iterable[T]) -> None:
        for item in items:
            self.append(item)

    def __len__(self) -> int:
        return len(self._spilled) if self._spilled is not None else len(self._items)

    @overload
    def __getitem__(self, index: int) -> T: ...

    @overload
    def __getitem__(self, index: slice) -> Iterable[T]: ...

    def __getitem__(self, index: int | slice) -> T | Iterable[T]:
        if self._spilled is not None:
            return self._spilled[index]
        if isinstance(index, slice):
            return iter(self._items[index])
        return self._items[index]

    def finish(self) -> list[T] | SpillableSequence[T]:
        return self if self.spilled else self._items

    def close(self) -> None:
        if self._spilled is not None:
            self._spilled.close()
            self._spilled = None
        self._items.clear()
