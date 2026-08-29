from __future__ import annotations

import secrets
import time


_CROCKFORD = "0123456789ABCDEFGHJKMNPQRSTVWXYZ"


def new_ulid(prefix: str = "") -> str:
    """Return an unpredictable, time-sortable ULID with an optional resource prefix."""
    timestamp_ms = int(time.time_ns() // 1_000_000)
    if timestamp_ms >= 1 << 48:
        raise OverflowError("ULID timestamp exceeds 48 bits")
    value = (timestamp_ms << 80) | secrets.randbits(80)
    encoded = ["0"] * 26
    for index in range(25, -1, -1):
        encoded[index] = _CROCKFORD[value & 31]
        value >>= 5
    return prefix + "".join(encoded)
