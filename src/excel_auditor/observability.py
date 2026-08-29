from __future__ import annotations

import json
import logging
import threading
from collections import Counter
from datetime import datetime, timezone
from typing import Any


class JsonFormatter(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        payload = {
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "level": record.levelname.lower(),
            "service": getattr(record, "service", "excel-auditor"),
            "trace_id": getattr(record, "trace_id", None),
            "job_id": getattr(record, "job_id", None),
            "stage": getattr(record, "stage", None),
            "event": getattr(record, "event", record.getMessage()),
            "duration_ms": getattr(record, "duration_ms", None),
            "safe_error_code": getattr(record, "safe_error_code", None),
            "rows_processed": getattr(record, "rows_processed", None),
            "differences": getattr(record, "differences", None),
            "temp_disk_bytes": getattr(record, "temp_disk_bytes", None),
        }
        return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def configure_logging() -> None:
    root = logging.getLogger()
    if any(getattr(handler, "_excel_auditor_json", False) for handler in root.handlers):
        return
    handler = logging.StreamHandler()
    handler.setFormatter(JsonFormatter())
    handler._excel_auditor_json = True  # type: ignore[attr-defined]
    root.handlers.clear()
    root.addHandler(handler)
    root.setLevel(logging.INFO)


def log_event(*, job_id: str | None, stage: str, event: str, trace_id: str | None = None, duration_ms: int | None = None, safe_error_code: str | None = None, rows_processed: int | None = None, differences: int | None = None, temp_disk_bytes: int | None = None) -> None:
    logging.getLogger("excel_auditor").info(event, extra={"service": "excel-auditor", "trace_id": trace_id, "job_id": job_id, "stage": stage, "event": event, "duration_ms": duration_ms, "safe_error_code": safe_error_code, "rows_processed": rows_processed, "differences": differences, "temp_disk_bytes": temp_disk_bytes})


class Metrics:
    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._counters: Counter[tuple[str, tuple[tuple[str, str], ...]]] = Counter()
        self._observations: dict[tuple[str, tuple[tuple[str, str], ...]], tuple[int, float]] = {}
        self._gauges: dict[tuple[str, tuple[tuple[str, str], ...]], float] = {}

    def increment(self, name: str, amount: int = 1, **labels: Any) -> None:
        safe_labels = tuple(sorted((str(key), str(value)) for key, value in labels.items()))
        with self._lock:
            self._counters[(name, safe_labels)] += amount

    def observe(self, name: str, value: float, **labels: Any) -> None:
        safe_labels = tuple(sorted((str(key), str(label)) for key, label in labels.items()))
        with self._lock:
            count, total = self._observations.get((name, safe_labels), (0, 0.0))
            self._observations[(name, safe_labels)] = (count + 1, total + float(value))

    def set_gauge(self, name: str, value: float, **labels: Any) -> None:
        safe_labels = tuple(sorted((str(key), str(label)) for key, label in labels.items()))
        with self._lock:
            self._gauges[(name, safe_labels)] = float(value)

    def prometheus(self) -> str:
        lines = []
        with self._lock:
            items = sorted(self._counters.items())
            observations = sorted(self._observations.items())
            gauges = sorted(self._gauges.items())
        for (name, labels), value in items:
            suffix = ""
            if labels:
                encoded = ",".join(f'{key}="{label.replace(chr(92), chr(92) * 2).replace(chr(34), chr(92) + chr(34))}"' for key, label in labels)
                suffix = "{" + encoded + "}"
            lines.append(f"excel_auditor_{name}{suffix} {value}")
        for (name, labels), (count, total) in observations:
            suffix = "{" + ",".join(f'{key}="{label}"' for key, label in labels) + "}" if labels else ""
            lines.append(f"excel_auditor_{name}_count{suffix} {count}")
            lines.append(f"excel_auditor_{name}_sum{suffix} {total}")
        for (name, labels), value in gauges:
            suffix = "{" + ",".join(f'{key}="{label}"' for key, label in labels) + "}" if labels else ""
            lines.append(f"excel_auditor_{name}{suffix} {value}")
        return "\n".join(lines) + ("\n" if lines else "")


metrics = Metrics()
