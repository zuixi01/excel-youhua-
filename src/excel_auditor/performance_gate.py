from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any


class PerformanceGateError(ValueError):
    pass


@dataclass(frozen=True)
class MetricLimit:
    name: str
    ratio: float
    absolute_slack: float


def compare_performance(
    current: dict[str, Any],
    reference: dict[str, Any],
    limits: list[MetricLimit],
) -> dict[str, Any]:
    _validate_scenario_identity(current, reference)
    comparisons: list[dict[str, Any]] = []
    regressions: list[str] = []
    for limit in limits:
        current_value = _positive_number(current, limit.name, allow_zero=True)
        reference_value = _positive_number(reference, limit.name)
        ceiling = max(
            reference_value * (1 + limit.ratio),
            reference_value + limit.absolute_slack,
        )
        passed = current_value <= ceiling
        comparisons.append(
            {
                "metric": limit.name,
                "reference": reference_value,
                "current": current_value,
                "ceiling": round(ceiling, 6),
                "passed": passed,
            }
        )
        if not passed:
            regressions.append(
                f"{limit.name} regressed: current={current_value}, reference={reference_value}, ceiling={ceiling:.6f}"
            )
    return {"passed": not regressions, "comparisons": comparisons, "regressions": regressions}


def _validate_scenario_identity(current: dict[str, Any], reference: dict[str, Any]) -> None:
    workbook_keys = ("rows", "columns", "sheets", "density", "difference_rate")
    http_keys = ("records", "page_size", "pages")
    keys = workbook_keys if "rows" in current or "rows" in reference else http_keys
    mismatches = [key for key in keys if current.get(key) != reference.get(key)]
    if mismatches:
        details = ", ".join(f"{key}={reference.get(key)!r}->{current.get(key)!r}" for key in mismatches)
        raise PerformanceGateError(f"performance scenarios do not match: {details}")


def _positive_number(payload: dict[str, Any], name: str, *, allow_zero: bool = False) -> float:
    value = payload.get(name)
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise PerformanceGateError(f"metric {name!r} is missing or not numeric")
    numeric = float(value)
    if numeric < 0 or (numeric == 0 and not allow_zero):
        raise PerformanceGateError(f"metric {name!r} must be {'non-negative' if allow_zero else 'positive'}")
    return numeric


def main() -> int:
    parser = argparse.ArgumentParser(description="Compare current performance metrics with a prior stable run")
    parser.add_argument("--current", type=Path, required=True)
    parser.add_argument("--reference", type=Path, required=True)
    parser.add_argument("--max-time-ratio", type=float, default=0.25)
    parser.add_argument("--time-slack-seconds", type=float, default=5.0)
    parser.add_argument("--max-memory-ratio", type=float)
    parser.add_argument("--memory-slack-mib", type=float, default=64.0)
    args = parser.parse_args()
    if args.max_time_ratio < 0 or args.time_slack_seconds < 0:
        parser.error("time ratio and slack must be non-negative")
    limits = [MetricLimit("elapsed_seconds", args.max_time_ratio, args.time_slack_seconds)]
    if args.max_memory_ratio is not None:
        if args.max_memory_ratio < 0 or args.memory_slack_mib < 0:
            parser.error("memory ratio and slack must be non-negative")
        limits.append(MetricLimit("peak_rss_delta_mib", args.max_memory_ratio, args.memory_slack_mib))
    current = json.loads(args.current.read_text(encoding="utf-8"))
    reference = json.loads(args.reference.read_text(encoding="utf-8"))
    result = compare_performance(current, reference, limits)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
