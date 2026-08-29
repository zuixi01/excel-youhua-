import json
import sys

import pytest

from excel_auditor.performance_gate import MetricLimit, PerformanceGateError, compare_performance, main


def _workbook(elapsed=10.0, memory=100.0, **overrides):
    payload = {
        "benchmark_version": 3,
        "rows": 10_000,
        "columns": 50,
        "sheets": 1,
        "density": 1.0,
        "difference_rate": 0.1,
        "elapsed_seconds": elapsed,
        "peak_rss_delta_mib": memory,
    }
    payload.update(overrides)
    return payload


def test_performance_gate_accepts_noise_within_ratio_or_absolute_slack():
    result = compare_performance(
        _workbook(elapsed=14.0, memory=150.0),
        _workbook(),
        [MetricLimit("elapsed_seconds", 0.25, 5.0), MetricLimit("peak_rss_delta_mib", 0.25, 64.0)],
    )
    assert result["passed"] is True
    assert all(item["passed"] for item in result["comparisons"])


def test_performance_gate_rejects_material_regression():
    result = compare_performance(
        _workbook(elapsed=16.0, memory=170.0),
        _workbook(),
        [MetricLimit("elapsed_seconds", 0.25, 5.0), MetricLimit("peak_rss_delta_mib", 0.25, 64.0)],
    )
    assert result["passed"] is False
    assert {item["metric"] for item in result["comparisons"] if not item["passed"]} == {
        "elapsed_seconds",
        "peak_rss_delta_mib",
    }


def test_performance_gate_refuses_mismatched_or_invalid_reference():
    with pytest.raises(PerformanceGateError, match="scenarios do not match"):
        compare_performance(_workbook(rows=20_000), _workbook(), [MetricLimit("elapsed_seconds", 0.25, 5.0)])
    with pytest.raises(PerformanceGateError, match="benchmark_version"):
        compare_performance(_workbook(benchmark_version=3), _workbook(benchmark_version=2), [MetricLimit("elapsed_seconds", 0.25, 5.0)])
    with pytest.raises(PerformanceGateError, match="must be positive"):
        compare_performance(_workbook(), _workbook(elapsed=0), [MetricLimit("elapsed_seconds", 0.25, 5.0)])


def test_cli_accepts_normalized_cpu_metric(tmp_path, monkeypatch):
    reference = _workbook(benchmark_version=4, normalized_cpu_units=50.0)
    current = _workbook(benchmark_version=4, normalized_cpu_units=55.0)
    reference_path = tmp_path / "reference.json"
    current_path = tmp_path / "current.json"
    reference_path.write_text(json.dumps(reference), encoding="utf-8")
    current_path.write_text(json.dumps(current), encoding="utf-8")
    monkeypatch.setattr(sys, "argv", [
        "performance-gate",
        "--current", str(current_path),
        "--reference", str(reference_path),
        "--time-metric", "normalized_cpu_units",
    ])
    assert main() == 0


def test_performance_gate_checks_uploaded_standard_identity():
    current = {
        "benchmark_version": 4,
        "source_format": "json_upload",
        "standard_records": 500_000,
        "normalized_cpu_units": 10.0,
    }
    with pytest.raises(PerformanceGateError, match="source_format"):
        compare_performance(
            current,
            {**current, "source_format": "csv_upload"},
            [MetricLimit("normalized_cpu_units", 0.25, 5.0)],
        )
