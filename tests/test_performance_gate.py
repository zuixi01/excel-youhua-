import pytest

from excel_auditor.performance_gate import MetricLimit, PerformanceGateError, compare_performance


def _workbook(elapsed=10.0, memory=100.0, **overrides):
    payload = {
        "benchmark_version": 2,
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
        compare_performance(_workbook(benchmark_version=2), _workbook(benchmark_version=1), [MetricLimit("elapsed_seconds", 0.25, 5.0)])
    with pytest.raises(PerformanceGateError, match="must be positive"):
        compare_performance(_workbook(), _workbook(elapsed=0), [MetricLimit("elapsed_seconds", 0.25, 5.0)])
