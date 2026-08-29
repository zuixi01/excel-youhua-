from __future__ import annotations

import json
from pathlib import Path

import pytest

from excel_auditor.models import RuleSet
from excel_auditor.snapshots import SpilledRecords
from excel_auditor.standard_files import load_standard_file


def _rules(*, max_records: int = 500_000, two_sheets: bool = False) -> RuleSet:
    sheets = [{
        "id": "data",
        "name": "Data",
        "primary_key": ["id"],
        "columns": [
            {"name": "id", "title": "ID", "aliases": ["Identifier"], "required": True},
            {"name": "name", "title": "Name"},
        ],
    }]
    if two_sheets:
        sheets.append({
            "id": "other",
            "name": "Other",
            "primary_key": ["id"],
            "columns": [{"name": "id", "title": "ID", "required": True}],
        })
    return RuleSet.model_validate({
        "schema_id": "streamed-standard",
        "schema_version": "1.0.0",
        "name": "Streamed standard",
        "workbook": {"max_standard_records": max_records},
        "sheets": sheets,
    })


def _close(standard):
    for rows in standard.values():
        close = getattr(rows, "close", None)
        if close is not None:
            close()


def test_json_object_streams_canonical_records_and_spills_without_read_text(tmp_path, monkeypatch):
    path = tmp_path / "standard.json"
    payload = {"Data": [
        {"Identifier": "E1", "Name": "Alice", "ignored": "x"},
        {"Identifier": "E2", "Name": "Bob"},
        {"Identifier": "E3", "Name": "Carol"},
    ]}
    path.write_bytes(b"\xef\xbb\xbf" + json.dumps(payload).encode("utf-8"))
    original_read_text = Path.read_text

    def reject_full_text_read(self, *args, **kwargs):
        if self == path:
            raise AssertionError("standard JSON must not be read as one full string")
        return original_read_text(self, *args, **kwargs)

    monkeypatch.setattr(Path, "read_text", reject_full_text_read)
    standard = load_standard_file(path, _rules(), spill_after_records=2)
    try:
        assert list(standard) == ["data"]
        assert isinstance(standard["data"], SpilledRecords)
        assert list(standard["data"]) == [
            {"id": "E1", "name": "Alice"},
            {"id": "E2", "name": "Bob"},
            {"id": "E3", "name": "Carol"},
        ]
    finally:
        _close(standard)


def test_root_array_and_csv_are_streamed_for_single_sheet(tmp_path):
    json_path = tmp_path / "standard.json"
    json_path.write_text(json.dumps([{"ID": "E1"}, {"ID": "E2"}]), encoding="utf-8")
    csv_path = tmp_path / "standard.csv"
    csv_path.write_text("ID,Name\nE1,Alice\nE2,Bob\nE3,Carol\n", encoding="utf-8")

    json_standard = load_standard_file(json_path, _rules(), spill_after_records=10)
    csv_standard = load_standard_file(csv_path, _rules(), spill_after_records=2)
    try:
        assert list(json_standard["data"]) == [{"id": "E1"}, {"id": "E2"}]
        assert isinstance(csv_standard["data"], SpilledRecords)
        assert csv_standard["data"][-1] == {"id": "E3", "name": "Carol"}
    finally:
        _close(json_standard)
        _close(csv_standard)


@pytest.mark.parametrize(
    ("payload", "message"),
    [
        ('{"unknown": []}', "unknown sheet key"),
        ('{"data": {}, "other": []}', "must be an array of objects"),
        ('{"data": [1]}', "must be an array of objects"),
        ('{"data": [], "Data": []}', "duplicate sheet mapping"),
        ('{"data": [], "data": []}', "duplicate sheet key"),
        ('{"data": [', "uploaded JSON is malformed"),
        ('null', "root must be an object or array"),
    ],
)
def test_streamed_json_rejects_ambiguous_or_malformed_shapes(tmp_path, payload, message):
    path = tmp_path / "invalid.json"
    path.write_text(payload, encoding="utf-8")
    with pytest.raises(ValueError, match=message):
        load_standard_file(path, _rules(two_sheets=True), spill_after_records=2)


def test_streamed_json_enforces_record_limit_before_full_parse(tmp_path, monkeypatch):
    path = tmp_path / "too-many.json"
    path.write_text(json.dumps({"data": [{"id": "E1"}, {"id": "E2"}, {"id": "E3"}]}), encoding="utf-8")
    closed: list[bool] = []
    original_close = SpilledRecords.close

    def recording_close(self):
        closed.append(True)
        original_close(self)

    monkeypatch.setattr(SpilledRecords, "close", recording_close)
    with pytest.raises(ValueError, match="STANDARD_TOO_LARGE"):
        load_standard_file(path, _rules(max_records=2), spill_after_records=1)
    assert closed
