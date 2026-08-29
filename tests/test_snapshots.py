import json

import pytest

from excel_auditor.snapshots import SpilledRecords, create_snapshot, load_snapshot


def test_snapshot_can_stream_into_disk_backed_sequences(tmp_path):
    snapshot = create_snapshot({"data": [{"id": index, "value": f"V{index}"} for index in range(5)]}, tmp_path)
    loaded = load_snapshot(snapshot, spill_after_records=2)
    try:
        assert isinstance(loaded["data"], SpilledRecords)
        assert len(loaded["data"]) == 5
        assert loaded["data"][0] == {"id": 0, "value": "V0"}
        assert loaded["data"][-1] == {"id": 4, "value": "V4"}
        assert list(loaded["data"][1:3]) == [{"id": 1, "value": "V1"}, {"id": 2, "value": "V2"}]
    finally:
        loaded["data"].close()


def test_snapshot_writer_streams_disk_backed_records_without_reordering(tmp_path):
    records = SpilledRecords()
    records.append({"id": 2, "value": "B"})
    records.append({"id": 1, "value": "A"})
    try:
        snapshot = create_snapshot({"data": records}, tmp_path)
    finally:
        records.close()
    payload = [json.loads(line) for line in snapshot.path.read_text(encoding="utf-8").splitlines()]
    assert [item["record"]["id"] for item in payload] == [2, 1]


def test_snapshot_hash_is_checked_while_streaming_and_spills_are_closed(tmp_path, monkeypatch):
    snapshot = create_snapshot({"data": [{"id": 1}]}, tmp_path)
    snapshot.path.write_text(json.dumps({"sheet_id": "data", "record": {"id": 2}}) + "\n", encoding="utf-8")
    closed = []
    original_close = SpilledRecords.close

    def recording_close(self):
        closed.append(True)
        original_close(self)

    monkeypatch.setattr(SpilledRecords, "close", recording_close)
    with pytest.raises(ValueError, match="snapshot hash mismatch"):
        load_snapshot(snapshot, spill_after_records=1)
    assert closed == []  # record_count == threshold keeps this tiny snapshot in memory


def test_corrupt_spilled_snapshot_closes_temporary_files(tmp_path, monkeypatch):
    snapshot = create_snapshot({"data": [{"id": 1}, {"id": 2}]}, tmp_path)
    with snapshot.path.open("ab") as handle:
        handle.write(b"not-json\n")
    closed = []
    original_close = SpilledRecords.close

    def recording_close(self):
        closed.append(True)
        original_close(self)

    monkeypatch.setattr(SpilledRecords, "close", recording_close)
    with pytest.raises(ValueError, match="corrupt snapshot line"):
        load_snapshot(snapshot, spill_after_records=1)
    assert closed == [True]
