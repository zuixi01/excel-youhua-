import pytest
from datetime import date, datetime, timezone
from decimal import Decimal
from hypothesis import given, settings, strategies as st

from excel_auditor.dataset_adapter import DataComPyAdapter
from excel_auditor.models import RuleSet
from excel_auditor.pandera_adapter import StandardDataValidator
from excel_auditor.partitioned_join import PolarsPartitionedKeyConnector


def test_datacompy_adapter_cross_checks_key_overlap():
    adapter = DataComPyAdapter()
    assert adapter.all_rows_overlap([(("string", "E1"),)], [(("string", "E1"),)])
    assert not adapter.all_rows_overlap([(("string", "E1"),)], [(("string", "E2"),)])


def test_polars_partitioned_connector_classifies_typed_composite_keys():
    shared = (("decimal", Decimal("1.20")), ("date", date(2026, 8, 29)))
    excel_only = (("datetime", datetime(2026, 8, 29, 1, 2, tzinfo=timezone.utc)),)
    standard_only = (("set", ("A", "B")),)
    joined = PolarsPartitionedKeyConnector(batch_size=1).classify(
        [shared, excel_only],
        [standard_only, shared],
    )
    assert joined.excel_only == [1]
    assert joined.standard_only == [0]
    assert joined.matched == [(0, 1)]


@given(st.sets(st.integers(min_value=-10_000, max_value=10_000), max_size=200), st.sets(st.integers(min_value=-10_000, max_value=10_000), max_size=200))
@settings(max_examples=25, deadline=None)
def test_polars_partitioned_connector_matches_set_algebra(excel_values, standard_values):
    excel_keys = [(('integer', value),) for value in sorted(excel_values)]
    standard_keys = [(('integer', value),) for value in sorted(standard_values)]
    joined = PolarsPartitionedKeyConnector(batch_size=17).classify(excel_keys, standard_keys)
    assert {excel_keys[index] for index in joined.excel_only} == set(excel_keys) - set(standard_keys)
    assert {standard_keys[index] for index in joined.standard_only} == set(standard_keys) - set(excel_keys)
    assert {
        (excel_keys[excel_index], standard_keys[standard_index])
        for excel_index, standard_index in joined.matched
    } == {(key, key) for key in set(excel_keys) & set(standard_keys)}


def test_pandera_standard_validator_rejects_duplicate_and_invalid_standard_values():
    rules = RuleSet.model_validate({
        "schema_id": "standard", "schema_version": "1.0.0", "name": "Standard",
        "sheets": [{"id": "data", "name": "Data", "primary_key": ["id"], "columns": [
            {"name": "id", "title": "ID", "required": True, "validation": {"nullable": False, "unique": True, "regex": "^E[0-9]+$"}},
            {"name": "status", "title": "状态", "type": "enum", "enum_values": ["在职", "离职"]}
        ]}],
    })
    with pytest.raises(ValueError, match="STANDARD_DATA_INVALID"):
        StandardDataValidator().validate({"data": [{"id": "bad", "status": "未知"}, {"id": "bad", "status": "在职"}]}, rules)


def test_pandera_standard_validator_accepts_string_dtype_and_checks_typed_range():
    rules = RuleSet.model_validate({
        "schema_id": "standard", "schema_version": "1.0.0", "name": "Standard",
        "sheets": [{"id": "data", "name": "Data", "primary_key": ["id"], "columns": [
            {"name": "id", "title": "ID", "required": True, "validation": {"nullable": False, "unique": True}},
            {"name": "amount", "title": "Amount", "type": "decimal", "validation": {"min": "0", "max": "100"}}
        ]}],
    })
    validator = StandardDataValidator()
    validator.validate({"data": [{"id": "E1", "amount": "12.50"}]}, rules)
    with pytest.raises(ValueError, match="typed failures.*amount"):
        validator.validate({"data": [{"id": "E1", "amount": "101"}]}, rules)


def test_standard_validator_enforces_composite_primary_key_and_sheet_set():
    rules = RuleSet.model_validate({
        "schema_id": "composite", "schema_version": "1.0.0", "name": "Composite",
        "sheets": [{"id": "data", "name": "Data", "primary_key": ["company", "employee"], "columns": [
            {"name": "company", "title": "Company", "required": True},
            {"name": "employee", "title": "Employee", "required": True},
        ]}],
    })
    validator = StandardDataValidator()
    with pytest.raises(ValueError, match="duplicate composite primary key"):
        validator.validate({"data": [
            {"company": "A", "employee": "1"},
            {"company": "A", "employee": "1"},
        ]}, rules)
    with pytest.raises(ValueError, match="sheet set mismatch"):
        validator.validate({"unknown": []}, rules)
