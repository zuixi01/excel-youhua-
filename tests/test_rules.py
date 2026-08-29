from pathlib import Path

import pytest
from pydantic import ValidationError

from excel_auditor.models import RuleSet
from excel_auditor.rules import DraftRegistry, RuleRegistry, load_rules
from excel_auditor.validators import register_validator


EXAMPLE = Path("configs/examples/employee-roster.yaml")


def test_example_rule_is_valid_and_hash_is_deterministic():
    first = load_rules(EXAMPLE)
    second = load_rules(EXAMPLE)
    assert first.content_sha256 == second.content_sha256
    assert first.sheets[0].primary_key == ["employee_id"]


@pytest.mark.parametrize(
    ("suffix", "document"),
    [
        (".json", '{"schema_id":"first","schema_id":"second"}'),
        (".yaml", "schema_id: first\nschema_id: second\n"),
    ],
)
def test_rule_document_loader_rejects_duplicate_keys(tmp_path, suffix, document):
    path = tmp_path / f"duplicate{suffix}"
    path.write_text(document, encoding="utf-8")

    with pytest.raises(ValueError, match="duplicate"):
        load_rules(path)


def test_published_version_is_immutable(tmp_path):
    registry = RuleRegistry(tmp_path)
    rules = load_rules(EXAMPLE)
    registry.publish(rules)
    changed = rules.model_copy(update={"name": "changed"})
    with pytest.raises(FileExistsError):
        registry.publish(changed)


def test_primary_key_must_be_required():
    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["columns"][0]["required"] = False
    with pytest.raises(ValidationError):
        RuleSet.model_validate(payload)


def test_primary_key_rejects_lossy_normalizers():
    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["columns"][0]["normalize"].append("collapse_spaces")
    with pytest.raises(ValidationError):
        RuleSet.model_validate(payload)


@pytest.mark.parametrize(
    ("field_update", "message"),
    [
        ({"compare": {"mode": "ignore_case"}}, "compare.mode=exact"),
        ({"compare": {"formula_mode": "formula"}}, "must reject formulas"),
        ({"type": "fuzzy_string"}, "cannot use fuzzy_string"),
    ],
)
def test_primary_key_rejects_non_deterministic_matching_modes(field_update, message):
    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["columns"][0].update(field_update)
    with pytest.raises(ValidationError, match=message):
        RuleSet.model_validate(payload)


def test_cross_field_rule_must_reference_existing_fields():
    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["cross_field_rules"] = [{"rule_id": "bad", "validator": "conditional_required", "params": {"when_field": "missing", "equals": "x", "required_field": "salary"}}]
    with pytest.raises(ValidationError):
        RuleSet.model_validate(payload)


def test_cross_field_validator_must_be_registered_by_trusted_code():
    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["cross_field_rules"] = [{"rule_id": "custom", "validator": "not_registered", "params": {}}]
    with pytest.raises(ValidationError, match="unregistered"):
        RuleSet.model_validate(payload)
    register_validator("always_valid_test", lambda _parsed, _params: None)
    payload["sheets"][0]["cross_field_rules"] = [{"rule_id": "custom", "validator": "always_valid_test", "params": {}}]
    assert RuleSet.model_validate(payload).sheets[0].cross_field_rules[0].validator == "always_valid_test"


def test_unknown_nested_rule_key_is_rejected():
    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["actions"]["mismatched_value_typo"] = "mark_yellow"
    with pytest.raises(ValidationError, match="mismatched_value_typo"):
        RuleSet.model_validate(payload)


def test_unsafe_schema_id_and_cross_sheet_alias_collision_are_rejected():
    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["schema_id"] = "../unsafe"
    with pytest.raises(ValidationError, match="schema_id"):
        RuleSet.model_validate(payload)

    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"].append({
        "id": "second",
        "name": "第二张表",
        "aliases": [payload["sheets"][0]["name"]],
        "primary_key": ["id"],
        "columns": [{"name": "id", "title": "ID", "required": True}],
    })
    with pytest.raises(ValidationError, match="shared by sheets"):
        RuleSet.model_validate(payload)

    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0].update({"primary_key": [], "primary_key_mode": "row_number", "row_number_field": "../row"})
    with pytest.raises(ValidationError, match="row_number_field"):
        RuleSet.model_validate(payload)

    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0].update({"primary_key": [], "primary_key_mode": "row_number", "row_number_field": "row", "empty_primary_key_action": "skip_row"})
    with pytest.raises(ValidationError, match="fallback action"):
        RuleSet.model_validate(payload)

    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["columns"][0]["aliases"].append("row")
    payload["sheets"][0].update({"primary_key": [], "primary_key_mode": "row_number", "row_number_field": "row"})
    with pytest.raises(ValidationError, match="collides with column"):
        RuleSet.model_validate(payload)


def test_decimal_tolerance_requires_string_and_risky_regex_is_rejected():
    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["columns"][2]["compare"]["absolute_tolerance"] = 0.1
    with pytest.raises(ValidationError, match="configured as a string"):
        RuleSet.model_validate(payload)
    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["columns"][0]["validation"]["regex"] = "^(a+)+$"
    with pytest.raises(ValidationError, match="high-complexity"):
        RuleSet.model_validate(payload)


@pytest.mark.parametrize("field", ["absolute_tolerance", "relative_tolerance"])
def test_negative_numeric_tolerance_is_rejected(field):
    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["columns"][2]["compare"][field] = "-0.01"
    with pytest.raises(ValidationError, match="non-negative"):
        RuleSet.model_validate(payload)


@pytest.mark.parametrize("value", ["NaN", "Infinity", "-Infinity"])
def test_non_finite_numeric_configuration_is_rejected(value):
    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["columns"][2]["compare"]["absolute_tolerance"] = value
    with pytest.raises(ValidationError, match="finite"):
        RuleSet.model_validate(payload)

    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["columns"][2]["validation"]["min"] = value
    with pytest.raises(ValidationError, match="finite"):
        RuleSet.model_validate(payload)


def test_conflicting_validation_and_type_specific_options_are_rejected():
    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["columns"][2]["validation"].update({"min": "10", "max": "1"})
    with pytest.raises(ValidationError, match="min cannot exceed max"):
        RuleSet.model_validate(payload)

    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["columns"][1]["validation"]["min"] = "1"
    with pytest.raises(ValidationError, match="numeric validation bounds are incompatible with string"):
        RuleSet.model_validate(payload)

    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["columns"][1]["compare"]["absolute_tolerance"] = "0.1"
    with pytest.raises(ValidationError, match="numeric comparison options are incompatible with string"):
        RuleSet.model_validate(payload)

    payload = load_rules(EXAMPLE).model_dump(mode="json")
    amount = payload["sheets"][0]["columns"][2]
    amount["compare"].update({"mode": "exact", "absolute_tolerance": "0.1"})
    with pytest.raises(ValidationError, match="require compare.mode=numeric"):
        RuleSet.model_validate(payload)

    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["columns"][1]["enum_values"] = ["ignored"]
    with pytest.raises(ValidationError, match="enum configuration is incompatible with string"):
        RuleSet.model_validate(payload)


def test_enum_aliases_and_static_repair_defaults_must_satisfy_the_rule():
    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["columns"][4]["enum_aliases"]["未知"] = "不存在"
    with pytest.raises(ValidationError, match="target unknown values"):
        RuleSet.model_validate(payload)

    payload = load_rules(EXAMPLE).model_dump(mode="json")
    name = payload["sheets"][0]["columns"][1]
    name.update({"fill_static_default": True, "static_default": "x"})
    name["validation"] = {"nullable": True, "regex": "^[A-Z]{2}$"}
    with pytest.raises(ValidationError, match="static_default.*validation regex"):
        RuleSet.model_validate(payload)

    payload = load_rules(EXAMPLE).model_dump(mode="json")
    name = payload["sheets"][0]["columns"][1]
    name.update({"fill_static_default": True, "static_default": " N/A ", "normalize": ["trim"], "value_aliases": {"N/A": ""}})
    with pytest.raises(ValidationError, match="normalizes to an empty value"):
        RuleSet.model_validate(payload)


def test_ignore_case_enum_configuration_rejects_folded_collisions():
    payload = load_rules(EXAMPLE).model_dump(mode="json")
    department = payload["sheets"][0]["columns"][4]
    department["compare"]["mode"] = "ignore_case"
    department["enum_values"] = ["HR", "hr"]
    department["enum_aliases"] = {}
    with pytest.raises(ValidationError, match="ambiguous under ignore_case"):
        RuleSet.model_validate(payload)


def test_standard_source_rejects_ignored_or_overlapping_http_configuration():
    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["standard_source"] = {"type": "upload", "path": "/ignored"}
    with pytest.raises(ValidationError, match="cannot contain managed HTTP configuration"):
        RuleSet.model_validate(payload)

    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["standard_source"] = {
        "type": "managed_http", "connection_id": "hr", "path": "/employees",
        "static_parameters": {"department": "D1"},
        "parameter_mapping": {"department": "department_id"},
    }
    with pytest.raises(ValidationError, match="parameters overlap"):
        RuleSet.model_validate(payload)

    payload["standard_source"] = {
        "type": "managed_http", "connection_id": "hr", "path": "/employees",
        "static_parameters": {"page": 1},
        "pagination": {"page_param": "page", "size_param": "size"},
    }
    with pytest.raises(ValidationError, match="pagination parameters overlap"):
        RuleSet.model_validate(payload)


@pytest.mark.parametrize("json_path", ["data.items", "$.data..items", "$.data."])
def test_standard_source_rejects_json_paths_runtime_cannot_interpret(json_path):
    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["standard_source"] = {
        "type": "managed_http", "connection_id": "hr", "path": "/employees",
        "data_json_path": json_path,
    }
    with pytest.raises(ValidationError, match="simple non-empty object path"):
        RuleSet.model_validate(payload)


def test_draft_mapping_confirmation_is_validated_and_published_draft_is_immutable(tmp_path):
    rules = load_rules(EXAMPLE)
    drafts = DraftRegistry(tmp_path / "drafts")
    record = drafts.create(rules.schema_id, rules.model_dump(mode="json"))
    drafts.confirm_mapping(rules.schema_id, record["draft_id"], "员工ID号", "employee_id", "employees")
    validated = drafts.validate(rules.schema_id, record["draft_id"])
    assert "员工ID号" in validated.sheets[0].columns[0].aliases
    drafts.mark_published(rules.schema_id, record["draft_id"])
    with pytest.raises(ValueError, match="immutable"):
        drafts.update(rules.schema_id, record["draft_id"], validated.model_dump(mode="json"))
