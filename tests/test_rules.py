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


def test_decimal_tolerance_requires_string_and_risky_regex_is_rejected():
    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["columns"][2]["compare"]["absolute_tolerance"] = 0.1
    with pytest.raises(ValidationError, match="configured as a string"):
        RuleSet.model_validate(payload)
    payload = load_rules(EXAMPLE).model_dump(mode="json")
    payload["sheets"][0]["columns"][0]["validation"]["regex"] = "^(a+)+$"
    with pytest.raises(ValidationError, match="high-complexity"):
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
