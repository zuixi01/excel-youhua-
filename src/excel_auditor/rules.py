from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .models import RuleSet
from .ids import new_ulid
from .strict_serialization import load_json_strict, load_yaml_strict


def load_rules(path: Path) -> RuleSet:
    text = path.read_text(encoding="utf-8")
    payload = parse_rule_document(text, path.suffix)
    return RuleSet.model_validate(payload)


def parse_rule_document(document: str, suffix: str) -> Any:
    normalized = suffix.lower()
    if normalized in {".yaml", ".yml"}:
        return load_yaml_strict(document, context="rule YAML")
    if normalized == ".json":
        return load_json_strict(document, context="rule JSON")
    raise ValueError("only JSON and YAML rule documents are supported")


class RuleRegistry:
    def __init__(self, root: Path) -> None:
        self.root = root
        self.root.mkdir(parents=True, exist_ok=True)

    def publish(self, rules: RuleSet) -> Path:
        directory = self.root / _safe_segment(rules.schema_id)
        directory.mkdir(parents=True, exist_ok=True)
        path = directory / f"{_safe_segment(rules.schema_version)}.json"
        payload = rules.model_dump(mode="json")
        if path.exists():
            existing = RuleSet.model_validate(load_json_strict(path.read_text(encoding="utf-8"), context="published rule JSON"))
            if existing.content_sha256 != rules.content_sha256:
                raise FileExistsError("published rule versions are immutable")
            return path
        path.write_text(json.dumps(payload, ensure_ascii=False, sort_keys=True, indent=2), encoding="utf-8")
        return path

    def get(self, schema_id: str, version: str) -> RuleSet:
        path = self.root / _safe_segment(schema_id) / f"{_safe_segment(version)}.json"
        if not path.is_file():
            raise FileNotFoundError(f"rule version not found: {schema_id}@{version}")
        return RuleSet.model_validate(load_json_strict(path.read_text(encoding="utf-8"), context="published rule JSON"))

    def versions(self, schema_id: str) -> list[dict[str, str]]:
        directory = self.root / _safe_segment(schema_id)
        if not directory.is_dir():
            return []
        result = []
        for path in sorted(directory.glob("*.json")):
            rules = RuleSet.model_validate(load_json_strict(path.read_text(encoding="utf-8"), context="published rule JSON"))
            result.append({"schema_id": rules.schema_id, "version": rules.schema_version, "config_sha256": rules.content_sha256, "status": "published"})
        return result


class DraftRegistry:
    def __init__(self, root: Path) -> None:
        self.root = root
        self.root.mkdir(parents=True, exist_ok=True)

    def create(self, schema_id: str, config: dict[str, Any], base_version: str | None = None) -> dict[str, Any]:
        _safe_segment(schema_id)
        draft_id = new_ulid("draft_")
        now = datetime.now(timezone.utc).isoformat()
        record = {"draft_id": draft_id, "schema_id": schema_id, "base_version": base_version, "status": "draft", "config": config, "created_at": now, "updated_at": now}
        self._write(record)
        return record

    def get(self, schema_id: str, draft_id: str) -> dict[str, Any]:
        path = self._path(schema_id, draft_id)
        if not path.is_file():
            raise FileNotFoundError(f"draft not found: {draft_id}")
        return load_json_strict(path.read_text(encoding="utf-8"), context="draft rule JSON")

    def update(self, schema_id: str, draft_id: str, config: dict[str, Any]) -> dict[str, Any]:
        record = self.get(schema_id, draft_id)
        if record["status"] != "draft":
            raise ValueError("published drafts are immutable")
        if config.get("schema_id") != schema_id:
            raise ValueError("draft schema_id does not match path")
        record["config"] = config
        record["updated_at"] = datetime.now(timezone.utc).isoformat()
        self._write(record)
        return record

    def validate(self, schema_id: str, draft_id: str) -> RuleSet:
        record = self.get(schema_id, draft_id)
        return RuleSet.model_validate(record["config"])

    def mark_published(self, schema_id: str, draft_id: str) -> dict[str, Any]:
        record = self.get(schema_id, draft_id)
        record["status"] = "published"
        record["updated_at"] = datetime.now(timezone.utc).isoformat()
        self._write(record)
        return record

    def confirm_mapping(self, schema_id: str, draft_id: str, raw_header: str, canonical_field: str, sheet_id: str | None = None) -> dict[str, Any]:
        record = self.get(schema_id, draft_id)
        if record["status"] != "draft":
            raise ValueError("published drafts are immutable")
        sheets = record["config"].get("sheets", [])
        candidates = [sheet for sheet in sheets if sheet_id is None or sheet.get("id") == sheet_id]
        matches = []
        for sheet in candidates:
            for column in sheet.get("columns", []):
                if column.get("name") == canonical_field:
                    matches.append(column)
        if len(matches) != 1:
            raise ValueError("canonical field must resolve to exactly one draft column; provide sheet_id when needed")
        aliases = matches[0].setdefault("aliases", [])
        if raw_header not in aliases:
            aliases.append(raw_header)
        RuleSet.model_validate(record["config"])
        record["updated_at"] = datetime.now(timezone.utc).isoformat()
        self._write(record)
        return record

    def _path(self, schema_id: str, draft_id: str) -> Path:
        if not draft_id.startswith("draft_") or not draft_id[6:].isalnum():
            raise ValueError("invalid draft id")
        return self.root / _safe_segment(schema_id) / f"{draft_id}.json"

    def _write(self, record: dict[str, Any]) -> None:
        path = self._path(record["schema_id"], record["draft_id"])
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = path.with_suffix(".tmp")
        temporary.write_text(json.dumps(record, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")
        temporary.replace(path)


def _safe_segment(value: str) -> str:
    if not value or any(char not in "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_." for char in value):
        raise ValueError("identifier contains unsafe characters")
    return value
