from __future__ import annotations

import argparse
from datetime import datetime
import hashlib
import json
from pathlib import Path, PurePath
import re
import sys
from typing import Any


SHA256 = re.compile(r"^[0-9a-f]{64}$")
REQUIRED_REVIEW_CHECKS = {
    "opens_without_repair_prompt",
    "sheet_names_and_layout_correct",
    "colors_comments_and_report_correct",
    "formulas_tables_filters_and_validation_correct",
    "macro_signature_activex_and_vml_correct",
}


def validate_automated_evidence(payload: Any) -> dict[str, Any]:
    if not isinstance(payload, dict):
        raise ValueError("automated evidence root must be an object")
    expected = {"schema_version", "generated_at", "excel_version", "macro_execution", "files", "summary"}
    _exact_keys(payload, expected, "automated evidence")
    if payload["schema_version"] != "1.0":
        raise ValueError("automated evidence schema_version must be 1.0")
    _timestamp(payload.get("generated_at"), "generated_at")
    if not isinstance(payload.get("excel_version"), str) or not payload["excel_version"].strip():
        raise ValueError("excel_version must be present")
    if payload["macro_execution"] != "force_disabled":
        raise ValueError("macro execution must be force_disabled")
    files = payload.get("files")
    if not isinstance(files, list) or len(files) < 2:
        raise ValueError("at least one .xlsx and one .xlsm result are required")
    seen: set[str] = set()
    xlsx_count = 0
    xlsm_count = 0
    critical_names: set[str] = set()
    for index, item in enumerate(files):
        if not isinstance(item, dict):
            raise ValueError(f"files[{index}] must be an object")
        _exact_keys(
            item,
            {
                "file_name",
                "extension",
                "input_sha256",
                "roundtrip_sha256",
                "opened",
                "roundtrip_opened",
                "saved_copy",
                "worksheet_names",
                "critical_part_count",
                "critical_part_names",
                "critical_parts_equal",
            },
            f"files[{index}]",
        )
        name = item.get("file_name")
        if not isinstance(name, str) or not name or PurePath(name).name != name or name in seen:
            raise ValueError(f"files[{index}].file_name must be a unique base name")
        seen.add(name)
        extension = item.get("extension")
        if extension not in {".xlsx", ".xlsm"} or not name.lower().endswith(extension):
            raise ValueError(f"files[{index}].extension is invalid")
        xlsx_count += extension == ".xlsx"
        xlsm_count += extension == ".xlsm"
        for field in ("input_sha256", "roundtrip_sha256"):
            if not isinstance(item.get(field), str) or not SHA256.fullmatch(item[field]):
                raise ValueError(f"files[{index}].{field} must be SHA-256")
        for field in ("opened", "roundtrip_opened", "saved_copy", "critical_parts_equal"):
            if item.get(field) is not True:
                raise ValueError(f"files[{index}].{field} must be true")
        names = item.get("worksheet_names")
        if not isinstance(names, list) or not names or not all(isinstance(name, str) and name for name in names):
            raise ValueError(f"files[{index}].worksheet_names must be non-empty strings")
        part_names = item.get("critical_part_names")
        if (
            not isinstance(part_names, list)
            or not all(isinstance(part, str) and part.startswith("xl/") for part in part_names)
            or part_names != sorted(set(part_names))
            or item.get("critical_part_count") != len(part_names)
        ):
            raise ValueError(f"files[{index}] critical part inventory is invalid")
        if extension == ".xlsm":
            critical_names.update(part_names)
    if xlsx_count < 1 or xlsm_count < 1:
        raise ValueError("evidence must include both .xlsx and .xlsm workbooks")
    required_macro_parts = {"xl/vbaProject.bin", "xl/vbaProjectSignature.bin"}
    if not required_macro_parts <= critical_names:
        raise ValueError(".xlsm evidence must include a VBA project and digital signature")
    if not any(name.startswith(("xl/activeX/", "xl/ctrlProps/")) or name.endswith(".vml") for name in critical_names):
        raise ValueError(".xlsm evidence must include an ActiveX or VML control part")
    summary = payload.get("summary")
    if not isinstance(summary, dict):
        raise ValueError("summary must be an object")
    _exact_keys(summary, {"total", "xlsx", "xlsm", "all_checks_passed"}, "summary")
    if summary != {"total": len(files), "xlsx": xlsx_count, "xlsm": xlsm_count, "all_checks_passed": True}:
        raise ValueError("summary does not match file results")
    return payload


def validate_human_review(payload: Any, automated_sha256: str) -> dict[str, Any]:
    if not isinstance(payload, dict):
        raise ValueError("human review root must be an object")
    _exact_keys(
        payload,
        {"schema_version", "automated_evidence_sha256", "reviewer", "reviewed_at", "decision", "checks", "notes"},
        "human review",
    )
    if payload["schema_version"] != "1.0":
        raise ValueError("human review schema_version must be 1.0")
    if payload.get("automated_evidence_sha256") != automated_sha256:
        raise ValueError("human review does not bind to the automated evidence SHA-256")
    if not isinstance(payload.get("reviewer"), str) or len(payload["reviewer"].strip()) < 2:
        raise ValueError("reviewer must identify the human approver")
    _timestamp(payload.get("reviewed_at"), "reviewed_at")
    if payload.get("decision") != "approved":
        raise ValueError("human review decision must be approved")
    checks = payload.get("checks")
    if not isinstance(checks, dict) or set(checks) != REQUIRED_REVIEW_CHECKS:
        raise ValueError(f"checks must contain exactly {sorted(REQUIRED_REVIEW_CHECKS)}")
    if not all(value is True for value in checks.values()):
        raise ValueError("every human review check must be true")
    if not isinstance(payload.get("notes"), str):
        raise ValueError("notes must be a string")
    return payload


def verify_acceptance(automated_path: Path, review_path: Path) -> dict[str, Any]:
    automated_bytes = automated_path.read_bytes()
    automated = validate_automated_evidence(json.loads(automated_bytes.decode("utf-8")))
    automated_sha256 = hashlib.sha256(automated_bytes).hexdigest()
    review = validate_human_review(json.loads(review_path.read_text(encoding="utf-8")), automated_sha256)
    return {
        "status": "approved",
        "automated_evidence_sha256": automated_sha256,
        "excel_version": automated["excel_version"],
        "workbooks": automated["summary"]["total"],
        "reviewer": review["reviewer"],
        "reviewed_at": review["reviewed_at"],
    }


def _exact_keys(payload: dict[str, Any], expected: set[str], label: str) -> None:
    missing = expected - payload.keys()
    unknown = payload.keys() - expected
    if missing:
        raise ValueError(f"{label} missing fields: {sorted(missing)}")
    if unknown:
        raise ValueError(f"{label} unknown fields: {sorted(unknown)}")


def _timestamp(value: Any, label: str) -> datetime:
    if not isinstance(value, str):
        raise ValueError(f"{label} must be an RFC 3339 timestamp")
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as exc:
        raise ValueError(f"{label} must be an RFC 3339 timestamp") from exc
    if parsed.tzinfo is None:
        raise ValueError(f"{label} must include a timezone")
    return parsed


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Verify Microsoft Excel desktop acceptance evidence")
    parser.add_argument("automated_evidence", type=Path)
    parser.add_argument("human_review", type=Path)
    args = parser.parse_args(argv)
    try:
        result = verify_acceptance(args.automated_evidence, args.human_review)
    except (OSError, UnicodeDecodeError, json.JSONDecodeError, ValueError) as exc:
        print(f"EXCEL_ACCEPTANCE_REJECTED: {exc}", file=sys.stderr)
        return 2
    print(json.dumps(result, ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
