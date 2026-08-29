from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest

from excel_auditor.excel_acceptance import (
    REQUIRED_REVIEW_CHECKS,
    main,
    validate_automated_evidence,
    validate_human_review,
    verify_acceptance,
)


def _automated() -> dict:
    files = []
    for index, extension in enumerate((".xlsx", ".xlsm"), start=1):
        critical_parts = [] if extension == ".xlsx" else [
            "xl/activeX/activeX1.bin",
            "xl/vbaProject.bin",
            "xl/vbaProjectSignature.bin",
        ]
        files.append({
            "file_name": f"rendered-{index}{extension}",
            "extension": extension,
            "input_sha256": f"{index}" * 64,
            "roundtrip_sha256": f"{index + 2}" * 64,
            "opened": True,
            "roundtrip_opened": True,
            "saved_copy": True,
            "worksheet_names": ["Data", "核验报告"],
            "critical_part_count": len(critical_parts),
            "critical_part_names": critical_parts,
            "critical_parts_equal": True,
        })
    return {
        "schema_version": "1.0",
        "generated_at": "2026-08-29T14:00:00Z",
        "excel_version": "16.0",
        "macro_execution": "force_disabled",
        "files": files,
        "summary": {"total": 2, "xlsx": 1, "xlsm": 1, "all_checks_passed": True},
    }


def _review(evidence_sha256: str) -> dict:
    return {
        "schema_version": "1.0",
        "automated_evidence_sha256": evidence_sha256,
        "reviewer": "Independent Reviewer",
        "reviewed_at": "2026-08-29T15:00:00+08:00",
        "decision": "approved",
        "checks": {name: True for name in REQUIRED_REVIEW_CHECKS},
        "notes": "Evidence archived in the controlled acceptance store.",
    }


def test_excel_acceptance_binds_passing_automation_to_human_review(tmp_path):
    automated_path = tmp_path / "automated.json"
    review_path = tmp_path / "review.json"
    automated_bytes = json.dumps(_automated(), ensure_ascii=False, indent=2).encode("utf-8")
    automated_path.write_bytes(automated_bytes)
    evidence_sha = hashlib.sha256(automated_bytes).hexdigest()
    review_path.write_text(json.dumps(_review(evidence_sha), ensure_ascii=False), encoding="utf-8")

    result = verify_acceptance(automated_path, review_path)
    assert result == {
        "status": "approved",
        "automated_evidence_sha256": evidence_sha,
        "excel_version": "16.0",
        "workbooks": 2,
        "reviewer": "Independent Reviewer",
        "reviewed_at": "2026-08-29T15:00:00+08:00",
    }
    assert main([str(automated_path), str(review_path)]) == 0


@pytest.mark.parametrize(
    ("mutate", "message"),
    [
        (lambda evidence: evidence.update(macro_execution="enabled"), "force_disabled"),
        (lambda evidence: evidence["files"][1].update(critical_parts_equal=False), "critical_parts_equal"),
        (lambda evidence: evidence.update(files=evidence["files"][:1]), "at least one"),
        (lambda evidence: evidence["summary"].update(total=3), "summary does not match"),
        (
            lambda evidence: evidence["files"][1].update(
                critical_part_count=1,
                critical_part_names=["xl/vbaProject.bin"],
            ),
            "digital signature",
        ),
    ],
)
def test_automated_excel_evidence_rejects_incomplete_checks(mutate, message):
    evidence = _automated()
    mutate(evidence)
    with pytest.raises(ValueError, match=message):
        validate_automated_evidence(evidence)


def test_human_excel_review_must_match_evidence_and_be_fully_approved():
    review = _review("a" * 64)
    with pytest.raises(ValueError, match="does not bind"):
        validate_human_review(review, "b" * 64)
    review["automated_evidence_sha256"] = "b" * 64
    review["checks"]["opens_without_repair_prompt"] = False
    with pytest.raises(ValueError, match="every human review check"):
        validate_human_review(review, "b" * 64)
    review["checks"]["opens_without_repair_prompt"] = True
    review["decision"] = "pending"
    with pytest.raises(ValueError, match="must be approved"):
        validate_human_review(review, "b" * 64)


def test_excel_com_harness_forces_macros_off_and_preserves_originals():
    script = Path("tools/Invoke-ExcelDesktopAcceptance.ps1").read_text(encoding="utf-8")
    assert "$excel.AutomationSecurity = 3" in script
    assert "$excel.EnableEvents = $false" in script
    assert "$excel.AskToUpdateLinks = $false" in script
    assert "Workbooks.Open($source, 0, $true)" not in script
    assert "$workbooks.Open($source, 0, $true)" in script
    assert "SaveCopyAs($roundtrip)" in script
    assert "vbaProject(?:Signature)?" in script and "xl/activeX/" in script and "vml" in script
    assert "Evidence output already exists" in script


def test_pending_review_template_cannot_pass_validation():
    template = json.loads(Path("docs/excel-acceptance-review.example.json").read_text(encoding="utf-8"))
    with pytest.raises(ValueError):
        validate_human_review(template, "a" * 64)
