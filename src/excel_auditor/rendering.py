from __future__ import annotations

import json
import re
import subprocess
import shutil
from abc import ABC, abstractmethod
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

from .engine import ComparisonResult
from .models import Difference, DifferenceType, RuleSet
from .workbook import WorkbookSnapshot, sha256_file


@dataclass
class RenderResult:
    output_path: Path
    sha256: str
    warnings: list[str] = field(default_factory=list)
    operation_results: list[dict[str, Any]] = field(default_factory=list)


class ExcelRenderer(ABC):
    @abstractmethod
    def render(self, source: Path, destination: Path, workbook: WorkbookSnapshot, rules: RuleSet, comparison: ComparisonResult, report_payload: dict[str, Any]) -> RenderResult:
        raise NotImplementedError


class OpenPyxlDevelopmentRenderer(ExcelRenderer):
    """Conservative development fallback; structural edits are refused on risky sheets."""

    def render(self, source: Path, destination: Path, workbook: WorkbookSnapshot, rules: RuleSet, comparison: ComparisonResult, report_payload: dict[str, Any]) -> RenderResult:
        if source.suffix.lower() == ".xlsm":
            raise RuntimeError("RENDER_FAILED: development renderer does not modify macro-enabled workbooks")
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)
        destination.chmod(0o600)
        book = load_workbook(destination, data_only=False, keep_links=False)
        warnings: list[str] = []
        color_by_action = {
            "mark_red": rules.colors.extra,
            "mark_row_red": rules.colors.extra,
            "mark_yellow": rules.colors.mismatch,
            "mark_orange": rules.colors.invalid,
            "mark_purple": rules.colors.ambiguous,
            "mark_row_purple": rules.colors.ambiguous,
        }
        for difference in comparison.differences:
            if difference.sheet_name not in book.sheetnames:
                continue
            sheet = book[difference.sheet_name]
            color = color_by_action.get(difference.render_action)
            if not color:
                continue
            fill = PatternFill("solid", fgColor=color)
            if difference.cell:
                cell = sheet[difference.cell]
                cell.fill = fill
                cell.comment = Comment(_comment(difference), "Excel Auditor")
            elif difference.excel_row:
                for cell in sheet[difference.excel_row]:
                    cell.fill = fill
                sheet.cell(difference.excel_row, 1).comment = Comment(_comment(difference), "Excel Auditor")
        for repair in comparison.repairs:
            if repair.type != "set_cell" or repair.sheet_name not in book.sheetnames or not repair.cell:
                continue
            cell = book[repair.sheet_name][repair.cell]
            _set_safe_value(cell, repair.value)
            cell.fill = PatternFill("solid", fgColor=rules.colors.inserted)
            cell.comment = Comment(f"自动修复；规则：{repair.rule_id}", "Excel Auditor")
        for sheet_rule in rules.sheets:
            if sheet_rule.name not in book.sheetnames:
                continue
            sheet = book[sheet_rule.name]
            risky = workbook.sheets[sheet_rule.name].risky_features
            missing = [item for item in comparison.differences if item.sheet_id == sheet_rule.id and item.type == DifferenceType.MISSING_HEADER and item.render_action == "insert_and_mark_green"]
            if missing and risky:
                warnings.append(f"{sheet_rule.name}: development renderer refused column insertion because of {', '.join(risky)}")
                continue
            _insert_missing_columns(sheet, sheet_rule, missing, rules.colors.inserted)
        for repair in comparison.repairs:
            if repair.sheet_name not in book.sheetnames:
                continue
            sheet = book[repair.sheet_name]
            if repair.type == "set_field" and repair.excel_row and repair.canonical_field:
                positions = _canonical_positions(sheet, next(item for item in rules.sheets if item.id == repair.sheet_id))
                if repair.canonical_field not in positions:
                    raise RuntimeError(f"RENDER_FAILED: repaired field has no final column: {repair.canonical_field}")
                cell = sheet.cell(repair.excel_row, positions[repair.canonical_field])
                _set_safe_value(cell, repair.value)
                cell.fill = PatternFill("solid", fgColor=rules.colors.inserted)
                cell.comment = Comment(f"自动修复；规则：{repair.rule_id}", "Excel Auditor")
            elif repair.type == "append_record" and repair.excel_row and repair.values is not None:
                sheet_rule = next(item for item in rules.sheets if item.id == repair.sheet_id)
                positions = _canonical_positions(sheet, sheet_rule)
                for field, value in repair.values.items():
                    if field not in positions:
                        continue
                    cell = sheet.cell(repair.excel_row, positions[field])
                    _set_safe_value(cell, value)
                    cell.fill = PatternFill("solid", fgColor=rules.colors.inserted)
                sheet.cell(repair.excel_row, 1).comment = Comment(f"自动追加标准记录；规则：{repair.rule_id}", "Excel Auditor")
        if "核验报告" in book.sheetnames:
            del book["核验报告"]
        report_sheet = book.create_sheet("核验报告")
        report_sheet.append(["项目", "值"])
        summary = report_payload.get("summary", {})
        for key, value in summary.items():
            report_sheet.append([key, value])
        report_sheet.append([])
        report_sheet.append(["类型", "工作表", "单元格", "业务主键", "说明"])
        for item in report_payload.get("differences", []):
            report_sheet.append([item["type"], item["sheet_name"], item.get("cell"), json.dumps(item.get("business_key"), ensure_ascii=False), item["message"]])
        report_sheet.sheet_state = "visible"
        book.save(destination)
        book.close()
        load_workbook(destination, read_only=True, data_only=False).close()
        applied = [{"difference_id": item.difference_id, "status": "applied"} for item in comparison.differences if item.repair_status == "planned"]
        return RenderResult(destination, sha256_file(destination), warnings, applied)


class DotNetOpenXmlRenderer(ExcelRenderer):
    def __init__(self, command: Path) -> None:
        if not command.is_file():
            raise FileNotFoundError(f"renderer command not found: {command}")
        self.command = command

    def self_check(self) -> str:
        completed = subprocess.run(
            [str(self.command), "--version"],
            capture_output=True,
            text=True,
            timeout=10,
            check=False,
        )
        if completed.returncode != 0:
            raise RuntimeError("renderer self-check failed")
        try:
            payload = json.loads(completed.stdout)
        except json.JSONDecodeError as exc:
            raise RuntimeError("renderer self-check returned invalid JSON") from exc
        version = payload.get("version")
        if payload.get("name") != "ExcelRenderer" or not isinstance(version, str) or not re.fullmatch(r"\d+\.\d+\.\d+", version):
            raise RuntimeError("renderer self-check returned an invalid identity or version")
        return version

    def render(self, source: Path, destination: Path, workbook: WorkbookSnapshot, rules: RuleSet, comparison: ComparisonResult, report_payload: dict[str, Any]) -> RenderResult:
        manifest = destination.parent / "render-manifest.private.json"
        completed = subprocess.run(
            [str(self.command), "--input", str(source), "--output", str(destination), "--manifest", str(manifest)],
            capture_output=True,
            text=True,
            timeout=rules.workbook.processing_timeout_seconds,
            check=False,
        )
        if completed.returncode != 0:
            raise RuntimeError(f"RENDER_FAILED: {completed.stderr[:1000]}")
        try:
            payload = json.loads(completed.stdout)
        except json.JSONDecodeError as exc:
            raise RuntimeError("RENDER_FAILED: renderer returned invalid JSON") from exc
        digest = sha256_file(destination)
        if not payload.get("success") or payload.get("output_sha256") != digest:
            raise RuntimeError("OUTPUT_VERIFICATION_FAILED: renderer hash mismatch")
        operation_results = payload.get("operation_results", [])
        if not isinstance(operation_results, list):
            raise RuntimeError("RENDER_FAILED: renderer returned invalid operation results")
        return RenderResult(destination, digest, operation_results=operation_results)


def _insert_missing_columns(sheet: Any, sheet_rule: Any, missing: list[Difference], color: str) -> None:
    missing_names = {item.canonical_field for item in missing}
    if not missing_names:
        return
    header_row = sheet_rule.header.row
    known_positions: dict[str, int] = {}
    normalized_headers = {str(sheet.cell(header_row, col).value).strip(): col for col in range(1, sheet.max_column + 1)}
    for rule in sheet_rule.columns:
        for candidate in [rule.title, rule.name, *rule.aliases]:
            if candidate in normalized_headers:
                known_positions[rule.name] = normalized_headers[candidate]
                break
    plans: list[tuple[int, int, Any]] = []
    for index, rule in enumerate(sheet_rule.columns):
        if rule.name not in missing_names:
            continue
        previous = [known_positions[previous_rule.name] for previous_rule in sheet_rule.columns[:index] if previous_rule.name in known_positions]
        before = previous[-1] + 1 if previous else 1
        plans.append((before, index, rule))
    for before, _index, rule in sorted(plans, key=lambda item: (item[0], item[1]), reverse=True):
        sheet.insert_cols(before)
        cell = sheet.cell(header_row, before, rule.title)
        if before > 1:
            source = sheet.cell(header_row, before - 1)
            cell.font, cell.border, cell.alignment, cell.number_format, cell.protection = copy(source.font), copy(source.border), copy(source.alignment), source.number_format, copy(source.protection)
        cell.fill = PatternFill("solid", fgColor=color)
        cell.comment = Comment(f"缺失表头；由规则 {sheet_rule.id} 插入", "Excel Auditor")


def _canonical_positions(sheet: Any, sheet_rule: Any) -> dict[str, int]:
    positions: dict[str, int] = {}
    normalized_headers = {str(sheet.cell(sheet_rule.header.row, col).value).strip(): col for col in range(1, sheet.max_column + 1)}
    for rule in sheet_rule.columns:
        for candidate in [rule.title, rule.name, *rule.aliases]:
            if candidate in normalized_headers:
                positions[rule.name] = normalized_headers[candidate]
                break
    return positions


def _set_safe_value(cell: Any, value: Any) -> None:
    cell.value = value
    if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):
        cell.data_type = "s"


def _comment(item: Difference) -> str:
    parts = [item.message]
    if item.excel_raw_value is not None:
        parts.append(f"Excel值：{item.excel_raw_value}")
    if item.standard_raw_value is not None:
        parts.append(f"标准值：{item.standard_raw_value}")
    if item.rule_id:
        parts.append(f"规则：{item.rule_id}")
    return "；".join(parts)[:32000]
