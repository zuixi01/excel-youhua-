from __future__ import annotations

import re
import json
from decimal import Decimal
from typing import Any, Sequence

import pandas as pd
import pandera.pandas as pa

from .models import RuleSet, SheetRule
from .normalization import is_formula_text, parse_value
from .validators import run_validator


class StandardDataValidator:
    def __init__(self, chunk_size: int = 50_000) -> None:
        if chunk_size < 1:
            raise ValueError("chunk_size must be positive")
        self.chunk_size = chunk_size

    def validate(self, standard: dict[str, Sequence[dict[str, Any]]], rules: RuleSet) -> None:
        record_count = sum(len(rows) for rows in standard.values())
        if record_count > rules.workbook.max_standard_records:
            raise ValueError(
                f"STANDARD_TOO_LARGE: {record_count} records exceed configured limit "
                f"{rules.workbook.max_standard_records}"
            )
        expected = {sheet.id for sheet in rules.sheets}
        unknown = set(standard) - expected
        missing = {sheet.id for sheet in rules.sheets if sheet.required} - set(standard)
        if unknown or missing:
            raise ValueError(f"STANDARD_DATA_INVALID: sheet set mismatch; missing={sorted(missing)}, unknown={sorted(unknown)}")
        for sheet in rules.sheets:
            rows = standard.get(sheet.id, [])
            typed_failures = self._validate_typed_values(rows, sheet)
            if typed_failures:
                columns = sorted(typed_failures)
                count = sum(typed_failures.values())
                raise ValueError(f"STANDARD_DATA_INVALID: {sheet.id} has {count} typed failures in {columns}")
            schema = self.compile_sheet(sheet)
            self._validate_unique_columns(rows, sheet)
            for start in range(0, len(rows), self.chunk_size):
                chunk = [
                    row for row in rows[start : start + self.chunk_size]
                    if not self._skip_empty_primary_key_row(row, sheet)
                ]
                frame = pd.DataFrame(chunk, columns=[column.name for column in sheet.columns])
                try:
                    schema.validate(frame, lazy=True)
                except pa.errors.SchemaErrors as exc:
                    count = len(exc.failure_cases)
                    columns = sorted({str(value) for value in exc.failure_cases.get("column", []) if value is not None})
                    raise ValueError(f"STANDARD_DATA_INVALID: {sheet.id} has {count} schema failures in {columns}") from exc
            self._validate_primary_key(rows, sheet)
            self._validate_cross_fields(rows, sheet)

    def compile_sheet(self, sheet: SheetRule) -> pa.DataFrameSchema:
        columns: dict[str, pa.Column] = {}
        for rule in sheet.columns:
            validation = rule.validation
            row_number_fallback = (
                sheet.primary_key_mode == "fields"
                and sheet.empty_primary_key_action == "use_row_number"
                and rule.name in sheet.primary_key
            )
            # Business types and rules are checked against parsed/normalized values
            # before Pandera sees a chunk. Pandera owns structural required/nullability
            # checks only; checking raw strings here would disagree with Excel-side
            # normalization (for example trim + uppercase + regex).
            columns[rule.name] = pa.Column(
                dtype=None,
                checks=[],
                nullable=row_number_fallback or (validation.nullable and not rule.required),
                # Cross-chunk uniqueness is checked explicitly in
                # _validate_unique_columns; Pandera would only see one chunk.
                unique=False,
                required=rule.required,
            )
        return pa.DataFrameSchema(columns, strict=False, coerce=False)

    @staticmethod
    def _validate_typed_values(rows: Sequence[dict[str, Any]], sheet: SheetRule) -> dict[str, int]:
        failures: dict[str, int] = {}
        for row in rows:
            if StandardDataValidator._skip_empty_primary_key_row(row, sheet):
                continue
            for rule in sheet.columns:
                raw = row.get(rule.name)
                parsed = parse_value(raw, rule)
                invalid = not parsed.valid
                if rule.compare.formula_mode == "formula" and is_formula_text(raw):
                    continue
                if parsed.valid and parsed.normalized is None:
                    row_number_fallback = (
                        sheet.primary_key_mode == "fields"
                        and sheet.empty_primary_key_action == "use_row_number"
                        and rule.name in sheet.primary_key
                    )
                    invalid = not row_number_fallback and (rule.required or not rule.validation.nullable)
                if parsed.valid and parsed.normalized is not None:
                    text = str(parsed.normalized)
                    minimum_length = rule.validation.min_length
                    maximum_length = rule.validation.max_length
                    invalid = invalid or (minimum_length is not None and len(text) < minimum_length)
                    invalid = invalid or (maximum_length is not None and len(text) > maximum_length)
                    invalid = invalid or (
                        rule.validation.regex is not None
                        and re.fullmatch(rule.validation.regex, text) is None
                    )
                    minimum = rule.validation.min
                    maximum = rule.validation.max
                    if minimum is not None or maximum is not None:
                        try:
                            value = Decimal(str(parsed.normalized))
                            invalid = invalid or (minimum is not None and value < minimum)
                            invalid = invalid or (maximum is not None and value > maximum)
                        except (ValueError, TypeError, ArithmeticError):
                            invalid = True
                if invalid:
                    failures[rule.name] = failures.get(rule.name, 0) + 1
        return failures

    @staticmethod
    def _validate_primary_key(rows: Sequence[dict[str, Any]], sheet: SheetRule) -> None:
        rules_by_name = {column.name: column for column in sheet.columns}
        observed: set[tuple[tuple[str, Any], ...]] = set()
        for row_index, row in enumerate(rows, start=1):
            if sheet.primary_key_mode == "row_number":
                try:
                    row_number = int(row.get(sheet.row_number_field))
                except (TypeError, ValueError) as exc:
                    raise ValueError(f"STANDARD_DATA_INVALID: {sheet.id} requires integer {sheet.row_number_field} at record {row_index}") from exc
                composite = (("row_number", row_number),)
                if row_number < 1:
                    raise ValueError(f"STANDARD_DATA_INVALID: {sheet.id} has invalid row number at record {row_index}")
                if composite in observed:
                    raise ValueError(f"STANDARD_DATA_INVALID: {sheet.id} has duplicate row number at record {row_index}")
                observed.add(composite)
                continue
            key: list[tuple[str, Any]] = []
            for name in sheet.primary_key:
                parsed = parse_value(row.get(name), rules_by_name[name])
                if not parsed.valid or parsed.normalized is None or parsed.normalized == "":
                    if sheet.empty_primary_key_action in {"skip_row", "use_row_number"}:
                        key = []
                        break
                    raise ValueError(f"STANDARD_DATA_INVALID: {sheet.id} primary key is empty or invalid at record {row_index}")
                key.append((rules_by_name[name].type.value, parsed.normalized))
            if not key and sheet.empty_primary_key_action in {"skip_row", "use_row_number"}:
                continue
            composite = tuple(key)
            if composite in observed:
                raise ValueError(f"STANDARD_DATA_INVALID: {sheet.id} has duplicate composite primary key at record {row_index}")
            observed.add(composite)

    @staticmethod
    def _validate_cross_fields(rows: Sequence[dict[str, Any]], sheet: SheetRule) -> None:
        rules_by_name = {column.name: column for column in sheet.columns}
        for row_index, row in enumerate(rows, start=1):
            if StandardDataValidator._skip_empty_primary_key_row(row, sheet):
                continue
            parsed = {name: parse_value(row.get(name), rule) for name, rule in rules_by_name.items()}
            for cross_rule in sheet.cross_field_rules:
                params = dict(cross_rule.params)
                when = params.get("when_field")
                if cross_rule.validator == "conditional_required" and when in rules_by_name:
                    params["equals"] = parse_value(params.get("equals"), rules_by_name[when]).normalized
                if run_validator(cross_rule.validator, parsed, params) is not None:
                    raise ValueError(f"STANDARD_DATA_INVALID: {sheet.id} violates {cross_rule.rule_id} at record {row_index}")

    @staticmethod
    def _validate_unique_columns(rows: Sequence[dict[str, Any]], sheet: SheetRule) -> None:
        for rule in sheet.columns:
            if not rule.validation.unique:
                continue
            observed: set[str] = set()
            for row_index, row in enumerate(rows, start=1):
                if StandardDataValidator._skip_empty_primary_key_row(row, sheet):
                    continue
                parsed = parse_value(row.get(rule.name), rule)
                if parsed.normalized is None:
                    continue
                token = json.dumps(
                    {"type": rule.type.value, "value": parsed.normalized},
                    ensure_ascii=False,
                    sort_keys=True,
                    separators=(",", ":"),
                    default=str,
                )
                if token in observed:
                    raise ValueError(f"STANDARD_DATA_INVALID: {sheet.id}.{rule.name} is not unique at record {row_index}")
                observed.add(token)

    @staticmethod
    def _skip_empty_primary_key_row(row: dict[str, Any], sheet: SheetRule) -> bool:
        if sheet.primary_key_mode != "fields" or sheet.empty_primary_key_action != "skip_row":
            return False
        for rule in sheet.columns:
            if rule.name not in sheet.primary_key:
                continue
            parsed = parse_value(row.get(rule.name), rule)
            if not parsed.valid or parsed.normalized is None or parsed.normalized == "":
                return True
        return False
