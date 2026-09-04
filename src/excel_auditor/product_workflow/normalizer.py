from __future__ import annotations

import json
import re
from collections import defaultdict
from collections.abc import Callable
from itertools import product
from typing import Any

from ..ids import new_ulid
from ..models import ColumnRule, FieldType, RuleSet
from ..normalization import (
    ParsedValue,
    excel_datetime_write_safe,
    excel_numeric_write_safe,
    normalized_uniqueness_key,
    parse_excel_value,
    parse_value,
)
from ..validators import run_validator
from ..workbook import WorkbookSnapshot, locate_header_row
from ..spill import SpillableSequence
from .catalog import CatalogAdapter
from .category import resolve_categories
from .mapping import map_product_headers
from .models import (
    CatalogFieldDefinition,
    CatalogFieldSource,
    CategoryCatalogSnapshot,
    CategoryResolution,
    NormalizedCategorySheet,
    ProductCellIssue,
    ProductNormalizationResult,
    ReviewItem,
    UnresolvedProductRow,
)
from .schema_builder import build_dynamic_schema


def normalize_product_workbook(
    workbook: WorkbookSnapshot,
    rules: RuleSet,
    catalog: CatalogAdapter,
    *,
    confirmed_aliases: dict[str, str] | None = None,
    confirmed_aliases_by_category: dict[str, dict[str, str]] | None = None,
    category_overrides: dict[int, str] | None = None,
    forced_extra_columns: set[int] | None = None,
    forced_extra_columns_by_category: dict[str, set[int]] | None = None,
    checkpoint: Callable[[], None] | None = None,
) -> ProductNormalizationResult:
    check = checkpoint or (lambda: None)
    check()
    config = rules.product_workflow
    if config is None:
        raise ValueError("product_workflow is not configured for this rule set")
    sheet_rule = next(sheet for sheet in rules.sheets if sheet.id == config.sheet_id)
    physical_names = {sheet_rule.name.casefold(), *(alias.casefold() for alias in sheet_rule.aliases)}
    matches = [snapshot for name, snapshot in workbook.sheets.items() if name.casefold() in physical_names]
    if len(matches) != 1:
        raise ValueError(
            f"PRODUCT_WORKFLOW_INPUT_INVALID: expected exactly one worksheet for {sheet_rule.id!r}, found {len(matches)}"
        )
    sheet = matches[0]
    header_row, header_problem = locate_header_row(sheet_rule, sheet)
    if header_problem:
        raise ValueError(f"PRODUCT_WORKFLOW_INPUT_INVALID: {header_problem}")
    header_values = next((values for row_number, values in sheet.rows if row_number == header_row), None)
    if header_values is None:
        raise ValueError("PRODUCT_WORKFLOW_INPUT_INVALID: header row is missing")
    headers = ["" if value is None else str(value) for value in header_values]
    data_start = sheet_rule.data_region.start_row or header_row + 1
    fixed_targets = [_fixed_target(column, index) for index, column in enumerate(sheet_rule.columns)]
    fixed_ids = {field.field_id for field in fixed_targets}
    fixed_confirmed_aliases = {
        alias: field_id
        for alias, field_id in (confirmed_aliases or {}).items()
        if field_id in fixed_ids
    }
    fixed_mappings = map_product_headers(
        headers,
        fixed_targets,
        confirmed_aliases=fixed_confirmed_aliases,
        forced_extra_columns=forced_extra_columns,
        fuzzy_threshold=config.mapping_fuzzy_threshold,
    )
    fixed_by_id = {
        mapping.field_id: mapping
        for mapping in fixed_mappings
        if mapping.status == "accepted" and mapping.field_id is not None
    }
    category_rows: list[dict[str, Any]] = []
    source_row_numbers: list[int] = []
    source_row_positions: dict[int, int] = {}
    for position, (row_number, values) in enumerate(sheet.rows):
        if row_number < data_start or not any(value not in {None, ""} for value in values):
            continue
        if len(source_row_numbers) % 1_000 == 0:
            check()
        source_row_numbers.append(row_number)
        source_row_positions[row_number] = position
        row: dict[str, Any] = {}
        for column in sheet_rule.columns:
            mapping = fixed_by_id.get(column.name)
            raw = _value_at(values, mapping.physical_column if mapping else None)
            parsed = parse_excel_value(raw, column, workbook.excel_epoch)
            row[column.name] = parsed.normalized if parsed.valid else raw
        category_rows.append(row)

    categories = catalog.list_categories()
    frozen_category_snapshot = getattr(catalog, "category_snapshot", None)
    category_catalog_snapshot = frozen_category_snapshot or CategoryCatalogSnapshot.create(
        snapshot_id=new_ulid("pcat_"),
        connection_id=str(getattr(catalog, "connection_id", config.catalog_connection_id)),
        categories=categories,
        source_metadata={"purpose": "product_category_resolution"},
    )
    resolutions = resolve_categories(
        category_rows,
        categories,
        source_field=config.category.source_field,
        id_field=config.category.id_field,
        first_excel_row=data_start,
        excel_rows=source_row_numbers,
        fuzzy_threshold=config.minimum_category_confidence,
        candidate_score_margin=config.category_candidate_score_margin,
    )
    check()
    categories_by_id = {category.category_id: category for category in categories if category.active}
    if category_overrides:
        row_numbers = set(source_row_numbers)
        unknown_rows = set(category_overrides) - row_numbers
        unknown_categories = set(category_overrides.values()) - set(categories_by_id)
        if unknown_rows:
            raise ValueError(f"category override references unknown Excel rows: {sorted(unknown_rows)}")
        if unknown_categories:
            raise ValueError(f"category override references unknown platform categories: {sorted(unknown_categories)}")
        resolutions = [
            CategoryResolution(
                excel_row=resolution.excel_row,
                raw_category_id=resolution.raw_category_id,
                raw_category=resolution.raw_category,
                category_id=categories_by_id[category_overrides[resolution.excel_row]].category_id,
                category_name=categories_by_id[category_overrides[resolution.excel_row]].name,
                status="resolved",
                match_type="confirmed",
                confidence=100,
            )
            if resolution.excel_row in category_overrides
            else resolution
            for resolution in resolutions
        ]
    resolution_by_row = {resolution.excel_row: resolution for resolution in resolutions}
    grouped: dict[str, list[int]] = defaultdict(list)
    unresolved_rows: list[UnresolvedProductRow] = []
    review_items: list[ReviewItem] = []
    for row_number in source_row_numbers:
        resolution = resolution_by_row[row_number]
        if resolution.status == "resolved" and resolution.category_id is not None:
            grouped[resolution.category_id].append(row_number)
        else:
            values = sheet.rows[source_row_positions[row_number]][1]
            unresolved_rows.append(UnresolvedProductRow(
                excel_row=row_number,
                values=list(values),
                category_resolution=resolution,
            ))
            review_items.append(_category_review(resolution))

    category_by_id = {category.category_id: category for category in categories}
    schema_category_ids = set(grouped)
    schema_category_ids.update(
        candidate.field_id
        for resolution in resolutions
        if resolution.status == "manual_review"
        for candidate in resolution.candidates
    )
    snapshots = []
    for category_id in sorted(schema_category_ids):
        check()
        snapshots.append(catalog.fetch_schema(category_id))
    snapshots_by_category = {snapshot.category_id: snapshot for snapshot in snapshots}
    sheets: list[NormalizedCategorySheet] = []
    issues: list[ProductCellIssue] = []
    used_sheet_names: set[str] = set()
    row_spill_threshold = max(
        1_000,
        rules.workbook.max_in_memory_cells // max(1, len(headers)),
    )
    for category_id in sorted(grouped):
        category = category_by_id[category_id]
        snapshot = snapshots_by_category[category_id]
        targets = [*fixed_targets, *snapshot.fields]
        target_ids = {field.field_id for field in targets}
        category_confirmed_aliases = {
            alias: field_id
            for alias, field_id in (confirmed_aliases or {}).items()
            if field_id in target_ids
        }
        for alias, field_id in (confirmed_aliases_by_category or {}).get(category_id, {}).items():
            if field_id not in target_ids:
                raise ValueError(
                    f"confirmed category mapping targets unknown field {field_id!r} for category {category_id!r}"
                )
            category_confirmed_aliases[alias] = field_id
        mappings = map_product_headers(
            headers,
            targets,
            confirmed_aliases=category_confirmed_aliases,
            forced_extra_columns={
                *(forced_extra_columns or set()),
                *(forced_extra_columns_by_category or {}).get(category_id, set()),
            },
            fuzzy_threshold=config.mapping_fuzzy_threshold,
        )
        plan = build_dynamic_schema(
            category_id=category_id,
            category_name=category.name,
            fixed_columns=sheet_rule.columns,
            platform_fields=snapshot.fields,
            headers=headers,
            mappings=mappings,
        )
        plan.worksheet_name = _unique_sheet_name(plan.worksheet_name, category_id, used_sheet_names)
        review_items.extend(
            item.model_copy(update={"key": f"category:{category_id}:{item.key}"})
            for item in plan.review_items
        )
        output_rows = SpillableSequence[dict[str, Any]](row_spill_threshold)
        sku_rows = SpillableSequence[dict[str, Any]](row_spill_threshold)
        sku_source_excel_rows: list[int] = []
        specification_plans = [
            planned
            for planned in plan.fields
            if planned.field.source == CatalogFieldSource.PLATFORM_SPECIFICATION
        ]
        for row_index, row_number in enumerate(grouped[category_id]):
            if row_index % 1_000 == 0:
                check()
            values = sheet.rows[source_row_positions[row_number]][1]
            output: dict[str, Any] = {}
            for planned in plan.fields:
                field = planned.field
                raw = _value_at(values, planned.source_column)
                normalized, error, write_unsafe = _normalize_field(raw, field, sheet_rule.columns, workbook)
                if error is not None:
                    output[field.field_id] = raw
                    issues.append(ProductCellIssue(
                        issue_type=(
                            "excel_write_unsafe"
                            if write_unsafe
                            else "invalid_value" if raw not in {None, ""} else "required_missing"
                        ),
                        excel_row=row_number,
                        category_id=category_id,
                        field_id=field.field_id,
                        physical_column=planned.source_column,
                        raw_value=raw,
                        message=error,
                        color=(
                            config.output.required_missing_color
                            if raw in {None, ""}
                            else config.output.invalid_value_color
                        ),
                    ))
                else:
                    output[field.field_id] = normalized
            output_rows.append(output)
            should_build_sku = (
                config.output.sku_sheet_mode == "always"
                or config.output.sku_sheet_mode == "when_present" and bool(specification_plans)
            )
            if should_build_sku:
                option_sets = [_sku_values(output.get(planned.field.field_id), planned.field) for planned in specification_plans]
                combination_count = 1
                for options in option_sets:
                    combination_count *= len(options)
                if combination_count > config.output.max_sku_combinations_per_product:
                    limiting_field = next((planned for planned in specification_plans if planned.field.multiple), specification_plans[0])
                    issues.append(ProductCellIssue(
                        issue_type="sku_combination_limit",
                        excel_row=row_number,
                        category_id=category_id,
                        field_id=limiting_field.field.field_id,
                        physical_column=limiting_field.ordinal,
                        raw_value=output.get(limiting_field.field.field_id),
                        message=(
                            f"SKU combinations {combination_count} exceed configured limit "
                            f"{config.output.max_sku_combinations_per_product}"
                        ),
                        color=config.output.ambiguous_color,
                    ))
                else:
                    fixed_values = {
                        planned.field.field_id: output[planned.field.field_id]
                        for planned in plan.fields
                        if planned.field.source == CatalogFieldSource.FIXED
                    }
                    for combination in product(*option_sets):
                        sku_row = dict(fixed_values)
                        sku_row.update({
                            planned.field.field_id: value
                            for planned, value in zip(specification_plans, combination, strict=True)
                        })
                        sku_rows.append(sku_row)
                        sku_source_excel_rows.append(row_number)
        if config.output.sku_sheet_mode == "disabled":
            sku_rows.close()
            sku_rows = []
        sheets.append(NormalizedCategorySheet(
            category_id=category_id,
            category_name=category.name,
            worksheet_name=plan.worksheet_name,
            plan=plan,
            source_excel_rows=list(grouped[category_id]),
            rows=output_rows.finish(),
            sku_rows=sku_rows.finish() if isinstance(sku_rows, SpillableSequence) else sku_rows,
            sku_source_excel_rows=sku_source_excel_rows,
        ))

    for mapping in fixed_mappings:
        if mapping.status == "manual_review":
            issues.append(ProductCellIssue(
                issue_type="ambiguous_mapping",
                excel_row=header_row,
                field_id=mapping.field_id,
                physical_column=mapping.physical_column,
                raw_value=mapping.raw_header,
                message=f"字段 {mapping.raw_header!r} 需要人工确认映射",
                color=config.output.ambiguous_color,
            ))
    check()
    issues.extend(_record_quality_issues(
        sheets,
        sheet_rule,
        workbook,
        config.output.invalid_value_color,
        existing_issues=issues,
        checkpoint=check,
    ))
    return ProductNormalizationResult(
        category_catalog_snapshot=category_catalog_snapshot,
        catalog_snapshots=snapshots,
        category_sheets=sheets,
        source_headers=headers,
        unresolved_rows=unresolved_rows,
        review_items=review_items,
        issues=issues,
        requires_manual_review=bool(
            review_items
            or unresolved_rows
            or any(issue.issue_type == "excel_write_unsafe" for issue in issues)
            or any(issue.issue_type == "sku_combination_limit" for issue in issues)
        ),
        merchant_extra_header_color=config.output.merchant_extra_header_color,
    )


def _record_quality_issues(
    sheets: list[NormalizedCategorySheet],
    sheet_rule: Any,
    workbook: WorkbookSnapshot,
    color: str,
    *,
    existing_issues: list[ProductCellIssue],
    checkpoint: Callable[[], None],
) -> list[ProductCellIssue]:
    """Apply fixed-field uniqueness, primary-key, and cross-field contracts globally."""
    rules_by_name = {column.name: column for column in sheet_rule.columns}
    unique_occurrences: dict[str, dict[tuple[str, Any], list[tuple[int, str, int]]]] = {
        column.name: defaultdict(list)
        for column in sheet_rule.columns
        if column.validation.unique
    }
    primary_occurrences: dict[tuple[Any, ...], list[tuple[int, str, int]]] = defaultdict(list)
    platform_unique_occurrences: dict[
        tuple[str, str],
        dict[tuple[str, Any], list[tuple[int, str, int]]],
    ] = {}
    invalid_cells = {
        (issue.excel_row, issue.category_id, issue.field_id)
        for issue in existing_issues
        if issue.issue_type in {"required_missing", "invalid_value"}
    }
    parsed_records: list[tuple[int, str, dict[str, ParsedValue], dict[str, int]]] = []
    for category_sheet in sheets:
        ordinals = {
            planned.field.field_id: planned.ordinal
            for planned in category_sheet.plan.fields
            if planned.field.source == CatalogFieldSource.FIXED
        }
        platform_unique_fields = {
            planned.field.field_id: planned.field
            for planned in category_sheet.plan.fields
            if planned.field.source != CatalogFieldSource.FIXED and planned.field.validation.unique
        }
        for field_id in platform_unique_fields:
            platform_unique_occurrences.setdefault(
                (category_sheet.category_id, field_id),
                defaultdict(list),
            )
        for row_index, (excel_row, output) in enumerate(
            zip(category_sheet.source_excel_rows, category_sheet.rows, strict=True)
        ):
            if row_index % 1_000 == 0:
                checkpoint()
            parsed = {
                name: parse_excel_value(output.get(name), rule, workbook.excel_epoch)
                for name, rule in rules_by_name.items()
            }
            parsed_records.append((excel_row, category_sheet.category_id, parsed, ordinals))
            if sheet_rule.primary_key_mode == "fields":
                key_values = [parsed[name] for name in sheet_rule.primary_key]
                if all(value.valid and value.normalized is not None for value in key_values):
                    key = tuple(
                        normalized_uniqueness_key(value, rules_by_name[name])
                        for name, value in zip(sheet_rule.primary_key, key_values, strict=True)
                    )
                    primary_occurrences[key].append(
                        (excel_row, category_sheet.category_id, ordinals.get(sheet_rule.primary_key[0], 1))
                    )
            for name, occurrences in unique_occurrences.items():
                value = parsed[name]
                key = normalized_uniqueness_key(value, rules_by_name[name]) if value.valid else None
                if key is not None:
                    occurrences[key].append((excel_row, category_sheet.category_id, ordinals.get(name, 1)))
            for field_id, field in platform_unique_fields.items():
                if (excel_row, category_sheet.category_id, field_id) in invalid_cells:
                    continue
                value = output.get(field_id)
                key = _platform_uniqueness_key(value, field)
                if key is not None:
                    ordinal = next(
                        planned.ordinal
                        for planned in category_sheet.plan.fields
                        if planned.field.field_id == field_id
                    )
                    platform_unique_occurrences[(category_sheet.category_id, field_id)][key].append(
                        (excel_row, category_sheet.category_id, ordinal)
                    )

    result: list[ProductCellIssue] = []
    primary_field = sheet_rule.primary_key[0] if sheet_rule.primary_key else None
    if primary_field is not None:
        for occurrences in primary_occurrences.values():
            if len(occurrences) <= 1:
                continue
            for excel_row, category_id, ordinal in occurrences:
                result.append(ProductCellIssue(
                    issue_type="duplicate_primary_key",
                    excel_row=excel_row,
                    category_id=category_id,
                    field_id=primary_field,
                    physical_column=ordinal,
                    message="商品主键重复，不允许作为唯一商品记录",
                    color=color,
                ))
    for name, values in unique_occurrences.items():
        for occurrences in values.values():
            if len(occurrences) <= 1:
                continue
            for excel_row, category_id, ordinal in occurrences:
                result.append(ProductCellIssue(
                    issue_type="unique_value",
                    excel_row=excel_row,
                    category_id=category_id,
                    field_id=name,
                    physical_column=ordinal,
                    message="字段值不唯一",
                    color=color,
                ))
    for (_category_id, field_id), values in platform_unique_occurrences.items():
        for occurrences in values.values():
            if len(occurrences) <= 1:
                continue
            for excel_row, category_id, ordinal in occurrences:
                result.append(ProductCellIssue(
                    issue_type="unique_value",
                    excel_row=excel_row,
                    category_id=category_id,
                    field_id=field_id,
                    physical_column=ordinal,
                    message="platform field value must be unique within its category",
                    color=color,
                ))

    for excel_row, category_id, parsed, ordinals in parsed_records:
        for cross_rule in sheet_rule.cross_field_rules:
            params = dict(cross_rule.params)
            when = params.get("when_field")
            if cross_rule.validator == "conditional_required" and when in rules_by_name:
                params["equals"] = parse_value(params.get("equals"), rules_by_name[when]).normalized
            outcome = run_validator(cross_rule.validator, parsed, params)
            if outcome is None:
                continue
            target, message = outcome
            if target not in rules_by_name:
                continue
            result.append(ProductCellIssue(
                issue_type="cross_field_error",
                excel_row=excel_row,
                category_id=category_id,
                field_id=target,
                physical_column=ordinals.get(target, 1),
                message=message,
                color=color,
            ))
    return result


def _platform_uniqueness_key(
    value: Any,
    field: CatalogFieldDefinition,
) -> tuple[str, Any] | None:
    if value is None:
        return None
    if isinstance(value, (dict, list)):
        canonical: Any = json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)
    elif isinstance(value, set):
        canonical = tuple(sorted(str(item) for item in value))
    else:
        canonical = value
    return field.field_type.value, canonical


def _sku_values(value: Any, field: CatalogFieldDefinition) -> list[Any]:
    if not field.multiple:
        return [value]
    if value is None:
        return [None]
    if isinstance(value, (list, tuple, set)):
        values = [item for item in value if item not in {None, ""}]
    else:
        values = [item.strip() for item in str(value).split("、") if item.strip()]
    return list(dict.fromkeys(values)) or [None]


def _fixed_target(column: ColumnRule, order: int) -> CatalogFieldDefinition:
    return CatalogFieldDefinition(
        field_id=column.name,
        title=column.title,
        aliases=column.aliases,
        source=CatalogFieldSource.FIXED,
        field_type=column.type,
        required=column.required,
        multiple=column.type == FieldType.SET,
        display_order=order,
        enum_values=column.enum_values,
        number_format=column.format,
        timezone=column.compare.timezone,
        validation=column.validation,
    )


def _normalize_field(
    raw: Any,
    field: CatalogFieldDefinition,
    fixed_columns: list[ColumnRule],
    workbook: WorkbookSnapshot,
) -> tuple[Any, str | None, bool]:
    if raw in {None, ""}:
        if field.required or not field.validation.nullable:
            return None, "必填字段为空", False
        return None, None, False
    if field.source == CatalogFieldSource.MERCHANT_EXTRA:
        return raw, None, False
    if field.source == CatalogFieldSource.FIXED:
        rule = next(column for column in fixed_columns if column.name == field.field_id)
    elif field.multiple:
        if isinstance(raw, (list, tuple, set)):
            items = [str(value).strip() for value in raw if str(value).strip()]
        else:
            items = [value.strip() for value in re.split(r"[,，;；、|]", str(raw)) if value.strip()]
        unique = list(dict.fromkeys(items))
        if field.enum_values:
            invalid = [value for value in unique if value not in field.enum_values]
            if invalid:
                return raw, f"值不在平台枚举范围：{invalid}", False
        joined = "、".join(unique)
        return joined, _validate_scalar(joined, field), False
    else:
        compare = (
            {"mode": "datetime", "timezone": field.timezone}
            if field.field_type == FieldType.DATETIME
            else {}
        )
        rule = ColumnRule(
            name=field.field_id,
            title=field.title,
            required=field.required,
            type=field.field_type,
            enum_values=field.enum_values,
            validation=field.validation,
            format=field.number_format,
            compare=compare,
        )
    parsed: ParsedValue = parse_excel_value(raw, rule, workbook.excel_epoch)
    if not parsed.valid:
        return raw, parsed.error or "字段值无效", False
    if field.field_type in {FieldType.INTEGER, FieldType.DECIMAL} and not excel_numeric_write_safe(parsed.normalized):
        return raw, "数值超出 Excel 可安全往返的 15 位有效数字或指数范围", True
    if (
        field.field_type == FieldType.DATETIME
        and not excel_datetime_write_safe(parsed.normalized, field.timezone)
    ):
        return raw, "日期时间写入 Excel 后会丢失时区或 DST 身份", True
    validation_error = _validate_scalar(parsed.normalized, field)
    return parsed.normalized, validation_error, False


def _validate_scalar(value: Any, field: CatalogFieldDefinition) -> str | None:
    config = field.validation
    if value is None:
        return "必填字段为空" if field.required or not config.nullable else None
    text = str(value)
    if config.min_length is not None and len(text) < config.min_length:
        return f"长度小于 {config.min_length}"
    if config.max_length is not None and len(text) > config.max_length:
        return f"长度大于 {config.max_length}"
    if config.regex and re.fullmatch(config.regex, text) is None:
        return "值不符合平台格式规则"
    if config.min is not None and value < config.min:
        return f"值小于最小值 {config.min}"
    if config.max is not None and value > config.max:
        return f"值大于最大值 {config.max}"
    return None


def _value_at(values: list[Any], physical_column: int | None) -> Any:
    if physical_column is None or physical_column > len(values):
        return None
    return values[physical_column - 1]


def _category_review(resolution: CategoryResolution) -> ReviewItem:
    if resolution.match_type == "id_name_conflict":
        message = (
            f"平台类目 ID {resolution.raw_category_id!r} 与商家类目名称 "
            f"{resolution.raw_category!r} 指向不同类目"
        )
    elif resolution.match_type == "invalid_id":
        message = f"平台类目 ID {resolution.raw_category_id!r} 无效，需要人工确认"
    elif resolution.status == "manual_review":
        message = f"类目 {resolution.raw_category!r} 需要人工确认"
    else:
        message = "未提供可解析的平台类目"
    return ReviewItem(
        review_type="category",
        key=f"row:{resolution.excel_row}",
        excel_row=resolution.excel_row,
        message=message,
        candidates=resolution.candidates,
    )


def _unique_sheet_name(base: str, category_id: str, used: set[str]) -> str:
    candidate = base[:31]
    if candidate.casefold() not in used:
        used.add(candidate.casefold())
        return candidate
    suffix = f"-{category_id}"[:15]
    candidate = f"{base[:31-len(suffix)]}{suffix}"
    counter = 2
    while candidate.casefold() in used:
        marker = f"-{counter}"
        candidate = f"{base[:31-len(suffix)-len(marker)]}{suffix}{marker}"
        counter += 1
    used.add(candidate.casefold())
    return candidate
