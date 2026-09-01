from __future__ import annotations

import re
from collections import defaultdict
from typing import Any

from ..models import ColumnRule, FieldType, RuleSet
from ..normalization import ParsedValue, parse_excel_value
from ..workbook import WorkbookSnapshot, locate_header_row
from .catalog import CatalogAdapter
from .category import resolve_categories
from .mapping import map_product_headers
from .models import (
    CatalogFieldDefinition,
    CatalogFieldSource,
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
) -> ProductNormalizationResult:
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
    source_rows = [
        (row_number, list(values))
        for row_number, values in sheet.rows
        if row_number >= data_start and any(value not in {None, ""} for value in values)
    ]

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
    for _row_number, values in source_rows:
        row: dict[str, Any] = {}
        for column in sheet_rule.columns:
            mapping = fixed_by_id.get(column.name)
            raw = _value_at(values, mapping.physical_column if mapping else None)
            parsed = parse_excel_value(raw, column, workbook.excel_epoch)
            row[column.name] = parsed.normalized if parsed.valid else raw
        category_rows.append(row)

    categories = catalog.list_categories()
    resolutions = resolve_categories(
        category_rows,
        categories,
        source_field=config.category.source_field,
        id_field=config.category.id_field,
        first_excel_row=data_start,
        excel_rows=[row_number for row_number, _values in source_rows],
        fuzzy_threshold=config.minimum_category_confidence,
    )
    categories_by_id = {category.category_id: category for category in categories if category.active}
    if category_overrides:
        row_numbers = {row_number for row_number, _values in source_rows}
        unknown_rows = set(category_overrides) - row_numbers
        unknown_categories = set(category_overrides.values()) - set(categories_by_id)
        if unknown_rows:
            raise ValueError(f"category override references unknown Excel rows: {sorted(unknown_rows)}")
        if unknown_categories:
            raise ValueError(f"category override references unknown platform categories: {sorted(unknown_categories)}")
        resolutions = [
            CategoryResolution(
                excel_row=resolution.excel_row,
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
    grouped: dict[str, list[tuple[int, list[Any]]]] = defaultdict(list)
    unresolved_rows: list[UnresolvedProductRow] = []
    review_items: list[ReviewItem] = []
    for row_number, values in source_rows:
        resolution = resolution_by_row[row_number]
        if resolution.status == "resolved" and resolution.category_id is not None:
            grouped[resolution.category_id].append((row_number, values))
        else:
            unresolved_rows.append(UnresolvedProductRow(
                excel_row=row_number,
                values=values,
                category_resolution=resolution,
            ))
            review_items.append(_category_review(resolution))

    category_by_id = {category.category_id: category for category in categories}
    snapshots = []
    sheets: list[NormalizedCategorySheet] = []
    issues: list[ProductCellIssue] = []
    used_sheet_names: set[str] = set()
    for category_id in sorted(grouped):
        category = category_by_id[category_id]
        snapshot = catalog.fetch_schema(category_id)
        snapshots.append(snapshot)
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
            forced_extra_columns=forced_extra_columns,
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
        output_rows: list[dict[str, Any]] = []
        sku_rows: list[dict[str, Any]] = []
        for row_number, values in grouped[category_id]:
            output: dict[str, Any] = {}
            for planned in plan.fields:
                field = planned.field
                raw = _value_at(values, planned.source_column)
                normalized, error = _normalize_field(raw, field, sheet_rule.columns, workbook)
                if error is not None:
                    output[field.field_id] = raw
                    issues.append(ProductCellIssue(
                        issue_type="invalid_value" if raw not in {None, ""} else "required_missing",
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
            specification_fields = {
                planned.field.field_id
                for planned in plan.fields
                if planned.field.source == CatalogFieldSource.PLATFORM_SPECIFICATION
            }
            if specification_fields:
                sku_rows.append({
                    planned.field.field_id: output[planned.field.field_id]
                    for planned in plan.fields
                    if planned.field.source == CatalogFieldSource.FIXED
                    or planned.field.field_id in specification_fields
                })
        if config.output.sku_sheet_mode == "disabled":
            sku_rows = []
        sheets.append(NormalizedCategorySheet(
            category_id=category_id,
            category_name=category.name,
            worksheet_name=plan.worksheet_name,
            plan=plan,
            source_excel_rows=[row_number for row_number, _values in grouped[category_id]],
            rows=output_rows,
            sku_rows=sku_rows,
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
    return ProductNormalizationResult(
        catalog_snapshots=snapshots,
        category_sheets=sheets,
        unresolved_rows=unresolved_rows,
        review_items=review_items,
        issues=issues,
        requires_manual_review=bool(review_items or unresolved_rows),
        merchant_extra_header_color=config.output.merchant_extra_header_color,
    )


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
        validation=column.validation,
    )


def _normalize_field(
    raw: Any,
    field: CatalogFieldDefinition,
    fixed_columns: list[ColumnRule],
    workbook: WorkbookSnapshot,
) -> tuple[Any, str | None]:
    if raw in {None, ""}:
        if field.required or not field.validation.nullable:
            return None, "必填字段为空"
        return None, None
    if field.source == CatalogFieldSource.MERCHANT_EXTRA:
        return raw, None
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
                return raw, f"值不在平台枚举范围：{invalid}"
        joined = "、".join(unique)
        return joined, _validate_scalar(joined, field)
    else:
        rule = ColumnRule(
            name=field.field_id,
            title=field.title,
            required=field.required,
            type=field.field_type,
            enum_values=field.enum_values,
            validation=field.validation,
            format=field.number_format,
        )
    parsed: ParsedValue = parse_excel_value(raw, rule, workbook.excel_epoch)
    if not parsed.valid:
        return raw, parsed.error or "字段值无效"
    validation_error = _validate_scalar(parsed.normalized, field)
    return parsed.normalized, validation_error


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
    return ReviewItem(
        review_type="category",
        key=f"row:{resolution.excel_row}",
        excel_row=resolution.excel_row,
        message=(
            f"类目 {resolution.raw_category!r} 需要人工确认"
            if resolution.status == "manual_review"
            else "未提供可解析的平台类目"
        ),
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
