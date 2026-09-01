from __future__ import annotations

import hashlib

from ..models import ColumnRule, normalize_header
from .models import (
    CatalogFieldDefinition,
    CatalogFieldSource,
    DynamicSchemaPlan,
    FieldMapping,
    PlannedField,
    ReviewItem,
)


_SOURCE_ORDER = {
    CatalogFieldSource.FIXED: 0,
    CatalogFieldSource.PLATFORM_ATTRIBUTE: 1,
    CatalogFieldSource.PLATFORM_SPECIFICATION: 2,
    CatalogFieldSource.MERCHANT_EXTRA: 3,
}


def _fixed_field(column: ColumnRule, display_order: int) -> CatalogFieldDefinition:
    return CatalogFieldDefinition(
        field_id=column.name,
        title=column.title,
        aliases=column.aliases,
        source=CatalogFieldSource.FIXED,
        field_type=column.type,
        required=column.required,
        multiple=column.type.value == "set",
        display_order=display_order,
        enum_values=column.enum_values,
        number_format=column.format,
        validation=column.validation,
    )


def _merchant_field(raw_header: str, physical_column: int, used_ids: set[str]) -> CatalogFieldDefinition:
    digest = hashlib.sha256(normalize_header(raw_header).casefold().encode("utf-8")).hexdigest()[:12]
    candidate = f"merchant_extra.{digest}"
    suffix = 2
    while candidate in used_ids:
        candidate = f"merchant_extra.{digest}.{suffix}"
        suffix += 1
    used_ids.add(candidate)
    return CatalogFieldDefinition(
        field_id=candidate,
        title=normalize_header(raw_header) or f"未命名字段{physical_column}",
        source=CatalogFieldSource.MERCHANT_EXTRA,
        display_order=physical_column,
    )


def build_dynamic_schema(
    *,
    category_id: str,
    category_name: str,
    fixed_columns: list[ColumnRule],
    platform_fields: list[CatalogFieldDefinition],
    headers: list[str],
    mappings: list[FieldMapping],
) -> DynamicSchemaPlan:
    """Build fixed -> attributes -> specifications -> merchant-extra output order."""
    if len(headers) != len(mappings):
        raise ValueError("headers and mappings must describe the same physical columns")
    if any(field.category_id != category_id for field in platform_fields):
        raise ValueError("platform_fields contains a field from a different category")
    if any(field.source not in {CatalogFieldSource.PLATFORM_ATTRIBUTE, CatalogFieldSource.PLATFORM_SPECIFICATION} for field in platform_fields):
        raise ValueError("platform_fields may contain only platform attributes and specifications")

    fixed = [_fixed_field(column, index) for index, column in enumerate(fixed_columns)]
    combined = [*fixed, *platform_fields]
    ids = [field.field_id for field in combined]
    if len(ids) != len(set(ids)):
        raise ValueError("fixed and platform fields contain duplicate field_id values")
    used_ids = set(ids)
    by_id = {field.field_id: field for field in combined}
    accepted_by_id = {
        mapping.field_id: mapping
        for mapping in mappings
        if mapping.status == "accepted" and mapping.field_id is not None
    }
    review_items: list[ReviewItem] = []
    for mapping in mappings:
        if mapping.status == "manual_review":
            review_items.append(ReviewItem(
                review_type="duplicate_header" if mapping.match_type == "duplicate" else "field_mapping",
                key=f"column:{mapping.physical_column}",
                physical_column=mapping.physical_column,
                raw_header=mapping.raw_header,
                category_id=category_id,
                message=(
                    f"字段 {mapping.raw_header!r} 重复映射到 {mapping.field_id!r}"
                    if mapping.match_type == "duplicate"
                    else f"字段 {mapping.raw_header!r} 需要人工确认映射"
                ),
                candidates=mapping.candidates,
            ))

    ordered_core = sorted(combined, key=lambda field: (_SOURCE_ORDER[field.source], field.display_order, field.field_id))
    planned: list[PlannedField] = []
    for field in ordered_core:
        mapping = accepted_by_id.get(field.field_id)
        planned.append(PlannedField(
            ordinal=len(planned) + 1,
            field=field,
            present=mapping is not None,
            source_header=mapping.raw_header if mapping else None,
            source_column=mapping.physical_column if mapping else None,
        ))

    for mapping in mappings:
        if mapping.status == "accepted":
            continue
        merchant = _merchant_field(mapping.raw_header, mapping.physical_column, used_ids)
        planned.append(PlannedField(
            ordinal=len(planned) + 1,
            field=merchant,
            present=True,
            source_header=mapping.raw_header,
            source_column=mapping.physical_column,
        ))

    safe_name = category_name.strip()[:31]
    if not safe_name or any(character in safe_name for character in "[]:*?/\\"):
        safe_name = f"类目-{category_id}"[:31]
    return DynamicSchemaPlan(
        category_id=category_id,
        worksheet_name=safe_name,
        fields=planned,
        mappings=mappings,
        review_items=review_items,
    )
