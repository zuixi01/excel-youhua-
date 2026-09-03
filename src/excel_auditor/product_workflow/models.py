from __future__ import annotations

import hashlib
import json
from datetime import UTC, datetime
from enum import Enum
from collections.abc import Sequence
from typing import Any, Literal

from pydantic import Field, SkipValidation, field_serializer, field_validator, model_validator

from ..models import FieldType, StrictModel, ValidationConfig, normalize_header
from ..spill import SpillableSequence


class CatalogFieldSource(str, Enum):
    FIXED = "fixed"
    PLATFORM_ATTRIBUTE = "platform_attribute"
    PLATFORM_SPECIFICATION = "platform_specification"
    MERCHANT_EXTRA = "merchant_extra"


class CategoryDefinition(StrictModel):
    category_id: str = Field(min_length=1, max_length=256)
    name: str = Field(min_length=1, max_length=512)
    parent_id: str | None = Field(default=None, max_length=256)
    path: list[str] = Field(default_factory=list, max_length=32)
    aliases: list[str] = Field(default_factory=list, max_length=128)
    active: bool = True
    raw: dict[str, Any] = Field(default_factory=dict)

    @field_validator("category_id", "name")
    @classmethod
    def non_blank(cls, value: str) -> str:
        if not normalize_header(value):
            raise ValueError("category identifier and name must not be blank")
        return value

    @field_validator("aliases")
    @classmethod
    def unique_aliases(cls, values: list[str]) -> list[str]:
        normalized = [normalize_header(value).casefold() for value in values]
        if any(not value for value in normalized) or len(normalized) != len(set(normalized)):
            raise ValueError("category aliases must be non-blank and unique after normalization")
        return values


class CatalogFieldDefinition(StrictModel):
    field_id: str = Field(min_length=1, max_length=256)
    title: str = Field(min_length=1, max_length=512)
    source: CatalogFieldSource
    aliases: list[str] = Field(default_factory=list, max_length=128)
    field_type: FieldType = FieldType.STRING
    required: bool = False
    multiple: bool = False
    display_order: int = Field(default=0, ge=0)
    category_id: str | None = Field(default=None, max_length=256)
    attribute_id: str | None = Field(default=None, max_length=256)
    enum_values: list[str] = Field(default_factory=list, max_length=10_000)
    number_format: str | None = Field(default=None, max_length=255)
    timezone: str | None = Field(default=None, max_length=255)
    validation: ValidationConfig = Field(default_factory=ValidationConfig)
    raw: dict[str, Any] = Field(default_factory=dict)

    @field_validator("field_id", "title")
    @classmethod
    def non_blank(cls, value: str) -> str:
        if not normalize_header(value) or any(ord(character) < 32 for character in value):
            raise ValueError("catalog field identifiers and titles must be printable and non-blank")
        return value

    @field_validator("aliases")
    @classmethod
    def unique_aliases(cls, values: list[str]) -> list[str]:
        normalized = [normalize_header(value).casefold() for value in values]
        if any(not value for value in normalized) or len(normalized) != len(set(normalized)):
            raise ValueError("field aliases must be non-blank and unique after normalization")
        return values

    @field_validator("timezone")
    @classmethod
    def valid_timezone(cls, value: str | None) -> str | None:
        if value is not None:
            from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

            try:
                ZoneInfo(value)
            except ZoneInfoNotFoundError as exc:
                raise ValueError(f"unknown IANA timezone: {value}") from exc
        return value

    @model_validator(mode="after")
    def source_contract(self) -> "CatalogFieldDefinition":
        platform_sources = {
            CatalogFieldSource.PLATFORM_ATTRIBUTE,
            CatalogFieldSource.PLATFORM_SPECIFICATION,
        }
        if self.source in platform_sources and (not self.category_id or not self.attribute_id):
            raise ValueError("platform catalog fields require category_id and attribute_id")
        if self.source not in platform_sources and (self.category_id is not None or self.attribute_id is not None):
            raise ValueError("fixed and merchant-extra fields cannot carry platform identifiers")
        if self.field_type == FieldType.ENUM and not self.enum_values:
            raise ValueError("enum catalog fields require enum_values")
        if self.field_type != FieldType.ENUM and self.enum_values:
            raise ValueError("enum_values are only valid for enum catalog fields")
        if len(self.enum_values) != len(set(self.enum_values)):
            raise ValueError("catalog field enum_values must be unique")
        if self.multiple and self.field_type not in {FieldType.STRING, FieldType.ENUM, FieldType.SET}:
            raise ValueError("multiple catalog fields must use string, enum, or set type")
        if (self.validation.min is not None or self.validation.max is not None) and self.field_type not in {
            FieldType.INTEGER,
            FieldType.DECIMAL,
        }:
            raise ValueError("catalog numeric validation bounds require integer or decimal type")
        if self.number_format is not None and (
            not self.number_format.strip()
            or any(ord(character) < 32 or 127 <= ord(character) <= 159 for character in self.number_format)
        ):
            raise ValueError("catalog number_format must be printable and non-blank")
        if self.timezone is not None and self.field_type != FieldType.DATETIME:
            raise ValueError("catalog timezone is only valid for datetime fields")
        return self

class CatalogSchemaSnapshot(StrictModel):
    snapshot_id: str = Field(min_length=1, max_length=128)
    connection_id: str = Field(min_length=1, max_length=128)
    category_id: str = Field(min_length=1, max_length=256)
    captured_at: datetime
    fields: list[CatalogFieldDefinition]
    content_sha256: str = Field(pattern=r"^[0-9a-f]{64}$")
    source_metadata: dict[str, Any] = Field(default_factory=dict)

    @classmethod
    def create(
        cls,
        *,
        snapshot_id: str,
        connection_id: str,
        category_id: str,
        fields: list[CatalogFieldDefinition],
        captured_at: datetime | None = None,
        source_metadata: dict[str, Any] | None = None,
    ) -> "CatalogSchemaSnapshot":
        ordered = sorted(fields, key=lambda field: (field.source.value, field.display_order, field.field_id))
        payload = [field.model_dump(mode="json") for field in ordered]
        content = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        return cls(
            snapshot_id=snapshot_id,
            connection_id=connection_id,
            category_id=category_id,
            captured_at=captured_at or datetime.now(UTC),
            fields=ordered,
            content_sha256=hashlib.sha256(content.encode("utf-8")).hexdigest(),
            source_metadata=source_metadata or {},
        )

    @model_validator(mode="after")
    def verify_snapshot(self) -> "CatalogSchemaSnapshot":
        if any(field.category_id != self.category_id for field in self.fields):
            raise ValueError("catalog snapshot contains a field from a different category")
        ids = [field.field_id for field in self.fields]
        if len(ids) != len(set(ids)):
            raise ValueError("catalog snapshot field_id values must be unique")
        payload = [field.model_dump(mode="json") for field in self.fields]
        content = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        expected = hashlib.sha256(content.encode("utf-8")).hexdigest()
        if self.content_sha256 != expected:
            raise ValueError("catalog snapshot content hash does not match its fields")
        return self


class CategoryCatalogSnapshot(StrictModel):
    snapshot_id: str = Field(min_length=1, max_length=128)
    connection_id: str = Field(min_length=1, max_length=128)
    captured_at: datetime
    categories: list[CategoryDefinition]
    content_sha256: str = Field(pattern=r"^[0-9a-f]{64}$")
    source_metadata: dict[str, Any] = Field(default_factory=dict)

    @classmethod
    def create(
        cls,
        *,
        snapshot_id: str,
        connection_id: str,
        categories: list[CategoryDefinition],
        captured_at: datetime | None = None,
        source_metadata: dict[str, Any] | None = None,
    ) -> "CategoryCatalogSnapshot":
        ordered = sorted(categories, key=lambda category: category.category_id)
        payload = [category.model_dump(mode="json") for category in ordered]
        content = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        return cls(
            snapshot_id=snapshot_id,
            connection_id=connection_id,
            captured_at=captured_at or datetime.now(UTC),
            categories=ordered,
            content_sha256=hashlib.sha256(content.encode("utf-8")).hexdigest(),
            source_metadata=source_metadata or {},
        )

    @model_validator(mode="after")
    def verify_snapshot(self) -> "CategoryCatalogSnapshot":
        identifiers = [category.category_id for category in self.categories]
        if len(identifiers) != len(set(identifiers)):
            raise ValueError("category catalog snapshot category_id values must be unique")
        payload = [category.model_dump(mode="json") for category in self.categories]
        content = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        expected = hashlib.sha256(content.encode("utf-8")).hexdigest()
        if self.content_sha256 != expected:
            raise ValueError("category catalog snapshot content hash does not match its categories")
        return self


class MappingCandidate(StrictModel):
    field_id: str
    title: str
    confidence: float = Field(ge=0, le=100)
    match_value: str


class FieldMapping(StrictModel):
    physical_column: int = Field(ge=1)
    raw_header: str
    normalized_header: str
    field_id: str | None = None
    match_type: Literal["exact", "confirmed_alias", "fuzzy_suggestion", "unmapped", "duplicate"]
    status: Literal["accepted", "manual_review", "unmapped"]
    confidence: float | None = Field(default=None, ge=0, le=100)
    candidates: list[MappingCandidate] = Field(default_factory=list)


class CategoryResolution(StrictModel):
    excel_row: int = Field(ge=1)
    raw_category_id: str | None = None
    raw_category: str | None = None
    category_id: str | None = None
    category_name: str | None = None
    status: Literal["resolved", "manual_review", "unresolved"]
    match_type: Literal["id", "exact", "confirmed", "fuzzy_suggestion", "missing", "ambiguous", "id_name_conflict", "invalid_id"]
    confidence: float | None = Field(default=None, ge=0, le=100)
    candidates: list[MappingCandidate] = Field(default_factory=list)


class ReviewItem(StrictModel):
    review_type: Literal["category", "field_mapping", "duplicate_header", "schema_conflict"]
    key: str
    message: str
    raw_header: str | None = Field(default=None, max_length=512)
    category_id: str | None = Field(default=None, max_length=256)
    physical_column: int | None = Field(default=None, ge=1)
    excel_row: int | None = Field(default=None, ge=1)
    candidates: list[MappingCandidate] = Field(default_factory=list)


class PlannedField(StrictModel):
    ordinal: int = Field(ge=1)
    field: CatalogFieldDefinition
    present: bool
    source_header: str | None = None
    source_column: int | None = Field(default=None, ge=1)


class DynamicSchemaPlan(StrictModel):
    category_id: str
    worksheet_name: str
    fields: list[PlannedField]
    mappings: list[FieldMapping]
    review_items: list[ReviewItem] = Field(default_factory=list)

    @model_validator(mode="after")
    def deterministic_contract(self) -> "DynamicSchemaPlan":
        if [field.ordinal for field in self.fields] != list(range(1, len(self.fields) + 1)):
            raise ValueError("planned field ordinals must be contiguous and one-based")
        ids = [field.field.field_id for field in self.fields]
        if len(ids) != len(set(ids)):
            raise ValueError("dynamic schema plan field_id values must be unique")
        return self


class ProductCellIssue(StrictModel):
    issue_type: Literal[
        "required_missing",
        "invalid_value",
        "excel_write_unsafe",
        "ambiguous_mapping",
        "duplicate_primary_key",
        "unique_value",
        "cross_field_error",
        "sku_combination_limit",
    ]
    excel_row: int = Field(ge=1)
    category_id: str | None = None
    field_id: str | None = None
    physical_column: int | None = Field(default=None, ge=1)
    raw_value: Any = None
    message: str
    color: str


class UnresolvedProductRow(StrictModel):
    excel_row: int = Field(ge=1)
    values: list[Any]
    category_resolution: CategoryResolution


class NormalizedCategorySheet(StrictModel):
    category_id: str
    category_name: str
    worksheet_name: str
    plan: DynamicSchemaPlan
    source_excel_rows: list[int]
    rows: SkipValidation[Sequence[dict[str, Any]]]
    sku_rows: SkipValidation[Sequence[dict[str, Any]]] = Field(default_factory=list)
    sku_source_excel_rows: list[int] = Field(default_factory=list)

    @field_validator("rows", "sku_rows", mode="before")
    @classmethod
    def validate_row_sequence(cls, value: Any) -> Sequence[dict[str, Any]]:
        if isinstance(value, SpillableSequence):
            return value
        if isinstance(value, (str, bytes, bytearray)) or not isinstance(value, Sequence):
            raise ValueError("product rows must be a sequence")
        return [dict(item) for item in value]

    @model_validator(mode="after")
    def aligned_rows(self) -> "NormalizedCategorySheet":
        if len(self.rows) != len(self.source_excel_rows):
            raise ValueError("product rows must align with source_excel_rows")
        if len(self.sku_rows) != len(self.sku_source_excel_rows):
            raise ValueError("SKU rows must align with sku_source_excel_rows")
        return self

    @field_serializer("rows", "sku_rows")
    def serialize_rows(self, value: Sequence[dict[str, Any]]) -> list[dict[str, Any]]:
        return list(value)

    def close(self) -> None:
        for sequence in (self.rows, self.sku_rows):
            close = getattr(sequence, "close", None)
            if close is not None:
                close()


class ProductNormalizationResult(StrictModel):
    category_catalog_snapshot: CategoryCatalogSnapshot
    catalog_snapshots: list[CatalogSchemaSnapshot]
    category_sheets: list[NormalizedCategorySheet]
    source_headers: list[str] = Field(default_factory=list)
    unresolved_rows: list[UnresolvedProductRow] = Field(default_factory=list)
    review_items: list[ReviewItem] = Field(default_factory=list)
    issues: list[ProductCellIssue] = Field(default_factory=list)
    requires_manual_review: bool = False
    merchant_extra_header_color: str = Field(default="D9D9D9", pattern=r"^[0-9A-Fa-f]{6}$")

    def close(self) -> None:
        for sheet in self.category_sheets:
            sheet.close()


class ProductReviewDecision(StrictModel):
    action: Literal["confirm_category", "confirm_mapping", "keep_extra", "reject"]
    category_id: str | None = Field(default=None, min_length=1, max_length=256)
    field_id: str | None = Field(default=None, min_length=1, max_length=256)
    raw_header: str | None = Field(default=None, min_length=1, max_length=512)
    note: str | None = Field(default=None, max_length=2000)

    @model_validator(mode="after")
    def decision_contract(self) -> "ProductReviewDecision":
        if self.action == "confirm_category" and self.category_id is None:
            raise ValueError("confirm_category requires category_id")
        if self.action == "confirm_mapping" and (self.field_id is None or self.raw_header is None):
            raise ValueError("confirm_mapping requires field_id and raw_header")
        if self.action in {"keep_extra", "reject"} and any(
            value is not None for value in (self.category_id, self.field_id, self.raw_header)
        ):
            raise ValueError(f"{self.action} cannot include a category or field target")
        return self
