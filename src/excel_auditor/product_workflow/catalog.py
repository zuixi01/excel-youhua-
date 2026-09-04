from __future__ import annotations

from collections.abc import Sequence
from decimal import Decimal, InvalidOperation
from typing import Any, Protocol
from urllib.parse import quote

from ..ids import new_ulid
from ..models import (
    FieldType,
    ProductCatalogEndpoint,
    ProductCategoryRecordMapping,
    ProductFieldRecordMapping,
    ProductWorkflowConfig,
    StandardSourceConfig,
    ValidationConfig,
)
from ..source_paths import validate_managed_http_path
from ..standard_sources import ManagedHttpSource
from .models import (
    CatalogFieldDefinition,
    CatalogFieldSource,
    CatalogSchemaSnapshot,
    CategoryCatalogSnapshot,
    CategoryDefinition,
)


class CatalogAdapter(Protocol):
    def list_categories(self) -> list[CategoryDefinition]: ...

    def fetch_schema(self, category_id: str) -> CatalogSchemaSnapshot: ...


class InMemoryCatalogAdapter:
    """Deterministic adapter for development, Golden tests, and offline demos."""

    def __init__(
        self,
        categories: list[CategoryDefinition],
        fields_by_category: dict[str, list[CatalogFieldDefinition]],
        *,
        connection_id: str = "in-memory",
    ) -> None:
        self._categories = list(categories)
        self._fields = {key: list(value) for key, value in fields_by_category.items()}
        self.connection_id = connection_id

    def list_categories(self) -> list[CategoryDefinition]:
        return list(self._categories)

    def fetch_schema(self, category_id: str) -> CatalogSchemaSnapshot:
        if category_id not in self._fields:
            raise KeyError(f"platform category schema not found: {category_id}")
        return CatalogSchemaSnapshot.create(
            snapshot_id=new_ulid("catalog_"),
            connection_id=self.connection_id,
            category_id=category_id,
            fields=list(self._fields[category_id]),
            source_metadata={"adapter": "in_memory"},
        )


class FrozenCatalogAdapter:
    """Replay one immutable category and field snapshot without calling the platform again."""

    def __init__(
        self,
        category_snapshot: CategoryCatalogSnapshot,
        schema_snapshots: list[CatalogSchemaSnapshot],
    ) -> None:
        self.category_snapshot = category_snapshot
        self.connection_id = category_snapshot.connection_id
        self._schemas = {snapshot.category_id: snapshot for snapshot in schema_snapshots}
        if len(self._schemas) != len(schema_snapshots):
            raise ValueError("frozen catalog contains duplicate category schema snapshots")
        if any(snapshot.connection_id != self.connection_id for snapshot in schema_snapshots):
            raise ValueError("frozen catalog snapshots must come from the same connection")

    def list_categories(self) -> list[CategoryDefinition]:
        return list(self.category_snapshot.categories)

    def fetch_schema(self, category_id: str) -> CatalogSchemaSnapshot:
        try:
            return self._schemas[category_id]
        except KeyError as exc:
            raise KeyError(f"frozen platform category schema not found: {category_id}") from exc


class ManagedHttpCatalogAdapter:
    """Load the platform catalog through the existing SSRF-safe managed connection layer."""

    def __init__(self, source: ManagedHttpSource, config: ProductWorkflowConfig) -> None:
        self.source = source
        self.config = config
        self.connection_id = config.catalog_connection_id

    def list_categories(self) -> list[CategoryDefinition]:
        source = StandardSourceConfig(
            type="managed_http",
            connection_id=self.config.catalog_connection_id,
            method=self.config.category.category_list_method,
            path=self.config.category.category_list_path,
            data_json_path=self.config.category.category_list_json_path,
            static_parameters=self.config.category.category_list_static_parameters,
            pagination=self.config.category.category_list_pagination,
        )
        records, _metadata = self.source.fetch_with_metadata(source)
        return [
            _parse_category(record, self.config.category.record_mapping)
            for record in _materialize_records(records, "category list")
        ]

    def fetch_schema(self, category_id: str) -> CatalogSchemaSnapshot:
        attributes, attribute_metadata = self._fetch_fields(
            category_id,
            self.config.category.attributes,
            CatalogFieldSource.PLATFORM_ATTRIBUTE,
        )
        specifications, specification_metadata = self._fetch_fields(
            category_id,
            self.config.category.specifications,
            CatalogFieldSource.PLATFORM_SPECIFICATION,
        )
        return CatalogSchemaSnapshot.create(
            snapshot_id=new_ulid("catalog_"),
            connection_id=self.config.catalog_connection_id,
            category_id=category_id,
            fields=[*attributes, *specifications],
            source_metadata={
                "attributes": attribute_metadata,
                "specifications": specification_metadata,
            },
        )

    def _fetch_fields(
        self,
        category_id: str,
        endpoint: ProductCatalogEndpoint,
        source_type: CatalogFieldSource,
    ) -> tuple[list[CatalogFieldDefinition], dict[str, Any]]:
        encoded_id = quote(category_id, safe="")
        path = endpoint.path_template.replace("{category_id}", encoded_id)
        validate_managed_http_path(path, field_name="resolved catalog endpoint")
        parameters = dict(endpoint.static_parameters)
        if endpoint.category_parameter:
            parameters[endpoint.category_parameter] = category_id
        source = StandardSourceConfig(
            type="managed_http",
            connection_id=self.config.catalog_connection_id,
            method=endpoint.method,
            path=path,
            data_json_path=endpoint.data_json_path,
            static_parameters=parameters,
            pagination=endpoint.pagination,
        )
        records, metadata = self.source.fetch_with_metadata(source)
        parsed = [
            _parse_field(
                record,
                category_id=category_id,
                source_type=source_type,
                mapping=endpoint.record_mapping,
            )
            for record in _materialize_records(records, source_type.value)
        ]
        return parsed, metadata


def _materialize_records(
    records: Sequence[dict[str, Any]] | dict[str, list[dict[str, Any]]],
    label: str,
) -> list[dict[str, Any]]:
    if isinstance(records, dict):
        raise ValueError(f"CATALOG_SOURCE_FAILED: {label} must be an array of objects")
    try:
        return list(records)
    finally:
        close = getattr(records, "close", None)
        if close is not None:
            close()


def _required_text(record: dict[str, Any], key: str, label: str) -> str:
    value = record.get(key)
    if isinstance(value, bool) or value is None or not str(value).strip():
        raise ValueError(f"CATALOG_SOURCE_FAILED: {label} requires non-blank {key}")
    return str(value).strip()


def _string_list(value: Any, label: str) -> list[str]:
    if value is None:
        return []
    if not isinstance(value, list) or any(not isinstance(item, str) or not item.strip() for item in value):
        raise ValueError(f"CATALOG_SOURCE_FAILED: {label} must be an array of non-blank strings")
    return value


def _strict_bool(value: Any, *, default: bool, label: str) -> bool:
    if value is None:
        return default
    if not isinstance(value, bool):
        raise ValueError(f"CATALOG_SOURCE_FAILED: {label} must be a boolean")
    return value


def _optional_decimal(value: Any, *, label: str) -> Decimal | None:
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, (int, float, str, Decimal)):
        raise ValueError(f"CATALOG_SOURCE_FAILED: {label} must be a finite decimal")
    if isinstance(value, str) and not value.strip():
        raise ValueError(f"CATALOG_SOURCE_FAILED: {label} must be a finite decimal")
    try:
        parsed = Decimal(str(value).strip())
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(f"CATALOG_SOURCE_FAILED: {label} must be a finite decimal") from exc
    if not parsed.is_finite():
        raise ValueError(f"CATALOG_SOURCE_FAILED: {label} must be a finite decimal")
    return parsed


def _optional_non_negative_int(value: Any, *, label: str) -> int | None:
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ValueError(f"CATALOG_SOURCE_FAILED: {label} must be a non-negative integer")
    return value


def _optional_string(value: Any, *, label: str) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"CATALOG_SOURCE_FAILED: {label} must be a non-blank string")
    return value


def _parse_category(record: dict[str, Any], mapping: ProductCategoryRecordMapping) -> CategoryDefinition:
    path = _string_list(_optional_value(record, mapping.path_key), "category.path")
    aliases = _string_list(_optional_value(record, mapping.aliases_key), "category.aliases")
    parent = _optional_value(record, mapping.parent_id_key)
    return CategoryDefinition(
        category_id=_required_text(record, mapping.id_key, "category"),
        name=_required_text(record, mapping.name_key, "category"),
        parent_id=str(parent).strip() if parent is not None else None,
        path=path,
        aliases=aliases,
        active=_strict_bool(
            _optional_value(record, mapping.active_key),
            default=True,
            label="category.active",
        ),
        raw=record,
    )


def _parse_field(
    record: dict[str, Any],
    *,
    category_id: str,
    source_type: CatalogFieldSource,
    mapping: ProductFieldRecordMapping,
) -> CatalogFieldDefinition:
    raw_type = _optional_value(record, mapping.type_key) or "string"
    if not isinstance(raw_type, str) or not raw_type.strip():
        raise ValueError("CATALOG_SOURCE_FAILED: field.type must be a non-blank string")
    normalized_type = raw_type.strip().casefold()
    type_aliases = {
        alias.strip().casefold(): target
        for alias, target in mapping.type_value_aliases.items()
    }
    try:
        field_type = type_aliases.get(normalized_type) or FieldType(normalized_type)
    except ValueError as exc:
        raise ValueError(f"CATALOG_SOURCE_FAILED: unsupported field type {raw_type!r}") from exc
    raw_order = _optional_value(record, mapping.display_order_key)
    raw_order = 0 if raw_order is None else raw_order
    if isinstance(raw_order, bool) or not isinstance(raw_order, int) or raw_order < 0:
        raise ValueError("CATALOG_SOURCE_FAILED: field.display_order must be a non-negative integer")
    attribute_id = _required_text(record, mapping.id_key, "field")
    validation = ValidationConfig(
        nullable=_strict_bool(
            _optional_value(record, mapping.nullable_key),
            default=True,
            label="field.nullable",
        ),
        unique=_strict_bool(
            _optional_value(record, mapping.unique_key),
            default=False,
            label="field.unique",
        ),
        min=_optional_decimal(_optional_value(record, mapping.min_key), label="field.min"),
        max=_optional_decimal(_optional_value(record, mapping.max_key), label="field.max"),
        min_length=_optional_non_negative_int(
            _optional_value(record, mapping.min_length_key),
            label="field.min_length",
        ),
        max_length=_optional_non_negative_int(
            _optional_value(record, mapping.max_length_key),
            label="field.max_length",
        ),
        regex=_optional_string(_optional_value(record, mapping.regex_key), label="field.regex"),
    )
    return CatalogFieldDefinition(
        field_id=f"platform.{source_type.value}.{attribute_id}",
        attribute_id=attribute_id,
        category_id=category_id,
        title=_required_text(record, mapping.title_key, "field"),
        source=source_type,
        aliases=_string_list(_optional_value(record, mapping.aliases_key), "field.aliases"),
        field_type=field_type,
        required=_strict_bool(_optional_value(record, mapping.required_key), default=False, label="field.required"),
        multiple=_strict_bool(_optional_value(record, mapping.multiple_key), default=False, label="field.multiple"),
        display_order=raw_order,
        enum_values=_string_list(_optional_value(record, mapping.enum_values_key), "field.enum_values"),
        number_format=_optional_string(
            _optional_value(record, mapping.number_format_key),
            label="field.number_format",
        ),
        timezone=_optional_string(
            _optional_value(record, mapping.timezone_key),
            label="field.timezone",
        ),
        validation=validation,
        raw=record,
    )


def _optional_value(record: dict[str, Any], key: str | None) -> Any:
    return record.get(key) if key is not None else None
