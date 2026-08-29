from __future__ import annotations

import hashlib
import json
import re
from datetime import datetime
from decimal import Decimal
from enum import Enum
from typing import Any, Literal, Sequence

from pydantic import BaseModel, ConfigDict, Field, SkipValidation, field_validator, model_validator

from .spill import SpillableSequence
from .source_paths import validate_managed_http_path, validate_parameter_name, validate_simple_json_path


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid")


class FieldType(str, Enum):
    STRING = "string"
    INTEGER = "integer"
    DECIMAL = "decimal"
    DATE = "date"
    DATETIME = "datetime"
    BOOLEAN = "boolean"
    ENUM = "enum"
    PHONE = "phone"
    ID_CODE = "id_code"
    POSTAL_CODE = "postal_code"
    SET = "set"
    JSON = "json"
    FUZZY_STRING = "fuzzy_string"


class DifferenceType(str, Enum):
    MISSING_SHEET = "MISSING_SHEET"
    EXTRA_SHEET = "EXTRA_SHEET"
    AMBIGUOUS_SHEET = "AMBIGUOUS_SHEET"
    MISSING_HEADER = "MISSING_HEADER"
    EXTRA_HEADER = "EXTRA_HEADER"
    DUPLICATE_HEADER = "DUPLICATE_HEADER"
    AMBIGUOUS_HEADER = "AMBIGUOUS_HEADER"
    HEADER_NOT_FOUND = "HEADER_NOT_FOUND"
    HEADER_ORDER_MISMATCH = "HEADER_ORDER_MISMATCH"
    EMPTY_PRIMARY_KEY = "EMPTY_PRIMARY_KEY"
    DUPLICATE_PRIMARY_KEY = "DUPLICATE_PRIMARY_KEY"
    EXTRA_RECORD = "EXTRA_RECORD"
    MISSING_RECORD = "MISSING_RECORD"
    VALUE_MISMATCH = "VALUE_MISMATCH"
    INVALID_VALUE = "INVALID_VALUE"
    VALIDATION_ERROR = "VALIDATION_ERROR"
    NORMALIZED_MATCH = "NORMALIZED_MATCH"
    UNSUPPORTED_FEATURE = "UNSUPPORTED_FEATURE"
    RENDER_WARNING = "RENDER_WARNING"


class CompareConfig(StrictModel):
    mode: Literal["exact", "numeric", "date", "datetime", "ignore_case", "set", "json"] = "exact"
    absolute_tolerance: Decimal = Decimal("0")
    relative_tolerance: Decimal = Decimal("0")
    decimal_places: int | None = Field(default=None, ge=0, le=28)
    timezone: str | None = None
    allow_naive_datetime: bool = False
    precision: Literal["second", "minute", "hour", "day"] = "second"
    formula_mode: Literal["formula", "cached_value", "reject"] = "reject"

    @field_validator("absolute_tolerance", "relative_tolerance", mode="before")
    @classmethod
    def decimal_configuration_must_be_string(cls, value: Any) -> Any:
        if not isinstance(value, (str, Decimal)):
            raise ValueError("decimal tolerance must be configured as a string")
        return value

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
    def non_negative_tolerances(self) -> "CompareConfig":
        if not self.absolute_tolerance.is_finite() or not self.relative_tolerance.is_finite():
            raise ValueError("numeric tolerances must be finite")
        if self.absolute_tolerance < 0 or self.relative_tolerance < 0:
            raise ValueError("numeric tolerances must be non-negative")
        return self


class ValidationConfig(StrictModel):
    nullable: bool = True
    unique: bool = False
    min: Decimal | None = None
    max: Decimal | None = None
    min_length: int | None = Field(default=None, ge=0)
    max_length: int | None = Field(default=None, ge=0)
    regex: str | None = None

    @field_validator("regex")
    @classmethod
    def safe_regex(cls, value: str | None) -> str | None:
        if value is not None:
            if len(value) > 256:
                raise ValueError("regex must be no longer than 256 characters")
            if re.search(r"\\[1-9]|\(\?<[=!]|\([^)]*[+*][^)]*\)\s*(?:[+*]|\{)", value):
                raise ValueError("regex contains a forbidden high-complexity construct")
            re.compile(value)
        return value

    @model_validator(mode="after")
    def ordered_bounds(self) -> "ValidationConfig":
        if any(bound is not None and not bound.is_finite() for bound in (self.min, self.max)):
            raise ValueError("validation numeric bounds must be finite")
        if self.min is not None and self.max is not None and self.min > self.max:
            raise ValueError("validation min cannot exceed max")
        if self.min_length is not None and self.max_length is not None and self.min_length > self.max_length:
            raise ValueError("validation min_length cannot exceed max_length")
        return self


class RegexReplacement(StrictModel):
    pattern: str
    replacement: str = ""
    ignore_case: bool = False

    @field_validator("pattern")
    @classmethod
    def safe_pattern(cls, value: str) -> str:
        return ValidationConfig.safe_regex(value) or ""

    @field_validator("replacement")
    @classmethod
    def bounded_replacement(cls, value: str) -> str:
        if len(value) > 256:
            raise ValueError("regex replacement must be no longer than 256 characters")
        return value


class ColumnRule(StrictModel):
    name: str
    title: str
    aliases: list[str] = Field(default_factory=list)
    required: bool = False
    type: FieldType = FieldType.STRING
    normalize: list[str] = Field(default_factory=list)
    compare: CompareConfig = Field(default_factory=CompareConfig)
    validation: ValidationConfig = Field(default_factory=ValidationConfig)
    enum_values: list[str] = Field(default_factory=list)
    enum_aliases: dict[str, str] = Field(default_factory=dict)
    value_aliases: dict[str, str] = Field(default_factory=dict)
    regex_replacements: list[RegexReplacement] = Field(default_factory=list, max_length=8)
    boolean_true_values: list[str] = Field(default_factory=lambda: ["true", "1", "yes", "y", "是", "真"])
    boolean_false_values: list[str] = Field(default_factory=lambda: ["false", "0", "no", "n", "否", "假"])
    parse_formats: list[str] = Field(default_factory=list)
    format: str | None = None
    formula_template: str | None = None
    separator: str = ","
    missing_column_action: Literal["insert", "report_only"] | None = None
    fill_static_default: bool = False
    static_default: Any = None
    sensitive: bool = False

    @model_validator(mode="after")
    def validate_field(self) -> "ColumnRule":
        allowed_normalizers = {"trim", "unicode_nfkc", "collapse_spaces", "uppercase", "lowercase", "casefold", "remove_group_separator", "remove_currency_symbol", "percent_to_decimal"}
        unknown = set(self.normalize) - allowed_normalizers
        if unknown:
            raise ValueError(f"unknown normalizers for {self.name!r}: {sorted(unknown)}")
        if self.type == FieldType.ENUM and not self.enum_values:
            raise ValueError(f"enum field {self.name!r} requires enum_values")
        if self.type != FieldType.ENUM and (self.enum_values or self.enum_aliases):
            raise ValueError(f"enum configuration is incompatible with {self.type.value}")
        if len(self.enum_values) != len(set(self.enum_values)):
            raise ValueError(f"enum field {self.name!r} has duplicate enum_values")
        invalid_enum_targets = set(self.enum_aliases.values()) - set(self.enum_values)
        if invalid_enum_targets:
            raise ValueError(f"enum aliases for {self.name!r} target unknown values: {sorted(invalid_enum_targets)}")
        remapped_canonical_values = {
            alias: target
            for alias, target in self.enum_aliases.items()
            if alias in self.enum_values and alias != target
        }
        if remapped_canonical_values:
            raise ValueError(f"enum aliases for {self.name!r} remap canonical values")
        if (self.value_aliases or self.regex_replacements) and self.type not in {
            FieldType.STRING, FieldType.PHONE, FieldType.ID_CODE, FieldType.POSTAL_CODE, FieldType.FUZZY_STRING
        }:
            raise ValueError(f"string aliases and regex replacements are incompatible with {self.type.value}")
        compatible_modes = {
            FieldType.STRING: {"exact", "ignore_case"},
            FieldType.PHONE: {"exact", "ignore_case"},
            FieldType.ID_CODE: {"exact", "ignore_case"},
            FieldType.POSTAL_CODE: {"exact", "ignore_case"},
            FieldType.FUZZY_STRING: {"exact", "ignore_case"},
            FieldType.INTEGER: {"exact", "numeric"},
            FieldType.DECIMAL: {"exact", "numeric"},
            FieldType.DATE: {"exact", "date"},
            FieldType.DATETIME: {"exact", "datetime"},
            FieldType.BOOLEAN: {"exact"},
            FieldType.ENUM: {"exact", "ignore_case"},
            FieldType.SET: {"exact", "set"},
            FieldType.JSON: {"exact", "json"},
        }
        if self.compare.mode not in compatible_modes[self.type]:
            raise ValueError(f"{self.compare.mode} comparison is incompatible with {self.type.value}")
        if self.type == FieldType.ENUM and self.compare.mode == "ignore_case":
            folded_values = [value.casefold() for value in self.enum_values]
            if len(folded_values) != len(set(folded_values)):
                raise ValueError(f"enum field {self.name!r} is ambiguous under ignore_case")
            folded_aliases: dict[str, str] = {}
            for alias, target in self.enum_aliases.items():
                folded = alias.casefold()
                owner = folded_aliases.get(folded)
                if owner is not None and owner != target:
                    raise ValueError(f"enum aliases for {self.name!r} are ambiguous under ignore_case")
                folded_aliases[folded] = target
            canonical_by_fold = {value.casefold(): value for value in self.enum_values}
            for alias, target in self.enum_aliases.items():
                canonical = canonical_by_fold.get(alias.casefold())
                if canonical is not None and canonical != target:
                    raise ValueError(f"enum aliases for {self.name!r} remap canonical values under ignore_case")
        numeric = self.type in {FieldType.INTEGER, FieldType.DECIMAL}
        if (self.validation.min is not None or self.validation.max is not None) and not numeric:
            raise ValueError(f"numeric validation bounds are incompatible with {self.type.value}")
        numeric_options_configured = (
            self.compare.absolute_tolerance != 0
            or self.compare.relative_tolerance != 0
            or self.compare.decimal_places is not None
        )
        if numeric_options_configured and not numeric:
            raise ValueError(f"numeric comparison options are incompatible with {self.type.value}")
        if numeric_options_configured and self.compare.mode != "numeric":
            raise ValueError("numeric comparison options require compare.mode=numeric")
        if self.compare.timezone is not None and self.type != FieldType.DATETIME:
            raise ValueError(f"compare.timezone is incompatible with {self.type.value}")
        if self.compare.allow_naive_datetime and self.type != FieldType.DATETIME:
            raise ValueError(f"allow_naive_datetime is incompatible with {self.type.value}")
        if self.parse_formats and self.type not in {FieldType.DATE, FieldType.DATETIME}:
            raise ValueError(f"parse_formats are incompatible with {self.type.value}")
        if self.fill_static_default and self.static_default is None:
            raise ValueError(f"field {self.name!r} enables fill_static_default without static_default")
        if self.fill_static_default:
            from .normalization import parse_value

            parsed = parse_value(self.static_default, self)
            if not parsed.valid:
                raise ValueError(f"static_default for {self.name!r} is invalid: {parsed.error}")
            if parsed.normalized is None:
                raise ValueError(f"static_default for {self.name!r} normalizes to an empty value")
            text = str(parsed.normalized)
            validation = self.validation
            if validation.min_length is not None and len(text) < validation.min_length:
                raise ValueError(f"static_default for {self.name!r} is shorter than min_length")
            if validation.max_length is not None and len(text) > validation.max_length:
                raise ValueError(f"static_default for {self.name!r} is longer than max_length")
            if validation.regex is not None and re.fullmatch(validation.regex, text) is None:
                raise ValueError(f"static_default for {self.name!r} does not match validation regex")
            if validation.min is not None and Decimal(str(parsed.normalized)) < validation.min:
                raise ValueError(f"static_default for {self.name!r} is below validation min")
            if validation.max is not None and Decimal(str(parsed.normalized)) > validation.max:
                raise ValueError(f"static_default for {self.name!r} exceeds validation max")
        if self.type == FieldType.BOOLEAN:
            truthy = {str(value).strip().casefold() for value in self.boolean_true_values}
            falsy = {str(value).strip().casefold() for value in self.boolean_false_values}
            if truthy & falsy:
                raise ValueError(f"boolean true/false aliases overlap for {self.name!r}")
        if self.formula_template is not None:
            formula = self.formula_template
            if len(formula) > 512 or not formula.startswith("="):
                raise ValueError(f"formula_template for {self.name!r} must start with '=' and be at most 512 characters")
            if re.search(r"\[[^\]]+\]|https?://|(?:WEBSERVICE|HYPERLINK|RTD|CALL)\s*\(", formula, re.IGNORECASE):
                raise ValueError(f"formula_template for {self.name!r} contains an external or forbidden function")
            if set(re.findall(r"\{([^{}]+)\}", formula)) - {"row"}:
                raise ValueError(f"formula_template for {self.name!r} only permits the {{row}} placeholder")
        return self


class HeaderRule(StrictModel):
    row: int = Field(default=1, ge=1)
    auto_detect: bool = False
    fuzzy_suggestion_threshold: int = Field(default=92, ge=0, le=100)


class DataRegionRule(StrictModel):
    start_row: int | None = Field(default=None, ge=1)
    end_strategy: Literal["used_range"] = "used_range"
    include_hidden_rows: bool = True


class SheetActions(StrictModel):
    extra_header: Literal["mark_red", "report_only"] = "mark_red"
    missing_required_header: Literal["insert_and_mark_green", "report_only"] = "insert_and_mark_green"
    missing_optional_header: Literal["insert_and_mark_green", "report_only"] = "report_only"
    extra_record: Literal["mark_red", "report_only"] = "mark_red"
    mismatched_value: Literal["mark_yellow", "report_only"] = "mark_yellow"
    invalid_value: Literal["mark_orange", "report_only"] = "mark_orange"
    duplicate_key: Literal["mark_purple", "report_only"] = "mark_purple"
    rename_confirmed_alias: bool = False
    fill_empty_from_standard: bool = False
    overwrite_mismatch: bool = False
    missing_record: Literal["report_only", "append_and_mark_green"] = "report_only"


class CrossFieldRule(StrictModel):
    rule_id: str
    validator: str
    params: dict[str, Any]
    severity: Literal["info", "warning", "error"] = "error"

    @field_validator("validator")
    @classmethod
    def registered_validator(cls, value: str) -> str:
        from .validators import validator_names

        if value not in validator_names():
            raise ValueError(f"unregistered cross-field validator: {value}")
        return value


class SheetRule(StrictModel):
    id: str = Field(min_length=1, max_length=128, pattern=r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")
    name: str
    aliases: list[str] = Field(default_factory=list)
    required: bool = True
    header: HeaderRule = Field(default_factory=HeaderRule)
    data_region: DataRegionRule = Field(default_factory=DataRegionRule)
    primary_key: list[str] = Field(default_factory=list)
    primary_key_mode: Literal["fields", "row_number"] = "fields"
    row_number_field: str = "__row_number__"
    empty_primary_key_action: Literal["report_invalid", "skip_row", "use_row_number"] = "report_invalid"
    columns: list[ColumnRule] = Field(min_length=1)
    cross_field_rules: list[CrossFieldRule] = Field(default_factory=list)
    actions: SheetActions = Field(default_factory=SheetActions)

    @field_validator("name")
    @classmethod
    def valid_excel_sheet_name(cls, value: str) -> str:
        if not value or len(value) > 31 or any(character in value for character in "[]:*?/\\") or value.startswith("'") or value.endswith("'"):
            raise ValueError("worksheet name must be a valid Excel sheet name")
        return value

    @field_validator("aliases")
    @classmethod
    def valid_excel_sheet_aliases(cls, values: list[str]) -> list[str]:
        for value in values:
            cls.valid_excel_sheet_name(value)
        return values

    @model_validator(mode="after")
    def validate_sheet(self) -> "SheetRule":
        names = [column.name for column in self.columns]
        if len(names) != len(set(names)):
            raise ValueError(f"duplicate canonical column in sheet {self.id!r}")
        if self.primary_key_mode == "fields" and not self.primary_key:
            raise ValueError("field primary key mode requires at least one primary key field")
        if self.primary_key_mode == "row_number" and self.primary_key:
            raise ValueError("row_number primary key mode must not also declare primary_key fields")
        missing = set(self.primary_key) - set(names)
        if missing:
            raise ValueError(f"primary key fields do not exist: {sorted(missing)}")
        by_name = {column.name: column for column in self.columns}
        for key in self.primary_key:
            key_rule = by_name[key]
            if not key_rule.required:
                raise ValueError(f"primary key field {key!r} must be required")
            if key_rule.type == FieldType.FUZZY_STRING:
                raise ValueError(f"primary key field {key!r} cannot use fuzzy_string")
            if key_rule.compare.mode != "exact":
                raise ValueError(
                    f"primary key field {key!r} must use compare.mode=exact; "
                    "matching keys cannot use comparison-only equivalence rules"
                )
            if key_rule.compare.formula_mode != "reject":
                raise ValueError(f"primary key field {key!r} must reject formulas")
            lossy = set(key_rule.normalize) & {"collapse_spaces", "remove_group_separator", "remove_currency_symbol"}
            if lossy:
                raise ValueError(f"primary key field {key!r} uses potentially lossy normalizers: {sorted(lossy)}")
        for cross_rule in self.cross_field_rules:
            referenced = {value for key, value in cross_rule.params.items() if key.endswith("_field") and isinstance(value, str)}
            missing_references = referenced - set(names)
            if missing_references:
                raise ValueError(f"cross-field rule {cross_rule.rule_id!r} references missing fields: {sorted(missing_references)}")
        normalized: dict[str, str] = {}
        for column in self.columns:
            for raw in [column.name, column.title, *column.aliases]:
                key = normalize_header(raw)
                owner = normalized.get(key)
                if owner is not None and owner != column.name:
                    raise ValueError(f"header alias collision: {raw!r} maps to {owner!r} and {column.name!r}")
                normalized[key] = column.name
        return self


class WorkbookRule(StrictModel):
    allowed_extensions: list[str] = Field(default_factory=lambda: ["xlsx"])
    preserve_macros: bool = False
    reject_protected: bool = True
    unsupported_feature_action: Literal["allow", "report", "manual_review", "reject"] = "manual_review"
    max_upload_mib: int = Field(default=30, ge=1, le=512)
    max_standard_upload_mib: int = Field(default=64, ge=1, le=1024)
    max_worksheets: int = Field(default=50, ge=1, le=500)
    max_rows_per_sheet: int = Field(default=100_000, ge=1, le=1_000_000)
    max_standard_records: int = Field(default=500_000, ge=1, le=5_000_000)
    max_columns_per_sheet: int = Field(default=200, ge=1, le=16_384)
    max_in_memory_cells: int = Field(default=2_000_000, ge=10_000)
    processing_timeout_seconds: int = Field(default=900, ge=1, le=86_400)
    large_file_action: Literal["report_only", "reject"] = "report_only"


class Colors(StrictModel):
    extra: str = "F4CCCC"
    inserted: str = "D9EAD3"
    mismatch: str = "FFE599"
    invalid: str = "F9CB9C"
    ambiguous: str = "D9D2E9"

    @field_validator("extra", "inserted", "mismatch", "invalid", "ambiguous")
    @classmethod
    def rgb(cls, value: str) -> str:
        value = value.upper()
        if not re.fullmatch(r"[0-9A-F]{6}", value):
            raise ValueError("color must be a six-digit RGB value")
        return value

    @model_validator(mode="after")
    def readable_with_black_text(self) -> "Colors":
        for name in ["extra", "inserted", "mismatch", "invalid", "ambiguous"]:
            value = getattr(self, name)
            channels = [int(value[index:index + 2], 16) / 255 for index in (0, 2, 4)]
            linear = [channel / 12.92 if channel <= 0.04045 else ((channel + 0.055) / 1.055) ** 2.4 for channel in channels]
            luminance = 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2]
            if (luminance + 0.05) / 0.05 < 4.5:
                raise ValueError(f"color {name} does not provide WCAG AA contrast with black text")
        return self


class PaginationConfig(StrictModel):
    type: Literal["page_number"] = "page_number"
    page_param: str = Field(default="page", pattern=r"^[A-Za-z0-9][A-Za-z0-9_.-]{0,127}$")
    size_param: str = Field(default="size", pattern=r"^[A-Za-z0-9][A-Za-z0-9_.-]{0,127}$")
    size: int = Field(default=500, ge=1, le=5000)
    total_json_path: str | None = None
    max_pages: int = Field(default=1000, ge=1, le=10000)

    @field_validator("total_json_path")
    @classmethod
    def valid_total_json_path(cls, value: str | None) -> str | None:
        return validate_simple_json_path(value, field_name="pagination.total_json_path")

    @model_validator(mode="after")
    def distinct_parameter_names(self) -> "PaginationConfig":
        if self.page_param == self.size_param:
            raise ValueError("pagination page_param and size_param must be distinct")
        return self


class StandardSourceConfig(StrictModel):
    type: Literal["upload", "managed_http"] = "upload"
    connection_id: str | None = Field(default=None, pattern=r"^[A-Za-z0-9][A-Za-z0-9_.-]{0,127}$")
    method: Literal["GET", "POST"] = "GET"
    path: str | None = None
    data_json_path: str = "$.data"
    static_parameters: dict[str, Any] = Field(default_factory=dict)
    parameter_mapping: dict[str, str] = Field(default_factory=dict)
    pagination: PaginationConfig | None = None

    @field_validator("path")
    @classmethod
    def valid_managed_path(cls, value: str | None) -> str | None:
        return validate_managed_http_path(value, field_name="standard_source.path") if value is not None else None

    @field_validator("data_json_path")
    @classmethod
    def valid_data_json_path(cls, value: str) -> str:
        return validate_simple_json_path(value, field_name="standard_source.data_json_path") or "$"

    @model_validator(mode="after")
    def validate_source(self) -> "StandardSourceConfig":
        if self.type == "managed_http" and (not self.connection_id or not self.path):
            raise ValueError("managed_http requires connection_id and path")
        if self.type == "upload":
            ignored = (
                self.connection_id is not None
                or self.path is not None
                or self.method != "GET"
                or bool(self.static_parameters)
                or bool(self.parameter_mapping)
                or self.pagination is not None
                or self.data_json_path != "$.data"
            )
            if ignored:
                raise ValueError("upload standard source cannot contain managed HTTP configuration")
            return self

        for name in self.static_parameters:
            validate_parameter_name(str(name), field_name="standard_source.static_parameters key")
        for request_name, task_name in self.parameter_mapping.items():
            validate_parameter_name(str(request_name), field_name="standard_source.parameter_mapping request name")
            validate_parameter_name(str(task_name), field_name="standard_source.parameter_mapping task name")
        overlap = set(self.static_parameters) & set(self.parameter_mapping)
        if overlap:
            raise ValueError(f"static and mapped HTTP parameters overlap: {sorted(overlap)}")
        if self.pagination is not None:
            occupied = set(self.static_parameters) | set(self.parameter_mapping)
            collisions = occupied & {self.pagination.page_param, self.pagination.size_param}
            if collisions:
                raise ValueError(f"pagination parameters overlap configured HTTP parameters: {sorted(collisions)}")
        return self


class RuleSet(StrictModel):

    schema_id: str = Field(min_length=1, max_length=128, pattern=r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")
    schema_version: str
    name: str
    workbook: WorkbookRule = Field(default_factory=WorkbookRule)
    sheets: list[SheetRule] = Field(min_length=1)
    colors: Colors = Field(default_factory=Colors)
    standard_source: StandardSourceConfig = Field(default_factory=StandardSourceConfig)

    @field_validator("schema_version")
    @classmethod
    def semver(cls, value: str) -> str:
        if not re.fullmatch(r"\d+\.\d+\.\d+", value):
            raise ValueError("schema_version must be semantic version x.y.z")
        return value

    @model_validator(mode="after")
    def validate_unique_sheets(self) -> "RuleSet":
        ids = [sheet.id for sheet in self.sheets]
        if len(ids) != len(set(ids)):
            raise ValueError("sheet ids must be unique")
        physical_names: dict[str, str] = {}
        for sheet in self.sheets:
            for physical_name in {sheet.name, *sheet.aliases}:
                key = physical_name.casefold()
                owner = physical_names.get(key)
                if owner is not None and owner != sheet.id:
                    raise ValueError(f"worksheet name or alias {physical_name!r} is shared by sheets {owner!r} and {sheet.id!r}")
                physical_names[key] = sheet.id
        return self

    @property
    def content_sha256(self) -> str:
        payload = json.dumps(self.model_dump(mode="json"), ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def normalize_header(value: Any) -> str:
    import unicodedata

    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKC", text).replace("\u00a0", " ").replace("\r", " ").replace("\n", " ")
    return " ".join(text.strip().split())


class HeaderMapping(StrictModel):
    sheet_id: str
    physical_column: int
    raw_header: str
    normalized_header: str
    canonical_field: str | None = None
    match_type: str
    confidence: float | None = None
    status: str


class Difference(StrictModel):
    difference_id: str
    job_id: str | None = None
    type: DifferenceType
    severity: Literal["info", "warning", "error"] = "error"
    sheet_id: str
    sheet_name: str
    cell: str | None = None
    excel_row: int | None = None
    canonical_field: str | None = None
    business_key: dict[str, Any] | None = None
    excel_raw_value: Any = None
    excel_normalized_value: Any = None
    standard_raw_value: Any = None
    standard_normalized_value: Any = None
    rule_id: str | None = None
    message: str
    render_action: str = "report_only"
    repair_status: str = "not_requested"


class ReportSummary(StrictModel):
    matched_records: int = 0
    extra_records: int = 0
    missing_records: int = 0
    mismatched_cells: int = 0
    validation_errors: int = 0
    differences: int = 0
    repairs_planned: int = 0
    repairs_applied: int = 0
    repair_failures: int = 0


class AuditReport(StrictModel):
    report_version: str = "1.0"
    job_id: str
    created_at: datetime
    schema_id: str
    schema_version: str
    schema_sha256: str
    input_sha256: str
    input_file_name: str | None = None
    input_file_size: int | None = None
    standard_snapshot_id: str
    standard_sha256: str
    standard_source_metadata: dict[str, Any] = Field(default_factory=dict)
    header_mappings: list[HeaderMapping] = Field(default_factory=list)
    # Large report-mode comparisons may keep the complete difference stream in
    # a disk-backed Sequence. Validation occurs when each Difference is created;
    # preserving the sequence here avoids materializing it again in the report.
    differences: SkipValidation[Sequence[Difference]] = Field(default_factory=list)
    summary: ReportSummary = Field(default_factory=ReportSummary)
    warnings: list[str] = Field(default_factory=list)
    workbook_structure: list[dict[str, Any]] = Field(default_factory=list)
    header_summary: dict[str, int] = Field(default_factory=dict)
    record_summary: dict[str, int] = Field(default_factory=dict)
    field_statistics: dict[str, dict[str, Any]] = Field(default_factory=dict)
    data_quality_summary: dict[str, int] = Field(default_factory=dict)
    output_sha256: str | None = None

    @field_validator("differences", mode="before")
    @classmethod
    def validate_difference_sequence(cls, value: Any) -> Sequence[Difference]:
        if isinstance(value, SpillableSequence):
            return value
        if isinstance(value, (str, bytes, bytearray)) or not isinstance(value, Sequence):
            raise ValueError("differences must be a sequence")
        return [item if isinstance(item, Difference) else Difference.model_validate(item) for item in value]
