"""Dynamic product-catalog normalization primitives."""

from .catalog import CatalogAdapter, InMemoryCatalogAdapter, ManagedHttpCatalogAdapter
from .category import resolve_categories
from .mapping import map_product_headers
from .normalizer import normalize_product_workbook
from .models import (
    CatalogFieldDefinition,
    CatalogFieldSource,
    CatalogSchemaSnapshot,
    CategoryDefinition,
    CategoryResolution,
    DynamicSchemaPlan,
    FieldMapping,
    MappingCandidate,
    NormalizedCategorySheet,
    PlannedField,
    ProductCellIssue,
    ProductNormalizationResult,
    ProductReviewDecision,
    ReviewItem,
    UnresolvedProductRow,
)
from .schema_builder import build_dynamic_schema

__all__ = [
    "CatalogFieldDefinition",
    "CatalogFieldSource",
    "CatalogSchemaSnapshot",
    "CategoryDefinition",
    "CategoryResolution",
    "CatalogAdapter",
    "DynamicSchemaPlan",
    "FieldMapping",
    "InMemoryCatalogAdapter",
    "ManagedHttpCatalogAdapter",
    "MappingCandidate",
    "NormalizedCategorySheet",
    "PlannedField",
    "ProductCellIssue",
    "ProductNormalizationResult",
    "ProductReviewDecision",
    "ReviewItem",
    "UnresolvedProductRow",
    "build_dynamic_schema",
    "map_product_headers",
    "normalize_product_workbook",
    "resolve_categories",
]
