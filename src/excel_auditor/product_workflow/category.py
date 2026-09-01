from __future__ import annotations

from collections import defaultdict
from typing import Any

from rapidfuzz import fuzz, process

from ..models import normalize_header
from .models import CategoryDefinition, CategoryResolution, MappingCandidate


def _key(value: Any) -> str:
    return normalize_header(value).casefold()


def resolve_categories(
    rows: list[dict[str, Any]],
    categories: list[CategoryDefinition],
    *,
    source_field: str = "merchant_category",
    id_field: str | None = "platform_category_id",
    first_excel_row: int = 2,
    excel_rows: list[int] | None = None,
    fuzzy_threshold: int = 95,
) -> list[CategoryResolution]:
    if excel_rows is not None and len(excel_rows) != len(rows):
        raise ValueError("excel_rows must align one-to-one with category rows")
    by_id = {category.category_id: category for category in categories if category.active}
    if len(by_id) != sum(1 for category in categories if category.active):
        raise ValueError("active platform category_id values must be unique")
    names: dict[str, set[str]] = defaultdict(set)
    for category in by_id.values():
        for value in (category.name, *category.aliases, *category.path):
            names[_key(value)].add(category.category_id)

    results: list[CategoryResolution] = []
    choices = list(names)
    for offset, row in enumerate(rows):
        excel_row = excel_rows[offset] if excel_rows is not None else first_excel_row + offset
        raw_id = normalize_header(row.get(id_field)) if id_field else ""
        if raw_id and raw_id in by_id:
            category = by_id[raw_id]
            results.append(CategoryResolution(
                excel_row=excel_row,
                raw_category=normalize_header(row.get(source_field)) or None,
                category_id=category.category_id,
                category_name=category.name,
                status="resolved",
                match_type="id",
                confidence=100,
            ))
            continue

        raw_name = normalize_header(row.get(source_field))
        if not raw_name:
            results.append(CategoryResolution(
                excel_row=excel_row,
                status="unresolved",
                match_type="missing",
            ))
            continue
        owners = names.get(_key(raw_name), set())
        if len(owners) == 1:
            category = by_id[next(iter(owners))]
            results.append(CategoryResolution(
                excel_row=excel_row,
                raw_category=raw_name,
                category_id=category.category_id,
                category_name=category.name,
                status="resolved",
                match_type="exact",
                confidence=100,
            ))
            continue

        candidates_by_id: dict[str, MappingCandidate] = {}
        if len(owners) > 1:
            for owner in owners:
                category = by_id[owner]
                candidates_by_id[owner] = MappingCandidate(
                    field_id=owner,
                    title=category.name,
                    confidence=100,
                    match_value=raw_name,
                )
            match_type = "ambiguous"
        else:
            for choice, score, _index in process.extract(
                _key(raw_name), choices, scorer=fuzz.ratio, score_cutoff=fuzzy_threshold, limit=12
            ):
                for owner in names[choice]:
                    category = by_id[owner]
                    candidate = MappingCandidate(
                        field_id=owner,
                        title=category.name,
                        confidence=score,
                        match_value=choice,
                    )
                    current = candidates_by_id.get(owner)
                    if current is None or candidate.confidence > current.confidence:
                        candidates_by_id[owner] = candidate
            match_type = "fuzzy_suggestion"
        candidates = sorted(candidates_by_id.values(), key=lambda item: (-item.confidence, item.field_id))[:3]
        results.append(CategoryResolution(
            excel_row=excel_row,
            raw_category=raw_name,
            status="manual_review" if candidates else "unresolved",
            match_type=match_type if candidates else "missing",
            confidence=candidates[0].confidence if candidates else None,
            candidates=candidates,
        ))
    return results
