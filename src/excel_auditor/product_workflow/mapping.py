from __future__ import annotations

from collections import defaultdict

from rapidfuzz import fuzz, process

from ..models import normalize_header
from .models import CatalogFieldDefinition, FieldMapping, MappingCandidate


def _key(value: str) -> str:
    return normalize_header(value).casefold()


def map_product_headers(
    headers: list[str],
    fields: list[CatalogFieldDefinition],
    *,
    confirmed_aliases: dict[str, str] | None = None,
    forced_extra_columns: set[int] | None = None,
    fuzzy_threshold: int = 92,
    max_candidates: int = 3,
) -> list[FieldMapping]:
    """Map exact names deterministically; fuzzy matches remain human-review suggestions."""
    by_id = {field.field_id: field for field in fields}
    if len(by_id) != len(fields):
        raise ValueError("mapping target field_id values must be unique")

    token_owners: dict[str, set[str]] = defaultdict(set)
    token_display: dict[str, str] = {}
    for field in fields:
        values = [field.field_id, field.title, *field.aliases]
        if field.attribute_id is not None:
            values.append(field.attribute_id)
        for value in values:
            token = _key(value)
            token_owners[token].add(field.field_id)
            token_display.setdefault(token, value)

    confirmed: dict[str, str] = {}
    for alias, field_id in (confirmed_aliases or {}).items():
        if field_id not in by_id:
            raise ValueError(f"confirmed alias targets unknown field_id: {field_id}")
        token = _key(alias)
        owner = confirmed.get(token)
        if owner is not None and owner != field_id:
            raise ValueError(f"confirmed alias is ambiguous after normalization: {alias!r}")
        confirmed[token] = field_id

    results: list[FieldMapping] = []
    accepted_columns: dict[str, int] = {}
    choices = list(token_owners)
    for column, raw_header in enumerate(headers, start=1):
        normalized = normalize_header(raw_header)
        token = normalized.casefold()
        if column in (forced_extra_columns or set()):
            results.append(FieldMapping(
                physical_column=column,
                raw_header=str(raw_header),
                normalized_header=normalized,
                match_type="unmapped",
                status="unmapped",
            ))
            continue
        field_id = confirmed.get(token)
        match_type = "confirmed_alias" if field_id is not None else "exact"
        owners = {field_id} if field_id is not None else token_owners.get(token, set())
        if len(owners) == 1:
            resolved_id = next(iter(owners))
            if resolved_id in accepted_columns:
                results.append(FieldMapping(
                    physical_column=column,
                    raw_header=str(raw_header),
                    normalized_header=normalized,
                    field_id=resolved_id,
                    match_type="duplicate",
                    status="manual_review",
                    confidence=100,
                ))
            else:
                accepted_columns[resolved_id] = column
                results.append(FieldMapping(
                    physical_column=column,
                    raw_header=str(raw_header),
                    normalized_header=normalized,
                    field_id=resolved_id,
                    match_type=match_type,
                    status="accepted",
                    confidence=100,
                ))
            continue
        if len(owners) > 1:
            candidates = [
                MappingCandidate(field_id=owner, title=by_id[owner].title, confidence=100, match_value=normalized)
                for owner in sorted(owners)
            ]
            results.append(FieldMapping(
                physical_column=column,
                raw_header=str(raw_header),
                normalized_header=normalized,
                match_type="exact",
                status="manual_review",
                confidence=100,
                candidates=candidates,
            ))
            continue

        best_by_field: dict[str, MappingCandidate] = {}
        for choice, score, _index in process.extract(
            token,
            choices,
            scorer=fuzz.ratio,
            score_cutoff=fuzzy_threshold,
            limit=max(max_candidates * 3, max_candidates),
        ):
            for owner in token_owners[choice]:
                candidate = MappingCandidate(
                    field_id=owner,
                    title=by_id[owner].title,
                    confidence=score,
                    match_value=token_display[choice],
                )
                current = best_by_field.get(owner)
                if current is None or candidate.confidence > current.confidence:
                    best_by_field[owner] = candidate
        candidates = sorted(best_by_field.values(), key=lambda item: (-item.confidence, item.field_id))[:max_candidates]
        if candidates:
            results.append(FieldMapping(
                physical_column=column,
                raw_header=str(raw_header),
                normalized_header=normalized,
                match_type="fuzzy_suggestion",
                status="manual_review",
                confidence=candidates[0].confidence,
                candidates=candidates,
            ))
        else:
            results.append(FieldMapping(
                physical_column=column,
                raw_header=str(raw_header),
                normalized_header=normalized,
                match_type="unmapped",
                status="unmapped",
            ))
    return results
