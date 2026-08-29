from __future__ import annotations

import argparse
from datetime import datetime
import json
import os
from pathlib import Path
import re
import sys
import tempfile
from typing import Any


GIT_SHA = re.compile(r"^[0-9a-f]{40}$")
RELEASE_TAG = re.compile(r"^v\d+\.\d+\.\d+(?:-[0-9A-Za-z]+(?:[.-][0-9A-Za-z]+)*)?$")
IMAGE_REF = re.compile(
    r"^ghcr\.io/(?:[a-z0-9][a-z0-9._-]*/)+excel-auditor-(api|web):"
    r"([0-9a-f]{40})@sha256:([0-9a-f]{64})$"
)


def validate_release_manifest(payload: Any) -> dict[str, Any]:
    if not isinstance(payload, dict):
        raise ValueError("root must be an object")
    required = {
        "manifest_version",
        "release_tag",
        "git_sha",
        "workflow_run_id",
        "generated_at",
        "images",
        "sboms",
        "deployment",
    }
    missing = required - payload.keys()
    unknown = payload.keys() - required
    if missing:
        raise ValueError(f"missing fields: {sorted(missing)}")
    if unknown:
        raise ValueError(f"unknown fields: {sorted(unknown)}")
    if payload["manifest_version"] != "1.0":
        raise ValueError("manifest_version must be 1.0")
    release_tag = _required_string(payload, "release_tag")
    git_sha = _required_string(payload, "git_sha")
    if not RELEASE_TAG.fullmatch(release_tag):
        raise ValueError("release_tag must be a v-prefixed semantic version")
    if not GIT_SHA.fullmatch(git_sha):
        raise ValueError("git_sha must be a lowercase 40-character SHA")
    run_id = payload["workflow_run_id"]
    if isinstance(run_id, bool) or not isinstance(run_id, (int, str)) or not str(run_id).isdigit() or int(run_id) < 1:
        raise ValueError("workflow_run_id must be a positive integer")
    generated_at = _required_string(payload, "generated_at")
    try:
        parsed_time = datetime.fromisoformat(generated_at.replace("Z", "+00:00"))
    except ValueError as exc:
        raise ValueError("generated_at must be an RFC 3339 timestamp") from exc
    if parsed_time.tzinfo is None:
        raise ValueError("generated_at must include a timezone")

    images = _object_with_keys(payload, "images", {"api", "web"})
    for kind in ("api", "web"):
        reference = _required_string(images, kind)
        match = IMAGE_REF.fullmatch(reference)
        if match is None or match.group(1) != kind:
            raise ValueError(f"images.{kind} must be an immutable GHCR {kind} reference")
        if match.group(2) != git_sha:
            raise ValueError(f"images.{kind} tag must equal git_sha")

    sboms = _object_with_keys(payload, "sboms", {"api", "web"})
    for kind in ("api", "web"):
        expected = f"excel-auditor-{kind}-{git_sha}.spdx.json"
        if sboms[kind] != expected:
            raise ValueError(f"sboms.{kind} must be {expected}")

    deployment = _object_with_keys(payload, "deployment", {"compose_file", "environment_file"})
    if deployment["compose_file"] != "deploy/docker-compose.prod.yaml":
        raise ValueError("deployment.compose_file is invalid")
    if deployment["environment_file"] != "release.env":
        raise ValueError("deployment.environment_file is invalid")
    return payload


def render_release_environment(payload: dict[str, Any]) -> str:
    validated = validate_release_manifest(payload)
    return "".join(
        [
            f"API_IMAGE={validated['images']['api']}\n",
            f"WEB_IMAGE={validated['images']['web']}\n",
            f"RELEASE_GIT_SHA={validated['git_sha']}\n",
            f"RELEASE_TAG={validated['release_tag']}\n",
        ]
    )


def write_release_environment(path: Path, payload: dict[str, Any]) -> None:
    content = render_release_environment(payload)
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(prefix=f".{path.name}.", dir=path.parent)
    temporary = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(content)
            handle.flush()
            os.fsync(handle.fileno())
        temporary.replace(path)
    except BaseException:
        temporary.unlink(missing_ok=True)
        raise


def _required_string(payload: dict[str, Any], key: str) -> str:
    value = payload.get(key)
    if not isinstance(value, str) or not value:
        raise ValueError(f"{key} must be a non-empty string")
    return value


def _object_with_keys(payload: dict[str, Any], key: str, expected: set[str]) -> dict[str, Any]:
    value = payload.get(key)
    if not isinstance(value, dict) or set(value) != expected:
        raise ValueError(f"{key} must contain exactly {sorted(expected)}")
    return value


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Validate an immutable Excel Auditor release manifest")
    parser.add_argument("manifest", type=Path)
    parser.add_argument("--env-output", type=Path)
    args = parser.parse_args(argv)
    try:
        payload = json.loads(args.manifest.read_text(encoding="utf-8"))
        validate_release_manifest(payload)
        if args.env_output is not None:
            write_release_environment(args.env_output, payload)
    except (OSError, json.JSONDecodeError, ValueError) as exc:
        print(f"INVALID_RELEASE_MANIFEST: {exc}", file=sys.stderr)
        return 2
    print(f"valid release manifest: {payload['release_tag']} {payload['git_sha']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
