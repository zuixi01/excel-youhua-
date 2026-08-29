from __future__ import annotations

import json
from pathlib import Path

import pytest
import yaml

from excel_auditor.release_manifest import main, render_release_environment, validate_release_manifest


SHA = "a" * 40
DIGESTS = {"api": "b" * 64, "web": "c" * 64}


def _manifest() -> dict:
    return {
        "manifest_version": "1.0",
        "release_tag": "v1.2.3-rc.1",
        "git_sha": SHA,
        "workflow_run_id": "12345",
        "generated_at": "2026-08-29T13:00:00Z",
        "images": {
            kind: f"ghcr.io/example/excel-auditor-{kind}:{SHA}@sha256:{DIGESTS[kind]}"
            for kind in ("api", "web")
        },
        "sboms": {
            kind: f"excel-auditor-{kind}-{SHA}.spdx.json"
            for kind in ("api", "web")
        },
        "deployment": {
            "compose_file": "deploy/docker-compose.prod.yaml",
            "environment_file": "release.env",
        },
    }


def test_release_manifest_produces_separate_digest_pinned_images():
    payload = validate_release_manifest(_manifest())
    environment = render_release_environment(payload)
    assert environment.splitlines() == [
        f"API_IMAGE=ghcr.io/example/excel-auditor-api:{SHA}@sha256:{DIGESTS['api']}",
        f"WEB_IMAGE=ghcr.io/example/excel-auditor-web:{SHA}@sha256:{DIGESTS['web']}",
        f"RELEASE_GIT_SHA={SHA}",
        "RELEASE_TAG=v1.2.3-rc.1",
    ]
    assert "latest" not in environment


@pytest.mark.parametrize(
    ("mutation", "message"),
    [
        (lambda manifest: manifest.update(release_tag="release-1"), "semantic version"),
        (lambda manifest: manifest["images"].update(api="ghcr.io/example/excel-auditor-api:latest"), "immutable GHCR"),
        (
            lambda manifest: manifest["images"].update(
                web=f"ghcr.io/example/excel-auditor-web:{'d' * 40}@sha256:{DIGESTS['web']}"
            ),
            "tag must equal git_sha",
        ),
        (lambda manifest: manifest.update(unexpected=True), "unknown fields"),
        (lambda manifest: manifest["sboms"].update(api="other.spdx.json"), "sboms.api"),
    ],
)
def test_release_manifest_rejects_untraceable_or_mutable_inputs(mutation, message):
    manifest = _manifest()
    mutation(manifest)
    with pytest.raises(ValueError, match=message):
        validate_release_manifest(manifest)


def test_release_manifest_cli_writes_validated_environment(tmp_path, capsys):
    manifest_path = tmp_path / "release-manifest.json"
    environment_path = tmp_path / "release.env"
    manifest_path.write_text(json.dumps(_manifest()), encoding="utf-8")
    assert main([str(manifest_path), "--env-output", str(environment_path)]) == 0
    assert environment_path.read_text(encoding="utf-8") == render_release_environment(_manifest())
    assert "valid release manifest" in capsys.readouterr().out

    invalid_path = tmp_path / "invalid.json"
    invalid_path.write_text("{}", encoding="utf-8")
    assert main([str(invalid_path), "--env-output", str(environment_path)]) == 2
    assert environment_path.read_text(encoding="utf-8") == render_release_environment(_manifest())


def test_production_compose_requires_independent_full_image_references():
    compose = yaml.safe_load(Path("deploy/docker-compose.prod.yaml").read_text(encoding="utf-8"))
    services = compose["services"]
    assert services["api"]["image"] == "${API_IMAGE:?set API_IMAGE to an immutable tag or digest}"
    assert services["worker"]["image"] == services["api"]["image"]
    assert services["web"]["image"] == "${WEB_IMAGE:?set WEB_IMAGE to an immutable tag or digest}"
    serialized = json.dumps(compose)
    assert "IMAGE_TAG" not in serialized and "REGISTRY" not in serialized and "latest" not in serialized


def test_release_identity_blocks_expensive_jobs_until_tag_validation():
    workflow = yaml.load(Path(".github/workflows/release.yml").read_text(encoding="utf-8"), Loader=yaml.BaseLoader)
    jobs = workflow["jobs"]
    assert jobs["tests"]["needs"] == "release-identity"
    assert jobs["performance"]["needs"] == "release-identity"
    assert set(jobs["images"]["needs"]) == {"release-identity", "tests", "performance"}
    identity_script = jobs["release-identity"]["steps"][1]["run"]
    assert "git rev-list" in identity_script and "GITHUB_SHA" in identity_script
