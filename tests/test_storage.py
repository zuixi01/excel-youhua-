from pathlib import Path

from excel_auditor.storage import ArtifactStore


class MemoryStore(ArtifactStore):
    def __init__(self):
        self.objects = {}

    def put_file(self, key: str, path: Path) -> str:
        self.objects[key] = path.read_bytes()
        return key

    def download_url(self, key: str, expires_seconds: int = 300) -> str:
        return f"https://download.invalid/{key}"


def test_artifact_store_contract(tmp_path):
    path = tmp_path / "artifact.json"
    path.write_text("{}", encoding="utf-8")
    store = MemoryStore()
    key = store.put_file("jobs/one/artifact.json", path)
    assert store.objects[key] == b"{}"
    assert store.download_url(key).endswith(key)
