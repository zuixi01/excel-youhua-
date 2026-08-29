from __future__ import annotations

import mimetypes
from abc import ABC, abstractmethod
from pathlib import Path

import boto3


class ArtifactStore(ABC):
    @abstractmethod
    def put_file(self, key: str, path: Path) -> str:
        raise NotImplementedError

    @abstractmethod
    def download_url(self, key: str, expires_seconds: int = 300) -> str:
        raise NotImplementedError

    def ping(self) -> None:
        return None

    def delete(self, key: str) -> None:
        return None


class S3ArtifactStore(ArtifactStore):
    def __init__(self, bucket: str, endpoint_url: str | None = None, region: str | None = None, server_side_encryption: str | None = "AES256") -> None:
        self.bucket = bucket
        self.server_side_encryption = server_side_encryption
        self.client = boto3.client("s3", endpoint_url=endpoint_url, region_name=region)

    def put_file(self, key: str, path: Path) -> str:
        content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        extra = {"ContentType": content_type}
        if self.server_side_encryption:
            extra["ServerSideEncryption"] = self.server_side_encryption
        self.client.upload_file(str(path), self.bucket, key, ExtraArgs=extra)
        return key

    def download_url(self, key: str, expires_seconds: int = 300) -> str:
        return self.client.generate_presigned_url("get_object", Params={"Bucket": self.bucket, "Key": key}, ExpiresIn=expires_seconds)

    def ping(self) -> None:
        self.client.head_bucket(Bucket=self.bucket)

    def delete(self, key: str) -> None:
        self.client.delete_object(Bucket=self.bucket, Key=key)
