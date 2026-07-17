import json
import os
from pathlib import Path


def cloud_mode():
    return os.getenv("DOCX_BUILDER_MODE", "local").strip().lower() == "cloud"


def gcs_bucket_name():
    bucket = os.getenv("GCS_BUCKET", "").strip()
    if not bucket:
        raise RuntimeError("Thiếu biến môi trường GCS_BUCKET")
    return bucket


class GCSStorage:
    def __init__(self, bucket_name=None):
        from google.cloud import storage

        self.client = storage.Client()
        self.bucket = self.client.bucket(bucket_name or gcs_bucket_name())

    def upload_file(self, local_path, object_name, content_type=None):
        blob = self.bucket.blob(object_name)
        blob.upload_from_filename(str(local_path), content_type=content_type)
        return object_name

    def download_file(self, object_name, local_path):
        local_path = Path(local_path)
        local_path.parent.mkdir(parents=True, exist_ok=True)
        self.bucket.blob(object_name).download_to_filename(str(local_path))
        return local_path

    def upload_json(self, object_name, data):
        blob = self.bucket.blob(object_name)
        blob.upload_from_string(
            json.dumps(data, ensure_ascii=False),
            content_type="application/json; charset=utf-8",
        )
        return object_name

    def download_json(self, object_name):
        text = self.bucket.blob(object_name).download_as_text(encoding="utf-8")
        return json.loads(text)

    def exists(self, object_name):
        return self.bucket.blob(object_name).exists()


def session_manifest_object(session_id):
    return f"sessions/{session_id}/session.json"


def job_manifest_object(job_id):
    return f"jobs/{job_id}/manifest.json"


def job_output_object(job_id, filename):
    safe = filename.replace("/", "_")
    return f"jobs/{job_id}/output/{safe}"


def job_preview_object(job_id, filename):
    safe = filename.replace("/", "_")
    return f"jobs/{job_id}/preview/{safe}"
