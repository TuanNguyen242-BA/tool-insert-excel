import json
import os
from datetime import datetime, timezone
from pathlib import Path


def utc_now_iso():
    return datetime.now(timezone.utc).isoformat()


class LocalJobStore:
    def __init__(self, root=None):
        self.root = Path(root or os.getenv("DOCX_BUILDER_JOBS_DIR", "/tmp/docx_builder_jobs"))
        self.root.mkdir(parents=True, exist_ok=True)

    def _path(self, job_id):
        return self.root / f"{job_id}.json"

    def create(self, job_id, data):
        now = utc_now_iso()
        record = {
            **data,
            "job_id": job_id,
            "created_at": now,
            "updated_at": now,
        }
        self._path(job_id).write_text(json.dumps(record, ensure_ascii=False), encoding="utf-8")
        return record

    def update(self, job_id, **fields):
        record = self.get(job_id) or {"job_id": job_id, "created_at": utc_now_iso()}
        record.update(fields)
        record["updated_at"] = utc_now_iso()
        self._path(job_id).write_text(json.dumps(record, ensure_ascii=False), encoding="utf-8")
        return record

    def get(self, job_id):
        path = self._path(job_id)
        if not path.exists():
            return None
        return json.loads(path.read_text(encoding="utf-8"))


class FirestoreJobStore:
    def __init__(self, collection=None):
        from google.cloud import firestore

        self.client = firestore.Client()
        self.collection = self.client.collection(collection or os.getenv("FIRESTORE_JOBS_COLLECTION", "docx_jobs"))

    def create(self, job_id, data):
        now = utc_now_iso()
        record = {
            **data,
            "job_id": job_id,
            "created_at": now,
            "updated_at": now,
        }
        self.collection.document(job_id).set(record)
        return record

    def update(self, job_id, **fields):
        fields["updated_at"] = utc_now_iso()
        self.collection.document(job_id).set(fields, merge=True)
        record = self.get(job_id)
        return record or {"job_id": job_id, **fields}

    def get(self, job_id):
        snapshot = self.collection.document(job_id).get()
        if not snapshot.exists:
            return None
        return snapshot.to_dict()


def get_job_store():
    mode = os.getenv("DOCX_BUILDER_MODE", "local").strip().lower()
    if mode == "cloud":
        return FirestoreJobStore()
    return LocalJobStore()
