import os
import tempfile
import traceback
from pathlib import Path

from generation_service import generate_docx_from_payload
from job_store import get_job_store
from storage_backend import GCSStorage, job_manifest_object


def main():
    job_id = os.getenv("JOB_ID", "").strip()
    if not job_id:
        raise RuntimeError("Thiếu biến môi trường JOB_ID")

    store = get_job_store()
    storage = GCSStorage()
    workdir = Path(tempfile.mkdtemp(prefix=f"docx_job_{job_id}_"))

    try:
        store.update(job_id, status="running", error="")
        manifest = storage.download_json(job_manifest_object(job_id))

        docx_path = workdir / manifest["docx_name"]
        excel_path = workdir / manifest["excel_name"]
        output_path = workdir / manifest["output_name"]

        storage.download_file(manifest["docx_object"], docx_path)
        storage.download_file(manifest["excel_object"], excel_path)

        def progress(percent, message):
            store.update(job_id, status="running", progress=percent, message=message)

        generate_docx_from_payload(
            docx_path=docx_path,
            excel_path=excel_path,
            output_path=output_path,
            payload=manifest["payload"],
            progress_callback=progress,
        )
        storage.upload_file(
            output_path,
            manifest["output_object"],
            content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        store.update(
            job_id,
            status="done",
            progress=100,
            message="Hoàn tất",
            output_object=manifest["output_object"],
            output_name=manifest["output_name"],
        )
    except Exception as exc:
        store.update(
            job_id,
            status="failed",
            error=str(exc),
            traceback=traceback.format_exc(),
        )
        raise


if __name__ == "__main__":
    main()
