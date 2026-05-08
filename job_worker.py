import os
import tempfile
import traceback
from pathlib import Path


def main():
    job_id = os.getenv("JOB_ID", "").strip()
    if not job_id:
        raise RuntimeError("Thiếu biến môi trường JOB_ID")

    from job_store import get_job_store
    from storage_backend import GCSStorage, job_manifest_object

    store = get_job_store()
    workdir = Path(tempfile.mkdtemp(prefix=f"docx_job_{job_id}_"))

    try:
        store.update(job_id, status="running", progress=2, message="Worker đã khởi động", error="")

        storage = GCSStorage()
        store.update(job_id, status="running", progress=4, message="Đang tải cấu hình job")
        manifest = storage.download_json(job_manifest_object(job_id))

        docx_path = workdir / manifest["docx_name"]
        excel_path = workdir / manifest["excel_name"]
        output_path = workdir / manifest["output_name"]

        store.update(job_id, status="running", progress=6, message="Đang tải file đầu vào")
        storage.download_file(manifest["docx_object"], docx_path)
        storage.download_file(manifest["excel_object"], excel_path)

        store.update(job_id, status="running", progress=8, message="Đang nạp engine generate")
        from generation_service import generate_docx_from_payload

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
