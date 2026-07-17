import os
import tempfile
import traceback
from pathlib import Path

from env_loader import load_dotenv_file


load_dotenv_file()


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
        excel_files = manifest.get("excel_files") or [
            {
                "id": manifest.get("excel_name", ""),
                "name": manifest.get("excel_name", ""),
                "object": manifest.get("excel_object", ""),
            }
        ]
        excel_paths = {}
        for excel_file in excel_files:
            if not excel_file.get("name") or not excel_file.get("object"):
                continue
            excel_paths[excel_file.get("id") or excel_file["name"]] = workdir / excel_file["name"]
        output_path = workdir / manifest["output_name"]

        store.update(job_id, status="running", progress=6, message="Đang tải file đầu vào")
        storage.download_file(manifest["docx_object"], docx_path)
        for excel_file in excel_files:
            local_path = excel_paths.get(excel_file.get("id") or excel_file.get("name"))
            if local_path:
                storage.download_file(excel_file["object"], local_path)

        store.update(job_id, status="running", progress=8, message="Đang nạp engine generate")
        from generation_service import generate_docx_from_payload

        def progress(percent, message):
            store.update(job_id, status="running", progress=percent, message=message)

        generate_docx_from_payload(
            docx_path=docx_path,
            excel_path=excel_paths,
            output_path=output_path,
            payload=manifest["payload"],
            progress_callback=progress,
        )
        preview_name = manifest.get("preview_name") or f"{output_path.stem}_preview.pdf"
        preview_path = workdir / preview_name
        preview_error = ""
        preview_available = False
        try:
            store.update(job_id, status="running", progress=98, message="Đang tạo preview PDF")
            from preview_service import PreviewConversionError, convert_docx_to_pdf

            generated_preview = convert_docx_to_pdf(output_path, preview_path.parent)
            if generated_preview != preview_path:
                generated_preview.replace(preview_path)
            preview_available = preview_path.exists()
        except PreviewConversionError as exc:
            preview_error = str(exc)
        storage.upload_file(
            output_path,
            manifest["output_object"],
            content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        if preview_available and manifest.get("preview_object"):
            storage.upload_file(
                preview_path,
                manifest["preview_object"],
                content_type="application/pdf",
            )
        store.update(
            job_id,
            status="done",
            progress=100,
            message="Hoàn tất",
            output_object=manifest["output_object"],
            output_name=manifest["output_name"],
            preview_object=manifest.get("preview_object", ""),
            preview_name=preview_name,
            preview_available=preview_available,
            preview_error=preview_error,
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
