import os


def execute_worker_job(job_id):
    project_id = os.getenv("PROJECT_ID", "").strip()
    region = os.getenv("REGION", os.getenv("GOOGLE_CLOUD_REGION", "")).strip()
    job_name = os.getenv("WORKER_JOB_NAME", "").strip()
    timeout = os.getenv("WORKER_TASK_TIMEOUT", "").strip()

    if not project_id:
        raise RuntimeError("Thiếu biến môi trường PROJECT_ID")
    if not region:
        raise RuntimeError("Thiếu biến môi trường REGION")
    if not job_name:
        raise RuntimeError("Thiếu biến môi trường WORKER_JOB_NAME")

    import google.auth
    from google.auth.transport.requests import AuthorizedSession

    credentials, _ = google.auth.default(scopes=["https://www.googleapis.com/auth/cloud-platform"])
    session = AuthorizedSession(credentials)
    url = f"https://run.googleapis.com/v2/projects/{project_id}/locations/{region}/jobs/{job_name}:run"
    env = [
        {"name": "JOB_ID", "value": job_id},
        {"name": "DOCX_BUILDER_MODE", "value": "cloud"},
    ]
    payload = {
        "overrides": {
            "containerOverrides": [
                {
                    "env": env,
                }
            ]
        }
    }
    if timeout:
        payload["overrides"]["timeout"] = timeout

    response = session.post(url, json=payload, timeout=30)
    response.raise_for_status()
    return response.json()
