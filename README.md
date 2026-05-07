# Tool Insert Excel

Web app va tool Python de tao file DOCX template, chon heading trong file Word va chen bang tu Excel.

## Chay local

```bash
python3 -m pip install -r requirements.txt
python3 -m uvicorn web_app:app --host 127.0.0.1 --port 8001 --reload
```

Mo trinh duyet tai:

```text
http://127.0.0.1:8001/
```

## Deploy

Project co san `Procfile`, `runtime.txt`, `requirements.txt` de deploy len Railway/Heroku-style platform.
