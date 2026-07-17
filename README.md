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

### Preview Word bang LibreOffice

Tinh nang preview can `soffice` cua LibreOffice de convert DOCX sang PDF.

Khi chay bang Docker/Cloud Run, image da cai LibreOffice san qua `Dockerfile`.

Khi chay local Windows, co 3 cach:

1. Cai LibreOffice binh thuong. App se tu tim:

```text
C:\Program Files\LibreOffice\program\soffice.exe
```

2. Dat LibreOffice portable vao mot trong cac thu muc sau cua project:

```text
tools\libreoffice\program\soffice.exe
tools\LibreOfficePortable\App\libreoffice\program\soffice.exe
```

3. Hoac tao file `.env` o thu muc project va dat duong dan tuy bien:

```text
LIBREOFFICE_BIN=C:\Program Files\LibreOffice\program\soffice.exe
```

```text
LIBREOFFICE_BIN=D:\Apps\LibreOfficePortable\App\libreoffice\program\soffice.exe
```

## Deploy

Project co san `Procfile`, `runtime.txt`, `requirements.txt` de deploy len Railway/Heroku-style platform.

Voi file Excel lon, dung kien truc Google Cloud Run Service + Cloud Run Jobs. Xem chi tiet trong `DEPLOY_GCP.md`.
