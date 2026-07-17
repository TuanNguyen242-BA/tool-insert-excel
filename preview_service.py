import os
import shutil
import subprocess
import tempfile
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent


class PreviewConversionError(RuntimeError):
    pass


def find_libreoffice():
    configured = os.getenv("LIBREOFFICE_BIN", "").strip()
    candidates = [
        configured,
        BASE_DIR / "tools" / "libreoffice" / "program" / "soffice.exe",
        BASE_DIR / "tools" / "LibreOfficePortable" / "App" / "libreoffice" / "program" / "soffice.exe",
        BASE_DIR / "tools" / "LibreOfficePortable" / "LibreOfficePortable.exe",
        r"C:\PortableApps\LibreOfficePortable\App\libreoffice\program\soffice.exe",
        r"C:\PortableApps\LibreOfficePortable\LibreOfficePortable.exe",
        r"C:\PortableAppsLibreOfficePortable\App\libreoffice\program\soffice.exe",
        r"C:\PortableAppsLibreOfficePortable\LibreOfficePortable.exe",
        r"C:\LOPortable\LibreOfficePortable\App\libreoffice\program\soffice.exe",
        r"C:\LOPortable\App\libreoffice\program\soffice.exe",
        shutil.which("soffice") or "",
        shutil.which("libreoffice") or "",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        path = Path(candidate) if candidate else None
        if path and path.exists():
            return str(path)
    raise PreviewConversionError(
        "Khong tim thay LibreOffice/soffice de tao preview PDF. "
        "Hay cai LibreOffice hoac dat bien LIBREOFFICE_BIN."
    )


def update_toc_page_numbers_with_word(docx_path, timeout=180):
    """Update only TOC page numbers via Microsoft Word, preserving TOC formatting."""
    if os.name != "nt":
        return False
    if str(os.getenv("DOCX_BUILDER_WORD_TOC_UPDATE", "1")).strip().lower() in {"0", "false", "no"}:
        return False

    docx_path = Path(docx_path).resolve()
    ps_path = str(docx_path).replace("'", "''")
    ps_script = f"""
$ErrorActionPreference = 'Stop'
$path = '{ps_path}'
$word = $null
$doc = $null
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($path, $false, $false)
    foreach ($toc in $doc.TablesOfContents) {{
        $toc.UpdatePageNumbers()
    }}
    $doc.Save()
}} finally {{
    if ($doc -ne $null) {{ $doc.Close($false) | Out-Null }}
    if ($word -ne $null) {{ $word.Quit() | Out-Null }}
}}
"""
    try:
        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                ps_script,
            ],
            check=False,
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False

    return result.returncode == 0


def convert_docx_to_pdf(docx_path, output_dir=None, timeout=180):
    docx_path = Path(docx_path).resolve()
    if not docx_path.exists():
        raise PreviewConversionError(f"Khong tim thay file DOCX de preview: {docx_path}")

    output_dir = Path(output_dir or docx_path.parent)
    output_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = output_dir / f"{docx_path.stem}.pdf"
    if pdf_path.exists():
        pdf_path.unlink()

    office_bin = find_libreoffice()

    # Prefer Word's page-number-only refresh. LibreOffice index update rebuilds
    # TOC layout and can change tab leaders/font/spacing, so do not use it here.
    update_toc_page_numbers_with_word(docx_path, timeout=timeout)

    profile_dir = Path(tempfile.mkdtemp(prefix="lo_pdf_profile_"))
    try:
        cmd = [
            office_bin,
            f"-env:UserInstallation={profile_dir.as_uri()}",
            "--headless",
            "--norestore",
            "--nofirststartwizard",
            "--convert-to",
            "pdf",
            "--outdir",
            str(output_dir),
            str(docx_path),
        ]
        env = os.environ.copy()
        env.setdefault("HOME", str(output_dir))
        env.setdefault("TMPDIR", str(output_dir))
        result = subprocess.run(
            cmd,
            check=False,
            capture_output=True,
            text=True,
            timeout=timeout,
            env=env,
        )
    except subprocess.TimeoutExpired as exc:
        raise PreviewConversionError("Convert preview PDF qua thoi gian cho phep.") from exc
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)

    if result.returncode != 0:
        details = (result.stderr or result.stdout or "").strip()
        raise PreviewConversionError(f"LibreOffice convert PDF loi: {details}")
    if not pdf_path.exists():
        raise PreviewConversionError("LibreOffice chay xong nhung khong tao file PDF preview.")
    return pdf_path
