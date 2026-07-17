"""
core/libreoffice.py - Detect LibreOffice (portable hoặc system) + chuyển xlsx → pdf.
"""
from __future__ import annotations
import os
import platform
import shutil
import subprocess
from pathlib import Path


def find_libreoffice() -> str | None:
    """Tìm executable của LibreOffice.

    Ưu tiên:
    1. ENV var ULNL_SOFFICE_PATH
    2. PATH lookup: soffice / libreoffice
    3. Vị trí cài đặt mặc định theo OS
    4. Portable trong thư mục hiện tại / parent / Desktop / Downloads
    """
    env = os.environ.get("ULNL_SOFFICE_PATH")
    if env and Path(env).exists():
        return env

    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p

    sys = platform.system()
    candidates = []
    if sys == "Windows":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        # Tìm trong các thư mục portable phổ biến
        for base in (
            os.path.expanduser("~/Desktop"),
            os.path.expanduser("~/Downloads"),
            os.path.expanduser("~/Documents"),
            "C:/",
            "D:/",
        ):
            if os.path.isdir(base):
                try:
                    for name in os.listdir(base):
                        n = name.lower()
                        if "libreoffice" in n or "lo_portable" in n or "loportable" in n:
                            for sub in ("program/soffice.exe", "App/libreoffice/program/soffice.exe",
                                        "soffice.exe"):
                                p = os.path.join(base, name, sub)
                                if os.path.isfile(p):
                                    candidates.append(p)
                except (PermissionError, OSError):
                    pass
    elif sys == "Darwin":
        candidates = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            os.path.expanduser("~/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        ]
    else:
        candidates = [
            "/usr/bin/soffice", "/usr/bin/libreoffice",
            "/usr/local/bin/soffice", "/snap/bin/libreoffice",
            "/opt/libreoffice/program/soffice",
        ]

    for c in candidates:
        if c and Path(c).exists():
            return c
    return None


def xlsx_to_pdf(xlsx_path: str | Path, out_dir: str | Path,
                soffice_path: str | None = None, timeout: int = 180) -> Path | None:
    """Convert xlsx → pdf bằng LibreOffice headless. Trả về Path của pdf hoặc None."""
    xlsx_path = Path(xlsx_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    soffice = soffice_path or find_libreoffice()
    if not soffice:
        return None

    cmd = [
        soffice, "--headless", "--norestore", "--nologo", "--nolockcheck",
        "--convert-to", "pdf", "--outdir", str(out_dir),
        str(xlsx_path),
    ]
    try:
        # Chỉ định profile riêng để không xung đột với LO đang mở
        import tempfile
        profile_dir = tempfile.mkdtemp(prefix="lo_profile_")
        cmd.insert(1, f"-env:UserInstallation=file://{profile_dir}")
        subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        # Cleanup profile
        try:
            import shutil as _sh
            _sh.rmtree(profile_dir, ignore_errors=True)
        except Exception:
            pass
    except (subprocess.TimeoutExpired, FileNotFoundError):
        return None

    pdf_path = out_dir / (xlsx_path.stem + ".pdf")
    return pdf_path if pdf_path.exists() else None


def open_with_default(path: str | Path) -> bool:
    """Mở file bằng ứng dụng mặc định của OS."""
    path = str(path)
    sys = platform.system()
    try:
        if sys == "Windows":
            os.startfile(path)
        elif sys == "Darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
        return True
    except Exception:
        return False
