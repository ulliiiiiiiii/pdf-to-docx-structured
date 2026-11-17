"""LibreOffice without a user interface as a converter, with a fresh user profile per call so an open
LibreOffice on the same machine does not get in the way."""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")
TIMEOUT_S = 120


class OfficeError(RuntimeError):
    pass


def convert(src: Path, fmt: str, out_dir: Path, infilter: str | None = None, run=subprocess.run) -> Path:
    """Convert `src` to `fmt` (for example "docx" or "pdf") in `out_dir` and return the new file."""
    if not SOFFICE:
        raise OfficeError("LibreOffice (soffice) is not installed")
    out_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="lo-profile-") as profile:
        cmd = [SOFFICE, f"-env:UserInstallation=file://{profile}", "--headless", "--norestore"]
        if infilter:
            cmd.append(f"--infilter={infilter}")
        cmd += ["--convert-to", fmt, "--outdir", str(out_dir), str(src)]
        try:
            run(cmd, check=True, capture_output=True, timeout=TIMEOUT_S)
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as exc:
            raise OfficeError(f"{src.name} to {fmt}: {exc}") from exc
    target = out_dir / f"{src.stem}.{fmt.split(':')[0]}"
    if not target.is_file() or target.stat().st_size == 0:
        raise OfficeError(f"{src.name} to {fmt}: no output")
    return target
