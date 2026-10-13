"""Formula recalculation via headless LibreOffice.

openpyxl cannot evaluate formulas, and when it saves a workbook it drops
all cached formula results. We therefore let LibreOffice open the file,
recalculate it and write it back so the computed values can be read.
"""

import shutil
import subprocess
import tempfile
from pathlib import Path


class RecalcError(RuntimeError):
    pass


def find_soffice():
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    for path in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if Path(path).exists():
            return path
    return None


def recalculate(path, timeout=120):
    """Recalculate all formulas in the workbook at `path` (in place)."""
    soffice = find_soffice()
    if soffice is None:
        raise RecalcError(
            "LibreOffice was not found. It is needed to calculate formulas. "
            "Install it from https://www.libreoffice.org/ and try again."
        )
    path = Path(path).resolve()
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        # A private profile avoids clashes with a running LibreOffice instance.
        profile = (tmp / "profile").as_uri()
        out_dir = tmp / "out"
        cmd = [
            soffice,
            f"-env:UserInstallation={profile}",
            "--headless",
            "--calc",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(out_dir),
            str(path),
        ]
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        except subprocess.TimeoutExpired:
            raise RecalcError(f"LibreOffice did not finish within {timeout}s.")
        result = out_dir / (path.stem + ".xlsx")
        if not result.exists():
            raise RecalcError(
                "LibreOffice could not recalculate the workbook.\n"
                f"stdout: {proc.stdout.strip()}\nstderr: {proc.stderr.strip()}"
            )
        shutil.copyfile(result, path)
