from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Iterable


def create_zip(
    files: Iterable[Path],
    output_zip_path: Path,
    *,
    base_dir: Path | None = None,
) -> Path:
    output_zip_path.parent.mkdir(parents=True, exist_ok=True)
    if output_zip_path.exists():
        output_zip_path.unlink()

    with zipfile.ZipFile(output_zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for file_path in sorted(files):
            arcname = file_path.relative_to(base_dir) if base_dir else Path(file_path.name)
            archive.write(file_path, arcname=arcname.as_posix())

    return output_zip_path
