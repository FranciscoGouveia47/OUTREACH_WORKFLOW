from __future__ import annotations

from pathlib import Path
from uuid import uuid4

from openpyxl import load_workbook


def save_workbook_atomically(workbook, workbook_path: Path) -> None:
    workbook_path = Path(workbook_path)
    workbook_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = workbook_path.with_name(
        f".{workbook_path.stem}.{uuid4().hex}.tmp{workbook_path.suffix}"
    )
    try:
        workbook.save(temp_path)
        check = load_workbook(temp_path, read_only=True, data_only=True)
        check.close()
        workbook.close()
        temp_path.replace(workbook_path)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise
