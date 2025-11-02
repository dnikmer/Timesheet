"""Excel helpers for the Timesheet application."""

from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable, List, Tuple

from openpyxl import load_workbook


REFERENCE_SHEET = "Справочник"
TIMESHEET_SHEET = "Учет времени"


class ExcelStructureError(RuntimeError):
    """Raised when the workbook structure does not match expectations."""


def _normalise(values: Iterable[str | None]) -> List[str]:
    items: List[str] = []
    for value in values:
        if value is None:
            continue
        text = str(value).strip()
        if text and text not in items:
            items.append(text)
    return items


def load_reference_data(path: Path | str) -> Tuple[List[str], List[str]]:
    """Return projects and work types from the reference sheet."""

    workbook_path = Path(path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Excel file not found: {workbook_path}")

    workbook = load_workbook(workbook_path, data_only=True)

    if REFERENCE_SHEET not in workbook:
        raise ExcelStructureError(
            f"Workbook must contain sheet '{REFERENCE_SHEET}'. Found: {', '.join(workbook.sheetnames)}"
        )

    sheet = workbook[REFERENCE_SHEET]

    projects: List[str] = []
    work_types: List[str] = []
    for idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        # Skip header row if present.
        if idx == 1:
            continue
        project, work_type, *_ = row + (None, None)
        projects.append(project)
        work_types.append(work_type)

    return _normalise(projects), _normalise(work_types)


def append_time_entry(
    path: Path | str,
    *,
    project: str,
    work_type: str,
    elapsed_seconds: float,
    finished_at: datetime | None = None,
) -> None:
    """Append a new row to the timesheet sheet with the supplied data."""

    workbook_path = Path(path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Excel file not found: {workbook_path}")

    workbook = load_workbook(workbook_path)

    if TIMESHEET_SHEET not in workbook:
        raise ExcelStructureError(
            f"Workbook must contain sheet '{TIMESHEET_SHEET}'. Found: {', '.join(workbook.sheetnames)}"
        )

    sheet = workbook[TIMESHEET_SHEET]
    timestamp = finished_at or datetime.now()

    duration = timedelta(seconds=elapsed_seconds)

    sheet.append(
        [
            timestamp.date(),
            project,
            work_type,
            duration,
        ]
    )

    workbook.save(workbook_path)
