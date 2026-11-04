"""Excel output shared by the Etsy bots: the rows that finished and the rows still left to send."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook


def read_source(path: Path) -> tuple[list, dict[int, list]]:
    """Header and every non-blank row of the uploaded Excel, keyed by Excel row number."""
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        lines = wb.active.iter_rows(values_only=True)
        headers = list(next(lines, None) or [])
        while headers and (headers[-1] is None or not str(headers[-1]).strip()):
            headers.pop()
        width = len(headers)
        body: dict[int, list] = {}
        for row_number, values in enumerate(lines, start=2):
            values = list(values or [])[:width]
            if any(value is not None and str(value).strip() for value in values):
                body[row_number] = values + [None] * (width - len(values))
    finally:
        wb.close()
    return headers, body


def finished(rows: list) -> set[int]:
    """A row is finished only when its status is OK."""
    return {row.row_number for row in rows if row.status == "OK"}


def _write(path: Path, headers: list, rows: list[list]) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "input"
    ws.append(headers)
    for row in rows:
        ws.append(row)
    try:
        wb.save(path)
    except PermissionError:
        # The file is open in Excel. Keep the list instead of losing it.
        path = path.with_name(f"{path.stem}-{datetime.now().strftime('%H%M%S')}{path.suffix}")
        wb.save(path)
    return path


def write_split(
    folder: Path,
    done_name: str,
    remain_name: str,
    headers: list,
    source: dict[int, list],
    done_numbers: set[int],
) -> tuple[Path, Path]:
    """Write today's finished rows and every other uploaded row, with the uploaded columns."""
    folder.mkdir(parents=True, exist_ok=True)
    day = datetime.now().strftime("%Y-%m-%d")
    done = [values for number, values in source.items() if number in done_numbers]
    remain = [values for number, values in source.items() if number not in done_numbers]
    done_path = _write(folder / f"{done_name}({day}).xlsx", headers, done)
    remain_path = _write(folder / f"{remain_name}({day}).xlsx", headers, remain)
    return done_path, remain_path
