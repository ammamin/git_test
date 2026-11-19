"""Remote MCP server that lets Claude read and edit one Excel workbook.

Configuration (environment variables):
  EXCEL_FILE        Path to the .xlsx/.xlsm workbook (default /project/todo.xlsx).
  MCP_SECRET        Long random string; becomes part of the URL path (required).
  MCP_HOST          Bind address (default 127.0.0.1 - put a TLS reverse proxy in front).
  MCP_PORT          Bind port (default 8000).
  MCP_ALLOWED_HOSTS Comma-separated public hostnames accepted in the Host header
                    (e.g. "excel.example.com"). Unset = no Host check.
  EXCEL_BACKUPS     How many backups to keep before each edit (default 10, 0 = off).

The connector URL to give Claude is:  https://<your-domain>/<MCP_SECRET>/mcp
"""

from __future__ import annotations

import fcntl
import os
import shutil
import sys
import tempfile
import threading
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.server.transport_security import TransportSecuritySettings
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

Scalar = str | int | float | bool | None

MAX_CELLS_PER_READ = 5000
MAX_SEARCH_HITS = 100


def _env(name: str, default: str | None = None) -> str:
    value = os.environ.get(name, default)
    if not value:
        sys.exit(f"Missing required environment variable {name}")
    return value


EXCEL_FILE = Path(_env("EXCEL_FILE", "/project/todo.xlsx")).expanduser().resolve()
SECRET = _env("MCP_SECRET")
HOST = os.environ.get("MCP_HOST", "127.0.0.1")
PORT = int(os.environ.get("MCP_PORT", "8000"))
ALLOWED_HOSTS = [h.strip() for h in os.environ.get("MCP_ALLOWED_HOSTS", "").split(",") if h.strip()]
BACKUPS = int(os.environ.get("EXCEL_BACKUPS", "10"))

if len(SECRET) < 24 or not SECRET.replace("-", "").replace("_", "").isalnum():
    sys.exit("MCP_SECRET must be at least 24 characters of letters, digits, '-' or '_'")
if not EXCEL_FILE.is_file():
    sys.exit(f"EXCEL_FILE does not exist: {EXCEL_FILE}")

mcp = MCPServer(
    "excel",
    instructions=(
        f"Tools for reading and editing the Excel workbook '{EXCEL_FILE.name}'. "
        "Call workbook_info first to see the sheets and their headers. "
        "Cell references use A1 notation. Strings starting with '=' are written as formulas."
    ),
)

_thread_lock = threading.Lock()


@contextmanager
def _locked():
    """Serialise access across threads and processes."""
    with _thread_lock, open(f"{EXCEL_FILE}.lock", "w") as lock:
        fcntl.flock(lock, fcntl.LOCK_EX)
        try:
            yield
        finally:
            fcntl.flock(lock, fcntl.LOCK_UN)


def _load(data_only: bool = False):
    return load_workbook(EXCEL_FILE, data_only=data_only, keep_vba=EXCEL_FILE.suffix.lower() == ".xlsm")


def _sheet(wb, name: str):
    if name not in wb.sheetnames:
        raise ToolError(f"No sheet named {name!r}. Sheets: {wb.sheetnames}")
    return wb[name]


def _backup() -> None:
    if BACKUPS <= 0:
        return
    folder = EXCEL_FILE.parent / f".{EXCEL_FILE.name}.backups"
    folder.mkdir(exist_ok=True)
    stamp = time.strftime("%Y%m%d-%H%M%S") + f"-{time.time_ns() % 1_000_000:06d}"
    shutil.copy2(EXCEL_FILE, folder / f"{EXCEL_FILE.stem}.{stamp}{EXCEL_FILE.suffix}")
    for old in sorted(folder.iterdir())[:-BACKUPS]:
        old.unlink()


def _save(wb) -> None:
    """Back up, then write atomically so a crash never leaves a half-written file."""
    _backup()
    fd, tmp = tempfile.mkstemp(dir=EXCEL_FILE.parent, prefix=".tmp-", suffix=EXCEL_FILE.suffix)
    os.close(fd)
    try:
        wb.save(tmp)
        shutil.copymode(EXCEL_FILE, tmp)
        os.replace(tmp, EXCEL_FILE)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)


def _jsonable(value: Any) -> Scalar:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value)


def _bounds(ws, cell_range: str | None) -> tuple[int, int, int, int]:
    if not cell_range:
        return ws.min_column, ws.min_row, ws.max_column, ws.max_row
    min_col, min_row, max_col, max_row = range_boundaries(cell_range.upper())
    return (
        min_col or ws.min_column,
        min_row or ws.min_row,
        max_col or ws.max_column,
        max_row or ws.max_row,
    )


@mcp.tool()
def workbook_info() -> dict:
    """List every sheet with its used range, size and first row (usually the headers)."""
    with _locked():
        wb = _load(data_only=True)
        sheets = []
        for ws in wb.worksheets:
            first_row = [_jsonable(c.value) for c in ws[ws.min_row]][:50]
            sheets.append(
                {
                    "name": ws.title,
                    "used_range": ws.dimensions,
                    "rows": ws.max_row,
                    "columns": ws.max_column,
                    "first_row": first_row,
                }
            )
        return {"file": EXCEL_FILE.name, "sheets": sheets}


@mcp.tool()
def read_range(sheet: str, cell_range: str | None = None, show_formulas: bool = False) -> dict:
    """Read cells from a sheet.

    Args:
        sheet: Sheet name.
        cell_range: A1-style range such as "A1:D20", "B:B" or "3:10". Omit for the whole used range.
        show_formulas: Return formula text (e.g. "=SUM(A1:A9)") instead of the last calculated value.
    """
    with _locked():
        ws = _sheet(_load(data_only=False), sheet)
        values = None if show_formulas else _sheet(_load(data_only=True), sheet)
        min_col, min_row, max_col, max_row = _bounds(ws, cell_range)
        width = max_col - min_col + 1
        total_rows = max_row - min_row + 1
        max_rows = max(1, MAX_CELLS_PER_READ // max(width, 1))
        last_row = min(max_row, min_row + max_rows - 1)

        rows = []
        for r in range(min_row, last_row + 1):
            row = []
            for c in range(min_col, max_col + 1):
                value = ws.cell(r, c).value
                if values is not None and isinstance(value, str) and value.startswith("="):
                    cached = values.cell(r, c).value
                    # No cached result (file not recalculated since last edit): show the formula.
                    value = cached if cached is not None else value
                row.append(_jsonable(value))
            rows.append(row)

        result: dict[str, Any] = {
            "sheet": sheet,
            "range": f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{last_row}",
            "columns": [get_column_letter(c) for c in range(min_col, max_col + 1)],
            "first_row_number": min_row,
            "rows": rows,
        }
        if last_row < max_row:
            result["truncated"] = (
                f"Returned {len(rows)} of {total_rows} rows. "
                f"Continue with a range starting at row {last_row + 1}."
            )
        return result


@mcp.tool()
def search(query: str, sheet: str | None = None) -> dict:
    """Find cells whose text contains `query` (case-insensitive). Searches all sheets unless one is given."""
    needle = query.lower()
    with _locked():
        wb = _load(data_only=True)
        sheets = [_sheet(wb, sheet)] if sheet else wb.worksheets
        hits = []
        for ws in sheets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is not None and needle in str(cell.value).lower():
                        hits.append({"sheet": ws.title, "cell": cell.coordinate, "value": _jsonable(cell.value)})
                        if len(hits) >= MAX_SEARCH_HITS:
                            return {"hits": hits, "truncated": True}
        return {"hits": hits, "truncated": False}


@mcp.tool()
def write_cells(sheet: str, cells: dict[str, Scalar]) -> str:
    """Set individual cells.

    Args:
        sheet: Sheet name.
        cells: Mapping of cell reference to value, e.g. {"B2": 42, "C2": "=B2*2", "D2": null}.
               null clears a cell; strings starting with "=" become formulas.
    """
    if not cells:
        raise ToolError("cells is empty")
    with _locked():
        wb = _load()
        ws = _sheet(wb, sheet)
        for ref, value in cells.items():
            try:
                coordinate_from_string(ref.upper())
            except Exception:
                raise ToolError(f"Invalid cell reference {ref!r}; use A1 notation like 'B2'") from None
            ws[ref.upper()] = value
        _save(wb)
    return f"Wrote {len(cells)} cell(s) on {sheet!r}."


@mcp.tool()
def write_range(sheet: str, start_cell: str, rows: list[list[Scalar]]) -> str:
    """Write a block of rows starting at `start_cell` (top-left corner), overwriting what is there.

    Args:
        sheet: Sheet name.
        start_cell: Top-left cell, e.g. "A2".
        rows: List of rows, each a list of values.
    """
    if not rows:
        raise ToolError("rows is empty")
    col_letter, first_row = coordinate_from_string(start_cell.upper())
    first_col = column_index_from_string(col_letter)
    with _locked():
        wb = _load()
        ws = _sheet(wb, sheet)
        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                ws.cell(first_row + r, first_col + c, value)
        _save(wb)
    return f"Wrote {len(rows)} row(s) on {sheet!r} starting at {start_cell.upper()}."


@mcp.tool()
def append_rows(sheet: str, rows: list[list[Scalar]]) -> str:
    """Add rows after the last used row of a sheet. Values go into columns A, B, C, ..."""
    if not rows:
        raise ToolError("rows is empty")
    with _locked():
        wb = _load()
        ws = _sheet(wb, sheet)
        for row in rows:
            ws.append(row)
        last = ws.max_row
        _save(wb)
    return f"Appended {len(rows)} row(s) to {sheet!r}; they end at row {last}."


@mcp.tool()
def insert_rows(sheet: str, before_row: int, count: int = 1) -> str:
    """Insert `count` empty rows above row number `before_row`. Formulas elsewhere are NOT re-pointed."""
    with _locked():
        wb = _load()
        _sheet(wb, sheet).insert_rows(before_row, count)
        _save(wb)
    return f"Inserted {count} row(s) above row {before_row} on {sheet!r}."


@mcp.tool()
def delete_rows(sheet: str, first_row: int, count: int = 1) -> str:
    """Delete `count` rows starting at row number `first_row`. Formulas elsewhere are NOT re-pointed."""
    with _locked():
        wb = _load()
        _sheet(wb, sheet).delete_rows(first_row, count)
        _save(wb)
    return f"Deleted {count} row(s) starting at row {first_row} on {sheet!r}."


def _headers(ws, header_row: int) -> dict[str, int]:
    """Map lower-cased header text to column index."""
    headers = {}
    for cell in ws[header_row]:
        if cell.value is not None and str(cell.value).strip():
            headers.setdefault(str(cell.value).strip().lower(), cell.column)
    if not headers:
        raise ToolError(f"Row {header_row} of {ws.title!r} has no headers")
    return headers


def _columns_for(ws, header_row: int, values: dict[str, Scalar]) -> dict[int, Scalar]:
    headers = _headers(ws, header_row)
    out = {}
    for name, value in values.items():
        col = headers.get(name.strip().lower())
        if col is None:
            names = [str(c.value) for c in ws[header_row] if c.value is not None]
            raise ToolError(f"No column named {name!r}. Columns: {names}")
        out[col] = value
    return out


@mcp.tool()
def list_rows(
    sheet: str,
    where: dict[str, Scalar] | None = None,
    contains: str | None = None,
    header_row: int = 1,
    limit: int = 200,
) -> dict:
    """List table rows (e.g. todo items) as records keyed by column header.

    Args:
        sheet: Sheet name.
        where: Only rows where these columns equal these values (text compared
               case-insensitively), e.g. {"Status": "Open"}. Use null to match empty cells.
        contains: Only rows where any cell contains this text (case-insensitive).
        header_row: Row number holding the column headers.
        limit: Maximum rows to return.
    """
    def norm(v: Any) -> Any:
        return v.strip().lower() if isinstance(v, str) else _jsonable(v)

    with _locked():
        ws = _sheet(_load(data_only=True), sheet)
        headers = _headers(ws, header_row)
        names = {col: str(ws.cell(header_row, col).value).strip() for col in headers.values()}
        wanted = {col: norm(v) for col, v in _columns_for(ws, header_row, where or {}).items()}
        needle = contains.lower() if contains else None
        rows, matched = [], 0
        for r in range(header_row + 1, ws.max_row + 1):
            record = {name: _jsonable(ws.cell(r, col).value) for col, name in names.items()}
            if all(v is None for v in record.values()):
                continue
            if any(norm(ws.cell(r, col).value) != v for col, v in wanted.items()):
                continue
            if needle and not any(needle in str(v).lower() for v in record.values() if v is not None):
                continue
            matched += 1
            if len(rows) < limit:
                rows.append({"row": r, **record})
        return {"sheet": sheet, "columns": list(names.values()), "matched": matched, "rows": rows}


@mcp.tool()
def add_row(sheet: str, values: dict[str, Scalar], header_row: int = 1) -> str:
    """Add one row (e.g. a new todo) after the last non-empty row, giving values by column header.

    Args:
        sheet: Sheet name.
        values: Column header -> value, e.g. {"Task": "Buy milk", "Status": "Open"}.
        header_row: Row number holding the column headers.
    """
    if not values:
        raise ToolError("values is empty")
    with _locked():
        wb = _load()
        ws = _sheet(wb, sheet)
        cols = _columns_for(ws, header_row, values)
        row = ws.max_row
        while row > header_row and all(c.value is None for c in ws[row]):
            row -= 1
        row += 1
        for col, value in cols.items():
            ws.cell(row, col, value)
        _save(wb)
    return f"Added row {row} on {sheet!r}."


@mcp.tool()
def update_row(sheet: str, row: int, values: dict[str, Scalar], header_row: int = 1) -> str:
    """Change some columns of one existing row, e.g. mark a todo done with {"Status": "Done"}.

    Args:
        sheet: Sheet name.
        row: Row number, as returned by list_rows.
        values: Column header -> new value. Columns not listed are left alone; null clears a cell.
        header_row: Row number holding the column headers.
    """
    if not values:
        raise ToolError("values is empty")
    with _locked():
        wb = _load()
        ws = _sheet(wb, sheet)
        if not header_row < row <= ws.max_row:
            raise ToolError(f"Row {row} is outside the data (rows {header_row + 1}-{ws.max_row})")
        for col, value in _columns_for(ws, header_row, values).items():
            ws.cell(row, col, value)
        _save(wb)
    return f"Updated row {row} on {sheet!r}."


@mcp.tool()
def add_sheet(name: str) -> str:
    """Create a new empty sheet at the end of the workbook."""
    with _locked():
        wb = _load()
        if name in wb.sheetnames:
            raise ToolError(f"Sheet {name!r} already exists")
        wb.create_sheet(name)
        _save(wb)
    return f"Created sheet {name!r}."


if __name__ == "__main__":
    print(f"Serving {EXCEL_FILE} on http://{HOST}:{PORT}/<MCP_SECRET>/mcp", file=sys.stderr)
    mcp.run(
        transport="streamable-http",
        host=HOST,
        port=PORT,
        streamable_http_path=f"/{SECRET}/mcp",
        stateless_http=True,
        json_response=True,
        transport_security=TransportSecuritySettings(
            enable_dns_rebinding_protection=bool(ALLOWED_HOSTS),
            allowed_hosts=ALLOWED_HOSTS,
        ),
    )
