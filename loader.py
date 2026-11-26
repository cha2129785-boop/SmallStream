# -*- coding: utf-8 -*-
"""loader v0.1 — xlsx/xlsm/xls/csv/HTML 위장 xls → Table (모든 값 str)"""
from __future__ import annotations

import csv
import datetime as dt
import io
import os
import re
from dataclasses import dataclass
from html.parser import HTMLParser

from numfmt import format_value

__version__ = "0.7"


class LoadError(Exception):
    pass


@dataclass
class Table:
    path: str
    sheet: str
    headers: list[str]
    rows: list[list[str]]
    row_numbers: list[int]  # 원본 행 번호 (1부터)
    fmt: str  # xlsx / xls / csv / html

    @property
    def label(self) -> str:
        base = os.path.basename(self.path)
        return f"{base}[{self.sheet}]" if self.sheet else base

    def column(self, name: str) -> list[str]:
        i = self.headers.index(name)
        return [r[i] for r in self.rows]


# ---------- 공통 ----------
def cell_str(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, dt.datetime):
        if v.time() == dt.time(0):
            return v.strftime("%Y-%m-%d")
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, dt.time):
        return v.strftime("%H:%M:%S")
    if isinstance(v, float):
        if v.is_integer() and abs(v) < 1e15:
            return str(int(v))
        return repr(v)
    return str(v).replace("\xa0", " ").strip()


def make_headers(raw: list[str]) -> list[str]:
    out: list[str] = []
    used: set[str] = set()
    for i, h in enumerate(raw, 1):
        h = re.sub(r"\s+", " ", h or "").strip() or f"열{i}"
        name, n = h, 1
        while name in used:
            n += 1
            name = f"{h}_{n}"
        used.add(name)
        out.append(name)
    return out


def _last_nonempty(row: list[str]) -> int:
    for i in range(len(row) - 1, -1, -1):
        if row[i] != "":
            return i + 1
    return 0


def grid_to_table(grid: list[list[str]], header_row: int, path: str, sheet: str, fmt: str) -> Table:
    if header_row < 1 or header_row > len(grid):
        raise LoadError(f"헤더 행 {header_row} 없음 (전체 {len(grid)}행)")
    hdr = grid[header_row - 1]
    data = grid[header_row:]
    width = max([_last_nonempty(hdr)] + [_last_nonempty(r) for r in data])
    if width == 0:
        raise LoadError("빈 시트")
    headers = make_headers([(hdr[i] if i < len(hdr) else "") for i in range(width)])
    rows, nums = [], []
    for idx, r in enumerate(data):
        r = [(r[i] if i < len(r) else "") for i in range(width)]
        if any(c != "" for c in r):
            rows.append(r)
            nums.append(header_row + 1 + idx)
    return Table(path, sheet, headers, rows, nums, fmt)


def detect_format(path: str) -> str:
    with open(path, "rb") as f:
        head = f.read(4096)
    if head.startswith(b"PK\x03\x04"):
        return "xlsx"
    if head.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return "xls"
    low = head.lstrip(b"\xef\xbb\xbf").lstrip().lower()
    if low.startswith(b"<") or b"<table" in low:
        return "html"
    ext = os.path.splitext(path)[1].lower()
    if ext in (".csv", ".tsv", ".txt"):
        return "csv"
    raise LoadError(f"지원하지 않는 형식: {os.path.basename(path)}")


# ---------- xlsx ----------
def _xlsx_grids(path: str):
    import openpyxl
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:  # noqa: BLE001
        raise LoadError(f"xlsx 읽기 실패: {e}") from e
    return wb


def _xlsx_sheets(path: str) -> list[str]:
    wb = _xlsx_grids(path)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _xlsx_grid(path: str, sheet: str | None, display: bool = True) -> tuple[str, list[list[str]]]:
    wb = _xlsx_grids(path)
    try:
        if sheet is None:
            ws = wb.worksheets[0]
        elif sheet in wb.sheetnames:
            ws = wb[sheet]
        else:
            raise LoadError(f"시트 없음: {sheet}")
        if not display:
            grid = [[cell_str(v) for v in row] for row in ws.iter_rows(values_only=True)]
            return ws.title, grid
        d1904 = getattr(getattr(wb, "epoch", None), "year", 1899) == 1904
        grid = []
        for row in ws.iter_rows():
            out = []
            for c in row:
                v = getattr(c, "value", None)
                shown = format_value(v, getattr(c, "number_format", None), d1904)
                out.append(cell_str(v) if shown is None else shown)
            grid.append(out)
        return ws.title, grid
    finally:
        wb.close()


# ---------- xls (BIFF) ----------
def _xls_book(path: str, formatting: bool = False):
    import xlrd
    try:
        return xlrd.open_workbook(path, on_demand=True, formatting_info=formatting)
    except Exception as e:  # noqa: BLE001
        raise LoadError(f"xls 읽기 실패 (암호 설정 파일일 수 있음): {e}") from e


def _xls_sheets(path: str) -> list[str]:
    book = _xls_book(path)
    try:
        return book.sheet_names()
    finally:
        book.release_resources()


def _xls_fmt(book, cell) -> str | None:
    try:
        xf = book.xf_list[cell.xf_index]
        return book.format_map[xf.format_key].format_str
    except (AttributeError, IndexError, KeyError, TypeError):
        return None


def _xls_grid(path: str, sheet: str | None, display: bool = True) -> tuple[str, list[list[str]]]:
    import xlrd
    try:
        book = _xls_book(path, formatting=display)
    except LoadError:
        if not display:
            raise
        book = _xls_book(path)  # 서식 정보 읽기 실패 → 값만
        display = False
    try:
        names = book.sheet_names()
        if sheet is None:
            sh = book.sheet_by_index(0)
        elif sheet in names:
            sh = book.sheet_by_name(sheet)
        else:
            raise LoadError(f"시트 없음: {sheet}")
        grid = []
        for r in range(sh.nrows):
            row = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                t = cell.ctype
                if t == xlrd.XL_CELL_DATE:
                    try:
                        row.append(cell_str(xlrd.xldate_as_datetime(cell.value, book.datemode)))
                    except Exception:  # noqa: BLE001
                        row.append(cell_str(cell.value))
                elif t == xlrd.XL_CELL_BOOLEAN:
                    row.append(cell_str(bool(cell.value)))
                elif t == xlrd.XL_CELL_ERROR:
                    row.append(xlrd.error_text_from_code.get(cell.value, "#ERR"))
                elif display and t in (xlrd.XL_CELL_NUMBER, xlrd.XL_CELL_TEXT):
                    shown = format_value(cell.value, _xls_fmt(book, cell), book.datemode == 1)
                    row.append(cell_str(cell.value) if shown is None else shown)
                else:
                    row.append(cell_str(cell.value))
            grid.append(row)
        return sh.name, grid
    finally:
        book.release_resources()


# ---------- csv ----------
def _decode(raw: bytes, prefer: str | None = None) -> str:
    encs = [e for e in (prefer, "utf-8-sig", "cp949") if e]
    for enc in encs:
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    raise LoadError("인코딩 판별 실패 (UTF-8 / CP949 아님)")


def _csv_grid(path: str) -> list[list[str]]:
    with open(path, "rb") as f:
        text = _decode(f.read())
    if path.lower().endswith(".tsv"):
        delim = "\t"
    else:
        try:
            delim = csv.Sniffer().sniff(text[:8192], delimiters=",\t;|").delimiter
        except csv.Error:
            delim = ","
    return [[c.replace("\xa0", " ").strip() for c in row] for row in csv.reader(io.StringIO(text), delimiter=delim)]


# ---------- HTML 위장 xls ----------
class _TableParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables: list[list[list[str]]] = []
        self._depth = 0
        self._rows = None
        self._row = None
        self._cell = None
        self._span = (1, 1)
        self._carry: dict[int, int] = {}  # 열 → 남은 rowspan

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "table":
            self._depth += 1
            if self._depth == 1:
                self._rows, self._carry = [], {}
        elif self._depth != 1:
            return
        elif tag == "tr":
            self._row = []
        elif tag in ("td", "th") and self._row is not None:
            self._cell = []
            self._span = (_int(a.get("colspan")), _int(a.get("rowspan")))
        elif tag == "br" and self._cell is not None:
            self._cell.append(" ")

    def handle_endtag(self, tag):
        if tag == "table":
            if self._depth == 1 and self._rows is not None:
                self.tables.append(self._rows)
                self._rows = None
            self._depth = max(0, self._depth - 1)
        elif self._depth != 1:
            return
        elif tag in ("td", "th") and self._cell is not None:
            self._place(re.sub(r"\s+", " ", "".join(self._cell)).strip())
            self._cell = None
        elif tag == "tr" and self._row is not None:
            self._fill_carry(len(self._row) + 10_000, until_end=True)
            self._rows.append(self._row)
            self._carry = {c: n - 1 for c, n in self._carry.items() if n - 1 > 0}
            self._row = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)

    def _fill_carry(self, limit, until_end=False):
        # rowspan으로 점유된 열은 빈 칸으로 채움
        while len(self._row) in self._carry and len(self._row) < limit:
            self._row.append("")
        if until_end:
            for c in sorted(self._carry):
                if c >= len(self._row):
                    self._row.extend([""] * (c - len(self._row) + 1))

    def _place(self, text):
        self._fill_carry(10**9)
        col = len(self._row)
        cs, rs = self._span
        self._row.append(text)
        self._row.extend([""] * (cs - 1))
        if rs > 1:
            for c in range(col, col + cs):
                self._carry[c] = rs  # 현재 행 포함 → tr 종료 시 1 감소


def _int(v) -> int:
    try:
        return max(1, int(v))
    except (TypeError, ValueError):
        return 1


def _html_tables(path: str) -> list[list[list[str]]]:
    with open(path, "rb") as f:
        raw = f.read()
    m = re.search(rb"charset\s*=\s*[\"']?([\w-]+)", raw[:4096], re.I)
    prefer = m.group(1).decode("ascii").lower() if m else None
    if prefer in ("euc-kr", "ks_c_5601-1987", "ksc5601"):
        prefer = "cp949"
    p = _TableParser()
    p.feed(_decode(raw, prefer))
    p.close()
    tables = [t for t in p.tables if t]
    if not tables:
        raise LoadError("HTML 안에 표 없음")
    return tables


# ---------- 공개 API ----------
def list_sheets(path: str) -> list[str]:
    fmt = detect_format(path)
    if fmt == "xlsx":
        return _xlsx_sheets(path)
    if fmt == "xls":
        return _xls_sheets(path)
    if fmt == "html":
        return [f"표{i}" for i in range(1, len(_html_tables(path)) + 1)]
    return [""]


def load_grid(path: str, sheet: str | None = None, display_format: bool = True) -> tuple[str, list[list[str]], str]:
    """시트 전체 칸 (헤더 처리 없음) → (시트 이름, 칸, 형식). v0.7: 열 삭제 도구용"""
    if not os.path.isfile(path):
        raise LoadError(f"파일 없음: {path}")
    fmt = detect_format(path)
    if fmt == "xlsx":
        name, grid = _xlsx_grid(path, sheet, display_format)
    elif fmt == "xls":
        name, grid = _xls_grid(path, sheet, display_format)
    elif fmt == "html":
        tables = _html_tables(path)
        idx = 0
        if sheet:
            m = re.fullmatch(r"표(\d+)", sheet)
            if not m or not 1 <= int(m.group(1)) <= len(tables):
                raise LoadError(f"표 없음: {sheet}")
            idx = int(m.group(1)) - 1
        name, grid = f"표{idx + 1}", tables[idx]
    else:
        name, grid = "", _csv_grid(path)
    return name, grid, fmt


def load_table(path: str, sheet: str | None = None, header_row: int = 1, display_format: bool = True) -> Table:
    """display_format=True: Excel 셀 표시 형식대로 읽음 (v0.6, 예: 46333 + "입원기간" 형식 → '입원기간')"""
    if not os.path.isfile(path):
        raise LoadError(f"파일 없음: {path}")
    fmt = detect_format(path)
    if fmt == "xlsx":
        name, grid = _xlsx_grid(path, sheet, display_format)
    elif fmt == "xls":
        name, grid = _xls_grid(path, sheet, display_format)
    elif fmt == "html":
        tables = _html_tables(path)
        idx = 0
        if sheet:
            m = re.fullmatch(r"표(\d+)", sheet)
            if not m or not 1 <= int(m.group(1)) <= len(tables):
                raise LoadError(f"표 없음: {sheet}")
            idx = int(m.group(1)) - 1
        name, grid = f"표{idx + 1}", tables[idx]
    else:
        name, grid = "", _csv_grid(path)
    return grid_to_table(grid, header_row, path, name, fmt)
