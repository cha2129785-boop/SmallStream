# -*- coding: utf-8 -*-
"""colprune v0.3 — 파일 정리: 열 삭제 · 서식 지우기(테두리 유지) · 중복 내용 지우기 → 새 파일 (원본 수정 없음)

열 지정: 열 이름(헤더 행 기준, 순서 바뀌어도 같은 열) + 고정 자리(A, C, F-H, 5)
xlsx/xlsm → 서식 유지 (병합 칸·열 너비·필터·틀 고정 위치 보정)
xls / HTML xls / csv → 표시값 + 기본 서식 xlsx
"""
from __future__ import annotations

import json
import os
import re
import datetime as dt
import unicodedata
from dataclasses import asdict, dataclass, field

from appcore import FONT_NAME, FONT_SIZE, MAX_W, MIN_W, atomic_write_json, sheet_title, text_width
from loader import LoadError, cell_str, detect_format, list_sheets, load_grid
from outpath import Resolver, backup, move_to_backup, next_free, save_wb

__version__ = "0.9"
SUFFIX = "_열삭제"          # 열 삭제만 할 때 (v0.2 이전과 같은 이름)
SUFFIX_CLEAN = "_정리"      # 서식·중복 작업 포함 (v0.3)
RESULT_SUFFIXES = (SUFFIX, SUFFIX_CLEAN)
# 중복 지우기 범위 (v0.3)
DEDUPE_MODES = {
    "col": "열마다 — 같은 열 안에서 첫 값만",
    "consecutive": "연속된 것만 — 바로 위 칸과 같을 때",
    "sheet": "시트 전체 — 처음 나온 칸만",
    "row": "행 중복 — 대상 열 값이 모두 같은 행",
}
# 중복 지우기 제외 종류 (v0.4): 칸 내용 분류
KINDS = {"text": "글자", "num": "숫자", "mixed": "혼합"}
# 셀 내용 가로 정렬 (v0.8). "" = 그대로
ALIGNS = {"": "그대로", "left": "왼쪽", "center": "가운데", "right": "오른쪽"}


def value_kind(v) -> str:
    """'num' = 숫자·기호만 (Excel 숫자·날짜 칸, 1,234 · 2026-11-10 · 010-1234-5678)
    'mixed' = 숫자 + 글자 (A101, 3층) / 'text' = 숫자 없음 (TRUE/FALSE 포함)"""
    if isinstance(v, bool):
        return "text"
    if isinstance(v, (int, float, dt.date, dt.time)):   # datetime 은 date 하위
        return "num"
    t = cell_str(v)
    if not any(ch.isdigit() for ch in t):
        return "text"
    return "mixed" if any(ch.isalpha() for ch in t) else "num"


class PruneError(Exception):
    pass


# ---------- 열 지정 ----------
def norm(s) -> str:
    """열 이름 비교용: 전각→반각, 공백 정리, 대소문자 무시"""
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", cell_str(s))).strip().casefold()


def col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(s: str) -> int:
    s = s.strip().upper()
    if s.isdigit():
        n = int(s)
    elif re.fullmatch(r"[A-Z]{1,3}", s):
        n = 0
        for ch in s:
            n = n * 26 + ord(ch) - 64
    else:
        raise PruneError(f"열 자리 형식 오류: '{s}' (예: A, C, F-H, 5)")
    if not 1 <= n <= 16384:
        raise PruneError(f"열 자리 범위 밖: '{s}'")
    return n


def parse_positions(text: str) -> list[int]:
    """'A, C, F-H, 5' → [1, 3, 5, 6, 7, 8] (중복 제거·정렬)"""
    out: set[int] = set()
    for part in re.split(r"[,\s;]+", (text or "").strip()):
        if not part:
            continue
        if "-" in part or "~" in part:
            a, b = re.split(r"[-~]", part, maxsplit=1)
            i, j = col_index(a), col_index(b)
            out.update(range(min(i, j), max(i, j) + 1))
        else:
            out.add(col_index(part))
    return sorted(out)


def parse_names(text: str) -> list[str]:
    """한 줄에 하나 (빈 줄 무시, 중복 제거, 순서 유지)"""
    seen, out = set(), []
    for line in (text or "").splitlines():
        line = line.strip()
        if line and norm(line) not in seen:
            seen.add(norm(line))
            out.append(line)
    return out


def plan_columns(header: list, names: list[str], positions: list[int]) -> tuple[list[int], list[str]]:
    """삭제할 열 번호(1부터) + 헤더에서 못 찾은 이름. 같은 이름 열이 여럿이면 모두 삭제"""
    want = {norm(n): n for n in names}
    cols, found = set(), set()
    for i, h in enumerate(header, 1):
        k = norm(h)
        if k and k in want:
            cols.add(i)
            found.add(k)
    width = len(header)
    cols.update(p for p in positions if p <= width)
    return sorted(cols), [n for k, n in want.items() if k not in found]


# ---------- 프리셋 ----------
@dataclass
class Preset:
    name: str = ""
    names: list[str] = field(default_factory=list)
    positions: str = ""          # 'A, C, F-H'
    header_row: int = 1
    all_sheets: bool = True
    formulas_to_values: bool = True
    # v0.3 서식 지우기 (테두리는 항상 유지). 옛 프리셋은 기본값 = 꺼짐
    clear_format: bool = False
    keep_numfmt: bool = True      # 표시 형식(날짜·쉼표) 유지
    clear_align: bool = True      # 정렬·줄바꿈도 지움
    unify_font: bool = True       # 맑은 고딕 11 (끄면 글꼴 지정 없음 = 파일 기본 글꼴)
    keep_size: bool = True        # 열 너비·행 높이 유지
    # v0.3 중복 내용 지우기 (헤더 행 아래만, 첫 칸만 남김)
    dedupe: bool = False
    dedupe_mode: str = "col"
    dedupe_all_cols: bool = False
    dedupe_except: bool = False   # v0.6: True → 이름·자리로 고른 열을 '빼고' 나머지 모든 열 (dedupe_all_cols 가 우선)
    dedupe_names: list[str] = field(default_factory=list)
    dedupe_positions: str = ""
    dedupe_skip: list[str] = field(default_factory=list)   # v0.4: 지우지 않을 종류 KINDS 키
    align: str = ""               # v0.8: 가로 정렬 ALIGNS 키 (값 있는 모든 칸, 세로 정렬·줄바꿈 유지)
    # v0.9 ④ 저장 설정 (화면 전용 — 엔진은 prune_file 인자로 받음). 옛 프리셋 = 지금 기본 동작
    save_where: str = "same"      # same / dir / orig
    save_dir: str = ""
    save_move: bool = False
    save_mode: str = ""           # 같은 이름 처리 outpath.MODES 키, "" = 바꾸지 않음
    name_custom: bool = False
    name_template: str = ""       # {원본} {날짜} {시각}

    def check(self) -> list[int]:
        """삭제할 고정 자리 반환. 할 작업이 없거나 설정 오류면 PruneError"""
        pos = parse_positions(self.positions)
        if self.align not in ALIGNS:
            raise PruneError(f"정렬 값 오류: {self.align}")
        if not self.names and not pos and not self.clear_format and not self.dedupe and not self.align:
            raise PruneError("할 작업을 하나 이상 고르세요 (열 삭제 · 서식 지우기 · 중복 지우기 · 정렬)")
        if self.header_row < 1:
            raise PruneError("헤더 행은 1 이상")
        if self.dedupe:
            if self.dedupe_mode not in DEDUPE_MODES:
                raise PruneError(f"중복 범위 오류: {self.dedupe_mode}")
            bad = [k for k in self.dedupe_skip if k not in KINDS]
            if bad:
                raise PruneError(f"제외 종류 오류: {', '.join(bad)}")
            if set(self.dedupe_skip) >= set(KINDS):
                raise PruneError("글자·숫자·혼합을 모두 제외하면 지울 칸이 없음 — 하나 이상 해제하세요")
            if not self.dedupe_all_cols and not self.dedupe_names and not parse_positions(self.dedupe_positions):
                raise PruneError("중복 대상: '고른 열'·'제외 열'은 열 이름·자리를 입력하세요 (또는 '모든 열')")
        return pos

    def suffix(self) -> str:
        return SUFFIX if not (self.clear_format or self.dedupe or self.align) else SUFFIX_CLEAN

    def dedupe_cols(self, header: list, width: int) -> tuple[list[int], list[str]]:
        """중복 지울 열 번호 + 못 찾은 이름"""
        if not self.dedupe:
            return [], []
        if self.dedupe_all_cols:
            return list(range(1, width + 1)), []
        cols, miss = plan_columns(list(header) + [""] * (width - len(header)), self.dedupe_names,
                                  parse_positions(self.dedupe_positions))
        if self.dedupe_except:
            if miss:   # v0.7: 제외할 열을 못 찾으면 그 열까지 지워지므로 이 시트는 중복 지우기 안 함
                return [], miss
            cols = [c for c in range(1, width + 1) if c not in set(cols)]
        return cols, miss

    def except_warning(self, sheet: str, miss: list[str]) -> str | None:
        if self.dedupe and self.dedupe_except and miss:
            return f"[{sheet}] 제외 열 못 찾음: {', '.join(miss)} → 이 시트는 중복 지우기 안 함 (헤더 행·열 이름 확인)"
        return None


def parse_name_list(text: str) -> list[str]:
    """'성명, 주소, "성명,이름"' → 이름 목록 (중복 제거, 순서 유지). v0.7: 쉼표 든 이름은 큰따옴표 (CSV 규칙)"""
    import csv
    parts: list[str] = []
    for row in csv.reader((text or "").splitlines(), skipinitialspace=True):
        parts.extend(row)
    return parse_names("\n".join(p.replace("\n", " ") for p in parts))


def join_name_list(names: list[str]) -> str:
    """parse_name_list 의 반대 (쉼표·따옴표 든 이름은 큰따옴표로 감쌈, v0.7)"""
    out = []
    for n in names:
        n = clean_name(n)
        out.append('"' + n.replace('"', '""') + '"' if ("," in n or '"' in n) else n)
    return ", ".join(out)


def clean_name(h) -> str:
    """헤더 → 입력용 이름: 줄바꿈(Alt+Enter)·연속 공백 → 공백 1칸 (비교는 norm 이 공백을 같게 봄, v0.7)"""
    return re.sub(r"\s+", " ", cell_str(h)).strip()


def dedupe_plan(rows: list[list[str]], first: int, cols: list[int], mode: str, skip=()) -> list[tuple[int, int]]:
    """지울 칸 (rows 안 행 번호 0부터, 열 번호 1부터). first = 첫 데이터 행. 빈 칸은 비교 안 함.
    비교: 앞뒤 공백만 무시 (대소문자·전각은 구분). skip: 지우지 않을 종류 (비교에는 포함, v0.4)"""
    def v(r: int, c: int) -> str:
        row = rows[r]
        return cell_str(row[c - 1]).strip() if c <= len(row) else ""

    out: list[tuple[int, int]] = []
    n = len(rows)
    if mode == "col":
        for c in cols:
            seen: set[str] = set()
            for r in range(first, n):
                x = v(r, c)
                if x:
                    if x in seen:
                        out.append((r, c))
                    else:
                        seen.add(x)
    elif mode == "consecutive":
        for c in cols:
            prev = ""
            for r in range(first, n):
                x = v(r, c)
                if x and x == prev:
                    out.append((r, c))
                prev = x
    elif mode == "sheet":
        seen = set()
        for r in range(first, n):
            for c in cols:
                x = v(r, c)
                if x:
                    if x in seen:
                        out.append((r, c))
                    else:
                        seen.add(x)
    elif mode == "row":
        seen_rows: set[tuple] = set()
        for r in range(first, n):
            key = tuple(v(r, c) for c in cols)
            if not any(key):
                continue
            if key in seen_rows:
                out.extend((r, c) for c in cols if v(r, c))
            else:
                seen_rows.add(key)
    else:
        raise PruneError(f"중복 범위 오류: {mode}")
    if skip:
        out = [(r, c) for r, c in out if value_kind(rows[r][c - 1]) not in skip]
    return sorted(out)


class PresetStore:
    def __init__(self, path: str):
        self.path = path

    def load(self) -> list[Preset]:
        try:
            with open(self.path, encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, ValueError):
            return []
        out = []
        for d in data.get("presets", []):
            try:
                d = {k: v for k, v in d.items() if k in Preset.__dataclass_fields__}
                if not isinstance(d.get("clear_align", True), bool):   # v0.18 버그로 글자 저장됨 → 기본값
                    d["clear_align"] = True
                out.append(Preset(**d))
            except TypeError:
                continue
        return out

    def save_all(self, presets: list[Preset]) -> None:
        atomic_write_json(self.path, {"format": 1, "presets": [asdict(p) for p in presets]})

    def put(self, p: Preset) -> None:
        items = [x for x in self.load() if x.name != p.name] + [p]
        self.save_all(sorted(items, key=lambda x: x.name))

    def delete(self, name: str) -> None:
        self.save_all([x for x in self.load() if x.name != name])


# ---------- 결과 ----------
@dataclass
class SheetReport:
    sheet: str
    deleted: list[str] = field(default_factory=list)   # 'C 연락처'
    missing: list[str] = field(default_factory=list)
    rows: int = 0
    cols: int = 0
    do_delete: bool = True     # v0.3: 열 삭제 작업 여부 (보고 표시용)
    cleared: int = -1          # v0.3: 중복으로 지운 칸 수 (-1 = 작업 안 함)


@dataclass
class FileReport:
    path: str
    out: str = ""
    sheets: list[SheetReport] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)    # v0.3: 서식 지우기 등 안내
    backup: str = ""                                    # v0.5: 원본 덮어쓰기 전 백업 경로
    moved: bool = False                                 # v0.6: 원본을 백업 폴더로 이동함
    error: str = ""

    @property
    def deleted_count(self) -> int:
        return sum(len(s.deleted) for s in self.sheets)


def out_path_for(path: str, out_dir: str | None, guard: list[str], suffix: str = SUFFIX) -> str:
    """새 이름 규칙 (기존 파일·원본이면 _2, _3 …)"""
    return next_free(target_path(path, out_dir, suffix), guard)


def target_path(path: str, out_dir: str | None, suffix: str, name: str = "") -> str:
    """name: render_name 결과 (확장자 없음). 비면 원본이름+suffix"""
    stem = os.path.splitext(os.path.basename(path))[0]
    return os.path.join(out_dir or os.path.dirname(os.path.abspath(path)), f"{name or stem + suffix}.xlsx")


BAD_CHARS = set('\\/:*?"<>|')
TOKENS = ("{원본}", "{날짜}", "{시각}")


def render_name(template: str, path: str, idx: int = 1, total: int = 1, now=None) -> str:
    """직접 입력 파일 이름 (v0.9). {원본}=원본 이름(확장자 빼고) {날짜}=YYYYMMDD {시각}=HHMM.
    .xlsx 등 확장자는 떼고, 여러 파일인데 {원본} 없으면 _1, _2 … 붙임. 잘못된 이름 → PruneError"""
    import datetime as _dt
    now = now or _dt.datetime.now()
    t = (template or "").strip()
    t = re.sub(r"\.(xlsx|xlsm|xls|csv)$", "", t, flags=re.I).strip()
    if not t:
        raise PruneError("파일 이름을 입력하세요")
    name = (t.replace("{원본}", os.path.splitext(os.path.basename(path))[0])
             .replace("{날짜}", f"{now:%Y%m%d}").replace("{시각}", f"{now:%H%M}"))
    bad = sorted({ch for ch in name if ch in BAD_CHARS})
    if bad:
        raise PruneError(f"파일 이름에 쓸 수 없는 글자: {' '.join(bad)}")
    if total > 1 and "{원본}" not in t:
        name += f"_{idx}"
    if name.rstrip(" .") != name:
        raise PruneError("파일 이름 끝에 공백·점은 쓸 수 없음")
    return name


# ---------- xlsx (서식 유지) ----------
def _remap(old: int, dels: list[int]) -> int | None:
    """삭제 후 열 번호 (삭제된 열이면 None)"""
    if old in dels:
        return None
    return old - sum(1 for d in dels if d < old)


def _span(a: int, b: int, dels: list[int]) -> tuple[int, int] | None:
    kept = [_remap(c, dels) for c in range(a, b + 1)]
    kept = [k for k in kept if k is not None]
    return (min(kept), max(kept)) if kept else None


def _prune_ws(ws, dels: list[int], warn: list[str]) -> None:
    from openpyxl.utils import get_column_letter, range_boundaries
    from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder

    maxc = max(ws.max_column, max(dels))
    # 열 너비·숨김 (범위로 묶인 항목 펼침)
    dims = {}
    for cd in list(ws.column_dimensions.values()):
        lo, hi = cd.min or 0, cd.max or 0
        if not lo:
            continue
        for c in range(lo, min(hi, maxc) + 1):
            dims[c] = (cd.width, cd.hidden, cd.customWidth)
    # 병합 칸: 모두 풀고 새 위치로 다시 병합
    merges = []
    for rng in list(ws.merged_cells.ranges):
        merges.append((rng.min_row, rng.min_col, rng.max_row, rng.max_col))
        ws.unmerge_cells(str(rng))
    for c in sorted(dels, reverse=True):
        ws.delete_cols(c)
    for r1, c1, r2, c2 in merges:
        sp = _span(c1, c2, dels)
        if sp and not (r1 == r2 and sp[0] == sp[1]):
            ws.merge_cells(start_row=r1, start_column=sp[0], end_row=r2, end_column=sp[1])
    # 열 너비 다시 배치
    holder = DimensionHolder(worksheet=ws)
    for c, (w, hid, cw) in sorted(dims.items()):
        n = _remap(c, dels)
        if n is None:
            continue
        L = get_column_letter(n)
        holder[L] = ColumnDimension(ws, index=L, width=w or 13, hidden=bool(hid), customWidth=bool(cw))
    ws.column_dimensions = holder
    # 자동 필터
    if ws.auto_filter.ref:
        c1, r1, c2, r2 = range_boundaries(ws.auto_filter.ref)
        sp = _span(c1, c2, dels)
        ws.auto_filter.ref = f"{get_column_letter(sp[0])}{r1}:{get_column_letter(sp[1])}{r2}" if sp else None
    # 틀 고정
    if ws.freeze_panes:
        c, r = range_boundaries(ws.freeze_panes + ":" + ws.freeze_panes)[:2]
        n = c - sum(1 for d in dels if d < c)
        ws.freeze_panes = f"{get_column_letter(max(1, n))}{r}"
    # 위치를 못 맞추는 항목은 알림
    nc = len(ws.conditional_formatting._cf_rules) if hasattr(ws.conditional_formatting, "_cf_rules") else 0
    nv = len(ws.data_validations.dataValidation) if ws.data_validations else 0
    if nc or nv:
        warn.append(f"[{ws.title}] 조건부 서식 {nc}개·데이터 유효성 {nv}개 위치가 어긋날 수 있음 — 확인 필요")  # ⚠️ openpyxl delete_cols 미보정


def clear_formats(ws, p: Preset) -> tuple[int, int, int]:
    """서식 지우기 (테두리 유지). 반환 (칸 수, 조건부 서식 수, 표 스타일 수)"""
    from openpyxl.formatting.formatting import ConditionalFormattingList
    from openpyxl.styles import Alignment, Font, PatternFill

    font = Font(name=FONT_NAME, size=FONT_SIZE) if p.unify_font else Font()
    n = 0
    for cell in list(ws._cells.values()):   # 빈 칸을 새로 만들지 않도록 기존 칸만
        if not cell.has_style and not (p.unify_font and cell.value is not None):
            continue
        cell.font = font
        cell.fill = PatternFill()
        if p.clear_align:
            cell.alignment = Alignment()
        if not p.keep_numfmt:
            cell.number_format = "General"
        n += 1                               # border·protection 은 그대로
    for d in list(ws.column_dimensions.values()) + list(ws.row_dimensions.values()):
        if d.has_style:                      # 열·행 전체 서식
            d.font = font
            d.fill = PatternFill()
            if p.clear_align:
                d.alignment = Alignment()
            if not p.keep_numfmt:
                d.number_format = "General"
    if not p.keep_size:
        for cd in ws.column_dimensions.values():
            cd.width = 0                     # width 0 → width 속성 안 씀 = Excel 기본 너비 (숨김은 유지)
        for rd in ws.row_dimensions.values():
            rd.ht = None
    ncf = len(ws.conditional_formatting)
    if ncf:
        ws.conditional_formatting = ConditionalFormattingList()
    nt = 0
    for t in ws.tables.values():             # 표(Ctrl+T) 줄무늬 색
        if t.tableStyleInfo is not None:
            t.tableStyleInfo = None
            nt += 1
    return n, ncf, nt


def apply_align(ws, h: str) -> int:
    """값 있는 칸의 가로 정렬만 바꿈 (세로·줄바꿈·들여쓰기 등 유지). 반환 칸 수 (v0.8)"""
    from openpyxl.styles import Alignment
    n = 0
    for cell in list(ws._cells.values()):
        if cell.value is None:
            continue
        a = cell.alignment
        cell.alignment = Alignment(horizontal=h, vertical=a.vertical, textRotation=a.textRotation,
                                   wrapText=a.wrapText, shrinkToFit=a.shrinkToFit,
                                   indent=a.indent if h != "center" else 0,
                                   relativeIndent=a.relativeIndent, justifyLastLine=a.justifyLastLine,
                                   readingOrder=a.readingOrder)
        n += 1
    return n


def _prune_xlsx(path: str, p: Preset, pos: list[int], out: str, rep: FileReport, skip=()) -> None:
    from openpyxl import load_workbook

    try:
        wb = load_workbook(path)
    except Exception as e:  # noqa: BLE001
        raise PruneError(f"xlsx 열기 실패: {e}") from e
    vals = None
    sheets = wb.worksheets if p.all_sheets else wb.worksheets[:1]
    formulas = 0
    fmt_cells = fmt_cf = fmt_tbl = n_align = 0
    do_delete = bool(p.names or pos)
    for ws in sheets:
        if ws.title in skip:
            continue
        width = ws.max_column
        header = [ws.cell(p.header_row, c).value for c in range(1, width + 1)]
        dels, missing = plan_columns(header, p.names, pos) if do_delete else ([], [])
        dcols, dmiss = p.dedupe_cols(header, width)
        sr = SheetReport(ws.title, missing=missing + [f"(중복) {m}" for m in dmiss], do_delete=do_delete)
        if w := p.except_warning(ws.title, dmiss):
            rep.warnings.append(w)
        if dels or dcols:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type == "f":
                        formulas += 1
                        if p.formulas_to_values:
                            if vals is None:
                                vals = load_workbook(path, data_only=True)
                            cell.value = vals[ws.title][cell.coordinate].value
        if p.dedupe:                          # 원래 열 자리 기준 → 열 삭제보다 먼저
            sr.cleared = 0
            if dcols:
                rows = [list(r) for r in ws.iter_rows(min_row=1, values_only=True)]
                for r, c in dedupe_plan(rows, p.header_row, dcols, p.dedupe_mode, p.dedupe_skip):
                    ws.cell(r + 1, c).value = None
                    sr.cleared += 1
        if dels:
            for c in dels:
                sr.deleted.append(f"{col_letter(c)} {cell_str(header[c - 1]) if c <= len(header) else ''}".rstrip())
            _prune_ws(ws, dels, rep.warnings)
        if p.clear_format:
            a, b, c = clear_formats(ws, p)
            fmt_cells, fmt_cf, fmt_tbl = fmt_cells + a, fmt_cf + b, fmt_tbl + c
        if p.align:                           # 서식 지우기 뒤 (정렬 지움 옵션보다 우선)
            n_align += apply_align(ws, p.align)
        sr.rows, sr.cols = ws.max_row, ws.max_column
        rep.sheets.append(sr)
    if p.clear_format:
        extra = "".join([f" · 조건부 서식 {fmt_cf}개" if fmt_cf else "", f" · 표 스타일 {fmt_tbl}개" if fmt_tbl else ""])
        rep.notes.append(f"서식 지움: 칸 {fmt_cells}개{extra} (테두리 유지)")
    if p.align:
        rep.notes.append(f"정렬 {ALIGNS[p.align]}: 칸 {n_align}개 (값 있는 칸, 세로 정렬·줄바꿈 유지)")
    if p.dedupe and formulas and not p.formulas_to_values:
        rep.warnings.append("수식 유지 상태 → 중복 비교는 수식 글자로 함 (결과값 기준 아님)")
    if formulas:
        rep.warnings.append(f"수식 {formulas}개 → 값으로 바꿈 (Excel에서 저장된 적 없는 파일은 빈 값일 수 있음)"
                            if p.formulas_to_values else
                            f"수식 {formulas}개 유지 — 삭제된 열을 참조하면 #REF! 가능, 확인 필요")
    if path.lower().endswith(".xlsm"):
        rep.warnings.append("매크로는 새 파일(.xlsx)에 포함되지 않음")
    try:
        save_wb(wb, out)
    except PermissionError as e:
        raise PruneError(f"저장 실패 — 파일이 열려 있거나 권한 없음: {out}") from e


# ---------- xls / HTML / csv (표시값) ----------
def _prune_grid(path: str, p: Preset, pos: list[int], out: str, rep: FileReport) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter

    names = list_sheets(path)
    if not p.all_sheets:
        names = names[:1]
    wb = Workbook()
    wb.remove(wb.active)
    used: set[str] = set()
    base, bold = Font(name=FONT_NAME, size=FONT_SIZE), Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
    for nm in names:
        sname, grid, _ = load_grid(path, nm or None, display_format=True)
        width = max((len(r) for r in grid), default=0)
        header = grid[p.header_row - 1] + [""] * width if len(grid) >= p.header_row else []
        header = header[:width]
        do_delete = bool(p.names or pos)
        dels, missing = plan_columns(header, p.names, pos) if do_delete else ([], [])
        dcols, dmiss = p.dedupe_cols(header, width)
        sr = SheetReport(sname or "시트", missing=missing + [f"(중복) {m}" for m in dmiss], do_delete=do_delete,
                         deleted=[f"{col_letter(c)} {cell_str(header[c - 1])}".rstrip() for c in dels])
        if w := p.except_warning(sr.sheet, dmiss):
            rep.warnings.append(w)
        if p.dedupe:
            plan = dedupe_plan(grid, p.header_row, dcols, p.dedupe_mode, p.dedupe_skip) if dcols else []
            for r, c in plan:
                grid[r][c - 1] = ""
            sr.cleared = len(plan)
        keep = [c for c in range(1, width + 1) if c not in dels]
        ws = wb.create_sheet(sheet_title(sname or os.path.splitext(os.path.basename(path))[0], used))
        widths = [0] * len(keep)
        for ri, row in enumerate(grid, 1):
            for ni, c in enumerate(keep):
                v = row[c - 1] if c <= len(row) else ""
                cell = ws.cell(ri, ni + 1, v if v != "" else None)
                cell.font = bold if ri == p.header_row and not p.clear_format else base
                if p.align and v != "":
                    cell.alignment = Alignment(horizontal=p.align)
                widths[ni] = max(widths[ni], max((text_width(x) for x in v.split("\n")), default=0))
        for ni, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(ni)].width = min(MAX_W, max(MIN_W, w + 2))
        sr.rows, sr.cols = len(grid), len(keep)
        rep.sheets.append(sr)
    rep.warnings.append("xls·csv 형식 → 화면 표시값을 글자로 저장 (원본 서식·수식 없음)")
    if p.clear_format:
        rep.notes.append("서식 지움: 원본 서식은 원래 옮겨지지 않음 → 헤더 굵게도 뺌")
    if p.align:
        rep.notes.append(f"정렬 {ALIGNS[p.align]}: 값 있는 칸")
    try:
        save_wb(wb, out)
    except PermissionError as e:
        raise PruneError(f"저장 실패 — 파일이 열려 있거나 권한 없음: {out}") from e


# ---------- 공개 API ----------
def preview(path: str, p: Preset) -> FileReport:
    """저장 없이 지울 열·못 찾은 이름 확인"""
    rep = FileReport(path)
    try:
        pos = p.check()
        do_delete = bool(p.names or pos)
        disp = detect_format(path) != "xlsx"   # xlsx 실행은 칸 값으로 비교 → 미리보기도 값 기준
        for nm in (list_sheets(path) if p.all_sheets else list_sheets(path)[:1]):
            sname, grid, _ = load_grid(path, nm or None, display_format=disp)
            width = max((len(r) for r in grid), default=0)
            header = (grid[p.header_row - 1] if len(grid) >= p.header_row else []) + [""] * width
            header = header[:width]
            dels, missing = plan_columns(header, p.names, pos) if do_delete else ([], [])
            dcols, dmiss = p.dedupe_cols(header, width)
            sr = SheetReport(sname or "시트", [f"{col_letter(c)} {cell_str(header[c - 1])}".rstrip() for c in dels],
                             missing + [f"(중복) {m}" for m in dmiss], len(grid), width - len(dels), do_delete)
            if w := p.except_warning(sr.sheet, dmiss):
                rep.warnings.append(w)
            if p.dedupe:
                sr.cleared = len(dedupe_plan(grid, p.header_row, dcols, p.dedupe_mode, p.dedupe_skip)) if dcols else 0
            rep.sheets.append(sr)
        if p.clear_format:
            rep.notes.append("서식 지우기: 실행 시 적용 (테두리 유지)")
        if p.align:
            rep.notes.append(f"정렬 {ALIGNS[p.align]}: 실행 시 적용")
    except (LoadError, PruneError) as e:
        rep.error = str(e)
    return rep


def prune_file(path: str, p: Preset, out_dir: str | None = None, guard: list[str] | None = None,
               skip_sheets=(), resolver: Resolver | None = None, to_original: bool = False,
               backup_dir: str = "", move_original: bool = False, name: str = "") -> FileReport:
    """skip_sheets: 손대지 않을 시트 이름 (xlsx만, v0.2 — 한 번에 실행의 결과 파일용)
    resolver: 같은 이름 결과 파일 처리 (기본 새 이름, v0.5)
    to_original: 원본에 덮어쓰기 (v0.5) — .xlsx 는 backup_dir 에 백업 후 교체, xls·csv·xlsm 은 원본이름.xlsx
    move_original: out_dir 에 원본이름.xlsx 로 저장 + 원본은 backup_dir 로 이동 (v0.6). 같은 이름 있으면 _1, _2
    name: 결과 파일 이름 (확장자 없음, render_name 결과, v0.9). 원본 덮어쓰기에는 안 씀"""
    rep = FileReport(path)
    try:
        pos = p.check()
        if not os.path.isfile(path):
            raise PruneError(f"파일 없음: {path}")
        others = [g for g in (guard or []) if os.path.normcase(os.path.abspath(g)) != os.path.normcase(os.path.abspath(path))]
        if move_original:
            return _prune_move(path, p, pos, out_dir, others, skip_sheets, backup_dir, rep, name)
        if to_original:
            stem, ext = os.path.splitext(path)
            target = path if ext.lower() == ".xlsx" else stem + ".xlsx"
            if os.path.normcase(os.path.abspath(target)) == os.path.normcase(os.path.abspath(path)):
                if not backup_dir:
                    raise PruneError("백업 폴더가 지정되지 않아 원본에 덮어쓸 수 없음")
                try:
                    rep.backup = backup(path, backup_dir)
                except OSError as e:
                    raise PruneError(f"백업 실패 → 원본 덮어쓰기 중단: {e}") from e
                out = path
            else:
                out = (resolver or Resolver("rename")).resolve(target, [path] + others)
                rep.notes.append(f"{ext.lower() or '이 형식'} 원본은 그대로 두고 같은 이름 .xlsx 로 저장")
        else:
            out = (resolver or Resolver("rename")).resolve(target_path(path, out_dir, p.suffix(), name), [path] + others)
        if out is None:
            raise PruneError("저장 취소")
        if detect_format(path) == "xlsx":
            _prune_xlsx(path, p, pos, out, rep, skip_sheets)
        else:
            _prune_grid(path, p, pos, out, rep)
        rep.out = out
    except (LoadError, PruneError) as e:
        rep.error = str(e)
    return rep


def _prune_move(path, p, pos, out_dir, others, skip, backup_dir, rep: FileReport, name: str = "") -> FileReport:
    """정리 결과를 임시 파일로 만든 뒤 → 원본을 백업 폴더로 이동 → 지정 폴더의 원본이름.xlsx 로 교체.
    처리·이동이 실패하면 원본은 제자리 (v0.6)"""
    import tempfile
    if not out_dir or not os.path.isdir(out_dir):
        raise PruneError("저장 폴더를 지정하세요")
    if not backup_dir:
        raise PruneError("백업 폴더가 지정되지 않아 원본을 옮길 수 없음")
    stem = name or os.path.splitext(os.path.basename(path))[0]
    out = next_free(os.path.join(out_dir, stem + ".xlsx"), guard=others, start=1, free=[path])
    fd, tmp = tempfile.mkstemp(prefix="~sm_", suffix=".xlsx", dir=out_dir)
    os.close(fd)
    try:
        if detect_format(path) == "xlsx":
            _prune_xlsx(path, p, pos, tmp, rep, skip)
        else:
            _prune_grid(path, p, pos, tmp, rep)
        try:
            rep.backup = move_to_backup(path, backup_dir)
        except OSError as e:
            raise PruneError(f"원본 이동 실패 (파일이 열려 있을 수 있음) → 저장 중단: {e}") from e
        rep.moved = True
        try:
            os.replace(tmp, out)
        except OSError as e:
            raise PruneError(f"저장 실패 — 결과는 임시 파일 {tmp} 에 있음, 원본은 {rep.backup} 로 이동됨: {e}") from e
        rep.out = out
    finally:
        if os.path.exists(tmp) and not (rep.moved and not rep.out):   # 이동 후 교체 실패 시 결과 임시 파일 보존
            os.remove(tmp)
    return rep


def prune_files(paths: list[str], p: Preset, out_dir: str | None = None, resolver: Resolver | None = None,
                to_original: bool = False, backup_dir: str = "", move_original: bool = False,
                name_template: str = "") -> list[FileReport]:
    """name_template: 직접 입력 파일 이름 (v0.9). 원본 덮어쓰기면 무시"""
    import datetime as _dt
    now = _dt.datetime.now()   # 여러 파일이 같은 {날짜}{시각}
    out = []
    for i, x in enumerate(paths, 1):
        name = ""
        if name_template and not to_original:
            try:
                name = render_name(name_template, x, i, len(paths), now)
            except PruneError as e:
                out.append(FileReport(x, error=str(e)))
                continue
        out.append(prune_file(x, p, out_dir, guard=paths, resolver=resolver, to_original=to_original,
                              backup_dir=backup_dir, move_original=move_original, name=name))
    return out
