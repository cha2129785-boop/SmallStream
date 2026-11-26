# -*- coding: utf-8 -*-
"""outpath v0.3 — 결과 파일 저장 경로: 같은 이름이 있을 때 덮어쓰기 / 새 이름(_2) / 매번 묻기, 백업, 안전 저장"""
from __future__ import annotations

import datetime as dt
import os
import shutil
import tempfile

__version__ = "0.3"
MODES = {"ask": "매번 묻기", "overwrite": "덮어쓰기", "rename": "새 이름 (_2, _3 …)"}
DEFAULT_MODE = "ask"
SETTING_KEY = "overwrite_mode"


def _n(p: str) -> str:
    return os.path.normcase(os.path.abspath(p))


# ---------- 설정 (settings.json) ----------
def get_mode() -> str:
    from appcore import load_settings
    try:
        m = load_settings().get(SETTING_KEY, DEFAULT_MODE)
    except Exception:  # noqa: BLE001 — 설정 파일 문제로 저장이 막히지 않게
        return DEFAULT_MODE
    return m if m in MODES else DEFAULT_MODE


def set_mode(mode: str) -> None:
    from appcore import load_settings, save_settings
    if mode not in MODES:
        raise ValueError(mode)
    s = load_settings()
    s[SETTING_KEY] = mode
    save_settings(s)


# ---------- 경로 ----------
def next_free(path: str, guard=(), start: int = 2, free=()) -> str:
    """path 가 없고 guard(원본)도 아니면 그대로, 아니면 _2, _3 … (start=1 → _1, _2 …, v0.2)
    free: 곧 비워질 경로 (있어도 쓸 수 있음, 예: 백업으로 옮길 원본)"""
    g, f = {_n(x) for x in guard}, {_n(x) for x in free}
    base, ext = os.path.splitext(path)
    cand, n = path, start - 1

    def busy(c):
        return (os.path.exists(c) and _n(c) not in f) or (_n(c) in g and _n(c) not in f)
    while busy(cand):
        n += 1
        cand = f"{base}_{n}{ext}"
    return cand


class Resolver:
    """같은 이름 파일 처리. ask(path, many) → (선택 'overwrite'|'rename'|None=취소, 나머지도 같게 bool)"""

    def __init__(self, mode: str = "rename", ask=None, many: bool = False):
        self.mode = mode if mode in MODES else "rename"
        self.ask = ask
        self.many = many
        self.fixed: str | None = None

    def resolve(self, path: str, guard=()) -> str | None:
        """저장할 경로 / None = 사용자 취소. guard(원본) 경로는 덮어쓰지 않고 새 이름"""
        if _n(path) in {_n(x) for x in guard}:
            return next_free(path, guard)
        if not os.path.exists(path):
            return path
        choice = self.fixed or (None if self.mode == "ask" else self.mode)
        if choice is None:
            if self.ask is None:
                choice = "rename"
            else:
                choice, remember = self.ask(path, self.many)
                if choice is None:
                    return None
                if remember:
                    self.fixed = choice
        return path if choice == "overwrite" else next_free(path, guard)


# ---------- 저장 ----------
def save_wb(wb, out: str) -> None:
    """같은 폴더 임시 파일에 저장 → 교체 (덮어쓰기 중 실패해도 기존 파일 유지). PermissionError 그대로 전달"""
    d = os.path.dirname(os.path.abspath(out))
    fd, tmp = tempfile.mkstemp(prefix="~sm_", suffix=".xlsx", dir=d)
    os.close(fd)
    try:
        wb.save(tmp)
        os.replace(tmp, out)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


def backup(path: str, backup_dir: str, now: dt.datetime | None = None) -> str:
    """원본 복사 → backup_dir/원본이름_YYYYMMDD_HHMMSS.확장자 (자동 삭제 안 함). 반환 백업 경로"""
    os.makedirs(backup_dir, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(path))
    dst = next_free(os.path.join(backup_dir, f"{stem}_{(now or dt.datetime.now()):%Y%m%d_%H%M%S}{ext}"))
    shutil.copy2(path, dst)
    return dst


def is_locked(path: str) -> bool:
    """다른 프로그램(Excel 등)이 쓰기 잠금 중인지 (v0.3). 쓰기 모드로 열었다 바로 닫음 — 내용·수정 시각 안 바뀜.
    ⚠️ Windows 공유 위반(PermissionError) 기준. Linux 는 잠금이 없어 항상 False"""
    if not os.path.isfile(path):
        return False
    try:
        with open(path, "r+b"):
            pass
    except PermissionError:
        return True
    except OSError:
        return False
    return False


def move_to_backup(path: str, backup_dir: str, now: dt.datetime | None = None) -> str:
    """원본을 backup_dir 로 이동 (이름_YYYYMMDD_HHMMSS.확장자, v0.2). 반환 이동한 경로"""
    os.makedirs(backup_dir, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(path))
    dst = next_free(os.path.join(backup_dir, f"{stem}_{(now or dt.datetime.now()):%Y%m%d_%H%M%S}{ext}"))
    shutil.move(path, dst)
    return dst


# ---------- 묻기 창 (Tk) ----------
def ask_overwrite(parent, path: str, many: bool = False) -> tuple[str | None, bool]:
    """[덮어쓰기] / [새 이름] / [취소]. many=True → '나머지 파일도 같게' 체크"""
    import tkinter as tk
    from tkinter import ttk

    res: dict = {"c": None}
    w = tk.Toplevel(parent)
    w.title("같은 이름 파일 있음")
    w.resizable(False, False)
    f = ttk.Frame(w, padding=14)
    f.pack(fill="both", expand=True)
    ttk.Label(f, text="같은 이름의 파일이 이미 있음. 덮어쓸까요?", font=("", 10, "bold")).pack(anchor="w")
    ttk.Label(f, text=path, foreground="#1f5fa8", wraplength=460, justify="left").pack(anchor="w", pady=(6, 4))
    ttk.Label(f, text="[아니오] → 번호를 붙인 새 이름(_2)으로 저장", foreground="#555").pack(anchor="w")
    v_all = tk.BooleanVar(value=False)
    if many:
        ttk.Checkbutton(f, text="나머지 파일도 같게", variable=v_all).pack(anchor="w", pady=(6, 0))
    b = ttk.Frame(f)
    b.pack(anchor="e", pady=(12, 0))

    def done(c):
        res["c"] = c
        w.destroy()
    ttk.Button(b, text="예 (덮어쓰기)", command=lambda: done("overwrite")).pack(side="left")
    ttk.Button(b, text="아니오 (새 이름)", command=lambda: done("rename")).pack(side="left", padx=6)
    ttk.Button(b, text="취소", command=lambda: done(None)).pack(side="left")
    w.bind("<Escape>", lambda e: done(None))
    w.protocol("WM_DELETE_WINDOW", lambda: done(None))
    w.transient(parent)
    w.grab_set()
    parent.wait_window(w)
    return res["c"], v_all.get()


def tk_resolver(parent, many: bool = False, mode: str | None = None) -> Resolver:
    """화면용: 설정 기본값 + 묻기 창. 테스트는 모듈의 ask_overwrite 를 바꿔 끼움"""
    import outpath as me
    return Resolver(mode or get_mode(), lambda p, m: me.ask_overwrite(parent, p, m), many)
