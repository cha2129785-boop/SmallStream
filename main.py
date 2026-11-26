# -*- coding: utf-8 -*-
"""스몰스트림 v0.3 — 스트림머지 '파일 정리 도구'만 떼어 낸 단독 프로그램 (열 삭제 · 서식 지우기 · 셀 정렬 · 중복 지우기)"""
from __future__ import annotations

import datetime as dt
import os
import subprocess
import sys
import traceback
import tkinter as tk
from tkinter import font as tkfont
from tkinter import messagebox, ttk

from appcore import APP_ID, APP_NAME, base_dir, import_from_parent, presets_path
from colprune_ui import ColPruneDialog

__version__ = "0.4.1"
AUTHOR = "Shugle(슈글)"
CREDIT = f"제작자. {AUTHOR}"


def resource_dir() -> str:
    """PyInstaller EXE → _MEIPASS (onedir: _internal), 소스 실행 → 이 파일 폴더"""
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))


def icon_path(name: str) -> str:
    return os.path.join(resource_dir(), "assets", name)


def ui_font() -> str:
    fams = set(tkfont.families())
    for f in ("맑은 고딕", "Malgun Gothic", "Noto Sans CJK KR", "NanumGothic"):
        if f in fams:
            return f
    return "TkDefaultFont"


def _report_error(root):
    def handler(exc, val, tb):
        text = "".join(traceback.format_exception(exc, val, tb))
        try:
            os.makedirs(base_dir(), exist_ok=True)
            with open(os.path.join(base_dir(), "error.log"), "a", encoding="utf-8") as f:
                f.write(f"\n[{dt.datetime.now():%Y-%m-%d %H:%M:%S}] v{__version__}\n{text}")
        except OSError:
            pass
        messagebox.showerror(APP_NAME, f"오류가 발생함 (error.log 기록)\n\n{val}", parent=root)
    return handler


def set_app_icon(root: tk.Tk) -> None:
    imgs = []
    for name in ("smallstream.png", "smallstream_32.png"):
        try:
            imgs.append(tk.PhotoImage(file=icon_path(name)))
        except tk.TclError:
            pass
    if imgs:
        root.iconphoto(True, *imgs)
        root._app_icons = imgs  # GC 방지
    if sys.platform == "win32" and os.path.isfile(icon_path("smallstream.ico")):
        try:
            root.iconbitmap(default=icon_path("smallstream.ico"))
        except tk.TclError:
            pass


def use_explicit_app_id() -> bool:
    """설치 EXE(PyInstaller)는 식별자 미설정 → Windows 가 EXE 경로로 묶어 고정 아이콘에 실행 창이 합쳐짐.
    식별자를 설정하면 바로가기(식별자 없음)와 달라 작업 표시줄 아이콘이 2개로 갈림"""
    return not getattr(sys, "frozen", False)


def setup_root() -> tk.Tk:
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # ⚠️추정: Windows 8.1+ 선명한 글자
        except Exception:  # noqa: BLE001
            pass
        if use_explicit_app_id():  # 개발용(python) 실행만: 작업 표시줄을 python 대신 이 프로그램으로 묶음
            try:
                import ctypes
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(f"Shugle.{APP_ID}")
            except Exception:  # noqa: BLE001
                pass
    root = tk.Tk()
    root.withdraw()
    root.report_callback_exception = _report_error(root)
    set_app_icon(root)
    f = ui_font()
    for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
        try:
            tkfont.nametofont(name).configure(family=f, size=10)
        except tk.TclError:
            pass
    return root


def open_folder(path: str) -> None:
    os.makedirs(path, exist_ok=True)
    if sys.platform == "win32":
        os.startfile(path)  # noqa: S606
    else:
        subprocess.Popen(["xdg-open", path])


def load_icon_image(size: int):
    """PNG 아이콘 → PhotoImage. 없으면 None"""
    p = icon_path("smallstream_128.png" if size <= 128 else "smallstream.png")
    try:
        img = tk.PhotoImage(file=p)
        if size < 128 and img.width() >= 128:
            img = img.subsample(max(1, img.width() // size))
        return img
    except tk.TclError:
        return None


class MainWindow(ColPruneDialog):
    """주 창 = 탭 2개: 파일 정리 / 정보 (v0.2). 닫으면 프로그램 종료"""

    def _make_body(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)
        tab = ttk.Frame(self.nb)
        self.nb.add(tab, text="파일 정리")
        return tab

    def build_info(self) -> ttk.Frame:
        t = ttk.Frame(self.nb, padding=16)
        self.nb.add(t, text="정보")
        self.tab_info = t
        head = ttk.Frame(t)
        head.pack(anchor="w")
        self._info_img = load_icon_image(128)
        if self._info_img is not None:
            ttk.Label(head, image=self._info_img).pack(side="left", padx=(0, 16))
        box = ttk.Frame(head)
        box.pack(side="left", anchor="n")
        ttk.Label(box, text=APP_NAME, font=(ui_font(), 20, "bold")).pack(anchor="w")
        ttk.Label(box, text=f"{APP_ID}  ·  버전 {__version__}", foreground="#555").pack(anchor="w", pady=(2, 8))
        ttk.Label(box, text=CREDIT, font=(ui_font(), 12, "bold")).pack(anchor="w")

        df = ttk.LabelFrame(t, text="데이터 폴더 (프리셋·설정·error.log)", padding=8)
        df.pack(fill="x", pady=(18, 8))
        ttk.Label(df, text=base_dir(), foreground="#1f5fa8").pack(anchor="w")
        ttk.Button(df, text="폴더 열기", command=lambda: open_folder(base_dir())).pack(anchor="w", pady=(6, 0))
        ttk.Label(df, text="첫 실행 때 스트림머지 프리셋·'같은 이름 파일 처리' 설정을 1회 복사함 (원본 유지, 이후 서로 영향 없음)",
                  foreground="#555", wraplength=640, justify="left").pack(anchor="w", pady=(4, 0))

        bf = ttk.LabelFrame(t, text="백업 폴더 (원본 이동·원본 덮어쓰기 때 원본 보관, 개인정보 포함 → 기관 PC 밖 반출 금지)", padding=8)
        bf.pack(fill="x", pady=(0, 8))
        ttk.Label(bf, text=self.backup_dir, foreground="#1f5fa8").pack(anchor="w")
        ttk.Button(bf, text="백업 폴더 열기", command=lambda: open_folder(self.backup_dir)).pack(anchor="w", pady=(6, 0))
        ttk.Label(bf, text="자동 삭제 안 함 → 필요 없는 백업은 직접 정리", foreground="#555").pack(anchor="w", pady=(4, 0))

        return t


def build_window(root: tk.Tk) -> MainWindow:
    """파일 정리 + 정보 탭 = 주 창. 닫히면 프로그램 종료"""
    dlg = MainWindow(root, presets_path(), f"{APP_NAME} v{__version__}", os.path.expanduser("~"))
    dlg.build_info()
    dlg.minsize(700, 740)
    dlg.bind("<Destroy>", lambda e: root.after_idle(root.destroy) if e.widget is dlg else None, add="+")
    return dlg


def main():
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):  # pythonw / EXE: stdout 없음
            pass
    if "--version" in sys.argv:  # 빌드 확인용
        print(f"{APP_NAME} {APP_ID} v{__version__} / data: {base_dir()}")
        return
    root = setup_root()
    try:
        got = import_from_parent()
    except OSError:
        got = []
    dlg = build_window(root)
    if got:
        messagebox.showinfo(APP_NAME, "스트림머지에서 가져옴 (원본 유지)\n\n" + "\n".join(f"• {g}" for g in got), parent=dlg)
        dlg.refresh_presets()
    root.mainloop()


if __name__ == "__main__":
    main()
