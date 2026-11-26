# -*- coding: utf-8 -*-
"""colprune_ui v0.10 — 파일 정리 도구 창 (여러 파일 → 열 삭제·서식 지우기·중복 지우기 → 새 파일)"""
from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

import colprune as cp
import outpath
from loader import LoadError, cell_str, list_sheets, load_grid

TITLE = "파일 정리 도구"
FILETYPES = [("Excel·CSV", "*.xlsx *.xlsm *.xls *.csv"), ("모든 파일", "*.*")]


class ColPruneDialog(tk.Toplevel):
    def __init__(self, master, presets_path: str, app_name: str = "스몰스트림", initial_dir: str = ""):
        super().__init__(master)
        self.title(f"{TITLE} — {app_name}")
        self.app_name = app_name
        self.store = cp.PresetStore(presets_path)
        self.backup_dir = os.path.join(os.path.dirname(os.path.abspath(presets_path)), "backup")  # v0.4
        self.files: list[str] = []
        self.last_dir = initial_dir if initial_dir and os.path.isdir(initial_dir) else os.path.expanduser("~")
        self.minsize(700, 700)
        self.reports: list[cp.FileReport] = []
        self.body = self._make_body()
        self._build()
        self.refresh_presets()
        self.bind("<Escape>", lambda e: self.destroy())

    # ---------- 화면 ----------
    def _make_body(self):
        """위젯을 놓을 컨테이너 (v0.10). 기본 = 창 자체, 스몰스트림은 Notebook 의 탭"""
        return self

    def _build(self):
        pad = {"padx": 10, "pady": 2}
        # 파일
        f1 = ttk.LabelFrame(self.body, text="① 파일 (원본은 그대로, 새 파일로 저장)", padding=6)
        f1.pack(fill="x", **pad)
        self.lb = tk.Listbox(f1, height=3, selectmode="extended")
        self.lb.pack(side="left", fill="both", expand=True)
        b = ttk.Frame(f1)
        b.pack(side="left", fill="y", padx=(6, 0))
        ttk.Button(b, text="파일 추가…", command=self.add_files).pack(fill="x")
        ttk.Button(b, text="선택 제거", command=self.remove_files).pack(fill="x", pady=4)
        # 프리셋
        f2 = ttk.LabelFrame(self.body, text="② 프리셋 · 삭제할 열 (비워 두면 열 삭제 안 함)", padding=6)
        f2.pack(fill="both", expand=True, **pad)
        row = ttk.Frame(f2)
        row.pack(fill="x")
        ttk.Label(row, text="프리셋").pack(side="left")
        self.v_preset = tk.StringVar()
        self.cb = ttk.Combobox(row, textvariable=self.v_preset, state="readonly", width=28)
        self.cb.pack(side="left", padx=6)
        self.cb.bind("<<ComboboxSelected>>", lambda e: self.load_preset())
        ttk.Button(row, text="저장…", command=self.save_preset).pack(side="left")
        ttk.Button(row, text="삭제", command=self.delete_preset).pack(side="left", padx=4)

        grid = ttk.Frame(f2)
        grid.pack(fill="both", expand=True, pady=(6, 0))
        grid.columnconfigure(0, weight=1)
        grid.rowconfigure(1, weight=1)
        hdr = ttk.Frame(grid)
        hdr.grid(row=0, column=0, sticky="ew")
        ttk.Label(hdr, text="열 이름 (한 줄에 하나 · 순서 바뀌어도 같은 열 · 띄어쓰기·대소문자 무시)").pack(side="left")
        ttk.Button(hdr, text="첫 파일에서 고르기…", command=self.pick_names).pack(side="right")
        self.txt = tk.Text(grid, height=3, width=40, undo=True)
        self.txt.grid(row=1, column=0, sticky="nsew", pady=2)

        opt = ttk.Frame(grid)
        opt.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        ttk.Label(opt, text="고정 자리").grid(row=0, column=0, sticky="w")
        self.v_pos = tk.StringVar()
        ttk.Entry(opt, textvariable=self.v_pos, width=18).grid(row=0, column=1, sticky="w", padx=6)
        ttk.Label(opt, text="예: A, C, F-H, 5", foreground="gray").grid(row=0, column=2, sticky="w")
        hr = ttk.Frame(opt)   # v0.9: 헤더 행을 같은 줄로 (창 높이 절약)
        hr.grid(row=0, column=3, sticky="w", padx=(16, 0))
        ttk.Label(hr, text="헤더 행").pack(side="left")
        self.v_hdr = tk.IntVar(value=1)
        ttk.Spinbox(hr, from_=1, to=50, textvariable=self.v_hdr, width=5).pack(side="left", padx=6)
        ttk.Label(hr, text="열 이름을 찾을 행", foreground="gray").pack(side="left")
        self.v_all = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="모든 시트에 적용 (끄면 첫 시트만)", variable=self.v_all).grid(row=2, column=0, columnspan=3, sticky="w")
        self.v_fv = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="수식을 값으로 바꾸기 (열 삭제로 수식이 깨지는 것 방지, 권장)",
                        variable=self.v_fv).grid(row=3, column=0, columnspan=3, sticky="w")
        self._build_clean(pad)
        # 저장 위치
        f3 = ttk.LabelFrame(self.body, text="④ 저장 — 원본이름_열삭제.xlsx (서식·중복 포함 시 _정리.xlsx)", padding=6)
        f3.pack(fill="x", **pad)
        r1 = ttk.Frame(f3)
        r1.pack(fill="x")
        self.v_where = tk.StringVar(value="same")   # same / dir / orig (v0.4)
        self.v_where.trace_add("write", lambda *a: self._sync_where())
        ttk.Radiobutton(r1, text="원본과 같은 폴더", variable=self.v_where, value="same").pack(side="left")
        ttk.Radiobutton(r1, text="폴더 지정:", variable=self.v_where, value="dir").pack(side="left", padx=(12, 0))
        self.v_dir = tk.StringVar()
        ttk.Entry(r1, textvariable=self.v_dir, width=26).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(r1, text="찾기…", command=self.pick_dir).pack(side="left")
        r1b = ttk.Frame(f3)
        r1b.pack(fill="x", padx=(20, 0))
        self.v_move = tk.BooleanVar(value=False)   # v0.5
        self.cb_move = ttk.Checkbutton(r1b, text="지정 폴더에 원본 이름(직접 입력 시 그 이름)으로 저장 + 원본은 백업 폴더로 이동 (같은 이름 _1, _2)",
                                       variable=self.v_move)
        self.cb_move.pack(side="left")
        r2 = ttk.Frame(f3)
        r2.pack(fill="x", pady=(4, 0))
        ttk.Radiobutton(r2, text="원본에 덮어쓰기 (백업 후)", variable=self.v_where, value="orig").pack(side="left")
        ttk.Label(r2, text="같은 이름 결과 파일이 있을 때").pack(side="left", padx=(18, 4))
        self.v_mode = tk.StringVar(value=outpath.MODES[outpath.get_mode()])
        cb = ttk.Combobox(r2, textvariable=self.v_mode, state="readonly", width=16, values=list(outpath.MODES.values()))
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda e: self._save_mode())
        r3 = ttk.Frame(f3)   # v0.9 파일 이름
        r3.pack(fill="x", pady=(4, 0))
        ttk.Label(r3, text="파일 이름:").pack(side="left")
        self.v_namec = tk.BooleanVar(value=False)
        ttk.Radiobutton(r3, text="자동", variable=self.v_namec, value=False,
                        command=self._sync_where).pack(side="left", padx=(8, 0))
        self.rb_namec = ttk.Radiobutton(r3, text="직접 입력:", variable=self.v_namec, value=True, command=self._sync_where)
        self.rb_namec.pack(side="left", padx=(8, 0))
        self.v_name = tk.StringVar()
        self.e_name = ttk.Entry(r3, textvariable=self.v_name, width=24)
        self.e_name.pack(side="left", padx=4)
        ttk.Label(r3, text="{원본} {날짜} {시각} 사용 가능 · .xlsx 자동", foreground="gray").pack(side="left", padx=4)
        self._sync_where()
        # 실행
        bar = ttk.Frame(self.body, padding=10)
        bar.pack(fill="x")
        ttk.Button(bar, text="닫기", command=self.destroy).pack(side="right")
        ttk.Button(bar, text="실행", command=self.run).pack(side="right", padx=4)
        ttk.Button(bar, text="미리보기", command=self.preview).pack(side="right")
        self.status = ttk.Label(bar, text="")
        self.status.pack(side="left")

    def _build_clean(self, pad):
        """③ 서식·중복 (v0.2)"""
        f = ttk.LabelFrame(self.body, text="③ 서식 · 중복 정리 (헤더 행 위·헤더 행은 중복 지우기 안 함)", padding=6)
        f.pack(fill="x", **pad)
        self.v_fmt = tk.BooleanVar(value=False)
        self.v_numfmt = tk.BooleanVar(value=True)
        self.v_align = tk.BooleanVar(value=True)
        self.v_font = tk.BooleanVar(value=True)
        self.v_size = tk.BooleanVar(value=True)
        ttk.Checkbutton(f, text="서식 모두 지우기 — 배경색·글자색·굵게·기울임·밑줄·취소선·글꼴·조건부 서식 (테두리 유지)",
                        variable=self.v_fmt, command=self._sync_clean).grid(row=0, column=0, columnspan=4, sticky="w")
        self.fmt_opts = []
        for i, (t, v) in enumerate([("표시 형식 유지 (날짜·쉼표)", self.v_numfmt), ("정렬·줄바꿈도 지움", self.v_align),
                                    ("글꼴 맑은 고딕 11", self.v_font), ("열 너비·행 높이 유지", self.v_size)]):
            cb = ttk.Checkbutton(f, text=t, variable=v)
            cb.grid(row=1, column=i, sticky="w", padx=(20 if i == 0 else 8, 0))
            self.fmt_opts.append(cb)
        ra = ttk.Frame(f)   # v0.8 가로 정렬 (서식 지우기와 별개)
        ra.grid(row=2, column=0, columnspan=4, sticky="w", pady=(6, 0))
        ttk.Label(ra, text="셀 내용 정렬:").pack(side="left")
        self.v_halign = tk.StringVar(value="")   # v0.9: v_align(정렬·줄바꿈 지움 체크)과 이름 분리 — v0.18 버그
        for k, t in cp.ALIGNS.items():
            ttk.Radiobutton(ra, text=t, variable=self.v_halign, value=k).pack(side="left", padx=(8, 0))
        ttk.Label(ra, text="값 있는 모든 칸 · 세로 정렬·줄바꿈 유지", foreground="gray").pack(side="left", padx=10)
        self.v_dd = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="중복 내용 지우기 (처음 칸만 남김)", variable=self.v_dd,
                        command=self._sync_clean).grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))
        self.v_ddmode = tk.StringVar(value=cp.DEDUPE_MODES["col"])
        self.cb_ddmode = ttk.Combobox(f, textvariable=self.v_ddmode, state="readonly", width=34,
                                      values=list(cp.DEDUPE_MODES.values()))
        self.cb_ddmode.grid(row=3, column=2, columnspan=2, sticky="w", pady=(6, 0))
        r3 = ttk.Frame(f)
        r3.grid(row=4, column=0, columnspan=4, sticky="ew", padx=(20, 0), pady=(2, 0))
        self.v_ddtgt = tk.StringVar(value="pick")   # all / pick / except (v0.5)
        self.dd_widgets = [self.cb_ddmode]
        for t, val in (("모든 열", "all"), ("고른 열", "pick"), ("제외 열:", "except")):
            rb = ttk.Radiobutton(r3, text=t, variable=self.v_ddtgt, value=val, command=self._sync_clean)
            rb.pack(side="left", padx=(0, 6))
            self.dd_widgets.append(rb)
        self.v_ddnames = tk.StringVar()
        self.e_ddnames = ttk.Entry(r3, textvariable=self.v_ddnames, width=22)
        self.e_ddnames.pack(side="left")
        self.b_ddpick = ttk.Button(r3, text="고르기…", command=self.pick_dd_names)
        self.b_ddpick.pack(side="left", padx=4)
        ttk.Label(r3, text="자리").pack(side="left", padx=(6, 2))
        self.v_ddpos = tk.StringVar()
        self.e_ddpos = ttk.Entry(r3, textvariable=self.v_ddpos, width=8)
        self.e_ddpos.pack(side="left")
        ttk.Label(r3, text="쉼표 구분", foreground="gray").pack(side="left", padx=6)
        r4 = ttk.Frame(f)
        r4.grid(row=5, column=0, columnspan=4, sticky="ew", padx=(20, 0), pady=(2, 0))
        ttk.Label(r4, text="제외 (지우지 않음):").pack(side="left")
        self.v_skip = {k: tk.BooleanVar(value=False) for k in cp.KINDS}
        for k, t in cp.KINDS.items():
            cb = ttk.Checkbutton(r4, text=t, variable=self.v_skip[k])
            cb.pack(side="left", padx=(8, 0))
            self.dd_widgets.append(cb)
        ttk.Label(r4, text="숫자: 1,234·날짜·전화번호 / 혼합: A101·3층 / 글자: 숫자 없음",
                  foreground="gray").pack(side="left", padx=10)
        self._sync_clean()

    def _sync_clean(self):
        st = "normal" if self.v_fmt.get() else "disabled"
        for w in self.fmt_opts:
            w.state(["!disabled"] if st == "normal" else ["disabled"])
        on = self.v_dd.get()
        for w in self.dd_widgets:
            w.state(["!disabled"] if on else ["disabled"])
        if on:
            self.cb_ddmode.state(["readonly"])
        pick = on and self.v_ddtgt.get() != "all"
        for w in (self.e_ddnames, self.b_ddpick, self.e_ddpos):
            w.state(["!disabled"] if pick else ["disabled"])

    def _ddmode_key(self) -> str:
        return next((k for k, t in cp.DEDUPE_MODES.items() if t == self.v_ddmode.get()), "col")

    def _first_header(self) -> list[str] | None:
        if not self.files:
            messagebox.showinfo(TITLE, "파일을 먼저 추가하세요", parent=self)
            return None
        try:
            hdr = max(1, int(self.v_hdr.get()))
            sheet = (list_sheets(self.files[0]) or [""])[0]
            _, grid, _ = load_grid(self.files[0], sheet or None)
        except (LoadError, tk.TclError, ValueError) as e:
            messagebox.showerror(TITLE, f"읽기 실패: {e}", parent=self)
            return None
        header = [cell_str(h) for h in (grid[hdr - 1] if len(grid) >= hdr else [])]
        if not any(header):
            messagebox.showinfo(TITLE, f"{hdr}행에 열 이름 없음 → 헤더 행 번호 확인", parent=self)
            return None
        return header

    def pick_dd_names(self):
        header = self._first_header()
        if header is None:
            return
        have = {cp.norm(n) for n in cp.parse_name_list(self.v_ddnames.get())}
        title = "중복 지우기에서 뺄 열 고르기" if self.v_ddtgt.get() == "except" else "중복을 지울 열 고르기"
        sel = NamePickDialog(self, header, have, title).result
        if sel is not None:
            self.v_ddnames.set(cp.join_name_list(sel))

    def _sync_where(self):
        if hasattr(self, "cb_move"):
            self.cb_move.state(["!disabled"] if self.v_where.get() == "dir" else ["disabled"])
        if hasattr(self, "e_name"):   # 원본 덮어쓰기는 원본 이름 그대로
            orig = self.v_where.get() == "orig"
            self.rb_namec.state(["disabled"] if orig else ["!disabled"])
            self.e_name.state(["!disabled"] if self.v_namec.get() and not orig else ["disabled"])

    def _mode_key(self) -> str:
        return next((k for k, t in outpath.MODES.items() if t == self.v_mode.get()), "ask")

    def _save_mode(self):
        """같은 이름 처리 = 전체 설정 (정보 탭과 같은 값)"""
        try:
            outpath.set_mode(self._mode_key())
        except OSError:
            pass

    # ---------- 파일 ----------
    def add_files(self):
        paths = filedialog.askopenfilenames(parent=self, title="정리할 파일 (여러 개 가능)",
                                            initialdir=self.last_dir, filetypes=FILETYPES)
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.lb.insert("end", os.path.basename(p))
        if paths:
            self.last_dir = os.path.dirname(paths[0])
        self._status()

    def remove_files(self):
        for i in sorted(self.lb.curselection(), reverse=True):
            self.lb.delete(i)
            del self.files[i]
        self._status()

    def _status(self, msg: str = ""):
        self.status.config(text=msg or f"파일 {len(self.files)}개")

    def pick_dir(self):
        d = filedialog.askdirectory(parent=self, title="저장할 폴더", initialdir=self.v_dir.get() or self.last_dir)
        if d:
            self.v_dir.set(d)
            self.v_where.set("dir")

    # ---------- 프리셋 ----------
    def current(self) -> cp.Preset:
        try:
            hdr = int(self.v_hdr.get())
        except (tk.TclError, ValueError):
            hdr = 0
        return cp.Preset(self.v_preset.get(), cp.parse_names(self.txt.get("1.0", "end")), self.v_pos.get().strip(),
                         hdr, self.v_all.get(), self.v_fv.get(),
                         clear_format=self.v_fmt.get(), keep_numfmt=self.v_numfmt.get(), clear_align=self.v_align.get(),
                         unify_font=self.v_font.get(), keep_size=self.v_size.get(), align=self.v_halign.get(),
                         dedupe=self.v_dd.get(), dedupe_mode=self._ddmode_key(), dedupe_all_cols=self.v_ddtgt.get() == "all",
                         dedupe_except=self.v_ddtgt.get() == "except",
                         dedupe_names=cp.parse_name_list(self.v_ddnames.get()), dedupe_positions=self.v_ddpos.get().strip(),
                         dedupe_skip=[k for k, v in self.v_skip.items() if v.get()],
                         save_where=self.v_where.get(), save_dir=self.v_dir.get().strip(), save_move=self.v_move.get(),
                         save_mode=self._mode_key(), name_custom=self.v_namec.get(), name_template=self.v_name.get().strip())

    def refresh_presets(self, select: str = ""):
        names = [p.name for p in self.store.load()]
        self.cb["values"] = names
        if select in names:
            self.v_preset.set(select)

    def load_preset(self):
        p = next((x for x in self.store.load() if x.name == self.v_preset.get()), None)
        if not p:
            return
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", "\n".join(p.names))
        self.v_pos.set(p.positions)
        self.v_hdr.set(p.header_row)
        self.v_all.set(p.all_sheets)
        self.v_fv.set(p.formulas_to_values)
        self.v_fmt.set(p.clear_format)
        self.v_numfmt.set(p.keep_numfmt)
        self.v_align.set(p.clear_align)
        self.v_font.set(p.unify_font)
        self.v_size.set(p.keep_size)
        self.v_halign.set(p.align if p.align in cp.ALIGNS else "")
        self.v_dd.set(p.dedupe)
        self.v_ddmode.set(cp.DEDUPE_MODES.get(p.dedupe_mode, cp.DEDUPE_MODES["col"]))
        self.v_ddtgt.set("all" if p.dedupe_all_cols else "except" if p.dedupe_except else "pick")
        self.v_ddnames.set(cp.join_name_list(p.dedupe_names))
        self.v_ddpos.set(p.dedupe_positions)
        for k, v in self.v_skip.items():
            v.set(k in p.dedupe_skip)
        self._sync_clean()
        # v0.9 ④ 저장 설정
        self.v_where.set(p.save_where if p.save_where in ("same", "dir", "orig") else "same")
        self.v_dir.set(p.save_dir)
        self.v_move.set(p.save_move)
        if p.save_mode in outpath.MODES:
            self.v_mode.set(outpath.MODES[p.save_mode])
            self._save_mode()
        self.v_namec.set(p.name_custom)
        self.v_name.set(p.name_template)
        self._sync_where()
        if p.save_where == "dir" and p.save_dir and not os.path.isdir(p.save_dir):
            self._status(f"⚠ 프리셋의 저장 폴더 없음: {p.save_dir}")

    def save_preset(self):
        p = self.current()
        try:
            p.check()
        except cp.PruneError as e:
            messagebox.showwarning(TITLE, str(e), parent=self)
            return
        name = simpledialog.askstring(TITLE, "프리셋 이름", initialvalue=self.v_preset.get(), parent=self)
        if not name or not name.strip():
            return
        p.name = name.strip()
        if p.name in self.cb["values"] and not messagebox.askyesno(TITLE, f"'{p.name}' 프리셋을 덮어쓸까요?", parent=self):
            return
        self.store.put(p)
        self.refresh_presets(p.name)
        self._status(f"프리셋 저장: {p.name}")

    def delete_preset(self):
        n = self.v_preset.get()
        if n and messagebox.askyesno(TITLE, f"'{n}' 프리셋을 삭제할까요?", parent=self):
            self.store.delete(n)
            self.v_preset.set("")
            self.refresh_presets()

    def pick_names(self):
        header = self._first_header()
        if header is None:
            return
        have = {cp.norm(n) for n in cp.parse_names(self.txt.get("1.0", "end"))}
        sel = NamePickDialog(self, header, have).result
        if sel is not None:
            self.txt.delete("1.0", "end")
            self.txt.insert("1.0", "\n".join(cp.clean_name(x) for x in sel))

    # ---------- 실행 ----------
    def _ready(self) -> cp.Preset | None:
        if not self.files:
            messagebox.showinfo(TITLE, "파일을 추가하세요", parent=self)
            return None
        p = self.current()
        try:
            p.check()
        except cp.PruneError as e:
            messagebox.showwarning(TITLE, str(e), parent=self)
            return None
        return p

    def preview(self):
        p = self._ready()
        if not p:
            return
        reps = [cp.preview(f, p) for f in self.files]
        if self.v_namec.get() and self.v_where.get() != "orig":   # v0.9 파일 이름 미리 보기
            for i, r in enumerate(reps, 1):
                try:
                    r.notes.append(f"저장 이름: {cp.render_name(self.v_name.get(), r.path, i, len(reps))}.xlsx")
                except cp.PruneError as e:
                    r.warnings.append(f"파일 이름: {e}")
        ResultDialog(self, "미리보기 (저장 안 함)", reps, saved=False)

    def run(self):
        p = self._ready()
        if not p:
            return
        out_dir, where = None, self.v_where.get()
        if where == "dir":
            out_dir = self.v_dir.get().strip()
            if not out_dir or not os.path.isdir(out_dir):
                messagebox.showwarning(TITLE, "저장 폴더를 지정하세요", parent=self)
                return
        orig = where == "orig"
        move = where == "dir" and self.v_move.get()
        tmpl = self.v_name.get().strip() if self.v_namec.get() and not orig else ""
        if self.v_namec.get() and not orig:
            try:
                ex = cp.render_name(tmpl, self.files[0], 1, len(self.files))
            except cp.PruneError as e:
                messagebox.showwarning(TITLE, str(e), parent=self)
                return
            self._status(f"파일 이름 예: {ex}.xlsx")
        if move:
            msg = (f"원본 {len(self.files)}개를 백업 폴더로 옮기고, 정리 결과를 지정 폴더에 원본 이름(.xlsx)으로 저장합니다.\n\n"
                   f"백업 폴더: {self.backup_dir}\n저장 폴더: {out_dir}\n"
                   "⚠ 원본 자리에는 파일이 없어짐 (저장 폴더가 원본 폴더면 같은 자리에 정리본)")
            if not messagebox.askyesno(TITLE, msg + "\n\n계속할까요?", parent=self):
                return
        if orig:
            n_x = sum(1 for f in self.files if f.lower().endswith(".xlsx"))
            msg = (f"원본 xlsx {n_x}개를 정리 결과로 덮어씁니다.\n\n"
                   f"덮어쓰기 전 백업: {self.backup_dir}\n"
                   "⚠ 원본의 그림·차트·피벗 표는 결과에 남지 않을 수 있음 (백업에는 남음)")
            if n_x < len(self.files):
                msg += f"\nxls·csv·xlsm {len(self.files) - n_x}개는 원본을 두고 같은 이름 .xlsx 로 저장"
            if not messagebox.askyesno(TITLE, msg + "\n\n계속할까요?", parent=self):
                return
        files, skipped = list(self.files), []
        if move or orig:   # v0.6: 옮기거나 덮어쓸 원본이 Excel에서 열려 있는지 미리 확인
            check = files if move else [f for f in files if f.lower().endswith(".xlsx")]
            while True:
                locked = [f for f in check if outpath.is_locked(f)]
                if not locked:
                    break
                c = ask_locked(self, locked)
                if c == "retry":
                    continue
                if c != "skip":
                    self._status("실행 취소 (열린 파일)")
                    return
                skipped = locked
                files = [f for f in files if f not in locked]
                break
        res = outpath.tk_resolver(self, many=len(files) > 1, mode=self._mode_key())
        self.config(cursor="watch")
        self.update_idletasks()
        try:
            done = {r.path: r for r in cp.prune_files(files, p, out_dir, res, orig, self.backup_dir, move, tmpl)}
        finally:
            self.config(cursor="")
        for f in skipped:
            done[f] = cp.FileReport(f, error="건너뜀 — Excel 등에서 열려 있음 (원본·결과 변경 없음)")
        self.reports = [done[f] for f in self.files if f in done]
        ok = sum(1 for r in self.reports if r.out)
        self._status(f"완료 {ok}개 / 실패 {len(self.reports) - ok}개")
        ResultDialog(self, "결과", self.reports, saved=True)


def ask_locked(parent, paths: list[str]) -> str | None:
    """열린 파일 안내 (v0.6). 반환 'retry' / 'skip' / None(취소)"""
    res: dict = {"c": None}
    w = tk.Toplevel(parent)
    w.title(f"{TITLE} — 열려 있는 파일")
    w.resizable(False, False)
    f = ttk.Frame(w, padding=14)
    f.pack(fill="both", expand=True)
    ttk.Label(f, text=f"다른 프로그램(Excel 등)에서 열려 있어 옮기거나 덮어쓸 수 없는 파일 {len(paths)}개",
              font=("", 10, "bold")).pack(anchor="w")
    ttk.Label(f, text="\n".join(f"  · {os.path.basename(x)}" for x in paths[:15]) + ("\n  …" if len(paths) > 15 else ""),
              justify="left").pack(anchor="w", pady=(6, 4))
    ttk.Label(f, text="Excel에서 닫은 뒤 [다시 시도]. [건너뛰기]는 이 파일들만 빼고 진행", foreground="#555").pack(anchor="w")
    b = ttk.Frame(f)
    b.pack(anchor="e", pady=(12, 0))

    def done(c):
        res["c"] = c
        w.destroy()
    ttk.Button(b, text="다시 시도", command=lambda: done("retry")).pack(side="left")
    ttk.Button(b, text="건너뛰기", command=lambda: done("skip")).pack(side="left", padx=6)
    ttk.Button(b, text="취소", command=lambda: done(None)).pack(side="left")
    w.bind("<Escape>", lambda e: done(None))
    w.protocol("WM_DELETE_WINDOW", lambda: done(None))
    w.transient(parent)
    w.grab_set()
    parent.wait_window(w)
    return res["c"]


def report_text(reps: list[cp.FileReport], saved: bool) -> str:
    lines = []
    for r in reps:
        lines.append(f"■ {os.path.basename(r.path)}")
        if r.error:
            lines.append(f"   [실패] {r.error}")
            continue
        for s in r.sheets:
            lines.append(f"   [{s.sheet}]")
            if s.do_delete:
                lines.append(f"      삭제 {len(s.deleted)}열: {', '.join(s.deleted) or '없음'}")
            if s.cleared >= 0:
                lines.append(f"      중복 지움 {s.cleared}칸")
            if s.missing:
                lines.append(f"      못 찾은 이름: {', '.join(s.missing)}")
            lines.append(f"      {'저장' if saved else '저장 시'} {s.rows}행 × {s.cols}열")
        for n in r.notes:
            lines.append(f"   · {n}")
        for w in r.warnings:
            lines.append(f"   ⚠ {w}")
        if r.backup:
            lines.append(f"   {'원본 이동' if r.moved else '백업'}: {r.backup}")
        if r.out:
            lines.append(f"   → {r.out}{'' if r.moved or not r.backup else ' (원본 덮어씀)'}")
    return "\n".join(lines)


class ResultDialog(tk.Toplevel):
    def __init__(self, master, title: str, reps: list[cp.FileReport], saved: bool):
        super().__init__(master)
        self.title(f"{TITLE} — {title}")
        self.transient(master)
        self.geometry("720x420")
        txt = tk.Text(self, wrap="word")
        sb = ttk.Scrollbar(self, command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        bar = ttk.Frame(self, padding=8)
        bar.pack(side="bottom", fill="x")
        sb.pack(side="right", fill="y")
        txt.pack(fill="both", expand=True)
        txt.insert("1.0", report_text(reps, saved))
        txt.configure(state="disabled")
        self.text = txt
        ttk.Button(bar, text="닫기", command=self.destroy).pack(side="right")
        outs = [r.out for r in reps if r.out]
        if outs and hasattr(os, "startfile"):
            ttk.Button(bar, text="저장 폴더 열기", command=lambda: os.startfile(os.path.dirname(outs[0]))).pack(side="right", padx=4)  # Windows 전용


class NamePickDialog(tk.Toplevel):
    """헤더 목록에서 지울 열 체크. result = 이름 목록 / None"""

    def __init__(self, master, header: list[str], checked: set[str], title: str = "지울 열 고르기"):
        super().__init__(master)
        self.title(title)
        self.transient(master)
        self.result = None
        outer = ttk.Frame(self, padding=8)
        outer.pack(fill="both", expand=True)
        cv = tk.Canvas(outer, height=320, width=320, highlightthickness=0)
        sb = ttk.Scrollbar(outer, command=cv.yview)
        inner = ttk.Frame(cv)
        inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.create_window((0, 0), window=inner, anchor="nw")
        cv.configure(yscrollcommand=sb.set)
        cv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.vars: list[tuple[str, tk.BooleanVar]] = []
        for i, h in enumerate(header, 1):
            if not h.strip():
                continue
            v = tk.BooleanVar(value=cp.norm(h) in checked)
            self.vars.append((h, v))
            ttk.Checkbutton(inner, text=f"{cp.col_letter(i)}  {cp.clean_name(h)}", variable=v).pack(anchor="w")
        bar = ttk.Frame(self, padding=8)
        bar.pack(fill="x")
        ttk.Button(bar, text="취소", command=self.destroy).pack(side="right")
        ttk.Button(bar, text="확인", command=self._ok).pack(side="right", padx=4)
        self.grab_set()
        master.wait_window(self)

    def _ok(self):
        self.result = [h for h, v in self.vars if v.get()]
        self.destroy()
