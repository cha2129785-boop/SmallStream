# -*- coding: utf-8 -*-
import os, sys, tempfile, unittest, csv
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT); sys.path.insert(0, os.path.join(ROOT, "tools"))
import colprune as cp  # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.styles import PatternFill  # noqa: E402


def make_xlsx(path):
    wb = Workbook(); ws = wb.active; ws.title = "명단"
    ws.append(["어르신 명단"]); ws.merge_cells("A1:F1")
    ws.append(["인정번호", "성명", "주민번호", "연 락 처", "비고", "메모"])
    for i in range(4):
        ws.append([f"L{i}", f"가{i}", "5001011234567", "010", None, f"m{i}"])
    ws["E3"] = "=A3&B3"
    ws.column_dimensions["B"].width = 20; ws.column_dimensions["D"].width = 30
    ws["D2"].fill = PatternFill("solid", fgColor="FFFF00")
    ws.merge_cells("C3:D3")
    ws.auto_filter.ref = "A2:F6"; ws.freeze_panes = "B3"
    ws2 = wb.create_sheet("두번째"); ws2.append(["주민번호", "x"]); ws2.append(["1", "2"])
    wb.save(path)


class T(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.mkdtemp()

    def test_parse(self):
        self.assertEqual(cp.parse_positions("A, c; F-H 5 h"), [1, 3, 5, 6, 7, 8])
        self.assertEqual(cp.parse_positions("AA~AB"), [27, 28])
        self.assertEqual(cp.parse_positions(""), [])
        with self.assertRaises(cp.PruneError):
            cp.parse_positions("A1")
        self.assertEqual(cp.parse_names(" 성명 \n\n성명\n연락처 "), ["성명", "연락처"])
        self.assertEqual(cp.col_letter(28), "AB")

    def test_plan_names_and_positions(self):
        h = ["인정번호", "성명", "연 락 처", "ＭＥＭＯ", "성명"]
        cols, miss = cp.plan_columns(h, ["연 락처", "memo", "성명", "없음"], [1, 9])
        self.assertEqual(cols, [1, 2, 4, 5])  # 공백·전각·대소문자 무시, 같은 이름 둘 다, 범위 밖 9 무시
        self.assertEqual(miss, ["연 락처", "없음"])  # '연 락처' ≠ '연 락 처' (공백 위치 다름)

    def test_xlsx_keeps_format(self):
        src = os.path.join(self.d, "명단.xlsx"); make_xlsx(src)
        before = os.path.getmtime(src)
        p = cp.Preset("t", ["주민번호"], "F", header_row=2)
        r = cp.prune_file(src, p)
        self.assertEqual(r.error, "")
        self.assertTrue(r.out.endswith("명단_열삭제.xlsx"))
        self.assertEqual(os.path.getmtime(src), before)
        wb = load_workbook(r.out)
        ws = wb["명단"]
        print("명단", ws.max_row, "행 ×", ws.max_column, "열")
        self.assertEqual([c.value for c in ws[2]], ["인정번호", "성명", "연 락 처", "비고"])
        self.assertEqual((ws.max_row, ws.max_column), (6, 4))
        self.assertEqual(r.sheets[0].deleted, ["C 주민번호", "F 메모"])
        self.assertEqual(ws.column_dimensions["B"].width, 20)
        self.assertEqual(ws.column_dimensions["C"].width, 30)
        self.assertEqual(ws["C2"].fill.fgColor.rgb, "00FFFF00")
        self.assertEqual({str(m) for m in ws.merged_cells.ranges}, {"A1:D1"})  # C3:D3 → C3 하나 → 병합 해제
        self.assertEqual(ws.auto_filter.ref, "A2:D6")
        self.assertEqual(ws.freeze_panes, "B3")
        self.assertIsNone(ws["D3"].value)  # 수식 → 값 (Excel 저장 이력 없음 → 빈 값)
        self.assertTrue(any("수식 1개" in w for w in r.warnings))
        # 둘째 시트: 헤더 행 2에 이름 없음 → 고정 자리만 (F 없음) → 그대로
        self.assertEqual(wb["두번째"].max_column, 2)
        self.assertEqual(r.sheets[1].missing, ["주민번호"])

    def test_first_sheet_only_and_keep_formula(self):
        src = os.path.join(self.d, "a.xlsx"); make_xlsx(src)
        r = cp.prune_file(src, cp.Preset("", ["비고"], "", 2, all_sheets=False, formulas_to_values=False))
        self.assertEqual(len(r.sheets), 1)
        self.assertTrue(any("#REF!" in w for w in r.warnings))

    def test_no_overwrite_and_guard(self):
        src = os.path.join(self.d, "a.xlsx"); make_xlsx(src)
        p = cp.Preset("", [], "A")
        r1, r2 = cp.prune_file(src, p), cp.prune_file(src, p)
        self.assertTrue(r2.out.endswith("a_열삭제_2.xlsx"))
        self.assertNotEqual(r1.out, r2.out)
        # 다른 원본 이름과 겹치면 차단
        other = os.path.join(self.d, "b.xlsx"); make_xlsx(other)
        clash = os.path.join(self.d, "b_열삭제.xlsx"); make_xlsx(clash)
        r = cp.prune_files([other, clash], p)
        self.assertTrue(all(x.out and x.out not in (other, clash) for x in r))

    def test_csv_and_xls(self):
        import make_samples as ms
        paths = ms.make_all(os.path.join(self.d, "s"))
        r = cp.prune_file(paths["hospital"], cp.Preset("", ["성명"], ""))
        self.assertEqual(r.error, "", r.error)
        ws = load_workbook(r.out).active
        self.assertNotIn("성명", [c.value for c in ws[1]])
        self.assertEqual((ws.max_row, ws.max_column), (r.sheets[0].rows, r.sheets[0].cols))
        self.assertEqual(ws["A1"].font.size, 11)
        r = cp.prune_file(paths["guardian"], cp.Preset("", ["보호자명"], "", header_row=2))
        self.assertEqual(r.error, "", r.error)
        self.assertIn("보호자명", " ".join(r.sheets[0].deleted))

    def test_preview_writes_nothing(self):
        src = os.path.join(self.d, "a.xlsx"); make_xlsx(src)
        r = cp.preview(src, cp.Preset("", ["성명"], "", 2))
        self.assertEqual(r.sheets[0].deleted, ["B 성명"])
        self.assertEqual(sorted(os.listdir(self.d)), ["a.xlsx"])
        self.assertTrue(cp.preview(src, cp.Preset()).error)

    def test_presets(self):
        st = cp.PresetStore(os.path.join(self.d, "p", "c.json"))
        self.assertEqual(st.load(), [])
        st.put(cp.Preset("월간", ["주민번호"], "C")); st.put(cp.Preset("가", [], "A"))
        st.put(cp.Preset("월간", ["연락처"], ""))
        self.assertEqual([p.name for p in st.load()], ["가", "월간"])
        self.assertEqual(st.load()[1].names, ["연락처"])
        st.delete("가")
        self.assertEqual([p.name for p in st.load()], ["월간"])


if __name__ == "__main__":
    unittest.main()


try:
    import tkinter as tk
    import colprune_ui as ui
    _r = tk.Tk(); _r.destroy(); HAVE_TK = True
except Exception:  # noqa: BLE001
    HAVE_TK = False


@unittest.skipUnless(HAVE_TK, "tkinter/디스플레이 없음")
class TUI(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.mkdtemp()
        self.root = tk.Tk(); self.root.withdraw()
        self.msgs = []
        self._orig = []
        for name in ("showinfo", "showwarning", "showerror"):
            self._patch(ui.messagebox, name, lambda *a, n=name, **k: self.msgs.append((n, a[1] if len(a) > 1 else "")))
        self._patch(ui.messagebox, "askyesno", lambda *a, **k: True)

    def _patch(self, obj, name, fn):
        self._orig.append((obj, name, getattr(obj, name))); setattr(obj, name, fn)

    def tearDown(self):
        for obj, name, fn in reversed(self._orig):
            setattr(obj, name, fn)
        self.root.destroy()

    def test_dialog_flow(self):
        src = os.path.join(self.d, "a.xlsx"); make_xlsx(src)
        self._patch(ui.filedialog, "askopenfilenames", lambda **k: (src,))
        dlg = ui.ColPruneDialog(self.root, os.path.join(self.d, "presets.json"), "스트림머지", self.d)
        self.root.update()
        dlg.run()
        self.assertEqual(self.msgs[-1][0], "showinfo")  # 파일 없음
        dlg.add_files()
        self.assertEqual(dlg.lb.get(0, "end"), ("a.xlsx",))
        dlg.run()
        self.assertEqual(self.msgs[-1][0], "showwarning")  # 열 지정 없음
        # 첫 파일에서 고르기 (헤더 2행)
        dlg.v_hdr.set(2)

        class Pick:
            def __init__(s, master, header, checked):
                self.assertEqual(header[:3], ["인정번호", "성명", "주민번호"]); s.result = ["주민번호", "메모"]
        self._patch(ui, "NamePickDialog", Pick)
        dlg.pick_names()
        self.assertEqual(dlg.txt.get("1.0", "end").split(), ["주민번호", "메모"])
        # 프리셋 저장 → 지우고 → 불러오기
        self._patch(ui.simpledialog, "askstring", lambda *a, **k: "월간 정리")
        dlg.v_pos.set("E")
        dlg.save_preset()
        self.assertEqual(list(dlg.cb["values"]), ["월간 정리"])
        dlg.txt.delete("1.0", "end"); dlg.v_pos.set(""); dlg.v_hdr.set(1)
        dlg.v_preset.set("월간 정리"); dlg.load_preset()
        self.assertEqual((dlg.v_pos.get(), int(dlg.v_hdr.get())), ("E", 2))
        # 미리보기 → 파일 안 생김
        dlg.preview(); self.root.update()
        self.assertFalse(any(f.endswith("_열삭제.xlsx") for f in os.listdir(self.d)))
        # 실행
        dlg.run(); self.root.update()
        self.assertEqual(len(dlg.reports), 1)
        out = dlg.reports[0].out
        self.assertEqual([c.value for c in load_workbook(out)["명단"][2]], ["인정번호", "성명", "연 락 처"])
        txt = ui.report_text(dlg.reports, True)
        self.assertIn("삭제 3열: C 주민번호, E 비고, F 메모", txt)
        self.assertIn("6행 × 3열", txt)
        dlg.destroy()
