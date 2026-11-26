# -*- coding: utf-8 -*-
"""파일 정리 도구 v0.3: 서식 지우기(테두리 유지) · 중복 내용 지우기"""
import csv, json, os, sys, tempfile, unittest
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT); sys.path.insert(0, os.path.join(ROOT, "tools"))
import colprune as cp  # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.formatting.rule import CellIsRule  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402

THIN = Side(style="thin")


def make_styled(path):
    wb = Workbook(); ws = wb.active; ws.title = "명단"
    ws.append(["제목"]); ws.merge_cells("A1:C1")
    ws.append(["성명", "주소", "생일"])
    rows = [("홍길동", "서울", 1), ("홍길동", "부산", 2), ("김철수", "서울", 3),
            ("홍길동 ", "서울", 4), ("김철수", "서울", 3)]
    for r in rows:
        ws.append(list(r))
    ws["A3"].font = Font(name="굴림", size=15, bold=True, italic=True, underline="single", strike=True, color="FF0000")
    ws["A3"].fill = PatternFill("solid", fgColor="FFFF00")
    ws["A3"].alignment = Alignment(horizontal="center", wrap_text=True)
    ws["A3"].border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    ws["C3"].number_format = "yyyy-mm-dd"
    ws["B10"].fill = PatternFill("solid", fgColor="00FF00")      # 값 없는 칸의 색
    ws.column_dimensions["B"].width = 30
    ws.row_dimensions[3].height = 40
    ws.conditional_formatting.add("C3:C7", CellIsRule(operator="greaterThan", formula=["1"],
                                                      fill=PatternFill("solid", fgColor="FF0000")))
    wb.save(path)


class TPlan(unittest.TestCase):
    rows = [["h1", "h2"], ["a", "x"], ["a", "x"], ["b", "a"], ["a", "x"], ["", ""], ["b", "y"]]

    def test_modes(self):
        p = cp.dedupe_plan
        self.assertEqual(p(self.rows, 1, [1], "col"), [(2, 1), (4, 1), (6, 1)])
        self.assertEqual(p(self.rows, 1, [1], "consecutive"), [(2, 1)])
        # 시트 전체: B4 'a' 는 A2 'a' 와 같음 (행→열 순서)
        self.assertEqual(p(self.rows, 1, [1, 2], "sheet"), [(2, 1), (2, 2), (3, 2), (4, 1), (4, 2), (6, 1)])
        self.assertEqual(p(self.rows, 1, [1, 2], "row"), [(2, 1), (2, 2), (4, 1), (4, 2)])
        self.assertEqual(p(self.rows, 0, [1], "col")[:1], [(2, 1)])  # 헤더 'h1' 은 다른 값
        with self.assertRaises(cp.PruneError):
            p(self.rows, 1, [1], "zzz")

    def test_kinds_skip(self):
        import datetime as dt
        k = cp.value_kind
        self.assertEqual([k(x) for x in ("홍길동", "TRUE", True, "1,234", "-3.5", "10%", "2026-11-10",
                                         "010-1234-5678", 3, 2.5, dt.datetime(2026, 1, 1), "A101", "3층", "2반 ")],
                         ["text", "text", "text", "num", "num", "num", "num",
                          "num", "num", "num", "num", "mixed", "mixed", "mixed"])
        rows = [["h"], ["가"], ["가"], ["1"], ["1"], ["A1"], ["A1"], ["가"]]
        self.assertEqual(cp.dedupe_plan(rows, 1, [1], "col"), [(2, 1), (4, 1), (6, 1), (7, 1)])
        self.assertEqual(cp.dedupe_plan(rows, 1, [1], "col", ["num"]), [(2, 1), (6, 1), (7, 1)])
        self.assertEqual(cp.dedupe_plan(rows, 1, [1], "col", ["text", "mixed"]), [(4, 1)])
        # 행 중복: 제외 칸도 비교엔 포함 → 숫자 다르면 중복 아님, 같으면 글자만 지움
        rr = [["a", "b"], ["가", 1], ["가", 2], ["가", 1]]
        self.assertEqual(cp.dedupe_plan(rr, 1, [1, 2], "row", ["num"]), [(3, 1)])
        with self.assertRaises(cp.PruneError):
            cp.Preset("x", dedupe=True, dedupe_all_cols=True, dedupe_skip=["text", "num", "mixed"]).check()
        with self.assertRaises(cp.PruneError):
            cp.Preset("x", dedupe=True, dedupe_all_cols=True, dedupe_skip=["zz"]).check()

    def test_names_newline_comma_and_except_missing(self):
        """v0.17 버그: 헤더 줄바꿈·쉼표 이름이 쪼개짐 / 제외 열 못 찾으면 모든 열 지움"""
        self.assertEqual(cp.clean_name("성\n명 "), "성 명")
        j = cp.join_name_list(["성\n명", "성명,이름", '별"명', "반"])
        self.assertEqual(j, '성 명, "성명,이름", "별""명", 반')
        self.assertEqual(cp.parse_name_list(j), ["성 명", "성명,이름", '별"명', "반"])
        self.assertEqual(cp.parse_name_list("성명, 주소\n성명"), ["성명", "주소"])
        hdr = ["성\n명", "성명,이름", "반"]
        self.assertEqual(cp.plan_columns(hdr, cp.parse_name_list(j), [])[0], [1, 2, 3])  # 모두 찾음
        p = cp.Preset("x", dedupe=True, dedupe_except=True, dedupe_names=["없는열"])
        self.assertEqual(p.dedupe_cols(["a", "b"], 2), ([], ["없는열"]))  # 못 찾음 → 지우지 않음
        self.assertIn("중복 지우기 안 함", p.except_warning("시트", ["없는열"]))
        self.assertIsNone(cp.Preset("x", dedupe=True, dedupe_names=["없는열"]).except_warning("시트", ["없는열"]))

    def test_check_and_compat(self):
        with self.assertRaises(cp.PruneError):
            cp.Preset("x").check()                                          # 할 작업 없음
        cp.Preset("x", clear_format=True).check()                           # 서식만 → OK
        with self.assertRaises(cp.PruneError):
            cp.Preset("x", dedupe=True).check()                             # 대상 열 없음
        cp.Preset("x", dedupe=True, dedupe_all_cols=True).check()
        cp.Preset("x", dedupe=True, dedupe_positions="A").check()
        self.assertEqual(cp.Preset("x", ["a"]).suffix(), "_열삭제")
        self.assertEqual(cp.Preset("x", ["a"], dedupe=True, dedupe_all_cols=True).suffix(), "_정리")
        self.assertEqual(cp.parse_name_list("성명, 주소\n성명"), ["성명", "주소"])
        # 옛 프리셋 파일 (v0.2) → 새 옵션 꺼짐
        d = tempfile.mkdtemp(); f = os.path.join(d, "p.json")
        json.dump({"format": 1, "presets": [{"name": "옛", "names": ["성명"], "positions": "", "header_row": 1,
                                              "all_sheets": True, "formulas_to_values": True}]},
                  open(f, "w", encoding="utf-8"))
        old = cp.PresetStore(f).load()[0]
        self.assertFalse(old.clear_format or old.dedupe)
        st = cp.PresetStore(f); st.put(cp.Preset("새", clear_format=True, dedupe=True, dedupe_mode="row",
                                                  dedupe_names=["성명"]))
        new = next(x for x in st.load() if x.name == "새")
        self.assertEqual((new.dedupe_mode, new.dedupe_names, new.clear_format), ("row", ["성명"], True))


class TFile(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.mkdtemp()
        self.src = os.path.join(self.d, "명단.xlsx"); make_styled(self.src)
        self.mtime = os.path.getmtime(self.src)

    def test_clear_format(self):
        r = cp.prune_file(self.src, cp.Preset("f", header_row=2, clear_format=True))
        self.assertEqual(r.error, "")
        self.assertTrue(r.out.endswith("명단_정리.xlsx"))
        self.assertEqual(os.path.getmtime(self.src), self.mtime)          # 원본 그대로
        ws = load_workbook(r.out)["명단"]
        print("서식 지움", ws.max_row, "행 ×", ws.max_column, "열", r.notes)
        a = ws["A3"]
        self.assertFalse(any((a.font.b, a.font.i, a.font.u, a.font.strike)))
        self.assertEqual((a.font.name, a.font.sz), ("맑은 고딕", 11))
        self.assertTrue(a.font.color is None or a.font.color.rgb != "00FF0000")
        self.assertEqual(a.fill.fill_type, None)
        self.assertEqual((a.alignment.horizontal, a.alignment.wrap_text), (None, None))
        self.assertEqual(a.border.left.style, "thin")                    # 테두리 유지
        self.assertEqual(ws["C3"].number_format, "yyyy-mm-dd")           # 표시 형식 유지
        self.assertEqual(ws["B10"].fill.fill_type, None)
        self.assertEqual(ws.column_dimensions["B"].width, 30)
        self.assertEqual(ws.row_dimensions[3].height, 40)
        self.assertEqual(len(ws.conditional_formatting), 0)
        self.assertEqual(list(ws.merged_cells.ranges)[0].coord, "A1:C1")
        self.assertEqual(ws["A3"].value, "홍길동")                         # 값 그대로
        self.assertTrue(any("조건부 서식 1개" in n for n in r.notes))

    def test_clear_format_options(self):
        p = cp.Preset("f", clear_format=True, keep_numfmt=False, clear_align=False, unify_font=False, keep_size=False)
        ws = load_workbook(cp.prune_file(self.src, p).out)["명단"]
        self.assertEqual(ws["C3"].number_format, "General")
        self.assertEqual(ws["A3"].alignment.horizontal, "center")
        self.assertNotEqual(ws["A3"].font.name, "굴림")
        self.assertFalse(ws["A3"].font.b)
        self.assertEqual(ws.row_dimensions[3].height, None)
        self.assertEqual(ws.column_dimensions["B"].width, 13)
        self.assertEqual(ws["A3"].border.left.style, "thin")

    def test_align(self):
        """v0.18: 가로 정렬 왼/가운데/오른쪽 — 세로·줄바꿈 유지, 서식 지우기와 함께도 적용"""
        from openpyxl.styles import Alignment as A
        wb = load_workbook(self.src); wb["명단"]["B3"].alignment = A(vertical="top", wrap_text=True); wb.save(self.src)
        for h in ("left", "center", "right"):
            r = cp.prune_file(self.src, cp.Preset("a", align=h))
            self.assertEqual(r.error, "")
            ws = load_workbook(r.out)["명단"]
            self.assertEqual({ws.cell(i, j).alignment.horizontal for i in range(1, 8) for j in range(1, 4)
                              if ws.cell(i, j).value is not None}, {h})
            self.assertEqual((ws["B3"].alignment.vertical, ws["B3"].alignment.wrap_text), ("top", True))
            self.assertIsNone(ws["B10"].alignment.horizontal)                  # 빈 칸은 그대로
            self.assertTrue(r.out.endswith("_정리.xlsx") or "_정리_" in r.out)
            self.assertTrue(any(cp.ALIGNS[h] in n for n in r.notes))
        # 서식 지우기(정렬 지움)와 함께 → 정렬이 이김
        ws = load_workbook(cp.prune_file(self.src, cp.Preset("a", clear_format=True, align="right")).out)["명단"]
        self.assertEqual(ws["A3"].alignment.horizontal, "right")
        with self.assertRaises(cp.PruneError):
            cp.Preset("a", align="middle").check()
        cp.Preset("a", align="center").check()                                 # 정렬만으로도 작업
        # csv 경로
        f = os.path.join(self.d, "a.csv")
        with open(f, "w", encoding="utf-8-sig", newline="") as fp:
            csv.writer(fp).writerows([["성명", "반"], ["가", ""]])
        ws = load_workbook(cp.prune_file(f, cp.Preset("a", align="center")).out).active
        self.assertEqual((ws["A1"].alignment.horizontal, ws["A2"].alignment.horizontal, ws["B2"].alignment.horizontal),
                         ("center", "center", None))
        self.assertIn("정렬 가운데", " ".join(cp.preview(self.src, cp.Preset("a", align="center")).notes))

    def test_dedupe_col_by_name(self):
        p = cp.Preset("d", header_row=2, dedupe=True, dedupe_names=["성명", "없는열"])
        r = cp.prune_file(self.src, p)
        ws = load_workbook(r.out)["명단"]
        col = [ws.cell(i, 1).value for i in range(1, 8)]
        print("중복 지움(열마다)", col)
        # '홍길동 ' 은 공백 무시 → 중복
        self.assertEqual(col, ["제목", "성명", "홍길동", None, "김철수", None, None])
        self.assertEqual([ws.cell(i, 2).value for i in range(3, 8)], ["서울", "부산", "서울", "서울", "서울"])
        self.assertEqual(r.sheets[0].cleared, 3)
        self.assertIn("(중복) 없는열", r.sheets[0].missing)
        # 제외 열 (v0.14): 성명 빼고 모든 열 → 주소·생일만 중복 지움
        r3 = cp.prune_file(self.src, cp.Preset("d", header_row=2, dedupe=True, dedupe_except=True, dedupe_names=["성명"]))
        ws3 = load_workbook(r3.out)["명단"]
        self.assertEqual([ws3.cell(i, 1).value for i in range(3, 8)], ["홍길동", "홍길동", "김철수", "홍길동 ", "김철수"])
        self.assertEqual([ws3.cell(i, 2).value for i in range(3, 8)], ["서울", "부산", None, None, None])
        self.assertEqual(ws3.cell(7, 3).value, None)                       # 생일 3 중복
        with self.assertRaises(cp.PruneError):
            cp.Preset("d", dedupe=True, dedupe_except=True).check()        # 제외 열 비어 있음
        # 제외 열 이름 틀림 → 파일에서도 중복 안 지움 + 경고
        r4 = cp.prune_file(self.src, cp.Preset("d", header_row=2, dedupe=True, dedupe_except=True, dedupe_names=["이름"]))
        ws4 = load_workbook(r4.out)["명단"]
        self.assertEqual(ws4.cell(4, 1).value, "홍길동")
        self.assertEqual(r4.sheets[0].cleared, 0)
        self.assertTrue(any("제외 열 못 찾음" in w for w in r4.warnings))
        self.assertTrue(any("제외 열 못 찾음" in w for w in cp.preview(self.src, cp.Preset("d", header_row=2, dedupe=True,
                            dedupe_except=True, dedupe_names=["이름"])).warnings))
        self.assertEqual(cp.Preset("d", dedupe=True, dedupe_except=True, dedupe_positions="A").dedupe_cols(["a", "b", "c"], 3)[0], [2, 3])
        # 생일(숫자) 열 포함 모든 열, 숫자 제외 → 생일 3 중복 유지
        r2 = cp.prune_file(self.src, cp.Preset("d", header_row=2, dedupe=True, dedupe_all_cols=True, dedupe_skip=["num"]))
        ws2 = load_workbook(r2.out)["명단"]
        self.assertEqual([ws2.cell(i, 3).value for i in range(4, 8)], [2, 3, 4, 3])   # C3 는 날짜 형식 칸
        self.assertEqual(ws2["A4"].value, None)
        self.assertEqual(ws["A3"].fill.fill_type, "solid")               # 서식 지우기 안 켬 → 유지
        # 미리보기 개수 = 실행 개수, 파일 안 만듦
        before = set(os.listdir(self.d))
        self.assertEqual(cp.preview(self.src, p).sheets[0].cleared, 3)
        self.assertEqual(set(os.listdir(self.d)), before)

    def test_dedupe_row_and_delete(self):
        # 행 중복(모든 열) + 주소 열 삭제: 중복 판정은 삭제 전 열 기준
        p = cp.Preset("d", ["주소"], header_row=2, dedupe=True, dedupe_mode="row", dedupe_all_cols=True)
        r = cp.prune_file(self.src, p)
        ws = load_workbook(r.out)["명단"]
        self.assertEqual([c.value for c in ws[2]], ["성명", "생일"])
        self.assertEqual([c.value for c in ws[7]], [None, None])          # 5행 = 3행 복제 → 지움
        self.assertEqual([c.value for c in ws[5]], ["김철수", 3])
        self.assertEqual(ws["A1"].value, "제목")                           # 헤더 위 손대지 않음
        txt = __import__("colprune_ui").report_text([r], True)
        self.assertIn("중복 지움 3칸", txt)                                # 7행 3칸
        self.assertIn("삭제 1열: B 주소", txt)

    def test_csv(self):
        f = os.path.join(self.d, "a.csv")
        with open(f, "w", encoding="utf-8-sig", newline="") as fp:
            csv.writer(fp).writerows([["성명", "반"], ["가", "1"], ["가", "1"], ["나", "1"]])
        r = cp.prune_file(f, cp.Preset("c", dedupe=True, dedupe_mode="consecutive", dedupe_positions="A-B",
                                       clear_format=True))
        ws = load_workbook(r.out).active
        self.assertEqual([[c.value for c in row] for row in ws.iter_rows()],
                         [["성명", "반"], ["가", "1"], [None, None], ["나", None]])
        self.assertFalse(ws["A1"].font.b)                                 # 서식 지움 → 헤더 굵게 뺌



try:
    import tkinter as tk
    tk.Tk().destroy()
    HAS_TK = True
except Exception:  # noqa: BLE001
    HAS_TK = False


@unittest.skipUnless(HAS_TK, "tkinter 없음")
class TGui(unittest.TestCase):
    def test_dialog_options(self):
        import colprune_ui as ui
        d = tempfile.mkdtemp(); src = os.path.join(d, "명단.xlsx"); make_styled(src)
        root = tk.Tk(); root.withdraw()
        orig = (ui.filedialog.askopenfilenames, ui.simpledialog.askstring, ui.NamePickDialog, ui.messagebox.showinfo)
        try:
            ui.filedialog.askopenfilenames = lambda **k: (src,)
            ui.simpledialog.askstring = lambda *a, **k: "정리"
            ui.messagebox.showinfo = lambda *a, **k: None
            dlg = ui.ColPruneDialog(root, os.path.join(d, "p.json"), "스트림머지", d)
            root.update()
            self.assertIn("파일 정리 도구", dlg.title())
            self.assertIn("disabled", dlg.e_ddnames.state())               # 중복 꺼짐 → 비활성
            dlg.add_files(); dlg.v_hdr.set(2)
            dlg.v_fmt.set(True); dlg.v_dd.set(True); dlg._sync_clean()
            self.assertNotIn("disabled", dlg.e_ddnames.state())

            class Pick:
                def __init__(s, master, header, checked, title="지울 열 고르기"):
                    s.result = ["성명"]; titles.append(title)
            titles = []
            ui.NamePickDialog = Pick
            dlg.pick_dd_names()
            self.assertEqual(dlg.v_ddnames.get(), "성명")
            self.assertEqual(titles, ["중복을 지울 열 고르기"])
            dlg.v_ddmode.set(cp.DEDUPE_MODES["row"]); dlg.v_skip["mixed"].set(True)
            dlg.v_ddtgt.set("except"); dlg._sync_clean(); dlg.v_halign.set("right"); dlg.v_align.set(False)
            dlg.v_where.set("dir"); dlg.v_dir.set(d); dlg.v_namec.set(True); dlg.v_name.set("{원본}_월간"); dlg._sync_where()
            self.assertNotIn("disabled", dlg.e_name.state())
            self.assertNotIn("disabled", dlg.e_ddnames.state())
            dlg.save_preset()
            dlg.v_fmt.set(False); dlg.v_dd.set(False); dlg.v_ddnames.set(""); dlg.v_skip["mixed"].set(False)
            dlg.v_ddtgt.set("all"); dlg.v_halign.set(""); dlg.v_align.set(True)
            dlg.v_where.set("same"); dlg.v_dir.set(""); dlg.v_namec.set(False); dlg.v_name.set("")
            dlg.v_preset.set("정리"); dlg.load_preset()
            p = dlg.current()
            self.assertEqual((p.clear_format, p.dedupe, p.dedupe_mode, p.dedupe_names, p.dedupe_skip, p.dedupe_except,
                              p.dedupe_all_cols, p.align, p.clear_align), (True, True, "row", ["성명"], ["mixed"], True, False, "right", False))
            self.assertEqual((p.save_where, p.save_dir, p.name_custom, p.name_template), ("dir", d, True, "{원본}_월간"))
            self.assertNotIn("disabled", dlg.e_name.state())
            dlg.v_where.set("orig"); self.assertIn("disabled", dlg.e_name.state())
            self.assertIn("disabled", dlg.cb_move.state())                 # 폴더 지정일 때만
            dlg.v_where.set("dir"); self.assertNotIn("disabled", dlg.cb_move.state())
            dlg.run(); root.update()
            self.assertTrue(dlg.reports[0].out.endswith("명단_월간.xlsx"))
            dlg.update_idletasks()
            print("창 요청 크기", dlg.winfo_reqwidth(), "×", dlg.winfo_reqheight())
            self.assertLessEqual(dlg.winfo_reqheight(), 728)   # 768 - 작업 표시줄
            dlg.destroy()
        finally:
            (ui.filedialog.askopenfilenames, ui.simpledialog.askstring, ui.NamePickDialog, ui.messagebox.showinfo) = orig
            root.destroy()


if __name__ == "__main__":
    unittest.main()
