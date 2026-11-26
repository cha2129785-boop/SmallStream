# -*- coding: utf-8 -*-
"""v0.13: 같은 이름 처리(덮어쓰기/새 이름/묻기) · 원본 덮어쓰기(백업)"""
import csv, os, sys, tempfile, unittest
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT); sys.path.insert(0, os.path.join(ROOT, "tools"))
import colprune as cp  # noqa: E402
import outpath as op  # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402


def xlsx(path, rows):
    wb = Workbook(); ws = wb.active
    for r in rows:
        ws.append(r)
    wb.save(path)


class T(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.mkdtemp()
        self._env = os.environ.get("SMALLSTREAM_HOME")
        os.environ["SMALLSTREAM_HOME"] = os.path.join(self.d, "앱")

    def tearDown(self):
        if self._env is None:
            os.environ.pop("SMALLSTREAM_HOME", None)
        else:
            os.environ["SMALLSTREAM_HOME"] = self._env

    def test_settings_and_resolver(self):
        self.assertEqual(op.get_mode(), "ask")                       # 기본 = 매번 묻기
        op.set_mode("overwrite"); self.assertEqual(op.get_mode(), "overwrite")
        with self.assertRaises(ValueError):
            op.set_mode("zz")
        a = os.path.join(self.d, "a.xlsx"); open(a, "w").close()
        self.assertEqual(op.Resolver("overwrite").resolve(a), a)
        self.assertTrue(op.Resolver("rename").resolve(a).endswith("a_2.xlsx"))
        self.assertTrue(op.Resolver("ask").resolve(a).endswith("a_2.xlsx"))  # 묻기 함수 없음 → 새 이름
        calls = []

        def ask(p, many):
            calls.append(p); return ("overwrite", True)
        r = op.Resolver("ask", ask, many=True)
        self.assertEqual(r.resolve(a), a); self.assertEqual(r.resolve(a), a)
        self.assertEqual(len(calls), 1)                               # 나머지도 같게
        self.assertIsNone(op.Resolver("ask", lambda p, m: (None, False)).resolve(a))
        self.assertTrue(op.Resolver("overwrite").resolve(a, guard=[a]).endswith("a_2.xlsx"))  # 원본은 덮지 않음
        self.assertEqual(op.Resolver("ask", ask).resolve(os.path.join(self.d, "new.xlsx")),
                         os.path.join(self.d, "new.xlsx"))            # 없는 이름 → 묻지 않음

    def test_prune_overwrite_result_and_original(self):
        src = os.path.join(self.d, "명단.xlsx")
        xlsx(src, [["성명", "반"], ["가", "1"], ["가", "1"]])
        p = cp.Preset("p", dedupe=True, dedupe_all_cols=True)
        r1 = cp.prune_file(src, p)
        r2 = cp.prune_file(src, p, resolver=op.Resolver("overwrite"))
        self.assertEqual(r1.out, r2.out)                              # 같은 이름 결과 덮어씀
        r3 = cp.prune_file(src, p)
        self.assertTrue(r3.out.endswith("명단_정리_2.xlsx"))
        r4 = cp.prune_file(src, p, resolver=op.Resolver("ask", lambda q, m: (None, False)))
        self.assertEqual((r4.error, r4.out), ("저장 취소", ""))
        # 원본 덮어쓰기: 백업 후 교체
        bdir = os.path.join(self.d, "backup")
        with self.assertRaises(AssertionError):
            self.assertEqual(cp.prune_file(src, p, to_original=True).error, "")  # 백업 폴더 없음 → 거부
        r = cp.prune_file(src, p, to_original=True, backup_dir=bdir)
        self.assertEqual((r.error, r.out), ("", src))
        self.assertTrue(os.path.isfile(r.backup) and r.backup.startswith(bdir))
        print("백업", os.path.basename(r.backup))
        self.assertEqual([c.value for c in load_workbook(src).active["A"]], ["성명", "가"])   # 3행 전부 지움
        self.assertEqual([c.value for c in load_workbook(r.backup).active["A"]], ["성명", "가", "가"])
        self.assertFalse([f for f in os.listdir(self.d) if f.startswith("~sm_")])
        txt = __import__("colprune_ui").report_text([r], True)
        self.assertIn("원본 덮어씀", txt); self.assertIn("백업:", txt)
        # 두 번째 백업은 이름 겹치지 않음
        r_b = cp.prune_file(src, p, to_original=True, backup_dir=bdir)
        self.assertNotEqual(r_b.backup, r.backup)
        self.assertEqual(len(os.listdir(bdir)), 2)

    def test_original_csv_xls(self):
        f = os.path.join(self.d, "a.csv")
        with open(f, "w", encoding="utf-8-sig", newline="") as fp:
            csv.writer(fp).writerows([["성명"], ["가"], ["가"]])
        bdir = os.path.join(self.d, "backup")
        r = cp.prune_file(f, cp.Preset("p", dedupe=True, dedupe_all_cols=True), to_original=True, backup_dir=bdir)
        self.assertEqual(r.out, os.path.join(self.d, "a.xlsx"))
        self.assertTrue(os.path.isfile(f))                            # csv 원본 유지
        self.assertEqual(r.backup, "")
        self.assertTrue(any("같은 이름 .xlsx" in n for n in r.notes))
        self.assertFalse(os.path.exists(bdir))
        # a.xlsx 가 이미 있음 → 같은 이름 규칙 (새 이름)
        r2 = cp.prune_file(f, cp.Preset("p", dedupe=True, dedupe_all_cols=True), to_original=True, backup_dir=bdir)
        self.assertTrue(r2.out.endswith("a_2.xlsx"))


class TMove(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.mkdtemp()
        self.src_dir = os.path.join(self.d, "원본"); self.out = os.path.join(self.d, "결과"); self.bk = os.path.join(self.d, "백업")
        os.makedirs(self.src_dir); os.makedirs(self.out)
        self.a = os.path.join(self.src_dir, "명단.xlsx")
        xlsx(self.a, [["성명"], ["가"], ["가"]])
        self.p = cp.Preset("p", dedupe=True, dedupe_all_cols=True)

    def test_move_flow(self):
        r = cp.prune_file(self.a, self.p, self.out, move_original=True, backup_dir=self.bk)
        self.assertEqual((r.error, r.out, r.moved), ("", os.path.join(self.out, "명단.xlsx"), True))
        self.assertFalse(os.path.exists(self.a))                                    # 원본 이동됨
        self.assertTrue(r.backup.startswith(self.bk) and os.path.isfile(r.backup))
        self.assertEqual([c.value for c in load_workbook(r.backup).active["A"]], ["성명", "가", "가"])
        self.assertEqual([c.value for c in load_workbook(r.out).active["A"]], ["성명", "가"])
        print("이동", os.path.basename(r.backup), "→", os.path.basename(r.out))
        txt = __import__("colprune_ui").report_text([r], True)
        self.assertIn("원본 이동:", txt); self.assertNotIn("원본 덮어씀", txt)
        # 같은 이름이 지정 폴더에 이미 있음 → _1, _2
        xlsx(self.a, [["성명"], ["나"], ["나"]])
        r2 = cp.prune_file(self.a, self.p, self.out, move_original=True, backup_dir=self.bk)
        self.assertTrue(r2.out.endswith("명단_1.xlsx"))
        xlsx(self.a, [["성명"]])
        self.assertTrue(cp.prune_file(self.a, self.p, self.out, move_original=True, backup_dir=self.bk).out.endswith("명단_2.xlsx"))
        self.assertEqual(len(os.listdir(self.bk)), 3)
        self.assertFalse([f for f in os.listdir(self.out) if f.startswith("~sm_")])

    def test_move_same_folder_csv_and_errors(self):
        # 저장 폴더 = 원본 폴더 → 같은 자리에 정리본 (번호 없음)
        r = cp.prune_file(self.a, self.p, self.src_dir, move_original=True, backup_dir=self.bk)
        self.assertEqual(r.out, self.a)
        self.assertEqual([c.value for c in load_workbook(self.a).active["A"]], ["성명", "가"])
        # csv → 명단2.xlsx, csv 원본은 백업으로
        f = os.path.join(self.src_dir, "명단2.csv")
        with open(f, "w", encoding="utf-8-sig", newline="") as fp:
            csv.writer(fp).writerows([["성명"], ["가"]])
        r = cp.prune_file(f, self.p, self.out, move_original=True, backup_dir=self.bk)
        self.assertEqual(r.out, os.path.join(self.out, "명단2.xlsx"))
        self.assertFalse(os.path.exists(f)); self.assertTrue(r.backup.endswith(".csv"))
        # 처리 실패 → 원본 그대로
        bad = os.path.join(self.src_dir, "깨짐.xlsx"); open(bad, "wb").write(b"PK\x03\x04 broken")
        r = cp.prune_file(bad, self.p, self.out, move_original=True, backup_dir=self.bk)
        self.assertTrue(r.error); self.assertTrue(os.path.exists(bad)); self.assertFalse(r.moved)
        self.assertFalse([x for x in os.listdir(self.out) if x.startswith("~sm_")])
        # 폴더·백업 미지정
        xlsx(self.a, [["성명"]])
        self.assertIn("저장 폴더", cp.prune_file(self.a, self.p, None, move_original=True, backup_dir=self.bk).error)
        self.assertIn("백업 폴더", cp.prune_file(self.a, self.p, self.out, move_original=True).error)
        self.assertTrue(os.path.exists(self.a))
        self.assertEqual(op.next_free(self.a, start=1), self.a[:-5] + "_1.xlsx")


class TName(unittest.TestCase):
    """v0.19: 직접 입력 파일 이름 · v0.18 프리셋 호환"""
    def test_render(self):
        import datetime as dt
        now = dt.datetime(2026, 11, 13, 9, 5)
        r = cp.render_name
        self.assertEqual(r("{원본}_정리_{날짜}", "C:/x/명단.xls", now=now), "명단_정리_20261113")
        self.assertEqual(r("월간보고.xlsx", "a.xlsx", now=now), "월간보고")
        self.assertEqual(r("월간_{시각}", "a.xlsx", 2, 3, now), "월간_0905_2")          # 여러 파일 + {원본} 없음
        self.assertEqual(r("{원본}_월간", "a.xlsx", 2, 3, now), "a_월간")
        for bad in ("", "  ", "a/b", "a?b", "끝점."):
            with self.assertRaises(cp.PruneError):
                r(bad, "a.xlsx", now=now)

    def test_files_with_name(self):
        d = tempfile.mkdtemp(); out = os.path.join(d, "결과"); os.makedirs(out)
        a, b = os.path.join(d, "a.xlsx"), os.path.join(d, "b.xlsx")
        for f in (a, b):
            xlsx(f, [["성명"], ["가"], ["가"]])
        p = cp.Preset("p", dedupe=True, dedupe_all_cols=True)
        reps = cp.prune_files([a, b], p, out, name_template="정리본")
        self.assertEqual([os.path.basename(r.out) for r in reps], ["정리본_1.xlsx", "정리본_2.xlsx"])
        reps = cp.prune_files([a], p, out, name_template="{원본}_완료")
        self.assertEqual(os.path.basename(reps[0].out), "a_완료.xlsx")
        reps = cp.prune_files([a], p, out, name_template="{원본}_완료")                  # 같은 이름 → 새 이름(기본)
        self.assertEqual(os.path.basename(reps[0].out), "a_완료_2.xlsx")
        self.assertIn("쓸 수 없는", cp.prune_files([a], p, out, name_template="a:b")[0].error)
        # 원본 덮어쓰기는 이름 무시
        r = cp.prune_files([a], p, None, to_original=True, backup_dir=os.path.join(d, "bk"), name_template="무시")[0]
        self.assertEqual(r.out, a)
        # 원본 이동 + 직접 이름
        r = cp.prune_files([b], p, out, move_original=True, backup_dir=os.path.join(d, "bk"), name_template="이동본")[0]
        self.assertEqual((os.path.basename(r.out), r.moved, os.path.exists(b)), ("이동본.xlsx", True, False))

    def test_preset_save_fields_and_v018_compat(self):
        import json
        d = tempfile.mkdtemp(); f = os.path.join(d, "p.json")
        st = cp.PresetStore(f)
        st.put(cp.Preset("저장", align="center", save_where="dir", save_dir="D:/결과", save_move=True, save_mode="overwrite",
                         name_custom=True, name_template="{원본}_월간"))
        q = st.load()[0]
        self.assertEqual((q.save_where, q.save_dir, q.save_move, q.save_mode, q.name_custom, q.name_template),
                         ("dir", "D:/결과", True, "overwrite", True, "{원본}_월간"))
        # v0.18 버그로 clear_align 이 글자 → 기본값 True, 저장 설정 없는 옛 프리셋 → same/자동
        json.dump({"format": 1, "presets": [{"name": "옛", "clear_format": True, "clear_align": "right", "align": "right"}]},
                  open(f, "w", encoding="utf-8"))
        q = cp.PresetStore(f).load()[0]
        self.assertEqual((q.clear_align, q.align, q.save_where, q.name_custom), (True, "right", "same", False))


class TLock(unittest.TestCase):
    def test_is_locked(self):
        from unittest import mock
        d = tempfile.mkdtemp(); f = os.path.join(d, "a.xlsx"); xlsx(f, [["x"]])
        t0 = os.path.getmtime(f)
        self.assertFalse(op.is_locked(f))
        self.assertEqual(os.path.getmtime(f), t0)                      # 수정 시각 안 바뀜
        self.assertFalse(op.is_locked(os.path.join(d, "없음.xlsx")))
        with mock.patch("builtins.open", side_effect=PermissionError(13, "sharing violation")):
            self.assertTrue(op.is_locked(f))                           # Windows 공유 위반 흉내

    def test_dialog_precheck(self):
        try:
            import tkinter as tk
            root = tk.Tk(); root.withdraw()
        except Exception:  # noqa: BLE001
            self.skipTest("tkinter 없음")
        import colprune_ui as ui
        d = tempfile.mkdtemp(); out = os.path.join(d, "결과"); os.makedirs(out)
        a, b = os.path.join(d, "a.xlsx"), os.path.join(d, "b.xlsx")
        xlsx(a, [["성명"], ["가"], ["가"]]); xlsx(b, [["성명"], ["나"], ["나"]])
        orig = (op.is_locked, ui.ask_locked, ui.messagebox.askyesno, ui.ResultDialog)
        locked = {a}
        answers = []
        try:
            op.is_locked = lambda f: f in locked
            ui.messagebox.askyesno = lambda *x, **k: True
            ui.ResultDialog = lambda *x, **k: None

            def ask(parent, paths):
                c = answers.pop(0); seen.append(list(paths))
                if c == "retry":
                    locked.clear()                                     # 사용자가 Excel에서 닫음
                return c
            seen = []
            ui.ask_locked = ask
            dlg = ui.ColPruneDialog(root, os.path.join(d, "p.json"), "스트림머지", d)
            dlg.files = [a, b]; dlg.v_dd.set(True); dlg.v_ddtgt.set("all")
            dlg.v_where.set("dir"); dlg.v_dir.set(out); dlg.v_move.set(True)
            # 취소 → 아무것도 안 바뀜
            answers[:] = [None]; dlg.reports = []
            dlg.run()
            self.assertEqual((os.listdir(out), dlg.reports), ([], []))
            self.assertTrue(os.path.exists(a) and os.path.exists(b))
            # 건너뛰기 → b 만 처리, a 는 그대로
            answers[:] = ["skip"]
            dlg.run()
            self.assertEqual(os.listdir(out), ["b.xlsx"])
            self.assertTrue(os.path.exists(a)); self.assertFalse(os.path.exists(b))
            self.assertEqual([r.path for r in dlg.reports], [a, b])
            self.assertIn("건너뜀", dlg.reports[0].error)
            self.assertIn("건너뜀", ui.report_text(dlg.reports, True))
            # 다시 시도 → 닫혔으니 a 처리
            dlg.files = [a]; answers[:] = ["retry"]
            dlg.run()
            self.assertEqual(sorted(os.listdir(out)), ["a.xlsx", "b.xlsx"])
            self.assertEqual(seen, [[a], [a], [a]])
            # 일반 저장(이동·덮어쓰기 아님)은 확인 안 함
            xlsx(a, [["성명"]]); locked.add(a); dlg.v_move.set(False); answers[:] = []
            dlg.run()
            self.assertEqual(len(seen), 3)
            dlg.destroy()
        finally:
            op.is_locked, ui.ask_locked, ui.messagebox.askyesno, ui.ResultDialog = orig
            root.destroy()


if __name__ == "__main__":
    unittest.main()
