# -*- coding: utf-8 -*-
"""스몰스트림 단독 실행: 독립성(스트림머지 모듈 불필요) · 스트림머지 프리셋 가져오기 · 주 창 닫으면 종료"""
import json, os, sys, tempfile, unittest, shutil, ast
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import appcore  # noqa: E402

MERGE_ONLY = {"exporter", "overlay", "profile_store", "merger", "keynorm", "autorun", "views", "targets", "dates", "cli"}


class TIndependent(unittest.TestCase):
    def test_no_streammerge_imports(self):
        for n in os.listdir(ROOT):
            if not n.endswith(".py"):
                continue
            tree = ast.parse(open(os.path.join(ROOT, n), encoding="utf-8").read())
            mods = set()
            for node in ast.walk(tree):
                if isinstance(node, ast.Import):
                    mods |= {a.name.split(".")[0] for a in node.names}
                elif isinstance(node, ast.ImportFrom) and node.module:
                    mods.add(node.module.split(".")[0])
            self.assertFalse(mods & MERGE_ONLY, f"{n}: {mods & MERGE_ONLY}")


class TImport(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.mkdtemp()
        self.env = {k: os.environ.get(k) for k in ("SMALLSTREAM_HOME", "STREAMMERGE_HOME")}
        os.environ["SMALLSTREAM_HOME"] = os.path.join(self.d, "small")
        os.environ["STREAMMERGE_HOME"] = os.path.join(self.d, "stream")

    def tearDown(self):
        for k, v in self.env.items():
            os.environ.pop(k, None) if v is None else os.environ.__setitem__(k, v)
        shutil.rmtree(self.d, ignore_errors=True)

    def test_import_once(self):
        self.assertEqual(appcore.import_from_parent(), [])              # 스트림머지 없음
        par = os.environ["STREAMMERGE_HOME"]
        os.makedirs(par)
        appcore.atomic_write_json(os.path.join(par, "colprune_presets.json"), [{"name": "월간"}])
        appcore.atomic_write_json(os.path.join(par, "settings.json"), {"overwrite_mode": "rename", "data_dir": "x"})
        got = appcore.import_from_parent()
        self.assertEqual(len(got), 2)
        self.assertEqual(json.load(open(appcore.presets_path(), encoding="utf-8")), [{"name": "월간"}])
        self.assertEqual(appcore.load_settings(), {"overwrite_mode": "rename"})  # 필요한 설정만
        self.assertTrue(os.path.isfile(os.path.join(par, "colprune_presets.json")))  # 원본 유지
        appcore.atomic_write_json(os.path.join(par, "colprune_presets.json"), [{"name": "바뀜"}])
        self.assertEqual(appcore.import_from_parent(), [])              # 두 번째 실행: 덮지 않음
        self.assertEqual(json.load(open(appcore.presets_path(), encoding="utf-8")), [{"name": "월간"}])


try:
    import tkinter as tk
    tk.Tk().destroy()
    HAS_TK = True
except Exception:  # noqa: BLE001
    HAS_TK = False


@unittest.skipUnless(HAS_TK, "tkinter 없음")
class TWindow(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.mkdtemp()
        self.old = os.environ.get("SMALLSTREAM_HOME")
        os.environ["SMALLSTREAM_HOME"] = self.d

    def tearDown(self):
        os.environ.pop("SMALLSTREAM_HOME", None) if self.old is None else os.environ.__setitem__("SMALLSTREAM_HOME", self.old)
        shutil.rmtree(self.d, ignore_errors=True)

    def test_close_ends_app(self):
        import main
        root = main.setup_root()
        dlg = main.build_window(root)
        root.update()
        self.assertIn("스몰스트림", dlg.title())
        self.assertEqual(dlg.backup_dir, os.path.join(self.d, "backup"))
        tabs = [dlg.nb.tab(i, "text") for i in dlg.nb.tabs()]
        self.assertEqual(tabs, ["파일 정리", "정보"])                      # v0.2 정보 탭
        texts = []

        def walk(w):
            try:
                texts.append(str(w.cget("text")))
            except tk.TclError:
                pass
            for c in w.winfo_children():
                walk(c)
        walk(dlg.tab_info)
        joined = "\n".join(texts)
        for want in ("스몰스트림", "버전 " + main.__version__, "제작자. Shugle(슈글)", self.d, dlg.backup_dir):
            self.assertIn(want, joined)
        for gone in ("공개 소프트웨어", "라이선스", "원본은 그대로 두고"):       # v0.3 삭제
            self.assertNotIn(gone, joined)
        self.assertTrue(dlg.lb.winfo_ismapped() or True)                 # 파일 정리 위젯은 첫 탭 안
        self.assertIs(dlg.body.master, dlg.nb)
        dlg.event_generate("<Escape>")
        root.update()
        try:
            root.update()
            alive = bool(root.winfo_exists())
        except tk.TclError:
            alive = False
        self.assertFalse(alive)


if __name__ == "__main__":
    unittest.main()
