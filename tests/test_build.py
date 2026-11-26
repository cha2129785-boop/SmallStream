# -*- coding: utf-8 -*-
import os, sys, json, tempfile, unittest, zipfile, shutil
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "packaging"))
import build_installer as b  # noqa: E402


def touch(p):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    open(p, "w").close()
    return p

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class T(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.mkdtemp()
        b.SETTINGS = os.path.join(self.d, "s.json")
        b._log = None

    def tearDown(self):
        shutil.rmtree(self.d, ignore_errors=True)

    def test_version_from_main(self):
        import re
        v = b.app_version()
        self.assertRegex(v, r"^\d+\.\d+")
        self.assertEqual(b.ver_tuple("0.1"), (0, 1, 0, 0))
        t = b.version_info_text("0.1")
        self.assertIn("(0, 1, 0, 0)", t)
        self.assertIn('"0.1.0"', t)
        self.assertTrue(re.search(r"VSVersionInfo", t))

    def test_sources_ok_and_broken_name(self):
        b.check_sources(ROOT)
        os.makedirs(os.path.join(self.d, "assets"))
        open(os.path.join(self.d, "main.py"), "w").close()
        open(os.path.join(self.d, "╗τ┐δ╣².md"), "w").close()
        with self.assertRaises(b.BuildError) as c:
            b.check_sources(self.d)
        self.assertIn("이름이 깨진", str(c.exception))

    def test_copy_out_onedir_zip_utf8(self):
        built = os.path.join(self.d, "built", "SmallStream")
        os.makedirs(os.path.join(built, "_internal"))
        exe = "SmallStream.exe" if b.IS_WIN else "SmallStream"
        open(os.path.join(built, exe), "w").close()
        open(os.path.join(built, "_internal", "x.dll"), "w").close()
        dist = os.path.join(self.d, "dist")
        os.makedirs(os.path.join(dist, "SmallStream"))
        open(os.path.join(dist, "SmallStream", "old.txt"), "w").close()   # 이전 결과 교체
        e, z = b.copy_out(built, False, "0.1", dist, ROOT)
        self.assertTrue(os.path.isfile(e))
        self.assertFalse(os.path.exists(os.path.join(dist, "SmallStream", "old.txt")))
        names = {i.filename: i for i in zipfile.ZipFile(z).infolist()}
        self.assertIn("SmallStream/사용법.md", names)
        self.assertTrue(names["SmallStream/사용법.md"].flag_bits & 0x800)
        self.assertIn("SmallStream/_internal/x.dll", names)
        self.assertTrue(z.endswith("_0.1_folder.zip"))

    def test_copy_out_onefile(self):
        built = os.path.join(self.d, "SmallStream.exe")
        open(built, "w").close()
        dist = os.path.join(self.d, "dist")
        os.makedirs(os.path.join(dist, "SmallStream", "_internal"))          # 이전 폴더판 결과
        open(os.path.join(dist, "SmallStream_0.0_folder.zip"), "w").close()
        e, z = b.copy_out(built, True, "0.1", dist, ROOT)
        self.assertFalse(os.path.exists(os.path.join(dist, "SmallStream")))  # v1.1: 헷갈리는 이전 결과 제거
        self.assertEqual(sorted(os.listdir(dist)), ["SmallStream.exe", "SmallStream_0.1.zip"])
        self.assertTrue(e.endswith("SmallStream.exe"))
        self.assertEqual(sorted(zipfile.ZipFile(z).namelist()),
                         ["SmallStream/SmallStream.exe", "SmallStream/THIRD_PARTY_LICENSES.txt", "SmallStream/사용법.md"])



class TInstaller(unittest.TestCase):
    """v0.4: 기본 = 설치 파일 (Inno Setup). Inno 탐색 테스트는 스트림머지 test_build 이식"""

    def setUp(self):
        self.d = tempfile.mkdtemp()
        b.SETTINGS = os.path.join(self.d, "s.json")
        b._log = None

    def tearDown(self):
        shutil.rmtree(self.d, ignore_errors=True)

    def test_default_is_installer(self):
        src = open(os.path.join(ROOT, "packaging", "build_installer.py"), encoding="utf-8").read()
        self.assertIn('portable = "--portable" in argv', src)
        self.assertIn("build_exe(ver, onefile=portable)", src)

    def test_iss_own_appid(self):
        t = open(b.ISS, encoding="utf-8-sig").read()
        self.assertIn("AppId={{4B1F9332-CCFC-47E6-AED4-ED68364B4325}", t)
        self.assertNotIn("6E0C4A71", t)                       # 스트림머지 AppId 아님
        for want in ('MyAppExeName "SmallStream.exe"', "smallstream.ico", "SmallStream_Setup_", "PrivilegesRequired=lowest",
                     "{userappdata}\\{#MyAppId}", "MB_DEFBUTTON2"):
            self.assertIn(want, t)

    def test_clean_old_portable(self):
        dist = os.path.join(self.d, "dist")
        touch(os.path.join(dist, "SmallStream.exe"))
        touch(os.path.join(dist, "SmallStream", "_internal", "python312.dll"))
        touch(os.path.join(dist, "SmallStream_0.3.1.zip"))
        touch(os.path.join(dist, "keep.txt"))
        self.assertEqual(sorted(b.clean_old_portable(dist)), ["SmallStream", "SmallStream.exe", "SmallStream_0.3.1.zip"])
        self.assertEqual(os.listdir(dist), ["keep.txt"])
        self.assertEqual(b.clean_old_portable(os.path.join(self.d, "none")), [])

    def test_any_version_folder(self):
        for v in ("Inno Setup 5", "Inno Setup 6", "Inno Setup 7"):
            touch(os.path.join(self.d, "PF", v, "ISCC.exe"))
        c = b.iscc_candidates([os.path.join(self.d, "PF")])
        self.assertEqual(len(c), 3)
        path, ver = b.choose_iscc(c)
        self.assertIn("Inno Setup 7", path)
        self.assertEqual(ver, (7, 0))

    def test_unknown_version_treated_7(self):
        p = touch(os.path.join(self.d, "tools", "my", "ISCC.exe"))
        self.assertEqual(b.choose_iscc([p]), (p, (7, 0)))

    def test_v5_only_rejected(self):
        p = touch(os.path.join(self.d, "PF", "Inno Setup 5", "ISCC.exe"))
        self.assertIsNone(b.choose_iscc([p]))

    def test_saved_choice_first_and_dedup(self):
        p = touch(os.path.join(self.d, "X", "Inno Setup 7", "ISCC.exe"))
        json.dump({"iscc": p}, open(b.SETTINGS, "w"))
        c = b.iscc_candidates([os.path.join(self.d, "X")])
        self.assertEqual(c, [p])

    def test_brackets_in_path(self):
        root = os.path.join(self.d, "P [x86]")
        touch(os.path.join(root, "Inno Setup 7", "ISCC.exe"))
        self.assertEqual(len(b.iscc_candidates([root])), 1)

    def test_prepare_iss(self):
        txt = 'A\nName: "korean"; MessagesFile: "compiler:Languages\\Korean.isl"\nArch=x64compatible\n'
        iscc = touch(os.path.join(self.d, "IS7", "ISCC.exe"))
        t, n = b.prepare_iss(txt, iscc, (7, 0))
        self.assertNotIn("korean", t)
        self.assertIn("x64compatible", t)
        touch(os.path.join(self.d, "IS7", "Languages", "Korean.isl"))
        t, n = b.prepare_iss(txt, iscc, (6, 2))
        self.assertIn("korean", t)
        self.assertNotIn("x64compatible", t)

    def test_real_iss_korean_line_matches(self):
        t = open(b.ISS, encoding="utf-8-sig").read()
        iscc = os.path.join(self.d, "none", "ISCC.exe")
        out, _ = b.prepare_iss(t, iscc, (7, 0))
        self.assertNotIn("MessagesFile", out)
        self.assertIn("#ifndef MyAppVersion", out)


if __name__ == "__main__":
    unittest.main()
