# -*- coding: utf-8 -*-
"""작업 표시줄 묶음: 설치 EXE(frozen)는 AppUserModelID 미설정, 개발용(python)은 설정"""
import os, sys, unittest
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    import main
except Exception:  # noqa: BLE001  tkinter 없는 환경
    main = None


@unittest.skipIf(main is None, "main import 불가")
class T(unittest.TestCase):
    def tearDown(self):
        if hasattr(sys, "frozen"):
            del sys.frozen

    def test_dev_sets_id(self):
        self.assertTrue(main.use_explicit_app_id())

    def test_frozen_exe_no_id(self):
        sys.frozen = True
        self.assertFalse(main.use_explicit_app_id())

    def test_call_guarded(self):
        src = open(main.__file__, encoding="utf-8").read()
        i = src.index("if use_explicit_app_id():")
        self.assertLess(i, src.index("SetCurrentProcessExplicitAppUserModelID"))
        self.assertEqual(src.count("SetCurrentProcessExplicitAppUserModelID"), 1)


if __name__ == "__main__":
    unittest.main()
