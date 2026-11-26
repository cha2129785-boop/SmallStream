# -*- coding: utf-8 -*-
"""v0.6 Excel 표시 형식 해석 + loader 연결 테스트"""
import datetime as dt
import os
import shutil
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from loader import load_table  # noqa: E402
from numfmt import format_value as f  # noqa: E402

CASES = [
    ('"입원기간"', 46333, "입원기간"),
    ('"입원 "0"일"', 12, "입원 12일"),
    ('0"일"', 12.0, "12일"),
    ("#,##0", 1234567, "1,234,567"),
    ("#,##0.00", 1234.5, "1,234.50"),
    ("0000000000", 123456, "0000123456"),
    ("000-0000-0000", 1012345678, "010-1234-5678"),
    ("0.0%", 0.256, "25.6%"),
    ("yyyy-mm-dd", 46333, "2026-11-07"),
    ('yyyy"년" m"월" d"일"', 46333, "2026-11-07"),
    ("h:mm", 0.5, "12:00"),
    ("yyyy-mm-dd hh:mm", 46333.5, "2026-11-07 12:00:00"),
    ("#,##0;[Red]-#,##0", -5, "-5"),
    ("0;(0)", -5, "(5)"),
    ('0;-0;"없음"', 0, "없음"),
    (";;;", 5, ""),
    ("[$-412]yyyy-mm-dd", 46333, "2026-11-07"),
    ("@", "abc", "abc"),
    ('0;0;0;"메모: "@', "x", "메모: x"),
    ("General", 5.0, None),
    ("0.00E+00", 12345, None),
    ("[>=100]0;0", 5, None),
    ("[h]:mm", 1.5, None),
    ("[$₩-412]#,##0", 1000, "₩1,000"),
    ("#,##0", -1234, "-1,234"),
    ("0.##", 3.5, "3.5"),
    ("0.##", 3, "3."),
    ('#,##0"원"', 0, "0원"),
    ('"입원기간"', -3, "입원기간"),
    ("0_);(0)", 5, "5"),
    ("#", 0, ""),
]


class TestFormat(unittest.TestCase):
    def test_cases(self):
        for fmt, v, exp in CASES:
            with self.subTest(fmt=fmt, v=v):
                self.assertEqual(f(v, fmt), exp)

    def test_passthrough(self):
        self.assertIsNone(f(None, "0"))
        self.assertIsNone(f(True, "0"))
        self.assertIsNone(f(dt.datetime(2026, 1, 1), "yyyy-mm-dd"))
        self.assertIsNone(f("text", "#,##0"))
        self.assertIsNone(f(5, None))
        self.assertEqual(f(0, "yyyy-mm-dd", date1904=True), "1904-01-01")


class TestLoaderDisplay(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="표시형식_")

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def test_xlsx_custom_formats(self):
        import openpyxl
        p = os.path.join(self.tmp, "형식.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["번호", "입원기간", "일수", "전화", "금액", "날짜"])
        ws.append([123456, 46333, 12, 1012345678, 1234567, dt.datetime(2026, 11, 7)])
        fmts = ["0000000000", '"입원기간"', '0"일"', "000-0000-0000", '#,##0"원"', "yyyy-mm-dd"]
        for c, fm in zip(ws[2], fmts):
            c.number_format = fm
        wb.save(p)
        t = load_table(p)
        self.assertEqual(t.rows[0], ["0000123456", "입원기간", "12일", "010-1234-5678", "1,234,567원", "2026-11-07"])
        raw = load_table(p, display_format=False)
        self.assertEqual(raw.rows[0][:3], ["123456", "46333", "12"])

    def test_xlsx_formula_cached_value(self):
        """수식 결과(캐시값) + 문자 고정 형식 — 사용자 보고 사례 재현 (46333 노출)"""
        import openpyxl
        import zipfile
        p = os.path.join(self.tmp, "수식.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["번호", "기간"])
        ws.append(["L1", "=TODAY()"])
        ws["B2"].number_format = '"입원기간"'
        wb.save(p)
        # Excel이 저장한 것처럼 수식 캐시값 46333 삽입
        tmp = p + ".tmp"
        with zipfile.ZipFile(p) as zin, zipfile.ZipFile(tmp, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    data = data.replace(b"<f>TODAY()</f><v></v>", b"<f>TODAY()</f><v>46333</v>")
                    data = data.replace(b"<f>TODAY()</f><v />", b"<f>TODAY()</f><v>46333</v>")
                    if b"<v>46333</v>" not in data:
                        data = data.replace(b"<f>TODAY()</f>", b"<f>TODAY()</f><v>46333</v>")
                zout.writestr(item, data)
        os.replace(tmp, p)
        self.assertEqual(load_table(p, display_format=False).rows[0][1], "46333")
        self.assertEqual(load_table(p).rows[0][1], "입원기간")

    def test_xls_formats(self):
        try:
            import xlwt
        except ImportError:
            self.skipTest("xlwt 없음")
        p = os.path.join(self.tmp, "형식.xls")
        wb = xlwt.Workbook(encoding="utf-8")
        ws = wb.add_sheet("s")
        for c, h in enumerate(["번호", "입원기간", "금액"]):
            ws.write(0, c, h)
        ws.write(1, 0, 123456, xlwt.easyxf(num_format_str="0000000000"))
        ws.write(1, 1, 46333, xlwt.easyxf(num_format_str='"입원기간"'))
        ws.write(1, 2, 1234567, xlwt.easyxf(num_format_str="#,##0"))
        wb.save(p)
        self.assertEqual(load_table(p).rows[0], ["0000123456", "입원기간", "1,234,567"])
        self.assertEqual(load_table(p, display_format=False).rows[0], ["123456", "46333", "1234567"])


if __name__ == "__main__":
    unittest.main()
