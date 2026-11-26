# -*- coding: utf-8 -*-
"""가상 데이터 생성기 v0.1 — 실존 인물 아님. 형식이 섞인 5개 파일 생성

  python tools/make_samples.py [출력폴더]   (기본: samples/)
"""
from __future__ import annotations

import datetime as dt
import os
import sys

__version__ = "0.1"

N = 30
_SUR = "김이박최정강조윤장임한오서신권황안송류홍"
_GIV = ["순자", "영자", "정숙", "옥순", "말순", "영수", "철수", "복순", "춘자", "갑돌",
        "정희", "명자", "길동", "순임", "덕수"]


def key(i: int) -> str:
    return f"L{i:010d}"


def name(i: int) -> str:
    return _SUR[i % len(_SUR)] + _GIV[(i * 7) % len(_GIV)]


def birth(i: int) -> dt.datetime:
    return dt.datetime(1928 + i % 30, 1 + i % 12, 1 + i % 27)


def renew(i: int) -> dt.datetime:
    return dt.datetime(2026, 11, 7) + dt.timedelta(days=i * 13 - 60)


def make_basic(path):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "기본"
    ws.append(["인정번호", "성명", "생년월일", "등급", "갱신예정일"])
    for i in range(1, N + 1):
        ws.append([key(i), name(i), birth(i), f"{1 + i % 5}등급", renew(i)])
    for c in ("C", "E"):
        for cell in ws[c][1:]:
            cell.number_format = "yyyy-mm-dd"
    wb.save(path)


def make_guardian_html(path):
    rows = []
    for i in range(1, 26):
        k = key(i)
        if i == 4:
            k = k.lower()
        elif i % 3 == 0:
            k = f"{k[:1]}-{k[1:6]}-{k[6:]}"
        rows.append((k, name(i), f"보호{name(i + 3)}", "자녀", f"010-{1000 + i:04d}-{2000 + i:04d}"))
        if i == 3:
            rows.append((k, name(i), f"보호{name(i + 9)}", "배우자", f"010-{3000 + i:04d}-0000"))
        if i == 5:
            rows.append(rows[-1])
    html = ["<html><head><meta http-equiv='Content-Type' content='text/html; charset=euc-kr'></head><body>",
            "<table border=1>",
            "<tr><td colspan=5>보호자 현황 (가상 데이터)</td></tr>",
            "<tr><th>수급자번호</th><th>어르신</th><th>보호자명</th><th>관계</th><th>연락처</th></tr>"]
    for r in rows:
        html.append("<tr>" + "".join(f"<td style='mso-number-format:\"@\"'>{v}</td>" for v in r) + "</tr>")
    html.append("</table></body></html>")
    with open(path, "w", encoding="cp949") as f:
        f.write("\n".join(html))


def make_hospital_csv(path):
    lines = ["장기요양번호,입원일,퇴원일,병원"]
    for i in (2, 7, 11, 15, 20):
        lines.append(f" L {i:010d} ,2026.10.0{i % 9 + 1},26-10-2{i % 9},가상병원{i}")
    lines.append(f"{key(99)},2026.09.01,,가상병원99")
    lines.append(f"{key(21)},2026.13.01,,가상병원21")
    lines.append(",2026.09.09,,키없음병원")
    lines.append("1.23E+10,2026.09.10,,지수표기병원")
    with open(path, "w", encoding="cp949", newline="") as f:
        f.write("\r\n".join(lines) + "\r\n")


def make_renew(path):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "갱신"
    ws.append(["갱신 관리표 (가상)"])
    ws.append([])
    ws.append(["인정번호", "성명", "갱신예정일", "갱신완료일"])
    full = str.maketrans("L0123456789", "Ｌ０１２３４５６７８９")
    for i in range(1, 21):
        nm = name(i) + ("미" if i == 2 else "")
        rd = renew(i) + (dt.timedelta(days=30) if i == 1 else dt.timedelta(0))
        done = (renew(i) - dt.timedelta(days=700)).strftime("%Y%m%d")
        ws.append([key(i).translate(full) if i % 2 else key(i), nm, rd, done])
    wb.save(path)


def make_service_xls(path) -> bool:
    try:
        import xlwt
    except ImportError:
        return False
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("서비스")
    dstyle = xlwt.easyxf(num_format_str="YYYY-MM-DD")
    for c, h in enumerate(["인정번호", "서비스", "이용시작일"]):
        ws.write(0, c, h)
    r = 1
    for i in range(1, 11):
        ws.write(r, 0, key(i))
        ws.write(r, 1, "주간보호" if i % 2 else "방문요양")
        ws.write(r, 2, dt.datetime(2025, i, 1), dstyle)
        r += 1
    ws.write(r, 0, 1.23e10)
    ws.write(r, 1, "지수표기")
    wb.save(path)
    return True


def make_all(out: str) -> dict[str, str]:
    os.makedirs(out, exist_ok=True)
    p = {
        "basic": os.path.join(out, "기본정보.xlsx"),
        "guardian": os.path.join(out, "보호자정보.xls"),
        "hospital": os.path.join(out, "입원정보.csv"),
        "renew": os.path.join(out, "갱신정보.xlsx"),
        "service": os.path.join(out, "서비스정보.xls"),
    }
    make_basic(p["basic"])
    make_guardian_html(p["guardian"])
    make_hospital_csv(p["hospital"])
    make_renew(p["renew"])
    if not make_service_xls(p["service"]):
        p.pop("service")
    return p


if __name__ == "__main__":
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except AttributeError:
            pass
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "samples")
    paths = make_all(out)
    for k, v in paths.items():
        print(f"{k}: {v}")
    if "service" not in paths:
        print("xlwt 미설치 → 서비스정보.xls(BIFF) 생략", file=sys.stderr)
