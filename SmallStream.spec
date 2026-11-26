# -*- mode: python ; coding: utf-8 -*-
# 스몰스트림 PyInstaller 빌드 설정 — packaging\build_installer.py 가 실행 (build.bat)
# SMALLSTREAM_ONEFILE=0 → onedir (설치 파일용, 기본) / 1 → EXE 1개 (build.bat --portable)
# SMALLSTREAM_VERINFO: build_installer.py 가 main.__version__ 으로 만든 버전 정보 파일
import os

ONEFILE = os.environ.get("SMALLSTREAM_ONEFILE") == "1"
VERINFO = os.environ.get("SMALLSTREAM_VERINFO") or None
ASSETS = ["smallstream.png", "smallstream_128.png", "smallstream_32.png", "smallstream.ico"]

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=[(os.path.join("assets", n), "assets") for n in ASSETS],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=["pytest", "numpy", "pandas", "matplotlib", "PIL", "xlwt", "tksheet"],
    noarchive=False,
)
pyz = PYZ(a.pure)
common = dict(name="SmallStream", debug=False, strip=False, upx=False, console=False,
              icon=os.path.join("assets", "smallstream.ico"), version=VERINFO)
if ONEFILE:
    exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], runtime_tmpdir=None, **common)
else:
    exe = EXE(pyz, a.scripts, [], exclude_binaries=True, **common)
    coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name="SmallStream")
