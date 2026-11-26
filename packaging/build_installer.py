# -*- coding: utf-8 -*-
"""스몰스트림 설치 파일 만들기 v2.0 — build.bat 이 실행 (bat 은 ASCII 만, 단계·한글 메시지는 모두 여기)

  기본 (설치 파일, '설치된 앱'에 등록):
    1) Python 확인  2) 부품 설치  3) EXE 폴더판 (PyInstaller, 중간 파일 %LOCALAPPDATA%\\SmallStream_build)
    4) Inno Setup 찾기 — 버전(6·7·이후)·설치 위치 무관, 못 찾으면 winget → 직접 선택 창 (선택 기억)
    5) 설치 파일 (ISCC) → dist_installer\\SmallStream_Setup_<버전>.exe → 탐색기에서 선택
  --portable : 설치 없이 쓰는 EXE 1개 → dist\\SmallStream.exe (+ zip)
  기록: build.log    (StreamMerge build_installer v5.1 의 Inno Setup 탐색 코드 재사용)
"""
from __future__ import annotations

import datetime as dt
import glob
import json
import os
import re
import shutil
import string
import subprocess
import sys
import tempfile
import traceback
import zipfile

__version__ = "2.0"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOG_PATH = os.path.join(ROOT, "build.log")
SETTINGS = os.path.join(ROOT, "packaging", "build_settings.json")
ISS = os.path.join(ROOT, "installer", "SmallStream.iss")
DIST = os.path.join(ROOT, "dist")
SETUP_DIR = os.path.join(ROOT, "dist_installer")
NAME = "SmallStream"
IS_WIN = sys.platform == "win32"
DOCS = ["사용법.md", "THIRD_PARTY_LICENSES.txt"]
REQUIRED = ["main.py", "SmallStream.spec", os.path.join("installer", "SmallStream.iss"), *DOCS] + [
    os.path.join("assets", n) for n in ("smallstream.ico", "smallstream.png", "smallstream_128.png", "smallstream_32.png")]
_log = None


class BuildError(Exception):
    pass


# ---------- 출력·기록 ----------
def say(msg: str = "") -> None:
    print(msg, flush=True)
    if _log:
        _log.write(msg + "\n")
        _log.flush()


def step(n: int, msg: str, total: int = 5) -> None:
    say(f"[{n}/{total}] {msg}")


def run(cmd: list[str], cwd: str = ROOT, env_add: dict | None = None) -> tuple[int, str]:
    """외부 명령: 출력을 화면·로그에 그대로 (리스트 인자 → 따옴표·한글 경로 문제 없음)"""
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1", **(env_add or {}))
    try:
        p = subprocess.Popen(cmd, cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, env=env,
                             text=True, encoding="utf-8", errors="replace")
    except OSError as e:
        say(f"  실행 불가: {cmd[0]} ({e})")
        return 127, ""
    out = []
    for line in p.stdout:
        line = line.rstrip()
        out.append(line)
        say(line)
    return p.wait(), "\n".join(out)


# ---------- 준비 ----------
def app_version(root: str = ROOT) -> str:
    """main.py 의 __version__ (버전 번호는 한 곳에서만 관리)"""
    with open(os.path.join(root, "main.py"), encoding="utf-8") as f:
        m = re.search(r'^__version__\s*=\s*"([^"]+)"', f.read(), re.M)
    return m.group(1) if m else "0.0"


def ver_tuple(ver: str) -> tuple[int, int, int, int]:
    nums = [int(x) for x in re.findall(r"\d+", ver)][:4]
    return tuple(nums + [0] * (4 - len(nums)))  # type: ignore[return-value]


def version_info_text(ver: str) -> str:
    """PyInstaller 버전 리소스 (탐색기 → 속성 → 자세히)"""
    t = ver_tuple(ver)
    s = ".".join(str(x) for x in t[:3])
    return f"""# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(filevers={t}, prodvers={t}, mask=0x3f, flags=0x0, OS=0x40004,
                    fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([StringTable("041204B0", [
      StringStruct("CompanyName", "Shugle(슈글)"),
      StringStruct("FileDescription", "스몰스트림"),
      StringStruct("FileVersion", "{s}"),
      StringStruct("InternalName", "{NAME}"),
      StringStruct("LegalCopyright", "© Shugle(슈글)"),
      StringStruct("OriginalFilename", "{NAME}.exe"),
      StringStruct("ProductName", "스몰스트림"),
      StringStruct("ProductVersion", "{s}")])]),
    VarFileInfo([VarStruct("Translation", [0x0412, 1200])])
  ]
)
"""


def check_sources(root: str = ROOT) -> None:
    """필요 파일 확인 — 압축 해제 시 한글 이름 깨짐(UTF-8 표시 없는 zip) 감지"""
    missing = [f for f in REQUIRED if not os.path.isfile(os.path.join(root, f))]
    if not missing:
        return
    odd = [n for n in os.listdir(root) if n.endswith(".md") and not n.isascii() and n not in REQUIRED]
    msg = "필요 파일 없음: " + ", ".join(missing)
    if odd:
        msg += f"\n  이름이 깨진 파일 발견: {', '.join(odd)} → 새 zip 으로 다시 압축 해제하거나 '사용법.md' 로 이름 변경"
    raise BuildError(msg)


def make_zip(src: str, zpath: str, arc_root: str, extra: list[str] = ()) -> int:
    """src(폴더 또는 파일) + extra 파일 → zip. zipfile 은 비 ASCII 이름에 UTF-8 표시(0x800) 자동. 반환 항목 수"""
    if os.path.exists(zpath):
        os.remove(zpath)
    n = 0
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        if os.path.isdir(src):
            for d, _, files in os.walk(src):
                for f in files:
                    p = os.path.join(d, f)
                    z.write(p, os.path.join(arc_root, os.path.relpath(p, src)))
                    n += 1
        else:
            z.write(src, os.path.join(arc_root, os.path.basename(src)))
            n += 1
        for p in extra:
            z.write(p, os.path.join(arc_root, os.path.basename(p)))
            n += 1
    bad = [i.filename for i in zipfile.ZipFile(zpath).infolist()
           if not i.filename.isascii() and not i.flag_bits & 0x800]
    if bad:
        raise BuildError(f"zip 한글 이름 표시 누락: {bad}")
    return n


def copy_out(built: str, onefile: bool, ver: str, dist: str = DIST, root: str = ROOT) -> tuple[str, str]:
    """빌드 결과 → dist. 반환 (실행 파일 경로, zip 경로). 기존 결과(우리가 만든 것)는 교체"""
    os.makedirs(dist, exist_ok=True)
    docs = [os.path.join(root, d) for d in DOCS]
    out = os.path.join(dist, NAME)
    exe1 = os.path.join(dist, f"{NAME}.exe")
    # 다른 형식의 이전 결과 제거 → 잘못된 파일을 복사해 가는 일 방지 (v1.1)
    for old in ([out] if onefile else [out, exe1]):
        if os.path.isdir(old):
            shutil.rmtree(old)
        elif os.path.isfile(old):
            os.remove(old)
    for old in os.listdir(dist):
        if old.startswith(f"{NAME}_") and old.endswith(".zip"):
            os.remove(os.path.join(dist, old))
    if onefile:
        shutil.copy2(built, exe1)
        zpath = os.path.join(dist, f"{NAME}_{ver}.zip")
        make_zip(exe1, zpath, NAME, docs)
        return exe1, zpath
    shutil.copytree(built, out)
    for d in docs:
        shutil.copy2(d, out)
    zpath = os.path.join(dist, f"{NAME}_{ver}_folder.zip")
    make_zip(out, zpath, NAME)
    return os.path.join(out, f"{NAME}.exe" if IS_WIN else NAME), zpath



# ---------- Inno Setup (StreamMerge build_installer v5.1) ----------
def file_version(path: str) -> tuple[int, int] | None:
    """EXE 파일 버전 (Windows 버전 리소스) → 실패 시 상위 폴더 이름의 숫자 (예: 'Inno Setup 7')"""
    if IS_WIN:
        try:
            import ctypes
            from ctypes import wintypes
            ver = ctypes.WinDLL("version")
            size = ver.GetFileVersionInfoSizeW(path, None)
            if size:
                buf = ctypes.create_string_buffer(size)
                if ver.GetFileVersionInfoW(path, 0, size, buf):
                    p = ctypes.c_void_p()
                    n = wintypes.UINT()
                    if ver.VerQueryValueW(buf, "\\", ctypes.byref(p), ctypes.byref(n)) and n.value:
                        ms = ctypes.cast(p, ctypes.POINTER(wintypes.DWORD * 4))[0][2]  # dwFileVersionMS
                        major, minor = ms >> 16, ms & 0xFFFF
                        if major:
                            return major, minor
        except Exception:  # noqa: BLE001
            pass
    m = re.search(r"(\d+)(?:\.(\d+))?", os.path.basename(os.path.dirname(path)))
    return (int(m.group(1)), int(m.group(2) or 0)) if m else None


def _registry_candidates() -> list[str]:
    if not IS_WIN:
        return []
    import winreg
    out = []
    roots = [(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
             (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
             (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")]
    for hive, base in roots:
        try:
            with winreg.OpenKey(hive, base) as k:
                for i in range(winreg.QueryInfoKey(k)[0]):
                    try:
                        with winreg.OpenKey(k, winreg.EnumKey(k, i)) as sub:
                            name = _reg_get(sub, "DisplayName")
                            if not name.lower().startswith("inno setup"):
                                continue
                            for v in ("InstallLocation", "DisplayIcon", "UninstallString"):
                                d = _reg_get(sub, v).split(",")[0].strip().strip('"')
                                if d:
                                    d = os.path.dirname(d) if d.lower().endswith(".exe") else d
                                    out.append(os.path.join(d, "ISCC.exe"))
                    except OSError:
                        continue
        except OSError:
            continue
    # .iss 파일 연결 (더블클릭 시 여는 Compil32.exe 와 같은 폴더)
    for hive, base in ((winreg.HKEY_CLASSES_ROOT, r"InnoSetupScriptFile\shell"),
                       (winreg.HKEY_CURRENT_USER, r"Software\Classes\InnoSetupScriptFile\shell")):
        try:
            with winreg.OpenKey(hive, base) as k:
                for i in range(winreg.QueryInfoKey(k)[0]):
                    try:
                        with winreg.OpenKey(k, winreg.EnumKey(k, i) + r"\command") as c:
                            m = re.search(r'"([^"]+\.exe)"', _reg_get(c, ""), re.I)
                            if m:
                                out.append(os.path.join(os.path.dirname(m.group(1)), "ISCC.exe"))
                    except OSError:
                        continue
        except OSError:
            continue
    return out


def _reg_get(key, name: str) -> str:
    import winreg
    try:
        return str(winreg.QueryValueEx(key, name)[0] or "")
    except OSError:
        return ""


def _folder_roots() -> list[str]:
    roots = []
    la = os.environ.get("LOCALAPPDATA")
    if la:
        roots.append(os.path.join(la, "Programs"))
    if IS_WIN:
        for d in string.ascii_uppercase:
            drive = f"{d}:\\"
            if os.path.exists(drive):
                roots += [os.path.join(drive, "Program Files"), os.path.join(drive, "Program Files (x86)"), drive]
    return roots


def iscc_candidates(extra_roots: list[str] | None = None) -> list[str]:
    cands: list[str] = []
    try:
        with open(SETTINGS, encoding="utf-8") as f:
            saved = json.load(f).get("iscc")
        if saved:
            cands.append(saved)
    except (OSError, ValueError):
        pass
    w = shutil.which("ISCC") or shutil.which("ISCC.exe")
    if w:
        cands.append(w)
    cands += _registry_candidates()
    for r in (extra_roots if extra_roots is not None else _folder_roots()):
        cands += glob.glob(os.path.join(glob.escape(r), "Inno Setup*", "ISCC.exe"))
        cands += glob.glob(os.path.join(glob.escape(r), "Inno Setup*", "iscc.exe"))
    seen, out = set(), []
    for c in cands:
        k = os.path.normcase(os.path.abspath(c))
        if k not in seen and os.path.isfile(c):
            seen.add(k)
            out.append(c)
    return out


def choose_iscc(cands: list[str]) -> tuple[str, tuple[int, int]] | None:
    """가장 높은 버전 (6 이상). 버전을 못 읽은 것만 있으면 그것을 7로 간주"""
    best, best_v, unknown = None, None, None
    for c in cands:
        v = file_version(c)
        say(f"  후보: {c}  (버전 {'.'.join(map(str, v)) if v else '알 수 없음'})")
        if v is None:
            unknown = unknown or c
        elif v[0] >= 6 and (best_v is None or v > best_v):
            best, best_v = c, v
    if best:
        return best, best_v
    if unknown:
        return unknown, (7, 0)
    return None


def pick_iscc_dialog() -> str | None:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        p = filedialog.askopenfilename(parent=root, title="Inno Setup 의 ISCC.exe 선택 (Inno Setup 설치 폴더 안)",
                                       initialdir=os.environ.get("ProgramFiles", "C:\\"),
                                       filetypes=[("ISCC.exe", "ISCC.exe"), ("실행 파일", "*.exe")])
        root.destroy()
        return p or None
    except Exception as e:  # noqa: BLE001
        say(f"  선택 창 열기 실패: {e}")
        return None


def find_iscc() -> tuple[str, tuple[int, int]]:
    found = choose_iscc(iscc_candidates())
    if not found and shutil.which("winget"):
        say("  Inno Setup 없음 → 최신 버전 자동 설치 (권한 확인 창이 뜨면 '예')")
        run(["winget", "install", "--id", "JRSoftware.InnoSetup", "-e", "--silent",
             "--accept-package-agreements", "--accept-source-agreements"])
        found = choose_iscc(iscc_candidates())
    if not found:
        say("  자동으로 못 찾음 → ISCC.exe 위치를 직접 선택하세요 (창이 안 보이면 작업 표시줄 확인)")
        p = pick_iscc_dialog()
        if p and os.path.isfile(p):
            found = (p, file_version(p) or (7, 0))
    if not found:
        raise BuildError("Inno Setup 을 찾지 못함 → https://jrsoftware.org/isdl.php 에서 설치 후 build.bat 다시 실행")
    os.makedirs(os.path.dirname(SETTINGS), exist_ok=True)
    with open(SETTINGS, "w", encoding="utf-8") as f:
        json.dump({"iscc": found[0]}, f, ensure_ascii=False, indent=1)  # 다음부터 바로 사용
    return found


# ---------- 설치 스크립트 준비 ----------
def prepare_iss(text: str, iscc: str, ver: tuple[int, int]) -> tuple[str, list[str]]:
    notes = []
    lang = os.path.join(os.path.dirname(iscc), "Languages", "Korean.isl")
    if not os.path.isfile(lang):
        text = "\n".join(l for l in text.splitlines() if not re.match(r'\s*Name:\s*"korean"', l))
        notes.append("Korean.isl 없음 → 설치 화면 영어로 진행")
    if ver < (6, 3):
        text = text.replace("x64compatible", "x64")
        notes.append("Inno Setup 6.3 미만 → x64 설정 사용")
    return text, notes




def clean_old_portable(dist: str = DIST) -> list[str]:
    """설치 파일 빌드 때: 예전 휴대용 결과(dist\\SmallStream.exe·폴더·zip) 삭제 → 잘못된 파일 배포 방지. 반환 삭제 목록"""
    gone = []
    if not os.path.isdir(dist):
        return gone
    for n in os.listdir(dist):
        p = os.path.join(dist, n)
        if n in (f"{NAME}.exe", NAME) or (n.startswith(f"{NAME}_") and n.endswith(".zip")):
            shutil.rmtree(p) if os.path.isdir(p) else os.remove(p)
            gone.append(n)
    return gone


def build_exe(ver: str, onefile: bool) -> str:
    """PyInstaller. 반환: onefile → EXE 경로 / 폴더판 → 폴더 경로"""
    if IS_WIN:
        subprocess.run(["taskkill", "/IM", f"{NAME}.exe", "/F"], capture_output=True)
    work = os.path.join(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir(), f"{NAME}_build")
    os.makedirs(work, exist_ok=True)
    vinfo = os.path.join(work, "version_info.txt")
    with open(vinfo, "w", encoding="utf-8") as f:
        f.write(version_info_text(ver))
    pdist = os.path.join(work, "dist")
    code, _ = run([sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--log-level", "WARN",
                   "--workpath", os.path.join(work, "build"), "--distpath", pdist, f"{NAME}.spec"],
                  env_add={"SMALLSTREAM_ONEFILE": "1" if onefile else "0", "SMALLSTREAM_VERINFO": vinfo})
    exe_name = f"{NAME}.exe" if IS_WIN else NAME
    built = os.path.join(pdist, exe_name) if onefile else os.path.join(pdist, NAME)
    if code or not os.path.isfile(built if onefile else os.path.join(built, exe_name)):
        raise BuildError("EXE 만들기 실패 (위 PyInstaller 메시지 확인)")
    say(f"  -> {built}")
    return built


# ---------- 단계 ----------
def main(argv: list[str]) -> int:
    global _log
    portable = "--portable" in argv
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    _log = open(LOG_PATH, "w", encoding="utf-8-sig")
    total = 4 if portable else 5
    say(f"스몰스트림 {'휴대용 EXE' if portable else '설치 파일'} 만들기 v{__version__} — {dt.datetime.now():%Y-%m-%d %H:%M:%S}")
    try:
        ver = app_version()
        step(1, "Python 확인", total)
        say(f"  {sys.version.split()[0]}  {sys.executable}")
        if sys.version_info < (3, 11):
            raise BuildError("Python 3.11 이상 필요 (3.12 권장)")

        step(2, "부품 설치 (openpyxl, xlrd, pyinstaller)", total)
        code, _ = run([sys.executable, "-m", "pip", "install", "-q", "--disable-pip-version-check",
                       "-r", "requirements.txt", "pyinstaller>=6.0"])
        if code:
            raise BuildError("부품 설치 실패 (위 pip 메시지 확인 — 인터넷 연결·회사 방화벽)")
        check_sources()

        step(3, f"EXE 만들기 v{ver} (1~3분, INFO·WARNING 메시지는 정상)", total)
        built = build_exe(ver, onefile=portable)

        if portable:
            step(4, "dist 폴더로 복사 · zip", total)
            try:
                exe, zpath = copy_out(built, True, ver)
            except PermissionError as e:
                raise BuildError(f"dist 안 파일이 사용 중 — 스몰스트림을 닫고 다시 실행 ({e})") from e
            say("")
            say(f"[완료] 휴대용 EXE: {exe}  (설치 없이 실행, '설치된 앱'에는 안 나옴)")
            say(f"       zip: {zpath}")
            if IS_WIN:
                subprocess.Popen(["explorer", "/select,", exe])
            return 0

        step(4, "Inno Setup 찾기", total)
        iscc, iver = find_iscc()
        say(f"  사용: Inno Setup {iver[0]}.{iver[1]}  {iscc}")

        step(5, "설치 파일 만들기", total)
        with open(ISS, encoding="utf-8-sig") as f:
            text, notes = prepare_iss(f.read(), iscc, iver)
        for n in notes:
            say(f"  {n}")
        tmp = os.path.join(ROOT, "installer", "_build.iss")  # 같은 폴더 → 상대 경로 유지
        with open(tmp, "w", encoding="utf-8-sig", newline="\r\n") as f:
            f.write(text)
        try:
            code, _ = run([iscc, f"/DDistDir={built}", f"/DMyAppVersion={ver}", tmp])
        finally:
            try:
                os.remove(tmp)
            except OSError:
                pass
        setup = os.path.join(SETUP_DIR, f"{NAME}_Setup_{ver}.exe")
        if code or not os.path.isfile(setup):
            raise BuildError("설치 파일 만들기 실패 (위 ISCC 메시지 확인)")
        gone = clean_old_portable()
        if gone:
            say(f"  예전 휴대용 결과 삭제: {', '.join(gone)}")
        say("")
        say(f"[완료] {setup}")
        say("다른 PC에 이 파일 하나 복사 → 실행하면 설치 · '설치된 앱'에 등록")
        say("삭제: Windows 설정 → 앱 → 설치된 앱 → 스몰스트림")
        if IS_WIN:
            subprocess.Popen(["explorer", "/select,", setup])
        return 0
    except BuildError as e:
        say("")
        say(f"[실패] {e}")
        say("build.log 파일 내용을 보내주세요.")
        return 1
    except Exception:  # noqa: BLE001
        say("")
        say("[실패] 예상하지 못한 오류")
        say(traceback.format_exc())
        say("build.log 파일 내용을 보내주세요.")
        return 1
    finally:
        _log.close()


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
