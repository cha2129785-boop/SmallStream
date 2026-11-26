# -*- coding: utf-8 -*-
"""appcore v0.1 — 스몰스트림 공통: 데이터 폴더·설정·JSON 저장·Excel 글꼴/열 너비 (스트림머지 exporter·overlay·profile_store 에서 필요한 부분만)"""
from __future__ import annotations

import json
import os
import re
import shutil
import unicodedata

APP_NAME = "스몰스트림"
APP_ID = "SmallStream"
PARENT_ID = "StreamMerge"          # 프리셋 가져올 원래 프로그램
PRESETS_FILE = "colprune_presets.json"

FONT_NAME = "맑은 고딕"
FONT_SIZE = 11
MIN_W, MAX_W = 6, 60


# ---------- 위치 ----------
def _appdata() -> str:
    return os.environ.get("APPDATA") or os.path.expanduser("~")


def base_dir() -> str:
    """%APPDATA%\\SmallStream — 프리셋·설정·backup·error.log (환경변수 SMALLSTREAM_HOME 로 바꿈, 테스트용)"""
    return os.environ.get("SMALLSTREAM_HOME") or os.path.join(_appdata(), APP_ID)


def presets_path() -> str:
    return os.path.join(base_dir(), PRESETS_FILE)


def settings_path() -> str:
    return os.path.join(base_dir(), "settings.json")


# ---------- 저장 ----------
def atomic_write_json(path: str, data) -> None:
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, path)


def load_settings() -> dict:
    try:
        with open(settings_path(), encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except (OSError, ValueError):
        return {}


def save_settings(d: dict) -> None:
    atomic_write_json(settings_path(), d)


# ---------- 스트림머지에서 가져오기 (첫 실행 1회) ----------
def parent_dir() -> str:
    return os.environ.get("STREAMMERGE_HOME") or os.path.join(_appdata(), PARENT_ID)


def import_from_parent() -> list[str]:
    """스몰스트림 프리셋 파일이 없고 스트림머지 프리셋이 있으면 복사 (원본 유지).
    같은 이름 파일 처리(overwrite_mode)도 아직 없으면 가져옴. 반환: 가져온 항목 설명"""
    done: list[str] = []
    src = os.path.join(parent_dir(), PRESETS_FILE)
    dst = presets_path()
    if os.path.isfile(src) and not os.path.exists(dst):
        os.makedirs(base_dir(), exist_ok=True)
        shutil.copy2(src, dst)
        done.append(f"파일 정리 프리셋: {src}")
    s = load_settings()
    if "overwrite_mode" not in s:
        try:
            with open(os.path.join(parent_dir(), "settings.json"), encoding="utf-8") as f:
                ps = json.load(f)
        except (OSError, ValueError):
            ps = {}
        if isinstance(ps, dict) and "overwrite_mode" in ps:
            s["overwrite_mode"] = ps["overwrite_mode"]
            save_settings(s)
            done.append("같은 이름 파일 처리 설정")
    return done


# ---------- Excel 표시 ----------
def text_width(s: str) -> int:
    w = 0
    for ch in str(s):
        w += 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
    return w


def sheet_title(name: str, used: set[str]) -> str:
    t = re.sub(r"[\[\]:*?/\\]", "_", name).strip("'") or "시트"
    t = t[:31]
    base, n = t, 1
    while t in used:
        n += 1
        suf = f"_{n}"
        t = base[:31 - len(suf)] + suf
    used.add(t)
    return t
