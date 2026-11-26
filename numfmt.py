# -*- coding: utf-8 -*-
"""numfmt v0.6 — Excel 사용자 지정 표시 형식 → 화면에 보이는 문자열

지원: 문자 고정("입원기간"), 숫자+문자(0"일"), 자릿수·천 단위·소수·%, 0 채우기(0000000000),
      자리 패턴(000-0000-0000), 양수;음수;0;텍스트 섹션, [색]·[$-412]·[$₩-412], @, 날짜·시간
미지원 → None 반환 (호출부가 원래 값 사용): 조건 [>=100], 지수 E+, 분수 ?/?, 경과 시간 [h]
날짜 형식은 기능(D-day·정렬) 유지를 위해 YYYY-MM-DD(시간 있으면 HH:MM:SS 추가)로 통일
"""
from __future__ import annotations

import datetime as dt
import re

__version__ = "0.6"

_EPOCH_1900 = dt.datetime(1899, 12, 30)
_EPOCH_1904 = dt.datetime(1904, 1, 1)
_COLOR = re.compile(r"\[(black|blue|cyan|green|magenta|red|white|yellow|color\s*\d+|검정|파랑|녹청|녹색|자홍|빨강|흰색|노랑)\]", re.I)
_LOCALE = re.compile(r"\[\$([^\]-]*)(-[0-9A-Fa-f]+)?\]")


def _sections(fmt: str) -> list[str]:
    out, cur, i, q = [], [], 0, False
    while i < len(fmt):
        ch = fmt[i]
        if ch == '"':
            q = not q
        elif ch == "\\" and not q and i + 1 < len(fmt):
            cur.append(fmt[i:i + 2])
            i += 2
            continue
        if ch == ";" and not q:
            out.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
        i += 1
    out.append("".join(cur))
    return out


def _tokens(sec: str) -> list[tuple[str, str]]:
    """[(종류, 값)] 종류: lit(문자) / ph(자리표시 0#?) / dot / comma / pct / at / raw(그 외 서식 문자)"""
    sec = _COLOR.sub("", sec)
    sec = _LOCALE.sub(lambda m: f'"{m.group(1)}"' if m.group(1) else "", sec)
    out, i = [], 0
    while i < len(sec):
        ch = sec[i]
        if ch == '"':
            j = sec.find('"', i + 1)
            j = len(sec) if j < 0 else j
            out.append(("lit", sec[i + 1:j]))
            i = j + 1
            continue
        if ch == "\\" and i + 1 < len(sec):
            out.append(("lit", sec[i + 1]))
            i += 2
            continue
        if ch in "_*" and i + 1 < len(sec):  # 정렬용 공백·채우기 → 무시
            i += 2
            continue
        if ch in "0#?":
            out.append(("ph", ch))
        elif ch == ".":
            out.append(("dot", ch))
        elif ch == ",":
            out.append(("comma", ch))
        elif ch == "%":
            out.append(("pct", ch))
        elif ch == "@":
            out.append(("at", ch))
        elif ch == "[":
            j = sec.find("]", i)
            out.append(("raw", sec[i:j + 1] if j > 0 else sec[i:]))
            i = (j + 1) if j > 0 else len(sec)
            continue
        elif ch in "$-+/():!^&'~{}<>=" or ch == " " or ord(ch) > 127:
            out.append(("lit", ch))
        else:
            out.append(("raw", ch))
        i += 1
    return out


def _is_date(tokens) -> bool:
    raw = "".join(v for k, v in tokens if k == "raw").lower()
    return bool(re.search(r"[ymdhs]|am/pm|a/p|aaa", raw))


def _date_text(value: float, tokens, date1904: bool) -> str | None:
    raw = "".join(v for k, v in tokens if k == "raw").lower()
    if "[" in raw:  # 경과 시간 [h] 등
        return None
    try:
        d = (_EPOCH_1904 if date1904 else _EPOCH_1900) + dt.timedelta(days=float(value))
    except (OverflowError, ValueError):
        return None
    has_date = bool(re.search(r"[yd]", raw)) or ("m" in raw and not re.search(r"[hs]", raw))
    has_time = bool(re.search(r"[hs]", raw))
    if has_date and has_time:
        return d.strftime("%Y-%m-%d %H:%M:%S")
    if has_time:
        return d.strftime("%H:%M:%S" if "s" in raw else "%H:%M")
    return d.strftime("%Y-%m-%d")


def _fill_int(digits: str, pattern: list[tuple[str, str]]) -> str:
    """정수 자리 채우기: 오른쪽부터 자리표시에 숫자 배치, 남는 숫자는 맨 왼쪽에"""
    out: list[str] = []
    di = len(digits)
    ph_left = [sum(1 for k2, v2 in pattern[:i] if k2 == "ph" and v2 == "0") for i in range(len(pattern))]
    for i in range(len(pattern) - 1, -1, -1):
        k, v = pattern[i]
        if k == "ph":
            if di > 0:
                di -= 1
                out.append(digits[di])
            elif v == "0":
                out.append("0")
            elif v == "?":
                out.append(" ")
        else:  # 자리 사이 문자 (000-0000)
            if di > 0 or ph_left[i] > 0:
                out.append(v)
    s = "".join(reversed(out))
    return digits[:di] + s if di > 0 else s


def _number_text(value: float, tokens) -> str | None:
    kinds = [k for k, _ in tokens]
    if "raw" in kinds:  # E+, 분수, 조건 등
        return None
    phs = [i for i, k in enumerate(kinds) if k == "ph"]
    pct = kinds.count("pct")
    if not phs:
        return "".join(v for k, v in tokens if k in ("lit", "pct"))
    first, last = phs[0], phs[-1]
    dot = next((i for i in range(first, last + 1) if kinds[i] == "dot"), None)
    if dot is None and first > 0 and kinds[first - 1] == "dot":  # .00 형태
        dot, first = first - 1, first - 1
    run = tokens[first:last + 1]
    int_part = run[:dot - first] if dot is not None else run
    frac_part = run[dot - first + 1:] if dot is not None else []
    grouping = any(k == "comma" for k, _ in int_part)
    int_pat = [(k, v) for k, v in int_part if k != "comma"]
    decimals = sum(1 for k, _ in frac_part if k == "ph")
    trailing_comma = last + 1 < len(tokens) and tokens[last + 1][0] == "comma"
    if trailing_comma:
        return None
    v = abs(value) * (100 ** pct)
    text = f"{v:.{decimals}f}"
    ip, _, fp = text.partition(".")
    if ip == "0" and not any(k == "ph" and x == "0" for k, x in int_pat):
        ip = ""
    if grouping and all(k == "ph" for k, _ in int_pat):
        body = f"{int(ip or 0):,}" if ip else ""
        min0 = sum(1 for k, x in int_pat if x == "0")
        if len(ip) < min0:
            body = f"{int(ip or 0):0{min0},}"
    else:
        body = _fill_int(ip, int_pat)
    if decimals:
        frac_ph = [x for k, x in frac_part if k == "ph"]
        fp = list(fp)
        for i in range(len(fp) - 1, -1, -1):  # # 자리 끝 0 제거
            if frac_ph[i] == "#" and fp[i] == "0":
                fp.pop()
            else:
                break
        body += "." + "".join(fp)  # Excel: 0.## + 정수 → "3."
    pre = "".join(v2 for k, v2 in tokens[:first] if k in ("lit", "pct"))
    post = "".join(v2 for k, v2 in tokens[last + 1:] if k in ("lit", "pct", "comma"))
    return pre + body + post


def format_value(value, fmt: str | None, date1904: bool = False) -> str | None:
    """표시 문자열 또는 None(원래 값 사용)"""
    if fmt is None or value is None or isinstance(value, (bool, dt.date, dt.time)):
        return None
    f = fmt.strip()
    if not f or f.lower() == "general":
        return None
    secs = _sections(f)
    if isinstance(value, str):
        if len(secs) >= 4:
            sec = secs[3]
        elif len(secs) == 1 and "@" in secs[0]:
            sec = secs[0]
        else:
            return None
        return "".join(value if k == "at" else v for k, v in _tokens(sec) if k in ("lit", "at"))
    if not isinstance(value, (int, float)):
        return None
    if any(re.match(r"\s*\[[<>=]", s) for s in secs):
        return None  # 조건 섹션 ⚠️미지원
    neg_own = len(secs) >= 2 and value < 0
    if value == 0 and len(secs) >= 3:
        sec = secs[2]
    elif value < 0 and len(secs) >= 2:
        sec = secs[1]
    else:
        sec = secs[0]
    toks = _tokens(sec)
    if not toks:
        return ""  # ;;; = 숨김
    if _is_date(toks):
        return _date_text(value, toks, date1904)
    s = _number_text(value, toks)
    if s is None:
        return None
    if value < 0 and not neg_own and any(k == "ph" for k, _ in toks):
        s = "-" + s
    return s
