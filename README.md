# 스몰스트림 (SmallStream)

Batch Excel/CSV cleaner for Windows: delete columns, clear formatting, align cells, remove duplicates. Originals stay safe.

여러 Excel·CSV 파일을 한 번에 정리하는 Windows 프로그램. 스트림머지의 '파일 정리 도구'를 단독 프로그램으로 분리.

## 기능
- 열 삭제 — 열 이름 또는 고정 자리(A, C, F-H)
- 서식 지우기 — 배경색·굵게·글꼴 등 (테두리 유지)
- 셀 내용 정렬 — 왼쪽 / 가운데 / 오른쪽
- 중복 내용 지우기 — 처음 값만 남김 (열마다·연속·시트 전체·행)
- 프리셋 저장 · 미리보기
- 결과는 새 `.xlsx` 파일. 원본 덮어쓰기는 선택 (자동 백업)
- 읽기: xlsx · xlsm · xls · HTML 형식 xls · csv

## 설치
[Releases](../../releases)에서 `SmallStream_Setup_<버전>.exe` 받아 실행 → Windows '설치된 앱'에 등록

## 소스에서 실행
```
install.bat          # Python 3.12 + openpyxl, xlrd
SmallStream.pyw      # 또는 run.bat
```

## 설치 파일 만들기
Python 3.12 + [Inno Setup](https://jrsoftware.org/isdl.php) 6 이상
```
build.bat            # → dist_installer\SmallStream_Setup_<버전>.exe
build.bat --portable # → dist\SmallStream.exe (설치 없이 실행)
```

## 테스트
```
pip install pytest xlwt
python -m pytest tests
```

## 데이터 위치
`%APPDATA%\SmallStream` — 프리셋·설정·백업(원본 사본)·error.log

자세한 사용법: [사용법.md](사용법.md) · 변경 기록: [CHANGELOG.md](CHANGELOG.md) · 사용한 공개 소프트웨어: [THIRD_PARTY_LICENSES.txt](THIRD_PARTY_LICENSES.txt)

제작자. Shugle(슈글)
