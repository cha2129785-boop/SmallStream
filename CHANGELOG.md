# 스몰스트림 버전 기록 (CHANGELOG)

제작자. Shugle(슈글)

## v0.4.1 — 2026-11-23
작업 표시줄: 고정 아이콘과 실행 창이 따로 표시되던 문제

- 원인: 설치 EXE가 AppUserModelID(`Shugle.SmallStream`)를 설정 → 식별자 없는 바로가기(고정 아이콘)와 다른 앱으로 묶임
- 조치: `main.use_explicit_app_id()` — 설치 EXE(frozen)는 식별자 미설정 → Windows가 EXE 경로로 묶어 고정 아이콘에 합쳐짐. 개발용(python) 실행은 기존대로 설정
- 스트림머지 v0.19.1 과 같은 수정
- 테스트 55건

## v0.4 — 2026-11-21
설치 파일 — Windows '설치된 앱'(프로그램 추가/제거) 등록

- `build.bat` 기본 = 설치 파일 `dist_installer\SmallStream_Setup_<버전>.exe` (Inno Setup 6·7)
- `installer\SmallStream.iss`: 전용 AppId(스트림머지와 별개), 관리자 권한 불필요(선택 시 모든 사용자), 바탕 화면 바로가기 선택, 시작 메뉴, 제거 시 데이터 삭제 확인(기본 남김)
- `packaging\build_installer.py` v2.0 (build_exe.py 대체): 스트림머지 v5.1 의 Inno Setup 탐색(버전·위치 무관, winget, 직접 선택·기억) 재사용
- 설치판 EXE = 폴더판 (설치 폴더째 들어가므로 _internal 누락 없음, 시작 빠름)
- `build.bat --portable`: 설치 없이 쓰는 EXE 1개 (v0.3.1 방식)
- 설치 파일 빌드 성공 시 예전 휴대용 결과(dist\SmallStream.exe 등) 삭제
- 테스트: Inno 탐색 7건 이식 + AppId·데이터 삭제 확인·이전 결과 정리 (52건)

## v0.3.1 — 2026-11-18
다른 PC 실행 오류 수정 — `Failed to load Python DLL 'D:\_internal\python312.dll'`

- 원인: 폴더판(onedir) `SmallStream.exe`만 복사, `_internal` 폴더 누락
- build_exe v1.1: **기본 = EXE 1개** (`dist\SmallStream.exe`, EXE만 복사해도 실행). 폴더판은 `build.bat --onedir`
- 빌드 때 다른 형식의 이전 결과(`dist\SmallStream\` 폴더·이전 zip) 삭제 → 잘못된 파일 복사 방지
- zip 이름: EXE 1개 `SmallStream_<버전>.zip` / 폴더판 `SmallStream_<버전>_folder.zip`
- 프로그램 기능 변경 없음

## v0.3 — 2026-11-17
정보 탭 정리

- 삭제: 제작자 아래 설명 문구, 공개 소프트웨어 고지 + [라이선스 보기]
- EXE 내부 THIRD_PARTY_LICENSES.txt 제외 (배포 폴더·zip에는 계속 포함)

## v0.2 — 2026-11-17
정보 탭 추가

- 주 창 = 탭 2개: **파일 정리** / **정보**
- 정보 탭: 아이콘 · 이름 · 버전 · 제작자 · 설명 / 데이터 폴더 경로 + [폴더 열기] / 백업 폴더 경로 + [백업 폴더 열기] / 공개 소프트웨어 고지 + [라이선스 보기]
- 창 아래 버전·버튼 줄 제거 (정보 탭으로 이동)
- colprune_ui v0.10: `_make_body()` 훅 — 위젯 컨테이너를 바꿀 수 있게 (기본 = 창 자체, 동작 동일)
- EXE에 THIRD_PARTY_LICENSES.txt 포함 (라이선스 보기용)
- 테스트: 정보 탭 내용·탭 구성 확인 추가

## v0.1 — 2026-11-17
스트림머지 v0.19 '파일 정리 도구'를 단독 프로그램으로 분리

- 포함: 열 삭제(이름·고정 자리) · 서식 지우기(테두리 유지) · 셀 내용 정렬 · 중복 내용 지우기 · 프리셋 · 저장 위치 4종(원본 폴더 / 폴더 지정 / 원본 이동 / 원본 덮어쓰기+백업) · 같은 이름 파일 처리 · 파일 이름 규칙 · 열린 파일 사전 확인
- 모듈: colprune v0.9 · colprune_ui v0.9 · outpath v0.3 · loader · numfmt (스트림머지와 동일 코드)
- 신규 appcore v0.1: 데이터 폴더 `%APPDATA%\SmallStream`, 설정, Excel 글꼴·열 너비 (스트림머지 exporter·overlay·profile_store 대체)
- 첫 실행: 스트림머지 프리셋·'같은 이름 파일 처리' 설정 1회 복사 (원본 유지)
- 창 아래: 버전·제작자, [데이터 폴더 열기] · [백업 폴더 열기]
- 아이콘: 스트림머지 아이콘 색상 변형 (작업 표시줄 구분)
- 빌드: `build.bat` (ASCII) → `packaging\build_exe.py` v1.0 → PyInstaller onedir (기본) / `--onefile` → `dist\` + 배포 zip (UTF-8 이름 표시 확인)
- 의존성: openpyxl, xlrd (tksheet 제외)
- 테스트 6파일
