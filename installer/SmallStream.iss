; 스몰스트림 설치 파일 (Inno Setup 6) — Windows '설치된 앱'(프로그램 추가/제거)에 등록
; 빌드: build.bat (Inno Setup 6·7 이상 모두 사용 가능)
#define MyAppName "스몰스트림"
#define MyAppId "SmallStream"
#ifndef MyAppVersion
  #define MyAppVersion "0.4"
#endif
#define MyAppPublisher "Shugle(슈글)"
#define MyAppExeName "SmallStream.exe"
; 빌드 결과 위치 (packaging\build_installer.py 가 /DDistDir=... 로 전달, 없으면 기본값)
#ifndef DistDir
  #define DistDir "..\dist\SmallStream"
#endif

[Setup]
; 스트림머지와 다른 AppId → 별개 프로그램으로 등록 (바꾸지 말 것: 바꾸면 업데이트가 새 프로그램으로 설치됨)
AppId={{4B1F9332-CCFC-47E6-AED4-ED68364B4325}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppId}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
; 관리자 권한 없이 설치 (사용자 폴더). 관리자는 설치 시 '모든 사용자' 선택 가능
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=..\dist_installer
OutputBaseFilename=SmallStream_Setup_{#MyAppVersion}
SetupIconFile=..\assets\smallstream.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoProductName={#MyAppName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
; Inno Setup 6.3 이상·7 문법 (6.0~6.2는 build_installer.py 가 x64로 자동 변경)
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes

[Languages]
; Korean.isl 없으면 build_installer.py 가 이 줄을 빼고 영어 화면으로 빌드
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
Name: "desktopicon"; Description: "바탕 화면에 바로가기 만들기"; GroupDescription: "추가 작업:"

[Files]
Source: "{#DistDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\THIRD_PARTY_LICENSES.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\사용법.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{#MyAppName} 제거"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{#MyAppName} 실행"; Flags: nowait postinstall skipifsilent

[Code]
// 제거 시: 데이터(프리셋·설정·백업 — 백업에는 원본 파일·개인정보 포함) 삭제 여부 확인. 기본 = 남김
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  DataDir: String;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    DataDir := ExpandConstant('{userappdata}\{#MyAppId}');
    if DirExists(DataDir) then
      if MsgBox('프리셋·설정·백업 데이터도 삭제할까요?' + #13#10 + '(백업 폴더에는 원본 파일이 들어 있음)' + #13#10 + #13#10 + DataDir + #13#10 + #13#10 +
                '[예] 삭제 (복구 불가)   [아니요] 남겨 둠 (다시 설치하면 그대로 사용)',
                mbConfirmation, MB_YESNO or MB_DEFBUTTON2) = IDYES then
        DelTree(DataDir, True, True, True);
  end;
end;
