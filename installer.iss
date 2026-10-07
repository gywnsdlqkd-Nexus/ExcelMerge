; ExcelMerge 설치 파일 — Inno Setup 6
;
; 왜 설치형인가. onefile 은 켤 때마다 번들 전체를 %TEMP% 에 풀어서 그게 시작 시간의
; 가장 큰 몫이었다(실측: onefile 4.6초 vs onedir 2.7초). onedir 로 바꾸면 그 비용이
; 사라지지만 배포물이 폴더가 된다 — zip 으로 주면 "압축을 안 풀고 안에서 exe 를
; 더블클릭" 하는 사고가 반복된다(_internal 을 못 찾아 실행 실패). 설치 파일이 그 문제를
; 없애고, 업데이트 때 실행 중인 앱을 닫고 파일을 바꾸는 일도 대신해 준다.
;
; 빌드:
;   python -m PyInstaller ExcelMerge.spec      → dist/ExcelMerge/
;   ISCC.exe /DMyVersion=<N> installer.iss     → dist/ExcelMerge_Setup.exe
; (release.py 가 두 단계를 묶어서 돌린다)

#ifndef MyVersion
  #error MyVersion 을 넘겨야 합니다 — ISCC /DMyVersion=216
#endif

#define MyAppName "ExcelMerge"
#define MyAppExeName "ExcelMerge.exe"
#define MyAppPublisher "gywnsdlqkd"

[Setup]
; AppId 는 **절대 바꾸지 않는다.** 이 값으로 같은 앱인지 판단해 덮어 설치한다.
; 바꾸면 기존 설치를 못 찾아 두 벌이 깔린다.
AppId={{8E6B1B2A-9C4D-4F77-B1E2-5A3D7C9F0A41}
AppName={#MyAppName}
AppVersion={#MyVersion}
AppVerName={#MyAppName} v{#MyVersion}
AppPublisher={#MyAppPublisher}
VersionInfoVersion=0.0.{#MyVersion}

; 사용자 영역에 설치한다 — 관리자 권한(UAC)이 필요 없다. 동료들이 받아서 바로
; 실행할 수 있어야 하고, 자동 업데이트도 조용히 돌아야 한다(UAC 가 뜨면 멈춘다).
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
; {autopf} 는 PrivilegesRequired=lowest 에서 %LOCALAPPDATA%\Programs 로 풀린다.
DefaultDirName={autopf}\{#MyAppName}
DisableDirPage=auto
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes

OutputDir=dist
; 파일 이름에 버전을 붙이지 않는다 — 받는 사람이 늘 같은 이름을 본다.
; 버전은 설치 마법사 제목과 '앱 및 기능' 목록(AppVerName)에 나온다.
OutputBaseFilename=ExcelMerge_Setup
SetupIconFile=images\app_icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

; 업데이트 때 실행 중이면 닫고 교체한다 — 이게 onedir 의 '폴더 교체' 문제를 대신 푼다.
CloseApplications=force
RestartApplications=no

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
Name: "desktopicon"; Description: "바탕 화면에 바로 가기 만들기"; \
    GroupDescription: "추가 작업:"; Flags: unchecked

[Files]
; dist/ExcelMerge/ 통째로. recursesubdirs 로 _internal 까지 따라간다.
Source: "dist\{#MyAppName}\*"; DestDir: "{app}"; \
    Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{#MyAppName} 제거"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; /SILENT 로 도는 자동 업데이트에서는 체크박스가 없으므로 postinstall 이 안 탄다 —
; 그 경우 updater 가 직접 다시 띄운다.
Filename: "{app}\{#MyAppExeName}"; Description: "{#MyAppName} 실행"; \
    Flags: nowait postinstall skipifsilent

[UninstallDelete]
; 설치 후 생긴 파이썬 캐시 등 — 남으면 폴더가 안 지워진다.
Type: filesandordirs; Name: "{app}\_internal\__pycache__"
Type: dirifempty; Name: "{app}"
