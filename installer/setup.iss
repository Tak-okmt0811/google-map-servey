; 競合分析ダッシュボード Windows インストーラ定義（Inno Setup）
;
; ビルド:
;   ISCC.exe /DMyAppVersion=1.0.0 installer\setup.iss
; （バージョンを省略した場合は "0.0.0-dev" になる）
;
; 出力: installer\output\CompetitorDashboardSetup.exe
;
; 管理者権限・UACプロンプトなしでインストールできるよう、インストール先は
; Program Filesではなく %LOCALAPPDATA% 配下（ユーザー単位）にしている。
; これにより、後からinputフォルダにCSVを置く操作も管理者権限なしで行える。

#ifndef MyAppVersion
  #define MyAppVersion "0.0.0-dev"
#endif

#define MyAppName "競合分析ダッシュボード"
#define MyAppExeName "competitor-dashboard.exe"
#define MyAppPublisher "Lumel Plan"

[Setup]
; 一度発行したら変更しないこと（変更するとアップグレードではなく別アプリ扱いになる）
AppId={{B7C1E9F0-2C7B-4E9D-9E4A-6C6F3E9D6A21}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\CompetitorDashboard
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=output
OutputBaseFilename=CompetitorDashboardSetup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
SetupLogging=yes

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "デスクトップにアイコンを作成する"; GroupDescription: "追加のアイコン:"

[Dirs]
Name: "{app}\input"

[Files]
Source: "..\dist\competitor-dashboard.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "input_readme.txt"; DestDir: "{app}\input"; DestName: "readme.txt"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
