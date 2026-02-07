#define MyAppName "DWG办公工具箱"
#define MyAppVersion "1.0.2.0"
#define MyAppPublisher "DWG办公工具箱"
#define MyAppExeName "DwgOfficeToolbox.App.exe"

[Setup]
AppId={{7B6F3C9F-3D6A-4E41-9F4D-8C6A2B1E2F61}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=..\dist\setup
OutputBaseFilename=DwgOfficeToolbox_Win7_Setup
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
MinVersion=0,6.1

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面图标"; GroupDescription: "附加任务"

[Files]
Source: "..\dist\app\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "启动 {#MyAppName}"; Flags: nowait postinstall skipifsilent
