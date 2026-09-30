; Inno Setup - PySide6 版
[Setup]
AppId=JinhuaJuhuo-Table-Tool
AppName=金华聚火表格处理
AppVersion=3.1.3
AppPublisher=订小易
DefaultDirName={autopf}\金华聚火表格处理
DefaultGroupName=金华聚火表格处理
UsePreviousAppDir=yes
DisableDirPage=no
OutputDir=.\installer
OutputBaseFilename=金华聚火表格处理_v3.1.3_Setup
SetupIconFile=.\favicon.ico
UninstallDisplayIcon={app}\_internal\favicon.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

[Languages]
Name: "chinese"; MessagesFile: "Default.isl"

[Files]
Source: "dist\JinhuaH\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{autoprograms}\金华聚火表格处理"; Filename: "{app}\JinhuaH.exe"; IconFilename: "{app}\_internal\favicon.ico"
Name: "{autodesktop}\金华聚火表格处理"; Filename: "{app}\JinhuaH.exe"; IconFilename: "{app}\_internal\favicon.ico"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "快捷方式:"; Flags: checkedonce

[Run]
Filename: "{app}\JinhuaH.exe"; Description: "启动 金华聚火表格处理"; Flags: nowait postinstall skipifsilent

[InstallDelete]
Type: files; Name: "{app}\jinhua_engine.exe"
Type: files; Name: "{app}\JinhuaJuhuo.exe"
Type: files; Name: "{app}\*.pyc"
Type: files; Name: "{app}\*.spec"

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
