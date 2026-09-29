; Inno Setup 安装脚本
; 下载 Inno Setup: https://jrsoftware.org/isinfo.php
; 编译: 右键 setup.iss → Compile，或在 Inno Setup Compiler 中打开

[Setup]
AppId=JinhuaJuhuo-Table-Tool
AppName=金华聚火表格处理
AppVersion=1.3.0
AppPublisher=订小易
DefaultDirName={autopf}\金华聚火表格处理
DefaultGroupName=金华聚火表格处理
OutputDir=.\installer
OutputBaseFilename=金华聚火表格处理_Setup
SetupIconFile=.\favicon.ico
UninstallDisplayIcon={app}\_internal\favicon.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

[Languages]
Name: "chinese"; MessagesFile: "D:\WFR\D2Y\Default.isl"

[Files]
Source: "dist\JinhuaJuhuo\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{autoprograms}\金华聚火表格处理"; Filename: "{app}\JinhuaJuhuo.exe"; IconFilename: "{app}\_internal\favicon.ico"
Name: "{autodesktop}\金华聚火表格处理"; Filename: "{app}\JinhuaJuhuo.exe"; IconFilename: "{app}\_internal\favicon.ico"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "快捷方式:"; Flags: checkedonce

[Run]
Filename: "{app}\JinhuaJuhuo.exe"; Description: "启动 金华聚火表格处理"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
