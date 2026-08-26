#ifndef MyAppVersion
  #error MyAppVersion must be provided by build_windows_installers.py
#endif
#ifndef SourceDir
  #error SourceDir must be provided by build_windows_installers.py
#endif
#ifndef OutputDir
  #error OutputDir must be provided by build_windows_installers.py
#endif
#ifndef SetupIconFile
  #error SetupIconFile must be provided by build_windows_installers.py
#endif
#ifndef InstallerSuffix
  #error InstallerSuffix must be provided by build_windows_installers.py
#endif
#ifndef MinWindowsVersion
  #error MinWindowsVersion must be provided by build_windows_installers.py
#endif

#define MyAppName "HRToolkit"
#define MyAppPublisher "xhzwjc"
#define MyAppExeName "HRToolkit.exe"

[Setup]
AppId={{8BBD86A8-CB7D-4D5D-A940-BD3F5942FBA6}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\HRToolkit
DefaultGroupName=HRToolkit
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion={#MinWindowsVersion}
#ifdef Win7Compatibility
OnlyBelowVersion=6.3
#endif
OutputDir={#OutputDir}
OutputBaseFilename=HRToolkit_{#MyAppVersion}_{#InstallerSuffix}-setup
SetupIconFile={#SetupIconFile}
UninstallDisplayIcon={app}\app\{#MyAppExeName}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no
UsePreviousAppDir=yes
ChangesAssociations=no
ChangesEnvironment=no
VersionInfoVersion={#MyAppVersion}.0
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}
#ifdef SignToolName
SignTool={#SignToolName}
SignedUninstaller=yes
#endif

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Default.isl,ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "其他选项："; Flags: unchecked

[Files]
; 安装器和卸载器保留在 {app}，可自更新 payload 独立放在 {app}\app。
; HRToolkitUpdater 只替换 sys.executable.parent，因此不会删除 unins*.exe。
Source: "{#SourceDir}\*"; DestDir: "{app}\app"; Flags: ignoreversion recursesubdirs createallsubdirs

[InstallDelete]
#ifdef CleanExistingPayload
; Win7 首次安装会替换原现代 payload，避免 Python 3.12 DLL 残留混用。
Type: filesandordirs; Name: "{app}\app"
#else
; 仅当现代包替换已安装的 Win7 payload 时清理；普通现代版安装不受影响。
Type: filesandordirs; Name: "{app}\app"; Check: ExistingPayloadIsWin7
#endif

[Icons]
Name: "{userprograms}\HRToolkit"; Filename: "{app}\app\{#MyAppExeName}"; WorkingDir: "{app}\app"
Name: "{userdesktop}\HRToolkit"; Filename: "{app}\app\{#MyAppExeName}"; WorkingDir: "{app}\app"; Tasks: desktopicon

[Run]
Filename: "{app}\app\{#MyAppExeName}"; Description: "启动 HRToolkit"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; 自更新可能改变 payload 文件集合，卸载时递归清理 app 子目录。
Type: filesandordirs; Name: "{app}\app"
Type: dirifempty; Name: "{app}"

#ifndef CleanExistingPayload
[Code]
function ExistingPayloadIsWin7: Boolean;
begin
  Result := FileExists(ExpandConstant('{app}\app\_internal\python38.dll'));
end;
#endif
