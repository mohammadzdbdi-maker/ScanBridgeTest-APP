; Inno Setup installer for ScanBridge
#define AppBuildDir "publish-single"
#define AppExecutableName "ScanBridgeTest.exe"

[Setup]
AppName=ScanBridge
AppVersion=1.0.2
AppPublisher=ScanBridge
AppPublisherURL=https://example.com/
AppSupportURL=https://example.com/
AppUpdatesURL=https://example.com/
DefaultDirName={localappdata}\Programs\ScanBridge
DefaultGroupName=ScanBridge
OutputBaseFilename=ScanBridge-Setup-1.0.2
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
CreateAppDir=yes
CreateUninstallRegKey=yes
AllowNoIcons=no
SetupIconFile=Assets\app-icon.ico
UninstallDisplayIcon={app}\ScanBridgeTest.exe
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
ShowLanguageDialog=no
UsePreviousAppDir=yes
UsePreviousTasks=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "{#AppBuildDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Run]
Filename: "{sys}\netsh.exe"; Parameters: "advfirewall firewall add rule name=""ScanBridge"" dir=in action=allow program=""{app}\ScanBridgeTest.exe"" enable=yes"; Flags: runhidden

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "ScanBridge"; ValueData: """{app}\ScanBridgeTest.exe"" --startup"; Flags: uninsdeletevalue

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: checkedonce

[Icons]
Name: "{group}\ScanBridge"; Filename: "{app}\{#AppExecutableName}"; WorkingDir: "{app}"; IconFilename: "{app}\{#AppExecutableName}"; IconIndex:0
Name: "{userdesktop}\ScanBridge"; Filename: "{app}\{#AppExecutableName}"; WorkingDir: "{app}"; Tasks: desktopicon; IconFilename: "{app}\{#AppExecutableName}"; IconIndex:0

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[UninstallRun]
Filename: "{sys}\netsh.exe"; Parameters: "advfirewall firewall delete rule name=""ScanBridge"""; Flags: runhidden
