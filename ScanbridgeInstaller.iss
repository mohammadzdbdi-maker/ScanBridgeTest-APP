; Scanbridge Installer - folder publish build
; 1) Publish first:
;    dotnet publish "ScanBridgeTest.csproj" -c Release -r win-x64 --self-contained true -p:PublishSingleFile=false -p:PublishReadyToRun=true -o ".\publish-folder"
; 2) Then compile this script with Inno Setup.

#define AppName "Scanbridge"
#define AppVersion "2.1.9"
#define AppPublisher "Scanbridge"
#define AppURL "https://scanbridge.ir"
#define AppExeName "ScanBridgeTest.exe"
#define AppBuildDir "publish-folder"

[Setup]
AppId={{A7B6F2D5-2A55-4F90-8AC7-5C4A4E425247}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
DefaultDirName={localappdata}\Programs\Scanbridge
DefaultGroupName=Scanbridge
OutputBaseFilename=Scanbridge-Setup-{#AppVersion}
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
CloseApplications=yes
RestartIfNeededByRun=no
CreateAppDir=yes
CreateUninstallRegKey=yes
AllowNoIcons=yes
SetupIconFile=Assets\app-icon.ico
UninstallDisplayIcon={app}\{#AppExeName}
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
ShowLanguageDialog=no
UsePreviousAppDir=yes
UsePreviousTasks=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
; If Persian.isl is not installed in your Inno Setup Languages folder, comment the next line.
;Name: "persian"; MessagesFile: "compiler:Languages\Persian.isl"

[Tasks]
Name: "desktopicon"; Description: "ساخت آیکون روی دسکتاپ"; GroupDescription: "گزینه‌های نصب:"; Flags: checkedonce

[Files]
Source: "{#AppBuildDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Scanbridge"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\Assets\app-icon.ico"
Name: "{userdesktop}\Scanbridge"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\Assets\app-icon.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "اجرای Scanbridge"; Flags: nowait postinstall skipifsilent

[Code]
procedure KillProcess(const ExeName: string);
var
  ResultCode: Integer;
begin
  // taskkill returns a non-zero exit code when the process isn't running — that's fine,
  // we don't care about the result either way, we just want it not to be running.
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM "' + ExeName + '" /T', '', SW_HIDE,
       ewWaitUntilTerminated, ResultCode);
end;

function InitializeSetup(): Boolean;
begin
  // adb.exe (Android Debug Bridge, used for mobile pairing) starts its own background
  // server process that keeps running even after Scanbridge itself is closed. Windows
  // refuses to overwrite a running .exe ("DeleteFile failed; Access is denied"), which
  // breaks the update on machines where it's still alive. Kill both it and the main app
  // before Setup starts extracting files.
  KillProcess('adb.exe');
  KillProcess('ScanBridgeTest.exe');
  Result := True;
end;
