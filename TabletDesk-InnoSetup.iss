; TabletDesk Inno Setup Script
; Creates a single EXE installer with EULA, shortcuts, and all files

[Setup]
AppName=TabletDesk
AppVersion=1.0
DefaultDirName={userpf}\TabletDesk
DefaultGroupName=TabletDesk
DisableProgramGroupPage=yes
OutputDir={userdesktop}
OutputBaseFilename=TabletDesk-Installer
SetupIconFile=TabletDeskTrayIcon.ico
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "TabletDesk-Launcher.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "TabletDesk.ahk"; DestDir: "{app}"; Flags: ignoreversion
Source: "TabletDeskTrayIcon.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "TabletDesk-EULA.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\TabletDesk"; Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\TabletDesk-Launcher.ps1"""; WorkingDir: "{app}"; IconFilename: "{app}\TabletDeskTrayIcon.ico"
Name: "{userdesktop}\TabletDesk"; Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\TabletDesk-Launcher.ps1"""; WorkingDir: "{app}"; IconFilename: "{app}\TabletDeskTrayIcon.ico"

[Run]
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\TabletDesk-Launcher.ps1"""; Description: "Launch TabletDesk now"; Flags: nowait postinstall skipifsilent

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
end;

[CustomMessages]
EULATitle=End User License Agreement
EULAAccept=I accept the agreement

[LicenseFile]
LicenseFile=TabletDesk-EULA.txt
