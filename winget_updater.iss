#define MyAppName "Winget Updater"
#define MyAppVersion "1.0"
#define MyAppPublisher "Visma IT"
#define MyAppExeName "WingetUpdater.exe"
#define MyServiceName "WingetUpdaterService"

[Setup]
AppId={{891E2B26-53C5-4371-A7BF-95A6AE3E944F}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
LicenseFile=README.md
OutputDir=installer_output
OutputBaseFilename=WingetUpdater_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\WingetUpdater\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Install and start the service
Filename: "{app}\{#MyAppExeName}"; Parameters: "--install"; Flags: runhidden waituntilterminated; StatusMsg: "Installing service..."
Filename: "powershell.exe"; Parameters: "Start-Service -Name {#MyServiceName}"; Flags: runhidden; StatusMsg: "Starting service..."
; Launch the application
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallRun]
; Stop and remove the service during uninstallation
Filename: "powershell.exe"; Parameters: "Stop-Service -Name {#MyServiceName} -Force -ErrorAction SilentlyContinue"; Flags: runhidden
Filename: "powershell.exe"; Parameters: "Remove-Service -Name {#MyServiceName} -ErrorAction SilentlyContinue"; Flags: runhidden

[Code]
// Check if the service exists before trying to install
function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
begin
  Result := True;
  
  // Try to stop the service if it exists
  Exec('powershell.exe', 'Stop-Service -Name ' + '{#MyServiceName}' + ' -Force -ErrorAction SilentlyContinue', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  
  // Try to remove existing service
  Exec('powershell.exe', 'Remove-Service -Name ' + '{#MyServiceName}' + ' -ErrorAction SilentlyContinue', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;
