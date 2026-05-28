; ASI QB Middleware -- Inno Setup installer script
;
; Build with scripts\build-installer.ps1 on the Windows workstation.
; Do not run iscc.exe directly -- the build script passes the version number
; and validates that bundled.env is populated before compiling.
;
; Prerequisites:
;   - Inno Setup 6 installed (https://jrsoftware.org/isinfo.php)
;   - installer\bundled.env populated with real values (see bundled.env.example)

#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif

#define MyAppName      "ASI QB Middleware"
#define MyAppPublisher "Bill Nienaber"

#ifndef MyOutputName
  #define MyOutputName "ASI-QB-Middleware-Setup"
#endif

[Setup]
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
; Inno Setup requires a DefaultDirName even when DisableDirPage=yes.
; We use a local app data path as the placeholder -- the actual middleware
; installs to C:\Services\asi-qb-middleware via install.ps1.
DefaultDirName={localappdata}\ASIQBMiddleware
DisableDirPage=yes
DisableProgramGroupPage=yes
; No entry in Programs & Features -- this is a background service, not a user app.
CreateUninstallRegKey=no
Uninstallable=no
OutputDir=Output
OutputBaseFilename={#MyOutputName}
Compression=lzma2
SolidCompression=yes
ShowLanguageDialog=no
; Admin required: install.ps1 creates C:\Services\, installs Git system-wide,
; and may need elevated access for Python installation.
PrivilegesRequired=admin
WizardStyle=modern
; Show a clean wizard -- suppress dir/group pages, keep Welcome and Finish.
DisableReadyPage=yes
SetupIconFile=

[Messages]
WelcomeLabel1=ASI QuickBooks Middleware
WelcomeLabel2=This installer will set up the ASI QuickBooks Middleware on this workstation.%n%nThe middleware runs in the background and allows FileMaker to post invoices to QuickBooks Desktop.%n%nClick Install to begin.
FinishedLabel=ASI QuickBooks Middleware has been installed successfully.%n%nThe middleware will start automatically each time you log in. QuickBooks Desktop must be open and authorized before FileMaker can post invoices.%n%nInstall log: C:\Services\asi-qb-middleware\install.log
ButtonInstall=Install

[Files]
; Both files are extracted to the Inno Setup temp folder and deleted after install.
Source: "..\scripts\install.ps1"; DestDir: "{tmp}"; Flags: deleteafterinstall
Source: "bundled.env";            DestDir: "{tmp}"; DestName: "bundled.env"; Flags: deleteafterinstall

[Run]
; Run install.ps1 with the bundled .env and -Silent flag.
; -Silent suppresses "Press any key" prompts so the installer flow stays clean.
; -ResultFile tells the script where to write its outcome so the [Code] section
;  below can show a proper error dialog if something went wrong.
Filename: "powershell.exe"; \
  Parameters: "-ExecutionPolicy Bypass -NoProfile -File ""{tmp}\install.ps1"" -BundledEnvPath ""{tmp}\bundled.env"" -Silent -ResultFile ""{tmp}\install-result.txt"""; \
  StatusMsg: "Installing ASI QB Middleware (this may take a few minutes)..."; \
  Flags: waituntilterminated

[Code]
// Read the first line of a file. Returns empty string if file doesn't exist.
function ReadResultFile(const Path: String): String;
var
  Lines: TArrayOfString;
begin
  Result := '';
  if LoadStringsFromFile(Path, Lines) and (GetArrayLength(Lines) > 0) then
    Result := Trim(Lines[0]);
end;

// After the installer finishes, check the result file written by install.ps1.
// If it's missing or doesn't contain "OK", show a clear error dialog so IT
// knows the installation did not complete -- not a silent false success.
procedure DeinitializeSetup();
var
  ResultPath, Status: String;
begin
  ResultPath := ExpandConstant('{tmp}\install-result.txt');
  if FileExists(ResultPath) then begin
    Status := ReadResultFile(ResultPath);
    if Status <> 'OK' then
      MsgBox(
        'Installation did not complete successfully.' + #13#10#13#10 +
        Status + #13#10#13#10 +
        'Log file: ' + ExpandConstant('{usertmp}') + '\asi-qb-install.log' + #13#10#13#10 +
        'Please send the log file to your administrator.',
        mbError, MB_OK);
  end else begin
    MsgBox(
      'Installation did not complete successfully (no result was recorded).' + #13#10#13#10 +
      'Log file: ' + ExpandConstant('{usertmp}') + '\asi-qb-install.log' + #13#10#13#10 +
      'Please send the log file to your administrator.',
      mbError, MB_OK);
  end;
end;
