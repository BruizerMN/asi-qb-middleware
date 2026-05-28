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

#define MyAppName    "ASI QB Middleware"
#define MyAppPublisher "Bill Nienaber"
#define MyOutputName "ASI-QB-Middleware-Setup"

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
FinishedLabel=ASI QuickBooks Middleware has been installed successfully.%n%nThe middleware will start automatically each time you log in. QuickBooks Desktop must be open and authorized before FileMaker can post invoices.
ButtonInstall=Install

[Files]
; Both files are extracted to the Inno Setup temp folder and deleted after install.
Source: "..\scripts\install.ps1"; DestDir: "{tmp}"; Flags: deleteafterinstall
Source: "bundled.env";            DestDir: "{tmp}"; DestName: "bundled.env"; Flags: deleteafterinstall

[Run]
; Run install.ps1 with the bundled .env and -Silent flag.
; -Silent suppresses "Press any key" prompts so the installer flow stays clean.
; The script downloads Python/Git if needed, clones the repo, installs deps,
; writes the .env, registers the Task Scheduler task, and starts the middleware.
Filename: "powershell.exe"; \
  Parameters: "-ExecutionPolicy Bypass -NoProfile -File ""{tmp}\install.ps1"" -BundledEnvPath ""{tmp}\bundled.env"" -Silent"; \
  StatusMsg: "Installing ASI QB Middleware (this may take a few minutes)..."; \
  Flags: waituntilterminated
