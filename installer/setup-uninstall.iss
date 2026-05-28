; ASI QB Middleware -- Inno Setup uninstaller script
;
; Build with scripts\build-installer.ps1 on the Windows workstation.
; Both the installer and uninstaller are produced in the same build run.

#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif

#define MyAppName      "ASI QB Middleware Uninstaller"
#define MyAppPublisher "Bill Nienaber"
#define MyOutputName   "ASI_FMQBsupport_Uninstall"

[Setup]
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={tmp}\asi-uninstall
DisableDirPage=yes
DisableProgramGroupPage=yes
CreateUninstallRegKey=no
Uninstallable=no
OutputDir=Output
OutputBaseFilename={#MyOutputName}
Compression=lzma2
SolidCompression=yes
ShowLanguageDialog=no
PrivilegesRequired=admin
WizardStyle=modern
DisableReadyPage=yes

[Messages]
WelcomeLabel1=ASI QB Middleware Uninstaller
WelcomeLabel2=This will remove the ASI QuickBooks Middleware from this workstation.%n%nThe scheduled task will be stopped and removed, and all middleware files will be deleted.%n%nUse the checkboxes on the next screen to also remove Python and Git if they are not needed for anything else on this machine.
WizardSelectTasks=Select components to remove
SelectTasksLabel2=The middleware will always be removed. Check the boxes below to also remove Python and Git.
ButtonInstall=Remove
FinishedLabel=ASI QB Middleware has been removed from this workstation.%n%nUninstall log: C:\ProgramData\ASI QB Middleware\uninstall.log

[Tasks]
Name: removepython; Description: "Remove Python  (uncheck if Python is used by other applications on this machine)"; Flags: unchecked
Name: removegit;    Description: "Remove Git  (uncheck if Git is used for other development on this machine)"; Flags: unchecked

[Files]
Source: "..\scripts\uninstall.ps1"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Code]
// Build the PowerShell parameter string based on which tasks were selected,
// then run uninstall.ps1.
//
// IMPORTANT: must use ssPostInstall, not ssInstall.
// ssInstall fires at the START of the install step, before files are extracted
// to {tmp}. ssPostInstall fires after extraction is complete, so uninstall.ps1
// is guaranteed to exist when we call it.
procedure CurStepChanged(CurStep: TSetupStep);
var
  PSExe, Params: String;
  ResultCode: Integer;
begin
  if CurStep = ssPostInstall then begin
    // Use full path to avoid PATH lookup issues in Inno Setup's Pascal context.
    PSExe := ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe');
    Params := '-ExecutionPolicy Bypass -NoProfile -File "' +
              ExpandConstant('{tmp}\uninstall.ps1') + '" -Silent';
    if IsTaskSelected('removepython') then
      Params := Params + ' -RemovePython';
    if IsTaskSelected('removegit') then
      Params := Params + ' -RemoveGit';
    if not Exec(PSExe, Params, '', SW_SHOW, ewWaitUntilTerminated, ResultCode) then
      MsgBox('Could not launch PowerShell. Uninstall may not have completed.' + #13#10 +
             'Run scripts\uninstall.ps1 manually from the middleware directory.',
             mbError, MB_OK);
  end;
end;
