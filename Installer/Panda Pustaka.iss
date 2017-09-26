#define MyAppName "Panda Pustaka"
#define MyAppVersion "1.0"
#define MyAppPublisher "Gurisa © 2015"
#define MyAppURL "http://www.Gurisa.Com/"
#define MyAppExeName "Perpustakaan.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{32DE5361-2A90-49FD-A1B6-05A634274AF3}
AppName={#MyAppName}
AppVersion=1.0
VersionInfoVersion= 1.0.0.0
VersionInfoDescription=Panda Pustaka
AppVerName={#MyAppName} {#MyAppVersion}
AppCopyright=Gurisa (C) 2015, Inc.
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf}\{#MyAppName}
DisableDirPage=yes
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=C:\Users\R\Desktop\Project\Installer\License.txt
InfoBeforeFile=C:\Users\R\Desktop\Project\Installer\Show Before.txt
InfoAfterFile=C:\Users\R\Desktop\Project\Installer\Show After.txt
OutputDir=C:\Users\R\Desktop\Project\Installer
OutputBaseFilename=Panda Pustaka
SetupIconFile=C:\Users\R\Desktop\Project\Icon\Panda Logo 4 Transparant.ico
Compression=lzma
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: quicklaunchicon; Description: "Create a &Quick Launch icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
;
Source: "C:\Users\R\Desktop\Project\Perpustakaan.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\R\Desktop\Project\Panel.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\R\Desktop\Project\Log.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\R\Desktop\Project\Panel.panpus"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\R\Desktop\Project\Crystl32.oca"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\R\Desktop\Project\Crystl32.OCX"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\R\Desktop\Project\Crystl32.oca"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist
Source: "C:\Users\R\Desktop\Project\Crystl32.OCX"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist
;
Source: "C:\Users\R\Desktop\Project\MSCOMCT2.DEP"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\R\Desktop\Project\MSCOMCT2.oca"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\R\Desktop\Project\MSCOMCT2.OCX"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\R\Desktop\Project\MSCOMCT2.SRG"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\R\Desktop\Project\MSCOMCT2.DEP"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist 
Source: "C:\Users\R\Desktop\Project\MSCOMCT2.oca"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist 
Source: "C:\Users\R\Desktop\Project\MSCOMCT2.OCX"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist 
Source: "C:\Users\R\Desktop\Project\MSCOMCT2.SRG"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist
Source: "C:\Users\R\Desktop\Project\MSCOMCT2.OCX"; DestDir: "{sys}"; Flags: regserver sharedfile restartreplace 
;
Source: "C:\Users\R\Desktop\Project\MSCOMCTL.DEP"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\R\Desktop\Project\MSCOMCTL.oca"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\R\Desktop\Project\MSCOMCTL.OCX"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\R\Desktop\Project\MSCOMCTL.SRG"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\R\Desktop\Project\MSCOMCTL.DEP"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist 
Source: "C:\Users\R\Desktop\Project\MSCOMCTL.oca"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist 
Source: "C:\Users\R\Desktop\Project\MSCOMCTL.OCX"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist 
Source: "C:\Users\R\Desktop\Project\MSCOMCTL.SRG"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist 
Source: "C:\Users\R\Desktop\Project\MSCOMCTL.OCX"; DestDir: "{sys}"; Flags: regserver sharedfile restartreplace

;
Source: "C:\Users\R\Desktop\Project\msvbvm60.dll"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist
Source: "C:\Users\R\Desktop\Project\msvbvm60.dll"; DestDir: "{sys}"; Flags: regserver sharedfile restartreplace
;
Source: "C:\Users\R\Desktop\Project\VB6.OLB"; DestDir: "C:\Program Files\Microsoft Visual Studio\VB98\"; Flags: uninsneveruninstall onlyifdoesntexist
;
Source: "C:\Users\R\Desktop\Project\stdole2.tlb"; DestDir: "{sys}"; Flags: uninsneveruninstall onlyifdoesntexist
;
Source: "C:\Users\R\Desktop\Project\DAO350.dll"; DestDir: "C:\Program Files\Common Files\Microsoft Shared\DAO\"; Flags: uninsneveruninstall onlyifdoesntexist
Source: "C:\Users\R\Desktop\Project\DAO350.dll"; DestDir: "C:\Program Files\Common Files\Microsoft Shared\DAO\"; Flags: regserver sharedfile restartreplace
;
Source: "C:\Users\R\Desktop\Project\msador28.tlb"; DestDir: "C:\Program Files\Common Files\System\ado\"; Flags: uninsneveruninstall onlyifdoesntexist
;
Source: "C:\Users\R\Desktop\Project\msado60.tlb"; DestDir: "C:\Program Files\Common Files\System\ado\"; Flags: uninsneveruninstall onlyifdoesntexist
;
Source: "C:\Users\R\Desktop\Project\Panduan\Panduan.pdf"; DestDir: "{app}\Panduan"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\R\Desktop\Project\Report\*"; DestDir: "{app}\Report"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\R\Desktop\Project\Component\Microsoft Visual Basic 6.0 Runtime\*"; DestDir: "{app}\Component\Microsoft Visual Basic 6.0 Runtime"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\R\Desktop\Project\Component\Crystal Report Redistribution 8.5 Runtime\*"; DestDir: "{app}\Component\Crystal Report Redistribution 8.5 Runtime"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\R\Desktop\Project\Component\MySQL Connector ODBC 5.2.2\*"; DestDir: "{app}\Component\MySQL Connector ODBC 5.2.2"; Flags: ignoreversion recursesubdirs createallsubdirs
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:ProgramOnTheWeb,{#MyAppName}}"; Filename: "{#MyAppURL}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Registry]
Root: HKLM; Subkey: "SOFTWARE\ODBC\"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\ODBC\"; ValueType: "none"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"; ValueType: "string"; ValueName: "Panda Pustaka"; ValueData: "MySQL ODBC 5.2w Driver"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\ODBC\ODBC.INI\Panda Pustaka"; ValueType: "none"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\ODBC\ODBC.INI\Panda Pustaka"; ValueType: "string"; ValueName: "DATABASE"; ValueData: "db_perpus"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\ODBC\ODBC.INI\Panda Pustaka"; ValueType: "string"; ValueName: "DESCRIPTION"; ValueData: "Panda Pustaka"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\ODBC\ODBC.INI\Panda Pustaka"; ValueType: "string"; ValueName: "Driver"; ValueData: "C:\\Program Files\\MySQL\\Connector ODBC 5.2\\Unicode\\myodbc5w.dll"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\ODBC\ODBC.INI\Panda Pustaka"; ValueType: "string"; ValueName: "PORT"; ValueData: "3306"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\ODBC\ODBC.INI\Panda Pustaka"; ValueType: "string"; ValueName: "PWD"; ValueData: "pustaka"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\ODBC\ODBC.INI\Panda Pustaka"; ValueType: "string"; ValueName: "SERVER"; ValueData: "192.168.100.1"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\ODBC\ODBC.INI\Panda Pustaka"; ValueType: "string"; ValueName: "UID"; ValueData: "panda"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\ODBC\ODBCINST.INI\MySQL ODBC 5.2w Driver"; ValueType: "none"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\ODBC\ODBCINST.INI\MySQL ODBC 5.2w Driver"; ValueType: "string"; ValueName: "CPTimeout"; ValueData: "<not pooled>"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\ODBC\ODBCINST.INI\MySQL ODBC 5.2w Driver"; ValueType: "string"; ValueName: "Driver"; ValueData: "C:\\Program Files\\MySQL\\Connector ODBC 5.2\\Unicode\\myodbc5w.dll"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\ODBC\ODBCINST.INI\MySQL ODBC 5.2w Driver"; ValueType: "string"; ValueName: "Setup"; ValueData: "C:\\Program Files\\MySQL\\Connector ODBC 5.2\\Unicode\\myodbc5S.dll"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\ODBC\ODBCINST.INI\MySQL ODBC 5.2w Driver"; ValueType: "dword"; ValueName: "UsageCount"; ValueData: "00000001"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\ODBC\ODBCINST.INI\ODBC Core"; ValueType: "none"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\ODBC\ODBCINST.INI\ODBC Core"; ValueType: "dword"; ValueName: "UsageCount"; ValueData: "00000001"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"; ValueType: "string"; ValueName: "MySQL ODBC 5.2w Driver"; ValueData: "Installed"; Flags: uninsdeletekey
Root: HKCU; Subkey: "SOFTWARE\Seagate Software\Crystal Reports\DatabaseOptions\ODBC\OuterJoin\"; ValueType: "string"; ValueName: "SQL2OuterJoin"; ValueData: "libmyodbc5w"; Flags: uninsdeletekey

[Run]
Filename: "{app}\Component\Microsoft Visual Basic 6.0 Runtime\Runtime Visual Basic 6.0.exe"; Description: "Menginstal Microsoft Visual Basic 6.0 Runtime Destribution"; StatusMsg: "Menginstal Microsoft Visual Basic 6.0 Runtime Destribution"; Flags: skipifsilent
Filename: "{app}\Component\Crystal Report Redistribution 8.5 Runtime\setup.exe"; Description: "Menginstal Crystal Report 8.5 Runtime Destribution"; StatusMsg: "Menginstal Crystal Reports 8.5"; Flags: skipifsilent
Filename: "msiexec.exe"; Parameters: "/i ""{app}\Component\MySQL Connector ODBC 5.2.2\MySQL Connector ODBC 5.2.2.msi"; Description: "Menginstal MySQL Connector ODBC 5.2.2"; StatusMsg: "Menginstall MySQL Connector ODBC 5.2.2";
Filename: "{app}\Panduan\Panduan.pdf"; Description: "Baca Panduan Penggunaan"; Flags: postinstall shellexec skipifsilent
Filename: "{app}\{#MyAppExeName}"; Description: "Jalankan Panda Pustaka"; Flags: nowait postinstall skipifsilent
