; -- AisDecoder.iss --

;SourcePath is where the .iss file is located
#pragma message SourcePath
#define MyAppName "NmeaRouter.exe" 
#pragma message "MyAppName info: " + MyAppName
#define MyAppFile SourcePath + MyAppName
#pragma message "MyAppFile info: " + MyAppFile
#define MyAppVersion GetFileVersion(MyAppFile)
#pragma message "Detailed version info: " + MyAppVersion
#define MyAppVersion StringChange(MyAppVersion, ".0.", "." )
#pragma message "Stripped version info: " + MyAppVersion

#define public MyFileDateTimeString GetFileDateTimeString(MyAppFile, 'dd/mm/yyyy hh:nn:ss', '-', ':');
#pragma message "File Date info: " + MyFileDateTimeString
#define MyDateTimeString GetDateTimeString('dd/mm/yyyy hh:nn:ss', '-', ':');

#define MyProgramData "C:\ProgramData"
#pragma message "MyProgramData: " + MyProgramData
#define MySys32 "C:\Windows\SysWOW64"
#pragma message "MySys32: " + MySys32


;#define ch FileOpen("c:\website\backup.bat")
;#define batcommand FileRead(ch)
;#define batcommand StringChange(batcommand, "Version", MyAppVersion )

; this works #define result Exec('cmd /c xcopy/s/y/q', '"e:\My Documents\Ais\NmeaRouterSource" "e:\My Documents\Ais\NmeaRouter_Backup\NmeaRouter_1.1.7\"')
#define result Exec('cmd /c xcopy/s/y/q', '"C:\Users\Admin\Documents\Ais\NmeaRouterSource" "C:\Users\Admin\Documents\Ais\NmeaRouterSourceBackup\NmeaRouter_' + MyAppVersion + '\"')
;Copy CommonSouce files into NmeaRouter backup as we may have changed a common routine
#define result Exec('cmd /c xcopy/s/y/q', '"C:\Users\Admin\Documents\Ais\CommonSource" "C:\Users\Admin\Documents\Ais\CommonSourceBackup\Common_' + MyAppVersion + '\"')
;#Define result Exec('cmd /c dir/p', '"e:\My Documents"')

[Setup]
;version explorer displays for setup.exe, recovered with VB6 app.major & app.minor
VersionInfoVersion={#MyAppVersion}
;minimum windows version sofware will run on (0=no Win98, 4.0= nt or 2000,XP upwards)
;MinVersion= 4.10,4.0
MinVersion= 0,5.0
AppName=NmeaRouter
AppId=NmeaRouter
;CreateUninstallRegKey=no
;UpdateUninstallLogAppName=no
;On INNO installer "This will install Ais Decoder Version x.x.x.x on your computer"
AppVerName=NmeaRouter
AppPublisher=Neal Arundale
AppPublisherURL=http://arundale.com/docs/ais/nmearouter.html
;where the users files are placed
DefaultDirName={pf}\Arundale\NmeaRouter
DefaultGroupName=NmeaRouter
;UsePreviousAppDir=No
;UninstallDisplayIcon=E:\jna\arundale\website\docs\arundale.ico
UninstallDisplayIcon=router.ico
;outputdir=E:\jna\Arundale\website\docs\ais\
outputdir= "C:\Users\Admin\Documents\DirectNic\Live Parent (ArundaleCom)\docs\ais"
OutputBaseFilename= NmeaRouter_setup_{#MyAppVersion}
setuplogging=yes
SetupIconFile=arundale.ico
;required for vbfiles installation
PrivilegesRequired=admin
LicenseFile=NmeaRouter EULA.rtf
;FileDateTimeString= (#MyFileDateTimeString)
AppMutex="NmeaRouter"

[Dirs]
;only required if creating an empty directory [files] creates the directory
;these get copied to userappdata when AisDecoderns new version
Name: "{userappdata}\Arundale\NmeaRouter"
[Files]
; begin VB system files
;dll'S CANNOT BE IN SYSTEM DIRECTORY
Source: "vbfiles\msvbvm60.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vbfiles\oleaut32.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vbfiles\olepro32.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vbfiles\asycfilt.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile
Source: "vbfiles\stdole2.tlb";  DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: "vbfiles\comcat.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
;Source: "vbfiles\Support\vb6stkit.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vbfiles\msstdfmt.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
;Source: "vbfiles\Support\msvcrt.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vbfiles\scrrun.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver

Source: "{#MySys32}\MSINET.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "{#MySys32}\MSWINSCK.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "{#MySys32}\mscomctl.OCX"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "{#MySys32}\ComDlg32.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "{#MySys32}\MSHFlxGd.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
;Windows 8
Source: "{#MySys32}\richtx32.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver

; end VB system files
Source: "{#MyAppName}"; DestDir: "{app}" ;flags: replacesameversion ignoreversion
;Source: "E:\My Documents\Ais\Decoder_v3\{#MyAppVersion}.txt"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Files" ;flags: replacesameversion ignoreversion
Source: "arundale.ico"; DestDir: "{app}"  ;flags: replacesameversion ignoreversion
Source: "router.ico"; DestDir: "{app}"  ;flags: replacesameversion ignoreversion
Source: "com0com\setup_com0com-3.0.0.0-i386-and-x64-unsigned.exe"; DestDir: "{app}\com0com"  ;flags: replacesameversion ignoreversion
Source: "com0com\setup_com0com_W7_x64_signed.exe"; DestDir: "{app}\com0com"  ;flags: replacesameversion ignoreversion
Source: "com0com\setup_com0com_W7_x86_signed.exe"; DestDir: "{app}\com0com"  ;flags: replacesameversion ignoreversion
Source: "com0com\setup_com0com-2.2.2.0-x64-fre-signed.exe"; DestDir: "{app}\com0com"  ;flags: replacesameversion ignoreversion
;Source: "com0com\com0com_setup_driver.bat"; DestDir: "{app}\com0com"  ;flags: replacesameversion ignoreversion
;Source: "com0com\com0com_setup_port.bat"; DestDir: "{app}\com0com"  ;flags: replacesameversion ignoreversion
;Source: "com0com\com0com_remove_port.bat"; DestDir: "{app}\com0com"  ;flags: replacesameversion ignoreversion

;help
;Source: "Help\NmeaRouter.chm"; DestDir: "{app}\Help"  ;flags: replacesameversion ignoreversion
;above need uncommenting
;Source: "Readme.txt"; DestDir: "{app}"; Flags: isreadme ignoreversion

;Creates the Shortcuts
[Icons]
Name: "{group}\NmeaRouter"; Filename: "{app}\NmeaRouter.exe"; IconFilename:"{app}\router.ico"
; NOTE: Most apps do not need registry entries to be pre-created. If you
; don't know what the registry is or if you need to use it, then chances are
; you don't need a [Registry] section.
;Name: "{userdocs}\Ais Decoder"; Filename: "{userappdata}\Arundale\Ais Decoder\Output"; Flags: foldershortcut ; IconFilename:"{app}\arundale.ico" ;Comment:"AisDecoder Files"

[InstallDelete]
Type: files; Name: "{app}\NmeaRouter.exe"

[Registry]
Root: HKCU; Subkey: "Software\Arundale"; Flags: uninsdeletekeyifempty
Root: HKCU; Subkey: "Software\Arundale\NmeaRouter"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Arundale\NmeaRouter\Profiles"; Flags: uninsdeletekey
;setting for NewVersion (2015) 
Root: HKLM; Subkey: "Software\Arundale\NmeaRouter\Settings"; ValueType: string; ValueName: "LastVersion"; ValueData: "{#MyAppVersion}"

[Run]
;Filename: "{app}\license.txt"; Description: "View the README file"; Flags: postinstall shellexec unchecked skipifsilent
;line below causes error invalid control array index in NmeaRouter
;Filename: "{app}\com0com\setup.exe";  Parameters: "/S /D={app}\com0com\"; Description: "Install Virtual Com Port (VCP) support"; Flags: postinstall nowait skipifsilent
;Filename: "{app}\com0com\setup.exe";  Parameters: "/S /D={app}\com0com\"; StatusMsg: "Installing Virtual Com Port (VCP) driver ..."; Flags: runminimized
;Filename: "{app}\com0com\com0com_setup_driver.bat";  Parameters: "/S /D={app}\com0com\"; StatusMsg: "Installing Virtual Com Port (VCP) driver ..."; Flags: runminimized


Filename: "{app}\NmeaRouter.exe"; Description: "Launch application"; Flags: postinstall nowait skipifsilent

