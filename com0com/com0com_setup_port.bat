::By default, the com0com excutable files are installed into Program Files\com0com
:: This is how to set up a port CNCA0 <> COM14 # means a free port
:: In Device Manager
:: CNCA0 will be under com0com device emulators
:: COM14 will be under Ports(COM & LPT)
:: next line required if run from the command prompt as opposed to ShellExecute
CD %ProgramFiles%\Arundale\NmeaRouter\com0com
:: sto add new hardawre wizard
reg add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v ConsentPromptBehaviorAdmin /t REG_DWORD /d 0 /f
reg add HKLM\Software\Policies\Microsoft\Windows\DeviceInstall\Settings /v SuppressNewHWUI /t REG_DWORD /d 1 /f
:: --no-update ensures virtual com ports all updated at the ned
setupc.exe --no-update install PortName=- PortName=COM#
:: must update if wizard is suppressed
setupc.exe update
reg add HKLM\Software\Policies\Microsoft\Windows\DeviceInstall\Settings /v SuppressNewHWUI /t REG_DWORD /d 0 /f
reg add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v ConsentPromptBehaviorAdmin /t REG_DWORD /d 2 /f
:: enable pause to debug messages
:: pause
::setupc.exe listfnames

