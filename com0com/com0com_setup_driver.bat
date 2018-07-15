::The master copy is in MD\Ais\NmeaRouterSource\com0com
::it is transferred by Inno Setup on install to the location below
CD %ProgramFiles%\Arundale\NmeaRouter\com0com
:: DisableUACforAdmin Vista/7 see http://www.howtogeek.com/howto/windows-vista/disable-user-account-controluac-for-administrators-only/
reg add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v ConsentPromptBehaviorAdmin /t REG_DWORD /d 0 /f
:: stop add new hardware wizard
reg add HKLM\Software\Policies\Microsoft\Windows\DeviceInstall\Settings /v SuppressNewHWUI /t REG_DWORD /d 1 /f
setup.exe /S /D=%ProgramFiles%\Arundale\NmeaRouter\com0com
reg add HKLM\Software\Policies\Microsoft\Windows\DeviceInstall\Settings /v SuppressNewHWUI /t REG_DWORD /d 0 /f
reg add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v ConsentPromptBehaviorAdmin /t REG_DWORD /d 2 /f
:: quit here on installation
