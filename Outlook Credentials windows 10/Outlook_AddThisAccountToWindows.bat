@echo off
if not "%1"=="am_admin" (powershell start -verb runas '%0' am_admin & exit)
cmdkey.exe /list > "%TEMP%\List.txt"
findstr.exe Target "%TEMP%\List.txt" > "%TEMP%\tokensonly.txt"
FOR /F "tokens=1,2 delims= " %%G IN (%TEMP%\tokensonly.txt) DO cmdkey.exe /delete:%%H
del "%TEMP%\List.txt" /s /f /q
del "%TEMP%\tokensonly.txt" /s /f /q
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Identity" /v DisableADALatopWAMOverride /t REG_DWORD /d 1
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 16
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 32
ipconfig /flushdns
ipconfig /registerdns
netsh winsock reset
netsh int ipv4 reset reset catalog
gpupdate /force
msg * Hello user: %username% from PC: %Computername% - The problem is resolved successfully.
msg * Please make a copy of this file to your local computer for next use. If the error persist then please contact IT Global Service Desk for further support ! - Thank you - VN OSS Ho Chi Minh

