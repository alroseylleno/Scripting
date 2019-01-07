robocopy "\\VNHCFPS01\TPV\A6_IT\Shared\Software\Unikey" "%UserProfile%\Desktop\Unikey" /mir

@echo off

set SCRIPT="%TEMP%\%RANDOM%-%RANDOM%-%RANDOM%-%RANDOM%.vbs"

echo Set oWS = WScript.CreateObject("WScript.Shell") >> %SCRIPT%
echo sLinkFile = "%USERPROFILE%\Desktop\Unikey.lnk" >> %SCRIPT%
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> %SCRIPT%
echo oLink.TargetPath = "%UserProfile%\Desktop\Unikey\UnikeyNT.exe" >> %SCRIPT%
echo oLink.Save >> %SCRIPT%

cscript /nologo %SCRIPT%
del %SCRIPT%

msg * Hello user: %username% from PC: %Computername% - The problem is resolved successfully.
msg * Please make a copy of this file to your local computer for next use. If the error persist then please contact IT Global Service Desk for further support ! - Thank you - VN OSS Ho Chi Minh