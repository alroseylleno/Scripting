@echo off
Color 08
set loopcount=100
:loop
Color 08
echo %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% %random% 
set /a loopcount=loopcount-1
if %loopcount%==0 goto exitloop
goto loop
:exitloop
@echo off
if not "%1"=="am_admin" (powershell start -verb runas '%0' am_admin & exit)
cd\
cd program files (x86)\bit9\parity agent
dascli allowuninstall 1
net stop "cb protection agent"
echo OSS VN HOCHIMINH
echo All Done - Press any key to close.
pause