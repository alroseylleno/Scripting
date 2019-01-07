@echo off
title OSS VN HOCHIMINH
color 8
if not "%1"=="am_admin" (powershell start -verb runas '%0' am_admin & exit)
cd\
cd program files (x86)\bit9\parity agent
dascli allowuninstall 1
net stop "cb protection agent"
net start "cb protection agent"
echo OSS VN HOCHIMINH.
echo All Done - Press any key to close.
pause