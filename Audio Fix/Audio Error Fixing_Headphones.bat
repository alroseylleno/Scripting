@echo off
if not "%1"=="am_admin" (powershell start -verb runas '%0' am_admin & exit)
net stop audiosrv
net stop AudioEndpointBuilder
net start audiosrv
net start AudioEndpointBuilder
Nircmd setdefaultsounddevice "Headphone" 1
msg * Hello user: %username% from PC: %Computername% - The problem is resolved successfully.
msg * Please make a copy of this file to your local computer for next use. If the error persist then please contact IT Global Service Desk for further support ! - Thank you - VN OSS Ho Chi Minh