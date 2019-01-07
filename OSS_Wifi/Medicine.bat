@echo off
echo -------------CLICK "YES" NOW !!-------------
timeout /t 5
if not "%1"=="am_admin" (powershell start -verb runas '%0' am_admin & exit)
echo -------------------------------------------------------
echo -------------------------------------------------------
echo WARNING: YOU ARE ABOUT TO DISCONNECT THE OSSVN WIFI ---
echo -------------------------------------------------------
echo -------------------------------------------------------
timeout /t 5
echo 1.Delete wireless network profile
netsh wlan delete profile "OSS VN"
echo ------------Done------------

echo 2.Reconfigurate IP Address:
netsh int ip set address "Wi-Fi" DHCP
echo 3.DHCP service enabled on IP Address
echo ------------Done------------
 
netsh int ip set dnsservers "Wi-Fi" DHCP
echo 4.DHCP service enabled on DNS Server
echo ------------Done------------

echo 5.Clearing profile cache. 
ipconfig /flushdns
echo ------------Done------------
del %userprofile%\desktop\OSSVN.bat

echo All done.
echo Closing the application. 
timeout /t 5




	
