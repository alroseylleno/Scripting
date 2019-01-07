@echo off
echo -------------CLICK "YES" NOW !!-------------
timeout /t 5
if not "%1"=="am_admin" (powershell start -verb runas '%0' am_admin & exit)
echo ---------------------------
SET /P IPorHost= How long will you need the fast connection? [In minute]  
SET /a second= 60 
SET /a Multi= "IPorHost * second"

echo 1.Making a copy of Wireless profile
echo ------------Done------------
copy "\\Vnhcfps01\tpv\A6_IT\Shared\_newcomer\2.Installation\OSS_Wifi\OSSVN.xml" "%UserProfile%\Desktop"

echo 2.Registering OSS Profile 
echo ------------Done------------
netsh wlan add profile filename ="%userprofile%\desktop\OSSVN.xml" interface="Wi-Fi"

echo 3.Delete temp Profile
echo ------------Done------------
del %userprofile%\desktop\OSSVN.xml

echo 4.Reconfigurate IP Address:
echo ------------Done------------
netsh interface ip set address "Wi-Fi" static 192.168.1.150 255.255.255.0 192.168.1.1
netsh interface ip set dnsservers "Wi-Fi" static 8.8.8.8

echo 5.Information Detail:
echo ------------Done------------
netsh int ip show config "Wi-Fi"

echo 5.Ready for high speed connection !!!
timeout /t 10
netsh wlan connect name="OSS VN" ssid="OSS VN" int="Wi-Fi"
echo -------------------------------------------------------
echo ------------------CONGRATULATIONS!!--------------------
echo -------------------------------------------------------
echo Counting down to the disconnect - DO NOT close this windows
timeout /t %Multi%
cls
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





	
