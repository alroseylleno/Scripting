@echo off
echo -------------CLICK "YES" NOW !!-------------
timeout /t 5
if not "%1"=="am_admin" (powershell start -verb runas '%0' am_admin & exit)
echo ---------------------------
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
pause





	
