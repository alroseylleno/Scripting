@echo off
 echo\
 echo --------------------------------------------------
 echo ------SYSTEM INFORMATION - TETRA PAK VIETNAM------
 echo --------------------------------------------------
 echo\  
 echo 1.USERNAME: %UserName%.
 echo --------------------------------
 echo 2.COMPUTER NAME: %ComputerName%.
 echo --------------------------------
 echo 3.IPv4 Address is : 
	ipconfig | find "IPv4"
 echo --------------------------------
 echo 4.IPv6 Address is : 
	ipconfig | find "IPv6"
 echo --------------------------------
 echo 5.Network: 
 echo --------------------------------
	ipconfig | find "Suffix"
 echo --------------------------------
systeminfo | findstr /c:"Original Install Date" 
systeminfo | findstr /c:"System Boot Time" 
systeminfo | findstr /c:"Time Zone" 
 echo --------------------------------
systeminfo | findstr /c:"Domain" 
 echo --------------------------------
systeminfo | findstr /c:"Logon Server"
 echo --------------------------------
systeminfo | findstr /c:"Registered Organization" 
 echo --------------------------------
systeminfo | findstr /c:"OS Name" 
systeminfo | findstr /c:"OS Version" 
systeminfo | findstr /c:"OS Manufacturer" 
systeminfo | findstr /c:"System Manufacturer" 
systeminfo | findstr /c:"Total Physical Memory" 
 echo --------------------------------
systeminfo | findstr /c:"Network Card(s)" 
 echo --------------------------------
 echo\
 echo Press the Space bar to close this window.
 pause > nul