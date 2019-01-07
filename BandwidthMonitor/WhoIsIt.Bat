@echo off
echo\
echo ------------------------------------
echo ----Who is it - OSS VN HoChiMinh---- 
echo ------------------------------------
echo\
SET /P IPorHost= [Ip Address OR Hostname:]
dir /O:D \\%IPorHost%\C$\Users
echo ------------------------------------
echo ----------MAC Address---------------
arp -a %IPorHost% 
echo ------------------------------------
echo ------------------------------------
echo ----------Host name-----------------
tracert %IPorHost%
Pause
