powercfg /batteryreport /output "%userprofile%\desktop\Battery_report.html"
start %userprofile%\desktop\Battery_Report.html
timeout /t 2
del %userprofile%\desktop\Battery_Report.html 