Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "%windir%\system32\control.exe /name Microsoft.DefaultPrograms /page pageDefaultProgram \pageAdvancedSettings?pszAppName=Internet%20Explorer"
WScript.Sleep 1200
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys " "
WshShell.SendKeys "{TAB}"
WshShell.SendKeys " "
WScript.Quit					 	 