strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
objWMIService.Security_.Privileges.AddAsString "SeLoadDriverPrivilege", True

Install "10.143.193.237"
sub Install(strIP)
    InstallPrinterPort strIP
end Sub


strBasePrinter = "Temp Office - A4 Printer "
strPrinterName = "HP LaserJet Enterprise 700 color MFP M775 PCL6 Class Driver"
strIPPort = "IP_10.143.193.237"
Set objShell = CreateObject("WScript.Shell")
strCommand = "cmd /c rundll32 printui.dll,PrintUIEntry /if /b """ & strBasePrinter & """ /r """ & strIPPort & """ /m """ & strPrinterName & """ & /Z"
objShell.Run strCommand, 1, True


Sub InstallPrinterPort(strIP)
   
    Set colInstalledPorts =  objWMIService.ExecQuery _
        ("Select Name from Win32_TCPIPPrinterPort")
    For each objPort in colInstalledPorts
        If objPort.Name="IP_" & strIP then exit sub 
    Next
  
    Set objNewPort = objWMIService.Get _
        ("Win32_TCPIPPrinterPort").SpawnInstance_
    objNewPort.Name = "IP_" & strIP
    objNewPort.Protocol = 1
    objNewPort.HostAddress = strIP
    objNewPort.PortNumber = "9100"
    objNewPort.SNMPEnabled = False
    objNewPort.Put_
end Sub
