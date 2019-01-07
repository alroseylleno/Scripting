stremail = UserInput( "Enter receiving e-mail:" )

Function UserInput( myPrompt )
 ' This function prompts the user for some input.
 ' When the script runs in CSCRIPT.EXE, StdIn is used,
 ' otherwise the VBScript InputBox( ) function is used.
 ' myPrompt is the the text used to prompt the user for input.
 ' The function returns the input typed either on StdIn or in InputBox( ).

     ' Check if the script runs in CSCRIPT.EXE
     If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
         ' If so, use StdIn and StdOut
         WScript.StdOut.Write myPrompt & " "
         UserInput = WScript.StdIn.ReadLine
     Else
         ' If not, use InputBox( )
         UserInput = InputBox( myPrompt )
     End If
End Function

Set MyEmail=CreateObject("CDO.Message")

MyEmail.Subject="[From IT] Tetra Pak Vietnam System introduction, for user: " &stremail
MyEmail.From="vanlong.truong@tetrapak.com"
MyEmail.To= stremail
MyEmail.TextBody="Dear newcomer," &vbCrLf&vbCrLf& "Hello and very much welcome to Tetra Pak !!" &vbCrLf& "This is an automatic e-mail sending from system for a quick introduction to the Tetra Pak Vietnam IT Tools. Please refer to the attached links below for each and every category, if you cannot open the link or having trouble in reading this e-mail then please contact us directly for help." &vbCrLf&vbCrLf& "IT Global Service Desk: 466 2010 (24/7)" &vbCrLf&vbCrLf& "I.REQUEST" &vbCrLf& "\\VNHCFPS01\TPV\A6_IT\Shared\ITGuidebook.html" &vbCrLf& "1.Software Request:" &vbCrLf& "\\VNHCFPS01\TPV\A6_IT\Shared\_newcomer\1.Manual-Guideline\3.Request\1.Software_Request" &vbCrLf& "2.NT account Request:" &vbCrLf& "\\VNHCFPS01\TPV\A6_IT\Shared\_newcomer\1.Manual-Guideline\3.Request\2.NT_account_request" &vbCrLf& "3.Logon Request:" &vbCrLf& "\\VNHCFPS01\TPV\A6_IT\Shared\_newcomer\1.Manual-Guideline\3.Request\3.Logon_Request" &vbCrLf& "4.Membership Request:" &vbCrLf& "\\VNHCFPS01\TPV\A6_IT\Shared\_newcomer\1.Manual-Guideline\3.Request\4.Membership_Request" &vbCrLf& "5.Create Folder - Distribution List Request:" &vbCrLf& "\\VNHCFPS01\TPV\A6_IT\Shared\_newcomer\1.Manual-Guideline\3.Request\5.Create-Folder-Distribution_List_Request" &vbCrLf&  "6.Hardware Request:" &vbCrLf& "\\VNHCFPS01\TPV\A6_IT\Shared\_newcomer\1.Manual-Guideline\3.Request\6.Hardware_Request" &vbCrLf& "7.Request e-mail on iPhone:" &vbCrLf& "\\VNHCFPS01\TPV\A6_IT\Shared\_newcomer\1.Manual-Guideline\4.TS_mobility_project-iPhone\Airwatch" &vbCrLf& "8.Request for permission to HOST a WebEx meeting:" &vbCrLf& "\\VNHCFPS01\TPV\A6_IT\Shared\_newcomer\1.Manual-Guideline\1.IT_Induction\IT_Service_Sharing\5.Webex,VideoConferencing" &vbCrLf&vbCrLf& "II.ALL ABOUT SMARTCARD" &vbCrLf& "1.How to unblock a smartcard: " &vbCrLf& "\\VNHCFPS01\TPV\A6_IT\Shared\_newcomer\1.Manual-Guideline\6.All_about_SmartCard\Unblock_smartcard" &vbCrLf& "2.How to active smartcard for printer: " &vbCrLf& "\\VNHCFPS01\TPV\A6_IT\Shared\_newcomer\1.Manual-Guideline\2.Printer" &vbCrLf& "3.How to use smartcard: " &vbCrLf& "\\VNHCFPS01\TPV\A6_IT\Shared\_newcomer\1.Manual-Guideline\6.All_about_SmartCard\VPN_pulse_secure" &vbCrLf& "4.How to request a Smart-card with your picture:" &vbCrLf& "\\VNHCFPS01\TPV\A6_IT\Shared\_newcomer\1.Manual-Guideline\3.Request\7.Request_SmartCard" &vbCrLf&vbCrLf& "III.HOW TO:" &vbCrLf& "1.Book a meeting room" &vbCrLf& "\\Vnhcfps01\tpv\A6_IT\Shared\_newcomer\1.Manual-Guideline\1.IT_Induction\IT_Service_Sharing\4.Meeting_Rooms"  &vbCrLf& "2.Contact Global Service Desk - Chat,Call,Raise a Ticket"  &vbCrLf& "\\Vnhcfps01\tpv\A6_IT\Shared\_newcomer\1.Manual-Guideline\1.IT_Induction\Contact_with_IT"  &vbCrLf& "3.Create and Join a WebEx meeting (On iPhone and Laptop)"  &vbCrLf& "\\Vnhcfps01\tpv\A6_IT\Shared\_newcomer\1.Manual-Guideline\1.IT_Induction\IT_Service_Sharing\5.Webex,VideoConferencing\WebEx" &vbCrLf& "4.Create and Join a Video Conference meeting" &vbCrLf& "\\Vnhcfps01\tpv\A6_IT\Shared\_newcomer\1.Manual-Guideline\1.IT_Induction\IT_Service_Sharing\5.Webex,VideoConferencing\VC" &vbCrLf& "5.Report Asset Lost procedure" &vbCrLf& "\\Vnhcfps01\tpv\A6_IT\Shared\_newcomer\1.Manual-Guideline\1.IT_Induction\IT_Service_Sharing\8.Asset_Lost_Reporting_Procedure" &vbCrLf&vbCrLf& "Notes:" &vbCrLf& "_You will not able to access: Jabber, VPN Pulse Secure, IP extension number for the first 3 Days, As Global Service Desk need time to resolved your access permission." &vbCrLf& "Do not attempt to shorten your Smartcard by cutting it as it will be damaged and may not work anymore." &vbCrLf& "_For Tetra Pak Wifi Access: Only for iPhone (using your windows username and password)" &vbCrLf& "_For Internet Guess Access: Contact Receptionist (Ms.Nguyen ThuyHoa - 848 0503)" &vbCrLf&vbCrLf& "Best Regards," &vbCrLf& "OSS VN HOCHIMINH" 

MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

'SMTP Server 
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="mail.tetrapak.com"

'SMTP Port
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 

MyEmail.Configuration.Fields.Update
MyEmail.Send

set MyEmail=nothing