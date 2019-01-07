strOSS = UserInput( "Enter your OSS e-mail:" )
struseraccount = UserInput( "User's NT account:" )
struseremail= UserInput( "User's E-mail address:" )
struseremail1= UserInput( "Line Manager's E-mail address:" )
strJabber= UserInput( "Permission to use Jabber? (YES or NO):" )
struserextension= UserInput( "User's Extension number:" )
strVPN= UserInput( "Permission to access VPN? (YES or NO):" )

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

MyEmail.Subject="Granting access for newcomer: " &struseraccount 
MyEmail.From=strOSS
MyEmail.cc=struseremail
MyEmail.bcc=struseremail1
MyEmail.To="ITservice@tetrapak.com"
MyEmail.TextBody="Dear Global Service Desk," &vbCrLf& vbCrLf& "Kindly help granting access for user to these following applications:" &vbCrLf&vbCrLf& "NT Account: " &struseraccount &vbCrLf& "E-mail: " &struseremail &vbCrLf& "Line Manager's E-mail: " &struseremail1 &vbCrLf&vbCrLf& "-------------------------" &vbCrLf& "Access IP Phone extension: " &struserextension &vbCrLf& "Access Jabber: " &strJabber &vbCrLf&  "Access VPN Junos Pulse Secure: " &strVPN &vbCrLf&"-------------------------" &vbCrLf&vbCrLf& "Best Regards," &vbCrLf&"OSS VN HOCHIMINH" &vbCrLf&"VanLong Truong"
 

MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

'SMTP Server
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="mail.tetrapak.com"

'SMTP Port
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 

MyEmail.Configuration.Fields.Update
MyEmail.Send

set MyEmail=nothing

Set MyEmail2=CreateObject("CDO.Message")

MyEmail2.Subject="[From IT] Tetra Pak Vietnam System introduction, for user: " &struseraccount
MyEmail2.From=strOSS
MyEmail2.cc=struseremail1
MyEmail2.To=struseremail
MyEmail2.Bcc="DLGIMOSSVNVietnam@tetrapak.com"
MyEmail2.TextBody="Dear newcomer:" &vbCrLf&vbCrLf& "Hello and very much welcome to Tetra Pak !!" &vbCrLf& "This is an automatic e-mail sending from system for a quick introduction to the Tetra Pak Vietnam IT Tools. Please refer to the attached links below for the IT Guidebook for the quick access to the documents & instructions (Can ONLY access with Tetra Pak network), if you cannot open the link or having trouble in opening the Guidebook then please contact us directly for help." &vbCrLf& "\\VNHCFPS01\TPV\A6_IT\Shared\ITGuidebook.html" &vbCrLf& "IT Global Service Desk: 466 2010 (24/7)" &vbCrLf& "Direct line:" &vbCrLf& "America:+19 40384 2100" &vbCrLf& "Sweden:+46 4636 2010" &vbCrLf& "Singapore:+65 6890 1911" &vbCrLf&vbCrLf& "First Recommanded setup:" &vbCrLf&"1. Request to DL VN HC Tetra Pak Vietnam e-mail group, follow the instruction of: \\VNHCFPS01\TPV\A6_IT\Shared\_newcomer\1.Manual-Guideline\3.Request\4.Membership_Request\Distribution " &vbCrLf&"2. Request for WebEx host permission (if needed), follow the instruction of: \\VNHCFPS01\TPV\A6_IT\Shared\_newcomer\1.Manual-Guideline\3.Request\1.Software_Request " &vbCrLf&"3. Unblock your smartcard PIN code, follow the instruction of: \\VNHCFPS01\TPV\A6_IT\Shared\_newcomer\1.Manual-Guideline\6.All_about_SmartCard\Unblock_smartcard " &vbCrLf&"4. Active PUK Code for smartcard printing, follow the instruction of: \\VNHCFPS01\TPV\A6_IT\Shared\_newcomer\1.Manual-Guideline\2.Printer " &vbCrLf&vbCrLf& "Special Notes:" &vbCrLf& "- You WILL NOT able to access: Jabber, VPN Pulse Secure, IP extension number for the first 2 Days, As Global Service Desk team need time to resolved your access permission." &vbCrLf& "- DO NOT attempt to shorten your Smartcard by cutting it as it will be damaged and may not work anymore." &vbCrLf& "- For Tetra Pak Wifi Access: Only for iPhone (using your windows username and password)" &vbCrLf& "- For Internet Guess Access: Please contact Receptionist" &vbCrLf&vbCrLf& "Best Regards," &vbCrLf& "OSS VN HOCHIMINH" 

MyEmail2.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

'SMTP Server 
MyEmail2.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="mail.tetrapak.com"

'SMTP Port
MyEmail2.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 

MyEmail2.Configuration.Fields.Update
MyEmail2.Send

set MyEmail2=nothing










