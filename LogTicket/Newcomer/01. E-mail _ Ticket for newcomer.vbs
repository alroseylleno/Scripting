struseraccount = UserInput( "User's NT account:" )
struseremail= UserInput( "User's E-mail address:" )
struseremail1= UserInput( "Line Manager's E-mail address:" )
strJabber= UserInput( "Permission to use Jabber? (YES or NO):" )
struserextension= UserInput( "User's Extension number:" )
strwebex= UserInput( "Permission to HOST a WebEx Meeting ? (YES or NO):" )
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
MyEmail.From="hien.nguyen@tetrapak.com"
MyEmail.cc="vanlong.truong@tetrapak.com"
MyEmail.Bcc=struseremail1
MyEmail.To="ITservice@tetrapak.com"
MyEmail.TextBody="Dear Global Service Desk," &vbCrLf& vbCrLf& "Kindly help granting access for user to these following applications:" &vbCrLf&vbCrLf& "NT Account: " &struseraccount &vbCrLf& "E-mail: " &struseremail &vbCrLf& "Line Manager's E-mail: " &struseremail1 &vbCrLf& "IP extension number: " &struserextension &vbCrLf&vbCrLf& "Access Jabber: " &strJabber &vbCrLf&  "Access VPN Junos Pulse Secure: " &strVPN  &vbCrLf& "Access WebEx permission to host a meeting: " &strwebex &vbCrLf&vbCrLf& "Best Regards," &vbCrLf&"OSS VN HOCHIMINH" &vbCrLf&"VanLong Truong"


MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

'SMTP Server
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="mail.tetrapak.com"

'SMTP Port
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 

MyEmail.Configuration.Fields.Update
MyEmail.Send

set MyEmail=nothing
