stremail = UserInput( "Enter your e-mail:" )

strsubject = UserInput( "Which Program/Hardware are you having issues with? :" )

strbody = UserInput( "Please describe your issue in detail:" )

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

MyEmail.Subject=strsubject
MyEmail.From=stremail
MyEmail.To="ITservice@tetrapak.com"
MyEmail.CC="vanlong.truong@tetrapak.com"
MyEmail.TextBody="Dear GSD team," &vbCrLf& vbCrLf& "Thank you so much for being helpful as always!." &vbCrLf&vbCrLf&  "Which Program/Hardware are you having issues with : " &vbCrLf &strsubject &vbCrLf& "Please describe your issue: " &vbCrLf &strbody &vbCrLf& vbCrLf& "Thank you so much" &vbCrLf& "Best Regards," &vbCrLf &stremail
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

'SMTP Server
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="mail.tetrapak.com"

'SMTP Port
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 

MyEmail.Configuration.Fields.Update
MyEmail.Send

set MyEmail=nothing