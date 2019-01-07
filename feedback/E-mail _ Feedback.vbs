strsubject = UserInput( "Please Send Us your feedback, your feedback is very much appriciated for our improvement!:" )

strbody = UserInput( "Please describe in more detail:" )

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

MyEmail.Subject="[IT Guidebook]FeedBack from user"
MyEmail.From="vanlong.truong@tetrapak.com"
MyEmail.To="vanlong.truong@tetrapak.com"
MyEmail.CC="thai.ha@tetrapak.com"
MyEmail.TextBody="Dear OSS team," &vbCrLf& vbCrLf& "This is the feedback form user, Please take a look." &vbCrLf&vbCrLf&  "Feedback General : " &vbCrLf &strsubject &vbCrLf& "Feedback in detail: " &vbCrLf &strbody &vbCrLf& vbCrLf& "Thank you so much" &vbCrLf& "Best Regards,"
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

'SMTP Server
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="mail.tetrapak.com"

'SMTP Port
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 

MyEmail.Configuration.Fields.Update
MyEmail.Send

set MyEmail=nothing