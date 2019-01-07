strlinemanager = UserInput( "Enter Line Manager e-mail:" )
struser = UserInput( "Enter User NT account: "&vbCrLf& "(Type ABC if you don't know the NT account)" )

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

MyEmail.Subject="[From IT] Enabling account and initial password for newcomer - " &struser
MyEmail.From="vanlong.truong@tetrapak.com"
MyEmail.To= strlinemanager
MyEmail.CC= "DLGIMOSSVNVietnam@tetrapak.com"
MyEmail.TextBody="Dear Line Manager," &vbCrLf&vbCrLf& "This is an automatic e-mail to remind you of enabling the account of user: " &struser &vbCrLf& "Currently locked out due to the Start Date Policy. As we need to log on to the new account and finish rest of the Computer Installation." &vbCrLf&vbCrLf& "Once you finish the following steps, you will also receive an e-mail from the system: UserProvisioning@tetrapak.com with enclosed NT account and Initial Password – Please forward this mentioned e-mail to us." &vbCrLf&vbCrLf& "1. Go to this website: https://userprovisioning.tetrapak.com/IdentityManagement/default.aspx " &vbCrLf& "2. Search for NT account or Full Name: " &struser &vbCrLf& "(Ex: VNHAT or Ha Thai)" &vbCrLf& "3. Click onto his/her account" &vbCrLf& "4. Switch to tab: Account Management" &vbCrLf& "5. Change from Disable to enable under the Account status" &vbCrLf& "6. Click Save/ Submit/ Confirm at the bottom right." &vbCrLf&vbCrLf& "Let us know if you need any help" &vbCrLf& "Thank you for your cooperation." &vbCrLf& "Best Regards," &vbCrLf& "OSS VIETNAM"

MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

'SMTP Server 
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="mail.tetrapak.com"

'SMTP Port
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 

MyEmail.Configuration.Fields.Update
MyEmail.Send

set MyEmail=nothing