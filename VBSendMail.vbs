Set Email = CreateObject("CDO.Message")

Set re = New RegExp
         With re
            .Pattern = "(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|""(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*"")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9]))\.){3}(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9])|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])"
         End With
        Do Until valid = True
		Email.To = InputBox("To:", "VBSendMail") 'To
         If re.Test(Email.To) = True Then
			valid = True
		Else
			input = MsgBox("'To' must be an email address.", vbCritical+vbRetryCancel, "VBSendMail")
			If input = vbRetry Then
				valid = False
			Else
				WScript.Quit
			End If
		End If
	Loop
        valid = False
		
Do Until valid = True
Email.From = InputBox("From:", "VBSendMail") 'From
         If re.Test(Email.From) = True Then
			valid = True
		Else
			input = MsgBox("'From' must be an email address.", vbCritical+vbRetryCancel, "VBSendMail")
			If input = vbRetry Then
				valid = False
			Else
				WScript.Quit
			End If
		End If
	Loop
        
Set re = Nothing

Email.Subject = InputBox("Subject:", "VBSendMail") 'Subject

Email.TextBody = InputBox("Body:", "VBSendMail") 'Body

Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

EmailDomain = Split(Split(Email.To, """")(1), "@")(1)

Set oShell = WScript.CreateObject("WScript.Shell")
sOutput = oShell.Exec("nslookup -q=mx " & EmailDomain).StdOut.ReadAll
sOutputStripped = Split(sOutput, "mail exchanger = ")(1)
SMTP = Split(sOutputStripped, vbCrLf)(0)
Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTP

Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

Email.Configuration.Fields.Update

Email.Send

MsgBox "Email sent successfully.", vbInformation, "VBSendMail"