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
			input = MsgBox("'" & Email.To & "' is not a valid email address.", vbCritical+vbRetryCancel, "VBSendMail")
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
			input = MsgBox("'" & Email.From & "' is not a valid email address.", vbCritical+vbRetryCancel, "VBSendMail")
			If input = vbRetry Then
				valid = False
			Else
				WScript.Quit
			End If
		End If
	Loop
valid = False		
Set re = Nothing

Email.Subject = InputBox("Subject:", "VBSendMail") 'Subject

Email.TextBody = InputBox("Body:", "VBSendMail") 'Body

Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

EmailDomain = Split(Split(Email.To, """")(1), "@")(1)

Set oShell = WScript.CreateObject("WScript.Shell")
        Do Until success = True
		Set sOutput = oShell.Exec("nslookup -q=mx " & EmailDomain)
		sOutputStd = sOutput.StdOut.ReadAll
		If InStr(sOutputStd, "mail exchanger") = False Then
		input = MsgBox("Error: MX record lookup failed." & vbCrLf & "Domain: " & EmailDomain & vbCrLf & "Description: " & sOutput.StdErr.ReadAll, vbCritical+vbRetryCancel, "VBSendMail")
			If input = vbRetry Then
				success = False
			Else
				WScript.Quit
			End If
        Else
			success = True
		End If
	Loop
success = False
SMTP = Split(Split(sOutputStd, "mail exchanger = ")(1), vbCrLf)(0)
Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTP

Email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

Email.Configuration.Fields.Update

On Error Resume Next
Email.Send
Do until success = True
If Err.Number <> 0 Then
input = MsgBox("Error: Email send failed." & vbCrLf & "To: " & Split(Email.To, """")(1) & vbCrLf & "From: " & Split(Email.From, """")(1) & vbCrLf & "Description: " & Err.Description & vbCrLf & "Note: Most likely the SMTP server has rejected your email due to your IP address not being whitelisted by spam protection as an SMTP server. You can try again or alternatively get your IP address whitelisted by the recipient's SMTP server.", vbCritical+vbRetryCancel, "VBSendMail")
			If input = vbRetry Then
				success = False
			Else
				WScript.Quit
			End If
End If
Loop
success = True
success = False

MsgBox "Email sent successfully.", vbInformation, "VBSendMail"