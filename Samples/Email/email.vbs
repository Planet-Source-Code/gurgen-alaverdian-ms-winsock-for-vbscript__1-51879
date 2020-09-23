


Option Explicit
Dim objWS, strRespCode, smtpHost, smtpPort, sendFrom, sendTo
Dim strSubject, strMessage, nameTo, nameFrom, ErrorTrap, strUC

On Error Resume Next
Set objWS = WScript.CreateObject("WinsckW.WinSock", "objWS_")
If Err.number <> 0 then msgbox Err.Number & ", " & _
                             Err.Description & ", " & _
                             Err.Source: Wscript.Quit
On Error Goto 0

'**** Replace Variables

smtpHost = "server"
smtpPort = 25   'change port if behind proxy
sendFrom = "email@email.com"
sendTo = "email@email.com"
strSubject = "Subject"
strMessage = "Message"
nameTo = "Recipient_Name"
nameFrom = "Sender_Name"

'***********************

With objWS
 .Protocol = 0
        .Connect smtpHost, smtpPort
 If CaptureResponse("220") = False Then trapError
Wscript.Echo objWS.State
 Wscript.Echo "Connecting....  OK"
 .SendData "HELO " & smtpHost & vbCrLf
 If CaptureResponse("250") = False Then trapError
 Wscript.Echo "Greet Server....  OK"
 .SendData "MAIL FROM:" & sendFrom & vbCrLf
 If CaptureResponse("250") = False Then trapError
 Wscript.Echo "Send Sender Address....  OK"
        .SendData "RCPT TO:" & sendTo & vbCrLf
 If CaptureResponse("250") = False Then trapError
 Wscript.Echo "Send Recipient Address....  OK"
        .SendData "DATA" & vbCrLf
 If CaptureResponse("354") = False Then trapError
 Wscript.Echo "Permission to send Data?....  OK"
 .SendData "To:" & nameTo & " <" & sendTO & ">" & vbCrLf
        .SendData "From:" & nameFrom & " <" & sendFrom & ">" & vbCrLf
        .SendData "Subject:" & strSubject & vbCrLf & vbCrLf
        .SendData strMessage & vbCrLf
 .SendData "." & vbCrLf
 If CaptureResponse("250") = False Then trapErrort
 Wscript.Echo "Data Sent.... OK"
 .SendData "QUIT" & vbCrLf
 If CaptureResponse("221") = False Then trapError
 Wscript.Echo "Quiting.... OK"
 .CloseCon
End With

Sub objWS_DataArrival(bytesTotal)
Dim I
strUC = objWS.GetData (strUC)
Wscript.Echo strUC
strRespCode = Left(strUC, 3)
End Sub

Function CaptureResponse(respCode)
Dim Timeout

Timeout = 1
Do
 Wscript.Sleep 50
     If Not strRespCode = "" And CStr(strRespCode) = respCode Then
  strRespCode = ""
  CaptureResponse = True
                Exit Function
        End If
 Timeout = Timeout + 1
Loop Until Timeout > 80
CaptureResponse = False
End Function

Sub trapError
	
	If not ErrorTrap = "" then 
		msgbox ErrorTrap
	Else:
		msgbox strUC
	End If
	objWS.CloseCon
	Wscript.Quit

End Sub

Sub objWS_Error(Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
 ErrorTrap = "Unable to send Mail. Error: " & Number & _
                     ", Source: " & Source & ", Description: " & Description

End Sub
 

