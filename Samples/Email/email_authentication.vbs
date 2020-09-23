
Option Explicit

Dim objWS, strRespCode, smtpHost, smtpPort, sendFrom, sendTo
Dim strSubject, strMessage, nameTo, nameFrom, ErrorTrap, strUC
Dim StateAuth, UserID, Password

On Error Resume Next
Set objWS = WScript.CreateObject("WinsckW.WinSock", "objWS_")
If Err.number <> 0 then Wscript.Echo Err.Number & ", " & _
                             Err.Description & ", " & _
                             Err.Source: Wscript.Quit
On Error Goto 0


'========================== Replace Variables ==========================


smtpHost = "server"
smtpPort = 25 
sendFrom = "email@domain.net"
sendTo = "email@domain.net"
strSubject = "SUBJECT"
strMessage = "MESSAGE"
nameTo = "RECIPIENT_NAME"
nameFrom = "SENDER_NAME"
userID = objWS.encode("User_ID")
Password = objWS.encode("Password")


'========================== End Replace Variables ======================


With objWS
 .Protocol = 0
        .Connect smtpHost, smtpPort
 If CaptureResponse("220") = False Then trapError
 Wscript.Echo "Connecting....  OK?"
 .SendData "EHLO " & userID & vbCrLf

 If CaptureResponse("250") = False Then trapError
 Wscript.Echo "Requesting Login...."
 .SendData "AUTH LOGIN" & vbCrLf

StateAuth = True
 If CaptureResponse("334") = False Then trapError
 Wscript.Echo "Sending User Name....  OK?"
 .SendData userID & vbCrLf

 If CaptureResponse("334") = False Then trapError
 Wscript.Echo "Sending Password....  Authenticated?"
 .SendData Password & vbCrLf

StateAuth = False

 If CaptureResponse("235") = False Then trapError
 Wscript.Echo "Sending Sender Email....  OK?"
 .SendData "MAIL FROM:" & sendFrom & vbCrLf

 If CaptureResponse("250") = False Then trapError
 Wscript.Echo "Sending Recipient Address....  OK?"
        .SendData "RCPT TO:" & sendTo & vbCrLf


 If CaptureResponse("250") = False Then trapError
 Wscript.Echo "Request Permission to send Data....  OK?"
        .SendData "DATA" & vbCrLf


 If CaptureResponse("354") = False Then trapError
Wscript.Echo "Permission is Granted! ...."

 .SendData "To:" & nameTo & " <" & sendTO & ">" & vbCrLf
        .SendData "From:" & nameFrom & " <" & sendFrom & ">" & vbCrLf
        .SendData "Subject:" & strSubject & vbCrLf & vbCrLf
        .SendData strMessage & vbCrLf
 .SendData "." & vbCrLf
 If CaptureResponse("250") = False Then trapErrort
 Wscript.Echo "Data Sent!...."
 .SendData "QUIT" & vbCrLf
 If CaptureResponse("221") = False Then trapError
 Wscript.Echo "Quiting.... Bye Bye..."
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

Select Case StateAuth
	Case False 
		If Not strRespCode = "" And CStr(strRespCode) = respCode Then
  			strRespCode = ""
  			CaptureResponse = True
                	Exit Function
        	End If
	Case True
		If CStr(strRespCode) = respCode Then
  			strRespCode = ""
  			CaptureResponse = True
                	Exit Function
		End If
End Select


Timeout = Timeout + 1
Loop Until Timeout > 80
CaptureResponse = False
End Function

Sub trapError
	
	If not ErrorTrap = "" then 
		Wscript.Echo ErrorTrap
	Else:
		Wscript.Echo strUC
	End If
	objWS.CloseCon
	Wscript.Quit

End Sub

Sub objWS_Error(Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
 ErrorTrap = "Unable to send Mail. Error: " & Number & _
                     ", Source: " & Source & ", Description: " & Description

End Sub
 

