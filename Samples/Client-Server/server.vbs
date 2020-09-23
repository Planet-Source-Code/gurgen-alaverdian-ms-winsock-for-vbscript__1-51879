Dim strData, objWS

sckClosed = 0
sckClosing = 8
sckConnected = 7
sckConnecting = 6
sckConnectionPending = 3
sckError = 9
sckHostResolved = 5
sckListening = 2
sckOpen = 1
sckResolvingHost = 4

sckTCPProtocol = 0
sckUDPProtocol = 1


Set objWS = WScript.CreateObject("WinsckW.WinSock", "objWS_")

objWS.Protocol = sckTCPProtocol

objWS.CloseCon
objWS.LocalPort = 189
objWS.Listen
Wscript.Echo "Listening port 189..."

Do 
   Wscript.Sleep 100
Loop While objWS.State = sckListening

If objWS.State = sckConnected Then
   Wscript.Echo "Connection with Client is established!..." & vbCrLf 


   Do 
      If not strData = Empty then
         Wscript.Echo "Receiving: " & strData
         strData = Empty
      End If
      Wscript.Sleep 100
   Loop While objWS.State = sckConnected
End If

objWS.CloseCon

Sub objWS_ConnectionRequest(requestID)

      If objWS.State <> sckClosed then
            objWS.CloseCon
            objWS.Accept requestID
      End If

End Sub

Sub objWS_DataArrival(bytesTotal)
      strData = objWS.GetData (strData)
End Sub
