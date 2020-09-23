Dim wsError, objWS
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

server = "localhost"


Set objWS = WScript.CreateObject("WinsckW.WinSock", "objWS_")
objWS.Protocol = sckTCPProtocol
objWS.Connect server, 189


Do
      Wscript.Sleep 50
      Timeout = Timeout + 1
      If Timeout > 40 then trapError
Loop Until objWS.State = 7 

Wscript.Echo "Connection with server is established!..." & vbCrLf

str = "Hello There...."
Wscript.Echo "Sending: " & "<Hello There....>"
objWS.SendData str
Wscript.Sleep 1000

str = "What is going On...."
Wscript.Echo "Sending: " & "<What is going On....>"
objWS.SendData str
Wscript.Sleep 1000

str = "Bye now...."
Wscript.Echo "Sending: " & "<Bye now....>"
objWS.SendData str
Wscript.Sleep 1000

Sub objWS_Error(Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
       wsError = "Error connecting... Number: " & Number & " Description: " & Description
End Sub

Sub trapError

     If not wsError = "" then 
           Wscript.Echo wsError
     Else: Wscript.Echo "Cannot Connect.. Unknown Error."
     End If
     objWS.CloseCon
     Wscript.Quit

End Sub
