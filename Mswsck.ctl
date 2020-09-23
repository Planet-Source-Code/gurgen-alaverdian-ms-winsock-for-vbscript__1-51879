VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.UserControl Winsock 
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   855
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   885
   ScaleWidth      =   855
   Windowless      =   -1  'True
   Begin MSWinsockLib.Winsock WSCK 
      Left            =   210
      Top             =   210
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Winsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum ProtocolConstants
    sckTCPProtocol = 0
    sckUDPProtocol = 1
End Enum

Public Enum StateConstants
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
End Enum

Public Enum ErrorConstants
    sckAlreadyComplete = &H2735
    sckAlreadyConnected = &H2748
    sckBadState = &H9C46
    sckConnectAborted = &H2745
    sckConnectionRefused = &H274D
    sckConnectionReset = &H2746
    sckGetNotSupported = &H18A
    sckHostNotFound = &H2AF9
    sckHostNotFoundTryAgain = &H2AFA
    sckInProgress = &H2734
    sckInvalidArg = &H9C4E
    sckInvalidArgument = &H271E
    sckInvalidOp = &H9C54
    sckInvalidPropertyValue = &H17C
    sckMsgTooBig = &H2738
    sckNetReset = &H2744
    sckNetworkSubsystemFailed = &H2742
    sckNetworkUnreachable = &H2743
    sckNoBufferSpace = &H2747
    sckNoData = &H2AFC
    sckNonRecoverableError = &H2AFB
    sckNotConnected = &H2749
    sckNotInitialized = &H276D
    sckNotSocket = &H2736
    sckOpCanceled = &H2714
    sckOutOfMemory = 7
    sckOutOfRange = &H9C55
    sckPortNotSupported = &H273B
    sckSetNotSupported = &H17F
    sckSocketShutdown = &H274A
    sckSuccess = &H9C51
    sckTimedout = &H274C
    sckUnsupported = &H9C52
    sckWouldBlock = &H2733
    sckWrongProtocol = &H9C5A
End Enum

Public Event CloseCon()
Public Event Connect()
Public Event ConnectionRequest(requestID As Long)
Public Event DataArrival(bytesTotal As Long)
Public Event Error(Number As Integer, Description As String, _
                    Scode As Long, Source As String, HelpFile As String, _
                    HelpContext As Long, CancelDisplay As Boolean)
Public Event SendComplete()
Public Event SendProgress(bytesSent As Long, bytesRemaining As Long)

Public Property Get Index() As Integer
Index = WSCK.Index
End Property
Public Property Get Tag() As String
Tag = WSCK.Tag
End Property
Public Property Get Name() As String
Name = WSCK.Name
End Property
Public Property Get Object() As Object
Object = WSCK.Object
End Property
Public Property Get Parent() As Object
Parent = WSCK.Parent
End Property
Public Property Get protocol() As ProtocolConstants
protocol = WSCK.protocol
End Property
Public Property Let protocol(pValue As ProtocolConstants)
WSCK.protocol = pValue
End Property
Public Property Get LocalIP() As String
LocalIP = WSCK.LocalIP
End Property
Public Property Get BytesReceived() As Long
BytesReceived = WSCK.BytesReceived
End Property
Public Property Get LocalHostName() As String
LocalHostName = WSCK.LocalHostName
End Property
Public Property Get LocalPort() As Long
LocalPort = WSCK.LocalPort
End Property
Public Property Let LocalPort(portValue As Long)
WSCK.LocalPort = portValue
End Property

Public Property Get RemoteHost() As String
RemoteHost = WSCK.RemoteHost
End Property
Public Property Let RemoteHost(rhValue As String)
WSCK.RemoteHost = rhValue
End Property

Public Property Get RemoteHostIP() As String
RemoteHostIP = WSCK.RemoteHostIP
End Property

Public Property Get RemotePort() As Long
RemotePort = WSCK.RemotePort
End Property
Public Property Let RemotePort(rpValue As Long)
WSCK.RemotePort = rpValue
End Property
Public Property Get SocketHandle() As Long
SocketHandle = WSCK.SocketHandle
End Property
Public Property Get State() As Integer
State = WSCK.State
End Property

Public Sub Accept(ByVal requestID As Long)
WSCK.Accept (requestID)
End Sub

Public Sub Bind(Optional ByVal LocalPort As Long, Optional ByVal LocalIP As String)
WSCK.Bind ByVal LocalPort, ByVal LocalIP
End Sub

Public Sub CloseCon()
WSCK.Close
End Sub

Public Sub Connect(Optional ByVal RemoteHost As String, Optional ByVal RemotePort As Long)
WSCK.Connect ByVal RemoteHost, ByVal RemotePort
End Sub
Public Function encode(ByVal strEncode As String)
encode = Base64Encode(ByVal strEncode)
End Function
Public Function decode(ByVal strDecode As String)
decode = Base64Decode(ByVal strDecode)
End Function
Public Function GetData(ByVal iData As String, Optional ByVal dType, Optional ByVal maxLen)
WSCK.GetData iData, ByVal dType, ByVal maxLen
GetData = iData
End Function
Public Sub Listen()
WSCK.Listen
End Sub

Public Sub PeekData(ByVal Data As String, Optional ByVal dType, Optional ByVal maxLen)
WSCK.PeekData Data, ByVal dType, ByVal maxLen
End Sub

Public Sub SendData(ByVal Data As String)
WSCK.SendData ByVal Data
End Sub

Private Sub WSCK_Close()
RaiseEvent CloseCon
End Sub

Private Sub WSCK_Connect()
RaiseEvent Connect
End Sub

Private Sub WSCK_ConnectionRequest(ByVal requestID As Long)
RaiseEvent ConnectionRequest(requestID)
End Sub

Private Sub WSCK_DataArrival(ByVal bytesTotal As Long)
RaiseEvent DataArrival(bytesTotal)
End Sub

Private Sub WSCK_Error(ByVal Number As Integer, _
                        Description As String, ByVal Scode As Long, _
                        ByVal Source As String, ByVal HelpFile As String, _
                        ByVal HelpContext As Long, CancelDisplay As Boolean)
RaiseEvent Error(Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub

Private Sub WSCK_SendComplete()
RaiseEvent SendComplete
End Sub

Private Sub WSCK_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
RaiseEvent SendProgress(bytesSent, bytesRemaining)
End Sub

