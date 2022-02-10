VERSION 5.00
Object = "{2120D62E-1B94-47CE-956E-F31CED1DA6C4}#19.3#0"; "aircutils.ocx"
Begin VB.Form frmDCCSend 
   Caption         =   "DCC"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDCCSend.frx":0000
   LinkTopic       =   "frmDCCSend"
   MDIChild        =   -1  'True
   ScaleHeight     =   2370
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   Begin aircutils.ProgressBar Progress 
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   1920
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
   End
   Begin VB.Timer timerAuto 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   120
   End
   Begin VB.Timer timerSendspeed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4440
      Top             =   120
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Reject"
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdResume 
      Caption         =   "Resume"
      Height          =   375
      Left            =   2400
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Elapsed:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label txtTimeElapsed 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Left:"
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   18
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label txtTimeLeft 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3360
      TabIndex        =   17
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   735
   End
   Begin VB.Label txtSendspeed 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   15
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label txtStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Label txtBuffered 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label txtReceived 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label txtSent 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label txtFileSize 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label txtNickname 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label txtFilename 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Sent:"
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Buffered:"
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Recieved:"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Nick:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Size:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Filename:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin aircutils.KeyFetch KeyFetch1 
      Height          =   465
      Left            =   120
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   820
   End
End
Attribute VB_Name = "frmDCCSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents DCC As CSocket
Attribute DCC.VB_VarHelpID = -1

Enum dccMsgTypes
    msgSend = 0
    msgReject = 1
    msgResume = 2
    msgAcceptResume = 3
    msgSendAck = 4 'Used for dcc passive
End Enum


'Big thanks goes to Erlend Sommerfelt Ervik (Again) !!

Public FReceived As Long, FSent As Long, FBuffer As Long
Dim ResumePos As Long
Public TimeElapsed As Double
Public TimeLeft As Double
Public SendSpeed As Long
Public PacketSize As Long
Private FNum As Integer

Public FSize As Long
Public FName As String
Public Nick As String
Public S_IP As String
Public S_Port As String
Public WindowNum As Integer
Public ServerNum As Integer
Public DCCLog As String

Public UniqueID As Long
Public DCCProtocol As dccProtocols
Public SendMsgByDCC As Boolean
Public MDSock As CSocket
Public OldP As String

Public IsSender As Boolean
Public IsReceiver As Boolean
Public DoResume As Boolean

Public maReady As Boolean

Sub DoReceive()
    timerAuto.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    DCC.CloseSocket
    timerSendspeed.Enabled = False
    If cmdCancel.Caption = "Cancel" Then
        txtStatus = dccStatusBroken
        cmdCancel.Caption = "Close"
        Close FNum
        ResetAll
    ElseIf cmdCancel.Caption = "Reject" Then
        DCCSendData msgReject
        Unload Me
    Else
        Unload Me
    End If
End Sub

Private Sub cmdResume_Click()
    DoResumeClick
End Sub

Private Sub cmdSend_Click()
    DoSendClick
End Sub

Sub InitResume()
    DoResume = True
    If IsSender Then
        DCCSendData msgAcceptResume
        txtStatus = dccStatusResumeRequest
    ElseIf IsReceiver Then
        txtStatus = dccStatusReceiving
        If DCCProtocol = dccNormal Then
            DCC.Connect DCC.RemoteHost, DCC.RemotePort
        Else
            DCCSendData msgSendAck
        End If
        DCCLog = DCCLog & "File was resumed at position " & FReceived & "." & vbCrLf
    End If
End Sub

Sub InitPassive(ByVal IP As String, ByVal Port As Long)
    'DCC.CloseSocket
    DCC.Connect IP, Port
End Sub

Private Sub DCC_OnClose()
    If FReceived = FSize Then
        txtStatus = dccStatusFinished
        DCCLog = DCCLog & "Entire file received, DCC transfer OK." & vbCrLf
        txtBuffered = ShortenBytes(0)
    Else
        txtStatus = dccStatusBroken
        DCCLog = DCCLog & "File partially received [" & FReceived & "/" & FSize & "], file transfer broken." & vbCrLf
    End If
    DCC.CloseSocket
    Close FNum
    cmdCancel.Caption = "Close"
    timerSendspeed.Enabled = False
    ResetAll
    If IsReceiver Then
        cmdSend.Caption = "Open"
        cmdSend.Enabled = True
    End If
    If ((DCCInfo.AutoAccept) And (IsReceiver)) Then Unload Me
End Sub

Private Sub DCC_OnConnect()
    ResumePos = FReceived
    If DCCProtocol = dccNormal Then
        FNum = FreeFile
        Open DCCInfo.DownloadDir & TrimPath(FName, True) For Binary As FNum
        If LOF(FNum) > 0 Then
            If DoResume Then
                Seek FNum, LOF(FNum) + 1
            Else
                Close FNum
                Kill DCCInfo.DownloadDir & TrimBad(FName)
                FNum = FreeFile
                Open DCCInfo.DownloadDir & TrimBad(FName) For Binary As FNum
            End If
        End If
        DCCLog = DCCLog & "Connected to host, writing to file " & DCCInfo.DownloadDir & FName & vbCrLf
        txtStatus = dccStatusReceiving
    Else 'Passive, we are sender
        txtStatus = dccStatusSending
        FNum = FreeFile
        Open FName For Binary As FNum
        If DoResume Then
            If FReceived >= FSize Then
                Close FNum
                txtStatus = dccStatusFinished
                cmdCancel.Caption = "Close"
                DCC.CloseSocket
                timerSendspeed.Enabled = False
                If ((DCCInfo.AutoAccept) And (IsReceiver)) Then Unload Me
                Exit Sub
            End If
            Seek FNum, FReceived + 1
        End If
        SendPacket
    End If
    cmdCancel.Caption = "Cancel"
    timerSendspeed.Enabled = True
End Sub

Private Sub DCC_OnConnectionRequest(ByVal requestID As Long)
    On Error GoTo exs
    Dim B() As Byte
    ResumePos = FReceived
    DCC.CloseSocket
    DCC.Accept requestID
    timerSendspeed.Enabled = True
    If DCCProtocol = dccNormal Then 'Passive, we are sender
        txtStatus = dccStatusSending
        cmdCancel.Caption = "Cancel"
        FNum = FreeFile
        Open FName For Binary As FNum
        If DoResume Then
            If FReceived >= FSize Then
                Close FNum
                txtStatus = dccStatusFinished
                cmdCancel.Caption = "Close"
                DCC.CloseSocket
                timerSendspeed.Enabled = False
                If ((DCCInfo.AutoAccept) And (IsReceiver)) Then Unload Me
                Exit Sub
            End If
            Seek FNum, FReceived + 1
        End If
    Else
        FNum = FreeFile
        DCCLog = DCCLog & "Client connected, writing to file " & DCCInfo.DownloadDir & FName & vbCrLf
        Open DCCInfo.DownloadDir & TrimPath(FName, True) For Binary As FNum
        If LOF(FNum) > 0 Then
            If DoResume Then
                Seek FNum, LOF(FNum) + 1
            Else
                Close FNum
                Kill DCCInfo.DownloadDir & TrimBad(FName)
                FNum = FreeFile
                Open DCCInfo.DownloadDir & TrimBad(FName) For Binary As FNum
            End If
        End If
        txtStatus = dccStatusReceiving
    End If
    If Not DCC.State = 7 Then
        DCC.CloseSocket
        Exit Sub
    End If
    If DCCProtocol = dccNormal Then
        SendPacket
    End If
    Exit Sub
exs:
    Close FNum
    Unload Me
End Sub

Private Sub DCC_OnDataArrival(ByVal bytesTotal As Long)
    Dim B() As Byte
    Dim A() As Byte
    Dim C As Long
    If Not DCC.State = 7 Then DoEvents
    If Not DCC.State = 7 Then
        DCC.CloseSocket
        DoEvents
        Exit Sub
    End If
    If IsSender Then
        For C = 1 To bytesTotal \ 4
            Dim GLoc As Long
            DCC.GetData B, vbArray + vbByte, 4
            GLoc = GetLong(CStr(B))
        Next
        FReceived = GLoc
        If FReceived = FSize Then
            txtStatus = dccStatusFinished
            cmdCancel.Caption = "Close"
            DCC.CloseSocket
            timerSendspeed.Enabled = False
            If ((DCCInfo.AutoAccept) And (IsReceiver)) Then Unload Me: Exit Sub
        Else
            FBuffer = FSent - FReceived
            txtBuffered = ShortenBytes(FBuffer)
            If DCCInfo.PumpDCC Then
                If FSent < FReceived Then
                    FSent = FReceived
                Else
                    If (FSent - FReceived <= DCCInfo.SendeBuffer) Then SendPacket
                End If
            Else
                If FReceived = FSent Then
                    SendPacket
                End If
            End If
        End If
    ElseIf IsReceiver Then
        DCC.GetData B, vbByte + vbArray, bytesTotal
        FReceived = FReceived + LenB(CStr(B))
        A = PutLong(FReceived)
        If CStr(A) = "" Then A = PutLong(FReceived)
        DCC.SendData A
        Put FNum, , B
    End If
    With Progress
        .SetRc FReceived
        .SetSn FSent
        .SetBf FBuffer
    End With
    txtReceived = ShortenBytes(FReceived)
    txtBuffered = ShortenBytes(FBuffer)
End Sub

Private Sub DCC_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    DCC.CloseSocket
    If IsReceiver Then
        DCCLog = DCCLog & "File partially received [" & FReceived & "/" & FSize & "], file transfer broken." & vbCrLf
    End If
    txtStatus = dccStatusBroken
    cmdCancel.Caption = "Close"
    timerSendspeed.Enabled = False
    Close FNum
    ResetAll
    If ((DCCInfo.AutoAccept) And (IsReceiver)) Then Unload Me
End Sub

Private Sub DCC_OnSendComplete()
    If FSent >= FSize Then
        Close FNum
        txtStatus = dccStatusFinished
        FSent = FReceived
        
        'txtStatus = dccStatusWaiting
        'DoEvents
    End If
End Sub

Private Sub Form_Activate()
    frmMain.WSwitch.ActWnd Me
End Sub

Private Sub Form_Load()
    Set DCC = New CSocket
    txtSent = ShortenBytes(0)
    txtReceived = ShortenBytes(0)
    txtBuffered = ShortenBytes(0)
    FNum = FreeFile
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then cmdCancel_Click
End Sub

Private Sub Form_Resize()
    If Not WindowState = 0 Then Exit Sub
    Width = 5145
    Height = 2775
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsReceiver Then
        DCCLog = DCCLog & "File was received in " & txtTimeElapsed & ", with average speed of " & txtSendspeed & "." & vbCrLf
        DCCLog = DCCLog & "Window closed, end of log entry."
        LogDCCEvent DCCLog
    End If
    UnloadDCCWnd WindowNum
End Sub

Public Function Percentage() As String
    If FReceived = 0 Then 'Prevent division by zero error
        Percentage = "0"
    Else
        Percentage = Format(100 - (1 - (FReceived / FSize)) * 100, "##")
    End If
    Percentage = Percentage & "%"
    If Percentage = "%" Then Percentage = "0%"
End Function

Private Sub SendPacket(Optional ByVal uPacketSize As Long)
    Dim B() As Byte
    Dim M As Long
    If uPacketSize = 0 Then PacketSize = DCCInfo.SendeBuffer Else PacketSize = uPacketSize
    If FileLen(FName) - FSent < PacketSize Then
        If FileLen(FName) - FSent > 0 Then
            ReDim B(1 To FileLen(FName) - FSent)
        Else
            txtSent = ShortenBytes(FSent)
            Exit Sub
        End If
    Else
        ReDim B(1 To PacketSize)
    End If
    Get FNum, , B
    Inc FSent, LenB(CStr(B))
    DCC.SendData B
    FBuffer = FSent - FReceived
    With Progress
        .SetRc FReceived
        .SetSn FSent
        .SetBf FBuffer
    End With
    txtSent = ShortenBytes(FSent)
    txtReceived = ShortenBytes(FReceived)
    txtBuffered = ShortenBytes(FBuffer)
End Sub

Private Sub KeyFetch1_ChangeWindow(ByVal WindowNum As Long)
    frmMain.WSwitch.NumWnd WindowNum
End Sub

Private Sub timerAuto_Timer()
    'Automatically accepted
    timerAuto.Enabled = False
    If DCCInfo.AutoAccept And IsReceiver Then
        If FReceived > 0 Then 'Resume
            DCCLog = DCCLog & "File was resumed automatically on byte position " & FReceived & " at " & CStr(Now) & vbCrLf
            DoResumeClick
        Else 'Accept
            DCCLog = DCCLog & "File was accepted automatically at " & CStr(Now) & vbCrLf
            DoSendClick
        End If
    End If
End Sub

Private Sub timerSendSpeed_Timer()
    Inc TimeElapsed
    SendSpeed = CLng((FReceived - ResumePos) \ TimeElapsed)
    txtTimeElapsed = ShortenTime(TimeElapsed)
    txtSendspeed = ShortenBytes(SendSpeed) & "/s"
    If Not SendSpeed = 0 Then
        TimeLeft = (FSize - FReceived) \ SendSpeed
        txtTimeLeft = ShortenTime(TimeLeft)
    End If
End Sub

Sub ResetAll()
    If Not IsSender Then Exit Sub
    cmdSend.Caption = "Resend"
    cmdSend.Enabled = True
    txtStatus = dccStatusReadySend
    txtTimeElapsed = ShortenTime(0)
    txtSendspeed = ShortenBytes(0) & "/s"
    txtTimeLeft = ShortenTime(0)
    Progress.SetBf 0
    Progress.SetRc 0
    Progress.SetSn 0
    FReceived = 0
    FSent = 0
    FBuffer = 0
    TimeElapsed = 0
    TimeLeft = 0
    SendSpeed = 0
    PacketSize = 0
    FNum = 0
    DoResume = False
End Sub

Sub DoSendClick()
    On Error GoTo ErrHandle
    Progress.Width = 2895
    cmdResume.Visible = False
    If cmdSend.Caption = "Open" Then
        ShellExecute frmMain.hwnd, vbNullString, DCCInfo.DownloadDir & FName, vbNull, DCCInfo.DownloadDir, 5

        Exit Sub
    End If
    If IsReceiver Then
        If DCCProtocol = dccNormal Then 'Normal dcc (clear)
            DCC.Connect DCC.RemoteHost, DCC.RemotePort
        ElseIf DCCProtocol = dccPassive Then 'Passive dcc
            DCC.CloseSocket
            DCC.vbSocket
            DCC.Bind NextDCCPort, DCCIP
            DCC.Listen
            DCCSendData msgSendAck
            txtStatus = dccStatusPassiveAck
        End If
    ElseIf IsSender Then
        If DCCProtocol = dccNormal Then
            DCC.CloseSocket
            DCC.vbSocket
            DCC.Bind NextDCCPort, DCCIP
            DCC.Listen
            DCCSendData msgSend
            txtStatus = dccStatusSendRequest
        ElseIf DCCProtocol = dccPassive Then
            DCCSendData msgSend
            txtStatus = dccStatusSendRequest
        End If
    End If
    cmdSend.Enabled = False
    cmdCancel.Caption = "Cancel"
    Exit Sub
ErrHandle:
    txtStatus = dccStatusError
    cmdSend.Enabled = False
    cmdCancel.Caption = "Close"
    Err.Clear
    DCC.CloseSocket
    If ((DCCInfo.AutoAccept) And (IsReceiver)) Then Unload Me
End Sub

Sub DoResumeClick()
    If DCCProtocol = dccPassive Then 'Passive dcc
        DCC.CloseSocket
        DCC.Bind , DCCIP
        DCC.Listen
    End If
    Progress.Width = 2895
    cmdResume.Visible = False
    DCCSendData msgResume
    txtStatus = dccStatusResumeSent
    cmdSend.Enabled = False
End Sub

Sub DCCSendData(ByVal L As dccMsgTypes)
    Dim s As String
    If ((DCCProtocol = dccNormal) And (L = msgSend)) Then 'Sender
        s = "DCC SEND """ & TrimPath(FName, True) & """ " & PutIP(DCCIP) & " " & DCC.LocalPort & " " & FileLen(FName)
    ElseIf ((DCCProtocol = dccNormal) And (L = msgReject)) Then 'Receiver
        s = "DCC REJECT """ & TrimPath(FName, True) & """ " & DCC.RemotePort
    ElseIf ((DCCProtocol = dccNormal) And (L = msgResume)) Then 'Receiver
        s = "DCC RESUME """ & TrimPath(FName, True) & """ " & DCC.RemotePort & " " & FileLen(DCCInfo.DownloadDir & FName)
    ElseIf ((DCCProtocol = dccNormal) And (L = msgAcceptResume)) Then 'Sender
        s = "DCC ACCEPT """ & TrimPath(FName, True) & """ " & DCC.LocalPort & " " & FReceived
    
    ElseIf ((DCCProtocol = dccPassive) And (L = msgSend)) Then 'Sender
        s = "DCC SEND """ & TrimPath(FName, True) & """ " & PutIP(DCCIP) & " 0 " & FileLen(FName) & " " & UniqueID
    ElseIf ((DCCProtocol = dccPassive) And (L = msgReject)) Then 'Receiver
        s = "DCC REJECT """ & TrimPath(FName, True) & """ " & DCC.LocalPort & " " & UniqueID
    ElseIf ((DCCProtocol = dccPassive) And (L = msgResume)) Then 'Receiver
        s = "DCC RESUME """ & TrimPath(FName, True) & """ " & DCC.LocalPort & " " & FileLen(DCCInfo.DownloadDir & FName) & " " & UniqueID
    ElseIf ((DCCProtocol = dccPassive) And (L = msgAcceptResume)) Then 'Sender
        s = "DCC ACCEPT """ & TrimPath(FName, True) & """ " & DCC.LocalPort & " " & FReceived & " " & UniqueID
    ElseIf ((DCCProtocol = dccPassive) And (L = msgSendAck)) Then 'Receiver
        s = "DCC SEND """ & TrimPath(FName, True) & """ " & PutIP(DCCIP) & " " & DCC.LocalPort & " " & FSize & " " & UniqueID
        
    Else
        Exit Sub
    End If
    s = "" & s & ""
    If SendMsgByDCC Then
        MDSock.SendData s & vbCrLf
    Else
        PutServ "PRIVMSG " & Nick & " :" & s, ServerNum
        ResetIdle ServerNum
    End If
End Sub

Sub maDCC()
    If Not maReady Then Exit Sub
    ModifyDCC WindowNum, FName, Nick, FSize, FSent, FReceived, txtSendspeed, Percentage, TimeElapsed, _
    TimeLeft, IIf(IsSender, "outgoing", "incoming"), Mid(txtStatus, 9)
End Sub

Private Sub txtBuffered_Change()
    maDCC
End Sub

Private Sub txtSendSpeed_Change()
    maDCC
End Sub

Private Sub txtReceived_Change()
    maDCC
End Sub

Private Sub txtSent_Change()
    maDCC
End Sub

Private Sub txtStatus_Change()
    maDCC
End Sub

Private Sub txtTimeElapsed_Change()
    maDCC
End Sub

Private Sub txtTimeLeft_Change()
    maDCC
End Sub
