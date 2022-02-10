VERSION 5.00
Object = "{2120D62E-1B94-47CE-956E-F31CED1DA6C4}#19.0#0"; "aircutils.ocx"
Begin VB.Form frmStatus 
   Caption         =   "Status"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStatus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmStatus"
   MDIChild        =   -1  'True
   MousePointer    =   1  'Arrow
   ScaleHeight     =   7530
   ScaleWidth      =   9540
   Tag             =   "Status"
   Begin VB.Timer tmrChkLag 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer tmrLag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer timerIdle 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      BackColor       =   &H00D2D2D2&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      Top             =   4320
      Width           =   6855
   End
   Begin aircutils.LogBox2 LogBox 
      Height          =   4335
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7646
   End
   Begin aircutils.KeyFetch KeyFetch1 
      Height          =   465
      Left            =   6960
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   820
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents IRC As CSocket
Attribute IRC.VB_VarHelpID = -1
Public WithEvents Ident As CSocket
Attribute Ident.VB_VarHelpID = -1

Private SCmd As String

Dim LastAscii As Integer
Dim ScrollHistory() As Variant
Dim ScrollHistoryMax As Integer
Dim ScrollHistoryCurrent As Integer

Public IsLag As Boolean 'Lagcounter enabled

Public LogNum As Long
Public LogOpen As Boolean
Public ServerNum As Integer
Public StatusLocked As Boolean

'####### IRCSTATUS VARIABLES #####
Public iServerName As String
Public IdleTime As Long
Public LagTime As Long
Public CurrentNick As String
Public AwayReason As String
Public ModeString As String
'#################################

Public HasQuit As Boolean '..?

Dim DiscReason As String

Dim uMode As Integer

Public AutoJoinChannels As String

Public Property Let ServerName(ByVal nServerName As String)
    iServerName = nServerName
End Property

Public Property Get ServerName() As String
    '
    'Below is failsafe if socket hasn't been initialized
    'If IRC Is Nothing Then ServerName = iServerName: Exit Property
    If IRC Is Nothing Then Exit Property
    '
    If iServerName = "" Then
        If IRC.RemoteHost = "" Then
            If IRC.RemoteHostIP = "" Then
                ServerName = IRCInfo.Server
            Else
                ServerName = IRC.RemoteHostIP
            End If
        Else
            ServerName = IRC.RemoteHost
        End If
    Else
        ServerName = iServerName
    End If
    '
End Property

Public Property Get IsOpen() As Boolean
    '
    'Returns TRUE if socket state is 0, or if socket hasn't been initialized
    '
    If IRC Is Nothing Then IsOpen = True: Exit Property
    If (IRC.State = 0) Or (IRC.State = 9) Then IsOpen = True
    '
End Property

Public Property Get IsConnected() As Boolean
    '
    'Returns TRUE if socket state is 7
    '
    If IRC Is Nothing Then Exit Property
    If IRC.State = 7 Then IsConnected = True
    '
End Property

Public Property Get IsRegistered() As Boolean
    '
    'Returns TRUE if connection to the server has been registered
    '
    If IRC Is Nothing Then Exit Property
    If IRC.State <> sckConnected Then Exit Property
    If CurrentNick <> "" Then IsRegistered = True
    '
End Property

Public Function TermSocket() As Boolean
    '
    'Terminates the socket
    '
    IRC.CloseSocket
    CloseIdent
    Call IRC_OnClose
    IRC.vbSocket
    '
End Function


Private Sub Ident_OnConnectionRequest(ByVal requestID As Long)
    Ident.CloseSocket
    Ident.Accept requestID
    If Not Ident.RemoteHostIP = IRC.RemoteHostIP Then
        Output "Ident request from wrong host, closing...", Me, statusc, True
        Ident.CloseSocket
    End If
End Sub

Private Sub Ident_OnDataArrival(ByVal bytesTotal As Long)
    Dim C As Long
    Dim s As String * 1
    Dim SCmd As String
    Dim SGet As String
    Ident.GetData SGet
    For C = 1 To Len(SGet)
        s = Mid(SGet, C, 1)
        If s = vbLf Or s = vbCr Then
            If Not SCmd = "" Then
                Output "Ident request: " & SCmd, Me, statusc, True
                Output "Ident response: " & SCmd & " : USERID : UNIX : " & IRCInfo.Ident, Me, statusc, True
                Ident.SendData SCmd & " : USERID : UNIX : " & IRCInfo.Ident & vbCrLf
                SCmd = ""
            End If
        Else
            SCmd = SCmd & s
        End If
    Next
End Sub

Sub CloseIdent(Optional ByVal Reason As String)
    Dim s As String
    If Ident.State = 0 Then Exit Sub 'Already closed
    Ident.CloseSocket
    If Not Reason = "" Then
        s = " (" & Reason & ")."
    Else
        s = "."
    End If
    Output "Ident server closed" & s, Me, statusc, True
End Sub

Private Sub Ident_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    Ident.CloseSocket
End Sub

Private Sub Ident_OnSendComplete()
    CloseIdent
End Sub

Private Sub IRC_OnClose()
    If DiscReason <> "" Then
        DiscReason = ": " & DiscReason
        If CurrentNick = "" Then 'Connection failed
            Output "*** Connection failed" & DiscReason, fActive, statusc
        Else
            OutputA "*** Disconnected" & DiscReason, CurrentNick, Me, statusc
        End If
    End If
    timerIdle.Enabled = False
    tmrChkLag.Enabled = False
    tmrLag.Enabled = False
    SendToScripts "disconnect", IRC.RemoteHost, IRC.RemotePort, ServerNum
    IdleTime = 0
    LagTime = 0
    CurrentNick = ""
    AwayReason = ""
    ModeString = ""
    With frmMain.IRCStatus
        .Reset
        .ChangeIdle ShortenTime(0)
        .Changelag ShortenTime(0)
    End With
End Sub

Private Sub IRC_OnConnect()
    HasQuit = False
    SendToScripts "connect", IRC.RemoteHost, IRC.RemotePort, ServerNum
    Output "*** Connected!", Me, statusc
    frmMain.IRCStatus.ChangeServer ServerName
    Me.Tag = ServerName
    Me.Caption = "Status: " & ServerName
    frmMain.WSwitch.Refresh
    timerIdle.Enabled = True
    With IRCInfo
        PutServ "NICK " & .Nick
        PutServ "USER " & .Ident & " " & (IIf(IRCInfo.ModeWallops, 2, 0) + IIf(IRCInfo.ModeInvisible, 4, 0)) & " * :" & .Realname
    End With
End Sub

Private Sub IRC_OnDataArrival(ByVal bytesTotal As Long)
    If Not IsConnected Then Exit Sub
    Dim C As Long
    Dim s As String * 1
    Dim SGet As String
    IRC.GetData SGet
    For C = 1 To Len(SGet)
        s = Mid(SGet, C, 1)
        If s = vbLf Or s = vbCr Then
            If s = vbCr And Mid(SGet, C + 1, 1) = vbLf Then C = C + 1
            If Not SCmd = "" Then ParseSrv SCmd, ServerNum
            SCmd = ""
        Else
            SCmd = SCmd & s
        End If
    Next
End Sub

Private Sub IRC_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    DiscReason = Description
    CloseIdent "timeout"
    Call IRC_OnClose
    IRC.vbSocket
End Sub

Private Sub KeyFetch1_ChangeWindow(ByVal WindowNum As Long)
    frmMain.WSwitch.NumWnd WindowNum
End Sub

Private Sub LogBox_Change()
    If Not Me Is fActive Then frmMain.WSwitch.ColWnd Me, vbRed
End Sub

Private Sub LogBox_GotFocus()
    txtInput.SetFocus
End Sub

Private Sub timerIdle_Timer()
    Inc IdleTime
    If ((AwayInfo.AAUse) And (IdleTime = CDbl(AwayInfo.AAMinutes) * 60)) Then
        If AwayReason = "" Then
            AwayReason = AwayInfo.AAMsg
            PutServ "AWAY :" & AwayInfo.AAMsg, ServerNum
            SendToScripts "autoaway", AwayReason
            If ServerNum = ActiveServer Then
                frmMain.IRCStatus.ChangeAway AwayInfo.AAMsg
            End If
        End If
    End If
    If ServerNum = ActiveServer Then frmMain.IRCStatus.ChangeIdle ShortenTime(CDbl(IdleTime))
End Sub

Private Sub timerITO_Timer()
    CloseIdent "timeout"
End Sub

'##### WINSOCK/IDENT STUFF OVER THIS LINE ######

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    uMode = UnloadMode
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not Me.Tag = "" Then UnloadStatusWnd ServerNum, uMode
End Sub

Private Sub Form_Activate()
    frmMain.WSwitch.ActWnd Me
    Set fActive = Me
    ActiveServer = ServerNum
    With frmMain.IRCStatus
        .ChangeServer ServerName
        .ChangeAway AwayReason
        .ChangeModes ModeString
        .ChangeNick CurrentNick
        .ChangeIdle ShortenTime(IdleTime)
        .Changelag ShortenTime(LagTime)
    End With
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    Me.Width = frmMain.ScaleWidth
    Me.Height = frmMain.ScaleHeight
    'Form_Activate
    If LagNewStatus Then IsLag = True
End Sub

Private Sub Form_Resize()
    If ScaleHeight < txtInput.Height Then Exit Sub
    LogBox.Width = ScaleWidth + 30
    LogBox.Height = ScaleHeight - txtInput.Height + 20
    txtInput.Top = LogBox.Height - 20
    txtInput.Width = ScaleWidth
End Sub

Private Sub LogBox_DblClick()
    PutServ "LUSERS"
End Sub

Private Sub tmrChkLag_Timer()
    If Not IsLag Then Exit Sub
    If tmrLag.Enabled Then 'High lag time
        Output "Warning: " & SC_Fill(IRC.RemoteHost) & " lag time is now " & BoldCode & _
               ShortenTime(LagTime + 1) & BoldCode & "!", fActive, statusc, True
    Else 'Check lag
        StartLagCount ServerNum
    End If
End Sub

Private Sub tmrLag_Timer()
    If Not IsLag Then Exit Sub
    LagTime = LagTime + 1
    If ServerNum = ActiveServer Then frmMain.IRCStatus.Changelag ShortenTime(LagTime)
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    LastAscii = 0
    If ((Shift = vbCtrlMask And 7) And ((KeyCode = vbKeyV) Or (KeyCode = Asc("v")))) Then
        Dim V As Variant
        Dim C As Long
        V = TrimCrLf_Out(Clipboard.GetText)
        If UBound(V) = -1 Then Exit Sub
        If UBound(V) > LBound(V) Then 'Multiline
            If UBound(V) - LBound(V) > 3 Then
                If MsgBox("Warning: paste " & UBound(V) - LBound(V) + 1 & " lines?", vbExclamation + vbYesNo, "Warning") = vbNo Then
                    KeyCode = 0
                    Shift = 0
                    Exit Sub
                End If
            End If
            For C = LBound(V) To UBound(V)
                V(C) = TrimCrLf(V(C))
                If Not Left(V(C), 1) = "/" Then V(C) = "/" & V(C)
                Parse V(C)
            Next
        End If
        KeyCode = 0
        Shift = 0
        Exit Sub
    End If
    If ChkFunction(KeyCode) Then
        KeyCode = 0
        Shift = 0
        Exit Sub
    End If
    If KeyCode = vbKeyUp Then
        If ScrollHistoryCurrent = ScrollHistoryMax + 1 Then txtInput = "": Exit Sub
        Inc ScrollHistoryCurrent
        If ScrollHistoryCurrent = ScrollHistoryMax + 1 Then txtInput = "": Exit Sub
        txtInput = ScrollHistory(ScrollHistoryCurrent)
        txtInput.SetFocus
        SendKeys "{END}"
    ElseIf KeyCode = vbKeyDown Then
        If ScrollHistoryCurrent <= 0 Then txtInput = "": Exit Sub
        Dec ScrollHistoryCurrent
        If ScrollHistoryCurrent <= 0 Then txtInput = "": Exit Sub
        txtInput = ScrollHistory(ScrollHistoryCurrent)
        txtInput.SetFocus
        SendKeys "{END}"
    End If
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = LastAscii Then Exit Sub
    LastAscii = KeyAscii
    If KeyAscii = 11 Then 'Ctrl+K
        If Not txtInput.SelLength = 0 Then
            txtInput.SelText = ColorCode & txtInput.SelText & ColorCode
        Else
            txtInput.SelText = ColorCode
        End If
        KeyAscii = 0
    ElseIf KeyAscii = 2 Then 'Ctrl+B
        If Not txtInput.SelLength = 0 Then
            txtInput.SelText = BoldCode & txtInput.SelText & BoldCode
        Else
            txtInput.SelText = BoldCode
        End If
        KeyAscii = 0
    ElseIf KeyAscii = 21 Then 'Ctrl+U
        If Not txtInput.SelLength = 0 Then
            txtInput.SelText = UnderlineCode & txtInput.SelText & UnderlineCode
        Else
            txtInput.SelText = UnderlineCode
        End If
        KeyAscii = 0
    ElseIf KeyAscii = 18 Then 'Ctrl+R
        If Not txtInput.SelLength = 0 Then
            txtInput.SelText = "^R" & txtInput.SelText & ReverseCode
        Else
            txtInput.SelText = "^R"
        End If
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        Dim C As Long
        If txtInput = "" Then Exit Sub
        'txtInput = IIf(Left(txtInput, 1) = "/", "", "/") & txtInput
        ScrollHistoryCurrent = 0
        Inc ScrollHistoryMax
        ReDim Preserve ScrollHistory(1 To ScrollHistoryMax)
        For C = ScrollHistoryMax To 2 Step -1
            ScrollHistory(C) = ScrollHistory(C - 1)
        Next
        ScrollHistory(1) = txtInput
        Parse txtInput
        If Me.Tag = "" Then Unload Me: Exit Sub
        txtInput = ""
        KeyAscii = 0
    End If
End Sub
