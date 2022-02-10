VERSION 5.00
Object = "{2120D62E-1B94-47CE-956E-F31CED1DA6C4}#19.0#0"; "aircutils.ocx"
Begin VB.Form frmDCCChat 
   Caption         =   "DCC Chat"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7785
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDCCChat.frx":0000
   LinkTopic       =   "frmDCCChat"
   MDIChild        =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   7785
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
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   820
   End
End
Attribute VB_Name = "frmDCCChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public WithEvents Chat As CSocket
Attribute Chat.VB_VarHelpID = -1

Dim LastAscii As Integer
Dim ScrollHistory() As Variant
Dim ScrollHistoryMax As Integer
Dim ScrollHistoryCurrent As Integer

Public LogNum As Long
Public LogOpen As Boolean

Public DCCProtocol As dccProtocols
Public UniqueID As Long

Public ServerNum As Integer 'fActive
Public WindowNum As Integer

Public IsConnected As Boolean

Private Sub Chat_OnConnectionRequest(ByVal requestID As Long)
    Chat.CloseSocket
    Chat.Accept requestID
    Output "*** Connected!", Me, statusc
    Me.Caption = Me.Tag & " (" & Chat.RemoteHostIP & ":" & Chat.RemotePort & ")"
End Sub

Private Sub Chat_OnClose()
    Chat.CloseSocket
    Output "*** Chat closed!", Me, statusc
End Sub

Private Sub Chat_OnConnect()
    Output "*** Connected!", Me, statusc
End Sub

Private Sub Chat_OnDataArrival(ByVal bytesTotal As Long)
    Dim C As Long
    Dim s As String * 1
    Dim SCmd As String
    Dim SGet As String
    Chat.GetData SGet
    For C = 1 To Len(SGet)
        s = Mid(SGet, C, 1)
        If s = vbLf Or s = vbCr Then
            If s = vbCr And Mid(SGet, C + 1, 1) = vbLf Then C = C + 1
            If Not SCmd = "" Then
                If Left(SCmd, 7) = "ACTION" Then
                    Output "* " & Me.Tag & " " & TrimCTCP(Mid(SCmd, 8)), Me, actionc
                Else
                    PDC SCmd 'Should work
                End If
            End If
            SCmd = ""
        Else
            SCmd = SCmd & s
        End If
    Next
End Sub

Private Sub Chat_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Chat.CloseSocket
    Output "*** Chat closed (Error " & Number & ": " & Description & ")", Me, statusc
End Sub

Private Sub Form_Activate()
    If Me.Tag = "" Then Exit Sub
    frmMain.WSwitch.ActWnd Me
    Set fActive = Me
    With frmMain.IRCStatus
        .ChangeServer StatusWnd(ActiveServer).ServerName
        .ChangeAway StatusWnd(ActiveServer).AwayReason
        .ChangeModes StatusWnd(ActiveServer).ModeString
        .ChangeNick StatusWnd(ActiveServer).CurrentNick
        .ChangeIdle ShortenTime(StatusWnd(ActiveServer).IdleTime)
        .Changelag ShortenTime(StatusWnd(ActiveServer).LagTime)
    End With
End Sub

Private Sub Form_Load()
    Set Chat = New CSocket
    Me.Left = 0
    Me.Top = 0
    Me.Width = frmMain.ScaleWidth
    Me.Height = frmMain.ScaleHeight
End Sub

Private Sub Form_Resize()
    If ScaleHeight < txtInput.Height Then Exit Sub
    LogBox.Width = ScaleWidth + 30
    LogBox.Height = ScaleHeight - txtInput.Height + 20
    txtInput.Top = LogBox.Height - 20
    txtInput.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadChatWnd WindowNum
End Sub

Private Sub LogBox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        ClickNick = Me.Tag
        Set ClickSock = Chat
        Me.PopupMenu frmMain.chatPopup
        Exit Sub
    End If
    txtInput.SetFocus
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
                Chat.SendData V(C) & vbCrLf
                Output "<" & IRCInfo.Nick & "> " & V(C), Me, ownc
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
        If (Chat.State = 0) And (txtInput = "") Then
            If DCCProtocol = dccPassive Then
                Chat.Bind NextDCCPort, DCCIP
                Chat.Listen
                SendCTCP Me.Tag, "DCC CHAT " & Me.Tag & " " & DCCIP & " " & Chat.LocalPort & " " & UniqueID, True
                Output "*** Setting up listening socket...", Me, statusc
            ElseIf DCCProtocol = dccNormal Then
                Debug.Print Chat.RemoteHost & Chat.RemotePort
                Chat.Connect
                Output "*** Attempting to connect...", Me, statusc
            End If
            KeyAscii = 0
            Exit Sub
        ElseIf Not Chat.State = 7 And Not txtInput = "" Then
            If Left(txtInput, 1) = "/" And Not Left(txtInput, 2) = "//" Then
                Parse txtInput, Me.Tag
            Else
                Output "*** Not connected!", Me, statusc
            End If
            txtInput = ""
            KeyAscii = 0
            Exit Sub
        End If
        If txtInput = "" Then KeyAscii = 0: Exit Sub
        ScrollHistoryCurrent = 0
        Inc ScrollHistoryMax
        ReDim Preserve ScrollHistory(1 To ScrollHistoryMax)
        For C = ScrollHistoryMax To 2 Step -1
            ScrollHistory(C) = ScrollHistory(C - 1)
        Next
        ScrollHistory(1) = txtInput
        If Left(txtInput, 4) = "/me " Then
            If Not Chat.State = 7 Then txtInput = "": Exit Sub
            txtInput = Mid(txtInput, 5)
            Chat.SendData "ACTION " & txtInput & "" & vbCrLf
            Output "* " & IRCInfo.Nick & " " & txtInput, Me, actionc
        ElseIf Left(txtInput, 1) = "/" And Not Left(txtInput, 2) = "//" Then
            Parse txtInput, Me.Tag
        Else
            If Left(txtInput, 2) = "//" Then txtInput = Mid(txtInput, 2)
            If Not Chat.State = 7 Then txtInput = "": Exit Sub
            Chat.SendData txtInput & vbCrLf
            Output "<" & IRCInfo.Nick & "> " & txtInput, Me, ownc
        End If
        txtInput = ""
        KeyAscii = 0
    ElseIf KeyAscii = 27 Then
        Unload Me
        KeyAscii = 0
    End If
End Sub

Sub PDC(ByVal s As String)
    Dim V As Variant
    Dim SrvTxt As String
    If Left(s, 1) = "" And Right(s, 1) = "" Then
        SrvTxt = "PRIVMSG " & Me.Tag & " " & s
        SrvTxt = "" & TrimCTCP(SrvTxt) & ""
        V = SplitCmd(SrvTxt)
        CTCPInterpret s, Me.Tag, "", "", SrvTxt, V
    Else
        Output "<" & Me.Tag & "> " & s, Me
    End If
End Sub
