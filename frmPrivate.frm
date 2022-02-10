VERSION 5.00
Object = "{2120D62E-1B94-47CE-956E-F31CED1DA6C4}#18.7#0"; "aircutils.ocx"
Begin VB.Form frmPrivate 
   Caption         =   "Chat"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrivate.frx":0000
   LinkTopic       =   "frmPrivate"
   MDIChild        =   -1  'True
   MousePointer    =   1  'Arrow
   ScaleHeight     =   5160
   ScaleWidth      =   7140
   Visible         =   0   'False
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
Attribute VB_Name = "frmPrivate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LastAscii As Integer
Dim ScrollHistory() As Variant
Dim ScrollHistoryMax As Integer
Dim ScrollHistoryCurrent As Integer

Public LogNum As Long
Public LogOpen As Boolean

Public ServerNum As Integer

Public WindowNum As Integer

Private Sub Form_Activate()
    Set fActive = Me
    frmMain.WSwitch.ActWnd Me
    ActiveServer = Me.ServerNum
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
    If Not Me.Tag = "" Then UnloadPrivateWnd WindowNum
End Sub

Private Sub KeyFetch1_ChangeWindow(ByVal WindowNum As Long)
    frmMain.WSwitch.NumWnd WindowNum
End Sub

Private Sub LogBox_Click(ByVal MouseButton As Integer, ByVal X As Long, ByVal Y As Long)
    Select Case MouseButton
        Case vbRightButton
            ClickNick = Me.Tag
            Me.PopupMenu frmMain.privPopup ', , X, Y
    End Select
End Sub

Private Sub LogBox_DblClick()
    WhoisColl.IsCollecting = True
    PutServ "WHOIS " & Me.Tag & " " & Me.Tag
End Sub

Private Sub LogBox_Change()
    If Not Me Is fActive Then frmMain.WSwitch.ColWnd Me, vbRed
End Sub

Private Sub LogBox_GotFocus()
    txtInput.SetFocus
End Sub

Private Sub LogBox_OLEDragDrop(ByRef Data As DataObject, ByRef Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim C As Long
    If Data.GetFormat(vbCFFiles) Then
        Effect = 4
        On Error Resume Next
        Do
            Inc C
            If Data.Files(C) = "" Then Exit Do
            If Err <> 0 Then Exit Do
            NewDCCWnd Me.Tag, Data.Files(C), FileLen(Data.Files(C)), DCCIP, 0, True
        Loop Until Err <> 0
        On Error GoTo 0
    Else
        Effect = 0
    End If
End Sub

Private Sub LogBox_OLEDragOver(ByRef Data As DataObject, ByRef Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, ByRef State As Integer)
    If Data.GetFormat(vbCFFiles) Then
        Effect = 4
    Else
        Effect = 0
    End If
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
                Parse V(C), Me.Tag
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
        ScrollHistoryCurrent = 0
        Inc ScrollHistoryMax
        ReDim Preserve ScrollHistory(1 To ScrollHistoryMax)
        For C = ScrollHistoryMax To 2 Step -1
            ScrollHistory(C) = ScrollHistory(C - 1)
        Next
        ScrollHistory(1) = txtInput
        Parse txtInput, Me.Tag
        If Me.Tag = "" Then Unload Me: Exit Sub
        txtInput = ""
        KeyAscii = 0
    ElseIf KeyAscii = 9 Then 'Tab
        If txtInput = "" Then Exit Sub
        Dim V As Variant
        Dim S As String
        V = Split(txtInput, " ")
        S = V(UBound(V))
        If Len(S) >= Len(Me.Tag) Then Exit Sub
        If LCase(S) = LCase(Left(Me.Tag, Len(S))) Then
            If S = V(0) Then
                txtInput = Me.Tag
            Else
                ReDim Preserve V(LBound(V) To UBound(V) - 1)
                txtInput = Merge(V, 0) & " " & Me.Tag
            End If
            txtInput.SetFocus
            SendKeys "{END}"
        End If
        KeyAscii = 0
    End If
End Sub

