VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{2120D62E-1B94-47CE-956E-F31CED1DA6C4}#18.7#0"; "aircutils.ocx"
Begin VB.Form frmChannel 
   Caption         =   "Channel"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9225
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChannel.frx":0000
   LinkTopic       =   "frmChannel"
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   9225
   Begin MSComctlLib.ListView listNick 
      Height          =   4335
      Left            =   5040
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   7646
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   13816530
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nick"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Timer timerIgnoreDCC 
      Enabled         =   0   'False
      Interval        =   10000
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   7646
   End
   Begin aircutils.KeyFetch KeyFetch1 
      Height          =   465
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   820
   End
End
Attribute VB_Name = "frmChannel"
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

Public WindowNum As Integer
Public HasParted As Boolean
Public ServerNum As Integer

Dim TabPosition As Integer
Dim TabList() As Variant

Function FindNickPos(ByVal Nick As String) As Long
    Dim C As Long
    Nick = TrimMode(Nick)
    With listNick.ListItems
        For C = 1 To .Count
            If Nick = TrimMode(.Item(C).Text) Then Exit For
        Next
        If C > .Count Then Exit Function
        FindNickPos = C
    End With
End Function

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
    listNick.ListItems.Clear
    Me.Left = 0
    Me.Top = 0
    Me.Width = frmMain.ScaleWidth
    Me.Height = frmMain.ScaleHeight
End Sub

Private Sub Form_Resize()
    If ScaleHeight < txtInput.Height Then Exit Sub
    If ScaleWidth < listNick.Width Then Exit Sub
    If DisplayInfo.ShowNicklist Then
        If ScaleHeight < txtInput.Height Then Exit Sub
        LogBox.Width = ScaleWidth + 30 - listNick.Width
        listNick.Height = ScaleHeight - txtInput.Height + 20
        listNick.Left = LogBox.Width - 30
    Else
        LogBox.Width = ScaleWidth + 30
    End If
    LogBox.Height = ScaleHeight - txtInput.Height + 20
    txtInput.Top = LogBox.Height - 20
    txtInput.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not Me.Tag = "" Then UnloadChannelWnd WindowNum, HasParted
End Sub

Private Sub KeyFetch1_ChangeWindow(ByVal WindowNum As Long)
    frmMain.WSwitch.NumWnd WindowNum
End Sub

Private Sub listNick_DblClick()
    Dim F As frmPrivate
    If listNick.SelectedItem Is Nothing Then Exit Sub
    If Not PrWnd(TrimMode(listNick.SelectedItem)) = 0 Then
        Set F = PrivateWnd(PrWnd(TrimMode(listNick.SelectedItem)))
        frmMain.WSwitch.ActWnd F
        F.SetFocus
    Else
        With Nicklist(WindowNum)
            NewPrivateWnd TrimMode(listNick.SelectedItem), .User_Host(.UserPos(TrimMode(listNick.SelectedItem))), True
        End With
    End If
End Sub

Private Sub listNick_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then listNick_DblClick
End Sub

Private Sub listNick_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set listNick.SelectedItem = listNick.HitTest(X, Y)
End Sub

Private Sub listNick_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If listNick.SelectedItem Is Nothing Then Exit Sub
    With Nicklist(WindowNum)
        If Button = 2 Then
            If Not listNick.SelectedItem Is Nothing Then ClickNick = TrimMode(listNick.SelectedItem)
            ClickChan = Me.Tag
            If IsIRCOP Then EnableIRCOP Else DisableIRCOP
            If .IsOp(.UserPos(StatusWnd(ActiveServer).CurrentNick)) Then
                EnableOp
                ReadyBans ClickNick, .User_Host(.UserPos(StatusWnd(ActiveServer).CurrentNick))
            Else
                DisableOp
            End If
            Me.PopupMenu frmMain.nickPopup
        End If
    End With
End Sub

Private Sub listNick_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim C As Long
    Set listNick.SelectedItem = listNick.HitTest(X, Y)
    If listNick.SelectedItem Is Nothing Then
        Effect = 0
    Else
        If Data.GetFormat(vbCFFiles) Then
            listNick.DropHighlight = Nothing
            Effect = 4
            On Error Resume Next
            Do
                Inc C
                If Data.Files(C) = "" Then Exit Do
                If Err <> 0 Then Exit Do
                NewDCCWnd TrimMode(listNick.SelectedItem), Data.Files(C), FileLen(Data.Files(C)), DCCIP, 0, True
            Loop Until Err <> 0
            On Error GoTo 0
        Else
            Effect = 0
        End If
    End If
End Sub

Private Sub listNick_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Set listNick.SelectedItem = listNick.HitTest(X, Y)
    listNick.DropHighlight = listNick.SelectedItem
    If listNick.SelectedItem Is Nothing Then
        Effect = 0
    Else
        If Data.GetFormat(vbCFFiles) Then
            Effect = 4
        Else
            Effect = 0
        End If
    End If
End Sub

Private Sub EnableIRCOP()
    frmMain.nickPopupIRCOPKill.Visible = True
    frmMain.nickPopupIRCOPKLine.Visible = True
End Sub

Private Sub EnableOp()
    frmMain.nickPopupOpMode.Visible = True
    frmMain.nickPopupOpBan.Visible = True
    frmMain.nickPopupOpKick.Visible = True
End Sub

Private Sub DisableIRCOP()
    frmMain.nickPopupIRCOPKill.Visible = False
    frmMain.nickPopupIRCOPKLine.Visible = False
End Sub

Private Sub DisableOp()
    frmMain.nickPopupOpMode.Visible = False
    frmMain.nickPopupOpBan.Visible = False
    frmMain.nickPopupOpKick.Visible = False
End Sub

Private Sub ReadyBans(Nick As String, Hostmask As String)
    Dim C As Long
    For C = 0 To 9
        frmMain.nickPopupBanMode(C).Caption = UserHostMode(Nick, Hostmask, C)
    Next
End Sub

Private Sub LogBox_Change()
    If Not Me Is fActive Then frmMain.WSwitch.ColWnd Me, vbRed
End Sub

Private Sub LogBox_GotFocus()
    txtInput.SetFocus
End Sub

Private Sub timerIgnoreDCC_Timer()
    timerIgnoreDCC.Enabled = False
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
    Dim DeleteTab As Boolean
    If KeyAscii = LastAscii Then Exit Sub
    LastAscii = KeyAscii
    DeleteTab = True
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
        If Tag = "" Then Unload Me: Exit Sub 'form was unloaded
        txtInput = ""
        KeyAscii = 0
    ElseIf KeyAscii = 9 Then 'Tab
        If txtInput = "" Then Exit Sub
        Dim V As Variant
        Dim S As String
        Dim T As String
        V = Split(txtInput, " ")
        S = V(UBound(V))
        If LCase(S) = LCase(Left(Me.Tag, Len(S))) Then
            If S = V(0) Then
                txtInput = Me.Tag
            Else
                ReDim Preserve V(LBound(V) To UBound(V) - 1)
                txtInput = Merge(V, 0) & " " & Me.Tag
            End If
            SendKeys "{END}"
            DeleteTab = False
            KeyAscii = 0
            Exit Sub
        End If
        If TabPosition = 0 Then
            GenerateTabList S
        End If
        T = NextTab
        If TabPosition > 0 Then
            If S = V(0) Then
                txtInput = T
            Else
                ReDim Preserve V(LBound(V) To UBound(V) - 1)
                txtInput = Merge(V, 0) & " " & T
            End If
            SendKeys "{END}"
            DeleteTab = False
        End If
        KeyAscii = 0
    End If
    If DeleteTab Then DeleteTabList
End Sub

Sub GenerateTabList(ByVal S As String)
    Dim C As Long
    Dim D As Long
    With Nicklist(WindowNum)
        For C = 1 To Nicklist(WindowNum).Count
            If Len(.User_Nick(C)) >= Len(S) Then
                If LCase(Left(.User_Nick(C), Len(S))) = LCase(S) Then 'Legg til
                    Inc D
                    ReDim Preserve TabList(1 To D)
                    TabList(D) = .User_Nick(C)
                End If
            End If
        Next
    End With
End Sub

Sub DeleteTabList()
    TabPosition = 0
    Erase TabList
End Sub

Function NextTab() As String
    TabPosition = TabPosition + 1
    If IsEmpty(TabList) Then Exit Function
    On Error Resume Next
    If TabPosition > UBound(TabList) Then TabPosition = 1
    If Err <> 0 Then
        On Error GoTo 0
        Err.Clear
        TabPosition = 0
        Exit Function
    End If
    On Error GoTo 0
    If TabPosition > UBound(TabList) Then TabPosition = 0: Exit Function
    NextTab = TabList(TabPosition)
End Function
