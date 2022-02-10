Attribute VB_Name = "modWindows"
Option Explicit

Public StatusWnd() As frmStatus
Public StatusWndU As Integer

Public PrivateWnd() As frmPrivate
Public ChannelWnd() As frmChannel
Public ChatWnd() As frmDCCChat
Public DCCWnd() As frmDCCSend
Public PrivateWndU As Integer
Public ChannelWndU As Integer
Public ChatWndU As Integer
Public DCCWndU As Integer

Public IgnoreU As Integer

Public ActiveServer As Integer


Sub ConnectNewStatusWnd()
    With IRCInfo
        If .AutoMode = 1 Then 'Show option dialog
            NewStatusWnd
            ShowOptionWnd
            If StatusWnd(StatusWndU).IsOpen Then 'User did not connect
                Unload StatusWnd(StatusWndU)
            End If
        ElseIf .AutoMode = 2 Then 'Auto connect
            If ((.Server = "") Or (.Port = "")) Then 'Show option dialog
                NewStatusWnd
                ShowOptionWnd
                If StatusWnd(StatusWndU).IsOpen Then 'User did not connect
                    Unload StatusWnd(StatusWndU)
                End If
            Else
                NewStatusWnd .Server, .Port
            End If
        ElseIf .AutoMode = 0 Then 'Create new server window
            NewStatusWnd
        End If
    End With
End Sub

Sub NewStatusWnd(Optional ByVal Server As String, Optional ByVal Port As String, Optional ByVal First As Boolean)
    Inc StatusWndU
    ReDim Preserve StatusWnd(1 To StatusWndU)
    Set StatusWnd(StatusWndU) = New frmStatus
    With StatusWnd(StatusWndU)
        .Tag = Server
        If .Tag = "" Then .Tag = "Status"
        .Caption = "Status: " & StatusWndU
        .ServerNum = StatusWndU
        .Visible = True
        Set .Ident = New CSocket
        Set .IRC = New CSocket
        If frmMain.IsBlocking Then
            Inc SavedWndsU
            ReDim Preserve SavedWnds(1 To SavedWndsU)
            Set SavedWnds(SavedWndsU) = StatusWnd(StatusWndU)
        Else
            .Show
            .txtInput.SetFocus
        End If
        .StatusLocked = True
        .timerIdle.Enabled = False
        .tmrLag.Enabled = False
        .tmrChkLag.Enabled = False
    End With
    SetColorWindows StatusWnd(StatusWndU)
    frmMain.WSwitch.AddWnd StatusWnd(StatusWndU), StatusWndU, wndStatus
    Set fActive = StatusWnd(StatusWndU)
    OpenLog logStatus, "Status"
    '########################§§§§§§§§§§§§§§§§§§§§§
    If ((Server = "") And (Port = "")) Then
    ElseIf ((Not Server = "") And (Port = "")) Then
        Port = "6667"
    ElseIf ((Server = "") And (Not Port = "")) Then
        Port = ""
    ElseIf Not IsNumeric(Port) Then
        Port = "6667"
    End If
    ActiveServer = StatusWndU
    InitLogo First, ActiveServer
    If ((Not Server = "") And (Not Port = "")) Then 'Connect now.
        IRCInfo.Server = Server
        IRCInfo.Port = Port
        InitConnect
    End If
End Sub

Sub UnloadStatusWnd(ByVal tServerNum As Integer, Optional ByVal UnloadMode As Integer = 0)
    Dim C As Long
    Dim D As Integer
    Dim oas As Integer
    If tServerNum > StatusWndU Then Exit Sub
    If tServerNum < 1 Then Exit Sub
    If StatusWnd(tServerNum) Is Nothing Then Exit Sub
    oas = ActiveServer
    ActiveServer = 1
    Set fActive = StatusWnd(ActiveServer)
    With StatusWnd(tServerNum)
        .HasQuit = True
        .timerIdle.Enabled = False
        .tmrLag.Enabled = False
        .tmrChkLag.Enabled = False
        CloseLog logStatus, .Tag
    End With
    PutServ "QUIT :Advanced IRC " & VerStr & ": don't ask, don't tell.", tServerNum: DoEvents
    For C = 1 To ChannelWndU
        If C > ChannelWndU Then Exit For
        If ChannelWnd(C).ServerNum = tServerNum Then 'unload
            ChannelWnd(C).HasParted = True 'unngå part på en annen server
            Unload ChannelWnd(C)
            Dec C
        End If
    Next
    For C = 1 To PrivateWndU
        If C > PrivateWndU Then Exit For
        If PrivateWnd(C).ServerNum = tServerNum Then 'unload
            Unload PrivateWnd(C)
            Dec C
        End If
    Next
    For C = 1 To ChannelWndU
        If ChannelWnd(C).ServerNum > tServerNum Then ChannelWnd(C).ServerNum = ChannelWnd(C).ServerNum - 1
    Next
    For C = 1 To PrivateWndU
        If PrivateWnd(C).ServerNum > tServerNum Then PrivateWnd(C).ServerNum = PrivateWnd(C).ServerNum - 1
    Next
    For C = 1 To ChatWndU
        If ChatWnd(C).ServerNum > tServerNum Then ChatWnd(C).ServerNum = ChatWnd(C).ServerNum - 1
    Next
    For C = 1 To DCCWndU
        If DCCWnd(C).ServerNum > tServerNum Then DCCWnd(C).ServerNum = DCCWnd(C).ServerNum - 1
    Next
    ResetIdle tServerNum
    With StatusWnd(tServerNum)
        .timerIdle.Enabled = False
        .tmrChkLag.Enabled = False
        .tmrLag.Enabled = False
        .StatusLocked = False
        Set .IRC = Nothing
        Set .Ident = Nothing
    End With
    frmMain.WSwitch.RemWnd StatusWnd(tServerNum)
    Dec StatusWndU
    For C = tServerNum To StatusWndU
         Set StatusWnd(C) = StatusWnd(C + 1)
         StatusWnd(C).ServerNum = C
    Next
    If StatusWndU < 1 Then
        Erase StatusWnd
    Else
        ReDim Preserve StatusWnd(1 To StatusWndU)
    End If
    If ((StatusWndU = 0) And (UnloadMode <> 4)) Then NewStatusWnd
    If oas >= StatusWndU Then ActiveServer = StatusWndU Else ActiveServer = oas
    If ActiveServer = 0 Then Exit Sub 'Program shutdown
    Set fActive = StatusWnd(ActiveServer)
    If Not frmMain.IsBlocking Then frmMain.WSwitch.ActWnd fActive
End Sub

Function StWnd(ByVal Srv As String) As Integer
    Dim C As Long
    For C = 1 To StatusWndU
        If LCase(StatusWnd(C).Tag) = LCase(Srv) Then StWnd = C: Exit Function
    Next
End Function

Sub NewPrivateWnd(ByVal Nick As String, ByVal Hostmask As String, Optional ByVal DoActivate As Boolean)
    Dim ofActive As Form
    If Not PrWnd(Nick) = 0 Then Exit Sub
    Inc PrivateWndU
    ReDim Preserve PrivateWnd(1 To PrivateWndU)
    Set PrivateWnd(PrivateWndU) = New frmPrivate
    With PrivateWnd(PrivateWndU)
        If Hostmask = "" Then .Caption = Nick Else .Caption = Nick & " (" & Hostmask & ")"
        .Tag = Nick
        .WindowNum = PrivateWndU
        .ServerNum = ActiveServer
        frmMain.WSwitch.AddWnd PrivateWnd(PrivateWndU), ActiveServer, wndPrivate
        If Not DoActivate Then Set ofActive = fActive
        If frmMain.IsBlocking Then
            Inc SavedWndsU
            ReDim Preserve SavedWnds(1 To SavedWndsU)
            Set SavedWnds(SavedWndsU) = PrivateWnd(PrivateWndU)
        Else
            If DoActivate Then
                .Show
                .txtInput.SetFocus
                Set fActive = PrivateWnd(PrivateWndU)
            End If
        End If
    End With
    SetColorWindows PrivateWnd(PrivateWndU)
    frmMain.WSwitch.ActWnd fActive
    If ((Not DoActivate) And (Not frmMain.IsBlocking)) Then
        Set fActive = ofActive
        frmMain.WSwitch.ColWnd PrivateWnd(PrivateWndU), vbRed
    End If
    OpenLog logPrivate, Nick
End Sub

Sub UnloadPrivateWnd(ByVal WindowNum As Integer)
    Dim C As Long
    CloseLog logPrivate, PrivateWnd(WindowNum).Tag
    If WindowNum = 0 Then Exit Sub
    frmMain.WSwitch.RemWnd PrivateWnd(WindowNum)
    frmMain.WSwitch.Refresh
    Dec PrivateWndU
    For C = WindowNum To PrivateWndU
        Set PrivateWnd(C) = PrivateWnd(C + 1)
        PrivateWnd(C).WindowNum = C
    Next
    If PrivateWndU > 0 Then
        ReDim Preserve PrivateWnd(1 To PrivateWndU)
    Else
        Erase PrivateWnd
    End If
    Set fActive = StatusWnd(ActiveServer)
End Sub

Function PrWnd(ByVal Nick As String) As Integer
    Dim C As Long
    If PrivateWndU = 0 Then PrWnd = 0: Exit Function
    For C = 1 To PrivateWndU
        If LCase(PrivateWnd(C).Tag) = LCase(Nick) Then If PrivateWnd(C).ServerNum = ActiveServer Then PrWnd = C: Exit Function
    Next
End Function

Sub NewChannelWnd(ByVal Chan As String)
    If Not ChWnd(Chan) = 0 Then Exit Sub
    Inc ChannelWndU
    ReDim Preserve ChannelWnd(1 To ChannelWndU)
    ReDim Preserve Ignore(1 To ChannelWndU)
    ReDim Preserve Nicklist(1 To ChannelWndU)
    ReDim Preserve ChanProps(1 To ChannelWndU)
    Set ChannelWnd(ChannelWndU) = New frmChannel
    With ChannelWnd(ChannelWndU)
        .Tag = Chan
        .WindowNum = ChannelWndU
        .Caption = Chan & "  - [" & ChanProps(ChannelWndU).Modes & "] - [" & ChanProps(ChannelWndU).topic & "]"
        .Visible = True
        frmMain.WSwitch.AddWnd ChannelWnd(ChannelWndU), ActiveServer, wndChannel
        SetColorWindows ChannelWnd(ChannelWndU), True
        If frmMain.IsBlocking Then
            Inc SavedWndsU
            ReDim Preserve SavedWnds(1 To SavedWndsU)
            Set SavedWnds(SavedWndsU) = ChannelWnd(ChannelWndU)
        Else
            .Show
        End If
        .ServerNum = ActiveServer
        If Not frmMain.IsBlocking Then
            frmMain.WSwitch.ActWnd ChannelWnd(ChannelWndU)
            .txtInput.SetFocus
        End If
        UpdateChannelWindows ChannelWndU
        OpenLog logChannel, Chan
    End With
End Sub

Sub UnloadChannelWnd(ByVal WindowNum As Integer, Optional ByVal HasParted As Boolean)
    Dim C As Long
    Dim Chan As String
    If WindowNum = 0 Then Exit Sub
    Chan = ChannelWnd(WindowNum).Tag
    CloseLog logChannel, Chan
    frmMain.WSwitch.RemWnd ChannelWnd(WindowNum)
    frmMain.WSwitch.Refresh
    Dec ChannelWndU
    For C = WindowNum To ChannelWndU
        Ignore(C) = Ignore(C + 1)
        ChanProps(C) = ChanProps(C + 1)
        Set Nicklist(C) = Nicklist(C + 1)
        Nicklist(C).SetChan C
        Set ChannelWnd(C) = ChannelWnd(C + 1)
        ChannelWnd(C).WindowNum = C
    Next
    If ChannelWndU > 0 Then
        ReDim Preserve ChannelWnd(1 To ChannelWndU)
        ReDim Preserve Ignore(1 To ChannelWndU)
        ReDim Preserve Nicklist(1 To ChannelWndU)
        ReDim Preserve ChanProps(1 To ChannelWndU)
    Else
        Erase ChannelWnd
        Erase Ignore
        Erase Nicklist
        Erase ChanProps
    End If
    Set fActive = StatusWnd(ActiveServer)
    If Not HasParted Then PutServ "PART " & Chan
End Sub

Function ChWnd(ByVal Chan As String) As Integer
    Dim C As Long
    If ChannelWndU = 0 Then ChWnd = 0: Exit Function
    For C = 1 To ChannelWndU
        If LCase(ChannelWnd(C).Tag) = LCase(Chan) Then If ChannelWnd(C).ServerNum = ActiveServer Then ChWnd = C: Exit Function
        If ChannelWnd(C).Tag = "" Then
            frmMain.WSwitch.RemWnd ChannelWnd(C)
            Unload ChannelWnd(C)
            Exit Function
        End If
    Next
End Function

Sub NewChatWnd(ByVal Nick As String, ByVal IP As String, ByVal Port As String, ByVal DoConnect As Boolean, Optional ByVal UID As Long)
    Inc ChatWndU
    ReDim Preserve ChatWnd(1 To ChatWndU)
    Set ChatWnd(ChatWndU) = New frmDCCChat
    With ChatWnd(ChatWndU)
        .Tag = Nick
        .WindowNum = ChatWndU
        .Caption = Nick & " (" & IP & ":" & Port & ")"
        .Visible = True
    End With
    frmMain.WSwitch.AddWnd ChatWnd(ChatWndU), ActiveServer, wndChat
    SetColorWindows ChatWnd(ChatWndU)
    If frmMain.IsBlocking Then
        Inc SavedWndsU
        ReDim Preserve SavedWnds(1 To SavedWndsU)
        Set SavedWnds(SavedWndsU) = StatusWnd(StatusWndU)
    Else
        'ChatWnd(ChatWndU).SetFocus
        'ChatWnd(ChatWndU).txtInput.SetFocus
        frmMain.WSwitch.ActWnd ChatWnd(ChatWndU)
    End If
    With ChatWnd(ChatWndU)
        .ServerNum = ActiveServer
        OpenLog logDCC, Nick
        If DoConnect Then 'Is client
            .UniqueID = UID
            If Not UID = 0 Then
                Output "*** " & Nick & " wants to start a passive DCC chat.", ChatWnd(ChatWndU), statusc
                .DCCProtocol = dccPassive
            Else
                Output "*** " & Nick & " wants to start a DCC chat (" & IP & ":" & Port & ")", ChatWnd(ChatWndU), statusc
                .Chat.RemoteHost = IP
                .Chat.RemotePort = Port
            End If
            Output "*** Press Enter to accept, Escape to decline", ChatWnd(ChatWndU), statusc
        Else 'Is server
            If DCCInfo.PassiveDCC Then 'Initalize passive DCC
                .DCCProtocol = dccPassive
                Inc DCCUnique
                .UniqueID = DCCUnique
                Output "*** Initializing passive DCC, waiting for acknowledgment...", ChatWnd(ChatWndU), statusc
                SendCTCP Nick, "DCC CHAT " & StatusWnd(ActiveServer).CurrentNick & " " & PutIP(DCCIP) & " 0 " & .UniqueID, True
            Else
                .Chat.Bind NextDCCPort, DCCIP
                .Chat.Listen
                Output "*** Waiting for reply...", ChatWnd(ChatWndU), statusc
                SendCTCP Nick, "DCC CHAT " & StatusWnd(ActiveServer).CurrentNick & " " & PutIP(DCCIP) & " " & .Chat.LocalPort, True
            End If
        End If
        ChatWnd(ChatWndU).ServerNum = ActiveServer
    End With
End Sub

Sub UnloadChatWnd(ByVal WindowNum As Integer)
    Dim C As Long
    If WindowNum = 0 Then Exit Sub
    CloseLog logDCC, ChatWnd(WindowNum).Tag
    frmMain.WSwitch.RemWnd ChatWnd(WindowNum)
    frmMain.WSwitch.Refresh
    Dec ChatWndU
    For C = WindowNum To ChatWndU
        Set ChatWnd(C) = ChatWnd(C + 1)
        ChatWnd(C).WindowNum = C
    Next
    If ChatWndU > 0 Then
        ReDim Preserve ChatWnd(1 To ChatWndU)
    Else
        Erase ChatWnd
    End If
    If Not ActiveServer = 0 Then Set fActive = StatusWnd(ActiveServer)
End Sub

Function DCWnd(ByVal Nick As String) As Integer
    Dim C As Long
    If ChatWndU = 0 Then DCWnd = 0: Exit Function
    For C = 1 To ChatWndU
        If LCase(ChatWnd(C).Tag) = LCase(Nick) Then DCWnd = C: Exit Function
    Next
End Function

Sub NewDCCWnd(ByVal Nick As String, ByVal Filename As String, ByVal Size As Currency, ByVal IP As String, _
              ByVal Port As String, ByVal IsSender As Boolean, Optional ByVal SendMsgByDCC As Boolean = False, _
              Optional ByVal UID As Long = 0)
    Dim V As Variant
    Dim Fn As Integer
    Inc DCCWndU
    ReDim Preserve DCCWnd(1 To DCCWndU)
    Set DCCWnd(DCCWndU) = New frmDCCSend
    With DCCWnd(DCCWndU)
        .Caption = "DCC " & Nick
        .WindowNum = DCCWndU
        .ServerNum = ActiveServer
        .Tag = Nick
        If Not IsSender Then
            .DCCLog = "Event log start: (" & CStr(Now) & ")" & vbCrLf & _
                      "Nick=" & Nick & " ; Filename=""" & Filename & """ ; Filesize=" & Size & _
                   " ; IP=" & IP & " ; Port=" & Port & vbCrLf
        End If
    End With
    frmMain.WSwitch.AddWnd DCCWnd(DCCWndU), ActiveServer, wndDCC
    If frmMain.IsBlocking Then
        Inc SavedWndsU
        ReDim Preserve SavedWnds(1 To SavedWndsU)
        Set SavedWnds(SavedWndsU) = DCCWnd(DCCWndU)
    Else
        DCCWnd(DCCWndU).Show
    End If
    With DCCWnd(DCCWndU)
        If IsSender Then
            .cmdSend.Caption = "Send"
            .cmdCancel.Caption = "Close"
            .txtStatus = dccStatusReadySend
            If DCCInfo.PassiveDCC Then
                .DCCProtocol = dccPassive
                Inc DCCUnique 'Never decreases, makes it unique
                .UniqueID = DCCUnique
            End If
        Else
            .txtStatus = dccStatusReadyReceive
            .cmdSend.Caption = "Receive"
            .cmdCancel.Caption = "Reject"
            .DCC.RemoteHost = IP
            .DCC.RemotePort = Port
            .UniqueID = UID
            If Not UID = 0 Then
                .DCCProtocol = dccPassive
                On Error Resume Next
                .DCC.Bind NextDCCPort, DCCIP
                Do While (Err.Number <> 0)
                    Err.Clear
                    .DCC.Bind NextDCCPort, DCCIP
                Loop
                On Error GoTo 0
                .DCCLog = .DCCLog & "Protocol used is DCC/PASV." & vbCrLf
            End If
        End If
    End With
    DCCWnd(DCCWndU).FName = Filename
    Filename = TrimPath(Filename, True)
    DCCWnd(DCCWndU).txtFilename = Filename
    If Not IsSender Then
        Fn = FreeFile 'GRR POKKER
        Open DCCInfo.DownloadDir & Filename For Random As Fn
        If LOF(Fn) > 0 Then 'aktiver resume
            DCCWnd(DCCWndU).Progress.Width = 2175
            DCCWnd(DCCWndU).cmdResume.Visible = True
        End If
        Close Fn
    End If
    With DCCWnd(DCCWndU)
        .txtFileSize = ShortenBytes(Size)
        .FSize = Size
        .Progress.SetMax Size
        .Nick = Nick
        .txtNickname = Nick
        .txtTimeElapsed = ShortenTime(0)
        .txtSendspeed = ShortenBytes(0) & "/s"
        .txtTimeLeft = ShortenTime(0)
        .IsSender = IsSender
        .IsReceiver = Not IsSender
        .SendMsgByDCC = SendMsgByDCC
        If .SendMsgByDCC Then Set .MDSock = ClickSock
        If .IsReceiver Then
            .DoReceive
        End If
        .maReady = True
    End With
    AddDCC DCCWndU
End Sub

Sub UnloadDCCWnd(ByVal WindowNum As Integer)
    Dim C As Long
    DCCWnd(WindowNum).maReady = False
    frmMain.WSwitch.RemWnd DCCWnd(WindowNum)
    frmMain.WSwitch.Refresh
    KillDCC WindowNum
    Dec DCCWndU
    For C = WindowNum To DCCWndU
        Set DCCWnd(C) = DCCWnd(C + 1)
        DCCWnd(C).WindowNum = C
    Next
    If ActiveServer = 0 Then Exit Sub
    Set fActive = StatusWnd(ActiveServer)
    frmMain.WSwitch.ActWnd frmMain.ActiveForm
End Sub

Function FindDCCWindow(Optional ByVal Port As Long, Optional ByVal UID As Long, Optional ByVal ByRec As Boolean = False) As frmDCCSend
    Set FindDCCWindow = Nothing
    Dim L_Port As Long
    If ((Port = 0) And (UID = 0)) Then Exit Function
    Dim C As Long
    For C = 1 To DCCWndU
        With DCCWnd(C)
            UID = .UniqueID
            If Not ByRec Then
                L_Port = .DCC.LocalPort
            Else
                L_Port = .DCC.RemotePort
            End If
            If .IsReceiver = ByRec Then
                If Port = 0 Then
                    If .UniqueID = UID Then Exit For
                ElseIf UID = 0 Then
                    If L_Port = Port Then Exit For
                Else
                    If ((.UniqueID = UID) And (L_Port = Port)) Then Exit For
                End If
            End If
        End With
    Next
    If C = DCCWndU + 1 Then Exit Function
    Set FindDCCWindow = DCCWnd(C)
End Function

Function FindChatWindow(ByVal UID As Long) As frmDCCChat
    Set FindChatWindow = Nothing
    Dim C As Long
    For C = 1 To ChatWndU
        If ChatWnd(C).UniqueID = UID Then Exit For
    Next
    If C > ChatWndU Then Exit Function
    Set FindChatWindow = ChatWnd(C)
End Function

Function MainFindWnd(ParamArray WName() As Variant) As Form
    Dim C As Long
    For C = LBound(WName) To UBound(WName)
        If ChWnd(WName(C)) > 0 Then Set MainFindWnd = ChannelWnd(ChWnd(WName(C))): Exit Function
        If PrWnd(WName(C)) > 0 Then Set MainFindWnd = PrivateWnd(PrWnd(WName(C))): Exit Function
        If StWnd(WName(C)) > 0 Then Set MainFindWnd = StatusWnd(StWnd(WName(C))): Exit Function
    Next
    Set MainFindWnd = StatusWnd(ActiveServer) 'If none found, set to active window
End Function

Function ChkWndChange(ByVal KeyCode As Integer) As Boolean
    'returns true if window has changed
    'returns false if keycode is not numeric
    'If Shift = 4 Then
        Select Case KeyCode
            Case 49, 50, 51, 52, 53, 54, 55, 56, 57
            '1, 2, 3, 4, 5, 6, 7, 8, 9
                ChkWndChange = True
                frmMain.WSwitch.NumWnd KeyCode - 48
            Case 48 '0
                ChkWndChange = True
                frmMain.WSwitch.NumWnd 10
            Case Else 'Check letters
                ChkWndChange = False
                Select Case Chr(KeyCode)
                    Case "Q", "q"
                        ChkWndChange = True
                        frmMain.WSwitch.NumWnd 11
                    Case "W", "w"
                        ChkWndChange = True
                        frmMain.WSwitch.NumWnd 12
                    Case "E", "e"
                        ChkWndChange = True
                        frmMain.WSwitch.NumWnd 13
                    Case "R", "r"
                        ChkWndChange = True
                        frmMain.WSwitch.NumWnd 14
                    Case "T", "t"
                        ChkWndChange = True
                        frmMain.WSwitch.NumWnd 15
                    Case "Y", "y"
                        ChkWndChange = True
                        frmMain.WSwitch.NumWnd 16
                    Case "U", "u"
                        ChkWndChange = True
                        frmMain.WSwitch.NumWnd 17
                    Case "I", "i"
                        ChkWndChange = True
                        frmMain.WSwitch.NumWnd 18
                    Case "O", "o"
                        ChkWndChange = True
                        frmMain.WSwitch.NumWnd 19
                    Case "P", "p"
                        ChkWndChange = True
                        frmMain.WSwitch.NumWnd 20
                    Case Else
                End Select
        End Select
    'End If
End Function

Sub ShowOptionWnd()
    frmConnect.Show
    frmConnect.comboServer.SetFocus
    frmConnect.Hide
    frmMain.ToggleBlock True
    frmConnect.Show vbModal, frmMain
    frmMain.ToggleBlock False
End Sub

'######### SCRIPT/PLUGIN SECTION #########

'Script
Function FindScript(ByVal ScriptName As String) As Integer
    Dim C As Long
    Dim s As String
    For C = 1 To ScriptArrayU
        s = TrimPath(LCase(ScriptArray(C).File_Name))
        If ((LCase(ScriptArray(C).Sc_Name) = LCase(ScriptName)) Or (Mid(s, 1, Len(s) - 4) = LCase(ScriptName))) Then
            FindScript = C
            Exit Function
        End If
    Next
End Function

'Plugin
Sub AddOCX(ByVal OCXPath As String)
    Dim C As Long
    Dim s As String
    Dim V As Variant
    On Error Resume Next
    If InStr(1, OCXPath, "\") = 0 Then OCXPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & OCXPath
    If Not LCase(Right(OCXPath, 4)) = ".dll" Then OCXPath = OCXPath & ".dll"
    If FileLen(OCXPath) = 0 Then 'Either error or 0-byte
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    Inc airc_AddInCount
    ReDim Preserve airc_AddIns(1 To airc_AddInCount)
    With airc_AddIns(airc_AddInCount)
        .Filename = OCXPath
        .AddinName = TrimPath(OCXPath): .AddinName = Left(.AddinName, Len(.AddinName) - 4)
        Set .AddinObj = CreateObject(.AddinName & ".clsMain")
        .AddinObj.Addin_Start s
        If Not s = "" Then
            V = TrimCrLf_Out(s)
            For C = LBound(V) To UBound(V)
                Output "PLUGIN> (" & .AddinName & ") " & V(C), fActive, statusc, True
            Next
        End If
    End With
End Sub

Sub RemoveOCX(ByVal OCXNumber As Long)
    Dim s As String
    Dim V As Variant
    If OCXNumber = 0 Then Exit Sub
    If OCXNumber > airc_AddInCount Then Exit Sub
    Dim C As Long
    With airc_AddIns(airc_AddInCount)
        .AddinObj.Addin_Close s
        If Not s = "" Then
            V = TrimCrLf_Out(s)
            For C = LBound(V) To UBound(V)
                Output "PLUGIN> (" & .AddinName & ") " & V(C), fActive, statusc, True
            Next
        End If
    End With
    Dec airc_AddInCount
    For C = OCXNumber To airc_AddInCount
        airc_AddIns(C) = airc_AddIns(C + 1)
    Next
    If airc_AddInCount = 0 Then
        Erase airc_AddIns
    Else
        ReDim Preserve airc_AddIns(1 To airc_AddInCount)
    End If
End Sub

Function FindOCX(ByVal OCXName As String) As Integer
    Dim C As Long
    Dim s As String
    For C = 1 To airc_AddInCount
        If LCase(airc_AddIns(C).AddinName) = LCase(OCXName) Then
            FindOCX = C
            Exit Function
        End If
    Next
End Function



'######### IGNORE SECTION #########

Sub NewIgnore(ByVal Nick As String)
    Inc IgnoreU
    ReDim Preserve IgnoreP(1 To IgnoreU)
    IgnoreP(IgnoreU).Nick = Nick
End Sub

Function IgnCC(ByVal Nick As String) As Integer
    Dim C As Long
    If IgnoreU = 0 Then IgnCC = 0: Exit Function
    For C = 1 To IgnoreU
        If LCase(IgnoreP(C).Nick) = LCase(Nick) Then IgnCC = C: Exit Function
    Next
End Function

