Attribute VB_Name = "modLogging"
Option Explicit

Public Enum LogTypes
    logStatus = 1
    logChannel = 2
    logPrivate = 3
    logDCC = 4
    logAll = 5
End Enum

Sub LogStr(F As Form, ByVal Text As String)
    If Not LogInfo.BrukLogg Then Exit Sub
    On Error GoTo ErrHandle
    Text = TrimCrLf(Text)
    If Not F.LogOpen Then Exit Sub
    Print #F.LogNum, Text
    On Error GoTo 0
    Exit Sub
ErrHandle:
    Output "Could not write to log.", fActive, statusc, True, True
    Resume Next
End Sub

Sub OpenLog(L As LogTypes, Optional N As String)
    Dim F As Form
    If Not LogInfo.BrukLogg Then Exit Sub
    On Error GoTo ErrHandle
    Select Case L
        Case logStatus
            If Not LogInfo.LoggStatus Then Exit Sub
            If N = "" Then Exit Sub
            Set F = StatusWnd(StWnd(N))
            If F.LogOpen Then Exit Sub
            F.LogNum = FreeFile
            Open LogInfo.LoggDir & N & ".log" For Append Shared As #F.LogNum
            F.LogOpen = True
            Print #F.LogNum, vbCrLf & "Log opened " & CStr(Now)
        Case logChannel
            If Not LogInfo.LoggKanaler Then Exit Sub
            If N = "" Then Exit Sub
            Set F = ChannelWnd(ChWnd(N))
            If F.LogOpen Then Exit Sub
            F.LogNum = FreeFile
            Open LogInfo.LoggDir & TrimBad(N) & ".log" For Append Shared As #F.LogNum
            F.LogOpen = True
            Print #F.LogNum, vbCrLf & "Log opened " & CStr(Now)
        Case logPrivate
            If Not LogInfo.LoggPrivat Then Exit Sub
            If N = "" Then Exit Sub
            Set F = PrivateWnd(PrWnd(N))
            If F.LogOpen Then Exit Sub
            F.LogNum = FreeFile
            Open LogInfo.LoggDir & TrimBad(N) & ".log" For Append Shared As #F.LogNum
            F.LogOpen = True
            Print #F.LogNum, vbCrLf & "Log opened " & CStr(Now)
        Case logDCC
            If Not LogInfo.LoggDCC Then Exit Sub
            If N = "" Then Exit Sub
            Set F = ChatWnd(DCWnd(N))
            If F.LogOpen Then Exit Sub
            F.LogNum = FreeFile
            Open LogInfo.LoggDir & TrimBad(N) & ".DCC.log" For Append Shared As #F.LogNum
            F.LogOpen = True
            Print #F.LogNum, vbCrLf & "Log opened " & CStr(Now)
        Case Else
            Exit Sub
    End Select
    Exit Sub
ErrHandle:
    Output "Could not open log for " & N & "!", fActive, statusc, True, True
    On Error GoTo 0
    Err.Clear
End Sub

Sub CloseLog(L As LogTypes, Optional N As String)
    On Error GoTo ErrHandle
    Dim F As Form
    Select Case L
        Case logStatus
            If N = "" Then Exit Sub
            Set F = StatusWnd(StWnd(N))
        Case logChannel
            If N = "" Then Exit Sub
            Set F = ChannelWnd(ChWnd(N))
        Case logPrivate
            If N = "" Then Exit Sub
            Set F = PrivateWnd(PrWnd(N))
        Case logDCC
            If N = "" Then Exit Sub
            Set F = ChatWnd(DCWnd(N))
        Case 5
            CloseAllLogs logAll
            Exit Sub
        Case Else
            Exit Sub
    End Select
    If Not F.LogOpen Then Exit Sub
    Print #F.LogNum, "Log closed " & CStr(Now) & vbCrLf
    Close #F.LogNum
    F.LogOpen = False
    Exit Sub
ErrHandle:
    'Output "Could not close log for " & N & "!", fActive, statusc, True, True
    'Unneccessary error message
    On Error GoTo 0
    Err.Clear
End Sub

Sub CloseAllLogs(L As LogTypes)
    Dim C As Long
    Select Case L
        Case logStatus
            For C = 1 To StatusWndU
                CloseLog logStatus, StatusWnd(C).Tag
            Next
        Case logChannel
            For C = 1 To ChannelWndU
                CloseLog logChannel, ChannelWnd(C).Tag
            Next
        Case logPrivate
            For C = 1 To PrivateWndU
                CloseLog logPrivate, PrivateWnd(C).Tag
            Next
        Case logDCC
            For C = 1 To ChatWndU
                CloseLog logDCC, ChatWnd(C).Tag
            Next
        Case logAll
            For C = 1 To StatusWndU
                CloseLog logStatus, StatusWnd(C).Tag
            Next
            For C = 1 To ChannelWndU
                CloseLog logChannel, ChannelWnd(C).Tag
            Next
            For C = 1 To PrivateWndU
                CloseLog logPrivate, PrivateWnd(C).Tag
            Next
            For C = 1 To ChatWndU
                CloseLog logDCC, ChatWnd(C).Tag
            Next
        Case Else
    End Select
End Sub


Sub OpenAllLogs(L As LogTypes)
    Dim C As Long
    Select Case L
        Case logStatus
            For C = 1 To StatusWndU
                OpenLog logStatus, StatusWnd(C).Tag
            Next
        Case logChannel
            For C = 1 To ChannelWndU
                OpenLog logChannel, ChannelWnd(C).Tag
            Next
        Case logPrivate
            For C = 1 To PrivateWndU
                OpenLog logPrivate, PrivateWnd(C).Tag
            Next
        Case logDCC
            For C = 1 To ChatWndU
                OpenLog logDCC, ChatWnd(C).Tag
            Next
        Case logAll
            For C = 1 To StatusWndU
                OpenLog logStatus, StatusWnd(C).Tag
            Next
            For C = 1 To ChannelWndU
                OpenLog logChannel, ChannelWnd(C).Tag
            Next
            For C = 1 To PrivateWndU
                OpenLog logPrivate, PrivateWnd(C).Tag
            Next
            For C = 1 To ChatWndU
                OpenLog logDCC, ChatWnd(C).Tag
            Next
        Case Else
    End Select
End Sub
