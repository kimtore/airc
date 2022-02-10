Attribute VB_Name = "modCTCP"
Option Explicit

'CTCP Character = ""

Sub StartLagCount(ByVal ServerNum As Long)
    With StatusWnd(ServerNum)
        .LagTime = 0
        .tmrLag.Enabled = True
        If ServerNum = ActiveServer Then frmMain.IRCStatus.Changelag ShortenTime(0)
        PutServ "PING " & timer, ServerNum
    End With
End Sub

Sub EndLagCount(ByVal ServerNum As Long)
    With StatusWnd(ServerNum)
        .tmrLag.Enabled = False
        .tmrChkLag.Enabled = False: .tmrChkLag.Enabled = True
        If ServerNum = ActiveServer Then frmMain.IRCStatus.Changelag ShortenTime(.LagTime)
    End With
End Sub

Function TCloak(ByVal T As String) As TypeCloak
    Select Case UCase(T)
        Case "PING"
            TCloak = Cloak.Ping
        Case "TIME"
            TCloak = Cloak.Time
        Case "VERSION"
            TCloak = Cloak.Version
        Case "URL"
            TCloak = Cloak.URL
    End Select
End Function

Sub STCloak(ByVal T As String, ByRef TC As TypeCloak)
    Select Case UCase(T)
        Case "PING"
            Cloak.Ping = TC
        Case "TIME"
            Cloak.Time = TC
        Case "VERSION"
            Cloak.Version = TC
        Case "URL"
            Cloak.URL = TC
    End Select
End Sub

Function TrimCTCP(ByVal s As String) As String
    TrimCTCP = Replace(s, "", "")
End Function

Sub SendCTCP(ByVal sendto As String, ByVal s As String, Optional ByVal HideRequest As Boolean = False)
    'Simple split operation to make first word uppercase
    Dim V As Variant
    If s = "" Then Exit Sub
    V = Split(s, " ")
    V(0) = UCase(V(0))
    s = Merge(V, 0)
    If Not HideRequest Then CTCPOut sendto, s, True, False
    If V(0) = "PING" Then 'Ping
        s = s & " " & timer
    End If
    PutCTCPChar s
    PutServ "PRIVMSG " & sendto & " :" & s
    ResetIdle
End Sub

Sub SendCTCPReply(ByVal sendto As String, ByVal Text As String)
    PutServ "NOTICE " & sendto & " :" & Text & ""
    ResetIdle
End Sub

Sub CTCPOut(ByVal Nick As String, ByVal Text As String, ByVal Outgoing As Boolean, ByVal IsReply As Boolean)
    Dim Arr As String
    Arr = IIf(Outgoing, "->", "<-")
    Output "[" & SecColNum & IIf(IsReply, "ctcpreply", "ctcp") & "]" & Arr & "[" & StdColNum & "" & Nick & "]" & Arr & " " & Text, fActive
End Sub

Function PutCTCPChar(ByRef Text As String) As String
    Text = TrimCTCP(Text)
    Text = "" & Text & ""
    PutCTCPChar = Text
End Function

Function IsCTCP(ByVal Text As String) As Boolean
    If ((Left(Text, 1) = "") And (Right(Text, 1) = "")) Then IsCTCP = True
End Function

'### DCC Handling ###'

Sub AddDCC(ByVal C As Long)
    If Not dcStat Then Exit Sub
    With frmDCCStatus.listTransfers.ListItems
        .Add , , DCCWnd(C).FName
        .Item(C).Tag = DCCWnd(C).WindowNum
        With DCCWnd(C)
            ModifyDCC C, .FName, .Nick, .FSize, .FSent, .FReceived, .txtSendspeed, .Percentage, .TimeElapsed, _
            .TimeLeft, IIf(DCCWnd(C).IsSender, "outgoing", "incoming"), Mid(DCCWnd(C).txtStatus, 9)
        End With
    End With
End Sub

Sub ModifyDCC(ByVal C As Long, ByVal dName As String, ByVal dNick As String, ByVal dSize As Long, ByVal dSent As Long, dReceived As Long, _
ByVal dSpeed As String, ByVal dPercent As String, ByVal dElapsed, ByVal dLeft As Long, ByVal dDirection As String, _
ByVal dStatus As String)
    If Not dcStat Then Exit Sub
    With frmDCCStatus.listTransfers.ListItems
        .Item(C) = TrimPath(dName)
        With .Item(C)
            .SubItems(1) = dNick
            .SubItems(2) = ShortenBytes(dSize)
            .SubItems(3) = ShortenBytes(dSent)
            .SubItems(4) = ShortenBytes(dReceived)
            .SubItems(5) = dSpeed
            .SubItems(6) = dPercent
            .SubItems(7) = ShortenTime(dElapsed)
            .SubItems(8) = ShortenTime(dLeft)
            .SubItems(9) = dDirection
            .SubItems(10) = dStatus
        End With
    End With
End Sub

Sub KillDCC(ByVal C As Long)
    If Not dcStat Then Exit Sub
    With frmDCCStatus.listTransfers.ListItems
        .Remove C
        For C = C To .Count
            .Item(C).Tag = C
        Next
    End With
End Sub
