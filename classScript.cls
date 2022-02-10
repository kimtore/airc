VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Functions which the script can use

Sub aecho(ByVal s As Variant, Optional ByVal StdColor As Variant, Optional ByVal DrawBrand As Variant)
    If IsMissing(StdColor) Then StdColor = -1
    If IsEmpty(StdColor) Then StdColor = -1
    If IsMissing(DrawBrand) Then DrawBrand = False
    If IsEmpty(DrawBrand) Then DrawBrand = False
    Output CStr(s), fActive, , CBool(DrawBrand)
End Sub

Sub secho(ByVal s As Variant, Optional ByVal StdColor As Variant, Optional ByVal DrawBrand As Variant)
    If IsMissing(StdColor) Then StdColor = -1
    If IsEmpty(StdColor) Then StdColor = -1
    If IsMissing(DrawBrand) Then DrawBrand = False
    If IsEmpty(DrawBrand) Then DrawBrand = False
    Output CStr(s), StatusWnd(modWindows.ActiveServer), , CBool(DrawBrand)
End Sub

Sub echo(ByVal s As Variant, ByVal Where As Variant, Optional ByVal Server As Variant, Optional ByVal StdColor As Variant, Optional ByVal DrawBrand As Variant)
    Dim oas As Integer
    Dim s_wnd As Form
    If ((Server < 1) Or (Server > StatusWndU)) Then Exit Sub
    oas = ActiveServer
    modWindows.ActiveServer = Server
    If IsChan(Where) Then
        If Not ChWnd(Where) = 0 Then
            Set s_wnd = ChannelWnd(ChWnd(Where))
        End If
    Else
        If Not PrWnd(Where) = 0 Then
            Set s_wnd = PrivateWnd(PrWnd(Where))
        End If
    End If
    modWindows.ActiveServer = oas
    If IsMissing(StdColor) Then StdColor = -1
    If IsEmpty(StdColor) Then StdColor = -1
    If IsMissing(DrawBrand) Then DrawBrand = False
    If IsEmpty(DrawBrand) Then DrawBrand = False
    Output CStr(s), s_wnd, , CBool(DrawBrand)
End Sub

Sub putignore(ByVal Where As Variant, ByVal Types As Variant)
    If IsChan(Where) Then
        modParsing.ParseIgnore Where, Types, True, DoSilent:=True
    Else
        modParsing.ParseIgnore Where, Types, True, True, , True
    End If
End Sub

Sub remignore(ByVal Where As Variant, ByVal Types As Variant)
    If IsChan(Where) Then
        modParsing.ParseIgnore Where, Types, False, DoSilent:=True
    Else
        modParsing.ParseIgnore Where, Types, False, True, , True
    End If
End Sub

Sub putmsg(ByVal Text As Variant, ByVal Where As Variant)
    If IsMissing(Text) Then Exit Sub
    If IsMissing(Where) Then Exit Sub
    PutServ "privmsg " & Where & " :" & Text
    ResetIdle
End Sub

Sub putnotice(ByVal Text As Variant, ByVal Where As Variant)
    If IsMissing(Text) Then Exit Sub
    If IsMissing(Where) Then Exit Sub
    PutServ "notice " & Where & " :" & Text
    ResetIdle
End Sub

Sub putserver(ByVal Text As Variant, Optional ByVal ServerNum As Variant)
    If IsMissing(ServerNum) Then ServerNum = ActiveServer
    PutServ Text, ServerNum
    If ((LCase(Left(Text, 8)) = "privmsg ") Or (LCase(Left(Text, 7)) = "notice ")) Then ResetIdle
End Sub

Sub runcmd(ByVal Text As Variant, Optional ByVal Where As Variant = "")
    If ((IsMissing(Where)) Or (IsEmpty(Where))) Then
        Parse Text
    Else
        Parse Text, CStr(Where)
    End If
End Sub

Sub say(ByVal Text As Variant)
    If fActive.name = frmStatus.name Then Exit Sub
    Parse Text, fActive.Tag
End Sub

Sub closequery(ByVal Nick As Variant)
    If Not PrWnd(Nick) = 0 Then Unload PrivateWnd(PrWnd(Nick))
End Sub

Sub putctcp(ByVal Text As Variant, ByVal Where As Variant)
    If IsMissing(Text) Then Exit Sub
    If IsMissing(Where) Then Exit Sub
    SendCTCP Text, Where
    ResetIdle
End Sub

Sub putctcpreply(ByVal Text As Variant, ByVal Where As Variant)
    If IsMissing(Text) Then Exit Sub
    If IsMissing(Where) Then Exit Sub
    PutServ "notice " & Where & " :" & Text & ""
    ResetIdle
End Sub

Sub SetAway(Optional ByVal Text As Variant, Optional ByVal ServerNum As Variant)
    If IsMissing(ServerNum) Then ServerNum = ActiveServer
    If IsMissing(Text) Then
        Text = ""
        If (ServerNum = ActiveServer) Then frmMain.IRCStatus.ChangeAway ""
    Else
        If (ServerNum = ActiveServer) Then frmMain.IRCStatus.ChangeAway CStr(Text)
        Text = " :" & Text
    End If
    StatusWnd(ActiveServer).AwayReason = Text
    PutServ "AWAY" & Text
End Sub




Function strsplit(s As Variant, splitchar As Variant) As Variant
    Dim M() As Variant
    If IsMissing(s) Then Exit Function
    If IsMissing(splitchar) Then splitchar = " "
    M = Split(s, splitchar)
    strsplit = M
End Function

Function searchstr(s1 As Variant, s2 As Variant) As Variant
    If InStr(1, LCase(s1), LCase(s2)) <> 0 Then searchstr = True Else searchstr = False
End Function

Function ActiveServer() As Variant
    ActiveServer = modWindows.ActiveServer
End Function

Function ServerName() As Variant
    ServerName = StatusWnd(ActiveServer).ServerName
End Function

Function fActive() As Form
    Set fActive = modVariables.fActive
End Function

Function CurNick() As Variant
    CurNick = StatusWnd(ActiveServer).CurrentNick
End Function

Function FetchHost(Nick As Variant, HostMode As Variant) As Variant
    Dim M As Integer
    Dim C As Long
    FetchHost = ""
    If ((IsMissing(Nick)) Or (IsMissing(HostMode))) Then Exit Function
    If HostMode = "" Then Exit Function
    If Not IsNumeric(HostMode) Then Exit Function
    If ((HostMode < 0) Or (HostMode > 9)) Then Exit Function
    If ChannelWndU = 0 Then Exit Function
    For C = 1 To ChannelWndU
        M = Nicklist(C).UserPos(Nick)
        If Not M = 0 Then Exit For
    Next
    If M = 0 Then Exit Function
    FetchHost = UserHostMode(Nick, Nicklist(C).User_Host(M), HostMode)
End Function


Function IsAway() As Variant
    If StatusWnd(ActiveServer).AwayReason = "" Then IsAway = False Else IsAway = True
End Function
