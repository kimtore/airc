Attribute VB_Name = "modSettings"
Option Explicit

Sub InitSysAccess()
    Dim L As Long
    frmMain.INIAccess.INIFile = App.Path & "\settings.ini"
    Open App.Path & "\settings.ini" For Binary As 1
    L = LOF(1)
    Close 1
    If L = 0 Then
        CreateSettings
    Else
        Exit Sub
    End If
End Sub

Sub CreateSettings()
    With IRCInfo
        .ModeInvisible = -1
        .AutoMode = 1
    End With
    With DCCInfo
        .ProtectVirus = True
        .DoIgnoreFiltyper = False
        .DownloadDir = App.Path
        If Not Right(.DownloadDir, 1) = "\" Then .DownloadDir = .DownloadDir & "\"
        .DownloadDir = .DownloadDir & "Download"
        .IgnoreFiltyper = ""
        .JoinIgnore = True
        .PumpDCC = False
        .SendeBuffer = "512"
    End With
    With IPInfo
        .UseCustomIP = True
        .LookupType = 0
        .IP = "127.0.0.1"
    End With
    With LogInfo
        .BrukLogg = False
        .LoggDCC = False
        .LoggDir = App.Path
        If Not Right(.LoggDir, 1) = "\" Then .LoggDir = .LoggDir & "\"
        .LoggDir = .LoggDir & "Logs"
        .LoggKanaler = False
        .LoggPrivat = False
        .LoggStatus = False
    End With
    With DisplayInfo
        .Timestamp = "[HH:nn]"
        .ColorActivity = True
        .ShowNicklist = True
    End With
    With ColorInfo
        'Initialize standard colors
        Set .Font = New StdFont
        .Font.name = "Courier New"
        .Font.Size = 9
        .cJoin = RGB(0, 127, 127)
        .cPart = RGB(0, 127, 127)
        .cQuit = RGB(255, 0, 0)
        .cNick = RGB(0, 127, 0)
        .cKick = RGB(255, 0, 0)
        .cMode = RGB(0, 0, 255)
        .cAction = RGB(127, 0, 127)
        .cStatus = RGB(0, 0, 127)
        .cTopic = RGB(0, 127, 0)
        .cNotice = RGB(127, 0, 0)
        .cNormal = RGB(0, 0, 0)
        .cOwn = RGB(0, 0, 0)
        .cBackColor = RGB(207, 207, 207)
        .cURLColor = RGB(0, 0, 255)
        .cBrandColor = RGB(0, 0, 255)
        .cStdColor = RGB(0, 0, 255)
        .cSecColor = RGB(127, 127, 127)
    End With
    With AwayInfo
        .AAUse = True
        .AAMinutes = 10
        .AACancelAway = True
        .AAMsg = "Advanced IRC: auto-away after 10 mins"
        .CancelAway = False
    End With
    With HighlightInfo
        .UseHighlight = True
        .HiNick = True
        .UseBold = True
    End With
    SaveIRCInfo True
    SaveCloakInfo
    SaveDCCInfo
    SaveIPInfo
    SaveLogInfo
    SaveDisplayInfo
    SaveColorInfo
    SaveAwayInfo
    SaveHighlightInfo
End Sub

Sub SysSaveServerList(ByVal SList As Variant)
    Dim C As Long
    With frmMain.INIAccess
        .INIEntry = "Serverliste"
        .INIDeleteEntry
        .INIEntry = "Serverliste"
        For C = 0 To UBound(SList)
            .INISaveSetting "n" & C, SList(C)
        Next
    End With
End Sub

Function SysGetServerList() As Variant
    On Error GoTo ErrHndl
    Dim C As Long
    Dim I As Integer
    Dim V As Variant
    Dim O As Integer
    O = FreeFile
    Open frmMain.INIAccess.INIFile For Random As O
    If LOF(O) = 0 Then Close O: ReDim V(-1 To -1): SysGetServerList = V: Exit Function
    Close O
    frmMain.INIAccess.INIEntry = "Serverliste"
    ReDim V(-1 To 0)
    V = frmMain.INIAccess.INIGetKeyList
    If Not UBound(V) = -1 Then
        I = UBound(frmMain.INIAccess.INIGetKeyList)
        ReDim V(0 To I - 1)
        For C = 1 To I
            V(C - 1) = frmMain.INIAccess.INIGetSetting("n" & C - 1)
        Next
    End If
    SysGetServerList = V
    Exit Function
ErrHndl:
    Err.Clear
    ReDim V(-1 To -1)
    SysGetServerList = V
End Function

Sub SysSaveIRCInfo(ByVal iList As Variant)
    With frmMain.INIAccess
        .INIEntry = "IRC Info"
        .INISaveSetting "Nick", iList(0)
        .INISaveSetting "Altnick", iList(1)
        .INISaveSetting "Ident", iList(2)
        .INISaveSetting "Realname", iList(3)
        .INISaveSetting "UseIdent", iList(4)
        .INISaveSetting "ActiveServer", iList(5)
        .INISaveSetting "ActivePort", iList(6)
        .INISaveSetting "ModeInvisible", iList(7)
        .INISaveSetting "ModeWallops", iList(8)
        .INISaveSetting "AutoMode", iList(9)
    End With
End Sub

Function SysGetIRCInfo() As Variant
    Dim V As Variant
    Dim O As Integer
    O = FreeFile
    Open frmMain.INIAccess.INIFile For Random As O
    If LOF(O) = 0 Then Close O: ReDim V(-1 To -1): SysGetIRCInfo = V: Exit Function
    Close O
    frmMain.INIAccess.INIEntry = "IRC Info"
    ReDim V(0 To 9)
    With frmMain.INIAccess
        V(0) = .INIGetSetting("Nick")
        V(1) = .INIGetSetting("Altnick")
        V(2) = .INIGetSetting("Ident")
        V(3) = .INIGetSetting("Realname")
        V(4) = .INIGetSetting("UseIdent")
        V(5) = .INIGetSetting("ActiveServer")
        V(6) = .INIGetSetting("ActivePort")
        V(7) = .INIGetSetting("ModeInvisible")
        V(8) = .INIGetSetting("ModeWallops")
        V(9) = .INIGetSetting("AutoMode")
    End With
    SysGetIRCInfo = V
End Function

Sub SysSaveCloakInfo(cList As Variant)
    With frmMain.INIAccess
        .INIEntry = "Cloak"
        .INISaveSetting "Ping_Hide", cList(0)
        .INISaveSetting "Ping_Reply", cList(1)
        .INISaveSetting "Ping_Custom", cList(2)
        .INISaveSetting "Time_Hide", cList(3)
        .INISaveSetting "Time_Reply", cList(4)
        .INISaveSetting "Time_Custom", cList(5)
        .INISaveSetting "Version_Hide", cList(6)
        .INISaveSetting "Version_Reply", cList(7)
        .INISaveSetting "Version_Custom", cList(8)
        .INISaveSetting "URL_Hide", cList(9)
        .INISaveSetting "URL_Reply", cList(10)
        .INISaveSetting "URL_Custom", cList(11)
    End With
End Sub

Function SysGetCloakInfo() As Variant
    Dim V As Variant
    Dim O As Integer
    O = FreeFile
    Open frmMain.INIAccess.INIFile For Random As O
    If LOF(O) = 0 Then Close O: ReDim V(-1 To -1): SysGetCloakInfo = V: Exit Function
    Close O
    ReDim V(0 To 11)
    With frmMain.INIAccess
        .INIEntry = "Cloak"
        V(0) = .INIGetSetting("Ping_Hide")
        V(1) = .INIGetSetting("Ping_Reply")
        V(2) = .INIGetSetting("Ping_Custom")
        V(3) = .INIGetSetting("Time_Hide")
        V(4) = .INIGetSetting("Time_Reply")
        V(5) = .INIGetSetting("Time_Custom")
        V(6) = .INIGetSetting("Version_Hide")
        V(7) = .INIGetSetting("Version_Reply")
        V(8) = .INIGetSetting("Version_Custom")
        V(9) = .INIGetSetting("URL_Hide")
        V(10) = .INIGetSetting("URL_Reply")
        V(11) = .INIGetSetting("URL_Custom")
    End With
    SysGetCloakInfo = V
End Function

Sub SysSaveDCCInfo(dList As Variant)
    With frmMain.INIAccess
        .INIEntry = "DCC"
        .INISaveSetting "DownloadDir", dList(0)
        .INISaveSetting "BeskyttVirus", dList(1)
        .INISaveSetting "JoinIgnore", dList(2)
        .INISaveSetting "DoIgnoreFiltyper", dList(3)
        .INISaveSetting "IgnoreFiltyper", dList(4)
        .INISaveSetting "AutoAccept", dList(5)
        .INISaveSetting "SendeBuffer", dList(6)
        .INISaveSetting "PumpDCC", dList(7)
        .INISaveSetting "PassiveDCC", dList(8)
        .INISaveSetting "UseDCCPorts", dList(9)
        .INISaveSetting "DCCPortRange", dList(10)
        .INISaveSetting "SafeMode", dList(11)
    End With
End Sub

Function SysGetDCCInfo() As Variant
    Dim V As Variant
    ReDim V(0 To 11)
    With frmMain.INIAccess
        .INIEntry = "DCC"
        V(0) = .INIGetSetting("DownloadDir")
        V(1) = .INIGetSetting("BeskyttVirus")
        V(2) = .INIGetSetting("JoinIgnore")
        V(3) = .INIGetSetting("DoIgnoreFiltyper")
        V(4) = .INIGetSetting("IgnoreFiltyper")
        V(5) = .INIGetSetting("AutoAccept")
        V(6) = .INIGetSetting("SendeBuffer")
        V(7) = .INIGetSetting("PumpDCC")
        V(8) = .INIGetSetting("PassiveDCC")
        V(9) = .INIGetSetting("UseDCCPorts")
        V(10) = .INIGetSetting("DCCPortRange")
        V(11) = .INIGetSetting("SafeMode")
    End With
    SysGetDCCInfo = V
End Function

Sub SysSaveIPInfo(iList As Variant)
    With frmMain.INIAccess
        .INIEntry = "IP"
        .INISaveSetting "IP", iList(0)
        .INISaveSetting "UseCustomIP", iList(1)
        .INISaveSetting "LookupType", iList(2)
    End With
End Sub

Function SysGetIPInfo() As Variant
    Dim V As Variant
    With frmMain.INIAccess
        .INIEntry = "IP"
        ReDim V(0 To 2)
        V(0) = .INIGetSetting("IP")
        V(1) = .INIGetSetting("UseCustomIP")
        V(2) = .INIGetSetting("LookupType")
    End With
    SysGetIPInfo = V
End Function

Sub SysSaveLogInfo(lList As Variant)
    With frmMain.INIAccess
        .INIEntry = "Logging"
        .INISaveSetting "BrukLogg", lList(0)
        .INISaveSetting "LoggDir", lList(1)
        .INISaveSetting "LoggStatus", lList(2)
        .INISaveSetting "LoggKanaler", lList(3)
        .INISaveSetting "LoggPrivat", lList(4)
        .INISaveSetting "LoggDCC", lList(5)
    End With
End Sub

Function SysGetLogInfo() As Variant
    Dim V As Variant
    With frmMain.INIAccess
        .INIEntry = "Logging"
        ReDim V(0 To 5)
        V(0) = .INIGetSetting("BrukLogg")
        V(1) = .INIGetSetting("LoggDir")
        V(2) = .INIGetSetting("LoggStatus")
        V(3) = .INIGetSetting("LoggKanaler")
        V(4) = .INIGetSetting("LoggPrivat")
        V(5) = .INIGetSetting("LoggDCC")
    End With
    SysGetLogInfo = V
End Function

Sub SysSaveDisplayInfo(iList As Variant)
    With frmMain.INIAccess
        .INIEntry = "Display"
        .INISaveSetting "Timestamp", iList(0)
        .INISaveSetting "StripCodes", iList(1)
        .INISaveSetting "StripC", iList(2)
        .INISaveSetting "StripB", iList(3)
        .INISaveSetting "StripU", iList(4)
        .INISaveSetting "StripA", iList(5)
        .INISaveSetting "FlashNew", iList(6)
        .INISaveSetting "FlashAny", iList(7)
        .INISaveSetting "ColorActivity", iList(8)
        .INISaveSetting "ShowNicklist", iList(9)
    End With
End Sub

Function SysGetDisplayInfo() As Variant
    Dim V As Variant
    With frmMain.INIAccess
        .INIEntry = "Display"
        ReDim V(0 To 9)
        V(0) = .INIGetSetting("Timestamp")
        V(1) = .INIGetSetting("StripCodes")
        V(2) = .INIGetSetting("StripC")
        V(3) = .INIGetSetting("StripB")
        V(4) = .INIGetSetting("StripU")
        V(5) = .INIGetSetting("StripA")
        V(6) = .INIGetSetting("FlashNew")
        V(7) = .INIGetSetting("FlashAny")
        V(8) = .INIGetSetting("ColorActivity")
        V(9) = .INIGetSetting("ShowNicklist")
    End With
    SysGetDisplayInfo = V
End Function

Sub SysSaveColorInfo(cList As Variant)
    With frmMain.INIAccess
    .INIEntry = "Colors"
    .INISaveSetting "cJoin", cList(0)
    .INISaveSetting "cPart", cList(1)
    .INISaveSetting "cQuit", cList(2)
    .INISaveSetting "cNick", cList(3)
    .INISaveSetting "cKick", cList(4)
    .INISaveSetting "cMode", cList(5)
    .INISaveSetting "cAction", cList(6)
    .INISaveSetting "cStatus", cList(7)
    .INISaveSetting "cTopic", cList(8)
    .INISaveSetting "cNormal", cList(9)
    .INISaveSetting "cOwn", cList(10)
    .INISaveSetting "cNotice", cList(11)
    .INISaveSetting "cBackColor", cList(12)
    .INISaveSetting "cURLColor", cList(13)
    .INISaveSetting "cBrandColor", cList(14)
    .INISaveSetting "cStdColor", cList(15)
    .INISaveSetting "cSecColor", cList(16)
    .INISaveSetting "UsemIRCColors", cList(17)
    .INISaveSetting "Font", cList(18)
    .INISaveSetting "FontSize", cList(19)
    .INISaveSetting "FontBold", cList(20)
    .INISaveSetting "FontUnderline", cList(21)
    .INISaveSetting "FontItalic", cList(22)
    End With
End Sub

Function SysGetColorInfo() As Variant
    Dim V As Variant
    With frmMain.INIAccess
        .INIEntry = "Colors"
        ReDim V(0 To 22)
        V(0) = .INIGetSetting("cJoin")
        V(1) = .INIGetSetting("cPart")
        V(2) = .INIGetSetting("cQuit")
        V(3) = .INIGetSetting("cNick")
        V(4) = .INIGetSetting("cKick")
        V(5) = .INIGetSetting("cMode")
        V(6) = .INIGetSetting("cAction")
        V(7) = .INIGetSetting("cStatus")
        V(8) = .INIGetSetting("cTopic")
        V(9) = .INIGetSetting("cNormal")
        V(10) = .INIGetSetting("cOwn")
        V(11) = .INIGetSetting("cNotice")
        V(12) = .INIGetSetting("cBackColor")
        V(13) = .INIGetSetting("cURLColor")
        V(14) = .INIGetSetting("cBrandColor")
        V(15) = .INIGetSetting("cStdColor")
        V(16) = .INIGetSetting("cSecColor")
        V(17) = .INIGetSetting("UsemIRCColors")
        V(18) = .INIGetSetting("Font")
        V(19) = .INIGetSetting("FontSize")
        V(20) = .INIGetSetting("FontBold")
        V(21) = .INIGetSetting("FontUnderline")
        V(22) = .INIGetSetting("FontItalic")
    End With
    SysGetColorInfo = V
End Function

Sub SysSaveAwayInfo(aList As Variant)
    With frmMain.INIAccess
        .INIEntry = "Away"
        .INISaveSetting "AAUse", aList(0)
        .INISaveSetting "AAMinutes", aList(1)
        .INISaveSetting "AACancelAway", aList(2)
        .INISaveSetting "AAMsg", aList(3)
        .INISaveSetting "CancelAway", aList(4)
    End With
End Sub

Function SysGetAwayInfo() As Variant
    Dim V As Variant
    With frmMain.INIAccess
        .INIEntry = "Away"
        ReDim V(0 To 4)
        V(0) = .INIGetSetting("AAUse")
        V(1) = .INIGetSetting("AAMinutes")
        V(2) = .INIGetSetting("AACancelAway")
        V(3) = .INIGetSetting("AAMsg")
        V(4) = .INIGetSetting("CancelAway")
    End With
    SysGetAwayInfo = V
End Function

Sub GetAwayInfo()
    Dim aList As Variant
    aList = SysGetAwayInfo
    If UBound(aList) = -1 Then Exit Sub
    With AwayInfo
        .AAUse = ToVal(aList(0))
        .AAMinutes = aList(1)
        .AACancelAway = ToVal(aList(2))
        .AAMsg = aList(3)
        .CancelAway = ToVal(aList(4))
    End With
End Sub

Sub SaveAwayInfo()
    Dim aList As Variant
    ReDim aList(0 To 4)
    With AwayInfo
        aList(0) = .AAUse
        aList(1) = .AAMinutes
        aList(2) = .AACancelAway
        aList(3) = .AAMsg
        aList(4) = .CancelAway
    End With
    SysSaveAwayInfo aList
End Sub

Sub SysSaveHighlightInfo(hList As Variant)
    With frmMain.INIAccess
        .INIEntry = "Highlight"
        .INISaveSetting "Use", hList(0)
        .INISaveSetting "Nick", hList(1)
        .INISaveSetting "Active", hList(2)
        .INISaveSetting "Words", hList(3)
        .INISaveSetting "WordList", hList(4)
        .INISaveSetting "UseColor", hList(5)
        .INISaveSetting "Color", hList(6)
        .INISaveSetting "UseBold", hList(7)
        .INISaveSetting "UseUnderline", hList(8)
    End With
End Sub

Function SysGetHighlightInfo() As Variant
    Dim V As Variant
    With frmMain.INIAccess
        .INIEntry = "Highlight"
        ReDim V(0 To 8)
        V(0) = .INIGetSetting("Use")
        V(1) = .INIGetSetting("Nick")
        V(2) = .INIGetSetting("Active")
        V(3) = .INIGetSetting("Words")
        V(4) = .INIGetSetting("WordList")
        V(5) = .INIGetSetting("UseColor")
        V(6) = .INIGetSetting("Color")
        V(7) = .INIGetSetting("UseBold")
        V(8) = .INIGetSetting("UseUnderline")
    End With
    SysGetHighlightInfo = V
End Function

Sub GetHighlightInfo()
    Dim hList As Variant
    hList = SysGetHighlightInfo
    If UBound(hList) = -1 Then Exit Sub
    With HighlightInfo
        .UseHighlight = ToVal(hList(0))
        .HiNick = ToVal(hList(1))
        .HiActive = ToVal(hList(2))
        .HiWords = ToVal(hList(3))
        .HiWordList = Split(hList(4), " ")
        .UseColor = ToVal(hList(5))
        .HiColor = ToVal(hList(6))
        .UseBold = ToVal(hList(7))
        .UseUnderline = ToVal(hList(8))
    End With
End Sub

Sub SaveHighlightInfo()
    Dim hList As Variant
    ReDim hList(0 To 8)
    With HighlightInfo
        hList(0) = .UseHighlight
        hList(1) = .HiNick
        hList(2) = .HiActive
        hList(3) = .HiWords
        hList(4) = join(.HiWordList, " ")
        hList(5) = .UseColor
        hList(6) = .HiColor
        hList(7) = .UseBold
        hList(8) = .UseUnderline
    End With
    SysSaveHighlightInfo hList
End Sub

Sub GetColorInfo()
    Dim cList As Variant
    Dim T As String
    With frmMain.INIAccess
        T = .INIFile
        If Not ApplyColorPath = "" Then
            .INIFile = ApplyColorPath
        End If
        cList = SysGetColorInfo
        .INIFile = T
        SysSaveColorInfo cList
        ApplyColorPath = ""
    End With
    If UBound(cList) = -1 Then Exit Sub
    With ColorInfo
        Set .Font = New StdFont
        With .Font
            .name = cList(18)
            If Not ToVal(cList(19)) = 0 Then .Size = ToVal(cList(19))
            .Bold = ToVal(cList(20))
            .Underline = ToVal(cList(21))
            .Italic = ToVal(cList(22))
        End With
        .cJoin = ToVal(cList(0))
        .cPart = ToVal(cList(1))
        .cQuit = ToVal(cList(2))
        .cNick = ToVal(cList(3))
        .cKick = ToVal(cList(4))
        .cMode = ToVal(cList(5))
        .cAction = ToVal(cList(6))
        .cStatus = ToVal(cList(7))
        .cTopic = ToVal(cList(8))
        .cNormal = ToVal(cList(9))
        .cOwn = ToVal(cList(10))
        .cNotice = ToVal(cList(11))
        .cBackColor = ToVal(cList(12))
        .cURLColor = ToVal(cList(13))
        .cBrandColor = ToVal(cList(14))
        .cStdColor = ToVal(cList(15))
        .cSecColor = ToVal(cList(16))
        .UsemIRCColors = ToVal(cList(17))
        If Not TestColor(.cURLColor) Then
            .cURLColor = RGB(0, 0, 255)
            SaveColorInfo
        End If
        If Not TestColor(.cStdColor) Then
            .cStdColor = RGB(0, 0, 255)
            SaveColorInfo
        End If
        If Not TestColor(.cSecColor) Then
            .cSecColor = RGB(128, 128, 128)
            SaveColorInfo
        End If
        URLColNum = GetColorNum(.cURLColor)
        StdColNum = GetColorNum(.cStdColor)
        SecColNum = GetColorNum(.cSecColor)
    End With
End Sub

Sub SaveColorInfo()
    Dim cList As Variant
    ReDim cList(0 To 22)
    With ColorInfo
        cList(0) = .cJoin
        cList(1) = .cPart
        cList(2) = .cQuit
        cList(3) = .cNick
        cList(4) = .cKick
        cList(5) = .cMode
        cList(6) = .cAction
        cList(7) = .cStatus
        cList(8) = .cTopic
        cList(9) = .cNormal
        cList(10) = .cOwn
        cList(11) = .cNotice
        cList(12) = .cBackColor
        cList(13) = .cURLColor
        cList(14) = .cBrandColor
        cList(15) = .cStdColor
        cList(16) = .cSecColor
        cList(17) = .UsemIRCColors
        cList(18) = .Font.name
        cList(19) = .Font.Size
        cList(20) = -.Font.Bold
        cList(21) = -.Font.Underline
        cList(22) = -.Font.Italic
        If Not TestColor(.cURLColor) Then .cURLColor = RGB(0, 0, 255)
        If Not TestColor(.cStdColor) Then .cStdColor = RGB(0, 0, 255)
        If Not TestColor(.cSecColor) Then .cSecColor = RGB(128, 128, 128)
        If Not TestColor(.cBrandColor) Then .cStdColor = RGB(0, 0, 255)
        URLColNum = GetColorNum(.cURLColor)
        StdColNum = GetColorNum(.cStdColor)
        SecColNum = GetColorNum(.cSecColor)
        BrandColNum = GetColorNum(.cBrandColor)
    End With
    SysSaveColorInfo cList
End Sub

Sub GetDisplayInfo()
    Dim dList As Variant
    dList = SysGetDisplayInfo
    If UBound(dList) = -1 Then Exit Sub
    With DisplayInfo
        .Timestamp = dList(0)
        .StripCodes = ToVal(dList(1))
        .StripC = ToVal(dList(2))
        .StripB = ToVal(dList(3))
        .StripU = ToVal(dList(4))
        .StripA = ToVal(dList(5))
        .FlashNew = ToVal(dList(6))
        .FlashAny = ToVal(dList(7))
        .ColorActivity = ToVal(dList(8))
        .ShowNicklist = ToVal(dList(9))
    End With
End Sub

Sub SaveDisplayInfo()
    Dim dList As Variant
    ReDim dList(0 To 9)
    With DisplayInfo
        dList(0) = .Timestamp
        dList(1) = .StripCodes
        dList(2) = .StripC
        dList(3) = .StripB
        dList(4) = .StripU
        dList(5) = .StripA
        dList(6) = .FlashNew
        dList(7) = .FlashAny
        dList(8) = .ColorActivity
        dList(9) = .ShowNicklist
    End With
    SysSaveDisplayInfo dList
End Sub

Sub GetLogInfo()
    Dim lList As Variant
    lList = SysGetLogInfo
    If UBound(lList) = -1 Then Exit Sub
    With LogInfo
        .BrukLogg = ToVal(lList(0))
        .LoggDir = lList(1)
        .LoggStatus = ToVal(lList(2))
        .LoggKanaler = ToVal(lList(3))
        .LoggPrivat = ToVal(lList(4))
        .LoggDCC = ToVal(lList(5))
        If .LoggDir = "" Then .LoggDir = App.Path
        If Not Right(.LoggDir, 1) = "\" Then .LoggDir = .LoggDir & "\"
    End With
End Sub

Sub SaveLogInfo()
    Dim lList As Variant
    ReDim lList(0 To 5)
    With LogInfo
        lList(0) = .BrukLogg
        lList(1) = .LoggDir
        lList(2) = .LoggStatus
        lList(3) = .LoggKanaler
        lList(4) = .LoggPrivat
        lList(5) = .LoggDCC
    End With
    SysSaveLogInfo lList
End Sub

Sub GetIPInfo()
    Dim iList As Variant
    iList = SysGetIPInfo
    If UBound(iList) = -1 Then Exit Sub
    With IPInfo
        .IP = iList(0)
        .UseCustomIP = ToVal(iList(1))
        .LookupType = ToVal(iList(2))
    End With
End Sub

Sub SaveIPInfo()
    Dim iList As Variant
    ReDim iList(0 To 2)
    With IPInfo
        iList(0) = .IP
        iList(1) = .UseCustomIP
        iList(2) = .LookupType
    End With
    SysSaveIPInfo iList
End Sub

Sub GetDCCInfo()
    Dim dList As Variant
    dList = SysGetDCCInfo
    If UBound(dList) = -1 Then Exit Sub
    With DCCInfo
        .DownloadDir = dList(0)
        .ProtectVirus = ToVal(dList(1))
        .JoinIgnore = ToVal(dList(2))
        .DoIgnoreFiltyper = ToVal(dList(3))
        .IgnoreFiltyper = dList(4)
        .AutoAccept = ToVal(dList(5))
        .SendeBuffer = ToVal(dList(6))
        .PumpDCC = ToVal(dList(7))
        .PassiveDCC = ToVal(dList(8))
        .UDCCPorts = ToVal(dList(9))
        .DCCPortRange = dList(10)
        .SafeMode = ToVal(dList(11))
    End With
End Sub

Sub SaveDCCInfo()
    Dim dList As Variant
    ReDim dList(0 To 11)
    With DCCInfo
        dList(0) = .DownloadDir
        dList(1) = .ProtectVirus
        dList(2) = .JoinIgnore
        dList(3) = .DoIgnoreFiltyper
        dList(4) = .IgnoreFiltyper
        dList(5) = .AutoAccept
        dList(6) = .SendeBuffer
        dList(7) = .PumpDCC
        dList(8) = .PassiveDCC
        dList(9) = .UDCCPorts
        dList(10) = .DCCPortRange
        dList(11) = .SafeMode
    End With
    SysSaveDCCInfo dList
End Sub

Sub GetCloakInfo()
    Dim cList As Variant
    cList = SysGetCloakInfo
    If UBound(cList) = -1 Then Exit Sub
    With Cloak
        With .Ping
            .HideRequest = ToVal(cList(0))
            .CloakType = ToVal(cList(1))
            .CustomReply = cList(2)
        End With
        With .Time
            .HideRequest = ToVal(cList(3))
            .CloakType = ToVal(cList(4))
            .CustomReply = cList(5)
        End With
        With .Version
            .HideRequest = ToVal(cList(6))
            .CloakType = ToVal(cList(7))
            .CustomReply = cList(8)
        End With
        With .URL
            .HideRequest = ToVal(cList(9))
            .CloakType = ToVal(cList(10))
            .CustomReply = cList(11)
        End With
    End With
End Sub

Sub SaveCloakInfo()
    Dim cList As Variant
    ReDim cList(0 To 11)
    With Cloak
        With .Ping
            cList(0) = .HideRequest
            cList(1) = .CloakType
            cList(2) = .CustomReply
        End With
        With .Time
            cList(3) = .HideRequest
            cList(4) = .CloakType
            cList(5) = .CustomReply
        End With
        With .Version
            cList(6) = .HideRequest
            cList(7) = .CloakType
            cList(8) = .CustomReply
        End With
        With .URL
            cList(9) = .HideRequest
            cList(10) = .CloakType
            cList(11) = .CustomReply
        End With
    End With
    SysSaveCloakInfo cList
End Sub

Sub GetIRCInfo()
    Dim SList As Variant, iList As Variant, V As Variant
    Dim C As Long
    SList = SysGetServerList
    iList = SysGetIRCInfo
    With IRCInfo
        V = SList
        If Not UBound(V) = -1 Then
            ReDim .SrvLst(0 To UBound(SList))
            ReDim .PortLst(0 To UBound(SList))
            For C = 0 To UBound(SList)
                V = Split(SList(C), ":")
                .SrvLst(C) = V(0)
                .PortLst(C) = V(1)
            Next
        End If
        V = iList
        If Not UBound(V) = -1 Then
            .Nick = iList(0)
            .Alternative = iList(1)
            .Ident = iList(2)
            .Realname = iList(3)
            .UseIdent = ToVal(iList(4))
            .Server = iList(5)
            .Port = ToVal(iList(6))
            .ModeInvisible = ToVal(iList(7))
            .ModeWallops = ToVal(iList(8))
            .AutoMode = ToVal(iList(9))
        End If
    End With
End Sub

Sub SaveIRCInfo(Optional ByVal ESL As Boolean = False)
    Dim SList As Variant, iList As Variant, V As Variant
    Dim C As Long
    With IRCInfo
        If Not ESL Then
            ReDim SList(0 To UBound(.SrvLst))
            For C = 0 To UBound(.SrvLst)
                SList(C) = .SrvLst(C) & ":" & .PortLst(C)
            Next
        End If
        ReDim iList(0 To 9)
        iList(0) = .Nick
        iList(1) = .Alternative
        iList(2) = .Ident
        iList(3) = .Realname
        iList(4) = .UseIdent
        iList(5) = .Server
        iList(6) = .Port
        iList(7) = .ModeInvisible
        iList(8) = .ModeWallops
        iList(9) = .AutoMode
    End With
    If Not ESL Then
        SysSaveServerList SList
    End If
    SysSaveIRCInfo iList
End Sub

Function ToVal(s As Variant) As Variant
    If s = "" Then s = "0"
    If Not IsNumeric(s) Then ToVal = 0: Exit Function
    ToVal = CLng(s)
End Function






'Script loading/saving on startup/exit

Sub LoadScripts()
    On Error Resume Next
    Dim s As String
    Dim C As Long
    frmMain.INIAccess.INIEntry = "Scripts"
    Do While Err = 0
        Inc C
        s = TrimCrLf(frmMain.INIAccess.INIGetSetting("n" & C))
        If Not s = "" Then
            If C = 1 Then Output "Loading scripts...", fActive, , True
            If TrimPath(s) = s Then
                If Right(App.Path, 1) = "\" Then
                    s = App.Path & s
                Else
                    s = App.Path & "\" & s
                End If
            End If
            frmScripts.DoAdd s
        Else
            If C > 1 Then Output "Finished loading scripts.", fActive, , True
            Exit Do
        End If
    Loop
    On Error GoTo 0
End Sub

Sub SaveScripts()
    Dim C As Long
    frmMain.INIAccess.INIEntry = "Scripts"
    frmMain.INIAccess.INIDeleteEntry
    For C = 1 To ScriptArrayU
        frmMain.INIAccess.INISaveSetting "n" & C, ScriptArray(C).File_Name
    Next
End Sub



'Above, only works with plugins instead

Sub LoadPlugins()
    On Error Resume Next
    Dim s As String
    Dim C As Long
    frmMain.INIAccess.INIEntry = "Plugins"
    Do While Err = 0
        Inc C
        s = frmMain.INIAccess.INIGetSetting("n" & C)
        If Not s = "" Then
            If TrimPath(s) = s Then
                If Right(App.Path, 1) = "\" Then
                    s = App.Path & s
                Else
                    s = App.Path & "\" & s
                End If
            End If
            AddOCX s
        Else
            Exit Do
        End If
    Loop
    On Error GoTo 0
End Sub

Sub SavePlugins()
    Dim C As Long
    frmMain.INIAccess.INIEntry = "Plugins"
    frmMain.INIAccess.INIDeleteEntry
    For C = 1 To airc_AddInCount
        frmMain.INIAccess.INISaveSetting "n" & C, airc_AddIns(C).Filename
        RemoveOCX C
    Next
End Sub



'IGNORE SAVING IN REGISTRY
'-------------------------

Sub SaveIgnore(ByVal IsChan As Boolean, ByVal N As String)
    Dim s As String
    If IsChan Then
        If Ignore(ChWnd(N)).join Then s = s & "JOIN "
        If Ignore(ChWnd(N)).part Then s = s & "PART "
        If Ignore(ChWnd(N)).quit Then s = s & "QUIT "
        If Ignore(ChWnd(N)).mode Then s = s & "MODE "
        If Ignore(ChWnd(N)).kick Then s = s & "KICK "
        If Ignore(ChWnd(N)).Nick Then s = s & "NICK "
        If Ignore(ChWnd(N)).Msg Then s = s & "MSG "
        s = Trim(s)
        SetStringValue "HKEY_CURRENT_USER\Software\Advanced IRC\IgnoreChan", N, s
    Else
        If IgnoreP(IgnCC(N)).Msg Then s = s & "MSG "
        If IgnoreP(IgnCC(N)).notice Then s = s & "NOTICE "
        If IgnoreP(IgnCC(N)).CTCP Then s = s & "CTCP "
        s = Trim(s)
        SetStringValue "HKEY_CURRENT_USER\Software\Advanced IRC\IgnorePriv", N, s
    End If
End Sub

Sub LoadChanIgnore(ByVal Chan As String)
    Dim s As String
    s = GetStringValue("HKEY_CURRENT_USER\Software\Advanced IRC\IgnoreChan", Chan)
    If Trim$(s) = "" Then Exit Sub
    If Trim$(UCase$(s)) = "ERROR" Then Exit Sub
    If InStr(1, s, Chr(0)) > 0 Then
        s = Mid(s, 1, InStr(1, s, Chr(0)) - 1)
    End If
    ParseIgnore Chan, s, True, False, True
End Sub

Sub LoadPrivIgnore(ByVal Nick As String)
    Dim s As String
    s = GetStringValue("HKEY_CURRENT_USER\Software\Advanced IRC\IgnorePriv", Nick)
    If Trim$(s) = "" Then Exit Sub
    If Trim$(UCase$(s)) = "ERROR" Then Exit Sub
    If InStr(1, s, Chr(0)) > 0 Then
        s = Mid(s, 1, InStr(1, s, Chr(0)) - 1)
    End If
    ParseIgnore Nick, s, True, True, True, True
End Sub

'#### autojoin

Sub LoadAutoJoin()
    Dim s As String, Srv As String
    Srv = StatusWnd(ActiveServer).ServerName
    s = GetStringValue("HKEY_CURRENT_USER\Software\Advanced IRC\Autojoin", Srv)
    If Trim$(s) = "" Then Exit Sub
    If Trim$(UCase$(s)) = "ERROR" Then Exit Sub
    If InStr(1, s, Chr(0)) > 0 Then
        s = Mid(s, 1, InStr(1, s, Chr(0)) - 1)
    End If
    s = TrimC(s, ",")
    StatusWnd(ActiveServer).AutoJoinChannels = s
    Output "Autojoin for " & Srv & " loaded, channels: " & StatusWnd(ActiveServer).AutoJoinChannels, fActive, , True
    Output "Press F11 at any time to join.", fActive, , True
End Sub

Sub AddAutoJoin(ByVal Chan As String)
    Dim Srv As String
    Srv = StatusWnd(ActiveServer).ServerName
    If Chan = "" Then Exit Sub
    If InStr(1, StatusWnd(ActiveServer).AutoJoinChannels, Chan) > 0 Then
        Output "Channel " & Chan & " already exists in " & Srv & " autojoin.", fActive, , True
        Exit Sub
    End If
    With StatusWnd(ActiveServer)
        .AutoJoinChannels = .AutoJoinChannels & "," & Chan
        .AutoJoinChannels = TrimC(.AutoJoinChannels, ",")
        SetStringValue "HKEY_CURRENT_USER\Software\Advanced IRC\Autojoin", Srv, .AutoJoinChannels
        Output "Channel " & Chan & " added to " & Srv & " autojoin.", fActive, , True
    End With
End Sub

Sub RemAutoJoin(ByVal Chan As String)
    Dim Srv As String
    If Chan = "" Then Exit Sub
    Srv = StatusWnd(ActiveServer).ServerName
    If InStr(1, StatusWnd(ActiveServer).AutoJoinChannels, Chan) = 0 Then
        Output "Channel " & Chan & " does not exist in " & Srv & " autojoin.", fActive, , True
        Exit Sub
    End If
    StatusWnd(ActiveServer).AutoJoinChannels = Replace(StatusWnd(ActiveServer).AutoJoinChannels, Chan, "")
    StatusWnd(ActiveServer).AutoJoinChannels = Replace(StatusWnd(ActiveServer).AutoJoinChannels, ",,", ",")
    StatusWnd(ActiveServer).AutoJoinChannels = TrimC(StatusWnd(ActiveServer).AutoJoinChannels, ",")
    SetStringValue "HKEY_CURRENT_USER\Software\Advanced IRC\Autojoin", Srv, StatusWnd(ActiveServer).AutoJoinChannels
    Output "Channel " & Chan & " removed from " & Srv & " autojoin.", fActive, , True
End Sub


