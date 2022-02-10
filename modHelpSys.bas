Attribute VB_Name = "modHelpSys"
Const HH_DISPLAY_TOPIC = &H0
Const HH_SET_WIN_TYPE = &H4
Const HH_GET_WIN_TYPE = &H5
Const HH_GET_WIN_HANDLE = &H6
Const HH_DISPLAY_TEXT_POPUP = &HE   ' Display string resource ID or
                                    ' text in a pop-up window.
Const HH_HELP_CONTEXT = &HF         ' Display mapped numeric value in
                                    ' dwData.
Const HH_TP_HELP_CONTEXTMENU = &H10 ' Text pop-up help, similar to
                                    ' WinHelp's HELP_CONTEXTMENU.
Const HH_TP_HELP_WM_HELP = &H11     ' text pop-up help, similar to
                                    ' WinHelp's HELP_WM_HELP.

Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
   (ByVal hwndCaller As Long, ByVal pszFile As String, _
   ByVal uCommand As Long, ByVal dwData As Long) As Long

Public Function ShowHelp(ByVal hWnd As Long, Optional ByVal ContextID As Long) As Long
    If ContextID = 0 Then
        ShowHelp = HtmlHelp(hWnd, "airc.chm", HH_DISPLAY_TOPIC, 0)
    Else
        ShowHelp = HtmlHelp(hWnd, "airc.chm", HH_HELP_CONTEXT, ContextID)
    End If
End Function
