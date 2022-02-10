Attribute VB_Name = "modColors"
#Const DRAWAPI = True
Option Explicit


Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function ScrollWindow Lib "user32" (ByVal hWnd As Long, ByVal XAmount As Long, ByVal YAmount As Long, lpRect As RECT, lpClipRect As RECT) As Long

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

'  Ternary raster operations
Public Const SRCCOPY = &HCC0020         ' (DWORD) dest = source
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const SRCAND = &H8800C6          ' (DWORD) dest = source AND dest
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Public Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)
Public Const MERGECOPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
Public Const MERGEPAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
Public Const PATCOPY = &HF00021         ' (DWORD) dest = pattern
Public Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
Public Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Public Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
Public Const BLACKNESS = &H42           ' (DWORD) dest = BLACK
Public Const WHITENESS = &HFF0062       ' (DWORD) dest = WHITE

Public Const DT_VCENTER = &H4
Public Const DT_CENTER = &H1
Public Const DT_NOCLIP = &H100

Public Const ColorCode As String = ""
Public Const BoldCode As String = ""
Public Const UnderlineCode As String = ""
Public Const ReverseCode As String = ""

Public StdFore As Long
Public StdBack As Long

Public Type TxtLine
    TxTxt As String
    TxFg As Long
    TxBg As Long
    TxEvent As Byte
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim R As RECT
Dim temp_x As Long
Dim temp_y As Long
Dim col_fore As Long
Dim col_back As Long
Dim hUseBrush As Long

Dim UseFore As Long
Dim UseBack As Long

Public CurFont As StdFont
Public StripTypes(1 To 3) As Boolean
Public Colors(0 To 15) As Long
Public EventColors(0 To 11) As OLE_COLOR

Function ParseControlCodes(ByRef T As TxtLine, ByVal P As PictureBox) As String
    Dim S As String
    Dim TxRect As RECT
    Dim C As Long
    Dim TS As String
    Dim LogText As String
    S = T.TxTxt
    'UseFore = IIf(-1 = T.TxFg, StdFore, T.TxFg)
    UseFore = EventColors(T.TxEvent)
    UseBack = IIf(-1 = T.TxBg, StdBack, T.TxBg)
    col_fore = UseFore
    col_back = UseBack
    
    With CurFont
        P.FontName = .Name
        P.FontBold = .Bold
        P.FontItalic = .Italic
        P.FontSize = .Size
        P.FontStrikethru = .Strikethrough
        P.FontUnderline = .Underline
    End With
    
    P.ForeColor = UseFore
    
    With R
        .Left = 0
        .Top = P.CurrentY
        .Right = P.ScaleWidth
        .Bottom = .Top + P.TextHeight("M")
    End With
    
    hUseBrush = CreateSolidBrush(UseBack)
    FillRect P.hdc, R, hUseBrush
    
    For C = 1 To Len(S)
        If Mid(S, C, 1) = ColorCode Then
            If Not StripTypes(1) Then
                DrawTextRect P, TS
                GetColors Mid(S, C + 1), C, col_fore, col_back
                P.ForeColor = col_fore
                DeleteObject hUseBrush
                hUseBrush = CreateSolidBrush(col_back)
                TS = ""
            Else
                GetColors Mid(S, C + 1), C, 0, 0
            End If
        ElseIf Mid(S, C, 1) = BoldCode Then
            If Not StripTypes(2) Then
                DrawTextRect P, TS
                DeleteObject hUseBrush
                hUseBrush = CreateSolidBrush(UseBack)
                TS = ""
                P.FontBold = Not P.FontBold
            End If
        ElseIf Mid(S, C, 1) = UnderlineCode Then
            If Not StripTypes(3) Then
                DrawTextRect P, TS
                DeleteObject hUseBrush
                hUseBrush = CreateSolidBrush(UseBack)
                TS = ""
                P.FontUnderline = Not P.FontUnderline
            End If
        ElseIf Mid(S, C, 1) = ReverseCode Then
        Else
            TS = TS & Mid(S, C, 1)
            LogText = LogText & Mid(S, C, 1)
        End If
    Next
    DeleteObject hUseBrush
    DrawTextRect P, TS, True
    P.ForeColor = UseFore
    ParseControlCodes = LogText
End Function

Sub DrawTextRect(ByVal P As PictureBox, ByVal S As String, Optional ByVal StrEnd As Boolean = False)
Dim TR As RECT
With TR
.Left = P.CurrentX
.Top = P.CurrentY
.Bottom = .Top + P.TextHeight(S)
.Right = .Left + P.TextWidth(S)
End With

#If DRAWAPI Then
S = Replace(S, "&", "&&")
FillRect P.hdc, TR, hUseBrush
DrawText P.hdc, S, Len(S), TR, 0
#Else
P.Line (TR.Left, TR.Top)-(TR.Right, TR.Bottom), col_back, BF
P.CurrentX = TR.Left
P.CurrentY = TR.Top
P.Print S;
#End If

If StrEnd Then
P.CurrentX = 0
P.CurrentY = TR.Bottom
Else
P.CurrentX = TR.Right
P.CurrentY = TR.Top
End If
End Sub

'Sub DrawRect(ByVal P As PictureBox, ByVal TS As String)
'If Not R.Right = 0 Then
'R.Bottom = R.Top + P.TextHeight(TS)
'temp_x = P.CurrentX: temp_y = P.CurrentY
'P.Line (R.Left, R.Top)-(R.Right, R.Bottom), col_back, BF
'P.CurrentY = temp_y
'R.Left = 0: R.Top = 0: R.Right = 0: R.Bottom = 0
'End If
'P.Print TS;
'End Sub

Sub GetColors(ByVal S As String, ByRef NextPos As Long, ByRef ForeColor As Long, ByRef BackColor As Long)
Dim M As Long
Dim S_Buf As String
Dim Cx As Long
Dim fore_c As Long
Dim back_c As Long
ForeColor = UseFore
BackColor = UseBack
M = InStr(1, S, ",") - 1
If M <= 0 Then
M = 2
ElseIf M > 2 Then
M = 2
Else
back_c = M + 2
End If
Do Until IsNumeric(Left(S, M))
M = M - 1
If M = 0 Then
Exit Do
End If
Loop
If M = 0 Then Exit Sub
fore_c = Left(S, M)
If fore_c > 15 Then Exit Sub
If fore_c < 0 Then Exit Sub
ForeColor = Colors(fore_c)
If back_c = 0 Then
NextPos = NextPos + M
Else
Cx = back_c
Do Until Not IsNumeric(Mid(S, Cx, 1))
S_Buf = S_Buf & Mid(S, Cx, 1)
Cx = Cx + 1
If Cx > 2 + back_c Then Exit Do
Loop
If IsNumeric(Mid(S, back_c, Cx - back_c)) Then back_c = Mid(S, back_c, Cx - back_c)
If back_c > 15 Then Exit Sub
If back_c < 0 Then Exit Sub
NextPos = NextPos + Cx - 1
BackColor = Colors(back_c)
End If
End Sub

Function StripCTRL(ByVal S As String) As String
Dim C As Long
For C = 1 To Len(S)
Select Case Mid(S, C, 1)
Case ColorCode
GetColors Mid(S, C + 1), C, 0, 0
Case BoldCode
Case UnderlineCode
Case ReverseCode
Case Else
StripCTRL = StripCTRL & Mid(S, C, 1)
End Select
Next
End Function

