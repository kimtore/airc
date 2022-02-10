VERSION 5.00
Begin VB.UserControl LogBox2 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3375
   ScaleWidth      =   7455
   ToolboxBitmap   =   "LogBox2.ctx":0000
   Begin VB.VScrollBar VScroll1 
      Height          =   3015
      Left            =   6720
      Max             =   0
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   199
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   447
      TabIndex        =   0
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "LogBox2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Event Click(ByVal MouseButton As Integer, ByVal x As Long, ByVal y As Long)
Event Copy(ByVal Text As String)
Event Change()
Event DblClick()
Event OLEDragOver(ByRef Data As DataObject, ByRef Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, ByRef State As Integer)
Event OLEDragDrop(ByRef Data As DataObject, ByRef Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

Dim C As Long 'Counter

Enum ClientEvents
cjoin = 0
cpart = 1
cquit = 2
cnick = 3
ckick = 4
cmode = 5
caction = 6
cstatus = 7
ctopic = 8
cnormal = 9
cown = 10
cnotice = 11
End Enum

Dim CharLen(1 To 255) As Long
Dim Text() As TxtLine 'Original non-wordwrap
Dim ActText() As TxtLine 'Word wrapped text
Dim LU As Long 'Lines used
Dim ALU As Long 'Active lines used
Dim LH As Long, LC As Long
Dim ScrollBuf As Long
Dim NoRefresh As Boolean

'Marking variables
Dim IsMarking As Boolean
Dim MarkLineStart As Long
Dim MarkLineStartVis As Long
Dim MarkLineStartPos As Long
Dim MarkLineEnd As Long
Dim MarkLineEndVis As Long
Dim MarkLineEndPos As Long
Dim MarkLastX As Long
Dim MarkLastY As Long

Public Sub AddLine(ByVal S As String, Optional ByVal WhichEvent As ClientEvents, Optional ByVal Fore_Color As Long = -1, Optional ByVal Back_Color As Long = -1)
Dim V() As String
LU = LU + 1
ALU = ALU + 1
ReDim Preserve Text(1 To LU)
With Text(LU)
.TxTxt = S
.TxFg = Fore_Color
.TxBg = Back_Color
.TxEvent = WhichEvent
End With
V = WordWrap(S)
For C = 1 To UBound(V)
ReDim Preserve ActText(1 To ALU)
With ActText(ALU)
.TxTxt = V(C)
.TxFg = Fore_Color
.TxBg = Back_Color
.TxEvent = WhichEvent
End With
ALU = ALU + 1
Next
ALU = ALU - 1
NoRefresh = True
VScroll1.Min = 1
VScroll1.Max = ALU
If VScroll1.Value = ALU - C + 1 Then
NoRefresh = True
    If Not IsMarking Then VScroll1.Value = VScroll1.Max
NoRefresh = False
    If Not IsMarking Then
    Scroll -C + 1
    DrawLineRange LC + 1 - C, LC
    End If
ElseIf ALU = 1 Then
NoRefresh = False
DrawLineRange LC, LC
End If
NoRefresh = False
ScrollBuf = VScroll1.Value
RaiseEvent Change
End Sub

Public Sub ClearScreen()
'Reverse vbBlack
SetLines
Erase Text
Erase ActText
LU = 0
ALU = 0
VScroll1.Min = 0
VScroll1.Value = 0
VScroll1.Max = 0
Refresh
End Sub

Public Sub SetFont(ByVal NewFont As StdFont)
SetLines
Set CurFont = NewFont
With NewFont
Picture1.FontName = .Name
Picture1.FontBold = .Bold
Picture1.FontItalic = .Italic
Picture1.FontSize = .Size
Picture1.FontStrikethru = .Strikethrough
Picture1.FontUnderline = .Underline
End With
MapCharLen
Refresh
End Sub

Public Sub SetEventColors(ByVal join As OLE_COLOR, ByVal part As OLE_COLOR, ByVal quit As OLE_COLOR, ByVal nick As OLE_COLOR, ByVal kick As OLE_COLOR, ByVal mode As OLE_COLOR, ByVal action As OLE_COLOR, ByVal status As OLE_COLOR, ByVal topic As OLE_COLOR, ByVal normal As OLE_COLOR, ByVal own As OLE_COLOR, ByVal notice As OLE_COLOR)
'/*** events ***
'join=0
'part=1
'quit=2
'nick=3
'kick=4
'mode=5
'action=6
'status=7
'topic=8
'normal=9
'own=10
'notice=11
'*** end ***
EventColors(0) = join
EventColors(1) = part
EventColors(2) = quit
EventColors(3) = nick
EventColors(4) = kick
EventColors(5) = mode
EventColors(6) = action
EventColors(7) = status
EventColors(8) = topic
EventColors(9) = normal
EventColors(10) = own
EventColors(11) = notice
End Sub

Public Sub SetBackground(ByVal NewColor As Long)
SetLines
StdBack = NewColor
Picture1.BackColor = NewColor
Refresh
End Sub

Public Sub SetTextColor(ByVal NewColor As Long)
SetLines
StdFore = NewColor
Picture1.ForeColor = NewColor
Refresh
End Sub

Public Sub SetColorList(ByRef NewColors() As Long)
SetLines
For C = 0 To 15
Colors(C) = NewColors(C)
Next
Refresh
End Sub

Public Sub SetStrip(ByVal sAll As Boolean, ByVal sColor As Boolean, ByVal sBold As Boolean, ByVal sUnderline As Boolean)
SetLines
If sAll Then
sColor = True
sBold = True
sUnderline = True
End If
StripTypes(1) = sColor
StripTypes(2) = sBold
StripTypes(3) = sUnderline
End Sub

Public Function FitChar() As Long
FitChar = Picture1.ScaleWidth \ Picture1.TextWidth("M") + 1
End Function

Private Sub Picture1_DblClick()
RaiseEvent DblClick
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If IsMarking = True Then IsMarking = False: Exit Sub
MarkLastX = x: MarkLastY = y
IsMarking = True
y = Picture1.ScaleHeight - y
MarkLineStartPos = (x \ Picture1.TextWidth("M"))    'MarkLineStartPos set
MarkLineStartVis = LineLoc(y) + 1                   'MarkLineStartVis set
MarkLineStart = VScroll1.Value - (LC - MarkLineStartVis)
If MarkLineStart <= 0 Then 'User didn't click any text,
IsMarking = False 'abort
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If IsMarking Then
If (x = MarkLastX) And (y = MarkLastY) Then Exit Sub
MarkLastX = x: MarkLastY = y
y = Picture1.ScaleHeight - y
MarkLineEndPos = (x \ Picture1.TextWidth("M"))
MarkLineEndVis = LineLoc(y) + 1
MarkLineEnd = VScroll1.Value - (LC - MarkLineEndVis)
If MarkLineEndVis < MarkLineStartVis Then Exit Sub
If (MarkLineEndVis = MarkLineStartVis) And (MarkLineEndPos <= MarkLineStartPos) Then Exit Sub
'Picture1_Paint 'Slows down VERY much, let's find another way
DrawMark
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If IsMarking Then
IsMarking = False
RaiseEvent Copy(CopyText) 'Change to copied text
Picture1_Paint
'DrawLineRange MarkLineStartVis - 1, MarkLineEndVis
Exit Sub
End If
'copy startx,starty,x,y
RaiseEvent Click(Button, x, y)
End Sub

Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub Picture1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub Picture1_Paint()
Dim M As Long
If NoRefresh Then Exit Sub
If LC > VScroll1.Value Then
DrawLineRange LC - VScroll1.Value + 1, LC
Else
DrawLineRange IIf(ALU - LC <= 0, VScroll1.Value - LC - 1, 0), LC
End If
End Sub

Private Sub DrawLineRange(ByVal LineStart As Long, ByVal LineEnd As Long)
'Draw range: LineStart to LineEnd
'Propotion: LineStart = VScroll1.Value + LC - LineStart; LineEnd = VScroll1.Value - LC + LineEnd

If LineEnd - LineStart < 0 Then Exit Sub
Dim Temp%
Temp = VScroll1.Value - LC + LineStart

If Temp < 0 Then 'Clear some area

Picture1.Line (0, 0)-(Picture1.ScaleWidth, LH * Abs(LineStart)), StdBack, BF
Picture1.CurrentX = 0

Temp = 1

ElseIf Temp = 0 Then

Temp = 1

End If

If (LineStart > 0) And (VScroll1.Value < LC + 1) Then
Picture1.Line (0, 0)-(Picture1.ScaleWidth, (LH * (LineEnd - LineStart))), StdBack, BF
Picture1.CurrentX = 0
End If
Picture1.CurrentY = (Abs(LineStart) * LH) + TopL

LineStart = Temp 'Physical line number

Temp = VScroll1.Value - LC + LineEnd
If Temp < 0 Then Exit Sub
LineEnd = Temp + 1 + tmLineStart 'Physical line number
If LineEnd > ALU Then LineEnd = ALU

If (LineEnd = 0) Or (LineStart = 0) Then Exit Sub

For C = LineStart To LineEnd
ParseControlCodes ActText(C), Picture1
Next

End Sub

Public Sub Refresh()
SetLines
Picture1_Paint
End Sub

Public Sub HardRefresh() 'Refresh with F5
SetLines
Picture1.Cls
Picture1_Paint
End Sub

Private Sub UserControl_Initialize()
Randomize
ClearScreen
End Sub

Private Sub SetLines()
LH = Picture1.TextHeight("M")
LC = Picture1.ScaleHeight \ LH - IIf(Picture1.ScaleHeight Mod LH = 0, 1, 0)
End Sub

Private Sub UserControl_Resize()
If (UserControl.ScaleWidth - VScroll1.Width) <= 0 Then
UserControl.Width = 3000
UserControl.Height = 3000
Exit Sub
End If
Picture1.Width = UserControl.ScaleWidth - VScroll1.Width
Picture1.Height = UserControl.ScaleHeight
VScroll1.Left = Picture1.Width - 10
VScroll1.Height = UserControl.ScaleHeight
SetLines
VScroll1.LargeChange = IIf(LC > 1, LC \ 2, 1)
BuildWordWrap
End Sub

Private Sub VScroll1_Change()
Dim TmpVal As Long
If NoRefresh Then Exit Sub
Scroll (ScrollBuf - VScroll1.Value)
If VScroll1.Value > ScrollBuf Then 'Ned
DrawLineRange LC - (VScroll1.Value - ScrollBuf), LC 'OK
ElseIf VScroll1.Value < ScrollBuf Then 'Opp
TmpVal = LC - VScroll1.Value + 1
DrawLineRange IIf(TmpVal > 0, TmpVal, 0), ScrollBuf - VScroll1.Value + IIf(TmpVal > 0, TmpVal, 0) 'OK
End If
ScrollBuf = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
VScroll1_Change
End Sub

' ####### EXPERIMENTAL #######

Private Sub Scroll(Optional ByVal N As Long = 1)
Dim PL As Long, RV As Long, R As RECT, CR As RECT
PL = Abs(N)
RV = Sgn(N)
R.Bottom = Picture1.ScaleHeight: R.Right = Picture1.ScaleWidth
CR = R
If RV = 1 Then 'Move window contents DOWN
    If PL > 1 Then 'Too much scroll, redraw using BitBlt
        BitBlt Picture1.hdc, 0, LH * PL, Picture1.ScaleWidth, Picture1.ScaleHeight - LH * PL, Picture1.hdc, 0, 0, SRCCOPY
    Else 'Use ScrollWindow, faster
        R.Bottom = Picture1.ScaleHeight - (LH * PL)
        ScrollWindow Picture1.hWnd, 0, LH * PL, R, CR
    End If
ElseIf RV = -1 Then 'Move window contents UP
    If PL > 1 Then 'Too much scroll, redraw using BitBlt
        BitBlt Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight - LH * PL, Picture1.hdc, 0, LH * PL, SRCCOPY
    Else 'Use ScrollWindow, faster
        R.Top = (LH * PL)
        ScrollWindow Picture1.hWnd, 0, -LH * PL, R, CR
    End If
End If ' 0 = 0
End Sub

Private Sub EraseLine(ByVal LineNum As Long, Optional ByVal Color As OLE_COLOR = -1)
If Color = -1 Then Color = StdFore
Dim bk_x As Long, bk_y As Long
bk_x = Picture1.CurrentX: bk_y = Picture1.CurrentY
Picture1.Line (0, TopL + (LH * LineNum) - LH)-(Picture1.ScaleWidth, TopL + (LH * LineNum)), Color, BF
Picture1.CurrentX = bk_x: Picture1.CurrentY = bk_y
End Sub

Private Sub ErasePartial(ByVal YPos As Long, Optional ByVal CharStartPos As Long = 0, Optional ByVal CharEndPos As Long = 0, Optional ByVal Color As OLE_COLOR = -1)
If Color = -1 Then Color = StdFore
Dim bk_x As Long, bk_y As Long
bk_x = Picture1.CurrentX: bk_y = Picture1.CurrentY
If CharEndPos = 0 Then CharEndPos = Picture1.ScaleWidth Else CharEndPos = (CharEndPos * TxtLen("M")) + TxtLen("M")
Picture1.Line (CharStartPos * TxtLen("M"), TopL + (LH * YPos) - LH)-(CharEndPos, TopL + (LH * YPos)), Color, BF
Picture1.CurrentX = bk_x: Picture1.CurrentY = bk_y
End Sub

Private Function TopL() As Long
TopL = Picture1.ScaleHeight Mod LH
Do Until TopL <= 0
TopL = TopL - LH
Loop
End Function

Private Sub MapCharLen()
'Not in use ATM

'For C = 1 To 255
'CharLen(C) = Picture1.TextWidth(Chr(C))
'Next
End Sub

Private Sub DrawPartialLine(ByVal YPos As Long, ByVal LPos As Long, Optional ByVal CharStartPos As Long = 0, Optional ByVal CharEndPos As Long = 0)
If (LPos > ALU) Or (LPos = 0) Then Exit Sub
Dim bk_x As Long, bk_y As Long
Dim TmpS$
bk_x = Picture1.CurrentX: bk_y = Picture1.CurrentY
If CharStartPos = 0 Then CharStartPos = 1
Picture1.CurrentX = (TxtLen("M") * CharStartPos) - TxtLen("M")
Picture1.CurrentY = (TopL + (YPos * LH)) - LH
If CharEndPos = 0 Then CharEndPos = Len(StripCTRL(ActText(LPos).TxTxt)) + 1
Picture1.ForeColor = StdBack
TmpS = StripCTRL(ActText(LPos).TxTxt)
If Not (CharEndPos - CharStartPos < 1) Then
Picture1.Print Mid(TmpS, CharStartPos, CharEndPos - CharStartPos)
End If
Picture1.ForeColor = StdFore
Picture1.CurrentX = bk_x: Picture1.CurrentY = bk_y
End Sub

Private Sub DrawMark()
'Refresh (Have to have an effective way to clear off unused copyfield)
If MarkLineStartVis = MarkLineEndVis Then 'One line only
ErasePartial MarkLineStartVis, MarkLineStartPos, MarkLineEndPos
DrawPartialLine MarkLineStartVis, MarkLineStart - 1, MarkLineStartPos + 1, MarkLineEndPos + 2
Else
ErasePartial MarkLineStartVis, MarkLineStartPos
DrawPartialLine MarkLineStartVis, MarkLineStart - 1, MarkLineStartPos + 1
For C = MarkLineStartVis + 1 To MarkLineEndVis - 1
EraseLine C
MarkLine C
Next
ErasePartial MarkLineEndVis, CharEndPos:=MarkLineEndPos
DrawPartialLine MarkLineEndVis, MarkLineEnd - 1, 0, MarkLineEndPos + 2
End If
End Sub

Private Function CopyText() As String
If (MarkLineStart <= 1) Or (MarkLineEnd <= 1) Then Exit Function
If MarkLineEnd < MarkLineStart Then Exit Function
If (MarkLineEnd = MarkLineStart) And (MarkLineEndPos <= MarkLineStartPos) Then Exit Function
If MarkLineStart = MarkLineEnd Then 'One line to copy
CopyText = Mid(StripCTRL(ActText(MarkLineStart - 1).TxTxt), MarkLineStartPos + 1, MarkLineEndPos - MarkLineStartPos + 1)
Else
CopyText = Mid(StripCTRL(ActText(MarkLineStart - 1).TxTxt), MarkLineStartPos + 1) & vbCrLf
If MarkLineEnd - 2 > ALU Then Exit Function
For C = MarkLineStart + 1 To MarkLineEnd - 1
CopyText = CopyText & StripCTRL(ActText(C - 1).TxTxt) & vbCrLf
Next
If (MarkLineEnd > 0) And (MarkLineEnd - 1 <= ALU) Then CopyText = CopyText & Mid(StripCTRL(ActText(MarkLineEnd - 1).TxTxt), 1, MarkLineEndPos + 1)
End If
Clipboard.Clear
Clipboard.SetText CopyText, vbCFText
End Function

Private Sub MarkLine(ByVal LineNum As Long)
NoRefresh = True
EraseLine LineNum
DrawPartialLine LineNum, VScroll1.Value - (LC - LineNum) - 1
NoRefresh = False
End Sub

Private Function LineLoc(ByVal y As Long)
LineLoc = LC - ((y \ LH) + 1) + 1
End Function

Private Sub BuildWordWrap()
Dim V() As String
Dim D As Long
Erase ActText
ALU = 1
For D = 1 To LU
V = WordWrap(Text(D).TxTxt)
For C = 1 To UBound(V)
ReDim Preserve ActText(1 To ALU)
With ActText(ALU)
.TxTxt = V(C)
.TxFg = Text(D).TxFg
.TxBg = Text(D).TxBg
.TxEvent = Text(D).TxEvent
End With
ALU = ALU + 1
Next
Next
ALU = ALU - 1
VScroll1.Max = ALU
VScroll1.Value = VScroll1.Max
DrawLineRange VScroll1.Value - LC - 1, LC
ScrollBuf = VScroll1.Value
End Sub

Private Function WordWrap(ByVal S As String) As String()
Dim i As Long
Dim I2 As Long
Dim V() As String
Dim U As Long
    i = InStr(1, S, " ")
    Do
        If Picture1.ScaleWidth > TxtLen(Mid(SL, 1, I2)) Then
            i = InStr(i + 1, S, " ")
            I2 = InStr(i + 1, S, " ")
            If I2 = 0 Then I2 = Len(S)
        End If
        If i = 0 Then
            U = U + 1
            ReDim Preserve V(1 To U)
            V(U) = S
            WordWrap = V
            Exit Function
        End If
        If Picture1.ScaleWidth <= TxtLen(Mid(S, 1, I2)) Then
            U = U + 1
            ReDim Preserve V(1 To U)
            V(U) = Mid(S, 1, i)
            S = " " & Mid(S, i)
            I2 = 2
            i = 2
        End If
        'Debug.Print Len(S)
    Loop
End Function

Private Function TxtLen(ByVal S As String) As Long
TxtLen = Picture1.TextWidth(S)
End Function
