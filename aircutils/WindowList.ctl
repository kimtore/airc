VERSION 5.00
Begin VB.UserControl WList 
   AccessKeys      =   "123456789"
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   178
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   252
   ToolboxBitmap   =   "WindowList.ctx":0000
   Windowless      =   -1  'True
End
Attribute VB_Name = "WList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Event RightClick()

Public DrawMode As Integer

Enum WndTypes
wndStatus = 1
wndChannel = 2
wndPrivate = 3
wndChat = 4
wndDCC = 5
wndOther = 99
End Enum

Enum SortOrder
   SortAscending = 0
   SortDescending = 1
End Enum

Private Type WndLDef
Wnd As Object
Col As Long
Indx As Integer
SrvNum As Integer 'Server number
Flags As Integer 'Window type
Ico As StdPicture
Title As String
End Type

Dim Wnds() As WndLDef
Dim WndsBK() As WndLDef

Dim CurSel As Integer, ItemCount As Integer, ButLen As Integer

Private Sub UserControl_Initialize()
CurSel = 1
ButLen = GetButLen
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim butnun As Integer
If Button = 2 Then 'Right click, raise rightclick event
RaiseEvent RightClick
Exit Sub
End If
If ((x < 0) Or (y < 0)) Or ((x > ScaleWidth) Or (y > ScaleHeight)) Then Exit Sub
If DrawMode = 0 Then
    butnun = Fix((y / 20) + 1)
Else
    butnun = Fix((x / ButLen) + 1)
End If
If (butnun < 1) Or (butnun > ItemCount) Then Exit Sub ' <= Ugyldig
Dim FX As Form
Set FX = Wnds(butnun).Wnd
If Shift = 0 Then
If Wnds(butnun).Wnd.Visible <> True Then Wnds(butnun).Wnd.Visible = True
Wnds(butnun).Wnd.SetFocus
ElseIf Shift = vbShiftMask Then
Unload Wnds(butnun).Wnd ' <= Gone.. eller ikke
Refresh
End If
CurSel = butnun
UserControl_Paint
End Sub

Sub SetDrawMode(ByVal i As Integer)
DrawMode = i
ButLen = GetButLen
Refresh
End Sub

Private Function GetButLen() As Integer
If DrawMode = 0 Then Exit Function
If ItemCount = 0 Then GetButLen = 133: Exit Function
GetButLen = ScaleWidth \ ItemCount
If GetButLen > 133 Then GetButLen = 133
End Function

Private Sub UserControl_Paint()
Dim x As Form
Dim V As Integer
If DrawMode = 0 Then '*** VERTICAL ***
    For V = 0 To ItemCount - 1
        If CurSel = V + 1 Then
            Line (0, V * 20)-(ScaleWidth - 2, V * 20 + 18), &H0, B
            Line (1, V * 20 + 1)-(ScaleWidth - 1, V * 20 + 19), &HFFFFFF, B
            Line (2, V * 20 + 2)-(ScaleWidth - 3, V * 20 + 17), &H808080, BF
            ForeColor = &HFFFFFF
            Wnds(V + 1).Col = 0
        Else
            Line (0, V * 20)-(ScaleWidth - 2, V * 20 + 18), &H0, B
            Line (1, V * 20 + 1)-(ScaleWidth - 1, V * 20 + 19), &HFFFFFF, B
            Line (2, V * 20 + 2)-(ScaleWidth - 3, V * 20 + 17), vbButtonFace, BF
            ForeColor = &H0
            If Not Wnds(V + 1).Col = 0 Then ForeColor = Wnds(V + 1).Col
        End If
        Set x = Wnds(V + 1).Wnd
        PaintPicture x.Icon, 2, V * 20 + 2
        CurrentX = 18: CurrentY = V * 20 + 3
        If x.Tag = "" Then Print x.Caption Else Print x.Tag
    Next
Else '*** HORIZONTAL ***
    For V = 0 To ItemCount - 1
        If CurSel = V + 1 Then
            Line (V * ButLen, 0)-(V * ButLen + ButLen - 2, ScaleHeight - 2), &H0, B
            Line (V * ButLen + 1, 1)-(V * ButLen + ButLen - 1, ScaleHeight - 1), &HFFFFFF, B
            Line (V * ButLen + 2, 2)-(V * ButLen + ButLen - 3, ScaleHeight - 3), &H808080, BF
            ForeColor = &HFFFFFF
            Wnds(V + 1).Col = 0
        Else
            Line (V * ButLen, 0)-(V * ButLen + ButLen - 2, ScaleHeight - 2), &H0, B
            Line (V * ButLen + 1, 1)-(V * ButLen + ButLen - 1, ScaleHeight - 1), &HFFFFFF, B
            Line (V * ButLen + 2, 2)-(V * ButLen + ButLen - 3, ScaleHeight - 3), &H8000000F, BF
            ForeColor = &H0
            If Not Wnds(V + 1).Col = 0 Then ForeColor = Wnds(V + 1).Col
        End If
        Set x = Wnds(V + 1).Wnd
        PaintPicture x.Icon, V * ButLen + 2, 2
        CurrentY = 3: CurrentX = V * ButLen + 18
        If x.Tag = "" Then Print x.Caption Else Print x.Tag
    Next
End If
End Sub

Function AddWnd(W As Object, ByVal ServerNum As Integer, Optional WndType As WndTypes = wndOther)
ItemCount = ItemCount + 1
ReDim Preserve Wnds(1 To ItemCount)
Set Wnds(ItemCount).Wnd = W
Wnds(ItemCount).Flags = WndType
Wnds(ItemCount).Title = Wnds(ItemCount).Wnd.Tag
If Wnds(ItemCount).Title = "" Then Wnds(ItemCount).Title = Wnds(ItemCount).Wnd.Caption
Wnds(ItemCount).SrvNum = ServerNum
ButLen = GetButLen
SortWnds
End Function

Sub NumWnd(ByVal i As Integer)
If ((ItemCount = 0) Or (i > ItemCount) Or (i < 1)) Then Exit Sub
CurSel = i
Wnds(i).Wnd.SetFocus
Refresh
End Sub

Sub ActWnd(W As Object)
Dim V As Integer
For V = 1 To ItemCount
If Wnds(V).Wnd Is W Then CurSel = V ': Wnds(V).Wnd.SetFocus
Next
Refresh
End Sub

Sub ColWnd(W As Object, Color As Long)
Dim V As Integer
For V = 1 To ItemCount
If Wnds(V).Wnd Is W Then Wnds(V).Col = Color: Exit For
Next
Refresh
End Sub

Function GetColWnd(W As Object) As Long
Dim V As Integer
For V = 1 To ItemCount
If Wnds(V).Wnd Is W Then GetColWnd = Wnds(V).Col: Exit For
Next
End Function

Sub RemWnd(W As Object)
Dim V As Integer
Dim Z As Integer
For V = 1 To ItemCount
If Wnds(V).Wnd Is W Then Exit For
Next
If V > ItemCount Then Exit Sub
For Z = V To ItemCount - 1
Wnds(Z) = Wnds(Z + 1)
Next
ItemCount = ItemCount - 1
If ItemCount <= 0 Then
ReDim Wnds(1 To 1)
Else
ReDim Preserve Wnds(1 To ItemCount)
'ActWnd Wnds(v).Wnd
End If
Cls
ButLen = GetButLen
Refresh
UserControl_Paint
End Sub

Sub Refresh()
Dim V As Integer
If ItemCount = 0 Then Exit Sub
For V = 1 To ItemCount
With Wnds(V)
If Not .Flags = 99 Then
.SrvNum = .Wnd.ServerNum
Else
.SrvNum = 1
End If
End With
Next
SortWnds
End Sub

Private Sub SortWnds()
    If UBound(Wnds) - LBound(Wnds) = 0 Then 'Array may be empty
        If Wnds(1).Wnd Is Nothing Then Exit Sub 'Array IS empty
    End If
    SortArray Wnds
    WndsBK = Wnds
    Dim C As Integer 'Counter
    Dim G As Integer 'Counter 2
    Dim S As Integer 'Counter 3
    Dim M As Integer 'WndsBK counter
    Dim LB As Integer 'LowBound server number
    Dim UB As Integer 'UpperBound server number
    LB = 32767 'Set to max
    For G = 1 To ItemCount
        With Wnds(G)
            If LB > .SrvNum Then LB = .SrvNum
            If UB < .SrvNum Then UB = .SrvNum
        End With
    Next
    For C = LB To UB
        For G = 1 To 5
            For S = 1 To ItemCount
                If ((Wnds(S).Flags = G) And (Wnds(S).SrvNum = C)) Then
                    M = M + 1
                    WndsBK(M) = Wnds(S)
                End If
                If M = ItemCount Then Exit For
            Next
        Next
    Next
    For S = 1 To ItemCount
        If Wnds(S).Flags = 99 Then
            M = M + 1
            WndsBK(M) = Wnds(S)
        End If
        If M = ItemCount Then Exit For
    Next
    Wnds = WndsBK
    UserControl_Paint
End Sub

Private Sub SortArray(ByRef sArray() As WndLDef, Optional ByVal SortOrder As SortOrder = SortAscending)
   Dim i          As Long   ' Loop Counter
   Dim j          As Long
   Dim iLBound    As Long
   Dim iUBound    As Long
   Dim iMax       As Long
   Dim sTemp      As WndLDef
   Dim distance   As Long
   Dim bSortOrder As Boolean
   
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)

   If Not iLBound = LBound(sArray) Then Exit Sub
   If Not iUBound = UBound(sArray) Then Exit Sub

   bSortOrder = IIf(SortOrder = SortAscending, False, True)
   iMax = iUBound - iLBound + 1
   
   Do
      distance = distance * 3 + 1
   Loop Until distance > iMax

   Do
      distance = distance \ 3
      For i = distance + iLBound To iUBound
         sTemp = sArray(i)
         j = i
         Do While (sArray(j - distance).Title > sTemp.Title) Xor bSortOrder
            sArray(j) = sArray(j - distance)
            j = j - distance
            If j - distance < iLBound Then Exit Do
         Loop
         sArray(j) = sTemp
      Next i
   Loop Until distance = 1

End Sub

